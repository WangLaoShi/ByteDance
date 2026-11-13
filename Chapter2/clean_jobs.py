"""清洗字节职位表，生成 QuickBI 可用的 Excel 和清洗说明。"""

import argparse
import re
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from address_dimensions import parse_address_excel
from date_dimensions import time_features, time_features_excel
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

EXPECTED_COLUMNS = (
    "id",
    "code",
    "title",
    "sub_title",
    "portal_type",
    "portal_name",
    "fetched_at",
    "publish_time",
    "category_id",
    "category_name",
    "category_parent_name",
    "city_code",
    "city_name",
    "city_names",
    "recruit_type_id",
    "recruit_type_name",
    "recruit_parent_name",
    "job_subject_name",
    "description",
    "requirement",
    "address",
)
TEXT_COLUMNS = (
    "id",
    "code",
    "category_id",
    "city_code",
    "portal_type",
    "recruit_type_id",
)
CLASS_COLUMNS = (
    "category_name",
    "category_parent_name",
    "city_name",
    "city_names",
    "portal_name",
    "recruit_type_name",
    "recruit_parent_name",
    "job_subject_name",
)
FREE_TEXT_COLUMNS = ("title", "description", "requirement", "address")
BUSINESS_COLUMNS = tuple(name for name in EXPECTED_COLUMNS if name != "id")
PLACEHOLDERS = {"null", "none", "nan", "n/a", "na", "undefined", "-", "--", "无"}
INVISIBLE = re.compile(r"[\u3000\u00a0\u200b\ufeff\u200c\u200d]")
CITY_SPLIT = re.compile(r"[、,，;；]+")
SHANGHAI = ZoneInfo("Asia/Shanghai")
RECRUIT_GROUPS = {
    ("2", "社招", "101", "正式", "社招"): "社招正式",
    ("3", "校招", "201", "正式", "校招"): "校招正式",
    ("3", "校招", "202", "实习", "校招"): "校招实习",
}
SOURCE_SHEET = "职位"

MAIN_COLUMNS = (
    "source_row",
    "id",
    "code",
    "title",
    "portal_type",
    "portal_name",
    "fetched_at",
    "publish_time",
    "category_id",
    "category_name",
    "category_parent_name",
    "city_code",
    "city_name",
    "city_names",
    "recruit_type_id",
    "recruit_type_name",
    "recruit_parent_name",
    "job_subject_name",
    "description",
    "requirement",
    "address",
    "recruit_group",
    "project_status",
    "analysis_category_l1",
    "analysis_category_l2",
    "category_mapping_status",
    "publish_datetime",
    "publish_date",
    "publish_month",
    "publish_time_status",
    "fetched_datetime",
    "publish_year",
    "publish_quarter",
    "publish_day",
    "publish_day_name",
    "publish_is_weekend",
    "publish_iso_year",
    "publish_weekofyear",
    "publish_week_start",
    "publish_dayofyear",
    "publish_ten_day_period",
    "address_province",
    "address_city",
    "address_district",
    "address_township",
    "address_adcode",
    "address_detail",
    "address_status",
    "address_parse_source",
    "address_candidate_codes",
    "address_data_version",
)
CITY_COLUMNS = (
    "id",
    "source_row",
    "city_raw",
    "city_standard",
    "province",
    "country_or_region",
    "geo_code",
    "city_mapping_status",
    "city_source",
    "city_code",
    "city_name",
    "recruit_group",
    "analysis_category_l1",
    "analysis_category_l2",
    "publish_month",
    "publish_time_status",
)
DATE_FORMATS = {
    "publish_datetime": "yyyy-mm-dd hh:mm:ss",
    "fetched_datetime": "yyyy-mm-dd hh:mm:ss",
    "publish_date": "yyyy-mm-dd",
    "publish_month": "yyyy-mm-dd",
    "publish_week_start": "yyyy-mm-dd",
}


def parse_args():
    script_dir = Path(__file__).resolve().parent
    root = script_dir.parent
    parser = argparse.ArgumentParser(description="清洗职位表并输出 QuickBI 工作簿")
    parser.add_argument(
        "--input",
        type=Path,
        default=root / "Chapter1" / "data" / "jobs.xlsx",
        help="源 Excel 路径",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=script_dir,
        help="输出目录",
    )
    parser.add_argument(
        "--category-catalog",
        type=Path,
        default=root / "Chapter1" / "data" / "job_categories.csv",
        help="职位类别表，用来确定大类和细类",
    )
    parser.add_argument(
        "--category-map",
        type=Path,
        default=None,
        help="人工类别映射 CSV，需含 category_id、analysis_category_l1、analysis_category_l2",
    )
    parser.add_argument(
        "--city-map",
        type=Path,
        default=None,
        help="城市映射 CSV，需含 city_raw、city_standard、province、country_or_region、geo_code",
    )
    parser.add_argument(
        "--early-before",
        default="2020-01-01",
        help="早于该日期（北京时间）的发布日期标为早期记录待复核",
    )
    return parser.parse_args()


def is_blank(value):
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() == ""


def whole_cell_placeholder(value):
    if is_blank(value):
        return True
    return str(value).strip().casefold() in PLACEHOLDERS


def clean_token(value):
    if is_blank(value):
        return None
    text = INVISIBLE.sub("", str(value)).strip()
    if text.casefold() in PLACEHOLDERS:
        return None
    return text


def as_text(value):
    if is_blank(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def load_source(path):
    workbook = load_workbook(path, read_only=True, data_only=False)
    if SOURCE_SHEET not in workbook.sheetnames:
        workbook.close()
        raise SystemExit(f"找不到工作表 {SOURCE_SHEET}")
    sheet = workbook[SOURCE_SHEET]
    row_iter = sheet.iter_rows()
    header = [cell.value for cell in next(row_iter)]
    missing = [name for name in EXPECTED_COLUMNS if name not in header]
    if missing:
        workbook.close()
        raise SystemExit("关键字段缺失: " + ", ".join(missing))
    extra = [name for name in header if name not in EXPECTED_COLUMNS]
    records = []
    suspect_ids = []
    numeric_cells = []
    placeholder_fields = {}
    for excel_row, cells in enumerate(row_iter, start=2):
        values = {name: cells[header.index(name)].value for name in header}
        record = {"source_row": excel_row}
        for name in EXPECTED_COLUMNS:
            cell = cells[header.index(name)]
            value = cell.value
            if name in TEXT_COLUMNS and cell.data_type == "n" and value is not None:
                digits = str(abs(int(value)))
                numeric_cells.append((excel_row, name, digits))
                if len(digits) > 15:
                    suspect_ids.append(
                        {
                            "source_row": excel_row,
                            "id": digits,
                            "字段": name,
                            "异常类型": "ID 精度可疑",
                            "原值": digits,
                            "处理动作": "从岗位主表排除，不猜测原始 ID",
                        }
                    )
            if whole_cell_placeholder(value) and not is_blank(value):
                placeholder_fields.setdefault(name, 0)
                placeholder_fields[name] += 1
            if name in FREE_TEXT_COLUMNS:
                record[name] = None if is_blank(value) else value if not isinstance(value, str) else value
                if isinstance(value, str) and value.strip() == "":
                    record[name] = None
            elif name in CLASS_COLUMNS or name in TEXT_COLUMNS:
                record[name] = clean_token(value) if name in CLASS_COLUMNS else as_text(value)
                if name in TEXT_COLUMNS and whole_cell_placeholder(value):
                    record[name] = None
            else:
                record[name] = None if whole_cell_placeholder(value) else value
        records.append(record)
    workbook.close()
    frame = pd.DataFrame.from_records(records)
    return frame, extra, suspect_ids, numeric_cells, placeholder_fields


def load_category_map(path):
    if path is None:
        return {}
    table = pd.read_csv(path, dtype=str)
    required = {"category_id", "analysis_category_l1", "analysis_category_l2"}
    missing = required - set(table.columns)
    if missing:
        raise SystemExit("类别映射缺少字段: " + ", ".join(sorted(missing)))
    if table["category_id"].duplicated().any():
        raise SystemExit("人工类别映射的 category_id 必须唯一")
    mapping = {}
    for row in table.itertuples(index=False):
        mapping[str(row.category_id)] = (
            clean_token(row.analysis_category_l1) or "待确认",
            clean_token(row.analysis_category_l2) or "待确认",
        )
    return mapping


def load_city_map(path):
    if path is None:
        return {}
    table = pd.read_csv(path, dtype=str)
    required = {"city_raw", "city_standard", "province", "country_or_region", "geo_code"}
    missing = required - set(table.columns)
    if missing:
        raise SystemExit("城市映射缺少字段: " + ", ".join(sorted(missing)))
    mapping = {}
    for row in table.itertuples(index=False):
        raw = clean_token(row.city_raw)
        if raw is None:
            continue
        mapping[raw] = {
            "city_standard": clean_token(row.city_standard) or raw,
            "province": clean_token(row.province),
            "country_or_region": clean_token(row.country_or_region),
            "geo_code": as_text(row.geo_code),
        }
    return mapping


def load_category_catalog(path):
    if path is None or not path.exists():
        return {}
    table = pd.read_csv(path, dtype=str).fillna("")
    required = {"id", "name", "parent_id"}
    missing = required - set(table.columns)
    if missing:
        raise SystemExit("职位类别表缺少字段: " + ", ".join(sorted(missing)))
    names = {}
    parents = {}
    for row in table.itertuples(index=False):
        category_id = clean_token(row.id)
        if category_id is None:
            continue
        names.setdefault(category_id, set()).add(clean_token(row.name) or "")
        parents.setdefault(category_id, set()).add(clean_token(row.parent_id) or "")
    catalog = {}
    for category_id, name_set in names.items():
        parent_ids = {parent_id for parent_id in parents[category_id] if parent_id}
        usable_names = {name for name in name_set if name}
        if len(usable_names) != 1 or len(parent_ids) > 1:
            continue
        parent_id = next(iter(parent_ids)) if parent_ids else ""
        catalog[category_id] = {
            "name": next(iter(usable_names)),
            "parent_id": parent_id,
        }
    for item in catalog.values():
        parent = catalog.get(item["parent_id"])
        item["parent_name"] = parent["name"] if parent else ""
    return catalog


def build_category_rules(frame, manual_map, catalog):
    rules = {}
    exceptions = []
    for category_id, group in frame.groupby("category_id", dropna=False):
        names = sorted({value for value in group["category_name"] if value})
        parents = sorted({value for value in group["category_parent_name"] if value})
        key = category_id or ""
        if len(names) != 1 or len(parents) > 1:
            rules[key] = {
                "l1": "待确认",
                "l2": "待确认",
                "status": "待确认",
                "name": "、".join(names),
                "parent": "、".join(parents),
            }
            exceptions.append(
                {
                    "source_row": "",
                    "id": "",
                    "字段": "category_id",
                    "异常类型": "类别名称或上级冲突",
                    "原值": f"{key} / {names} / {parents}",
                    "处理动作": "分析类别标为待确认",
                }
            )
            continue
        name = names[0] if names else None
        if key in manual_map:
            l1, l2 = manual_map[key]
            rules[key] = {
                "l1": l1,
                "l2": l2,
                "status": "人工映射",
                "name": name or "",
                "parent": parents[0] if parents else "",
            }
            continue
        catalog_item = catalog.get(key)
        if catalog_item:
            category_name = name or catalog_item["name"]
            if catalog_item["parent_name"]:
                rules[key] = {
                    "l1": catalog_item["parent_name"],
                    "l2": category_name,
                    "status": "类别表",
                    "name": category_name,
                    "parent": catalog_item["parent_name"],
                }
            else:
                rules[key] = {
                    "l1": category_name,
                    "l2": category_name,
                    "status": "类别表顶级",
                    "name": category_name,
                    "parent": "",
                }
            continue
        if len(parents) == 1:
            empty_parent = group["category_parent_name"].isna().any()
            rules[key] = {
                "l1": parents[0],
                "l2": name or "待确认",
                "status": "同 ID 补齐" if empty_parent else "已映射",
                "name": name or "",
                "parent": parents[0],
            }
            continue
        rules[key] = {
            "l1": "待确认",
            "l2": name or "待确认",
            "status": "待确认",
            "name": name or "",
            "parent": "",
        }
    return rules, exceptions


def parse_fetched_at(value):
    if is_blank(value):
        return None
    text = str(value).strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=SHANGHAI)
    return parsed.astimezone(SHANGHAI)


def choose_rows(frame, suspect_row_ids):
    """每个有效 ID 留一条。返回主表行、快照行、异常和核数。"""
    exceptions = list(suspect_row_ids)
    suspect_rows = {item["source_row"] for item in suspect_row_ids}
    kept_indexes = []
    snapshot_indexes = []
    duplicate_removed = 0
    replaced_by_newer = 0
    excluded_rows = 0
    multiple_fetch_times = frame["fetched_at"].dropna().nunique() > 1

    for job_id, group in frame.groupby("id", dropna=False):
        if not job_id or group["source_row"].isin(suspect_rows).all():
            excluded_rows += len(group)
            if not job_id:
                for row in group.itertuples(index=False):
                    exceptions.append(
                        {
                            "source_row": row.source_row,
                            "id": "",
                            "字段": "id",
                            "异常类型": "ID 缺失",
                            "原值": "",
                            "处理动作": "从岗位主表排除",
                        }
                    )
            continue
        usable = group.loc[~group["source_row"].isin(suspect_rows)].copy()
        excluded_rows += len(group) - len(usable)
        parsed_times = usable["fetched_at"].map(parse_fetched_at)
        if parsed_times.isna().any():
            excluded_rows += len(usable)
            exceptions.append(
                {
                    "source_row": "、".join(map(str, usable["source_row"])),
                    "id": job_id,
                    "字段": "fetched_at",
                    "异常类型": "抓取时间冲突",
                    "原值": "、".join(sorted({str(value) for value in usable["fetched_at"]})),
                    "处理动作": "无法判断先后，从岗位主表排除",
                }
            )
            continue
        usable = usable.assign(_fetched_dt=parsed_times)
        latest_time = usable["_fetched_dt"].max()
        latest = usable.loc[usable["_fetched_dt"].eq(latest_time)]
        conflict_fields = differing_fields(latest)
        if conflict_fields:
            excluded_rows += len(usable)
            exceptions.append(
                {
                    "source_row": "、".join(map(str, sorted(latest["source_row"]))),
                    "id": job_id,
                    "字段": "、".join(conflict_fields),
                    "异常类型": "同一抓取时间字段冲突",
                    "原值": "、".join(map(str, sorted(latest["source_row"]))),
                    "处理动作": "冲突涉及最新抓取，从岗位主表排除",
                }
            )
            continue
        latest_kept = latest.sort_values("source_row").iloc[0]
        kept_indexes.append(latest_kept.name)
        duplicate_removed += len(latest) - 1
        older = usable.loc[usable["_fetched_dt"].lt(latest_time)]
        replaced_by_newer += len(older)
        if multiple_fetch_times:
            for _, snapshot_group in usable.groupby("_fetched_dt"):
                snapshot_fields = differing_fields(snapshot_group)
                if snapshot_fields:
                    exceptions.append(
                        {
                            "source_row": "、".join(map(str, sorted(snapshot_group["source_row"]))),
                            "id": job_id,
                            "字段": "、".join(snapshot_fields),
                            "异常类型": "历史抓取字段冲突",
                            "原值": "、".join(map(str, sorted(snapshot_group["source_row"]))),
                            "处理动作": "不进入该抓取时间的岗位快照",
                        }
                    )
                    continue
                snapshot_indexes.append(snapshot_group.sort_values("source_row").iloc[0].name)
    stats = {
        "raw_rows": len(frame),
        "duplicate_removed": duplicate_removed,
        "replaced_by_newer": replaced_by_newer,
        "excluded_rows": excluded_rows,
    }
    main = frame.loc[kept_indexes].copy()
    snapshots = frame.loc[snapshot_indexes].copy() if multiple_fetch_times else pd.DataFrame()
    return main, snapshots, exceptions, stats, multiple_fetch_times


def differing_fields(group):
    fields = []
    for name in BUSINESS_COLUMNS:
        values = {None if is_blank(value) else str(value) for value in group[name]}
        if len(values) > 1:
            fields.append(name)
    return fields


def recruit_group_for(row):
    key = (
        row.portal_type or "",
        row.portal_name or "",
        row.recruit_type_id or "",
        row.recruit_type_name or "",
        row.recruit_parent_name or "",
    )
    group = RECRUIT_GROUPS.get(key)
    if group:
        return group, ""
    reason = (
        f"portal_type={key[0]}，portal_name={key[1]}，recruit_type_id={key[2]}，"
        f"recruit_type_name={key[3]}，recruit_parent_name={key[4]}"
    )
    return "待确认", reason


def project_status_for(row, society_subjects_all_empty):
    if row.job_subject_name:
        return "已填写"
    if row.portal_name == "社招" and society_subjects_all_empty:
        return "不适用"
    return "待确认"


def publish_status(publish_dt, fetched_dt, early_before):
    if publish_dt is None:
        return "缺失"
    if fetched_dt is not None and publish_dt > fetched_dt:
        return "晚于抓取时间"
    if publish_dt.date() < early_before:
        return "早期记录待复核"
    return "有效"


def enrich_main(main, category_rules, early_before):
    society = main.loc[main["portal_name"].eq("社招"), "job_subject_name"]
    society_subjects_all_empty = society.empty or society.isna().all()
    exceptions = []
    records = []
    for row in main.itertuples(index=False):
        group_name, group_reason = recruit_group_for(row)
        if group_name == "待确认":
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "recruit_group",
                    "异常类型": "招聘分组待确认",
                    "原值": group_reason,
                    "处理动作": "保留在主表，recruit_group=待确认",
                }
            )
        rule = category_rules.get(row.category_id or "", rules_fallback())
        if rule["status"] == "待确认" and (not rule["parent"] or "、" in rule["parent"]):
            if "、" in rule["name"] or "、" in rule["parent"]:
                exceptions.append(
                    {
                        "source_row": row.source_row,
                        "id": row.id,
                        "字段": "category_id",
                        "异常类型": "类别名称或上级冲突",
                        "原值": row.category_id or "",
                        "处理动作": "分析类别标为待确认",
                    }
                )
        fetched_dt = parse_fetched_at(row.fetched_at)
        publish_dt = None
        features = None
        status = "缺失"
        if is_blank(row.publish_time):
            status = "缺失"
        else:
            try:
                features = time_features_excel(int(str(row.publish_time)), unit="ms")
                publish_dt = time_features(int(str(row.publish_time)), unit="ms")["dt"]
                status = publish_status(publish_dt, fetched_dt, early_before)
            except (TypeError, ValueError):
                status = "格式错误"
                exceptions.append(
                    {
                        "source_row": row.source_row,
                        "id": row.id,
                        "字段": "publish_time",
                        "异常类型": "发布时间格式错误",
                        "原值": row.publish_time,
                        "处理动作": "保留岗位，时间字段留空",
                    }
                )
        if status == "晚于抓取时间":
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "publish_time",
                    "异常类型": "晚于抓取时间",
                    "原值": row.publish_time,
                    "处理动作": "保留岗位，时间图不纳入",
                }
            )
        elif status == "早期记录待复核":
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "publish_time",
                    "异常类型": "早期记录待复核",
                    "原值": features["date_only"].isoformat() if features else row.publish_time,
                    "处理动作": "保留在主表并标记，时间图不纳入",
                }
            )
        elif status == "缺失":
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "publish_time",
                    "异常类型": "发布时间缺失",
                    "原值": "",
                    "处理动作": "保留岗位，时间图不纳入",
                }
            )
        address_fields = parse_address_fields(row.address, row.city_name)
        item = {name: getattr(row, name) for name in MAIN_COLUMNS if hasattr(row, name)}
        item.update(
            {
                "recruit_group": group_name,
                "project_status": project_status_for(row, society_subjects_all_empty),
                "analysis_category_l1": rule["l1"],
                "analysis_category_l2": rule["l2"],
                "category_mapping_status": rule["status"],
                "publish_datetime": features["dt"] if features else None,
                "publish_date": features["date_only"] if features else None,
                "publish_month": features["month_start"] if features else None,
                "publish_time_status": status,
                "fetched_datetime": None if fetched_dt is None else fetched_dt.replace(tzinfo=None),
                "publish_year": features["year"] if features else None,
                "publish_quarter": features["quarter"] if features else None,
                "publish_day": features["day"] if features else None,
                "publish_day_name": features["day_name"] if features else None,
                "publish_is_weekend": features["is_weekend"] if features else None,
                "publish_iso_year": features["iso_year"] if features else None,
                "publish_weekofyear": features["weekofyear"] if features else None,
                "publish_week_start": features["week_start"] if features else None,
                "publish_dayofyear": features["dayofyear"] if features else None,
                "publish_ten_day_period": features["ten_day_period"] if features else None,
                **address_fields,
            }
        )
        records.append(item)
    result = pd.DataFrame.from_records(records, columns=MAIN_COLUMNS)
    return result, exceptions, society_subjects_all_empty


def rules_fallback():
    return {"l1": "待确认", "l2": "待确认", "status": "待确认", "name": "", "parent": ""}


def address_result(parsed, source):
    return {
        "address_province": parsed["province"],
        "address_city": parsed["city"],
        "address_district": parsed["district"],
        "address_township": parsed["township"],
        "address_adcode": parsed["adcode"],
        "address_detail": parsed["detail"],
        "address_status": parsed["status"],
        "address_parse_source": source,
        "address_candidate_codes": parsed["candidate_codes"] or None,
        "address_data_version": parsed["data_version"],
    }


def parse_address_fields(value, city_name=None):
    empty = {
        "address_province": None,
        "address_city": None,
        "address_district": None,
        "address_township": None,
        "address_adcode": None,
        "address_detail": None,
        "address_status": None,
        "address_parse_source": None,
        "address_candidate_codes": None,
        "address_data_version": None,
    }
    if is_blank(value):
        empty["address_status"] = "未提供"
        return empty
    text = INVISIBLE.sub("", str(value)).strip()
    parsed = parse_address_excel(text)
    source = "address"
    city = None if is_blank(city_name) else INVISIBLE.sub("", str(city_name)).strip()
    # 解析器只认「北京市」这类正式前缀，短名「北京」拼上去仍是 unmatched。
    if parsed["status"] == "unmatched" and city:
        trials = []
        if not text.startswith(city):
            trials.append((city + text, "city_name+address"))
        formal = city if city.endswith("市") else city + "市"
        if formal != city and not text.startswith(formal):
            trials.append((formal + text, "city_name市+address"))
        for trial_text, trial_source in trials:
            retried = parse_address_excel(trial_text)
            if retried["status"] != "unmatched":
                parsed = retried
                source = trial_source
                break
    return address_result(parsed, source)


def split_cities(value):
    if is_blank(value):
        return []
    seen = []
    for part in CITY_SPLIT.split(str(value)):
        city = clean_token(part)
        if city and city not in seen:
            seen.append(city)
    return seen


def build_cities(main, city_map):
    rows = []
    exceptions = []
    mismatch = 0
    for row in main.itertuples(index=False):
        names = split_cities(row.city_names)
        source = "city_names"
        if not names and row.city_name:
            names = [row.city_name]
            source = "city_name"
        if not names:
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "city_names",
                    "异常类型": "城市缺失",
                    "原值": "",
                    "处理动作": "岗位保留在主表，不进入岗位城市",
                }
            )
            continue
        if row.city_name and row.city_name not in names and source == "city_names":
            mismatch += 1
            exceptions.append(
                {
                    "source_row": row.source_row,
                    "id": row.id,
                    "字段": "city_name",
                    "异常类型": "表单城市不在城市列表中",
                    "原值": f"{row.city_name} / {row.city_names}",
                    "处理动作": "保留差异供核对，城市行仍按 city_names 拆分",
                }
            )
        for raw in names:
            mapped = city_map.get(raw)
            rows.append(
                {
                    "id": row.id,
                    "source_row": row.source_row,
                    "city_raw": raw,
                    "city_standard": mapped["city_standard"] if mapped else raw,
                    "province": mapped["province"] if mapped else None,
                    "country_or_region": mapped["country_or_region"] if mapped else None,
                    "geo_code": mapped["geo_code"] if mapped else None,
                    "city_mapping_status": "已映射" if mapped else "待映射",
                    "city_source": source,
                    "city_code": row.city_code,
                    "city_name": row.city_name,
                    "recruit_group": row.recruit_group,
                    "analysis_category_l1": row.analysis_category_l1,
                    "analysis_category_l2": row.analysis_category_l2,
                    "publish_month": row.publish_month,
                    "publish_time_status": row.publish_time_status,
                }
            )
    cities = pd.DataFrame.from_records(rows, columns=CITY_COLUMNS)
    return cities, exceptions, mismatch


def category_sheet(main, category_rules):
    counts = main.groupby("category_id").size().to_dict()
    rows = []
    for category_id, rule in sorted(category_rules.items(), key=lambda item: item[0]):
        if category_id not in counts and category_id not in set(main["category_id"].fillna("")):
            continue
        rows.append(
            {
                "原类别ID": category_id,
                "原名称": rule["name"],
                "原上级": rule["parent"],
                "分析大类": rule["l1"],
                "分析细类": rule["l2"],
                "映射状态": rule["status"],
                "岗位数": int(counts.get(category_id, 0)),
            }
        )
    return pd.DataFrame(rows)


def audit_sheets(main, cities):
    city_cross = (
        cities.groupby(["city_standard", "recruit_group"], dropna=False)["id"]
        .nunique()
        .reset_index(name="岗位数")
        .rename(columns={"city_standard": "标准城市", "recruit_group": "招聘分组"})
        .sort_values(["岗位数", "标准城市", "招聘分组"], ascending=[False, True, True])
    )
    category_cross = (
        main.groupby(["recruit_group", "analysis_category_l1", "analysis_category_l2"], dropna=False)["id"]
        .nunique()
        .reset_index(name="岗位数")
        .rename(
            columns={
                "recruit_group": "招聘分组",
                "analysis_category_l1": "岗位大类",
                "analysis_category_l2": "岗位细类",
            }
        )
        .sort_values(["招聘分组", "岗位数"], ascending=[True, False])
    )
    valid = main.loc[main["publish_time_status"].eq("有效")]
    month_cross = (
        valid.groupby(["publish_month", "recruit_group"], dropna=False)["id"]
        .nunique()
        .reset_index(name="岗位数")
        .rename(columns={"publish_month": "发布月", "recruit_group": "招聘分组"})
        .sort_values(["发布月", "招聘分组"])
    )
    month_cross["纳入规则"] = "仅统计 publish_time_status=有效。本次抓取岗位按发布月份分布"
    return city_cross, category_cross, month_cross


def build_confirmations(main, cities, categories, early_before, fetch_label):
    """把每次运行都要人工看的事项收成固定列，并按类型和标识排序。"""
    rows = []
    pending_categories = categories.loc[categories["映射状态"].eq("待确认")].sort_values("原类别ID")
    for row in pending_categories.itertuples(index=False):
        rows.append(
            {
                "确认类型": "类别是否为顶级类",
                "标识": row.原类别ID,
                "名称": row.原名称,
                "岗位数": int(row.岗位数),
                "当前处理": "分析大类标为待确认，细类保留原名称",
                "待决定事项": "确认是顶级类后，用 --category-map 指定分析大类和细类再运行",
            }
        )
    recruit = (
        main.loc[main["recruit_group"].eq("待确认")]
        .groupby(["recruit_type_id", "recruit_type_name", "portal_name"], dropna=False)
        .agg(岗位数=("id", "nunique"), 标识=("id", lambda values: "、".join(sorted(values))))
        .reset_index()
        .sort_values(["portal_name", "recruit_type_id", "recruit_type_name"])
    )
    for row in recruit.itertuples(index=False):
        rows.append(
            {
                "确认类型": "招聘分组",
                "标识": row.标识,
                "名称": f"{row.portal_name} / {row.recruit_type_name}",
                "岗位数": int(row.岗位数),
                "当前处理": "保留在主表，recruit_group=待确认",
                "待决定事项": "确认这种招聘类型应归入哪个分组",
            }
        )
    if cities.empty:
        city_groups = pd.DataFrame(columns=["city_raw", "岗位数"])
    else:
        city_groups = (
            cities.loc[cities["city_mapping_status"].eq("待映射")]
            .groupby("city_raw", dropna=False)["id"]
            .nunique()
            .reset_index(name="岗位数")
            .sort_values("city_raw")
        )
    for row in city_groups.itertuples(index=False):
        rows.append(
            {
                "确认类型": "城市标准名",
                "标识": row.city_raw or "",
                "名称": row.city_raw or "",
                "岗位数": int(row.岗位数),
                "当前处理": "未传 --city-map，标准城市等于原名，省份和地图编码为空",
                "待决定事项": "在城市映射中给出标准名、省份、国家或地区、地图编码",
            }
        )
    addresses = main.loc[~main["address_status"].isin(["matched", "未提供"])]
    address_groups = (
        addresses.groupby(["address_status", "address"], dropna=False)["id"]
        .nunique()
        .reset_index(name="岗位数")
        .sort_values(["address_status", "address"])
    )
    for row in address_groups.itertuples(index=False):
        rows.append(
            {
                "确认类型": "地址解析",
                "标识": row.address_status,
                "名称": row.address or "",
                "岗位数": int(row.岗位数),
                "当前处理": f"address_status={row.address_status}，原文保留",
                "待决定事项": "核对这条地址应归到哪个省市区",
            }
        )
    early_count = int(main["publish_time_status"].eq("早期记录待复核").sum())
    rows.append(
        {
            "确认类型": "早期发布时间阈值",
            "标识": early_before.isoformat(),
            "名称": f"早于 {early_before.isoformat()}",
            "岗位数": early_count,
            "当前处理": "保留在主表并标记早期记录待复核，不进入发布月核数。逐条明细在异常明细",
            "待决定事项": "确认阈值是否用 --early-before 调整",
        }
    )
    for field_name in ("薪资", "行业", "公司分级", "招聘人数", "目标值"):
        rows.append(
            {
                "确认类型": "待补字段",
                "标识": field_name,
                "名称": field_name,
                "岗位数": 0,
                "当前处理": "源表没有该字段，本次不生成",
                "待决定事项": "有可核验来源后再接入清洗",
            }
        )
    rows.append(
        {
            "确认类型": "跨期快照",
            "标识": fetch_label,
            "名称": "抓取批次",
            "岗位数": len(main),
            "当前处理": "只有一个抓取时间，不输出岗位快照",
            "待决定事项": "出现第二批抓取后，再判断能否做跨期对比",
        }
    )
    return pd.DataFrame(rows)[
        ["确认类型", "标识", "名称", "岗位数", "当前处理", "待决定事项"]
    ]


def field_dictionary():
    rows = [
        ("岗位主表", "source_row", "整数", "源 Excel 行号", "读取时的实际行号，表头为第 1 行", "无", "维度"),
        ("岗位主表", "id", "文本", "岗位 ID", "按文本读取和写出", "异常时不进入主表", "维度，去重计数作为岗位数"),
        ("岗位主表", "code", "文本", "职位编码", "按文本保留", "空", "维度"),
        ("岗位主表", "title", "文本", "岗位标题", "保留原文", "空", "维度"),
        ("岗位主表", "portal_type", "文本", "渠道编码", "2 社招，3 校招", "空", "维度"),
        ("岗位主表", "portal_name", "文本", "渠道名称", "去除首尾和不可见空白", "空", "维度"),
        ("岗位主表", "fetched_at", "文本", "原始抓取时间", "保留源字符串", "空", "维度"),
        ("岗位主表", "publish_time", "文本", "原始发布时间", "Unix 毫秒时间戳原文", "空", "维度"),
        ("岗位主表", "category_id", "文本", "职位类别 ID", "按文本保留", "空", "维度"),
        ("岗位主表", "category_name", "文本", "职位类别名称", "去除首尾和不可见空白", "空", "维度"),
        ("岗位主表", "category_parent_name", "文本", "上级类别", "去除首尾和不可见空白", "未填写上级", "维度"),
        ("岗位主表", "city_code", "文本", "表单城市编码", "源表单字段，不是多城市拆分结果", "空", "维度"),
        ("岗位主表", "city_name", "文本", "表单城市", "源表单字段。未确认为主城市，不直接做主城市占比", "空", "维度"),
        ("岗位主表", "city_names", "文本", "城市列表原文", "保留原文，拆分见岗位城市", "空", "维度"),
        ("岗位主表", "recruit_type_id", "文本", "招聘类型编码", "按文本保留", "空", "维度"),
        ("岗位主表", "recruit_type_name", "文本", "招聘类型", "正式、实习等", "空", "维度"),
        ("岗位主表", "recruit_parent_name", "文本", "招聘渠道上级", "应与 portal_name 一致", "空", "维度"),
        ("岗位主表", "job_subject_name", "文本", "招聘项目", "校招项目名称", "社招在本次全量中为空", "维度"),
        ("岗位主表", "description", "文本", "岗位职责", "保留原文", "空", "文本"),
        ("岗位主表", "requirement", "文本", "岗位要求", "保留原文", "空", "文本"),
        ("岗位主表", "address", "文本", "地址原文", "保留原文", "空", "文本"),
        ("岗位主表", "recruit_group", "文本", "招聘分组", "由渠道和招聘类型交叉核对得到", "待确认表示组合未识别或字段冲突", "维度"),
        ("岗位主表", "project_status", "文本", "项目填写状态", "已填写、不适用、待确认", "无", "维度"),
        ("岗位主表", "analysis_category_l1", "文本", "分析大类", "有上级时用上级；未确认的顶级类为待确认", "无", "维度"),
        ("岗位主表", "analysis_category_l2", "文本", "分析细类", "有上级时用类别名；待确认时保留原类别名", "无", "维度"),
        ("岗位主表", "category_mapping_status", "文本", "类别映射状态", "已映射、同 ID 补齐、人工映射或待确认", "无", "维度"),
        ("岗位主表", "publish_datetime", "日期时间", "发布时间", "毫秒时间戳先按 UTC 再转北京时间，Excel 不带时区标记", "时间无效时为空", "时间维度"),
        ("岗位主表", "publish_date", "日期", "发布日期", "北京时间的日期", "时间无效时为空", "时间维度"),
        ("岗位主表", "publish_month", "日期", "发布月", "月初日期，由 date-dimensions 的 month_start 得到", "时间无效时为空", "时间维度"),
        ("岗位主表", "publish_time_status", "文本", "发布时间状态", "有效、缺失、格式错误、晚于抓取时间、早期记录待复核", "无", "维度"),
        ("岗位主表", "fetched_datetime", "日期时间", "抓取时间", "ISO 时间统一到北京时间后去掉时区标记", "空", "时间维度"),
        ("岗位主表", "publish_year", "整数", "发布年", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_quarter", "整数", "发布季度", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_day", "整数", "发布日", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_day_name", "文本", "星期", "date-dimensions，英文星期名", "时间无效时为空", "维度"),
        ("岗位主表", "publish_is_weekend", "文本", "是否周末", "date-dimensions 的 Excel 取值：是或否", "时间无效时为空", "维度"),
        ("岗位主表", "publish_iso_year", "整数", "ISO 年", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_weekofyear", "整数", "ISO 周", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_week_start", "日期", "周起始日", "date-dimensions，周一", "时间无效时为空", "时间维度"),
        ("岗位主表", "publish_dayofyear", "整数", "年内第几天", "date-dimensions", "时间无效时为空", "维度"),
        ("岗位主表", "publish_ten_day_period", "文本", "旬", "上旬、中旬、下旬", "时间无效时为空", "维度"),
        ("岗位主表", "address_province", "文本", "地址省", "address-dimensions", "未解析到", "维度"),
        ("岗位主表", "address_city", "文本", "地址地级", "address-dimensions。直辖市通常为空", "未解析到", "维度"),
        ("岗位主表", "address_district", "文本", "地址区县", "address-dimensions", "未解析到", "维度"),
        ("岗位主表", "address_township", "文本", "地址乡镇", "address-dimensions", "未解析到", "维度"),
        ("岗位主表", "address_adcode", "文本", "区划代码", "命中终点的区划代码，按文本保存", "未确定", "维度"),
        ("岗位主表", "address_detail", "文本", "地址剩余详情", "路名和门牌等", "空", "文本"),
        ("岗位主表", "address_status", "文本", "地址解析状态", "matched、ambiguous、conflict、historical、unmatched、未提供", "无", "维度"),
        ("岗位主表", "address_parse_source", "文本", "地址解析输入", "address 表示原文；city_name+address 表示短名直接拼接成功；city_name市+address 表示短名补「市」后成功", "未解析", "维度"),
        ("岗位主表", "address_candidate_codes", "文本", "地址候选区划", "歧义时用顿号连接", "无歧义", "维度"),
        ("岗位主表", "address_data_version", "文本", "区划数据版本", "address-dimensions 的快照日期", "空", "维度"),
        ("岗位城市", "id", "文本", "岗位 ID", "回连岗位主表", "无", "维度，去重计数作为岗位数"),
        ("岗位城市", "source_row", "整数", "源 Excel 行号", "来自主表保留的那一行", "无", "维度"),
        ("岗位城市", "city_raw", "文本", "拆分出的原城市", "按顿号、中英文逗号和分号拆分，不去空格拆分", "无", "维度"),
        ("岗位城市", "city_standard", "文本", "标准城市", "没有映射时等于原城市", "无", "维度"),
        ("岗位城市", "province", "文本", "省份", "只来自城市映射", "待映射时为空", "维度"),
        ("岗位城市", "country_or_region", "文本", "国家或地区", "只来自城市映射", "待映射时为空", "维度"),
        ("岗位城市", "geo_code", "文本", "地图编码", "只来自城市映射", "待映射时为空", "维度"),
        ("岗位城市", "city_mapping_status", "文本", "城市映射状态", "已映射或待映射", "无", "维度"),
        ("岗位城市", "city_source", "文本", "城市来源", "city_names，为空时回退 city_name", "无", "维度"),
        ("岗位城市", "city_code", "文本", "表单城市编码", "源字段，重复在每个城市行上", "空", "维度"),
        ("岗位城市", "city_name", "文本", "表单城市", "源字段，未确认为主城市", "空", "维度"),
        ("岗位城市", "recruit_group", "文本", "招聘分组", "从主表带入", "待确认", "维度"),
        ("岗位城市", "analysis_category_l1", "文本", "分析大类", "从主表带入", "待确认", "维度"),
        ("岗位城市", "analysis_category_l2", "文本", "分析细类", "从主表带入", "待确认", "维度"),
        ("岗位城市", "publish_month", "日期", "发布月", "从主表带入", "时间无效时为空", "时间维度"),
        ("岗位城市", "publish_time_status", "文本", "发布时间状态", "从主表带入", "无", "维度"),
        ("城市招聘交叉核数", "标准城市", "文本", "标准城市", "来自岗位城市", "无", "维度"),
        ("城市招聘交叉核数", "招聘分组", "文本", "招聘分组", "来自岗位城市", "待确认", "维度"),
        ("城市招聘交叉核数", "岗位数", "整数", "去重岗位数", "组内 ID 去重", "无", "指标"),
        ("类别层级核数", "招聘分组", "文本", "招聘分组", "来自岗位主表", "待确认", "维度"),
        ("类别层级核数", "岗位大类", "文本", "分析大类", "来自岗位主表", "待确认", "维度"),
        ("类别层级核数", "岗位细类", "文本", "分析细类", "来自岗位主表", "待确认", "维度"),
        ("类别层级核数", "岗位数", "整数", "去重岗位数", "主表 ID 去重", "无", "指标"),
        ("发布月核数", "发布月", "日期", "发布月月初", "只含发布时间状态为有效的岗位", "无", "时间维度"),
        ("发布月核数", "招聘分组", "文本", "招聘分组", "来自岗位主表", "待确认", "维度"),
        ("发布月核数", "岗位数", "整数", "去重岗位数", "有效发布时间内的 ID 去重", "无", "指标"),
        ("发布月核数", "纳入规则", "文本", "口径说明", "固定说明本次只纳入有效发布时间", "无", "文本"),
        ("异常明细", "source_row", "文本", "源行号", "可能包含多行", "无", "维度"),
        ("异常明细", "id", "文本", "岗位 ID", "能回查时填写", "ID 缺失时为空", "维度"),
        ("异常明细", "字段", "文本", "异常字段", "一个或多个字段名", "无", "维度"),
        ("异常明细", "异常类型", "文本", "异常类型", "清洗规则命中的类型", "无", "维度"),
        ("异常明细", "原值", "文本", "原值摘要", "便于回查，不改写源文件", "空", "文本"),
        ("异常明细", "处理动作", "文本", "处理动作", "保留、排除或不纳入时间图", "无", "文本"),
        ("待确认清单", "确认类型", "文本", "需要人工决定的类别", "每次运行按固定规则生成", "无", "维度"),
        ("待确认清单", "标识", "文本", "类别 ID、岗位 ID、城市或阈值", "同一类事项内按标识排序", "空", "维度"),
        ("待确认清单", "名称", "文本", "展示名称", "类别名、城市名或地址原文", "空", "文本"),
        ("待确认清单", "岗位数", "整数", "涉及的去重岗位数", "待补字段固定为 0", "无", "指标"),
        ("待确认清单", "当前处理", "文本", "本次清洗已经采取的处理", "不依赖人工备注", "无", "文本"),
        ("待确认清单", "待决定事项", "文本", "下一次运行前可以改的决定", "对应命令参数或映射文件", "无", "文本"),
    ]
    return pd.DataFrame(rows, columns=["表名", "字段名", "类型", "含义", "生成规则", "空值含义", "QuickBI配置"])


def write_workbook(path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets:
            frame.to_excel(writer, sheet_name=name, index=False)
    workbook = load_workbook(path)
    for name, frame in sheets:
        sheet = workbook[name]
        headers = list(frame.columns)
        for column_name in TEXT_COLUMNS + ("address_adcode", "geo_code", "原类别ID"):
            if column_name not in headers:
                continue
            column_index = headers.index(column_name) + 1
            for row_index in range(2, sheet.max_row + 1):
                sheet.cell(row=row_index, column=column_index).number_format = "@"
        for column_name, number_format in DATE_FORMATS.items():
            if column_name not in headers:
                continue
            column_index = headers.index(column_name) + 1
            letter = get_column_letter(column_index)
            for row_index in range(2, sheet.max_row + 1):
                sheet.cell(row=row_index, column=column_index).number_format = number_format
            sheet.column_dimensions[letter].width = 22
    workbook.save(path)


def missing_counts(frame):
    counts = {}
    for name in EXPECTED_COLUMNS:
        if name not in frame.columns:
            continue
        counts[name] = int(frame[name].map(is_blank).sum())
    return counts


def same_code_multiple_ids(frame):
    usable = frame.loc[frame["code"].notna() & frame["id"].notna()]
    counts = usable.groupby("code")["id"].nunique()
    return counts.loc[counts.gt(1)]


def validate(main, cities, raw_stats):
    errors = []
    if main["id"].isna().any() or main["id"].eq("").any():
        errors.append("主表存在空 ID")
    if main["id"].duplicated().any():
        errors.append("主表 ID 不唯一")
    explained = (
        len(main)
        + raw_stats["duplicate_removed"]
        + raw_stats["replaced_by_newer"]
        + raw_stats["excluded_rows"]
    )
    if explained != raw_stats["raw_rows"]:
        errors.append(f"行数差额无法解释: 原始 {raw_stats['raw_rows']}，解释后 {explained}")
    if not cities.empty:
        if cities.duplicated(["id", "city_standard"]).any():
            errors.append("岗位城市的 id + 标准城市 有重复")
        missing_link = set(cities["id"]) - set(main["id"])
        if missing_link:
            errors.append("岗位城市存在无法回连主表的 ID")
    if int(main.groupby("recruit_group")["id"].nunique().sum()) != len(main):
        errors.append("招聘分组岗位数加总不等于主表")
    category_total = int(
        main.groupby(["analysis_category_l1", "analysis_category_l2"])["id"].nunique().sum()
    )
    if category_total != len(main):
        errors.append("类别层级岗位数加总不等于主表")
    return errors


def verify_workbook(path, expected_rows, sample_id):
    workbook = load_workbook(path, data_only=False)
    sheet = workbook["岗位主表"]
    headers = [cell.value for cell in sheet[1]]
    id_index = headers.index("id") + 1
    date_index = headers.index("publish_datetime") + 1
    found = None
    for row_index in range(2, sheet.max_row + 1):
        if sheet.cell(row=row_index, column=id_index).value == sample_id:
            found = sheet.cell(row=row_index, column=date_index).value
            break
    counts = {name: workbook[name].max_row - 1 for name in expected_rows}
    workbook.close()
    problems = []
    if found is None:
        problems.append("回读时没有找到抽样 ID")
    elif not isinstance(found, datetime):
        problems.append(f"publish_datetime 不是日期类型: {type(found).__name__}")
    for name, count in expected_rows.items():
        if counts.get(name) != count:
            problems.append(f"{name} 行数 {counts.get(name)}，期望 {count}")
    return problems, counts


def write_report(path, context):
    missing = context["missing"]
    stats = context["stats"]
    lines = [
        "# 职位清洗说明",
        "",
        "## 输入和范围",
        "",
        f"- 输入文件：`{context['input_path']}`",
        f"- 工作表：`{SOURCE_SHEET}`",
        f"- 本次读取 {stats['raw_rows']} 行、{len(EXPECTED_COLUMNS) + len(context['extra'])} 列。",
        f"- 新增字段：{('、'.join(context['extra']) if context['extra'] else '无')}",
        f"- 整格缺失占位符：{context['placeholder_text']}",
        f"- 早期复核阈值：发布日期早于 {context['early_before'].isoformat()}（北京时间）。",
        f"- 抓取批次：{context['fetch_batches']} 批。{context['batch_text']}",
        "",
        "## 关键规则",
        "",
        "- `id`、`code`、`category_id`、`city_code`、`portal_type`、`recruit_type_id` 按文本处理。超过 15 位的 Excel 数值视为精度可疑，不猜测原值。",
        "- `sub_title` 全量缺失，未放入岗位主表。",
        "- 分类字段去掉首尾空白、全角空格和不可见空白。标题、职责、要求和地址保留原文。",
        "- 招聘分组只确认三种组合：社招正式、校招正式、校招实习。其余标为待确认。",
        "- 社招项目名称在全量中均为空，`project_status` 标为不适用。校招均已填写。",
        "- 职位类别默认读取 `Chapter1/data/job_categories.csv`。有上级时，大类用上级、细类用自身，状态为类别表。没有上级时视为顶级类，大类和细类都用自身，状态为类别表顶级。`--category-map` 仍可覆盖。",
        "- 发布时间使用 `date-dimensions`：毫秒时间戳换算为北京时间，并展开年、季、周、旬等字段。未启用可选的中国节假日包。",
        "- 地址使用 `address-dimensions` 的 `parse_address_excel`。原文为 unmatched 时，先试 city_name 加 address。短名不以「市」结尾且仍 unmatched 时，再试 city_name 加「市」加 address。",
        "- 城市按 `、`、中英文逗号和分号拆分。本次没有传入城市映射，标准城市等于原城市，状态为待映射。",
        "- 只有一个抓取时间，因此不输出岗位快照。时间图口径是「本次抓取岗位按发布月份分布」。",
        "",
        "## 全量核数",
        "",
        f"- 原始行数：{stats['raw_rows']}",
        f"- 合并的完全重复行：{stats['duplicate_removed']}",
        f"- 被较新抓取替代的历史行：{stats['replaced_by_newer']}",
        f"- 因异常排除的行：{stats['excluded_rows']}",
        f"- 岗位主表：{context['main_rows']}",
        f"- 差额：{stats['raw_rows']} = {context['main_rows']} + {stats['duplicate_removed']} + {stats['replaced_by_newer']} + {stats['excluded_rows']}",
        f"- 主表 ID 唯一且非空：{'是' if context['id_ok'] else '否'}",
        f"- 岗位城市组合：{context['city_rows']}。没有城市的岗位：{context['jobs_without_city']}",
        f"- 城市待映射：{context['unmapped_cities']} 行。表单城市不在城市列表中：{context['city_mismatch']}",
        f"- 招聘分组加总：{context['recruit_total']}，等于主表 {context['main_rows']}。",
        f"- 类别层级加总：{context['category_total']}，等于主表 {context['main_rows']}。",
        f"- 多城市抽查：岗位 `{context['multi_city_id']}` 在主表 1 行，在岗位城市 {context['multi_city_count']} 行。",
        f"- 有效发布时间范围：{context['valid_range']}",
        f"- 时间状态：{context['time_status']}",
        f"- 地址解析状态：{context['address_status']}",
        f"- 地址原文 unmatched 后，用城市名重解析成功：{context['address_retry_recovered']} 条。来源：{context['address_parse_source']}",
        f"- 类别表映射状态：{context['category_status']}",
        f"- 同一 code 对应多个 id：{context['code_conflicts']}",
        f"- 长 ID 数值单元格：{context['numeric_cells']}",
        "",
        "各列空值数：",
        "",
        "| 字段 | 空值数 |",
        "| --- | --- |",
    ]
    for name, count in missing.items():
        lines.append(f"| {name} | {count} |")
    lines.extend(
        [
            "",
            "## 复现",
            "",
            "使用同一份输入、同一版本的清洗脚本和相同参数，工作表里的单元格值保持一致。命令是：",
            "",
            "```bash",
            "python Chapter2/clean_jobs.py --input Chapter1/data/jobs.xlsx --output-dir Chapter2",
            "```",
            "",
            "待确认事项不在对话里临时列举，每次都完整写入工作表 `待确认清单`，并按确认类型、标识、名称排序。异常明细同样全量写入，不在报告里截断。",
            "",
            "## 待确认清单",
            "",
            "| 确认类型 | 标识 | 名称 | 岗位数 | 当前处理 | 待决定事项 |",
            "| --- | --- | --- | --- | --- | --- |",
        ]
    )
    for row in context["confirmations"].itertuples(index=False):
        name = str(row.名称).replace("|", "\\|")
        lines.append(
            f"| {row.确认类型} | {row.标识} | {name} | {int(row.岗位数)} | {row.当前处理} | {row.待决定事项} |"
        )
    lines.extend(
        [
            "",
            "## 待补数据",
            "",
            "- 源表没有可核验的薪资、行业、公司分级、招聘人数和目标值。",
            "- 进度条目标值没有来源，不能在 QuickBI 里直接做完成率。",
            "- 本次只有一个抓取时间，不能判断跨期快照是否完整。",
            "",
            "## QuickBI 首张图",
            "",
            "用 `岗位城市` 做交叉表：行拖「标准城市」，列拖「招聘分组」，值拖 `id` 并选择去重计数。岗位会在多个城市各计一次，跨城市总计需要重新对 ID 去重。",
            "",
            "总岗位 KPI 用 `岗位主表` 的 `id` 去重计数。",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8")


def main():
    args = parse_args()
    early_before = date.fromisoformat(args.early_before)
    frame, extra, suspect_ids, numeric_cells, placeholder_fields = load_source(args.input)
    manual_categories = load_category_map(args.category_map)
    city_map = load_city_map(args.city_map)
    category_catalog = load_category_catalog(args.category_catalog)
    category_rules, category_exceptions = build_category_rules(
        frame, manual_categories, category_catalog
    )
    main_source, snapshots, choice_exceptions, stats, multiple_fetches = choose_rows(
        frame, suspect_ids
    )
    main, enrich_exceptions, society_empty = enrich_main(main_source, category_rules, early_before)
    main = main.sort_values(["source_row", "id"], kind="mergesort").reset_index(drop=True)
    cities, city_exceptions, city_mismatch = build_cities(main, city_map)
    if not cities.empty:
        cities = cities.sort_values(["source_row", "city_standard", "id"], kind="mergesort").reset_index(drop=True)
    exceptions = choice_exceptions + enrich_exceptions + city_exceptions
    if category_exceptions:
        exceptions.extend(
            item for item in category_exceptions if item["异常类型"] == "类别名称或上级冲突"
        )
    exception_frame = pd.DataFrame(
        exceptions,
        columns=["source_row", "id", "字段", "异常类型", "原值", "处理动作"],
    )
    if not exception_frame.empty:
        exception_frame = exception_frame.sort_values(
            ["异常类型", "source_row", "id"], kind="mergesort"
        ).reset_index(drop=True)
    categories = category_sheet(main, category_rules)
    city_cross, category_cross, month_cross = audit_sheets(main, cities)
    dictionary = field_dictionary()
    fetch_label = "、".join(sorted({str(value) for value in frame["fetched_at"].dropna()})) or "(空)"
    confirmations = build_confirmations(main, cities, categories, early_before, fetch_label)
    checks = validate(main, cities, stats)
    if checks:
        raise SystemExit("清洗校验失败: " + "；".join(checks))

    sheets = [
        ("岗位主表", main),
        ("岗位城市", cities),
        ("分类映射", categories),
        ("城市招聘交叉核数", city_cross),
        ("类别层级核数", category_cross),
        ("发布月核数", month_cross),
        ("异常明细", exception_frame),
        ("待确认清单", confirmations),
        ("字段字典", dictionary),
    ]
    if multiple_fetches and not snapshots.empty:
        sheets.insert(1, ("岗位快照", snapshots))
    output_dir = args.output_dir
    workbook_path = output_dir / "jobs_quickbi_cleaned.xlsx"
    report_path = output_dir / "jobs_cleaning_report.md"
    write_workbook(workbook_path, sheets)
    expected = {name: len(table) for name, table in sheets}
    sample_id = main.iloc[0]["id"]
    problems, _ = verify_workbook(workbook_path, expected, sample_id)
    if problems:
        raise SystemExit("回读校验失败: " + "；".join(problems))

    valid = main.loc[main["publish_time_status"].eq("有效"), "publish_date"]
    valid_range = "无有效发布时间"
    if not valid.empty:
        valid_range = f"{min(valid).isoformat()} 至 {max(valid).isoformat()}"
    fetch_counts = frame["fetched_at"].fillna("(空)").value_counts().sort_index()
    batch_text = "；".join(f"{name} {count} 行" for name, count in fetch_counts.items())
    if len(fetch_counts) == 1:
        batch_text += "。时间图统一表述为「本次抓取岗位按发布月份分布」。"
    code_conflicts = same_code_multiple_ids(frame)
    multi = cities.groupby("id", sort=True).size()
    multi = multi.loc[multi.gt(1)].sort_index()
    if multi.empty:
        multi_id = main.iloc[0]["id"]
        multi_count = 1
    else:
        multi_source = cities.loc[cities["id"].isin(multi.index)].groupby("id")["source_row"].min().sort_values()
        multi_id = multi_source.index[0]
        multi_count = int(multi.loc[multi_id])
    write_report(
        report_path,
        {
            "input_path": args.input,
            "extra": extra,
            "placeholder_text": (
                "、".join(f"{name} {count} 格" for name, count in sorted(placeholder_fields.items()))
                if placeholder_fields
                else "没有发现整格占位符"
            ),
            "early_before": early_before,
            "fetch_batches": int(fetch_counts.shape[0]),
            "batch_text": batch_text,
            "missing": missing_counts(frame),
            "stats": stats,
            "main_rows": len(main),
            "id_ok": True,
            "city_rows": len(cities),
            "jobs_without_city": int(main["id"].nunique() - cities["id"].nunique()) if len(cities) else len(main),
            "unmapped_cities": int(cities["city_mapping_status"].eq("待映射").sum()) if len(cities) else 0,
            "city_mismatch": city_mismatch,
            "recruit_total": int(main.groupby("recruit_group")["id"].nunique().sum()),
            "category_total": int(
                main.groupby(["analysis_category_l1", "analysis_category_l2"])["id"].nunique().sum()
            ),
            "multi_city_id": multi_id,
            "multi_city_count": multi_count,
            "valid_range": valid_range,
            "time_status": "，".join(
                f"{name} {count}"
                for name, count in main["publish_time_status"].value_counts().sort_index().items()
            ),
            "address_status": "，".join(
                f"{name} {count}"
                for name, count in main["address_status"].value_counts().sort_index().items()
            ),
            "address_retry_recovered": int(
                main["address_parse_source"].isin(["city_name+address", "city_name市+address"]).sum()
            ),
            "address_parse_source": "，".join(
                f"{name} {count}"
                for name, count in main["address_parse_source"].value_counts().sort_index().items()
            ),
            "category_status": "，".join(
                f"{name} {count}"
                for name, count in main["category_mapping_status"].value_counts().sort_index().items()
            ),
            "code_conflicts": "无" if code_conflicts.empty else "、".join(code_conflicts.index.astype(str)),
            "numeric_cells": (
                "无"
                if not numeric_cells
                else f"{len(numeric_cells)} 个，其中超过 15 位的 {len(suspect_ids)} 个"
            ),
            "confirmations": confirmations,
        },
    )
    print(f"已写入 {workbook_path}", flush=True)
    print(f"已写入 {report_path}", flush=True)
    print(
        f"主表 {len(main)} 行，岗位城市 {len(cities)} 行，异常 {len(exception_frame)} 行",
        flush=True,
    )


if __name__ == "__main__":
    main()
