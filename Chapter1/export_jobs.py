"""把已抓取的职位导出为 CSV 和 Excel。"""

import csv
import json
import re
from pathlib import Path

from openpyxl import Workbook

DATA_DIR = Path(__file__).resolve().parent / "data"
JSONL_PATH = DATA_DIR / "jobs.jsonl"
CSV_PATH = DATA_DIR / "jobs.csv"
EXCEL_PATH = DATA_DIR / "jobs.xlsx"
ILLEGAL_EXCEL_CHARS = re.compile(r"[\000-\010\013\014\016-\037]")

COLUMNS = (
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


def iter_jobs(path):
    with path.open(encoding="utf-8") as source:
        for line in source:
            line = line.strip()
            if line:
                yield json.loads(line)


def text(value):
    if value is None:
        return ""
    return str(value)


def subject_name(subject):
    if not isinstance(subject, dict):
        return ""
    name = subject.get("name")
    if isinstance(name, dict):
        return name.get("zh_cn") or name.get("i18n") or ""
    return text(name)


def to_row(job):
    category = job.get("job_category") or {}
    category_parent = category.get("parent") or {}
    city = job.get("city_info") or {}
    city_names = []
    for item in job.get("city_list") or []:
        name = (item or {}).get("name")
        if name and name not in city_names:
            city_names.append(name)
    recruit = job.get("recruit_type") or {}
    recruit_parent = recruit.get("parent") or {}
    post = job.get("job_post_info") or {}
    return {
        "id": text(job.get("id")),
        "code": text(job.get("code")),
        "title": text(job.get("title")),
        "sub_title": text(job.get("sub_title")),
        "portal_type": text(job.get("portal_type")),
        "portal_name": text(job.get("portal_name")),
        "fetched_at": text(job.get("fetched_at")),
        "publish_time": text(job.get("publish_time")),
        "category_id": text(category.get("id")),
        "category_name": text(category.get("name")),
        "category_parent_name": text(category_parent.get("name")),
        "city_code": text(city.get("code")),
        "city_name": text(city.get("name")),
        "city_names": "、".join(city_names),
        "recruit_type_id": text(recruit.get("id")),
        "recruit_type_name": text(recruit.get("name")),
        "recruit_parent_name": text(recruit_parent.get("name")),
        "job_subject_name": subject_name(job.get("job_subject")),
        "description": text(job.get("description")),
        "requirement": text(job.get("requirement")),
        "address": text(post.get("address")),
    }


def write_csv(path, jobs):
    count = 0
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=COLUMNS)
        writer.writeheader()
        for job in jobs:
            writer.writerow(to_row(job))
            count += 1
    return count


def write_excel(path, jobs):
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("职位")
    sheet.append(list(COLUMNS))
    count = 0
    for job in jobs:
        row = to_row(job)
        sheet.append([ILLEGAL_EXCEL_CHARS.sub("", row[name]) for name in COLUMNS])
        count += 1
    workbook.save(path)
    return count


def main():
    if not JSONL_PATH.exists():
        raise SystemExit(f"找不到抓取结果: {JSONL_PATH}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    csv_count = write_csv(CSV_PATH, iter_jobs(JSONL_PATH))
    excel_count = write_excel(EXCEL_PATH, iter_jobs(JSONL_PATH))
    print(
        f"已写入 {CSV_PATH}，{csv_count} 条，{CSV_PATH.stat().st_size} 字节",
        flush=True,
    )
    print(
        f"已写入 {EXCEL_PATH}，{excel_count} 条，{EXCEL_PATH.stat().st_size} 字节",
        flush=True,
    )


if __name__ == "__main__":
    main()
