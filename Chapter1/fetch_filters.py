"""抓取社招、校招的职位筛选项，分开存成类别、城市和招聘项目。

GET /api/v1/config/job/filters/{portal_type}
路径里的 2 是社招，3 是校招，两套结果不同。
Cookie 和 _signature 不是必需的。
"""

import csv
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

FILTERS_URL = "https://jobs.bytedance.com/api/v1/config/job/filters/{portal_type}"
DATA_DIR = Path(__file__).resolve().parent / "data"

PORTALS = (
    {
        "portal_type": 2,
        "portal_name": "社招",
        "channel": "office",
        "website_path": "society",
        "referer": "https://jobs.bytedance.com/experienced/position",
    },
    {
        "portal_type": 3,
        "portal_name": "校招",
        "channel": "campus",
        "website_path": "campus",
        "referer": "https://jobs.bytedance.com/campus/position",
    },
)

CATEGORY_COLUMNS = (
    "portal_type",
    "portal_name",
    "id",
    "name",
    "en_name",
    "depth",
    "parent_id",
    "filter_count",
    "fetched_at",
)
CITY_COLUMNS = (
    "portal_type",
    "portal_name",
    "code",
    "name",
    "en_name",
    "py_name",
    "fetched_at",
)
SUBJECT_COLUMNS = (
    "portal_type",
    "portal_name",
    "id",
    "name",
    "subject_group_id",
    "subject_group_name",
    "active_status",
    "fetched_at",
)


def text(value):
    if value is None:
        return ""
    return str(value)


def fetch_filters(portal):
    response = requests.get(
        FILTERS_URL.format(portal_type=portal["portal_type"]),
        headers={
            "Accept": "application/json, text/plain, */*",
            "Referer": portal["referer"],
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/156.0.0.0 Safari/537.36"
            ),
            "accept-language": "zh-CN",
            "Portal-Channel": portal["channel"],
            "Portal-Platform": "pc",
            "website-path": portal["website_path"],
        },
        timeout=30,
    )
    response.raise_for_status()
    body = response.json()
    if body.get("code") != 0:
        raise RuntimeError(body.get("message") or body)
    return body.get("data") or {}


def flatten_categories(nodes, parent_id=""):
    rows = []
    for node in nodes or []:
        category_id = text(node.get("id"))
        rows.append((node, parent_id))
        rows.extend(flatten_categories(node.get("children"), category_id))
    return rows


def category_rows(portal, data, fetched_at):
    counts = data.get("job_type_count_map") or {}
    rows = []
    for node, parent_id in flatten_categories(data.get("job_type_list")):
        category_id = text(node.get("id"))
        count = counts.get(category_id)
        rows.append(
            {
                "portal_type": portal["portal_type"],
                "portal_name": portal["portal_name"],
                "id": category_id,
                "name": text(node.get("name")),
                "en_name": text(node.get("en_name")),
                "depth": text(node.get("depth")),
                "parent_id": parent_id,
                "filter_count": text(count),
                "fetched_at": fetched_at,
            }
        )
    return rows


def city_rows(portal, data, fetched_at):
    rows = []
    for city in data.get("city_list") or []:
        rows.append(
            {
                "portal_type": portal["portal_type"],
                "portal_name": portal["portal_name"],
                "code": text(city.get("code")),
                "name": text(city.get("name")),
                "en_name": text(city.get("en_name")),
                "py_name": text(city.get("py_name")),
                "fetched_at": fetched_at,
            }
        )
    return rows


def subject_name(value):
    if isinstance(value, dict):
        return text(value.get("zh_cn") or value.get("i18n"))
    return text(value)


def subject_rows(portal, data, fetched_at):
    rows = []
    for subject in data.get("job_subject_list") or []:
        group = subject.get("subject_group_info") or {}
        rows.append(
            {
                "portal_type": portal["portal_type"],
                "portal_name": portal["portal_name"],
                "id": text(subject.get("id")),
                "name": subject_name(subject.get("name")),
                "subject_group_id": text(group.get("id")),
                "subject_group_name": text(group.get("name")),
                "active_status": text(subject.get("active_status")),
                "fetched_at": fetched_at,
            }
        )
    return rows


def write_csv(path, columns, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def main():
    fetched_at = datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")
    categories = []
    cities = []
    subjects = []
    for portal in PORTALS:
        data = fetch_filters(portal)
        categories.extend(category_rows(portal, data, fetched_at))
        cities.extend(city_rows(portal, data, fetched_at))
        subjects.extend(subject_rows(portal, data, fetched_at))

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    outputs = (
        (DATA_DIR / "job_categories.csv", CATEGORY_COLUMNS, categories),
        (DATA_DIR / "cities.csv", CITY_COLUMNS, cities),
        (DATA_DIR / "job_subjects.csv", SUBJECT_COLUMNS, subjects),
    )
    for path, columns, rows in outputs:
        count = write_csv(path, columns, rows)
        print(f"已写入 {path}，{count} 条，抓取时间 {fetched_at}", flush=True)


if __name__ == "__main__":
    main()
