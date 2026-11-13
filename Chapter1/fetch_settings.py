"""抓取招聘站共用枚举，单独存成字典表。

GET /api/v1/common/setting 不返回职位。社招和校招内容相同，
这里只请求一次。Cookie 和 _signature 不是必需的。
"""

import csv
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

SETTING_URL = "https://jobs.bytedance.com/api/v1/common/setting"
OUTPUT_PATH = Path(__file__).resolve().parent / "data" / "settings.csv"
COLUMNS = (
    "dict_name",
    "key",
    "label",
    "label_en",
    "country",
    "py_country_name",
    "fetched_at",
)


def fetch_setting():
    response = requests.get(
        SETTING_URL,
        headers={
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://jobs.bytedance.com/experienced/position",
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/156.0.0.0 Safari/537.36"
            ),
            "accept-language": "zh-CN",
            "Portal-Channel": "office",
            "Portal-Platform": "pc",
            "website-path": "society",
        },
        timeout=30,
    )
    response.raise_for_status()
    body = response.json()
    if body.get("code") != 0:
        raise RuntimeError(body.get("message") or body)
    return body.get("data") or {}


def text(value):
    if value is None:
        return ""
    return str(value)


def to_row(dict_name, item, fetched_at):
    value = item.get("val")
    label = ""
    label_en = ""
    if isinstance(value, dict):
        label = text(value.get("zh_cn") or value.get("i18n"))
        label_en = text(value.get("en_us"))
    else:
        label = text(value)
    return {
        "dict_name": dict_name,
        "key": text(item.get("key")),
        "label": label,
        "label_en": label_en,
        "country": text(item.get("country")),
        "py_country_name": text(item.get("py_country_name")),
        "fetched_at": fetched_at,
    }


def main():
    fetched_at = datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")
    data = fetch_setting()
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=COLUMNS)
        writer.writeheader()
        for dict_name, items in data.items():
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, dict):
                    continue
                writer.writerow(to_row(dict_name, item, fetched_at))
                count += 1
    print(f"已写入 {OUTPUT_PATH}，{count} 条，抓取时间 {fetched_at}", flush=True)


if __name__ == "__main__":
    main()
