"""抓取字节跳动社招、校招当前在招职位。

社招和校招共用 POST /api/v1/search/job/posts，响应字段一致。
两者只靠 portal_type 和渠道请求头区分：社招 portal_type=2，校招 portal_type=3。
未筛选列表最多返回 10000 条。社招分类合计超过这个上限，因此按职位类别逐个抓取。
"""

import json
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

SEARCH_URL = "https://jobs.bytedance.com/api/v1/search/job/posts"
FILTERS_URL = "https://jobs.bytedance.com/api/v1/config/job/filters/{portal_type}"
PAGE_SIZE = 100
LIST_CAP = 10000
REQUEST_INTERVAL = 0.3
OUTPUT_DIR = Path(__file__).resolve().parent / "data"

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


def portal_headers(portal):
    return {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://jobs.bytedance.com",
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
    }


def search_payload(portal, offset, category_id):
    return {
        "keyword": "",
        "limit": PAGE_SIZE,
        "offset": offset,
        "job_category_id_list": [category_id] if category_id else [],
        "tag_id_list": [],
        "location_code_list": [],
        "subject_id_list": [],
        "recruitment_id_list": [],
        "portal_type": portal["portal_type"],
        "job_function_id_list": [],
        "storefront_id_list": [],
        "portal_entrance": 1,
    }


def query_params(payload):
    params = {}
    for key, value in payload.items():
        if isinstance(value, list):
            params[key] = ",".join(value)
        else:
            params[key] = value
    return params


def request_json(session, method, url, portal, **kwargs):
    last_error = None
    for attempt in range(4):
        try:
            response = session.request(
                method,
                url,
                headers=portal_headers(portal),
                timeout=30,
                **kwargs,
            )
            response.raise_for_status()
            body = response.json()
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
            continue
        if body.get("code") != 0:
            last_error = RuntimeError(body.get("message") or body)
            time.sleep(1.5 * (attempt + 1))
            continue
        return body.get("data") or {}
    raise RuntimeError(f"请求失败: {url}") from last_error


def flatten_categories(nodes):
    categories = []
    for node in nodes or []:
        categories.append({"id": node["id"], "name": node.get("name") or ""})
        categories.extend(flatten_categories(node.get("children")))
    return categories


def load_categories(session, portal):
    data = request_json(
        session,
        "GET",
        FILTERS_URL.format(portal_type=portal["portal_type"]),
        portal,
    )
    return flatten_categories(data.get("job_type_list"))


def fetch_pages(session, portal, category_id, fetched_at):
    """按 offset 翻页，直到本查询没有更多职位。"""
    offset = 0
    reported = None
    jobs = []
    seen = set()
    while offset < LIST_CAP:
        payload = search_payload(portal, offset, category_id)
        data = request_json(
            session,
            "POST",
            SEARCH_URL,
            portal,
            params=query_params(payload),
            json=payload,
        )
        reported = data.get("count") or 0
        batch = data.get("job_post_list") or []
        if not batch:
            break
        for job in batch:
            job_id = job.get("id")
            if job_id in seen:
                continue
            seen.add(job_id)
            job["portal_type"] = portal["portal_type"]
            job["portal_name"] = portal["portal_name"]
            job["fetched_at"] = fetched_at
            jobs.append(job)
        offset += len(batch)
        if offset >= reported:
            break
        time.sleep(REQUEST_INTERVAL)
    return jobs, reported


def fetch_portal(session, portal, fetched_at):
    payload = search_payload(portal, 0, None)
    probe_data = request_json(
        session,
        "POST",
        SEARCH_URL,
        portal,
        params=query_params(payload),
        json=payload,
    )
    unfiltered_count = probe_data.get("count") or 0
    time.sleep(REQUEST_INTERVAL)
    if unfiltered_count < LIST_CAP:
        jobs, reported = fetch_pages(session, portal, None, fetched_at)
        return jobs, {
            "portal_type": portal["portal_type"],
            "portal_name": portal["portal_name"],
            "mode": "direct",
            "api_count": reported,
            "fetched_count": len(jobs),
            "capped": False,
        }

    categories = load_categories(session, portal)
    jobs_by_id = {}
    category_stats = []
    capped_categories = []
    for category in categories:
        batch, api_count = fetch_pages(session, portal, category["id"], fetched_at)
        category_stats.append(
            {
                "id": category["id"],
                "name": category["name"],
                "api_count": api_count,
                "fetched_count": len(batch),
            }
        )
        if api_count >= LIST_CAP:
            capped_categories.append(category["name"])
        for job in batch:
            jobs_by_id.setdefault(job["id"], job)
        print(
            f"  {portal['portal_name']} {category['name']}: "
            f"{len(batch)}/{api_count}，累计 {len(jobs_by_id)}",
            flush=True,
        )
        time.sleep(REQUEST_INTERVAL)

    jobs = list(jobs_by_id.values())
    return jobs, {
        "portal_type": portal["portal_type"],
        "portal_name": portal["portal_name"],
        "mode": "by_category",
        "api_count": unfiltered_count,
        "fetched_count": len(jobs),
        "capped": True,
        "capped_categories": capped_categories,
        "categories": category_stats,
    }


def sort_jobs(jobs):
    def sort_key(job):
        publish_time = job.get("publish_time")
        if not isinstance(publish_time, (int, float)):
            publish_time = -1
        return publish_time, job.get("id") or ""

    return sorted(jobs, key=sort_key, reverse=True)


def main():
    fetched_at = datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")
    session = requests.Session()
    all_jobs = []
    portal_meta = []
    for portal in PORTALS:
        print(f"开始抓取{portal['portal_name']}", flush=True)
        jobs, meta = fetch_portal(session, portal, fetched_at)
        all_jobs.extend(jobs)
        portal_meta.append(meta)
        print(
            f"{portal['portal_name']}完成: {meta['fetched_count']} 条"
            f"（接口未筛选 count={meta['api_count']}）",
            flush=True,
        )

    all_jobs = sort_jobs(all_jobs)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    jobs_path = OUTPUT_DIR / "jobs.jsonl"
    meta_path = OUTPUT_DIR / "fetch_meta.json"
    with jobs_path.open("w", encoding="utf-8") as output:
        for job in all_jobs:
            output.write(json.dumps(job, ensure_ascii=False) + "\n")
    meta = {
        "fetched_at": fetched_at,
        "total": len(all_jobs),
        "page_size": PAGE_SIZE,
        "list_cap": LIST_CAP,
        "output": str(jobs_path),
        "portals": portal_meta,
    }
    meta_path.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"已写入 {jobs_path}，共 {len(all_jobs)} 条，抓取时间 {fetched_at}", flush=True)


if __name__ == "__main__":
    main()
