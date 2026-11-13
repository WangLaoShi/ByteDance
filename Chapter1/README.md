# Chapter 1 职位抓取

社招和校招是同一套搜索接口，写在同一个脚本里。

## 请求差异

两端都是 `POST https://jobs.bytedance.com/api/v1/search/job/posts`。分页只改 `offset`，`limit` 最大可用 100。浏览器里的 `_signature` 和 Cookie 不是必需的。

| | 社招 | 校招 |
| --- | --- | --- |
| `portal_type` | 2 | 3 |
| `Portal-Channel` | office | campus |
| `website-path` | society | campus |
| Referer | `/experienced/position` | `/campus/position` |

## 返回结构

两边的 `job_post_list` 字段相同。每条职位额外写入：

- `portal_type`：2 社招，3 校招
- `portal_name`：`社招` 或 `校招`
- `fetched_at`：本次抓取时间（东八区）

`recruit_type.parent` 里也有「社招 / 校招」，额外字段用来直接区分渠道。

## 为什么社招要按类别抓

未筛选列表的 `count` 到 10000 后不再返回数据。社招按类别树逐个翻页，再用职位 `id` 去重。一级类别的结果已经包含其子类别。校招总数低于 10000，直接翻页。

单个类别如果也顶到 10000，`fetch_meta.json` 的 `capped_categories` 会记下来。

## 运行

在项目根目录执行：

```bash
python Chapter1/fetch_jobs.py
python Chapter1/export_jobs.py
python Chapter1/fetch_settings.py
python Chapter1/fetch_filters.py
```

结果：

- `Chapter1/data/jobs.jsonl`：当前抓到的全部职位，按 `publish_time` 从新到旧
- `Chapter1/data/fetch_meta.json`：最后抓取时间、分渠道条数、社招各类别条数
- `Chapter1/data/jobs.csv`：同一份职位的 CSV
- `Chapter1/data/jobs.xlsx`：同一份职位的 Excel
- `Chapter1/data/settings.csv`：简历和表单共用枚举。列是 `dict_name`、`key`、`label`，手机区号另有国家名和拼音，币种另有英文代码。不并入职位表
- `Chapter1/data/job_categories.csv`：社招、校招各自的职位类别。`filter_count` 只是筛选项上的计数，不是抓到的职位条数
- `Chapter1/data/cities.csv`：两个渠道各自可选的城市
- `Chapter1/data/job_subjects.csv`：招聘项目，目前只有校招有

`publish_time` 是毫秒时间戳。`job_categories.csv` 会在 [Chapter2](../Chapter2/README.md) 里用来确定职位大类和细类。

## 本次结果

抓取时间 `2026-11-12T16:12:57+08:00`。

| 渠道 | 条数 | 方式 |
| --- | --- | --- |
| 社招 | 10414 | 按类别抓取。未筛选接口只给出 10000 |
| 校招 | 6675 | 直接翻页 |
| 合计 | 17089 | 职位 `id` 不重复 |
