# Chapter 2 职位清洗

把 `Chapter1/data/jobs.xlsx` 清洗成 QuickBI 工作簿。原始 Excel 只读。每次运行都会重写输出，核数和待确认清单以当次代码结果为准。

## 运行

在项目根目录执行：

```bash
python Chapter2/clean_jobs.py --input Chapter1/data/jobs.xlsx --output-dir Chapter2
```

可选参数：

| 参数 | 默认 | 作用 |
| --- | --- | --- |
| `--category-catalog` | `Chapter1/data/job_categories.csv` | 确定分析大类和细类 |
| `--category-map` | 无 | 人工覆盖类别。CSV 需含 `category_id`、`analysis_category_l1`、`analysis_category_l2`，类别 ID 唯一 |
| `--city-map` | 无 | 城市标准名。CSV 需含 `city_raw`、`city_standard`、`province`、`country_or_region`、`geo_code` |
| `--early-before` | `2020-01-01` | 早于该日期（北京时间）的发布日期标为早期记录待复核 |

## 规则

- 职位类别有上级时，大类用上级、细类用自身。没有上级时视为顶级类，大类和细类都用自身。
- 发布时间用 `date-dimensions` 把毫秒时间戳换成北京时间，并展开年、季、周、旬。
- 地址用 `address-dimensions` 解析。原文为 unmatched 时，先试 `city_name + address`；短名不以「市」结尾且仍 unmatched 时，再试 `city_name + 市 + address`。
- 城市按顿号、中英文逗号和分号拆分。未传 `--city-map` 时，标准城市等于原名，状态为待映射。

## 输出

- `Chapter2/jobs_quickbi_cleaned.xlsx`：岗位主表、岗位城市、分类映射、核数表、异常明细、待确认清单、字段字典
- `Chapter2/jobs_cleaning_report.md`：输入范围、规则、全量核数和完整待确认清单

只有一个抓取时间时不输出岗位快照。完整核数见清洗说明，不在这里另抄一份。
