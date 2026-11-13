# ByteDance

字节职位分析。先抓取招聘职位和筛选项，再清洗成可供 QuickBI 使用的工作簿。

## 环境

Python 3.10 及以上。

## 安装

```bash
pip install -r requirements.txt
```

`address-dimensions` 从 Git 安装。若 HTTPS 克隆失败，改用 SSH：

```bash
pip install "address-dimensions @ git+ssh://git@github.com/WangLaoShi/address-dimensions.git"
```

## 章节

- [Chapter1](Chapter1/README.md)：抓取社招、校招职位，以及设置枚举和筛选项。
- [Chapter2](Chapter2/README.md)：清洗职位表，输出 QuickBI 工作簿和清洗说明。
- Chapter3、Chapter4：尚未开始。
