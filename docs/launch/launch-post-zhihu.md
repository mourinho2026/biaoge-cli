看到这个问题，必须来安利一下我做的开源工具：**biaoge**（表格）。

## 相见恨晚的理由

作为一个经常要处理各种 Excel/CSV 的人，我以前也是：

- 不会 pandas，只能手动复制粘贴合并报表
- VLOOKUP 写到头晕，还经常出错
- 想做个透视表，要在 Excel 里拖半天
- 数据敏感，不敢上传到在线工具

直到我做了 biaoge，一个命令行工具，20 个命令覆盖 80% 的表格日常。

## 几个让我自己都觉得"早该这样做"的功能

### 1. SQL 直查 Excel

不会写公式？直接写 SQL：



零依赖，sqlite3 驱动，中文列名随便用。

### 2. 一键合并整个文件夹

月底了，要把 30 个分公司的报表合成一张？



自动加"来源文件"列，方便追溯。

### 3. 数据体检

拿到一张表，先体检再看：



缺失率、重复值、离群点、隐私列，一键诊断，出 Markdown 报告。

### 4. 文件监控自动汇总

收集箱场景：大家把日报丢进一个文件夹，自动汇入总表：



配合 cron，定时任务搞定。

## 20 个命令一览

merge/split/dedup/convert/summary/join/groupby/pivot/unpivot/filter/clean/report/watch/diff/sql/chart/template/fill/sort/transpose

覆盖：合并拆分、VLOOKUP、透视表、数据清洗、体检报告、文件监控、SQL 查询、批量生成文档、智能排序、转置……

## 为什么不用 pandas？

- 学习成本高（DataFrame、merge、groupby 语法）
- 要装环境（Python、Jupyter、各种库）
- 脚本不好分享给同事

biaoge 就是为了解决这些问题：

- 中文命令行，看得懂就会用
- pip install 一行搞定
- 纯本地离线，数据不出电脑

## 安装

```
pip install biaoge
```

GitHub: https://github.com/mourinho2026/biaoge-cli

欢迎试用，欢迎提需求。觉得好用可以点个 Star 或到爱发电赞助一杯奶茶。

以上。
