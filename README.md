# biaoge — Excel/CSV 命令行瑞士军刀

> 20 个命令覆盖表格日常：SQL 直查、合并拆分、透视表、VLOOKUP 平替、数据体检、文件监控、终端出图……
> 不会 pandas 也能 3 分钟上手，纯本地离线，数据不出电脑。

[![PyPI](https://img.shields.io/pypi/v/biaoge)](https://pypi.org/project/biaoge/)
[![Python](https://img.shields.io/pypi/pyversions/biaoge)](https://pypi.org/project/biaoge/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
[![Tests](https://img.shields.io/badge/tests-143%20passed-brightgreen)](./tests/)

---

## 30 秒安装上手

```bash
pip install biaoge
```

三条最常用命令：

```bash
# 把一整个文件夹的表格合成一张
biaoge merge ./月度报表/ -o 全年汇总.xlsx --source-col

# 用 SQL 查 Excel（零依赖，sqlite3 驱动）
biaoge sql "SELECT 部门, SUM(金额) FROM data GROUP BY 部门" -i 数据.xlsx -o 结果.xlsx

# 数据体检：缺失/重复/离群值/隐私列一键诊断
biaoge report 客户表.xlsx -o 体检报告.md
```

## 20 个命令一览

| # | 命令 | 一句话用途 | 典型场景 |
|---|------|-----------|---------|
| 1 | `merge` | 多表合一 / 两表关联合并 | 月度报表合并；员工表+工资表 VLOOKUP |
| 2 | `split` | 按列值拆分成多文件 | 全员名单按部门拆 |
| 3 | `dedup` | 去重 | 客户名单按手机号去重 |
| 4 | `convert` | xlsx ↔ csv 批量互转 | 整目录格式统一 |
| 5 | `summary` | 快速摸底（行列数/表头/空值率） | 合并前先看看 |
| 6 | `join` | 两表 VLOOKUP 式合并 | 订单表匹配客户信息 |
| 7 | `groupby` | 分组汇总（一维透视） | 按产品求和+数单量 |
| 8 | `pivot` | 二维交叉透视表 | 行=地区，列=月份，值=金额 |
| 9 | `unpivot` | 宽表转长表（逆透视） | 1月/2月/3月列拆成数据库格式 |
| 10 | `filter` | 按条件筛选行 | 金额>1000 且 地区~华东 |
| 11 | `clean` | 数据清洗 | 去空格/统一日期/填缺失 |
| 12 | `report` | 数据体检报告（Markdown） | 拿到表先体检 |
| 13 | `watch` | 监控文件夹自动汇总 | 收集箱场景，日报/周报自动汇入 |
| 14 | `diff` | 两表对比（新增/删除/修改） | 版本间差异、数据核对 |
| 15 | `sql` | SQL 查 Excel/CSV（sqlite3 零依赖） | 复杂查询、聚合、GROUP BY |
| 16 | `chart` | 从数据出图（bar/line/pie/scatter） | 终端出图，支持 SQL 查询出图 |
| 17 | `template` | 数据驱动批量文档生成 | 邮件合并平替，批量通知书 |
| 18 | `fill` | 空值向下填充 | 合并单元格拆分+前向填充 |
| 19 | `sort` | 智能排序 | 数值感知+自然排序+空值沉底+Top-N |
| 20 | `transpose` | 转置表格（行列互换） | 矩阵转置，支持无表头/index-col |

完整参数：`biaoge <命令> --help`

## 为什么用 biaoge？

| 你的情况 | biaoge 怎么帮 |
|---------|--------------|
| 不会 pandas / openpyxl | 中文命令行，看得懂就会用 |
| 不想开 Excel 等半天 | 一条命令秒级完成，可脚本化 |
| 数据敏感，不能上传 | 纯本地离线，数据不出电脑 |
| 表格多、重复劳动多 | 文件夹批量、监控自动汇总 |
| 想写 SQL 不想写公式 | `sql` 命令直查 Excel，sqlite3 零依赖 |

## 使用案例

### 案例 1：发票台账汇总

```bash
# 1. 先摸底看看每个文件长什么样
biaoge summary ./发票/

# 2. 合并所有发票，加上「来源文件」列
biaoge merge ./发票/ -o 全年发票.xlsx --source-col

# 3. 按供应商分组汇总金额
biaoge groupby 全年发票.xlsx --by 供应商 --agg 金额:sum,发票号:count --sort 金额 -o 供应商汇总.xlsx

# 4. 体检一下有没有缺失或异常
biaoge report 全年发票.xlsx -o 体检报告.md
```

### 案例 2：HR 简历初筛

```bash
# 1. 合并所有简历表
biaoge merge ./简历/ -o 全部简历.xlsx --source-col

# 2. 筛选：工作年限>3 且 期望城市~上海
biaoge filter 全部简历.xlsx --where "工作年限>3" --where "期望城市~上海" -o 初筛.xlsx

# 3. 按工作年限降序取前 20
biaoge sort 初筛.xlsx --by 工作年限:desc --limit 20 -o top20.xlsx

# 4. 去重（按手机号）
biaoge dedup top20.xlsx --columns 手机号 -o 去重后.xlsx
```

### 案例 3：销售报表透视

```bash
# 1. SQL 查出每个地区的销售总额
biaoge sql "SELECT 地区, SUM(金额) AS 总额 FROM data GROUP BY 地区" -i 销售表.xlsx -o 地区汇总.xlsx

# 2. 二维透视表：行=地区，列=月份，值=金额求和，加总计
biaoge pivot 销售表.xlsx --rows 地区 --cols 月份 --values 金额:sum --margins -o 透视表.xlsx

# 3. 宽表转长表，方便导入数据库
biaoge unpivot 透视表.xlsx --id 地区 --name 月份 --value 金额 -o 长表.xlsx

# 4. 出个柱状图
biaoge chart -i 地区汇总.xlsx --type bar --x 地区 --y 总额 --title "各地区销售额" --out 图表.png
```

## 路线图

- [x] v0.1-v0.16：20 个命令全部上线（merge/split/dedup/convert/summary/join/groupby/pivot/unpivot/filter/clean/report/watch/diff/sql/chart/template/fill/sort/transpose）
- [ ] v0.17：打包为独立 exe（免装 Python）
- [ ] v0.18：插件系统（自定义命令）
- [ ] v0.19：交互式 TUI 模式

功能建议 & Bug 反馈 → [Issues](https://github.com/mourinho2026/biaoge-cli/issues)。

## 贡献指南

欢迎 PR！代码风格保持简单，新增命令请附测试和文档。

```bash
git clone https://github.com/mourinho2026/biaoge-cli.git
cd biaoge-cli
python -m venv .venv && source .venv/bin/activate
pip install -e ".[chart,template]" pytest
pytest tests/ -v
```

发布前跑一遍：

```bash
bash release-check.sh
```

## 赞助

觉得好用？到[爱发电](https://afdian.com/@<待定>)赞助一杯奶茶，支持持续更新。

| 档位 | 金额 | 权益 |
|------|------|------|
| 请杯咖啡 | 5 元 | 感谢 + 名字进 README |
| 月赞助 | 30 元/月 | 新功能投票权 |
| 年赞助 | 100 元/年 | 专属工具定制一次 |

## License

[MIT](./LICENSE) — Copyright 2026 biaoge contributors
