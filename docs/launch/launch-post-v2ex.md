# [分享创造] biaoge - 一个让表格处理变简单的命令行工具

大家好，分享一个我做的开源命令行工具：**biaoge**（表格）。

## 痛点

办公室里经常要处理各种 Excel/CSV：合并报表、去重、VLOOKUP、透视表……不会 pandas 的人要么手动搞半天，要么求会 Python 的同事帮忙。

## 解决方案

biaoge 把 20 个常用表格操作做成了命令行，中文报错、中文文档、支持中文路径：

```bash
# 合并一整个文件夹的表格
biaoge merge ./月度报表/ -o 全年汇总.xlsx --source-col

# 用 SQL 查 Excel（零依赖）
biaoge sql "SELECT 部门, SUM(金额) FROM data GROUP BY 部门" -i 数据.xlsx

# 数据体检：缺失/重复/离群值一键诊断
biaoge report 客户表.xlsx -o 体检报告.md
```

## 20 个命令

merge/split/dedup/convert/summary/join/groupby/pivot/unpivot/filter/clean/report/watch/diff/sql/chart/template/fill/sort/transpose

覆盖：合并拆分、VLOOKUP、透视表、数据清洗、体检报告、文件监控、SQL 查询、批量生成文档、智能排序、转置……

## 特点

- **纯本地离线**：数据不出电脑
- **零学习成本**：看得懂中文就会用
- **安全**：只读源文件，产出写到新文件

## 安装

```bash
pip install biaoge
```

GitHub: https://github.com/mourinho2026/biaoge-cli

欢迎试用、提建议、报 Bug。觉得好用可以点个 ⭐ 或到爱发电赞助一杯奶茶 🙏
