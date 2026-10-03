# 用 stdlib sqlite3 实现 SQL 直查 Excel：biaoge sql 命令的设计思路

## 背景

[biaoge](https://github.com/mourinho2026/biaoge-cli) 是一个 Excel/CSV 命令行工具，其中 sql 命令可以用 SQL 查询 Excel 文件，零外部依赖（只用 Python 标准库 sqlite3）。

本文分享实现思路。

## 核心思路

1. 读取 Excel 到内存数据：用 openpyxl 读取 .xlsx，得到 headers 和 rows
2. 创建临时 SQLite 数据库：:memory: 模式，不写磁盘
3. 动态建表：根据 headers 推断列类型（数值/文本），生成 CREATE TABLE
4. 批量插入：executemany 一次性灌入数据
5. 执行用户 SQL：表名固定为 data
6. 输出结果：写回 Excel/CSV 或终端打印

## 关键代码

import sqlite3
from openpyxl import load_workbook

def cmd_sql(args):
    # 1. 读 Excel
    wb = load_workbook(args.input, read_only=True)
    ws = wb.active
    rows = list(ws.values)
    headers = [str(h) for h in rows[0]]
    data_rows = rows[1:]
    
    # 2. 推断列类型
    types = []
    for i in range(len(headers)):
        values = [r[i] for r in data_rows if r[i] is not None]
        if all(isinstance(v, (int, float)) for v in values[:10]):
            types.append('REAL')
        else:
            types.append('TEXT')
    
    # 3. 建表
    conn = sqlite3.connect(':memory:')
    cols = ', '.join(f'"{h}" {t}' for h, t in zip(headers, types))
    conn.execute(f'CREATE TABLE data ({cols})')
    
    # 4. 插入数据
    placeholders = ', '.join(['?'] * len(headers))
    conn.executemany(f'INSERT INTO data VALUES ({placeholders})', data_rows)
    
    # 5. 执行查询
    cursor = conn.execute(args.sql)
    result_headers = [d[0] for d in cursor.description]
    result_rows = cursor.fetchall()

## 设计取舍

- 表名固定 data：简化用户心智，不用记表名
- 中文列名：SQLite 支持，但要用双引号包裹
- 类型推断只看前 10 行：性能考虑，大文件不遍历全部
- 内存数据库：不写临时文件，安全快速

## 使用示例

biaoge sql "SELECT 部门, SUM(金额) FROM data GROUP BY 部门" -i 数据.xlsx -o 结果.xlsx

## 局限

- 不支持 Excel 公式、合并单元格、多 sheet 关联
- 大文件（>100MB）会占较多内存
- 类型推断可能不准（只看前 10 行）

但对于日常 80% 的查询场景，够用了。

欢迎 Star、PR、提需求。
