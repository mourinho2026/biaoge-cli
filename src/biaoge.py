#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/CSV 批量处理工具箱
功能：合并 / 拆分 / 去重 / 格式转换 / 汇总统计
作者：夜间工作室 | 版本 1.0
"""
import argparse
import csv
import datetime
import hashlib
import math
import os
import re
import sqlite3
import sys
from pathlib import Path

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少依赖：请先运行  pip install openpyxl")
    sys.exit(1)

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None


# ---------- 通用读取 ----------
def read_table(path):
    """读取单个 xlsx/csv，返回 (headers, rows[list[list]])"""
    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        data = [[("" if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    elif ext == ".csv":
        with open(p, newline="", encoding="utf-8-sig") as f:
            data = [row for row in csv.reader(f)]
    elif ext == ".txt":
        # 尝试常见分隔符
        with open(p, encoding="utf-8-sig") as f:
            sample = f.read(4096)
        delim = "\t" if "\t" in sample else ("," if "," in sample else None)
        with open(p, newline="", encoding="utf-8-sig") as f:
            data = [row for row in csv.reader(f, delimiter=delim)] if delim else [
                [line.rstrip("\n")] for line in f]
    else:
        raise ValueError(f"不支持的文件类型: {p.name}（支持 .xlsx/.csv/.txt）")
    if not data:
        return [], []
    return data[0], data[1:]


def write_table(path, headers, rows):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix.lower() == ".csv":
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for r in rows:
            ws.append(r)
        wb.save(p)


def to_number(v):
    """尽量把单元格值转成数字（去千分位逗号/￥/$/空格），失败返回 None"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "").replace("￥", "").replace("¥", "").replace("$", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def collect_files(paths, recursive=False):
    """展开目录/文件列表为表格文件列表"""
    exts = {".xlsx", ".xlsm", ".csv", ".txt"}
    out = []
    for item in paths:
        p = Path(item)
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            out.extend(sorted(f for f in it if f.suffix.lower() in exts and not f.name.startswith("~$")))
        elif p.is_file():
            out.append(p)
        else:
            print(f"[警告] 路径不存在，跳过: {item}")
    return out


def _read_sheet(path, sheet_name):
    """读取 xlsx 指定 sheet，返回 (headers, rows)；csv 或无 sheet_name 时走 read_table"""
    p = Path(path)
    if not sheet_name or p.suffix.lower() not in ('.xlsx', '.xlsm'):
        return read_table(p)
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    data = [[('' if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    if not data:
        return [], []
    return data[0], data[1:]


# ---------- 子命令 ----------
def cmd_merge(args):
    """合并多个表格为一个文件，或两表按关键列关联（--on 模式，VLOOKUP 平替）"""
    if getattr(args, 'on', None):
        return _merge_join(args)
    files = collect_files(args.inputs, args.recursive)
    if not files:
        print("没有找到可合并的文件"); return 1
    headers = None
    all_rows = []
    for f in files:
        h, rows = read_table(f)
        if not h:
            continue
        if headers is None:
            headers = h
            if args.source_col:
                headers = h + ["来源文件"]
        base = f.name
        for r in rows:
            if args.source_col:
                all_rows.append(list(r) + [base])
            else:
                all_rows.append(r)
        print(f"  读取 {f.name}: {len(rows)} 行")
    write_table(args.output, headers, all_rows)
    print(f"✅ 合并完成：{len(files)} 个文件 → {args.output}（共 {len(all_rows)} 行）")
    return 0


def _merge_join(args):
    """两表按 key 关联（VLOOKUP 平替）"""
    on_cols = [c.strip() for c in args.on.split(",")]
    how = args.how
    suffix = args.suffix
    dup = args.dup

    left_h, left_rows = _read_sheet(args.inputs[0], getattr(args, 'left_sheet', None))
    right_h, right_rows = _read_sheet(args.inputs[1], getattr(args, 'right_sheet', None))

    if not left_h or not right_h:
        print("错误：读不到表头（空文件？）"); return 1

    for c in on_cols:
        if c not in left_h:
            print(f"错误：左表里没有关键列「{c}」，现有列：{left_h}"); return 1
        if c not in right_h:
            print(f"错误：右表里没有关键列「{c}」，现有列：{right_h}"); return 1

    on_left = [left_h.index(c) for c in on_cols]
    on_right = [right_h.index(c) for c in on_cols]

    right_non_key = [c for c in right_h if c not in on_cols]
    right_non_key_idx = [right_h.index(c) for c in right_non_key]
    right_out_names = [c + suffix if c in left_h and c not in on_cols else c for c in right_non_key]

    def _norm_key(v):
        if v is None:
            return ""
        s = str(v).strip().lower()
        n = to_number(s)
        if n is not None and isinstance(n, float) and n == int(n):
            return str(int(n))
        return s if n is None else str(n)

    def _make_key(row, idxs):
        return tuple(_norm_key(row[i] if i < len(row) else "") for i in idxs)

    right_index = {}
    right_dup_count = 0
    for r in right_rows:
        k = _make_key(r, on_right)
        if k in right_index:
            if dup == "error":
                print(f"错误：右表关键列重复：{k}（可用 --dup first/concat 处理）"); return 1
            elif dup == "concat":
                entry = right_index[k]
                for ci, col_name in zip(right_non_key_idx, right_non_key):
                    v = r[ci] if ci < len(r) else ""
                    existing = entry[col_name]
                    try:
                        float(str(existing).strip())
                    except (TypeError, ValueError):
                        entry[col_name] = str(existing) + "|" + str(v) if existing != "" else str(v)
                continue
            else:
                right_dup_count += 1
                continue
        vals = {}
        for ci, col_name in zip(right_non_key_idx, right_non_key):
            vals[col_name] = r[ci] if ci < len(r) else ""
        right_index[k] = vals

    if right_dup_count:
        print(f"⚠ 右表有 {right_dup_count} 个重复 key，已取第一条")

    left_index = {}
    for r in left_rows:
        k = _make_key(r, on_left)
        if k not in left_index:
            left_index[k] = r

    left_keys = set(left_index.keys())
    right_keys = set(right_index.keys())
    common_keys = left_keys & right_keys

    out_headers = list(left_h) + right_out_names
    out_rows = []
    matched = unmatched = 0

    if how == "left":
        for r in left_rows:
            k = _make_key(r, on_left)
            if k in right_index:
                matched += 1
                rv = right_index[k]
                out_rows.append(list(r) + [rv[c] for c in right_non_key])
            else:
                unmatched += 1
                out_rows.append(list(r) + [""] * len(right_non_key))
    elif how == "inner":
        for r in left_rows:
            k = _make_key(r, on_left)
            if k in right_index:
                matched += 1
                rv = right_index[k]
                out_rows.append(list(r) + [rv[c] for c in right_non_key])
            else:
                unmatched += 1
    elif how == "right":
        for r in right_rows:
            k = _make_key(r, on_right)
            if k in left_index:
                matched += 1
                lr = left_index[k]
                out_rows.append(list(lr) + [r[ci] if ci < len(r) else "" for ci in right_non_key_idx])
            else:
                unmatched += 1
                left_empty = [""] * len(left_h)
                for ki, col in zip(on_left, on_cols):
                    ri = right_h.index(col)
                    left_empty[ki] = r[ri] if ri < len(r) else ""
                out_rows.append(left_empty + [r[ci] if ci < len(r) else "" for ci in right_non_key_idx])

    write_table(Path(args.output), out_headers, out_rows)

    total = len(left_rows) if how in ("left", "inner") else len(right_rows)
    rate = matched / total * 100 if total else 0
    print(f"📊 关联统计：")
    print(f"   左表：{len(left_rows)} 行")
    print(f"   右表：{len(right_rows)} 行")
    print(f"   匹配：{matched} 行")
    print(f"   未匹配：{unmatched} 行")
    print(f"   匹配率：{rate:.1f}%")

    if how == "left":
        unmatched_right = right_keys - left_keys
        if unmatched_right:
            print(f"   ⚠ 右表有 {len(unmatched_right)} 个 key 在左表未匹配")

    return 0


def cmd_split(args):
    """按列值把一个大表拆成多个文件"""
    headers, rows = read_table(args.input)
    if args.column not in headers:
        print(f"错误：找不到列「{args.column}」，可用列：{headers}"); return 1
    idx = headers.index(args.column)
    groups = {}
    for r in rows:
        key = str(r[idx]) if idx < len(r) else "空"
        groups.setdefault(key, []).append(r)
    stem = Path(args.input).stem
    out_dir = Path(args.output_dir)
    for key, grp in sorted(groups.items()):
        safe = "".join(c if c.isalnum() or c in "-_（）() " else "_" for c in key)[:40] or "空"
        out = out_dir / f"{stem}_{safe}{args.fmt}"
        write_table(out, headers, grp)
        print(f"  {out.name}: {len(grp)} 行")
    print(f"✅ 拆分完成：{len(groups)} 个文件 → {out_dir}/")
    return 0


def cmd_dedup(args):
    """去重：按指定列或整行去重"""
    headers, rows = read_table(args.input)
    if args.columns:
        idxs = []
        for c in args.columns.split(","):
            c = c.strip()
            if c in headers:
                idxs.append(headers.index(c))
            else:
                print(f"错误：找不到列「{c}」，可用列：{headers}"); return 1
    else:
        idxs = list(range(len(headers)))
    seen, kept = set(), []
    for r in rows:
        key = tuple(str(r[i]) if i < len(r) else "" for i in idxs)
        if key in seen:
            continue
        seen.add(key)
        kept.append(r)
    write_table(args.output, headers, kept)
    print(f"✅ 去重完成：{len(rows)} → {len(kept)} 行（去掉 {len(rows)-len(kept)} 条重复）→ {args.output}")
    return 0


def cmd_convert(args):
    """格式互转：xlsx↔csv（可批量、可递归目录）"""
    files = collect_files(args.inputs, args.recursive)
    if not files:
        print("没有找到可转换的文件"); return 1
    out_dir = Path(args.output_dir); out_dir.mkdir(parents=True, exist_ok=True)
    n = 0
    for f in files:
        h, rows = read_table(f)
        out = out_dir / (f.stem + args.fmt)
        write_table(out, h, rows)
        n += 1
        print(f"  {f.name} → {out.name}")
    print(f"✅ 转换完成：{n} 个文件 → {out_dir}/")
    return 0


def cmd_summary(args):
    """快速摸底：每个文件的行列数、表头、数值列统计"""
    files = collect_files(args.inputs, args.recursive)
    if not files:
        print("没有找到文件"); return 1
    for f in files:
        try:
            h, rows = read_table(f)
        except Exception as e:
            print(f"❌ {f.name}: 读取失败 {e}"); continue
        print(f"\n📄 {f.name}  ({len(rows)} 行 × {len(h)} 列)")
        print("   表头: " + " | ".join(str(x) for x in h))
        for i, name in enumerate(h):
            nums = []
            for r in rows[:5000]:
                if i < len(r):
                    v = r[i]
                    try:
                        nums.append(float(v))
                    except (TypeError, ValueError):
                        pass
            if nums and len(nums) >= max(1, len(rows[:5000])) * 0.7:
                print(f"   [{name}] 数值列: 合计 {sum(nums):.2f} 平均 {sum(nums)/len(nums):.2f} "
                      f"最小 {min(nums):.2f} 最大 {max(nums):.2f}")
    return 0


def cmd_join(args):
    """两表按关键列 VLOOKUP 式合并"""
    left_headers, left_rows = read_table(args.left)
    right_headers, right_rows = read_table(args.right)

    if args.on not in left_headers:
        print(f"错误：左表里没有列「{args.on}」，现有列：{left_headers}")
        return 2
    if args.on not in right_headers:
        print(f"错误：右表里没有列「{args.on}」，现有列：{right_headers}")
        return 2

    want_cols = [c.strip() for c in args.columns.split(",")] if args.columns else \
        [c for c in right_headers if c != args.on]
    missing = [c for c in want_cols if c not in right_headers]
    if missing:
        print(f"错误：右表里没有列 {missing}，现有列：{right_headers}")
        return 2

    # 建右表索引（重复键取第一条，除非 --multi）
    on_idx = right_headers.index(args.on)
    col_idx = [right_headers.index(c) for c in want_cols]
    index = {}
    for r in right_rows:
        key = str(r[on_idx]) if on_idx < len(r) else ""
        if key in index and not args.multi:
            continue
        vals = [r[i] if i < len(r) else "" for i in col_idx]
        if args.multi:
            index.setdefault(key, []).append(vals)
        else:
            index[key] = vals

    left_on = left_headers.index(args.on)
    prefix = args.prefix or ""
    out_headers = list(left_headers) + [prefix + c for c in want_cols]
    out_rows = []
    matched = unmatched = 0
    for r in left_rows:
        key = str(r[left_on]) if left_on < len(r) else ""
        hit = index.get(key)
        if hit:
            matched += 1
            if args.multi:
                for vals in hit:
                    out_rows.append(list(r) + list(vals))
            else:
                out_rows.append(list(r) + list(hit))
        else:
            unmatched += 1
            if not args.only_matched:
                out_rows.append(list(r) + [""] * len(want_cols))

    write_table(Path(args.output), out_headers, out_rows)
    print(f"完成：左表 {len(left_rows)} 行，匹配 {matched} 行，未匹配 {unmatched} 行"
          f"{'（未匹配行已跳过）' if args.only_matched else ''}")
    print(f"输出：{args.output}（{len(out_rows)} 行 × {len(out_headers)} 列）")
    return 0


AGG_FUNCS = {
    "sum": lambda xs: sum(xs),
    "count": lambda xs: len(xs),
    "avg": lambda xs: sum(xs) / len(xs) if xs else 0,
    "max": lambda xs: max(xs) if xs else "",
    "min": lambda xs: min(xs) if xs else "",
}


def cmd_groupby(args):
    """分组汇总（透视表平替）：按一列分组，对数值列求和/计数/平均/最大最小"""
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    if args.by not in headers:
        print(f"错误：表里没有列「{args.by}」，现有列：{headers}")
        return 2

    by_idx = headers.index(args.by)
    # 解析聚合定义：列名:函数，可多个，逗号分隔；默认对所有数值列 sum
    if args.agg:
        specs = []
        for part in args.agg.split(","):
            part = part.strip()
            if ":" in part:
                col, fn = part.rsplit(":", 1)
                col, fn = col.strip(), fn.strip().lower()
            else:
                col, fn = part, "sum"
            if fn not in AGG_FUNCS:
                print(f"错误：不支持的聚合函数「{fn}」，可选：{'/'.join(AGG_FUNCS)}")
                return 2
            if col not in headers:
                print(f"错误：表里没有列「{col}」，现有列：{headers}")
                return 2
            specs.append((col, fn))
    else:
        # 默认：找出所有数值列，全部 sum
        specs = []
        for i, name in enumerate(headers):
            if i == by_idx:
                continue
            nums = 0
            for r in rows[:200]:
                if i < len(r):
                    try:
                        float(r[i]); nums += 1
                    except (TypeError, ValueError):
                        pass
            if nums >= max(1, min(len(rows), 200)) * 0.7:
                specs.append((name, "sum"))
        if not specs:
            specs = [(headers[min(1, len(headers)-1)], "count")]

    col_idx = {name: headers.index(name) for name, _ in specs}
    groups = {}
    order = []
    for r in rows:
        key = str(r[by_idx]) if by_idx < len(r) else ""
        if key not in groups:
            groups[key] = [[] for _ in specs]
            order.append(key)
        for si, (name, _) in enumerate(specs):
            v = r[col_idx[name]] if col_idx[name] < len(r) else ""
            try:
                groups[key][si].append(float(v))
            except (TypeError, ValueError):
                if v != "":
                    groups[key][si].append(v)  # 非数值保留给 count
    out_headers = [args.by] + [f"{fn}({name})" for name, fn in specs]
    out_rows = []
    for key in order:
        row = [key]
        for si, (name, fn) in enumerate(specs):
            vals = groups[key][si]
            if fn == "count":
                row.append(len(vals))
            else:
                nums = [v for v in vals if isinstance(v, float)]
                result = AGG_FUNCS[fn](nums)
                row.append(round(result, 4) if isinstance(result, float) else result)
        out_rows.append(row)
    if args.sort:
        names = [n for n, _ in specs]
        idx = -1
        if args.sort in names:
            idx = 1 + names.index(args.sort)
        else:
            # 也允许直接写输出列名，如 sum(金额)
            if args.sort in out_headers:
                idx = out_headers.index(args.sort)
        if idx > 0:
            def sort_key(r):
                v = r[idx]
                return v if isinstance(v, (int, float)) else float("-inf")
            out_rows.sort(key=sort_key, reverse=not args.asc)
    write_table(Path(args.output), out_headers, out_rows)
    print(f"完成：{len(rows)} 行 → {len(out_rows)} 组（按「{args.by}」）")
    print(f"聚合：" + "，".join(f"{fn}({name})" for name, fn in specs))
    print(f"输出：{args.output}")
    return 0


def cmd_filter(args):
    """按条件筛选行：列=值 / 列!=值 / 列>值 / 列>=值 / 列<值 / 列<=值 / 列~包含"""
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    headers = [str(h) for h in headers]

    conds = []
    for expr in args.where:
        m = re.match(r"^\s*([^!=<>~]+?)\s*(!=|>=|<=|=|>|<|~)\s*(.*)$", expr)
        if not m:
            print(f"错误：看不懂条件「{expr}」。写法：列名=值 / 列名>100 / 列名~包含字 / 列名!=值")
            return 2
        col, op, val = m.group(1).strip(), m.group(2), m.group(3).strip()
        if col not in headers:
            print(f"错误：表里没有列「{col}」，现有列：{headers}")
            return 2
        conds.append((headers.index(col), op, val))

    def row_ok(r):
        for idx, op, val in conds:
            cell = r[idx] if idx < len(r) else ""
            cs = "" if cell is None else str(cell).strip()
            if op == "=":
                if cs != val:
                    return False
            elif op == "!=":
                if cs == val:
                    return False
            elif op == "~":
                if val not in cs:
                    return False
            else:  # 数值比较
                a, b = to_number(cell), to_number(val)
                if a is None or b is None:
                    return False
                if op == ">" and not a > b: return False
                if op == ">=" and not a >= b: return False
                if op == "<" and not a < b: return False
                if op == "<=" and not a <= b: return False
        return True

    out_rows = [r for r in rows if row_ok(r)]
    if args.drop_empty:
        out_rows = [r for r in out_rows if any(str(c).strip() for c in r if c is not None)]
    write_table(Path(args.output), headers, out_rows)
    print(f"完成：{len(rows)} 行 → 筛出 {len(out_rows)} 行")
    print("条件：" + " 且 ".join(args.where))
    print(f"输出：{args.output}")
    return 0


DATE_PATTERNS = [
    (re.compile(r"^(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?$"), lambda m: f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"),
    (re.compile(r"^(\d{1,2})[-/月.](\d{1,2})[-/日.](\d{4})$"), None),  # 美式不猜，跳过
]


def cmd_clean(args):
    """清洗：去首尾空格/去空行/统一日期/填缺失值/去全角空格"""
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    stats = {"空行": 0, "去空格": 0, "日期统一": 0, "填缺失": 0}

    norm_headers = []
    for h in headers:
        nh = str(h).replace("\u3000", " ").strip()
        if nh != str(h):
            stats["去空格"] += 1
        norm_headers.append(nh)

    targets = None
    if args.columns:
        targets = set()
        for c in args.columns.split(","):
            c = c.strip()
            if c not in norm_headers:
                print(f"错误：表里没有列「{c}」，现有列：{norm_headers}")
                return 2
            targets.add(norm_headers.index(c))

    out_rows = []
    for r in rows:
        cells = list(r) + [""] * (len(norm_headers) - len(r))
        if not any(str(c).strip() for c in cells if c is not None):
            stats["空行"] += 1
            continue
        for i in range(min(len(cells), len(norm_headers))):
            if targets is not None and i not in targets:
                continue
            v = cells[i]
            if v is None:
                v = ""
            if isinstance(v, str):
                nv = v.replace("\u3000", " ").strip()
                if nv != v:
                    stats["去空格"] += 1
                v = nv
                if args.dates:
                    for pat, fmt in DATE_PATTERNS:
                        if fmt is None:
                            continue
                        m = pat.match(v)
                        if m:
                            nv = fmt(m)
                            if nv != v:
                                stats["日期统一"] += 1
                            v = nv
                            break
            if v == "" and args.fill:
                v = args.fill
                stats["填缺失"] += 1
            cells[i] = v
        out_rows.append(cells[:len(norm_headers)])

    write_table(Path(args.output), norm_headers, out_rows)
    print(f"完成：{len(rows)} 行 → {len(out_rows)} 行")
    print("清洗统计：" + "，".join(f"{k} {v}" for k, v in stats.items() if v))
    print(f"输出：{args.output}")
    return 0


# ---------- report：数据体检 ----------
DATE_RE = re.compile(r"\d{4}[-/年.]\d{1,2}[-/月.]\d{1,2}")
PHONE_RE = re.compile(r"^1[3-9]\d{9}$")


def _guess_type(values):
    """推断列类型：空/数字/日期/文本，返回 (类型名, 数字列表)"""
    nums, dates, texts = 0, 0, 0
    num_vals = []
    for v in values:
        s = str(v).strip() if v is not None else ""
        if s == "":
            continue
        n = to_number(v)
        if n is not None and not DATE_RE.fullmatch(s):
            nums += 1
            num_vals.append(n)
        elif DATE_RE.fullmatch(s):
            dates += 1
        else:
            texts += 1
    total = nums + dates + texts
    if total == 0:
        return "空列", []
    if dates / total >= 0.7:
        return "日期", []
    if nums / total >= 0.7:
        return "数字", num_vals
    return "文本", []


def cmd_report(args):
    """数据体检报告：行列/缺失/重复/类型/异常，一键出 Markdown"""
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    ncol = len(headers)
    nrow = len(rows)
    # 补齐短行
    norm_rows = [list(r) + [""] * (ncol - len(r)) for r in rows]

    lines = []
    lines.append(f"# 数据体检报告：{Path(args.input).name}")
    lines.append("")
    lines.append(f"- 行数：**{nrow}**　列数：**{ncol}**")

    # 整行重复
    seen, dup_full = set(), 0
    for r in norm_rows:
        key = tuple(str(c).strip() for c in r)
        if key in seen:
            dup_full += 1
        seen.add(key)
    lines.append(f"- 完全重复的行：**{dup_full}**" + ("　（用 `dedup` 一键去掉）" if dup_full else ""))

    # 全空行
    empty_rows = sum(1 for r in norm_rows if not any(str(c).strip() for c in r))
    if empty_rows:
        lines.append(f"- 全空行：**{empty_rows}**　（用 `clean` 一键去掉）")

    lines.append("")
    lines.append("## 逐列体检")
    lines.append("")
    lines.append("| 列名 | 类型 | 缺失 | 缺失率 | 重复值 | 带空格 | 备注 |")
    lines.append("|------|------|------|--------|--------|--------|------|")

    issues = []
    for i, h in enumerate(headers):
        vals = [r[i] if i < len(r) else "" for r in norm_rows]
        missing = sum(1 for v in vals if v is None or str(v).strip() == "")
        rate = missing / nrow * 100 if nrow else 0
        kind, num_vals = _guess_type(vals)

        # ID 类数字列（手机号等）：求和无意义，改标记
        str_vals = [str(v).strip() for v in vals if v is not None and str(v).strip()]
        id_like = str_vals and sum(1 for v in str_vals if PHONE_RE.fullmatch(v)) / len(str_vals) >= 0.8
        if id_like:
            kind = "号码"
            num_vals = []

        # 重复值（非空）
        cnt = {}
        for v in vals:
            s = str(v).strip() if v is not None else ""
            if s:
                cnt[s] = cnt.get(s, 0) + 1
        dup_vals = sum(c - 1 for c in cnt.values() if c > 1)

        # 首尾空格/全角空格
        spaced = sum(1 for v in vals
                     if isinstance(v, str) and v and (v != v.strip() or "\u3000" in v))

        notes = []
        if kind == "号码":
            uniq_rate = len(cnt) / max(len(str_vals), 1)
            if uniq_rate < 1 and cnt:
                notes.append(f"有 {len(str_vals) - len(cnt)} 个重复号码")
            else:
                notes.append("全表唯一")
        if kind == "数字" and num_vals:
            uniq = len(set(num_vals))
            if uniq == nrow - missing and missing == 0 and nrow > 1:
                notes.append("可作主键（全表唯一）")
            neg = sum(1 for n in num_vals if n < 0)
            if neg:
                notes.append(f"有 {neg} 个负数")
            if len(num_vals) >= 3:
                srt = sorted(num_vals)
                mid = srt[len(srt) // 2] if len(srt) % 2 else (srt[len(srt)//2 - 1] + srt[len(srt)//2]) / 2
                avg = sum(num_vals) / len(num_vals)
                mx, mn = max(num_vals), min(num_vals)
                if mid > 0 and mx > mid * 10:
                    notes.append(f"疑似离群值（最大 {mx:g} ≫ 中位数 {mid:g}）")
                notes.append(f"合计 {sum(num_vals):g}，范围 {mn:g}~{mx:g}")
        if kind == "日期":
            fmts = set()
            for v in vals:
                s = str(v).strip() if v is not None else ""
                if DATE_RE.fullmatch(s):
                    fmts.add(re.sub(r"\d", "N", s))
            if len(fmts) > 1:
                notes.append(f"日期格式不统一（{len(fmts)} 种），用 `clean --dates` 归一")
        if kind == "文本" and cnt:
            top = max(cnt.items(), key=lambda kv: kv[1])
            if top[1] >= 2 and nrow <= 200:
                notes.append(f"最常见：「{top[0][:12]}」×{top[1]}")
        if spaced:
            notes.append(f"{spaced} 个单元格首尾带空格，用 `clean` 修")
        if rate >= 50:
            notes.append("缺失过半，考虑删列或核实来源")
        elif rate >= 10:
            notes.append("缺失较多，可 `clean --fill` 填占位")
        if dup_vals and kind != "空列":
            issues.append((str(h), dup_vals))

        lines.append(f"| {h} | {kind} | {missing} | {rate:.0f}% | {dup_vals or '-'} "
                     f"| {spaced or '-'} | {'；'.join(notes)} |")

    # 关键列候选提示（像手机号/身份证的列提醒隐私合规）
    for i, h in enumerate(headers):
        vals = [str(r[i]).strip() if i < len(r) and r[i] is not None else "" for r in norm_rows]
        nonempty = [v for v in vals if v]
        if nonempty and sum(1 for v in nonempty if PHONE_RE.fullmatch(v)) / len(nonempty) >= 0.8:
            lines.append("")
            lines.append(f"> ⚠ 列「{h}」疑似手机号，属个人隐私数据：对外分发前请先脱敏，遵守个人信息保护法。")
            break

    # 总评
    lines.append("")
    score = 100
    if dup_full:
        score -= min(30, dup_full)
    if empty_rows:
        score -= min(10, empty_rows * 2)
    spaced_total = sum(1 for r in norm_rows for v in r
                     if isinstance(v, str) and v and v != v.strip())
    if spaced_total:
        score -= min(15, spaced_total)
    missing_total = sum(1 for r in norm_rows for v in r if v is None or str(v).strip() == "")
    if nrow * ncol:
        mr = missing_total / (nrow * ncol) * 100
        if mr > 20:
            score -= 20
        elif mr > 5:
            score -= 10
        elif mr > 1:
            score -= 5
    verdict = "很干净，直接用" if score >= 90 else ("小问题，`clean`+`dedup` 两条命令能修好" if score >= 70 else "问题较多，建议先清洗再用")
    lines.append(f"## 总评：{max(0, score)} 分 —— {verdict}")

    report = "\n".join(lines)
    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        Path(args.output).write_text(report, encoding="utf-8")
        print(f"报告已生成：{args.output}")
    else:
        print(report)
    return 0


def cmd_pivot(args):
    """二维交叉透视表（Excel 数据透视表平替）：行=--rows 列=--cols 值=--values:函数
    示例：pivot 订单.xlsx --rows 地区 --cols 产品 --values 金额:sum --margins
    """
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    headers = [str(h) for h in headers]

    def col_idx(name, label):
        if name not in headers:
            print(f"错误：表里没有{label}「{name}」，现有列：{headers}")
            sys.exit(2)
        return headers.index(name)

    # 解析值定义 列名:函数（默认 sum）
    if ":" in args.values:
        vcol, fn = args.values.rsplit(":", 1)
        vcol, fn = vcol.strip(), fn.strip().lower()
    else:
        vcol, fn = args.values.strip(), "sum"
    if fn not in AGG_FUNCS:
        print(f"错误：不支持的聚合函数「{fn}」，可选：{'/'.join(AGG_FUNCS)}")
        return 2
    vi = col_idx(vcol, "值列")

    row_cols = [c.strip() for c in args.rows.split(",") if c.strip()]
    if not row_cols:
        print("错误：--rows 不能为空")
        return 2
    row_idxs = [col_idx(c, "行字段") for c in row_cols]

    col_field = args.cols.strip()
    ci = col_idx(col_field, "列字段")

    def cell(r, i):
        v = r[i] if i < len(r) else ""
        return "" if v is None else str(v).strip()

    # 收集数据
    col_keys, col_order = {}, []
    row_keys, row_order = {}, []
    for r in rows:
        if not any(cell(r, i) for i in range(len(headers))):
            continue  # 跳过全空行
        ck = cell(r, ci) or "(空白)"
        rk = tuple(cell(r, i) or "(空白)" for i in row_idxs)
        if ck not in col_keys:
            col_keys[ck] = []
            col_order.append(ck)
        if rk not in row_keys:
            row_keys[rk] = []
            row_order.append(rk)
        v = r[vi] if vi < len(r) else ""
        n = to_number(v)
        if fn == "count":
            row_keys[rk].append((ck, v if str(v).strip() != "" else None))
        else:
            if n is not None:
                row_keys[rk].append((ck, n))

    # 聚合
    def agg(vals):
        if fn == "count":
            return len(vals)
        nums = [v for _, v in vals]
        if not nums:
            return 0 if fn in ("sum", "count") else ""
        res = AGG_FUNCS[fn](nums)
        return round(res, 4) if isinstance(res, float) else res

    # 排序：列键尝试数字序，否则字典序
    def sort_keys(keys):
        def numkey(k):
            try:
                return (0, float(k))
            except ValueError:
                return (1, k)
        return sorted(keys, key=numkey)
    col_order = sort_keys(col_order)
    if args.sort_rows:
        row_order = sorted(row_order)
    # 行合计（用于默认按合计降序排）
    if not args.keep_order:
        def row_total(rk):
            t = 0.0
            for ck, v in row_keys[rk]:
                if isinstance(v, (int, float)):
                    t += v
            return t if fn == "sum" else len(row_keys[rk])
        row_order = sorted(row_order, key=row_total, reverse=True)

    multi_row = len(row_cols) > 1
    out_headers = (row_cols if multi_row else [row_cols[0]]) + [f"{ck}" for ck in col_order]
    if args.margins:
        out_headers.append("总计")

    out_rows = []
    col_totals = [0.0] * len(col_order)
    col_counts = [0] * len(col_order)
    grand = 0.0
    grand_n = 0
    for rk in row_order:
        base = list(rk) if multi_row else [rk[0]]
        line = base + [""] * len(col_order)
        buckets = {}
        for ck, v in row_keys[rk]:
            buckets.setdefault(ck, []).append((ck, v))
        for j, ck in enumerate(col_order):
            if ck in buckets:
                val = agg(buckets[ck])
                line[len(base) + j] = val
                if isinstance(val, (int, float)):
                    if fn == "count":
                        col_counts[j] += val
                    else:
                        col_totals[j] += val
        if args.margins:
            if fn == "count":
                tot = sum(v for v in line[len(base):] if isinstance(v, (int, float)))
            else:
                tot = round(sum(v for v in line[len(base):] if isinstance(v, (int, float))), 4)
            line.append(tot)
            grand += tot
            grand_n += 1
        out_rows.append(line)

    if args.margins:
        total_line = (["总计"] + [""] * (len(row_cols) - 1) if multi_row else ["总计"])
        for j in range(len(col_order)):
            total_line.append(col_counts[j] if fn == "count" else round(col_totals[j], 4))
        if fn == "count":
            total_line.append(sum(col_counts))
        else:
            total_line.append(round(grand, 4))
        out_rows.append(total_line)

    write_table(Path(args.output), out_headers, out_rows)
    print(f"完成：{len(rows)} 行 → 透视表 {len(out_rows)} 行 × {len(out_headers)} 列"
          f"（行={'+'.join(row_cols)}，列={col_field}[{len(col_order)}种]，值={fn}({vcol})）")
    print(f"输出：{args.output}")
    return 0


def cmd_unpivot(args):
    """宽表转长表（逆透视，pivot 的逆操作）：把多列指标拆成「指标名+值」两列
    示例：unpivot 销售.xlsx --id 门店 --cols 1月,2月,3月 --name 月份 --value 销售额
    宽表：门店|1月|2月|3月 → 长表：门店|月份|销售额
    """
    headers, rows = read_table(args.input)
    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 2
    headers = [str(h) for h in headers]

    def col_idx(name, label):
        if name not in headers:
            print(f"错误：表里没有{label}「{name}」，现有列：{headers}")
            sys.exit(2)
        return headers.index(name)

    id_cols = [c.strip() for c in args.id.split(",") if c.strip()] if args.id else []
    id_idxs = [col_idx(c, "ID列") for c in id_cols]

    # 要拆的列：显式指定或自动推断（除 ID 外全部）
    if args.cols:
        melt_names = [c.strip() for c in args.cols.split(",") if c.strip()]
        melt_idxs = [col_idx(c, "指标列") for c in melt_names]
    else:
        melt_idxs = [i for i in range(len(headers)) if i not in id_idxs]
        melt_names = [headers[i] for i in melt_idxs]
    if not melt_idxs:
        print("错误：没有可拆的指标列（用 --cols 指定，或确认 --id 没把列全包进去）")
        return 2

    name_col = args.name or "指标"
    value_col = args.value or "值"
    if name_col in id_cols or value_col in id_cols:
        print(f"错误：新列名「{name_col}」/「{value_col}」不能与 ID 列重名")
        return 2

    skip_empty = not args.keep_empty
    out_headers = id_cols + [name_col, value_col]
    out_rows = []
    skipped = 0
    for r in rows:
        if not any((r[i] if i < len(r) else "") not in (None, "") for i in range(len(headers))):
            continue  # 全空行直接跳过
        id_vals = [(r[i] if i < len(r) else "") for i in id_idxs]
        for mi, mname in zip(melt_idxs, melt_names):
            v = r[mi] if mi < len(r) else ""
            if v is None or (isinstance(v, str) and not v.strip()):
                if skip_empty:
                    skipped += 1
                    continue
                v = ""
            out_rows.append(list(id_vals) + [mname, v])

    write_table(Path(args.output), out_headers, out_rows)
    print(f"完成：宽表 {len(rows)} 行 × {len(headers)} 列 → 长表 {len(out_rows)} 行"
          f"（ID={'+'.join(id_cols) or '无'}，拆了 {len(melt_names)} 列：{melt_names}）")
    if skipped:
        print(f"跳过空值单元格：{skipped} 个（--keep-empty 可保留）")
    print(f"输出：{args.output}")
    return 0


def cmd_watch(args):
    """监控文件夹，新表格自动汇入总表（零依赖轮询）

    机制：
    - 每 interval 秒扫描一次目录，用 mtime+size 判断新增/修改
    - 有变化则全量重建输出文件（简单可靠，附带来源文件列）
    - 状态记录在输出文件旁 .biaoge-watch-state.json
    - 自动跳过：输出文件本身、状态文件、~$ 临时文件、隐藏文件
    - --once 单次模式：扫一遍就退出（可配合 cron/任务计划程序）
    """
    import json, time
    watch_dir = Path(args.folder).expanduser().resolve()
    if not watch_dir.is_dir():
        print(f"目录不存在：{watch_dir}"); return 1
    out_path = Path(args.output).expanduser().resolve()
    state_path = out_path.parent / ".biaoge-watch-state.json"

    def scan():
        files = collect_files([watch_dir], args.recursive)
        return [f for f in files
                if f.resolve() != out_path and f.resolve() != state_path
                and not f.name.startswith("~$") and not f.name.startswith(".")]

    def fingerprint(files):
        fp = {}
        for f in files:
            try:
                st = f.stat()
                fp[str(f)] = [st.st_mtime, st.st_size]
            except OSError:
                continue
        return fp

    def load_state():
        try:
            return json.loads(state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def rebuild(files):
        """全量重建汇总表，返回 (文件数, 行数)"""
        headers = None
        all_rows = []
        for f in files:
            try:
                h, rows = read_table(f)
            except Exception as e:
                print(f"  ⚠ 跳过读不了的文件 {f.name}: {e}", flush=True)
                continue
            if not h:
                continue
            if headers is None:
                headers = h + ["来源文件"]
            for r in rows:
                all_rows.append(list(r) + [f.name])
        if headers is None:
            return 0, 0
        # 先写临时文件再替换，避免写一半被读
        tmp = out_path.with_suffix(out_path.suffix + ".tmp")
        write_table(tmp, headers, all_rows)
        os.replace(tmp, out_path)
        return len(files), len(all_rows)

    state = load_state()
    files = scan()
    fp = fingerprint(files)
    if fp != state or not out_path.exists():
        nf, nr = rebuild(files)
        state = fp
        state_path.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"📊 已汇总 {nf} 个文件 / {nr} 行 → {out_path}", flush=True)
    else:
        print(f"无变化（{len(files)} 个文件已汇总）→ {out_path}", flush=True)

    if args.once:
        return 0

    print(f"👀 监控中：{watch_dir}（每 {args.interval}s 扫描，Ctrl+C 退出）", flush=True)
    try:
        while True:
            time.sleep(args.interval)
            files = scan()
            fp = fingerprint(files)
            if fp == state:
                continue
            # 找出变化项，日志友好
            added = [Path(k).name for k in fp if k not in state]
            changed = [Path(k).name for k in fp if k in state and fp[k] != state[k]]
            removed = [Path(k).name for k in state if k not in fp]
            for n in added: print(f"  ➕ 新文件：{n}", flush=True)
            for n in changed: print(f"  ✏️ 修改：{n}", flush=True)
            for n in removed: print(f"  ➖ 移除：{n}", flush=True)
            nf, nr = rebuild(files)
            state = fp
            state_path.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
            ts = time.strftime("%H:%M:%S")
            print(f"[{ts}] 📊 重新汇总 {nf} 个文件 / {nr} 行 → {out_path}", flush=True)
    except KeyboardInterrupt:
        print("\n已停止监控。", flush=True)
    return 0



def _read_sheet_for_sql(path, sheet_name=None):
    """Read a sheet from xlsx/csv for sql command, return (headers, rows)"""
    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
        data = [[("" if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    elif ext == ".csv":
        with open(p, newline="", encoding="utf-8-sig") as f:
            data = [row for row in csv.reader(f)]
    else:
        raise ValueError(f"不支持的文件类型: {p.name}（sql 命令支持 .xlsx/.csv）")
    if not data:
        return [], []
    headers = [str(h).strip() for h in data[0]]
    return headers, data[1:]


def _infer_column_types(headers, rows):
    """Infer column types: INTEGER if all values are int, REAL if all numeric, else TEXT"""
    types = []
    for i, h in enumerate(headers):
        col_vals = []
        for r in rows:
            if i < len(r) and r[i] is not None and str(r[i]).strip() != "":
                col_vals.append(r[i])
        if not col_vals:
            types.append("TEXT")
            continue
        all_int = True
        all_num = True
        for v in col_vals:
            if isinstance(v, bool):
                all_int = False
                all_num = False
                break
            if isinstance(v, int):
                continue
            if isinstance(v, float):
                all_int = False
                if v != int(v):
                    all_num = True
                    continue
                continue
            s = str(v).strip()
            try:
                n = float(s)
                if n != int(n):
                    all_int = False
            except (ValueError, TypeError):
                all_int = False
                all_num = False
                break
        if all_int:
            types.append("INTEGER")
        elif all_num:
            types.append("REAL")
        else:
            types.append("TEXT")
    return types


def cmd_sql(args):
    """用 SQL 查询 Excel/CSV（sqlite3 零依赖实现）"""
    sql_text = args.query.strip()

    # 只允许 SELECT 语句
    upper = sql_text.upper().lstrip()
    if not upper.startswith("SELECT") and not upper.startswith("WITH"):
        print("错误：sql 命令只支持 SELECT 查询（这是查询工具，不是写库工具）")
        print("支持的语句：SELECT ... FROM data WHERE ... GROUP BY ... ORDER BY ...")
        return 1

    # 读取数据
    try:
        headers, rows = _read_sheet_for_sql(args.input, getattr(args, 'sheet', None))
    except Exception as e:
        print(f"错误：读取文件失败 - {e}")
        return 1

    if not headers:
        print("错误：读不到表头（空文件？）")
        return 1

    # 类型推断
    col_types = _infer_column_types(headers, rows)

    # 建内存 sqlite 表
    conn = sqlite3.connect(":memory:")
    try:
        # 建表：列名用双引号包裹以支持中文和特殊字符
        col_defs = ", ".join(
            f'"{h}" {t}' for h, t in zip(headers, col_types)
        )
        conn.execute(f'CREATE TABLE data ({col_defs})')

        # 插入数据
        placeholders = ", ".join(["?"] * len(headers))
        insert_sql = f'INSERT INTO data VALUES ({placeholders})'
        for r in rows:
            padded = list(r) + [None] * (len(headers) - len(r))
            vals = []
            for i, v in enumerate(padded[:len(headers)]):
                if v is None or (isinstance(v, str) and v.strip() == ""):
                    vals.append(None)
                else:
                    vals.append(v)
            conn.execute(insert_sql, vals)
        conn.commit()

        # 执行查询
        try:
            cursor = conn.execute(sql_text)
        except sqlite3.OperationalError as e:
            msg = str(e)
            # 列名不存在时给出友好提示
            if "no such column" in msg.lower() or "has no column named" in msg.lower():
                print(f"错误：SQL 列名不存在 - {msg}")
                print(f"可用列名：{headers}")
                return 1
            print(f"错误：SQL 执行失败 - {msg}")
            print(f"可用列名：{headers}")
            return 1
        except sqlite3.Error as e:
            print(f"错误：SQL 语法错误 - {e}")
            return 1

        # 获取结果
        result_headers = [desc[0] for desc in cursor.description] if cursor.description else []
        result_rows = cursor.fetchall()

        # --print 预览模式
        if args.print_rows is not None:
            n = args.print_rows
            if not result_headers:
                print("查询结果为空")
                return 0
            print("	".join(result_headers))
            for r in result_rows[:n]:
                print("	".join(str(v) if v is not None else "" for v in r))
            total = len(result_rows)
            if total > n:
                print(f"... 共 {total} 行，已显示前 {n} 行")
            else:
                print(f"共 {total} 行")
            return 0

        # 输出到文件
        if args.output:
            write_table(args.output, result_headers, [list(r) for r in result_rows])
            print(f"✅ SQL 查询完成：{len(result_rows)} 行 → {args.output}")
        else:
            # 默认打印到终端
            if not result_headers:
                print("查询结果为空")
                return 0
            print("	".join(result_headers))
            for r in result_rows:
                print("	".join(str(v) if v is not None else "" for v in r))
            print(f"共 {len(result_rows)} 行")

    finally:
        conn.close()
    return 0


def _normalize_key(v):
    """key 归一化：去首尾空格 + 大小写不敏感"""
    if v is None:
        return ""
    return str(v).strip().lower()


def _normalize_value(v):
    """值归一化：数字 1.0 == 1 == '1'，避免类型误报"""
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        # 整数型浮点转 int（1.0 → 1）
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if s == "":
        return ""
    n = to_number(s)
    if n is not None:
        if isinstance(n, float) and n == int(n):
            return str(int(n))
        return str(n)
    return s


def cmd_diff(args):
    """两表对比：按 --key 列对齐，找出新增/删除/修改"""
    old_h, old_rows = read_table(args.old)
    new_h, new_rows = read_table(args.new)

    if not old_h or not new_h:
        print("错误：读不到表头（空文件？）"); return 1

    # key 列检查
    if args.key not in old_h:
        print(f"错误：旧表里没有关键列「{args.key}」，现有列：{old_h}"); return 1
    if args.key not in new_h:
        print(f"错误：新表里没有关键列「{args.key}」，现有列：{new_h}"); return 1

    old_key_idx = old_h.index(args.key)
    new_key_idx = new_h.index(args.key)

    # 对比列：缺省=两表共同列（排除 key）
    if args.columns:
        cols = [c.strip() for c in args.columns.split(",") if c.strip()]
        missing_old = [c for c in cols if c not in old_h]
        missing_new = [c for c in cols if c not in new_h]
        if missing_old:
            print(f"错误：旧表里没有列 {missing_old}，现有列：{old_h}"); return 1
        if missing_new:
            print(f"错误：新表里没有列 {missing_new}，现有列：{new_h}"); return 1
    else:
        cols = [c for c in old_h if c != args.key and c in new_h]
        if not cols:
            print(f"错误：两表没有共同列可对比（旧表列：{old_h}，新表列：{new_h}）"); return 1

    old_col_idxs = {c: old_h.index(c) for c in cols}
    new_col_idxs = {c: new_h.index(c) for c in cols}

    # 建索引：key → 行（重复 key 取第一条，发警告）
    def build_index(headers, rows, key_idx, col_idxs):
        idx = {}
        dup_count = 0
        for r in rows:
            raw_key = r[key_idx] if key_idx < len(r) else ""
            k = _normalize_key(raw_key)
            if k in idx:
                dup_count += 1
                continue
            vals = {}
            for c, ci in col_idxs.items():
                v = r[ci] if ci < len(r) else ""
                vals[c] = v
            idx[k] = (raw_key, vals)
        return idx, dup_count

    old_idx, old_dup = build_index(old_h, old_rows, old_key_idx, old_col_idxs)
    new_idx, new_dup = build_index(new_h, new_rows, new_key_idx, new_col_idxs)

    if old_dup:
        print(f"⚠ 警告：旧表有 {old_dup} 个重复 key，已取第一条")
    if new_dup:
        print(f"⚠ 警告：新表有 {new_dup} 个重复 key，已取第一条")

    old_keys = set(old_idx.keys())
    new_keys = set(new_idx.keys())

    added_keys = new_keys - old_keys
    removed_keys = old_keys - new_keys
    common_keys = old_keys & new_keys

    # 找 modified
    modified = []  # [(key_display, col, old_val, new_val), ...]
    modified_keys = set()
    for k in common_keys:
        old_raw, old_vals = old_idx[k]
        new_raw, new_vals = new_idx[k]
        for c in cols:
            ov = _normalize_value(old_vals[c])
            nv = _normalize_value(new_vals[c])
            if ov != nv:
                modified.append((old_raw, c, old_vals[c], new_vals[c]))
                modified_keys.add(k)

    added_rows = [(new_idx[k][0], new_idx[k][1]) for k in added_keys]
    removed_rows = [(old_idx[k][0], old_idx[k][1]) for k in removed_keys]

    # 终端摘要
    print(f"📊 对比结果：旧表 {len(old_rows)} 行，新表 {len(new_rows)} 行")
    print(f"   新增：{len(added_keys)} 行")
    print(f"   删除：{len(removed_keys)} 行")
    print(f"   修改：{len(modified_keys)} 个 key，{len(modified)} 处变化")

    show_n = 5
    if added_rows:
        print(f"\n【新增前 {min(show_n, len(added_rows))} 条】")
        for raw_key, vals in added_rows[:show_n]:
            print(f"   + {raw_key}")
    if removed_rows:
        print(f"\n【删除前 {min(show_n, len(removed_rows))} 条】")
        for raw_key, vals in removed_rows[:show_n]:
            print(f"   - {raw_key}")
    if modified:
        print(f"\n【修改前 {min(show_n * 2, len(modified))} 处】")
        for key_disp, col, ov, nv in modified[:show_n * 2]:
            print(f"   ~ {key_disp}  {col}: {ov!r} → {nv!r}")

    # -o 输出 xlsx
    if args.output:
        wb = openpyxl.Workbook()
        # 摘要 sheet
        ws_sum = wb.active
        ws_sum.title = "摘要"
        ws_sum.append(["项目", "数量"])
        ws_sum.append(["旧表行数", len(old_rows)])
        ws_sum.append(["新表行数", len(new_rows)])
        ws_sum.append(["新增", len(added_keys)])
        ws_sum.append(["删除", len(removed_keys)])
        ws_sum.append(["修改（key数）", len(modified_keys)])
        ws_sum.append(["修改（变化处）", len(modified)])

        # 新增 sheet
        ws_add = wb.create_sheet("新增")
        add_headers = [args.key] + cols
        ws_add.append(add_headers)
        for raw_key, vals in added_rows:
            ws_add.append([raw_key] + [vals[c] for c in cols])

        # 删除 sheet
        ws_del = wb.create_sheet("删除")
        del_headers = [args.key] + cols
        ws_del.append(del_headers)
        for raw_key, vals in removed_rows:
            ws_del.append([raw_key] + [vals[c] for c in cols])

        # 修改 sheet：key, 列名, 旧值, 新值
        ws_mod = wb.create_sheet("修改")
        ws_mod.append([args.key, "列名", "旧值", "新值"])
        for key_disp, col, ov, nv in modified:
            ws_mod.append([key_disp, col, ov, nv])

        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
        print(f"\n✅ 对比结果已输出：{args.output}（4 个 sheet：摘要/新增/删除/修改）")

    return 0



def _setup_cjk_font():
    """配置 CJK 字体回退链，支持中文标题/标签"""
    import matplotlib.font_manager as fm
    import matplotlib.pyplot as plt
    candidates = [
        "Microsoft YaHei",
        "PingFang SC", "PingFang TC",
        "Noto Sans CJK SC", "Noto Sans CJK TC", "Noto Sans CJK JP",
        "Source Han Sans SC", "Source Han Sans CN",
        "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",
        "SimHei",
        "STHeiti",
        "Arial Unicode MS",
    ]
    available = {f.name for f in fm.fontManager.ttflist}
    for name in candidates:
        if name in available:
            plt.rcParams["font.sans-serif"] = [name] + list(plt.rcParams["font.sans-serif"])
            plt.rcParams["axes.unicode_minus"] = False
            return name
    import warnings
    warnings.warn("未找到 CJK 字体，中文标签可能显示为乱码。可安装 fonts-noto-cjk 或微软雅黑。")
    plt.rcParams["axes.unicode_minus"] = False
    return None


def _load_data_for_chart(input_path, sheet_name, sql_text=None, x_col=None, y_cols=None):
    """加载图表数据，返回 (x_labels, y_data_dict, y_col_names)。"""
    if sql_text:
        headers, rows = _read_sheet_for_sql(input_path, sheet_name)
        if not headers:
            raise ValueError("读不到表头（空文件？）")
        col_types = _infer_column_types(headers, rows)
        conn = sqlite3.connect(":memory:")
        try:
            col_defs = ", ".join(f'"{h}" {t}' for h, t in zip(headers, col_types))
            conn.execute(f'CREATE TABLE data ({col_defs})')
            placeholders = ", ".join(["?"] * len(headers))
            insert_sql = f'INSERT INTO data VALUES ({placeholders})'
            for r in rows:
                padded = list(r) + [None] * (len(headers) - len(r))
                vals = []
                for i, v in enumerate(padded[:len(headers)]):
                    if v is None or (isinstance(v, str) and v.strip() == ""):
                        vals.append(None)
                    else:
                        vals.append(v)
                conn.execute(insert_sql, vals)
            conn.commit()
            cursor = conn.execute(sql_text)
            result_headers = [desc[0] for desc in cursor.description] if cursor.description else []
            result_rows = cursor.fetchall()
        finally:
            conn.close()
        if not result_headers or not result_rows:
            raise ValueError("SQL 查询结果为空，无法出图")
        x_labels = [str(r[0]) for r in result_rows]
        y_data = {}
        for i in range(1, len(result_headers)):
            vals = []
            for r in result_rows:
                v = r[i]
                if v is None:
                    vals.append(0.0)
                elif isinstance(v, (int, float)):
                    vals.append(float(v))
                else:
                    try:
                        vals.append(float(str(v).strip()))
                    except (ValueError, TypeError):
                        vals.append(0.0)
            y_data[result_headers[i]] = vals
        return x_labels, y_data, list(y_data.keys())
    else:
        headers, rows = read_table(input_path)
        if not headers:
            raise ValueError("读不到表头（空文件？）")
        headers = [str(h).strip() for h in headers]
        if x_col not in headers:
            raise ValueError(f"列 '{x_col}' 不存在。可用列：{headers}")
        xi = headers.index(x_col)
        x_labels = [str(r[xi]) if xi < len(r) else "" for r in rows]
        y_data = {}
        for yc in y_cols:
            if yc not in headers:
                raise ValueError(f"列 '{yc}' 不存在。可用列：{headers}")
            ci = headers.index(yc)
            vals = []
            for r in rows:
                v = r[ci] if ci < len(r) else None
                if v is None or (isinstance(v, str) and v.strip() == ""):
                    vals.append(0.0)
                elif isinstance(v, (int, float)):
                    vals.append(float(v))
                else:
                    n = to_number(v)
                    vals.append(float(n) if n is not None else 0.0)
            y_data[yc] = vals
        return x_labels, y_data, list(y_cols)


def cmd_chart(args):
    """从 Excel/CSV 数据生成图表（PNG）"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        print("错误：chart 命令需要 matplotlib，请运行：")
        print("  pip install biaoge[chart]")
        print("或：")
        print("  pip install matplotlib")
        return 1

    chart_type = args.type
    out_path = Path(args.out)
    title = args.title or ""
    theme = getattr(args, "theme", None)

    sql_text = getattr(args, "sql", None)
    x_col = getattr(args, "x", None)
    y_cols_str = getattr(args, "y", None)

    if sql_text:
        sql_text = sql_text.strip()
        upper = sql_text.upper().lstrip()
        if not upper.startswith("SELECT") and not upper.startswith("WITH"):
            print("错误：--sql 只支持 SELECT 查询")
            return 1
    elif x_col and y_cols_str:
        y_cols = [c.strip() for c in y_cols_str.split(",") if c.strip()]
        if not y_cols:
            print("错误：--y 至少指定一个列名")
            return 1
    else:
        print("错误：请指定 --sql 查询语句，或 --x 和 --y 列名")
        return 1

    try:
        if sql_text:
            x_labels, y_data, y_names = _load_data_for_chart(
                args.input, getattr(args, "sheet", None), sql_text=sql_text)
        else:
            x_labels, y_data, y_names = _load_data_for_chart(
                args.input, getattr(args, "sheet", None), x_col=x_col, y_cols=y_cols)
    except Exception as e:
        print(f"错误：{e}")
        return 1

    if not x_labels:
        print("错误：数据为空，无法出图")
        return 1

    _setup_cjk_font()

    if theme:
        try:
            plt.style.use(theme)
        except OSError:
            print(f"警告：主题 '{theme}' 不存在，使用默认主题")

    fig, ax = plt.subplots(figsize=(10, 6))

    if chart_type == "pie":
        if len(x_labels) > 8:
            y_name = y_names[0]
            values = list(zip(x_labels, y_data[y_name]))
            values.sort(key=lambda t: t[1], reverse=True)
            top7 = values[:7]
            other_sum = sum(v for _, v in values[7:])
            labels = [t[0] for t in top7] + ["其他"]
            sizes = [t[1] for t in top7] + [other_sum]
        else:
            y_name = y_names[0]
            labels = x_labels
            sizes = y_data[y_name]
        ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=90)
        ax.set_title(title or y_names[0], fontsize=14)
    elif chart_type == "bar":
        x_pos = list(range(len(x_labels)))
        width = 0.8 / len(y_names)
        for i, yn in enumerate(y_names):
            offset = (i - len(y_names) / 2 + 0.5) * width
            ax.bar([p + offset for p in x_pos], y_data[yn], width=width, label=yn)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(x_labels, rotation=45 if len(x_labels) > 6 else 0,
                           ha="right" if len(x_labels) > 6 else "center")
        ax.set_title(title, fontsize=14)
        ax.set_xlabel(x_col or "")
        ax.set_ylabel(", ".join(y_names))
        if len(y_names) > 1:
            ax.legend()
    elif chart_type == "line":
        for yn in y_names:
            ax.plot(x_labels, y_data[yn], marker="o", label=yn)
        if len(x_labels) > 6:
            plt.xticks(rotation=45, ha="right")
        ax.set_title(title, fontsize=14)
        ax.set_xlabel(x_col or "")
        ax.set_ylabel(", ".join(y_names))
        if len(y_names) > 1:
            ax.legend()
    elif chart_type == "scatter":
        if len(y_names) < 1:
            print("错误：scatter 图至少需要一个 y 列")
            return 1
        y_name = y_names[0]
        x_nums = []
        for v in x_labels:
            try:
                x_nums.append(float(v))
            except (ValueError, TypeError):
                x_nums.append(None)
        if any(n is None for n in x_nums):
            print("错误：scatter 图的 x 列必须为数值")
            return 1
        ax.scatter(x_nums, y_data[y_name], alpha=0.7)
        ax.set_title(title, fontsize=14)
        ax.set_xlabel(x_col or "")
        ax.set_ylabel(y_name)
    else:
        print(f"错误：不支持的图表类型 '{chart_type}'（支持 bar/line/pie/scatter）")
        return 1

    plt.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(out_path), dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"✅ 图表已生成：{out_path.resolve()}")
    return 0



# ---------- template ----------
def _fmt_cell_value(v):
    if v is None:
        return ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v).strip()
    if s.lower() in ("none", "nan", ""):
        return ""
    return s


def _sanitize_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name)
    name = name.strip('. ')
    return name or "output"


def _merge_runs_replace(paragraph, mapping):
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return False
    changed = False
    for key, val in mapping.items():
        placeholder = "{{" + key + "}}"
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, val)
            changed = True
    if changed:
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ""
    return changed


def _replace_in_docx(doc, mapping):
    for para in doc.paragraphs:
        _merge_runs_replace(para, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _merge_runs_replace(para, mapping)
    for section in doc.sections:
        for para in section.header.paragraphs:
            _merge_runs_replace(para, mapping)
        for para in section.footer.paragraphs:
            _merge_runs_replace(para, mapping)


def _replace_in_xlsx(wb, mapping):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "{{" in cell.value:
                    val = cell.value
                    for key, r_val in mapping.items():
                        val = val.replace("{{" + key + "}}", r_val)
                    cell.value = val


def _find_placeholders_in_docx(doc):
    found = set()
    pat = re.compile(r"\{\{([^}]+)\}\}")
    for para in doc.paragraphs:
        full = "".join(r.text for r in para.runs)
        found.update(pat.findall(full))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full = "".join(r.text for r in para.runs)
                    found.update(pat.findall(full))
    for section in doc.sections:
        for para in section.header.paragraphs:
            full = "".join(r.text for r in para.runs)
            found.update(pat.findall(full))
        for para in section.footer.paragraphs:
            full = "".join(r.text for r in para.runs)
            found.update(pat.findall(full))
    return found


def _find_placeholders_in_xlsx(wb):
    found = set()
    pat = re.compile(r"\{\{([^}]+)\}\}")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    found.update(pat.findall(cell.value))
    return found


def cmd_template(args):
    """数据驱动批量文档生成（邮件合并平替）"""
    data_path = Path(args.data)
    tpl_path = Path(args.tpl)
    out_dir = Path(args.out)
    tpl_ext = tpl_path.suffix.lower()

    if not data_path.exists():
        print(f"错误：数据表不存在：{data_path}")
        return 1
    if not tpl_path.exists():
        print(f"错误：模板文件不存在：{tpl_path}")
        return 1
    if tpl_ext not in (".docx", ".xlsx"):
        print(f"错误：模板只支持 .docx 或 .xlsx，当前：{tpl_ext}")
        return 1
    if tpl_ext == ".docx" and DocxDocument is None:
        print("错误：docx 模板需要 python-docx，请运行：pip install python-docx")
        return 1

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        test_file = out_dir / ".biaoge_write_test"
        test_file.write_text("test")
        test_file.unlink()
    except Exception:
        print(f"错误：输出目录不可写：{out_dir}")
        return 1

    sheet_name = getattr(args, "sheet", None)
    if sheet_name and data_path.suffix.lower() in (".xlsx", ".xlsm"):
        headers, rows = _read_sheet(data_path, sheet_name)
    else:
        headers, rows = read_table(data_path)
    headers = [str(h) for h in headers]

    if not headers:
        print("错误：数据表读不到表头（空文件？）")
        return 1

    if tpl_ext == ".docx":
        tpl_doc = DocxDocument(str(tpl_path))
        placeholders = _find_placeholders_in_docx(tpl_doc)
    else:
        tpl_doc = openpyxl.load_workbook(str(tpl_path))
        placeholders = _find_placeholders_in_xlsx(tpl_doc)

    if not placeholders:
        print("⚠ 警告：模板中没有找到 {{占位符}}，生成的文档将与模板完全相同")

    missing_cols = placeholders - set(headers)
    if missing_cols:
        msg = f"⚠ 警告：模板中的占位符在数据表里没有对应列：{sorted(missing_cols)}（将保留原样）"
        if getattr(args, "strict", False):
            print(f"错误：{msg}")
            return 1
        print(msg)

    filename_tpl = getattr(args, "filename_tpl", None)
    name_col = getattr(args, "name_col", None)

    if getattr(args, "dry_run", False):
        n_preview = min(3, len(rows))
        print(f"📋 预览模式（前 {n_preview} 行）：")
        print(f"   数据表：{data_path.name}（{len(rows)} 行 x {len(headers)} 列）")
        ph_str = str(sorted(placeholders)) if placeholders else "无"
        print(f"   模板：{tpl_path.name}（占位符：{ph_str}）")
        print(f"   输出目录：{out_dir}")
        print()
        for i, row in enumerate(rows[:3]):
            mapping = {}
            for h in headers:
                idx = headers.index(h)
                v = row[idx] if idx < len(row) else ""
                mapping[h] = _fmt_cell_value(v)
            if filename_tpl:
                try:
                    fname = filename_tpl.format(**mapping)
                except KeyError:
                    fname = str(i + 1)
            elif name_col and name_col in headers:
                fname = mapping.get(name_col, str(i + 1))
            else:
                fname = str(i + 1)
            fname = _sanitize_filename(fname) + tpl_ext
            print(f"   行 {i+1} -> {fname}")
            for k in sorted(mapping.keys()):
                if k in placeholders:
                    print(f"     {{{{{k}}}}} -> {mapping[k]}")
        return 0

    used_names = {}
    generated = 0
    for i, row in enumerate(rows):
        mapping = {}
        for h in headers:
            idx = headers.index(h)
            v = row[idx] if idx < len(row) else ""
            mapping[h] = _fmt_cell_value(v)

        if filename_tpl:
            try:
                fname = filename_tpl.format(**mapping)
            except KeyError:
                fname = str(i + 1)
        elif name_col and name_col in headers:
            fname = mapping.get(name_col, str(i + 1))
        else:
            fname = str(i + 1)
        fname = _sanitize_filename(fname)

        base_name = fname
        if base_name in used_names:
            used_names[base_name] += 1
            fname = f"{base_name}_{used_names[base_name]}{tpl_ext}"
        else:
            used_names[base_name] = 0
            fname = base_name + tpl_ext

        out_path = out_dir / fname

        if tpl_ext == ".docx":
            doc = DocxDocument(str(tpl_path))
            _replace_in_docx(doc, mapping)
            doc.save(str(out_path))
        else:
            wb = openpyxl.load_workbook(str(tpl_path))
            _replace_in_xlsx(wb, mapping)
            wb.save(str(out_path))
            wb.close()

        generated += 1

    print(f"✅ 批量生成完成：{generated} 个文件 -> {out_dir}/")
    return 0



# ---------- fill 空值向下填充 ----------
def cmd_fill(args):
    """空值向下填充（ffill），支持合并单元格拆分、分组填充、间隙限制"""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在：{input_path}")
        return 1

    is_xlsx = input_path.suffix.lower() in ('.xlsx', '.xlsm')
    is_csv = input_path.suffix.lower() == '.csv'

    if args.unmerge and is_csv:
        print("提示：CSV 文件没有合并单元格，--unmerge 已忽略")

    if not args.output and not args.report:
        print("错误：请指定 -o 输出文件，或使用 --report 仅查看报告")
        return 1

    col_specs = [c.strip() for c in args.cols.split(',')] if args.cols else None

    if is_xlsx:
        return _fill_xlsx(input_path, args, col_specs)
    else:
        return _fill_csv(input_path, args, col_specs)


def _resolve_col_indices(headers, col_specs):
    """把列名或字母（A/B/C）解析为 0-based 列索引列表"""
    indices = []
    for spec in col_specs:
        if re.match(r"^[A-Za-z]{1,3}$", spec):
            idx = 0
            for ch in spec.upper():
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            idx -= 1
            if idx >= len(headers):
                raise ValueError(f"错误：列字母「{spec}」超出范围（共 {len(headers)} 列）")
            indices.append(idx)
        else:
            if spec not in headers:
                avail = ", ".join(str(h) for h in headers)
                raise ValueError(f"错误：找不到列「{spec}」，可用列：{avail}")
            indices.append(headers.index(spec))
    return indices


def _do_fill_rows(headers, rows, fill_col_indices, group_by_idx, min_gap):
    """对 rows 就地 ffill，返回 (filled_count, remaining_total, remaining_per_col)"""
    n_cols = len(headers)
    filled_total = 0
    remaining_per_col = [0] * n_cols

    data_col_indices = [ci for ci in fill_col_indices if ci != group_by_idx]

    original_group_values = None
    if group_by_idx is not None:
        original_group_values = []
        for row in rows:
            cell = row[group_by_idx] if group_by_idx < len(row) else None
            original_group_values.append(cell if cell is not None and cell != "" else None)

    if group_by_idx is not None and group_by_idx in fill_col_indices:
        group_val = None
        for row in rows:
            cell = row[group_by_idx] if group_by_idx < len(row) else None
            if cell is not None and cell != "":
                group_val = cell
            elif group_val is not None:
                row[group_by_idx] = group_val
                filled_total += 1
            else:
                remaining_per_col[group_by_idx] += 1

    for ci in data_col_indices:
        if ci >= n_cols:
            continue
        last_val = None
        empty_start = None
        current_group = None
        for ri, row in enumerate(rows):
            is_group_boundary = False
            if group_by_idx is not None:
                orig_gv = original_group_values[ri] if ri < len(original_group_values) else None
                if orig_gv is not None:
                    if current_group is None or orig_gv != current_group:
                        is_group_boundary = True
                        current_group = orig_gv
                        if empty_start is not None:
                            run_len = ri - empty_start
                            if last_val is not None and (min_gap is None or run_len <= min_gap):
                                for j in range(empty_start, ri):
                                    rows[j][ci] = last_val
                                    filled_total += 1
                            else:
                                remaining_per_col[ci] += run_len
                            empty_start = None
                        last_val = None

            if ci >= len(row):
                continue
            cell = row[ci]
            if cell is None or cell == "":
                if empty_start is None and not is_group_boundary:
                    empty_start = ri
            else:
                if empty_start is not None:
                    run_len = ri - empty_start
                    if last_val is not None and (min_gap is None or run_len <= min_gap):
                        for j in range(empty_start, ri):
                            rows[j][ci] = last_val
                            filled_total += 1
                    else:
                        remaining_per_col[ci] += run_len
                    empty_start = None
                last_val = cell
        if empty_start is not None:
            run_len = len(rows) - empty_start
            if last_val is not None and (min_gap is None or run_len <= min_gap):
                for j in range(empty_start, len(rows)):
                    rows[j][ci] = last_val
                    filled_total += 1
            else:
                remaining_per_col[ci] += run_len

    remaining_total = sum(remaining_per_col[ci] for ci in fill_col_indices)
    return filled_total, remaining_total, remaining_per_col


def _fill_xlsx(input_path, args, col_specs):
    wb = openpyxl.load_workbook(input_path)

    if args.sheet:
        sheet_ref = args.sheet
        if sheet_ref in wb.sheetnames:
            target_sheets = [sheet_ref]
        else:
            try:
                idx = int(sheet_ref)
                if 0 <= idx < len(wb.sheetnames):
                    target_sheets = [wb.sheetnames[idx]]
                else:
                    print(f"错误：sheet 索引 {idx} 超出范围（共 {len(wb.sheetnames)} 个 sheet）")
                    wb.close()
                    return 1
            except ValueError:
                print(f"错误：找不到 sheet「{sheet_ref}」")
                wb.close()
                return 1
    else:
        target_sheets = list(wb.sheetnames)

    report_lines = []
    total_filled = 0
    total_remaining = 0

    auto_unmerge_notified = False
    for sheet_name in target_sheets:
        ws = wb[sheet_name]

        has_merged = len(ws.merged_cells.ranges) > 0
        if has_merged and not args.unmerge and not auto_unmerge_notified:
            import sys
            print("检测到合并单元格，已自动拆分填充", file=sys.stderr)
            auto_unmerge_notified = True

        if args.unmerge or has_merged:
            for merge_range in list(ws.merged_cells.ranges):
                top_left_value = ws.cell(merge_range.min_row, merge_range.min_col).value
                ws.unmerge_cells(str(merge_range))
                for row_idx in range(merge_range.min_row, merge_range.max_row + 1):
                    for col_idx in range(merge_range.min_col, merge_range.max_col + 1):
                        ws.cell(row_idx, col_idx).value = top_left_value

        data_rows = list(ws.iter_rows(values_only=True))
        if not data_rows:
            continue
        headers = list(data_rows[0])
        rows = [list(r) for r in data_rows[1:]]

        try:
            if col_specs:
                fill_indices = _resolve_col_indices(headers, col_specs)
            else:
                fill_indices = list(range(len(headers)))
        except ValueError as e:
            print(str(e))
            wb.close()
            return 1

        group_by_idx = None
        if args.group_by:
            gb = args.group_by
            if gb not in headers:
                avail = ", ".join(str(h) for h in headers)
                print(f"错误：找不到分组列「{gb}」，可用列：{avail}")
                wb.close()
                return 1
            group_by_idx = headers.index(gb)
            if group_by_idx not in fill_indices:
                fill_indices.append(group_by_idx)

        min_gap = args.min_gap

        filled, remaining, remaining_per_col = _do_fill_rows(
            headers, rows, fill_indices, group_by_idx, min_gap)

        total_filled += filled
        total_remaining += remaining

        if args.report:
            report_lines.append(f"📋 Sheet「{sheet_name}」:")
            for ci in sorted(fill_indices):
                col_name = headers[ci]
                col_remaining = remaining_per_col[ci]
                report_lines.append(f"  列「{col_name}」: 剩余空值 {col_remaining}")
            report_lines.append("")
            continue

        for row_idx, row_data in enumerate(rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row_idx, col_idx).value = value

    if args.report:
        report_text = "\n".join(report_lines)
        print(report_text)
        if args.output:
            out_path = Path(args.output)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(report_text, encoding="utf-8")
        wb.close()
        return 0

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    wb.close()
    print(f"✅ 填充完成：共填充 {total_filled} 个空值 -> {out_path}")
    return 0


def _fill_csv(input_path, args, col_specs):
    with open(input_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        data = list(reader)

    if not data:
        print("错误：空文件")
        return 1

    headers = data[0]
    n_cols = len(headers)
    rows = []
    for r in data[1:]:
        row = list(r)
        while len(row) < n_cols:
            row.append("")
        rows.append(row)

    try:
        if col_specs:
            fill_indices = _resolve_col_indices(headers, col_specs)
        else:
            fill_indices = list(range(len(headers)))
    except ValueError as e:
        print(str(e))
        return 1

    group_by_idx = None
    if args.group_by:
        gb = args.group_by
        if gb not in headers:
            avail = ", ".join(str(h) for h in headers)
            print(f"错误：找不到分组列「{gb}」，可用列：{avail}")
            avail = ", ".join(str(h) for h in headers)
            print(f"错误：找不到分组列「{gb}」，可用列：{avail}")
        group_by_idx = headers.index(gb)
        if group_by_idx not in fill_indices:
            fill_indices.append(group_by_idx)

    min_gap = args.min_gap

    filled, remaining, remaining_per_col = _do_fill_rows(
        headers, rows, fill_indices, group_by_idx, min_gap)

    if args.report:
        report_lines = []
        report_lines.append("📋 CSV 填充报告:")
        for ci in sorted(fill_indices):
            col_name = headers[ci]
            col_remaining = remaining_per_col[ci]
            report_lines.append(f"  列「{col_name}」: 剩余空值 {col_remaining}")
        report_text = "\n".join(report_lines)
        print(report_text)
        if args.output:
            out_path = Path(args.output)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(report_text, encoding="utf-8")
        return 0

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    print(f"✅ 填充完成：共填充 {filled} 个空值 -> {out_path}")
    return 0



# ---------- sort 智能排序 ----------
_NATURAL_RE = re.compile(r"(\d+)")


def _natural_key(s):
    parts = _NATURAL_RE.split(str(s))
    result = []
    for i, p in enumerate(parts):
        if i % 2 == 1:
            result.append((0, int(p), ""))
        else:
            result.append((1, 0, p.lower()))
    return result


def _parse_date_value(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_sort_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s:
        return None
    if s.endswith("%"):
        s = s[:-1]
    s = s.replace(",", "").replace("￥", "").replace("¥", "").replace("$", "").replace(" ", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _detect_column_type(rows, col_idx):
    num_count = 0
    date_count = 0
    total = 0
    for r in rows:
        v = r[col_idx] if col_idx < len(r) else None
        if _is_empty(v):
            continue
        total += 1
        if _parse_sort_number(v) is not None:
            num_count += 1
        elif _parse_date_value(v) is not None:
            date_count += 1
    if total == 0:
        return "text"
    if num_count / total >= 0.7:
        return "number"
    if date_count / total >= 0.7:
        return "date"
    return "text"


def _sort_key_for_value(v, col_type):
    if _is_empty(v):
        return None
    if col_type == "number":
        n = _parse_sort_number(v)
        if n is not None:
            return ("num", n)
        return ("text_nat", _natural_key(str(v)))
    if col_type == "date":
        d = _parse_date_value(v)
        if d is not None:
            return ("date", d)
        return ("text_nat", _natural_key(str(v)))
    return ("text_nat", _natural_key(str(v)))


def cmd_sort(args):
    """智能排序：数值感知 + 自然排序 + 空值沉底 + 多关键字"""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在：{input_path}")
        return 1

    sheet_ref = getattr(args, "sheet", None)
    if sheet_ref and input_path.suffix.lower() in (".xlsx", ".xlsm"):
        headers, rows = _read_sheet(input_path, sheet_ref)
    else:
        headers, rows = read_table(input_path)

    if not headers:
        print(f"错误：{args.input} 读不到表头")
        return 1
    headers = [str(h) for h in headers]

    by_specs = []
    for part in args.by.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            col_name, direction = part.rsplit(":", 1)
            col_name = col_name.strip()
            direction = direction.strip().lower()
            if direction not in ("asc", "desc"):
                print(f"错误：排序方向只支持 asc/desc，当前：{direction}")
                return 1
        else:
            col_name = part
            direction = "asc"
        if col_name not in headers:
            print(f"错误：表里没有列「{col_name}」，现有列：{headers}")
            return 1
        by_specs.append((headers.index(col_name), direction))

    if not by_specs:
        print("错误：--by 不能为空")
        return 1

    limit = getattr(args, "limit", None)
    if limit is not None and limit <= 0:
        print(f"错误：--limit 必须为正整数，当前：{limit}")
        return 1

    empty_last = not getattr(args, "empty_first", False)

    col_types = {}
    for col_idx, _ in by_specs:
        col_types[col_idx] = _detect_column_type(rows, col_idx)

    def make_key(row):
        keys = []
        for col_idx, direction in by_specs:
            v = row[col_idx] if col_idx < len(row) else None
            is_emp = _is_empty(v)
            sv = _sort_key_for_value(v, col_types[col_idx])
            if is_emp:
                emp_sort = (1,) if empty_last else (0,)
                keys.append(emp_sort)
                continue
            emp_sort = (0,) if empty_last else (1,)
            if direction == "desc":
                if sv and sv[0] == "num":
                    keys.append(emp_sort + (0, -sv[1]))
                elif sv and sv[0] == "date":
                    max_date = datetime.date(9999, 12, 31)
                    keys.append(emp_sort + (0, max_date - sv[1]))
                elif sv and sv[0] == "text_nat":
                    keys.append(emp_sort + (1,))
                else:
                    keys.append(emp_sort + (2,))
            else:
                if sv and sv[0] == "num":
                    keys.append(emp_sort + (0, sv[1]))
                elif sv and sv[0] == "date":
                    keys.append(emp_sort + (0, sv[1]))
                elif sv and sv[0] == "text_nat":
                    keys.append(emp_sort + (1,) + tuple(sv[1]))
                else:
                    keys.append(emp_sort + (2,))
        return keys

    indexed_rows_with_idx = [(r, i) for i, r in enumerate(rows)]

    def combined_key(ri):
        row, orig_idx = ri
        mk = make_key(row)
        return (mk, orig_idx)

    indexed_rows_with_idx.sort(key=combined_key)
    sorted_rows = [r for r, _ in indexed_rows_with_idx]

    if limit is not None:
        sorted_rows = sorted_rows[:limit]

    write_table(Path(args.output), headers, sorted_rows)
    print(f"✅ 排序完成：{len(rows)} 行 → {len(sorted_rows)} 行（按 {args.by}）→ {args.output}")
    return 0



# ---------- transpose 行列互换 ----------
def _read_raw_rows(path, sheet_name=None):
    """读取所有行，不把首行当表头"""
    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
        data = [[("" if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    elif ext == ".csv":
        with open(p, newline="", encoding="utf-8-sig") as f:
            data = [row for row in csv.reader(f)]
    elif ext == ".txt":
        with open(p, encoding="utf-8-sig") as f:
            sample = f.read(4096)
        delim = "	" if "	" in sample else ("," if "," in sample else None)
        with open(p, newline="", encoding="utf-8-sig") as f:
            data = [row for row in csv.reader(f, delimiter=delim)] if delim else [
                [line.rstrip("\n")] for line in f]
    else:
        raise ValueError(f"不支持的文件类型: {p.name}（支持 .xlsx/.csv/.txt）")
    return data


def _col_letter_to_idx(ref, max_cols):
    """列字母(A/B/.../Z/AA/...)转 0-based 索引，不合法返回 None"""
    ref = ref.strip().upper()
    if not ref or not ref.isalpha():
        return None
    idx = 0
    for ch in ref:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    idx -= 1
    return idx if 0 <= idx < max_cols else None


def _dedup_headers(headers):
    """重复表头自动加 _2/_3 后缀"""
    seen = {}
    result = []
    for h in headers:
        h = str(h)
        if h in seen:
            seen[h] += 1
            result.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 1
            result.append(h)
    return result


def cmd_transpose(args):
    """行列互换/转置表"""
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在：{input_path}")
        return 1

    sheet_ref = getattr(args, "sheet", None)
    no_header = getattr(args, "no_header", False)
    index_col = getattr(args, "index_col", None)

    if no_header:
        try:
            all_data = _read_raw_rows(input_path, sheet_ref)
        except Exception as e:
            print(f"错误：读取文件失败 - {e}")
            return 1
        if not all_data:
            if args.output:
                write_table(Path(args.output), [""], [])
                print(f"✅ 转置完成：空表 → {args.output}")
            else:
                print("空表（无数据）")
            return 0
        max_cols = max(len(r) for r in all_data)
        all_data = [list(r) + [""] * (max_cols - len(r)) for r in all_data]
        transposed = [list(r) for r in zip(*all_data)]
        new_headers = [str(v) for v in transposed[0]]
        new_rows = [[str(v) for v in r] for r in transposed[1:]] if len(transposed) > 1 else []
        orig_rows_count = len(all_data)
        orig_cols_count = max_cols
    else:
        try:
            if sheet_ref and input_path.suffix.lower() in (".xlsx", ".xlsm"):
                headers, rows = _read_sheet(input_path, sheet_ref)
            else:
                headers, rows = read_table(input_path)
        except Exception as e:
            print(f"错误：读取文件失败 - {e}")
            return 1

        if not headers and not rows:
            if args.output:
                write_table(Path(args.output), [""], [])
                print(f"✅ 转置完成：空表 → {args.output}")
            else:
                print("空表（无数据）")
            return 0

        headers = [str(h) for h in headers]
        orig_rows_count = len(rows)
        orig_cols_count = len(headers)
        full = [headers]
        for r in rows:
            padded = list(r) + [""] * (len(headers) - len(r))
            full.append([str(c) for c in padded[:len(headers)]])
        transposed = [list(r) for r in zip(*full)]
        new_headers = [str(v) for v in transposed[0]]
        new_rows = [[str(v) for v in r] for r in transposed[1:]] if len(transposed) > 1 else []

    if index_col:
        col_idx = None
        if index_col in new_headers:
            col_idx = new_headers.index(index_col)
        else:
            letter_idx = _col_letter_to_idx(index_col, len(new_headers))
            if letter_idx is not None:
                col_idx = letter_idx
        if col_idx is None:
            print(f"错误：转置后表里没有列「{index_col}」，现有列：{new_headers}")
            return 1
        index_values = [row[col_idx] for row in new_rows]
        remaining_col_headers = [h for i, h in enumerate(new_headers) if i != col_idx]
        remaining_col_data = []
        for j in range(len(new_headers)):
            if j == col_idx:
                continue
            col_data = [row[j] for row in new_rows]
            remaining_col_data.append(col_data)
        new_headers = [index_col] + [str(v) for v in index_values]
        new_rows = [[remaining_col_headers[i]] + [str(v) for v in remaining_col_data[i]] for i in range(len(remaining_col_headers))] if remaining_col_headers else []

    new_headers = _dedup_headers(new_headers)

    if args.output:
        write_table(Path(args.output), new_headers, new_rows)
        print(f"✅ 转置完成：{orig_rows_count} 行 × {orig_cols_count} 列 → {len(new_rows)} 行 × {len(new_headers)} 列 → {args.output}")
    else:
        print("	".join(new_headers))
        for r in new_rows[:20]:
            print("	".join(str(v) for v in r))
        total = len(new_rows)
        if total > 20:
            print(f"... 共 {total} 行，已显示前 20 行")
        else:
            print(f"共 {total} 行")
    return 0



# ---------- mask：数据脱敏 ----------
_COMPOUND_SURNAMES = {
    "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫",
    "万俟", "闻人", "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台",
    "皇甫", "宗政", "濮阳", "公冶", "太叔", "申屠", "公孙", "慕容",
    "仲孙", "钟离", "长孙", "宇文", "司徒", "鲜于", "司空", "闾丘",
    "子车", "亓官", "司寇", "巫马", "公西", "颛孙", "壤驷", "公良",
    "漆雕", "乐正", "宰父", "谷梁", "拓跋", "夹谷", "轩辕", "令狐",
    "段干", "百里", "呼延", "东郭", "南门", "羊舌", "微生", "公户",
    "公玉", "公仪", "梁丘", "公仲", "公上", "公门", "公山", "公坚",
    "左丘", "公伯", "西门", "公祖", "第五", "公乘", "贯丘", "公皙",
    "南荣", "东里", "东宫", "仲长", "子书", "子桑", "即墨", "达奚",
    "褚师",
}
_ADDRESS_PREFIXES = (
    "北京", "天津", "上海", "重庆",
    "河北", "山西", "辽宁", "吉林", "黑龙江",
    "江苏", "浙江", "安徽", "福建", "江西", "山东",
    "河南", "湖北", "湖南", "广东", "海南",
    "四川", "贵州", "云南", "陕西", "甘肃",
    "青海", "台湾",
    "内蒙古", "广西", "西藏", "宁夏", "新疆",
    "香港", "澳门",
)


def _mask_phone(v):
    s = str(v).strip()
    if not s:
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        s = str(int(v)) if v == int(v) else str(v)
    if re.fullmatch(r"1[3-9]\d{9}", s):
        return s[:3] + "****" + s[7:]
    return v


def _mask_idcard(v):
    s = str(v).strip()
    if not s:
        return v
    if re.fullmatch(r"\d{17}[\dXx]", s):
        return s[:6] + "*" * 8 + s[-4:]
    if re.fullmatch(r"\d{15}", s):
        return s[:6] + "*" * 3 + s[-3:]
    return v


def _mask_name(v):
    s = str(v).strip()
    if not s:
        return v
    if not any("\u4e00" <= c <= "\u9fff" for c in s):
        return v
    if len(s) >= 3 and s[:2] in _COMPOUND_SURNAMES:
        return s[:2] + "*" * (len(s) - 2)
    if len(s) >= 2 and "\u4e00" <= s[0] <= "\u9fff":
        return s[0] + "*" * (len(s) - 1)
    if len(s) == 1 and "\u4e00" <= s[0] <= "\u9fff":
        return s + "*"
    return v


def _mask_email(v):
    s = str(v).strip()
    if not s or "@" not in s:
        return v
    local, domain = s.rsplit("@", 1)
    if not local:
        return v
    return local[0] + "***@" + domain


def _mask_bank(v):
    s = re.sub(r"[\s-]", "", str(v).strip())
    if not s:
        return v
    if len(s) < 9 or not s.isdigit():
        return v
    return s[:4] + " **** **** " + s[-4:]


def _mask_address(v):
    s = str(v).strip()
    if not s or len(s) < 6:
        return v
    prefix_len = 0
    for pfx in _ADDRESS_PREFIXES:
        if s.startswith(pfx):
            prefix_len = len(pfx)
            # include 市/省/自治区 suffix
            rest = s[prefix_len:]
            if rest.startswith("市") or rest.startswith("省"):
                prefix_len += 1
            elif rest.startswith("自治区"):
                prefix_len += 3
            elif rest.startswith("特别行政区"):
                prefix_len += 5
            break
    if prefix_len == 0:
        prefix_len = min(3, len(s))
    detail = s[prefix_len:]
    if not detail:
        return v
    return s[:prefix_len] + "*" * len(detail)


def _mask_text(v):
    s = str(v)
    if len(s) <= 2:
        return v
    return s[0] + "*" * (len(s) - 2) + s[-1]


def _mask_hash(v, seed=None):
    s = str(v)
    raw = f"{seed}:{s}" if seed is not None else s
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12]


def cmd_mask(args):
    """数据脱敏（手机号/身份证/姓名/邮箱/银行卡/地址/通用/哈希/删除列）"""
    VALID_MODES = {"phone", "idcard", "name", "email", "bank", "address", "text", "hash", "drop"}
    inp = Path(args.input)
    if not inp.exists():
        print(f"错误：文件不存在：{inp}")
        return 1
    mode = args.mode
    if mode not in VALID_MODES:
        print(f"错误：不支持的脱敏模式「{mode}」，可选：{', '.join(sorted(VALID_MODES))}")
        return 1

    sheet_name = getattr(args, "sheet", None)
    if sheet_name and inp.suffix.lower() in (".xlsx", ".xlsm"):
        wb_check = openpyxl.load_workbook(inp, read_only=True)
        if sheet_name not in wb_check.sheetnames:
            available = list(wb_check.sheetnames)
            wb_check.close()
            print(f"错误：sheet「{sheet_name}」不存在，可选：{', '.join(available)}")
            return 1
        wb_check.close()

    headers, rows = _read_sheet(str(inp), sheet_name)
    if not headers:
        print("错误：表格为空或无法读取")
        return 1

    col_specs = [c.strip() for c in args.cols.split(",") if c.strip()]
    try:
        col_indices = _resolve_col_indices(headers, col_specs)
    except ValueError as e:
        print(str(e))
        return 1

    seed = getattr(args, "seed", None)

    if mode == "drop":
        new_headers = [h for i, h in enumerate(headers) if i not in col_indices]
        new_rows = [[v for i, v in enumerate(r) if i not in col_indices] for r in rows]
    else:
        mask_fn = {
            "phone": lambda v: _mask_phone(v),
            "idcard": lambda v: _mask_idcard(v),
            "name": lambda v: _mask_name(v),
            "email": lambda v: _mask_email(v),
            "bank": lambda v: _mask_bank(v),
            "address": lambda v: _mask_address(v),
            "text": lambda v: _mask_text(v),
            "hash": lambda v: _mask_hash(v, seed),
        }[mode]
        new_headers = list(headers)
        new_rows = [list(r) for r in rows]
        for r in new_rows:
            for ci in col_indices:
                if ci < len(r):
                    v = r[ci]
                    if v is None or (isinstance(v, str) and not v.strip()):
                        continue
                    r[ci] = mask_fn(v)

    if getattr(args, "dry_run", False):
        print(f"脱敏预览（模式：{mode}，列：{', '.join(col_specs)}）：")
        print()
        for ci in col_indices:
            if ci >= len(headers):
                continue
            print(f"  列「{headers[ci]}」：")
            preview_count = 0
            for r in rows:
                if preview_count >= 3:
                    break
                if ci < len(r):
                    orig = r[ci]
                    if orig is None or (isinstance(orig, str) and not orig.strip()):
                        continue
                    if mode == "drop":
                        masked = "(已删除)"
                    else:
                        mask_fn_preview = {
                            "phone": _mask_phone,
                            "idcard": _mask_idcard,
                            "name": _mask_name,
                            "email": _mask_email,
                            "bank": _mask_bank,
                            "address": _mask_address,
                            "text": _mask_text,
                            "hash": lambda v, s=seed: _mask_hash(v, s),
                        }[mode]
                        masked = mask_fn_preview(orig)
                    print(f"    {orig} → {masked}")
                    preview_count += 1
        return 0

    if args.output:
        out_path = Path(args.output)
    else:
        out_path = inp.parent / f"{inp.stem}_masked{inp.suffix}"
    write_table(out_path, new_headers, new_rows)
    print(f"已脱敏并保存至：{out_path}")
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Excel/CSV 批量处理工具箱（合并/拆分/去重/转换/关联/统计/对比/出图/模板批量生成/空值填充/智能排序/转置/脱敏）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  合并目录下所有表格：   python excel_tool.py merge ./表格 -o 合并结果.xlsx --source-col
  按「部门」列拆分：     python excel_tool.py split 员工表.xlsx --column 部门 -d ./拆分结果
  按「手机号」列去重：   python excel_tool.py dedup 客户表.csv --columns 手机号 -o 去重结果.csv
  xlsx 转 csv：          python excel_tool.py convert ./表格 -d ./csv输出 --fmt .csv
  快速摸底统计：         python excel_tool.py summary ./表格
  两表VLOOKUP式合并：     python excel_tool.py join 订单表.xlsx 客户表.xlsx --on 客户编号 -o 结果.xlsx
  分组汇总(透视表平替)：  python excel_tool.py groupby 订单表.xlsx --by 产品 --agg 销售额:sum,订单号:count -o 汇总.xlsx
  按条件筛选行：        python excel_tool.py filter 订单表.xlsx --where "销售额>1000" --where "地区~华东" -o 大单.xlsx
  数据清洗：            python excel_tool.py clean 名单.xlsx --dates --fill "-" -o 名单_清洗.xlsx
  数据体检报告：         python excel_tool.py report 客户表.xlsx -o 体检报告.md
  二维交叉透视表：        python excel_tool.py pivot 订单表.xlsx --rows 地区 --cols 月份 --values 金额:sum --margins -o 透视表.xlsx
  监控文件夹自动汇总：     python excel_tool.py watch ./收集箱 --out 总表.xlsx --interval 5
  两表关联合并（VLOOKUP 平替）：python excel_tool.py merge 员工表.xlsx 工资表.xlsx --on 工号 --how left -o 合并结果.xlsx
  两表对比（找新增/删除/修改）：python excel_tool.py diff 旧表.xlsx 新表.xlsx --key 编号 -o 对比结果.xlsx
  用 SQL 查 Excel：           python excel_tool.py sql "SELECT 部门, SUM(金额) FROM data GROUP BY 部门" -i 数据.xlsx -o 结果.xlsx
  从数据出图：              python excel_tool.py chart -i 数据.xlsx --type bar --x 部门 --y 金额 --title "各部门销售额"
  空值向下填充：              python excel_tool.py fill 数据.xlsx -o 填充后.xlsx
  拆分合并单元格并分组填充：    python excel_tool.py fill 数据.xlsx --unmerge --group-by 部门 -o 填充后.xlsx
  智能排序（金额降序取前10）：    python excel_tool.py sort 订单表.xlsx --by 金额:desc --limit 10 -o top10.xlsx
  转置表格（行列互换）：            python excel_tool.py transpose 数据.xlsx -o 转置结果.xlsx
  数据脱敏（手机号/身份证等）：     python excel_tool.py mask 客户表.xlsx --cols 手机号,身份证 --mode phone -o 脱敏.xlsx
""")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("merge", help="合并多个表格 / 两表按关键列关联（--on 模式）")
    p.add_argument("inputs", nargs="+", help="文件或目录（关联模式：左表 右表）")
    p.add_argument("-o", "--output", default="合并结果.xlsx")
    p.add_argument("--source-col", action="store_true", help="追加「来源文件」列")
    p.add_argument("-r", "--recursive", action="store_true", help="递归扫描子目录")
    p.add_argument("--on", help="关联关键列（逗号分隔多列），启用后为 VLOOKUP 模式")
    p.add_argument("--how", default="left", choices=["left", "inner", "right"], help="关联方式（默认 left）")
    p.add_argument("--suffix", default="_右", help="重名列后缀（默认 _右）")
    p.add_argument("--dup", default="first", choices=["first", "error", "concat"], help="右表 key 重复策略（默认 first）")
    p.add_argument("--left-sheet", help="左表 xlsx 读取指定 sheet 名")
    p.add_argument("--right-sheet", help="右表 xlsx 读取指定 sheet 名")
    p.set_defaults(fn=cmd_merge)

    p = sub.add_parser("split", help="按列值拆分表格")
    p.add_argument("input", help="输入表格")
    p.add_argument("--column", required=True, help="拆分依据的列名")
    p.add_argument("-d", "--output-dir", default="./拆分结果")
    p.add_argument("--fmt", default=".xlsx", choices=[".xlsx", ".csv"])
    p.set_defaults(fn=cmd_split)

    p = sub.add_parser("dedup", help="去重")
    p.add_argument("input", help="输入表格")
    p.add_argument("--columns", help="去重依据列（逗号分隔，默认整行）")
    p.add_argument("-o", "--output", default="去重结果.xlsx")
    p.set_defaults(fn=cmd_dedup)

    p = sub.add_parser("convert", help="格式互转 xlsx↔csv")
    p.add_argument("inputs", nargs="+")
    p.add_argument("-d", "--output-dir", default="./转换结果")
    p.add_argument("--fmt", default=".csv", choices=[".xlsx", ".csv"])
    p.add_argument("-r", "--recursive", action="store_true")
    p.set_defaults(fn=cmd_convert)

    p = sub.add_parser("summary", help="快速摸底统计")
    p.add_argument("inputs", nargs="+")
    p.add_argument("-r", "--recursive", action="store_true")
    p.set_defaults(fn=cmd_summary)

    p = sub.add_parser("join", help="两表按关键列 VLOOKUP 式合并")
    p.add_argument("left", help="左表（主表，每行保留）")
    p.add_argument("right", help="右表（查找表）")
    p.add_argument("--on", required=True, help="关键列名（两表都要有）")
    p.add_argument("--columns", help="从右表取哪些列（逗号分隔，默认除关键列外全部）")
    p.add_argument("--prefix", help="新列名前缀（防重名，如 '员工_'）")
    p.add_argument("--only-matched", action="store_true", help="只保留匹配上的行（默认未匹配留空）")
    p.add_argument("--multi", action="store_true", help="右表同键多行时全部保留（默认取第一条）")
    p.add_argument("-o", "--output", default="join结果.xlsx")
    p.set_defaults(fn=cmd_join)

    p = sub.add_parser("groupby", help="分组汇总（透视表平替）")
    p.add_argument("input", help="输入表格")
    p.add_argument("--by", required=True, help="分组依据的列名")
    p.add_argument("--agg", help="聚合定义：列名:函数（sum/count/avg/max/min），逗号分隔；缺省对所有数值列 sum")
    p.add_argument("--sort", help="结果按这个聚合列排序（写列名即可）")
    p.add_argument("--asc", action="store_true", help="升序（默认降序）")
    p.add_argument("-o", "--output", default="分组汇总.xlsx")
    p.set_defaults(fn=cmd_groupby)

    p = sub.add_parser("filter", help="按条件筛选行")
    p.add_argument("input", help="输入表格")
    p.add_argument("--where", action="append", required=True,
                   help="条件：列名=值 / 列名!=值 / 列名>100 / 列名~包含字；可叠加多个（且的关系）")
    p.add_argument("--drop-empty", action="store_true", help="顺便去掉全空行")
    p.add_argument("-o", "--output", default="筛选结果.xlsx")
    p.set_defaults(fn=cmd_filter)

    p = sub.add_parser("clean", help="数据清洗（去空格/去空行/统一日期/填缺失）")
    p.add_argument("input", help="输入表格")
    p.add_argument("--columns", help="只清洗这几列（逗号分隔，默认全部）")
    p.add_argument("--dates", action="store_true", help="统一日期格式为 YYYY-MM-DD")
    p.add_argument("--fill", default=None, help="空单元格填这个值（默认不填）")
    p.add_argument("-o", "--output", default="清洗结果.xlsx")
    p.set_defaults(fn=cmd_clean)

    p = sub.add_parser("report", help="数据体检报告（缺失/重复/类型/异常，出 Markdown）")
    p.add_argument("input", help="输入表格")
    p.add_argument("-o", "--output", help="报告输出到 .md 文件（默认打印到终端）")
    p.set_defaults(fn=cmd_report)

    p = sub.add_parser("pivot", help="二维交叉透视表（Excel数据透视表平替）")
    p.add_argument("input", help="输入表格")
    p.add_argument("--rows", required=True, help="行字段（多字段用逗号分隔）")
    p.add_argument("--cols", required=True, help="列字段（该列每个值变成一列）")
    p.add_argument("--values", required=True, help="值列:函数（sum/count/avg/max/min，默认sum）")
    p.add_argument("--margins", action="store_true", help="加行/列总计")
    p.add_argument("--keep-order", action="store_true", help="保持原始出现顺序（默认按合计降序）")
    p.add_argument("--sort-rows", action="store_true", help="行按键字典序排")
    p.add_argument("-o", "--output", default="透视表.xlsx")
    p.set_defaults(fn=cmd_pivot)

    p = sub.add_parser("unpivot", help="宽表转长表（逆透视，pivot 的逆操作）")
    p.add_argument("input", help="输入表格")
    p.add_argument("--id", default="", help="保持不动的 ID 列（多列逗号分隔，如：门店,区域）")
    p.add_argument("--cols", default="", help="要拆开的指标列（逗号分隔；缺省=除 ID 外全部列）")
    p.add_argument("--name", default="指标", help="新列：指标名的列名（默认「指标」）")
    p.add_argument("--value", default="值", help="新列：值的列名（默认「值」）")
    p.add_argument("--keep-empty", action="store_true", help="保留空值单元格（默认跳过）")
    p.add_argument("-o", "--output", default="长表.xlsx")
    p.set_defaults(fn=cmd_unpivot)

    p = sub.add_parser("watch", help="监控文件夹，新表格自动汇入总表（收集箱场景）")
    p.add_argument("folder", help="被监控的文件夹")
    p.add_argument("--out", "--output", dest="output", default="汇总结果.xlsx", help="汇总表输出路径")
    p.add_argument("--interval", type=float, default=5.0, help="扫描间隔秒数，默认5")
    p.add_argument("-r", "--recursive", action="store_true", help="递归监控子目录")
    p.add_argument("--once", action="store_true", help="只扫描汇总一次就退出（配合定时任务用）")
    p.set_defaults(fn=cmd_watch)

    p = sub.add_parser("diff", help="两表对比：按关键列找出新增/删除/修改")
    p.add_argument("old", help="旧表（基准）")
    p.add_argument("new", help="新表（对比目标）")
    p.add_argument("--key", required=True, help="关键列名（两表都要有，按此列对齐行）")
    p.add_argument("--columns", help="对比哪些列（逗号分隔，缺省=两表共同列排除 key）")
    p.add_argument("-o", "--output", help="输出 xlsx（4 个 sheet：摘要/新增/删除/修改）")
    p.set_defaults(fn=cmd_diff)


    p = sub.add_parser("sql", help="用 SQL 查询 Excel/CSV（sqlite3 零依赖）")
    p.add_argument("query", help="SQL 查询语句（表名为 data，如：SELECT 部门, SUM(金额) FROM data GROUP BY 部门）")
    p.add_argument("-i", "--input", required=True, help="输入文件（.xlsx/.csv）")
    p.add_argument("-o", "--output", help="输出文件（.xlsx/.csv），不指定则打印到终端")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名（默认第一个）")
    p.add_argument("--print", dest="print_rows", type=int, help="预览前 N 行（打印到终端）")
    p.set_defaults(fn=cmd_sql)


    p = sub.add_parser("template", help="数据驱动批量文档生成（邮件合并平替）")
    p.add_argument("data", help="数据表（.xlsx/.csv）")
    p.add_argument("--tpl", required=True, help="模板文件（.docx/.xlsx）")
    p.add_argument("--out", required=True, help="输出目录")
    p.add_argument("--name-col", help="用作文件名的列")
    p.add_argument("--filename-tpl", help="文件名模板，支持 {列名} 占位")
    p.add_argument("--sheet", help="xlsx 数据表指定 sheet 名")
    p.add_argument("--strict", action="store_true", help="缺失列时报错退出")
    p.add_argument("--dry-run", action="store_true", help="预览前 3 行替换结果，不写盘")
    p.set_defaults(fn=cmd_template)

    p = sub.add_parser("fill", help="空值向下填充（合并单元格/前向填充 ffill）")
    p.add_argument("input", help="输入表格（.xlsx/.csv）")
    p.add_argument("-o", "--output", help="输出文件（--report 时可选）")
    p.add_argument("--cols", help="只填充指定列（逗号分隔，支持列名或 A/B/C 字母）")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名或索引（默认全部 sheet）")
    p.add_argument("--unmerge", action="store_true", help="先拆分合并单元格再填充")
    p.add_argument("--group-by", help="分组内填充（换组时重置，防止跨组污染）")
    p.add_argument("--min-gap", type=int, default=None, help="只填充连续空值长度<=N 的段")
    p.add_argument("--report", action="store_true", help="只报告不写文件")
    p.set_defaults(fn=cmd_fill)


    p = sub.add_parser("sort", help="智能排序（数值感知+自然排序+空值沉底+多关键字）")
    p.add_argument("input", help="输入表格（.xlsx/.csv）")
    p.add_argument("--by", required=True, help="排序键：列名 或 列名:asc/desc，多键逗号分隔（如 部门:asc,金额:desc）")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名或索引（默认第一个）")
    p.add_argument("--limit", type=int, help="只输出排序后前 N 行（Top-N 场景）")
    p.add_argument("--empty-first", action="store_true", help="空值排在最前（默认空值沉底/排最后）")
    p.add_argument("--stable", action="store_true", help="稳定排序（默认即稳定，此 flag 仅作显式声明）")
    p.add_argument("-o", "--output", default="排序结果.xlsx", help="输出文件")
    p.set_defaults(fn=cmd_sort)

    p = sub.add_parser("transpose", help="转置表格（行列互换）")
    p.add_argument("input", help="输入表格（.xlsx/.csv）")
    p.add_argument("-o", "--output", default=None, help="输出文件（缺省打印预览到终端）")
    p.add_argument("--no-header", action="store_true", help="输入无表头，纯矩阵转置")
    p.add_argument("--index-col", help="转置后把该列值作为新表头（列名或字母如 A/B）")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名")
    p.set_defaults(fn=cmd_transpose)

    p = sub.add_parser("chart", help="从 Excel/CSV 数据生成图表（需 pip install biaoge[chart]）")
    p.add_argument("-i", "--input", required=True, help="输入文件（.xlsx/.csv）")
    p.add_argument("--type", required=True, choices=["bar", "line", "pie", "scatter"], help="图表类型")
    p.add_argument("--x", help="x 轴列名（与 --y 配合使用）")
    p.add_argument("--y", help="y 轴列名（逗号分隔多个）")
    p.add_argument("--sql", help="SQL 查询语句出图（第一列做 x，其余数值列做 y）")
    p.add_argument("--title", help="图表标题")
    p.add_argument("--out", default="chart.png", help="输出 PNG 路径（默认 chart.png）")
    p.add_argument("--theme", help="matplotlib 主题名")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名（默认第一个）")
    p.set_defaults(fn=cmd_chart)


    p = sub.add_parser("mask", help="数据脱敏（手机号/身份证/姓名/邮箱/银行卡/地址/通用/哈希/删列）")
    p.add_argument("input", help="输入表格（.xlsx/.csv）")
    p.add_argument("--cols", required=True, help="脱敏列（逗号分隔，支持列名或 A/B/C 字母）")
    p.add_argument("--mode", required=True,
                   choices=["phone", "idcard", "name", "email", "bank", "address", "text", "hash", "drop"],
                   help="脱敏模式")
    p.add_argument("--sheet", help="xlsx 指定 sheet 名（默认第一个）")
    p.add_argument("--seed", type=int, default=None, help="hash 模式的种子值（同 seed 同值同输出）")
    p.add_argument("--dry-run", action="store_true", help="预览脱敏效果（前 3 行对照表），不写盘")
    p.add_argument("-o", "--output", default=None, help="输出文件（缺省：<名>_masked.xlsx）")
    p.set_defaults(fn=cmd_mask)

    args = ap.parse_args()
    sys.exit(args.fn(args))


if __name__ == "__main__":
    main()
