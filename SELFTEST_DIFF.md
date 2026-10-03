# SELFTEST_DIFF.md — biaoge v0.9.0 diff 命令自测报告

## 1. 单元测试（pytest）

```bash
cd /home/admin/.openclaw/workspace/money/opensource/biaoge-cli
/home/admin/.openclaw/workspace/money/.venv/bin/python -m pytest tests/ -q
```

**结果：44 passed in 19.59s**（原有 37 个 + 新增 7 个 diff 测试，全部通过，无回归）

### 新增 diff 测试清单

| # | 测试函数 | 覆盖场景 |
|---|---------|----------|
| 1 | `test_diff_basic_three_categories` | 基础三类差异：新增/删除/修改 |
| 2 | `test_diff_custom_columns` | `--columns` 只对比指定列 |
| 3 | `test_diff_key_not_found` | key 列不存在 → exit=1 友好报错 |
| 4 | `test_diff_number_normalization` | 数字归一：1 == 1.0 不误报 |
| 5 | `test_diff_key_case_and_space` | key 忽略大小写和首尾空格 |
| 6 | `test_diff_output_xlsx_sheets` | `-o` 输出 xlsx 4 个 sheet 结构验证 |
| 7 | `test_diff_no_common_columns` | 两表列完全不同 → 友好提示 |

## 2. E2E 测试（干净 venv）

```bash
# 建干净 venv
python3 -m venv /tmp/bgtest09/venv
/tmp/bgtest09/venv/bin/pip install dist/biaoge-0.9.0-py3-none-any.whl

# 造两个小表
cd /tmp/bgtest09
# old.csv: E001~E004（4行）
# new.csv: E001(工资改),E003,E004,E005(新增)；E002被删

# 跑 diff
/tmp/bgtest09/venv/bin/biaoge diff old.csv new.csv --key 编号 -o diff_result.xlsx
```

**E2E 输出：**

```
📊 对比结果：旧表 4 行，新表 4 行
   新增：1 行
   删除：1 行
   修改：1 个 key，1 处变化

【新增前 1 条】
   + E005

【删除前 1 条】
   - E002

【修改前 1 处】
   ~ E001  工资: '8000' → '8500'

✅ 对比结果已输出：/tmp/bgtest09/diff_result.xlsx（4 个 sheet：摘要/新增/删除/修改）
```

**xlsx 输出验证（4 个 sheet）：**

- `摘要`：旧表行数=4，新表行数=4，新增=1，删除=1，修改(key数)=1，修改(变化处)=1
- `新增`：E005,孙七,运营,7500
- `删除`：E002,李四,技术,12000
- `修改`：(编号=E001, 列名=工资, 旧值=8000, 新值=8500)

## 3. 构建产物

```
dist/biaoge-0.9.0-py3-none-any.whl   (21K)
dist/biaoge-0.9.0.tar.gz             (26K)
```

路径：`/home/admin/.openclaw/workspace/money/opensource/biaoge-cli/dist/`

## 4. 版本变更

- `pyproject.toml`：version 0.8.0 → 0.9.0，description 追加「两表对比」
- `README.md`：中文/英文双区加 diff 说明 + 路线图打勾 `[x] diff 命令（v0.9）`
- `src/biaoge.py`：新增 `cmd_diff` + `_normalize_key` + `_normalize_value`，注册 `diff` 子命令
