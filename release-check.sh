#!/usr/bin/env bash
set -euo pipefail

FAIL=0
PASS_COUNT=0
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

if [ -f .venv/bin/activate ]; then
  source .venv/bin/activate
fi

pass() { echo "  OK $1"; PASS_COUNT=$((PASS_COUNT+1)); }
fail() { echo "  FAIL $1"; FAIL=1; }

echo "biaoge release check"
echo "===================="

echo ""
echo "[1/4] pytest..."
if python -m pytest tests/ -q 2>&1 | tail -3; then
  pass "pytest"
else
  fail "pytest"
fi

echo ""
echo "[2/4] version..."
PV=$(grep -E "^version\s*=" pyproject.toml | head -1 | cut -d'"' -f2)
DW=$(ls dist/biaoge-*.whl 2>/dev/null | sort -V | tail -1 | grep -oP "biaoge-\K[^-]+" || echo "")
echo "  pyproject: $PV"
[ -n "$DW" ] && echo "  wheel: $DW"
if [ -n "$DW" ] && [ "$PV" != "$DW" ]; then
  fail "version mismatch"
else
  pass "version $PV"
fi

echo ""
echo "[3/4] smoke..."
SD=$(mktemp -d)
python smoke_data.py "$SD"

SO=1
rs() { local name="$1"; shift; if "$@" >/dev/null 2>&1; then echo "  OK $name"; else echo "  FAIL $name"; SO=0; fi; }
rs sql python src/biaoge.py sql "SELECT 地区,金额 FROM data LIMIT 3" -i "$SD/data.csv" -o "$SD/o.xlsx"
rs sort python src/biaoge.py sort "$SD/data.csv" --by 金额:desc --limit 3 -o "$SD/s.csv"
rs transpose python src/biaoge.py transpose "$SD/data.csv" -o "$SD/t.csv"
rs groupby python src/biaoge.py groupby "$SD/data.csv" --by 地区 --agg 金额:sum -o "$SD/g.xlsx"
if [ "$SO" = "1" ]; then pass smoke; else fail smoke; fi
rm -rf "$SD"

echo ""
echo "[4/4] wheel..."
WH=$(ls dist/biaoge-*.whl 2>/dev/null | sort -V | tail -1)
if [ -z "$WH" ]; then
  fail "no wheel"
else
  VD=$(mktemp -d)
  if python -m venv "$VD" && "$VD/bin/pip" install -q "$WH" && "$VD/bin/biaoge" --help >/dev/null 2>&1; then
    pass "wheel"
  else
    fail "wheel"
  fi
  rm -rf "$VD"
fi

echo ""
echo "===================="
echo "pass: $PASS_COUNT"
if [ "$FAIL" = "0" ]; then
  echo "RELEASE-READY"
  exit 0
else
  echo "HAS FAILURES"
  exit 1
fi
