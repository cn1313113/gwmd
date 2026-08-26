#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""前后端一致性测试（解析 + 结构校验，可多样例）

用法:
    python3 tests/consistency.py [样例文件...]
    # 默认: tests/cases.gwmd tests/cases_warnings.gwmd
    # 退出码 0=全部一致, 1=存在不一致, 2=执行错误

说明:
    - 修改 gwmd2docx.py 或 app.js 的解析/校验逻辑后必须运行本测试
    - 新增语法特性时，先在 tests/cases.gwmd 中补样例
    - 结构校验样例放 tests/cases_warnings.gwmd
"""
import difflib
import json
import os
import random
import subprocess
import sys
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_CASES = [os.path.join(HERE, 'cases.gwmd'),
                 os.path.join(HERE, 'cases_warnings.gwmd')]
CASES = [os.path.abspath(c) for c in (sys.argv[1:] or DEFAULT_CASES)]
JS_PATH = os.environ.get('GWMD_JS', os.path.join(ROOT, 'app.js'))

os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
sys.path.insert(0, ROOT)
from gwmd2docx import parse_gwmd, validate_gwmd  # noqa: E402


def node_run(script):
    proc = subprocess.run(['node', '-e', script],
                          capture_output=True, text=True, timeout=30)
    if proc.returncode != 0:
        print('!! Node 执行失败:', proc.stderr.strip())
        sys.exit(2)
    return json.loads(proc.stdout)


ok = True

# ---------- 解析一致性（每个样例） ----------
for CASE in CASES:
    with open(CASE, 'r', encoding='utf-8') as f:
        src = f.read()
    blocks, footnotes = parse_gwmd(src)
    py_out = {'blocks': blocks, 'footnotes': dict(footnotes)}
    py_json = json.dumps(py_out, ensure_ascii=False, indent=2, sort_keys=True)
    js = node_run(
        'const fs = require("fs");'
        f'const src = fs.readFileSync({json.dumps(CASE)}, "utf8");'
        f'const GWMD = require({json.dumps(JS_PATH)});'
        'const r = GWMD.parse(src);'
        'console.log(JSON.stringify({blocks: r.blocks, footnotes: r.footnotes}));')
    js_json = json.dumps(js, ensure_ascii=False, indent=2, sort_keys=True)
    if py_json == js_json:
        print(f'PASS 解析一致: {os.path.basename(CASE)} '
              f'({len(blocks)} 块, {len(footnotes)} 条脚注)')
    else:
        ok = False
        print(f'FAIL 解析不一致: {os.path.basename(CASE)}')
        diff = list(difflib.unified_diff(
            py_json.splitlines(), js_json.splitlines(),
            'Python', 'Node', lineterm=''))
        print('\n'.join(diff[:60]))

# ---------- 字符宽度判定一致性抽查 ----------
rng = random.Random(20260826)
width_samples = [0x00, 0x20, 0x41, 0x7E, 0xA1, 0x2014, 0x2018, 0x201C, 0x2026,
                 0x20AC, 0x2192, 0x2460, 0x2605, 0x2E80, 0x3000, 0x3001, 0x4E00,
                 0x9FFF, 0xFF01, 0xFF60, 0xFF61, 0xFFE0, 0x1100, 0x3131, 0xAC00,
                 0xFE30, 0xE000, 0x10FFFF]
width_samples += [rng.randrange(0x110000) for _ in range(1500)]
js_widths = node_run(
    f'const GWMD = require({json.dumps(JS_PATH)});'
    'const cps = ' + json.dumps(width_samples) + ';'
    'console.log(JSON.stringify(cps.map(c => GWMD.isWideChar(String.fromCodePoint(c)))));')
mismatch = []
for cp, js_w in zip(width_samples, js_widths):
    py_w = unicodedata.east_asian_width(chr(cp)) in ('F', 'W', 'A')
    if js_w != py_w:
        mismatch.append((cp, py_w, js_w))
if mismatch:
    ok = False
    print('FAIL: 字符宽度判定不一致', len(mismatch), '处，前 10 处:')
    for cp, pw, jw in mismatch[:10]:
        print(f'  U+{cp:04X} {chr(cp)!r}: Python={pw} Node={jw}')
else:
    print(f'PASS 宽度一致: {len(width_samples)} 个码位抽样（F/W/A 规则）')

# ---------- 结构校验一致性（每个样例） ----------
for CASE in CASES:
    with open(CASE, 'r', encoding='utf-8') as f:
        src = f.read()
    py_w = validate_gwmd(src)
    js_w = node_run(
        'const fs = require("fs");'
        f'const src = fs.readFileSync({json.dumps(CASE)}, "utf8");'
        f'const GWMD = require({json.dumps(JS_PATH)});'
        'console.log(JSON.stringify(GWMD.validate(src)));')
    if py_w == js_w:
        print(f'PASS 校验一致: {os.path.basename(CASE)} ({len(py_w)} 条警告)')
    else:
        ok = False
        print(f'FAIL 校验不一致: {os.path.basename(CASE)}')
        print('  Python:', json.dumps(py_w, ensure_ascii=False))
        print('  Node  :', json.dumps(js_w, ensure_ascii=False))

sys.exit(0 if ok else 1)