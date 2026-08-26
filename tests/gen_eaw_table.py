#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 East Asian Width 全宽区间表（F/W/A），嵌入 app.js 的 isWideChar。

目的：前后端字符宽度判定一致。
后端 dw() 用 unicodedata.east_asian_width(c) in ('F','W','A') 判定全宽；
本脚本把同样规则固化为 JS 区间表，替换前端手写的近似规则
（旧规则对 …、""、——、'' 等 Ambiguous 字符判定为半宽，与后端不一致）。

用法:
    python3 tests/gen_eaw_table.py          # 自动更新 app.js 中的表
    python3 tests/gen_eaw_table.py --check  # 只校验现有表与 unicodedata 一致

修改规则：更新本脚本后重新生成并运行 tests/consistency.py。
"""
import os
import random
import re
import sys
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
APP_JS = os.path.join(ROOT, 'app.js')


def build_ranges():
    """返回所有 east_asian_width ∈ {F, W, A} 的码位区间 [(start, end), ...]"""
    ranges = []
    start = None
    prev = None
    for cp in range(0x110000):
        wide = unicodedata.east_asian_width(chr(cp)) in ('F', 'W', 'A')
        if wide:
            if start is None:
                start = cp
            prev = cp
        elif start is not None:
            ranges.append((start, prev))
            start = None
    if start is not None:
        ranges.append((start, prev))
    return ranges


def in_ranges(cp, ranges):
    lo, hi = 0, len(ranges) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        a, b = ranges[mid]
        if cp < a:
            hi = mid - 1
        elif cp > b:
            lo = mid + 1
        else:
            return True
    return False


def self_check(ranges):
    """抽样验证区间表与 unicodedata 判定一致（含边界、随机、重点字符）"""
    samples = [0x00, 0x20, 0x41, 0x7E, 0x7F, 0xA0, 0xA1, 0xAD, 0x2014, 0x2018,
               0x201C, 0x2026, 0x20AC, 0x2192, 0x2460, 0x2605, 0x2B1B, 0x2E80,
               0x3000, 0x3001, 0x3040, 0x4E00, 0x9FFF, 0xFF01, 0xFF60, 0xFF61,
               0xFFE0, 0xFFE6, 0x1100, 0x115F, 0x3131, 0xAC00, 0xD7A3, 0xFE30,
               0xFE50, 0xE000, 0xF8FF, 0x10000, 0x1F600, 0x10FFFF]
    rng = random.Random(20260826)
    samples += [rng.randrange(0x110000) for _ in range(3000)]
    for cp in samples:
        expect = unicodedata.east_asian_width(chr(cp)) in ('F', 'W', 'A')
        if in_ranges(cp, ranges) != expect:
            raise AssertionError(
                f'自校验失败 cp=U+{cp:04X} expect={expect}')
    print(f'自校验通过：{len(ranges)} 个区间，{len(samples)} 个抽样码位一致')


def fmt_js(ranges):
    lines = ['  ' + ', '.join(f'[{a},{b}]' for a, b in chunk) + ','
             for chunk in (ranges[i:i + 8] for i in range(0, len(ranges), 8))]
    return '\n'.join(lines)


def extract_current_table(src):
    m = re.search(r'const EAW_WIDE = \[(.*?)\n\];', src, re.S)
    return m


def patch_app_js(ranges):
    with open(APP_JS, 'r', encoding='utf-8') as f:
        src = f.read()
    table = fmt_js(ranges)
    new_block = (
        '// East Asian Width 全宽区间表（F/W/A），由 tests/gen_eaw_table.py 生成\n'
        '// 与后端 gwmd2docx.py 的 unicodedata.east_asian_width 判定保持一致\n'
        '// 修改：重跑 tests/gen_eaw_table.py（勿手改，否则 tests/consistency.py 会失败）\n'
        'const EAW_WIDE = [\n' + table + '\n];\n'
        'function isWideChar(ch) {\n'
        '  const c = ch.codePointAt(0);\n'
        '  let lo = 0, hi = EAW_WIDE.length - 1;\n'
        '  while (lo <= hi) {\n'
        '    const mid = (lo + hi) >> 1;\n'
        '    const r = EAW_WIDE[mid];\n'
        '    if (c < r[0]) hi = mid - 1;\n'
        '    else if (c > r[1]) lo = mid + 1;\n'
        '    else return true;\n'
        '  }\n'
        '  return false;\n'
        '}'
    )
    old_iswide = re.search(
        r'// East Asian Width 全宽区间表.*?}\n}\n', src, re.S)
    if old_iswide:
        src = src[:old_iswide.start()] + new_block + '\n' + src[old_iswide.end():]
    else:
        # 兼容旧实现（无表，仅有近似规则）
        old_fn = re.search(r'function isWideChar\(ch\) \{\n.*?\n\}', src, re.S)
        if not old_fn:
            raise RuntimeError('app.js 中找不到 isWideChar')
        src = src[:old_fn.start()] + new_block + '\n' + src[old_fn.end():]
    with open(APP_JS, 'w', encoding='utf-8') as f:
        f.write(src)
    print(f'app.js 已更新：EAW_WIDE 表 {len(ranges)} 个区间')


def main():
    ranges = build_ranges()
    self_check(ranges)
    if '--check' in sys.argv:
        return 0
    patch_app_js(ranges)
    return 0


if __name__ == '__main__':
    sys.exit(main())