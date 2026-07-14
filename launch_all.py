#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pathlib import Path
import sys

print('启动全量匹配...')

# 直接导入并执行英文命名的脚本
try:
    print('\n[1/2] 运行OMS全量...')
    exec(open('match_oms_1_9_months.py', encoding='utf-8').read())
except FileNotFoundError:
    print('未找到OMS脚本: match_oms_1_9_months.py')
    sys.exit(1)

try:
    print('\n[2/2] 运行DMS全量...')
    exec(open('match_dms_1_9_months_from_sql.py', encoding='utf-8').read())
except FileNotFoundError:
    print('未找到DMS脚本: match_dms_1_9_months_from_sql.py')
    sys.exit(1)

print('\n全部完成。')
