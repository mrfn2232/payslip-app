#!/usr/bin/env python3
"""
ساخت ماه جدید از ماه قبل:  python new_month.py payroll_1405.xlsx --from آبان --to آذر [--keep-monthly]
- همه ستون‌ها (حقوق پایه، افراد جدید، ضرایب و …) از ماه قبل کپی می‌شود.
- ستون‌های متغیر ماهانه (H ساعت اضافه، I کسر کار، J ماموریت، K پاداش، AH و AI مزایا، AJ و AK کسور) صفر می‌شوند، مگر با --keep-monthly.
- مبلغ توافقی (F) و ستون «سایر مزایا» (V، فرمول) از ماه قبل می‌آیند.
- فرمول‌های مالیات تجمعی و تعداد روز ماه مقصد دست‌نخورده می‌ماند. اگر ردیف جدیدی در ماه قبل اضافه شده باشد، فرمول‌هایش هم ساخته می‌شود.
- بعد از اجرا فایل را در Excel باز و ذخیره کنید (تا فرمول‌ها محاسبه شوند) و در شیت تنظیمات «ماه جاری» را عوض کنید.
"""
import argparse, copy, re
import openpyxl
from openpyxl.formula.translate import Translator

MONTHS = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
SKIP_COLS = {'AD', 'AE', 'AF', 'AG', 'BF', 'BH', 'BI'}      # مخصوص هر ماه (مالیات تجمعی؛ اسفند جدول سالانه دارد)
SKIP_CELLS = {'N1'}
RESET = ('H', 'I', 'J', 'K', 'AH', 'AI', 'AJ', 'AK')

ap = argparse.ArgumentParser(); ap.add_argument('xlsx'); ap.add_argument('--from', dest='src', required=True)
ap.add_argument('--to', dest='dst', required=True); ap.add_argument('--keep-monthly', action='store_true')
ap.add_argument('--out'); a = ap.parse_args()
wb = openpyxl.load_workbook(a.xlsx)
src, dst = wb[a.src], wb[a.dst]
n = MONTHS.index(a.dst) + 1
S = "'تنظیمات'!"
last_dst = max(r for r in range(4, 200) if dst[f'A{r}'].value is not None)
rows = [r for r in range(4, 200) if src[f'A{r}'].value is not None]
copied = 0
for r in rows:
    for c in src[r]:
        col = c.column_letter
        if col in SKIP_COLS or f'{col}{r}' in SKIP_CELLS: continue
        t = dst[f'{col}{r}']
        t.value = c.value
        if r > last_dst: t._style = copy.copy(c._style)
        copied += 1
    if not a.keep_monthly:
        for col in RESET: dst[f'{col}{r}'].value = 0
        # overtime amount typed as a number (e.g. copied from an official list) -> back to the hours formula
        if not str(dst[f'Z{r}'].value or '').startswith('='):
            dst[f'Z{r}'].value = f'=(H{r}*{S}$B$13-I{r})*O{r}/{S}$B$12'
    if r > last_dst:   # ردیف جدید: فرمول‌های مالیات تجمعی
        for col in sorted(SKIP_COLS):
            f = dst[f'{col}{last_dst}'].value
            dst[f'{col}{r}'].value = Translator(f, origin=f'{col}{last_dst}').translate_formula(f'{col}{r}')
            dst[f'{col}{r}']._style = copy.copy(dst[f'{col}{last_dst}']._style)
pe = wb['پرسنل']
have = {pe[f'A{r}'].value for r in range(4, 300) if pe[f'A{r}'].value is not None}
nxt = max([x for x in have if isinstance(x, (int, float))] or [0]) + 1
added = 0
for r in rows:
    i = src[f'A{r}'].value
    if i not in have:
        row = next(x for x in range(4, 300) if pe[f'A{x}'].value is None)
        pe[f'A{row}'].value = i; have.add(i); added += 1
wb.save(a.out or a.xlsx)
if added: print(f'{added} کد جدید در شیت پرسنل ثبت شد؛ ایمیل و شبا را کامل کنید.')
print(f'{len(rows)} ردیف از «{a.src}» به «{a.dst}» کپی شد (شماره ماه مقصد: {n}).')
