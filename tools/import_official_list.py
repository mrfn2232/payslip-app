#!/usr/bin/env python3
"""Fill all 12 month sheets of payroll v4 from an official monthly payroll list (e.g. Shahrivar 1405).
usage: python tools/import_official_list.py payroll_v4.xlsx official_list.xlsx out.xlsx report.csv"""
import copy, csv, re, sys
import openpyxl
from openpyxl.formula.translate import Translator

MONTHS = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
SRC_MONTH = 'شهریور'
src_path, sh_path, out_path, rep_path = sys.argv[1:5]

norm = lambda s: re.sub(r'[\s‌]', '', str(s or '').replace('ي', 'ی').replace('ك', 'ک').replace('ء', '').replace('أ', 'ا'))
num = lambda v: float(str(v).replace(',', '')) if str(v or '').replace(',', '').replace('.', '', 1).lstrip('-').isdigit() else 0.0
clean = lambda s: ' '.join(str(s or '').split())

# ---- official Shahrivar list
sh = openpyxl.load_workbook(sh_path, data_only=True)['sheet1']
hdr = [c.value for c in sh[1]]
col = {h.strip(): i for i, h in enumerate(hdr) if h}
people = []
for r in sh.iter_rows(min_row=2, values_only=True):
    if r[0] in (None, ''): continue
    g = lambda h: r[col[h]]
    people.append(dict(code=int(num(g('کد پرسنلی'))), name=clean(g('نام')), family=clean(g('نام خانوادگی')), father=clean(g('نام پدر')),
                       nid=clean(g('کد ملی')), kids=int(num(g('تعداد فرزندان'))), wage=num(g('مزد شغل روزانه')),
                       seniority=num(g('پایه سنوات روزانه')), days=num(g('روزهای کارکرد')), ot_h=num(g('اضافه کار واحد')),
                       ot=num(g('اضافه کار')), car=num(g('حق خودرو')), bonus=num(g('پاداش')), other=num(g('سایر')),
                       married=num(g('حق تاهل')) > 0, tax=num(g('مالیات حقوق')), ins=num(g('بیمه سهم کارمندی')),
                       gross=num(g('جمع ناخالص')), net=num(g('خالص پرداختی'))))
people.sort(key=lambda p: p['code'])

# ---- existing info by name (latest month first)
wb = openpyxl.load_workbook(src_path)
wv = openpyxl.load_workbook(src_path, data_only=True)
KEEP = ('D', 'E', 'F', 'G', 'AM', 'AO', 'AP', 'AS', 'AU')
old = {}
for m in reversed(MONTHS[:8]):
    w = wb[m]
    for r in range(4, w.max_row + 1):
        if w[f'A{r}'].value in (None, ''): continue
        k = norm(w[f'B{r}'].value) + norm(w[f'C{r}'].value)
        if k not in old:
            old[k] = dict({c: w[f'{c}{r}'].value for c in KEEP}, oldcode=w[f'A{r}'].value)
pers_old = {}
pe = wb['پرسنل']
for r in range(4, pe.max_row + 1):
    if pe[f'A{r}'].value not in (None, ''):
        pers_old[str(pe[f'A{r}'].value)] = {c: pe[f'{c}{r}'].value for c in 'CDEF'}
for p in people:
    p['old'] = old.get(norm(p['name']) + norm(p['family']))

# ---- write month sheets
st = wb['تنظیمات']
days_in = {st[f'G{r}'].value: st[f'H{r}'].value for r in range(3, 15)}
first = 4
for m in MONTHS:
    w = wb[m]
    tmpl = {c.column_letter: (c.value, copy.copy(c._style)) for c in w[first]}
    last_old = max(r for r in range(4, w.max_row + 1) if w[f'A{r}'].value not in (None, ''))
    for i, p in enumerate(people):
        r = first + i
        for L_, (v, sty) in tmpl.items():
            cell = w[f'{L_}{r}']
            if isinstance(v, str) and v.startswith('='):
                cell.value = Translator(v, origin=f'{L_}{first}').translate_formula(f'{L_}{r}')
            elif r > last_old:
                cell.value = None
            if r > last_old: cell._style = copy.copy(sty)
        o = p['old'] or {}
        full = days_in[m]
        d = p['days'] if p['days'] < days_in[SRC_MONTH] else full
        d = min(d, full)
        vals = {'A': p['code'], 'B': p['name'], 'C': p['family'], 'D': o.get('D', 1), 'E': o.get('E'), 'F': o.get('F'),
                'G': o.get('G'), 'H': p['ot_h'], 'I': 0, 'J': 0, 'K': p['bonus'] / 1e6, 'N': 1 if p['married'] else 0,
                'O': p['wage'] + p['seniority'], 'Y': p['car'], 'AH': p['other'], 'AI': 0, 'AJ': 0, 'AK': 0,
                'AM': o.get('AM') or p['nid'], 'AN': p['father'], 'AO': o.get('AO'), 'AP': o.get('AP'), 'AQ': None, 'AR': None,
                'AS': o.get('AS'), 'AT': p['kids'], 'AU': 0}
        for c, v in vals.items(): w[f'{c}{r}'].value = v
        if d != full: w[f'M{r}'].value = d                      # partial month: days as a number
        if p['ot'] > 0:                                          # overtime amount exactly as in the official list
            w[f'Z{r}'].value = round(p['ot'])
        if p['ins'] == 0 and p['gross'] > 0:                     # official list deducts no insurance for this person
            w[f'AU{r}'].value = 1
        w[f'Q{r}'].value = f'=AT{r}*تنظیمات!$B$10*M{r}/M$1'       # child allowance prorated by days (as official list)
        v = w[f'V{r}'].value
        assert v.startswith(f'=IF(BG{r}="",0,'), v[:40]
        w[f'V{r}'].value = v.replace(f'=IF(BG{r}="",0,', f'=IF(OR(BG{r}="",N(M{r})=0),0,', 1)   # no top-up without work days
    for r in range(first + len(people), max(w.max_row, last_old) + 1):   # clear leftover old rows
        for c in w[r]: c.value = None
    w['AM1'].value = 'شماره شناسنامه / کد ملی'
    w['AU1'].value = 'بدون کسر بیمه (۱ = بله؛ بازنشسته یا معاف)'

# ---- personnel sheet: permanent code list in the new numbering
for r in range(4, pe.max_row + 1):
    for c in 'ACDEF': pe[f'{c}{r}'].value = None
for i, p in enumerate(people):
    r = 4 + i
    pe[f'A{r}'].value = p['code']
    o = pers_old.get(str((p['old'] or {}).get('oldcode')), {})
    for c in 'CDEF': pe[f'{c}{r}'].value = o.get(c)

# ---- settings: the official list has no daily company benefit
st['B15'].value = 0
st['C15'].value = 'در لیست رسمی شهریور ۱۴۰۵ این قلم نیست؛ ۰ گذاشته شد'
for r in (53, 54, 55):                       # official list taxes housing, grocery and child allowance
    st[f'C{r}'].value = 1
st['B59'].value = 1                          # official list: overtime is subject to insurance
st['A66'].value = ('مشمولیت مالیات طبق لیست رسمی شهریور ۱۴۰۵ تنظیم شد: مسکن، بن و حق اولاد هم مشمول مالیات‌اند. '
                   'حق اولاد مثل مسکن و بن به نسبت روز کارکرد حساب می‌شود. اضافه‌کار مشمول بیمه است.')
wb.calculation.fullCalcOnLoad = True
wb.save(out_path)

with open(rep_path, 'w', newline='', encoding='utf-8-sig') as fh:
    wr = csv.writer(fh)
    wr.writerow(['کد پرسنلی', 'نام', 'نام خانوادگی', 'کد قبلی در payroll', 'کد نوع', 'مبلغ توافقی (هزار ریال)', 'وضعیت'])
    for p in people:
        o = p['old']
        wr.writerow([p['code'], p['name'], p['family'], o['oldcode'] if o else '', o.get('E') if o else '', o.get('F') if o else '',
                     'پیدا شد' if o else 'جدید — کد نوع و توافقی را وارد کنید'])
print('people', len(people), 'matched', sum(1 for p in people if p['old']))
