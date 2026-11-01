#!/usr/bin/env python3
"""v3 -> v4: delete month-sheet columns E, AG, AH; move AP:AV before AF; agreed-amount gross-up; 4 personnel types.
usage: python3 transform_xlsx.py in.xlsx out.xlsx"""
import copy, re, sys
import openpyxl
from openpyxl.formula.tokenizer import Tokenizer, Token
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

MONTHS = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
S = 'تنظیمات!'


def newcol(c):
    """old column index -> new index (None = deleted)"""
    if c <= 4: return c
    if c == 5: return None                 # E  code (old)
    if c <= 31: return c - 1               # F..AE -> E..AD
    if c == 32: return 38                  # AF net legal -> AL
    if c in (33, 34): return None          # AG payable, AH shortfall
    if c <= 41: return c + 4               # AI..AO -> AM..AS
    if c <= 48: return c - 11              # AP..AV -> AE..AK
    return c - 3                           # AW.. -> AT..


CELL = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d*)$')


def map_part(p):
    m = CELL.match(p)
    if not m: return p, None
    d1, col, d2, row = m.groups()
    n = newcol(CI(col))
    if n is None: raise ValueError(f'reference to deleted column: {p}')
    return f'{d1}{L(n)}{d2}{row}', (CI(col), n)


def map_ref(tok, host):
    sheet, ref = None, tok
    if '!' in tok:
        sheet, ref = tok.rsplit('!', 1)
    target = sheet.strip("'") if sheet else host
    if target not in MONTHS: return tok
    parts = ref.split(':')
    out, cols = [], []
    for p in parts:
        q, c = map_part(p); out.append(q); cols.append(c)
    if len(parts) == 2 and cols[0] and cols[1] and (cols[1][1] - cols[0][1]) != (cols[1][0] - cols[0][0]):
        raise ValueError(f'range split by move: {tok}')
    return (sheet + '!' if sheet else '') + ':'.join(out)


def map_formula(f, host):
    tok = Tokenizer(f)
    out = '='
    for t in tok.items:
        v = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            v = map_ref(v, host)
        out += v
    return out


def remove(f, part):
    assert part in f, (part, f)
    return f.replace(part, '', 1)


def main(src, dst):
    wb = openpyxl.load_workbook(src)

    # 1) summary column K (sum of old AG payable) -> sum of auto "other benefits" (old W)
    ys = wb['خلاصه سالانه']
    for r in range(4, ys.max_row + 1):
        v = ys[f'K{r}'].value
        if isinstance(v, str) and v.startswith('='):
            # payable = agreed target (old BJ -> new BG) when present, else legal net (old AF -> new AL)
            ys[f'K{r}'].value = re.sub(
                r"SUMIF\(([^!()]+)!\$A\$4:\$A\$(\d+),\$A(\d+),\1!\$AG\$4:\$AG\$\2\)",
                lambda m: (f'(SUMIF({m[1]}!$A$4:$A${m[2]},$A{m[3]},{m[1]}!$BJ$4:$BJ${m[2]})'
                           f'+SUMIFS({m[1]}!$AF$4:$AF${m[2]},{m[1]}!$A$4:$A${m[2]},$A{m[3]},{m[1]}!$BJ$4:$BJ${m[2]},""))'), v)
            assert '$AG$' not in ys[f'K{r}'].value, ys[f'K{r}'].value[:200]
    ys['K3'].value = 'جمع پرداختنی (توافقی، یا خالص قانونی اگر توافقی نیست)'

    for m in MONTHS:
        wb[m]['BD3'].value = 0          # old check on AH (deleted); rebuilt in step 5
    # 2) remap every formula that points into a month sheet (before moving cells; host = sheet title)
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith('=')): continue
                if ws.title in MONTHS and newcol(c.column) is None: continue   # cell is going away
                c.value = map_formula(v, ws.title)

    # 3) physically move month-sheet cells and column widths
    for m in MONTHS:
        ws = wb[m]
        widths = {}
        for k, d in list(ws.column_dimensions.items()):
            for i in range(d.min or CI(k), (d.max or CI(k)) + 1):
                widths[i] = d.width
        new = {}
        for (r, c), cell in ws._cells.items():
            n = newcol(c)
            if n is None: continue
            cell.column = n
            new[(r, n)] = cell
        ws._cells = new
        ws.column_dimensions.clear()
        for old, w in widths.items():
            n = newcol(old)
            if n and w: ws.column_dimensions[L(n)].width = w

    # 4) settings: tax helper (g = threshold - tax at threshold), personnel types, notes
    st = wb['تنظیمات']
    st['F22'].value, st['G22'].value = 'g ماهانه (آستانه − مالیات)', 'g سالانه (آستانه − مالیات)'
    for r in range(23, 28):
        st[f'F{r}'].value = f'=A{r}-SUMPRODUCT((A{r}>$A$23:$A$27)*(A{r}-$A$23:$A$27)*$C$23:$C$27)'
        st[f'G{r}'].value = f'=E{r}-SUMPRODUCT((E{r}>$E$23:$E$27)*(E{r}-$E$23:$E$27)*$C$23:$C$27)'
        for c in 'FG': st[f'{c}{r}']._style = copy.copy(st[f'D{r}']._style)
    for c in 'FG': st[f'{c}22']._style = copy.copy(st['D22']._style)
    types = [(1, 'مهندس ناظر کارگاهی', 0), (2, 'پرسنل فنی — استفاده از مدرک (حق مدرک؛ واریز در مقاطعی از سال)', 2),
             (3, 'پرسنل دفتری', 0), (4, 'مدیران', 0)]
    for i, (code, title, kind) in enumerate(types):
        st[f'A{32 + i}'].value, st[f'B{32 + i}'].value, st[f'C{32 + i}'].value = code, title, kind
    for c in 'ABC': st[f'{c}36'].value = None
    st['A30'].value = ('نوع پرسنل (ستون E «کد») — هر کد قرارداد مخصوص خودش را دارد. ستون سوم: ۰ = پرداخت ماهانه، '
                       '۲ = حق مدرک (حقوق هر ماه تعلق می‌گیرد و گزارش می‌شود ولی در مقاطعی از سال واریز می‌شود)')
    st['C3'].value = 'ستون AB (بیمه)'
    st['C7'].value = 'ستون R (مسکن)'
    st['C8'].value = 'ستون S (خواروبار)'
    st['C9'].value = 'ضرب در ستون تاهل (N)'
    st['C13'].value = 'ستون Z (اضافه‌کاری)'
    st['C15'].value = 'ستون‌های T و AA و AC'
    st['A44'].value = ('ستون‌های مزایا و کسور (قبل از خالص): AH مزایای مشمول بیمه و مالیات، AI مزایای معاف، AJ مساعده/وام، AK سایر کسور. '
                       'ستون‌های AE تا AG درآمد و مالیات تجمعی‌اند.')
    st['A45'].value = 'ستون AT = تعداد اولاد (حق اولاد = تعداد × B10)، ستون AU = بازنشسته (۱ = بیمه کسر نمی‌شود).'
    st['A46'].value = ('مبلغ توافقی (ستون F، خالص دریافتی به هزار ریال): اگر خالی باشد فقط فرمول‌های قانونی کار می‌کنند. '
                       'اگر پر باشد ستون V «سایر مزایا» خودکار طوری حساب می‌شود که خالص قانونی (AL) برابر توافقی + اضافه‌کار/ماموریت/مزایا − کسور شود. '
                       'اگر توافقی کمتر از خالص قانونی باشد، لیست قانونی (AL) جدا می‌ماند و همان توافقی پرداخت می‌شود (شیت «لیست پرداخت»). توافقی ۰ یعنی خالی.')
    st['A62'].value = 'سایر مزایا — خودکار تا مبلغ توافقی (ستون V)'
    st['D62'].value = 'مشمول بیمه را ۰ نگه دارید'
    st['A63'].value = 'مزایای اضافی مشمول بیمه و مالیات (ستون AH)'
    st['A64'].value = 'مزایای معاف از بیمه و مالیات (ستون AI)'
    st['A58'].value = 'پاداش (ستون K)'

    # 5) month sheets: new formulas
    tax_m = lambda x: f'SUMPRODUCT(({x}>{S}$A$23:$A$27)*({x}-{S}$A$23:$A$27)*{S}$C$23:$C$27)'
    tax_y = lambda x: f'SUMPRODUCT(({x}>{S}$E$23:$E$27)*({x}-{S}$E$23:$E$27)*{S}$C$23:$C$27)'
    for m in MONTHS:
        ws = wb[m]
        esf = m == 'اسفند'
        rows = [r for r in range(4, ws.max_row + 1) if isinstance(ws[f'AL{r}'].value, str) and ws[f'AL{r}'].value.startswith('=')]
        last = rows[-1]
        for r in rows:
            aa, t, ac = ws[f'AA{r}'].value, ws[f'T{r}'].value, ws[f'AC{r}'].value
            ws[f'BC{r}'] = remove(aa, f'+V{r}')
            ws[f'BD{r}'] = remove(t, f'+V{r}*{S}$B$62')
            ws[f'BE{r}'] = remove(ac, f'+V{r}*{S}$C$62')
            ws[f'AA{r}'] = f'=BC{r}+V{r}'
            ws[f'T{r}'] = f'=BD{r}+V{r}*{S}$B$62'
            ws[f'AC{r}'] = f'=BE{r}+V{r}*{S}$C$62'
            tax0 = f'({tax_y(f"(AE{r}+BE{r})")}-AF{r})' if esf else tax_m(f'BE{r}')
            ws[f'BF{r}'] = (f'=BC{r}-IF(AU{r}=1,0,{S}$B$3*MIN({S}$B$14*M{r},BD{r}+P{r}*{S}$B$52))-{tax0}+X{r}-AJ{r}-AK{r}')
            ws[f'BG{r}'] = f'=IF(L{r}="","",L{r}+0.9*Y{r}+Z{r}+X{r}+AH{r}+AI{r}-AJ{r}-AK{r})'
            ws[f'BH{r}'] = f'=AE{r}+BE{r}' if esf else f'=BE{r}'
            tx = tax_y(f'BH{r}') if esf else tax_m(f'BH{r}')
            ws[f'BI{r}'] = f'=IF(BG{r}="","",BH{r}-{tx}+BG{r}-BF{r})'
            g, thr = ('$G$23:$G$27', '$E$23:$E$27') if esf else ('$F$23:$F$27', '$A$23:$A$27')
            g1 = '$G$23' if esf else '$F$23'
            tstar = (f'IF(BI{r}<{S}{g1},BI{r},LOOKUP(BI{r},{S}{g},{S}{thr})+(BI{r}-LOOKUP(BI{r},{S}{g},{S}{g}))'
                     f'/(1-LOOKUP(BI{r},{S}{g},{S}$B$23:$B$27)))')
            ws[f'V{r}'] = f'=IF(BG{r}="",0,MAX(0,IF({S}$C$62=0,BG{r}-BF{r},{tstar}-BH{r})))'
            ws[f'L{r}'] = f'=IF(N(F{r})=0,"",F{r}*1000+K{r}*1000000)'
            ws[f'V{r}']._style = copy.copy(ws[f'U{r}']._style)
            for c in ('BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI'):
                ws[f'{c}{r}']._style = copy.copy(ws[f'AL{r}']._style)
        heads = {'BC': 'جمع مزایا بدون سایر مزایا', 'BD': 'مزایای مشمول بیمه بدون سایر مزایا', 'BE': 'مشمول مالیات بدون سایر مزایا',
                 'BF': 'خالص قانونی بدون سایر مزایا', 'BG': 'پرداختنی توافقی (توافقی + متغیرها − کسور)',
                 'BH': 'پایه مالیات' + (' سالانه' if esf else ''), 'BI': 'کمکی محاسبه سایر مزایا'}
        for c, h in heads.items():
            ws[f'{c}1'] = h; ws[f'{c}1']._style = copy.copy(ws['AL1']._style)
            ws.column_dimensions[c].width = 16
        ws['BB1'] = 'ستون‌های کمکی محاسبه «سایر مزایا» (دست نزنید) ←'
        ws['F1'] = 'مبلغ توافقی (خالص دریافتی، هزار ریال)'
        ws['L1'] = 'مبلغ توافقی (ریال)'
        ws['V1'] = 'سایر مزایا (خودکار تا مبلغ توافقی)'
        ws['AL1'] = 'خالص قابل پرداخت قانونی (لیست بیمه و مالیات)'
        ws['E1'] = 'کد نوع (۱ ناظر، ۲ مدرک، ۳ دفتری، ۴ مدیر)'
        # checks: old "payable < legal" -> "agreed < legal net"
        n = re.search(r'A4:A(\d+)', str(ws['AV3'].value)).group(1)
        ws['BA3'] = f'=SUMPRODUCT((A4:A{n}<>"")*(BG4:BG{n}<>"")*(BG4:BG{n}<BF4:BF{n}-0.5))'
        ws['BA2'] = 'توافقی<قانونی'
        ws['B3'] = ws['B3'].value.replace('نفر پرداختنی کمتر از خالص قانونی', 'نفر توافقی کمتر از خالص قانونی (توافقی پرداخت می‌شود)')
        if ws.print_area:
            ws.print_area = f'A1:AS{last}'

    # 6) payment-list sheet: INDIRECT strings use text column letters
    pl = wb['لیست پرداخت']
    MI = lambda col, r: f'INDIRECT("\'"&{S}$B$19&"\'!{col}{r}")'
    for r in range(4, pl.max_row + 1):
        if not isinstance(pl[f'A{r}'].value, str): continue
        for c in 'CEHIJ':
            v = pl[f'{c}{r}'].value
            if isinstance(v, str):
                v = (v.replace(f"'!F{r}\"", f"'!E{r}\"").replace(f"'!N{r}\"", f"'!M{r}\"").replace(f"'!AG{r}\"", f"'!AL{r}\"")
                      .replace(f"'!AF{r}\"", f"'!AL{r}\"").replace(f"'!P{r}\"", f"'!O{r}\""))
                pl[f'{c}{r}'].value = v
        pay = f'IF({MI("BG", r)}="",N({MI("AL", r)}),N({MI("BG", r)}))'
        pl[f'E{r}'] = pl[f'E{r}'].value.replace(f'N({MI("AL", r)})>0', f'{pay}>0')
        assert pay in pl[f'E{r}'].value
        pl[f'H{r}'] = f'=IF(E{r}="پرداخت",{pay},0)'
        pe = lambda col: f'INDEX(پرسنل!${col}$4:${col}$300,MATCH($A{r},پرسنل!$A$4:$A$300,0))'
        pl[f'F{r}'] = f'=IF(A{r}="","",IF(IFERROR({pe("D")}&"","")<>"",{pe("D")}&"",{MI("AS", r)}&""))'
        pl[f'G{r}'] = f'=IF(A{r}="","",IFERROR({pe("C")}&"",""))'
        j = pl[f'J{r}'].value
        assert j.endswith('"")))'), j[-30:]
        add = (f'&IF(AND(E{r}<>"قطع همکاری",N({MI("BG", r)})>0,N({MI("BG", r)})<N({MI("BF", r)})-0.5),'
               f'"مبلغ توافقی کمتر از خالص قانونی (توافقی پرداخت می‌شود)؛ ","")'
               f'&IF(AND(C{r}=4,{MI("BG", r)}=""),"مبلغ توافقی مدیر وارد نشده؛ ","")')
        pl[f'J{r}'] = j[:-2] + add + j[-2:]     # keep the closing TRIM( and IF( parens
    wb.calculation.fullCalcOnLoad = True
    wb.save(dst)
    print('saved', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
