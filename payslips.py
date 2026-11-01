#!/usr/bin/env python3
"""
ساخت فیش حقوقی، قرارداد (چهار نوع)، ایمیل و لیست پرداخت از روی payroll_1405_v4.xlsx

پیش‌نیاز:  pip install openpyxl            (برای PDF:  pip install playwright && playwright install chromium)
فایل اکسل باید یک‌بار در Excel ذخیره شده باشد تا مقدار فرمول‌ها داخلش باشد.

نمونه‌ها:
  python payslips.py payroll_1405.xlsx --period "آبان ۱۴۰۵" --out slips            # فقط HTML
  python payslips.py payroll_1405.xlsx --period "آبان ۱۴۰۵" --out slips --pdf      # HTML + PDF
  python payslips.py payroll_1405.xlsx --period "آبان ۱۴۰۵" --out slips --pdf --email --dry-run
  python payslips.py payroll_1405.xlsx --period "آبان ۱۴۰۵" --out slips --pdf --email   # ارسال واقعی
  python payslips.py payroll_1405.xlsx --payment-list pay.csv                          # فقط لیست پرداخت
تنظیم ایمیل با متغیرهای محیطی: SMTP_HOST SMTP_PORT SMTP_USER SMTP_PASS SMTP_FROM
"""
import argparse, csv, html, os, smtplib, ssl, sys
from email.message import EmailMessage
import openpyxl

PERSIAN = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')

# ستون‌های شیت ماه (نسخه v4 اکسل: E و AG و AH حذف شده‌اند، AP:AV قبل از خالص آمده‌اند)
COL = dict(name='B', family='C', code='E', agreed_k='F', note='G', overtime_hours='H', deduct_hours='I', bonus='K',
           contract_total='L', days='M', daily='O', monthly='P', child='Q', housing='R', grocery='S', benefits='T',
           insured_total='U', nonins_tax='V', mission='X', car='Y', overtime='Z', gross_legal='AA', insurance='AB',
           taxable='AC', tax='AD', extra_ins='AH', exempt='AI', loan='AJ', other_ded='AK', net_legal='AL',
           birth_cert='AM', father='AN', ins_no='AO', job='AP', account='AS', target='BG')
LAYOUT_CHECK = {'AL': 'قابل پرداخت', 'F': 'توافقی', 'V': 'سایر مزایا'}
DEFAULT_TYPES = {1: ('مهندس ناظر کارگاهی', 0), 2: ('پرسنل فنی — استفاده از مدرک (حق مدرک)', 2),
                 3: ('پرسنل دفتری', 0), 4: ('مدیران', 0)}


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def txt(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return ' '.join(str(v).split())


def load_types(wb):
    st = wb['تنظیمات']; types = {}
    for r in range(32, 37):
        c = st[f'A{r}'].value
        if isinstance(c, (int, float)):
            types[int(c)] = (txt(st[f'B{r}'].value), int(num(st[f'C{r}'].value)))
    return types or dict(DEFAULT_TYPES)


def load(path, sheet=None):
    """همه‌چیز از شیت همان ماه خوانده می‌شود (نه از «لیست پرداخت» که فقط ماه جاری را نشان می‌دهد)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    st = wb['تنظیمات']
    sheet = sheet or st['B19'].value                       # پیش‌فرض: ماه جاری در تنظیمات
    if sheet not in wb.sheetnames: raise ValueError(f'شیت «{sheet}» در فایل نیست')
    main = wb[sheet]
    bad = [c for c, h in LAYOUT_CHECK.items() if h not in txt(main[f'{c}1'].value)]
    if bad: raise ValueError('ستون‌های این فایل با نسخه v4 جور نیست (ستون‌های ' + '، '.join(bad) +
                             '). فایل را با نسخه جدید payroll_1405_v4.xlsx جایگزین کنید.')
    if main['AL4'].value is None and main['AL5'].value is None:
        raise ValueError('مقدار فرمول‌ها در فایل نیست؛ فایل را یک بار در Excel باز و ذخیره کنید.')
    types, min_daily = load_types(wb), num(st['B17'].value)
    pe, pers = wb['پرسنل'], {}
    for r in range(4, pe.max_row + 1):
        k = pe[f'A{r}'].value
        if k not in (None, ''):
            pers[txt(k)] = dict(email=txt(pe[f'C{r}'].value), iban=txt(pe[f'D{r}'].value), manual=txt(pe[f'E{r}'].value))
    people = []
    for r in range(4, main.max_row + 1):
        pid = main[f'A{r}'].value
        if pid in (None, '') or not (txt(main[f'B{r}'].value) or txt(main[f'C{r}'].value)): continue
        p = {k: main[f'{c}{r}'].value for k, c in COL.items()}
        info = pers.get(txt(pid), {})
        code = int(num(p['code']))
        title, kind = types.get(code, ('کد ناشناخته', 0))
        agreed = p['target'] not in (None, '')
        payable = num(p['target']) if agreed else num(p['net_legal'])
        status = info.get('manual') or ('حق مدرک' if kind == 2 else 'بدون کارکرد' if num(p['days']) == 0
                                        else 'پرداخت' if payable > 0.5 else 'صفر - بررسی')
        iban = info.get('iban') or txt(p['account'])
        email = info.get('email', '')
        w = []
        if code not in types: w.append('کد نوع ناشناخته')
        if status not in ('قطع همکاری', 'بدون کارکرد') and not email: w.append('ایمیل ندارد')
        if status == 'پرداخت' and not iban: w.append('شبا ندارد')
        if status == 'صفر - بررسی': w.append('پرداختنی صفر است')
        if agreed and payable < num(p['net_legal']) - 0.5: w.append('مبلغ توافقی کمتر از خالص قانونی (توافقی پرداخت می‌شود)')
        if code == 4 and not agreed: w.append('مبلغ توافقی مدیر وارد نشده')
        if status not in ('حق مدرک', 'قطع همکاری') and 0 < num(p['daily']) < min_daily: w.append('دستمزد روزانه کمتر از حداقل')
        p.update(row=r, id=pid, code=code, full=' '.join(f"{txt(p['name'])} {txt(p['family'])}".split()), type=title,
                 kind=kind, status=status, iban=iban, email=email, payable=payable, agreed=agreed,
                 pay=payable if status == 'پرداخت' else 0.0, accrued=num(p['net_legal']) if status == 'حق مدرک' else 0.0,
                 warn='؛ '.join(w), month=sheet)
        people.append(p)
    return people


CSS = """
body{font-family:Vazirmatn,'B Nazanin',Tahoma,'DejaVu Sans',sans-serif;direction:rtl;margin:0;color:#111}
.slip{width:182mm;margin:8mm auto;border:1.5px solid #333;padding:7mm;box-sizing:border-box}
h1{font-size:17px;margin:0 0 2px;text-align:center}.sub{text-align:center;color:#555;font-size:12px;margin-bottom:8px}
table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}
th,td{border:1px solid #999;padding:4px 7px}th{background:#e8eef7;text-align:right}
td.n{text-align:left;direction:ltr;font-variant-numeric:tabular-nums}
.info td{border:none;padding:2px 4px}.tot td{font-weight:bold;background:#f3f3f3}
.net{font-size:15px;background:#fff6cc!important}.foot{margin-top:14px;font-size:11px;color:#555;display:flex;justify-content:space-between}
.slip{page-break-after:always;break-after:page}.slip:last-child{page-break-after:auto;break-after:auto}
@page{size:A4;margin:8mm}
"""


def build_html(p, period, company, layout, digits):
    def f(x):
        s_ = f"{num(x):,.0f}"
        s_ = s_.translate(PERSIAN) if digits else s_
        return '\u200e' + s_ if s_.startswith('-') else s_
    t = lambda s: (str(s).translate(PERSIAN) if digits else str(s))
    e = lambda s: html.escape(str(s if s is not None else ''))
    other = num(p['gross_legal']) - sum(num(p[k]) for k in ('monthly', 'housing', 'grocery', 'child', 'overtime', 'car', 'nonins_tax', 'extra_ins', 'exempt'))
    rows_e = [('حقوق ماهانه (دستمزد روزانه × روز کارکرد)', p['monthly']), ('حق مسکن', p['housing']), ('بن کارگری', p['grocery']),
              ('حق اولاد', p['child']), ('سایر (حق تاهل، مزایای روزانه، پاداش)', other), ('اضافه‌کاری (پس از کسر کارکرد)', p['overtime']),
              ('حق خودرو', p['car']), ('سایر مزایا', p['nonins_tax']), ('مزایای اضافی مشمول', p['extra_ins']),
              ('مزایای معاف از بیمه و مالیات', p['exempt'])]
    rows_e = [(k, v) for k, v in rows_e if abs(num(v)) > 0.5]
    rows_d = [('بیمه سهم کارگر', p['insurance']), ('مالیات بر حقوق', p['tax']), ('مساعده / قسط وام', p['loan']), ('سایر کسور', p['other_ded'])]
    rows_d = [(k, v) for k, v in rows_d if abs(num(v)) > 0.5]
    ded_total = sum(num(v) for _, v in rows_d)
    deferred = p['status'] == 'حق مدرک'
    legal = layout in ('legal', 'both') or deferred
    actual = layout in ('actual', 'both') and not deferred
    body = ''
    if legal:
        body += '<table><tr><th colspan=2>حقوق و مزایا (طبق لیست قانونی)</th></tr>'
        body += ''.join(f'<tr><td>{k}</td><td class=n>{f(v)}</td></tr>' for k, v in rows_e)
        body += f'<tr class=tot><td>جمع حقوق و مزایا</td><td class=n>{f(p["gross_legal"])}</td></tr></table>'
        body += '<table><tr><th colspan=2>کسورات</th></tr>'
        body += ''.join(f'<tr><td>{k}</td><td class=n>{f(v)}</td></tr>' for k, v in rows_d)
        body += f'<tr class=tot><td>جمع کسورات</td><td class=n>{f(ded_total)}</td></tr></table>'
        if abs(num(p['mission'])) > 0.5:
            body += f'<table><tr><td>حق ماموریت (افزوده به خالص)</td><td class=n>{f(p["mission"])}</td></tr></table>'
        body += f'<table><tr class="tot net"><td>خالص قابل پرداخت (طبق لیست)</td><td class=n>{f(p["net_legal"])}</td></tr></table>'
    if actual:
        body += '<table><tr><th colspan=2>پرداخت</th></tr>'
        if p['agreed']:
            body += f'<tr><td>مبلغ توافقی (خالص دریافتی)</td><td class=n>{f(p["contract_total"])}</td></tr>'
            for k, v in (('اضافه‌کاری', num(p['overtime'])), ('حق ماموریت', num(p['mission'])), ('حق خودرو', 0.9 * num(p['car'])),
                         ('مزایای اضافی مشمول', num(p['extra_ins'])), ('مزایای معاف', num(p['exempt'])),
                         ('مساعده / قسط وام (کسر)', -num(p['loan'])), ('سایر کسور', -num(p['other_ded']))):
                if abs(v) > 0.5: body += f'<tr><td>{k}</td><td class=n>{f(v)}</td></tr>'
        else:
            body += f'<tr><td>خالص حقوق طبق لیست قانونی</td><td class=n>{f(p["net_legal"])}</td></tr>'
        body += f'<tr class="tot net"><td>مبلغ قابل پرداخت (ریال)</td><td class=n>{f(p["payable"])}</td></tr></table>'
    if deferred:
        body += ('<table><tr><th colspan=2>وضعیت پرداخت</th></tr>'
                 f'<tr><td>پرداخت نقدی این ماه</td><td class=n>{f(0)}</td></tr>'
                 f'<tr><td>مبلغ تعلق‌گرفته این ماه (به حساب حق مدرک منظور شد)</td><td class=n>{f(p["accrued"])}</td></tr>'
                 f'<tr class="tot net"><td>این مبلغ در مقاطعی از سال واریز می‌شود</td><td class=n>{f(0)}</td></tr></table>')
    info = (f'<table class=info><tr><td>نام: <b>{e(p["full"])}</b></td><td>شغل: {e(p["job"])}</td></tr>'
            f'<tr><td>شماره بیمه: {e(t(p["ins_no"] or ""))}</td><td>روز کارکرد: {t(int(num(p["days"])))}</td></tr>'
            f'<tr><td>ساعت اضافه‌کار: {t(p["overtime_hours"] or 0)}</td><td>ساعت کسر کار: {t(p["deduct_hours"] or 0)}</td></tr></table>')
    return (f'<!doctype html><html lang=fa dir=rtl><meta charset=utf-8><title>فیش حقوقی {e(p["full"])}</title><style>{CSS}</style>'
            f'<div class=slip><h1>فیش حقوقی — {e(t(period))}</h1><div class=sub>{e(company)}</div>{info}{body}'
            f'<div class=foot><span>مبالغ به ریال</span><span>این فیش به صورت خودکار تولید شده است.</span></div></div></html>')


import re
DATE = re.compile(r'1[34]\d\d/\d{1,2}/\d{1,2}')

# چهار الگوی قرارداد بر اساس کد نوع (ستون E). {job} و {workplace} جایگزین می‌شوند.
CONTRACTS = {
    1: dict(title='قرارداد کار — مهندس ناظر کارگاهی',
            subject='انجام وظایف مهندس ناظر در سمت «{job}» در کارگاه‌های پروژه‌های کارفرما (کارگاه فعلی: {workplace})، شامل نظارت بر اجرای کار '
                    'مطابق نقشه‌ها و مشخصات فنی، تهیه گزارش روزانه و کمک به تنظیم صورت‌وضعیت‌ها، و کنترل رعایت ضوابط ایمنی و بهداشت کار در کارگاه.',
            hours='ساعات کار تابع برنامه کارگاه و حداکثر ۴۴ ساعت در هفته است. اضافه‌کاری با توافق و با ضریب ۱٫۴ مزد ساعتی پرداخت می‌شود.',
            extra=['محل کار و جابه‌جایی: محل کار کارگاه است. کارفرما می‌تواند کارگر را با اطلاع قبلی به کارگاه‌های دیگر خود منتقل کند؛ '
                   'هزینه ماموریت و ایاب‌وذهاب طبق ضوابط شرکت پرداخت می‌شود.',
                   'ایمنی: کارگر موظف است دستورالعمل‌های ایمنی کارگاه را رعایت کند و موارد خطر را بی‌درنگ به سرپرست کارگاه گزارش دهد.']),
    2: dict(title='قرارداد کار — پرسنل فنی (استفاده از مدرک)',
            subject='همکاری فنی در سمت «{job}» و استفاده کارفرما از مدرک تحصیلی و صلاحیت حرفه‌ای کارگر برای احراز و حفظ صلاحیت و رتبه شرکت '
                    'نزد مراجع ذی‌ربط؛ و حضور و انجام کار فنی در محل {workplace} هر زمان که کارفرما بخواهد.',
            hours='ساعات کار طبق برنامه‌ای است که کارفرما ابلاغ می‌کند و حداکثر ۴۴ ساعت در هفته است.',
            extra=['تعهد مدرک: کارگر متعهد است در مدت این قرارداد مدرک و صلاحیت خود را برای احراز صلاحیت شخص یا شرکت دیگری به کار نگیرد، '
                   'مدارک لازم را در اختیار کارفرما بگذارد و برای پایان همکاری دست‌کم یک ماه قبل کتباً اطلاع دهد.']),
    3: dict(title='قرارداد کار — پرسنل دفتری',
            subject='انجام امور اداری و دفتری در سمت «{job}» در دفتر کارفرما ({workplace}) طبق شرح وظایفی که کارفرما ابلاغ می‌کند.',
            hours='ساعات کار طبق ساعات اداری شرکت و حداکثر ۴۴ ساعت در هفته است. اضافه‌کاری با توافق و با ضریب ۱٫۴ مزد ساعتی پرداخت می‌شود.',
            extra=['محرمانگی: کارگر اسناد، اطلاعات مالی و اطلاعات پرسنلی شرکت را محرمانه نگه می‌دارد و بدون اجازه کتبی در اختیار دیگران نمی‌گذارد.']),
    4: dict(title='قرارداد کار — مدیران',
            subject='تصدی سمت مدیریتی «{job}» شامل برنامه‌ریزی، هدایت و نظارت بر واحد مربوط و گزارش به مدیرعامل، در حدود اختیاراتی '
                    'که کتباً ابلاغ می‌شود؛ محل کار: {workplace}.',
            hours='ساعات کار طبق برنامه شرکت است. کار بیش از ساعات قانونی با توافق طرفین و طبق قانون کار انجام می‌شود.',
            extra=['محرمانگی و تعارض منافع: کارگر اطلاعات شرکت را محرمانه نگه می‌دارد و در مدت قرارداد بدون اجازه کتبی کارفرما برای شرکت‌های '
                   'هم‌رشته فعالیت نمی‌کند.',
                   'تحویل امور: هنگام پایان قرارداد، کارگر اسناد، اموال و امور در جریان را صورت‌جلسه کرده و به جانشین یا کارفرما تحویل می‌دهد.']),
}
GENERIC_CONTRACT = dict(title='قرارداد کار', subject='انجام کار در سمت «{job}» در محل: {workplace}.',
                        hours='حداکثر ۴۴ ساعت در هفته؛ اضافه‌کاری با توافق و با ضریب ۱٫۴ پرداخت می‌شود.', extra=[])
DEFERRED_CLAUSE = ('نحوه پرداخت (حق مدرک): حقوق و مزایای هر ماه طبق ماده ۵ محاسبه و تعلق می‌گیرد و هر ماه به سازمان تأمین اجتماعی و '
                   'سازمان امور مالیاتی اعلام می‌شود؛ پرداخت نقدی جمع آن با توافق طرفین در مقاطعی از سال واریز می‌شود.')


def build_contract(p, a, digits=True):
    t = lambda x: (str(x).translate(PERSIAN) if digits else str(x))
    m = lambda x: t(f"{num(x):,.0f}")
    e = lambda x: html.escape(str(x if x else '………………'))
    tpl = CONTRACTS.get(p['code'], GENERIC_CONTRACT)
    d = DATE.search(str(p.get('note') or ''))
    start = d.group(0) if d else a.contract_start
    fill = lambda s_: html.escape(s_).replace('{job}', e(p['job'])).replace('{workplace}', e(a.workplace))
    rows = [('دستمزد روزانه', p['daily']), ('حق مسکن (ماهانه)', p['housing']), ('بن کارگری (ماهانه)', p['grocery']),
            ('حق اولاد', p['child'])]
    if p['agreed']: rows.append(('مبلغ توافقی ماهانه (خالص دریافتی)', p['contract_total']))
    wage = ''.join(f'<tr><td>{k}</td><td class=n>{m(v)} ریال</td></tr>' for k, v in rows if abs(num(v)) > 0.5)
    extra = list(tpl['extra']) + ([DEFERRED_CLAUSE] if p['status'] == 'حق مدرک' else [])
    arts = ''.join(f'<p><b>ماده {t(7 + i)} — {html.escape(x.split(":", 1)[0])}:</b>{html.escape(x.split(":", 1)[1])}</p>'
                   for i, x in enumerate(extra))
    last = t(7 + len(extra))
    return f"""<!doctype html><html lang=fa dir=rtl><meta charset=utf-8><title>قرارداد {html.escape(p['full'])}</title><style>{CSS}
.slip{{width:180mm}}p{{font-size:13px;line-height:2;margin:4px 0}}.sig{{display:flex;justify-content:space-between;margin-top:40px;font-size:13px}}</style>
<div class=slip><h1>{html.escape(tpl['title'])}</h1><div class=sub>{e(a.company)} — پیش‌نویس الگو؛ پیش از امضا توسط مشاور حقوقی بررسی شود</div>
<p><b>ماده ۱ — طرفین:</b> کارفرما: {e(a.company)} (نماینده: {e(a.rep)}) و کارگر: <b>{html.escape(p['full'])}</b>، نام پدر: {e(p['father'])}،
شماره شناسنامه: {e(t(txt(p['birth_cert'])))}، شماره بیمه: {e(t(txt(p['ins_no'])))}.</p>
<p><b>ماده ۲ — موضوع:</b> {fill(tpl['subject'])}</p>
<p><b>ماده ۳ — مدت:</b> از تاریخ {e(t(start))} به مدت {e(t(a.duration))}.</p>
<p><b>ماده ۴ — ساعات کار:</b> {html.escape(tpl['hours'])}</p>
<p><b>ماده ۵ — حق‌السعی:</b></p><table>{wage}</table>
<p><b>ماده ۶ — بیمه و مالیات:</b> حق بیمه سهم کارگر (۷٪) و مالیات بر حقوق طبق قوانین جاری از مبلغ پرداختی کسر و توسط کارفرما واریز می‌شود.</p>
{arts}<p><b>ماده {last} — سایر:</b> سایر موارد مطابق قانون کار جمهوری اسلامی ایران است. قرارداد در دو نسخه تنظیم و هر دو نسخه حکم واحد دارد.</p>
<div class=sig><span>کارفرما: {e(a.rep)}<br><br>امضا</span><span>کارگر: {html.escape(p['full'])}<br><br>امضا</span></div></div></html>"""


def find_browser():
    import shutil
    cands = [r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
             r'C:\Program Files\Google\Chrome\Application\chrome.exe', r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe']
    for c in cands:
        if os.path.exists(c): return c
    for n in ('msedge', 'google-chrome', 'chromium', 'chromium-browser', 'chrome'):
        w = shutil.which(n)
        if w: return w
    return None


def edge_pdf(br, html_path, pdf_path, timeout=60):
    """یک PDF با Edge/Chrome. با پروفایل موقت (تا اگر مرورگر باز است گیر نکند) و دو حالت headless پشت‌سرهم."""
    import subprocess, tempfile, shutil, time, pathlib
    uri = pathlib.Path(os.path.abspath(html_path)).as_uri()
    for mode in ('--headless=new', '--headless'):
        prof = tempfile.mkdtemp(prefix='ps_')
        if os.path.exists(pdf_path): os.remove(pdf_path)
        cmd = [br, mode, '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-sync',
               '--disable-background-networking', '--disable-component-update', f'--user-data-dir={prof}', '--no-pdf-header-footer',
               f'--print-to-pdf={os.path.abspath(pdf_path)}', uri]
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        t0 = time.time()
        try:
            while time.time() - t0 < timeout:
                if proc.poll() is not None: break
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    time.sleep(0.7); break
                time.sleep(0.3)
        finally:
            if proc.poll() is None: proc.kill()
            shutil.rmtree(prof, ignore_errors=True)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0: return True
    return False


def combine(htmls):
    """چند فایل HTML فیش/قرارداد را در یک سند پشت‌سرهم (هر کدام یک صفحه) ادغام می‌کند."""
    head = htmls[0][:htmls[0].index('<div class=slip>')]
    return head + ''.join(h[h.index('<div class=slip>'):h.rindex('</html>')] for h in htmls) + '</html>'


def to_pdf(html_files):
    """PDF: اول Edge/Chrome نصب‌شده روی سیستم؛ اگر نبود playwright. خروجی: لیست PDFهای ساخته‌شده."""
    br = find_browser()
    out = []
    if br:
        for h in html_files:
            pdf = h[:-5] + '.pdf'
            if edge_pdf(br, h, pdf): out.append(pdf)
        return out
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        b = pw.chromium.launch(); pg = b.new_page()
        for h in html_files:
            pg.goto('file://' + os.path.abspath(h)); pdf = h[:-5] + '.pdf'
            pg.pdf(path=pdf, format='A4', print_background=True); out.append(pdf)
        b.close()
    return out


def send(p, attach, period, dry):
    msg = EmailMessage()
    msg['Subject'] = f'فیش حقوقی {period}'; msg['From'] = os.environ.get('SMTP_FROM', os.environ.get('SMTP_USER', ''))
    msg['To'] = p['email']
    msg.set_content(f'{p["full"]} عزیز،\nفیش حقوقی {period} شما پیوست است.\nدر صورت هرگونه مغایرت با واحد مالی تماس بگیرید.')
    mt = ('application', 'pdf') if attach.endswith('.pdf') else ('text', 'html')
    with open(attach, 'rb') as fh:
        msg.add_attachment(fh.read(), maintype=mt[0], subtype=mt[1], filename=os.path.basename(attach))
    if dry:
        return 'dry-run'
    ctx = ssl.create_default_context()
    with smtplib.SMTP(os.environ['SMTP_HOST'], int(os.environ.get('SMTP_PORT', 587))) as s:
        s.starttls(context=ctx); s.login(os.environ['SMTP_USER'], os.environ['SMTP_PASS']); s.send_message(msg)
    return 'sent'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('xlsx'); ap.add_argument('--period', default=''); ap.add_argument('--company', default='')
    ap.add_argument('--out', default='slips'); ap.add_argument('--layout', choices=['actual', 'legal', 'both'], default='actual')
    ap.add_argument('--pdf', action='store_true'); ap.add_argument('--email', action='store_true')
    ap.add_argument('--dry-run', action='store_true'); ap.add_argument('--english-digits', action='store_true')
    ap.add_argument('--only', help='فقط این ردیف‌ها (شماره ردیف اکسل، با کاما)')
    ap.add_argument('--contracts', action='store_true', help='قرارداد کار (الگوی هر کد نوع) برای همه به‌جز قطع همکاری')
    ap.add_argument('--contract-register', metavar='CSV', help='گزارش وضعیت قرارداد و فیش همه پرسنل')
    ap.add_argument('--rep', default=''); ap.add_argument('--workplace', default=''); ap.add_argument('--duration', default='یک سال')
    ap.add_argument('--contract-start', default='')
    ap.add_argument('--payment-list', metavar='CSV'); ap.add_argument('--sheet', help='نام شیت ماه (پیش‌فرض: ماه جاری در تنظیمات)')
    a = ap.parse_args()
    people = load(a.xlsx, a.sheet)
    if a.payment_list:
        with open(a.payment_list, 'w', newline='', encoding='utf-8-sig') as fh:
            w = csv.writer(fh); w.writerow(['ردیف', 'کد', 'نام', 'شبا/حساب', 'مبلغ (ریال)', 'وضعیت'])
            for p in people:
                if p['status'] == 'پرداخت': w.writerow([p['row'], p['id'], p['full'], p['iban'], int(round(p['pay'])), p['status']])
        print('payment list written:', a.payment_list)
        if not a.period: return
    os.makedirs(a.out, exist_ok=True)
    if a.contracts or a.contract_register:
        cf = {}
        if a.contracts:
            for p in people:
                if p['status'] == 'قطع همکاری': continue
                fn = os.path.join(a.out, f"contract_{p['row']:03d}.html")
                with open(fn, 'w', encoding='utf-8') as fh:
                    fh.write(build_contract(p, a, not a.english_digits))
                cf[p['row']] = fn
            if a.pdf:
                for (r, _), pdf in zip(list(cf.items()), to_pdf(list(cf.values()))): cf[r] = pdf
        if a.contract_register:
            with open(a.contract_register, 'w', newline='', encoding='utf-8-sig') as fh:
                w = csv.writer(fh); w.writerow(['ردیف', 'نام', 'نوع', 'وضعیت', 'پرداخت نقدی این ماه', 'تعلق‌گرفته این ماه (حق مدرک)', 'ایمیل', 'شبا', 'فایل قرارداد', 'هشدار'])
                for p in people:
                    note = cf.get(p['row'], '')
                    w.writerow([p['row'], p['full'], p['type'], p['status'], int(round(p['pay'])), int(round(p['accrued'])), p['email'], p['iban'], note, p['warn']])
            print('register written:', a.contract_register)
        print(f'{len(cf)} قرارداد ساخته شد')
        if not a.period: return
    only = {int(x) for x in a.only.split(',')} if a.only else None
    log, files = [], []
    for p in people:
        if only and p['row'] not in only: continue
        if p['status'] not in ('پرداخت', 'حق مدرک') and not only: continue
        fn = os.path.join(a.out, f"slip_{p['row']:03d}.html")
        with open(fn, 'w', encoding='utf-8') as fh:
            fh.write(build_html(p, a.period, a.company, a.layout, not a.english_digits))
        files.append((p, fn))
    if a.pdf:
        pdfs = to_pdf([f for _, f in files]); files = [(p, pdf) for (p, _), pdf in zip(files, pdfs)]
    for p, fn in files:
        st = 'created'
        if a.email:
            if not p['email']: st = 'NO EMAIL'
            else:
                try: st = send(p, fn, a.period, a.dry_run)
                except Exception as ex: st = f'ERROR {ex}'
        log.append([p['row'], p['full'], p['email'], fn, st])
    with open(os.path.join(a.out, 'log.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
        csv.writer(fh).writerows([['ردیف', 'نام', 'ایمیل', 'فایل', 'نتیجه']] + log)
    print(f'{len(log)} فیش ساخته شد؛ گزارش: {a.out}/log.csv')


if __name__ == '__main__':
    main()
