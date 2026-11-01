# -*- coding: utf-8 -*-
"""برنامه‌ی ویندوز فیش‌ساز حقوق: انتخاب فایل اکسل، ماه، و ساخت فیش/قرارداد/لیست پرداخت (روی پایتون tkinter)."""
import os, csv, sys, traceback, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from types import SimpleNamespace
import openpyxl
import payslips as P

MONTHS = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
SETTINGS = os.path.join(os.path.expanduser('~'), '.payslip_app.txt')


def load_settings():
    d = {}
    try:
        for ln in open(SETTINGS, encoding='utf-8'):
            if '=' in ln: k, v = ln.rstrip('\n').split('=', 1); d[k] = v
    except OSError: pass
    return d


def save_settings(d):
    try:
        with open(SETTINGS, 'w', encoding='utf-8') as f:
            for k, v in d.items(): f.write(f'{k}={v}\n')
    except OSError: pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Payslip Maker — فیش‌ساز حقوق'); self.geometry('640x600')
        self.cfg = load_settings(); self.v = {}
        pad = dict(padx=8, pady=4)
        f = ttk.Frame(self); f.pack(fill='both', expand=True, padx=12, pady=10)
        ttk.Label(f, text='Excel file / فایل اکسل (xlsx):').grid(row=0, column=0, sticky='w', **pad)
        self.v['file'] = tk.StringVar(value=self.cfg.get('file', ''))
        ttk.Entry(f, textvariable=self.v['file'], width=52).grid(row=1, column=0, sticky='we', **pad)
        ttk.Button(f, text='Browse…', command=self.pick).grid(row=1, column=1, **pad)
        ttk.Label(f, text='Month / ماه:').grid(row=2, column=0, sticky='w', **pad)
        self.month = ttk.Combobox(f, values=MONTHS, state='readonly', width=20); self.month.grid(row=3, column=0, sticky='w', **pad)
        rows = [('period', 'Slip title / عنوان فیش (مثلاً مهر ۱۴۰۵)'), ('company', 'Company / نام شرکت'), ('rep', 'Representative / نماینده کارفرما'),
                ('workplace', 'Workplace / محل کار'), ('duration', 'Contract duration / مدت قرارداد'), ('cstart', 'Start date / تاریخ شروع قرارداد (1405/01/01)')]
        r = 4
        for k, lab in rows:
            ttk.Label(f, text=lab).grid(row=r, column=0, sticky='w', **pad)
            self.v[k] = tk.StringVar(value=self.cfg.get(k, 'یک سال' if k == 'duration' else ''))
            ttk.Entry(f, textvariable=self.v[k], width=52).grid(row=r + 1, column=0, sticky='we', **pad); r += 2
        self.pdf = tk.BooleanVar(value=True); ttk.Checkbutton(f, text='One combined PDF (fast) / یک PDF برای همه (هر نفر یک صفحه)', variable=self.pdf).grid(row=r, column=0, sticky='w', **pad)
        self.pdf_each = tk.BooleanVar(value=False); ttk.Checkbutton(f, text='Separate PDF per person (slow, ~3 s each) / PDF جدا برای هر نفر', variable=self.pdf_each).grid(row=r + 1, column=0, sticky='w', **pad); r += 1
        bar = ttk.Frame(f); bar.grid(row=r + 1, column=0, columnspan=2, sticky='w', pady=8)
        for t, c in (('Payslips / فیش‌ها', self.slips), ('Contracts / قراردادها', self.contracts), ('Payment list / لیست پرداخت', self.paylist)):
            ttk.Button(bar, text=t, command=c).pack(side='left', padx=4)
        self.status = tk.StringVar(value=''); ttk.Label(f, textvariable=self.status, wraplength=600, foreground='#0a6b6b').grid(row=r + 2, column=0, columnspan=2, sticky='w', **pad)
        f.columnconfigure(0, weight=1)
        self.on_file()

    def pick(self):
        p = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        if p: self.v['file'].set(p); self.on_file()

    def on_file(self):
        p = self.v['file'].get()
        cur = ''
        if p and os.path.exists(p):
            try:
                wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
                cur = str(wb['تنظیمات']['B19'].value or ''); wb.close()
            except Exception: pass
        self.month.set(cur if cur in MONTHS else self.cfg.get('month', 'مهر'))
        if not self.v['period'].get() and self.month.get(): self.v['period'].set(self.month.get() + ' ۱۴۰۵')

    def args(self):
        s = {k: v.get().strip() for k, v in self.v.items()}
        s['month'] = self.month.get(); save_settings(s)
        if not os.path.exists(s['file']): raise ValueError('فایل اکسل انتخاب نشده / Excel file not selected')
        if not s['month']: raise ValueError('ماه انتخاب نشده / month not selected')
        out = os.path.join(os.path.dirname(os.path.abspath(s['file'])), 'خروجی-' + s['month'])
        os.makedirs(out, exist_ok=True)
        return s, out, SimpleNamespace(company=s['company'], rep=s['rep'], workplace=s['workplace'], duration=s['duration'] or 'یک سال', contract_start=s['cstart'])

    def run(self, fn):
        try:
            self.status.set('Working… / در حال ساخت…'); self.update_idletasks()
            msg = fn(); self.status.set(msg); messagebox.showinfo('Done / انجام شد', msg)
        except Exception as e:
            traceback.print_exc(); self.status.set('Error: ' + str(e)); messagebox.showerror('Error / خطا', str(e))

    def make_pdfs(self, out, files, name):
        """از روی فایل‌های HTML: یک PDF مشترک و/یا PDF جدا. خروجی: توضیح کوتاه."""
        if not files or not (self.pdf.get() or self.pdf_each.get()): return ''
        br = P.find_browser()
        if not br: return '\nPDF ساخته نشد: Edge یا Chrome پیدا نشد (فایل‌های HTML ساخته شد).'
        note = ''
        if self.pdf.get():
            htmls = [open(f, encoding='utf-8').read() for f in files]
            allf = os.path.join(out, name + '_all.html'); open(allf, 'w', encoding='utf-8').write(P.combine(htmls))
            ok = P.edge_pdf(br, allf, allf[:-5] + '.pdf', timeout=90)
            note += f'\nPDF مشترک: {name}_all.pdf' if ok else '\nPDF مشترک ساخته نشد (فایل HTML آن ساخته شد: آن را در مرورگر باز کنید و Ctrl+P بزنید).'
        if self.pdf_each.get():
            done = P.to_pdf(files); note += f'\nPDF جدا: {len(done)} از {len(files)}'
        return note

    def slips(self):
        def go():
            s, out, a = self.args(); people = P.load(s['file'], s['month']); files = []
            for p in people:
                if p['status'] not in ('پرداخت', 'حق مدرک'): continue
                fn = os.path.join(out, f"slip_{p['row']:03d}.html")
                open(fn, 'w', encoding='utf-8').write(P.build_html(p, s['period'], s['company'], 'actual', True)); files.append(fn)
            extra = self.make_pdfs(out, files, 'slips')
            return f'{len(files)} فیش ساخته شد در:\n{out}' + extra
        self.run(go)

    def contracts(self):
        def go():
            s, out, a = self.args(); people = P.load(s['file'], s['month']); files = []
            for p in people:
                if p['status'] == 'قطع همکاری': continue
                fn = os.path.join(out, f"contract_{p['row']:03d}.html")
                open(fn, 'w', encoding='utf-8').write(P.build_contract(p, a, True)); files.append(fn)
            extra = self.make_pdfs(out, files, 'contracts')
            return f'{len(files)} قرارداد ساخته شد در:\n{out}' + extra
        self.run(go)

    def paylist(self):
        def go():
            s, out, a = self.args(); people = P.load(s['file'], s['month'])
            fn = os.path.join(out, 'payment_list.csv')
            with open(fn, 'w', newline='', encoding='utf-8-sig') as fh:
                w = csv.writer(fh); w.writerow(['ردیف', 'کد', 'نام', 'شبا/حساب', 'مبلغ (ریال)', 'وضعیت'])
                n = 0
                for p in people:
                    if p['status'] == 'پرداخت': w.writerow([p['row'], p['id'], p['full'], p['iban'], int(round(p['pay'])), p['status']]); n += 1
            return f'لیست پرداخت ({n} نفر) ساخته شد:\n{fn}'
        self.run(go)


if __name__ == '__main__':
    App().mainloop()
