#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بيحول ملف APP.xlsx (الشيكات / العملاء والموردين / الرواتب / اجمالي متأخرات /
اخر سعر شراء وبيع) لـ 5 ملفات JSON (checks.json / parties.json /
salaries.json / overdue.json / lastPrices.json) بنفس الشكل اللي كود
Apps Script بيتوقعه بالظبط.

الاستخدام:
    python3 xlsx_to_json.py [مسار APP.xlsx] [مجلد الإخراج]
لو من غير أي parameters، بيدور على APP.xlsx في نفس مجلد السكريبت ويحط
الملفات الناتجة جنبه.
"""
import sys
import json
import os
from datetime import datetime, date, timezone
import openpyxl
from openpyxl.utils.datetime import from_excel

SKIP_LABELS = {'الاجمالي', 'الإجمالي', 'الصافي', 'الإجمالى'}
OVERDUE_GRAND_TOTAL_LABEL = 'الإجمالي العام'
OVERDUE_EMPTY_LABEL = 'لا يوجد مستحقات'
# لاحظ: اسم الشيت فيه مسافتين بين "شراء" و"وبيع" بالظبط زي ما هو متسمي
# في APP.xlsx — لو غيّرت الاسم في الشيت لازم تغيّره هنا كمان بنفس الشكل.
LAST_PRICES_SHEET_NAME = 'اخر سعر شراء  وبيع'
INSTALLMENTS_SHEET_NAME = 'الاقساط'
# ترتيب/أسماء أعمدة شيت الأقساط بالظبط زي ما الواجهة (index.html) بتتوقعها:
# index 5 = تاريخ الاستحقاق، index 8 = المتبقّي — الواجهة بتحسب "مدة
# التأخير"/"الحالة" لوحدها من التاريخ، فمش لازم تيجي صح من هنا.
INSTALLMENTS_COLUMNS = [
    'أمر البيع', 'العميل', 'تاريخ التحميل', 'إجمالي الفاتورة', 'القسط',
    'تاريخ الاستحقاق', 'قيمة القسط', 'المسدّد', 'المتبقّي',
    'المتبقّي المتأخر', 'مدة التأخير', 'الحالة',
]


def is_valid_name(v):
    if v is None:
        return False
    s = str(v).strip()
    if not s or s == '0':
        return False
    if s in SKIP_LABELS:
        return False
    return True


def to_number(v):
    if v is None or v == '':
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0


def cell_to_ddmmyyyy(v):
    """بيحول أي خلية تاريخ (Date حقيقي أو رقم تسلسلي إكسل) لنص dd/MM/yyyy،
    أو None لو الخلية فاضية/مش قابلة للتحويل."""
    if v is None or v == '':
        return None
    if isinstance(v, (datetime, date)):
        d = v
    elif isinstance(v, (int, float)):
        try:
            d = from_excel(v)
        except Exception:
            return None
    else:
        # نص جاهز بصيغة معروفة (يوم/شهر/سنة) — نسيبه زي ما هو
        s = str(v).strip()
        return s if s else None
    return d.strftime('%d/%m/%Y')


def find_header_row(rows, required_headers, max_scan=10):
    """بيدوّر على أول صف من أول max_scan صف فيه كل الأسماء المطلوبة، ويرجع
    (رقم الصف index صفري، ديكشنري اسم العمود -> index)."""
    for i, row in enumerate(rows[:max_scan]):
        cells = [str(c).strip() if c is not None else '' for c in row]
        if all(h in cells for h in required_headers):
            col_idx = {}
            for j, c in enumerate(cells):
                if c:
                    col_idx[c] = j
            return i, col_idx
    return -1, {}


def convert_checks(ws):
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(rows, ['العميل', 'تاريخ الاستحقاق'])
    if header_i == -1:
        return []

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return -1

    c_client = col('العميل', 'اسم العميل')
    c_number = col('الرقم')
    c_value = col('القيمة', 'القيمه')
    c_due = col('تاريخ الاستحقاق', 'تاريخه الاستحقاق')
    c_bank = col('البنك', 'البنك المسحوب عليه')
    c_type = col('الموقف', 'النوع', 'نوع الشيك'); c_notes = col('ملاحظات')

    out = []
    for row in rows[header_i + 1:]:
        client = row[c_client] if c_client > -1 and c_client < len(row) else None
        if not is_valid_name(client):
            continue
        due = cell_to_ddmmyyyy(row[c_due]) if c_due > -1 else None
        if not due:
            continue
        out.append({
            'client': str(client).strip(),
            'number': row[c_number] if c_number > -1 else '',
            'value': row[c_value] if c_value > -1 else '',
            'dueDate': due,
            'bank': str(row[c_bank]).strip() if c_bank > -1 and row[c_bank] is not None else '',
            'notes': str(row[c_notes]).strip() if c_notes > -1 and row[c_notes] is not None else '', 'type': str(row[c_type]).strip() if c_type > -1 and c_type < len(row) and row[c_type] is not None else ''
        })
    return out


def convert_parties(ws):
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(rows, ['اسم العميل', 'اسم المورد'])
    if header_i == -1:
        return {'clients': [], 'suppliers': []}

    client_col = cols.get('اسم العميل', -1)
    supplier_col = cols.get('اسم المورد', -1)

    clients, suppliers = [], []
    for row in rows[header_i + 1:]:
        if client_col > -1 and client_col < len(row):
            name = row[client_col]
            if is_valid_name(name):
                debit = to_number(row[client_col + 1]) if client_col + 1 < len(row) else 0
                credit = to_number(row[client_col + 2]) if client_col + 2 < len(row) else 0
                clients.append({'name': str(name).strip(), 'debit': debit, 'credit': credit})
        if supplier_col > -1 and supplier_col < len(row):
            name = row[supplier_col]
            if is_valid_name(name):
                debit = to_number(row[supplier_col + 1]) if supplier_col + 1 < len(row) else 0
                credit = to_number(row[supplier_col + 2]) if supplier_col + 2 < len(row) else 0
                suppliers.append({'name': str(name).strip(), 'debit': debit, 'credit': credit})
    return {'clients': clients, 'suppliers': suppliers}


SALARY_KEYS = [
    ('الاسم', 'name'), ('الاساسي', 'basic'), ('بدل انتقال', 'transport'),
    ('بدل بنزين', 'fuel'), ('مستحق اخر', 'otherDue'), ('اجمالي استحقاق', 'totalDue'),
    ('الاستقطاعات', 'deductions'), ('السلف', 'advances'), ('اجمالي استقطاع', 'totalDeduct'),
    ('صافي المرتب', 'net'), ('الشهر', 'month'), ('السنه', 'year')
]


def convert_salaries(ws):
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(rows, ['الاسم'])
    if header_i == -1:
        return []

    out = []
    for row in rows[header_i + 1:]:
        name_col = cols.get('الاسم', -1)
        name = row[name_col] if name_col > -1 and name_col < len(row) else None
        if not is_valid_name(name):
            continue
        rec = {}
        for header, key in SALARY_KEYS:
            idx = cols.get(header, -1)
            val = row[idx] if idx > -1 and idx < len(row) else None
            if key == 'name':
                rec[key] = str(name).strip()
            elif key in ('month',):
                rec[key] = str(val).strip() if val is not None else ''
            elif key in ('year',):
                rec[key] = val
            else:
                rec[key] = val
        out.append(rec)
    return out


def convert_overdue(ws):
    """بيقرأ شيت 'اجمالي متأخرات' (نسخة طبق الأصل من 'إجمالي التأخيرات'):
    صف عنوان، صف 'الإجمالي العام' (عدد/مبلغ/متوسط تأخير)، صف هيدر
    ('العميل'/'عدد المستحقات'/'إجمالي المبلغ المستحق'/'متوسط التأخير (يوم)')،
    وبعدها صفوف البيانات (أو صف 'لا يوجد مستحقات' لو فاضي)."""
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(
        rows, ['العميل', 'عدد المستحقات', 'إجمالي المبلغ المستحق']
    )

    grand_total = {'count': 0, 'amount': 0, 'avgDelay': 0}
    scan_upto = header_i if header_i > -1 else len(rows)
    for row in rows[:scan_upto]:
        if row and row[0] is not None and str(row[0]).strip() == OVERDUE_GRAND_TOTAL_LABEL:
            grand_total = {
                'count': to_number(row[1]) if len(row) > 1 else 0,
                'amount': to_number(row[2]) if len(row) > 2 else 0,
                'avgDelay': to_number(row[3]) if len(row) > 3 else 0,
            }
            break

    if header_i == -1:
        return {'rows': [], 'grandTotal': grand_total}

    c_client = cols.get('العميل', 0)
    c_count = cols.get('عدد المستحقات', 1)
    c_amount = cols.get('إجمالي المبلغ المستحق', 2)
    c_avg = cols.get('متوسط التأخير (يوم)', 3)

    out = []
    for row in rows[header_i + 1:]:
        name = row[c_client] if c_client < len(row) else None
        if name is None:
            continue
        s = str(name).strip()
        if not s or s in SKIP_LABELS or s == OVERDUE_EMPTY_LABEL:
            continue
        out.append({
            'client': s,
            'count': to_number(row[c_count]) if c_count < len(row) else 0,
            'amount': to_number(row[c_amount]) if c_amount < len(row) else 0,
            'avgDelay': to_number(row[c_avg]) if c_avg < len(row) else 0,
        })
    return {'rows': out, 'grandTotal': grand_total}


def convert_last_prices(ws):
    """بيقرأ شيت 'اخر سعر شراء وبيع' جوه APP.xlsx (بيتملي أوتوماتيك من شيت
    'بيانات التوريد' في بيانات التوريد.xlsm عن طريق ماكرو SyncToApp):
    صف هيدر ('اسم العميل'/'اخر سعر شراء'/'اخر سعر بيع'/'رقم امر البيع'/
    'اخر طريقة السداد'/'تاريخ اخر فاتورة') وتحته صفوف البيانات (عميل واحد
    في كل صف — آخر عملية توريد ليه)."""
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(
        rows, ['اسم العميل', 'اخر سعر شراء', 'اخر سعر بيع']
    )
    if header_i == -1:
        return []

    c_client = cols.get('اسم العميل', 0)
    c_buy = cols.get('اخر سعر شراء', 1)
    c_sell = cols.get('اخر سعر بيع', 2)
    c_order = cols.get('رقم امر البيع', 3)
    c_payment = cols.get('اخر طريقة السداد', 4)
    c_date = cols.get('تاريخ اخر فاتورة', 5)

    out = []
    for row in rows[header_i + 1:]:
        name = row[c_client] if c_client < len(row) else None
        if not is_valid_name(name):
            continue
        out.append({
            'client': str(name).strip(),
            'lastBuyPrice': to_number(row[c_buy]) if c_buy < len(row) else 0,
            'lastSellPrice': to_number(row[c_sell]) if c_sell < len(row) else 0,
            'saleOrderNo': (row[c_order] if c_order < len(row) and row[c_order] is not None else ''),
            'paymentMethod': (str(row[c_payment]).strip() if c_payment < len(row) and row[c_payment] is not None else ''),
            'lastInvoiceDate': cell_to_ddmmyyyy(row[c_date]) if c_date < len(row) else None,
        })
    return out



def convert_all_sales(ws):
    """بيقرأ شيت 'اخر سعر شراء وبيع' في APP.xlsx ويجمّع الصفوف حسب (اسم
    العميل + رقم أمر البيع) في سطر واحد لكل أمر — لأن الماكرو بيكتب صف لكل
    صنف (قطر) في نفس أمر البيع، فنجمع الإجماليات ونحتفظ بأول سعر ظهر لكل
    أمر. النتيجة مرتبة تنازليًا حسب التاريخ ثم رقم الأمر، ومفلترة لإخفاء أي
    أمر بدون سعر بيع (زي ما كان في الشاشة الأصلية). بيدعم أسماء الأعمدة
    القديمة والجديدة (يعني اخر سعر شراء / سعر شراء / إلخ) عشان يشتغل قبل
    وبعد أي تغيير في أسماء الهيدر."""
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(rows, ['اسم العميل'])
    if header_i == -1:
        return []

    def col(*names, default=None):
        for n in names:
            if n in cols:
                return cols[n]
        return default

    c_client = col('اسم العميل', default=0)
    c_buy = col('سعر شراء', 'اخر سعر شراء', default=1)
    c_sell = col('سعر بيع', 'اخر سعر بيع', default=2)
    c_order = col('رقم امر البيع', 'رقم أمر البيع', default=3)
    c_payment = col('طريقة السداد', 'اخر طريقة السداد', default=4)
    c_date = col('تاريخ الفاتورة', 'التاريخ', 'تاريخ اخر فاتورة', default=5)
    c_buy_total_col = col('إجمالي سعر الشراء', 'اجمالي سعر الشراء', 'اجمالي سعرالشراء', default=6)
    c_sell_total_col = col('إجمالي سعر البيع', 'اجمالي سعر البيع', 'اجمالي سعرالبيع', default=7)

    # Group by (client + orderNo) — كل مفتاح فيه: totals مُجمَّعة + أول سعر
    # شراء/بيع لقيته + مجموعة الأسعار المختلفة (لو الأمر فيه أصناف بأسعار
    # مختلفة، نعرضهم في العمود مفصولين بـ "/") + تاريخ الأول + طريقة السداد.
    grouped = {}
    order_list = []
    for row in rows[header_i + 1:]:
        name = row[c_client] if c_client < len(row) else None
        if not is_valid_name(name):
            continue
        order_val = row[c_order] if c_order < len(row) else None
        order_str = str(order_val).strip() if order_val is not None else ''
        key = str(name).strip() + '|' + order_str
        if key not in grouped:
            grouped[key] = {
                'date': cell_to_ddmmyyyy(row[c_date]) if c_date < len(row) else None,
                'orderNo': order_val if order_val is not None else '',
                'client': str(name).strip(),
                'buy_prices': [],
                'sell_prices': [],
                'buyTotal': 0.0,
                'sellTotal': 0.0,
                'paymentMethod': '',
            }
            order_list.append(key)
        g = grouped[key]
        # جمع الإجماليات
        bt = to_number(row[c_buy_total_col]) if c_buy_total_col < len(row) else 0
        st = to_number(row[c_sell_total_col]) if c_sell_total_col < len(row) else 0
        g['buyTotal'] += bt
        g['sellTotal'] += st
        # لو الأمر بأكتر من سعر (أصناف مختلفة)، نجمع الأسعار المتميزة
        bp = to_number(row[c_buy]) if c_buy < len(row) else 0
        sp = to_number(row[c_sell]) if c_sell < len(row) else 0
        if bp and bp not in g['buy_prices']:
            g['buy_prices'].append(bp)
        if sp and sp not in g['sell_prices']:
            g['sell_prices'].append(sp)
        # أول طريقة سداد غير فاضية
        if not g['paymentMethod'] and c_payment < len(row) and row[c_payment] is not None:
            pm = str(row[c_payment]).strip()
            if pm:
                g['paymentMethod'] = pm
        # لو التاريخ الأول كان فاضي وفيه تاريخ في الصف الحالي، ناخده
        if not g['date'] and c_date < len(row):
            g['date'] = cell_to_ddmmyyyy(row[c_date])

    # نبني الإخراج مع فلترة الأوامر بدون سعر بيع
    out = []
    for k in order_list:
        g = grouped[k]
        # إخفاء الأمر لو مفيش أي سعر بيع (لا في العمود العادي ولا في الإجمالي)
        if not g['sell_prices'] and not g['sellTotal']:
            continue
        # الأسعار: لو أكتر من واحد نعرضهم كنص "40000 / 41000"
        buy_price = g['buy_prices'][0] if g['buy_prices'] else 0
        sell_price = g['sell_prices'][0] if g['sell_prices'] else 0
        buy_price2 = g['buy_prices'][1] if len(g['buy_prices']) > 1 else 0
        sell_price2 = g['sell_prices'][1] if len(g['sell_prices']) > 1 else 0
        out.append({
            'date': g['date'],
            'orderNo': g['orderNo'],
            'client': g['client'],
            'buyPrice': buy_price,
            'sellPrice': sell_price,
            'buyPrice2': buy_price2,
            'sellPrice2': sell_price2,
            'buyTotal': g['buyTotal'],
            'sellTotal': g['sellTotal'],
            'paymentMethod': g['paymentMethod'],
        })

    # ترتيب تنازلي حسب التاريخ (الأحدث أولاً)، ثم حسب رقم الأمر تنازلي
    def sort_key(o):
        d = o.get('date') or ''
        try:
            p = d.split('/')
            dt = (int(p[2]), int(p[1]), int(p[0]))
        except Exception:
            dt = (0, 0, 0)
        try:
            ord_num = float(o.get('orderNo') or 0)
        except Exception:
            ord_num = 0
        return (dt, ord_num)
    out.sort(key=sort_key, reverse=True)
    return out

def convert_installments(ws):
    """بيقرأ شيت 'الاقساط' ويرجّع نفس شكل {sheet, columns, rows} اللي
    الواجهة (loadInstallmentsData_ في index.html) بتتوقعه: كل صف فيه
    'row' (رقم الصف في الإكسل) و'values' (12 قيمة بنفس ترتيب
    INSTALLMENTS_COLUMNS بالظبط). الواجهة بتحسب مدة التأخير/الحالة/المتبقي
    المتأخر بنفسها من عمود 'تاريخ الاستحقاق' وعمود 'المتبقّي'، فمش محتاجين
    نحسبهم هنا ولا نبعت أي تنسيقات/ألوان."""
    rows = list(ws.iter_rows(values_only=True))
    header_i, cols = find_header_row(rows, ['العميل', 'تاريخ الاستحقاق'])
    if header_i == -1:
        return {'sheet': INSTALLMENTS_SHEET_NAME, 'columns': INSTALLMENTS_COLUMNS, 'rows': []}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return -1

    c_order = col('أمر البيع', 'رقم أمر البيع')
    c_client = col('العميل')
    c_load = col('تاريخ التحميل')
    c_invoice = col('إجمالي الفاتورة')
    c_label = col('القسط')
    c_due = col('تاريخ الاستحقاق')
    c_value = col('قيمة القسط')
    c_paid = col('المسدّد')
    c_remaining = col('المتبقّي')
    c_late_remaining = col('المتبقّي المتأخر')
    c_delay = col('مدة التأخير')
    c_status = col('الحالة')

    def get(row, idx):
        return row[idx] if idx > -1 and idx < len(row) else None

    out = []
    for row_num, row in enumerate(rows[header_i + 1:], start=header_i + 2):
        client = get(row, c_client)
        if not is_valid_name(client):
            continue
        values = [
            get(row, c_order),
            str(client).strip(),
            cell_to_ddmmyyyy(get(row, c_load)),
            to_number(get(row, c_invoice)),
            get(row, c_label),
            cell_to_ddmmyyyy(get(row, c_due)),
            to_number(get(row, c_value)),
            to_number(get(row, c_paid)),
            to_number(get(row, c_remaining)),
            get(row, c_late_remaining),
            get(row, c_delay),
            get(row, c_status),
        ]
        out.append({'row': row_num, 'values': values})
    return {'sheet': INSTALLMENTS_SHEET_NAME, 'columns': INSTALLMENTS_COLUMNS, 'rows': out}


def main():
    xlsx_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), 'APP.xlsx')
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(xlsx_path))

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    checks = convert_checks(wb['الشيكات'])
    parties = convert_parties(wb['العملاء والموردين'])
    salaries = convert_salaries(wb['الرواتب'])
    overdue = convert_overdue(wb['اجمالي متأخرات']) if 'اجمالي متأخرات' in wb.sheetnames else {'rows': [], 'grandTotal': {'count': 0, 'amount': 0, 'avgDelay': 0}}
    last_prices = convert_last_prices(wb[LAST_PRICES_SHEET_NAME]) if LAST_PRICES_SHEET_NAME in wb.sheetnames else []
    all_sales = convert_all_sales(wb[LAST_PRICES_SHEET_NAME]) if LAST_PRICES_SHEET_NAME in wb.sheetnames else []
    installments = convert_installments(wb[INSTALLMENTS_SHEET_NAME]) if INSTALLMENTS_SHEET_NAME in wb.sheetnames else {'sheet': INSTALLMENTS_SHEET_NAME, 'columns': INSTALLMENTS_COLUMNS, 'rows': []}

    with open(os.path.join(out_dir, 'checks.json'), 'w', encoding='utf-8') as f:
        json.dump(checks, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'parties.json'), 'w', encoding='utf-8') as f:
        json.dump(parties, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'salaries.json'), 'w', encoding='utf-8') as f:
        json.dump(salaries, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'overdue.json'), 'w', encoding='utf-8') as f:
        json.dump(overdue, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'lastPrices.json'), 'w', encoding='utf-8') as f:
        json.dump(last_prices, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'allSales.json'), 'w', encoding='utf-8') as f:
        json.dump(all_sales, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'installments.json'), 'w', encoding='utf-8') as f:
        json.dump(installments, f, ensure_ascii=False, indent=2)

    # وقت التحويل ده (مش وقت آخر تعديل في الإكسل نفسه) — بيتقرا في الواجهة
    # عشان اليوزر يعرف بيانات الشيكات/العملاء/الأسعار/المتأخرات دي جايه من
    # امتى بالظبط (آخر مرة اتحول فيها APP.xlsx)، بدل ما يفتكرها لحظية.
    with open(os.path.join(out_dir, 'dataUpdatedAt.json'), 'w', encoding='utf-8') as f:
        json.dump({'updatedAt': datetime.now(timezone.utc).isoformat()}, f, ensure_ascii=False, indent=2)

    print('checks:', len(checks), '| clients:', len(parties['clients']), '| suppliers:', len(parties['suppliers']), '| salaries:', len(salaries), '| overdue rows:', len(overdue['rows']), '| last prices rows:', len(last_prices), '| installments rows:', len(installments['rows']))


if __name__ == '__main__':
    main()
