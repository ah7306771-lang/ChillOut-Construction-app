#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بيحول ملف APP.xlsx (الشيكات / العملاء والموردين / الرواتب) لـ 3 ملفات JSON
(checks.json / parties.json / salaries.json) بنفس الشكل اللي كود Apps Script
بيتوقعه بالظبط.

الاستخدام:
    python3 xlsx_to_json.py [مسار APP.xlsx] [مجلد الإخراج]
لو من غير أي parameters، بيدور على APP.xlsx في نفس مجلد السكريبت ويحط
الملفات الناتجة جنبه.
"""
import sys
import json
import os
from datetime import datetime, date
import openpyxl
from openpyxl.utils.datetime import from_excel

SKIP_LABELS = {'الاجمالي', 'الإجمالي', 'الصافي', 'الإجمالى'}


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


def main():
    xlsx_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), 'APP.xlsx')
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(xlsx_path))

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    checks = convert_checks(wb['الشيكات'])
    parties = convert_parties(wb['العملاء والموردين'])
    salaries = convert_salaries(wb['الرواتب'])

    with open(os.path.join(out_dir, 'checks.json'), 'w', encoding='utf-8') as f:
        json.dump(checks, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'parties.json'), 'w', encoding='utf-8') as f:
        json.dump(parties, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, 'salaries.json'), 'w', encoding='utf-8') as f:
        json.dump(salaries, f, ensure_ascii=False, indent=2)

    print('checks:', len(checks), '| clients:', len(parties['clients']), '| suppliers:', len(parties['suppliers']), '| salaries:', len(salaries))


if __name__ == '__main__':
    main()
