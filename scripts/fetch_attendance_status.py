import json
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

ATTENDANCE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyx1csVo37UKOdvZzQD_KGrGjtvWMOM8g3cPte1FQ92lhK7dR5KlvxMaWCLZBbak1k/exec'


def cairo_today_str():
    # مصر بتوقيت UTC+2 ثابت (من غير توقيت صيفي)
    cairo = timezone(timedelta(hours=2))
    now = datetime.now(cairo)
    return now.strftime('%Y-%m-%d')


def fetch_json(url):
    # جوجل بيرفض بعض الطلبات اللي من غير User-Agent شبه المتصفح
    # وبيرجع 404 بدل ما ينفذ الطلب، فبنبعت هيدرز شبه المتصفح.
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                      '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/plain, */*'
    })
    with urllib.request.urlopen(req, timeout=45) as resp:
        return json.loads(resp.read().decode('utf-8'))


def main():
    date_str = cairo_today_str()

    quick_url = ATTENDANCE_SCRIPT_URL + '?action=attendanceQuickStatus&date=' + date_str
    dash_url = ATTENDANCE_SCRIPT_URL + '?action=dashboardData&date=' + date_str

    try:
        quick = fetch_json(quick_url)
    except Exception as e:
        print('ERROR fetching attendanceQuickStatus:', e, file=sys.stderr)
        sys.exit(0)  # لا تفشل الـ workflow، سيبها تحاول تاني بعد 10 دقايق

    if not isinstance(quick, dict) or not quick.get('ok'):
        print('attendanceQuickStatus response not ok, skipping write:', quick, file=sys.stderr)
        sys.exit(0)

    try:
        dash = fetch_json(dash_url)
    except Exception as e:
        print('ERROR fetching dashboardData:', e, file=sys.stderr)
        sys.exit(0)

    if not isinstance(dash, dict) or not dash.get('ok'):
        print('dashboardData response not ok, skipping write:', dash, file=sys.stderr)
        sys.exit(0)

    # كل طلبات الإذن/المأمورية/الإجازة (كل التواريخ) عشان نحسب عدد اللي
    # لسه محتاج اعتماد — البادچ ده بيبان في صفحة الدخول والصفحة الرئيسية
    # وبيتحدث كل شوية، فتحويله لملف ثابت بيلغي الحاجة لضرب السيرفر كل
    # مرة.
    requests_url = ATTENDANCE_SCRIPT_URL + '?action=requestsReport'
    try:
        reqs = fetch_json(requests_url)
    except Exception as e:
        print('ERROR fetching requestsReport:', e, file=sys.stderr)
        sys.exit(0)

    if not isinstance(reqs, dict) or not reqs.get('ok'):
        print('requestsReport response not ok, skipping write:', reqs, file=sys.stderr)
        sys.exit(0)

    pending_count = sum(
        1 for r in reqs.get('rows', [])
        if r.get('status') not in ('موافق', 'مرفوض')
    )

    # قايمة الموظفين (الأسماء + الأرقام السرية + الأيقونات المخفية لكل
    # واحد + تاريخ الميلاد) — بتتحمل من السيرفر أول ما التطبيق يفتح وده
    # كان بطيء أحياناً. البيانات دي أصلاً عمومية وغير محمية (نفس اللي
    # بيوصله أي حد بيفتح رابط السكريبت مباشرة)، فتخزينها في ملف ثابت مش
    # بيكشف حاجة جديدة.
    employees_url = ATTENDANCE_SCRIPT_URL + '?action=employeesList'
    try:
        emps = fetch_json(employees_url)
    except Exception as e:
        print('ERROR fetching employeesList:', e, file=sys.stderr)
        sys.exit(0)

    if not isinstance(emps, dict) or not emps.get('ok'):
        print('employeesList response not ok, skipping write:', emps, file=sys.stderr)
        sys.exit(0)

    out = {
        'ok': True,
        'date': date_str,
        'updatedAt': datetime.now(timezone.utc).isoformat(),
        'absent': quick.get('absent', []),
        'late': quick.get('late', []),
        'dashboardRows': dash.get('rows', []),
        'dashboardRequestRows': dash.get('requestRows', []),
        'dashboardOfficialOut': dash.get('officialOut', {}),
        'pendingRequestsCount': pending_count,
        'employees': emps.get('employees', [])
    }

    with open('attendanceStatus.json', 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write('\n')

    print('Wrote attendanceStatus.json for', date_str)


if __name__ == '__main__':
    main()
