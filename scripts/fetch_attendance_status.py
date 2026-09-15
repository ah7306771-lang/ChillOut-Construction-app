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
        return resp.read().decode('utf-8')


def main():
    date_str = cairo_today_str()
    url = ATTENDANCE_SCRIPT_URL + '?action=attendanceQuickStatus&date=' + date_str
    try:
        raw = fetch_json(url)
        data = json.loads(raw)
    except Exception as e:
        print('ERROR fetching attendance status:', e, file=sys.stderr)
        sys.exit(0)  # لا تفشل الـ workflow، سيبها تحاول تاني بعد 10 دقايق

    if not isinstance(data, dict) or not data.get('ok'):
        print('Response not ok, skipping write:', data, file=sys.stderr)
        sys.exit(0)

    out = {
        'ok': True,
        'date': date_str,
        'updatedAt': datetime.now(timezone.utc).isoformat(),
        'absent': data.get('absent', []),
        'late': data.get('late', [])
    }

    with open('attendanceStatus.json', 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write('\n')

    print('Wrote attendanceStatus.json for', date_str)


if __name__ == '__main__':
    main()
