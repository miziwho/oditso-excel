# -*- coding: utf-8 -*-
"""릴리스 내려받기 수를 하루에 한 번 적어 둔다.

GitHub 은 누적 숫자만 주고 날짜별 내역을 주지 않는다. 그래서 매일 찍어
두고, 이웃한 두 날의 차이로 '그날 몇 개'를 본다. 과거는 되살릴 수 없으니
빠뜨린 날은 그냥 빈 칸으로 남는다.
"""
import json, os, sys, urllib.request, urllib.error
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE = os.path.join(os.path.dirname(__file__), '..', 'stats')
OUT = os.path.join(BASE, 'downloads.json')
TABLE = os.path.join(BASE, '일별.md')

REPOS = {
    'excel': 'miziwho/oditso-excel',
    'window-switcher': 'miziwho/oditso-window-switcher',
}


def api(path):
    req = urllib.request.Request('https://api.github.com' + path)
    req.add_header('Accept', 'application/vnd.github+json')
    token = os.environ.get('GITHUB_TOKEN')
    if token:
        req.add_header('Authorization', 'Bearer ' + token)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def snapshot(repo):
    """저장소 하나의 자산별 내려받기 수. 못 읽으면 None."""
    try:
        releases = api('/repos/%s/releases?per_page=100' % repo)
    except urllib.error.HTTPError as e:
        print('  %s 를 못 읽음 (%s). 건너뛴다.' % (repo, e.code))
        return None
    assets, en, ko = {}, 0, 0
    for rel in releases:
        if rel.get('draft'):
            continue
        tag = rel['tag_name']
        is_en = tag.lower().endswith('-en')
        for a in rel.get('assets', []):
            n = a.get('download_count', 0)
            assets['%s/%s' % (tag, a['name'])] = n
            if is_en:
                en += n
            else:
                ko += n
    return {'total': en + ko, 'en': en, 'ko': ko, 'assets': assets}


def write_table(days):
    """사람이 읽을 표. 그날치는 어제와의 차이다."""
    lines = ['# 날짜별 내려받기', '',
             '`downloads.json` 을 보기 좋게 옮긴 것이다. 손으로 고치지 말 것 — 매일 다시 쓴다.',
             '',
             '**그날치**는 전날과의 차이다. 기록이 빠진 날이 있으면 그 다음 날에 합쳐져 보이고,',
             '릴리스 자산을 교체한 날은 숫자가 0 으로 돌아가 음수로 보인다.',
             '']
    # 엑셀만 릴리스 파일이 한글판·영문판으로 갈라져 있다. 창전환은 앱 하나에
    # 두 언어가 다 들어 있어 나눌 것이 없다 — 영문 칸은 늘 0, 한글 칸은 합계와
    # 같은 숫자라 열만 차지한다. 그래서 창전환은 합계만 적는다.
    for key, title, split in (('excel', '엑셀 업무도구', True),
                              ('window-switcher', '창전환', False)):
        rows = [d for d in days if key in d]
        if not rows:
            continue
        if split:
            head = '| 날짜 | 합계 | 그날치 | 영문 | 그날치 | 한글 | 그날치 |'
            rule = '| --- | ---: | ---: | ---: | ---: | ---: | ---: |'
        else:
            head = '| 날짜 | 합계 | 그날치 |'
            rule = '| --- | ---: | ---: |'
        lines += ['## %s' % title, '', head, rule]
        prev = None
        for d in rows:
            c = d[key]
            def delta(field):
                if prev is None:
                    return ''
                n = c[field] - prev[field]
                return '+%d' % n if n > 0 else ('%d' % n if n < 0 else '·')
            if split:
                lines.append('| %s | %d | %s | %d | %s | %d | %s |' % (
                    d['date'], c['total'], delta('total'),
                    c['en'], delta('en'), c['ko'], delta('ko')))
            else:
                lines.append('| %s | %d | %s |' % (
                    d['date'], c['total'], delta('total')))
            prev = c
        lines += ['']
    lines += ['마지막 기록: %s' % (days[-1].get('at', '') if days else '')]
    with open(TABLE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')


def main():
    today = datetime.now(KST).strftime('%Y-%m-%d')

    data = {'설명': '하루 한 번 기록한 릴리스 내려받기 누적 수. 이웃한 두 날의 차이가 그날치다.',
            'days': []}
    if os.path.exists(OUT):
        with open(OUT, encoding='utf-8') as f:
            data = json.load(f)
    data.setdefault('days', [])

    row = {'date': today, 'at': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}
    for key, repo in REPOS.items():
        s = snapshot(repo)
        if s is not None:
            row[key] = s
            print('  %-16s 합계 %d (영문 %d / 한글 %d)' % (key, s['total'], s['en'], s['ko']))

    if len(row) == 2:
        print('읽은 게 없다. 아무것도 쓰지 않는다.')
        return 1

    # 같은 날 두 번 돌면 덮어쓴다
    data['days'] = [d for d in data['days'] if d.get('date') != today]
    data['days'].append(row)
    data['days'].sort(key=lambda d: d['date'])
    data['갱신'] = row['at']

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write('\n')
    write_table(data['days'])
    print('%s 기록함. 모두 %d일치.' % (today, len(data['days'])))
    return 0


if __name__ == '__main__':
    sys.exit(main())
