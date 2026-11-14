# -*- coding: utf-8 -*-
"""릴리스 내려받기 수를 하루에 한 번 적어 둔다.

GitHub 은 누적 숫자만 주고 날짜별 내역을 주지 않는다. 그래서 매일 찍어
두고, 이웃한 두 날의 차이로 '그날 몇 개'를 본다. 과거는 되살릴 수 없으니
빠뜨린 날은 그냥 빈 칸으로 남는다.
"""
import json, os, sys, urllib.request, urllib.error
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
OUT = os.path.join(os.path.dirname(__file__), '..', 'stats', 'downloads.json')

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
    print('%s 기록함. 모두 %d일치.' % (today, len(data['days'])))
    return 0


if __name__ == '__main__':
    sys.exit(main())
