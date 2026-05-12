# news_monitor.py
# 설치: pip install requests jinja2 schedule openpyxl

import requests
import json
import schedule
import time
import os
from datetime import datetime
from jinja2 import Template
import openpyxl

# ─── 설정 ────────────────────────────────────────────────────
NAVER_CLIENT_ID     = "YOUR_CLIENT_ID"       # 네이버 API ID
NAVER_CLIENT_SECRET = "YOUR_CLIENT_SECRET"   # 네이버 API Secret
KEYWORD             = "정원오"
OUTPUT_HTML         = "report.html"          # 생성될 HTML 파일명
OUTPUT_EXCEL        = "archive.xlsx"         # 엑셀 아카이브 파일
SEARCH_DISPLAY      = 20                     # 한 번에 가져올 기사 수
# ─────────────────────────────────────────────────────────────

archive = []   # 메모리 내 아카이브 (필요시 DB로 교체)

def search_naver_news(keyword, display=20):
    """네이버 뉴스 API 검색"""
    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id":     NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    params = {
        "query":   keyword,
        "display": display,
        "sort":    "date",   # 최신순
    }
    res = requests.get(url, headers=headers, params=params)
    if res.status_code == 200:
        return res.json().get("items", [])
    else:
        print(f"[ERROR] API 오류: {res.status_code}")
        return []

def clean_html_tags(text):
    """네이버 API 결과의 HTML 태그 제거"""
    import re
    return re.sub(r'<[^>]+>', '', text)

def group_by_media(articles):
    """언론사별로 기사 그룹화"""
    from collections import defaultdict
    groups = defaultdict(list)
    for a in articles:
        groups[a["media"]].append(a)
    return dict(groups)

def save_to_excel(articles, timestamp):
    """엑셀 아카이브 저장"""
    if not os.path.exists(OUTPUT_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "아카이브"
        ws.append(["수집시간", "언론사", "제목", "날짜", "링크", "요약"])
        wb.save(OUTPUT_EXCEL)

    wb = openpyxl.load_workbook(OUTPUT_EXCEL)
    ws = wb.active
    for a in articles:
        ws.append([
            timestamp,
            a.get("media", ""),
            a.get("title", ""),
            a.get("pubDate", ""),
            a.get("link", ""),
            a.get("description", ""),
        ])
    wb.save(OUTPUT_EXCEL)
    print(f"[EXCEL] {len(articles)}건 아카이브 저장 완료")

def generate_html(articles, timestamp):
    """정적 HTML 파일 생성"""
    grouped = group_by_media(articles)
    media_counts = {m: len(v) for m, v in grouped.items()}

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>정원오 뉴스 모니터 - {timestamp}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@400;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  /* (위 HTML의 CSS와 동일 — 생략 없이 그대로 포함됩니다) */
  :root {{
    --bg:#0d0f11;--surface:#151920;--border:#1f2730;
    --accent:#c8a96e;--accent2:#4a9eff;--text:#e2ddd6;
    --muted:#6b7585;--danger:#e05c5c;--green:#4ecb71;
  }}
  *{{margin:0;padding:0;box-sizing:border-box;}}
  body{{background:var(--bg);color:var(--text);font-family:'Noto Serif KR',serif;}}
  .header{{background:var(--surface);border-bottom:1px solid var(--border);padding:14px 24px;
    display:flex;justify-content:space-between;align-items:center;}}
  .logo{{font-family:'JetBrains Mono',monospace;font-size:13px;color:var(--accent);}}
  .meta{{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted);}}
  .hero{{border-bottom:2px solid var(--accent);padding:32px 24px;background:var(--surface);}}
  .hero h1{{font-size:36px;font-weight:700;}}
  .hero h1 em{{color:var(--accent);font-style:normal;}}
  .hero-sub{{font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--muted);margin-top:6px;}}
  .stats{{display:flex;gap:0;border:1px solid var(--border);margin-top:20px;display:inline-flex;}}
  .stat{{padding:12px 24px;border-right:1px solid var(--border);text-align:center;}}
  .stat:last-child{{border-right:none;}}
  .stat-n{{font-family:'JetBrains Mono',monospace;font-size:24px;color:var(--accent);display:block;}}
  .stat-l{{font-size:11px;color:var(--muted);font-family:'JetBrains Mono',monospace;}}
  .layout{{display:grid;grid-template-columns:200px 1fr;gap:24px;padding:24px;max-width:1100px;margin:0 auto;}}
  .sidebar{{}}
  .sidebar-block{{background:var(--surface);border:1px solid var(--border);margin-bottom:16px;}}
  .shead{{padding:8px 12px;background:var(--border);font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted);letter-spacing:.1em;}}
  .mlist{{list-style:none;}}
  .mitem{{display:flex;justify-content:space-between;padding:8px 12px;border-bottom:1px solid var(--border);font-size:12px;}}
  .mitem:last-child{{border-bottom:none;}}
  .mcnt{{font-family:'JetBrains Mono',monospace;font-size:11px;background:var(--border);padding:1px 6px;color:var(--accent);}}
  .main{{display:flex;flex-direction:column;gap:14px;}}
  .sdiv{{display:flex;align-items:center;gap:10px;padding:4px 0;}}
  .sdiv span{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;white-space:nowrap;}}
  .sdiv-line{{flex:1;height:1px;background:var(--border);}}
  .card{{background:var(--surface);border:1px solid var(--border);border-left:3px solid var(--border);padding:18px 20px;}}
  .card:hover{{border-left-color:var(--accent);}}
  .cmeta{{display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap;}}
  .mtag{{background:var(--border);font-family:'JetBrains Mono',monospace;font-size:10px;padding:2px 7px;color:var(--accent);}}
  .dtag{{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted);}}
  .ctitle{{font-size:15px;font-weight:600;margin-bottom:10px;}}
  .ctitle a{{color:var(--text);text-decoration:none;}}
  .ctitle a:hover{{color:var(--accent);}}
  .cdesc{{font-size:12px;color:var(--muted);line-height:1.7;border-left:2px solid var(--border);padding-left:10px;margin-bottom:12px;}}
  .cfoot{{display:flex;justify-content:flex-end;}}
  .rbtn{{font-family:'JetBrains Mono',monospace;font-size:11px;border:1px solid var(--border);
    color:var(--muted);padding:4px 12px;text-decoration:none;}}
  .rbtn:hover{{border-color:var(--accent);color:var(--accent);}}
  footer{{border-top:1px solid var(--border);padding:18px;text-align:center;
    font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted);}}
</style>
</head>
<body>
<div class="header">
  <div class="logo">NEWS MONITOR / 정원오 REPORT</div>
  <div class="meta">생성시각: {timestamp}</div>
</div>
<div class="hero" style="max-width:1100px;margin:0 auto;padding:32px 24px;">
  <h1>정<em>원오</em> 뉴스 모니터링</h1>
  <div class="hero-sub">키워드: 정원오 &nbsp;|&nbsp; 생성: {timestamp}</div>
  <div class="stats">
    <div class="stat"><span class="stat-n">{len(articles)}</span><span class="stat-l">총 기사</span></div>
    <div class="stat"><span class="stat-n">{len(grouped)}</span><span class="stat-l">언론사</span></div>
  </div>
</div>
<div class="layout">
  <aside>
    <div class="sidebar-block">
      <div class="shead">언론사별 건수</div>
      <ul class="mlist">
        <li class="mitem"><span>전체</span><span class="mcnt">{len(articles)}</span></li>
"""
    for media, cnt in media_counts.items():
        html += f'        <li class="mitem"><span>{media}</span><span class="mcnt">{cnt}</span></li>\n'

    html += """      </ul>
    </div>
  </aside>
  <main class="main">
"""
    for media, items in grouped.items():
        html += f"""    <div class="sdiv"><div class="sdiv-line"></div><span>{media}</span><div class="sdiv-line"></div></div>\n"""
        for a in items:
            html += f"""    <div class="card">
      <div class="cmeta">
        <span class="mtag">{media}</span>
        <span class="dtag">{a.get('pubDate','')}</span>
      </div>
      <div class="ctitle"><a href="{a.get('link','#')}" target="_blank" rel="noopener">{a.get('title','')}</a></div>
      <div class="cdesc">{a.get('description','')}</div>
      <div class="cfoot"><a class="rbtn" href="{a.get('link','#')}" target="_blank" rel="noopener">원문 보기 →</a></div>
    </div>\n"""

    html += f"""  </main>
</div>
<footer>Generated by NEWS MONITOR &mdash; keyword: 정원오 &mdash; {timestamp}</footer>
</body>
</html>"""

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[HTML] {OUTPUT_HTML} 생성 완료")

def run_job():
    """메인 수집 작업"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n[START] {timestamp} - '{KEYWORD}' 검색 시작")

    raw = search_naver_news(KEYWORD, SEARCH_DISPLAY)
    articles = []
    for item in raw:
        articles.append({
            "title":       clean_html_tags(item.get("title", "")),
            "link":        item.get("link", ""),
            "description": clean_html_tags(item.get("description", "")),
            "pubDate":     item.get("pubDate", ""),
            "media":       item.get("originallink", "").split("/")[2].replace("www.", "")
                           if item.get("originallink") else "기타",
        })

    archive.extend(articles)
    generate_html(articles, timestamp)
    save_to_excel(articles, timestamp)
    print(f"[DONE] {len(articles)}건 처리 완료\n")

# ─── 스케줄 설정 ──────────────────────────────────────────────
if __name__ == "__main__":
    # 원하는 시간에 실행 (여러 개 추가 가능)
    schedule.every().day.at("09:00").do(run_job)
    schedule.every().day.at("11:00").do(run_job)
    schedule.every().day.at("13:00").do(run_job)
    schedule.every().day.at("15:00").do(run_job)

    print("== NEWS MONITOR 시작 ==")
    print(f"키워드: {KEYWORD}")
    print("스케줄: 09:00 / 11:00 / 13:00 / 15:00")
    print("Ctrl+C 로 종료\n")

    run_job()   # 시작하자마자 1회 즉시 실행

    while True:
        schedule.run_pending()
        time.sleep(30)
