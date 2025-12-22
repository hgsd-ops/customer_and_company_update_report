# ======================================================
# Company Updates – Monthly Report (Standalone + PDF)
# ======================================================

import requests
import json
import re
import shutil
import subprocess
import tempfile
import webbrowser
from datetime import datetime, timedelta, timezone
from pathlib import Path
from collections import Counter
import html as html_lib
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
import calendar

# ------------------------------------------------------
# CONFIG
# ------------------------------------------------------

MONDAY_API_KEY = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQ3MDA3NDM5MywiYWFpIjoxMSwidWlkIjoyNTQ3ODk2NiwiaWFkIjoiMjAyNS0wMi0xMFQwOTowMToxMi4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjQ2NTk0NSwicmduIjoidXNlMSJ9.bVWyZPs_iK2IFftxN6qaMIxJIZCR87EPDM8NGesluRI"
MONDAY_TEAM_URL = "https://abyssworkshop.monday.com"
COMPANIES_BOARD_ID = 3401154685
MONDAY_API_URL = "https://api.monday.com/v2"

DEFAULT_AVATAR = "https://cdn-icons-png.flaticon.com/512/149/149071.png"

# ---- Keyword groups (English + Norwegian) ----
KEYWORD_GROUPS = {
    "offer": {
        "words": ["offer", "tilbud"],
        "color": "#fff3bf",
    },
    "meeting": {
        "words": ["meeting", "møte", "forhandling"],
        "color": "#e3f2fd",
    },
    "price": {
        "words": ["price", "pris", "kostnad"],
        "color": "#e8f5e9",
    },
    "contract": {
        "words": ["contract", "kontrakt", "avtale"],
        "color": "#fbe9e7",
    },
}

SCRIPT_DIR = Path(__file__).resolve().parent
PREVIEW_HTML = SCRIPT_DIR / "company_updates_preview.html"

HEADERS = {
    "Authorization": MONDAY_API_KEY,
    "Content-Type": "application/json",
}

# ------------------------------------------------------
# MONTH WINDOW (last 31 days)
# ------------------------------------------------------

NOW_UTC = datetime.now(timezone.utc)

MONTH_END_UTC = NOW_UTC
MONTH_START_UTC = NOW_UTC - timedelta(days=31)

# Format dates for heading
START_DATE_STR = MONTH_START_UTC.strftime("%b %d")
END_DATE_STR = MONTH_END_UTC.strftime("%b %d, %Y")

HEADING_TEXT = f"Company updates – {START_DATE_STR} to {END_DATE_STR}"
PDF_FILENAME = f"Company updates – {START_DATE_STR} to {END_DATE_STR}.pdf"

# ------------------------------------------------------
# GRAPHQL QUERY
# ------------------------------------------------------

QUERY_COMPANY_UPDATES = f"""
{{
  boards(ids: {COMPANIES_BOARD_ID}) {{
    items_page(limit: 500) {{
      items {{
        id
        name
        updates {{
          id
          body
          created_at
          creator {{
            id
            name
            photo_small
          }}
        }}
      }}
    }}
  }}
}}
"""

# ------------------------------------------------------
# TEXT CLEANING + TRUNCATION
# ------------------------------------------------------

def clean_update_text(raw: str) -> str:
    if not raw:
        return ""

    text = raw

    text = re.sub(
        r'<a[^>]+class="user_mention_editor[^"]*"[^>]*>(.*?)</a>',
        r'\1',
        text,
        flags=re.IGNORECASE | re.DOTALL
    )

    text = re.sub(r'<img[^>]*>', '', text, flags=re.IGNORECASE)
    text = re.sub(r'</?(div|p|br|hr)[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html_lib.unescape(text)

    text = re.sub(
        r'(?im)^(from|fra|sent|sendt|to|til|subject|emne):.*$',
        '',
        text
    )

    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def truncate_update_text(text: str, max_lines: int = 14) -> str:
    lines = [l.rstrip() for l in text.splitlines() if l.strip()]

    output = []
    from_blocks_seen = 0

    for line in lines:
        if line.lower().startswith("fra:"):
            from_blocks_seen += 1
            if from_blocks_seen > 1:
                output.append("— tidligere e-postutveksling skjult —")
                break

        output.append(line)

        if len(output) >= max_lines:
            output.append("— teksten er forkortet —")
            break

    return "\n".join(output)

# ------------------------------------------------------
# LINKS
# ------------------------------------------------------

def strip_tracking_params(url: str) -> str:
    try:
        parsed = urlparse(url)
        query = parse_qs(parsed.query)

        cleaned_query = {
            k: v for k, v in query.items()
            if not (
                k.lower().startswith("utm_") or
                k.lower() in ("fbclid", "gclid", "mc_cid", "mc_eid")
            )
        }

        return urlunparse(parsed._replace(
            query=urlencode(cleaned_query, doseq=True)
        ))
    except Exception:
        return url


def extract_links(text: str):
    urls = re.findall(r'https?://\S+', text)
    return [strip_tracking_params(u) for u in urls]

# ------------------------------------------------------
# KEYWORD HIGHLIGHTING (LANGUAGE-AWARE)
# ------------------------------------------------------

def highlight_keywords(text: str) -> str:
    escaped = html_lib.escape(text)

    for group in KEYWORD_GROUPS.values():
        words = group["words"]
        color = group["color"]

        pattern = r"\b(" + "|".join(map(re.escape, words)) + r")\b"

        escaped = re.sub(
            pattern,
            lambda m: (
                f"<span style='background:{color};"
                f"padding:1px 4px;border-radius:4px;'>"
                f"{m.group(0)}</span>"
            ),
            escaped,
            flags=re.IGNORECASE
        )

    return escaped

# ------------------------------------------------------
# FETCH UPDATES
# ------------------------------------------------------

def fetch_company_updates():
    r = requests.post(
        MONDAY_API_URL,
        json={"query": QUERY_COMPANY_UPDATES},
        headers=HEADERS,
        timeout=30
    ).json()

    items = r["data"]["boards"][0]["items_page"]["items"]
    updates = []

    for item in items:
        company_url = f"{MONDAY_TEAM_URL}/boards/{COMPANIES_BOARD_ID}/pulses/{item['id']}"

        for u in item.get("updates", []):
            try:
                created = datetime.fromisoformat(
                    u["created_at"].replace("Z", "+00:00")
                )
            except Exception:
                continue

            if not (MONTH_START_UTC <= created <= MONTH_END_UTC):
                continue

            creator = u.get("creator") or {}

            updates.append({
                "company": item["name"],
                "company_url": company_url,
                "update_url": f"{company_url}?update_id={u['id']}",
                "text": u.get("body", ""),
                "created": created,
                "user": creator.get("name", "Unknown"),
                "avatar": creator.get("photo_small") or DEFAULT_AVATAR,
            })

    updates.sort(key=lambda x: x["created"], reverse=True)
    return updates

# ------------------------------------------------------
# USER HEATMAP
# ------------------------------------------------------

def render_user_heatmap(updates):
    counts = Counter(u["user"] for u in updates)

    if not counts:
        return "<div style='color:#777;'>No updates this month.</div>"

    html = "<div style='margin-top:6px;'>"
    for user, count in sorted(counts.items(), key=lambda x: x[1], reverse=True):
        intensity = min(1.0, count / 12)
        html += f"""
<div style="margin:4px 0;font-size:13px;">
  <span style="
      display:inline-block;
      width:10px;
      height:10px;
      border-radius:50%;
      background:rgba(46,125,50,{intensity});
      margin-right:6px;"></span>
  {html_lib.escape(user)} — {count}
</div>
"""
    html += "</div>"
    return html

# ------------------------------------------------------
# HTML BUILDER
# ------------------------------------------------------

def build_html(updates):
    heatmap_html = render_user_heatmap(updates)

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{HEADING_TEXT}</title>

<style>
@page {{
  size: A4 portrait;
  margin: 20mm;
}}

body {{
  font-family: Arial, sans-serif;
  max-width: 700px;
  margin: 0 auto;
}}

.lead {{
  border:1px solid #ddd;
  border-radius:8px;
  padding:12px;
  margin-bottom:16px;
  break-inside: avoid;
}}

.lead-title {{
  font-size:16px;
  font-weight:700;
}}

.owner-line {{
  display:flex;
  align-items:center;
  margin-top:6px;
  font-size:13px;
}}

.avatar {{
  width:26px;
  height:26px;
  border-radius:50%;
  margin-right:8px;
}}

.meta {{
  font-size:13px;
  margin-top:6px;
  white-space:pre-wrap;
  overflow-wrap:anywhere;
  word-break:break-word;
}}

.button {{
  display:inline-block;
  margin-top:8px;
  padding:6px 10px;
  font-size:12px;
  border-radius:6px;
  background:#0a4b8f;
  color:#fff;
  text-decoration:none;
}}

a {{
  color:#0a4b8f;
  text-decoration:none;
}}
</style>
</head>

<body>

<h1>{HEADING_TEXT}</h1>

<div class="lead">
  <b>User activity</b>
  {heatmap_html}
</div>
"""

    current_company = None

    for u in updates:
        if u["company"] != current_company:
            html += "</div>" if current_company else ""
            html += f"""
<div class="lead">
  <div class="lead-title">{html_lib.escape(u["company"])}</div>
"""
            current_company = u["company"]

        cleaned = clean_update_text(u["text"])
        truncated = truncate_update_text(cleaned)
        highlighted = highlight_keywords(truncated)
        links = extract_links(cleaned)

        links_html = ""
        if links:
            links_html = "<div class='meta'><b>Links:</b><ul>"
            for l in links:
                short = l.split("?")[0]
                links_html += f"<li><a href='{l}'>{html_lib.escape(short)}</a></li>"
            links_html += "</ul></div>"

        created_str = u["created"].strftime("%d %b %Y %H:%M")

        html += f"""
  <div class="owner-line">
    <img src="{u['avatar']}" class="avatar">
    <b>{html_lib.escape(u["user"])}</b>
    <span style="margin-left:8px;color:#777;">{created_str}</span>
  </div>

  <div class="meta">{highlighted}</div>
  {links_html}

  <a class="button" href="{u['update_url']}">Open full update in Monday</a>
"""

    if current_company:
        html += "</div>"

    html += "</body></html>"
    return html

# ------------------------------------------------------
# HTML → PDF
# ------------------------------------------------------

def html_to_pdf(html: str, output_path: Path):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp:
        tmp.write(html.encode("utf-8"))
        tmp_path = Path(tmp.name)

    candidates = [
        shutil.which("msedge"),
        shutil.which("chrome"),
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]

    browser = next((c for c in candidates if c and Path(c).exists()), None)
    if not browser:
        raise RuntimeError("Chrome or Edge not found.")

    cmd = [
        browser,
        "--headless=new",
        "--disable-gpu",
        "--print-to-pdf-no-header",
        "--no-pdf-header-footer",
        "--window-size=800,2000",
        "--virtual-time-budget=20000",
        f"--print-to-pdf={output_path}",
        f"file:///{tmp_path.resolve().as_posix()}",
    ]

    subprocess.run(cmd, check=True)
    tmp_path.unlink(missing_ok=True)

# ------------------------------------------------------
# MAIN
# ------------------------------------------------------

def main():
    updates = fetch_company_updates()

    html = build_html(updates)
    PREVIEW_HTML.write_text(html, encoding="utf-8")
    webbrowser.open(PREVIEW_HTML.as_uri())

    pdf_path = SCRIPT_DIR / PDF_FILENAME
    html_to_pdf(html, pdf_path)

    print(f"Monthly PDF generated: {pdf_path}")

if __name__ == "__main__":
    main()
