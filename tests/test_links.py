#!/usr/bin/env python3
"""
連結測試：報告裡每一個指向本機檔案的連結，點下去都要找得到檔案；信裡的網址要抓得完整、乾淨。

刻意製造最容易出錯的情況：
- 資料夾和檔名有中文、空白、#（Office 會把 # 當成「檔案內的位置」而切斷路徑）
- 報告輸出到另一個資料夾（要算相對路徑）
- 用相對路徑執行 parse_emails.py，而且是在別的目錄執行 build_report.py
- report.json 裡手填的 folder 是錯的（連結要以 emails.json 為準）
- 網址後面直接黏中文、包在全形括號裡、Outlook SafeLinks

用法：python tests/test_links.py
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import json
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "email-case-file" / "scripts"
FIXTURES = ROOT / "tests" / "fixtures" / "mail"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def run(script, *args, cwd=None):
    p = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)], cwd=cwd,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode != 0:
        print(f"失敗：{script} 結束碼 {p.returncode}\n{p.stdout}\n{p.stderr}")
        sys.exit(1)
    return p.stdout


def check(name, cond, detail=""):
    if not cond:
        print(f"失敗：{name}" + (f"\n  {detail}" if detail else ""))
        sys.exit(1)
    print(f"通過：{name}")


def office_targets(path):
    """Word／Excel 檔裡所有對外連結（relationship 的 Target）。"""
    out = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith(".rels"):
                xml = z.read(n).decode("utf-8")
                for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?TargetMode="External"', xml):
                    out.append(m.group(1).replace("&amp;", "&"))
                for m in re.finditer(r'<Relationship [^>]*?TargetMode="External"[^>]*?Target="([^"]+)"', xml):
                    out.append(m.group(1).replace("&amp;", "&"))
    return out


def office_resolve(t, base):
    """照 Word／Excel 2013 實測的行為解讀：只認得 %23(#)、%25(%)，中文與空白要原樣。"""
    t = t.replace("%23", "#").replace("%25", "%")
    p = Path(t)
    return p if p.is_absolute() else base / p


LINK_MAIL = """From: Amy Chen <amy@example.com>
To: John Smith <john@example.com>
Subject: 規格連結
Date: Fri, 11 Sep 2026 09:00:00 +0800
Message-ID: <links@example.com>
Content-Type: text/plain; charset=utf-8

John，
規格放這裡 https://drive.google.com/file/d/abc123/view請查收。
報價單（https://example.com/quote_v2.pdf）
https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.dropbox.com%2Fs%2Fq1%2Fquote.xlsx%3Fdl%3D0&data=05%7C01&reserved=0

Amy
"""

tmp = Path(tempfile.mkdtemp(prefix="ecf_link_"))
try:
    mail = tmp / "客戶 X#2026" / "往來信件"
    work, outdir, elsewhere = tmp / "work", tmp / "報告 輸出", tmp / "別的目錄"
    shutil.copytree(FIXTURES, mail / "A機種")
    (mail / "A機種" / "02_reply.eml").rename(mail / "A機種" / "RE_ 打樣 #2 (50%).eml")
    (mail / "A機種" / "08_links.eml").write_text(LINK_MAIL, encoding="utf-8")
    elsewhere.mkdir()

    # 用相對路徑解析（在 tmp 底下執行）
    run("parse_emails.py", Path("客戶 X#2026") / "往來信件", "--out", "work", cwd=tmp)
    data = json.loads((work / "emails.json").read_text(encoding="utf-8"))
    check("emails.json 記的是信件資料夾的絕對路徑", Path(data["folder"]).is_absolute(), data["folder"])

    lm = [e for e in data["emails"] if e["subject"] == "規格連結"][0]
    check("網址：後面黏中文、全形括號、SafeLinks 都抓對",
          lm["links"] == ["https://drive.google.com/file/d/abc123/view", "https://example.com/quote_v2.pdf",
                          "https://www.dropbox.com/s/q1/quote.xlsx?dl=0"], str(lm["links"]))

    emails = [e for e in data["emails"] if not e["system"]]
    report = {
        "title": "連結測試", "folder": "C:/這個路徑是錯的", "generated_date": "2026-09-12",
        "email_count": len(emails), "date_range": ["2026-09-01", "2026-09-11"], "overview_summary": "測試。",
        "topics": [{
            "name": "A 機種", "email_count": len(emails), "status_summary": "測試。",
            "decisions": [], "pending": [], "completed": [],
            "todos": [{"item": f"看 {e['id']}", "owner": "Amy", "due": "2026-09-15", "confidence": "高",
                       "source": {"id": e["id"]}} for e in emails],
            "timeline": [{"id": e["id"], "date": e["date_display"], "sender": e["sender"], "recipients": e["to"],
                          "subject": e["subject"], "headline": e["subject"], "tag": "資訊", "points": ["測試"],
                          "actions": [], "attachments": [], "links": e["links"]} for e in emails],
        }],
    }
    rj = work / "report.json"
    rj.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
    # 在別的目錄執行、用相對路徑指定輸出位置
    out = run("build_report.py", rj, "--outdir", Path("..") / "報告 輸出", "--emails", work / "emails.json",
              "--formats", "docx,html,xlsx,brief", cwd=elsewhere)
    check("沒有「檔案不存在」的警告", "點了會打不開" not in out, out)

    local, web = [], []
    for f in list(outdir.glob("*.docx")) + list(outdir.glob("*.xlsx")):
        for t in office_targets(f):
            (web if t.startswith(("http://", "https://")) else local).append((f.name, t))
    check("Word／Excel 裡有本機連結", len(local) >= len(emails), str(local[:3]))
    bad = [(n, t) for n, t in local if not office_resolve(t, outdir).is_file()]
    check("Word／Excel 的本機連結全都找得到檔案", not bad, str(bad[:5]))
    check("Word／Excel 的本機連結沒有 %E4 這種中文編碼（Office 解不開）",
          not any(re.search(r"%[89A-F][0-9A-F]", t) for _, t in local), str(local[:3]))
    check("Word 報告裡的網址是可點的連結", "https://www.dropbox.com/s/q1/quote.xlsx?dl=0" in [t for _, t in web])

    html = next(outdir.glob("案卷整理_*.html")).read_text(encoding="utf-8")
    hrefs = [urllib.parse.unquote(h.replace("&amp;", "&")) for h in re.findall(r"href='([^']+)'", html)]
    lh = [h for h in hrefs if not h.startswith(("http://", "https://", "#", "mailto:"))]
    check("網頁版有本機連結", len(lh) >= len(emails), str(lh[:3]))
    badh = [h for h in lh if not (outdir / h).is_file()]
    check("網頁版的本機連結全都找得到檔案", not badh, str(badh[:5]))

    # 檔案真的不見時要警告
    (mail / "A機種" / "01_request.eml").unlink()
    out = run("build_report.py", rj, "--outdir", outdir, "--emails", work / "emails.json", "--formats", "docx")
    check("連結指向的檔案不存在 → 印出警告", "點了會打不開" in out and "01_request.eml" in out, out)

    print("全部通過")
finally:
    shutil.rmtree(tmp, ignore_errors=True)
