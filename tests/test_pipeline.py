#!/usr/bin/env python3
"""
整條流程的測試：用 tests/fixtures/mail 裡的假信件（人名、公司、金額都是虛構的），
實際跑 parse_emails.py → build_report.py → case_query.py，檢查結果。

用法：python tests/test_pipeline.py
改過腳本後跑一次。全部通過會印「全部通過」；有一項不對就停下來並說明是哪一項。
測試只在暫存資料夾裡寫檔，不會動到 fixtures。
"""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "email-case-file" / "scripts"
FIXTURES = ROOT / "tests" / "fixtures" / "mail"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def run(script, *args, expect=0):
    p = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode != expect:
        print(f"失敗：{script} 結束碼 {p.returncode}（預期 {expect}）\n{p.stdout}\n{p.stderr}")
        sys.exit(1)
    return p.stdout


def check(name, cond, detail=""):
    if not cond:
        print(f"失敗：{name}" + (f"\n  {detail}" if detail else ""))
        sys.exit(1)
    print(f"通過：{name}")


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def by_subject(emails, subject, src="top"):
    hits = [e for e in emails if e["subject"] == subject and e["src"] == src]
    assert len(hits) == 1, (subject, src, len(hits))
    return hits[0]


LONG_TAIL = "Final point after a very long mail: we accept 30,000 as the final tooling amount."

tmp = Path(tempfile.mkdtemp(prefix="ecf_test_"))
try:
    mail, work = tmp / "mail", tmp / "work"
    shutil.copytree(FIXTURES, mail)
    # 長信：重點寫在 8,000 字以後
    filler = "\n".join(f"Line {i}: background detail about shipment batch {i}, nothing to decide here." for i in range(200))
    (mail / "07_long.eml").write_text(
        "From: Tom Lee <tom@example.net>\nTo: Amy Chen <amy@example.com>\nSubject: Tooling cost - long summary\n"
        "Date: Thu, 10 Sep 2026 15:00:00 +0800\nMessage-ID: <e7@example.net>\n"
        "Content-Type: text/plain; charset=utf-8\n\nAmy,\n\n" + filler + "\n\n" + LONG_TAIL + "\n\nTom\n", encoding="utf-8")

    # ---------- 1. 解析 ----------
    run("parse_emails.py", mail, "--out", work)
    data = load(work / "emails.json")
    emails = data["emails"]
    ids = [e["id"] for e in emails]

    check("7 個信件檔 → 7 封信（1 封是從轉寄信拆出來的，重複的不算）", len(emails) == 7, str(ids))
    check("編號照寄件時間排", ids == sorted(ids) and [e["date"] for e in emails] == sorted(e["date"] for e in emails))

    first = by_subject(emails, "A model colour options")
    check("同一封信存兩次、又被引用在回信裡 → 只留一份，並記下也出現在哪",
          first["id"] == "E001" and "02_reply.eml" in first["also_in"])

    tom = by_subject(emails, "Tooling cost", src="quote")
    fwd = by_subject(emails, "FW: Tooling cost")
    check("轉寄信裡夾的舊信拆成獨立一封，時間用引用標頭（9/3 16:15）",
          tom["date_display"] == "2026-09-03 16:15" and tom["container"] == fwd["id"] and tom["date_approx"])
    check("拆出來的信只有自己的內容", "PI-0001" in tom["body"] and "FYI" not in tom["body"])
    check("轉寄信本身只留新寫的那幾句", "FYI" in fwd["body"] and "PI-0001" not in fwd["body"])

    reply = by_subject(emails, "RE: A model colour options")
    check("回信只留新寫的內容", "go ahead with option B" in reply["body"] and "two colour options" not in reply["body"])
    check("回信和原信串成同一串", reply["thread"] == first["thread"] and reply["thread_size"] == 2)

    boss = by_subject(emails, "RE: FW: Tooling cost")
    check("中文引用標頭認得、中文內文沒有亂碼", "模具費就照付一半去談" in boss["body"] and "FYI" not in boss["body"])
    check("轉寄、回覆、拆出的舊信串成同一串", len({tom["thread"], fwd["thread"], boss["thread"]}) == 1)

    auto = by_subject(emails, "Automatic reply: A model colour options")
    check("自動回覆標成系統信", auto["system"] and not first["system"])

    long_ = by_subject(emails, "Tooling cost - long summary")
    full = (work / long_["body_file"]).read_text(encoding="utf-8") if long_.get("body_file") else ""
    check("長信：emails.json 只放開頭，全文另存", long_["body_truncated"] and LONG_TAIL not in long_["body"] and LONG_TAIL in full)
    check("emails.txt 會指出全文在哪", "bodies" in (work / "emails.txt").read_text(encoding="utf-8"))

    docs = data["documents"]
    check("資料夾裡的文字檔編成 D001", len(docs) == 1 and docs[0]["id"] == "D001" and docs[0]["file"].endswith(".txt"))

    people = load(work / "people.json")["people"]
    david = [p for p in people if "david@example.com" in json.dumps(p)]
    check("同一個信箱的中英文名字合成一個人", len(david) == 1 and "王大明" in json.dumps(david[0], ensure_ascii=False))

    # ---------- 2. 產出報告 ----------
    report = {
        "title": "測試案", "folder": str(mail), "generated_date": "2026-09-12", "email_count": len(emails),
        "date_range": ["2026-09-01", "2026-09-10"], "overview_summary": "外觀已定 B 方案；模具費還在談。",
        "figures": [{"name": "模具費", "unit": "US$", "versions": [
            {"date": "2026-09-01", "by": "Amy（我方）", "value": "24,000", "status": "提出", "source": {"id": first["id"]}},
            {"date": "2026-09-05", "by": "John", "value": "36,500", "status": "還價", "source": {"id": reply["id"]}}]}],
        "topics": [{
            "name": "A 機種", "email_count": len(emails), "status_summary": "測試。",
            "decisions": [
                {"date": "2026-09-05", "decider": "John Smith", "decision": "外觀用 B 方案。",
                 "quote": "OK, please go ahead with option B.", "confidence": "高", "source": {"id": reply["id"]}},
                {"date": "2026-09-10", "decider": "Tom Lee", "decision": "模具費最後接受 30,000。",
                 "quote": "we accept 30,000 as the final tooling amount", "confidence": "高", "source": {"id": long_["id"]}},
                {"date": "2026-09-02", "decider": "管理層", "decision": "模具費我方最多付一半。",
                 "quote": "管理層決議：模具費我方最多付一半。", "confidence": "中",
                 "source": {"id": "D001", "attachment": "meeting_2026-09-02.txt"}}],
            "pending": [{"issue": "模具費付多少", "raised_by": "Tom", "raised_date": "2026-09-03",
                         "waiting_on": "我方", "confidence": "高", "source": {"id": tom["id"]}}],
            "todos": [{"item": "回覆 Tom 模具費方案", "owner": "Amy", "due": "2026-09-15", "confidence": "高",
                       "source": {"id": boss["id"]}}],
            "completed": [],
            "timeline": [{"id": e["id"], "date": e["date_display"], "sender": e["sender"], "recipients": e["to"],
                          "subject": e["subject"], "headline": e["subject"], "tag": "資訊", "points": ["測試"],
                          "actions": [], "attachments": [], "links": []} for e in emails if not e["system"]],
        }],
    }
    rj = work / "report.json"
    rj.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
    out = run("build_report.py", rj, "--outdir", mail, "--emails", work / "emails.json", "--formats", "docx,html,xlsx,brief,ics")
    check("三條原文都核對到（一般信、8,000 字以後的長信、資料夾裡的檔案）", "原文關鍵句 3 條，其中 3 條" in out, out)
    made = sorted(p.name for p in mail.glob("案卷整理_*"))
    check("Word、網頁、Excel、速覽、行事曆、資料檔都有產出",
          all(any(n.endswith(x) for n in made) for x in (".docx", ".html", "_待辦.xlsx", "_會議速覽.docx", ".ics", ".data.json")), str(made))
    side1 = load(next(mail.glob("案卷整理_*.data.json")))
    keys = {k: [it.get("key") for it in side1["report"]["topics"][0][k]] for k in ("decisions", "pending", "todos")}
    check("每條都拿到固定編號", sorted(keys["decisions"]) == ["DEC-001", "DEC-002", "DEC-003"]
          and keys["pending"] == ["PEN-001"] and keys["todos"] == ["TODO-001"], str(keys))

    bad = json.loads(json.dumps(report))
    bad["topics"][0]["decisions"][0]["quote"] = "OK, we approve option B."
    bj = work / "bad.json"
    bj.write_text(json.dumps(bad, ensure_ascii=False), encoding="utf-8")
    out = run("build_report.py", bj, "--outdir", mail, "--emails", work / "emails.json", expect=2)
    check("原文被改寫 → 不產出報告", "沒有產出報告" in out and len(list(mail.glob("案卷整理_*.data.json"))) == 1)

    # ---------- 3. 快問 ----------
    out = run("case_query.py", mail, "模具")
    check("快問：查「模具」找得到決策、未決、待辦和數字",
          "30,000" in out and "PEN-001" in out and "TODO-001" in out and "36,500" in out and "外觀用 B" not in out, out)
    check("快問：沒有新信時不警告", "⚠" not in out)
    run("case_query.py", work, expect=3)
    print("通過：快問：資料夾沒有資料檔 → 結束碼 3")

    # ---------- 4. 增量 ----------
    (mail / "08_new.eml").write_text(
        "From: Amy Chen <amy@example.com>\nTo: Tom Lee <tom@example.net>\nSubject: RE: Tooling cost - long summary\n"
        "Date: Mon, 14 Sep 2026 10:00:00 +0800\nMessage-ID: <f8@example.com>\nIn-Reply-To: <e7@example.net>\n"
        "Content-Type: text/plain; charset=utf-8\n\nTom,\n\nAgreed. We will pay 30,000. Thank you.\n\nAmy\n", encoding="utf-8")
    out = run("case_query.py", mail)
    check("快問：有新信時會警告答案可能不是最新", "⚠" in out and "1 個信件檔" in out, out)

    work2 = tmp / "work2"
    out = run("parse_emails.py", mail, "--out", work2)
    e2 = load(work2 / "emails.json")["emails"]
    new = [e for e in e2 if e["is_new"]]
    old_ids = {e["uid"]: e["id"] for e in emails}
    check("增量：只有 1 封新信，編號接在最後", len(new) == 1 and new[0]["id"] == "E008", str([e["id"] for e in new]))
    check("增量：舊信編號不變", all(old_ids[e["uid"]] == e["id"] for e in e2 if not e["is_new"]))

    r2 = side1["report"]
    r2["generated_date"] = "2026-09-16"
    t = r2["topics"][0]
    t["todos"][0]["item"] = "回覆 Tom：模具費方案（已口頭說明）"      # 文字改寫，key 還在 → 是同一條
    del t["pending"][0]["key"]                                      # 忘了帶 key、文字沒變 → 照樣接回舊編號
    t["decisions"].append({"date": "2026-09-14", "decider": "Amy Chen", "decision": "我方同意付 30,000。",
                           "quote": "We will pay 30,000.", "confidence": "高", "source": {"id": "E008"}})
    t["todos"].append({"item": "安排 30,000 付款", "owner": "Amy", "due": "", "confidence": "中", "source": {"id": "E008"}})
    t["timeline"].append({"id": "E008", "date": new[0]["date_display"], "sender": "Amy", "recipients": "Tom",
                          "subject": new[0]["subject"], "headline": "Amy 同意付 30,000", "tag": "決策", "points": ["同意"],
                          "actions": [], "attachments": [], "links": []})
    rj2 = work2 / "report.json"
    rj2.write_text(json.dumps(r2, ensure_ascii=False), encoding="utf-8")
    out = run("build_report.py", rj2, "--outdir", mail, "--emails", work2 / "emails.json", "--formats", "html")
    side2 = load(sorted(mail.glob("案卷整理_2026-09-16*.data.json"))[-1])
    t2 = side2["report"]["topics"][0]
    check("增量：改寫過文字的待辦沿用 TODO-001，新待辦拿到 TODO-002",
          sorted(x["key"] for x in t2["todos"]) == ["TODO-001", "TODO-002"])
    check("增量：沒帶 key 但文字沒變的未決，接回 PEN-001", t2["pending"][0]["key"] == "PEN-001")
    check("增量：新決策拿到 DEC-004", sorted(d["key"] for d in t2["decisions"])[-1] == "DEC-004")
    html = next(mail.glob("案卷整理_2026-09-16*.html")).read_text(encoding="utf-8")
    check("增量：變化清單有新決策，而且不會把改寫過的待辦誤報成「已移出」",
          "自上次報告以來" in html and "我方同意付 30,000" in html and "已移出待辦" not in html)
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print("全部通過")
