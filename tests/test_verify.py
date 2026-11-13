#!/usr/bin/env python3
"""
出處核對（build_report.verify_report）的測試。資料全是虛構的。

用法：python tests/test_verify.py
全部通過會印「全部通過」；有一項不對就會停下來並說明是哪一項。
"""
import copy
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "email-case-file" / "scripts"))
from build_report import verify_report  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

EMAILS = {
    "emails": [
        {"id": "E001", "date_display": "2026-09-01 10:00", "sender": "Amy Chen <amy@example.com>",
         "subject": "A 機種 外觀方案", "body": "附上兩個配色，請選擇。\n模具費我方提議 US$ 24,000。",
         "quoted_excerpt": "", "body_truncated": False},
        {"id": "E002", "date_display": "2026-09-05 14:32", "sender": "John Smith <john@example.net>",
         "subject": "RE: A 機種 外觀方案",
         "body": "Hi Amy,\n\nOK, please go ahead\nwith option B.  We can’t accept 24,000; our number is 36.500.\n\nJohn",
         "quoted_excerpt": "附上兩個配色，請選擇。", "body_truncated": False},
        {"id": "E003", "date_display": "2026-09-08 09:00", "sender": "John Smith <john@example.net>",
         "subject": "Signed quotation", "body": "As attached.", "quoted_excerpt": "", "body_truncated": False,
         "attachment_details": [{"name": "quote_v2.pdf", "status": "ok", "text_file": "att/quote_v2.pdf.txt",
                                 "excerpt": ""}]},
        {"id": "E004", "date_display": "2026-09-09 09:00", "sender": "Amy Chen <amy@example.com>",
         "subject": "很長的信", "body": "開頭而已", "quoted_excerpt": "", "body_truncated": True},
    ],
    "documents": [{"id": "D001", "name": "會議記錄.pdf", "attachment_details": []}],
}


def decision(quote, sid="E002", date="2026-09-05", **src):
    return {"date": date, "decider": "John Smith", "decision": "外觀用 B 方案", "quote": quote,
            "confidence": "高", "source": {"id": sid, **src}}


def run(decisions=(), todos=(), timeline=(), figures=()):
    report = {"topics": [{"name": "A 機種", "decisions": list(decisions), "todos": list(todos),
                          "timeline": list(timeline)}], "figures": list(figures)}
    with tempfile.TemporaryDirectory() as wd:
        (Path(wd) / "att").mkdir()
        (Path(wd) / "att" / "quote_v2.pdf.txt").write_text("Unit price USD 125\nConfirmed and signed by John Smith",
                                                           encoding="utf-8")
        return verify_report(report, copy.deepcopy(EMAILS), wd), report


def check(name, cond):
    if not cond:
        print(f"失敗：{name}")
        sys.exit(1)
    print(f"通過：{name}")


v, _ = run([decision("OK, please go ahead with option B.")])
check("原文一字不差（只差換行）→ 通過", not v["errors"] and v["quotes_ok"] == 1)

v, _ = run([decision("“ok, please go ahead with option b”")])
check("引號樣式、大小寫不同 → 通過", not v["errors"] and v["quotes_ok"] == 1)

v, _ = run([decision("OK, please go ahead … our number is 36.500")])
check("中間用「…」省略 → 通過", not v["errors"] and v["quotes_ok"] == 1)

v, _ = run([decision("We can't accept 24,000")])
check("直引號對彎引號 → 通過", not v["errors"] and v["quotes_ok"] == 1)

v, _ = run([decision("OK, please proceed with option B.")])
check("原文被改寫 → 擋下", len(v["errors"]) == 1 and "找不到" in v["errors"][0])

v, _ = run([decision("OK, please go ahead with option B.", sid="E001", date="2026-09-01")])
check("編號寫錯 → 擋下，並提示真正在哪一封", len(v["errors"]) == 1 and "E002" in v["errors"][0])

v, _ = run([decision("OK", sid="E099")])
check("編號不存在 → 擋下", len(v["errors"]) == 1 and "不存在" in v["errors"][0])

v, _ = run([decision("附上兩個配色，請選擇。")])
check("原文其實在引用的舊信裡 → 提醒，不擋", not v["errors"] and any("引用的舊信" in n for n in v["notes"]))

v, _ = run([decision("這句在後半段", sid="E004", date="2026-09-09")])
check("信太長被截斷 → 提醒無法核對，不擋", not v["errors"] and any("截斷" in n for n in v["notes"]))

v, _ = run([decision("Confirmed and signed by John Smith", sid="E003", date="2026-09-08", attachment="quote_v2.pdf")])
check("原文在附件裡 → 通過", not v["errors"] and v["quotes_ok"] == 1)

v, _ = run([decision("Unit price USD 130", sid="E003", date="2026-09-08", attachment="quote_v2.pdf")])
check("附件裡沒有這句 → 擋下", len(v["errors"]) == 1)

v, _ = run([decision("主管同意付一半", sid="D001", date="2026-09-02")])
check("檔案內容沒讀到 → 提醒無法核對，不擋", not v["errors"] and any("無法核對" in n for n in v["notes"]))

wrong = decision("OK, please go ahead with option B.")
wrong["source"]["date"] = "2026-09-15 14:32"
v, rep = run([wrong])
check("出處日期寫錯 → 自動改成信件的日期",
      len(v["fixed"]) == 1 and rep["topics"][0]["decisions"][0]["source"]["date"] == "2026-09-05 14:32")

v, _ = run([decision("OK, please go ahead with option B.", date="2026-09-20")])
check("決策日期和寄件日差很多 → 提醒", any("不同" in n for n in v["notes"]))

v, _ = run(todos=[{"item": "寄樣品", "source": {"id": "E050"}}, {"item": "沒出處"}])
check("待辦的出處不存在或沒填 → 擋下", len(v["errors"]) == 2)

v, rep = run(timeline=[{"id": "E002", "date": "2026-09-06 08:00"}, {"id": "E077", "date": "2026-09-07"}])
check("時間序：日期自動更正、不存在的編號擋下",
      rep["topics"][0]["timeline"][0]["date"] == "2026-09-05 14:32" and len(v["errors"]) == 1)

fig = lambda value, sid: {"name": "模具費", "versions": [{"date": "2026-09-05", "value": value, "source": {"id": sid}}]}
v, _ = run(figures=[fig("36,500", "E002"), fig("24,000", "E001")])
check("數字寫法不同（36,500 對 36.500）→ 通過", not v["errors"] and not v["notes"])

v, _ = run(figures=[fig("38,500", "E002")])
check("數字在信裡找不到 → 提醒", any("找不到" in n for n in v["notes"]))

print("全部通過")
