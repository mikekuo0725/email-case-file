#!/usr/bin/env python3
"""
快問：不重新整理信件，直接從上次的案卷資料檔（案卷整理_*.data.json）查答案。

用法：
    python case_query.py <信件資料夾>                  # 現況一覽：摘要、數字、未決、待辦、最近的決策
    python case_query.py <信件資料夾> 模具 tooling     # 只列提到這些字的條目（任一個字符合就列）
    python case_query.py <信件資料夾> --all            # 全部條目，含已完成與每封信的一句話

只讀不寫。資料夾裡沒有資料檔時，結束碼是 3，代表要先跑完整流程。
"""
import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

EMAIL_EXT = (".msg", ".eml", ".pst")


def parse_date(s):
    try:
        return date.fromisoformat((s or "")[:10])
    except ValueError:
        return None


def src(it):
    s = it.get("source") or {}
    bits = [s.get("id", ""), str(s.get("date", ""))[:10], s.get("file", "")]
    if s.get("attachment"):
        bits.append(f"附件《{s['attachment']}》")
    return "　出處：" + " ".join(b for b in bits if b) if s.get("id") else ""


def tags(it):
    out = [it.get("key", "")]
    if it.get("confidence") == "低":
        out.append("待確認")
    return "".join(f"[{t}] " for t in out if t)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("words", nargs="*", help="關鍵字；任一個符合就列出。不給就列現況一覽")
    ap.add_argument("--all", action="store_true", help="列出全部條目（含已完成與時間序）")
    ap.add_argument("--data", help="指定資料檔；預設用資料夾裡最新的 案卷整理_*.data.json")
    args = ap.parse_args()
    try:   # Windows 主控台預設 cp950
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    folder = Path(args.folder)
    if args.data:
        path = Path(args.data)
    else:
        cands = sorted(folder.glob("案卷整理_*.data.json"), key=lambda p: p.stat().st_mtime)
        if not cands:
            print("這個資料夾還沒有案卷資料檔（案卷整理_*.data.json），沒辦法快問。請先跑完整的整理流程。")
            sys.exit(3)
        path = cands[-1]
    r = json.loads(path.read_text(encoding="utf-8")).get("report", {})
    today = date.today()
    gen = parse_date(r.get("generated_date"))

    made = path.stat().st_mtime
    newer = [p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() in EMAIL_EXT and p.stat().st_mtime > made]
    print(f"案卷：{r.get('title', '')}｜資料檔：{path.name}｜整理日期：{r.get('generated_date', '')}"
          + (f"（{(today - gen).days} 天前）" if gen else ""))
    if newer:
        print(f"⚠ 資料夾裡有 {len(newer)} 個信件檔比這份資料新（最新：{datetime.fromtimestamp(max(p.stat().st_mtime for p in newer)):%Y-%m-%d}），"
              "下面的答案不含這些信。要最新狀況請重新整理（增量模式只會讀新信）。")
    else:
        print("資料夾裡沒有比這份資料更新的信件檔。")

    words = [w.casefold() for w in args.words]
    hit = lambda it: not words or any(w in json.dumps(it, ensure_ascii=False).casefold() for w in words)
    full = args.all or bool(words)

    if not words:
        print(f"\n【現況】{r.get('overview_summary', '')}")

    figs = [f for f in r.get("figures") or [] if hit(f)]
    if figs:
        print("\n【數字】")
        for f in figs:
            vs = sorted(f.get("versions", []), key=lambda v: v.get("date", ""))
            for v in (vs if full else vs[-1:]):
                print(f"- {f.get('name', '')}（{f.get('unit', '')}）{v.get('date', '')}　{v.get('by', '')}　{v.get('value', '')}"
                      f"　{v.get('status', '')}{'　' + v['note'] if v.get('note') else ''}{src(v)}")
            if not full and len(vs) > 1:
                print(f"  （共 {len(vs)} 版，這是最新一版；加關鍵字或 --all 看全部）")

    rows = {k: [] for k in ("decisions", "pending", "todos", "completed", "timeline")}
    for t in r.get("topics", []):
        for k in rows:
            rows[k] += [{**it, "_t": t.get("short_name") or t.get("name", "")} for it in t.get(k, []) if hit(it)]

    decs = sorted(rows["decisions"], key=lambda d: d.get("date", ""))
    if decs:
        shown = decs if full else decs[-5:]
        print(f"\n【已決策】{len(decs)} 條" + ("" if full or len(decs) <= 5 else "，列最近 5 條"))
        for d in shown:
            role = f"（{d['decider_role']}）" if d.get("decider_role") else ""
            print(f"- {tags(d)}{d.get('date', '')}　{d.get('decider', '')}{role}：{d.get('decision', '')}〔{d['_t']}〕")
            if d.get("quote"):
                print(f"  原文：「{d['quote']}」")
            if d.get("note"):
                print(f"  備註：{d['note']}")
            if src(d):
                print(f" {src(d)}")

    if rows["pending"]:
        pend = sorted(rows["pending"], key=lambda p: p.get("raised_date") or "9999")
        print(f"\n【尚未決策】{len(pend)} 條（天數算到今天 {today}）")
        for p in pend:
            rd = parse_date(p.get("raised_date"))
            print(f"- {tags(p)}{p.get('issue', '')}｜已等 {(today - rd).days if rd else '?'} 天｜提出：{p.get('raised_by', '')} "
                  f"{p.get('raised_date', '')}｜等待：{p.get('waiting_on', '')}〔{p['_t']}〕"
                  f"{'｜' + p['note'] if p.get('note') else ''}{src(p)}")

    if rows["todos"]:
        todos = sorted(rows["todos"], key=lambda x: x.get("due") or "9999")
        print(f"\n【待辦】{len(todos)} 條")
        for x in todos:
            d = parse_date(x.get("due"))
            state = "期限未定" if not d else (f"已逾期 {(today - d).days} 天" if d < today else f"{x['due']} 到期")
            print(f"- {tags(x)}{x.get('item', '')}｜{x.get('owner') or '負責人未定'}｜{state}〔{x['_t']}〕"
                  f"{'｜' + x['note'] if x.get('note') else ''}{src(x)}")

    if full and rows["completed"]:
        print(f"\n【已完成】{len(rows['completed'])} 條")
        for x in rows["completed"]:
            print(f"- {tags(x)}{x.get('item', '')}｜{x.get('completed_date', '')}｜{x.get('evidence', '')}〔{x['_t']}〕{src(x)}")

    if full and rows["timeline"]:
        tl = sorted(rows["timeline"], key=lambda x: x.get("date", ""))
        print(f"\n【相關信件】{len(tl)} 封")
        for x in tl:
            line = x.get("headline") or (x.get("points") or [x.get("summary", "")])[0]
            print(f"- {x.get('id', '')} {x.get('date', '')}　{x.get('sender', '')}：{line}〔{x['_t']}〕　{x.get('file', '')}")

    if words and not (figs or any(rows.values())):
        print(f"\n上次的案卷裡沒有提到「{'、'.join(args.words)}」的條目。換個說法再查，或重新整理信件。")


if __name__ == "__main__":
    main()
