#!/usr/bin/env python3
"""
把 report.json 產出成案卷報告。

用法：
    python build_report.py report.json --outdir <信件資料夾> --emails <工作目錄>/emails.json \
        [--formats docx,pdf,html,brief,xlsx,ics]

格式：docx（完整報告）、pdf、html、brief（一頁會議速覽 docx）、xlsx（待辦／未決／決策表）、ics（待辦行事曆檔）。
檔名：案卷整理_YYYY-MM-DD.docx；同一天已有同名檔時自動加 _2、_3，不覆蓋舊版。
另會存一份 案卷整理_YYYY-MM-DD.data.json（下次增量整理用，請勿刪除）。
"""
import argparse
import html as H
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "Microsoft JhengHei"   # 微軟正黑體；沒有這個字型的環境會自動替換
ACCENT = RGBColor(0x1F, 0x4E, 0x79)
GREY = RGBColor(0x59, 0x59, 0x59)
WARN = RGBColor(0xC0, 0x50, 0x00)
HEADER_FILL = "DCE6F0"

# ---------- 資料整理 ----------

def parse_date(s):
    try:
        return date.fromisoformat((s or "")[:10])
    except ValueError:
        return None


def conf_label(c):
    return "待確認" if c == "低" else ""


def src_text(s):
    if not s:
        return ""
    parts = [s.get("id", ""), s.get("date", ""), s.get("sender", ""), f"「{s.get('subject', '')}」"]
    if s.get("attachment"):
        parts.append(f"附件《{s['attachment']}》")
    return "　".join(p for p in parts if p)


def _norm(p):
    return str(p).replace("\\", "/").rstrip("/")


def file_href(r, f):
    """信件檔案相對於報告所在資料夾的連結（URL 編碼）。.pst 內的信連到 .pst 檔本身；資訊不足回傳空字串。"""
    if not f:
        return ""
    f = str(f).split(" › ")[0].replace("\\", "/")
    folder, outdir = r.get("folder"), r.get("_outdir")
    if not folder or not outdir:
        return ""
    if _norm(folder) == _norm(outdir):
        rel = f
    else:
        try:
            rel = posixpath.relpath(_norm(folder) + "/" + f, _norm(outdir))
        except ValueError:
            return ""
    return urllib.parse.quote(rel)


def src_html(r, e, s, id_only=False):
    s = s or {}
    text = (s.get("id", "") if id_only else src_text(s))
    href = file_href(r, s.get("file"))
    return f"<a href='{e(href)}'>{e(text)}</a>" if href and text else e(text)


def src_cell(r, s):
    """表格用：回傳 (顯示文字, 連結)；沒有出處回傳空字串。"""
    s = s or {}
    return (s.get("id", ""), file_href(r, s.get("file"))) if s.get("id") else ""


def fill_files(report, files):
    """用 emails.json 的 id→檔案對照，補齊來源與時間序缺少的 file 欄位（超連結要用）。"""
    if not files:
        return
    for t in report.get("topics", []):
        for k in ("decisions", "pending", "todos", "completed"):
            for it in t.get(k, []):
                s = it.get("source")
                if isinstance(s, dict) and s.get("id") and not s.get("file") and s["id"] in files:
                    s["file"] = files[s["id"]]
        for x in t.get("timeline", []):
            if x.get("id") and not x.get("file") and x["id"] in files:
                x["file"] = files[x["id"]]
    for f in report.get("figures") or []:
        for v in f.get("versions", []):
            s = v.get("source")
            if isinstance(s, dict) and s.get("id") and not s.get("file") and s["id"] in files:
                s["file"] = files[s["id"]]


def prepare(report, files=None):
    gen = parse_date(report.get("generated_date")) or date.today()
    report["generated_date"] = gen.isoformat()
    fill_files(report, files)
    all_todos, all_pending = [], []
    for t in report.get("topics", []):
        for k in ("decisions", "pending", "todos", "completed", "timeline"):
            t.setdefault(k, [])
        t["decisions"].sort(key=lambda d: d.get("date", ""))
        for p in t["pending"]:
            rd = parse_date(p.get("raised_date"))
            p["days_pending"] = (gen - rd).days if rd else None
        t["pending"].sort(key=lambda p: -(p["days_pending"] or 0))
        t["todos"].sort(key=lambda x: x.get("due") or "9999")
        t["timeline"].sort(key=lambda x: x.get("date", ""))
        for x in t["todos"]:
            all_todos.append({**x, "topic": t["name"]})
        for p in t["pending"]:
            all_pending.append({**p, "topic": t["name"]})
    all_todos.sort(key=lambda x: x.get("due") or "9999")
    all_pending.sort(key=lambda p: -(p["days_pending"] or 0))
    report["_todos"], report["_pending"] = all_todos, all_pending
    return report


def due_text(x, gen):
    d = parse_date(x.get("due"))
    if not d:
        return "未定"
    diff = (d - gen).days
    if diff < 0:
        return f"{d.isoformat()}（已逾期 {-diff} 天）"
    if diff <= 7:
        return f"{d.isoformat()}（剩 {diff} 天）"
    return d.isoformat()


def pending_days_text(p):
    n = p.get("days_pending")
    return f"{n} 天" if n is not None else "不明"


def entry_points(x):
    """時間序條目的重點：優先用 points（條列），沒有就退回 summary 一句話。"""
    pts = [q for q in (x.get("points") or []) if q]
    return pts or ([x["summary"]] if x.get("summary") else [])


def entry_actions_text(a):
    parts = [a.get("item", "")]
    if a.get("owner"):
        parts.append(f"負責人：{a['owner']}")
    parts.append(f"期限：{a['due']}" if a.get("due") else "期限：未定")
    return "｜".join(p for p in parts if p)


def all_chrono(r):
    """合併所有議題的時間序條目，依信件編號去重，由舊到新。"""
    seen, out = set(), []
    for t in r.get("topics", []):
        for x in t["timeline"]:
            k = x.get("id") or id(x)
            if k in seen:
                continue
            seen.add(k)
            out.append({**x, "_topic": t["name"], "_short": t.get("short_name") or t["name"]})
    out.sort(key=lambda x: x.get("date", ""))
    return out


def short_name(s):
    """「陳小華 Amy（我方業務部）」→「陳小華 Amy」，總時間軸用。"""
    return re.sub(r"[（(].*?[）)]|<.*?>", "", s or "").strip() or (s or "")


def headline(x):
    """總時間軸一行：優先用 headline，否則第一條重點。"""
    if x.get("headline"):
        return x["headline"]
    pts = entry_points(x)
    return pts[0] if pts else ""


def master_rows(r):
    """總時間軸：每封信一行，由舊到新。"""
    return [{"date": (x.get("date") or "")[:16], "who": short_name(x.get("sender", "")), "what": headline(x),
             "tag": x.get("tag", ""), "topic": x["_short"], "id": x.get("id", ""), "file": x.get("file")}
            for x in all_chrono(r)]


def figure_rows(f):
    vs = sorted(f.get("versions", []), key=lambda v: v.get("date", ""))
    return vs


def render_master_docx(doc, r):
    rows = master_rows(r)
    if not rows:
        return
    heading(doc, "總時間軸（一行一件事，由舊到新）", 2)
    multi = len(r.get("topics", [])) > 1
    if multi:
        table(doc, ["日期", "誰", "發生什麼", "類型", "議題", "信件"],
              [[x["date"], x["who"], x["what"], x["tag"], x["topic"], (x["id"], file_href(r, x["file"]))] for x in rows],
              [2.4, 2.4, 7.6, 1.3, 1.8, 1.5])
    else:
        table(doc, ["日期", "誰", "發生什麼", "類型", "信件"],
              [[x["date"], x["who"], x["what"], x["tag"], (x["id"], file_href(r, x["file"]))] for x in rows],
              [2.4, 2.6, 9.0, 1.4, 1.6])


def render_figures_docx(doc, r):
    figs = r.get("figures") or []
    if not figs:
        return
    heading(doc, "數字怎麼變（金額／數量的每一版）", 2)
    para(doc, "同一個數字被誰、在哪天提出，對方接不接受。最下面一列是目前最新的版本。", size=9.5, color=GREY)
    for f in figs:
        p = para(doc, space_after=2)
        set_font(p.add_run(f.get("name", "")), 10.5, True, ACCENT)
        if f.get("unit"):
            set_font(p.add_run(f"（{f['unit']}）"), 9.5, False, GREY)
        if f.get("note"):
            para(doc, f["note"], size=9.5, color=GREY, indent=0.4, space_after=2)
        table(doc, ["日期", "誰提出", "數字", "狀態", "說明", "出處"],
              [[v.get("date", ""), short_name(v.get("by", "")), v.get("value", ""), v.get("status", ""),
                v.get("note", ""), src_cell(r, v.get("source"))] for v in figure_rows(f)],
              [2.5, 2.4, 3.6, 1.4, 5.2, 1.9])


def render_master_html(r, e):
    rows = master_rows(r)
    if not rows:
        return ""
    multi = len(r.get("topics", [])) > 1
    link = lambda x: (f"<a href='{e(file_href(r, x['file']))}'>{e(x['id'])}</a>" if file_href(r, x["file"]) else e(x["id"]))
    head = ["日期", "誰", "發生什麼", "類型"] + (["議題"] if multi else []) + ["信件"]
    body = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in
                                    [e(x["date"]), e(x["who"]), e(x["what"]), e(x["tag"])]
                                    + ([e(x["topic"])] if multi else []) + [link(x)]) + "</tr>" for x in rows)
    return ("<h3>總時間軸（一行一件事，由舊到新）</h3><div class='scroll'><table><tr>"
            + "".join(f"<th>{h}</th>" for h in head) + "</tr>" + body + "</table></div>")


def render_figures_html(r, e):
    figs = r.get("figures") or []
    if not figs:
        return ""
    out = ["<h3>數字怎麼變（金額／數量的每一版）</h3><p class='muted'>最下面一列是目前最新的版本。</p>"]
    for f in figs:
        out.append(f"<p><b>{e(f.get('name'))}</b> <span class='muted'>{e(f.get('unit'))}</span></p>")
        if f.get("note"):
            out.append(f"<p class='muted'>{e(f['note'])}</p>")
        rows = "".join(
            f"<tr><td>{e(v.get('date'))}</td><td>{e(short_name(v.get('by', '')))}</td><td><b>{e(v.get('value'))}</b></td>"
            f"<td>{e(v.get('status'))}</td><td>{e(v.get('note'))}</td><td>{src_html(r, e, v.get('source'), True)}</td></tr>"
            for v in figure_rows(f))
        out.append("<div class='scroll'><table><tr><th>日期</th><th>誰提出</th><th>數字</th><th>狀態</th><th>說明</th><th>出處</th></tr>"
                   + rows + "</table></div>")
    return "".join(out)


def render_entries_docx(doc, r, entries, show_topic=False):
    """每封信一個區塊：時間與寄/收件人、重點條列、關鍵行動、附件與連結。"""
    for x in entries:
        p = para(doc, space_after=1)
        set_font(p.add_run(x.get("date", "")), 10.5, True, ACCENT)
        if x.get("id"):
            set_font(p.add_run("　"), 9.5, False, GREY)
            href = file_href(r, x.get("file"))
            if href:
                add_hyperlink(p, href, f"[{x['id']}]", size=9.5)
            else:
                set_font(p.add_run(f"[{x['id']}]"), 9.5, False, GREY)
        if show_topic and x.get("_topic"):
            set_font(p.add_run(f"　〔{x['_topic']}〕"), 9.5, False, GREY)
        who = f"{x.get('sender', '')} → {x.get('recipients', '') or '（收件者不明）'}"
        if x.get("cc"):
            who += f"　（副本：{x['cc']}）"
        labelled(doc, "寄／收：", who, indent=0.4)
        if x.get("subject"):
            labelled(doc, "主旨：", x["subject"], indent=0.4)
        for q in entry_points(x):
            bp = para(doc, indent=0.9, space_after=1)
            bp.paragraph_format.first_line_indent = Cm(-0.4)
            set_font(bp.add_run("• " + q), 10)
        if x.get("actions"):
            labelled(doc, "關鍵行動：", "", indent=0.4)
            for a in x["actions"]:
                bp = para(doc, indent=1.0, space_after=1)
                bp.paragraph_format.first_line_indent = Cm(-0.4)
                set_font(bp.add_run("▶ " + entry_actions_text(a)), 10, False, WARN)
        if x.get("attachments"):
            labelled(doc, "附件：", "、".join(x["attachments"]), indent=0.4)
        if x.get("attachment_notes"):
            labelled(doc, "附件重點：", "", indent=0.4)
            for q in x["attachment_notes"]:
                bp = para(doc, indent=1.0, space_after=1)
                bp.paragraph_format.first_line_indent = Cm(-0.4)
                set_font(bp.add_run("▪ " + q), 10)
        if x.get("links"):
            labelled(doc, "連結：", "　".join(x["links"]), indent=0.4)
        para(doc, space_after=4)


def render_entries_html(r, entries, e, show_topic=False):
    out = []
    for x in entries:
        who = f"{e(x.get('sender'))} → {e(x.get('recipients') or '（收件者不明）')}"
        if x.get("cc"):
            who += f"（副本：{e(x['cc'])}）"
        topic = f" 〔{e(x['_topic'])}〕" if show_topic and x.get("_topic") else ""
        href = file_href(r, x.get("file"))
        idt = f"<a href='{e(href)}'>[{e(x.get('id'))}]</a>" if href else f"[{e(x.get('id'))}]"
        h = [f"<div class='card'><div class='ch'>{e(x.get('date'))} <span class='muted'>{idt}{topic}</span></div>",
             f"<p class='muted'>寄／收：{who}</p>"]
        if x.get("subject"):
            h.append(f"<p class='muted'>主旨：{e(x['subject'])}</p>")
        pts = entry_points(x)
        if pts:
            h.append("<ul>" + "".join(f"<li>{e(q)}</li>" for q in pts) + "</ul>")
        if x.get("actions"):
            h.append("<p><b class='warnc'>關鍵行動</b></p><ul>"
                     + "".join(f"<li class='warnc'>{e(entry_actions_text(a))}</li>" for a in x["actions"]) + "</ul>")
        if x.get("attachments"):
            h.append(f"<p>附件：{e('、'.join(x['attachments']))}</p>")
        if x.get("attachment_notes"):
            h.append("<p><b>附件重點</b></p><ul>" + "".join(f"<li>{e(q)}</li>" for q in x["attachment_notes"]) + "</ul>")
        if x.get("links"):
            h.append("<p>連結：" + "　".join(f"<a href='{e(u)}'>{e(u)}</a>" for u in x["links"]) + "</p>")
        out.append("".join(h) + "</div>")
    return "".join(out)

# ---------- Word ----------

def set_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def para(doc, text="", size=10.5, bold=False, color=None, italic=False, indent=None, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if text:
        set_font(p.add_run(text), size, bold, color, italic)
    return p


def heading(doc, text, level):
    h = doc.add_heading(level=level)
    set_font(h.add_run(text), {0: 20, 1: 15, 2: 12.5}[level], True, ACCENT)
    h.paragraph_format.space_before = Pt(14 if level < 2 else 10)
    h.paragraph_format.space_after = Pt(6)
    return h


def shade(cell, fill):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def table(doc, headers, rows, widths_cm):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        set_font(c.paragraphs[0].add_run(h), 9.5, True)
        shade(c, HEADER_FILL)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            if isinstance(v, tuple):          # (顯示文字, 連結)
                if v[1]:
                    add_hyperlink(cells[i].paragraphs[0], v[1], v[0], size=9.5)
                else:
                    set_font(cells[i].paragraphs[0].add_run(v[0]), 9.5)
                continue
            color = WARN if v == "待確認" else None
            set_font(cells[i].paragraphs[0].add_run(str(v)), 9.5, v == "待確認", color)
    # 欄寬：Word 讀 cell 寬度，LibreOffice 讀 tblGrid，兩邊都設定並固定版面
    tblpr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:type"), "dxa")
    tblw.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in t.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def empty(doc):
    para(doc, "目前沒有。", color=GREY)


def labelled(doc, label, value, indent=0.6, color=None):
    p = para(doc, indent=indent, space_after=2)
    set_font(p.add_run(label), 10, True, GREY)
    set_font(p.add_run(value), 10, False, color)
    return p


def add_hyperlink(p, url, text, size=10, color=ACCENT):
    r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    run = p.add_run(text)
    set_font(run, size, False, color)
    run.font.underline = True
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    h.append(run._r)
    p._p.append(h)
    return h


def src_line(doc, r, s, indent=0.6):
    p = para(doc, indent=indent, space_after=2)
    set_font(p.add_run("出處："), 10, True, GREY)
    text = src_text(s)
    href = file_href(r, (s or {}).get("file"))
    if href and text:
        add_hyperlink(p, href, text)
    else:
        set_font(p.add_run(text), 10)
    return p


def build_docx(r, path):
    gen = parse_date(r["generated_date"])
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)

    # 封面資訊
    heading(doc, f"案卷整理：{r.get('title', '')}", 0)
    dr = r.get("date_range") or []
    info = [
        ("資料夾：", r.get("folder", "")),
        ("信件數：", f"{r.get('email_count', '')} 封"),
        ("期間：", f"{dr[0]} 至 {dr[1]}" if len(dr) == 2 else "不明"),
        ("產出日期：", r["generated_date"]),
    ]
    for k, v in info:
        labelled(doc, k, v, indent=0)

    # 總覽
    heading(doc, "總覽", 1)
    para(doc, r.get("overview_summary", ""), space_after=8)
    render_changes_docx(doc, r)
    render_master_docx(doc, r)
    render_figures_docx(doc, r)

    heading(doc, "待辦事項（全部議題，依期限排序）", 2)
    if r["_todos"]:
        table(doc, ["事項", "負責人", "期限", "議題", "備註", "出處"],
              [[x.get("item", ""), x.get("owner", ""), due_text(x, gen), x["topic"],
                conf_label(x.get("confidence")), src_cell(r, x.get("source"))] for x in r["_todos"]],
              [5.0, 2.6, 3.2, 2.6, 1.4, 2.2])
    else:
        empty(doc)

    heading(doc, "尚未決策的議題（依懸而未決天數排序）", 2)
    if r["_pending"]:
        table(doc, ["議題", "提出者／日期", "等待", "懸而未決", "所屬議題"],
              [[p.get("issue", ""), f"{p.get('raised_by', '')}\n{p.get('raised_date', '')}",
                p.get("waiting_on", ""), pending_days_text(p), p["topic"]] for p in r["_pending"]],
              [6.0, 3.4, 3.2, 1.8, 2.6])
    else:
        empty(doc)

    # 各議題
    for n, t in enumerate(r.get("topics", []), 1):
        doc.add_page_break()
        heading(doc, f"議題 {n}：{t['name']}", 1)
        para(doc, f"相關信件 {t.get('email_count', '')} 封", size=9.5, color=GREY)
        para(doc, t.get("status_summary", ""), space_after=8)

        heading(doc, "已決策事項", 2)
        if not t["decisions"]:
            empty(doc)
        for d in t["decisions"]:
            p = para(doc, space_after=2)
            role = f"（{d['decider_role']}）" if d.get("decider_role") else ""
            set_font(p.add_run(f"{d.get('date', '')}　{d.get('decider', '')}{role}"), 10.5, True, ACCENT)
            if d.get("confidence") == "低":
                set_font(p.add_run("　待確認"), 10, True, WARN)
            elif d.get("confidence") == "中":
                set_font(p.add_run("　（信心：中）"), 9.5, False, GREY)
            para(doc, d.get("decision", ""), indent=0.6, space_after=2)
            if d.get("quote"):
                para(doc, f"原文：「{d['quote']}」", size=10, italic=True, color=GREY, indent=0.6, space_after=2)
            if d.get("note"):
                labelled(doc, "備註：", d["note"])
            src_line(doc, r, d.get("source"))
            para(doc, space_after=4)

        heading(doc, "尚未決策", 2)
        if not t["pending"]:
            empty(doc)
        for p_ in t["pending"]:
            p = para(doc, space_after=2)
            set_font(p.add_run(p_.get("issue", "")), 10.5, True)
            set_font(p.add_run(f"　已懸而未決 {pending_days_text(p_)}"), 10, True, WARN)
            labelled(doc, "提出：", f"{p_.get('raised_by', '')}，{p_.get('raised_date', '')}")
            if p_.get("options"):
                labelled(doc, "選項：", "；".join(p_["options"]))
            labelled(doc, "等待：", p_.get("waiting_on", ""))
            if p_.get("note"):
                labelled(doc, "備註：", p_["note"])
            src_line(doc, r, p_.get("source"))
            para(doc, space_after=4)

        heading(doc, "待辦事項", 2)
        if t["todos"]:
            table(doc, ["事項", "負責人", "期限", "備註", "出處"],
                  [[x.get("item", ""), x.get("owner", ""), due_text(x, gen),
                    "；".join(v for v in (conf_label(x.get("confidence")), x.get("note", "")) if v),
                    src_cell(r, x.get("source"))]
                   for x in t["todos"]],
                  [5.6, 2.8, 3.2, 3.4, 2.0])
        else:
            empty(doc)

        heading(doc, "已完成事項", 2)
        if t["completed"]:
            table(doc, ["事項", "完成日期", "依據", "備註", "出處"],
                  [[x.get("item", ""), x.get("completed_date", ""), x.get("evidence", ""),
                    conf_label(x.get("confidence")), src_cell(r, x.get("source"))] for x in t["completed"]],
                  [5.0, 2.3, 6.2, 1.5, 2.0])
        else:
            empty(doc)

        heading(doc, "時間序報告（由舊到新）", 2)
        if t["timeline"]:
            render_entries_docx(doc, r, t["timeline"])
        else:
            empty(doc)

    # 完整時間序（跨議題）
    chrono = all_chrono(r)
    if len(r.get("topics", [])) > 1 and chrono:
        doc.add_page_break()
        heading(doc, "附錄：完整時間序報告（全部議題，由舊到新）", 1)
        render_entries_docx(doc, r, chrono, show_topic=True)

    # 人物與角色
    if r.get("people"):
        doc.add_page_break()
        heading(doc, "附錄：人物與角色", 1)
        para(doc, "「可拍板」表示該人在信件中的說法能否視為正式決策，是判斷「已決策事項」信心程度的依據。",
             size=9.5, color=GREY)
        table(doc, ["姓名（別名）", "單位", "角色", "可拍板", "Email"],
              [[p_.get("name", "") + (f"（{'、'.join(p_['aliases'])}）" if p_.get("aliases") else ""),
                p_.get("org", ""), p_.get("role", ""), p_.get("can_decide", ""),
                "\n".join(p_.get("emails", []))] for p_ in r["people"]],
              [3.8, 3.2, 3.6, 1.8, 4.6])

    # 附註
    doc.add_page_break()
    heading(doc, "附註", 1)
    for line in [
        "信心程度：「高」表示信中有明確文字且說話者有權限；「中」表示需要一點推論（轉述、推算期限、只看到附件已寄出）；「低」主要是推論，報告中標示「待確認」。",
        "懸而未決天數：從問題最早被提出的日期，計算到本報告產出日期。",
        "出處的信件編號（E001 等）依真實寄件時間排序，對應資料夾中的原始信件檔；轉寄信裡引用的舊信也拆成獨立編號，連結指向收錄它的那個檔案，時間取自引用標頭（時區為推估）。增量整理時舊信沿用原編號，新信往後接號。出處與信件編號可直接點選開啟原信（需報告與信件資料夾保持原本的相對位置；Word 可能跳出安全提示）。原文引用保留原語言。",
        "本報告由 AI 依信件內容整理，重要決策開會前請再核對原始信件。",
    ]:
        para(doc, line, size=9.5, color=GREY)

    doc.save(path)

# ---------- HTML ----------

def build_html(r, path):
    gen = parse_date(r["generated_date"])
    e = lambda s: H.escape(str(s or ""))
    warn = lambda c: '<span class="warn">待確認</span>' if c == "低" else ""

    def tbl(headers, rows):
        out = ["<div class='scroll'><table><tr>" + "".join(f"<th>{e(h)}</th>" for h in headers) + "</tr>"]
        for row in rows:
            out.append("<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>")
        return "".join(out) + "</table></div>"

    none = "<p class='muted'>目前沒有。</p>"
    mid = lambda d: "<span class='muted'>（信心：中）</span>" if d.get('confidence') == '中' else ''
    dr = r.get("date_range") or []
    b = [f"<h1>案卷整理：{e(r.get('title'))}</h1>",
         f"<p class='muted'>資料夾：{e(r.get('folder'))}　｜　信件 {e(r.get('email_count'))} 封　｜　"
         f"期間：{e(dr[0] if dr else '')} 至 {e(dr[1] if len(dr) > 1 else '')}　｜　產出：{e(r['generated_date'])}</p>",
         "<h2>總覽</h2>", f"<p>{e(r.get('overview_summary'))}</p>", render_changes_html(r, e),
         render_master_html(r, e), render_figures_html(r, e),
         "<h3>待辦事項（全部議題）</h3>",
         tbl(["事項", "負責人", "期限", "議題", "", "出處"],
             [[e(x.get("item")), e(x.get("owner")), e(due_text(x, gen)), e(x["topic"]), warn(x.get("confidence")),
               src_html(r, e, x.get("source"), True)]
              for x in r["_todos"]]) if r["_todos"] else none,
         "<h3>尚未決策的議題</h3>",
         tbl(["議題", "提出者／日期", "等待", "懸而未決", "所屬議題"],
             [[e(p.get("issue")), f"{e(p.get('raised_by'))}<br>{e(p.get('raised_date'))}", e(p.get("waiting_on")),
               f"<b class='warnc'>{e(pending_days_text(p))}</b>", e(p["topic"])] for p in r["_pending"]])
         if r["_pending"] else none]

    for n, t in enumerate(r.get("topics", []), 1):
        b += [f"<h2>議題 {n}：{e(t['name'])}</h2>",
              f"<p class='muted'>相關信件 {e(t.get('email_count'))} 封</p><p>{e(t.get('status_summary'))}</p>",
              "<h3>已決策事項</h3>"]
        if not t["decisions"]:
            b.append(none)
        for d in t["decisions"]:
            role = f"（{e(d.get('decider_role'))}）" if d.get("decider_role") else ""
            b.append(f"<div class='card'><div class='ch'>{e(d.get('date'))}　{e(d.get('decider'))}{role} {warn(d.get('confidence'))}{mid(d)}</div>"
                     f"<p>{e(d.get('decision'))}</p>"
                     + (f"<blockquote>「{e(d.get('quote'))}」</blockquote>" if d.get("quote") else "")
                     + (f"<p class='muted'>備註：{e(d.get('note'))}</p>" if d.get("note") else "")
                     + f"<p class='muted'>出處：{src_html(r, e, d.get('source'))}</p></div>")
        b.append("<h3>尚未決策</h3>")
        if not t["pending"]:
            b.append(none)
        for p in t["pending"]:
            b.append(f"<div class='card'><div class='ch'>{e(p.get('issue'))} <span class='warnc'>已懸而未決 {e(pending_days_text(p))}</span></div>"
                     f"<p class='muted'>提出：{e(p.get('raised_by'))}，{e(p.get('raised_date'))}</p>"
                     + (f"<p>選項：{e('；'.join(p['options']))}</p>" if p.get("options") else "")
                     + f"<p>等待：{e(p.get('waiting_on'))}</p>"
                     + (f"<p class='muted'>備註：{e(p.get('note'))}</p>" if p.get("note") else "")
                     + f"<p class='muted'>出處：{src_html(r, e, p.get('source'))}</p></div>")
        b += ["<h3>待辦事項</h3>",
              tbl(["事項", "負責人", "期限", "備註", "出處"],
                  [[e(x.get("item")), e(x.get("owner")), e(due_text(x, gen)),
                    warn(x.get("confidence")) + e(x.get("note")), src_html(r, e, x.get("source"), True)]
                   for x in t["todos"]]) if t["todos"] else none,
              "<h3>已完成事項</h3>",
              tbl(["事項", "完成日期", "依據", "", "出處"],
                  [[e(x.get("item")), e(x.get("completed_date")), e(x.get("evidence")), warn(x.get("confidence")),
                    src_html(r, e, x.get("source"), True)]
                   for x in t["completed"]]) if t["completed"] else none,
              "<h3>時間序報告（由舊到新）</h3>",
              render_entries_html(r, t["timeline"], e) if t["timeline"] else none]

    chrono = all_chrono(r)
    if len(r.get("topics", [])) > 1 and chrono:
        b += ["<h2>附錄：完整時間序報告（全部議題，由舊到新）</h2>", render_entries_html(r, chrono, e, show_topic=True)]

    if r.get("people"):
        b += ["<h2>附錄：人物與角色</h2>",
              tbl(["姓名（別名）", "單位", "角色", "可拍板", "Email"],
                  [[e(q.get("name", "") + (f"（{'、'.join(q['aliases'])}）" if q.get("aliases") else "")),
                    e(q.get("org")), e(q.get("role")), e(q.get("can_decide")), e("、".join(q.get("emails", [])))]
                   for q in r["people"]])]
    b.append("<h2>附註</h2><p class='muted'>信心程度「低」的項目標示為「待確認」。懸而未決天數從問題最早提出日計算到報告產出日。"
             "出處的信件編號依寄件時間排序。本報告由 AI 整理，重要決策請核對原始信件。</p>")

    css = """
:root{--bg:#fff;--fg:#222;--muted:#666;--accent:#1f4e79;--line:#d0d7de;--head:#dce6f0;--card:#f7f9fb;--warn:#c05000}
@media (prefers-color-scheme:dark){:root{--bg:#16181c;--fg:#e6e6e6;--muted:#9aa0a6;--accent:#7fb3e6;--line:#333a44;--head:#22303f;--card:#1e2228;--warn:#ff9a4d}}
body{background:var(--bg);color:var(--fg);font-family:"Microsoft JhengHei","PingFang TC","Noto Sans TC",sans-serif;
max-width:1000px;margin:0 auto;padding:24px 16px;line-height:1.7}
h1,h2,h3{color:var(--accent)} h2{border-bottom:2px solid var(--line);padding-bottom:4px;margin-top:40px}
.muted{color:var(--muted);font-size:.92em} .scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:.93em;margin:8px 0 16px}
th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top;text-align:left} th{background:var(--head)}
.card{background:var(--card);border-left:4px solid var(--accent);padding:8px 14px;margin:10px 0;border-radius:4px}
.card p{margin:4px 0} .card ul{margin:4px 0 4px 1.2em;padding:0} .ch{font-weight:bold}
blockquote{margin:6px 0;padding-left:10px;border-left:3px solid var(--line);font-style:italic;color:var(--muted)}
.warn{color:var(--warn);font-weight:bold;font-size:.9em;border:1px solid var(--warn);border-radius:3px;padding:0 4px}
.warnc{color:var(--warn);font-weight:bold}"""
    page = (f"<!DOCTYPE html><html lang='zh-Hant'><head><meta charset='utf-8'>"
            f"<meta name='viewport' content='width=device-width, initial-scale=1'>"
            f"<title>案卷整理：{e(r.get('title'))}</title><style>{css}</style></head><body>{''.join(b)}</body></html>")
    Path(path).write_text(page, encoding="utf-8")

# ---------- 增量：自上次報告以來的變化 ----------

def _items(report):
    out = {k: [] for k in ("decisions", "pending", "todos", "completed")}
    for t in report.get("topics", []):
        for k in out:
            for it in t.get(k, []):
                out[k].append({**it, "_topic": t["name"]})
    return out


def _sid(it):
    return (it.get("source") or {}).get("id", "")


def compute_changes(r, prev, new_ids):
    """比對上次報告與本次報告。「新」的判斷依據是出處信件是否為本次新增的信（new_ids）。"""
    pr = prev.get("report", {})
    prev_gen, gen = parse_date(pr.get("generated_date")), parse_date(r["generated_date"])
    cur, old = _items(r), _items(pr)
    isnew = lambda it: _sid(it) in new_ids

    def todo_line(x, extra=""):
        return f"{x.get('item', '')}｜{x.get('owner') or '負責人未定'}｜{x.get('due') or '期限未定'}{extra}（{x['_topic']}）"

    cur_todo_keys = {(_sid(x), str(x.get("item", "")).strip()) for x in cur["todos"]}
    cur_pend_keys = {(_sid(x), str(x.get("issue", "")).strip()) for x in cur["pending"]}
    cur_pend_src = {_sid(x) for x in cur["pending"]}

    newly, still = [], 0
    for x in cur["todos"]:
        d = parse_date(x.get("due"))
        if d and prev_gen and d < prev_gen:
            still += 1
        elif d and prev_gen and prev_gen <= d < gen:
            newly.append(todo_line(x, f"｜已逾期 {(gen - d).days} 天"))

    sections = [
        ("新增決策", [f"{d.get('date', '')}　{d.get('decider', '')}：{d.get('decision', '')}（{d['_topic']}）"
                  for d in cur["decisions"] if isnew(d)]),
        ("新增尚未決策", [f"{p.get('issue', '')}｜提出：{p.get('raised_by', '')}｜等待：{p.get('waiting_on', '')}（{p['_topic']}）"
                    for p in cur["pending"] if isnew(p)]),
        ("新增待辦", [todo_line(x) for x in cur["todos"] if isnew(x)]),
        ("剛逾期的待辦（上次報告後到期，尚未看到完成跡象）", newly),
        ("新增已完成", [f"{x.get('item', '')}｜{x.get('completed_date', '')}（{x['_topic']}）"
                   for x in cur["completed"] if isnew(x)]),
        ("已移出待辦（可能已完成或取消，請確認）",
         [todo_line(x) for x in old["todos"] if (_sid(x), str(x.get("item", "")).strip()) not in cur_todo_keys]),
        ("已不在未決清單（可能已決策或結案）",
         [f"{p.get('issue', '')}（{p['_topic']}）" for p in old["pending"]
          if (_sid(p), str(p.get("issue", "")).strip()) not in cur_pend_keys and _sid(p) not in cur_pend_src]),
    ]
    return {"prev_gen": pr.get("generated_date", ""), "new_email_count": len(new_ids),
            "sections": [(a, b) for a, b in sections if b], "still_overdue": still}


def render_changes_docx(doc, r):
    c = r.get("_changes")
    if not c:
        return
    heading(doc, f"自上次報告以來的變化（上次 {c['prev_gen']}；新增 {c['new_email_count']} 封信）", 2)
    if r.get("changes_summary"):
        para(doc, r["changes_summary"], space_after=6)
    if not c["sections"]:
        para(doc, "上次報告後沒有新的決策、待辦或未決議題。", color=GREY)
    for label, lines in c["sections"]:
        p = para(doc, space_after=1)
        set_font(p.add_run(f"{label}（{len(lines)}）"), 10.5, True, ACCENT)
        for ln in lines:
            bp = para(doc, indent=0.9, space_after=1)
            bp.paragraph_format.first_line_indent = Cm(-0.4)
            set_font(bp.add_run("• " + ln), 10)
    if c["still_overdue"]:
        para(doc, f"另有 {c['still_overdue']} 項待辦在上次報告前就已逾期，仍未看到完成跡象。", size=9.5, color=WARN)


def render_changes_html(r, e):
    c = r.get("_changes")
    if not c:
        return ""
    h = [f"<h3>自上次報告以來的變化（上次 {e(c['prev_gen'])}；新增 {c['new_email_count']} 封信）</h3>"]
    if r.get("changes_summary"):
        h.append(f"<p>{e(r['changes_summary'])}</p>")
    if not c["sections"]:
        h.append("<p class='muted'>上次報告後沒有新的決策、待辦或未決議題。</p>")
    for label, lines in c["sections"]:
        h.append(f"<p><b>{e(label)}（{len(lines)}）</b></p><ul>" + "".join(f"<li>{e(x)}</li>" for x in lines) + "</ul>")
    if c["still_overdue"]:
        h.append(f"<p class='muted'>另有 {c['still_overdue']} 項待辦在上次報告前就已逾期，仍未看到完成跡象。</p>")
    return "".join(h)

# ---------- 一頁會議速覽 ----------

def build_brief(r, path):
    gen = parse_date(r["generated_date"])
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(1.6)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)

    def h(text):
        p = para(doc, space_after=2)
        p.paragraph_format.space_before = Pt(6)
        set_font(p.add_run(text), 11, True, ACCENT)

    def bullet(text, color=None):
        bp = para(doc, indent=0.6, space_after=1)
        bp.paragraph_format.first_line_indent = Cm(-0.4)
        set_font(bp.add_run("• " + text), 10, False, color)

    p = para(doc, space_after=1)
    set_font(p.add_run(f"會議速覽｜{r.get('title', '')}"), 16, True, ACCENT)
    dr = r.get("date_range") or []
    para(doc, f"資料至 {dr[1] if len(dr) > 1 else '不明'}　｜　產出 {gen}　｜　信件 {r.get('email_count', '')} 封",
         size=9, color=GREY, space_after=4)
    para(doc, r.get("overview_summary", ""), size=10, space_after=2)

    c = r.get("_changes")
    if c:
        h(f"自上次報告（{c['prev_gen']}）以來")
        cnt = {a.split("（")[0]: len(b) for a, b in c["sections"]}
        bits = [f"新增 {c['new_email_count']} 封信"] + [f"{k} {v} 項" for k, v in cnt.items()]
        bullet("；".join(bits))
        if r.get("changes_summary"):
            bullet(r["changes_summary"])

    if r.get("figures"):
        h("關鍵數字（目前最新版本）")
        for f in r["figures"]:
            vs = figure_rows(f)
            if vs:
                v = vs[-1]
                bullet(f"{f.get('name', '')}：{v.get('value', '')}（{v.get('status', '')}，{v.get('date', '')} {short_name(v.get('by', ''))}）"
                       + (f"；共 {len(vs)} 版" if len(vs) > 1 else ""))

    h("需要在會議上決議（尚未決策，懸而未決最久優先）")
    if r["_pending"]:
        for q in r["_pending"][:5]:
            bullet(f"{q.get('issue', '')}　→ 等待：{q.get('waiting_on', '')}　（已 {pending_days_text(q)}）")
        if len(r["_pending"]) > 5:
            para(doc, f"另有 {len(r['_pending']) - 5} 項，見完整報告。", size=9, color=GREY, indent=0.6)
    else:
        para(doc, "目前沒有。", size=10, color=GREY, indent=0.6)

    h("逾期與 7 天內到期的待辦")
    near = [x for x in r["_todos"] if parse_date(x.get("due")) and (parse_date(x["due"]) - gen).days <= 7]
    if near:
        table(doc, ["事項", "負責人", "期限"],
              [[x.get("item", ""), x.get("owner", ""), due_text(x, gen)] for x in near[:8]], [8.5, 3.6, 4.9])
    else:
        para(doc, "目前沒有。", size=10, color=GREY, indent=0.6)
    extra = len(near) - 8
    nodue = sum(1 for x in r["_todos"] if not parse_date(x.get("due")))
    notes = ([f"另有 {extra} 項近期待辦"] if extra > 0 else []) + ([f"{nodue} 項待辦未訂期限"] if nodue else [])
    if notes:
        para(doc, "；".join(notes) + "，見完整報告。", size=9, color=GREY, indent=0.6)

    h("最近的決策")
    decs = sorted([{**d, "_topic": t["name"]} for t in r.get("topics", []) for d in t["decisions"]],
                  key=lambda d: d.get("date", ""), reverse=True)[:4]
    if decs:
        for d in decs:
            warn = "（待確認）" if d.get("confidence") == "低" else ""
            bullet(f"{d.get('date', '')}　{d.get('decider', '')}：{d.get('decision', '')}{warn}")
    else:
        para(doc, "目前沒有。", size=10, color=GREY, indent=0.6)

    low = sum(1 for t in r.get("topics", []) for k in ("decisions", "pending", "todos", "completed")
              for i in t[k] if i.get("confidence") == "低")
    para(doc, f"標示「待確認」共 {low} 條。本速覽由 AI 整理，重要決策請核對完整報告與原始信件。",
         size=8.5, color=GREY, space_after=0)
    doc.save(path)

# ---------- 待辦匯出：xlsx / ics ----------

def todo_status(x, gen):
    d = parse_date(x.get("due"))
    if not d:
        return "未定"
    diff = (d - gen).days
    return "已逾期" if diff < 0 else ("7 天內到期" if diff <= 7 else "進行中")


def build_xlsx(r, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    gen = parse_date(r["generated_date"])
    wb = Workbook()
    head_fill = PatternFill("solid", fgColor=HEADER_FILL)
    link_font = Font(color="1F4E79", underline="single")

    def sheet(ws, headers, rows, widths, link_col=None, hrefs=None, date_cols=()):
        ws.append(headers)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True), head_fill
        for i, row in enumerate(rows):
            ws.append(row)
            for ci in date_cols:
                v = parse_date(row[ci])
                if v:
                    ws.cell(row=i + 2, column=ci + 1).value = v
                    ws.cell(row=i + 2, column=ci + 1).number_format = "yyyy-mm-dd"
            if link_col is not None and hrefs and hrefs[i]:
                cell = ws.cell(row=i + 2, column=link_col + 1)
                cell.hyperlink, cell.font = hrefs[i], link_font
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    ws = wb.active
    ws.title = "待辦"
    todos = r["_todos"]
    sheet(ws, ["議題", "事項", "負責人", "期限", "狀態", "信心", "備註", "出處", "完成？"],
          [[x["topic"], x.get("item", ""), x.get("owner", ""), x.get("due") or "未定", todo_status(x, gen),
            x.get("confidence", ""), x.get("note", ""), _sid(x), ""] for x in todos],
          [16, 44, 18, 12, 12, 7, 28, 8, 9], 7, [file_href(r, (x.get("source") or {}).get("file")) for x in todos],
          date_cols=(3,))
    pend = r["_pending"]
    sheet(wb.create_sheet("尚未決策"), ["議題", "問題", "提出者", "提出日", "等待", "懸而未決（天）", "選項", "備註", "出處"],
          [[p["topic"], p.get("issue", ""), p.get("raised_by", ""), p.get("raised_date", ""), p.get("waiting_on", ""),
            p.get("days_pending"), "；".join(p.get("options") or []), p.get("note", ""), _sid(p)] for p in pend],
          [16, 44, 16, 12, 20, 12, 30, 28, 8], 8, [file_href(r, (p.get("source") or {}).get("file")) for p in pend],
          date_cols=(3,))
    decs = sorted([{**d, "_topic": t["name"]} for t in r.get("topics", []) for d in t["decisions"]],
                  key=lambda d: d.get("date", ""))
    sheet(wb.create_sheet("已決策"), ["議題", "日期", "決定者", "角色", "決策", "原文", "信心", "備註", "出處"],
          [[d["_topic"], d.get("date", ""), d.get("decider", ""), d.get("decider_role", ""), d.get("decision", ""),
            d.get("quote", ""), d.get("confidence", ""), d.get("note", ""), _sid(d)] for d in decs],
          [16, 12, 14, 16, 44, 40, 7, 28, 8], 8, [file_href(r, (d.get("source") or {}).get("file")) for d in decs],
          date_cols=(1,))
    figs = [(f.get("name", ""), f.get("unit", ""), v) for f in r.get("figures") or [] for v in figure_rows(f)]
    if figs:
        sheet(wb.create_sheet("數字演變"), ["項目", "單位", "日期", "誰提出", "數字", "狀態", "說明", "出處"],
              [[n, u, v.get("date", ""), short_name(v.get("by", "")), v.get("value", ""), v.get("status", ""),
                v.get("note", ""), (v.get("source") or {}).get("id", "")] for n, u, v in figs],
              [24, 8, 12, 14, 22, 8, 40, 8], 7, [file_href(r, (v.get("source") or {}).get("file")) for _, _, v in figs],
              date_cols=(2,))
    mr = master_rows(r)
    if mr:
        sheet(wb.create_sheet("總時間軸"), ["日期", "誰", "發生什麼", "類型", "議題", "信件"],
              [[x["date"], x["who"], x["what"], x["tag"], x["topic"], x["id"]] for x in mr],
              [17, 16, 60, 8, 18, 8], 5, [file_href(r, x["file"]) for x in mr])
    wb.save(path)


def _ics_esc(t):
    return str(t).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\r", "").replace("\n", "\\n")


def _ics_fold(line, limit=70):
    b, parts = line.encode("utf-8"), []
    while len(b) > limit:
        cut = limit
        while cut > 0 and (b[cut] & 0xC0) == 0x80:
            cut -= 1
        parts.append(b[:cut].decode("utf-8"))
        b = b[cut:]
    parts.append(b.decode("utf-8"))
    return "\r\n ".join(parts)


def build_ics(r, path):
    """有期限的待辦 → 全天事件，到期前一天早上 9 點提醒。回傳 (事件數, 無期限而略過的數量)。"""
    import hashlib
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//email-case-file//待辦//ZH", "CALSCALE:GREGORIAN",
             "X-WR-CALNAME:" + _ics_esc(f"案卷待辦：{r.get('title', '')}")]
    n = skipped = 0
    for x in r["_todos"]:
        d = parse_date(x.get("due"))
        if not d:
            skipped += 1
            continue
        uid = hashlib.sha1(f"{x['topic']}|{_sid(x)}|{x.get('item', '')}".encode("utf-8")).hexdigest()[:20]
        desc = [f"議題：{x['topic']}", f"負責人：{x.get('owner') or '未定'}", f"出處：{src_text(x.get('source'))}"]
        if x.get("note"):
            desc.append(f"備註：{x['note']}")
        if x.get("confidence") == "低":
            desc.append("（待確認）")
        title = f"[{x['topic']}] {x.get('item', '')}" + (f"（{x['owner']}）" if x.get("owner") else "")
        lines += ["BEGIN:VEVENT", f"UID:{uid}@email-case-file", f"DTSTAMP:{stamp}",
                  f"DTSTART;VALUE=DATE:{d.strftime('%Y%m%d')}",
                  f"DTEND;VALUE=DATE:{(d + timedelta(days=1)).strftime('%Y%m%d')}",
                  "SUMMARY:" + _ics_esc(title), "DESCRIPTION:" + _ics_esc("\n".join(desc)),
                  "BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:" + _ics_esc("待辦明日到期"), "TRIGGER:-PT15H",
                  "END:VALARM", "END:VEVENT"]
        n += 1
    lines.append("END:VCALENDAR")
    Path(path).write_bytes(("\r\n".join(_ics_fold(l) for l in lines) + "\r\n").encode("utf-8"))
    return n, skipped

# ---------- PDF ----------

def build_pdf(docx_path, pdf_path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(docx_path)],
                           capture_output=True, timeout=180)
            out = Path(tmp) / (Path(docx_path).stem + ".pdf")
            if out.exists():
                shutil.move(str(out), pdf_path)
                return True
    try:
        from docx2pdf import convert  # Windows 上用本機 Word 轉檔
        convert(str(docx_path), str(pdf_path))
        return Path(pdf_path).exists()
    except Exception:
        return False

# ---------- 主程式 ----------

def unique_base(outdir, gen):
    base = f"案卷整理_{gen}"
    n = 1
    while any((outdir / f"{base if n == 1 else f'{base}_{n}'}.{ext}").exists() for ext in ("docx", "pdf", "html", "data.json")):
        n += 1
    return base if n == 1 else f"{base}_{n}"


def load_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception as ex:
        print(f"讀取失敗：{path}（{ex}）")
        return None


def save_sidecar(r, emails_data, path):
    """存下本次報告與信件索引，下次增量整理時用來比對新舊信件與報告內容。"""
    rep_ = {k: v for k, v in r.items() if not k.startswith("_")}
    idx = [{"uid": e.get("uid"), "alt_uids": e.get("alt_uids", []), "src": e.get("src", "top"),
            "id": e["id"], "file": e["file"], "date": e.get("date_display", "")}
           for e in (emails_data or {}).get("emails", [])]
    Path(path).write_text(json.dumps({"generated_date": r["generated_date"], "report": rep_, "email_index": idx},
                                     ensure_ascii=False, indent=1), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--formats", default="docx", help="docx,pdf,html,brief,xlsx,ics（逗號分隔）")
    ap.add_argument("--emails", help="parse_emails.py 產出的 emails.json；提供後才有來源超連結、增量比對與下次可比對的索引")
    ap.add_argument("--prev", help="指定上次報告資料檔（*.data.json）；預設用 emails.json 記錄的那一份")
    ap.add_argument("--no-prev", action="store_true", help="不產生「自上次報告以來的變化」")
    args = ap.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    emails_data = load_json(args.emails) if args.emails else None
    files = {e["id"]: e["file"] for e in emails_data["emails"]} if emails_data else {}

    r = json.loads(Path(args.report).read_text(encoding="utf-8"))
    r["_outdir"] = str(outdir)
    r = prepare(r, files)

    prev = None
    if not args.no_prev and emails_data:
        if args.prev:
            prev = load_json(args.prev)
        else:
            name = (emails_data.get("prev") or {}).get("file")
            if name and (outdir / name).exists():
                prev = load_json(outdir / name)
    if prev:
        r["_changes"] = compute_changes(r, prev, {e["id"] for e in emails_data["emails"] if e.get("is_new")})

    fmts = [f.strip().lower() for f in args.formats.split(",") if f.strip()]
    base = unique_base(outdir, r["generated_date"])

    made = []
    docx_path = outdir / f"{base}.docx"
    if "docx" in fmts or "pdf" in fmts:
        build_docx(r, docx_path)
        if "docx" in fmts:
            made.append(docx_path)
    if "pdf" in fmts:
        pdf_path = outdir / f"{base}.pdf"
        if build_pdf(docx_path, pdf_path):
            made.append(pdf_path)
        else:
            print("PDF 轉檔失敗：找不到 LibreOffice 或 Word。請用 Word 開啟 docx 後「另存為 PDF」。")
        if "docx" not in fmts:
            docx_path.unlink(missing_ok=True)
    if "html" in fmts:
        html_path = outdir / f"{base}.html"
        build_html(r, html_path)
        made.append(html_path)
    if "brief" in fmts:
        bp = outdir / f"{base}_會議速覽.docx"
        build_brief(r, bp)
        made.append(bp)
    if "xlsx" in fmts:
        xp = outdir / f"{base}_待辦.xlsx"
        build_xlsx(r, xp)
        made.append(xp)
    if "ics" in fmts:
        ip = outdir / f"{base}_待辦.ics"
        n, skipped = build_ics(r, ip)
        made.append(ip)
        print(f"行事曆檔：{n} 項待辦已轉為全天事件" + (f"；{skipped} 項沒有期限，未納入" if skipped else ""))

    sidecar = outdir / f"{base}.data.json"
    save_sidecar(r, emails_data, sidecar)
    if not emails_data:
        print("注意：沒有提供 --emails，來源超連結與下次增量比對無法使用。")

    t = r.get("topics", [])
    print(f"議題 {len(t)} 個；決策 {sum(len(x['decisions']) for x in t)} 條；"
          f"尚未決策 {len(r['_pending'])} 條；待辦 {len(r['_todos'])} 條")
    low = sum(1 for x in t for k in ("decisions", "pending", "todos", "completed") for i in x[k] if i.get("confidence") == "低")
    print(f"標示「待確認」：{low} 條")
    if r.get("_changes"):
        c = r["_changes"]
        print(f"自上次報告（{c['prev_gen']}）以來：新增 {c['new_email_count']} 封信；變化項目 "
              + "、".join(f"{a.split('（')[0]} {len(b)}" for a, b in c["sections"]))
    for m in made:
        print(f"已產出：{m}")
    print(f"增量資料檔（下次整理要用，請保留）：{sidecar}")


if __name__ == "__main__":
    main()
