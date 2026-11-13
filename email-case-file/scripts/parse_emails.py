#!/usr/bin/env python3
"""
解析 .msg / .eml / .pst 信件：把轉寄／回覆信裡的引用前文拆成一封封獨立的信，跨檔去重後依真實寄件時間排序。

用法：
    python parse_emails.py <信件資料夾> --out <工作目錄> [--no-recursive] [--tz-offset 8]
                           [--read-attachments] [--attach-chars 500]

輸出：
    <工作目錄>/emails.json   完整資料
    <工作目錄>/emails.txt    方便閱讀的純文字版
"""
import argparse
import base64
import difflib
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path

BODY_LIMIT = 8000
QUOTE_EXCERPT_LIMIT = 1500

# ---------- 文字處理 ----------

def html_to_text(s):
    if not s:
        return ""
    if isinstance(s, bytes):
        for enc in ("utf-8", "cp950", "big5", "latin-1"):
            try:
                s = s.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    s = re.sub(r"(?is)<(script|style|head).*?</\1>", "", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</(p|div|tr|li|h\d)>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    s = re.sub(r"[ \t\xa0]+", " ", s)
    s = re.sub(r"\n\s*\n\s*\n+", "\n\n", s)
    return s.strip()


HARD_MARKERS = [
    r"^-{2,}\s*Original Message\s*-{2,}",
    r"^-{2,}\s*原始(郵件|訊息|邮件)\s*-{2,}",
    r"^-{2,}\s*(Forwarded message|轉寄的郵件|轉寄郵件)\s*-{2,}",
    r"^_{10,}\s*$",
    r"^On .{5,250}wrote:\s*$",
    r"^在 .{5,250}(寫道|写道)\s*[:：]?\s*$",
    r"^於 .{5,250}(寫道|写道)\s*[:：]?\s*$",
]
FROM_LINE = r"^\s*\**\s*(From|寄件者|寄件人|發件人|发件人)\s*\**\s*[:：]"
HEADER_FOLLOW = r"^\s*\**\s*(Sent|Date|To|Subject|傳送時間|寄件日期|日期|收件者|收件人|主旨|主題)\s*\**\s*[:：]"


def split_quoted(body):
    """回傳 (新寫的內容, 被引用的前文)"""
    lines = body.splitlines()
    cut = None
    for i, line in enumerate(lines):
        s = line.strip()
        if any(re.match(p, s, re.I) for p in HARD_MARKERS):
            cut = i
            break
        if re.match(FROM_LINE, s, re.I):
            window = lines[i + 1:i + 6]
            if any(re.match(HEADER_FOLLOW, w.strip(), re.I) for w in window):
                cut = i
                break
        # 多語言引用標頭（義大利文 Da:/Data:、德文 Von:/Gesendet: 等），與 split_thread 同一套規則
        if H_FROM.match(s) and any(H_DATE.match(w) for w in lines[i + 1:i + 7]):
            cut = i
            break
        if WROTE_RE.match(s) or WROTE_ZH.match(s):
            cut = i
            break
    new_part = lines if cut is None else lines[:cut]
    quoted = [] if cut is None else lines[cut:]
    new_part = [l for l in new_part if not l.lstrip().startswith(">")]
    new_text = "\n".join(new_part).strip()
    quoted_text = "\n".join(quoted).strip()
    new_text = re.sub(r"\n\s*\n\s*\n+", "\n\n", new_text)
    return new_text, quoted_text


# ---------- 拆信串：把引用前文拆成一封封獨立的信 ----------
# 轉寄／回覆信常把幾個月的往來整串夾在下面。只看最上面那段會漏掉大半歷史，
# 所以逐段找出引用標頭（From/Sent/To/Subject，含中、義、德、法、西文），拆成子信件。

_K_FROM = r"From|寄件者|寄件人|發件人|发件人|Von|Da|De|Van"
_K_DATE = r"Sent|Date|傳送時間|寄件日期|日期|發送時間|发送时间|Gesendet|Datum|Data|Inviato|Envoyé|Enviado|Fecha"
_K_TO = r"To|收件者|收件人|An|A|À|Para|Aan"
_K_CC = r"Cc|CC|副本|抄送|Kopie"
_K_SUBJ = r"Subject|主旨|主題|主题|Betreff|Oggetto|Objet|Asunto|Onderwerp"
_HKEY = lambda keys: re.compile(rf"^\s*\**\s*({keys})\s*\**\s*[:：]\s*(.*)$", re.I)
H_FROM, H_DATE, H_TO, H_CC, H_SUBJ = (_HKEY(k) for k in (_K_FROM, _K_DATE, _K_TO, _K_CC, _K_SUBJ))
H_OTHER = _HKEY(r"Reply-To|Rispondi a|Antwort an|Répondre à|Importance|Wichtigkeit|Priorità|重要性|Attachments|附件")
TZ_ABBR = {"cest": 2, "mesz": 2, "cet": 1, "mez": 1, "bst": 1, "gmt": 0, "utc": 0, "edt": -4, "est": -5,
           "pdt": -7, "pst": -8, "jst": 9}
WROTE_RE = re.compile(r"^\s*(?:On|Am|Il|Le|El)\s+(.{5,200}?),?\s+(.{1,120}?)\s+(?:wrote|schrieb|ha scritto|a écrit|escribió)\s*:\s*$", re.I)
WROTE_ZH = re.compile(r"^\s*[在於]\s*(.{5,120}?)[，,]?\s*(.{1,80}?)\s*(?:寫道|写道)\s*[:：]?\s*$")
SEPARATOR_RE = re.compile(r"^\s*(-{2,}.*(Original Message|原始(郵件|訊息|邮件)|Forwarded message|轉寄|Messaggio originale|Ursprüngliche Nachricht|Message d'origine).*-{2,}|_{10,})\s*$", re.I)

MONTHS = {}
for _i, _names in enumerate([
    "january jan gennaio gen januar jän janvier janv enero ene",
    "february feb febbraio februar février févr febrero",
    "march mar marzo märz mär mars",
    "april apr aprile avril abril",
    "may maggio mag mai mayo",
    "june jun giugno giu juni juin junio",
    "july jul luglio lug juli juillet juil julio",
    "august aug agosto ago août aout",
    "september sep sept settembre set septembre septiembre",
    "october oct ottobre ott oktober okt octobre octubre",
    "november nov novembre novembre noviembre",
    "december dec dicembre dic dezember dez décembre déc diciembre"], 1):
    for _n in _names.split():
        MONTHS.setdefault(_n, _i)
MONTH_RE = "|".join(sorted(MONTHS, key=len, reverse=True))
EURO_HINT = re.compile(r"alle ore|\bum\b|lunedì|martedì|mercoledì|giovedì|venerdì|sabato|domenica|montag|dienstag|mittwoch|"
                       r"donnerstag|freitag|samstag|sonntag|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|"
                       r"settembre|ottobre|novembre|dicembre|januar|februar|märz|juni|juli|oktober|dezember|lundi|mardi|"
                       r"mercredi|jeudi|vendredi|lunes|martes|miércoles|jueves|viernes")


def parse_loose_date(s, tz):
    """解析引用標頭裡的日期（各種語言）。標頭通常沒有時區，假設為 tz。解析失敗回傳 None。"""
    s = (s or "").strip()
    if not s:
        return None
    low = s.lower()
    mz = re.search(r"\b(" + "|".join(TZ_ABBR) + r")\b", low)
    mo_ = re.search(r"(?:gmt|utc)?\s*([+-])(\d{2}):?(\d{2})\b", low)
    if mz:
        tz = timezone(timedelta(hours=TZ_ABBR[mz.group(1)]))
    elif mo_ and re.search(r"\d{1,2}:\d{2}", low[:mo_.start()]):
        sign = 1 if mo_.group(1) == "+" else -1
        tz = timezone(sign * timedelta(hours=int(mo_.group(2)), minutes=int(mo_.group(3))))
    hh = mm = 0
    tm = re.search(r"(?<![\d./])(\d{1,2})[:：.](\d{2})(?![./]\d)(?:[:：]\d{2})?\s*(am|pm|a\.m\.|p\.m\.)?", low)
    pm_zh = re.search(r"(上午|下午|晚上|中午|凌晨)\s*(\d{1,2})[:：](\d{2})", s)
    y = mo = d = None
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", s) or re.search(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.search(rf"(\d{{1,2}})\.?\s+({MONTH_RE})\.?,?\s+(\d{{4}})", low)        # 5 ottobre 2026 / 5. Oktober 2026
        if m:
            d, mo, y = int(m.group(1)), MONTHS[m.group(2)], int(m.group(3))
        else:
            m = re.search(rf"({MONTH_RE})\.?\s+(\d{{1,2}}),?\s+(\d{{4}})", low)       # October 6, 2026
            if m:
                mo, d, y = MONTHS[m.group(1)], int(m.group(2)), int(m.group(3))
            else:
                m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", low)            # 6.10.2026（日.月.年）
                if m:
                    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                    if mo > 12:
                        d, mo = mo, d
    if not y:
        try:                      # 最後才用 RFC 解析（它會忽略 AM/PM，所以不能放前面）
            dt = parsedate_to_datetime(s)
            return dt if dt.tzinfo else dt.replace(tzinfo=tz)
        except Exception:
            return None
    if pm_zh:
        hh, mm = int(pm_zh.group(2)), int(pm_zh.group(3))
        if pm_zh.group(1) in ("下午", "晚上") and hh < 12:
            hh += 12
    elif tm:
        hh, mm = int(tm.group(1)), int(tm.group(2))
        ap = (tm.group(3) or "").replace(".", "")
        if ap == "pm" and hh < 12:
            hh += 12
        elif ap == "am" and hh == 12:
            hh = 0
    try:
        if not (mz or mo_) and EURO_HINT.search(low):
            # 義／德／法／西文標頭沒寫時區時，多半是歐洲當地時間（夏令 4–10 月 +2，其餘 +1，約略）
            tz = timezone(timedelta(hours=2 if 4 <= mo <= 10 else 1))
        return datetime(y, mo, d, hh, mm, tzinfo=tz)
    except ValueError:
        return None


def _clean_field(v):
    v = re.sub(r"<mailto:[^>]*>|\[mailto:[^\]]*\]?|mailto:\S+", "", v or "")
    return re.sub(r"\s+", " ", v).strip(" ;,'\"")


def _strip_quote_marks(lines):
    return [re.sub(r"^\s*(>\s?)+", "", l) for l in lines]


def _trim_separators(lines):
    while lines and (not lines[-1].strip() or SEPARATOR_RE.match(lines[-1])):
        lines = lines[:-1]
    while lines and not lines[0].strip():
        lines = lines[1:]
    return lines


def split_thread(quoted, tz):
    """把引用前文拆成 [{sender,to,cc,subject,date,date_raw,body}]，由新到舊（信串原本的順序）。"""
    lines = _strip_quote_marks((quoted or "").splitlines())
    starts = []          # (標頭起始行, 標頭結束後第一行, header dict)
    i = 0
    while i < len(lines):
        s = lines[i]
        mw = WROTE_RE.match(s) or WROTE_ZH.match(s)
        if mw:
            starts.append((i, i + 1, {"date_raw": mw.group(1), "sender": _clean_field(mw.group(2))}))
            i += 1
            continue
        mf = H_FROM.match(s)
        if mf and any(H_DATE.match(w) for w in lines[i + 1:i + 7]):
            h, last, j, seen_subj, blank = {"sender": mf.group(2)}, "sender", i + 1, False, False
            while j < len(lines) and j < i + 15:
                w = lines[j]
                if not w.strip():
                    nxt = next((x for x in lines[j + 1:j + 3] if x.strip()), "")
                    if seen_subj and "date_raw" in h and "to" in h and not H_OTHER.match(nxt):
                        break
                    blank = True
                    j += 1
                    continue
                if H_OTHER.match(w):           # Reply-To、Importance 之類，略過
                    blank, last = False, "_other"
                    j += 1
                    continue
                if blank and not any(rx.match(w) for rx in (H_DATE, H_TO, H_CC, H_SUBJ)):
                    break                      # 空行後不是標頭 → 標頭結束，這行是內文
                blank = False
                for key, rx in (("date_raw", H_DATE), ("to", H_TO), ("cc", H_CC), ("subject", H_SUBJ)):
                    mm = rx.match(w)
                    if mm:
                        h[key], last = mm.group(2), key
                        seen_subj = seen_subj or key == "subject"
                        break
                else:
                    if last in ("subject", "_other"):     # 主旨後面緊接的非標頭行 → 內文開始
                        break
                    h[last] = h.get(last, "") + " " + w.strip()     # 標頭換行的續行
                j += 1
            starts.append((i, j, {k: _clean_field(v) for k, v in h.items()}))
            i = j
            continue
        i += 1
    segs = []
    for n, (a, b, h) in enumerate(starts):
        end = starts[n + 1][0] if n + 1 < len(starts) else len(lines)
        body = "\n".join(_trim_separators(lines[b:end])).strip()
        body = re.sub(r"\n\s*\n\s*\n+", "\n\n", body)
        segs.append({"sender": h.get("sender", ""), "to": h.get("to", ""), "cc": h.get("cc", ""),
                     "subject": h.get("subject", ""), "date_raw": h.get("date_raw", ""),
                     "date": parse_loose_date(h.get("date_raw", ""), tz), "body": body})
    return segs


def fingerprint(text):
    """內文指紋：去掉連結、標點、空白後的前 160 字。同一封信被不同人轉寄／引用，指紋會相同。"""
    t = re.sub(r"<mailto:[^>]*>|\[mailto:[^\]]*\]|https?://\S+|mailto:\S+", "", text or "")
    t = "".join(ch for ch in t.lower() if ch.isalnum())
    return t[:160]


def sender_key(s):
    """寄件者比對用：取 email，沒有就取名字的英數字＋中文。"""
    m = re.search(r"[\w.+-]+@[\w-]+(\.[\w-]+)+", s or "")
    if m:
        return m.group(0).lower()
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())[:20]


SUBJECT_PREFIX = re.compile(
    r"^\s*((re|fw|fwd|aw|sv|答覆|答复|回覆|回复|轉寄|轉發|转寄|转发)\s*(\[\d+\])?\s*[:：]\s*)+", re.I
)


def normalize_subject(s):
    s = s or ""
    prev = None
    while prev != s:
        prev = s
        s = SUBJECT_PREFIX.sub("", s).strip()
    return s or "(無主旨)"


def to_dt(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = parsedate_to_datetime(str(value))
        except Exception:
            try:
                dt = datetime.fromisoformat(str(value))
            except Exception:
                return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt

# ---------- 各格式解析 ----------

def msg_ids(s):
    """從 In-Reply-To / References 抓出所有 <...> 形式的 Message-ID。"""
    s = str(s or "")
    found = re.findall(r"<[^<>\s]+>", s)
    return [x.lower() for x in found] or ([s.strip().lower()] if s.strip() else [])


def conv_key(raw):
    """Outlook 的 ConversationIndex（.eml 裡叫 Thread-Index）前 22 bytes 是整串信共用的識別碼。"""
    if not raw:
        return ""
    if isinstance(raw, str):
        try:
            raw = base64.b64decode(raw.strip() + "===")
        except Exception:
            return ""
    return bytes(raw[:22]).hex() if len(raw) >= 22 else ""


def parse_msg(path, keep=False):
    import extract_msg
    m = extract_msg.openMsg(str(path))
    try:
        hdr = getattr(m, "header", None)
        refs = hdr.get("References", "") if hdr is not None else ""
        try:
            cidx = m.getStream("__substg1.0_00710102")
        except Exception:
            cidx = None
        if not cidx and hdr is not None:
            cidx = hdr.get("Thread-Index", "")
        body = m.body or ""
        if not body.strip():
            body = html_to_text(getattr(m, "htmlBody", None))
        attachments, blobs = [], []
        for a in getattr(m, "attachments", []) or []:
            name = getattr(a, "longFilename", None) or getattr(a, "shortFilename", None) or getattr(a, "name", None)
            if name:
                attachments.append(str(name))
                if keep:
                    data = getattr(a, "data", None)
                    if isinstance(data, (bytes, bytearray)):
                        blobs.append((str(name), bytes(data)))
        return {
            "date": to_dt(getattr(m, "date", None)),
            "sender": str(getattr(m, "sender", "") or ""),
            "to": str(getattr(m, "to", "") or ""),
            "cc": str(getattr(m, "cc", "") or ""),
            "subject": str(getattr(m, "subject", "") or ""),
            "message_id": str(getattr(m, "messageId", "") or ""),
            "in_reply_to": msg_ids(getattr(m, "inReplyTo", "")),
            "references": msg_ids(refs),
            "conv": conv_key(cidx),
            "body": body,
            "attachments": attachments,
            "att_blobs": blobs,
        }
    finally:
        try:
            m.close()
        except Exception:
            pass


def decode_part(part):
    try:
        return part.get_content()
    except Exception:
        payload = part.get_payload(decode=True) or b""
        for enc in (part.get_content_charset() or "utf-8", "cp950", "big5", "utf-8", "latin-1"):
            try:
                return payload.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return payload.decode("utf-8", errors="replace")


def parse_eml(path, keep=False):
    with open(path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    body = ""
    plain = msg.get_body(preferencelist=("plain",))
    if plain is not None:
        body = decode_part(plain)
    if not body.strip():
        h = msg.get_body(preferencelist=("html",))
        if h is not None:
            body = html_to_text(decode_part(h))
    attachments, blobs = [], []
    for part in msg.iter_attachments():
        fn = part.get_filename()
        if not fn:
            continue
        attachments.append(fn)
        if keep:
            try:
                data = part.get_payload(decode=True)
            except Exception:
                data = None
            if isinstance(data, (bytes, bytearray)):
                blobs.append((fn, bytes(data)))
    return {
        "date": to_dt(msg.get("Date")),
        "sender": str(msg.get("From", "")),
        "to": str(msg.get("To", "")),
        "cc": str(msg.get("Cc", "")),
        "subject": str(msg.get("Subject", "")),
        "message_id": str(msg.get("Message-ID", "")),
        "in_reply_to": msg_ids(msg.get("In-Reply-To", "")),
        "references": msg_ids(msg.get("References", "")),
        "conv": conv_key(str(msg.get("Thread-Index", "") or "")),
        "body": body or "",
        "attachments": attachments,
        "att_blobs": blobs,
    }


def find_soffice():
    """找 LibreOffice。Windows、macOS 裝好後通常不在 PATH 裡，要到預設安裝位置找。"""
    hit = shutil.which("soffice") or shutil.which("libreoffice")
    if hit:
        return hit
    cands = [Path(os.environ[k]) / "LibreOffice" / "program" / "soffice.exe"
             for k in ("ProgramFiles", "ProgramFiles(x86)") if os.environ.get(k)]
    cands.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    return next((str(c) for c in cands if c.is_file()), None)


def expand_pst(path, errors):
    """用 readpst 把 .pst 展開成 .eml，回傳 (暫存目錄, eml 路徑清單)"""
    if not shutil.which("readpst"):
        errors.append(f"{path.name}: 找不到 readpst。" + (
            "Windows 沒有現成的 readpst，請在 Outlook 把這個 .pst 裡的信選起來，拖到資料夾存成 .msg" if os.name == "nt"
            else "請安裝 pst-utils（apt-get install -y pst-utils 或 brew install libpst），或在 Outlook 另存為 .msg"))
        return None, []
    tmp = Path(tempfile.mkdtemp(prefix="pst_"))
    r = subprocess.run(["readpst", "-e", "-q", "-o", str(tmp), str(path)], capture_output=True, text=True)
    if r.returncode != 0:
        errors.append(f"{path.name}: readpst 失敗：{r.stderr.strip()[:200]}")
        return tmp, []
    return tmp, sorted(tmp.rglob("*.eml"))

# ---------- 主程式 ----------

URL_RE = re.compile(r"https?://[^\s<>\"')\]」』）]+")


def extract_links(text):
    """抓出內文（新寫的部分）裡的外部連結，保留順序、去重。"""
    seen, out = set(), []
    for u in URL_RE.findall(text or ""):
        u = u.rstrip(".,;:、，。")
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


# ---------- 附件內容（--read-attachments） ----------

class MissingDep(Exception):
    pass


TEXT_EXT = {".txt", ".csv", ".tsv", ".md", ".log", ".json", ".xml"}
HTML_EXT = {".html", ".htm"}
LEGACY = {".doc": "docx", ".rtf": "docx", ".xls": "xlsx", ".ppt": "pptx"}   # 用 LibreOffice 轉成新格式再讀
SUPPORTED = {".pdf", ".docx", ".xlsx", ".xlsm", ".pptx"} | TEXT_EXT | HTML_EXT | set(LEGACY)
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg", ".emf", ".wmf"}
MAX_ATT_BYTES = 30 * 1024 * 1024
MAX_TEXT_CHARS = 200_000
SMALL_IMG = 30 * 1024          # 小圖片多半是簽名檔 logo，不在 emails.txt 列出


def safe_name(n):
    return re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", n).strip(" .") or "attachment"


def decode_bytes(b):
    if b.startswith((b"\xff\xfe", b"\xfe\xff")):
        return b.decode("utf-16", errors="replace")
    for enc in ("utf-8-sig", "cp950", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def ex_pdf(path):
    text, pages = "", None
    if shutil.which("pdftotext"):
        try:
            r = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, timeout=120)
            if r.returncode == 0:
                text = r.stdout.decode("utf-8", "replace")
                pages = text.count("\f") or None
        except Exception:
            pass
    if not text.strip():
        try:
            from pypdf import PdfReader
        except ImportError:
            if not shutil.which("pdftotext"):
                raise MissingDep("pypdf（pip install pypdf）")
            return "", ""
        rd = PdfReader(str(path))
        if rd.is_encrypted:
            try:
                ok = rd.decrypt("")
            except Exception:
                ok = 0
            if not ok:
                raise ValueError("PDF 已加密，無法讀取")
        pages = len(rd.pages)
        text = "\n".join((pg.extract_text() or "") for pg in rd.pages)
    return text, (f"{pages} 頁" if pages else "")


def ex_docx(path):
    try:
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        raise MissingDep("python-docx（pip install python-docx）")
    d = Document(str(path))
    out = []
    for ch in d.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            t = Paragraph(ch, d).text.strip()
            if t:
                out.append(t)
        elif ch.tag.endswith("}tbl"):
            for row in Table(ch, d).rows:
                cells, prev = [], None
                for c in row.cells:
                    if c._tc is prev:          # 合併儲存格只取一次
                        continue
                    prev = c._tc
                    cells.append(c.text.strip().replace("\n", " "))
                if any(cells):
                    out.append(" | ".join(cells))
    return "\n".join(out), ""


def ex_xlsx(path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise MissingDep("openpyxl（pip install openpyxl）")
    wb = load_workbook(str(path), read_only=True, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            out.append(f"## 工作表：{ws.title}")
            n = 0
            for row in ws.iter_rows(values_only=True):
                vals = []
                for v in row[:40]:
                    if v is None:
                        vals.append("")
                    elif hasattr(v, "strftime"):
                        vals.append(v.strftime("%Y-%m-%d"))
                    elif isinstance(v, float) and v.is_integer():
                        vals.append(str(int(v)))
                    else:
                        vals.append(str(v).strip())
                while vals and not vals[-1]:
                    vals.pop()
                if not any(vals):
                    continue
                out.append(" | ".join(vals))
                n += 1
                if n >= 500:
                    out.append("……（此工作表超過 500 列，已截斷）")
                    break
        meta = f"{len(wb.worksheets)} 個工作表"
    finally:
        wb.close()
    return "\n".join(out), meta


def ex_pptx(path):
    try:
        from pptx import Presentation
    except ImportError:
        raise MissingDep("python-pptx（pip install python-pptx）")
    prs = Presentation(str(path))
    out = []
    for i, sl in enumerate(prs.slides, 1):
        out.append(f"## 投影片 {i}")
        for sh in sl.shapes:
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    out.append(t)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    out.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
        if sl.has_notes_slide and sl.notes_slide.notes_text_frame is not None:
            nt = sl.notes_slide.notes_text_frame.text.strip()
            if nt:
                out.append("備註：" + nt)
    return "\n".join(out), f"{len(prs.slides)} 張投影片"


def extract_text(path, ext):
    if ext == ".pdf":
        return ex_pdf(path)
    if ext == ".docx":
        return ex_docx(path)
    if ext in (".xlsx", ".xlsm"):
        return ex_xlsx(path)
    if ext == ".pptx":
        return ex_pptx(path)
    if ext in TEXT_EXT:
        return decode_bytes(Path(path).read_bytes()), ""
    if ext in HTML_EXT:
        return html_to_text(Path(path).read_bytes()), ""
    if ext in LEGACY:
        soffice = find_soffice()
        if not soffice:
            raise MissingDep("LibreOffice（用來讀 .doc/.xls/.ppt 舊版檔案）")
        target = LEGACY[ext]
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", target, "--outdir", tmp, str(path)],
                           capture_output=True, timeout=180)
            conv = Path(tmp) / (Path(path).stem + "." + target)
            if not conv.exists():
                raise ValueError("舊版檔案轉檔失敗")
            return extract_text(conv, "." + target)
    raise ValueError("不支援的格式")


_VER = (r"v\s?\d+(?:\.\d+)*|ver\.?\s?\d+(?:\.\d+)*|rev\.?\s?[a-z0-9]+|r\d+|第\s?\d+\s?版|修正\s?\d*|修訂\s?\d*|改版\s?\d*|"
        r"final|最終版?|定稿|\d{8}|\d{4}[-_.]\d{1,2}[-_.]\d{1,2}")
VER_RE = re.compile(rf"(?:^|(?<=[\s_\-.(（]))(?:{_VER})(?=$|[\s_\-.)）])", re.I)


def version_hint(name):
    m = VER_RE.search(Path(name).stem)
    return m.group(0) if m else ""


def family_key(name):
    """去掉版本字樣後的檔名＋副檔名，同一組視為同系列附件（spec_v1.pdf、spec_v2.pdf）。"""
    base = re.sub(r"[\s_\-.()（）]+", "", VER_RE.sub("", Path(name).stem)).lower()
    return (base, Path(name).suffix.lower()) if len(base) >= 2 else None


def _norm_lines(text):
    return [re.sub(r"\s+", " ", l).strip() for l in text.splitlines() if l.strip()]


def process_attachments(pairs, out, excerpt_chars, subdir="attachments"):
    """把附件存到 <工作目錄>/<subdir>/<編號>/，抽出文字，並對同系列附件產生版本差異。
    信件附件用 attachments/；資料夾裡的其他檔案（D001…）用 documents/。"""
    adir = out / subdir
    if adir.exists():
        shutil.rmtree(adir, ignore_errors=True)
    adir.mkdir(parents=True, exist_ok=True)
    seen, families, texts = {}, defaultdict(list), {}
    stats, missing, problems = Counter(), set(), []

    for item, blobs in pairs:
        details = []
        for name, data in blobs:
            ext = Path(name).suffix.lower()
            sha = hashlib.sha1(data).hexdigest()
            d = {"name": name, "ext": ext, "size": len(data), "sha1": sha[:10], "status": "", "note": "",
                 "meta": "", "version_hint": version_hint(name), "saved": "", "text_file": "",
                 "text_chars": 0, "excerpt": ""}
            details.append(d)
            if ext in IMG_EXT:
                d.update(status="image", note="圖片，未讀取內容")
                stats["image"] += 1
                continue
            if sha in seen:
                d.update(status="duplicate", duplicate_of=seen[sha][0],
                         note=f"與 {seen[sha][0]}《{seen[sha][1]}》內容相同，不重複讀取")
                stats["duplicate"] += 1
                continue
            seen[sha] = (item["id"], name)
            if len(data) > MAX_ATT_BYTES:
                d.update(status="too_large", note=f"檔案過大（{len(data) // 1048576} MB），未讀取")
                stats["skipped"] += 1
                continue
            folder = adir / item["id"]
            folder.mkdir(exist_ok=True)
            fpath, k = folder / safe_name(name), 1
            while fpath.exists():
                fpath = folder / f"{k}_{safe_name(name)}"
                k += 1
            fpath.write_bytes(data)
            d["saved"] = str(fpath.relative_to(out)).replace("\\", "/")
            if ext not in SUPPORTED:
                d.update(status="unsupported", note=f"{ext or '無副檔名'} 格式未讀取內容（原檔存在 {d['saved']}，圖檔／CAD 需另行檢視）")
                stats["skipped"] += 1
                continue
            try:
                text, meta = extract_text(fpath, ext)
            except MissingDep as ex:
                missing.add(str(ex))
                d.update(status="missing_dep", note=f"缺少 {ex}，無法讀取")
                stats["error"] += 1
                continue
            except Exception as ex:
                d.update(status="error", note=f"讀取失敗：{type(ex).__name__}: {str(ex)[:80]}")
                problems.append(f"{item['id']} {name}: {d['note']}")
                stats["error"] += 1
                continue
            text = re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+\n", "\n", (text or "").replace("\r\n", "\n"))).strip()
            if not text:
                d.update(status="empty", note=("沒有可讀的文字（PDF 可能是掃描檔，內容在圖片裡，需用圖片檢視：原檔存在 " + d["saved"] + "）")
                         if ext == ".pdf" else "檔案沒有可讀的文字內容")
                stats["empty"] += 1
                continue
            body = text[:MAX_TEXT_CHARS] + ("\n……（內容過長已截斷）" if len(text) > MAX_TEXT_CHARS else "")
            tf = Path(str(fpath) + ".txt")
            tf.write_text(body, encoding="utf-8")
            d.update(status="ok", meta=meta, text_file=str(tf.relative_to(out)).replace("\\", "/"),
                     text_chars=len(text), excerpt=text[:excerpt_chars])
            texts[id(d)] = text
            stats["ok"] += 1
            key = family_key(name)
            if key:
                families[key].append((item, d))
        item["attachment_details"] = details

    # 同系列附件（不同版本）逐版比對
    ddir, n_diff = adir / "diffs", 0
    for key, members in families.items():
        for (it0, d0), (it1, d1) in zip(members, members[1:]):
            a, b = _norm_lines(texts[id(d0)]), _norm_lines(texts[id(d1)])
            if a == b:
                continue
            dl = [l for l in difflib.unified_diff(a, b, n=0, lineterm="") if not l.startswith(("---", "+++", "@@"))]
            added, removed = sum(l.startswith("+") for l in dl), sum(l.startswith("-") for l in dl)
            ddir.mkdir(exist_ok=True)
            n_diff += 1
            df = ddir / f"{n_diff:02d}_{safe_name(key[0])[:30]}_{it0['id']}-{it1['id']}.diff.txt"
            head = f"{it0['id']}《{d0['name']}》 → {it1['id']}《{d1['name']}》\n（+ 新版新增／變更的行，- 舊版被刪除／變更的行）\n\n"
            df.write_text(head + "\n".join(dl[:3000]) + ("\n……（差異過多已截斷）" if len(dl) > 3000 else ""), encoding="utf-8")
            d1["diff_from"] = {"email_id": it0["id"], "name": d0["name"], "added": added, "removed": removed,
                               "file": str(df.relative_to(out)).replace("\\", "/"),
                               "preview": [l[:160] for l in dl[:20]]}
    return {"total": sum(stats.values()), "ok": stats["ok"], "duplicate": stats["duplicate"],
            "image": stats["image"], "empty": stats["empty"], "skipped": stats["skipped"],
            "error": stats["error"], "diffs": n_diff, "missing": sorted(missing), "problems": problems[:10]}


def write_att_details(f, x, label="附件內容", keep_small_images=False):
    shown = [d for d in x.get("attachment_details", [])
             if keep_small_images or not (d["status"] == "image" and d["size"] < SMALL_IMG)]
    if not shown:
        return
    f.write(f"{label}：\n")
    for d in shown:
        tag = f"［{d['version_hint']}］" if d.get("version_hint") else ""
        if d["status"] == "ok":
            f.write(f"  ▸ {d['name']} {tag}　{d['meta']}　約 {d['text_chars']:,} 字\n    全文：{d['text_file']}\n")
            ex = d["excerpt"].replace("\n", " ⏎ ")
            f.write(f"    摘錄：{ex}{'…' if d['text_chars'] > len(d['excerpt']) else ''}\n")
            df = d.get("diff_from")
            if df:
                f.write(f"    與 {df['email_id']}《{df['name']}》相比：新增 {df['added']} 行、刪除 {df['removed']} 行（完整差異：{df['file']}）\n")
                for l in df["preview"]:
                    f.write("      " + l + "\n")
                if df["added"] + df["removed"] > len(df["preview"]):
                    f.write("      ……\n")
        else:
            f.write(f"  ▸ {d['name']} {tag}　（{d['note']}）\n")


# ---------- 系統信、人物表、上次報告 ----------

SYSTEM_SENDER_KW = ["no-reply", "noreply", "no_reply", "donotreply", "do-not-reply",
                    "do_not_reply", "mailer-daemon", "postmaster"]
# 注意：不要放「通知」「notification」——「出貨通知」是正式業務信，不能被排除
SYSTEM_SUBJECT_KW = ["自動回覆", "自动回复", "automatic reply", "auto-reply", "autoreply",
                     "out of office", "undeliverable", "delivery status notification",
                     "無法傳遞", "投遞失敗", "退信"]


def load_exclude(path):
    kws = []
    if path:
        for line in Path(path).read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                kws.append(line.lower())
    return kws


def system_reason(sender, subject, extra):
    s, t = (sender or "").lower(), (subject or "").lower()
    for k in SYSTEM_SENDER_KW + extra:
        if k in s:
            return f"寄件者含「{k}」"
    for k in SYSTEM_SUBJECT_KW:
        if k in t:
            return f"主旨含「{k}」"
    return ""


def split_addrs(s):
    """把寄件者/收件者字串拆成 [(顯示名稱, email)]。支援分號、逗號分隔，以及只有名字沒有 email 的情況。"""
    s = s or ""
    if not s.strip():
        return []
    parts = [p for p in s.split(";") if p.strip()] if ";" in s else [s]
    out = []
    for part in parts:
        for name, addr in getaddresses([part]):
            name, addr = (name or "").strip().strip('"'), (addr or "").strip()
            if not addr and not name:
                continue
            if addr and "@" not in addr:      # 只有名字
                name, addr = (name or addr), ""
            out.append((name, addr.lower()))
    return out


def build_people(items, prev_people):
    ppl = {}

    def rec(name, addr):
        key = addr or ("name:" + name.lower())
        r = ppl.setdefault(key, {
            "key": key, "emails": [addr] if addr else [], "names": [],
            "domain": addr.split("@")[-1] if addr else "",
            "sent": 0, "received": 0, "first": "", "last": "", "signature_hint": "",
        })
        if name and name not in r["names"]:
            r["names"].append(name)
        return r

    for x in items:
        if x["system"]:
            continue
        d = x["date_display"][:10] if x["date"] else ""
        for nm, ad in split_addrs(x["sender"]):
            r = rec(nm, ad)
            r["sent"] += 1
            if d:
                r["first"] = r["first"] or d
                r["last"] = d
            lines = [l.strip() for l in (x["body"] or "").splitlines() if l.strip()]
            if len(lines) >= 3:
                r["signature_hint"] = " / ".join(lines[-4:])[:160]
        for nm, ad in split_addrs(x["to"]) + split_addrs(x["cc"]):
            rec(nm, ad)["received"] += 1

    # 引用標頭常只有名字沒有 email：名字（去空白後）只對到一個有 email 的人，就直接併過去
    squash = lambda n: re.sub(r"\s+", "", n.lower())
    for k in [k for k in ppl if k.startswith("name:")]:
        r = ppl[k]
        cands = [a for a, x in ppl.items() if not a.startswith("name:")
                 and {squash(n) for n in x["names"]} & {squash(n) for n in r["names"]}]
        if len(cands) != 1:
            continue
        t = ppl[cands[0]]
        t["sent"] += r["sent"]
        t["received"] += r["received"]
        t["names"] += [n for n in r["names"] if n not in t["names"]]
        t["first"] = min(x for x in (t["first"], r["first"]) if x) if (t["first"] or r["first"]) else ""
        t["last"] = max(t["last"], r["last"])
        t["signature_hint"] = t["signature_hint"] or r["signature_hint"]
        del ppl[k]

    # 同一個顯示名稱出現在多個位址 → 可能是同一人（需使用者/Claude 確認）
    by_name = defaultdict(set)
    for k, r in ppl.items():
        for n in r["names"]:
            by_name[n.lower()].add(k)
    alias_candidates = [{"name": n, "keys": sorted(ks)} for n, ks in by_name.items() if len(ks) > 1]

    # 帶入上次報告已確認的人物資料
    for pp in prev_people or []:
        emails = {e.lower() for e in pp.get("emails", [])}
        names = {n.lower() for n in [pp.get("name", "")] + list(pp.get("aliases", [])) if n}
        for r in ppl.values():
            if "confirmed" in r:
                continue
            if emails & set(r["emails"]) or names & {n.lower() for n in r["names"]}:
                r["confirmed"] = {k: pp.get(k) for k in ("name", "aliases", "org", "role", "can_decide") if pp.get(k) is not None}

    people = sorted(ppl.values(), key=lambda r: -(r["sent"] + r["received"]))
    return people, alias_candidates


def assign_threads(items):
    """用信件標頭把信串起來：Message-ID ↔ In-Reply-To / References，加上 Outlook 的 ConversationIndex。
    從引用前文拆出來的信，跟它所在的那封信同一串。主旨改了也不會斷。
    每封信得到 thread（T01、T02…，依該串最早的信排序）與 thread_size。編號每次重算，不跨報告沿用。"""
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    by_id = {it["id"]: it for it in items}
    for it in items:
        node = "e:" + it["id"]
        find(node)
        for kind, val in it.pop("_links"):
            union(node, f"{kind}:{val}")
        if it["container"] in by_id:
            union(node, "e:" + it["container"])
    groups = defaultdict(list)
    for it in items:
        groups[find("e:" + it["id"])].append(it)
    ordered = sorted(groups.values(), key=lambda g: min(x["date"] or "9999" for x in g))
    for n, g in enumerate(ordered, 1):
        for it in g:
            it["thread"] = f"T{n:02d}"
            it["thread_size"] = len(g)
    return len(ordered)


def find_prev(folder):
    """找資料夾中最新的上次報告資料檔（build_report.py 產出的 案卷整理_*.data.json）。"""
    cands = sorted(folder.glob("案卷整理_*.data.json"))
    if not cands:
        return None, None
    try:
        return cands[-1], json.loads(cands[-1].read_text(encoding="utf-8"))
    except Exception:
        return cands[-1], None

# ---------- 資料夾裡的其他檔案（PDF、Word、Excel…，不是信件） ----------

EMAIL_EXT = (".msg", ".eml", ".pst")
DOC_SKIP_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}


def _is_under(p, root):
    try:
        p.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def list_documents(folder, pattern, out):
    """信件以外的檔案：排除本 skill 的報告、Office 暫存檔（~$）、隱藏檔與工作目錄。"""
    docs = []
    for p in folder.glob(pattern):
        if not p.is_file() or p.suffix.lower() in EMAIL_EXT:
            continue
        if p.name.startswith(("案卷整理_", "~$", ".")) or p.name.lower() in DOC_SKIP_NAMES:
            continue
        if any(part.startswith(".") for part in p.relative_to(folder).parts[:-1]) or _is_under(p, out):
            continue
        docs.append(p)
    return docs


def collect_documents(paths, folder, out, prev, show_tz, excerpt_chars):
    """每個檔案給一個 D 編號（依修改時間排序）。增量模式下，內容沒變的檔案沿用上次的編號。"""
    prev_map = {d["uid"]: d["id"] for d in (prev or {}).get("document_index", []) if d.get("uid") and d.get("id")}
    next_n = max([int(i[1:]) for i in prev_map.values() if i[1:].isdigit()] or [0])
    entries = []
    for p in paths:
        st = p.stat()
        rel = str(p.relative_to(folder)).replace("\\", "/")
        ext = p.suffix.lower()
        data = None
        if ext not in IMG_EXT and st.st_size <= MAX_ATT_BYTES:
            data = p.read_bytes()
            uid = "d" + hashlib.sha1(data).hexdigest()[:11]
        else:
            uid = "d" + hashlib.sha1(f"{rel}|{st.st_size}|{int(st.st_mtime)}".encode("utf-8", "ignore")).hexdigest()[:11]
        entries.append((st.st_mtime, rel, p, ext, st.st_size, uid, data))
    entries.sort(key=lambda e: (e[0], e[1]))

    items, pairs, used = [], [], set()
    for i, (mtime, rel, p, ext, size, uid, data) in enumerate(entries, 1):
        old = prev_map.get(uid) if prev else None
        if old and old not in used:
            did, is_new = old, False
        elif prev:
            next_n += 1
            did, is_new = f"D{next_n:03d}", True
        else:
            did, is_new = f"D{i:03d}", False
        used.add(did)
        local = datetime.fromtimestamp(mtime, tz=timezone.utc).astimezone(show_tz)
        it = {"id": did, "uid": uid, "is_new": is_new, "file": rel, "name": p.name, "ext": ext, "size": size,
              "modified": local.isoformat(), "modified_display": local.strftime("%Y-%m-%d %H:%M"),
              "attachment_details": []}
        items.append(it)
        base = {"name": p.name, "ext": ext, "size": size, "sha1": "", "status": "", "note": "", "meta": "",
                "version_hint": version_hint(p.name), "saved": "", "text_file": "", "text_chars": 0, "excerpt": ""}
        if ext in IMG_EXT:
            it["attachment_details"] = [dict(base, status="image", note="圖片，未讀取內容；需要時用 view 直接看原檔")]
        elif data is None:
            it["attachment_details"] = [dict(base, status="too_large", note=f"檔案過大（{size // 1048576} MB），未讀取")]
        else:
            pairs.append((it, [(p.name, data)]))
    stats = process_attachments(pairs, out, excerpt_chars, subdir="documents") if pairs else None
    return items, stats


# ---------- 主程式 ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--out", required=True)
    ap.add_argument("--no-recursive", action="store_true")
    ap.add_argument("--tz-offset", type=float, default=8, help="顯示時區，預設台灣 UTC+8")
    ap.add_argument("--exclude-file", help="額外的系統信寄件者關鍵字，一行一個（# 開頭為註解）")
    ap.add_argument("--prev", help="指定上次報告資料檔（*.data.json）；預設自動找信件資料夾裡最新的")
    ap.add_argument("--no-prev", action="store_true", help="不要和上次報告比對，視為全新整理")
    ap.add_argument("--read-attachments", action="store_true",
                    help="讀取附件內容（PDF、Word、Excel、PowerPoint、文字檔），並比對同系列附件的版本差異")
    ap.add_argument("--attach-chars", type=int, default=500, help="emails.txt 中每個附件的摘錄字數，預設 500")
    ap.add_argument("--no-split", action="store_true", help="不要把引用前文拆成獨立信件（舊版行為）")
    ap.add_argument("--no-documents", action="store_true",
                    help="不要讀資料夾裡信件以外的檔案（PDF、Word、Excel、PowerPoint、文字檔…）")
    args = ap.parse_args()
    try:   # Windows 主控台預設 cp950，印到「［」這類字會當掉
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    show_tz = timezone(timedelta(hours=args.tz_offset))

    folder = Path(args.folder)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    if not folder.is_dir():
        sys.exit(f"找不到資料夾：{folder}")
    extra_kw = load_exclude(args.exclude_file)

    pattern = "*" if args.no_recursive else "**/*"
    files = [p for p in folder.glob(pattern) if p.is_file() and p.suffix.lower() in EMAIL_EXT]
    files = [p for p in files if not p.name.startswith("案卷整理_")]   # 排除本 skill 產出的報告
    doc_paths = [] if args.no_documents else list_documents(folder, pattern, out)

    errors, raw, tmpdirs = [], [], []
    for p in sorted(files):
        rel = str(p.relative_to(folder)).replace("\\", "/")
        ext = p.suffix.lower()
        try:
            if ext == ".msg":
                raw.append((rel, parse_msg(p, args.read_attachments)))
            elif ext == ".eml":
                raw.append((rel, parse_eml(p, args.read_attachments)))
            elif ext == ".pst":
                tmp, emls = expand_pst(p, errors)
                if tmp:
                    tmpdirs.append(tmp)
                for e in emls:
                    try:
                        raw.append((f"{rel} › {e.relative_to(tmp)}", parse_eml(e, args.read_attachments)))
                    except Exception as ex:
                        errors.append(f"{rel} › {e.name}: {ex}")
        except Exception as ex:
            errors.append(f"{rel}: {type(ex).__name__}: {ex}")

    # 第一層去重（同一個檔案存了兩次）
    seen, tops, dup = set(), [], 0
    for rel, d in raw:
        new_text, quoted = split_quoted(d["body"].replace("\r\n", "\n"))
        mid = d["message_id"].strip()
        if mid:
            key = "mid:" + mid
        else:
            dt = d["date"].strftime("%Y%m%d%H%M") if d["date"] else ""
            key = "h:" + hashlib.sha1(
                (normalize_subject(d["subject"]) + d["sender"] + dt + new_text[:200]).encode("utf-8", "ignore")
            ).hexdigest()
        if key in seen:
            dup += 1
            continue
        seen.add(key)
        tops.append({"rel": rel, "d": d, "body": new_text, "quoted": quoted, "src": "top",
                     "uid": hashlib.sha1(key.encode("utf-8", "ignore")).hexdigest()[:12]})

    # 拆信串：引用前文裡的每一封舊信變成獨立的一封（src = quote）
    quotes = []
    if not args.no_split:
        for t in tops:
            for seg in split_thread(t["quoted"], show_tz):
                if not seg["body"] and not seg["sender"]:
                    continue
                fp = fingerprint(seg["body"])
                day = seg["date"].strftime("%Y%m%d") if seg["date"] else ""
                uid = "q" + hashlib.sha1((fp + sender_key(seg["sender"]) + day).encode("utf-8", "ignore")).hexdigest()[:11]
                quotes.append({"rel": t["rel"], "src": "quote", "uid": uid, "container_uid": t["uid"],
                               "body": seg["body"], "quoted": "", "date_raw": seg["date_raw"],
                               "d": {"date": seg["date"], "sender": seg["sender"], "to": seg["to"], "cc": seg["cc"],
                                     "subject": seg["subject"], "message_id": "", "in_reply_to": [],
                                     "references": [], "conv": "", "attachments": [], "att_blobs": []}})

    # 跨檔去重：同一封信可能是 A 檔的本文、又出現在 B、C 檔的引用裡。
    # 比對內文指紋＋寄件時間差 2 天內（引用標頭的時間常被換成轉寄者的時區）；本文版優先保留。
    def same(a, b):
        fa, fb = a["_fp"], b["_fp"]
        if not fa or fa != fb:
            if not (len(fa) >= 60 and len(fb) >= 60 and (fa.startswith(fb[:60]) or fb.startswith(fa[:60]))):
                return False
        da, db = a["d"]["date"], b["d"]["date"]
        if da and db and abs((da - db).total_seconds()) > 2 * 86400:
            return False
        if len(fa) < 40:          # 內文太短（如「Thanks」）時，寄件者也要相同
            return sender_key(a["d"]["sender"])[:6] == sender_key(b["d"]["sender"])[:6]
        return True

    kept, xdup = [], 0
    for m in tops + quotes:
        m["_fp"] = fingerprint(m["body"])
        m["also_in"], m["alt_uids"] = [], []
        hit = next((k for k in kept if same(k, m)), None) if m["src"] == "quote" else None
        if hit is None:
            kept.append(m)
            continue
        xdup += 1
        if m["rel"] != hit["rel"] and m["rel"] not in hit["also_in"]:
            hit["also_in"].append(m["rel"])
        if m["uid"] not in hit["alt_uids"] and m["uid"] != hit["uid"]:
            hit["alt_uids"].append(m["uid"])
        if hit["src"] == "quote" and not hit["d"]["date"] and m["d"]["date"]:
            hit["d"]["date"] = m["d"]["date"]

    far_future = datetime.max.replace(tzinfo=timezone.utc)
    kept.sort(key=lambda x: x["d"]["date"] or far_future)

    # 上次報告：沿用舊編號（E001 不會因為插入新信而變動），只有新信往後接號
    prev_path, prev = (None, None)
    if not args.no_prev:
        if args.prev:
            prev_path = Path(args.prev)
            try:
                prev = json.loads(prev_path.read_text(encoding="utf-8"))
            except Exception as ex:
                errors.append(f"上次報告資料檔讀取失敗：{ex}")
        else:
            prev_path, prev = find_prev(folder)
    prev_idx = (prev or {}).get("email_index", [])
    prev_map = {}
    for e in prev_idx:
        for u in [e.get("uid")] + list(e.get("alt_uids") or []):
            if u and e.get("id"):
                prev_map[u] = e["id"]
    prev_warn = ""
    if prev and not prev_map:
        prev_warn = "上次報告資料檔沒有信件索引（email_index），無法比對新舊信件，視為全新整理。"
        prev = None
    elif prev and not args.no_split and not any("src" in e for e in prev_idx):
        prev_warn = ("上次報告是舊版格式（沒有拆信串），這次會把引用裡的舊信都當成新信。"
                     "建議加 --no-prev 重新整理一次，之後就能正常增量比對。")
    next_n = max([int(i[1:]) for i in prev_map.values() if i[1:].isdigit()] or [0])

    items, by_uid = [], {}
    for i, m in enumerate(kept, 1):
        d, new_text = m["d"], m["body"]
        dt = d["date"]
        local = dt.astimezone(show_tz) if dt else None
        old = next((prev_map[u] for u in [m["uid"]] + m["alt_uids"] if u in prev_map), None) if prev else None
        if old:
            eid, is_new = old, False
        elif prev:
            next_n += 1
            eid, is_new = f"E{next_n:03d}", True
        else:
            eid, is_new = f"E{i:03d}", False
        reason = system_reason(d["sender"], d["subject"], extra_kw)
        it = {
            "id": eid,
            "uid": m["uid"],
            "alt_uids": m["alt_uids"],
            "src": m["src"],
            "is_new": is_new,
            "system": bool(reason),
            "system_reason": reason,
            "file": m["rel"],
            "also_in": m["also_in"],
            "date": local.isoformat() if local else "",
            "date_display": local.strftime("%Y-%m-%d %H:%M") if local else "日期不明",
            "date_approx": m["src"] == "quote",
            "date_raw": m.get("date_raw", ""),
            "sender": d["sender"],
            "to": d["to"],
            "cc": d["cc"],
            "subject": d["subject"],
            "subject_norm": normalize_subject(d["subject"]),
            "attachments": d["attachments"],
            "links": extract_links(new_text),
            "body": new_text[:BODY_LIMIT],
            "body_truncated": len(new_text) > BODY_LIMIT,
            "body_file": "",
            "_full": new_text if len(new_text) > BODY_LIMIT else None,
            "quoted_excerpt": m["quoted"][:QUOTE_EXCERPT_LIMIT],
            "_blobs": d.get("att_blobs", []),
            "_container_uid": m.get("container_uid"),
            "_links": ([("mid", d["message_id"].strip().lower())] if d.get("message_id", "").strip() else [])
                      + [("mid", r) for r in d.get("in_reply_to", []) + d.get("references", [])]
                      + ([("conv", d["conv"])] if d.get("conv") else []),
        }
        items.append(it)
        by_uid[m["uid"]] = it
    for it in items:
        cu = it.pop("_container_uid")
        it["container"] = by_uid[cu]["id"] if cu in by_uid else ""
    thread_count = assign_threads(items)
    if prev:     # 增量模式下新舊編號交錯，報告仍依時間排序
        items.sort(key=lambda x: x["date"] or "9999")
    quote_count = sum(1 for x in items if x["src"] == "quote")

    att_stats = None
    if args.read_attachments:
        att_stats = process_attachments([(it, it["_blobs"]) for it in items], out, args.attach_chars)
    long_count = 0
    for it in items:
        it.pop("_blobs", None)
        full = it.pop("_full", None)
        if full:   # 內文超過 BODY_LIMIT：emails.json／emails.txt 只放開頭，全文另存，判讀與出處核對都用全文
            (out / "bodies").mkdir(exist_ok=True)
            (out / "bodies" / f"{it['id']}.txt").write_text(full, encoding="utf-8")
            it["body_file"], it["body_chars"] = f"bodies/{it['id']}.txt", len(full)
            long_count += 1

    documents, doc_stats = collect_documents(doc_paths, folder, out, prev, show_tz, args.attach_chars)
    new_doc_count = sum(1 for d in documents if d["is_new"])

    dated = [x["date_display"][:10] for x in items if x["date"]]
    new_count = sum(1 for x in items if x["is_new"])
    sys_count = sum(1 for x in items if x["system"])
    prev_report = (prev or {}).get("report", {})
    prev_info = None
    if prev:
        prev_info = {"file": prev_path.name, "generated_date": prev_report.get("generated_date", ""),
                     "email_count": len(prev_map), "new_count": new_count}
    people, alias_candidates = build_people(items, prev_report.get("people"))

    result = {
        "folder": str(folder),
        "parsed_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "file_count": len(files),
        "email_count": len(items),
        "duplicates_removed": dup + xdup,
        "top_count": len(items) - quote_count,
        "quote_count": quote_count,
        "cross_file_duplicates": xdup,
        "system_count": sys_count,
        "thread_count": thread_count,
        "attachment_stats": att_stats,
        "date_range": [dated[0], dated[-1]] if dated else [],
        "prev": prev_info,
        "errors": errors,
        "emails": items,
        "document_count": len(documents),
        "document_stats": doc_stats,
        "documents": documents,
    }
    (out / "emails.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "people.json").write_text(json.dumps(
        {"people": people, "alias_candidates": alias_candidates}, ensure_ascii=False, indent=2), encoding="utf-8")

    with open(out / "emails.txt", "w", encoding="utf-8") as f:
        if prev_info:
            f.write(f"# 上次報告：{prev_info['generated_date']}（{prev_info['file']}）；本次新增 {new_count} 封，標 【新】\n\n")
        for x in items:
            tag = " 【新】" if x["is_new"] else ""
            if x["system"]:
                f.write(f"===== {x['id']} | {x['date_display']} | 系統信（略）{tag} =====\n"
                        f"寄件者：{x['sender']}｜主旨：{x['subject']}｜原因：{x['system_reason']}\n\n")
                continue
            approx = "（引用標頭時間，時區可能不準）" if x["date_approx"] else ""
            f.write(f"===== {x['id']} | {x['date_display']}{approx}{tag} =====\n")
            if x["src"] == "quote":
                f.write(f"來源：{x['container']} 的引用前文（{x['file']}）；原始標頭時間：{x['date_raw'] or '無'}\n")
            f.write(f"寄件者：{x['sender']}\n收件者：{x['to']}\n")
            if x["cc"]:
                f.write(f"副本：{x['cc']}\n")
            f.write(f"主旨：{x['subject']}\n檔案：{x['file']}\n")
            if x["thread_size"] > 1:
                f.write(f"信串：{x['thread']}（共 {x['thread_size']} 封）\n")
            if x["also_in"]:
                f.write(f"也出現在：{'、'.join(x['also_in'])}\n")
            if x["attachments"]:
                f.write(f"附件：{', '.join(x['attachments'])}\n")
            if args.read_attachments and (not prev or x["is_new"]):
                write_att_details(f, x)
            if x["links"]:
                f.write(f"連結：{' | '.join(x['links'])}\n")
            f.write("\n" + (x["body"] or "（沒有新內容，可能是純轉寄）") + "\n")
            if x["body_truncated"]:
                f.write(f"……（內文過長，這裡只到前 {BODY_LIMIT} 字。全文 {x.get('body_chars', '?')} 字在 "
                        f"{out / x['body_file'] if x.get('body_file') else '（未存）'}，判讀這封信前要把全文讀完）\n")
            if x["quoted_excerpt"] and len(x["body"]) < 80 and args.no_split:
                f.write("\n[引用前文摘錄]\n" + x["quoted_excerpt"] + "\n")
            f.write("\n")

        if documents:
            f.write(f"##### 資料夾裡的其他檔案（不是信件）：共 {len(documents)} 個"
                    + (f"，本次新增或內容有變 {new_doc_count} 個，標 【新】" if prev_info else "")
                    + "。時間是檔案修改時間，不是寄出時間。\n\n")
            for d in documents:
                tag = " 【新】" if d["is_new"] else ""
                if prev_info and not d["is_new"]:
                    f.write(f"===== {d['id']} | {d['modified_display']} | {d['file']}（上次已讀過，內容沒變） =====\n\n")
                    continue
                f.write(f"===== {d['id']} | 修改時間 {d['modified_display']}{tag} =====\n檔案：{d['file']}\n")
                write_att_details(f, d, label="檔案內容", keep_small_images=True)
                f.write("\n")

    for t in tmpdirs:
        shutil.rmtree(t, ignore_errors=True)

    print(f"讀取檔案：{len(files)} 個；信件：{len(items)} 封（檔案本文 {len(items) - quote_count}、從引用前文拆出 {quote_count}）；"
          f"去除重複：{dup + xdup} 封（跨檔重複 {xdup}）；系統信（已標記）：{sys_count} 封")
    if long_count:
        print(f"長信：{long_count} 封內文超過 {BODY_LIMIT} 字，emails.txt 只列開頭，全文存在 {out / 'bodies'}（判讀時要讀全文）")
    nodate = sum(1 for x in items if not x["date"])
    if nodate:
        print(f"注意：{nodate} 封引用信讀不到日期，排在最後，請人工判斷時間")
    if dated:
        print(f"期間：{dated[0]} 至 {dated[-1]}")
    if prev_info:
        print(f"\n【增量模式】比對上次報告 {prev_info['generated_date']}：新增 {new_count} 封"
              + (("（沒有新信，也沒有新檔案，不需要重新整理）" if new_doc_count == 0
                  else f"；但資料夾裡有 {new_doc_count} 個新增或內容有變的檔案")
                 if new_count == 0 else "，舊信沿用原編號"))
    if prev_warn:
        print("\n注意：" + prev_warn)
    if att_stats:
        print(f"\n【附件】共 {att_stats['total']} 個：已讀取 {att_stats['ok']}、重複 {att_stats['duplicate']}、圖片 {att_stats['image']}、"
              f"無文字 {att_stats['empty']}、未支援 {att_stats['skipped']}、失敗 {att_stats['error']}；版本差異 {att_stats['diffs']} 組"
              + ("（增量模式只在 emails.txt 顯示新信的附件內容）" if prev_info else ""))
        for m in att_stats["missing"]:
            print(f"  缺少套件：{m}")
        for pr in att_stats["problems"]:
            print("  - " + pr)
        print(f"  附件與抽出的文字存在：{out/'attachments'}")
    if documents:
        ds = doc_stats or {}
        print(f"\n【資料夾裡的其他檔案】共 {len(documents)} 個（D001 起編號）：已讀取 {ds.get('ok', 0)}、"
              f"內容相同 {ds.get('duplicate', 0)}、無文字 {ds.get('empty', 0)}、未支援 {ds.get('skipped', 0)}、"
              f"失敗 {ds.get('error', 0)}；圖片與過大的檔案 "
              f"{sum(1 for d in documents for x in d['attachment_details'] if x['status'] in ('image', 'too_large'))}；"
              f"版本差異 {ds.get('diffs', 0)} 組"
              + (f"；本次新增或內容有變 {new_doc_count} 個" if prev_info else ""))
        for m in ds.get("missing", []):
            print(f"  缺少套件：{m}")
        for pr in ds.get("problems", []):
            print("  - " + pr)
        print(f"  抽出的文字存在：{out/'documents'}")
    print("\n主旨分組統計（分議題的起點，不是答案；不含系統信）：")
    for subj, n in Counter(x["subject_norm"] for x in items if not x["system"]).most_common(30):
        print(f"  {n:>3} 封  {subj}")
    threads = defaultdict(list)
    for x in items:
        if not x["system"]:
            threads[x["thread"]].append(x)
    multi = sorted((g for g in threads.values() if len(g) > 1), key=len, reverse=True)
    loose = sum(1 for g in threads.values() if len(g) == 1)
    print(f"\n信串統計（依信件標頭的回覆關係串起來，不看主旨；不含系統信）：{len(multi)} 串，另有 {loose} 封單獨的信")
    for g in multi[:30]:
        subs = list(dict.fromkeys(x["subject_norm"] for x in g if x["subject_norm"]))
        dts = sorted(x["date_display"][:10] for x in g if x["date"])
        span = f"{dts[0]}~{dts[-1]}" if dts else "日期不明"
        flag = "［主旨換過］" if len(subs) > 1 else ""
        print(f"  {g[0]['thread']}  {len(g):>3} 封  {span}  {' ／ '.join(subs[:3])}{' …' if len(subs) > 3 else ''} {flag}")
    if items and not multi:
        print("  （沒有串起來的信：可能是存檔缺回覆標頭，這時分議題只能靠主旨和內容）")
    print("\n主要人物（寄件+收件次數，完整資料見 people.json）：")
    for r in people[:12]:
        nm = " / ".join(r["names"]) or "(無名稱)"
        mark = "［已確認］" if r.get("confirmed") else ""
        print(f"  寄{r['sent']:>3} 收{r['received']:>3}  {nm}  {r['emails'][0] if r['emails'] else ''} {mark}")
    if alias_candidates:
        print(f"\n可能是同一人的多個位址（需確認）：{len(alias_candidates)} 組")
        for a in alias_candidates[:8]:
            print(f"  {a['name']}: {', '.join(a['keys'])}")
    if errors:
        print(f"\n無法讀取 {len(errors)} 個：")
        for e in errors[:20]:
            print("  - " + e)
    print(f"\n輸出：{out/'emails.json'}、{out/'emails.txt'}、{out/'people.json'}")


if __name__ == "__main__":
    main()
