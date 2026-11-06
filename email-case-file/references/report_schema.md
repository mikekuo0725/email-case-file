# report.json 格式

`build_report.py` 讀取這個 JSON 產出報告。所有日期用 `YYYY-MM-DD`，時間軸可用 `YYYY-MM-DD HH:MM`。

## 共用：source（出處）

每條決策、未決議題、待辦、已完成都要附上出處。值可以直接從 `emails.json` 複製。

```json
{"id": "E014", "subject": "RE: A機種 打樣確認", "sender": "John Smith <john@example.com>", "date": "2026-09-05 14:32", "file": "A機種/RE_ A機種 打樣確認.msg"}
```

## 完整結構

```json
{
  "title": "A 機種開發案",
  "folder": "G:/信件/客戶X",
  "generated_date": "2026-09-23",
  "email_count": 52,
  "date_range": ["2026-06-02", "2026-09-20"],
  "overview_summary": "3–5 句的整體現況摘要，點出最需要注意的事。",
  "changes_summary": "增量模式才填：1–3 句，這次新信推進了什麼。第一次整理省略。",

  "figures": [
    {
      "name": "模具費用分攤",
      "unit": "US$",
      "note": "選填，一句話說明這個數字是什麼",
      "versions": [
        {"date": "2026-03-10", "by": "Acme（Tom）", "value": "48,000", "status": "提出", "note": "PI-0001 全額求償", "source": {"id": "E001"}},
        {"date": "2026-07-20", "by": "陳小華 Amy（我方）", "value": "24,000", "status": "提出", "note": "依管理層決議付一半", "source": {"id": "E017"}},
        {"date": "2026-09-24", "by": "Acme", "value": "36,500", "status": "拒絕", "note": "拒絕我方方案並還價", "source": {"id": "E035"}}
      ]
    }
  ],

  "people": [
    {
      "name": "王大明",
      "aliases": ["David Wang"],
      "emails": ["wang@example.com", "david.wang@example.net"],
      "org": "我方公司",
      "role": "副總",
      "can_decide": "是"
    }
  ],

  "topics": [
    {
      "name": "A 機種開發",
      "short_name": "A 機種",
      "email_count": 31,
      "status_summary": "2–4 句，這個議題目前走到哪、卡在哪。",

      "decisions": [
        {
          "date": "2026-09-05",
          "decider": "John Smith",
          "decider_role": "客戶採購經理",
          "decision": "外觀採用 B 方案（霧黑烤漆）。",
          "quote": "OK, please go ahead with option B.",
          "confidence": "高",
          "note": "",
          "source": { "...": "見上方 source 格式；決策或數字出自附件時，加 \"attachment\": \"spec_v2.docx\"" }
        }
      ],

      "pending": [
        {
          "issue": "包裝外箱尺寸是否可加高 5 cm",
          "raised_by": "陳業務（我方）",
          "raised_date": "2026-08-20",
          "options": ["加高 5 cm，貨櫃裝載數減少 8%", "維持原尺寸，改內部緩衝"],
          "waiting_on": "客戶（John, 採購）",
          "note": "9/10 追問一次，仍未回覆",
          "confidence": "高",
          "source": { "...": "..." }
        }
      ],

      "todos": [
        {
          "item": "寄出修改後的打樣 2 台",
          "owner": "王工程師（我方研發）",
          "due": "2026-09-30",
          "confidence": "中",
          "note": "期限由「下週五」推算",
          "source": { "...": "..." }
        }
      ],

      "completed": [
        {
          "item": "第一版規格書確認",
          "completed_date": "2026-07-12",
          "evidence": "客戶回覆「Spec confirmed, no further changes.」",
          "confidence": "高",
          "source": { "...": "..." }
        }
      ],

      "timeline": [
        {
          "id": "E001",
          "date": "2026-06-02 09:15",
          "sender": "John Smith（客戶採購）",
          "recipients": "陳業務、王工程師",
          "cc": "",
          "subject": "A 機種開發需求",
          "file": "A機種/01.msg",
          "headline": "客戶提出 A 機種開發需求，要 7 月底前報價",
          "tag": "提案",
          "points": ["客戶提出 A 機種開發需求，附初版規格", "要求 7 月底前提供打樣報價"],
          "actions": [
            {"item": "提供打樣報價", "owner": "陳業務（我方）", "due": "2026-07-31"}
          ],
          "attachments": ["A_spec_v1.pdf"],
          "attachment_notes": ["A_spec_v1.pdf：單價 USD 120，交期 45 天（讀取附件時才填）"],
          "links": ["https://example.com/share/spec"]
        }
      ]
    }
  ]
}
```

## 欄位說明

- `timeline`：每封有實質內容的信一條，由舊到新（腳本會依 `date` 排序）。`points` 為條列重點（缺少時退回舊的單句 `summary`）；`actions` 的 `owner`、`due` 信中沒寫就留空；`links` 取自 `emails.json`。有 2 個以上議題時，腳本會另外合併成附錄的完整時間序。

- `short_name`：選填，議題的短名（4–8 字），用在總時間軸的「議題」欄，避免長名稱把表格撐開。沒填就用 `name`。
- `timeline[].headline`：選填但強烈建議。總時間軸（總覽第一張表）一封信一行用的那句話，20–40 字，寫「誰做了什麼」，例如「Acme 書面拒絕我方方案，還價 36,500」。沒填就用 `points` 第一條。
- `timeline[].tag`：選填，總時間軸的「類型」欄。只用這幾個：`提案`、`回覆`、`拒絕`、`決策`、`待辦`、`完成`、`資訊`。
- `figures`：選填，「數字怎麼變」表。案子的核心是金額、數量、單價、交期在來回談時就要填（報價、分攤、求償、議價）。每個要追的數字一組，`versions` 依日期列出每一次被提出或修改的版本；`status` 只用 `提出`、`還價`、`拒絕`、`接受`、`定案`、`撤回`；`value` 寫原幣別與原數字，不要換算；`source.id` 必填。同一個數字前後單位不同（US$ 與 €）就照原文寫，在 `note` 註明。沒有談數字的案子省略。報告會放在總覽，速覽列出每組最新版本，xlsx 多一張「數字演變」。
- `people`：人物表。`can_decide` 填「是」「否」「部分」或「不明」；`emails` 要列出這個人用過的所有位址，下次增量整理時靠它（和名稱）認出同一個人。`people.json` 裡帶 `confirmed` 的人是上次已確認過的，直接沿用。
- `changes_summary`：只在增量模式填。「自上次報告以來的變化」那一節的清單（新增決策、新增待辦、剛逾期、已移出待辦、已不在未決清單）由 `build_report.py` 比對上次的 `*.data.json` 自動產生，不用自己列。
- `attachment_notes`：選填，只有用 `--read-attachments` 讀過附件內容才填。寫附件裡的重點（價格、規格、數量、交期），以及和前一版相比變了什麼（可參考 `emails.txt` 的版本差異），一條一句。附件沒讀、讀不到或沒有重點就省略。
- `source.id` 用 D 編號：依據來自資料夾裡的其他檔案（不是信件，也不是信的附件）時，`id` 填 `D001` 這類編號，`attachment` 填檔名，`subject`、`sender` 可省略；`date` 填檔案內文寫的日期，沒有才用修改時間。`build_report.py --emails` 會補上 `file`，報告裡點了直接開那個檔案。
- `source.attachment`：選填。決策、未決議題、待辦、已完成的依據是附件內容而不是信件本文時填附件檔名，報告出處會顯示「附件《檔名》」。
- 信件編號：`parse_emails.py` 會把轉寄信裡引用的舊信拆成獨立的信（`src: "quote"`，`container` 是收錄它的那封），跨檔去重後依真實寄件時間編號。引用信的時間取自引用標頭，時區是推估的（`date_approx: true`）；同一天內先後順序若關係到「誰先說」，回原信確認。
- `source.file`、`timeline[].file`：信件檔案相對於信件資料夾的路徑，用來產生超連結。可以省略，`build_report.py --emails` 會依信件編號補齊。
- **增量模式的穩定性**：腳本用「出處信件編號＋事項文字」（待辦用 `item`，未決用 `issue`）比對新舊報告。沒有變動的條目請**原樣沿用上次的文字與出處**，否則會被誤判為「已移出」加「新增」。
- `confidence`：只能是 `高`、`中`、`低`。`低` 的項目在報告中會顯示「待確認」。
- `days_pending`：不用填，腳本依 `raised_date` 與 `generated_date` 計算。
- `due`：沒有明確期限就填空字串 `""`，報告會顯示「未定」並排在最後。
- `note`：選填，沒有就填 `""` 或省略。
- 任何清單沒有內容時給空陣列 `[]`，報告會顯示「目前沒有」。
- 不要自己寫總覽的待辦與未決清單，腳本會從各議題彙整。
