# email-case-file

A Claude skill that turns a folder of saved Outlook emails (.msg, .eml, .pst) into a "case file" report.

- Groups emails by topic and lists decisions (who decided, when, with the original quote as evidence), open questions (with days pending), to-dos, and completed items
- A chronological report, one entry per email, oldest first
- Thread splitting: old emails quoted inside forwards and replies become separate emails, duplicates across files are removed, and everything is sorted by real send time
- An overall timeline (one line per email) and a "how the numbers changed" table for prices and quantities under negotiation
- Thread linking: emails are linked by their reply headers (In-Reply-To / References, Outlook ConversationIndex), so a thread stays together even when the subject changes
- Large folders (120+ emails) are split across several subagents that read in parallel, then merged
- Incremental updates: only new emails are processed, with a "changes since the last report" section
- People list with alias merging; auto-replies and system notifications are filtered out
- Optional: read attachment contents and compare versions (`--read-attachments`)
- Output: Word (plus PDF and HTML), a one-page meeting brief, a to-do Excel file, and a calendar .ics file; every source link opens the original email

The skill's instructions and report output are in Traditional Chinese.

## Install
Zip the `email-case-file/` folder into a `.skill` file, or import it directly into Claude's Skills.

## Dependencies
`pip install extract-msg python-docx openpyxl` (reading attachments also needs `pypdf python-pptx`)

---

# email-case-file（中文）

Claude skill：把資料夾裡的 Outlook 信件（.msg、.eml、.pst）整理成「案卷報告」。

- 分議題，整理已決策事項（決定者、日期、原文佐證）、尚未決策（附懸而未決天數）、待辦、已完成
- 逐封由舊到新的時間序報告
- 拆信串：轉寄／回覆信夾帶的整串舊信拆成獨立信件，跨檔去重，依真實寄件時間排序
- 總時間軸（一封一行）與「數字怎麼變」表（金額、數量的議價過程）
- 串信串：用信件標頭的回覆關係（In-Reply-To／References、Outlook ConversationIndex）把同一串信接起來，主旨改了也不會斷
- 信很多（超過 120 封）時，分給多個小幫手（subagent）分段判讀，再由主流程合併
- 增量更新：只處理新信，列出「自上次報告以來的變化」
- 人物表與別名合併、系統信自動排除
- 可選：讀取附件內容並比對版本差異（`--read-attachments`）
- 輸出：Word（可加 PDF、HTML）、一頁會議速覽、待辦 Excel、行事曆 .ics；出處可點選開啟原信

## 安裝
把 `email-case-file/` 資料夾打包成 `.skill`（或直接匯入 Claude 的 Skills）。

## 相依套件
`pip install extract-msg python-docx openpyxl`（讀附件另需 `pypdf python-pptx`）
