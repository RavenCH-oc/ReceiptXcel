# DocXcel Manual Acceptance Checklist

## A. 建立模板

- 選擇 Excel
- 選擇 Word 模板
- 讀取欄位對應列
- 建立並保存 `.docxcel.json`

## B. 載入模板

- 關閉程式後重新開啟
- 載入模板設定檔
- 選擇另一份相同 schema 的 Excel
- 確認 schema validation success
- 確認缺少的最近模板不會阻止程式啟動

## C. Word

- 正文 placeholder
- 表格 placeholder
- 頁首與頁尾
- 中文、前導零、日期

## D. Selection

- Marker
- Excel Rows
- Latest N
- Single-mapping runtime marker safety

## E. Batch

- 產生多份 Word
- filename pattern
- duplicate filename suffix
- partial failure
- 開啟輸出資料夾

## F. Cancel

- 大批次開始
- 中途取消
- 已完成文件保留
- 未處理列不產生文件
- 沒有殘留 `.tmp` 或損壞 `.docx`

## G. Safety

- Excel 欄位順序交換時阻止產生
- Word 模板不存在時阻止產生
- invalid row expression 顯示友善訊息
- 輸出資料夾權限不足時顯示友善訊息

## H. Restart

- settings 正常載入
- malformed settings 使用 defaults
- settings 中最近模板不存在時不 crash
