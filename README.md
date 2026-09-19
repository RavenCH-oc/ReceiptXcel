# ReceiptXcel

ReceiptXcel｜自行收納款項收據產生工具，是只讀 Excel、產生固定「自行收納款項統一收據」三聯 DOCX 的 Windows WPF 工具。

一般使用者不需要編輯 Word 模板、設定欄位 mapping、輸入 placeholder 或設定檔名樣式。程式會使用內建固定收據格式，一個 Excel 資料列產生一份包含第一聯、第二聯、第三聯的 Word 文件。

## Excel 格式

只接受工作表 `工作表1`，第 2 列必須完全符合以下欄位順序：

```text
年 / 月 / 日 / 編號 / 繳款人 / 數字金額 / 事由 / 承辦人
```

第 1 列可作為標題，第 3 列起為資料。年、月、日、編號、繳款人、數字金額與事由必須可解析。H「承辦人」欄目前僅為相容既有 Excel 登記表格式而保留，不會輸出到收據；Word 的「經手人」填寫格目前固定留白。

## 使用方式

1. 啟動 ReceiptXcel。
2. 按「選擇」載入 `.xlsx` 收據登記表。
3. 確認工作表與「Excel 格式正確」訊息。
4. 選擇「最新資料」或「Excel 列號」。
5. 確認預計產生筆數，選擇輸出資料夾。
6. 按「產生收據」。

「最新資料」會從最下面的非空白資料列取出指定筆數，再按照 Excel 列號由小到大產生。「Excel 列號」支援單列、逗號清單與範圍，例如 `3`、`3,5,8`、`3-6`、`3,5,8-10`。

## 金額與輸出安全

- 只允許正整數元，最大值為 `9,999,999`。
- 程式自動處理固定表單的七格金額欄位與中文財務大寫。
- 每列固定輸出為 `收據_{收據編號}.docx`。
- 已存在的檔案不會覆寫，也不會自動加上 `(2)` 等尾碼。
- 同一批次出現重複收據編號時，整批會在產生前停止。
- 個別資料列格式錯誤或輸出檔已存在時，該列失敗，其他列仍會繼續。
- Excel 原檔永遠不會被寫回。
- 取消批次時，已完成文件保留，未處理列不產生，暫存檔會清理。

## 系統需求

產生 DOCX 不需要安裝 Microsoft Excel 或 Word；ReceiptXcel 使用 ClosedXML 與 Open XML SDK 讀取／產生檔案。查看與列印產出的 DOCX 需要 Microsoft Word 或相容軟體。

本工具不使用 Office Interop、PDF、LibreOffice、資料庫或網路服務。

## 設定與 diagnostics

ReceiptXcel 使用自己的產品資料夾，不會與 DocXcel 共用：

```text
%LocalAppData%\ReceiptXcel\settings.json
%LocalAppData%\ReceiptXcel\Logs\
```

只保存低風險偏好：最後輸出資料夾、最後選擇模式與視窗大小。不會保存 Excel 內容、選取列運算式、收據資料、金額、繳款人、事由或批次結果。diagnostics 可以記錄時間、版本、操作、內部錯誤與例外堆疊，但不會記錄業務欄位或 Word 正文。

## 建置與測試

```powershell
dotnet restore
dotnet build
dotnet test
dotnet build -c Release
dotnet test -c Release
```

Portable Windows staging publish 使用：

```text
src/XlsxDocxGenerator/Properties/PublishProfiles/WinX64Portable.pubxml
```

輸出設定為 `win-x64`、self-contained、未啟用 trimming、未 single-file，staging 位置為 `artifacts/publish/win-x64-phase2/`；staging executable 為 `ReceiptXcel 收據產生工具.exe`。本專案不提供 installer。

## 內建固定格式

內建格式位於 `src/XlsxDocxGenerator/Assets/Templates/receipt-template.docx`。Debug、Release 與 publish 都會由專案內容複製到執行檔旁的 `Assets/Templates/`。若檔案遺失或 contract 損壞，產生按鈕會停用並顯示重新安裝訊息；一般使用者不能選擇其他 Word 模板補救。
