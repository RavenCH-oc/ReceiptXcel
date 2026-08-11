# ReceiptXcel

ReceiptXcel 是「自行收納款項統一收據」的固定表單產生工具。它沿用 DocXcel v0.1.0 已驗證的 Windows WPF、ClosedXML 與 Open XML 基礎，但本 repository 有自己的 Git history，產品版本為 `0.1.0-dev`。

本 Phase 0 建立固定 Excel schema、收據 record、日期/編號規則、金額解析與三聯 Word 內建模板；正式 specialized UI 與完整生成 workflow 留待後續 Phase。

## 第一次使用

1. 準備符合 `工作表1`、Row 2 固定欄位的 Excel。
2. 使用內建 `src/XlsxDocxGenerator/Assets/Templates/receipt-template.docx`。
3. 由 specialized receipt workflow 讀取資料列並產生三聯 DOCX（正式 UI 尚未在本 Phase 實作）。

## 目前保留的 inherited baseline

現有通用模板 UI、`.docxcel.json` 與 batch infrastructure 暫時保留，以維持 DocXcel baseline 的 build/test regression；它們不代表 ReceiptXcel 的最終產品流程。

模板設定檔保存模板名稱、Word 路徑、preferred worksheet、欄位 mapping 與輸出檔名樣式，不保存 Excel 資料內容、selection rule 或 batch result。

## Runtime marker safety

兩個以上欄位 mapping 會使用 strict automatic detection：必須完整匹配所有 mapping placeholder，且其他非 mapping 欄位不可有非空內容。單一 mapping 無法可靠分辨 `{{NAME}}` 是 marker 或真實資料，因此不會自動排除；若目前 Excel 仍保留該列，請在畫面指定「目前 Excel 欄位對應列」。留白代表不排除任何 runtime marker row。此指定值只存在於目前 workbook/session，不會寫入模板 JSON。

建立模板時的 `CreationMappingMarkerRowNumber` 只是來源 Excel metadata，不會永久排除新 Excel 的相同列號。

## 支援的 Word 部位

支援正文、表格、頁首、頁尾，以及跨 Run/Text node 的 placeholder。頁首與頁尾包含 first/even page parts。

目前不支援文字方塊、Shapes、圖片 placeholder、註腳、章節附註、Comments、Content Controls，以及特殊 tracked-changes 結構。

## 批次與取消

產生批次期間可以按「取消」。目前文件會安全完成或清理 temporary file；已完成的 DOCX 保留，尚未處理的資料列不會產生輸出。

## 本機設定與 diagnostics

低風險使用偏好保存於：

```text
%LocalAppData%\ReceiptXcel\settings.json
```

diagnostics log 保存於：

```text
%LocalAppData%\ReceiptXcel\Logs\
```

設定與 log 都不保存 Excel cell、Word 正文或 batch 業務資料。

## Portable Windows build

Release portable profile：

```text
src/XlsxDocxGenerator/Properties/PublishProfiles/WinX64Portable.pubxml
```

目標為 `win-x64`、self-contained、未啟用 aggressive trimming，輸出至 `artifacts/publish/win-x64/`。本專案不提供 MSI、MSIX 或其他 installer。

人工驗收項目請參考 [docs/ACCEPTANCE_CHECKLIST.md](docs/ACCEPTANCE_CHECKLIST.md)。
