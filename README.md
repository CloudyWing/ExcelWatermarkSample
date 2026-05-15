# ExcelWatermarkSample

這是一個 .NET 10 Console 範例，搭配 NPOI 寫入帶浮水印的 Excel。

實作背景與設計考量可參考這篇筆記：[使用 .NET 產生帶有浮水印的 Excel](https://note.cloudywing.net/backend/%E4%BD%BF%E7%94%A8%20.NET%20%E7%94%A2%E7%94%9F%E5%B8%B6%E6%9C%89%E6%B5%AE%E6%B0%B4%E5%8D%B0%E7%9A%84%20Excel)。

## 技術重點

- Target Framework：`.NET 10`
- Excel：`NPOI 2.8.0`
- 圖片產生：`SkiaSharp 3.119.2`
- 測試：`NUnit 4.6.0` + `NSubstitute 5.3.0`
- 專案結構：`src` / `tests`

## 版本策略

此範例採用新版 NPOI 單一路線，不同時維護舊版與新版兩套程式碼。

原因：

- GitHub 範例應以目前可維護、可重現的版本為主。
- 舊版 NPOI 的低階 API 差異會讓範例偏向版本相容性整理，讓主題失焦。
- NPOI 2.8.0 已支援 `net10.0`，可直接驗證背景圖片與 header 圖片的低階 OOXML/VML 寫法。

NPOI 2.8.0 會帶入 `System.Security.Cryptography.Xml` 相依性。此範例直接參考 `System.Security.Cryptography.Xml 10.0.8`，避免使用含已知弱點的 transitive 版本。

## 功能

- 使用 SkiaSharp 產生 PNG 浮水印圖片。
- 將圖片設定為工作表背景，支援一般檢視模式。
- 將圖片設定到 header，支援整頁檢視與列印。
- 支援 `Background`、`Header`、`BackgroundAndHeader` 三種寫入模式。
- 支援多工作表，每張工作表各自建立正確的 relationship。
- 支援 A4、Letter 與自訂紙張尺寸，並可切換直向與橫向。
- 設定 Excel 預設字型，避免欄寬受預設字型差異影響。
- 測試直接檢查 xlsx 內部 zip entry、sheet XML、sheet relationship 與 VML relationship。

## 程式結構

| 類別 | 職責 |
| --- | --- |
| `ReportWorkbookBuilder` | 建立範例報表內容與工作表。 |
| `WatermarkImageBuilder` | 使用 SkiaSharp 產生滿版 PNG 浮水印圖片。 |
| `ExcelWatermarkWriter` | 將浮水印圖片寫入 NPOI XSSF 工作表的背景與 header。 |
| `WorkbookSerializer` | 將 `XSSFWorkbook` 序列化為 xlsx bytes。 |

## 使用範例

```csharp
WatermarkOptions options = new(
    Text: "CONFIDENTIAL",
    PageSize: WatermarkPageSize.A4,
    Orientation: WatermarkPageOrientation.Portrait,
    FontSize: 72
);

WatermarkImage image = new WatermarkImageBuilder().Build(options);
ReportWorkbookBuilder workbookBuilder = new();
ExcelWatermarkWriter watermarkWriter = new();

using XSSFWorkbook workbook = workbookBuilder.CreateWorkbook();
for (int i = 0; i < workbook.NumberOfSheets; i++) {
    watermarkWriter.Apply(
        (XSSFSheet)workbook.GetSheetAt(i),
        image,
        WatermarkPlacement.BackgroundAndHeader
    );
}

byte[] workbookBytes = WorkbookSerializer.ToBytes(workbook);
```

## 限制

- Excel 本身沒有等同 Word 的浮水印功能，此範例改用背景圖片與 header 圖片來模擬。
- Background 圖片對應一般檢視。
- Header 圖片對應整頁檢視與列印。
- 分頁預覽的呈現會因 Excel 版本與檢視模式而有差異。

## 執行

```powershell
dotnet run --project .\src\ExcelWatermarkSample\ExcelWatermarkSample.csproj
```

輸出檔案會建立在執行目錄的 `artifacts` 資料夾。

## 測試

```powershell
dotnet test .\ExcelWatermarkSample.slnx
```

## 授權條款

本專案採用 [MIT License](LICENSE.md)。
