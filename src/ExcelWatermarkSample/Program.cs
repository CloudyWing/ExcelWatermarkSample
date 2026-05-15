using CloudyWing.ExcelWatermarkSample;
using NPOI.XSSF.UserModel;

WatermarkOptions options = new(
    Text: "CONFIDENTIAL",
    PageSize: WatermarkPageSize.A4,
    Orientation: WatermarkPageOrientation.Portrait,
    FontSize: 72
);
WatermarkImage image = new WatermarkImageBuilder().Build(options);
ReportWorkbookBuilder workbookBuilder = new();
ExcelWatermarkWriter watermarkWriter = new();

string outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts");
Directory.CreateDirectory(outputDirectory);

string outputPath = Path.Combine(outputDirectory, "report-with-watermark.xlsx");
using XSSFWorkbook workbook = workbookBuilder.CreateWorkbook();
for (int i = 0; i < workbook.NumberOfSheets; i++) {
    watermarkWriter.Apply(
        (XSSFSheet)workbook.GetSheetAt(i),
        image,
        WatermarkPlacement.BackgroundAndHeader
    );
}

await File.WriteAllBytesAsync(outputPath, WorkbookSerializer.ToBytes(workbook));

Console.WriteLine("ExcelWatermarkSample");
Console.WriteLine($"Workbook: {outputPath}");
