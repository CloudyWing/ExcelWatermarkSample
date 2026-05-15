using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Creates the sample report workbook content.
/// </summary>
public sealed class ReportWorkbookBuilder {
    /// <summary>
    /// Creates a workbook with multiple sheets.
    /// </summary>
    /// <returns>The generated workbook.</returns>
    public XSSFWorkbook CreateWorkbook() {
        XSSFWorkbook workbook = new();
        ConfigureDefaultFont(workbook);
        CreateSalesSheet(workbook);
        CreateInventorySheet(workbook);

        return workbook;
    }

    private static void ConfigureDefaultFont(XSSFWorkbook workbook) {
        IFont defaultFont = workbook.GetFontAt(0);
        defaultFont.FontName = "微軟正黑體";
        defaultFont.FontHeightInPoints = 12;
    }

    private static void CreateSalesSheet(IWorkbook workbook) {
        ISheet sheet = workbook.CreateSheet("Sales");
        IRow titleRow = sheet.CreateRow(0);
        titleRow.CreateCell(0).SetCellValue("Monthly Sales Report");

        IRow headerRow = sheet.CreateRow(2);
        headerRow.CreateCell(0).SetCellValue("Region");
        headerRow.CreateCell(1).SetCellValue("Amount");
        headerRow.CreateCell(2).SetCellValue("Status");

        string[] regions = ["North", "South", "East", "West"];
        for (int i = 0; i < regions.Length; i++) {
            IRow row = sheet.CreateRow(i + 3);
            row.CreateCell(0).SetCellValue(regions[i]);
            row.CreateCell(1).SetCellValue(12000 + (i * 1300));
            row.CreateCell(2).SetCellValue(i % 2 == 0 ? "Ready" : "Review");
        }

        sheet.SetColumnWidth(0, 16 * 256);
        sheet.SetColumnWidth(1, 14 * 256);
        sheet.SetColumnWidth(2, 18 * 256);
    }

    private static void CreateInventorySheet(IWorkbook workbook) {
        ISheet sheet = workbook.CreateSheet("Inventory");
        IRow titleRow = sheet.CreateRow(0);
        titleRow.CreateCell(0).SetCellValue("Inventory Snapshot");

        IRow headerRow = sheet.CreateRow(2);
        headerRow.CreateCell(0).SetCellValue("Product");
        headerRow.CreateCell(1).SetCellValue("Quantity");
        headerRow.CreateCell(2).SetCellValue("Warehouse");

        string[] products = ["Laptop", "Monitor", "Keyboard", "Mouse"];
        for (int i = 0; i < products.Length; i++) {
            IRow row = sheet.CreateRow(i + 3);
            row.CreateCell(0).SetCellValue(products[i]);
            row.CreateCell(1).SetCellValue(25 - (i * 3));
            row.CreateCell(2).SetCellValue(i % 2 == 0 ? "Taipei" : "Kaohsiung");
        }

        sheet.SetColumnWidth(0, 18 * 256);
        sheet.SetColumnWidth(1, 14 * 256);
        sheet.SetColumnWidth(2, 18 * 256);
    }
}
