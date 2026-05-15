using System.IO.Compression;

namespace CloudyWing.ExcelWatermarkSample.Tests;

[TestFixture]
internal sealed class ReportWorkbookBuilderTests {
    [Test]
    public void CreateWorkbook_WhenCalled_CreatesMultipleSheets() {
        ReportWorkbookBuilder builder = new();

        using NPOI.XSSF.UserModel.XSSFWorkbook workbook = builder.CreateWorkbook();

        Assert.Multiple(() => {
            Assert.That(workbook.NumberOfSheets, Is.EqualTo(2));
            Assert.That(workbook.GetSheetAt(0).SheetName, Is.EqualTo("Sales"));
            Assert.That(workbook.GetSheetAt(1).SheetName, Is.EqualTo("Inventory"));
        });
    }

    [Test]
    public void CreateWorkbook_WhenSerialized_ConfiguresDefaultWorkbookFont() {
        ReportWorkbookBuilder builder = new();

        using NPOI.XSSF.UserModel.XSSFWorkbook workbook = builder.CreateWorkbook();
        using ZipArchive archive = new(new MemoryStream(WorkbookSerializer.ToBytes(workbook)), ZipArchiveMode.Read);
        string stylesXml = ReadEntry(archive, "xl/styles.xml");

        Assert.Multiple(() => {
            Assert.That(stylesXml, Does.Contain("微軟正黑體"));
            Assert.That(stylesXml, Does.Contain("val=\"12\""));
        });
    }

    private static string ReadEntry(ZipArchive archive, string entryName) {
        ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException($"Entry not found: {entryName}");
        using Stream stream = entry.Open();
        using StreamReader reader = new(stream);
        return reader.ReadToEnd();
    }
}
