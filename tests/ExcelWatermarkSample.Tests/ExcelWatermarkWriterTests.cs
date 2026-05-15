using System.IO.Compression;
using System.Xml.Linq;
using NPOI.XSSF.UserModel;

namespace CloudyWing.ExcelWatermarkSample.Tests;

[TestFixture]
internal sealed class ExcelWatermarkWriterTests {
    private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace VmlNamespace = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace OfficeNamespace = "urn:schemas-microsoft-com:office:office";

    [Test]
    public void Apply_WhenAppliedToAllSheets_WritesWatermarkPartsForEverySheet() {
        byte[] workbookBytes = CreateWorkbook(WatermarkPlacement.BackgroundAndHeader);
        using ZipArchive archive = OpenWorkbook(workbookBytes);

        Assert.Multiple(() => {
            AssertSheetHasWatermark(archive, 1);
            AssertSheetHasWatermark(archive, 2);
            Assert.That(archive.GetEntry("xl/media/image1.png"), Is.Not.Null);
            Assert.That(archive.GetEntry("xl/media/image2.png"), Is.Not.Null);
        });
    }

    [Test]
    public void Apply_WhenBackgroundOnly_DoesNotWriteHeaderDrawing() {
        byte[] workbookBytes = CreateWorkbook(WatermarkPlacement.Background);
        using ZipArchive archive = OpenWorkbook(workbookBytes);
        XDocument sheet = ReadXml(archive, "xl/worksheets/sheet1.xml");

        Assert.Multiple(() => {
            Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "picture"), Is.Not.Null);
            Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "legacyDrawingHF"), Is.Null);
            Assert.That(archive.Entries.Any(x => x.FullName.StartsWith("xl/drawings/vmlDrawing", StringComparison.Ordinal)), Is.False);
        });
    }

    [Test]
    public void Apply_WhenHeaderOnly_DoesNotWriteBackgroundPicture() {
        byte[] workbookBytes = CreateWorkbook(WatermarkPlacement.Header);
        using ZipArchive archive = OpenWorkbook(workbookBytes);
        XDocument sheet = ReadXml(archive, "xl/worksheets/sheet1.xml");

        Assert.Multiple(() => {
            Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "picture"), Is.Null);
            Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "legacyDrawingHF"), Is.Not.Null);
            Assert.That(ReadEntry(archive, "xl/worksheets/sheet1.xml"), Does.Contain("<![CDATA[&C&G]]>"));
        });
    }

    [Test]
    public void Apply_WhenBackgroundAndHeader_RelationshipsPointToExpectedParts() {
        byte[] workbookBytes = CreateWorkbook(WatermarkPlacement.BackgroundAndHeader);
        using ZipArchive archive = OpenWorkbook(workbookBytes);

        AssertSheetRelationshipConsistency(archive, 1);
    }

    private static byte[] CreateWorkbook(WatermarkPlacement placement) {
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
            watermarkWriter.Apply((XSSFSheet)workbook.GetSheetAt(i), image, placement);
        }

        return WorkbookSerializer.ToBytes(workbook);
    }

    private static void AssertSheetHasWatermark(ZipArchive archive, int sheetNumber) {
        XDocument sheet = ReadXml(archive, $"xl/worksheets/sheet{sheetNumber}.xml");
        string sheetXml = ReadEntry(archive, $"xl/worksheets/sheet{sheetNumber}.xml");

        Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "picture"), Is.Not.Null);
        Assert.That(sheet.Root?.Element(SpreadsheetNamespace + "legacyDrawingHF"), Is.Not.Null);
        Assert.That(sheetXml, Does.Contain("<![CDATA[&C&G]]>"));
    }

    private static void AssertSheetRelationshipConsistency(ZipArchive archive, int sheetNumber) {
        string sheetPath = $"xl/worksheets/sheet{sheetNumber}.xml";
        XDocument sheet = ReadXml(archive, sheetPath);
        IReadOnlyDictionary<string, RelationshipInfo> sheetRelationships = ReadRelationships(
            archive, $"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels"
        );

        string pictureRelationshipId = GetRelationshipId(sheet, "picture");
        string headerRelationshipId = GetRelationshipId(sheet, "legacyDrawingHF");

        RelationshipInfo backgroundImageRelationship = sheetRelationships[pictureRelationshipId];
        RelationshipInfo vmlRelationship = sheetRelationships[headerRelationshipId];
        string vmlPartPath = $"xl/drawings/{Path.GetFileName(vmlRelationship.Target)}";
        string vmlRelationshipsPath = $"xl/drawings/_rels/{Path.GetFileName(vmlRelationship.Target)}.rels";
        XDocument vmlDrawing = ReadXml(archive, vmlPartPath);
        IReadOnlyDictionary<string, RelationshipInfo> vmlRelationships = ReadRelationships(archive, vmlRelationshipsPath);
        string vmlImageRelationshipId = vmlDrawing
            .Descendants(VmlNamespace + "imagedata")
            .Single()
            .Attribute(OfficeNamespace + "relid")!
            .Value;

        Assert.Multiple(() => {
            Assert.That(backgroundImageRelationship.Type, Does.EndWith("/image"));
            Assert.That(backgroundImageRelationship.Target, Does.StartWith("../media/image"));
            Assert.That(vmlRelationship.Type, Does.EndWith("/vmlDrawing"));
            Assert.That(vmlRelationships[vmlImageRelationshipId].Type, Does.EndWith("/image"));
            Assert.That(vmlRelationships[vmlImageRelationshipId].Target, Does.StartWith("../media/image"));
            Assert.That(archive.GetEntry(vmlPartPath), Is.Not.Null);
        });
    }

    private static string GetRelationshipId(XDocument document, string elementName) {
        return document.Root
            ?.Element(SpreadsheetNamespace + elementName)
            ?.Attribute(RelationshipsNamespace + "id")
            ?.Value
            ?? throw new InvalidOperationException($"Relationship id not found: {elementName}");
    }

    private static IReadOnlyDictionary<string, RelationshipInfo> ReadRelationships(ZipArchive archive, string entryName) {
        XDocument document = ReadXml(archive, entryName);
        return document
            .Root!
            .Elements(PackageRelationshipsNamespace + "Relationship")
            .ToDictionary(
                x => x.Attribute("Id")!.Value,
                x => new RelationshipInfo(
                    x.Attribute("Type")!.Value,
                    x.Attribute("Target")!.Value
                )
            );
    }

    private static ZipArchive OpenWorkbook(byte[] workbookBytes) {
        return new ZipArchive(new MemoryStream(workbookBytes), ZipArchiveMode.Read);
    }

    private static XDocument ReadXml(ZipArchive archive, string entryName) {
        using StringReader reader = new(ReadEntry(archive, entryName));
        return XDocument.Load(reader);
    }

    private static string ReadEntry(ZipArchive archive, string entryName) {
        ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException($"Entry not found: {entryName}");
        using Stream stream = entry.Open();
        using StreamReader reader = new(stream);
        return reader.ReadToEnd();
    }

    private sealed record RelationshipInfo(string Type, string Target);
}
