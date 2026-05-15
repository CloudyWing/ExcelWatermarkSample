using System.Globalization;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Writes watermark images into an XSSF worksheet.
/// </summary>
public sealed class ExcelWatermarkWriter {
    /// <summary>
    /// Applies a watermark image to a worksheet.
    /// </summary>
    /// <param name="sheet">The target worksheet.</param>
    /// <param name="image">The watermark image.</param>
    /// <param name="placement">The target placement.</param>
    public void Apply(XSSFSheet sheet, WatermarkImage image, WatermarkPlacement placement) {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(image);

        if (placement == WatermarkPlacement.None) {
            return;
        }

        if (sheet.Workbook is not XSSFWorkbook workbook) {
            throw new InvalidOperationException("Watermark writer only supports XSSFWorkbook.");
        }

        POIXMLDocumentPart imagePart = AddImagePart(workbook, image);

        if ((placement & WatermarkPlacement.Background) == WatermarkPlacement.Background) {
            ApplyBackground(sheet, imagePart);
        }

        if ((placement & WatermarkPlacement.Header) == WatermarkPlacement.Header) {
            ApplyHeader(workbook, sheet, image, imagePart);
        }
    }

    private static POIXMLDocumentPart AddImagePart(XSSFWorkbook workbook, WatermarkImage image) {
        int pictureIndex = workbook.AddPicture(image.PngBytes, PictureType.PNG);
        if (workbook.GetAllPictures()[pictureIndex] is not POIXMLDocumentPart imagePart) {
            throw new InvalidOperationException("NPOI did not return an OOXML image part.");
        }

        return imagePart;
    }

    private static void ApplyBackground(XSSFSheet sheet, POIXMLDocumentPart imagePart) {
        POIXMLDocumentPart.RelationPart backgroundRelation = sheet.AddRelation(null, XSSFRelation.IMAGES, imagePart);
        sheet.GetCTWorksheet().picture = new CT_SheetBackgroundPicture {
            id = backgroundRelation.Relationship.Id
        };
    }

    private static void ApplyHeader(
        XSSFWorkbook workbook,
        XSSFSheet sheet,
        WatermarkImage image,
        POIXMLDocumentPart imagePart
    ) {
        int drawingNumber = workbook.GetPackagePart()
            .Package
            .GetPartsByContentType(XSSFRelation.VML_DRAWINGS.ContentType)
            .Count + 1;
        VmlDrawing drawing = (VmlDrawing)sheet.CreateRelationship(
            VmlRelation.Instance,
            XSSFFactory.GetInstance(),
            drawingNumber
        );
        POIXMLDocumentPart.RelationPart headerRelation = drawing.AddRelation(null, XSSFRelation.IMAGES, imagePart);

        drawing.PictureRelationshipId = headerRelation.Relationship.Id;
        drawing.WidthPoints = image.WidthPoints;
        drawing.HeightPoints = image.HeightPoints;

        sheet.Header.Center = NPOI.HSSF.UserModel.HeaderFooter.PICTURE_FIELD.sequence;
        sheet.GetCTWorksheet().legacyDrawingHF = new CT_LegacyDrawing {
            id = sheet.GetRelationId(drawing)
        };
    }

    private sealed class VmlRelation : POIXMLRelation {
        private static readonly Lazy<VmlRelation> LazyInstance = new(() => new VmlRelation(
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            "/xl/drawings/vmlDrawing#.vml",
            typeof(VmlDrawing)
        ));

        private VmlRelation(string type, string rel, string defaultName, Type cls) : base(type, rel, defaultName, cls) {
        }

        public static VmlRelation Instance => LazyInstance.Value;
    }

    private sealed class VmlDrawing : POIXMLDocumentPart {
        public string PictureRelationshipId { get; set; } = "";

        public double WidthPoints { get; set; }

        public double HeightPoints { get; set; }

        protected override void Commit() {
            PackagePart part = GetPackagePart();
            using Stream output = part.GetOutputStream();
            Write(output);
        }

        private void Write(Stream stream) {
            string width = WidthPoints.ToString("0.##", CultureInfo.InvariantCulture);
            string height = HeightPoints.ToString("0.##", CultureInfo.InvariantCulture);
            using StreamWriter writer = new(stream);
            writer.Write($"""
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                  <o:shapelayout v:ext="edit">
                    <o:idmap v:ext="edit" data="1" />
                  </o:shapelayout>
                  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
                    <v:stroke joinstyle="miter" />
                    <v:path gradientshapeok="t" o:connecttype="rect" />
                  </v:shapetype>
                  <v:shape id="CH" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:{width}pt;height:{height}pt;z-index:1">
                    <v:imagedata o:relid="{PictureRelationshipId}" o:title="" />
                    <o:lock v:ext="edit" rotation="t" />
                  </v:shape>
                </xml>
                """);
        }
    }
}
