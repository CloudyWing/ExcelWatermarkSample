using NPOI.XSSF.UserModel;

namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Serializes NPOI workbooks to xlsx bytes.
/// </summary>
public static class WorkbookSerializer {
    /// <summary>
    /// Serializes a workbook to bytes.
    /// </summary>
    /// <param name="workbook">The workbook to serialize.</param>
    /// <returns>The serialized xlsx bytes.</returns>
    public static byte[] ToBytes(XSSFWorkbook workbook) {
        ArgumentNullException.ThrowIfNull(workbook);

        using MemoryStream stream = new();
        workbook.Write(stream, leaveOpen: true);
        return stream.ToArray();
    }
}
