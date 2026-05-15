namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Defines where the watermark image should be written in the worksheet.
/// </summary>
[Flags]
public enum WatermarkPlacement {
    /// <summary>
    /// Does not write the watermark.
    /// </summary>
    None = 0,

    /// <summary>
    /// Writes the watermark as a worksheet background image.
    /// </summary>
    Background = 1,

    /// <summary>
    /// Writes the watermark as a centered header image.
    /// </summary>
    Header = 2,

    /// <summary>
    /// Writes the watermark both as background and header image.
    /// </summary>
    BackgroundAndHeader = Background | Header
}
