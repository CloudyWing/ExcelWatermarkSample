namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Represents an Excel page size used to generate a full-page watermark image.
/// </summary>
/// <param name="Name">The page size name.</param>
/// <param name="WidthPixels">The portrait width in pixels.</param>
/// <param name="HeightPixels">The portrait height in pixels.</param>
public sealed record WatermarkPageSize(string Name, int WidthPixels, int HeightPixels) {
    /// <summary>
    /// Gets A4 page size.
    /// </summary>
    public static WatermarkPageSize A4 { get; } = new("A4", 827, 1169);

    /// <summary>
    /// Gets Letter page size.
    /// </summary>
    public static WatermarkPageSize Letter { get; } = new("Letter", 850, 1100);
}
