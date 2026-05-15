namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Represents options for generating the watermark image.
/// </summary>
/// <param name="Text">The watermark text.</param>
/// <param name="PageSize">The target page size.</param>
/// <param name="Orientation">The target page orientation.</param>
/// <param name="FontSize">The watermark font size.</param>
/// <param name="Opacity">The watermark opacity from 0 to 1.</param>
public sealed record WatermarkOptions(
    string Text,
    WatermarkPageSize PageSize,
    WatermarkPageOrientation Orientation,
    float FontSize,
    float Opacity = 0.22f
) {
    /// <summary>
    /// Initializes a new instance of the <see cref="WatermarkOptions"/> record.
    /// </summary>
    /// <param name="text">The watermark text.</param>
    /// <param name="pageWidthPixels">The custom page width in pixels.</param>
    /// <param name="pageHeightPixels">The custom page height in pixels.</param>
    /// <param name="fontSize">The watermark font size.</param>
    /// <param name="opacity">The watermark opacity from 0 to 1.</param>
    public WatermarkOptions(
        string text,
        int pageWidthPixels,
        int pageHeightPixels,
        float fontSize,
        float opacity = 0.22f
    ) : this(text, new WatermarkPageSize("Custom", pageWidthPixels, pageHeightPixels), WatermarkPageOrientation.Portrait, fontSize, opacity) {
    }

    /// <summary>
    /// Gets the effective page width in pixels.
    /// </summary>
    public int PageWidthPixels => Orientation == WatermarkPageOrientation.Landscape
        ? PageSize.HeightPixels
        : PageSize.WidthPixels;

    /// <summary>
    /// Gets the effective page height in pixels.
    /// </summary>
    public int PageHeightPixels => Orientation == WatermarkPageOrientation.Landscape
        ? PageSize.WidthPixels
        : PageSize.HeightPixels;
}
