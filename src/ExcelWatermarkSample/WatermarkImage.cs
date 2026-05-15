namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Represents a generated watermark image.
/// </summary>
/// <param name="PngBytes">The PNG bytes.</param>
/// <param name="WidthPixels">The image width in pixels.</param>
/// <param name="HeightPixels">The image height in pixels.</param>
public sealed record WatermarkImage(byte[] PngBytes, int WidthPixels, int HeightPixels) {
    /// <summary>
    /// Gets the image width in points assuming 96 DPI.
    /// </summary>
    public double WidthPoints => WidthPixels * 72d / 96d;

    /// <summary>
    /// Gets the image height in points assuming 96 DPI.
    /// </summary>
    public double HeightPoints => HeightPixels * 72d / 96d;
}
