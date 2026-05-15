using SkiaSharp;

namespace CloudyWing.ExcelWatermarkSample;

/// <summary>
/// Generates PNG watermark images without relying on System.Drawing.Common.
/// </summary>
public sealed class WatermarkImageBuilder {
    /// <summary>
    /// Builds a watermark PNG image.
    /// </summary>
    /// <param name="options">The watermark options.</param>
    /// <returns>The generated watermark image.</returns>
    public WatermarkImage Build(WatermarkOptions options) {
        ArgumentNullException.ThrowIfNull(options);
        Validate(options);

        using SKBitmap bitmap = new(options.PageWidthPixels, options.PageHeightPixels);
        using SKCanvas canvas = new(bitmap);
        canvas.Clear(SKColors.White);

        byte alpha = (byte)Math.Clamp(options.Opacity * 255, 0, 255);
        using SKPaint paint = new() {
            Color = new SKColor(120, 120, 120, alpha),
            IsAntialias = true
        };
        using SKFont font = new(SKTypeface.Default, options.FontSize);
        SKRect textBounds = new();
        font.MeasureText(options.Text, out textBounds);

        canvas.Translate(options.PageWidthPixels / 2f, options.PageHeightPixels / 2f);
        canvas.RotateDegrees(-35);
        canvas.DrawText(options.Text, -textBounds.MidX, -textBounds.MidY, font, paint);

        using SKImage image = SKImage.FromBitmap(bitmap);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);

        return new WatermarkImage(data.ToArray(), options.PageWidthPixels, options.PageHeightPixels);
    }

    private static void Validate(WatermarkOptions options) {
        ArgumentException.ThrowIfNullOrWhiteSpace(options.Text);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.PageWidthPixels);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.PageHeightPixels);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.FontSize);

        if (options.Opacity is < 0 or > 1) {
            throw new ArgumentOutOfRangeException(nameof(options), options.Opacity, "Opacity must be between 0 and 1.");
        }
    }
}
