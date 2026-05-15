namespace CloudyWing.ExcelWatermarkSample.Tests;

[TestFixture]
internal sealed class WatermarkImageBuilderTests {
    [Test]
    public void Build_WithValidOptions_ReturnsPngImage() {
        WatermarkImageBuilder builder = new();
        WatermarkOptions options = new(
            "CONFIDENTIAL",
            WatermarkPageSize.A4,
            WatermarkPageOrientation.Portrait,
            72
        );

        WatermarkImage image = builder.Build(options);

        Assert.Multiple(() => {
            Assert.That(image.WidthPixels, Is.EqualTo(827));
            Assert.That(image.HeightPixels, Is.EqualTo(1169));
            Assert.That(image.PngBytes.Take(8), Is.EqualTo(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }));
        });
    }

    [Test]
    public void Build_WithInvalidOpacity_ThrowsArgumentOutOfRangeException() {
        WatermarkImageBuilder builder = new();
        WatermarkOptions options = new(
            "CONFIDENTIAL",
            WatermarkPageSize.A4,
            WatermarkPageOrientation.Portrait,
            72,
            1.1f
        );

        Assert.Throws<ArgumentOutOfRangeException>(() => builder.Build(options));
    }

    [Test]
    public void Build_WithLandscapeOrientation_SwapsPageDimensions() {
        WatermarkImageBuilder builder = new();
        WatermarkOptions options = new(
            "CONFIDENTIAL",
            WatermarkPageSize.A4,
            WatermarkPageOrientation.Landscape,
            72
        );

        WatermarkImage image = builder.Build(options);

        Assert.Multiple(() => {
            Assert.That(image.WidthPixels, Is.EqualTo(1169));
            Assert.That(image.HeightPixels, Is.EqualTo(827));
        });
    }
}
