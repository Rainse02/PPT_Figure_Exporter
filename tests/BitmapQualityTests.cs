using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using System.Windows.Media.Imaging;
using PPTFigureExporter;

internal static class BitmapQualityTests
{
    public static void Run(string output, Action<bool, string> assert)
    {
        string folder = Path.Combine(output, "bitmap-quality");
        Directory.CreateDirectory(folder);
        string png = Path.Combine(folder, "original.png");
        string jpeg = Path.Combine(folder, "original.jpg");
        string tiff = Path.Combine(folder, "original.tiff");
        using (var image = new Bitmap(2048, 1024))
        using (Graphics graphics = Graphics.FromImage(image))
        {
            graphics.Clear(Color.White);
            for (int x = 0; x < image.Width; x += 8)
                graphics.FillRectangle(Brushes.Black, x, 0, 4, image.Height);
            image.Save(png, ImageFormat.Png);
            image.Save(jpeg, ImageFormat.Jpeg);
            image.Save(tiff, ImageFormat.Tiff);
        }

        dynamic app = null;
        dynamic presentation = null;
        dynamic slide = null;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            presentation = app.Presentations.Add(0);
            slide = presentation.Slides.Add(1, 12);
            dynamic picture = slide.Shapes.AddPicture(png, 0, -1, 40F, 40F, 144F, 72F);
            picture.Name = "HighResolutionBitmap";
            dynamic label = slide.Shapes.AddTextbox(1, 40F, 120F, 180F, 30F);
            label.Name = "VectorLabel";
            label.TextFrame.TextRange.Text = "Vector text and symbols: alpha beta";
            label.TextFrame.TextRange.Font.Size = 14F;
            dynamic group = slide.Shapes.Range(new object[] { "HighResolutionBitmap", "VectorLabel" }).Group();
            CheckSvg(group, presentation, slide, png, "mixed-group", folder, assert);

            dynamic nestedLabel = slide.Shapes.AddShape(1, 250F, 40F, 10F, 10F);
            nestedLabel.Name = "NestingAnchor";
            string groupName = group.Name;
            dynamic outer = slide.Shapes.Range(new object[] { groupName, "NestingAnchor" }).Group();
            CheckSvg(outer, presentation, slide, png, "nested-group", folder, assert);

            dynamic single = slide.Shapes.AddPicture(png, 0, -1, 40F, 200F, 144F, 72F);
            string baseline = Path.Combine(folder, "single-baseline.svg");
            single.Export(baseline, 6);
            string restored = CheckSvg(single, presentation, slide, png, "single", folder, assert);
            byte[] originalBytes = File.ReadAllBytes(png);
            assert(EmbeddedImages(restored).Any(data => data.SequenceEqual(originalBytes)), "PNG source bytes preserved exactly");
            string oldTiff = ConvertToTiff(baseline, folder, "baseline");
            string fixedTiff = ConvertToTiff(restored, folder, "restored");
            double before = Contrast(oldTiff);
            double after = Contrast(fixedTiff);
            assert(after > 40 && after > before * 3, "TIFF retains source bitmap detail: " + before + " -> " + after);
            using (var result = new Bitmap(fixedTiff))
            {
                assert(result.Width == 2000 && result.Height == 1000, "TIFF output pixel size unchanged");
                assert(Math.Abs(result.HorizontalResolution - 1000) < 0.1, "TIFF output DPI unchanged");
            }

            string nativeTiff = NativeRasterTests.Check((object)single, (object)presentation, folder,
                "native-detail", assert, new ExportSettings
                {
                    Format = "tiff",
                    Dpi = 1000,
                    Transparent = true
                });
            double nativeContrast = Contrast(nativeTiff);
            assert(nativeContrast > 40 && nativeContrast > before * 3,
                "Native TIFF retains source detail above the old SVG preview: " + before + " -> " + nativeContrast);
            NativeRasterTests.Check((object)single, (object)presentation, folder, "custom-stretch",
                assert, new ExportSettings
                {
                    Format = "tiff",
                    SizeMode = "custom",
                    Unit = "inch",
                    Width = 1.25,
                    Height = 1.25,
                    Dpi = 160,
                    Transparent = true
                });


            single.PictureFormat.CropLeft = 12F;
            single.PictureFormat.CropRight = 6F;
            single.Flip(0);
            CheckSvg(single, presentation, slide, png, "cropped-flipped", folder, assert);
            single.Name = "CroppedBitmap";
            dynamic cropLabel = slide.Shapes.AddShape(1, 200F, 200F, 10F, 10F);
            cropLabel.Name = "CropAnchor";
            dynamic croppedGroup = slide.Shapes.Range(new object[] { "CroppedBitmap", "CropAnchor" }).Group();
            CheckSvg(croppedGroup, presentation, slide, png, "cropped-child", folder, assert);
            dynamic jpgPicture = slide.Shapes.AddPicture(jpeg, 0, -1, 250F, 200F, 144F, 72F);
            CheckSvg(jpgPicture, presentation, slide, jpeg, "jpeg", folder, assert);
            dynamic tifPicture = slide.Shapes.AddPicture(tiff, 0, -1, 250F, 300F, 144F, 72F);
            string tifSvg = CheckSvg(tifPicture, presentation, slide, null, "tiff-source", folder, assert);
            using (var stream = new MemoryStream(EmbeddedImages(tifSvg).Single()))
            using (Image decoded = Image.FromStream(stream))
                assert(decoded.Width == 2048 && decoded.Height == 1024, "TIFF source retains full pixel dimensions");

            string grayPath = Path.Combine(folder, "gray16.tiff");
            byte[] grayPixels = new byte[64 * 32 * 2];
            for (int index = 0; index < 64 * 32; index++)
            {
                ushort value = (ushort)((index * 997) % 65536);
                grayPixels[index * 2] = (byte)(value & 255);
                grayPixels[index * 2 + 1] = (byte)(value >> 8);
            }
            var graySource = BitmapSource.Create(64, 32, 96, 96, System.Windows.Media.PixelFormats.Gray16,
                null, grayPixels, 64 * 2);
            var grayEncoder = new TiffBitmapEncoder();
            grayEncoder.Frames.Add(BitmapFrame.Create(graySource));
            using (var file = File.Create(grayPath))
                grayEncoder.Save(file);
            dynamic grayPicture = slide.Shapes.AddPicture(grayPath, 0, -1, 420F, 40F, 144F, 72F);
            string graySvg = CheckSvg(grayPicture, presentation, slide, null, "gray16", folder, assert);
            using (var stream = new MemoryStream(EmbeddedImages(graySvg).Single()))
            {
                BitmapFrame frame = BitmapFrame.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                assert(frame.Format == System.Windows.Media.PixelFormats.Gray16, "16-bit TIFF stays 16-bit in SVG");
                var restoredPixels = new byte[grayPixels.Length];
                frame.CopyPixels(restoredPixels, 64 * 2, 0);
                assert(restoredPixels.SequenceEqual(grayPixels), "16-bit TIFF sample values preserved exactly");
            }

            dynamic fill = slide.Shapes.AddShape(1, 40F, 320F, 144F, 72F);
            fill.Fill.UserPicture(png);
            fill.TextFrame.TextRange.Text = "Picture fill";
            CheckSvg(fill, presentation, slide, png, "picture-fill", folder, assert);
            CheckCompatibilityMatrix(presentation, slide, folder, assert);
            CheckExtendedFormats(presentation, slide, folder, assert);
            assert(Directory.GetFiles(folder, "*.pptx").Length == 0, "Temporary presentation snapshots cleaned");
            Console.WriteLine("Bitmap regression output: " + folder);
            Console.WriteLine("TIFF row contrast (standard deviation): " + before + " -> " + after);
            Console.WriteLine("Native TIFF row contrast (standard deviation): " + before + " -> " + nativeContrast);
        }
        finally
        {
            if (presentation != null)
            {
                presentation.Saved = -1;
                presentation.Close();
            }
            Release(slide);
            Release(presentation);
            Release(app);
        }
    }


    private static void CheckCompatibilityMatrix(dynamic presentation, dynamic slide, string folder,
        Action<bool, string> assert)
    {
        string transparent = Path.Combine(folder, "transparent-blocks.png");
        string bitmap = Path.Combine(folder, "color-blocks.bmp");
        string gif = Path.Combine(folder, "color-blocks.gif");
        string wide = Path.Combine(folder, "wide-blocks.png");
        string tall = Path.Combine(folder, "tall-blocks.png");
        string small = Path.Combine(folder, "small-blocks.png");
        CreateColorBlocks(transparent, 512, 256, ImageFormat.Png, true, false);
        CreateColorBlocks(bitmap, 257, 129, ImageFormat.Bmp, false, false);
        CreateColorBlocks(gif, 384, 192, ImageFormat.Gif, false, true);
        CreateColorBlocks(wide, 2048, 32, ImageFormat.Png, false, false);
        CreateColorBlocks(tall, 32, 2048, ImageFormat.Png, false, false);
        CreateColorBlocks(small, 9, 7, ImageFormat.Png, false, false);

        var cases = new[]
        {
            new { Path = transparent, Width = 512, Height = 256, Name = "transparent-png", ExactBytes = true },
            new { Path = bitmap, Width = 257, Height = 129, Name = "bmp", ExactBytes = false },
            new { Path = gif, Width = 384, Height = 192, Name = "gif", ExactBytes = false },
            new { Path = wide, Width = 2048, Height = 32, Name = "wide-image", ExactBytes = true },
            new { Path = tall, Width = 32, Height = 2048, Name = "tall-image", ExactBytes = true },
            new { Path = small, Width = 9, Height = 7, Name = "small-image", ExactBytes = true }
        };
        foreach (var item in cases)
        {
            dynamic picture = slide.Shapes.AddPicture(item.Path, 0, -1, 40F, 40F, 144F, 72F);
            string restored = CheckSvg(picture, presentation, slide,
                item.ExactBytes ? item.Path : null, item.Name, folder, assert);
            assert(EmbeddedImages(restored).Any(data => HasDimensions(data, item.Width, item.Height)),
                item.Name + ": original pixel dimensions retained");
            string rendered = RenderComparison(restored, folder, item.Name + "-rendered",
                (double)picture.Width, (double)picture.Height);
            CheckSourcePatches(item.Path, rendered, item.Name, assert);
            Release((object)picture);
        }

        dynamic rotated = slide.Shapes.AddPicture(transparent, 0, -1, 40F, 40F, 144F, 72F);
        rotated.Rotation = 31F;
        CheckAppearance(rotated, presentation, slide, transparent, "rotated-transparent", folder, assert);

        dynamic cropped = slide.Shapes.AddPicture(bitmap, 0, -1, 240F, 40F, 144F, 72F);
        cropped.PictureFormat.CropLeft = 11F;
        cropped.PictureFormat.CropTop = 5F;
        cropped.Flip(1);
        cropped.Name = "MatrixCrop";
        dynamic anchor = slide.Shapes.AddShape(1, 240F, 120F, 12F, 12F);
        anchor.Name = "MatrixCropAnchor";
        dynamic croppedGroup = slide.Shapes.Range(new object[] { "MatrixCrop", "MatrixCropAnchor" }).Group();
        croppedGroup.Rotation = 17F;
        CheckAppearance(croppedGroup, presentation, slide, null, "rotated-group-crop", folder, assert);

        dynamic fill = slide.Shapes.AddShape(1, 40F, 40F, 144F, 72F);
        fill.Line.Visible = 0;
        fill.Fill.UserPicture(transparent);
        fill.Rotation = 23F;
        CheckAppearance(fill, presentation, slide, transparent, "transparent-picture-fill", folder, assert);

        dynamic first = slide.Shapes.AddPicture(transparent, 0, -1, 40F, 40F, 144F, 72F);
        dynamic repeated = slide.Shapes.AddPicture(transparent, 0, -1, 220F, 40F, 144F, 72F);
        dynamic different = slide.Shapes.AddPicture(gif, 0, -1, 400F, 40F, 144F, 72F);
        first.Name = "MatrixFirst";
        repeated.Name = "MatrixRepeated";
        different.Name = "MatrixDifferent";
        dynamic duplicates = slide.Shapes.Range(new object[]
        {
            "MatrixFirst", "MatrixRepeated", "MatrixDifferent"
        }).Group();
        first.Name = "RepeatedPictureName";
        repeated.Name = "RepeatedPictureName";
        different.Name = "RepeatedPictureName";
        string duplicateSvg = CheckAppearance(duplicates, presentation, slide, transparent,
            "repeated-resource-and-name", folder, assert);
        assert(EmbeddedImages(duplicateSvg).Any(data => HasDimensions(data, 384, 192)),
            "Repeated names: distinct GIF resource remains distinct");

        dynamic brightness = slide.Shapes.AddPicture(bitmap, 0, -1, 40F, 40F, 144F, 72F);
        brightness.PictureFormat.Brightness = 0.68F;
        string brightSvg = CheckAppearance(brightness, presentation, slide, null,
            "picture-brightness", folder, assert);
        assert(EmbeddedImages(brightSvg).Any(data => HasDimensions(data, 257, 129)),
            "Brightness correction retains original pixel dimensions");

        dynamic contrast = slide.Shapes.AddPicture(bitmap, 0, -1, 240F, 40F, 144F, 72F);
        contrast.PictureFormat.Contrast = 0.72F;
        string contrastSvg = CheckAppearance(contrast, presentation, slide, null,
            "picture-contrast", folder, assert);
        assert(EmbeddedImages(contrastSvg).Any(data => HasDimensions(data, 257, 129)),
            "Contrast correction retains original pixel dimensions");

        Release((object)contrast);
        Release((object)brightness);
        Release((object)duplicates);
        Release((object)different);
        Release((object)repeated);
        Release((object)first);
        Release((object)fill);
        Release((object)croppedGroup);
        Release((object)anchor);
        Release((object)cropped);
        Release((object)rotated);
    }


    private static void CheckExtendedFormats(dynamic presentation, dynamic slide, string folder,
        Action<bool, string> assert)
    {
        string bitmap = Path.Combine(folder, "color-blocks.bmp");
        string gray = Path.Combine(folder, "gray8.png");
        string cmyk = Path.Combine(folder, "cmyk.jpg");
        string jpegXr = Path.Combine(folder, "lossless-jpeg-xr.wdp");
        CreateGrayPng(gray);
        CreateCmykJpeg(cmyk);

        string graySvg = CheckEncodedPicture(presentation, slide, gray, gray, gray,
            192, 96, "gray8-png", folder, assert);
        assert(EmbeddedImages(graySvg).Any(data => HasPixelFormat(data, System.Windows.Media.PixelFormats.Gray8)),
            "Grayscale PNG retains its 8-bit sample format");
        CheckEncodedPicture(presentation, slide, cmyk, cmyk, null,
            192, 96, "cmyk-jpeg", folder, assert);
        if (TryCreateJpegXr(bitmap, jpegXr))
            CheckEncodedPicture(presentation, slide, jpegXr, bitmap, null,
                257, 129, "jpeg-xr", folder, assert);

        foreach (int colorType in new[] { 2, 3 })
        {
            dynamic picture = slide.Shapes.AddPicture(bitmap, 0, -1, 40F, 40F, 144F, 72F);
            try
            {
                picture.PictureFormat.ColorType = colorType;
                string name = colorType == 2 ? "picture-grayscale" : "picture-black-and-white";
                string svg = CheckAppearance(picture, presentation, slide, null, name, folder, assert);
                assert((int)picture.PictureFormat.ColorType == colorType,
                    name + ": source color operation unchanged");
                assert(EmbeddedImages(svg).Any(data => HasDimensions(data, 257, 129)),
                    name + ": original pixel dimensions retained");
                using (var image = new Bitmap(Path.Combine(folder, name + "-rendered.png")))
                {
                    for (int block = 0; block < 4; block++)
                    {
                        double[] channels = PatchMean(image, block % 2 * 2 + 1, block / 2 * 2 + 1);
                        assert(Math.Abs(channels[0] - channels[1]) <= 2 &&
                            Math.Abs(channels[1] - channels[2]) <= 2,
                            name + ": neutral rendered channels in block " + block);
                        if (colorType == 3)
                            assert(channels.Take(3).All(value => value <= 7 || value >= 248),
                                name + ": black or white rendered block " + block);
                    }
                }
            }
            finally
            {
                Release((object)picture);
            }
        }

        dynamic shadowed = slide.Shapes.AddPicture(bitmap, 0, -1, 40F, 40F, 144F, 72F);
        try
        {
            shadowed.Shadow.Type = 1;
            shadowed.Shadow.Visible = -1;
            shadowed.Shadow.Transparency = 0.35F;
            shadowed.Shadow.OffsetX = 6F;
            shadowed.Shadow.OffsetY = 5F;
            shadowed.Shadow.Blur = 4F;
            string shadowSvg = CheckAppearance(shadowed, presentation, slide, null,
                "picture-shadow", folder, assert);
            assert(EmbeddedImages(shadowSvg).Any(data => HasDimensions(data, 257, 129)),
                "Picture shadow: original pixel dimensions retained");
        }
        finally
        {
            Release((object)shadowed);
        }
    }

    private static string CheckEncodedPicture(dynamic presentation, dynamic slide, string source,
        string patchSource, string exactSource, int width, int height, string name, string folder,
        Action<bool, string> assert)
    {
        dynamic picture = slide.Shapes.AddPicture(source, 0, -1, 40F, 40F, 144F, 72F);
        try
        {
            string svg = CheckSvg(picture, presentation, slide, exactSource, name, folder, assert);
            assert(EmbeddedImages(svg).Any(data => HasDimensions(data, width, height)),
                name + ": original pixel dimensions retained");
            string rendered = RenderComparison(svg, folder, name + "-rendered",
                (double)picture.Width, (double)picture.Height);
            CheckSourcePatches(patchSource, rendered, name, assert);
            return svg;
        }
        finally
        {
            Release((object)picture);
        }
    }

    private static void CreateGrayPng(string path)
    {
        const int width = 192;
        const int height = 96;
        byte[] shades = { 32, 112, 176, 232 };
        var pixels = new byte[width * height];
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int row = y < height / 2 ? 0 : 1;
                int column = x < width / 2 ? 0 : 1;
                int block = row * 2 + column;
                pixels[y * width + x] = shades[block];
            }
        }
        BitmapSource source = BitmapSource.Create(width, height, 96, 96,
            System.Windows.Media.PixelFormats.Gray8, null, pixels, width);
        SaveEncodedFrame(path, source, new PngBitmapEncoder());
    }

    private static void CreateCmykJpeg(string path)
    {
        const int width = 192;
        const int height = 96;
        byte[][] colors = { new byte[] { 20, 100, 160, 40 }, new byte[] { 160, 40, 20, 24 },
            new byte[] { 60, 160, 60, 32 }, new byte[] { 80, 40, 160, 64 } };
        var pixels = new byte[width * height * 4];
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int row = y < height / 2 ? 0 : 1;
                int column = x < width / 2 ? 0 : 1;
                int block = row * 2 + column;
                Buffer.BlockCopy(colors[block], 0, pixels, (y * width + x) * 4, 4);
            }
        }
        BitmapSource source = BitmapSource.Create(width, height, 96, 96,
            System.Windows.Media.PixelFormats.Cmyk32, null, pixels, width * 4);
        SaveEncodedFrame(path, source, new JpegBitmapEncoder { QualityLevel = 98 });
    }

    private static bool TryCreateJpegXr(string sourcePath, string path)
    {
        try
        {
            using (var input = File.OpenRead(sourcePath))
            {
                BitmapFrame source = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnLoad);
                SaveEncodedFrame(path, source, new WmpBitmapEncoder { Lossless = true });
            }
            return true;
        }
        catch (Exception exception) when (exception is NotSupportedException ||
            exception is COMException || exception is FileFormatException)
        {
            Console.WriteLine("SKIP JPEG XR: local WIC encoder unavailable: " + exception.Message);
            return false;
        }
    }

    private static void SaveEncodedFrame(string path, BitmapSource source, BitmapEncoder encoder)
    {
        encoder.Frames.Add(BitmapFrame.Create(source));
        using (var output = new MemoryStream())
        {
            encoder.Save(output);
            File.WriteAllBytes(path, output.ToArray());
        }
    }

    private static bool HasPixelFormat(byte[] data, System.Windows.Media.PixelFormat expected)
    {
        using (var input = new MemoryStream(data, false))
        {
            BitmapFrame frame = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat,
                BitmapCacheOption.OnLoad);
            return frame.Format == expected;
        }
    }

    private static void CreateColorBlocks(string path, int width, int height, ImageFormat format,
        bool transparent, bool reverseColors)
    {
        using (var image = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        using (Graphics graphics = Graphics.FromImage(image))
        {
            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
            for (int row = 0; row < 2; row++)
            {
                for (int column = 0; column < 2; column++)
                {
                    int index = row * 2 + column;
                    Color color = BlockColor(reverseColors ? 3 - index : index, transparent);
                    int left = column * width / 2;
                    int top = row * height / 2;
                    int right = (column + 1) * width / 2;
                    int bottom = (row + 1) * height / 2;
                    using (var brush = new SolidBrush(color))
                        graphics.FillRectangle(brush, left, top, right - left, bottom - top);
                }
            }
            image.Save(path, format);
        }
    }

    private static Color BlockColor(int index, bool transparent)
    {
        Color[] colors = { Color.FromArgb(36, 76, 120), Color.FromArgb(180, 120, 50),
            Color.FromArgb(70, 160, 100), Color.FromArgb(210, 195, 165) };
        int[] alpha = { 0, 96, 192, 255 };
        return Color.FromArgb(transparent ? alpha[index] : 255, colors[index]);
    }

    private static bool HasDimensions(byte[] data, int width, int height)
    {
        using (var input = new MemoryStream(data, false))
        {
            BitmapFrame frame = BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat,
                BitmapCacheOption.OnLoad);
            return frame.PixelWidth == width && frame.PixelHeight == height;
        }
    }

    private static string CheckAppearance(dynamic shape, dynamic presentation, dynamic slide,
        string expectedSource, string name, string folder, Action<bool, string> assert)
    {
        string restored = CheckSvg(shape, presentation, slide, expectedSource, name, folder, assert);
        string baseline = RenderComparison(Path.Combine(folder, name + "-before.svg"), folder,
            name + "-baseline-rendered", (double)shape.Width, (double)shape.Height);
        string result = RenderComparison(restored, folder, name + "-rendered",
            (double)shape.Width, (double)shape.Height);
        double maximumCellError;
        double meanError = AppearanceDifference(baseline, result, out maximumCellError);
        assert(meanError <= 8 && maximumCellError <= 36,
            name + ": rendered colors and alpha match PowerPoint SVG (mean " + meanError.ToString("F2") +
            ", maximum cell " + maximumCellError.ToString("F2") + ")");
        return restored;
    }

    private static string RenderComparison(string svg, string folder, string name, double width, double height)
    {
        string output = Path.Combine(folder, name + ".png");
        string request = Path.Combine(folder, name + ".json");
        File.WriteAllText(request, new JavaScriptSerializer().Serialize(new
        {
            input_svg = svg,
            output_path = output,
            format = "png",
            source_width_pt = width,
            source_height_pt = height,
            size_mode = "keep",
            dpi = 144,
            transparent = true,
            overwrite = true
        }), new UTF8Encoding(false));
        ExportService.RunAsync(request, CancellationToken.None).GetAwaiter().GetResult();
        return output;
    }

    private static double AppearanceDifference(string baseline, string restored, out double maximumCellError)
    {
        using (var before = new Bitmap(baseline))
        using (var after = new Bitmap(restored))
        {
            if (before.Width != after.Width || before.Height != after.Height)
                throw new InvalidOperationException("Comparison renders have different dimensions.");
            double total = 0;
            maximumCellError = 0;
            // Block averages compare appearance without requiring identical resampling at image edges.
            const int cells = 16;
            for (int row = 0; row < cells; row++)
            {
                for (int column = 0; column < cells; column++)
                {
                    double[] beforeMean = MeanChannels(before, column, row, cells);
                    double[] afterMean = MeanChannels(after, column, row, cells);
                    double error = 0;
                    for (int channel = 0; channel < beforeMean.Length; channel++)
                        error += Math.Abs(beforeMean[channel] - afterMean[channel]);
                    error /= beforeMean.Length;
                    total += error;
                    maximumCellError = Math.Max(maximumCellError, error);
                }
            }
            return total / (cells * cells);
        }
    }

    private static double[] MeanChannels(Bitmap image, int column, int row, int cells)
    {
        var result = new double[4];
        int left = column * image.Width / cells;
        int top = row * image.Height / cells;
        int right = (column + 1) * image.Width / cells;
        int bottom = (row + 1) * image.Height / cells;
        int count = (right - left) * (bottom - top);
        for (int y = top; y < bottom; y++)
        {
            for (int x = left; x < right; x++)
            {
                Color pixel = image.GetPixel(x, y);
                // Transparent RGB values are undefined; compare the visible premultiplied channels.
                double alpha = pixel.A / 255.0;
                result[0] += pixel.R * alpha;
                result[1] += pixel.G * alpha;
                result[2] += pixel.B * alpha;
                result[3] += pixel.A;
            }
        }
        for (int channel = 0; channel < result.Length; channel++)
            result[channel] /= count;
        return result;
    }

    private static void CheckSourcePatches(string source, string rendered, string name, Action<bool, string> assert)
    {
        // Use original constant-color regions, averaging palette dithering in GIF sources.
        // The old COM preview may have only a few JPEG rows for extreme aspect ratios.
        using (var original = new Bitmap(source))
        using (var image = new Bitmap(rendered))
        {
            for (int index = 0; index < 4; index++)
            {
                int xFraction = index % 2 * 2 + 1;
                int yFraction = index / 2 * 2 + 1;
                double[] expected = PatchMean(original, xFraction, yFraction);
                double[] actual = PatchMean(image, xFraction, yFraction);
                assert(expected.Zip(actual, (left, right) => Math.Abs(left - right)).All(error => error <= 7),
                    name + ": original color and alpha in block " + index);
            }
        }
    }

    private static double[] PatchMean(Bitmap image, int xFraction, int yFraction)
    {
        int centerX = xFraction * image.Width / 4;
        int centerY = yFraction * image.Height / 4;
        int radiusX = image.Width / 16;
        int radiusY = image.Height / 16;
        int left = Math.Max(0, centerX - radiusX);
        int right = Math.Min(image.Width - 1, centerX + radiusX);
        int top = Math.Max(0, centerY - radiusY);
        int bottom = Math.Min(image.Height - 1, centerY + radiusY);
        var channels = new double[4];
        for (int y = top; y <= bottom; y++)
        {
            for (int x = left; x <= right; x++)
            {
                Color pixel = image.GetPixel(x, y);
                double alpha = pixel.A / 255.0;
                channels[0] += pixel.R * alpha;
                channels[1] += pixel.G * alpha;
                channels[2] += pixel.B * alpha;
                channels[3] += pixel.A;
            }
        }
        int count = (right - left + 1) * (bottom - top + 1);
        return channels.Select(channel => channel / count).ToArray();
    }

    private static string CheckSvg(dynamic shape, dynamic presentation, dynamic slide, string expectedSource,
        string name, string folder, Action<bool, string> assert)
    {
        string baseline = Path.Combine(folder, name + "-before.svg");
        string restored = Path.Combine(folder, name + ".svg");
        shape.Export(baseline, 6);
        string before = Snapshot(shape, presentation);
        PowerPointSvgExporter.Export((object)shape, (object)presentation, (int)slide.SlideID, restored);
        assert(Snapshot(shape, presentation) == before, name + ": source state and geometry unchanged");
        assert(Structure(baseline) == Structure(restored), name + ": SVG text, clipping and transforms unchanged");
        if (expectedSource != null)
        {
            byte[] expected = File.ReadAllBytes(expectedSource);
            assert(EmbeddedImages(restored).Any(data => data.SequenceEqual(expected)), name + ": original bitmap file embedded");
        }
        NativeRasterTests.Check((object)shape, (object)presentation, folder, name, assert);
        return restored;
    }

    private static string Snapshot(dynamic shape, dynamic presentation)
    {
        return NativeRasterTests.Snapshot((object)shape, (object)presentation);
    }

    private static string Structure(string path)
    {
        XDocument document = XDocument.Load(path);
        foreach (var image in document.Descendants().Where(node => node.Name.LocalName == "image"))
            foreach (var attribute in image.Attributes().Where(attribute => attribute.Name.LocalName == "href"))
                attribute.Value = "[bitmap]";
        return document.Root.ToString(SaveOptions.DisableFormatting);
    }

    private static byte[][] EmbeddedImages(string path)
    {
        return XDocument.Load(path).Descendants().Where(node => node.Name.LocalName == "image")
            .SelectMany(node => node.Attributes().Where(attribute => attribute.Name.LocalName == "href"))
            .Select(attribute => attribute.Value).Where(value => value.StartsWith("data:image/"))
            .Select(value => Convert.FromBase64String(value.Substring(value.IndexOf(',') + 1))).ToArray();
    }

    private static string ConvertToTiff(string svg, string folder, string name)
    {
        string output = Path.Combine(folder, name + ".tiff");
        string request = Path.Combine(folder, name + ".json");
        File.WriteAllText(request, new JavaScriptSerializer().Serialize(new
        {
            input_svg = svg, output_path = output, format = "tiff", source_width_pt = 144,
            source_height_pt = 72, size_mode = "keep", dpi = 1000, transparent = false, overwrite = true
        }), new UTF8Encoding(false));
        ExportService.RunAsync(request, CancellationToken.None).GetAwaiter().GetResult();
        return output;
    }

    private static double Contrast(string path)
    {
        using (var image = new Bitmap(path))
        {
            double sum = 0;
            double squares = 0;
            int samples = image.Width - 100;
            for (int x = 50; x < image.Width - 50; x++)
            {
                double value = image.GetPixel(x, image.Height / 2).R;
                sum += value;
                squares += value * value;
            }
            return Math.Sqrt(Math.Max(0, squares / samples - Math.Pow(sum / samples, 2)));
        }
    }

    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
