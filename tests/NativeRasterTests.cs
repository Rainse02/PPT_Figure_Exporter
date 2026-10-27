using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Media.Imaging;
using PPTFigureExporter;

internal static class NativeRasterTests
{
    public static string Check(object shape, object presentation, string folder, string name,
        Action<bool, string> assert, ExportSettings settings = null)
    {
        settings = settings ?? new ExportSettings { Format = "tiff", Dpi = 144, Transparent = true };
        double width = (double)((dynamic)shape).Width;
        double height = (double)((dynamic)shape).Height;
        SizePoints target = settings.Target(width, height);
        bool stretch = settings.SizeMode == "custom";
        string png = Path.Combine(folder, name + "-native.png");
        string tiff = Path.Combine(folder, name + "-native.tiff");
        string request = Path.Combine(folder, name + "-native-request.json");
        string before = Snapshot(shape, presentation);

        PowerPointRasterExporter.Export(shape, presentation, png, target, stretch);
        assert(Snapshot(shape, presentation) == before,
            name + ": native export preserves source objects and saved state");

        BitmapFrame rendered = ReadFrame(png);
        assert(rendered.PixelWidth > 0 && rendered.PixelWidth <= target.PixelWidth &&
            rendered.PixelHeight > 0 && rendered.PixelHeight <= target.PixelHeight,
            name + ": native PNG fits the target canvas");
        int gapX = target.PixelWidth - rendered.PixelWidth;
        int gapY = target.PixelHeight - rendered.PixelHeight;
        assert(stretch ? gapX <= 2 && gapY <= 2 : Math.Min(gapX, gapY) <= 2,
            name + ": native render fills the available target dimensions");

        File.WriteAllText(request, new JavaScriptSerializer().Serialize(new
        {
            input_path = png,
            output_path = tiff,
            format = "tiff",
            source_width_pt = width,
            source_height_pt = height,
            size_mode = settings.SizeMode,
            width = settings.Width,
            height = settings.Height,
            unit = settings.Unit,
            percent = settings.Percent,
            dpi = settings.Dpi,
            transparent = settings.Transparent,
            overwrite = true
        }), new UTF8Encoding(false));
        ExportService.RunAsync(request, CancellationToken.None).GetAwaiter().GetResult();

        BitmapFrame result = ReadFrame(tiff);
        assert(result.PixelWidth == target.PixelWidth && result.PixelHeight == target.PixelHeight,
            name + ": TIFF has the exact requested pixel dimensions");
        assert(Math.Abs(result.DpiX - settings.Dpi) < 0.05 &&
            Math.Abs(result.DpiY - settings.Dpi) < 0.05,
            name + ": TIFF has the requested horizontal and vertical DPI");
        assert(settings.Transparent, name + ": exact pixel regression uses transparent output");

        byte[] expected = CenteredPixels(rendered, target.PixelWidth, target.PixelHeight);
        byte[] actual = Pixels(result);
        int mismatch = FirstMismatch(expected, actual);
        string location = mismatch < 0 ? "" : " at (" + mismatch / 4 % target.PixelWidth + ", " +
            mismatch / 4 / target.PixelWidth + "), BGRA channel " + mismatch % 4 +
            ": expected " + expected[mismatch] + ", actual " + actual[mismatch];
        assert(mismatch < 0, name + ": TIFF preserves centered native PNG pixels exactly" + location);
        return tiff;
    }

    internal static string Snapshot(object shape, object presentation)
    {
        var values = new List<string>();
        object slides = null;
        object pageSetup = null;
        try
        {
            slides = ((dynamic)presentation).Slides;
            pageSetup = ((dynamic)presentation).PageSetup;
            Append(values, ((dynamic)presentation).Saved, ((dynamic)slides).Count,
                ((dynamic)pageSetup).SlideWidth, ((dynamic)pageSetup).SlideHeight);
            CaptureShape(shape, values);
            return string.Join("|", values);
        }
        finally
        {
            Release(pageSetup);
            Release(slides);
        }
    }

    private static void CaptureShape(object shape, List<string> values)
    {
        dynamic current = shape;
        int type = (int)current.Type;
        Append(values, current.Id, type, current.Left, current.Top, current.Width, current.Height,
            current.Rotation, current.Name, current.Visible, current.HorizontalFlip, current.VerticalFlip);
        if (type == 11 || type == 13)
        {
            object picture = null;
            try
            {
                picture = current.PictureFormat;
                Append(values, ((dynamic)picture).CropLeft, ((dynamic)picture).CropTop,
                    ((dynamic)picture).CropRight, ((dynamic)picture).CropBottom,
                    ((dynamic)picture).Brightness, ((dynamic)picture).Contrast, ((dynamic)picture).ColorType);
            }
            finally
            {
                Release(picture);
            }
        }
        if (type != 6)
            return;

        object children = null;
        try
        {
            children = current.GroupItems;
            int count = (int)((dynamic)children).Count;
            Append(values, count);
            for (int index = 1; index <= count; index++)
            {
                object child = null;
                try
                {
                    child = ((dynamic)children)[index];
                    CaptureShape(child, values);
                }
                finally
                {
                    Release(child);
                }
            }
        }
        finally
        {
            Release(children);
        }
    }

    private static void Append(List<string> values, params object[] properties)
    {
        foreach (object property in properties)
            values.Add(Convert.ToString(property, CultureInfo.InvariantCulture));
    }

    private static BitmapFrame ReadFrame(string path)
    {
        using (var input = File.OpenRead(path))
            return BitmapFrame.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
    }

    private static byte[] Pixels(BitmapSource image)
    {
        // Straight BGRA avoids GDI+ alpha premultiplication and TIFF color conversion.
        var converted = new FormatConvertedBitmap(image, System.Windows.Media.PixelFormats.Bgra32, null, 0);
        int stride = image.PixelWidth * 4;
        var result = new byte[stride * image.PixelHeight];
        converted.CopyPixels(result, stride, 0);
        return result;
    }

    private static byte[] CenteredPixels(BitmapSource image, int width, int height)
    {
        byte[] source = Pixels(image);
        var result = new byte[width * height * 4];
        int offsetX = (width - image.PixelWidth) / 2;
        int offsetY = (height - image.PixelHeight) / 2;
        int sourceStride = image.PixelWidth * 4;
        for (int row = 0; row < image.PixelHeight; row++)
        {
            int destination = ((row + offsetY) * width + offsetX) * 4;
            Buffer.BlockCopy(source, row * sourceStride, result, destination, sourceStride);
        }
        return result;
    }

    private static int FirstMismatch(byte[] expected, byte[] actual)
    {
        if (expected.Length != actual.Length)
            throw new InvalidOperationException("PNG and TIFF comparison buffers have different dimensions.");
        for (int index = 0; index < expected.Length; index++)
        {
            if (expected[index] != actual[index])
                return index;
        }
        return -1;
    }

    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
