using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PPTFigureExporter
{
    internal static class PowerPointRasterExporter
    {
        public static void Export(object shape, object presentation, string path, SizePoints target, bool stretch)
        {
            if (target.PixelWidth < 1 || target.PixelHeight < 1)
                throw new ArgumentException("目标像素尺寸必须大于零。", "target");

            object pageSetup = null;
            string candidate = Path.Combine(Path.GetDirectoryName(path), Guid.NewGuid().ToString("N") + ".png");
            try
            {
                pageSetup = ((dynamic)presentation).PageSetup;
                double slideWidth = (double)((dynamic)pageSetup).SlideWidth;
                double slideHeight = (double)((dynamic)pageSetup).SlideHeight;
                double width = (double)((dynamic)shape).Width;
                double height = (double)((dynamic)shape).Height;

                // Shape.Export's scale arguments describe the slide, not the shape.
                // A small native render measures the actual bounds, including rotation,
                // shadows and text overhang. Its pixels are never used in the output.
                double pilotScale = Math.Min(512.0 / Math.Max(width, height),
                    Math.Min(target.PixelWidth / width, target.PixelHeight / height));
                double scaleX = slideWidth * 72 / 96 * pilotScale;
                double scaleY = slideHeight * 72 / 96 * pilotScale;
                scaleX = ScaleArgument(scaleX);
                scaleY = ScaleArgument(scaleY);
                int[] measured = Render(shape, candidate, (int)scaleX, (int)scaleY);
                AdjustScale(ref scaleX, ref scaleY, measured, target, stretch);

                long bestArea = 0;
                int previousX = 0;
                int previousY = 0;
                for (int attempt = 0; attempt < 6; attempt++)
                {
                    int argumentX = ScaleArgument(scaleX);
                    int argumentY = ScaleArgument(scaleY);
                    if (argumentX == previousX && argumentY == previousY)
                        break;
                    previousX = argumentX;
                    previousY = argumentY;
                    measured = Render(shape, candidate, argumentX, argumentY);
                    bool fits = measured[0] <= target.PixelWidth && measured[1] <= target.PixelHeight;
                    long area = (long)measured[0] * measured[1];
                    if (fits && area > bestArea)
                    {
                        File.Copy(candidate, path, true);
                        bestArea = area;
                    }
                    int gapX = target.PixelWidth - measured[0];
                    int gapY = target.PixelHeight - measured[1];
                    if (fits && (stretch ? gapX <= 1 && gapY <= 1 : Math.Min(gapX, gapY) <= 1))
                        return;

                    // Re-render at corrected native dimensions; do not resize a PNG.
                    scaleX = argumentX;
                    scaleY = argumentY;
                    AdjustScale(ref scaleX, ref scaleY, measured, target, stretch);
                }
                if (bestArea == 0)
                    throw new InvalidOperationException("PowerPoint 无法按目标像素范围导出此对象，请调整尺寸或 DPI。");
            }
            finally
            {
                if (File.Exists(candidate))
                    File.Delete(candidate);
                if (pageSetup != null && Marshal.IsComObject(pageSetup))
                    Marshal.ReleaseComObject(pageSetup);
            }
        }

        private static void AdjustScale(ref double scaleX, ref double scaleY, int[] actual,
            SizePoints target, bool stretch)
        {
            // Leave a fractional-pixel margin only when rounding exceeded the canvas.
            double width = target.PixelWidth - (actual[0] > target.PixelWidth ? 0.5 : 0);
            double height = target.PixelHeight - (actual[1] > target.PixelHeight ? 0.5 : 0);
            double ratioX = width / actual[0];
            double ratioY = height / actual[1];
            if (!stretch)
                ratioX = ratioY = Math.Min(ratioX, ratioY);
            scaleX *= ratioX;
            scaleY *= ratioY;
        }

        private static int ScaleArgument(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value > int.MaxValue)
                throw new ArgumentException("目标尺寸超出 PowerPoint 的导出范围，请减小尺寸或 DPI。");
            return Math.Max(1, (int)Math.Round(value, MidpointRounding.AwayFromZero));
        }

        private static int[] Render(object shape, string path, int scaleWidth, int scaleHeight)
        {
            // ppShapeFormatPNG = 2; ppRelativeToSlide = 1.
            ((dynamic)shape).Export(path, 2, scaleWidth, scaleHeight, 1);
            using (var input = File.OpenRead(path))
            {
                var header = new byte[24];
                if (input.Read(header, 0, header.Length) != header.Length ||
                    header[0] != 137 || header[1] != 80 || header[2] != 78 || header[3] != 71 ||
                    header[4] != 13 || header[5] != 10 || header[6] != 26 || header[7] != 10 ||
                    header[12] != 73 || header[13] != 72 || header[14] != 68 || header[15] != 82)
                    throw new IOException("PowerPoint 未生成有效的 PNG 图像。");
                int width = ReadBigEndianInt(header, 16);
                int height = ReadBigEndianInt(header, 20);
                if (width < 1 || height < 1)
                    throw new IOException("PowerPoint 返回了无效的 PNG 像素尺寸。");
                return new[] { width, height };
            }
        }

        private static int ReadBigEndianInt(byte[] bytes, int offset)
        {
            return bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];
        }
    }
}
