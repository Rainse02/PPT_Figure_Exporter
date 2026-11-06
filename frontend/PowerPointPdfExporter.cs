using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Media.Imaging;

namespace PPTFigureExporter
{
    internal sealed class PdfCrop
    {
        public double left_pt { get; set; }
        public double top_pt { get; set; }
        public double width_pt { get; set; }
        public double height_pt { get; set; }
    }

    internal static class PowerPointPdfExporter
    {
        private const double FrameUnitsPerInch = 2540;
        private const double CanvasPadding = 1;
        private const double MaximumCanvas = 4032;
        private const double ImagePpiLimit = 180;

        public static PdfCrop Export(object shape, object presentation, int slideId, string path)
        {
            string folder = Path.Combine(Path.GetDirectoryName(path), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            object application = null;
            object presentations = null;
            object copy = null;
            object pageSetup = null;
            object slides = null;
            object slide = null;
            object background = null;
            object fill = null;
            object shapes = null;
            object figure = null;
            object printOptions = null;
            object ranges = null;
            object range = null;
            try
            {
                // Capture the entire selected unit natively. EMF keeps drawing commands,
                // glyph placement and bitmap pixels together, including group transformations.
                // Scaling this unit also scales fonts, strokes and effects consistently.
                string metafile = Path.Combine(folder, "figure.emf");
                PdfCrop bounds = ExportMetafile(shape, metafile);
                string source = (string)((dynamic)presentation).FullName;
                bool savedPackage = (int)((dynamic)presentation).Saved == -1 && File.Exists(source) &&
                    new[] { ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm" }
                        .Contains(Path.GetExtension(source).ToLowerInvariant());
                if (!savedPackage)
                {
                    source = Path.Combine(folder, "source-snapshot.pptx");
                    ((dynamic)presentation).SaveCopyAs(source, 24, 0);
                }
                List<NativePicture> pictures = PresentationPictures.Read(source, slideId, (int)((dynamic)shape).Id);
                double bitmapScale = RequiredBitmapScale(shape, pictures);
                double maximumScale = Math.Min(
                    (MaximumCanvas - 2 * CanvasPadding) / bounds.width_pt,
                    (MaximumCanvas - 2 * CanvasPadding) / bounds.height_pt);
                if (bitmapScale > maximumScale)
                    throw new InvalidOperationException("保持原图像素所需的 PDF 临时画布超出 PowerPoint 的尺寸上限，已停止导出以避免降低图片清晰度。");
                double scale = Math.Min(maximumScale, Math.Max(3, bitmapScale));
                var crop = new PdfCrop
                {
                    left_pt = CanvasPadding,
                    top_pt = CanvasPadding,
                    width_pt = bounds.width_pt * scale,
                    height_pt = bounds.height_pt * scale
                };

                application = ((dynamic)presentation).Application;
                presentations = ((dynamic)application).Presentations;
                copy = ((dynamic)presentations).Add(0);
                pageSetup = ((dynamic)copy).PageSetup;
                // PowerPoint requires a one-inch minimum page. The core crops this canvas
                // to the measured figure bounds and sets the requested physical dimensions.
                ((dynamic)pageSetup).SlideWidth = (float)Math.Max(72, Math.Ceiling(crop.width_pt + 2 * CanvasPadding));
                ((dynamic)pageSetup).SlideHeight = (float)Math.Max(72, Math.Ceiling(crop.height_pt + 2 * CanvasPadding));
                slides = ((dynamic)copy).Slides;
                slide = ((dynamic)slides).Add(1, 12);
                ((dynamic)slide).DisplayMasterShapes = 0;
                ((dynamic)slide).FollowMasterBackground = 0;
                background = ((dynamic)slide).Background;
                fill = ((dynamic)background).Fill;
                ((dynamic)fill).Visible = 0;
                shapes = ((dynamic)slide).Shapes;
                figure = ((dynamic)shapes).AddPicture(metafile, 0, -1,
                    (float)crop.left_pt, (float)crop.top_pt, (float)crop.width_pt, (float)crop.height_pt);
                printOptions = ((dynamic)copy).PrintOptions;
                ((dynamic)printOptions).HighQuality = -1;
                ranges = ((dynamic)printOptions).Ranges;
                range = ((dynamic)ranges).Add(1, 1);

                // PDF, print intent, one slide. Source objects and print settings are untouched.
                ((dynamic)copy).ExportAsFixedFormat(path, 2, 2, 0, 1, 1, 0, range, 4,
                    "", false, true, false, true, false);
                using (var output = File.OpenRead(path))
                {
                    var signature = new byte[5];
                    if (output.Read(signature, 0, signature.Length) != signature.Length ||
                        Encoding.ASCII.GetString(signature) != "%PDF-")
                        throw new IOException("PowerPoint 未生成有效的 PDF 文件。");
                }
                return crop;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        ((dynamic)copy).Saved = -1;
                        ((dynamic)copy).Close();
                    }
                    catch (COMException ex) { SettingsStore.Log("关闭 PDF 临时文稿失败。", ex); }
                }
                Release(range);
                Release(ranges);
                Release(printOptions);
                Release(figure);
                Release(shapes);
                Release(fill);
                Release(background);
                Release(slide);
                Release(slides);
                Release(pageSetup);
                Release(copy);
                Release(presentations);
                Release(application);
                try { Directory.Delete(folder, true); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    SettingsStore.Log("清理 PDF 临时文件失败。", ex);
                }
            }
        }

        private static PdfCrop ExportMetafile(object shape, string path)
        {
            // ppShapeFormatEMF = 5. ENHMETAHEADER.rclFrame uses 0.01 mm coordinates.
            ((dynamic)shape).Export(path, 5);
            using (var reader = new BinaryReader(File.OpenRead(path)))
            {
                if (reader.BaseStream.Length < 88 || reader.ReadUInt32() != 1 || reader.ReadUInt32() < 88)
                    throw new IOException("PowerPoint 未生成有效的矢量图形单元。");
                reader.BaseStream.Position = 24;
                int left = reader.ReadInt32();
                int top = reader.ReadInt32();
                int right = reader.ReadInt32();
                int bottom = reader.ReadInt32();
                if (reader.ReadUInt32() != 0x464D4520 || right <= left || bottom <= top)
                    throw new IOException("PowerPoint 返回的图形绘制边界为空或无效。");
                double points = 72 / FrameUnitsPerInch;
                return new PdfCrop
                {
                    width_pt = (right - (double)left) * points,
                    height_pt = (bottom - (double)top) * points
                };
            }
        }

        private static double RequiredBitmapScale(object root, List<NativePicture> pictures)
        {
            // Native PDF export downsamples ordinary pictures near 200 PPI. Keep the
            // original pixels below that threshold on the disposable canvas, then let
            // the core resize the PDF page with a vector transform. No bitmap is enlarged.
            double required = 0;
            foreach (NativePicture picture in pictures)
            {
                object current = (int)((dynamic)root).Id == picture.ShapeId ? root :
                    PowerPointSvgExporter.FindPicture(root, picture.ShapeId, picture.ShapeName);
                if (current == null)
                    throw new InvalidDataException("无法测量选区中的原始图片。");
                try
                {
                    if ((int)((dynamic)current).Visible == 0)
                        continue;
                    double width = (double)((dynamic)current).Width;
                    double height = (double)((dynamic)current).Height;
                    if (width <= 0 || height <= 0)
                        throw new InvalidDataException("选区中有尺寸无效的图片。");
                    using (var input = new MemoryStream(picture.Data, false))
                    {
                        BitmapFrame frame = BitmapFrame.Create(input,
                            BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.None);
                        required = Math.Max(required, Math.Max(
                            frame.PixelWidth * 72 / (width * ImagePpiLimit),
                            frame.PixelHeight * 72 / (height * ImagePpiLimit)));
                    }
                }
                finally
                {
                    if (!ReferenceEquals(current, root))
                        Release(current);
                }
            }
            return required;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
