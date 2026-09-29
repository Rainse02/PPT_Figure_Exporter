using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace PPTFigureExporter
{
    public sealed class ExportSettings
    {
        public string SizeMode { get; set; } = "keep";
        public string Format { get; set; } = "tiff";
        public string Unit { get; set; } = "cm";
        public double Width { get; set; } = 14.35;
        public double Height { get; set; } = 12.0;
        public double Percent { get; set; } = 100;
        public int Dpi { get; set; } = 1000;
        public bool Transparent { get; set; }
        public int Quality { get; set; } = 95;
        public string OutputDirectory { get; set; } = "";

        public bool IsRaster { get { return Format != "svg" && Format != "pdf"; } }
        public bool SupportsAlpha { get { return Format == "png" || Format == "tiff" || Format == "webp"; } }
        public bool UsesQuality { get { return Format == "jpeg" || Format == "webp"; } }
        public string Extension { get { return Format == "jpeg" ? "jpg" : Format; } }

        public ExportSettings Copy()
        {
            return (ExportSettings)MemberwiseClone();
        }

        public void Validate()
        {
            if (Array.IndexOf(new[] { "keep", "width", "height", "custom", "scale" }, SizeMode) < 0)
                throw new ArgumentException("尺寸模式无效。");
            if (Array.IndexOf(new[] { "svg", "pdf", "png", "tiff", "jpeg", "webp" }, Format) < 0)
                throw new ArgumentException("输出格式无效。");
            if (Array.IndexOf(new[] { "mm", "cm", "inch" }, Unit) < 0)
                throw new ArgumentException("尺寸单位无效。");
            if (!Positive(Width) || !Positive(Height) || !Positive(Percent))
                throw new ArgumentException("尺寸与缩放比例必须为有限正数。");
            if (Dpi < 1 || Dpi > 9600 || Quality < 1 || Quality > 100)
                throw new ArgumentException("DPI 必须为 1–9600，质量必须为 1–100。");
        }

        public SizePoints Target(double widthPt, double heightPt)
        {
            Validate();
            if (!Positive(widthPt) || !Positive(heightPt))
                throw new ArgumentException("选中对象的宽高必须大于零。请将线条与其他对象组合后导出。");
            double factor = Unit == "mm" ? 72 / 25.4 : Unit == "cm" ? 72 / 2.54 : 72;
            double width = widthPt;
            double height = heightPt;
            if (SizeMode == "width")
            {
                width = Width * factor;
                height = width * heightPt / widthPt;
            }
            else if (SizeMode == "height")
            {
                height = Height * factor;
                width = height * widthPt / heightPt;
            }
            else if (SizeMode == "custom")
            {
                width = Width * factor;
                height = Height * factor;
            }
            else if (SizeMode == "scale")
            {
                width *= Percent / 100;
                height *= Percent / 100;
            }
            if (!Positive(width) || !Positive(height))
                throw new ArgumentException("目标尺寸超出可用范围。");
            var result = new SizePoints { Width = width, Height = height };
            if (IsRaster)
            {
                double pixelsX = Math.Max(1, Math.Round(width / 72 * Dpi, MidpointRounding.ToEven));
                double pixelsY = Math.Max(1, Math.Round(height / 72 * Dpi, MidpointRounding.ToEven));
                if (!Positive(pixelsX * pixelsY) || pixelsX * pixelsY > 80000000)
                    throw new ArgumentException("输出超过 8000 万像素，请减小尺寸或 DPI。");
                if (Format == "webp" && (pixelsX > 16383 || pixelsY > 16383))
                    throw new ArgumentException("WebP 每边不能超过 16383 像素。");
                result.PixelWidth = (int)pixelsX;
                result.PixelHeight = (int)pixelsY;
            }
            return result;
        }

        private static bool Positive(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value) && value > 0;
        }
    }

    public sealed class SizePoints
    {
        public double Width;
        public double Height;
        public int PixelWidth;
        public int PixelHeight;

        public string Description(bool raster, int dpi)
        {
            string size = string.Format(CultureInfo.CurrentCulture, "{0:0.###} × {1:0.###} cm", Width / 72 * 2.54, Height / 72 * 2.54);
            return raster ? size + string.Format("；{0} × {1} px，{2} DPI", PixelWidth, PixelHeight, dpi) : size;
        }
    }

    internal static class SettingsStore
    {
        internal static readonly string Folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PPTFigureExporter");
        private static readonly string SettingsPath = Path.Combine(Folder, "settings.json");

        public static ExportSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                    return new ExportSettings();
                var settings = new JavaScriptSerializer().Deserialize<ExportSettings>(File.ReadAllText(SettingsPath, Encoding.UTF8));
                if (settings == null)
                    return new ExportSettings();
                settings.Validate();
                return settings;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is InvalidOperationException)
            {
                Log("读取设置失败，使用默认值。", ex);
                return new ExportSettings();
            }
        }

        public static void Save(ExportSettings settings)
        {
            settings.Validate();
            Directory.CreateDirectory(Folder);
            string temporary = Path.Combine(Folder, Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllText(temporary, new JavaScriptSerializer().Serialize(settings), new UTF8Encoding(false));
                if (File.Exists(SettingsPath))
                    File.Replace(temporary, SettingsPath, null);
                else
                    File.Move(temporary, SettingsPath);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        public static void Log(string message, Exception exception)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                string path = Path.Combine(Folder, "export.log");
                if (File.Exists(path) && new FileInfo(path).Length > 1048576)
                    File.Delete(path);
                File.AppendAllText(path, DateTime.Now.ToString("s", CultureInfo.InvariantCulture) + " " + message + " " + exception + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Logging must never interrupt PowerPoint shutdown or mask the export error.
            }
        }
    }
}
