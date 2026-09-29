using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml;

namespace PPTFigureExporter
{
    internal sealed class SelectedFigure : IDisposable
    {
        private object shape;
        public double WidthPt { get; private set; }
        public double HeightPt { get; private set; }
        public string Name { get; private set; }

        private SelectedFigure(object value)
        {
            shape = value;
            dynamic selected = value;
            WidthPt = (double)selected.Width;
            HeightPt = (double)selected.Height;
            Name = (string)selected.Name;
        }

        public static SelectedFigure Read(object application)
        {
            object window = null;
            object selection = null;
            object range = null;
            object selected = null;
            try
            {
                window = ((dynamic)application).ActiveWindow;
                if (window == null)
                    throw new InvalidOperationException("请打开演示文稿，并选中一个图形或组合。");
                selection = ((dynamic)window).Selection;
                if ((int)((dynamic)selection).Type != 2)
                    throw new InvalidOperationException("请选中一个图形或组合；若正在编辑文字，请先按 Esc。");
                range = (bool)((dynamic)selection).HasChildShapeRange
                    ? ((dynamic)selection).ChildShapeRange
                    : ((dynamic)selection).ShapeRange;
                if ((int)((dynamic)range).Count != 1)
                    throw new InvalidOperationException("选中了多个独立对象。请先按 Ctrl+G 组合，再导出。");
                selected = ((dynamic)range)[1];
                var figure = new SelectedFigure(selected);
                selected = null;
                return figure;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法读取选中图形。请在普通编辑视图中选中一个图形或组合。", ex);
            }
            finally
            {
                Release(selected);
                Release(range);
                Release(selection);
                Release(window);
            }
        }

        public void ExportSvg(string path)
        {
            // Only reads/exports the selected shape. No slide, shape or clipboard mutations.
            try
            {
                ((dynamic)shape).Export(path, 6);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("PowerPoint 无法通过自动化接口导出 SVG。请使用支持该接口的新版 PowerPoint，并检查所选图形内容。", ex);
            }
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
                throw new IOException("PowerPoint 未生成有效的 SVG 文件。");
            // Older Office builds silently write JPEG for the unsupported SVG enum value.
            // Inspect content: the file extension alone does not establish SVG support.
            try
            {
                var options = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var reader = XmlReader.Create(path, options))
                {
                    reader.MoveToContent();
                    if (reader.LocalName != "svg" || reader.NamespaceURI != "http://www.w3.org/2000/svg")
                        throw new XmlException("根元素不是 SVG。");
                }
            }
            catch (XmlException ex)
            {
                throw new InvalidOperationException("此 PowerPoint 自动化接口没有生成真实 SVG。手动另存为 SVG 与自动化导出的版本支持可能不同，请更新 Office。", ex);
            }
        }

        public void Dispose()
        {
            Release(shape);
            shape = null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }

    public sealed class CoreResult
    {
        public bool ok { get; set; }
        public string output_path { get; set; }
        public double width_pt { get; set; }
        public double height_pt { get; set; }
        public int? pixel_width { get; set; }
        public int? pixel_height { get; set; }
        public string error { get; set; }
    }

    internal static class ExportService
    {
        public static string CorePath
        {
            get { return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "figure-export-core.exe"); }
        }

        public static string WriteRequest(string folder, string input, string output, SelectedFigure figure, ExportSettings settings, bool overwrite)
        {
            var request = new
            {
                input_svg = input,
                output_path = output,
                format = settings.Format,
                source_width_pt = figure.WidthPt,
                source_height_pt = figure.HeightPt,
                size_mode = settings.SizeMode,
                width = settings.Width,
                height = settings.Height,
                unit = settings.Unit,
                percent = settings.Percent,
                dpi = settings.Dpi,
                transparent = settings.Transparent,
                quality = settings.Quality,
                overwrite = overwrite
            };
            string path = Path.Combine(folder, "request.json");
            File.WriteAllText(path, new JavaScriptSerializer().Serialize(request), new UTF8Encoding(false));
            return path;
        }

        public static async Task<CoreResult> RunAsync(string requestPath, CancellationToken cancellation)
        {
            if (!File.Exists(CorePath))
                throw new FileNotFoundException("未找到导出核心，请重新安装 PPT Figure Exporter。", CorePath);
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = CorePath,
                    Arguments = "--request \"" + requestPath + "\"",
                    WorkingDirectory = Path.GetDirectoryName(CorePath),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                process.Start();
                Task<string> stdout = process.StandardOutput.ReadToEndAsync();
                Task<string> stderr = process.StandardError.ReadToEndAsync();
                Task finished = Task.Run(() => process.WaitForExit());
                Task timeout = Task.Delay(TimeSpan.FromMinutes(5), cancellation);
                if (await Task.WhenAny(finished, timeout) != finished && !process.HasExited)
                {
                    try { process.Kill(); }
                    catch (InvalidOperationException) { }
                    await finished;
                    await Task.WhenAll(stdout, stderr);
                    cancellation.ThrowIfCancellationRequested();
                    throw new TimeoutException("导出超过 5 分钟，已停止。请减小尺寸或简化复杂滤镜。");
                }
                await finished;
                string output = await stdout;
                string errors = await stderr;
                CoreResult result;
                try
                {
                    result = new JavaScriptSerializer().Deserialize<CoreResult>(output);
                }
                catch (ArgumentException ex)
                {
                    throw new IOException("导出核心未返回有效结果。" + errors, ex);
                }
                if (process.ExitCode != 0 || result == null || !result.ok)
                    throw new IOException(result != null && !string.IsNullOrWhiteSpace(result.error) ? result.error : "导出失败。" + errors);
                if (string.IsNullOrWhiteSpace(result.output_path) || !File.Exists(result.output_path))
                    throw new IOException("导出核心报告成功，但找不到输出文件。");
                return result;
            }
        }
    }
}
