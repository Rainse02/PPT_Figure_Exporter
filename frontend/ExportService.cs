using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace PPTFigureExporter
{
    internal sealed class SelectedFigure : IDisposable
    {
        private object shape;
        private object presentation;
        private readonly int slideId;
        public double WidthPt { get; private set; }
        public double HeightPt { get; private set; }
        public string Name { get; private set; }

        private SelectedFigure(object value, object sourcePresentation, int sourceSlideId)
        {
            shape = value;
            presentation = sourcePresentation;
            slideId = sourceSlideId;
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
            object presentation = null;
            object view = null;
            object slide = null;
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
                presentation = ((dynamic)window).Presentation;
                view = ((dynamic)window).View;
                slide = ((dynamic)view).Slide;
                var figure = new SelectedFigure(selected, presentation, (int)((dynamic)slide).SlideID);
                selected = null;
                presentation = null;
                return figure;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法读取选中图形。请在普通编辑视图中选中一个图形或组合。", ex);
            }
            finally
            {
                Release(slide);
                Release(view);
                Release(presentation);
                Release(selected);
                Release(range);
                Release(selection);
                Release(window);
            }
        }

        public void ExportSvg(string path)
        {
            PowerPointSvgExporter.Export(shape, presentation, slideId, path);
        }

        public void ExportRaster(string path, SizePoints target, bool stretch)
        {
            PowerPointRasterExporter.Export(shape, presentation, path, target, stretch);
        }

        public void Dispose()
        {
            Release(shape);
            shape = null;
            Release(presentation);
            presentation = null;
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
                input_path = input,
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
