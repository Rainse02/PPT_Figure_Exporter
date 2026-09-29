using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PPTFigureExporter
{
    [ComVisible(true), Guid("A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76")]
    [ProgId("PPTFigureExporter.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCallbacks))]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IRibbonCallbacks
    {
        private object application;
        private IRibbonUI ribbon;
        private ExportSettings settings = SettingsStore.Load();
        private System.Windows.Forms.Timer selectionTimer;
        private bool busy;
        private string selectionLabel = "请选择一个图形或组合";
        private bool selectionAvailable;

        public void OnConnection(object app, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            application = app;
            SettingsStore.Log("PowerPoint add-in connected.", null);
            selectionTimer = new System.Windows.Forms.Timer { Interval = 800 };
            selectionTimer.Tick += (sender, args) => RefreshSelection();
            selectionTimer.Start();
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (selectionTimer != null)
            {
                selectionTimer.Stop();
                selectionTimer.Dispose();
                selectionTimer = null;
            }
            ribbon = null;
            application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { RefreshSelection(); }
        public void OnBeginShutdown(ref Array custom) { OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom); }

        public string GetCustomUI(string ribbonId)
        {
            SettingsStore.Log("Ribbon XML requested: " + ribbonId, null);
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("PPTFigureExporter.Ribbon.xml"))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
                return reader.ReadToEnd();
        }

        public void OnRibbonLoad(object value)
        {
            ribbon = (IRibbonUI)value;
            SettingsStore.Log("Ribbon loaded.", null);
            RefreshSelection();
        }

        public string GetSelectionLabel(object control) { return selectionLabel; }
        public string GetFormat(object control) { return settings.Format; }
        public string GetDpi(object control) { return settings.Dpi.ToString(CultureInfo.InvariantCulture); }
        public bool GetRasterEnabled(object control) { return settings.IsRaster && !busy; }
        public bool GetTransparencyEnabled(object control) { return settings.SupportsAlpha && !busy; }
        public bool GetTransparent(object control) { return settings.Transparent; }
        public string GetMode(object control) { return settings.SizeMode; }
        public bool GetExportEnabled(object control) { return application != null && selectionAvailable && !busy; }

        public void OnFormat(object control, string selectedId, int selectedIndex)
        {
            UpdateSettings(() => settings.Format = selectedId);
        }

        public void OnDpi(object control, string value)
        {
            int dpi;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out dpi) || dpi < 1 || dpi > 9600)
            {
                MessageBox.Show(Owner(), "DPI 必须是 1–9600 的整数。", "图形导出", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Invalidate();
                return;
            }
            UpdateSettings(() => settings.Dpi = dpi);
        }

        public void OnTransparent(object control, bool value)
        {
            UpdateSettings(() => settings.Transparent = value);
        }

        public void OnMode(object control, string selectedId, int selectedIndex)
        {
            UpdateSettings(() => settings.SizeMode = selectedId);
            if (selectedId != "keep")
                OnSettings(control);
        }

        public void OnSettings(object control)
        {
            if (busy)
                return;
            double width = 0;
            double height = 0;
            try
            {
                using (var figure = SelectedFigure.Read(application))
                {
                    width = figure.WidthPt;
                    height = figure.HeightPt;
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is COMException)
            {
                // Settings remain available when no figure is selected.
            }
            try
            {
                using (var dialog = new SettingsDialog(settings, width, height))
                {
                    if (dialog.ShowDialog(Owner()) == DialogResult.OK)
                    {
                        settings = dialog.Result;
                        SaveSettings();
                        Invalidate();
                    }
                }
            }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnExport(object control)
        {
            if (busy || application == null)
                return;
            string temporary = null;
            busy = true;
            Invalidate();
            try
            {
                using (var figure = SelectedFigure.Read(application))
                {
                    SizePoints target = settings.Target(figure.WidthPt, figure.HeightPt);
                    if (!File.Exists(ExportService.CorePath))
                        throw new FileNotFoundException("未找到导出核心，请重新安装插件。");
                    using (var save = new SaveFileDialog())
                    {
                        save.Title = "导出图形 — " + target.Description(settings.IsRaster, settings.Dpi);
                        save.Filter = settings.Format.ToUpperInvariant() + " 文件|*." + settings.Extension;
                        save.DefaultExt = settings.Extension;
                        save.AddExtension = true;
                        save.OverwritePrompt = true;
                        save.FileName = SafeName(figure.Name) + "." + settings.Extension;
                        save.InitialDirectory = Directory.Exists(settings.OutputDirectory)
                            ? settings.OutputDirectory
                            : Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                        if (save.ShowDialog(Owner()) != DialogResult.OK)
                            return;
                        if (!string.Equals(Path.GetExtension(save.FileName), "." + settings.Extension, StringComparison.OrdinalIgnoreCase))
                            throw new ArgumentException("文件扩展名必须为 ." + settings.Extension + "，请在格式菜单中选择目标格式。");
                        bool overwrite = File.Exists(save.FileName);
                        temporary = Path.Combine(Path.GetTempPath(), "PPTFigureExporter", Guid.NewGuid().ToString("N"));
                        Directory.CreateDirectory(temporary);
                        string svg = Path.Combine(temporary, "figure.svg");
                        figure.ExportSvg(svg);
                        string request = ExportService.WriteRequest(temporary, svg, save.FileName, figure, settings.Copy(), overwrite);
                        using (var progress = new ExportProgress(request))
                        {
                            progress.ShowDialog(Owner());
                            if (progress.Error is OperationCanceledException)
                                return;
                            if (progress.Error != null)
                                throw progress.Error;
                            settings.OutputDirectory = Path.GetDirectoryName(progress.Result.output_path);
                            SaveSettings();
                            MessageBox.Show(Owner(), "已导出：\r\n" + progress.Result.output_path + "\r\n\r\n" + target.Description(settings.IsRaster, settings.Dpi),
                                "图形导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex) { ShowError(ex); }
            finally
            {
                if (temporary != null)
                {
                    try
                    {
                        // Only this invocation's freshly created private directory is removed.
                        Directory.Delete(temporary, true);
                    }
                    catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    {
                        SettingsStore.Log("临时文件清理失败。", ex);
                    }
                }
                busy = false;
                RefreshSelection();
                Invalidate();
            }
        }

        private void RefreshSelection()
        {
            if (application == null || busy)
                return;
            string label;
            bool available;
            try
            {
                using (var figure = SelectedFigure.Read(application))
                {
                    label = string.Format(CultureInfo.CurrentCulture, "选中尺寸：{0:0.00} × {1:0.00} cm", figure.WidthPt / 72 * 2.54, figure.HeightPt / 72 * 2.54);
                    available = true;
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is COMException || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                label = "请选择一个图形或组合";
                // Keep Export available for a useful explanation of multiple selections.
                available = true;
            }
            if (label != selectionLabel || available != selectionAvailable)
            {
                selectionLabel = label;
                selectionAvailable = available;
                Invalidate();
            }
        }

        private void UpdateSettings(Action change)
        {
            if (busy)
                return;
            ExportSettings before = settings.Copy();
            try
            {
                change();
                settings.Validate();
                SaveSettings();
            }
            catch (Exception ex)
            {
                settings = before;
                ShowError(ex);
            }
            Invalidate();
        }

        private void SaveSettings()
        {
            try { SettingsStore.Save(settings); }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                SettingsStore.Log("保存设置失败。", ex);
                MessageBox.Show(Owner(), "文件导出功能可继续使用，但无法保存设置：" + ex.Message,
                    "设置保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Invalidate()
        {
            try { if (ribbon != null) ribbon.Invalidate(); }
            catch (COMException) { }
        }

        private void ShowError(Exception exception)
        {
            SettingsStore.Log("导出操作失败。", exception);
            MessageBox.Show(Owner(), exception.Message, "图形导出失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private IWin32Window Owner()
        {
            // Ribbon callbacks run inside PowerPoint's UI thread. Its native active
            // window is reliable even when Application.HWND rejects IDispatch access.
            IntPtr handle = GetActiveWindow();
            if (handle == IntPtr.Zero)
            {
                using (var process = System.Diagnostics.Process.GetCurrentProcess())
                    handle = process.MainWindowHandle;
            }
            return handle == IntPtr.Zero ? null : new OfficeWindow(handle);
        }

        private static string SafeName(string name)
        {
            string result = string.IsNullOrWhiteSpace(name) ? "Figure" : name;
            foreach (char value in Path.GetInvalidFileNameChars())
                result = result.Replace(value, '_');
            result = result.Trim().TrimEnd('.');
            if (string.IsNullOrEmpty(result))
                result = "Figure";
            string stem = result.Split('.')[0].ToUpperInvariant();
            if (System.Text.RegularExpressions.Regex.IsMatch(stem, @"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$"))
                result = "Figure_" + result;
            return result.Length > 100 ? result.Substring(0, 100) : result;
        }

        private sealed class OfficeWindow : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public OfficeWindow(IntPtr handle) { Handle = handle; }
        }
    }
}
