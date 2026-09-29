using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTFigureExporter
{
    internal sealed class SettingsDialog : Form
    {
        private readonly ExportSettings settings;
        private readonly double sourceWidth;
        private readonly double sourceHeight;
        private readonly ComboBox mode;
        private readonly ComboBox unit;
        private readonly ComboBox format;
        private readonly NumericUpDown width;
        private readonly NumericUpDown height;
        private readonly NumericUpDown percent;
        private readonly NumericUpDown dpi;
        private readonly NumericUpDown quality;
        private readonly CheckBox transparent;
        private readonly Label preview;
        private readonly Label note;
        private readonly Label qualityLabel;
        public ExportSettings Result { get; private set; }

        public SettingsDialog(ExportSettings initial, double widthPt, double heightPt)
        {
            settings = initial.Copy();
            sourceWidth = widthPt;
            sourceHeight = heightPt;
            Text = "图形导出设置";
            Font = new Font("Microsoft YaHei UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(560, 600);
            MinimumSize = new Size(560, 600);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            var table = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24, 16, 24, 16), ColumnCount = 2, RowCount = 13 };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(table);
            mode = Combo(new[] { "保持 PPT 尺寸", "按宽度等比", "按高度等比", "自定义宽高（拉伸）", "缩放百分比" });
            mode.SelectedIndex = Array.IndexOf(new[] { "keep", "width", "height", "custom", "scale" }, settings.SizeMode);
            unit = Combo(new[] { "mm", "cm", "inch" });
            unit.SelectedItem = settings.Unit;
            width = Number(settings.Width, 0.001M, 100000M, 3);
            height = Number(settings.Height, 0.001M, 100000M, 3);
            percent = Number(settings.Percent, 0.01M, 10000M, 2);
            format = Combo(new[] { "TIFF", "PNG", "SVG", "PDF", "JPEG", "WebP" });
            format.SelectedIndex = Array.IndexOf(new[] { "tiff", "png", "svg", "pdf", "jpeg", "webp" }, settings.Format);
            dpi = Number(settings.Dpi, 1, 9600, 0);
            quality = Number(settings.Quality, 1, 100, 0);
            transparent = new CheckBox { Text = "保留透明背景", AutoSize = true, Checked = settings.Transparent };
            AddRow(table, 0, "尺寸模式", mode);
            AddRow(table, 1, "物理尺寸单位", unit);
            AddRow(table, 2, "目标宽度", width);
            AddRow(table, 3, "目标高度", height);
            AddRow(table, 4, "缩放百分比 (%)", percent);
            AddRow(table, 5, "输出格式", format);
            AddRow(table, 6, "栅格分辨率 (DPI)", dpi);
            qualityLabel = AddRow(table, 7, "压缩质量 (1–100)", quality);
            AddRow(table, 8, "背景", transparent);
            preview = new Label { Dock = DockStyle.Fill, AutoSize = false, ForeColor = Color.FromArgb(30, 70, 100), Padding = new Padding(0, 8, 0, 0) };
            table.Controls.Add(preview, 0, 9);
            table.SetColumnSpan(preview, 2);
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            note = new Label { Dock = DockStyle.Fill, ForeColor = Color.DimGray, AutoSize = false };
            table.Controls.Add(note, 0, 10);
            table.SetColumnSpan(note, 2);
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var ok = new Button { Text = "保存设置", AutoSize = true, Height = 32 };
            var cancel = new Button { Text = "取消", AutoSize = true, Height = 32, DialogResult = DialogResult.Cancel };
            actions.Controls.Add(ok);
            actions.Controls.Add(cancel);
            table.Controls.Add(actions, 0, 11);
            table.SetColumnSpan(actions, 2);
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            AcceptButton = ok;
            CancelButton = cancel;
            ok.Click += (sender, args) => SaveAndClose();
            mode.SelectedIndexChanged += Changed;
            unit.SelectedIndexChanged += Changed;
            format.SelectedIndexChanged += Changed;
            foreach (var field in new[] { width, height, percent, dpi, quality })
                field.ValueChanged += Changed;
            transparent.CheckedChanged += Changed;
            UpdatePreview();
        }

        private static ComboBox Combo(string[] items)
        {
            var box = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            box.Items.AddRange(items);
            return box;
        }

        private static NumericUpDown Number(double value, decimal minimum, decimal maximum, int decimals)
        {
            return new NumericUpDown
            {
                Dock = DockStyle.Fill,
                DecimalPlaces = decimals,
                Minimum = minimum,
                Maximum = maximum,
                Value = (decimal)Math.Max((double)minimum, Math.Min((double)maximum, value)),
                ThousandsSeparator = true
            };
        }

        private static Label AddRow(TableLayoutPanel table, int row, string caption, Control control)
        {
            var label = new Label { Text = caption, AutoSize = true, Anchor = AnchorStyles.Left };
            table.Controls.Add(label, 0, row);
            control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            table.Controls.Add(control, 1, row);
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
            return label;
        }

        private void ReadFields()
        {
            settings.SizeMode = new[] { "keep", "width", "height", "custom", "scale" }[mode.SelectedIndex];
            settings.Format = new[] { "tiff", "png", "svg", "pdf", "jpeg", "webp" }[format.SelectedIndex];
            settings.Unit = (string)unit.SelectedItem;
            settings.Width = (double)width.Value;
            settings.Height = (double)height.Value;
            settings.Percent = (double)percent.Value;
            settings.Dpi = (int)dpi.Value;
            settings.Quality = (int)quality.Value;
            settings.Transparent = transparent.Checked;
        }

        private void Changed(object sender, EventArgs args)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            ReadFields();
            width.Enabled = settings.SizeMode == "width" || settings.SizeMode == "custom";
            height.Enabled = settings.SizeMode == "height" || settings.SizeMode == "custom";
            unit.Enabled = width.Enabled || height.Enabled;
            percent.Enabled = settings.SizeMode == "scale";
            dpi.Enabled = settings.IsRaster;
            quality.Visible = settings.UsesQuality;
            qualityLabel.Visible = settings.UsesQuality;
            transparent.Enabled = settings.SupportsAlpha;
            note.Text = settings.SizeMode == "custom" ? "自定义宽高会拉伸图形，改变原始宽高比例。\r\n" : "";
            note.Text += settings.IsRaster
                ? "像素数由物理尺寸与 DPI 计算，最多 8000 万像素。JPEG 始终使用白底。WebP 不保存印刷 DPI。"
                : "SVG / PDF 保留原图透明度。PDF 保留可转换的矢量内容；滤镜效果可能局部栅格化。";
            try
            {
                preview.Text = sourceWidth > 0 && sourceHeight > 0
                    ? "预计输出：" + settings.Target(sourceWidth, sourceHeight).Description(settings.IsRaster, settings.Dpi)
                    : "选中一个图形或组合后显示预计输出尺寸。";
            }
            catch (ArgumentException ex)
            {
                preview.Text = ex.Message;
            }
        }

        private void SaveAndClose()
        {
            try
            {
                ReadFields();
                settings.Validate();
                if (sourceWidth > 0 && sourceHeight > 0)
                    settings.Target(sourceWidth, sourceHeight);
                Result = settings.Copy();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    internal sealed class ExportProgress : Form
    {
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private bool completed;
        public CoreResult Result { get; private set; }
        public Exception Error { get; private set; }

        public ExportProgress(string request)
        {
            Text = "正在导出图形";
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(420, 145);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            var label = new Label { Text = "正在生成输出文件…", AutoSize = true, Location = new Point(20, 20) };
            var progress = new ProgressBar { Style = ProgressBarStyle.Marquee, Location = new Point(20, 53), Size = new Size(380, 20) };
            var cancel = new Button { Text = "取消", Location = new Point(310, 95), Size = new Size(90, 30) };
            Controls.AddRange(new Control[] { label, progress, cancel });
            cancel.Click += (sender, args) => cancellation.Cancel();
            FormClosing += (sender, args) =>
            {
                if (!completed)
                {
                    cancellation.Cancel();
                    label.Text = "正在停止…";
                    cancel.Enabled = false;
                    args.Cancel = true;
                }
            };
            Shown += async (sender, args) => await Run(request);
        }

        private async Task Run(string request)
        {
            try { Result = await ExportService.RunAsync(request, cancellation.Token); }
            catch (Exception ex) { Error = ex; }
            finally
            {
                completed = true;
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                cancellation.Dispose();
            base.Dispose(disposing);
        }
    }
}
