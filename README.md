# PPT Figure Exporter

将 PowerPoint 中选中的图形或组合导出为 SVG、PDF、TIFF、PNG、JPEG 和 WebP，适合制作论文插图、报告配图和透明背景素材。

可以保留图形在 PowerPoint 中的物理尺寸，也可以指定宽高、缩放比例和 DPI。导出在本机完成，原幻灯片保持不变。

## 安装

从 [Releases](https://github.com/Rainse02/PPT_Figure_Exporter/releases) 下载 `PPTFigureExporter_Setup.exe`，关闭 PowerPoint 后运行安装器。重新打开 PowerPoint，功能区中会出现 **Figure Export** 选项卡。

运行环境：

- Windows 10/11 x64。
- 64 位桌面版 PowerPoint，支持 SVG 自动化导出。
- .NET Framework 4.8 或更新的 4.x 版本。

PowerPoint 的 SVG 自动化接口从 Windows 版 Version 2302 开始提供，见[微软接口文档](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)。旧版 Office 遇到 SVG 导出错误时，需要更新 PowerPoint。

安装包与 .NET 运行库分开提供。缺少运行库时，从[微软下载页](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)选择 **Runtime** 安装。离线安装可单独准备微软离线运行库安装器。

插件按当前用户安装。升级或卸载前关闭 PowerPoint，卸载入口位于 Windows 的“已安装的应用”。

## 使用

1. 选中要导出的图形。多个对象先按 `Ctrl+G` 组合。
2. 打开 **Figure Export**，选择格式和尺寸模式。
3. 在“尺寸与质量设置”中填写宽高、单位或缩放比例；栅格格式可设置 DPI 和透明背景。
4. 点击“导出图形”，选择保存位置。

默认输出为 TIFF、1000 DPI、白底，尺寸与 PowerPoint 中的对象一致。JPEG 和 WebP 的默认质量为 95。

### 尺寸

支持 mm、cm、inch 三种单位，以及以下尺寸模式：

- 保持 PPT 尺寸。
- 指定宽度或高度，另一边等比缩放。
- 自定义宽高，按指定比例拉伸。
- 按百分比缩放。

栅格图的像素数由物理尺寸和 DPI 共同决定。例如，10.16 × 6.35 cm 的图形以 300 DPI 导出，得到 1200 × 750 px。

### 格式

| 格式 | 用途与设置 |
| --- | --- |
| SVG | 矢量图，适合继续编辑和网页使用 |
| PDF | 矢量排版，页面大小与目标图形尺寸一致 |
| TIFF | LZW 无损压缩，写入 DPI，支持透明背景 |
| PNG | 无损压缩，写入分辨率信息，支持透明背景 |
| JPEG | 白色背景，质量可设为 1–100 |
| WebP | 支持透明背景，质量可设为 1–100 |

SVG 和 PDF 按物理尺寸输出；TIFF、PNG、JPEG 和 WebP 按尺寸与 DPI 渲染。嵌入图片保留其原始图像内容。

## 从源码构建

准备以下工具：

- PowerShell 5.1 或更新版本。
- Visual Studio 2022 Build Tools，包含 MSBuild、MSVC x64 和 Windows SDK。
- Rust 的 `x86_64-pc-windows-msvc` 工具链。
- PowerPoint 及微软 Office 互操作程序集 `office.dll`、`Extensibility.dll`。

在项目根目录执行：

```powershell
.\scripts\check-environment.ps1 -ForBuild -Strict
cargo fetch --manifest-path .\native\Cargo.toml --locked --target x86_64-pc-windows-msvc
.\scripts\generate-notices.ps1
.\scripts\build.ps1 -BootstrapInno
```

构建脚本准备 .NET Framework 4.8 引用程序集，编译 C# 前端和 Rust 核心，再用 Inno Setup 打包。`-BootstrapInno` 会下载并使用项目内的便携编译器。首次构建需要联网下载依赖。

产物位于 `dist/`：

| 路径 | 内容 |
| --- | --- |
| `PPTFigureExporter_Setup.exe` | 可分发的安装包 |
| `app/` | 插件 DLL、转换核心和构建清单 |
| `licenses/` | Apache 2.0 许可证与第三方声明 |

已有 `dist/app/` 时，可以只重新打包：

```powershell
.\scripts\build.ps1 -RepackageOnly -BootstrapInno
```

发布时将 Setup 上传到 GitHub Releases。需要随包提供独立的 .NET 离线安装器时，通过 `-DotNetOfflineInstaller` 指定文件路径，脚本会将其复制到 `dist/prereq/`。

### 实现

C# COM 加载项负责 PowerPoint 功能区、设置和选区读取，调用 PowerPoint 的 `Shape.Export` 接口生成 SVG。Rust 核心接收转换参数，使用 usvg/resvg 渲染图形、svg2pdf 生成 PDF，并编码其他图像格式。

`frontend/` 是加载项，`native/` 是转换核心，`installer/` 是安装器定义，`scripts/` 提供构建工具，`tests/` 包含前端和 Office 联调测试。`Directory.Build.props` 配置 .NET 引用程序集路径，`native/Cargo.lock` 固定 Rust 依赖版本。

## 设置与临时文件

设置保存在 `%APPDATA%\PPTFigureExporter\settings.json`，错误日志位于同目录的 `export.log`。插件会记住导出参数和最后使用的输出目录。

每次导出的中间 SVG 和请求文件放在 `%TEMP%\PPTFigureExporter\` 的独立子目录中，任务结束后清理。崩溃或强制结束进程可能留下文件；取消转换时，输出目录也可能留下带 `.figure-export-` 标记的临时文件。当前版本没有自动清扫历史残留的功能。手动清理前关闭 PowerPoint，并确认 `figure-export-core.exe` 已退出，只删除这些已确认的临时文件。

## 许可证

项目采用 [Apache License 2.0](LICENSE)。第三方组件保留各自许可证，相关声明随安装包放在 `licenses/THIRD_PARTY_NOTICES.txt` 中。
