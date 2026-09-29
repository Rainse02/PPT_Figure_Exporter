# PPT Figure Exporter

PPT Figure Exporter 是面向 Windows 桌面版 PowerPoint 的图形导出插件。选中一个图形或组合后，可以按对象的实际物理尺寸，导出 SVG、PDF、TIFF、PNG、JPEG 或 WebP。

项目采用 C#/.NET Framework COM 加载项与 Rust 原生转换核心。导出在本机完成，无需上传演示文稿。项目自有代码采用 **Apache License 2.0**；第三方依赖保留各自许可证。

## 适用场景

| 场景 | 用法与边界 |
| --- | --- |
| 论文与报告插图 | 将 PowerPoint 中排版好的文字、线条、箭头、嵌入图像等组合，按目标版面宽度输出 |
| 期刊要求的栅格图 | 指定物理尺寸和 DPI，导出带分辨率元数据的 TIFF、PNG 或 JPEG；具体参数按投稿要求设置 |
| 矢量交付与后续排版 | 导出 SVG 或 PDF，保留可转换的矢量路径和文本；嵌入位图仍是位图 |
| 透明背景素材 | 使用 PNG、TIFF 或 WebP 并启用透明背景 |
| 网页与普通演示素材 | 使用 PNG、JPEG 或 WebP，按需要设置尺寸、透明度与有损质量 |

当前版本一次导出一个选中对象或组合，不提供整套幻灯片批量导出、自动命名模板或科研原始数据处理。导出不会改变源对象的位置、尺寸、组合关系、文稿保存状态或剪贴板。

## 运行环境

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows 10/11 x64；本机验证环境为 Windows 10 Pro 22H2 |
| PowerPoint | Windows 桌面版、**64 位**，且 `Shape.Export(..., 6)` 能输出真实 SVG |
| .NET | .NET Framework **4.8 或更新的 4.x 版本**；.NET 6/8 等不能替代此项 |
| 字体 | 安装图形实际使用的字体；字体缺失可能导致替换或布局变化 |
| 安装权限 | 插件按当前用户安装和注册，正常安装不要求管理员权限；系统运行库安装按微软安装器要求执行 |
| 磁盘与内存 | 输出目录和临时目录可写；较高 DPI 会显著增加像素数、内存与输出文件体积 |

运行插件无需 Python、Conda、Rust、Node、完整 Visual Studio、VSTO Runtime、Office PIA 单独部署，或 ImageMagick、Inkscape、Ghostscript 等转换程序。32 位 Office、macOS PowerPoint、网页版 PowerPoint 和 WPS 不在当前实现范围内。

### Office 版本与兼容性

所有六种输出格式均先调用 PowerPoint 导出 SVG，再由转换核心处理。因此，自动化 SVG 接口是所有格式的共同前提。

| Office / PowerPoint | 已知状态 |
| --- | --- |
| **Office 2024 x64，PowerPoint 16.0.17932.20130** | 本项目已实测：原生 SVG 导出、六格式转换、COM 加载与功能区操作通过 |
| Windows PowerPoint Version 2302 及后续版本 | 微软文档标明该版本范围提供 SVG 过滤器；不等同于本项目已逐个版本验收 |
| 旧构建 **16.0.14334.20918** | 本机曾观察到 `Shape.Export(..., 6)` 生成 JPEG 内容，无法作为当前转换链路的有效输入 |
| 其他 Microsoft 365、Office 永久版或 LTSC 构建 | 需以实际自动化导出结果为准；仅凭“2019 / 2021 / 2024”等产品名称不能确认兼容性 |

接口依据：[Microsoft Learn — Shape.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)。版本 2302 指 PowerPoint 的应用版本，不是 Windows 版本。

手动“另存为图片 → SVG”可用，并不等于同一构建支持此自动化调用。插件检查导出文件的 XML 根元素，拒绝将 JPEG 等内容误作 SVG；不会自动回退到截图或 XPS。安装器检查 PowerPoint 位数与注册情况，实际 SVG 能力在导出时验证。

上述本机验证记录日期为 2026-09-29。尚未完成独立干净电脑及所有 Office 发布通道的兼容性验收。

## 安装与使用

1. 关闭 PowerPoint，运行 `dist/PPTFigureExporter_Setup.exe`。
2. 重新打开 PowerPoint，在功能区找到 **Figure Export**。
3. 选中一个图形或组合。多个独立对象须先按 `Ctrl+G` 组合。
4. 选择输出格式和尺寸模式；需要输入宽高、单位、缩放比例或有损质量时，打开“尺寸与质量设置”。
5. 栅格格式可设置 DPI；PNG、TIFF、WebP 可选择保留透明背景。
6. 点击“导出图形”，在保存对话框中选择目标位置；已有文件需要确认覆盖。

默认参数为：**保持 PPT 尺寸、TIFF、1000 DPI、白底**。JPEG / WebP 的默认质量为 95。默认值不代表所有期刊或业务场景都需要 1000 DPI。

### 安装位置与数据

| 内容 | 默认位置 / 行为 |
| --- | --- |
| 插件与原生核心 | `%LOCALAPPDATA%\Programs\PPTFigureExporter\` |
| 用户设置与最后输出目录 | `%APPDATA%\PPTFigureExporter\settings.json` |
| 错误与加载日志 | `%APPDATA%\PPTFigureExporter\export.log` |
| 转换临时文件 | `%TEMP%\PPTFigureExporter\` 下的独立子目录，任务结束后尝试清理，异常残留见下文 |
| COM 加载项注册 | 当前用户 `HKCU` 下的插件注册项 |
| 安装后的许可文本 | 安装目录的 `licenses\LICENSE` 与 `licenses\THIRD_PARTY_NOTICES.txt` |

升级和卸载前也需要关闭 PowerPoint，安装器不会强制结束 PowerPoint。通过 Windows“应用和功能 / 已安装的应用”卸载插件；用户设置和日志保留在 AppData，便于后续恢复或排查。

### .NET 独立依赖与分发

普通用户只需接收 **`PPTFigureExporter_Setup.exe`**。主包约 3.8 MB，包含插件、原生核心、说明和许可，不包含 .NET 离线运行库，也不会自动下载或启动它。

- 已有 .NET Framework 4.8 或更新的 4.x 版本：直接安装插件。
- 缺少运行库：从[微软 .NET Framework 4.8 下载页](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)选择 **Runtime**，按提示安装、必要时重启，再运行插件 Setup。终端用户无需 Developer Pack。
- 离线电脑：单独传输微软离线 Runtime 安装器与插件 Setup，先安装运行库。

`dist/app/` 是发布文件暂存目录，用于重新打包；复制该目录不会自动完成 COM 注册。普通分发不需要将整个项目或 `dist` 目录一并交付。

## 临时文件、缓存与异常残留

**正常使用不会按导出次数持续保存中间 SVG；异常中断时仍可能留下临时文件。** 当前实现为“每次任务单独建目录，结束后清理”，没有跨任务的 SVG 缓存，也没有启动时扫描并删除历史残留的机制。

| 文件类型 | 位置与生命周期 |
| --- | --- |
| 中间 SVG 与请求 JSON | `%TEMP%\PPTFigureExporter\<随机标识>\figure.svg`、`request.json`；仅本次转换使用 |
| 尚未提交的输出文件 | 用户选择的输出目录中，命名如 `.figure.tiff.figure-export-1234.0tmp`；编码数据写入完成后改名为正式输出 |
| 用户设置 | `%APPDATA%\PPTFigureExporter\settings.json`；有意长期保留，不属于缓存垃圾 |
| 错误日志 | 同目录 `export.log`；写入新记录前若已有文件超过 1 MiB，就删除旧日志再写入，因此并非严格的 1 MiB 硬上限 |
| 开发缓存 | 项目 `.build/`、`native/target/`、`bin/obj/`、`tests/output/`；仅开发构建或测试生成，普通安装不会带入 |

### 哪些情况会清理

- 在保存对话框直接取消：尚未创建转换临时目录。
- 导出成功：完成提示关闭后，前端的 `finally` 清理本次独立目录；最终导出的文件保留。
- SVG 导出失败、转换报错或用户取消：前端只要正常返回异常处理流程，就仍会尝试清理本次目录。
- 原生核心正常报告写入或提交失败：尝试删除目标目录中的临时输出文件。
- 用户取消或超过 5 分钟：前端终止并等待核心进程退出，再执行前端清理；若恰好在写临时输出时被终止，目标目录仍可能留下该临时文件。

### 哪些情况可能残留

PowerPoint 崩溃、任务管理器强制结束进程、系统断电，以及文件占用或权限错误，都可能使清理无法完成。强制结束 PowerPoint 时，已经启动的转换核心也可能短暂继续运行。前端清理失败会尝试写入 `export.log`，但程序崩溃时不保证能留下日志。

残留主要有两处：`%TEMP%\PPTFigureExporter\` 内的旧任务子目录，以及输出目录内带 `.figure-export-` 标记、以 `tmp` 结尾的临时文件。设置写入同样采用临时文件替换；极端中断时 AppData 设置目录也可能留下随机名称的 `.tmp` 文件。空的 `%TEMP%\PPTFigureExporter\` 根目录可以保留，几乎不占空间。

当前版本不自动清扫这些历史残留。若需要手动处理，先关闭 PowerPoint，并确认 `figure-export-core.exe` 已结束，再仅删除已确认属于本插件、且不再使用的旧任务目录或临时文件。不要清空整个系统 Temp、整个输出目录或整个 AppData；`settings.json`、需要保留的日志与正式导出文件不属于应清除的中间缓存。中间 SVG 可能包含图形、文字和嵌入图像，异常残留的内容与原图有关。

这里描述的是插件界面的临时文件管理。直接调用核心命令行时，调用方提供的输入 SVG 和请求 JSON 由调用方管理，核心不会删除它们。本节结论来自当前清理代码检查，未通过强杀 PowerPoint 或断电进行故障注入测试。

## 尺寸、DPI 与输出格式

目标物理尺寸来自 PowerPoint 对象的 `Width` 和 `Height`，单位为 pt（点）：**72 pt = 1 inch = 2.54 cm**。SVG 自带的像素宽高或 `viewBox` 不用于推断最终物理尺寸。

| 尺寸模式 | 行为 |
| --- | --- |
| 保持 PPT 尺寸 | 使用选中对象的当前宽高 |
| 按宽度等比 / 按高度等比 | 指定一边，另一边按原始比例计算 |
| 自定义宽高 | 同时指定两边，允许拉伸并改变比例 |
| 缩放百分比 | 宽高均按相同比例缩放 |

尺寸输入支持 mm、cm、inch。栅格像素数按下式计算，中点采用取偶数舍入，最少为 1 像素：

```text
像素宽度 = round(目标宽度(pt) / 72 × DPI)
像素高度 = round(目标高度(pt) / 72 × DPI)
```

例如，288 × 180 pt 的对象为 10.16 × 6.35 cm；保持尺寸并选择 300 DPI 时，输出为 1200 × 750 px。

| 格式 | 尺寸与分辨率 | 透明度与编码 |
| --- | --- | --- |
| SVG | 根 width/height 写入目标 pt；不设置输出 DPI | 保留内部矢量几何和原始透明度 |
| PDF | 页面尺寸使用目标 pt；不设置输出 DPI | 可转换的路径和文本保持矢量；滤镜可能局部栅格化 |
| TIFF | 写入像素尺寸、X/YResolution 和 inch 单位 | 8 位 RGB/RGBA，LZW 无损压缩 |
| PNG | 写入像素尺寸与 pHYs 分辨率 | 8 位 RGB/RGBA，无损；DPI 存在整数 pixels/metre 的量化误差 |
| JPEG | 写入像素尺寸与 JFIF DPI | 白底，有损质量 1–100 |
| WebP | 写入像素尺寸，不承诺印刷 DPI 元数据 | 有损质量 1–100，可保留透明度 |

DPI 范围为 1–9600，单幅栅格输出最多 **80,000,000 像素**；WebP 每边最多 16383 像素，JPEG 每边最多 65535 像素。SVG/PDF 不受栅格输出 DPI 设置控制。

SVG 已有的 `viewBox` 会保留；缺失时根据原始 SVG 视口补充，之后再写入目标物理尺寸。自定义宽高模式允许内容随目标比例拉伸。

嵌入的低分辨率位图不会因提高 DPI 而产生更多细节。复杂 Office 特效、字体、OLE 对象与 SVG 滤镜需要检查实际输出。8 位展示图不用于替代高位深显微图像、测量矩阵等科学原始数据。

## 实现与构建逻辑

### 导出流程

1. **C# 前端**在 PowerPoint 的 UI/COM 线程读取选区，要求恰好一个 Shape 或 GroupShape，记录对象尺寸。
2. 调用 `Shape.Export(临时路径, 6)` 取得 SVG，并验证 XML 内容，不复制幻灯片或创建临时演示文稿。
3. 将格式、尺寸、DPI、透明度和质量写入 UTF-8 JSON 请求，启动随包的 `figure-export-core.exe`。
4. **Rust 核心**规范化 SVG；使用 usvg/resvg 解析与栅格渲染、svg2pdf 生成 PDF，以及内置编码库生成栅格文件。
5. 核心先在目标目录写临时输出，再提交为目标文件。转换失败不会将部分结果写入已有目标文件。前端收集结果、保存设置并清理临时目录。

转换过程支持取消，核心进程超时为 5 分钟。请求结构、返回 JSON 与命令行调用示例见源码中的 `docs/DEVELOPMENT.md`。

### 开发环境

| 组件 | 用途与本机验证版本 |
| --- | --- |
| PowerShell 5.1 或更新版本 | 执行环境检查、构建与打包脚本 |
| Visual Studio 2022 Build Tools | C# MSBuild，以及“使用 C++ 的桌面开发”对应的 MSVC x64 与 Windows SDK；无需完整 Visual Studio IDE |
| MSBuild / MSVC | 本机验证：MSBuild 17.14.51.32402、MSVC 14.44.35207 |
| Rust | 使用 `x86_64-pc-windows-msvc` 工具链；本机验证 rustc 1.97.1，依赖版本由 `native/Cargo.lock` 固定 |
| .NET Framework 4.8 引用程序集 | 构建脚本从微软 NuGet 包 `Microsoft.NETFramework.ReferenceAssemblies.net48` 1.0.3 准备到 `.build/tools/` |
| Office 互操作程序集 | 微软官方 `office.dll` 与 `Extensibility.dll`，供编译使用；通过 `EmbedInteropTypes=true` 嵌入必要类型 |
| Inno Setup | 仅打包需要；`-BootstrapInno` 准备项目内的 6.7.3 便携编译器 |

上述版本是已验证的构建组合，不作为所有工具的最低版本承诺。现有脚本优先探测本机 VS 2022 Build Tools、Windows SDK 和 GAC 路径，非标准安装需检查脚本和项目引用配置。前端支持通过 MSBuild 属性 `OfficeInteropPath`、`ExtensibilityInteropPath` 指定官方程序集路径。

完整构建需要预先准备 Office、Build Tools、Rust 和系统运行库；脚本不会自动安装它们。首次构建还需要网络取得 NuGet 引用程序集、Cargo 依赖，以及显式启用引导时的 Inno Setup。导出插件本身无需联网。

### 从源码构建

以下命令均在项目根目录执行。先检查本机环境：

```powershell
.\scripts\check-environment.ps1 -ForBuild -Strict
```

完整发布流程如下。这里将 Cargo 缓存设在项目内，便于集中清理；该环境变量仅影响当前 PowerShell 会话及其子进程。

```powershell
$env:CARGO_HOME = Join-Path (Get-Location).Path '.build\cargo-home'
New-Item -ItemType Directory -Force -Path $env:CARGO_HOME | Out-Null

# 编译前端和原生核心，暂不制作 Setup。
.\scripts\build.ps1 -SkipInstaller

# 根据锁定依赖及其原始许可证生成第三方声明。
.\scripts\generate-notices.ps1

# 复用发布二进制，更新随包文档并制作安装器。
.\scripts\build.ps1 -RepackageOnly -BootstrapInno
```

依赖和第三方声明均已同步时，也可直接运行：

```powershell
.\scripts\build.ps1 -BootstrapInno
```

构建脚本的主要步骤为：

1. 检查本机环境；按参数准备便携 Inno Setup 与 .NET 4.8 引用程序集。
2. 在 `native/` 下执行 `cargo build --locked --release`，加载 `native/.cargo/config.toml`，静态链接 MSVC CRT；默认原生输出位于 `.build/native-release-target/`。
3. 用 MSBuild 编译 x64 Release COM 前端，嵌入 Office 接口类型，整理输出到 `dist/app/`。
4. 复制 README、Apache 许可证、第三方声明，生成 `build-manifest.json`。
5. 调用 Inno Setup 生成 `dist/PPTFigureExporter_Setup.exe`，配置当前用户 COM 注册与依赖检查。

依赖更新后需要重新生成第三方声明；普通构建会复制现有声明，不会自动刷新其内容。便携 Inno 不注册系统卸载项，也不创建开始菜单快捷方式。

### 重新打包与独立运行库

只修改说明、许可或安装脚本时，复用经过验证的 `dist/app/`：

```powershell
.\scripts\build.ps1 -RepackageOnly -BootstrapInno
```

此模式不运行 MSBuild/Cargo，不下载 .NET 引用程序集；必须已存在插件 DLL 和原生核心。已有可用的 Inno 编译器时可省略 `-BootstrapInno`。

只有需要离线运行库附件时，才显式提供微软安装器：

```powershell
.\scripts\build.ps1 -RepackageOnly -BootstrapInno `
  -DotNetOfflineInstaller 'C:\Downloads\dotNet48-x86-x64-AllOS-ENU.exe'
```

附件通过微软 Authenticode 签名检查后复制到 `dist/prereq/`，始终不嵌入主 Setup。输入文件须位于生成目录 `dist/prereq/` 之外，避免被暂存清理删除。

常用参数还包括 `-SkipInstaller`（只编译和整理发布文件）、`-EnvironmentOnly`（仅环境报告）、`-Version`（安装包及清单版本）。`-Version` 不会自动修改 C# 程序集版本或 Cargo 包版本；产品升级时应明确管理各处版本。

## 验证与故障定位

当前保留的验证结论包括：C# 基础检查 35 项断言、PowerPoint 实机联调 55 项断言（含该次基础检查）、Rust 测试 11 项，以及安装后的 COM 加载、功能区和用户完成的 TIFF 导出。不同测试运行的断言数不作简单相加。完整结果与验收边界见 `docs/VALIDATION.md`。

Rust 的开发检查从 `native/` 目录执行：

```powershell
cargo fmt --check
cargo clippy --locked --all-targets -- -D warnings
cargo test --locked
```

前端测试项目为 `tests/PPTFigureExporter.Tests.csproj`，支持基础检查、`--office` 实机联调和 `--render-ui` 布局输出。实机测试会创建专用文稿，运行前需关闭 PowerPoint；操作说明见 `docs/DEVELOPMENT.md`。安装后可运行 `tests/Verify-Installed.ps1` 检查 COM 加载与功能区，该脚本不自动点击保存对话框。

| 现象 | 检查方向 |
| --- | --- |
| 安装器要求先关闭 PowerPoint | 保存文稿并关闭所有 PowerPoint 窗口，再重试 |
| 提示缺少 .NET Framework 4.8 | 安装微软 Runtime；不要用 .NET 6/8 或 Developer Pack 下载替代运行库检查 |
| 找不到 Figure Export | 确认 Office 为 64 位，在 PowerPoint 的 COM 加载项管理中查看插件状态，并检查 `export.log`；组织策略可能限制加载项 |
| 导出按钮不可用或要求组合 | 选中一个有效对象；多个对象先 Ctrl+G，零宽或零高线条需与其他对象组合 |
| 提示并非真实 SVG | 核对 PowerPoint 具体构建及自动化接口能力，不能仅以手动 SVG 导出菜单判断 |
| 超过像素上限或转换超时 | 减小物理尺寸或 DPI；减少复杂滤镜与大幅嵌入位图 |
| 字体、裁剪或效果存在差异 | 检查本机字体与源对象，比较 SVG/PDF 和栅格结果；最终以导出文件为准 |

## 项目结构与文档

```text
frontend/                  C# COM 加载项、Ribbon、设置与导出调用
native/                    Rust 转换核心、Cargo.lock 与核心测试
installer/                 Inno Setup 安装器定义
scripts/                   环境检查、构建与第三方声明生成
tests/                    前端及 Office 联调测试
docs/                     开发、分发、环境与验证记录
dist/                     生成的安装包、发布暂存文件与许可
LICENSE                   Apache License 2.0 正式文本
```

源码文档：

- [开发接口与测试](docs/DEVELOPMENT.md)
- [分发与独立运行库](docs/DISTRIBUTION.md)
- [已验证的构建环境](docs/ENVIRONMENT.md)
- [验证记录与边界](docs/VALIDATION.md)
- [参考实现行为](docs/REFERENCE_BEHAVIOR.md)

`docs/PPT_Figure_Exporter_Design_Brief.md` 保留原始需求；实际实现以代码、本文和验证记录为准。`docs/scripts_refernce/` 仅保留参考源码，不参与插件运行或常规构建。

`.build/`、`native/target/`、各项目 `bin/obj/` 和 `tests/output/` 为可重建的缓存或输出。发布时保留源码、许可证和最终分发文件；`dist/` 本身是生成目录，不是唯一的源码备份。随安装包附带的 README 中，上述开发文档路径均相对于源码项目根目录。

## 许可证与第三方声明

本项目自有源码和文档采用 **Apache License, Version 2.0**，SPDX 标识为 `Apache-2.0`。完整正式条款见根目录 [LICENSE](LICENSE)，原文来源为 [Apache Software Foundation](https://www.apache.org/licenses/LICENSE-2.0.txt)。

第三方组件、其许可证文本以及微软 Office/.NET 等独立产品不因本项目选择 Apache 2.0 而改变许可。Rust 依赖的版本、来源及许可全文保留于 [THIRD_PARTY_NOTICES.txt](docs/THIRD_PARTY_NOTICES.txt)。分发插件或修改版本时，应一并提供项目许可证和适用的第三方声明，并遵守各自条款。

构建会将项目许可证和第三方声明复制到 `dist/licenses/`，安装后放置于应用目录的 `licenses/`。许可证选择不会改变导出文件所含图形、文字、图像或字体素材本身的权利归属。
