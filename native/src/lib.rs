use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs::{self, OpenOptions};
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use roxmltree::{Document, ParsingOptions};
use serde::{Deserialize, Serialize};

const MAX_PIXELS: u64 = 80_000_000;
const MAX_WEBP_DIMENSION: u32 = 16_383;
const DEFAULT_DPI: u32 = 1_000;
const DEFAULT_QUALITY: u8 = 95;
const PDF_FILTER_DPI: f32 = 300.0;
const CSS_DPI: f32 = 96.0;

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

#[derive(Debug)]
pub enum CoreError {
    InvalidRequest(String),
    Input(String),
    Svg(String),
    Render(String),
    Encode(String),
    Output(String),
}

impl Display for CoreError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidRequest(message)
            | Self::Input(message)
            | Self::Svg(message)
            | Self::Render(message)
            | Self::Encode(message)
            | Self::Output(message) => formatter.write_str(message),
        }
    }
}

impl Error for CoreError {}

impl From<io::Error> for CoreError {
    fn from(error: io::Error) -> Self {
        Self::Output(error.to_string())
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum OutputFormat {
    Svg,
    Pdf,
    Png,
    Tiff,
    Jpeg,
    Webp,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum SizeMode {
    Keep,
    Width,
    Height,
    Custom,
    Scale,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum Unit {
    #[default]
    Mm,
    Cm,
    Inch,
}

#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Request {
    pub input_svg: String,
    pub output_path: String,
    pub format: OutputFormat,
    pub source_width_pt: f64,
    pub source_height_pt: f64,
    pub size_mode: SizeMode,
    #[serde(default)]
    pub width: Option<f64>,
    #[serde(default)]
    pub height: Option<f64>,
    #[serde(default)]
    pub unit: Unit,
    #[serde(default)]
    pub percent: Option<f64>,
    #[serde(default = "default_dpi")]
    pub dpi: u32,
    #[serde(default)]
    pub transparent: bool,
    #[serde(default = "default_quality")]
    pub quality: u8,
    #[serde(default)]
    pub overwrite: bool,
}

#[derive(Debug, Serialize)]
pub struct Response {
    pub ok: bool,
    pub output_path: String,
    pub width_pt: f64,
    pub height_pt: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pixel_width: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pixel_height: Option<u32>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TargetSize {
    pub width_pt: f64,
    pub height_pt: f64,
    pub custom_stretch: bool,
    pub pixel_width: Option<u32>,
    pub pixel_height: Option<u32>,
}

fn default_dpi() -> u32 {
    DEFAULT_DPI
}

fn default_quality() -> u8 {
    DEFAULT_QUALITY
}

pub fn process_request(request: &Request) -> Result<Response, CoreError> {
    validate_request_basics(request)?;
    let target = calculate_target_size(request)?;
    let input_path = Path::new(&request.input_svg);
    let output_path = Path::new(&request.output_path);
    validate_paths(input_path, output_path)?;

    let source = fs::read(input_path)
        .map_err(|error| CoreError::Input(format!("无法读取 SVG 输入：{error}")))?;
    let source_text = std::str::from_utf8(&source)
        .map_err(|_| CoreError::Input("SVG 输入不是 UTF-8 文本".to_owned()))?;
    let normalized = normalize_svg(source_text, target)?;

    let output = match request.format {
        OutputFormat::Svg => normalized.into_bytes(),
        OutputFormat::Pdf => render_pdf(&normalized)?,
        OutputFormat::Png | OutputFormat::Tiff | OutputFormat::Jpeg | OutputFormat::Webp => {
            let pixels = render_raster(&normalized, target)?;
            encode_raster(&pixels, target, request)?
        }
    };

    write_atomic(output_path, &output, request.overwrite)?;

    Ok(Response {
        ok: true,
        output_path: request.output_path.clone(),
        width_pt: target.width_pt,
        height_pt: target.height_pt,
        pixel_width: target.pixel_width,
        pixel_height: target.pixel_height,
    })
}

fn validate_request_basics(request: &Request) -> Result<(), CoreError> {
    if request.input_svg.trim().is_empty() {
        return Err(CoreError::InvalidRequest("input_svg 不能为空".to_owned()));
    }
    if request.output_path.trim().is_empty() {
        return Err(CoreError::InvalidRequest("output_path 不能为空".to_owned()));
    }
    validate_positive_finite(request.source_width_pt, "source_width_pt")?;
    validate_positive_finite(request.source_height_pt, "source_height_pt")?;
    if !(1..=9_600).contains(&request.dpi) {
        return Err(CoreError::InvalidRequest(
            "dpi 必须在 1 到 9600 之间".to_owned(),
        ));
    }
    if !(1..=100).contains(&request.quality) {
        return Err(CoreError::InvalidRequest(
            "quality 必须在 1 到 100 之间".to_owned(),
        ));
    }
    Ok(())
}

fn calculate_target_size(request: &Request) -> Result<TargetSize, CoreError> {
    let source_width = request.source_width_pt;
    let source_height = request.source_height_pt;
    let (width_pt, height_pt, custom_stretch) = match request.size_mode {
        SizeMode::Keep => (source_width, source_height, false),
        SizeMode::Width => {
            let width = required_value(request.width, "width")?;
            let width_pt = to_points(width, request.unit, "width")?;
            let height_pt = width_pt * source_height / source_width;
            (width_pt, height_pt, false)
        }
        SizeMode::Height => {
            let height = required_value(request.height, "height")?;
            let height_pt = to_points(height, request.unit, "height")?;
            let width_pt = height_pt * source_width / source_height;
            (width_pt, height_pt, false)
        }
        SizeMode::Custom => {
            let width = required_value(request.width, "width")?;
            let height = required_value(request.height, "height")?;
            let width_pt = to_points(width, request.unit, "width")?;
            let height_pt = to_points(height, request.unit, "height")?;
            (width_pt, height_pt, true)
        }
        SizeMode::Scale => {
            let percent = required_value(request.percent, "percent")?;
            validate_positive_finite(percent, "percent")?;
            let factor = percent / 100.0;
            (source_width * factor, source_height * factor, false)
        }
    };

    validate_positive_finite(width_pt, "目标宽度")?;
    validate_positive_finite(height_pt, "目标高度")?;

    let (pixel_width, pixel_height) = match request.format {
        OutputFormat::Svg | OutputFormat::Pdf => (None, None),
        OutputFormat::Png | OutputFormat::Tiff | OutputFormat::Jpeg | OutputFormat::Webp => {
            let pixel_width = pixels_from_points(width_pt, request.dpi)?;
            let pixel_height = pixels_from_points(height_pt, request.dpi)?;
            let area = u64::from(pixel_width)
                .checked_mul(u64::from(pixel_height))
                .ok_or_else(|| CoreError::InvalidRequest("输出像素数超过可表示范围".to_owned()))?;
            if area > MAX_PIXELS {
                return Err(CoreError::InvalidRequest(
                    "输出超过 8000 万像素，请减小尺寸或 DPI".to_owned(),
                ));
            }
            if request.format == OutputFormat::Webp
                && (pixel_width > MAX_WEBP_DIMENSION || pixel_height > MAX_WEBP_DIMENSION)
            {
                return Err(CoreError::InvalidRequest(
                    "WebP 单边尺寸不能超过 16383 像素".to_owned(),
                ));
            }
            if request.format == OutputFormat::Jpeg
                && (pixel_width > u32::from(u16::MAX) || pixel_height > u32::from(u16::MAX))
            {
                return Err(CoreError::InvalidRequest(
                    "JPEG 单边尺寸不能超过 65535 像素".to_owned(),
                ));
            }
            (Some(pixel_width), Some(pixel_height))
        }
    };

    Ok(TargetSize {
        width_pt,
        height_pt,
        custom_stretch,
        pixel_width,
        pixel_height,
    })
}

fn required_value(value: Option<f64>, name: &str) -> Result<f64, CoreError> {
    let value =
        value.ok_or_else(|| CoreError::InvalidRequest(format!("size_mode 需要提供 {name}")))?;
    validate_positive_finite(value, name)?;
    Ok(value)
}

fn validate_positive_finite(value: f64, name: &str) -> Result<(), CoreError> {
    if !value.is_finite() || value <= 0.0 {
        return Err(CoreError::InvalidRequest(format!("{name} 必须是有限正数")));
    }
    Ok(())
}

fn to_points(value: f64, unit: Unit, name: &str) -> Result<f64, CoreError> {
    let factor = match unit {
        Unit::Mm => 72.0 / 25.4,
        Unit::Cm => 72.0 / 2.54,
        Unit::Inch => 72.0,
    };
    let result = value * factor;
    validate_positive_finite(result, name)?;
    Ok(result)
}

fn pixels_from_points(points: f64, dpi: u32) -> Result<u32, CoreError> {
    let raw = points / 72.0 * f64::from(dpi);
    validate_positive_finite(raw, "像素尺寸")?;
    let rounded = round_ties_even(raw).max(1.0);
    if rounded > f64::from(u32::MAX) {
        return Err(CoreError::InvalidRequest(
            "输出像素尺寸超出可表示范围".to_owned(),
        ));
    }
    Ok(rounded as u32)
}

fn round_ties_even(value: f64) -> f64 {
    let floor = value.floor();
    let fraction = value - floor;
    if fraction < 0.5 {
        floor
    } else if fraction > 0.5 {
        floor + 1.0
    } else if (floor as u64).is_multiple_of(2) {
        floor
    } else {
        floor + 1.0
    }
}

fn validate_paths(input: &Path, output: &Path) -> Result<(), CoreError> {
    if !input.is_file() {
        return Err(CoreError::Input(format!(
            "输入 SVG 不存在或不是文件：{}",
            input.display()
        )));
    }
    let input_absolute = canonical_or_absolute(input)?;
    let output_absolute = canonical_or_absolute(output)?;
    if paths_equal(&input_absolute, &output_absolute) {
        return Err(CoreError::InvalidRequest(
            "input_svg 与 output_path 不能相同".to_owned(),
        ));
    }
    if output.exists() && output.is_dir() {
        return Err(CoreError::Output(format!(
            "输出路径是目录：{}",
            output.display()
        )));
    }
    let parent = output.parent().filter(|path| !path.as_os_str().is_empty());
    let parent = parent.unwrap_or_else(|| Path::new("."));
    if !parent.is_dir() {
        return Err(CoreError::Output(format!(
            "输出目录不存在：{}",
            parent.display()
        )));
    }
    Ok(())
}

fn canonical_or_absolute(path: &Path) -> Result<PathBuf, CoreError> {
    if path.exists() {
        return fs::canonicalize(path)
            .map_err(|error| CoreError::Output(format!("解析路径失败：{error}")));
    }
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let parent = fs::canonicalize(parent)
        .map_err(|error| CoreError::Output(format!("解析输出目录失败：{error}")))?;
    Ok(parent.join(
        path.file_name()
            .ok_or_else(|| CoreError::InvalidRequest("output_path 文件名无效".to_owned()))?,
    ))
}

fn paths_equal(left: &Path, right: &Path) -> bool {
    #[cfg(windows)]
    {
        left.to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy())
    }
    #[cfg(not(windows))]
    {
        left == right
    }
}

fn normalize_svg(source: &str, target: TargetSize) -> Result<String, CoreError> {
    let lower = source.to_ascii_lowercase();
    if lower.contains("<!doctype") || lower.contains("<!entity") {
        return Err(CoreError::Svg("SVG 不允许 DTD 或外部实体引用".to_owned()));
    }
    if lower.contains("<?xml-stylesheet") {
        return Err(CoreError::Svg("SVG 不允许外部样式表处理指令".to_owned()));
    }
    let document = Document::parse_with_options(
        source,
        ParsingOptions {
            allow_dtd: false,
            ..ParsingOptions::default()
        },
    )
    .map_err(|error| CoreError::Svg(format!("SVG XML 解析失败：{error}")))?;
    let root = document.root_element();
    if root.tag_name().name() != "svg" {
        return Err(CoreError::Svg("SVG 根元素必须是 svg".to_owned()));
    }
    let inferred_view_box = if let Some(value) = root.attribute("viewBox") {
        validate_view_box(Some(value))?;
        None
    } else {
        Some(infer_view_box(
            root.attribute("width"),
            root.attribute("height"),
        )?)
    };
    validate_external_references(&document)?;

    let opening_start = root.range().start;
    let opening_end = find_opening_tag_end(source, opening_start)?;
    let replacement = rewrite_root_opening(
        &source[opening_start..opening_end],
        target.width_pt,
        target.height_pt,
        target.custom_stretch,
        inferred_view_box,
    )?;
    let mut result = String::with_capacity(source.len() + replacement.len());
    result.push_str(&source[..opening_start]);
    result.push_str(&replacement);
    result.push_str(&source[opening_end..]);
    Ok(result)
}

fn infer_view_box(width: Option<&str>, height: Option<&str>) -> Result<(f64, f64), CoreError> {
    let width = parse_svg_length_px(
        width.ok_or_else(|| CoreError::Svg("SVG 缺少 viewBox，且根元素缺少 width".to_owned()))?,
    )?;
    let height = parse_svg_length_px(
        height.ok_or_else(|| CoreError::Svg("SVG 缺少 viewBox，且根元素缺少 height".to_owned()))?,
    )?;
    if width <= 0.0 || height <= 0.0 {
        return Err(CoreError::Svg(
            "SVG 缺少 viewBox，且根 width/height 无效".to_owned(),
        ));
    }
    Ok((width, height))
}

fn parse_svg_length_px(value: &str) -> Result<f64, CoreError> {
    let value = value.trim();
    let number_end = value
        .find(|character: char| {
            !character.is_ascii_digit()
                && character != '+'
                && character != '-'
                && character != '.'
                && character != 'e'
                && character != 'E'
        })
        .unwrap_or(value.len());
    let number = value[..number_end]
        .parse::<f64>()
        .map_err(|_| CoreError::Svg("SVG 根 width/height 不是有效数字".to_owned()))?;
    let unit = value[number_end..].trim().to_ascii_lowercase();
    let factor = match unit.as_str() {
        "" | "px" => 1.0,
        "pt" => CSS_DPI as f64 / 72.0,
        "pc" => CSS_DPI as f64 / 6.0,
        "mm" => CSS_DPI as f64 / 25.4,
        "cm" => CSS_DPI as f64 / 2.54,
        "in" => CSS_DPI as f64,
        _ => {
            return Err(CoreError::Svg(
                "SVG 根 width/height 必须使用绝对长度单位".to_owned(),
            ));
        }
    };
    let result = number * factor;
    if !result.is_finite() || result <= 0.0 {
        return Err(CoreError::Svg("SVG 根 width/height 无效".to_owned()));
    }
    Ok(result)
}
fn validate_view_box(value: Option<&str>) -> Result<(), CoreError> {
    let value = value.ok_or_else(|| CoreError::Svg("SVG 必须包含 viewBox".to_owned()))?;
    let replaced = value.replace(',', " ");
    let numbers = replaced
        .split_whitespace()
        .map(str::parse::<f64>)
        .collect::<Result<Vec<_>, _>>()
        .map_err(|_| CoreError::Svg("SVG viewBox 必须包含四个数字".to_owned()))?;
    if numbers.len() != 4
        || !numbers.iter().all(|number| number.is_finite())
        || numbers[2] <= 0.0
        || numbers[3] <= 0.0
    {
        return Err(CoreError::Svg("SVG viewBox 尺寸无效".to_owned()));
    }
    Ok(())
}

fn validate_external_references(document: &Document<'_>) -> Result<(), CoreError> {
    for node in document.descendants() {
        for attribute in node.attributes() {
            let name = attribute.name();
            let value = attribute.value().trim();
            if name.rsplit(':').next() == Some("href") && !is_allowed_reference(value) {
                return Err(CoreError::Svg(format!(
                    "SVG 包含不允许的外部 href：{value}"
                )));
            }
            if value.to_ascii_lowercase().contains("url(") {
                validate_url_functions(value)?;
            }
        }
    }
    Ok(())
}

fn is_allowed_reference(value: &str) -> bool {
    let lower = value.to_ascii_lowercase();
    value.starts_with('#') || lower.starts_with("data:")
}

fn validate_url_functions(value: &str) -> Result<(), CoreError> {
    let lower = value.to_ascii_lowercase();
    let mut offset = 0;
    while let Some(relative) = lower[offset..].find("url(") {
        let start = offset + relative + 4;
        let end = lower[start..]
            .find(')')
            .map(|index| start + index)
            .ok_or_else(|| CoreError::Svg("SVG CSS url() 缺少右括号".to_owned()))?;
        let reference = value[start..end]
            .trim()
            .trim_matches('"')
            .trim_matches('\'');
        if !is_allowed_reference(reference) {
            return Err(CoreError::Svg(format!(
                "SVG 包含不允许的外部 url()：{reference}"
            )));
        }
        offset = end + 1;
    }
    Ok(())
}

fn find_opening_tag_end(source: &str, start: usize) -> Result<usize, CoreError> {
    let bytes = source.as_bytes();
    let mut quote = None;
    for (index, byte) in bytes.iter().enumerate().skip(start) {
        match (quote, *byte) {
            (Some(current), byte) if byte == current => quote = None,
            (None, b'\'' | b'"') => quote = Some(*byte),
            (None, b'>') => return Ok(index + 1),
            _ => {}
        }
    }
    Err(CoreError::Svg("SVG 根元素开始标签不完整".to_owned()))
}

fn rewrite_root_opening(
    opening: &str,
    width_pt: f64,
    height_pt: f64,
    custom_stretch: bool,
    inferred_view_box: Option<(f64, f64)>,
) -> Result<String, CoreError> {
    let bytes = opening.as_bytes();
    let name_end = opening
        .find("svg")
        .ok_or_else(|| CoreError::Svg("无法定位 SVG 根元素开始标签".to_owned()))?
        + 3;
    let gt = opening
        .len()
        .checked_sub(1)
        .ok_or_else(|| CoreError::Svg("SVG 根元素开始标签为空".to_owned()))?;
    let mut output = String::from(&opening[..name_end]);
    let mut position = name_end;
    let mut suffix_start = gt;

    while position < gt {
        let whitespace_start = position;
        while position < gt && bytes[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= gt || bytes[position] == b'/' {
            suffix_start = whitespace_start.min(position);
            break;
        }
        let attribute_start = position;
        while position < gt
            && !bytes[position].is_ascii_whitespace()
            && bytes[position] != b'='
            && bytes[position] != b'/'
        {
            position += 1;
        }
        let name = &opening[attribute_start..position];
        while position < gt && bytes[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= gt || bytes[position] != b'=' {
            return Err(CoreError::Svg("SVG 根元素属性格式无效".to_owned()));
        }
        position += 1;
        while position < gt && bytes[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= gt || !matches!(bytes[position], b'\'' | b'"') {
            return Err(CoreError::Svg("SVG 根元素属性值必须加引号".to_owned()));
        }
        let quote = bytes[position];
        position += 1;
        while position < gt && bytes[position] != quote {
            position += 1;
        }
        if position >= gt {
            return Err(CoreError::Svg("SVG 根元素属性值未闭合".to_owned()));
        }
        position += 1;
        let local_name = name.rsplit(':').next().unwrap_or(name);
        let replaced = local_name == "width"
            || local_name == "height"
            || (custom_stretch && local_name == "preserveAspectRatio");
        if !replaced {
            output.push_str(&opening[whitespace_start..position]);
        }
    }

    output.push_str(&format!(
        " width=\"{}pt\" height=\"{}pt\"",
        format_length(width_pt),
        format_length(height_pt)
    ));
    if let Some((view_box_width, view_box_height)) = inferred_view_box {
        output.push_str(&format!(
            " viewBox=\"0 0 {} {}\"",
            format_length(view_box_width),
            format_length(view_box_height),
        ));
    }
    if custom_stretch {
        output.push_str(" preserveAspectRatio=\"none\"");
    }
    output.push_str(&opening[suffix_start..]);
    Ok(output)
}

fn format_length(value: f64) -> String {
    let mut result = format!("{value:.9}");
    while result.contains('.') && result.ends_with('0') {
        result.pop();
    }
    if result.ends_with('.') {
        result.pop();
    }
    result
}

fn parse_tree(svg: &str, input_parent: Option<&Path>) -> Result<usvg::Tree, CoreError> {
    let mut options = usvg::Options {
        font_family: "Arial".to_owned(),
        dpi: CSS_DPI,
        resources_dir: input_parent.map(Path::to_path_buf),
        ..usvg::Options::default()
    };
    options.fontdb_mut().load_system_fonts();
    usvg::Tree::from_str(svg, &options)
        .map_err(|error| CoreError::Render(format!("SVG 渲染树解析失败：{error}")))
}

fn render_pdf(svg: &str) -> Result<Vec<u8>, CoreError> {
    let tree = parse_tree(svg, None)?;
    let conversion_options = svg2pdf::ConversionOptions {
        raster_scale: PDF_FILTER_DPI / CSS_DPI,
        ..svg2pdf::ConversionOptions::default()
    };
    svg2pdf::to_pdf(
        &tree,
        conversion_options,
        svg2pdf::PageOptions { dpi: CSS_DPI },
    )
    .map_err(|error| CoreError::Render(format!("SVG 转 PDF 失败：{error}")))
}

struct RasterPixels {
    width: u32,
    height: u32,
    rgba: Vec<u8>,
}

fn render_raster(svg: &str, target: TargetSize) -> Result<RasterPixels, CoreError> {
    let width = target
        .pixel_width
        .ok_or_else(|| CoreError::Render("栅格输出缺少像素宽度".to_owned()))?;
    let height = target
        .pixel_height
        .ok_or_else(|| CoreError::Render("栅格输出缺少像素高度".to_owned()))?;
    let tree = parse_tree(svg, None)?;
    let tree_width = f64::from(tree.size().width());
    let tree_height = f64::from(tree.size().height());
    if !tree_width.is_finite()
        || !tree_height.is_finite()
        || tree_width <= 0.0
        || tree_height <= 0.0
    {
        return Err(CoreError::Render("SVG 渲染树尺寸无效".to_owned()));
    }
    let mut pixmap = resvg::tiny_skia::Pixmap::new(width, height)
        .ok_or_else(|| CoreError::Render("无法分配栅格像素缓冲区".to_owned()))?;
    let scale_x = f64::from(width) / tree_width;
    let scale_y = f64::from(height) / tree_height;
    let transform = resvg::tiny_skia::Transform::from_scale(scale_x as f32, scale_y as f32);
    resvg::render(&tree, transform, &mut pixmap.as_mut());

    let premultiplied = pixmap.data();
    let mut rgba = Vec::with_capacity(premultiplied.len());
    for pixel in premultiplied.chunks_exact(4) {
        let alpha = pixel[3];
        rgba.push(unpremultiply(pixel[0], alpha));
        rgba.push(unpremultiply(pixel[1], alpha));
        rgba.push(unpremultiply(pixel[2], alpha));
        rgba.push(alpha);
    }
    Ok(RasterPixels {
        width,
        height,
        rgba,
    })
}

fn unpremultiply(value: u8, alpha: u8) -> u8 {
    if alpha == 0 {
        return 0;
    }
    let result = (u16::from(value) * 255 + u16::from(alpha) / 2) / u16::from(alpha);
    result.min(255) as u8
}

fn encode_raster(
    pixels: &RasterPixels,
    target: TargetSize,
    request: &Request,
) -> Result<Vec<u8>, CoreError> {
    let wants_alpha = request.transparent
        && matches!(
            request.format,
            OutputFormat::Png | OutputFormat::Tiff | OutputFormat::Webp
        );
    let data = composite_white(&pixels.rgba, wants_alpha);
    let output = match request.format {
        OutputFormat::Png => {
            encode_png(pixels.width, pixels.height, &data, wants_alpha, request.dpi)
        }
        OutputFormat::Tiff => {
            encode_tiff(pixels.width, pixels.height, &data, wants_alpha, request.dpi)
        }
        OutputFormat::Jpeg => encode_jpeg(pixels.width, pixels.height, &data, request),
        OutputFormat::Webp => encode_webp(pixels.width, pixels.height, &data, wants_alpha, request),
        OutputFormat::Svg | OutputFormat::Pdf => Err(CoreError::Encode(
            "内部错误：矢量格式进入栅格编码器".to_owned(),
        )),
    }?;
    debug_assert_eq!(target.pixel_width, Some(pixels.width));
    debug_assert_eq!(target.pixel_height, Some(pixels.height));
    Ok(output)
}
fn composite_white(rgba: &[u8], preserve_alpha: bool) -> Vec<u8> {
    if preserve_alpha {
        return rgba.to_vec();
    }
    let mut rgb = Vec::with_capacity(rgba.len() / 4 * 3);
    for pixel in rgba.chunks_exact(4) {
        let alpha = u16::from(pixel[3]);
        let inverse = 255 - alpha;
        rgb.push(((u16::from(pixel[0]) * alpha + 255 * inverse + 127) / 255) as u8);
        rgb.push(((u16::from(pixel[1]) * alpha + 255 * inverse + 127) / 255) as u8);
        rgb.push(((u16::from(pixel[2]) * alpha + 255 * inverse + 127) / 255) as u8);
    }
    rgb
}

fn encode_png(
    width: u32,
    height: u32,
    data: &[u8],
    alpha: bool,
    dpi: u32,
) -> Result<Vec<u8>, CoreError> {
    let mut output = Vec::new();
    let mut encoder = png::Encoder::new(&mut output, width, height);
    encoder.set_depth(png::BitDepth::Eight);
    encoder.set_color(if alpha {
        png::ColorType::Rgba
    } else {
        png::ColorType::Rgb
    });
    let ppm = (f64::from(dpi) * 39.37007874015748).round();
    if ppm > f64::from(u32::MAX) {
        return Err(CoreError::Encode("PNG DPI 元数据超出范围".to_owned()));
    }
    encoder.set_pixel_dims(Some(png::PixelDimensions {
        xppu: ppm as u32,
        yppu: ppm as u32,
        unit: png::Unit::Meter,
    }));
    let mut writer = encoder
        .write_header()
        .map_err(|error| CoreError::Encode(format!("PNG 头写入失败：{error}")))?;
    writer
        .write_image_data(data)
        .map_err(|error| CoreError::Encode(format!("PNG 编码失败：{error}")))?;
    drop(writer);
    Ok(output)
}

fn encode_tiff(
    width: u32,
    height: u32,
    data: &[u8],
    alpha: bool,
    dpi: u32,
) -> Result<Vec<u8>, CoreError> {
    use tiff::encoder::colortype::{RGB8, RGBA8};
    use tiff::encoder::compression::Lzw;
    use tiff::encoder::{Rational, TiffEncoder};
    use tiff::tags::{ResolutionUnit, Tag};

    let mut output = std::io::Cursor::new(Vec::new());
    let mut encoder = TiffEncoder::new(&mut output)
        .map_err(|error| CoreError::Encode(format!("TIFF 初始化失败：{error}")))?;
    let resolution = Rational { n: dpi, d: 1 };
    if alpha {
        let mut image = encoder
            .new_image_with_compression::<RGBA8, Lzw>(width, height, Lzw)
            .map_err(|error| CoreError::Encode(format!("TIFF 图像初始化失败：{error}")))?;
        image.resolution(ResolutionUnit::Inch, resolution);
        image
            .encoder()
            .write_tag(Tag::ExtraSamples, &[2u16][..])
            .map_err(|error| CoreError::Encode(format!("TIFF alpha 标签写入失败：{error}")))?;
        image
            .write_data(data)
            .map_err(|error| CoreError::Encode(format!("TIFF 编码失败：{error}")))?;
    } else {
        let mut image = encoder
            .new_image_with_compression::<RGB8, Lzw>(width, height, Lzw)
            .map_err(|error| CoreError::Encode(format!("TIFF 图像初始化失败：{error}")))?;
        image.resolution(ResolutionUnit::Inch, resolution);
        image
            .write_data(data)
            .map_err(|error| CoreError::Encode(format!("TIFF 编码失败：{error}")))?;
    }
    Ok(output.into_inner())
}

fn encode_jpeg(
    width: u32,
    height: u32,
    data: &[u8],
    request: &Request,
) -> Result<Vec<u8>, CoreError> {
    use jpeg_encoder::{ColorType, Density, Encoder};

    let mut output = Vec::new();
    let mut encoder = Encoder::new(&mut output, request.quality);
    let density = u16::try_from(request.dpi)
        .map_err(|_| CoreError::Encode("JPEG DPI 超出编码器范围".to_owned()))?;
    encoder.set_density(Density::Inch {
        x: density,
        y: density,
    });
    encoder
        .encode(data, width as u16, height as u16, ColorType::Rgb)
        .map_err(|error| CoreError::Encode(format!("JPEG 编码失败：{error}")))?;
    Ok(output)
}

fn encode_webp(
    width: u32,
    height: u32,
    data: &[u8],
    alpha: bool,
    request: &Request,
) -> Result<Vec<u8>, CoreError> {
    let encoded = if alpha {
        webp::Encoder::from_rgba(data, width, height).encode(f32::from(request.quality))
    } else {
        webp::Encoder::from_rgb(data, width, height).encode(f32::from(request.quality))
    };
    Ok(encoded.to_vec())
}

fn write_atomic(path: &Path, data: &[u8], overwrite: bool) -> Result<(), CoreError> {
    if path.exists() && !overwrite {
        return Err(CoreError::Output(format!(
            "输出文件已存在且 overwrite=false：{}",
            path.display()
        )));
    }
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| CoreError::Output("输出文件名不是有效的 UTF-8 文件名".to_owned()))?;
    let temporary = loop {
        let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".{file_name}.figure-export-{}.{}tmp",
            std::process::id(),
            counter
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(mut file) => {
                let write_result = file
                    .write_all(data)
                    .and_then(|_| file.flush())
                    .and_then(|_| file.sync_all());
                if let Err(error) = write_result {
                    drop(file);
                    let _ = fs::remove_file(&candidate);
                    return Err(CoreError::Output(format!("写入临时输出失败：{error}")));
                }
                break candidate;
            }
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(CoreError::Output(format!("创建临时输出失败：{error}")));
            }
        }
    };

    let commit_result = commit_temp(&temporary, path, overwrite);
    if commit_result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    commit_result
}

#[cfg(windows)]
fn commit_temp(temporary: &Path, destination: &Path, overwrite: bool) -> Result<(), CoreError> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Storage::FileSystem::{
        MoveFileExW, MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH,
    };

    if !overwrite && destination.exists() {
        return Err(CoreError::Output(format!(
            "输出文件已存在且 overwrite=false：{}",
            destination.display()
        )));
    }
    let source_wide = temporary
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let destination_wide = destination
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let mut flags = MOVEFILE_WRITE_THROUGH;
    if overwrite {
        flags |= MOVEFILE_REPLACE_EXISTING;
    }
    let result = unsafe { MoveFileExW(source_wide.as_ptr(), destination_wide.as_ptr(), flags) };
    if result == 0 {
        return Err(CoreError::Output(format!(
            "原子替换输出文件失败：{}",
            destination.display()
        )));
    }
    Ok(())
}

#[cfg(not(windows))]
fn commit_temp(temporary: &Path, destination: &Path, overwrite: bool) -> Result<(), CoreError> {
    if !overwrite && destination.exists() {
        return Err(CoreError::Output(format!(
            "输出文件已存在且 overwrite=false：{}",
            destination.display()
        )));
    }
    if overwrite && destination.exists() {
        fs::remove_file(destination)?;
    }
    fs::rename(temporary, destination)
        .map_err(|error| CoreError::Output(format!("原子替换输出文件失败：{error}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn request(mode: SizeMode) -> Request {
        Request {
            input_svg: "input.svg".to_owned(),
            output_path: "output.png".to_owned(),
            format: OutputFormat::Png,
            source_width_pt: 144.0,
            source_height_pt: 72.0,
            size_mode: mode,
            width: Some(25.4),
            height: Some(12.7),
            unit: Unit::Mm,
            percent: Some(50.0),
            dpi: 144,
            transparent: true,
            quality: 95,
            overwrite: false,
        }
    }

    #[test]
    fn target_size_keep_uses_source_points() {
        let target = calculate_target_size(&request(SizeMode::Keep)).unwrap();
        assert_eq!(target.width_pt, 144.0);
        assert_eq!(target.height_pt, 72.0);
        assert_eq!(target.pixel_width, Some(288));
        assert_eq!(target.pixel_height, Some(144));
    }

    #[test]
    fn target_size_width_preserves_ratio_and_units() {
        let target = calculate_target_size(&request(SizeMode::Width)).unwrap();
        assert!((target.width_pt - 72.0).abs() < 1e-10);
        assert!((target.height_pt - 36.0).abs() < 1e-10);
    }

    #[test]
    fn target_size_height_preserves_ratio() {
        let target = calculate_target_size(&request(SizeMode::Height)).unwrap();
        assert!((target.width_pt - 72.0).abs() < 1e-10);
        assert!((target.height_pt - 36.0).abs() < 1e-10);
    }

    #[test]
    fn target_size_custom_stretches() {
        let target = calculate_target_size(&request(SizeMode::Custom)).unwrap();
        assert_eq!(target.width_pt, 72.0);
        assert_eq!(target.height_pt, 36.0);
        assert!(target.custom_stretch);
    }

    #[test]
    fn target_size_scale_uses_percent() {
        let target = calculate_target_size(&request(SizeMode::Scale)).unwrap();
        assert_eq!(target.width_pt, 72.0);
        assert_eq!(target.height_pt, 36.0);
    }

    #[test]
    fn round_ties_to_even_matches_python_round() {
        assert_eq!(round_ties_even(2.5), 2.0);
        assert_eq!(round_ties_even(3.5), 4.0);
        assert_eq!(round_ties_even(2.49), 2.0);
        assert_eq!(round_ties_even(2.51), 3.0);
    }

    #[test]
    fn invalid_and_large_raster_inputs_are_rejected() {
        let mut bad_dpi = request(SizeMode::Keep);
        bad_dpi.dpi = 0;
        assert!(validate_request_basics(&bad_dpi).is_err());

        let mut too_large = request(SizeMode::Keep);
        too_large.source_width_pt = 72_000.0;
        too_large.source_height_pt = 72_000.0;
        too_large.dpi = 1_000;
        assert!(calculate_target_size(&too_large).is_err());
    }

    #[test]
    fn svg_rewrite_preserves_geometry_and_sets_custom_aspect_ratio() {
        let svg = r#"<?xml version="1.0"?><svg viewBox="0 0 100 50" width="100" height="50" preserveAspectRatio="xMidYMid meet"><g transform="scale(2)"><path d="M0 0L10 10"/></g></svg>"#;
        let target = TargetSize {
            width_pt: 144.0,
            height_pt: 72.0,
            custom_stretch: true,
            pixel_width: None,
            pixel_height: None,
        };
        let result = normalize_svg(svg, target).unwrap();
        assert!(result.contains("width=\"144pt\""));
        assert!(result.contains("height=\"72pt\""));
        assert!(result.contains("preserveAspectRatio=\"none\""));
        assert!(result.contains("<path d=\"M0 0L10 10\"/>"));
        assert!(result.contains("transform=\"scale(2)\""));
    }

    #[test]
    fn svg_external_references_and_missing_viewbox_are_rejected() {
        let external = r#"<svg viewBox="0 0 10 10"><image href="file:///x.png"/></svg>"#;
        assert!(normalize_svg(
            external,
            TargetSize {
                width_pt: 10.0,
                height_pt: 10.0,
                custom_stretch: false,
                pixel_width: None,
                pixel_height: None,
            }
        )
        .is_err());

        let missing = r#"<svg width="195" height="126"></svg>"#;
        let normalized = normalize_svg(
            missing,
            TargetSize {
                width_pt: 10.0,
                height_pt: 10.0,
                custom_stretch: false,
                pixel_width: None,
                pixel_height: None,
            },
        )
        .unwrap();
        assert!(normalized.contains("viewBox=\"0 0 195 126\""));

        let no_size = r#"<svg></svg>"#;
        assert!(normalize_svg(
            no_size,
            TargetSize {
                width_pt: 10.0,
                height_pt: 10.0,
                custom_stretch: false,
                pixel_width: None,
                pixel_height: None,
            }
        )
        .is_err());
    }

    #[test]
    fn alpha_unpremultiplication_and_white_compositing_are_correct() {
        assert_eq!(unpremultiply(64, 128), 128);
        assert_eq!(unpremultiply(0, 0), 0);
        assert_eq!(
            composite_white(&[100, 50, 0, 128], false),
            vec![177, 152, 127]
        );
        assert_eq!(
            composite_white(&[100, 50, 0, 128], true),
            vec![100, 50, 0, 128]
        );
    }
}
