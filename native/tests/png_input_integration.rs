use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use figure_export_core::{process_request, OutputFormat, Request, SizeMode, Unit};
use tiff::decoder::{Decoder as TiffDecoder, DecodingResult};
use tiff::tags::Tag;

static TEST_COUNTER: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Self {
        let counter = TEST_COUNTER.fetch_add(1, Ordering::Relaxed);
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../.build/native-tests")
            .join(format!(
                "figure-export-png-test-{}-{counter}",
                std::process::id()
            ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn join(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn raster_request(input: &Path, output: &Path, width: u32, height: u32, dpi: u32) -> Request {
    Request {
        input_path: input.to_string_lossy().into_owned(),
        output_path: output.to_string_lossy().into_owned(),
        format: OutputFormat::Tiff,
        source_width_pt: f64::from(width) * 72.0 / f64::from(dpi),
        source_height_pt: f64::from(height) * 72.0 / f64::from(dpi),
        size_mode: SizeMode::Keep,
        width: None,
        height: None,
        unit: Unit::Inch,
        percent: None,
        dpi,
        transparent: true,
        quality: 95,
        overwrite: false,
    }
}

fn png_bytes(
    width: u32,
    height: u32,
    color: png::ColorType,
    depth: png::BitDepth,
    data: &[u8],
) -> Vec<u8> {
    let mut output = Vec::new();
    let mut encoder = png::Encoder::new(&mut output, width, height);
    encoder.set_color(color);
    encoder.set_depth(depth);
    let mut writer = encoder.write_header().unwrap();
    writer.write_image_data(data).unwrap();
    drop(writer);
    output
}

fn png_header(width: u32, height: u32) -> Vec<u8> {
    let mut output = Vec::new();
    let mut encoder = png::Encoder::new(&mut output, width, height);
    encoder.set_color(png::ColorType::Rgba);
    encoder.set_depth(png::BitDepth::Eight);
    let writer = encoder.write_header().unwrap();
    drop(writer);
    output
}

fn tiff_decoder(path: &Path) -> TiffDecoder<Cursor<Vec<u8>>> {
    TiffDecoder::new(Cursor::new(fs::read(path).unwrap())).unwrap()
}

fn tiff_pixels(decoder: &mut TiffDecoder<Cursor<Vec<u8>>>) -> Vec<u8> {
    match decoder.read_image().unwrap() {
        DecodingResult::U8(pixels) => pixels,
        other => panic!("unexpected TIFF sample type: {other:?}"),
    }
}

#[test]
fn rgba_png_to_tiff_preserves_every_sample_and_requested_dpi() {
    let directory = TestDirectory::new();
    let input = directory.join("rendered-object.data");
    let output = directory.join("figure.tiff");
    let width = 317;
    let height = 183;
    let dpi = 1_000;
    let alpha_values = [0, 1, 127, 128, 254, 255];
    let mut expected = Vec::with_capacity((width * height * 4) as usize);
    for y in 0..height {
        for x in 0..width {
            expected.extend_from_slice(&[
                ((x * 17 + y * 3) % 256) as u8,
                ((x * 5 + y * 19) % 256) as u8,
                ((x * 31 + y * 7) % 256) as u8,
                alpha_values[((x + y) as usize) % alpha_values.len()],
            ]);
        }
    }
    fs::write(
        &input,
        png_bytes(
            width,
            height,
            png::ColorType::Rgba,
            png::BitDepth::Eight,
            &expected,
        ),
    )
    .unwrap();

    let response = process_request(&raster_request(&input, &output, width, height, dpi)).unwrap();
    assert_eq!(response.pixel_width, Some(width));
    assert_eq!(response.pixel_height, Some(height));
    let mut decoder = tiff_decoder(&output);
    assert_eq!(decoder.dimensions().unwrap(), (width, height));
    assert!(matches!(
        decoder.colortype().unwrap(),
        tiff::ColorType::RGBA(8)
    ));
    assert_eq!(decoder.get_tag_u32(Tag::ExtraSamples).unwrap(), 2);
    assert_eq!(decoder.get_tag_u32(Tag::ResolutionUnit).unwrap(), 2);
    assert_eq!(
        decoder.get_tag_u32_vec(Tag::XResolution).unwrap(),
        vec![dpi, 1]
    );
    assert_eq!(
        decoder.get_tag_u32_vec(Tag::YResolution).unwrap(),
        vec![dpi, 1]
    );
    assert_eq!(tiff_pixels(&mut decoder), expected);
}

#[test]
fn smaller_png_is_centered_without_scaling_and_composites_over_white() {
    let directory = TestDirectory::new();
    let input = directory.join("small.png");
    fs::write(
        &input,
        png_bytes(
            2,
            1,
            png::ColorType::Rgba,
            png::BitDepth::Eight,
            &[100, 50, 0, 128, 7, 9, 11, 0],
        ),
    )
    .unwrap();
    let transparent_row = [0u8; 20];
    let transparent_content_row = [
        0, 0, 0, 0, 100, 50, 0, 128, 7, 9, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    ];
    let transparent_expected = [
        transparent_row.as_slice(),
        transparent_content_row.as_slice(),
        transparent_row.as_slice(),
        transparent_row.as_slice(),
    ]
    .concat();
    let white_row = [255u8; 15];
    let white_content_row = [
        255, 255, 255, 177, 152, 127, 255, 255, 255, 255, 255, 255, 255, 255, 255,
    ];
    let white_expected = [
        white_row.as_slice(),
        white_content_row.as_slice(),
        white_row.as_slice(),
        white_row.as_slice(),
    ]
    .concat();

    for (transparent, expected) in [(true, transparent_expected), (false, white_expected)] {
        let output = directory.join(&format!("padded-{transparent}.tiff"));
        let mut request = raster_request(&input, &output, 5, 4, 72);
        request.transparent = transparent;
        process_request(&request).unwrap();
        let mut decoder = tiff_decoder(&output);
        assert_eq!(decoder.dimensions().unwrap(), (5, 4));
        if transparent {
            assert!(matches!(
                decoder.colortype().unwrap(),
                tiff::ColorType::RGBA(8)
            ));
        } else {
            assert!(matches!(
                decoder.colortype().unwrap(),
                tiff::ColorType::RGB(8)
            ));
        }
        assert_eq!(tiff_pixels(&mut decoder), expected);
    }
}

#[test]
fn png_color_types_and_bit_depths_decode_to_rgba() {
    let directory = TestDirectory::new();
    let fixtures = [
        (
            png::ColorType::Rgb,
            png::BitDepth::Eight,
            vec![3, 17, 251, 200, 99, 0],
            vec![3, 17, 251, 255, 200, 99, 0, 255],
        ),
        (
            png::ColorType::Grayscale,
            png::BitDepth::Eight,
            vec![17, 203],
            vec![17, 17, 17, 255, 203, 203, 203, 255],
        ),
        (
            png::ColorType::GrayscaleAlpha,
            png::BitDepth::Eight,
            vec![17, 128, 203, 0],
            vec![17, 17, 17, 128, 203, 203, 203, 0],
        ),
        (
            png::ColorType::Rgba,
            png::BitDepth::Sixteen,
            vec![
                3, 99, 17, 88, 251, 77, 128, 66, 200, 55, 99, 44, 0, 33, 255, 22,
            ],
            vec![3, 17, 251, 128, 200, 99, 0, 255],
        ),
        (
            png::ColorType::Grayscale,
            png::BitDepth::One,
            vec![0b0100_0000],
            vec![0, 0, 0, 255, 255, 255, 255, 255],
        ),
    ];
    for (index, (color, depth, samples, expected)) in fixtures.into_iter().enumerate() {
        let input = directory.join(&format!("fixture-{index}.png"));
        let output = directory.join(&format!("fixture-{index}.tiff"));
        fs::write(&input, png_bytes(2, 1, color, depth, &samples)).unwrap();
        process_request(&raster_request(&input, &output, 2, 1, 72)).unwrap();
        let mut decoder = tiff_decoder(&output);
        assert_eq!(decoder.dimensions().unwrap(), (2, 1));
        assert_eq!(tiff_pixels(&mut decoder), expected);
    }
}

#[test]
fn png_dimensions_are_checked_before_decoding_image_data() {
    let directory = TestDirectory::new();
    let cases = [
        (80_000_001, 1, "8000 万像素"),
        (6, 2, "超过目标画布"),
        (2, 5, "超过目标画布"),
    ];
    for (index, (width, height, message)) in cases.into_iter().enumerate() {
        let input = directory.join(&format!("oversized-{index}.png"));
        let output = directory.join(&format!("oversized-{index}.tiff"));
        fs::write(&input, png_header(width, height)).unwrap();
        let error = process_request(&raster_request(&input, &output, 5, 4, 72)).unwrap_err();
        assert!(error.to_string().contains(message), "{error}");
        assert!(!output.exists());
    }
}

#[test]
fn png_input_rejects_vector_output() {
    let directory = TestDirectory::new();
    let input = directory.join("fixture.png");
    fs::write(
        &input,
        png_bytes(
            1,
            1,
            png::ColorType::Rgba,
            png::BitDepth::Eight,
            &[10, 20, 30, 255],
        ),
    )
    .unwrap();
    for (format, name) in [
        (OutputFormat::Svg, "output.svg"),
        (OutputFormat::Pdf, "output.pdf"),
    ] {
        let output = directory.join(name);
        let mut request = raster_request(&input, &output, 1, 1, 72);
        request.format = format;
        let error = process_request(&request).unwrap_err();
        assert!(error.to_string().contains("PNG 输入仅支持"), "{error}");
        assert!(!output.exists());
    }
}

#[test]
fn request_json_accepts_input_path_and_legacy_input_svg() {
    let directory = TestDirectory::new();
    let input = directory.join("legacy.svg");
    fs::write(
        &input,
        r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 72 36"><rect width="72" height="36"/></svg>"#,
    )
    .unwrap();
    for key in ["input_path", "input_svg"] {
        let output = directory.join(&format!("{key}.svg"));
        let mut json = serde_json::json!({
            "output_path": output,
            "format": "svg",
            "source_width_pt": 72.0,
            "source_height_pt": 36.0,
            "size_mode": "keep"
        });
        json[key] = serde_json::json!(input);
        let request: Request = serde_json::from_value(json).unwrap();
        assert_eq!(Path::new(&request.input_path), input);
        process_request(&request).unwrap();
        let result = fs::read_to_string(output).unwrap();
        assert!(result.contains("width=\"72pt\" height=\"36pt\""));
    }
}
