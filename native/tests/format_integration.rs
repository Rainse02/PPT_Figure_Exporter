use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use figure_export_core::{process_request, OutputFormat, Request, SizeMode, Unit};

static TEST_COUNTER: AtomicU64 = AtomicU64::new(0);

fn test_directory() -> PathBuf {
    let counter = TEST_COUNTER.fetch_add(1, Ordering::Relaxed);
    let directory = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../.build/native-tests")
        .join(format!(
            "figure-export-core-test-{}-{counter}",
            std::process::id()
        ));
    fs::create_dir_all(&directory).unwrap();
    directory
}

fn request(input: &Path, output: &Path, format: OutputFormat) -> Request {
    Request {
        input_path: input.to_string_lossy().into_owned(),
        output_path: output.to_string_lossy().into_owned(),
        format,
        source_width_pt: 144.0,
        source_height_pt: 72.0,
        size_mode: SizeMode::Keep,
        width: None,
        height: None,
        unit: Unit::Inch,
        percent: None,
        dpi: 72,
        transparent: true,
        quality: 80,
        overwrite: false,
    }
}

#[test]
fn all_six_formats_produce_valid_output() {
    let directory = test_directory();
    let input = directory.join("fixture.svg");
    let svg = r##"<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="2in" height="1in" viewBox="0 0 200 100">
  <rect x="20" y="20" width="160" height="60" fill="#ff0000" fill-opacity="0.5"/>
  <rect x="100" y="20" width="80" height="60" fill="#0000ff" fill-opacity="0.25"/>
  <path d="M 10 10 L 190 90" stroke="#ffffff" stroke-width="2"/>
</svg>
"##;
    fs::write(&input, svg).unwrap();

    let formats = [
        (OutputFormat::Svg, "figure.svg"),
        (OutputFormat::Pdf, "figure.pdf"),
        (OutputFormat::Png, "figure.png"),
        (OutputFormat::Tiff, "figure.tiff"),
        (OutputFormat::Jpeg, "figure.jpg"),
        (OutputFormat::Webp, "figure.webp"),
    ];
    for (format, name) in formats {
        let output = directory.join(name);
        let response = process_request(&request(&input, &output, format)).unwrap();
        assert!(response.ok);
        assert_eq!(response.width_pt, 144.0);
        assert_eq!(response.height_pt, 72.0);
        if matches!(
            format,
            OutputFormat::Png | OutputFormat::Tiff | OutputFormat::Jpeg | OutputFormat::Webp
        ) {
            assert_eq!(response.pixel_width, Some(144));
            assert_eq!(response.pixel_height, Some(72));
        } else {
            assert_eq!(response.pixel_width, None);
            assert_eq!(response.pixel_height, None);
        }
        assert!(output.is_file());

        let bytes = fs::read(&output).unwrap();
        match format {
            OutputFormat::Svg => {
                let text = String::from_utf8(bytes).unwrap();
                assert!(text.contains("width=\"144pt\""));
                assert!(text.contains("<path d=\"M 10 10 L 190 90\""));
            }
            OutputFormat::Pdf => {
                assert!(bytes.starts_with(b"%PDF-"));
                let text = String::from_utf8_lossy(&bytes);
                assert!(text.contains("/MediaBox [0 0 144 72]"));
            }
            OutputFormat::Png => inspect_png(&bytes),
            OutputFormat::Tiff => inspect_tiff(&bytes),
            OutputFormat::Jpeg => inspect_jpeg(&bytes),
            OutputFormat::Webp => inspect_webp(&bytes),
        }
    }
    let _ = fs::remove_dir_all(directory);
}

fn inspect_png(bytes: &[u8]) {
    let decoder = png::Decoder::new(Cursor::new(bytes));
    let mut reader = decoder.read_info().unwrap();
    let info = reader.info();
    assert_eq!((info.width, info.height), (144, 72));
    assert_eq!(info.color_type, png::ColorType::Rgba);
    let dimensions = info.pixel_dims.unwrap();
    assert_eq!(dimensions.unit, png::Unit::Meter);
    assert_eq!(dimensions.xppu, dimensions.yppu);
    let mut buffer = vec![0; reader.output_buffer_size()];
    let output_info = reader.next_frame(&mut buffer).unwrap();
    assert_eq!(output_info.buffer_size(), 144 * 72 * 4);
    assert_eq!(buffer[3], 0);
    let center = ((36 * 144 + 36) * 4) as usize;
    assert!(buffer[center + 3] > 0);
    assert!(buffer[center + 3] < 255);
}

fn inspect_tiff(bytes: &[u8]) {
    use tiff::decoder::Decoder;
    use tiff::tags::Tag;

    let mut decoder = Decoder::new(Cursor::new(bytes)).unwrap();
    assert_eq!(decoder.dimensions().unwrap(), (144, 72));
    assert!(matches!(
        decoder.colortype().unwrap(),
        tiff::ColorType::RGBA(8)
    ));
    assert_eq!(decoder.get_tag_u32(Tag::ResolutionUnit).unwrap(), 2);
    assert_eq!(
        decoder.get_tag_u32_vec(Tag::XResolution).unwrap(),
        vec![72, 1]
    );
    assert_eq!(
        decoder.get_tag_u32_vec(Tag::YResolution).unwrap(),
        vec![72, 1]
    );
    assert_eq!(decoder.get_tag_u32(Tag::ExtraSamples).unwrap(), 2);
    let image = decoder.read_image().unwrap();
    match image {
        tiff::decoder::DecodingResult::U8(data) => {
            assert_eq!(data[3], 0);
            let center = ((36 * 144 + 36) * 4) as usize;
            assert!(data[center + 3] > 0);
            assert!(data[center + 3] < 255);
        }
        other => panic!("unexpected TIFF sample type: {other:?}"),
    }
}

fn inspect_jpeg(bytes: &[u8]) {
    let mut decoder = jpeg_decoder::Decoder::new(Cursor::new(bytes));
    let pixels = decoder.decode().unwrap();
    let info = decoder.info().unwrap();
    assert_eq!((info.width, info.height), (144, 72));
    assert_eq!(info.pixel_format, jpeg_decoder::PixelFormat::RGB24);
    assert!(pixels[0] > 245);
    assert!(pixels[1] > 245);
    assert!(pixels[2] > 245);
}

fn inspect_webp(bytes: &[u8]) {
    assert_eq!(&bytes[..4], b"RIFF");
    assert_eq!(&bytes[8..12], b"WEBP");
    let image = webp::Decoder::new(bytes).decode().unwrap();
    assert_eq!((image.width(), image.height()), (144, 72));
    assert!(image.is_alpha());
    assert_eq!(image[3], 0);
    let center = ((36 * 144 + 36) * 4) as usize;
    assert!(image[center + 3] > 0);
    assert!(image[center + 3] < 255);
}
