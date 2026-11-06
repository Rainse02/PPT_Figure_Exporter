use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use figure_export_core::{process_request, OutputFormat, PdfCrop, Request, SizeMode, Unit};
use lopdf::content::{Content, Operation};
use lopdf::{dictionary, Dictionary, Document, Object, ObjectId, Stream};

static TEST_COUNTER: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Self {
        let counter = TEST_COUNTER.fetch_add(1, Ordering::Relaxed);
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../.build/native-tests")
            .join(format!(
                "figure-export-pdf-test-{}-{counter}",
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

struct Fixture {
    document: Document,
    page_id: ObjectId,
    pages_id: ObjectId,
    content_ids: Vec<ObjectId>,
    font_id: ObjectId,
    image_id: Option<ObjectId>,
}

impl Fixture {
    fn new(with_image: bool) -> Self {
        let mut document = Document::with_version("1.7");
        let pages_id = document.new_object_id();
        let font_id = document.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
        });
        let mut resources = dictionary! {
            "Font" => dictionary! { "F1" => font_id },
        };
        let image_id = if with_image {
            let mut encoded = Vec::new();
            jpeg_encoder::Encoder::new(&mut encoded, 90)
                .encode(
                    &[20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240],
                    2,
                    2,
                    jpeg_encoder::ColorType::Rgb,
                )
                .unwrap();
            let id = document.add_object(Stream::new(
                dictionary! {
                    "Type" => "XObject",
                    "Subtype" => "Image",
                    "Width" => 2,
                    "Height" => 2,
                    "ColorSpace" => "DeviceRGB",
                    "BitsPerComponent" => 8,
                    "Filter" => "DCTDecode",
                },
                encoded,
            ));
            resources.set("XObject", dictionary! { "Im0" => id });
            Some(id)
        } else {
            None
        };
        let mut text = Stream::new(
            Dictionary::new(),
            b"BT /F1 12 Tf 30 150 Td (x = 1) Tj ET\n".to_vec(),
        );
        text.compress().unwrap();
        let text_id = document.add_object(text);
        let mut drawing = b"10 10 100 50 re S\n".to_vec();
        if with_image {
            drawing.extend_from_slice(b"q\n40 0 0 20 40 50 cm\n/Im0 Do\nQ\n");
        }
        let mut drawing = Stream::new(Dictionary::new(), drawing);
        drawing.compress().unwrap();
        let drawing_id = document.add_object(drawing);
        let contents_id = document.add_object(vec![Object::Reference(text_id), drawing_id.into()]);
        let annotation_id = document.add_object(dictionary! {
            "Type" => "Annot",
            "Subtype" => "Text",
            "Rect" => vec![10.into(), 20.into(), 30.into(), 40.into()],
            "Contents" => Object::string_literal("annotation"),
        });
        let page_id = document.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => contents_id,
            "Annots" => vec![annotation_id.into()],
            "BleedBox" => vec![10.into(), 20.into(), 310.into(), 220.into()],
            "TrimBox" => vec![10.into(), 20.into(), 310.into(), 220.into()],
            "ArtBox" => vec![10.into(), 20.into(), 310.into(), 220.into()],
        });
        document.objects.insert(
            pages_id,
            dictionary! {
                "Type" => "Pages",
                "Kids" => vec![page_id.into()],
                "Count" => 1,
                "MediaBox" => vec![10.into(), 20.into(), 310.into(), 220.into()],
                "Resources" => resources,
            }
            .into(),
        );
        let catalog_id = document.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        document.trailer.set("Root", catalog_id);
        Self {
            document,
            page_id,
            pages_id,
            content_ids: vec![text_id, drawing_id],
            font_id,
            image_id,
        }
    }

    fn write(&mut self, path: &Path) {
        let mut bytes = Vec::new();
        self.document.save_to(&mut bytes).unwrap();
        fs::write(path, bytes).unwrap();
    }
}

fn request(input: &Path, output: &Path) -> Request {
    Request {
        input_path: input.to_string_lossy().into_owned(),
        pdf_crop: Some(PdfCrop {
            left_pt: 20.0,
            top_pt: 30.0,
            width_pt: 100.0,
            height_pt: 50.0,
        }),
        output_path: output.to_string_lossy().into_owned(),
        format: OutputFormat::Pdf,
        source_width_pt: 200.0,
        source_height_pt: 200.0,
        size_mode: SizeMode::Keep,
        width: None,
        height: None,
        unit: Unit::Inch,
        percent: None,
        dpi: 1_000,
        transparent: false,
        quality: 95,
        overwrite: false,
    }
}

fn read_pdf(path: &Path) -> Document {
    Document::load_mem(&fs::read(path).unwrap()).unwrap()
}

fn wrapper_operations(document: &Document, page_id: ObjectId) -> Vec<Operation> {
    let contents = document.get_page_contents(page_id);
    let stream = document
        .get_object(contents[0])
        .unwrap()
        .as_stream()
        .unwrap();
    Content::decode(&stream.content).unwrap().operations
}

fn assert_numbers(actual: &[Object], expected: &[f64]) {
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(expected) {
        let actual = f64::from(actual.as_float().unwrap());
        assert!((actual - expected).abs() < 1e-5, "{actual} != {expected}");
    }
}

fn image_count(document: &Document) -> usize {
    document
        .objects
        .values()
        .filter(|object| {
            object
                .as_stream()
                .ok()
                .and_then(|stream| stream.dict.get(b"Subtype").ok())
                .and_then(|subtype| subtype.as_name().ok())
                == Some(b"Image".as_slice())
        })
        .count()
}

#[test]
fn native_pdf_retains_vectors_and_centers_the_crop_with_inherited_media_box() {
    let directory = TestDirectory::new();
    let input = directory.join("native.data");
    let output = directory.join("cropped.pdf");
    let mut fixture = Fixture::new(false);
    fixture.write(&input);
    let response = process_request(&request(&input, &output)).unwrap();
    assert_eq!((response.width_pt, response.height_pt), (200.0, 200.0));
    assert_eq!(response.pixel_width, None);
    assert_eq!(response.pixel_height, None);
    let result = read_pdf(&output);
    let page = result.get_dictionary(fixture.page_id).unwrap();
    for key in [b"MediaBox".as_slice(), b"CropBox".as_slice()] {
        assert_numbers(
            page.get(key).unwrap().as_array().unwrap(),
            &[0.0, 0.0, 200.0, 200.0],
        );
    }
    for key in [
        b"Annots".as_slice(),
        b"BleedBox".as_slice(),
        b"TrimBox".as_slice(),
        b"ArtBox".as_slice(),
    ] {
        assert!(page.get(key).is_err());
    }
    let operations = wrapper_operations(&result, fixture.page_id);
    let operators: Vec<_> = operations
        .iter()
        .map(|operation| operation.operator.as_str())
        .collect();
    assert_eq!(operators, vec!["q", "re", "W", "n", "cm"]);
    assert_numbers(&operations[1].operands, &[0.0, 50.0, 200.0, 100.0]);
    assert_numbers(
        &operations[4].operands,
        &[2.0, 0.0, 0.0, 2.0, -60.0, -230.0],
    );
    let result_contents = result.get_page_contents(fixture.page_id);
    assert_eq!(&result_contents[1..3], fixture.content_ids.as_slice());
    for id in &fixture.content_ids {
        let original = fixture
            .document
            .get_object(*id)
            .unwrap()
            .as_stream()
            .unwrap();
        let retained = result.get_object(*id).unwrap().as_stream().unwrap();
        assert_eq!(retained.content, original.content);
    }
    assert_eq!(
        result.get_object(fixture.font_id).unwrap(),
        fixture.document.get_object(fixture.font_id).unwrap()
    );
    let content = Content::decode(&result.get_page_content(fixture.page_id)).unwrap();
    for operator in ["Tf", "Tj", "re", "S"] {
        assert!(content
            .operations
            .iter()
            .any(|operation| operation.operator == operator));
    }
    assert_eq!(image_count(&result), 0);
    let suffix = result
        .get_object(*result_contents.last().unwrap())
        .unwrap()
        .as_stream()
        .unwrap();
    assert_eq!(suffix.content, b"\nQ\n");
}

#[test]
fn embedded_image_stream_is_preserved_byte_for_byte() {
    let directory = TestDirectory::new();
    let input = directory.join("image.pdf");
    let output = directory.join("cropped.pdf");
    let mut fixture = Fixture::new(true);
    fixture.write(&input);
    process_request(&request(&input, &output)).unwrap();
    let result = read_pdf(&output);
    let image_id = fixture.image_id.unwrap();
    let original = fixture
        .document
        .get_object(image_id)
        .unwrap()
        .as_stream()
        .unwrap();
    let retained = result.get_object(image_id).unwrap().as_stream().unwrap();
    assert_eq!(retained.content, original.content);
    assert_eq!(retained.dict, original.dict);
    assert_eq!(image_count(&result), 1);
}

#[test]
fn custom_pdf_size_scales_axes_independently() {
    let directory = TestDirectory::new();
    let input = directory.join("native.pdf");
    let output = directory.join("custom.pdf");
    let mut fixture = Fixture::new(false);
    fixture.write(&input);
    let mut request = request(&input, &output);
    request.size_mode = SizeMode::Custom;
    request.width = Some(200.0 / 72.0);
    request.height = Some(300.0 / 72.0);
    process_request(&request).unwrap();
    let result = read_pdf(&output);
    let operations = wrapper_operations(&result, fixture.page_id);
    assert_numbers(&operations[1].operands, &[0.0, 0.0, 200.0, 300.0]);
    assert_numbers(
        &operations[4].operands,
        &[2.0, 0.0, 0.0, 6.0, -60.0, -840.0],
    );
    let page = result.get_dictionary(fixture.page_id).unwrap();
    assert_numbers(
        page.get(b"MediaBox").unwrap().as_array().unwrap(),
        &[0.0, 0.0, 200.0, 300.0],
    );
}

#[test]
fn absent_pdf_crop_uses_the_complete_media_box_and_single_content_reference() {
    let directory = TestDirectory::new();
    let input = directory.join("native.pdf");
    let output = directory.join("complete.pdf");
    let mut fixture = Fixture::new(false);
    fixture
        .document
        .get_dictionary_mut(fixture.page_id)
        .unwrap()
        .set("Contents", fixture.content_ids[0]);
    fixture.write(&input);
    let mut request = request(&input, &output);
    request.pdf_crop = None;
    request.source_width_pt = 300.0;
    request.source_height_pt = 200.0;
    process_request(&request).unwrap();
    let result = read_pdf(&output);
    let operations = wrapper_operations(&result, fixture.page_id);
    assert_numbers(&operations[1].operands, &[0.0, 0.0, 300.0, 200.0]);
    assert_numbers(&operations[4].operands, &[1.0, 0.0, 0.0, 1.0, -10.0, -20.0]);
    assert_eq!(result.get_page_contents(fixture.page_id).len(), 3);
}

#[test]
fn invalid_pdf_crop_coordinates_and_dimensions_are_rejected() {
    let directory = TestDirectory::new();
    let input = directory.join("native.pdf");
    let mut fixture = Fixture::new(false);
    fixture.write(&input);
    let crops = [
        (-1.0, 0.0, 10.0, 10.0),
        (0.0, -1.0, 10.0, 10.0),
        (0.0, 0.0, 0.0, 10.0),
        (0.0, 0.0, 10.0, -1.0),
        (f64::NAN, 0.0, 10.0, 10.0),
        (0.0, 0.0, f64::INFINITY, 10.0),
        (250.0, 0.0, 100.0, 10.0),
        (0.0, 180.0, 10.0, 50.0),
    ];
    for (index, (left_pt, top_pt, width_pt, height_pt)) in crops.into_iter().enumerate() {
        let output = directory.join(&format!("invalid-{index}.pdf"));
        let mut request = request(&input, &output);
        request.pdf_crop = Some(PdfCrop {
            left_pt,
            top_pt,
            width_pt,
            height_pt,
        });
        let error = process_request(&request).unwrap_err();
        assert!(error.to_string().contains("pdf_crop"), "{error}");
        assert!(!output.exists());
    }
}

#[test]
fn multi_page_rotated_and_nonstandard_unit_pdf_inputs_are_rejected() {
    let directory = TestDirectory::new();
    for (case, message) in [
        ("multi", "仅包含一页"),
        ("rotated", "Rotate"),
        ("unit", "UserUnit"),
    ] {
        let input = directory.join(&format!("{case}.pdf"));
        let output = directory.join(&format!("{case}-output.pdf"));
        let mut fixture = Fixture::new(false);
        match case {
            "multi" => {
                let second = fixture
                    .document
                    .get_dictionary(fixture.page_id)
                    .unwrap()
                    .clone();
                let second_id = fixture.document.add_object(second);
                let pages = fixture
                    .document
                    .get_dictionary_mut(fixture.pages_id)
                    .unwrap();
                pages.set(
                    "Kids",
                    vec![Object::Reference(fixture.page_id), second_id.into()],
                );
                pages.set("Count", 2);
            }
            "rotated" => fixture
                .document
                .get_dictionary_mut(fixture.pages_id)
                .unwrap()
                .set("Rotate", 90),
            "unit" => fixture
                .document
                .get_dictionary_mut(fixture.page_id)
                .unwrap()
                .set("UserUnit", 2),
            _ => unreachable!(),
        }
        fixture.write(&input);
        let error = process_request(&request(&input, &output)).unwrap_err();
        assert!(error.to_string().contains(message), "{error}");
        assert!(!output.exists());
    }
}

#[test]
fn pdf_input_rejects_every_non_pdf_output_format() {
    let directory = TestDirectory::new();
    let input = directory.join("native.pdf");
    let mut fixture = Fixture::new(false);
    fixture.write(&input);
    for (index, format) in [
        OutputFormat::Svg,
        OutputFormat::Png,
        OutputFormat::Tiff,
        OutputFormat::Jpeg,
        OutputFormat::Webp,
    ]
    .into_iter()
    .enumerate()
    {
        let output = directory.join(&format!("invalid-{index}.data"));
        let mut request = request(&input, &output);
        request.format = format;
        let error = process_request(&request).unwrap_err();
        assert!(
            error.to_string().contains("PDF 输入仅支持 PDF 输出"),
            "{error}"
        );
        assert!(!output.exists());
    }
}
