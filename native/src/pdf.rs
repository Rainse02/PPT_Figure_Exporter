use std::collections::HashSet;

use lopdf::content::{Content, Operation};
use lopdf::{Dictionary, Document, Object, ObjectId, Stream};

use crate::{CoreError, PdfCrop, TargetSize};

pub(crate) fn crop_and_resize(
    source: &[u8],
    crop: Option<PdfCrop>,
    target: TargetSize,
) -> Result<Vec<u8>, CoreError> {
    let mut document = Document::load_mem(source)
        .map_err(|error| CoreError::Input(format!("PDF 解析失败：{error}")))?;
    if document.is_encrypted() {
        return Err(CoreError::Input("PDF 输入不支持加密文档".to_owned()));
    }
    let pages = document.get_pages();
    if pages.len() != 1 {
        return Err(CoreError::Input(format!(
            "PDF 输入必须仅包含一页，实际为 {} 页",
            pages.len()
        )));
    }
    let page_id = *pages.values().next().unwrap();
    validate_page_units(&document, page_id)?;
    let bounds = media_box(&document, page_id)?;
    let page_width = bounds[2] - bounds[0];
    let page_height = bounds[3] - bounds[1];
    let crop = crop.unwrap_or(PdfCrop {
        left_pt: 0.0,
        top_pt: 0.0,
        width_pt: page_width,
        height_pt: page_height,
    });
    validate_crop(crop, page_width, page_height)?;

    let mut scale_x = target.width_pt / crop.width_pt;
    let mut scale_y = target.height_pt / crop.height_pt;
    if !target.custom_stretch {
        let scale = scale_x.min(scale_y);
        scale_x = scale;
        scale_y = scale;
    }
    if !scale_x.is_finite() || !scale_y.is_finite() || scale_x <= 0.0 || scale_y <= 0.0 {
        return Err(CoreError::InvalidRequest("PDF 缩放比例无效".to_owned()));
    }
    let output_width = crop.width_pt * scale_x;
    let output_height = crop.height_pt * scale_y;
    let offset_x = (target.width_pt - output_width) / 2.0;
    let offset_y = (target.height_pt - output_height) / 2.0;
    let crop_left = bounds[0] + crop.left_pt;
    let crop_bottom = bounds[3] - crop.top_pt - crop.height_pt;
    let translation_x = offset_x - crop_left * scale_x;
    let translation_y = offset_y - crop_bottom * scale_y;
    let prefix = Content {
        operations: vec![
            Operation::new("q", vec![]),
            Operation::new(
                "re",
                pdf_numbers(&[offset_x, offset_y, output_width, output_height])?,
            ),
            Operation::new("W", vec![]),
            Operation::new("n", vec![]),
            Operation::new(
                "cm",
                pdf_numbers(&[scale_x, 0.0, 0.0, scale_y, translation_x, translation_y])?,
            ),
        ],
    }
    .encode()
    .map_err(|error| CoreError::Encode(format!("PDF 裁剪变换写入失败：{error}")))?;
    let original_contents = page_contents(&document, page_id)?;
    let prefix_id = document.add_object(Stream::new(Dictionary::new(), prefix));
    let suffix_id = document.add_object(Stream::new(Dictionary::new(), b"\nQ\n".to_vec()));
    let mut contents = Vec::with_capacity(original_contents.len() + 2);
    contents.push(prefix_id.into());
    contents.extend(original_contents);
    contents.push(suffix_id.into());
    let output_box = pdf_numbers(&[0.0, 0.0, target.width_pt, target.height_pt])?;
    let page = document
        .get_dictionary_mut(page_id)
        .map_err(pdf_structure_error)?;
    page.set("Contents", contents);
    page.set("MediaBox", output_box.clone());
    page.set("CropBox", output_box);
    page.set("Rotate", 0);
    page.set("UserUnit", 1);
    // Annotation rectangles are not part of the content transformation.
    page.remove(b"Annots");
    page.remove(b"BleedBox");
    page.remove(b"TrimBox");
    page.remove(b"ArtBox");

    let mut output = Vec::new();
    document
        .save_to(&mut output)
        .map_err(|error| CoreError::Encode(format!("PDF 写出失败：{error}")))?;
    Ok(output)
}

fn pdf_structure_error(error: lopdf::Error) -> CoreError {
    CoreError::Input(format!("PDF 结构无效：{error}"))
}

fn inherited_attribute<'a>(
    document: &'a Document,
    page_id: ObjectId,
    name: &[u8],
) -> Result<Option<&'a Object>, CoreError> {
    let mut current = page_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current) {
            return Err(CoreError::Input("PDF 页面继承链存在循环".to_owned()));
        }
        let dictionary = document
            .get_dictionary(current)
            .map_err(pdf_structure_error)?;
        if let Ok(value) = dictionary.get(name) {
            return document
                .dereference(value)
                .map(|(_, value)| Some(value))
                .map_err(pdf_structure_error);
        }
        let Ok(parent) = dictionary.get(b"Parent") else {
            return Ok(None);
        };
        current = parent.as_reference().map_err(pdf_structure_error)?;
    }
}

fn numeric_value(document: &Document, value: &Object) -> Result<f64, CoreError> {
    let (_, value) = document.dereference(value).map_err(pdf_structure_error)?;
    let number = match value {
        Object::Integer(number) => *number as f64,
        Object::Real(number) => f64::from(*number),
        _ => return Err(CoreError::Input("PDF 页面数值类型无效".to_owned())),
    };
    if !number.is_finite() {
        return Err(CoreError::Input("PDF 页面数值必须有限".to_owned()));
    }
    Ok(number)
}

fn media_box(document: &Document, page_id: ObjectId) -> Result<[f64; 4], CoreError> {
    let value = inherited_attribute(document, page_id, b"MediaBox")?
        .ok_or_else(|| CoreError::Input("PDF 页面缺少 MediaBox".to_owned()))?;
    let values = value.as_array().map_err(pdf_structure_error)?;
    if values.len() != 4 {
        return Err(CoreError::Input("PDF MediaBox 必须包含四个数值".to_owned()));
    }
    let mut bounds = [0.0; 4];
    for (destination, value) in bounds.iter_mut().zip(values) {
        *destination = numeric_value(document, value)?;
    }
    if bounds[2] <= bounds[0] || bounds[3] <= bounds[1] {
        return Err(CoreError::Input("PDF MediaBox 尺寸必须为正数".to_owned()));
    }
    Ok(bounds)
}

fn validate_page_units(document: &Document, page_id: ObjectId) -> Result<(), CoreError> {
    if let Some(rotation) = inherited_attribute(document, page_id, b"Rotate")? {
        if numeric_value(document, rotation)? != 0.0 {
            return Err(CoreError::Input("PDF 输入不支持非零 Rotate".to_owned()));
        }
    }
    let page = document
        .get_dictionary(page_id)
        .map_err(pdf_structure_error)?;
    if let Ok(unit) = page.get(b"UserUnit") {
        if numeric_value(document, unit)? != 1.0 {
            return Err(CoreError::Input("PDF 输入仅支持 UserUnit=1".to_owned()));
        }
    }
    Ok(())
}

fn validate_crop(crop: PdfCrop, page_width: f64, page_height: f64) -> Result<(), CoreError> {
    let values = [crop.left_pt, crop.top_pt, crop.width_pt, crop.height_pt];
    if !values.iter().all(|value| value.is_finite())
        || crop.left_pt < 0.0
        || crop.top_pt < 0.0
        || crop.width_pt <= 0.0
        || crop.height_pt <= 0.0
        || crop.left_pt + crop.width_pt > page_width
        || crop.top_pt + crop.height_pt > page_height
    {
        return Err(CoreError::InvalidRequest(
            "pdf_crop 必须包含有限坐标和正尺寸，且完整位于 PDF MediaBox 内".to_owned(),
        ));
    }
    Ok(())
}

fn page_contents(document: &Document, page_id: ObjectId) -> Result<Vec<Object>, CoreError> {
    let page = document
        .get_dictionary(page_id)
        .map_err(pdf_structure_error)?;
    let Ok(value) = page.get(b"Contents") else {
        return Ok(vec![]);
    };
    let (_, resolved) = document.dereference(value).map_err(pdf_structure_error)?;
    let contents = match resolved {
        Object::Null => vec![],
        Object::Array(contents) => contents.clone(),
        Object::Stream(_) if matches!(value, Object::Reference(_)) => vec![value.clone()],
        _ => return Err(CoreError::Input("PDF Contents 必须引用内容流".to_owned())),
    };
    for content in &contents {
        let id = content.as_reference().map_err(pdf_structure_error)?;
        document
            .get_object(id)
            .and_then(Object::as_stream)
            .map_err(pdf_structure_error)?;
    }
    Ok(contents)
}

fn pdf_numbers(values: &[f64]) -> Result<Vec<Object>, CoreError> {
    values
        .iter()
        .map(|value| {
            let real = *value as f32;
            if !real.is_finite() || (real == 0.0 && *value != 0.0) {
                return Err(CoreError::InvalidRequest(
                    "PDF 裁剪或目标尺寸超出可表示范围".to_owned(),
                ));
            }
            Ok(Object::Real(real))
        })
        .collect()
}
