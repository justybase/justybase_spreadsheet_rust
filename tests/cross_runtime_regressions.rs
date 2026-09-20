use chrono::NaiveDate;
use spreadsheet::{
    CellValue, SheetOptions, XlsbReader, XlsbSheetOptions, XlsbUpdater, XlsbWriter, XlsmUpdater,
    XlsxReader, XlsxUpdater, XlsxWriter, F,
};
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

fn date() -> chrono::NaiveDateTime {
    NaiveDate::from_ymd_opt(2024, 2, 29)
        .unwrap()
        .and_hms_opt(10, 20, 30)
        .unwrap()
}

fn headers() -> Vec<String> {
    [
        "Text",
        "Integer",
        "Number",
        "Boolean",
        "Date",
        "Custom date",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect()
}

fn rows() -> Vec<Vec<CellValue>> {
    vec![
        vec![
            CellValue::Text("żółć & <tag> 😀\n".into()),
            CellValue::Integer(-536_870_912),
            CellValue::Number(-0.001),
            CellValue::Boolean(false),
            CellValue::DateTime(date()),
            CellValue::formatted(CellValue::DateTime(date()), F::DATE_ISO),
        ],
        vec![
            CellValue::Text("  spaced  ".into()),
            CellValue::Integer(536_870_911),
            CellValue::Number(1.7976931348623157e20),
            CellValue::Boolean(true),
            CellValue::DateTime(date()),
            CellValue::formatted(CellValue::DateTime(date()), F::DATETIME_ISO),
        ],
    ]
}

fn write_workbook(path: &Path, xlsb: bool, streaming: bool) {
    let header = headers();
    let data = rows();
    if xlsb {
        let mut writer = XlsbWriter::create(path).unwrap();
        if streaming {
            writer
                .start_sheet("Data", header.len(), Some(&header), XlsbSheetOptions::new())
                .unwrap();
            for row in &data {
                writer.write_row(row).unwrap();
            }
            writer.end_sheet().unwrap();
        } else {
            writer.add_sheet("Data", false);
            writer.write_sheet(data, Some(&header), true).unwrap();
        }
        writer.add_sheet("Other", true);
        writer
            .write_sheet(
                vec![vec![CellValue::Text("untouched".into())]],
                Some(&["Value".into()]),
                false,
            )
            .unwrap();
        writer.finalize().unwrap();
    } else {
        let mut writer = XlsxWriter::create(path).unwrap();
        if streaming {
            writer
                .start_sheet("Data", header.len(), Some(&header), SheetOptions::new())
                .unwrap();
            for row in &data {
                writer.write_row(row).unwrap();
            }
            writer.end_sheet().unwrap();
        } else {
            writer.add_sheet("Data", false);
            writer.write_sheet(data, Some(&header), true).unwrap();
        }
        writer.add_sheet("Other", true);
        writer
            .write_sheet(
                vec![vec![CellValue::Text("untouched".into())]],
                Some(&["Value".into()]),
                false,
            )
            .unwrap();
        writer.finalize().unwrap();
    }
}

fn assert_data_rows(rows: &[Vec<CellValue>]) {
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[0][0], CellValue::Text("Text".into()));
    assert_eq!(rows[1][0], CellValue::Text("żółć & <tag> 😀\n".into()));
    assert_eq!(rows[1][1], CellValue::Number(-536_870_912.0));
    assert_eq!(rows[1][2], CellValue::Number(-0.001));
    assert_eq!(rows[1][3], CellValue::Boolean(false));
    assert_eq!(rows[1][4], CellValue::DateTime(date()));
    assert_eq!(rows[1][5], CellValue::DateTime(date()));
}

fn read_xlsx(path: &Path, sheet: &str) -> Vec<Vec<CellValue>> {
    let mut reader = XlsxReader::new();
    reader.open(path, true).unwrap();
    reader.select_sheet(sheet).unwrap();
    let mut rows = Vec::new();
    while reader.read().unwrap() {
        rows.push(reader.current_row().to_vec());
    }
    rows
}

fn read_xlsb(path: &Path, sheet: &str) -> Vec<Vec<CellValue>> {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    reader.select_sheet(sheet).unwrap();
    let mut rows = Vec::new();
    while reader.read().unwrap() {
        rows.push(reader.current_row().to_vec());
    }
    rows
}

#[test]
fn writers_and_readers_cover_types_unicode_and_all_sheets() {
    let dir = tempfile::tempdir().unwrap();
    for (extension, xlsb) in [("xlsx", false), ("xlsb", true)] {
        for streaming in [false, true] {
            let path = dir.path().join(format!(
                "types-{}-{streaming}.{extension}",
                if xlsb { "b" } else { "x" }
            ));
            write_workbook(&path, xlsb, streaming);
            let data = if xlsb {
                read_xlsb(&path, "Data")
            } else {
                read_xlsx(&path, "Data")
            };
            assert_data_rows(&data);
            let other = if xlsb {
                read_xlsb(&path, "Other")
            } else {
                read_xlsx(&path, "Other")
            };
            assert_eq!(other[1][0], CellValue::Text("untouched".into()));
        }
    }
}

fn prefix_xml(name: &str, mut xml: String) -> String {
    if name == "xl/workbook.xml" {
        xml = xml.replace("<workbook ", "<x:workbook xmlns:x=\"urn:test-main\" ");
        xml = xml.replace("</workbook>", "</x:workbook>");
        xml = xml.replace("<workbookPr ", "<x:workbookPr ");
        xml = xml.replace("<sheets>", "<x:sheets>");
        xml = xml.replace("</sheets>", "</x:sheets>");
        xml = xml.replace("<sheet ", "<x:sheet ");
    } else if name == "xl/_rels/workbook.xml.rels" {
        xml = xml.replace(
            "<Relationships ",
            "<x:Relationships xmlns:x=\"urn:test-rels\" ",
        );
        xml = xml.replace("</Relationships>", "</x:Relationships>");
        xml = xml.replace("<Relationship ", "<x:Relationship ");
    } else if name == "xl/styles.xml" {
        xml = xml.replace("<styleSheet ", "<x:styleSheet xmlns:x=\"urn:test-style\" ");
        xml = xml.replace("</styleSheet>", "</x:styleSheet>");
        for tag in ["numFmt", "cellXfs", "xf"] {
            xml = xml.replace(&format!("<{tag} "), &format!("<x:{tag} "));
            xml = xml.replace(&format!("</{tag}>"), &format!("</x:{tag}>"));
        }
    } else if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
        xml = xml.replace("<worksheet ", "<x:worksheet xmlns:x=\"urn:test-sheet\" ");
        xml = xml.replace("</worksheet>", "</x:worksheet>");
        for tag in ["sheetData", "row", "c", "v", "t"] {
            xml = xml.replace(&format!("<{tag} "), &format!("<x:{tag} "));
            xml = xml.replace(&format!("<{tag}>"), &format!("<x:{tag}>"));
            xml = xml.replace(&format!("</{tag}>"), &format!("</x:{tag}>"));
        }
    } else if name == "xl/sharedStrings.xml" {
        xml = xml.replace("<sst ", "<x:sst xmlns:x=\"urn:test-sst\" ");
        xml = xml.replace("</sst>", "</x:sst>");
        for tag in ["si", "t"] {
            xml = xml.replace(&format!("<{tag} "), &format!("<x:{tag} "));
            xml = xml.replace(&format!("<{tag}>"), &format!("<x:{tag}>"));
            xml = xml.replace(&format!("</{tag}>"), &format!("</x:{tag}>"));
        }
    }
    xml
}

fn rewrite_zip(source: &Path, target: &Path, extra: Option<(&str, &[u8])>) {
    let source_file = File::open(source).unwrap();
    let mut archive = zip::ZipArchive::new(source_file).unwrap();
    let target_file = File::create(target).unwrap();
    let mut writer = zip::ZipWriter::new(target_file);
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    let mut has_extra = false;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().to_owned();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        drop(entry);
        if name == extra.map(|(name, _)| name).unwrap_or("") {
            has_extra = true;
        }
        writer.start_file(&name, options).unwrap();
        if name.ends_with(".xml") {
            let xml = String::from_utf8(bytes).unwrap();
            writer.write_all(prefix_xml(&name, xml).as_bytes()).unwrap();
        } else {
            writer.write_all(&bytes).unwrap();
        }
    }
    if let Some((name, bytes)) = extra {
        if !has_extra {
            writer.start_file(name, options).unwrap();
            writer.write_all(bytes).unwrap();
        }
    }
    writer.finish().unwrap();
}

#[test]
fn xlsx_reader_and_updater_accept_namespace_prefixes() {
    let dir = tempfile::tempdir().unwrap();
    let source = dir.path().join("source.xlsx");
    let prefixed = dir.path().join("prefixed.xlsx");
    let updated = dir.path().join("updated.xlsx");
    write_workbook(&source, false, false);
    rewrite_zip(&source, &prefixed, None);

    assert_data_rows(&read_xlsx(&prefixed, "Data"));
    let mut updater = XlsxUpdater::open(&prefixed).unwrap();
    assert_eq!(updater.sheet_names(), vec!["Data", "Other"]);
    updater
        .replace_sheet_data(
            "Data",
            vec![vec![CellValue::Text("prefix-safe".into())]],
            None,
        )
        .unwrap();
    updater.save(Some(&updated)).unwrap();
    let rows = read_xlsx(&updated, "Data");
    assert_eq!(rows[0][0], CellValue::Text("prefix-safe".into()));
}

#[test]
fn streaming_updaters_roll_back_bad_rows_and_can_be_reused() {
    let dir = tempfile::tempdir().unwrap();
    for (extension, xlsb) in [("xlsx", false), ("xlsb", true)] {
        let source = dir.path().join(format!("rollback.{extension}"));
        let output = dir.path().join(format!("rollback-out.{extension}"));
        write_workbook(&source, xlsb, false);
        if xlsb {
            let mut updater = XlsbUpdater::open(&source).unwrap();
            updater
                .replace_sheet_data_stream(
                    "Data",
                    vec![vec![CellValue::Text("staged-before-error".into()); 6]].into_iter(),
                    None,
                )
                .unwrap();
            assert!(updater
                .replace_sheet_data_fallible_stream(
                    "Data",
                    vec![
                        Ok(vec![CellValue::Text("valid".into()); 6]),
                        Err("generator failed"),
                    ]
                    .into_iter(),
                    None,
                )
                .is_err());
            updater
                .replace_sheet_data_stream(
                    "Data",
                    vec![vec![CellValue::Text("reused".into()); 6]].into_iter(),
                    None,
                )
                .unwrap();
            updater.save_streaming(Some(&output)).unwrap();
            assert_eq!(
                read_xlsb(&output, "Data")[0][0],
                CellValue::Text("reused".into())
            );
        } else {
            let mut updater = XlsxUpdater::open(&source).unwrap();
            updater
                .replace_sheet_data_stream(
                    "Data",
                    vec![vec![CellValue::Text("staged-before-error".into()); 6]].into_iter(),
                    None,
                )
                .unwrap();
            assert!(updater
                .replace_sheet_data_fallible_stream(
                    "Data",
                    vec![
                        Ok(vec![CellValue::Text("valid".into()); 6]),
                        Err("generator failed"),
                    ]
                    .into_iter(),
                    None,
                )
                .is_err());
            updater
                .replace_sheet_data_stream(
                    "Data",
                    vec![vec![CellValue::Text("reused".into()); 6]].into_iter(),
                    None,
                )
                .unwrap();
            updater.save_streaming(Some(&output)).unwrap();
            assert_eq!(
                read_xlsx(&output, "Data")[0][0],
                CellValue::Text("reused".into())
            );
        }
    }
}

#[test]
fn xlsm_updater_preserves_opaque_vba_payload() {
    let dir = tempfile::tempdir().unwrap();
    let source = dir.path().join("macro.xlsm");
    let output = dir.path().join("macro-out.xlsm");
    let payload = b"synthetic-vba\0\xff\x10";
    let plain = dir.path().join("plain.xlsx");
    write_workbook(&plain, false, false);
    rewrite_zip(&plain, &source, Some(("xl/vbaProject.bin", payload)));

    let mut updater = XlsmUpdater::open(&source).unwrap();
    updater
        .replace_sheet_data(
            "Data",
            vec![vec![CellValue::Text("macro-safe".into())]],
            None,
        )
        .unwrap();
    updater.save(Some(&output)).unwrap();

    let file = File::open(&output).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut vba = Vec::new();
    archive
        .by_name("xl/vbaProject.bin")
        .unwrap()
        .read_to_end(&mut vba)
        .unwrap();
    assert_eq!(vba, payload);
}
