use clap::Parser;
use dirs;
use docx_rs::*;
use pdf_extract::extract_text;
use printpdf::*;
use serde_json::Value;
use serde_xml_rs::to_string;
use slicestring::Slice;
use std::error::Error;
use std::fs::{self, File};
use std::io::{BufWriter, Read, Write};
use std::path::Path;
use std::process::exit;
use quickxml_to_serde::{xml_string_to_json, Config, NullValue};
use csv;

type MyResultBox = Result<(), Box<dyn Error>>;

#[derive(Parser)]
#[command(author="Diyaan Pulikkal <66010991@kmitl.ac.th>", version="0.1.0", about="Document Conversion Tool", long_about = None)]
pub struct Instruction {
    /// Enter desired input file's full path [REQUIRED]
    #[arg(long, short)]
    pub input_path: Option<String>,

    /// Enter desired conversion format, e.g., txt, pdf, docx, etc. Leaving this blank will just make a copy of the input file.
    #[arg(long, short, default_value_t = String::from(""))]
    extension: String,

    /// Enter desired path for file conversion result. (Default path is downloads folder)
    #[arg(long, short, default_value_t = String::from(""))]
    output_path: String,

    /// Customize output file name only for single file input
    #[arg(long, short, default_value_t = String::from(""))]
    name_file: String,
}

struct ConversionInfo {
    input_path: String,
    input_extension: String,
    output_path: String,
    output_extension: String,
    name_file: String,
}

impl ConversionInfo {
    fn new(
        input_path: String,
        input_extension: String,
        output_path: String,
        output_extension: String,
        name_file: String,
    ) -> ConversionInfo {
        return ConversionInfo {
            input_path,
            input_extension,
            output_path,
            output_extension,
            name_file,
        };
    }
}

//----------------------------------------------------------------------------------------------------------------------------------------

//Format raw arguments to be understandable by the file
pub fn verify_args(instruction: Instruction){
    let input_path: String = instruction.input_path.unwrap();
    let mut output_extension: String = instruction.extension;
    let mut output_path: String = instruction.output_path;
    let mut name_file: String = instruction.name_file;
    let mut input_extension: String = "".to_string();

    // Check if the input path exists
    if fs::metadata(&input_path).is_err() {
        eprintln!("Input file not found.");
        exit(0);
    }

    // Check if output path exists
    if output_path != "".to_string() && fs::metadata(&output_path).is_err() {
        eprintln!("Output folder not found.");
        exit(0);
    }

    // If no output path was provided, direct the converted file to the download folder
    if output_path == "" {
        if let Some(download_dir) = dirs::download_dir() {
            output_path = download_dir.to_string_lossy().to_string();
        } else {
            eprintln!("Specify output folder (unable to find default Downloads folder).");
            exit(0);
        }
    }

    // Check for the extension of input file and also check if output extention is available, if not, then the default will be same as of the input file's.
    let file_path = Path::new(&input_path);
    if let Some(ext) = file_path.extension() {
        if let Some(ext_str) = ext.to_str() {
            if output_extension == "".to_string() {
                output_extension = ext_str.to_string();
            }
            input_extension = ext_str.to_string().clone();
        }
    }

    let avail_ext = [
        "txt".to_string(),
        "pdf".to_string(),
        "docx".to_string(),
        "csv".to_string(),
        "xlsx".to_string(),
        "json".to_string(),
        "xml".to_string(),
        "html".to_string(),
    ];

    if !avail_ext.contains(&output_extension) || !avail_ext.contains(&input_extension) {
        eprintln!("Invalid extension");
        exit(0);
    }
    drop(avail_ext);

    // See if name_file was entered
    if name_file == "".to_string() {
        let path_for_name = Path::new(&input_path);
        if let Some(path_name) = path_for_name.file_stem() {
            if let Some(str_name) = path_name.to_str() {
                if input_extension == output_extension {
                    name_file = format!("{}-copied", str_name.to_string());
                } else {
                    name_file = format!("{}-converted", str_name.to_string());
                }
            }
        }
    }

    // Construct the verified input for the next process
    let finalized_info = ConversionInfo::new(
        input_path,
        input_extension,
        output_path,
        output_extension,
        name_file,
    );

    if let Err(e) = direct_info(finalized_info) {
        eprintln!("{}", e);
    }
}

//-----------------------------------------------------------------------------------------------------------------------------------------

// Direct the processed argument to their respective conversion function
fn direct_info(info: ConversionInfo) -> MyResultBox {
    if info.input_extension == info.output_extension {
        if let Err(e) = duplicate_file(info) {
            eprintln!("{}", e)
        }
        return Ok(());
    }

    match (
        info.input_extension.as_str(),
        info.output_extension.as_str(),
    ) {
        ("txt", "pdf") => txt_to_pdf(info),
        ("pdf", "txt") => pdf_to_txt(info),
        ("json", "xml") => json_to_xml(info),
        ("xml", "json") => xml_to_json(info),
        ("csv", "html") => csv_to_html(info),
        ("txt", "docx") => txt_to_docx(info),
        _ => Ok(inconvertible_formats(info)),
    }
}

//------------------------------------------------------------------------------------------------------------------------------

fn inconvertible_formats(info: ConversionInfo) {
    println!(
        "Cannot convert {} to {}",
        info.input_extension, info.output_extension
    );
    println!("See possible conversions:");
    println!("txt => pdf");
    println!("txt => docx");
    println!("pdf(text only) => txt");
    println!("xml <=> json");
    println!("csv => html");
    
    exit(0)
}

//---------------------------------------------------------------------------------------------------------------------------

// Conversion Functions:

fn duplicate_file(info: ConversionInfo) -> MyResultBox {
    let out_str_path: String = format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    );
    fs::copy(info.input_path, out_str_path)?;

    Ok(())
}

fn txt_to_pdf(info: ConversionInfo) -> MyResultBox {
    // Get String from txt file
    let mut text: String = String::new();
    File::open(info.input_path)?.read_to_string(&mut text)?;

    // Create a document object
    let (doc, page1, layer1) = PdfDocument::new("", Mm(210.0), Mm(297.0), "Layer 1");
    let font = doc.add_builtin_font(BuiltinFont::TimesRoman)?;
    let mut layer = doc.get_page(page1).get_layer(layer1);

    // Divide the text into lines with 110 characters each
    let text_char: Vec<char> = text.chars().collect();
    let char_per_lines: usize = 110;
    let no_of_lines: usize = text_char.len() / char_per_lines;
    let mut count: f32 = 0.0;

    for i in 0..no_of_lines {
        layer.use_text(
            text.slice((char_per_lines * i)..(char_per_lines * (i + 1))),
            12.0,
            Mm(10.0),
            Mm(287.0 - (10.0 * count)),
            &font,
        );
        if i == no_of_lines - 1 {
            layer.use_text(
                text.slice((char_per_lines * (i + 1))..),
                12.0,
                Mm(10.0),
                Mm(287.0 - (10.0 * (count + 1.0))),
                &font,
            );
        }
        // If a page already contains 27 lines, a new page begins
        if count == 27.0 {
            let (page1, layer1) = doc.add_page(Mm(210.0), Mm(297.0), "Layer 1");
            layer = doc.get_page(page1).get_layer(layer1);
            count = 0.0;
        } else {
            count = count + 1.0;
        }
    }

    // Create the new pdf file and apply the object created
    let file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;
    let mut file = BufWriter::new(file);
    doc.save(&mut file)?;

    Ok(())
}

fn pdf_to_txt(info: ConversionInfo) -> MyResultBox {
    let text = extract_text(&info.input_path).unwrap_or("".to_string());

    let mut out_file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;

    out_file.write_all(text.as_bytes())?;

    Ok(())
}

fn json_to_xml(info: ConversionInfo) -> MyResultBox {
    let mut json_content = String::new();
    File::open(info.input_path)?.read_to_string(&mut json_content)?;

    let json_data: Value = serde_json::from_str(&json_content)?;
    let xml_content = to_string(&json_data)?;

    let mut xml_file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;
    xml_file.write_all(xml_content.as_bytes())?;

    Ok(())
}

fn xml_to_json(info: ConversionInfo) -> MyResultBox {
    let mut xml_data = String::new();
    File::open(info.input_path)?.read_to_string(&mut xml_data)?;

    let conf = Config::new_with_custom_values(true, "", "txt", NullValue::Null);
    let json = xml_string_to_json(xml_data.to_owned(), &conf);

    let mut json_file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;
    json_file.write_all(json.unwrap().to_string().as_bytes())?;

    Ok(())
}

fn txt_to_docx(info: ConversionInfo) -> MyResultBox {
    let mut txt_content = String::new();
    File::open(&info.input_path)?.read_to_string(&mut txt_content)?;

    let word_docx = Docx::new();
    let paragraph = Paragraph::new().add_run(Run::new().add_text(txt_content.clone()));

    let docx_file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;
    word_docx.add_paragraph(paragraph).build().pack(docx_file)?;

    Ok(())
}

fn csv_to_html(info: ConversionInfo) -> MyResultBox {
    // Set up html file
    let mut html_file = File::create(format!(
        "{}/{}.{}",
        info.output_path, info.name_file, info.output_extension
    ))?;
    html_file.write_all(format!("<!DOCTYPE html><html><body><table><tr>").as_bytes())?;

    // Construct csv reader
    let mut csv_file = csv::Reader::from_path(info.input_path)?;

    // Make column headers
    let headers = csv_file.headers()?.clone();
    let vec_header: Vec<String> = headers.iter().map(|s| s.to_string()).collect();
    for i in vec_header {
        html_file.write_all(format!("<th>{}</th>", i).as_bytes())?;
    }
    html_file.write_all("</tr>".as_bytes())?;

    // Insert all data
    for i in csv_file.records() {
        let vec_data: Vec<String> = i?.iter().map(|s| s.to_string()).collect();
        html_file.write_all("<tr>".as_bytes())?;
        for j in vec_data {
            html_file.write_all(format!("<td>{}</td>", j).as_bytes())?
        }
        html_file.write_all("</tr>".as_bytes())?;
    }

    // Finish up the html structure
    html_file.write_all(format!("</table></body></html>").as_bytes())?;

    Ok(())
}
