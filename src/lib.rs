mod test;

use std::time::Instant;
use calamine::{ open_workbook, Reader, Xlsx};
use rayon::iter::{IntoParallelRefIterator, ParallelIterator};
use xlsxwriter::prelude::*;
use pyo3::prelude::*;
use unicode_segmentation::UnicodeSegmentation;


#[pyfunction]
pub fn check_bad_val(value: &str) -> bool {
    (value != "") && (value != "Nil") && (value != "--") && (value.to_ascii_lowercase() != "nan") && (value != "NMF")
}

#[pyfunction]
pub fn contains_character(input: &str) -> bool {
    for character in input.graphemes(true) {
        if character.parse::<u8>().is_err() {
            return true;
        }
    }
    false
}
#[pyfunction]
fn contains_number(input: &str) -> bool {
    for character in input.graphemes(true) {
        if character.parse::<u8>().is_ok() {
            return true;
        }
    }
    false
}

#[pyfunction]
fn extract_value(input: &str) -> (String, String, String) {
    let mut prefix: String = "".to_string();
    let mut postfix: String = "".to_string();
    let mut value: String = "".to_string();

    if input.parse::<f32>().is_ok() {
        value = input.to_string();
    } else {
        let characters = input.graphemes(true);
        for char in characters.clone() {
            if (char.parse::<f32>().is_ok()) || (char == ".".to_string()) {
                break;
            } else {
                prefix += char;
            }
        }
        for char in characters.rev() {
            if (char.parse::<f32>().is_ok()) || (char == ".".to_string()) {
                break;
            } else {
                postfix = char.to_string() + &postfix;
            }
        }

        // assert!(prefix.len() <= (input.len()-postfix.len()));
        // println!("{} {} {}", input, input.len(), postfix.len());
        value = input[prefix.len()..(input.len() - postfix.len())].to_string();
    }

    (value, prefix, postfix)
}

#[pyfunction]
fn count_postfix(input: &str, character: char) -> u32 {
    let mut zero_counter = 0u32;
    let characters = input.graphemes(true);
    for char in characters.rev() {
        if char == character.to_string() {
            zero_counter += 1;
        } else {
            return zero_counter;
        }
    }
    zero_counter
}

#[pyfunction]
fn reg_replace_common(input: String) -> String {
    if input.to_lowercase().contains("nme") {
        return "NMF".to_string();
    }
    let output = input.replace(",", ".");

   output
}


#[pyfunction]
pub fn process(fp: &str, input_dir: &str, output_dir: &str) -> PyResult<()> {
    let output_workbook = Workbook::new(&(output_dir.to_string() + fp)).unwrap();
    let mut workbook: Xlsx<_> = open_workbook(&(input_dir.to_string() + fp)).unwrap();
    let names = workbook.sheet_names();
    for sheet_name in names {
        let input_sheet = workbook.worksheet_range(&sheet_name).unwrap().unwrap(); // get sheet contents
        let mut output_sheet = output_workbook.add_worksheet(Some(&sheet_name)).unwrap(); // create sheet inside output file

        if sheet_name == "Table_0" {
            for y in 0usize..24{

                // does value from 'labels' column contain "sales"
                let label;
                if let Some(_val) = input_sheet.get_value((y as u32, 8u32)) {
                    label = _val.to_string();
                } else {
                    label = "".to_string();
                }
                if !((label.contains("sales")) || (label.contains("Sales"))) {

                    // check if column 14 has a decimal
                    let col14_cell;
                    if let Some(label) = input_sheet.get_value((y as u32, 9u32)) {
                        col14_cell = label.to_string();
                    } else {
                        col14_cell = "".to_string()
                    }
                    if check_bad_val(&col14_cell) {

                        // determine whether column 14 has a decimal point + approximate size
                        if contains_number(&col14_cell) {

                            let (col14_value, _, _) = extract_value(&col14_cell);
                            if col14_value.parse::<f32>().is_ok() {
                                let reference_base = f32::log10(col14_value.parse::<f32>().unwrap());

                                for x in 0..8u16 {
                                    // println!("{cell_value}");
                                    let cell_value;
                                    if let Some(label) = input_sheet.get_value((y as u32, x as u32)) {
                                        cell_value = reg_replace_common(label.to_string());
                                    } else {
                                        cell_value = "".to_string()
                                    }
                                    if !(cell_value.contains(".")) && (!cell_value.is_empty()) && contains_number(&cell_value) {
                                        let (value_tmp, prefix, postfix) = extract_value(&cell_value);
                                        // ensure numerical only
                                        if let Ok(input_value) = value_tmp.clone().parse::<f32>() {
                                            let value_base = f32::log10(input_value);

                                            let a = (reference_base - value_base).floor();
                                            let mut output_value = input_value.to_string();
                                            if a <= 0f32 {
                                                output_value = (input_value.round() * 10f32.powf(a)).to_string();
                                            }
                                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &(prefix + &output_value + &postfix), Some(Format::new().set_bg_color(FormatColor::Orange))).unwrap();
                                        } else {
                                            // inner val unparsable -> copy raw cell val
                                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, Some(Format::new().set_bg_color(FormatColor::Orange))).unwrap();
                                        }
                                    } else {
                                        // execute if cell contains 'special value' ∊ NULL, nill, NMF, etc
                                        output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                                    }
                                }

                                for x in 8u16..10u16 {
                                    // copies last two columns because std code cannot be applied
                                    let cell_value;
                                    if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                                    } else { cell_value = "".to_string() }
                                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                                }
                            } else {
                                // if 'last' column cannot be parsed for number
                                for x in 0u16..10u16 {
                                    let cell_value;
                                    if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                                    } else { cell_value = "".to_string() }
                                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                                }
                            }
                        } else {
                            for x in 0u16..10u16 {
                                // let cell_value = get_cell_value(y, x as usize, &input_sheet);
                                let cell_value;
                                if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                                } else { cell_value = "".to_string() }
                                output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                            }
                        }

                    } else {
                        // generic copy if
                        for x in 0u16..10u16 {
                            let cell_value;
                            if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                            } else { cell_value = "".to_string() }
                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                        }
                    }
                } else {
                    // generic copy (if label 'Sales')
                    for x in 0u16..10u16 {
                        let cell_value;
                        if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                        } else { cell_value = "".to_string() }
                        output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                    }
                }
            }
        } else {
            // generically copies table if NOT "Table_0"
            for y in 0..input_sheet.height(){
                for x in 0u16..input_sheet.width() as u16 {
                    let cell_value;
                    if let Some(label) = input_sheet.get_value((y as u32, x as u32)) { cell_value = label.to_string();
                    } else { cell_value = "".to_string() }
                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None).unwrap();
                }
            }
        }
    }
    Ok(())
}

#[pyfunction]
pub fn process_all(input_dir: &str, output_dir: &str, sequential: bool) -> PyResult<()> {
    let mut files: Vec<String> = vec![];
    for file in std::fs::read_dir(input_dir).unwrap() {
        files.push( file.unwrap().file_name().to_str().unwrap().to_string() )
    }
    let start_time = Instant::now();

    if sequential {
        for (index, file) in files.iter().enumerate() {
            process(&file, input_dir, output_dir).unwrap();
            println!("{file} ({}/{})", index+1, files.len())
        }

    } else {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(4) // Set the number of threads in the pool
            .build()
            .unwrap();


        let _ = pool.install(|| {
            files.par_iter() // Use parallel iterator
                .map(|item| process(item, input_dir, output_dir).unwrap()) // Replace with your computation
        });
    }

    println!("{:#?}", Instant::now() - start_time);
    Ok(())
}

#[pymodule]
fn decimal_processing_xlsx_vesmar(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(process, m)?)?;
    m.add_function(wrap_pyfunction!(process_all, m)?)?;
    Ok(())
}