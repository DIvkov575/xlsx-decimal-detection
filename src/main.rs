use std::error::Error;
use std::time::Instant;
use calamine::{ open_workbook, Reader, Xlsx};
use rayon::iter::{IntoParallelRefIterator, ParallelIterator};
use xlsxwriter::prelude::*;

mod lib;
use lib::{extract_value, get_cell_value, reg_replace_common, check_bad_val};
use crate::lib::contains_number;


fn process(fp: &str) -> Result<(), Box<dyn Error>> {
    let mut output_workbook = Workbook::new(&("output/".to_string() + fp))?;

    let mut workbook: Xlsx<_> = open_workbook(&("input/".to_string() + fp))?;
    let mut names = workbook.sheet_names();
    for sheet_name in names {
        let input_sheet = workbook.worksheet_range(&sheet_name).unwrap()?; // get sheet contents
        let mut output_sheet = output_workbook.add_worksheet(Some(&sheet_name))?; // create sheet inside output file

        if sheet_name == "Table_0" {
            for y in 0usize..24{

                // does value from 'labels' column contain "sales"
                let label =  get_cell_value(y, 8usize, &input_sheet);
                if !((label.contains("sales")) || (label.contains("Sales"))) {

                    // check if column 14 has a decimal
                    let col14_cell = get_cell_value(y, 9usize, &input_sheet);
                    if check_bad_val(&col14_cell) {

                        // determine whether column 14 has a decimal point + approximate size
                        if contains_number(&col14_cell) {

                            let (col14_value, _, _) = extract_value(&col14_cell);
                            if col14_value.parse::<f32>().is_ok() {
                                let reference_base = f32::log10(col14_value.parse::<f32>()?);

                                for x in 0..8u16 {
                                    let cell_value = reg_replace_common(get_cell_value(y, x as usize, &input_sheet));
                                    // println!("{cell_value}");
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
                                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &(prefix + &output_value + &postfix), Some(Format::new().set_bg_color(FormatColor::Orange)))?;
                                        } else {
                                            // inner val unparsable -> copy raw cell val

                                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, Some(Format::new().set_bg_color(FormatColor::Orange)))?;
                                        }
                                    } else {
                                        // execute if cell contains 'special value' ∊ NULL, nill, NMF, etc
                                        output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                                    }
                                }

                                for x in 8u16..10u16 {
                                    // copies last two columns because std code cannot be applied
                                    let cell_value = get_cell_value(y, x as usize, &input_sheet);
                                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                                }
                            } else {
                                // if 'last' column cannot be parsed for number
                                for x in 0u16..10u16 {
                                    let cell_value = get_cell_value(y, x as usize, &input_sheet);
                                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                                }
                            }
                        } else {
                            for x in 0u16..10u16 {
                                let cell_value = get_cell_value(y, x as usize, &input_sheet);
                                output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                            }
                        }

                    } else {
                        // generic copy if
                        for x in 0u16..10u16 {
                            let cell_value= get_cell_value(y, x as usize, &input_sheet);
                            output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                        }
                    }
                } else {
                    // generic copy (if label 'Sales')
                    for x in 0u16..10u16 {
                        let cell_value= get_cell_value(y, x as usize, &input_sheet);
                        output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                    }
                }
            }
        } else {
            // generically copies table if NOT "Table_0"
            for y in 0..input_sheet.height(){
                for x in 0u16..input_sheet.width() as u16 {
                    let cell_value= get_cell_value(y, x as usize, &input_sheet);
                    output_sheet.write_string(WorksheetRow::from(y as u32), WorksheetCol::from(x), &cell_value, None)?;
                }
            }
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut files: Vec<String> = vec![];

    let pool = rayon::ThreadPoolBuilder::new()
        .num_threads(4) // Set the number of threads in the pool
        .build()
        .unwrap();

    for file in std::fs::read_dir("./input/").unwrap() {
        files.push( file.unwrap().file_name().to_str().unwrap().to_string() )
    }



    let start_time = Instant::now();

    for (index, file) in files.iter().enumerate() {
        process(&file).unwrap();
        println!("{file} ({}/{})", index+1, files.len())
    }

    // let _ = pool.install(|| {
    //     files.par_iter() // Use parallel iterator
    //         .map(|item| process(item).unwrap()) // Replace with your computation
    // });


    println!("{:#?}", Instant::now() - start_time);
    Ok(())
}
