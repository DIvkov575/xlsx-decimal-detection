use calamine::{CellType, DataType};
use unicode_segmentation::UnicodeSegmentation;
use regex::{Regex, RegexSet};


pub fn check_bad_val(value: &str) -> bool {
    (value != "") && (value != "Nil") && (value != "--") && (value.to_ascii_lowercase() != "nan") && (value != "NMF")
}

pub fn contains_character(input: &str) -> bool {
    for character in input.graphemes(true) {
        if character.parse::<u8>().is_err() {
            return true;
        }
    }
    false
}
pub fn contains_number(input: &str) -> bool {
    for character in input.graphemes(true) {
        if character.parse::<u8>().is_ok() {
            return true;
        }
    }
    false
}

pub fn extract_value(input: &str) -> (String, String, String) {
    let mut prefix: String = "".to_string();
    let mut postfix: String = "".to_string();
    #[warn(unused_assignments)]
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

pub fn count_postfix(input: &str, character: char) -> u32 {
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

pub fn get_cell_value<T>(y: usize, x: usize, input_sheet: &calamine::Range<T>) -> String
where
    T: CellType,
    T: std::fmt::Display,
    // &T: Default
{
    if let Some(a) = input_sheet.get_value((y as u32, x as u32)) {
        return a.to_string()
    } else {
        "".to_string()
    }
}

pub fn reg_replace_common(input: String) -> String {
    if input.to_lowercase().contains("nme") {
        return "NMF".to_string();
    }
    let output = input.replace(",", ".");

   output
}