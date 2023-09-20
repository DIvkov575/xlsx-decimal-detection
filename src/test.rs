#[cfg(test)]
mod test {
    use std::fs;
    use super::*;

    #[test]
    fn correct_file_reading() {
        assert!(std::fs::Path::new("input").is_dir());

        if std::fs::Path::new("output").is_dir() {
            fs::remove_dir_all("output").unwrap()
        }
        fs::create_dir("output");

        assert!(process_all("output", false).is_ok()0);


    }
}