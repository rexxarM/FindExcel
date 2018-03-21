extern crate glob;
extern crate calamine;
extern crate time;
extern crate regex;
extern crate clap;

use regex::Regex;
use glob::glob;
use calamine::{Reader, open_workbook, Xlsx, CellType};
use time::PreciseTime;

fn main() {
    println!("find excel begin...");
    let path = "./**/*.xlsx";
    let excel = "test/Seal.xlsx";
    let re = Regex::new(r"Tips\d+").unwrap();

    // visit(path);
    let start = PreciseTime::now();
    let mut workbook: Xlsx<_> = open_workbook(excel).expect("Cannot open file");
    let sheets = workbook.sheet_names().to_owned();
    for s in sheets {
        if let Some(Ok(range)) = workbook.worksheet_range(&s) {
            for row in range.rows() { 
                for (i, c) in row.iter().enumerate() {
                    let value = c.to_string();
                    if re.is_match(&value){
                        println!("cell in {} {}=> {:?}",s, i,value);
                    }
                }
            }
        }
    }
    

    let end = PreciseTime::now();
    println!("=> {} seconds for whatever you did.", start.to(end));
}

fn visit(pathpattern:&str){
    for entry in glob(pathpattern).expect("Failed to read glob pattern") {
    match entry {
        Ok(path) => println!("{:?}", path.display()),
        Err(e) => println!("{:?}", e),
    }
    }
}