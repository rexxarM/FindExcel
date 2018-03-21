extern crate glob;
extern crate calamine;
extern crate time;
extern crate regex;

use regex::Regex;
use glob::glob;
use calamine::{Reader, open_workbook, Xlsx, CellType};
use time::PreciseTime;

fn main() {
    println!("find excel begin...");
    let path = "./**/*.xlsx";
    let excel = "test/Localization_Skill.xlsx";
    // visit(path);
    let start = PreciseTime::now();
    let mut workbook: Xlsx<_> = open_workbook(excel).expect("Cannot open file");
    let sheets = workbook.sheet_names().to_owned();
    for s in sheets {
        if let Some(Ok(range)) = workbook.worksheet_range(&s) {
            let total_cells = range.get_size().0 * range.get_size().1;
            let non_empty_cells: usize = range.used_cells().count();
            for row in range.rows() { 
                for cell in row{
                    println!("cell in {} => {:?}",s, cell);
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