use chrono::{Datelike, NaiveDate};

fn main() {
    let excel_date: i64 = 45422;
    let date = NaiveDate::from_num_days_from_ce_opt(excel_date.try_into().unwrap());
    println!("{}", format!("{:?}", date.unwrap().year()));
}