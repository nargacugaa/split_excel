use calamine::{open_workbook, Reader, Xlsx};
use chrono::NaiveDate;
use rust_xlsxwriter::Workbook;
use std::error::Error;
use std::{fs, io};
use std::path::{Path, PathBuf};

fn split_excel_file(
    input_file: &str,
    output_prefix: &str,
    rows_per_file: usize,
) -> Result<(), Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(input_file)?;

    // 使用 match 处理可能的错误
    let range = match workbook.worksheet_range("Sheet1") {
        Ok(r) => r,
        Err(_) => return Err(Box::from("Cannot find 'Sheet1'")),
    };

    let mut current_row = 1;
    let mut new_start_row = 1;
    let mut file_index = 1;

    // 提取第一行数据
    let first_row = range.rows().next().unwrap_or(&[]).to_vec();

    loop {
        let mut new_workbook = Workbook::new();
        let new_worksheet = new_workbook.add_worksheet();

        // 写入第一行数据
        for (col, cell) in first_row.iter().enumerate() {
            let value = match cell {
                calamine::Data::Int(i) => i.to_string(),
                calamine::Data::Float(f) => f.to_string(),
                calamine::Data::String(s) => s.clone(),
                calamine::Data::Bool(b) => {
                    if *b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                calamine::Data::Error(e) => e.to_string(),
                calamine::Data::Empty => "".to_string(),
                calamine::Data::DateTime(_) => panic!("DateTimeIso is not supported"),
                calamine::Data::DateTimeIso(_) => panic!("DateTimeIso is not supported"),
                calamine::Data::DurationIso(_) => panic!("DurationIso is not supported"),
            };
            new_worksheet.write_string(0, col as u16, &value)?;
        }

        for row in range.rows().skip(current_row).take(rows_per_file) {
            for (col, cell) in row.iter().enumerate() {
                let value = match cell {
                    calamine::Data::Int(i) => i.to_string(),
                    calamine::Data::Float(f) => f.to_string(),
                    calamine::Data::String(s) => s.clone(),
                    calamine::Data::Bool(b) => {
                        if *b {
                            "TRUE".to_string()
                        } else {
                            "FALSE".to_string()
                        }
                    }
                    calamine::Data::Error(e) => e.to_string(),
                    calamine::Data::Empty => "".to_string(),
                    calamine::Data::DateTime(excel_time) => {
                        let serial = excel_time.as_f64();
                        let date = NaiveDate::from_ymd_opt(1899, 12, 31)
                            .and_then(|base_date| {
                                base_date.checked_add_signed(chrono::Duration::days(
                                    (serial - 1.0).floor() as i64,
                                ))
                            })
                            .expect("Failed to convert Excel date to NaiveDate");
                        // 使用转换后的日期
                        date.format("%Y/%m/%d").to_string()
                    }
                    calamine::Data::DateTimeIso(_) => panic!("DateTimeIso is not supported"),
                    calamine::Data::DurationIso(_) => panic!("DurationIso is not supported"),
                };
                new_worksheet.write_string(new_start_row as u32, col as u16, &value)?;
            }
            current_row += 1;
            new_start_row += 1;
        }

        let output_file = format!("{}_{:03}.xlsx", output_prefix, file_index);
        new_workbook.save(Path::new(&output_file))?;


        // 新Excel从第0行开始
        new_start_row = 1;

        println!("已拆分为：{} 个Excel",file_index);

        if current_row >= range.height() {
            break;
        }

        file_index += 1;
    }

    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let current_dir = std::env::current_dir()?;
    let mut input_file: Option<PathBuf> = None;

    // 查找当前目录中的第一个 .xlsx 或 .xls 文件
    for entry in fs::read_dir(current_dir.clone())? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|s| s.to_str()) == Some("xlsx")
            || path.extension().and_then(|s| s.to_str()) == Some("xls")
        {
            input_file = Some(path);
            break;
        }
    }

    let input_file = match input_file {
        Some(path) => path,
        None => {
            return Err(Box::from(
                "No .xlsx or .xls file found in the current directory",
            ))
        }
    };

    // 创建 result 目录
    let output_dir = current_dir.join("result");
    fs::create_dir_all(&output_dir)?;

    // 设置输出文件前缀
    let output_prefix = output_dir.join("output").to_str().unwrap().to_string();

    let rows_per_file = 10000;

    split_excel_file(input_file.to_str().unwrap(), &output_prefix, rows_per_file)?;

    // 完成
    println!("完成！！ Excel 文件已成功拆分 \r\n按回车键关闭窗口");
    // 从标准输入读取一次输入，不关心输入的具体内容
    let mut buffer = String::new();
    let _ = io::stdin().read_line(&mut buffer).ok();

    Ok(())

// 使用终端原始模式 未尝试
//     use crossterm::{
//     event::{read, Event, KeyCode},
//     execute,
//     terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
// };
// use std::io::{self, Write};

// fn main() -> crossterm::Result<()> {
//     // 启用终端的原始模式
//     enable_raw_mode()?;

//     // 你的程序逻辑
//     println!("Hello, world!");

//     // 程序执行完毕，等待用户按键
//     println!("Press any key to continue...");

//     // 循环等待按键事件
//     loop {
//         if let Event::Key(_) = read()? {
//             break;
//         }
//     }

//     // 禁用原始模式
//     disable_raw_mode()?;

//     Ok(())
// }

}
