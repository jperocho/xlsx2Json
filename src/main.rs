use calamine::{Reader, Xlsx, open_workbook};
use std::collections::HashMap;
use std::time::Instant;
use serde_json;
use clap::Parser;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// xlsx-path
    #[arg(short, long)]
    path: String,

    /// json output
    #[arg(short, long)]
    output: Option<String>,
}

fn main() {
    let start = Instant::now(); // Start the timer

    let args = Args::parse();
    let path = args.path;

    // Check if the file exists
    if !std::path::Path::new(&path).exists() {
        eprintln!("File not found: {}", path);
        return;
    }

    let mut excel: Xlsx<_> = match open_workbook(&path) {
        Ok(workbook) => workbook,
        Err(e) => {
            eprintln!("Cannot open Excel file: {}", e);
            return;
        }
    };

    let sheet_names = excel.sheet_names().to_owned();
    if sheet_names.is_empty() {
        eprintln!("No sheets found in the Excel file.");
        return;
    }

    let first_sheet = &sheet_names[0];

    match excel.worksheet_range(first_sheet) {
        Ok(range) => {
            if let Some(headers) = range.headers() {
                // Create empty Vector to store the hashmap
                let mut data: Vec<HashMap<String, String>> = Vec::new();

                for (_, row) in range.rows().enumerate().skip(1) {
                    let mut hashmap: HashMap<String, String> = HashMap::new();

                    for (header, cell) in headers.iter().zip(row.iter()) {
                        hashmap.insert(header.clone(), cell.to_string().trim().to_string());
                    }

                    data.push(hashmap);
                }

                // Save data to json file
                let json_data = serde_json::to_string(&data).unwrap();

                // if output flag is provided, write to output path, otherwise write to data.json
                let outputfile = if let Some(output) = args.output {
                    output
                } else {
                    "data.json".to_string()
                };

                std::fs::write(outputfile, json_data).unwrap();

                let duration = start.elapsed(); // Calculate the elapsed time
                println!("Time taken: {:?}", duration);
            } else {
                eprintln!("Headers row not found.");
            }
        },
        Err(e) => {
            eprintln!("Error reading sheet '{}': {}", first_sheet, e);
        }
    }
}
