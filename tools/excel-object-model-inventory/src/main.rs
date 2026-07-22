#![cfg(windows)]

mod markdown;
mod model;
mod typelib;
mod validate;

use std::env;
use std::path::PathBuf;

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-object-model-inventory: {error}");
        std::process::exit(1);
    }
}

fn run(arguments: Vec<String>) -> Result<(), String> {
    let command = arguments.first().map(String::as_str).ok_or_else(usage)?;
    let root = arguments
        .windows(2)
        .find_map(|pair| (pair[0] == "--root").then(|| PathBuf::from(&pair[1])))
        .unwrap_or_else(|| PathBuf::from("."));
    match command {
        "extract" => {
            let summary = typelib::extract(&root)?;
            println!(
                "extracted {} type infos: {} objects, {} members, {} enums",
                summary.type_infos, summary.objects, summary.members, summary.enums
            );
        }
        "generate" => {
            let summary = markdown::generate(&root)?;
            println!(
                "generated {} detailed pages, {} enum pages, and {} indexes",
                summary.pages, summary.enum_pages, summary.indexes
            );
        }
        "check" => {
            validate::check(&root)?;
            println!(
                "Excel object-model metadata and generated documentation are current and deterministic"
            );
        }
        _ => return Err(usage()),
    }
    Ok(())
}

fn usage() -> String {
    "usage: cargo run -p excel-object-model-inventory -- <extract|generate|check> [--root <repository-root>]".to_owned()
}
