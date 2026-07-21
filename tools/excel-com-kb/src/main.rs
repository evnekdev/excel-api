use std::env;
use std::path::PathBuf;

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-com-kb: {error}");
        std::process::exit(1);
    }
}

fn run(arguments: Vec<String>) -> Result<(), String> {
    let Some(command) = arguments.first().map(String::as_str) else {
        return Err(usage());
    };
    let options = parse_options(&arguments[1..])?;
    match command {
        "ingest" => {
            let source = option_path(&options, "source")?;
            let manifest = option_path(&options, "manifest")?;
            let output = option_path(&options, "output")?;
            let result = excel_com_kb::ingest(&source, &manifest, &output)?;
            println!(
                "ingested {} objects, {} members, {} enums, and {} relationships from {} selected files",
                result.objects,
                result.members,
                result.enums,
                result.relationships,
                result.coverage.selected_files
            );
        }
        "validate" => {
            let root = option_path(&options, "root")?;
            let result = excel_com_kb::validate(&root)?;
            println!(
                "validated {} objects, {} members, {} enums, and {} relationships",
                result.objects, result.members, result.enums, result.relationships
            );
        }
        "generate" => {
            let root = option_path(&options, "root")?;
            let result = excel_com_kb::generate(&root)?;
            println!("generated {} reports", result.reports);
        }
        "analyze" => {
            let root = option_path(&options, "root")?;
            let result = excel_com_kb::analyze(&root)?;
            println!("generated {} analysis reports", result.reports);
        }
        "check" => {
            let root = option_path(&options, "root")?;
            excel_com_kb::check(&root)?;
            println!("knowledge base is valid and generated output is current and deterministic");
        }
        _ => return Err(usage()),
    }
    Ok(())
}

fn parse_options(
    arguments: &[String],
) -> Result<std::collections::BTreeMap<String, String>, String> {
    let mut result = std::collections::BTreeMap::new();
    let mut index = 0;
    while index < arguments.len() {
        let key = arguments[index]
            .strip_prefix("--")
            .ok_or_else(usage)?
            .to_owned();
        let value = arguments.get(index + 1).ok_or_else(usage)?.to_owned();
        if result.insert(key, value).is_some() {
            return Err("duplicate CLI option".to_owned());
        }
        index += 2;
    }
    Ok(result)
}

fn option_path(
    options: &std::collections::BTreeMap<String, String>,
    name: &str,
) -> Result<PathBuf, String> {
    options.get(name).map(PathBuf::from).ok_or_else(usage)
}

fn usage() -> String {
    "usage: excel-com-kb <ingest|validate|generate|analyze|check> --source <checkout> --manifest <SOURCE_MANIFEST.toml> --output <data-dir>; validation commands use --root <knowledge-root>".to_owned()
}
