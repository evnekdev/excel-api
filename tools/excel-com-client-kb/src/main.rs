use std::env;
use std::path::PathBuf;

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-com-client-kb: {error}");
        std::process::exit(1);
    }
}

fn run(arguments: Vec<String>) -> Result<(), String> {
    let Some(command) = arguments.first().map(String::as_str) else {
        return Err(usage());
    };
    let options = parse_options(&arguments[1..])?;
    let root = option_path(&options, "root")?;
    match command {
        "generate" => {
            let summary = excel_com_client_kb::generate(&root)?;
            println!(
                "generated {} source-derived records and {} deterministic reports",
                summary.records, summary.reports
            );
        }
        "check" => {
            let summary = excel_com_client_kb::check(&root)?;
            println!(
                "client-implementation knowledge base is valid, joined to {} typelib members, and deterministic",
                summary.typelib_joins
            );
        }
        "diagnose" => {
            let mode = options.get("mode").map(String::as_str).ok_or_else(usage)?;
            excel_com_client_kb::diagnose(&root, mode, options.get("python").map(String::as_str))?;
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
    "usage: excel-com-client-kb <generate|check> --root <knowledge-root>; diagnose requires --mode <pywin32-dynamic|pywin32-generated|comtypes|comtypes-generated> and accepts optional --python <interpreter>".to_owned()
}
