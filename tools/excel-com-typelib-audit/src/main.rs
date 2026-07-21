use std::env;
use std::path::PathBuf;

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-com-typelib-audit: {error}");
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
        "audit" => {
            let input = excel_com_typelib_audit::AuditInput {
                typelib_path: options.get("typelib").map(PathBuf::from),
                windows_version: options
                    .get("windows-version")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
                excel_file_version: options
                    .get("excel-file-version")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
                office_bitness: options
                    .get("office-bitness")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
            };
            let summary = excel_com_typelib_audit::audit(&root, &input)?;
            println!(
                "audited {} coclasses, {} interfaces, {} members, {} parameters, and {} enums",
                summary.coclasses,
                summary.interfaces,
                summary.members,
                summary.parameters,
                summary.enums
            );
        }
        "check" => {
            let input = excel_com_typelib_audit::AuditInput {
                typelib_path: options.get("typelib").map(PathBuf::from),
                windows_version: options
                    .get("windows-version")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
                excel_file_version: options
                    .get("excel-file-version")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
                office_bitness: options
                    .get("office-bitness")
                    .cloned()
                    .unwrap_or_else(|| "not-recorded".to_owned()),
            };
            excel_com_typelib_audit::check(&root, &input)?;
            println!("type-library evidence is current and deterministic");
        }
        "check-historical" => {
            excel_com_typelib_audit::check_historical(&root)?;
            println!("historical type-library evidence is current and deterministic");
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
        if result.insert(key.clone(), value).is_some() {
            return Err(format!("duplicate option --{key}"));
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
    "usage: excel-com-typelib-audit <audit|check|check-historical> --root <knowledge-root> [--typelib <path>] [--windows-version <value>] [--excel-file-version <value>] [--office-bitness <value>]".to_owned()
}
