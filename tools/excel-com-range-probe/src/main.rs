use std::collections::BTreeMap;
use std::env;
use std::path::PathBuf;

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-com-range-probe: {error}");
        std::process::exit(1);
    }
}

fn run(arguments: Vec<String>) -> Result<(), String> {
    let Some(command) = arguments.first().map(String::as_str) else {
        return Err(usage());
    };
    let options = parse_options(&arguments[1..])?;
    let root = options.get("root").map(PathBuf::from).ok_or_else(usage)?;
    match command {
        "live" => {
            let control_script = options.get("control-script").map(PathBuf::from);
            let summary = excel_com_range_probe::live(&root, control_script.as_deref())?;
            println!(
                "captured {} runtime observations across {} completed cases ({} inconclusive)",
                summary.observations, summary.completed_cases, summary.inconclusive_cases
            );
        }
        "diagnose" => {
            let control_script = options.get("control-script").map(PathBuf::from);
            let summary = excel_com_range_probe::diagnose(&root, control_script.as_deref())?;
            println!(
                "captured {} diagnostic/runtime observations across {} completed cases ({} inconclusive)",
                summary.observations, summary.completed_cases, summary.inconclusive_cases
            );
        }
        "check" => {
            excel_com_range_probe::check(&root)?;
            println!("runtime evidence and reports are current and deterministic");
        }
        "refresh" => {
            excel_com_range_probe::refresh(&root)?;
            println!("runtime reports refreshed from existing evidence without opening Excel");
        }
        _ => return Err(usage()),
    }
    Ok(())
}

fn parse_options(arguments: &[String]) -> Result<BTreeMap<String, String>, String> {
    let mut options = BTreeMap::new();
    let mut index = 0;
    while index < arguments.len() {
        let key = arguments[index]
            .strip_prefix("--")
            .ok_or_else(usage)?
            .to_owned();
        let value = arguments.get(index + 1).ok_or_else(usage)?.to_owned();
        if options.insert(key.clone(), value).is_some() {
            return Err(format!("duplicate option --{key}"));
        }
        index += 2;
    }
    Ok(options)
}

fn usage() -> String {
    "usage: excel-com-range-probe <live|diagnose|refresh|check> --root <knowledge-root> [--control-script <path>]"
        .to_owned()
}
