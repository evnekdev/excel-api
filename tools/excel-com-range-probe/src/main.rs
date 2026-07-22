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
        "parity" => {
            let mode = options
                .get("mode")
                .ok_or_else(usage)
                .and_then(|mode| excel_com_range_probe::ParityMode::parse(mode))?;
            let fixture = options.get("fixture").map(PathBuf::from);
            let run_id = options
                .get("run-id")
                .map(String::as_str)
                .unwrap_or("05d-2026-07-21");
            let summary = excel_com_range_probe::parity(&root, mode, fixture.as_deref(), run_id)?;
            println!(
                "captured {} parity observations across {} completed cases ({} inconclusive)",
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
        "kernel-init" => {
            excel_com_range_probe::raw::initialize(&root)?;
            println!("windows-sys kernel evidence initialized without opening Excel");
        }
        "kernel" => {
            let backend = options
                .get("backend")
                .map(String::as_str)
                .unwrap_or("raw-windows-sys");
            let mode = options
                .get("mode")
                .map(String::as_str)
                .unwrap_or("all");
            let fixture = options.get("fixture").map(PathBuf::from);
            let action = options.get("action").map(String::as_str).unwrap_or("single");
            let summary = excel_com_range_probe::raw::run(
                &root,
                backend,
                mode,
                fixture.as_deref(),
                action,
            )?;
            println!("{summary}");
        }
        "value-matrix" => {
            let mode = options.get("mode").map(String::as_str).unwrap_or("L");
            let action = options
                .get("action")
                .map(String::as_str)
                .unwrap_or("scalar-value2");
            let case = options.get("case").map(String::as_str);
            let summary = excel_com_range_probe::raw::value_matrix(&root, mode, action, case)?;
            println!("{summary}");
        }
        "value-matrix-refresh" => {
            excel_com_range_probe::raw::refresh_value_matrix(&root)?;
            println!("value-runtime reports refreshed from existing evidence without opening Excel");
        }
        "value-matrix-check" => {
            excel_com_range_probe::raw::value_matrix_check(&root)?;
            println!("value-runtime evidence and reports are current and deterministic");
        }
        "error-scode" => {
            let mode = options.get("mode").map(String::as_str).unwrap_or("L");
            let summary = excel_com_range_probe::raw::error_scode(&root, mode)?;
            println!("{summary}");
        }
        "error-scode-refresh" => {
            excel_com_range_probe::raw::error_scode_refresh(&root)?;
            println!("error-SCODE reports refreshed without opening Excel");
        }
        "error-scode-python" => {
            let python = options.get("python").map(PathBuf::from).ok_or_else(usage)?;
            let client = options.get("client").map(String::as_str).ok_or_else(usage)?;
            let summary = excel_com_range_probe::raw::error_scode_python(&root, &python, client)?;
            println!("{summary}");
        }
        "error-scode-check" => {
            excel_com_range_probe::raw::error_scode_check(&root)?;
            println!("error-SCODE evidence and reports are current and deterministic");
        }
        "automation-value-check" => {
            excel_com_range_probe::automation_value_check(&root)?;
            println!("automation value-layer evidence and reports are current and deterministic");
        }
        "automation-value-live" => {
            let summary = excel_com_range_probe::automation_value_live(
                &root,
                options.get("case").map(String::as_str),
            )?;
            println!("{summary}");
        }
        "python-differential" => {
            let python = options
                .get("python")
                .map(PathBuf::from)
                .ok_or_else(usage)?;
            let client = options.get("client").map(String::as_str).ok_or_else(usage)?;
            let wrapper = options.get("wrapper").map(String::as_str).ok_or_else(usage)?;
            let family = options.get("family").map(String::as_str).unwrap_or("all");
            let environment_id = options
                .get("environment-id")
                .map(String::as_str)
                .ok_or_else(usage)?;
            let cache_dir = options
                .get("cache-dir")
                .map(PathBuf::from)
                .ok_or_else(usage)?;
            let summary = excel_com_range_probe::python_differential::capture(
                &root,
                &python,
                client,
                wrapper,
                family,
                environment_id,
                &cache_dir,
            )?;
            println!("{summary}");
        }
        "python-differential-refresh" => {
            excel_com_range_probe::python_differential::refresh(&root)?;
            println!("Python-client differential reports refreshed without opening Excel or Python");
        }
        "python-differential-check" => {
            excel_com_range_probe::python_differential::check(&root)?;
            println!("Python-client differential evidence and reports are current and deterministic");
        }
        "python-differential-raw" => {
            let mode = options.get("mode").map(String::as_str).unwrap_or("L");
            let family = options.get("family").map(String::as_str).ok_or_else(usage)?;
            let run_id = options.get("run-id").map(String::as_str);
            let result = excel_com_range_probe::raw::differential(mode, family)?;
            let summary = excel_com_range_probe::python_differential::capture_raw(
                &root, family, &result, run_id,
            )?;
            println!("{summary}");
        }
        "kernel-check" => {
            excel_com_range_probe::raw::check(&root)?;
            println!("windows-sys kernel evidence and reports are current and deterministic");
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
    "usage: excel-com-range-probe <live|diagnose|parity|refresh|check|kernel-init|kernel|kernel-check|value-matrix|value-matrix-refresh|value-matrix-check|error-scode|error-scode-refresh|error-scode-python|error-scode-check|automation-value-check|automation-value-live|python-differential|python-differential-raw|python-differential-refresh|python-differential-check> --root <knowledge-root> [--control-script <path>] [--mode <rust-baseline|pywin32-dynamic|pywin32-generated|comtypes-dynamic|comtypes-generated|L|S|X|all>] [--backend <raw-windows-sys|high-level-windows>] [--action <single|repeatability|compare|retry|scalar-value2>] [--fixture <controlled-fixture>] [--case <case-id>] [--run-id <id>] [--python <path> --client <pywin32|comtypes> --wrapper <dynamic|generated|all> --family <all|mixed|date|shape|dynamic> --environment-id <id> --cache-dir <path>]"
        .to_owned()
}
