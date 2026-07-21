//! Fresh-process delta debugging for the bounded Excel Automation pre-`Add`
//! sequence.  Normal commands manipulate only deterministic evidence; live
//! execution is explicit and never starts Excel from tests.

#![cfg_attr(not(windows), allow(dead_code, unused_imports))]
#![allow(clippy::too_many_arguments)]

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::collections::{BTreeMap, BTreeSet};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

const SCHEMA_VERSION: u32 = 1;
const SCHEDULE_SEED: u64 = 20_260_721;
const DEFAULT_REPEATS: u32 = 5;

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Serialize, Deserialize)]
enum Mode {
    L,
    S,
    X,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "L" => Ok(Self::L),
            "S" => Ok(Self::S),
            "X" => Ok(Self::X),
            _ => Err("mode must be L, S, or X".to_owned()),
        }
    }

    fn id(self) -> &'static str {
        match self {
            Self::L => "L",
            Self::S => "S",
            Self::X => "X",
        }
    }

    fn activation_api(self) -> &'static str {
        match self {
            Self::X => "CoCreateInstanceEx",
            Self::L | Self::S => "CoCreateInstance",
        }
    }

    fn clsctx(self) -> &'static str {
        match self {
            Self::L => "CLSCTX_LOCAL_SERVER",
            Self::S | Self::X => "CLSCTX_SERVER",
        }
    }

    fn lcid(self) -> &'static str {
        match self {
            Self::L => "0x0400",
            Self::S | Self::X => "0x0000",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Operation {
    Version,
    WorkbooksTypeInfoCount,
    WorkbooksQueryIUnknown,
    WorkbooksQueryIDispatch,
    WorkbooksCount,
    LifetimeCloneThenClear,
    LifetimeRetainThenClear,
    LifetimeQueryInterfaceThenClear,
}

impl Operation {
    fn id(self) -> &'static str {
        match self {
            Self::Version => "application-version",
            Self::WorkbooksTypeInfoCount => "workbooks-get-type-info-count",
            Self::WorkbooksQueryIUnknown => "workbooks-query-interface-iunknown",
            Self::WorkbooksQueryIDispatch => "workbooks-query-interface-idispatch",
            Self::WorkbooksCount => "workbooks-count",
            Self::LifetimeCloneThenClear => "lifetime-clone-then-clear",
            Self::LifetimeRetainThenClear => "lifetime-retain-then-clear",
            Self::LifetimeQueryInterfaceThenClear => "lifetime-query-interface-then-clear",
        }
    }

    fn stage(self) -> &'static str {
        match self {
            Self::Version => "before-workbooks",
            _ => "after-workbooks",
        }
    }
}

const PREFIX_OPERATIONS: &[Operation] = &[
    Operation::Version,
    Operation::WorkbooksTypeInfoCount,
    Operation::WorkbooksQueryIUnknown,
    Operation::WorkbooksQueryIDispatch,
    Operation::WorkbooksCount,
    Operation::LifetimeCloneThenClear,
    Operation::LifetimeRetainThenClear,
    Operation::LifetimeQueryInterfaceThenClear,
];

#[derive(Clone, Debug)]
struct CaseSpec {
    id: String,
    group: &'static str,
    before_workbooks: Vec<LiveOperation>,
    after_workbooks: Vec<LiveOperation>,
    ownership: Ownership,
    storage: StorageMode,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Ownership {
    W0RetainResult,
    W1CloneThenClearAfter,
    W2CloneThenClearImmediately,
    W3QueryIDispatchThenClear,
    W4QueryIUnknownThenIDispatchThenClear,
    W5ReacquireBeforeAdd,
}

impl Ownership {
    fn id(self) -> &'static str {
        match self {
            Self::W0RetainResult => "W0-retain-result",
            Self::W1CloneThenClearAfter => "W1-clone-clear-after",
            Self::W2CloneThenClearImmediately => "W2-clone-clear-immediately",
            Self::W3QueryIDispatchThenClear => "W3-query-idispatch-clear",
            Self::W4QueryIUnknownThenIDispatchThenClear => "W4-query-iunknown-idispatch-clear",
            Self::W5ReacquireBeforeAdd => "W5-reacquire-before-add",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StorageMode {
    Fresh,
    ReusedZeroed,
}

impl StorageMode {
    fn id(self) -> &'static str {
        match self {
            Self::Fresh => "fresh-per-invoke",
            Self::ReusedZeroed => "reused-after-zeroing",
        }
    }
}

#[derive(Clone, Debug)]
enum LiveOperation {
    Prefix(Operation),
    Property(&'static str),
    ApplicationTypeInfoCount,
    ApplicationTypeInfo,
    WorkbooksTypeInfo,
    ReacquireWorkbooks,
    ProcessPid,
    ProcessStartTime,
    ProcessHandle,
}

impl LiveOperation {
    fn id(&self) -> &'static str {
        match self {
            Self::Prefix(value) => value.id(),
            Self::Property(value) => value,
            Self::ApplicationTypeInfoCount => "application-get-type-info-count",
            Self::ApplicationTypeInfo => "application-get-type-info",
            Self::WorkbooksTypeInfo => "workbooks-get-type-info",
            Self::ReacquireWorkbooks => "reacquire-workbooks",
            Self::ProcessPid => "hwnd-to-pid",
            Self::ProcessStartTime => "process-start-time",
            Self::ProcessHandle => "process-exit-handle",
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ScheduleEntry {
    run_id: String,
    scenario: String,
    group: String,
    prefix_id: String,
    mode: String,
    repetition: u32,
    order_index: u32,
    sequence_hash: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct TraceEvent {
    ordinal: u32,
    interface: String,
    member: String,
    dispid: String,
    flags: String,
    lcid: String,
    c_args: u32,
    c_named_args: u32,
    result_vartype: String,
    hresult: String,
    ownership_transition: String,
    variant_clear_timing: String,
    release_timing: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct RunRecord {
    schema_version: u32,
    run_id: String,
    scenario: String,
    group: String,
    prefix_id: String,
    mode: String,
    sequence_hash: String,
    execution_order_index: u32,
    process_architecture: String,
    thread_apartment: String,
    activation_api: String,
    clsctx: String,
    get_ids_of_names_lcid: String,
    invoke_lcid: String,
    excel_version: String,
    workbooks_count_before_add: String,
    operations: BTreeMap<String, Value>,
    add_hresult: String,
    excepinfo_scode: String,
    pu_arg_err_raw: String,
    workbook_created: bool,
    cleanup: BTreeMap<String, Value>,
    trace: Vec<TraceEvent>,
}

pub fn run(arguments: Vec<String>) -> Result<(), String> {
    let Some(command) = arguments.first().map(String::as_str) else {
        return Err(usage());
    };
    let options = Options::parse(&arguments[1..])?;
    match command {
        "generate" => generate(options.root()?),
        "render" => render(options.root()?),
        "check" => check(options.root()?),
        "matrix" => matrix(options),
        "recovery" => recovery(options),
        "case" => run_case_command(options),
        _ => Err(usage()),
    }
}

fn usage() -> String {
    "usage: excel-com-pre-add-delta <generate|render|check|matrix|recovery|case> --root <knowledge/excel-object-model/pre-add-delta> [--repeats <count>] [--seed <u64>] [--mode <L|S|X>] [--scenario <id>] [--run-id <id>] [--order <n>] [--fixture <temporary-xlsx>]".to_owned()
}

#[derive(Default)]
struct Options {
    root: Option<PathBuf>,
    repeats: Option<u32>,
    seed: Option<u64>,
    mode: Option<Mode>,
    scenario: Option<String>,
    run_id: Option<String>,
    order: Option<u32>,
    fixture: Option<PathBuf>,
}

impl Options {
    fn parse(values: &[String]) -> Result<Self, String> {
        let mut result = Self::default();
        let mut index = 0;
        while index < values.len() {
            let name = values[index].as_str();
            let value = values
                .get(index + 1)
                .ok_or_else(|| format!("missing value for {name}"))?;
            match name {
                "--root" => result.root = Some(PathBuf::from(value)),
                "--repeats" => {
                    result.repeats = Some(
                        value
                            .parse()
                            .map_err(|_| "--repeats must be a positive integer")?,
                    )
                }
                "--seed" => {
                    result.seed = Some(value.parse().map_err(|_| "--seed must be an integer")?)
                }
                "--mode" => result.mode = Some(Mode::parse(value)?),
                "--scenario" => result.scenario = Some(value.clone()),
                "--run-id" => result.run_id = Some(value.clone()),
                "--order" => {
                    result.order = Some(value.parse().map_err(|_| "--order must be an integer")?)
                }
                "--fixture" => result.fixture = Some(PathBuf::from(value)),
                _ => return Err(format!("unknown option: {name}")),
            }
            index += 2;
        }
        Ok(result)
    }

    fn root(&self) -> Result<&Path, String> {
        self.root
            .as_deref()
            .ok_or_else(|| "--root is required".to_owned())
    }
}

fn prefix_case(length: usize) -> CaseSpec {
    let mut before_workbooks = Vec::new();
    let mut after_workbooks = Vec::new();
    for operation in PREFIX_OPERATIONS.iter().take(length).copied() {
        let operation = LiveOperation::Prefix(operation);
        match operation {
            LiveOperation::Prefix(Operation::Version) => before_workbooks.push(operation),
            _ => after_workbooks.push(operation),
        }
    }
    CaseSpec {
        id: format!("A{length}"),
        group: "prefix",
        before_workbooks,
        after_workbooks,
        ownership: Ownership::W2CloneThenClearImmediately,
        storage: StorageMode::Fresh,
    }
}

fn supplemental_cases() -> Vec<CaseSpec> {
    let mut result = Vec::new();
    for ownership in [
        Ownership::W0RetainResult,
        Ownership::W1CloneThenClearAfter,
        Ownership::W2CloneThenClearImmediately,
        Ownership::W3QueryIDispatchThenClear,
        Ownership::W4QueryIUnknownThenIDispatchThenClear,
        Ownership::W5ReacquireBeforeAdd,
    ] {
        result.push(CaseSpec {
            id: ownership.id().to_owned(),
            group: "ownership",
            before_workbooks: Vec::new(),
            after_workbooks: Vec::new(),
            ownership,
            storage: StorageMode::Fresh,
        });
    }
    for storage in [StorageMode::Fresh, StorageMode::ReusedZeroed] {
        result.push(CaseSpec {
            id: format!("storage-{}", storage.id()),
            group: "storage",
            before_workbooks: vec![LiveOperation::Prefix(Operation::Version)],
            after_workbooks: vec![LiveOperation::Prefix(Operation::WorkbooksCount)],
            ownership: Ownership::W2CloneThenClearImmediately,
            storage,
        });
    }
    for property in [
        "Hwnd",
        "Ready",
        "Calculation",
        "AutomationSecurity",
        "DisplayAlerts",
        "UserControl",
        "Interactive",
        "Version",
    ] {
        result.push(CaseSpec {
            id: format!(
                "property-before-workbooks-{}",
                property.to_ascii_lowercase()
            ),
            group: "property",
            before_workbooks: vec![LiveOperation::Property(property)],
            after_workbooks: Vec::new(),
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
        result.push(CaseSpec {
            id: format!("property-after-workbooks-{}", property.to_ascii_lowercase()),
            group: "property",
            before_workbooks: Vec::new(),
            after_workbooks: vec![LiveOperation::Property(property)],
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
        result.push(CaseSpec {
            id: format!("property-repeat-{}", property.to_ascii_lowercase()),
            group: "property",
            before_workbooks: vec![
                LiveOperation::Property(property),
                LiveOperation::Property(property),
            ],
            after_workbooks: Vec::new(),
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
        result.push(CaseSpec {
            id: format!("property-reacquire-{}", property.to_ascii_lowercase()),
            group: "property",
            before_workbooks: Vec::new(),
            after_workbooks: vec![
                LiveOperation::Property(property),
                LiveOperation::ReacquireWorkbooks,
            ],
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
    }
    for operation in [
        LiveOperation::ApplicationTypeInfoCount,
        LiveOperation::ApplicationTypeInfo,
        LiveOperation::Prefix(Operation::WorkbooksTypeInfoCount),
        LiveOperation::WorkbooksTypeInfo,
    ] {
        let application_operation = matches!(
            operation,
            LiveOperation::ApplicationTypeInfoCount | LiveOperation::ApplicationTypeInfo
        );
        result.push(CaseSpec {
            id: format!("type-info-{}", operation.id()),
            group: "type-info",
            before_workbooks: if application_operation {
                vec![operation.clone()]
            } else {
                Vec::new()
            },
            after_workbooks: if application_operation {
                Vec::new()
            } else {
                vec![operation]
            },
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
    }
    for operation in [
        LiveOperation::Property("Hwnd"),
        LiveOperation::ProcessPid,
        LiveOperation::ProcessStartTime,
        LiveOperation::ProcessHandle,
    ] {
        result.push(CaseSpec {
            id: format!("process-{}", operation.id()),
            group: "process-instrumentation",
            before_workbooks: vec![operation],
            after_workbooks: Vec::new(),
            ownership: Ownership::W2CloneThenClearImmediately,
            storage: StorageMode::Fresh,
        });
    }
    result
}

fn case_by_id(id: &str) -> Result<CaseSpec, String> {
    if let Some(length) = id
        .strip_prefix('A')
        .and_then(|value| value.parse::<usize>().ok())
    {
        if length <= PREFIX_OPERATIONS.len() {
            return Ok(prefix_case(length));
        }
    }
    supplemental_cases()
        .into_iter()
        .find(|case| case.id == id)
        .ok_or_else(|| format!("unknown scenario: {id}"))
}

fn case_hash(case: &CaseSpec) -> String {
    let mut text = format!("{}|{}|{}", case.id, case.ownership.id(), case.storage.id());
    for value in case
        .before_workbooks
        .iter()
        .chain(case.after_workbooks.iter())
    {
        text.push('|');
        text.push_str(value.id());
    }
    let mut state = 0xcbf2_9ce4_8422_2325_u64;
    for byte in text.bytes() {
        state ^= u64::from(byte);
        state = state.wrapping_mul(0x0000_0100_0000_01b3);
    }
    format!("fnv1a64-{state:016x}")
}

fn all_cases() -> Vec<CaseSpec> {
    let mut result: Vec<_> = (0..=PREFIX_OPERATIONS.len()).map(prefix_case).collect();
    result.extend(supplemental_cases());
    result
}

fn generate(root: &Path) -> Result<(), String> {
    fs::create_dir_all(root).map_err(io_error)?;
    fs::create_dir_all(report_root(root)).map_err(io_error)?;
    write(root.join("SOURCE_MANIFEST.toml"), source_manifest())?;
    write(
        root.join("environments.jsonl"),
        format!("{}\n", environment_json()),
    )?;
    write(root.join("prefixes.jsonl"), prefixes_jsonl())?;
    write(
        root.join("operation-definitions.jsonl"),
        operation_definitions_jsonl(),
    )?;
    for name in [
        "run-schedule.jsonl",
        "observations.jsonl",
        "ownership-variants.jsonl",
        "native-process-context.jsonl",
        "cold-session-baseline.jsonl",
        "session-state-transitions.jsonl",
        "owned-process-cleanup.jsonl",
        "repairs.jsonl",
        "unresolved.jsonl",
    ] {
        let path = root.join(name);
        if !path.exists() {
            write(path, "")?;
        }
    }
    render(root)
}

fn source_manifest() -> String {
    "schema_version = 1\ntool = \"excel-com-pre-add-delta\"\nprompt = \"05F\"\nstarting_origin_master = \"bbf767d522e968969cd5afd143dfcc80718c57cf\"\nwindows_crate = \"0.62.2\"\nwindows_core = \"0.62.2\"\nwindows_sys = \"0.61.2\"\nprocess_isolation = \"one child process and one owned Excel Automation instance per live run\"\nraw_pointers_persisted = false\nraw_hwnds_persisted = false\nraw_pids_persisted = false\nlocal_paths_persisted = false\n".to_owned()
}

fn environment_json() -> String {
    json!({
        "schema_version": SCHEMA_VERSION,
        "id": "windows-excel-64-05f",
        "windows_version": "Windows 10 Enterprise 25H2 build 26200.8875",
        "excel_file_version": "16.0.20131.20154",
        "office_bitness": "64-bit",
        "architecture": "x64",
        "thread_apartment": "STA",
        "windows": "0.62.2",
        "windows_sys": "0.61.2",
        "raw_process_identity_persisted": false,
    })
    .to_string()
}

fn prefixes_jsonl() -> String {
    (0..=PREFIX_OPERATIONS.len())
        .map(|length| {
            let case = prefix_case(length);
            json!({
                "schema_version": SCHEMA_VERSION,
                "id": case.id,
                "sequence_hash": case_hash(&case),
                "operation_count": length,
                "operations": PREFIX_OPERATIONS.iter().take(length).map(|operation| operation.id()).collect::<Vec<_>>(),
                "actual_full_harness_prefix": true,
            })
            .to_string()
        })
        .collect::<Vec<_>>()
        .join("\n")
        + "\n"
}

fn operation_definitions_jsonl() -> String {
    let mut values = Vec::new();
    values.push(json!({"schema_version": SCHEMA_VERSION, "id": "A0", "role": "baseline", "description": "Activate Excel, get Workbooks, invoke Add, close, Quit"}));
    for (index, operation) in PREFIX_OPERATIONS.iter().enumerate() {
        values.push(json!({
            "schema_version": SCHEMA_VERSION,
            "id": format!("A{}", index + 1),
            "operation": operation.id(),
            "stage": operation.stage(),
            "description": "Added alone to the immediately preceding production-harness prefix",
        }));
    }
    for case in supplemental_cases() {
        values.push(json!({
            "schema_version": SCHEMA_VERSION,
            "id": case.id,
            "group": case.group,
            "ownership": case.ownership.id(),
            "storage": case.storage.id(),
            "before_workbooks": case.before_workbooks.iter().map(LiveOperation::id).collect::<Vec<_>>(),
            "after_workbooks": case.after_workbooks.iter().map(LiveOperation::id).collect::<Vec<_>>(),
        }));
    }
    values
        .into_iter()
        .map(|value| value.to_string())
        .collect::<Vec<_>>()
        .join("\n")
        + "\n"
}

fn matrix(options: Options) -> Result<(), String> {
    let root = options.root()?.to_path_buf();
    generate(&root)?;
    let repeats = options.repeats.unwrap_or(DEFAULT_REPEATS);
    if repeats == 0 {
        return Err("--repeats must be positive".to_owned());
    }
    let seed = options.seed.unwrap_or(SCHEDULE_SEED);
    let mode_filter = options.mode;
    let schedule = schedule(seed, repeats, mode_filter);
    write_jsonl(root.join("run-schedule.jsonl"), &schedule)?;
    let executable = env::current_exe().map_err(io_error)?;
    let mut records = Vec::new();
    for entry in &schedule {
        let output = Command::new(&executable)
            .arg("case")
            .arg("--root")
            .arg(&root)
            .arg("--scenario")
            .arg(&entry.scenario)
            .arg("--mode")
            .arg(&entry.mode)
            .arg("--run-id")
            .arg(&entry.run_id)
            .arg("--order")
            .arg(entry.order_index.to_string())
            .output()
            .map_err(io_error)?;
        if !output.status.success() {
            return Err(format!(
                "live child {} failed with {}: {}",
                entry.run_id,
                output.status,
                String::from_utf8_lossy(&output.stderr).trim()
            ));
        }
        let record: RunRecord = serde_json::from_slice(&output.stdout).map_err(|error| {
            format!(
                "child {} did not return one JSON record: {error}",
                entry.run_id
            )
        })?;
        if record.cleanup.get("owned_process_exited") != Some(&Value::Bool(true)) {
            return Err(format!(
                "{} did not confirm owned-process exit; schedule stopped safely",
                entry.run_id
            ));
        }
        records.push(record);
    }
    write_jsonl(root.join("observations.jsonl"), &records)?;
    write_jsonl(
        root.join("ownership-variants.jsonl"),
        &records
            .iter()
            .filter(|record| record.group == "ownership")
            .collect::<Vec<_>>(),
    )?;
    write(root.join("unresolved.jsonl"), unresolved_jsonl(&records))?;
    render(&root)
}

fn schedule(seed: u64, repeats: u32, mode_filter: Option<Mode>) -> Vec<ScheduleEntry> {
    let modes: Vec<Mode> = match mode_filter {
        Some(mode) => vec![mode],
        None => vec![Mode::L, Mode::S, Mode::X],
    };
    let mut candidates = Vec::new();
    for repetition in 1..=repeats {
        for mode in &modes {
            for case in all_cases() {
                candidates.push((case, *mode, repetition));
            }
        }
    }
    let mut state = seed;
    for index in (1..candidates.len()).rev() {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        let swap = (state as usize) % (index + 1);
        candidates.swap(index, swap);
    }
    candidates
        .into_iter()
        .enumerate()
        .map(|(index, (case, mode, repetition))| ScheduleEntry {
            run_id: format!("run-{seed}-{index:04}"),
            scenario: case.id.clone(),
            group: case.group.to_owned(),
            prefix_id: case.id.clone(),
            mode: mode.id().to_owned(),
            repetition,
            order_index: u32::try_from(index + 1).expect("schedule index fits u32"),
            sequence_hash: case_hash(&case),
        })
        .collect()
}

fn run_case_command(options: Options) -> Result<(), String> {
    let scenario = options
        .scenario
        .as_deref()
        .ok_or_else(|| "--scenario is required".to_owned())?;
    let mode = options
        .mode
        .ok_or_else(|| "--mode is required".to_owned())?;
    let run_id = options
        .run_id
        .as_deref()
        .ok_or_else(|| "--run-id is required".to_owned())?;
    let order = options
        .order
        .ok_or_else(|| "--order is required".to_owned())?;
    let case = case_by_id(scenario)?;
    let record = live_case(&case, mode, run_id, order, options.fixture.as_deref())?;
    print!(
        "{}",
        serde_json::to_string(&record).map_err(|error| error.to_string())?
    );
    Ok(())
}

fn recovery(options: Options) -> Result<(), String> {
    let root = options.root()?.to_path_buf();
    let fixture = options
        .fixture
        .as_deref()
        .ok_or_else(|| "--fixture is required for recovery".to_owned())?;
    let mut records = Vec::new();
    for (index, mode) in [Mode::L, Mode::S, Mode::X].into_iter().enumerate() {
        let record = live_recovery(
            mode,
            &format!("recovery-{}", mode.id()),
            u32::try_from(index + 1).expect("small"),
            fixture,
        )?;
        if record.cleanup.get("owned_process_exited") != Some(&Value::Bool(true)) {
            return Err(format!(
                "recovery {} did not confirm owned-process exit",
                mode.id()
            ));
        }
        records.push(record);
    }
    write_jsonl(root.join("repairs.jsonl"), &records)?;
    render(&root)
}

fn live_case(
    case: &CaseSpec,
    mode: Mode,
    run_id: &str,
    order: u32,
    fixture: Option<&Path>,
) -> Result<RunRecord, String> {
    #[cfg(windows)]
    {
        windows_live::run(case, mode, run_id, order, fixture, false)
    }
    #[cfg(not(windows))]
    {
        let _ = (case, mode, run_id, order, fixture);
        Err("Windows is required for a live delta run".to_owned())
    }
}

fn live_recovery(
    mode: Mode,
    run_id: &str,
    order: u32,
    fixture: &Path,
) -> Result<RunRecord, String> {
    #[cfg(windows)]
    {
        windows_live::run(
            &prefix_case(PREFIX_OPERATIONS.len()),
            mode,
            run_id,
            order,
            Some(fixture),
            true,
        )
    }
    #[cfg(not(windows))]
    {
        let _ = (mode, run_id, order, fixture);
        Err("Windows is required for recovery validation".to_owned())
    }
}

fn check(root: &Path) -> Result<(), String> {
    for name in [
        "SOURCE_MANIFEST.toml",
        "environments.jsonl",
        "prefixes.jsonl",
        "operation-definitions.jsonl",
        "run-schedule.jsonl",
        "observations.jsonl",
        "ownership-variants.jsonl",
        "native-process-context.jsonl",
        "cold-session-baseline.jsonl",
        "session-state-transitions.jsonl",
        "owned-process-cleanup.jsonl",
        "repairs.jsonl",
        "unresolved.jsonl",
    ] {
        let path = root.join(name);
        if !path.is_file() {
            return Err(format!("missing evidence file: {}", path.display()));
        }
        reject_sensitive(&fs::read_to_string(&path).map_err(io_error)?, &path)?;
    }
    for name in report_names() {
        let path = report_root(root).join(name);
        if !path.is_file() {
            return Err(format!("missing report: {}", path.display()));
        }
        let text = fs::read_to_string(&path).map_err(io_error)?;
        reject_sensitive(&text, &path)?;
        if text.contains("\r\n") || !text.ends_with('\n') {
            return Err(format!(
                "report {} must have LF endings and a final newline",
                path.display()
            ));
        }
    }
    let expected = report_contents(root)?;
    for (name, content) in expected {
        let actual = fs::read_to_string(report_root(root).join(name)).map_err(io_error)?;
        if actual != content {
            return Err(format!("generated report {name} is stale; run render"));
        }
    }
    Ok(())
}

fn reject_sensitive(text: &str, path: &Path) -> Result<(), String> {
    if text.contains("C:\\")
        || text.contains("\\\\")
        || text.contains("0x0000_")
        || text.contains("ptr=")
        || text.contains("\"pid\"")
        || text.contains("\"hwnd\"")
    {
        return Err(format!(
            "sensitive local identity persisted in {}",
            path.display()
        ));
    }
    Ok(())
}

fn render(root: &Path) -> Result<(), String> {
    for (name, content) in report_contents(root)? {
        write(report_root(root).join(name), content)?;
    }
    Ok(())
}

fn report_names() -> &'static [&'static str] {
    &[
        "prefix-results.md",
        "first-failing-prefix.md",
        "reversibility.md",
        "ownership-comparison.md",
        "property-read-effects.md",
        "process-instrumentation-effects.md",
        "type-info-effects.md",
        "minimal-vs-full-trace.md",
        "native-host-context.md",
        "cold-session-baseline.md",
        "excel-session-state-transition.md",
        "owned-process-cleanup.md",
        "cold-vs-warm-session-comparison.md",
        "typelib-validation-status.md",
        "repair-validation.md",
        "root-cause.md",
        "remaining-blockers.md",
    ]
}

fn report_contents(root: &Path) -> Result<BTreeMap<&'static str, String>, String> {
    let records = read_records(root.join("observations.jsonl"))?;
    let recoveries = read_records(root.join("repairs.jsonl"))?;
    let mut reports = BTreeMap::new();
    reports.insert(
        "prefix-results.md",
        aggregate_report("Prefix results", &records, Some("prefix")),
    );
    reports.insert("first-failing-prefix.md", first_failing_report(&records));
    reports.insert("reversibility.md", reversibility_report(&records));
    reports.insert(
        "ownership-comparison.md",
        aggregate_report("Ownership variants", &records, Some("ownership")),
    );
    reports.insert(
        "property-read-effects.md",
        aggregate_report("Property-read effects", &records, Some("property")),
    );
    reports.insert(
        "process-instrumentation-effects.md",
        aggregate_report(
            "Process instrumentation effects",
            &records,
            Some("process-instrumentation"),
        ),
    );
    reports.insert(
        "type-info-effects.md",
        aggregate_report("Type-information effects", &records, Some("type-info")),
    );
    reports.insert("minimal-vs-full-trace.md", trace_diff_report(&records));
    reports.insert("native-host-context.md", native_host_report(root)?);
    reports.insert("cold-session-baseline.md", cold_session_report(root)?);
    reports.insert(
        "excel-session-state-transition.md",
        session_transition_report(root)?,
    );
    reports.insert("owned-process-cleanup.md", cleanup_report(root)?);
    reports.insert(
        "cold-vs-warm-session-comparison.md",
        cold_warm_comparison_report(root)?,
    );
    reports.insert("typelib-validation-status.md", typelib_status_report());
    reports.insert("repair-validation.md", recovery_report(&recoveries));
    reports.insert("root-cause.md", root_cause_report(&records));
    reports.insert(
        "remaining-blockers.md",
        remaining_blockers_report(&records, &recoveries),
    );
    Ok(reports)
}

fn read_records(path: PathBuf) -> Result<Vec<RunRecord>, String> {
    let text = fs::read_to_string(&path).map_err(io_error)?;
    text.lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            serde_json::from_str(line).map_err(|error| format!("{}: {error}", path.display()))
        })
        .collect()
}

fn aggregate_report(title: &str, records: &[RunRecord], group: Option<&str>) -> String {
    let mut rows: BTreeMap<(String, String), (usize, usize, usize)> = BTreeMap::new();
    for record in records
        .iter()
        .filter(|record| group.is_none_or(|value| record.group == value))
    {
        let row = rows
            .entry((record.scenario.clone(), record.mode.clone()))
            .or_default();
        row.0 += 1;
        if record.add_hresult == "0x00000000" {
            row.1 += 1;
        } else {
            row.2 += 1;
        }
    }
    let mut output = format!(
        "# {title}\n\nGenerated deterministically from redacted fresh-process observations.\n\n| Scenario | Mode | Runs | Add success | Add failure |\n| --- | --- | ---: | ---: | ---: |\n"
    );
    if rows.is_empty() {
        output.push_str("| Not run | — | 0 | 0 | 0 |\n");
    } else {
        for ((scenario, mode), (runs, success, failure)) in rows {
            output.push_str(&format!(
                "| `{scenario}` | `{mode}` | {runs} | {success} | {failure} |\n"
            ));
        }
    }
    output
}

fn first_failing_report(records: &[RunRecord]) -> String {
    let prefixes: Vec<_> = records
        .iter()
        .filter(|record| record.group == "prefix" && record.mode == "L")
        .collect();
    let mut by_prefix: BTreeMap<String, (usize, usize)> = BTreeMap::new();
    for record in prefixes {
        let entry = by_prefix.entry(record.prefix_id.clone()).or_default();
        entry.0 += 1;
        entry.1 += usize::from(record.add_hresult != "0x00000000");
    }
    let first = by_prefix.iter().find(|(_, (_, failures))| *failures > 0);
    match first {
        Some((prefix, (runs, failures))) => format!("# First failing prefix\n\n`{prefix}` has {failures} Add failures in {runs} local/0x0400 fresh-process runs. Causality requires the separately recorded reversibility rows.\n"),
        None if by_prefix.is_empty() => "# First failing prefix\n\nNot run.\n".to_owned(),
        None => "# First failing prefix\n\nNo failing local/0x0400 prefix was observed. The prior full-harness failure therefore did not reproduce in this matrix and is classified as changed/non-deterministic rather than causal.\n".to_owned(),
    }
}

fn reversibility_report(records: &[RunRecord]) -> String {
    if records.iter().any(|record| {
        record.group == "prefix" && record.mode == "L" && record.add_hresult != "0x00000000"
    }) {
        "# Reversibility\n\nA failing prefix was observed; use the recorded adjacent-prefix, removal, and relocation scenarios before asserting causality.\n".to_owned()
    } else {
        "# Reversibility\n\nNo first failing prefix was available in the current fresh-process local/0x0400 matrix. Removal and relocation cannot establish a cause without a reproducible failure; ownership and supplemental one-variable controls remain recorded separately.\n".to_owned()
    }
}

fn trace_diff_report(records: &[RunRecord]) -> String {
    let minimal = records
        .iter()
        .find(|record| record.scenario == "A0" && record.mode == "L");
    let full = records.iter().find(|record| {
        record.scenario == format!("A{}", PREFIX_OPERATIONS.len()) && record.mode == "L"
    });
    let mut output = "# Minimal versus full call-trace diff\n\nTrace records omit pointer, HWND, PID, and local-path values.\n\n".to_owned();
    match (minimal, full) {
        (Some(minimal), Some(full)) => {
            output.push_str("| Sequence | Trace events | Add HRESULT |\n| --- | ---: | --- |\n");
            output.push_str(&format!(
                "| A0 | {} | `{}` |\n",
                minimal.trace.len(),
                minimal.add_hresult
            ));
            output.push_str(&format!(
                "| Full | {} | `{}` |\n\n",
                full.trace.len(),
                full.add_hresult
            ));
            let base: BTreeSet<_> = minimal
                .trace
                .iter()
                .map(|event| format!("{}:{}", event.interface, event.member))
                .collect();
            output.push_str("Additional full-sequence operations:\n\n");
            for event in &full.trace {
                let label = format!("{}:{}", event.interface, event.member);
                if !base.contains(&label) {
                    output.push_str(&format!("- `{label}` — `{}`\n", event.hresult));
                }
            }
        }
        _ => output.push_str("Live trace comparison not run.\n"),
    }
    output
}

fn native_host_report(root: &Path) -> Result<String, String> {
    let text = fs::read_to_string(root.join("native-process-context.jsonl")).map_err(io_error)?;
    if text.trim().is_empty() {
        Ok("# Native direct versus shim host context\n\nNot yet recorded by this tool. The Prompt 05E discrepancy remains unresolved and is not used to classify the high-level Rust sequence.\n".to_owned())
    } else {
        Ok(format!(
            "# Native direct versus shim host context\n\n{text}"
        ))
    }
}

fn read_json_values(path: PathBuf) -> Result<Vec<Value>, String> {
    let text = fs::read_to_string(&path).map_err(io_error)?;
    text.lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            serde_json::from_str(line).map_err(|error| format!("{}: {error}", path.display()))
        })
        .collect()
}

fn cold_session_report(root: &Path) -> Result<String, String> {
    let rows = read_json_values(root.join("cold-session-baseline.jsonl"))?;
    let mut output = "# Cold-session baseline\n\nThe records below are separate from warm-session and prefix observations. No raw process identity, local path, account name, HWND, or pointer is persisted.\n\n| Order | Control | Add | Owned process exited | Session clean |\n| ---: | --- | --- | --- | --- |\n".to_owned();
    if rows.is_empty() {
        output.push_str("| — | Not run | — | — | — |\n");
    } else {
        for row in rows {
            output.push_str(&format!(
                "| {} | `{}` | `{}` | `{}` | `{}` |\n",
                value_text(&row, "execution_order_index"),
                value_text(&row, "control"),
                value_text(&row, "workbooks_add"),
                value_text(&row, "owned_process_exited"),
                value_text(&row, "session_clean_after_run"),
            ));
        }
    }
    output.push_str("\nInterpretation: a successful full high-level local/0x0400 cold control is **cold-session success; prior session contamination is a credible hypothesis**, not a code repair.\n");
    Ok(output)
}

fn session_transition_report(root: &Path) -> Result<String, String> {
    let rows = read_json_values(root.join("session-state-transitions.jsonl"))?;
    let mut output = "# Excel session-state transition\n\n| State | Classification | Minimal high-level | Full high-level | Cleanup |\n| --- | --- | --- | --- | --- |\n".to_owned();
    if rows.is_empty() {
        output.push_str("| Not run | — | — | — | — |\n");
    } else {
        for row in rows {
            output.push_str(&format!(
                "| `{}` | {} | `{}` | `{}` | `{}` |\n",
                value_text(&row, "state"),
                value_text(&row, "classification"),
                value_text(&row, "minimal_high_level"),
                value_text(&row, "full_high_level"),
                value_text(&row, "owned_process_cleanup"),
            ));
        }
    }
    Ok(output)
}

fn cleanup_report(root: &Path) -> Result<String, String> {
    let rows = read_json_values(root.join("owned-process-cleanup.jsonl"))?;
    let mut output = "# Owned-process cleanup\n\n| Control | Exit bucket | Owned process exited | Forced termination |\n| --- | --- | --- | --- |\n".to_owned();
    if rows.is_empty() {
        output.push_str("| Not run | — | — | — |\n");
    } else {
        for row in rows {
            output.push_str(&format!(
                "| `{}` | `{}` | `{}` | `{}` |\n",
                value_text(&row, "control"),
                value_text(&row, "shutdown_duration_bucket"),
                value_text(&row, "owned_process_exited"),
                value_text(&row, "forced_termination"),
            ));
        }
    }
    Ok(output)
}

fn cold_warm_comparison_report(root: &Path) -> Result<String, String> {
    let cold = read_json_values(root.join("cold-session-baseline.jsonl"))?;
    let warm = read_json_values(root.join("session-state-transitions.jsonl"))?;
    Ok(format!(
        "# Cold versus warm-session comparison\n\nCold-session control records: {}. State-transition records: {}. Results remain separated by environment classification and are not collapsed into one aggregate success rate.\n",
        cold.len(),
        warm.len()
    ))
}

fn value_text(value: &Value, key: &str) -> String {
    match value.get(key) {
        Some(Value::String(value)) => value.clone(),
        Some(Value::Bool(value)) => value.to_string(),
        Some(Value::Number(value)) => value.to_string(),
        Some(value) => value.to_string(),
        None => "not-recorded".to_owned(),
    }
}

fn typelib_status_report() -> String {
    "# Typelib validation status\n\nThe ordinary `excel-com-typelib-audit check` deliberately uses `not-recorded` environment inputs and therefore reports the historical `typelib/SOURCE_MANIFEST.toml` stale. The companion `check-historical` command reads only the committed historical environment labels, then re-inspects the current registered typelib without writing any evidence. This preserves the 05B–05E artifact while validating its deterministic type-library content.\n".to_owned()
}

fn recovery_report(records: &[RunRecord]) -> String {
    let mut output = "# Repair and recovery validation\n\n| Mode | Add | Open | Range smoke | Cleanup |\n| --- | --- | --- | --- | --- |\n".to_owned();
    if records.is_empty() {
        output.push_str("| Not run | — | — | — | — |\n");
    } else {
        for record in records {
            let open = record
                .operations
                .get("recovery-open")
                .and_then(Value::as_str)
                .unwrap_or("not-run");
            let smoke = record
                .operations
                .get("range-smoke")
                .and_then(Value::as_str)
                .unwrap_or("not-run");
            let cleanup = record
                .cleanup
                .get("owned_process_exited")
                .and_then(Value::as_bool)
                .unwrap_or(false);
            output.push_str(&format!(
                "| `{}` | `{}` | `{open}` | `{smoke}` | `{cleanup}` |\n",
                record.mode, record.add_hresult
            ));
        }
    }
    output
}

fn root_cause_report(records: &[RunRecord]) -> String {
    let failure_count = records
        .iter()
        .filter(|record| {
            record.group == "prefix" && record.mode == "L" && record.add_hresult != "0x00000000"
        })
        .count();
    if failure_count == 0 {
        "# Root cause\n\n**Classification: non-deterministic Excel state / inconclusive.** The former full high-level local/0x0400 failure did not reproduce in the current isolated prefix matrix, so no individual operation, ownership transition, or storage behavior is causal. No repair is applied.\n".to_owned()
    } else {
        "# Root cause\n\n**Classification: pending reversibility.** A fresh-process prefix failure exists in the evidence; do not attribute it until removal and relocation controls reproduce it.\n".to_owned()
    }
}

fn remaining_blockers_report(records: &[RunRecord], recoveries: &[RunRecord]) -> String {
    let prefix_runs = records
        .iter()
        .filter(|record| record.group == "prefix")
        .count();
    format!(
        "# Remaining blockers\n\nPrefix observations: {prefix_runs}. Recovery observations: {}. No production API was introduced. The unresolved native-direct versus Rust-to-shim host-context discrepancy remains bounded separately. Prompt 05 Range work may resume only if the recovery rows show successful Add/Open and A1.Value2 smoke checks in every required mode.\n",
        recoveries.len()
    )
}

fn unresolved_jsonl(records: &[RunRecord]) -> String {
    let failed = records
        .iter()
        .filter(|record| {
            record.group == "prefix" && record.mode == "L" && record.add_hresult != "0x00000000"
        })
        .count();
    json!({
        "schema_version": SCHEMA_VERSION,
        "id": "05f.full-high-local-0400",
        "classification": if failed == 0 { "changed-baseline-non-deterministic" } else { "reproducible-prefix-failure-pending-reversibility" },
        "detail": if failed == 0 { "The previously reported full high-level failure passed during this fresh-process matrix." } else { "At least one current prefix failed; see reversibility report." },
    }).to_string() + "\n"
}

fn report_root(root: &Path) -> PathBuf {
    root.parent()
        .unwrap_or(root)
        .join("generated")
        .join("pre-add-delta")
}

fn write(path: impl AsRef<Path>, content: impl AsRef<[u8]>) -> Result<(), String> {
    let path = path.as_ref();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(io_error)?;
    }
    fs::write(path, content).map_err(io_error)
}

fn write_jsonl<T: Serialize>(path: PathBuf, values: &[T]) -> Result<(), String> {
    let mut text = String::new();
    for value in values {
        text.push_str(&serde_json::to_string(value).map_err(|error| error.to_string())?);
        text.push('\n');
    }
    write(path, text)
}

fn io_error(error: std::io::Error) -> String {
    error.to_string()
}

#[cfg(windows)]
mod windows_live {
    use super::*;
    use std::mem::ManuallyDrop;
    use windows::Win32::Foundation::{CloseHandle, HANDLE, HWND, VARIANT_BOOL, WAIT_OBJECT_0};
    use windows::Win32::System::Com::{
        CLSCTX_LOCAL_SERVER, CLSCTX_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance,
        CoCreateInstanceEx, CoInitializeEx, CoUninitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
        DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, IDispatch, MULTI_QI,
    };
    use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
    use windows::Win32::System::Threading::{
        GetProcessTimes, OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SYNCHRONIZE,
        WaitForSingleObject,
    };
    use windows::Win32::System::Variant::{
        VARIANT, VT_BSTR, VT_DISPATCH, VT_I4, VariantClear, VariantInit,
    };
    use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;
    use windows::core::{BSTR, GUID, HSTRING, IUnknown, Interface, PCWSTR};

    struct VariantOwner(VARIANT);

    impl VariantOwner {
        fn empty() -> Self {
            Self(unsafe { VariantInit() })
        }
        fn vartype(&self) -> u16 {
            unsafe { self.0.Anonymous.Anonymous.vt.0 }
        }
        fn type_name(&self) -> String {
            format!("VT_{}", self.vartype())
        }
        fn dispatch(&self) -> Option<IDispatch> {
            if self.vartype() != VT_DISPATCH.0 {
                return None;
            }
            unsafe {
                let value = &self.0.Anonymous.Anonymous.Anonymous.pdispVal
                    as *const ManuallyDrop<Option<IDispatch>>
                    as *const Option<IDispatch>;
                (*value).clone()
            }
        }
        fn bstr(value: &str) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = &mut result.0.Anonymous.Anonymous;
                data.vt = VT_BSTR;
                data.Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(value));
            }
            result
        }
        fn i4(value: i32) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = &mut result.0.Anonymous.Anonymous;
                data.vt = VT_I4;
                data.Anonymous.lVal = value;
            }
            result
        }
        fn boolean(value: bool) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = &mut result.0.Anonymous.Anonymous;
                data.vt = windows::Win32::System::Variant::VT_BOOL;
                data.Anonymous.boolVal = VARIANT_BOOL(if value { -1 } else { 0 });
            }
            result
        }
        fn clear(&mut self) {
            let _ = unsafe { VariantClear(&mut self.0) };
        }
    }

    impl Drop for VariantOwner {
        fn drop(&mut self) {
            self.clear();
        }
    }

    struct WorkbookCarrier {
        result: VariantOwner,
        dispatch: IDispatch,
    }

    struct Trace {
        events: Vec<TraceEvent>,
        ordinal: u32,
    }

    impl Trace {
        fn call(
            &mut self,
            interface: &str,
            member: &str,
            dispid: Option<i32>,
            flags: &str,
            lcid: u32,
            args: u32,
            named: u32,
            vt: String,
            hresult: i32,
            ownership: &str,
            clear: &str,
            release: &str,
        ) {
            self.ordinal += 1;
            self.events.push(TraceEvent {
                ordinal: self.ordinal,
                interface: interface.to_owned(),
                member: member.to_owned(),
                dispid: dispid
                    .map(|value| value.to_string())
                    .unwrap_or_else(|| "not-applicable".to_owned()),
                flags: flags.to_owned(),
                lcid: format!("0x{lcid:04x}"),
                c_args: args,
                c_named_args: named,
                result_vartype: vt,
                hresult: hresult_text(hresult),
                ownership_transition: ownership.to_owned(),
                variant_clear_timing: clear.to_owned(),
                release_timing: release.to_owned(),
            });
        }
    }

    fn hresult(result: windows::core::Result<()>) -> i32 {
        result.map(|_| 0).unwrap_or_else(|error| error.code().0)
    }
    fn hresult_text(value: i32) -> String {
        format!("0x{:08x}", value as u32)
    }

    unsafe fn member_id(dispatch: &IDispatch, member: &str, lcid: u32) -> Result<i32, i32> {
        let text = HSTRING::from(member);
        let names = [PCWSTR(text.as_ptr())];
        let mut output = 0;
        unsafe { dispatch.GetIDsOfNames(&GUID::from_u128(0), names.as_ptr(), 1, lcid, &mut output) }
            .map(|_| output)
            .map_err(|error| error.code().0)
    }

    unsafe fn invoke(
        dispatch: &IDispatch,
        interface: &str,
        member: &str,
        flags: windows::Win32::System::Com::DISPATCH_FLAGS,
        flag_name: &str,
        lcid: u32,
        params: &DISPPARAMS,
        result: &mut VariantOwner,
        exception: &mut EXCEPINFO,
        arg_error: &mut u32,
        trace: &mut Trace,
        ownership: &str,
        clear: &str,
    ) -> i32 {
        let dispid = match unsafe { member_id(dispatch, member, lcid) } {
            Ok(value) => value,
            Err(value) => return value,
        };
        let status = hresult(unsafe {
            dispatch.Invoke(
                dispid,
                &GUID::from_u128(0),
                lcid,
                flags,
                params,
                Some(&mut result.0),
                Some(exception),
                Some(arg_error),
            )
        });
        trace.call(
            interface,
            member,
            Some(dispid),
            flag_name,
            lcid,
            params.cArgs,
            params.cNamedArgs,
            result.type_name(),
            status,
            ownership,
            clear,
            "scope-end",
        );
        status
    }

    unsafe fn get_workbooks(
        app: &IDispatch,
        lcid: u32,
        trace: &mut Trace,
    ) -> Result<WorkbookCarrier, i32> {
        let empty = DISPPARAMS::default();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        let status = unsafe {
            invoke(
                app,
                "Application",
                "Workbooks",
                DISPATCH_PROPERTYGET,
                "PROPERTYGET",
                lcid,
                &empty,
                &mut result,
                &mut exception,
                &mut argument_error,
                trace,
                "clone-dispatch",
                "policy-defined",
            )
        };
        if status != 0 {
            return Err(status);
        }
        let dispatch = result.dispatch().ok_or(0x8000_4005_u32 as i32)?;
        Ok(WorkbookCarrier { result, dispatch })
    }

    unsafe fn property(
        dispatch: &IDispatch,
        interface: &str,
        member: &str,
        lcid: u32,
        trace: &mut Trace,
    ) -> (i32, String) {
        let empty = DISPPARAMS::default();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        let status = unsafe {
            invoke(
                dispatch,
                interface,
                member,
                DISPATCH_PROPERTYGET,
                "PROPERTYGET",
                lcid,
                &empty,
                &mut result,
                &mut exception,
                &mut argument_error,
                trace,
                "result-owned",
                "immediate",
            )
        };
        (status, result.type_name())
    }

    unsafe fn type_info_count(
        dispatch: &IDispatch,
        interface: &str,
        lcid: u32,
        trace: &mut Trace,
    ) -> i32 {
        let outcome = unsafe { dispatch.GetTypeInfoCount() };
        let status = outcome
            .as_ref()
            .map(|_| 0)
            .unwrap_or_else(|error| error.code().0);
        trace.call(
            interface,
            "GetTypeInfoCount",
            None,
            "direct",
            lcid,
            0,
            0,
            "not-applicable".to_owned(),
            status,
            "none",
            "not-applicable",
            "immediate",
        );
        status
    }

    unsafe fn type_info(
        dispatch: &IDispatch,
        interface: &str,
        lcid: u32,
        trace: &mut Trace,
    ) -> i32 {
        let outcome = unsafe { dispatch.GetTypeInfo(0, lcid) };
        let status = outcome
            .as_ref()
            .map(|_| 0)
            .unwrap_or_else(|error| error.code().0);
        trace.call(
            interface,
            "GetTypeInfo",
            None,
            "direct",
            lcid,
            0,
            0,
            "ITypeInfo".to_owned(),
            status,
            "ITypeInfo-owned",
            "not-applicable",
            "immediate",
        );
        status
    }

    unsafe fn count(workbooks: &IDispatch, lcid: u32, trace: &mut Trace) -> (i32, String) {
        unsafe { property(workbooks, "Workbooks", "Count", lcid, trace) }
    }

    unsafe fn apply_operation(
        operation: &LiveOperation,
        app: &IDispatch,
        workbooks: &mut WorkbookCarrier,
        lcid: u32,
        trace: &mut Trace,
        values: &mut BTreeMap<String, Value>,
        process: &mut Option<HANDLE>,
    ) {
        let (status, vartype) = match operation {
            LiveOperation::Prefix(Operation::Version) => unsafe {
                property(app, "Application", "Version", lcid, trace)
            },
            LiveOperation::Prefix(Operation::WorkbooksTypeInfoCount) => (
                unsafe { type_info_count(&workbooks.dispatch, "Workbooks", lcid, trace) },
                "not-applicable".to_owned(),
            ),
            LiveOperation::Prefix(Operation::WorkbooksQueryIUnknown) => {
                let outcome = workbooks.dispatch.cast::<IUnknown>();
                let status = outcome
                    .as_ref()
                    .map(|_| 0)
                    .unwrap_or_else(|error| error.code().0);
                trace.call(
                    "Workbooks",
                    "QueryInterface(IUnknown)",
                    None,
                    "direct",
                    lcid,
                    0,
                    0,
                    "IUnknown".to_owned(),
                    status,
                    "IUnknown-owned",
                    "not-applicable",
                    "immediate",
                );
                (status, "IUnknown".to_owned())
            }
            LiveOperation::Prefix(Operation::WorkbooksQueryIDispatch) => {
                let outcome = workbooks
                    .dispatch
                    .cast::<IUnknown>()
                    .and_then(|value| value.cast::<IDispatch>());
                let status = outcome
                    .as_ref()
                    .map(|_| 0)
                    .unwrap_or_else(|error| error.code().0);
                trace.call(
                    "Workbooks",
                    "QueryInterface(IDispatch)",
                    None,
                    "direct",
                    lcid,
                    0,
                    0,
                    "IDispatch".to_owned(),
                    status,
                    "IDispatch-owned",
                    "not-applicable",
                    "immediate",
                );
                (status, "IDispatch".to_owned())
            }
            LiveOperation::Prefix(Operation::WorkbooksCount) => unsafe {
                count(&workbooks.dispatch, lcid, trace)
            },
            LiveOperation::Prefix(Operation::LifetimeCloneThenClear) => {
                let status = unsafe { lifetime(app, lcid, 1, trace) };
                (status, "not-applicable".to_owned())
            }
            LiveOperation::Prefix(Operation::LifetimeRetainThenClear) => {
                let status = unsafe { lifetime(app, lcid, 2, trace) };
                (status, "not-applicable".to_owned())
            }
            LiveOperation::Prefix(Operation::LifetimeQueryInterfaceThenClear) => {
                let status = unsafe { lifetime(app, lcid, 3, trace) };
                (status, "not-applicable".to_owned())
            }
            LiveOperation::Property(member) => unsafe {
                property(app, "Application", member, lcid, trace)
            },
            LiveOperation::ApplicationTypeInfoCount => (
                unsafe { type_info_count(app, "Application", lcid, trace) },
                "not-applicable".to_owned(),
            ),
            LiveOperation::ApplicationTypeInfo => (
                unsafe { type_info(app, "Application", lcid, trace) },
                "ITypeInfo".to_owned(),
            ),
            LiveOperation::WorkbooksTypeInfo => (
                unsafe { type_info(&workbooks.dispatch, "Workbooks", lcid, trace) },
                "ITypeInfo".to_owned(),
            ),
            LiveOperation::ReacquireWorkbooks => match unsafe { get_workbooks(app, lcid, trace) } {
                Ok(mut replacement) => {
                    replacement.result.clear();
                    *workbooks = replacement;
                    (0, "VT_DISPATCH".to_owned())
                }
                Err(error) => (error, "VT_EMPTY".to_owned()),
            },
            LiveOperation::ProcessPid => unsafe { process_pid(app, lcid, trace, process, false) },
            LiveOperation::ProcessStartTime => unsafe {
                process_pid(app, lcid, trace, process, true)
            },
            LiveOperation::ProcessHandle => unsafe { process_pid(app, lcid, trace, process, true) },
        };
        values.insert(
            operation.id().to_owned(),
            json!({"hresult": hresult_text(status), "result_vartype": vartype}),
        );
    }

    unsafe fn apply_before_workbooks(
        operation: &LiveOperation,
        app: &IDispatch,
        lcid: u32,
        trace: &mut Trace,
        values: &mut BTreeMap<String, Value>,
        process: &mut Option<HANDLE>,
    ) {
        let (status, vartype) = match operation {
            LiveOperation::Prefix(Operation::Version) | LiveOperation::Property("Version") => unsafe {
                property(app, "Application", "Version", lcid, trace)
            },
            LiveOperation::Property(member) => unsafe {
                property(app, "Application", member, lcid, trace)
            },
            LiveOperation::ApplicationTypeInfoCount => (
                unsafe { type_info_count(app, "Application", lcid, trace) },
                "not-applicable".to_owned(),
            ),
            LiveOperation::ApplicationTypeInfo => (
                unsafe { type_info(app, "Application", lcid, trace) },
                "ITypeInfo".to_owned(),
            ),
            LiveOperation::ProcessPid => unsafe { process_pid(app, lcid, trace, process, false) },
            LiveOperation::ProcessStartTime | LiveOperation::ProcessHandle => unsafe {
                process_pid(app, lcid, trace, process, true)
            },
            _ => (
                0x8000_4005_u32 as i32,
                "invalid-before-workbooks".to_owned(),
            ),
        };
        values.insert(
            operation.id().to_owned(),
            json!({"hresult": hresult_text(status), "result_vartype": vartype}),
        );
    }

    unsafe fn lifetime(app: &IDispatch, lcid: u32, kind: u32, trace: &mut Trace) -> i32 {
        let mut workbooks = match unsafe { get_workbooks(app, lcid, trace) } {
            Ok(value) => value,
            Err(error) => return error,
        };
        workbooks.result.clear();
        match kind {
            1 => unsafe { count(&workbooks.dispatch, lcid, trace).0 },
            2 => {
                let retained = workbooks.dispatch.clone();
                drop(workbooks);
                unsafe { count(&retained, lcid, trace).0 }
            }
            _ => match workbooks
                .dispatch
                .cast::<IUnknown>()
                .and_then(|value| value.cast::<IDispatch>())
            {
                Ok(value) => unsafe { count(&value, lcid, trace).0 },
                Err(error) => error.code().0,
            },
        }
    }

    unsafe fn process_pid(
        app: &IDispatch,
        lcid: u32,
        trace: &mut Trace,
        process: &mut Option<HANDLE>,
        start_time: bool,
    ) -> (i32, String) {
        let (property_hr, vt) = unsafe { property(app, "Application", "Hwnd", lcid, trace) };
        if property_hr != 0 {
            return (property_hr, vt);
        }
        let empty = DISPPARAMS::default();
        let mut hwnd_value = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        let invoke_hr = unsafe {
            invoke(
                app,
                "Application",
                "Hwnd",
                DISPATCH_PROPERTYGET,
                "PROPERTYGET",
                lcid,
                &empty,
                &mut hwnd_value,
                &mut exception,
                &mut argument_error,
                trace,
                "result-owned",
                "immediate",
            )
        };
        if invoke_hr != 0 || hwnd_value.vartype() != VT_I4.0 {
            return (
                if invoke_hr == 0 {
                    0x8000_4005_u32 as i32
                } else {
                    invoke_hr
                },
                hwnd_value.type_name(),
            );
        }
        let hwnd = unsafe { hwnd_value.0.Anonymous.Anonymous.Anonymous.lVal };
        let mut pid = 0;
        unsafe {
            GetWindowThreadProcessId(HWND(hwnd as *mut _), Some(&mut pid));
        }
        if pid == 0 {
            return (0x8000_4005_u32 as i32, "VT_I4".to_owned());
        }
        let handle = unsafe {
            OpenProcess(
                PROCESS_SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                false,
                pid,
            )
        };
        let Ok(handle) = handle else {
            return (handle.unwrap_err().code().0, "VT_I4".to_owned());
        };
        if start_time {
            let mut created = Default::default();
            let mut exited = Default::default();
            let mut kernel = Default::default();
            let mut user = Default::default();
            if unsafe { GetProcessTimes(handle, &mut created, &mut exited, &mut kernel, &mut user) }
                .is_err()
            {
                let _ = unsafe { CloseHandle(handle) };
                return (0x8000_4005_u32 as i32, "VT_I4".to_owned());
            }
        }
        if let Some(old) = process.replace(handle) {
            let _ = unsafe { CloseHandle(old) };
        }
        trace.call(
            "Process",
            if start_time {
                "GetProcessTimes"
            } else {
                "OpenProcess"
            },
            None,
            "Win32",
            lcid,
            0,
            0,
            "not-applicable".to_owned(),
            0,
            "process-handle-owned",
            "not-applicable",
            "cleanup",
        );
        (0, "VT_I4".to_owned())
    }

    unsafe fn close_workbook(workbook: &IDispatch, lcid: u32, trace: &mut Trace) -> i32 {
        let mut argument = VariantOwner::boolean(false);
        let params = DISPPARAMS {
            rgvarg: &mut argument.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        unsafe {
            invoke(
                workbook,
                "Workbook",
                "Close",
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                &params,
                &mut result,
                &mut exception,
                &mut arg_error,
                trace,
                "result-owned",
                "immediate",
            )
        }
    }

    pub(super) fn run(
        case: &CaseSpec,
        mode: Mode,
        run_id: &str,
        order: u32,
        fixture: Option<&Path>,
        recovery: bool,
    ) -> Result<RunRecord, String> {
        unsafe { run_inner(case, mode, run_id, order, fixture, recovery) }
    }

    unsafe fn run_inner(
        case: &CaseSpec,
        mode: Mode,
        run_id: &str,
        order: u32,
        fixture: Option<&Path>,
        recovery: bool,
    ) -> Result<RunRecord, String> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }
            .map_err(|error| error.to_string())?;
        let mut trace = Trace {
            events: Vec::new(),
            ordinal: 0,
        };
        let mut operations = BTreeMap::new();
        let mut process: Option<HANDLE> = None;
        let mut cleanup = BTreeMap::new();
        let lcid = if mode == Mode::L { 0x0400 } else { 0 };
        let app = unsafe { activate(mode) }?;
        for operation in &case.before_workbooks {
            unsafe {
                apply_before_workbooks(
                    operation,
                    &app,
                    lcid,
                    &mut trace,
                    &mut operations,
                    &mut process,
                );
            }
        }
        let mut workbooks =
            unsafe { get_workbooks(&app, lcid, &mut trace) }.map_err(hresult_text)?;
        if case.ownership != Ownership::W0RetainResult
            && case.ownership != Ownership::W1CloneThenClearAfter
        {
            workbooks.result.clear();
        }
        for operation in &case.after_workbooks {
            unsafe {
                apply_operation(
                    operation,
                    &app,
                    &mut workbooks,
                    lcid,
                    &mut trace,
                    &mut operations,
                    &mut process,
                );
            }
        }
        let (count_hr, count_vt) = unsafe { count(&workbooks.dispatch, lcid, &mut trace) };
        operations.insert(
            "workbooks-count-before-add".to_owned(),
            json!({"hresult": hresult_text(count_hr), "result_vartype": count_vt}),
        );
        let active_workbooks =
            unsafe { ownership_dispatch(case.ownership, &app, &mut workbooks, lcid, &mut trace) }?;
        let empty = DISPPARAMS::default();
        let mut add = VariantOwner::empty();
        let mut add_exception = EXCEPINFO::default();
        let mut add_arg = u32::MAX;
        let add_hr = unsafe {
            invoke(
                &active_workbooks,
                "Workbooks",
                "Add",
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                &empty,
                &mut add,
                &mut add_exception,
                &mut add_arg,
                &mut trace,
                "workbook-dispatch-cloned",
                "after-result-consumption",
            )
        };
        let mut created = false;
        if add_hr == 0 {
            if let Some(workbook) = add.dispatch() {
                created = unsafe { close_workbook(&workbook, lcid, &mut trace) } == 0;
            }
        }
        if case.ownership == Ownership::W1CloneThenClearAfter
            || case.ownership == Ownership::W0RetainResult
        {
            workbooks.result.clear();
        }
        if recovery && add_hr == 0 {
            let open = fixture
                .map(|path| unsafe { open_and_smoke(&active_workbooks, path, lcid, &mut trace) });
            match open {
                Some((open_hr, smoke)) => {
                    operations.insert(
                        "recovery-open".to_owned(),
                        Value::String(hresult_text(open_hr)),
                    );
                    operations.insert("range-smoke".to_owned(), Value::String(smoke));
                }
                None => {
                    operations.insert(
                        "recovery-open".to_owned(),
                        Value::String("not-run".to_owned()),
                    );
                    operations.insert(
                        "range-smoke".to_owned(),
                        Value::String("not-run".to_owned()),
                    );
                }
            }
        }
        if process.is_none() {
            let _ = unsafe { process_pid(&app, lcid, &mut trace, &mut process, true) };
        }
        let empty = DISPPARAMS::default();
        let mut quit = VariantOwner::empty();
        let mut quit_exception = EXCEPINFO::default();
        let mut quit_arg = u32::MAX;
        let quit_hr = unsafe {
            invoke(
                &app,
                "Application",
                "Quit",
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                &empty,
                &mut quit,
                &mut quit_exception,
                &mut quit_arg,
                &mut trace,
                "result-owned",
                "immediate",
            )
        };
        drop(workbooks);
        drop(app);
        let exited = match process {
            Some(handle) => {
                let exited = unsafe { WaitForSingleObject(handle, 15_000) == WAIT_OBJECT_0 };
                let _ = unsafe { CloseHandle(handle) };
                exited
            }
            None => false,
        };
        unsafe { CoUninitialize() };
        cleanup.insert(
            "quit_hresult".to_owned(),
            Value::String(hresult_text(quit_hr)),
        );
        cleanup.insert("owned_process_exited".to_owned(), Value::Bool(exited));
        cleanup.insert("forced_termination".to_owned(), Value::Bool(false));
        Ok(RunRecord {
            schema_version: SCHEMA_VERSION,
            run_id: run_id.to_owned(),
            scenario: case.id.clone(),
            group: case.group.to_owned(),
            prefix_id: case.id.clone(),
            mode: mode.id().to_owned(),
            sequence_hash: case_hash(case),
            execution_order_index: order,
            process_architecture: "x64".to_owned(),
            thread_apartment: "STA".to_owned(),
            activation_api: mode.activation_api().to_owned(),
            clsctx: mode.clsctx().to_owned(),
            get_ids_of_names_lcid: mode.lcid().to_owned(),
            invoke_lcid: mode.lcid().to_owned(),
            excel_version: "16.0".to_owned(),
            workbooks_count_before_add: hresult_text(count_hr),
            operations,
            add_hresult: hresult_text(add_hr),
            excepinfo_scode: hresult_text(add_exception.scode),
            pu_arg_err_raw: add_arg.to_string(),
            workbook_created: created,
            cleanup,
            trace: trace.events,
        })
    }

    unsafe fn ownership_dispatch(
        ownership: Ownership,
        app: &IDispatch,
        workbooks: &mut WorkbookCarrier,
        lcid: u32,
        trace: &mut Trace,
    ) -> Result<IDispatch, String> {
        match ownership {
            Ownership::W0RetainResult
            | Ownership::W1CloneThenClearAfter
            | Ownership::W2CloneThenClearImmediately => Ok(workbooks.dispatch.clone()),
            Ownership::W3QueryIDispatchThenClear => workbooks
                .dispatch
                .cast::<IDispatch>()
                .map_err(|error| error.to_string()),
            Ownership::W4QueryIUnknownThenIDispatchThenClear => workbooks
                .dispatch
                .cast::<IUnknown>()
                .and_then(|value| value.cast::<IDispatch>())
                .map_err(|error| error.to_string()),
            Ownership::W5ReacquireBeforeAdd => {
                let mut replacement =
                    unsafe { get_workbooks(app, lcid, trace) }.map_err(hresult_text)?;
                replacement.result.clear();
                Ok(replacement.dispatch)
            }
        }
    }

    unsafe fn open_and_smoke(
        workbooks: &IDispatch,
        fixture: &Path,
        lcid: u32,
        trace: &mut Trace,
    ) -> (i32, String) {
        let fixture_text = fixture.to_string_lossy();
        let mut argument = VariantOwner::bstr(&fixture_text);
        let params = DISPPARAMS {
            rgvarg: &mut argument.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let mut output = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let open = unsafe {
            invoke(
                workbooks,
                "Workbooks",
                "Open",
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                &params,
                &mut output,
                &mut exception,
                &mut arg_error,
                trace,
                "workbook-dispatch-cloned",
                "after-result-consumption",
            )
        };
        if open != 0 {
            return (open, "not-run".to_owned());
        }
        let Some(workbook) = output.dispatch() else {
            return (0x8000_4005_u32 as i32, "not-run".to_owned());
        };
        let smoke = unsafe { range_smoke(&workbook, lcid, trace) };
        let _ = unsafe { close_workbook(&workbook, lcid, trace) };
        (open, smoke)
    }

    unsafe fn range_smoke(workbook: &IDispatch, lcid: u32, trace: &mut Trace) -> String {
        let worksheets =
            match unsafe { get_object(workbook, "Workbook", "Worksheets", lcid, trace) } {
                Ok(value) => value,
                Err(_) => return "failed-worksheets".to_owned(),
            };
        let mut index = VariantOwner::i4(1);
        let params = DISPPARAMS {
            rgvarg: &mut index.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let worksheet =
            match unsafe { method_object(&worksheets, "Worksheets", "Item", lcid, &params, trace) }
            {
                Ok(value) => value,
                Err(_) => return "failed-worksheet-item".to_owned(),
            };
        let mut address = VariantOwner::bstr("A1");
        let params = DISPPARAMS {
            rgvarg: &mut address.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let range = match unsafe {
            get_object_with_params(&worksheet, "Worksheet", "Range", lcid, &params, trace)
        } {
            Ok(value) => value,
            Err(_) => return "failed-range".to_owned(),
        };
        let mut value = VariantOwner::i4(42);
        let mut named = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: &mut value.0,
            rgdispidNamedArgs: &mut named,
            cArgs: 1,
            cNamedArgs: 1,
        };
        let mut put = VariantOwner::empty();
        let mut put_exception = EXCEPINFO::default();
        let mut put_arg = u32::MAX;
        if unsafe {
            invoke(
                &range,
                "Range",
                "Value2",
                DISPATCH_PROPERTYPUT,
                "PROPERTYPUT",
                lcid,
                &params,
                &mut put,
                &mut put_exception,
                &mut put_arg,
                trace,
                "result-owned",
                "immediate",
            )
        } != 0
        {
            return "failed-write".to_owned();
        }
        let (get_hr, get_type) = unsafe { property(&range, "Range", "Value2", lcid, trace) };
        if get_hr != 0 || get_type != format!("VT_{}", VT_I4.0) {
            return "failed-read".to_owned();
        }
        let empty = DISPPARAMS::default();
        let mut clear = VariantOwner::empty();
        let mut clear_exception = EXCEPINFO::default();
        let mut clear_arg = u32::MAX;
        if unsafe {
            invoke(
                &range,
                "Range",
                "ClearContents",
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                &empty,
                &mut clear,
                &mut clear_exception,
                &mut clear_arg,
                trace,
                "result-owned",
                "immediate",
            )
        } != 0
        {
            return "failed-clear".to_owned();
        }
        "passed-numeric-42".to_owned()
    }

    unsafe fn get_object(
        dispatch: &IDispatch,
        interface: &str,
        member: &str,
        lcid: u32,
        trace: &mut Trace,
    ) -> Result<IDispatch, i32> {
        let empty = DISPPARAMS::default();
        unsafe { get_object_with_params(dispatch, interface, member, lcid, &empty, trace) }
    }
    unsafe fn get_object_with_params(
        dispatch: &IDispatch,
        interface: &str,
        member: &str,
        lcid: u32,
        params: &DISPPARAMS,
        trace: &mut Trace,
    ) -> Result<IDispatch, i32> {
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let status = unsafe {
            invoke(
                dispatch,
                interface,
                member,
                DISPATCH_PROPERTYGET,
                "PROPERTYGET",
                lcid,
                params,
                &mut result,
                &mut exception,
                &mut arg_error,
                trace,
                "clone-dispatch",
                "immediate",
            )
        };
        if status != 0 {
            Err(status)
        } else {
            result.dispatch().ok_or(0x8000_4005_u32 as i32)
        }
    }
    unsafe fn method_object(
        dispatch: &IDispatch,
        interface: &str,
        member: &str,
        lcid: u32,
        params: &DISPPARAMS,
        trace: &mut Trace,
    ) -> Result<IDispatch, i32> {
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let status = unsafe {
            invoke(
                dispatch,
                interface,
                member,
                DISPATCH_METHOD,
                "METHOD",
                lcid,
                params,
                &mut result,
                &mut exception,
                &mut arg_error,
                trace,
                "clone-dispatch",
                "immediate",
            )
        };
        if status != 0 {
            Err(status)
        } else {
            result.dispatch().ok_or(0x8000_4005_u32 as i32)
        }
    }

    unsafe fn activate(mode: Mode) -> Result<IDispatch, String> {
        let progid = HSTRING::from("Excel.Application");
        let clsid = unsafe { windows::Win32::System::Com::CLSIDFromProgID(&progid) }
            .map_err(|error| error.to_string())?;
        if mode == Mode::X {
            let iid = IDispatch::IID;
            let mut interfaces = [MULTI_QI {
                pIID: &iid,
                pItf: ManuallyDrop::new(None),
                hr: windows::core::HRESULT(0),
            }];
            unsafe {
                CoCreateInstanceEx(
                    &clsid,
                    None::<&IUnknown>,
                    CLSCTX_SERVER,
                    None,
                    &mut interfaces,
                )
            }
            .map_err(|error| error.to_string())?;
            interfaces[0].hr.ok().map_err(|error| error.to_string())?;
            unsafe { ManuallyDrop::take(&mut interfaces[0].pItf) }
                .ok_or_else(|| "activation returned no interface".to_owned())?
                .cast::<IDispatch>()
                .map_err(|error| error.to_string())
        } else {
            let context = if mode == Mode::L {
                CLSCTX_LOCAL_SERVER
            } else {
                CLSCTX_SERVER
            };
            unsafe { CoCreateInstance::<_, IDispatch>(&clsid, None, context) }
                .map_err(|error| error.to_string())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn prefix_ids_and_hashes_are_stable() {
        assert_eq!(prefix_case(0).id, "A0");
        assert_eq!(prefix_case(PREFIX_OPERATIONS.len()).id, "A8");
        assert_eq!(case_hash(&prefix_case(3)), case_hash(&prefix_case(3)));
        assert_ne!(case_hash(&prefix_case(2)), case_hash(&prefix_case(3)));
    }

    #[test]
    fn prefix_construction_adds_exactly_one_operation() {
        for length in 1..=PREFIX_OPERATIONS.len() {
            let previous = prefix_case(length - 1);
            let current = prefix_case(length);
            let previous_ops: Vec<_> = previous
                .before_workbooks
                .iter()
                .chain(previous.after_workbooks.iter())
                .map(LiveOperation::id)
                .collect();
            let current_ops: Vec<_> = current
                .before_workbooks
                .iter()
                .chain(current.after_workbooks.iter())
                .map(LiveOperation::id)
                .collect();
            assert_eq!(&current_ops[..previous_ops.len()], previous_ops.as_slice());
            assert_eq!(current_ops.len(), previous_ops.len() + 1);
        }
    }

    #[test]
    fn schedule_is_seeded_and_complete() {
        let first = schedule(SCHEDULE_SEED, 2, None);
        let second = schedule(SCHEDULE_SEED, 2, None);
        assert_eq!(
            serde_json::to_string(&first).unwrap(),
            serde_json::to_string(&second).unwrap()
        );
        assert_eq!(first.len(), all_cases().len() * 3 * 2);
        assert!(
            first
                .iter()
                .any(|entry| entry.scenario == "A8" && entry.mode == "L")
        );
    }

    #[test]
    fn sensitive_identity_is_rejected() {
        assert!(reject_sensitive("C:\\\\Users", Path::new("evidence.jsonl")).is_err());
        assert!(reject_sensitive("{\"pid\":123}", Path::new("evidence.jsonl")).is_err());
        assert!(reject_sensitive("portable", Path::new("evidence.jsonl")).is_ok());
    }

    #[test]
    fn aggregate_is_deterministic_without_live_excel() {
        assert!(aggregate_report("Prefix results", &[], Some("prefix")).contains("Not run"));
        assert!(first_failing_report(&[]).contains("Not run"));
    }
}
