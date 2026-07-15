//! Bounded diagnostic recording for COM entry points.
//!
//! # Safety contract
//! `GetCurrentThreadId` takes no pointers. `CoGetApartmentType` receives valid
//! stack out-pointers and its HRESULT is observed rather than assumed.

use crate::model::ServerPhase;
use std::fs::OpenOptions;
use std::io::Write;
use std::sync::Mutex;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};
use windows::Win32::System::Com::{APTTYPE, APTTYPEQUALIFIER, CoGetApartmentType};
use windows::Win32::System::Threading::{GetCurrentProcessId, GetCurrentThreadId};

const MAX_EVENTS: u64 = 4_096;
static EVENT_SEQUENCE: AtomicU64 = AtomicU64::new(0);
static FILE_LOCK: Mutex<()> = Mutex::new(());

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(crate) struct ResourceCounters {
    pub objects: u32,
    pub class_locks: u32,
    pub servers: u32,
    pub producers: u32,
    pub callback_cookies: u32,
    pub notification_calls: u32,
}

pub(crate) fn record(
    method: &str,
    phase: ServerPhase,
    result: i32,
    counters: ResourceCounters,
    detail: &str,
) {
    let sequence = EVENT_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    if sequence >= MAX_EVENTS || !valid_label(method) {
        return;
    }
    let Some(path) = std::env::var_os("EXCEL_API_MINIMAL_RTD_DIAGNOSTICS") else {
        return;
    };
    if !valid_detail(detail) {
        return;
    }
    let process_id = unsafe { GetCurrentProcessId() };
    let thread_id = unsafe { GetCurrentThreadId() };
    let timestamp_ms = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |duration| duration.as_millis());
    let mut apartment = APTTYPE::default();
    let mut qualifier = APTTYPEQUALIFIER::default();
    let apartment_result = unsafe { CoGetApartmentType(&mut apartment, &mut qualifier) };
    let (apartment, qualifier) = if apartment_result.is_ok() {
        (apartment.0, qualifier.0)
    } else {
        (-1, -1)
    };
    let Ok(_guard) = FILE_LOCK.lock() else {
        return;
    };
    let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) else {
        return;
    };
    let _ = writeln!(
        file,
        "{{\"sequence\":{sequence},\"timestamp_ms\":{timestamp_ms},\"process_id\":{process_id},\"thread_id\":{thread_id},\"apartment\":{apartment},\"qualifier\":{qualifier},\"method\":\"{method}\",\"detail\":\"{detail}\",\"phase\":\"{}\",\"result\":{result},\"counters\":{{\"objects\":{},\"class_locks\":{},\"servers\":{},\"producers\":{},\"callback_cookies\":{},\"notification_calls\":{}}}}}",
        phase_name(phase),
        counters.objects,
        counters.class_locks,
        counters.servers,
        counters.producers,
        counters.callback_cookies,
        counters.notification_calls,
    );
}

fn valid_detail(detail: &str) -> bool {
    detail.len() <= 160
        && detail.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(byte, b'_' | b'-' | b':' | b'{' | b'}' | b'.' | b' ')
        })
}

fn valid_label(label: &str) -> bool {
    !label.is_empty()
        && label.len() <= 48
        && label
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'_')
}

fn phase_name(phase: ServerPhase) -> &'static str {
    match phase {
        ServerPhase::Created => "Created",
        ServerPhase::Started => "Started",
        ServerPhase::Active => "Active",
        ServerPhase::Stopping => "Stopping",
        ServerPhase::CallbackRevocationPending => "CallbackRevocationPending",
        ServerPhase::Terminated => "Terminated",
    }
}
