//! Raw in-process COM class server for the compatibility prototype.
//!
//! # Safety contract
//! COM supplies pointers under the audited Office 1.9 ABI. Entries validate
//! nullable pointers, initialize outs, contain panics, and never unwind across
//! COM. Objects have a vtable first and are reclaimed only by atomic final
//! Release; interface returns AddRef once. Only GIT cookies cross threads, and
//! worker COM initialization brackets every retrieved proxy. Tests preserve
//! locally constructed object lifetimes around each raw call.

use super::abi::{
    CLSID_MINIMAL_RTD, EXCEL_TYPELIB, EXCEL_TYPELIB_MAJOR, EXCEL_TYPELIB_MINOR, IID_IRTD_SERVER,
    IRtdServer_Vtbl, IRtdUpdateEvent,
};
use super::automation::{OwnedSafeArray, initial_counter, read_topic_components};
use super::diagnostics;
use crate::model::{ModelError, RefreshItem, ServerModel, ServerPhase};
use core::ffi::c_void;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::ptr::null_mut;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex, MutexGuard};
use std::thread::{self, JoinHandle};
use std::time::Duration;
use windows::Win32::Foundation::{
    CLASS_E_CLASSNOTAVAILABLE, CLASS_E_NOAGGREGATION, E_FAIL, E_INVALIDARG, E_NOINTERFACE,
    E_OUTOFMEMORY, E_POINTER, E_UNEXPECTED, S_FALSE, S_OK, VARIANT_BOOL,
};
use windows::Win32::System::Com::{
    CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
    DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, IClassFactory, IClassFactory_Vtbl, IDispatch,
    IDispatch_Vtbl, IGlobalInterfaceTable, ITypeInfo, SAFEARRAY,
};
use windows::Win32::System::Ole::{DispGetIDsOfNames, DispInvoke, LoadRegTypeLib};
use windows::Win32::System::Variant::VARIANT;
use windows::core::{BOOL, GUID, HRESULT, IUnknown, IUnknown_Vtbl, Interface};

const CLSID_STD_GLOBAL_INTERFACE_TABLE: GUID =
    GUID::from_u128(0x00000323_0000_0000_c000_000000000046);
const PRODUCER_PERIOD: Duration = Duration::from_millis(500);
const PRODUCER_TICKS: usize = 120;

static ACTIVE_OBJECTS: AtomicU32 = AtomicU32::new(0);
static SERVER_LOCKS: AtomicU32 = AtomicU32::new(0);
static ACTIVE_SERVERS: AtomicU32 = AtomicU32::new(0);
static ACTIVE_PRODUCERS: AtomicU32 = AtomicU32::new(0);
static CALLBACK_COOKIES: AtomicU32 = AtomicU32::new(0);
static NOTIFICATION_CALLS: AtomicU32 = AtomicU32::new(0);
static NEXT_GENERATION: AtomicU64 = AtomicU64::new(1);

struct Shared {
    model: Mutex<ServerModel>,
    wake: Condvar,
}

impl Shared {
    fn new() -> Self {
        Self {
            model: Mutex::new(ServerModel::new(
                NEXT_GENERATION.fetch_add(1, Ordering::Relaxed),
            )),
            wake: Condvar::new(),
        }
    }

    fn lock(&self) -> MutexGuard<'_, ServerModel> {
        self.model.lock().unwrap_or_else(|error| error.into_inner())
    }
}

#[repr(C)]
struct FactoryObject {
    vtable: *const IClassFactory_Vtbl,
    refs: AtomicU32,
}

#[repr(C)]
struct RtdObject {
    vtable: *const IRtdServer_Vtbl,
    refs: AtomicU32,
    shared: Arc<Shared>,
    producer: Mutex<Option<JoinHandle<()>>>,
    callback_cookie: AtomicU32,
    active_started: AtomicBool,
}

impl RtdObject {
    fn new() -> Box<Self> {
        ACTIVE_OBJECTS.fetch_add(1, Ordering::AcqRel);
        Box::new(Self {
            vtable: &RTD_VTABLE,
            refs: AtomicU32::new(1),
            shared: Arc::new(Shared::new()),
            producer: Mutex::new(None),
            callback_cookie: AtomicU32::new(0),
            active_started: AtomicBool::new(false),
        })
    }

    fn phase(&self) -> ServerPhase {
        self.shared.lock().phase()
    }

    fn start(&self, callback: *mut c_void) -> HRESULT {
        if callback.is_null() {
            return E_POINTER;
        }
        {
            let mut model = self.shared.lock();
            if let Err(error) = model.start() {
                return model_hresult(error);
            }
        }

        let cookie = match register_callback(callback) {
            Ok(cookie) => cookie,
            Err(code) => {
                self.shared.lock().rollback_start();
                return code;
            }
        };
        self.callback_cookie.store(cookie, Ordering::Release);
        CALLBACK_COOKIES.fetch_add(1, Ordering::AcqRel);

        let (ready_tx, ready_rx) = std::sync::mpsc::sync_channel(1);
        let shared = Arc::clone(&self.shared);
        let spawn = thread::Builder::new()
            .name("excel-api-rtd-producer".into())
            .spawn(move || producer_main(shared, cookie, ready_tx));
        let handle = match spawn {
            Ok(handle) => handle,
            Err(_) => {
                let revoke = revoke_callback(cookie);
                self.callback_cookie.store(0, Ordering::Release);
                if revoke.is_ok() {
                    CALLBACK_COOKIES.fetch_sub(1, Ordering::AcqRel);
                }
                self.shared.lock().rollback_start();
                return revoke.map_or_else(|code| code, |()| E_OUTOFMEMORY);
            }
        };
        let ready = ready_rx.recv_timeout(Duration::from_secs(5));
        if !matches!(ready, Ok(Ok(()))) {
            {
                let mut model = self.shared.lock();
                model.begin_stop();
            }
            self.shared.wake.notify_all();
            let _ = handle.join();
            let revoke = revoke_callback(cookie);
            self.callback_cookie.store(0, Ordering::Release);
            if revoke.is_ok() {
                CALLBACK_COOKIES.fetch_sub(1, Ordering::AcqRel);
            }
            let mut model = self.shared.lock();
            model.rollback_failed_start();
            return E_FAIL;
        }
        *self
            .producer
            .lock()
            .unwrap_or_else(|error| error.into_inner()) = Some(handle);
        self.active_started.store(true, Ordering::Release);
        ACTIVE_SERVERS.fetch_add(1, Ordering::AcqRel);
        S_OK
    }

    fn terminate(&self) -> HRESULT {
        let should_stop = self.shared.lock().begin_stop();
        if !should_stop {
            return S_OK;
        }
        self.shared.wake.notify_all();
        let handle = self
            .producer
            .lock()
            .unwrap_or_else(|error| error.into_inner())
            .take();
        if let Some(handle) = handle {
            let _ = handle.join();
        }

        let cookie = self.callback_cookie.swap(0, Ordering::AcqRel);
        let revoke = if cookie == 0 {
            Ok(())
        } else {
            revoke_callback(cookie)
        };
        if cookie != 0 && revoke.is_ok() {
            CALLBACK_COOKIES.fetch_sub(1, Ordering::AcqRel);
        }
        self.shared.lock().finish_stop();
        if self.active_started.swap(false, Ordering::AcqRel) {
            ACTIVE_SERVERS.fetch_sub(1, Ordering::AcqRel);
        }
        revoke.map_or_else(|code| code, |()| S_OK)
    }
}

impl Drop for RtdObject {
    fn drop(&mut self) {
        let _ = self.terminate();
        ACTIVE_OBJECTS.fetch_sub(1, Ordering::AcqRel);
    }
}

struct ComApartment;

impl ComApartment {
    fn initialize_mta() -> Result<Self, HRESULT> {
        let result = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };
        result.ok().map_err(|error| error.code())?;
        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

fn create_git() -> Result<IGlobalInterfaceTable, HRESULT> {
    unsafe {
        CoCreateInstance(
            &CLSID_STD_GLOBAL_INTERFACE_TABLE,
            None::<&IUnknown>,
            CLSCTX_INPROC_SERVER,
        )
    }
    .map_err(|error| error.code())
}

fn register_callback(callback: *mut c_void) -> Result<u32, HRESULT> {
    let git = create_git()?;
    let borrowed = unsafe { IRtdUpdateEvent::from_raw_borrowed(&callback) }.ok_or(E_POINTER)?;
    let unknown = borrowed.cast::<IUnknown>().map_err(|error| error.code())?;
    unsafe { git.RegisterInterfaceInGlobal(&unknown, &super::abi::IID_IRTD_UPDATE_EVENT) }
        .map_err(|error| error.code())
}

fn revoke_callback(cookie: u32) -> Result<(), HRESULT> {
    let git = create_git()?;
    unsafe { git.RevokeInterfaceFromGlobal(cookie) }.map_err(|error| error.code())
}

trait NotificationSink {
    fn update_notify(&self) -> Result<(), HRESULT>;
}

impl NotificationSink for IRtdUpdateEvent {
    fn update_notify(&self) -> Result<(), HRESULT> {
        unsafe { self.update_notify() }.map_err(|error| error.code())
    }
}

fn publish_and_notify(shared: &Shared, callback: &impl NotificationSink) -> bool {
    let mut model = shared.lock();
    let notify = model.publish_counter_tick();
    let phase = model.phase();
    drop(model);
    if !notify {
        return false;
    }
    NOTIFICATION_CALLS.fetch_add(1, Ordering::AcqRel);
    diagnostics::record("UpdateNotify_begin", phase, 0);
    let result = callback.update_notify();
    NOTIFICATION_CALLS.fetch_sub(1, Ordering::AcqRel);
    let code = result.as_ref().map_or_else(|error| error.0, |()| 0);
    diagnostics::record("UpdateNotify_end", phase, code);
    if result.is_err() {
        shared.lock().notification_failed();
    }
    true
}

fn producer_main(
    shared: Arc<Shared>,
    cookie: u32,
    ready: std::sync::mpsc::SyncSender<Result<(), HRESULT>>,
) {
    ACTIVE_PRODUCERS.fetch_add(1, Ordering::AcqRel);
    let run = || -> Result<(), HRESULT> {
        let _apartment = ComApartment::initialize_mta()?;
        let git = create_git()?;
        let mut raw = null_mut();
        unsafe { git.GetInterfaceFromGlobal(cookie, &super::abi::IID_IRTD_UPDATE_EVENT, &mut raw) }
            .map_err(|error| error.code())?;
        let callback = unsafe { IRtdUpdateEvent::from_raw(raw) };
        let _ = ready.send(Ok(()));
        for _ in 0..PRODUCER_TICKS {
            let guard = shared.lock();
            let (guard, _) = shared
                .wake
                .wait_timeout(guard, PRODUCER_PERIOD)
                .unwrap_or_else(|error| error.into_inner());
            if matches!(
                guard.phase(),
                ServerPhase::Stopping | ServerPhase::Terminated
            ) {
                break;
            }
            drop(guard);
            publish_and_notify(&shared, &callback);
        }
        Ok(())
    };
    let result = run();
    if result.is_err() {
        let _ = ready.send(result);
    }
    ACTIVE_PRODUCERS.fetch_sub(1, Ordering::AcqRel);
}

struct RefreshGuard<'a> {
    shared: &'a Shared,
    items: Vec<RefreshItem>,
    finished: bool,
}

impl<'a> RefreshGuard<'a> {
    fn begin(shared: &'a Shared) -> Result<Self, HRESULT> {
        let items = shared.lock().begin_refresh().map_err(model_hresult)?;
        Ok(Self {
            shared,
            items,
            finished: false,
        })
    }

    fn finish(mut self, succeeded: bool) {
        self.shared.lock().finish_refresh(&self.items, succeeded);
        self.finished = true;
    }
}

impl Drop for RefreshGuard<'_> {
    fn drop(&mut self) {
        if !self.finished {
            self.shared.lock().finish_refresh(&self.items, false);
        }
    }
}

fn model_hresult(error: ModelError) -> HRESULT {
    match error {
        ModelError::InvalidTopic | ModelError::DuplicateTopicId | ModelError::UnknownTopic => {
            E_INVALIDARG
        }
        ModelError::TopicLimit => E_OUTOFMEMORY,
        ModelError::AlreadyStarted
        | ModelError::InvalidState
        | ModelError::RefreshAlreadyRunning => E_FAIL,
    }
}

fn type_info() -> Result<ITypeInfo, HRESULT> {
    let library =
        unsafe { LoadRegTypeLib(&EXCEL_TYPELIB, EXCEL_TYPELIB_MAJOR, EXCEL_TYPELIB_MINOR, 0) }
            .map_err(|error| error.code())?;
    unsafe { library.GetTypeInfoOfGuid(&IID_IRTD_SERVER) }.map_err(|error| error.code())
}

fn catch_hresult(action: impl FnOnce() -> HRESULT) -> HRESULT {
    catch_unwind(AssertUnwindSafe(action)).unwrap_or(E_UNEXPECTED)
}

fn add_reference(refs: &AtomicU32) -> u32 {
    match refs.fetch_update(Ordering::Relaxed, Ordering::Relaxed, |value| {
        value.checked_add(1)
    }) {
        Ok(previous) => previous + 1,
        Err(current) => current,
    }
}

fn release_reference(refs: &AtomicU32) -> Option<u32> {
    refs.fetch_update(Ordering::Release, Ordering::Relaxed, |value| {
        value.checked_sub(1)
    })
    .ok()
    .map(|previous| previous - 1)
}

unsafe extern "system" fn factory_query_interface(
    this: *mut c_void,
    iid: *const GUID,
    output: *mut *mut c_void,
) -> HRESULT {
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    catch_hresult(|| {
        if this.is_null() || iid.is_null() || output.is_null() {
            return E_POINTER;
        }
        let iid = unsafe { &*iid };
        if iid != &IUnknown::IID && iid != &IClassFactory::IID {
            return E_NOINTERFACE;
        }
        let object = unsafe { &*(this.cast::<FactoryObject>()) };
        add_reference(&object.refs);
        unsafe { output.write(this) };
        S_OK
    })
}

unsafe extern "system" fn factory_add_ref(this: *mut c_void) -> u32 {
    catch_unwind(AssertUnwindSafe(|| {
        if this.is_null() {
            0
        } else {
            add_reference(&unsafe { &*(this.cast::<FactoryObject>()) }.refs)
        }
    }))
    .unwrap_or(0)
}

unsafe extern "system" fn factory_release(this: *mut c_void) -> u32 {
    catch_unwind(AssertUnwindSafe(|| {
        if this.is_null() {
            return 0;
        }
        let object = unsafe { &*(this.cast::<FactoryObject>()) };
        let Some(remaining) = release_reference(&object.refs) else {
            return 0;
        };
        if remaining == 0 {
            std::sync::atomic::fence(Ordering::Acquire);
            ACTIVE_OBJECTS.fetch_sub(1, Ordering::AcqRel);
            unsafe { drop(Box::from_raw(this.cast::<FactoryObject>())) };
        }
        remaining
    }))
    .unwrap_or(0)
}

unsafe extern "system" fn factory_create_instance(
    _this: *mut c_void,
    outer: *mut c_void,
    iid: *const GUID,
    output: *mut *mut c_void,
) -> HRESULT {
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    catch_hresult(|| {
        if iid.is_null() || output.is_null() {
            return E_POINTER;
        }
        if !outer.is_null() {
            return CLASS_E_NOAGGREGATION;
        }
        let raw = Box::into_raw(RtdObject::new()).cast::<c_void>();
        let result = unsafe { rtd_query_interface(raw, iid, output) };
        unsafe { rtd_release(raw) };
        result
    })
}

unsafe extern "system" fn factory_lock_server(this: *mut c_void, lock: BOOL) -> HRESULT {
    if this.is_null() {
        return E_POINTER;
    }
    catch_hresult(|| {
        if lock.as_bool() {
            SERVER_LOCKS.fetch_add(1, Ordering::AcqRel);
        } else if SERVER_LOCKS
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |value| {
                value.checked_sub(1)
            })
            .is_err()
        {
            return E_UNEXPECTED;
        }
        S_OK
    })
}

unsafe extern "system" fn rtd_query_interface(
    this: *mut c_void,
    iid: *const GUID,
    output: *mut *mut c_void,
) -> HRESULT {
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    catch_hresult(|| {
        if this.is_null() || iid.is_null() || output.is_null() {
            return E_POINTER;
        }
        let iid = unsafe { &*iid };
        if iid != &IUnknown::IID && iid != &IDispatch::IID && iid != &IID_IRTD_SERVER {
            return E_NOINTERFACE;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        add_reference(&object.refs);
        unsafe { output.write(this) };
        S_OK
    })
}

unsafe extern "system" fn rtd_add_ref(this: *mut c_void) -> u32 {
    catch_unwind(AssertUnwindSafe(|| {
        if this.is_null() {
            0
        } else {
            add_reference(&unsafe { &*(this.cast::<RtdObject>()) }.refs)
        }
    }))
    .unwrap_or(0)
}

unsafe extern "system" fn rtd_release(this: *mut c_void) -> u32 {
    catch_unwind(AssertUnwindSafe(|| {
        if this.is_null() {
            return 0;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        let Some(remaining) = release_reference(&object.refs) else {
            return 0;
        };
        if remaining == 0 {
            std::sync::atomic::fence(Ordering::Acquire);
            unsafe { drop(Box::from_raw(this.cast::<RtdObject>())) };
        }
        remaining
    }))
    .unwrap_or(0)
}

unsafe extern "system" fn dispatch_get_type_info_count(
    _this: *mut c_void,
    count: *mut u32,
) -> HRESULT {
    if count.is_null() {
        E_POINTER
    } else {
        unsafe { count.write(1) };
        S_OK
    }
}

unsafe extern "system" fn dispatch_get_type_info(
    _this: *mut c_void,
    index: u32,
    _lcid: u32,
    output: *mut *mut c_void,
) -> HRESULT {
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    catch_hresult(|| {
        if output.is_null() {
            return E_POINTER;
        }
        if index != 0 {
            return E_INVALIDARG;
        }
        match type_info() {
            Ok(info) => {
                unsafe { output.write(info.into_raw()) };
                S_OK
            }
            Err(code) => code,
        }
    })
}

unsafe extern "system" fn dispatch_get_ids_of_names(
    _this: *mut c_void,
    iid: *const GUID,
    names: *const windows::core::PCWSTR,
    count: u32,
    _lcid: u32,
    dispids: *mut i32,
) -> HRESULT {
    catch_hresult(|| {
        if iid.is_null() || names.is_null() || dispids.is_null() {
            return E_POINTER;
        }
        if unsafe { *iid } != GUID::zeroed() {
            return E_INVALIDARG;
        }
        match type_info() {
            Ok(info) => unsafe { DispGetIDsOfNames(&info, names, count, dispids) }
                .map_or_else(|error| error.code(), |()| S_OK),
            Err(code) => code,
        }
    })
}

unsafe extern "system" fn dispatch_invoke(
    this: *mut c_void,
    dispid: i32,
    iid: *const GUID,
    _lcid: u32,
    flags: DISPATCH_FLAGS,
    params: *const DISPPARAMS,
    result: *mut VARIANT,
    exception: *mut EXCEPINFO,
    arg_error: *mut u32,
) -> HRESULT {
    catch_hresult(|| {
        if this.is_null() || iid.is_null() || params.is_null() {
            return E_POINTER;
        }
        if unsafe { *iid } != GUID::zeroed() {
            return E_INVALIDARG;
        }
        match type_info() {
            Ok(info) => unsafe {
                DispInvoke(
                    this,
                    &info,
                    dispid,
                    flags.0,
                    params.cast_mut(),
                    result,
                    exception,
                    arg_error,
                )
            }
            .map_or_else(|error| error.code(), |()| S_OK),
            Err(code) => code,
        }
    })
}

unsafe extern "system" fn rtd_server_start(
    this: *mut c_void,
    callback: *mut c_void,
    result: *mut i32,
) -> HRESULT {
    if !result.is_null() {
        unsafe { result.write(0) };
    }
    let code = catch_hresult(|| {
        if this.is_null() || result.is_null() {
            return E_POINTER;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        let code = object.start(callback);
        if code.is_ok() {
            unsafe { result.write(1) };
        }
        code
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("ServerStart", object.phase(), code.0);
    }
    code
}

unsafe extern "system" fn rtd_connect_data(
    this: *mut c_void,
    topic_id: i32,
    strings: *mut *mut SAFEARRAY,
    get_new_values: *mut VARIANT_BOOL,
    output: *mut VARIANT,
) -> HRESULT {
    if !get_new_values.is_null() {
        unsafe { get_new_values.write(VARIANT_BOOL(0)) };
    }
    if !output.is_null() {
        unsafe { output.write(VARIANT::default()) };
    }
    let code = catch_hresult(|| {
        if this.is_null() || get_new_values.is_null() || output.is_null() {
            return E_POINTER;
        }
        let components = match unsafe { read_topic_components(strings) } {
            Ok(components) => components,
            Err(code) => return code,
        };
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        match object.shared.lock().connect(topic_id, components) {
            Ok(initial) => {
                unsafe {
                    get_new_values.write(VARIANT_BOOL(-1));
                    output.write(initial_counter(initial));
                }
                object.shared.wake.notify_all();
                S_OK
            }
            Err(error) => model_hresult(error),
        }
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("ConnectData", object.phase(), code.0);
    }
    code
}

unsafe extern "system" fn rtd_refresh_data(
    this: *mut c_void,
    topic_count: *mut i32,
    output: *mut *mut SAFEARRAY,
) -> HRESULT {
    if !topic_count.is_null() {
        unsafe { topic_count.write(0) };
    }
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    let code = catch_hresult(|| {
        if this.is_null() || topic_count.is_null() || output.is_null() {
            return E_POINTER;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        let guard = match RefreshGuard::begin(&object.shared) {
            Ok(guard) => guard,
            Err(code) => return code,
        };
        let payload = match OwnedSafeArray::refresh_payload(&guard.items) {
            Ok(payload) => payload,
            Err(code) => return code,
        };
        let count = match i32::try_from(guard.items.len()) {
            Ok(count) => count,
            Err(_) => return E_UNEXPECTED,
        };
        unsafe {
            topic_count.write(count);
            output.write(payload.into_raw());
        }
        guard.finish(true);
        S_OK
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("RefreshData", object.phase(), code.0);
    }
    code
}

unsafe extern "system" fn rtd_disconnect_data(this: *mut c_void, topic_id: i32) -> HRESULT {
    let code = catch_hresult(|| {
        if this.is_null() {
            return E_POINTER;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        object
            .shared
            .lock()
            .disconnect(topic_id)
            .map_or_else(model_hresult, |()| S_OK)
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("DisconnectData", object.phase(), code.0);
    }
    code
}

unsafe extern "system" fn rtd_heartbeat(this: *mut c_void, result: *mut i32) -> HRESULT {
    if !result.is_null() {
        unsafe { result.write(0) };
    }
    let code = catch_hresult(|| {
        if this.is_null() || result.is_null() {
            return E_POINTER;
        }
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        unsafe { result.write(object.shared.lock().heartbeat()) };
        S_OK
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("Heartbeat", object.phase(), code.0);
    }
    code
}

unsafe extern "system" fn rtd_server_terminate(this: *mut c_void) -> HRESULT {
    let code = catch_hresult(|| {
        if this.is_null() {
            E_POINTER
        } else {
            unsafe { &*(this.cast::<RtdObject>()) }.terminate()
        }
    });
    if !this.is_null() {
        let object = unsafe { &*(this.cast::<RtdObject>()) };
        diagnostics::record("ServerTerminate", object.phase(), code.0);
    }
    code
}

static FACTORY_VTABLE: IClassFactory_Vtbl = IClassFactory_Vtbl {
    base__: IUnknown_Vtbl {
        QueryInterface: factory_query_interface,
        AddRef: factory_add_ref,
        Release: factory_release,
    },
    CreateInstance: factory_create_instance,
    LockServer: factory_lock_server,
};

static RTD_VTABLE: IRtdServer_Vtbl = IRtdServer_Vtbl {
    base__: IDispatch_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: rtd_query_interface,
            AddRef: rtd_add_ref,
            Release: rtd_release,
        },
        GetTypeInfoCount: dispatch_get_type_info_count,
        GetTypeInfo: dispatch_get_type_info,
        GetIDsOfNames: dispatch_get_ids_of_names,
        Invoke: dispatch_invoke,
    },
    server_start: rtd_server_start,
    connect_data: rtd_connect_data,
    refresh_data: rtd_refresh_data,
    disconnect_data: rtd_disconnect_data,
    heartbeat: rtd_heartbeat,
    server_terminate: rtd_server_terminate,
};

#[unsafe(no_mangle)]
pub unsafe extern "system" fn DllGetClassObject(
    clsid: *const GUID,
    iid: *const GUID,
    output: *mut *mut c_void,
) -> HRESULT {
    if !output.is_null() {
        unsafe { output.write(null_mut()) };
    }
    catch_hresult(|| {
        if clsid.is_null() || iid.is_null() || output.is_null() {
            return E_POINTER;
        }
        if unsafe { *clsid } != CLSID_MINIMAL_RTD {
            return CLASS_E_CLASSNOTAVAILABLE;
        }
        ACTIVE_OBJECTS.fetch_add(1, Ordering::AcqRel);
        let factory = Box::new(FactoryObject {
            vtable: &FACTORY_VTABLE,
            refs: AtomicU32::new(1),
        });
        let raw = Box::into_raw(factory).cast::<c_void>();
        let result = unsafe { factory_query_interface(raw, iid, output) };
        unsafe { factory_release(raw) };
        result
    })
}

#[unsafe(no_mangle)]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    let unloadable = ACTIVE_OBJECTS.load(Ordering::Acquire) == 0
        && SERVER_LOCKS.load(Ordering::Acquire) == 0
        && ACTIVE_SERVERS.load(Ordering::Acquire) == 0
        && ACTIVE_PRODUCERS.load(Ordering::Acquire) == 0
        && CALLBACK_COOKIES.load(Ordering::Acquire) == 0
        && NOTIFICATION_CALLS.load(Ordering::Acquire) == 0;
    if unloadable { S_OK } else { S_FALSE }
}

#[cfg(test)]
mod tests {
    use super::*;

    struct MockNotification<'a> {
        shared: &'a Shared,
        calls: AtomicU32,
        fail: AtomicBool,
    }

    impl NotificationSink for MockNotification<'_> {
        fn update_notify(&self) -> Result<(), HRESULT> {
            assert!(self.shared.model.try_lock().is_ok());
            self.calls.fetch_add(1, Ordering::AcqRel);
            if self.fail.load(Ordering::Acquire) {
                Err(E_FAIL)
            } else {
                Ok(())
            }
        }
    }

    unsafe fn class_factory() -> *mut c_void {
        let mut output = null_mut();
        assert_eq!(
            unsafe { DllGetClassObject(&CLSID_MINIMAL_RTD, &IClassFactory::IID, &mut output) },
            S_OK
        );
        output
    }

    #[test]
    fn class_factory_identity_aggregation_and_unload_are_exact() {
        assert_eq!(DllCanUnloadNow(), S_OK);
        let factory = unsafe { class_factory() };
        assert_eq!(DllCanUnloadNow(), S_FALSE);
        let mut server = null_mut();
        assert_eq!(
            unsafe { factory_create_instance(factory, null_mut(), &IID_IRTD_SERVER, &mut server) },
            S_OK
        );
        assert!(!server.is_null());
        let mut unsupported = null_mut();
        let unsupported_iid = GUID::from_u128(0x11111111_2222_3333_4444_555555555555);
        assert_eq!(
            unsafe { rtd_query_interface(server, &unsupported_iid, &mut unsupported) },
            E_NOINTERFACE
        );
        assert!(unsupported.is_null());
        let mut aggregated = null_mut();
        assert_eq!(
            unsafe { factory_create_instance(factory, factory, &IID_IRTD_SERVER, &mut aggregated) },
            CLASS_E_NOAGGREGATION
        );
        unsafe { rtd_release(server) };
        unsafe { factory_release(factory) };
        assert_eq!(DllCanUnloadNow(), S_OK);
    }

    #[test]
    fn lock_server_controls_unload_and_underflow_is_rejected() {
        let factory = unsafe { class_factory() };
        assert_eq!(unsafe { factory_lock_server(factory, BOOL(1)) }, S_OK);
        assert_eq!(DllCanUnloadNow(), S_FALSE);
        assert_eq!(unsafe { factory_lock_server(factory, BOOL(0)) }, S_OK);
        assert_eq!(
            unsafe { factory_lock_server(factory, BOOL(0)) },
            E_UNEXPECTED
        );
        assert_eq!(
            unsafe { factory_lock_server(null_mut(), BOOL(0)) },
            E_POINTER
        );
        unsafe { factory_release(factory) };
        assert_eq!(DllCanUnloadNow(), S_OK);
    }

    #[test]
    fn heartbeat_and_terminate_are_controlled_without_start() {
        let raw = Box::into_raw(RtdObject::new()).cast::<c_void>();
        let mut heartbeat = -1;
        assert_eq!(unsafe { rtd_heartbeat(raw, &mut heartbeat) }, S_OK);
        assert_eq!(heartbeat, 0);
        assert_eq!(unsafe { rtd_server_terminate(raw) }, S_OK);
        assert_eq!(unsafe { rtd_server_terminate(raw) }, S_OK);
        unsafe { rtd_release(raw) };
        assert_eq!(DllCanUnloadNow(), S_OK);
    }

    #[test]
    fn notification_is_lock_free_retryable_and_suppressed_after_stop() {
        let shared = Shared::new();
        {
            let mut model = shared.lock();
            model.start().unwrap();
            model
                .connect(1, vec!["COUNTER".encode_utf16().collect()])
                .unwrap();
        }
        let callback = MockNotification {
            shared: &shared,
            calls: AtomicU32::new(0),
            fail: AtomicBool::new(true),
        };
        assert!(publish_and_notify(&shared, &callback));
        callback.fail.store(false, Ordering::Release);
        assert!(publish_and_notify(&shared, &callback));
        assert_eq!(callback.calls.load(Ordering::Acquire), 2);
        assert!(shared.lock().begin_stop());
        assert!(!publish_and_notify(&shared, &callback));
        assert_eq!(callback.calls.load(Ordering::Acquire), 2);
        shared.lock().finish_stop();
    }

    #[test]
    fn class_lookup_and_null_outputs_fail_explicitly() {
        let other = GUID::from_u128(1);
        let mut output = null_mut();
        assert_eq!(
            unsafe { DllGetClassObject(&other, &IClassFactory::IID, &mut output) },
            CLASS_E_CLASSNOTAVAILABLE
        );
        assert_eq!(
            unsafe { DllGetClassObject(&CLSID_MINIMAL_RTD, &IClassFactory::IID, null_mut()) },
            E_POINTER
        );
    }
}
