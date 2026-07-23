use std::ffi::c_void;
use std::marker::PhantomData;
use std::rc::Rc;
use std::sync::atomic::{AtomicU32, Ordering};
use std::thread::ThreadId;

use windows_sys::Win32::Media::Audio::CoRegisterMessageFilter;
use windows_sys::Win32::System::Com::INTERFACEINFO;
use windows_sys::core::{GUID, IUnknown_Vtbl};

use super::retry::{ComRetryPolicy, replace_active_policy};
use crate::error::ExcelRuntimeError;
use crate::internal::{ComPtr, Unknown};
use crate::{ComApartment, ExcelComError};

const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_c000_000000000046);
const IID_IMESSAGE_FILTER: GUID = GUID::from_u128(0x00000016_0000_0000_c000_000000000046);
const E_NOINTERFACE: i32 = 0x8000_4002_u32 as i32;
const E_POINTER: i32 = 0x8000_4003_u32 as i32;
const S_OK: i32 = 0;
const PENDINGMSG_WAITDEFPROCESS: u32 = 2;
const SERVERCALL_ISHANDLED: u32 = 0;
const RETRY_CANCEL_CALL: u32 = u32::MAX;

/// Temporarily installs the crate's conservative COM message filter on one STA.
///
/// The filter never replays calls on COM's behalf. Safe retries are performed
/// by the private dispatch layer, where the member kind is known. Dropping this
/// guard restores the prior filter best-effort; use [`Self::restore`] when a
/// registration-restoration failure must be observed.
pub struct ComMessageFilterGuard {
    _filter: Box<MessageFilter>,
    previous_filter: Option<ComPtr<Unknown>>,
    previous_policy: Option<ComRetryPolicy>,
    thread: ThreadId,
    restored: bool,
    _not_send_or_sync: PhantomData<Rc<()>>,
}

impl std::fmt::Debug for ComMessageFilterGuard {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ComMessageFilterGuard")
            .field("restored", &self.restored)
            .finish_non_exhaustive()
    }
}

impl ComMessageFilterGuard {
    /// Registers a filter and makes `policy` available to safe dispatch retries
    /// on `apartment`'s current STA thread.
    pub fn install(
        apartment: &ComApartment,
        policy: ComRetryPolicy,
    ) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        let mut filter = Box::new(MessageFilter::new());
        let mut previous = std::ptr::null_mut();
        // SAFETY: `filter` stays pinned in this guard until it is unregistered;
        // output storage accepts the previous IUnknown-compatible filter pointer.
        let status = unsafe {
            CoRegisterMessageFilter((&mut *filter as *mut MessageFilter).cast(), &mut previous)
        };
        if ExcelComError::failed(status) {
            return Err(ExcelComError::Runtime(
                ExcelRuntimeError::MessageFilterRegistrationFailed { hresult: status },
            ));
        }
        // SAFETY: successful registration transfers one owned reference for a
        // non-null previous filter, as documented by CoRegisterMessageFilter.
        let previous_filter = if previous.is_null() {
            None
        } else {
            // SAFETY: successful registration transferred one owned reference
            // for this non-null previous IUnknown-compatible filter pointer.
            Some(unsafe { ComPtr::from_owned(previous) }?)
        };
        Ok(Self {
            _filter: filter,
            previous_filter,
            previous_policy: replace_active_policy(Some(policy)),
            thread: std::thread::current().id(),
            restored: false,
            _not_send_or_sync: PhantomData,
        })
    }

    /// Restores the filter that was active before [`Self::install`].
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.restore_inner()?;
        Ok(())
    }

    fn restore_inner(&mut self) -> Result<(), ExcelComError> {
        if self.restored {
            return Ok(());
        }
        if self.thread != std::thread::current().id() {
            return Err(ExcelComError::Ownership {
                detail: "COM message filter restored from a different thread",
            });
        }
        let previous = self
            .previous_filter
            .as_ref()
            .map_or(std::ptr::null_mut(), ComPtr::raw);
        // SAFETY: the prior filter pointer, if any, remains owned by this
        // guard until after the replacement call completes.
        let status = unsafe { CoRegisterMessageFilter(previous, std::ptr::null_mut()) };
        if ExcelComError::failed(status) {
            return Err(ExcelComError::Runtime(
                ExcelRuntimeError::MessageFilterRegistrationFailed { hresult: status },
            ));
        }
        replace_active_policy(self.previous_policy.take());
        self.restored = true;
        Ok(())
    }
}

impl Drop for ComMessageFilterGuard {
    fn drop(&mut self) {
        let _ = self.restore_inner();
    }
}

#[repr(C)]
struct MessageFilterVtbl {
    base: IUnknown_Vtbl,
    handle_incoming_call:
        unsafe extern "system" fn(*mut c_void, u32, *mut c_void, u32, *const INTERFACEINFO) -> u32,
    retry_rejected_call: unsafe extern "system" fn(*mut c_void, *mut c_void, u32, u32) -> u32,
    message_pending: unsafe extern "system" fn(*mut c_void, *mut c_void, u32, u32) -> u32,
}

#[repr(C)]
struct MessageFilter {
    vtbl: *const MessageFilterVtbl,
    references: AtomicU32,
}

impl MessageFilter {
    fn new() -> Self {
        Self {
            vtbl: &MESSAGE_FILTER_VTBL,
            references: AtomicU32::new(1),
        }
    }
}

static MESSAGE_FILTER_VTBL: MessageFilterVtbl = MessageFilterVtbl {
    base: IUnknown_Vtbl {
        QueryInterface: message_filter_query_interface,
        AddRef: message_filter_add_ref,
        Release: message_filter_release,
    },
    handle_incoming_call: message_filter_handle_incoming_call,
    retry_rejected_call: message_filter_retry_rejected_call,
    message_pending: message_filter_message_pending,
};

unsafe extern "system" fn message_filter_query_interface(
    this: *mut c_void,
    iid: *const GUID,
    interface: *mut *mut c_void,
) -> i32 {
    if iid.is_null() || interface.is_null() {
        return E_POINTER;
    }
    // SAFETY: validated non-null pointers are supplied by COM.
    if unsafe { guid_eq(&*iid, &IID_IUNKNOWN) || guid_eq(&*iid, &IID_IMESSAGE_FILTER) } {
        // SAFETY: `interface` is valid output storage for the COM caller.
        unsafe { *interface = this };
        // SAFETY: COM supplied our own instance pointer.
        unsafe { message_filter_add_ref(this) };
        S_OK
    } else {
        // SAFETY: `interface` is valid output storage for the COM caller.
        unsafe { *interface = std::ptr::null_mut() };
        E_NOINTERFACE
    }
}

fn guid_eq(left: &GUID, right: &GUID) -> bool {
    left.data1 == right.data1
        && left.data2 == right.data2
        && left.data3 == right.data3
        && left.data4 == right.data4
}

unsafe extern "system" fn message_filter_add_ref(this: *mut c_void) -> u32 {
    // SAFETY: COM invokes the vtable only for the MessageFilter object.
    unsafe {
        (&*(this.cast::<MessageFilter>()))
            .references
            .fetch_add(1, Ordering::Relaxed)
            + 1
    }
}

unsafe extern "system" fn message_filter_release(this: *mut c_void) -> u32 {
    // The guard owns allocation lifetime. COM reference accounting is retained
    // only to honor IUnknown; it must not free the object under the guard.
    // SAFETY: COM invokes the vtable only for the MessageFilter object.
    unsafe {
        (&*(this.cast::<MessageFilter>()))
            .references
            .fetch_sub(1, Ordering::Relaxed)
            .saturating_sub(1)
    }
}

unsafe extern "system" fn message_filter_handle_incoming_call(
    _this: *mut c_void,
    _call_type: u32,
    _caller: *mut c_void,
    _tick_count: u32,
    _interface_info: *const INTERFACEINFO,
) -> u32 {
    SERVERCALL_ISHANDLED
}

unsafe extern "system" fn message_filter_retry_rejected_call(
    _this: *mut c_void,
    _callee: *mut c_void,
    _tick_count: u32,
    _reject_type: u32,
) -> u32 {
    // Member safety is not available in IMessageFilter. Returning the cancel
    // sentinel makes COM report the original HRESULT to the typed dispatch
    // layer, which retries only reads and idempotent property puts.
    RETRY_CANCEL_CALL
}

unsafe extern "system" fn message_filter_message_pending(
    _this: *mut c_void,
    _callee: *mut c_void,
    _tick_count: u32,
    _pending_type: u32,
) -> u32 {
    PENDINGMSG_WAITDEFPROCESS
}
