use std::ffi::c_void;
use std::time::Instant;

use windows_sys::Win32::Foundation::{SysFreeString, SysStringLen};
use windows_sys::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, DISPPARAMS, EXCEPINFO,
};
use windows_sys::core::GUID;

use super::{
    InvocationRetrySafety, MemberDescriptor, MemberKind, OwnedVariant, active_policy,
    classify_com_hresult, reverse_for_com,
};
use crate::{
    ExcelComError, InvocationError,
    internal::{ComPtr, Dispatch, wide_nul},
};

const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
const LOCALE_USER_DEFAULT: u32 = 1024;
const LOCALE_SYSTEM_DEFAULT: u32 = 2048;
const DISPID_PROPERTYPUT: i32 = -3;

pub(crate) fn activate_excel() -> Result<ComPtr<Dispatch>, ExcelComError> {
    let name = wide_nul("Excel.Application");
    let mut class = GUID::default();
    // SAFETY: the NUL-terminated ProgID and output GUID are valid for this call.
    let status = unsafe { CLSIDFromProgID(name.as_ptr(), &mut class) };
    if ExcelComError::failed(status) {
        return Err(ExcelComError::Activation { hresult: status });
    }
    let mut raw: *mut c_void = std::ptr::null_mut();
    // SAFETY: class/IID/output storage are valid and activation requests a local server only.
    let status = unsafe {
        CoCreateInstance(
            &class,
            std::ptr::null_mut(),
            CLSCTX_LOCAL_SERVER,
            &IID_IDISPATCH,
            &mut raw,
        )
    };
    if ExcelComError::failed(status) {
        return Err(ExcelComError::Activation { hresult: status });
    }
    // SAFETY: successful CoCreateInstance returned one owned IDispatch reference.
    unsafe { ComPtr::from_owned(raw) }
}

pub(crate) fn property_get(
    target: &ComPtr<Dispatch>,
    descriptor: MemberDescriptor,
    args: Vec<OwnedVariant>,
) -> Result<OwnedVariant, ExcelComError> {
    invoke(target, descriptor, args, false)
}

pub(crate) fn property_put(
    target: &ComPtr<Dispatch>,
    descriptor: MemberDescriptor,
    value: OwnedVariant,
) -> Result<OwnedVariant, ExcelComError> {
    invoke(target, descriptor, vec![value], true)
}

pub(crate) fn invoke(
    target: &ComPtr<Dispatch>,
    descriptor: MemberDescriptor,
    mut args: Vec<OwnedVariant>,
    property_put: bool,
) -> Result<OwnedVariant, ExcelComError> {
    let _inventory_member = descriptor.id;
    let member = descriptor.name;
    let name = wide_nul(member);
    let names = [name.as_ptr()];
    let mut dispid = 0;
    // SAFETY: the vtable is validated by ComPtr and all lookup buffers outlive the call.
    let lookup = unsafe {
        (target.vtbl().get_ids_of_names)(
            target.raw(),
            &GUID::default(),
            names.as_ptr(),
            1,
            LOCALE_USER_DEFAULT,
            &mut dispid,
        )
    };
    if ExcelComError::failed(lookup) {
        return Err(ExcelComError::NameLookup {
            member,
            hresult: lookup,
        });
    }
    if !property_put {
        reverse_for_com(&mut args);
    }

    let retry_safety = retry_safety(descriptor, property_put);
    let policy = active_policy();
    let started = Instant::now();
    let mut attempts = 0;
    loop {
        attempts += 1;
        let mut named = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: if args.is_empty() {
                std::ptr::null_mut()
            } else {
                args.as_mut_ptr().cast()
            },
            rgdispidNamedArgs: if property_put {
                &mut named
            } else {
                std::ptr::null_mut()
            },
            cArgs: args.len() as u32,
            cNamedArgs: u32::from(property_put),
        };
        let mut result = OwnedVariant::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument = u32::MAX;
        // SAFETY: DISPPARAMS, result, EXCEPINFO, and argument-error storage remain valid through Invoke.
        let status = unsafe {
            (target.vtbl().invoke)(
                target.raw(),
                dispid,
                &GUID::default(),
                LOCALE_SYSTEM_DEFAULT,
                descriptor.kind.flags(),
                &params,
                &mut result.0,
                &mut exception,
                &mut argument,
            )
        };
        let exception_source = bstr_text(exception.bstrSource);
        let exception_description = bstr_text(exception.bstrDescription);
        let scode = exception.scode;
        // SAFETY: EXCEPINFO BSTR fields are owned by this call result and are released exactly once.
        unsafe {
            for value in [
                &mut exception.bstrSource,
                &mut exception.bstrDescription,
                &mut exception.bstrHelpFile,
            ] {
                if !(*value).is_null() {
                    SysFreeString(*value);
                    *value = std::ptr::null_mut();
                }
            }
        }
        if !ExcelComError::failed(status) {
            return Ok(result);
        }
        let elapsed = started.elapsed();
        let disposition = classify_com_hresult(status);
        let error = ExcelComError::Invocation(Box::new(InvocationError {
            object_type: object_type(descriptor.id.as_str()),
            member,
            dispid,
            hresult: status,
            exception_scode: (scode != 0).then_some(scode),
            argument_index: (argument != u32::MAX).then_some(argument),
            dispatch_flags: descriptor.kind.flags(),
            disposition,
            retry_safety,
            attempts,
            elapsed,
            exception_source,
            exception_description,
        }));
        let Some(policy) = policy.as_ref() else {
            return Err(error);
        };
        let safe = matches!(
            retry_safety,
            InvocationRetrySafety::SafeRead | InvocationRetrySafety::IdempotentWrite
        );
        let permitted = match disposition {
            super::ComCallDisposition::RetryableBusy => policy.retry_server_busy,
            super::ComCallDisposition::RetryableRejected => policy.retry_call_rejected,
            super::ComCallDisposition::PermanentFailure | super::ComCallDisposition::Unknown => {
                false
            }
        };
        let delay = policy.delay_for_retry(attempts);
        if !safe
            || !permitted
            || attempts >= policy.max_attempts.max(1)
            || elapsed.saturating_add(delay) > policy.total_timeout
        {
            return Err(error);
        }
        std::thread::sleep(delay);
    }
}

fn retry_safety(descriptor: MemberDescriptor, property_put: bool) -> InvocationRetrySafety {
    // The Excel typelib records some mutating members (notably Workbooks.Add)
    // as PROPERTYGET. Known object-creation and destructive descriptors must
    // therefore override the generic DISPATCH classification.
    if matches!(
        descriptor.id.as_str(),
        "excel.workbooks.add"
            | "excel.worksheets.add"
            | "excel.charts.add"
            | "excel.chartobjects.add"
            | "excel.seriescollection.add"
            | "excel.seriescollection.newseries"
            | "excel.workbookconnection.delete"
    ) || descriptor.name == "Delete"
    {
        return InvocationRetrySafety::NonIdempotentWrite;
    }
    match (descriptor.kind, property_put) {
        (MemberKind::PropertyGet, false) => InvocationRetrySafety::SafeRead,
        (MemberKind::PropertyPut | MemberKind::PropertyPutRef, true) => {
            InvocationRetrySafety::IdempotentWrite
        }
        (MemberKind::Method, _) => InvocationRetrySafety::NonIdempotentWrite,
        _ => InvocationRetrySafety::Unknown,
    }
}

fn bstr_text(value: *const u16) -> Option<String> {
    if value.is_null() {
        return None;
    }
    // SAFETY: the BSTR belongs to the current EXCEPINFO until SysFreeString;
    // SysStringLen gives the exact UTF-16 length including embedded NULs.
    let length = unsafe { SysStringLen(value) } as usize;
    // SAFETY: BSTR storage remains valid until the caller frees EXCEPINFO.
    let units = unsafe { std::slice::from_raw_parts(value, length) };
    Some(String::from_utf16_lossy(units))
}

fn object_type(id: &str) -> &'static str {
    if id.starts_with("excel.application.") {
        "Application"
    } else if id.starts_with("excel.workbooks.") {
        "Workbooks"
    } else if id.starts_with("excel.workbook.") {
        "Workbook"
    } else if id.starts_with("excel.worksheets.") {
        "Worksheets"
    } else if id.starts_with("excel.worksheet.") {
        "Worksheet"
    } else if id.starts_with("excel.range.") {
        "Range"
    } else if id.starts_with("excel.names.") {
        "Names"
    } else if id.starts_with("excel.name.") {
        "Name"
    } else if id.starts_with("excel.chartgroup.") {
        "ChartGroup"
    } else if id.starts_with("excel.point.") {
        "Point"
    } else {
        "IDispatch"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn positional_argument_order_is_com_reverse_order() {
        let mut values = [1, 2, 3];
        values.reverse();
        assert_eq!(values, [3, 2, 1]);
    }

    #[test]
    fn only_reads_and_property_puts_are_retry_safe() {
        assert_eq!(
            retry_safety(
                MemberDescriptor {
                    id: crate::MemberId::new("excel.application.version"),
                    name: "Version",
                    kind: MemberKind::PropertyGet,
                },
                false,
            ),
            InvocationRetrySafety::SafeRead
        );
        assert_eq!(
            retry_safety(
                MemberDescriptor {
                    id: crate::MemberId::new("excel.application.visible"),
                    name: "Visible",
                    kind: MemberKind::PropertyPut,
                },
                true,
            ),
            InvocationRetrySafety::IdempotentWrite
        );
        assert_eq!(
            retry_safety(
                MemberDescriptor {
                    id: crate::MemberId::new("excel.application.quit"),
                    name: "Quit",
                    kind: MemberKind::Method,
                },
                false,
            ),
            InvocationRetrySafety::NonIdempotentWrite
        );
    }

    #[test]
    fn workbooks_add_is_never_classified_as_a_read() {
        assert_eq!(
            retry_safety(
                MemberDescriptor {
                    id: crate::MemberId::new("excel.workbooks.add"),
                    name: "Add",
                    kind: MemberKind::PropertyGet,
                },
                false,
            ),
            InvocationRetrySafety::NonIdempotentWrite
        );
    }
}
