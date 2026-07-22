use std::ffi::c_void;

use windows_sys::Win32::Foundation::SysFreeString;
use windows_sys::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, DISPPARAMS, EXCEPINFO,
};
use windows_sys::core::GUID;

use super::{MemberDescriptor, OwnedVariant};
use crate::{
    ExcelComError,
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
        args.reverse();
    }
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
    if ExcelComError::failed(status) {
        return Err(ExcelComError::Invocation {
            member,
            dispid,
            hresult: status,
            exception_scode: (scode != 0).then_some(scode),
            argument_index: (argument != u32::MAX).then_some(argument),
        });
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    #[test]
    fn positional_argument_order_is_com_reverse_order() {
        let mut values = [1, 2, 3];
        values.reverse();
        assert_eq!(values, [3, 2, 1]);
    }
}
