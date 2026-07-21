#include "excel_com_native_control.h"

#include <oaidl.h>
#include <oleauto.h>
#include <windows.h>

#include <cstddef>
#include <cstring>
#include <type_traits>

namespace {

constexpr HRESULT kNotTested = static_cast<HRESULT>(0x80004001L);
constexpr HRESULT kNoInterfaceResult = static_cast<HRESULT>(0x80004002L);

static_assert(sizeof(void*) == 8, "The native differential is x64-only.");
static_assert(sizeof(VARIANT) == 24, "x64 Windows VARIANT layout changed.");
static_assert(alignof(VARIANT) == 8, "x64 Windows VARIANT alignment changed.");
static_assert(sizeof(DISPPARAMS) == 24, "x64 Windows DISPPARAMS layout changed.");
static_assert(alignof(DISPPARAMS) == 8, "x64 Windows DISPPARAMS alignment changed.");
static_assert(sizeof(EXCEPINFO) == 64, "x64 Windows EXCEPINFO layout changed.");
static_assert(alignof(EXCEPINFO) == 8, "x64 Windows EXCEPINFO alignment changed.");
static_assert(sizeof(GUID) == 16, "Windows GUID layout changed.");
static_assert(std::is_standard_layout_v<ExcelRawResult>);
static_assert(std::is_standard_layout_v<ExcelAbiLayout>);

struct DispatchRef {
    IDispatch* value = nullptr;
    DispatchRef() = default;
    explicit DispatchRef(IDispatch* initial) : value(initial) {}
    DispatchRef(const DispatchRef&) = delete;
    DispatchRef& operator=(const DispatchRef&) = delete;
    DispatchRef(DispatchRef&& other) noexcept : value(other.value) { other.value = nullptr; }
    DispatchRef& operator=(DispatchRef&& other) noexcept {
        if (this != &other) {
            reset();
            value = other.value;
            other.value = nullptr;
        }
        return *this;
    }
    ~DispatchRef() { reset(); }
    void reset(IDispatch* replacement = nullptr) {
        if (value != nullptr) value->Release();
        value = replacement;
    }
    IDispatch* get() const { return value; }
    IDispatch** put() { reset(); return &value; }
};

struct VariantOwner {
    VARIANT value;
    VariantOwner() { VariantInit(&value); }
    VariantOwner(const VariantOwner&) = delete;
    VariantOwner& operator=(const VariantOwner&) = delete;
    ~VariantOwner() { VariantClear(&value); }
};

struct ExceptionOwner {
    EXCEPINFO value{};
    ~ExceptionOwner() {
        SysFreeString(value.bstrSource);
        SysFreeString(value.bstrDescription);
        SysFreeString(value.bstrHelpFile);
    }
    HRESULT fill_and_copy_scode(int32_t* out_scode) {
        HRESULT deferred = S_OK;
        if (value.pfnDeferredFillIn != nullptr) deferred = value.pfnDeferredFillIn(&value);
        if (out_scode != nullptr) *out_scode = value.scode;
        return deferred;
    }
};

struct ProcessHandle {
    HANDLE value = nullptr;
    ~ProcessHandle() { if (value != nullptr) CloseHandle(value); }
};

LCID lcid_for_mode(uint32_t mode) {
    return mode == EXCEL_RAW_COCREATE_LOCAL_LCID_0400 ? 0x0400 : 0;
}

DWORD clsctx_for_mode(uint32_t mode) {
    return mode == EXCEL_RAW_COCREATE_LOCAL_LCID_0400 ? CLSCTX_LOCAL_SERVER : CLSCTX_SERVER;
}

HRESULT activate(uint32_t mode, REFCLSID clsid, DispatchRef* out) {
    if (out == nullptr) return E_POINTER;
    if (mode == EXCEL_RAW_COCREATEEX_SERVER_LCID_0000) {
        MULTI_QI query{};
        query.pIID = &IID_IDispatch;
        HRESULT hr = CoCreateInstanceEx(clsid, nullptr, CLSCTX_SERVER, nullptr, 1, &query);
        if (FAILED(hr)) return hr;
        if (FAILED(query.hr)) return query.hr;
        out->reset(static_cast<IDispatch*>(query.pItf));
        return out->get() == nullptr ? E_NOINTERFACE : S_OK;
    }
    return CoCreateInstance(clsid, nullptr, clsctx_for_mode(mode), IID_IDispatch,
                            reinterpret_cast<void**>(out->put()));
}

HRESULT dispid(IDispatch* dispatch, const wchar_t* name, LCID lcid, DISPID* out) {
    if (dispatch == nullptr || out == nullptr) return E_POINTER;
    LPOLESTR names[] = { const_cast<LPOLESTR>(name) };
    return dispatch->GetIDsOfNames(IID_NULL, names, 1, lcid, out);
}

HRESULT invoke(IDispatch* dispatch, const wchar_t* name, LCID lcid, WORD flags,
               DISPPARAMS* params, VARIANT* result, ExceptionOwner* exception,
               UINT* arg_error, DISPID* resolved) {
    DISPID member = DISPID_UNKNOWN;
    HRESULT hr = dispid(dispatch, name, lcid, &member);
    if (resolved != nullptr) *resolved = member;
    if (FAILED(hr)) return hr;
    return dispatch->Invoke(member, IID_NULL, lcid, flags, params, result,
                            exception == nullptr ? nullptr : &exception->value, arg_error);
}

HRESULT get_dispatch(IDispatch* parent, const wchar_t* name, LCID lcid, DispatchRef* out,
                     uint16_t* vartype, ExceptionOwner* exception, UINT* arg_error,
                     DISPID* resolved) {
    VariantOwner result;
    DISPPARAMS empty{};
    HRESULT hr = invoke(parent, name, lcid, DISPATCH_PROPERTYGET, &empty, &result.value,
                        exception, arg_error, resolved);
    if (vartype != nullptr) *vartype = V_VT(&result.value);
    if (FAILED(hr)) return hr;
    if (V_VT(&result.value) != VT_DISPATCH || V_DISPATCH(&result.value) == nullptr) return kNoInterfaceResult;
    // The returned VARIANT owns this interface reference. Keep an independent
    // reference before the VariantOwner clears the result at function exit.
    IDispatch* returned = V_DISPATCH(&result.value);
    returned->AddRef();
    out->reset(returned);
    return S_OK;
}

HRESULT workbooks_count(IDispatch* workbooks, LCID lcid) {
    VariantOwner result;
    ExceptionOwner exception;
    DISPPARAMS empty{};
    UINT arg_error = UINT_MAX;
    return invoke(workbooks, L"Count", lcid, DISPATCH_PROPERTYGET, &empty, &result.value,
                  &exception, &arg_error, nullptr);
}

HRESULT close_workbook(IDispatch* workbook, LCID lcid) {
    VariantOwner save_changes;
    V_VT(&save_changes.value) = VT_BOOL;
    V_BOOL(&save_changes.value) = VARIANT_FALSE;
    DISPPARAMS params{};
    params.rgvarg = &save_changes.value;
    params.cArgs = 1;
    VariantOwner result;
    ExceptionOwner exception;
    UINT arg_error = UINT_MAX;
    return invoke(workbook, L"Close", lcid, DISPATCH_METHOD, &params, &result.value,
                  &exception, &arg_error, nullptr);
}

HRESULT lifetime_sequence(IDispatch* app, LCID lcid, int sequence) {
    VariantOwner result;
    DISPPARAMS empty{};
    ExceptionOwner exception;
    UINT arg_error = UINT_MAX;
    HRESULT hr = invoke(app, L"Workbooks", lcid, DISPATCH_PROPERTYGET, &empty, &result.value,
                        &exception, &arg_error, nullptr);
    if (FAILED(hr)) return hr;
    if (V_VT(&result.value) != VT_DISPATCH || V_DISPATCH(&result.value) == nullptr) return kNoInterfaceResult;
    if (sequence == 1) {
        DispatchRef clone(V_DISPATCH(&result.value));
        clone.get()->AddRef();
        VariantClear(&result.value);
        VariantInit(&result.value);
        return workbooks_count(clone.get(), lcid);
    }
    if (sequence == 2) {
        hr = workbooks_count(V_DISPATCH(&result.value), lcid);
        return hr;
    }
    DispatchRef queried;
    hr = V_DISPATCH(&result.value)->QueryInterface(IID_IDispatch,
                                                    reinterpret_cast<void**>(queried.put()));
    if (FAILED(hr)) return hr;
    VariantClear(&result.value);
    VariantInit(&result.value);
    return workbooks_count(queried.get(), lcid);
}

void fill_layout(ExcelAbiLayout* out) {
    std::memset(out, 0, sizeof(*out));
    out->pointer_width = sizeof(void*);
    out->variant_size = sizeof(VARIANT);
    out->variant_align = alignof(VARIANT);
    out->dispparams_size = sizeof(DISPPARAMS);
    out->dispparams_align = alignof(DISPPARAMS);
    out->excepinfo_size = sizeof(EXCEPINFO);
    out->excepinfo_align = alignof(EXCEPINFO);
    out->guid_size = sizeof(GUID);
    out->guid_align = alignof(GUID);
    out->variant_vt_offset = offsetof(VARIANT, vt);
    out->variant_data_offset = offsetof(VARIANT, lVal);
    out->variant_error_offset = offsetof(VARIANT, scode);
    out->variant_i4_offset = offsetof(VARIANT, lVal);
    out->variant_bstr_offset = offsetof(VARIANT, bstrVal);
    out->variant_dispatch_offset = offsetof(VARIANT, pdispVal);
    out->dispparams_rgvarg_offset = offsetof(DISPPARAMS, rgvarg);
    out->dispparams_named_offset = offsetof(DISPPARAMS, rgdispidNamedArgs);
    out->dispparams_args_offset = offsetof(DISPPARAMS, cArgs);
    out->dispparams_named_count_offset = offsetof(DISPPARAMS, cNamedArgs);
    out->excepinfo_wcode_offset = offsetof(EXCEPINFO, wCode);
    out->excepinfo_source_offset = offsetof(EXCEPINFO, bstrSource);
    out->excepinfo_description_offset = offsetof(EXCEPINFO, bstrDescription);
    out->excepinfo_help_file_offset = offsetof(EXCEPINFO, bstrHelpFile);
    out->excepinfo_help_context_offset = offsetof(EXCEPINFO, dwHelpContext);
    out->excepinfo_reserved_offset = offsetof(EXCEPINFO, pvReserved);
    out->excepinfo_deferred_offset = offsetof(EXCEPINFO, pfnDeferredFillIn);
    out->excepinfo_scode_offset = offsetof(EXCEPINFO, scode);
    out->result_size = sizeof(ExcelRawResult);
    out->result_align = alignof(ExcelRawResult);
}

} // namespace

extern "C" EXCEL_COM_NATIVE_API int32_t excel_raw_abi_layout(ExcelAbiLayout* out_layout) {
    if (out_layout == nullptr) return E_POINTER;
    fill_layout(out_layout);
    return S_OK;
}

extern "C" EXCEL_COM_NATIVE_API int32_t excel_raw_run(
    uint32_t mode, const wchar_t* fixture_path, ExcelRawResult* out_result) {
    if (out_result == nullptr) return E_POINTER;
    std::memset(out_result, 0, sizeof(*out_result));
    out_result->schema_version = 1;
    out_result->activation_hresult = kNotTested;
    out_result->version_hresult = kNotTested;
    out_result->workbooks_hresult = kNotTested;
    out_result->count_hresult = kNotTested;
    out_result->add_hresult = kNotTested;
    out_result->open_hresult = kNotTested;
    out_result->quit_hresult = kNotTested;
    out_result->inner_scode = 0;
    out_result->deferred_fill_in_hresult = S_OK;
    out_result->pu_arg_err_raw = UINT_MAX;
    out_result->lifetime_clone_then_clear = kNotTested;
    out_result->lifetime_retain_then_clear = kNotTested;
    out_result->lifetime_query_interface_then_clear = kNotTested;

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return hr;
    ProcessHandle process;
    DispatchRef app;
    const LCID lcid = lcid_for_mode(mode);
    do {
        CLSID clsid{};
        hr = CLSIDFromProgID(L"Excel.Application", &clsid);
        if (FAILED(hr)) { out_result->activation_hresult = hr; break; }
        hr = activate(mode, clsid, &app);
        out_result->activation_hresult = hr;
        if (FAILED(hr)) break;

        // Version is a canonical, non-mutating property get.
        VariantOwner version;
        ExceptionOwner version_exception;
        DISPPARAMS empty{};
        UINT version_arg_error = UINT_MAX;
        out_result->version_hresult = invoke(app.get(), L"Version", lcid, DISPATCH_PROPERTYGET,
                                             &empty, &version.value, &version_exception,
                                             &version_arg_error, nullptr);
        ExceptionOwner workbooks_exception;
        UINT workbooks_arg_error = UINT_MAX;
        DispatchRef workbooks;
        out_result->workbooks_hresult = get_dispatch(app.get(), L"Workbooks", lcid, &workbooks,
                                                      &out_result->workbooks_vt,
                                                      &workbooks_exception, &workbooks_arg_error, nullptr);
        if (FAILED(out_result->workbooks_hresult)) {
            out_result->inner_scode = workbooks_exception.value.scode;
            out_result->deferred_fill_in_hresult = workbooks_exception.fill_and_copy_scode(&out_result->inner_scode);
            break;
        }
        VariantOwner add_result;
        ExceptionOwner add_exception;
        UINT add_arg_error = UINT_MAX;
        out_result->add_hresult = invoke(workbooks.get(), L"Add", lcid, DISPATCH_METHOD, &empty,
                                         &add_result.value, &add_exception, &add_arg_error, nullptr);
        out_result->result_vt = V_VT(&add_result.value);
        out_result->pu_arg_err_raw = add_arg_error;
        out_result->inner_scode = add_exception.value.scode;
        out_result->deferred_fill_in_hresult = add_exception.fill_and_copy_scode(&out_result->inner_scode);
        if (SUCCEEDED(out_result->add_hresult) && V_VT(&add_result.value) == VT_DISPATCH && V_DISPATCH(&add_result.value) != nullptr) {
            DispatchRef workbook(V_DISPATCH(&add_result.value));
            workbook.get()->AddRef();
            out_result->workbook_created = SUCCEEDED(close_workbook(workbook.get(), lcid)) ? 1 : 0;
        }

        if (fixture_path != nullptr && fixture_path[0] != L'\0') {
            VariantOwner argument;
            V_VT(&argument.value) = VT_BSTR;
            V_BSTR(&argument.value) = SysAllocString(fixture_path);
            DISPPARAMS open_params{};
            open_params.rgvarg = &argument.value;
            open_params.cArgs = 1;
            VariantOwner open_result;
            ExceptionOwner open_exception;
            UINT open_arg_error = UINT_MAX;
            out_result->open_hresult = invoke(workbooks.get(), L"Open", lcid, DISPATCH_METHOD,
                                              &open_params, &open_result.value, &open_exception,
                                              &open_arg_error, nullptr);
            if (FAILED(out_result->open_hresult)) {
                out_result->inner_scode = open_exception.value.scode;
                out_result->deferred_fill_in_hresult = open_exception.fill_and_copy_scode(&out_result->inner_scode);
            } else if (V_VT(&open_result.value) == VT_DISPATCH && V_DISPATCH(&open_result.value) != nullptr) {
                DispatchRef workbook(V_DISPATCH(&open_result.value));
                workbook.get()->AddRef();
                out_result->workbook_opened = SUCCEEDED(close_workbook(workbook.get(), lcid)) ? 1 : 0;
            }
        }
        // Hwnd is used only after the primary operations to obtain a process
        // handle for bounded cleanup; its raw value is never copied to output.
        VariantOwner hwnd;
        ExceptionOwner hwnd_exception;
        UINT hwnd_arg_error = UINT_MAX;
        HRESULT hwnd_hr = invoke(app.get(), L"Hwnd", lcid, DISPATCH_PROPERTYGET, &empty,
                                 &hwnd.value, &hwnd_exception, &hwnd_arg_error, nullptr);
        if (SUCCEEDED(hwnd_hr) && V_VT(&hwnd.value) == VT_I4) {
            DWORD pid = 0;
            GetWindowThreadProcessId(reinterpret_cast<HWND>(static_cast<intptr_t>(V_I4(&hwnd.value))), &pid);
            if (pid != 0) process.value = OpenProcess(SYNCHRONIZE, FALSE, pid);
        }
        // Capability probes are deliberately after the canonical Add/Open
        // operations. They are recorded, but cannot become an accidental
        // precondition for the primary operation under comparison.
        UINT type_info_count = 0;
        out_result->type_info_hresult = workbooks.get()->GetTypeInfoCount(&type_info_count);
        if (SUCCEEDED(out_result->type_info_hresult)) out_result->type_info_count = type_info_count;
        IUnknown* as_unknown = nullptr;
        out_result->workbooks_query_iunknown_hresult = workbooks.get()->QueryInterface(
            IID_IUnknown, reinterpret_cast<void**>(&as_unknown));
        if (as_unknown != nullptr) as_unknown->Release();
        IDispatch* as_dispatch = nullptr;
        out_result->workbooks_query_idispatch_hresult = workbooks.get()->QueryInterface(
            IID_IDispatch, reinterpret_cast<void**>(&as_dispatch));
        if (as_dispatch != nullptr) as_dispatch->Release();
        out_result->count_hresult = workbooks_count(workbooks.get(), lcid);
        out_result->lifetime_clone_then_clear = lifetime_sequence(app.get(), lcid, 1);
        out_result->lifetime_retain_then_clear = lifetime_sequence(app.get(), lcid, 2);
        out_result->lifetime_query_interface_then_clear = lifetime_sequence(app.get(), lcid, 3);
        VariantOwner quit_result;
        ExceptionOwner quit_exception;
        UINT quit_arg_error = UINT_MAX;
        out_result->quit_hresult = invoke(app.get(), L"Quit", lcid, DISPATCH_METHOD, &empty,
                                          &quit_result.value, &quit_exception, &quit_arg_error, nullptr);
    } while (false);
    app.reset();
    if (process.value != nullptr) out_result->process_exited =
        WaitForSingleObject(process.value, 15000) == WAIT_OBJECT_0 ? 1 : 0;
    CoUninitialize();
    return S_OK;
}
