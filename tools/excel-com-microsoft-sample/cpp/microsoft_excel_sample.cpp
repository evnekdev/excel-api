// Independently written, Microsoft-sample-faithful C++ control.
//
// Based on Microsoft's "How to automate Excel from C++ without using MFC or
// #import" sample. This file deliberately preserves its raw IDispatch call
// classifications, locale asymmetry, CoInitialize(NULL), CLSCTX_LOCAL_SERVER,
// and AutoWrap DISPPARAMS construction. It replaces only UI/error exits with
// structured JSON so the experiment is repeatable and non-interactive.

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <tlhelp32.h>

#include <chrono>
#include <cstdarg>
#include <cstring>
#include <iostream>
#include <string>
#include <thread>
#include <vector>

namespace {

constexpr int kDispatchMethod = DISPATCH_METHOD;
constexpr int kDispatchPropertyGet = DISPATCH_PROPERTYGET;
constexpr int kDispatchPropertyPut = DISPATCH_PROPERTYPUT;

struct Trace {
    unsigned sequence = 0;
    std::string member;
    std::string role;
    LONG get_ids_hr = E_FAIL;
    LONG invoke_hr = E_FAIL;
    DISPID dispid = DISPID_UNKNOWN;
    int flags = 0;
    unsigned args = 0;
    unsigned named_args = 0;
    std::vector<unsigned> argument_vartypes;
    bool result_requested = false;
    unsigned result_vartype = VT_EMPTY;
    bool observational = false;
    LONG excepinfo_scode = 0;
    bool argerr_present = false;
    unsigned argerr = 0;
};

std::vector<Trace> g_trace;
bool g_observational = false;

std::string JsonEscape(const std::string& value) {
    std::string escaped;
    for (const char character : value) {
        if (character == '"' || character == '\\') {
            escaped.push_back('\\');
        }
        escaped.push_back(character);
    }
    return escaped;
}

std::string NarrowAscii(const wchar_t* value) {
    std::string result;
    if (value == nullptr) {
        return result;
    }
    while (*value != L'\0') {
        result.push_back(static_cast<char>(*value));
        ++value;
    }
    return result;
}

void ClearExcepInfo(EXCEPINFO* info) {
    if (info == nullptr) {
        return;
    }
    if (info->pfnDeferredFillIn != nullptr) {
        (void)info->pfnDeferredFillIn(info);
    }
    if (info->bstrSource != nullptr) {
        SysFreeString(info->bstrSource);
    }
    if (info->bstrDescription != nullptr) {
        SysFreeString(info->bstrDescription);
    }
    if (info->bstrHelpFile != nullptr) {
        SysFreeString(info->bstrHelpFile);
    }
}

bool IsExcelName(const WCHAR* name) {
    return _wcsicmp(name, L"EXCEL.EXE") == 0;
}

unsigned ExcelProcessCount() {
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snapshot == INVALID_HANDLE_VALUE) {
        return 0;
    }
    PROCESSENTRY32W entry{};
    entry.dwSize = sizeof(entry);
    unsigned count = 0;
    if (Process32FirstW(snapshot, &entry)) {
        do {
            if (IsExcelName(entry.szExeFile)) {
                ++count;
            }
        } while (Process32NextW(snapshot, &entry));
    }
    CloseHandle(snapshot);
    return count;
}

bool WaitForZeroExcelProcesses() {
    for (unsigned attempt = 0; attempt != 200; ++attempt) {
        if (ExcelProcessCount() == 0) {
            return true;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
    }
    return ExcelProcessCount() == 0;
}

// This is a faithful behavioural translation of the article's AutoWrap:
// DISPPARAMS starts zeroed; property puts alone receive DISPID_PROPERTYPUT;
// callers provide multiple arguments in COM reverse order; lookup/invoke use
// the article's distinct LCIDs. The variadic extraction is intentionally kept.
HRESULT AutoWrap(
    int autoType,
    VARIANT* pvResult,
    IDispatch* pDisp,
    LPOLESTR ptName,
    const char* role,
    int cArgs,
    ...) {
    Trace trace{};
    trace.sequence = static_cast<unsigned>(g_trace.size() + 1);
    trace.member = NarrowAscii(ptName);
    trace.role = role;
    trace.flags = autoType;
    trace.args = static_cast<unsigned>(cArgs);
    trace.result_requested = pvResult != nullptr;
    trace.observational = g_observational;

    if (pDisp == nullptr) {
        trace.get_ids_hr = E_POINTER;
        trace.invoke_hr = E_POINTER;
        g_trace.push_back(trace);
        return E_POINTER;
    }

    DISPPARAMS dp = {nullptr, nullptr, 0, 0};
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID = DISPID_UNKNOWN;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    trace.get_ids_hr = hr;
    trace.dispid = dispID;
    if (FAILED(hr)) {
        trace.invoke_hr = hr;
        g_trace.push_back(trace);
        return hr;
    }

    va_list marker;
    va_start(marker, cArgs);
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    for (int index = 0; index < cArgs; ++index) {
        pArgs[index] = va_arg(marker, VARIANT);
        trace.argument_vartypes.push_back(pArgs[index].vt);
    }
    va_end(marker);

    dp.cArgs = cArgs;
    dp.rgvarg = cArgs == 0 ? nullptr : pArgs;
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }
    trace.named_args = dp.cNamedArgs;

    EXCEPINFO exception{};
    UINT argerr = UINT_MAX;
    hr = pDisp->Invoke(
        dispID,
        IID_NULL,
        LOCALE_SYSTEM_DEFAULT,
        static_cast<WORD>(autoType),
        &dp,
        pvResult,
        g_observational ? &exception : nullptr,
        g_observational ? &argerr : nullptr);
    trace.invoke_hr = hr;
    if (pvResult != nullptr) {
        trace.result_vartype = pvResult->vt;
    }
    if (g_observational) {
        trace.excepinfo_scode = exception.scode;
        trace.argerr_present = argerr != UINT_MAX;
        trace.argerr = argerr;
        ClearExcepInfo(&exception);
    }
    delete[] pArgs;
    g_trace.push_back(trace);
    return hr;
}

bool IsExpectedNumber(const VARIANT& value, long expected) {
    if (value.vt == VT_I4) {
        return value.lVal == expected;
    }
    if (value.vt == VT_R8) {
        return value.dblVal == static_cast<double>(expected);
    }
    return false;
}

HRESULT ReadCell(IDispatch* sheet, const wchar_t* address, long expected, bool* matches) {
    VARIANT argument{};
    VariantInit(&argument);
    argument.vt = VT_BSTR;
    argument.bstrVal = SysAllocString(address);
    if (argument.bstrVal == nullptr) {
        return E_OUTOFMEMORY;
    }
    VARIANT range_result{};
    VariantInit(&range_result);
    HRESULT hr = AutoWrap(
        kDispatchPropertyGet,
        &range_result,
        sheet,
        const_cast<LPOLESTR>(L"Range"),
        "worksheet",
        1,
        argument);
    VariantClear(&argument);
    if (FAILED(hr) || range_result.vt != VT_DISPATCH || range_result.pdispVal == nullptr) {
        VariantClear(&range_result);
        return FAILED(hr) ? hr : DISP_E_TYPEMISMATCH;
    }
    IDispatch* range = range_result.pdispVal;
    range_result.vt = VT_EMPTY;
    range_result.pdispVal = nullptr;
    VARIANT value{};
    VariantInit(&value);
    hr = AutoWrap(kDispatchPropertyGet, &value, range, const_cast<LPOLESTR>(L"Value"), "range", 0);
    if (SUCCEEDED(hr)) {
        *matches = IsExpectedNumber(value, expected);
        if (!*matches) {
            hr = E_FAIL;
        }
    }
    VariantClear(&value);
    range->Release();
    return hr;
}

void EmitJson(
    const char* classification,
    const char* failure_stage,
    HRESULT failure_hr,
    bool readback_a1,
    bool readback_b3,
    bool readback_o15,
    bool process_exited,
    unsigned preexisting_count) {
    std::cout << "{\"control\":\"official-cpp\",\"classification\":\"" << classification
              << "\",\"failure_stage\":\"" << failure_stage
              << "\",\"failure_hresult\":" << static_cast<LONG>(failure_hr)
              << ",\"preexisting_excel_process_count\":" << preexisting_count
              << ",\"readback\":{\"A1\":" << (readback_a1 ? "true" : "false")
              << ",\"B3\":" << (readback_b3 ? "true" : "false")
              << ",\"O15\":" << (readback_o15 ? "true" : "false") << "}"
              << ",\"cleanup\":{\"owned_process_exit_verified\":"
              << (process_exited ? "true" : "false")
              << ",\"forced_termination\":false}"
              << ",\"trace\":[";
    for (size_t index = 0; index < g_trace.size(); ++index) {
        const Trace& trace = g_trace[index];
        if (index != 0) {
            std::cout << ',';
        }
        std::cout << "{\"sequence\":" << trace.sequence
                  << ",\"member\":\"" << JsonEscape(trace.member)
                  << "\",\"object_role\":\"" << JsonEscape(trace.role)
                  << "\",\"get_ids_of_names_lcid\":" << LOCALE_USER_DEFAULT
                  << ",\"get_ids_of_names_hresult\":" << trace.get_ids_hr
                  << ",\"dispid\":" << trace.dispid
                  << ",\"invoke_lcid\":" << LOCALE_SYSTEM_DEFAULT
                  << ",\"invoke_flags\":" << trace.flags
                  << ",\"argument_count\":" << trace.args
                  << ",\"named_argument_count\":" << trace.named_args
                  << ",\"argument_vartypes\":[";
        for (size_t argument = 0; argument < trace.argument_vartypes.size(); ++argument) {
            if (argument != 0) {
                std::cout << ',';
            }
            std::cout << trace.argument_vartypes[argument];
        }
        std::cout << "]"
                  << ",\"result_hresult\":" << trace.invoke_hr
                  << ",\"result_requested\":" << (trace.result_requested ? "true" : "false")
                  << ",\"result_vartype\":" << trace.result_vartype
                  << ",\"excepinfo_mode\":\""
                  << (trace.observational ? "observational" : "official-null") << "\""
                  << ",\"excepinfo_scode\":" << trace.excepinfo_scode
                  << ",\"pu_arg_err\":";
        if (trace.argerr_present) {
            std::cout << trace.argerr;
        } else {
            std::cout << "null";
        }
        std::cout << '}';
    }
    std::cout << "]}\n";
}

}  // namespace

int wmain(int argc, wchar_t** argv) {
    g_observational = argc > 1 && wcscmp(argv[1], L"--observational") == 0;
    const unsigned preexisting = ExcelProcessCount();
    if (preexisting != 0) {
        EmitJson("activation failure", "preexisting-excel", HRESULT_FROM_WIN32(ERROR_BUSY), false, false, false, false, preexisting);
        return 0;
    }

    HRESULT failure = S_OK;
    const char* failure_stage = "";
    bool com_initialized = false;
    IDispatch* app = nullptr;
    IDispatch* books = nullptr;
    IDispatch* book = nullptr;
    IDispatch* sheet = nullptr;
    IDispatch* range = nullptr;
    VARIANT array{};
    VariantInit(&array);
    bool a1 = false;
    bool b3 = false;
    bool o15 = false;
    bool quit_requested = false;

    failure = CoInitialize(nullptr);
    if (FAILED(failure)) {
        failure_stage = "CoInitialize";
        goto cleanup;
    }
    com_initialized = true;
    {
        CLSID clsid{};
        failure = CLSIDFromProgID(L"Excel.Application", &clsid);
        if (FAILED(failure)) {
            failure_stage = "CLSIDFromProgID";
            goto cleanup;
        }
        failure = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, reinterpret_cast<void**>(&app));
        if (FAILED(failure) || app == nullptr) {
            failure_stage = "CoCreateInstance";
            if (SUCCEEDED(failure)) {
                failure = E_POINTER;
            }
            goto cleanup;
        }
    }
    {
        VARIANT visible{};
        VariantInit(&visible);
        visible.vt = VT_I4;
        visible.lVal = 1;
        failure = AutoWrap(kDispatchPropertyPut, nullptr, app, const_cast<LPOLESTR>(L"Visible"), "application", 1, visible);
        if (FAILED(failure)) {
            failure_stage = "Visible";
            goto cleanup;
        }
    }
    {
        VARIANT result{};
        VariantInit(&result);
        failure = AutoWrap(kDispatchPropertyGet, &result, app, const_cast<LPOLESTR>(L"Workbooks"), "application", 0);
        if (FAILED(failure) || result.vt != VT_DISPATCH || result.pdispVal == nullptr) {
            failure_stage = "Workbooks";
            if (SUCCEEDED(failure)) failure = DISP_E_TYPEMISMATCH;
            goto cleanup;
        }
        books = result.pdispVal;
        result.vt = VT_EMPTY;
        result.pdispVal = nullptr;
    }
    {
        VARIANT result{};
        VariantInit(&result);
        failure = AutoWrap(kDispatchPropertyGet, &result, books, const_cast<LPOLESTR>(L"Add"), "workbooks", 0);
        if (FAILED(failure) || result.vt != VT_DISPATCH || result.pdispVal == nullptr) {
            failure_stage = "Add";
            if (SUCCEEDED(failure)) failure = DISP_E_TYPEMISMATCH;
            goto cleanup;
        }
        book = result.pdispVal;
        result.vt = VT_EMPTY;
        result.pdispVal = nullptr;
    }
    {
        SAFEARRAYBOUND bounds[2]{};
        bounds[0].lLbound = 1; bounds[0].cElements = 15;
        bounds[1].lLbound = 1; bounds[1].cElements = 15;
        array.vt = VT_ARRAY | VT_VARIANT;
        array.parray = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (array.parray == nullptr) {
            failure = E_OUTOFMEMORY;
            failure_stage = "SAFEARRAY";
            goto cleanup;
        }
        for (long row = 1; row <= 15; ++row) {
            for (long column = 1; column <= 15; ++column) {
                VARIANT value{};
                VariantInit(&value);
                value.vt = VT_I4;
                value.lVal = row * column;
                long indices[] = {row, column};
                failure = SafeArrayPutElement(array.parray, indices, &value);
                if (FAILED(failure)) {
                    failure_stage = "SAFEARRAY";
                    goto cleanup;
                }
            }
        }
    }
    {
        VARIANT result{};
        VariantInit(&result);
        failure = AutoWrap(kDispatchPropertyGet, &result, app, const_cast<LPOLESTR>(L"ActiveSheet"), "application", 0);
        if (FAILED(failure) || result.vt != VT_DISPATCH || result.pdispVal == nullptr) {
            failure_stage = "ActiveSheet";
            if (SUCCEEDED(failure)) failure = DISP_E_TYPEMISMATCH;
            goto cleanup;
        }
        sheet = result.pdispVal;
        result.vt = VT_EMPTY;
        result.pdispVal = nullptr;
    }
    {
        VARIANT address{};
        VariantInit(&address);
        address.vt = VT_BSTR;
        address.bstrVal = SysAllocString(L"A1:O15");
        if (address.bstrVal == nullptr) {
            failure = E_OUTOFMEMORY;
            failure_stage = "Range";
            goto cleanup;
        }
        VARIANT result{};
        VariantInit(&result);
        failure = AutoWrap(kDispatchPropertyGet, &result, sheet, const_cast<LPOLESTR>(L"Range"), "worksheet", 1, address);
        VariantClear(&address);
        if (FAILED(failure) || result.vt != VT_DISPATCH || result.pdispVal == nullptr) {
            failure_stage = "Range";
            if (SUCCEEDED(failure)) failure = DISP_E_TYPEMISMATCH;
            goto cleanup;
        }
        range = result.pdispVal;
        result.vt = VT_EMPTY;
        result.pdispVal = nullptr;
    }
    failure = AutoWrap(kDispatchPropertyPut, nullptr, range, const_cast<LPOLESTR>(L"Value"), "range", 1, array);
    if (FAILED(failure)) {
        failure_stage = "Value";
        goto cleanup;
    }
    failure = ReadCell(sheet, L"A1", 1, &a1);
    if (FAILED(failure)) { failure_stage = "read-back"; goto cleanup; }
    failure = ReadCell(sheet, L"B3", 6, &b3);
    if (FAILED(failure)) { failure_stage = "read-back"; goto cleanup; }
    failure = ReadCell(sheet, L"O15", 225, &o15);
    if (FAILED(failure)) { failure_stage = "read-back"; goto cleanup; }
    {
        VARIANT saved{};
        VariantInit(&saved);
        saved.vt = VT_I4;
        saved.lVal = 1;
        failure = AutoWrap(kDispatchPropertyPut, nullptr, book, const_cast<LPOLESTR>(L"Saved"), "workbook", 1, saved);
        if (FAILED(failure)) { failure_stage = "Saved"; goto cleanup; }
    }
    failure = AutoWrap(kDispatchMethod, nullptr, app, const_cast<LPOLESTR>(L"Quit"), "application", 0);
    quit_requested = SUCCEEDED(failure);
    if (FAILED(failure)) { failure_stage = "Quit"; }

cleanup:
    if (app != nullptr && !quit_requested) {
        (void)AutoWrap(kDispatchMethod, nullptr, app, const_cast<LPOLESTR>(L"Quit"), "application", 0);
    }
    if (range != nullptr) range->Release();
    if (sheet != nullptr) sheet->Release();
    if (book != nullptr) book->Release();
    if (books != nullptr) books->Release();
    if (app != nullptr) app->Release();
    VariantClear(&array);
    if (com_initialized) CoUninitialize();
    const bool exited = WaitForZeroExcelProcesses();
    const bool complete = SUCCEEDED(failure) && failure_stage[0] == '\0' && a1 && b3 && o15 && exited;
    EmitJson(complete ? "complete" : "failure", failure_stage, failure, a1, b3, o15, exited, preexisting);
    return 0;
}
