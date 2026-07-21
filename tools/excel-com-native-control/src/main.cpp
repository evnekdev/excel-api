#include "excel_com_native_control.h"

#include <cstdio>
#include <cstring>
#include <cwchar>

namespace {

uint32_t parse_mode(const wchar_t* value) {
    if (std::wcscmp(value, L"native-cocreate-local-lcid-0400") == 0) return EXCEL_RAW_COCREATE_LOCAL_LCID_0400;
    if (std::wcscmp(value, L"native-cocreate-server-lcid-0000") == 0) return EXCEL_RAW_COCREATE_SERVER_LCID_0000;
    if (std::wcscmp(value, L"native-cocreateex-server-lcid-0000") == 0) return EXCEL_RAW_COCREATEEX_SERVER_LCID_0000;
    return 0;
}

void print_result(uint32_t mode, const ExcelRawResult& r) {
    std::printf("{\"schema_version\":1,\"path\":\"native-cpp-direct\",\"mode\":%u,"
                "\"activation_hresult\":%d,\"version_hresult\":%d,\"workbooks_hresult\":%d,"
                "\"count_hresult\":%d,\"add_hresult\":%d,\"open_hresult\":%d,\"quit_hresult\":%d,"
                "\"inner_scode\":%d,\"deferred_fill_in_hresult\":%d,\"result_vt\":%u,"
                "\"workbooks_vt\":%u,\"pu_arg_err_raw\":%u,\"workbook_created\":%d,"
                "\"workbook_opened\":%d,\"process_exited\":%d,\"lifetime_clone_then_clear\":%d,"
                "\"lifetime_retain_then_clear\":%d,\"lifetime_query_interface_then_clear\":%d,"
                "\"type_info_count\":%u,\"type_info_hresult\":%d,\"workbooks_query_iunknown_hresult\":%d,"
                "\"workbooks_query_idispatch_hresult\":%d,\"raw_paths_recorded\":false,"
                "\"raw_hwnd_recorded\":false,\"raw_pointer_values_recorded\":false}\n",
                mode, r.activation_hresult, r.version_hresult, r.workbooks_hresult, r.count_hresult,
                r.add_hresult, r.open_hresult, r.quit_hresult, r.inner_scode,
                r.deferred_fill_in_hresult, r.result_vt, r.workbooks_vt, r.pu_arg_err_raw,
                r.workbook_created, r.workbook_opened, r.process_exited, r.lifetime_clone_then_clear,
                r.lifetime_retain_then_clear, r.lifetime_query_interface_then_clear, r.type_info_count,
                r.type_info_hresult, r.workbooks_query_iunknown_hresult,
                r.workbooks_query_idispatch_hresult);
}

void print_layout(const ExcelAbiLayout& l) {
    std::printf("{\"schema_version\":1,\"path\":\"native-cpp-sdk-layout\","
                "\"pointer_width\":%u,\"variant\":{\"size\":%u,\"align\":%u,"
                "\"vt_offset\":%u,\"data_offset\":%u,\"error_offset\":%u,\"i4_offset\":%u,"
                "\"bstr_offset\":%u,\"dispatch_offset\":%u},\"dispparams\":{\"size\":%u,"
                "\"align\":%u,\"rgvarg_offset\":%u,\"named_offset\":%u,\"args_offset\":%u,"
                "\"named_count_offset\":%u},\"excepinfo\":{\"size\":%u,\"align\":%u,"
                "\"wcode_offset\":%u,\"source_offset\":%u,\"description_offset\":%u,"
                "\"help_file_offset\":%u,\"help_context_offset\":%u,\"reserved_offset\":%u,"
                "\"deferred_offset\":%u,\"scode_offset\":%u},\"result_size\":%u,\"result_align\":%u,"
                "\"raw_pointer_values_recorded\":false}\n",
                l.pointer_width, l.variant_size, l.variant_align, l.variant_vt_offset, l.variant_data_offset,
                l.variant_error_offset, l.variant_i4_offset, l.variant_bstr_offset, l.variant_dispatch_offset,
                l.dispparams_size, l.dispparams_align, l.dispparams_rgvarg_offset, l.dispparams_named_offset,
                l.dispparams_args_offset, l.dispparams_named_count_offset, l.excepinfo_size, l.excepinfo_align,
                l.excepinfo_wcode_offset, l.excepinfo_source_offset, l.excepinfo_description_offset,
                l.excepinfo_help_file_offset, l.excepinfo_help_context_offset, l.excepinfo_reserved_offset,
                l.excepinfo_deferred_offset, l.excepinfo_scode_offset, l.result_size, l.result_align);
}

} // namespace

int wmain(int argc, wchar_t** argv) {
    uint32_t mode = EXCEL_RAW_COCREATE_LOCAL_LCID_0400;
    const wchar_t* fixture = nullptr;
    bool layout_only = false;
    for (int i = 1; i < argc; ++i) {
        if (std::wcscmp(argv[i], L"--mode") == 0 && i + 1 < argc) mode = parse_mode(argv[++i]);
        else if (std::wcscmp(argv[i], L"--fixture") == 0 && i + 1 < argc) fixture = argv[++i];
        else if (std::wcscmp(argv[i], L"--layout") == 0) layout_only = true;
    }
    if (layout_only) {
        ExcelAbiLayout layout{};
        if (excel_raw_abi_layout(&layout) != 0) return 4;
        print_layout(layout);
        return 0;
    }
    if (mode == 0) return 2;
    ExcelRawResult result{};
    const int32_t status = excel_raw_run(mode, fixture, &result);
    if (status != 0) return 3;
    print_result(mode, result);
    return 0;
}
