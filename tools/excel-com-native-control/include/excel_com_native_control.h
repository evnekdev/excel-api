#pragma once

#include <stdint.h>

#if defined(_WIN32)
#if defined(EXCEL_COM_NATIVE_BUILD_DLL)
#define EXCEL_COM_NATIVE_API __declspec(dllexport)
#else
#define EXCEL_COM_NATIVE_API __declspec(dllimport)
#endif
#else
#define EXCEL_COM_NATIVE_API
#endif

enum ExcelRawMode : uint32_t {
    EXCEL_RAW_COCREATE_LOCAL_LCID_0400 = 1,
    EXCEL_RAW_COCREATE_SERVER_LCID_0000 = 2,
    EXCEL_RAW_COCREATEEX_SERVER_LCID_0000 = 3,
};

// Fixed-width, copied diagnostics only. This intentionally carries neither a
// COM interface pointer nor a BSTR ownership obligation across the ABI.
struct ExcelRawResult {
    uint32_t schema_version;
    int32_t activation_hresult;
    int32_t version_hresult;
    int32_t workbooks_hresult;
    int32_t count_hresult;
    int32_t add_hresult;
    int32_t open_hresult;
    int32_t quit_hresult;
    int32_t inner_scode;
    int32_t deferred_fill_in_hresult;
    uint16_t result_vt;
    uint16_t workbooks_vt;
    uint32_t pu_arg_err_raw;
    int32_t workbook_created;
    int32_t workbook_opened;
    int32_t process_exited;
    int32_t lifetime_clone_then_clear;
    int32_t lifetime_retain_then_clear;
    int32_t lifetime_query_interface_then_clear;
    uint32_t type_info_count;
    int32_t type_info_hresult;
    int32_t workbooks_query_iunknown_hresult;
    int32_t workbooks_query_idispatch_hresult;
};

struct ExcelAbiLayout {
    uint32_t pointer_width;
    uint32_t variant_size;
    uint32_t variant_align;
    uint32_t dispparams_size;
    uint32_t dispparams_align;
    uint32_t excepinfo_size;
    uint32_t excepinfo_align;
    uint32_t guid_size;
    uint32_t guid_align;
    uint32_t variant_vt_offset;
    uint32_t variant_data_offset;
    uint32_t variant_error_offset;
    uint32_t variant_i4_offset;
    uint32_t variant_bstr_offset;
    uint32_t variant_dispatch_offset;
    uint32_t dispparams_rgvarg_offset;
    uint32_t dispparams_named_offset;
    uint32_t dispparams_args_offset;
    uint32_t dispparams_named_count_offset;
    uint32_t excepinfo_wcode_offset;
    uint32_t excepinfo_source_offset;
    uint32_t excepinfo_description_offset;
    uint32_t excepinfo_help_file_offset;
    uint32_t excepinfo_help_context_offset;
    uint32_t excepinfo_reserved_offset;
    uint32_t excepinfo_deferred_offset;
    uint32_t excepinfo_scode_offset;
    uint32_t result_size;
    uint32_t result_align;
};

extern "C" EXCEL_COM_NATIVE_API int32_t excel_raw_run(
    uint32_t mode,
    const wchar_t* fixture_path,
    ExcelRawResult* out_result);

extern "C" EXCEL_COM_NATIVE_API int32_t excel_raw_abi_layout(ExcelAbiLayout* out_layout);
