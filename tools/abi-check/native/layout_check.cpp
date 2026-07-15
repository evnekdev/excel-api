#ifdef _WIN32
/* Compiled as C++ so decltype can name the SDK's anonymous nested types. */
#include <windows.h>
#include <stddef.h>
#include <stdint.h>
#include <type_traits>
#include "XLCALL.H"

#ifdef __cplusplus
extern "C" {
#endif

#define SIZEOF_FN(name, type) size_t name(void) { return sizeof(type); }
#define ALIGNOF_FN(name, type) size_t name(void) { return __alignof(type); }
#define OFFSETOF_FN(name, type, field) size_t name(void) { return offsetof(type, field); }
#define CONST_I32_FN(name, value) int32_t name(void) { return (int32_t)(value); }
#define CONST_U32_FN(name, value) uint32_t name(void) { return (uint32_t)(value); }

SIZEOF_FN(excel_sdk_sizeof_bool, BOOL)
SIZEOF_FN(excel_sdk_sizeof_xchar, XCHAR)
SIZEOF_FN(excel_sdk_sizeof_rw, RW)
SIZEOF_FN(excel_sdk_sizeof_col, COL)
SIZEOF_FN(excel_sdk_sizeof_idsheet, IDSHEET)

SIZEOF_FN(excel_sdk_sizeof_xlref12, XLREF12)
ALIGNOF_FN(excel_sdk_alignof_xlref12, XLREF12)
OFFSETOF_FN(excel_sdk_offsetof_xlref12_rw_first, XLREF12, rwFirst)
OFFSETOF_FN(excel_sdk_offsetof_xlref12_rw_last, XLREF12, rwLast)
OFFSETOF_FN(excel_sdk_offsetof_xlref12_col_first, XLREF12, colFirst)
OFFSETOF_FN(excel_sdk_offsetof_xlref12_col_last, XLREF12, colLast)

SIZEOF_FN(excel_sdk_sizeof_xlmref12, XLMREF12)
ALIGNOF_FN(excel_sdk_alignof_xlmref12, XLMREF12)
OFFSETOF_FN(excel_sdk_offsetof_xlmref12_count, XLMREF12, count)
OFFSETOF_FN(excel_sdk_offsetof_xlmref12_reftbl, XLMREF12, reftbl)

SIZEOF_FN(excel_sdk_sizeof_fp12, FP12)
ALIGNOF_FN(excel_sdk_alignof_fp12, FP12)
OFFSETOF_FN(excel_sdk_offsetof_fp12_rows, FP12, rows)
OFFSETOF_FN(excel_sdk_offsetof_fp12_columns, FP12, columns)
OFFSETOF_FN(excel_sdk_offsetof_fp12_array, FP12, array)

size_t excel_sdk_sizeof_xloper12_value(void) { return sizeof(((XLOPER12 *)0)->val); }
size_t excel_sdk_alignof_xloper12_value(void) { return __alignof(decltype(((XLOPER12 *)0)->val)); }
size_t excel_sdk_sizeof_xloper12_num(void) { return sizeof(((XLOPER12 *)0)->val.num); }
size_t excel_sdk_sizeof_xloper12_str(void) { return sizeof(((XLOPER12 *)0)->val.str); }
size_t excel_sdk_sizeof_xloper12_xbool(void) { return sizeof(((XLOPER12 *)0)->val.xbool); }
size_t excel_sdk_sizeof_xloper12_err(void) { return sizeof(((XLOPER12 *)0)->val.err); }
size_t excel_sdk_sizeof_xloper12_w(void) { return sizeof(((XLOPER12 *)0)->val.w); }
size_t excel_sdk_sizeof_xloper12_sref(void) { return sizeof(((XLOPER12 *)0)->val.sref); }
size_t excel_sdk_alignof_xloper12_sref(void) { return __alignof(decltype(((XLOPER12 *)0)->val.sref)); }
size_t excel_sdk_offsetof_xloper12_sref_count(void) { return offsetof(XLOPER12, val.sref.count); }
size_t excel_sdk_offsetof_xloper12_sref_ref(void) { return offsetof(XLOPER12, val.sref.ref); }
size_t excel_sdk_sizeof_xloper12_mref(void) { return sizeof(((XLOPER12 *)0)->val.mref); }
size_t excel_sdk_alignof_xloper12_mref(void) { return __alignof(decltype(((XLOPER12 *)0)->val.mref)); }
size_t excel_sdk_offsetof_xloper12_mref_lpmref(void) { return offsetof(XLOPER12, val.mref.lpmref); }
size_t excel_sdk_offsetof_xloper12_mref_idsheet(void) { return offsetof(XLOPER12, val.mref.idSheet); }
size_t excel_sdk_sizeof_xloper12_array(void) { return sizeof(((XLOPER12 *)0)->val.array); }
size_t excel_sdk_alignof_xloper12_array(void) { return __alignof(decltype(((XLOPER12 *)0)->val.array)); }
size_t excel_sdk_offsetof_xloper12_array_lparray(void) { return offsetof(XLOPER12, val.array.lparray); }
size_t excel_sdk_offsetof_xloper12_array_rows(void) { return offsetof(XLOPER12, val.array.rows); }
size_t excel_sdk_offsetof_xloper12_array_columns(void) { return offsetof(XLOPER12, val.array.columns); }
size_t excel_sdk_sizeof_xloper12_flow_value(void) { return sizeof(((XLOPER12 *)0)->val.flow.valflow); }
size_t excel_sdk_alignof_xloper12_flow_value(void) { return __alignof(decltype(((XLOPER12 *)0)->val.flow.valflow)); }
size_t excel_sdk_sizeof_xloper12_flow_level(void) { return sizeof(((XLOPER12 *)0)->val.flow.valflow.level); }
size_t excel_sdk_sizeof_xloper12_flow_tbctrl(void) { return sizeof(((XLOPER12 *)0)->val.flow.valflow.tbctrl); }
size_t excel_sdk_sizeof_xloper12_flow_idsheet(void) { return sizeof(((XLOPER12 *)0)->val.flow.valflow.idSheet); }
size_t excel_sdk_sizeof_xloper12_flow(void) { return sizeof(((XLOPER12 *)0)->val.flow); }
size_t excel_sdk_alignof_xloper12_flow(void) { return __alignof(decltype(((XLOPER12 *)0)->val.flow)); }
size_t excel_sdk_offsetof_xloper12_flow_valflow(void) { return offsetof(XLOPER12, val.flow.valflow); }
size_t excel_sdk_offsetof_xloper12_flow_rw(void) { return offsetof(XLOPER12, val.flow.rw); }
size_t excel_sdk_offsetof_xloper12_flow_col(void) { return offsetof(XLOPER12, val.flow.col); }
size_t excel_sdk_offsetof_xloper12_flow_xlflow(void) { return offsetof(XLOPER12, val.flow.xlflow); }
size_t excel_sdk_sizeof_xloper12_bigdata_handle(void) { return sizeof(((XLOPER12 *)0)->val.bigdata.h); }
size_t excel_sdk_alignof_xloper12_bigdata_handle(void) { return __alignof(decltype(((XLOPER12 *)0)->val.bigdata.h)); }
size_t excel_sdk_sizeof_xloper12_bigdata_lpbdata(void) { return sizeof(((XLOPER12 *)0)->val.bigdata.h.lpbData); }
size_t excel_sdk_sizeof_xloper12_bigdata_hdata(void) { return sizeof(((XLOPER12 *)0)->val.bigdata.h.hdata); }
size_t excel_sdk_sizeof_xloper12_bigdata(void) { return sizeof(((XLOPER12 *)0)->val.bigdata); }
size_t excel_sdk_alignof_xloper12_bigdata(void) { return __alignof(decltype(((XLOPER12 *)0)->val.bigdata)); }
size_t excel_sdk_offsetof_xloper12_bigdata_h(void) { return offsetof(XLOPER12, val.bigdata.h); }
size_t excel_sdk_offsetof_xloper12_bigdata_cbdata(void) { return offsetof(XLOPER12, val.bigdata.cbData); }

SIZEOF_FN(excel_sdk_sizeof_xloper12, XLOPER12)
ALIGNOF_FN(excel_sdk_alignof_xloper12, XLOPER12)
OFFSETOF_FN(excel_sdk_offsetof_xloper12_val, XLOPER12, val)
OFFSETOF_FN(excel_sdk_offsetof_xloper12_xltype, XLOPER12, xltype)

CONST_U32_FN(excel_sdk_xltype_num, xltypeNum)
CONST_U32_FN(excel_sdk_xltype_str, xltypeStr)
CONST_U32_FN(excel_sdk_xltype_bool, xltypeBool)
CONST_U32_FN(excel_sdk_xltype_ref, xltypeRef)
CONST_U32_FN(excel_sdk_xltype_err, xltypeErr)
CONST_U32_FN(excel_sdk_xltype_flow, xltypeFlow)
CONST_U32_FN(excel_sdk_xltype_multi, xltypeMulti)
CONST_U32_FN(excel_sdk_xltype_missing, xltypeMissing)
CONST_U32_FN(excel_sdk_xltype_nil, xltypeNil)
CONST_U32_FN(excel_sdk_xltype_sref, xltypeSRef)
CONST_U32_FN(excel_sdk_xltype_int, xltypeInt)
CONST_U32_FN(excel_sdk_xltype_bigdata, xltypeBigData)
CONST_U32_FN(excel_sdk_xlbit_xlfree, xlbitXLFree)
CONST_U32_FN(excel_sdk_xlbit_dllfree, xlbitDLLFree)

CONST_I32_FN(excel_sdk_xlerr_null, xlerrNull)
CONST_I32_FN(excel_sdk_xlerr_div0, xlerrDiv0)
CONST_I32_FN(excel_sdk_xlerr_value, xlerrValue)
CONST_I32_FN(excel_sdk_xlerr_ref, xlerrRef)
CONST_I32_FN(excel_sdk_xlerr_name, xlerrName)
CONST_I32_FN(excel_sdk_xlerr_num, xlerrNum)
CONST_I32_FN(excel_sdk_xlerr_na, xlerrNA)
CONST_I32_FN(excel_sdk_xlerr_getting_data, xlerrGettingData)

CONST_I32_FN(excel_sdk_xlflow_halt, xlflowHalt)
CONST_I32_FN(excel_sdk_xlflow_goto, xlflowGoto)
CONST_I32_FN(excel_sdk_xlflow_restart, xlflowRestart)
CONST_I32_FN(excel_sdk_xlflow_pause, xlflowPause)
CONST_I32_FN(excel_sdk_xlflow_resume, xlflowResume)

CONST_I32_FN(excel_sdk_xlret_success, xlretSuccess)
CONST_I32_FN(excel_sdk_xlret_abort, xlretAbort)
CONST_I32_FN(excel_sdk_xlret_inv_xlfn, xlretInvXlfn)
CONST_I32_FN(excel_sdk_xlret_inv_count, xlretInvCount)
CONST_I32_FN(excel_sdk_xlret_inv_xloper, xlretInvXloper)
CONST_I32_FN(excel_sdk_xlret_stack_ovfl, xlretStackOvfl)
CONST_I32_FN(excel_sdk_xlret_failed, xlretFailed)
CONST_I32_FN(excel_sdk_xlret_uncalced, xlretUncalced)
CONST_I32_FN(excel_sdk_xlret_not_thread_safe, xlretNotThreadSafe)
CONST_I32_FN(excel_sdk_xlret_inv_async_context, xlretInvAsynchronousContext)
CONST_I32_FN(excel_sdk_xlret_not_cluster_safe, xlretNotClusterSafe)
CONST_I32_FN(excel_sdk_xlhpc_ret_success, xlHpcRetSuccess)
CONST_I32_FN(excel_sdk_xlhpc_ret_session_invalid, xlHpcRetSessionIdInvalid)
CONST_I32_FN(excel_sdk_xlhpc_ret_call_failed, xlHpcRetCallFailed)

CONST_I32_FN(excel_sdk_xlcommand, xlCommand)
CONST_I32_FN(excel_sdk_xlspecial, xlSpecial)
CONST_I32_FN(excel_sdk_xlintl, xlIntl)
CONST_I32_FN(excel_sdk_xlprompt, xlPrompt)
CONST_I32_FN(excel_sdk_xlfree, xlFree)
CONST_I32_FN(excel_sdk_xlstack, xlStack)
CONST_I32_FN(excel_sdk_xlcoerce, xlCoerce)
CONST_I32_FN(excel_sdk_xlset, xlSet)
CONST_I32_FN(excel_sdk_xlsheetid, xlSheetId)
CONST_I32_FN(excel_sdk_xlsheetnm, xlSheetNm)
CONST_I32_FN(excel_sdk_xlabort, xlAbort)
CONST_I32_FN(excel_sdk_xlfnow, xlfNow)
CONST_I32_FN(excel_sdk_xlcontime, xlcOnTime)
CONST_I32_FN(excel_sdk_xlgetname, xlGetName)
CONST_I32_FN(excel_sdk_xlasyncreturn, xlAsyncReturn)
CONST_I32_FN(excel_sdk_xleventregister, xlEventRegister)
CONST_I32_FN(excel_sdk_xlgetinstptr, xlGetInstPtr)
CONST_I32_FN(excel_sdk_xlfsetname, xlfSetName)
CONST_I32_FN(excel_sdk_xlfcaller, xlfCaller)
CONST_I32_FN(excel_sdk_xlfregister, xlfRegister)
CONST_I32_FN(excel_sdk_xlfunregister, xlfUnregister)
CONST_I32_FN(excel_sdk_xlfregisterid, xlfRegisterId)
CONST_I32_FN(excel_sdk_xludf, xlUDF)
CONST_I32_FN(excel_sdk_xlevent_calculation_ended, xleventCalculationEnded)
CONST_I32_FN(excel_sdk_xlevent_calculation_canceled, xleventCalculationCanceled)

typedef int (__cdecl *excel12_signature)(int, LPXLOPER12, int, ...);
typedef int (PASCAL *excel12v_signature)(int, LPXLOPER12, int, LPXLOPER12[]);
typedef int (PASCAL *excel12_entry_signature)(int, int, LPXLOPER12 *, LPXLOPER12);

static_assert(std::is_same<decltype(&Excel12), excel12_signature>::value,
    "Excel12 signature mismatch");
static_assert(std::is_same<decltype(&Excel12v), excel12v_signature>::value,
    "Excel12v signature mismatch");

int excel_sdk_excel12_signatures_compile(void)
{
    excel12_entry_signature entry_ptr = (excel12_entry_signature)0;
    return entry_ptr == 0;
}

typedef int (WINAPI *xl_auto_simple_signature)(void);
typedef LPXLOPER12 (WINAPI *xl_auto_value_signature)(LPXLOPER12);
typedef void (WINAPI *xl_auto_free12_signature)(LPXLOPER12);

int excel_sdk_lifecycle_signatures_compile(
    xl_auto_simple_signature auto_open,
    xl_auto_simple_signature auto_close,
    xl_auto_simple_signature auto_add,
    xl_auto_simple_signature auto_remove,
    xl_auto_value_signature addin_manager_info12,
    xl_auto_value_signature auto_register12,
    xl_auto_free12_signature auto_free12)
{
    return auto_open != 0 && auto_close != 0 && auto_add != 0 &&
        auto_remove != 0 && addin_manager_info12 != 0 &&
        auto_register12 != 0 && auto_free12 != 0;
}

#ifdef __cplusplus
}
#endif
#endif
