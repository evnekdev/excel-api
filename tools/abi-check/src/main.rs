#[cfg(all(target_os = "windows", target_arch = "x86_64", target_env = "msvc"))]
unsafe extern "system" fn lifecycle_simple_probe() -> i32 {
    1
}

#[cfg(all(target_os = "windows", target_arch = "x86_64", target_env = "msvc"))]
unsafe extern "system" fn lifecycle_value_probe(
    value: excel_api_sys::LPXLOPER12,
) -> excel_api_sys::LPXLOPER12 {
    value
}

#[cfg(all(target_os = "windows", target_arch = "x86_64", target_env = "msvc"))]
unsafe extern "system" fn lifecycle_free_probe(_value: excel_api_sys::LPXLOPER12) {}

#[cfg(all(target_os = "windows", target_arch = "x86_64", target_env = "msvc"))]
fn main() {
    use core::mem::{align_of, offset_of, size_of};
    use excel_api_sys::*;

    unsafe extern "C" {
        fn excel_sdk_sizeof_bool() -> usize;
        fn excel_sdk_sizeof_xchar() -> usize;
        fn excel_sdk_sizeof_rw() -> usize;
        fn excel_sdk_sizeof_col() -> usize;
        fn excel_sdk_sizeof_idsheet() -> usize;
        fn excel_sdk_sizeof_xlref12() -> usize;
        fn excel_sdk_alignof_xlref12() -> usize;
        fn excel_sdk_offsetof_xlref12_rw_first() -> usize;
        fn excel_sdk_offsetof_xlref12_rw_last() -> usize;
        fn excel_sdk_offsetof_xlref12_col_first() -> usize;
        fn excel_sdk_offsetof_xlref12_col_last() -> usize;
        fn excel_sdk_sizeof_xlmref12() -> usize;
        fn excel_sdk_alignof_xlmref12() -> usize;
        fn excel_sdk_offsetof_xlmref12_count() -> usize;
        fn excel_sdk_offsetof_xlmref12_reftbl() -> usize;
        fn excel_sdk_sizeof_fp12() -> usize;
        fn excel_sdk_alignof_fp12() -> usize;
        fn excel_sdk_offsetof_fp12_rows() -> usize;
        fn excel_sdk_offsetof_fp12_columns() -> usize;
        fn excel_sdk_offsetof_fp12_array() -> usize;
        fn excel_sdk_sizeof_xloper12_value() -> usize;
        fn excel_sdk_alignof_xloper12_value() -> usize;
        fn excel_sdk_sizeof_xloper12_num() -> usize;
        fn excel_sdk_sizeof_xloper12_str() -> usize;
        fn excel_sdk_sizeof_xloper12_xbool() -> usize;
        fn excel_sdk_sizeof_xloper12_err() -> usize;
        fn excel_sdk_sizeof_xloper12_w() -> usize;
        fn excel_sdk_sizeof_xloper12_sref() -> usize;
        fn excel_sdk_alignof_xloper12_sref() -> usize;
        fn excel_sdk_offsetof_xloper12_sref_count() -> usize;
        fn excel_sdk_offsetof_xloper12_sref_ref() -> usize;
        fn excel_sdk_sizeof_xloper12_mref() -> usize;
        fn excel_sdk_alignof_xloper12_mref() -> usize;
        fn excel_sdk_offsetof_xloper12_mref_lpmref() -> usize;
        fn excel_sdk_offsetof_xloper12_mref_idsheet() -> usize;
        fn excel_sdk_sizeof_xloper12_array() -> usize;
        fn excel_sdk_alignof_xloper12_array() -> usize;
        fn excel_sdk_offsetof_xloper12_array_lparray() -> usize;
        fn excel_sdk_offsetof_xloper12_array_rows() -> usize;
        fn excel_sdk_offsetof_xloper12_array_columns() -> usize;
        fn excel_sdk_sizeof_xloper12_flow_value() -> usize;
        fn excel_sdk_alignof_xloper12_flow_value() -> usize;
        fn excel_sdk_sizeof_xloper12_flow_level() -> usize;
        fn excel_sdk_sizeof_xloper12_flow_tbctrl() -> usize;
        fn excel_sdk_sizeof_xloper12_flow_idsheet() -> usize;
        fn excel_sdk_sizeof_xloper12_flow() -> usize;
        fn excel_sdk_alignof_xloper12_flow() -> usize;
        fn excel_sdk_offsetof_xloper12_flow_valflow() -> usize;
        fn excel_sdk_offsetof_xloper12_flow_rw() -> usize;
        fn excel_sdk_offsetof_xloper12_flow_col() -> usize;
        fn excel_sdk_offsetof_xloper12_flow_xlflow() -> usize;
        fn excel_sdk_sizeof_xloper12_bigdata_handle() -> usize;
        fn excel_sdk_alignof_xloper12_bigdata_handle() -> usize;
        fn excel_sdk_sizeof_xloper12_bigdata_lpbdata() -> usize;
        fn excel_sdk_sizeof_xloper12_bigdata_hdata() -> usize;
        fn excel_sdk_sizeof_xloper12_bigdata() -> usize;
        fn excel_sdk_alignof_xloper12_bigdata() -> usize;
        fn excel_sdk_offsetof_xloper12_bigdata_h() -> usize;
        fn excel_sdk_offsetof_xloper12_bigdata_cbdata() -> usize;
        fn excel_sdk_sizeof_xloper12() -> usize;
        fn excel_sdk_alignof_xloper12() -> usize;
        fn excel_sdk_offsetof_xloper12_val() -> usize;
        fn excel_sdk_offsetof_xloper12_xltype() -> usize;

        fn excel_sdk_xltype_num() -> u32;
        fn excel_sdk_xltype_str() -> u32;
        fn excel_sdk_xltype_bool() -> u32;
        fn excel_sdk_xltype_ref() -> u32;
        fn excel_sdk_xltype_err() -> u32;
        fn excel_sdk_xltype_flow() -> u32;
        fn excel_sdk_xltype_multi() -> u32;
        fn excel_sdk_xltype_missing() -> u32;
        fn excel_sdk_xltype_nil() -> u32;
        fn excel_sdk_xltype_sref() -> u32;
        fn excel_sdk_xltype_int() -> u32;
        fn excel_sdk_xltype_bigdata() -> u32;
        fn excel_sdk_xlbit_xlfree() -> u32;
        fn excel_sdk_xlbit_dllfree() -> u32;

        fn excel_sdk_xlerr_null() -> i32;
        fn excel_sdk_xlerr_div0() -> i32;
        fn excel_sdk_xlerr_value() -> i32;
        fn excel_sdk_xlerr_ref() -> i32;
        fn excel_sdk_xlerr_name() -> i32;
        fn excel_sdk_xlerr_num() -> i32;
        fn excel_sdk_xlerr_na() -> i32;
        fn excel_sdk_xlerr_getting_data() -> i32;
        fn excel_sdk_xlflow_halt() -> i32;
        fn excel_sdk_xlflow_goto() -> i32;
        fn excel_sdk_xlflow_restart() -> i32;
        fn excel_sdk_xlflow_pause() -> i32;
        fn excel_sdk_xlflow_resume() -> i32;
        fn excel_sdk_xlret_success() -> i32;
        fn excel_sdk_xlret_abort() -> i32;
        fn excel_sdk_xlret_inv_xlfn() -> i32;
        fn excel_sdk_xlret_inv_count() -> i32;
        fn excel_sdk_xlret_inv_xloper() -> i32;
        fn excel_sdk_xlret_stack_ovfl() -> i32;
        fn excel_sdk_xlret_failed() -> i32;
        fn excel_sdk_xlret_uncalced() -> i32;
        fn excel_sdk_xlret_not_thread_safe() -> i32;
        fn excel_sdk_xlret_inv_async_context() -> i32;
        fn excel_sdk_xlret_not_cluster_safe() -> i32;
        fn excel_sdk_xlhpc_ret_success() -> i32;
        fn excel_sdk_xlhpc_ret_session_invalid() -> i32;
        fn excel_sdk_xlhpc_ret_call_failed() -> i32;

        fn excel_sdk_xlcommand() -> i32;
        fn excel_sdk_xlspecial() -> i32;
        fn excel_sdk_xlintl() -> i32;
        fn excel_sdk_xlprompt() -> i32;
        fn excel_sdk_xlfree() -> i32;
        fn excel_sdk_xlstack() -> i32;
        fn excel_sdk_xlcoerce() -> i32;
        fn excel_sdk_xlset() -> i32;
        fn excel_sdk_xlsheetid() -> i32;
        fn excel_sdk_xlsheetnm() -> i32;
        fn excel_sdk_xlabort() -> i32;
        fn excel_sdk_xlfnow() -> i32;
        fn excel_sdk_xlcontime() -> i32;
        fn excel_sdk_xlgetname() -> i32;
        fn excel_sdk_xlasyncreturn() -> i32;
        fn excel_sdk_xleventregister() -> i32;
        fn excel_sdk_xlgetinstptr() -> i32;
        fn excel_sdk_xlfsetname() -> i32;
        fn excel_sdk_xlfcaller() -> i32;
        fn excel_sdk_xlfregister() -> i32;
        fn excel_sdk_xlfunregister() -> i32;
        fn excel_sdk_xlfregisterid() -> i32;
        fn excel_sdk_xludf() -> i32;
        fn excel_sdk_xlevent_calculation_ended() -> i32;
        fn excel_sdk_xlevent_calculation_canceled() -> i32;
        fn excel_sdk_excel12_signatures_compile() -> i32;
        fn excel_sdk_lifecycle_signatures_compile(
            auto_open: XlAutoOpenFn,
            auto_close: XlAutoCloseFn,
            auto_add: XlAutoAddFn,
            auto_remove: XlAutoRemoveFn,
            addin_manager_info12: XlAddInManagerInfo12Fn,
            auto_register12: XlAutoRegister12Fn,
            auto_free12: XlAutoFree12Fn,
        ) -> i32;
    }

    let mut failures = 0usize;
    let mut checks = 0usize;
    macro_rules! check {
        ($name:expr, $rust:expr, $sdk:expr) => {{
            let rust = $rust;
            let sdk = $sdk;
            checks += 1;
            if rust == sdk {
                println!("PASS {}: {:?}", $name, rust);
            } else {
                eprintln!("FAIL {}: Rust={:?}, SDK={:?}", $name, rust, sdk);
                failures += 1;
            }
        }};
    }

    // SAFETY: Every foreign function is a no-argument probe compiled into the
    // checker from the checked-in SDK header. Each returns a scalar value and
    // neither reads nor writes through a pointer.
    unsafe {
        check!("sizeof(BOOL)", size_of::<BOOL>(), excel_sdk_sizeof_bool());
        check!(
            "sizeof(XCHAR)",
            size_of::<XCHAR>(),
            excel_sdk_sizeof_xchar()
        );
        check!("sizeof(RW)", size_of::<RW>(), excel_sdk_sizeof_rw());
        check!("sizeof(COL)", size_of::<COL>(), excel_sdk_sizeof_col());
        check!(
            "sizeof(IDSHEET)",
            size_of::<IDSHEET>(),
            excel_sdk_sizeof_idsheet()
        );

        check!(
            "sizeof(XLREF12)",
            size_of::<XLREF12>(),
            excel_sdk_sizeof_xlref12()
        );
        check!(
            "alignof(XLREF12)",
            align_of::<XLREF12>(),
            excel_sdk_alignof_xlref12()
        );
        check!(
            "offsetof(XLREF12.rwFirst)",
            offset_of!(XLREF12, rwFirst),
            excel_sdk_offsetof_xlref12_rw_first()
        );
        check!(
            "offsetof(XLREF12.rwLast)",
            offset_of!(XLREF12, rwLast),
            excel_sdk_offsetof_xlref12_rw_last()
        );
        check!(
            "offsetof(XLREF12.colFirst)",
            offset_of!(XLREF12, colFirst),
            excel_sdk_offsetof_xlref12_col_first()
        );
        check!(
            "offsetof(XLREF12.colLast)",
            offset_of!(XLREF12, colLast),
            excel_sdk_offsetof_xlref12_col_last()
        );

        check!(
            "sizeof(XLMREF12)",
            size_of::<XLMREF12>(),
            excel_sdk_sizeof_xlmref12()
        );
        check!(
            "alignof(XLMREF12)",
            align_of::<XLMREF12>(),
            excel_sdk_alignof_xlmref12()
        );
        check!(
            "offsetof(XLMREF12.count)",
            offset_of!(XLMREF12, count),
            excel_sdk_offsetof_xlmref12_count()
        );
        check!(
            "offsetof(XLMREF12.reftbl)",
            offset_of!(XLMREF12, reftbl),
            excel_sdk_offsetof_xlmref12_reftbl()
        );

        check!("sizeof(FP12)", size_of::<FP12>(), excel_sdk_sizeof_fp12());
        check!(
            "alignof(FP12)",
            align_of::<FP12>(),
            excel_sdk_alignof_fp12()
        );
        check!(
            "offsetof(FP12.rows)",
            offset_of!(FP12, rows),
            excel_sdk_offsetof_fp12_rows()
        );
        check!(
            "offsetof(FP12.columns)",
            offset_of!(FP12, columns),
            excel_sdk_offsetof_fp12_columns()
        );
        check!(
            "offsetof(FP12.array)",
            offset_of!(FP12, array),
            excel_sdk_offsetof_fp12_array()
        );

        check!(
            "sizeof(XLOPER12.val)",
            size_of::<XLOPER12Value>(),
            excel_sdk_sizeof_xloper12_value()
        );
        check!(
            "alignof(XLOPER12.val)",
            align_of::<XLOPER12Value>(),
            excel_sdk_alignof_xloper12_value()
        );
        check!(
            "sizeof(val.num)",
            size_of::<f64>(),
            excel_sdk_sizeof_xloper12_num()
        );
        check!(
            "sizeof(val.str)",
            size_of::<*mut XCHAR>(),
            excel_sdk_sizeof_xloper12_str()
        );
        check!(
            "sizeof(val.xbool)",
            size_of::<BOOL>(),
            excel_sdk_sizeof_xloper12_xbool()
        );
        check!(
            "sizeof(val.err)",
            size_of::<i32>(),
            excel_sdk_sizeof_xloper12_err()
        );
        check!(
            "sizeof(val.w)",
            size_of::<i32>(),
            excel_sdk_sizeof_xloper12_w()
        );
        check!(
            "sizeof(XLOPER12.val.sref)",
            size_of::<XLOPER12SRef>(),
            excel_sdk_sizeof_xloper12_sref()
        );
        check!(
            "alignof(XLOPER12.val.sref)",
            align_of::<XLOPER12SRef>(),
            excel_sdk_alignof_xloper12_sref()
        );
        check!(
            "offsetof(sref.count)",
            offset_of!(XLOPER12SRef, count),
            excel_sdk_offsetof_xloper12_sref_count()
        );
        check!(
            "offsetof(sref.ref)",
            offset_of!(XLOPER12SRef, reference),
            excel_sdk_offsetof_xloper12_sref_ref()
        );
        check!(
            "sizeof(XLOPER12.val.mref)",
            size_of::<XLOPER12MRef>(),
            excel_sdk_sizeof_xloper12_mref()
        );
        check!(
            "alignof(XLOPER12.val.mref)",
            align_of::<XLOPER12MRef>(),
            excel_sdk_alignof_xloper12_mref()
        );
        check!(
            "offsetof(mref.lpmref)",
            offset_of!(XLOPER12MRef, lpmref),
            excel_sdk_offsetof_xloper12_mref_lpmref()
        );
        check!(
            "offsetof(mref.idSheet)",
            offset_of!(XLOPER12MRef, idSheet),
            excel_sdk_offsetof_xloper12_mref_idsheet()
        );
        check!(
            "sizeof(XLOPER12.val.array)",
            size_of::<XLOPER12Array>(),
            excel_sdk_sizeof_xloper12_array()
        );
        check!(
            "alignof(XLOPER12.val.array)",
            align_of::<XLOPER12Array>(),
            excel_sdk_alignof_xloper12_array()
        );
        check!(
            "offsetof(array.lparray)",
            offset_of!(XLOPER12Array, lparray),
            excel_sdk_offsetof_xloper12_array_lparray()
        );
        check!(
            "offsetof(array.rows)",
            offset_of!(XLOPER12Array, rows),
            excel_sdk_offsetof_xloper12_array_rows()
        );
        check!(
            "offsetof(array.columns)",
            offset_of!(XLOPER12Array, columns),
            excel_sdk_offsetof_xloper12_array_columns()
        );
        check!(
            "sizeof(flow.valflow)",
            size_of::<XLOPER12FlowValue>(),
            excel_sdk_sizeof_xloper12_flow_value()
        );
        check!(
            "alignof(flow.valflow)",
            align_of::<XLOPER12FlowValue>(),
            excel_sdk_alignof_xloper12_flow_value()
        );
        check!(
            "sizeof(flow.level)",
            size_of::<i32>(),
            excel_sdk_sizeof_xloper12_flow_level()
        );
        check!(
            "sizeof(flow.tbctrl)",
            size_of::<i32>(),
            excel_sdk_sizeof_xloper12_flow_tbctrl()
        );
        check!(
            "sizeof(flow.idSheet)",
            size_of::<IDSHEET>(),
            excel_sdk_sizeof_xloper12_flow_idsheet()
        );
        check!(
            "sizeof(XLOPER12.val.flow)",
            size_of::<XLOPER12Flow>(),
            excel_sdk_sizeof_xloper12_flow()
        );
        check!(
            "alignof(XLOPER12.val.flow)",
            align_of::<XLOPER12Flow>(),
            excel_sdk_alignof_xloper12_flow()
        );
        check!(
            "offsetof(flow.valflow)",
            offset_of!(XLOPER12Flow, valflow),
            excel_sdk_offsetof_xloper12_flow_valflow()
        );
        check!(
            "offsetof(flow.rw)",
            offset_of!(XLOPER12Flow, rw),
            excel_sdk_offsetof_xloper12_flow_rw()
        );
        check!(
            "offsetof(flow.col)",
            offset_of!(XLOPER12Flow, col),
            excel_sdk_offsetof_xloper12_flow_col()
        );
        check!(
            "offsetof(flow.xlflow)",
            offset_of!(XLOPER12Flow, xlflow),
            excel_sdk_offsetof_xloper12_flow_xlflow()
        );
        check!(
            "sizeof(bigdata.h)",
            size_of::<XLOPER12BigDataHandle>(),
            excel_sdk_sizeof_xloper12_bigdata_handle()
        );
        check!(
            "alignof(bigdata.h)",
            align_of::<XLOPER12BigDataHandle>(),
            excel_sdk_alignof_xloper12_bigdata_handle()
        );
        check!(
            "sizeof(bigdata.lpbData)",
            size_of::<*mut BYTE>(),
            excel_sdk_sizeof_xloper12_bigdata_lpbdata()
        );
        check!(
            "sizeof(bigdata.hdata)",
            size_of::<HANDLE>(),
            excel_sdk_sizeof_xloper12_bigdata_hdata()
        );
        check!(
            "sizeof(XLOPER12.val.bigdata)",
            size_of::<XLOPER12BigData>(),
            excel_sdk_sizeof_xloper12_bigdata()
        );
        check!(
            "alignof(XLOPER12.val.bigdata)",
            align_of::<XLOPER12BigData>(),
            excel_sdk_alignof_xloper12_bigdata()
        );
        check!(
            "offsetof(bigdata.h)",
            offset_of!(XLOPER12BigData, h),
            excel_sdk_offsetof_xloper12_bigdata_h()
        );
        check!(
            "offsetof(bigdata.cbData)",
            offset_of!(XLOPER12BigData, cbData),
            excel_sdk_offsetof_xloper12_bigdata_cbdata()
        );
        check!(
            "sizeof(XLOPER12)",
            size_of::<XLOPER12>(),
            excel_sdk_sizeof_xloper12()
        );
        check!(
            "alignof(XLOPER12)",
            align_of::<XLOPER12>(),
            excel_sdk_alignof_xloper12()
        );
        check!(
            "offsetof(XLOPER12.val)",
            offset_of!(XLOPER12, val),
            excel_sdk_offsetof_xloper12_val()
        );
        check!(
            "offsetof(XLOPER12.xltype)",
            offset_of!(XLOPER12, xltype),
            excel_sdk_offsetof_xloper12_xltype()
        );

        check!("xltypeNum", xltypeNum, excel_sdk_xltype_num());
        check!("xltypeStr", xltypeStr, excel_sdk_xltype_str());
        check!("xltypeBool", xltypeBool, excel_sdk_xltype_bool());
        check!("xltypeRef", xltypeRef, excel_sdk_xltype_ref());
        check!("xltypeErr", xltypeErr, excel_sdk_xltype_err());
        check!("xltypeFlow", xltypeFlow, excel_sdk_xltype_flow());
        check!("xltypeMulti", xltypeMulti, excel_sdk_xltype_multi());
        check!("xltypeMissing", xltypeMissing, excel_sdk_xltype_missing());
        check!("xltypeNil", xltypeNil, excel_sdk_xltype_nil());
        check!("xltypeSRef", xltypeSRef, excel_sdk_xltype_sref());
        check!("xltypeInt", xltypeInt, excel_sdk_xltype_int());
        check!("xltypeBigData", xltypeBigData, excel_sdk_xltype_bigdata());
        check!("xlbitXLFree", xlbitXLFree, excel_sdk_xlbit_xlfree());
        check!("xlbitDLLFree", xlbitDLLFree, excel_sdk_xlbit_dllfree());

        check!("xlerrNull", xlerrNull, excel_sdk_xlerr_null());
        check!("xlerrDiv0", xlerrDiv0, excel_sdk_xlerr_div0());
        check!("xlerrValue", xlerrValue, excel_sdk_xlerr_value());
        check!("xlerrRef", xlerrRef, excel_sdk_xlerr_ref());
        check!("xlerrName", xlerrName, excel_sdk_xlerr_name());
        check!("xlerrNum", xlerrNum, excel_sdk_xlerr_num());
        check!("xlerrNA", xlerrNA, excel_sdk_xlerr_na());
        check!(
            "xlerrGettingData",
            xlerrGettingData,
            excel_sdk_xlerr_getting_data()
        );
        check!("xlflowHalt", xlflowHalt, excel_sdk_xlflow_halt());
        check!("xlflowGoto", xlflowGoto, excel_sdk_xlflow_goto());
        check!("xlflowRestart", xlflowRestart, excel_sdk_xlflow_restart());
        check!("xlflowPause", xlflowPause, excel_sdk_xlflow_pause());
        check!("xlflowResume", xlflowResume, excel_sdk_xlflow_resume());
        check!("xlretSuccess", xlretSuccess, excel_sdk_xlret_success());
        check!("xlretAbort", xlretAbort, excel_sdk_xlret_abort());
        check!("xlretInvXlfn", xlretInvXlfn, excel_sdk_xlret_inv_xlfn());
        check!("xlretInvCount", xlretInvCount, excel_sdk_xlret_inv_count());
        check!(
            "xlretInvXloper",
            xlretInvXloper,
            excel_sdk_xlret_inv_xloper()
        );
        check!(
            "xlretStackOvfl",
            xlretStackOvfl,
            excel_sdk_xlret_stack_ovfl()
        );
        check!("xlretFailed", xlretFailed, excel_sdk_xlret_failed());
        check!("xlretUncalced", xlretUncalced, excel_sdk_xlret_uncalced());
        check!(
            "xlretNotThreadSafe",
            xlretNotThreadSafe,
            excel_sdk_xlret_not_thread_safe()
        );
        check!(
            "xlretInvAsynchronousContext",
            xlretInvAsynchronousContext,
            excel_sdk_xlret_inv_async_context()
        );
        check!(
            "xlretNotClusterSafe",
            xlretNotClusterSafe,
            excel_sdk_xlret_not_cluster_safe()
        );
        check!(
            "xlHpcRetSuccess",
            xlHpcRetSuccess,
            excel_sdk_xlhpc_ret_success()
        );
        check!(
            "xlHpcRetSessionIdInvalid",
            xlHpcRetSessionIdInvalid,
            excel_sdk_xlhpc_ret_session_invalid()
        );
        check!(
            "xlHpcRetCallFailed",
            xlHpcRetCallFailed,
            excel_sdk_xlhpc_ret_call_failed()
        );

        check!("xlCommand", xlCommand, excel_sdk_xlcommand());
        check!("xlSpecial", xlSpecial, excel_sdk_xlspecial());
        check!("xlIntl", xlIntl, excel_sdk_xlintl());
        check!("xlPrompt", xlPrompt, excel_sdk_xlprompt());
        check!("xlFree", xlFree, excel_sdk_xlfree());
        check!("xlStack", xlStack, excel_sdk_xlstack());
        check!("xlCoerce", xlCoerce, excel_sdk_xlcoerce());
        check!("xlSet", xlSet, excel_sdk_xlset());
        check!("xlSheetId", xlSheetId, excel_sdk_xlsheetid());
        check!("xlSheetNm", xlSheetNm, excel_sdk_xlsheetnm());
        check!("xlAbort", xlAbort, excel_sdk_xlabort());
        check!("xlfNow", xlfNow, excel_sdk_xlfnow());
        check!("xlcOnTime", xlcOnTime, excel_sdk_xlcontime());
        check!("xlGetName", xlGetName, excel_sdk_xlgetname());
        check!("xlAsyncReturn", xlAsyncReturn, excel_sdk_xlasyncreturn());
        check!(
            "xlEventRegister",
            xlEventRegister,
            excel_sdk_xleventregister()
        );
        check!("xlGetInstPtr", xlGetInstPtr, excel_sdk_xlgetinstptr());
        check!("xlfSetName", xlfSetName, excel_sdk_xlfsetname());
        check!("xlfCaller", xlfCaller, excel_sdk_xlfcaller());
        check!("xlfRegister", xlfRegister, excel_sdk_xlfregister());
        check!("xlfUnregister", xlfUnregister, excel_sdk_xlfunregister());
        check!("xlfRegisterId", xlfRegisterId, excel_sdk_xlfregisterid());
        check!("xlUDF", xlUDF, excel_sdk_xludf());
        check!(
            "xleventCalculationEnded",
            xleventCalculationEnded,
            excel_sdk_xlevent_calculation_ended()
        );
        check!(
            "xleventCalculationCanceled",
            xleventCalculationCanceled,
            excel_sdk_xlevent_calculation_canceled()
        );
        check!(
            "Excel12/Excel12v C signatures",
            1,
            excel_sdk_excel12_signatures_compile()
        );
        check!(
            "Excel 12 lifecycle callback signatures",
            1,
            excel_sdk_lifecycle_signatures_compile(
                lifecycle_simple_probe,
                lifecycle_simple_probe,
                lifecycle_simple_probe,
                lifecycle_simple_probe,
                lifecycle_value_probe,
                lifecycle_value_probe,
                lifecycle_free_probe,
            )
        );
    }

    check!(
        "Q value-only registration code",
        XLL_TYPE_XLOPER12_VALUE,
        "Q"
    );
    check!(
        "U reference registration code",
        XLL_TYPE_XLOPER12_REFERENCE,
        "U"
    );
    check!(
        "counted XCHAR registration code",
        XLL_TYPE_XCHAR_COUNTED,
        "D%"
    );
    check!(
        "NUL XCHAR registration code",
        XLL_TYPE_XCHAR_NULL_TERMINATED,
        "C%"
    );
    check!("async handle registration code", XLL_TYPE_ASYNC_HANDLE, "X");
    check!("async void registration code", XLL_TYPE_ASYNC_VOID, ">");
    check!("thread-safe modifier", XLL_MODIFIER_THREAD_SAFE, '$');
    check!("macro-sheet modifier", XLL_MODIFIER_MACRO_SHEET, '#');
    check!("volatile modifier", XLL_MODIFIER_VOLATILE, '!');
    check!("cluster-safe modifier", XLL_MODIFIER_CLUSTER_SAFE, '&');
    check!("Excel 12 maximum arguments", EXCEL12_MAX_ARGUMENTS, 255);
    check!("Excel 12 maximum rows", EXCEL12_MAX_ROWS, 1_048_576);
    check!("Excel 12 maximum columns", EXCEL12_MAX_COLUMNS, 16_384);
    check!(
        "Excel 12 maximum string units",
        EXCEL12_MAX_STRING_CODE_UNITS,
        32_767
    );

    if failures > 0 {
        eprintln!("{failures} of {checks} Excel SDK ABI checks failed.");
        std::process::exit(1);
    }
    println!("All {checks} Excel SDK ABI checks passed.");
}

#[cfg(not(all(target_os = "windows", target_arch = "x86_64", target_env = "msvc")))]
fn main() {
    println!("Excel ABI C verification skipped: Windows x86_64 MSVC required.");
}
