#[cfg(target_os = "windows")]
fn main() {
    use core::mem::{align_of, offset_of, size_of};
    use excel_api_sys::{XLBIT_DLL_FREE, XLBIT_XL_FREE, XLREF12, XLTYPE_INT, XLTYPE_REF, XLOPER12};
    unsafe extern "C" {
        fn excel_sdk_sizeof_xlref12() -> usize;
        fn excel_sdk_alignof_xlref12() -> usize;
        fn excel_sdk_sizeof_fp12_header() -> usize;
        fn excel_sdk_sizeof_xloper12() -> usize;
        fn excel_sdk_alignof_xloper12() -> usize;
        fn excel_sdk_offsetof_xloper12_xltype() -> usize;
        fn excel_sdk_xltype_ref() -> u32;
        fn excel_sdk_xltype_int() -> u32;
        fn excel_sdk_xlbit_xlfree() -> u32;
        fn excel_sdk_xlbit_dllfree() -> u32;
    }
    let mut failures=0usize;
    macro_rules! check {($n:expr,$r:expr,$s:expr)=>{{let r=$r;let s=$s;if r==s{println!("PASS {}: {:?}",$n,r);}else{eprintln!("FAIL {}: Rust={:?}, SDK={:?}",$n,r,s);failures+=1;}}}}
    unsafe {
        check!("sizeof(XLREF12)",size_of::<XLREF12>(),excel_sdk_sizeof_xlref12());
        check!("alignof(XLREF12)",align_of::<XLREF12>(),excel_sdk_alignof_xlref12());
        check!("sizeof(XLOPER12)",size_of::<XLOPER12>(),excel_sdk_sizeof_xloper12());
        check!("alignof(XLOPER12)",align_of::<XLOPER12>(),excel_sdk_alignof_xloper12());
        check!("offsetof(XLOPER12.xltype)",offset_of!(XLOPER12,xltype),excel_sdk_offsetof_xloper12_xltype());
        println!("INFO offsetof(FP12.array): {}",excel_sdk_sizeof_fp12_header());
        check!("xltypeRef",XLTYPE_REF,excel_sdk_xltype_ref());
        check!("xltypeInt",XLTYPE_INT,excel_sdk_xltype_int());
        check!("xlbitXLFree",XLBIT_XL_FREE,excel_sdk_xlbit_xlfree());
        check!("xlbitDLLFree",XLBIT_DLL_FREE,excel_sdk_xlbit_dllfree());
    }
    if failures>0 { std::process::exit(1); }
    println!("Initial Excel SDK ABI checks passed.");
}
#[cfg(not(target_os = "windows"))]
fn main(){println!("Excel ABI C verification skipped: Windows required.");}
