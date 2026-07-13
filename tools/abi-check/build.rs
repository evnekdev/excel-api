use std::{env, path::PathBuf};
fn main() {
    println!("cargo:rerun-if-env-changed=EXCEL_XLL_SDK_DIR");
    println!("cargo:rerun-if-changed=native/layout_check.c");
    if !cfg!(target_os = "windows") {
        println!("cargo:warning=Excel ABI C checker is Windows-only");
        return;
    }
    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let sdk = env::var_os("EXCEL_XLL_SDK_DIR").map(PathBuf::from)
        .unwrap_or_else(|| manifest.join("../Excel2013XLLSDK"));
    let include = sdk.join("INCLUDE");
    let header = include.join("XLCALL.H");
    if !header.is_file() {
        panic!("XLCALL.H not found at {}", header.display());
    }
    println!("cargo:rerun-if-changed={}", header.display());
    cc::Build::new().file("native/layout_check.c").include(include)
        .warnings(true).compile("excel_api_layout_check");
}
