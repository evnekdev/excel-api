use std::{env, path::PathBuf};
fn main() {
    println!("cargo:rerun-if-env-changed=EXCEL_XLL_SDK_DIR");
    println!("cargo:rerun-if-changed=native/layout_check.cpp");
    let target_os = env::var("CARGO_CFG_TARGET_OS").unwrap();
    let target_arch = env::var("CARGO_CFG_TARGET_ARCH").unwrap();
    let target_env = env::var("CARGO_CFG_TARGET_ENV").unwrap();
    if target_os != "windows" {
        println!("cargo:warning=Excel ABI C checker requires Windows x86_64 MSVC");
        return;
    }
    assert_eq!(target_arch, "x86_64", "Excel ABI checker requires x86_64");
    assert_eq!(
        target_env, "msvc",
        "Excel ABI checker requires the MSVC ABI"
    );
    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let sdk = env::var_os("EXCEL_XLL_SDK_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| manifest.join("../Excel2013XLLSDK"));
    let include = sdk.join("INCLUDE");
    let header = include.join("XLCALL.H");
    if !header.is_file() {
        panic!("XLCALL.H not found at {}", header.display());
    }
    println!("cargo:rerun-if-changed={}", header.display());
    cc::Build::new()
        .file("native/layout_check.cpp")
        .cpp(true)
        .include(include)
        .warnings(true)
        .compile("excel_api_layout_check");
}
