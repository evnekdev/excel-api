pub(crate) fn wide_nul(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(Some(0)).collect()
}
