/// Raw Excel error discriminants used in an `XLOPER12`.
#[repr(i32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum XlError {
    Null = crate::xlerrNull,
    Div0 = crate::xlerrDiv0,
    Value = crate::xlerrValue,
    Ref = crate::xlerrRef,
    Name = crate::xlerrName,
    Num = crate::xlerrNum,
    Na = crate::xlerrNA,
    GettingData = crate::xlerrGettingData,
}
