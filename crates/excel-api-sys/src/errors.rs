/// Raw Excel error discriminants used in an `XLOPER12`.
#[repr(i32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum XlError {
    /// Intersection of two ranges is empty.
    Null = crate::xlerrNull,
    /// Division by zero.
    Div0 = crate::xlerrDiv0,
    /// An argument or operation has an invalid value.
    Value = crate::xlerrValue,
    /// A reference is invalid.
    Ref = crate::xlerrRef,
    /// Excel cannot resolve a name.
    Name = crate::xlerrName,
    /// A number is invalid or outside Excel's supported range.
    Num = crate::xlerrNum,
    /// A value is not available.
    Na = crate::xlerrNA,
    /// Excel is still obtaining external data.
    GettingData = crate::xlerrGettingData,
}
