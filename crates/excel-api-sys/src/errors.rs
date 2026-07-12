/// Raw Excel error discriminants used in an `XLOPER12`.
#[repr(i32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum XlError {
    Null = 0,
    Div0 = 7,
    Value = 15,
    Ref = 23,
    Name = 29,
    Num = 36,
    Na = 42,
    GettingData = 43,
}
