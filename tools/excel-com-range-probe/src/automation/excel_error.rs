//! Lossless physical Excel error SCODE handling.

#[allow(dead_code)]
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub(crate) struct ExcelError(pub(super) i32);

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub(crate) enum ExcelErrorKind {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NotAvailable,
    Other(i32),
}

impl ExcelError {
    pub(crate) const NULL: Self = Self(0x800A_07D0_u32 as i32);
    pub(crate) const DIV0: Self = Self(0x800A_07D7_u32 as i32);
    pub(crate) const VALUE: Self = Self(0x800A_07DF_u32 as i32);
    #[allow(dead_code)]
    pub(crate) const REF: Self = Self(0x800A_07E7_u32 as i32);
    #[allow(dead_code)]
    pub(crate) const NAME: Self = Self(0x800A_07ED_u32 as i32);
    #[allow(dead_code)]
    pub(crate) const NUM: Self = Self(0x800A_07F4_u32 as i32);
    pub(crate) const NOT_AVAILABLE: Self = Self(0x800A_07FA_u32 as i32);

    pub(crate) const fn from_scode(scode: i32) -> Self { Self(scode) }
    pub(crate) const fn scode(self) -> i32 { self.0 }

    #[allow(dead_code)]
    pub(crate) fn kind(self) -> ExcelErrorKind {
        match self {
            Self::NULL => ExcelErrorKind::Null,
            Self::DIV0 => ExcelErrorKind::Div0,
            Self::VALUE => ExcelErrorKind::Value,
            Self::REF => ExcelErrorKind::Ref,
            Self::NAME => ExcelErrorKind::Name,
            Self::NUM => ExcelErrorKind::Num,
            Self::NOT_AVAILABLE => ExcelErrorKind::NotAvailable,
            Self(value) => ExcelErrorKind::Other(value),
        }
    }

    #[allow(dead_code)]
    pub(crate) fn excel_number(self) -> Option<u16> {
        ((self.0 as u32 & 0xffff_0000) == 0x800A_0000)
            .then_some((self.0 as u32 & 0xffff) as u16)
    }

    pub(super) fn valid_for_direct_write(self) -> bool { self.0 < 0 }
}
