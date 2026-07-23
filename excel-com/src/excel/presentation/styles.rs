//! Styles and theme formatting wrappers.
#![allow(missing_docs)]

use super::*;
use crate::excel::{Borders, Font, Interior, MixedValue};

#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct ThemeColor(i32);
impl ThemeColor {
    pub const LIGHT1: Self = Self(0);
    pub const DARK1: Self = Self(1);
    pub const LIGHT2: Self = Self(2);
    pub const DARK2: Self = Self(3);
    pub const ACCENT1: Self = Self(4);
    pub const ACCENT2: Self = Self(5);
    pub const ACCENT3: Self = Self(6);
    pub const ACCENT4: Self = Self(7);
    pub const ACCENT5: Self = Self(8);
    pub const ACCENT6: Self = Self(9);
    pub const HYPERLINK: Self = Self(10);
    pub const FOLLOWED_HYPERLINK: Self = Self(11);
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }
    pub const fn raw(self) -> i32 {
        self.0
    }
}

#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct ThemeFont(i32);
impl ThemeFont {
    pub const NONE: Self = Self(0);
    pub const MAJOR: Self = Self(1);
    pub const MINOR: Self = Self(2);
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }
    pub const fn raw(self) -> i32 {
        self.0
    }
}

const STYLES_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Styles",
    count: MemberId::new("excel.styles.count"),
    item: MemberId::new("excel.styles.item"),
    new_enum: MemberId::new("excel.styles.newenum"),
};

#[derive(Clone, Debug)]
pub struct Styles {
    inner: DispatchObject,
}
impl Styles {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Styles",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, STYLES_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<Style, ExcelComError> {
        style_item(&self.inner, index)
    }
    /// Returns the one-based style item at `index`.
    pub fn item(&self, index: usize) -> Result<Style, ExcelComError> {
        self.item_by_index(index)
    }
    pub fn item_by_name(&self, name: &str) -> Result<Style, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_required(text_bstr(name)?);
        let mut value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.styles.item"), false),
            a.into_inner(),
        )?;
        Ok(Style::from_dispatch(value.take_dispatch()?))
    }
    pub fn iter(&self) -> Result<StylesIter, ExcelComError> {
        Ok(StylesIter {
            enumerator: enumerator(&self.inner, STYLES_DESCRIPTOR)?,
            index: 0,
            terminal: false,
        })
    }
    pub fn add(&self, name: &str, based_on: Option<&Range>) -> Result<Style, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_required(text_bstr(name)?);
        a.push_optional_object(based_on.map(Range::dispatch_object));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.styles.add"), false),
            a.into_inner(),
            false,
        )?;
        Ok(Style::from_dispatch(value.take_dispatch()?))
    }
}
pub struct StylesIter {
    enumerator: crate::automation::EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for StylesIter {
    type Item = Result<Style, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "Styles", index)
                        .and_then(Style::from_dispatch_result),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for StylesIter {}

#[derive(Clone, Debug)]
pub struct Style {
    inner: DispatchObject,
}
impl Style {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Style",
            },
        }
    }
    fn from_dispatch_result(dispatch: ComPtr<Dispatch>) -> Result<Self, ExcelComError> {
        Ok(Self::from_dispatch(dispatch))
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.style.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.style.name", text_bstr(value)?)
    }
    pub fn built_in(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.builtin")
    }
    pub fn include_number(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includenumber")
    }
    pub fn set_include_number(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includenumber",
            OwnedVariant::bool(value),
        )
    }
    pub fn include_font(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includefont")
    }
    pub fn set_include_font(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includefont",
            OwnedVariant::bool(value),
        )
    }
    pub fn include_alignment(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includealignment")
    }
    pub fn set_include_alignment(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includealignment",
            OwnedVariant::bool(value),
        )
    }
    pub fn include_border(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includeborder")
    }
    pub fn set_include_border(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includeborder",
            OwnedVariant::bool(value),
        )
    }
    pub fn include_protection(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includeprotection")
    }
    pub fn set_include_protection(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includeprotection",
            OwnedVariant::bool(value),
        )
    }
    pub fn include_patterns(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.style.includepatterns")
    }
    pub fn set_include_patterns(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.style.includepatterns",
            OwnedVariant::bool(value),
        )
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_object(&self.inner, "excel.style.font", Font::from_dispatch)
    }
    pub fn interior(&self) -> Result<Interior, ExcelComError> {
        get_object(&self.inner, "excel.style.interior", Interior::from_dispatch)
    }
    pub fn borders(&self) -> Result<Borders, ExcelComError> {
        get_object(&self.inner, "excel.style.borders", Borders::from_dispatch)
    }
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.style.numberformat")
    }
    pub fn set_number_format(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.style.numberformat", text_bstr(value)?)
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.style.delete", vec![])
    }
}

impl Workbook {
    pub fn styles(&self) -> Result<Styles, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.workbook.styles",
            Styles::from_dispatch,
        )
    }
}
impl Range {
    /// Returns Excel's `Range.Style` name.
    pub fn style(&self) -> Result<MixedValue<String>, ExcelComError> {
        self.style_name()
    }
    pub fn style_name(&self) -> Result<MixedValue<String>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.style", |value| {
            value.as_string().map(MixedValue::Uniform)
        })
    }
    pub fn set_style_by_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.style",
            text_bstr(value)?,
        )
    }
    pub fn set_style(&self, value: &Style) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.style",
            OwnedVariant::dispatch_borrowed(&value.inner.dispatch),
        )
    }
}

impl Tab {
    pub fn theme_color(&self) -> Result<ThemeColor, ExcelComError> {
        Ok(ThemeColor::from_raw(get_i32(
            &self.inner,
            "excel.tab.themecolor",
            "Tab.ThemeColor",
        )?))
    }
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.tab.themecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn tint_and_shade(&self) -> Result<f64, ExcelComError> {
        get_f64(&self.inner, "excel.tab.tintandshade", "Tab.TintAndShade")
    }
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        tint_put(&self.inner, "excel.tab.tintandshade", value)
    }
}

fn style_item(target: &DispatchObject, index: usize) -> Result<Style, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(one_based(index, "Styles.Item")?);
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new("excel.styles.item"), false),
        a.into_inner(),
    )?;
    Ok(Style::from_dispatch(value.take_dispatch()?))
}

fn tint_put(target: &DispatchObject, id: &'static str, value: f64) -> Result<(), ExcelComError> {
    finite(value)?;
    if !(-1.0..=1.0).contains(&value) {
        return Err(ExcelComError::Unsupported {
            detail: "TintAndShade must be between -1.0 and 1.0",
        });
    }
    put(target, id, OwnedVariant::f64(value))
}
