//! Hyperlink wrappers.
#![allow(missing_docs)]

use super::*;
use windows_sys::Win32::System::Variant::{VT_EMPTY, VT_NULL};

const HYPERLINKS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Hyperlinks",
    count: MemberId::new("excel.hyperlinks.count"),
    item: MemberId::new("excel.hyperlinks.item"),
    new_enum: MemberId::new("excel.hyperlinks.newenum"),
};

/// Arguments for [`Hyperlinks::add`].
///
/// `anchor` is required by Excel. At least one of `address` or `sub_address`
/// must be meaningful to Excel; this wrapper leaves that semantic validation
/// to the host.
#[derive(Clone, Debug)]
pub struct HyperlinkAddOptions<'a> {
    /// The range or shape anchor for the new hyperlink.
    pub anchor: &'a Range,
    /// Optional external address. Live acceptance tests use no external URL.
    pub address: Option<&'a str>,
    /// Optional workbook-relative destination such as `Sheet1!A1`.
    pub sub_address: Option<&'a str>,
    /// Optional Excel screen tip.
    pub screen_tip: Option<&'a str>,
    /// Optional text Excel displays at the anchor.
    pub text_to_display: Option<&'a str>,
}

#[derive(Clone, Debug)]
pub struct Hyperlinks {
    inner: DispatchObject,
}
impl Hyperlinks {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Hyperlinks",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, HYPERLINKS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<Hyperlink, ExcelComError> {
        hyperlink_item(&self.inner, index)
    }
    /// Returns the one-based hyperlink at `index`.
    pub fn item(&self, index: usize) -> Result<Hyperlink, ExcelComError> {
        self.item_by_index(index)
    }
    pub fn iter(&self) -> Result<HyperlinksIter, ExcelComError> {
        Ok(HyperlinksIter {
            enumerator: enumerator(&self.inner, HYPERLINKS_DESCRIPTOR)?,
            index: 0,
            terminal: false,
        })
    }
    /// Adds a hyperlink using Excel's positional `Hyperlinks.Add` contract.
    pub fn add(&self, options: &HyperlinkAddOptions<'_>) -> Result<Hyperlink, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_object(options.anchor.dispatch_object());
        a.push_optional(options.address.map(text_bstr).transpose()?);
        a.push_optional(options.sub_address.map(text_bstr).transpose()?);
        a.push_optional(options.screen_tip.map(text_bstr).transpose()?);
        a.push_optional(options.text_to_display.map(text_bstr).transpose()?);
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.hyperlinks.add"), false),
            a.into_inner(),
            false,
        )?;
        Ok(Hyperlink::from_dispatch(value.take_dispatch()?))
    }
}
pub struct HyperlinksIter {
    enumerator: crate::automation::EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for HyperlinksIter {
    type Item = Result<Hyperlink, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "Hyperlinks", index)
                        .map(Hyperlink::from_dispatch),
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
impl FusedIterator for HyperlinksIter {}

#[derive(Clone, Debug)]
pub struct Hyperlink {
    inner: DispatchObject,
}
impl Hyperlink {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Hyperlink",
            },
        }
    }
    pub fn address(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.hyperlink.address")
    }
    pub fn set_address(&self, value: Option<&str>) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hyperlink.address",
            value
                .map(text_bstr)
                .transpose()?
                .unwrap_or_else(OwnedVariant::empty),
        )
    }
    pub fn sub_address(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.hyperlink.subaddress")
    }
    pub fn set_sub_address(&self, value: Option<&str>) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hyperlink.subaddress",
            value
                .map(text_bstr)
                .transpose()?
                .unwrap_or_else(OwnedVariant::empty),
        )
    }
    pub fn screen_tip(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.hyperlink.screentip")
    }
    pub fn set_screen_tip(&self, value: Option<&str>) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hyperlink.screentip",
            value
                .map(text_bstr)
                .transpose()?
                .unwrap_or_else(OwnedVariant::empty),
        )
    }
    pub fn text_to_display(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.hyperlink.texttodisplay")
    }
    pub fn set_text_to_display(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hyperlink.texttodisplay",
            text_bstr(value)?,
        )
    }
    pub fn range(&self) -> Result<Option<Range>, ExcelComError> {
        let mut value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.hyperlink.range"), false),
            vec![],
        )?;
        if matches!(value.vt(), VT_EMPTY | VT_NULL) {
            Ok(None)
        } else {
            Ok(Some(Range::from_dispatch(value.take_dispatch()?)))
        }
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.hyperlink.delete", vec![])
    }
}
impl Range {
    pub fn hyperlinks(&self) -> Result<Hyperlinks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.range.hyperlinks",
            Hyperlinks::from_dispatch,
        )
    }
}
impl Worksheet {
    pub fn hyperlinks(&self) -> Result<Hyperlinks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.hyperlinks",
            Hyperlinks::from_dispatch,
        )
    }
}

fn hyperlink_item(target: &DispatchObject, index: usize) -> Result<Hyperlink, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(one_based(index, "Hyperlinks.Item")?);
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new("excel.hyperlinks.item"), false),
        a.into_inner(),
    )?;
    Ok(Hyperlink::from_dispatch(value.take_dispatch()?))
}
