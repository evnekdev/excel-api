//! Typed text-connection identity where the installed Excel version exposes it.

use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use std::fmt::{Debug, Formatter};

/// Apartment-bound text-connection detail object.
///
/// Text parsing configuration is exposed through [`super::QueryTable`] in
/// this slice; Excel versions differ in standalone TextConnection members.
pub struct TextConnection {
    pub(crate) inner: DispatchObject,
}
impl Debug for TextConnection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TextConnection").field(&self.inner).finish()
    }
}
impl Clone for TextConnection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl TextConnection {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "TextConnection",
            },
        }
    }
}
