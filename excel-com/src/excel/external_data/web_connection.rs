//! Typed web-connection identity without network access.

use crate::excel::DispatchObject;
use std::fmt::{Debug, Formatter};

/// Apartment-bound web-connection detail object.
///
/// The crate never follows endpoints or stores credentials; version-specific
/// web metadata remains intentionally unavailable until it has local evidence.
pub struct WebConnection {
    pub(crate) inner: DispatchObject,
}
impl Debug for WebConnection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("WebConnection").field(&self.inner).finish()
    }
}
impl Clone for WebConnection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
