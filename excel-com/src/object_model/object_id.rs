/// Stable inventory object identifier.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash, Ord, PartialOrd)]
pub struct ObjectId(&'static str);
impl ObjectId {
    /// Creates an identifier from a static inventory key.
    pub const fn new(value: &'static str) -> Self {
        Self(value)
    }
    /// Returns the static inventory key.
    pub const fn as_str(self) -> &'static str {
        self.0
    }
}
