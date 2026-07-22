/// Controlled implementation-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ImplementationStatus {
    NotPlanned,
    NotStarted,
    MetadataOnly,
    Stub,
    Partial,
    Implemented,
    Blocked,
    Unsupported,
}
/// Controlled documentation-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DocumentationStatus {
    NotStarted,
    Generated,
    ReviewNeeded,
    Reviewed,
}
/// Controlled testing-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TestStatus {
    NotTested,
    UnitTested,
    IntegrationTested,
    LiveTested,
    Blocked,
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn vocabulary_is_explicit() {
        assert_eq!(ImplementationStatus::Implemented as u8, 5);
        assert_eq!(TestStatus::LiveTested as u8, 3);
    }
}
