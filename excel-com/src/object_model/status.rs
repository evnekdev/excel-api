/// Controlled implementation-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ImplementationStatus {
    /// The member is intentionally not on the implementation roadmap.
    NotPlanned,
    /// The member has not been started.
    NotStarted,
    /// Only its extracted metadata is available.
    MetadataOnly,
    /// A public or internal stub exists without complete behavior.
    Stub,
    /// Some supported behavior exists but the contract is incomplete.
    Partial,
    /// The selected implementation contract is complete.
    Implemented,
    /// Completion is prevented by a documented blocker.
    Blocked,
    /// The member is deliberately unsupported.
    Unsupported,
}
/// Controlled documentation-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DocumentationStatus {
    /// No member documentation is available.
    NotStarted,
    /// Documentation was generated from inventory metadata.
    Generated,
    /// Generated documentation requires human review.
    ReviewNeeded,
    /// Documentation has been reviewed.
    Reviewed,
}
/// Controlled testing-completion vocabulary shared with inventory metadata.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TestStatus {
    /// No tests cover the member.
    NotTested,
    /// Deterministic unit tests cover the member.
    UnitTested,
    /// Integration tests cover the member without a live Excel server.
    IntegrationTested,
    /// An opt-in live Excel test covers the member.
    LiveTested,
    /// Testing is prevented by a documented blocker.
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
