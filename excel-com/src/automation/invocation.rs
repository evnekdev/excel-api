use crate::object_model::MemberId;
/// Explicit Automation invocation classification from structural and runtime evidence.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum MemberKind {
    PropertyGet,
    PropertyPut,
    PropertyPutRef,
    Method,
}
impl MemberKind {
    pub(crate) const fn flags(self) -> u16 {
        match self {
            Self::PropertyGet => 2,
            Self::PropertyPut => 4,
            Self::PropertyPutRef => 8,
            Self::Method => 1,
        }
    }
}
/// One implementation descriptor; wrappers use these rather than scattered literals.
#[derive(Clone, Copy, Debug)]
pub(crate) struct MemberDescriptor {
    pub(crate) id: MemberId,
    pub(crate) name: &'static str,
    pub(crate) kind: MemberKind,
}
