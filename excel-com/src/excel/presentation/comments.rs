//! Legacy Notes and optional threaded-comment wrappers.
#![allow(missing_docs)]

use super::*;
use windows_sys::Win32::System::Variant::{VT_EMPTY, VT_NULL};

const COMMENTS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Comments",
    count: MemberId::new("excel.comments.count"),
    item: MemberId::new("excel.comments.item"),
    new_enum: MemberId::new("excel.comments.newenum"),
};
const THREADED_COMMENTS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "CommentsThreaded",
    count: MemberId::new("excel.commentsthreaded.count"),
    item: MemberId::new("excel.commentsthreaded.item"),
    new_enum: MemberId::new("excel.commentsthreaded.newenum"),
};

/// Excel's legacy comment object, called a Note in modern Excel UI.
#[derive(Clone, Debug)]
pub struct Comment {
    inner: DispatchObject,
}
impl Comment {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Comment",
            },
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(None);
        a.push_optional(None);
        let value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.comment.text"), false),
            a.into_inner(),
            false,
        )?;
        value.as_string()
    }
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_required(text_bstr(value)?);
        a.push_optional(None);
        a.push_optional(None);
        call(&self.inner, "excel.comment.text", a.into_inner())
    }
    pub fn author(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.comment.author")
    }
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.comment.visible")
    }
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.comment.visible",
            OwnedVariant::bool(value),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.comment.delete", vec![])
    }
}

#[derive(Clone, Debug)]
pub struct Comments {
    inner: DispatchObject,
}
impl Comments {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Comments",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, COMMENTS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<Comment, ExcelComError> {
        comment_item(&self.inner, index)
    }
    /// Returns the one-based legacy comment at `index`.
    pub fn item(&self, index: usize) -> Result<Comment, ExcelComError> {
        self.item_by_index(index)
    }
    pub fn iter(&self) -> Result<CommentsIter, ExcelComError> {
        Ok(CommentsIter {
            enumerator: enumerator(&self.inner, COMMENTS_DESCRIPTOR)?,
            index: 0,
            terminal: false,
        })
    }
}
pub struct CommentsIter {
    enumerator: crate::automation::EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for CommentsIter {
    type Item = Result<Comment, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "Comments", index)
                        .map(Comment::from_dispatch),
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
impl FusedIterator for CommentsIter {}

impl Range {
    pub fn comment(&self) -> Result<Option<Comment>, ExcelComError> {
        optional_dispatch(
            self.dispatch_object(),
            "excel.range.comment",
            Comment::from_dispatch,
        )
    }
    pub fn add_comment(&self, text: Option<&str>) -> Result<Comment, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(text.map(text_bstr).transpose()?);
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.addcomment"), false),
            a.into_inner(),
            false,
        )?;
        Ok(Comment::from_dispatch(value.take_dispatch()?))
    }
    pub fn threaded_comment(&self) -> Result<Option<ThreadedComment>, ExcelComError> {
        optional_dispatch(
            self.dispatch_object(),
            "excel.range.commentthreaded",
            ThreadedComment::from_dispatch,
        )
    }
}
impl Worksheet {
    pub fn comments(&self) -> Result<Comments, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.comments",
            Comments::from_dispatch,
        )
    }
    pub fn threaded_comments(&self) -> Result<ThreadedComments, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.commentsthreaded",
            ThreadedComments::from_dispatch,
        )
    }
}

/// Read-only threaded-comment view. Creating comments requires an Office identity and is intentionally unsupported.
#[derive(Clone, Debug)]
pub struct ThreadedComment {
    inner: DispatchObject,
}
impl ThreadedComment {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "CommentThreaded",
            },
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.commentthreaded.text")
    }
    pub fn date(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.commentthreaded.date")
    }
    pub fn author(&self) -> Result<CommentAuthor, ExcelComError> {
        get_object(
            &self.inner,
            "excel.commentthreaded.author",
            CommentAuthor::from_dispatch,
        )
    }
    pub fn replies(&self) -> Result<ThreadedComments, ExcelComError> {
        get_object(
            &self.inner,
            "excel.commentthreaded.replies",
            ThreadedComments::from_dispatch,
        )
    }
}
#[derive(Clone, Debug)]
pub struct ThreadedComments {
    inner: DispatchObject,
}
impl ThreadedComments {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "CommentsThreaded",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, THREADED_COMMENTS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ThreadedComment, ExcelComError> {
        threaded_comment_item(&self.inner, index)
    }
    /// Returns the one-based threaded comment at `index`.
    pub fn item(&self, index: usize) -> Result<ThreadedComment, ExcelComError> {
        self.item_by_index(index)
    }
    pub fn iter(&self) -> Result<ThreadedCommentsIter, ExcelComError> {
        Ok(ThreadedCommentsIter {
            enumerator: enumerator(&self.inner, THREADED_COMMENTS_DESCRIPTOR)?,
            index: 0,
            terminal: false,
        })
    }
}
pub struct ThreadedCommentsIter {
    enumerator: crate::automation::EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for ThreadedCommentsIter {
    type Item = Result<ThreadedComment, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "CommentsThreaded", index)
                        .map(ThreadedComment::from_dispatch),
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
impl FusedIterator for ThreadedCommentsIter {}

#[derive(Clone, Debug)]
pub struct CommentAuthor {
    inner: DispatchObject,
}
impl CommentAuthor {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Author",
            },
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.author.name")
    }
    pub fn provider_id(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.author.providerid")
    }
    pub fn user_id(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.author.userid")
    }
}

fn optional_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<Option<T>, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    match value.vt() {
        VT_EMPTY | VT_NULL => Ok(None),
        _ => Ok(Some(from(value.take_dispatch()?))),
    }
}
fn comment_item(target: &DispatchObject, index: usize) -> Result<Comment, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(one_based(index, "Comments.Item")?);
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new("excel.comments.item"), false),
        a.into_inner(),
    )?;
    Ok(Comment::from_dispatch(value.take_dispatch()?))
}
fn threaded_comment_item(
    target: &DispatchObject,
    index: usize,
) -> Result<ThreadedComment, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(one_based(index, "CommentsThreaded.Item")?);
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new("excel.commentsthreaded.item"), false),
        a.into_inner(),
    )?;
    Ok(ThreadedComment::from_dispatch(value.take_dispatch()?))
}
