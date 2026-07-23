//! Focused drawing implementation component.
#![allow(missing_docs)]
use super::export::copy_picture;
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// Excel Shapes on a worksheet or chart.
pub struct Shapes {
    inner: DispatchObject,
}
impl Debug for Shapes {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Shapes").field(&self.inner).finish()
    }
}
impl Clone for Shapes {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Shapes {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Shapes", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SHAPES)
    }
    pub fn item_by_index(&self, index: usize) -> Result<Shape, ExcelComError> {
        collection_item(&self.inner, SHAPES, index, Shape::from_dispatch)
    }
    pub fn item_by_name(&self, name: &str) -> Result<Shape, ExcelComError> {
        collection_named(&self.inner, SHAPES, name, Shape::from_dispatch)
    }
    pub fn iter(&self) -> Result<ShapesIter, ExcelComError> {
        Ok(ShapesIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SHAPES.new_enum),
                SHAPES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add_shape(
        &self,
        shape_type: AutoShapeType,
        bounds: ShapeBounds,
    ) -> Result<Shape, ExcelComError> {
        shape_bounds(bounds)?;
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(shape_type.raw()));
        for value in [bounds.left, bounds.top, bounds.width, bounds.height] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addshape"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_line(&self, start: ShapePoint, end: ShapePoint) -> Result<Shape, ExcelComError> {
        for value in [start.x, start.y, end.x, end.y] {
            finite(value, "shape geometry must be finite")?;
        }
        let mut args = PositionalArguments::new();
        for value in [start.x, start.y, end.x, end.y] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addline"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_text_box(&self, options: &TextBoxAddOptions<'_>) -> Result<Shape, ExcelComError> {
        shape_bounds(options.bounds)?;
        let _ = text_bstr(options.text)?;
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.orientation.raw()));
        for value in [
            options.bounds.left,
            options.bounds.top,
            options.bounds.width,
            options.bounds.height,
        ] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addtextbox"), false),
            args.into_inner(),
            false,
        )?;
        let shape = Shape::from_dispatch(value.take_dispatch()?);
        if let Some(frame) = shape.text_frame()? {
            frame.text_range()?.set_text(options.text)?;
        }
        Ok(shape)
    }
    pub fn add_picture(&self, options: &PictureAddOptions<'_>) -> Result<Shape, ExcelComError> {
        if !options.path.is_file() {
            return Err(ExcelComError::InvalidPath {
                detail: "picture path must name an existing local file",
            });
        }
        if !options.link_to_file && !options.save_with_document {
            return Err(ExcelComError::Unsupported {
                detail: "picture must be linked or saved with the workbook",
            });
        }
        shape_bounds(options.bounds)?;
        let mut args = PositionalArguments::new();
        args.push_result(path_bstr(options.path))?;
        args.push_required(OwnedVariant::i32(if options.link_to_file { -1 } else { 0 }));
        args.push_required(OwnedVariant::i32(if options.save_with_document {
            -1
        } else {
            0
        }));
        for value in [
            options.bounds.left,
            options.bounds.top,
            options.bounds.width,
            options.bounds.height,
        ] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addpicture"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_shape_over_range(
        &self,
        shape_type: AutoShapeType,
        range: &Range,
    ) -> Result<Shape, ExcelComError> {
        self.add_shape(shape_type, range.shape_bounds()?)
    }
    pub fn add_picture_over_range(
        &self,
        path: &Path,
        range: &Range,
    ) -> Result<Shape, ExcelComError> {
        self.add_picture(&PictureAddOptions {
            path,
            link_to_file: false,
            save_with_document: true,
            bounds: range.shape_bounds()?,
        })
    }
}
pub struct ShapesIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for ShapesIter {
    type Item = Result<Shape, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(enumerated_dispatch(&mut value, "Shapes", index).map(Shape::from_dispatch))
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ShapesIter {}

/// An Excel or Office shape.
pub struct Shape {
    inner: DispatchObject,
}
impl Debug for Shape {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Shape").field(&self.inner).finish()
    }
}
impl Clone for Shape {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Shape {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Shape", value),
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.shape.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.shape.name", text_bstr(value)?)
    }
    pub fn shape_type(&self) -> Result<ShapeType, ExcelComError> {
        Ok(ShapeType::from_raw(get_i32(
            &self.inner,
            "excel.shape.type",
            "Shape.Type was not an integer",
        )?))
    }
    pub fn left(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.left",
            "Shape.Left was not numeric",
        )
    }
    pub fn set_left(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape position must be finite")?;
        put(&self.inner, "excel.shape.left", OwnedVariant::f64(value))
    }
    pub fn top(&self) -> Result<f64, ExcelComError> {
        get_f64(&self.inner, "excel.shape.top", "Shape.Top was not numeric")
    }
    pub fn set_top(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape position must be finite")?;
        put(&self.inner, "excel.shape.top", OwnedVariant::f64(value))
    }
    pub fn width(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.width",
            "Shape.Width was not numeric",
        )
    }
    pub fn set_width(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "shape width must be positive")?;
        put(&self.inner, "excel.shape.width", OwnedVariant::f64(value))
    }
    pub fn height(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.height",
            "Shape.Height was not numeric",
        )
    }
    pub fn set_height(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "shape height must be positive")?;
        put(&self.inner, "excel.shape.height", OwnedVariant::f64(value))
    }
    pub fn rotation(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.rotation",
            "Shape.Rotation was not numeric",
        )
    }
    pub fn set_rotation(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape rotation must be finite")?;
        put(
            &self.inner,
            "excel.shape.rotation",
            OwnedVariant::f64(value),
        )
    }
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.shape.visible")
    }
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.visible",
            OwnedVariant::bool(value),
        )
    }
    pub fn placement(&self) -> Result<ShapePlacement, ExcelComError> {
        Ok(ShapePlacement::from_raw(get_i32(
            &self.inner,
            "excel.shape.placement",
            "Shape.Placement was not an integer",
        )?))
    }
    pub fn set_placement(&self, value: ShapePlacement) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.placement",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn lock_aspect_ratio(&self) -> Result<bool, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.shape.lockaspectratio",
            "Shape.LockAspectRatio was not an integer",
        )
        .map(|value| value != 0)
    }
    pub fn set_lock_aspect_ratio(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.lockaspectratio",
            OwnedVariant::i32(if value { -1 } else { 0 }),
        )
    }
    pub fn fill(&self) -> Result<FillFormat, ExcelComError> {
        get_dispatch(&self.inner, "excel.shape.fill", FillFormat::from_dispatch)
    }
    pub fn line(&self) -> Result<LineFormat, ExcelComError> {
        get_dispatch(&self.inner, "excel.shape.line", LineFormat::from_dispatch)
    }
    pub fn z_order(&self, value: ZOrderCommand) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.shape.zorder",
            vec![OwnedVariant::i32(value.raw())],
        )
    }
    pub fn copy(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.copy", vec![])
    }
    pub fn cut(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.cut", vec![])
    }
    pub fn text_frame(&self) -> Result<Option<TextFrame>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.shape.textframe2",
            TextFrame::from_dispatch,
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.delete", vec![])
    }
}
/// A typed Office text frame returned by `Shape.TextFrame2`.
pub struct TextFrame {
    inner: DispatchObject,
}
impl Debug for TextFrame {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TextFrame").field(&self.inner).finish()
    }
}
impl TextFrame {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("TextFrame2", value),
        }
    }
    pub fn text_range(&self) -> Result<TextRange, ExcelComError> {
        get_dispatch(
            &self.inner,
            "office.textframe2.textrange",
            TextRange::from_dispatch,
        )
    }
}
/// Typed Office text range for a Shape text frame.
pub struct TextRange {
    inner: DispatchObject,
}
impl Debug for TextRange {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TextRange").field(&self.inner).finish()
    }
}
impl TextRange {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("TextRange2", value),
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "office.textrange2.text")
    }
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "office.textrange2.text", text_bstr(value)?)
    }
}

impl Range {
    /// Returns point-based bounds derived from Excel's Range geometry.
    pub fn shape_bounds(&self) -> Result<ShapeBounds, ExcelComError> {
        let left = get_f64(
            self.dispatch_object(),
            "excel.range.left",
            "Range.Left was not numeric",
        )?;
        let top = get_f64(
            self.dispatch_object(),
            "excel.range.top",
            "Range.Top was not numeric",
        )?;
        let width = get_f64(
            self.dispatch_object(),
            "excel.range.width",
            "Range.Width was not numeric",
        )?;
        let height = get_f64(
            self.dispatch_object(),
            "excel.range.height",
            "Range.Height was not numeric",
        )?;
        let result = ShapeBounds {
            left,
            top,
            width,
            height,
        };
        shape_bounds(result)?;
        Ok(result)
    }
    /// Copies this Range as an Excel-native picture. Clear `CutCopyMode` after controlled use.
    pub fn copy_picture(&self, options: &CopyPictureOptions) -> Result<(), ExcelComError> {
        copy_picture(self.dispatch_object(), "excel.range.copypicture", options)
    }
    /// Returns cell-bound Sparkline groups associated with this Range.
    pub fn sparkline_groups(&self) -> Result<SparklineGroups, ExcelComError> {
        get_dispatch(
            self.dispatch_object(),
            "excel.range.sparklinegroups",
            SparklineGroups::from_dispatch,
        )
    }
}
impl Application {
    /// Returns Excel's current copy/cut state without reading the system clipboard.
    pub fn cut_copy_mode(&self) -> Result<CutCopyMode, ExcelComError> {
        Ok(CutCopyMode::from_raw(get_i32(
            self.dispatch_object(),
            "excel.application.cutcopymode",
            "Application.CutCopyMode was not an integer",
        )?))
    }
    /// Clears Excel's copy/cut state without extracting arbitrary clipboard contents.
    pub fn clear_cut_copy_mode(&self) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.application.cutcopymode",
            OwnedVariant::bool(false),
        )
    }
}
