//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// A cell-bound Excel SparklineGroups collection.
pub struct SparklineGroups {
    inner: DispatchObject,
}
impl Debug for SparklineGroups {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SparklineGroups").field(&self.inner).finish()
    }
}
impl SparklineGroups {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SparklineGroups", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SPARKLINE_GROUPS)
    }
    pub fn item(&self, index: usize) -> Result<SparklineGroup, ExcelComError> {
        collection_item(
            &self.inner,
            SPARKLINE_GROUPS,
            index,
            SparklineGroup::from_dispatch,
        )
    }
    pub fn iter(&self) -> Result<SparklineGroupsIter, ExcelComError> {
        Ok(SparklineGroupsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SPARKLINE_GROUPS.new_enum),
                SPARKLINE_GROUPS.name,
            )?,
            index: 0,
            done: false,
        })
    }
    /// Adds a cell-bound group using Excel source and location addresses.
    pub fn add(
        &self,
        sparkline_type: SparklineType,
        source_data: &Range,
        location: &Range,
    ) -> Result<SparklineGroup, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(sparkline_type.raw()));
        args.push_result(text_bstr(&source_data.address_a1()?))?;
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sparklinegroups.add"), false),
            args.into_inner(),
            false,
        )?;
        let group = SparklineGroup::from_dispatch(value.take_dispatch()?);
        group.set_location(location)?;
        Ok(group)
    }
}
pub struct SparklineGroupsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for SparklineGroupsIter {
    type Item = Result<SparklineGroup, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "SparklineGroups", index)
                        .map(SparklineGroup::from_dispatch),
                )
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
impl FusedIterator for SparklineGroupsIter {}
/// A group of cell-bound Excel sparklines.
pub struct SparklineGroup {
    inner: DispatchObject,
}
impl Debug for SparklineGroup {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SparklineGroup").field(&self.inner).finish()
    }
}
impl SparklineGroup {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SparklineGroup", value),
        }
    }
    pub fn sparkline_type(&self) -> Result<SparklineType, ExcelComError> {
        Ok(SparklineType::from_raw(get_i32(
            &self.inner,
            "excel.sparklinegroup.type",
            "SparklineGroup.Type was not an integer",
        )?))
    }
    pub fn set_sparkline_type(&self, value: SparklineType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.type",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn source_data(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.sparklinegroup.sourcedata")
    }
    pub fn set_source_data(&self, source: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.sourcedata",
            text_bstr(&source.address_a1()?)?,
        )
    }
    fn set_location(&self, location: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.location",
            OwnedVariant::dispatch_borrowed(&location.dispatch_object().dispatch),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.sparklinegroup.delete", vec![])
    }
}
