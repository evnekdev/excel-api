# Shapes

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Shapes` |
| GUID | `{0002443a-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Shapes` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Item` | `excel.shape` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Implemented | Reviewed | Live Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | ShapeRange | declared | 197 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _Default | Shape | 1 | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Item | Shape | 1 | declared | 170 | Implemented | Reviewed | Live Tested | |
| AddCallout | Shape | 5 | declared | 1713 | Metadata Only | Reviewed | Not Tested | |
| AddConnector | Shape | 5 | declared | 1714 | Metadata Only | Reviewed | Not Tested | |
| AddCurve | Shape | 1 | declared | 1719 | Metadata Only | Reviewed | Not Tested | |
| AddLabel | Shape | 5 | declared | 1721 | Metadata Only | Reviewed | Not Tested | |
| AddLine | Shape | 4 | declared | 1722 | Implemented | Reviewed | Live Tested | |
| AddPicture | Shape | 7 | declared | 1723 | Implemented | Reviewed | Live Tested | |
| AddPolyline | Shape | 1 | declared | 1726 | Metadata Only | Reviewed | Not Tested | |
| AddShape | Shape | 5 | declared | 1727 | Implemented | Reviewed | Live Tested | |
| AddTextEffect | Shape | 8 | declared | 1728 | Metadata Only | Reviewed | Not Tested | |
| AddTextbox | Shape | 5 | declared | 1734 | Implemented | Reviewed | Live Tested | |
| BuildFreeform | FreeformBuilder | 3 | declared | 1735 | Metadata Only | Reviewed | Not Tested | |
| SelectAll | Unknown | 0 | declared | 1737 | Metadata Only | Reviewed | Not Tested | |
| AddFormControl | Shape | 5 | declared | 1738 | Metadata Only | Reviewed | Not Tested | |
| AddOLEObject | Shape | 11 | declared | 1739 | Metadata Only | Reviewed | Not Tested | |
| AddDiagram | Shape | 5 | declared | 2176 | Metadata Only | Reviewed | Not Tested | |
| AddCanvas | Shape | 4 | declared | 2177 | Metadata Only | Reviewed | Not Tested | |
| AddChart | Shape | 5 | declared | 2665 | Metadata Only | Reviewed | Not Tested | |
| AddSmartArt | Shape | 5 | declared | 2920 | Metadata Only | Reviewed | Not Tested | |
| AddChart2 | Shape | 7 | declared | 3088 | Metadata Only | Reviewed | Not Tested | |
| AddPicture2 | Shape | 8 | declared | 3159 | Metadata Only | Reviewed | Not Tested | |
| Add3DModel | Shape | 7 | declared | 3359 | Metadata Only | Reviewed | Not Tested | |
| QueryInterface | Unknown | 2 | inherited-iunknown | 1610612736 | Metadata Only | Reviewed | Not Tested | |
| AddRef | Unknown | 0 | inherited-iunknown | 1610612737 | Metadata Only | Reviewed | Not Tested | |
| Release | Unknown | 0 | inherited-iunknown | 1610612738 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfoCount | Unknown | 1 | inherited-idispatch | 1610678272 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfo | Unknown | 3 | inherited-idispatch | 1610678273 | Metadata Only | Reviewed | Not Tested | |
| GetIDsOfNames | Unknown | 5 | inherited-idispatch | 1610678274 | Metadata Only | Reviewed | Not Tested | |
| Invoke | Unknown | 8 | inherited-idispatch | 1610678275 | Metadata Only | Reviewed | Not Tested | |

## Events

| Event | Arguments | DISPID | Implementation | Docs | Tests |
|---|---:|---:|---|---|---|
| -- | -- | -- | Not started | Generated | Not tested |

## Unsupported or deferred behaviour

See the global unsupported index for unimplemented object-model areas.
<!-- END GENERATED MEMBERS -->
