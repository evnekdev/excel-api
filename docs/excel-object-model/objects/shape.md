# Shape

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Shape` |
| GUID | `{00024439-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Shape` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `BottomRightCell` | `excel.range` | Metadata Only |
| `TopLeftCell` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Chart | PROPERTYGET | Chart | declared | 7 | Metadata Only | Reviewed | Not Tested | |
| Rotation | PROPERTYGET/PROPERTYPUT | f32 | declared | 59 | Implemented | Reviewed | Live Tested | |
| Shadow | PROPERTYGET | ShadowFormat | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET | MsoShapeType | declared | 108 | Implemented | Reviewed | Live Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f32 | declared | 122 | Implemented | Reviewed | Live Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f32 | declared | 123 | Implemented | Reviewed | Live Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f32 | declared | 126 | Implemented | Reviewed | Live Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f32 | declared | 127 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Title | PROPERTYGET/PROPERTYPUT | String | declared | 199 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | bool | declared | 269 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | MsoTriState | declared | 558 | Implemented | Reviewed | Live Tested | |
| ID | PROPERTYGET | i32 | declared | 570 | Metadata Only | Reviewed | Not Tested | |
| OnAction | PROPERTYGET/PROPERTYPUT | String | declared | 596 | Metadata Only | Reviewed | Not Tested | |
| BottomRightCell | PROPERTYGET | Range | declared | 615 | Metadata Only | Reviewed | Not Tested | |
| Placement | PROPERTYGET/PROPERTYPUT | XlPlacement | declared | 617 | Implemented | Reviewed | Live Tested | |
| TopLeftCell | PROPERTYGET | Range | declared | 620 | Metadata Only | Reviewed | Not Tested | |
| Vertices | PROPERTYGET | AutomationValue | declared | 621 | Metadata Only | Reviewed | Not Tested | |
| Line | PROPERTYGET | LineFormat | declared | 817 | Implemented | Reviewed | Live Tested | |
| PictureFormat | PROPERTYGET | PictureFormat | declared | 1631 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | FillFormat | declared | 1663 | Implemented | Reviewed | Live Tested | |
| Adjustments | PROPERTYGET | Adjustments | declared | 1691 | Metadata Only | Reviewed | Not Tested | |
| TextFrame | PROPERTYGET | TextFrame | declared | 1692 | Metadata Only | Reviewed | Not Tested | |
| AutoShapeType | PROPERTYGET/PROPERTYPUT | MsoAutoShapeType | declared | 1693 | Metadata Only | Reviewed | Not Tested | |
| Callout | PROPERTYGET | CalloutFormat | declared | 1694 | Metadata Only | Reviewed | Not Tested | |
| ConnectionSiteCount | PROPERTYGET | i32 | declared | 1695 | Metadata Only | Reviewed | Not Tested | |
| Connector | PROPERTYGET | MsoTriState | declared | 1696 | Metadata Only | Reviewed | Not Tested | |
| ConnectorFormat | PROPERTYGET | ConnectorFormat | declared | 1697 | Metadata Only | Reviewed | Not Tested | |
| GroupItems | PROPERTYGET | GroupShapes | declared | 1698 | Metadata Only | Reviewed | Not Tested | |
| HorizontalFlip | PROPERTYGET | MsoTriState | declared | 1699 | Metadata Only | Reviewed | Not Tested | |
| LockAspectRatio | PROPERTYGET/PROPERTYPUT | MsoTriState | declared | 1700 | Implemented | Reviewed | Live Tested | |
| Nodes | PROPERTYGET | ShapeNodes | declared | 1701 | Metadata Only | Reviewed | Not Tested | |
| TextEffect | PROPERTYGET | TextEffectFormat | declared | 1702 | Metadata Only | Reviewed | Not Tested | |
| ThreeD | PROPERTYGET | ThreeDFormat | declared | 1703 | Metadata Only | Reviewed | Not Tested | |
| VerticalFlip | PROPERTYGET | MsoTriState | declared | 1704 | Metadata Only | Reviewed | Not Tested | |
| ZOrderPosition | PROPERTYGET | i32 | declared | 1705 | Metadata Only | Reviewed | Not Tested | |
| Hyperlink | PROPERTYGET | Hyperlink | declared | 1706 | Metadata Only | Reviewed | Not Tested | |
| BlackWhiteMode | PROPERTYGET/PROPERTYPUT | MsoBlackWhiteMode | declared | 1707 | Metadata Only | Reviewed | Not Tested | |
| DrawingObject | PROPERTYGET | Object | declared | 1708 | Metadata Only | Reviewed | Not Tested | |
| ControlFormat | PROPERTYGET | ControlFormat | declared | 1709 | Metadata Only | Reviewed | Not Tested | |
| LinkFormat | PROPERTYGET | LinkFormat | declared | 1710 | Metadata Only | Reviewed | Not Tested | |
| OLEFormat | PROPERTYGET | OLEFormat | declared | 1711 | Metadata Only | Reviewed | Not Tested | |
| FormControlType | PROPERTYGET | XlFormControl | declared | 1712 | Metadata Only | Reviewed | Not Tested | |
| AlternativeText | PROPERTYGET/PROPERTYPUT | String | declared | 1891 | Metadata Only | Reviewed | Not Tested | |
| Script | PROPERTYGET | Script | declared | 1892 | Metadata Only | Reviewed | Not Tested | |
| DiagramNode | PROPERTYGET | DiagramNode | declared | 2165 | Metadata Only | Reviewed | Not Tested | |
| HasDiagramNode | PROPERTYGET | MsoTriState | declared | 2166 | Metadata Only | Reviewed | Not Tested | |
| Diagram | PROPERTYGET | Diagram | declared | 2167 | Metadata Only | Reviewed | Not Tested | |
| HasDiagram | PROPERTYGET | MsoTriState | declared | 2168 | Metadata Only | Reviewed | Not Tested | |
| Child | PROPERTYGET | MsoTriState | declared | 2169 | Metadata Only | Reviewed | Not Tested | |
| ParentGroup | PROPERTYGET | Shape | declared | 2170 | Metadata Only | Reviewed | Not Tested | |
| CanvasItems | PROPERTYGET | CanvasShapes | declared | 2171 | Metadata Only | Reviewed | Not Tested | |
| HasChart | PROPERTYGET | MsoTriState | declared | 2658 | Metadata Only | Reviewed | Not Tested | |
| TextFrame2 | PROPERTYGET | TextFrame2 | declared | 2659 | Implemented | Reviewed | Live Tested | |
| ShapeStyle | PROPERTYGET/PROPERTYPUT | MsoShapeStyleIndex | declared | 2660 | Metadata Only | Reviewed | Not Tested | |
| BackgroundStyle | PROPERTYGET/PROPERTYPUT | MsoBackgroundStyleIndex | declared | 2661 | Metadata Only | Reviewed | Not Tested | |
| SoftEdge | PROPERTYGET | SoftEdgeFormat | declared | 2662 | Metadata Only | Reviewed | Not Tested | |
| Glow | PROPERTYGET | GlowFormat | declared | 2663 | Metadata Only | Reviewed | Not Tested | |
| Reflection | PROPERTYGET | ReflectionFormat | declared | 2664 | Metadata Only | Reviewed | Not Tested | |
| HasSmartArt | PROPERTYGET | MsoTriState | declared | 2918 | Metadata Only | Reviewed | Not Tested | |
| SmartArt | PROPERTYGET | SmartArt | declared | 2919 | Metadata Only | Reviewed | Not Tested | |
| GraphicStyle | PROPERTYGET/PROPERTYPUT | MsoGraphicStyleIndex | declared | 3272 | Metadata Only | Reviewed | Not Tested | |
| Model3D | PROPERTYGET | Model3DFormat | declared | 3357 | Metadata Only | Reviewed | Not Tested | |
| Decorative | PROPERTYGET/PROPERTYPUT | MsoTriState | declared | 3358 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| CopyPicture | Unknown | 2 | declared | 213 | Metadata Only | Reviewed | Not Tested | |
| Select | Unknown | 1 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Ungroup | ShapeRange | 0 | declared | 244 | Metadata Only | Reviewed | Not Tested | |
| Copy | Unknown | 0 | declared | 551 | Implemented | Reviewed | Live Tested | |
| Cut | Unknown | 0 | declared | 565 | Implemented | Reviewed | Live Tested | |
| ZOrder | Unknown | 1 | declared | 622 | Implemented | Reviewed | Live Tested | |
| Duplicate | Shape | 0 | declared | 1039 | Metadata Only | Reviewed | Not Tested | |
| Apply | Unknown | 0 | declared | 1675 | Metadata Only | Reviewed | Not Tested | |
| Flip | Unknown | 1 | declared | 1676 | Metadata Only | Reviewed | Not Tested | |
| IncrementLeft | Unknown | 1 | declared | 1678 | Metadata Only | Reviewed | Not Tested | |
| IncrementRotation | Unknown | 1 | declared | 1680 | Metadata Only | Reviewed | Not Tested | |
| IncrementTop | Unknown | 1 | declared | 1681 | Metadata Only | Reviewed | Not Tested | |
| PickUp | Unknown | 0 | declared | 1682 | Metadata Only | Reviewed | Not Tested | |
| RerouteConnections | Unknown | 0 | declared | 1683 | Metadata Only | Reviewed | Not Tested | |
| ScaleHeight | Unknown | 3 | declared | 1684 | Metadata Only | Reviewed | Not Tested | |
| ScaleWidth | Unknown | 3 | declared | 1688 | Metadata Only | Reviewed | Not Tested | |
| SetShapesDefaultProperties | Unknown | 0 | declared | 1689 | Metadata Only | Reviewed | Not Tested | |
| CanvasCropLeft | Unknown | 1 | declared | 2172 | Metadata Only | Reviewed | Not Tested | |
| CanvasCropTop | Unknown | 1 | declared | 2173 | Metadata Only | Reviewed | Not Tested | |
| CanvasCropRight | Unknown | 1 | declared | 2174 | Metadata Only | Reviewed | Not Tested | |
| CanvasCropBottom | Unknown | 1 | declared | 2175 | Metadata Only | Reviewed | Not Tested | |
| PlacePictureInCell | Unknown | 0 | declared | 3409 | Metadata Only | Reviewed | Not Tested | |
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
