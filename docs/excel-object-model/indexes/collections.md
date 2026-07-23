# Collection inventory

Collections are detected structurally from Count and Item. Iterator status is independent from the broader wrapper status.

## Priority collection status

| Collection | Iterator status |
|---|---|
| Workbooks | implemented |
| Worksheets | implemented |
| Areas | implemented |
| Names | implemented |
| Borders | implemented |
| Charts | implemented |
| Shapes | implemented |
| ListObjects | implemented |
| FormatConditions | implemented |
| ColorScaleCriteria | implemented |
| IconCriteria | implemented |
| Styles | implemented |
| Comments | implemented |
| CommentsThreaded | implemented |
| Hyperlinks | implemented |

## All structurally identified collections

| Collection | Element | Count member | Item member | Enumerator | Index kinds | Iterator |
|---|---|---|---|---|---|---|
| Actions | Unknown | excel.actions.count | excel.actions.item | excel.actions.newenum | variant-key | metadata-only |
| AddIns | Unknown | excel.addins.count | excel.addins.item | excel.addins.newenum | variant-key | metadata-only |
| AddIns2 | Unknown | excel.addins2.count | excel.addins2.item | excel.addins2.newenum | variant-key | metadata-only |
| Adjustments | Unknown | excel.adjustments.count | excel.adjustments.item | -- | variant-key | metadata-only |
| AllowEditRanges | Unknown | excel.alloweditranges.count | excel.alloweditranges.item | excel.alloweditranges.newenum | variant-key | metadata-only |
| Arcs | Unknown | excel.arcs.count | excel.arcs.item | excel.arcs.newenum | variant-key | metadata-only |
| Areas | Range | excel.areas.count | excel.areas.item | excel.areas.newenum | one-based-integer | implemented |
| Axes | Unknown | excel.axes.count | excel.axes.item | excel.axes.newenum | variant-key | metadata-only |
| Borders | Border | excel.borders.count | excel.borders.item | excel.borders.newenum | enum-key | implemented |
| Buttons | Unknown | excel.buttons.count | excel.buttons.item | excel.buttons.newenum | variant-key | metadata-only |
| CalculatedFields | Unknown | excel.calculatedfields.count | excel.calculatedfields.item | excel.calculatedfields.newenum | variant-key | metadata-only |
| CalculatedItems | Unknown | excel.calculateditems.count | excel.calculateditems.item | excel.calculateditems.newenum | variant-key | metadata-only |
| CalculatedMembers | Unknown | excel.calculatedmembers.count | excel.calculatedmembers.item | excel.calculatedmembers.newenum | variant-key | metadata-only |
| CategoryCollection | Unknown | excel.categorycollection.count | excel.categorycollection.item | -- | variant-key | metadata-only |
| ChartGroups | Unknown | excel.chartgroups.count | excel.chartgroups.item | excel.chartgroups.newenum | variant-key | metadata-only |
| ChartObjects | ChartObject | excel.chartobjects.count | excel.chartobjects.item | excel.chartobjects.newenum | one-based-integer, string-key | implemented |
| Charts | Chart | excel.charts.count | excel.charts.item | excel.charts.newenum | one-based-integer, string-key | implemented |
| CheckBoxes | Unknown | excel.checkboxes.count | excel.checkboxes.item | excel.checkboxes.newenum | variant-key | metadata-only |
| ColorScaleCriteria | ColorScaleCriterion | excel.colorscalecriteria.count | excel.colorscalecriteria.item | excel.colorscalecriteria.newenum | one-based-integer | implemented |
| ColorStops | Unknown | excel.colorstops.count | excel.colorstops.item | excel.colorstops.newenum | variant-key | metadata-only |
| Comments | Comment | excel.comments.count | excel.comments.item | excel.comments.newenum | one-based-integer | implemented |
| CommentsThreaded | CommentThreaded | excel.commentsthreaded.count | excel.commentsthreaded.item | excel.commentsthreaded.newenum | one-based-integer | implemented |
| Connections | Unknown | excel.connections.count | excel.connections.item | excel.connections.newenum | variant-key | metadata-only |
| CubeFields | Unknown | excel.cubefields.count | excel.cubefields.item | excel.cubefields.newenum | variant-key | metadata-only |
| CustomProperties | Unknown | excel.customproperties.count | excel.customproperties.item | excel.customproperties.newenum | variant-key | metadata-only |
| CustomViews | Unknown | excel.customviews.count | excel.customviews.item | excel.customviews.newenum | variant-key | metadata-only |
| DataLabels | DataLabel | excel.datalabels.count | excel.datalabels.item | excel.datalabels.newenum | one-based-integer | metadata-only |
| DiagramNodeChildren | Unknown | excel.diagramnodechildren.count | excel.diagramnodechildren.item | excel.diagramnodechildren.newenum | variant-key | metadata-only |
| DiagramNodes | Unknown | excel.diagramnodes.count | excel.diagramnodes.item | excel.diagramnodes.newenum | variant-key | metadata-only |
| Dialogs | Unknown | excel.dialogs.count | excel.dialogs.item | excel.dialogs.newenum | variant-key | metadata-only |
| DialogSheets | Unknown | excel.dialogsheets.count | excel.dialogsheets.item | excel.dialogsheets.newenum | variant-key | metadata-only |
| DrawingObjects | Unknown | excel.drawingobjects.count | excel.drawingobjects.item | excel.drawingobjects.newenum | variant-key | metadata-only |
| Drawings | Unknown | excel.drawings.count | excel.drawings.item | excel.drawings.newenum | variant-key | metadata-only |
| DropDowns | Unknown | excel.dropdowns.count | excel.dropdowns.item | excel.dropdowns.newenum | variant-key | metadata-only |
| EditBoxes | Unknown | excel.editboxes.count | excel.editboxes.item | excel.editboxes.newenum | variant-key | metadata-only |
| FileExportConverters | Unknown | excel.fileexportconverters.count | excel.fileexportconverters.item | excel.fileexportconverters.newenum | variant-key | metadata-only |
| Filters | Filter | excel.filters.count | excel.filters.item | excel.filters.newenum | one-based-integer | implemented |
| FormatConditions | ConditionalFormat | excel.formatconditions.count | excel.formatconditions.item | excel.formatconditions.newenum | one-based-integer | implemented |
| FullSeriesCollection | Unknown | excel.fullseriescollection.count | excel.fullseriescollection.item | excel.fullseriescollection.newenum | variant-key | metadata-only |
| GroupBoxes | Unknown | excel.groupboxes.count | excel.groupboxes.item | excel.groupboxes.newenum | variant-key | metadata-only |
| GroupObjects | Unknown | excel.groupobjects.count | excel.groupobjects.item | excel.groupobjects.newenum | variant-key | metadata-only |
| GroupShapes | Unknown | excel.groupshapes.count | excel.groupshapes.item | excel.groupshapes.newenum | variant-key | metadata-only |
| HPageBreaks | Unknown | excel.hpagebreaks.count | excel.hpagebreaks.item | excel.hpagebreaks.newenum | variant-key | metadata-only |
| Hyperlinks | Hyperlink | excel.hyperlinks.count | excel.hyperlinks.item | excel.hyperlinks.newenum | one-based-integer | implemented |
| IActions | Unknown | excel.iactions.count | excel.iactions.item | excel.iactions.newenum | variant-key | metadata-only |
| IAddIns | Unknown | excel.iaddins.count | excel.iaddins.item | excel.iaddins.newenum | variant-key | metadata-only |
| IAddIns2 | Unknown | excel.iaddins2.count | excel.iaddins2.item | excel.iaddins2.newenum | variant-key | metadata-only |
| IAllowEditRanges | Unknown | excel.ialloweditranges.count | excel.ialloweditranges.item | excel.ialloweditranges.newenum | variant-key | metadata-only |
| IArcs | Unknown | excel.iarcs.count | excel.iarcs.item | excel.iarcs.newenum | variant-key | metadata-only |
| IAreas | Unknown | excel.iareas.count | excel.iareas.item | excel.iareas.newenum | variant-key | metadata-only |
| IAxes | Unknown | excel.iaxes.count | excel.iaxes.item | excel.iaxes.newenum | variant-key | metadata-only |
| IBorders | Unknown | excel.iborders.count | excel.iborders.item | excel.iborders.newenum | variant-key | metadata-only |
| IButtons | Unknown | excel.ibuttons.count | excel.ibuttons.item | excel.ibuttons.newenum | variant-key | metadata-only |
| ICalculatedFields | Unknown | excel.icalculatedfields.count | excel.icalculatedfields.item | excel.icalculatedfields.newenum | variant-key | metadata-only |
| ICalculatedItems | Unknown | excel.icalculateditems.count | excel.icalculateditems.item | excel.icalculateditems.newenum | variant-key | metadata-only |
| ICalculatedMembers | Unknown | excel.icalculatedmembers.count | excel.icalculatedmembers.item | excel.icalculatedmembers.newenum | variant-key | metadata-only |
| ICategoryCollection | Unknown | excel.icategorycollection.count | excel.icategorycollection.item | -- | variant-key | metadata-only |
| IChartGroups | Unknown | excel.ichartgroups.count | excel.ichartgroups.item | excel.ichartgroups.newenum | variant-key | metadata-only |
| IChartObjects | Unknown | excel.ichartobjects.count | excel.ichartobjects.item | excel.ichartobjects.newenum | variant-key | metadata-only |
| ICharts | Unknown | excel.icharts.count | excel.icharts.item | excel.icharts.newenum | variant-key | metadata-only |
| ICheckBoxes | Unknown | excel.icheckboxes.count | excel.icheckboxes.item | excel.icheckboxes.newenum | variant-key | metadata-only |
| IColorScaleCriteria | Unknown | excel.icolorscalecriteria.count | excel.icolorscalecriteria.item | excel.icolorscalecriteria.newenum | variant-key | metadata-only |
| IColorStops | Unknown | excel.icolorstops.count | excel.icolorstops.item | excel.icolorstops.newenum | variant-key | metadata-only |
| IComments | Unknown | excel.icomments.count | excel.icomments.item | excel.icomments.newenum | variant-key | metadata-only |
| ICommentsThreaded | Unknown | excel.icommentsthreaded.count | excel.icommentsthreaded.item | excel.icommentsthreaded.newenum | variant-key | metadata-only |
| IconCriteria | IconCriterion | excel.iconcriteria.count | excel.iconcriteria.item | excel.iconcriteria.newenum | one-based-integer | implemented |
| IConnections | Unknown | excel.iconnections.count | excel.iconnections.item | excel.iconnections.newenum | variant-key | metadata-only |
| IconSet | Unknown | excel.iconset.count | excel.iconset.item | excel.iconset.newenum | variant-key | metadata-only |
| IconSets | Unknown | excel.iconsets.count | excel.iconsets.item | excel.iconsets.newenum | variant-key | metadata-only |
| ICustomProperties | Unknown | excel.icustomproperties.count | excel.icustomproperties.item | excel.icustomproperties.newenum | variant-key | metadata-only |
| ICustomViews | Unknown | excel.icustomviews.count | excel.icustomviews.item | excel.icustomviews.newenum | variant-key | metadata-only |
| IDataLabels | Unknown | excel.idatalabels.count | excel.idatalabels.item | excel.idatalabels.newenum | variant-key | metadata-only |
| IDialogs | Unknown | excel.idialogs.count | excel.idialogs.item | excel.idialogs.newenum | variant-key | metadata-only |
| IDialogSheets | Unknown | excel.idialogsheets.count | excel.idialogsheets.item | excel.idialogsheets.newenum | variant-key | metadata-only |
| IDrawingObjects | Unknown | excel.idrawingobjects.count | excel.idrawingobjects.item | excel.idrawingobjects.newenum | variant-key | metadata-only |
| IDrawings | Unknown | excel.idrawings.count | excel.idrawings.item | excel.idrawings.newenum | variant-key | metadata-only |
| IDropDowns | Unknown | excel.idropdowns.count | excel.idropdowns.item | excel.idropdowns.newenum | variant-key | metadata-only |
| IEditBoxes | Unknown | excel.ieditboxes.count | excel.ieditboxes.item | excel.ieditboxes.newenum | variant-key | metadata-only |
| IFileExportConverters | Unknown | excel.ifileexportconverters.count | excel.ifileexportconverters.item | excel.ifileexportconverters.newenum | variant-key | metadata-only |
| IFilters | Unknown | excel.ifilters.count | excel.ifilters.item | excel.ifilters.newenum | variant-key | metadata-only |
| IFormatConditions | Unknown | excel.iformatconditions.count | excel.iformatconditions.item | excel.iformatconditions.newenum | variant-key | metadata-only |
| IFullSeriesCollection | Unknown | excel.ifullseriescollection.count | excel.ifullseriescollection.item | excel.ifullseriescollection.newenum | variant-key | metadata-only |
| IGroupBoxes | Unknown | excel.igroupboxes.count | excel.igroupboxes.item | excel.igroupboxes.newenum | variant-key | metadata-only |
| IGroupObjects | Unknown | excel.igroupobjects.count | excel.igroupobjects.item | excel.igroupobjects.newenum | variant-key | metadata-only |
| IGroupShapes | Unknown | excel.igroupshapes.count | excel.igroupshapes.item | excel.igroupshapes.newenum | variant-key | metadata-only |
| IHPageBreaks | Unknown | excel.ihpagebreaks.count | excel.ihpagebreaks.item | excel.ihpagebreaks.newenum | variant-key | metadata-only |
| IHyperlinks | Unknown | excel.ihyperlinks.count | excel.ihyperlinks.item | excel.ihyperlinks.newenum | variant-key | metadata-only |
| IIconCriteria | Unknown | excel.iiconcriteria.count | excel.iiconcriteria.item | excel.iiconcriteria.newenum | variant-key | metadata-only |
| IIconSet | Unknown | excel.iiconset.count | excel.iiconset.item | excel.iiconset.newenum | variant-key | metadata-only |
| IIconSets | Unknown | excel.iiconsets.count | excel.iiconsets.item | excel.iiconsets.newenum | variant-key | metadata-only |
| ILabels | Unknown | excel.ilabels.count | excel.ilabels.item | excel.ilabels.newenum | variant-key | metadata-only |
| ILegendEntries | Unknown | excel.ilegendentries.count | excel.ilegendentries.item | excel.ilegendentries.newenum | variant-key | metadata-only |
| ILines | Unknown | excel.ilines.count | excel.ilines.item | excel.ilines.newenum | variant-key | metadata-only |
| IListBoxes | Unknown | excel.ilistboxes.count | excel.ilistboxes.item | excel.ilistboxes.newenum | variant-key | metadata-only |
| IListColumns | Unknown | excel.ilistcolumns.count | excel.ilistcolumns.item | excel.ilistcolumns.newenum | variant-key | metadata-only |
| IListObjects | Unknown | excel.ilistobjects.count | excel.ilistobjects.item | excel.ilistobjects.newenum | variant-key | metadata-only |
| IListRows | Unknown | excel.ilistrows.count | excel.ilistrows.item | excel.ilistrows.newenum | variant-key | metadata-only |
| IMenuBars | Unknown | excel.imenubars.count | excel.imenubars.item | excel.imenubars.newenum | variant-key | metadata-only |
| IMenuItems | Unknown | excel.imenuitems.count | excel.imenuitems.item | excel.imenuitems.newenum | variant-key | metadata-only |
| IMenus | Unknown | excel.imenus.count | excel.imenus.item | excel.imenus.newenum | variant-key | metadata-only |
| IModelColumnChanges | Unknown | excel.imodelcolumnchanges.count | excel.imodelcolumnchanges.item | excel.imodelcolumnchanges.newenum | variant-key | metadata-only |
| IModelColumnNames | Unknown | excel.imodelcolumnnames.count | excel.imodelcolumnnames.item | excel.imodelcolumnnames.newenum | variant-key | metadata-only |
| IModelMeasureNames | Unknown | excel.imodelmeasurenames.count | excel.imodelmeasurenames.item | excel.imodelmeasurenames.newenum | variant-key | metadata-only |
| IModelMeasures | Unknown | excel.imodelmeasures.count | excel.imodelmeasures.item | excel.imodelmeasures.newenum | variant-key | metadata-only |
| IModelRelationships | Unknown | excel.imodelrelationships.count | excel.imodelrelationships.item | excel.imodelrelationships.newenum | variant-key | metadata-only |
| IModelTableColumns | Unknown | excel.imodeltablecolumns.count | excel.imodeltablecolumns.item | excel.imodeltablecolumns.newenum | variant-key | metadata-only |
| IModelTableNameChanges | Unknown | excel.imodeltablenamechanges.count | excel.imodeltablenamechanges.item | excel.imodeltablenamechanges.newenum | variant-key | metadata-only |
| IModelTableNames | Unknown | excel.imodeltablenames.count | excel.imodeltablenames.item | -- | variant-key | metadata-only |
| IModelTables | Unknown | excel.imodeltables.count | excel.imodeltables.item | excel.imodeltables.newenum | variant-key | metadata-only |
| IModules | Unknown | excel.imodules.count | excel.imodules.item | excel.imodules.newenum | variant-key | metadata-only |
| INames | Unknown | excel.inames.count | excel.inames.item | excel.inames.newenum | variant-key | metadata-only |
| IODBCErrors | Unknown | excel.iodbcerrors.count | excel.iodbcerrors.item | excel.iodbcerrors.newenum | variant-key | metadata-only |
| IOLEDBErrors | Unknown | excel.ioledberrors.count | excel.ioledberrors.item | excel.ioledberrors.newenum | variant-key | metadata-only |
| IOLEObjects | Unknown | excel.ioleobjects.count | excel.ioleobjects.item | excel.ioleobjects.newenum | variant-key | metadata-only |
| IOptionButtons | Unknown | excel.ioptionbuttons.count | excel.ioptionbuttons.item | excel.ioptionbuttons.newenum | variant-key | metadata-only |
| IOvals | Unknown | excel.iovals.count | excel.iovals.item | excel.iovals.newenum | variant-key | metadata-only |
| IPages | Unknown | excel.ipages.count | excel.ipages.item | excel.ipages.newenum | variant-key | metadata-only |
| IPanes | Unknown | excel.ipanes.count | excel.ipanes.item | -- | variant-key | metadata-only |
| IParameters | Unknown | excel.iparameters.count | excel.iparameters.item | excel.iparameters.newenum | variant-key | metadata-only |
| IPhonetics | Unknown | excel.iphonetics.count | excel.iphonetics.item | excel.iphonetics.newenum | variant-key | metadata-only |
| IPictures | Unknown | excel.ipictures.count | excel.ipictures.item | excel.ipictures.newenum | variant-key | metadata-only |
| IPivotCaches | Unknown | excel.ipivotcaches.count | excel.ipivotcaches.item | excel.ipivotcaches.newenum | variant-key | metadata-only |
| IPivotFields | Unknown | excel.ipivotfields.count | excel.ipivotfields.item | excel.ipivotfields.newenum | variant-key | metadata-only |
| IPivotFilters | Unknown | excel.ipivotfilters.count | excel.ipivotfilters.item | excel.ipivotfilters.newenum | variant-key | metadata-only |
| IPivotFormulas | Unknown | excel.ipivotformulas.count | excel.ipivotformulas.item | excel.ipivotformulas.newenum | variant-key | metadata-only |
| IPivotItemList | Unknown | excel.ipivotitemlist.count | excel.ipivotitemlist.item | excel.ipivotitemlist.newenum | variant-key | metadata-only |
| IPivotItems | Unknown | excel.ipivotitems.count | excel.ipivotitems.item | excel.ipivotitems.newenum | variant-key | metadata-only |
| IPivotLineCells | Unknown | excel.ipivotlinecells.count | excel.ipivotlinecells.item | excel.ipivotlinecells.newenum | variant-key | metadata-only |
| IPivotLines | Unknown | excel.ipivotlines.count | excel.ipivotlines.item | excel.ipivotlines.newenum | variant-key | metadata-only |
| IPivotTableChangeList | Unknown | excel.ipivottablechangelist.count | excel.ipivottablechangelist.item | excel.ipivottablechangelist.newenum | variant-key | metadata-only |
| IPivotTables | Unknown | excel.ipivottables.count | excel.ipivottables.item | excel.ipivottables.newenum | variant-key | metadata-only |
| IPoints | Unknown | excel.ipoints.count | excel.ipoints.item | excel.ipoints.newenum | variant-key | metadata-only |
| IProtectedViewWindows | Unknown | excel.iprotectedviewwindows.count | excel.iprotectedviewwindows.item | excel.iprotectedviewwindows.newenum | variant-key | metadata-only |
| IPublishedDocs | Unknown | excel.ipublisheddocs.count | excel.ipublisheddocs.item | excel.ipublisheddocs.newenum | variant-key | metadata-only |
| IPublishObjects | Unknown | excel.ipublishobjects.count | excel.ipublishobjects.item | excel.ipublishobjects.newenum | variant-key | metadata-only |
| IQueries | Unknown | excel.iqueries.count | excel.iqueries.item | excel.iqueries.newenum | variant-key | metadata-only |
| IQueryTables | Unknown | excel.iquerytables.count | excel.iquerytables.item | excel.iquerytables.newenum | variant-key | metadata-only |
| IRange | Unknown | excel.irange.count | excel.irange.item | excel.irange.newenum | variant-key | metadata-only |
| IRanges | Unknown | excel.iranges.count | excel.iranges.item | excel.iranges.newenum | variant-key | metadata-only |
| IRecentFiles | Unknown | excel.irecentfiles.count | excel.irecentfiles.item | excel.irecentfiles.newenum | variant-key | metadata-only |
| IRectangles | Unknown | excel.irectangles.count | excel.irectangles.item | excel.irectangles.newenum | variant-key | metadata-only |
| IScenarios | Unknown | excel.iscenarios.count | excel.iscenarios.item | excel.iscenarios.newenum | variant-key | metadata-only |
| IScrollBars | Unknown | excel.iscrollbars.count | excel.iscrollbars.item | excel.iscrollbars.newenum | variant-key | metadata-only |
| ISeriesCollection | Unknown | excel.iseriescollection.count | excel.iseriescollection.item | excel.iseriescollection.newenum | variant-key | metadata-only |
| IServerViewableItems | Unknown | excel.iserverviewableitems.count | excel.iserverviewableitems.item | excel.iserverviewableitems.newenum | variant-key | metadata-only |
| IShapeRange | Unknown | excel.ishaperange.count | excel.ishaperange.item | excel.ishaperange.newenum | variant-key | metadata-only |
| IShapes | Unknown | excel.ishapes.count | excel.ishapes.item | excel.ishapes.newenum | variant-key | metadata-only |
| ISheetViews | Unknown | excel.isheetviews.count | excel.isheetviews.item | excel.isheetviews.newenum | variant-key | metadata-only |
| ISlicerCacheLevels | Unknown | excel.islicercachelevels.count | excel.islicercachelevels.item | excel.islicercachelevels.newenum | variant-key | metadata-only |
| ISlicerCaches | Unknown | excel.islicercaches.count | excel.islicercaches.item | excel.islicercaches.newenum | variant-key | metadata-only |
| ISlicerItems | Unknown | excel.isliceritems.count | excel.isliceritems.item | excel.isliceritems.newenum | variant-key | metadata-only |
| ISlicerPivotTables | Unknown | excel.islicerpivottables.count | excel.islicerpivottables.item | excel.islicerpivottables.newenum | variant-key | metadata-only |
| ISlicers | Unknown | excel.islicers.count | excel.islicers.item | excel.islicers.newenum | variant-key | metadata-only |
| ISmartTagActions | Unknown | excel.ismarttagactions.count | excel.ismarttagactions.item | excel.ismarttagactions.newenum | variant-key | metadata-only |
| ISmartTagRecognizers | Unknown | excel.ismarttagrecognizers.count | excel.ismarttagrecognizers.item | excel.ismarttagrecognizers.newenum | variant-key | metadata-only |
| ISortFields | Unknown | excel.isortfields.count | excel.isortfields.item | excel.isortfields.newenum | variant-key | metadata-only |
| ISparklineGroup | Unknown | excel.isparklinegroup.count | excel.isparklinegroup.item | excel.isparklinegroup.newenum | variant-key | metadata-only |
| ISparklineGroups | Unknown | excel.isparklinegroups.count | excel.isparklinegroups.item | excel.isparklinegroups.newenum | variant-key | metadata-only |
| ISpinners | Unknown | excel.ispinners.count | excel.ispinners.item | excel.ispinners.newenum | variant-key | metadata-only |
| IStyles | Unknown | excel.istyles.count | excel.istyles.item | excel.istyles.newenum | variant-key | metadata-only |
| ITableStyleElements | Unknown | excel.itablestyleelements.count | excel.itablestyleelements.item | excel.itablestyleelements.newenum | variant-key | metadata-only |
| ITableStyles | Unknown | excel.itablestyles.count | excel.itablestyles.item | excel.itablestyles.newenum | variant-key | metadata-only |
| ITextBoxes | Unknown | excel.itextboxes.count | excel.itextboxes.item | excel.itextboxes.newenum | variant-key | metadata-only |
| IToolbarButtons | Unknown | excel.itoolbarbuttons.count | excel.itoolbarbuttons.item | excel.itoolbarbuttons.newenum | variant-key | metadata-only |
| IToolbars | Unknown | excel.itoolbars.count | excel.itoolbars.item | excel.itoolbars.newenum | variant-key | metadata-only |
| ITrendlines | Unknown | excel.itrendlines.count | excel.itrendlines.item | excel.itrendlines.newenum | variant-key | metadata-only |
| IUsedObjects | Unknown | excel.iusedobjects.count | excel.iusedobjects.item | excel.iusedobjects.newenum | variant-key | metadata-only |
| IUserAccessList | Unknown | excel.iuseraccesslist.count | excel.iuseraccesslist.item | excel.iuseraccesslist.newenum | variant-key | metadata-only |
| IVPageBreaks | Unknown | excel.ivpagebreaks.count | excel.ivpagebreaks.item | excel.ivpagebreaks.newenum | variant-key | metadata-only |
| IWatches | Unknown | excel.iwatches.count | excel.iwatches.item | excel.iwatches.newenum | variant-key | metadata-only |
| IWindows | Unknown | excel.iwindows.count | excel.iwindows.item | excel.iwindows.newenum | variant-key | metadata-only |
| IWorksheets | Unknown | excel.iworksheets.count | excel.iworksheets.item | excel.iworksheets.newenum | variant-key | metadata-only |
| IXmlMaps | Unknown | excel.ixmlmaps.count | excel.ixmlmaps.item | excel.ixmlmaps.newenum | variant-key | metadata-only |
| IXmlNamespaces | Unknown | excel.ixmlnamespaces.count | excel.ixmlnamespaces.item | excel.ixmlnamespaces.newenum | variant-key | metadata-only |
| IXmlSchemas | Unknown | excel.ixmlschemas.count | excel.ixmlschemas.item | excel.ixmlschemas.newenum | variant-key | metadata-only |
| Labels | Unknown | excel.labels.count | excel.labels.item | excel.labels.newenum | variant-key | metadata-only |
| LegendEntries | Unknown | excel.legendentries.count | excel.legendentries.item | excel.legendentries.newenum | variant-key | metadata-only |
| Lines | Unknown | excel.lines.count | excel.lines.item | excel.lines.newenum | variant-key | metadata-only |
| ListBoxes | Unknown | excel.listboxes.count | excel.listboxes.item | excel.listboxes.newenum | variant-key | metadata-only |
| ListColumns | ListColumn | excel.listcolumns.count | excel.listcolumns.item | excel.listcolumns.newenum | one-based-integer, string-key | implemented |
| ListObjects | ListObject | excel.listobjects.count | excel.listobjects.item | excel.listobjects.newenum | one-based-integer, string-key | implemented |
| ListRows | ListRow | excel.listrows.count | excel.listrows.item | excel.listrows.newenum | one-based-integer | implemented |
| MenuBars | Unknown | excel.menubars.count | excel.menubars.item | excel.menubars.newenum | variant-key | metadata-only |
| MenuItems | Unknown | excel.menuitems.count | excel.menuitems.item | excel.menuitems.newenum | variant-key | metadata-only |
| Menus | Unknown | excel.menus.count | excel.menus.item | excel.menus.newenum | variant-key | metadata-only |
| ModelColumnChanges | Unknown | excel.modelcolumnchanges.count | excel.modelcolumnchanges.item | excel.modelcolumnchanges.newenum | variant-key | metadata-only |
| ModelColumnNames | Unknown | excel.modelcolumnnames.count | excel.modelcolumnnames.item | excel.modelcolumnnames.newenum | variant-key | metadata-only |
| ModelMeasureNames | Unknown | excel.modelmeasurenames.count | excel.modelmeasurenames.item | excel.modelmeasurenames.newenum | variant-key | metadata-only |
| ModelMeasures | Unknown | excel.modelmeasures.count | excel.modelmeasures.item | excel.modelmeasures.newenum | variant-key | metadata-only |
| ModelRelationships | Unknown | excel.modelrelationships.count | excel.modelrelationships.item | excel.modelrelationships.newenum | variant-key | metadata-only |
| ModelTableColumns | Unknown | excel.modeltablecolumns.count | excel.modeltablecolumns.item | excel.modeltablecolumns.newenum | variant-key | metadata-only |
| ModelTableNameChanges | Unknown | excel.modeltablenamechanges.count | excel.modeltablenamechanges.item | excel.modeltablenamechanges.newenum | variant-key | metadata-only |
| ModelTableNames | Unknown | excel.modeltablenames.count | excel.modeltablenames.item | -- | variant-key | metadata-only |
| ModelTables | Unknown | excel.modeltables.count | excel.modeltables.item | excel.modeltables.newenum | variant-key | metadata-only |
| Modules | Unknown | excel.modules.count | excel.modules.item | excel.modules.newenum | variant-key | metadata-only |
| Names | Name | excel.names.count | excel.names.item | excel.names.newenum | one-based-integer, string-key | implemented |
| ODBCErrors | Unknown | excel.odbcerrors.count | excel.odbcerrors.item | excel.odbcerrors.newenum | variant-key | metadata-only |
| OLEDBErrors | Unknown | excel.oledberrors.count | excel.oledberrors.item | excel.oledberrors.newenum | variant-key | metadata-only |
| OLEObjects | Unknown | excel.oleobjects.count | excel.oleobjects.item | excel.oleobjects.newenum | variant-key | metadata-only |
| OptionButtons | Unknown | excel.optionbuttons.count | excel.optionbuttons.item | excel.optionbuttons.newenum | variant-key | metadata-only |
| Ovals | Unknown | excel.ovals.count | excel.ovals.item | excel.ovals.newenum | variant-key | metadata-only |
| Pages | Unknown | excel.pages.count | excel.pages.item | excel.pages.newenum | variant-key | metadata-only |
| Panes | Unknown | excel.panes.count | excel.panes.item | -- | variant-key | metadata-only |
| Parameters | Unknown | excel.parameters.count | excel.parameters.item | excel.parameters.newenum | variant-key | metadata-only |
| Phonetics | Unknown | excel.phonetics.count | excel.phonetics.item | excel.phonetics.newenum | variant-key | metadata-only |
| Pictures | Unknown | excel.pictures.count | excel.pictures.item | excel.pictures.newenum | variant-key | metadata-only |
| PivotCaches | Unknown | excel.pivotcaches.count | excel.pivotcaches.item | excel.pivotcaches.newenum | variant-key | metadata-only |
| PivotFields | Unknown | excel.pivotfields.count | excel.pivotfields.item | excel.pivotfields.newenum | variant-key | metadata-only |
| PivotFilters | Unknown | excel.pivotfilters.count | excel.pivotfilters.item | excel.pivotfilters.newenum | variant-key | metadata-only |
| PivotFormulas | Unknown | excel.pivotformulas.count | excel.pivotformulas.item | excel.pivotformulas.newenum | variant-key | metadata-only |
| PivotItemList | Unknown | excel.pivotitemlist.count | excel.pivotitemlist.item | excel.pivotitemlist.newenum | variant-key | metadata-only |
| PivotItems | Unknown | excel.pivotitems.count | excel.pivotitems.item | excel.pivotitems.newenum | variant-key | metadata-only |
| PivotLineCells | Unknown | excel.pivotlinecells.count | excel.pivotlinecells.item | excel.pivotlinecells.newenum | variant-key | metadata-only |
| PivotLines | Unknown | excel.pivotlines.count | excel.pivotlines.item | excel.pivotlines.newenum | variant-key | metadata-only |
| PivotTableChangeList | Unknown | excel.pivottablechangelist.count | excel.pivottablechangelist.item | excel.pivottablechangelist.newenum | variant-key | metadata-only |
| PivotTables | Unknown | excel.pivottables.count | excel.pivottables.item | excel.pivottables.newenum | variant-key | metadata-only |
| Points | Unknown | excel.points.count | excel.points.item | excel.points.newenum | variant-key | metadata-only |
| ProtectedViewWindows | Unknown | excel.protectedviewwindows.count | excel.protectedviewwindows.item | excel.protectedviewwindows.newenum | variant-key | metadata-only |
| PublishedDocs | Unknown | excel.publisheddocs.count | excel.publisheddocs.item | excel.publisheddocs.newenum | variant-key | metadata-only |
| PublishObjects | Unknown | excel.publishobjects.count | excel.publishobjects.item | excel.publishobjects.newenum | variant-key | metadata-only |
| Queries | Unknown | excel.queries.count | excel.queries.item | excel.queries.newenum | variant-key | metadata-only |
| QueryTables | Unknown | excel.querytables.count | excel.querytables.item | excel.querytables.newenum | variant-key | metadata-only |
| Range | Unknown | excel.range.count | excel.range.item | excel.range.newenum | variant-key | metadata-only |
| Ranges | Unknown | excel.ranges.count | excel.ranges.item | excel.ranges.newenum | variant-key | metadata-only |
| RecentFiles | Unknown | excel.recentfiles.count | excel.recentfiles.item | excel.recentfiles.newenum | variant-key | metadata-only |
| Rectangles | Unknown | excel.rectangles.count | excel.rectangles.item | excel.rectangles.newenum | variant-key | metadata-only |
| Scenarios | Unknown | excel.scenarios.count | excel.scenarios.item | excel.scenarios.newenum | variant-key | metadata-only |
| ScrollBars | Unknown | excel.scrollbars.count | excel.scrollbars.item | excel.scrollbars.newenum | variant-key | metadata-only |
| SeriesCollection | Series | excel.seriescollection.count | excel.seriescollection.item | excel.seriescollection.newenum | variant-key | implemented |
| ServerViewableItems | Unknown | excel.serverviewableitems.count | excel.serverviewableitems.item | excel.serverviewableitems.newenum | variant-key | metadata-only |
| ShapeNodes | Unknown | excel.shapenodes.count | excel.shapenodes.item | excel.shapenodes.newenum | variant-key | metadata-only |
| ShapeRange | Unknown | excel.shaperange.count | excel.shaperange.item | excel.shaperange.newenum | variant-key | metadata-only |
| Shapes | Shape | excel.shapes.count | excel.shapes.item | excel.shapes.newenum | one-based-integer, string-key | implemented |
| Sheets | Unknown | excel.sheets.count | excel.sheets.item | excel.sheets.newenum | variant-key | metadata-only |
| SheetViews | Unknown | excel.sheetviews.count | excel.sheetviews.item | excel.sheetviews.newenum | variant-key | metadata-only |
| SlicerCacheLevels | Unknown | excel.slicercachelevels.count | excel.slicercachelevels.item | excel.slicercachelevels.newenum | variant-key | metadata-only |
| SlicerCaches | Unknown | excel.slicercaches.count | excel.slicercaches.item | excel.slicercaches.newenum | variant-key | metadata-only |
| SlicerItems | Unknown | excel.sliceritems.count | excel.sliceritems.item | excel.sliceritems.newenum | variant-key | metadata-only |
| SlicerPivotTables | Unknown | excel.slicerpivottables.count | excel.slicerpivottables.item | excel.slicerpivottables.newenum | variant-key | metadata-only |
| Slicers | Unknown | excel.slicers.count | excel.slicers.item | excel.slicers.newenum | variant-key | metadata-only |
| SmartTagActions | Unknown | excel.smarttagactions.count | excel.smarttagactions.item | excel.smarttagactions.newenum | variant-key | metadata-only |
| SmartTagRecognizers | Unknown | excel.smarttagrecognizers.count | excel.smarttagrecognizers.item | excel.smarttagrecognizers.newenum | variant-key | metadata-only |
| SortFields | SortField | excel.sortfields.count | excel.sortfields.item | excel.sortfields.newenum | one-based-integer | metadata-only |
| SparklineGroup | Unknown | excel.sparklinegroup.count | excel.sparklinegroup.item | excel.sparklinegroup.newenum | variant-key | metadata-only |
| SparklineGroups | SparklineGroup | excel.sparklinegroups.count | excel.sparklinegroups.item | excel.sparklinegroups.newenum | one-based-integer | implemented |
| Spinners | Unknown | excel.spinners.count | excel.spinners.item | excel.spinners.newenum | variant-key | metadata-only |
| Styles | Style | excel.styles.count | excel.styles.item | excel.styles.newenum | one-based-integer, string-key | implemented |
| TableStyleElements | Unknown | excel.tablestyleelements.count | excel.tablestyleelements.item | excel.tablestyleelements.newenum | variant-key | metadata-only |
| TableStyles | Unknown | excel.tablestyles.count | excel.tablestyles.item | excel.tablestyles.newenum | variant-key | metadata-only |
| TextBoxes | Unknown | excel.textboxes.count | excel.textboxes.item | excel.textboxes.newenum | variant-key | metadata-only |
| ToolbarButtons | Unknown | excel.toolbarbuttons.count | excel.toolbarbuttons.item | excel.toolbarbuttons.newenum | variant-key | metadata-only |
| Toolbars | Unknown | excel.toolbars.count | excel.toolbars.item | excel.toolbars.newenum | variant-key | metadata-only |
| Trendlines | Trendline | excel.trendlines.count | excel.trendlines.item | excel.trendlines.newenum | one-based-integer | implemented |
| UsedObjects | Unknown | excel.usedobjects.count | excel.usedobjects.item | excel.usedobjects.newenum | variant-key | metadata-only |
| UserAccessList | Unknown | excel.useraccesslist.count | excel.useraccesslist.item | excel.useraccesslist.newenum | variant-key | metadata-only |
| VPageBreaks | Unknown | excel.vpagebreaks.count | excel.vpagebreaks.item | excel.vpagebreaks.newenum | variant-key | metadata-only |
| Watches | Unknown | excel.watches.count | excel.watches.item | excel.watches.newenum | variant-key | metadata-only |
| Windows | Unknown | excel.windows.count | excel.windows.item | excel.windows.newenum | variant-key | metadata-only |
| Workbooks | Workbook | excel.workbooks.count | excel.workbooks.item | excel.workbooks.newenum | one-based-integer, string-key | implemented |
| Worksheets | Worksheet | excel.worksheets.count | excel.worksheets.item | excel.worksheets.newenum | one-based-integer, string-key | implemented |
| XmlMaps | Unknown | excel.xmlmaps.count | excel.xmlmaps.item | excel.xmlmaps.newenum | variant-key | metadata-only |
| XmlNamespaces | Unknown | excel.xmlnamespaces.count | excel.xmlnamespaces.item | excel.xmlnamespaces.newenum | variant-key | metadata-only |
| XmlSchemas | Unknown | excel.xmlschemas.count | excel.xmlschemas.item | excel.xmlschemas.newenum | variant-key | metadata-only |

## Prompt 16 typed collection policy

| Collection | Element | Index kinds | Iterator | Heterogeneous policy |
|---|---|---|---|---|
| ColorScaleCriteria | ColorScaleCriterion | one-based-integer | implemented | homogeneous |
| Comments | Comment | one-based-integer | implemented | homogeneous |
| CommentsThreaded | CommentThreaded | one-based-integer | implemented | homogeneous |
| FormatConditions | ConditionalFormat | one-based-integer | implemented | typed-subtype-by-type-property |
| Hyperlinks | Hyperlink | one-based-integer | implemented | homogeneous |
| IconCriteria | IconCriterion | one-based-integer | implemented | homogeneous |
| Styles | Style | one-based-integer, string-key | implemented | homogeneous |
