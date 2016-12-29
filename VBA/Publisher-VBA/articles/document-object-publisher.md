---
title: Document Object (Publisher)
keywords: vbapb10.chm553713663
f1_keywords: vbapb10.chm553713663
ms.prod: PUBLISHER
api_name: Publisher.Document
ms.assetid: 44f02255-ff5b-bcfe-900f-61c8fdf61ef3
ms.locale: en-US
---


# Document Object (Publisher)

Represents a publication. 


## Example

Use the  **[ActiveDocument](http://msdn.microsoft.com/library/application.activedocument-property-publisher%28Office.15%29.aspx)** property to refer to the current publication. This example adds a table to the first page of the active publication.


```
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```

You can also write the above routine by using a reference to the  **ThisDocument** module. This example uses a **ThisDocument** reference instead of **ActiveDocument**.




```
Sub PrintPublication() 
 With ThisDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```


## Events



|**Name**|
|:-----|
|[BeforeClose](http://msdn.microsoft.com/library/document.beforeclose-event-publisher%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/document.open-event-publisher%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/document.redo-event-publisher%28Office.15%29.aspx)|
|[ShapesAdded](http://msdn.microsoft.com/library/document.shapesadded-event-publisher%28Office.15%29.aspx)|
|[ShapesRemoved](http://msdn.microsoft.com/library/document.shapesremoved-event-publisher%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/document.undo-event-publisher%28Office.15%29.aspx)|
|[WizardAfterChange](http://msdn.microsoft.com/library/document.wizardafterchange-event-publisher%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[BeginCustomUndoAction](http://msdn.microsoft.com/library/document.begincustomundoaction-method-publisher%28Office.15%29.aspx)|
|[ChangeDocument](http://msdn.microsoft.com/library/document.changedocument-method-publisher%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/document.close-method-publisher%28Office.15%29.aspx)|
|[ConvertPublicationType](http://msdn.microsoft.com/library/document.convertpublicationtype-method-publisher%28Office.15%29.aspx)|
|[EndCustomUndoAction](http://msdn.microsoft.com/library/document.endcustomundoaction-method-publisher%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/document.exportasfixedformat-method-publisher%28Office.15%29.aspx)|
|[FindShapeByWizardTag](http://msdn.microsoft.com/library/document.findshapebywizardtag-method-publisher%28Office.15%29.aspx)|
|[FindShapesByTag](http://msdn.microsoft.com/library/document.findshapesbytag-method-publisher%28Office.15%29.aspx)|
|[PrintOutEx](http://msdn.microsoft.com/library/document.printoutex-method-publisher%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/document.redo-method-publisher%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/document.save-method-publisher%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/document.saveas-method-publisher%28Office.15%29.aspx)|
|[SetBusinessInformation](http://msdn.microsoft.com/library/document.setbusinessinformation-method-publisher%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/document.undo-method-publisher%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/document.undoclear-method-publisher%28Office.15%29.aspx)|
|[UpdateOLEObjects](http://msdn.microsoft.com/library/document.updateoleobjects-method-publisher%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/document.webpagepreview-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveView](http://msdn.microsoft.com/library/document.activeview-property-publisher%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/document.activewindow-property-publisher%28Office.15%29.aspx)|
|[AdvancedPrintOptions](http://msdn.microsoft.com/library/document.advancedprintoptions-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/document.application-property-publisher%28Office.15%29.aspx)|
|[AvailableBuildingBlocks](http://msdn.microsoft.com/library/document.availablebuildingblocks-property-publisher%28Office.15%29.aspx)|
|[BorderArts](http://msdn.microsoft.com/library/document.borderarts-property-publisher%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/document.colorscheme-property-publisher%28Office.15%29.aspx)|
|[DefaultTabStop](http://msdn.microsoft.com/library/document.defaulttabstop-property-publisher%28Office.15%29.aspx)|
|[DocumentDirection](http://msdn.microsoft.com/library/document.documentdirection-property-publisher%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/document.envelopevisible-property-publisher%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/document.find-property-publisher%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/document.fullname-property-publisher%28Office.15%29.aspx)|
|[IsDataSourceConnected](http://msdn.microsoft.com/library/document.isdatasourceconnected-property-publisher%28Office.15%29.aspx)|
|[IsWizard](http://msdn.microsoft.com/library/document.iswizard-property-publisher%28Office.15%29.aspx)|
|[LayoutGuides](http://msdn.microsoft.com/library/document.layoutguides-property-publisher%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/document.mailenvelope-property-publisher%28Office.15%29.aspx)|
|[MailMerge](http://msdn.microsoft.com/library/document.mailmerge-property-publisher%28Office.15%29.aspx)|
|[MasterPages](http://msdn.microsoft.com/library/document.masterpages-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/document.name-property-publisher%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/document.pages-property-publisher%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/document.pagesetup-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/document.parent-property-publisher%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/document.path-property-publisher%28Office.15%29.aspx)|
|[PrintPageBackgrounds](http://msdn.microsoft.com/library/document.printpagebackgrounds-property-publisher%28Office.15%29.aspx)|
|[PrintStyle](http://msdn.microsoft.com/library/document.printstyle-property-publisher%28Office.15%29.aspx)|
|[PublicationType](http://msdn.microsoft.com/library/document.publicationtype-property-publisher%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/document.readonly-property-publisher%28Office.15%29.aspx)|
|[RedoActionsAvailable](http://msdn.microsoft.com/library/document.redoactionsavailable-property-publisher%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/document.removepersonalinformation-property-publisher%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/document.saved-property-publisher%28Office.15%29.aspx)|
|[SaveFormat](http://msdn.microsoft.com/library/document.saveformat-property-publisher%28Office.15%29.aspx)|
|[ScratchArea](http://msdn.microsoft.com/library/document.scratcharea-property-publisher%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/document.sections-property-publisher%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/document.selection-property-publisher%28Office.15%29.aspx)|
|[Stories](http://msdn.microsoft.com/library/document.stories-property-publisher%28Office.15%29.aspx)|
|[SurplusShapes](http://msdn.microsoft.com/library/document.surplusshapes-property-publisher%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/document.tags-property-publisher%28Office.15%29.aspx)|
|[TextStyles](http://msdn.microsoft.com/library/document.textstyles-property-publisher%28Office.15%29.aspx)|
|[UndoActionsAvailable](http://msdn.microsoft.com/library/document.undoactionsavailable-property-publisher%28Office.15%29.aspx)|
|[ViewBoundaries](http://msdn.microsoft.com/library/document.viewboundaries-property-publisher%28Office.15%29.aspx)|
|[ViewGuides](http://msdn.microsoft.com/library/document.viewguides-property-publisher%28Office.15%29.aspx)|
|[ViewHorizontalBaseLineGuides](http://msdn.microsoft.com/library/document.viewhorizontalbaselineguides-property-publisher%28Office.15%29.aspx)|
|[ViewTwoPageSpread](http://msdn.microsoft.com/library/document.viewtwopagespread-property-publisher%28Office.15%29.aspx)|
|[ViewVerticalBaseLineGuides](http://msdn.microsoft.com/library/document.viewverticalbaselineguides-property-publisher%28Office.15%29.aspx)|
|[WebNavigationBarSets](http://msdn.microsoft.com/library/document.webnavigationbarsets-property-publisher%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/document.wizard-property-publisher%28Office.15%29.aspx)|

