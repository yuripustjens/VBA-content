---
title: Application Object (Publisher)
keywords: vbapb10.chm536936447
f1_keywords: vbapb10.chm536936447
ms.prod: PUBLISHER
api_name: Publisher.Application
ms.assetid: acfc7efb-e6a5-a89a-3aee-3cb4af2f3508
ms.locale: en-US
---


# Application Object (Publisher)

Represents the Microsoft Publisher application. The  **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.


## Remarks

When using Microsoft Visual Basic for Applications in Publisher, all of the properties and methods of the  **Application** object can be used without the **Application** object qualifier. For example, instead of typing `Application.ActiveDocument.PrintOut`, you can type  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **<globals>** at the top of the list in the **Classes** box. When accessing the Publisher object model from a non-Publisher project, all properties and methods must be fully qualified.


## Example

Use the  **[Application](http://msdn.microsoft.com/library/application.application-property-publisher%28Office.15%29.aspx)** property to return the **Application** object. The following example displays the application name.


```
Sub ShowAppName() 
 MsgBox Application.Name 
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterPrint](http://msdn.microsoft.com/library/application.afterprint-event-publisher%28Office.15%29.aspx)|
|[BeforePrint](http://msdn.microsoft.com/library/application.beforeprint-event-publisher%28Office.15%29.aspx)|
|[DocumentBeforeClose](http://msdn.microsoft.com/library/application.documentbeforeclose-event-publisher%28Office.15%29.aspx)|
|[DocumentOpen](http://msdn.microsoft.com/library/application.documentopen-event-publisher%28Office.15%29.aspx)|
|[HideCatalogUI](http://msdn.microsoft.com/library/application.hidecatalogui-event-publisher%28Office.15%29.aspx)|
|[MailMergeAfterMerge](http://msdn.microsoft.com/library/application.mailmergeaftermerge-event-publisher%28Office.15%29.aspx)|
|[MailMergeAfterRecordMerge](http://msdn.microsoft.com/library/application.mailmergeafterrecordmerge-event-publisher%28Office.15%29.aspx)|
|[MailMergeBeforeMerge](http://msdn.microsoft.com/library/application.mailmergebeforemerge-event-publisher%28Office.15%29.aspx)|
|[MailMergeBeforeRecordMerge](http://msdn.microsoft.com/library/application.mailmergebeforerecordmerge-event-publisher%28Office.15%29.aspx)|
|[MailMergeDataSourceLoad](http://msdn.microsoft.com/library/application.mailmergedatasourceload-event-publisher%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate](http://msdn.microsoft.com/library/application.mailmergedatasourcevalidate-event-publisher%28Office.15%29.aspx)|
|[MailMergeGenerateBarcode](http://msdn.microsoft.com/library/application.mailmergegeneratebarcode-event-publisher%28Office.15%29.aspx)|
|[MailMergeInsertBarcode](http://msdn.microsoft.com/library/application.mailmergeinsertbarcode-event-publisher%28Office.15%29.aspx)|
|[MailMergeRecipientListClose](http://msdn.microsoft.com/library/application.mailmergerecipientlistclose-event-publisher%28Office.15%29.aspx)|
|[MailMergeWizardFollowUpCustom](http://msdn.microsoft.com/library/application.mailmergewizardfollowupcustom-event-publisher%28Office.15%29.aspx)|
|[MailMergeWizardStateChange](http://msdn.microsoft.com/library/application.mailmergewizardstatechange-event-publisher%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/application.newdocument-event-publisher%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application.quit-event-publisher%28Office.15%29.aspx)|
|[ShowCatalogUI](http://msdn.microsoft.com/library/application.showcatalogui-event-publisher%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application.windowactivate-event-publisher%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/application.windowdeactivate-event-publisher%28Office.15%29.aspx)|
|[WindowPageChange](http://msdn.microsoft.com/library/application.windowpagechange-event-publisher%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[CentimetersToPoints](http://msdn.microsoft.com/library/application.centimeterstopoints-method-publisher%28Office.15%29.aspx)|
|[ChangeFileOpenDirectory](http://msdn.microsoft.com/library/application.changefileopendirectory-method-publisher%28Office.15%29.aspx)|
|[EmusToPoints](http://msdn.microsoft.com/library/application.emustopoints-method-publisher%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/application.help-method-publisher%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/application.inchestopoints-method-publisher%28Office.15%29.aspx)|
|[IsValidObject](http://msdn.microsoft.com/library/application.isvalidobject-method-publisher%28Office.15%29.aspx)|
|[LinesToPoints](http://msdn.microsoft.com/library/application.linestopoints-method-publisher%28Office.15%29.aspx)|
|[MillimetersToPoints](http://msdn.microsoft.com/library/application.millimeterstopoints-method-publisher%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/application.newdocument-method-publisher%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/application.open-method-publisher%28Office.15%29.aspx)|
|[PicasToPoints](http://msdn.microsoft.com/library/application.picastopoints-method-publisher%28Office.15%29.aspx)|
|[PixelsToPoints](http://msdn.microsoft.com/library/application.pixelstopoints-method-publisher%28Office.15%29.aspx)|
|[PointsToCentimeters](http://msdn.microsoft.com/library/application.pointstocentimeters-method-publisher%28Office.15%29.aspx)|
|[PointsToEmus](http://msdn.microsoft.com/library/application.pointstoemus-method-publisher%28Office.15%29.aspx)|
|[PointsToInches](http://msdn.microsoft.com/library/application.pointstoinches-method-publisher%28Office.15%29.aspx)|
|[PointsToLines](http://msdn.microsoft.com/library/application.pointstolines-method-publisher%28Office.15%29.aspx)|
|[PointsToMillimeters](http://msdn.microsoft.com/library/application.pointstomillimeters-method-publisher%28Office.15%29.aspx)|
|[PointsToPicas](http://msdn.microsoft.com/library/application.pointstopicas-method-publisher%28Office.15%29.aspx)|
|[PointsToPixels](http://msdn.microsoft.com/library/application.pointstopixels-method-publisher%28Office.15%29.aspx)|
|[PointsToTwips](http://msdn.microsoft.com/library/application.pointstotwips-method-publisher%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application.quit-method-publisher%28Office.15%29.aspx)|
|[ShowWizardCatalog](http://msdn.microsoft.com/library/application.showwizardcatalog-method-publisher%28Office.15%29.aspx)|
|[TwipsToPoints](http://msdn.microsoft.com/library/application.twipstopoints-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveDocument](http://msdn.microsoft.com/library/application.activedocument-property-publisher%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application.activewindow-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/application.application-property-publisher%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application.assistance-property-publisher%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application.automationsecurity-property-publisher%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application.build-property-publisher%28Office.15%29.aspx)|
|[CaptionStyles](http://msdn.microsoft.com/library/application.captionstyles-property-publisher%28Office.15%29.aspx)|
|[ColorSchemes](http://msdn.microsoft.com/library/application.colorschemes-property-publisher%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application.comaddins-property-publisher%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application.commandbars-property-publisher%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/application.documents-property-publisher%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/application.filedialog-property-publisher%28Office.15%29.aspx)|
|[InsertBarcodeVisible](http://msdn.microsoft.com/library/application.insertbarcodevisible-property-publisher%28Office.15%29.aspx)|
|[InstalledPrinters](http://msdn.microsoft.com/library/application.installedprinters-property-publisher%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/application.language-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application.name-property-publisher%28Office.15%29.aspx)|
|[OfficeDataSourceObject](http://msdn.microsoft.com/library/application.officedatasourceobject-property-publisher%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/application.options-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application.parent-property-publisher%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application.path-property-publisher%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/application.pathseparator-property-publisher%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/application.printpreview-property-publisher%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application.productcode-property-publisher%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/application.screenupdating-property-publisher%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/application.selection-property-publisher%28Office.15%29.aspx)|
|[ShowFollowUpCustom](http://msdn.microsoft.com/library/application.showfollowupcustom-property-publisher%28Office.15%29.aspx)|
|[SnapToGuides](http://msdn.microsoft.com/library/application.snaptoguides-property-publisher%28Office.15%29.aspx)|
|[SnapToObjects](http://msdn.microsoft.com/library/application.snaptoobjects-property-publisher%28Office.15%29.aspx)|
|[TemplateFolderPath](http://msdn.microsoft.com/library/application.templatefolderpath-property-publisher%28Office.15%29.aspx)|
|[ValidateAddressVisible](http://msdn.microsoft.com/library/application.validateaddressvisible-property-publisher%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application.version-property-publisher%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/application.weboptions-property-publisher%28Office.15%29.aspx)|
|[WizardCatalogVisible](http://msdn.microsoft.com/library/application.wizardcatalogvisible-property-publisher%28Office.15%29.aspx)|

## See also


#### Other resources


[Application Object Members](http://msdn.microsoft.com/library/application-members-publisher%28Office.15%29.aspx)
