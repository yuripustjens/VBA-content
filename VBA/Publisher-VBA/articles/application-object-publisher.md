---
title: Application Object (Publisher)
keywords: vbapb10.chm536936447
f1_keywords:
- vbapb10.chm536936447
ms.prod: PUBLISHER
api_name:
- Publisher.Application
ms.assetid: acfc7efb-e6a5-a89a-3aee-3cb4af2f3508
ms.locale: en-US
---


# Application Object (Publisher)

Represents the Microsoft Publisher application. The  **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.
 


## Remarks

When using Microsoft Visual Basic for Applications in Publisher, all of the properties and methods of the  **Application** object can be used without the **Application** object qualifier. For example, instead of typing `Application.ActiveDocument.PrintOut`, you can type  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **<globals>** at the top of the list in the **Classes** box. When accessing the Publisher object model from a non-Publisher project, all properties and methods must be fully qualified.
 

 

## Example

Use the  ** [Application](application.application-property-publisher.md)** property to return the **Application** object. The following example displays the application name.
 

 

```
Sub ShowAppName() 
 MsgBox Application.Name 
End Sub
```


## Events



|**Name**|
|:-----|
| [AfterPrint](application.afterprint-event-publisher.md)|
| [BeforePrint](application.beforeprint-event-publisher.md)|
| [DocumentBeforeClose](application.documentbeforeclose-event-publisher.md)|
| [DocumentOpen](application.documentopen-event-publisher.md)|
| [HideCatalogUI](application.hidecatalogui-event-publisher.md)|
| [MailMergeAfterMerge](application.mailmergeaftermerge-event-publisher.md)|
| [MailMergeAfterRecordMerge](application.mailmergeafterrecordmerge-event-publisher.md)|
| [MailMergeBeforeMerge](application.mailmergebeforemerge-event-publisher.md)|
| [MailMergeBeforeRecordMerge](application.mailmergebeforerecordmerge-event-publisher.md)|
| [MailMergeDataSourceLoad](application.mailmergedatasourceload-event-publisher.md)|
| [MailMergeDataSourceValidate](application.mailmergedatasourcevalidate-event-publisher.md)|
| [MailMergeGenerateBarcode](application.mailmergegeneratebarcode-event-publisher.md)|
| [MailMergeInsertBarcode](application.mailmergeinsertbarcode-event-publisher.md)|
| [MailMergeRecipientListClose](application.mailmergerecipientlistclose-event-publisher.md)|
| [MailMergeWizardFollowUpCustom](application.mailmergewizardfollowupcustom-event-publisher.md)|
| [MailMergeWizardStateChange](application.mailmergewizardstatechange-event-publisher.md)|
| [NewDocument](application.newdocument-event-publisher.md)|
| [Quit](application.quit-event-publisher.md)|
| [ShowCatalogUI](application.showcatalogui-event-publisher.md)|
| [WindowActivate](application.windowactivate-event-publisher.md)|
| [WindowDeactivate](application.windowdeactivate-event-publisher.md)|
| [WindowPageChange](application.windowpagechange-event-publisher.md)|

## Methods



|**Name**|
|:-----|
| [CentimetersToPoints](application.centimeterstopoints-method-publisher.md)|
| [ChangeFileOpenDirectory](application.changefileopendirectory-method-publisher.md)|
| [EmusToPoints](application.emustopoints-method-publisher.md)|
| [Help](application.help-method-publisher.md)|
| [InchesToPoints](application.inchestopoints-method-publisher.md)|
| [IsValidObject](application.isvalidobject-method-publisher.md)|
| [LinesToPoints](application.linestopoints-method-publisher.md)|
| [MillimetersToPoints](application.millimeterstopoints-method-publisher.md)|
| [NewDocument](application.newdocument-method-publisher.md)|
| [Open](application.open-method-publisher.md)|
| [PicasToPoints](application.picastopoints-method-publisher.md)|
| [PixelsToPoints](application.pixelstopoints-method-publisher.md)|
| [PointsToCentimeters](application.pointstocentimeters-method-publisher.md)|
| [PointsToEmus](application.pointstoemus-method-publisher.md)|
| [PointsToInches](application.pointstoinches-method-publisher.md)|
| [PointsToLines](application.pointstolines-method-publisher.md)|
| [PointsToMillimeters](application.pointstomillimeters-method-publisher.md)|
| [PointsToPicas](application.pointstopicas-method-publisher.md)|
| [PointsToPixels](application.pointstopixels-method-publisher.md)|
| [PointsToTwips](application.pointstotwips-method-publisher.md)|
| [Quit](application.quit-method-publisher.md)|
| [ShowWizardCatalog](application.showwizardcatalog-method-publisher.md)|
| [TwipsToPoints](application.twipstopoints-method-publisher.md)|

## Properties



|**Name**|
|:-----|
| [ActiveDocument](application.activedocument-property-publisher.md)|
| [ActiveWindow](application.activewindow-property-publisher.md)|
| [Application](application.application-property-publisher.md)|
| [Assistance](application.assistance-property-publisher.md)|
| [AutomationSecurity](application.automationsecurity-property-publisher.md)|
| [Build](application.build-property-publisher.md)|
| [CaptionStyles](application.captionstyles-property-publisher.md)|
| [ColorSchemes](application.colorschemes-property-publisher.md)|
| [COMAddIns](application.comaddins-property-publisher.md)|
| [CommandBars](application.commandbars-property-publisher.md)|
| [Documents](application.documents-property-publisher.md)|
| [FileDialog](application.filedialog-property-publisher.md)|
| [InsertBarcodeVisible](application.insertbarcodevisible-property-publisher.md)|
| [InstalledPrinters](application.installedprinters-property-publisher.md)|
| [Language](application.language-property-publisher.md)|
| [Name](application.name-property-publisher.md)|
| [OfficeDataSourceObject](application.officedatasourceobject-property-publisher.md)|
| [Options](application.options-property-publisher.md)|
| [Parent](application.parent-property-publisher.md)|
| [Path](application.path-property-publisher.md)|
| [PathSeparator](application.pathseparator-property-publisher.md)|
| [PrintPreview](application.printpreview-property-publisher.md)|
| [ProductCode](application.productcode-property-publisher.md)|
| [ScreenUpdating](application.screenupdating-property-publisher.md)|
| [Selection](application.selection-property-publisher.md)|
| [ShowFollowUpCustom](application.showfollowupcustom-property-publisher.md)|
| [SnapToGuides](application.snaptoguides-property-publisher.md)|
| [SnapToObjects](application.snaptoobjects-property-publisher.md)|
| [TemplateFolderPath](application.templatefolderpath-property-publisher.md)|
| [ValidateAddressVisible](application.validateaddressvisible-property-publisher.md)|
| [Version](application.version-property-publisher.md)|
| [WebOptions](application.weboptions-property-publisher.md)|
| [WizardCatalogVisible](application.wizardcatalogvisible-property-publisher.md)|

## See also


#### Other resources


 
 [Application Object Members](http://msdn.microsoft.com/library/application-members-publisher%28Office.15%29.aspx)
