---
title: Options Object (Publisher)
keywords: vbapb10.chm1114111
f1_keywords: vbapb10.chm1114111
ms.prod: PUBLISHER
api_name: Publisher.Options
ms.assetid: 2554cd33-9d94-2622-6fab-19ca33d5a561
ms.locale: en-US
---


# Options Object (Publisher)

Represents application and publication options in Microsoft Publisher. Many of the properties for the  **Options** object correspond to items in the **Options** dialog box ( **Tools** menu).


## Example

Use the  **[Options](http://msdn.microsoft.com/library/application.options-property-publisher%28Office.15%29.aspx)** property to return the **Options** object. The following example sets four application options for Publisher.


```
Sub SetSpecialOptions() 
 With Options 
 .AllowBackgroundSave = True 
 .DragAndDropText = True 
 .AutoHyphenate = True 
 .MeasurementUnit = pbUnitInch 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ResetTips](http://msdn.microsoft.com/library/options.resettips-method-publisher%28Office.15%29.aspx)|
|[ResetWizardSynchronizing](http://msdn.microsoft.com/library/options.resetwizardsynchronizing-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddHebDoubleQuote](http://msdn.microsoft.com/library/options.addhebdoublequote-property-publisher%28Office.15%29.aspx)|
|[AllowBackgroundSave](http://msdn.microsoft.com/library/options.allowbackgroundsave-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/options.application-property-publisher%28Office.15%29.aspx)|
|[AutoFormatWord](http://msdn.microsoft.com/library/options.autoformatword-property-publisher%28Office.15%29.aspx)|
|[AutoHyphenate](http://msdn.microsoft.com/library/options.autohyphenate-property-publisher%28Office.15%29.aspx)|
|[AutoKeyboardSwitching](http://msdn.microsoft.com/library/options.autokeyboardswitching-property-publisher%28Office.15%29.aspx)|
|[AutoSelectWord](http://msdn.microsoft.com/library/options.autoselectword-property-publisher%28Office.15%29.aspx)|
|[DefaultPubDirection](http://msdn.microsoft.com/library/options.defaultpubdirection-property-publisher%28Office.15%29.aspx)|
|[DefaultTextFlowDirection](http://msdn.microsoft.com/library/options.defaulttextflowdirection-property-publisher%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/options.displaystatusbar-property-publisher%28Office.15%29.aspx)|
|[DragAndDropText](http://msdn.microsoft.com/library/options.draganddroptext-property-publisher%28Office.15%29.aspx)|
|[HyphenationZone](http://msdn.microsoft.com/library/options.hyphenationzone-property-publisher%28Office.15%29.aspx)|
|[MeasurementUnit](http://msdn.microsoft.com/library/options.measurementunit-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/options.parent-property-publisher%28Office.15%29.aspx)|
|[PathForPictures](http://msdn.microsoft.com/library/options.pathforpictures-property-publisher%28Office.15%29.aspx)|
|[PathForPublications](http://msdn.microsoft.com/library/options.pathforpublications-property-publisher%28Office.15%29.aspx)|
|[SaveAutoRecoverInfo](http://msdn.microsoft.com/library/options.saveautorecoverinfo-property-publisher%28Office.15%29.aspx)|
|[SaveAutoRecoverInfoInterval](http://msdn.microsoft.com/library/options.saveautorecoverinfointerval-property-publisher%28Office.15%29.aspx)|
|[SequenceCheck](http://msdn.microsoft.com/library/options.sequencecheck-property-publisher%28Office.15%29.aspx)|
|[ShowBasicColors](http://msdn.microsoft.com/library/options.showbasiccolors-property-publisher%28Office.15%29.aspx)|
|[ShowScreenTipsOnObjects](http://msdn.microsoft.com/library/options.showscreentipsonobjects-property-publisher%28Office.15%29.aspx)|
|[ShowTipPages](http://msdn.microsoft.com/library/options.showtippages-property-publisher%28Office.15%29.aspx)|
|[TypeNReplace](http://msdn.microsoft.com/library/options.typenreplace-property-publisher%28Office.15%29.aspx)|
|[UseCatalogAtStartup](http://msdn.microsoft.com/library/options.usecatalogatstartup-property-publisher%28Office.15%29.aspx)|
|[UseWizardForBlankPublication](http://msdn.microsoft.com/library/options.usewizardforblankpublication-property-publisher%28Office.15%29.aspx)|

