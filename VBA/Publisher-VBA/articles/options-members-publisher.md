---
title: Options Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: da3ce703-1c0c-d683-c8d7-aa1c8f1ab928
ms.locale: en-US
---


# Options Members (Publisher)
Represents application and publication options in Microsoft Publisher. Many of the properties for the  **Options** object correspond to items in the **Options** dialog box ( **Tools** menu).

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [ResetTips](options.resettips-method-publisher.md)|Resets tippages so that a user can view them when using features that have been used before.|
| [ResetWizardSynchronizing](options.resetwizardsynchronizing-method-publisher.md)|Resets the data that Microsoft Publisher uses to automatically change similar objects to have the same formatting or content.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [AddHebDoubleQuote](options.addhebdoublequote-property-publisher.md)| **True** for Microsoft Publisher to display double quotes for Hebrew alphabet numbering. Default is **False**. Read/write  **Boolean**.|
| [AllowBackgroundSave](options.allowbackgroundsave-property-publisher.md)| **True** (default) for Microsoft Publisher to save publications in the background, allowing users to perform other actions at the same time. Read/write **Boolean**.|
| [Application](options.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoFormatWord](options.autoformatword-property-publisher.md)| **True** for Microsoft Publisher to automatically format the entire word at the cursor position, even when no text is selected. Read/write **Boolean**.|
| [AutoHyphenate](options.autohyphenate-property-publisher.md)| **True** (default) for Microsoft Publisher to automatically hyphenate text in text frames. Read/write **Boolean**.|
| [AutoKeyboardSwitching](options.autokeyboardswitching-property-publisher.md)| **True** for Microsoft Publisher to automatically switch the keyboard language to the language used for the text at the cursor position. Read/write **Boolean**.|
| [AutoSelectWord](options.autoselectword-property-publisher.md)| **True** for Microsoft Publisher to automatically select the entire word when selecting text. Read/write **Boolean**.|
| [DefaultPubDirection](options.defaultpubdirection-property-publisher.md)|Returns or sets a  **PbDirectionType** constant that represents the default direction in which text flows when a new publication is created. Read/write.|
| [DefaultTextFlowDirection](options.defaulttextflowdirection-property-publisher.md)|Returns or sets a  **PbDirectionType** constant that represents a global Microsoft Publisher option, indicating whether text flows from left to right or from right to left in a publication. Read/write.|
| [DisplayStatusBar](options.displaystatusbar-property-publisher.md)| **True** for Microsoft Publisher to show the status bar at the bottom of the Publisher window. Read/write **Boolean**.|
| [DragAndDropText](options.draganddroptext-property-publisher.md)| **True** to enable dragging of text. Read/write **Boolean**.|
| [HyphenationZone](options.hyphenationzone-property-publisher.md)|Returns or sets a  **Variant** that represents the maximum amount of space that Microsoft Publisher leaves between the end of the last word in a line and the right margin. Read/write.|
| [MeasurementUnit](options.measurementunit-property-publisher.md)|Returns or sets a  **PbUnitType** constant representing the standard measurement unit for Microsoft Publisher. Read/write.|
| [Parent](options.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [PathForPictures](options.pathforpictures-property-publisher.md)|Returns a  **String** that represents the default path for picture files. Read.|
| [PathForPublications](options.pathforpublications-property-publisher.md)|Returns a  **String** that represents the default folder for publications. Read.|
| [SaveAutoRecoverInfo](options.saveautorecoverinfo-property-publisher.md)| **True** if Microsoft Publisher automatically saves publications for recovery if the application is unexpectedly shut down. Read/write **Boolean**.|
| [SaveAutoRecoverInfoInterval](options.saveautorecoverinfointerval-property-publisher.md)|Returns or sets a  **Long** that represents the time interval in minutes for automatically saving a publication for recovery if the application is unexpectedly shut down. Read/write.|
| [SequenceCheck](options.sequencecheck-property-publisher.md)| **True** to check the sequence of independent characters for Asian text. Read/write **Boolean**.|
| [ShowBasicColors](options.showbasiccolors-property-publisher.md)|Returns or sets a  **Boolean** indicating whether Microsoft Publisher shows basic colors in the color palette; **True** to show basic colors in the palette. Read/write.|
| [ShowScreenTipsOnObjects](options.showscreentipsonobjects-property-publisher.md)| **True** for Microsoft Publisher to display ScreenTips when the mouse pointer hovers over a text box, shape or other object. Read/write **Boolean**.|
| [ShowTipPages](options.showtippages-property-publisher.md)| **True** for Microsoft Publisher to display tippages in balloons. Read/write **Boolean**.|
| [TypeNReplace](options.typenreplace-property-publisher.md)| **True** for Microsoft Publisher to replace unreadable Asian character clusters resulting from invalid keyboard sequences. Read/write **Boolean**.|
| [UseCatalogAtStartup](options.usecatalogatstartup-property-publisher.md)| **True** for Microsoft Publisher to show the catalog when starting. Read/write **Boolean**.|
| [UseWizardForBlankPublication](options.usewizardforblankpublication-property-publisher.md)|Gets or sets whether to use a wizard for blank publications. Read/Write.|

