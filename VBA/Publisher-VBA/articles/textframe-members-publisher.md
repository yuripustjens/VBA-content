---
title: TextFrame Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: c74bdb5b-4699-614b-acc3-da0bcafa4138
ms.locale: en-US
---


# TextFrame Members (Publisher)
Represents the text frame in a  **[Shape](shape-object-publisher.md)** object. Contains the text in the text frame and the properties that control the margins and orientation of the text frame.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [BreakForwardLink](textframe.breakforwardlink-method-publisher.md)|Breaks the forward link for the specified text frame, if such a link exists.|
| [ValidLinkTarget](textframe.validlinktarget-method-publisher.md)|Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns  **True** if **_LinkTarget_** is a valid target, **False** if **_LinkTarget_** already contains text or is already linked, or if the shape does not support attached text.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](textframe.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoFitText](textframe.autofittext-property-publisher.md)|Sets or returns a  **PbTextAutoFitType**constant that represents how Microsoft Publisher automatically adjusts the text font size and the  **TextFrame** objects size for best viewing. Read/write.|
| [Columns](textframe.columns-property-publisher.md)|Returns or sets a  **Long** that represents the number of guide columns on a page or the number of columns in a text frame. Read/write.|
| [ColumnSpacing](textframe.columnspacing-property-publisher.md)|Returns or sets a  **Variant** that represents the amount of space between text columns. Read/write.|
| [HasNextLink](textframe.hasnextlink-property-publisher.md)|Indicates whether the specified text frame has a valid forward text-box link. Read-only.|
| [HasPreviousLink](textframe.haspreviouslink-property-publisher.md)|Returns  **msoTrue** if the specified text frame has a valid link to a backward text box and **msoFalse** if it does not. Read-only.|
| [HasText](textframe.hastext-property-publisher.md)|Returns an  **MsoTriState** constant indicating whether the specified shape has text associated with it. Read-only.|
| [IncludeContinuedFromPage](textframe.includecontinuedfrompage-property-publisher.md)|Determines whether the text "Continued from page  _pagenumber_" appears in a text box when Microsoft Publisher links text boxes. Read/write.|
| [IncludeContinuedOnPage](textframe.includecontinuedonpage-property-publisher.md)|Determines whether the text "Continued on page  _pagenumber_" appears in a text box when Microsoft Publisher links text boxes. Read/write.|
| [MarginBottom](textframe.marginbottom-property-publisher.md)|Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the bottom edge of a cell, text frame, or page. Read/write.|
| [MarginLeft](textframe.marginleft-property-publisher.md)|Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the left edge of a cell, text frame, or page. Read/write.|
| [MarginRight](textframe.marginright-property-publisher.md)|Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the right edge of a cell, text frame, or page. Read/write.|
| [MarginTop](textframe.margintop-property-publisher.md)|Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the top edge of a cell, text frame, or page. Read/write.|
| [NextLinkedTextFrame](textframe.nextlinkedtextframe-property-publisher.md)|Returns or sets a  **[TextFrame](textframe-object-publisher.md)** object representing the text frame to which text flows from the specified text frame. Read/write.|
| [Orientation](textframe.orientation-property-publisher.md)|Returns or sets a  **PbTextOrientation**constant that represents how text flows in a text box. Read/write.|
| [Overflowing](textframe.overflowing-property-publisher.md)|Indicates whether the text frame contains more text than can fit into the text frame. Read-only.|
| [Parent](textframe.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [PreviousLinkedTextFrame](textframe.previouslinkedtextframe-property-publisher.md)|Returns a  **[TextFrame](textframe-object-publisher.md)** object representing the text frame from which text flows to the specified text frame.|
| [Story](textframe.story-property-publisher.md)|Returns a  **Story** object that represents the story properties in a text range.|
| [TextRange](textframe.textrange-property-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the text that is attached to a shape, and properties and methods for manipulating the text.|
| [VerticalTextAlignment](textframe.verticaltextalignment-property-publisher.md)|Returns or sets a  **PbVerticalTextAlignmentType**constant that represents the vertical alignment of text in a text box. Read/write.|

