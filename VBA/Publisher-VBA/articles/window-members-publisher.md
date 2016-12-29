---
title: Window Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 7498ccd9-8b08-45f4-187e-87781871cdae
ms.locale: en-US
---


# Window Members (Publisher)
Represents a window. Many publication characteristics, such as scroll bars and rulers, are actually properties of the window.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Activate](window.activate-method-publisher.md)|Activates a window or OLE object.|
| [Move](window.move-method-publisher.md)|Moves the active document window.|
| [Resize](window.resize-method-publisher.md)|Sizes the Microsoft Publisher application window.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](window.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Caption](window.caption-property-publisher.md)|Returns or sets a  **String** indicating the caption at the top of the Microsoft Publisher application window. Read/write.|
| [Height](window.height-property-publisher.md)|Returns or sets a  **Long** that represents the height (in points) of the window. Read/write.|
| [Hwnd](window.hwnd-property-publisher.md)|Returns a  **Long** indicating the handle to the Microsoft Publisher application window. Read-only.|
| [Left](window.left-property-publisher.md)|Returns or sets a  **Long** indicating the position (in points) of the left edge of the application window relative to the left edge of the screen. Read/write.|
| [Parent](window.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Top](window.top-property-publisher.md)|Returns or sets a  **Long** that represents the distance between the top edge of the screen and the application window. Read/write.|
| [Visible](window.visible-property-publisher.md)| **True** if the window is visible. Read/write **Boolean**.|
| [Width](window.width-property-publisher.md)|Returns or sets a  **Long** that represents the width (in points) of the window. Read/write.|
| [WindowState](window.windowstate-property-publisher.md)|Returns or sets a  **PbWindowState** constant indicating the state of the Microsoft Publisher window. Read/write.|

