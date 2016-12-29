---
title: ReaderSpread Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 34aee36d-1c3b-7aa6-3594-eb76e5602174
ms.locale: en-US
---


# ReaderSpread Members (Publisher)
Represents the reader spread (not the printer spread) for the page. A reader spread generally contains one or two pages. The  **ReaderSpread** object properties provide information about whether pages are facing and how those pages are laid out. For example, in facing page view, pages two and three can be side-by-side or one on top of the other.

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](readerspread.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Height](readerspread.height-property-publisher.md)|Returns a  **Single** that represents the height, in points, of the page. Read-only.|
| [Left](readerspread.left-property-publisher.md)|Returns a  **Single** indicating the position (in points) of the left edge of the reader spread from the workspace. Read-only.|
| [PageCount](readerspread.pagecount-property-publisher.md)|Returns a  **Long** indicating the number of pages in the specified reader spread. Read-only.|
| [Pages](readerspread.pages-property-publisher.md)|Returns a  **[Page](page-object-publisher.md)** object representing one of the pages that compose the specified reader spread.|
| [Parent](readerspread.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Top](readerspread.top-property-publisher.md)|Returns the a  **Single** that represents the distance (in points) from the top edge of the workspace to the top edge of the page. Read-only.|
| [Width](readerspread.width-property-publisher.md)|Returns a  **Single** that represents the width, in points, of the page. Read-only.|
|Name|Description|

