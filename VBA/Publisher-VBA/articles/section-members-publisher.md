---
title: Section Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: ce13a263-d9c1-2919-fd0e-a7e8decd1bc8
ms.locale: en-US
---


# Section Members (Publisher)
Represents a Section of a publication or document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Delete](section.delete-method-publisher.md)|Deletes the specified object.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](section.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [ContinueNumbersFromPreviousSection](section.continuenumbersfromprevioussection-property-publisher.md)| **True** if the specified section continues the numbering from the prvious section. Read/write **Boolean**.|
| [PageNumberFormat](section.pagenumberformat-property-publisher.md)|Sets or returns a  **PbPageNumberFormat** constant that reperesents the formatting of the page numbering. Read/write.|
| [PageNumberStart](section.pagenumberstart-property-publisher.md)|Sets or returns the page number that the specified section starts with. Read/write  **Long**.|
| [Parent](section.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ShowHeaderFooterOnFirstPage](section.showheaderfooteronfirstpage-property-publisher.md)| **True** if the header and footer of the specified section will be visible. Read/write **Boolean**.|
| [StartPageIndex](section.startpageindex-property-publisher.md)|Returns the page number of the page that the specified  **Section** object begins on. Read-only.|

