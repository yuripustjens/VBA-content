---
title: TabStop Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 1f7fac1b-6a80-963c-d714-ef75be225fe6
ms.locale: en-US
---


# TabStop Members (Publisher)
Represents a single tab stop. The  **TabStop** object is a member of the **[TabStops](tabstops-object-publisher.md)** collection. The **TabStops** collection represents all the custom and default tab stops in a paragraph or group of paragraphs.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Clear](tabstop.clear-method-publisher.md)|Removes the specified custom tab stop.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Alignment](tabstop.alignment-property-publisher.md)|Returns or sets a  **PbTabAlignmentType** constant that represents the alignment for the specified tab stop. Read/write.|
| [Application](tabstop.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Leader](tabstop.leader-property-publisher.md)|Sets or returns a  **PbTabLeaderType** constant that represents the leader character for a tab stop. Read/write.|
| [Parent](tabstop.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Position](tabstop.position-property-publisher.md)|Returns or sets a  **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline, negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.|

