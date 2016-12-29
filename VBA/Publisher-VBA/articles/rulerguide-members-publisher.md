---
title: RulerGuide Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 052845fa-a5c5-114d-8fad-39bfe99cffb3
ms.locale: en-US
---


# RulerGuide Members (Publisher)
Represents a gridline used to align objects on a page. The  **RulerGuide** object is a member of the **[RulerGuides](rulerguides-object-publisher.md)** collection.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Delete](rulerguide.delete-method-publisher.md)|Deletes the specified object.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](rulerguide.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Parent](rulerguide.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Position](rulerguide.position-property-publisher.md)|Returns or sets a  **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline, negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.|
| [Type](rulerguide.type-property-publisher.md)|Returns or sets a  **PbRulerGuideType** constant that represents the ruler guide type. Read/write.|

