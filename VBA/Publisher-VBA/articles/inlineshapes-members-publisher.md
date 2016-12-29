---
title: InlineShapes Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 64d6fb5e-2a65-a57b-55c3-e32f3a58f281
ms.locale: en-US
---


# InlineShapes Members (Publisher)
Contains a collection of  **[Shape](shape-object-publisher.md)** objects, which represent objects in the drawing layer, where **Shape.IsInline** is **True**. The collection of shapes is limited to shapes within a given text range.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Item](inlineshapes.item-method-publisher.md)|Returns a  **[Shape](shape-object-publisher.md)** object that represents an inline shape contained in a text range. This method is the default member of the **InlineShapes** collection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](inlineshapes.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](inlineshapes.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](inlineshapes.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Range](inlineshapes.range-property-publisher.md)|Returns a  **[ShapeRange](shaperange-object-publisher.md)** collection that represents the same set of inline shapes as the **InlineShapes** collection whose method was called. This allows for miscellaneous formatting of the contained shapes. Read-only.|

