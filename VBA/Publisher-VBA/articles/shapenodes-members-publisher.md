---
title: ShapeNodes Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: fe25740b-f4e0-0e10-803b-564dba7c86b8
ms.locale: en-US
---


# ShapeNodes Members (Publisher)
A collection of all the  **[ShapeNode](shapenode-object-publisher.md)** objects in the specified freeform. Each  **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the **[BuildFreeform](shapes.buildfreeform-method-publisher.md)** and  **[ConvertToShape](freeformbuilder.converttoshape-method-publisher.md)** methods.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Delete](shapenodes.delete-method-publisher.md)|Deletes the specified shape node object.|
| [Insert](shapenodes.insert-method-publisher.md)|Inserts a new segment after the specified node of the freeform drawing.|
| [Item](shapenodes.item-method-publisher.md)|Returns an individual object in a specified collection.|
| [SetEditingType](shapenodes.seteditingtype-method-publisher.md)|Sets the editing type of the specified node. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Depending on the editing type, this method may affect the position of adjacent nodes.|
| [SetPosition](shapenodes.setposition-method-publisher.md)|Sets the position of the specified node. Depending on the editing type of the node, this method may affect the position of adjacent nodes.|
| [SetSegmentType](shapenodes.setsegmenttype-method-publisher.md)|Sets the segment type of the segment that follows the specified node. If the node is a control point for a curved segment, this method sets the segment type for that curve; this may affect the total number of nodes by inserting or deleting adjacent nodes.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](shapenodes.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](shapenodes.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](shapenodes.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

