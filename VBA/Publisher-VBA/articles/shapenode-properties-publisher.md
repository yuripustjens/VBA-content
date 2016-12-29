---
title: ShapeNode Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 4e76d72a-9fdf-4b98-b6f5-d0610b4cce9a
ms.locale: en-US
---


# ShapeNode Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](shapenode.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [EditingType](shapenode.editingtype-property-publisher.md)|If the specified node is a vertex, this property returns an  **MsoEditingType** constant indicating how changes made to the node affect the two segments connected to the node. If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. Read-only.|
| [Parent](shapenode.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Points](shapenode.points-property-publisher.md)|Gets the  _x-_ and _y-_ coordinates of the shape node. Read-only.|
| [SegmentType](shapenode.segmenttype-property-publisher.md)|Returns an  **MsoSegmentType** constant that indicates whether the segment associated with the specified node is straight or curved. Read-only.|
|Name|Description|

