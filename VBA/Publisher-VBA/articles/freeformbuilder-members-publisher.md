---
title: FreeformBuilder Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: ef008b58-823e-28a4-8e50-c19be6deafb3
ms.locale: en-US
---


# FreeformBuilder Members (Publisher)
Represents the geometry of a freeform while it is being built.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddNodes](freeformbuilder.addnodes-method-publisher.md)|Inserts a new segment at the end of the freeform that is being created, and adds the nodes that define the segment. You can use this method as many times as you want to add nodes to the freeform you are creating. When you finish adding nodes, use the  **[ConvertToShape](freeformbuilder.converttoshape-method-publisher.md)** method to create the freeform you just defined.|
| [ConvertToShape](freeformbuilder.converttoshape-method-publisher.md)|Creates a shape that has the geometric characteristics of the specified  **[FreeformBuilder](freeformbuilder-object-publisher.md)** object. Returns a **[Shape](shape-object-publisher.md)** object that represents the new shape.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](freeformbuilder.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Parent](freeformbuilder.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

