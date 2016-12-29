---
title: GroupShapes Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 72e9f731-fba3-c8d3-fe5f-3d81d59321b8
ms.locale: en-US
---


# GroupShapes Members (Publisher)
Represents the individual shapes within a grouped shape. Each shape is represented by a  **[Shape](shape-object-publisher.md)** object. Using the  **[Item](groupshapes.item-method-publisher.md)** method with this object, you can work with single shapes within a group without having to ungroup them.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Item](groupshapes.item-method-publisher.md)|Returns an individual object in a specified collection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](groupshapes.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](groupshapes.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](groupshapes.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

