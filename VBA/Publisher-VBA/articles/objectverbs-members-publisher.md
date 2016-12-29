---
title: ObjectVerbs Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 8805a9b5-daba-fc4d-7bcb-90216ca49282
ms.locale: en-US
---


# ObjectVerbs Members (Publisher)
Represents the collection of OLE verbs for the specified OLE object. OLE verbs are the operations supported by an OLE object. Commonly used OLE verbs are play and edit.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Item](objectverbs.item-method-publisher.md)|Returns an individual object in a specified collection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](objectverbs.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](objectverbs.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](objectverbs.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

