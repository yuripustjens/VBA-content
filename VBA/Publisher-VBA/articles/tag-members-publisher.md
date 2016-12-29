---
title: Tag Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 6847ce7d-746f-d5da-aa16-9371ea4b8bfe
ms.locale: en-US
---


# Tag Members (Publisher)
Represents a tag or a custom property that you can create for a shape, shape range, page, or publication. Each  **Tag** object contains the name of a custom property and a value for that property. **Tag** objects are members of the **[Tags](tags-object-publisher.md)** collection. Create a tag when you want to be able to selectively work with specific members of a collection, based on an attribute that isn't already represented by a built-in property.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Delete](tag.delete-method-publisher.md)|Deletes the specified object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](tag.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Name](tag.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](tag.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Value](tag.value-property-publisher.md)|Returns or sets a  **Variant** that represents the value of a tag of a shape, page, or publication. Read/write.|

