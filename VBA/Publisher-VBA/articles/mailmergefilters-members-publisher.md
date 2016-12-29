---
title: MailMergeFilters Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: c700d39b-d940-6c7b-b423-c11ebc43734f
ms.locale: en-US
---


# MailMergeFilters Members (Publisher)
Represents all the filters to apply to the data source attached to the mail merge or catalog merge publication. The  **MailMergeFilters** object is composed of **MailMergeFilterCriterion** objects.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Add](mailmergefilters.add-method-publisher.md)|Adds a new filter criterion to the specified  **MailMergeFilters** object.|
| [Delete](mailmergefilters.delete-method-publisher.md)|Deletes the specified object.|
| [Item](mailmergefilters.item-method-publisher.md)|Returns an individual object in a specified collection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](mailmergefilters.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](mailmergefilters.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Creator](mailmergefilters.creator-property-publisher.md)|Returns a  **Long** that represents the application in which the specified object was created. For example, if the object was created in Microsoft Publisher, this property returns the hexadecimal number 4D505542, which represents the string "MSPB." This value can also be represented by the constant.|
| [Parent](mailmergefilters.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

