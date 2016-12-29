---
title: MailMergeDataField Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: af9decbd-04f0-4f12-80a2-b4edb808aef3
ms.locale: en-US
---


# MailMergeDataField Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](mailmergedatafield.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Creator](mailmergedatafield.creator-property-publisher.md)|Returns a  **Long** that represents the application in which the specified object was created. For example, if the object was created in Microsoft Publisher, this property returns the hexadecimal number 4D505542, which represents the string "MSPB." This value can also be represented by the constant.|
| [FieldType](mailmergedatafield.fieldtype-property-publisher.md)||
| [Index](mailmergedatafield.index-property-publisher.md)|Returns a  **Long** that represents the position of a particular item in a specified collection. .|
| [IsMapped](mailmergedatafield.ismapped-property-publisher.md)|Indicates if the parent  **MailMergeDataField** object is mapped to a recipient field in the master data source (combined mail-merge recipient list). Read-only.|
| [MappedTo](mailmergedatafield.mappedto-property-publisher.md)|Returns the name of the recipient field (column) in the master data source (combined mail-merge recipient list) that the parent  **MailMergeDataField** object is mapped to. Read-only.|
| [Name](mailmergedatafield.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](mailmergedatafield.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Value](mailmergedatafield.value-property-publisher.md)|Returns a  **String** that represents the value of a mail merge data field record or a mapped data field. Read-only.|

