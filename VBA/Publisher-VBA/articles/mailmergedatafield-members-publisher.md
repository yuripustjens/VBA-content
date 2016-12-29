---
title: MailMergeDataField Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 7fbf8b02-5168-b6e6-34c1-6b504da02b81
ms.locale: en-US
---


# MailMergeDataField Members (Publisher)
Represents a single merge field in a data source. The  **MailMergeDataField** object is a member of the **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** collection. The  **MailMergeDataFields** collection includes all the data fields in a mail merge or catalog merge data source (for example, Name, Address, and City).

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddToRecipientFields](mailmergedatafield.addtorecipientfields-method-publisher.md)|Adds the parent  **MailMergeDataField** object from a particular data source to the master data source (collection of data fields) for a mail-merge publication.|
| [Insert](mailmergedatafield.insert-method-publisher.md)|Returns a  **[Shape](shape-object-publisher.md)** object that represents a data field inserted into a publication.|
| [MapToRecipientField](mailmergedatafield.maptorecipientfield-method-publisher.md)|Maps a field (column) in a particular data source represented by the parent  **MailMergeDataField** object to a recipient field (column) in the master data source (combined mail-merge recipient list).|
| [UnMapRecipientField](mailmergedatafield.unmaprecipientfield-method-publisher.md)|Undoes the mapping between the parent  **MailMergeDataField** object in a particular data source and the recipient field in the master data source (combined mail-merge recipient list) to which it is currently mapped.|

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

