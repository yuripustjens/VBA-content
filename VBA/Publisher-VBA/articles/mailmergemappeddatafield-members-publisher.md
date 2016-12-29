---
title: MailMergeMappedDataField Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: fae683c3-3bee-fc67-1f17-8c104d8c1ef6
ms.locale: en-US
---


# MailMergeMappedDataField Members (Publisher)
Represents a single mapped data field. The  **MailMergeMappedDataField** object is a member of the **[MailMergeMappedDataFields](mailmergemappeddatafields-object-publisher.md)** collection. A mapped data field is a field contained within Microsoft Publisher that represents commonly used name or address information, such as First Name. If a data source contains a First Name field or a variation (such as First_Name, FirstName, First, or FName), the field in the data source will automatically map to the corresponding mapped data field. If a publication is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the publication to agree with the field names in the database.

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](mailmergemappeddatafield.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [DataFieldName](mailmergemappeddatafield.datafieldname-property-publisher.md)|Returns or sets a  **String** which represents the name of the field in the mail merge data source to which a mapped data field maps. An empty string is returned if the specified data field is not mapped to a mapped data field. Read/write.|
| [Index](mailmergemappeddatafield.index-property-publisher.md)|Returns a  **Long** that represents the position of a particular item in a specified collection. .|
| [Name](mailmergemappeddatafield.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](mailmergemappeddatafield.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Value](mailmergemappeddatafield.value-property-publisher.md)|Returns a  **String** that represents the value of a mail merge data field record or a mapped data field. Read-only.|

