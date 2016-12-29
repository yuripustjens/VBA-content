---
title: MailMergeFilterCriterion Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 6e17db17-85fc-2e92-1b95-e25656e26b84
ms.locale: en-US
---


# MailMergeFilterCriterion Members (Publisher)
Represents a filter to be applied to an attached mail merge or catalog merge data source. The  **MailMergeFilterCriterion** object is a member of the **MailMergeFilters** object.

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](mailmergefiltercriterion.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Column](mailmergefiltercriterion.column-property-publisher.md)|Returns a  **String** that represents the name of the field in the mail merge data source to use in the filter. Read/write.|
| [CompareTo](mailmergefiltercriterion.compareto-property-publisher.md)|Returns or sets a  **String** that represents the text to compare in the query filter criterion. Read/write.|
| [Comparison](mailmergefiltercriterion.comparison-property-publisher.md)|Returns or sets an  **MsoFilterComparison** constant that represents how to compare the [Column](cell.column-property-publisher.md) and **[CompareTo](mailmergefiltercriterion.compareto-property-publisher.md)** properties. Read/write.|
| [Conjunction](mailmergefiltercriterion.conjunction-property-publisher.md)|Returns or sets an  **MsoFilterConjunction** constant that represents how a filter criterion relates to other filter criteria in the **[MailMergeFilters](mailmergefilters-object-publisher.md)** object. Read/write.|
| [Creator](mailmergefiltercriterion.creator-property-publisher.md)|Returns a  **Long** that represents the application in which the specified object was created. For example, if the object was created in Microsoft Publisher, this property returns the hexadecimal number 4D505542, which represents the string "MSPB." This value can also be represented by the constant.|
| [Index](mailmergefiltercriterion.index-property-publisher.md)|Returns a  **Long** that represents the position of a particular item in a specified collection. .|
| [Parent](mailmergefiltercriterion.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
|Name|Description|

