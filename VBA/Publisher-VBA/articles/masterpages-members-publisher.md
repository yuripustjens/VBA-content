---
title: MasterPages Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 380658ad-3326-075b-cddc-95a54a927dda
ms.locale: en-US
---


# MasterPages Members (Publisher)
Represents the page master for a publication after which all pages in the publication will be designed. The  **MasterPages** object is a collection of **[Page](page-object-publisher.md)** objects.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Add](masterpages.add-method-publisher.md)|Adds a new  **Page** object to the specified **MasterPages** object and returns the new **Page** object.|
| [FindByPageID](masterpages.findbypageid-method-publisher.md)|Returns a  **[Page](page-object-publisher.md)** object that represents the page with the specified page ID number. Each page is automatically assigned a unique ID number when it is created. Use the **[PageID](page.pageid-property-publisher.md)** property to return a page's ID number.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](masterpages.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](masterpages.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Item](masterpages.item-property-publisher.md)|Returns the specified  **[Page](page-object-publisher.md)** object from a **Pages** or **MasterPages** collection. Read-only.|
| [Parent](masterpages.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

