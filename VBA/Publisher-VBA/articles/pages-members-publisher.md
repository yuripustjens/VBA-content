---
title: Pages Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 149036ce-8686-da05-d4f8-f89e6bd78570
ms.locale: en-US
---


# Pages Members (Publisher)
Represents all the pages in a publication. The  **Pages** collection contains all the **[Page](page-object-publisher.md)** objects in a publication.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Add](pages.add-method-publisher.md)|Adds a new  **Page** object to the specified **Pages** object and returns the new **Page** object.|
| [AddWizardPage](pages.addwizardpage-method-publisher.md)|Adds the specified new wizard page to a specified location in a publication.|
| [FindByPageID](pages.findbypageid-method-publisher.md)|Returns a  **[Page](page-object-publisher.md)** object that represents the page with the specified page ID number. Each page is automatically assigned a unique ID number when it is created. Use the **[PageID](page.pageid-property-publisher.md)** property to return a page's ID number.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](pages.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](pages.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Item](pages.item-property-publisher.md)|Returns the specified  **[Page](page-object-publisher.md)** object from a **Pages** or **MasterPages** collection. Read-only.|
| [Parent](pages.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

