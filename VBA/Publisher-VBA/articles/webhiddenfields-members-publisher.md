---
title: WebHiddenFields Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 69ebb67f-dce8-b708-2c8b-3d395e24dddc
ms.locale: en-US
---


# WebHiddenFields Members (Publisher)
Represents hidden Web fields that allow a Web page to pass non-visible data to the Web server when a Web page is submitted. The  **WebHiddenFields** object enables control of all the hidden fields attached to a Submit command button.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Add](webhiddenfields.add-method-publisher.md)|Adds a new hidden field to a Web form and returns a  **Long** indicating the number of the new field in the **WebHiddenFields** collection. New fields are always placed at the end of the current field list.|
| [Delete](webhiddenfields.delete-method-publisher.md)|Deletes the specified hidden Web field or Web list box item object.|
| [Item](webhiddenfields.item-method-publisher.md)|Returns a  **String** corresponding to the value of a hidden field in a Web form or a list item in a Web list box control.|
| [Name](webhiddenfields.name-method-publisher.md)|Returns a  **String** that represents the name of a hidden Web field for a Web command button.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](webhiddenfields.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](webhiddenfields.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](webhiddenfields.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

