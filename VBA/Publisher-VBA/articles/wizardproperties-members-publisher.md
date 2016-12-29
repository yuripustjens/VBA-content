---
title: WizardProperties Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 5c6868b8-1bd2-68fb-b23e-d507294c8928
ms.locale: en-US
---


# WizardProperties Members (Publisher)
Represents the settings available in a publication design or in a Design Gallery object's wizard.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [FindPropertyById](wizardproperties.findpropertybyid-method-publisher.md)|Returns a  **[WizardProperty](wizardproperty-object-publisher.md)** object, based on the specified ID, from the collection of wizard properties associated with a publication design or a Design Gallery object's wizard.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](wizardproperties.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Count](wizardproperties.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Item](wizardproperties.item-property-publisher.md)|Returns an individual object from a specified collection. Read-only.|
| [Parent](wizardproperties.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

