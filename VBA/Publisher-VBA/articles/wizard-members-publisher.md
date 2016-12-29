---
title: Wizard Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: cc935bea-db91-eba8-062e-0be94683d020
ms.locale: en-US
---


# Wizard Members (Publisher)
Represents the publication design associated with a publication or the wizard associated with a Design Gallery object.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [SetId](wizard.setid-method-publisher.md)|Specifies the type of the wizard (template) to which to convert the current publication type.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](wizard.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [ID](wizard.id-property-publisher.md)|Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.|
| [Name](wizard.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](wizard.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Properties](wizard.properties-property-publisher.md)|Returns a  **[WizardProperties](wizardproperties-object-publisher.md)** collection representing all the settings that are part of the specified publication design or Design Gallery object's wizard.|

