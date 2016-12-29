---
title: OLEFormat Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 8c7b8865-fa1e-4b73-6404-393181612e12
ms.locale: en-US
---


# OLEFormat Members (Publisher)
Represents the OLE characteristics, other than linking (see the  **[LinkFormat](linkformat-object-publisher.md)** object), for an OLE object, ActiveX control, or field.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Activate](oleformat.activate-method-publisher.md)|Activates a window or OLE object.|
| [DoVerb](oleformat.doverb-method-publisher.md)|Requests that an OLE object perform one of its verbs.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](oleformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Object](oleformat.object-property-publisher.md)|Returns an  **Object** that represents the specified OLE object's top-level interface. This property allows you to access the properties and methods of an ActiveX control or the application in which an OLE object was created. The OLE object must support OLE Automation for this property to work. Read-only.|
| [ObjectVerbs](oleformat.objectverbs-property-publisher.md)|Returns an  **[ObjectVerbs](objectverbs-object-publisher.md)** collection that contains all the OLE verbs for the specified OLE object. Read-only.|
| [Parent](oleformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ProgId](oleformat.progid-property-publisher.md)|Returns a  **String** that represents the programmatic identifier (ProgID) for the specified OLE object. Read-only.|

