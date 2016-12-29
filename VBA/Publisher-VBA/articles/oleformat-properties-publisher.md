---
title: OLEFormat Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: f857dd85-086e-4312-97b3-a34a27600c1f
ms.locale: en-US
---


# OLEFormat Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](oleformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Object](oleformat.object-property-publisher.md)|Returns an  **Object** that represents the specified OLE object's top-level interface. This property allows you to access the properties and methods of an ActiveX control or the application in which an OLE object was created. The OLE object must support OLE Automation for this property to work. Read-only.|
| [ObjectVerbs](oleformat.objectverbs-property-publisher.md)|Returns an  **[ObjectVerbs](objectverbs-object-publisher.md)** collection that contains all the OLE verbs for the specified OLE object. Read-only.|
| [Parent](oleformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ProgId](oleformat.progid-property-publisher.md)|Returns a  **String** that represents the programmatic identifier (ProgID) for the specified OLE object. Read-only.|
|Name|Description|

