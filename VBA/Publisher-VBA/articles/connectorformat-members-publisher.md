---
title: ConnectorFormat Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: afc6b3bc-213a-48d3-668e-876102cfb3fc
ms.locale: en-US
---


# ConnectorFormat Members (Publisher)
Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [BeginConnect](connectorformat.beginconnect-method-publisher.md)|Attaches the beginning of the specified connector to a specified shape.|
| [BeginDisconnect](connectorformat.begindisconnect-method-publisher.md)|Detaches the beginning of the specified connector from the shape to which it is attached.|
| [EndConnect](connectorformat.endconnect-method-publisher.md)|Attaches the end of the specified connector to a specified shape.|
| [EndDisconnect](connectorformat.enddisconnect-method-publisher.md)|Detaches the end of the specified connector from the shape to which it is attached.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](connectorformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BeginConnected](connectorformat.beginconnected-property-publisher.md)|Returns an  **MsoTriState**constant indicating whether the beginning of the specified connector is connected to a shape. Read-only.|
| [BeginConnectedShape](connectorformat.beginconnectedshape-property-publisher.md)|Returns a  **[Shape](shape-object-publisher.md)** object that represents the shape to which the beginning of the specified connector is attached.|
| [BeginConnectionSite](connectorformat.beginconnectionsite-property-publisher.md)|Returns a  **Long** indicating the connection site to which the beginning of a connector is connected. Read-only.|
| [EndConnected](connectorformat.endconnected-property-publisher.md)|Returns an  **MsoTriState** constant indicating whether the end of the specified connector is connected to a shape. Read-only.|
| [EndConnectedShape](connectorformat.endconnectedshape-property-publisher.md)|Returns a  **[Shape](shape-object-publisher.md)** object that represents the shape to which the end of the specified connector is attached.|
| [EndConnectionSite](connectorformat.endconnectionsite-property-publisher.md)|Returns a  **Long** indicating the connection site to which the end of a connector is connected. Read-only.|
| [Parent](connectorformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Type](connectorformat.type-property-publisher.md)|Returns or sets an  **MsoConnectorType** constant that represents the connector type. Read/write.|

