---
title: Plate Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: bcff2eb7-98d2-48d2-be7d-cd1ce27e3bd6
ms.locale: en-US
---


# Plate Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](plate.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Color](plate.color-property-publisher.md)|Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the color information for the specified object.|
| [Index](plate.index-property-publisher.md)|Returns a  **Long** that represents the position of a particular item in a specified collection. .|
| [InkName](plate.inkname-property-publisher.md)|Returns a  **PbInkName** constant that represents the name of the ink to be printed using this plate. Read-only.|
| [InUse](plate.inuse-property-publisher.md)|Returns  **True** if the specified ink (represented by the plate) is used in the publication. Read-only **Boolean**.|
| [Luminance](plate.luminance-property-publisher.md)|Returns or sets a  **Long** indicating a calculated luminance value for the specified plate; used for spot-color trapping. Valid values are from 0 to 100. Read/write.|
| [Name](plate.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](plate.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

