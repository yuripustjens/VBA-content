---
title: ColorScheme Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 46c2e0a5-3ea3-c15c-7419-036c5aadff1b
ms.locale: en-US
---


# ColorScheme Members (Publisher)
Represents a color scheme, which is a set of eight colors used for the different elements of a publication. Each color is represented by a  **[ColorFormat](colorformat-object-publisher.md)** object. The  **ColorScheme** object is a member of the **[ColorSchemes](colorschemes-object-publisher.md)** collection. The  **ColorSchemes** collection contains all the color schemes available to Microsoft Publisher.

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](colorscheme.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Colors](colorscheme.colors-property-publisher.md)|Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing a color from the specified color scheme.|
| [Name](colorscheme.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](colorscheme.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

