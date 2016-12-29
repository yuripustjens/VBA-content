---
title: ColorFormat Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 2c5305de-e49a-e9c5-1371-5809b02d3797
ms.locale: en-US
---


# ColorFormat Members (Publisher)
Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill. You can set colors to an explicit red-green-blue value by using the  **[RGB](colorformat.rgb-property-publisher.md)** property.

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](colorformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BaseCMYK](colorformat.basecmyk-property-publisher.md)|Returns the base cyan-magenta-yellow-black (CMYK) color value of the parent  **ColorFormat** object before any tinting or shading is applied to the color. Read-only.|
| [BaseRGB](colorformat.basergb-property-publisher.md)|Returns or sets an  **MsoRGBType** constant that represents the original RGB color format before color-changing properties, such as tinting and shading, are applied. Read/write.|
| [CMYK](colorformat.cmyk-property-publisher.md)|Returns a  **ColorCMYK** object that represents CMYK color properties.|
| [Ink](colorformat.ink-property-publisher.md)|Returns or sets a  **Long** indicating whether the specified color is a spot color, and if so, the spot plate to which it belongs. Valid values are **pbInkNone** (default; meaning that the color is not a spot color) or a number between 1 and _n_ where _n_ is the number of spot plates. Read/write.|
| [Parent](colorformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [RGB](colorformat.rgb-property-publisher.md)|Returns or sets an  **MsoRGBType** that represents the red-green-blue (RGB) value of the specified color. Read/write.|
| [SchemeColor](colorformat.schemecolor-property-publisher.md)|Specifies the color of the current color scheme. Read/write.|
| [TintAndShade](colorformat.tintandshade-property-publisher.md)|Returns or sets a  **Single** that represents the lightening or darkening of a specified shape's color. Read/write.|
| [Transparency](colorformat.transparency-property-publisher.md)|Gets or sets the degree of transparency of the color represented by the parent  **ColorFormat** object as a value between 0.0 (opaque) and 1.0 (clear). Read/write.|
| [Type](colorformat.type-property-publisher.md)|Returns or sets a  **PbColorType** constant that represents the shape color type. Read-only.|
|Name|Description|

