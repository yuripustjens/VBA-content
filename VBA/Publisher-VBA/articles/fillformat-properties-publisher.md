---
title: FillFormat Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 7080f2f9-b5d6-4439-8052-1194fb7ddbbd
ms.locale: en-US
---


# FillFormat Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](fillformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BackColor](fillformat.backcolor-property-publisher.md)|Returns or sets a  **[ColorFormat](colorformat-object-publisher.md)** object representing the background color for the specified fill or patterned line. .|
| [ForeColor](fillformat.forecolor-property-publisher.md)|Returns or sets a  **[ColorFormat](colorformat-object-publisher.md)** object representing the foreground color for the fill, line, or shadow. Read/write.|
| [GradientAngle](fillformat.gradientangle-property-publisher.md)|Returns or sets the angle of the gradient fill for the specified fill format. Read/write.|
| [GradientColorType](fillformat.gradientcolortype-property-publisher.md)|Returns an  **MsoGradientColorType** constant indicating the gradient color type for the specified fill. Read-only.|
| [GradientDegree](fillformat.gradientdegree-property-publisher.md)|Returns a  **Single** indicating how dark or light a one-color gradient fill is. A value of 0 (zero) means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in; and values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in. Read-only.|
| [GradientStyle](fillformat.gradientstyle-property-publisher.md)|Returns an  **MsoGradientStyle** constant indicating the gradient style for the specified fill. Read-only.|
| [GradientVariant](fillformat.gradientvariant-property-publisher.md)|Returns a  **Long** indicating the gradient variant for the specified fill. Generally, values are integers from 1 to 4 for most gradient fills. If the gradient style is **msoGradientFromTitle** or **msoGradientFromCenter**, this property returns either 1 or 2. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the  **Gradient** tab in the **Fill Effects** dialog box. Read-only.|
| [Parent](fillformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Pattern](fillformat.pattern-property-publisher.md)|Returns an  **MsoPatternType** constant that represents the pattern applied to the specified fill or line.|
| [PresetGradientType](fillformat.presetgradienttype-property-publisher.md)|Returns an  **MsoPresetGradientType** constant that represents the preset gradient type for the specified fill. Read-only.|
| [PresetTexture](fillformat.presettexture-property-publisher.md)|Returns an  **MsoPresetTexture** constant that represents the preset texture for the specified fill. Read-only.|
| [RotateWithObject](fillformat.rotatewithobject-property-publisher.md)|Returns or sets whether the fill rotates with the specified shape. Read/write.|
| [TextureAlignment](fillformat.texturealignment-property-publisher.md)|Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.|
| [TextureHorizontalScale](fillformat.texturehorizontalscale-property-publisher.md)|Returns or sets a  **Single** that specifies the horizontal scaling factor for the texture fill. Read/write.|
| [TextureName](fillformat.texturename-property-publisher.md)|Returns a  **String** indicating the name of the custom texture file for the specified fill. Read-only.|
| [TextureOffsetX](fillformat.textureoffsetx-property-publisher.md)|Returns or sets a  **Long** that specifies the horizontal offset of the texture from the origin in points. Read/write.|
| [TextureOffsetY](fillformat.textureoffsety-property-publisher.md)|Returns or sets a  **Long** that specifies the vertical offset of the texture from the origin in points. Read/write.|
| [TextureType](fillformat.texturetype-property-publisher.md)|Returns an  **MsoTextureType** constant indicating the texture type for the specified fill. Read-only.|
| [TextureVerticalScale](fillformat.textureverticalscale-property-publisher.md)|Returns or sets a  **Single** that specifies the vertical scaling factor for the texture fill. Read/write.|
| [Transparency](fillformat.transparency-property-publisher.md)|Returns or sets the degree of transparency of the specified fill for a shape as a value between 0.0 (opaque) and 1.0 (clear). Read/write.|
| [Type](fillformat.type-property-publisher.md)|Returns an  **MsoFillType** constant that represents the fill format type of a shape. Read-only.|
| [Visible](fillformat.visible-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether the specified object or the formatting applied to the specified object is visible. Read/write.|
|Name|Description|

