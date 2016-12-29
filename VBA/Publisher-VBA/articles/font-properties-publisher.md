---
title: Font Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 44b90349-ac17-44d0-8326-8cb78b169a47
ms.locale: en-US
---


# Font Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [AllCaps](font.allcaps-property-publisher.md)|Returns or sets  **msoTrue** if the font is formatted as all capital letters, or returns one of the other **MsoTriState** constants if it is not. Read/write.|
| [Application](font.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AttachedToText](font.attachedtotext-property-publisher.md)| **True** if the **Font** or **ParagraphFormat** object is attached to a **TextRange** object. If the object is attached to a **TextRange** object, the document will be updated when properties of the object are changed. If the object is not attached, nothing in the document will be changed until the object is applied to a **TextRange** or **Style** object. Read-only **Boolean**.|
| [AutomaticPairKerningThreshold](font.automaticpairkerningthreshold-property-publisher.md)|Returns or sets a  **Variant** value that represents the point size above which kerning is automatically adjusted for characters in the specified text range. Read/write.|
| [Bold](font.bold-property-publisher.md)|Returns or sets an  **MsoTriState**constant that represents the state of the  **Bold** property on the characters in a text range. Read/write.|
| [BoldBi](font.boldbi-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether the font is bold; used with text in a right-to-left language. Read/write.|
| [ContextualAlternates](font.contextualalternates-property-publisher.md)|Returns or sets an  **MsoTriState** constant that represents the state of the **ContextualAlternates** property on the characters in a text range. The **ContextualAlternates** property enables different shape choices for some characters depending on the context of the character and the design of the selected font. Read/write.|
| [DiacriticColor](font.diacriticcolor-property-publisher.md)|Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the 24-bit color used for diacritics in a right-to-left language publication.|
| [ExpandUsingKashida](font.expandusingkashida-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether to apply kashida rules while applying tracking to Arabic text. Read/write.|
| [Fill](font.fill-property-publisher.md)|Returns a  [FillFormat](fillformat-object-publisher.md) object that contains fill formatting properties for the font used by the specified range of text. Read-only.|
| [Glow](font.glow-property-publisher.md)|Returns a  [GlowFormat](glowformat-object-publisher.md) object that represents the glow formatting for the font used by the specified range of text. Read-only.|
| [Italic](font.italic-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether the specified text is formatted as italic. Read/write.|
| [ItalicBi](font.italicbi-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether the specified text is formatted as italic; applies to text in a right-to-left language. Read/write.|
| [Kerning](font.kerning-property-publisher.md)|Returns or sets a  **Variant** indicating the amount of horizontal spacing Microsoft Publisher applies to the characters in the text range. Read/write.|
| [Ligature](font.ligature-property-publisher.md)|Returns or sets a  **[pbLigaturePresetType](pbligaturepresettype-enumeration-publisher.md)** constant that represents the state of the **Ligature** property on the characters in a text range. The **Ligature** property enables embellishments to the characters, often in the form of bigger and more flamboyant serifs. Read/write.|
| [Line](font.line-property-publisher.md)|Returns a  [LineFormat](lineformat-object-publisher.md) object that specifies the formatting for a line. Read/write.|
| [Name](font.name-property-publisher.md)|Indicates the name of the specified font. Read/write.|
| [NumberStyle](font.numberstyle-property-publisher.md)|Returns or sets a  **[pbNumberStyles](pbnumberstylestype-enumeration-publisher.md)** constant that represents the state of the **NumberStyles** property on the numerical characters in a text range. Read/write.|
| [Parent](font.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Position](font.position-property-publisher.md)|Returns or sets a  **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline, negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.|
| [Reflection](font.reflection-property-publisher.md)|Returns a  [ReflectionFormat](reflectionformat-object-publisher.md) object that represents the reflection formatting for a shape. Read-only.|
| [Scaling](font.scaling-property-publisher.md)|Returns or sets a  **Variant** value used to scale the width of the characters in the text range as a percentage of the current font size. Read/write.|
| [Size](font.size-property-publisher.md)|Represents the size of the characters in the text range in points. Read/write.|
| [SizeBi](font.sizebi-property-publisher.md)|Returns or sets a  **Variant** value representing the size, in points, of the **Font** object for text in a right-to-left language. Valid range is 0.5 points to 999.5 points. Read/write.|
| [SmallCaps](font.smallcaps-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether the specified text is formatted as small caps. Read/write.|
| [StrikeThrough](font.strikethrough-property-publisher.md)|Returns or sets an  **MsoTriState** constant that represents the state of the **StrikeThrough** property on the characters in a text range. Read/write.|
| [StylisticAlternates](font.stylisticalternates-property-publisher.md)|Returns or sets a  **Variant** that represents the state of the **StylisticAlternates** property on the characters in a text range. The **StylisticAlternates** property allows you to select an alternate look for the look of the characters you have selected, if the font designer has created these alternates. Read/write.|
| [StylisticSets](font.stylisticsets-property-publisher.md)|Returns or sets a  **Variant** that represents the state of the **StylisticSets** property on the characters in a text range. Read/write.|
| [SubScript](font.subscript-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether characters are formatted as subscript in the specified text range. Read/write.|
| [SuperScript](font.superscript-property-publisher.md)|Returns or sets an  **MsoTriState** constant indicating whether characters are formatted as superscript in the specified text range. Read/write.|
| [Swash](font.swash-property-publisher.md)|Returns or sets an  **MsoTriState** constant that represents the state of the **Swash** property on the characters in a text range. The **Swash** property enables embellishments to the characters, often in the form of bigger and more flamboyant serifs. Read/write.|
| [TextShadow](font.textshadow-property-publisher.md)|Returns a  [ShadowFormat](shadowformat-object-publisher.md) object that specifies the shadow formatting for the specified font.|
| [ThreeD](font.threed-property-publisher.md)|Returns a  [ThreeDFormat](threedformat-object-publisher.md) object that contains 3-D effect formatting properties for the specified font. Read-only.|
| [Tracking](font.tracking-property-publisher.md)|Returns or sets a  **Variant** indicating the tracking value used to display space between the characters in the specified text range. Read/write.|
| [TrackingPreset](font.trackingpreset-property-publisher.md)|Returns or sets a  **PbTrackingPresetType** constant representing the preset tracking type for characters in the specified font in a text range. Read/write.|
| [Underline](font.underline-property-publisher.md)|Returns or sets an  **PbUnderlineType** constant that indicates the type of underline for the selected characters in the specified font in a text range. Read/write.|
| [UseDiacriticColor](font.usediacriticcolor-property-publisher.md)|Returns or sets  **MsoTriState** constant indicating whether you can set the color of diacritics in the specified text range. Read/write.|

