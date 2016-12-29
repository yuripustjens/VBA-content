---
title: Font Object (Publisher)
keywords: vbapb10.chm5439487
f1_keywords: vbapb10.chm5439487
ms.prod: PUBLISHER
api_name: Publisher.Font
ms.assetid: 992fda94-2820-d665-0d78-efd4b5434731
ms.locale: en-US
---


# Font Object (Publisher)

Contains font attributes (font name, font size, color, and so on) for an object.


## Example

Use the  **[Font](http://msdn.microsoft.com/library/textstyle.font-property-publisher%28Office.15%29.aspx)** property to return the **Font** object. The following instruction applies bold formatting to the selection.


```
Sub BoldText() 
 Selection.TextRange.Font.Bold = True 
End Sub
```

The following example formats the first paragraph in the active publication as 24-point Arial and italic.




```
Sub FormatText() 
 Dim txtRange As TextRange 
 Set txtRange = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 With txtRange.Font 
 .Bold = True 
 .Name = "Arial" 
 .Size = 24 
 End With 
End Sub
```

The following example changes the formatting of the Heading 2 style in the active publication to Arial and bold.




```
Sub FormatStyle() 
 With ActiveDocument.TextStyles("Normal").Font 
 .Name = "Tahoma" 
 .Italic = True 
 .Size = 15 
 End With 
End Sub
```

You can also duplicate a  **Font** object by using the **[Duplicate](http://msdn.microsoft.com/library/textrange.duplicate-property-publisher%28Office.15%29.aspx)** property. The following example creates a new character style with the character formatting from the selection in addition to italic formatting. The formatting of the selection is not changed.




```
Sub DuplicateFont() 
 Dim fntNew As Font 
 Set fntNew = Selection.TextRange.Font.Duplicate 
 fntNew.Italic = True 
 ActiveDocument.TextStyles.Add(StyleName:="Italics").Font = fntNew 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Duplicate](http://msdn.microsoft.com/library/font.duplicate-method-publisher%28Office.15%29.aspx)|
|[GetScriptName](http://msdn.microsoft.com/library/font.getscriptname-method-publisher%28Office.15%29.aspx)|
|[Grow](http://msdn.microsoft.com/library/font.grow-method-publisher%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/font.reset-method-publisher%28Office.15%29.aspx)|
|[SetScriptName](http://msdn.microsoft.com/library/font.setscriptname-method-publisher%28Office.15%29.aspx)|
|[Shrink](http://msdn.microsoft.com/library/font.shrink-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AllCaps](http://msdn.microsoft.com/library/font.allcaps-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/font.application-property-publisher%28Office.15%29.aspx)|
|[AttachedToText](http://msdn.microsoft.com/library/font.attachedtotext-property-publisher%28Office.15%29.aspx)|
|[AutomaticPairKerningThreshold](http://msdn.microsoft.com/library/font.automaticpairkerningthreshold-property-publisher%28Office.15%29.aspx)|
|[Bold](http://msdn.microsoft.com/library/font.bold-property-publisher%28Office.15%29.aspx)|
|[BoldBi](http://msdn.microsoft.com/library/font.boldbi-property-publisher%28Office.15%29.aspx)|
|[ContextualAlternates](http://msdn.microsoft.com/library/font.contextualalternates-property-publisher%28Office.15%29.aspx)|
|[DiacriticColor](http://msdn.microsoft.com/library/font.diacriticcolor-property-publisher%28Office.15%29.aspx)|
|[ExpandUsingKashida](http://msdn.microsoft.com/library/font.expandusingkashida-property-publisher%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/font.fill-property-publisher%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/font.glow-property-publisher%28Office.15%29.aspx)|
|[Italic](http://msdn.microsoft.com/library/font.italic-property-publisher%28Office.15%29.aspx)|
|[ItalicBi](http://msdn.microsoft.com/library/font.italicbi-property-publisher%28Office.15%29.aspx)|
|[Kerning](http://msdn.microsoft.com/library/font.kerning-property-publisher%28Office.15%29.aspx)|
|[Ligature](http://msdn.microsoft.com/library/font.ligature-property-publisher%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/font.line-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/font.name-property-publisher%28Office.15%29.aspx)|
|[NumberStyle](http://msdn.microsoft.com/library/font.numberstyle-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/font.parent-property-publisher%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/font.position-property-publisher%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/font.reflection-property-publisher%28Office.15%29.aspx)|
|[Scaling](http://msdn.microsoft.com/library/font.scaling-property-publisher%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/font.size-property-publisher%28Office.15%29.aspx)|
|[SizeBi](http://msdn.microsoft.com/library/font.sizebi-property-publisher%28Office.15%29.aspx)|
|[SmallCaps](http://msdn.microsoft.com/library/font.smallcaps-property-publisher%28Office.15%29.aspx)|
|[StrikeThrough](http://msdn.microsoft.com/library/font.strikethrough-property-publisher%28Office.15%29.aspx)|
|[StylisticAlternates](http://msdn.microsoft.com/library/font.stylisticalternates-property-publisher%28Office.15%29.aspx)|
|[StylisticSets](http://msdn.microsoft.com/library/font.stylisticsets-property-publisher%28Office.15%29.aspx)|
|[SubScript](http://msdn.microsoft.com/library/font.subscript-property-publisher%28Office.15%29.aspx)|
|[SuperScript](http://msdn.microsoft.com/library/font.superscript-property-publisher%28Office.15%29.aspx)|
|[Swash](http://msdn.microsoft.com/library/font.swash-property-publisher%28Office.15%29.aspx)|
|[TextShadow](http://msdn.microsoft.com/library/font.textshadow-property-publisher%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/font.threed-property-publisher%28Office.15%29.aspx)|
|[Tracking](http://msdn.microsoft.com/library/font.tracking-property-publisher%28Office.15%29.aspx)|
|[TrackingPreset](http://msdn.microsoft.com/library/font.trackingpreset-property-publisher%28Office.15%29.aspx)|
|[Underline](http://msdn.microsoft.com/library/font.underline-property-publisher%28Office.15%29.aspx)|
|[UseDiacriticColor](http://msdn.microsoft.com/library/font.usediacriticcolor-property-publisher%28Office.15%29.aspx)|

