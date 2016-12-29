---
title: PictureFormat Object (Publisher)
keywords: vbapb10.chm3670015
f1_keywords: vbapb10.chm3670015
ms.prod: PUBLISHER
api_name: Publisher.PictureFormat
ms.assetid: aa30ea9d-b91f-acdf-2e60-8a9f506f28b4
ms.locale: en-US
---


# PictureFormat Object (Publisher)

Contains properties and methods that apply to pictures.


## Example

Use the  **[PictureFormat](http://msdn.microsoft.com/library/shape.pictureformat-property-publisher%28Office.15%29.aspx)** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Sub FormatPicture() 
 With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .Brightness = 0.6 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayscale 
 .CropBottom = 18 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ClearCrop](http://msdn.microsoft.com/library/pictureformat.clearcrop-method-publisher%28Office.15%29.aspx)|
|[FillFrame](http://msdn.microsoft.com/library/pictureformat.fillframe-method-publisher%28Office.15%29.aspx)|
|[FitFrame](http://msdn.microsoft.com/library/pictureformat.fitframe-method-publisher%28Office.15%29.aspx)|
|[IncrementBrightness](http://msdn.microsoft.com/library/pictureformat.incrementbrightness-method-publisher%28Office.15%29.aspx)|
|[IncrementContrast](http://msdn.microsoft.com/library/pictureformat.incrementcontrast-method-publisher%28Office.15%29.aspx)|
|[Recolor](http://msdn.microsoft.com/library/pictureformat.recolor-method-publisher%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/pictureformat.remove-method-publisher%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/pictureformat.replace-method-publisher%28Office.15%29.aspx)|
|[ReplaceEx](http://msdn.microsoft.com/library/pictureformat.replaceex-method-publisher%28Office.15%29.aspx)|
|[RestoreOriginalColors](http://msdn.microsoft.com/library/pictureformat.restoreoriginalcolors-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pictureformat.application-property-publisher%28Office.15%29.aspx)|
|[Brightness](http://msdn.microsoft.com/library/pictureformat.brightness-property-publisher%28Office.15%29.aspx)|
|[ColorModel](http://msdn.microsoft.com/library/pictureformat.colormodel-property-publisher%28Office.15%29.aspx)|
|[ColorsInPalette](http://msdn.microsoft.com/library/pictureformat.colorsinpalette-property-publisher%28Office.15%29.aspx)|
|[ColorType](http://msdn.microsoft.com/library/pictureformat.colortype-property-publisher%28Office.15%29.aspx)|
|[Contrast](http://msdn.microsoft.com/library/pictureformat.contrast-property-publisher%28Office.15%29.aspx)|
|[CropBottom](http://msdn.microsoft.com/library/pictureformat.cropbottom-property-publisher%28Office.15%29.aspx)|
|[CropLeft](http://msdn.microsoft.com/library/pictureformat.cropleft-property-publisher%28Office.15%29.aspx)|
|[CropRight](http://msdn.microsoft.com/library/pictureformat.cropright-property-publisher%28Office.15%29.aspx)|
|[CropTop](http://msdn.microsoft.com/library/pictureformat.croptop-property-publisher%28Office.15%29.aspx)|
|[EffectiveResolution](http://msdn.microsoft.com/library/pictureformat.effectiveresolution-property-publisher%28Office.15%29.aspx)|
|[Filename](http://msdn.microsoft.com/library/pictureformat.filename-property-publisher%28Office.15%29.aspx)|
|[FileSize](http://msdn.microsoft.com/library/pictureformat.filesize-property-publisher%28Office.15%29.aspx)|
|[HasAlphaChannel](http://msdn.microsoft.com/library/pictureformat.hasalphachannel-property-publisher%28Office.15%29.aspx)|
|[HasTransparencyColor](http://msdn.microsoft.com/library/pictureformat.hastransparencycolor-property-publisher%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/pictureformat.height-property-publisher%28Office.15%29.aspx)|
|[HorizontalPictureLocking](http://msdn.microsoft.com/library/pictureformat.horizontalpicturelocking-property-publisher%28Office.15%29.aspx)|
|[HorizontalScale](http://msdn.microsoft.com/library/pictureformat.horizontalscale-property-publisher%28Office.15%29.aspx)|
|[ImageFormat](http://msdn.microsoft.com/library/pictureformat.imageformat-property-publisher%28Office.15%29.aspx)|
|[IsEmpty](http://msdn.microsoft.com/library/pictureformat.isempty-property-publisher%28Office.15%29.aspx)|
|[IsGreyScale](http://msdn.microsoft.com/library/pictureformat.isgreyscale-property-publisher%28Office.15%29.aspx)|
|[IsLinked](http://msdn.microsoft.com/library/pictureformat.islinked-property-publisher%28Office.15%29.aspx)|
|[IsRecolored](http://msdn.microsoft.com/library/pictureformat.isrecolored-property-publisher%28Office.15%29.aspx)|
|[IsTrueColor](http://msdn.microsoft.com/library/pictureformat.istruecolor-property-publisher%28Office.15%29.aspx)|
|[LeaveBlackAsBlack](http://msdn.microsoft.com/library/pictureformat.leaveblackasblack-property-publisher%28Office.15%29.aspx)|
|[LinkedFileStatus](http://msdn.microsoft.com/library/pictureformat.linkedfilestatus-property-publisher%28Office.15%29.aspx)|
|[OriginalColorsInPalette](http://msdn.microsoft.com/library/pictureformat.originalcolorsinpalette-property-publisher%28Office.15%29.aspx)|
|[OriginalFileSize](http://msdn.microsoft.com/library/pictureformat.originalfilesize-property-publisher%28Office.15%29.aspx)|
|[OriginalHasAlphaChannel](http://msdn.microsoft.com/library/pictureformat.originalhasalphachannel-property-publisher%28Office.15%29.aspx)|
|[OriginalHeight](http://msdn.microsoft.com/library/pictureformat.originalheight-property-publisher%28Office.15%29.aspx)|
|[OriginalIsTrueColor](http://msdn.microsoft.com/library/pictureformat.originalistruecolor-property-publisher%28Office.15%29.aspx)|
|[OriginalResolution](http://msdn.microsoft.com/library/pictureformat.originalresolution-property-publisher%28Office.15%29.aspx)|
|[OriginalWidth](http://msdn.microsoft.com/library/pictureformat.originalwidth-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/pictureformat.parent-property-publisher%28Office.15%29.aspx)|
|[RecoloredPictureColor](http://msdn.microsoft.com/library/pictureformat.recoloredpicturecolor-property-publisher%28Office.15%29.aspx)|
|[TransparencyColor](http://msdn.microsoft.com/library/pictureformat.transparencycolor-property-publisher%28Office.15%29.aspx)|
|[TransparentBackground](http://msdn.microsoft.com/library/pictureformat.transparentbackground-property-publisher%28Office.15%29.aspx)|
|[VerticalPictureLocking](http://msdn.microsoft.com/library/pictureformat.verticalpicturelocking-property-publisher%28Office.15%29.aspx)|
|[VerticalScale](http://msdn.microsoft.com/library/pictureformat.verticalscale-property-publisher%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/pictureformat.width-property-publisher%28Office.15%29.aspx)|

