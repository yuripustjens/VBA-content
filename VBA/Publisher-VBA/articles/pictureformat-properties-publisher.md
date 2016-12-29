---
title: PictureFormat Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 4a41da55-3f9d-45eb-b6d6-20e43eab98ec
ms.locale: en-US
---


# PictureFormat Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](pictureformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Brightness](pictureformat.brightness-property-publisher.md)|Returns or sets a  **Single** indicating the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write.|
| [ColorModel](pictureformat.colormodel-property-publisher.md)|Returns a  **PbColorModel** constant that represents the color model of the picture. Read-only.|
| [ColorsInPalette](pictureformat.colorsinpalette-property-publisher.md)| Returns a **Long** that represents the number of colors in the picture's palette. Read-only.|
| [ColorType](pictureformat.colortype-property-publisher.md)|Returns or sets an  **MsoPictureColorType** constant indicating the type of color transformation applied to the specified picture or OLE object. Read/write.|
| [Contrast](pictureformat.contrast-property-publisher.md)|Returns or sets a  **Single** indicating the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write.|
| [CropBottom](pictureformat.cropbottom-property-publisher.md)|Returns or sets a  **Variant** indicating the amount by which the bottom edge of a picture or OLE object is cropped. Read/write.|
| [CropLeft](pictureformat.cropleft-property-publisher.md)|Returns or sets a  **Variant** indicating the amount by which the left edge of a picture or OLE object is cropped. Read/write.|
| [CropRight](pictureformat.cropright-property-publisher.md)|Returns or sets a  **Variant** indicating the amount by which the right edge of a picture or OLE object is cropped. Read/write.|
| [CropTop](pictureformat.croptop-property-publisher.md)|Returns or sets a  **Variant** indicating the amount by which the top edge of a picture or OLE object is cropped. Read/write.|
| [EffectiveResolution](pictureformat.effectiveresolution-property-publisher.md)|Returns a  **Long** that represents, in dots per inch (dpi), the effective resolution of the picture. Read-only.|
| [Filename](pictureformat.filename-property-publisher.md)|Returns a  **String** that represents the file name of the specified picture or OLE object. Read-only.|
| [FileSize](pictureformat.filesize-property-publisher.md)|Returns a  **Long** that represents, in bytes, the size of the picture or OLE object as it appears in the specified publication. Read-only.|
| [HasAlphaChannel](pictureformat.hasalphachannel-property-publisher.md)|Returns an  **MsoTriState** constant indicating whether the specified picture contains an alpha channel. Read-only.|
| [HasTransparencyColor](pictureformat.hastransparencycolor-property-publisher.md)|Returns a  **Boolean** that indicates whether a transparency color has been applied to the specified picture. Read-only.|
| [Height](pictureformat.height-property-publisher.md)|Returns a  **Variant** that represents the height, in points, of the specified picture or OLE object. Read-only.|
| [HorizontalPictureLocking](pictureformat.horizontalpicturelocking-property-publisher.md)|Returns or sets a  **PbHorizontalPictureLocking** constant indicating where newly inserted pictures appear in relation to the specified frame. Read/write.|
| [HorizontalScale](pictureformat.horizontalscale-property-publisher.md)|Returns a  **Long** that represents the scaling of the picture along its horizontal axis. The scaling is expressed as a percentage (for example, 200 equals 200 percent scaling). Read-only.|
| [ImageFormat](pictureformat.imageformat-property-publisher.md)|Returns a  **PbImageFormat** constant that represents the image format of a picture as determined by Microsoft Windows Graphics Device Interface (GDI+). Read-only.|
| [IsEmpty](pictureformat.isempty-property-publisher.md)|Returns a  **MsoTriState** constant that represents whether the specified shape is an empty picture frame. Read-only.|
| [IsGreyScale](pictureformat.isgreyscale-property-publisher.md)|Returns a  **MsoTriState** constant that indicates whether the picture is a greyscale image. Read-only.|
| [IsLinked](pictureformat.islinked-property-publisher.md)|Returns a  **MsoTriState** constant indicating whether the specified picture is a linked picture or OLE object. Read-only.|
| [IsRecolored](pictureformat.isrecolored-property-publisher.md)|Returns  **True** if the image represented by the parent **PictureFormat** object has been recolored, either in the user interface or by using the ** [PictureFormat.Recolor](pictureformat.recolor-method-publisher.md)** method. Read-only.|
| [IsTrueColor](pictureformat.istruecolor-property-publisher.md)|Returns an  **MsoTriState** constant indicating whether the specified picture or OLE object contains color data of 24 bits per channel or greater. Read-only.|
| [LeaveBlackAsBlack](pictureformat.leaveblackasblack-property-publisher.md)|Returns  **True** if, when you recolor the image represented by the parent **PictureFormat** object, the black parts of the original image should remain black. Read-only.|
| [LinkedFileStatus](pictureformat.linkedfilestatus-property-publisher.md)|Returns a  **PbLinkedFileStatus** constant that indicates the status of the file linked to the specified picture. Read-only.|
| [OriginalColorsInPalette](pictureformat.originalcolorsinpalette-property-publisher.md)|Returns a  **Long** that represents the number of colors in the specified linked picture's palette. Read-only.|
| [OriginalFileSize](pictureformat.originalfilesize-property-publisher.md)|Returns a  **Long** representing the size, in bytes, of the linked picture or OLE object. Read-only.|
| [OriginalHasAlphaChannel](pictureformat.originalhasalphachannel-property-publisher.md)|Returns an  **MsoTriState** constant depending on whether the original, linked picture contains an alpha channel. Read-only.|
| [OriginalHeight](pictureformat.originalheight-property-publisher.md)|Returns a  **Variant** representing the height, in points, of the specified linked picture or OLE object. Read-only.|
| [OriginalIsTrueColor](pictureformat.originalistruecolor-property-publisher.md)|Returns an  **MsoTriState** constant indicating whether the specified linked picture or OLE object contains color data of 24 bits per channel or greater. Read-only.|
| [OriginalResolution](pictureformat.originalresolution-property-publisher.md)|Returns a  **Long** that represents, in dots per inch (dpi), the resolution at which the linked picture was originally scanned. Read-only.|
| [OriginalWidth](pictureformat.originalwidth-property-publisher.md)|Returns a  **Variant** that represents, in points, the width of the specified linked picture or OLE object. Read-only.|
| [Parent](pictureformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [RecoloredPictureColor](pictureformat.recoloredpicturecolor-property-publisher.md)|Returns the color that has been applied to the image represented by the parent  **PictureFormat** object. Read-only.|
| [TransparencyColor](pictureformat.transparencycolor-property-publisher.md)|Returns or sets an  **MsoRGBType** constant that represents the transparency color. Read/write.|
| [TransparentBackground](pictureformat.transparentbackground-property-publisher.md)|Indicates whether the parts of the specified picture that are defined as the transparent color appear transparent. Read/write.|
| [VerticalPictureLocking](pictureformat.verticalpicturelocking-property-publisher.md)|Returns or sets a  **PbVerticalPictureLocking** constant indicating where newly inserted pictures appear in relation to the specified frame. Read/write.|
| [VerticalScale](pictureformat.verticalscale-property-publisher.md)|Returns a  **Long** that represents the scaling of the picture along its vertical axis. The scaling is expressed as a percentage (for example, 200 equals 200 percent scaling). Read-only.|
| [Width](pictureformat.width-property-publisher.md)|Returns a  **Variant** that represents the width, in points, of the specified picture. Read-only.|
|Name|Description|

