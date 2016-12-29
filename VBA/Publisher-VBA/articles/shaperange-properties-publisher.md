---
title: ShapeRange Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: ed51c793-07f2-48ea-92f6-75f4013c5265
ms.locale: en-US
---


# ShapeRange Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Adjustments](shaperange.adjustments-property-publisher.md)|Returns an  **[Adjustments](adjustments-object-publisher.md)** collection representing all adjustment handles for the specified **Shape** or **ShapeRange** object.|
| [AlternativeText](shaperange.alternativetext-property-publisher.md)|Returns or sets a  **String** representing the text displayed by a Web browser in place of the **Shape** object while the **Shape** object is being downloaded or when graphics are turned off. Read/write.|
| [Application](shaperange.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoShapeType](shaperange.autoshapetype-property-publisher.md)|Returns or sets an  **MsoAutoShapeType**constant that specifies a  **ShapeRange** object's AutoShape type.|
| [BlackWhiteMode](shaperange.blackwhitemode-property-publisher.md)|Returns or sets an  **MsoBlackWhiteMode**constant indicating how the specified shape or shape range appears when the publication is viewed in black-and-white mode. Read/write.|
| [Callout](shaperange.callout-property-publisher.md)|Returns a  **[CalloutFormat](calloutformat-object-publisher.md)** object representing the formatting of a line callout.|
| [ConnectionSiteCount](shaperange.connectionsitecount-property-publisher.md)|Returns a  **Long** indicating the count of connection sites on the current **Shape** object. Read-only.|
| [Connector](shaperange.connector-property-publisher.md)|Returns an  **MsoTriState**value indicating whether the specified shape is a connector. Read-only.|
| [ConnectorFormat](shaperange.connectorformat-property-publisher.md)|Returns a  **[ConnectorFormat](connectorformat-object-publisher.md)** object that contains connector formatting properties. Applies to  **Shape** or **ShapeRange** objects that represent connectors.|
| [Count](shaperange.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Fill](shaperange.fill-property-publisher.md)| Returns a **[FillFormat](fillformat-object-publisher.md)** object representing the fill for the specified shape or table cell.|
| [Glow](shaperange.glow-property-publisher.md)|Returns a  **[GlowFormat](glowformat-object-publisher.md)** object that represents the glow formatting for a range of shapes. Read-only.|
| [GroupItems](shaperange.groupitems-property-publisher.md)|Returns a  **[GroupShapes](groupshapes-object-publisher.md)** collection if the specified shape is a group.|
| [HasTable](shaperange.hastable-property-publisher.md)|Returns  **msoTrue** if the shape represents a **TableFrame** object or **msoFalse** if the shape represents any other object type. Read-only.|
| [HasTextFrame](shaperange.hastextframe-property-publisher.md)|Indicates whether the specified shape has a  **TextFrame** object associated with it. Read-only.|
| [Height](shaperange.height-property-publisher.md)|Returns a  **Variant** that represents the height (in points) of a specified range of shapes. Read-only.|
| [HorizontalFlip](shaperange.horizontalflip-property-publisher.md)|Indicates whether the specified shape has been flipped around its horizontal axis. Read-only.|
| [Hyperlink](shaperange.hyperlink-property-publisher.md)|Returns a  **[Hyperlink](hyperlink-object-publisher.md)** object representing the hyperlink associated with the specified shape.|
| [ID](shaperange.id-property-publisher.md)|Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.|
| [InlineAlignment](shaperange.inlinealignment-property-publisher.md)|Returns or sets a  **PbInlineAlignment** constant that indicates whether an inline shape has left, right, or in-text alignment. Read/write.|
| [InlineTextRange](shaperange.inlinetextrange-property-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that reflects the position of the inline shape in its containing text range. Read-only.|
| [IsInline](shaperange.isinline-property-publisher.md)|Returns an  **MsoTriState** constant that specifies whether a shape is inline. Read-only.|
| [Left](shaperange.left-property-publisher.md)|Returns a  **Variant** indicating the distance from the left edge of the page to the leftmost edge of all the shapes in the specified shape range. Numeric values are in points; all other values are in any measurement supported by Publisher (for example, "2.5 in"). Read-only.|
| [Line](shaperange.line-property-publisher.md)|Returns a  **[LineFormat](lineformat-object-publisher.md)** object that contains line formatting properties for the specified shape. (For a line, the  **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.).|
| [LinkFormat](shaperange.linkformat-property-publisher.md)|Returns a  [LinkFormat](linkformat-object-publisher.md)object that contains the properties that are unique to linked OLE objects. Read-only.|
| [LockAspectRatio](shaperange.lockaspectratio-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether the specified shape retains its original proportions when you resize it. Read/write.|
| [Name](shaperange.name-property-publisher.md)|Returns or sets a  **String** value indicating the name of the specified object. Read/write.|
| [Nodes](shaperange.nodes-property-publisher.md)|Returns a  **[ShapeNodes](shapenodes-object-publisher.md)** collection that represents the geometric description of the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent freeform drawings.|
| [OLEFormat](shaperange.oleformat-property-publisher.md)|Returns an  **[OLEFormat](oleformat-object-publisher.md)** object that contains OLE formatting properties for the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent OLE objects.|
| [Parent](shaperange.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [PictureFormat](shaperange.pictureformat-property-publisher.md)|Returns a  **[PictureFormat](pictureformat-object-publisher.md)** object that contains picture formatting properties for the specified object. Applies to  **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** objects that represent pictures or OLE objects. Read-only.|
| [Reflection](shaperange.reflection-property-publisher.md)|Returns a  **[ReflectionFormat](reflectionformat-object-publisher.md)** object that represents the reflection formatting for a range of shapes. Read-only.|
| [Rotation](shaperange.rotation-property-publisher.md)|Returns or sets a  **Single** that represents the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write.|
| [Shadow](shaperange.shadow-property-publisher.md)|Returns a  **[ShadowFormat](shadowformat-object-publisher.md)** object that represents the shadow formatting for the specified shape.|
| [SoftEdge](shaperange.softedge-property-publisher.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-publisher.md)** object that represents the soft edge formatting for a range of shapes. Read-only.|
| [Table](shaperange.table-property-publisher.md)|Returns a  **Table** object that represents a table in Microsoft Publisher.|
| [Tags](shaperange.tags-property-publisher.md)|Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.|
| [TextEffect](shaperange.texteffect-property-publisher.md)|Returns a  **[TextEffectFormat](texteffectformat-object-publisher.md)** object that represents the text formatting properties of a WordArt object.|
| [TextFrame](shaperange.textframe-property-publisher.md)|Returns a  **[TextFrame](textframe-object-publisher.md)** object that represents the text in a shape and the properties that control the margins and orientation of the text.|
| [TextWrap](shaperange.textwrap-property-publisher.md)|Returns a  **[WrapFormat](wrapformat-object-publisher.md)** object that represents the properties for wrapping text around a shape or shape range.|
| [ThreeD](shaperange.threed-property-publisher.md)|Returns a  **[ThreeDFormat](threedformat-object-publisher.md)** object.|
| [Top](shaperange.top-property-publisher.md)|Returns a  **Variant** that represents the distance between the top of the page and the top shape in a range of shapes. Read-only.|
| [Type](shaperange.type-property-publisher.md)|Specifies the shape type. Read-only.|
| [VerticalFlip](shaperange.verticalflip-property-publisher.md)|Returns  **msoTrue** if the specified shape has been flipped around its vertical axis. Read-only.|
| [Vertices](shaperange.vertices-property-publisher.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. Read-only  **Variant**.|
| [Width](shaperange.width-property-publisher.md)|Returns a  **Variant** that represents the width (in points) of a specified range of shapes. Read-only.|
| [Wizard](shaperange.wizard-property-publisher.md)|Returns a  **[Wizard](wizard-object-publisher.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.|
| [WizardTag](shaperange.wizardtag-property-publisher.md)|Returns or sets a  **PbWizardTag**constant indicating the function of a specified shape with respect to its publication design. Read/write.|
| [WizardTagInstance](shaperange.wizardtaginstance-property-publisher.md)|Returns or sets a  **Long** indicating the instance of the specified shape compared with other shapes having the same wizard tag. Read/write.|
| [ZOrderPosition](shaperange.zorderposition-property-publisher.md)|Returns a  **Long** indicating the position of the specified shape or shape range in the z-order. Read-only.|

