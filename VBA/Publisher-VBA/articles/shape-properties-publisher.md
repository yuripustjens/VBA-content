---
title: Shape Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 50baf097-df2a-4453-a4db-96191047a50b
ms.locale: en-US
---


# Shape Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Adjustments](shape.adjustments-property-publisher.md)|Returns an  **[Adjustments](adjustments-object-publisher.md)** collection representing all adjustment handles for the specified **Shape** or **ShapeRange** object.|
| [AlternativeText](shape.alternativetext-property-publisher.md)|Returns or sets a  **String** representing the text displayed by a Web browser in place of the **Shape** object while the **Shape** object is being downloaded or when graphics are turned off. Read/write.|
| [Application](shape.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoShapeType](shape.autoshapetype-property-publisher.md)|Returns or sets an  **MsoAutoShapeType**constant that specifies a  **Shape** object's AutoShape type.|
| [BlackWhiteMode](shape.blackwhitemode-property-publisher.md)|Returns or sets an  **MsoBlackWhiteMode**constant indicating how the specified shape or shape range appears when the publication is viewed in black-and-white mode. Read/write.|
| [BorderArt](shape.borderart-property-publisher.md)|Returns a  **[BorderArtFormat](borderartformat-object-publisher.md)** object that represents the BorderArt type applied to the specified shape. Returns "Permission Denied" if BorderArt has not been applied to the shape. Read-only.|
| [Callout](shape.callout-property-publisher.md)|Returns a  **[CalloutFormat](calloutformat-object-publisher.md)** object representing the formatting of a line callout.|
| [CatalogMergeItems](shape.catalogmergeitems-property-publisher.md)|Returns a  **CatalogMergeShapes** collection that represents the shapes included in the catalog merge area. Read-only.|
| [ConnectionSiteCount](shape.connectionsitecount-property-publisher.md)|Returns a  **Long** indicating the count of connection sites on the current **Shape** object. Read-only.|
| [Connector](shape.connector-property-publisher.md)|Returns an  **MsoTriState**value indicating whether the specified shape is a connector. Read-only.|
| [ConnectorFormat](shape.connectorformat-property-publisher.md)|Returns a  **[ConnectorFormat](connectorformat-object-publisher.md)** object that contains connector formatting properties. Applies to  **Shape** or **ShapeRange** objects that represent connectors. Read-only.|
| [Fill](shape.fill-property-publisher.md)| Returns a **[FillFormat](fillformat-object-publisher.md)** object representing the fill for the specified shape or table cell.|
| [Glow](shape.glow-property-publisher.md)|Returns a  [GlowFormat](glowformat-object-publisher.md) object that represents the glow formatting for a shape. Read-only.|
| [GroupItems](shape.groupitems-property-publisher.md)|Returns a  **[GroupShapes](groupshapes-object-publisher.md)** collection if the specified shape is a group.|
| [HasTable](shape.hastable-property-publisher.md)|Returns  **msoTrue** if the shape represents a **TableFrame** object or **msoFalse** if the shape represents any other object type. Read-only.|
| [HasTextFrame](shape.hastextframe-property-publisher.md)|Returns an  **MsoTriState** constant if the specified shape has a **TextFrame** object associated with it. Read-only.|
| [Height](shape.height-property-publisher.md)|Returns or sets a  **Variant** that represents the height (in points) of a specified table row or shape. Read/write.|
| [HorizontalFlip](shape.horizontalflip-property-publisher.md)|Indicates whether the specified shape has been flipped around its horizontal axis. Read-only.|
| [Hyperlink](shape.hyperlink-property-publisher.md)|Returns a  **[Hyperlink](hyperlink-object-publisher.md)** object representing the hyperlink associated with the specified shape.|
| [ID](shape.id-property-publisher.md)|Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.|
| [InlineAlignment](shape.inlinealignment-property-publisher.md)|Returns or sets a  **PbInlineAlignment** constant that indicates whether an inline shape has left, right, or in-text alignment. Read/write.|
| [InlineTextRange](shape.inlinetextrange-property-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that reflects the position of the inline shape in its containing text range. Read-only.|
| [IsExcess](shape.isexcess-property-publisher.md)|Indicates whether the parent  **Shape** object is an excess shape after the document template (wizard) is changed by using the ** [Document.ChangeDocument](document.changedocument-method-publisher.md)** method or by using the **Change Template** command in the user interface. Microsoft Publisher places any excess shape under **Extra Content** in the **Format Publication** task pane. Read-only.|
| [IsGroupMember](shape.isgroupmember-property-publisher.md)|Returns  **True** if the specified shape is a member of a group, **False** otherwise. Read-only **Boolean**.|
| [IsInline](shape.isinline-property-publisher.md)|Returns an  **MsoTriState** constant that specifies whether a shape is inline (contained in a text run). Read-only.|
| [Left](shape.left-property-publisher.md)|Returns or sets a  **Variant** indicating the distance from the left edge of the page to the leftmost edge of the specified shape. Numeric values are in points; all other values are in any measurement supported by Publisher (for example, "2.5 in"). Read/write.|
| [Line](shape.line-property-publisher.md)|Returns a  **[LineFormat](lineformat-object-publisher.md)** object that contains line formatting properties for the specified shape. (For a line, the  **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.).|
| [LinkFormat](shape.linkformat-property-publisher.md)|Returns a  [LinkFormat](linkformat-object-publisher.md)object that contains the properties that are unique to linked OLE objects. Read-only.|
| [LockAspectRatio](shape.lockaspectratio-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether the specified shape retains its original proportions when you resize it. Read/write.|
| [Name](shape.name-property-publisher.md)|Returns or sets a  **String** value indicating the name of the specified object. Read/write.|
| [Nodes](shape.nodes-property-publisher.md)|Returns a  **[ShapeNodes](shapenodes-object-publisher.md)** collection that represents the geometric description of the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent freeform drawings.|
| [OLEFormat](shape.oleformat-property-publisher.md)|Returns an  **[OLEFormat](oleformat-object-publisher.md)** object that contains OLE formatting properties for the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent OLE objects.|
| [Parent](shape.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ParentGroupShape](shape.parentgroupshape-property-publisher.md)|Returns a  **[Shape](shape-object-publisher.md)** object that represents the common parent shape of a child shape or a range of child shapes.|
| [PictureFormat](shape.pictureformat-property-publisher.md)|Returns a  **[PictureFormat](pictureformat-object-publisher.md)** object that contains picture formatting properties for the specified object. Applies to  **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** objects that represent pictures or OLE objects. Read-only.|
| [Reflection](shape.reflection-property-publisher.md)|Returns a  [ReflectionFormat](reflectionformat-object-publisher.md) object that represents the reflection formatting for a shape. Read-only.|
| [Rotation](shape.rotation-property-publisher.md)|Returns or sets a  **Single** that represents the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write.|
| [Shadow](shape.shadow-property-publisher.md)|Returns a  **[ShadowFormat](shadowformat-object-publisher.md)** object that represents the shadow formatting for the specified shape.|
| [SoftEdge](shape.softedge-property-publisher.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-publisher.md)** object that represents the soft edge formatting for a shape. Read-only.|
| [Table](shape.table-property-publisher.md)|Returns a  **Table** object that represents a table in Microsoft Publisher.|
| [Tags](shape.tags-property-publisher.md)|Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.|
| [TextEffect](shape.texteffect-property-publisher.md)|Returns a  **[TextEffectFormat](texteffectformat-object-publisher.md)** object that represents the text formatting properties of a WordArt object.|
| [TextFrame](shape.textframe-property-publisher.md)|Returns a  **[TextFrame](textframe-object-publisher.md)** object that represents the text in a shape and the properties that control the margins and orientation of the text.|
| [TextWrap](shape.textwrap-property-publisher.md)|Returns a  **[WrapFormat](wrapformat-object-publisher.md)** object that represents the properties for wrapping text around a shape or shape range.|
| [ThreeD](shape.threed-property-publisher.md)|Returns a  **[ThreeDFormat](threedformat-object-publisher.md)** object.|
| [Top](shape.top-property-publisher.md)|Returns or sets a  **Variant** that represents the distance between the top of the page and the top of a shape. Read/write.|
| [Type](shape.type-property-publisher.md)|Specifies the shape type. Read-only.|
| [VerticalFlip](shape.verticalflip-property-publisher.md)|Returns  **msoTrue** if the specified shape has been flipped around its vertical axis. Read-only.|
| [Vertices](shape.vertices-property-publisher.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. Read-only  **Variant**.|
| [WebCheckBox](shape.webcheckbox-property-publisher.md)|Returns the  **[WebCheckBox](webcheckbox-object-publisher.md)** object associated with the specified shape.|
| [WebCommandButton](shape.webcommandbutton-property-publisher.md)|Returns the  **[WebCommandButton](webcommandbutton-object-publisher.md)** object associated with the specified shape.|
| [WebListBox](shape.weblistbox-property-publisher.md)|Returns the  **[WebListBox](weblistbox-object-publisher.md)** object associated with the specified shape.|
| [WebNavigationBarSetName](shape.webnavigationbarsetname-property-publisher.md)|Returns a  **String** that represents the name of the Web navigation bar set the specified shape is an instance of. Read-only.|
| [WebOptionButton](shape.weboptionbutton-property-publisher.md)|Returns the  **[WebOptionButton](weboptionbutton-object-publisher.md)** object associated with the specified shape.|
| [WebTextBox](shape.webtextbox-property-publisher.md)|Returns the  **[WebTextBox](webtextbox-object-publisher.md)** object associated with the specified shape.|
| [Width](shape.width-property-publisher.md)|Returns or sets a  **Variant** that represents the width (in points) of a specified table column or shape. Read/write.|
| [Wizard](shape.wizard-property-publisher.md)|Returns a  **[Wizard](wizard-object-publisher.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.|
| [WizardTag](shape.wizardtag-property-publisher.md)|Returns or sets a  **PbWizardTag**constant indicating the function of a specified shape with respect to its publication design. Read/write.|
| [WizardTagInstance](shape.wizardtaginstance-property-publisher.md)|Returns or sets a  **Long** indicating the instance of the specified shape compared with other shapes having the same wizard tag. Read/write.|
| [ZOrderPosition](shape.zorderposition-property-publisher.md)|Returns a  **Long** indicating the position of the specified shape or shape range in the z-order. Read-only.|

