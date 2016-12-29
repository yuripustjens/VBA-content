---
title: Shape Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 7a6cf332-6f30-6444-bcd9-2fd98ecd611d
ms.locale: en-US
---


# Shape Members (Publisher)
Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, ActiveX control, or picture. The  **Shape** object is a member of the **[Shapes](shapes-object-publisher.md)** collection, which includes all the shapes on a page or in a selection.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddToCatalogMergeArea](shape.addtocatalogmergearea-method-publisher.md)|Adds the specified shape or shapes to the publication page's catalog merge area.|
| [Apply](shape.apply-method-publisher.md)|Applies formatting copied from another shape or shape range using the  **[PickUp](shape.pickup-method-publisher.md)** method.|
| [Copy](shape.copy-method-publisher.md)|Copies the specified object to the Clipboard.|
| [Cut](shape.cut-method-publisher.md)|Deletes the specified object and places it on the Clipboard.|
| [Delete](shape.delete-method-publisher.md)|Deletes the specified object.|
| [Duplicate](shape.duplicate-method-publisher.md)|Creates a duplicate of the specified  **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** object, adds the new shape or range of shapes to the **Shapes** collection immediately after the shape or range of shapes specified originally, and then returns the new **Shape** or **ShapeRange** object.|
| [Flip](shape.flip-method-publisher.md)|Flips the specified shape around its horizontal or vertical axis, or flips all the shapes in the specified shape range around their horizontal or vertical axes.|
| [GetHeight](shape.getheight-method-publisher.md)|Returns the height of the shape or shape range as a  **Single** in the specified units.|
| [GetLeft](shape.getleft-method-publisher.md)|Returns the distance of the shape's or shape range's left edge from the left edge of the leftmost page in the current view as a  **Single** in the specified units.|
| [GetTop](shape.gettop-method-publisher.md)|Returns the distance of the shape's or shape range's top edge from the top edge of the leftmost page in the current view as a  **Single** in the specified units.|
| [GetWidth](shape.getwidth-method-publisher.md)|Returns the width of the shape or shape range as a  **Single** in the specified units.|
| [IncrementLeft](shape.incrementleft-method-publisher.md)|Moves the specified shape or shape range horizontally by the specified distance.|
| [IncrementRotation](shape.incrementrotation-method-publisher.md)|Changes the rotation of the specified shape around the z-axis (extends outward from the plane of the publication) by the specified number of degrees.|
| [IncrementTop](shape.incrementtop-method-publisher.md)|Moves the specified shape or shape range vertically by the specified distance.|
| [MoveIntoTextFlow](shape.moveintotextflow-method-publisher.md)|Moves a given shape into the text flow defined by  ** [TextRange Object](textrange-object-publisher.md)**. The shape will always be inserted inline at the beginning of the text flow.|
| [MoveOutOfTextFlow](shape.moveoutoftextflow-method-publisher.md)|Moves a given inline shape out of its containing text range, defined by  ** [TextRange Object](textrange-object-publisher.md)**, and makes the shape fixed.|
| [MoveToPage](shape.movetopage-method-publisher.md)|Moves a shape to the specified page.|
| [PickUp](shape.pickup-method-publisher.md)|Copies formatting from a shape or shape range so that it can be copied to another shape or shape range using the  **[Apply](shaperange.apply-method-publisher.md)** method.|
| [RemoveCatalogMergeArea](shape.removecatalogmergearea-method-publisher.md)|Deletes the catalog merge area from the specified publication page. All shapes contained in the catalog merge area remain in place on the page, but are no longer connected to the catalog merge data source.|
| [RemoveFromCatalogMergeArea](shape.removefromcatalogmergearea-method-publisher.md)|Removes a shape from the specified page's catalog merge area. Removed shapes are not deleted, but instead remain in place on the page containing the catalog merge area.|
| [RerouteConnections](shape.rerouteconnections-method-publisher.md)|Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the  **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.|
| [SaveAsBuildingBlock](shape.saveasbuildingblock-method-publisher.md)|Saves a single shape as a building block. Returns the resulting  **[BuildingBlock](buildingblock-object-publisher.md)** object.|
| [SaveAsPicture](shape.saveaspicture-method-publisher.md)|Saves a single shape as a picture file.|
| [ScaleHeight](shape.scaleheight-method-publisher.md)|Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.|
| [ScaleWidth](shape.scalewidth-method-publisher.md)|Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.|
| [Select](shape.select-method-publisher.md)|Selects the specified object.|
| [SetCaption](shape.setcaption-method-publisher.md)||
| [SetShapesDefaultProperties](shape.setshapesdefaultproperties-method-publisher.md)|Applies the formatting for the specified shape or shape range to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.|
| [Ungroup](shape.ungroup-method-publisher.md)|Ungroups the specified group of shapes or any groups of shapes in the specified shape range. If the specified shape is a picture or OLE object, Microsoft Publisher will break it apart and convert it to an ungrouped set of shapes. (For example, an embedded Microsoft Office Excel spreadsheet is converted into lines and text boxes.) Returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-publisher.md)** object.|
| [ZOrder](shape.zorder-method-publisher.md)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|

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

