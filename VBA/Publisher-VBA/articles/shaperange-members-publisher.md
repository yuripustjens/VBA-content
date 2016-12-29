---
title: ShapeRange Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 50146f74-1df4-8123-f42f-d4f5d1efba7b
ms.locale: en-US
---


# ShapeRange Members (Publisher)
Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. You can include whichever shapes you want — chosen from among all the shapes in the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddToCatalogMergeArea](shaperange.addtocatalogmergearea-method-publisher.md)|Adds the specified shape or shapes to the publication page's catalog merge area.|
| [Align](shaperange.align-method-publisher.md)|Aligns all the shapes in the specified  **ShapeRange** object.|
| [Apply](shaperange.apply-method-publisher.md)|Applies formatting copied from another shape or shape range using the  **[PickUp](shaperange.pickup-method-publisher.md)** method.|
| [Copy](shaperange.copy-method-publisher.md)|Copies the specified object to the Clipboard.|
| [Cut](shaperange.cut-method-publisher.md)|Deletes the specified object and places it on the Clipboard.|
| [Delete](shaperange.delete-method-publisher.md)|Deletes the specified object.|
| [Distribute](shaperange.distribute-method-publisher.md)|Evenly distributes the shapes in the specified shape range.|
| [Duplicate](shaperange.duplicate-method-publisher.md)|Creates a duplicate of the specified  **[Shape](shape-object-publisher.md)** or **[ShapeRange](shaperange-object-publisher.md)** object, adds the new shape or range of shapes to the **Shapes** collection immediately after the shape or range of shapes specified originally, and then returns the new **Shape** or **ShapeRange** object.|
| [Flip](shaperange.flip-method-publisher.md)|Flips the specified shape around its horizontal or vertical axis, or flips all the shapes in the specified shape range around their horizontal or vertical axes.|
| [GetHeight](shaperange.getheight-method-publisher.md)|Returns the height of the shape or shape range as a  **Single** in the specified units.|
| [GetLeft](shaperange.getleft-method-publisher.md)|Returns the distance of the shape's or shape range's left edge from the left edge of the leftmost page in the current view as a  **Single** in the specified units.|
| [GetTop](shaperange.gettop-method-publisher.md)|Returns the distance of the shape's or shape range's top edge from the top edge of the leftmost page in the current view as a  **Single** in the specified units.|
| [GetWidth](shaperange.getwidth-method-publisher.md)|Returns the width of the shape or shape range as a  **Single** in the specified units. .|
| [Group](shaperange.group-method-publisher.md)|Groups the shapes in the specified shape range. Returns the grouped shapes as a single  **[Shape](shape-object-publisher.md)** object.|
| [IncrementLeft](shaperange.incrementleft-method-publisher.md)|Moves the specified shape or shape range horizontally by the specified distance.|
| [IncrementRotation](shaperange.incrementrotation-method-publisher.md)|Changes the rotation of the specified shape around the z-axis (extends outward from the plane of the publication) by the specified number of degrees.|
| [IncrementTop](shaperange.incrementtop-method-publisher.md)|Moves the specified shape or shape range vertically by the specified distance.|
| [Item](shaperange.item-method-publisher.md)|Returns an individual object in a specified collection.|
| [MoveIntoTextFlow](shaperange.moveintotextflow-method-publisher.md)|Moves a given shape into the text flow defined by  ** [TextRange Object](textrange-object-publisher.md)**. The shape will always be inserted inline at the beginning of the text flow.|
| [MoveOutOfTextFlow](shaperange.moveoutoftextflow-method-publisher.md)|Moves a given inline shape out of its containing text range, defined by  ** [TextRange Object](textrange-object-publisher.md)**, and makes the shape fixed.|
| [PickUp](shaperange.pickup-method-publisher.md)|Copies formatting from a shape or shape range so that it can be copied to another shape or shape range using the  **[Apply](shaperange.apply-method-publisher.md)** method.|
| [Regroup](shaperange.regroup-method-publisher.md)|Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single  **Shape** object.|
| [RemoveFromCatalogMergeArea](shaperange.removefromcatalogmergearea-method-publisher.md)|Removes a shape from the specified page's catalog merge area. Removed shapes are not deleted, but instead remain in place on the page containing the catalog merge area.|
| [RerouteConnections](shaperange.rerouteconnections-method-publisher.md)|Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the  **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.|
| [SaveAsBuildingBlock](shaperange.saveasbuildingblock-method-publisher.md)|Saves a single shape range as a building block. Returns the resulting  **[BuildingBlock](buildingblock-object-publisher.md)** object.|
| [SaveAsPicture](shaperange.saveaspicture-method-publisher.md)|Saves a range of one or more shapes as a picture file.|
| [ScaleHeight](shaperange.scaleheight-method-publisher.md)|Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.|
| [ScaleWidth](shaperange.scalewidth-method-publisher.md)|Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.|
| [Select](shaperange.select-method-publisher.md)|Selects the specified object.|
| [SetShapesDefaultProperties](shaperange.setshapesdefaultproperties-method-publisher.md)|Applies the formatting for the specified shape or shape range to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.|
| [Ungroup](shaperange.ungroup-method-publisher.md)|Ungroups the specified group of shapes or any groups of shapes in the specified shape range. If the specified shape is a picture or OLE object, Microsoft Publisher will break it apart and convert it to an ungrouped set of shapes. (For example, an embedded Microsoft Excel spreadsheet is converted into lines and text boxes.) Returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-publisher.md)** object.|
| [ZOrder](shaperange.zorder-method-publisher.md)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|

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

