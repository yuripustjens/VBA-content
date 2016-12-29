---
title: Shape Methods (Publisher)
ms.prod: PUBLISHER
ms.assetid: a8593165-57d0-448c-a404-2a05892d54d9
ms.locale: en-US
---


# Shape Methods (Publisher)

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
|Name|Description|

