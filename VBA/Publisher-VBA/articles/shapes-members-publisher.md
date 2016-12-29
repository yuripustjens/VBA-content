---
title: Shapes Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 0d1e8129-531e-fb5c-901a-66dee3744dbd
ms.locale: en-US
---


# Shapes Members (Publisher)
A collection of  **[Shape](shape-object-publisher.md)** objects that represent all the shapes on a page of a publication. Each **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddBuildingBlock](shapes.addbuildingblock-method-publisher.md)|Adds a  **[BuildingBlock](buildingblock-object-publisher.md)** object and returns a **[Shape](shape-object-publisher.md)** object on the page that represents the building block.|
| [AddCallout](shapes.addcallout-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing a borderless line callout to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddCatalogMergeArea](shapes.addcatalogmergearea-method-publisher.md)|Adds a  **Shape** object that represents the specified publication's catalog merge area.|
| [AddCatalogMergeFieldToCanvas](shapes.addcatalogmergefieldtocanvas-method-publisher.md)|Adds a catalog merge field of the specified type to the canvas. Returns nothing.|
| [AddConnector](shapes.addconnector-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing a connector to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddCurve](shapes.addcurve-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing a Bézier curve to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddEmptyPictureFrame](shapes.addemptypictureframe-method-publisher.md)|Returns a  **Shape** object that represents an empty picture frame inserted at the specified coordinates.|
| [AddGroupWizard](shapes.addgroupwizard-method-publisher.md)|Adds a  **Shape** object representing a Design Gallery object to the publication.|
| [AddLabel](shapes.addlabel-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing a text label to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddLine](shapes.addline-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing a line to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddOLEObject](shapes.addoleobject-method-publisher.md)|Adds a new  **[Shape](shape-object-publisher.md)** object representing an OLE object to the specified **[Shapes](shapes-object-publisher.md)** collection.|
| [AddPicture](shapes.addpicture-method-publisher.md)|Adds a new  **Shape** object representing a picture to the specified **Shapes** collection.|
| [AddPolyline](shapes.addpolyline-method-publisher.md)|Adds a new  **Shape** object representing an open polyline or a closed polygon to the specified **Shapes** collection.|
| [AddShape](shapes.addshape-method-publisher.md)|Adds a new  **Shape** object representing an AutoShape to the specified **Shapes** collection.|
| [AddTable](shapes.addtable-method-publisher.md)|Adds a new  **Shape** object representing a table to the specified **Shapes** collection.|
| [AddTextbox](shapes.addtextbox-method-publisher.md)|Adds a new  **Shape** object representing a text box to the specified **Shapes** collection.|
| [AddTextEffect](shapes.addtexteffect-method-publisher.md)|Adds a new  **Shape** object representing a WordArt object to the specified **Shapes** collection.|
| [AddWebControl](shapes.addwebcontrol-method-publisher.md)|Adds a new  **Shape** object representing a Web form control to the specified **Shapes** collection.|
| [AddWebNavigationBar](shapes.addwebnavigationbar-method-publisher.md)|Adds a  **Shape** object of type **pbWebNavigationBar** to the current page of a publication.|
| [AddWordArt](shapes.addwordart-method-publisher.md)|Returns a  **Shape** object that represents the WordArt to be added to the publication.|
| [BuildFreeform](shapes.buildfreeform-method-publisher.md)|Builds a freeform object. Returns a  [FreeformBuilder](freeformbuilder-object-publisher.md)object that represents the freeform as it is being built.|
| [FindShapeByWizardTag](shapes.findshapebywizardtag-method-publisher.md)|Returns a  **ShapeRange** object representing one or all of the shapes placed in a publication by a wizard and bearing the specified wizard tag.|
| [Item](shapes.item-method-publisher.md)|Returns an individual object in a specified collection.|
| [Paste](shapes.paste-method-publisher.md)|Pastes the shapes or text on the Clipboard into the specified  **[Shapes](shapes-object-publisher.md)** collection, at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a **[ShapeRange](shaperange-object-publisher.md)** object that represents the pasted objects.|
| [Range](shapes.range-method-publisher.md)|Returns a  **[ShapeRange](shaperange-object-publisher.md)** object that represents a subset of the shapes in a **Shapes** collection.|
| [SelectAll](shapes.selectall-method-publisher.md)|Selects all the shapes in the specified  **[Shapes](shapes-object-publisher.md)** collection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](shapes.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [CanvasArrangementType](shapes.canvasarrangementtype-property-publisher.md)|A  **[pbCanvasArrangementType](pbcanvasarrangementtype-enumeration-publisher.md)** constant that specifies the arrangement type of the canvas. Read/write.|
| [CanvasesCount](shapes.canvasescount-property-publisher.md)|A  **Long** that specifies the number of canvases in the **Shapes** collection. Read-only.|
| [Count](shapes.count-property-publisher.md)|Returns a  **Long** that represents the number of items in the specified collection.|
| [Parent](shapes.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

