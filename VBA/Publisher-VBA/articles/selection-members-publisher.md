---
title: Selection Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: bcc0f7af-be9f-917a-1920-d0c2946d95f7
ms.locale: en-US
---


# Selection Members (Publisher)
Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the publication, or it represents the cursor if nothing in the publication is selected. There can only be one  **Selection** object per publication window pane, and only one **Selection** object in the entire application can be active.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Unselect](selection.unselect-method-publisher.md)|Cancels the current selection.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](selection.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [ChildShapeRange](selection.childshaperange-property-publisher.md)|Returns a  **[ShapeRange](shaperange-object-publisher.md)** object representing the child shapes of a selection.|
| [Parent](selection.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ShapeRange](selection.shaperange-property-publisher.md)|Returns a  **[ShapeRange](shaperange-object-publisher.md)** collection that represents all the **Shape** objects in the specified range or selection. The shape range can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts.|
| [TableCellRange](selection.tablecellrange-property-publisher.md)|Returns a  **CellRange** object that represents the cells in a table selection.|
| [TextRange](selection.textrange-property-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the text that is attached to a shape and properties and methods for manipulating the text.|
| [Type](selection.type-property-publisher.md)|Returns a  **PbSelectionType** constant that represents the selection type. Read-only.|

