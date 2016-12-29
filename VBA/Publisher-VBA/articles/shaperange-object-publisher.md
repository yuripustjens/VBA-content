---
title: ShapeRange Object (Publisher)
keywords: vbapb10.chm2359295
f1_keywords: vbapb10.chm2359295
ms.prod: PUBLISHER
api_name: Publisher.ShapeRange
ms.assetid: c85967c9-af43-747d-7e0b-64ddc22c84be
ms.locale: en-US
---


# ShapeRange Object (Publisher)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. You can include whichever shapes you want — chosen from among all the shapes in the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document.


 **Note**  Most operations that you can do with a  **[Shape](http://msdn.microsoft.com/library/shape-object-publisher%28Office.15%29.aspx)** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, will cause an error. This section describes how to:


- Return a set of shapes.
    
- Return a  **ShapeRange** object within a selection or range.
    
- Align, distribute, and group shapes in a  **ShapeRange** object.
    

## Example

Use  **Shapes.Range** (index), where index is the index number of the shape or an array that contains index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes in a publication. You can use Visual Basic's **Array** function to construct an array of index numbers. The following example sets the fill pattern for shapes one through three on the active publication.


```
Sub ChangeFillPattern() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2, 3)) _ 
        .Fill.PresetGradient Style:=msoGradientDiagonalDown, _ 
        Variant:=1, PresetGradientType:=msoGradientHorizon 
End Sub
```

Although you can use the  **[Range](http://msdn.microsoft.com/library/shapes.range-method-publisher%28Office.15%29.aspx)** method to return any number of shapes, it is simpler to use the **[Item](http://msdn.microsoft.com/library/shaperange.item-method-publisher%28Office.15%29.aspx)** method if you want to return only a single member of the collection. For example, **Shapes** (1) is simpler than **Shapes.Range** (1).

Use  **Selection.ShapeRange** (index), where index is the index number of the shape, to return a **Shape** object that represents a shape within a selection. The following example selects the first two shapes on the first page of the active publication and then sets the fill for the first shape in the selection.




```
Sub ChangeFillForShapeRange() 
    ActiveDocument.Pages(1).Shapes.Range(Array(1, 2)).Select 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

This example selects all the shapes on the first page of the active publication, then adds and formats text in the second shape in the range.




```
Sub SelectShapesOnPageOne() 
    ActiveDocument.Pages(1).Shapes.Range.Select 
    With Selection.ShapeRange(2).TextFrame.TextRange 
        .Text = "Shape Number 2" 
        .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
        .Font.Size = 25 
    End With 
End Sub
```

Use the  **[Align](http://msdn.microsoft.com/library/shaperange.align-method-publisher%28Office.15%29.aspx)**, **[Distribute](http://msdn.microsoft.com/library/shaperange.distribute-method-publisher%28Office.15%29.aspx)**, or **[ZOrder](http://msdn.microsoft.com/library/shaperange.zorder-method-publisher%28Office.15%29.aspx)** method to position a set of shapes relative to each other or relative to the document. This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.




```
Sub AlignDistibuteShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
 
    With rngShapes 
        .Align AlignCmd:=msoAlignLefts, RelativeTo:=msoFalse 
        .Distribute DistributeCmd:=msoDistributeVertically, RelativeTo:=msoTrue 
    End With 
End Sub
```

Use the  **[Group](http://msdn.microsoft.com/library/shaperange.group-method-publisher%28Office.15%29.aspx)**, **[Regroup](http://msdn.microsoft.com/library/shaperange.regroup-method-publisher%28Office.15%29.aspx)**, or **[Ungroup](http://msdn.microsoft.com/library/shaperange.ungroup-method-publisher%28Office.15%29.aspx)** method to create and work with a single shape formed from a shape range. The **[GroupItems](http://msdn.microsoft.com/library/shaperange.groupitems-property-publisher%28Office.15%29.aspx)** property for a **Shape** object returns the **[GroupShapes](http://msdn.microsoft.com/library/groupshapes-object-publisher%28Office.15%29.aspx)** object, which represents all the shapes that were grouped to form one shape. This example specifies a shape range and left-aligns and vertically distributes the shapes on the page.




```
Sub GroupShapes() 
    Dim rngShapes As ShapeRange 
    Set rngShapes = ActiveDocument.Pages(1).Shapes.Range 
    rngShapes.Group 
 
    rngShapes(1).Fill.OneColorGradient _ 
        Style:=msoGradientFromCenter, _ 
        Variant:=2, Degree:=1 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToCatalogMergeArea](http://msdn.microsoft.com/library/shaperange.addtocatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[Align](http://msdn.microsoft.com/library/shaperange.align-method-publisher%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/shaperange.apply-method-publisher%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/shaperange.copy-method-publisher%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/shaperange.cut-method-publisher%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/shaperange.delete-method-publisher%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/shaperange.distribute-method-publisher%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/shaperange.duplicate-method-publisher%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/shaperange.flip-method-publisher%28Office.15%29.aspx)|
|[GetHeight](http://msdn.microsoft.com/library/shaperange.getheight-method-publisher%28Office.15%29.aspx)|
|[GetLeft](http://msdn.microsoft.com/library/shaperange.getleft-method-publisher%28Office.15%29.aspx)|
|[GetTop](http://msdn.microsoft.com/library/shaperange.gettop-method-publisher%28Office.15%29.aspx)|
|[GetWidth](http://msdn.microsoft.com/library/shaperange.getwidth-method-publisher%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/shaperange.group-method-publisher%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/shaperange.incrementleft-method-publisher%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/shaperange.incrementrotation-method-publisher%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/shaperange.incrementtop-method-publisher%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/shaperange.item-method-publisher%28Office.15%29.aspx)|
|[MoveIntoTextFlow](http://msdn.microsoft.com/library/shaperange.moveintotextflow-method-publisher%28Office.15%29.aspx)|
|[MoveOutOfTextFlow](http://msdn.microsoft.com/library/shaperange.moveoutoftextflow-method-publisher%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/shaperange.pickup-method-publisher%28Office.15%29.aspx)|
|[Regroup](http://msdn.microsoft.com/library/shaperange.regroup-method-publisher%28Office.15%29.aspx)|
|[RemoveFromCatalogMergeArea](http://msdn.microsoft.com/library/shaperange.removefromcatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/shaperange.rerouteconnections-method-publisher%28Office.15%29.aspx)|
|[SaveAsBuildingBlock](http://msdn.microsoft.com/library/shaperange.saveasbuildingblock-method-publisher%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/shaperange.saveaspicture-method-publisher%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/shaperange.scaleheight-method-publisher%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/shaperange.scalewidth-method-publisher%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/shaperange.select-method-publisher%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/shaperange.setshapesdefaultproperties-method-publisher%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/shaperange.ungroup-method-publisher%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/shaperange.zorder-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Adjustments](http://msdn.microsoft.com/library/shaperange.adjustments-property-publisher%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/shaperange.alternativetext-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/shaperange.application-property-publisher%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/shaperange.autoshapetype-property-publisher%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/shaperange.blackwhitemode-property-publisher%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/shaperange.callout-property-publisher%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/shaperange.connectionsitecount-property-publisher%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/shaperange.connector-property-publisher%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/shaperange.connectorformat-property-publisher%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/shaperange.count-property-publisher%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/shaperange.fill-property-publisher%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/shaperange.glow-property-publisher%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/shaperange.groupitems-property-publisher%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/shaperange.hastable-property-publisher%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/shaperange.hastextframe-property-publisher%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/shaperange.height-property-publisher%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/shaperange.horizontalflip-property-publisher%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/shaperange.hyperlink-property-publisher%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/shaperange.id-property-publisher%28Office.15%29.aspx)|
|[InlineAlignment](http://msdn.microsoft.com/library/shaperange.inlinealignment-property-publisher%28Office.15%29.aspx)|
|[InlineTextRange](http://msdn.microsoft.com/library/shaperange.inlinetextrange-property-publisher%28Office.15%29.aspx)|
|[IsInline](http://msdn.microsoft.com/library/shaperange.isinline-property-publisher%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/shaperange.left-property-publisher%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/shaperange.line-property-publisher%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/shaperange.linkformat-property-publisher%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/shaperange.lockaspectratio-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/shaperange.name-property-publisher%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/shaperange.nodes-property-publisher%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/shaperange.oleformat-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shaperange.parent-property-publisher%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/shaperange.pictureformat-property-publisher%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/shaperange.reflection-property-publisher%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/shaperange.rotation-property-publisher%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/shaperange.shadow-property-publisher%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/shaperange.softedge-property-publisher%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/shaperange.table-property-publisher%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/shaperange.tags-property-publisher%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/shaperange.texteffect-property-publisher%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/shaperange.textframe-property-publisher%28Office.15%29.aspx)|
|[TextWrap](http://msdn.microsoft.com/library/shaperange.textwrap-property-publisher%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/shaperange.threed-property-publisher%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/shaperange.top-property-publisher%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/shaperange.type-property-publisher%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/shaperange.verticalflip-property-publisher%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/shaperange.vertices-property-publisher%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/shaperange.width-property-publisher%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/shaperange.wizard-property-publisher%28Office.15%29.aspx)|
|[WizardTag](http://msdn.microsoft.com/library/shaperange.wizardtag-property-publisher%28Office.15%29.aspx)|
|[WizardTagInstance](http://msdn.microsoft.com/library/shaperange.wizardtaginstance-property-publisher%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/shaperange.zorderposition-property-publisher%28Office.15%29.aspx)|

