---
title: Shape Object (Publisher)
keywords: vbapb10.chm2293759
f1_keywords: vbapb10.chm2293759
ms.prod: PUBLISHER
api_name: Publisher.Shape
ms.assetid: 666cb7f0-62a8-f419-9838-007ef29506ee
ms.locale: en-US
---


# Shape Object (Publisher)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, ActiveX control, or picture. The  **Shape** object is a member of the **[Shapes](http://msdn.microsoft.com/library/shapes-object-publisher%28Office.15%29.aspx)** collection, which includes all the shapes on a page or in a selection.


 **Note**  There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a document; the **[ShapeRange](http://msdn.microsoft.com/library/shaperange-object-publisher%28Office.15%29.aspx)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); the **Shape** object, which represents a single shape on a document. If you want to work with several shape at the same time or with shapes within the selection, use a **ShapeRange** collection. This section describes how to:


- Return an existing shape on a document.
    
- Return a shape or shapes within a selection.
    
- Return a newly created shape.
    
- Work with a group of shapes.
    
- Format a shape.
    
- Use other important shape properties.
    

## Example

Use  **[Shapes](http://msdn.microsoft.com/library/shapes-object-publisher%28Office.15%29.aspx)** (index), where index is the name or the index number, to return a single **Shape** object. The following example horizontally flips shape one on the active document.


```
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

The following example horizontally flips the shape named "Rectangle 1" on the active document.




```
Sub FlipShapeByName() 
    ActiveDocument.Pages(1).Shapes("Rectangle 1") _ 
        .Flip FlipCmd:=msoFlipHorizontal 
End Sub
```

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named "Rectangle 2," "TextBox 3," and "Oval 4." To give a shape a more meaningful name, set the  **Name** property of the shape.

Use  **Selection.ShapeRange** (index), where index is the name or the index number, to return a **Shape** object that represents a shape within a selection. The following example sets the fill for the first shape in the selection, assuming that the selection contains at least one shape.




```
Sub FillSelectedShape() 
    Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0) 
End Sub
```

The following example sets the fill for all the shapes in the selection, assuming that the selection contains at least one shape.




```
Sub FillAllSelectedShapes() 
    Dim shpShape As Shape 
    For Each
```




```
shpShape In Selection.ShapeRange 
       
```




```
shpShape.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
    Next shpShape 
End Sub
```

To add a  **Shape** object to the collection of shapes for the specified document and return a **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection: **[AddCallout](http://msdn.microsoft.com/library/shapes.addcallout-method-publisher%28Office.15%29.aspx)**, **[AddConnector](http://msdn.microsoft.com/library/shapes.addconnector-method-publisher%28Office.15%29.aspx)**, **[AddCurve](http://msdn.microsoft.com/library/shapes.addcurve-method-publisher%28Office.15%29.aspx)**, **[AddLabel](http://msdn.microsoft.com/library/shapes.addlabel-method-publisher%28Office.15%29.aspx)**, **[AddLine](http://msdn.microsoft.com/library/shapes.addline-method-publisher%28Office.15%29.aspx)**, **[AddOLEObject](http://msdn.microsoft.com/library/shapes.addoleobject-method-publisher%28Office.15%29.aspx)**, **[AddPolyline](http://msdn.microsoft.com/library/shapes.addpolyline-method-publisher%28Office.15%29.aspx)**, **[AddShape](http://msdn.microsoft.com/library/shapes.addshape-method-publisher%28Office.15%29.aspx)**, **[AddTextBox](http://msdn.microsoft.com/library/shapes.addtextbox-method-publisher%28Office.15%29.aspx)** or **[AddTextEffect](http://msdn.microsoft.com/library/shapes.addtexteffect-method-publisher%28Office.15%29.aspx)**. The following example adds a rectangle to the active document.




```
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeRectangle, _ 
        Left:=400, Top:=72, Width:=100, Height:=200 
End Sub
```

Use  **[GroupItems](http://msdn.microsoft.com/library/shape.groupitems-property-publisher%28Office.15%29.aspx)** (index), where index is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape. Use the **[Group](http://msdn.microsoft.com/library/shaperange.group-method-publisher%28Office.15%29.aspx)** or **[Regroup](http://msdn.microsoft.com/library/shaperange.regroup-method-publisher%28Office.15%29.aspx)** method to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape. This example adds three shapes to the active publication, groups the shapes, and sets the fill color for each of the shapes in the group




```
Sub WorkWithGroupShapes() 
 
    With ActiveDocument.Pages(1).Shapes 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=100, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=250, _ 
            Top:=72, Width:=100, Height:=100 
        .AddShape Type:=msoShapeIsoscelesTriangle, Left:=400, _ 
            Top:=72, Width:=100, Height:=100 
        .SelectAll 
 
        With Selection.ShapeRange 
            .Group 
            .GroupItems(1).Fill.ForeColor _ 
                .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
            .GroupItems(2).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=255, Blue:=0) 
            .GroupItems(3).Fill.ForeColor _ 
                .RGB = RGB(Red:=0, Green:=0, Blue:=255) 
        End With 
    End With 
End Sub
```

Use the  **[Fill](http://msdn.microsoft.com/library/shape.fill-property-publisher%28Office.15%29.aspx)** property to return the **[FillFormat](http://msdn.microsoft.com/library/fillformat-object-publisher%28Office.15%29.aspx)** object, which contains all the properties and methods for formatting the fill of a closed shape. The **[Shadow](http://msdn.microsoft.com/library/shape.shadow-property-publisher%28Office.15%29.aspx)** property returns the **[ShadowFormat](http://msdn.microsoft.com/library/shadowformat-object-publisher%28Office.15%29.aspx)** object, which you use to format a shadow. Use the **[Line](http://msdn.microsoft.com/library/shape.line-property-publisher%28Office.15%29.aspx)** property to return a **[LineFormat](http://msdn.microsoft.com/library/lineformat-object-publisher%28Office.15%29.aspx)** object, which contains properties and methods for formatting lines and arrows. The **[TextEffect](http://msdn.microsoft.com/library/shape.texteffect-property-publisher%28Office.15%29.aspx)** property returns the **[TextEffectFormat](http://msdn.microsoft.com/library/texteffectformat-object-publisher%28Office.15%29.aspx)** object, which you use to format WordArt. The **[Callout](http://msdn.microsoft.com/library/shape.callout-property-publisher%28Office.15%29.aspx)** property returns the **[CalloutFormat](http://msdn.microsoft.com/library/calloutformat-object-publisher%28Office.15%29.aspx)** object, which you use to format line callouts. The **[TextWrap](http://msdn.microsoft.com/library/shape.textwrap-property-publisher%28Office.15%29.aspx)** property returns the **[WrapFormat](http://msdn.microsoft.com/library/wrapformat-object-publisher%28Office.15%29.aspx)** object, which you use to define how text wraps around shapes. The **[ThreeD](http://msdn.microsoft.com/library/shape.threed-property-publisher%28Office.15%29.aspx)** property returns the **[ThreeDFormat](http://msdn.microsoft.com/library/threedformat-object-publisher%28Office.15%29.aspx)** object, which you use to create 3-D shapes. You can use the **[PickUp](http://msdn.microsoft.com/library/shape.pickup-method-publisher%28Office.15%29.aspx)** and **[Apply](http://msdn.microsoft.com/library/shape.apply-method-publisher%28Office.15%29.aspx)** methods to transfer formatting from one shape to another.



Use the  **[SetShapesDefaultProperties](http://msdn.microsoft.com/library/shape.setshapesdefaultproperties-method-publisher%28Office.15%29.aspx)** method for a **Shape** object to set the formatting for the default shape for the document. New shapes inherit many of their attributes from the default shape.

Use the  **[Type](http://msdn.microsoft.com/library/shape.type-property-publisher%28Office.15%29.aspx)** property to specify the type of shape: freeform, AutoShape, OLE object, callout, or linked picture, for instance. Use the **[AutoShapeType](http://msdn.microsoft.com/library/shape.autoshapetype-property-publisher%28Office.15%29.aspx)** property to specify the type of AutoShape: oval, rectangle, or balloon, for instance.



Use the  **[Width](http://msdn.microsoft.com/library/shape.width-property-publisher%28Office.15%29.aspx)** and **[Height](http://msdn.microsoft.com/library/shape.height-property-publisher%28Office.15%29.aspx)** properties to specify the size of the shape.



Use  **[TextFrame](http://msdn.microsoft.com/library/shape.textframe-property-publisher%28Office.15%29.aspx)** and **[TextRange](http://msdn.microsoft.com/library/cell.textrange-property-publisher%28Office.15%29.aspx)** properties to return the **[TextFrame](http://msdn.microsoft.com/library/textframe-object-publisher%28Office.15%29.aspx)** and **[TextRange](http://msdn.microsoft.com/library/textrange-object-publisher%28Office.15%29.aspx)** objects, respectively, which contain all the properties and methods for inserting and formatting text within shapes and publications and linking the text frames together. The following example adds a text box to the first page of the active publication, then adds text to it and formats the text.




```
Sub CreateNewTextBox() 
    With ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
        Orientation:=pbTextOrientationHorizontal, Left:=100, _ 
        Top:=100, Width:=200, Height:=100).TextFrame.TextRange 
        .Text = "This is a textbox." 
        With .Font 
            .Name = "Stencil" 
            .Bold = msoTrue 
            .Size = 30 
        End With 
    End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddToCatalogMergeArea](http://msdn.microsoft.com/library/shape.addtocatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/shape.apply-method-publisher%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/shape.copy-method-publisher%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/shape.cut-method-publisher%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/shape.delete-method-publisher%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/shape.duplicate-method-publisher%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/shape.flip-method-publisher%28Office.15%29.aspx)|
|[GetHeight](http://msdn.microsoft.com/library/shape.getheight-method-publisher%28Office.15%29.aspx)|
|[GetLeft](http://msdn.microsoft.com/library/shape.getleft-method-publisher%28Office.15%29.aspx)|
|[GetTop](http://msdn.microsoft.com/library/shape.gettop-method-publisher%28Office.15%29.aspx)|
|[GetWidth](http://msdn.microsoft.com/library/shape.getwidth-method-publisher%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/shape.incrementleft-method-publisher%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/shape.incrementrotation-method-publisher%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/shape.incrementtop-method-publisher%28Office.15%29.aspx)|
|[MoveIntoTextFlow](http://msdn.microsoft.com/library/shape.moveintotextflow-method-publisher%28Office.15%29.aspx)|
|[MoveOutOfTextFlow](http://msdn.microsoft.com/library/shape.moveoutoftextflow-method-publisher%28Office.15%29.aspx)|
|[MoveToPage](http://msdn.microsoft.com/library/shape.movetopage-method-publisher%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/shape.pickup-method-publisher%28Office.15%29.aspx)|
|[RemoveCatalogMergeArea](http://msdn.microsoft.com/library/shape.removecatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[RemoveFromCatalogMergeArea](http://msdn.microsoft.com/library/shape.removefromcatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/shape.rerouteconnections-method-publisher%28Office.15%29.aspx)|
|[SaveAsBuildingBlock](http://msdn.microsoft.com/library/shape.saveasbuildingblock-method-publisher%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/shape.saveaspicture-method-publisher%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/shape.scaleheight-method-publisher%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/shape.scalewidth-method-publisher%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/shape.select-method-publisher%28Office.15%29.aspx)|
|[SetCaption](http://msdn.microsoft.com/library/shape.setcaption-method-publisher%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/shape.setshapesdefaultproperties-method-publisher%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/shape.ungroup-method-publisher%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/shape.zorder-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Adjustments](http://msdn.microsoft.com/library/shape.adjustments-property-publisher%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/shape.alternativetext-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/shape.application-property-publisher%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/shape.autoshapetype-property-publisher%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/shape.blackwhitemode-property-publisher%28Office.15%29.aspx)|
|[BorderArt](http://msdn.microsoft.com/library/shape.borderart-property-publisher%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/shape.callout-property-publisher%28Office.15%29.aspx)|
|[CatalogMergeItems](http://msdn.microsoft.com/library/shape.catalogmergeitems-property-publisher%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/shape.connectionsitecount-property-publisher%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/shape.connector-property-publisher%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/shape.connectorformat-property-publisher%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/shape.fill-property-publisher%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/shape.glow-property-publisher%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/shape.groupitems-property-publisher%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/shape.hastable-property-publisher%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/shape.hastextframe-property-publisher%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/shape.height-property-publisher%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/shape.horizontalflip-property-publisher%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/shape.hyperlink-property-publisher%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/shape.id-property-publisher%28Office.15%29.aspx)|
|[InlineAlignment](http://msdn.microsoft.com/library/shape.inlinealignment-property-publisher%28Office.15%29.aspx)|
|[InlineTextRange](http://msdn.microsoft.com/library/shape.inlinetextrange-property-publisher%28Office.15%29.aspx)|
|[IsExcess](http://msdn.microsoft.com/library/shape.isexcess-property-publisher%28Office.15%29.aspx)|
|[IsGroupMember](http://msdn.microsoft.com/library/shape.isgroupmember-property-publisher%28Office.15%29.aspx)|
|[IsInline](http://msdn.microsoft.com/library/shape.isinline-property-publisher%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/shape.left-property-publisher%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/shape.line-property-publisher%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/shape.linkformat-property-publisher%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/shape.lockaspectratio-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/shape.name-property-publisher%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/shape.nodes-property-publisher%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/shape.oleformat-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shape.parent-property-publisher%28Office.15%29.aspx)|
|[ParentGroupShape](http://msdn.microsoft.com/library/shape.parentgroupshape-property-publisher%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/shape.pictureformat-property-publisher%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/shape.reflection-property-publisher%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/shape.rotation-property-publisher%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/shape.shadow-property-publisher%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/shape.softedge-property-publisher%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/shape.table-property-publisher%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/shape.tags-property-publisher%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/shape.texteffect-property-publisher%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/shape.textframe-property-publisher%28Office.15%29.aspx)|
|[TextWrap](http://msdn.microsoft.com/library/shape.textwrap-property-publisher%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/shape.threed-property-publisher%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/shape.top-property-publisher%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/shape.type-property-publisher%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/shape.verticalflip-property-publisher%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/shape.vertices-property-publisher%28Office.15%29.aspx)|
|[WebCheckBox](http://msdn.microsoft.com/library/shape.webcheckbox-property-publisher%28Office.15%29.aspx)|
|[WebCommandButton](http://msdn.microsoft.com/library/shape.webcommandbutton-property-publisher%28Office.15%29.aspx)|
|[WebListBox](http://msdn.microsoft.com/library/shape.weblistbox-property-publisher%28Office.15%29.aspx)|
|[WebNavigationBarSetName](http://msdn.microsoft.com/library/shape.webnavigationbarsetname-property-publisher%28Office.15%29.aspx)|
|[WebOptionButton](http://msdn.microsoft.com/library/shape.weboptionbutton-property-publisher%28Office.15%29.aspx)|
|[WebTextBox](http://msdn.microsoft.com/library/shape.webtextbox-property-publisher%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/shape.width-property-publisher%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/shape.wizard-property-publisher%28Office.15%29.aspx)|
|[WizardTag](http://msdn.microsoft.com/library/shape.wizardtag-property-publisher%28Office.15%29.aspx)|
|[WizardTagInstance](http://msdn.microsoft.com/library/shape.wizardtaginstance-property-publisher%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/shape.zorderposition-property-publisher%28Office.15%29.aspx)|

