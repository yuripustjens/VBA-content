---
title: Shapes Object (Publisher)
keywords: vbapb10.chm2228223
f1_keywords: vbapb10.chm2228223
ms.prod: PUBLISHER
api_name: Publisher.Shapes
ms.assetid: 52e069a6-d54b-a11a-1cba-96174329cb02
ms.locale: en-US
---


# Shapes Object (Publisher)

A collection of  **[Shape](http://msdn.microsoft.com/library/shape-object-publisher%28Office.15%29.aspx)** objects that represent all the shapes on a page of a publication. Each **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


 **Note**  If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](shaperange-object-publisher.md)** collection that contains the shapes with which you want to work.


## Example

Use the  **[Shapes](http://msdn.microsoft.com/library/page.shapes-property-publisher%28Office.15%29.aspx)** property to return the **Shapes** collection. The following example selects all the shapes on the first page of the active publication.


```
Sub SelectAllShapes() 
    ActiveDocument.Pages(1).Shapes.SelectAll 
End Sub
```


 **Note**  If you want to do something (like delete or set a property) to all the shapes in a publication at the same time, use the  **[Range](http://msdn.microsoft.com/library/shapes.range-method-publisher%28Office.15%29.aspx)** method to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.



Use one of the following methods of the  **Shapes** collection: **[AddCallout](http://msdn.microsoft.com/library/shapes.addcallout-method-publisher%28Office.15%29.aspx)**, **[AddConnector](http://msdn.microsoft.com/library/shapes.addconnector-method-publisher%28Office.15%29.aspx)**, **[AddCurve](http://msdn.microsoft.com/library/shapes.addcurve-method-publisher%28Office.15%29.aspx)**, **[AddLabel](http://msdn.microsoft.com/library/shapes.addlabel-method-publisher%28Office.15%29.aspx)**, **[AddLine](http://msdn.microsoft.com/library/shapes.addline-method-publisher%28Office.15%29.aspx)**, **[AddOLEObject](http://msdn.microsoft.com/library/shapes.addoleobject-method-publisher%28Office.15%29.aspx)**, **[AddPolyline](http://msdn.microsoft.com/library/shapes.addpolyline-method-publisher%28Office.15%29.aspx)**, **[AddShape](http://msdn.microsoft.com/library/shapes.addshape-method-publisher%28Office.15%29.aspx)**, **[AddTextbox](http://msdn.microsoft.com/library/shapes.addtextbox-method-publisher%28Office.15%29.aspx)**, or **[AddTextEffect](http://msdn.microsoft.com/library/shapes.addtexteffect-method-publisher%28Office.15%29.aspx)** to add a shape to a publication and return a **Shape** object that represents the newly created shape. The following example adds a new shape to the active publication.




```
Sub AddNewShape() 
    ActiveDocument.Pages(1).Shapes.AddShape Type:=msoShapeFoldedCorner, _ 
        Left:=50, Top:=50, Width:=100, Height:=200 
End Sub
```

Use  **Shapes** (index), where index is the index number, to return a single **Shape** object. The following example horizontally flips shape one on the first page of the active publication.




```
Sub FlipShape() 
    ActiveDocument.Pages(1).Shapes(1).Flip FlipCmd:=msoFlipHorizontal 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddBuildingBlock](http://msdn.microsoft.com/library/shapes.addbuildingblock-method-publisher%28Office.15%29.aspx)|
|[AddCallout](http://msdn.microsoft.com/library/shapes.addcallout-method-publisher%28Office.15%29.aspx)|
|[AddCatalogMergeArea](http://msdn.microsoft.com/library/shapes.addcatalogmergearea-method-publisher%28Office.15%29.aspx)|
|[AddCatalogMergeFieldToCanvas](http://msdn.microsoft.com/library/shapes.addcatalogmergefieldtocanvas-method-publisher%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/shapes.addconnector-method-publisher%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/shapes.addcurve-method-publisher%28Office.15%29.aspx)|
|[AddEmptyPictureFrame](http://msdn.microsoft.com/library/shapes.addemptypictureframe-method-publisher%28Office.15%29.aspx)|
|[AddGroupWizard](http://msdn.microsoft.com/library/shapes.addgroupwizard-method-publisher%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/shapes.addlabel-method-publisher%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/shapes.addline-method-publisher%28Office.15%29.aspx)|
|[AddOLEObject](http://msdn.microsoft.com/library/shapes.addoleobject-method-publisher%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/shapes.addpicture-method-publisher%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/shapes.addpolyline-method-publisher%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/shapes.addshape-method-publisher%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/shapes.addtable-method-publisher%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/shapes.addtextbox-method-publisher%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/shapes.addtexteffect-method-publisher%28Office.15%29.aspx)|
|[AddWebControl](http://msdn.microsoft.com/library/shapes.addwebcontrol-method-publisher%28Office.15%29.aspx)|
|[AddWebNavigationBar](http://msdn.microsoft.com/library/shapes.addwebnavigationbar-method-publisher%28Office.15%29.aspx)|
|[AddWordArt](http://msdn.microsoft.com/library/shapes.addwordart-method-publisher%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/shapes.buildfreeform-method-publisher%28Office.15%29.aspx)|
|[FindShapeByWizardTag](http://msdn.microsoft.com/library/shapes.findshapebywizardtag-method-publisher%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/shapes.item-method-publisher%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/shapes.paste-method-publisher%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/shapes.range-method-publisher%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/shapes.selectall-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/shapes.application-property-publisher%28Office.15%29.aspx)|
|[CanvasArrangementType](http://msdn.microsoft.com/library/shapes.canvasarrangementtype-property-publisher%28Office.15%29.aspx)|
|[CanvasesCount](http://msdn.microsoft.com/library/shapes.canvasescount-property-publisher%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/shapes.count-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shapes.parent-property-publisher%28Office.15%29.aspx)|

