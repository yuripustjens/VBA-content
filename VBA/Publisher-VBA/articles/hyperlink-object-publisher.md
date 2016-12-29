---
title: Hyperlink Object (Publisher)
keywords: vbapb10.chm4653055
f1_keywords: vbapb10.chm4653055
ms.prod: PUBLISHER
api_name: Publisher.Hyperlink
ms.assetid: 1cc6d95b-357a-c169-a5d2-6850a1a3bbd6
ms.locale: en-US
---


# Hyperlink Object (Publisher)

Represents a hyperlink. The  **Hyperlink** object is a member of the **[Hyperlinks](hyperlinks-object-publisher.md)** collection and the **[Shape](http://msdn.microsoft.com/library/shape-object-publisher%28Office.15%29.aspx)** and **[ShapeRange](shaperange-object-publisher.md)** objects.


## Example

Use the  **[Hyperlink](http://msdn.microsoft.com/library/shape.hyperlink-property-publisher%28Office.15%29.aspx)** property to return a **Hyperlink** object associated with a shape (a shape can have only one hyperlink). The following example deletes the hyperlink associated with the first shape in the active document.


```
Sub DeleteHyperlink() 
 ActiveDocument.Pages(1).Shapes(1).Hyperlink.Delete 
End Sub
```

Use  **Hyperlinks** (index), where index is the index number, to return a single **Hyperlink** object from a document, range, or selection. The following example deletes the first hyperlink in the selection.




```
Sub DeleteSelectedHyperlink() 
 If Selection.TextRange.Hyperlinks.Count >= 1 Then 
 Selection.TextRange.Hyperlinks(1).Delete 
 End If 
End Sub
```

Use the  **[Add](http://msdn.microsoft.com/library/hyperlinks.add-method-publisher%28Office.15%29.aspx)** method to add a hyperlink. The following example adds a hyperlink to the selected text.




```
Sub AddHyperlinkToSelectedText() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="http://www.tailspintoys.com/" 
End Sub
```

Use the  **[Address](http://msdn.microsoft.com/library/hyperlink.address-property-publisher%28Office.15%29.aspx)** property to add or change the address to a hyperlink. The following example adds a shape to the active publication and then adds a hyperlink to the shape.




```
Sub AddHyperlinkToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=200, Width:=300, Height:=300) 
 .Hyperlink.Address = "http://www.tailspintoys.com/" 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/hyperlink.delete-method-publisher%28Office.15%29.aspx)|
|[SetPageRelative](http://msdn.microsoft.com/library/hyperlink.setpagerelative-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/hyperlink.address-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/hyperlink.application-property-publisher%28Office.15%29.aspx)|
|[EmailSubject](http://msdn.microsoft.com/library/hyperlink.emailsubject-property-publisher%28Office.15%29.aspx)|
|[PageID](http://msdn.microsoft.com/library/hyperlink.pageid-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/hyperlink.parent-property-publisher%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/hyperlink.range-property-publisher%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/hyperlink.shape-property-publisher%28Office.15%29.aspx)|
|[TargetType](http://msdn.microsoft.com/library/hyperlink.targettype-property-publisher%28Office.15%29.aspx)|
|[TextToDisplay](http://msdn.microsoft.com/library/hyperlink.texttodisplay-property-publisher%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/hyperlink.type-property-publisher%28Office.15%29.aspx)|

