---
title: TextFrame Object (Publisher)
keywords: vbapb10.chm3932159
f1_keywords: vbapb10.chm3932159
ms.prod: PUBLISHER
api_name: Publisher.TextFrame
ms.assetid: 95e88f5a-b3dc-272e-7c1d-5282c97ae11e
ms.locale: en-US
---


# TextFrame Object (Publisher)

Represents the text frame in a  **[Shape](http://msdn.microsoft.com/library/shape-object-publisher%28Office.15%29.aspx)** object. Contains the text in the text frame and the properties that control the margins and orientation of the text frame.


## Example

Use the  **[TextFrame](http://msdn.microsoft.com/library/shape.textframe-property-publisher%28Office.15%29.aspx)** property to return the **TextFrame** object for a shape. The **[TextRange](http://msdn.microsoft.com/library/textframe.textrange-property-publisher%28Office.15%29.aspx)** property returns a **[TextRange](textrange-object-publisher.md)** object that represents the range of text inside the specified text frame. The following example adds text to the text frame of shape one in the active publication, and then formats the new text.


```
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```


 **Note**  Some shapes do not support attached text (lines, freeforms, pictures, and OLE objects, for example). If you attempt to return or set properties that control text in a text frame for those objects, an error occurs.

Use the  **[HasTextFrame](http://msdn.microsoft.com/library/shape.hastextframe-property-publisher%28Office.15%29.aspx)** property to determine whether the shape has a text frame and use the **[HasText](http://msdn.microsoft.com/library/textframe.hastext-property-publisher%28Office.15%29.aspx)** property to determine whether the text frame contains text, as shown in the following example.




```
Sub GetTextFromTextFrame() 
 Dim shpText As Shape 
 
 For Each shpText In ActiveDocument.Pages(1).Shapes 
 If shpText.HasTextFrame = msoTrue Then 
 With shpText.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
 End If 
 Next 
End Sub
```

Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the  **[NextLinkedTextFrame](http://msdn.microsoft.com/library/textframe.nextlinkedtextframe-property-publisher%28Office.15%29.aspx)** and **[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/textframe.previouslinkedtextframe-property-publisher%28Office.15%29.aspx)** properties to link text frames. The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one.




```
Sub LinkTextBoxes() 
 Dim shpTextBox1 As Shape 
 Dim shpTextBox2 As Shape 
 
 Set shpTextBox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
 shpTextBox1.TextFrame.TextRange.Text = _ 
 "This is some text. This is some more text." 
 
 Set shpTextBox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
 shpTextBox1.TextFrame.NextLinkedTextFrame = shpTextBox2 _ 
 .TextFrame 
End Sub
```


## Methods



|**Name**|
|:-----|
|[BreakForwardLink](http://msdn.microsoft.com/library/textframe.breakforwardlink-method-publisher%28Office.15%29.aspx)|
|[ValidLinkTarget](http://msdn.microsoft.com/library/textframe.validlinktarget-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/textframe.application-property-publisher%28Office.15%29.aspx)|
|[AutoFitText](http://msdn.microsoft.com/library/textframe.autofittext-property-publisher%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/textframe.columns-property-publisher%28Office.15%29.aspx)|
|[ColumnSpacing](http://msdn.microsoft.com/library/textframe.columnspacing-property-publisher%28Office.15%29.aspx)|
|[HasNextLink](http://msdn.microsoft.com/library/textframe.hasnextlink-property-publisher%28Office.15%29.aspx)|
|[HasPreviousLink](http://msdn.microsoft.com/library/textframe.haspreviouslink-property-publisher%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/textframe.hastext-property-publisher%28Office.15%29.aspx)|
|[IncludeContinuedFromPage](http://msdn.microsoft.com/library/textframe.includecontinuedfrompage-property-publisher%28Office.15%29.aspx)|
|[IncludeContinuedOnPage](http://msdn.microsoft.com/library/textframe.includecontinuedonpage-property-publisher%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/textframe.marginbottom-property-publisher%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/textframe.marginleft-property-publisher%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/textframe.marginright-property-publisher%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/textframe.margintop-property-publisher%28Office.15%29.aspx)|
|[NextLinkedTextFrame](http://msdn.microsoft.com/library/textframe.nextlinkedtextframe-property-publisher%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/textframe.orientation-property-publisher%28Office.15%29.aspx)|
|[Overflowing](http://msdn.microsoft.com/library/textframe.overflowing-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/textframe.parent-property-publisher%28Office.15%29.aspx)|
|[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/textframe.previouslinkedtextframe-property-publisher%28Office.15%29.aspx)|
|[Story](http://msdn.microsoft.com/library/textframe.story-property-publisher%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/textframe.textrange-property-publisher%28Office.15%29.aspx)|
|[VerticalTextAlignment](http://msdn.microsoft.com/library/textframe.verticaltextalignment-property-publisher%28Office.15%29.aspx)|

