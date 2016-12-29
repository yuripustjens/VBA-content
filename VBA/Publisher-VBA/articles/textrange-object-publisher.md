---
title: TextRange Object (Publisher)
keywords: vbapb10.chm5373951
f1_keywords: vbapb10.chm5373951
ms.prod: PUBLISHER
api_name: Publisher.TextRange
ms.assetid: 566f240b-d2a6-8cb3-9eb7-68328d6c28bd
ms.locale: en-US
---


# TextRange Object (Publisher)

Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text. This topic describes how to: 


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Insert text, the date and time, or the page number into a text range.
    

## Example

Use the  **[TextRange](http://msdn.microsoft.com/library/textframe.textrange-property-publisher%28Office.15%29.aspx)** property of the **[TextFrame](textframe-object-publisher.md)** object to return a **TextRange** object for any shape you specify. Use the **[Text](http://msdn.microsoft.com/library/textrange.text-property-publisher%28Office.15%29.aspx)** property to return the string of text in the **TextRange** object. The following example adds a rectangle to the active publication and sets the text it contains.


```
Sub AddTextToShape() 
    With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
        Left:=72, Top:=72, Width:=250, Height:=140) 
        .TextFrame.TextRange.Text = "Here is some test text" 
    End With 
End Sub
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange.text = "Here is some test text" 
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange = "Here is some test text"
```

Use the  **[HasTextFrame](http://msdn.microsoft.com/library/shaperange.hastextframe-property-publisher%28Office.15%29.aspx)** property to determine whether a shape has a text frame, and use the **[HasText](http://msdn.microsoft.com/library/textframe.hastext-property-publisher%28Office.15%29.aspx)** property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```
Sub CopyAndPasteText() 
    With ActiveDocument 
        .Selection.TextRange.Copy 
        .Pages(1).Shapes(1).TextFrame.TextRange.Paste 
    End With 
End Sub
```

Use one of the following methods to return a portion of the text of a  **TextRange** object: **[Characters](http://msdn.microsoft.com/library/textrange.characters-method-publisher%28Office.15%29.aspx)**, **[Lines](http://msdn.microsoft.com/library/textrange.lines-method-publisher%28Office.15%29.aspx)**, **[Paragraphs](http://msdn.microsoft.com/library/textrange.paragraphs-method-publisher%28Office.15%29.aspx)**, or **[Words](http://msdn.microsoft.com/library/textrange.words-method-publisher%28Office.15%29.aspx)**. The following example formats the second word in the first shape on the first page of the active publication. For this example to work, the specified shape must contain text.




```
Sub FormatWords() 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange.Words(2).Font 
        .Bold = msoTrue 
        .Size = 15 
        .Name = "Text Name" 
    End With 
End Sub
```

Use one of the following methods to insert characters into a  **TextRange** object: **[InsertAfter](http://msdn.microsoft.com/library/textrange.insertafter-method-publisher%28Office.15%29.aspx)**, **[InsertBefore](http://msdn.microsoft.com/library/textrange.insertbefore-method-publisher%28Office.15%29.aspx)**, **[InsertDateTime](http://msdn.microsoft.com/library/textrange.insertdatetime-method-publisher%28Office.15%29.aspx)**, **[InsertPageNumber](http://msdn.microsoft.com/library/textrange.insertpagenumber-method-publisher%28Office.15%29.aspx)**, or **[InsertSymbol](http://msdn.microsoft.com/library/textrange.insertsymbol-method-publisher%28Office.15%29.aspx)**. This example inserts a new line with text after any existing text in the first shape on the first page of the active publication.




```
Sub InsertNewText() 
    Dim intCount As Integer 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange 
        For intCount = 1 To 3 
            .InsertAfter vbLf &amp; "This is a test." 
        Next intCount 
    End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Characters](http://msdn.microsoft.com/library/textrange.characters-method-publisher%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/textrange.collapse-method-publisher%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/textrange.copy-method-publisher%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/textrange.cut-method-publisher%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/textrange.delete-method-publisher%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/textrange.expand-method-publisher%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/textrange.insertafter-method-publisher%28Office.15%29.aspx)|
|[InsertBarcode](http://msdn.microsoft.com/library/textrange.insertbarcode-method-publisher%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/textrange.insertbefore-method-publisher%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/textrange.insertdatetime-method-publisher%28Office.15%29.aspx)|
|[InsertMailMergeField](http://msdn.microsoft.com/library/textrange.insertmailmergefield-method-publisher%28Office.15%29.aspx)|
|[InsertPageNumber](http://msdn.microsoft.com/library/textrange.insertpagenumber-method-publisher%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/textrange.insertsymbol-method-publisher%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/textrange.lines-method-publisher%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/textrange.move-method-publisher%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/textrange.moveend-method-publisher%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/textrange.movestart-method-publisher%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/textrange.paragraphs-method-publisher%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/textrange.paste-method-publisher%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/textrange.select-method-publisher%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/textrange.words-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/textrange.application-property-publisher%28Office.15%29.aspx)|
|[BoundHeight](http://msdn.microsoft.com/library/textrange.boundheight-property-publisher%28Office.15%29.aspx)|
|[BoundLeft](http://msdn.microsoft.com/library/textrange.boundleft-property-publisher%28Office.15%29.aspx)|
|[BoundTop](http://msdn.microsoft.com/library/textrange.boundtop-property-publisher%28Office.15%29.aspx)|
|[BoundWidth](http://msdn.microsoft.com/library/textrange.boundwidth-property-publisher%28Office.15%29.aspx)|
|[ContainingObject](http://msdn.microsoft.com/library/textrange.containingobject-property-publisher%28Office.15%29.aspx)|
|[DropCap](http://msdn.microsoft.com/library/textrange.dropcap-property-publisher%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/textrange.duplicate-property-publisher%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/textrange.end-property-publisher%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/textrange.fields-property-publisher%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/textrange.find-property-publisher%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/textrange.font-property-publisher%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/textrange.hyperlinks-property-publisher%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/textrange.inlineshapes-property-publisher%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/textrange.languageid-property-publisher%28Office.15%29.aspx)|
|[Length](http://msdn.microsoft.com/library/textrange.length-property-publisher%28Office.15%29.aspx)|
|[LinesCount](http://msdn.microsoft.com/library/textrange.linescount-property-publisher%28Office.15%29.aspx)|
|[MajorityFont](http://msdn.microsoft.com/library/textrange.majorityfont-property-publisher%28Office.15%29.aspx)|
|[MajorityParagraphFormat](http://msdn.microsoft.com/library/textrange.majorityparagraphformat-property-publisher%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/textrange.paragraphformat-property-publisher%28Office.15%29.aspx)|
|[ParagraphsCount](http://msdn.microsoft.com/library/textrange.paragraphscount-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/textrange.parent-property-publisher%28Office.15%29.aspx)|
|[Script](http://msdn.microsoft.com/library/textrange.script-property-publisher%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/textrange.start-property-publisher%28Office.15%29.aspx)|
|[Story](http://msdn.microsoft.com/library/textrange.story-property-publisher%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/textrange.text-property-publisher%28Office.15%29.aspx)|
|[WordsCount](http://msdn.microsoft.com/library/textrange.wordscount-property-publisher%28Office.15%29.aspx)|

