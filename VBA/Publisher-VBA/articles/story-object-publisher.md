---
title: Story Object (Publisher)
keywords: vbapb10.chm5898239
f1_keywords: vbapb10.chm5898239
ms.prod: PUBLISHER
api_name: Publisher.Story
ms.assetid: 0385b4be-0046-9198-a186-0d992601780e
ms.locale: en-US
---


# Story Object (Publisher)

Represents the text in an unlinked text frame, text flowing between linked text frames, or text in a table cell. The  **Story** object is a member of the **TextFrame** and **TextRange** objects and the **Stories** collection.


## Example

Use the  **Story** property to return the **Story** object in a text range or text frame. This example returns the story in the selected text range and, if it is in a text frame, inserts text into the text range.


```
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then .TextRange _ 
 .InsertAfter NewText:=vbLf &amp; "This is a test." 
 End With 
End Sub
```

Use  **Stories** (index), where index is the number of the story, to return an individual **Story** object. This example determines if the first story in the active publication has a text frame and, if it does, formats the paragraphs in the story with a half inch first line indent and a six-point spacing before each paragraph.




```
Sub StoryParagraphFirstLineIndent() 
 With ActiveDocument.Stories(1) 
 If .HasTextFrame Then 
 With .TextFrame.TextRange.ParagraphFormat 
 .FirstLineIndent = InchesToPoints(0.5) 
 .SpaceBefore = 6 
 End With 
 End If 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/story.application-property-publisher%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/story.hastable-property-publisher%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/story.hastextframe-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/story.parent-property-publisher%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/story.table-property-publisher%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/story.textframe-property-publisher%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/story.textrange-property-publisher%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/story.type-property-publisher%28Office.15%29.aspx)|

