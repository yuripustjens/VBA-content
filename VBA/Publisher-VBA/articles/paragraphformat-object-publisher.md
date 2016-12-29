---
title: ParagraphFormat Object (Publisher)
keywords: vbapb10.chm5505023
f1_keywords: vbapb10.chm5505023
ms.prod: PUBLISHER
api_name: Publisher.ParagraphFormat
ms.assetid: 0e5b1c20-564e-ef5c-f24d-1143dcaadcd8
ms.locale: en-US
---


# ParagraphFormat Object (Publisher)

Represents all the formatting for a paragraph.


## Example

Use the  **[ParagraphFormat](http://msdn.microsoft.com/library/textstyle.paragraphformat-property-publisher%28Office.15%29.aspx)** property to return the **ParagraphFormat** object for a paragraph or paragraphs. The **ParagraphFormat** property returns the **ParagraphFormat** object for a selection, range, or style. The following example centers the paragraph at the cursor position. This example assumes that the first shape is a text box and not another type of shape.


```
Sub CenterParagraph() 
 Selection.TextRange.ParagraphFormat _ 
 .Alignment = pbParagraphAlignmentCenter 
End Sub
```

Use the  **[Duplicate](http://msdn.microsoft.com/library/textrange.duplicate-property-publisher%28Office.15%29.aspx)** property to copy an existing **ParagraphFormat** object. The following example duplicates the paragraph formatting of the first paragraph in the active publication and stores the formatting in a variable. This example duplicates an existing **ParagraphFormat** object and then changes the left indent to one inch, creates a new textbox, inserts text into it, and applies the paragraph formatting of the duplicated paragraph format to the text.




```
Sub DuplicateParagraphFormating() 
 Dim pfmtDup As ParagraphFormat 
 
 Set pfmtDup = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 
 pfmtDup.LeftIndent = Application.InchesToPoints(1) 
 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddTextbox(pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=200, Height:=100) 
 With .TextFrame.TextRange 
 .Text = "This is a test of how to use " &amp; _ 
 "the ParagraphFormat object." 
 .ParagraphFormat = pfmtDup 
 End With 
 End With 
 End With 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Duplicate](http://msdn.microsoft.com/library/paragraphformat.duplicate-method-publisher%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/paragraphformat.reset-method-publisher%28Office.15%29.aspx)|
|[SetLineSpacing](http://msdn.microsoft.com/library/paragraphformat.setlinespacing-method-publisher%28Office.15%29.aspx)|
|[SetListType](http://msdn.microsoft.com/library/paragraphformat.setlisttype-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Alignment](http://msdn.microsoft.com/library/paragraphformat.alignment-property-publisher%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/paragraphformat.application-property-publisher%28Office.15%29.aspx)|
|[AttachedToText](http://msdn.microsoft.com/library/paragraphformat.attachedtotext-property-publisher%28Office.15%29.aspx)|
|[CharBasedFirstLineIndent](http://msdn.microsoft.com/library/paragraphformat.charbasedfirstlineindent-property-publisher%28Office.15%29.aspx)|
|[FirstLineIndent](http://msdn.microsoft.com/library/paragraphformat.firstlineindent-property-publisher%28Office.15%29.aspx)|
|[KashidaPercentage](http://msdn.microsoft.com/library/paragraphformat.kashidapercentage-property-publisher%28Office.15%29.aspx)|
|[KeepLinesTogether](http://msdn.microsoft.com/library/paragraphformat.keeplinestogether-property-publisher%28Office.15%29.aspx)|
|[KeepWithNext](http://msdn.microsoft.com/library/paragraphformat.keepwithnext-property-publisher%28Office.15%29.aspx)|
|[LeftIndent](http://msdn.microsoft.com/library/paragraphformat.leftindent-property-publisher%28Office.15%29.aspx)|
|[LineSpacing](http://msdn.microsoft.com/library/paragraphformat.linespacing-property-publisher%28Office.15%29.aspx)|
|[LineSpacingRule](http://msdn.microsoft.com/library/paragraphformat.linespacingrule-property-publisher%28Office.15%29.aspx)|
|[ListBulletFontName](http://msdn.microsoft.com/library/paragraphformat.listbulletfontname-property-publisher%28Office.15%29.aspx)|
|[ListBulletFontSize](http://msdn.microsoft.com/library/paragraphformat.listbulletfontsize-property-publisher%28Office.15%29.aspx)|
|[ListBulletText](http://msdn.microsoft.com/library/paragraphformat.listbullettext-property-publisher%28Office.15%29.aspx)|
|[ListIndent](http://msdn.microsoft.com/library/paragraphformat.listindent-property-publisher%28Office.15%29.aspx)|
|[ListNumberSeparator](http://msdn.microsoft.com/library/paragraphformat.listnumberseparator-property-publisher%28Office.15%29.aspx)|
|[ListNumberStart](http://msdn.microsoft.com/library/paragraphformat.listnumberstart-property-publisher%28Office.15%29.aspx)|
|[ListType](http://msdn.microsoft.com/library/paragraphformat.listtype-property-publisher%28Office.15%29.aspx)|
|[LockToBaseLine](http://msdn.microsoft.com/library/paragraphformat.locktobaseline-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/paragraphformat.parent-property-publisher%28Office.15%29.aspx)|
|[RightIndent](http://msdn.microsoft.com/library/paragraphformat.rightindent-property-publisher%28Office.15%29.aspx)|
|[SpaceAfter](http://msdn.microsoft.com/library/paragraphformat.spaceafter-property-publisher%28Office.15%29.aspx)|
|[SpaceBefore](http://msdn.microsoft.com/library/paragraphformat.spacebefore-property-publisher%28Office.15%29.aspx)|
|[StartInNextTextBox](http://msdn.microsoft.com/library/paragraphformat.startinnexttextbox-property-publisher%28Office.15%29.aspx)|
|[Tabs](http://msdn.microsoft.com/library/paragraphformat.tabs-property-publisher%28Office.15%29.aspx)|
|[TextDirection](http://msdn.microsoft.com/library/paragraphformat.textdirection-property-publisher%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/paragraphformat.textstyle-property-publisher%28Office.15%29.aspx)|
|[UseCharBasedFirstLineIndent](http://msdn.microsoft.com/library/paragraphformat.usecharbasedfirstlineindent-property-publisher%28Office.15%29.aspx)|
|[WidowControl](http://msdn.microsoft.com/library/paragraphformat.widowcontrol-property-publisher%28Office.15%29.aspx)|

