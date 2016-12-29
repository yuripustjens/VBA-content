---
title: Page Object (Publisher)
keywords: vbapb10.chm458751
f1_keywords: vbapb10.chm458751
ms.prod: PUBLISHER
api_name: Publisher.Page
ms.assetid: 9b2e8f29-26c3-1008-0ffd-eea2147abca4
ms.locale: en-US
---


# Page Object (Publisher)

Represents a page in a publication. The  **[Pages](http://msdn.microsoft.com/library/pages-object-publisher%28Office.15%29.aspx)** collection contains all the **Page** objects in a publication.


## Example

Use  **Pages** (index) to return a single **Page** object. The following example adds new text to the first shape on the first page in the active publication.


```
Sub AddPageNumberField() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .InsertAfter " This text is added after the existing text." 
 .Font.Size = 15 
 End With 
End Sub
```

Use the  **[FindBypageID](http://msdn.microsoft.com/library/pages.findbypageid-method-publisher%28Office.15%29.aspx)** property to locate a **Page** object using the application assigned page ID. Use the **[Add](http://msdn.microsoft.com/library/pages.add-method-publisher%28Office.15%29.aspx)** method to create a new page and add it to the publication. The following example adds a new page to the active publication and then looks for that page using the page ID.




```
Sub FindPage() 
 Dim lngPageID As Long 
 
 'Get page ID 
 lngPageID = ActiveDocument.Pages.Add(Count:=1, After:=1).PageID 
 
 'Use page ID to add a new shape to the page 
 ActiveDocument.Pages.FindByPageID(PageID:=lngPageID) _ 
 .Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=200, Top:=72, Width:=50, Height:=50 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/page.delete-method-publisher%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/page.duplicate-method-publisher%28Office.15%29.aspx)|
|[ExportEmailHTML](http://msdn.microsoft.com/library/page.exportemailhtml-method-publisher%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/page.move-method-publisher%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/page.saveaspicture-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/page.application-property-publisher%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/page.background-property-publisher%28Office.15%29.aspx)|
|[Footer](http://msdn.microsoft.com/library/page.footer-property-publisher%28Office.15%29.aspx)|
|[Header](http://msdn.microsoft.com/library/page.header-property-publisher%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/page.height-property-publisher%28Office.15%29.aspx)|
|[IgnoreMaster](http://msdn.microsoft.com/library/page.ignoremaster-property-publisher%28Office.15%29.aspx)|
|[IsLeading](http://msdn.microsoft.com/library/page.isleading-property-publisher%28Office.15%29.aspx)|
|[IsTrailing](http://msdn.microsoft.com/library/page.istrailing-property-publisher%28Office.15%29.aspx)|
|[IsTwoPageMaster](http://msdn.microsoft.com/library/page.istwopagemaster-property-publisher%28Office.15%29.aspx)|
|[IsWizardPage](http://msdn.microsoft.com/library/page.iswizardpage-property-publisher%28Office.15%29.aspx)|
|[LayoutGuides](http://msdn.microsoft.com/library/page.layoutguides-property-publisher%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/page.master-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/page.name-property-publisher%28Office.15%29.aspx)|
|[PageID](http://msdn.microsoft.com/library/page.pageid-property-publisher%28Office.15%29.aspx)|
|[PageIndex](http://msdn.microsoft.com/library/page.pageindex-property-publisher%28Office.15%29.aspx)|
|[PageNumber](http://msdn.microsoft.com/library/page.pagenumber-property-publisher%28Office.15%29.aspx)|
|[PageType](http://msdn.microsoft.com/library/page.pagetype-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/page.parent-property-publisher%28Office.15%29.aspx)|
|[ReaderSpread](http://msdn.microsoft.com/library/page.readerspread-property-publisher%28Office.15%29.aspx)|
|[RulerGuides](http://msdn.microsoft.com/library/page.rulerguides-property-publisher%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/page.shapes-property-publisher%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/page.tags-property-publisher%28Office.15%29.aspx)|
|[WebPageOptions](http://msdn.microsoft.com/library/page.webpageoptions-property-publisher%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/page.width-property-publisher%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/page.wizard-property-publisher%28Office.15%29.aspx)|
|[XOffsetWithinReaderSpread](http://msdn.microsoft.com/library/page.xoffsetwithinreaderspread-property-publisher%28Office.15%29.aspx)|
|[YOffsetWithinReaderSpread](http://msdn.microsoft.com/library/page.yoffsetwithinreaderspread-property-publisher%28Office.15%29.aspx)|

