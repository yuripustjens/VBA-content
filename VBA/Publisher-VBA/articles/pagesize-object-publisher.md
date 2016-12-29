---
title: PageSize Object (Publisher)
keywords: vbapb10.chm8912895
f1_keywords: vbapb10.chm8912895
ms.prod: PUBLISHER
api_name: Publisher.PageSize
ms.assetid: 80767524-6f0c-0d3f-388a-a38891b2d04a
ms.locale: en-US
---


# PageSize Object (Publisher)

Represents the page size of the current Microsoft Publisher publication.


## Remarks

The page size represented by the  **PageSize** object corresponds to one of the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Name** property of the **PageSize** object to get a list of the names of all the page sizes available in the current document and print the list in the **Immediate** window.


```
Public Sub PageSizes_Example() 
 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim pubPageSize As Publisher.PageSize 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 For Each pubPageSize In pubPageSizes 
 Debug.Print pubPageSize.Name 
 Next 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pagesize.application-property-publisher%28Office.15%29.aspx)|
|[HasBackgroundImage](http://msdn.microsoft.com/library/pagesize.hasbackgroundimage-property-publisher%28Office.15%29.aspx)|
|[HorizontalGap](http://msdn.microsoft.com/library/pagesize.horizontalgap-property-publisher%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/pagesize.leftmargin-property-publisher%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/pagesize.name-property-publisher%28Office.15%29.aspx)|
|[PageHeight](http://msdn.microsoft.com/library/pagesize.pageheight-property-publisher%28Office.15%29.aspx)|
|[PageWidth](http://msdn.microsoft.com/library/pagesize.pagewidth-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/pagesize.parent-property-publisher%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/pagesize.topmargin-property-publisher%28Office.15%29.aspx)|
|[VerticalGap](http://msdn.microsoft.com/library/pagesize.verticalgap-property-publisher%28Office.15%29.aspx)|

