---
title: PageSetup Object (Publisher)
keywords: vbapb10.chm7012351
f1_keywords: vbapb10.chm7012351
ms.prod: PUBLISHER
api_name: Publisher.PageSetup
ms.assetid: 23fe3235-88ea-ac93-6d7d-850298263046
ms.locale: en-US
---


# PageSetup Object (Publisher)

Contains information about the page setup for the pages in a publication.


## Example

Use the  **[PageSetup](http://msdn.microsoft.com/library/document.pagesetup-property-publisher%28Office.15%29.aspx)** property to return the **PageSetup** object. The following example sets all pages in the active publication to be 8.5 inches wide and 11 inches high.


```
Sub SetPageSetupOptions() 
 With ActiveDocument.PageSetup 
 .PageHeight = 11 * 72 
 .PageWidth = 8.5 * 72 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pagesetup.application-property-publisher%28Office.15%29.aspx)|
|[AvailablePageSizes](http://msdn.microsoft.com/library/pagesetup.availablepagesizes-property-publisher%28Office.15%29.aspx)|
|[HorizontalGap](http://msdn.microsoft.com/library/pagesetup.horizontalgap-property-publisher%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/pagesetup.leftmargin-property-publisher%28Office.15%29.aspx)|
|[PageHeight](http://msdn.microsoft.com/library/pagesetup.pageheight-property-publisher%28Office.15%29.aspx)|
|[PageSize](http://msdn.microsoft.com/library/pagesetup.pagesize-property-publisher%28Office.15%29.aspx)|
|[PageWidth](http://msdn.microsoft.com/library/pagesetup.pagewidth-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/pagesetup.parent-property-publisher%28Office.15%29.aspx)|
|[PublicationLayout](http://msdn.microsoft.com/library/pagesetup.publicationlayout-property-publisher%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/pagesetup.topmargin-property-publisher%28Office.15%29.aspx)|
|[VerticalGap](http://msdn.microsoft.com/library/pagesetup.verticalgap-property-publisher%28Office.15%29.aspx)|

