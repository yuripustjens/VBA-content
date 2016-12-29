---
title: Cell Object (Publisher)
keywords: vbapb10.chm5177343
f1_keywords: vbapb10.chm5177343
ms.prod: PUBLISHER
api_name: Publisher.Cell
ms.assetid: 5baafaa6-368e-9eae-30b9-90d2d89d5a5b
ms.locale: en-US
---


# Cell Object (Publisher)

Represents a single table cell. The  **Cell** object is a member of the **[CellRange](http://msdn.microsoft.com/library/cellrange-object-publisher%28Office.15%29.aspx)** collection. The **CellRange** collection represents all the cells in the specified object.


## Example

Use  **Cells** (index), where index is the cell number, to return a **Cell** object. This example merges the first two cells of the first column of the specified table.


```
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

This example applies a thick border around the first cell in the second column of the specified table.




```
Sub OutlineBorderCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(2).Cells(1) 
 .BorderLeft.Weight = 5 
 .BorderRight.Weight = 5 
 .BorderTop.Weight = 5 
 .BorderBottom.Weight = 5 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Merge](http://msdn.microsoft.com/library/cell.merge-method-publisher%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/cell.select-method-publisher%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/cell.split-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cell.application-property-publisher%28Office.15%29.aspx)|
|[BorderBottom](http://msdn.microsoft.com/library/cell.borderbottom-property-publisher%28Office.15%29.aspx)|
|[BorderDiagonal](http://msdn.microsoft.com/library/cell.borderdiagonal-property-publisher%28Office.15%29.aspx)|
|[BorderLeft](http://msdn.microsoft.com/library/cell.borderleft-property-publisher%28Office.15%29.aspx)|
|[BorderRight](http://msdn.microsoft.com/library/cell.borderright-property-publisher%28Office.15%29.aspx)|
|[BorderTop](http://msdn.microsoft.com/library/cell.bordertop-property-publisher%28Office.15%29.aspx)|
|[CellTextOrientation](http://msdn.microsoft.com/library/cell.celltextorientation-property-publisher%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/cell.column-property-publisher%28Office.15%29.aspx)|
|[Diagonal](http://msdn.microsoft.com/library/cell.diagonal-property-publisher%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/cell.fill-property-publisher%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/cell.hastext-property-publisher%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/cell.height-property-publisher%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/cell.marginbottom-property-publisher%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/cell.marginleft-property-publisher%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/cell.marginright-property-publisher%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/cell.margintop-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cell.parent-property-publisher%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/cell.row-property-publisher%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/cell.selected-property-publisher%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/cell.textrange-property-publisher%28Office.15%29.aspx)|
|[VerticalTextAlignment](http://msdn.microsoft.com/library/cell.verticaltextalignment-property-publisher%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/cell.width-property-publisher%28Office.15%29.aspx)|

