---
title: Table Object (Publisher)
keywords: vbapb10.chm4849663
f1_keywords: vbapb10.chm4849663
ms.prod: PUBLISHER
api_name: Publisher.Table
ms.assetid: 09da4a0a-2230-067e-1cac-55321ea044c5
ms.locale: en-US
---


# Table Object (Publisher)

Represents a single table.


## Example

Use the  **[Table](http://msdn.microsoft.com/library/shape.table-property-publisher%28Office.15%29.aspx)** property to return a **Table** object. The following example selects the specified table in the active publication.


```
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then _ 
 .Table.Cells.Select 
 End With 
End Sub
```

Use the  **[AddTable](http://msdn.microsoft.com/library/shapes.addtable-method-publisher%28Office.15%29.aspx)** method to add a **Shape** object representing a table at the specified range. The following example adds a 5x5 table on the first page of the active publication, and then selects the first column of the new table.




```
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(1).Cells.Select 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyAutoFormat](http://msdn.microsoft.com/library/table.applyautoformat-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/table.application-property-publisher%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/table.cells-property-publisher%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/table.columns-property-publisher%28Office.15%29.aspx)|
|[GrowToFitText](http://msdn.microsoft.com/library/table.growtofittext-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/table.parent-property-publisher%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/table.rows-property-publisher%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/table.tabledirection-property-publisher%28Office.15%29.aspx)|

