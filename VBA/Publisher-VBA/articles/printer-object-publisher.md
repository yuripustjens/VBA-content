---
title: Printer Object (Publisher)
keywords: vbapb10.chm9043967
f1_keywords: vbapb10.chm9043967
ms.prod: PUBLISHER
api_name: Publisher.Printer
ms.assetid: 46f8c6a2-4cf1-bb6a-1214-a751440870f2
ms.locale: en-US
---


# Printer Object (Publisher)

A  **Printer** object represents a printer installed on your computer.


## Remarks

Many of the properties, such as  **PaperSize**, **PaperSource**, and **PaperOrientation**, of the **Printer** object correspond to the settings in the **Print Setup** dialog box ( **File** menu) in the Microsoft Publisher user interface .

The collection of all the printers installed on your computer is represented by the  **InstalledPrinters** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **PrinterName** and **IsActivePrinter** properties of the **Printer** object to get a list of all the installed printers on the computer, determine which of them is currently the active printer, and get some of the settings of the active printer. The macro displays the results in the **Immediate** window.


```
Public Sub Printer_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer" 
 Debug.Print "Paper size is ", pubPrinter.PaperSize 
 Debug.Print "Paper orientation is ", pubPrinter.PaperOrientation 
 Debug.Print "Paper source is ", pubPrinter.PaperSource 
 End If 
 Next 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/printer.application-property-publisher%28Office.15%29.aspx)|
|[DriverType](http://msdn.microsoft.com/library/printer.drivertype-property-publisher%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/printer.index-property-publisher%28Office.15%29.aspx)|
|[IsActivePrinter](http://msdn.microsoft.com/library/printer.isactiveprinter-property-publisher%28Office.15%29.aspx)|
|[IsColor](http://msdn.microsoft.com/library/printer.iscolor-property-publisher%28Office.15%29.aspx)|
|[IsDuplex](http://msdn.microsoft.com/library/printer.isduplex-property-publisher%28Office.15%29.aspx)|
|[PaperHeight](http://msdn.microsoft.com/library/printer.paperheight-property-publisher%28Office.15%29.aspx)|
|[PaperOrientation](http://msdn.microsoft.com/library/printer.paperorientation-property-publisher%28Office.15%29.aspx)|
|[PaperSize](http://msdn.microsoft.com/library/printer.papersize-property-publisher%28Office.15%29.aspx)|
|[PaperSource](http://msdn.microsoft.com/library/printer.papersource-property-publisher%28Office.15%29.aspx)|
|[PaperWidth](http://msdn.microsoft.com/library/printer.paperwidth-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/printer.parent-property-publisher%28Office.15%29.aspx)|
|[PrintableRect](http://msdn.microsoft.com/library/printer.printablerect-property-publisher%28Office.15%29.aspx)|
|[PrinterName](http://msdn.microsoft.com/library/printer.printername-property-publisher%28Office.15%29.aspx)|
|[PrintMode](http://msdn.microsoft.com/library/printer.printmode-property-publisher%28Office.15%29.aspx)|

