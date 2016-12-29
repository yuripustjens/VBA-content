---
title: TextRange.LanguageID Property (Publisher)
keywords: vbapb10.chm5308471
f1_keywords: vbapb10.chm5308471
ms.prod: PUBLISHER
api_name: Publisher.TextRange.LanguageID
ms.assetid: 1007c821-cafd-0cb3-94f4-4ac25decad30
ms.locale: en-US
---


# TextRange.LanguageID Property (Publisher)

Returns or sets a  **MsoLanguageID** constant that represents the language for the specified object. Read/write.


## Syntax

 _expression_. **LanguageID**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

MsoLanguageID


## Remarks

The  **LanguageID** property value can be one of the ** [MsoLanguageID](http://msdn.microsoft.com/library/msolanguageid-enumeration-office%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example formats the specified selection as French. This example assumes that the cursor is in a text box.


```vb
Sub SetLanguage() 
 Selection.TextRange.LanguageID = msoLanguageIDFrench 
End Sub
```


