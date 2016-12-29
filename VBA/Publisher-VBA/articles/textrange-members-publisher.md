---
title: TextRange Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: b565a83f-a2a9-8e2a-f7d0-28fbe5435e20
ms.locale: en-US
---


# TextRange Members (Publisher)
Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text. This topic describes how to: 

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Characters](textrange.characters-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the specified subset of text characters.|
| [Collapse](textrange.collapse-method-publisher.md)|Collapses a range or selection to the starting position or ending position. After a range or selection is collapsed, the starting point and the ending point are equal.|
| [Copy](textrange.copy-method-publisher.md)|Copies the specified object to the Clipboard.|
| [Cut](textrange.cut-method-publisher.md)|Deletes the specified object and places it on the Clipboard.|
| [Delete](textrange.delete-method-publisher.md)|Deletes the specified object.|
| [Expand](textrange.expand-method-publisher.md)|Expands the specified range or selection. Returns or sets a  **Long** that represents the number of specified units added to the range or selection.|
| [InsertAfter](textrange.insertafter-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents text appended to the end of a text range.|
| [InsertBarcode](textrange.insertbarcode-method-publisher.md)|Inserts a bar code field at the end of the text range represented by the parent  **TextRange** object.|
| [InsertBefore](textrange.insertbefore-method-publisher.md)|Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.|
| [InsertDateTime](textrange.insertdatetime-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the date and time inserted into a specified text range.|
| [InsertMailMergeField](textrange.insertmailmergefield-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a text data field for a mail merge or catalog merge.|
| [InsertPageNumber](textrange.insertpagenumber-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a page number field in a publication.|
| [InsertSymbol](textrange.insertsymbol-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a symbol inserted in place of the specified range or selection.|
| [Lines](textrange.lines-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the specified lines.|
| [Move](textrange.move-method-publisher.md)|Collapses the specified range to its start position or end position and then moves the collapsed object by the specified number of units. This method returns a  **Long** that represents the number of units by which the object was actually moved, or it returns 0 (zero) if the move was unsuccessful.|
| [MoveEnd](textrange.moveend-method-publisher.md)|Moves the ending character position of a range. This method returns a  **Long** that represents the number of units the range or selection actually moved or returns 0 (zero) if the move was unsuccessful.|
| [MoveStart](textrange.movestart-method-publisher.md)|Moves the start position of the specified range. This method returns a  **Long** that indicates the number of units by which the start position or the range or selection actually moved, or it returns 0 (zero) if the move was unsuccessful.|
| [Paragraphs](textrange.paragraphs-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the specified paragraphs.|
| [Paste](textrange.paste-method-publisher.md)|Pastes the text on the Clipboard into the specified text range, and returns a  **[TextRange](textrange-object-publisher.md)** object that represents the pasted text.|
| [Select](textrange.select-method-publisher.md)|Selects the specified object.|
| [Words](textrange.words-method-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the specified subset of text words.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](textrange.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BoundHeight](textrange.boundheight-property-publisher.md)|Returns a  **Single** indicating the height, in points, of the bounding box for the specified text range. Read-only.|
| [BoundLeft](textrange.boundleft-property-publisher.md)|Returns a  **Single** indicating the distance, in points, from the left edge of the leftmost page to the left edge of the bounding box for the specified text range. Read-only.|
| [BoundTop](textrange.boundtop-property-publisher.md)|Returns a  **Single** indicating the distance, in points, from the top edge of the topmost page to the top edge of the bounding box for the specified text range. Read-only.|
| [BoundWidth](textrange.boundwidth-property-publisher.md)|Returns a  **Single** indicating the width, in points, of the bounding box for the specified text range. Read-only.|
| [ContainingObject](textrange.containingobject-property-publisher.md)|Returns an  **Object** that represents the object that contains the text range. Read-only.|
| [DropCap](textrange.dropcap-property-publisher.md)|Returns a  **[DropCap](dropcap-object-publisher.md)** object that represents a dropped capital letter for the paragraphs in the specified text frame.|
| [Duplicate](textrange.duplicate-property-publisher.md)|Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a duplicate of the specified text range.|
| [End](textrange.end-property-publisher.md)|Sets or returns a  **Long** that represents the ending character position of a selection or text range. Read/write.|
| [Fields](textrange.fields-property-publisher.md)|Returns a  **Fields** object that represents all the fields in the specified text range.|
| [Find](textrange.find-property-publisher.md)|Returns a  **FindReplace** object from the specified **TextRange** object. The **FindReplace** object is used to perform a text search and replace in the specified text range.|
| [Font](textrange.font-property-publisher.md)|Sets or returns a  **[Font](font-object-publisher.md)** object that represents character formatting attributes applied to the specified object. Read/write.|
| [Hyperlinks](textrange.hyperlinks-property-publisher.md)|Returns a  **[Hyperlinks](hyperlinks-object-publisher.md)** collection representing all the hyperlinks in the specified text range.|
| [InlineShapes](textrange.inlineshapes-property-publisher.md)|Returns an  **[InlineShapes](inlineshapes-object-publisher.md)** collection, which represents the inline shapes contained within a text range. Read-only.|
| [LanguageID](textrange.languageid-property-publisher.md)|Returns or sets a  **MsoLanguageID** constant that represents the language for the specified object. Read/write.|
| [Length](textrange.length-property-publisher.md)|Returns a  **Long** value indicating the length of the specified text range, in characters. Read-only.|
| [LinesCount](textrange.linescount-property-publisher.md)|Returns the number of lines of text in the text range represented by the parent  **TextRange** object. Read-only.|
| [MajorityFont](textrange.majorityfont-property-publisher.md)|Returns a  **[Font](font-object-publisher.md)** object that represents the font name most in use in a text range.|
| [MajorityParagraphFormat](textrange.majorityparagraphformat-property-publisher.md)|Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object that represents the paragraph formatting applied to most of the paragraphs in a text range.|
| [ParagraphFormat](textrange.paragraphformat-property-publisher.md)|Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object representing the paragraph formatting for the specified text range or text style.|
| [ParagraphsCount](textrange.paragraphscount-property-publisher.md)|Returns the number of paragraphs of text in the text range represented by the parent  **TextRange** object. Read-only.|
| [Parent](textrange.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Script](textrange.script-property-publisher.md)|Returns a  **PbFontScriptType** constant that represents the font script for a text range. Read-only.|
| [Start](textrange.start-property-publisher.md)|Returns or sets a  **Long** that represents the starting character position of a text range. Read/write.|
| [Story](textrange.story-property-publisher.md)|Returns a  **Story** object that represents the story properties in a text range.|
| [Text](textrange.text-property-publisher.md)|Returns or sets a  **String** that represents the text in a text range or WordArt shape. Read/write.|
| [WordsCount](textrange.wordscount-property-publisher.md)|Returns the number of words in the text range represented by the parent  **TextRange** object. Read-only.|

