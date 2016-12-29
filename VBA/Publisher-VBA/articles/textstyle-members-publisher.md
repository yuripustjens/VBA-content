---
title: TextStyle Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 32662844-441e-ac73-0cb4-a246af9a394b
ms.locale: en-US
---


# TextStyle Members (Publisher)
Represents a single built-in or user-defined style. The  **TextStyle** object includes style attributes (font, font style, paragraph spacing, and so on) as properties of the **TextStyle** object. The **TextStyle** object is a member of the **[TextStyles](textstyles-object-publisher.md)** collection. The  **TextStyles** collection includes all the styles in the specified document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Delete](textstyle.delete-method-publisher.md)|Deletes the specified object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](textstyle.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BaseStyle](textstyle.basestyle-property-publisher.md)|Returns or sets a  **String** that represents the style upon which the formatting of another style is based. Read/write.|
| [Description](textstyle.description-property-publisher.md)|Returns a  **String** that represents the description of the specified style. For example, a typical description for the Normal style might be "(Default) Times New Roman, (Asian) MS Mincho, 10 pt, Main (Black), Kerning 14 pt, Left, Line spacing 1 sp." Read-only.|
| [Font](textstyle.font-property-publisher.md)|Sets or returns a  **[Font](font-object-publisher.md)** object that represents character formatting attributes applied to the specified object. Read/write.|
| [Name](textstyle.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [NextParagraphStyle](textstyle.nextparagraphstyle-property-publisher.md)|Returns or sets a  **String** that represents the paragraph style that follows the specified text style when a user presses ENTER. Read/write.|
| [ParagraphFormat](textstyle.paragraphformat-property-publisher.md)|Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object representing the paragraph formatting for the specified text range or text style.|
| [Parent](textstyle.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|

