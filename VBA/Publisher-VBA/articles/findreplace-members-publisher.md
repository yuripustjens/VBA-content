---
title: FindReplace Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 14fa0975-7499-80c3-aa8b-e7cf52fe40e7
ms.locale: en-US
---


# FindReplace Members (Publisher)
Represents the criteria for a find operation. The properties and methods of the  **FindReplace** object correspond to the options in the **Find and Replace** dialog box.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Clear](findreplace.clear-method-publisher.md)|Removes the specified search criteria in a find or replace operation.|
| [Execute](findreplace.execute-method-publisher.md)|Performs the specified Find or Replace operation.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](findreplace.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [FindText](findreplace.findtext-property-publisher.md)|Sets or retrieves a  **String** representing the text to find in the specified range or selection. Read/write.|
| [Forward](findreplace.forward-property-publisher.md)|Sets or retrieves a  **Boolean** representing the direction of the text search. **True** if the find operation searches forward through the document. **False** if it searches backward through the document. Read/write.|
| [FoundTextRange](findreplace.foundtextrange-property-publisher.md)|Returns a  **TextRange** object that represents the found text or replaced text of a find operation. Read-only.|
| [MatchAlefHamza](findreplace.matchalefhamza-property-publisher.md)|Sets or returns a  **Boolean** representing whether or not a search operation will match alefs and hamzas. Read/write.|
| [MatchCase](findreplace.matchcase-property-publisher.md)|Sets or returns a  **Boolean** that represents the case sensitivity of the search operation. Read/write.|
| [MatchDiacritics](findreplace.matchdiacritics-property-publisher.md)|Sets or returns a  **Boolean** representing whether or not a search operation will match diacritics. Read/write.|
| [MatchKashida](findreplace.matchkashida-property-publisher.md)|Sets or returns a  **Boolean** representing whether or not a search operation will match kashidas. Read/write.|
| [MatchWholeWord](findreplace.matchwholeword-property-publisher.md)|Sets or returns a  **Boolean** that represents whether the whole word will be matched in the search operation. Read/write. **Boolean**.|
| [MatchWidth](findreplace.matchwidth-property-publisher.md)|Sets or returns a  **Boolean** representing whether or not a search operation will match the character width of the searched text. Read/Write.|
| [Parent](findreplace.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ReplaceScope](findreplace.replacescope-property-publisher.md)|Specifies how many scope replacements are to be made: one, all, or none. Read/write.|
| [ReplaceWithText](findreplace.replacewithtext-property-publisher.md)|Sets or retrieves a  **String** representing the replacement text in the specified range or selection. Read/write.|

