---
title: WebPageOptions Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: 7e3e4207-4f1a-1a6b-e7f6-2fd02a9cbb5d
ms.locale: en-US
---


# WebPageOptions Members (Publisher)
Represents the properties of a single Web page within a Web publication, including options for adding the title and description of the page, background sounds, in addition to other options. The  **WebPageOptions** object is a member of the **[Page](page-object-publisher.md)** object.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [SetBackgroundSoundRepeat](webpageoptions.setbackgroundsoundrepeat-method-publisher.md)|Specifies whether the background sound attached to a Web page should be played infinitely after the page is loaded in a Web browser, and if it should not, optionally specifies the number of times the background sound should be played.|
|Name|Description|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](webpageoptions.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [BackgroundSound](webpageoptions.backgroundsound-property-publisher.md)|Returns or sets a  **String** that specifies the path to a sound file that is played when the Web page is loaded in a Web browser. Read/write.|
| [BackgroundSoundLoopCount](webpageoptions.backgroundsoundloopcount-property-publisher.md)|Returns a  **Long** value that specifies the number of times the background sound attached to a Web page is played when the page is loaded in a Web browser. Read-only.|
| [BackgroundSoundLoopForever](webpageoptions.backgroundsoundloopforever-property-publisher.md)|Returns a  **Boolean** value that specifies whether the background sound attached to the Web page should be repeated infinitely. Read-only.|
| [Description](webpageoptions.description-property-publisher.md)|Returns or sets a  **String** that represents the description of a Web page within a Web publication. Read/write.|
| [IncludePageOnNewWebNavigationBars](webpageoptions.includepageonnewwebnavigationbars-property-publisher.md)|Returns or sets a  **Boolean** value that specifies whether a link to a Web page will be added to the automatic navigation bars of new pages. Read/write.|
| [Keywords](webpageoptions.keywords-property-publisher.md)|Returns or sets a  **String** that represents the keywords for a Web page within a Web publication. Read/write.|
| [Parent](webpageoptions.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [PublishFileName](webpageoptions.publishfilename-property-publisher.md)|Returns or sets a  **String** that represents the file name of a Web page (within a Web publication) that is being saved as filtered HTML.|

