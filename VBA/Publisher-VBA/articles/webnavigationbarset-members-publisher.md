---
title: WebNavigationBarSet Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: c7ef5285-7d69-b753-b2b6-45db62dc2d42
ms.locale: en-US
---


# WebNavigationBarSet Members (Publisher)
Represents a Web navigation bar set for the current document. The  **WebNavigationBarSet** object is a member of the **WebNavigationBarSets** collection, which includes all of the Web navigation bar sets in the current document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddToEveryPage](webnavigationbarset.addtoeverypage-method-publisher.md)|Adds a  **ShapeRange** of type **pbWebNavigationBar** to each page of the current document.|
| [ChangeOrientation](webnavigationbarset.changeorientation-method-publisher.md)|Sets a  **PbNavBarOrientation** constant that represents the alignment of the navigation bar; vertical or horizontal.|
| [DeleteSetAndInstances](webnavigationbarset.deletesetandinstances-method-publisher.md)|Deletes a Web navigation bar set and all instances of it in the current document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](webnavigationbarset.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoUpdate](webnavigationbarset.autoupdate-property-publisher.md)| **True** if all pages will be added to the specified Web navigation bar set and that adding new pages will update the navigation bar with a corresponding item. Pages must have the **AddHyperlinkToWebNavbar** set to **True** or **WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True** to be added or updated within the specified **WebNavigationBarSet**. Read/write  **Boolean**.|
| [ButtonStyle](webnavigationbarset.buttonstyle-property-publisher.md)|Sets or returns a  **PbWizardNavBarButtonStyle** constant that represents the style of the navigation bar buttons: large, small, or text-only. Read/write.|
| [Design](webnavigationbarset.design-property-publisher.md)|Sets or returns a  **PbWizardNavBarDesign** constant representing the design of the specified Web navigation bar set. Read/write.|
| [HorizontalAlignment](webnavigationbarset.horizontalalignment-property-publisher.md)|Sets or returns a  **PbWizardNavBarAlignment** constant that represents the horizontal alignment of the buttons in a Web navigation bar set. Read/write.|
| [HorizontalButtonCount](webnavigationbarset.horizontalbuttoncount-property-publisher.md)|Sets or returns a  **Long** representing the number of buttons in each row of buttons for a Web navigation bar set. Read/write. **Long**.|
| [IsHorizontal](webnavigationbarset.ishorizontal-property-publisher.md)| **True** if the specified **WebNavigationBarSet** has a horizontal orientation. Read-only **Boolean**.|
| [Links](webnavigationbarset.links-property-publisher.md)|Returns a  **WebNavigationBarHyperlinks** collection containing all of the hyperlinks in the specified Web navigation bar set. Read/write.|
| [Name](webnavigationbarset.name-property-publisher.md)|Indicates the name of the specified Web navigation bar set. Read/write.|
| [Parent](webnavigationbarset.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ShowSelected](webnavigationbarset.showselected-property-publisher.md)| **True** if the selected button is highlighted for the specified **WebNavigationBarSet** object. Read/write **Boolean**.|

