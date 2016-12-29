---
title: Page Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: cc2406e2-9f50-4cee-a602-4e5a3cf16f13
ms.locale: en-US
---


# Page Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](page.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Background](page.background-property-publisher.md)|Sets or returns a  **PageBackground** object representing the background of the specified page.|
| [Footer](page.footer-property-publisher.md)|Returns a  **HeaderFooter** object representing the footer of the specified **Page** object. Read-only.|
| [Header](page.header-property-publisher.md)|Returns a  **HeaderFooter** object representing the header of the specified **Page** object. Read-only.|
| [Height](page.height-property-publisher.md)|Returns a  **Long** that represent the height (in points) of a cell, range of cells, or page. Read-only.|
| [IgnoreMaster](page.ignoremaster-property-publisher.md)| **True** for Microsoft Publisher to ignore the master page formatting for the specified page. Read/write **Boolean**.|
| [IsLeading](page.isleading-property-publisher.md)| **True** if the specified **Page** object is a leading page of a two-page spread. Read-only **Boolean**.|
| [IsTrailing](page.istrailing-property-publisher.md)| **True** if the specified **Page** object is a trailing page of a two-page spread. Read-only **Boolean**.|
| [IsTwoPageMaster](page.istwopagemaster-property-publisher.md)| **True** if the specified **Page** object is a two-page master. Read/write **Boolean**.|
| [IsWizardPage](page.iswizardpage-property-publisher.md)|Returns  **True** if the specified page is a Microsoft Publisher wizard page. Read-only **Boolean**.|
| [LayoutGuides](page.layoutguides-property-publisher.md)|Returns a  **[LayoutGuides](layoutguides-object-publisher.md)** object consisting of the margin and grid layout guides for all pages including master pages in the publication.|
| [Master](page.master-property-publisher.md)|Sets or returns a  **[Page](page-object-publisher.md)** object that represents the master page properties for the specified page.|
| [Name](page.name-property-publisher.md)|Returns or sets a  **String** value indicating the name of the specified object. Read/write.|
| [PageID](page.pageid-property-publisher.md)|Returns a  **Long** indicating the unique identifier for a page in a publication. Read-only.|
| [PageIndex](page.pageindex-property-publisher.md)|Gets the index of the page in the  **Pages** collection of the active document. Read-only.|
| [PageNumber](page.pagenumber-property-publisher.md)|Returns a  **String** that represents the current page number. Read-only.|
| [PageType](page.pagetype-property-publisher.md)|Returns a  **PbPageType** constant that represents the page type. Read-only.|
| [Parent](page.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [ReaderSpread](page.readerspread-property-publisher.md)|Returns a  **[ReaderSpread](readerspread-object-publisher.md)** object that represents the reader spread of the specified page.|
| [RulerGuides](page.rulerguides-property-publisher.md)|Returns a  **[RulerGuides](rulerguides-object-publisher.md)** collection that represents gridlines used to align objects on a page.|
| [Shapes](page.shapes-property-publisher.md)|Returns a  **[Shapes](shapes-object-publisher.md)** collection that represents all the  **Shape** objects in the specified publication. This collection can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts.|
| [Tags](page.tags-property-publisher.md)|Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.|
| [WebPageOptions](page.webpageoptions-property-publisher.md)|Returns a  **[WebPageOptions](webpageoptions-object-publisher.md)** object, which represents the properties of a single Web page within a Web publication. Read-only.|
| [Width](page.width-property-publisher.md)|Returns a  **Long** that represent the width (in points) of a cell, range of cells, or page. Read-only.|
| [Wizard](page.wizard-property-publisher.md)|Returns a  **[Wizard](wizard-object-publisher.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.|
| [XOffsetWithinReaderSpread](page.xoffsetwithinreaderspread-property-publisher.md)|Returns a  **Single** that represents the distance (in points) from the left edge of the reader spread to the left edge of the page. Read-only.|
| [YOffsetWithinReaderSpread](page.yoffsetwithinreaderspread-property-publisher.md)|Returns a  **Single** that represents the distance (in points) from the top edge of the reader spread to the top edge of the page. Read-only.|

