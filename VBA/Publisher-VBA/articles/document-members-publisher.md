---
title: Document Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: d41cc0c0-b6f8-078b-6d29-0dab4962b06d
ms.locale: en-US
---


# Document Members (Publisher)
Represents a publication. 

## Events



|**Name**|**Description**|
|:-----|:-----|
| [BeforeClose](document.beforeclose-event-publisher.md)|Occurs immediately before any open document closes.|
| [Open](document.open-event-publisher.md)|Occurs when a publication is opening.|
| [Redo](document.redo-event-publisher.md)|Occurs when reversing the last action that was undone.|
| [ShapesAdded](document.shapesadded-event-publisher.md)|Occurs when one or more new shapes are added to a publication. This event occurs whether shapes are added manually or programmatically.|
| [ShapesRemoved](document.shapesremoved-event-publisher.md)|Occurs when a shape is deleted from a publication.|
| [Undo](document.undo-event-publisher.md)|Occurs when a user undoes the last action performed.|
| [WizardAfterChange](document.wizardafterchange-event-publisher.md)|Occurs after the user chooses an option in the wizard pane that changes any of the following settings in the publication: page layout (page size, fold type, orientation, label product), print setup (paper size, print tiling), adding or deleting objects, adding or deleting pages, or object or page formatting (size, position, fill, border, background, default text, text formatting).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [BeginCustomUndoAction](document.begincustomundoaction-method-publisher.md)|Specifies the starting point and label (textual description) of a group of actions that are wrapped to create a single undo action. The  **[EndCustomUndoAction](document.endcustomundoaction-method-publisher.md)** method is used to specify the endpoint of the actions used to create the single undo action. The wrapped group of actions can be undone with a single undo.|
| [ChangeDocument](document.changedocument-method-publisher.md)|Changes the current publication to one that uses the wizard, and optionally the design, that you specify.|
| [Close](document.close-method-publisher.md)|Closes the current publication and creates a blank publication in its place.|
| [ConvertPublicationType](document.convertpublicationtype-method-publisher.md)|Converts the specified publication to the specified publication type.|
| [EndCustomUndoAction](document.endcustomundoaction-method-publisher.md)|Specifies the endpoint of a group of actions that are wrapped to create a single undo action. The  ** [BeginCustomUndoAction Method](document.begincustomundoaction-method-publisher.md)** method is used to specify the starting point and label (textual description) of the actions used to create the single undo action. The wrapped group of actions can be undone with a single undo.|
| [ExportAsFixedFormat](document.exportasfixedformat-method-publisher.md)|Saves a Microsoft Publisher publication in PDF or XPS format. The conversion readies the document to be sent to commercial presses, to copy shops, for desktop printing, or for electronic distribution.|
| [FindShapeByWizardTag](document.findshapebywizardtag-method-publisher.md)|Returns a  **ShapeRange** object representing one or all of the shapes placed in a publication by a wizard and bearing the specified wizard tag.|
| [FindShapesByTag](document.findshapesbytag-method-publisher.md)|Returns a  **[ShapeRange](shaperange-object-publisher.md)** object that represents the shapes with the specified tag.|
| [PrintOutEx](document.printoutex-method-publisher.md)|Prints all or part of the specified publication.|
| [Redo](document.redo-method-publisher.md)|Redoes the last action or a specified number of actions. Corresponds to the list of items that appears when you click the arrow beside the  **Redo** button on the **Standard** toolbar. Calling this method reverses the ** [Undo Method](document.undo-method-publisher.md)** method.|
| [Save](document.save-method-publisher.md)|Saves the specified publication.|
| [SaveAs](document.saveas-method-publisher.md)|Saves the specified publication with a new name or format.|
| [SetBusinessInformation](document.setbusinessinformation-method-publisher.md)|Applies the specified business information set, which consists of a logo image and business contact information (such as the company name and address), to the current publication.|
| [Undo](document.undo-method-publisher.md)|Undoes the last action or a specified number of actions. Corresponds to the list of items that appears when you click the arrow beside the  **Undo** button on the **Standard** toolbar.|
| [UndoClear](document.undoclear-method-publisher.md)|Clears the list of actions that can be undone for the specified publication. Corresponds to the list of items that appears when you click the arrow beside the  **Undo** button on the **Standard** toolbar.|
| [UpdateOLEObjects](document.updateoleobjects-method-publisher.md)|Updates linked and embedded OLE objects.|
| [WebPagePreview](document.webpagepreview-method-publisher.md)|Generates a Web page preview of the specified publication in Internet Explorer.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [ActiveView](document.activeview-property-publisher.md)|Returns a  **[View](view-object-publisher.md)** object representing the view attributes for the specified document. Read-only.|
| [ActiveWindow](document.activewindow-property-publisher.md)|Returns a  **[Window](window-object-publisher.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.|
| [AdvancedPrintOptions](document.advancedprintoptions-property-publisher.md)|Returns an  **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object that represents the advanced print settings for a publication. Read-only.|
| [Application](document.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AvailableBuildingBlocks](document.availablebuildingblocks-property-publisher.md)|Returns the available members of the  **[BuildingBlocks](buildingblocks-object-publisher.md)** collection of the current document. Read-only.|
| [BorderArts](document.borderarts-property-publisher.md)|Returns a  **[BorderArts](borderarts-object-publisher.md)** collection that represents the BorderArt types available for use in the specified publication. Read-only.|
| [ColorScheme](document.colorscheme-property-publisher.md)|Returns or sets the  **[ColorScheme](colorscheme-object-publisher.md)** object that represents the scheme colors for the specified publication. Read/write.|
| [DefaultTabStop](document.defaulttabstop-property-publisher.md)|Returns or sets a  **Variant** corresponding to the default tab stop for all text in the active publication. Valid range is 1 to 1584 points (0.014" to 22"). Once set, numeric values are considered to be in points. **String** values may be in any unit supported by Microsoft Publisher. Point values are always returned. If values are outside the valid range, an error is returned. Read/write.|
| [DocumentDirection](document.documentdirection-property-publisher.md)|Returns or sets a  **PbDirectionType** constant that indicates whether text in the document is read from left to right or from right to left. Read/write.|
| [EnvelopeVisible](document.envelopevisible-property-publisher.md)|Returns or sets a  **Boolean** indicating whether the e-mail message header is visible in the publication window. Read/write.|
| [Find](document.find-property-publisher.md)||
| [FullName](document.fullname-property-publisher.md)|Returns a  **String** representing the full file name of the saved active publication, including its path and file name. Read-only.|
| [IsDataSourceConnected](document.isdatasourceconnected-property-publisher.md)| **True** if the specified publication is connected to a data source. Read-only.|
| [IsWizard](document.iswizard-property-publisher.md)|Returns  **True** if the specified publication is a publication generated by a Microsoft Publisher wizard. Read-only **Boolean**.|
| [LayoutGuides](document.layoutguides-property-publisher.md)|Returns a  **[LayoutGuides](layoutguides-object-publisher.md)** object consisting of the margin and grid layout guides for all pages including master pages in the publication.|
| [MailEnvelope](document.mailenvelope-property-publisher.md)|Returns an  **MsoEnvelope** object that represents an e-mail header for a publication.|
| [MailMerge](document.mailmerge-property-publisher.md)|Returns a  **[MailMerge](mailmerge-object-publisher.md)** object that represents the mail merge functionality for the specified publication.|
| [MasterPages](document.masterpages-property-publisher.md)|Returns the  **[MasterPages](masterpages-object-publisher.md)** collection for the specified publication.|
| [Name](document.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Pages](document.pages-property-publisher.md)|Returns a  **[Pages](pages-object-publisher.md)** collection representing all the pages in the specified publication.|
| [PageSetup](document.pagesetup-property-publisher.md)|Returns a  **[PageSetup](pagesetup-object-publisher.md)** object representing a publication's page size, page layout and paper settings. Read-only.|
| [Parent](document.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Path](document.path-property-publisher.md)|Returns a  **String** indicating the full path to the file of the saved active publication, not including the last separator or file name.|
| [PrintPageBackgrounds](document.printpagebackgrounds-property-publisher.md)|Returns or sets  **True** to include page backgrounds when printing pages from the specified publication. Default is **True**. Read/write  **Boolean**.|
| [PrintStyle](document.printstyle-property-publisher.md)|Read-only|
| [PublicationType](document.publicationtype-property-publisher.md)|Returns a  **PbPublicationType** constant that represents the type of the specified publication. Read-only.|
| [ReadOnly](document.readonly-property-publisher.md)|Returns  **True** if the publication is read-only; returns **False** if it is read/write. Read-only **Boolean**.|
| [RedoActionsAvailable](document.redoactionsavailable-property-publisher.md)|Returns the number of actions available on the redo stack. Read-only  **Long**.|
| [RemovePersonalInformation](document.removepersonalinformation-property-publisher.md)|Returns or sets a  **Boolean** that represents whether to save personal information when the file is saved. Read/write.|
| [Saved](document.saved-property-publisher.md)|Returns  **True** if no changes have been made to a publication since it was last saved. Read-only **Boolean**.|
| [SaveFormat](document.saveformat-property-publisher.md)|Indicates the file format of the specified document. Read-only.|
| [ScratchArea](document.scratcharea-property-publisher.md)|Returns a  **[ScratchArea](scratcharea-object-publisher.md)** object for an a given document.|
| [Sections](document.sections-property-publisher.md)|Returns a  **Sections** object representing a collection of **Section** objects in the specified document. Read-only **Sections**.|
| [Selection](document.selection-property-publisher.md)|Returns a  **[Selection](selection-object-publisher.md)** object that represents a selected range or the cursor.|
| [Stories](document.stories-property-publisher.md)|Returns a  **[Stories](stories-object-publisher.md)** collection containing all stories in the publication.|
| [SurplusShapes](document.surplusshapes-property-publisher.md)|Returns a  **ShapeRange** object that represents the collection of surplus shapes that Microsoft Publisher places under **Extra Content**in the  **Format Publication** task pane after the document template (wizard) is changed by using the ** [Document.ChangeDocument](document.changedocument-method-publisher.md)** method or by using the **Change Template** command in the user interface. Read-only.|
| [Tags](document.tags-property-publisher.md)|Returns a  **[Tags](tags-object-publisher.md)** collection representing tags or custom properties applied to a shape, shape range, page, or publication.|
| [TextStyles](document.textstyles-property-publisher.md)|Returns a  **[TextStyles](textstyles-object-publisher.md)** collection that contains a publication's text styles.|
| [UndoActionsAvailable](document.undoactionsavailable-property-publisher.md)|Returns the number of actions available on the undo stack. Read-only  **Long**.|
| [ViewBoundaries](document.viewboundaries-property-publisher.md)|Returns  **True** if object boundaries are visible in the specified publication. Read/write **Boolean**.|
| [ViewGuides](document.viewguides-property-publisher.md)|Returns  **True** if guides are visible in the specified publication. Read/write **Boolean**.|
| [ViewHorizontalBaseLineGuides](document.viewhorizontalbaselineguides-property-publisher.md)|Sets or returns a  **Boolean** that represents whether or not the horizontal baseline guides are visible in the specified **Document** object. **True** if they are visible. **False** if they are not visible. Read/write.|
| [ViewTwoPageSpread](document.viewtwopagespread-property-publisher.md)|Returns  **True** if the specified publication should be viewed as a two-page spread. Read/write **Boolean**.|
| [ViewVerticalBaseLineGuides](document.viewverticalbaselineguides-property-publisher.md)|Sets or returns a  **Boolean** that represents whether or not the vertical baseline guides are visible in the specified **Document** object. **True** if they are visible. **False** if they are not visible. Read/write.|
| [WebNavigationBarSets](document.webnavigationbarsets-property-publisher.md)|Returns a  **WebNavigationBarSets** object representing a collection of all **WebNavigationBarSet** objects in the specified document. Read-only.|
| [Wizard](document.wizard-property-publisher.md)|Returns a  **[Wizard](wizard-object-publisher.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.|

