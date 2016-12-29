---
title: Application Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: f66bb6c0-e5e8-4ac7-bfdc-94344cba132e
ms.locale: en-US
---


# Application Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [ActiveDocument](application.activedocument-property-publisher.md)|Returns a  **[Document](document-object-publisher.md)** object that represents the active publication. If there are no documents open, an error occurs.|
| [ActiveWindow](application.activewindow-property-publisher.md)|Returns a  **[Window](window-object-publisher.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.|
| [Application](application.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Assistance](application.assistance-property-publisher.md)|Gets a reference to the Microsoft Office (MSO) **IAssistance** object, which provides a means for developers to create a customized Help experience for users within Microsoft Office. Read-only.|
| [AutomationSecurity](application.automationsecurity-property-publisher.md)|Specifies the security mode that Microsoft Publisher uses when programmatically opening files. Read/write.|
| [Build](application.build-property-publisher.md)|Returns a  **Long** that represents the Microsoft Publisher build number. Read-only.|
| [CaptionStyles](application.captionstyles-property-publisher.md)|Returns a  **[CaptionStyles](documents-object-publisher.md)** collection that represents all the available caption styles. Read-only.|
| [ColorSchemes](application.colorschemes-property-publisher.md)|Returns a  **[ColorSchemes](colorschemes-object-publisher.md)** collection that represents the color schemes available.|
| [COMAddIns](application.comaddins-property-publisher.md)|Returns a  **COMAddIns** collection that represents a reference to the Component Object Model (COM) add-ins currently loaded in Microsoft Publisher.|
| [CommandBars](application.commandbars-property-publisher.md)|Sets or returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Publisher.|
| [Documents](application.documents-property-publisher.md)|Returns a  **[Documents](documents-object-publisher.md)** collection that represents all open publications. Read-only.|
| [FileDialog](application.filedialog-property-publisher.md)|Returns a  **FileDialog** object that represents a single instance of a file dialog box.|
| [InsertBarcodeVisible](application.insertbarcodevisible-property-publisher.md)|Determines whether  **Add a postal bar code** is available under **More Items** on the **Mail Merge** and **Catalog Merge** task panes in the Microsoft Publisher user interface (UI); and whether **Add postal bar codes** is available under **Prepare for Mailing** on the **Publisher Tasks** task pane in the Publisher UI. Read/write.|
| [InstalledPrinters](application.installedprinters-property-publisher.md)|Returns an  **[InstalledPrinters](installedprinters-object-publisher.md)** collection that represents the names of all the printers installed on the computer and to which Microsoft Publisher can print the publication. Read-only.|
| [Language](application.language-property-publisher.md)|Returns a  **Long** that represents the language selected for the Microsoft Publisher user interface. Read-only.|
| [Name](application.name-property-publisher.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [OfficeDataSourceObject](application.officedatasourceobject-property-publisher.md)|Returns an  **OfficeDataSourceObject** object representing the data source in a mail merge or catalog merge operation. Read-only.|
| [Options](application.options-property-publisher.md)|Returns an  **[Options](options-object-publisher.md)** object that represents application settings you can set in Microsoft Publisher.|
| [Parent](application.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Path](application.path-property-publisher.md)|Returns a  **String** indicating the full path to the file of the saved active publication, not including the last separator or file name.|
| [PathSeparator](application.pathseparator-property-publisher.md)|Returns a  **String** that represents the character used to separate folder names. Read-only.|
| [PrintPreview](application.printpreview-property-publisher.md)| **True** to display in Print Preview the publication in the current view. Read/write **Boolean**.|
| [ProductCode](application.productcode-property-publisher.md)|Returns a  **String** indicating the Microsoft Publisher globally unique identifier (GUID). Read-only.|
| [ScreenUpdating](application.screenupdating-property-publisher.md)|Returns or sets a  **Boolean** indicating whether Microsoft Publisher refreshes the screen display during run time; **True** to refresh the screen display. Read/write.|
| [Selection](application.selection-property-publisher.md)|Returns a  **[Selection](selection-object-publisher.md)** object that represents a selected range or the cursor.|
| [ShowFollowUpCustom](application.showfollowupcustom-property-publisher.md)|Gets or sets the string, if any, that appears as the fourth item under  **Prepare to follow-up on this mailing** on the third **Mail Merge** task pane in the Microsoft Publisher user interface. Read/write.|
| [SnapToGuides](application.snaptoguides-property-publisher.md)| **True** for Microsoft Publisher to use the guides to align objects on a page in a publication. Read/write **Boolean**.|
| [SnapToObjects](application.snaptoobjects-property-publisher.md)| **True** for Microsoft Publisher to use objects on a page to line up other objects. Read/write **Boolean**.|
| [TemplateFolderPath](application.templatefolderpath-property-publisher.md)|Returns a  **String** that represents the location where Microsoft Publisher templates are stored. Read-only.|
| [ValidateAddressVisible](application.validateaddressvisible-property-publisher.md)|Determines whether  **Validate addresses** is available under **Refine recipients** in the **Mail Merge Recipients** dialog box in the Microsoft Publisher user interface (UI); and whether **Validate addresses** is available under **Prepare for Mailing** on the **Publisher Tasks** task pane in the Publisher UI. Read/write.|
| [Version](application.version-property-publisher.md)|Returns a  **String** indicating the version number of the currently-installed copy of Microsoft Publisher. Read-only.|
| [WebOptions](application.weboptions-property-publisher.md)|Returns a  **[WebOptions](weboptions-object-publisher.md)** object, which represents the properties of Web publications. Read-only.|
| [WizardCatalogVisible](application.wizardcatalogvisible-property-publisher.md)|Returns or sets a  **Boolean** indicating whether the Wizard Catalog is visible. Read/write.|

