---
title: Application Members (Publisher)
ms.prod: PUBLISHER
ms.assetid: aa4d515b-f779-b8b5-968a-8e5f7466fb56
ms.locale: en-US
---


# Application Members (Publisher)
Represents the Microsoft Publisher application. The  **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.

## Events



|**Name**|**Description**|
|:-----|:-----|
| [AfterPrint](application.afterprint-event-publisher.md)|Fires after all variables and fields print.|
| [BeforePrint](application.beforeprint-event-publisher.md)|Occurs before the publication is printed or previewed. .|
| [DocumentBeforeClose](application.documentbeforeclose-event-publisher.md)|Occurs immediately before any open document closes.|
| [DocumentOpen](application.documentopen-event-publisher.md)|Occurs when opening a document.|
| [HideCatalogUI](application.hidecatalogui-event-publisher.md)|Occurs when the catalog of publication wizards is hidden in the Microsoft Publisher user interface.|
| [MailMergeAfterMerge](application.mailmergeaftermerge-event-publisher.md)|Occurs after all records in a mail merge have merged successfully.|
| [MailMergeAfterRecordMerge](application.mailmergeafterrecordmerge-event-publisher.md)|Occurs after each record in the data source successfully merges in a mail merge.|
| [MailMergeBeforeMerge](application.mailmergebeforemerge-event-publisher.md)|Occurs when a merge is executed before any records in a mail merge have merged.|
| [MailMergeBeforeRecordMerge](application.mailmergebeforerecordmerge-event-publisher.md)|Occurs as a merge is executed for the individual records in a merge.|
| [MailMergeDataSourceLoad](application.mailmergedatasourceload-event-publisher.md)|Occurs when the data source is loaded for a mail merge.|
| [MailMergeDataSourceValidate](application.mailmergedatasourcevalidate-event-publisher.md)|Occurs when a user performs address verification by clicking  **Validate** in the **Mail Merge Recipients** dialog box.|
| [MailMergeGenerateBarcode](application.mailmergegeneratebarcode-event-publisher.md)|Occurs when Microsoft Publisher requires data to generate barcodes in a mail-merge publication, in particular when the mail-merge recipient list changes.|
| [MailMergeInsertBarcode](application.mailmergeinsertbarcode-event-publisher.md)|Occurs when the user issues the command to insert postal barcodes into a mail-merge publication, either in the Microsoft Publisher user interface (UI), or programmatically.|
| [MailMergeRecipientListClose](application.mailmergerecipientlistclose-event-publisher.md)|Fires when the user closes the  **Mail Merge Recipients** dialog box. (From the **Mail Merge** or **E-mail Merge** task pane, click **Edit Recipient List**). Also fires when the user closes the  **Catalog Merge Product List** dialog box, which opens when the user clicks **Edit Product List** in the **Catalog Merge** task pane.|
| [MailMergeWizardFollowUpCustom](application.mailmergewizardfollowupcustom-event-publisher.md)|Fires when the string that appears as the fourth item under  **Prepare to follow-up on this mailing** on the third **Mail Merge** task pane in the Microsoft Publisher user interface is clicked.|
| [MailMergeWizardStateChange](application.mailmergewizardstatechange-event-publisher.md)|Occurs when a user changes from a specified step to a specified step in the Mail Merge Wizard.|
| [NewDocument](application.newdocument-event-publisher.md)|Occurs when a new publication is created.|
| [Quit](application.quit-event-publisher.md)|Occurs when the user exits Microsoft Publisher.|
| [ShowCatalogUI](application.showcatalogui-event-publisher.md)|Fires when the catalog of publication wizards is displayed in the Microsoft Publisher user interface.|
| [WindowActivate](application.windowactivate-event-publisher.md)|Occurs when the application window is activated.|
| [WindowDeactivate](application.windowdeactivate-event-publisher.md)|Occurs when the application window is deactivated.|
| [WindowPageChange](application.windowpagechange-event-publisher.md)|Occurs when switching the view from one page to another page in a publication.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
| [CentimetersToPoints](application.centimeterstopoints-method-publisher.md)|Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single**.|
| [ChangeFileOpenDirectory](application.changefileopendirectory-method-publisher.md)|Sets the folder in which Microsoft Publisher searches for documents. The specified folder's contents are listed the next time the  **Open Publication** dialog box ( **File** menu) is displayed.|
| [EmusToPoints](application.emustopoints-method-publisher.md)|Converts a measurement from emus to points (12700 emus = 1 point). Returns the converted measurement as a  **Single**.|
| [Help](application.help-method-publisher.md)|Displays online Help information.|
| [InchesToPoints](application.inchestopoints-method-publisher.md)|Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single**.|
| [IsValidObject](application.isvalidobject-method-publisher.md)|Determines whether the specified object variable references a valid object and returns a  **Boolean** value: **True** if the specified variable that references an object is valid, **False** if the object referenced by the variable has been deleted.|
| [LinesToPoints](application.linestopoints-method-publisher.md)|Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single**.|
| [MillimetersToPoints](application.millimeterstopoints-method-publisher.md)|Converts a measurement from millimeters to points (1 mm = 2.835 points). Returns the converted measurement as a  **Single**.|
| [NewDocument](application.newdocument-method-publisher.md)|Returns a  **Document** object that represents a new publication.|
| [Open](application.open-method-publisher.md)|Returns a  **[Document](document-object-publisher.md)** object that represents the newly opened publication.|
| [PicasToPoints](application.picastopoints-method-publisher.md)|Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single**.|
| [PixelsToPoints](application.pixelstopoints-method-publisher.md)|Converts a measurement from pixels to points (1 pixel = 0.75 points). Returns the converted measurement as a  **Single**.|
| [PointsToCentimeters](application.pointstocentimeters-method-publisher.md)|Converts a measurement from points to centimeters (1 cm = 28.35 points). Returns the converted measurement as a  **Single**.|
| [PointsToEmus](application.pointstoemus-method-publisher.md)|Converts a measurement from points to emus (12700 emus = 1 point). Returns the converted measurement as a  **Single**.|
| [PointsToInches](application.pointstoinches-method-publisher.md)|Converts a measurement from points to inches (1 in = 72 points). Returns the converted measurement as a  **Single**.|
| [PointsToLines](application.pointstolines-method-publisher.md)|Converts a measurement from points to lines (1 line = 12 points). Returns the converted measurement as a  **Single**.|
| [PointsToMillimeters](application.pointstomillimeters-method-publisher.md)|Converts a measurement from points to millimeters (1 mm = 2.835 points). Returns the converted measurement as a  **Single**.|
| [PointsToPicas](application.pointstopicas-method-publisher.md)|Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a  **Single**.|
| [PointsToPixels](application.pointstopixels-method-publisher.md)|Converts a measurement from points to pixels (1 pixel = 0.75 points). Returns the converted measurement as a  **Single**.|
| [PointsToTwips](application.pointstotwips-method-publisher.md)|Converts a measurement from points to twips (20 twips = 1 point). Returns the converted measurement as a  **Single**.|
| [Quit](application.quit-method-publisher.md)|Quits Microsoft Publisher. This is equivalent to clicking  **Exit** on the **File** menu.|
| [ShowWizardCatalog](application.showwizardcatalog-method-publisher.md)|Displays the  **Publication Types** catalog for the wizard of the specified type.|
| [TwipsToPoints](application.twipstopoints-method-publisher.md)|Converts a measurement from twips to points (20 twips = 1 point). Returns the converted measurement as a  **Single**.|

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

