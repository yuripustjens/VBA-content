---
title: MailMerge Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 63efb9c7-65ba-4b6f-8322-461eec6ac1f5
ms.locale: en-US
---


# MailMerge Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](mailmerge.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [DataSource](mailmerge.datasource-property-publisher.md)|Returns a  **[MailMergeDataSource](mailmergedatasource-object-publisher.md)** object that refers to the data source attached to a mail merge or catalog merge main publication.|
| [DocumentUpdating](mailmerge.documentupdating-property-publisher.md)|Returns or sets a  **Boolean** indicating whether the screen is updated while mail merge code is running. Default is **True** (the screen is updated). Read/write.|
| [EmailMergeEnvelope](mailmerge.emailmergeenvelope-property-publisher.md)|Returns the  **EmailMergeEnvelope** object associated with the parent **MailMerge** object. Read-only.|
| [Parent](mailmerge.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [SuppressBlankLines](mailmerge.suppressblanklines-property-publisher.md)| **True** to suppress blank lines when mail merge fields in a mail merge main document are empty. Read/write **Boolean**.|
| [Type](mailmerge.type-property-publisher.md)|Gets or sets the type of mail merge represented by the parent  **MailMerge** object. Read/write.|
| [ViewMailMergeFieldCodes](mailmerge.viewmailmergefieldcodes-property-publisher.md)| **True** if merge field names are displayed in a mail merge publication; **False** if information from the current record is displayed. Read/write **Boolean**. .|
| [WizardState](mailmerge.wizardstate-property-publisher.md)|Returns or sets a  **Long** indicating the current Mail Merge wizard step for a publication. The **WizardState** property returns a number that equates to the current Mail Merge wizard step; a zero (0) means the Mail Merge wizard is closed. Read/write.|
|Name|Description|

