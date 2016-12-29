---
title: MailMerge Object (Publisher)
keywords: vbapb10.chm6291455
f1_keywords: vbapb10.chm6291455
ms.prod: PUBLISHER
api_name: Publisher.MailMerge
ms.assetid: 028e1e42-c61c-9b2b-4aec-d6a184504ec1
ms.locale: en-US
---


# MailMerge Object (Publisher)

Represents the mail merge and catalog merge functionality in Microsoft Publisher.


## Example

Use the  **[MailMerge](http://msdn.microsoft.com/library/document.mailmerge-property-publisher%28Office.15%29.aspx)** property to return the **MailMerge** object. The **MailMerge** object is always available regardless of whether the mail merge or catalog merge operation has begun. The following example merges and prints the main publication with the first three records in the attached data source.


```
Sub SelectiveMerge() 
 Dim mrgMain As MailMerge 
 Set mrgMain = ActiveDocument.MailMerge 
 With mrgMain.DataSource 
 .FirstRecord = 1 
 .LastRecord = 3 
 End With 
 mrgMain.Execute True 
End Sub
```


## Methods



|**Name**|
|:-----|
|[CreateShortcut](http://msdn.microsoft.com/library/mailmerge.createshortcut-method-publisher%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/mailmerge.execute-method-publisher%28Office.15%29.aspx)|
|[ExportRecipientList](http://msdn.microsoft.com/library/mailmerge.exportrecipientlist-method-publisher%28Office.15%29.aspx)|
|[OpenDataSource](http://msdn.microsoft.com/library/mailmerge.opendatasource-method-publisher%28Office.15%29.aspx)|
|[ShowWizardEx](http://msdn.microsoft.com/library/mailmerge.showwizardex-method-publisher%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/mailmerge.application-property-publisher%28Office.15%29.aspx)|
|[DataSource](http://msdn.microsoft.com/library/mailmerge.datasource-property-publisher%28Office.15%29.aspx)|
|[DocumentUpdating](http://msdn.microsoft.com/library/mailmerge.documentupdating-property-publisher%28Office.15%29.aspx)|
|[EmailMergeEnvelope](http://msdn.microsoft.com/library/mailmerge.emailmergeenvelope-property-publisher%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/mailmerge.parent-property-publisher%28Office.15%29.aspx)|
|[SuppressBlankLines](http://msdn.microsoft.com/library/mailmerge.suppressblanklines-property-publisher%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/mailmerge.type-property-publisher%28Office.15%29.aspx)|
|[ViewMailMergeFieldCodes](http://msdn.microsoft.com/library/mailmerge.viewmailmergefieldcodes-property-publisher%28Office.15%29.aspx)|
|[WizardState](http://msdn.microsoft.com/library/mailmerge.wizardstate-property-publisher%28Office.15%29.aspx)|

