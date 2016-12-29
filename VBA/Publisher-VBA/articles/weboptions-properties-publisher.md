---
title: WebOptions Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: 44af230f-34ed-453a-b390-e12f7ae860b8
ms.locale: en-US
---


# WebOptions Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [AlwaysSaveInDefaultEncoding](weboptions.alwayssaveindefaultencoding-property-publisher.md)|Returns or sets a  **Boolean** value that specifies whether Web pages within a Web publication should always be saved using default encoding. If **True**, Web pages within a publication will always be saved using the default encoding of the client computer. If  **False**, Web pages will not be saved using default encoding. The default value is  **False**. Read/write.|
| [Application](weboptions.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [EmailAsImg](weboptions.emailasimg-property-publisher.md)| **True** to send the entire publication page as a single JPEG image. Read/write **Boolean**.|
| [EnableIncrementalUpload](weboptions.enableincrementalupload-property-publisher.md)|Returns or sets a  **Boolean** value that specifies whether changes made to a Web publication can be uploaded to a Web server independent of the entire publication. If **True**, only changes made to a publication will be uploaded to the Web server when published. If  **False**, the entire publication will be uploaded to the Web server. The default value is  **True**. Read/write.|
| [Encoding](weboptions.encoding-property-publisher.md)|Returns an  **MsoEncoding** constant that specifies the encoding of the Web publication. Read/write.|
| [OrganizeInFolder](weboptions.organizeinfolder-property-publisher.md)|Returns or sets a  **Boolean** value that specifies whether a Web publication will be saved in a flat structure or hierarchical structure. If **False**, all files in the Web publication will be saved in a flat structure within the root folder. If  **True**, the files will be saved in a hierarchical structure within the root folder. The default value is  **True**. Read/write.|
| [Parent](weboptions.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [RelyOnVML](weboptions.relyonvml-property-publisher.md)|Returns or sets a  **Boolean** value that specifies whether image files are generated from drawing objects when a Web publication is saved. If **True**, image files are not generated. If  **False**, images are generated. The default value is  **False**. Read/write.|
| [ShowOnlyWebFonts](weboptions.showonlywebfonts-property-publisher.md)|Returns or sets a **Boolean** value that specifies whether only Web-safe fonts and font schemes should be used when the Web site is viewed in a browser. If **True**, only Web-safe fonts and font schemes are used. If  **False**, display is not limited to Web-safe fonts and font schemes. The default value is  **False**. Read/write.|
|Name|Description|

