---
title: CalloutFormat Properties (Publisher)
ms.prod: PUBLISHER
ms.assetid: e4510734-4b1c-4363-8f89-07d66c352b9f
ms.locale: en-US
---


# CalloutFormat Properties (Publisher)

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Accent](calloutformat.accent-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether a vertical accent bar separates the callout text from the callout line. Read/write.|
| [Angle](calloutformat.angle-property-publisher.md)|Returns or sets an  **MsoCalloutAngleType** constant that represents the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.|
| [Application](calloutformat.application-property-publisher.md)|Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [AutoAttach](calloutformat.autoattach-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points) is to the left or right of the callout text box. Read/write.|
| [AutoLength](calloutformat.autolength-property-publisher.md)|Returns an  **MsoTriState**constant indicating whether the first segment of the callout line is scaled when the callout is moved. Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour**). Read-only.|
| [Border](calloutformat.border-property-publisher.md)|Returns or sets an  **MsoTriState**constant indicating whether the text in the specified callout is surrounded by a border. Read/write.|
| [Drop](calloutformat.drop-property-publisher.md)|For callouts with an explicitly set drop value, this property returns the vertical distance from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (where the callout points). In this case, the drop distance is measured from the bottom of the text box. Read-only **Variant**.|
| [DropType](calloutformat.droptype-property-publisher.md)|Returns an  **MsoCalloutDropType** constant indicating where the callout line attaches to the callout text box. Read-only.|
| [Gap](calloutformat.gap-property-publisher.md)|Returns or sets a  **Variant** indicating the horizontal distance between the end of the callout line and the text bounding box. Read/write.|
| [Length](calloutformat.length-property-publisher.md)|Returns a  **Variant** indicating the length (in points) of the first segment of the callout line (the segment attached to the text callout box) if the **[AutoLength](calloutformat.autolength-property-publisher.md)** property of the specified callout is set to **False**. Otherwise, an error occurs. Read-only.|
| [Parent](calloutformat.parent-property-publisher.md)|Returns an object that represents the parent object of the specified object. For example, for a  **[TextFrame](textframe-object-publisher.md)** object, returns a **[Shape](shape-object-publisher.md)** object representing the parent shape of the text frame. Read-only.|
| [Type](calloutformat.type-property-publisher.md)|Returns or sets an  **MsoCalloutType** constant that represents the callout type. Read/write.|
|Name|Description|

