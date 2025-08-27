---
title: Office Scripts preview APIs
description: Details about upcoming Office Scripts APIs.
ms.topic: whats-new
ms.date: 08/26/2025
---

# Office Scripts preview APIs

New Office Scripts APIs are first introduced in "preview", and then later released to general availability after sufficient testing occurs and user feedback is acquired.

> [!IMPORTANT]
> Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.

## API list

| Namespace | Class | Fields | Description |
|:---|:---|:---|:---|
| [OfficeScript](/javascript/api/office-scripts/officescript) | [EmailAttachment](/javascript/api/office-scripts/officescript/officescript.emailattachment) | | The attachment to send with the email. A value must be specified for at least one of the to, cc, or bcc parameters. If no recipient is specified, the following error is shown: "The message has no recipient. Please enter a value for at least one of the "to", "cc", or "bcc" parameters." |
| | | content | The contents of the file. |
| | | name | The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name. |
| | [MailProperties](/javascript/api/office-scripts/officescript/officescript.mailproperties) | | The properties of the email to be sent. |
| | | [attachments](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-attachments-member) | A file (such as a text file or Excel workbook) attached to a message. Optional. |
| | | [bcc](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-bcc-member) | The blind carbon copy (BCC) recipient or recipients of the email. Optional. |
| | | [cc](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-cc-member) | The carbon copy (CC) recipient or recipients of the email. Optional. |
| | | [content](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-content-member) | The content of the email. Optional. |
| | | [contentType](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-contentType-member) | The type of the content in the email. Possible values are text or HTML. Optional. |
| | | [importance](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-importance-member) | The importance of the email. The possible values are low, normal, and high. Default value is normal. Optional. |
| | | [subject](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-subject-member) | The subject of the email. Optional. |
| | | [to](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-to-member) | The direct recipient or recipients of the email. Optional. |

## See also

- [Office Scripts API reference](overview.md)
