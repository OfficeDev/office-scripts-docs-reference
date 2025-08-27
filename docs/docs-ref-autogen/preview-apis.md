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

The following table lists the Office Scripts APIs currently in preview.

| Namespace | Class | Fields | Description |
|:---|:---|:---|:---|
| [OfficeScript](/javascript/api/office-scripts/officescript) | | [convertToPdf()](/javascript/api/office-scripts/officescript#officescript-officescript-converttopdf-function(1)) | Return the text encoding of the document as a PDF. If the document is empty, then the following error is shown: "We didn't find anything to print". Some actions made prior to using this API may not be captured in the PDF in Excel on the web. |
| | | [downloadFile({ name, content, })](javascript/api/office-scripts/officescript#officescript-officescript-downloadfile-function(1)) | Downloads a specified file to the default download location specified by the local machine. |
| | | [Metadata.getScriptName()](/javascript/api/office-scripts/officescript#officescript-officescript-metadata-getscriptname-function(1)) | Get the name of the currently running script. |
| | | [saveCopyAs(filename)](/javascript/api/office-scripts/officescript#officescript-officescript-savecopyas-function(1)) | Saves a copy of the current workbook in OneDrive, in the same directory as the original file, with the specified file name. The API has a timeout limit of 30 seconds. This limit is rarely exceeded. |
| | | [sendMail(mailProperties)](/javascript/api/office-scripts/officescript#officescript-officescript-sendmail-function(1)) | Send an email with an Office Script. Use MailProperties to specify the content and recipients of the email. If the request body includes content, this method returns 400 Bad request. |
| [OfficeScript](/javascript/api/office-scripts/officescript) | [EmailAttachment](/javascript/api/office-scripts/officescript/officescript.emailattachment) | [content](/javascript/api/office-scripts/officescript/officescript.emailattachment#officescript-officescript-emailattachment-content-member) | The contents of the file. |
| | | [name](/javascript/api/office-scripts/officescript/officescript.emailattachment#officescript-officescript-emailattachment-name-member) | The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name. |
| [OfficeScript](/javascript/api/office-scripts/officescript) | [MailProperties](/javascript/api/office-scripts/officescript/officescript.mailproperties) | [attachments](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-attachments-member) | A file (such as a text file or Excel workbook) attached to a message. Optional. |
| | | [bcc](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-bcc-member) | The blind carbon copy (BCC) recipient or recipients of the email. Optional. |
| | | [cc](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-cc-member) | The carbon copy (CC) recipient or recipients of the email. Optional. |
| | | [content](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-content-member) | The content of the email. Optional. |
| | | [contentType](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-contentType-member) | The type of the content in the email. Possible values are text or HTML. Optional. |
| | | [importance](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-importance-member) | The importance of the email. The possible values are `low`, `normal`, and `high`. Default value is `normal`. Optional. |
| | | [subject](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-subject-member) | The subject of the email. Optional. |
| | | [to](/javascript/api/office-scripts/officescript/officescript.mailproperties#officescript-officescript-mailproperties-to-member) | The direct recipient or recipients of the email. Optional. |

## See also

- [Office Scripts API reference](overview.md)
