---
title: Office Scripts preview APIs
description: Details about upcoming Office Scripts APIs.
ms.topic: whats-new
ms.date: 08/26/2025
---

# Office Scripts preview APIs

New Office Scripts APis are first introduced in "preview", and then later released to general availability after sufficient testing occurs and user feedback is acquired.

> [!IMPORTANT]
> Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.

## API list

| Namespace | Class | Fields | Description |
|:---|:---|:---|
| [OfficeScript](/officescript) | [EmailAttachment](/officescript/officescript.emailattachment) | | The attachment to send with the email. A value must be specified for at least one of the to, cc, or bcc parameters. If no recipient is specified, the following error is shown: "The message has no recipient. Please enter a value for at least one of the "to", "cc", or "bcc" parameters." |
| [OfficeScript](/officescript) | [EmailAttachment](/officescript/officescript.emailattachment) | content | The contents of the file. |
| [OfficeScript](/officescript) | [EmailAttachment](/officescript/officescript.emailattachment) | name | The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name. |

## See also

- [Office Scripts API reference](overview.md)
- 