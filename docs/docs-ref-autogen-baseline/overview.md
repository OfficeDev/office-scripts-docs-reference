---
title: Office Scripts API reference 
description: An overview of the Office Scripts JavaScript APIs.
ms.topic: overview
ms.date: 12/04/2025
---

# Office Scripts API reference

The Office Scripts API lets you automate common tasks in Excel. Use this reference documentation to learn more about the classes, methods, and other types available for your scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

> [!NOTE]
> If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview&preserve-view=true).

## Namespaces

Office Scripts APIs use two namespaces: [OfficeScript](/javascript/api/office-scripts/officescript) for APIs that are not connected to an Excel workbook, and [ExcelScript](/javascript/api/office-scripts/excelscript) for APIs that work with Excel workbooks.

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excelscript/excelscript.range) objects.
- A [Range](/javascript/api/office-scripts/excelscript/excelscript.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excelscript/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excelscript/excelscript.table), [Charts](/javascript/api/office-scripts/excelscript/excelscript.chart), [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.
- A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.

For more information about the Office Scripts object model, visit [Fundamentals for Office Scripts in Excel](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [About Office Scripts](/office/dev/scripts/overview/excel)
- [Record, edit, and create Office Scripts in Excel](/office/dev/scripts/tutorials/excel-tutorial)
- [Fundamentals for Office Scripts in Excel](/office/dev/scripts/develop/scripting-fundamentals)