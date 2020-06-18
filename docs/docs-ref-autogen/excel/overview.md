---
title: Office Scripts API reference 
description: 'An overview of the Office Scripts JavaScript APIs.'
ms.date: 06/17/2020
---

# Office Scripts API reference

The Office Scripts API lets you automate common tasks in Excel on the web. Use this reference documentation to learn more about the classes, methods, and other types available for your scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

> [!NOTE]
> If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet).
- A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.
- A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.
- A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.

For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [About Office Scripts](/office/dev/scripts/overview/excel)
- [Record, edit, and create Office Scripts in Excel on the web](/office/dev/scripts/tutorials/excel-tutorial)
- [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)
