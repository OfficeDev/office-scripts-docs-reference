---
title: Office Scripts API reference 
description: 'An overview of the Office Scripts JavaScript APIs'
ms.date: 06/16/2020
---

# Office Scripts API reference

The Office Scripts API lets you automate common tasks in Excel on the web. Use this reference documentation to learn more about the classes, methods, and other types available for your scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excel/excel.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excel.worksheet).
- A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excel.range) objects.
- A [Range](/javascript/api/office-scripts/excel/excel.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excel/excel.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excel.table), [Charts](/javascript/api/office-scripts/excel/excel.chart), [Shapes](/javascript/api/office-scripts/excel/excel.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) contains arrays filled with those objects that are present in the individual sheet.
- A [Workbook](/javascript/api/office-scripts/excel/excel.workbook) contains arrays of some of those data objects for the entire Workbook.

For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [About Office Scripts](/office/dev/scripts/overview/excel)
- [Record, edit, and create Office Scripts in Excel on the web](/office/dev/scripts/tutorials/excel-tutorial)
- [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)
