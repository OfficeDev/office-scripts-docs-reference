---
title: Office Scripts API reference 
description: 'An overview of the Office Scripts Async JavaScript APIs'
ms.date: 06/17/2020
---

# Office Scripts Async API reference

The Office Scripts Async API supports older scripts made during the Office Scripts preview phase. Use this reference documentation to learn more about the classes, methods, and other types used by these older scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

> [!IMPORTANT]
> We strongly recommend creating new scripts with the standard Office Scripts APIs. If you're making or editing a new script, please switch to the [non-async version](?view=office-scripts-async) of the APIs.

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excel/excel.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excel.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).
- A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excel.range) objects.
- A [Range](/javascript/api/office-scripts/excel/excel.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excel/excel.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excel.table), [Charts](/javascript/api/office-scripts/excel/excel.chart), [Shapes](/javascript/api/office-scripts/excel/excel.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)) that are present in the individual sheet.
- [Workbooks](/javascript/api/office-scripts/excel/excel.workbook). contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel/excel.workbook).

For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [Using the Office Scripts Async APIs to support legacy scripts](/office/dev/scripts/develop/excel-async-model)
- [About Office Scripts](/office/dev/scripts/overview/excel)
- [Record, edit, and create Office Scripts in Excel on the web](/office/dev/scripts/tutorials/excel-tutorial)
- [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)
