---
title: Office Scripts API reference 
description: 'An overview of the Office Scripts Async JavaScript APIs.'
ms.date: 06/29/2020
---

# Office Scripts Async API reference

The Office Scripts Async API supports older scripts made during the Office Scripts preview phase. Use this reference documentation to learn more about the classes, methods, and other types used by these older scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

> [!IMPORTANT]
> We strongly recommend creating new scripts with the standard Office Scripts APIs. If you're making or editing a new script, please switch to the [non-async version](?view=office-scripts) of the APIs.

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excelscript/excelscript.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excelscript/excelscript.worksheetcollection).
- A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excelscript/excelscript.range) objects.
- A [Range](/javascript/api/office-scripts/excelscript/excelscript.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excelscript/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excelscript/excelscript.table), [Charts](/javascript/api/office-scripts/excelscript/excelscript.chart), [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excelscript/excelscript.chartcollection)) that are present in the individual sheet.
- [Workbooks](/javascript/api/office-scripts/excelscript/excelscript.workbook) contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excelscript/excelscript.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook).

For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [Using the Office Scripts Async APIs to support legacy scripts](/office/dev/scripts/develop/excel-async-model)
- [About Office Scripts](/office/dev/scripts/overview/excel)
- [Record, edit, and create Office Scripts in Excel on the web](/office/dev/scripts/tutorials/excel-tutorial)
- [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)
