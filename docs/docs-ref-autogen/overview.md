---
title: Office Scripts API reference 
description: 'An overview of the Office Scripts JavaScript APIs'
ms.date: 12/12/2019
---

# Office Scripts API reference

The Office Scripts API lets you automate common tasks in Excel on the web. Use this reference documentation to learn more about the classes, methods, and other types available for your scripts. All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.

## Common classes

The following list breaks down the basics of the Office Scripts object model. This shows the common classes and how they relate to one another.

- A [Workbook](/javascript/api/office-scripts/excel-scripts/excel.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel-scripts/excel.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel-scripts/excel.worksheetcollection).
- A [Worksheet](/javascript/api/office-scripts/excel-scripts/excel.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel-scripts/excel.range) objects.
- A [Range](/javascript/api/office-scripts/excel-scripts/excel.range) represents a group of contiguous cells.
- [Ranges](/javascript/api/office-scripts/excel-scripts/excel.range) are used to create and place [Tables](/javascript/api/office-scripts/excel-scripts/excel.table), [Charts](/javascript/api/office-scripts/excel-scripts/excel.chart), [Shapes*](/javascript/api/office-scripts/excel-scripts/excel.shape), and other data visualization or organization objects.
- A [Worksheet](/javascript/api/office-scripts/excel-scripts/excel.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel-scripts/excel.chartcollection)) that are present in the individual sheet.
- [Workbooks](/javascript/api/office-scripts/excel-scripts/excel. contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel-scripts/excel.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel-scripts/excel.workbook).

For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)

## See also

- [About Office Scripts](/office/dev/scripts/overview/overview)
- [Record, edit, and create Office Scripts in Excel on the web](/office/dev/scripts/tutorials/excel-office-scripts-tutorial)
- [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)
