declare namespace Excel {
    /**
     * run function
     */
    function run<T>(
        batch: (context: {
            sync: () => Promise<void>;
            workbook: Workbook;
        }) => Promise<T>
    ): Promise<T>;

    /**
     * Contains the result for methods that return primitive types.
     * The object's value property is retrieved from the document after `context.sync()` is invoked.
     */
    interface ClientResult<T> {
        /**
         * The value of the result that is retrieved from the document after `context.sync()` is invoked.
         */
        value: T;
    }

    /**
     * The RequestContext object facilitates requests to the Excel workbook.
     * The script and the Excel workbook run in separate processes.
     * The request context is used to synchronize data and actions between the script and the workbook.
     */
    interface RequestContext {
        readonly workbook: Workbook;
        sync(): Promise<void>;
    }

    //
    // Class
    //

    /**
     * Represents the Excel application that manages the workbook.
     */
    interface Application {
        /**
         * Returns the Excel calculation engine version used for the last full recalculation.
         */
        readonly calculationEngineVersion: number;

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        calculationMode:
            | CalculationMode
            | "Automatic"
            | "AutomaticExceptTables"
            | "Manual";

        /**
         * Returns the calculation state of the application. See Excel.CalculationState for details.
         */
        readonly calculationState:
            | CalculationState
            | "Done"
            | "Calculating"
            | "Pending";

        /**
         * Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.
         */
        readonly cultureInfo: CultureInfo;

        /**
         * Gets the string used as the decimal separator for numeric values. This is based on Excel's local settings.
         */
        readonly decimalSeparator: string;

        /**
         * Returns the Iterative Calculation settings.
         * In Excel on Windows and Mac, the settings will apply to the Excel Application.
         * In Excel on the web and other platforms, the settings will apply to the active workbook.
         */
        readonly iterativeCalculation: IterativeCalculation;

        /**
         * Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on Excel's local settings.
         */
        readonly thousandsSeparator: string;

        /**
         * Specifies if the system separators of Excel are enabled.
         * System separators include the decimal separator and thousands separator.
         */
        readonly useSystemSeparators: boolean;

        /**
         * Recalculate all currently opened workbooks in Excel.
         * @param calculationType Specifies the calculation type to use. See Excel.CalculationType for details.
         */
        calculate(
            calculationType:
                | CalculationType
                | "Recalculate"
                | "Full"
                | "FullRebuild"
        ): void;

        /**
         * Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.
         */
        suspendApiCalculationUntilNextSync(): void;

        /**
         * Suspends screen updating until the next `context.sync()` is called.
         *
         * **Note**: Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop). Repeated calls will cause the Excel window to flicker.
         */
        suspendScreenUpdatingUntilNextSync(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Iterative Calculation settings.
     */
    interface IterativeCalculation {
        /**
         * True if Excel will use iteration to resolve circular references.
         */
        enabled: boolean;

        /**
         * Specifies the maximum amount of change between each iteration as Excel resolves circular references.
         */
        maxChange: number;

        /**
         * Specifies the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        maxIteration: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
     * To learn more about the workbook object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks | Work with workbooks using the Excel JavaScript API}.
     */
    interface Workbook {
        /**
         * Represents the Excel application instance that contains this workbook.
         */
        readonly application: Application;

        /**
         * Specifies if the workbook is in autosave mode.
         */
        readonly autoSave: boolean;

        /**
         * Represents a collection of bindings that are part of the workbook.
         */
        readonly bindings: BindingCollection;

        /**
         * Returns a number about the version of Excel Calculation Engine.
         */
        readonly calculationEngineVersion: number;

        /**
         * True if all charts in the workbook are tracking the actual data points to which they are attached.
         * False if the charts track the index of the data points.
         */
        chartDataPointTrack: boolean;

        /**
         * Represents a collection of Comments associated with the workbook.
         */
        readonly comments: CommentCollection;

        /**
         * Represents the collection of custom XML parts contained by this workbook.
         */
        readonly customXmlParts: CustomXmlPartCollection;

        /**
         * Represents all data connections in the workbook.
         */
        readonly dataConnections: DataConnectionCollection;

        /**
         * Specifies if changes have been made since the workbook was last saved.
         * You can set this property to true if you want to close a modified workbook without either saving it or being prompted to save it.
         */
        isDirty: boolean;

        /**
         * Gets the workbook name.
         */
        readonly name: string;

        /**
         * Represents a collection of workbook scoped named items (named ranges and constants).
         */
        readonly names: NamedItemCollection;

        /**
         * Represents a collection of PivotTableStyles associated with the workbook.
         */
        readonly pivotTableStyles: PivotTableStyleCollection;

        /**
         * Represents a collection of PivotTables associated with the workbook.
         */
        readonly pivotTables: PivotTableCollection;

        /**
         * Specifies if the workbook has ever been saved locally or online.
         */
        readonly previouslySaved: boolean;

        /**
         * Gets the workbook properties.
         */
        readonly properties: DocumentProperties;

        /**
         * Returns the protection object for a workbook.
         */
        readonly protection: WorkbookProtection;

        /**
         * True if the workbook is open in Read-only mode.
         */
        readonly readOnly: boolean;

        /**
         * Represents a collection of SlicerStyles associated with the workbook.
         */
        readonly slicerStyles: SlicerStyleCollection;

        /**
         * Represents a collection of Slicers associated with the workbook.
         */
        readonly slicers: SlicerCollection;

        /**
         * Represents a collection of styles associated with the workbook.
         */
        readonly styles: StyleCollection;

        /**
         * Represents a collection of TableStyles associated with the workbook.
         */
        readonly tableStyles: TableStyleCollection;

        /**
         * Represents a collection of tables associated with the workbook.
         */
        readonly tables: TableCollection;

        /**
         * Represents a collection of TimelineStyles associated with the workbook.
         */
        readonly timelineStyles: TimelineStyleCollection;

        /**
         * True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
         * Data will permanently lose accuracy when switching this property from false to true.
         */
        usePrecisionAsDisplayed: boolean;

        /**
         * Represents a collection of worksheets associated with the workbook.
         */
        readonly worksheets: WorksheetCollection;

        /**
         * Gets the currently active cell from the workbook.
         */
        getActiveCell(): Range;

        /**
         * Gets the currently active chart in the workbook. If there is no active chart, an `ItemNotFound` exception is thrown.
         */
        getActiveChart(): Chart;

        /**
         * Gets the currently active chart in the workbook. If there is no active chart, a null object is returned.
         */
        getActiveChartOrNullObject(): Chart;

        /**
         * Gets the currently active slicer in the workbook. If there is no active slicer, an `ItemNotFound` exception is thrown.
         */
        getActiveSlicer(): Slicer;

        /**
         * Gets the currently active slicer in the workbook. If there is no active slicer, a null object is returned.
         */
        getActiveSlicerOrNullObject(): Slicer;

        /**
         * Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.
         */
        getSelectedRange(): Range;

        /**
         * Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.
         */
        getSelectedRanges(): RangeAreas;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the protection of a workbook object.
     */
    interface WorkbookProtection {
        /**
         * Specifies if the workbook is protected.
         */
        readonly protected: boolean;

        /**
         * Protects a workbook. Fails if the workbook has been protected.
         * @param password workbook protection password.
         */
        protect(password?: string): void;

        /**
         * Unprotects a workbook.
         * @param password workbook protection password.
         */
        unprotect(password?: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
     * To learn more about the worksheet object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets | Work with worksheets using the Excel JavaScript API}.
     */
    interface Worksheet {
        /**
         * Represents the AutoFilter object of the worksheet.
         */
        readonly autoFilter: AutoFilter;

        /**
         * Returns a collection of charts that are part of the worksheet.
         */
        readonly charts: ChartCollection;

        /**
         * Returns a collection of all the Comments objects on the worksheet.
         */
        readonly comments: CommentCollection;

        /**
         * Determines if Excel should recalculate the worksheet when necessary.
         * True if Excel recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.
         */
        enableCalculation: boolean;

        /**
         * Gets an object that can be used to manipulate frozen panes on the worksheet.
         */
        readonly freezePanes: WorksheetFreezePanes;

        /**
         * Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.
         */
        readonly horizontalPageBreaks: PageBreakCollection;

        /**
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
         */
        readonly id: string;

        /**
         * The display name of the worksheet.
         */
        name: string;

        /**
         * Collection of names scoped to the current worksheet.
         */
        readonly names: NamedItemCollection;

        /**
         * Gets the PageLayout object of the worksheet.
         */
        readonly pageLayout: PageLayout;

        /**
         * Collection of PivotTables that are part of the worksheet.
         */
        readonly pivotTables: PivotTableCollection;

        /**
         * The zero-based position of the worksheet within the workbook.
         */
        position: number;

        /**
         * Returns sheet protection object for a worksheet.
         */
        readonly protection: WorksheetProtection;

        /**
         * Returns the collection of all the Shape objects on the worksheet.
         */
        readonly shapes: ShapeCollection;

        /**
         * Specifies if gridlines are visible to the user.
         */
        showGridlines: boolean;

        /**
         * Specifies if headings are visible to the user.
         */
        showHeadings: boolean;

        /**
         * Returns a collection of slicers that are part of the worksheet.
         */
        readonly slicers: SlicerCollection;

        /**
         * Returns the standard (default) height of all the rows in the worksheet, in points.
         */
        readonly standardHeight: number;

        /**
         * Specifies the standard (default) width of all the columns in the worksheet.
         * One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.
         */
        standardWidth: number;

        /**
         * The tab color of the worksheet.
         * When retrieving the tab color, if the worksheet is invisible, the value will be null. If the worksheet is visible but the tab color is set to auto, an empty string will be returned. Otherwise, the property will be set to a color, in the form "#123456"
         * When setting the color, use an empty-string to set an "auto" color, or a real color otherwise.
         */
        tabColor: string;

        /**
         * Collection of tables that are part of the worksheet.
         */
        readonly tables: TableCollection;

        /**
         * Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.
         */
        readonly verticalPageBreaks: PageBreakCollection;

        /**
         * The Visibility of the worksheet.
         */
        visibility: SheetVisibility | "Visible" | "Hidden" | "VeryHidden";

        /**
         * Activate the worksheet in the Excel UI.
         */
        activate(): void;

        /**
         * Calculates all cells on a worksheet.
         * @param markAllDirty True, to mark all as dirty.
         */
        calculate(markAllDirty: boolean): void;

        /**
         * Copies a worksheet and places it at the specified position.
         * @param positionType The location in the workbook to place the newly created worksheet. The default value is "None", which inserts the worksheet at the beginning of the worksheet.
         * @param relativeTo The existing worksheet which determines the newly created worksheet's position. This is only needed if `positionType` is "Before" or "After".
         */
        copy(
            positionType?:
                | WorksheetPositionType
                | "None"
                | "Before"
                | "After"
                | "Beginning"
                | "End",
            relativeTo?: Worksheet
        ): Worksheet;

        /**
         * Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with an `InvalidOperation` exception. You should first change its visibility to hidden or visible before deleting it.
         */
        delete(): void;

        /**
         * Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.
         * @param text The string to find.
         * @param criteria Additional search criteria, including whether the search needs to match the entire cell or be case sensitive.
         */
        findAll(text: string, criteria: WorksheetSearchCriteria): RangeAreas;

        /**
         * Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.
         * @param text The string to find.
         * @param criteria Additional search criteria, including whether the search needs to match the entire cell or be case sensitive.
         */
        findAllOrNullObject(
            text: string,
            criteria: WorksheetSearchCriteria
        ): RangeAreas;

        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.
         * @param row The row number of the cell to be retrieved. Zero-indexed.
         * @param column the column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getNext(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getNextOrNullObject(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getPrevious(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getPreviousOrNullObject(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the range object, representing a single rectangular block of cells, specified by the address or name.
         * @param address Optional. The string representing the address or name of the range. For example, "A1:B2". If not specified, the entire worksheet range is returned.
         */
        getRange(address?: string): Range;

        /**
         * Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.
         * @param startRow Start row (zero-indexed).
         * @param startColumn Start column (zero-indexed).
         * @param rowCount Number of rows to include in the range.
         * @param columnCount Number of columns to include in the range.
         */
        getRangeByIndexes(
            startRow: number,
            startColumn: number,
            rowCount: number,
            columnCount: number
        ): Range;

        /**
         * Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.
         * @param address Optional. A string containing the comma-separated addresses or names of the individual ranges. For example, "A1:B2, A5:B5". If not specified, an RangeArea object for the entire worksheet is returned.
         */
        getRanges(address?: string): RangeAreas;

        /**
         * The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e. it will *not* throw an error).
         * @param valuesOnly Optional. If true, considers only cells with values as used cells (ignoring formatting).
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.
         * @param valuesOnly Optional. Considers only cells with values as used cells.
         */
        getUsedRangeOrNullObject(valuesOnly?: boolean): Range;

        /**
         * Finds and replaces the given string based on the criteria specified within the current worksheet.
         * @param text String to find.
         * @param replacement String to replace the original with.
         * @param criteria Additional Replace Criteria.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): ClientResult<number>;

        /**
         * Shows row or column groups by their outline levels.
         * Outlines group and summarize a list of data in the worksheet.
         * The `rowLevels` and `columnLevels` parameters specify how many levels of the outline will be displayed.
         * The acceptable argument range is between 0 and 8.
         * A value of 0 does not change the current display. A value greater than the current number of levels displays all the levels.
         * @param rowLevels The number of row levels of an outline to display.
         * @param columnLevels The number of column levels of an outline to display.
         */
        showOutlineLevels(rowLevels: number, columnLevels: number): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of worksheet objects that are part of the workbook.
     */
    interface WorksheetCollection {
        /**
         * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
         * @param name Optional. The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.
         */
        add(name?: string): Worksheet;

        /**
         * Gets the currently active worksheet in the workbook.
         */
        getActiveWorksheet(): Worksheet;

        /**
         * Gets the number of worksheets in the collection.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getCount(visibleOnly?: boolean): ClientResult<number>;

        /**
         * Gets the first worksheet in the collection.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getFirst(visibleOnly?: boolean): Worksheet;

        /**
         * Gets a worksheet object using its Name or ID.
         * @param key The Name or ID of the worksheet.
         */
        getItem(key: string): Worksheet;

        /**
         * Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.
         * @param key The Name or ID of the worksheet.
         */
        getItemOrNullObject(key: string): Worksheet;

        /**
         * Gets the last worksheet in the collection.
         * @param visibleOnly Optional. If true, considers only visible worksheets, skipping over any hidden ones.
         */
        getLast(visibleOnly?: boolean): Worksheet;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the protection of a sheet object.
     */
    interface WorksheetProtection {
        /**
         * Specifies the protection options for the worksheet.
         */
        readonly options: WorksheetProtectionOptions;

        /**
         * Specifies if the worksheet is protected.
         */
        readonly protected: boolean;

        /**
         * Protects a worksheet. Fails if the worksheet has already been protected.
         * @param options Optional. Sheet protection options.
         * @param password Optional. Sheet protection password.
         */
        protect(options?: WorksheetProtectionOptions, password?: string): void;

        /**
         * Unprotects a worksheet.
         * @param password sheet protection password.
         */
        unprotect(password?: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface WorksheetFreezePanes {
        /**
         * Sets the frozen cells in the active worksheet view.
         * The range provided corresponds to cells that will be frozen in the top- and left-most pane.
         * @param frozenRange A range that represents the cells to be frozen, or null to remove all frozen panes.
         */
        freezeAt(frozenRange: Range | string): void;

        /**
         * Freeze the first column(s) of the worksheet in place.
         * @param count Optional number of columns to freeze, or zero to unfreeze all columns
         */
        freezeColumns(count?: number): void;

        /**
         * Freeze the top row(s) of the worksheet in place.
         * @param count Optional number of rows to freeze, or zero to unfreeze all rows
         */
        freezeRows(count?: number): void;

        /**
         * Gets a range that describes the frozen cells in the active worksheet view.
         * The frozen range is corresponds to cells that are frozen in the top- and left-most pane.
         */
        getLocation(): Range;

        /**
         * Gets a range that describes the frozen cells in the active worksheet view.
         * The frozen range is corresponds to cells that are frozen in the top- and left-most pane.
         * If there is no frozen pane, returns a null object.
         */
        getLocationOrNullObject(): Range;

        /**
         * Removes all frozen panes in the worksheet.
         */
        unfreeze(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
     * To learn more about how ranges are used throughout the API, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges | Work with ranges using the Excel JavaScript API}
     * and {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges-advanced | Work with ranges using the Excel JavaScript API (advanced)}.
     */
    interface Range {
        /**
         * Specifies the range reference in A1-style. Address value will contain the Sheet reference (e.g., "Sheet1!A1:B4").
         */
        readonly address: string;

        /**
         * Specifies the range reference for the specified range in the language of the user.
         */
        readonly addressLocal: string;

        /**
         * Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).
         */
        readonly cellCount: number;

        /**
         * Specifies the total number of columns in the range.
         */
        readonly columnCount: number;

        /**
         * Represents if all columns of the current range are hidden.
         */
        columnHidden: boolean;

        /**
         * Specifies the column number of the first cell in the range. Zero-indexed.
         */
        readonly columnIndex: number;

        /**
         * The collection of ConditionalFormats that intersect the range.
         */
        readonly conditionalFormats: ConditionalFormatCollection;

        /**
         * Returns a data validation object.
         */
        readonly dataValidation: DataValidation;

        /**
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
         */
        readonly format: RangeFormat;

        /**
         * Represents the formula in A1-style notation.
         */
        formulas: any[][];

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         */
        formulasLocal: any[][];

        /**
         * Represents the formula in R1C1-style notation.
         */
        formulasR1C1: any[][];

        /**
         * Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range.
         */
        readonly height: number;

        /**
         * Represents if all cells of the current range are hidden.
         */
        readonly hidden: boolean;

        /**
         * Represents the hyperlink for the current range.
         */
        hyperlink: RangeHyperlink;

        /**
         * Represents if the current range is an entire column.
         */
        readonly isEntireColumn: boolean;

        /**
         * Represents if the current range is an entire row.
         */
        readonly isEntireRow: boolean;

        /**
         * Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range.
         */
        readonly left: number;

        /**
         * Represents the data type state of each cell.
         */
        readonly linkedDataTypeState: LinkedDataTypeState[][];

        /**
         * Represents Excel's number format code for the given range.
         */
        numberFormat: any[][];

        /**
         * Represents Excel's number format code for the given range, based on the language settings of the user.​
         * Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property.
         * Any returned text uses the locally-formatted strings based on the language specified in the system settings.
         */
        numberFormatLocal: any[][];

        /**
         * Returns the total number of rows in the range.
         */
        readonly rowCount: number;

        /**
         * Represents if all rows of the current range are hidden.
         */
        rowHidden: boolean;

        /**
         * Returns the row number of the first cell in the range. Zero-indexed.
         */
        readonly rowIndex: number;

        /**
         * Represents the range sort of the current range.
         */
        readonly sort: RangeSort;

        /**
         * Represents the style of the current range.
         * If the styles of the cells are inconsistent, null will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the BuiltInStyle enum will be returned.
         */
        style: string;

        /**
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API.
         */
        readonly text: string[][];

        /**
         * Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range.
         */
        readonly top: number;

        /**
         * Specifies the type of data in each cell.
         */
        readonly valueTypes: RangeValueType[][];

        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        values: any[][];

        /**
         * Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range.
         */
        readonly width: number;

        /**
         * The worksheet containing the current range.
         */
        readonly worksheet: Worksheet;

        /**
         * Fills range from the current range to the destination range using the specified AutoFill logic.
         * The destination range can be null, or can extend the source either horizontally or vertically.
         * Discontiguous ranges are not supported.
         *
         * For more information, read {@link https://support.office.com/article/video-use-autofill-and-flash-fill-2e79a709-c814-4b27-8bc2-c4dc84d49464 | Use AutoFill and Flash Fill}.
         * @param destinationRange The destination range to autofill. If the destination range is null, data is filled out based on the surrounding cells (which is the behavior when double-clicking the UI’s range fill handle).
         * @param autoFillType The type of autofill. Specifies how the destination range is to be filled, based on the contents of the current range. Default is "FillDefault".
         */
        autoFill(
            destinationRange?: Range | string,
            autoFillType?:
                | AutoFillType
                | "FillDefault"
                | "FillCopy"
                | "FillSeries"
                | "FillFormats"
                | "FillValues"
                | "FillDays"
                | "FillWeekdays"
                | "FillMonths"
                | "FillYears"
                | "LinearTrend"
                | "GrowthTrend"
                | "FlashFill"
        ): void;

        /**
         * Calculates a range of cells on a worksheet.
         */
        calculate(): void;

        /**
         * Clear range values, format, fill, border, etc.
         * @param applyTo Optional. Determines the type of clear action. See Excel.ClearApplyTo for details.
         */
        clear(
            applyTo?:
                | ClearApplyTo
                | "All"
                | "Formats"
                | "Contents"
                | "Hyperlinks"
                | "RemoveHyperlinks"
        ): void;

        /**
         * Converts the range cells with datatypes into text.
         */
        convertDataTypeToText(): void;

        /**
         * Converts the range cells into linked datatype in the worksheet.
         * @param serviceID The Service ID which will be used to query the data.
         * @param languageCulture Language Culture to query the service for.
         */
        convertToLinkedDataType(
            serviceID: number,
            languageCulture: string
        ): void;

        /**
         * Copies cell data or formatting from the source range or RangeAreas to the current range.
         * The destination range can be a different size than the source range or RangeAreas. The destination will be expanded automatically if it is smaller than the source.
         * @param sourceRange The source range or RangeAreas to copy from. When the source RangeAreas has multiple ranges, their form must be able to be created by removing full rows or columns from a rectangular range.
         * @param copyType The type of cell data or formatting to copy over. Default is "All".
         * @param skipBlanks True if to skip blank cells in the source range. Default is false.
         * @param transpose True if to transpose the cells in the destination range. Default is false.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?:
                | RangeCopyType
                | "All"
                | "Formulas"
                | "Values"
                | "Formats",
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Deletes the cells associated with the range.
         * @param shift Specifies which way to shift the cells. See Excel.DeleteShiftDirection for details.
         */
        delete(shift: DeleteShiftDirection | "Up" | "Left"): void;

        /**
         * Finds the given string based on the criteria specified.
         * If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell.
         * @param text The string to find.
         * @param criteria Additional search criteria, including the search direction and whether the search needs to match the entire cell or be case sensitive.
         */
        find(text: string, criteria: SearchCriteria): Range;

        /**
         * Finds the given string based on the criteria specified.
         * If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell.
         * If there are no matches, this function will return a null object.
         * @param text The string to find.
         * @param criteria Additional search criteria, including the search direction and whether the search needs to match the entire cell or be case sensitive.
         */
        findOrNullObject(text: string, criteria: SearchCriteria): Range;

        /**
         * Does FlashFill to current range.Flash Fill will automatically fills data when it senses a pattern, so the range must be single column range and have data around in order to find pattern.
         */
        flashFill(): void;

        /**
         * Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.
         * @param numRows The number of rows of the new range size.
         * @param numColumns The number of columns of the new range size.
         */
        getAbsoluteResizedRange(numRows: number, numColumns: number): Range;

        /**
         * Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".
         * @param anotherRange The range object or address or range name.
         */
        getBoundingRect(anotherRange: Range | string): Range;

        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         * @param row Row number of the cell to be retrieved. Zero-indexed.
         * @param column Column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets a column contained in the range.
         * @param column Column number of the range to be retrieved. Zero-indexed.
         */
        getColumn(column: number): Range;

        /**
         * Gets a certain number of columns to the right of the current Range object.
         * @param count Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getColumnsAfter(count?: number): Range;

        /**
         * Gets a certain number of columns to the left of the current Range object.
         * @param count Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getColumnsBefore(count?: number): Range;

        /**
         * Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").
         */
        getEntireColumn(): Range;

        /**
         * Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").
         */
        getEntireRow(): Range;

        /**
         * Renders the range as a base64-encoded png image.
         */
        getImage(): ClientResult<string>;

        /**
         * Gets the range object that represents the rectangular intersection of the given ranges.
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         */
        getIntersection(anotherRange: Range | string): Range;

        /**
         * Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         */
        getIntersectionOrNullObject(anotherRange: Range | string): Range;

        /**
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         */
        getLastCell(): Range;

        /**
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         */
        getLastColumn(): Range;

        /**
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         */
        getLastRow(): Range;

        /**
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
         * @param rowOffset The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Range;

        /**
         * Gets a scoped collection of PivotTables that overlap with the range.
         * @param fullyContained If true, returns only PivotTables that are fully contained within the range bounds. The default value is false.
         */
        getPivotTables(fullyContained?: boolean): PivotTableScopedCollection;

        /**
         * Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
         * @param deltaRows The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         * @param deltaColumns The number of columns by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         */
        getResizedRange(deltaRows: number, deltaColumns: number): Range;

        /**
         * Gets a row contained in the range.
         * @param row Row number of the range to be retrieved. Zero-indexed.
         */
        getRow(row: number): Range;

        /**
         * Gets a certain number of rows above the current Range object.
         * @param count Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getRowsAbove(count?: number): Range;

        /**
         * Gets a certain number of rows below the current Range object.
         * @param count Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getRowsBelow(count?: number): Range;

        /**
         * Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.
         * If no special cells are found, an ItemNotFound error will be thrown.
         * @param cellType The type of cells to include.
         * @param cellValueType If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCells(
            cellType:
                | SpecialCellType
                | "ConditionalFormats"
                | "DataValidations"
                | "Blanks"
                | "Constants"
                | "Formulas"
                | "SameConditionalFormat"
                | "SameDataValidation"
                | "Visible",
            cellValueType?:
                | SpecialCellValueType
                | "All"
                | "Errors"
                | "ErrorsLogical"
                | "ErrorsNumbers"
                | "ErrorsText"
                | "ErrorsLogicalNumber"
                | "ErrorsLogicalText"
                | "ErrorsNumberText"
                | "Logical"
                | "LogicalNumbers"
                | "LogicalText"
                | "LogicalNumbersText"
                | "Numbers"
                | "NumbersText"
                | "Text"
        ): RangeAreas;

        /**
         * Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.
         * If no special cells are found, a null object will be returned.
         * @param cellType The type of cells to include.
         * @param cellValueType If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCellsOrNullObject(
            cellType:
                | SpecialCellType
                | "ConditionalFormats"
                | "DataValidations"
                | "Blanks"
                | "Constants"
                | "Formulas"
                | "SameConditionalFormat"
                | "SameDataValidation"
                | "Visible",
            cellValueType?:
                | SpecialCellValueType
                | "All"
                | "Errors"
                | "ErrorsLogical"
                | "ErrorsNumbers"
                | "ErrorsText"
                | "ErrorsLogicalNumber"
                | "ErrorsLogicalText"
                | "ErrorsNumberText"
                | "Logical"
                | "LogicalNumbers"
                | "LogicalText"
                | "LogicalNumbersText"
                | "Numbers"
                | "NumbersText"
                | "Text"
        ): RangeAreas;

        /**
         * Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
         */
        getSurroundingRegion(): Range;

        /**
         * Gets a scoped collection of tables that overlap with the range.
         * @param fullyContained If true, returns only tables that are fully contained within the range bounds. The default value is false.
         */
        getTables(fullyContained?: boolean): TableScopedCollection;

        /**
         * Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.
         * @param valuesOnly Considers only cells with values as used cells.
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.
         * @param valuesOnly Considers only cells with values as used cells.
         */
        getUsedRangeOrNullObject(valuesOnly?: boolean): Range;

        /**
         * Represents the visible rows of the current range.
         */
        getVisibleView(): RangeView;

        /**
         * Groups columns and rows for an outline.
         * @param groupOption Specifies how the range can be grouped by rows or columns.
         * An `InvalidArgument` error is thrown when the group option differs from the range's
         * `isEntireRow` or `isEntireColumn` property (i.e., `range.isEntireRow` is true and `groupOption` is "ByColumns"
         * or `range.isEntireColumn` is true and `groupOption` is "ByRows").
         */
        group(groupOption: GroupOption | "ByRows" | "ByColumns"): void;

        /**
         * Hide details of the row or column group.
         * @param groupOption Specifies whether to hide details of grouped rows or grouped columns.
         */
        hideGroupDetails(
            groupOption: GroupOption | "ByRows" | "ByColumns"
        ): void;

        /**
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
         * @param shift Specifies which way to shift the cells. See Excel.InsertShiftDirection for details.
         */
        insert(shift: InsertShiftDirection | "Down" | "Right"): Range;

        /**
         * Merge the range cells into one region in the worksheet.
         * @param across Optional. Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.
         */
        merge(across?: boolean): void;

        /**
         * Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.
         * The destination range will be expanded automatically if it is smaller than the current range. Any cells in the destination range that are outside of the original range's area are not changed.
         * @param destinationRange destinationRange Specifies the range to where the information in this range will be moved.
         */
        moveTo(destinationRange: Range | string): void;

        /**
         * Removes duplicate values from the range specified by the columns.
         * @param columns The columns inside the range that may contain duplicates. At least one column needs to be specified. Zero-indexed.
         * @param includesHeader True if the input data contains header. Default is false.
         */
        removeDuplicates(
            columns: number[],
            includesHeader: boolean
        ): RemoveDuplicatesResult;

        /**
         * Finds and replaces the given string based on the criteria specified within the current range.
         * @param text String to find.
         * @param replacement String to replace the original with.
         * @param criteria Additional Replace Criteria.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): ClientResult<number>;

        /**
         * Selects the specified range in the Excel UI.
         */
        select(): void;

        /**
         * Set a range to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        /**
         * Displays the card for an active cell if it has rich value content.
         */
        showCard(): void;

        /**
         * Show details of the row or column group.
         * @param groupOption Specifies whether to show details of grouped rows or grouped columns.
         */
        showGroupDetails(
            groupOption: GroupOption | "ByRows" | "ByColumns"
        ): void;

        /**
         * Ungroups columns and rows for an outline.
         * @param groupOption Specifies how the range can be ungrouped by rows or columns.
         */
        ungroup(groupOption: GroupOption | "ByRows" | "ByColumns"): void;

        /**
         * Unmerge the range cells into separate cells.
         */
        unmerge(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * RangeAreas represents a collection of one or more rectangular ranges in the same worksheet.
     * To learn how to use discontinguous ranges, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-multiple-ranges | Work with multiple ranges simultaneously in Excel add-ins}.
     */
    interface RangeAreas {
        /**
         * Returns the RangeAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!A1:B4, Sheet1!D1:D4").
         */
        readonly address: string;

        /**
         * Returns the RangeAreas reference in the user locale.
         */
        readonly addressLocal: string;

        /**
         * Returns the number of rectangular ranges that comprise this RangeAreas object.
         */
        readonly areaCount: number;

        /**
         * Returns a collection of rectangular ranges that comprise this RangeAreas object.
         */
        readonly areas: RangeCollection;

        /**
         * Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647).
         */
        readonly cellCount: number;

        /**
         * Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object.
         */
        readonly conditionalFormats: ConditionalFormatCollection;

        /**
         * Returns a dataValidation object for all ranges in the RangeAreas.
         */
        readonly dataValidation: DataValidation;

        /**
         * Returns a RangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object.
         */
        readonly format: RangeFormat;

        /**
         * Specifies if all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z").
         */
        readonly isEntireColumn: boolean;

        /**
         * Specifies if all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7").
         */
        readonly isEntireRow: boolean;

        /**
         * Represents the style for all ranges in this RangeAreas object.
         * If the styles of the cells are inconsistent, null will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the BuiltInStyle enum will be returned.
         */
        style: string;

        /**
         * Returns the worksheet for the current RangeAreas.
         */
        readonly worksheet: Worksheet;

        /**
         * Calculates all cells in the RangeAreas.
         */
        calculate(): void;

        /**
         * Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.
         * @param applyTo Optional. Determines the type of clear action. See Excel.ClearApplyTo for details. Default is "All".
         */
        clear(
            applyTo?:
                | ClearApplyTo
                | "All"
                | "Formats"
                | "Contents"
                | "Hyperlinks"
                | "RemoveHyperlinks"
        ): void;

        /**
         * Converts all cells in the RangeAreas with datatypes into text.
         */
        convertDataTypeToText(): void;

        /**
         * Converts all cells in the RangeAreas into linked datatype.
         * @param serviceID The Service ID which will be used to query the data.
         * @param languageCulture Language Culture to query the service for.
         */
        convertToLinkedDataType(
            serviceID: number,
            languageCulture: string
        ): void;

        /**
         * Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.
         * The destination rangeAreas can be a different size than the source range or RangeAreas. The destination will be expanded automatically if it is smaller than the source.
         * @param sourceRange The source range or RangeAreas to copy from. When the source RangeAreas has multiple ranges, their form must able to be created by removing full rows or columns from a rectangular range.
         * @param copyType The type of cell data or formatting to copy over. Default is "All".
         * @param skipBlanks True if to skip blank cells in the source range or RangeAreas. Default is false.
         * @param transpose True if to transpose the cells in the destination RangeAreas. Default is false.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?:
                | RangeCopyType
                | "All"
                | "Formulas"
                | "Values"
                | "Formats",
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").
         */
        getEntireColumn(): RangeAreas;

        /**
         * Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").
         */
        getEntireRow(): RangeAreas;

        /**
         * Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.
         * @param anotherRange The range, RangeAreas object or range address that will be used to determine the intersection.
         */
        getIntersection(anotherRange: Range | RangeAreas | string): RangeAreas;

        /**
         * Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.
         * @param anotherRange The range, RangeAreas, or address that will be used to determine the intersection.
         */
        getIntersectionOrNullObject(
            anotherRange: Range | RangeAreas | string
        ): RangeAreas;

        /**
         * Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.
         * @param rowOffset The number of rows (positive, negative, or 0) by which the RangeAreas is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the RangeAreas is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRangeAreas(
            rowOffset: number,
            columnOffset: number
        ): RangeAreas;

        /**
         * Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.
         * @param cellType The type of cells to include.
         * @param cellValueType If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCells(
            cellType:
                | SpecialCellType
                | "ConditionalFormats"
                | "DataValidations"
                | "Blanks"
                | "Constants"
                | "Formulas"
                | "SameConditionalFormat"
                | "SameDataValidation"
                | "Visible",
            cellValueType?:
                | SpecialCellValueType
                | "All"
                | "Errors"
                | "ErrorsLogical"
                | "ErrorsNumbers"
                | "ErrorsText"
                | "ErrorsLogicalNumber"
                | "ErrorsLogicalText"
                | "ErrorsNumberText"
                | "Logical"
                | "LogicalNumbers"
                | "LogicalText"
                | "LogicalNumbersText"
                | "Numbers"
                | "NumbersText"
                | "Text"
        ): RangeAreas;

        /**
         * Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.
         * @param cellType The type of cells to include.
         * @param cellValueType If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCellsOrNullObject(
            cellType:
                | SpecialCellType
                | "ConditionalFormats"
                | "DataValidations"
                | "Blanks"
                | "Constants"
                | "Formulas"
                | "SameConditionalFormat"
                | "SameDataValidation"
                | "Visible",
            cellValueType?:
                | SpecialCellValueType
                | "All"
                | "Errors"
                | "ErrorsLogical"
                | "ErrorsNumbers"
                | "ErrorsText"
                | "ErrorsLogicalNumber"
                | "ErrorsLogicalText"
                | "ErrorsNumberText"
                | "Logical"
                | "LogicalNumbers"
                | "LogicalText"
                | "LogicalNumbersText"
                | "Numbers"
                | "NumbersText"
                | "Text"
        ): RangeAreas;

        /**
         * Returns a scoped collection of tables that overlap with any range in this RangeAreas object.
         * @param fullyContained If true, returns only tables that are fully contained within the range bounds. Default is false.
         */
        getTables(fullyContained?: boolean): TableScopedCollection;

        /**
         * Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.
         * If there are no used cells within the RangeAreas, the ItemNotFound error will be thrown.
         * @param valuesOnly Whether to only consider cells with values as used cells. Default is false.
         */
        getUsedRangeAreas(valuesOnly?: boolean): RangeAreas;

        /**
         * Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.
         * If there are no used cells within the RangeAreas, a null object will be returned.
         * @param valuesOnly Whether to only consider cells with values as used cells.
         */
        getUsedRangeAreasOrNullObject(valuesOnly?: boolean): RangeAreas;

        /**
         * Sets the RangeAreas to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * RangeView represents a set of visible cells of the parent range.
     */
    interface RangeView {
        /**
         * Represents the cell addresses of the RangeView.
         */
        readonly cellAddresses: any[][];

        /**
         * The number of visible columns.
         */
        readonly columnCount: number;

        /**
         * Represents the formula in A1-style notation.
         */
        formulas: any[][];

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         */
        formulasLocal: any[][];

        /**
         * Represents the formula in R1C1-style notation.
         */
        formulasR1C1: any[][];

        /**
         * Returns a value that represents the index of the RangeView.
         */
        readonly index: number;

        /**
         * Represents Excel's number format code for the given cell.
         */
        numberFormat: any[][];

        /**
         * The number of visible rows.
         */
        readonly rowCount: number;

        /**
         * Represents a collection of range views associated with the range.
         */
        readonly rows: RangeViewCollection;

        /**
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API.
         */
        readonly text: string[][];

        /**
         * Represents the type of data of each cell.
         */
        readonly valueTypes: RangeValueType[][];

        /**
         * Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        values: any[][];

        /**
         * Gets the parent range associated with the current RangeView.
         */
        getRange(): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of RangeView objects.
     */
    interface RangeViewCollection {
        /**
         * Gets the number of RangeView objects in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a RangeView Row via its index. Zero-Indexed.
         * @param index Index of the visible row.
         */
        getItemAt(index: number): RangeView;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of key-value pair setting objects that are part of the workbook. The scope is limited to per file and add-in (task-pane or content) combination.
     */
    interface SettingCollection {
        /**
         * Sets or adds the specified setting to the workbook.
         * @param key The Key of the new setting.
         * @param value The Value for the new setting.
         */
        add(
            key: string,
            value: string | number | boolean | Date | Array<any> | any
        ): Setting;

        /**
         * Gets the number of Settings in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a Setting entry via the key.
         * @param key Key of the setting.
         */
        getItem(key: string): Setting;

        /**
         * Gets a Setting entry via the key. If the Setting does not exist, will return a null object.
         * @param key The key of the setting.
         */
        getItemOrNullObject(key: string): Setting;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Setting represents a key-value pair of a setting persisted to the document (per file per add-in). These custom key-value pair can be used to store state or lifecycle information needed by the content or task-pane add-in. Note that settings are persisted in the document and hence it is not a place to store any sensitive or protected information such as user information and password.
     */
    interface Setting {
        /**
         * The key that represents the id of the Setting.
         */
        readonly key: string;

        /**
         * Represents the value stored for this setting.
         */
        value: any;

        /**
         * Deletes the setting.
         */
        delete(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A collection of all the NamedItem objects that are part of the workbook or worksheet, depending on how it was reached.
     */
    interface NamedItemCollection {
        /**
         * Adds a new name to the collection of the given scope.
         * @param name The name of the named item.
         * @param reference The formula or the range that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        add(
            name: string,
            reference: Range | string,
            comment?: string
        ): NamedItem;

        /**
         * Adds a new name to the collection of the given scope using the user's locale for the formula.
         * @param name The "name" of the named item.
         * @param formula The formula in the user's locale that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        addFormulaLocal(
            name: string,
            formula: string,
            comment?: string
        ): NamedItem;

        /**
         * Gets the number of named items in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a NamedItem object using its name.
         * @param name Nameditem name.
         */
        getItem(name: string): NamedItem;

        /**
         * Gets a NamedItem object using its name. If the nameditem object does not exist, will return a null object.
         * @param name Nameditem name.
         */
        getItemOrNullObject(name: string): NamedItem;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
     */
    interface NamedItem {
        /**
         * Returns an object containing values and types of the named item.
         */
        readonly arrayValues: NamedItemArrayValues;

        /**
         * Specifies the comment associated with this name.
         */
        comment: string;

        /**
         * The formula of the named item. Formula always starts with a '=' sign.
         */
        formula: any;

        /**
         * The name of the object.
         */
        readonly name: string;

        /**
         * Specifies if the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook.
         */
        readonly scope: NamedItemScope | "Worksheet" | "Workbook";

        /**
         * Specifies the type of the value returned by the name's formula. See Excel.NamedItemType for details.
         */
        readonly type:
            | NamedItemType
            | "String"
            | "Integer"
            | "Double"
            | "Boolean"
            | "Range"
            | "Error"
            | "Array";

        /**
         * Represents the value computed by the name's formula. For a named range, will return the range address.
         */
        readonly value: any;

        /**
         * Specifies if the object is visible.
         */
        visible: boolean;

        /**
         * Returns the worksheet on which the named item is scoped to. Throws an error if the item is scoped to the workbook instead.
         */
        readonly worksheet: Worksheet;

        /**
         * Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.
         */
        readonly worksheetOrNullObject: Worksheet;

        /**
         * Deletes the given name.
         */
        delete(): void;

        /**
         * Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.
         */
        getRange(): Range;

        /**
         * Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.
         */
        getRangeOrNullObject(): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an object containing values and types of a named item.
     */
    interface NamedItemArrayValues {
        /**
         * Represents the types for each item in the named item array
         */
        readonly types: RangeValueType[][];

        /**
         * Represents the values of each item in the named item array.
         */
        readonly values: any[][];

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an Office.js binding that is defined in the workbook.
     */
    interface Binding {
        /**
         * Represents binding identifier.
         */
        readonly id: string;

        /**
         * Returns the type of the binding. See Excel.BindingType for details.
         */
        readonly type: BindingType | "Range" | "Table" | "Text";

        /**
         * Deletes the binding.
         */
        delete(): void;

        /**
         * Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getRange(): Range;

        /**
         * Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getTable(): Table;

        /**
         * Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
         */
        getText(): ClientResult<string>;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the collection of all the binding objects that are part of the workbook.
     */
    interface BindingCollection {
        /**
         * Returns the number of bindings in the collection.
         */
        readonly count: number;

        /**
         * Add a new binding to a particular Range.
         * @param range Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         */
        add(
            range: Range | string,
            bindingType: BindingType | "Range" | "Table" | "Text",
            id: string
        ): Binding;

        /**
         * Add a new binding based on a named item in the workbook.
         * If the named item references to multiple areas, the "InvalidReference" error will be returned.
         * @param name Name from which to create binding.
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         */
        addFromNamedItem(
            name: string,
            bindingType: BindingType | "Range" | "Table" | "Text",
            id: string
        ): Binding;

        /**
         * Add a new binding based on the current selection.
         * If the selection has multiple areas, the "InvalidReference" error will be returned.
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         */
        addFromSelection(
            bindingType: BindingType | "Range" | "Table" | "Text",
            id: string
        ): Binding;

        /**
         * Gets the number of bindings in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a binding object by ID.
         * @param id Id of the binding object to be retrieved.
         */
        getItem(id: string): Binding;

        /**
         * Gets a binding object based on its position in the items array.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Binding;

        /**
         * Gets a binding object by ID. If the binding object does not exist, will return a null object.
         * @param id Id of the binding object to be retrieved.
         */
        getItemOrNullObject(id: string): Binding;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
     */
    interface TableCollection {
        /**
         * Returns the number of tables in the workbook.
         */
        readonly count: number;

        /**
         * Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         * @param address A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.
         * @param hasHeaders Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.
         */
        add(address: Range | string, hasHeaders: boolean): Table;

        /**
         * Gets the number of tables in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a table by Name or ID.
         * @param key Name or ID of the table to be retrieved.
         */
        getItem(key: string): Table;

        /**
         * Gets a table based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Table;

        /**
         * Gets a table by Name or ID. If the table does not exist, will return a null object.
         * @param key Name or ID of the table to be retrieved.
         */
        getItemOrNullObject(key: string): Table;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a scoped collection of tables. For each table its top-left corner is considered its anchor location and the tables are sorted top to bottom and then left to right.
     */
    interface TableScopedCollection {
        /**
         * Gets the number of tables in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.
         */
        getFirst(): Table;

        /**
         * Gets a table by Name or ID.
         * @param key Name or ID of the table to be retrieved.
         */
        getItem(key: string): Table;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an Excel table.
     * To learn more about the table object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables | Work with tables using the Excel JavaScript API}.
     */
    interface Table {
        /**
         * Represents the AutoFilter object of the table.
         */
        readonly autoFilter: AutoFilter;

        /**
         * Represents a collection of all the columns in the table.
         */
        readonly columns: TableColumnCollection;

        /**
         * Specifies if the first column contains special formatting.
         */
        highlightFirstColumn: boolean;

        /**
         * Specifies if the last column contains special formatting.
         */
        highlightLastColumn: boolean;

        /**
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
         */
        readonly id: string;

        /**
         * Returns a numeric id.
         */
        readonly legacyId: string;

        /**
         * Name of the table.
         *
         * The set name of the table must follow the guidelines specified in the {@link https://support.office.com/article/Rename-an-Excel-table-FBF49A4F-82A3-43EB-8BA2-44D21233B114 | Rename an Excel table} article.
         */
        name: string;

        /**
         * Represents a collection of all the rows in the table.
         */
        readonly rows: TableRowCollection;

        /**
         * Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
         */
        showBandedColumns: boolean;

        /**
         * Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
         */
        showBandedRows: boolean;

        /**
         * Specifies if the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
         */
        showFilterButton: boolean;

        /**
         * Specifies if the header row is visible. This value can be set to show or remove the header row.
         */
        showHeaders: boolean;

        /**
         * Specifies if the total row is visible. This value can be set to show or remove the total row.
         */
        showTotals: boolean;

        /**
         * Represents the sorting for the table.
         */
        readonly sort: TableSort;

        /**
         * Constant value that represents the Table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
         */
        style: string;

        /**
         * The worksheet containing the current table.
         */
        readonly worksheet: Worksheet;

        /**
         * Clears all the filters currently applied on the table.
         */
        clearFilters(): void;

        /**
         * Converts the table into a normal range of cells. All data is preserved.
         */
        convertToRange(): Range;

        /**
         * Deletes the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the table.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the range object associated with header row of the table.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire table.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with totals row of the table.
         */
        getTotalRowRange(): Range;

        /**
         * Reapplies all the filters currently on the table.
         */
        reapplyFilters(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the columns that are part of the table.
     */
    interface TableColumnCollection {
        /**
         * Returns the number of columns in the table.
         */
        readonly count: number;

        /**
         * Adds a new column to the table.
         * @param index Optional. Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.
         * @param values Optional. A 2-dimensional array of unformatted values of the table column.
         * @param name Optional. Specifies the name of the new column. If null, the default name will be used.
         */
        add(
            index?: number,
            values?:
                | Array<Array<boolean | string | number>>
                | boolean
                | string
                | number,
            name?: string
        ): TableColumn;

        /**
         * Gets the number of columns in the table.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a column object by Name or ID.
         * @param key Column Name or ID.
         */
        getItem(key: number | string): TableColumn;

        /**
         * Gets a column based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): TableColumn;

        /**
         * Gets a column object by Name or ID. If the column does not exist, will return a null object.
         * @param key Column Name or ID.
         */
        getItemOrNullObject(key: number | string): TableColumn;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a column in a table.
     */
    interface TableColumn {
        /**
         * Retrieve the filter applied to the column.
         */
        readonly filter: Filter;

        /**
         * Returns a unique key that identifies the column within the table.
         */
        readonly id: number;

        /**
         * Returns the index number of the column within the columns collection of the table. Zero-indexed.
         */
        readonly index: number;

        /**
         * Specifies the name of the table column.
         */
        name: string;

        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        values: any[][];

        /**
         * Deletes the column from the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the column.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the range object associated with the header row of the column.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire column.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with the totals row of the column.
         */
        getTotalRowRange(): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the rows that are part of the table.
     *
     * Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
     * a TableRow object represent the physical location of the table row, but not the data.
     * That is, if the data is sorted or if new rows are added, a table row will continue
     * to point at the index for which it was created.
     */
    interface TableRowCollection {
        /**
         * Returns the number of rows in the table.
         */
        readonly count: number;

        /**
         * Adds one or more rows to the table. The return object will be the top of the newly added row(s).
         *
         * Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
         * a TableRow object represent the physical location of the table row, but not the data.
         * That is, if the data is sorted or if new rows are added, a table row will continue
         * to point at the index for which it was created.
         * @param index Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values Optional. A 2-dimensional array of unformatted values of the table row.
         */
        add(
            index?: number,
            values?:
                | Array<Array<boolean | string | number>>
                | boolean
                | string
                | number
        ): TableRow;

        /**
         * Gets the number of rows in the table.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a row based on its position in the collection.
         *
         * Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
         * a TableRow object represent the physical location of the table row, but not the data.
         * That is, if the data is sorted or if new rows are added, a table row will continue
         * to point at the index for which it was created.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): TableRow;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a row in a table.
     *
     * Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
     * a TableRow object represent the physical location of the table row, but not the data.
     * That is, if the data is sorted or if new rows are added, a table row will continue
     * to point at the index for which it was created.
     */
    interface TableRow {
        /**
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        values: any[][];

        /**
         * Deletes the row from the table.
         */
        delete(): void;

        /**
         * Returns the range object associated with the entire row.
         */
        getRange(): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the data validation applied to the current range.
     * To learn more about the data validation object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation | Add data validation to Excel ranges}.
     */
    interface DataValidation {
        /**
         * Error alert when user enters invalid data.
         */
        errorAlert: DataValidationErrorAlert;

        /**
         * Specifies if data validation will be performed on blank cells, it defaults to true.
         */
        ignoreBlanks: boolean;

        /**
         * Prompt when users select a cell.
         */
        prompt: DataValidationPrompt;

        /**
         * Data validation rule that contains different type of data validation criteria.
         */
        rule: DataValidationRule;

        /**
         * Type of the data validation, see Excel.DataValidationType for details.
         */
        readonly type:
            | DataValidationType
            | "None"
            | "WholeNumber"
            | "Decimal"
            | "List"
            | "Date"
            | "Time"
            | "TextLength"
            | "Custom"
            | "Inconsistent"
            | "MixedCriteria";

        /**
         * Represents if all cell values are valid according to the data validation rules.
         * Returns true if all cell values are valid, or false if all cell values are invalid.
         * Returns null if there are both valid and invalid cell values within the range.
         */
        readonly valid: boolean;

        /**
         * Clears the data validation from the current range.
         */
        clear(): void;

        /**
         * Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.
         */
        getInvalidCells(): RangeAreas;

        /**
         * Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.
         */
        getInvalidCellsOrNullObject(): RangeAreas;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the results from the removeDuplicates method on range
     */
    interface RemoveDuplicatesResult {
        /**
         * Number of duplicated rows removed by the operation.
         */
        readonly removed: number;

        /**
         * Number of remaining unique rows present in the resulting range.
         */
        readonly uniqueRemaining: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     */
    interface RangeFormat {
        /**
         * Specifies if text is automatically indented when text alignment is set to equal distribution.
         */
        autoIndent: boolean;

        /**
         * Collection of border objects that apply to the overall range.
         */
        readonly borders: RangeBorderCollection;

        /**
         * Specifies the width of all colums within the range. If the column widths are not uniform, null will be returned.
         */
        columnWidth: number;

        /**
         * Returns the fill object defined on the overall range.
         */
        readonly fill: RangeFill;

        /**
         * Returns the font object defined on the overall range.
         */
        readonly font: RangeFont;

        /**
         * Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
         */
        horizontalAlignment:
            | HorizontalAlignment
            | "General"
            | "Left"
            | "Center"
            | "Right"
            | "Fill"
            | "Justify"
            | "CenterAcrossSelection"
            | "Distributed";

        /**
         * An integer from 0 to 250 that indicates the indent level.
         */
        indentLevel: number;

        /**
         * Returns the format protection object for a range.
         */
        readonly protection: FormatProtection;

        /**
         * The reading order for the range.
         */
        readingOrder: ReadingOrder | "Context" | "LeftToRight" | "RightToLeft";

        /**
         * The height of all rows in the range. If the row heights are not uniform, null will be returned.
         */
        rowHeight: number;

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        shrinkToFit: boolean;

        /**
         * The text orientation of all the cells within the range.
         * The text orientation should be an integer either from -90 to 90, or 180 for vertically-oriented text.
         * If the orientation within a range are not uniform, then null will be returned.
         */
        textOrientation: number;

        /**
         * Determines if the row height of the Range object equals the standard height of the sheet.
         * Returns True if the row height of the Range object equals the standard height of the sheet.
         * Returns Null if the range contains more than one row and the rows aren't all the same height.
         * Returns False otherwise.
         */
        useStandardHeight: boolean;

        /**
         * Specifies if the column width of the Range object equals the standard width of the sheet.
         * Returns True if the column width of the Range object equals the standard width of the sheet.
         * Returns Null if the range contains more than one column and the columns aren't all the same height.
         * Returns False otherwise.
         */
        useStandardWidth: boolean;

        /**
         * Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
         */
        verticalAlignment:
            | VerticalAlignment
            | "Top"
            | "Center"
            | "Bottom"
            | "Justify"
            | "Distributed";

        /**
         * Specifies if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
         */
        wrapText: boolean;

        /**
         * Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250 and is measured in characters.
         * @param amount The number of character spaces by which the current indent is adjusted. This value should be between -250 and 250.
         * **Note**: If the amount would raise the indent level above 250, the indent level stays with 250.
         * Similarly, if the amount would lower the indent level below 0, the indent level stays 0.
         */
        adjustIndent(amount: number): void;

        /**
         * Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitColumns(): void;

        /**
         * Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitRows(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the format protection of a range object.
     */
    interface FormatProtection {
        /**
         * Specifies if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
         */
        formulaHidden: boolean;

        /**
         * Specifies if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
         */
        locked: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the background of a range object.
     */
    interface RangeFill {
        /**
         * HTML color code representing the color of the background, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")
         */
        color: string;

        /**
         * The pattern of a range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.
         * A null value indicates that the entire range doesn't have uniform pattern setting.
         */
        pattern:
            | FillPattern
            | "None"
            | "Solid"
            | "Gray50"
            | "Gray75"
            | "Gray25"
            | "Horizontal"
            | "Vertical"
            | "Down"
            | "Up"
            | "Checker"
            | "SemiGray75"
            | "LightHorizontal"
            | "LightVertical"
            | "LightDown"
            | "LightUp"
            | "Grid"
            | "CrissCross"
            | "Gray16"
            | "Gray8"
            | "LinearGradient"
            | "RectangularGradient";

        /**
         * The HTML color code representing the color of the range pattern, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        patternColor: string;

        /**
         * Specifies a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * If the pattern tintAndShades are not uniform, null will be returned.
         */
        patternTintAndShade: number;

        /**
         * Specifies a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * If the tintAndShades are not uniform, null will be returned.
         */
        tintAndShade: number;

        /**
         * Resets the range background.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the border of an object.
     */
    interface RangeBorder {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        color: string;

        /**
         * Constant value that indicates the specific side of the border. See Excel.BorderIndex for details.
         */
        readonly sideIndex:
            | BorderIndex
            | "EdgeTop"
            | "EdgeBottom"
            | "EdgeLeft"
            | "EdgeRight"
            | "InsideVertical"
            | "InsideHorizontal"
            | "DiagonalDown"
            | "DiagonalUp";

        /**
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         */
        style:
            | BorderLineStyle
            | "None"
            | "Continuous"
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "Dot"
            | "Double"
            | "SlantDashDot";

        /**
         * Specifies a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A null value indicates that the border doesn't have uniform tintAndShade setting.
         */
        tintAndShade: number;

        /**
         * Specifies the weight of the border around a range. See Excel.BorderWeight for details.
         */
        weight: BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the border objects that make up the range border.
     */
    interface RangeBorderCollection {
        /**
         * Number of border objects in the collection.
         */
        readonly count: number;

        /**
         * Specifies a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A null value indicates that the entire border collections don't have uniform tintAndShade setting.
         */
        tintAndShade: number;

        /**
         * Gets a border object using its name.
         * @param index Index value of the border object to be retrieved. See Excel.BorderIndex for details.
         */
        getItem(
            index:
                | BorderIndex
                | "EdgeTop"
                | "EdgeBottom"
                | "EdgeLeft"
                | "EdgeRight"
                | "InsideVertical"
                | "InsideHorizontal"
                | "DiagonalDown"
                | "DiagonalUp"
        ): RangeBorder;

        /**
         * Gets a border object using its index.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): RangeBorder;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     */
    interface RangeFont {
        /**
         * Represents the bold status of font.
         */
        bold: boolean;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        color: string;

        /**
         * Specifies the italic status of the font.
         */
        italic: boolean;

        /**
         * Font name (e.g., "Calibri")
         */
        name: string;

        /**
         * Font size.
         */
        size: number;

        /**
         * Specifies the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.
         */
        strikethrough: boolean;

        /**
         * Specifies the Subscript status of font.
         * Returns True if all the fonts of the range are Subscript.
         * Returns False if all the fonts of the range are Superscript or normal (neither Superscript, nor Subscript).
         * Returns Null otherwise.
         */
        subscript: boolean;

        /**
         * Specifies the Superscript status of font.
         * Returns True if all the fonts of the range are Superscript.
         * Returns False if all the fonts of the range are Subscript or normal (neither Superscript, nor Subscript).
         * Returns Null otherwise.
         */
        superscript: boolean;

        /**
         * Specifies a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A null value indicates that the entire range doesn't have uniform font tintAndShade setting.
         */
        tintAndShade: number;

        /**
         * Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
         */
        underline:
            | RangeUnderlineStyle
            | "None"
            | "Single"
            | "Double"
            | "SingleAccountant"
            | "DoubleAccountant";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A collection of all the chart objects on a worksheet.
     */
    interface ChartCollection {
        /**
         * Returns the number of charts in the worksheet.
         */
        readonly count: number;

        /**
         * Creates a new chart.
         * @param type Represents the type of a chart. See Excel.ChartType for details.
         * @param sourceData The Range object corresponding to the source data.
         * @param seriesBy Optional. Specifies the way columns or rows are used as data series on the chart. See Excel.ChartSeriesBy for details.
         */
        add(
            type:
                | ChartType
                | "Invalid"
                | "ColumnClustered"
                | "ColumnStacked"
                | "ColumnStacked100"
                | "3DColumnClustered"
                | "3DColumnStacked"
                | "3DColumnStacked100"
                | "BarClustered"
                | "BarStacked"
                | "BarStacked100"
                | "3DBarClustered"
                | "3DBarStacked"
                | "3DBarStacked100"
                | "LineStacked"
                | "LineStacked100"
                | "LineMarkers"
                | "LineMarkersStacked"
                | "LineMarkersStacked100"
                | "PieOfPie"
                | "PieExploded"
                | "3DPieExploded"
                | "BarOfPie"
                | "XyscatterSmooth"
                | "XyscatterSmoothNoMarkers"
                | "XyscatterLines"
                | "XyscatterLinesNoMarkers"
                | "AreaStacked"
                | "AreaStacked100"
                | "3DAreaStacked"
                | "3DAreaStacked100"
                | "DoughnutExploded"
                | "RadarMarkers"
                | "RadarFilled"
                | "Surface"
                | "SurfaceWireframe"
                | "SurfaceTopView"
                | "SurfaceTopViewWireframe"
                | "Bubble"
                | "Bubble3DEffect"
                | "StockHLC"
                | "StockOHLC"
                | "StockVHLC"
                | "StockVOHLC"
                | "CylinderColClustered"
                | "CylinderColStacked"
                | "CylinderColStacked100"
                | "CylinderBarClustered"
                | "CylinderBarStacked"
                | "CylinderBarStacked100"
                | "CylinderCol"
                | "ConeColClustered"
                | "ConeColStacked"
                | "ConeColStacked100"
                | "ConeBarClustered"
                | "ConeBarStacked"
                | "ConeBarStacked100"
                | "ConeCol"
                | "PyramidColClustered"
                | "PyramidColStacked"
                | "PyramidColStacked100"
                | "PyramidBarClustered"
                | "PyramidBarStacked"
                | "PyramidBarStacked100"
                | "PyramidCol"
                | "3DColumn"
                | "Line"
                | "3DLine"
                | "3DPie"
                | "Pie"
                | "Xyscatter"
                | "3DArea"
                | "Area"
                | "Doughnut"
                | "Radar"
                | "Histogram"
                | "Boxwhisker"
                | "Pareto"
                | "RegionMap"
                | "Treemap"
                | "Waterfall"
                | "Sunburst"
                | "Funnel",
            sourceData: Range,
            seriesBy?: ChartSeriesBy | "Auto" | "Columns" | "Rows"
        ): Chart;

        /**
         * Returns the number of charts in the worksheet.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
         * @param name Name of the chart to be retrieved.
         */
        getItem(name: string): Chart;

        /**
         * Gets a chart based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Chart;

        /**
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
         * If the chart does not exist, will return a null object.
         * @param name Name of the chart to be retrieved.
         */
        getItemOrNullObject(name: string): Chart;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a chart object in a workbook.
     * To learn more about the Chart object model, see {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts | Work with charts using the Excel JavaScript API}.
     */
    interface Chart {
        /**
         * Represents chart axes.
         */
        readonly axes: ChartAxes;

        /**
         * Specifies a ChartCategoryLabelLevel enumeration constant referring to
         * the level of where the category labels are being sourced from.
         */
        categoryLabelLevel: number;

        /**
         * Specifies the type of the chart. See Excel.ChartType for details.
         */
        chartType:
            | ChartType
            | "Invalid"
            | "ColumnClustered"
            | "ColumnStacked"
            | "ColumnStacked100"
            | "3DColumnClustered"
            | "3DColumnStacked"
            | "3DColumnStacked100"
            | "BarClustered"
            | "BarStacked"
            | "BarStacked100"
            | "3DBarClustered"
            | "3DBarStacked"
            | "3DBarStacked100"
            | "LineStacked"
            | "LineStacked100"
            | "LineMarkers"
            | "LineMarkersStacked"
            | "LineMarkersStacked100"
            | "PieOfPie"
            | "PieExploded"
            | "3DPieExploded"
            | "BarOfPie"
            | "XyscatterSmooth"
            | "XyscatterSmoothNoMarkers"
            | "XyscatterLines"
            | "XyscatterLinesNoMarkers"
            | "AreaStacked"
            | "AreaStacked100"
            | "3DAreaStacked"
            | "3DAreaStacked100"
            | "DoughnutExploded"
            | "RadarMarkers"
            | "RadarFilled"
            | "Surface"
            | "SurfaceWireframe"
            | "SurfaceTopView"
            | "SurfaceTopViewWireframe"
            | "Bubble"
            | "Bubble3DEffect"
            | "StockHLC"
            | "StockOHLC"
            | "StockVHLC"
            | "StockVOHLC"
            | "CylinderColClustered"
            | "CylinderColStacked"
            | "CylinderColStacked100"
            | "CylinderBarClustered"
            | "CylinderBarStacked"
            | "CylinderBarStacked100"
            | "CylinderCol"
            | "ConeColClustered"
            | "ConeColStacked"
            | "ConeColStacked100"
            | "ConeBarClustered"
            | "ConeBarStacked"
            | "ConeBarStacked100"
            | "ConeCol"
            | "PyramidColClustered"
            | "PyramidColStacked"
            | "PyramidColStacked100"
            | "PyramidBarClustered"
            | "PyramidBarStacked"
            | "PyramidBarStacked100"
            | "PyramidCol"
            | "3DColumn"
            | "Line"
            | "3DLine"
            | "3DPie"
            | "Pie"
            | "Xyscatter"
            | "3DArea"
            | "Area"
            | "Doughnut"
            | "Radar"
            | "Histogram"
            | "Boxwhisker"
            | "Pareto"
            | "RegionMap"
            | "Treemap"
            | "Waterfall"
            | "Sunburst"
            | "Funnel";

        /**
         * Represents the datalabels on the chart.
         */
        readonly dataLabels: ChartDataLabels;

        /**
         * Specifies the way that blank cells are plotted on a chart.
         */
        displayBlanksAs:
            | ChartDisplayBlanksAs
            | "NotPlotted"
            | "Zero"
            | "Interplotted";

        /**
         * Encapsulates the format properties for the chart area.
         */
        readonly format: ChartAreaFormat;

        /**
         * Specifies the height, in points, of the chart object.
         */
        height: number;

        /**
         * The unique id of chart.
         */
        readonly id: string;

        /**
         * The distance, in points, from the left side of the chart to the worksheet origin.
         */
        left: number;

        /**
         * Represents the legend for the chart.
         */
        readonly legend: ChartLegend;

        /**
         * Specifies the name of a chart object.
         */
        name: string;

        /**
         * Encapsulates the options for a pivot chart.
         */
        readonly pivotOptions: ChartPivotOptions;

        /**
         * Represents the plotArea for the chart.
         */
        readonly plotArea: ChartPlotArea;

        /**
         * Specifies the way columns or rows are used as data series on the chart.
         */
        plotBy: ChartPlotBy | "Rows" | "Columns";

        /**
         * True if only visible cells are plotted. False if both visible and hidden cells are plotted.
         */
        plotVisibleOnly: boolean;

        /**
         * Represents either a single series or collection of series in the chart.
         */
        readonly series: ChartSeriesCollection;

        /**
         * Specifies a ChartSeriesNameLevel enumeration constant referring to
         * the level of where the series names are being sourced from.
         */
        seriesNameLevel: number;

        /**
         * Specifies whether to display all field buttons on a PivotChart.
         */
        showAllFieldButtons: boolean;

        /**
         * Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.
         * If value axis became smaller than the size of data points, you can use this property to set whether to show the data labels.
         * This property applies to 2-D charts only.
         */
        showDataLabelsOverMaximum: boolean;

        /**
         * Specifies the chart style for the chart.
         */
        style: number;

        /**
         * Specifies the title of the specified chart, including the text, visibility, position, and formatting of the title.
         */
        readonly title: ChartTitle;

        /**
         * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         */
        top: number;

        /**
         * Specifies the width, in points, of the chart object.
         */
        width: number;

        /**
         * The worksheet containing the current chart.
         */
        readonly worksheet: Worksheet;

        /**
         * Activates the chart in the Excel UI.
         */
        activate(): void;

        /**
         * Deletes the chart object.
         */
        delete(): void;

        /**
         * Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
         * The aspect ratio is preserved as part of the resizing.
         * @param height (Optional) The desired height of the resulting image.
         * @param width (Optional) The desired width of the resulting image.
         * @param fittingMode (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set).
         */
        getImage(
            width?: number,
            height?: number,
            fittingMode?: ImageFittingMode | "Fit" | "FitAndCenter" | "Fill"
        ): ClientResult<string>;

        /**
         * Resets the source data for the chart.
         * @param sourceData The range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, and Columns. See Excel.ChartSeriesBy for details.
         */
        setData(
            sourceData: Range,
            seriesBy?: ChartSeriesBy | "Auto" | "Columns" | "Rows"
        ): void;

        /**
         * Positions the chart relative to cells on the worksheet.
         * @param startCell The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
         * @param endCell (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
         */
        setPosition(startCell: Range | string, endCell?: Range | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the options for the pivot chart.
     */
    interface ChartPivotOptions {
        /**
         * Specifies whether to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.
         */
        showAxisFieldButtons: boolean;

        /**
         * Specifies whether to display the legend field buttons on a PivotChart.
         */
        showLegendFieldButtons: boolean;

        /**
         * Specifies whether to display the report filter field buttons on a PivotChart.
         */
        showReportFilterFieldButtons: boolean;

        /**
         * Specifies whether to display the show value field buttons on a PivotChart.
         */
        showValueFieldButtons: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for the overall chart area.
     */
    interface ChartAreaFormat {
        /**
         * Represents the border format of chart area, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Specifies the color scheme of the chart.
         */
        colorScheme:
            | ChartColorScheme
            | "ColorfulPalette1"
            | "ColorfulPalette2"
            | "ColorfulPalette3"
            | "ColorfulPalette4"
            | "MonochromaticPalette1"
            | "MonochromaticPalette2"
            | "MonochromaticPalette3"
            | "MonochromaticPalette4"
            | "MonochromaticPalette5"
            | "MonochromaticPalette6"
            | "MonochromaticPalette7"
            | "MonochromaticPalette8"
            | "MonochromaticPalette9"
            | "MonochromaticPalette10"
            | "MonochromaticPalette11"
            | "MonochromaticPalette12"
            | "MonochromaticPalette13";

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for the current object.
         */
        readonly font: ChartFont;

        /**
         * Specifies if the chart area of the chart has rounded corners.
         */
        roundedCorners: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of chart series.
     */
    interface ChartSeriesCollection {
        /**
         * Returns the number of series in the collection.
         */
        readonly count: number;

        /**
         * Add a new series to the collection. The new added series is not visible until set values/x axis values/bubble sizes for it (depending on chart type).
         * @param name Optional. Name of the series.
         * @param index Optional. Index value of the series to be added. Zero-indexed.
         */
        add(name?: string, index?: number): ChartSeries;

        /**
         * Returns the number of series in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Retrieves a series based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): ChartSeries;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a series in a chart.
     */
    interface ChartSeries {
        /**
         * Specifies the group for the specified series.
         */
        axisGroup: ChartAxisGroup | "Primary" | "Secondary";

        /**
         * Encapsulates the bin options for histogram charts and pareto charts.
         */
        readonly binOptions: ChartBinOptions;

        /**
         * Encapsulates the options for the box and whisker charts.
         */
        readonly boxwhiskerOptions: ChartBoxwhiskerOptions;

        /**
         * This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts.
         */
        bubbleScale: number;

        /**
         * Represents the chart type of a series. See Excel.ChartType for details.
         */
        chartType:
            | ChartType
            | "Invalid"
            | "ColumnClustered"
            | "ColumnStacked"
            | "ColumnStacked100"
            | "3DColumnClustered"
            | "3DColumnStacked"
            | "3DColumnStacked100"
            | "BarClustered"
            | "BarStacked"
            | "BarStacked100"
            | "3DBarClustered"
            | "3DBarStacked"
            | "3DBarStacked100"
            | "LineStacked"
            | "LineStacked100"
            | "LineMarkers"
            | "LineMarkersStacked"
            | "LineMarkersStacked100"
            | "PieOfPie"
            | "PieExploded"
            | "3DPieExploded"
            | "BarOfPie"
            | "XyscatterSmooth"
            | "XyscatterSmoothNoMarkers"
            | "XyscatterLines"
            | "XyscatterLinesNoMarkers"
            | "AreaStacked"
            | "AreaStacked100"
            | "3DAreaStacked"
            | "3DAreaStacked100"
            | "DoughnutExploded"
            | "RadarMarkers"
            | "RadarFilled"
            | "Surface"
            | "SurfaceWireframe"
            | "SurfaceTopView"
            | "SurfaceTopViewWireframe"
            | "Bubble"
            | "Bubble3DEffect"
            | "StockHLC"
            | "StockOHLC"
            | "StockVHLC"
            | "StockVOHLC"
            | "CylinderColClustered"
            | "CylinderColStacked"
            | "CylinderColStacked100"
            | "CylinderBarClustered"
            | "CylinderBarStacked"
            | "CylinderBarStacked100"
            | "CylinderCol"
            | "ConeColClustered"
            | "ConeColStacked"
            | "ConeColStacked100"
            | "ConeBarClustered"
            | "ConeBarStacked"
            | "ConeBarStacked100"
            | "ConeCol"
            | "PyramidColClustered"
            | "PyramidColStacked"
            | "PyramidColStacked100"
            | "PyramidBarClustered"
            | "PyramidBarStacked"
            | "PyramidBarStacked100"
            | "PyramidCol"
            | "3DColumn"
            | "Line"
            | "3DLine"
            | "3DPie"
            | "Pie"
            | "Xyscatter"
            | "3DArea"
            | "Area"
            | "Doughnut"
            | "Radar"
            | "Histogram"
            | "Boxwhisker"
            | "Pareto"
            | "RegionMap"
            | "Treemap"
            | "Waterfall"
            | "Sunburst"
            | "Funnel";

        /**
         * Represents a collection of all dataLabels in the series.
         */
        readonly dataLabels: ChartDataLabels;

        /**
         * Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.
         * Throws an invalid argument exception on invalid charts.
         */
        doughnutHoleSize: number;

        /**
         * Specifies the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).
         */
        explosion: number;

        /**
         * Specifies if the series is filtered. Not applicable for surface charts.
         */
        filtered: boolean;

        /**
         * Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.
         */
        firstSliceAngle: number;

        /**
         * Represents the formatting of a chart series, which includes fill and line formatting.
         */
        readonly format: ChartSeriesFormat;

        /**
         * Represents the gap width of a chart series.  Only valid on bar and column charts, as well as
         * specific classes of line and pie charts.  Throws an invalid argument exception on invalid charts.
         */
        gapWidth: number;

        /**
         * Specifies the color for maximum value of a region map chart series.
         */
        gradientMaximumColor: string;

        /**
         * Specifies the type for maximum value of a region map chart series.
         */
        gradientMaximumType:
            | ChartGradientStyleType
            | "ExtremeValue"
            | "Number"
            | "Percent";

        /**
         * Specifies the maximum value of a region map chart series.
         */
        gradientMaximumValue: number;

        /**
         * Specifies the color for midpoint value of a region map chart series.
         */
        gradientMidpointColor: string;

        /**
         * Specifies the type for midpoint value of a region map chart series.
         */
        gradientMidpointType:
            | ChartGradientStyleType
            | "ExtremeValue"
            | "Number"
            | "Percent";

        /**
         * Specifies the midpoint value of a region map chart series.
         */
        gradientMidpointValue: number;

        /**
         * Specifies the color for minimum value of a region map chart series.
         */
        gradientMinimumColor: string;

        /**
         * Specifies the type for minimum value of a region map chart series.
         */
        gradientMinimumType:
            | ChartGradientStyleType
            | "ExtremeValue"
            | "Number"
            | "Percent";

        /**
         * Specifies the minimum value of a region map chart series.
         */
        gradientMinimumValue: number;

        /**
         * Specifies the series gradient style of a region map chart.
         */
        gradientStyle: ChartGradientStyle | "TwoPhaseColor" | "ThreePhaseColor";

        /**
         * Specifies if the series has data labels.
         */
        hasDataLabels: boolean;

        /**
         * Specifies the fill color for negative data points in a series.
         */
        invertColor: string;

        /**
         * True if Excel inverts the pattern in the item when it corresponds to a negative number.
         */
        invertIfNegative: boolean;

        /**
         * Encapsulates the options for a region map chart.
         */
        readonly mapOptions: ChartMapOptions;

        /**
         * Specifies the markers background color of a chart series.
         */
        markerBackgroundColor: string;

        /**
         * Specifies the markers foreground color of a chart series.
         */
        markerForegroundColor: string;

        /**
         * Specifies the marker size of a chart series.
         */
        markerSize: number;

        /**
         * Specifies the marker style of a chart series. See Excel.ChartMarkerStyle for details.
         */
        markerStyle:
            | ChartMarkerStyle
            | "Invalid"
            | "Automatic"
            | "None"
            | "Square"
            | "Diamond"
            | "Triangle"
            | "X"
            | "Star"
            | "Dot"
            | "Dash"
            | "Circle"
            | "Plus"
            | "Picture";

        /**
         * Specifies the name of a series in a chart.
         */
        name: string;

        /**
         * Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts.
         */
        overlap: number;

        /**
         * Specifies the series parent label strategy area for a treemap chart.
         */
        parentLabelStrategy:
            | ChartParentLabelStrategy
            | "None"
            | "Banner"
            | "Overlapping";

        /**
         * Specifies the plot order of a chart series within the chart group.
         */
        plotOrder: number;

        /**
         * Returns a collection of all points in the series.
         */
        readonly points: ChartPointsCollection;

        /**
         * Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.
         */
        secondPlotSize: number;

        /**
         * Specifies whether connector lines are shown in waterfall charts.
         */
        showConnectorLines: boolean;

        /**
         * Specifies whether leader lines are displayed for each data label in the series.
         */
        showLeaderLines: boolean;

        /**
         * Specifies if the series has a shadow.
         */
        showShadow: boolean;

        /**
         * Specifies if the series is smooth. Only applicable to line and scatter charts.
         */
        smooth: boolean;

        /**
         * Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.
         */
        splitType:
            | ChartSplitType
            | "SplitByPosition"
            | "SplitByValue"
            | "SplitByPercentValue"
            | "SplitByCustomSplit";

        /**
         * Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.
         */
        splitValue: number;

        /**
         * The collection of trendlines in the series.
         */
        readonly trendlines: ChartTrendlineCollection;

        /**
         * True if Excel assigns a different color or pattern to each data marker. The chart must contain only one series.
         */
        varyByCategories: boolean;

        /**
         * Represents the error bar object of a chart series.
         */
        readonly xErrorBars: ChartErrorBars;

        /**
         * Represents the error bar object of a chart series.
         */
        readonly yErrorBars: ChartErrorBars;

        /**
         * Deletes the chart series.
         */
        delete(): void;

        /**
         * Sets the bubble sizes for a chart series. Only works for bubble charts.
         * @param sourceData The Range object corresponding to the source data.
         */
        setBubbleSizes(sourceData: Range): void;

        /**
         * Sets the values for a chart series. For scatter chart, it means Y axis values.
         * @param sourceData The Range object corresponding to the source data.
         */
        setValues(sourceData: Range): void;

        /**
         * Sets the values of the X axis for a chart series. Only works for scatter charts.
         * @param sourceData The Range object corresponding to the source data.
         */
        setXAxisValues(sourceData: Range): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for the chart series
     */
    interface ChartSeriesFormat {
        /**
         * Represents the fill format of a chart series, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Represents line formatting.
         */
        readonly line: ChartLineFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A collection of all the chart points within a series inside a chart.
     */
    interface ChartPointsCollection {
        /**
         * Returns the number of chart points in the series.
         */
        readonly count: number;

        /**
         * Returns the number of chart points in the series.
         */
        getCount(): ClientResult<number>;

        /**
         * Retrieve a point based on its position within the series.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): ChartPoint;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a point of a series in a chart.
     */
    interface ChartPoint {
        /**
         * Returns the data label of a chart point.
         */
        readonly dataLabel: ChartDataLabel;

        /**
         * Encapsulates the format properties chart point.
         */
        readonly format: ChartPointFormat;

        /**
         * Represents whether a data point has a data label. Not applicable for surface charts.
         */
        hasDataLabel: boolean;

        /**
         * HTML color code representation of the marker background color of data point (e.g., #FF0000 represents Red).
         */
        markerBackgroundColor: string;

        /**
         * HTML color code representation of the marker foreground color of data point (e.g., #FF0000 represents Red).
         */
        markerForegroundColor: string;

        /**
         * Represents marker size of data point.
         */
        markerSize: number;

        /**
         * Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.
         */
        markerStyle:
            | ChartMarkerStyle
            | "Invalid"
            | "Automatic"
            | "None"
            | "Square"
            | "Diamond"
            | "Triangle"
            | "X"
            | "Star"
            | "Dot"
            | "Dash"
            | "Circle"
            | "Plus"
            | "Picture";

        /**
         * Returns the value of a chart point.
         */
        readonly value: any;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents formatting object for chart points.
     */
    interface ChartPointFormat {
        /**
         * Represents the border format of a chart data point, which includes color, style, and weight information.
         */
        readonly border: ChartBorder;

        /**
         * Represents the fill format of a chart, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the chart axes.
     */
    interface ChartAxes {
        /**
         * Represents the category axis in a chart.
         */
        readonly categoryAxis: ChartAxis;

        /**
         * Represents the series axis of a 3-dimensional chart.
         */
        readonly seriesAxis: ChartAxis;

        /**
         * Represents the value axis in an axis.
         */
        readonly valueAxis: ChartAxis;

        /**
         * Returns the specific axis identified by type and group.
         * @param type Specifies the axis type. See Excel.ChartAxisType for details.
         * @param group Optional. Specifies the axis group. See Excel.ChartAxisGroup for details.
         */
        getItem(
            type: ChartAxisType | "Invalid" | "Category" | "Value" | "Series",
            group?: ChartAxisGroup | "Primary" | "Secondary"
        ): ChartAxis;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a single axis in a chart.
     */
    interface ChartAxis {
        /**
         * Specifies the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.
         */
        alignment: ChartTickLabelAlignment | "Center" | "Left" | "Right";

        /**
         * Specifies the group for the specified axis. See Excel.ChartAxisGroup for details.
         */
        readonly axisGroup: ChartAxisGroup | "Primary" | "Secondary";

        /**
         * Specifies the base unit for the specified category axis.
         */
        baseTimeUnit: ChartAxisTimeUnit | "Days" | "Months" | "Years";

        /**
         * Specifies the category axis type.
         */
        categoryType:
            | ChartAxisCategoryType
            | "Automatic"
            | "TextAxis"
            | "DateAxis";

        /**
         * Specifies the custom axis display unit value. To set this property, please use the SetCustomDisplayUnit(double) method.
         */
        readonly customDisplayUnit: number;

        /**
         * Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.
         */
        displayUnit:
            | ChartAxisDisplayUnit
            | "None"
            | "Hundreds"
            | "Thousands"
            | "TenThousands"
            | "HundredThousands"
            | "Millions"
            | "TenMillions"
            | "HundredMillions"
            | "Billions"
            | "Trillions"
            | "Custom";

        /**
         * Represents the formatting of a chart object, which includes line and font formatting.
         */
        readonly format: ChartAxisFormat;

        /**
         * Specifies the height, in points, of the chart axis. Null if the axis is not visible.
         */
        readonly height: number;

        /**
         * Specifies if the value axis crosses the category axis between categories.
         */
        isBetweenCategories: boolean;

        /**
         * Specifies the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible.
         */
        readonly left: number;

        /**
         * Specifies if the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.
         */
        linkNumberFormat: boolean;

        /**
         * Specifies the base of the logarithm when using logarithmic scales.
         */
        logBase: number;

        /**
         * Returns a Gridlines object that represents the major gridlines for the specified axis.
         */
        readonly majorGridlines: ChartGridlines;

        /**
         * Specifies the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.
         */
        majorTickMark:
            | ChartAxisTickMark
            | "None"
            | "Cross"
            | "Inside"
            | "Outside";

        /**
         * Specifies the major unit scale value for the category axis when the CategoryType property is set to TimeScale.
         */
        majorTimeUnitScale: ChartAxisTimeUnit | "Days" | "Months" | "Years";

        /**
         * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
         */
        majorUnit: any;

        /**
         * Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         */
        maximum: any;

        /**
         * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         */
        minimum: any;

        /**
         * Returns a Gridlines object that represents the minor gridlines for the specified axis.
         */
        readonly minorGridlines: ChartGridlines;

        /**
         * Specifies the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.
         */
        minorTickMark:
            | ChartAxisTickMark
            | "None"
            | "Cross"
            | "Inside"
            | "Outside";

        /**
         * Specifies the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.
         */
        minorTimeUnitScale: ChartAxisTimeUnit | "Days" | "Months" | "Years";

        /**
         * Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         */
        minorUnit: any;

        /**
         * Specifies if an axis is multilevel.
         */
        multiLevel: boolean;

        /**
         * Specifies the format code for the axis tick label.
         */
        numberFormat: string;

        /**
         * Specifies the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.
         */
        offset: number;

        /**
         * Specifies the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.
         */
        position:
            | ChartAxisPosition
            | "Automatic"
            | "Maximum"
            | "Minimum"
            | "Custom";

        /**
         * Specifies the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property.
         */
        readonly positionAt: number;

        /**
         * Specifies if Excel plots data points from last to first.
         */
        reversePlotOrder: boolean;

        /**
         * Specifies the value axis scale type. See Excel.ChartAxisScaleType for details.
         */
        scaleType: ChartAxisScaleType | "Linear" | "Logarithmic";

        /**
         * Specifies if the axis display unit label is visible.
         */
        showDisplayUnitLabel: boolean;

        /**
         * Specifies the angle to which the text is oriented for the chart axis tick label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: any;

        /**
         * Specifies the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.
         */
        tickLabelPosition:
            | ChartAxisTickLabelPosition
            | "NextToAxis"
            | "High"
            | "Low"
            | "None";

        /**
         * Specifies the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.
         */
        tickLabelSpacing: any;

        /**
         * Specifies the number of categories or series between tick marks.
         */
        tickMarkSpacing: number;

        /**
         * Represents the axis title.
         */
        readonly title: ChartAxisTitle;

        /**
         * Specifies the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible.
         */
        readonly top: number;

        /**
         * Specifies the axis type. See Excel.ChartAxisType for details.
         */
        readonly type:
            | ChartAxisType
            | "Invalid"
            | "Category"
            | "Value"
            | "Series";

        /**
         * Specifies if the axis is visible.
         */
        visible: boolean;

        /**
         * Specifies the width, in points, of the chart axis. Null if the axis is not visible.
         */
        readonly width: number;

        /**
         * Sets all the category names for the specified axis.
         * @param sourceData The Range object corresponding to the source data.
         */
        setCategoryNames(sourceData: Range): void;

        /**
         * Sets the axis display unit to a custom value.
         * @param value Custom value of the display unit.
         */
        setCustomDisplayUnit(value: number): void;

        /**
         * Sets the specified axis position where the other axis crosses at.
         * @param value Custom value of the crosses at
         */
        setPositionAt(value: number): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for the chart axis.
     */
    interface ChartAxisFormat {
        /**
         * Specifies chart fill formatting.
         */
        readonly fill: ChartFill;

        /**
         * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
         */
        readonly font: ChartFont;

        /**
         * Specifies chart line formatting.
         */
        readonly line: ChartLineFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the title of a chart axis.
     */
    interface ChartAxisTitle {
        /**
         * Specifies the formatting of chart axis title.
         */
        readonly format: ChartAxisTitleFormat;

        /**
         * Specifies the axis title.
         */
        text: string;

        /**
         * Specifies the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: number;

        /**
         * Specifies if the axis title is visibile.
         */
        visible: boolean;

        /**
         * A string value that represents the formula of chart axis title using A1-style notation.
         * @param formula a string that present the formula to set
         */
        setFormula(formula: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the chart axis title formatting.
     */
    interface ChartAxisTitleFormat {
        /**
         * Specifies the chart axis title's border format, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Specifies the chart axis title's fill formatting.
         */
        readonly fill: ChartFill;

        /**
         * Specifies the chart axis title's font attributes, such as font name, font size, color, etc. of chart axis title object.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the data labels on a chart point.
     */
    interface ChartDataLabels {
        /**
         * Specifies if data labels automatically generate appropriate text based on context.
         */
        autoText: boolean;

        /**
         * Specifies the format of chart data labels, which includes fill and font formatting.
         */
        readonly format: ChartDataLabelFormat;

        /**
         * Specifies the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of data label is 0.
         */
        horizontalAlignment:
            | ChartTextHorizontalAlignment
            | "Center"
            | "Left"
            | "Right"
            | "Justify"
            | "Distributed";

        /**
         * Specifies if the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.
         */
        linkNumberFormat: boolean;

        /**
         * Specifies the format code for data labels.
         */
        numberFormat: string;

        /**
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         */
        position:
            | ChartDataLabelPosition
            | "Invalid"
            | "None"
            | "Center"
            | "InsideEnd"
            | "InsideBase"
            | "OutsideEnd"
            | "Left"
            | "Right"
            | "Top"
            | "Bottom"
            | "BestFit"
            | "Callout";

        /**
         * String representing the separator used for the data labels on a chart.
         */
        separator: string;

        /**
         * Specifies if the data label bubble size is visible.
         */
        showBubbleSize: boolean;

        /**
         * Specifies if the data label category name is visible.
         */
        showCategoryName: boolean;

        /**
         * Specifies if the data label legend key is visible.
         */
        showLegendKey: boolean;

        /**
         * Specifies if the data label percentage is visible.
         */
        showPercentage: boolean;

        /**
         * Specifies if the data label series name is visible.
         */
        showSeriesName: boolean;

        /**
         * Specifies if the data label value is visible.
         */
        showValue: boolean;

        /**
         * Represents the angle to which the text is oriented for data labels. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: number;

        /**
         * Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of data label is -90, 90, or 180.
         */
        verticalAlignment:
            | ChartTextVerticalAlignment
            | "Center"
            | "Bottom"
            | "Top"
            | "Justify"
            | "Distributed";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the data label of a chart point.
     */
    interface ChartDataLabel {
        /**
         * Specifies if the data label automatically generates appropriate text based on context.
         */
        autoText: boolean;

        /**
         * Represents the format of chart data label.
         */
        readonly format: ChartDataLabelFormat;

        /**
         * String value that represents the formula of chart data label using A1-style notation.
         */
        formula: string;

        /**
         * Returns the height, in points, of the chart data label. Null if chart data label is not visible.
         */
        readonly height: number;

        /**
         * Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of data label is -90, 90, or 180.
         */
        horizontalAlignment:
            | ChartTextHorizontalAlignment
            | "Center"
            | "Left"
            | "Right"
            | "Justify"
            | "Distributed";

        /**
         * Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.
         */
        left: number;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        linkNumberFormat: boolean;

        /**
         * String value that represents the format code for data label.
         */
        numberFormat: string;

        /**
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         */
        position:
            | ChartDataLabelPosition
            | "Invalid"
            | "None"
            | "Center"
            | "InsideEnd"
            | "InsideBase"
            | "OutsideEnd"
            | "Left"
            | "Right"
            | "Top"
            | "Bottom"
            | "BestFit"
            | "Callout";

        /**
         * String representing the separator used for the data label on a chart.
         */
        separator: string;

        /**
         * Specifies if the data label bubble size is visible.
         */
        showBubbleSize: boolean;

        /**
         * Specifies if the data label category name is visible.
         */
        showCategoryName: boolean;

        /**
         * Specifies if the data label legend key is visible.
         */
        showLegendKey: boolean;

        /**
         * Specifies if the data label percentage is visible.
         */
        showPercentage: boolean;

        /**
         * Specifies if the data label series name is visible.
         */
        showSeriesName: boolean;

        /**
         * Specifies if the data label value is visible.
         */
        showValue: boolean;

        /**
         * String representing the text of the data label on a chart.
         */
        text: string;

        /**
         * Represents the angle to which the text is oriented for the chart data label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: number;

        /**
         * Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.
         */
        top: number;

        /**
         * Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of data label is 0.
         */
        verticalAlignment:
            | ChartTextVerticalAlignment
            | "Center"
            | "Bottom"
            | "Top"
            | "Justify"
            | "Distributed";

        /**
         * Returns the width, in points, of the chart data label. Null if chart data label is not visible.
         */
        readonly width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for the chart data labels.
     */
    interface ChartDataLabelFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Represents the fill format of the current chart data label.
         */
        readonly fill: ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for a chart data label.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the attributes for a chart's error bars.
     */
    interface ChartErrorBars {
        /**
         * Specifies if error bars have an end style cap.
         */
        endStyleCap: boolean;

        /**
         * Specifies the formatting type of the error bars.
         */
        readonly format: ChartErrorBarsFormat;

        /**
         * Specifies which parts of the error bars to include.
         */
        include: ChartErrorBarsInclude | "Both" | "MinusValues" | "PlusValues";

        /**
         * The type of range marked by the error bars.
         */
        type:
            | ChartErrorBarsType
            | "FixedValue"
            | "Percent"
            | "StDev"
            | "StError"
            | "Custom";

        /**
         * Specifies whether the error bars are displayed.
         */
        visible: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for chart error bars.
     */
    interface ChartErrorBarsFormat {
        /**
         * Represents the chart line formatting.
         */
        readonly line: ChartLineFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents major or minor gridlines on a chart axis.
     */
    interface ChartGridlines {
        /**
         * Represents the formatting of chart gridlines.
         */
        readonly format: ChartGridlinesFormat;

        /**
         * Specifies if the axis gridlines are visible.
         */
        visible: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for chart gridlines.
     */
    interface ChartGridlinesFormat {
        /**
         * Represents chart line formatting.
         */
        readonly line: ChartLineFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the legend in a chart.
     */
    interface ChartLegend {
        /**
         * Represents the formatting of a chart legend, which includes fill and font formatting.
         */
        readonly format: ChartLegendFormat;

        /**
         * Specifies the height, in points, of the legend on the chart. Null if legend is not visible.
         */
        height: number;

        /**
         * Specifies the left, in points, of the legend on the chart. Null if legend is not visible.
         */
        left: number;

        /**
         * Represents a collection of legendEntries in the legend.
         */
        readonly legendEntries: ChartLegendEntryCollection;

        /**
         * Specifies if the chart legend should overlap with the main body of the chart.
         */
        overlay: boolean;

        /**
         * Specifies the position of the legend on the chart. See Excel.ChartLegendPosition for details.
         */
        position:
            | ChartLegendPosition
            | "Invalid"
            | "Top"
            | "Bottom"
            | "Left"
            | "Right"
            | "Corner"
            | "Custom";

        /**
         * Specifies if the legend has a shadow on the chart.
         */
        showShadow: boolean;

        /**
         * Specifies the top of a chart legend.
         */
        top: number;

        /**
         * Specifies if the ChartLegend is visible.
         */
        visible: boolean;

        /**
         * Specifies the width, in points, of the legend on the chart. Null if legend is not visible.
         */
        width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the legendEntry in legendEntryCollection.
     */
    interface ChartLegendEntry {
        /**
         * Specifies the height of the legendEntry on the chart legend.
         */
        readonly height: number;

        /**
         * Specifies the index of the legendEntry in the chart legend.
         */
        readonly index: number;

        /**
         * Specifies the left of a chart legendEntry.
         */
        readonly left: number;

        /**
         * Specifies the top of a chart legendEntry.
         */
        readonly top: number;

        /**
         * Represents the visible of a chart legend entry.
         */
        visible: boolean;

        /**
         * Represents the width of the legendEntry on the chart Legend.
         */
        readonly width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of legendEntries.
     */
    interface ChartLegendEntryCollection {
        /**
         * Returns the number of legendEntry in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Returns a legendEntry at the given index.
         * @param index Index of the legendEntry to be retrieved.
         */
        getItemAt(index: number): ChartLegendEntry;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties of a chart legend.
     */
    interface ChartLegendFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Represents the font attributes such as font name, font size, color, etc. of a chart legend.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the properties for a region map chart.
     */
    interface ChartMapOptions {
        /**
         * Specifies the series map labels strategy of a region map chart.
         */
        labelStrategy: ChartMapLabelStrategy | "None" | "BestFit" | "ShowAll";

        /**
         * Specifies the series mapping level of a region map chart.
         */
        level:
            | ChartMapAreaLevel
            | "Automatic"
            | "DataOnly"
            | "City"
            | "County"
            | "State"
            | "Country"
            | "Continent"
            | "World";

        /**
         * Specifies the series projection type of a region map chart.
         */
        projectionType:
            | ChartMapProjectionType
            | "Automatic"
            | "Mercator"
            | "Miller"
            | "Robinson"
            | "Albers";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a chart title object of a chart.
     */
    interface ChartTitle {
        /**
         * Represents the formatting of a chart title, which includes fill and font formatting.
         */
        readonly format: ChartTitleFormat;

        /**
         * Returns the height, in points, of the chart title. Null if chart title is not visible.
         */
        readonly height: number;

        /**
         * Specifies the horizontal alignment for chart title.
         */
        horizontalAlignment:
            | ChartTextHorizontalAlignment
            | "Center"
            | "Left"
            | "Right"
            | "Justify"
            | "Distributed";

        /**
         * Specifies the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.
         */
        left: number;

        /**
         * Specifies if the chart title will overlay the chart.
         */
        overlay: boolean;

        /**
         * Represents the position of chart title. See Excel.ChartTitlePosition for details.
         */
        position:
            | ChartTitlePosition
            | "Automatic"
            | "Top"
            | "Bottom"
            | "Left"
            | "Right";

        /**
         * Represents a boolean value that determines if the chart title has a shadow.
         */
        showShadow: boolean;

        /**
         * Specifies the chart's title text.
         */
        text: string;

        /**
         * Specifies the angle to which the text is oriented for the chart title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: number;

        /**
         * Specifies the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.
         */
        top: number;

        /**
         * Specifies the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.
         */
        verticalAlignment:
            | ChartTextVerticalAlignment
            | "Center"
            | "Bottom"
            | "Top"
            | "Justify"
            | "Distributed";

        /**
         * Specifies if the chart title is visibile.
         */
        visible: boolean;

        /**
         * Specifies the width, in points, of the chart title. Null if chart title is not visible.
         */
        readonly width: number;

        /**
         * Get the substring of a chart title. Line break '\n' also counts one character.
         * @param start Start position of substring to be retrieved. Position start with 0.
         * @param length Length of substring to be retrieved.
         */
        getSubstring(start: number, length: number): ChartFormatString;

        /**
         * Sets a string value that represents the formula of chart title using A1-style notation.
         * @param formula A string that represents the formula to set.
         */
        setFormula(formula: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the substring in chart related objects that contains text, like ChartTitle object, ChartAxisTitle object, etc.
     */
    interface ChartFormatString {
        /**
         * Represents the font attributes, such as font name, font size, color, etc. of chart characters object.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Provides access to the office art formatting for chart title.
     */
    interface ChartTitleFormat {
        /**
         * Represents the border format of chart title, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for an object.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the fill formatting for a chart element.
     */
    interface ChartFill {
        /**
         * Clear the fill color of a chart element.
         */
        clear(): void;

        /**
         * Sets the fill formatting of a chart element to a uniform color.
         * @param color HTML color code representing the color of the background, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the border formatting of a chart element.
     */
    interface ChartBorder {
        /**
         * HTML color code representing the color of borders in the chart.
         */
        color: string;

        /**
         * Represents the line style of the border. See Excel.ChartLineStyle for details.
         */
        lineStyle:
            | ChartLineStyle
            | "None"
            | "Continuous"
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "Dot"
            | "Grey25"
            | "Grey50"
            | "Grey75"
            | "Automatic"
            | "RoundDot";

        /**
         * Represents weight of the border, in points.
         */
        weight: number;

        /**
         * Clear the border format of a chart element.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the bin options for histogram charts and pareto charts.
     */
    interface ChartBinOptions {
        /**
         * Specifies if bin overflow is enabled in a histogram chart or pareto chart.
         */
        allowOverflow: boolean;

        /**
         * Specifies if bin underflow is enabled in a histogram chart or pareto chart.
         */
        allowUnderflow: boolean;

        /**
         * Specifies the bin count of a histogram chart or pareto chart.
         */
        count: number;

        /**
         * Specifies the bin overflow value of a histogram chart or pareto chart.
         */
        overflowValue: number;

        /**
         * Specifies the bin's type for a histogram chart or pareto chart.
         */
        type: ChartBinType | "Category" | "Auto" | "BinWidth" | "BinCount";

        /**
         * Specifies the bin underflow value of a histogram chart or pareto chart.
         */
        underflowValue: number;

        /**
         * Specifies the bin width value of a histogram chart or pareto chart.
         */
        width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the properties of a box and whisker chart.
     */
    interface ChartBoxwhiskerOptions {
        /**
         * Specifies if the quartile calculation type of a box and whisker chart.
         */
        quartileCalculation:
            | ChartBoxQuartileCalculation
            | "Inclusive"
            | "Exclusive";

        /**
         * Specifies if inner points are shown in a box and whisker chart.
         */
        showInnerPoints: boolean;

        /**
         * Specifies if the mean line is shown in a box and whisker chart.
         */
        showMeanLine: boolean;

        /**
         * Specifies if the mean marker is shown in a box and whisker chart.
         */
        showMeanMarker: boolean;

        /**
         * Specifies if outlier points are shown in a box and whisker chart.
         */
        showOutlierPoints: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the formatting options for line elements.
     */
    interface ChartLineFormat {
        /**
         * HTML color code representing the color of lines in the chart.
         */
        color: string;

        /**
         * Represents the line style. See Excel.ChartLineStyle for details.
         */
        lineStyle:
            | ChartLineStyle
            | "None"
            | "Continuous"
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "Dot"
            | "Grey25"
            | "Grey50"
            | "Grey75"
            | "Automatic"
            | "RoundDot";

        /**
         * Represents weight of the line, in points.
         */
        weight: number;

        /**
         * Clear the line format of a chart element.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the font attributes (font name, font size, color, etc.) for a chart object.
     */
    interface ChartFont {
        /**
         * Represents the bold status of font.
         */
        bold: boolean;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        color: string;

        /**
         * Represents the italic status of the font.
         */
        italic: boolean;

        /**
         * Font name (e.g., "Calibri")
         */
        name: string;

        /**
         * Size of the font (e.g., 11)
         */
        size: number;

        /**
         * Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
         */
        underline: ChartUnderlineStyle | "None" | "Single";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the attributes for a chart trendline object.
     */
    interface ChartTrendline {
        /**
         * Represents the number of periods that the trendline extends backward.
         */
        backwardPeriod: number;

        /**
         * Represents the formatting of a chart trendline.
         */
        readonly format: ChartTrendlineFormat;

        /**
         * Represents the number of periods that the trendline extends forward.
         */
        forwardPeriod: number;

        /**
         * Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.
         */
        intercept: any;

        /**
         * Represents the label of a chart trendline.
         */
        readonly label: ChartTrendlineLabel;

        /**
         * Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.
         */
        movingAveragePeriod: number;

        /**
         * Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string
         */
        name: string;

        /**
         * Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.
         */
        polynomialOrder: number;

        /**
         * True if the equation for the trendline is displayed on the chart.
         */
        showEquation: boolean;

        /**
         * True if the R-squared for the trendline is displayed on the chart.
         */
        showRSquared: boolean;

        /**
         * Represents the type of a chart trendline.
         */
        type:
            | ChartTrendlineType
            | "Linear"
            | "Exponential"
            | "Logarithmic"
            | "MovingAverage"
            | "Polynomial"
            | "Power";

        /**
         * Delete the trendline object.
         */
        delete(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of Chart Trendlines.
     */
    interface ChartTrendlineCollection {
        /**
         * Adds a new trendline to trendline collection.
         * @param type Specifies the trendline type. The default value is "Linear". See Excel.ChartTrendline for details.
         */
        add(
            type?:
                | ChartTrendlineType
                | "Linear"
                | "Exponential"
                | "Logarithmic"
                | "MovingAverage"
                | "Polynomial"
                | "Power"
        ): ChartTrendline;

        /**
         * Returns the number of trendlines in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Get trendline object by index, which is the insertion order in items array.
         * @param index Represents the insertion order in items array.
         */
        getItem(index: number): ChartTrendline;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the format properties for chart trendline.
     */
    interface ChartTrendlineFormat {
        /**
         * Represents chart line formatting.
         */
        readonly line: ChartLineFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the attributes for a chart trendline lable object.
     */
    interface ChartTrendlineLabel {
        /**
         * Specifies if trendline label automatically generate appropriate text based on context.
         */
        autoText: boolean;

        /**
         * The format of chart trendline label.
         */
        readonly format: ChartTrendlineLabelFormat;

        /**
         * String value that represents the formula of chart trendline label using A1-style notation.
         */
        formula: string;

        /**
         * Returns the height, in points, of the chart trendline label. Null if chart trendline label is not visible.
         */
        readonly height: number;

        /**
         * Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.
         * This property is valid only when TextOrientation of trendline label is -90, 90, or 180.
         */
        horizontalAlignment:
            | ChartTextHorizontalAlignment
            | "Center"
            | "Left"
            | "Right"
            | "Justify"
            | "Distributed";

        /**
         * Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.
         */
        left: number;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        linkNumberFormat: boolean;

        /**
         * String value that represents the format code for trendline label.
         */
        numberFormat: string;

        /**
         * String representing the text of the trendline label on a chart.
         */
        text: string;

        /**
         * Represents the angle to which the text is oriented for the chart trendline label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        textOrientation: number;

        /**
         * Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.
         */
        top: number;

        /**
         * Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.
         * This property is valid only when TextOrientation of trendline label is 0.
         */
        verticalAlignment:
            | ChartTextVerticalAlignment
            | "Center"
            | "Bottom"
            | "Top"
            | "Justify"
            | "Distributed";

        /**
         * Returns the width, in points, of the chart trendline label. Null if chart trendline label is not visible.
         */
        readonly width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Encapsulates the format properties for the chart trendline label.
     */
    interface ChartTrendlineLabelFormat {
        /**
         * Specifies the border format, which includes color, linestyle, and weight.
         */
        readonly border: ChartBorder;

        /**
         * Specifies the fill format of the current chart trendline label.
         */
        readonly fill: ChartFill;

        /**
         * Specifies the font attributes (font name, font size, color, etc.) for a chart trendline label.
         */
        readonly font: ChartFont;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the attributes for a chart plotArea object.
     */
    interface ChartPlotArea {
        /**
         * Specifies the formatting of a chart plotArea.
         */
        readonly format: ChartPlotAreaFormat;

        /**
         * Specifies the height value of plotArea.
         */
        height: number;

        /**
         * Specifies the insideHeight value of plotArea.
         */
        insideHeight: number;

        /**
         * Specifies the insideLeft value of plotArea.
         */
        insideLeft: number;

        /**
         * Specifies the insideTop value of plotArea.
         */
        insideTop: number;

        /**
         * Specifies the insideWidth value of plotArea.
         */
        insideWidth: number;

        /**
         * Specifies the left value of plotArea.
         */
        left: number;

        /**
         * Specifies the position of plotArea.
         */
        position: ChartPlotAreaPosition | "Automatic" | "Custom";

        /**
         * Specifies the top value of plotArea.
         */
        top: number;

        /**
         * Specifies the width value of plotArea.
         */
        width: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the format properties for chart plotArea.
     */
    interface ChartPlotAreaFormat {
        /**
         * Specifies the border attributes of a chart plotArea.
         */
        readonly border: ChartBorder;

        /**
         * Specifies the fill format of an object, which includes background formatting information.
         */
        readonly fill: ChartFill;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Manages sorting operations on Range objects.
     */
    interface RangeSort {
        /**
         * Perform a sort operation.
         * @param fields The list of conditions to sort on.
         * @param matchCase Optional. Whether to have the casing impact string ordering.
         * @param hasHeaders Optional. Whether the range has a header.
         * @param orientation Optional. Whether the operation is sorting rows or columns.
         * @param method Optional. The ordering method used for Chinese characters.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            hasHeaders?: boolean,
            orientation?: SortOrientation | "Rows" | "Columns",
            method?: SortMethod | "PinYin" | "StrokeCount"
        ): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Manages sorting operations on Table objects.
     */
    interface TableSort {
        /**
         * Specifies the current conditions used to last sort the table.
         */
        readonly fields: SortField[];

        /**
         * Specifies if the casing impacts the last sort of the table.
         */
        readonly matchCase: boolean;

        /**
         * Represents Chinese character ordering method last used to sort the table.
         */
        readonly method: SortMethod | "PinYin" | "StrokeCount";

        /**
         * Perform a sort operation.
         * @param fields The list of conditions to sort on.
         * @param matchCase Optional. Whether to have the casing impact string ordering.
         * @param method Optional. The ordering method used for Chinese characters.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            method?: SortMethod | "PinYin" | "StrokeCount"
        ): void;

        /**
         * Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
         */
        clear(): void;

        /**
         * Reapplies the current sorting parameters to the table.
         */
        reapply(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Manages the filtering of a table's column.
     */
    interface Filter {
        /**
         * The currently applied filter on the given column.
         */
        readonly criteria: FilterCriteria;

        /**
         * Apply the given filter criteria on the given column.
         * @param criteria The criteria to apply.
         */
        apply(criteria: FilterCriteria): void;

        /**
         * Apply a "Bottom Item" filter to the column for the given number of elements.
         * @param count The number of elements from the bottom to show.
         */
        applyBottomItemsFilter(count: number): void;

        /**
         * Apply a "Bottom Percent" filter to the column for the given percentage of elements.
         * @param percent The percentage of elements from the bottom to show.
         */
        applyBottomPercentFilter(percent: number): void;

        /**
         * Apply a "Cell Color" filter to the column for the given color.
         * @param color The background color of the cells to show.
         */
        applyCellColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given criteria strings.
         * @param criteria1 The first criteria string.
         * @param criteria2 Optional. The second criteria string.
         * @param oper Optional. The operator that describes how the two criteria are joined.
         */
        applyCustomFilter(
            criteria1: string,
            criteria2?: string,
            oper?: FilterOperator | "And" | "Or"
        ): void;

        /**
         * Apply a "Dynamic" filter to the column.
         * @param criteria The dynamic criteria to apply.
         */
        applyDynamicFilter(
            criteria:
                | DynamicFilterCriteria
                | "Unknown"
                | "AboveAverage"
                | "AllDatesInPeriodApril"
                | "AllDatesInPeriodAugust"
                | "AllDatesInPeriodDecember"
                | "AllDatesInPeriodFebruray"
                | "AllDatesInPeriodJanuary"
                | "AllDatesInPeriodJuly"
                | "AllDatesInPeriodJune"
                | "AllDatesInPeriodMarch"
                | "AllDatesInPeriodMay"
                | "AllDatesInPeriodNovember"
                | "AllDatesInPeriodOctober"
                | "AllDatesInPeriodQuarter1"
                | "AllDatesInPeriodQuarter2"
                | "AllDatesInPeriodQuarter3"
                | "AllDatesInPeriodQuarter4"
                | "AllDatesInPeriodSeptember"
                | "BelowAverage"
                | "LastMonth"
                | "LastQuarter"
                | "LastWeek"
                | "LastYear"
                | "NextMonth"
                | "NextQuarter"
                | "NextWeek"
                | "NextYear"
                | "ThisMonth"
                | "ThisQuarter"
                | "ThisWeek"
                | "ThisYear"
                | "Today"
                | "Tomorrow"
                | "YearToDate"
                | "Yesterday"
        ): void;

        /**
         * Apply a "Font Color" filter to the column for the given color.
         * @param color The font color of the cells to show.
         */
        applyFontColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given icon.
         * @param icon The icons of the cells to show.
         */
        applyIconFilter(icon: Icon): void;

        /**
         * Apply a "Top Item" filter to the column for the given number of elements.
         * @param count The number of elements from the top to show.
         */
        applyTopItemsFilter(count: number): void;

        /**
         * Apply a "Top Percent" filter to the column for the given percentage of elements.
         * @param percent The percentage of elements from the top to show.
         */
        applyTopPercentFilter(percent: number): void;

        /**
         * Apply a "Values" filter to the column for the given values.
         * @param values The list of values to show. This must be an array of strings or an array of Excel.FilterDateTime objects.
         */
        applyValuesFilter(values: Array<string | FilterDatetime>): void;

        /**
         * Clear the filter on the given column.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the AutoFilter object.
     * AutoFilter turns the values in Excel column into specific filters based on the cell contents.
     */
    interface AutoFilter {
        /**
         * An array that holds all the filter criteria in the autofiltered range.
         */
        readonly criteria: FilterCriteria[];

        /**
         * Specifies if the AutoFilter is enabled.
         */
        readonly enabled: boolean;

        /**
         * Specifies if the AutoFilter has filter criteria.
         */
        readonly isDataFiltered: boolean;

        /**
         * Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.
         * @param range The range over which the AutoFilter will apply on.
         * @param columnIndex The zero-based column index to which the AutoFilter is applied.
         * @param criteria The filter criteria.
         */
        apply(
            range: Range | string,
            columnIndex?: number,
            criteria?: FilterCriteria
        ): void;

        /**
         * Clears the filter criteria of the AutoFilter.
         */
        clearCriteria(): void;

        /**
         * Returns the Range object that represents the range to which the AutoFilter applies.
         */
        getRange(): Range;

        /**
         * Returns the Range object that represents the range to which the AutoFilter applies.
         * If there is no Range object associated with the AutoFilter, this method returns a null object.
         */
        getRangeOrNullObject(): Range;

        /**
         * Applies the specified Autofilter object currently on the range.
         */
        reapply(): void;

        /**
         * Removes the AutoFilter for the range.
         */
        remove(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.
     */
    interface CultureInfo {
        /**
         * Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us"). This is based on current system settings.
         */
        readonly name: string;

        /**
         * Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.
         */
        readonly numberFormat: NumberFormatInfo;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.
     */
    interface NumberFormatInfo {
        /**
         * Gets the string used as the decimal separator for numeric values. This is based on current system settings.
         */
        readonly numberDecimalSeparator: string;

        /**
         * Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.
         */
        readonly numberGroupSeparator: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A scoped collection of custom XML parts.
     * A scoped collection is the result of some operation (e.g., filtering by namespace).
     * A scoped collection cannot be scoped any further.
     */
    interface CustomXmlPartScopedCollection {
        /**
         * Gets the number of CustomXML parts in this collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a custom XML part based on its ID.
         * @param id ID of the object to be retrieved.
         */
        getItem(id: string): CustomXmlPart;

        /**
         * Gets a custom XML part based on its ID.
         * If the CustomXmlPart does not exist, the return object's isNull property will be true.
         * @param id ID of the object to be retrieved.
         */
        getItemOrNullObject(id: string): CustomXmlPart;

        /**
         * If the collection contains exactly one item, this method returns it.
         * Otherwise, this method produces an error.
         */
        getOnlyItem(): CustomXmlPart;

        /**
         * If the collection contains exactly one item, this method returns it.
         * Otherwise, this method returns Null.
         */
        getOnlyItemOrNullObject(): CustomXmlPart;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A collection of custom XML parts.
     */
    interface CustomXmlPartCollection {
        /**
         * Adds a new custom XML part to the workbook.
         * @param xml XML content. Must be a valid XML fragment.
         */
        add(xml: string): CustomXmlPart;

        /**
         * Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
         * @param namespaceUri This must be a fully qualified schema URI; for example, "http://schemas.contoso.com/review/1.0".
         */
        getByNamespace(namespaceUri: string): CustomXmlPartScopedCollection;

        /**
         * Gets the number of CustomXml parts in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a custom XML part based on its ID.
         * @param id ID of the object to be retrieved.
         */
        getItem(id: string): CustomXmlPart;

        /**
         * Gets a custom XML part based on its ID.
         * If the CustomXmlPart does not exist, the return object's isNull property will be true.
         * @param id ID of the object to be retrieved.
         */
        getItemOrNullObject(id: string): CustomXmlPart;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a custom XML part object in a workbook.
     */
    interface CustomXmlPart {
        /**
         * The custom XML part's ID.
         */
        readonly id: string;

        /**
         * The custom XML part's namespace URI.
         */
        readonly namespaceUri: string;

        /**
         * Deletes the custom XML part.
         */
        delete(): void;

        /**
         * Gets the custom XML part's full XML content.
         */
        getXml(): ClientResult<string>;

        /**
         * Sets the custom XML part's full XML content.
         * @param xml XML content for the part.
         */
        setXml(xml: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a scoped collection of PivotTables. The PivotTables are sorted based on the location of the PivotTable's top-left corner. They are ordered top to bottom and then left to right.
     */
    interface PivotTableScopedCollection {
        /**
         * Gets the number of PivotTables in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the first PivotTable in the collection. The PivotTables in the collection are sorted top to bottom and left to right, such that top-left table is the first PivotTable in the collection.
         */
        getFirst(): PivotTable;

        /**
         * Gets a PivotTable by name.
         * @param key Name of the PivotTable to be retrieved.
         */
        getItem(key: string): PivotTable;

        /**
         * Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.
         * @param name Name of the PivotTable to be retrieved.
         */
        getItemOrNullObject(name: string): PivotTable;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the PivotTables that are part of the workbook or worksheet.
     */
    interface PivotTableCollection {
        /**
         * Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.
         * @param name The name of the new PivotTable.
         * @param source The source data for the new PivotTable, this can either be a range (or string address including the worksheet name) or a table.
         * @param destination The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed).
         */
        add(
            name: string,
            source: Range | string | Table,
            destination: Range | string
        ): PivotTable;

        /**
         * Gets the number of pivot tables in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a PivotTable by name.
         * @param name Name of the PivotTable to be retrieved.
         */
        getItem(name: string): PivotTable;

        /**
         * Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.
         * @param name Name of the PivotTable to be retrieved.
         */
        getItemOrNullObject(name: string): PivotTable;

        /**
         * Refreshes all the pivot tables in the collection.
         */
        refreshAll(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an Excel PivotTable.
     * To learn more about the PivotTable object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables | Work with PivotTables using the Excel JavaScript API}.
     */
    interface PivotTable {
        /**
         * The Column Pivot Hierarchies of the PivotTable.
         */
        readonly columnHierarchies: RowColumnPivotHierarchyCollection;

        /**
         * The Data Pivot Hierarchies of the PivotTable.
         */
        readonly dataHierarchies: DataPivotHierarchyCollection;

        /**
         * Specifies if the PivotTable allows values in the data body to be edited by the user.
         */
        enableDataValueEditing: boolean;

        /**
         * The Filter Pivot Hierarchies of the PivotTable.
         */
        readonly filterHierarchies: FilterPivotHierarchyCollection;

        /**
         * The Pivot Hierarchies of the PivotTable.
         */
        readonly hierarchies: PivotHierarchyCollection;

        /**
         * Id of the PivotTable.
         */
        readonly id: string;

        /**
         * The PivotLayout describing the layout and visual structure of the PivotTable.
         */
        readonly layout: PivotLayout;

        /**
         * Name of the PivotTable.
         */
        name: string;

        /**
         * The Row Pivot Hierarchies of the PivotTable.
         */
        readonly rowHierarchies: RowColumnPivotHierarchyCollection;

        /**
         * Specifies if the PivotTable uses custom lists when sorting.
         */
        useCustomSortLists: boolean;

        /**
         * The worksheet containing the current PivotTable.
         */
        readonly worksheet: Worksheet;

        /**
         * Deletes the PivotTable.
         */
        delete(): void;

        /**
         * Refreshes the PivotTable.
         */
        refresh(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the visual layout of the PivotTable.
     */
    interface PivotLayout {
        /**
         * Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.
         */
        autoFormat: boolean;

        /**
         * Specifies if the field list can be shown in the UI.
         */
        enableFieldList: boolean;

        /**
         * This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        layoutType: PivotLayoutType | "Compact" | "Tabular" | "Outline";

        /**
         * Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
         */
        preserveFormatting: boolean;

        /**
         * Specifies if the PivotTable report shows grand totals for columns.
         */
        showColumnGrandTotals: boolean;

        /**
         * Specifies if the PivotTable report shows grand totals for rows.
         */
        showRowGrandTotals: boolean;

        /**
         * This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        subtotalLocation: SubtotalLocationType | "AtTop" | "AtBottom" | "Off";

        /**
         * Returns the range where the PivotTable's column labels reside.
         */
        getColumnLabelRange(): Range;

        /**
         * Returns the range where the PivotTable's data values reside.
         */
        getDataBodyRange(): Range;

        /**
         * Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.
         * @param cell A single cell within the PivotTable data body.
         */
        getDataHierarchy(cell: Range | string): DataPivotHierarchy;

        /**
         * Returns the range of the PivotTable's filter area.
         */
        getFilterAxisRange(): Range;

        /**
         * Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.
         * @param axis The axis from which to get the PivotItems. Must be either "row" or "column."
         * @param cell A single cell within the PivotTable's data body.
         */
        getPivotItems(
            axis: PivotAxis | "Unknown" | "Row" | "Column" | "Data" | "Filter",
            cell: Range | string
        ): ClientResult<PivotItem[]>;

        /**
         * Returns the range the PivotTable exists on, excluding the filter area.
         */
        getRange(): Range;

        /**
         * Returns the range where the PivotTable's row labels reside.
         */
        getRowLabelRange(): Range;

        /**
         * Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.
         * @param cell A single cell to use get the criteria from for applying the autosort.
         * @param sortBy The direction of the sort.
         */
        setAutoSortOnCell(
            cell: Range | string,
            sortBy: SortBy | "Ascending" | "Descending"
        ): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the PivotHierarchies that are part of the PivotTable.
     */
    interface PivotHierarchyCollection {
        /**
         * Gets the number of pivot hierarchies in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a PivotHierarchy by its name or id.
         * @param name Name of the PivotHierarchy to be retrieved.
         */
        getItem(name: string): PivotHierarchy;

        /**
         * Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.
         * @param name Name of the PivotHierarchy to be retrieved.
         */
        getItemOrNullObject(name: string): PivotHierarchy;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel PivotHierarchy.
     */
    interface PivotHierarchy {
        /**
         * Returns the PivotFields associated with the PivotHierarchy.
         */
        readonly fields: PivotFieldCollection;

        /**
         * Id of the PivotHierarchy.
         */
        readonly id: string;

        /**
         * Name of the PivotHierarchy.
         */
        name: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of RowColumnPivotHierarchy items associated with the PivotTable.
     */
    interface RowColumnPivotHierarchyCollection {
        /**
         * Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,
         * or filter axis, it will be removed from that location.
         */
        add(pivotHierarchy: PivotHierarchy): RowColumnPivotHierarchy;

        /**
         * Gets the number of pivot hierarchies in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a RowColumnPivotHierarchy by its name or id.
         * @param name Name of the RowColumnPivotHierarchy to be retrieved.
         */
        getItem(name: string): RowColumnPivotHierarchy;

        /**
         * Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.
         * @param name Name of the RowColumnPivotHierarchy to be retrieved.
         */
        getItemOrNullObject(name: string): RowColumnPivotHierarchy;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel RowColumnPivotHierarchy.
     */
    interface RowColumnPivotHierarchy {
        /**
         * Returns the PivotFields associated with the RowColumnPivotHierarchy.
         */
        readonly fields: PivotFieldCollection;

        /**
         * Id of the RowColumnPivotHierarchy.
         */
        readonly id: string;

        /**
         * Name of the RowColumnPivotHierarchy.
         */
        name: string;

        /**
         * Position of the RowColumnPivotHierarchy.
         */
        position: number;

        /**
         * Reset the RowColumnPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of FilterPivotHierarchy items associated with the PivotTable.
     */
    interface FilterPivotHierarchyCollection {
        /**
         * Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,
         * or filter axis, it will be removed from that location.
         */
        add(pivotHierarchy: PivotHierarchy): FilterPivotHierarchy;

        /**
         * Gets the number of pivot hierarchies in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a FilterPivotHierarchy by its name or id.
         * @param name Name of the FilterPivotHierarchy to be retrieved.
         */
        getItem(name: string): FilterPivotHierarchy;

        /**
         * Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.
         * @param name Name of the FilterPivotHierarchy to be retrieved.
         */
        getItemOrNullObject(name: string): FilterPivotHierarchy;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        remove(filterPivotHierarchy: FilterPivotHierarchy): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel FilterPivotHierarchy.
     */
    interface FilterPivotHierarchy {
        /**
         * Determines whether to allow multiple filter items.
         */
        enableMultipleFilterItems: boolean;

        /**
         * Returns the PivotFields associated with the FilterPivotHierarchy.
         */
        readonly fields: PivotFieldCollection;

        /**
         * Id of the FilterPivotHierarchy.
         */
        readonly id: string;

        /**
         * Name of the FilterPivotHierarchy.
         */
        name: string;

        /**
         * Position of the FilterPivotHierarchy.
         */
        position: number;

        /**
         * Reset the FilterPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of DataPivotHierarchy items associated with the PivotTable.
     */
    interface DataPivotHierarchyCollection {
        /**
         * Adds the PivotHierarchy to the current axis.
         */
        add(pivotHierarchy: PivotHierarchy): DataPivotHierarchy;

        /**
         * Gets the number of pivot hierarchies in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a DataPivotHierarchy by its name or id.
         * @param name Name of the DataPivotHierarchy to be retrieved.
         */
        getItem(name: string): DataPivotHierarchy;

        /**
         * Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.
         * @param name Name of the DataPivotHierarchy to be retrieved.
         */
        getItemOrNullObject(name: string): DataPivotHierarchy;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        remove(DataPivotHierarchy: DataPivotHierarchy): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel DataPivotHierarchy.
     */
    interface DataPivotHierarchy {
        /**
         * Returns the PivotFields associated with the DataPivotHierarchy.
         */
        readonly field: PivotField;

        /**
         * Id of the DataPivotHierarchy.
         */
        readonly id: string;

        /**
         * Name of the DataPivotHierarchy.
         */
        name: string;

        /**
         * Number format of the DataPivotHierarchy.
         */
        numberFormat: string;

        /**
         * Position of the DataPivotHierarchy.
         */
        position: number;

        /**
         * Specifies if the data should be shown as a specific summary calculation.
         */
        showAs: ShowAsRule;

        /**
         * Specifies if all items of the DataPivotHierarchy are shown.
         */
        summarizeBy:
            | AggregationFunction
            | "Unknown"
            | "Automatic"
            | "Sum"
            | "Count"
            | "Average"
            | "Max"
            | "Min"
            | "Product"
            | "CountNumbers"
            | "StandardDeviation"
            | "StandardDeviationP"
            | "Variance"
            | "VarianceP";

        /**
         * Reset the DataPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the PivotFields that are part of a PivotTable's hierarchy.
     */
    interface PivotFieldCollection {
        /**
         * Gets the number of pivot fields in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a PivotField by its name or id.
         * @param name Name of the PivotField to be retrieved.
         */
        getItem(name: string): PivotField;

        /**
         * Gets a PivotField by name. If the PivotField does not exist, will return a null object.
         * @param name Name of the PivotField to be retrieved.
         */
        getItemOrNullObject(name: string): PivotField;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel PivotField.
     */
    interface PivotField {
        /**
         * Id of the PivotField.
         */
        readonly id: string;

        /**
         * Returns the PivotFields associated with the PivotField.
         */
        readonly items: PivotItemCollection;

        /**
         * Name of the PivotField.
         */
        name: string;

        /**
         * Determines whether to show all items of the PivotField.
         */
        showAllItems: boolean;

        /**
         * Subtotals of the PivotField.
         */
        subtotals: Subtotals;

        /**
         * Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.
         * @param sortBy Specifies if the sorting is done in ascending or descending order.
         */
        sortByLabels(sortBy: SortBy | "Ascending" | "Descending"): void;

        /**
         * Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when
         * there are multiple values from the same DataPivotHierarchy.
         * @param sortBy Specifies if the sorting is done in ascending or descending order.
         * @param valuesHierarchy Specifies the values hierarchy on the data axis to be used for sorting.
         * @param pivotItemScope The items that should be used for the scope of the sorting. These will be the
         * items that make up the row or column that you want to sort on. If a string is used instead of a PivotItem,
         * the string represents the ID of the PivotItem. If there are no items other than data hierarchy on the axis
         * you want to sort on, this can be empty.
         */
        sortByValues(
            sortBy: SortBy | "Ascending" | "Descending",
            valuesHierarchy: DataPivotHierarchy,
            pivotItemScope?: Array<PivotItem | string>
        ): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the PivotItems related to their parent PivotField.
     */
    interface PivotItemCollection {
        /**
         * Gets the number of PivotItems in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a PivotItem by its name or id.
         * @param name Name of the PivotItem to be retrieved.
         */
        getItem(name: string): PivotItem;

        /**
         * Gets a PivotItem by name. If the PivotItem does not exist, will return a null object.
         * @param name Name of the PivotItem to be retrieved.
         */
        getItemOrNullObject(name: string): PivotItem;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the Excel PivotItem.
     */
    interface PivotItem {
        /**
         * Id of the PivotItem.
         */
        readonly id: string;

        /**
         * Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.
         */
        isExpanded: boolean;

        /**
         * Name of the PivotItem.
         */
        name: string;

        /**
         * Specifies if the PivotItem is visible.
         */
        visible: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents workbook properties.
     */
    interface DocumentProperties {
        /**
         * The author of the workbook.
         */
        author: string;

        /**
         * The category of the workbook.
         */
        category: string;

        /**
         * The comments of the workbook.
         */
        comments: string;

        /**
         * The company of the workbook.
         */
        company: string;

        /**
         * Gets the creation date of the workbook. Read only.
         */
        readonly creationDate: Date;

        /**
         * Gets the collection of custom properties of the workbook. Read only.
         */
        readonly custom: CustomPropertyCollection;

        /**
         * The keywords of the workbook.
         */
        keywords: string;

        /**
         * Gets the last author of the workbook. Read only.
         */
        readonly lastAuthor: string;

        /**
         * The manager of the workbook.
         */
        manager: string;

        /**
         * Gets the revision number of the workbook. Read only.
         */
        revisionNumber: number;

        /**
         * The subject of the workbook.
         */
        subject: string;

        /**
         * The title of the workbook.
         */
        title: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a custom property.
     */
    interface CustomProperty {
        /**
         * The key of the custom property. The key is limited to 255 characters outside of Excel on the web (larger keys are automatically trimmed to 255 characters on other platforms).
         */
        readonly key: string;

        /**
         * The type of the value used for the custom property.
         */
        readonly type:
            | DocumentPropertyType
            | "Number"
            | "Boolean"
            | "Date"
            | "String"
            | "Float";

        /**
         * The value of the custom property. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).
         */
        value: any;

        /**
         * Deletes the custom property.
         */
        delete(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Contains the collection of customProperty objects.
     */
    interface CustomPropertyCollection {
        /**
         * Creates a new or sets an existing custom property.
         * @param key Required. The custom property's key, which is case-insensitive. The key is limited to 255 characters outside of Excel on the web (larger keys are automatically trimmed to 255 characters on other platforms).
         * @param value Required. The custom property's value. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).
         */
        add(key: string, value: any): CustomProperty;

        /**
         * Deletes all custom properties in this collection.
         */
        deleteAll(): void;

        /**
         * Gets the count of custom properties.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.
         * @param key The key that identifies the custom property object.
         */
        getItem(key: string): CustomProperty;

        /**
         * Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.
         * @param key Required. The key that identifies the custom property object.
         */
        getItemOrNullObject(key: string): CustomProperty;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the conditional formats that are overlap the range.
     */
    interface ConditionalFormatCollection {
        /**
         * Adds a new conditional format to the collection at the first/top priority.
         * @param type The type of conditional format being added. See Excel.ConditionalFormatType for details.
         */
        add(
            type:
                | ConditionalFormatType
                | "Custom"
                | "DataBar"
                | "ColorScale"
                | "IconSet"
                | "TopBottom"
                | "PresetCriteria"
                | "ContainsText"
                | "CellValue"
        ): ConditionalFormat;

        /**
         * Clears all conditional formats active on the current specified range.
         */
        clearAll(): void;

        /**
         * Returns the number of conditional formats in the workbook.
         */
        getCount(): ClientResult<number>;

        /**
         * Returns a conditional format for the given ID.
         * @param id The id of the conditional format.
         */
        getItem(id: string): ConditionalFormat;

        /**
         * Returns a conditional format at the given index.
         * @param index Index of the conditional formats to be retrieved.
         */
        getItemAt(index: number): ConditionalFormat;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * An object encapsulating a conditional format's range, format, rule, and other properties.
     * To learn more about the conditional formatting object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-conditional-formatting | Apply conditional formatting to Excel ranges}.
     */
    interface ConditionalFormat {
        /**
         * Returns the cell value conditional format properties if the current conditional format is a CellValue type.
         * For example to format all cells between 5 and 10.
         */
        readonly cellValue: CellValueConditionalFormat;

        /**
         * Returns the cell value conditional format properties if the current conditional format is a CellValue type.
         * For example to format all cells between 5 and 10.
         */
        readonly cellValueOrNullObject: CellValueConditionalFormat;

        /**
         * Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.
         */
        readonly colorScale: ColorScaleConditionalFormat;

        /**
         * Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.
         */
        readonly colorScaleOrNullObject: ColorScaleConditionalFormat;

        /**
         * Returns the custom conditional format properties if the current conditional format is a custom type.
         */
        readonly custom: CustomConditionalFormat;

        /**
         * Returns the custom conditional format properties if the current conditional format is a custom type.
         */
        readonly customOrNullObject: CustomConditionalFormat;

        /**
         * Returns the data bar properties if the current conditional format is a data bar.
         */
        readonly dataBar: DataBarConditionalFormat;

        /**
         * Returns the data bar properties if the current conditional format is a data bar.
         */
        readonly dataBarOrNullObject: DataBarConditionalFormat;

        /**
         * Returns the IconSet conditional format properties if the current conditional format is an IconSet type.
         */
        readonly iconSet: IconSetConditionalFormat;

        /**
         * Returns the IconSet conditional format properties if the current conditional format is an IconSet type.
         */
        readonly iconSetOrNullObject: IconSetConditionalFormat;

        /**
         * The Priority of the Conditional Format within the current ConditionalFormatCollection.
         */
        readonly id: string;

        /**
         * Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.
         */
        readonly preset: PresetCriteriaConditionalFormat;

        /**
         * Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.
         */
        readonly presetOrNullObject: PresetCriteriaConditionalFormat;

        /**
         * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
         * changes other conditional formats' priorities, to allow for a contiguous priority order.
         * Use a negative priority to begin from the back.
         * Priorities greater than than bounds will get and set to the maximum (or minimum if negative) priority.
         * Also note that if you change the priority, you have to re-fetch a new copy of the object at that new priority location if you want to make further changes to it.
         */
        priority: number;

        /**
         * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
         * Null on databars, icon sets, and colorscales as there's no concept of StopIfTrue for these
         */
        stopIfTrue: boolean;

        /**
         * Returns the specific text conditional format properties if the current conditional format is a text type.
         * For example to format cells matching the word "Text".
         */
        readonly textComparison: TextConditionalFormat;

        /**
         * Returns the specific text conditional format properties if the current conditional format is a text type.
         * For example to format cells matching the word "Text".
         */
        readonly textComparisonOrNullObject: TextConditionalFormat;

        /**
         * Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.
         * For example to format the top 10% or bottom 10 items.
         */
        readonly topBottom: TopBottomConditionalFormat;

        /**
         * Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.
         * For example to format the top 10% or bottom 10 items.
         */
        readonly topBottomOrNullObject: TopBottomConditionalFormat;

        /**
         * A type of conditional format. Only one can be set at a time.
         */
        readonly type:
            | ConditionalFormatType
            | "Custom"
            | "DataBar"
            | "ColorScale"
            | "IconSet"
            | "TopBottom"
            | "PresetCriteria"
            | "ContainsText"
            | "CellValue";

        /**
         * Deletes this conditional format.
         */
        delete(): void;

        /**
         * Returns the range the conditonal format is applied to. Throws an error if the conditional format is applied to multiple ranges.
         */
        getRange(): Range;

        /**
         * Returns the range the conditonal format is applied to, or a null object if the conditional format is applied to multiple ranges.
         */
        getRangeOrNullObject(): Range;

        /**
         * Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to.
         */
        getRanges(): RangeAreas;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an Excel Conditional Data Bar Type.
     */
    interface DataBarConditionalFormat {
        /**
         * HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * "" (empty string) if no axis is present or set.
         */
        axisColor: string;

        /**
         * Representation of how the axis is determined for an Excel data bar.
         */
        axisFormat:
            | ConditionalDataBarAxisFormat
            | "Automatic"
            | "None"
            | "CellMidPoint";

        /**
         * Specifies the direction that the data bar graphic should be based on.
         */
        barDirection:
            | ConditionalDataBarDirection
            | "Context"
            | "LeftToRight"
            | "RightToLeft";

        /**
         * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.lowerBoundRule = {...}` instead of `x.lowerBoundRule.formula = ...`).
         */
        lowerBoundRule: ConditionalDataBarRule;

        /**
         * Representation of all values to the left of the axis in an Excel data bar.
         */
        readonly negativeFormat: ConditionalDataBarNegativeFormat;

        /**
         * Representation of all values to the right of the axis in an Excel data bar.
         */
        readonly positiveFormat: ConditionalDataBarPositiveFormat;

        /**
         * If true, hides the values from the cells where the data bar is applied.
         */
        showDataBarOnly: boolean;

        /**
         * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.upperBoundRule = {...}` instead of `x.upperBoundRule.formula = ...`).
         */
        upperBoundRule: ConditionalDataBarRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a conditional format DataBar Format for the positive side of the data bar.
     */
    interface ConditionalDataBarPositiveFormat {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * "" (empty string) if no border is present or set.
         */
        borderColor: string;

        /**
         * HTML color code representing the fill color, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        fillColor: string;

        /**
         * Specifies if the DataBar has a gradient.
         */
        gradientFill: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a conditional format DataBar Format for the negative side of the data bar.
     */
    interface ConditionalDataBarNegativeFormat {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * "Empty String" if no border is present or set.
         */
        borderColor: string;

        /**
         * HTML color code representing the fill color, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        fillColor: string;

        /**
         * Specifies if the negative DataBar has the same border color as the positive DataBar.
         */
        matchPositiveBorderColor: boolean;

        /**
         * Specifies if the negative DataBar has the same fill color as the positive DataBar.
         */
        matchPositiveFillColor: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a custom conditional format type.
     */
    interface CustomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        readonly format: ConditionalRangeFormat;

        /**
         * Specifies the Rule object on this conditional format.
         */
        readonly rule: ConditionalFormatRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a rule, for all traditional rule/format pairings.
     */
    interface ConditionalFormatRule {
        /**
         * The formula, if required, to evaluate the conditional format rule on.
         */
        formula: string;

        /**
         * The formula, if required, to evaluate the conditional format rule on in the user's language.
         */
        formulaLocal: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an IconSet criteria for conditional formatting.
     */
    interface IconSetConditionalFormat {
        /**
         * An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.
         */
        criteria: ConditionalIconCriterion[];

        /**
         * If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.
         */
        reverseIconOrder: boolean;

        /**
         * If true, hides the values and only shows icons.
         */
        showIconOnly: boolean;

        /**
         * If set, displays the IconSet option for the conditional format.
         */
        style:
            | IconSet
            | "Invalid"
            | "ThreeArrows"
            | "ThreeArrowsGray"
            | "ThreeFlags"
            | "ThreeTrafficLights1"
            | "ThreeTrafficLights2"
            | "ThreeSigns"
            | "ThreeSymbols"
            | "ThreeSymbols2"
            | "FourArrows"
            | "FourArrowsGray"
            | "FourRedToBlack"
            | "FourRating"
            | "FourTrafficLights"
            | "FiveArrows"
            | "FiveArrowsGray"
            | "FiveRating"
            | "FiveQuarters"
            | "ThreeStars"
            | "ThreeTriangles"
            | "FiveBoxes";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents ColorScale criteria for conditional formatting.
     */
    interface ColorScaleConditionalFormat {
        /**
         * The criteria of the color scale. Midpoint is optional when using a two point color scale.
         */
        criteria: ConditionalColorScaleCriteria;

        /**
         * If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).
         */
        readonly threeColorScale: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a Top/Bottom conditional format.
     */
    interface TopBottomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        readonly format: ConditionalRangeFormat;

        /**
         * The criteria of the Top/Bottom conditional format.
         */
        rule: ConditionalTopBottomRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the the preset criteria conditional format such as above average, below average, unique values, contains blank, nonblank, error, and noerror.
     */
    interface PresetCriteriaConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        readonly format: ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        rule: ConditionalPresetCriteriaRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a specific text conditional format.
     */
    interface TextConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        readonly format: ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        rule: ConditionalTextComparisonRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a cell value conditional format.
     */
    interface CellValueConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        readonly format: ConditionalRangeFormat;

        /**
         * Specifies the Rule object on this conditional format.
         */
        rule: ConditionalCellValueRule;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * A format object encapsulating the conditional formats range's font, fill, borders, and other properties.
     */
    interface ConditionalRangeFormat {
        /**
         * Collection of border objects that apply to the overall conditional format range.
         */
        readonly borders: ConditionalRangeBorderCollection;

        /**
         * Returns the fill object defined on the overall conditional format range.
         */
        readonly fill: ConditionalRangeFill;

        /**
         * Returns the font object defined on the overall conditional format range.
         */
        readonly font: ConditionalRangeFont;

        /**
         * Represents Excel's number format code for the given range. Cleared if null is passed in.
         */
        numberFormat: any;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * This object represents the font attributes (font style, color, etc.) for an object.
     */
    interface ConditionalRangeFont {
        /**
         * Specifies if the font is bold.
         */
        bold: boolean;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        color: string;

        /**
         * Specifies if the font is italic.
         */
        italic: boolean;

        /**
         * Specifies the strikethrough status of the font.
         */
        strikethrough: boolean;

        /**
         * The type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.
         */
        underline:
            | ConditionalRangeFontUnderlineStyle
            | "None"
            | "Single"
            | "Double";

        /**
         * Resets the font formats.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the background of a conditional range object.
     */
    interface ConditionalRangeFill {
        /**
         * HTML color code representing the color of the fill, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        color: string;

        /**
         * Resets the fill.
         */
        clear(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the border of an object.
     */
    interface ConditionalRangeBorder {
        /**
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        color: string;

        /**
         * Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details.
         */
        readonly sideIndex:
            | ConditionalRangeBorderIndex
            | "EdgeTop"
            | "EdgeBottom"
            | "EdgeLeft"
            | "EdgeRight";

        /**
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         */
        style:
            | ConditionalRangeBorderLineStyle
            | "None"
            | "Continuous"
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "Dot";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the border objects that make up range border.
     */
    interface ConditionalRangeBorderCollection {
        /**
         * Gets the bottom border.
         */
        readonly bottom: ConditionalRangeBorder;

        /**
         * Number of border objects in the collection.
         */
        readonly count: number;

        /**
         * Gets the left border.
         */
        readonly left: ConditionalRangeBorder;

        /**
         * Gets the right border.
         */
        readonly right: ConditionalRangeBorder;

        /**
         * Gets the top border.
         */
        readonly top: ConditionalRangeBorder;

        /**
         * Gets a border object using its name.
         * @param index Index value of the border object to be retrieved. See Excel.ConditionalRangeBorderIndex for details.
         */
        getItem(
            index:
                | ConditionalRangeBorderIndex
                | "EdgeTop"
                | "EdgeBottom"
                | "EdgeLeft"
                | "EdgeRight"
        ): ConditionalRangeBorder;

        /**
         * Gets a border object using its index.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): ConditionalRangeBorder;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * An object encapsulating a style's format and other properties.
     */
    interface Style {
        /**
         * Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.
         */
        autoIndent: boolean;

        /**
         * A Border collection of four Border objects that represent the style of the four borders.
         */
        readonly borders: RangeBorderCollection;

        /**
         * Specifies if the style is a built-in style.
         */
        readonly builtIn: boolean;

        /**
         * The Fill of the style.
         */
        readonly fill: RangeFill;

        /**
         * A Font object that represents the font of the style.
         */
        readonly font: RangeFont;

        /**
         * Specifies if the formula will be hidden when the worksheet is protected.
         */
        formulaHidden: boolean;

        /**
         * Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.
         */
        horizontalAlignment:
            | HorizontalAlignment
            | "General"
            | "Left"
            | "Center"
            | "Right"
            | "Fill"
            | "Justify"
            | "CenterAcrossSelection"
            | "Distributed";

        /**
         * Specifies if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.
         */
        includeAlignment: boolean;

        /**
         * Specifies if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.
         */
        includeBorder: boolean;

        /**
         * Specifies if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.
         */
        includeFont: boolean;

        /**
         * Specifies if the style includes the NumberFormat property.
         */
        includeNumber: boolean;

        /**
         * Specifies if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.
         */
        includePatterns: boolean;

        /**
         * Specifies if the style includes the FormulaHidden and Locked protection properties.
         */
        includeProtection: boolean;

        /**
         * An integer from 0 to 250 that indicates the indent level for the style.
         */
        indentLevel: number;

        /**
         * Specifies if the object is locked when the worksheet is protected.
         */
        locked: boolean;

        /**
         * The name of the style.
         */
        readonly name: string;

        /**
         * The format code of the number format for the style.
         */
        numberFormat: string;

        /**
         * The localized format code of the number format for the style.
         */
        numberFormatLocal: string;

        /**
         * The reading order for the style.
         */
        readingOrder: ReadingOrder | "Context" | "LeftToRight" | "RightToLeft";

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        shrinkToFit: boolean;

        /**
         * The text orientation for the style.
         */
        textOrientation: number;

        /**
         * Specifies the vertical alignment for the style. See Excel.VerticalAlignment for details.
         */
        verticalAlignment:
            | VerticalAlignment
            | "Top"
            | "Center"
            | "Bottom"
            | "Justify"
            | "Distributed";

        /**
         * Specifies if Excel wraps the text in the object.
         */
        wrapText: boolean;

        /**
         * Deletes this style.
         */
        delete(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the styles.
     */
    interface StyleCollection {
        /**
         * Adds a new style to the collection.
         * @param name Name of the style to be added.
         */
        add(name: string): void;

        /**
         * Gets the number of styles in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a style by name.
         * @param name Name of the style to be retrieved.
         */
        getItem(name: string): Style;

        /**
         * Gets a style based on its position in the collection.
         * @param index Index value of the style object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Style;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of TableStyles.
     */
    interface TableStyleCollection {
        /**
         * Creates a blank TableStyle with the specified name.
         * @param name The unique name for the new TableStyle. Will throw an invalid argument exception if the name is already in use.
         * @param makeUniqueName Optional, defaults to false. If true, will append numbers to the name in order to make it unique, if needed.
         */
        add(name: string, makeUniqueName?: boolean): TableStyle;

        /**
         * Gets the number of table styles in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the default TableStyle for the parent object's scope.
         */
        getDefault(): TableStyle;

        /**
         * Gets a TableStyle by name.
         * @param name Name of the TableStyle to be retrieved.
         */
        getItem(name: string): TableStyle;

        /**
         * Gets a TableStyle by name. If the TableStyle does not exist, will return a null object.
         * @param name Name of the TableStyle to be retrieved.
         */
        getItemOrNullObject(name: string): TableStyle;

        /**
         * Sets the default TableStyle for use in the parent object's scope.
         * @param newDefaultStyle The TableStyle object or name of the TableStyle object that should be the new default.
         */
        setDefault(newDefaultStyle: TableStyle | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a TableStyle, which defines the style elements by region of the Table.
     */
    interface TableStyle {
        /**
         * Gets the name of the TableStyle.
         */
        name: string;

        /**
         * Specifies if this TableStyle object is read-only.
         */
        readonly readOnly: boolean;

        /**
         * Deletes the TableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this TableStyle with copies of all the style elements.
         */
        duplicate(): TableStyle;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of PivotTable styles.
     */
    interface PivotTableStyleCollection {
        /**
         * Creates a blank PivotTableStyle with the specified name.
         * @param name The unique name for the new PivotTableStyle. Will throw an invalid argument exception if the name is already in use.
         * @param makeUniqueName Optional, defaults to false. If true, will append numbers to the name in order to make it unique, if needed.
         */
        add(name: string, makeUniqueName?: boolean): PivotTableStyle;

        /**
         * Gets the number of PivotTable styles in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the default PivotTableStyle for the parent object's scope.
         */
        getDefault(): PivotTableStyle;

        /**
         * Gets a PivotTableStyle by name.
         * @param name Name of the PivotTableStyle to be retrieved.
         */
        getItem(name: string): PivotTableStyle;

        /**
         * Gets a PivotTableStyle by name. If the PivotTableStyle does not exist, will return a null object.
         * @param name Name of the PivotTableStyle to be retrieved.
         */
        getItemOrNullObject(name: string): PivotTableStyle;

        /**
         * Sets the default PivotTableStyle for use in the parent object's scope.
         * @param newDefaultStyle The PivotTableStyle object or name of the PivotTableStyle object that should be the new default.
         */
        setDefault(newDefaultStyle: PivotTableStyle | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a PivotTable Style, which defines style elements by PivotTable region.
     */
    interface PivotTableStyle {
        /**
         * Gets the name of the PivotTableStyle.
         */
        name: string;

        /**
         * Specifies if this PivotTableStyle object is read-only.
         */
        readonly readOnly: boolean;

        /**
         * Deletes the PivotTableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this PivotTableStyle with copies of all the style elements.
         */
        duplicate(): PivotTableStyle;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of SlicerStyle objects.
     */
    interface SlicerStyleCollection {
        /**
         * Creates a blank SlicerStyle with the specified name.
         * @param name The unique name for the new SlicerStyle. Will throw an invalid argument exception if the name is already in use.
         * @param makeUniqueName Optional, defaults to false. If true, will append numbers to the name in order to make it unique, if needed.
         */
        add(name: string, makeUniqueName?: boolean): SlicerStyle;

        /**
         * Gets the number of slicer styles in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the default SlicerStyle for the parent object's scope.
         */
        getDefault(): SlicerStyle;

        /**
         * Gets a SlicerStyle by name.
         * @param name Name of the SlicerStyle to be retrieved.
         */
        getItem(name: string): SlicerStyle;

        /**
         * Gets a SlicerStyle by name. If the SlicerStyle does not exist, will return a null object.
         * @param name Name of the SlicerStyle to be retrieved.
         */
        getItemOrNullObject(name: string): SlicerStyle;

        /**
         * Sets the default SlicerStyle for use in the parent object's scope.
         * @param newDefaultStyle The SlicerStyle object or name of the SlicerStyle object that should be the new default.
         */
        setDefault(newDefaultStyle: SlicerStyle | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a Slicer Style, which defines style elements by region of the slicer.
     */
    interface SlicerStyle {
        /**
         * Gets the name of the SlicerStyle.
         */
        name: string;

        /**
         * Specifies if this SlicerStyle object is read-only.
         */
        readonly readOnly: boolean;

        /**
         * Deletes the SlicerStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this SlicerStyle with copies of all the style elements.
         */
        duplicate(): SlicerStyle;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of TimelineStyles.
     */
    interface TimelineStyleCollection {
        /**
         * Creates a blank TimelineStyle with the specified name.
         * @param name The unique name for the new TimelineStyle. Will throw an invalid argument exception if the name is already in use.
         * @param makeUniqueName Optional, defaults to false. If true, will append numbers to the name in order to make it unique, if needed.
         */
        add(name: string, makeUniqueName?: boolean): TimelineStyle;

        /**
         * Gets the number of timeline styles in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets the default TimelineStyle for the parent object's scope.
         */
        getDefault(): TimelineStyle;

        /**
         * Gets a TimelineStyle by name.
         * @param name Name of the TimelineStyle to be retrieved.
         */
        getItem(name: string): TimelineStyle;

        /**
         * Gets a TimelineStyle by name. If the TimelineStyle does not exist, will return a null object.
         * @param name Name of the TimelineStyle to be retrieved.
         */
        getItemOrNullObject(name: string): TimelineStyle;

        /**
         * Sets the default TimelineStyle for use in the parent object's scope.
         * @param newDefaultStyle The TimelineStyle object or name of the TimelineStyle object that should be the new default.
         */
        setDefault(newDefaultStyle: TimelineStyle | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a Timeline style, which defines style elements by region in the Timeline.
     */
    interface TimelineStyle {
        /**
         * Gets the name of the TimelineStyle.
         */
        name: string;

        /**
         * Specifies if this TimelineStyle object is read-only.
         */
        readonly readOnly: boolean;

        /**
         * Deletes the TableStyle.
         */
        delete(): void;

        /**
         * Creates a duplicate of this TimelineStyle with copies of all the style elements.
         */
        duplicate(): TimelineStyle;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents layout and print settings that are not dependent any printer-specific implementation. These settings include margins, orientation, page numbering, title rows, and print area.
     */
    interface PageLayout {
        /**
         * The worksheet's black and white print option.
         */
        blackAndWhite: boolean;

        /**
         * The worksheet's bottom page margin to use for printing in points.
         */
        bottomMargin: number;

        /**
         * The worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.
         */
        centerHorizontally: boolean;

        /**
         * The worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.
         */
        centerVertically: boolean;

        /**
         * The worksheet's draft mode option. If true the sheet will be printed without graphics.
         */
        draftMode: boolean;

        /**
         * The worksheet's first page number to print. Null value represents "auto" page numbering.
         */
        firstPageNumber: number | "";

        /**
         * The worksheet's footer margin, in points, for use when printing.
         */
        footerMargin: number;

        /**
         * The worksheet's header margin, in points, for use when printing.
         */
        headerMargin: number;

        /**
         * Header and footer configuration for the worksheet.
         */
        readonly headersFooters: HeaderFooterGroup;

        /**
         * The worksheet's left margin, in points, for use when printing.
         */
        leftMargin: number;

        /**
         * The worksheet's orientation of the page.
         */
        orientation: PageOrientation | "Portrait" | "Landscape";

        /**
         * The worksheet's paper size of the page.
         */
        paperSize:
            | PaperType
            | "Letter"
            | "LetterSmall"
            | "Tabloid"
            | "Ledger"
            | "Legal"
            | "Statement"
            | "Executive"
            | "A3"
            | "A4"
            | "A4Small"
            | "A5"
            | "B4"
            | "B5"
            | "Folio"
            | "Quatro"
            | "Paper10x14"
            | "Paper11x17"
            | "Note"
            | "Envelope9"
            | "Envelope10"
            | "Envelope11"
            | "Envelope12"
            | "Envelope14"
            | "Csheet"
            | "Dsheet"
            | "Esheet"
            | "EnvelopeDL"
            | "EnvelopeC5"
            | "EnvelopeC3"
            | "EnvelopeC4"
            | "EnvelopeC6"
            | "EnvelopeC65"
            | "EnvelopeB4"
            | "EnvelopeB5"
            | "EnvelopeB6"
            | "EnvelopeItaly"
            | "EnvelopeMonarch"
            | "EnvelopePersonal"
            | "FanfoldUS"
            | "FanfoldStdGerman"
            | "FanfoldLegalGerman";

        /**
         * Specifies if the worksheet's comments should be displayed when printing.
         */
        printComments: PrintComments | "NoComments" | "EndSheet" | "InPlace";

        /**
         * The worksheet's print errors option.
         */
        printErrors:
            | PrintErrorType
            | "AsDisplayed"
            | "Blank"
            | "Dash"
            | "NotAvailable";

        /**
         * Specifies if the worksheet's gridlines will be printed.
         */
        printGridlines: boolean;

        /**
         * Specifies if the worksheet's headings will be printed.
         */
        printHeadings: boolean;

        /**
         * The worksheet's page print order option. This specifies the order to use for processing the page number printed.
         */
        printOrder: PrintOrder | "DownThenOver" | "OverThenDown";

        /**
         * The worksheet's right margin, in points, for use when printing.
         */
        rightMargin: number;

        /**
         * The worksheet's top margin, in points, for use when printing.
         */
        topMargin: number;

        /**
         * The worksheet's print zoom options.
         * The `PageLayoutZoomOptions` object must be set as a JSON object (use `x.zoom = {...}` instead of `x.zoom.scale = ...`).
         */
        zoom: PageLayoutZoomOptions;

        /**
         * Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, an ItemNotFound error will be thrown.
         */
        getPrintArea(): RangeAreas;

        /**
         * Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, a null object will be returned.
         */
        getPrintAreaOrNullObject(): RangeAreas;

        /**
         * Gets the range object representing the title columns.
         */
        getPrintTitleColumns(): Range;

        /**
         * Gets the range object representing the title columns. If not set, this will return a null object.
         */
        getPrintTitleColumnsOrNullObject(): Range;

        /**
         * Gets the range object representing the title rows.
         */
        getPrintTitleRows(): Range;

        /**
         * Gets the range object representing the title rows. If not set, this will return a null object.
         */
        getPrintTitleRowsOrNullObject(): Range;

        /**
         * Sets the worksheet's print area.
         * @param printArea The range, or RangeAreas of the content to print.
         */
        setPrintArea(printArea: Range | RangeAreas | string): void;

        /**
         * Sets the worksheet's page margins with units.
         * @param unit Measurement unit for the margins provided.
         * @param marginOptions Margin values to set, margins not provided will remain unchanged.
         */
        setPrintMargins(
            unit: PrintMarginUnit | "Points" | "Inches" | "Centimeters",
            marginOptions: PageLayoutMarginOptions
        ): void;

        /**
         * Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.
         * @param printTitleColumns The columns to be repeated to the left of each page, range must span the entire column to be valid.
         */
        setPrintTitleColumns(printTitleColumns: Range | string): void;

        /**
         * Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.
         * @param printTitleRows The rows to be repeated at the top of each page, range must span the entire row to be valid.
         */
        setPrintTitleRows(printTitleRows: Range | string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface HeaderFooter {
        /**
         * The center footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        centerFooter: string;

        /**
         * The center header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        centerHeader: string;

        /**
         * The left footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        leftFooter: string;

        /**
         * The left header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        leftHeader: string;

        /**
         * The right footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        rightFooter: string;

        /**
         * The right header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        rightHeader: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface HeaderFooterGroup {
        /**
         * The general header/footer, used for all pages unless even/odd or first page is specified.
         */
        readonly defaultForAllPages: HeaderFooter;

        /**
         * The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.
         */
        readonly evenPages: HeaderFooter;

        /**
         * The first page header/footer, for all other pages general or even/odd is used.
         */
        readonly firstPage: HeaderFooter;

        /**
         * The header/footer to use for odd pages, even header/footer needs to be specified for even pages.
         */
        readonly oddPages: HeaderFooter;

        /**
         * The state by which headers/footers are set. See Excel.HeaderFooterState for details.
         */
        state:
            | HeaderFooterState
            | "Default"
            | "FirstAndDefault"
            | "OddAndEven"
            | "FirstOddAndEven";

        /**
         * Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.
         */
        useSheetMargins: boolean;

        /**
         * Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.
         */
        useSheetScale: boolean;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface PageBreak {
        /**
         * Specifies the column index for the page break
         */
        readonly columnIndex: number;

        /**
         * Deletes a page break object.
         */
        delete(): void;

        /**
         * Gets the first cell after the page break.
         */
        getCellAfterBreak(): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface PageBreakCollection {
        /**
         * Adds a page break before the top-left cell of the range specified.
         * @param pageBreakRange The range immediately after the page break to be added.
         */
        add(pageBreakRange: Range | string): PageBreak;

        /**
         * Gets the number of page breaks in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a page break object via the index.
         * @param index Index of the page break.
         */
        getItem(index: number): PageBreak;

        /**
         * Resets all manual page breaks in the collection.
         */
        removePageBreaks(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the Data Connections that are part of the workbook or worksheet.
     */
    interface DataConnectionCollection {
        /**
         * Refreshes all the Data Connections in the collection.
         */
        refreshAll(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    interface RangeCollection {
        /**
         * Returns the number of ranges in the RangeCollection.
         */
        getCount(): ClientResult<number>;

        /**
         * Returns the range object based on its position in the RangeCollection.
         * @param index Index value of the range object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Range;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of comment objects that are part of the workbook.
     */
    interface CommentCollection {
        /**
         * Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param cellAddress The cell to which the comment is added. This can be a Range object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param content The comment's content. This can be either a string or CommentRichContent object. Strings are used for plain text. CommentRichContent objects allow for other comment features, such as mentions.
         * @param contentType Optional. The type of content contained within the comment. The default value is enum `ContentType.Plain`.
         */
        add(
            cellAddress: Range | string,
            content: CommentRichContent | string,
            contentType?: ContentType | "Plain" | "Mention"
        ): Comment;

        /**
         * Gets the number of comments in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a comment from the collection based on its ID.
         * @param commentId The identifier for the comment.
         */
        getItem(commentId: string): Comment;

        /**
         * Gets a comment from the collection based on its position.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Comment;

        /**
         * Gets the comment from the specified cell.
         * @param cellAddress The cell which the comment is on. This can be a Range object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         */
        getItemByCell(cellAddress: Range | string): Comment;

        /**
         * Gets the comment to which the given reply is connected.
         * @param replyId The identifier of comment reply.
         */
        getItemByReplyId(replyId: string): Comment;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a comment in the workbook.
     */
    interface Comment {
        /**
         * Gets the email of the comment's author.
         */
        readonly authorEmail: string;

        /**
         * Gets the name of the comment's author.
         */
        readonly authorName: string;

        /**
         * The comment's content. The string is plain text.
         */
        content: string;

        /**
         * Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.
         */
        readonly creationDate: Date;

        /**
         * Specifies the comment identifier.
         */
        readonly id: string;

        /**
         * Gets the entities (e.g., people) that are mentioned in comments.
         */
        readonly mentions: CommentMention[];

        /**
         * Represents a collection of reply objects associated with the comment.
         */
        readonly replies: CommentReplyCollection;

        /**
         * The comment thread status. A value of "true" means the comment thread is resolved.
         */
        resolved: boolean;

        /**
         * Gets the rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        readonly richContent: string;

        /**
         * Deletes the comment and all the connected replies.
         */
        delete(): void;

        /**
         * Gets the cell where this comment is located.
         */
        getLocation(): Range;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         * @param contentWithMentions The content for the comment. This contains a specially formatted string and a list of mentions that will be parsed into the string when displayed by Excel.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of comment reply objects that are part of the comment.
     */
    interface CommentReplyCollection {
        /**
         * Creates a comment reply for comment.
         * @param content The comment's content. This can be either a string or Interface CommentRichContent (e.g., for comments with mentions).
         * @param contentType Optional. The type of content contained within the comment. The default value is enum `ContentType.Plain`.
         */
        add(
            content: CommentRichContent | string,
            contentType?: ContentType | "Plain" | "Mention"
        ): CommentReply;

        /**
         * Gets the number of comment replies in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Returns a comment reply identified by its ID.
         * @param commentReplyId The identifier for the comment reply.
         */
        getItem(commentReplyId: string): CommentReply;

        /**
         * Gets a comment reply based on its position in the collection.
         * @param index The index value of the comment reply to be retrieved. The collection uses zero-based indexing.
         */
        getItemAt(index: number): CommentReply;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a comment reply in the workbook.
     */
    interface CommentReply {
        /**
         * Gets the email of the comment reply's author.
         */
        readonly authorEmail: string;

        /**
         * Gets the name of the comment reply's author.
         */
        readonly authorName: string;

        /**
         * The comment reply's content. The string is plain text.
         */
        content: string;

        /**
         * Gets the creation time of the comment reply.
         */
        readonly creationDate: Date;

        /**
         * Specifies the comment reply identifier.
         */
        readonly id: string;

        /**
         * The entities (e.g., people) that are mentioned in comments.
         */
        readonly mentions: CommentMention[];

        /**
         * The comment reply status. A value of "true" means the reply is in the resolved state.
         */
        readonly resolved: boolean;

        /**
         * The rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        readonly richContent: string;

        /**
         * Deletes the comment reply.
         */
        delete(): void;

        /**
         * Gets the cell where this comment reply is located.
         */
        getLocation(): Range;

        /**
         * Gets the parent comment of this reply.
         */
        getParentComment(): Comment;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         * @param contentWithMentions The content for the comment. This contains a specially formatted string and a list of mentions that will be parsed into the string when displayed by Excel.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the shapes in the worksheet.
     */
    interface ShapeCollection {
        /**
         * Adds a geometric shape to the worksheet. Returns a Shape object that represents the new shape.
         * @param geometricShapeType Represents the type of the geometric shape. See Excel.GeometricShapeType for details.
         */
        addGeometricShape(
            geometricShapeType:
                | GeometricShapeType
                | "LineInverse"
                | "Triangle"
                | "RightTriangle"
                | "Rectangle"
                | "Diamond"
                | "Parallelogram"
                | "Trapezoid"
                | "NonIsoscelesTrapezoid"
                | "Pentagon"
                | "Hexagon"
                | "Heptagon"
                | "Octagon"
                | "Decagon"
                | "Dodecagon"
                | "Star4"
                | "Star5"
                | "Star6"
                | "Star7"
                | "Star8"
                | "Star10"
                | "Star12"
                | "Star16"
                | "Star24"
                | "Star32"
                | "RoundRectangle"
                | "Round1Rectangle"
                | "Round2SameRectangle"
                | "Round2DiagonalRectangle"
                | "SnipRoundRectangle"
                | "Snip1Rectangle"
                | "Snip2SameRectangle"
                | "Snip2DiagonalRectangle"
                | "Plaque"
                | "Ellipse"
                | "Teardrop"
                | "HomePlate"
                | "Chevron"
                | "PieWedge"
                | "Pie"
                | "BlockArc"
                | "Donut"
                | "NoSmoking"
                | "RightArrow"
                | "LeftArrow"
                | "UpArrow"
                | "DownArrow"
                | "StripedRightArrow"
                | "NotchedRightArrow"
                | "BentUpArrow"
                | "LeftRightArrow"
                | "UpDownArrow"
                | "LeftUpArrow"
                | "LeftRightUpArrow"
                | "QuadArrow"
                | "LeftArrowCallout"
                | "RightArrowCallout"
                | "UpArrowCallout"
                | "DownArrowCallout"
                | "LeftRightArrowCallout"
                | "UpDownArrowCallout"
                | "QuadArrowCallout"
                | "BentArrow"
                | "UturnArrow"
                | "CircularArrow"
                | "LeftCircularArrow"
                | "LeftRightCircularArrow"
                | "CurvedRightArrow"
                | "CurvedLeftArrow"
                | "CurvedUpArrow"
                | "CurvedDownArrow"
                | "SwooshArrow"
                | "Cube"
                | "Can"
                | "LightningBolt"
                | "Heart"
                | "Sun"
                | "Moon"
                | "SmileyFace"
                | "IrregularSeal1"
                | "IrregularSeal2"
                | "FoldedCorner"
                | "Bevel"
                | "Frame"
                | "HalfFrame"
                | "Corner"
                | "DiagonalStripe"
                | "Chord"
                | "Arc"
                | "LeftBracket"
                | "RightBracket"
                | "LeftBrace"
                | "RightBrace"
                | "BracketPair"
                | "BracePair"
                | "Callout1"
                | "Callout2"
                | "Callout3"
                | "AccentCallout1"
                | "AccentCallout2"
                | "AccentCallout3"
                | "BorderCallout1"
                | "BorderCallout2"
                | "BorderCallout3"
                | "AccentBorderCallout1"
                | "AccentBorderCallout2"
                | "AccentBorderCallout3"
                | "WedgeRectCallout"
                | "WedgeRRectCallout"
                | "WedgeEllipseCallout"
                | "CloudCallout"
                | "Cloud"
                | "Ribbon"
                | "Ribbon2"
                | "EllipseRibbon"
                | "EllipseRibbon2"
                | "LeftRightRibbon"
                | "VerticalScroll"
                | "HorizontalScroll"
                | "Wave"
                | "DoubleWave"
                | "Plus"
                | "FlowChartProcess"
                | "FlowChartDecision"
                | "FlowChartInputOutput"
                | "FlowChartPredefinedProcess"
                | "FlowChartInternalStorage"
                | "FlowChartDocument"
                | "FlowChartMultidocument"
                | "FlowChartTerminator"
                | "FlowChartPreparation"
                | "FlowChartManualInput"
                | "FlowChartManualOperation"
                | "FlowChartConnector"
                | "FlowChartPunchedCard"
                | "FlowChartPunchedTape"
                | "FlowChartSummingJunction"
                | "FlowChartOr"
                | "FlowChartCollate"
                | "FlowChartSort"
                | "FlowChartExtract"
                | "FlowChartMerge"
                | "FlowChartOfflineStorage"
                | "FlowChartOnlineStorage"
                | "FlowChartMagneticTape"
                | "FlowChartMagneticDisk"
                | "FlowChartMagneticDrum"
                | "FlowChartDisplay"
                | "FlowChartDelay"
                | "FlowChartAlternateProcess"
                | "FlowChartOffpageConnector"
                | "ActionButtonBlank"
                | "ActionButtonHome"
                | "ActionButtonHelp"
                | "ActionButtonInformation"
                | "ActionButtonForwardNext"
                | "ActionButtonBackPrevious"
                | "ActionButtonEnd"
                | "ActionButtonBeginning"
                | "ActionButtonReturn"
                | "ActionButtonDocument"
                | "ActionButtonSound"
                | "ActionButtonMovie"
                | "Gear6"
                | "Gear9"
                | "Funnel"
                | "MathPlus"
                | "MathMinus"
                | "MathMultiply"
                | "MathDivide"
                | "MathEqual"
                | "MathNotEqual"
                | "CornerTabs"
                | "SquareTabs"
                | "PlaqueTabs"
                | "ChartX"
                | "ChartStar"
                | "ChartPlus"
        ): Shape;

        /**
         * Groups a subset of shapes in this collection's worksheet. Returns a Shape object that represents the new group of shapes.
         * @param values An array of shape ID or shape objects.
         */
        addGroup(values: Array<string | Shape>): Shape;

        /**
         * Creates an image from a base64-encoded string and adds it to the worksheet. Returns the Shape object that represents the new image.
         * @param base64ImageString A base64-encoded string representing an image in either JPEG or PNG format.
         */
        addImage(base64ImageString: string): Shape;

        /**
         * Adds a line to worksheet. Returns a Shape object that represents the new line.
         * @param startLeft The distance, in points, from the start of the line to the left side of the worksheet.
         * @param startTop The distance, in points, from the start of the line to the top of the worksheet.
         * @param endLeft The distance, in points, from the end of the line to the left of the worksheet.
         * @param endTop The distance, in points, from the end of the line to the top of the worksheet.
         * @param connectorType Represents the connector type. See Excel.ConnectorType for details.
         */
        addLine(
            startLeft: number,
            startTop: number,
            endLeft: number,
            endTop: number,
            connectorType?: ConnectorType | "Straight" | "Elbow" | "Curve"
        ): Shape;

        /**
         * Adds a text box to the worksheet with the provided text as the content. Returns a Shape object that represents the new text box.
         * @param text Represents the text that will be shown in the created text box.
         */
        addTextBox(text?: string): Shape;

        /**
         * Returns the number of shapes in the worksheet.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a shape using its Name or ID.
         * @param key Name or ID of the shape to be retrieved.
         */
        getItem(key: string): Shape;

        /**
         * Gets a shape using its position in the collection.
         * @param index The zero-based index of the shape to be retrieved.
         */
        getItemAt(index: number): Shape;

        /**
         * Returns a shape identified by the shape id.
         * @param shapeId The identifier for the shape.
         */
        _GetItem(shapeId: string): Shape;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a generic shape object in the worksheet. A shape could be a geometric shape, a line, a group of shapes, etc.
     * To learn more about the shape object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-shapes | Work with shapes using the Excel JavaScript API}.
     */
    interface Shape {
        /**
         * Specifies the alternative description text for a Shape object.
         */
        altTextDescription: string;

        /**
         * Specifies the alternative title text for a Shape object.
         */
        altTextTitle: string;

        /**
         * Returns the number of connection sites on this shape.
         */
        readonly connectionSiteCount: number;

        /**
         * Returns the fill formatting of this shape.
         */
        readonly fill: ShapeFill;

        /**
         * Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".
         */
        readonly geometricShape: GeometricShape;

        /**
         * Specifies the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".
         */
        geometricShapeType:
            | GeometricShapeType
            | "LineInverse"
            | "Triangle"
            | "RightTriangle"
            | "Rectangle"
            | "Diamond"
            | "Parallelogram"
            | "Trapezoid"
            | "NonIsoscelesTrapezoid"
            | "Pentagon"
            | "Hexagon"
            | "Heptagon"
            | "Octagon"
            | "Decagon"
            | "Dodecagon"
            | "Star4"
            | "Star5"
            | "Star6"
            | "Star7"
            | "Star8"
            | "Star10"
            | "Star12"
            | "Star16"
            | "Star24"
            | "Star32"
            | "RoundRectangle"
            | "Round1Rectangle"
            | "Round2SameRectangle"
            | "Round2DiagonalRectangle"
            | "SnipRoundRectangle"
            | "Snip1Rectangle"
            | "Snip2SameRectangle"
            | "Snip2DiagonalRectangle"
            | "Plaque"
            | "Ellipse"
            | "Teardrop"
            | "HomePlate"
            | "Chevron"
            | "PieWedge"
            | "Pie"
            | "BlockArc"
            | "Donut"
            | "NoSmoking"
            | "RightArrow"
            | "LeftArrow"
            | "UpArrow"
            | "DownArrow"
            | "StripedRightArrow"
            | "NotchedRightArrow"
            | "BentUpArrow"
            | "LeftRightArrow"
            | "UpDownArrow"
            | "LeftUpArrow"
            | "LeftRightUpArrow"
            | "QuadArrow"
            | "LeftArrowCallout"
            | "RightArrowCallout"
            | "UpArrowCallout"
            | "DownArrowCallout"
            | "LeftRightArrowCallout"
            | "UpDownArrowCallout"
            | "QuadArrowCallout"
            | "BentArrow"
            | "UturnArrow"
            | "CircularArrow"
            | "LeftCircularArrow"
            | "LeftRightCircularArrow"
            | "CurvedRightArrow"
            | "CurvedLeftArrow"
            | "CurvedUpArrow"
            | "CurvedDownArrow"
            | "SwooshArrow"
            | "Cube"
            | "Can"
            | "LightningBolt"
            | "Heart"
            | "Sun"
            | "Moon"
            | "SmileyFace"
            | "IrregularSeal1"
            | "IrregularSeal2"
            | "FoldedCorner"
            | "Bevel"
            | "Frame"
            | "HalfFrame"
            | "Corner"
            | "DiagonalStripe"
            | "Chord"
            | "Arc"
            | "LeftBracket"
            | "RightBracket"
            | "LeftBrace"
            | "RightBrace"
            | "BracketPair"
            | "BracePair"
            | "Callout1"
            | "Callout2"
            | "Callout3"
            | "AccentCallout1"
            | "AccentCallout2"
            | "AccentCallout3"
            | "BorderCallout1"
            | "BorderCallout2"
            | "BorderCallout3"
            | "AccentBorderCallout1"
            | "AccentBorderCallout2"
            | "AccentBorderCallout3"
            | "WedgeRectCallout"
            | "WedgeRRectCallout"
            | "WedgeEllipseCallout"
            | "CloudCallout"
            | "Cloud"
            | "Ribbon"
            | "Ribbon2"
            | "EllipseRibbon"
            | "EllipseRibbon2"
            | "LeftRightRibbon"
            | "VerticalScroll"
            | "HorizontalScroll"
            | "Wave"
            | "DoubleWave"
            | "Plus"
            | "FlowChartProcess"
            | "FlowChartDecision"
            | "FlowChartInputOutput"
            | "FlowChartPredefinedProcess"
            | "FlowChartInternalStorage"
            | "FlowChartDocument"
            | "FlowChartMultidocument"
            | "FlowChartTerminator"
            | "FlowChartPreparation"
            | "FlowChartManualInput"
            | "FlowChartManualOperation"
            | "FlowChartConnector"
            | "FlowChartPunchedCard"
            | "FlowChartPunchedTape"
            | "FlowChartSummingJunction"
            | "FlowChartOr"
            | "FlowChartCollate"
            | "FlowChartSort"
            | "FlowChartExtract"
            | "FlowChartMerge"
            | "FlowChartOfflineStorage"
            | "FlowChartOnlineStorage"
            | "FlowChartMagneticTape"
            | "FlowChartMagneticDisk"
            | "FlowChartMagneticDrum"
            | "FlowChartDisplay"
            | "FlowChartDelay"
            | "FlowChartAlternateProcess"
            | "FlowChartOffpageConnector"
            | "ActionButtonBlank"
            | "ActionButtonHome"
            | "ActionButtonHelp"
            | "ActionButtonInformation"
            | "ActionButtonForwardNext"
            | "ActionButtonBackPrevious"
            | "ActionButtonEnd"
            | "ActionButtonBeginning"
            | "ActionButtonReturn"
            | "ActionButtonDocument"
            | "ActionButtonSound"
            | "ActionButtonMovie"
            | "Gear6"
            | "Gear9"
            | "Funnel"
            | "MathPlus"
            | "MathMinus"
            | "MathMultiply"
            | "MathDivide"
            | "MathEqual"
            | "MathNotEqual"
            | "CornerTabs"
            | "SquareTabs"
            | "PlaqueTabs"
            | "ChartX"
            | "ChartStar"
            | "ChartPlus";

        /**
         * Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".
         */
        readonly group: ShapeGroup;

        /**
         * Specifies the height, in points, of the shape.
         * Throws an invalid argument exception when set with a negative value or zero as input.
         */
        height: number;

        /**
         * Specifies the shape identifier.
         */
        readonly id: string;

        /**
         * Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".
         */
        readonly image: Image;

        /**
         * The distance, in points, from the left side of the shape to the left side of the worksheet.
         * Throws an invalid argument exception when set with a negative value as input.
         */
        left: number;

        /**
         * Specifies the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.
         */
        readonly level: number;

        /**
         * Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".
         */
        readonly line: Line;

        /**
         * Returns the line formatting of this shape.
         */
        readonly lineFormat: ShapeLineFormat;

        /**
         * Specifies if the aspect ratio of this shape is locked.
         */
        lockAspectRatio: boolean;

        /**
         * Specifies the name of the shape.
         */
        name: string;

        /**
         * Specifies the parent group of this shape.
         */
        readonly parentGroup: Shape;

        /**
         * Represents how the object is attached to the cells below it.
         */
        placement: Placement | "TwoCell" | "OneCell" | "Absolute";

        /**
         * Specifies the rotation, in degrees, of the shape.
         */
        rotation: number;

        /**
         * Returns the text frame object of this shape. Read only.
         */
        readonly textFrame: TextFrame;

        /**
         * The distance, in points, from the top edge of the shape to the top edge of the worksheet.
         * Throws an invalid argument exception when set with a negative value as input.
         */
        top: number;

        /**
         * Returns the type of this shape. See Excel.ShapeType for details.
         */
        readonly type:
            | ShapeType
            | "Unsupported"
            | "Image"
            | "GeometricShape"
            | "Group"
            | "Line";

        /**
         * Specifies if the shape is visible.
         */
        visible: boolean;

        /**
         * Specifies the width, in points, of the shape.
         * Throws an invalid argument exception when set with a negative value or zero as input.
         */
        width: number;

        /**
         * Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.
         */
        readonly zOrderPosition: number;

        /**
         * Copies and pastes a Shape object.
         * The pasted shape is copied to the same pixel location as this shape.
         * @param destinationSheet The sheet to which the shape object will be pasted. The default value is the copied Shape's worksheet.
         */
        copyTo(destinationSheet?: Worksheet | string): Shape;

        /**
         * Removes the shape from the worksheet.
         */
        delete(): void;

        /**
         * Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, and `Excel.PictureFormat.GIF`.
         * @param format Specifies the format of the image.
         */
        getAsImage(
            format:
                | PictureFormat
                | "Unknown"
                | "Bmp"
                | "Jpeg"
                | "Gif"
                | "Png"
                | "Svg"
        ): ClientResult<string>;

        /**
         * Moves the shape horizontally by the specified number of points.
         * @param increment The increment, in points, the shape will be horizontally moved. A positive value moves the shape to the right and a negative value moves it to the left. If the sheet is right-to-left oriented, this is reversed: positive values will move the shape to the left and negative values will move it to the right.
         */
        incrementLeft(increment: number): void;

        /**
         * Rotates the shape clockwise around the z-axis by the specified number of degrees.
         * Use the `rotation` property to set the absolute rotation of the shape.
         * @param increment How many degrees the shape will be rotated. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.
         */
        incrementRotation(increment: number): void;

        /**
         * Moves the shape vertically by the specified number of points.
         * @param increment The increment, in points, the shape will be vertically moved. in points. A positive value moves the shape down and a negative value moves it up.
         */
        incrementTop(increment: number): void;

        /**
         * Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
         * @param scaleFactor Specifies the ratio between the height of the shape after you resize it and the current or original height.
         * @param scaleType Specifies whether the shape is scaled relative to its original or current size. The original size scaling option only works for images.
         * @param scaleFrom Optional. Specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.
         */
        scaleHeight(
            scaleFactor: number,
            scaleType: ShapeScaleType | "CurrentSize" | "OriginalSize",
            scaleFrom?:
                | ShapeScaleFrom
                | "ScaleFromTopLeft"
                | "ScaleFromMiddle"
                | "ScaleFromBottomRight"
        ): void;

        /**
         * Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.
         * @param scaleFactor Specifies the ratio between the width of the shape after you resize it and the current or original width.
         * @param scaleType Specifies whether the shape is scaled relative to its original or current size. The original size scaling option only works for images.
         * @param scaleFrom Optional. Specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.
         */
        scaleWidth(
            scaleFactor: number,
            scaleType: ShapeScaleType | "CurrentSize" | "OriginalSize",
            scaleFrom?:
                | ShapeScaleFrom
                | "ScaleFromTopLeft"
                | "ScaleFromMiddle"
                | "ScaleFromBottomRight"
        ): void;

        /**
         * Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.
         * @param position Where to move the shape in the z-order stack relative to the other shapes. See Excel.ShapeZOrder for details.
         */
        setZOrder(
            position:
                | ShapeZOrder
                | "BringToFront"
                | "BringForward"
                | "SendToBack"
                | "SendBackward"
        ): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a geometric shape inside a worksheet. A geometric shape can be a rectangle, block arrow, equation symbol, flowchart item, star, banner, callout, or any other basic shape in Excel.
     */
    interface GeometricShape {
        /**
         * Returns the shape identifier.
         */
        readonly id: string;

        /**
         * Returns the Shape object for the geometric shape.
         */
        readonly shape: Shape;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents an image in the worksheet. To get the corresponding Shape object, use Image.shape.
     */
    interface Image {
        /**
         * Specifies the shape identifier for the image object.
         */
        readonly id: string;

        /**
         * Returns the Shape object associated with the image.
         */
        readonly shape: Shape;

        /**
         * Returns the format of the image.
         */
        readonly format:
            | PictureFormat
            | "Unknown"
            | "Bmp"
            | "Jpeg"
            | "Gif"
            | "Png"
            | "Svg";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a shape group inside a worksheet. To get the corresponding Shape object, use `ShapeGroup.shape`.
     */
    interface ShapeGroup {
        /**
         * Specifies the shape identifier.
         */
        readonly id: string;

        /**
         * Returns the Shape object associated with the group.
         */
        readonly shape: Shape;

        /**
         * Returns the collection of Shape objects.
         */
        readonly shapes: GroupShapeCollection;

        /**
         * Ungroups any grouped shapes in the specified shape group.
         */
        ungroup(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the shape collection inside a shape group.
     */
    interface GroupShapeCollection {
        /**
         * Returns the number of shapes in the shape group.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a shape using its Name or ID.
         * @param key The Name or ID of the shape to be retrieved.
         */
        getItem(key: string): Shape;

        /**
         * Gets a shape based on its position in the collection.
         * @param index The zero-based index value of the object to be retrieved.
         */
        getItemAt(index: number): Shape;

        /**
         * Returns a shape identified by the shape id.
         * @param shapeId The identifier for the shape.
         */
        _GetItem(shapeId: string): Shape;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a line inside a worksheet. To get the corresponding Shape object, use `Line.shape`.
     */
    interface Line {
        /**
         * Represents the length of the arrowhead at the beginning of the specified line.
         */
        beginArrowheadLength: ArrowheadLength | "Short" | "Medium" | "Long";

        /**
         * Represents the style of the arrowhead at the beginning of the specified line.
         */
        beginArrowheadStyle:
            | ArrowheadStyle
            | "None"
            | "Triangle"
            | "Stealth"
            | "Diamond"
            | "Oval"
            | "Open";

        /**
         * Represents the width of the arrowhead at the beginning of the specified line.
         */
        beginArrowheadWidth: ArrowheadWidth | "Narrow" | "Medium" | "Wide";

        /**
         * Represents the shape to which the beginning of the specified line is attached.
         */
        readonly beginConnectedShape: Shape;

        /**
         * Represents the connection site to which the beginning of a connector is connected. Returns null when the beginning of the line is not attached to any shape.
         */
        readonly beginConnectedSite: number;

        /**
         * Represents the length of the arrowhead at the end of the specified line.
         */
        endArrowheadLength: ArrowheadLength | "Short" | "Medium" | "Long";

        /**
         * Represents the style of the arrowhead at the end of the specified line.
         */
        endArrowheadStyle:
            | ArrowheadStyle
            | "None"
            | "Triangle"
            | "Stealth"
            | "Diamond"
            | "Oval"
            | "Open";

        /**
         * Represents the width of the arrowhead at the end of the specified line.
         */
        endArrowheadWidth: ArrowheadWidth | "Narrow" | "Medium" | "Wide";

        /**
         * Represents the shape to which the end of the specified line is attached.
         */
        readonly endConnectedShape: Shape;

        /**
         * Represents the connection site to which the end of a connector is connected. Returns null when the end of the line is not attached to any shape.
         */
        readonly endConnectedSite: number;

        /**
         * Specifies the shape identifier.
         */
        readonly id: string;

        /**
         * Specifies if the beginning of the specified line is connected to a shape.
         */
        readonly isBeginConnected: boolean;

        /**
         * Specifies if the end of the specified line is connected to a shape.
         */
        readonly isEndConnected: boolean;

        /**
         * Returns the Shape object associated with the line.
         */
        readonly shape: Shape;

        /**
         * Represents the connector type for the line.
         */
        connectorType: ConnectorType | "Straight" | "Elbow" | "Curve";

        /**
         * Attaches the beginning of the specified connector to a specified shape.
         * @param shape The shape to connect.
         * @param connectionSite The connection site on the shape to which the beginning of the connector is attached. Must be an integer between 0 (inclusive) and the connection-site count of the specified shape (exclusive).
         */
        connectBeginShape(shape: Shape, connectionSite: number): void;

        /**
         * Attaches the end of the specified connector to a specified shape.
         * @param shape The shape to connect.
         * @param connectionSite The connection site on the shape to which the end of the connector is attached. Must be an integer between 0 (inclusive) and the connection-site count of the specified shape (exclusive).
         */
        connectEndShape(shape: Shape, connectionSite: number): void;

        /**
         * Detaches the beginning of the specified connector from a shape.
         */
        disconnectBeginShape(): void;

        /**
         * Detaches the end of the specified connector from a shape.
         */
        disconnectEndShape(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the fill formatting of a shape object.
     */
    interface ShapeFill {
        /**
         * Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")
         */
        foregroundColor: string;

        /**
         * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         */
        transparency: number;

        /**
         * Returns the fill type of the shape. See Excel.ShapeFillType for details.
         */
        readonly type:
            | ShapeFillType
            | "NoFill"
            | "Solid"
            | "Gradient"
            | "Pattern"
            | "PictureAndTexture"
            | "Mixed";

        /**
         * Clears the fill formatting of this shape.
         */
        clear(): void;

        /**
         * Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".
         * @param color A string that represents the fill color in HTML color format, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
     */
    interface ShapeLineFormat {
        /**
         * Represents the line color in HTML color format, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        color: string;

        /**
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.
         */
        dashStyle:
            | ShapeLineDashStyle
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "LongDash"
            | "LongDashDot"
            | "RoundDot"
            | "Solid"
            | "SquareDot"
            | "LongDashDotDot"
            | "SystemDash"
            | "SystemDot"
            | "SystemDashDot";

        /**
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.
         */
        style:
            | ShapeLineStyle
            | "Single"
            | "ThickBetweenThin"
            | "ThickThin"
            | "ThinThick"
            | "ThinThin";

        /**
         * Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
         */
        transparency: number;

        /**
         * Specifies if the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.
         */
        visible: boolean;

        /**
         * Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.
         */
        weight: number;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the text frame of a shape object.
     */
    interface TextFrame {
        /**
         * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         */
        autoSizeSetting:
            | ShapeAutoSize
            | "AutoSizeNone"
            | "AutoSizeTextToFitShape"
            | "AutoSizeShapeToFitText"
            | "AutoSizeMixed";

        /**
         * Represents the bottom margin, in points, of the text frame.
         */
        bottomMargin: number;

        /**
         * Specifies if the text frame contains text.
         */
        readonly hasText: boolean;

        /**
         * Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.
         */
        horizontalAlignment:
            | ShapeTextHorizontalAlignment
            | "Left"
            | "Center"
            | "Right"
            | "Justify"
            | "JustifyLow"
            | "Distributed"
            | "ThaiDistributed";

        /**
         * Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.
         */
        horizontalOverflow: ShapeTextHorizontalOverflow | "Overflow" | "Clip";

        /**
         * Represents the left margin, in points, of the text frame.
         */
        leftMargin: number;

        /**
         * Represents the angle to which the text is oriented for the text frame. See Excel.ShapeTextOrientation for details.
         */
        orientation:
            | ShapeTextOrientation
            | "Horizontal"
            | "Vertical"
            | "Vertical270"
            | "WordArtVertical"
            | "EastAsianVertical"
            | "MongolianVertical"
            | "WordArtVerticalRTL";

        /**
         * Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.
         */
        readingOrder: ShapeTextReadingOrder | "LeftToRight" | "RightToLeft";

        /**
         * Represents the right margin, in points, of the text frame.
         */
        rightMargin: number;

        /**
         * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See Excel.TextRange for details.
         */
        readonly textRange: TextRange;

        /**
         * Represents the top margin, in points, of the text frame.
         */
        topMargin: number;

        /**
         * Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.
         */
        verticalAlignment:
            | ShapeTextVerticalAlignment
            | "Top"
            | "Middle"
            | "Bottom"
            | "Justified"
            | "Distributed";

        /**
         * Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.
         */
        verticalOverflow:
            | ShapeTextVerticalOverflow
            | "Overflow"
            | "Ellipsis"
            | "Clip";

        /**
         * Deletes all the text in the text frame.
         */
        deleteText(): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
     */
    interface TextRange {
        /**
         * Returns a ShapeFont object that represents the font attributes for the text range.
         */
        readonly font: ShapeFont;

        /**
         * Represents the plain text content of the text range.
         */
        text: string;

        /**
         * Returns a TextRange object for the substring in the given range.
         * @param start The zero-based index of the first character to get from the text range.
         * @param length Optional. The number of characters to be returned in the new text range. If length is omitted, all the characters from start to the end of the text range's last paragraph will be returned.
         */
        getSubstring(start: number, length?: number): TextRange;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
     */
    interface ShapeFont {
        /**
         * Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.
         */
        bold: boolean;

        /**
         * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.
         */
        color: string;

        /**
         * Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.
         */
        italic: boolean;

        /**
         * Represents font name (e.g., "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
         */
        name: string;

        /**
         * Represents font size in points (e.g., 11). Returns null if the TextRange includes text fragments with different font sizes.
         */
        size: number;

        /**
         * Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.
         */
        underline:
            | ShapeFontUnderlineStyle
            | "None"
            | "Single"
            | "Double"
            | "Heavy"
            | "Dotted"
            | "DottedHeavy"
            | "Dash"
            | "DashHeavy"
            | "DashLong"
            | "DashLongHeavy"
            | "DotDash"
            | "DotDashHeavy"
            | "DotDotDash"
            | "DotDotDashHeavy"
            | "Wavy"
            | "WavyHeavy"
            | "WavyDouble";

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a slicer object in the workbook.
     */
    interface Slicer {
        /**
         * Represents the caption of slicer.
         */
        caption: string;

        /**
         * Represents the height, in points, of the slicer.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value or zero as input.
         */
        height: number;

        /**
         * Represents the unique id of slicer.
         */
        readonly id: string;

        /**
         * True if all filters currently applied on the slicer are cleared.
         */
        readonly isFilterCleared: boolean;

        /**
         * Represents the distance, in points, from the left side of the slicer to the left of the worksheet.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value as input.
         */
        left: number;

        /**
         * Represents the name of slicer.
         */
        name: string;

        /**
         * Represents the collection of SlicerItems that are part of the slicer.
         */
        readonly slicerItems: SlicerItemCollection;

        /**
         * Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".
         */
        sortBy: SlicerSortType | "DataSourceOrder" | "Ascending" | "Descending";

        /**
         * Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.
         */
        style: string;

        /**
         * Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value as input.
         */
        top: number;

        /**
         * Represents the width, in points, of the slicer.
         * Throws an "The argument is invalid or missing or has an incorrect format." exception when set with negative value or zero as input.
         */
        width: number;

        /**
         * Represents the worksheet containing the slicer.
         */
        readonly worksheet: Worksheet;

        /**
         * Clears all the filters currently applied on the slicer.
         */
        clearFilters(): void;

        /**
         * Deletes the slicer.
         */
        delete(): void;

        /**
         * Returns an array of selected items' keys.
         */
        getSelectedItems(): ClientResult<string[]>;

        /**
         * Selects slicer items based on their keys. The previous selections are cleared.
         * All items will be selected by default if the array is empty.
         * @param items Optional. The specified slicer item names to be selected.
         */
        selectItems(items?: string[]): void;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the slicer objects on the workbook or a worksheet.
     */
    interface SlicerCollection {
        /**
         * Adds a new slicer to the workbook.
         * @param slicerSource The data source that the new slicer will be based on. It can be a PivotTable object, a Table object or a string. When a PivotTable object is passed, the data source is the source of the PivotTable object. When a Table object is passed, the data source is the Table object. When a string is passed, it is interpreted as the name/id of a PivotTable/Table.
         * @param sourceField The field in the data source to filter by. It can be a PivotField object, a TableColumn object, the id of a PivotField or the id/name of TableColumn.
         * @param slicerDestination Optional. The worksheet where the new slicer will be created in. It can be a Worksheet object or the name/id of a worksheet. This parameter can be omitted if the slicer collection is retrieved from worksheet.
         */
        add(
            slicerSource: string | PivotTable | Table,
            sourceField: string | PivotField | number | TableColumn,
            slicerDestination?: string | Worksheet
        ): Slicer;

        /**
         * Returns the number of slicers in the collection.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a slicer object using its name or id.
         * @param key The name or id of the slicer.
         */
        getItem(key: string): Slicer;

        /**
         * Gets a slicer based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Slicer;

        /**
         * Gets a slicer using its name or id. If the slicer does not exist, will return a null object.
         * @param key Name or Id of the slicer to be retrieved.
         */
        getItemOrNullObject(key: string): Slicer;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a slicer item in a slicer.
     */
    interface SlicerItem {
        /**
         * True if the slicer item has data.
         */
        readonly hasData: boolean;

        /**
         * True if the slicer item is selected.
         * Setting this value will not clear other SlicerItems' selected state.
         * By default, if the slicer item is the only one selected, when it is deselected, all items will be selected.
         */
        isSelected: boolean;

        /**
         * Represents the unique value representing the slicer item.
         */
        readonly key: string;

        /**
         * Represents the title displayed in the UI.
         */
        readonly name: string;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    /**
     * Represents a collection of all the slicer item objects on the slicer.
     */
    interface SlicerItemCollection {
        /**
         * Returns the number of slicer items in the slicer.
         */
        getCount(): ClientResult<number>;

        /**
         * Gets a slicer item object using its key or name.
         * @param key The key or name of the slicer item.
         */
        getItem(key: string): SlicerItem;

        /**
         * Gets a slicer item based on its position in the collection.
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): SlicerItem;

        /**
         * Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.
         * @param key Key or name of the slicer to be retrieved.
         */
        getItemOrNullObject(key: string): SlicerItem;

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): void;
    }

    //
    // Interface
    //

    /**
     * Represents the options in sheet protection.
     */
    interface WorksheetProtectionOptions {
        /**
         * Represents the worksheet protection option of allowing using auto filter feature.
         */
        allowAutoFilter?: boolean;
        /**
         * Represents the worksheet protection option of allowing deleting columns.
         */
        allowDeleteColumns?: boolean;
        /**
         * Represents the worksheet protection option of allowing deleting rows.
         */
        allowDeleteRows?: boolean;
        /**
         * Represents the worksheet protection option of allowing editing objects.
         */
        allowEditObjects?: boolean;
        /**
         * Represents the worksheet protection option of allowing editing scenarios.
         */
        allowEditScenarios?: boolean;
        /**
         * Represents the worksheet protection option of allowing formatting cells.
         */
        allowFormatCells?: boolean;
        /**
         * Represents the worksheet protection option of allowing formatting columns.
         */
        allowFormatColumns?: boolean;
        /**
         * Represents the worksheet protection option of allowing formatting rows.
         */
        allowFormatRows?: boolean;
        /**
         * Represents the worksheet protection option of allowing inserting columns.
         */
        allowInsertColumns?: boolean;
        /**
         * Represents the worksheet protection option of allowing inserting hyperlinks.
         */
        allowInsertHyperlinks?: boolean;
        /**
         * Represents the worksheet protection option of allowing inserting rows.
         */
        allowInsertRows?: boolean;
        /**
         * Represents the worksheet protection option of allowing using PivotTable feature.
         */
        allowPivotTables?: boolean;
        /**
         * Represents the worksheet protection option of allowing using sort feature.
         */
        allowSort?: boolean;
        /**
         * Represents the worksheet protection option of selection mode.
         */
        selectionMode?:
            | ProtectionSelectionMode
            | "Normal"
            | "Unlocked"
            | "None";
    }
    /**
     * Represents a string reference of the form SheetName!A1:B5, or a global or local named range.
     */
    interface RangeReference {
        /**
         * The address of the range; for example 'SheetName!A1:B5'.
         */
        address: string;
    }
    /**
     * Represents the necessary strings to get/set a hyperlink (XHL) object.
     */
    interface RangeHyperlink {
        /**
         * Represents the url target for the hyperlink.
         */
        address?: string;
        /**
         * Represents the document reference target for the hyperlink.
         */
        documentReference?: string;
        /**
         * Represents the string displayed when hovering over the hyperlink.
         */
        screenTip?: string;
        /**
         * Represents the string that is displayed in the top left most cell in the range.
         */
        textToDisplay?: string;
    }
    /**
     * Represents the search criteria to be used.
     */
    interface SearchCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is false (partial).
         */
        completeMatch?: boolean;
        /**
         * Specifies if the match is case sensitive. Default is false (insensitive).
         */
        matchCase?: boolean;
        /**
         * Specifies the search direction. Default is forward. See Excel.SearchDirection.
         */
        searchDirection?: SearchDirection | "Forward" | "Backwards";
    }
    /**
     * Represents the worksheet search criteria to be used.
     */
    interface WorksheetSearchCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is false (partial).
         */
        completeMatch?: boolean;
        /**
         * Specifies if the match is case sensitive. Default is false (insensitive).
         */
        matchCase?: boolean;
    }
    /**
     * Represents the replace criteria to be used.
     */
    interface ReplaceCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is false (partial).
         */
        completeMatch?: boolean;
        /**
         * Specifies if the match is case sensitive. Default is false (insensitive).
         */
        matchCase?: boolean;
    }
    /**
     * Specifies which properties to load on the `format.fill` object.
     */
    interface CellPropertiesFillLoadOptions {
        /**
         * Specifies whether to load the `color` property.
         */
        color?: boolean;
        /**
         * Specifies whether to load the `pattern` property.
         */
        pattern?: boolean;
        /**
         * Specifies whether to load the `patternColor` property.
         */
        patternColor?: boolean;
        /**
         * Specifies whether to load the `patternTintAndShade` property.
         */
        patternTintAndShade?: boolean;
        /**
         * Specifies whether to load the `tintAndShade` property.
         */
        tintAndShade?: boolean;
    }
    /**
     * Specifies which properties to load on the `format.font` object.
     */
    interface CellPropertiesFontLoadOptions {
        /**
         * Specifies whether to load on the `bold` property.
         */
        bold?: boolean;
        /**
         * Specifies whether to load on the `color` property.
         */
        color?: boolean;
        /**
         * Specifies whether to load on the `italic` property.
         */
        italic?: boolean;
        /**
         * Specifies whether to load on the `name` property.
         */
        name?: boolean;
        /**
         * Specifies whether to load on the `size` property.
         */
        size?: boolean;
        /**
         * Specifies whether to load on the `strikethrough` property.
         */
        strikethrough?: boolean;
        /**
         * Specifies whether to load on the `subscript` property.
         */
        subscript?: boolean;
        /**
         * Specifies whether to load on the `superscript` property.
         */
        superscript?: boolean;
        /**
         * Specifies whether to load on the `tintAndShade` property.
         */
        tintAndShade?: boolean;
        /**
         * Specifies whether to load on the `underline` property.
         */
        underline?: boolean;
    }
    /**
     * Specifies which properties to load on the `format.borders` object.
     */
    interface CellPropertiesBorderLoadOptions {
        /**
         * Specifies whether to load on the `color` property.
         */
        color?: boolean;
        /**
         * Specifies whether to load on the `style` property.
         */
        style?: boolean;
        /**
         * Specifies whether to load on the `tintAndShade` property.
         */
        tintAndShade?: boolean;
        /**
         * Specifies whether to load on the `weight` property.
         */
        weight?: boolean;
    }
    /**
     * Represents the `format.protection` properties of `getCellProperties`, `getRowProperties`, and `getColumnProperties` or the `format.protection` input parameter of `setCellProperties`, `setRowProperties`, and `setColumnProperties`.
     */
    interface CellPropertiesProtection {
        /**
         * Represents the `format.protection.formulaHidden` property.
         */
        formulaHidden?: boolean;
        /**
         * Represents the `format.protection.locked` property.
         */
        locked?: boolean;
    }
    /**
     * Represents the `format.fill` properties of `getCellProperties`, `getRowProperties`, and `getColumnProperties` or the `format.fill` input parameter of `setCellProperties`, `setRowProperties`, and `setColumnProperties`.
     */
    interface CellPropertiesFill {
        /**
         * Represents the `format.fill.color` property.
         */
        color?: string;
        /**
         * Represents the `format.fill.pattern` property.
         */
        pattern?:
            | FillPattern
            | "None"
            | "Solid"
            | "Gray50"
            | "Gray75"
            | "Gray25"
            | "Horizontal"
            | "Vertical"
            | "Down"
            | "Up"
            | "Checker"
            | "SemiGray75"
            | "LightHorizontal"
            | "LightVertical"
            | "LightDown"
            | "LightUp"
            | "Grid"
            | "CrissCross"
            | "Gray16"
            | "Gray8"
            | "LinearGradient"
            | "RectangularGradient";
        /**
         * Represents the `format.fill.patternColor` property.
         */
        patternColor?: string;
        /**
         * Represents the `format.fill.patternTintAndShade` property.
         */
        patternTintAndShade?: number;
        /**
         * Represents the `format.fill.tintAndShade` property.
         */
        tintAndShade?: number;
    }
    /**
     * Represents the `format.font` properties of `getCellProperties`, `getRowProperties`, and `getColumnProperties` or the `format.font` input parameter of `setCellProperties`, `setRowProperties`, and `setColumnProperties`.
     */
    interface CellPropertiesFont {
        /**
         * Represents the `format.font.bold` property.
         */
        bold?: boolean;
        /**
         * Represents the `format.font.color` property.
         */
        color?: string;
        /**
         * Represents the `format.font.italic` property.
         */
        italic?: boolean;
        /**
         * Represents the `format.font.name` property.
         */
        name?: string;
        /**
         * Represents the `format.font.size` property.
         */
        size?: number;
        /**
         * Represents the `format.font.strikethrough` property.
         */
        strikethrough?: boolean;
        /**
         * Represents the `format.font.subscript` property.
         */
        subscript?: boolean;
        /**
         * Represents the `format.font.superscript` property.
         */
        superscript?: boolean;
        /**
         * Represents the `format.font.tintAndShade` property.
         */
        tintAndShade?: number;
        /**
         * Represents the `format.font.underline` property.
         */
        underline?:
            | RangeUnderlineStyle
            | "None"
            | "Single"
            | "Double"
            | "SingleAccountant"
            | "DoubleAccountant";
    }
    /**
     * Represents the `format.borders` properties of `getCellProperties`, `getRowProperties`, and `getColumnProperties` or the `format.borders` input parameter of `setCellProperties`, `setRowProperties`, and `setColumnProperties`.
     */
    interface CellBorderCollection {
        /**
         * Represents the `format.borders.bottom` property.
         */
        bottom?: CellBorder;
        /**
         * Represents the `format.borders.diagonalDown` property.
         */
        diagonalDown?: CellBorder;
        /**
         * Represents the `format.borders.diagonalUp` property.
         */
        diagonalUp?: CellBorder;
        /**
         * Represents the `format.borders.horizontal` property.
         */
        horizontal?: CellBorder;
        /**
         * Represents the `format.borders.left` property.
         */
        left?: CellBorder;
        /**
         * Represents the `format.borders.right` property.
         */
        right?: CellBorder;
        /**
         * Represents the `format.borders.top` property.
         */
        top?: CellBorder;
        /**
         * Represents the `format.borders.vertical` property.
         */
        vertical?: CellBorder;
    }
    /**
     * Represents the properties of a single border returned by `getCellProperties`, `getRowProperties`, and `getColumnProperties` or the border property input parameter of `setCellProperties`, `setRowProperties`, and `setColumnProperties`.
     */
    interface CellBorder {
        /**
         * Represents the `color` property of a single border.
         */
        color?: string;
        /**
         * Represents the `style` property of a single border.
         */
        style?:
            | BorderLineStyle
            | "None"
            | "Continuous"
            | "Dash"
            | "DashDot"
            | "DashDotDot"
            | "Dot"
            | "Double"
            | "SlantDashDot";
        /**
         * Represents the `tintAndShade` property of a single border.
         */
        tintAndShade?: number;
        /**
         * Represents the `weight` property of a single border.
         */
        weight?: BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";
    }
    /**
     * Data validation rule contains different types of data validation. You can only use one of them at a time according the Excel.DataValidationType.
     */
    interface DataValidationRule {
        /**
         * Custom data validation criteria.
         */
        custom?: CustomDataValidation;
        /**
         * Date data validation criteria.
         */
        date?: DateTimeDataValidation;
        /**
         * Decimal data validation criteria.
         */
        decimal?: BasicDataValidation;
        /**
         * List data validation criteria.
         */
        list?: ListDataValidation;
        /**
         * TextLength data validation criteria.
         */
        textLength?: BasicDataValidation;
        /**
         * Time data validation criteria.
         */
        time?: DateTimeDataValidation;
        /**
         * WholeNumber data validation criteria.
         */
        wholeNumber?: BasicDataValidation;
    }
    /**
     * Represents the Basic Type data validation criteria.
     */
    interface BasicDataValidation {
        /**
         * Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.
         * For example, setting formula1 to 10 and operator to GreaterThan means that valid data for the range must be greater than 10.
         * When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula1: string | number | Range;
        /**
         * With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.
         * When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula2?: string | number | Range;
        /**
         * The operator to use for validating the data.
         */
        operator:
            | DataValidationOperator
            | "Between"
            | "NotBetween"
            | "EqualTo"
            | "NotEqualTo"
            | "GreaterThan"
            | "LessThan"
            | "GreaterThanOrEqualTo"
            | "LessThanOrEqualTo";
    }
    /**
     * Represents the Date data validation criteria.
     */
    interface DateTimeDataValidation {
        /**
         * Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.
         * When setting the value, it can be passed in as a Date, a Range object, or a string formula (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula1: string | Date | Range;
        /**
         * With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.
         * When setting the value, it can be passed in as a Date, a Range object, or a string (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula2?: string | Date | Range;
        /**
         * The operator to use for validating the data.
         */
        operator:
            | DataValidationOperator
            | "Between"
            | "NotBetween"
            | "EqualTo"
            | "NotEqualTo"
            | "GreaterThan"
            | "LessThan"
            | "GreaterThanOrEqualTo"
            | "LessThanOrEqualTo";
    }
    /**
     * Represents the List data validation criteria.
     */
    interface ListDataValidation {
        /**
         * Displays the list in cell drop down or not, it defaults to true.
         */
        inCellDropDown: boolean;
        /**
         * Source of the list for data validation
         * When setting the value, it can be passed in as a Excel Range object, or a string that contains comma separated number, boolean or date.
         */
        source: string | Range;
    }
    /**
     * Represents the Custom data validation criteria.
     */
    interface CustomDataValidation {
        /**
         * A custom data validation formula. This creates special input rules, such as preventing duplicates, or limiting the total in a range of cells.
         */
        formula: string;
    }
    /**
     * Represents the error alert properties for the data validation.
     */
    interface DataValidationErrorAlert {
        /**
         * Represents error alert message.
         */
        message: string;
        /**
         * Specifies whether to show an error alert dialog when a user enters invalid data. The default is true.
         */
        showAlert: boolean;
        /**
         * The data validation alert type, please see Excel.DataValidationAlertStyle for details.
         */
        style: DataValidationAlertStyle | "Stop" | "Warning" | "Information";
        /**
         * Represents error alert dialog title.
         */
        title: string;
    }
    /**
     * Represents the user prompt properties for the data validation.
     */
    interface DataValidationPrompt {
        /**
         * Specifies the message of the prompt.
         */
        message: string;
        /**
         * Specifies if a prompt is shown when a user selects a cell with data validation.
         */
        showPrompt: boolean;
        /**
         * Specifies the title for the prompt.
         */
        title: string;
    }
    /**
     * Represents a condition in a sorting operation.
     */
    interface SortField {
        /**
         * Specifies if the sorting is done in an ascending fashion.
         */
        ascending?: boolean;
        /**
         * Specifies the color that is the target of the condition if the sorting is on font or cell color.
         */
        color?: string;
        /**
         * Represents additional sorting options for this field.
         */
        dataOption?: SortDataOption | "Normal" | "TextAsNumber";
        /**
         * Specifies the icon that is the target of the condition if the sorting is on the cell's icon.
         */
        icon?: Icon;
        /**
         * Specifies the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
         */
        key: number;
        /**
         * Specifies the type of sorting of this condition.
         */
        sortOn?: SortOn | "Value" | "CellColor" | "FontColor" | "Icon";
        /**
         * Specifies the subfield that is the target property name of a rich value to sort on.
         */
        subField?: string;
    }
    /**
     * Represents the filtering criteria applied to a column.
     */
    interface FilterCriteria {
        /**
         * The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.
         */
        color?: string;
        /**
         * The first criterion used to filter data. Used as an operator in the case of "custom" filtering.
         * For example ">50" for number greater than 50 or "=*s" for values ending in "s".
         *
         * Used as a number in the case of top/bottom items/percents (e.g., "5" for the top 5 items if filterOn is set to "topItems").
         */
        criterion1?: string;
        /**
         * The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.
         */
        criterion2?: string;
        /**
         * The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering.
         */
        dynamicCriteria?:
            | DynamicFilterCriteria
            | "Unknown"
            | "AboveAverage"
            | "AllDatesInPeriodApril"
            | "AllDatesInPeriodAugust"
            | "AllDatesInPeriodDecember"
            | "AllDatesInPeriodFebruray"
            | "AllDatesInPeriodJanuary"
            | "AllDatesInPeriodJuly"
            | "AllDatesInPeriodJune"
            | "AllDatesInPeriodMarch"
            | "AllDatesInPeriodMay"
            | "AllDatesInPeriodNovember"
            | "AllDatesInPeriodOctober"
            | "AllDatesInPeriodQuarter1"
            | "AllDatesInPeriodQuarter2"
            | "AllDatesInPeriodQuarter3"
            | "AllDatesInPeriodQuarter4"
            | "AllDatesInPeriodSeptember"
            | "BelowAverage"
            | "LastMonth"
            | "LastQuarter"
            | "LastWeek"
            | "LastYear"
            | "NextMonth"
            | "NextQuarter"
            | "NextWeek"
            | "NextYear"
            | "ThisMonth"
            | "ThisQuarter"
            | "ThisWeek"
            | "ThisYear"
            | "Today"
            | "Tomorrow"
            | "YearToDate"
            | "Yesterday";
        /**
         * The property used by the filter to determine whether the values should stay visible.
         */
        filterOn:
            | FilterOn
            | "BottomItems"
            | "BottomPercent"
            | "CellColor"
            | "Dynamic"
            | "FontColor"
            | "Values"
            | "TopItems"
            | "TopPercent"
            | "Icon"
            | "Custom";
        /**
         * The icon used to filter cells. Used with "icon" filtering.
         */
        icon?: Icon;
        /**
         * The operator used to combine criterion 1 and 2 when using "custom" filtering.
         */
        operator?: FilterOperator | "And" | "Or";
        /**
         * The property used by the filter to do rich filter on richvalues.
         */
        subField?: string;
        /**
         * The set of values to be used as part of "values" filtering.
         */
        values?: Array<string | FilterDatetime>;
    }
    /**
     * Represents how to filter a date when filtering on values.
     */
    interface FilterDatetime {
        /**
         * The date in ISO8601 format used to filter data.
         */
        date: string;
        /**
         * How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009.
         */
        specificity:
            | FilterDatetimeSpecificity
            | "Year"
            | "Month"
            | "Day"
            | "Hour"
            | "Minute"
            | "Second";
    }
    /**
     * Represents a cell icon.
     */
    interface Icon {
        /**
         * Specifies the index of the icon in the given set.
         */
        index: number;
        /**
         * Specifies the set that the icon is part of.
         */
        set:
            | IconSet
            | "Invalid"
            | "ThreeArrows"
            | "ThreeArrowsGray"
            | "ThreeFlags"
            | "ThreeTrafficLights1"
            | "ThreeTrafficLights2"
            | "ThreeSigns"
            | "ThreeSymbols"
            | "ThreeSymbols2"
            | "FourArrows"
            | "FourArrowsGray"
            | "FourRedToBlack"
            | "FourRating"
            | "FourTrafficLights"
            | "FiveArrows"
            | "FiveArrowsGray"
            | "FiveRating"
            | "FiveQuarters"
            | "ThreeStars"
            | "ThreeTriangles"
            | "FiveBoxes";
    }
    interface ShowAsRule {
        /**
         * The base PivotField to base the ShowAs calculation, if applicable based on the ShowAsCalculation type, else null.
         */
        baseField?: PivotField;
        /**
         * The base Item to base the ShowAs calculation on, if applicable based on the ShowAsCalculation type, else null.
         */
        baseItem?: PivotItem;
        /**
         * The ShowAs Calculation to use for the Data PivotField. See Excel.ShowAsCalculation for Details.
         */
        calculation:
            | ShowAsCalculation
            | "Unknown"
            | "None"
            | "PercentOfGrandTotal"
            | "PercentOfRowTotal"
            | "PercentOfColumnTotal"
            | "PercentOfParentRowTotal"
            | "PercentOfParentColumnTotal"
            | "PercentOfParentTotal"
            | "PercentOf"
            | "RunningTotal"
            | "PercentRunningTotal"
            | "DifferenceFrom"
            | "PercentDifferenceFrom"
            | "RankAscending"
            | "RankDecending"
            | "Index";
    }
    /**
     * Subtotals for the Pivot Field.
     */
    interface Subtotals {
        /**
         * If Automatic is set to true, then all other values will be ignored when setting the Subtotals.
         */
        automatic?: boolean;
        /**
         * Average
         */
        average?: boolean;
        /**
         * Count
         */
        count?: boolean;
        /**
         * CountNumbers
         */
        countNumbers?: boolean;
        /**
         * Max
         */
        max?: boolean;
        /**
         * Min
         */
        min?: boolean;
        /**
         * Product
         */
        product?: boolean;
        /**
         * StandardDeviation
         */
        standardDeviation?: boolean;
        /**
         * StandardDeviationP
         */
        standardDeviationP?: boolean;
        /**
         * Sum
         */
        sum?: boolean;
        /**
         * Variance
         */
        variance?: boolean;
        /**
         * VarianceP
         */
        varianceP?: boolean;
    }
    /**
     * Represents a rule-type for a Data Bar.
     */
    interface ConditionalDataBarRule {
        /**
         * The formula, if required, to evaluate the databar rule on.
         */
        formula?: string;
        /**
         * The type of rule for the databar.
         */
        type:
            | ConditionalFormatRuleType
            | "Invalid"
            | "Automatic"
            | "LowestValue"
            | "HighestValue"
            | "Number"
            | "Percent"
            | "Formula"
            | "Percentile";
    }
    /**
     * Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.
     */
    interface ConditionalIconCriterion {
        /**
         * The custom icon for the current criterion if different from the default IconSet, else null will be returned.
         */
        customIcon?: Icon;
        /**
         * A number or a formula depending on the type.
         */
        formula: string;
        /**
         * GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.
         */
        operator:
            | ConditionalIconCriterionOperator
            | "Invalid"
            | "GreaterThan"
            | "GreaterThanOrEqual";
        /**
         * What the icon conditional formula should be based on.
         */
        type:
            | ConditionalFormatIconRuleType
            | "Invalid"
            | "Number"
            | "Percent"
            | "Formula"
            | "Percentile";
    }
    /**
     * Represents the criteria of the color scale.
     */
    interface ConditionalColorScaleCriteria {
        /**
         * The maximum point Color Scale Criterion.
         */
        maximum: ConditionalColorScaleCriterion;
        /**
         * The midpoint Color Scale Criterion if the color scale is a 3-color scale.
         */
        midpoint?: ConditionalColorScaleCriterion;
        /**
         * The minimum point Color Scale Criterion.
         */
        minimum: ConditionalColorScaleCriterion;
    }
    /**
     * Represents a Color Scale Criterion which contains a type, value, and a color.
     */
    interface ConditionalColorScaleCriterion {
        /**
         * HTML color code representation of the color scale color (e.g., #FF0000 represents Red).
         */
        color?: string;
        /**
         * A number, a formula, or null (if Type is LowestValue).
         */
        formula?: string;
        /**
         * What the criterion conditional formula should be based on.
         */
        type:
            | ConditionalFormatColorCriterionType
            | "Invalid"
            | "LowestValue"
            | "HighestValue"
            | "Number"
            | "Percent"
            | "Formula"
            | "Percentile";
    }
    /**
     * Represents the rule of the top/bottom conditional format.
     */
    interface ConditionalTopBottomRule {
        /**
         * The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.
         */
        rank: number;
        /**
         * Format values based on the top or bottom rank.
         */
        type:
            | ConditionalTopBottomCriterionType
            | "Invalid"
            | "TopItems"
            | "TopPercent"
            | "BottomItems"
            | "BottomPercent";
    }
    /**
     * Represents the Preset Criteria Conditional Format Rule
     */
    interface ConditionalPresetCriteriaRule {
        /**
         * The criterion of the conditional format.
         */
        criterion:
            | ConditionalFormatPresetCriterion
            | "Invalid"
            | "Blanks"
            | "NonBlanks"
            | "Errors"
            | "NonErrors"
            | "Yesterday"
            | "Today"
            | "Tomorrow"
            | "LastSevenDays"
            | "LastWeek"
            | "ThisWeek"
            | "NextWeek"
            | "LastMonth"
            | "ThisMonth"
            | "NextMonth"
            | "AboveAverage"
            | "BelowAverage"
            | "EqualOrAboveAverage"
            | "EqualOrBelowAverage"
            | "OneStdDevAboveAverage"
            | "OneStdDevBelowAverage"
            | "TwoStdDevAboveAverage"
            | "TwoStdDevBelowAverage"
            | "ThreeStdDevAboveAverage"
            | "ThreeStdDevBelowAverage"
            | "UniqueValues"
            | "DuplicateValues";
    }
    /**
     * Represents a Cell Value Conditional Format Rule
     */
    interface ConditionalTextComparisonRule {
        /**
         * The operator of the text conditional format.
         */
        operator:
            | ConditionalTextOperator
            | "Invalid"
            | "Contains"
            | "NotContains"
            | "BeginsWith"
            | "EndsWith";
        /**
         * The Text value of conditional format.
         */
        text: string;
    }
    /**
     * Represents a cell value conditional format rule.
     */
    interface ConditionalCellValueRule {
        /**
         * The formula, if required, to evaluate the conditional format rule on.
         */
        formula1: string;
        /**
         * The formula, if required, to evaluate the conditional format rule on.
         */
        formula2?: string;
        /**
         * The operator of the cell value conditional format.
         */
        operator:
            | ConditionalCellValueOperator
            | "Invalid"
            | "Between"
            | "NotBetween"
            | "EqualTo"
            | "NotEqualTo"
            | "GreaterThan"
            | "LessThan"
            | "GreaterThanOrEqual"
            | "LessThanOrEqual";
    }
    /**
     * Represents page zoom properties.
     */
    interface PageLayoutZoomOptions {
        /**
         * Number of pages to fit horizontally. This value can be null if percentage scale is used.
         */
        horizontalFitToPages?: number;
        /**
         * Print page scale value can be between 10 and 400. This value can be null if fit to page tall or wide is specified.
         */
        scale?: number;
        /**
         * Number of pages to fit vertically. This value can be null if percentage scale is used.
         */
        verticalFitToPages?: number;
    }
    /**
     * Represents the options in page layout margins.
     */
    interface PageLayoutMarginOptions {
        /**
         * Specifies the page layout bottom margin in the unit specified to use for printing.
         */
        bottom?: number;
        /**
         * Specifies the page layout footer margin in the unit specified to use for printing.
         */
        footer?: number;
        /**
         * Specifies the page layout header margin in the unit specified to use for printing.
         */
        header?: number;
        /**
         * Specifies the page layout left margin in the unit specified to use for printing.
         */
        left?: number;
        /**
         * Specifies the page layout right margin in the unit specified to use for printing.
         */
        right?: number;
        /**
         * Specifies the page layout top margin in the unit specified to use for printing.
         */
        top?: number;
    }
    /**
     * Represents the entity that is mentioned in comments.
     */
    interface CommentMention {
        /**
         * The email address of the entity that is mentioned in comment.
         */
        email: string;
        /**
         * The id of the entity. The id matches one of the ids in `CommentRichContent.richContent`.
         */
        id: number;
        /**
         * The name of the entity that is mentioned in comment.
         */
        name: string;
    }
    /**
     * Represents the content contained within a comment or comment reply. Rich content incudes the text string and any other objects contained within the comment body, such as mentions.
     */
    interface CommentRichContent {
        /**
         * An array containing all the entities (e.g., people) mentioned within the comment.
         */
        mentions?: CommentMention[];
        /**
         * Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an id attribute of 0, and the second mentioned entity has an id attribute of 1).
         */
        richContent: string;
    }

    //
    // Enum
    //
    /**
     * Represents the criteria for the top/bottom values filter.
     */
    enum PivotFilterTopBottomCriterion {
        invalid = "Invalid",

        topItems = "TopItems",

        topPercent = "TopPercent",

        topSum = "TopSum",

        bottomItems = "BottomItems",

        bottomPercent = "BottomPercent",

        bottomSum = "BottomSum",
    }
    /**
     * Represents the sort direction.
     */
    enum SortBy {
        /**
         * Ascending sort. Smallest to largest or A to Z.
         */
        ascending = "Ascending",

        /**
         * Descending sort. Largest to smallest or Z to A.
         */
        descending = "Descending",
    }
    /**
     * Aggregation Function for the Data Pivot Field.
     */
    enum AggregationFunction {
        /**
         * Aggregation function is unknown or unsupported.
         */
        unknown = "Unknown",

        /**
         * Excel will automatically select the aggregation based on the data items.
         */
        automatic = "Automatic",

        /**
         * Aggregate using the sum of the data, equivalent to the SUM function.
         */
        sum = "Sum",

        /**
         * Aggregate using the count of items in the data, equivalent to the COUNTA function.
         */
        count = "Count",

        /**
         * Aggregate using the average of the data, equivalent to the AVERAGE function.
         */
        average = "Average",

        /**
         * Aggregate using the maximum value of the data, equivalent to the MAX function.
         */
        max = "Max",

        /**
         * Aggregate using the minimum value of the data, equivalent to the MIN function.
         */
        min = "Min",

        /**
         * Aggregate using the product of the data, equivalent to the PRODUCT function.
         */
        product = "Product",

        /**
         * Aggregate using the count of numbers in the data, equivalent to the COUNT function.
         */
        countNumbers = "CountNumbers",

        /**
         * Aggregate using the standard deviation of the data, equivalent to the STDEV function.
         */
        standardDeviation = "StandardDeviation",

        /**
         * Aggregate using the standard deviation of the data, equivalent to the STDEVP function.
         */
        standardDeviationP = "StandardDeviationP",

        /**
         * Aggregate using the variance of the data, equivalent to the VAR function.
         */
        variance = "Variance",

        /**
         * Aggregate using the variance of the data, equivalent to the VARP function.
         */
        varianceP = "VarianceP",
    }
    /**
     * The ShowAs Calculation function for the Data Pivot Field.
     */
    enum ShowAsCalculation {
        /**
         * Calculation is unknown or unsupported.
         */
        unknown = "Unknown",

        /**
         * No calculation is applied.
         */
        none = "None",

        /**
         * Percent of the grand total.
         */
        percentOfGrandTotal = "PercentOfGrandTotal",

        /**
         * Percent of the row total.
         */
        percentOfRowTotal = "PercentOfRowTotal",

        /**
         * Percent of the column total.
         */
        percentOfColumnTotal = "PercentOfColumnTotal",

        /**
         * Percent of the row total for the specified Base Field.
         */
        percentOfParentRowTotal = "PercentOfParentRowTotal",

        /**
         * Percent of the column total for the specified Base Field.
         */
        percentOfParentColumnTotal = "PercentOfParentColumnTotal",

        /**
         * Percent of the grand total for the specified Base Field.
         */
        percentOfParentTotal = "PercentOfParentTotal",

        /**
         * Percent of the specified Base Field and Base Item.
         */
        percentOf = "PercentOf",

        /**
         * Running Total of the specified Base Field.
         */
        runningTotal = "RunningTotal",

        /**
         * Percent Running Total of the specified Base Field.
         */
        percentRunningTotal = "PercentRunningTotal",

        /**
         * Difference from the specified Base Field and Base Item.
         */
        differenceFrom = "DifferenceFrom",

        /**
         * Difference from the specified Base Field and Base Item.
         */
        percentDifferenceFrom = "PercentDifferenceFrom",

        /**
         * Ascending Rank of the specified Base Field.
         */
        rankAscending = "RankAscending",

        /**
         * Descending Rank of the specified Base Field.
         */
        rankDecending = "RankDecending",

        /**
         * Calculates the values as follows:
         * ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total))
         */
        index = "Index",
    }
    /**
     * Represents the axis from which to get the PivotItems.
     */
    enum PivotAxis {
        /**
         * The axis or region is unknown or unsupported.
         */
        unknown = "Unknown",

        /**
         * The row axis.
         */
        row = "Row",

        /**
         * The column axis.
         */
        column = "Column",

        /**
         * The data axis.
         */
        data = "Data",

        /**
         * The filter axis.
         */
        filter = "Filter",
    }
    enum ChartAxisType {
        invalid = "Invalid",

        /**
         * Axis displays categories.
         */
        category = "Category",

        /**
         * Axis displays values.
         */
        value = "Value",

        /**
         * Axis displays data series.
         */
        series = "Series",
    }
    enum ChartAxisGroup {
        primary = "Primary",

        secondary = "Secondary",
    }
    enum ChartAxisScaleType {
        linear = "Linear",

        logarithmic = "Logarithmic",
    }
    enum ChartAxisPosition {
        automatic = "Automatic",

        maximum = "Maximum",

        minimum = "Minimum",

        custom = "Custom",
    }
    enum ChartAxisTickMark {
        none = "None",

        cross = "Cross",

        inside = "Inside",

        outside = "Outside",
    }
    /**
     * Represents the state of calculation across the entire Excel application.
     */
    enum CalculationState {
        /**
         * Calculations complete.
         */
        done = "Done",

        /**
         * Calculations in progress.
         */
        calculating = "Calculating",

        /**
         * Changes that trigger calculation have been made, but a recalculation has not yet been performed.
         */
        pending = "Pending",
    }
    enum ChartAxisTickLabelPosition {
        nextToAxis = "NextToAxis",

        high = "High",

        low = "Low",

        none = "None",
    }
    enum ChartAxisDisplayUnit {
        /**
         * Default option. This will reset display unit to the axis, and set unit label invisible.
         */
        none = "None",

        /**
         * This will set the axis in units of hundreds.
         */
        hundreds = "Hundreds",

        /**
         * This will set the axis in units of thousands.
         */
        thousands = "Thousands",

        /**
         * This will set the axis in units of tens of thousands.
         */
        tenThousands = "TenThousands",

        /**
         * This will set the axis in units of hundreds of thousands.
         */
        hundredThousands = "HundredThousands",

        /**
         * This will set the axis in units of millions.
         */
        millions = "Millions",

        /**
         * This will set the axis in units of tens of millions.
         */
        tenMillions = "TenMillions",

        /**
         * This will set the axis in units of hundreds of millions.
         */
        hundredMillions = "HundredMillions",

        /**
         * This will set the axis in units of billions.
         */
        billions = "Billions",

        /**
         * This will set the axis in units of trillions.
         */
        trillions = "Trillions",

        /**
         * This will set the axis in units of custom value.
         */
        custom = "Custom",
    }
    /**
     * Specifies the unit of time for chart axes and data series.
     */
    enum ChartAxisTimeUnit {
        days = "Days",

        months = "Months",

        years = "Years",
    }
    /**
     * Represents the quartile calculation type of chart series layout. Only applies to a box and whisker chart.
     */
    enum ChartBoxQuartileCalculation {
        inclusive = "Inclusive",

        exclusive = "Exclusive",
    }
    /**
     * Specifies the type of the category axis.
     */
    enum ChartAxisCategoryType {
        /**
         * Excel controls the axis type.
         */
        automatic = "Automatic",

        /**
         * Axis groups data by an arbitrary set of categories.
         */
        textAxis = "TextAxis",

        /**
         * Axis groups data on a time scale.
         */
        dateAxis = "DateAxis",
    }
    /**
     * Specifies the bin's type of a histogram chart or pareto chart series.
     */
    enum ChartBinType {
        category = "Category",

        auto = "Auto",

        binWidth = "BinWidth",

        binCount = "BinCount",
    }
    enum ChartLineStyle {
        none = "None",

        continuous = "Continuous",

        dash = "Dash",

        dashDot = "DashDot",

        dashDotDot = "DashDotDot",

        dot = "Dot",

        grey25 = "Grey25",

        grey50 = "Grey50",

        grey75 = "Grey75",

        automatic = "Automatic",

        roundDot = "RoundDot",
    }
    enum ChartDataLabelPosition {
        invalid = "Invalid",

        none = "None",

        center = "Center",

        insideEnd = "InsideEnd",

        insideBase = "InsideBase",

        outsideEnd = "OutsideEnd",

        left = "Left",

        right = "Right",

        top = "Top",

        bottom = "Bottom",

        bestFit = "BestFit",

        callout = "Callout",
    }
    /**
     * Represents which parts of the error bar to include.
     */
    enum ChartErrorBarsInclude {
        both = "Both",

        minusValues = "MinusValues",

        plusValues = "PlusValues",
    }
    /**
     * Represents the range type for error bars.
     */
    enum ChartErrorBarsType {
        fixedValue = "FixedValue",

        percent = "Percent",

        stDev = "StDev",

        stError = "StError",

        custom = "Custom",
    }
    /**
     * Represents the mapping level of a chart series. This only applies to region map charts.
     */
    enum ChartMapAreaLevel {
        automatic = "Automatic",

        dataOnly = "DataOnly",

        city = "City",

        county = "County",

        state = "State",

        country = "Country",

        continent = "Continent",

        world = "World",
    }
    /**
     * Represents the gradient style of a chart series. This is only applicable for region map charts.
     */
    enum ChartGradientStyle {
        twoPhaseColor = "TwoPhaseColor",

        threePhaseColor = "ThreePhaseColor",
    }
    /**
     * Represents the gradient style type of a chart series. This is only applicable for region map charts.
     */
    enum ChartGradientStyleType {
        extremeValue = "ExtremeValue",

        number = "Number",

        percent = "Percent",
    }
    /**
     * Represents the position of chart title.
     */
    enum ChartTitlePosition {
        automatic = "Automatic",

        top = "Top",

        bottom = "Bottom",

        left = "Left",

        right = "Right",
    }
    enum ChartLegendPosition {
        invalid = "Invalid",

        top = "Top",

        bottom = "Bottom",

        left = "Left",

        right = "Right",

        corner = "Corner",

        custom = "Custom",
    }
    enum ChartMarkerStyle {
        invalid = "Invalid",

        automatic = "Automatic",

        none = "None",

        square = "Square",

        diamond = "Diamond",

        triangle = "Triangle",

        x = "X",

        star = "Star",

        dot = "Dot",

        dash = "Dash",

        circle = "Circle",

        plus = "Plus",

        picture = "Picture",
    }
    enum ChartPlotAreaPosition {
        automatic = "Automatic",

        custom = "Custom",
    }
    /**
     * Represents the region level of a chart series layout. This only applies to region map charts.
     */
    enum ChartMapLabelStrategy {
        none = "None",

        bestFit = "BestFit",

        showAll = "ShowAll",
    }
    /**
     * Represents the region projection type of a chart series layout. This only applies to region map charts.
     */
    enum ChartMapProjectionType {
        automatic = "Automatic",

        mercator = "Mercator",

        miller = "Miller",

        robinson = "Robinson",

        albers = "Albers",
    }
    /**
     * Represents the parent label strategy of the chart series layout. This only applies to treemap charts
     */
    enum ChartParentLabelStrategy {
        none = "None",

        banner = "Banner",

        overlapping = "Overlapping",
    }
    /**
     * Specifies whether the series are by rows or by columns. On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; in Excel on the web, "auto" will simply default to "columns".
     */
    enum ChartSeriesBy {
        /**
         * On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; in Excel on the web, "auto" will simply default to "columns".
         */
        auto = "Auto",

        columns = "Columns",

        rows = "Rows",
    }
    /**
     * Represents the horizontal alignment for the specified object.
     */
    enum ChartTextHorizontalAlignment {
        center = "Center",

        left = "Left",

        right = "Right",

        justify = "Justify",

        distributed = "Distributed",
    }
    /**
     * Represents the vertical alignment for the specified object.
     */
    enum ChartTextVerticalAlignment {
        center = "Center",

        bottom = "Bottom",

        top = "Top",

        justify = "Justify",

        distributed = "Distributed",
    }
    enum ChartTickLabelAlignment {
        center = "Center",

        left = "Left",

        right = "Right",
    }
    enum ChartType {
        invalid = "Invalid",

        columnClustered = "ColumnClustered",

        columnStacked = "ColumnStacked",

        columnStacked100 = "ColumnStacked100",

        _3DColumnClustered = "3DColumnClustered",

        _3DColumnStacked = "3DColumnStacked",

        _3DColumnStacked100 = "3DColumnStacked100",

        barClustered = "BarClustered",

        barStacked = "BarStacked",

        barStacked100 = "BarStacked100",

        _3DBarClustered = "3DBarClustered",

        _3DBarStacked = "3DBarStacked",

        _3DBarStacked100 = "3DBarStacked100",

        lineStacked = "LineStacked",

        lineStacked100 = "LineStacked100",

        lineMarkers = "LineMarkers",

        lineMarkersStacked = "LineMarkersStacked",

        lineMarkersStacked100 = "LineMarkersStacked100",

        pieOfPie = "PieOfPie",

        pieExploded = "PieExploded",

        _3DPieExploded = "3DPieExploded",

        barOfPie = "BarOfPie",

        xyscatterSmooth = "XyscatterSmooth",

        xyscatterSmoothNoMarkers = "XyscatterSmoothNoMarkers",

        xyscatterLines = "XyscatterLines",

        xyscatterLinesNoMarkers = "XyscatterLinesNoMarkers",

        areaStacked = "AreaStacked",

        areaStacked100 = "AreaStacked100",

        _3DAreaStacked = "3DAreaStacked",

        _3DAreaStacked100 = "3DAreaStacked100",

        doughnutExploded = "DoughnutExploded",

        radarMarkers = "RadarMarkers",

        radarFilled = "RadarFilled",

        surface = "Surface",

        surfaceWireframe = "SurfaceWireframe",

        surfaceTopView = "SurfaceTopView",

        surfaceTopViewWireframe = "SurfaceTopViewWireframe",

        bubble = "Bubble",

        bubble3DEffect = "Bubble3DEffect",

        stockHLC = "StockHLC",

        stockOHLC = "StockOHLC",

        stockVHLC = "StockVHLC",

        stockVOHLC = "StockVOHLC",

        cylinderColClustered = "CylinderColClustered",

        cylinderColStacked = "CylinderColStacked",

        cylinderColStacked100 = "CylinderColStacked100",

        cylinderBarClustered = "CylinderBarClustered",

        cylinderBarStacked = "CylinderBarStacked",

        cylinderBarStacked100 = "CylinderBarStacked100",

        cylinderCol = "CylinderCol",

        coneColClustered = "ConeColClustered",

        coneColStacked = "ConeColStacked",

        coneColStacked100 = "ConeColStacked100",

        coneBarClustered = "ConeBarClustered",

        coneBarStacked = "ConeBarStacked",

        coneBarStacked100 = "ConeBarStacked100",

        coneCol = "ConeCol",

        pyramidColClustered = "PyramidColClustered",

        pyramidColStacked = "PyramidColStacked",

        pyramidColStacked100 = "PyramidColStacked100",

        pyramidBarClustered = "PyramidBarClustered",

        pyramidBarStacked = "PyramidBarStacked",

        pyramidBarStacked100 = "PyramidBarStacked100",

        pyramidCol = "PyramidCol",

        _3DColumn = "3DColumn",

        line = "Line",

        _3DLine = "3DLine",

        _3DPie = "3DPie",

        pie = "Pie",

        xyscatter = "Xyscatter",

        _3DArea = "3DArea",

        area = "Area",

        doughnut = "Doughnut",

        radar = "Radar",

        histogram = "Histogram",

        boxwhisker = "Boxwhisker",

        pareto = "Pareto",

        regionMap = "RegionMap",

        treemap = "Treemap",

        waterfall = "Waterfall",

        sunburst = "Sunburst",

        funnel = "Funnel",
    }
    enum ChartUnderlineStyle {
        none = "None",

        single = "Single",
    }
    enum ChartDisplayBlanksAs {
        notPlotted = "NotPlotted",

        zero = "Zero",

        interplotted = "Interplotted",
    }
    enum ChartPlotBy {
        rows = "Rows",

        columns = "Columns",
    }
    enum ChartSplitType {
        splitByPosition = "SplitByPosition",

        splitByValue = "SplitByValue",

        splitByPercentValue = "SplitByPercentValue",

        splitByCustomSplit = "SplitByCustomSplit",
    }
    enum ChartColorScheme {
        colorfulPalette1 = "ColorfulPalette1",

        colorfulPalette2 = "ColorfulPalette2",

        colorfulPalette3 = "ColorfulPalette3",

        colorfulPalette4 = "ColorfulPalette4",

        monochromaticPalette1 = "MonochromaticPalette1",

        monochromaticPalette2 = "MonochromaticPalette2",

        monochromaticPalette3 = "MonochromaticPalette3",

        monochromaticPalette4 = "MonochromaticPalette4",

        monochromaticPalette5 = "MonochromaticPalette5",

        monochromaticPalette6 = "MonochromaticPalette6",

        monochromaticPalette7 = "MonochromaticPalette7",

        monochromaticPalette8 = "MonochromaticPalette8",

        monochromaticPalette9 = "MonochromaticPalette9",

        monochromaticPalette10 = "MonochromaticPalette10",

        monochromaticPalette11 = "MonochromaticPalette11",

        monochromaticPalette12 = "MonochromaticPalette12",

        monochromaticPalette13 = "MonochromaticPalette13",
    }
    enum ChartTrendlineType {
        linear = "Linear",

        exponential = "Exponential",

        logarithmic = "Logarithmic",

        movingAverage = "MovingAverage",

        polynomial = "Polynomial",

        power = "Power",
    }
    /**
     * Specifies where in the z-order a shape should be moved relative to other shapes.
     */
    enum ShapeZOrder {
        bringToFront = "BringToFront",

        bringForward = "BringForward",

        sendToBack = "SendToBack",

        sendBackward = "SendBackward",
    }
    /**
     * Specifies the type of a shape.
     */
    enum ShapeType {
        unsupported = "Unsupported",

        image = "Image",

        geometricShape = "GeometricShape",

        group = "Group",

        line = "Line",
    }
    /**
     * Specifies whether the shape is scaled relative to its original or current size.
     */
    enum ShapeScaleType {
        currentSize = "CurrentSize",

        originalSize = "OriginalSize",
    }
    /**
     * Specifies which part of the shape retains its position when the shape is scaled.
     */
    enum ShapeScaleFrom {
        scaleFromTopLeft = "ScaleFromTopLeft",

        scaleFromMiddle = "ScaleFromMiddle",

        scaleFromBottomRight = "ScaleFromBottomRight",
    }
    /**
     * Specifies a shape's fill type.
     */
    enum ShapeFillType {
        /**
         * No fill.
         */
        noFill = "NoFill",

        /**
         * Solid fill.
         */
        solid = "Solid",

        /**
         * Gradient fill.
         */
        gradient = "Gradient",

        /**
         * Pattern fill.
         */
        pattern = "Pattern",

        /**
         * Picture and texture fill.
         */
        pictureAndTexture = "PictureAndTexture",

        /**
         * Mixed fill.
         */
        mixed = "Mixed",
    }
    /**
     * The type of underline applied to a font.
     */
    enum ShapeFontUnderlineStyle {
        none = "None",

        single = "Single",

        double = "Double",

        heavy = "Heavy",

        dotted = "Dotted",

        dottedHeavy = "DottedHeavy",

        dash = "Dash",

        dashHeavy = "DashHeavy",

        dashLong = "DashLong",

        dashLongHeavy = "DashLongHeavy",

        dotDash = "DotDash",

        dotDashHeavy = "DotDashHeavy",

        dotDotDash = "DotDotDash",

        dotDotDashHeavy = "DotDotDashHeavy",

        wavy = "Wavy",

        wavyHeavy = "WavyHeavy",

        wavyDouble = "WavyDouble",
    }
    /**
     * The format of the image.
     */
    enum PictureFormat {
        unknown = "Unknown",

        /**
         * Bitmap image.
         */
        bmp = "Bmp",

        /**
         * Joint Photographic Experts Group.
         */
        jpeg = "Jpeg",

        /**
         * Graphics Interchange Format.
         */
        gif = "Gif",

        /**
         * Portable Network Graphics.
         */
        png = "Png",

        /**
         * Scalable Vector Graphic.
         */
        svg = "Svg",
    }
    /**
     * The style for a line.
     */
    enum ShapeLineStyle {
        /**
         * Single line.
         */
        single = "Single",

        /**
         * Thick line with a thin line on each side.
         */
        thickBetweenThin = "ThickBetweenThin",

        /**
         * Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
         */
        thickThin = "ThickThin",

        /**
         * Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
         */
        thinThick = "ThinThick",

        /**
         * Two thin lines.
         */
        thinThin = "ThinThin",
    }
    /**
     * The dash style for a line.
     */
    enum ShapeLineDashStyle {
        dash = "Dash",

        dashDot = "DashDot",

        dashDotDot = "DashDotDot",

        longDash = "LongDash",

        longDashDot = "LongDashDot",

        roundDot = "RoundDot",

        solid = "Solid",

        squareDot = "SquareDot",

        longDashDotDot = "LongDashDotDot",

        systemDash = "SystemDash",

        systemDot = "SystemDot",

        systemDashDot = "SystemDashDot",
    }
    enum ArrowheadLength {
        short = "Short",

        medium = "Medium",

        long = "Long",
    }
    enum ArrowheadStyle {
        none = "None",

        triangle = "Triangle",

        stealth = "Stealth",

        diamond = "Diamond",

        oval = "Oval",

        open = "Open",
    }
    enum ArrowheadWidth {
        narrow = "Narrow",

        medium = "Medium",

        wide = "Wide",
    }
    enum BindingType {
        range = "Range",

        table = "Table",

        text = "Text",
    }
    enum BorderIndex {
        edgeTop = "EdgeTop",

        edgeBottom = "EdgeBottom",

        edgeLeft = "EdgeLeft",

        edgeRight = "EdgeRight",

        insideVertical = "InsideVertical",

        insideHorizontal = "InsideHorizontal",

        diagonalDown = "DiagonalDown",

        diagonalUp = "DiagonalUp",
    }
    enum BorderLineStyle {
        none = "None",

        continuous = "Continuous",

        dash = "Dash",

        dashDot = "DashDot",

        dashDotDot = "DashDotDot",

        dot = "Dot",

        double = "Double",

        slantDashDot = "SlantDashDot",
    }
    enum BorderWeight {
        hairline = "Hairline",

        thin = "Thin",

        medium = "Medium",

        thick = "Thick",
    }
    enum CalculationMode {
        /**
         * The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
         */
        automatic = "Automatic",

        /**
         * Calculates new formula results every time the relevant data is changed, unless the formula is in a data table.
         */
        automaticExceptTables = "AutomaticExceptTables",

        /**
         * Calculations only occur when the user or add-in requests them.
         */
        manual = "Manual",
    }
    enum CalculationType {
        /**
         * Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty.
         */
        recalculate = "Recalculate",

        /**
         * This will mark all cells as dirty and then recalculate them.
         */
        full = "Full",

        /**
         * This will rebuild the full dependency chain, mark all cells as dirty and then recalculate them.
         */
        fullRebuild = "FullRebuild",
    }
    enum ClearApplyTo {
        all = "All",

        /**
         * Clears all formatting for the range.
         */
        formats = "Formats",

        /**
         * Clears the contents of the range.
         */
        contents = "Contents",

        /**
         * Clears all hyperlinks, but leaves all content and formatting intact.
         */
        hyperlinks = "Hyperlinks",

        /**
         * Removes hyperlinks and formatting for the cell but leaves content, conditional formats, and data validation intact.
         */
        removeHyperlinks = "RemoveHyperlinks",
    }
    /**
     * Represents the format options for a Data Bar Axis.
     */
    enum ConditionalDataBarAxisFormat {
        automatic = "Automatic",

        none = "None",

        cellMidPoint = "CellMidPoint",
    }
    /**
     * Represents the Data Bar direction within a cell.
     */
    enum ConditionalDataBarDirection {
        context = "Context",

        leftToRight = "LeftToRight",

        rightToLeft = "RightToLeft",
    }
    /**
     * Represents the direction for a selection.
     */
    enum ConditionalFormatDirection {
        top = "Top",

        bottom = "Bottom",
    }
    enum ConditionalFormatType {
        custom = "Custom",

        dataBar = "DataBar",

        colorScale = "ColorScale",

        iconSet = "IconSet",

        topBottom = "TopBottom",

        presetCriteria = "PresetCriteria",

        containsText = "ContainsText",

        cellValue = "CellValue",
    }
    /**
     * Represents the types of conditional format values.
     */
    enum ConditionalFormatRuleType {
        invalid = "Invalid",

        automatic = "Automatic",

        lowestValue = "LowestValue",

        highestValue = "HighestValue",

        number = "Number",

        percent = "Percent",

        formula = "Formula",

        percentile = "Percentile",
    }
    /**
     * Represents the types of icon conditional format.
     */
    enum ConditionalFormatIconRuleType {
        invalid = "Invalid",

        number = "Number",

        percent = "Percent",

        formula = "Formula",

        percentile = "Percentile",
    }
    /**
     * Represents the types of color criterion for conditional formatting.
     */
    enum ConditionalFormatColorCriterionType {
        invalid = "Invalid",

        lowestValue = "LowestValue",

        highestValue = "HighestValue",

        number = "Number",

        percent = "Percent",

        formula = "Formula",

        percentile = "Percentile",
    }
    /**
     * Represents the criteria for the above/below average conditional format type.
     */
    enum ConditionalTopBottomCriterionType {
        invalid = "Invalid",

        topItems = "TopItems",

        topPercent = "TopPercent",

        bottomItems = "BottomItems",

        bottomPercent = "BottomPercent",
    }
    /**
     * Represents the criteria for the Preset Criteria conditional format type.
     */
    enum ConditionalFormatPresetCriterion {
        invalid = "Invalid",

        blanks = "Blanks",

        nonBlanks = "NonBlanks",

        errors = "Errors",

        nonErrors = "NonErrors",

        yesterday = "Yesterday",

        today = "Today",

        tomorrow = "Tomorrow",

        lastSevenDays = "LastSevenDays",

        lastWeek = "LastWeek",

        thisWeek = "ThisWeek",

        nextWeek = "NextWeek",

        lastMonth = "LastMonth",

        thisMonth = "ThisMonth",

        nextMonth = "NextMonth",

        aboveAverage = "AboveAverage",

        belowAverage = "BelowAverage",

        equalOrAboveAverage = "EqualOrAboveAverage",

        equalOrBelowAverage = "EqualOrBelowAverage",

        oneStdDevAboveAverage = "OneStdDevAboveAverage",

        oneStdDevBelowAverage = "OneStdDevBelowAverage",

        twoStdDevAboveAverage = "TwoStdDevAboveAverage",

        twoStdDevBelowAverage = "TwoStdDevBelowAverage",

        threeStdDevAboveAverage = "ThreeStdDevAboveAverage",

        threeStdDevBelowAverage = "ThreeStdDevBelowAverage",

        uniqueValues = "UniqueValues",

        duplicateValues = "DuplicateValues",
    }
    /**
     * Represents the operator of the text conditional format type.
     */
    enum ConditionalTextOperator {
        invalid = "Invalid",

        contains = "Contains",

        notContains = "NotContains",

        beginsWith = "BeginsWith",

        endsWith = "EndsWith",
    }
    /**
     * Represents the operator of the text conditional format type.
     */
    enum ConditionalCellValueOperator {
        invalid = "Invalid",

        between = "Between",

        notBetween = "NotBetween",

        equalTo = "EqualTo",

        notEqualTo = "NotEqualTo",

        greaterThan = "GreaterThan",

        lessThan = "LessThan",

        greaterThanOrEqual = "GreaterThanOrEqual",

        lessThanOrEqual = "LessThanOrEqual",
    }
    /**
     * Represents the operator for each icon criteria.
     */
    enum ConditionalIconCriterionOperator {
        invalid = "Invalid",

        greaterThan = "GreaterThan",

        greaterThanOrEqual = "GreaterThanOrEqual",
    }
    enum ConditionalRangeBorderIndex {
        edgeTop = "EdgeTop",

        edgeBottom = "EdgeBottom",

        edgeLeft = "EdgeLeft",

        edgeRight = "EdgeRight",
    }
    enum ConditionalRangeBorderLineStyle {
        none = "None",

        continuous = "Continuous",

        dash = "Dash",

        dashDot = "DashDot",

        dashDotDot = "DashDotDot",

        dot = "Dot",
    }
    enum ConditionalRangeFontUnderlineStyle {
        none = "None",

        single = "Single",

        double = "Double",
    }
    /**
     * Represents Data validation type enum.
     */
    enum DataValidationType {
        /**
         * None means allow any value and so there is no data validation in the range.
         */
        none = "None",

        /**
         * Whole number data validation type
         */
        wholeNumber = "WholeNumber",

        /**
         * Decimal data validation type
         */
        decimal = "Decimal",

        /**
         * List data validation type
         */
        list = "List",

        /**
         * Date data validation type
         */
        date = "Date",

        /**
         * Time data validation type
         */
        time = "Time",

        /**
         * Text length data validation type
         */
        textLength = "TextLength",

        /**
         * Custom data validation type
         */
        custom = "Custom",

        /**
         * Inconsistent means that the range has inconsistent data validation (there are different rules on different cells)
         */
        inconsistent = "Inconsistent",

        /**
         * MixedCriteria means that the range has data validation present on some but not all cells
         */
        mixedCriteria = "MixedCriteria",
    }
    /**
     * Represents Data validation operator enum.
     */
    enum DataValidationOperator {
        between = "Between",

        notBetween = "NotBetween",

        equalTo = "EqualTo",

        notEqualTo = "NotEqualTo",

        greaterThan = "GreaterThan",

        lessThan = "LessThan",

        greaterThanOrEqualTo = "GreaterThanOrEqualTo",

        lessThanOrEqualTo = "LessThanOrEqualTo",
    }
    /**
     * Represents Data validation error alert style. The default is "Stop".
     */
    enum DataValidationAlertStyle {
        stop = "Stop",

        warning = "Warning",

        information = "Information",
    }
    enum DeleteShiftDirection {
        up = "Up",

        left = "Left",
    }
    enum DynamicFilterCriteria {
        unknown = "Unknown",

        aboveAverage = "AboveAverage",

        allDatesInPeriodApril = "AllDatesInPeriodApril",

        allDatesInPeriodAugust = "AllDatesInPeriodAugust",

        allDatesInPeriodDecember = "AllDatesInPeriodDecember",

        allDatesInPeriodFebruray = "AllDatesInPeriodFebruray",

        allDatesInPeriodJanuary = "AllDatesInPeriodJanuary",

        allDatesInPeriodJuly = "AllDatesInPeriodJuly",

        allDatesInPeriodJune = "AllDatesInPeriodJune",

        allDatesInPeriodMarch = "AllDatesInPeriodMarch",

        allDatesInPeriodMay = "AllDatesInPeriodMay",

        allDatesInPeriodNovember = "AllDatesInPeriodNovember",

        allDatesInPeriodOctober = "AllDatesInPeriodOctober",

        allDatesInPeriodQuarter1 = "AllDatesInPeriodQuarter1",

        allDatesInPeriodQuarter2 = "AllDatesInPeriodQuarter2",

        allDatesInPeriodQuarter3 = "AllDatesInPeriodQuarter3",

        allDatesInPeriodQuarter4 = "AllDatesInPeriodQuarter4",

        allDatesInPeriodSeptember = "AllDatesInPeriodSeptember",

        belowAverage = "BelowAverage",

        lastMonth = "LastMonth",

        lastQuarter = "LastQuarter",

        lastWeek = "LastWeek",

        lastYear = "LastYear",

        nextMonth = "NextMonth",

        nextQuarter = "NextQuarter",

        nextWeek = "NextWeek",

        nextYear = "NextYear",

        thisMonth = "ThisMonth",

        thisQuarter = "ThisQuarter",

        thisWeek = "ThisWeek",

        thisYear = "ThisYear",

        today = "Today",

        tomorrow = "Tomorrow",

        yearToDate = "YearToDate",

        yesterday = "Yesterday",
    }
    enum FilterDatetimeSpecificity {
        year = "Year",

        month = "Month",

        day = "Day",

        hour = "Hour",

        minute = "Minute",

        second = "Second",
    }
    enum FilterOn {
        bottomItems = "BottomItems",

        bottomPercent = "BottomPercent",

        cellColor = "CellColor",

        dynamic = "Dynamic",

        fontColor = "FontColor",

        values = "Values",

        topItems = "TopItems",

        topPercent = "TopPercent",

        icon = "Icon",

        custom = "Custom",
    }
    enum FilterOperator {
        and = "And",

        or = "Or",
    }
    enum HorizontalAlignment {
        general = "General",

        left = "Left",

        center = "Center",

        right = "Right",

        fill = "Fill",

        justify = "Justify",

        centerAcrossSelection = "CenterAcrossSelection",

        distributed = "Distributed",
    }
    enum IconSet {
        invalid = "Invalid",

        threeArrows = "ThreeArrows",

        threeArrowsGray = "ThreeArrowsGray",

        threeFlags = "ThreeFlags",

        threeTrafficLights1 = "ThreeTrafficLights1",

        threeTrafficLights2 = "ThreeTrafficLights2",

        threeSigns = "ThreeSigns",

        threeSymbols = "ThreeSymbols",

        threeSymbols2 = "ThreeSymbols2",

        fourArrows = "FourArrows",

        fourArrowsGray = "FourArrowsGray",

        fourRedToBlack = "FourRedToBlack",

        fourRating = "FourRating",

        fourTrafficLights = "FourTrafficLights",

        fiveArrows = "FiveArrows",

        fiveArrowsGray = "FiveArrowsGray",

        fiveRating = "FiveRating",

        fiveQuarters = "FiveQuarters",

        threeStars = "ThreeStars",

        threeTriangles = "ThreeTriangles",

        fiveBoxes = "FiveBoxes",
    }
    enum ImageFittingMode {
        fit = "Fit",

        fitAndCenter = "FitAndCenter",

        fill = "Fill",
    }
    enum InsertShiftDirection {
        down = "Down",

        right = "Right",
    }
    enum NamedItemScope {
        worksheet = "Worksheet",

        workbook = "Workbook",
    }
    enum NamedItemType {
        string = "String",

        integer = "Integer",

        double = "Double",

        boolean = "Boolean",

        range = "Range",

        error = "Error",

        array = "Array",
    }
    enum RangeUnderlineStyle {
        none = "None",

        single = "Single",

        double = "Double",

        singleAccountant = "SingleAccountant",

        doubleAccountant = "DoubleAccountant",
    }
    enum SheetVisibility {
        visible = "Visible",

        hidden = "Hidden",

        veryHidden = "VeryHidden",
    }
    enum RangeValueType {
        unknown = "Unknown",

        empty = "Empty",

        string = "String",

        integer = "Integer",

        double = "Double",

        boolean = "Boolean",

        error = "Error",

        richValue = "RichValue",
    }
    /**
     * Specifies the search direction.
     */
    enum SearchDirection {
        /**
         * Search in forward order.
         */
        forward = "Forward",

        /**
         * Search in reverse order.
         */
        backwards = "Backwards",
    }
    enum SortOrientation {
        rows = "Rows",

        columns = "Columns",
    }
    enum SortOn {
        value = "Value",

        cellColor = "CellColor",

        fontColor = "FontColor",

        icon = "Icon",
    }
    enum SortDataOption {
        normal = "Normal",

        textAsNumber = "TextAsNumber",
    }
    enum SortMethod {
        pinYin = "PinYin",

        strokeCount = "StrokeCount",
    }
    enum VerticalAlignment {
        top = "Top",

        center = "Center",

        bottom = "Bottom",

        justify = "Justify",

        distributed = "Distributed",
    }
    enum DocumentPropertyType {
        number = "Number",

        boolean = "Boolean",

        date = "Date",

        string = "String",

        float = "Float",
    }
    enum SubtotalLocationType {
        /**
         * Subtotals are at the top.
         */
        atTop = "AtTop",

        /**
         * Subtotals are at the bottom.
         */
        atBottom = "AtBottom",

        /**
         * Subtotals are off.
         */
        off = "Off",
    }
    enum PivotLayoutType {
        /**
         * A horizontally compressed form with labels from the next field in the same column.
         */
        compact = "Compact",

        /**
         * Inner fields' items are always on a new line relative to the outer fields' items.
         */
        tabular = "Tabular",

        /**
         * Inner fields' items are on same row as outer fields' items and subtotals are always on the bottom.
         */
        outline = "Outline",
    }
    enum ProtectionSelectionMode {
        /**
         * Selection is allowed for all cells.
         */
        normal = "Normal",

        /**
         * Selection is allowed only for cells that are not locked.
         */
        unlocked = "Unlocked",

        /**
         * Selection is not allowed for all cells.
         */
        none = "None",
    }
    enum PageOrientation {
        portrait = "Portrait",

        landscape = "Landscape",
    }
    enum PaperType {
        letter = "Letter",

        letterSmall = "LetterSmall",

        tabloid = "Tabloid",

        ledger = "Ledger",

        legal = "Legal",

        statement = "Statement",

        executive = "Executive",

        a3 = "A3",

        a4 = "A4",

        a4Small = "A4Small",

        a5 = "A5",

        b4 = "B4",

        b5 = "B5",

        folio = "Folio",

        quatro = "Quatro",

        paper10x14 = "Paper10x14",

        paper11x17 = "Paper11x17",

        note = "Note",

        envelope9 = "Envelope9",

        envelope10 = "Envelope10",

        envelope11 = "Envelope11",

        envelope12 = "Envelope12",

        envelope14 = "Envelope14",

        csheet = "Csheet",

        dsheet = "Dsheet",

        esheet = "Esheet",

        envelopeDL = "EnvelopeDL",

        envelopeC5 = "EnvelopeC5",

        envelopeC3 = "EnvelopeC3",

        envelopeC4 = "EnvelopeC4",

        envelopeC6 = "EnvelopeC6",

        envelopeC65 = "EnvelopeC65",

        envelopeB4 = "EnvelopeB4",

        envelopeB5 = "EnvelopeB5",

        envelopeB6 = "EnvelopeB6",

        envelopeItaly = "EnvelopeItaly",

        envelopeMonarch = "EnvelopeMonarch",

        envelopePersonal = "EnvelopePersonal",

        fanfoldUS = "FanfoldUS",

        fanfoldStdGerman = "FanfoldStdGerman",

        fanfoldLegalGerman = "FanfoldLegalGerman",
    }
    enum ReadingOrder {
        /**
         * Reading order is determined by the language of the first character entered.
         * If a right-to-left language character is entered first, reading order is right to left.
         * If a left-to-right language character is entered first, reading order is left to right.
         */
        context = "Context",

        /**
         * Left to right reading order
         */
        leftToRight = "LeftToRight",

        /**
         * Right to left reading order
         */
        rightToLeft = "RightToLeft",
    }
    enum BuiltInStyle {
        normal = "Normal",

        comma = "Comma",

        currency = "Currency",

        percent = "Percent",

        wholeComma = "WholeComma",

        wholeDollar = "WholeDollar",

        hlink = "Hlink",

        hlinkTrav = "HlinkTrav",

        note = "Note",

        warningText = "WarningText",

        emphasis1 = "Emphasis1",

        emphasis2 = "Emphasis2",

        emphasis3 = "Emphasis3",

        sheetTitle = "SheetTitle",

        heading1 = "Heading1",

        heading2 = "Heading2",

        heading3 = "Heading3",

        heading4 = "Heading4",

        input = "Input",

        output = "Output",

        calculation = "Calculation",

        checkCell = "CheckCell",

        linkedCell = "LinkedCell",

        total = "Total",

        good = "Good",

        bad = "Bad",

        neutral = "Neutral",

        accent1 = "Accent1",

        accent1_20 = "Accent1_20",

        accent1_40 = "Accent1_40",

        accent1_60 = "Accent1_60",

        accent2 = "Accent2",

        accent2_20 = "Accent2_20",

        accent2_40 = "Accent2_40",

        accent2_60 = "Accent2_60",

        accent3 = "Accent3",

        accent3_20 = "Accent3_20",

        accent3_40 = "Accent3_40",

        accent3_60 = "Accent3_60",

        accent4 = "Accent4",

        accent4_20 = "Accent4_20",

        accent4_40 = "Accent4_40",

        accent4_60 = "Accent4_60",

        accent5 = "Accent5",

        accent5_20 = "Accent5_20",

        accent5_40 = "Accent5_40",

        accent5_60 = "Accent5_60",

        accent6 = "Accent6",

        accent6_20 = "Accent6_20",

        accent6_40 = "Accent6_40",

        accent6_60 = "Accent6_60",

        explanatoryText = "ExplanatoryText",
    }
    enum PrintErrorType {
        asDisplayed = "AsDisplayed",

        blank = "Blank",

        dash = "Dash",

        notAvailable = "NotAvailable",
    }
    enum WorksheetPositionType {
        none = "None",

        before = "Before",

        after = "After",

        beginning = "Beginning",

        end = "End",
    }
    enum PrintComments {
        /**
         * Comments will not be printed.
         */
        noComments = "NoComments",

        /**
         * Comments will be printed as end notes at the end of the worksheet.
         */
        endSheet = "EndSheet",

        /**
         * Comments will be printed where they were inserted in the worksheet.
         */
        inPlace = "InPlace",
    }
    enum PrintOrder {
        /**
         * Process down the rows before processing across pages or page fields to the right.
         */
        downThenOver = "DownThenOver",

        /**
         * Process across pages or page fields to the right before moving down the rows.
         */
        overThenDown = "OverThenDown",
    }
    enum PrintMarginUnit {
        /**
         * Assign the page margins in points. A point is 1/72 of an inch.
         */
        points = "Points",

        /**
         * Assign the page margins in inches.
         */
        inches = "Inches",

        /**
         * Assign the page margins in centimeters.
         */
        centimeters = "Centimeters",
    }
    enum HeaderFooterState {
        /**
         * Only one general header/footer is used for all pages printed.
         */
        default = "Default",

        /**
         * There is a separate first page header/footer, and a general header/footer used for all other pages.
         */
        firstAndDefault = "FirstAndDefault",

        /**
         * There is a different header/footer for odd and even pages.
         */
        oddAndEven = "OddAndEven",

        /**
         * There is a separate first page header/footer, then there is a separate header/footer for odd and even pages.
         */
        firstOddAndEven = "FirstOddAndEven",
    }
    /**
     * The behavior types when AutoFill is used on a range in the workbook.
     */
    enum AutoFillType {
        /**
         * Populates the adjacent cells with data the selected data.
         */
        fillDefault = "FillDefault",

        /**
         * Populates the adjacent cells with data the selected data.
         */
        fillCopy = "FillCopy",

        /**
         * Populates the adjacent cells with data that follows a pattern in the copied cells.
         */
        fillSeries = "FillSeries",

        /**
         * Populates the adjacent cells with the selected formulas.
         */
        fillFormats = "FillFormats",

        /**
         * Populates the adjacent cells with the selected values.
         */
        fillValues = "FillValues",

        /**
         * A version of "FillSeries" for dates that bases the pattern on either the day of the month or the day of the week, depending on the context.
         */
        fillDays = "FillDays",

        /**
         * A version of "FillSeries" for dates that bases the pattern on the day of the week and only includes weekdays.
         */
        fillWeekdays = "FillWeekdays",

        /**
         * A version of "FillSeries" for dates that bases the pattern on the month.
         */
        fillMonths = "FillMonths",

        /**
         * A version of "FillSeries" for dates that bases the pattern on the year.
         */
        fillYears = "FillYears",

        /**
         * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a linear trend model.
         */
        linearTrend = "LinearTrend",

        /**
         * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a growth trend model.
         */
        growthTrend = "GrowthTrend",

        /**
         * Populates the adjacent cells by using Excel's FlashFill feature.
         */
        flashFill = "FlashFill",
    }
    enum GroupOption {
        /**
         * Group by rows.
         */
        byRows = "ByRows",

        /**
         * Group by columns.
         */
        byColumns = "ByColumns",
    }
    enum RangeCopyType {
        all = "All",

        formulas = "Formulas",

        values = "Values",

        formats = "Formats",
    }
    enum LinkedDataTypeState {
        none = "None",

        validLinkedData = "ValidLinkedData",

        disambiguationNeeded = "DisambiguationNeeded",

        brokenLinkedData = "BrokenLinkedData",

        fetchingData = "FetchingData",
    }
    /**
     * Specifies the shape type for a GeometricShape object.
     */
    enum GeometricShapeType {
        lineInverse = "LineInverse",

        triangle = "Triangle",

        rightTriangle = "RightTriangle",

        rectangle = "Rectangle",

        diamond = "Diamond",

        parallelogram = "Parallelogram",

        trapezoid = "Trapezoid",

        nonIsoscelesTrapezoid = "NonIsoscelesTrapezoid",

        pentagon = "Pentagon",

        hexagon = "Hexagon",

        heptagon = "Heptagon",

        octagon = "Octagon",

        decagon = "Decagon",

        dodecagon = "Dodecagon",

        star4 = "Star4",

        star5 = "Star5",

        star6 = "Star6",

        star7 = "Star7",

        star8 = "Star8",

        star10 = "Star10",

        star12 = "Star12",

        star16 = "Star16",

        star24 = "Star24",

        star32 = "Star32",

        roundRectangle = "RoundRectangle",

        round1Rectangle = "Round1Rectangle",

        round2SameRectangle = "Round2SameRectangle",

        round2DiagonalRectangle = "Round2DiagonalRectangle",

        snipRoundRectangle = "SnipRoundRectangle",

        snip1Rectangle = "Snip1Rectangle",

        snip2SameRectangle = "Snip2SameRectangle",

        snip2DiagonalRectangle = "Snip2DiagonalRectangle",

        plaque = "Plaque",

        ellipse = "Ellipse",

        teardrop = "Teardrop",

        homePlate = "HomePlate",

        chevron = "Chevron",

        pieWedge = "PieWedge",

        pie = "Pie",

        blockArc = "BlockArc",

        donut = "Donut",

        noSmoking = "NoSmoking",

        rightArrow = "RightArrow",

        leftArrow = "LeftArrow",

        upArrow = "UpArrow",

        downArrow = "DownArrow",

        stripedRightArrow = "StripedRightArrow",

        notchedRightArrow = "NotchedRightArrow",

        bentUpArrow = "BentUpArrow",

        leftRightArrow = "LeftRightArrow",

        upDownArrow = "UpDownArrow",

        leftUpArrow = "LeftUpArrow",

        leftRightUpArrow = "LeftRightUpArrow",

        quadArrow = "QuadArrow",

        leftArrowCallout = "LeftArrowCallout",

        rightArrowCallout = "RightArrowCallout",

        upArrowCallout = "UpArrowCallout",

        downArrowCallout = "DownArrowCallout",

        leftRightArrowCallout = "LeftRightArrowCallout",

        upDownArrowCallout = "UpDownArrowCallout",

        quadArrowCallout = "QuadArrowCallout",

        bentArrow = "BentArrow",

        uturnArrow = "UturnArrow",

        circularArrow = "CircularArrow",

        leftCircularArrow = "LeftCircularArrow",

        leftRightCircularArrow = "LeftRightCircularArrow",

        curvedRightArrow = "CurvedRightArrow",

        curvedLeftArrow = "CurvedLeftArrow",

        curvedUpArrow = "CurvedUpArrow",

        curvedDownArrow = "CurvedDownArrow",

        swooshArrow = "SwooshArrow",

        cube = "Cube",

        can = "Can",

        lightningBolt = "LightningBolt",

        heart = "Heart",

        sun = "Sun",

        moon = "Moon",

        smileyFace = "SmileyFace",

        irregularSeal1 = "IrregularSeal1",

        irregularSeal2 = "IrregularSeal2",

        foldedCorner = "FoldedCorner",

        bevel = "Bevel",

        frame = "Frame",

        halfFrame = "HalfFrame",

        corner = "Corner",

        diagonalStripe = "DiagonalStripe",

        chord = "Chord",

        arc = "Arc",

        leftBracket = "LeftBracket",

        rightBracket = "RightBracket",

        leftBrace = "LeftBrace",

        rightBrace = "RightBrace",

        bracketPair = "BracketPair",

        bracePair = "BracePair",

        callout1 = "Callout1",

        callout2 = "Callout2",

        callout3 = "Callout3",

        accentCallout1 = "AccentCallout1",

        accentCallout2 = "AccentCallout2",

        accentCallout3 = "AccentCallout3",

        borderCallout1 = "BorderCallout1",

        borderCallout2 = "BorderCallout2",

        borderCallout3 = "BorderCallout3",

        accentBorderCallout1 = "AccentBorderCallout1",

        accentBorderCallout2 = "AccentBorderCallout2",

        accentBorderCallout3 = "AccentBorderCallout3",

        wedgeRectCallout = "WedgeRectCallout",

        wedgeRRectCallout = "WedgeRRectCallout",

        wedgeEllipseCallout = "WedgeEllipseCallout",

        cloudCallout = "CloudCallout",

        cloud = "Cloud",

        ribbon = "Ribbon",

        ribbon2 = "Ribbon2",

        ellipseRibbon = "EllipseRibbon",

        ellipseRibbon2 = "EllipseRibbon2",

        leftRightRibbon = "LeftRightRibbon",

        verticalScroll = "VerticalScroll",

        horizontalScroll = "HorizontalScroll",

        wave = "Wave",

        doubleWave = "DoubleWave",

        plus = "Plus",

        flowChartProcess = "FlowChartProcess",

        flowChartDecision = "FlowChartDecision",

        flowChartInputOutput = "FlowChartInputOutput",

        flowChartPredefinedProcess = "FlowChartPredefinedProcess",

        flowChartInternalStorage = "FlowChartInternalStorage",

        flowChartDocument = "FlowChartDocument",

        flowChartMultidocument = "FlowChartMultidocument",

        flowChartTerminator = "FlowChartTerminator",

        flowChartPreparation = "FlowChartPreparation",

        flowChartManualInput = "FlowChartManualInput",

        flowChartManualOperation = "FlowChartManualOperation",

        flowChartConnector = "FlowChartConnector",

        flowChartPunchedCard = "FlowChartPunchedCard",

        flowChartPunchedTape = "FlowChartPunchedTape",

        flowChartSummingJunction = "FlowChartSummingJunction",

        flowChartOr = "FlowChartOr",

        flowChartCollate = "FlowChartCollate",

        flowChartSort = "FlowChartSort",

        flowChartExtract = "FlowChartExtract",

        flowChartMerge = "FlowChartMerge",

        flowChartOfflineStorage = "FlowChartOfflineStorage",

        flowChartOnlineStorage = "FlowChartOnlineStorage",

        flowChartMagneticTape = "FlowChartMagneticTape",

        flowChartMagneticDisk = "FlowChartMagneticDisk",

        flowChartMagneticDrum = "FlowChartMagneticDrum",

        flowChartDisplay = "FlowChartDisplay",

        flowChartDelay = "FlowChartDelay",

        flowChartAlternateProcess = "FlowChartAlternateProcess",

        flowChartOffpageConnector = "FlowChartOffpageConnector",

        actionButtonBlank = "ActionButtonBlank",

        actionButtonHome = "ActionButtonHome",

        actionButtonHelp = "ActionButtonHelp",

        actionButtonInformation = "ActionButtonInformation",

        actionButtonForwardNext = "ActionButtonForwardNext",

        actionButtonBackPrevious = "ActionButtonBackPrevious",

        actionButtonEnd = "ActionButtonEnd",

        actionButtonBeginning = "ActionButtonBeginning",

        actionButtonReturn = "ActionButtonReturn",

        actionButtonDocument = "ActionButtonDocument",

        actionButtonSound = "ActionButtonSound",

        actionButtonMovie = "ActionButtonMovie",

        gear6 = "Gear6",

        gear9 = "Gear9",

        funnel = "Funnel",

        mathPlus = "MathPlus",

        mathMinus = "MathMinus",

        mathMultiply = "MathMultiply",

        mathDivide = "MathDivide",

        mathEqual = "MathEqual",

        mathNotEqual = "MathNotEqual",

        cornerTabs = "CornerTabs",

        squareTabs = "SquareTabs",

        plaqueTabs = "PlaqueTabs",

        chartX = "ChartX",

        chartStar = "ChartStar",

        chartPlus = "ChartPlus",
    }
    enum ConnectorType {
        straight = "Straight",

        elbow = "Elbow",

        curve = "Curve",
    }
    enum ContentType {
        /**
         * Indicates plain format type of the comment content.
         */
        plain = "Plain",

        /**
         * Comment content containing mentions.
         */
        mention = "Mention",
    }
    enum SpecialCellType {
        /**
         * All cells with conditional formats
         */
        conditionalFormats = "ConditionalFormats",

        /**
         * Cells having validation criteria.
         */
        dataValidations = "DataValidations",

        /**
         * Cells with no content.
         */
        blanks = "Blanks",

        /**
         * Cells containing constants.
         */
        constants = "Constants",

        /**
         * Cells containing formulas.
         */
        formulas = "Formulas",

        /**
         * Cells having the same conditional format as the first cell in the range.
         */
        sameConditionalFormat = "SameConditionalFormat",

        /**
         * Cells having the same data validation criteria as the first cell in the range.
         */
        sameDataValidation = "SameDataValidation",

        /**
         * Cells that are visible.
         */
        visible = "Visible",
    }
    enum SpecialCellValueType {
        /**
         * Cells that have errors, true/false, numeric, or a string value.
         */
        all = "All",

        /**
         * Cells that have errors.
         */
        errors = "Errors",

        /**
         * Cells that have errors, or a true/false value.
         */
        errorsLogical = "ErrorsLogical",

        /**
         * Cells that have errors, or a numeric value.
         */
        errorsNumbers = "ErrorsNumbers",

        /**
         * Cells that have errors, or a string value.
         */
        errorsText = "ErrorsText",

        /**
         * Cells that have errors, true/false, or a numeric value.
         */
        errorsLogicalNumber = "ErrorsLogicalNumber",

        /**
         * Cells that have errors, true/false, or a string value.
         */
        errorsLogicalText = "ErrorsLogicalText",

        /**
         * Cells that have errors, numeric, or a string value.
         */
        errorsNumberText = "ErrorsNumberText",

        /**
         * Cells that have a true/false value.
         */
        logical = "Logical",

        /**
         * Cells that have a true/false, or a numeric value.
         */
        logicalNumbers = "LogicalNumbers",

        /**
         * Cells that have a true/false, or a string value.
         */
        logicalText = "LogicalText",

        /**
         * Cells that have a true/false, numeric, or a string value.
         */
        logicalNumbersText = "LogicalNumbersText",

        /**
         * Cells that have a numeric value.
         */
        numbers = "Numbers",

        /**
         * Cells that have a numeric, or a string value.
         */
        numbersText = "NumbersText",

        /**
         * Cells that have a string value.
         */
        text = "Text",
    }
    /**
     * Specifies the way that an object is attached to its underlying cells.
     */
    enum Placement {
        /**
         * The object is moved with the cells.
         */
        twoCell = "TwoCell",

        /**
         * The object is moved and sized with the cells.
         */
        oneCell = "OneCell",

        /**
         * The object is free floating.
         */
        absolute = "Absolute",
    }
    enum FillPattern {
        none = "None",

        solid = "Solid",

        gray50 = "Gray50",

        gray75 = "Gray75",

        gray25 = "Gray25",

        horizontal = "Horizontal",

        vertical = "Vertical",

        down = "Down",

        up = "Up",

        checker = "Checker",

        semiGray75 = "SemiGray75",

        lightHorizontal = "LightHorizontal",

        lightVertical = "LightVertical",

        lightDown = "LightDown",

        lightUp = "LightUp",

        grid = "Grid",

        crissCross = "CrissCross",

        gray16 = "Gray16",

        gray8 = "Gray8",

        linearGradient = "LinearGradient",

        rectangularGradient = "RectangularGradient",
    }
    /**
     * Specifies the horizontal alignment for the text frame in a shape.
     */
    enum ShapeTextHorizontalAlignment {
        left = "Left",

        center = "Center",

        right = "Right",

        justify = "Justify",

        justifyLow = "JustifyLow",

        distributed = "Distributed",

        thaiDistributed = "ThaiDistributed",
    }
    /**
     * Specifies the vertical alignment for the text frame in a shape.
     */
    enum ShapeTextVerticalAlignment {
        top = "Top",

        middle = "Middle",

        bottom = "Bottom",

        justified = "Justified",

        distributed = "Distributed",
    }
    /**
     * Specifies the vertical overflow for the text frame in a shape.
     */
    enum ShapeTextVerticalOverflow {
        /**
         * Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment).
         */
        overflow = "Overflow",

        /**
         * Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text.
         */
        ellipsis = "Ellipsis",

        /**
         * Hide text that does not fit vertically within the text frame.
         */
        clip = "Clip",
    }
    /**
     * Specifies the horizontal overflow for the text frame in a shape.
     */
    enum ShapeTextHorizontalOverflow {
        overflow = "Overflow",

        clip = "Clip",
    }
    /**
     * Specifies the reading order for the text frame in a shape.
     */
    enum ShapeTextReadingOrder {
        leftToRight = "LeftToRight",

        rightToLeft = "RightToLeft",
    }
    /**
     * Specifies the orientation for the text frame in a shape.
     */
    enum ShapeTextOrientation {
        horizontal = "Horizontal",

        vertical = "Vertical",

        vertical270 = "Vertical270",

        wordArtVertical = "WordArtVertical",

        eastAsianVertical = "EastAsianVertical",

        mongolianVertical = "MongolianVertical",

        wordArtVerticalRTL = "WordArtVerticalRTL",
    }
    /**
     * Determines the type of automatic sizing allowed.
     */
    enum ShapeAutoSize {
        /**
         * No autosizing.
         */
        autoSizeNone = "AutoSizeNone",

        /**
         * The text is adjusted to fit the shape.
         */
        autoSizeTextToFitShape = "AutoSizeTextToFitShape",

        /**
         * The shape is adjusted to fit the text.
         */
        autoSizeShapeToFitText = "AutoSizeShapeToFitText",

        /**
         * A combination of automatic sizing schemes are used.
         */
        autoSizeMixed = "AutoSizeMixed",
    }
    /**
     * Specifies the slicer sort behavior for Slicer.sortBy API.
     */
    enum SlicerSortType {
        /**
         * Sort slicer items in the order provided by the data source.
         */
        dataSourceOrder = "DataSourceOrder",

        /**
         * Sort slicer items in ascending order by item captions.
         */
        ascending = "Ascending",

        /**
         * Sort slicer items in descending order by item captions.
         */
        descending = "Descending",
    }
}
