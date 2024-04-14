// src/SheetsAPI.ts
import { Spreadsheet } from './Spreadsheet'

/**
 * This class provides a set of static methods to interact with Google Sheets.
 * @class
 * @static
 * @hideconstructor
 * @see {@link https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app}
 */
export class SheetsAPI {
  /*****************************************************/
  /* Properties */
  /*****************************************************/
  /**
   * The AutoFillSeries enumeration.
   */
  static AutoFillSeries = {
    DEFAULT_SERIES: SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES,
    ALTERNATE_SERIES: SpreadsheetApp.AutoFillSeries.ALTERNATE_SERIES
  }

  /**
   * The BandingTheme enumeration.
   */
  static BandingTheme = {
    LIGHT_GREY: SpreadsheetApp.BandingTheme.LIGHT_GREY,
    CYAN: SpreadsheetApp.BandingTheme.CYAN,
    GREEN: SpreadsheetApp.BandingTheme.GREEN,
    YELLOW: SpreadsheetApp.BandingTheme.YELLOW,
    ORANGE: SpreadsheetApp.BandingTheme.ORANGE,
    BLUE: SpreadsheetApp.BandingTheme.BLUE,
    TEAL: SpreadsheetApp.BandingTheme.TEAL,
    GREY: SpreadsheetApp.BandingTheme.GREY,
    BROWN: SpreadsheetApp.BandingTheme.BROWN,
    LIGHT_GREEN: SpreadsheetApp.BandingTheme.LIGHT_GREEN,
    INDIGO: SpreadsheetApp.BandingTheme.INDIGO,
    PINK: SpreadsheetApp.BandingTheme.PINK
  }

  /**
   * The BooleanCriteria enumeration.
   */
  static BooleanCriteria = {
    CELL_EMPTY: SpreadsheetApp.BooleanCriteria.CELL_EMPTY,
    CELL_NOT_EMPTY: SpreadsheetApp.BooleanCriteria.CELL_NOT_EMPTY,
    DATE_AFTER: SpreadsheetApp.BooleanCriteria.DATE_AFTER,
    DATE_BEFORE: SpreadsheetApp.BooleanCriteria.DATE_BEFORE,
    DATE_EQUAL_TO: SpreadsheetApp.BooleanCriteria.DATE_EQUAL_TO,
    // DATE_NOT_EQUAL_TO: SpreadsheetApp.BooleanCriteria.DATE_NOT_EQUAL_TO,
    DATE_AFTER_RELATIVE: SpreadsheetApp.BooleanCriteria.DATE_AFTER_RELATIVE,
    DATE_BEFORE_RELATIVE: SpreadsheetApp.BooleanCriteria.DATE_BEFORE_RELATIVE,
    DATE_EQUAL_TO_RELATIVE: SpreadsheetApp.BooleanCriteria.DATE_EQUAL_TO_RELATIVE,
    NUMBER_BETWEEN: SpreadsheetApp.BooleanCriteria.NUMBER_BETWEEN,
    NUMBER_EQUAL_TO: SpreadsheetApp.BooleanCriteria.NUMBER_EQUAL_TO,
    NUMBER_GREATER_THAN: SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN,
    NUMBER_GREATER_THAN_OR_EQUAL_TO: SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO,
    NUMBER_LESS_THAN: SpreadsheetApp.BooleanCriteria.NUMBER_LESS_THAN,
    NUMBER_LESS_THAN_OR_EQUAL_TO: SpreadsheetApp.BooleanCriteria.NUMBER_LESS_THAN_OR_EQUAL_TO,
    NUMBER_NOT_BETWEEN: SpreadsheetApp.BooleanCriteria.NUMBER_NOT_BETWEEN,
    NUMBER_NOT_EQUAL_TO: SpreadsheetApp.BooleanCriteria.NUMBER_NOT_EQUAL_TO,
    TEXT_CONTAINS: SpreadsheetApp.BooleanCriteria.TEXT_CONTAINS,
    TEXT_DOES_NOT_CONTAIN: SpreadsheetApp.BooleanCriteria.TEXT_DOES_NOT_CONTAIN,
    TEXT_EQUAL_TO: SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO,
    // TEXT_NOT_EQUAL_TO: SpreadsheetApp.BooleanCriteria.TEXT_NOT_EQUAL_TO,
    TEXT_STARTS_WITH: SpreadsheetApp.BooleanCriteria.TEXT_STARTS_WITH,
    TEXT_ENDS_WITH: SpreadsheetApp.BooleanCriteria.TEXT_ENDS_WITH,
    CUSTOM_FORMULA: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA
  }

  /**
   * The BorderStyle enumeration.
   */
  static BorderStyle = {
    DOTTED: SpreadsheetApp.BorderStyle.DOTTED,
    DASHED: SpreadsheetApp.BorderStyle.DASHED,
    SOLID: SpreadsheetApp.BorderStyle.SOLID,
    SOLID_MEDIUM: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
    SOLID_THICK: SpreadsheetApp.BorderStyle.SOLID_THICK,
    DOUBLE: SpreadsheetApp.BorderStyle.DOUBLE
  }

  /**
   * The ColorType enumeration.
   */
  static ColorType = {
    UNSUPPORTED: SpreadsheetApp.ColorType.UNSUPPORTED,
    RGB: SpreadsheetApp.ColorType.RGB,
    THEME: SpreadsheetApp.ColorType.THEME
  }

  /**
   * The CopyPasteType enumeration.
   */
  static CopyPasteType = {
    PASTE_NORMAL: SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    PASTE_NO_BORDERS: SpreadsheetApp.CopyPasteType.PASTE_NO_BORDERS,
    PASTE_FORMAT: SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    PASTE_FORMULA: SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
    PASTE_DATA_VALIDATION: SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    PASTE_VALUES: SpreadsheetApp.CopyPasteType.PASTE_VALUES,
    PASTE_CONDITIONAL_FORMATTING: SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING,
    PASTE_COLUMN_WIDTHS: SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS
  }

  /**
   * The DataExecutionErrorCode enumeration.
   */
  static DataExecutionErrorCode = {
    DATA_EXECUTION_ERROR_CODE_UNSUPPORTED:
      SpreadsheetApp.DataExecutionErrorCode.DATA_EXECUTION_ERROR_CODE_UNSUPPORTED,
    NONE: SpreadsheetApp.DataExecutionErrorCode.NONE,
    TIME_OUT: SpreadsheetApp.DataExecutionErrorCode.TIME_OUT,
    TOO_MANY_ROWS: SpreadsheetApp.DataExecutionErrorCode.TOO_MANY_ROWS,
    // TOO_MANY_COLUMNS: SpreadsheetApp.DataExecutionErrorCode.TOO_MANY_COLUMNS,
    TOO_MANY_CELLS: SpreadsheetApp.DataExecutionErrorCode.TOO_MANY_CELLS,
    ENGINE: SpreadsheetApp.DataExecutionErrorCode.ENGINE,
    PARAMETER_INVALID: SpreadsheetApp.DataExecutionErrorCode.PARAMETER_INVALID,
    UNSUPPORTED_DATA_TYPE: SpreadsheetApp.DataExecutionErrorCode.UNSUPPORTED_DATA_TYPE,
    DUPLICATE_COLUMN_NAMES: SpreadsheetApp.DataExecutionErrorCode.DUPLICATE_COLUMN_NAMES,
    INTERRUPTED: SpreadsheetApp.DataExecutionErrorCode.INTERRUPTED,
    OTHER: SpreadsheetApp.DataExecutionErrorCode.OTHER
    // TOO_MANY_CHARS: SpreadsheetApp.DataExecutionErrorCode.TOO_MANY_CHARS,
    // DATA_NOT_FOUND: SpreadsheetApp.DataExecutionErrorCode.DATA_NOT_FOUND,
    // PERMISSION_DENIED: SpreadsheetApp.DataExecutionErrorCode.PERMISSION_DENIED
  }

  /**
   * The DataExecutionState enumeration.
   */
  static DataExecutionState = {
    DATA_EXECUTION_STATE_UNSUPPORTED:
      SpreadsheetApp.DataExecutionState.DATA_EXECUTION_STATE_UNSUPPORTED,
    RUNNING: SpreadsheetApp.DataExecutionState.RUNNING,
    SUCCESS: SpreadsheetApp.DataExecutionState.SUCCESS,
    ERROR: SpreadsheetApp.DataExecutionState.ERROR,
    NOT_STARTED: SpreadsheetApp.DataExecutionState.NOT_STARTED
  }

  /**
   * The DataSourceParameterType enumeration.
   */
  static DataSourceParameterType = {
    DATA_SOURCE_PARAMETER_TYPE_UNSUPPORTED:
      SpreadsheetApp.DataSourceParameterType.DATA_SOURCE_PARAMETER_TYPE_UNSUPPORTED,
    CELL: SpreadsheetApp.DataSourceParameterType.CELL
  }

  /**
   * The DataSourceRefreshScope enumeration.
   */
  // static DataSourceRefreshScope = {
  //   DATA_SOURCE_REFRESH_SCOPE_UNSUPPORTED:
  //     SpreadsheetApp.DataSourceRefreshScope
  //       .DATA_SOURCE_REFRESH_SCOPE_UNSUPPORTED,
  //   DATA_SOURCE: SpreadsheetApp.DataSourceRefreshScope.DATA_SOURCE,
  //   DATA_SOURCE_AND_OBJECTS:
  //     SpreadsheetApp.DataSourceRefreshScope.DATA_SOURCE_AND_OBJECTS
  // }

  /**
   * The DataSourceType enumeration.
   */
  static DataSourceType = {
    DATA_SOURCE_TYPE_UNSUPPORTED: SpreadsheetApp.DataSourceType.DATA_SOURCE_TYPE_UNSUPPORTED,
    BIGQUERY: SpreadsheetApp.DataSourceType.BIGQUERY
  }

  /**
   * The DataValidationCriteria enumeration.
   */
  static DataValidationCriteria = {
    DATE_AFTER: SpreadsheetApp.DataValidationCriteria.DATE_AFTER,
    DATE_BEFORE: SpreadsheetApp.DataValidationCriteria.DATE_BEFORE,
    DATE_BETWEEN: SpreadsheetApp.DataValidationCriteria.DATE_BETWEEN,
    DATE_EQUAL_TO: SpreadsheetApp.DataValidationCriteria.DATE_EQUAL_TO,
    DATE_IS_VALID_DATE: SpreadsheetApp.DataValidationCriteria.DATE_IS_VALID_DATE,
    DATE_NOT_BETWEEN: SpreadsheetApp.DataValidationCriteria.DATE_NOT_BETWEEN,
    DATE_ON_OR_AFTER: SpreadsheetApp.DataValidationCriteria.DATE_ON_OR_AFTER,
    DATE_ON_OR_BEFORE: SpreadsheetApp.DataValidationCriteria.DATE_ON_OR_BEFORE,
    NUMBER_BETWEEN: SpreadsheetApp.DataValidationCriteria.NUMBER_BETWEEN,
    NUMBER_EQUAL_TO: SpreadsheetApp.DataValidationCriteria.NUMBER_EQUAL_TO,
    NUMBER_GREATER_THAN: SpreadsheetApp.DataValidationCriteria.NUMBER_GREATER_THAN,
    NUMBER_GREATER_THAN_OR_EQUAL_TO:
      SpreadsheetApp.DataValidationCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO,
    NUMBER_LESS_THAN: SpreadsheetApp.DataValidationCriteria.NUMBER_LESS_THAN,
    NUMBER_LESS_THAN_OR_EQUAL_TO:
      SpreadsheetApp.DataValidationCriteria.NUMBER_LESS_THAN_OR_EQUAL_TO,
    NUMBER_NOT_BETWEEN: SpreadsheetApp.DataValidationCriteria.NUMBER_NOT_BETWEEN,
    NUMBER_NOT_EQUAL_TO: SpreadsheetApp.DataValidationCriteria.NUMBER_NOT_EQUAL_TO,
    TEXT_CONTAINS: SpreadsheetApp.DataValidationCriteria.TEXT_CONTAINS,
    TEXT_DOES_NOT_CONTAIN: SpreadsheetApp.DataValidationCriteria.TEXT_DOES_NOT_CONTAIN,
    TEXT_EQUAL_TO: SpreadsheetApp.DataValidationCriteria.TEXT_EQUAL_TO,
    TEXT_IS_VALID_EMAIL: SpreadsheetApp.DataValidationCriteria.TEXT_IS_VALID_EMAIL,
    TEXT_IS_VALID_URL: SpreadsheetApp.DataValidationCriteria.TEXT_IS_VALID_URL,
    VALUE_IN_LIST: SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST,
    VALUE_IN_RANGE: SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE,
    CUSTOM_FORMULA: SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA,
    CHECKBOX: SpreadsheetApp.DataValidationCriteria.CHECKBOX
  }

  /**
   * The DateTimeGroupingRuleType enumeration.
   */
  // static DateTimeGroupingRuleType = {
  //   UNSUPPORTED: SpreadsheetApp.DateTimeGroupingRuleType.UNSUPPORTED,
  //   SECOND: SpreadsheetApp.DateTimeGroupingRuleType.SECOND,
  //   MINUTE: SpreadsheetApp.DateTimeGroupingRuleType.MINUTE,
  //   HOUR: SpreadsheetApp.DateTimeGroupingRuleType.HOUR,
  //   HOUR_MINUTE: SpreadsheetApp.DateTimeGroupingRuleType.HOUR_MINUTE,
  //   HOUR_MINUTE_AMPM: SpreadsheetApp.DateTimeGroupingRuleType.HOUR_MINUTE_AMPM,
  //   DAY_OF_WEEK: SpreadsheetApp.DateTimeGroupingRuleType.DAY_OF_WEEK,
  //   DAY_OF_YEAR: SpreadsheetApp.DateTimeGroupingRuleType.DAY_OF_YEAR,
  //   DAY_OF_MONTH: SpreadsheetApp.DateTimeGroupingRuleType.DAY_OF_MONTH,
  //   DAY_MONTH: SpreadsheetApp.DateTimeGroupingRuleType.DAY_MONTH,
  //   MONTH: SpreadsheetApp.DateTimeGroupingRuleType.MONTH,
  //   QUARTER: SpreadsheetApp.DateTimeGroupingRuleType.QUARTER,
  //   YEAR: SpreadsheetApp.DateTimeGroupingRuleType.YEAR,
  //   YEAR_MONTH: SpreadsheetApp.DateTimeGroupingRuleType.YEAR_MONTH,
  //   YEAR_QUARTER: SpreadsheetApp.DateTimeGroupingRuleType.YEAR_QUARTER,
  //   YEAR_MONTH_DAY: SpreadsheetApp.DateTimeGroupingRuleType.YEAR_MONTH_DAY
  // }

  /**
   * The DeveloperMetadataLocationType enumeration.
   */
  static DeveloperMetadataLocationType = {
    SPREADSHEET: SpreadsheetApp.DeveloperMetadataLocationType.SPREADSHEET,
    SHEET: SpreadsheetApp.DeveloperMetadataLocationType.SHEET,
    ROW: SpreadsheetApp.DeveloperMetadataLocationType.ROW,
    COLUMN: SpreadsheetApp.DeveloperMetadataLocationType.COLUMN
  }

  /**
   * The DeveloperMetadataVisibility enumeration.
   */
  static DeveloperMetadataVisibility = {
    DOCUMENT: SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT,
    PROJECT: SpreadsheetApp.DeveloperMetadataVisibility.PROJECT
  }

  /**
   * The Dimension enumeration.
   */
  static Dimension = {
    COLUMNS: SpreadsheetApp.Dimension.COLUMNS,
    ROWS: SpreadsheetApp.Dimension.ROWS
  }

  /**
   * The Direction enumeration.
   */
  static Direction = {
    UP: SpreadsheetApp.Direction.UP,
    DOWN: SpreadsheetApp.Direction.DOWN,
    PREVIOUS: SpreadsheetApp.Direction.PREVIOUS,
    NEXT: SpreadsheetApp.Direction.NEXT
  }

  /**
   * The FrequencyType enumeration.
   */
  // static FrequencyType = {
  //   FREQUENCY_TYPE_UNSUPPORTED: SpreadsheetApp.FrequencyType.FREQUENCY_TYPE_UNSUPPORTED,
  //   DAYLY: SpreadsheetApp.FrequencyType.DAYLY,
  //   WEEKLY: SpreadsheetApp.FrequencyType.WEEKLY,
  //   MONTHLY: SpreadsheetApp.FrequencyType.MONTHLY
  // }

  /**
   * The GroupControlTogglePosition enumeration.
   */
  static GroupControlTogglePosition = {
    BEFORE: SpreadsheetApp.GroupControlTogglePosition.BEFORE,
    AFTER: SpreadsheetApp.GroupControlTogglePosition.AFTER
  }

  /**
   * The InterpolationType enumeration.
   */
  static InterpolationType = {
    NUMBER: SpreadsheetApp.InterpolationType.NUMBER,
    PERCENT: SpreadsheetApp.InterpolationType.PERCENT,
    PERCENTILE: SpreadsheetApp.InterpolationType.PERCENTILE,
    MIN: SpreadsheetApp.InterpolationType.MIN,
    MAX: SpreadsheetApp.InterpolationType.MAX
  }

  /**
   * The PivotTableSummarizeFunction enumeration.
   */
  static PivotTableSummarizeFunction = {
    CUSTOM: SpreadsheetApp.PivotTableSummarizeFunction.CUSTOM,
    SUM: SpreadsheetApp.PivotTableSummarizeFunction.SUM,
    COUNTA: SpreadsheetApp.PivotTableSummarizeFunction.COUNTA,
    COUNT: SpreadsheetApp.PivotTableSummarizeFunction.COUNT,
    COUNTUNIQUE: SpreadsheetApp.PivotTableSummarizeFunction.COUNTUNIQUE,
    AVERAGE: SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE,
    MAX: SpreadsheetApp.PivotTableSummarizeFunction.MAX,
    MIN: SpreadsheetApp.PivotTableSummarizeFunction.MIN,
    MEDIAN: SpreadsheetApp.PivotTableSummarizeFunction.MEDIAN,
    PRODUCT: SpreadsheetApp.PivotTableSummarizeFunction.PRODUCT,
    STDEV: SpreadsheetApp.PivotTableSummarizeFunction.STDEV,
    STDEVP: SpreadsheetApp.PivotTableSummarizeFunction.STDEVP,
    VAR: SpreadsheetApp.PivotTableSummarizeFunction.VAR,
    VARP: SpreadsheetApp.PivotTableSummarizeFunction.VARP
  }

  /**
   * The PivotValueDisplayType enumeration.
   */
  static PivotValueDisplayType = {
    DEFAULT: SpreadsheetApp.PivotValueDisplayType.DEFAULT,
    PERCENT_OF_ROW_TOTAL: SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_ROW_TOTAL,
    PERCENT_OF_COLUMN_TOTAL: SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_COLUMN_TOTAL,
    PERCENT_OF_GRAND_TOTAL: SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_GRAND_TOTAL
  }

  /**
   * The ProtectionType enumeration.
   */
  static ProtectionType = {
    RANGE: SpreadsheetApp.ProtectionType.RANGE,
    SHEET: SpreadsheetApp.ProtectionType.SHEET
  }

  /**
   * The RecalculationInterval enumeration.
   */
  static RecalculationInterval = {
    ON_CHANGE: SpreadsheetApp.RecalculationInterval.ON_CHANGE,
    MINUTE: SpreadsheetApp.RecalculationInterval.MINUTE,
    HOUR: SpreadsheetApp.RecalculationInterval.HOUR
  }

  /**
   * The RelativeDate enumeration.
   */
  static RelativeDate = {
    TODAY: SpreadsheetApp.RelativeDate.TODAY,
    TOMORROW: SpreadsheetApp.RelativeDate.TOMORROW,
    YESTERDAY: SpreadsheetApp.RelativeDate.YESTERDAY,
    PAST_WEEK: SpreadsheetApp.RelativeDate.PAST_WEEK,
    PAST_MONTH: SpreadsheetApp.RelativeDate.PAST_MONTH,
    PAST_YEAR: SpreadsheetApp.RelativeDate.PAST_YEAR
  }

  /**
   * The SheetType enumeration.
   */
  static SheetType = {
    GRID: SpreadsheetApp.SheetType.GRID,
    OBJECT: SpreadsheetApp.SheetType.OBJECT
    // DATASOURCE: SpreadsheetApp.SheetType.DATASOURCE
  }

  /**
   * The SortOrder enumeration.
   */
  // static SortOrder = {
  //   ASCENDING: SpreadsheetApp.SortOrder.ASCENDING,
  //   DESCENDING: SpreadsheetApp.SortOrder.DESCENDING
  // }

  /**
   * The TextDirection enumeration.
   */
  static TextDirection = {
    LEFT_TO_RIGHT: SpreadsheetApp.TextDirection.LEFT_TO_RIGHT,
    RIGHT_TO_LEFT: SpreadsheetApp.TextDirection.RIGHT_TO_LEFT
  }

  /**
   * The TextToColumnsDelimiter enumeration.
   */
  static TextToColumnsDelimiter = {
    COMMA: SpreadsheetApp.TextToColumnsDelimiter.COMMA,
    SEMICOLON: SpreadsheetApp.TextToColumnsDelimiter.SEMICOLON,
    PERIOD: SpreadsheetApp.TextToColumnsDelimiter.PERIOD,
    SPACE: SpreadsheetApp.TextToColumnsDelimiter.SPACE
  }

  /**
   * The ThemeColorType enumeration.
   */
  static ThemeColorType = {
    UNSUPPORTED: SpreadsheetApp.ThemeColorType.UNSUPPORTED,
    TEXT: SpreadsheetApp.ThemeColorType.TEXT,
    BACKGROUND: SpreadsheetApp.ThemeColorType.BACKGROUND,
    ACCENT1: SpreadsheetApp.ThemeColorType.ACCENT1,
    ACCENT2: SpreadsheetApp.ThemeColorType.ACCENT2,
    ACCENT3: SpreadsheetApp.ThemeColorType.ACCENT3,
    ACCENT4: SpreadsheetApp.ThemeColorType.ACCENT4,
    ACCENT5: SpreadsheetApp.ThemeColorType.ACCENT5,
    ACCENT6: SpreadsheetApp.ThemeColorType.ACCENT6,
    HYPERLINK: SpreadsheetApp.ThemeColorType.HYPERLINK
  }

  /**
   * The ValueType enumeration.
   */
  static ValueType = {
    IMAGE: SpreadsheetApp.ValueType.IMAGE
  }

  /**
   * The WrapStrategy enumeration.
   */
  static WrapStrategy = {
    WRAP: SpreadsheetApp.WrapStrategy.WRAP,
    OVERFLOW: SpreadsheetApp.WrapStrategy.OVERFLOW,
    CLIP: SpreadsheetApp.WrapStrategy.CLIP
  }

  /*****************************************************/
  /* Static methods */
  /*****************************************************/

  /**
   * Creates a new spreadsheet with a given name.
   * @param {string} name - The name of the new spreadsheet.
   * @returns {Spreadsheet} The new spreadsheet.
   */
  static create(name: string): Spreadsheet

  /**
   * Creates a new spreadsheet with a given name, rows, and columns.
   * @param {string} name - The name of the new spreadsheet.
   * @param {number} rows - The number of rows in the new spreadsheet.
   * @param {number} columns - The number of columns in the new spreadsheet.
   * @returns {Spreadsheet} The new spreadsheet.
   */
  static create(name: string, rows: number, columns: number): Spreadsheet

  /**
   * Concrete implementation of the create method.
   * @param name
   * @param rows
   * @param columns
   * @returns
   */
  static create(name: string, rows?: number, columns?: number): Spreadsheet {
    if (rows != null && columns != null) {
      const spreadsheet = SpreadsheetApp.create(name, rows, columns)
      if (!spreadsheet) {
        throw new Error('Failed to create a new spreadsheet with specified rows and columns.')
      }
      return new Spreadsheet(spreadsheet)
    } else {
      const spreadsheet = SpreadsheetApp.create(name)
      if (!spreadsheet) {
        throw new Error('Failed to create a new spreadsheet.')
      }
      return new Spreadsheet(spreadsheet)
    }
  }

  /**
   * Enables the execution of all data sources in the spreadsheet.
   */
  static enableAllDataSourcesExecution(): void {
    SpreadsheetApp.enableAllDataSourcesExecution()
  }

  /**
   * Enables the execution of BigQuery queries in the spreadsheet.
   */
  static enableBigQueryExecution(): void {
    SpreadsheetApp.enableBigQueryExecution()
  }

  /**
   * Flushes all pending Spreadsheet changes.
   */
  static flush(): void {
    SpreadsheetApp.flush()
  }

  /**
   * Returns the active spreadsheet.
   * @returns {Spreadsheet} The active spreadsheet.
   */
  static getActive(): Spreadsheet {
    const spreadsheet = SpreadsheetApp.getActive()
    if (!spreadsheet) {
      throw new Error('No active spreadsheet found.')
    }
    return new Spreadsheet(spreadsheet)
  }

  /**
   * Returns the active range.
   * @returns {GoogleAppsScript.Spreadsheet.Range} The active range.
   */
  static getActiveRange(): GoogleAppsScript.Spreadsheet.Range {
    const range = SpreadsheetApp.getActiveRange()
    if (!range) {
      throw new Error('No active range found.')
    }
    return range
  }

  /**
   * Returns the active range list.
   * @returns {GoogleAppsScript.Spreadsheet.RangeList} The active range list.
   */
  static getActiveRangeList(): GoogleAppsScript.Spreadsheet.RangeList {
    const rangeList = SpreadsheetApp.getActiveRangeList()
    if (!rangeList) {
      throw new Error('No active range list found.')
    }
    return rangeList
  }

  /**
   * Returns the active sheet.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet.
   */
  static getActiveSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const sheet = SpreadsheetApp.getActiveSheet()
    if (!sheet) {
      throw new Error('No active sheet found.')
    }
    return sheet
  }

  /**
   * Returns the active spreadsheet.
   * @returns {Spreadsheet} The active spreadsheet.
   */
  static getActiveSpreadsheet(): Spreadsheet {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    if (!spreadsheet) {
      throw new Error('No active spreadsheet found.')
    }
    return new Spreadsheet(spreadsheet)
  }

  /**
   * Returns the current cell in the active sheet.
   * @returns {GoogleAppsScript.Spreadsheet.Range} The current cell.
   */
  static getCurrentCell(): GoogleAppsScript.Spreadsheet.Range {
    return SpreadsheetApp.getCurrentCell()
  }

  /**
   * Returns the selection in the active spreadsheet.
   * @returns {GoogleAppsScript.Spreadsheet.Selection} The selection.
   */
  static getSelection(): GoogleAppsScript.Spreadsheet.Selection {
    return SpreadsheetApp.getSelection()
  }

  /**
   * Returns the UI instance for the active spreadsheet.
   * @returns {GoogleAppsScript.Base.Ui} The UI instance.
   */
  static getUi(): GoogleAppsScript.Base.Ui {
    return SpreadsheetApp.getUi()
  }

  static newCellImage(): void {
    throw new Error('Method not implemented.')
  }

  static newColor(): void {
    throw new Error('Method not implemented.')
  }

  static newConditionalFormatRule(): void {
    throw new Error('Method not implemented.')
  }

  static newDataSourceSpec(): void {
    throw new Error('Method not implemented.')
  }

  static newDataValidation(): void {
    throw new Error('Method not implemented.')
  }

  static newFilterCriteria(): void {
    throw new Error('Method not implemented.')
  }

  static newRichTextValue(): void {
    throw new Error('Method not implemented.')
  }

  static newTextStyle(): void {
    throw new Error('Method not implemented.')
  }

  /**
   * Opens a spreadsheet by its name.
   * @param {GoogleAppsScript.Drive.File} name - The name of the spreadsheet to open.
   * @returns {Spreadsheet} The opened spreadsheet.
   */
  static open(name: GoogleAppsScript.Drive.File): Spreadsheet {
    const spreadsheet = SpreadsheetApp.open(name)
    if (!spreadsheet) {
      throw new Error('No spreadsheet found with the given name.')
    }
    return new Spreadsheet(spreadsheet)
  }

  /**
   * Opens a spreadsheet by its ID.
   * @param {string} id - The ID of the spreadsheet to open.
   * @returns {Spreadsheet} The opened spreadsheet.
   */
  static openById(id: string): Spreadsheet {
    const spreadsheet = SpreadsheetApp.openById(id)
    if (!spreadsheet) {
      throw new Error('No spreadsheet found with the given ID.')
    }
    return new Spreadsheet(spreadsheet)
  }

  /**
   * Opens a spreadsheet by its URL.
   * @param {string} url - The URL of the spreadsheet to open.
   * @returns {Spreadsheet} The opened spreadsheet.
   */
  static openByUrl(url: string): Spreadsheet {
    const spreadsheet = SpreadsheetApp.openByUrl(url)
    if (!spreadsheet) {
      throw new Error('No spreadsheet found with the given URL.')
    }
    return new Spreadsheet(spreadsheet)
  }

  /**
   * Sets the active range.
   * @param range
   */
  static setActiveRange(range: GoogleAppsScript.Spreadsheet.Range): void {
    SpreadsheetApp.setActiveRange(range)
  }

  /**
   * Sets the active range list.
   * @param rangeList
   */
  static setActiveRangeList(rangeList: GoogleAppsScript.Spreadsheet.RangeList): void {
    SpreadsheetApp.setActiveRangeList(rangeList)
  }

  /**
   * Sets the active sheet.
   * @param sheet
   */
  static setActiveSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    SpreadsheetApp.setActiveSheet(sheet)
  }

  /**
   * Sets the active spreadsheet.
   * @param newActiveSpreadsheet
   */
  static setActiveSpreadsheet(
    newActiveSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
  ): void {
    SpreadsheetApp.setActiveSpreadsheet(newActiveSpreadsheet)
  }

  /**
   * Sets the current cell in the active sheet.
   * @param cell
   */
  static setCurrentCell(cell: GoogleAppsScript.Spreadsheet.Range): void {
    SpreadsheetApp.setCurrentCell(cell)
  }
}
