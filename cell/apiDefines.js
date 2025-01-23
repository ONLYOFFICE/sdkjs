/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
// Используем [] вместо new Array() для ускорения (http://jsperf.com/creation-array)
// Используем {} вместо new Object() для ускорения (http://jsperf.com/creation-object)

  // Import
  var CColor = AscCommon.CColor;

var c_oAscMergeOptions = {
  Disabled: -1,
  None: 0,
  Merge: 1,
  MergeCenter: 2,
  MergeAcross: 3
};

var c_oAscSortOptions = {
  Ascending: 1,
  Descending: 2,
  ByColorFill: 3,
  ByColorFont: 4,
  ByIcon: 5,
  ByValue: 6
};

var c_oAscBorderOptions = {
  Top: 0,
  Right: 1,
  Bottom: 2,
  Left: 3,
  DiagD: 4,
  DiagU: 5,
  InnerV: 6,
  InnerH: 7
};

var c_oAscCleanOptions = {
  All: 0,
  Text: 1,
  Format: 2,
  Formula: 4,
  Comments: 5,
  Hyperlinks: 6,
  Sparklines: 7,
  SparklineGroups: 8
};

// ToDo Add including cells not marked as needing to be calculated
var c_oAscCalculateType = {
  WorkbookOnlyChanged: 1,
  ActiveSheet: 2,
  Workbook: 3,
  All: 4
};

var c_oAscDrawDepOptions = {
  Master: 0,
  Slave: 1,
  Clear: 2
};

var c_oAscSelectionDialogType = {
  None: 0,
  FormatTable: 1,
  Chart: 2,
  FormatTableChangeRange: 4,
  CustomSort: 5,
  PivotTableData: 6,
  PivotTableReport: 7,
  PrintTitles: 8,
  Function: 9,
  DataValidation: 10,
  ConditionalFormattingRule: 11,
  ImportXml: 12,
  GoalSeek_Cell: 13,
  GoalSeek_ChangingCell: 14
};

var c_oAscScrollType = {
  ScrollVertical: 1,
  ScrollHorizontal: 2,
  ScrollInitRowsColsCount: 4
};

var c_oAscHyperlinkType = {
  WebLink: 1,
  RangeLink: 2,
  FileLink: 3
};

var c_oAscMouseMoveType = {
  None: 0,
  Hyperlink: 1,
  Comment: 2,
  LockedObject: 3,
  ResizeColumn: 4,
  ResizeRow: 5,
  Filter: 6,
  Tooltip: 7,
  ForeignSelect: 8,
  Eyedropper: 9
};

var c_oAscMouseMoveLockedObjectType = {
  None: -1,
  Range: 0,
  TableProperties: 1,
  Sheet: 2
};

var c_oAscLockTypeElem = {
  Range: 1,
  Object: 2,
  Sheet: 3
};

var c_oAscLockTypeElemSubType = {
  DeleteColumns: 1,
  InsertColumns: 2,
  DeleteRows: 3,
  InsertRows: 4,
  ChangeProperties: 5,
  DefinedNames: 6,
  NamedSheetView: 7,
  UserProtectedRange: 8
};

var c_oAscRecalcIndexTypes = {
  RecalcIndexAdd: 1,
  RecalcIndexRemove: 2
};

/** @enum */
var c_oAscCustomAutoFilter = {
  equals: 1,
  isGreaterThan: 2,
  isGreaterThanOrEqualTo: 3,
  isLessThan: 4,
  isLessThanOrEqualTo: 5,
  doesNotEqual: 6,
  beginsWith: 7,
  doesNotBeginWith: 8,
  endsWith: 9,
  doesNotEndWith: 10,
  contains: 11,
  doesNotContain: 12
};

    /** @enum */
var c_oAscDynamicAutoFilter = {
    aboveAverage: 1,
    belowAverage: 2,
    lastMonth: 3,
    lastQuarter: 4,
    lastWeek: 5,
    lastYear: 6,
    m1: 7,
    m10: 8,
    m11: 9,
    m12: 10,
    m2: 11,
    m3: 12,
    m4: 13,
    m5: 14,
    m6: 15,
    m7: 16,
    m8: 17,
    m9: 18,
    nextMonth: 19,
    nextQuarter: 20,
    nextWeek: 21,
    nextYear: 22,
    nullType: 23,
    q1: 24,
    q2: 25,
    q3: 26,
    q4: 27,
    thisMonth: 28,
    thisQuarter: 29,
    thisWeek: 30,
    thisYear: 31,
    today: 32,
    tomorrow: 33,
    yearToDate: 34,
    yesterday: 35
};

var c_oAscTop10AutoFilter = {
    max: 1,
    min: 2
};

var c_oAscChangeFilterOptions = {
  filter: 1,
  style: 2
};

var c_oAscChangeSelectionFormatTable = {
	all: 1,
	data: 2,
	row: 3,
	column: 4,
	dataColumn: 5
};

var c_oAscChangeTableStyleInfo = {
	columnFirst: 1,
	columnLast: 2,
	columnBanded: 3,
	rowHeader: 4,
	rowTotal: 5,
	rowBanded: 6,
	filterButton: 7,
	advancedSettings: 8
};

// Состояние редактора ячейки
var c_oAscCellEditorState = {
  editEnd: 0,				// Окончание редактирования
  editStart: 1,				// Начало редактирования
  editEmptyCell: 2,			// Редактирование пустой ячейки (доступны функции и свойства текста)
  editText: 3,				// Редактирование текста, числа, даты и др. формата, кроме формулы
  editFormula: 4,			// Редактирование формулы
  editInFormulaBar: 5,		// Редактирование в строке формул
  editInCell: 6				// Редактирование в ячейке
};

// Состояние select-а
var c_oAscCellEditorSelectState = {
  no    : 0,
  char  : 1,
  word  : 2
};

// Пересчитывать ли ширину столбца
var c_oAscCanChangeColWidth = {
  none: 0,	// not recalc
  numbers: 1,	// only numbers
  all: 2	// numbers + text
};

// Merge cell type
var c_oAscMergeType = {
  none: 0,
  cols: 1, // Замержены ли колонки (если да, то автоподбор ширины не должен работать)
  rows: 2  // Замержены ли строки (если да, то автоподбор высоты не должен работать)
};

var c_oAscPaneState = {
  Frozen: "frozen",
  FrozenSplit: "frozenSplit",
  Split: "split"
};

var c_oAscFindLookIn = {
  Formulas: 1,
  Value: 2,
  Annotations: 3
};

var c_oTargetType = {
  None: 0,
  ColumnResize: 1,
  RowResize: 2,
  FillHandle: 3,
  MoveRange: 4,
  MoveResizeRange: 5,
  FilterObject: 6,
  ColumnHeader: 7,
  RowHeader: 8,
  Corner: 9,
  Hyperlink: 10,
  Cells: 11,
  Shape: 12,
  FrozenAnchorH: 14,
  FrozenAnchorV: 15,
  GroupRow: 16,
  GroupCol: 17,
  TableSelectionChange: 18,
  Placeholder: 19,
  ColumnRowHeaderMove: 20,
  TraceDependents: 21
};

var c_oAscAutoFilterTypes = {
  ColorFilter: 0,
  CustomFilters: 1,
  DynamicFilter: 2,
  Top10: 3,
  Filters: 4,
  None: 5
};

var c_oAscCoAuthoringMeBorderColor = new CColor(22, 156, 0);
var c_oAscCoAuthoringOtherBorderColor = new CColor(238, 53, 37);
var c_oAscCoAuthoringLockTablePropertiesBorderColor = new CColor(255, 144, 0);
var c_oAscCoAuthoringDottedWidth = 4;
var c_oAscCoAuthoringDottedDistance = 2;

var c_oAscFormulaRangeBorderColor = [
  new CColor(95, 140, 237),
  new CColor(235, 94, 96),
  new CColor(141, 97, 194),
  new CColor(45, 150, 57),
  new CColor(191, 76, 145),
  new CColor(227, 130, 34),
  new CColor(55, 127, 158)
];

var c_oAscVisibleAreaOleEditorBorderColor = new CColor(32, 139, 255);

  var selectionLineType = {
    None        : 0,
    Selection   : 1,
    ActiveCell  : 2,
    Resize      : 4,
    Promote     : 8,
    Dash        : 16,
    DashThick   : 32,
    ResizeRange : 64,
    NotStroke   : 128
  };

  var docChangedType = {
    cellValue: 0,
    rangeValues: 1,
    sheetContent: 2,
    sheetRemove: 3,
    sheetRename: 4,
    sheetChangeIndex: 5,
    markModifiedSearch: 6,
    mergeRange: 7,
    removeRows: 8
  };

  var c_oAscLockNameFrozenPane = "frozenPane";
  var c_oAscLockNameTabColor = "tabColor";
  var c_oAscLockAddSheet = "addSheet";
  var c_oAscLockLayoutOptions = "layoutOptions";
  var c_oAscHeaderFooterEdit = "headerFooterEdit";
  var c_oAscLockPrintScaleOptions = "printScaleOptions";

var c_oAscGetDefinedNamesList = {
  Worksheet: 0,
  WorksheetWorkbook: 1,
  All: 2
};

var c_oAscDefinedNameReason = {
  WrongName: -1,
  IsLocked: -2,
  Existed: -3,
  LockDefNameManager: -4,
  NameReserved: -5,
  OK: 0
};

var c_oAscPopUpSelectorType = {
  None: 0,
  Func: 1,
  Range: 2,
  Table: 3,
  Slicer: 4,
  TotalRowFunc: 5,
  TableColumnName: 6,
  TableThisRow: 7,
  TableAll: 8,
  TableData: 9,
  TableHeaders: 10,
  TableTotals: 11
};
  /** @enum */
  var c_oSerFormat = {
    Version		: 2, //1.0.0.2
    Signature	: "XLSY"
  };

  var c_oAscSparklineType = {
    Line: 0,
    Column: 1,
    Stacked: 2
  };
  var c_oAscEDispBlanksAs = {
    Span: 0,
    Gap: 1,
    Zero: 2
  };
  var c_oAscSparklineAxisMinMax = {
    Individual: 0,
    Group: 1,
    Custom: 2
  };
  var c_oAscAutoCorrectOptions = {
	UndoTableAutoExpansion: 0,
	RedoTableAutoExpansion: 1
  };

  //изменяем Print_Area
  var c_oAscChangePrintAreaType = {
      set: 0,
      clear: 1,
      add: 2,
      change: 3
  };

  //поля header/footer
  var c_oAscHeaderFooterField = {
      pageNumber: 0,
      pageCount: 1,
      sheetName: 2,
      fileName: 3,
      filePath: 4,
      date: 5,
      time: 6,
      lineBreak: 7,
	  picture: 8,
	  text: 9
  };

  var c_oAscPageHFType = {
      firstHeader: 0,
      oddHeader: 1,
      evenHeader: 2,
      firstFooter: 3,
      oddFooter: 4,
      evenFooter: 5
  };

  var c_oAscHeaderFooterType = {
      first: 0,
      odd: 1,
      even: 2
  };

  var c_oAscHeaderFooterPresets = {
      none: 0,
      page: 1,
      pageOfQuestion: 2,
      sheet: 3,
      confidential: 4,
      bookName: 5,
      //bookPath: 5
      sheetPage: 6,
      sheetConfidentialPage: 7,
      bookNamePage: 8,
      pageSheet: 9,
      pageBook: 10,
      //bookPathPage: 11;
      pageBookName: 11,
      userPageDate: 12,
      //bookPathPagePathFile: 12;
      preparedUserDatePage: 13,
      custom: 14
  };

  var c_oAscPrintTitlesRangeType = {
      first: 0,
      frozen: 1,
      current: 2
  };

  var c_oAscFormulaArgumentType = {
    number: 0,
    text: 1,
    reference: 2,
    any: 3,
    logical: 4,
    array: 5
  };

  var c_oAscSelectionForCFType = {
    selection: 1,
    worksheet: 2,
    table: 3,
    pivot: 4
  };

  var c_kMaxPrintPages = 1500;

  var c_oAscFrozenPaneBorderType = {
    shadow: 1,
    line: 2
  };

  var c_oAscCFRuleTypeSettings = {
    dataBar: 1,
    colorScale: 2,
    icons: 3,
    format: 4
  };
  
  var c_oAscSheetProtectType = {
    objects: 1,
    scenarios: 2,
    formatCells: 3,
    formatColumns: 4,
    formatRows: 5,
    insertColumns: 6,
    insertRows: 7,
    insertHyperlinks: 8,
    deleteColumns: 9,
    deleteRows: 10,
    selectLockedCells: 11,
    sort: 12,
    autoFilter: 13,
    pivotTables: 14,
    selectUnlockedCells: 15
  };

  var c_oAscFrozenPaneAddType = {
    firstRow: 1,
    firstCol: 2
  };

  var ETableType = {
	  queryTable: 0,
	  worksheet: 1,
	  xml: 2
  };

	var c_oAscWorkbookProtectType = {
		lockStructure: 0,
		lockWindows: 1,
		lockRevisions: 2
	};

	var c_oAscSelectionSortExpand = {
		expandAndNotShowMessage: 0,
		notExpandAndNotShowMessage: 1,
		showExpandMessage: 2,
		showLockMessage: 3
	};

  var c_oAscSearchBy = {
    Workbook: 0,
    Sheet: 1,
    Range: 2
  };

  var c_nAscMaxAddCellWatchesCount = 10000;
  var c_oAscExternalReferenceType = {
    referenceData: 0,
    link: 1,
    path: 2
  };

  var c_oAscPageBreaksDisableType = {
    none: 0,
    all: 1,
    insertRemove: 2,
    reset: 3
  };

  var c_oAscSeriesInType = {
    rows: 0,
    columns: 1
  };

  var c_oAscSeriesType = {
    linear: 0,
    growth: 1,
    date: 2,
    autoFill: 3
  };

  var c_oAscDateUnitType = {
    day: 0,
    weekday: 1,
    month: 2,
    year: 3
  };

  var c_oAscRemoveArrowsType = {
    all: 0,
    precedent: 1,
    dependent: 2
  };

  var c_oAscFillType = {
    copyCells: 0,
    fillSeries: 1,
    fillFormattingOnly: 2,
    fillWithoutFormatting: 3,
    fillDays: 4,
    fillWeekdays: 5,
    fillMonths: 6,
    fillYears: 7,
    linearTrend: 8,
    growthTrend: 9,
    flashFill: 10,
    series: 11,
    fillDown: 12,
    fillRight: 13,
    fillUp: 14,
    fillLeft: 15,
    justify: 16
  };

  /** @enum {number} */
  var c_oAscContextMenuTypes = {
    common       : 0, // default context menu
    changeSeries : 1  // fill right click mouse - series menu
  };

  var c_oAscCalcMode = {
    auto: 0,
    autoNoTable: 1,
    manual: 2
  };

  const c_oAscCellShortcutType = {
    refreshAllConnections     : 0,
    refreshSelectedConnections: 1,
    changeFormatTableInfo     : 2,
    calculateAll              : 3,
    calculateActiveSheet      : 5,
    focusOnCellEditor         : 7,
    addDate                   : 8,
    addTime                   : 9,
    removeActiveCell          : 10,
    emptyRange                : 11,
    moveActiveCellToLeft      : 12,
    moveActiveCellToRight     : 13,
    moveActiveCellToDown      : 14,
    moveActiveCellToUp        : 15,
    reset                     : 16,
    disableNumLock            : 17,
    disableScrollLock         : 18,
    selectColumn              : 19,
    selectRow                 : 20,
    selectSheet               : 21,
    addSeparator              : 22,
    goToPreviousSheet         : 23,
    moveToTopCell             : 24,
    moveToNextSheet           : 25,
    moveToLeftEdgeCell        : 26,
    selectToLeftEdgeCell      : 27,
    moveToLeftCell            : 28,
    selectToLeftCell          : 29,
    moveToRightEdgeCell       : 30,
    selectToRightEdgeCell     : 31,
    moveToRightCell           : 32,
    selectToRightCell         : 33,
    selectToTopCell           : 34,
    moveToUpCell              : 35,
    selectToUpCell            : 36,
    moveToBottomCell          : 37,
    selectToBottomCell        : 38,
    moveToDownCell            : 39,
    selectToDownCell          : 40,
    moveToFirstColumn         : 41,
    selectToFirstColumn       : 42,
    moveToLeftEdgeTop         : 43,
    selectToLeftEdgeTop       : 44,
    moveToRightBottomEdge     : 45,
    selectToRightBottomEdge   : 46,
    setNumberFormat           : 47,
    setTimeFormat             : 48,
    setDateFormat             : 49,
    setCurrencyFormat         : 50,
    setPercentFormat          : 51,
    setStrikethrough          : 52,
    setExponentialFormat      : 53,
    setBold                   : 54,
    setItalic                 : 55,
    setUnderline              : 56,
    setGeneralFormat          : 57,
    redo                      : 58,
    undo                      : 59,
    print                     : 60,
    addSum                    : 61,
    moveToUpperCell           : 62,
    contextMenu               : 63,
    moveToLowerCell           : 64,
    selectToLowerCell         : 65,
    selectToUpperCell         : 66,
    showFilterOptions         : 67,
    showAutoComplete          : 68,
    showDataValidation        : 69,
    increaseFontSize          : 70,
    decreaseFontSize          : 71,
    selectAll                 : 72,
    save                      : 73
  };

  //----------------------------------------------------------export----------------------------------------------------
  window['AscCommonExcel'] = window['AscCommonExcel'] || {};
  window['AscCommonExcel'].c_oAscDrawDepOptions = c_oAscDrawDepOptions;
  window['AscCommonExcel'].c_oAscScrollType = c_oAscScrollType;
  window['AscCommonExcel'].c_oAscLockTypeElem = c_oAscLockTypeElem;
  window['AscCommonExcel'].c_oAscLockTypeElemSubType = c_oAscLockTypeElemSubType;
  window['AscCommonExcel'].c_oAscRecalcIndexTypes = c_oAscRecalcIndexTypes;
  window['AscCommonExcel'].c_oAscCellEditorSelectState = c_oAscCellEditorSelectState;
  window['AscCommonExcel'].c_oAscCanChangeColWidth = c_oAscCanChangeColWidth;
  window['AscCommonExcel'].c_oAscMergeType = c_oAscMergeType;
  window['AscCommonExcel'].c_oAscPaneState = c_oAscPaneState;
  window['AscCommonExcel'].c_oTargetType = c_oTargetType;
  window['AscCommonExcel'].c_oAscCoAuthoringMeBorderColor = c_oAscCoAuthoringMeBorderColor;
  window['AscCommonExcel'].c_oAscCoAuthoringOtherBorderColor = c_oAscCoAuthoringOtherBorderColor;
  window['AscCommonExcel'].c_oAscCoAuthoringLockTablePropertiesBorderColor = c_oAscCoAuthoringLockTablePropertiesBorderColor;
  window['AscCommonExcel'].c_oAscCoAuthoringDottedWidth = c_oAscCoAuthoringDottedWidth;
  window['AscCommonExcel'].c_oAscCoAuthoringDottedDistance = c_oAscCoAuthoringDottedDistance;
  window['AscCommonExcel'].c_oAscFormulaRangeBorderColor = c_oAscFormulaRangeBorderColor;
  window['AscCommonExcel'].c_oAscVisibleAreaOleEditorBorderColor = c_oAscVisibleAreaOleEditorBorderColor;
  window['AscCommonExcel'].selectionLineType = selectionLineType;
  window['AscCommonExcel'].c_oAscLockNameFrozenPane = c_oAscLockNameFrozenPane;
  window['AscCommonExcel'].c_oAscLockNameTabColor = c_oAscLockNameTabColor;
  window['AscCommonExcel'].c_oAscLockAddSheet = c_oAscLockAddSheet;
  window['AscCommonExcel'].c_oAscLockLayoutOptions = c_oAscLockLayoutOptions;
  window['AscCommonExcel'].c_oAscHeaderFooterEdit = c_oAscHeaderFooterEdit;
  window['AscCommonExcel'].c_oAscLockPrintScaleOptions = c_oAscLockPrintScaleOptions;
  window['AscCommonExcel'].docChangedType = docChangedType;


  window['AscCommonExcel'].c_kMaxPrintPages = c_kMaxPrintPages;
  window['AscCommonExcel'].filteringMode = true;

  window['AscCommon'] = window['AscCommon'] || {};
  window['AscCommon'].c_oSerFormat = c_oSerFormat;
  window['AscCommon'].CurFileVersion = c_oSerFormat.Version;

  var prot;
  window['Asc'] = window['Asc'] || {};
  window['Asc']['c_oAscSortOptions'] = window['Asc'].c_oAscSortOptions = c_oAscSortOptions;
  prot = c_oAscSortOptions;
  prot['Ascending'] = prot.Ascending;
  prot['Descending'] = prot.Descending;
  prot['ByColorFill'] = prot.ByColorFill;
  prot['ByColorFont'] = prot.ByColorFont;
  prot['ByIcon'] = prot.ByIcon;
  prot['ByValue'] = prot.ByValue;

  prot['ConfirmReplaceFormulaInTable'] = prot.ConfirmReplaceFormulaInTable;
  window['Asc']['c_oAscMergeOptions'] = window['Asc'].c_oAscMergeOptions = c_oAscMergeOptions;
  prot = c_oAscMergeOptions;
  prot['Disabled'] = prot.Disabled;
  prot['None'] = prot.None;
  prot['Merge'] = prot.Merge;
  prot['MergeCenter'] = prot.MergeCenter;
  prot['MergeAcross'] = prot.MergeAcross;
  window['Asc']['c_oAscBorderOptions'] = window['Asc'].c_oAscBorderOptions = c_oAscBorderOptions;
  prot = c_oAscBorderOptions;
  prot['Top'] = prot.Top;
  prot['Right'] = prot.Right;
  prot['Bottom'] = prot.Bottom;
  prot['Left'] = prot.Left;
  prot['DiagD'] = prot.DiagD;
  prot['DiagU'] = prot.DiagU;
  prot['InnerV'] = prot.InnerV;
  prot['InnerH'] = prot.InnerH;
  window['Asc']['c_oAscCleanOptions'] = window['Asc'].c_oAscCleanOptions = c_oAscCleanOptions;
  prot = c_oAscCleanOptions;
  prot['All'] = prot.All;
  prot['Text'] = prot.Text;
  prot['Format'] = prot.Format;
  prot['Formula'] = prot.Formula;
  prot['Comments'] = prot.Comments;
  prot['Hyperlinks'] = prot.Hyperlinks;
  prot['Sparklines'] = prot.Sparklines;
  prot['SparklineGroups'] = prot.SparklineGroups;
  window['Asc']['c_oAscCalculateType'] = window['Asc'].c_oAscCalculateType = c_oAscCalculateType;
  prot = c_oAscCalculateType;
  prot['WorkbookOnlyChanged'] = prot.WorkbookOnlyChanged;
  prot['ActiveSheet'] = prot.ActiveSheet;
  prot['Workbook'] = prot.Workbook;
  prot['All'] = prot.All;
  window['Asc']['c_oAscSelectionDialogType'] = window['Asc'].c_oAscSelectionDialogType = c_oAscSelectionDialogType;
  prot = c_oAscSelectionDialogType;
  prot['None'] = prot.None;
  prot['FormatTable'] = prot.FormatTable;
  prot['Chart'] = prot.Chart;
  prot['FormatTableChangeRange'] = prot.FormatTableChangeRange;
  prot['CustomSort'] = prot.CustomSort;
  prot['PivotTableData'] = prot.PivotTableData;
  prot['PivotTableReport'] = prot.PivotTableReport;
  prot['PrintTitles'] = prot.PrintTitles;
  prot['Function'] = prot.Function;
  prot['DataValidation'] = prot.DataValidation;
  prot['ImportXml'] = prot.ImportXml;
  prot['ConditionalFormattingRule'] = prot.ConditionalFormattingRule;
  prot['GoalSeek_Cell'] = prot.GoalSeek_Cell;
  prot['GoalSeek_ChangingCell'] = prot.GoalSeek_ChangingCell;


  window['Asc']['c_oAscHyperlinkType'] = window['Asc'].c_oAscHyperlinkType = c_oAscHyperlinkType;
  prot = c_oAscHyperlinkType;
  prot['WebLink'] = prot.WebLink;
  prot['RangeLink'] = prot.RangeLink;
  prot['FileLink'] = prot.FileLink;
  window['Asc']['c_oAscMouseMoveType'] = window['Asc'].c_oAscMouseMoveType = c_oAscMouseMoveType;
  prot = c_oAscMouseMoveType;
  prot['None'] = prot.None;
  prot['Hyperlink'] = prot.Hyperlink;
  prot['Comment'] = prot.Comment;
  prot['LockedObject'] = prot.LockedObject;
  prot['ResizeColumn'] = prot.ResizeColumn;
  prot['ResizeRow'] = prot.ResizeRow;
  prot['Filter'] = prot.Filter;
  prot['Tooltip'] = prot.Tooltip;
  prot['ForeignSelect'] = prot.ForeignSelect;
  prot['Eyedropper'] = prot.Eyedropper;
  window['Asc']['c_oAscMouseMoveLockedObjectType'] = window['Asc'].c_oAscMouseMoveLockedObjectType = c_oAscMouseMoveLockedObjectType;
  prot = c_oAscMouseMoveLockedObjectType;
  prot['None'] = prot.None;
  prot['Range'] = prot.Range;
  prot['TableProperties'] = prot.TableProperties;
  prot['Sheet'] = prot.Sheet;
  window['Asc']['c_oAscCustomAutoFilter'] = window['Asc'].c_oAscCustomAutoFilter = c_oAscCustomAutoFilter;
  prot = c_oAscCustomAutoFilter;
  prot['equals'] = prot.equals;
  prot['isGreaterThan'] = prot.isGreaterThan;
  prot['isGreaterThanOrEqualTo'] = prot.isGreaterThanOrEqualTo;
  prot['isLessThan'] = prot.isLessThan;
  prot['isLessThanOrEqualTo'] = prot.isLessThanOrEqualTo;
  prot['doesNotEqual'] = prot.doesNotEqual;
  prot['beginsWith'] = prot.beginsWith;
  prot['doesNotBeginWith'] = prot.doesNotBeginWith;
  prot['endsWith'] = prot.endsWith;
  prot['doesNotEndWith'] = prot.doesNotEndWith;
  prot['contains'] = prot.contains;
  prot['doesNotContain'] = prot.doesNotContain;
  window['Asc']['c_oAscDynamicAutoFilter'] = window['Asc'].c_oAscDynamicAutoFilter = c_oAscDynamicAutoFilter;
  prot = c_oAscDynamicAutoFilter;
  prot['aboveAverage'] = prot.aboveAverage;
  prot['belowAverage'] = prot.belowAverage;
  prot['lastMonth']    = prot.lastMonth;
  prot['lastQuarter']  = prot.lastQuarter;
  prot['lastWeek']     = prot.lastWeek;
  prot['lastYear']     = prot.lastYear;
  prot['m1']           = prot.m1;
  prot['m11']          = prot.m11;
  prot['m12']          = prot.m12;
  prot['m2']           = prot.m2;
  prot['m3']           = prot.m3;
  prot['m4']           = prot.m4;
  prot['m5']           = prot.m5;
  prot['m6']           = prot.m6;
  prot['m7']           = prot.m7;
  prot['m8']           = prot.m8;
  prot['m9']           = prot.m9;
  prot['nextMonth']    = prot.nextMonth;
  prot['nextQuarter']  = prot.nextQuarter;
  prot['nextWeek']     = prot.nextWeek;
  prot['nextYear']     = prot.nextYear;
  prot['nullType']     = prot.nullType;
  prot['q1']           = prot.q1;
  prot['q2']           = prot.q2;
  prot['q3']           = prot.q3;
  prot['q4']           = prot.q4;
  prot['thisMonth']    = prot.thisMonth;
  prot['thisQuarter']  = prot.thisQuarter;
  prot['thisWeek']     = prot.thisWeek;
  prot['thisYear']     = prot.thisYear;
  prot['today']        = prot.today;
  prot['tomorrow']     = prot.tomorrow;
  prot['yearToDate']   = prot.yearToDate;
  prot['yesterday']    = prot.yesterday;
  window['Asc']['c_oAscTop10AutoFilter'] = window['Asc'].c_oAscTop10AutoFilter = c_oAscTop10AutoFilter;
  prot = c_oAscTop10AutoFilter;
  prot['max'] = prot.max;
  prot['min'] = prot.min;
  window['Asc']['c_oAscChangeFilterOptions'] = window['Asc'].c_oAscChangeFilterOptions = c_oAscChangeFilterOptions;
  prot = c_oAscChangeFilterOptions;
  prot['filter'] = prot.filter;
  prot['style'] = prot.style;
  window['Asc']['c_oAscCellEditorState'] = window['Asc'].c_oAscCellEditorState = c_oAscCellEditorState;
  prot = c_oAscCellEditorState;
  prot['editEnd'] = prot.editEnd;
  prot['editStart'] = prot.editStart;
  prot['editEmptyCell'] = prot.editEmptyCell;
  prot['editText'] = prot.editText;
  prot['editFormula'] = prot.editFormula;
  prot['editInFormulaBar'] = prot.editInFormulaBar;
  prot['editInCell'] = prot.editInCell;
  window['Asc']['c_oAscChangeSelectionFormatTable'] = window['Asc'].c_oAscChangeSelectionFormatTable = c_oAscChangeSelectionFormatTable;
  prot = c_oAscChangeSelectionFormatTable;
  prot['all'] = prot.all;
  prot['data'] = prot.data;
  prot['row'] = prot.row;
  prot['column'] = prot.column;
  prot['dataColumn'] = prot.dataColumn;
  window['Asc']['c_oAscChangeTableStyleInfo'] = window['Asc'].c_oAscChangeTableStyleInfo = c_oAscChangeTableStyleInfo;
  prot = c_oAscChangeTableStyleInfo;
  prot['columnFirst'] = prot.columnFirst;
  prot['columnLast'] = prot.columnLast;
  prot['columnBanded'] = prot.columnBanded;
  prot['rowHeader'] = prot.rowHeader;
  prot['rowTotal'] = prot.rowTotal;
  prot['rowBanded'] = prot.rowBanded;
  prot['filterButton'] = prot.filterButton;
  prot['advancedSettings'] = prot.advancedSettings;
  window['Asc']['c_oAscAutoFilterTypes'] = window['Asc'].c_oAscAutoFilterTypes = c_oAscAutoFilterTypes;
  prot = c_oAscAutoFilterTypes;
  prot['ColorFilter'] = prot.ColorFilter;
  prot['CustomFilters'] = prot.CustomFilters;
  prot['DynamicFilter'] = prot.DynamicFilter;
  prot['Top10'] = prot.Top10;
  prot['Filters'] = prot.Filters;
  prot['None'] = prot.None;
  window['Asc']['c_oAscFindLookIn'] = window['Asc'].c_oAscFindLookIn = c_oAscFindLookIn;
  prot = c_oAscFindLookIn;
  prot['Formulas'] = prot.Formulas;
  prot['Value'] = prot.Value;
  prot['Annotations'] = prot.Annotations;
  window['Asc']['c_oAscGetDefinedNamesList'] = window['Asc'].c_oAscGetDefinedNamesList = c_oAscGetDefinedNamesList;
  prot = c_oAscGetDefinedNamesList;
  prot['Worksheet'] = prot.Worksheet;
  prot['WorksheetWorkbook'] = prot.WorksheetWorkbook;
  prot['All'] = prot.All;
  window['Asc']['c_oAscDefinedNameReason'] = window['Asc'].c_oAscDefinedNameReason = c_oAscDefinedNameReason;
  prot = c_oAscDefinedNameReason;
  prot['WrongName'] = prot.WrongName;
  prot['IsLocked'] = prot.IsLocked;
  prot['Existed'] = prot.Existed;
  prot['LockDefNameManager'] = prot.LockDefNameManager;
  prot['NameReserved'] = prot.NameReserved;
  prot['OK'] = prot.OK;
  window['Asc']['c_oAscPopUpSelectorType'] = window['Asc'].c_oAscPopUpSelectorType = c_oAscPopUpSelectorType;
  prot = c_oAscPopUpSelectorType;
  prot['None'] = prot.None;
  prot['Func'] = prot.Func;
  prot['Range'] = prot.Range;
  prot['Table'] = prot.Table;
  prot['Slicer'] = prot.Slicer;
  prot['TotalRowFunc'] = prot.TotalRowFunc;
  prot['TableColumnName'] = prot.TableColumnName;
  prot['TableThisRow'] = prot.TableThisRow;
  prot['TableAll'] = prot.TableAll;
  prot['TableData'] = prot.TableData;
  prot['TableHeaders'] = prot.TableHeaders;
  prot['TableTotals'] = prot.TableTotals;

  window['Asc']['c_oAscSparklineType'] = window['Asc'].c_oAscSparklineType = c_oAscSparklineType;
  prot = c_oAscSparklineType;
  prot['Line'] = prot.Line;
  prot['Column'] = prot.Column;
  prot['Stacked'] = prot.Stacked;
  window['Asc']['c_oAscEDispBlanksAs'] = window['Asc'].c_oAscEDispBlanksAs = c_oAscEDispBlanksAs;
  prot = c_oAscEDispBlanksAs;
  prot['Span'] = prot.Span;
  prot['Gap'] = prot.Gap;
  prot['Zero'] = prot.Zero;
  window['Asc']['c_oAscSparklineAxisMinMax'] = window['Asc'].c_oAscSparklineAxisMinMax = c_oAscSparklineAxisMinMax;
  prot = c_oAscSparklineAxisMinMax;
  prot['Individual'] = prot.Individual;
  prot['Group'] = prot.Group;
  prot['Custom'] = prot.Custom;
  window['Asc']['c_oAscAutoCorrectOptions'] = window['Asc'].c_oAscAutoCorrectOptions = c_oAscAutoCorrectOptions;
  prot = c_oAscAutoCorrectOptions;
  prot['UndoTableAutoExpansion'] = prot.UndoTableAutoExpansion;
  prot['RedoTableAutoExpansion'] = prot.RedoTableAutoExpansion;
  window['Asc']['c_oAscChangePrintAreaType'] = window['Asc'].c_oAscChangePrintAreaType = c_oAscChangePrintAreaType;
  prot = c_oAscChangePrintAreaType;
  prot['set'] = prot.set;
  prot['clear'] = prot.clear;
  prot['add'] = prot.add;
  window['Asc']['c_oAscHeaderFooterField'] = window['Asc'].c_oAscHeaderFooterField = c_oAscHeaderFooterField;
  prot = c_oAscHeaderFooterField;
  prot['pageNumber'] = prot.pageNumber;
  prot['pageCount'] = prot.pageCount;
  prot['sheetName'] = prot.sheetName;
  prot['fileName'] = prot.fileName;
  prot['filePath'] = prot.filePath;
  prot['date'] = prot.date;
  prot['time'] = prot.time;
  prot['lineBreak'] = prot.lineBreak;
  prot['picture'] = prot.picture;
  prot['text'] = prot.text;
  window['Asc']['c_oAscPageHFType'] = window['Asc'].c_oAscPageHFType = c_oAscPageHFType;
  prot = c_oAscPageHFType;
  prot['firstHeader'] = prot.firstHeader;
  prot['oddHeader'] = prot.oddHeader;
  prot['evenHeader'] = prot.evenHeader;
  prot['firstFooter'] = prot.firstFooter;
  prot['oddFooter'] = prot.oddFooter;
  prot['evenFooter'] = prot.evenFooter;

  window['Asc']['c_oAscHeaderFooterType'] = window['Asc'].c_oAscHeaderFooterType = c_oAscHeaderFooterType;
  prot = c_oAscHeaderFooterType;
  prot['first'] = prot.first;
  prot['odd'] = prot.odd;
  prot['even'] = prot.even;

  window['Asc']['c_oAscHeaderFooterPresets'] = window['Asc'].c_oAscHeaderFooterPresets = c_oAscHeaderFooterPresets;
  prot = c_oAscHeaderFooterPresets;
  prot['none'] = prot.none;
  prot['page'] = prot.page;
  prot['pageOfQuestion'] = prot.pageOfQuestion;
  prot['sheet'] = prot.sheet;
  prot['confidential'] = prot.confidential;
  prot['bookName'] = prot.bookName;
  prot['sheetPage'] = prot.sheetPage;
  prot['sheetConfidentialPage'] = prot.sheetConfidentialPage;
  prot['bookNamePage'] = prot.bookNamePage;
  prot['pageSheet'] = prot.pageSheet;
  prot['pageBook'] = prot.pageBook;
  prot['pageBookName'] = prot.pageBookName;
  prot['userPageDate'] = prot.userPageDate;
  prot['preparedUserDatePage'] = prot.preparedUserDatePage;
  prot['custom'] = prot.custom;

  window['Asc']['c_oAscPrintTitlesRangeType'] = window['Asc'].c_oAscPrintTitlesRangeType = c_oAscPrintTitlesRangeType;
  prot = c_oAscPrintTitlesRangeType;
  prot['first'] = prot.first;
  prot['frozen'] = prot.frozen;
  prot['current'] = prot.current;

  window['Asc']['c_oAscFormulaArgumentType'] = window['Asc'].c_oAscFormulaArgumentType = c_oAscFormulaArgumentType;
  prot = c_oAscFormulaArgumentType;
  prot['number'] = prot.number;
  prot['text'] = prot.text;
  prot['reference'] = prot.reference;
  prot['any'] = prot.any;
  prot['logical'] = prot.logical;
  prot['array'] = prot.array;

  window['Asc']['c_oAscSelectionForCFType'] = window['Asc'].c_oAscSelectionForCFType = c_oAscSelectionForCFType;
  prot = c_oAscSelectionForCFType;
  prot['selection'] = prot.selection;
  prot['worksheet'] = prot.worksheet;
  prot['table'] = prot.table;
  prot['pivot'] = prot.pivot;


  window['Asc']['c_oAscFrozenPaneBorderType'] = window['Asc'].c_oAscFrozenPaneBorderType = c_oAscFrozenPaneBorderType;
  prot = c_oAscFrozenPaneBorderType;
  prot['shadow'] = prot.shadow;
  prot['line'] = prot.line;

  window['Asc']['c_oAscCFRuleTypeSettings'] = window['Asc'].c_oAscCFRuleTypeSettings = c_oAscCFRuleTypeSettings;
  prot = c_oAscCFRuleTypeSettings;
  prot['dataBar'] = prot.dataBar;
  prot['colorScale'] = prot.colorScale;
  prot['icons'] = prot.icons;
  prot['format'] = prot.format;
  
  window['Asc']['c_oAscSheetProtectType'] = window['Asc'].c_oAscSheetProtectType = c_oAscSheetProtectType;
  prot = c_oAscSheetProtectType;
  prot['objects'] = prot.objects;
  prot['scenarios'] = prot.scenarios;
  prot['formatCells'] = prot.formatCells;
  prot['formatColumns'] = prot.formatColumns;
  prot['formatRows'] = prot.formatRows;
  prot['insertColumns'] = prot.insertColumns;
  prot['insertRows'] = prot.insertRows;
  prot['insertHyperlinks'] = prot.insertHyperlinks;
  prot['deleteColumns'] = prot.deleteColumns;
  prot['deleteRows'] = prot.deleteRows;
  prot['deleteRows'] = prot.deleteRows;
  prot['selectLockedCells'] = prot.selectLockedCells;
  prot['sort'] = prot.sort;
  prot['autoFilter'] = prot.autoFilter;
  prot['pivotTables'] = prot.pivotTables;
  prot['selectUnlockedCells'] = prot.selectUnlockedCells;

  window['Asc']['c_oAscFrozenPaneAddType'] = window['Asc'].c_oAscFrozenPaneAddType = c_oAscFrozenPaneAddType;
  prot = c_oAscFrozenPaneAddType;
  prot['firstRow'] = prot.firstRow;
  prot['firstCol'] = prot.firstCol;
  window['Asc']['c_oAscWorkbookProtectType'] = window['Asc'].c_oAscWorkbookProtectType = c_oAscWorkbookProtectType;
  prot = c_oAscWorkbookProtectType;
  prot['lockStructure'] = prot.lockStructure;
  prot['lockWindows'] = prot.lockWindows;
  prot['lockRevisions'] = prot.lockRevisions;

  window['Asc']['c_oAscSelectionSortExpand'] = window['Asc'].c_oAscSelectionSortExpand = c_oAscSelectionSortExpand;
  prot = c_oAscSelectionSortExpand;
  prot['expandAndNotShowMessage'] = prot.expandAndNotShowMessage;
  prot['notExpandAndNotShowMessage'] = prot.notExpandAndNotShowMessage;
  prot['showExpandMessage'] = prot.showExpandMessage;
  prot['showLockMessage'] = prot.showLockMessage;

  window['Asc']['c_oAscSearchBy'] = window['Asc'].c_oAscSearchBy = c_oAscSearchBy;
  prot = c_oAscSearchBy;
  prot['Workbook'] = prot.Workbook;
  prot['Sheet'] = prot.Sheet;
  prot['Range'] = prot.Range;

  window['Asc']['c_nAscMaxAddCellWatchesCount'] = window['Asc'].c_nAscMaxAddCellWatchesCount = c_nAscMaxAddCellWatchesCount;
  window['Asc']['c_oAscExternalReferenceType'] = window['Asc'].c_oAscExternalReferenceType = c_oAscExternalReferenceType;
  prot = c_oAscExternalReferenceType;
  prot['referenceData'] = prot.referenceData;
  prot['link'] = prot.link;
  prot['path'] = prot.path;

  window['Asc']['c_oAscPageBreaksDisableType'] = window['Asc'].c_oAscPageBreaksDisableType = c_oAscPageBreaksDisableType;
  prot = c_oAscPageBreaksDisableType;
  prot['none'] = prot.none;
  prot['all'] = prot.all;
  prot['insertRemove'] = prot.insertRemove;
  prot['reset'] = prot.reset;
  window['Asc']['c_oAscRemoveArrowsType'] = window['Asc'].c_oAscRemoveArrowsType = c_oAscRemoveArrowsType;
  prot = c_oAscRemoveArrowsType;
  prot['all'] = prot.all;
  prot['precedent'] = prot.precedent;
  prot['dependent'] = prot.dependent;

  window['Asc']['c_oAscSeriesInType'] = window['Asc'].c_oAscSeriesInType = c_oAscSeriesInType;
  prot = c_oAscSeriesInType;
  prot['rows'] = prot.rows;
  prot['columns'] = prot.columns;

  window['Asc']['c_oAscSeriesType'] = window['Asc'].c_oAscSeriesType = c_oAscSeriesType;
  prot = c_oAscSeriesType;
  prot['linear'] = prot.linear;
  prot['growth'] = prot.growth;
  prot['date'] = prot.date;
  prot['autoFill'] = prot.autoFill;

  window['Asc']['c_oAscDateUnitType'] = window['Asc'].c_oAscDateUnitType = c_oAscDateUnitType;
  prot = c_oAscDateUnitType;
  prot['day'] = prot.day;
  prot['weekday'] = prot.weekday;
  prot['month'] = prot.month;
  prot['year'] = prot.year;

  window['Asc']['c_oAscFillType'] = window['Asc'].c_oAscFillType = c_oAscFillType;
  prot = c_oAscFillType;
  prot['copyCells'] = prot.copyCells;
  prot['fillSeries'] = prot.fillSeries;
  prot['fillFormattingOnly'] = prot.fillFormattingOnly;
  prot['fillWithoutFormatting'] = prot.fillWithoutFormatting;
  prot['fillDays'] = prot.fillDays;
  prot['fillWeekdays'] = prot.fillWeekdays;
  prot['fillMonths'] = prot.fillMonths;
  prot['fillYears'] = prot.fillYears;
  prot['linearTrend'] = prot.linearTrend;
  prot['growthTrend'] = prot.growthTrend;
  prot['flashFill'] = prot.flashFill;
  prot['series'] = prot.series;
  prot['fillDown'] = prot.fillDown;
  prot['fillRight'] = prot.fillRight;
  prot['fillUp'] = prot.fillUp;
  prot['fillLeft'] = prot.fillLeft;
  prot['justify'] = prot.justify;

  window['Asc']['c_oAscContextMenuTypes'] = window['Asc'].c_oAscContextMenuTypes = c_oAscContextMenuTypes;
  prot = c_oAscContextMenuTypes;
  prot['common'] = prot.common;
  prot['changeSeries'] = prot.changeSeries;

  window['Asc']['c_oAscCalcMode'] = window['Asc'].c_oAscCalcMode = c_oAscCalcMode;
  prot = c_oAscCalcMode;
  prot['auto'] = prot.auto;
  prot['autoNoTable'] = prot.autoNoTable;
  prot['manual'] = prot.manual;


  window['Asc']['c_oAscCellShortcutType'] = window['Asc'].c_oAscCellShortcutType = c_oAscCellShortcutType;
  prot = c_oAscCellShortcutType;
  prot['refreshAllConnections']       = c_oAscCellShortcutType.refreshAllConnections;
  prot['refreshSelectedConnections']  = c_oAscCellShortcutType.refreshSelectedConnections;
  prot['changeFormatTableInfo']       = c_oAscCellShortcutType.changeFormatTableInfo;
  prot['calculateAll']                = c_oAscCellShortcutType.calculateAll;
  prot['calculateActiveSheet']        = c_oAscCellShortcutType.calculateActiveSheet;
  prot['focusOnCellEditor']           = c_oAscCellShortcutType.focusOnCellEditor;
  prot['addDate']                     = c_oAscCellShortcutType.addDate;
  prot['addTime']                     = c_oAscCellShortcutType.addTime;
  prot['removeActiveCell']            = c_oAscCellShortcutType.removeActiveCell;
  prot['emptyRange']                  = c_oAscCellShortcutType.emptyRange;
  prot['moveActiveCellToLeft']        = c_oAscCellShortcutType.moveActiveCellToLeft;
  prot['moveActiveCellToRight']       = c_oAscCellShortcutType.moveActiveCellToRight;
  prot['moveActiveCellToDown']        = c_oAscCellShortcutType.moveActiveCellToDown;
  prot['moveActiveCellToUp']          = c_oAscCellShortcutType.moveActiveCellToUp;
  prot['reset']                       = c_oAscCellShortcutType.reset;
  prot['disableNumLock']              = c_oAscCellShortcutType.disableNumLock;
  prot['disableScrollLock']           = c_oAscCellShortcutType.disableScrollLock;
  prot['selectColumn']                = c_oAscCellShortcutType.selectColumn;
  prot['selectRow']                   = c_oAscCellShortcutType.selectRow;
  prot['selectSheet']                 = c_oAscCellShortcutType.selectSheet;
  prot['addSeparator']                = c_oAscCellShortcutType.addSeparator;
  prot['goToPreviousSheet']           = c_oAscCellShortcutType.goToPreviousSheet;
  prot['moveToTopCell']               = c_oAscCellShortcutType.moveToTopCell;
  prot['moveToNextSheet']             = c_oAscCellShortcutType.moveToNextSheet;
  prot['moveToLeftEdgeCell']          = c_oAscCellShortcutType.moveToLeftEdgeCell;
  prot['selectToLeftEdgeCell']        = c_oAscCellShortcutType.selectToLeftEdgeCell;
  prot['moveToLeftCell']              = c_oAscCellShortcutType.moveToLeftCell;
  prot['selectToLeftCell']            = c_oAscCellShortcutType.selectToLeftCell;
  prot['moveToRightEdgeCell']         = c_oAscCellShortcutType.moveToRightEdgeCell;
  prot['selectToRightEdgeCell']       = c_oAscCellShortcutType.selectToRightEdgeCell;
  prot['moveToRightCell']             = c_oAscCellShortcutType.moveToRightCell;
  prot['selectToRightCell']           = c_oAscCellShortcutType.selectToRightCell;
  prot['selectToTopCell']             = c_oAscCellShortcutType.selectToTopCell;
  prot['moveToUpCell']                = c_oAscCellShortcutType.moveToUpCell;
  prot['selectToUpCell']              = c_oAscCellShortcutType.selectToUpCell;
  prot['moveToBottomCell']            = c_oAscCellShortcutType.moveToBottomCell;
  prot['selectToBottomCell']          = c_oAscCellShortcutType.selectToBottomCell;
  prot['moveToDownCell']              = c_oAscCellShortcutType.moveToDownCell;
  prot['selectToDownCell']            = c_oAscCellShortcutType.selectToDownCell;
  prot['moveToFirstColumn']           = c_oAscCellShortcutType.moveToFirstColumn;
  prot['selectToFirstColumn']         = c_oAscCellShortcutType.selectToFirstColumn;
  prot['moveToLeftEdgeTop']           = c_oAscCellShortcutType.moveToLeftEdgeTop;
  prot['selectToLeftEdgeTop']         = c_oAscCellShortcutType.selectToLeftEdgeTop;
  prot['moveToRightBottomEdge']       = c_oAscCellShortcutType.moveToRightBottomEdge;
  prot['selectToRightBottomEdge']     = c_oAscCellShortcutType.selectToRightBottomEdge;
  prot['setNumberFormat']             = c_oAscCellShortcutType.setNumberFormat;
  prot['setTimeFormat']               = c_oAscCellShortcutType.setTimeFormat;
  prot['setDateFormat']               = c_oAscCellShortcutType.setDateFormat;
  prot['setCurrencyFormat']           = c_oAscCellShortcutType.setCurrencyFormat;
  prot['setPercentFormat']            = c_oAscCellShortcutType.setPercentFormat;
  prot['setStrikethrough']            = c_oAscCellShortcutType.setStrikethrough;
  prot['setExponentialFormat']        = c_oAscCellShortcutType.setExponentialFormat;
  prot['setBold']                     = c_oAscCellShortcutType.setBold;
  prot['setItalic']                   = c_oAscCellShortcutType.setItalic;
  prot['setUnderline']                = c_oAscCellShortcutType.setUnderline;
  prot['setGeneralFormat']            = c_oAscCellShortcutType.setGeneralFormat;
  prot['redo']                        = c_oAscCellShortcutType.redo;
  prot['undo']                        = c_oAscCellShortcutType.undo;
  prot['print']                       = c_oAscCellShortcutType.print;
  prot['addSum']                      = c_oAscCellShortcutType.addSum;
  prot['moveToUpperCell']             = c_oAscCellShortcutType.moveToUpperCell;
  prot['contextMenu']                 = c_oAscCellShortcutType.contextMenu;
  prot['moveToLowerCell']             = c_oAscCellShortcutType.moveToLowerCell;
  prot['selectToLowerCell']           = c_oAscCellShortcutType.selectToLowerCell;
  prot['selectToUpper Cell']          = c_oAscCellShortcutType.selectToUpperCell;
  prot['showFilterOptions']           = c_oAscCellShortcutType.showFilterOptions;
  prot['showAutoComplete']            = c_oAscCellShortcutType.showAutoComplete;
  prot['showDataValidation']          = c_oAscCellShortcutType.showDataValidation;
  prot['increaseFontSize']            = c_oAscCellShortcutType.increaseFontSize;
  prot['decreaseFontSize']            = c_oAscCellShortcutType.decreaseFontSize;
  prot['selectAll']                   = c_oAscCellShortcutType.selectAll;
  prot['save']                        = c_oAscCellShortcutType.save;

})(window);
