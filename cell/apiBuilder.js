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

(function (window, builder) {
	function checkFormat(value) {
		if (value.getTime) {
			return new AscCommonExcel.cNumber(new Asc.cDate(value.getTime()).getExcelDateWithTime(true));
		} else {
			return new AscCommonExcel.cString(value + '');
		}
	}

	/**
	 * Base class.
	 * @global
	 * @class
	 * @name Api
	 * @property {Array} Sheets - Returns the Sheets collection that represents all the sheets in the active workbook.
	 * @property {ApiWorksheet} ActiveSheet - Returns an object that represents the active sheet.
	 * @property {ApiRange} Selection - Returns an object that represents the selected range.
	 * @property {ApiComment[]} Comments - Returns all comments related to the whole workbook.
	 * @property {FreezePaneType} FreezePanes - Returns or sets the type of freeze panes.
	 * @property {ApiComment[]} AllComments - Returns all comments from the current workbook including comments from all worksheets.
	 * @property {ReferenceStyle} ReferenceStyle - Returns or sets the reference style.
	 * @property {ApiWorksheetFunction} WorksheetFunction - Returns an object that represents the function list.
	 * @property {ApiPivotTable[]} PivotTables - Returns all PivotTables.
	 */
	var Api = window["Asc"]["spreadsheet_api"];

	/**
	 * The callback function which is called when the specified range of the current sheet changes.
	 * <note>Please note that the event is not called for the undo/redo operations.</note>
	 * @event Api#onWorksheetChange
	 * @param {ApiRange} range - The modified range represented as the ApiRange object.
	 */

	/**
	 * Class representing a sheet.
	 * @constructor
	 * @property {boolean} Visible - Returns or sets the state of sheet visibility.
	 * @property {number} Active - Makes the current sheet active.
	 * @property {ApiRange} ActiveCell - Returns an object that represents an active cell.
	 * @property {ApiRange} Selection - Returns an object that represents the selected range.
	 * @property {ApiRange} Cells - Returns ApiRange that represents all the cells on the worksheet (not just the cells that are currently in use).
	 * @property {ApiRange} Rows - Returns ApiRange that represents all the cells of the rows range.
	 * @property {ApiRange} Cols - Returns ApiRange that represents all the cells of the columns range.
	 * @property {ApiRange} UsedRange - Returns ApiRange that represents the used range on the specified worksheet.
	 * @property {string} Name - Returns or sets a name of the active sheet.
	 * @property {number} Index - Returns a sheet index.
	 * @property {number} LeftMargin - Returns or sets the size of the sheet left margin measured in points.
	 * @property {number} RightMargin - Returns or sets the size of the sheet right margin measured in points.
	 * @property {number} TopMargin - Returns or sets the size of the sheet top margin measured in points.
	 * @property {number} BottomMargin - Returns or sets the size of the sheet bottom margin measured in points.
	 * @property {PageOrientation} PageOrientation - Returns or sets the page orientation.
	 * @property {boolean} PrintHeadings - Returns or sets the page PrintHeadings property.
	 * @property {boolean} PrintGridlines - Returns or sets the page PrintGridlines property.
	 * @property {Array} Defnames - Returns an array of the ApiName objects.
	 * @property {Array} Comments - Returns all comments from the current worksheet.
	 * @property {ApiFreezePanes} FreezePanes - Returns the freeze panes for the current worksheet.
	 * @property {ApiProtectedRange[]} AllProtectedRanges - Returns all protected ranges from the current worksheet.
	 * @property {ApiPivotTable[]} PivotTables - Returns all PivotTables on worksheet.
	 */
	function ApiWorksheet(worksheet) {
		this.worksheet = worksheet;
	}

	/**
	 * Class representing a range.
	 * @constructor
	 * @property {number} Row - Returns the row number for the selected cell.
	 * @property {number} Col - Returns the column number for the selected cell.
	 * @property {ApiRange} Rows - Returns the ApiRange object that represents the rows of the specified range.
	 * @property {ApiRange} Cols - Returns the ApiRange object that represents the columns of the specified range.
	 * @property {ApiRange} Cells - Returns a Range object that represents all the cells in the specified range or a specified cell.
	 * @property {number} Count - Returns the rows or columns count.
	 * @property {string} Address - Returns the range address.
	 * @property {string} Value - Returns a value from the first cell of the specified range or sets it to this cell.
	 * @property {string} Formula - Returns a formula from the first cell of the specified range or sets it to this cell.
	 * @property {string} Value2 - Returns the value2 (value without format) from the first cell of the specified range or sets it to this cell.
	 * @property {string} Text - Returns the text from the first cell of the specified range or sets it to this cell.
	 * @property {ApiColor} FontColor - Sets the text color to the current cell range with the previously created color object.
	 * @property {boolean} Hidden - Returns or sets the value hiding property.
	 * @property {number} ColumnWidth - Returns or sets the width of all the columns in the specified range measured in points.
	 * @property {number} Width - Returns a value that represents the range width measured in points.
	 * @property {number} RowHeight - Returns or sets the height of the first row in the specified range measured in points.
	 * @property {number} Height - Returns a value that represents the range height measured in points.
	 * @property {number} FontSize - Sets the font size to the characters of the current cell range.
	 * @property {string} FontName - Sets the specified font family as the font name for the current cell range.
	 * @property {'center' | 'bottom' | 'top' | 'distributed' | 'justify'} AlignVertical - Sets the text vertical alignment to the current cell range.
	 * @property {'left' | 'right' | 'center' | 'justify'} AlignHorizontal - Sets the text horizontal alignment to the current cell range.
	 * @property {boolean} Bold - Sets the bold property to the text characters from the current cell or cell range.
	 * @property {boolean} Italic - Sets the italic property to the text characters in the current cell or cell range.
	 * @property {'none' | 'single' | 'singleAccounting' | 'double' | 'doubleAccounting'} Underline - Sets the type of underline applied to the font.
	 * @property {boolean} Strikeout - Sets a value that indicates whether the contents of the current cell or cell range are displayed struck through.
	 * @property {boolean} WrapText - Returns the information about the wrapping cell style or specifies whether the words in the cell must be wrapped to fit the cell size or not.
	 * @property {ApiColor|'No Fill'} FillColor - Returns or sets the background color of the current cell range.
	 * @property {string} NumberFormat - Sets a value that represents the format code for the object.
	 * @property {ApiRange} MergeArea - Returns the cell or cell range from the merge area.
	 * @property {ApiWorksheet} Worksheet - Returns the ApiWorksheet object that represents the worksheet containing the specified range.
	 * @property {ApiName} DefName - Returns the ApiName object.
	 * @property {ApiComment | null} Comments - Returns the ApiComment collection that represents all the comments from the specified worksheet.
	 * @property {'xlDownward' | 'xlHorizontal' | 'xlUpward' | 'xlVertical'} Orientation - Sets an angle to the current cell range.
	 * @property {ApiAreas} Areas - Returns a collection of the areas.
	 * @property {ApiCharacters} Characters - Returns the ApiCharacters object that represents a range of characters within the object text. Use the ApiCharacters object to format characters within a text string.
	 * @property {ApiPivotTable | null} PivotTable - Returns a PivotTable object that represents the PivotTable report containing the upper-left corner of the specified range.
	 */
	function ApiRange(range, areas) {
		this.range = range;
		this.areas = areas || null;
	}


	/**
	 * Class representing a graphical object.
	 * @constructor
	 */
	function ApiDrawing(Drawing) {
		this.Drawing = Drawing;
	}

	/**
	 * Class representing a shape.
	 * @constructor
	 */
	function ApiShape(oShape) {
		ApiDrawing.call(this, oShape);
		this.Shape = oShape;
	}

	ApiShape.prototype = Object.create(ApiDrawing.prototype);
	ApiShape.prototype.constructor = ApiShape;

	/**
	 * Class representing an image.
	 * @constructor
	 */
	function ApiImage(oImage) {
		ApiDrawing.call(this, oImage);
	}

	ApiImage.prototype = Object.create(ApiDrawing.prototype);
	ApiImage.prototype.constructor = ApiImage;



	/**
	 * Class representing an OLE object.
	 * @constructor
	 */
	function ApiOleObject(OleObject) {
		ApiDrawing.call(this, OleObject);
	}

	ApiOleObject.prototype = Object.create(ApiDrawing.prototype);
	ApiOleObject.prototype.constructor = ApiOleObject;

	/**
	 * The available preset color names.
	 * @typedef {("aliceBlue" | "antiqueWhite" | "aqua" | "aquamarine" | "azure" | "beige" | "bisque" | "black" |
	 *     "blanchedAlmond" | "blue" | "blueViolet" | "brown" | "burlyWood" | "cadetBlue" | "chartreuse" | "chocolate"
	 *     | "coral" | "cornflowerBlue" | "cornsilk" | "crimson" | "cyan" | "darkBlue" | "darkCyan" | "darkGoldenrod" |
	 *     "darkGray" | "darkGreen" | "darkGrey" | "darkKhaki" | "darkMagenta" | "darkOliveGreen" | "darkOrange" |
	 *     "darkOrchid" | "darkRed" | "darkSalmon" | "darkSeaGreen" | "darkSlateBlue" | "darkSlateGray" |
	 *     "darkSlateGrey" | "darkTurquoise" | "darkViolet" | "deepPink" | "deepSkyBlue" | "dimGray" | "dimGrey" |
	 *     "dkBlue" | "dkCyan" | "dkGoldenrod" | "dkGray" | "dkGreen" | "dkGrey" | "dkKhaki" | "dkMagenta" |
	 *     "dkOliveGreen" | "dkOrange" | "dkOrchid" | "dkRed" | "dkSalmon" | "dkSeaGreen" | "dkSlateBlue" |
	 *     "dkSlateGray" | "dkSlateGrey" | "dkTurquoise" | "dkViolet" | "dodgerBlue" | "firebrick" | "floralWhite" |
	 *     "forestGreen" | "fuchsia" | "gainsboro" | "ghostWhite" | "gold" | "goldenrod" | "gray" | "green" |
	 *     "greenYellow" | "grey" | "honeydew" | "hotPink" | "indianRed" | "indigo" | "ivory" | "khaki" | "lavender" |
	 *     "lavenderBlush" | "lawnGreen" | "lemonChiffon" | "lightBlue" | "lightCoral" | "lightCyan" |
	 *     "lightGoldenrodYellow" | "lightGray" | "lightGreen" | "lightGrey" | "lightPink" | "lightSalmon" |
	 *     "lightSeaGreen" | "lightSkyBlue" | "lightSlateGray" | "lightSlateGrey" | "lightSteelBlue" | "lightYellow" |
	 *     "lime" | "limeGreen" | "linen" | "ltBlue" | "ltCoral" | "ltCyan" | "ltGoldenrodYellow" | "ltGray" |
	 *     "ltGreen" | "ltGrey" | "ltPink" | "ltSalmon" | "ltSeaGreen" | "ltSkyBlue" | "ltSlateGray" | "ltSlateGrey"|
	 *     "ltSteelBlue" | "ltYellow" | "magenta" | "maroon" | "medAquamarine" | "medBlue" | "mediumAquamarine" |
	 *     "mediumBlue" | "mediumOrchid" | "mediumPurple" | "mediumSeaGreen" | "mediumSlateBlue" |
	 *     "mediumSpringGreen" | "mediumTurquoise" | "mediumVioletRed" | "medOrchid" | "medPurple" | "medSeaGreen" |
	 *     "medSlateBlue" | "medSpringGreen" | "medTurquoise" | "medVioletRed" | "midnightBlue" | "mintCream" |
	 *     "mistyRose" | "moccasin" | "navajoWhite" | "navy" | "oldLace" | "olive" | "oliveDrab" | "orange" |
	 *     "orangeRed" | "orchid" | "paleGoldenrod" | "paleGreen" | "paleTurquoise" | "paleVioletRed" | "papayaWhip"|
	 *     "peachPuff" | "peru" | "pink" | "plum" | "powderBlue" | "purple" | "red" | "rosyBrown" | "royalBlue" |
	 *     "saddleBrown" | "salmon" | "sandyBrown" | "seaGreen" | "seaShell" | "sienna" | "silver" | "skyBlue" |
	 *     "slateBlue" | "slateGray" | "slateGrey" | "snow" | "springGreen" | "steelBlue" | "tan" | "teal" |
	 *     "thistle" | "tomato" | "turquoise" | "violet" | "wheat" | "white" | "whiteSmoke" | "yellow" |
	 *     "yellowGreen")} PresetColor
	 * @see office-js-api/Examples/Enumerations/PresetColor.js
	 */

	/**
	 * Possible values for the position of chart tick labels (either horizontal or vertical).
	 * <b>"none"</b> - does not display the selected tick labels.
	 * <b>"nextTo"</b> - sets the position of the selected tick labels next to the main label.
	 * <b>"low"</b> - sets the position of the selected tick labels in the part of the chart with lower values.
	 * <b>"high"</b> - sets the position of the selected tick labels in the part of the chart with higher values.
	 * @typedef {("none" | "nextTo" | "low" | "high")} TickLabelPosition
	 * @see office-js-api/Examples/Enumerations/TickLabelPosition.js
	 */

	/**
	 * The page orientation type.
	 * @typedef {("xlLandscape" | "xlPortrait")} PageOrientation
	 * @see office-js-api/Examples/Enumerations/PageOrientation.js
	 */

	/**
	 * The type of tick mark appearance.
	 * @typedef {("cross" | "in" | "none" | "out")} TickMark
	 * @see office-js-api/Examples/Enumerations/TickMark.js
	 */

	/**
	 * Text transform type.
	 * @typedef {("textArchDown" | "textArchDownPour" | "textArchUp" | "textArchUpPour" | "textButton" | "textButtonPour" | "textCanDown"
	 * | "textCanUp" | "textCascadeDown" | "textCascadeUp" | "textChevron" | "textChevronInverted" | "textCircle" | "textCirclePour"
	 * | "textCurveDown" | "textCurveUp" | "textDeflate" | "textDeflateBottom" | "textDeflateInflate" | "textDeflateInflateDeflate" | "textDeflateTop"
	 * | "textDoubleWave1" | "textFadeDown" | "textFadeLeft" | "textFadeRight" | "textFadeUp" | "textInflate" | "textInflateBottom" | "textInflateTop"
	 * | "textPlain" | "textRingInside" | "textRingOutside" | "textSlantDown" | "textSlantUp" | "textStop" | "textTriangle" | "textTriangleInverted"
	 * | "textWave1" | "textWave2" | "textWave4" | "textNoShape")} TextTransform
	 * @see office-js-api/Examples/Enumerations/TextTransform.js
	 */

	/**
	 * Axis position in the chart.
	 * @typedef {("top" | "bottom" | "right" | "left")} AxisPos
	 * @see office-js-api/Examples/Enumerations/AxisPos.js
	 */

	/**
	 * Standard numeric format.
	 * @typedef {("General" | "0" | "0.00" | "#,##0" | "#,##0.00" | "0%" | "0.00%" |
	 * "0.00E+00" | "# ?/?" | "# ??/??" | "m/d/yyyy" | "d-mmm-yy" | "d-mmm" | "mmm-yy" | "h:mm AM/PM" |
	 * "h:mm:ss AM/PM" | "h:mm" | "h:mm:ss" | "m/d/yyyy h:mm" | "#,##0_);(#,##0)" | "#,##0_);[Red](#,##0)" | 
	 * "#,##0.00_);(#,##0.00)" | "#,##0.00_);[Red](#,##0.00)" | "mm:ss" | "[h]:mm:ss" | "mm:ss.0" | "##0.0E+0" | "@")} NumFormat
	 * @see office-js-api/Examples/Enumerations/NumFormat.js
	 */

	/**
	 * The cell reference type.
	 * @typedef {('xlA1' | 'xlR1C1')} ReferenceStyle
	 * @see office-js-api/Examples/Enumerations/ReferenceStyle.js
	 */

	//TODO not support "xlPasteAllMergingConditionalFormats" / "xlPasteAllUsingSourceTheme" / "xlPasteValidation"
	/**
	 * Specifies the part of the range to be pasted.
	 * @typedef {("xlPasteAll" | "xlPasteAllExceptBorders"
	 * | "xlPasteColumnWidths" | "xlPasteComments"
	 * | "xlPasteFormats" | "xlPasteFormulas" | "xlPasteFormulasAndNumberFormats"
	 * | "xlPasteValues" | "xlPasteValuesAndNumberFormats" )} PasteType
	 * @see office-js-api/Examples/Enumerations/PasteType.js
	 */

	/**
	 * The mathematical operation which will be applied to the copied data.
	 * @typedef {("xlPasteSpecialOperationAdd" | "xlPasteSpecialOperationDivide" | "xlPasteSpecialOperationMultiply"|
	 * "xlPasteSpecialOperationNone" | "xlPasteSpecialOperationSubtract" )} PasteSpecialOperation
	 * @see office-js-api/Examples/Enumerations/PasteSpecialOperation.js
	 */

	/**
	* Specifies how to shift cells to replace deleted cells.
	* @typedef {("up" | "left")} DeleteShiftDirection
	* @see office-js-api/Examples/Enumerations/DeleteShiftDirection.js
	*/

	/**
	 * Class representing a base class for the color types.
	 * @constructor
	 */
	function ApiColor(color) {
		this.color = color;
	}
	/**
	 * Returns a color value in RGB format.
	 * @memberof ApiColor
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiColor/Methods/GetRGB.js
	 */
	ApiColor.prototype.GetRGB = function () {
		if (!this.color) {
			return 0;
		}
		return this.color.getRgb();
	};

	/**
	 * Class representing a name.
	 * @constructor
	 * @property {string} Name - Sets a name to the active sheet.
	 * @property {string} RefersTo - Returns or sets a formula that the name is defined to refer to.
	 * @property {ApiRange} RefersToRange - Returns the ApiRange object by reference.
	 */
	function ApiName(DefName) {
		this.DefName = DefName;
	}

	/**
	 * Class representing a comment.
	 * @constructor
	 * @property {string} Text - Returns or sets the comment text.
	 * @property {string} Id - Returns the current comment ID.
	 * @property {string} AuthorName - Returns or sets the comment author's name.
	 * @property {string} UserId - Returns or sets the user ID of the comment author.
	 * @property {boolean} Solved - Checks if a comment is solved or not or marks a comment as solved.
	 * @property {number | string} TimeUTC - Returns or sets the timestamp of the comment creation in UTC format.
	 * @property {number | string} Time - Returns or sets the timestamp of the comment creation in the current time zone format.
	 * @property {string} QuoteText - Returns the quote text of the current comment.
	 * @property {Number} RepliesCount - Returns a number of the comment replies.
	 */
	function ApiComment(comment, wb) {
		this.Comment = comment.clone();
		this.WB = wb;
	}

	/**
	 * Class representing a comment reply.
	 * @constructor
	 * @property {string} Text - Returns or sets the comment reply text.
	 * @property {string} AuthorName - Returns or sets the comment reply author's name.
	 * @property {string} UserId - Returns or sets the user ID of the comment reply author.
	 * @property {number | string} TimeUTC - Returns or sets the timestamp of the comment reply creation in UTC format.
	 * @property {number | string} Time - Returns or sets the timestamp of the comment reply creation in the current time zone format.
	 */
	function ApiCommentReply(oParentComm, oCommentReply) {
		this.Comment = oParentComm;
		this.Data = oCommentReply;
	}

	/**
	 * Class representing the areas.
	 * @constructor
	 * @property {number} Count - Returns a value that represents the number of objects in the collection.
	 * @property {ApiRange} Parent - Returns the parent object for the specified collection.
	 */
	function ApiAreas(items, parent) {
		this.Items = [];
		this._parent = parent;
		for (var i = 0; i < items.length; i++) {
			this.Items.push(new ApiRange(items[i]));
		}
	}

	/**
	 * Class representing a PivotTable.
	 * @constructor
	 * @property {string} Name - Returns or sets name of the PivotTable.
	 * @property {boolean} ColumnGrand - Returns or sets PivotTable Grand Totals settings for columns.
	 * @property {boolean} RowGrand - Returns or sets PivotTable Grand Totals settings for rows.
	 * @property {boolean} DisplayFieldCaptions  - Returns or set show field headers for rows and columns.
	 * @property {string} Title - Returns or sets the title of the PivotTable.
	 * @property {string} Description - Returns or sets the description of the PivotTable.
	 * @property {string} StyleName - Returns or sets the style name of the PivotTable.
	 * @property {boolean} ShowTableStyleRowHeaders - Returns or sets the row headers style for PivotTable.
	 * @property {boolean} ShowTableStyleColumnHeaders - Returns or sets the column headers style for PivotTable.
	 * @property {boolean} ShowTableStyleRowStripes - Returns or sets the banded rows style for PivotTable.
	 * @property {boolean} ShowTableStyleColumnStripes - Returns or sets the banded columns style for PivotTable.
	 * @property {ApiRange} Source - Returns or sets source for the PivotTable.
	 * @property {ApiRange | null} ColumnRange - Returns a Range object that represents the range that contains the column area in the PivotTable report.
	 * @property {ApiRange | null} RowRange - Returns a Range object that represents the range that contains the row area in the PivotTable report.
	 * @property {ApiRange} DataBodyRange - Returns a Range object that represents the range of values in a PivotTable.
	 * @property {ApiRange | null} TableRange1 - Returns a Range object that represents the range containing the entire PivotTable report, but doesn't include page fields.
	 * @property {ApiRange | null} TableRange2 - Returns a Range object that represents the range containing the entire PivotTable report, including page fields
	 * @property {string} GrandTotalName - Sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report.
	 * @property {boolean} RepeatAllLabels - Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.
	 * @property {object} RowAxisLayout - Sets the way the specified PivotTable items appear—in table format or in outline format.
	 * @property {boolean} LayoutBlankLine - Sets PivotTable setting insert blank rows after each item.
	 * @property {boolean} LayoutSubtotals - Sets PivotTable setting show subtotals.
	 * @property {number} SubtotalLocation - Sets PivotTable setting layout subtotal location.
	 * @property {ApiPivotField[]} PivotFields - Returns all PivotFields in PivotTable.
	 * @property {ApiPivotField[]} ColumnFields - Returns array that are currently showing as column fields in PivotTable.
	 * @property {ApiPivotField[]} DataFields - Returns array that are currently showing values fields in PivotTable.
	 * @property {ApiPivotField[]} HiddenFields - Returns array that represents all not added fields in PivotTable.
	 * @property {ApiPivotField[]} VisibleFields - Returns array that represents all not added fields in PivotTable.
	 * @property {ApiPivotField[]} PageFields - Returns array that are currently showing as page fields in PivotTable.
	 * @property {ApiPivotField[]} RowFields - Returns array that are currently showing as row fields in PivotTable.
	 */
	function ApiPivotTable(pivot, api) {
		/** @type {CT_pivotTableDefinition} */
		this.pivot = pivot;
		this.api = api;
	}

	/**
	 * Class representing a PivotTable field.
	 * @constructor
	 * @property {number} Position - Returns or sets a value that represents the position of the field (first, second, third, and so on) among all the fields in its orientation (Rows, Columns, Pages, Data).
	 * @property {number} Orientation - A pivot field orientation value that represents the location of the field in the specified PivotTable report.
	 * @property {string} Caption - Returns or sets a value that represents the label text for the pivot field.
	 * @property {string} Name - Returns or sets a value representing the name of the object.
	 * @property {string} Value - Returns or sets a value representing the name of the object.
	 * @property {string} SourceName - Returns source name.
	 * @property {number} Index - Returns index.
	 * @property {ApiPivotTable} Table - Returns ApiPivotField.
	 * @property {ApiPivotTable} Parent - Returns parent object to this field.
	 * @property {boolean} LayoutCompactRow - Returns or sets whether or not a PivotField is compacted.
	 * @property {number} LayoutForm - Returns or sets the way the specified PivotTable items appear—in table format or in outline format.
	 * @property {boolean} LayoutPageBreak - Returns or set the flag that represents page break is inserted after each field.
	 * @property {boolean} ShowingInAxis - Returns if the PivotField is currently visible in the PivotTable or not.
	 * @property {boolean} RepeatLabels - setting repeat items labels at each row.
	 * @property {boolean} LayoutBlankLine - Returns or sets setting insert blank rows after each item.
	 * @property {boolean} ShowAllItems - Returns or sets setting show items with no data.
	 * @property {boolean} LayoutSubtotals - Returns or sets setting show subtotals.
	 * @property {number} LayoutSubtotalLocation - Returns or sets setting layout subtotal location.
	 * @property {string} SubtotalName - Returns or sets text string label displayed in the subtotal column or row heading in the specified PivotTable report.
	 * @property {object} Subtotals - Returns or sets subtotals.
	 * @property {number} Formula - Returns or sets a value that represents the object's formula.
	 * @property {boolean} DragToColumn - Returns or sets the specified field, can be dragged to the column position or not.
	 * @property {boolean} DragToRow - Returns or sets the specified field, can be dragged to the row position or not.
	 * @property {boolean} DragToData - Returns or sets the specified field, can be dragged to the data position or not.
	 * @property {boolean} DragToPage - Returns or sets the specified field, can be dragged to the page position or not.
	 * @property {string | null} NumberFormat - Returns or sets a value that represents the format code for the object.
	 * @property {string | number} CurrentPage - Returns the current page showing for the page field (valid only for page fields).
	 */
	function ApiPivotField(table, index, pivotField) {
		/** @type {ApiPivotTable} */
		this.table = table;
		/** @type {number} */
		this.index = index;
		/** @type {CT_PivotField} */
		this.pivotField = pivotField;
	}

	/**
	 * Class representing a PivotTable data field.
	 * @constructor
	 * @extends ApiPivotField
	 * @property {DataConsolidateFunctionType} Function - Returns the function performed in the data field.
	 * @property {number} Position - Returns a value that represents the position of the field in category.
	 * @property {PivotFieldOrientationType} Orientation - Returns a data field orientation value that represents the location.
	 * @property {string} Name - Returns a value representing the name of the object.
	 * @property {string} Value - Returns a value representing the value of the object.
	 * @property {string} Caption - Returns a value that represents the label text for the data field.
	 * @property {string | null} NumberFormat - Returns a value that represents the format code for the object.
	 * @property {number} Index - Returns index of the data field.
	 * @property {ApiPivotField} PivotField - Returns the pivotField from which the data field was created.
	 */
	function ApiPivotDataField(table, dataIndex, dataField) {
		const pivotIndex = dataField.asc_getIndex();
		const pivotField = table.pivot.asc_getPivotFields()[pivotIndex];
		ApiPivotField.call(this, table, pivotIndex, pivotField);
		/** @type {number} */
		this.dataIndex = dataIndex;
		/** @type {CT_DataField} */
		this.dataField = dataField;

	}
	ApiPivotDataField.prototype = Object.create(ApiPivotField.prototype);
	ApiPivotDataField.prototype.constructor = ApiPivotDataField;

	/**
	 * Class representing a PivotTable field item.
	 * @constructor
	 * @property {string} Name - Returns the name of Pivot Item.
	 * @property {string} Caption - Returns the caption of Pivot Item.
	 * @property {string} Value - Returns the value of Pivot Item.
	 * @property {string} Parent - Returns the parent of Pivot Item.
	 * @property {string} Field - Returns the field of Pivot Item.
	 */
	function ApiPivotItem(field, item) {
		/** @type{ApiPivotField} */
		this.field = field;
		/** @type{CT_Item} */
		this.pivotItem = item;
	}


	/**
	 * Class representing characters in an object that contains text.
	 * @constructor
	 * @property {number} Count - The number of characters in the collection.
	 * @property {ApiRange} Parent - The parent object of the specified characters.
	 * @property {string} Caption - The text of the specified range of characters.
	 * @property {string} Text - The string value representing the text of the specified range of characters.
	 * @property {ApiFont} Font - The font of the specified characters.
	 */
	function ApiCharacters(options, parent) {
		this._options = options;
		this._parent = parent;
	}

	/**
	 * Class that contains the font attributes (font name, font size, color, and so on).
	 * @constructor
	 * @property {ApiCharacters} Parent - The parent object of the specified font object.
	 * @property {boolean | null} Bold - The font bold property.
	 * @property {boolean | null} Italic - The font italic property.
	 * @property {number | null} Size - The font size property.
	 * @property {boolean | null} Strikethrough - The font strikethrough property.
	 * @property {string | null} Underline - The font type of underline.
	 * @property {boolean | null} Subscript - The font subscript property.
	 * @property {boolean | null} Superscript - The font superscript property.
	 * @property {string | null} Name - The font name.
	 * @property {ApiColor | null} Color - The font color property.
	 */
	function ApiFont(object) {
		this._object = object;
	}

	/**
	 * Class representing freeze panes.
	 * @constructor
	 */
	function ApiFreezePanes(ws) {
		this.ws = ws;
	}

	/**
	 * Returns a class formatted according to the instructions contained in the format expression.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} expression - Any valid expression.
	 * @param {string} [format] - A valid named or user-defined format expression.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/Format.js
	 */
	Api.prototype.Format = function (expression, format) {
		format = null == format ? '' : format;
		return AscCommonExcel.cTEXT.prototype.Calculate([checkFormat(expression), new AscCommonExcel.cString(format)])
			.getValue();
	};


	/**
	 * Creates a new custom function.
	 * The description of the function parameters and result is specified using JSDoc. The <em>@customfunction</em> tag is required in JSDoc.
	 * Parameters and results can be specified as the <em>number / string / bool / any / number[][] / string[][] / bool[][] / any[][]</em> types.
	 * Parameters can be required or optional. A user can also set a default value.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {Function} fCustom - A new function for calculating.
	 */
	// Example with description:
	// Calculates the sum of the specified numbers.
	// @customfunction
	// @param {number} first Required first number.
	// @param {number} [second] Optional second number to add.
	// @returns {number} The sum of the numbers.
	// Api.AddCustomFunction(function add(first, second) {
	//     if (second === null) {
	//         second = 0;
	//     }
	//     return first + second;
	// })
	Api.prototype.AddCustomFunction = function (fCustom) {
		// get parsedJSDoc from a macros (we receive it from the Api class)
		// take the first element and validate it
		const parsedJSDoc = this.parsedJSDoc && this.parsedJSDoc.shift();
		const isValidJsDoc = parsedJSDoc ? private_ValidateParamsForCustomFunction(parsedJSDoc) : false;
		//const isValidOptions = options ? private_ValidateParamsForCustomFunction(options) : false;
		if (!isValidJsDoc/* && !isValidOptions*/) {
			throwException(new Error('Invalid parameters type in JSDOC or options.'));
		}
		// remove it from this class and use it from the variable (only if it was the last)
		// we don't remove it immediately, because we can have there data for another function
		if (!this.parsedJSDoc.length) {
			delete this.parsedJSDoc;
		}


		// now we have to decide what we're going to use (make the priority order) - parsedJSDoc or options


		//1. jsdoc params:
		//
		//  * Calculates the sum of the specified numbers
		//  * @customfunction
		//  * @param {number} first First number.
		//  * @param {number} second Second number.
		//  * @param {number} [third] Third number to add. If omitted, third = 0.
		//  * @returns {number} The sum of the numbers.
		//
		/*Api.AddCustomFunction(function add(first, second, third) {
			if (third === null) {
				third = 0;
			}
			return first + second + third;
		})*/

		//2. object params - removed it at the moment
		/*

		(function()
		{
    		function add(first, second, third) {
				if (third === null) {
					third = 0;
				}
				return first + second + third;
			}
			Api.AddCustomFunction(add,
				{
					"params":[
						{
							 "defaultValue": "",
							 "description": "First number. *",
							 "name": "first",
							 "optional": false,
							 "parentName": "",
							 "type": "number" // "string", "bool"
						 },
						 {
							 "defaultValue": "",
							 "description": "Second number. *",
							 "name": "second",
							 "optional": false,
							 "parentName": "",
							 "type": "number"
						 },
						 {
							 "defaultValue": "",
							 "description": "Second number. *",
							 "name": "second",
							 "optional": true,
							 "parentName": "",
							 "type": "number"
						 }
					 ]
				}
			);
		})();*/

		this.addCustomFunction(fCustom, parsedJSDoc/*isValidJsDoc ? parsedJSDoc : options*/);
	};

	/**
	 * Register a new custom functions library (see SetCustomFunctions plugin method).
	 * The description of the function parameters and result is specified using JSDoc. The <em>@customfunction</em> tag is required in JSDoc.
	 * Parameters and results can be specified as the <em>number / string / bool / any / number[][] / string[][] / bool[][] / any[][]</em> types.
	 * Parameters can be required or optional. A user can also set a default value.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - A name of library
	 * @param {Function} Func - A custom functions library code.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddCustomFunction.js
	 */
	Api.prototype.AddCustomFunctionLibrary = function(sName, Func) {
		this.addCustomFunctionsLibrary(sName, Func);
	};

	/**
	 * Removes a custom function.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The name of a custom function.
	 * @returns {boolean} - returns false if such a function does not exist.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/RemoveCustomFunction.js
	 */
	Api.prototype.RemoveCustomFunction = function (sName) {
		return this.removeCustomFunction(sName);
	};
	/**
	 * Clears all custom functions.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {boolean} - returns false if such functions do not exist.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/ClearCustomFunctions.js
	 */
	Api.prototype.ClearCustomFunctions = function () {
		return this.clearCustomFunctions();
	};

	/**
	 * Creates a new worksheet. The new worksheet becomes the active sheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The name of a new worksheet.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddSheet.js
	 */
	Api.prototype.AddSheet = function (sName) {
		if (this.GetSheet(sName))
			throwException(new Error('Worksheet with such a name already exists.'));
		else
			this.asc_addWorksheet(sName);
	};

	/**
	 * Returns a sheet collection that represents all the sheets in the active workbook.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheet[]}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetSheets.js
	 */
	Api.prototype.GetSheets = function () {
		var result = [];
		for (var i = 0; i < this.wbModel.getWorksheetCount(); ++i) {
			result.push(new ApiWorksheet(this.wbModel.getWorksheet(i)));
		}
		return result;
	};
	Object.defineProperty(Api.prototype, "Sheets", {
		get: function () {
			return this.GetSheets();
		}
	});

	/**
	 * Sets a locale to the document.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {number} LCID - The locale specified.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/SetLocale.js
	 */
	Api.prototype.SetLocale = function (LCID) {
		this.asc_setLocale(LCID, null, null);
	};

	/**
	 * Returns the current locale ID.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetLocale.js
	 */
	Api.prototype.GetLocale = function () {
		return this.asc_getLocale();
	};

	/**
	 * Returns an object that represents the active sheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheet}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetActiveSheet.js
	 */
	Api.prototype.GetActiveSheet = function () {
		var index = this.wbModel.getActive();
		return new ApiWorksheet(this.wbModel.getWorksheet(index));
	};
	Object.defineProperty(Api.prototype, "ActiveSheet", {
		get: function () {
			return this.GetActiveSheet();
		}
	});

	/**
	 * Returns an object that represents a sheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string | number} nameOrIndex - Sheet name or sheet index.
	 * @returns {ApiWorksheet | null}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetSheet.js
	 */
	Api.prototype.GetSheet = function (nameOrIndex) {
		var ws = ('string' === typeof nameOrIndex) ? this.wbModel.getWorksheetByName(nameOrIndex) :
			this.wbModel.getWorksheet(nameOrIndex);
		return ws ? new ApiWorksheet(ws) : null;
	};

	/**
	 * Returns a list of all the available theme colors for the spreadsheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {string[]}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetThemesColors.js
	 */
	Api.prototype.GetThemesColors = function () {
		var result = [];
		AscCommon.g_oUserColorScheme.forEach(function (item) {
			result.push(item.get_name());
		});

		return result;
	};

	/**
	 * Sets the theme colors to the current spreadsheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sTheme - The color scheme that will be set to the current spreadsheet.
	 * @returns {boolean} - returns false if sTheme isn't a string.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/SetThemeColors.js
	 */
	Api.prototype.SetThemeColors = function (sTheme) {
		if ('string' === typeof sTheme) {
			this.wbModel.changeColorScheme(sTheme);
			return true;
		}
		return false;
	};

	/**
	 * Creates a new history point.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateNewHistoryPoint.js
	 */
	Api.prototype.CreateNewHistoryPoint = function () {
		History.Create_NewPoint();
	};

	/**
	 * Creates an RGB color setting the appropriate values for the red, green and blue color components.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiColor}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateColorFromRGB.js
	 */
	Api.prototype.CreateColorFromRGB = function (r, g, b) {
		return new ApiColor(AscCommonExcel.createRgbColor(r, g, b));
	};

	/**
	 * Creates a color selecting it from one of the available color presets.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {PresetColor} sPresetColor - A preset selected from the list of the available color preset names.
	 * @returns {ApiColor}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/CreateColorByName.js
	 */
	Api.prototype.CreateColorByName = function (sPresetColor) {
		var rgb = AscFormat.mapPrstColor[sPresetColor];
		return new ApiColor(AscCommonExcel.createRgbColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF));
	};

	/**
	 * Returns the ApiRange object that represents the rectangular intersection of two or more ranges. If one or more ranges from a different worksheet are specified, an error will be returned.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} Range1 - One of the intersecting ranges. At least two Range objects must be specified.
	 * @param {ApiRange} Range2 - One of the intersecting ranges. At least two Range objects must be specified.
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/Intersect.js
	 */
	Api.prototype.Intersect = function (Range1, Range2) {
		let result = null;
		if (Range1.GetWorksheet().Id === Range2.GetWorksheet().Id) {
			var res = Range1.range.bbox.intersection(Range2.range.bbox);
			if (!res) {
				logError(new Error('Ranges do not intersect.'));
			} else {
				result = new ApiRange(this.GetActiveSheet().worksheet.getRange3(res.r1, res.c1, res.r2, res.c2));
			}
		} else {
			logError(new Error('Ranges should be from one worksheet.'));
		}
		return result;
	};

	/**
	 * Returns an object that represents the selected range.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetSelection.js
	 */
	Api.prototype.GetSelection = function () {
		return this.GetActiveSheet().GetSelection();
	};
	Object.defineProperty(Api.prototype, "Selection", {
		get: function () {
			return this.GetSelection();
		}
	});

	/**
	 * Adds a new name to a range of cells.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The range name.
	 * @param {string} sRef - The reference to the specified range. It must contain the sheet name, followed by sign ! and a range of cells.
	 * Example: "Sheet1!$A$1:$B$2".
	 * @param {boolean} isHidden - Defines if the range name is hidden or not.
	 * @returns {boolean} - returns false if sName or sRef are invalid.
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddDefName.js
	 */
	Api.prototype.AddDefName = function (sName, sRef, isHidden) {
		return private_AddDefName(this.wbModel, sName, sRef, null, isHidden);
	};

	/**
	 * Returns the ApiName object by the range name.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} defName - The range name.
	 * @returns {ApiName}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetDefName.js
	 */
	Api.prototype.GetDefName = function (defName) {
		if (defName && typeof defName === "string") {
			defName = this.wbModel.getDefinesNames(defName);
		}
		return new ApiName(defName);
	};

	/**
	 * Saves changes to the specified document.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/Api/Methods/Save.js
	 */
	Api.prototype.Save = function () {
		this.SaveAfterMacros = true;
	};

	/**
	 * Returns the ApiRange object by the range reference.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - The range of cells from the current sheet.
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetRange.js
	 */
	Api.prototype.GetRange = function (sRange) {
		var ws;
		var res = AscCommon.parserHelp.parse3DRef(sRange);
		if (res) {
			ws = this.wbModel.getWorksheetByName(res.sheet);
			sRange = res.range;
		} else {
			ws = this.wbModel.getActiveWs();
		}
		return new ApiRange(ws ? ws.getRange2(sRange) : null);
	};

	/**
	 * Returns the ApiWorksheetFunction object.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheetFunction}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetWorksheetFunction.js
	 */
	Api.prototype.GetWorksheetFunction = function () {
		if (!this.oWorksheetFunction) {
			this.oWorksheetFunction = new ApiWorksheetFunction(this);
			//this.oWorksheetFunction.init();
		}
		return this.oWorksheetFunction;
	};
	Object.defineProperty(Api.prototype, "WorksheetFunction", {
		get: function () {
			return this.GetWorksheetFunction();
		}
	});

	//vba + js
	let supportedFunctionsMap = {
		"ABS": "ABS",
		"ACCRINT": "AccrInt",
		"ACCRINTM": "AccrIntM",
		"ACOS": "Acos",
		"ACOSH": "Acosh",
		"ACOT": "Acot",
		"ACOTH": "Acoth",
		"AGGREGATE": "Aggregate",//not js
		"AMORDEGRC": "AmorDegrc",
		"AMORLINC": "AmorLinc",
		"AND": "And",
		"ARABIC": "Arabic",
		"ASC": "Asc",
		"ASIN": "Asin",
		"ASINH": "Asinh",
		"ATAN2": "Atan2",
		"ATANH": "Atanh",
		"AVEDEV": "AveDev",
		"AVERAGE": "Average",
		"AVERAGEIF": "AverageIf",
		"AVERAGEIFS": "AverageIfs",
		//"BAHTTEXT": "BahtText",//not support by OO
		"BASE": "Base",
		"BESSELI": "BesselI",
		"BESSELJ": "BesselJ",
		"BESSELK": "BesselK",
		"BESSELY": "BesselY",
		"BETA.DIST": "Beta.Dist",
		"BETA.INV": "Beta.Inv",
		"BETADIST": "BetaDist",//not js
		"BETAINV": "BetaInv",//not js
		"BIN2DEC": "Bin2Dec",
		"BIN2HEX": "Bin2Hex",
		"BIN2OCT": "Bin2Oct",
		"BINOM.DIST": "Binom.Dist",
		"BINOM.DIST.RANGE": "Binom.Dist.Range",
		"BINOM.INV": "Binom.Inv",
		"BINOMDIST": "BinomDist",//not js
		"BITAND": "Bitand",
		"BITLSHIFT": "Bitlshift",
		"BITOR": "Bitor",
		"BITRSHIFT": "Bitrshift",
		"BITXOR": "Bitxor",
		"CEILING": "Ceiling",//not js
		"CEILING.MATH": "Ceiling.Math",
		"CEILING.PRECISE": "Ceiling.Precise",
		"CHIDIST": "ChiDist",//not js
		"CHIINV": "ChiInv",//not js
		"CHISQ.DIST": "ChiSq.Dist",
		"CHISQ.DIST.RT": "ChiSq.Dist.RT",
		"CHISQ.INV": "ChiSq.Inv",
		"CHISQ.INV.RT": "ChiSq.Inv.RT",
		"CHISQ.TEST": "ChiSq.Test",//not js
		"CHITEST": "ChiTest",//not js
		"CHOOSE": "Choose",
		"CLEAN": "Clean",
		"COMBIN": "Combin",
		"COMBINA": "Combina",
		"COMPLEX": "Complex",
		"CONFIDENCE": "Confidence",//not js
		"CONFIDENCE.NORM": "Confidence.Norm",
		"CONFIDENCE.T": "Confidence.T",
		"CONVERT": "Convert",
		"CORREL": "Correl",//not js
		"COSH": "Cosh",
		"COT": "Cot",
		"COTH": "Coth",
		"COUNT": "Count",
		"COUNTA": "CountA",
		"COUNTBLANK": "CountBlank",
		"COUNTIF": "CountIf",
		"COUNTIFS": "CountIfs",
		"COUPDAYBS": "CoupDayBs",
		"COUPDAYS": "CoupDays",
		"COUPDAYSNC": "CoupDaysNc",
		"COUPNCD": "CoupNcd",
		"COUPNUM": "CoupNum",
		"COUPPCD": "CoupPcd",
		"COVAR": "Covar",//not js
		"COVARIANCE.P": "Covariance.P",//not js
		"COVARIANCE.S": "Covariance.S",//not js
		"CRITBINOM": "CritBinom",//not js
		"CSC": "Csc",
		"CSCH": "Csch",
		"CUMIPMT": "CumIPmt",
		"CUMPRINC": "CumPrinc",
		"DAVERAGE": "DAverage",
		"DAYS": "Days",
		"DAYS360": "Days360",
		"DB": "Db",
		//"DBCS": "Dbcs", //not support by OO
		"DCOUNT": "DCount",
		"DCOUNTA": "DCountA",
		"DDB": "Ddb",
		"DEC2BIN": "Dec2Bin",
		"DEC2HEX": "Dec2Hex",
		"DEC2OCT": "Dec2Oct",
		"DECIMAL": "Decimal",
		"DEGREES": "Degrees",
		"DELTA": "Delta",
		"DEVSQ": "DevSq",
		"DGET": "DGet",
		"DISC": "Disc",
		"DMAX": "DMax",
		"DMIN": "DMin",
		"DOLLAR": "Dollar",
		"DOLLARDE": "DollarDe",
		"DOLLARFR": "DollarFr",
		"DPRODUCT": "DProduct",
		"DSTDEV": "DStDev",
		"DSTDEVP": "DStDevP",
		"DSUM": "DSum",
		"DURATION": "Duration",
		"DVAR": "DVar",
		"DVARP": "DVarP",
		"EDATE": "EDate",
		"EFFECT": "Effect",
		"ENCODEURL": "EncodeUrl",//not js
		"EOMONTH": "EoMonth",
		"ERF": "Erf",
		"ERF.PRECISE": "Erf.Precise",
		"ERFC": "ErfC",
		"ERFC.PRECISE": "ErfC.Precise",
		"EVEN": "Even",
		"EXPON.DIST": "Expon.Dist",
		"EXPONDIST": "ExponDist",//not js
		"F.DIST": "F.Dist",
		"F.DIST.RT": "F.Dist.RT",
		"F.INV": "F.Inv",
		"F.INV.RT": "F.Inv.RT",
		"F.TEST": "F.Test",//not js
		"FACT": "Fact",
		"FACTDOUBLE": "FactDouble",
		"FDIST": "FDist",//not js
		"FILTERXML": "FilterXML",//not js
		"FIND": "Find",
		"FINDB": "FindB",
		"FINV": "FInv",//not js
		"FISHER": "Fisher",
		"FISHERINV": "FisherInv",
		"FIXED": "Fixed",
		"FLOOR": "Floor",//not js
		"FLOOR.MATH": "Floor.Math",
		"FLOOR.PRECISE": "Floor.Precise",
		"FORECAST": "Forecast",//not js
		"FORECAST.ETS": "Forecast.ETS",//not js
		"FORECAST.ETS.CONFINT": "Forecast.ETS.ConfInt",//not js
		"FORECAST.ETS.SEASONALITY": "Forecast.ETS.Seasonality",//not js
		"FORECAST.ETS.STAT": "Forecast.ETS.STAT",//not js
		"FORECAST.LINEAR": "Forecast.Linear",//not js
		"FREQUENCY": "Frequency",//not js
		"FTEST": "FTest",//not js
		"FV": "Fv",
		"FVSCHEDULE": "FVSchedule",
		"GAMMA": "Gamma",
		"GAMMA.DIST": "Gamma.Dist",
		"GAMMA.INV": "Gamma.Inv",
		"GAMMADIST": "GammaDist",//not js
		"GAMMAINV": "GammaInv",//not js
		"GAMMALN": "GammaLn",
		"GAMMALN.PRECISE": "GammaLn.Precise",
		"GAUSS": "Gauss",
		"GCD": "Gcd",
		"GEOMEAN": "GeoMean",
		"GESTEP": "GeStep",
		"GROWTH": "Growth",//not js
		"HARMEAN": "HarMean",
		"HEX2BIN": "Hex2Bin",
		"HEX2DEC": "Hex2Dec",
		"HEX2OCT": "Hex2Oct",
		"HLOOKUP": "HLookup",
		"HYPGEOM.DIST": "HypGeom.Dist",
		"HYPGEOMDIST": "HypGeomDist",//not js
		"IFERROR": "IfError",//not js
		"IFNA": "IfNa",//not js
		"IMABS": "ImAbs",
		"IMAGINARY": "Imaginary",
		"IMARGUMENT": "ImArgument",
		"IMCONJUGATE": "ImConjugate",
		"IMCOS": "ImCos",
		"IMCOSH": "ImCosh",
		"IMCOT": "ImCot",
		"IMCSC": "ImCsc",
		"IMCSCH": "ImCsch",
		"IMDIV": "ImDiv",
		"IMEXP": "ImExp",
		"IMLN": "ImLn",
		"IMLOG10": "ImLog10",
		"IMLOG2": "ImLog2",
		"IMPOWER": "ImPower",
		"IMPRODUCT": "ImProduct",
		"IMREAL": "ImReal",
		"IMSEC": "ImSec",
		"IMSECH": "ImSech",
		"IMSIN": "ImSin",
		"IMSINH": "ImSinh",
		"IMSQRT": "ImSqrt",
		"IMSUB": "ImSub",
		"IMSUM": "ImSum",
		"IMTAN": "ImTan",
		"INDEX": "Index",//not js
		"INTERCEPT": "Intercept",//not js
		"INTRATE": "IntRate",
		"IPMT": "Ipmt",
		"IRR": "Irr",
		"ISERR": "IsErr",
		"ISERROR": "IsError",
		"ISEVEN": "IsEven",
		"ISFORMULA": "IsFormula",
		"ISLOGICAL": "IsLogical",
		"ISNA": "IsNA",
		"ISNONTEXT": "IsNonText",
		"ISNUMBER": "IsNumber",
		"ISO.CEILING": "ISO.Ceiling",
		"ISODD": "IsOdd",
		"ISOWEEKNUM": "IsoWeekNum",
		"ISPMT": "Ispmt",
		"ISTEXT": "IsText",
		"KURT": "Kurt",
		"LARGE": "Large",
		"LCM": "Lcm",
		"LINEST": "LinEst",//not js
		"LN": "Ln",
		"LOG": "Log",
		"LOG10": "Log10",
		"LOGEST": "LogEst",//not js
		"LOGINV": "LogInv",//not js
		"LOGNORM.DIST": "LogNorm.Dist",
		"LOGNORM.INV": "LogNorm.Inv",
		"LOGNORMDIST": "LogNormDist",//not js
		"LOOKUP": "Lookup",
		"MATCH": "Match",
		"MAX": "Max",
		"MDETERM": "MDeterm",//not js
		"MDURATION": "MDuration",
		"MEDIAN": "Median",
		"MIN": "Min",
		"MINVERSE": "MInverse",//not js
		"MIRR": "MIrr",
		"MMULT": "MMult",//not js
		"MODE": "Mode",//not js
		"MODE.MULT": "Mode.Mult",//not js
		"MODE.SNGL": "Mode.Sngl",//not js
		"MROUND": "MRound",
		"MULTINOMIAL": "MultiNomial",
		"MUNIT": "Munit",//not js
		"NEGBINOM.DIST": "NegBinom.Dist",
		"NEGBINOMDIST": "NegBinomDist",//not js
		"NETWORKDAYS": "NetworkDays",
		"NETWORKDAYS.INTL": "NetworkDays.Intl",
		"NOMINAL": "Nominal",
		"NORM.DIST": "Norm.Dist",
		"NORM.INV": "Norm.Inv",
		"NORM.S.DIST": "Norm.S.Dist",
		"NORM.S.INV": "Norm.S.Inv",
		"NORMDIST": "NormDist",//not js
		"NORMINV": "NormInv",//not js
		"NORMSDIST": "NormSDist",//not js
		"NORMSINV": "NormSInv",//not js
		"NPER": "NPer",
		"NPV": "Npv",
		"NUMBERVALUE": "NumberValue",
		"OCT2BIN": "Oct2Bin",
		"OCT2DEC": "Oct2Dec",
		"OCT2HEX": "Oct2Hex",
		"ODD": "Odd",
		"ODDFPRICE": "OddFPrice",
		"ODDFYIELD": "OddFYield",
		"ODDLPRICE": "OddLPrice",
		"ODDLYIELD": "OddLYield",
		"OR": "Or",
		"PDURATION": "PDuration",
		"PEARSON": "Pearson",//not js
		"PERCENTILE": "Percentile",//not js
		"PERCENTILE.EXC": "Percentile.Exc",
		"PERCENTILE.INC": "Percentile.Inc",
		"PERCENTRANK": "PercentRank",//not js
		"PERCENTRANK.EXC": "PercentRank.Exc",
		"PERCENTRANK.INC": "PercentRank.Inc",
		"PERMUT": "Permut",
		"PERMUTATIONA": "Permutationa",
		"PHI": "Phi",
		"PHONETIC": "Phonetic",//not js
		"PI": "Pi",
		"PMT": "Pmt",
		"POISSON": "Poisson",//not js
		"POISSON.DIST": "Poisson.Dist",
		"POWER": "Power",
		"PPMT": "Ppmt",
		"PRICE": "Price",
		"PRICEDISC": "PriceDisc",
		"PRICEMAT": "PriceMat",
		"PROB": "Prob",//not js
		"PRODUCT": "Product",
		"PROPER": "Proper",
		"PV": "Pv",
		"QUARTILE": "Quartile",//not js
		"QUARTILE.EXC": "Quartile.Exc",
		"QUARTILE.INC": "Quartile.Inc",
		"QUOTIENT": "Quotient",
		"RADIANS": "Radians",
		"RANDBETWEEN": "RandBetween",
		"RANK": "Rank",//not js
		"RANK.AVG": "Rank.Avg",
		"RANK.EQ": "Rank.Eq",
		"RATE": "Rate",
		"RECEIVED": "Received",
		"REPLACE": "Replace",
		"REPLACEB": "ReplaceB",
		"REPT": "Rept",
		"ROMAN": "Roman",
		"ROUND": "Round",
		"ROUNDDOWN": "RoundDown",
		"ROUNDUP": "RoundUp",
		"RRI": "Rri",
		"RSQ": "RSq",//not js
		"RTD": "RTD",//not js
		"SEARCH": "Search",//not js
		"SEARCHB": "SearchB",//not js
		"SEC": "Sec",
		"SECH": "Sech",
		"SERIESSUM": "SeriesSum",
		"SINH": "Sinh",
		"SKEW": "Skew",
		"SKEW.P": "Skew.p",
		"SLN": "Sln",
		"SLOPE": "Slope",//not js
		"SMALL": "Small",
		"SQRTPI": "SqrtPi",
		"STANDARDIZE": "Standardize",
		"STDEV": "StDev",//not js
		"STDEV.P": "StDev.P",//not js
		"STDEV.S": "StDev.S",
		"STDEVP": "StDevP",
		"STEYX": "StEyx",//not js
		"SUBSTITUTE": "Substitute",
		"SUBTOTAL": "Subtotal",
		"SUM": "Sum",
		"SUMIF": "SumIf",
		"SUMIFS": "SumIfs",
		"SUMPRODUCT": "SumProduct",//not js
		"SUMSQ": "SumSq",
		"SUMX2MY2": "SumX2MY2",//not js
		"SUMX2PY2": "SumX2PY2",//not js
		"SUMXMY2": "SumXMY2",//not js
		"SYD": "Syd",
		"T.DIST": "T.Dist",
		"T.DIST.2T": "T.Dist.2T",
		"T.DIST.RT": "T.Dist.RT",
		"T.INV": "T.Inv",
		"T.INV.2T": "T.Inv.2T",
		"T.TEST": "T.Test",//not js
		"TANH": "Tanh",
		"TBILLEQ": "TBillEq",
		"TBILLPRICE": "TBillPrice",
		"TBILLYIELD": "TBillYield",
		"TDIST": "TDist",//not js
		"TEXT": "Text",
		"TINV": "TInv",//not js
		"TRANSPOSE": "Transpose",//not js
		"TREND": "Trend",//not js
		"TRIM": "Trim",
		"TRIMMEAN": "TrimMean",
		"TTEST": "TTest",//not js
		"UNICHAR": "Unichar",
		"UNICODE": "Unicode",
		//"USDOLLAR": "USDollar", //not support by OO
		"VAR": "Var",//not js
		"VAR.P": "Var.P",//not js
		"VAR.S": "Var.S",
		"VARP": "VarP",
		"VDB": "Vdb",
		"VLOOKUP": "VLookup",
		"WEBSERVICE": "WebService",//not js
		"WEEKDAY": "Weekday",
		"WEEKNUM": "WeekNum",
		"WEIBULL": "Weibull",//not js
		"WEIBULL.DIST": "Weibull.Dist",
		"WORKDAY": "WorkDay",
		"WORKDAY.INTL": "WorkDay.Intl",
		"XIRR": "Xirr",
		"XNPV": "Xnpv",
		"XOR": "Xor",
		"YEARFRAC": "YearFrac",
		"YIELDDISC": "YieldDisc",
		"YIELDMAT": "YieldMat",
		"Z.TEST": "Z.Test",
		"ZTEST": "ZTest",//not js


		//not in vba, only js:
		//"INIT": "init", //not support by OO
		//"AREAS": "areas", //not support by OO
		"ATAN": "atan",
		"AVERAGEA": "averageA",
		"CHAR": "char",
		"CODE": "code",
		"COLUMNS": "columns",
		"CONCATENATE": "concatenate",
		"COS": "cos",
		"DATE": "date",
		"DATEVALUE": "datevalue",
		"DAY": "day",
		"ECMA.CEILING": "ecma.Ceiling",
		"ERROR.TYPE": "error.Type",
		"EXACT": "exact",
		"EXP": "exp",
		"FALSE": "false",
		"HOUR": "hour",
		"HYPERLINK": "hyperlink",
		"IF": "if",
		"INT": "int",
		"ISREF": "isref",
		"LEFT": "left",
		"LEFTB": "leftb",
		"LEN": "len",
		"LENB": "lenb",
		"LOWER": "lower",
		"MAXA": "maxA",
		"MID": "mid",
		"MIDB": "midb",
		"MINA": "minA",
		"MINUTE": "minute",
		"MOD": "mod",
		"MONTH": "month",
		"N": "n",
		"NA": "na",
		"NOT": "not",
		"NOW": "now",
		"RAND": "rand",
		"RIGHT": "right",
		"RIGHTB": "rightb",
		"ROWS": "rows",
		"SECOND": "second",
		"SHEET": "sheet",
		"SHEETS": "sheets",
		"SIGN": "sign",
		"SIN": "sin",
		"SQRT": "sqrt",
		"STDEVA": "stDevA",
		"STDEVPA": "stDevPA",
		"T": "t",
		"TAN": "tan",
		"TIME": "time",
		"TIMEVALUE": "timevalue",
		"TODAY": "today",
		"TRUE": "true",
		"TRUNC": "trunc",
		"TYPE": "type",
		"UPPER": "upper",
		"VALUE": "value",
		"VARA": "varA",
		"VARPA": "varPA",
		"YEAR": "year",
		"YIELD": "yield"
	};

	/**
	 * Class representing a worksheet function.
	 * @constructor
	 */
	function ApiWorksheetFunction(api) {
		this.api = api;
	}

	// ApiWorksheetFunction.prototype.init = function () {
	// 	let getArgType = function (_arg) {
	//
	// 		let res = "any";
	// 		if (_arg === Asc.c_oAscFormulaArgumentType.number) {
	// 			res = "number"
	// 		} else if (_arg === Asc.c_oAscFormulaArgumentType.text) {
	// 			res = "string"
	// 		} else if (_arg === Asc.c_oAscFormulaArgumentType.reference) {
	// 			res = "ApiRange"
	// 		} else if (_arg === Asc.c_oAscFormulaArgumentType.logical) {
	// 			res = "boolean"
	// 		}
	// 		return res;
	// 	};
	//
	// 	let test = ""
	// 	for (let i in AscCommonExcel.cFormulaFunction) {
	// 		if (supportedFunctionsMap[i]) {
	//
	// 			//if (i === "SERIESSUM") {
	// 				let test1 = "";
	// 				let test2 = ""
	// 				let maxArg = AscCommonExcel.cFormulaFunction[i].prototype.argumentsMax;
	// 				let minArg = AscCommonExcel.cFormulaFunction[i].prototype.argumentsMin;
	// 				if (maxArg < 10) {
	// 					for (let j = 1; j <= maxArg; j++) {
	// 						let test12 = "arg" + (j)
	//
	// 						test1 += "\t * @param {" + (j <= minArg ? "" : "?") + getArgType(AscCommonExcel.cFormulaFunction[i].prototype.argumentsType && AscCommonExcel.cFormulaFunction[i].prototype.argumentsType[j-1]) + "} ";
	// 						if (j <= minArg) {
	// 							test1 += test12;
	// 						} else {
	// 							test1 += test12;
	// 						}
	//
	// 						let argInfo = window.map2[i] && window.map2[i][j - 1];
	// 						if (argInfo) {
	// 							argInfo = argInfo.charAt(0).toUpperCase() + argInfo.slice(1)
	// 						} else {
	// 							argInfo = "";
	// 						}
	//
	// 						test1 += (argInfo !== "" ? " " : "") + argInfo + ".\n";
	//
	// 						test2 += j === maxArg ? test12 : (test12 + ",")
	// 					}
	// 				}
	//
	//
	// 				let funcInfo = window.test0[i] ? " " + window.test0[i].d : "Returns the result of calculating the function"
	//
	// 				test += "/**\n" + "\t *" + funcInfo + ".\n" + "\t * @memberof ApiWorksheetFunction\n" + "\t * @typeofeditors [\"CSE\"]\n" + test1
	// 					+ "\t * @returns {number | string | boolean}\n" + "\t */\n" + "\tApiWorksheetFunction.prototype." + i.replaceAll(".","_")  + "= function (" + test2 + ") {\n" + "\t\treturn this.private_calculateFunction(\"" + i + "\", arguments);\n" + "\t};"
	// 			//}
	//
	// 			test += "\n";
	//
	//
	//
	// 			ApiWorksheetFunction.prototype[i] = function () {
	// 				return this.private_calculateFunction(AscCommonExcel.cFormulaFunction[i].prototype, arguments);
	// 			}
	// 		}
	// 	}
	// 	console.log(test)
	// };


	ApiWorksheetFunction.prototype.private_calculateFunction = function (sFunc, arg) {
		//check
		let func = AscCommonExcel.cFormulaFunction[sFunc].prototype;
		if (!func) {
			return;
		}
		let argsCount = arg.length;
		if (!func.checkArguments(argsCount)) {
			throwException(new Error('Arguments count error.'));
			return null;
		}

		//prepare arguments
		let newArguments = [];
		for (let i = 0; i < argsCount; i++) {
			if ('number' === typeof arg[i]) {
				newArguments.push(new AscCommonExcel.cNumber(arg[i]));
			} else if ('string' === typeof arg[i]) {
				newArguments.push(new AscCommonExcel.cString(arg[i]));
			} else if ('boolean' === typeof arg[i]) {
				newArguments.push(new AscCommonExcel.cBool(arg[i]));
			} else if (arg[i] instanceof ApiRange ) {
				//cArea/cRef/cArea3D/cRef3d
				if (arg[i].range && arg[i].range.bbox && arg[i].range.worksheet) {
					newArguments.push(new AscCommonExcel.cArea3D(arg[i].range.bbox.getName(), arg[i].range.worksheet, arg[i].range.worksheet));
				} else {
					throwException(new Error('Arguments type error.'));
					return null;
				}
			} else {
				throwException(new Error('Arguments type error.'));
				return null;
			}
		}

		//prepare result
		let ws = this.api && this.api.wb && this.api.wb.getWorksheet();
		if (ws) {
			ws = ws.model;
		}
		let result = func.Calculate(newArguments, new Asc.Range(0, 0, 0, 0), null, ws);

		if (!result) {
			throwException(new Error('Result type error.'));
			return null;
		}


		if (AscCommonExcel.cElementType.cell === result.type || AscCommonExcel.cElementType.cell3D === result.type) {
			result = result.getValue();
			if (AscCommonExcel.cElementType.empty === result.type) {
				result = new AscCommonExcel.cNumber(0);
			}
		} else if (AscCommonExcel.cElementType.array === result.type) {
			result = result.getElement(0);
		} else if (AscCommonExcel.cElementType.cellsRange === result.type || AscCommonExcel.cElementType.cellsRange3D === result.type) {
			if (AscCommonExcel.cElementType.cellsRange === result.type) {
				result = result.getValue2(0, 0);
			} else {
				result = result.getValue2(new AscCommon.CellAddress(result.getBBox0().r1, result.getBBox0().c1, 0));
			}
		}

		if (result && result.getValue) {
			result = result.getValue();
		} else {
			throwException(new Error('Result type error.'));
			return null;
		}

		return result;
	};

	ApiWorksheetFunction.prototype.private_simpleTestAllFunctions = function () {
		let obj = Object.getPrototypeOf(this)

		let t = this
		let checkType = function (type) {
			let arg
			if (type === Asc.c_oAscFormulaArgumentType.reference) {
				arg = t.api.GetRange("A1:B2");
			} else if (type === Asc.c_oAscFormulaArgumentType.number) {
				arg = 111;
			} else if (type === Asc.c_oAscFormulaArgumentType.text) {
				arg = "test1"
			} else if (type === Asc.c_oAscFormulaArgumentType.any) {
				arg = t.api.GetRange("A1:B2");
			} else if (type === Asc.c_oAscFormulaArgumentType.logical) {
				arg = true;
			}
			return arg
		}

		for (let i in obj) {
			let arrayTypeFunc = false;

			let iReplace = i.replace("_", ".")
			if (!AscCommonExcel.cFormulaFunction[iReplace]) {
				continue
			}
			if (!AscCommonExcel.cFormulaFunction[iReplace]) {
				console.log(i)
				continue;
			}
			let props = AscCommonExcel.cFormulaFunction[iReplace].prototype
			let args = [];
			for (let j in props.argumentsType) {

				let arg;
				let type;
				if (Array.isArray(props.argumentsType[j])) {
					for (let n  =0 ; n < props.argumentsType[j].length; n++) {
						type = props.argumentsType[j][n]
						arg = checkType(type);
						if (arg == null) {
							console.log("undefined type: " + i)
							arrayTypeFunc = true
							continue;
						}
						args.push(arg);
					}
				} else {
					type = props.argumentsType[j]
					arg = checkType(type);
					if (arg == null) {
						console.log("undefined type: " + i)
						arrayTypeFunc = true
						continue;
					}
					args.push(arg);
				}
			}
			if (!arrayTypeFunc) {
				obj[i].apply(this, args);
			}
		}
	};

	/**
	 * For double-byte character set (DBCS) languages, the function changes full-width (double-byte) characters to half-width (single-byte) characters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text or a reference to a cell containing the text to change.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ASC.js
	 */
	ApiWorksheetFunction.prototype.ASC = function (arg1) {
		return this.private_calculateFunction("ASC", arguments);
	};
	/**
	 * Returns the character specified by the code number from your computer's character set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A number between 1 and 255 specifying a character from the computer character set.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHAR.js
	 */
	ApiWorksheetFunction.prototype.CHAR = function (arg1) {
		return this.private_calculateFunction("CHAR", arguments);
	};
	/**
	 * Removes all the nonprintable characters from the text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - A string from which nonprintable characters will be removed.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CLEAN.js
	 */
	ApiWorksheetFunction.prototype.CLEAN = function (arg1) {
		return this.private_calculateFunction("CLEAN", arguments);
	};
	/**
	 * Returns the code number from your computer's character set for the first character in the specified text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text for which to get the code of the first character.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CODE.js
	 */
	ApiWorksheetFunction.prototype.CODE = function (arg1) {
		return this.private_calculateFunction("CODE", arguments);
	};
	/**
	 * Combines multiple text strings into one text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg_n - Up to 255 data values that will be combined.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CONCATENATE.js
	 */
	ApiWorksheetFunction.prototype.CONCATENATE = function () {
		return this.private_calculateFunction("CONCATENATE", arguments);
	};
	/**
	 * Converts a number to text, using a currency format $#.##.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string} arg1 - A number, a reference to a cell containing a number, or a formula that returns a number.
	 * @param {?number} arg2 - A number of digits to the right of the decimal point. The number is rounded as necessary.
	 * If it is omitted, the function will assume it to be 2.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DOLLAR.js
	 */
	ApiWorksheetFunction.prototype.DOLLAR = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLAR", arguments);
	};
	/**
	 * Checks whether two text strings are exactly the same, and returns <b>true</b> or <b>false</b>. This function is case-sensitive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The first text string.
	 * @param {string} arg2 - The second text string.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EXACT.js
	 */
	ApiWorksheetFunction.prototype.EXACT = function (arg1, arg2) {
		return this.private_calculateFunction("EXACT", arguments);
	};
	/**
	 * Returns the starting position of one text string within another text string. This function is case-sensitive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text to find. Use double quotes (empty text) to match the first character in the search string.
	 * Wildcard characters are not allowed.
	 * @param {string} arg2 - The text containing the text to find.
	 * @param {?number} arg3 - Specifies the character at which to start the search. The first character in the search string is character number 1.
	 * If omitted, this parameter is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FIND.js
	 */
	ApiWorksheetFunction.prototype.FIND = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FIND", arguments);
	};
	/**
	 * Finds the specified substring within another string and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text to find. Use double quotes (empty text) to match the first character in the search string.
	 * Wildcard characters are not allowed.
	 * @param {string} arg2 - The text containing the text to find.
	 * @param {?number} arg3 - Specifies the character at which to start the search. The first character in the search string is character number 1.
	 * If omitted, this parameter is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FINDB.js
	 */
	ApiWorksheetFunction.prototype.FINDB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FINDB", arguments);
	};
	/**
	 * Rounds a number to the specified number of decimals and returns the result as text with or without commas.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number to round and convert to text.
	 * @param {?number} arg2 - The number of digits to the right of the decimal point. If omitted, the function will assume it to be 2.
	 * @param {?boolean} arg3 - Specifies whether do display commas in the returned text (<b>false</b> or omitted) or not (<b>true</b>).
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FIXED.js
	 */
	ApiWorksheetFunction.prototype.FIXED = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FIXED", arguments);
	};
	/**
	 * Returns the specified number of characters from the start of a text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string containing the characters to extract.
	 * @param {?number} arg2 - A number of the substring characters. It must be greater than or equal to 0.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LEFT.js
	 */
	ApiWorksheetFunction.prototype.LEFT = function (arg1, arg2) {
		return this.private_calculateFunction("LEFT", arguments);
	};
	/**
	 * Extracts the substring from the specified string starting from the left character and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string containing the characters to extract.
	 * @param {?number} arg2 - A number of the substring characters, based on bytes.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LEFTB.js
	 */
	ApiWorksheetFunction.prototype.LEFTB = function (arg1, arg2) {
		return this.private_calculateFunction("LEFTB", arguments);
	};
	/**
	 * Returns the number of characters in a text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text whose length will be returned. Spaces are considered as characters.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LEN.js
	 */
	ApiWorksheetFunction.prototype.LEN = function (arg1) {
		return this.private_calculateFunction("LEN", arguments);
	};
	/**
	 * Analyses the specified string and returns the number of characters it contains and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text whose length will be returned. Spaces are considered as characters.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LENB.js
	 */
	ApiWorksheetFunction.prototype.LENB = function (arg1) {
		return this.private_calculateFunction("LENB", arguments);
	};
	/**
	 * Converts all letters in a text string to lowercase.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text to convert to lowercase. The text characters that are not letters are not changed.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOWER.js
	 */
	ApiWorksheetFunction.prototype.LOWER = function (arg1) {
		return this.private_calculateFunction("LOWER", arguments);
	};
	/**
	 * Returns the characters from the middle of a text string, given a starting position and length.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string from which to extract the characters.
	 * @param {number} arg2 - The position of the first character to extract. The first text character is 1.
	 * @param {number} arg3 - A number of the characters to extract.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MID.js
	 */
	ApiWorksheetFunction.prototype.MID = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MID", arguments);
	};
	/**
	 * Extracts the characters from the specified string starting from any position and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string from which to extract the characters.
	 * @param {number} arg2 - The position of the first character to extract. The first text character is 1.
	 * @param {number} arg3 - A number of the characters to extract, based on bytes.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MIDB.js
	 */
	ApiWorksheetFunction.prototype.MIDB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MIDB", arguments);
	};
	/**
	 * Converts text to a number, in a locale-independent way.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The string representing a number to convert.
	 * @param {?string} arg2 - The character used as the decimal separator in the string.
	 * @param {?string} arg3 - The character used as the group separator in the string.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NUMBERVALUE.js
	 */
	ApiWorksheetFunction.prototype.NUMBERVALUE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NUMBERVALUE", arguments);
	};
	/**
	 * Converts a text string to proper case: the first letter in each word to uppercase, and all other letters to lowercase.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text enclosed in quotation marks, a formula that returns text, or a reference to a cell containing text to partially capitalize.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PROPER.js
	 */
	ApiWorksheetFunction.prototype.PROPER = function (arg1) {
		return this.private_calculateFunction("PROPER", arguments);
	};
	/**
	 * Replaces part of a text string with a different text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text where some characters will be replaced.
	 * @param {number} arg2 - The position of the character in the original text that will be replaced with the new text.
	 * @param {number} arg3 - The number of characters in the original text that will be replaced.
	 * @param {string} arg4 - The text that will replace characters in the original text.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/REPLACE.js
	 */
	ApiWorksheetFunction.prototype.REPLACE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("REPLACE", arguments);
	};
	/**
	 * Replaces a set of characters, based on the number of characters and the start position specified, with a new set of characters and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text where some characters will be replaced.
	 * @param {number} arg2 - The position of the character in the original text that will be replaced with the new text.
	 * @param {number} arg3 - The number of characters in the original text that will be replaced, based on bytes.
	 * @param {string} arg4 - The text that will replace characters in the original text.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/REPLACEB.js
	 */
	ApiWorksheetFunction.prototype.REPLACEB = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("REPLACEB", arguments);
	};
	/**
	 * Repeats text a given number of times. Use this function to fill a cell with a number of instances of a text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text that will be repeated.
	 * @param {number} arg2 - A positive number specifying the number of times to repeat text.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/REPT.js
	 */
	ApiWorksheetFunction.prototype.REPT = function (arg1, arg2) {
		return this.private_calculateFunction("REPT", arguments);
	};
	/**
	 * Returns the specified number of characters from the end of a text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string that contains the characters to extract.
	 * @param {?number} arg2 - A number of the substring characters. If it is omitted, the function will assume it to be 1.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RIGHT.js
	 */
	ApiWorksheetFunction.prototype.RIGHT = function (arg1, arg2) {
		return this.private_calculateFunction("RIGHT", arguments);
	};
	/**
	 * Extracts a substring from a string starting from the right-most character, based on the specified number of characters and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text string that contains the characters to extract.
	 * @param {?number} arg2 - A number of the substring characters, based on bytes.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RIGHTB.js
	 */
	ApiWorksheetFunction.prototype.RIGHTB = function (arg1, arg2) {
		return this.private_calculateFunction("RIGHTB", arguments);
	};
	/**
	 * Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text to find. The ? and * wildcard characters can be used. Use ~? and ~* to find the ? and * characters.
	 * @param {string} arg2 - The text where to search for the specified text.
	 * @param {?number} arg3 - The character number in the search text, counting from the left, at which to start searching. If omitted, 1 is used.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SEARCH.js
	 */
	ApiWorksheetFunction.prototype.SEARCH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SEARCH", arguments);
	};
	/**
	 * Returns the location of the specified substring in a string and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text to find. The ? and * wildcard characters can be used. Use ~? and ~* to find the ? and * characters.
	 * @param {string} arg2 - The text where to search for the specified text.
	 * @param {?number} arg3 - The character number in the search text, counting from the left, at which to start searching. If omitted, 1 is used.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SEARCHB.js
	 */
	ApiWorksheetFunction.prototype.SEARCHB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SEARCHB", arguments);
	};
	/**
	 * Replaces existing text with new text in a text string.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text or the reference to a cell containing text in which the characters will be substituted.
	 * @param {string} arg2 - The existing text to replace. If the case of the original text does not match the case of text, the function will not replace the text.
	 * @param {string} arg3 - The text to replace the original text with.
	 * @param {?string} arg4 - Specifies which occurrence of the original text to replace. If omitted, every instance of the original text will be replaced.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUBSTITUTE.js
	 */
	ApiWorksheetFunction.prototype.SUBSTITUTE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SUBSTITUTE", arguments);
	};
	/**
	 * Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 - The value to test.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T.js
	 */
	ApiWorksheetFunction.prototype.T = function (arg1) {
		return this.private_calculateFunction("T", arguments);
	};
	/**
	 * Converts a value to text in a specific number format.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string} arg1 - A number, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
	 * @param {string} arg2 - A number format in the text form from the <b>Number format</b> combo box on the <b>Home</b> tab.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TEXT.js
	 */
	ApiWorksheetFunction.prototype.TEXT = function (arg1, arg2) {
		return this.private_calculateFunction("TEXT", arguments);
	};
	/**
	 * Removes all spaces from a text string except for single spaces between words.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text from which the spaces will be removed.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TRIM.js
	 */
	ApiWorksheetFunction.prototype.TRIM = function (arg1) {
		return this.private_calculateFunction("TRIM", arguments);
	};
	/**
	 * Returns the Unicode character referenced by the given numeric value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The Unicode number representing a character.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/UNICHAR.js
	 */
	ApiWorksheetFunction.prototype.UNICHAR = function (arg1) {
		return this.private_calculateFunction("UNICHAR", arguments);
	};
	/**
	 * Returns the number (code point) corresponding to the first character of the text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The character for which the Unicode value will be returned.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/UNICODE.js
	 */
	ApiWorksheetFunction.prototype.UNICODE = function (arg1) {
		return this.private_calculateFunction("UNICODE", arguments);
	};
	/**
	 * Converts a text string to all uppercase letters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text which will be converted to uppercase, a reference or a text string.
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/UPPER.js
	 */
	ApiWorksheetFunction.prototype.UPPER = function (arg1) {
		return this.private_calculateFunction("UPPER", arguments);
	};
	/**
	 * Converts a text string that represents a number to a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text enclosed in quotation marks or a reference to a cell containing the text which will be converted to a number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VALUE.js
	 */
	ApiWorksheetFunction.prototype.VALUE = function (arg1) {
		return this.private_calculateFunction("VALUE", arguments);
	};
	/**
	 * Returns the average of the absolute deviations of data points from their mean.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[]} args - Up to 255 numeric values for which the average of the absolute deviations will be returned. The first argument is required,
	 * subsequent arguments are optional. Arguments can be numbers or names, arrays or references that contain numbers.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AVEDEV.js
	 */
	ApiWorksheetFunction.prototype.AVEDEV = function () {
		return this.private_calculateFunction("AVEDEV", arguments);
	};
	/**
	 * Returns the average (arithmetic mean) of the specified arguments.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[]} args - Up to 255 numeric values for which the average value will be returned. The first argument is required,
	 * subsequent arguments are optional. Arguments can be numbers or names, arrays or references that contain numbers.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AVERAGE.js
	 */
	ApiWorksheetFunction.prototype.AVERAGE = function () {
		return this.private_calculateFunction("AVERAGE", arguments);
	};
	/**
	 * Returns the average (arithmetic mean) of the specified arguments, evaluating text and <b>false</b> in arguments as 0; <b>true</b> evaluates as 1.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[]} args - Up to 255 numeric values for which the average value will be returned. The first argument is required,
	 * subsequent arguments are optional. Arguments can be numbers or names, arrays or references that contain numbers, text representations of numbers,
	 * or logical values, such as <b>true</b> and <b>false</b>.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AVERAGEA.js
	 */
	ApiWorksheetFunction.prototype.AVERAGEA = function () {
		return this.private_calculateFunction("AVERAGEA", arguments);
	};
	/**
	 * Finds the average (arithmetic mean) for the cells specified by a given condition or criteria.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The range of cells which will be evaluated.
	 * @param {number | string} arg2 - The condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
	 * @param {?ApiRange} arg3 - The actual cells to be used to find the average. If omitted, the cells in the range are used.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AVERAGEIF.js
	 */
	ApiWorksheetFunction.prototype.AVERAGEIF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("AVERAGEIF", arguments);
	};
	/**
	 * Finds the average (arithmetic mean) for the cells specified by a given set of conditions or criteria.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The range of cells which will be evaluated.
	 * @param {number | string} arg2 - The first condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
	 * @param {?ApiRange} arg3 - The actual cells to be used to find the average. If omitted, the cells in the range are used.
	 * @param {number | string} arg4 - Up to 127 additional conditions or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
	 * These arguments are optional.
	 * @param {?ApiRange} arg5 - Up to 127 actual ranges to be used to find the average. If omitted, the cells in the range are used. These arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AVERAGEIFS.js
	 */
	ApiWorksheetFunction.prototype.AVERAGEIFS = function () {
		return this.private_calculateFunction("AVERAGEIFS", arguments);
	};
	/**
	 * Returns the cumulative beta probability density function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value between A and B at which to evaluate the function.
	 * @param {number} arg2 - The alpha parameter of the distribution which must be greater than 0.
	 * @param {number} arg3 - The beta parameter of the distribution which must be greater than 0.
	 * @param {?number} arg4 - An optional lower bound to the interval of x (A). If omitted, it is equal to 0.
	 * @param {?number} arg5 - An optional upper bound to the interval of x (B). If omitted, it is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BETADIST.js
	 */
	ApiWorksheetFunction.prototype.BETADIST = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETADIST", arguments);
	};
	/**
	 * Returns the beta probability distribution function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value between A and B at which to evaluate the function.
	 * @param {number} arg2 - The alpha parameter of the distribution which must be greater than 0.
	 * @param {number} arg3 - The beta parameter of the distribution which must be greater than 0.
	 * @param {boolean} arg4 - Specifies if this is the cumulative distribution function (<b>true</b>) or the probability density function (<b>false</b>).
	 * @param {?number} arg5 - An optional lower bound to the interval of x (A). If omitted, it is equal to 0.
	 * @param {?number} arg6 - An optional upper bound to the interval of x (B). If omitted, it is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BETA_DIST.js
	 */
	ApiWorksheetFunction.prototype.BETA_DIST = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("BETA.DIST", arguments);
	};
	/**
	 * Returns the inverse of the cumulative beta probability density function (BETA_DIST).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the beta distribution.
	 * @param {number} arg2 - The alpha parameter of the distribution which must be greater than 0.
	 * @param {number} arg3 - The beta parameter of the distribution which must be greater than 0.
	 * @param {?number} arg4 - An optional lower bound to the interval of x (A). If omitted, it is equal to 0.
	 * @param {?number} arg5 - An optional upper bound to the interval of x (B). If omitted, it is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BETA_INV.js
	 */
	ApiWorksheetFunction.prototype.BETA_INV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETA.INV", arguments);
	};
	/**
	 * Returns the inverse of the cumulative beta probability density function for a specified beta distribution (BETADIST).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the beta distribution.
	 * @param {number} arg2 - The alpha parameter of the distribution which must be greater than 0.
	 * @param {number} arg3 - The beta parameter of the distribution which must be greater than 0.
	 * @param {?number} arg4 - An optional lower bound to the interval of x (A). If omitted, it is equal to 0.
	 * @param {?number} arg5 - An optional upper bound to the interval of x (B). If omitted, it is equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BETAINV.js
	 */
	ApiWorksheetFunction.prototype.BETAINV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETAINV", arguments);
	};
	/**
	 * Returns the individual term binomial distribution probability.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of successes in trials.
	 * @param {number} arg2 - The number of independent trials.
	 * @param {number} arg3 - The probability of success on each trial.
	 * @param {boolean} arg4 - Specifies if this is the cumulative distribution function (<b>true</b>) or the probability mass function (<b>false</b>).
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BINOMDIST.js
	 */
	ApiWorksheetFunction.prototype.BINOMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOMDIST", arguments);
	};
	/**
	 * Returns the individual term binomial distribution probability.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of successes in trials.
	 * @param {number} arg2 - The number of independent trials.
	 * @param {number} arg3 - The probability of success on each trial.
	 * @param {boolean} arg4 - Specifies if this is the cumulative distribution function (<b>true</b>) or the probability mass function (<b>false</b>).
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BINOM_DIST.js
	 */
	ApiWorksheetFunction.prototype.BINOM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOM.DIST", arguments);
	};
	/**
	 * Returns the probability of a trial result using a binomial distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of independent trials.
	 * @param {number} arg2 - The probability of success on each trial.
	 * @param {number} arg3 - The minimum number of successes in the trials to calculate probability for, a numeric value greater than or equal to 0.
	 * @param {?number} arg4 - The maximum number of successes in the trials to calculate probability for,
	 * a numeric value greater than the minimum number of successes and less than or equal to trials.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BINOM_DIST_RANGE.js
	 */
	ApiWorksheetFunction.prototype.BINOM_DIST_RANGE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOM.DIST.RANGE", arguments);
	};
	/**
	 * Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of Bernoulli trials.
	 * @param {number} arg2 - The probability of success on each trial, a number between 0 and 1 inclusive.
	 * @param {number} arg3 - The criterion value, a number between 0 and 1 inclusive.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BINOM_INV.js
	 */
	ApiWorksheetFunction.prototype.BINOM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("BINOM.INV", arguments);
	};
	/**
	 * Returns the right-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which the distribution will be evaluated, a nonnegative number.
	 * @param {number} arg2 - The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHIDIST.js
	 */
	ApiWorksheetFunction.prototype.CHIDIST = function (arg1, arg2) {
		return this.private_calculateFunction("CHIDIST", arguments);
	};
	/**
	 * Returns the inverse of the right-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
	 * @param {number} arg2 - The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHIINV.js
	 */
	ApiWorksheetFunction.prototype.CHIINV = function (arg1, arg2) {
		return this.private_calculateFunction("CHIINV", arguments);
	};
	/**
	 * Returns the left-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which the distribution will be evaluated, a nonnegative number.
	 * @param {number} arg2 - The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {boolean} arg3 - A logical value that determines the form of the function. If this argument is equal to <b>true</b>,
	 * the cumulative distribution function is returned; if  it is equal to <b>false</b>, the probability density function is returned.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHISQ_DIST.js
	 */
	ApiWorksheetFunction.prototype.CHISQ_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CHISQ.DIST", arguments);
	};
	/**
	 * Returns the right-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which the distribution will be evaluated, a nonnegative number.
	 * @param {number} arg2 - The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHISQ_DIST_RT.js
	 */
	ApiWorksheetFunction.prototype.CHISQ_DIST_RT = function (arg1, arg2) {
		return this.private_calculateFunction("CHISQ.DIST.RT", arguments);
	};
	/**
	 * Returns the inverse of the left-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
	 * @param {number} arg2- The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHISQ_INV.js
	 */
	ApiWorksheetFunction.prototype.CHISQ_INV = function (arg1, arg2) {
		return this.private_calculateFunction("CHISQ.INV", arguments);
	};
	/**
	 * Returns the inverse of the right-tailed probability of the chi-squared distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
	 * @param {number} arg2 - The number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHISQ_INV_RT.js
	 */
	ApiWorksheetFunction.prototype.CHISQ_INV_RT = function (arg1, arg2) {
		return this.private_calculateFunction("CHISQ.INV.RT", arguments);
	};

	//todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.CHITEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("CHITEST", arguments);
	// };


	/**
	 * Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 - The range of data that contains observations to test against expected values.
	 * @param {any} arg2 - The range of data that contains the ratio of the product of row totals and column totals to the grand total.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHITEST.js
	 */
	ApiWorksheetFunction.prototype.CHITEST = function (arg1, arg2) {
		return this.private_calculateFunction("CHITEST", arguments);
	};
	// todo need array
	// /**
	//  * Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the range of data that contains observations to test against expected values.
	//  * @param {any} arg2 Is the range of data that contains the ratio of the product of row totals and column totals to the grand total.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.CHISQ_TEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("CHISQ.TEST", arguments);
	// };
	/**
	 * Returns the confidence interval for a population mean, using a normal distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The significance level used to compute the confidence level, a number greater than 0 and less than 1.
	 * @param {number} arg2 - The population standard deviation for the data range and is assumed to be known. This value must be greater than 0.
	 * @param {number} arg3 - The sample size.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CONFIDENCE.js
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE", arguments);
	};
	/**
	 * Returns the confidence interval for a population mean, using a normal distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The significance level used to compute the confidence level, a number greater than 0 and less than 1.
	 * @param {number} arg2 - The population standard deviation for the data range and is assumed to be known. This value must be greater than 0.
	 * @param {number} arg3 - The sample size.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CONFIDENCE_NORM.js
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE_NORM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE.NORM", arguments);
	};
	/**
	 * Returns the confidence interval for a population mean, using a Student's t distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The significance level used to compute the confidence level, a number greater than 0 and less than 1.
	 * @param {number} arg2 - The population standard deviation for the data range and is assumed to be known. This value must be greater than 0.
	 * @param {number} arg3 - The sample size.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CONFIDENCE_T.js
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE_T = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE.T", arguments);
	};
	// todo need array
	// /**
	//  * Returns the correlation coefficient between two data sets.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is a cell range of values. The values should be numbers, names, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is a second cell range of values. The values should be numbers, names, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.CORREL = function (arg1, arg2) {
	// 	return this.private_calculateFunction("CORREL", arguments);
	// };
	/**
	 * Counts a number of cells in a range that contains numbers ignoring empty cells or those contaning text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string | array | ApiRange} args - Up to 255 items, cell references, or ranges to count numbers. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUNT.js
	 */
	ApiWorksheetFunction.prototype.COUNT = function () {
		return this.private_calculateFunction("COUNT", arguments);
	};
	/**
	 * Counts a number of cells in a range that are not empty.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string | array | ApiRange} args - Up to 255 items, cell references, or ranges to count values. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUNTA.js
	 */
	ApiWorksheetFunction.prototype.COUNTA = function () {
		return this.private_calculateFunction("COUNTA", arguments);
	};
	/**
	 * Counts a number of empty cells in a specified range of cells.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The range to count the empty cells.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUNTBLANK.js
	 */
	ApiWorksheetFunction.prototype.COUNTBLANK = function (arg1) {
		return this.private_calculateFunction("COUNTBLANK", arguments);
	};
	/**
	 * Counts a number of cells within a range that meet the given condition.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The range of cells to count nonblank cells.
	 * @param {number | string} arg2 - The condition in the form of a number, expression, or text that defines which cells will be counted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUNTIF.js
	 */
	ApiWorksheetFunction.prototype.COUNTIF = function (arg1, arg2) {
		return this.private_calculateFunction("COUNTIF", arguments);
	};
	/**
	 * Counts a number of cells specified by a given set of conditions or criteria.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The first range of cells to count nonblank cells.
	 * @param {number | string} arg2 - The first condition in the form of a number, expression, or text that defines which cells will be counted.
	 * @param {ApiRange} arg3 - Up to 127 additional ranges of cells to count nonblank cells. This argument is optional.
	 * @param {number | string} arg4 - Up to 127 additional conditions in the form of a number, expression, or text that define which cells will be counted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUNTIFS.js
	 */
	ApiWorksheetFunction.prototype.COUNTIFS = function () {
		return this.private_calculateFunction("COUNTIFS", arguments);
	};

	// todo need array
	// /**
	//  * Returns covariance, the average of the products of deviations for each data point pair in two data sets.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is the second cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVAR = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVAR", arguments);
	// };

	// todo need array
	// /**
	//  * Returns population covariance, the average of the products of deviations for each data point pair in two data sets.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is the second cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVARIANCE_P = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVARIANCE.P", arguments);
	// };

	// todo need array
	// /**
	//  * Returns sample covariance, the average of the products of deviations for each data point pair in two data sets.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is the second cell range of integers and must be numbers, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVARIANCE_S = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVARIANCE.S", arguments);
	// };
	/**
	 * Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of Bernoulli trials.
	 * @param {number} arg2 - The probability of success on each trial, a number between 0 and 1 inclusive.
	 * @param {number} arg3 - The criterion value, a number between 0 and 1 inclusive.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CRITBINOM.js
	 */
	ApiWorksheetFunction.prototype.CRITBINOM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CRITBINOM", arguments);
	};
	/**
	 * Returns the sum of squares of deviations of data points from their sample mean.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} args - Up to 255 numerical values for which to find the sum of squares of deviations. The first argument is required, subsequent arguments are optional. 
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DEVSQ.js
	 */
	ApiWorksheetFunction.prototype.DEVSQ = function () {
		return this.private_calculateFunction("DEVSQ", arguments);
	};
	/**
	 * Returns the exponential distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value of the x function, a nonnegative number.
	 * @param {number} arg2 - The lambda parameter value, a positive number.
	 * @param {boolean} arg3 - A logical value that determines the function form. If this parameter is <b>true</b>,
	 * the function will return the cumulative distribution function, if it is <b>false</b>, it will return the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EXPON_DIST.js
	 */
	ApiWorksheetFunction.prototype.EXPON_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("EXPON.DIST", arguments);
	};
	/**
	 * Returns the exponential distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value of the x function, a nonnegative number.
	 * @param {number} arg2 - The lambda parameter value, a positive number.
	 * @param {boolean} arg3 - A logical value that determines the function form. If this parameter is <b>true</b>,
	 * the function will return the cumulative distribution function, if it is <b>false</b>, it will return the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EXPONDIST.js
	 */
	ApiWorksheetFunction.prototype.EXPONDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("EXPONDIST", arguments);
	};
	/**
	 * Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a nonnegative number.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {boolean} arg4 - A logical value that determines the function form. If this parameter is <b>true</b>,
	 * the function will return the cumulative distribution function, if it is <b>false</b>, it will return the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/F_DIST.js
	 */
	ApiWorksheetFunction.prototype.F_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("F.DIST", arguments);
	};
	/**
	 * Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a nonnegative number.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FDIST.js
	 */
	ApiWorksheetFunction.prototype.FDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FDIST", arguments);
	};
	/**
	 * Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a nonnegative number.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/F_DIST_RT.js
	 */
	ApiWorksheetFunction.prototype.F_DIST_RT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.DIST.RT", arguments);
	};
	/**
	 * Returns the inverse of the (left-tailed) F probability distribution: if p = F.DIST(x,...), then F.INV(p,...) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/F_INV.js
	 */
	ApiWorksheetFunction.prototype.F_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.INV", arguments);
	};
	/**
	 * Returns the inverse of the (right-tailed) F probability distribution: if p = FDIST(x,...), then FINV(p,...) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FINV.js
	 */
	ApiWorksheetFunction.prototype.FINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FINV", arguments);
	};
	/**
	 * Returns the inverse of the (right-tailed) F probability distribution: if p = F.DIST.RT(x,...), then F.INV.RT(p,...) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - The numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @param {number} arg3 - The denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/F_INV_RT.js
	 */
	ApiWorksheetFunction.prototype.F_INV_RT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.INV.RT", arguments);
	};
	/**
	 * Returns the Fisher transformation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for the transformation, a number between -1 and 1, excluding -1 and 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FISHER.js
	 */
	ApiWorksheetFunction.prototype.FISHER = function (arg1) {
		return this.private_calculateFunction("FISHER", arguments);
	};
	/**
	 * Returns the inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value to perform the inverse of the transformation.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FISHERINV.js
	 */
	ApiWorksheetFunction.prototype.FISHERINV = function (arg1) {
		return this.private_calculateFunction("FISHERINV", arguments);
	};
	//todo need array
	// /**
	//  * Calculates, or predicts, a future value along a linear trend by using existing values.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {number} arg1 Is the data point for which you want to predict a value and must be a numeric value.
	//  * @param {any} arg2 Is the dependent array or range of numeric data.
	//  * @param {any} arg3 Is the independent array or range of numeric data. The variance of Known_x's must not be zero.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FORECAST = function (arg1, arg2, arg3) {
	// 	return this.private_calculateFunction("FORECAST", arguments);
	// };
	/**
	 * A numeric value that specifies which function should be used to aggregate identical time values in the timeline data range.
	 * <b>1</b> (or omitted) - AVERAGE.
	 * <b>2</b> - COUNT.
	 * <b>3</b> - COUNTA.
	 * <b>4</b> - MAX.
	 * <b>5</b> - MEDIAN.
	 * <b>6</b> - MIN.
	 * <b>7</b> - SUM.
	 * @typedef {(1 | 2 | 3 | 4 | 5 | 6 | 7)} Aggregation
	 * @see office-js-api/Examples/Enumerations/Aggregation.js
	 */
	
	/**
	 * Сalculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A date for which a new value will be predicted. Must be after the last date in the timeline.
	 * @param {ApiRange | number[]} arg2 - A range or an array of numeric data that determines the historical values for which a new point will be predicted.
	 * @param {ApiRange} arg3 - A range of date/time values that correspond to the historical values.
	 * The timeline range must be of the same size as the second argument. Date/time values must have a constant step between them and can't be zero.
	 * @param {?number} arg4 - An optional numeric value that specifies the length of the seasonal pattern. The default value of 1 indicates seasonality is detected automatically.
	 * The 0 value means no seasonality.
	 * @param {?number} arg5 - An optional numeric value to handle missing values. The default value of 1 replaces missing values by interpolation, and 0 replaces them with zeros.
	 * @param {Aggregation} arg6 - An optional numeric value to aggregate multiple values with the same time stamp.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FORECAST_ETS.js
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("FORECAST.ETS", arguments);
	};
	/**
	 * Returns a confidence interval for the forecast value at the specified target date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A date for which a new value will be predicted. Must be after the last date in the timeline.
	 * @param {ApiRange | number[]} arg2 - A range or an array of numeric data that determines the historical values for which a new point will be predicted.
	 * @param {ApiRange} arg3 - A range of date/time values that correspond to the historical values.
	 * The timeline range must be of the same size as the second argument. Date/time values must have a constant step between them and can't be zero.
	 * @param {?number} arg4 - A number between 0 and 1 that shows the confidence level for the calculated confidence interval. The default value is .95.
	 * @param {?number} arg5 - An optional numeric value that specifies the length of the seasonal pattern. The default value of 1 indicates seasonality is detected automatically.
	 * The 0 value means no seasonality.
	 * @param {?number} arg6 - An optional numeric value to handle missing values. The default value of 1 replaces missing values by interpolation, and 0 replaces them with zeros.
	 * @param {Aggregation} arg7 - An optional numeric value to aggregate multiple values with the same time stamp.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FORECAST_ETS_CONFINT.js
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS_CONFINT = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("FORECAST.ETS.CONFINT", arguments);
	};
	/**
	 * Returns the length of the repetitive pattern an application detects for the specified time series.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - A range or an array of numeric data that determines the historical values for which a new point will be predicted.
	 * @param {ApiRange} arg2 - A range of date/time values that correspond to the historical values.
	 * The timeline range must be of the same size as the second argument. Date/time values must have a constant step between them and can't be zero.
	 * @param {?number} arg3 - An optional numeric value to handle missing values. The default value of 1 replaces missing values by interpolation, and 0 replaces them with zeros.
	 * @param {Aggregation} arg4 - An optional numeric value to aggregate multiple values with the same time stamp.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FORECAST_ETS_SEASONALITY.js
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS_SEASONALITY = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("FORECAST.ETS.SEASONALITY", arguments);
	};

	/**
	 * A numeric value between 1 and 8 that specifies which statistic will be returned.
	 * <b>1</b> - Alpha parameter of ETS algorithm - the base value parameter.
	 * <b>2</b> - Beta parameter of ETS algorithm - the trend value parameter.
	 * <b>3</b> - Gamma parameter of ETS algorithm - the seasonality value parameter.
	 * <b>4</b> - MASE (mean absolute scaled error) metric - a measure of the accuracy of forecasts.
	 * <b>5</b> - SMAPE (symmetric mean absolute percentage error) metric - a measure of the accuracy based on percentage errors.
	 * <b>6</b> - MAE (mean absolute error) metric - a measure of the accuracy of forecasts.
	 * <b>7</b> - RMSE (root mean squared error) metric - a measure of the differences between predicted and observed values.
	 * <b>8</b> - Step size detected in the timeline.
	 * @typedef {(1 | 2 | 3 | 4 | 5 | 6 | 7 | 8)} StatisticType
	 * @see office-js-api/Examples/Enumerations/StatisticType.js
	 */

	/**
	 * Returns the requested statistic for the forecast.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - A range or an array of numeric data that determines the historical values for which a new point will be predicted.
	 * @param {ApiRange} arg2 - A range of date/time values that correspond to the historical values.
	 * The timeline range must be of the same size as the second argument. Date/time values must have a constant step between them and can't be zero.
	 * @param {StatisticType} arg3 - A number between 1 and 8, indicating which statistic will be returned for the calculated forecast.
	 * @param {?number} arg4 - An optional numeric value that specifies the length of the seasonal pattern. The default value of 1 indicates seasonality is detected automatically.
	 * The 0 value means no seasonality.
	 * @param {?number} arg5 - An optional numeric value to handle missing values. The default value of 1 replaces missing values by interpolation, and 0 replaces them with zeros.
	 * @param {Aggregation} arg6 - An optional numeric value to aggregate multiple values with the same time stamp.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FORECAST_ETS_STAT.js
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS_STAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("FORECAST.ETS.STAT", arguments);
	};
	//todo need array
	// /**
	//  * Calculates, or predicts, a future value along a linear trend by using existing values.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {number} arg1 Is the data point for which you want to predict a value and must be a numeric value.
	//  * @param {any} arg2 Is the dependent array or range of numeric data.
	//  * @param {any} arg3 Is the independent array or range of numeric data. The variance of Known_x's must not be zero.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FORECAST_LINEAR = function (arg1, arg2, arg3) {
	// 	return this.private_calculateFunction("FORECAST.LINEAR", arguments);
	// };
	/**
	 * Calculates how often values occur within a range of values and then returns the first value of the returned vertical array of numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - An array of values or the selected range for which the frequencies will be counted (blanks and text are ignored).
	 * @param {ApiRange | number[]} arg2 - An array of intervals or the selected range into which the values in the first array will be grouped.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FREQUENCY.js
	 */
	ApiWorksheetFunction.prototype.FREQUENCY = function (arg1, arg2) {
		return this.private_calculateFunction("FREQUENCY", arguments);
	};
	// //todo need array
	// /**
	//  * Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored).
	//  * @param {any} arg2 Is the second array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored).
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FTEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("FTEST", arguments);
	// };
	// //todo need array
	// /**
	//  * Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored).
	//  * @param {any} arg2 Is the second array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored).
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.F_TEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("F.TEST", arguments);
	// };
	/**
	 * Returns the gamma function value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the gamma function will be calculated.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMA.js
	 */
	ApiWorksheetFunction.prototype.GAMMA = function (arg1) {
		return this.private_calculateFunction("GAMMA", arguments);
	};
	/**
	 * Returns the gamma distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which the distribution will be calculated, a nonnegative number.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number. If this parameter is equal to 1, the function returns the standard gamma distribution.
	 * @param {boolean} arg4 - A logical value (<b>true</b>> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function. If it is <b>false</b>, the function returns the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMA_DIST.js
	 */
	ApiWorksheetFunction.prototype.GAMMA_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GAMMA.DIST", arguments);
	};
	/**
	 * Returns the gamma distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which the distribution will be calculated, a nonnegative number.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number. If this parameter is equal to 1, the function returns the standard gamma distribution.
	 * @param {boolean} arg4 - A logical value (<b>true</b>> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function. If it is <b>false</b>, the function returns the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMADIST.js
	 */
	ApiWorksheetFunction.prototype.GAMMADIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GAMMADIST", arguments);
	};
	/**
	 * Returns the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The probability associated with the gamma distribution, a number between 0 and 1, inclusive.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number. If this parameter is equal to 1, the function returns the standard gamma distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMA_INV.js
	 */
	ApiWorksheetFunction.prototype.GAMMA_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("GAMMA.INV", arguments);
	};
	/**
	 * Returns the inverse of the gamma cumulative distribution: if p = GAMMADIST(x,...), then GAMMAINV(p,...) = x.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The probability associated with the gamma distribution, a number between 0 and 1, inclusive.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number. If this parameter is equal to 1, the function returns the standard gamma distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMAINV.js
	 */
	ApiWorksheetFunction.prototype.GAMMAINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("GAMMAINV", arguments);
	};
	/**
	 * Returns the natural logarithm of the gamma function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the natural logarithm of the gamma function will be calculated, a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMALN.js
	 */
	ApiWorksheetFunction.prototype.GAMMALN = function (arg1) {
		return this.private_calculateFunction("GAMMALN", arguments);
	};
	/**
	 * Returns the natural logarithm of the gamma function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the natural logarithm of the gamma function will be calculated, a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAMMALN_PRECISE.js
	 */
	ApiWorksheetFunction.prototype.GAMMALN_PRECISE = function (arg1) {
		return this.private_calculateFunction("GAMMALN.PRECISE", arguments);
	};
	/**
	 * Calculates the probability that a member of a standard normal population will fall between the mean and arg1 standard deviations from the mean.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the distribution will be calculated.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GAUSS.js
	 */
	ApiWorksheetFunction.prototype.GAUSS = function (arg1) {
		return this.private_calculateFunction("GAUSS", arguments);
	};
	/**
	 * Returns the geometric mean of an array or range of positive numeric data.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} args - Up to 255 numeric values for which the geometric mean will be calculated.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GEOMEAN.js
	 */
	ApiWorksheetFunction.prototype.GEOMEAN = function () {
		return this.private_calculateFunction("GEOMEAN", arguments);
	};
	/**
	 * Calculates predicted exponential growth by using existing data.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - The set of y-values from the <em>y = b*m^x</em> equation, an array or range of positive numbers.
	 * @param {?ApiRange | number[]} arg2 - An optional set of x-values from the <em>y = b*m^x</em> equation, an array or range of positive numbers that has the same size as the set of y-values.
	 * @param {?ApiRange | number[]} arg3 - New x-values for which the function will return the corresponding y-values.
	 * @param {?boolean} arg4 - A logical value: the constant <em>b</em> is calculated normally if this parameter is set to <b>true</b>,
	 * and <em>b</em> is set equal to 1 if the parameter is <b>false</b> or omitted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GROWTH.js
	 */
	ApiWorksheetFunction.prototype.GROWTH = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GROWTH", arguments);
	};
	/**
	 * Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} args - Up to 255 numeric values for which the harmonic mean will be calculated.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HARMEAN.js
	 */
	ApiWorksheetFunction.prototype.HARMEAN = function () {
		return this.private_calculateFunction("HARMEAN", arguments);
	};
	/**
	 * Returns the hypergeometric distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of successes in the sample.
	 * @param {number} arg2 - The size of the sample.
	 * @param {number} arg3 - The number of successes in the population.
	 * @param {number} arg4 - The population size.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HYPGEOMDIST.js
	 */
	ApiWorksheetFunction.prototype.HYPGEOMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("HYPGEOMDIST", arguments);
	};
	/**
	 * Returns the hypergeometric distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of successes in the sample.
	 * @param {number} arg2 - The size of the sample.
	 * @param {number} arg3 - The number of successes in the population.
	 * @param {number} arg4 - The population size.
	 * @param {boolean} arg5 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function. If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HYPGEOM_DIST.js
	 */
	ApiWorksheetFunction.prototype.HYPGEOM_DIST = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("HYPGEOM.DIST", arguments);
	};
	// todo need array
	// /**
	//  * Calculates the point at which a line will intersect the y-axis by using a best-fit regression line plotted through the known x-values and y-values.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the dependent set of observations or data and can be numbers or names, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is the independent set of observations or data and can be numbers or names, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.INTERCEPT = function (arg1, arg2) {
	// 	return this.private_calculateFunction("INTERCEPT", arguments);
	// };
	/**
	 * Returns the kurtosis of a data set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} args - Up to 255 numeric values for which the kurtosis will be calculated.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/KURT.js
	 */
	ApiWorksheetFunction.prototype.KURT = function () {
		return this.private_calculateFunction("KURT", arguments);
	};
	/**
	 * Returns the k-th largest value in a data set. For example, the fifth largest number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - The array or range of data for which the k-th largest value will be determined.
	 * @param {number} arg2 - The position (from the largest) in the array or cell range of data to return.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LARGE.js
	 */
	ApiWorksheetFunction.prototype.LARGE = function (arg1, arg2) {
		return this.private_calculateFunction("LARGE", arguments);
	};
	/**
	 * Returns statistics that describe a linear trend matching known data points, by fitting a straight line using the least squares method.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The set of y-values from the <em>y = mx + b</em> equation.
	 * @param {?ApiRange} arg2 - An optional set of x-values from the <em>y = mx + b</em> equation.
	 * @param {?boolean} arg3 - A logical value: the constant <em>b</em> is calculated normally if this parameter is set to <b>true</b> or omitted,
	 * and <em>b</em> is set equal to 0 if the parameter is <b>false</b>.
	 * @param {?boolean} arg4 - A logical value: return additional regression statistics if this parameter is set to <b>true</b>,
	 * and return m-coefficients and the constant <em>b</em> if the parameter is <b>false</b> or omitted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LINEST.js
	 */
	ApiWorksheetFunction.prototype.LINEST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LINEST", arguments);
	};
	/**
	 * Returns statistics that describe an exponential curve matching known data points.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 - The set of y-values from the <em>y = b*m^x</em> equation.
	 * @param {?ApiRange} arg2 - An optional set of x-values from the <em>y = b*m^x</em> equation.
	 * @param {?boolean} arg3 - A logical value: the constant <em>b</em> is calculated normally if this parameter is set to <b>true</b> or omitted,
	 * and <em>b</em> is set equal to 1 if the parameter is <b>false</b>.
	 * @param {?boolean} arg4 - A logical value: return additional regression statistics if this parameter is set to <b>true</b>,
	 * and return m-coefficients and the constant <em>b</em> if the parameter is <b>false</b> or omitted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOGEST.js
	 */
	ApiWorksheetFunction.prototype.LOGEST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LOGEST", arguments);
	};
	/**
	 * Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with the specified parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the lognormal distribution, a number between 0 and 1, inclusive.
	 * @param {number} arg2 - The mean of ln(x).
	 * @param {number} arg3 - The standard deviation of ln(x), a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOGINV.js
	 */
	ApiWorksheetFunction.prototype.LOGINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGINV", arguments);
	};
	/**
	 * Returns the lognormal distribution of x, where ln(x) is normally distributed with the specified parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a positive number.
	 * @param {number} arg2 - The mean of ln(x).
	 * @param {number} arg3 - The standard deviation of ln(x), a positive number.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOGNORM_DIST.js
	 */
	ApiWorksheetFunction.prototype.LOGNORM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LOGNORM.DIST", arguments);
	};
	/**
	 * Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with the specified parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability associated with the lognormal distribution, a number between 0 and 1, inclusive.
	 * @param {number} arg2 - The mean of ln(x).
	 * @param {number} arg3 - The standard deviation of ln(x), a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOGNORM_INV.js
	 */
	ApiWorksheetFunction.prototype.LOGNORM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGNORM.INV", arguments);
	};
	/**
	 * Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with the specified parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a positive number.
	 * @param {number} arg2 - The mean of ln(x).
	 * @param {number} arg3 - The standard deviation of ln(x), a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOGNORMDIST.js
	 */
	ApiWorksheetFunction.prototype.LOGNORMDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGNORMDIST", arguments);
	};
	/**
	 * Returns the largest value in a set of values. Ignores logical values and text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | array | ApiRange} args - Up to 255 numeric values for which the largest number will be returned.
	 * The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MAX.js
	 */
	ApiWorksheetFunction.prototype.MAX = function () {
		return this.private_calculateFunction("MAX", arguments);
	};
	/**
	 * Returns the largest value in a set of values. Does not ignore logical values and text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | array | ApiRange} args - Up to 255 values (number, text, logical value) for which the largest value will be returned.
	 * The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MAXA.js
	 */
	ApiWorksheetFunction.prototype.MAXA = function () {
		return this.private_calculateFunction("MAXA", arguments);
	};
	/**
	 * Returns the median, or the number in the middle of the set of given numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | array | ApiRange} args - Up to 255 numeric values for which the median will be calculated.
	 * The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MEDIAN.js
	 */
	ApiWorksheetFunction.prototype.MEDIAN = function () {
		return this.private_calculateFunction("MEDIAN", arguments);
	};
	/**
	 * Returns the smallest number in a set of values. Ignores logical values and text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | array | ApiRange} args - Up to 255 numeric values for which the smallest number will be returned.
	 * The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MIN.js
	 */
	ApiWorksheetFunction.prototype.MIN = function () {
		return this.private_calculateFunction("MIN", arguments);
	};
	/**
	 * Returns the smallest value in a set of values. Does not ignore logical values and text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | array | ApiRange} args - Up to 255 values (number, text, logical value) for which the smallest value will be returned.
	 * The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MINA.js
	 */
	ApiWorksheetFunction.prototype.MINA = function () {
		return this.private_calculateFunction("MINA", arguments);
	};
	// todo need array
	// /**
	//  * Returns the most frequently occurring, or repetitive, value in an array or range of data.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE = function () {
	// 	return this.private_calculateFunction("MODE", arguments);
	// };
	// todo need array
	// /**
	//  * Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data. For a horizontal array, use =TRANSPOSE(MODE.MULT(number1,number2,...)).
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE_MULT = function () {
	// 	return this.private_calculateFunction("MODE.MULT", arguments);
	// };
	// todo need array
	// /**
	//  * Returns the most frequently occurring, or repetitive, value in an array or range of data.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE_SNGL = function () {
	// 	return this.private_calculateFunction("MODE.SNGL", arguments);
	// };
	/**
	 * Returns the negative binomial distribution, the probability that there will be the specified number of failures before the last success, with the specified probability of a success.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of failures.
	 * @param {number} arg2 - The threshold number of successes.
	 * @param {number} arg3 - The probability of a success; a number between 0 and 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NEGBINOMDIST.js
	 */
	ApiWorksheetFunction.prototype.NEGBINOMDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NEGBINOMDIST", arguments);
	};
	/**
	 * Returns the negative binomial distribution, the probability that there will be the specified number of failures before the last success, with the specified probability of a success.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of failures.
	 * @param {number} arg2 - The threshold number of successes.
	 * @param {number} arg3 - The probability of a success; a number between 0 and 1.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NEGBINOM_DIST.js
	 */
	ApiWorksheetFunction.prototype.NEGBINOM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NEGBINOM.DIST", arguments);
	};
	/**
	 * Returns the normal cumulative distribution for the specified mean and standard deviation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the distribution will be returned.
	 * @param {number} arg2 - The arithmetic mean of the distribution.
	 * @param {number} arg3 - The standard deviation of the distribution, a positive number.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORMDIST.js
	 */
	ApiWorksheetFunction.prototype.NORMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NORMDIST", arguments);
	};
	/**
	 * Returns the normal distribution for the specified mean and standard deviation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the distribution will be returned.
	 * @param {number} arg2 - The arithmetic mean of the distribution.
	 * @param {number} arg3 - The standard deviation of the distribution, a positive number.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORM_DIST.js
	 */
	ApiWorksheetFunction.prototype.NORM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NORM.DIST", arguments);
	};
	/**
	 * Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - The arithmetic mean of the distribution.
	 * @param {number} arg3 - The standard deviation of the distribution, a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORMINV.js
	 */
	ApiWorksheetFunction.prototype.NORMINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NORMINV", arguments);
	};
	/**
	 * Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - The arithmetic mean of the distribution.
	 * @param {number} arg3 - The standard deviation of the distribution, a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORM_INV.js
	 */
	ApiWorksheetFunction.prototype.NORM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NORM.INV", arguments);
	};
	/**
	 * Returns the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the distribution will be returned.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORMSDIST.js
	 */
	ApiWorksheetFunction.prototype.NORMSDIST = function (arg1) {
		return this.private_calculateFunction("NORMSDIST", arguments);
	};
	/**
	 * Returns the standard normal distribution (has a mean of zero and a standard deviation of one).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value for which the distribution will be returned.
	 * @param {boolean} arg2 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORM_S_DIST.js
	 */
	ApiWorksheetFunction.prototype.NORM_S_DIST = function (arg1, arg2) {
		return this.private_calculateFunction("NORM.S.DIST", arguments);
	};
	/**
	 * Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORMSINV.js
	 */
	ApiWorksheetFunction.prototype.NORMSINV = function (arg1) {
		return this.private_calculateFunction("NORMSINV", arguments);
	};
	/**
	 * Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NORM_S_INV.js
	 */
	ApiWorksheetFunction.prototype.NORM_S_INV = function (arg1) {
		return this.private_calculateFunction("NORM.S.INV", arguments);
	};
	// todo need array
	// /**
	//  * Returns the Pearson product moment correlation coefficient, r.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is a set of independent values.
	//  * @param {any} arg2 Is a set of dependent values.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.PEARSON = function (arg1, arg2) {
	// 	return this.private_calculateFunction("PEARSON", arguments);
	// };

	/**
	 * Returns the k-th percentile of values in a range.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data that defines relative standing.
	 * @param {number} arg2 - The percentile value that is equal to 0 but less than or equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTILE.js
	 */
	ApiWorksheetFunction.prototype.PERCENTILE = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE", arguments);
	};
	/**
	 * Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data that defines relative standing.
	 * @param {number} arg2 - The percentile value that is greater than 0 but less than 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTILE_EXC.js
	 */
	ApiWorksheetFunction.prototype.PERCENTILE_EXC = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE.EXC", arguments);
	};
	/**
	 * Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data that defines relative standing.
	 * @param {number} arg2 - The percentile value that is equal to 0 but less than or equal to 1.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTILE_INC.js
	 */
	ApiWorksheetFunction.prototype.PERCENTILE_INC = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE.INC", arguments);
	};
	/**
	 * Returns the rank of a value in a data set as a percentage of the data set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data with numeric values that defines relative standing.
	 * @param {number} arg2 - The value for which the rank will be returned.
	 * @param {?number} arg3 - An optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTRANK.js
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK", arguments);
	};
	/**
	 * Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data with numeric values that defines relative standing.
	 * @param {number} arg2 - The value for which the rank will be returned.
	 * @param {?number} arg3 - An optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTRANK_EXC.js
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK_EXC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK.EXC", arguments);
	};
	/**
	 * Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data with numeric values that defines relative standing.
	 * @param {number} arg2 - The value for which the rank will be returned.
	 * @param {?number} arg3 - An optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERCENTRANK_INC.js
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK_INC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK.INC", arguments);
	};
	/**
	 * Returns the number of permutations for a given number of objects that can be selected from the total objects.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The total number of objects.
	 * @param {number} arg2 - The number of objects in each permutation.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERMUT.js
	 */
	ApiWorksheetFunction.prototype.PERMUT = function (arg1, arg2) {
		return this.private_calculateFunction("PERMUT", arguments);
	};
	/**
	 * Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The total number of objects.
	 * @param {number} arg2 - The number of objects in each permutation.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PERMUTATIONA.js
	 */
	ApiWorksheetFunction.prototype.PERMUTATIONA = function (arg1, arg2) {
		return this.private_calculateFunction("PERMUTATIONA", arguments);
	};
	/**
	 * Returns the value of the density function for a standard normal distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number for which the density of the standard normal distribution will be returned.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PHI.js
	 */
	ApiWorksheetFunction.prototype.PHI = function (arg1) {
		return this.private_calculateFunction("PHI", arguments);
	};
	/**
	 * Returns the Poisson distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of events.
	 * @param {number} arg2 - The expected numeric value, a positive number.
	 * @param {boolean} arg3 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative Poisson probability.
	 * If it is <b>false</b>, the function returns the Poisson probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/POISSON.js
	 */
	ApiWorksheetFunction.prototype.POISSON = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("POISSON", arguments);
	};
	/**
	 * Returns the Poisson distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number of events.
	 * @param {number} arg2 - The expected numeric value, a positive number.
	 * @param {boolean} arg3 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative Poisson probability.
	 * If it is <b>false</b>, the function returns the Poisson probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/POISSON_DIST.js
	 */
	ApiWorksheetFunction.prototype.POISSON_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("POISSON.DIST", arguments);
	};
	// todo need array
	// /**
	//  * Returns the probability that values in a range are between two limits or equal to a lower limit.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the range of numeric values of x with which there are associated probabilities.
	//  * @param {any} arg2 Is the set of probabilities associated with values in X_range, values between 0 and 1 and excluding 0.
	//  * @param {number} arg3 Is the lower bound on the value for which you want a probability.
	//  * @param {?number} arg4 Is the optional upper bound on the value. If omitted, PROB returns the probability that X_range values are equal to Lower_limit.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.PROB = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("PROB", arguments);
	// };
	/**
	 * Returns the quartile of a data set.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or cell range of numeric values for which the quartile value will be returned.
	 * @param {number} arg2 - The quartile value to return: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/QUARTILE.js
	 */
	ApiWorksheetFunction.prototype.QUARTILE = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE", arguments);
	};
	/**
	 * Returns the quartile of a data set, based on percentile values from 0..1, exclusive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or cell range of numeric values for which the quartile value will be returned.
	 * @param {number} arg2 - The quartile value to return: 1st quartile = 1; median value = 2; 3rd quartile = 3.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/QUARTILE_EXC.js
	 */
	ApiWorksheetFunction.prototype.QUARTILE_EXC = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE.EXC", arguments);
	};
	/**
	 * Returns the quartile of a data set, based on percentile values from 0..1, inclusive.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or cell range of numeric values for which the quartile value will be returned.
	 * @param {number} arg2 - The quartile value to return: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/QUARTILE_INC.js
	 */
	ApiWorksheetFunction.prototype.QUARTILE_INC = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE.INC", arguments);
	};
	/**
	 * Returns the rank of a number in a list of numbers: its size relative to other values in the list.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number for which the rank will be returned.
	 * @param {number[] | ApiRange} arg2 - An array of numbers or a reference to a list of numbers. Nonnumeric values are ignored.
	 * @param {?boolean} arg3 - The numeric value that specifyes how to order the numbers. If it is 0 or omitted, the rank in the list will be sorted in descending order.
	 * Any other numeric value means that the rank in the list will be sorted in ascending order.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RANK.js
	 */
	ApiWorksheetFunction.prototype.RANK = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK", arguments);
	};
	/**
	 * Returns the rank of a number in a list of numbers: its size relative to other values in the list. If more than one value has the same rank, the average rank is returned.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number for which the rank will be returned.
	 * @param {number[] | ApiRange} arg2 - An array of numbers or a reference to a list of numbers. Nonnumeric values are ignored.
	 * @param {?boolean} arg3 - The numeric value that specifyes how to order the numbers. If it is 0 or omitted, the rank in the list will be sorted in descending order.
	 * Any other numeric value means that the rank in the list will be sorted in ascending order.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RANK_AVG.js
	 */
	ApiWorksheetFunction.prototype.RANK_AVG = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK.AVG", arguments);
	};
	/**
	 * Returns the rank of a number in a list of numbers: its size relative to other values in the list. If more than one value has the same rank, the top rank of that set of values is returned.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The number for which the rank will be returned.
	 * @param {number[] | ApiRange} arg2 - An array of numbers or a reference to a list of numbers. Nonnumeric values are ignored.
	 * @param {?boolean} arg3 - The numeric value that specifyes how to order the numbers. If it is 0 or omitted, the rank in the list will be sorted in descending order.
	 * Any other numeric value means that the rank in the list will be sorted in ascending order.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RANK_EQ.js
	 */
	ApiWorksheetFunction.prototype.RANK_EQ = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK.EQ", arguments);
	};
	// todo need array
	// /**
	//  * Returns the square of the Pearson product moment correlation coefficient through the given data points.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is an array or range of data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is an array or range of data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.RSQ = function (arg1, arg2) {
	// 	return this.private_calculateFunction("RSQ", arguments);
	// };

	/**
	 * Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the skewness of a distribution will be returned. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SKEW.js
	 */
	ApiWorksheetFunction.prototype.SKEW = function () {
		return this.private_calculateFunction("SKEW", arguments);
	};
	/**
	 * Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the skewness of a distribution will be returned. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SKEW_P.js
	 */
	ApiWorksheetFunction.prototype.SKEW_P = function () {
		return this.private_calculateFunction("SKEW.P", arguments);
	};
	// todo need array
	// /**
	//  * Returns the slope of the linear regression line through the given data points.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is an array or cell range of numeric dependent data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is the set of independent data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SLOPE = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SLOPE", arguments);
	// };
	/**
	 * Returns the k-th smallest value in a data set. For example, the fifth smallest number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - An array or range of numerical data for which the k-th smallest value will be determined.
	 * @param {number} arg2 - The position (from the smallest) in the array or range of the value to return.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SMALL.js
	 */
	ApiWorksheetFunction.prototype.SMALL = function (arg1, arg2) {
		return this.private_calculateFunction("SMALL", arguments);
	};
	/**
	 * Returns a normalised value from a distribution characterised by a mean and standard deviation.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value to normalize.
	 * @param {number} arg2 - The arithmetic mean of the distribution.
	 * @param {number} arg3 - The standard deviation of the distribution, a positive number.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STANDARDIZE.js
	 */
	ApiWorksheetFunction.prototype.STANDARDIZE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("STANDARDIZE", arguments);
	};
	/**
	 * Estimates standard deviation based on a sample (ignores logical values and text in the sample).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEV.js
	 */
	ApiWorksheetFunction.prototype.STDEV = function () {
		return this.private_calculateFunction("STDEV", arguments);
	};
	/**
	 * Estimates standard deviation based on a sample (ignores logical values and text in the sample).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEV_S.js
	 */
	ApiWorksheetFunction.prototype.STDEV_S = function () {
		return this.private_calculateFunction("STDEV.S", arguments);
	};
	/**
	 * Estimates standard deviation based on a sample, including logical values and text. Text and the <b>false</b> logical value have the value 0; the <b>true</b> logical value has the value 1.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | number[] | ApiRange} args - Up to 255 values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEVA.js
	 */
	ApiWorksheetFunction.prototype.STDEVA = function () {
		return this.private_calculateFunction("STDEVA", arguments);
	};
	/**
	 * Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEVP.js
	 */
	ApiWorksheetFunction.prototype.STDEVP = function () {
		return this.private_calculateFunction("STDEVP", arguments);
	};
	/**
	 * Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | number[] | ApiRange} args - Up to 255 numeric values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEV_P.js
	 */
	ApiWorksheetFunction.prototype.STDEV_P = function () {
		return this.private_calculateFunction("STDEV.P", arguments);
	};
	/**
	 * Calculates standard deviation based on an entire population, including logical values and text. Text and the <b>false</b> logical value have the value 0; the <b>true</b> logical value has the value 1.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | number[] | ApiRange} args - Up to 255 values for which the standard deviation will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/STDEVPA.js
	 */
	ApiWorksheetFunction.prototype.STDEVPA = function () {
		return this.private_calculateFunction("STDEVPA", arguments);
	};
	// todo need array
	// /**
	//  * Returns the standard error of the predicted y-value for each x in a regression.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is an array or range of dependent data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @param {any} arg2 Is an array or range of independent data points and can be numbers or names, arrays, or references that contain numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.STEYX = function (arg1, arg2) {
	// 	return this.private_calculateFunction("STEYX", arguments);
	// };
	/**
	 * Returns the Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The numeric value at which to evaluate the distribution.
	 * @param {number} arg2 - An integer indicating the number of degrees of freedom that characterize the distribution.
	 * @param {number} arg3 - Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TDIST.js
	 */
	ApiWorksheetFunction.prototype.TDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TDIST", arguments);
	};
	/**
	 * Returns the left-tailed Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The numeric value at which to evaluate the distribution.
	 * @param {number} arg2 - An integer indicating the number of degrees of freedom that characterize the distribution.
	 * @param {boolean} arg3 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability density function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T_DIST.js
	 */
	ApiWorksheetFunction.prototype.T_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("T.DIST", arguments);
	};
	/**
	 * Returns the two-tailed Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The numeric value at which to evaluate the distribution.
	 * @param {number} arg2 - An integer indicating the number of degrees of freedom that characterize the distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T_DIST_2T.js
	 */
	ApiWorksheetFunction.prototype.T_DIST_2T = function (arg1, arg2) {
		return this.private_calculateFunction("T.DIST.2T", arguments);
	};
	/**
	 * Returns the right-tailed Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The numeric value at which to evaluate the distribution.
	 * @param {number} arg2 - An integer indicating the number of degrees of freedom that characterize the distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T_DIST_RT.js
	 */
	ApiWorksheetFunction.prototype.T_DIST_RT = function (arg1, arg2) {
		return this.private_calculateFunction("T.DIST.RT", arguments);
	};
	/**
	 * Returns the left-tailed inverse of the Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - A positive integer indicating the number of degrees of freedom to characterize the distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T_INV.js
	 */
	ApiWorksheetFunction.prototype.T_INV = function (arg1, arg2) {
		return this.private_calculateFunction("T.INV", arguments);
	};
	/**
	 * Returns the two-tailed inverse of the Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - A positive integer indicating the number of degrees of freedom to characterize the distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/T_INV_2T.js
	 */
	ApiWorksheetFunction.prototype.T_INV_2T = function (arg1, arg2) {
		return this.private_calculateFunction("T.INV.2T", arguments);
	};
	/**
	 * Returns the two-tailed inverse of the Student's t-distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
	 * @param {number} arg2 - A positive integer indicating the number of degrees of freedom to characterize the distribution.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TINV.js
	 */
	ApiWorksheetFunction.prototype.TINV = function (arg1, arg2) {
		return this.private_calculateFunction("TINV", arguments);
	};
	/**
	 * Returns numbers in a linear trend matching known data points, using the least squares method.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - A range or array of y-values from the <em>y = mx + b</em> equation.
	 * @param {?ApiRange | number[]} arg2 - An optional range or array of x-values from the <em>y = mx + b</em> equation, an array of the same size as an array of y-values.
	 * @param {?ApiRange | number[]} arg3 - A range or array of new x-values for which this function will return corresponding y-values.
	 * @param {?boolean} arg4 - A logical value: the constant <em>b</em> is calculated normally if this parameter is set to <b>true</b> or omitted,
	 * and <em>b</em> is set equal to 0 if the parameter is <b>false</b>.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TREND.js
	 */
	ApiWorksheetFunction.prototype.TREND = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("TREND", arguments);
	};
	/**
	 * Returns the mean of the interior portion of a set of data values.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | number[]} arg1 - The range or array of values to trim and average.
	 * @param {number} arg2 - The fractional number of data points to exclude from the top and bottom of the data set.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TRIMMEAN.js
	 */
	ApiWorksheetFunction.prototype.TRIMMEAN = function (arg1, arg2) {
		return this.private_calculateFunction("TRIMMEAN", arguments);
	};
	// todo need array
	// /**
	//  * Returns the probability associated with a Student's t-Test.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first data set.
	//  * @param {any} arg2 Is the second data set.
	//  * @param {number} arg3 Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2.
	//  * @param {number} arg4 Is the kind of t-test: paired = 1, two-sample equal variance (homoscedastic) = 2, two-sample unequal variance = 3.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.TTEST = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("TTEST", arguments);
	// };
	// todo need array
	// /**
	//  * Returns the probability associated with a Student's t-Test.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first data set.
	//  * @param {any} arg2 Is the second data set.
	//  * @param {number} arg3 Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2.
	//  * @param {number} arg4 Is the kind of t-test: paired = 1, two-sample equal variance (homoscedastic) = 2, two-sample unequal variance = 3.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.T_TEST = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("T.TEST", arguments);
	// };
	/**
	 * Estimates variance based on a sample (ignores logical values and text in the sample).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | ApiRange | number[]} args - Up to 255 numeric values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VAR.js
	 */
	ApiWorksheetFunction.prototype.VAR = function () {
		return this.private_calculateFunction("VAR", arguments);
	};
	/**
	 * Estimates variance based on a sample, including logical values and text. Text and the <b>false</b> logical value have the value 0; the <b>true</b> logical value has the value 1.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | ApiRange | number[]} args - Up to 255 values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VARA.js
	 */
	ApiWorksheetFunction.prototype.VARA = function () {
		return this.private_calculateFunction("VARA", arguments);
	};
	/**
	 * Calculates variance based on the entire population (ignores logical values and text in the population).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | ApiRange | number[]} args - Up to 255 numeric values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VARP.js
	 */
	ApiWorksheetFunction.prototype.VARP = function () {
		return this.private_calculateFunction("VARP", arguments);
	};
	/**
	 * Calculates variance based on the entire population (ignores logical values and text in the population).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | ApiRange | number[]} args - Up to 255 numeric values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VAR_P.js
	 */
	ApiWorksheetFunction.prototype.VAR_P = function () {
		return this.private_calculateFunction("VAR.P", arguments);
	};
	/**
	 * Estimates variance based on a sample (ignores logical values and text in the sample).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | ApiRange | number[]} args - Up to 255 numeric values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VAR_S.js
	 */
	ApiWorksheetFunction.prototype.VAR_S = function () {
		return this.private_calculateFunction("VAR.S", arguments);
	};
	/**
	 * Calculates variance based on the entire population, including logical values and text. Text and the <b>false</b> logical value have the value 0; the <b>true</b> logical value has the value 1.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number | string | boolean | ApiRange | number[]} args - Up to 255 values for which the variance will be calculated. The first argument is required, subsequent arguments are optional.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VARPA.js
	 */
	ApiWorksheetFunction.prototype.VARPA = function () {
		return this.private_calculateFunction("VARPA", arguments);
	};
	/**
	 * Returns the Weibull distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a nonnegative number.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WEIBULL.js
	 */
	ApiWorksheetFunction.prototype.WEIBULL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WEIBULL", arguments);
	};
	/**
	 * Returns the Weibull distribution.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - The value at which to evaluate the function, a nonnegative number.
	 * @param {number} arg2 - The alpha parameter of the distribution, a positive number.
	 * @param {number} arg3 - The beta parameter of the distribution, a positive number.
	 * @param {boolean} arg4 - A logical value (<b>true</b> or <b>false</b>) that determines the function form.
	 * If it is <b>true</b>, the function returns the cumulative distribution function.
	 * If it is <b>false</b>, the function returns the probability mass function.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WEIBULL_DIST.js
	 */
	ApiWorksheetFunction.prototype.WEIBULL_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WEIBULL.DIST", arguments);
	};
	/**
	 * Returns the one-tailed P-value of a z-test.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data against which to test X.
	 * @param {number} arg2 - The value to test.
	 * @param {?number} arg3 - The population (known) standard deviation. If omitted, the sample standard deviation is used.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ZTEST.js
	 */
	ApiWorksheetFunction.prototype.ZTEST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("ZTEST", arguments);
	};
	/**
	 * Returns the one-tailed P-value of a z-test.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number[] | ApiRange} arg1 - The array or range of data against which to test X.
	 * @param {number} arg2 - The value to test.
	 * @param {?number} arg3 - The population (known) standard deviation. If omitted, the sample standard deviation is used.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/Z_TEST.js
	 */
	ApiWorksheetFunction.prototype.Z_TEST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("Z.TEST", arguments);
	};
	/**
	 * Returns a number that represents the date in the date-time code.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A number from 1900 or 1904 (depending on the workbook's date system) to 9999.
	 * @param {number} arg2 - A number from 1 to 12 representing the month of the year.
	 * @param {number} arg3 - A number from 1 to 31 representing the day of the month.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DATE.js
	 */
	ApiWorksheetFunction.prototype.DATE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DATE", arguments);
	};
	/**
	 * Converts a date in the form of text to a number that represents the date in the date-time code.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 - The text that represents a date, between 1/1/1900 or 1/1/1904 (depending on the workbook's date system) and 12/31/9999.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DATEVALUE.js
	 */
	ApiWorksheetFunction.prototype.DATEVALUE = function (arg1) {
		return this.private_calculateFunction("DATEVALUE", arguments);
	};
	/**
	 * Returns the day of the date given in the numerical format, a number from 1 to 31.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - A number in the date-time code.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DAY.js
	 */
	ApiWorksheetFunction.prototype.DAY = function (arg1) {
		return this.private_calculateFunction("DAY", arguments);
	};
	/**
	 * Returns the number of days between the two dates.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - Start date from which days will be counted.
	 * @param {number} arg2 - End date until which days will be counted.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DAYS.js
	 */
	ApiWorksheetFunction.prototype.DAYS = function (arg1, arg2) {
		return this.private_calculateFunction("DAYS", arguments);
	};
	/**
	 * Returns the number of days between two dates based on a 360-day year (twelve 30-day months).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 - Start date from which days will be counted.
	 * @param {number} arg2 - End date until which days will be counted.
	 * @param {?boolean} arg3 - A logical value that specifies whether to use the U.S. (NASD) (false or omitted) or European (true) method in the calculation.
	 * According to the European method, the start and end dates that occur on the 31st of a month become equal to the 30th of the same month.
	 * According to the U.S. method, the start date is the last day of a month, it becomes equal to the 30th of the same month.
	 * If the end date is the last day of a month and the start date is earlier than the 30th of a month, the end date becomes equal to the 1st of the next month.
	 * Otherwise the end date becomes equal to the 30th of the same month.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DAYS360.js
	 */
	ApiWorksheetFunction.prototype.DAYS360 = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DAYS360", arguments);
	};
	/**
	 * Returns the serial number of the date which comes the indicated number of months before or after the start date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 - A serial date number that represents the start date.
	 * @param {any} arg2 - The number of months before or after the start date.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EDATE.js
	 */
	ApiWorksheetFunction.prototype.EDATE = function (arg1, arg2) {
		return this.private_calculateFunction("EDATE", arguments);
	};
	/**
	 * Returns the serial number of the last day of the month before or after a specified number of months.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is the number of months before or after the start_date.
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EOMONTH.js
	 */
	ApiWorksheetFunction.prototype.EOMONTH = function (arg1, arg2) {
		return this.private_calculateFunction("EOMONTH", arguments);
	};
	/**
	 * Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.)..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number in the date-time code used by Microsoft Excel, or text in time format, such as 16:48:00 or 4:48:00 PM.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HOUR.js
	 */
	ApiWorksheetFunction.prototype.HOUR = function (arg1) {
		return this.private_calculateFunction("HOUR", arguments);
	};
	/**
	 * Returns the ISO week number in the year for a given date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the date-time code used by Microsoft Excel for date and time calculation.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISOWEEKNUM.js
	 */
	ApiWorksheetFunction.prototype.ISOWEEKNUM = function (arg1) {
		return this.private_calculateFunction("ISOWEEKNUM", arguments);
	};
	/**
	 * Returns the minute, a number from 0 to 59..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:00 or 4:48:00 PM.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MINUTE.js
	 */
	ApiWorksheetFunction.prototype.MINUTE = function (arg1) {
		return this.private_calculateFunction("MINUTE", arguments);
	};
	/**
	 * Returns the month, a number from 1 (January) to 12 (December)..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number in the date-time code used by Microsoft Excel.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MONTH.js
	 */
	ApiWorksheetFunction.prototype.MONTH = function (arg1) {
		return this.private_calculateFunction("MONTH", arguments);
	};
	/**
	 * Returns the number of whole workdays between two dates.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is a serial date number that represents the end date.
	 * @param {?any} arg3 Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NETWORKDAYS.js
	 */
	ApiWorksheetFunction.prototype.NETWORKDAYS = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NETWORKDAYS", arguments);
	};
	/**
	 * Returns the number of whole workdays between two dates with custom weekend parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is a serial date number that represents the end date.
	 * @param {?number} arg3 Is a number or string specifying when weekends occur.
	 * @param {?any} arg4 Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NETWORKDAYS_INTL.js
	 */
	ApiWorksheetFunction.prototype.NETWORKDAYS_INTL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NETWORKDAYS.INTL", arguments);
	};
	/**
	 * Returns the current date and time formatted as a date and time..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NOW.js
	 */
	ApiWorksheetFunction.prototype.NOW = function () {
		return this.private_calculateFunction("NOW", arguments);
	};
	/**
	 * Returns the second, a number from 0 to 59..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:23 or 4:48:47 PM.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SECOND.js
	 */
	ApiWorksheetFunction.prototype.SECOND = function (arg1) {
		return this.private_calculateFunction("SECOND", arguments);
	};
	/**
	 * Converts hours, minutes and seconds given as numbers to a serial number, formatted with a time format.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number from 0 to 23 representing the hour.
	 * @param {number} arg2 Is a number from 0 to 59 representing the minute.
	 * @param {number} arg3 Is a number from 0 to 59 representing the second.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TIME.js
	 */
	ApiWorksheetFunction.prototype.TIME = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TIME", arguments);
	};
	/**
	 * Converts a text time to a serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a text string that gives a time in any one of the Microsoft Excel time formats (date information in the string is ignored).
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TIMEVALUE.js
	 */
	ApiWorksheetFunction.prototype.TIMEVALUE = function (arg1) {
		return this.private_calculateFunction("TIMEVALUE", arguments);
	};
	/**
	 * Returns the current date formatted as a date..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TODAY.js
	 */
	ApiWorksheetFunction.prototype.TODAY = function () {
		return this.private_calculateFunction("TODAY", arguments);
	};
	/**
	 * Returns a number from 1 to 7 identifying the day of the week of a date..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number that represents a date.
	 * @param {?number} arg2 Is a number: for Sunday=1 through Saturday=7, use 1; for Monday=1 through Sunday=7, use 2; for Monday=0 through Sunday=6, use 3.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WEEKDAY.js
	 */
	ApiWorksheetFunction.prototype.WEEKDAY = function (arg1, arg2) {
		return this.private_calculateFunction("WEEKDAY", arguments);
	};
	/**
	 * Returns the week number in the year.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the date-time code used by Microsoft Excel for date and time calculation.
	 * @param {?any} arg2 Is a number (1 or 2) that determines the type of the return value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WEEKNUM.js
	 */
	ApiWorksheetFunction.prototype.WEEKNUM = function (arg1, arg2) {
		return this.private_calculateFunction("WEEKNUM", arguments);
	};
	/**
	 * Returns the serial number of the date before or after a specified number of workdays.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is the number of nonweekend and non-holiday days before or after start_date.
	 * @param {?any} arg3 Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WORKDAY.js
	 */
	ApiWorksheetFunction.prototype.WORKDAY = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("WORKDAY", arguments);
	};
	/**
	 * Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is the number of nonweekend and non-holiday days before or after start_date.
	 * @param {?number} arg3 Is a number or string specifying when weekends occur.
	 * @param {?any} arg4 Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/WORKDAY_INTL.js
	 */
	ApiWorksheetFunction.prototype.WORKDAY_INTL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WORKDAY.INTL", arguments);
	};
	/**
	 * Returns the year of a date, an integer in the range 1900-9999..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is a number in the date-time code used by Microsoft Excel.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/YEAR.js
	 */
	ApiWorksheetFunction.prototype.YEAR = function (arg1) {
		return this.private_calculateFunction("YEAR", arguments);
	};
	/**
	 * Returns the year fraction representing the number of whole days between start_date and end_date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a serial date number that represents the start date.
	 * @param {any} arg2 Is a serial date number that represents the end date.
	 * @param {?any} arg3 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/YEARFRAC.js
	 */
	ApiWorksheetFunction.prototype.YEARFRAC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("YEARFRAC", arguments);
	};
	/**
	 * Returns the modified Bessel function In(x).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value at which to evaluate the function.
	 * @param {any} arg2 Is the order of the Bessel function.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BESSELI.js
	 */
	ApiWorksheetFunction.prototype.BESSELI = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELI", arguments);
	};
	/**
	 * Returns the Bessel function Jn(x).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value at which to evaluate the function.
	 * @param {any} arg2 Is the order of the Bessel function.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BESSELJ.js
	 */
	ApiWorksheetFunction.prototype.BESSELJ = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELJ", arguments);
	};
	/**
	 * Returns the modified Bessel function Kn(x).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value at which to evaluate the function.
	 * @param {any} arg2 Is the order of the function.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BESSELK.js
	 */
	ApiWorksheetFunction.prototype.BESSELK = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELK", arguments);
	};
	/**
	 * Returns the Bessel function Yn(x).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value at which to evaluate the function.
	 * @param {any} arg2 Is the order of the function.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BESSELY.js
	 */
	ApiWorksheetFunction.prototype.BESSELY = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELY", arguments);
	};
	/**
	 * Converts a binary number to decimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the binary number you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BIN2DEC.js
	 */
	ApiWorksheetFunction.prototype.BIN2DEC = function (arg1) {
		return this.private_calculateFunction("BIN2DEC", arguments);
	};
	/**
	 * Converts a binary number to hexadecimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the binary number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BIN2HEX.js
	 */
	ApiWorksheetFunction.prototype.BIN2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("BIN2HEX", arguments);
	};
	/**
	 * Converts a binary number to octal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the binary number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BIN2OCT.js
	 */
	ApiWorksheetFunction.prototype.BIN2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("BIN2OCT", arguments);
	};
	/**
	 * Returns a bitwise 'And' of two numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the decimal representation of the binary number you want to evaluate.
	 * @param {number} arg2 Is the decimal representation of the binary number you want to evaluate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BITAND.js
	 */
	ApiWorksheetFunction.prototype.BITAND = function (arg1, arg2) {
		return this.private_calculateFunction("BITAND", arguments);
	};
	/**
	 * Returns a number shifted left by shift_amount bits.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the decimal representation of the binary number you want to evaluate.
	 * @param {number} arg2 Is the number of bits that you want to shift Number left by.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BITLSHIFT.js
	 */
	ApiWorksheetFunction.prototype.BITLSHIFT = function (arg1, arg2) {
		return this.private_calculateFunction("BITLSHIFT", arguments);
	};
	/**
	 * Returns a bitwise 'Or' of two numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the decimal representation of the binary number you want to evaluate.
	 * @param {number} arg2 Is the decimal representation of the binary number you want to evaluate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BITOR.js
	 */
	ApiWorksheetFunction.prototype.BITOR = function (arg1, arg2) {
		return this.private_calculateFunction("BITOR", arguments);
	};
	/**
	 * Returns a number shifted right by shift_amount bits.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the decimal representation of the binary number you want to evaluate.
	 * @param {number} arg2 Is the number of bits that you want to shift Number right by.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BITRSHIFT.js
	 */
	ApiWorksheetFunction.prototype.BITRSHIFT = function (arg1, arg2) {
		return this.private_calculateFunction("BITRSHIFT", arguments);
	};
	/**
	 * Returns a bitwise 'Exclusive Or' of two numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the decimal representation of the binary number you want to evaluate.
	 * @param {number} arg2 Is the decimal representation of the binary number you want to evaluate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BITXOR.js
	 */
	ApiWorksheetFunction.prototype.BITXOR = function (arg1, arg2) {
		return this.private_calculateFunction("BITXOR", arguments);
	};
	/**
	 * Converts real and imaginary coefficients into a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the real coefficient of the complex number.
	 * @param {any} arg2 Is the imaginary coefficient of the complex number.
	 * @param {?any} arg3 Is the suffix for the imaginary component of the complex number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COMPLEX.js
	 */
	ApiWorksheetFunction.prototype.COMPLEX = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("COMPLEX", arguments);
	};
	/**
	 * Converts a number from one measurement system to another.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value in from_units to convert.
	 * @param {any} arg2 Is the units for number.
	 * @param {any} arg3 Is the units for the result.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CONVERT.js
	 */
	ApiWorksheetFunction.prototype.CONVERT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONVERT", arguments);
	};
	/**
	 * Converts a decimal number to binary.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the decimal integer you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DEC2BIN.js
	 */
	ApiWorksheetFunction.prototype.DEC2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2BIN", arguments);
	};
	/**
	 * Converts a decimal number to hexadecimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the decimal integer you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DEC2HEX.js
	 */
	ApiWorksheetFunction.prototype.DEC2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2HEX", arguments);
	};
	/**
	 * Converts a decimal number to octal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the decimal integer you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DEC2OCT.js
	 */
	ApiWorksheetFunction.prototype.DEC2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2OCT", arguments);
	};
	/**
	 * Tests whether two numbers are equal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the first number.
	 * @param {?any} arg2 Is the second number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DELTA.js
	 */
	ApiWorksheetFunction.prototype.DELTA = function (arg1, arg2) {
		return this.private_calculateFunction("DELTA", arguments);
	};
	/**
	 * Returns the error function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the lower bound for integrating ERF.
	 * @param {?any} arg2 Is the upper bound for integrating ERF.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ERF.js
	 */
	ApiWorksheetFunction.prototype.ERF = function (arg1, arg2) {
		return this.private_calculateFunction("ERF", arguments);
	};
	/**
	 * Returns the error function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the lower bound for integrating ERF.PRECISE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ERF_PRECISE.js
	 */
	ApiWorksheetFunction.prototype.ERF_PRECISE = function (arg1) {
		return this.private_calculateFunction("ERF.PRECISE", arguments);
	};
	/**
	 * Returns the complementary error function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the lower bound for integrating ERF.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ERFC.js
	 */
	ApiWorksheetFunction.prototype.ERFC = function (arg1) {
		return this.private_calculateFunction("ERFC", arguments);
	};
	/**
	 * Returns the complementary error function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the lower bound for integrating ERFC.PRECISE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ERFC_PRECISE.js
	 */
	ApiWorksheetFunction.prototype.ERFC_PRECISE = function (arg1) {
		return this.private_calculateFunction("ERFC.PRECISE", arguments);
	};
	/**
	 * Tests whether a number is greater than a threshold value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to test against step.
	 * @param {?any} arg2 Is the threshold value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GESTEP.js
	 */
	ApiWorksheetFunction.prototype.GESTEP = function (arg1, arg2) {
		return this.private_calculateFunction("GESTEP", arguments);
	};
	/**
	 * Converts a Hexadecimal number to binary.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the hexadecimal number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HEX2BIN.js
	 */
	ApiWorksheetFunction.prototype.HEX2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("HEX2BIN", arguments);
	};
	/**
	 * Converts a hexadecimal number to decimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the hexadecimal number you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HEX2DEC.js
	 */
	ApiWorksheetFunction.prototype.HEX2DEC = function (arg1) {
		return this.private_calculateFunction("HEX2DEC", arguments);
	};
	/**
	 * Converts a hexadecimal number to octal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the hexadecimal number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HEX2OCT.js
	 */
	ApiWorksheetFunction.prototype.HEX2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("HEX2OCT", arguments);
	};
	/**
	 * Returns the absolute value (modulus) of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the absolute value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMABS.js
	 */
	ApiWorksheetFunction.prototype.IMABS = function (arg1) {
		return this.private_calculateFunction("IMABS", arguments);
	};
	/**
	 * Returns the imaginary coefficient of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the imaginary coefficient.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMAGINARY.js
	 */
	ApiWorksheetFunction.prototype.IMAGINARY = function (arg1) {
		return this.private_calculateFunction("IMAGINARY", arguments);
	};
	/**
	 * Returns the argument q, an angle expressed in radians.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the argument.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMARGUMENT.js
	 */
	ApiWorksheetFunction.prototype.IMARGUMENT = function (arg1) {
		return this.private_calculateFunction("IMARGUMENT", arguments);
	};
	/**
	 * Returns the complex conjugate of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the conjugate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCONJUGATE.js
	 */
	ApiWorksheetFunction.prototype.IMCONJUGATE = function (arg1) {
		return this.private_calculateFunction("IMCONJUGATE", arguments);
	};
	/**
	 * Returns the cosine of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the cosine.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCOS.js
	 */
	ApiWorksheetFunction.prototype.IMCOS = function (arg1) {
		return this.private_calculateFunction("IMCOS", arguments);
	};
	/**
	 * Returns the hyperbolic cosine of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the hyperbolic cosine.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCOSH.js
	 */
	ApiWorksheetFunction.prototype.IMCOSH = function (arg1) {
		return this.private_calculateFunction("IMCOSH", arguments);
	};
	/**
	 * Returns the cotangent of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the cotangent.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCOT.js
	 */
	ApiWorksheetFunction.prototype.IMCOT = function (arg1) {
		return this.private_calculateFunction("IMCOT", arguments);
	};
	/**
	 * Returns the cosecant of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the cosecant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCSC.js
	 */
	ApiWorksheetFunction.prototype.IMCSC = function (arg1) {
		return this.private_calculateFunction("IMCSC", arguments);
	};
	/**
	 * Returns the hyperbolic cosecant of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the hyperbolic cosecant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMCSCH.js
	 */
	ApiWorksheetFunction.prototype.IMCSCH = function (arg1) {
		return this.private_calculateFunction("IMCSCH", arguments);
	};
	/**
	 * Returns the quotient of two complex numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the complex numerator or dividend.
	 * @param {any} arg2 Is the complex denominator or divisor.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMDIV.js
	 */
	ApiWorksheetFunction.prototype.IMDIV = function (arg1, arg2) {
		return this.private_calculateFunction("IMDIV", arguments);
	};
	/**
	 * Returns the exponential of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the exponential.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMEXP.js
	 */
	ApiWorksheetFunction.prototype.IMEXP = function (arg1) {
		return this.private_calculateFunction("IMEXP", arguments);
	};
	/**
	 * Returns the natural logarithm of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the natural logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMLN.js
	 */
	ApiWorksheetFunction.prototype.IMLN = function (arg1) {
		return this.private_calculateFunction("IMLN", arguments);
	};
	/**
	 * Returns the base-10 logarithm of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the common logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMLOG10.js
	 */
	ApiWorksheetFunction.prototype.IMLOG10 = function (arg1) {
		return this.private_calculateFunction("IMLOG10", arguments);
	};
	/**
	 * Returns the base-2 logarithm of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the base-2 logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMLOG2.js
	 */
	ApiWorksheetFunction.prototype.IMLOG2 = function (arg1) {
		return this.private_calculateFunction("IMLOG2", arguments);
	};
	/**
	 * Returns a complex number raised to an integer power.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number you want to raise to a power.
	 * @param {any} arg2 Is the power to which you want to raise the complex number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMPOWER.js
	 */
	ApiWorksheetFunction.prototype.IMPOWER = function (arg1, arg2) {
		return this.private_calculateFunction("IMPOWER", arguments);
	};
	/**
	 * Returns the product of 1 to 255 complex numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMPRODUCT.js
	 */
	ApiWorksheetFunction.prototype.IMPRODUCT = function () {
		return this.private_calculateFunction("IMPRODUCT", arguments);
	};
	/**
	 * Returns the real coefficient of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the real coefficient.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMREAL.js
	 */
	ApiWorksheetFunction.prototype.IMREAL = function (arg1) {
		return this.private_calculateFunction("IMREAL", arguments);
	};
	/**
	 * Returns the secant of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the secant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSEC.js
	 */
	ApiWorksheetFunction.prototype.IMSEC = function (arg1) {
		return this.private_calculateFunction("IMSEC", arguments);
	};
	/**
	 * Returns the hyperbolic secant of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the hyperbolic secant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSECH.js
	 */
	ApiWorksheetFunction.prototype.IMSECH = function (arg1) {
		return this.private_calculateFunction("IMSECH", arguments);
	};
	/**
	 * Returns the sine of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the sine.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSIN.js
	 */
	ApiWorksheetFunction.prototype.IMSIN = function (arg1) {
		return this.private_calculateFunction("IMSIN", arguments);
	};
	/**
	 * Returns the hyperbolic sine of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the hyperbolic sine.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSINH.js
	 */
	ApiWorksheetFunction.prototype.IMSINH = function (arg1) {
		return this.private_calculateFunction("IMSINH", arguments);
	};
	/**
	 * Returns the square root of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the square root.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSQRT.js
	 */
	ApiWorksheetFunction.prototype.IMSQRT = function (arg1) {
		return this.private_calculateFunction("IMSQRT", arguments);
	};
	/**
	 * Returns the difference of two complex numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the complex number from which to subtract inumber2.
	 * @param {any} arg2 Is the complex number to subtract from inumber1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSUB.js
	 */
	ApiWorksheetFunction.prototype.IMSUB = function (arg1, arg2) {
		return this.private_calculateFunction("IMSUB", arguments);
	};
	/**
	 * Returns the sum of complex numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMSUM.js
	 */
	ApiWorksheetFunction.prototype.IMSUM = function () {
		return this.private_calculateFunction("IMSUM", arguments);
	};
	/**
	 * Returns the tangent of a complex number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a complex number for which you want the tangent.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IMTAN.js
	 */
	ApiWorksheetFunction.prototype.IMTAN = function (arg1) {
		return this.private_calculateFunction("IMTAN", arguments);
	};
	/**
	 * Converts an octal number to binary.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the octal number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/OCT2BIN.js
	 */
	ApiWorksheetFunction.prototype.OCT2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("OCT2BIN", arguments);
	};
	/**
	 * Converts an octal number to decimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the octal number you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/OCT2DEC.js
	 */
	ApiWorksheetFunction.prototype.OCT2DEC = function (arg1) {
		return this.private_calculateFunction("OCT2DEC", arguments);
	};
	/**
	 * Converts an octal number to hexadecimal.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the octal number you want to convert.
	 * @param {?any} arg2 Is the number of characters to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/OCT2HEX.js
	 */
	ApiWorksheetFunction.prototype.OCT2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("OCT2HEX", arguments);
	};
	/**
	 * Averages the values in a column in a list or database that match conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DAVERAGE.js
	 */
	ApiWorksheetFunction.prototype.DAVERAGE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DAVERAGE", arguments);
	};
	/**
	 * Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DCOUNT.js
	 */
	ApiWorksheetFunction.prototype.DCOUNT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DCOUNT", arguments);
	};
	/**
	 * Counts nonblank cells in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DCOUNTA.js
	 */
	ApiWorksheetFunction.prototype.DCOUNTA = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DCOUNTA", arguments);
	};
	/**
	 * Extracts from a database a single record that matches the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DGET.js
	 */
	ApiWorksheetFunction.prototype.DGET = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DGET", arguments);
	};
	/**
	 * Returns the largest number in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DMAX.js
	 */
	ApiWorksheetFunction.prototype.DMAX = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DMAX", arguments);
	};
	/**
	 * Returns the smallest number in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DMIN.js
	 */
	ApiWorksheetFunction.prototype.DMIN = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DMIN", arguments);
	};
	/**
	 * Multiplies the values in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DPRODUCT.js
	 */
	ApiWorksheetFunction.prototype.DPRODUCT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DPRODUCT", arguments);
	};
	/**
	 * Estimates the standard deviation based on a sample from selected database entries.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DSTDEV.js
	 */
	ApiWorksheetFunction.prototype.DSTDEV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSTDEV", arguments);
	};
	/**
	 * Calculates the standard deviation based on the entire population of selected database entries.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DSTDEVP.js
	 */
	ApiWorksheetFunction.prototype.DSTDEVP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSTDEVP", arguments);
	};
	/**
	 * Adds the numbers in the field (column) of records in the database that match the conditions you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DSUM.js
	 */
	ApiWorksheetFunction.prototype.DSUM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSUM", arguments);
	};
	/**
	 * Estimates variance based on a sample from selected database entries.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DVAR.js
	 */
	ApiWorksheetFunction.prototype.DVAR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DVAR", arguments);
	};
	/**
	 * Calculates variance based on the entire population of selected database entries.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells that makes up the list or database. A database is a list of related data.
	 * @param {number} arg2 Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
	 * @param {string} arg3 Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DVARP.js
	 */
	ApiWorksheetFunction.prototype.DVARP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DVARP", arguments);
	};
	/**
	 * Returns the accrued interest for a security that pays periodic interest..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's first interest date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's annual coupon rate.
	 * @param {any} arg5 Is the security's par value.
	 * @param {any} arg6 Is the number of coupon payments per year.
	 * @param {?any} arg7 Is the type of day count basis to use.
	 * @param {?any} arg8 Is a logical value: to accrued interest from issue date = TRUE or omitted; to calculate from last coupon payment date = FALSE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACCRINT.js
	 */
	ApiWorksheetFunction.prototype.ACCRINT = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ACCRINT", arguments);
	};
	/**
	 * Returns the accrued interest for a security that pays interest at maturity.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's annual coupon rate.
	 * @param {any} arg4 Is the security's par value.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACCRINTM.js
	 */
	ApiWorksheetFunction.prototype.ACCRINTM = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("ACCRINTM", arguments);
	};
	/**
	 * Returns the prorated linear depreciation of an asset for each accounting period..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {?any} arg7.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AMORDEGRC.js
	 */
	ApiWorksheetFunction.prototype.AMORDEGRC = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("AMORDEGRC", arguments);
	};
	/**
	 * Returns the prorated linear depreciation of an asset for each accounting period..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the cost of the asset.
	 * @param {any} arg2 Is the date the asset is purchased.
	 * @param {any} arg3 Is the date of the end of the first period.
	 * @param {any} arg4 Is the salvage value at the end of life of the asset..
	 * @param {any} arg5 Is the period.
	 * @param {any} arg6 Is the rate of depreciation.
	 * @param {?any} arg7 Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days..
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AMORLINC.js
	 */
	ApiWorksheetFunction.prototype.AMORLINC = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("AMORLINC", arguments);
	};
	/**
	 * Returns the number of days from the beginning of the coupon period to the settlement date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPDAYBS.js
	 */
	ApiWorksheetFunction.prototype.COUPDAYBS = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYBS", arguments);
	};
	/**
	 * Returns the number of days in the coupon period that contains the settlement date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPDAYS.js
	 */
	ApiWorksheetFunction.prototype.COUPDAYS = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYS", arguments);
	};
	/**
	 * Returns the number of days from the settlement date to the next coupon date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPDAYSNC.js
	 */
	ApiWorksheetFunction.prototype.COUPDAYSNC = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYSNC", arguments);
	};
	/**
	 * Returns the next coupon date after the settlement date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPNCD.js
	 */
	ApiWorksheetFunction.prototype.COUPNCD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPNCD", arguments);
	};
	/**
	 * Returns the number of coupons payable between the settlement date and maturity date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPNUM.js
	 */
	ApiWorksheetFunction.prototype.COUPNUM = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPNUM", arguments);
	};
	/**
	 * Returns the previous coupon date before the settlement date.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the number of coupon payments per year.
	 * @param {?any} arg4 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COUPPCD.js
	 */
	ApiWorksheetFunction.prototype.COUPPCD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPPCD", arguments);
	};
	/**
	 * Returns the cumulative interest paid between two periods.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the interest rate.
	 * @param {any} arg2 Is the total number of payment periods.
	 * @param {any} arg3 Is the present value.
	 * @param {any} arg4 Is the first period in the calculation.
	 * @param {any} arg5 Is the last period in the calculation.
	 * @param {any} arg6 Is the timing of the payment.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CUMIPMT.js
	 */
	ApiWorksheetFunction.prototype.CUMIPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("CUMIPMT", arguments);
	};
	/**
	 * Returns the cumulative principal paid on a loan between two periods.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the interest rate.
	 * @param {any} arg2 Is the total number of payment periods.
	 * @param {any} arg3 Is the present value.
	 * @param {any} arg4 Is the first period in the calculation.
	 * @param {any} arg5 Is the last period in the calculation.
	 * @param {any} arg6 Is the timing of the payment.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CUMPRINC.js
	 */
	ApiWorksheetFunction.prototype.CUMPRINC = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("CUMPRINC", arguments);
	};
	/**
	 * Returns the depreciation of an asset for a specified period using the fixed-declining balance method.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the initial cost of the asset.
	 * @param {number} arg2 Is the salvage value at the end of the life of the asset.
	 * @param {number} arg3 Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
	 * @param {number} arg4 Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
	 * @param {?number} arg5 Is the number of months in the first year. If month is omitted, it is assumed to be 12.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DB.js
	 */
	ApiWorksheetFunction.prototype.DB = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DB", arguments);
	};
	/**
	 * Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the initial cost of the asset.
	 * @param {number} arg2 Is the salvage value at the end of the life of the asset.
	 * @param {number} arg3 Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
	 * @param {number} arg4 Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
	 * @param {?number} arg5 Is the rate at which the balance declines. If Factor is omitted, it is assumed to be 2 (the double-declining balance method).
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DDB.js
	 */
	ApiWorksheetFunction.prototype.DDB = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DDB", arguments);
	};
	/**
	 * Returns the discount rate for a security.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's price per $100 face value.
	 * @param {any} arg4 Is the security's redemption value per $100 face value.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DISC.js
	 */
	ApiWorksheetFunction.prototype.DISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DISC", arguments);
	};
	/**
	 * Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a number expressed as a fraction.
	 * @param {any} arg2 Is the integer to use in the denominator of the fraction.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DOLLARDE.js
	 */
	ApiWorksheetFunction.prototype.DOLLARDE = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLARDE", arguments);
	};
	/**
	 * Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a decimal number.
	 * @param {any} arg2 Is the integer to use in the denominator of a fraction.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DOLLARFR.js
	 */
	ApiWorksheetFunction.prototype.DOLLARFR = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLARFR", arguments);
	};
	/**
	 * Returns the annual duration of a security with periodic interest payments.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's annual coupon rate.
	 * @param {any} arg4 Is the security's annual yield.
	 * @param {any} arg5 Is the number of coupon payments per year.
	 * @param {?any} arg6 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DURATION.js
	 */
	ApiWorksheetFunction.prototype.DURATION = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("DURATION", arguments);
	};
	/**
	 * Returns the effective annual interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the nominal interest rate.
	 * @param {any} arg2 Is the number of compounding periods per year.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EFFECT.js
	 */
	ApiWorksheetFunction.prototype.EFFECT = function (arg1, arg2) {
		return this.private_calculateFunction("EFFECT", arguments);
	};
	/**
	 * Returns the future value of an investment based on periodic, constant payments and a constant interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Is the total number of payment periods in the investment.
	 * @param {number} arg3 Is the payment made each period; it cannot change over the life of the investment.
	 * @param {?number} arg4 Is the present value, or the lump-sum amount that a series of future payments is worth now. If omitted, Pv = 0.
	 * @param {?number} arg5 Is a value representing the timing of payment: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FV.js
	 */
	ApiWorksheetFunction.prototype.FV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("FV", arguments);
	};
	/**
	 * Returns the future value of an initial principal after applying a series of compound interest rates.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the present value.
	 * @param {any} arg2 Is an array of interest rates to apply.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FVSCHEDULE.js
	 */
	ApiWorksheetFunction.prototype.FVSCHEDULE = function (arg1, arg2) {
		return this.private_calculateFunction("FVSCHEDULE", arguments);
	};
	/**
	 * Returns the interest rate for a fully invested security.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the amount invested in the security.
	 * @param {any} arg4 Is the amount to be received at maturity.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/INTRATE.js
	 */
	ApiWorksheetFunction.prototype.INTRATE = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("INTRATE", arguments);
	};
	/**
	 * Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Is the period for which you want to find the interest and must be in the range 1 to Nper.
	 * @param {number} arg3 Is the total number of payment periods in an investment.
	 * @param {number} arg4 Is the present value, or the lump-sum amount that a series of future payments is worth now.
	 * @param {?number} arg5 Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, Fv = 0.
	 * @param {?number} arg6 Is a logical value representing the timing of payment: at the end of the period = 0 or omitted, at the beginning of the period = 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IPMT.js
	 */
	ApiWorksheetFunction.prototype.IPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("IPMT", arguments);
	};
	/**
	 * Returns the internal rate of return for a series of cash flows.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return.
	 * @param {?number} arg2 Is a number that you guess is close to the result of IRR; 0.1 (10 percent) if omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IRR.js
	 */
	ApiWorksheetFunction.prototype.IRR = function (arg1, arg2) {
		return this.private_calculateFunction("IRR", arguments);
	};
	/**
	 * Returns the interest paid during a specific period of an investment.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Period for which you want to find the interest.
	 * @param {number} arg3 Number of payment periods in an investment.
	 * @param {number} arg4 Lump sum amount that a series of future payments is right now.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISPMT.js
	 */
	ApiWorksheetFunction.prototype.ISPMT = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("ISPMT", arguments);
	};
	/**
	 * Returns the Macauley modified duration for a security with an assumed par value of $100.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's annual coupon rate.
	 * @param {any} arg4 Is the security's annual yield.
	 * @param {any} arg5 Is the number of coupon payments per year.
	 * @param {?any} arg6 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MDURATION.js
	 */
	ApiWorksheetFunction.prototype.MDURATION = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("MDURATION", arguments);
	};
	/**
	 * Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is an array or a reference to cells that contain numbers that represent a series of payments (negative) and income (positive) at regular periods.
	 * @param {number} arg2 Is the interest rate you pay on the money used in the cash flows.
	 * @param {number} arg3 Is the interest rate you receive on the cash flows as you reinvest them.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MIRR.js
	 */
	ApiWorksheetFunction.prototype.MIRR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MIRR", arguments);
	};
	/**
	 * Returns the annual nominal interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the effective interest rate.
	 * @param {any} arg2 Is the number of compounding periods per year.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NOMINAL.js
	 */
	ApiWorksheetFunction.prototype.NOMINAL = function (arg1, arg2) {
		return this.private_calculateFunction("NOMINAL", arguments);
	};
	/**
	 * Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Is the payment made each period; it cannot change over the life of the investment.
	 * @param {number} arg3 Is the present value, or the lump-sum amount that a series of future payments is worth now.
	 * @param {?number} arg4 Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, zero is used.
	 * @param {?number} arg5 Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NPER.js
	 */
	ApiWorksheetFunction.prototype.NPER = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("NPER", arguments);
	};
	/**
	 * Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NPV.js
	 */
	ApiWorksheetFunction.prototype.NPV = function () {
		return this.private_calculateFunction("NPV", arguments);
	};
	/**
	 * Returns the price per $100 face value of a security with an odd first period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's first coupon date, expressed as a serial date number.
	 * @param {any} arg5 Is the security's interest rate.
	 * @param {any} arg6 Is the security's annual yield.
	 * @param {any} arg7 Is the security's redemption value per $100 face value.
	 * @param {any} arg8 Is the number of coupon payments per year.
	 * @param {?any} arg9 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ODDFPRICE.js
	 */
	ApiWorksheetFunction.prototype.ODDFPRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9) {
		return this.private_calculateFunction("ODDFPRICE", arguments);
	};
	/**
	 * Returns the yield of a security with an odd first period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's first coupon date, expressed as a serial date number.
	 * @param {any} arg5 Is the security's interest rate.
	 * @param {any} arg6 Is the security's price.
	 * @param {any} arg7 Is the security's redemption value per $100 face value.
	 * @param {any} arg8 Is the number of coupon payments per year.
	 * @param {?any} arg9 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ODDFYIELD.js
	 */
	ApiWorksheetFunction.prototype.ODDFYIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9) {
		return this.private_calculateFunction("ODDFYIELD", arguments);
	};
	/**
	 * Returns the price per $100 face value of a security with an odd last period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's last coupon date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's interest rate.
	 * @param {any} arg5 Is the security's annual yield.
	 * @param {any} arg6 Is the security's redemption value per $100 face value.
	 * @param {any} arg7 Is the number of coupon payments per year.
	 * @param {?any} arg8 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ODDLPRICE.js
	 */
	ApiWorksheetFunction.prototype.ODDLPRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ODDLPRICE", arguments);
	};
	/**
	 * Returns the yield of a security with an odd last period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's last coupon date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's interest rate.
	 * @param {any} arg5 Is the security's price.
	 * @param {any} arg6 Is the security's redemption value per $100 face value.
	 * @param {any} arg7 Is the number of coupon payments per year.
	 * @param {?any} arg8 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ODDLYIELD.js
	 */
	ApiWorksheetFunction.prototype.ODDLYIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ODDLYIELD", arguments);
	};
	/**
	 * Returns the number of periods required by an investment to reach a specified value.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period..
	 * @param {number} arg2 Is the present value of the investment.
	 * @param {number} arg3 Is the desired future value of the investment.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PDURATION.js
	 */
	ApiWorksheetFunction.prototype.PDURATION = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PDURATION", arguments);
	};
	/**
	 * Calculates the payment for a loan based on constant payments and a constant interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period for the loan. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Is the total number of payments for the loan.
	 * @param {number} arg3 Is the present value: the total amount that a series of future payments is worth now.
	 * @param {?number} arg4 Is the future value, or a cash balance you want to attain after the last payment is made, 0 (zero) if omitted.
	 * @param {?number} arg5 Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PMT.js
	 */
	ApiWorksheetFunction.prototype.PMT = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PMT", arguments);
	};
	/**
	 * Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Specifies the period and must be in the range 1 to nper.
	 * @param {number} arg3 Is the total number of payment periods in an investment.
	 * @param {number} arg4 Is the present value: the total amount that a series of future payments is worth now.
	 * @param {?number} arg5 Is the future value, or cash balance you want to attain after the last payment is made.
	 * @param {?number} arg6 Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PPMT.js
	 */
	ApiWorksheetFunction.prototype.PPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("PPMT", arguments);
	};
	/**
	 * Returns the price per $100 face value of a security that pays periodic interest.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's annual coupon rate.
	 * @param {any} arg4 Is the security's annual yield.
	 * @param {any} arg5 Is the security's redemption value per $100 face value.
	 * @param {any} arg6 Is the number of coupon payments per year.
	 * @param {?any} arg7 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PRICE.js
	 */
	ApiWorksheetFunction.prototype.PRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("PRICE", arguments);
	};
	/**
	 * Returns the price per $100 face value of a discounted security.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's discount rate.
	 * @param {any} arg4 Is the security's redemption value per $100 face value.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PRICEDISC.js
	 */
	ApiWorksheetFunction.prototype.PRICEDISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PRICEDISC", arguments);
	};
	/**
	 * Returns the price per $100 face value of a security that pays interest at maturity.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's interest rate at date of issue.
	 * @param {any} arg5 Is the security's annual yield.
	 * @param {?any} arg6 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PRICEMAT.js
	 */
	ApiWorksheetFunction.prototype.PRICEMAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("PRICEMAT", arguments);
	};
	/**
	 * Returns the present value of an investment: the total amount that a series of future payments is worth now.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @param {number} arg2 Is the total number of payment periods in an investment.
	 * @param {number} arg3 Is the payment made each period and cannot change over the life of the investment.
	 * @param {?number} arg4 Is the future value, or a cash balance you want to attain after the last payment is made.
	 * @param {?number} arg5 Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PV.js
	 */
	ApiWorksheetFunction.prototype.PV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PV", arguments);
	};
	/**
	 * Returns the interest rate per period of a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the total number of payment periods for the loan or investment.
	 * @param {number} arg2 Is the payment made each period and cannot change over the life of the loan or investment.
	 * @param {number} arg3 Is the present value: the total amount that a series of future payments is worth now.
	 * @param {?number} arg4 Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, uses Fv = 0.
	 * @param {?number} arg5 Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
	 * @param {?number} arg6 Is your guess for what the rate will be; if omitted, Guess = 0.1 (10 percent).
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RATE.js
	 */
	ApiWorksheetFunction.prototype.RATE = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("RATE", arguments);
	};
	/**
	 * Returns the amount received at maturity for a fully invested security.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the amount invested in the security.
	 * @param {any} arg4 Is the security's discount rate.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RECEIVED.js
	 */
	ApiWorksheetFunction.prototype.RECEIVED = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("RECEIVED", arguments);
	};
	/**
	 * Returns an equivalent interest rate for the growth of an investment.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number of periods for the investment.
	 * @param {number} arg2 Is the present value of the investment.
	 * @param {number} arg3 Is the future value of the investment.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RRI.js
	 */
	ApiWorksheetFunction.prototype.RRI = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RRI", arguments);
	};
	/**
	 * Returns the straight-line depreciation of an asset for one period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the initial cost of the asset.
	 * @param {number} arg2 Is the salvage value at the end of the life of the asset.
	 * @param {number} arg3 Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SLN.js
	 */
	ApiWorksheetFunction.prototype.SLN = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SLN", arguments);
	};
	/**
	 * Returns the sum-of-years' digits depreciation of an asset for a specified period.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the initial cost of the asset.
	 * @param {number} arg2 Is the salvage value at the end of the life of the asset.
	 * @param {number} arg3 Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
	 * @param {number} arg4 Is the period and must use the same units as Life.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SYD.js
	 */
	ApiWorksheetFunction.prototype.SYD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SYD", arguments);
	};
	/**
	 * Returns the bond-equivalent yield for a treasury bill.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the Treasury bill's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the Treasury bill's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the Treasury bill's discount rate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TBILLEQ.js
	 */
	ApiWorksheetFunction.prototype.TBILLEQ = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLEQ", arguments);
	};
	/**
	 * Returns the price per $100 face value for a treasury bill.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the Treasury bill's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the Treasury bill's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the Treasury bill's discount rate.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TBILLPRICE.js
	 */
	ApiWorksheetFunction.prototype.TBILLPRICE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLPRICE", arguments);
	};
	/**
	 * Returns the yield for a treasury bill.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the Treasury bill's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the Treasury bill's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the Treasury Bill's price per $100 face value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TBILLYIELD.js
	 */
	ApiWorksheetFunction.prototype.TBILLYIELD = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLYIELD", arguments);
	};
	/**
	 * Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the initial cost of the asset.
	 * @param {number} arg2 Is the salvage value at the end of the life of the asset.
	 * @param {number} arg3 Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
	 * @param {number} arg4 Is the starting period for which you want to calculate the depreciation, in the same units as Life.
	 * @param {number} arg5 Is the ending period for which you want to calculate the depreciation, in the same units as Life.
	 * @param {?number} arg6 Is the rate at which the balance declines, 2 (double-declining balance) if omitted.
	 * @param {?boolean} arg7 Switch to straight-line depreciation when depreciation is greater than the declining balance = FALSE or omitted; do not switch = TRUE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VDB.js
	 */
	ApiWorksheetFunction.prototype.VDB = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("VDB", arguments);
	};
	/**
	 * Returns the internal rate of return for a schedule of cash flows.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a series of cash flows that correspond to a schedule of payments in dates.
	 * @param {any} arg2 Is a schedule of payment dates that corresponds to the cash flow payments.
	 * @param {?any} arg3 Is a number that you guess is close to the result of XIRR.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/XIRR.js
	 */
	ApiWorksheetFunction.prototype.XIRR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("XIRR", arguments);
	};
	/**
	 * Returns the net present value for a schedule of cash flows.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the discount rate to apply to the cash flows.
	 * @param {any} arg2 Is a series of cash flows that correspond to a schedule of payments in dates.
	 * @param {any} arg3 Is a schedule of payment dates that corresponds to the cash flow payments.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/XNPV.js
	 */
	ApiWorksheetFunction.prototype.XNPV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("XNPV", arguments);
	};
	/**
	 * Returns the yield on a security that pays periodic interest.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's annual coupon rate.
	 * @param {any} arg4 Is the security's price per $100 face value.
	 * @param {any} arg5 Is the security's redemption value per $100 face value.
	 * @param {any} arg6 Is the number of coupon payments per year.
	 * @param {?any} arg7 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/YIELD.js
	 */
	ApiWorksheetFunction.prototype.YIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("YIELD", arguments);
	};
	/**
	 * Returns the annual yield for a discounted security. For example, a treasury bill.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's price per $100 face value.
	 * @param {any} arg4 Is the security's redemption value per $100 face value.
	 * @param {?any} arg5 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/YIELDDISC.js
	 */
	ApiWorksheetFunction.prototype.YIELDDISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("YIELDDISC", arguments);
	};
	/**
	 * Returns the annual yield of a security that pays interest at maturity.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the security's settlement date, expressed as a serial date number.
	 * @param {any} arg2 Is the security's maturity date, expressed as a serial date number.
	 * @param {any} arg3 Is the security's issue date, expressed as a serial date number.
	 * @param {any} arg4 Is the security's interest rate at date of issue.
	 * @param {any} arg5 Is the security's price per $100 face value.
	 * @param {?any} arg6 Is the type of day count basis to use.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/YIELDMAT.js
	 */
	ApiWorksheetFunction.prototype.YIELDMAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("YIELDMAT", arguments);
	};
	/**
	 * Returns the absolute value of a number, a number without its sign.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the real number for which you want the absolute value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ABS.js
	 */
	ApiWorksheetFunction.prototype.ABS = function (arg1) {
		return this.private_calculateFunction("ABS", arguments);
	};
	/**
	 * Returns the arccosine of a number, in radians in the range 0 to Pi. The arccosine is the angle whose cosine is Number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the cosine of the angle you want and must be from -1 to 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACOS.js
	 */
	ApiWorksheetFunction.prototype.ACOS = function (arg1) {
		return this.private_calculateFunction("ACOS", arguments);
	};
	/**
	 * Returns the inverse hyperbolic cosine of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number equal to or greater than 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACOSH.js
	 */
	ApiWorksheetFunction.prototype.ACOSH = function (arg1) {
		return this.private_calculateFunction("ACOSH", arguments);
	};
	/**
	 * Returns the arccotangent of a number, in radians in the range 0 to Pi..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the cotangent of the angle you want.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACOT.js
	 */
	ApiWorksheetFunction.prototype.ACOT = function (arg1) {
		return this.private_calculateFunction("ACOT", arguments);
	};
	/**
	 * Returns the inverse hyperbolic cotangent of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the hyperbolic cotangent of the angle that you want.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ACOTH.js
	 */
	ApiWorksheetFunction.prototype.ACOTH = function (arg1) {
		return this.private_calculateFunction("ACOTH", arguments);
	};
	/**
	 * Returns an aggregate in a list or database.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AGGREGATE.js
	 */
	ApiWorksheetFunction.prototype.AGGREGATE = function () {
		return this.private_calculateFunction("AGGREGATE", arguments);
	};
	/**
	 * Converts a Roman numeral to Arabic.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 Is the Roman numeral you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ARABIC.js
	 */
	ApiWorksheetFunction.prototype.ARABIC = function (arg1) {
		return this.private_calculateFunction("ARABIC", arguments);
	};
	/**
	 * Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the sine of the angle you want and must be from -1 to 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ASIN.js
	 */
	ApiWorksheetFunction.prototype.ASIN = function (arg1) {
		return this.private_calculateFunction("ASIN", arguments);
	};
	/**
	 * Returns the inverse hyperbolic sine of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number equal to or greater than 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ASINH.js
	 */
	ApiWorksheetFunction.prototype.ASINH = function (arg1) {
		return this.private_calculateFunction("ASINH", arguments);
	};
	/**
	 * Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the tangent of the angle you want.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ATAN.js
	 */
	ApiWorksheetFunction.prototype.ATAN = function (arg1) {
		return this.private_calculateFunction("ATAN", arguments);
	};
	/**
	 * Returns the arctangent of the specified x and y coordinates, in radians between -Pi and Pi, excluding -Pi.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the x-coordinate of the point.
	 * @param {number} arg2 Is the y-coordinate of the point.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ATAN2.js
	 */
	ApiWorksheetFunction.prototype.ATAN2 = function (arg1, arg2) {
		return this.private_calculateFunction("ATAN2", arguments);
	};
	/**
	 * Returns the inverse hyperbolic tangent of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number between -1 and 1 excluding -1 and 1.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ATANH.js
	 */
	ApiWorksheetFunction.prototype.ATANH = function (arg1) {
		return this.private_calculateFunction("ATANH", arguments);
	};
	/**
	 * Converts a number into a text representation with the given radix (base).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number that you want to convert.
	 * @param {number} arg2 Is the base Radix that you want to convert the number into.
	 * @param {?number} arg3 Is the minimum length of the returned string.  If omitted leading zeros are not added.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/BASE.js
	 */
	ApiWorksheetFunction.prototype.BASE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("BASE", arguments);
	};
	/**
	 * Rounds a number up, to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the value you want to round.
	 * @param {number} arg2 Is the multiple to which you want to round.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CEILING.js
	 */
	ApiWorksheetFunction.prototype.CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("CEILING", arguments);
	};
	/**
	 * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the value you want to round.
	 * @param {?number} arg2 Is the multiple to which you want to round.
	 * @param {?number} arg3 When given and nonzero this function will round away from zero.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CEILING_MATH.js
	 */
	ApiWorksheetFunction.prototype.CEILING_MATH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CEILING.MATH", arguments);
	};
	/**
	 * Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {?number} arg2.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CEILING_PRECISE.js
	 */
	ApiWorksheetFunction.prototype.CEILING_PRECISE = function (arg1, arg2) {
		return this.private_calculateFunction("CEILING.PRECISE", arguments);
	};
	/**
	 * Returns the number of combinations for a given number of items.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the total number of items.
	 * @param {number} arg2 Is the number of items in each combination.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COMBIN.js
	 */
	ApiWorksheetFunction.prototype.COMBIN = function (arg1, arg2) {
		return this.private_calculateFunction("COMBIN", arguments);
	};
	/**
	 * Returns the number of combinations with repetitions for a given number of items.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the total number of items.
	 * @param {number} arg2 Is the number of items in each combination.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COMBINA.js
	 */
	ApiWorksheetFunction.prototype.COMBINA = function (arg1, arg2) {
		return this.private_calculateFunction("COMBINA", arguments);
	};
	/**
	 * Returns the cosine of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the cosine.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COS.js
	 */
	ApiWorksheetFunction.prototype.COS = function (arg1) {
		return this.private_calculateFunction("COS", arguments);
	};
	/**
	 * Returns the hyperbolic cosine of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COSH.js
	 */
	ApiWorksheetFunction.prototype.COSH = function (arg1) {
		return this.private_calculateFunction("COSH", arguments);
	};
	/**
	 * Returns the cotangent of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the cotangent.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COT.js
	 */
	ApiWorksheetFunction.prototype.COT = function (arg1) {
		return this.private_calculateFunction("COT", arguments);
	};
	/**
	 * Returns the hyperbolic cotangent of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the hyperbolic cotangent.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COTH.js
	 */
	ApiWorksheetFunction.prototype.COTH = function (arg1) {
		return this.private_calculateFunction("COTH", arguments);
	};
	/**
	 * Returns the cosecant of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the cosecant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CSC.js
	 */
	ApiWorksheetFunction.prototype.CSC = function (arg1) {
		return this.private_calculateFunction("CSC", arguments);
	};
	/**
	 * Returns the hyperbolic cosecant of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the hyperbolic cosecant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CSCH.js
	 */
	ApiWorksheetFunction.prototype.CSCH = function (arg1) {
		return this.private_calculateFunction("CSCH", arguments);
	};
	/**
	 * Converts a text representation of a number in a given base into a decimal number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 Is the number that you want to convert.
	 * @param {number} arg2 Is the base Radix of the number you are converting.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DECIMAL.js
	 */
	ApiWorksheetFunction.prototype.DECIMAL = function (arg1, arg2) {
		return this.private_calculateFunction("DECIMAL", arguments);
	};
	/**
	 * Converts radians to degrees.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians that you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/DEGREES.js
	 */
	ApiWorksheetFunction.prototype.DEGREES = function (arg1) {
		return this.private_calculateFunction("DEGREES", arguments);
	};
	/**
	 * Rounds the number up to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ECMA_CEILING.js
	 */
	ApiWorksheetFunction.prototype.ECMA_CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("ECMA.CEILING", arguments);
	};
	/**
	 * Rounds a positive number up and negative number down to the nearest even integer.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the value to round.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EVEN.js
	 */
	ApiWorksheetFunction.prototype.EVEN = function (arg1) {
		return this.private_calculateFunction("EVEN", arguments);
	};
	/**
	 * Returns e raised to the power of a given number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the exponent applied to the base e. The constant e equals 2.71828182845904, the base of the natural logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/EXP.js
	 */
	ApiWorksheetFunction.prototype.EXP = function (arg1) {
		return this.private_calculateFunction("EXP", arguments);
	};
	/**
	 * Returns the factorial of a number, equal to 1*2*3*...* Number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the nonnegative number you want the factorial of.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FACT.js
	 */
	ApiWorksheetFunction.prototype.FACT = function (arg1) {
		return this.private_calculateFunction("FACT", arguments);
	};
	/**
	 * Returns the double factorial of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value for which to return the double factorial.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FACTDOUBLE.js
	 */
	ApiWorksheetFunction.prototype.FACTDOUBLE = function (arg1) {
		return this.private_calculateFunction("FACTDOUBLE", arguments);
	};
	/**
	 * Rounds a number down to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the numeric value you want to round.
	 * @param {number} arg2 Is the multiple to which you want to round. Number and Significance must either both be positive or both be negative.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FLOOR.js
	 */
	ApiWorksheetFunction.prototype.FLOOR = function (arg1, arg2) {
		return this.private_calculateFunction("FLOOR", arguments);
	};
	/**
	 * Returns a number that is rounded down to the nearest integer or to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {?number} arg2.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FLOOR_PRECISE.js
	 */
	ApiWorksheetFunction.prototype.FLOOR_PRECISE = function (arg1, arg2) {
		return this.private_calculateFunction("FLOOR.PRECISE", arguments);
	};
	/**
	 * Rounds a number down, to the nearest integer or to the nearest multiple of significance.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the value you want to round.
	 * @param {?number} arg2 Is the multiple to which you want to round.
	 * @param {?number} arg3 When given and nonzero this function will round towards zero.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FLOOR_MATH.js
	 */
	ApiWorksheetFunction.prototype.FLOOR_MATH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FLOOR.MATH", arguments);
	};
	/**
	 * Returns the greatest common divisor.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/GCD.js
	 */
	ApiWorksheetFunction.prototype.GCD = function () {
		return this.private_calculateFunction("GCD", arguments);
	};
	/**
	 * Rounds a number down to the nearest integer.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the real number you want to round down to an integer.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/INT.js
	 */
	ApiWorksheetFunction.prototype.INT = function (arg1) {
		return this.private_calculateFunction("INT", arguments);
	};
	/**
	 * Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance regardless of the sign of the number. However, if the number or the significance is zero, zero is returned..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {?number} arg2.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISO_CEILING.js
	 */
	ApiWorksheetFunction.prototype.ISO_CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("ISO.CEILING", arguments);
	};
	/**
	 * Returns the least common multiple.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LCM.js
	 */
	ApiWorksheetFunction.prototype.LCM = function () {
		return this.private_calculateFunction("LCM", arguments);
	};
	/**
	 * Returns the natural logarithm of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the positive real number for which you want the natural logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LN.js
	 */
	ApiWorksheetFunction.prototype.LN = function (arg1) {
		return this.private_calculateFunction("LN", arguments);
	};
	/**
	 * Returns the logarithm of a number to the base you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the positive real number for which you want the logarithm.
	 * @param {?number} arg2 Is the base of the logarithm; 10 if omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOG.js
	 */
	ApiWorksheetFunction.prototype.LOG = function (arg1, arg2) {
		return this.private_calculateFunction("LOG", arguments);
	};
	/**
	 * Returns the base-10 logarithm of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the positive real number for which you want the base-10 logarithm.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOG10.js
	 */
	ApiWorksheetFunction.prototype.LOG10 = function (arg1) {
		return this.private_calculateFunction("LOG10", arguments);
	};
	// todo need array
	// /**
	//  * Returns the matrix determinant of an array.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is a numeric array with an equal number of rows and columns, either a cell range or an array constant.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MDETERM = function (arg1) {
	// 	return this.private_calculateFunction("MDETERM", arguments);
	// };
	// todo need array
	// /**
	//  * Returns the inverse matrix for the matrix stored in an array.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is a numeric array with an equal number of rows and columns, either a cell range or an array constant.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MINVERSE = function (arg1) {
	// 	return this.private_calculateFunction("MINVERSE", arguments);
	// };
	// todo need array
	// /**
	//  * Returns the matrix product of two arrays, an array with the same number of rows as array1 and columns as array2.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first array of numbers to multiply and must have the same number of columns as Array2 has rows.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MMULT = function (arg1, arg2) {
	// 	return this.private_calculateFunction("MMULT", arguments);
	// };
	/**
	 * Returns the remainder after a number is divided by a divisor.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number for which you want to find the remainder after the division is performed.
	 * @param {number} arg2 Is the number by which you want to divide Number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MOD.js
	 */
	ApiWorksheetFunction.prototype.MOD = function (arg1, arg2) {
		return this.private_calculateFunction("MOD", arguments);
	};
	/**
	 * Returns a number rounded to the desired multiple.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to round.
	 * @param {any} arg2 Is the multiple to which you want to round number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MROUND.js
	 */
	ApiWorksheetFunction.prototype.MROUND = function (arg1, arg2) {
		return this.private_calculateFunction("MROUND", arguments);
	};
	/**
	 * Returns the multinomial of a set of numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MULTINOMIAL.js
	 */
	ApiWorksheetFunction.prototype.MULTINOMIAL = function () {
		return this.private_calculateFunction("MULTINOMIAL", arguments);
	};
	/**
	 * Returns the unit matrix for the specified dimension.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is an integer specifying the dimension of the unit matrix that you want to return.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MUNIT.js
	 */
	ApiWorksheetFunction.prototype.MUNIT = function (arg1) {
		return this.private_calculateFunction("MUNIT", arguments);
	};
	/**
	 * Rounds a positive number up and negative number down to the nearest odd integer.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the value to round.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ODD.js
	 */
	ApiWorksheetFunction.prototype.ODD = function (arg1) {
		return this.private_calculateFunction("ODD", arguments);
	};
	/**
	 * Returns the value of Pi, 3.14159265358979, accurate to 15 digits.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PI.js
	 */
	ApiWorksheetFunction.prototype.PI = function () {
		return this.private_calculateFunction("PI", arguments);
	};
	/**
	 * Returns the result of a number raised to a power.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the base number, any real number.
	 * @param {number} arg2 Is the exponent, to which the base number is raised.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/POWER.js
	 */
	ApiWorksheetFunction.prototype.POWER = function (arg1, arg2) {
		return this.private_calculateFunction("POWER", arguments);
	};
	/**
	 * Multiplies all the numbers given as arguments.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/PRODUCT.js
	 */
	ApiWorksheetFunction.prototype.PRODUCT = function () {
		return this.private_calculateFunction("PRODUCT", arguments);
	};
	/**
	 * Returns the integer portion of a division.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the dividend.
	 * @param {any} arg2 Is the divisor.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/QUOTIENT.js
	 */
	ApiWorksheetFunction.prototype.QUOTIENT = function (arg1, arg2) {
		return this.private_calculateFunction("QUOTIENT", arguments);
	};
	/**
	 * Converts degrees to radians.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is an angle in degrees that you want to convert.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RADIANS.js
	 */
	ApiWorksheetFunction.prototype.RADIANS = function (arg1) {
		return this.private_calculateFunction("RADIANS", arguments);
	};
	/**
	 * Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RAND.js
	 */
	ApiWorksheetFunction.prototype.RAND = function () {
		return this.private_calculateFunction("RAND", arguments);
	};
	/**
	 * Returns a random number between the numbers you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the smallest integer RANDBETWEEN will return.
	 * @param {any} arg2 Is the largest integer RANDBETWEEN will return.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/RANDBETWEEN.js
	 */
	ApiWorksheetFunction.prototype.RANDBETWEEN = function (arg1, arg2) {
		return this.private_calculateFunction("RANDBETWEEN", arguments);
	};
	/**
	 * Converts an Arabic numeral to Roman, as text.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the Arabic numeral you want to convert.
	 * @param {?number} arg2 Is the number specifying the type of Roman numeral you want..
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ROMAN.js
	 */
	ApiWorksheetFunction.prototype.ROMAN = function (arg1, arg2) {
		return this.private_calculateFunction("ROMAN", arguments);
	};
	/**
	 * Rounds a number to a specified number of digits.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number you want to round.
	 * @param {number} arg2 Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero to the nearest integer.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ROUND.js
	 */
	ApiWorksheetFunction.prototype.ROUND = function (arg1, arg2) {
		return this.private_calculateFunction("ROUND", arguments);
	};
	/**
	 * Rounds a number down, towards zero.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number that you want rounded down.
	 * @param {number} arg2 Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ROUNDDOWN.js
	 */
	ApiWorksheetFunction.prototype.ROUNDDOWN = function (arg1, arg2) {
		return this.private_calculateFunction("ROUNDDOWN", arguments);
	};
	/**
	 * Rounds a number up, away from zero.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number that you want rounded up.
	 * @param {number} arg2 Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ROUNDUP.js
	 */
	ApiWorksheetFunction.prototype.ROUNDUP = function (arg1, arg2) {
		return this.private_calculateFunction("ROUNDUP", arguments);
	};
	/**
	 * Returns the secant of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the secant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SEC.js
	 */
	ApiWorksheetFunction.prototype.SEC = function (arg1) {
		return this.private_calculateFunction("SEC", arguments);
	};
	/**
	 * Returns the hyperbolic secant of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the hyperbolic secant.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SECH.js
	 */
	ApiWorksheetFunction.prototype.SECH = function (arg1) {
		return this.private_calculateFunction("SECH", arguments);
	};
	/**
	 * Returns the sum of a power series based on the formula.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the input value to the power series.
	 * @param {any} arg2 Is the initial power to which you want to raise x.
	 * @param {any} arg3 Is the step by which to increase n for each term in the series.
	 * @param {any} arg4 Is a set of coefficients by which each successive power of x is multiplied.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SERIESSUM.js
	 */
	ApiWorksheetFunction.prototype.SERIESSUM = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SERIESSUM", arguments);
	};
	/**
	 * Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SIGN.js
	 */
	ApiWorksheetFunction.prototype.SIGN = function (arg1) {
		return this.private_calculateFunction("SIGN", arguments);
	};
	/**
	 * Returns the sine of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the sine. Degrees * PI()/180 = radians.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SIN.js
	 */
	ApiWorksheetFunction.prototype.SIN = function (arg1) {
		return this.private_calculateFunction("SIN", arguments);
	};
	/**
	 * Returns the hyperbolic sine of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SINH.js
	 */
	ApiWorksheetFunction.prototype.SINH = function (arg1) {
		return this.private_calculateFunction("SINH", arguments);
	};
	/**
	 * Returns the square root of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number for which you want the square root.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SQRT.js
	 */
	ApiWorksheetFunction.prototype.SQRT = function (arg1) {
		return this.private_calculateFunction("SQRT", arguments);
	};
	/**
	 * Returns the square root of (number * Pi).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the number by which p is multiplied.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SQRTPI.js
	 */
	ApiWorksheetFunction.prototype.SQRTPI = function (arg1) {
		return this.private_calculateFunction("SQRTPI", arguments);
	};
	/**
	 * Returns a subtotal in a list or database.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUBTOTAL.js
	 */
	ApiWorksheetFunction.prototype.SUBTOTAL = function () {
		return this.private_calculateFunction("SUBTOTAL", arguments);
	};
	/**
	 * Adds all the numbers in a range of cells.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUM.js
	 */
	ApiWorksheetFunction.prototype.SUM = function () {
		return this.private_calculateFunction("SUM", arguments);
	};
	/**
	 * Adds the cells specified by a given condition or criteria.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is the range of cells you want evaluated.
	 * @param {any} arg2 Is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.
	 * @param {?ApiRange} arg3 Are the actual cells to sum. If omitted, the cells in range are used.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUMIF.js
	 */
	ApiWorksheetFunction.prototype.SUMIF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SUMIF", arguments);
	};
	/**
	 * Adds the cells specified by a given set of conditions or criteria.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUMIFS.js
	 */
	ApiWorksheetFunction.prototype.SUMIFS = function () {
		return this.private_calculateFunction("SUMIFS", arguments);
	};
	// todo need array
	// /**
	//  * Returns the sum of the products of corresponding ranges or arrays.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMPRODUCT = function () {
	// 	return this.private_calculateFunction("SUMPRODUCT", arguments);
	// };
	/**
	 * Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names or references to cells that contain numbers.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SUMSQ.js
	 */
	ApiWorksheetFunction.prototype.SUMSQ = function () {
		return this.private_calculateFunction("SUMSQ", arguments);
	};
	// todo need array
	// /**
	//  * Sums the differences between the squares of two corresponding ranges or arrays.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first range or array of numbers and can be a number or name, array, or reference that contains numbers.
	//  * @param {any} arg2 Is the second range or array of numbers and can be a number or name, array, or reference that contains numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMX2MY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMX2MY2", arguments);
	// };
	// todo need array
	// /**
	//  * Returns the sum total of the sums of squares of numbers in two corresponding ranges or arrays.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first range or array of numbers and can be a number or name, array, or reference that contains numbers.
	//  * @param {any} arg2 Is the second range or array of numbers and can be a number or name, array, or reference that contains numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMX2PY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMX2PY2", arguments);
	// };
	// todo need array
	// /**
	//  * Sums the squares of the differences in two corresponding ranges or arrays.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1 Is the first range or array of values and can be a number or name, array, or reference that contains numbers.
	//  * @param {any} arg2 Is the second range or array of values and can be a number or name, array, or reference that contains numbers.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMXMY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMXMY2", arguments);
	// };
	/**
	 * Returns the tangent of an angle.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the angle in radians for which you want the tangent. Degrees * PI()/180 = radians.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TAN.js
	 */
	ApiWorksheetFunction.prototype.TAN = function (arg1) {
		return this.private_calculateFunction("TAN", arguments);
	};
	/**
	 * Returns the hyperbolic tangent of a number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is any real number.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TANH.js
	 */
	ApiWorksheetFunction.prototype.TANH = function (arg1) {
		return this.private_calculateFunction("TANH", arguments);
	};
	/**
	 * Truncates a number to an integer by removing the decimal, or fractional, part of the number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1 Is the number you want to truncate.
	 * @param {?number} arg2 Is a number specifying the precision of the truncation, 0 (zero) if omitted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TRUNC.js
	 */
	ApiWorksheetFunction.prototype.TRUNC = function (arg1, arg2) {
		return this.private_calculateFunction("TRUNC", arguments);
	};
	/**
	 * Chooses a value or action to perform from a list of values, based on an index number.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/CHOOSE.js
	 */
	ApiWorksheetFunction.prototype.CHOOSE = function () {
		return this.private_calculateFunction("CHOOSE", arguments);
	};
	/**
	 * Returns the number of columns in an array or reference.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is an array or array formula, or a reference to a range of cells for which you want the number of columns.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/COLUMNS.js
	 */
	ApiWorksheetFunction.prototype.COLUMNS = function (arg1) {
		return this.private_calculateFunction("COLUMNS", arguments);
	};
	/**
	 * Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to be found in the first row of the table and can be a value, a reference, or a text string.
	 * @param {number} arg2 Is a table of text, numbers, or logical values in which data is looked up. Table_array can be a reference to a range or a range name.
	 * @param {number} arg3 Is the row number in table_array from which the matching value should be returned. The first row of values in the table is row 1.
	 * @param {?boolean} arg4 Is a logical value: to find the closest match in the top row (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HLOOKUP.js
	 */
	ApiWorksheetFunction.prototype.HLOOKUP = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("HLOOKUP", arguments);
	};
	/**
	 * Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1 Is the text giving the path and file name to the document to be opened, a hard drive location, UNC address, or URL path.
	 * @param {?any} arg2 Is text or a number that is displayed in the cell. If omitted, the cell displays the Link_location text.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/HYPERLINK.js
	 */
	ApiWorksheetFunction.prototype.HYPERLINK = function (arg1, arg2) {
		return this.private_calculateFunction("HYPERLINK", arguments);
	};
	/**
	 * Returns a value or reference of the cell at the intersection of a particular row and column, in a given range.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is a range of cells or an array constant..
	 * @param {number} arg2 Selects the row in Array or Reference from which to return a value. If omitted, Column_num is required.
	 * @param {?number} arg3 Selects the column in Array or Reference from which to return a value. If omitted, Row_num is required.
	 * @param {?any} arg4 Is a reference to one or more cell ranges.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/INDEX.js
	 */
	ApiWorksheetFunction.prototype.INDEX = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("INDEX", arguments);
	};
	/**
	 * Looks up a value either from a one-row or one-column range or from an array. Provided for backwards compatibility.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a value that LOOKUP searches for in Lookup_vector and can be a number, text, a logical value, or a name or reference to a value.
	 * @param {ApiRange} arg2 Is a range that contains only one row or one column of text, numbers, or logical values, placed in ascending order.
	 * @param {?ApiRange} arg3 Is a range that contains only one row or column, the same size as Lookup_vector.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/LOOKUP.js
	 */
	ApiWorksheetFunction.prototype.LOOKUP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOOKUP", arguments);
	};
	/**
	 * Returns the relative position of an item in an array that matches a specified value in a specified order.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you use to find the value you want in the array, a number, text, or logical value, or a reference to one of these.
	 * @param {number} arg2 Is a contiguous range of cells containing possible lookup values, an array of values, or a reference to an array.
	 * @param {?number} arg3 Is a number 1, 0, or -1 indicating which value to return..
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/MATCH.js
	 */
	ApiWorksheetFunction.prototype.MATCH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MATCH", arguments);
	};
	/**
	 * Returns the number of rows in a reference or array.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is an array, an array formula, or a reference to a range of cells for which you want the number of rows.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ROWS.js
	 */
	ApiWorksheetFunction.prototype.ROWS = function (arg1) {
		return this.private_calculateFunction("ROWS", arguments);
	};
	/**
	 * Converts a vertical range of cells to a horizontal range, or vice versa.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is a range of cells on a worksheet or an array of values that you want to transpose.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TRANSPOSE.js
	 */
	ApiWorksheetFunction.prototype.TRANSPOSE = function (arg1) {
		return this.private_calculateFunction("TRANSPOSE", arguments);
	};
	/**
	 * Looks for a value in the leftmost column of a table and then returns a value in the same row from a column that you specify. By default, the table must be sorted in an ascending order.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to be found in the first column of the table, and can be a value, a reference, or a text string.
	 * @param {number} arg2 Is a table of text, numbers, or logical values, in which data is retrieved. Table_array can be a reference to a range or a range name.
	 * @param {number} arg3 Is the column number in table_array from which the matching value should be returned. The first column of values in the table is column 1.
	 * @param {?boolean} arg4 Is a logical value: to find the closest match in the first column (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/VLOOKUP.js
	 */
	ApiWorksheetFunction.prototype.VLOOKUP = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("VLOOKUP", arguments);
	};
	/**
	 * Returns a number matching an error value..
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the error value for which you want the identifying number, and can be an actual error value or a reference to a cell containing an error value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ERROR_TYPE.js
	 */
	ApiWorksheetFunction.prototype.ERROR_TYPE = function (arg1) {
		return this.private_calculateFunction("ERROR.TYPE", arguments);
	};
	/**
	 * Checks whether a value is an error other than #N/A, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISERR.js
	 */
	ApiWorksheetFunction.prototype.ISERR = function (arg1) {
		return this.private_calculateFunction("ISERR", arguments);
	};
	/**
	 * Checks whether a value is an error, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISERROR.js
	 */
	ApiWorksheetFunction.prototype.ISERROR = function (arg1) {
		return this.private_calculateFunction("ISERROR", arguments);
	};
	/**
	 * Returns TRUE if the number is even.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to test.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISEVEN.js
	 */
	ApiWorksheetFunction.prototype.ISEVEN = function (arg1) {
		return this.private_calculateFunction("ISEVEN", arguments);
	};
	/**
	 * Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1 Is a reference to the cell you want to test.  Reference can be a cell reference, a formula, or name that refers to a cell.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISFORMULA.js
	 */
	ApiWorksheetFunction.prototype.ISFORMULA = function (arg1) {
		return this.private_calculateFunction("ISFORMULA", arguments);
	};
	/**
	 * Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISLOGICAL.js
	 */
	ApiWorksheetFunction.prototype.ISLOGICAL = function (arg1) {
		return this.private_calculateFunction("ISLOGICAL", arguments);
	};
	/**
	 * Checks whether a value is #N/A, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISNA.js
	 */
	ApiWorksheetFunction.prototype.ISNA = function (arg1) {
		return this.private_calculateFunction("ISNA", arguments);
	};
	/**
	 * Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want tested: a cell; a formula; or a name referring to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISNONTEXT.js
	 */
	ApiWorksheetFunction.prototype.ISNONTEXT = function (arg1) {
		return this.private_calculateFunction("ISNONTEXT", arguments);
	};
	/**
	 * Checks whether a value is a number, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISNUMBER.js
	 */
	ApiWorksheetFunction.prototype.ISNUMBER = function (arg1) {
		return this.private_calculateFunction("ISNUMBER", arguments);
	};
	/**
	 * Returns TRUE if the number is odd.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value to test.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISODD.js
	 */
	ApiWorksheetFunction.prototype.ISODD = function (arg1) {
		return this.private_calculateFunction("ISODD", arguments);
	};
	/**
	 * Checks whether a value is a reference, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISREF.js
	 */
	ApiWorksheetFunction.prototype.ISREF = function (arg1) {
		return this.private_calculateFunction("ISREF", arguments);
	};
	/**
	 * Checks whether a value is text, and returns TRUE or FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/ISTEXT.js
	 */
	ApiWorksheetFunction.prototype.ISTEXT = function (arg1) {
		return this.private_calculateFunction("ISTEXT", arguments);
	};
	/**
	 * Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is the value you want converted.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/N.js
	 */
	ApiWorksheetFunction.prototype.N = function (arg1) {
		return this.private_calculateFunction("N", arguments);
	};
	/**
	 * Returns the error value #N/A (value not available).
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NA.js
	 */
	ApiWorksheetFunction.prototype.NA = function () {
		return this.private_calculateFunction("NA", arguments);
	};
	/**
	 * Returns the sheet number of the referenced sheet.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {?string} arg1 Is the name of a sheet or a reference that you want the sheet number of.  If omitted the number of the sheet containing the function is returned.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SHEET.js
	 */
	ApiWorksheetFunction.prototype.SHEET = function (arg1) {
		return this.private_calculateFunction("SHEET", arguments);
	};
	/**
	 * Returns the number of sheets in a reference.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {?ApiRange} arg1 Is a reference for which you want to know the number of sheets it contains.  If omitted the number of sheets in the workbook containing the function is returned.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/SHEETS.js
	 */
	ApiWorksheetFunction.prototype.SHEETS = function (arg1) {
		return this.private_calculateFunction("SHEETS", arguments);
	};
	/**
	 * Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64; compound data = 128.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Can be any value.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TYPE.js
	 */
	ApiWorksheetFunction.prototype.TYPE = function (arg1) {
		return this.private_calculateFunction("TYPE", arguments);
	};
	/**
	 * Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/AND.js
	 */
	ApiWorksheetFunction.prototype.AND = function () {
		return this.private_calculateFunction("AND", arguments);
	};
	/**
	 * Returns the logical value FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/FALSE.js
	 */
	ApiWorksheetFunction.prototype.FALSE = function () {
		return this.private_calculateFunction("FALSE", arguments);
	};
	/**
	 * Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {boolean} arg1 Is any value or expression that can be evaluated to TRUE or FALSE.
	 * @param {any} arg2 Is the value that is returned if Logical_test is TRUE. If omitted, TRUE is returned. You can nest up to seven IF functions.
	 * @param {?any} arg3 Is the value that is returned if Logical_test is FALSE. If omitted, FALSE is returned.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IF.js
	 */
	ApiWorksheetFunction.prototype.IF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("IF", arguments);
	};
	/**
	 * Returns value_if_error if expression is an error and the value of the expression itself otherwise.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is any value or expression or reference.
	 * @param {any} arg2 Is any value or expression or reference.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IFERROR.js
	 */
	ApiWorksheetFunction.prototype.IFERROR = function (arg1, arg2) {
		return this.private_calculateFunction("IFERROR", arguments);
	};
	/**
	 * Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1 Is any value or expression or reference.
	 * @param {any} arg2 Is any value or expression or reference.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/IFNA.js
	 */
	ApiWorksheetFunction.prototype.IFNA = function (arg1, arg2) {
		return this.private_calculateFunction("IFNA", arguments);
	};
	/**
	 * Changes FALSE to TRUE, or TRUE to FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {boolean} arg1 Is a value or expression that can be evaluated to TRUE or FALSE.
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/NOT.js
	 */
	ApiWorksheetFunction.prototype.NOT = function (arg1) {
		return this.private_calculateFunction("NOT", arguments);
	};
	/**
	 * Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/OR.js
	 */
	ApiWorksheetFunction.prototype.OR = function () {
		return this.private_calculateFunction("OR", arguments);
	};
	/**
	 * Returns the logical value TRUE.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/TRUE.js
	 */
	ApiWorksheetFunction.prototype.TRUE = function () {
		return this.private_calculateFunction("TRUE", arguments);
	};
	/**
	 * Returns a logical 'Exclusive Or' of all arguments.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheetFunction/Methods/XOR.js
	 */
	ApiWorksheetFunction.prototype.XOR = function () {
		return this.private_calculateFunction("XOR", arguments);
	};



	/**
	 * Returns an object that represents the range of the specified sheet using the maximum and minimum row/column coordinates.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {Range} range - The internal Range class (not a ApiRange). For more details see any new ApiRange.
	 * @param {Range[]} areas - A collection of the ranges (not a ApiRange) from the specified range. For more details see any new ApiRange.
	 * @returns {ApiRange}
	 */
	Api.prototype.private_GetRange = function (range, areas) {
		return new ApiRange(range, areas);
	};

	/**
	 * Returns the mail merge fields.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {number} nSheet - The sheet index.
	 * @returns {string[]}
	 */
	Api.prototype.private_GetMailMergeFields = function (nSheet) {
		var oSheet = this.GetSheet(nSheet);
		var arrFields = [];
		var colIndex = 0;
		var colsCount = 0;
		var oRange = oSheet.GetRangeByNumber(1, colIndex);
		var fieldValue = undefined;

		while (oRange.GetValue() !== "") {
			colsCount++;
			colIndex++;
			oRange = oSheet.GetRangeByNumber(1, colIndex);
		}

		for (var nCol = 0; nCol < colsCount; nCol++) {
			oRange = oSheet.GetRangeByNumber(0, nCol);
			fieldValue = oRange.GetValue();

			if (fieldValue !== "")
				arrFields.push(oRange.GetValue());
			else
				arrFields.push("F" + String(nCol + 1));
		}

		return arrFields;
	};

	/**
	 * Returns the mail merge map.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {number} nSheet - The sheet index.
	 * @param {boolean} [bWithFormat=false] - Specifies that the data will be received with the format.
	 * @returns {string[][]}
	 */
	Api.prototype.private_GetMailMergeMap = function (nSheet, bWithFormat) {
		var oSheet = this.GetSheet(nSheet);
		var arrMailMergeMap = [];
		var valuesInRow = null;

		var rowIndex = 1;
		var rowsCount = 0;
		var colIndex = 0;
		var colsCount = 0;

		var mergeValue = undefined;

		var oRange = oSheet.GetRangeByNumber(rowIndex, 0);

		// определяем количество строк с данными
		while (oRange.GetValue() !== "") {
			rowsCount++;
			rowIndex++;
			oRange = oSheet.GetRangeByNumber(rowIndex, 0);
		}

		oRange = oSheet.GetRangeByNumber(1, colIndex);
		// определяем количество столбцов с данными
		while (oRange.GetValue() !== "") {
			colsCount++;
			colIndex++;
			oRange = oSheet.GetRangeByNumber(1, colIndex);
		}

		for (var nRow = 1; nRow < rowsCount + 1; nRow++) {
			valuesInRow = [];

			for (var nCol = 0; nCol < colsCount; nCol++) {
				oRange = oSheet.GetRangeByNumber(nRow, nCol);
				mergeValue = bWithFormat ? oRange.GetText() : oRange.GetValue();

				valuesInRow.push(mergeValue);
			}

			arrMailMergeMap.push(valuesInRow);
		}


		return arrMailMergeMap;
	};

	/**
	 * Returns the mail merge data.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {number} nSheet - The sheet index.
	 * @param {boolean} [bWithFormat=false] - Specifies that the data will be received with the format.
	 * @returns {string[][]}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetMailMergeData.js
	 */
	Api.prototype.GetMailMergeData = function (nSheet, bWithFormat) {
		if (bWithFormat !== true)
			bWithFormat = false;

		var arrFields = this.private_GetMailMergeFields(nSheet);
		var arrMailMergeMap = this.private_GetMailMergeMap(nSheet, arrFields, bWithFormat);
		var resultList = [arrFields];

		for (var nMailMergeMap = 0; nMailMergeMap < arrMailMergeMap.length; nMailMergeMap++) {
			resultList.push(arrMailMergeMap[nMailMergeMap]);
		}

		return resultList;
	};

	/**
	 * Recalculates all formulas in the active workbook.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {Function} fLogger - A function which specifies the logger object for checking recalculation of formulas.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/RecalculateAllFormulas.js
	 */
	Api.prototype.RecalculateAllFormulas = function (fLogger) {
		var formulas = this.wbModel.getAllFormulas(true);
		var _compare = function (_val1, _val2) {
			if (!isNaN(parseFloat(_val1)) && isFinite(_val1) && !isNaN(parseFloat(_val2)) && isFinite(_val2)) {
				var eps = 1e-12;
				if (Math.abs(_val2 - _val1) < eps) {
					return true;
				}

				var _slice = function (_val) {
					var sVal = _val.toString();
					if (sVal) {
						var aVal1 = sVal.split(".");
						if (aVal1[1]) {
							aVal1[1] = aVal1[1].slice(0, 9);
							sVal = aVal1[0] + "." + aVal1[1];
						}
						sVal = sVal.slice(0, 14);
						_val = parseFloat(sVal);
					}
					return _val;
				};

				_val1 = _slice(_val1);
				_val2 = _slice(_val2);
			} else {
				if (_val1 && _val2) {

					var complexVal1 = AscCommonExcel.Complex.prototype.ParseString(_val1 + "");
					if (complexVal1 && complexVal1.real && complexVal1.img) {
						var complexVal2 = AscCommonExcel.Complex.prototype.ParseString(_val2 + "");
						if (complexVal2 && complexVal2.real && complexVal2.img) {
							if (_compare(complexVal1.real, complexVal2.real) && _compare(complexVal1.img, complexVal2.img)) {
								return true;
							}
						}
					}
				}
			}
			return _val1 == _val2;
		};
		for (var i = 0; i < formulas.length; ++i) {
			var formula = formulas[i];
			var nRow;
			var nCol;
			if (formula.f && formula.r !== undefined && formula.c !== undefined) {
				nRow = formula.r;
				nCol = formula.c;
				formula = formula.f;
			}

			if (formula.parent) {
				nRow = formula.parent.nRow;
				nCol = formula.parent.nCol;
			}

			if (formula.parent && nRow !== undefined && nCol !== undefined) {
				var cell = formula.ws.getCell3(nRow, nCol);
				var oldValue = cell.getValue();
				let caTemp = formula.ca;
				formula.setFormula(formula.getFormula());
				formula.parse();
				formula.ca = caTemp;
				var formulaRes = formula.calculate();
				var newValue = formula.simplifyRefType(formulaRes, formula.ws, nRow, nCol);
				if (fLogger) {
					if (!_compare(oldValue, newValue)) {
						//error
						fLogger({
							sheet: formula.ws.sName,
							r: formula.parent.nRow,
							c: formula.parent.nCol,
							f: formula.Formula,
							oldValue: oldValue,
							newValue: newValue
						});
					}
				}
			}
		}
	};

	/**
	 * Inserts PivotTable on existing worksheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} dataRef - Source data range.
	 * @param {ApiRange} pivotRef - Range in which the table will be located.
	 * @param {bool} confirmation - Replace the data in pivotRef (if it exists) or create a dialog box for this (if it exists).
	 * @returns {ApiPivotTable}
	 * @see office-js-api/Examples/Cell/Api/Methods/InsertPivotExistingWorksheet.js
	 */
	Api.prototype.InsertPivotExistingWorksheet = function (dataRef, pivotRef, confirmation) {
		if (dataRef) {
			dataRef = dataRef.GetWorksheet().GetName() + "!" + dataRef.GetAddress(true, true);
		} else {
			var options = this.asc_getAddPivotTableOptions();
			dataRef = options.range;
		}
		if (pivotRef) {
			pivotRef = pivotRef.GetWorksheet().GetName() + "!" + pivotRef.GetAddress(true, true);
		} else {
			private_MakeError('"pivotRef" is undefined.');
		}
		var pivot = this.asc_insertPivotExistingWorksheet(dataRef, pivotRef, confirmation);
		if (pivot) {
			return new ApiPivotTable(pivot, this);
		}
		private_MakeError('Error! Bad pivotRef!');
	};

	/**
	 * Inserts PivotTable on new worksheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} dataRef - Source data range.
	 * @param {ApiRange} [newSheetName] - New worksheet name.
	 * @returns {ApiPivotTable}
	 * @see office-js-api/Examples/Cell/Api/Methods/InsertPivotNewWorksheet.js
	 */
	Api.prototype.InsertPivotNewWorksheet = function (dataRef, newSheetName) {
		if (dataRef) {
			dataRef = dataRef.GetWorksheet().GetName() + "!" + dataRef.GetAddress(true, true);
		} else {
			var options = this.asc_getAddPivotTableOptions();
			dataRef = options.range;
		}
		if (!newSheetName) {
			var items = [], wc = this.asc_getWorksheetsCount();
			while (wc--) {
				items.push(this.asc_getWorksheetName(wc).toLowerCase());
			}

			var index = 0, name;
			while(++index < 1000) {
				name = 'Sheet' + index;
				if (items.indexOf(name.toLowerCase()) < 0) break;
			}

			newSheetName = name;
		}
		var pivot = this.asc_insertPivotNewWorksheet(dataRef, newSheetName);
		if (pivot) {
			return new ApiPivotTable(pivot, this);
		}
		private_MakeError('An error occurred while creating the pivot table!');
	};

	/**
	 * Returns PivotTable by name or null.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name of PivotTable.
	 * @returns {ApiPivotTable|null}
	 * @see office-js-api/Examples/Cell/Api/Methods/GetPivotByName.js
	 */
	Api.prototype.GetPivotByName = function (name) {
		var res = null;
		if (typeof name == "string" && name.trim().length) {
			var pivot = this.wbModel.getPivotTableByName( name.trim() );
			if (pivot)
				res = new ApiPivotTable(pivot, this);
		}
		return res;
	};

	/**
	 * Refreshs all PivotTables.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/Cell/Api/Methods/RefreshAllPivots.js
	 */
	Api.prototype.RefreshAllPivots = function () {
		this.asc_refreshAllPivots();
	};

	/**
	 * Returns all PivotTables.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotTable[]}
	 * @see office-js-api/Examples/Cell/Api/Methods/GetAllPivotTables.js
	 */
	Api.prototype.GetAllPivotTables = function () {
		var res = [];
		var sheets = this.GetSheets();
		sheets.forEach(function(ws) {
			res = res.concat( ws.GetAllPivotTables() )
		});
		return res;
	};

	Object.defineProperty(Api.prototype, "PivotTables", {
		get: function () {
			return this.GetAllPivotTables();
		}
	});

	/**
	 * Subscribes to the specified event and calls the callback function when the event fires.
	 * @function
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} eventName - The event name.
	 * @param {function} callback - Function to be called when the event fires.
	 * @fires Api#onWorksheetChange
	 * @see office-js-api/Examples/{Editor}/Api/Methods/attachEvent.js
	 */
	Api.prototype["attachEvent"] = Api.prototype.attachEvent;

	/**
	 * Unsubscribes from the specified event.
	 * @function
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} eventName - The event name.
	 * @fires Api#onWorksheetChange
	 * @see office-js-api/Examples/{Editor}/Api/Methods/detachEvent.js
	 */
	Api.prototype["detachEvent"] = Api.prototype.detachEvent;

	/**
	 * Returns an array of ApiComment objects.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sText - The comment text.
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment | null}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/Api/Methods/AddComment.js
	 */
	Api.prototype.AddComment = function (sText, sAuthor) {
		let result = null;
		let isValidData = typeof (sText) === 'string' && sText.trim() !== '';
		if (isValidData) {
			var comment = new Asc.asc_CCommentData();
			comment.asc_putText(sText);
			let author = ((typeof (sAuthor) === 'string' && sAuthor.trim() !== '') ? sAuthor : Asc['editor'].User.asc_getUserName());
			comment.asc_putUserName(author);
			// todo проверить как в документа добавлются (надо ли выставлять этот параметр)
			// comment.asc_putUserId(Asc['editor'].User.asc_getId());
			comment.asc_putDocumentFlag(true);
			this.asc_addComment(comment);
			result = new ApiComment(comment, Asc['editor'].wb);
		}

		return result;
	};

	/**
	 * Returns a comment from the current document by its ID.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sId - The comment ID.
	 * @returns {?ApiComment}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetCommentById.js
	 */
	Api.prototype.GetCommentById = function (sId) {
		let comment = this.asc_findComment(sId);
		if (!comment)
			comment = this.wb.cellCommentator.findComment(sId);

		return comment ? new ApiComment(comment, Asc['editor'].wb) : null;
	};

	/**
	 * Returns all comments related to the whole workbook.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiComment[]}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetComments.js
	 */
	Api.prototype.GetComments = function () {
		var comments = [];
		for (var i = 0; i < this.wbModel.aComments.length; i++) {
			comments.push(new ApiComment(this.wbModel.aComments[i], this.wb));
		}
		return comments;
	};
	Object.defineProperty(Api.prototype, "Comments", {
		get: function () {
			return this.GetComments();
		}
	});


	/**
	 * Returns all comments from the current workbook including comments from all worksheets.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiComment[]}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetAllComments.js
	 */
	Api.prototype.GetAllComments = function () {
		let aApiComments = this.GetComments();

		let aWS = this.GetSheets();
		for(let nWS = 0; nWS < aWS.length; ++nWS) {
			aApiComments = aApiComments.concat(aWS[nWS].GetComments())
		}
		return aApiComments;
	};
	Object.defineProperty(Api.prototype, "AllComments", {
		get: function () {
			return this.GetAllComments();
		}
	});

	/**
	 * Specifies a type of freeze panes.
	 * @typedef {("row" | "column" | "cell" | null )} FreezePaneType
	 * @see office-js-api/Examples/Enumerations/FreezePaneType.js
	 */

	/**
	 * Sets a type to the freeze panes.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {FreezePaneType} FreezePaneType - The freeze panes type ("null" to unfreeze).
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/Api/Methods/SetFreezePanesType.js
	 */
	Api.prototype.SetFreezePanesType = function (FreezePaneType) {
		if (typeof FreezePaneType === 'string' || FreezePaneType === null) {
			//detect current freeze type
			let curType = this.GetFreezePanesType();

			let type = null;
			if (FreezePaneType === 'cell' && ((curType && curType !== 'cell') || (!curType))) {
				// make unfreeze and freeze then
				if (curType)
					this.asc_freezePane(undefined);

				type = undefined;
			} else if (FreezePaneType === null && curType) {
				type = undefined;
			} else if (FreezePaneType === 'row' && curType !== 'row') {
				type = 1;
			} else if (FreezePaneType === 'column' && curType !== 'column') {
				type = 2;
			}

			if (type !== null)
				this.asc_freezePane(type);

		} else {
			logError(new Error('Invalid parameter "FreezePaneType".'));
		}
	};

	/**
	 * Returns the freeze panes type.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {FreezePaneType} FreezePaneType - The freeze panes type ("null" if there are no freeze panes).
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetFreezePanesType.js
	 */
	Api.prototype.GetFreezePanesType = function () {
		let cell = this.wb.getWorksheet().topLeftFrozenCell;
		//detect current freeze type
		let curType = null;
		if (cell) {
			let c = cell.getCol0();
			let r = cell.getRow0();
			if (c == 0) {
				// hole row
				curType = 'row';
			} else if (r == 0) {
				// whole column
				curType = 'column';
			} else {
				// cell
				curType = 'cell';
			}
		}
		return curType;
	};

	Object.defineProperty(Api.prototype, "FreezePanes", {
		get: function () {
			return this.GetFreezePanesType();
		},
		set: function (FreezePaneType) {
			this.SetFreezePanesType(FreezePaneType);
		}
	});

	/**
	 * Returns the cell reference style.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ReferenceStyle} - The cell reference style.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetReferenceStyle.js
	 */
	Api.prototype.GetReferenceStyle = function () {
		let bReferenceStyle = this.asc_getR1C1Mode();
		return bReferenceStyle ? "xlR1C1" : "xlA1";
	};

	/**
	 * Sets the cell reference style.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {ReferenceStyle} sReferenceStyle - The cell reference style.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/Api/Methods/SetReferenceStyle.js
	 */
	Api.prototype.SetReferenceStyle = function (sReferenceStyle) {
		let bReferenceMode = null;
		switch (sReferenceStyle) {
			case "xlA1":
				bReferenceMode = false;
				break;
			case "xlR1C1":
				bReferenceMode = true;
				break;
		}

		if (bReferenceMode !== null) {
			this.asc_setR1C1Mode(bReferenceMode);
		} else {
			logError(new Error('Invalid parameter "ReferenceStyle"'));
		}
	};

	Object.defineProperty(Api.prototype, "ReferenceStyle", {
		get: function () {
			return this.GetReferenceStyle();
		},
		set: function (ReferenceStyle) {
			this.SetReferenceStyle(ReferenceStyle);
		}
	});


	/**
	 * Returns the document information:
	 * <b>Application</b> - the application the document has been created with.
	 * <b>CreatedRaw</b> - the date and time when the file was created.
	 * <b>Created</b> - the parsed date and time when the file was created.
	 * <b>LastModifiedRaw</b> - the date and time when the file was last modified.
	 * <b>LastModified</b> - the parsed date and time when the file was last modified.
	 * <b>LastModifiedBy</b> - the name of the user who has made the latest change to the document.
	 * <b>Autrors</b> - the persons who has created the file.
	 * <b>Title</b> - this property allows you to simplify your documents classification.
	 * <b>Tags</b> - this property allows you to simplify your documents classification.
	 * <b>Subject</b> - this property allows you to simplify your documents classification.
	 * <b>Comment</b> - this property allows you to simplify your documents classification.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {object}
	 * @see office-js-api/Examples/{Editor}/Api/Methods/GetDocumentInfo.js
	 */
	Api.prototype.GetDocumentInfo = function()
	{
		const oDocInfo = {
			Application: '',
			CreatedRaw: null,
			Created: '',
			LastModifiedRaw: null,
			LastModified: '',
			LastModifiedBy: '',
			Autrors: [],
			Title: '',
			Tags: '',
			Subject: '',
			Comment: ''
		};

		let props = (this) ? this.asc_getAppProps() : null;
		oDocInfo.Application = (props.asc_getApplication() || '') + (props.asc_getAppVersion() ? ' ' : '') + (props.asc_getAppVersion() || '');

		let langCode = 1033; // en-US
		let langName = 'en-us';
		if (AscCommon.g_oDefaultCultureInfo.Name) {
			langName = AscCommon.g_oDefaultCultureInfo.Name.replace('_', '-').toLowerCase();
		} else if (this.defaultLanguage && window['Common']) {
			langCode = this.defaultLanguage;
			langName = window['Common']['util']['LanguageInfo']['getLocalLanguageName'](langCode)[0].toLowerCase();

		}

		props = this.asc_getCoreProps();
		oDocInfo.CreatedRaw = props.asc_getCreated();
		oDocInfo.LastModifiedRaw = props.asc_getModified();

		try {
			if (oDocInfo.CreatedRaw)
				oDocInfo.Created = (oDocInfo.CreatedRaw.toLocaleString(langName, {year: 'numeric', month: '2-digit', day: '2-digit'}) + ' ' +oDocInfo. CreatedRaw.toLocaleString(langName, {timeStyle: 'short'}));
			
			if (oDocInfo.LastModifiedRaw)
				oDocInfo.LastModified = (oDocInfo.LastModifiedRaw.toLocaleString(langName, {year: 'numeric', month: '2-digit', day: '2-digit'}) + ' ' + oDocInfo.LastModifiedRaw.toLocaleString(langName, {timeStyle: 'short'}));
		} catch (e) {
			langName = 'en';
			if (oDocInfo.CreatedRaw)
				oDocInfo.Created = (oDocInfo.CreatedRaw.toLocaleString(langName, {year: 'numeric', month: '2-digit', day: '2-digit'}) + ' ' + oDocInfo.CreatedRaw.toLocaleString(langName, {timeStyle: 'short'}));

			if (oDocInfo.LastModifiedRaw)
				oDocInfo.LastModified = (oDocInfo.LastModifiedRaw.toLocaleString(langName, {year: 'numeric', month: '2-digit', day: '2-digit'}) + ' ' + oDocInfo.LastModifiedRaw.toLocaleString(langName, {timeStyle: 'short'}));
		}

		const LastModifiedBy = props.asc_getLastModifiedBy();
		oDocInfo.LastModifiedBy = AscCommon.UserInfoParser.getParsedName(LastModifiedBy);

		oDocInfo.Title = (props.asc_getTitle() || '');
		oDocInfo.Tags = (props.asc_getKeywords() || '');
		oDocInfo.Subject = (props.asc_getSubject() || '');
		oDocInfo.Comment = (props.asc_getDescription() || '');

		const authors = props.asc_getCreator();
		if (authors)
			oDocInfo.Autrors = authors.split(/\s*[,;]\s*/);

		return oDocInfo;
	};
	/**
	 * Returns the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetVisible.js
	 */
	ApiWorksheet.prototype.GetVisible = function () {
		return !this.worksheet.getHidden();
	};

	/**
	 * Sets the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isVisible - Specifies if the sheet is visible or not.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetVisible.js
	 */
	ApiWorksheet.prototype.SetVisible = function (isVisible) {
		this.worksheet.setHidden(!isVisible);
	};
	Object.defineProperty(ApiWorksheet.prototype, "Visible", {
		get: function () {
			return this.GetVisible();
		},
		set: function (isVisible) {
			this.SetVisible(isVisible);
		}
	});

	/**
	 * Makes the current sheet active.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetActive.js
	 */
	ApiWorksheet.prototype.SetActive = function () {
		this.worksheet.workbook.setActive(this.worksheet.index);
	};
	Object.defineProperty(ApiWorksheet.prototype, "Active", {
		set: function () {
			this.SetActive();
		}
	});

	/**
	 * Returns an object that represents an active cell.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetActiveCell.js
	 */
	ApiWorksheet.prototype.GetActiveCell = function () {
		let cell = this.worksheet.getCell3(this.worksheet.selectionRange.activeCell.row, this.worksheet.selectionRange.activeCell.col);
		let merged = cell.hasMerged();
		if (merged)
			cell = this.worksheet.getCell3(merged.r1, merged.c1);

		return new ApiRange(cell);
	};
	Object.defineProperty(ApiWorksheet.prototype, "ActiveCell", {
		get: function () {
			return this.GetActiveCell();
		}
	});

	/**
	 * Returns an object that represents the selected range.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetSelection.js
	 */
	ApiWorksheet.prototype.GetSelection = function () {
		var r = this.worksheet.selectionRange.getLast();
		var ranges = this.worksheet.selectionRange.ranges;
		var arr = [];
		for (var i = 0; i < ranges.length; i++) {
			arr.push(this.worksheet.getRange3(ranges[i].r1, ranges[i].c1, ranges[i].r2, ranges[i].c2));
		}
		return new ApiRange(this.worksheet.getRange3(r.r1, r.c1, r.r2, r.c2), arr);
	};
	Object.defineProperty(ApiWorksheet.prototype, "Selection", {
		get: function () {
			return this.GetSelection();
		}
	});

	/**
	 * Returns the ApiRange that represents all the cells on the worksheet (not just the cells that are currently in use).
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} row - The row number or the cell number (if only row is defined).
	 * @param {number} col - The column number.
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetCells.js
	 */
	ApiWorksheet.prototype.GetCells = function (row, col) {
		let result;
		if (typeof col == "number" && typeof row == "number") {
			if (col < 1 || row < 1 || col > AscCommon.gc_nMaxCol0 || row > AscCommon.gc_nMaxRow0) {
				logError(new Error('Invalid paremert "row" or "col".'));
				result = null;
			} else {
				row--;
				col--;
				result = new ApiRange(this.worksheet.getRange3(row, col, row, col));
			}
		} else if (typeof row == "number") {
			if (row < 1 || row > AscCommon.gc_nMaxRow0) {
				logError(new Error('Invalid paremert "row".'));
				result = null;
			} else {
				row--;
				let r = (row) ? (row / AscCommon.gc_nMaxCol0) >> 0 : row;
				let c = (row) ? row % AscCommon.gc_nMaxCol0 : row;
				if (r && c) c--;
				result = new ApiRange(this.worksheet.getRange3(r, c, r, c));
			}

		} else if (typeof col == "number") {
			if (col < 1 || col > AscCommon.gc_nMaxCol0) {
				logError(new Error('Invalid paremert "col".'));
				result = null;
			} else {
				col--;
				result = new ApiRange(this.worksheet.getRange3(0, col, 0, col));
			}
		} else {
			result = new ApiRange(this.worksheet.getRange3(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0));
		}

		return result;
	};
	Object.defineProperty(ApiWorksheet.prototype, "Cells", {
		get: function () {
			return this.GetCells();
		}
	});
	Object.defineProperty(ApiWorksheet.prototype, "Rows", {
		get: function () {
			return this.GetCells();
		}
	});

	/**
	 * Returns the ApiRange object that represents all the cells on the rows range.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string | number} value - Specifies the rows range in the string or number format.
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetRows.js
	 */
	ApiWorksheet.prototype.GetRows = function (value) {
		if (typeof value === "undefined") {
			return this.GetCells();
		} else if (typeof value == "number" || value.indexOf(':') == -1) {
			value = parseInt(value);
			if (value > 0 && value <= AscCommon.gc_nMaxRow0 + 1 && value[0] !== NaN) {
				value--;
			} else {
				logError(new Error('The nRow must be greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1)));
				return null;
			}
			return new ApiRange(this.worksheet.getRange3(value, 0, value, AscCommon.gc_nMaxCol0));
		} else {
			value = value.split(':');
			var isError = false;
			for (var i = 0; i < value.length; ++i) {
				value[i] = parseInt(value[i]);
				if (value[i] > 0 && value[i] <= AscCommon.gc_nMaxRow0 + 1 && value[0] !== NaN) {
					value[i]--;
				} else {
					isError = true;
				}
			}
			if (isError) {
				logError(new Error('The nRow must be greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1)));
				return null;
			} else {
				return new ApiRange(this.worksheet.getRange3(value[0], 0, value[1], AscCommon.gc_nMaxCol0));
			}
		}
	};

	/**
	 * Returns the ApiRange object that represents all the cells on the columns range.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - Specifies the columns range in the string format.
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetCols.js
	 */
	ApiWorksheet.prototype.GetCols = function (sRange) {
		if (sRange.indexOf(':') == -1) {
			sRange += ':' + sRange;
		}
		return new ApiRange(this.worksheet.getRange2(sRange));
	};
	Object.defineProperty(ApiWorksheet.prototype, "Cols", {
		get: function () {
			return this.GetCells();
		}
	});

	/**
	 * Returns the ApiRange object that represents the used range on the specified worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetUsedRange.js
	 */
	ApiWorksheet.prototype.GetUsedRange = function () {
		var rEnd = this.worksheet.getRowsCount() - 1;
		var cEnd = this.worksheet.getColsCount() - 1;
		return new ApiRange(this.worksheet.getRange3(0, 0, (rEnd < 0) ? 0 : rEnd,
			(cEnd < 0) ? 0 : cEnd));
	};
	Object.defineProperty(ApiWorksheet.prototype, "UsedRange", {
		get: function () {
			return this.GetUsedRange();
		}
	});

	/**
	 * Returns a sheet name.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetName.js
	 */
	ApiWorksheet.prototype.GetName = function () {
		return this.worksheet.getName();
	};

	/**
	 * Sets a name to the current active sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The name which will be displayed for the current sheet at the sheet tab.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetName.js
	 */
	ApiWorksheet.prototype.SetName = function (sName) {
		let sOldName = this.worksheet.getName();
		this.worksheet.setName(sName);
		// let oWorkbookView = this.worksheet.workbook.oApi.wb;
		// it's temporary solution (we should use oWorkbookView instead of oWorkbook)
		let oWorkbook = this.worksheet.workbook;
		oWorkbook.oApi.sheetsChanged();
		if (oWorkbook)
			oWorkbook.handleChartsOnChangeSheetName(this.worksheet, sOldName, sName);
	};
	Object.defineProperty(ApiWorksheet.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (sName) {
			this.SetName(sName);
		}
	});

	/**
	 * Returns a sheet index.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetIndex.js
	 */
	ApiWorksheet.prototype.GetIndex = function () {
		return this.worksheet.getIndex();
	};
	Object.defineProperty(ApiWorksheet.prototype, "Index", {
		get: function () {
			return this.GetIndex();
		}
	});

	/**
	 * Returns an object that represents the selected range of the current sheet. Can be a single cell - <b>A1</b>, or cells
	 * from a single row - <b>A1:E1</b>, or cells from a single column - <b>A1:A10</b>, or cells from several rows and columns - <b>A1:E10</b>.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string | ApiRange} Range1 - The range of cells from the current sheet.
	 * @param {string | ApiRange} Range2 - The range of cells from the current sheet.
	 * @returns {ApiRange | null} - returns null if such a range does not exist.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetRange.js
	 */
	ApiWorksheet.prototype.GetRange = function (Range1, Range2) {
		var Range, r1, c1, r2, c2;
		Range1 = (Range1 instanceof ApiRange) ? Range1.range : (typeof Range1 == 'string') ? this.worksheet.getRange2(Range1) : null;

		if (!Range1) {
			logError(new Error('Incorrect "Range1" or it is empty.'));
			return null;
		}

		Range2 = (Range2 instanceof ApiRange) ? Range2.range : (typeof Range2 == 'string') ? this.worksheet.getRange2(Range2) : null;

		if (Range2) {
			r1 = Math.min(Range1.bbox.r1, Range2.bbox.r1);
			c1 = Math.min(Range1.bbox.c1, Range2.bbox.c1);
			r2 = Math.max(Range1.bbox.r2, Range2.bbox.r2);
			c2 = Math.max(Range1.bbox.c1, Range2.bbox.c2);
		} else {
			r1 = Range1.bbox.r1;
			c1 = Range1.bbox.c1;
			r2 = Range1.bbox.r2;
			c2 = Range1.bbox.c2;
		}

		Range = this.worksheet.getRange3(r1, c1, r2, c2);

		if (!Range)
			return null;

		return new ApiRange(Range);
	};

	/**
	 * Returns an object that represents the selected range of the current sheet using the <b>row/column</b> coordinates for the cell selection.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nRow - The row number.
	 * @param {number} nCol - The column number.
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetRangeByNumber.js
	 */
	ApiWorksheet.prototype.GetRangeByNumber = function (nRow, nCol) {
		return new ApiRange(this.worksheet.getCell3(nRow, nCol));
	};

	/**
	 * Formats the selected range of cells from the current sheet as a table (with the first row formatted as a header).
	 * <note>As the first row is always formatted as a table header, you need to select at least two rows for the table to be formed correctly.</note>
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - The range of cells from the current sheet which will be formatted as a table.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/FormatAsTable.js
	 */
	ApiWorksheet.prototype.FormatAsTable = function (sRange) {
		this.worksheet.autoFilters.addAutoFilter('TableStyleLight9', AscCommonExcel.g_oRangeCache.getAscRange(sRange));
	};

	/**
	 * Sets the width of the specified column.
	 * One unit of column width is equal to the width of one character in the Normal style.
	 * For proportional fonts, the width of the character 0 (zero) is used.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nColumn - The number of the column to set the width to.
	 * @param {number} nWidth - The width of the column divided by 7 pixels.
	 * @param {boolean} [bWithotPaddings=false] - Specifies whether nWidth will be set without standard paddings.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetColumnWidth.js
	 */
	ApiWorksheet.prototype.SetColumnWidth = function (nColumn, nWidth, bWithotPaddings) {
		if (bWithotPaddings) {
			let wb = this.worksheet.workbook;
			nWidth = (nWidth * wb.maxDigitWidth - wb.paddingPlusBorder) / wb.maxDigitWidth;
		}
		this.worksheet.setColWidth(nWidth, nColumn, nColumn);
	};

	/**
	 * Sets the height of the specified row measured in points.
	 * A point is 1/72 inch.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nRow - The number of the row to set the height to.
	 * @param {number} nHeight - The height of the row measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetRowHeight.js
	 */
	ApiWorksheet.prototype.SetRowHeight = function (nRow, nHeight) {
		this.worksheet.setRowHeight(nHeight, nRow, nRow, true);
	};

	/**
	 * Specifies whether the current sheet gridlines must be displayed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isDisplayed - Specifies whether the current sheet gridlines must be displayed or not. The default value is <b>true</b>.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetDisplayGridlines.js
	 */
	ApiWorksheet.prototype.SetDisplayGridlines = function (isDisplayed) {
		this.worksheet.setDisplayGridlines(!!isDisplayed);
	};

	/**
	 * Specifies whether the current sheet row/column headers must be displayed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isDisplayed - Specifies whether the current sheet row/column headers must be displayed or not. The default value is <b>true</b>.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetDisplayHeadings.js
	 */
	ApiWorksheet.prototype.SetDisplayHeadings = function (isDisplayed) {
		this.worksheet.setDisplayHeadings(!!isDisplayed);
	};

	/**
	 * Sets the left margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nPoints - The left margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetLeftMargin.js
	 */
	ApiWorksheet.prototype.SetLeftMargin = function (nPoints) {
		nPoints = (typeof nPoints !== 'number') ? 0 : nPoints;
		this.worksheet.PagePrintOptions.pageMargins.asc_setLeft(nPoints);
	};
	/**
	 * Returns the left margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {number} - The left margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetLeftMargin.js
	 */
	ApiWorksheet.prototype.GetLeftMargin = function () {
		return this.worksheet.PagePrintOptions.pageMargins.asc_getLeft();
	};
	Object.defineProperty(ApiWorksheet.prototype, "LeftMargin", {
		get: function () {
			return this.GetLeftMargin();
		},
		set: function (nPoints) {
			this.SetLeftMargin(nPoints);
		}
	});

	/**
	 * Sets the right margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nPoints - The right margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetRightMargin.js
	 */
	ApiWorksheet.prototype.SetRightMargin = function (nPoints) {
		nPoints = (typeof nPoints !== 'number') ? 0 : nPoints;
		this.worksheet.PagePrintOptions.pageMargins.asc_setRight(nPoints);
	};
	/**
	 * Returns the right margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {number} - The right margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetRightMargin.js
	 */
	ApiWorksheet.prototype.GetRightMargin = function () {
		return this.worksheet.PagePrintOptions.pageMargins.asc_getRight();
	};
	Object.defineProperty(ApiWorksheet.prototype, "RightMargin", {
		get: function () {
			return this.GetRightMargin();
		},
		set: function (nPoints) {
			this.SetRightMargin(nPoints);
		}
	});

	/**
	 * Sets the top margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nPoints - The top margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetTopMargin.js
	 */
	ApiWorksheet.prototype.SetTopMargin = function (nPoints) {
		nPoints = (typeof nPoints !== 'number') ? 0 : nPoints;
		this.worksheet.PagePrintOptions.pageMargins.asc_setTop(nPoints);
	};
	/**
	 * Returns the top margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {number} - The top margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetTopMargin.js
	 */
	ApiWorksheet.prototype.GetTopMargin = function () {
		return this.worksheet.PagePrintOptions.pageMargins.asc_getTop();
	};
	Object.defineProperty(ApiWorksheet.prototype, "TopMargin", {
		get: function () {
			return this.GetTopMargin();
		},
		set: function (nPoints) {
			this.SetTopMargin(nPoints);
		}
	});

	/**
	 * Sets the bottom margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nPoints - The bottom margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetBottomMargin.js
	 */
	ApiWorksheet.prototype.SetBottomMargin = function (nPoints) {
		nPoints = (typeof nPoints !== 'number') ? 0 : nPoints;
		this.worksheet.PagePrintOptions.pageMargins.asc_setBottom(nPoints);
	};
	/**
	 * Returns the bottom margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {number} - The bottom margin size measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetBottomMargin.js
	 */
	ApiWorksheet.prototype.GetBottomMargin = function () {
		return this.worksheet.PagePrintOptions.pageMargins.asc_getBottom();
	};
	Object.defineProperty(ApiWorksheet.prototype, "BottomMargin", {
		get: function () {
			return this.GetBottomMargin();
		},
		set: function (nPoints) {
			this.SetBottomMargin(nPoints);
		}
	});

	/**
	 * Sets the page orientation.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {PageOrientation} sPageOrientation - The page orientation type.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetPageOrientation.js
	 */
	ApiWorksheet.prototype.SetPageOrientation = function (sPageOrientation) {
		this.worksheet.PagePrintOptions.pageSetup.asc_setOrientation('xlLandscape' === sPageOrientation ? 1 : 0);
	};

	/**
	 * Returns the page orientation.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {PageOrientation}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetPageOrientation.js
	 */
	ApiWorksheet.prototype.GetPageOrientation = function () {
		var PageOrientation = this.worksheet.PagePrintOptions.pageSetup.asc_getOrientation();
		return (PageOrientation) ? 'xlLandscape' : 'xlPortrait';
	};

	Object.defineProperty(ApiWorksheet.prototype, "PageOrientation", {
		get: function () {
			return this.GetPageOrientation();
		},
		set: function (sPageOrientation) {
			this.SetPageOrientation(sPageOrientation);
		}
	});


	/**
	 * Returns the page PrintHeadings property which specifies whether the current sheet row/column headings must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {boolean} - Specifies whether the current sheet row/column headings must be printed or not.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetPrintHeadings.js
	 */
	ApiWorksheet.prototype.GetPrintHeadings = function () {
		return this.worksheet.PagePrintOptions.asc_getHeadings();
	};

	/**
	 * Specifies whether the current sheet row/column headers must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bPrint - Specifies whether the current sheet row/column headers must be printed or not.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetPrintHeadings.js
	 */
	ApiWorksheet.prototype.SetPrintHeadings = function (bPrint) {
		this.worksheet.PagePrintOptions.asc_setHeadings(!!bPrint);
	};

	Object.defineProperty(ApiWorksheet.prototype, "PrintHeadings", {
		get: function () {
			return this.GetPrintHeadings();
		},
		set: function (bPrint) {
			this.SetPrintHeadings(bPrint)
		}
	});

	/**
	 * Returns the page PrintGridlines property which specifies whether the current sheet gridlines must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {boolean} - True if cell gridlines are printed on this page.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetPrintGridlines.js
	 */
	ApiWorksheet.prototype.GetPrintGridlines = function () {
		return this.worksheet.PagePrintOptions.asc_getGridLines();
	};

	/**
	 * Specifies whether the current sheet gridlines must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bPrint - Defines if cell gridlines are printed on this page or not.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetPrintGridlines.js
	 */
	ApiWorksheet.prototype.SetPrintGridlines = function (bPrint) {
		this.worksheet.PagePrintOptions.asc_setGridLines(!!bPrint);
	};

	Object.defineProperty(ApiWorksheet.prototype, "PrintGridlines", {
		get: function () {
			return this.GetPrintGridlines();
		},
		set: function (bPrint) {
			this.SetPrintGridlines(bPrint)
		}
	});

	/**
	 * Returns an array of ApiName objects.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiName[]}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetDefNames.js
	 */
	ApiWorksheet.prototype.GetDefNames = function () {
		var res = this.worksheet.workbook.getDefinedNamesWS(this.worksheet.getId());
		var name = [];
		if (!res.length) {
			return [new ApiName(undefined)]
		}
		for (var i = 0; i < res.length; i++) {
			name.push(new ApiName(res[i]));
		}
		return name;
	};

	/**
	 * Returns the ApiName object by the worksheet name.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} defName - The worksheet name.
	 * @returns {ApiName | null} - returns null if definition name doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetDefName.js
	 */
	ApiWorksheet.prototype.GetDefName = function (defName) {
		if (defName && typeof defName === "string") {
			defName = this.worksheet.workbook.getDefinesNames(defName, this.worksheet.getId());
			return new ApiName(defName);
		}

		return null;
	};

	/**
	 * Adds a new name to the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The range name.
	 * @param {string} sRef  - Must contain the sheet name, followed by sign ! and a range of cells.
	 * Example: "Sheet1!$A$1:$B$2".
	 * @param {boolean} isHidden - Defines if the range name is hidden or not.
	 * @returns {boolean} - returns false if sName or sRef are invalid.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddDefName.js
	 */
	ApiWorksheet.prototype.AddDefName = function (sName, sRef, isHidden) {
		return private_AddDefName(this.worksheet.workbook, sName, sRef, this.worksheet.getIndex(), isHidden);
	};

	Object.defineProperty(ApiWorksheet.prototype, "DefNames", {
		get: function () {
			return this.GetDefNames();
		}
	});

	/**
	 * Returns all comments from the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiComment[]}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetComments.js
	 */
	ApiWorksheet.prototype.GetComments = function () {
		var comments = [];
		for (var i = 0; i < this.worksheet.aComments.length; i++) {
			comments.push(new ApiComment(this.worksheet.aComments[i], this.worksheet.workbook.oApi.wb));
		}
		return comments;
	};
	Object.defineProperty(ApiWorksheet.prototype, "Comments", {
		get: function () {
			return this.GetComments();
		}
	});

	/**
	 * Deletes the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/Delete.js
	 */
	ApiWorksheet.prototype.Delete = function () {
		this.worksheet.workbook.oApi.asc_deleteWorksheet([this.worksheet.getIndex()]);
	};

	/**
	 * Adds a hyperlink to the specified range.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - The range where the hyperlink will be added to.
	 * @param {string} sAddress - The link address.
	 * @param {string} subAddress - The link subaddress to insert internal sheet hyperlinks.
	 * @param {string} sScreenTip - The screen tip text.
	 * @param {string} sTextToDisplay - The link text that will be displayed on the sheet.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/SetHyperlink.js
	 */
	ApiWorksheet.prototype.SetHyperlink = function (sRange, sAddress, subAddress, sScreenTip, sTextToDisplay) {
		var range = new ApiRange(this.worksheet.getRange2(sRange));
		var address;
		if (range && range.range.isOneCell() && (sAddress || subAddress)) {
			var externalLink = sAddress ? AscCommon.rx_allowedProtocols.test(sAddress) : false;
			if (externalLink && AscCommonExcel.getFullHyperlinkLength(sAddress) > Asc.c_nMaxHyperlinkLength) {
				throwException(new Error('Incorrect "sAddress".'));
			}
			if (!externalLink) {
				address = subAddress.split("!");
				if (address.length == 1)
					address.unshift(this.GetName());
				else if (this.worksheet.workbook.getWorksheetByName(address[0]) === null) {
					throwException(new Error('Invalid "subAddress".'));
				}
				var res = this.worksheet.workbook.oApi.asc_checkDataRange(Asc.c_oAscSelectionDialogType.FormatTable, address[1], false);
				if (res === Asc.c_oAscError.ID.DataRangeError) {
					throwException(new Error('Invalid "subAddress".'));
				}
			}
			this.worksheet.selectionRange.assign2(range.range.bbox);
			var Hyperlink = new Asc.asc_CHyperlink();
			if (sScreenTip) {
				Hyperlink.asc_setText(sScreenTip);
			} else {
				Hyperlink.asc_setText((externalLink ? sAddress : subAddress));
			}
			if (sTextToDisplay) {
				Hyperlink.asc_setTooltip(sTextToDisplay);
			}
			if (externalLink) {
				Hyperlink.asc_setHyperlinkUrl(sAddress);
			} else {
				Hyperlink.asc_setRange(address[1]);
				Hyperlink.asc_setSheet(address[0]);
			}
			this.worksheet.workbook.oApi.wb.insertHyperlink(Hyperlink, this.GetIndex());
		}
	};

	/**
	 * Creates a chart of the specified type from the selected data range of the current sheet.
	 * <note>Please note that the horizontal and vertical offsets are calculated within the limits of the specified column and
	 * row cells only. If this value exceeds the cell width or height, another vertical/horizontal position will be set.</note>
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sDataRange - The selected cell range which will be used to get the data for the chart, formed specifically and including the sheet name.
	 * @param {boolean} bInRows - Specifies whether to take the data from the rows or from the columns. If true, the data from the rows will be used.
	 * @param {ChartType} sType - The chart type used for the chart display.
	 * @param {number} nStyleIndex - The chart color style index (can be <b>1 - 48</b>, as described in OOXML specification).
	 * @param {EMU} nExtX - The chart width in English measure units
	 * @param {EMU} nExtY - The chart height in English measure units.
	 * @param {number} nFromCol - The number of the column where the beginning of the chart will be placed.
	 * @param {EMU} nColOffset - The offset from the nFromCol column to the left part of the chart measured in English measure units.
	 * @param {number} nFromRow - The number of the row where the beginning of the chart will be placed.
	 * @param {EMU} nRowOffset - The offset from the nFromRow row to the upper part of the chart measured in English measure units.
	 * @returns {ApiChart}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddChart.js
	 */
	ApiWorksheet.prototype.AddChart =
		function (sDataRange, bInRows, sType, nStyleIndex, nExtX, nExtY, nFromCol, nColOffset, nFromRow, nRowOffset) {
			var settings = new Asc.asc_ChartSettings();
			settings.type = AscFormat.ChartBuilderTypeToInternal(sType);
			settings.style = nStyleIndex;
			settings.inColumns = !bInRows;
			settings.putRange(sDataRange);
			var oChart = AscFormat.DrawingObjectsController.prototype.getChartSpace(settings);
			if (arguments.length === 8) {//support old variant
				oChart.setBDeleted(false);
				oChart.setWorksheet(this.worksheet);
				oChart.addToDrawingObjects();
				oChart.setDrawingBaseCoords(arguments[4], 0, arguments[5], 0, arguments[6], 0, arguments[7], 0, 0, 0, 0, 0);
			} else {
				private_SetCoords(oChart, this.worksheet, nExtX, nExtY, nFromCol, nColOffset, nFromRow, nRowOffset);
			}
			if (AscFormat.isRealNumber(nStyleIndex)) {
				oChart.setStyle(nStyleIndex);
			}
			oChart.recalculateReferences();
			return Asc.editor.private_CreateApiChart(oChart);
		};


	/**
	 * Adds a shape to the current sheet with the parameters specified.
	 * <note>Please note that the horizontal and vertical offsets are
	 * calculated within the limits of the specified column and row cells
	 * only. If this value exceeds the cell width or height, another vertical/horizontal position will be set.</note>
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {ShapeType} [sType="rect"] - The shape type which specifies the preset shape geometry.
	 * @param {EMU} nWidth - The shape width in English measure units.
	 * @param {EMU} nHeight - The shape height in English measure units.
	 * @param {ApiFill} oFill - The color or pattern used to fill the shape.
	 * @param {ApiStroke} oStroke - The stroke used to create the element shadow.
	 * @param {number} nFromCol - The number of the column where the beginning of the shape will be placed.
	 * @param {EMU} nColOffset - The offset from the nFromCol column to the left part of the shape measured in English measure units.
	 * @param {number} nFromRow - The number of the row where the beginning of the shape will be placed.
	 * @param {EMU} nRowOffset - The offset from the nFromRow row to the upper part of the shape measured in English measure units.
	 * @returns {ApiShape}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddShape.js
	 */
	ApiWorksheet.prototype.AddShape = function (sType, nWidth, nHeight, oFill, oStroke, nFromCol, nColOffset, nFromRow, nRowOffset) {
		var oShape = AscFormat.builder_CreateShape(sType, nWidth / 36000, nHeight / 36000, oFill.UniFill, oStroke.Ln, null, this.worksheet.workbook.theme, this.worksheet.getDrawingDocument(), false, this.worksheet);
		private_SetCoords(oShape, this.worksheet, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset);
		return new ApiShape(oShape);
	};


	/**
	 * Adds an image to the current sheet with the parameters specified.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The image width in English measure units.
	 * @param {EMU} nHeight - The image height in English measure units.
	 * @param {number} nFromCol - The number of the column where the beginning of the image will be placed.
	 * @param {EMU} nColOffset - The offset from the nFromCol column to the left part of the image measured in English measure units.
	 * @param {number} nFromRow - The number of the row where the beginning of the image will be placed.
	 * @param {EMU} nRowOffset - The offset from the nFromRow row to the upper part of the image measured in English measure units.
	 * @returns {ApiImage}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddImage.js
	 */
	ApiWorksheet.prototype.AddImage = function (sImageSrc, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset) {
		var oImage = AscFormat.DrawingObjectsController.prototype.createImage(sImageSrc, 0, 0, nWidth / 36000, nHeight / 36000);
		private_SetCoords(oImage, this.worksheet, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset);
		return new ApiImage(oImage);
	};

	/**
	 * Adds a Text Art object to the current sheet with the parameters specified.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {ApiTextPr} [oTextPr=Api.CreateTextPr()] - The text properties.
	 * @param {string} [sText="Your text here"] - The text for the Text Art object.
	 * @param {TextTransform} [sTransform="textNoShape"] - Text transform type.
	 * @param {ApiFill} [oFill=Api.CreateNoFill()] - The color or pattern used to fill the Text Art object.
	 * @param {ApiStroke} [oStroke=Api.CreateStroke(0, Api.CreateNoFill())] - The stroke used to create the Text Art object shadow.
	 * @param {number} [nRotAngle=0] - Rotation angle.
	 * @param {EMU} [nWidth=1828800] - The Text Art width measured in English measure units.
	 * @param {EMU} [nHeight=1828800] - The Text Art heigth measured in English measure units.
	 * @param {number} [nFromCol=0] - The column number where the beginning of the Text Art object will be placed.
	 * @param {number} [nFromRow=0] - The row number where the beginning of the Text Art object will be placed.
	 * @param {EMU} [nColOffset=0] - The offset from the nFromCol column to the left part of the Text Art object measured in English measure units.
	 * @param {EMU} [nRowOffset=0] - The offset from the nFromRow row to the upper part of the Text Art object measured in English measure units.
	 * @returns {ApiDrawing}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddWordArt.js
	 */
	ApiWorksheet.prototype.AddWordArt = function (oTextPr, sText, sTransform, oFill, oStroke, nRotAngle, nWidth, nHeight, nFromCol, nFromRow, nColOffset, nRowOffset) {
		oTextPr = oTextPr && oTextPr.TextPr ? oTextPr.TextPr : null;
		nRotAngle = typeof (nRotAngle) === "number" && nRotAngle > 0 ? nRotAngle : 0;
		nWidth = typeof (nWidth) === "number" && nWidth > 0 ? nWidth : 1828800;
		nHeight = typeof (nHeight) === "number" && nHeight > 0 ? nHeight : 1828800;
		oFill = oFill && oFill.UniFill ? oFill.UniFill : Asc.editor.CreateNoFill().UniFill;
		oStroke = oStroke && oStroke.Ln ? oStroke.Ln : Asc.editor.CreateStroke(0, Asc.editor.CreateNoFill()).Ln;
		nFromCol = typeof (nFromCol) === "number" && nFromCol > 0 ? nFromCol : 0;
		nFromRow = typeof (nFromRow) === "number" && nFromRow > 0 ? nFromRow : 0;
		nColOffset = typeof (nColOffset) === "number" && nColOffset > 0 ? nColOffset : 0;
		nRowOffset = typeof (nRowOffset) === "number" && nRowOffset > 0 ? nRowOffset : 0;
		sTransform = typeof (sTransform) === "string" && sTransform !== "" ? sTransform : "textNoShape";

		var oArt = Asc.editor.private_createWordArt(oTextPr, sText, sTransform, oFill, oStroke, nRotAngle, nWidth, nHeight);

		private_SetCoords(oArt, this.worksheet, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset);

		return new ApiDrawing(oArt);
	};

	/**
	 * Adds an OLE object to the current sheet with the parameters specified.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently, only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The OLE object width in English measure units.
	 * @param {EMU} nHeight - The OLE object height in English measure units.
	 * @param {string} sData - The OLE object string data.
	 * @param {string} sAppId - The application ID associated with the current OLE object.
	 * @param {number} nFromCol - The number of the column where the beginning of the OLE object will be placed.
	 * @param {EMU} nColOffset - The offset from the nFromCol column to the left part of the OLE object measured in English measure units.
	 * @param {number} nFromRow - The number of the row where the beginning of the OLE object will be placed.
	 * @param {EMU} nRowOffset - The offset from the nFromRow row to the upper part of the OLE object measured in English measure units.
	 * @returns {ApiOleObject}
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddOleObject.js
	 */
	ApiWorksheet.prototype.AddOleObject = function (sImageSrc, nWidth, nHeight, sData, sAppId, nFromCol, nColOffset, nFromRow, nRowOffset) {
		if (typeof sImageSrc === "string" && sImageSrc.length > 0 && typeof sData === "string"
			&& typeof sAppId === "string" && sAppId.length > 0
			&& AscFormat.isRealNumber(nWidth) && AscFormat.isRealNumber(nHeight)
		)

			var nW = nWidth / 36000.0;
		var nH = nHeight / 36000.0;

		var oImage = AscFormat.DrawingObjectsController.prototype.createOleObject(sData, sAppId, sImageSrc, 0, 0, nW, nH);
		private_SetCoords(oImage, this.worksheet, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset);
		return new ApiOleObject(oImage);
	};

	/**
	 * Replaces the current image with a new one.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sImageUrl - The image source where the image to be inserted should be taken from (currently only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The image width in English measure units.
	 * @param {EMU} nHeight - The image height in English measure units.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/ReplaceCurrentImage.js
	 */
	ApiWorksheet.prototype.ReplaceCurrentImage = function (sImageUrl, nWidth, nHeight) {
		let oWorksheet = Asc['editor'].wb.getWorksheet();
		if (oWorksheet && oWorksheet.objectRender && oWorksheet.objectRender.controller) {
			let oController = oWorksheet.objectRender.controller;
			let dK = 1 / 36000 / AscCommon.g_dKoef_pix_to_mm;
			oController.putImageToSelection(sImageUrl, nWidth * dK, nHeight * dK);
		}
	};

	/**
	 * Returns all drawings from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiDrawing[]}.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllDrawings.js
	 */
	ApiWorksheet.prototype.GetAllDrawings = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject) {
				allApiDrawings.push(new ApiDrawing(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Returns all images from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiImage[]}.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllImages.js
	 */
	ApiWorksheet.prototype.GetAllImages = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject && allDrawings[nDrawing].isImage()) {
				allApiDrawings.push(new ApiImage(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Returns all shapes from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiShape[]}.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllShapes.js
	 */
	ApiWorksheet.prototype.GetAllShapes = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject && allDrawings[nDrawing].isShape()) {
				allApiDrawings.push(new ApiShape(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Returns all charts from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiChart[]}.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllCharts.js
	 */
	ApiWorksheet.prototype.GetAllCharts = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject && allDrawings[nDrawing].isChart()) {
				allApiDrawings.push(Asc.editor.private_CreateApiChart(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Returns all OLE objects from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiOleObject[]}.
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllOleObjects.js
	 */
	ApiWorksheet.prototype.GetAllOleObjects = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject && allDrawings[nDrawing].graphicObject instanceof AscFormat.COleObject) {
				allApiDrawings.push(new ApiOleObject(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Moves the current sheet to another location in the workbook.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {ApiWorksheet} before - The sheet before which the current sheet will be placed. You cannot specify "before" if you specify "after".
	 * @param {ApiWorksheet} after - The sheet after which the current sheet will be placed. You cannot specify "after" if you specify "before".
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/Move.js
	 */
	ApiWorksheet.prototype.Move = function (before, after) {
		let bb = before instanceof ApiWorksheet;
		let ba = after instanceof ApiWorksheet;
		if ((bb && ba) || (!bb && !ba)) {
			throwException(new Error('Incorrect parametrs.'));
		} else {
			let curIndex = this.GetIndex();
			let newIndex = (bb ? (before.GetIndex()) : (after.GetIndex() + 1));
			this.worksheet.workbook.oApi.asc_moveWorksheet(newIndex, [curIndex]);
		}
	};

	/**
	 * Return PivotTable by name.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name of PivotTable.
	 * @returns {ApiPivotTable|null}
	 * @see office-js-api/Examples/Cell/ApiWorksheet/Methods/GetPivotByName.js
	 */
	ApiWorksheet.prototype.GetPivotByName = function (name) {
		var res = null;
		if (name) {
			var pivot = this.worksheet.getPivotTableByName(name);
			if (pivot)
				res = new ApiPivotTable(pivot, this.worksheet.workbook.oApi);
		}
		return res;
	};

	/**
	 * Returns all PivotTables on worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotTable[]}
	 * @see office-js-api/Examples/Cell/ApiWorksheet/Methods/GetAllPivotTables.js
	 */
	ApiWorksheet.prototype.GetAllPivotTables = function () {
		var res = [];
		var ws = this.worksheet;
		ws.pivotTables.forEach(function(piv) {
			res.push( new ApiPivotTable(piv, ws.workbook.oApi) );
		});
		return res;
	};

	Object.defineProperty(ApiWorksheet.prototype, "PivotTables", {
		get: function () {
			return this.GetAllPivotTables();
		}
	});

	/**
	 * Refresh all PivotTables on current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/Cell/ApiWorksheet/Methods/RefreshAllPivots.js
	 */
	ApiWorksheet.prototype.RefreshAllPivots = function () {
		const t = this;
		this.worksheet.pivotTables.forEach(function(pivot) {
			pivot.asc_refresh(t.worksheet.workbook.oApi);
		});
	};

	/**
	 * Returns the freeze panes from the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiFreezePanes}
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetFreezePanes.js
	 */
	ApiWorksheet.prototype.GetFreezePanes = function () {
		return new ApiFreezePanes(this.worksheet);
	};

	Object.defineProperty(ApiWorksheet.prototype, "FreezePanes", {
		get: function () {
			return this.GetFreezePanes();
		}
	});

	/**
	 * Creates a protected range of the specified type from the selected data range of the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sTitle - The title which will be displayed for the current protected range.
	 * @param {string} sDataRange - The selected cell range which will be used to get the data for the protected range.
	 * @returns {ApiProtectedRange | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/AddProtectedRange.js
	 */
	ApiWorksheet.prototype.AddProtectedRange = function (sTitle, sDataRange) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let isValidRef = typeof (sDataRange) === 'string' && sDataRange.trim() !== '';
		let result = null;
		if (isValidTitle && isValidRef) {
			let settings = new Asc.CUserProtectedRange(this.worksheet);
			settings.asc_setName(sTitle);
			settings.asc_setRef(sDataRange);

			let docInfo = this.worksheet.workbook.oApi && this.worksheet.workbook.oApi.DocInfo;
			if (docInfo) {
				let userInfo = docInfo.UserInfo;
				if (userInfo) {
					let users = [];
					let user = new Asc.CUserProtectedRangeUserInfo();

					user.asc_setId(userInfo.asc_getId());
					user.asc_setName(userInfo.get_FullName());

					users.push(user);
					settings.asc_setUsers(users);
				}
			}
			let editRes = this.worksheet.editUserProtectedRanges(null, settings, true);
			if (typeof editRes === "object") {
				result = new ApiProtectedRange(editRes);
			} else {
				logError(new Error('Protected range cannot be added.'));
			}
		} else {
			logError(new Error('The title or dataRange is invalid'));
		}

		return result;
	};


	/**
	 * Returns a protected range object by its title.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sTitle - The title of the protected range that will be returned.
	 * @returns {ApiProtectedRange | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetProtectedRange.js
	 */
	ApiWorksheet.prototype.GetProtectedRange = function (sTitle) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let result = null;
		if (isValidTitle) {
			let protectedRange = this.worksheet.getUserProtectedRangeByName(sTitle);
			result = protectedRange && protectedRange.obj ? new ApiProtectedRange(protectedRange.obj.clone(protectedRange.obj._ws, true)) : null;
			if (result === null) {
				logError(new Error('The range not found'));
			}
		} else {
			logError(new Error('The title is invalid'));
		}

		return result;
	};

	/**
	 * Returns all protected ranges from the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiProtectedRange[] | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/GetAllProtectedRanges.js
	 */
	ApiWorksheet.prototype.GetAllProtectedRanges = function () {
		let protectedRanges = this.worksheet && this.worksheet.workbook && this.worksheet.workbook.oApi.asc_getUserProtectedRanges(this.worksheet.sName);
		let result = null;
		if (protectedRanges) {
			result = [];
			for (let i  = 0; i < protectedRanges.length; i++) {
				result.push(new ApiProtectedRange(protectedRanges[i].clone(protectedRanges[i]._ws, true)));
			}
		} else {
			logError(new Error('Ranges not found'));
		}

		return result;
	};
	Object.defineProperty(ApiWorksheet.prototype, "AllProtectedRanges", {
		get: function () {
			return this.GetAllProtectedRanges();
		}
	});

	/**
	 * Pastes the contents of the clipboard to the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - The cell range where the clipboard contents should be pasted. If this argument is omitted, the current selection is used.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiWorksheet/Methods/Paste.js
	 */
	ApiWorksheet.prototype.Paste = function (destination) {
		var oApi = Asc["editor"];
		if (destination) {
			if (destination instanceof ApiRange) {
				AscCommon.g_specialPasteHelper && AscCommon.g_specialPasteHelper.Special_Paste_Hide_Button();
				let ws =  destination.range.worksheet;
				private_executeOtherActiveSheet(ws, destination.areas || [destination.range], function () {
					oApi && oApi.asc_Paste();
				});
			} else {
				logError(new Error('Invalid destination'));
			}
		} else {
			AscCommon.g_specialPasteHelper && AscCommon.g_specialPasteHelper.Special_Paste_Hide_Button();
			let ws = this.worksheet;
			private_executeOtherActiveSheet(ws, null, function () {
				oApi && oApi.asc_Paste();
			});
		}
	};



	/**
	 * Specifies the cell border position.
	 * @typedef {("DiagonalDown" | "DiagonalUp" | "Bottom" | "Left" | "Right" | "Top" | "InsideHorizontal" | "InsideVertical")} BordersIndex
	 * @see office-js-api/Examples/Enumerations/BordersIndex.js
	 */

	/**
	 * Specifies the line style used to form the cell border.
	 * @typedef {("None" | "Double" | "Hair" | "DashDotDot" | "DashDot" | "Dotted" | "Dashed" | "Thin" | "MediumDashDotDot" | "SlantDashDot" | "MediumDashDot" | "MediumDashed" | "Medium" | "Thick")} LineStyle
	 * @see office-js-api/Examples/Enumerations/LineStyle.js
	 */

	//TODO xlManual param
	/**
	 * Specifies the sort order.
	 * @typedef {("xlAscending" | "xlDescending")}  SortOrder
	 * @see office-js-api/Examples/Enumerations/SortOrder.js
	 */

	//TODO xlGuess param
	/**
	 * Specifies whether the first row of the sort range contains the header information.
	 * @typedef {("xlNo" | "xlYes")} SortHeader
	 * @see office-js-api/Examples/Enumerations/SortHeader.js
	 */

	/**
	 * Specifies if the sort should be by row or column.
	 * @typedef {("xlSortColumns" | "xlSortRows")} SortOrientation
	 * @see office-js-api/Examples/Enumerations/SortOrientation.js
	 */

	/**
	 * Specifies the range angle.
	 * @typedef {("xlDownward" | "xlHorizontal" | "xlUpward" | "xlVertical")} Angle
	 * @see office-js-api/Examples/Enumerations/Angle.js
	 */

	/**
	 * Specifies the direction of end in the specified range.
	 * @typedef {("xlUp" | "xlDown" | "xlToRight" | "xlToLeft")} Direction
	 * @see office-js-api/Examples/Enumerations/Direction.js
	 */

	/**
	 * Returns a type of the ApiRange class.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {"range"}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetClassType.js
	 */
	ApiRange.prototype.GetClassType = function () {
		return "range";
	};

	/**
	 * Returns a row number for the selected cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetRow.js
	 */
	ApiRange.prototype.GetRow = function () {
		return (this.range.bbox.r1 + 1);
	};
	Object.defineProperty(ApiRange.prototype, "Row", {
		get: function () {
			return this.GetRow();
		}
	});
	/**
	 * Returns a column number for the selected cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetCol.js
	 */
	ApiRange.prototype.GetCol = function () {
		return (this.range.bbox.c1 + 1);
	};
	Object.defineProperty(ApiRange.prototype, "Col", {
		get: function () {
			return this.GetCol();
		}
	});

	/**
	 * Clears the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Clear.js
	 */
	ApiRange.prototype.Clear = function () {
		let range = this.range,
			bbox = range.bbox,
			ws = range.worksheet,
			wsView = Asc['editor'].wb.getWorksheet(ws.getIndex());
		range.cleanAll();
		ws.deletePivotTables(bbox);
		ws.removeSparklines(bbox);
		ws.clearDataValidation([bbox], true);
		ws.clearConditionalFormattingRulesByRanges([bbox]);
		wsView.cellCommentator.deleteCommentsRange(bbox, null);
	};

	/**
	 * Returns a Range object that represents the rows in the specified range. If the specified row is outside the Range object, a new Range will be returned that represents the cells between the columns of the original range in the specified row.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nRow - The row number (starts counting from 1, the 0 value returns an error).
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetRows.js
	 */
	ApiRange.prototype.GetRows = function (nRow) {
		let result = null;
		if (typeof nRow === "undefined") {
			result = this;
		} else {
			if (typeof nRow === "number") {
				nRow--;
				let r = this.range.bbox.r1 + nRow;
				if (r > AscCommon.gc_nMaxRow0) r = AscCommon.gc_nMaxRow0;
				if (r < 0) r = 0;
				result = new ApiRange(this.range.worksheet.getRange3(r, this.range.bbox.c1, r, this.range.bbox.c2));
			} else {
				logError(new Error('The nRow must be a number that greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1)));
			}
		}
		return result;
	};
	Object.defineProperty(ApiRange.prototype, "Rows", {
		get: function () {
			return this.GetRows();
		}
	});

	/**
	 * Returns a Range object that represents the columns in the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nCol - The column number. *
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetCols.js
	 */
	ApiRange.prototype.GetCols = function (nCol) {
		let result = null;
		if (typeof nCol === "undefined") {
			result = this;
		} else {
			if (typeof nCol === "number") {
				nCol--;
				let c = this.range.bbox.c1 + nCol;
				if (c > AscCommon.gc_nMaxCol0) c = AscCommon.gc_nMaxCol0;
				if (c < 0) c = 0;
				result = new ApiRange(this.range.worksheet.getRange3(this.range.bbox.r1, c, this.range.bbox.r2, c));
			} else {
				logError(new Error('The nCol must be a number that greater than 0 and less then ' + (AscCommon.gc_nMaxCol0 + 1)));
			}
		}
		return result;
	};
	Object.defineProperty(ApiRange.prototype, "Cols", {
		get: function () {
			return this.GetCols();
		}
	});

	/**
	 * Returns a Range object that represents the end in the specified direction in the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {Direction} direction - The direction of end in the specified range. *
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/End.js
	 */
	ApiRange.prototype.End = function (direction) {
		let bbox = this.range.bbox;
		let row, col, res;
		switch (direction) {
			case "xlUp":
				row = (bbox.r1 > 0 ? bbox.r1 - 1 : bbox.r1);
				res = this.range.worksheet.getRange3(0, bbox.c1, 0, bbox.c1);
				while (row) {
					let cell = this.range.worksheet.getRange3(row, bbox.c1, row, bbox.c1);
					if (cell.getValue() !== "") {
						res = cell;
						break;
					}
					row--;
				}
				break;
			case "xlDown":
				row = (bbox.r1 < AscCommon.gc_nMaxRow0 ? bbox.r1 + 1 : bbox.r1);
				res = this.range.worksheet.getRange3(AscCommon.gc_nMaxRow0, bbox.c1, AscCommon.gc_nMaxRow0, bbox.c1);
				while (row < AscCommon.gc_nMaxRow0) {
					let cell = this.range.worksheet.getRange3(row, bbox.c1, row, bbox.c1);
					if (cell.getValue() !== "") {
						res = cell;
						break;
					}
					row++;
				}
				break;
			case "xlToRight":
				col = (bbox.c1 < AscCommon.gc_nMaxCol0 ? bbox.c1 + 1 : bbox.c1);
				res = this.range.worksheet.getRange3(bbox.r1, AscCommon.gc_nMaxCol0, bbox.r1, AscCommon.gc_nMaxCol0);
				while (col < AscCommon.gc_nMaxCol0) {
					let cell = this.range.worksheet.getRange3(bbox.r1, col, bbox.r1, col);
					if (cell.getValue() !== "") {
						res = cell;
						break;
					}
					col++;
				}
				break;
			case "xlToLeft":
				col = (bbox.c1 > 0 ? bbox.c1 - 1 : bbox.c1);
				res = this.range.worksheet.getRange3(bbox.r1, 0, bbox.r1, 0);
				while (col) {
					let cell = this.range.worksheet.getRange3(bbox.r1, col, bbox.r1, col);
					if (cell.getValue() !== "") {
						res = cell;
						break;
					}
					col--;
				}
				break;
			default:
				res = this.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r2, bbox.c2);
				break;
		}
		return new ApiRange(res);
	};

	/**
	 * Returns a Range object that represents all the cells in the specified range or a specified cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} row - The row number or the cell number (if only row is defined).
	 * @param {number} col - The column number.
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetCells.js
	 */
	ApiRange.prototype.GetCells = function (row, col) {
		let bbox = this.range.bbox;
		let r1, c1, result;
		if (typeof col == "number" && typeof row == "number") {
			row--;
			col--;
			r1 = bbox.r1 + row;
			c1 = bbox.c1 + col;
		} else if (typeof row == "number") {
			row--;
			let cellCount = bbox.c2 - bbox.c1 + 1;
			r1 = bbox.r1 + ((row) ? (row / cellCount) >> 0 : row);
			c1 = bbox.c1 + ((r1) ? 1 : 0) + ((row) ? row % cellCount : row);
			if (r1 && c1) c1--;
		} else if (typeof col == "number") {
			col--;
			r1 = bbox.r1;
			c1 = bbox.c1 + col;
		} else {
			result = new ApiRange(this.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r2, bbox.c2));
		}

		if (!result) {
			if (r1 > AscCommon.gc_nMaxRow0) r1 = AscCommon.gc_nMaxRow0;
			if (r1 < 0) r1 = 0;
			if (c1 > AscCommon.gc_nMaxCol0) c1 = AscCommon.gc_nMaxCol0;
			if (c1 < 0) c1 = 0;
			result = new ApiRange(this.range.worksheet.getRange3(r1, c1, r1, c1));
		}
		return result;
	};
	Object.defineProperty(ApiRange.prototype, "Cells", {
		get: function () {
			return this.GetCells();
		}
	});

	/**
	 * Sets the cell offset.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nRow - The row number.
	 * @param {number} nCol - The column number.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetOffset.js
	 */
	ApiRange.prototype.SetOffset = function (nRow, nCol) {
		this.range.setOffset({row: nRow, col: nCol});
	};

	/**
	 * Returns the range address.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} RowAbs - Defines if the link to the row is absolute or not.
	 * @param {boolean} ColAbs - Defines if the link to the column is absolute or not.
	 * @param {string} RefStyle - The reference style.
	 * @param {boolean} External - Defines if the range is in the current file or not.
	 * @param {range} RelativeTo - The range which the current range is relative to.
	 * @returns {string | null} - returns address of range as string.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetAddress.js
	 */
	ApiRange.prototype.GetAddress = function (RowAbs, ColAbs, RefStyle, External, RelativeTo) {
		// todo поправить, чтобы возвращал адреса всех areas внутри range
		var range = this.range.bbox;
		var isOneCell = this.range.isOneCell();
		var isOneCol = (this.range.bbox.c1 === this.range.bbox.c2 && this.range.bbox.r1 === 0 && this.range.bbox.r2 === AscCommon.gc_nMaxRow0);
		var isOneRow = (this.range.bbox.r1 === this.range.bbox.r2 && this.range.bbox.c1 === 0 && this.range.bbox.c2 === AscCommon.gc_nMaxCol0);
		var ws = this.range.worksheet;
		var value;
		var row1 = range.r1 + ((RowAbs || RefStyle != "xlR1C1") ? 1 : 0),
			col1 = range.c1 + ((ColAbs || RefStyle != "xlR1C1") ? 1 : 0),
			row2 = range.r2 + ((RowAbs || RefStyle != "xlR1C1") ? 1 : 0),
			col2 = range.c2 + ((ColAbs || RefStyle != "xlR1C1") ? 1 : 0);
		if (RefStyle == 'xlR1C1') {
			if (RowAbs) {
				row1 = "R" + row1;
				row2 = isOneCell ? "" : ":R" + row2;
			} else {
				var tmpR = (RelativeTo instanceof ApiRange) ? RelativeTo.range.bbox.r1 : 0;
				row1 = "R" + ((row1 - tmpR) !== 0 ? "[" + (row1 - tmpR) + "]" : "");
				row2 = isOneCell ? "" : ":R" + ((row2 - tmpR) !== 0 ? "[" + (row2 - tmpR) + "]" : "");
			}

			if (ColAbs) {
				col1 = "C" + col1;
				col2 = isOneCell ? "" : "C" + col2;
			} else {
				var tmpC = (RelativeTo instanceof ApiRange) ? RelativeTo.range.bbox.c1 : 0;
				col1 = "C" + ((col1 - tmpC) !== 0 ? "[" + (col1 - tmpC) + "]" : "");
				col2 = isOneCell ? "" : "C" + ((col2 - tmpC) !== 0 ? "[" + (col2 - tmpC) + "]" : "");
			}
			value = isOneCol ? col1 : isOneRow ? row1 : row1 + col1 + row2 + col2;
		} else {
			// xlA1 - default
			row1 = (RowAbs ? "$" : "") + row1;
			col1 = (ColAbs ? "$" : "") + AscCommon.g_oCellAddressUtils.colnumToColstr(col1);
			row2 = isOneCell ? "" : ((RowAbs ? "$" : "") + row2);
			col2 = isOneCell ? "" : ((ColAbs ? ":$" : ":") + AscCommon.g_oCellAddressUtils.colnumToColstr(col2));
			value = isOneCol ? col1 + col2 : isOneRow ? row1 + ":" + row2 : col1 + row1 + col2 + row2;
		}
		return (External) ? '[' + ws.workbook.oApi.DocInfo.Title + ']' + AscCommon.parserHelp.get3DRef(ws.sName, value) : value;
	};
	Object.defineProperty(ApiRange.prototype, "Address", {
		get: function () {
			return this.GetAddress(true, true);
		}
	});

	/**
	 * Returns the rows or columns count.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetCount.js
	 */
	ApiRange.prototype.GetCount = function () {
		var range = this.range.bbox;
		var count;
		switch (range.getType()) {
			case Asc.c_oAscSelectionType.RangeCells:
				count = (range.c2 - range.c1 + 1) * (range.r2 - range.r1 + 1);
				break;

			case Asc.c_oAscSelectionType.RangeCol:
				count = range.c2 - range.c1 + 1;
				break;

			case Asc.c_oAscSelectionType.RangeRow:
				count = range.r2 - range.r1 + 1;
				break;

			case Asc.c_oAscSelectionType.RangeMax:
				count = range.r2 * range.c2;
				break;
		}
		return count;
	};
	Object.defineProperty(ApiRange.prototype, "Count", {
		get: function () {
			return this.GetCount();
		}
	});

	/**
	 * Returns a value of the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {string | string[][]}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetValue.js
	 */
	ApiRange.prototype.GetValue = function () {
		var bbox = this.range.bbox;
		var nCol = bbox.c2 - bbox.c1 + 1;
		var nRow = bbox.r2 - bbox.r1 + 1;
		var res;
		if (this.range.isOneCell()) {
			res = this.range.getValue();
		} else {
			res = [];
			for (var i = 0; i < nRow; i++) {
				var arr = [];
				for (var k = 0; k < nCol; k++) {
					var cell = this.range.worksheet.getRange3((bbox.r1 + i), (bbox.c1 + k), (bbox.r1 + i), (bbox.c1 + k));
					arr.push(cell.getValue());
				}
				res.push(arr);
			}
		}
		return res;
	};

	/**
	 * Sets a value to the current cell or cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string | bool | number | Array[] | Array[][]} data - The general value for the cell or cell range.
	 * @returns {boolean} - returns false if such a range does not exist.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetValue.js
	 */
	ApiRange.prototype.SetValue = function (data) {
		if (!this.range)
			return false;

		let worksheet = this.range.worksheet;

		if (Array.isArray(data)) {
			let checkDepth = function (x) {
				return Array.isArray(x) ? 1 + Math.max.apply(this, x.map(checkDepth)) : 0;
			};
			let maxDepth = checkDepth(data);
			if (maxDepth <= 2) {
				if (this.range.isOneCell()) {
					data = maxDepth == 1 ? data[0] : data[0][0];
				} else {
					let bbox = this.range.bbox;
					let nRow = bbox.r2 - bbox.r1 + 1;
					let nCol = bbox.c2 - bbox.c1 + 1;
					for (let indC = 0; indC < nCol; indC++) {
						for (let indR = 0; indR < nRow; indR++) {
							let value = (maxDepth == 1 ? data[indC] : data[indR] ? data[indR][indC] : null);
							if (value === undefined || value === null)
								value = AscCommon.cErrorLocal["na"];

							let cell = this.range.worksheet.getRange3((bbox.r1 + indR), (bbox.c1 + indC), (bbox.r1 + indR), (bbox.c1 + indC));
							let merged = cell.hasMerged();
							if (merged)
								cell = this.range.worksheet.getRange3(merged.r1, merged.c1, merged.r1, merged.c1);

							value = checkFormat(value.toString());
							cell.setValue(value.toString());
							if (value.type === AscCommonExcel.cElementType.number)
								cell.setNumFormat(AscCommon.getShortDateFormat());
						}
					}
					worksheet.workbook.handlers.trigger("cleanCellCache", worksheet.getId(), [this.range.bbox], true);
					worksheet.workbook.oApi.onWorksheetChange(this.range.bbox);
					return true;
				}
			}
		}
		if (data === undefined || data === null)
			data = AscCommon.cErrorLocal["na"];

		data = checkFormat(data);
		let range = this.range;
		let merged = range.hasMerged();
		if (merged)
			range = this.range.worksheet.getRange3(merged.r1, merged.c1, merged.r1, merged.c1);

		range.setValue(data.toString());
		if (data.type === AscCommonExcel.cElementType.number)
			range.setNumFormat(AscCommon.getShortDateFormat());

		worksheet.workbook.handlers.trigger("cleanCellCache", worksheet.getId(), [range.bbox], true);
		worksheet.workbook.oApi.onWorksheetChange(range.bbox);
		return true;
	};

	Object.defineProperty(ApiRange.prototype, "Value", {
		get: function () {
			return this.GetValue();
		},
		set: function (sValue) {
			this.SetValue(sValue);
		}
	});

	/**
	 * Returns a formula of the specified range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @returns {string | string[][]} - return Value2 property (value without format) if formula doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetFormula.js
	 */
	ApiRange.prototype.GetFormula = function () {
		if (this.range.isFormula())
			return "= " + this.range.getFormula();
		else
			return this.GetValue2();
	};

	Object.defineProperty(ApiRange.prototype, "Formula", {
		get: function () {
			return this.GetFormula();
		},
		set: function (value) {
			this.SetValue(value);
		}
	});

	/**
	 * Returns the Value2 property (value without format) of the specified range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @returns {string | string[][]}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetValue2.js
	 */
	ApiRange.prototype.GetValue2 = function () {
		var bbox = this.range.bbox;
		var nCol = bbox.c2 - bbox.c1 + 1;
		var nRow = bbox.r2 - bbox.r1 + 1;
		var res;
		if (this.range.isOneCell()) {
			res = this.range.getValueWithoutFormat();
		} else {
			res = [];
			for (var i = 0; i < nRow; i++) {
				var arr = [];
				for (var k = 0; k < nCol; k++) {
					var cell = this.range.worksheet.getRange3((bbox.r1 + i), (bbox.c1 + k), (bbox.r1 + i), (bbox.c1 + k));
					arr.push(cell.getValueWithoutFormat());
				}
				res.push(arr);
			}
		}
		return res;
	};

	Object.defineProperty(ApiRange.prototype, "Value2", {
		get: function () {
			return this.GetValue2();
		},
		set: function (value) {
			this.SetValue(value);
		}
	});

	/**
	 * Returns the text of the specified range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @returns {string | string[][]}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetText.js
	 */
	ApiRange.prototype.GetText = function () {
		var bbox = this.range.bbox;
		var nCol = bbox.c2 - bbox.c1 + 1;
		var nRow = bbox.r2 - bbox.r1 + 1;
		var res;
		if (this.range.isOneCell()) {
			res = this.range.getValueWithFormat();
		} else {
			res = [];
			for (var i = 0; i < nRow; i++) {
				var arr = [];
				for (var k = 0; k < nCol; k++) {
					var cell = this.range.worksheet.getRange3((bbox.r1 + i), (bbox.c1 + k), (bbox.r1 + i), (bbox.c1 + k));
					arr.push(cell.getValueWithFormat());
				}
				res.push(arr);
			}
		}
		return res;
	};

	Object.defineProperty(ApiRange.prototype, "Text", {
		get: function () {
			return this.GetText();
		},
		set: function (value) {
			this.SetValue(value);
		}
	});

	/**
	 * Sets the text color to the current cell range with the previously created color object.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiColor} oColor - The color object which specifies the color to be set to the text in the cell / cell range.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetFontColor.js
	 */
	ApiRange.prototype.SetFontColor = function (oColor) {
		this.range.setFontcolor(oColor.color);
	};
	Object.defineProperty(ApiRange.prototype, "FontColor", {
		set: function (oColor) {
			return this.SetFontColor(oColor);
		}
	});

	/**
	 * Returns the value hiding property. The specified range must span an entire column or row.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {boolean} - returns true if the values in the range specified are hidden.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetHidden.js
	 */
	ApiRange.prototype.GetHidden = function () {
		var range = this.range;
		var worksheet = range.worksheet;
		var bbox = range.bbox;
		switch (bbox.getType()) {
			case Asc.c_oAscSelectionType.RangeCol:
				return worksheet.getColHidden(bbox.c1);

			case Asc.c_oAscSelectionType.RangeRow:
				return worksheet.getRowHidden(bbox.r1);

			default:
				return false;
		}
	};
	/**
	 * Sets the value hiding property. The specified range must span an entire column or row.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isHidden - Specifies if the values in the current range are hidden or not.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetHidden.js
	 */
	ApiRange.prototype.SetHidden = function (isHidden) {
		var range = this.range;
		var worksheet = range.worksheet;
		var bbox = range.bbox;
		switch (bbox.getType()) {
			case Asc.c_oAscSelectionType.RangeCol:
				worksheet.setColHidden(isHidden, bbox.c1, bbox.c2);
				break;

			case Asc.c_oAscSelectionType.RangeRow:
				worksheet.setRowHidden(isHidden, bbox.r1, bbox.r2);
				break;
		}
	};
	Object.defineProperty(ApiRange.prototype, "Hidden", {
		get: function () {
			return this.GetHidden();
		},
		set: function (isHidden) {
			this.SetHidden(isHidden);
		}
	});

	/**
	 * Returns the column width value.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetColumnWidth.js
	 */
	ApiRange.prototype.GetColumnWidth = function () {
		var ws = this.range.worksheet;
		var width = ws.getColWidth(this.range.bbox.c1);
		width = (width < 0) ? AscCommonExcel.oDefaultMetrics.ColWidthChars : width;
		return ws.colWidthToCharCount(ws.modelColWidthToColWidth(width));
	};
	/**
	 * Sets the width of all the columns in the current range.
	 * One unit of column width is equal to the width of one character in the Normal style.
	 * For proportional fonts, the width of the character 0 (zero) is used.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nWidth - The width of the column divided by 7 pixels.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetColumnWidth.js
	 */
	ApiRange.prototype.SetColumnWidth = function (nWidth) {
		this.range.worksheet.setColWidth(nWidth, this.range.bbox.c1, this.range.bbox.c2);
	};
	Object.defineProperty(ApiRange.prototype, "ColumnWidth", {
		get: function () {
			return this.GetColumnWidth();
		},
		set: function (nWidth) {
			this.SetColumnWidth(nWidth);
		}
	});
	Object.defineProperty(ApiRange.prototype, "Width", {
		get: function () {
			var max = this.range.bbox.c2 - this.range.bbox.c1;
			var ws = this.range.worksheet;
			var sum = 0;
			var width;
			for (var i = 0; i <= max; i++) {
				width = ws.getColWidth(i);
				width = (width < 0) ? AscCommonExcel.oDefaultMetrics.ColWidthChars : width;
				sum += ws.modelColWidthToColWidth(width);
			}
			return sum;
		}
	});

	/**
	 * Returns the row height value.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {pt} - The row height in the range specified, measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetRowHeight.js
	 */
	ApiRange.prototype.GetRowHeight = function () {
		return this.range.worksheet.getRowHeight(this.range.bbox.r1);
	};

	/**
	 * Sets the row height value.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {pt} nHeight - The row height in the current range measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetRowHeight.js
	 */
	ApiRange.prototype.SetRowHeight = function (nHeight) {
		this.range.worksheet.setRowHeight(nHeight, this.range.bbox.r1, this.range.bbox.r2, true);
	};
	Object.defineProperty(ApiRange.prototype, "RowHeight", {
		get: function () {
			return this.GetRowHeight();
		},
		set: function (nHeight) {
			this.SetRowHeight(nHeight);
		}
	});
	Object.defineProperty(ApiRange.prototype, "Height", {
		get: function () {
			var max = this.range.bbox.r2 - this.range.bbox.r1;
			var sum = 0;
			for (var i = 0; i <= max; i++) {
				sum += this.range.worksheet.getRowHeight(i);
			}
			return sum;
		}
	});

	/**
	 * Sets the font size to the characters of the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nSize - The font size value measured in points.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetFontSize.js
	 */
	ApiRange.prototype.SetFontSize = function (nSize) {
		this.range.setFontsize(nSize);
	};
	Object.defineProperty(ApiRange.prototype, "FontSize", {
		set: function (nSize) {
			return this.SetFontSize(nSize);
		}
	});

	/**
	 * Sets the specified font family as the font name for the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The font family name used for the current cell range.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetFontName.js
	 */
	ApiRange.prototype.SetFontName = function (sName) {
		this.range.setFontname(sName);
	};
	Object.defineProperty(ApiRange.prototype, "FontName", {
		set: function (sName) {
			return this.SetFontName(sName);
		}
	});

	/**
	 * Sets the vertical alignment of the text in the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {'center' | 'bottom' | 'top' | 'distributed' | 'justify'} sAligment - The vertical alignment that will be applied to the cell contents.
	 * @returns {boolean} - return false if sAligment doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetAlignVertical.js
	 */
	ApiRange.prototype.SetAlignVertical = function (sAligment) {
		switch (sAligment) {
			case "center": {
				this.range.setAlignVertical(Asc.c_oAscVAlign.Center);
				break;
			}
			case "bottom": {
				this.range.setAlignVertical(Asc.c_oAscVAlign.Bottom);
				break;
			}
			case "top": {
				this.range.setAlignVertical(Asc.c_oAscVAlign.Top);
				break;
			}
			case "distributed": {
				this.range.setAlignVertical(Asc.c_oAscVAlign.Dist);
				break;
			}
			case "justify": {
				this.range.setAlignVertical(Asc.c_oAscVAlign.Just);
				break;
			}
			default :
				return false;
		}

		return true;
	};
	Object.defineProperty(ApiRange.prototype, "AlignVertical", {
		set: function (sAligment) {
			return this.SetAlignVertical(sAligment);
		}
	});

	/**
	 * Sets the horizontal alignment of the text in the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {'left' | 'right' | 'center' | 'justify'} sAlignment - The horizontal alignment that will be applied to the cell contents.
	 * @returns {boolean} - return false if sAligment doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetAlignHorizontal.js
	 */
	ApiRange.prototype.SetAlignHorizontal = function (sAlignment) {
		switch (sAlignment) {
			case "left": {
				this.range.setAlignHorizontal(AscCommon.align_Left);
				break;
			}
			case "right": {
				this.range.setAlignHorizontal(AscCommon.align_Right);
				break;
			}
			case "justify": {
				this.range.setAlignHorizontal(AscCommon.align_Justify);
				break;
			}
			case "center": {
				this.range.setAlignHorizontal(AscCommon.align_Center);
				break;
			}
			default :
				return false;
		}

		return true;
	};
	Object.defineProperty(ApiRange.prototype, "AlignHorizontal", {
		set: function (sAlignment) {
			return this.SetAlignHorizontal(sAlignment);
		}
	});

	/**
	 * Sets the bold property to the text characters in the current cell or cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isBold - Specifies that the contents of the current cell / cell range are displayed bold.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetBold.js
	 */
	ApiRange.prototype.SetBold = function (isBold) {
		this.range.setBold(!!isBold);
	};
	Object.defineProperty(ApiRange.prototype, "Bold", {
		set: function (isBold) {
			return this.SetBold(isBold);
		}
	});

	/**
	 * Sets the italic property to the text characters in the current cell or cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current cell / cell range are displayed italicized.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetItalic.js
	 */
	ApiRange.prototype.SetItalic = function (isItalic) {
		this.range.setItalic(!!isItalic);
	};
	Object.defineProperty(ApiRange.prototype, "Italic", {
		set: function (isItalic) {
			return this.SetItalic(isItalic);
		}
	});

	/**
	 * Specifies that the contents of the current cell / cell range are displayed along with a line appearing directly below the character.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {'none' | 'single' | 'singleAccounting' | 'double' | 'doubleAccounting'} undelineType - Specifies the type of the
	 * line displayed under the characters. The following values are available:
	 * <b>"none"</b> - for no underlining;
	 * <b>"single"</b> - for a single line underlining the cell contents;
	 * <b>"singleAccounting"</b> - for a single line underlining the cell contents but not protruding beyond the cell borders;
	 * <b>"double"</b> - for a double line underlining the cell contents;
	 * <b>"doubleAccounting"</b> - for a double line underlining the cell contents but not protruding beyond the cell borders.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetUnderline.js
	 */
	ApiRange.prototype.SetUnderline = function (undelineType) {
		var val;
		switch (undelineType) {
			case 'single':
				val = Asc.EUnderline.underlineSingle;
				break;
			case 'singleAccounting':
				val = Asc.EUnderline.underlineSingleAccounting;
				break;
			case 'double':
				val = Asc.EUnderline.underlineDouble;
				break;
			case 'doubleAccounting':
				val = Asc.EUnderline.underlineDoubleAccounting;
				break;
			case 'none':
			default:
				val = Asc.EUnderline.underlineNone;
				break;
		}
		this.range.setUnderline(val);
	};
	Object.defineProperty(ApiRange.prototype, "Underline", {
		set: function (undelineType) {
			return this.SetUnderline(undelineType);
		}
	});

	/**
	 * Specifies that the contents of the cell / cell range are displayed with a single horizontal line through the center of the contents.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isStrikeout - Specifies if the contents of the current cell / cell range are displayed struck through.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetStrikeout.js
	 */
	ApiRange.prototype.SetStrikeout = function (isStrikeout) {
		this.range.setStrikeout(!!isStrikeout);
	};
	Object.defineProperty(ApiRange.prototype, "Strikeout", {
		set: function (isStrikeout) {
			return this.SetStrikeout(isStrikeout);
		}
	});

	/**
	 * Specifies whether the words in the cell must be wrapped to fit the cell size or not.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isWrap - Specifies if the words in the cell will be wrapped to fit the cell size.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetWrap.js
	 */
	ApiRange.prototype.SetWrap = function (isWrap) {
		this.range.setWrap(!!isWrap);
	};

	/**
	 * Returns the information about the wrapping cell style.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetWrapText.js
	 */
	ApiRange.prototype.GetWrapText = function () {
		return this.range.getAlign().getWrap();
	};
	Object.defineProperty(ApiRange.prototype, "WrapText", {
		set: function (isWrap) {
			this.SetWrap(isWrap);
		},
		get: function () {
			return this.GetWrapText();
		}
	});

	/**
	 * Sets the background color to the current cell range with the previously created color object.
	 * Sets 'No Fill' when previously created color object is null.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiColor} oColor - The color object which specifies the color to be set to the background in the cell / cell range.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetFillColor.js
	 */
	ApiRange.prototype.SetFillColor = function (oColor) {
		this.range.setFillColor('No Fill' === oColor ? null : oColor.color);
	};
	/**
	 * Returns the background color for the current cell range. Returns 'No Fill' when the color of the background in the cell / cell range is null.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiColor|'No Fill'} - return 'No Fill' when the color to the background in the cell / cell range is null.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetFillColor.js
	 */
	ApiRange.prototype.GetFillColor = function () {
		var oColor = this.range.getFillColor();
		return oColor ? new ApiColor(oColor) : 'No Fill';
	};
	Object.defineProperty(ApiRange.prototype, "FillColor", {
		set: function (oColor) {
			return this.SetFillColor(oColor);
		},
		get: function () {
			return this.GetFillColor();
		}
	});

	/**
	 * Returns a value that represents the format code for the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {string | null} This property returns null if all cells in the specified range don't have the same number format.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetNumberFormat.js
	 */
	ApiRange.prototype.GetNumberFormat = function () {
		var bbox = this.range.bbox;
		var nCol = bbox.c2 - bbox.c1 + 1;
		var nRow = bbox.r2 - bbox.r1 + 1;
		var res = this.range.getNumFormatStr();
		if (!this.range.isOneCell()) {
			for (var i = 0; i < nRow; i++) {
				for (var k = 0; k < nCol; k++) {
					var cell = this.range.worksheet.getRange3((bbox.r1 + i), (bbox.c1 + k), (bbox.r1 + i), (bbox.c1 + k));
					if (res !== cell.getNumFormatStr())
						return null;
				}
			}
		}
		return res;
	};
	/**
	 * Specifies whether a number in the cell should be treated like number, currency, date, time, etc. or just like text.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sFormat - Specifies the mask applied to the number in the cell.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetNumberFormat.js
	 */
	ApiRange.prototype.SetNumberFormat = function (sFormat) {
		this.range.setNumFormat(sFormat);
	};
	Object.defineProperty(ApiRange.prototype, "NumberFormat", {
		get: function () {
			return this.GetNumberFormat();
		},
		set: function (sFormat) {
			return this.SetNumberFormat(sFormat);
		}
	});

	/**
	 * Sets the border to the cell / cell range with the parameters specified.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {BordersIndex} bordersIndex - Specifies the cell border position.
	 * @param {LineStyle} lineStyle - Specifies the line style used to form the cell border.
	 * @param {ApiColor} oColor - The color object which specifies the color to be set to the cell border.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetBorders.js
	 */
	ApiRange.prototype.SetBorders = function (bordersIndex, lineStyle, oColor) {
		var borders = new AscCommonExcel.Border();
		borders.initDefault();
		switch (bordersIndex) {
			case 'DiagonalDown':
				borders.dd = true;
				borders.d = private_MakeBorder(lineStyle, oColor);
				break;
			case 'DiagonalUp':
				borders.du = true;
				borders.d = private_MakeBorder(lineStyle, oColor);
				break;
			case 'Bottom':
				borders.b = private_MakeBorder(lineStyle, oColor);
				break;
			case 'Left':
				borders.l = private_MakeBorder(lineStyle, oColor);
				break;
			case 'Right':
				borders.r = private_MakeBorder(lineStyle, oColor);
				break;
			case 'Top':
				borders.t = private_MakeBorder(lineStyle, oColor);
				break;
			case 'InsideHorizontal':
				borders.ih = private_MakeBorder(lineStyle, oColor);
				break;
			case 'InsideVertical':
				borders.iv = private_MakeBorder(lineStyle, oColor);
				break;
		}
		this.range.setBorder(borders);
	};

	/**
	 * Merges the selected cell range into a single cell or a cell row.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isAcross - When set to <b>true</b>, the cells within the selected range will be merged along the rows,
	 * but remain split in the columns. When set to <b>false</b>, the whole selected range of cells will be merged into a single cell.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Merge.js
	 */
	ApiRange.prototype.Merge = function (isAcross) {
		if (isAcross) {
			var ws = this.range.worksheet;
			var bbox = this.range.getBBox0();
			for (var r = bbox.r1; r <= bbox.r2; ++r) {
				ws.getRange3(r, bbox.c1, r, bbox.c2).merge(null);
			}
		} else {
			this.range.merge(null);
		}
	};

	/**
	 * Splits the selected merged cell range into the single cells.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/UnMerge.js
	 */
	ApiRange.prototype.UnMerge = function () {
		this.range.unmerge();
	};

	/**
	 * Returns one cell or cells from the merge area.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null} - returns null if range isn't one cell
	 */
	Object.defineProperty(ApiRange.prototype, "MergeArea", {
		get: function () {
			let result = null;
			if (this.range.isOneCell()) {
				var bb = this.range.hasMerged();
				result = new ApiRange((bb) ? AscCommonExcel.Range.prototype.createFromBBox(this.range.worksheet, bb) : this.range);
			} else {
				logError(new Error('Range must be is one cell.'));
			}
			return result;
		}
	});

	/**
	 * Executes a provided function once for each cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {Function} fCallback - A function which will be executed for each cell.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/ForEach.js
	 */
	ApiRange.prototype.ForEach = function (fCallback) {
		if (fCallback instanceof Function) {
			var ws = this.range.getWorksheet();
			this.range._foreach(function (cell) {
				fCallback(new ApiRange(ws.getCell3(cell.nRow, cell.nCol)));
			});
		}
	};

	/**
	 * Adds a comment to the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sText - The comment text.
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment | null} - returns false if comment can't be added.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/AddComment.js
	 */
	ApiRange.prototype.AddComment = function (sText, sAuthor) {
		let result = null;
		let ws = Asc['editor'].wb.getWorksheet(this.range.getWorksheet().getIndex());
		let isValidData = typeof (sText) === 'string' && sText.trim() !== '';
		if (ws && isValidData) {
			var comment = new Asc.asc_CCommentData();
			comment.asc_putText(sText);
			let author = ((typeof (sAuthor) === 'string' && sAuthor.trim() !== '') ? sAuthor : Asc['editor'].User.asc_getUserName());
			comment.asc_putUserName(author);
			// todo проверить как в документа добавлются (надо ли выставлять этот параметр)
			// comment.asc_putUserId(Asc['editor'].User.asc_getId());
			comment.asc_putCol(this.range.bbox.c1);
			comment.asc_putRow(this.range.bbox.r1);
			comment.asc_putDocumentFlag(false);
			ws.cellCommentator.addComment(comment, true);
			// Asc['editor'].wb.Api.asc_addComment(comment);
			result = new ApiComment(comment, Asc['editor'].wb);
		}

		return result;
	};

	/**
	 * Returns the Worksheet object that represents the worksheet containing the specified range. It will be available in the read-only mode.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheet}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetWorksheet.js
	 */
	ApiRange.prototype.GetWorksheet = function () {
		return new ApiWorksheet(this.range.worksheet);
	};
	Object.defineProperty(ApiRange.prototype, "Worksheet", {
		get: function () {
			return this.GetWorksheet();
		}
	});

	/**
	 * Returns the ApiName object of the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiName}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetDefName.js
	 */
	ApiRange.prototype.GetDefName = function () {
		var defName = this.range.worksheet.getName() + "!" + this.range.bbox.getAbsName();
		var SheetId = this.range.worksheet.getId();
		defName = this.range.worksheet.workbook.findDefinesNames(defName, SheetId);
		if (defName) {
			defName = this.range.worksheet.workbook.getDefinesNames(defName, SheetId);
		}
		return new ApiName(defName);
	};
	Object.defineProperty(ApiRange.prototype, "DefName", {
		get: function () {
			return this.GetDefName();
		}
	});

	/**
	 * Returns the ApiComment object of the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiComment | null} - returns null if range does not consist of one cell.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetComment.js
	 */
	ApiRange.prototype.GetComment = function () {
		if (!this.range.isOneCell()) {
			return null;
		}
		var ws = this.range.worksheet.workbook.oApi.wb.getWorksheet(this.range.worksheet.getIndex());
		var comment = ws.cellCommentator.getComment(this.range.bbox.c1, this.range.bbox.r1, false);
		var res = comment ? new ApiComment(comment, this.range.worksheet.workbook.oApi.wb) : null;
		return res;
	};
	Object.defineProperty(ApiRange.prototype, "Comments", {
		get: function () {
			return this.GetComment();
		}
	});

	/**
	 * Selects the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Select.js
	 */
	ApiRange.prototype.Select = function () {
		if (this.range.worksheet.getId() === this.range.worksheet.workbook.getActiveWs().getId()) {
			var newSelection = new AscCommonExcel.SelectionRange(this.range.worksheet);
			let bbox = this.range.bbox;
			newSelection.assign2(bbox);
			if (this.areas) {
				this.areas.forEach(function (el) {
					if (!bbox.isEqual(el.bbox))
						newSelection.ranges.push(el.bbox);
				})
			}
			newSelection.Select();
		}
	};

	/**
	 * Returns the current range angle.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {Angle}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetOrientation.js
	 */
	ApiRange.prototype.GetOrientation = function () {
		return this.range.getAngle();
	};

	/**
	 * Sets an angle to the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {Angle} angle - Specifies the range angle.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetOrientation.js
	 */
	ApiRange.prototype.SetOrientation = function (angle) {
		switch (angle) {
			case 'xlDownward':
				angle = -90;
				break;
			case 'xlHorizontal':
				angle = 0;
				break;
			case 'xlUpward':
				angle = 90;
				break;
			case 'xlVertical':
				angle = 255;
				break;
		}
		this.range.setAngle(angle);
	};

	Object.defineProperty(ApiRange.prototype, "Orientation", {
		get: function () {
			return this.GetOrientation();
		},
		set: function () {
			return this.SetOrientation();
		}
	});

	/**
	 * Sorts the cells in the given range by the parameters specified in the request.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | String} key1 - First sort field.
	 * @param {SortOrder} sSortOrder1 - The sort order for the values specified in Key1.
	 * @param {ApiRange | String} key2 - Second sort field.
	 * @param {SortOrder} sSortOrder2 - The sort order for the values specified in Key2.
	 * @param {ApiRange | String} key3 - Third sort field.
	 * @param {SortOrder} sSortOrder3 - The sort order for the values specified in Key3.
	 * @param {SortHeader} sHeader - Specifies whether the first row contains header information.
	 * @param {SortOrientation} sOrientation - Specifies if the sort should be by row (default) or column.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/SetSort.js
	 */
	ApiRange.prototype.SetSort = function (key1, sSortOrder1, key2, /*Type,*/ sSortOrder2, key3, sSortOrder3, sHeader, /*OrderCustom, MatchCase,*/ sOrientation/*, SortMethod, DataOption1, DataOption2, DataOption3*/) {
		var ws = this.range.worksheet;
		var sortSettings = new Asc.CSortProperties(ws);
		var range = this.range.bbox;

		var aMerged = ws.mergeManager.get(range);
		if (aMerged.outer.length > 0 || (aMerged.inner.length > 0 && null == window['AscCommonExcel']._isSameSizeMerged(range, aMerged.inner, true))) {
			return;
		}

		sortSettings.hasHeaders = sHeader === "xlYes";
		var columnSort = sortSettings.columnSort = sOrientation !== "xlSortRows";

		var getSortLevel = function (_key, _order) {
			var index = null;
			if (_key instanceof ApiRange) {
				index = columnSort ? _key.range.bbox.c1 - range.c1 : _key.range.bbox.r1 - range.r1;
			} else if (typeof _key === "string") {
				//named range
				var _defName = ws.workbook.getDefinesNames(_key);
				if (_defName) {
					var defNameRef;
					AscCommonExcel.executeInR1C1Mode(false, function () {
						defNameRef = AscCommonExcel.getRangeByRef(_defName.ref, ws, true, true)
					});
					if (defNameRef && defNameRef[0] && defNameRef[0].worksheet) {
						if (range.contains(defNameRef[0].bbox.c1, defNameRef[0].bbox.r1)) {
							if (defNameRef[0].worksheet.Id === ws.Id) {
								index = columnSort ? defNameRef[0].bbox.c1 - range.c1 : defNameRef[0].bbox.r1 - range.r1;
							}
						} else {
							//error
							return false;
						}
					}
				}
			}

			if (null === index) {
				return null;
			}

			var level = new Asc.CSortPropertiesLevel();
			level.index = index;
			level.descending = _order === "xlDescending" ? Asc.c_oAscSortOptions.Descending : Asc.c_oAscSortOptions.Ascending;
			sortSettings.levels.push(level);
		};

		sortSettings.levels = [];
		if (key1 && false === getSortLevel(key1, sSortOrder1)) {
			return;
		}
		if (key2 && false === getSortLevel(key2, sSortOrder2)) {
			return;
		}
		if (key3 && false === getSortLevel(key3, sSortOrder3)) {
			return;
		}

		var oWorksheet = Asc['editor'].wb.getWorksheet();
		var tables = ws.autoFilters.getTablesIntersectionRange(range);
		var obj;
		if (tables && tables.length) {
			obj = tables[0];
		} else if (ws.AutoFilter && ws.AutoFilter.Ref && ws.AutoFilter.Ref.intersection(range)) {
			obj = ws.AutoFilter;
		}
		ws.setCustomSort(sortSettings, obj, null, oWorksheet && oWorksheet.cellCommentator, range);
	};

	/*Object.defineProperty(ApiRange.prototype, "Sort", {
		set: function (obj) {
			return this.SetSort(obj.Key1, obj.Order1, obj.Key2, obj.Type, obj.Order2, obj.Key3, obj.Order3, obj.Header,
				obj.OrderCustom, obj.MatchCase, obj.Orientation, obj.SortMethod, obj.DataOption1, obj.DataOption2,
				obj.DataOption3);
		}
	});*/

	/**
	 * Deletes the Range object.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {?DeleteShiftDirection} shift - Specifies how to shift cells to replace the deleted cells.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Delete.js
	 */
	ApiRange.prototype.Delete = function (shift) {
		let preDeleteAction = function() {
			cellCommentator.updateCommentsDependencies(false, val, checkRange);
			wsView.shiftCellWatches(false, val, bbox);
			wsView.model.shiftDataValidation(false, val, checkRange, true);
			wsView._cleanCache(lockRange);
		};
		let val;
		let ws = this.GetWorksheet().worksheet;
		let wsView = Asc['editor'].wb.getWorksheet(ws.getIndex());
		let cellCommentator = wsView.cellCommentator;
		let bbox = this.range.bbox;
		let checkRange = bbox.clone();
		let lockRange;
		if (shift && shift.toLocaleLowerCase) {
			shift = shift.toLocaleLowerCase();
		} else {
			let rows = bbox.r2 - bbox.r1 + 1;
			let cols = bbox.c2 - bbox.c1 + 1;
			shift = (rows <= cols) ? "up" : "left";
		}
		if (shift === "up") {
			val = Asc.c_oAscDeleteOptions.DeleteCellsAndShiftTop;
			lockRange = ws.getRange3(bbox.r1, bbox.c1, bbox.r2, AscCommon.gc_nMaxCol0);
			this.range.deleteCellsShiftUp(preDeleteAction);
		} else {
			val = Asc.c_oAscDeleteOptions.DeleteCellsAndShiftLeft;
			lockRange = ws.getRange3(bbox.r1, bbox.c1, AscCommon.gc_nMaxRow0, AscCommon.c2);
			this.range.deleteCellsShiftLeft(preDeleteAction);
		}
	};

	/**
	 * Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {?string} shift - Specifies which way to shift the cells ("right", "down").
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Insert.js
	 */
	ApiRange.prototype.Insert = function (shift) {
		if (shift && shift.toLocaleLowerCase) {
			shift = shift.toLocaleLowerCase();
		} else {
			var bbox = this.range.bbox;
			var rows = bbox.r2 - bbox.r1 + 1;
			var cols = bbox.c2 - bbox.c1 + 1;
			shift = (rows <= cols) ? "down" : "right";
		}
		if (shift == "down")
			this.range.addCellsShiftBottom();
		else
			this.range.addCellsShiftRight()
	};

	/**
	 * Changes the width of the columns or the height of the rows in the range to achieve the best fit.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {?bool} bRows - Specifies if the width of the columns will be autofit.
	 * @param {?bool} bCols - Specifies if the height of the rows will be autofit.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/AutoFit.js
	 */
	ApiRange.prototype.AutoFit = function (bRows, bCols) {
		var index = this.range.worksheet.getIndex();
		if (bRows)
			this.range.worksheet.workbook.oApi.wb.getWorksheet(index).autoFitRowHeight(this.range.bbox.r1, this.range.bbox.r2);

		for (var i = this.range.bbox.c1; i <= this.range.bbox.c2 && bCols; i++)
			this.range.worksheet.workbook.oApi.wb.getWorksheet(index).autoFitColumnsWidth(i);
	};

	/**
	 * Returns a collection of the ranges.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiAreas}
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetAreas.js
	 */
	ApiRange.prototype.GetAreas = function () {
		return new ApiAreas(this.areas || [this.range], this);
	};
	Object.defineProperty(ApiRange.prototype, "Areas", {
		get: function () {
			return this.GetAreas();
		}
	});

	/**
	 * Copies the range to the specified range or to the clipboard.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - Specifies the new range to which the specified range will be copied. If this argument is omitted, the range will be copied to the clipboard.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Copy.js
	 */
	ApiRange.prototype.Copy = function (destination) {
		var oApi = Asc["editor"];
		if (destination) {
			if (destination instanceof ApiRange) {
				let bboxFrom = this.range.bbox;
				let cols = bboxFrom.c2 - bboxFrom.c1;
				let rows = bboxFrom.r2 - bboxFrom.r1;
				let bbox = destination.range.bbox;
				let range = destination.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols));
				this.range.move(range.bbox, true, destination.range.worksheet);
				//AscCommon.g_clipboardBase && AscCommon.g_clipboardBase.ClearBuffer();
			} else {
				logError(new Error('Invalid destination'));
			}
		} else {
			let ws =  this.range.worksheet;
			private_executeOtherActiveSheet(ws, this.areas || [this.range], function () {
				oApi && oApi.asc_Copy();
			});
			oApi && oApi.wb.cleanCopyData();
		}
	};

	/**
	 * Cuts the range and save it to the clipboard or paste it to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - Specifies the new range to which the cut range will be pasted. If this argument is omitted, the range will be copied to the clipboard.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Cut.js
	 */
	ApiRange.prototype.Cut = function (destination) {
		var oApi = Asc["editor"];
		if (destination) {
			if (destination instanceof ApiRange) {
				let bboxFrom = this.range.bbox;
				let cols = bboxFrom.c2 - bboxFrom.c1;
				let rows = bboxFrom.r2 - bboxFrom.r1;
				let bbox = destination.range.bbox;
				let range = destination.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols));
				this.range.move(range.bbox, false, destination.range.worksheet);
				//AscCommon.g_clipboardBase && AscCommon.g_clipboardBase.ClearBuffer();
			} else {
				logError(new Error('Invalid destination'));
			}
		} else {
			let ws =  this.range.worksheet;
			private_executeOtherActiveSheet(ws, [this.range], function () {
				AscCommon.g_clipboardBase.forceCutSelection = true;
				oApi && oApi.asc_Cut();
				AscCommon.g_clipboardBase.forceCutSelection = false;
			});
			oApi && oApi.wb.cleanCutData();
		}
	};

	/**
	 * Pastes the Range object to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} rangeFrom - Specifies the range to be pasted to the current range
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Paste.js
	 */
	ApiRange.prototype.Paste = function (rangeFrom) {
		if (rangeFrom && rangeFrom instanceof ApiRange) {
			let bboxFrom = rangeFrom.range.bbox;
			let cols = bboxFrom.c2 - bboxFrom.c1;
			let rows = bboxFrom.r2 - bboxFrom.r1;
			let bbox = this.range.bbox;
			let range = this.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols));
			rangeFrom.range.move(range.bbox, true, range.worksheet);
		} else {
			logError(new Error('Invalid range'));
		}
	};

	/**
	 * Pastes the Range object to the specified range using the special paste options.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {PasteType} [sPasteType="xlPasteAll"]  - Paste option.
	 * @param {PasteSpecialOperation} [sPasteSpecialOperation="xlPasteSpecialOperationNone"] - The mathematical operation which will be applied to the copied data.
	 * @param {boolean} bSkipBlanks [bSkipBlanks=false] - Specifies whether to avoid replacing values in the paste area when blank cells occur in the copy area.
	 * @param {boolean} bTranspose [bTranspose=false] - Specifies whether the pasted data will be transposed from rows to columns.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/PasteSpecial.js
	 */
	ApiRange.prototype.PasteSpecial = function (sPasteType, sPasteSpecialOperation, bSkipBlanks, bTranspose) {
		if (sPasteType && typeof sPasteType !== 'string') {
			logError(new Error('Invalid type of parameter "sPasteType".'));
			return;
		}
		if (sPasteSpecialOperation && typeof sPasteSpecialOperation !== 'string') {
			logError(new Error('Invalid type of parameter "sPasteSpecialOperation".'));
			return;
		}


		let nPasteType = null;
		if (sPasteType) {
			switch (sPasteType) {
				// case "xlPasteAllMergingConditionalFormats":
				// 	break;
				case "xlPasteAll":
					break;
				case "xlPasteAllExceptBorders":
					nPasteType = Asc.c_oSpecialPasteProps.formulaWithoutBorders;
					break;
				//case "xlPasteAllUsingSourceTheme":
				//	break;
				case "xlPasteColumnWidths":
					nPasteType = Asc.c_oSpecialPasteProps.formulaColumnWidth;
					break;
				case "xlPasteComments":
					nPasteType = Asc.c_oSpecialPasteProps.comments;
					break;
				case "xlPasteFormats":
					nPasteType = Asc.c_oSpecialPasteProps.pasteOnlyFormating;
					break;
				case "xlPasteFormulas":
					nPasteType = Asc.c_oSpecialPasteProps.pasteOnlyFormula;
					break;
				case "xlPasteFormulasAndNumberFormats":
					nPasteType = Asc.c_oSpecialPasteProps.formulaNumberFormat;
					break;
				// case "xlPasteValidation":
				// 	nPasteType = Asc.c_oSpecialPasteProps.formulaColumnWidth;
				// 	break;
				case "xlPasteValues":
					nPasteType = Asc.c_oSpecialPasteProps.pasteOnlyValues;
					break;
				case "xlPasteValuesAndNumberFormats":
					nPasteType = Asc.c_oSpecialPasteProps.valueNumberFormat;
					break;
			}
		}

		let nPasteSpecialOperation = null;
		if (sPasteSpecialOperation) {
			switch (sPasteSpecialOperation) {
				case "xlPasteSpecialOperationAdd":
					nPasteSpecialOperation = Asc.c_oSpecialPasteOperation.add;
					break;
				case "xlPasteSpecialOperationDivide":
					nPasteSpecialOperation = Asc.c_oSpecialPasteOperation.divide;
					break;
				case "xlPasteSpecialOperationMultiply":
					nPasteSpecialOperation = Asc.c_oSpecialPasteOperation.multiply;
					break;
				case "xlPasteSpecialOperationNone":
					break;
				case "xlPasteSpecialOperationSubtract":
					nPasteSpecialOperation = Asc.c_oSpecialPasteOperation.subtract;
					break;
			}
		}

		let specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
		if (!specialPasteHelper.specialPasteProps) {
			specialPasteHelper.specialPasteProps = new Asc.SpecialPasteProps();
		}
		let specialPasteProps = specialPasteHelper.specialPasteProps;

		if (nPasteType != null) {
			specialPasteProps.asc_setProps(nPasteType);
		}
		if (nPasteSpecialOperation != null) {
			specialPasteProps.asc_setOperation(nPasteSpecialOperation);
		}
		specialPasteProps.asc_setSkipBlanks(!!bSkipBlanks);
		specialPasteProps.asc_setTranspose(!!bTranspose);

		let oApi = Asc["editor"];
		AscCommon.g_specialPasteHelper && AscCommon.g_specialPasteHelper.Special_Paste_Hide_Button();
		let ws =  this.range.worksheet;
		private_executeOtherActiveSheet(ws, this.areas || [this.range], function () {
			oApi && oApi.asc_Paste();
		});
	};

	/**
	 * Returns a PivotTable object that represents the PivotTable report containing the upper-left corner of the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotTable | null}
	 * @see office-js-api/Examples/Cell/ApiRange/Methods/GetPivotTable.js
	 */
	ApiRange.prototype.GetPivotTable = function() {
		var bbox = this.range.isOneCell() ?  this.range.bbox : this.range.worksheet.getRange3(this.range.bbox.r1, this.range.bbox.c1, this.range.bbox.r1, this.range.bbox.c1).bbox;
		var pivotTables = this.range.worksheet.pivotTables;
		var foundTable = null;
		for (var i = 0; i < pivotTables.length; i++) {
			var table = pivotTables[i];
			if (table.intersection(bbox)) {
				foundTable = new ApiPivotTable(table, this.range.worksheet.workbook.oApi);
				break;
			}
		}
		return foundTable;
	};

	Object.defineProperty(ApiRange.prototype, "PivotTable", {
		get: function () {
			return this.GetPivotTable();
		}
	});



	/**
	 * Search data type (formulas or values).
	 * @typedef {("xlFormulas" | "xlValues")} XlFindLookIn
	 * @see office-js-api/Examples/Enumerations/XlFindLookIn.js
	 */

	/**
	 * Specifies whether the whole search text or any part of the search text is matched.
	 * @typedef {("xlWhole" | "xlPart")} XlLookAt
	 * @see office-js-api/Examples/Enumerations/XlLookAt.js
	 */

	/**
	 * Range search order - by rows or by columns.
	 * @typedef {("xlByRows" | "xlByColumns")} XlSearchOrder
	 * @see office-js-api/Examples/Enumerations/XlSearchOrder.js
	 */

	/**
	 * Range search direction - next match or previous match.
	 * @typedef {("xlNext" | "xlPrevious")} XlSearchDirection
	 * @see office-js-api/Examples/Enumerations/XlSearchDirection.js
	 */

	/**
	 * Properties to make search.
	 * @typedef {Object} SearchData
	 * @property {string | undefined} What - The data to search for.
	 * @property {ApiRange} After - The cell after which you want the search to begin. If this argument is not specified, the search starts after the cell in the upper-left corner of the range.
	 * @property {XlFindLookIn} LookIn - Search data type (formulas or values).
	 * @property {XlLookAt} LookAt - Specifies whether the whole search text or any part of the search text is matched.
	 * @property {XlSearchOrder} SearchOrder - Range search order - by rows or by columns.
	 * @property {XlSearchDirection} SearchDirection - Range search direction - next match or previous match.
	 * @property {boolean} MatchCase - Case sensitive or not. The default value is "false".
	 * @see office-js-api/Examples/Enumerations/SearchData.js
	 */

	/**
	 * Properties to make search and replace.
	 * @typedef {Object} ReplaceData
	 * @property {string | undefined} What - The data to search for.
	 * @property {string} Replacement - The replacement string.
	 * @property {XlLookAt} LookAt - Specifies whether the whole search text or any part of the search text is matched.
	 * @property {XlSearchOrder} SearchOrder - Range search order - by rows or by columns.
	 * @property {XlSearchDirection} SearchDirection - Range search direction - next match or previous match.
	 * @property {boolean} MatchCase - Case sensitive or not. The default value is "false".
	 * @property {boolean} ReplaceAll - Specifies if all the found data will be replaced or not. The default value is "true".
	 * @see office-js-api/Examples/Enumerations/ReplaceData.js
	 */

	/**
	 * Finds specific information in the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {SearchData} oSearchData - The search data used to make search.
	 * @returns {ApiRange | null} - Returns null if the current range does not contain such text.
	 * @also
	 * Finds specific information in the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string | undefined} What - The data to search for.
	 * @param {ApiRange} After - The cell after which you want the search to begin. If this argument is not specified, the search starts after the cell in the upper-left corner of the range.
	 * @param {XlFindLookIn} LookIn - Search data type (formulas or values).
	 * @param {XlLookAt} LookAt - Specifies whether the whole search text or any part of the search text is matched.
	 * @param {XlSearchOrder} SearchOrder - Range search order - by rows or by columns.
	 * @param {XlSearchDirection} SearchDirection - Range search direction - next match or previous match.
	 * @param {boolean} MatchCase - Case sensitive or not. The default value is "false".
	 * @returns {ApiRange | null} - Returns null if the current range does not contain such text.
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Find.js
	 */
	ApiRange.prototype.Find = function (oSearchData) {
		let What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase;

		if (arguments.length === 1) {
			if (AscCommon.isRealObject(oSearchData)) {
				What = oSearchData['What'];
				After = oSearchData['After'];
				LookIn = oSearchData['LookIn'];
				LookAt = oSearchData['LookAt'];
				SearchOrder = oSearchData['SearchOrder'];
				SearchDirection = oSearchData['SearchDirection'];
				MatchCase = oSearchData['MatchCase'];
			} else {
				return null;
			}
		} else {
			What = arguments[0];
			After = arguments[1];
			LookIn = arguments[2];
			LookAt = arguments[3];
			SearchOrder = arguments[4];
			SearchDirection = arguments[5];
			MatchCase = arguments[6];
		}

		if (typeof What === 'string' || What === undefined) {
			let res = null;
			let options = new Asc.asc_CFindOptions();
			options.asc_setFindWhat(What);
			options.asc_setScanForward(SearchDirection != 'xlPrevious');
			MatchCase && options.asc_setIsMatchCase(MatchCase);
			options.asc_setIsWholeCell(LookAt === 'xlWhole');
			options.asc_setScanOnOnlySheet(Asc.c_oAscSearchBy.Range);
			options.asc_setSpecificRange(this["Address"]);
			options.asc_setScanByRows(SearchOrder === 'xlByRows');
			options.asc_setLookIn((LookIn === 'xlValues' ? 2 : 1));
			options.asc_setNotSearchEmptyCells(!(What === "" && !options.isWholeCell));
			let start = (After instanceof ApiRange && After.range.isOneCell() && this.range.containsRange(After.range))
				? {row: After.range.bbox.r1, col: After.range.bbox.c1}
				: {row: this.range.bbox.r1, col: this.range.bbox.c1};

			start.row += (options.scanByRows ? (options.scanForward ? 1 : -1) : 0);
			start.col += (!options.scanByRows ? (options.scanForward ? 1 : -1) : 0);
			options.asc_setActiveCell(start);
			let engine = this.range.worksheet.workbook.oApi.wb.Search(options);
			let id = this.range.worksheet.workbook.oApi.wb.GetSearchElementId(options.scanForward);
			if (id != null) {
				let elem = engine.Elements[id];
				res = new ApiRange(this.range.worksheet.getRange3(elem.row, elem.col, elem.row, elem.col));
			}
			this._searchOptions = options;
			return res;
		} else {
			logError(new Error('Invalid parametr "What".'));
			return null;
		}
	};

	/**
	 * Continues a search that was begun with the {@link ApiRange#Find} method. Finds the next cell that matches those same conditions and returns the ApiRange object that represents that cell. This does not affect the selection or the active cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} After - The cell after which the search will start. If this argument is not specified, the search starts from the last cell found.
	 * @returns {ApiRange | null} - Returns null if the range does not contain such text.
	 *
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/FindNext.js
	 */
	ApiRange.prototype.FindNext = function (After) {
		if (this._searchOptions) {
			let res = null;
			let activeCell;
			let engine;
			this._searchOptions.asc_setScanForward(true);
			if (After instanceof ApiRange && After.range.isOneCell() && this.range.containsRange(After.range)) {
				activeCell = {row: After.range.bbox.r1, col: After.range.bbox.c1};
				activeCell.row += (this._searchOptions.scanByRows ? 1 : 0);
				activeCell.col += (!this._searchOptions.scanByRows ? 1 : 0);
			} else {
				activeCell = {row: this.range.bbox.r1, col: this.range.bbox.c1};
			}
			if (JSON.stringify(this._searchOptions.activeCell) !== JSON.stringify(activeCell)) {
				this._searchOptions.asc_setActiveCell(activeCell);
			} else {
				engine = this.range.worksheet.workbook.oApi.wb.Search(this._searchOptions);
				engine.Reset();
			}
			engine = this.range.worksheet.workbook.oApi.wb.Search(this._searchOptions);
			let id = this.range.worksheet.workbook.oApi.wb.GetSearchElementId(true);
			if (id != null) {
				let elem = engine.Elements[id];
				res = new ApiRange(this.range.worksheet.getRange3(elem.row, elem.col, elem.row, elem.col));
			}
			return res;
		} else {
			logError(new Error('You should use "Find" method before this.'));
			return null;
		}
	};

	/**
	 * Continues a search that was begun with the {@link ApiRange#Find} method. Finds the previous cell that matches those same conditions and returns the ApiRange object that represents that cell. This does not affect the selection or the active cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} Before - The cell before which the search will start. If this argument is not specified, the search starts from the last cell found.
	 * @returns {ApiRange | null} - Returns null if the range does not contain such text.
	 *
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/FindPrevious.js
	 */
	ApiRange.prototype.FindPrevious = function (Before) {
		if (this._searchOptions) {
			let res = null;
			let activeCell;
			let engine;
			this._searchOptions.asc_setScanForward(false);
			if (Before instanceof ApiRange && Before.range.isOneCell() && this.range.containsRange(Before.range)) {
				activeCell = {row: Before.range.bbox.r1, col: Before.range.bbox.c1};
				activeCell.row += (this._searchOptions.scanByRows ? -1 : 0);
				activeCell.col += (!this._searchOptions.scanByRows ? -1 : 0);
			} else {
				activeCell = {row: this.range.bbox.r1, col: this.range.bbox.c1};
			}
			if (JSON.stringify(this._searchOptions.activeCell) !== JSON.stringify(activeCell)) {
				this._searchOptions.asc_setActiveCell(activeCell);
			} else {
				engine = this.range.worksheet.workbook.oApi.wb.Search(this._searchOptions);
				engine.Reset();
			}
			engine = this.range.worksheet.workbook.oApi.wb.Search(this._searchOptions);
			let id = this.range.worksheet.workbook.oApi.wb.GetSearchElementId(false);
			if (id != null) {
				let elem = engine.Elements[id];
				res = new ApiRange(this.range.worksheet.getRange3(elem.row, elem.col, elem.row, elem.col));
			}
			return res;
		} else {
			logError(new Error('You should use "Find" method before this.'));
			return null;
		}
	};

	/**
	 * Replaces specific information to another one in a range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ReplaceData} oReplaceData - The data used to make search and replace.
	 * @returns {ApiRange | null} - Returns null if the current range does not contain such text.
	 * @also
	 * Replaces specific information to another one in a range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string | undefined} What - The data to search for.
	 * @param {string} Replacement - The replacement string.
	 * @param {XlLookAt} LookAt - Specifies whether the whole search text or any part of the search text is matched.
	 * @param {XlSearchOrder} SearchOrder - Range search order - by rows or by columns.
	 * @param {XlSearchDirection} SearchDirection - Range search direction - next match or previous match.
	 * @param {boolean} MatchCase - Case sensitive or not. The default value is "false".
	 * @param {boolean} ReplaceAll - Specifies if all the found data will be replaced or not. The default value is "true".
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/Replace.js
	 */
	ApiRange.prototype.Replace = function (oReplaceData) {
		let What, Replacement, LookAt, SearchOrder, SearchDirection, MatchCase, ReplaceAll;

		if (arguments.length === 1) {
			if (AscCommon.isRealObject(oReplaceData)) {
				What = oReplaceData['What'];
				Replacement = oReplaceData['Replacement'];
				LookAt = oReplaceData['LookAt'];
				SearchOrder = oReplaceData['SearchOrder'];
				SearchDirection = oReplaceData['SearchDirection'];
				MatchCase = oReplaceData['MatchCase'];
				ReplaceAll = oReplaceData['ReplaceAll'];
			} else {
				return null;
			}
		} else {
			What = arguments[0];
			Replacement = arguments[1];
			LookAt = arguments[2];
			SearchOrder = arguments[3];
			SearchDirection = arguments[4];
			MatchCase = arguments[5];
			ReplaceAll = arguments[6];
		}

		if (typeof What === 'string' && typeof Replacement === 'string') {
			let options = new Asc.asc_CFindOptions();
			options.asc_setFindWhat(What);
			options.asc_setReplaceWith(Replacement);
			options.asc_setScanForward(SearchDirection != 'xlPrevious');
			MatchCase && options.asc_setIsMatchCase(MatchCase);
			options.asc_setIsWholeCell(LookAt === 'xlWhole');
			options.asc_setScanOnOnlySheet(Asc.c_oAscSearchBy.Range);
			options.asc_setSpecificRange(this["Address"]);
			options.asc_setScanByRows(SearchOrder === 'xlByRows');
			options.asc_setLookIn(Asc.c_oAscFindLookIn.Formulas);
			if (typeof ReplaceAll !== 'boolean')
				ReplaceAll = true;

			options.asc_setIsReplaceAll((ReplaceAll === true));
			this.range.worksheet.workbook.oApi.isReplaceAll = options.isReplaceAll;
			let engine = this.range.worksheet.workbook.oApi.wb.Search(options);
			engine.Reset();
			engine = this.range.worksheet.workbook.oApi.wb.Search(options);
			let id = this.range.worksheet.workbook.oApi.wb.GetSearchElementId(SearchDirection != 'xlPrevious');
			options.asc_setIsForMacros(true);
			if (id != null) {
				if (ReplaceAll)
					engine.SetCurrent(id);
				else
					this.range.worksheet.workbook.oApi.wb.SelectSearchElement(id);

				this.range.worksheet.workbook.oApi.wb.replaceCellText(options);
			}
		} else {
			logError(new Error('Invalid type of parametr "What" or "Replacement".'));
		}
	};

	/**
	 * Returns the ApiCharacters object that represents a range of characters within the object text. Use the ApiCharacters object to format characters within a text string.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} Start - The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.
	 * @param {number} Length - The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).
	 * @returns {ApiCharacters}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiRange/Methods/GetCharacters.js
	 */
	ApiRange.prototype.GetCharacters = function (Start, Length) {
		let options = {
			fragments: this.range.getValueForEdit2(),
			// user start
			uStart: Start,
			// user length
			uLength: Length,
			// real start
			start: (typeof Start !== "number" || Start < 1) ? 1 : Start,
			// user corrected length
			length: Length,
			// real length
			len: 0
		};

		options.len = AscCommonExcel.getFragmentsCharCodesLength(options.fragments);
		if (typeof Length !== "number" || options.len < (options.start + Length)) {
			options.length = options.len - options.start + 1;
		}

		return new ApiCharacters(options, this);
	};

	Object.defineProperty(ApiRange.prototype, "Characters", {
		get: function () {
			return this.GetCharacters();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDrawing
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiDrawing class.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CSE"]
	 * @returns {"drawing"}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetClassType.js
	 */
	ApiDrawing.prototype.GetClassType = function () {
		return "drawing";
	};

	/**
	 * Sets a size of the object (image, shape, chart) bounding box.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CSE"]
	 * @param {EMU} nWidth - The object width measured in English measure units.
	 * @param {EMU} nHeight - The object height measured in English measure units.
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetSize.js
	 */
	ApiDrawing.prototype.SetSize = function (nWidth, nHeight) {
		var fWidth = nWidth / 36000.0;
		var fHeight = nHeight / 36000.0;
		if (this.Drawing && this.Drawing.spPr && this.Drawing.spPr.xfrm) {
			this.Drawing.spPr.xfrm.setExtX(fWidth);
			this.Drawing.spPr.xfrm.setExtY(fHeight);
			this.Drawing.setDrawingBaseExt(fWidth, fHeight);

		}
	};

	/**
	 * Changes the position for the drawing object.
	 * <note>Please note that the horizontal and vertical offsets are calculated within the limits of
	 * the specified column and row cells only. If this value exceeds the cell width or height, another vertical/horizontal position will be set.</note>
	 * @memberof ApiDrawing
	 * @typeofeditors ["CSE"]
	 * @param {number} nFromCol - The number of the column where the beginning of the drawing object will be placed.
	 * @param {EMU} nColOffset - The offset from the nFromCol column to the left part of the drawing object measured in English measure units.
	 * @param {number} nFromRow - The number of the row where the beginning of the drawing object will be placed.
	 * @param {EMU} nRowOffset - The offset from the nFromRow row to the upper part of the drawing object measured in English measure units.
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetPosition.js
	 */
	ApiDrawing.prototype.SetPosition = function (nFromCol, nColOffset, nFromRow, nRowOffset) {
		var extX = null, extY = null;
		if (this.Drawing.drawingBase) {
			if (this.Drawing.drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorOneCell ||
				this.Drawing.drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorAbsolute) {
				extX = this.Drawing.drawingBase.ext.cx;
				extY = this.Drawing.drawingBase.ext.cy;
			}
		}
		if (!AscFormat.isRealNumber(extX) || !AscFormat.isRealNumber(extY)) {
			if (this.Drawing.spPr && this.Drawing.spPr.xfrm) {
				extX = this.Drawing.spPr.xfrm.extX;
				extY = this.Drawing.spPr.xfrm.extY;
			} else {
				extX = 5;
				extY = 5;
			}
		}
		this.Drawing.setDrawingBaseType(AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
		this.Drawing.setDrawingBaseCoords(nFromCol, nColOffset / 36000.0, nFromRow, nRowOffset / 36000.0, 0, 0, 0, 0, 0, 0, extX, extY);
	};

	/**
	 * Returns the width of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetWidth.js
	 */
	ApiDrawing.prototype.GetWidth = function () {
		return private_MM2EMU(this.Drawing.GetWidth());
	};
	/**
	 * Returns the height of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {EMU}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetHeight.js
	 */
	ApiDrawing.prototype.GetHeight = function () {
		return private_MM2EMU(this.Drawing.GetHeight());
	};
	/**
	 * Returns the lock value for the specified lock type of the current drawing.
	 * @typeofeditors ["CSE"]
	 * @param {DrawingLockType} sType - Lock type in the string format.
	 * @returns {bool}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/GetLockValue.js
	 */
	ApiDrawing.prototype.GetLockValue = function (sType) {
		var nLockType = private_GetDrawingLockType(sType);

		if (nLockType === -1)
			return false;

		if (this.Drawing)
			return this.Drawing.getLockValue(nLockType);

		return false;
	};

	/**
	 * Sets the lock value to the specified lock type of the current drawing.
	 * @typeofeditors ["CSE"]
	 * @param {DrawingLockType} sType - Lock type in the string format.
	 * @param {bool} bValue - Specifies if the specified lock is applied to the current drawing.
	 * @returns {bool}
	 * @see office-js-api/Examples/{Editor}/ApiDrawing/Methods/SetLockValue.js
	 */
	ApiDrawing.prototype.SetLockValue = function (sType, bValue) {
		var nLockType = private_GetDrawingLockType(sType);

		if (nLockType === -1)
			return false;

		if (this.Drawing) {
			this.Drawing.setLockValue(nLockType, bValue);
			return true;
		}


		return false;
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiImage
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiImage class.
	 * @memberof ApiImage
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {"image"}
	 * @see office-js-api/Examples/{Editor}/ApiImage/Methods/GetClassType.js
	 */
	ApiImage.prototype.GetClassType = function () {
		return "image";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiShape
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiShape class.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @returns {"shape"}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetClassType.js
	 */
	ApiShape.prototype.GetClassType = function () {
		return "shape";
	};

	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @returns {?ApiDocumentContent}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetContent.js
	 */
	ApiShape.prototype.GetContent = function () {
		var oApi = Asc["editor"];
		if (oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content) {
			return oApi.private_CreateApiDocContent(this.Drawing.txBody.content);
		}
		return null;
	};

	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @returns {?ApiDocumentContent}
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/GetDocContent.js
	 */
	ApiShape.prototype.GetDocContent = function () {
		var oApi = Asc["editor"];
		if (oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content) {
			return oApi.private_CreateApiDocContent(this.Drawing.txBody.content);
		}
		return null;
	};

	/**
	 * Sets the vertical alignment to the shape content where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @param {"top" | "center" | "bottom" } sVerticalAlign - The vertical alignment type for the shape inner contents.
	 * @returns {boolean} - returns false if shape or aligment doesn't exist.
	 * @see office-js-api/Examples/{Editor}/ApiShape/Methods/SetVerticalTextAlign.js
	 */
	ApiShape.prototype.SetVerticalTextAlign = function (sVerticalAlign) {
		if (this.Shape) {
			switch (sVerticalAlign) {
				case "top": {
					this.Shape.setVerticalAlign(4);
					break;
				}
				case "center": {
					this.Shape.setVerticalAlign(1);
					break;
				}
				case "bottom": {
					this.Shape.setVerticalAlign(0);
					break;
				}
				default:
					return false;
			}
			return true;
		}

		return false;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiChart
	//
	//------------------------------------------------------------------------------------------------------------------














	let ApiChart = AscBuilder.ApiChart;

	/**
	 * Sets values from the specified range to the specified series.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with series values. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column,
	 * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiChart/Methods/SetSeriaValues.js
	 */
	ApiChart.prototype.SetSeriaValues = function (sRange, nSeria) {
		return this.Chart.SetSeriaValues(sRange, nSeria);
	};

	/**
	 * Sets the x-axis values from the specified range to the specified series. It is used with the scatter charts only.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with series x-axis values. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column,
	 * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiChart/Methods/SetSeriaXValues.js
	 */
	ApiChart.prototype.SetSeriaXValues = function (sRange, nSeria) {
		return this.Chart.SetSeriaXValues(sRange, nSeria);
	};

	/**
	 * Sets a name to the specified series.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sNameRange - The series name. Can be a range of cells or usual text. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column,
	 * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiChart/Methods/SetSeriaName.js
	 */
	ApiChart.prototype.SetSeriaName = function (sNameRange, nSeria) {
		return this.Chart.SetSeriaName(sNameRange, nSeria);
	};

	/**
	 * Sets a range with the category values to the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with the category names. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column.
	 * @see office-js-api/Examples/{Editor}/ApiChart/Methods/SetCatFormula.js
	 */
	ApiChart.prototype.SetCatFormula = function (sRange) {
		return this.Chart.SetCatFormula(sRange);
	};

	/**
	 * Adds a new series to the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sNameRange - The series name. Can be a range of cells or usual text. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column,
	 * "Example series".
	 * @param {string} sValuesRange - A range of cells from the sheet with series values. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column.
	 * @param {string} [sXValuesRange=undefined] - A range of cells from the sheet with series x-axis values. It is used with the scatter charts only. For example:
	 * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * "A1:A5" - must be a single cell, row or column.
	 * @see office-js-api/Examples/{Editor}/ApiChart/Methods/AddSeria.js
	 */
	ApiChart.prototype.AddSeria = function (sNameRange, sValuesRange, sXValuesRange) {
		if (this.Chart.isScatterChartType() && typeof (sXValuesRange) === "string" && sXValuesRange !== "") {
			this.Chart.addScatterSeries(sNameRange, sXValuesRange, sValuesRange);
		} else
			this.Chart.addSeries(sNameRange, sValuesRange);
	};






	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiOleObject
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiOleObject class.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {"oleObject"}
	 * @see office-js-api/Examples/{Editor}/ApiOleObject/Methods/GetClassType.js
	 */
	ApiOleObject.prototype.GetClassType = function () {
		return "oleObject";
	};

	/**
	 * Sets the data to the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {string} sData - The OLE object string data.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiOleObject/Methods/SetData.js
	 */
	ApiOleObject.prototype.SetData = function (sData) {
		if (typeof (sData) !== "string" || sData === "")
			return false;

		this.Drawing.setData(sData);
		return true;
	};

	/**
	 * Returns the string data from the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiOleObject/Methods/GetData.js
	 */
	ApiOleObject.prototype.GetData = function () {
		if (typeof (this.Drawing.m_sData) === "string")
			return this.Drawing.m_sData;

		return "";
	};

	/**
	 * Sets the application ID to the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {string} sAppId - The application ID associated with the current OLE object.
	 * @returns {boolean}
	 * @see office-js-api/Examples/{Editor}/ApiOleObject/Methods/SetApplicationId.js
	 */
	ApiOleObject.prototype.SetApplicationId = function (sAppId) {
		if (typeof (sAppId) !== "string" || sAppId === "")
			return false;

		this.Drawing.setApplicationId(sAppId);
		return true;
	};

	/**
	 * Returns the application ID from the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiOleObject/Methods/GetApplicationId.js
	 */
	ApiOleObject.prototype.GetApplicationId = function () {
		if (typeof (this.Drawing.m_sApplicationId) === "string")
			return this.Drawing.m_sApplicationId;

		return "";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiColor class.
	 * @memberof ApiColor
	 * @typeofeditors ["CSE"]
	 * @returns {"color"}
	 * @see office-js-api/Examples/{Editor}/ApiColor/Methods/GetClassType.js
	 */
	ApiColor.prototype.GetClassType = function () {
		return "color";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiName
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiName class.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/GetName.js
	 */
	ApiName.prototype.GetName = function () {
		if (this.DefName) {
			return this.DefName.name
		} else {
			return this.DefName;
		}
	};

	/**
	 * Sets a string value representing the object name.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - New name for the range.
	 * @returns {boolean} - returns false if sName is invalid.
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/SetName.js
	 */
	ApiName.prototype.SetName = function (sName) {
		if (!sName || typeof sName !== 'string' || !this.DefName) {
			logError(new Error('Invalid name or Defname is undefined.'));
			return false;
		}
		var res = this.DefName.wb.checkDefName(sName);
		if (!res.status) {
			logError(new Error('Invalid name.')); // invalid name
			return false;
		}
		var oldName = this.DefName.getAscCDefName(false);
		var newName = this.DefName.getAscCDefName(false);
		newName.Name = sName;
		this.DefName.wb.editDefinesNames(oldName, newName);

		return true;
	};

	Object.defineProperty(ApiName.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (sName) {
			return this.SetName(sName);
		}
	});

	/**
	 * Deletes the DefName object.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/Delete.js
	 */
	ApiName.prototype.Delete = function () {
		this.DefName.wb.delDefinesNames(this.DefName.getAscCDefName(false));
	};

	/**
	 * Sets a formula that the name is defined to refer to.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @param {string} sRef    - The range reference which must contain the sheet name, followed by sign ! and a range of cells.
	 * Example: "Sheet1!$A$1:$B$2".
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/SetRefersTo.js
	 */
	ApiName.prototype.SetRefersTo = function (sRef) {
		this.DefName.setRef(sRef);
	};

	/**
	 * Returns a formula that the name is defined to refer to.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/GetRefersTo.js
	 */
	ApiName.prototype.GetRefersTo = function () {
		return (this.DefName) ? this.DefName.ref : this.DefName;
	};

	Object.defineProperty(ApiName.prototype, "RefersTo", {
		get: function () {
			return this.GetRefersTo();
		},
		set: function (sRef) {
			return this.SetRefersTo(sRef);
		}
	});

	/**
	 * Returns the ApiRange object by its name.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiName/Methods/GetRefersToRange.js
	 */
	ApiName.prototype.GetRefersToRange = function () {
		var range;
		if (this.DefName) {
			range = AscCommonExcel.getRangeByRef(this.DefName.ref, this.DefName.wb.getActiveWs(), true, true)[0];
		}
		return new ApiRange(range);
	};

	Object.defineProperty(ApiName.prototype, "RefersToRange", {
		get: function () {
			return this.GetRefersToRange();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiComment
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiComment class.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {"comment"}
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetClassType.js
	 */
	ApiComment.prototype.GetClassType = function () {
		return "comment";
	};

	/**
	 * Returns the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetText.js
	 */
	ApiComment.prototype.GetText = function () {
		return this.Comment.asc_getText();
	};

	/**
	 * Sets the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {string} text - New text for comment.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetText.js
	 */
	ApiComment.prototype.SetText = function (text) {
		if (typeof text === 'string' && text.trim() !== '') {
			this.Comment.asc_putText(text);
			this.private_OnChange();
		}
	};

	Object.defineProperty(ApiComment.prototype, "Text", {
		get: function () {
			return this.GetText();
		},
		set: function (text) {
			return this.SetText(text);
		}
	});

	/**
	 * Returns the current comment ID.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetId.js
	 */
	ApiComment.prototype.GetId = function () {
		return this.Comment.asc_getId();
	};

	Object.defineProperty(ApiComment.prototype, "Id", {
		get: function () {
			return this.GetId();
		}
	});

	/**
	 * Returns the comment author's name.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetAuthorName.js
	 */
	ApiComment.prototype.GetAuthorName = function () {
		return this.Comment.asc_getUserName();
	};

	/**
	 * Sets the comment author's name.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {string} sAuthorName - The comment author's name.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetAuthorName.js
	 */
	ApiComment.prototype.SetAuthorName = function (sAuthorName) {
		this.Comment.asc_putUserName(sAuthorName);
		this.private_OnChange();
	};

	Object.defineProperty(ApiComment.prototype, "AuthorName", {
		get: function () {
			return this.GetAuthorName();
		},
		set: function (sAuthorName) {
			return this.SetAuthorName(sAuthorName);
		}
	});

	/**
	 * Returns the user ID of the comment author.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetUserId.js
	 */
	ApiComment.prototype.GetUserId = function () {
		return this.Comment.asc_getUserId();
	};

	/**
	 * Sets the user ID to the comment author.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {string} sUserId - The user ID of the comment author.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetUserId.js
	 */
	ApiComment.prototype.SetUserId = function (sUserId) {
		this.Comment.asc_putUserId(sUserId);
		this.private_OnChange();
	};

	Object.defineProperty(ApiComment.prototype, "UserId", {
		get: function () {
			return this.GetUserId();
		},
		set: function (sUserId) {
			return this.SetUserId(sUserId);
		}
	});

	/**
	 * Checks if a comment is solved or not.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/IsSolved.js
	 */
	ApiComment.prototype.IsSolved = function () {
		return this.Comment.getSolved();
	};

	/**
	 * Marks a comment as solved.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bSolved - Specifies if a comment is solved or not.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetSolved.js
	 */
	ApiComment.prototype.SetSolved = function (bSolved) {
		this.Comment.setSolved(bSolved);
		this.private_OnChange();
	};

	Object.defineProperty(ApiComment.prototype, "Solved", {
		get: function () {
			return this.IsSolved();
		},
		set: function (bSolved) {
			return this.SetSolved(bSolved);
		}
	});

	/**
	 * Returns the timestamp of the comment creation in UTC format.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {Number}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetTimeUTC.js
	 */
	ApiComment.prototype.GetTimeUTC = function () {
		let nTime = parseInt(this.Comment.asc_getOnlyOfficeTime());
		if (isNaN(nTime))
			return 0;
		return nTime;
	};

	/**
	 * Sets the timestamp of the comment creation in UTC format.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment creation in UTC format.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetTimeUTC.js
	 */
	ApiComment.prototype.SetTimeUTC = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Comment.asc_putOnlyOfficeTime("0");
		else
			this.Comment.asc_putOnlyOfficeTime(String(nTime));

		this.private_OnChange();
	};

	Object.defineProperty(ApiComment.prototype, "TimeUTC", {
		get: function () {
			return this.GetTimeUTC();
		},
		set: function (timeStamp) {
			return this.SetTimeUTC(timeStamp);
		}
	});

	/**
	 * Returns the timestamp of the comment creation in the current time zone format.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {Number}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetTime.js
	 */
	ApiComment.prototype.GetTime = function () {
		let nTime = parseInt(this.Comment.asc_getTime());
		if (isNaN(nTime))
			return 0;
		return nTime;
	};

	/**
	 * Sets the timestamp of the comment creation in the current time zone format.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment creation in the current time zone format.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/SetTime.js
	 */
	ApiComment.prototype.SetTime = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Comment.asc_putTime("0");
		else
			this.Comment.asc_putTime(String(nTime));

		this.private_OnChange();
	};

	Object.defineProperty(ApiComment.prototype, "Time", {
		get: function () {
			return this.GetTime();
		},
		set: function (timeStamp) {
			return this.SetTime(timeStamp);
		}
	});

	/**
	 * Returns the quote text of the current comment.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {String | null}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetQuoteText.js
	 */
	ApiComment.prototype.GetQuoteText = function () {
		let text = null;
		let ws = this.WB.getWorksheetById(this.Comment.wsId);
		if (!this.Comment.asc_getDocumentFlag() && ws)
			text = ws._getRange(this.Comment.nCol, this.Comment.nRow, this.Comment.nCol, this.Comment.nRow).getValue();

		return text;
	};

	Object.defineProperty(ApiComment.prototype, "QuoteText", {
		get: function () {
			return this.GetQuoteText();
		}
	});

	/**
	 * Returns a number of the comment replies.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {Number?}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetRepliesCount.js
	 */
	ApiComment.prototype.GetRepliesCount = function () {
		return this.Comment.asc_getRepliesCount()
	};

	Object.defineProperty(ApiComment.prototype, "RepliesCount", {
		get: function () {
			return this.GetRepliesCount();
		}
	});

	/**
	 * Returns the specified comment reply.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {Number} [nIndex = 0] - The comment reply index.
	 * @returns {ApiCommentReply?}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/GetReply.js
	 */
	ApiComment.prototype.GetReply = function (nIndex) {
		if (typeof (nIndex) != "number" || nIndex < 0 || nIndex >= this.GetRepliesCount())
			nIndex = 0;

		let oReply = this.Comment.asc_getReply(nIndex);
		if (!oReply)
			return null;

		return new ApiCommentReply(this, oReply);
	};

	/**
	 * Adds a reply to a comment.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {String} sText - The comment reply text (required).
	 * @param {String} sAuthorName - The name of the comment reply author (optional).
	 * @param {String} sUserId - The user ID of the comment reply author (optional).
	 * @param {Number} [nPos=this.GetRepliesCount()] - The comment reply position.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/AddReply.js
	 */
	ApiComment.prototype.AddReply = function (sText, sAuthorName, sUserId, nPos) {
		if (typeof (sText) !== "string" || sText.trim() === "")
			return null;

		if (typeof (nPos) !== "number" || nPos < 0 || nPos > this.GetRepliesCount())
			nPos = this.GetRepliesCount();

		let oReply = new Asc.asc_CCommentData();
		oReply.asc_putText(sText);

		if (typeof (sAuthorName) === "string" && sAuthorName !== "")
			oReply.asc_putUserName(sAuthorName);

		if (sUserId != undefined && typeof (sUserId) === "string" && sUserId.trim() !== "")
			oReply.asc_putUserId(sUserId);

		this.Comment.aReplies.splice(nPos, 0, oReply);
		this.private_OnChange();
	};

	/**
	 * Removes the specified comment replies.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @param {Number} [nPos = 0] - The position of the first comment reply to remove.
	 * @param {Number} [nCount = 1] - A number of comment replies to remove.
	 * @param {boolean} [bRemoveAll = false] - Specifies whether to remove all comment replies or not.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/RemoveReplies.js
	 */
	ApiComment.prototype.RemoveReplies = function (nPos, nCount, bRemoveAll) {
		if (typeof (nPos) !== "number" || nPos < 0 || nPos > this.GetRepliesCount())
			nPos = 0;

		if (typeof (nCount) !== "number" || nCount < 0)
			nCount = 1;

		if (typeof (bRemoveAll) !== "boolean")
			bRemoveAll = false;

		if (bRemoveAll) {
			nPos = 0
			nCount = this.GetRepliesCount();
		}

		this.Comment.aReplies.splice(nPos, nCount);
		this.private_OnChange();
	};

	/**
	 * Deletes the ApiComment object.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/{Editor}/ApiComment/Methods/Delete.js
	 */
	ApiComment.prototype.Delete = function () {
		this.WB.Api.asc_removeComment(this.Comment.asc_getId());
	};

	ApiComment.prototype.private_OnChange = function () {
		this.WB.Api.asc_changeComment(this.Comment.asc_getId(), this.Comment);
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCommentReply
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiCommentReply class.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {"commentReply"}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetClassType.js
	 */
	ApiCommentReply.prototype.GetClassType = function () {
		return "commentReply";
	};

	/**
	 * Returns the comment reply text.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetText.js
	 */
	ApiCommentReply.prototype.GetText = function () {
		return this.Data.asc_getText();
	};

	/**
	 * Sets the comment reply text.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @param {string} sText - The comment reply text.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/SetText.js
	 */
	ApiCommentReply.prototype.SetText = function (sText) {
		this.Data.asc_putText(sText);
		this.private_OnChange();
	};

	Object.defineProperty(ApiCommentReply.prototype, "Text", {
		get: function () {
			return this.GetText();
		},
		set: function (sText) {
			return this.SetText(sText);
		}
	});

	/**
	 * Returns the comment reply author's name.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetAuthorName.js
	 */
	ApiCommentReply.prototype.GetAuthorName = function () {
		return this.Data.asc_getUserName();
	};

	/**
	 * Sets the comment reply author's name.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @param {string} sAuthorName - The comment reply author's name.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/SetAuthorName.js
	 */
	ApiCommentReply.prototype.SetAuthorName = function (sAuthorName) {
		this.Data.asc_putUserName(sAuthorName);
		this.private_OnChange();
	};

	Object.defineProperty(ApiCommentReply.prototype, "AuthorName", {
		get: function () {
			return this.GetAuthorName();
		},
		set: function (sAuthorName) {
			return this.SetAuthorName(sAuthorName);
		}
	});

	/**
	 * Returns the user ID of the comment reply author.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetUserId.js
	 */
	ApiCommentReply.prototype.GetUserId = function () {
		return this.Data.asc_getUserId();
	};

	/**
	 * Sets the user ID to the comment reply author.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @param {string} sUserId - The user ID of the comment reply author.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/SetUserId.js
	 */
	ApiCommentReply.prototype.SetUserId = function (sUserId) {
		this.Data.asc_putUserId(sUserId);
		this.private_OnChange();
	};

	Object.defineProperty(ApiCommentReply.prototype, "UserId", {
		get: function () {
			return this.GetUserId();
		},
		set: function (sUserId) {
			return this.SetUserId(sUserId);
		}
	});

	/**
	 * Returns the timestamp of the comment reply creation in UTC format.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {Number}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetTimeUTC.js
	 */
	ApiCommentReply.prototype.GetTimeUTC = function () {
		let nTime = parseInt(this.Data.asc_getOnlyOfficeTime());
		if (isNaN(nTime))
			return 0;
		return nTime;
	};

	/**
	 * Sets the timestamp of the comment reply creation in UTC format.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment reply creation in UTC format.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/SetTimeUTC.js
	 */
	ApiCommentReply.prototype.SetTimeUTC = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Data.asc_putOnlyOfficeTime("0");
		else
			this.Data.asc_putOnlyOfficeTime(String(nTime));

		this.private_OnChange();
	};

	Object.defineProperty(ApiCommentReply.prototype, "TimeUTC", {
		get: function () {
			return this.GetTimeUTC();
		},
		set: function (timeStamp) {
			return this.SetTimeUTC(timeStamp);
		}
	});

	/**
	 * Returns the timestamp of the comment reply creation in the current time zone format.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @returns {Number}
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/GetTime.js
	 */
	ApiCommentReply.prototype.GetTime = function () {
		let nTime = parseInt(this.Data.asc_getTime());
		if (isNaN(nTime))
			return 0;
		return nTime;
	};

	/**
	 * Sets the timestamp of the comment reply creation in the current time zone format.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CSE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment reply creation in the current time zone format.
	 * @since 7.5.0
	 * @see office-js-api/Examples/{Editor}/ApiCommentReply/Methods/SetTime.js
	 */
	ApiCommentReply.prototype.SetTime = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Data.asc_putTime("0");
		else
			this.Data.asc_putTime(String(nTime));

		this.private_OnChange();
	};

	Object.defineProperty(ApiCommentReply.prototype, "Time", {
		get: function () {
			return this.GetTime();
		},
		set: function (timeStamp) {
			return this.SetTime(timeStamp);
		}
	});

	ApiCommentReply.prototype.private_OnChange = function () {
		this.Comment.private_OnChange();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiAreas
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a value that represents the number of objects in the collection.
	 * @memberof ApiAreas
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiAreas/Methods/GetCount.js
	 */
	ApiAreas.prototype.GetCount = function () {
		return this.Items.length;
	};

	Object.defineProperty(ApiAreas.prototype, "Count", {
		get: function () {
			return this.GetCount();
		}
	});

	/**
	 * Returns a single object from a collection by its ID.
	 * @memberof ApiAreas
	 * @typeofeditors ["CSE"]
	 * @param {number} ind - The index number of the object.
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/{Editor}/ApiAreas/Methods/GetItem.js
	 */
	ApiAreas.prototype.GetItem = function (ind) {
		return this.Items[ind - 1] || null;
	};

	/**
	 * Returns the parent object for the specified collection.
	 * @memberof ApiAreas
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/{Editor}/ApiAreas/Methods/GetParent.js
	 */
	ApiAreas.prototype.GetParent = function () {
		return this._parent;
	};

	Object.defineProperty(ApiAreas.prototype, "Parent", {
		get: function () {
			return this.GetParent();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCharacters
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a value that represents a number of objects in the collection.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/GetCount.js
	 */
	ApiCharacters.prototype.GetCount = function () {
		return this._options.length < 0 ? 0 : this._options.length;
	};

	Object.defineProperty(ApiCharacters.prototype, "Count", {
		get: function () {
			return this.GetCount();
		}
	});

	/**
	 * Returns the parent object of the specified characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/GetParent.js
	 */
	ApiCharacters.prototype.GetParent = function () {
		return this._parent;
	};

	Object.defineProperty(ApiCharacters.prototype, "Parent", {
		get: function () {
			return this.GetParent();
		}
	});

	/**
	 * Deletes the ApiCharacters object.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/Delete.js
	 */
	ApiCharacters.prototype.Delete = function () {
		if (this._options.start <= this._options.len) {
			let editor = this._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let begin = this._options.start - 1;
			if (begin > this._options.len) begin = this._options.len - 1;
			let end = this._options.start + this._options.length - 1;
			if (end > this._options.len) end = this._options.len - 1;
			let fragments = this._options.fragments;
			let first = editor._findFragment(begin, fragments);
			let last = editor._findFragment(end, fragments);
			if (first && last) {
				if (first.index === last.index) {
					let codes = fragments[first.index].getCharCodes();
					fragments[first.index].setCharCodes(codes.slice(0, begin - first.begin).concat(codes.slice(end - first.begin)));
				} else {
					fragments[first.index].setCharCodes(fragments[first.index].getCharCodes().slice(0, begin - first.begin));
					fragments[last.index].setCharCodes(fragments[last.index].getCharCodes().slice(end - last.begin));
					let len = last.index - first.index;
					if (len > 1) {
						fragments.splice(first.index + 1, len - 1);
					}
				}
				editor._mergeFragments(fragments);
				let range = this._parent.range.worksheet.getRange3(this._parent.range.bbox.r1, this._parent.range.bbox.c1, this._parent.range.bbox.r1, this._parent.range.bbox.c1);
				range.setValue2(fragments);
				this._options = this._parent.GetCharacters(this._options.uStart, this._options.uLength)._options;
			}
		}
	};

	/**
	 * Inserts a string replacing the specified characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @param {string} String - The string to insert.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/Insert.js
	 */
	ApiCharacters.prototype.Insert = function (String) {
		this.Delete();
		let begin = this._options.start - 1;
		if (begin > this._options.len) begin = this._options.len - 1;
		let end = this._options.start + this._options.length - 1;
		if (end > this._options.len) end = this._options.len - 1;
		let fragments = this._options.fragments;
		let editor = this._parent.range.worksheet.workbook.oApi.wb.cellEditor;
		let copyFragment = editor._findFragmentToInsertInto((begin < this._options.len ? begin + 1 : begin), fragments);
		let textFormat = fragments[copyFragment.index].format.clone();

		String = AscCommon.convertUTF16toUnicode(String);
		let length = String.length;
		if (length) {
			// limit count characters
			let excess = AscCommonExcel.getFragmentsCharCodesLength(fragments) + length - Asc.c_oAscMaxCellOrCommentLength;
			if (0 > excess) excess = 0;

			if (excess) {
				length -= excess;
				if (!length) {
					logError(new Error('Max symbols in one cell.'));
					return;
				}
				String = String.slice(0, length);
			}

			let pos = this._options.start <= this._options.len ? begin : this._options.len;
			let fr;
			if (textFormat) {
				let newFr = new AscCommonExcel.Fragment({format: textFormat, charCodes: String});
				fr = editor._findFragment(pos, fragments);
				if (fr && pos < fr.end) {
					editor._splitFragment(fr, pos, fragments);
					fr = editor._findFragment(pos, fragments);
					Array.prototype.splice.apply(fragments, [fr.index, 0].concat(newFr));
				} else {
					fragments = fragments.concat(newFr);
				}
				editor._mergeFragments(fragments);
			} else {
				fr = editor._findFragmentToInsertInto(pos);
				if (fr) {
					let len = pos - fr.begin;
					let codes = fragments[fr.index].getCharCodes();
					fragments[fr.index].setCharCodes(codes.slice(0, len).concat(String).concat(codes.slice(len)));
					codes = fragments[fr.index].getCharCodes();
				}
			}
			let range = this._parent.range.worksheet.getRange3(this._parent.range.bbox.r1, this._parent.range.bbox.c1, this._parent.range.bbox.r1, this._parent.range.bbox.c1);
			range.setValue2(fragments);
			this._options = this._parent.GetCharacters(this._options.uStart, this._options.uLength)._options;
		}
	};

	/**
	 * Sets a string value that represents the text of the specified range of characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @param {string} Caption - A string value that represents the text of the specified range of characters.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/SetCaption.js
	 */
	ApiCharacters.prototype.SetCaption = function (Caption) {
		this.Insert(Caption);
	};

	/**
	 * Returns a string value that represents the text of the specified range of characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @returns {string} - A string value that represents the text of the specified range of characters.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/GetCaption.js
	 */
	ApiCharacters.prototype.GetCaption = function () {
		let value = this._parent.range.getValue();
		let begin = this._options.start - 1;
		let end = this._options.start + this._options.length - 1;
		let str = value.slice(begin, end);
		return str;
	};

	Object.defineProperty(ApiCharacters.prototype, "Caption", {
		get: function () {
			return this.GetCaption();
		},
		set: function (Caption) {
			return this.SetCaption(Caption);
		}
	});

	/**
	 * Sets the text for the specified characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @param {string} Text - The text to be set.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/SetText.js
	 */
	ApiCharacters.prototype.SetText = function (Text) {
		this.Insert(Text)
	};

	/**
	 * Returns the text of the specified range of characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @returns {string} - The text of the specified range of characters.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/GetText.js
	 */
	ApiCharacters.prototype.GetText = function () {
		return this.GetCaption();
	};

	Object.defineProperty(ApiCharacters.prototype, "Text", {
		get: function () {
			return this.GetText();
		},
		set: function (Text) {
			return this.SetText(Text);
		}
	});

	/**
	 * Returns the ApiFont object that represents the font of the specified characters.
	 * @memberof ApiCharacters
	 * @typeofeditors ["CSE"]
	 * @returns {ApiFont}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiCharacters/Methods/GetFont.js
	 */
	ApiCharacters.prototype.GetFont = function () {
		return new ApiFont(this);
	};

	Object.defineProperty(ApiCharacters.prototype, "Font", {
		get: function () {
			return this.GetFont();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFont
	//
	//------------------------------------------------------------------------------------------------------------------


	/**
	 * Returns the parent ApiCharacters object of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {ApiCharacters} - The parent ApiCharacters object.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetParent.js
	 */
	ApiFont.prototype.GetParent = function () {
		return this._object;
	};

	Object.defineProperty(ApiFont.prototype, "Parent", {
		get: function () {
			return this.GetParent();
		}
	});

	/**
	 * Returns the bold property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {boolean | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetBold.js
	 */
	ApiFont.prototype.GetBold = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let isBold = true;
			if (first && last) {
				for (let i = first.index; i <= last.index; i++) {
					if (!opt.fragments[i].format.getBold()) {
						isBold = null;
						break;
					}
				}
			} else {
				isBold = null;
			}
			return isBold;
		}
	};

	/**
	 * Sets the bold property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isBold - Specifies that the text characters are displayed bold.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetBold.js
	 */
	ApiFont.prototype.SetBold = function (isBold) {
		if (typeof isBold !== 'boolean') {
			logError(new Error('Invalid type of parametr "isBold".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "b", isBold);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Bold", {
		get: function () {
			return this.GetBold();
		},
		set: function (isBold) {
			return this.SetBold(isBold);
		}
	});

	/**
	 * Returns the italic property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {boolean | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetItalic.js
	 */
	ApiFont.prototype.GetItalic = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let isItalic = true;
			if (first && last) {
				for (let i = first.index; i <= last.index; i++) {
					if (!opt.fragments[i].format.getItalic()) {
						isItalic = null;
						break;
					}
				}
			} else {
				isItalic = null;
			}
			return isItalic;
		}
	};

	/**
	 * Sets the italic property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isItalic - Specifies that the text characters are displayed italic.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetItalic.js
	 */
	ApiFont.prototype.SetItalic = function (isItalic) {
		if (typeof isItalic !== 'boolean') {
			logError(new Error('Invalid type of parametr "isItalic".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "i", isItalic);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Italic", {
		get: function () {
			return this.GetItalic();
		},
		set: function (isItalic) {
			return this.SetItalic(isItalic);
		}
	});

	/**
	 * Returns the font size property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {number | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetSize.js
	 */
	ApiFont.prototype.GetSize = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let size = opt.fragments[first.index].format.getSize();
			if (first && last) {
				for (let i = first.index + 1; i <= last.index; i++) {
					if (size !== opt.fragments[i].format.getSize()) {
						size = null;
						break;
					}
				}
			} else {
				size = null;
			}
			return size;
		}
	};

	/**
	 * Sets the font size property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {number} Size - Font size.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetSize.js
	 */
	ApiFont.prototype.SetSize = function (Size) {
		if (typeof Size !== 'number' || Size < 0 || Size > 409) {
			logError(new Error('Invalid type of parametr "Size".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "fs", Size);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Size", {
		get: function () {
			return this.GetSize();
		},
		set: function (Size) {
			return this.SetSize(Size);
		}
	});

	/**
	 * Returns the strikethrough property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {boolean | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetStrikethrough.js
	 */
	ApiFont.prototype.GetStrikethrough = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let isStrikethrough = true;
			if (first && last) {
				for (let i = first.index; i <= last.index; i++) {
					if (!opt.fragments[i].format.getStrikeout()) {
						isStrikethrough = null;
						break;
					}
				}
			} else {
				isStrikethrough = null;
			}
			return isStrikethrough;
		}
	};

	/**
	 * Sets the strikethrough property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isStrikethrough - Specifies that the text characters are displayed strikethrough.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetStrikethrough.js
	 */
	ApiFont.prototype.SetStrikethrough = function (isStrikethrough) {
		if (typeof isStrikethrough !== 'boolean') {
			logError(new Error('Invalid type of parametr "isStrikethrough".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "s", isStrikethrough);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Strikethrough", {
		get: function () {
			return this.GetStrikethrough();
		},
		set: function (isStrikethrough) {
			return this.SetStrikethrough(isStrikethrough);
		}
	});

	/**
	 * Underline type.
	 * @typedef {("xlUnderlineStyleDouble" | "xlUnderlineStyleDoubleAccounting" | "xlUnderlineStyleNone" | "xlUnderlineStyleSingle" | "xlUnderlineStyleSingleAccounting")} XlUnderlineStyle
	 * @see office-js-api/Examples/Enumerations/XlUnderlineStyle.js
	 */

	/**
	 * Returns the type of underline applied to the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {XlUnderlineStyle | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetUnderline.js
	 */
	ApiFont.prototype.GetUnderline = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let Underline = opt.fragments[first.index].format.getUnderline();
			if (first && last) {
				for (let i = first.index + 1; i <= last.index; i++) {
					if (Underline !== opt.fragments[i].format.getUnderline()) {
						Underline = null;
						break;
					}
				}
			} else {
				Underline = null;
			}

			switch (Underline) {
				case Asc.EUnderline.underlineDouble:
					// todo doesn't work
					Underline = "xlUnderlineStyleDouble";
					break;
				case Asc.EUnderline.underlineDoubleAccounting:
					// todo doesn't work
					Underline = "xlUnderlineStyleDoubleAccounting";
					break;
				case Asc.EUnderline.underlineNone:
					Underline = "xlUnderlineStyleNone";
					break;
				case Asc.EUnderline.underlineSingle:
					Underline = "xlUnderlineStyleSingle";
					break;
				case Asc.EUnderline.underlineSingleAccounting:
					// todo doesn't work
					Underline = "xlUnderlineStyleSingleAccounting";
					break;

				default:
					Underline = null;
					break;
			}

			return Underline;
		}
	};

	/**
	 * Sets an underline of the type specified in the request to the current font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {XlUnderlineStyle} Underline - Underline type.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetUnderline.js
	 */
	ApiFont.prototype.SetUnderline = function (Underline) {
		if (typeof Underline !== 'string') {
			logError(new Error('Invalid type of parametr "isUnderline".'));
			return;
		}
		switch (Underline) {
			case "xlUnderlineStyleDouble":
				// todo doesn't work
				Underline = Asc.EUnderline.underlineDouble;
				break;
			case "xlUnderlineStyleDoubleAccounting":
				// todo doesn't work
				Underline = Asc.EUnderline.underlineDoubleAccounting;
				break;
			case "xlUnderlineStyleNone":
				Underline = Asc.EUnderline.underlineNone;
				break;
			case "xlUnderlineStyleSingle":
				Underline = Asc.EUnderline.underlineSingle;
				break;
			case "xlUnderlineStyleSingleAccounting":
				// todo doesn't work
				Underline = Asc.EUnderline.underlineSingleAccounting;
				break;

			default:
				Underline = Asc.EUnderline.underlineNone;
				break;
		}

		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "u", Underline);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Underline", {
		get: function () {
			return this.GetUnderline();
		},
		set: function (Underline) {
			return this.SetUnderline(Underline);
		}
	});

	/**
	 * Returns the subscript property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {boolean | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetSubscript.js
	 */
	ApiFont.prototype.GetSubscript = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let isSubscript = true;
			if (first && last) {
				for (let i = first.index; i <= last.index; i++) {
					if (AscCommon.vertalign_SubScript !== opt.fragments[i].format.getVerticalAlign()) {
						isSubscript = null;
						break;
					}
				}
			} else {
				isSubscript = null;
			}
			return isSubscript;
		}
	};

	/**
	 * Sets the subscript property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isSubscript - Specifies that the text characters are displayed subscript.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetSubscript.js
	 */
	ApiFont.prototype.SetSubscript = function (isSubscript) {
		if (typeof isSubscript !== 'boolean') {
			logError(new Error('Invalid type of parametr "isSubscript".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "fa", (isSubscript ? AscCommon.vertalign_SubScript : AscCommon.vertalign_Baseline));
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Subscript", {
		get: function () {
			return this.GetSubscript();
		},
		set: function (isSubscript) {
			return this.SetSubscript(isSubscript);
		}
	});

	/**
	 * Returns the superscript property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {boolean | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetSuperscript.js
	 */
	ApiFont.prototype.GetSuperscript = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let isSuperscript = true;
			if (first && last) {
				for (let i = first.index; i <= last.index; i++) {
					if (AscCommon.vertalign_SuperScript !== opt.fragments[i].format.getVerticalAlign()) {
						isSuperscript = null;
						break;
					}
				}
			} else {
				isSuperscript = null;
			}
			return isSuperscript;
		}
	};

	/**
	 * Sets the superscript property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isSuperscript - Specifies that the text characters are displayed superscript.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetSuperscript.js
	 */
	ApiFont.prototype.SetSuperscript = function (isSuperscript) {
		if (typeof isSuperscript !== 'boolean') {
			logError(new Error('Invalid type of parametr "isSuperscript".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "fa", (isSuperscript ? AscCommon.vertalign_SuperScript : AscCommon.vertalign_Baseline));
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Superscript", {
		get: function () {
			return this.GetSuperscript();
		},
		set: function (isSuperscript) {
			return this.SetSuperscript(isSuperscript);
		}
	});

	/**
	 * Returns the font name property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetName.js
	 */
	ApiFont.prototype.GetName = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let FontName = opt.fragments[first.index].format.getName();
			if (first && last) {
				for (let i = first.index + 1; i <= last.index; i++) {
					if (FontName !== opt.fragments[i].format.getName()) {
						FontName = null;
						break;
					}
				}
			} else {
				FontName = null;
			}
			return FontName;
		}
	};

	/**
	 * Sets the font name property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {string} FontName - Font name.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetName.js
	 */
	ApiFont.prototype.SetName = function (FontName) {
		if (typeof FontName !== 'string') {
			logError(new Error('Invalid type of parametr "FontName".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			// todo ms 19 allows to set any string, but maybe we shoud check it before set
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "fn", FontName);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (FontName) {
			return this.SetName(FontName);
		}
	});

	/**
	 * Returns the font color property of the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {ApiColor | null}
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/GetColor.js
	 */
	ApiFont.prototype.GetColor = function () {
		if (this._object instanceof ApiCharacters) {
			let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
			let opt = this._object._options;
			let begin = opt.start - 1;
			if (begin > opt.len) begin = opt.len - 1;
			let end = opt.start + opt.length - 1;
			if (end > opt.len) end = opt.len - 1;
			let first = editor._findFragment(begin, opt.fragments);
			let last = editor._findFragmentToInsertInto(end, opt.fragments);
			let color = opt.fragments[first.index].format.getColor();
			if (first && last) {
				for (let i = first.index + 1; i <= last.index; i++) {
					if (color.rgb !== opt.fragments[i].format.getColor().rgb) {
						color = null;
						break;
					}
				}
			} else {
				color = null;
			}
			return (color !== null ? new ApiColor(color) : null);
		}
	};

	/**
	 * Sets the font color property to the specified font.
	 * <note>This method will work only with the text format of the cell.</note>
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @param {ApiColor} Color - Font color.
	 * @since 7.4.0
	 * @see office-js-api/Examples/{Editor}/ApiFont/Methods/SetColor.js
	 */
	ApiFont.prototype.SetColor = function (Color) {
		if (!Color instanceof ApiColor) {
			logError(new Error('Invalid type of parametr "Color".'));
			return;
		}
		if (this._object instanceof ApiCharacters) {
			// todo ms 19 allows to set any string, but maybe we shoud check it before set
			let opt = this._object._options;
			if (opt.start <= opt.len) {
				let editor = this._object._parent.range.worksheet.workbook.oApi.wb.cellEditor;
				let begin = opt.start - 1;
				if (begin > opt.len) begin = opt.len - 1;
				let end = opt.start + opt.length - 1;
				if (end > opt.len) end = opt.len - 1;
				if (begin != end) {
					editor._extractFragments(begin, end - begin, opt.fragments);
					let first = editor._findFragment(begin, opt.fragments);
					let last = editor._findFragment(end - 1, opt.fragments);
					if (first && last) {
						for (let i = first.index; i <= last.index; ++i) {
							editor._setFormatProperty(opt.fragments[i].format, "c", Color.color);
						}
						editor._mergeFragments(opt.fragments);
						let bbox = this._object._parent.range.bbox;
						let range = this._object._parent.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c1);
						range.setValue2(opt.fragments);
						this._object._options = this._object._parent.GetCharacters(opt.uStart, opt.uLength)._options;
					}
				}
			}
		}
	};

	Object.defineProperty(ApiFont.prototype, "Color", {
		get: function () {
			return this.GetColor();
		},
		set: function (Color) {
			return this.SetColor(Color);
		}
	});


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFreezePanes
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Sets the frozen cells in the active worksheet view. The range provided corresponds to the cells that will be frozen in the top- and left-most pane.
	 * @memberof ApiFreezePanes
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange | String} frozenRange - A range that represents the cells to be frozen.
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiFreezePanes/Methods/FreezeAt.js
	 */
	ApiFreezePanes.prototype.FreezeAt = function (frozenRange) {
		let api = this.ws.workbook.oApi;
		let tempRange = (typeof frozenRange === 'string') ? api.GetRange(frozenRange) : frozenRange;

		if (tempRange.range) {
			let bbox = tempRange.range.bbox;
			let r = bbox.r2 < AscCommon.gc_nMaxRow0 ? bbox.r2 + 1 : bbox.r2;
			let c = bbox.c2 < AscCommon.gc_nMaxCol0 ? bbox.c2 + 1 : bbox.c2;
			api.asc_freezePane(null, c, r);
		} else {
			logError(new Error('Invalid parametr "frozenRange".'));
		}
	};

	/**
	 * Freezes the first column or columns of the current worksheet.
	 * @memberof ApiFreezePanes
	 * @typeofeditors ["CSE"]
	 * @param {Number?} [count=0] - Optional number of columns to freeze, or zero to unfreeze all columns.
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiFreezePanes/Methods/FreezeColumns.js
	 */
	ApiFreezePanes.prototype.FreezeColumns = function (count) {
		let api = this.ws.workbook.oApi;
		if (count == undefined) count = 0;
		if (typeof count === 'number' && count > 0 && count <= AscCommon.gc_nMaxCol0) {
			api.asc_freezePane(null, count, 0);
		} else if (!!api.wb.getWorksheet().topLeftFrozenCell && count === 0) {
			api.asc_freezePane(undefined);
		} else {
			logError(new Error('Invalid parametr "count".'));
		}
	};

	/**
	 * Freezes the top row or rows of the current worksheet.
	 * @memberof ApiFreezePanes
	 * @typeofeditors ["CSE"]
	 * @param {Number?} [count=0] - Optional number of rows to freeze, or zero to unfreeze all rows.
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiFreezePanes/Methods/FreezeRows.js
	 */
	ApiFreezePanes.prototype.FreezeRows = function (count) {
		let api = this.ws.workbook.oApi;
		if (count == undefined) count = 0;
		if (typeof count === 'number' && count > 0 && count <= AscCommon.gc_nMaxRow0) {
			api.asc_freezePane(null, 0, count);
		} else if (!!api.wb.getWorksheet().topLeftFrozenCell && count === 0) {
			api.asc_freezePane(undefined);
		} else {
			logError(new Error('Invalid parametr "count".'));
		}
	};

	/**
	 * Returns a range that describes the frozen cells in the active worksheet view.
	 * @memberof ApiFreezePanes
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null} - Returns null if there is no frozen pane.
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiFreezePanes/Methods/GetLocation.js
	 */
	ApiFreezePanes.prototype.GetLocation = function () {
		let result = null;
		let api = this.ws.workbook.oApi;
		let cell = api.wb.getWorksheet().topLeftFrozenCell;
		if (cell) {
			let c = cell.getCol0();
			let r = cell.getRow0();
			if (c == 0) {
				// hole row
				r--;
				c = AscCommon.gc_nMaxCol0;
			} else if (r == 0) {
				// whole column
				c--;
				r = AscCommon.gc_nMaxRow0;
			} else {
				// cell
				r--;
				c--;
			}
			result = new ApiRange(this.ws.getRange3(0, 0, r, c));
		}
		return result;
	};

	/**
	 * Removes all frozen panes in the current worksheet.
	 * @memberof ApiFreezePanes
	 * @typeofeditors ["CSE"]
	 * @since 8.0.0
	 * @see office-js-api/Examples/{Editor}/ApiFreezePanes/Methods/Unfreeze.js
	 */
	ApiFreezePanes.prototype.Unfreeze = function () {
		if (!!this.ws.workbook.oApi.wb.getWorksheet().topLeftFrozenCell)
			this.ws.workbook.oApi.asc_freezePane(undefined);
	};

	/**
	 * Class representing a user-protected range.
	 * @constructor
	 */
	function ApiProtectedRange(protectedRange) {
		this.protectedRange = protectedRange;
	}

	/**
	 * Sets a title to the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sTitle - The title which will be displayed for the current protected range.
	 * @returns {boolean} - Returns false if a user doesn't have permission to modify the protected range.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/SetTitle.js
	 */
	ApiProtectedRange.prototype.SetTitle = function (sTitle) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let result = false;
		if (isValidTitle && sTitle !== this.protectedRange.asc_getName()) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
				newProtectedRange.asc_setName(sTitle);
				let editRes = worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true);
				if (typeof editRes === "object") {
					this.protectedRange = editRes;
				}
				if (editRes) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Sets a range to the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - The cell range which will be set for the current protected range.
	 * @returns {boolean} - Returns false if a user doesn't have permission to modify the protected range.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/SetRange.js
	 */
	ApiProtectedRange.prototype.SetRange = function (sRange) {
		let isValidRange = typeof (sRange) === 'string' && sRange.trim() !== '';
		let result = false;
		if (isValidRange /*asc_getRef !==*/) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
				newProtectedRange.asc_setRef(sRange);
				let editRes = worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true);
				if (typeof editRes === "object") {
					this.protectedRange = editRes;
				}
				if (editRes) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Specifies the user type of the protected range.
	 * @typedef {("CanEdit" | "CanView" | "NotView")} ProtectedRangeUserType
	 * @see office-js-api/Examples/Enumerations/ProtectedRangeUserType.js
	 */


	/**
	 * Sets a user to the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sId - The user ID.
	 * @param {string} sName - The user name.
	 * @param {ProtectedRangeUserType} protectedRangeUserType - The user type of the protected range.
	 * @returns {ApiProtectedRangeUserInfo | null} - Returns null if a user doesn't have permission to modify the protected range.
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/AddUser.js
	 */
	ApiProtectedRange.prototype.AddUser = function (sId, sName, protectedRangeUserType) {
		let isValidIdTitle = typeof (sId) === 'string' && sId.trim() !== '';
		let isValidTitle = typeof (sName) === 'string' && sName.trim() !== '';
		let result = null;
		if (isValidTitle && isValidIdTitle) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);

				let newUser = new Asc.CUserProtectedRangeUserInfo();
				newUser.asc_setId(sId);
				newUser.asc_setName(sName);
				let nType = Asc.c_oSerUserProtectedRangeType.edit;
				if (protectedRangeUserType === "CanView") {
					nType = Asc.c_oSerUserProtectedRangeType.view;
				} else if (protectedRangeUserType === "NotView") {
					nType = Asc.c_oSerUserProtectedRangeType.notView;
				}
				newUser.asc_setType(nType);

				let users = this.protectedRange.asc_getUsers();
				if (!users) {
					users = [];
				}
				users.push(newUser);
				newProtectedRange.asc_setUsers(users);
				let editRes = worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true);
				if (typeof editRes === "object") {
					this.protectedRange = editRes;
				}
				result = new ApiProtectedRangeUserInfo(newUser, this.protectedRange);
			}
		}
		return result;
	};

	/**
	 * Removes a user from the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sId - The user ID.
	 * @returns {bool}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/DeleteUser.js
	 */
	ApiProtectedRange.prototype.DeleteUser = function (sId) {
		let isValidId = typeof (sId) === 'string' && sId.trim() !== '';
		let result = false;
		if (isValidId) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let userInfo = this.protectedRange.getUserById(sId);
				if (userInfo) {
					let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
					let users = this.protectedRange.asc_getUsers();
					if (users) {
						let newUsers = [];
						for (let i = 0; i < users.length; i++) {
							if (i !== userInfo.index) {
								newUsers.push(users[i].clone());
							}
						}
						newProtectedRange.asc_setUsers(newUsers);

						if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
							result = true;
						}
					}
				}
			}
		}
		return result;
	};

	/**
	 * Returns all users from the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiProtectedRangeUserInfo[] | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/GetAllUsers.js
	 */
	ApiProtectedRange.prototype.GetAllUsers = function () {
		let worksheet = this.protectedRange._ws;
		let result = null;
		if (worksheet) {
			let users = this.protectedRange.asc_getUsers();
			if (users) {
				let newUsers = [];
				for (let i = 0; i < users.length; i++) {
					newUsers.push(new ApiProtectedRangeUserInfo(users[i], this.protectedRange));
				}
				result = newUsers;
			}
		}

		return result;
	};

	/**
	 * Sets the type of the "Anyone" user to the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {ProtectedRangeUserType} protectedRangeUserType - The user type of the protected range.
	 * @returns {bool}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/SetAnyoneType.js
	 */
	ApiProtectedRange.prototype.SetAnyoneType = function (protectedRangeUserType) {
		let nType = Asc.c_oSerUserProtectedRangeType.edit;
		if (protectedRangeUserType === "CanView") {
			nType = Asc.c_oSerUserProtectedRangeType.view;
		} else if (protectedRangeUserType === "NotView") {
			nType = Asc.c_oSerUserProtectedRangeType.notView;
		}
		let result = false;
		if (this.protectedRange.asc_getType() !== nType) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
				newProtectedRange.asc_setType(nType);
				let editRes = worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true);
				if (typeof editRes === "object") {
					this.protectedRange = editRes;
				}
				if (editRes) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Returns an object that represents a user from the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sId - The user ID.
	 * @returns {ApiProtectedRangeUserInfo | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRange/Methods/GetUser.js
	 */
	ApiProtectedRange.prototype.GetUser = function (sId) {
		let isValidRange = typeof (sId) === 'string' && sId.trim() !== '';
		let result = null;
		if (isValidRange) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let userInfo = this.protectedRange.getUserById(sId);
				if (userInfo) {
					result = new ApiProtectedRangeUserInfo(userInfo.obj, this.protectedRange)
				}
			}
		}
		return result;
	};

	/**
	 * Class representing a user from the current protected range.
	 * @constructor
	 */
	function ApiProtectedRangeUserInfo(userInfo, protectedRange) {
		this.userInfo = userInfo;
		this.protectedRange = protectedRange;
	}

	/**
	 * Returns the name property of the current user's information.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRangeUserInfo/Methods/GetName.js
	 */
	ApiProtectedRangeUserInfo.prototype.GetName = function () {
		//the sets methods are available from the parent
		// not adding by ApiProtectedRangeUserInfo because need change id/name together
		return this.userInfo.asc_getName();
	};

	/**
	 * Returns the type property of the current user's information.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {ProtectedRangeUserType}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRangeUserInfo/Methods/GetType.js
	 */
	ApiProtectedRangeUserInfo.prototype.GetType = function () {
		let nType = this.userInfo.asc_getType();
		let protectedRangeUserType = "CanEdit";//default
		if (nType === Asc.c_oSerUserProtectedRangeType.view) {
			protectedRangeUserType = "CanView"
		} else if (nType === Asc.c_oSerUserProtectedRangeType.notView) {
			protectedRangeUserType = "NotView";
		}
		return protectedRangeUserType;
	};


	/**
	 * Returns the ID property of the current user's information.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
	 * @since 8.1.0
	 * @see office-js-api/Examples/{Editor}/ApiProtectedRangeUserInfo/Methods/GetId.js
	 */
	ApiProtectedRangeUserInfo.prototype.GetId = function () {
		return this.userInfo.asc_getId();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPivotTable
	//
	//------------------------------------------------------------------------------------------------------------------

	/* Methods */

	/**
	 * Adds a data field to a PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number | string} field - The index number or name of the field.
	 * @returns {ApiPivotDataField}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/AddDataField.js
	 */
	ApiPivotTable.prototype.AddDataField = function (field) {
		let pivotIndex = -1;
		const pivotFields = this.pivot.asc_getPivotFields();
		if (typeof field === 'string') {
			pivotIndex = this.pivot.getFieldIndexByValue(field);
		} else if (typeof field === 'number') {
			pivotIndex = field - 1;
		} else {
			private_MakeError('Bad field indentifier type.')
			return null;
		}
		if (pivotFields[pivotIndex]) {
			this.pivot.asc_addDataField(this.api, pivotIndex);
			const dataFields = this.pivot.asc_getDataFields();
			return new ApiPivotDataField(this, dataFields.length - 1, dataFields[dataFields.length - 1]);
		}
		private_MakeError("Field with such an identifier does not exist.");
		return null;
	};
	/**
	 * Adds row, column, and page fields to a PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {Object} options
	 * @param {number | string | number[] | string[]} [options.rows] - Specifies an array of field names or ids to be added as rows or added to the category axis.
	 * @param {number | string | number[] | string[]} [options.columns] - Specifies an array of field names or ids to be added as columns or added to the series axis.
	 * @param {number | string | number[] | string[]} [options.pages] - Specifies an array of field names or ids to be added as pages or added to the page area.
	 * @param {boolean} [options.addToTable] - Applies only to PivotTable reports. True to add the specified fields to the report (none of the existing fields are replaced).
	 * False to replace existing fields with the new fields. The default value is False.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/AddFields.js
	 */
	ApiPivotTable.prototype.AddFields = function (options) {
		options['rows'] = options['rows'] != null ? options['rows'] : [];
		options['columns'] = options['columns'] != null ? options['columns'] : [];
		options['pages'] = options['pages'] != null ? options['pages'] : [];

		const rows = Array.isArray(options['rows']) ? options['rows'] : [options['rows']];
		const cols = Array.isArray(options['columns']) ? options['columns'] : [options['columns']];
		const pages = Array.isArray(options['pages']) ? options['pages'] : [options['pages']];
		const cacheFields = this.pivot.asc_getCacheFields();
		const t = this;

		function processField(field, callback) {
			let index = null;
			if (typeof field == "number" && field > 0 && field - 1 < cacheFields.length) {
				index = field - 1;
			} else if (typeof field == "string") {
				index = t.pivot.getFieldIndexByValue(field.trim());
				if (index < 0) {
					index = null;
				}
			}
			if (index !== null) {
				callback(index);
			} else {
				private_MakeError("There is no field with such an identifier.");
			}
		}

		if (!options.addToTable) {
			const pivotFields = this.GetPivotFields()
			pivotFields.forEach(function (pivotField) {
				pivotField.Remove()
			})

		}
		rows.forEach(function(row) {
			processField(row, function(index) {
				t.pivot.asc_addRowField(t.api, index);
			});
		});
		cols.forEach(function(col) {
			processField(col, function(index) {
				t.pivot.asc_addColField(t.api, index);
			});
		});
		pages.forEach(function(page) {
			processField(page, function(index) {
				t.pivot.asc_addPageField(t.api, index);
			});
		});
	};
	/**
	 * Deletes all filters currently applied to the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/ClearAllFilters.js
	 */
	ApiPivotTable.prototype.ClearAllFilters = function () {
		this.pivot.asc_removeFilters(this.api);
	};
	/**
	 * Clears the Pivot Table.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/ClearTable.js
	 */
	ApiPivotTable.prototype.ClearTable = function () {
		const ws = this.pivot.worksheet;
		const name = this.pivot.asc_getName();
		const range = this.pivot.getRange();
		const bbox = new Asc.Range(range.c1, range.r1, range.c1, range.r1);
		const dataRef = this.pivot.cacheDefinition.cacheSource.worksheetSource.getDataRef();
		let index = -1;
		for (let i = 0; i < ws.pivotTables.length; i += 1) {
			if (ws.pivotTables[i].Get_Id() === this.pivot.Get_Id()) {
				index = i;
				break;
			}
		}
		if (index !== -1) {
			ws._deletePivotTable(ws.pivotTables, this.pivot, index);
			this.pivot = this.api._asc_insertPivot(ws.workbook, dataRef, ws, bbox, false);
			this.pivot.asc_setName(name);
			return;
		}
		private_MakeError("Unknown error!");
	};
	/**
	 * Returns the value for the data filed in a PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string[]} items - Describes a single cell in the PivotTable report.
	 * @returns {number | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetData.js
	 */
	ApiPivotTable.prototype.GetData = function (items) {
		const params = this.pivot.asc_getDataToGetPivotData(items);
		const cell = this.pivot.getCellByGetPivotDataParams(params);
		if (cell) {
			return this.pivot.worksheet.getCell3(cell.row, cell.col).getValue();
		}
		private_MakeError('There is no data with that params.');
		return null;
	};
	/**
	 * Returns a Range object with information about a data item in a PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string} [dataField]
	 * @param {string[]} [fieldItemsArray]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetPivotData.js
	 */
	ApiPivotTable.prototype.GetPivotData = function (dataField, fieldItemsArray) {
		const cell = this.pivot.getCellByGetPivotDataParams({
			dataFieldName: dataField,
			optParams: fieldItemsArray
		});
		if (cell) {
			return new ApiRange(this.pivot.worksheet.getCell3(cell.row, cell.col));
		}
		return null;
	};
	/**
	 * Returns an collection that represents either a single PivotTable field
	 * or a collection of both the visible and hidden fields in the PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string | number} [field] - The name or index of the field to be returned.
	 * @returns {ApiPivotField[] | ApiPivotField | ApiPivotDataField | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetPivotFields.js
	 */
	ApiPivotTable.prototype.GetPivotFields = function (field) {
		const pivotFields = this.pivot.asc_getPivotFields();
		if (field != null) {
			let pivotIndex = -1;
			if (typeof field === 'number') {
				pivotIndex = field - 1;
				if (pivotFields[pivotIndex]) {
					return new ApiPivotField(this, pivotIndex, pivotFields[pivotIndex]);
				}
			} else if (typeof field === 'string') {
				pivotIndex = this.pivot.getFieldIndexByValue(field.trim());
				if (pivotIndex !== -1) {
					return new ApiPivotField(this, pivotIndex, pivotFields[pivotIndex]);
				}
				return this.GetDataFields(field);
			}
		}
		const t = this;
		return pivotFields.map(function(pivotField, i) {
			return new ApiPivotField(t, i, pivotField);
		});
	};
	Object.defineProperty(ApiPivotTable.prototype, "PivotFields", {
		get: function (field) {
			this.GetPivotFields(field);
		}
	});
	/**
	 * Returns the value of a pivot table cell.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number} rowLine
	 * @param {number} colLine
	 * @returns {number | string | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/PivotValueCell.js
	 */
	ApiPivotTable.prototype.PivotValueCell = function (rowLine, colLine) {
		if (rowLine > 0 && colLine > 0) {
			const pivotRange = this.pivot.getRange();
			const location = this.pivot.location;
			const baseCol = pivotRange.c1 + location.firstDataCol;
			const baseRow = pivotRange.r1 + location.firstDataRow;
			const curRow = rowLine + baseRow - 1;
			const curCol = colLine + baseCol - 1;
			if (curRow <= pivotRange.r2 && curCol <= pivotRange.c2) {
				return this.pivot.worksheet.getCell3(curRow, curCol).getValue();
			}
		}
		private_MakeError('Cell is out of range');
		return null;
	};
	/**
	 * Show details of pivot cell.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number} rowLine - report area row
	 * @param {number} colLine - report area column
	 * @returns {boolean} success
	 * @see office-js-api/Examples/ApiPivotTable/Methods/ShowDetails.js
	 */
	ApiPivotTable.prototype.ShowDetails = function (rowLine, colLine) {
		if (rowLine > 0 && colLine > 0) {
			const pivotRange = this.pivot.getRange();
			const location = this.pivot.location;
			const baseCol = pivotRange.c1 + location.firstDataCol;
			const baseRow = pivotRange.r1 + location.firstDataRow;
			const curRow = rowLine + baseRow - 1;
			const curCol = colLine + baseCol - 1;
			if (curRow <= pivotRange.r2 && curCol <= pivotRange.c2) {
				return this.api.asc_pivotShowDetails(this.pivot, {row: curRow, col: curCol});
			}
		}
		private_MakeError('Cell is out of range');
		return false;
	};
	/**
	 * Refreshes the PivotTable report from the source data.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/RefreshTable.js
	 */
	ApiPivotTable.prototype.RefreshTable = function () {
		this.pivot.asc_refresh(this.api);
	};
	/**
	 * Updates the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/Update.js
	 */
	ApiPivotTable.prototype.Update = function () {
		this.pivot.asc_refresh(this.api);
	};
	/**
	 * Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.
	 * @memberof ApiPivotTable
	 * @param {boolean} repeat - Specifies whether to repeat all field item labels in a PivotTable report.
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetRepeatAllLabels.js
	 */
	ApiPivotTable.prototype.SetRepeatAllLabels = function (repeat) {
		if (typeof repeat == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setFillDownLabelsDefault(repeat);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "repeat".');
		}
	};
	Object.defineProperty(ApiPivotTable.prototype, "RepeatAllLabels", {
		set: function (repeat) {
			this.SetRepeatAllLabels(repeat);
		}
	});
	/**
	 * Sets the way the specified PivotTable items appear—in table format or in outline format.
	 * @memberof ApiPivotTable
	 * @param {PivotLayoutType} type - Type of layot report form.
	 * @param {boolean} compact - Use compact form or not.
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetRowAxisLayout.js
	 */
	ApiPivotTable.prototype.SetRowAxisLayout = function (type, compact) {
		let props = null;
		if (typeof type === "string" && (type === "Tabular" || type === "Outline")) {
			props = new Asc.CT_pivotTableDefinition();
			props.asc_setOutline((type == "Outline"));
		} else {
			private_MakeError('Invalid type of "type" or invalid value.');
		}
		if (compact != null) {
			if (typeof compact === "boolean") {
				if (!props) {
					props = new Asc.CT_pivotTableDefinition();
				}
				props.asc_setCompact(compact);
			} else {
				private_MakeError('Invalid type of "compact".');
			}
		}
		if (props) {
			this.pivot.asc_set(this.api, props);
		}
	};
	Object.defineProperty(ApiPivotTable.prototype, "RowAxisLayout", {
		set: function (type, compact) {
			this.SetRowAxisLayout(type, compact);
		}
	});
	/**
	 * @typedef { "Hidden" | "Top" | "Bottom" } PivotSubtotalLayoutType
	 */
	/**
	 * Sets PivotTable setting layout subtotal location.
	 * @memberof ApiPivotTable
	 * @param {PivotSubtotalLayoutType} type - Type of subtotal layout
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetSubtotalLocation.js
	 */
	ApiPivotTable.prototype.SetSubtotalLocation = function (type) {
		if ( typeof type == "string" && (type === "Hidden" || type === "Bottom" || type === "Top") ) {
			const props = new Asc.CT_pivotTableDefinition();
			if (type == "Hidden") {
				props.asc_setDefaultSubtotal(false);
			} else {
				props.asc_setDefaultSubtotal(true);
				props.asc_setSubtotalTop( (type == "Top") );
			}
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "type" or invalid value.');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "SubtotalLocation", {
		set: function (type) {
			this.SetSubtotalLocation(type);
		}
	});
	/**
	 * Removes field from all PivotTable categories.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number | string} identifier - The index number or name of the field.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/RemoveField.js
	 */
	ApiPivotTable.prototype.RemoveField = function (identifier) {
		const pivotField = this.GetPivotFields(identifier);
		if (pivotField) {
			this.pivot.asc_removeField(this.api, pivotField.index);
		}
	};

	/**
	 * @typedef { "Up" | "Down" | "Begin" | "End" } PivotMoveFieldType
	 */
	/**
	 * @typedef {"Rows" | "Columns" | "Filters" | "Values" | "Hidden" } PivotFieldOrientationType
	 */

	/**
	 * Moves field from one category to another.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number | string} identifier - The index number or name of the field.
	 * @param {PivotMoveFieldType | PivotFieldOrientationType} type - The type of the field to move.
	 * @param {number} [index] - The index of the field in new category.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/MoveField.js
	 */
	ApiPivotTable.prototype.MoveField = function (identifier, type, index) {
		const pivotField = this.GetPivotFields(identifier);
		if (pivotField) {
			pivotField.Move(type, index)
		}
	};
	/**
	 * Selects the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/Select.js
	 */
	ApiPivotTable.prototype.Select = function () {
		this.pivot.asc_select(this.api);
	};

	/* Attributes */

	/**
	 * Returns an collection that represents a collection of all column fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField[]}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetColumnFields.js
	 */
	ApiPivotTable.prototype.GetColumnFields = function (field) {
		const pivotFields = this.pivot.asc_getPivotFields();
		const colFields = this.pivot.asc_getColumnFields();
		const t = this;
		return colFields.map(function(colField, i) {
			const index = colField.asc_getIndex();
			return new ApiPivotField(t, index, pivotFields[index]);
		});
	};

	Object.defineProperty(ApiPivotTable.prototype, "ColumnFields", {
		get: function (field) {
			return this.GetColumnFields(field);
		}
	});
	/**
	 * Returns an collection that represents either a single PivotTable data field
	 * or a collection of all visible data fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {number | string |undefined} field - The name or index of the field to be returned.
	 * @returns {ApiPivotDataField[] | ApiPivotDataField | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetDataFields.js
	 */
	ApiPivotTable.prototype.GetDataFields = function (field) {
		const dataFields = this.pivot.asc_getDataFields();
		if (field != null) {
			let dataIndex = -1;
			if (typeof field === 'number') {
				dataIndex = field - 1;
				const dataField = dataFields[dataIndex];
				if (dataField) {
					return new ApiPivotDataField(this, dataIndex, dataField)
				}
			} else if (typeof field === 'string') {
				const dataIndex = this.pivot.dataFields.getIndexByName(field.trim())
				if (dataIndex !== -1) {
					return  new ApiPivotDataField(this, dataIndex, dataFields[dataIndex]);
				}
			}
			private_MakeError("A field with such an identifier does not exist.");
			return null;
		}
		const t = this;
		return dataFields.map(function(dataField, i) {
			return new ApiPivotDataField(t, i, dataField);
		});
	};

	Object.defineProperty(ApiPivotTable.prototype, "DataFields", {
		get: function (field) {
			return this.GetDataFields(field);
		}
	});

	/**
	 * Returns array that represents all not added fields in PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField[]}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetHiddenFields.js
	 */
	ApiPivotTable.prototype.GetHiddenFields = function () {
		var fields = this.pivot.asc_getPivotFields();
		var hidden = [];
		for (var i = 0; i < fields.length; i++)
			if (fields[i].axis === null && !fields[i].dataField)
				hidden.push( new ApiPivotField(this, i, fields[i]));

		return hidden;
	};

	Object.defineProperty(ApiPivotTable.prototype, "HiddenFields", {
		get: function () {
			return this.GetHiddenFields();
		}
	});

	/**
	 * Returns array that represents all added fields in PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField[]}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetVisibleFields.js
	 */
	ApiPivotTable.prototype.GetVisibleFields = function () {
		const pivotFields = this.pivot.asc_getPivotFields();
		const visible = [];
		for (var i = 0; i < pivotFields.length; i++) {
			if (pivotFields[i].axis !== null && !pivotFields[i].dataField) {
				visible.push( new ApiPivotField(this, i, pivotFields[i]) );
			}
		}
		return visible.concat(this.GetDataFields());
	};

	Object.defineProperty(ApiPivotTable.prototype, "VisibleFields", {
		get: function () {
			return this.GetVisibleFields();
		}
	});

	/**
	 * Returns an collection that represents either a single PivotTable page field
	 * or a collection of all visible page fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField[]}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetPageFields.js
	 */
	ApiPivotTable.prototype.GetPageFields = function (field) {
		const pivotFields = this.pivot.asc_getPivotFields();
		const pageFields = this.pivot.asc_getPageFields();
		const t = this;
		return pageFields.map(function(pageField, i) {
			const index = pageField.asc_getIndex();
			return new ApiPivotField(t, index, pivotFields[index]);
		});
	};

	Object.defineProperty(ApiPivotTable.prototype, "PageFields", {
		get: function (field) {
			return this.GetPageFields(field);
		}
	});
	/**
	 * Returns an collection that represents a collection of all row fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField[]}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetRowFields.js
	 */
	ApiPivotTable.prototype.GetRowFields = function (field) {
		const pivotFields = this.pivot.asc_getPivotFields();
		const rowFields = this.pivot.asc_getRowFields();
		const t = this;
		return rowFields.map(function(rowField, i) {
			const index = rowField.asc_getIndex();
			return new ApiPivotField(t, index, pivotFields[index]);
		});
	};

	Object.defineProperty(ApiPivotTable.prototype, "RowFields", {
		get: function () {
			return this.GetRowFields();
		}
	});
	/**
	 * Returns name of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetName.js
	 */
	ApiPivotTable.prototype.GetName = function () {
		return this.pivot.asc_getName();
	};

	/**
	 * Sets name of the PivotTable.
	 * @memberof ApiPivotTable
	 * @param {string} name - Name for PivotTable.
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetName.js
	 */
	ApiPivotTable.prototype.SetName = function (name) {
		if (typeof name == "string" && name.trim().length) {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setName(name.trim());
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "name" or "name" is empty.');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (name) {
			this.SetName(name);
		}
	});

	/**
	 * Returns PivotTable Grand Totals settings for columns.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetColumnGrand.js
	 */
	ApiPivotTable.prototype.GetColumnGrand = function () {
		return this.pivot.asc_getColGrandTotals();
	};

	/**
	 * Sets PivotTable Grand Totals settings for columns.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Shows grand totals for columns or not.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetColumnGrand.js
	 */
	ApiPivotTable.prototype.SetColumnGrand = function (show) {
		if (typeof show == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setColGrandTotals(show);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "ColumnGrand", {
		get: function () {
			return this.GetColumnGrand();
		},
		set: function (show) {
			this.SetColumnGrand(show);
		}
	});

	/**
	 * Returns PivotTable Grand Totals settings for rows.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetRowGrand.js
	 */
	ApiPivotTable.prototype.GetRowGrand = function () {
		return this.pivot.asc_getRowGrandTotals();
	};

	/**
	 * Sets PivotTable Grand Totals settings for rows.
	 * @memberof ApiPivotTable
	 * @param {boolean} show - Shows grand totals for rows or not.
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetRowGrand.js
	 */
	ApiPivotTable.prototype.SetRowGrand = function (show) {
		if (typeof show == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setRowGrandTotals(show);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "RowGrand", {
		get: function () {
			return this.GetRowGrand();
		},
		set: function (show) {
			this.SetRowGrand(show);
		}
	});

	/** @typedef {"OverThenDown" | "DownThenOver"} FieldsInReportFilterType */

	/**
	 * Returns PivotTable display fields in report filter area settings.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {{Type: FieldsInReportFilterType, ReportFilterFields: number}}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetDisplayFieldsInReportFilterArea.js
	 */
	ApiPivotTable.prototype.GetDisplayFieldsInReportFilterArea = function () {
		return {
			"Type": (this.pivot.asc_getPageOverThenDown() ? "OverThenDown" : "DownThenOver"),
			"ReportFilterFields": this.pivot.asc_getPageWrap()
		};
	};

	/**
	 * Returns PivotTable display fields in report filter area settings.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {FieldsInReportFilterType} type
	 * @param {number} fields - Count of report filter fields.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetDisplayFieldsInReportFilterArea.js
	 */
	ApiPivotTable.prototype.SetDisplayFieldsInReportFilterArea = function (type, fields) {
		var props = null;
		if (type != undefined) {
			if ( typeof type == "string" && (type == 'OverThenDown' || type == 'DownThenOver') ) {
				props = new Asc.CT_pivotTableDefinition();
				props.asc_setPageOverThenDown( (type == 'OverThenDown') );
			} else {
				private_MakeError('Invalid type of "type".');
			}
		}

		if (fields != undefined) {
			if (typeof fields == "number" && fields >= 0 && fields <= 255) {
				if (!props)
					props = new Asc.CT_pivotTableDefinition();

				props.asc_setPageWrap(fields);
			} else {
				private_MakeError('Invalid type of "fields" or invalid value.');
			}
		}

		if (props)
			this.pivot.asc_set(this.api, props);
	};

	/**
	 * Returns PivotTable setting show field headers for rows and columns.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetDisplayFieldCaptions.js
	 */
	ApiPivotTable.prototype.GetDisplayFieldCaptions = function () {
		return this.pivot.asc_getShowHeaders();
	};

	/**
	 * Sets show field headers for rows and columns.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show or not field headers for rows and columns.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetDisplayFieldCaptions.js
	 */
	ApiPivotTable.prototype.SetDisplayFieldCaptions = function (show) {
		if (typeof show == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setShowHeaders(show);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "DisplayFieldCaptions", {
		get: function () {
			return this.GetDisplayFieldCaptions ();
		},
		set: function (show) {
			this.SetDisplayFieldCaptions (show);
		}
	});

	/**
	 * Returns the title of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTitle.js
	 */
	ApiPivotTable.prototype.GetTitle = function () {
		return this.pivot.asc_getTitle() || "";
	};

	/**
	 * Sets the title of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string} title - Title for PivotTable.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetTitle.js
	 */
	ApiPivotTable.prototype.SetTitle = function (title) {
		if (typeof title == "string") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setTitle(title.trim());
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "title".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "Title", {
		get: function () {
			return this.GetTitle();
		},
		set: function (title) {
			this.SetTitle(title);
		}
	});

	/**
	 * Returns the description of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetDescription.js
	 */
	ApiPivotTable.prototype.GetDescription = function () {
		return this.pivot.asc_getDescription() || "";
	};

	/**
	 * Sets the description of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string} description - Description for PivotTable.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetDescription.js
	 */
	ApiPivotTable.prototype.SetDescription = function (description) {
		if (typeof description == "string") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setDescription(description.trim());
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "description".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "Description", {
		get: function () {
			return this.GetDescription();
		},
		set: function (description) {
			this.SetDescription(description);
		}
	});

	/**
	 * Returns the style name of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetStyleName.js
	 */
	ApiPivotTable.prototype.GetStyleName = function () {
		return this.pivot.asc_getStyleInfo().asc_getName();
	};

	/**
	 * Sets the style of the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Style name.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetStyleName.js
	 */
	ApiPivotTable.prototype.SetStyleName = function (name) {
		if (typeof name == "string" && name.trim().length) {
			this.pivot.asc_getStyleInfo().asc_setName(this.api, this.pivot, name.trim());
		} else {
			private_MakeError('Invalid type of "name" or "name" is empty.');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "StyleName", {
		get: function () {
			return this.GetStyleName();
		},
		set: function (name) {
			this.SetStyleName(name);
		}
	});

	/**
	 * Returns the row headers style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableStyleRowHeaders.js
	 */
	ApiPivotTable.prototype.GetTableStyleRowHeaders = function () {
		return this.pivot.asc_getStyleInfo().asc_getShowRowHeaders();
	};

	/**
	 * Sets the row headers style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show or not row headers.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetTableStyleRowHeaders.js
	 */
	ApiPivotTable.prototype.SetTableStyleRowHeaders = function (show) {
		if (typeof show == "boolean") {
			this.pivot.asc_getStyleInfo().asc_setShowRowHeaders(this.api, this.pivot, show);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "ShowTableStyleRowHeaders", {
		get: function () {
			return this.GetTableStyleRowHeaders();
		},
		set: function (show) {
			this.SetTableStyleRowHeaders(show);
		}
	});

	/**
	 * Returns the column headers style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableStyleColumnHeaders.js
	 */
	ApiPivotTable.prototype.GetTableStyleColumnHeaders = function () {
		return this.pivot.asc_getStyleInfo().asc_getShowColHeaders();
	};

	/**
	 * Sets the column headers style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show or not column headers.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetTableStyleColumnHeaders.js
	 */
	ApiPivotTable.prototype.SetTableStyleColumnHeaders = function (show) {
		if (typeof show == "boolean") {
			this.pivot.asc_getStyleInfo().asc_setShowColHeaders(this.api, this.pivot, show);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "ShowTableStyleColumnHeaders", {
		get: function () {
			return this.GetTableStyleColumnHeaders();
		},
		set: function (show) {
			this.SetTableStyleColumnHeaders(show);
		}
	});

	/**
	 * Returns the banded rows style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableStyleRowStripes.js
	 */
	ApiPivotTable.prototype.GetTableStyleRowStripes = function () {
		return this.pivot.asc_getStyleInfo().asc_getShowRowStripes();
	};

	/**
	 * Sets the banded rows style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show or not banded rows.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetTableStyleRowStripes.js
	 */
	ApiPivotTable.prototype.SetTableStyleRowStripes = function (show) {
		if (typeof show == "boolean") {
			this.pivot.asc_getStyleInfo().asc_setShowRowStripes(this.api, this.pivot, show);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "ShowTableStyleRowStripes", {
		get: function () {
			return this.GetTableStyleRowStripes();
		},
		set: function (show) {
			this.SetTableStyleRowStripes(show);
		}
	});

	/**
	 * Returns the banded columns style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableStyleColumnStripes.js
	 */
	ApiPivotTable.prototype.GetTableStyleColumnStripes = function () {
		return this.pivot.asc_getStyleInfo().asc_getShowColStripes();
	};

	/**
	 * Sets the banded columns style for PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show or not banded columns.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetTableStyleColumnStripes.js
	 */
	ApiPivotTable.prototype.SetTableStyleColumnStripes = function (show) {
		if (typeof show == "boolean") {
			this.pivot.asc_getStyleInfo().asc_setShowColStripes(this.api, this.pivot, show);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "ShowTableStyleColumnStripes", {
		get: function () {
			return this.GetTableStyleColumnStripes();
		},
		set: function (show) {
			this.SetTableStyleColumnStripes(show);
		}
	});

	/**
	 * Returns source for the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetSource.js
	 */
	ApiPivotTable.prototype.GetSource = function () {
		var location = this.pivot.getDataLocation();
		return new ApiRange( location.ws.getRange3(location.bbox.r1, location.bbox.c1, location.bbox.r2, location.bbox.c2) );
	};

	/**
	 * Sets source for the PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} source - Range where pivot table will be located.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetSource.js
	 */
	ApiPivotTable.prototype.SetSource = function (source) {
		if (source instanceof ApiRange) {
			var ref = source.GetWorksheet().GetName() + "!" + source.GetAddress(true, true);
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setDataRef(ref);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Source must be instance of ApiRange.');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "Source", {
		get: function () {
			return this.GetSource();
		},
		set: function (source) {
			this.SetSource(source);
		}
	});

	/**
	 * Returns a Range object that represents the range that contains the column area in the PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetColumnRange.js
	 */
	ApiPivotTable.prototype.GetColumnRange = function () {
		const res = this.pivot.asc_getColumnRange();
		if (res) {
			return new ApiRange(res);
		}
		return null;
	};

	Object.defineProperty(ApiPivotTable.prototype, "ColumnRange", {
		get: function () {
			return this.GetColumnRange();
		}
	});

	/**
	 * Returns a Range object that represents the range that contains the row area in the PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetRowRange.js
	 */
	ApiPivotTable.prototype.GetRowRange = function () {
		const res = this.pivot.asc_getRowRange();
		if (res) {
			return new ApiRange(res);
		}
		return null;
	};

	Object.defineProperty(ApiPivotTable.prototype, "RowRange", {
		get: function () {
			return this.GetRowRange();
		}
	});

	/**
	 * Returns a Range object that represents the range of values in a PivotTable.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetDataBodyRange.js
	 */
	ApiPivotTable.prototype.GetDataBodyRange = function () {
		const res = this.pivot.asc_getDataBodyRange();
		if (res) {
			return new ApiRange(res);
		}
		return null;
	};

	Object.defineProperty(ApiPivotTable.prototype, "DataBodyRange", {
		get: function () {
			return this.GetDataBodyRange();
		}
	});

	/**
	 * Returns a Range object that represents the range containing the entire PivotTable report,
	 * but doesn't include page fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableRange1.js
	 */
	ApiPivotTable.prototype.GetTableRange1 = function () {
		var ref =  (this.pivot.location ? this.pivot.location.ref : null);
		return (ref ? new ApiRange( this.pivot.worksheet.getRange3(ref.r1, ref.c1, ref.r2, ref.c2) ) : null);
	};

	Object.defineProperty(ApiPivotTable.prototype, "TableRange1", {
		get: function () {
			return this.GetTableRange1();
		}
	});

	/**
	 * Returns a Range object that represents the range containing the entire PivotTable report,
	 * including page fields.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiRange | null}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetTableRange2.js
	 */
	ApiPivotTable.prototype.GetTableRange2 = function () {
		var ref =  (this.pivot.location ? this.pivot.location.ref : null);
		var firsPFP, lastPFP;
		if (this.pivot.pageFieldsPositions.length) {
			firsPFP = this.pivot.pageFieldsPositions[0];
			lastPFP = this.pivot.pageFieldsPositions[ (this.pivot.pageFieldsPositions.length - 1) ];
		}
		var r1, c1, r2, c2;
		if (ref) {
			r1 = (firsPFP ? Math.min(firsPFP.row, ref.r1) : ref.r1);
			c1 = (firsPFP ? Math.min(firsPFP.col, ref.c1) : ref.c1);
			r2 = (lastPFP ? Math.max(lastPFP.row, ref.r2) : ref.r2);
			c2 = (lastPFP ? Math.max( (lastPFP.col + 1), ref.c2 ) : ref.c2);
		}
		return (ref ? new ApiRange( this.pivot.worksheet.getRange3(r1, c1, r2, c2) ) : null);
	};

	Object.defineProperty(ApiPivotTable.prototype, "TableRange2", {
		get: function () {
			return this.GetTableRange2();
		}
	});

	/**
	 * Returns the text string label that is displayed in the grand total column
	 * or row heading in the specified PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetGrandTotalName.js
	 */
	ApiPivotTable.prototype.GetGrandTotalName = function () {
		return ( this.pivot.asc_getGrandTotalCaption() || AscCommon.translateManager.getValue(AscCommonExcel.GRAND_TOTAL_CAPTION) );
	};

	/**
	 * Sets the text string label that is displayed in the grand total column
	 * or row heading in the specified PivotTable report.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Grand total name.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetGrandTotalName.js
	 */
	ApiPivotTable.prototype.SetGrandTotalName = function (name) {
		if (typeof name == "string") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setGrandTotalCaption( name.trim() );
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "name".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "GrandTotalName", {
		get: function () {
			return this.GetGrandTotalName();
		},
		set: function (name) {
			this.SetGrandTotalName(name);
		}
	});
	/**
	 * Sets PivotTable setting insert blank rows after each item.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} insert - Flag insert blank row after each item or not.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetLayoutBlankLine.js
	 */
	ApiPivotTable.prototype.SetLayoutBlankLine = function (insert) {
		if (typeof insert == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setInsertBlankRow(insert);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "insert".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "LayoutBlankLine", {
		set: function (insert) {
			this.SetLayoutBlankLine(insert);
		}
	});

	/**
	 * Sets PivotTable setting show subtotals.
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show subtotals or not.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/SetLayoutSubtotals.js
	 */
	ApiPivotTable.prototype.SetLayoutSubtotals = function (show) {
		if (typeof show == "boolean") {
			var props = new Asc.CT_pivotTableDefinition();
			props.asc_setDefaultSubtotal(show);
			this.pivot.asc_set(this.api, props);
		} else {
			private_MakeError('Invalid type of "show".');
		}
	};

	Object.defineProperty(ApiPivotTable.prototype, "LayoutSubtotals", {
		set: function (show) {
			this.SetLayoutSubtotals(show);
		}
	});

	/**
	 * Returns Pivot Table parent
	 * @memberof ApiPivotTable
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheet} - Pivot Table parent.
	 * @see office-js-api/Examples/ApiPivotTable/Methods/GetParent.js
	 */
	ApiPivotTable.prototype.GetParent = function () {
		return new ApiWorksheet(this.pivot.worksheet);
	};

	Object.defineProperty(ApiPivotTable.prototype, "Parent", {
		set: function () {
			this.GetParent();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPivotDataField
	//
	//------------------------------------------------------------------------------------------------------------------

	/** Methods */

	/**
	 * Removes data field from category.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/Remove.js
	 */
	ApiPivotDataField.prototype.Remove = function () {
		this.table.pivot.asc_removeDataField(this.table.api, this.index, this.dataIndex);
	};
	/**
	 * Moves field inside category.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {PivotMoveFieldType | PivotFieldOrientationType} type - The type of the field to move.
	 * @param {number} [index] - The index of the field in new category.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/Move.js
	 */
	ApiPivotDataField.prototype.Move = function (type, index) {
		function getIndexTo(type, indexFrom, fields) {
			switch (type) {
				case "Up":
					return (indexFrom > 0) ? indexFrom - 1 : indexFrom;
				case "Down":
					return (indexFrom < fields.length - 1) ? indexFrom + 1 : fields.length - 1;
				case "Begin":
					return 0;
				case "End":
					return fields.length - 1;
				default:
					return null;
			}
		}
		switch (type) {
			case "Rows":
				this.table.pivot.asc_moveToRowField(this.table.api, this.index, this.dataIndex, index);
				break;
			case "Columns":
				this.table.pivot.asc_moveToColField(this.table.api, this.index, this.dataIndex, index);
				break;
			case "Filters":
				this.table.pivot.asc_moveToPageField(this.table.api, this.index, this.dataIndex, index);
				break;
			case "Values":
				this.SetPosition(index);
				break;
			default:
				const fields = this.table.pivot.asc_getDataFields();
				let indexFrom = this.dataIndex;
				let indexTo = getIndexTo(type, indexFrom, fields);
				if (indexTo != null) {
					this.SetPosition(indexTo + 1);
				} else {
					private_MakeError("Bad move type.");
				}
				break;
		}
	}

	/** Attributes */

	/**
	 * Type of calculation to perform on the data field items.
	 * @typedef {"Average" | "CountNumbers" | "Count" | "Max" | "Min" | "Product" |
	 * "StdDev" | "StdDevP" | "Sum" | "Var" | "VarP"} DataConsolidateFunctionType
	 */

	/**
	 * Sets the function performed in the data field.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {DataConsolidateFunctionType} func - The function performed in the added data field.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetFunction.js
	 */
	ApiPivotDataField.prototype.SetFunction = function (func) {
		const field = new Asc.CT_DataField();
		switch (func) {
			case "Average":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Average);
				break;
			case "Count":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Count);
				break;
			case "CountNumbers":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.CountNums);
				break;
			case "Max":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Max);
				break;
			case "Min":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Min);
				break;
			case "Product":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Product);
				break;
			case "StdDev":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.StdDev);
				break;
			case "StdDevP":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.StdDevp);
				break;
			case "Sum":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Sum);
				break;
			case "Var":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Var);
				break;
			case "VarP":
				field.asc_setSubtotal(Asc.c_oAscDataConsolidateFunction.Varp);
				break;
			default:
				private_MakeError('Invalid function type.');
				return;
		}
		this.dataField.asc_set(this.table.api, this.table.pivot, this.dataIndex, field);
	};
	/**
	 * Returns the function performed in the data field.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {DataConsolidateFunctionType} func - The function performed in the added data field.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetFunction.js
	 */
	ApiPivotDataField.prototype.GetFunction = function () {
		const subtotal = this.dataField.asc_getSubtotal();
		switch (subtotal) {
			case Asc.c_oAscDataConsolidateFunction.Average:
				return "Average";
			case Asc.c_oAscDataConsolidateFunction.Count:
				return "Count";
			case Asc.c_oAscDataConsolidateFunction.CountNums:
				return "CountNumbers";
			case Asc.c_oAscDataConsolidateFunction.Max:
				return "Max";
			case Asc.c_oAscDataConsolidateFunction.Min:
				return "Min";
			case Asc.c_oAscDataConsolidateFunction.Product:
				return "Product";
			case Asc.c_oAscDataConsolidateFunction.StdDev:
				return "StdDev";
			case Asc.c_oAscDataConsolidateFunction.StdDevP:
				return "StdDevP";
			case Asc.c_oAscDataConsolidateFunction.Sum:
				return "Sum";
			case Asc.c_oAscDataConsolidateFunction.Var:
				return "Var";
			case Asc.c_oAscDataConsolidateFunction.VarP:
				return "VarP";
		}
	};

	/**
	 * Returns a value that represents the position of the field in category.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetPosition.js
	 */
	ApiPivotDataField.prototype.GetPosition = function () {
		return this.dataIndex + 1;
	};

	/**
	 * Sets a value that represents the position of the field in dataField category.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {number} position - Position.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetPosition.js
	 */
	ApiPivotDataField.prototype.SetPosition = function (position) {
		const dataFields = this.table.pivot.asc_getDataFields();
		if (typeof position === "number") {
			if (dataFields[position] && this.index !== position - 1) {
				this.table.pivot.asc_moveDataField(this.table.api, this.dataIndex, position - 1);
				this.dataIndex = position - 1;
			} else {
				private_MakeError('Invalid position (out of range or the same).');
			}
		} else {
			private_MakeError('Invalid type of "position".');
		}
	};

	Object.defineProperty(ApiPivotDataField.prototype, "Position", {
		get: function () {
			return this.GetPosition();
		},
		set: function (position) {
			this.SetPosition(position);
		}
	});
	/**
	 * Returns a data field orientation value that represents the location
	 * of the field in the specified PivotTable report.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {PivotFieldOrientationType}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetOrientation.js
	 */
	ApiPivotDataField.prototype.GetOrientation = function () {
		if (this.dataField) {
			return "Values";
		}
		return null;
	};
	Object.defineProperty(ApiPivotDataField.prototype, "Orientation", {
		get: function () {
			return this.GetOrientation();
		},
	});
	/**
	 * Returns a value representing the name of the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetValue.js
	 */
	ApiPivotDataField.prototype.GetValue = function () {
		return this.GetName();
	};

	/**
	 * Sets a value representing the name of the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetValue.js
	 */
	ApiPivotDataField.prototype.SetValue = function (name) {
		this.SetName(name);
	};

	Object.defineProperty(ApiPivotDataField.prototype, "Value", {
		get: function () {
			return this.GetValue();
		},
		set: function (name) {
			this.SetValue(name);
		}
	});
	/**
	 * Returns a value that represents the label text for the data field.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetCaption.js
	 */
	ApiPivotDataField.prototype.GetCaption = function () {
		return this.GetName();
	};
	/**
	 * Set value that represents the label text for the pivot field.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {string} caption
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetCaption.js
	 */
	ApiPivotDataField.prototype.SetCaption = function (caption) {
		return this.SetName(caption);
	};

	Object.defineProperty(ApiPivotDataField.prototype, "Caption", {
		get: function () {
			return this.GetCaption();
		},
		set: function(caption) {
			this.SetCaption(caption);
		}
	});

	/**
	 * Returns a value representing the name of the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetName.js
	 */
	ApiPivotDataField.prototype.GetName = function () {
		return this.dataField.asc_getName();
	};

	/**
	 * Sets a value representing the name of the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name.
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetName.js
	 */
	ApiPivotDataField.prototype.SetName = function (name) {
		if (typeof name === 'string' && name.length > 0) {
			const field = new Asc.CT_DataField();
			field.asc_setName(name);
			this.dataField.asc_set(this.table.api, this.table.pivot, this.dataIndex, field);
		} else {
			private_MakeError('Bad name type or empty.');
		}

	};

	Object.defineProperty(ApiPivotDataField.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (name) {
			this.SetName(name);
		}
	});
	/**
	 * Returns a value that represents the format code for the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetNumberFormat.js
	 */
	ApiPivotDataField.prototype.GetNumberFormat = function () {
		return this.dataField.asc_getNumFormat();
	};
	/**
	 * Sets value that represents the format code for the object.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @param {string} format
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/SetNumberFormat.js
	 */
	ApiPivotDataField.prototype.SetNumberFormat = function (format) {
		const newField = new Asc.CT_DataField();
		newField.asc_setNumFormat(format);
		this.dataField.asc_set(this.table.api, this.table.pivot, this.dataIndex, newField);
	};
	Object.defineProperty(ApiPivotDataField.prototype, "NumberFormat", {
		get: function () {
			return this.GetNumberFormat();
		},
		set: function (format) {
			return this.SetNumberFormat(format);
		}
	});
	/**
	 * Returns index of the data field.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetIndex.js
	 */
	ApiPivotDataField.prototype.GetIndex = function () {
		return this.dataIndex + 1;
	};

	Object.defineProperty(ApiPivotDataField.prototype, "Index", {
		get: function () {
			return this.GetIndex();
		}
	});
	/**
	 * Returns the pivotField from which the data field was created.
	 * @memberof ApiPivotDataField
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField}
	 * @see office-js-api/Examples/ApiPivotDataField/Methods/GetPivotField.js
	 */
	ApiPivotDataField.prototype.GetPivotField = function () {
		return new ApiPivotField(this.table, this.index, this.pivotField);
	};

	Object.defineProperty(ApiPivotDataField.prototype, "PivotField", {
		get: function () {
			return this.GetPivotField();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPivotField
	//
	//------------------------------------------------------------------------------------------------------------------

	/** Methods */

	/**
	 * Calling this method deletes all filters currently applied to the PivotField.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotField/Methods/ClearAllFilters.js
	 */
	ApiPivotField.prototype.ClearAllFilters  = function () {
		this.table.pivot.removeFiltersWithLock(this.table.api, [this.index], false);
	};
	/**
	 * This method deletes all label filters or all date filters in the PivotFilters collection.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotField/Methods/ClearLabelFilters.js
	 */
	ApiPivotField.prototype.ClearLabelFilters  = function () {
		this.table.pivot.asc_removePivotFilter(this.table.api, this.index, false, true, false);
	};
	/**
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotField/Methods/ClearManualFilters.js
	 */
	ApiPivotField.prototype.ClearManualFilters  = function () {
		this.table.pivot.asc_removePivotFilter(this.table.api, this.index, true, false, false);
	};
	/**
	 * Calling this method deletes all value filters in the PivotFilters collection of the PivotField.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotField/Methods/ClearValueFilters.js
	 */
	ApiPivotField.prototype.ClearValueFilters  = function () {
		this.table.pivot.asc_removePivotFilter(this.table.api, this.index, false, false, true);
	};
	/**
	 * Returns an object that represents either a single PivotTable item (a PivotItem object)
	 * or a collection of all the visible and hidden items (a PivotItems object) in the specified field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {number} [index]
	 * @returns {ApiPivotItem[] | ApiPivotItem | null}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetPivotItems.js
	 */
	ApiPivotField.prototype.GetPivotItems = function (index) {
		const pivotFields = this.table.pivot.asc_getPivotFields();
		const pivotField = pivotFields[this.index];
		if (index != null) {
			const item = pivotField[index];
			if (item && item.t === Asc.c_oAscItemType.Data) {
				return new ApiPivotItem(this, item);
			}
			private_MakeError('Invalid item index.');
			return null;
		}
		const items = pivotField.getItems();
		const t = this;
		return items.filter(function (item) {
			return Asc.c_oAscItemType.Data === item.t;
		}).map(function (item, index) {
			return new ApiPivotItem(t, item);
		})
	};

	Object.defineProperty(ApiPivotField.prototype, "PivotItems", {
		get: function (index) {
			return this.GetPivotItems(index);
		}
	});
	/**
	 * Moves field inside category.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {PivotMoveFieldType | PivotFieldOrientationType} type - The type of the field to move.
	 * @param {number | undefined} index - The index of the field in new category.
	 * @see office-js-api/Examples/ApiPivotField/Methods/Move.js
	 */
	ApiPivotField.prototype.Move = function (type, index) {
		function getIndexTo(type, indexFrom, fields) {
			switch (type) {
				case "Up":
					return (indexFrom > 0) ? indexFrom - 1 : indexFrom;
				case "Down":
					return (indexFrom < fields.length - 1) ? indexFrom + 1 : fields.length - 1;
				case "Begin":
					return 0;
				case "End":
					return fields.length - 1;
				default:
					return null;
			}
		}
		if (index == null) {
			index = 0;
		}
		switch (type) {
			case "Rows":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisRow) {
					this.table.pivot.asc_moveToRowField(this.table.api, this.index, undefined, index - 1);
				} else {
					this.SetPosition(index)
				}
				break;
			case "Columns":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisCol) {
					this.table.pivot.asc_moveToColField(this.table.api, this.index, undefined, index - 1);
				} else {
					this.SetPosition(index)
				}
				break;
			case "Filters":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisPage) {
					this.table.pivot.asc_moveToPageField(this.table.api, this.index, undefined, index - 1);
				} else {
					this.SetPosition(index)
				}
				break;
			case "Values":
				this.table.pivot.asc_moveToDataField(this.table.api, this.index, undefined, index - 1);
				break;
			case "Hidden":
				this.Remove();
				break;
			default:
				const fields = this.table.pivot.getAxisFields(this.pivotField.axis);
				if (fields) {
					let indexFrom = null;
					for (let i = 0; i < fields.length; i += 1) {
						if (fields[i].asc_getIndex() === this.index) {
							indexFrom = i;
							break;
						}
					}
					let indexTo = getIndexTo(type, indexFrom, fields);
					if (indexTo != null) {
						this.SetPosition(indexTo + 1);
					} else {
						private_MakeError("Bad move type.");
					}
				} else {
					private_MakeError("Field is hidden.");
				}
				break;
		}
	};
	/**
	 * Removes field from PivotTable.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @see office-js-api/Examples/ApiPivotField/Methods/Remove.js
	 */
	ApiPivotField.prototype.Remove = function () {
		this.table.pivot.asc_removeNoDataField(this.table.api, this.index);
	};

	/** Attributes */

	/**
	 * Returns a value that represents the position of the field in category.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetPosition.js
	 */
	ApiPivotField.prototype.GetPosition = function () {
		const fields = this.table.pivot.getAxisFields(this.pivotField.axis);
		if (fields) {
			for (let i = 0; i < fields.length; i += 1) {
				if (fields[i].asc_getIndex() === this.index) {
					return i + 1;
				}
			}
		}
		private_MakeError('The field is hidden.\n' +
		'If you need to get the position of the data field then use ApiPivotDataField.GetPosition.\n' +
		'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ' +
		'ApiPivotDataField object');
	};

	/**
	 * Sets a value that represents the position of the field
	 * among all the fields in its orientation (Rows, Columns, Pages).
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {number} position - Position.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetPosition.js
	 */
	ApiPivotField.prototype.SetPosition = function (position) {
		if (typeof position == "number") {
			if (this.pivotField.axis === null) {
				private_MakeError('The field is hidden.\n' +
					'If you need to set the position of the data field then use ApiPivotDataField.SetPosition.\n' +
					'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ' +
					'ApiPivotDataField object');
				return;
			}
			if (!this.table.pivot.moveFieldInAxis(this.table.api, this.index, this.pivotField.axis, position - 1)) {
				private_MakeError('Invalid position (out of range or the same).')
			}
		} else {
			private_MakeError('Invalid type of "position".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "Position", {
		get: function () {
			return this.GetPosition();
		},
		set: function (position) {
			this.SetPosition(position);
		}
	});

	/**
	 * Returns a pivot field orientation value that represents the location
	 * of the field in the specified PivotTable report.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {PivotFieldOrientationType}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetOrientation.js
	 */
	ApiPivotField.prototype.GetOrientation = function () {
		if (this.pivotField.axis === Asc.c_oAscAxis.AxisRow) {
			return "Rows";
		} else if (this.pivotField.axis === Asc.c_oAscAxis.AxisCol) {
			return "Columns";
		} else if (this.pivotField.axis === Asc.c_oAscAxis.AxisPage) {
			return "Filters"
		} else {
			return "Hidden"
		}
	};

	/**
	 * Sets a pivot field orientation value that represents the location
	 * of the field in the specified PivotTable report.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {PivotFieldOrientationType} type - Field orientation type.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetOrientation.js
	 */
	ApiPivotField.prototype.SetOrientation = function (type) {
		switch (type) {
			case "Rows":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisRow) {
					this.table.pivot.asc_moveToRowField(this.table.api, this.index);
				} else {
					private_MakeError('The field already has that orientation.')
				}
				break;
			case "Columns":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisCol) {
					this.table.pivot.asc_moveToColField(this.table.api, this.index);
				} else {
					private_MakeError('The field already has that orientation.')
				}
				break;
			case "Filters":
				if (this.pivotField.axis !== Asc.c_oAscAxis.AxisPage) {
					this.table.pivot.asc_moveToPageField(this.table.api, this.index);
				} else {
					private_MakeError('The field already has that orientation.')
				}
				break;
			case "Values":
				this.table.pivot.asc_moveToDataField(this.table.api, this.index);
				break;
			case "Hidden":
				this.Remove();
				break;
			default:
				private_MakeError('Invalid "type" value.');
				break;
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "Orientation", {
		get: function () {
			return this.GetOrientation();
		},
		set: function (type) {
			this.SetOrientation(type);
		}
	});
	/**
	 * Returns a value representing the name of the object.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetValue.js
	 */
	ApiPivotField.prototype.GetValue = function () {
		return this.GetName();
	};

	/**
	 * Sets a value representing the name of the object.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetValue.js
	 */
	ApiPivotField.prototype.SetValue = function (name) {
		this.SetName(name)
	};

	Object.defineProperty(ApiPivotField.prototype, "Value", {
		get: function () {
			return this.GetValue();
		},
		set: function (name) {
			this.SetValue(name);
		}
	});
	/**
	 * Returns a value that represents the label text for the pivot field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetCaption.js
	 */
	ApiPivotField.prototype.GetCaption = function () {
		return this.GetName();
	};
	/**
	 * Set value that represents the label text for the pivot field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {string} caption
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetCaption.js
	 */
	ApiPivotField.prototype.SetCaption = function (caption) {
		return this.SetName(caption);
	};

	Object.defineProperty(ApiPivotField.prototype, "Caption", {
		get: function () {
			return this.GetCaption();
		},
		set: function(caption) {
			this.SetCaption(caption);
		}
	});

	/**
	 * Returns a value representing the name of the object.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetName.js
	 */
	ApiPivotField.prototype.GetName = function () {
		return this.pivotField.asc_getName() || this.GetSourceName();
	};

	/**
	 * Sets a value representing the name of the object.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {string} name - Name.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetName.js
	 */
	ApiPivotField.prototype.SetName = function (name) {
		if (typeof name === 'string' && name.length > 0) {
			const field = new Asc.CT_PivotField();
			field.asc_setName(name);
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Bad name type or empty.');
		}

	};

	Object.defineProperty(ApiPivotField.prototype, "Name", {
		get: function () {
			return this.GetName();
		},
		set: function (name) {
			this.SetName(name);
		}
	});
	/**
	 * Returns source name
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetSourceName.js
	 */
	ApiPivotField.prototype.GetSourceName = function () {
		return this.table.pivot.getCacheFieldName(this.index);
	};

	Object.defineProperty(ApiPivotField.prototype, "SourceName", {
		get: function () {
			return this.GetSourceName();
		}
	});

	/**
	 * Returns index of the PivotField.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetIndex.js
	 */
	ApiPivotField.prototype.GetIndex = function () {
		return this.index + 1;
	};

	Object.defineProperty(ApiPivotField.prototype, "Index", {
		get: function () {
			return this.GetIndex();
		}
	});

	/**
	 * Returns ApiPivotTable.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotTable}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetTable.js
	 */
	ApiPivotField.prototype.GetTable = function () {
		return this.table;
	};

	Object.defineProperty(ApiPivotField.prototype, "Table", {
		get: function () {
			return this.GetTable();
		}
	});

	/**
	 * Returns parent object to this field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotTable}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetParent.js
	 */
	ApiPivotField.prototype.GetParent = function () {
		return this.table;
	};

	Object.defineProperty(ApiPivotField.prototype, "Parent", {
		get: function () {
			return this.GetParent();
		}
	});

	/**
	 * Returns whether or not a PivotField is compacted.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutCompactRow.js
	 */
	ApiPivotField.prototype.GetLayoutCompactRow = function () {
		const pivField = this.table.pivot.asc_getPivotFields()[this.index];
		return (pivField.asc_getOutline() && pivField.asc_getCompact());
	};

	/**
	 * Sets whether or not a PivotField is compacted.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} compact - Use compact form or not.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutCompactRow.js
	 */
	ApiPivotField.prototype.SetLayoutCompactRow = function (compact) {
		if (typeof compact == "boolean") {
			const field = new Asc.CT_PivotField();
			const pivField = this.table.pivot.asc_getPivotFields()[this.index];
			field.asc_setCompact( (pivField.asc_getOutline() && compact) );
			pivField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "compact".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutCompactRow", {
		get: function () {
			return this.GetLayoutCompactRow();
		},
		set: function (compact) {
			this.SetLayoutCompactRow(compact);
		}
	});

	/**
	 * @typedef {"Tabular" | "Outline"} PivotLayoutType
	 */

	/**
	 * Returns the way the specified PivotTable items appear—in table format or in outline format.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {PivotLayoutType}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutForm.js
	 */
	ApiPivotField.prototype.GetLayoutForm = function () {
		return this.pivotField.asc_getOutline() ? "Outline" : "Tabular";
	};

	/**
	 * Sets the way the specified PivotTable items appear—in table format or in outline format.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {PivotLayoutType} type - Type of layot report form.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutForm.js
	 */
	ApiPivotField.prototype.SetLayoutForm = function (type) {
		if (type === "Tabular" || type === "Outline") {
			const newField = new Asc.CT_PivotField();
			newField.asc_setOutline(type === "Outline");
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, newField);
		} else {
			private_MakeError('Invalid type of "type" or invalid value.')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutForm", {
		get: function () {
			return this.GetLayoutForm();
		},
		set: function (type) {
			this.SetLayoutForm(type);
		}
	});

	/**
	 * Returns the flag that represents page break is inserted after each field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutPageBreak.js
	 */
	ApiPivotField.prototype.GetLayoutPageBreak = function () {
		return this.pivotField.insertPageBreak;
	};

	/**
	 * Sets the flag that represents page break is inserted after each field.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} insert - Flag insert or not page break after each field.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutPageBreak.js
	 */
	ApiPivotField.prototype.SetLayoutPageBreak = function (insert) {
		if ( typeof insert == "boolean") {
			this.pivotField.insertPageBreak = insert;
		} else {
			private_MakeError('Invalid type of "insert".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutPageBreak", {
		get: function () {
			return this.GetLayoutPageBreak();
		},
		set: function (type) {
			this.SetLayoutPageBreak(type);
		}
	});

	/**
	 * Returns if the PivotField is currently visible in the PivotTable or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetShowingInAxis.js
	 */
	ApiPivotField.prototype.GetShowingInAxis = function () {
		return this.pivotField.axis !== null || this.pivotField.dataField;
	};

	Object.defineProperty(ApiPivotField.prototype, "ShowingInAxis", {
		get: function () {
			return this.GetShowingInAxis();
		}
	});

	/**
	 * Returns setting repeat items labels at each row.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetRepeatLabels.js
	 */
	ApiPivotField.prototype.GetRepeatLabels = function () {
		return this.pivotField.asc_getFillDownLabelsDefault();
	};

	/**
	 * Sets setting repeat items labels at each row.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} repeat - Flag repeat items labels at each row or not.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetRepeatLabels.js
	 */
	ApiPivotField.prototype.SetRepeatLabels = function (repeat) {
		if (typeof repeat == "boolean") {
			const field = new Asc.CT_PivotField();
			field.asc_setFillDownLabelsDefault(repeat);
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "repeat".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "RepeatLabels", {
		get: function () {
			return this.GetRepeatLabels();
		},
		set: function (repeat) {
			this.SetRepeatLabels(repeat);
		}
	});

	/**
	 * Returns setting insert blank rows after each item.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutBlankLine.js
	 */
	ApiPivotField.prototype.GetLayoutBlankLine = function () {
		return this.pivotField.asc_getInsertBlankRow();
	};

	/**
	 * Sets setting insert blank rows after each item.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} insert - Flag insert blank rows after each item or not.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutBlankLine.js
	 */
	ApiPivotField.prototype.SetLayoutBlankLine = function (insert) {
		if (typeof insert == "boolean") {
			const field = new Asc.CT_PivotField();
			field.asc_setInsertBlankRow(insert);
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "insert".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutBlankLine", {
		get: function () {
			return this.GetLayoutBlankLine();
		},
		set: function (insert) {
			this.SetLayoutBlankLine(insert);
		}
	});

	/**
	 * Returns setting show items with no data.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetShowAllItems.js
	 */
	ApiPivotField.prototype.GetShowAllItems = function () {
		return this.pivotField.asc_getShowAll();
	};

	/**
	 * Sets setting show items with no data.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show items with no data or not.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetShowAllItems.js
	 */
	ApiPivotField.prototype.SetShowAllItems = function (show) {
		if (typeof show == "boolean") {
			const field = new Asc.CT_PivotField();
			field.asc_setShowAll(show);
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "show".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "ShowAllItems", {
		get: function () {
			return this.GetShowAllItems();
		},
		set: function (show) {
			this.SetShowAllItems(show);
		}
	});

	/**
	 * Returns setting show subtotals.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutSubtotals.js
	 */
	ApiPivotField.prototype.GetLayoutSubtotals = function () {
		return this.pivotField.asc_getDefaultSubtotal();
	};

	/**
	 * Set setting show subtotals.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} show - Flag show subtotals or not.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutSubtotals.js
	 */
	ApiPivotField.prototype.SetLayoutSubtotals = function (show) {
		if (typeof show == "boolean") {
			const field = new Asc.CT_PivotField();
			field.asc_setDefaultSubtotal(show);
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "show".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutSubtotals", {
		get: function () {
			return this.GetLayoutSubtotals();
		},
		set: function (show) {
			this.SetLayoutSubtotals(show);
		}
	});

	/** @typedef { "Top" | "Bottom" } LayoutSubtotalLocationType */

	/**
	 * Returns setting layout subtotal location.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {LayoutSubtotalLocationType}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetLayoutSubtotalLocation.js
	 */
	ApiPivotField.prototype.GetLayoutSubtotalLocation = function () {
		return ( this.pivotField.asc_getSubtotalTop() ? "Top" : "Bottom" );
	};

	/**
	 * Sets setting layout subtotal location.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {LayoutSubtotalLocationType} type - Type of layot subtotal location
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetLayoutSubtotalLocation.js
	 */
	ApiPivotField.prototype.SetLayoutSubtotalLocation = function (type) {
		if (typeof type == "string" && ( type == "Top" || type == "Bottom")) {
			const field = new Asc.CT_PivotField();
			field.asc_setSubtotalTop( (type == "Top") );
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "type" or invalid value.')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "LayoutSubtotalLocation", {
		get: function () {
			return this.GetLayoutSubtotalLocation();
		},
		set: function (type) {
			this.SetLayoutSubtotalLocation(type);
		}
	});

	/**
	 * Returns the text string label displayed in the subtotal column or row heading in the specified PivotTable report.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetSubtotalName.js
	 */
	ApiPivotField.prototype.GetSubtotalName = function () {
		return (this.pivotField.subtotalCaption);
	};

	/**
	 * Sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {string} caption - Caption.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetSubtotalName.js
	 */
	ApiPivotField.prototype.SetSubtotalName = function (caption) {
		if ( typeof caption == "string") {
			const field = new Asc.CT_PivotField();
			field.subtotalCaption = caption.trim();
			this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
		} else {
			private_MakeError('Invalid type of "caption".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "SubtotalName", {
		get: function () {
			return this.GetSubtotalName();
		},
		set: function (caption) {
			this.SetSubtotalName(caption);
		}
	});

	/**
	 * Subtotal pivot field types
	 * @typedef {Object} PivotFieldSubtotals
	 * @property {boolean} Sum
	 * @property {boolean} Count
	 * @property {boolean} Average
	 * @property {boolean} Max
	 * @property {boolean} Min
	 * @property {boolean} Product
	 * @property {boolean} CountNumbers
	 * @property {boolean} StdDev
	 * @property {boolean} StdDevP
	 * @property {boolean} Var
	 * @property {boolean} VarP
	 */

	/**
	 * Returns object that represent all subtotals.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {PivotFieldSubtotals}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetSubtotals.js
	 */
	ApiPivotField.prototype.GetSubtotals = function () {
		const res = {
			'Sum': false,
			'Count': false,
			'Average': false,
			'Max': false,
			'Min': false,
			'Product': false,
			'CountNumbers': false,
			'StdDev': false,
			'StdDevP': false,
			'Var': false,
			'VarP': false
		};
		if (this.pivotField.asc_getDefaultSubtotal()) {
			const subtotals = this.pivotField.asc_getSubtotals();
			for (var i = 0; i < subtotals.length; i++) {
				switch (subtotals[i]) {
					case Asc.c_oAscItemType.Sum:
						res['Sum'] = true;
						break;
					case Asc.c_oAscItemType.CountA:
						res['Count'] = true;
						break;
					case Asc.c_oAscItemType.Avg:
						res['Average'] = true;
						break;
					case Asc.c_oAscItemType.Max:
						res['Max'] = true;
						break;
					case Asc.c_oAscItemType.Min:
						res['Min'] = true;
						break;
					case Asc.c_oAscItemType.Product:
						res['Product'] = true;
						break;
					case Asc.c_oAscItemType.Count:
						res['CountNumbers'] = true;
						break;
					case Asc.c_oAscItemType.StdDev:
						res['StdDev'] = true;
						break;
					case Asc.c_oAscItemType.StdDevP:
						res['StdDevP'] = true;
						break;
					case Asc.c_oAscItemType.Var:
						res['Var'] = true;
						break;
					case Asc.c_oAscItemType.VarP:
						res['VarP'] = true;
						break;
				}
			}
		}
		return res;
	};

	/**
	 * Sets Subtotals.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {PivotFieldSubtotals} subtotals - Object that represent all subtotals or some of them.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetSubtotals.js
	 */
	ApiPivotField.prototype.SetSubtotals = function (subtotals) {
			if (typeof subtotals == "object") {
				const field = new Asc.CT_PivotField();
				const arr = [];
				if (this.pivotField.asc_getDefaultSubtotal()) {
					if (subtotals.hasOwnProperty('Sum') && subtotals['Sum']) {
						arr.push(Asc.c_oAscItemType.Sum);
					}
					if (subtotals.hasOwnProperty('Count') && subtotals['Count']) {
						arr.push(Asc.c_oAscItemType.CountA);
					}
					if (subtotals.hasOwnProperty('Average') && subtotals['Average']) {
						arr.push(Asc.c_oAscItemType.Avg);
					}
					if (subtotals.hasOwnProperty('Max') && subtotals['Max']) {
						arr.push(Asc.c_oAscItemType.Max);
					}
					if (subtotals.hasOwnProperty('Min') && subtotals['Min']) {
						arr.push(Asc.c_oAscItemType.Min);
					}
					if (subtotals.hasOwnProperty('Product') && subtotals['Product']) {
						arr.push(Asc.c_oAscItemType.Product);
					}
					if (subtotals.hasOwnProperty('CountNumbers') && subtotals['CountNumbers']) {
						arr.push(Asc.c_oAscItemType.Count);
					}
					if (subtotals.hasOwnProperty('StdDev') && subtotals['StdDev']) {
						arr.push(Asc.c_oAscItemType.StdDev);
					}
					if (subtotals.hasOwnProperty('StdDevP') && subtotals['StdDevP']) {
						arr.push(Asc.c_oAscItemType.StdDevP);
					}
					if (subtotals.hasOwnProperty('Var') && subtotals['Var']) {
						arr.push(Asc.c_oAscItemType.Var);
					}
					if (subtotals.hasOwnProperty('VarP') && subtotals['VarP']) {
						arr.push(Asc.c_oAscItemType.VarP);
					}
				}
				if (arr.length) {
					field.asc_setSubtotals(arr);
					this.pivotField.asc_set(this.table.api, this.table.pivot, this.index, field);
				}
			} else {
				private_MakeError('Invalid type of "subtotals".')
			}
	};

	Object.defineProperty(ApiPivotField.prototype, "Subtotals", {
		get: function () {
			return this.GetSubtotals();
		},
		set: function (subtotals) {
			this.SetSubtotals(subtotals);
		}
	});

	/**
	 * Returns the specified field, can be dragged to the column position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetDragToColumn.js
	 */
	ApiPivotField.prototype.GetDragToColumn = function () {
		return this.pivotField.dragToCol;
	};

	/**
	 * Sets the specified field, can be dragged to the column position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} flag - Flag whether this field can be moved to columns.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetDragToColumn.js
	 */
	ApiPivotField.prototype.SetDragToColumn = function (flag) {
		if (typeof flag == "boolean") {
			this.pivotField.dragToCol = flag;
		} else {
			private_MakeError('Invalid type of "flag".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "DragToColumn", {
		get: function () {
			return this.GetDragToColumn();
		},
		set: function (flag) {
			this.SetDragToColumn(flag);
		}
	});

	/**
	 * Returns the specified field, can be dragged to the row position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetDragToRow.js
	 */
	ApiPivotField.prototype.GetDragToRow = function () {
		return this.pivotField.dragToRow;
	};

	/**
	 * Sets the specified field, can be dragged to the row position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} flag - Flag whether this field can be moved to rows.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetDragToRow.js
	 */
	ApiPivotField.prototype.SetDragToRow = function (flag) {
		if (typeof flag == "boolean") {
			this.pivotField.dragToRow = flag;
		} else {
			private_MakeError('Invalid type of "flag".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "DragToRow", {
		get: function () {
			return this.GetDragToRow();
		},
		set: function (flag) {
			this.SetDragToRow(flag);
		}
	});

	/**
	 * Returns the specified field, can be dragged to the data position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetDragToData.js
	 */
	ApiPivotField.prototype.GetDragToData = function () {
		return this.pivotField.dragToData;
	};

	/**
	 * Sets the specified field, can be dragged to the data position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} flag - Flag whether this field can be moved to data.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetDragToData.js
	 */
	ApiPivotField.prototype.SetDragToData = function (flag) {
		if (typeof flag == "boolean") {
			this.pivotField.dragToData = flag;
		} else {
			private_MakeError('Invalid type of "flag".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "DragToData", {
		get: function () {
			return this.GetDragToData();
		},
		set: function (flag) {
			this.SetDragToData(flag);
		}
	});

	/**
	 * Returns the specified field, can be dragged to the page position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetDragToPage.js
	 */
	ApiPivotField.prototype.GetDragToPage = function () {
		return this.pivotField.dragToPage;
	};

	/**
	 * Sets the specified field, can be dragged to the page position or not.
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @param {boolean} flag - Flag whether this field can be moved to page.
	 * @see office-js-api/Examples/ApiPivotField/Methods/SetDragToPage.js
	 */
	ApiPivotField.prototype.SetDragToPage = function (flag) {
		if (typeof flag == "boolean") {
			this.pivotField.dragToPage = flag;
		} else {
			private_MakeError('Invalid type of "flag".')
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "DragToPage", {
		get: function () {
			return this.GetDragToPage();
		},
		set: function (flag) {
			this.SetDragToPage(flag);
		}
	});

	/**
	 * Returns the current page showing for the page field (valid only for page fields).
	 * @memberof ApiPivotField
	 * @typeofeditors ["CSE"]
	 * @returns {string | number}
	 * @see office-js-api/Examples/ApiPivotField/Methods/GetCurrentPage.js
	 */
	ApiPivotField.prototype.GetCurrentPage = function () {
		const pageFields = this.table.pivot.asc_getPageFields();
		const t = this;
		const pageIndex = pageFields.findIndex(function(pageField) {
			return pageField.asc_getIndex() === t.index;
		});
		if (this.pivotField.axis === Asc.c_oAscAxis.AxisPage) {
			const val = this.table.pivot.getPageFieldCellValue(pageIndex);
			return val.text || val.number || val.multiText;
		} else {
			private_MakeError("It is not possible from this field.");
		}
	};

	Object.defineProperty(ApiPivotField.prototype, "CurrentPage", {
		get: function () {
			return this.GetCurrentPage();
		}
	});

	ApiPivotField.prototype.GetFunction = function () {
		private_MakeError('This method can only be called on a data field.\n' +
			'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ApiPivotDataField object');
		return null;
	};
	ApiPivotField.prototype.SetFunction = function () {
		private_MakeError('This method can only be called on a data field.\n' +
			'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ApiPivotDataField object');
		return null;
	};

	Object.defineProperty(ApiPivotField.prototype, "Function", {
		get: function () {
			return this.GetFunction();
		},
		set: function () {
			this.SetFunction()
		}
	});

	ApiPivotField.prototype.GetNumberFormat = function () {
		private_MakeError('This method can only be called on a data field.\n' +
			'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ApiPivotDataField object');
		return null;
	};
	ApiPivotField.prototype.SetNumberFormat = function () {
		private_MakeError('This method can only be called on a data field.\n' +
			'See ApiPivotTable.GetDataFields or ApiPivotTable.GetPivotFields with dataField identifier to get ApiPivotDataField object');
		return null;
	};

	Object.defineProperty(ApiPivotField.prototype, "NumberFormat", {
		get: function () {
			return this.GetNumberFormat();
		},
		set: function () {
			this.SetNumberFormat()
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPivotItem
	//
	//------------------------------------------------------------------------------------------------------------------

	/* Attributes */

	/**
	 * Returns a String value representing the name of the object.
	 * @memberof ApiPivotItem
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotItem/Methods/GetName.js
	 */
	ApiPivotItem.prototype.GetName = function () {
		if (this.pivotItem.n) {
			return this.pivotItem.n;
		}
		const pivot = this.field.table.pivot;
		const cacheField = pivot.asc_getCacheFields()[this.field.index];
		const sharedItem = cacheField.getGroupOrSharedItem(this.pivotItem.x);
		if (sharedItem) {
			return sharedItem.getCellValue().getTextValue();
		}
	};

	Object.defineProperty(ApiPivotItem.prototype, "Name", {
		get: function () {
			return this.GetName();
		}
	});

	/**
	 * Returns a String value that represents the label text for the pivot item.
	 * @memberof ApiPivotItem
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotItem/Methods/GetCaption.js
	 */
	ApiPivotItem.prototype.GetCaption = function () {
		return this.GetName();
	};

	Object.defineProperty(ApiPivotItem.prototype, "Caption", {
		get: function () {
			return this.GetCaption();
		}
	});

	/**
	 * Returns a String value that represents the name of the specified item in the PivotTable field.
	 * @memberof ApiPivotItem
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 * @see office-js-api/Examples/ApiPivotItem/Methods/GetValue.js
	 */
	ApiPivotItem.prototype.GetValue = function () {
		return this.GetName();
	};

	Object.defineProperty(ApiPivotItem.prototype, "Value", {
		get: function () {
			return this.GetValue();
		}
	});

	/**
	 * Returns the parent object for the specified object.
	 * @memberof ApiPivotItem
	 * @typeofeditors ["CSE"]
	 * @returns {ApiPivotField}
	 * @see office-js-api/Examples/ApiPivotItem/Methods/GetParent.js
	 */
	ApiPivotItem.prototype.GetParent = function () {
		return this.field;
	};

	Object.defineProperty(ApiPivotItem.prototype, "Parent", {
		get: function () {
			return this.GetParent();
		}
	});

	Api.prototype["Format"]                = Api.prototype.Format;
	Api.prototype["AddSheet"]              = Api.prototype.AddSheet;
	Api.prototype["GetSheets"]             = Api.prototype.GetSheets;
	Api.prototype["GetActiveSheet"]        = Api.prototype.GetActiveSheet;
	Api.prototype["GetLocale"]             = Api.prototype.GetLocale;
	Api.prototype["SetLocale"]             = Api.prototype.SetLocale;
	Api.prototype["GetSheet"]              = Api.prototype.GetSheet;
	Api.prototype["GetThemesColors"]       = Api.prototype.GetThemesColors;
	Api.prototype["SetThemeColors"]        = Api.prototype.SetThemeColors;
	Api.prototype["CreateNewHistoryPoint"] = Api.prototype.CreateNewHistoryPoint;
	Api.prototype["CreateColorFromRGB"]    = Api.prototype.CreateColorFromRGB;
	Api.prototype["CreateColorByName"]     = Api.prototype.CreateColorByName;
	Api.prototype["Intersect"]             = Api.prototype.Intersect;
	Api.prototype["GetSelection"]          = Api.prototype.GetSelection;
	Api.prototype["AddDefName"]            = Api.prototype.AddDefName;
	Api.prototype["GetDefName"]            = Api.prototype.GetDefName;
	Api.prototype["Save"]                  = Api.prototype.Save;
	Api.prototype["GetMailMergeData"]      = Api.prototype.GetMailMergeData;
	
	Api.prototype["GetRange"] = Api.prototype.GetRange;

	Api.prototype["RecalculateAllFormulas"] = Api.prototype.RecalculateAllFormulas;
	Api.prototype["AddComment"]  = Api.prototype.AddComment;
	Api.prototype["GetComments"] = Api.prototype.GetComments;
	Api.prototype["GetAllComments"] = Api.prototype.GetAllComments;
	Api.prototype["GetCommentById"] = Api.prototype.GetCommentById;
	Api.prototype["SetFreezePanesType"] = Api.prototype.SetFreezePanesType;
	Api.prototype["GetFreezePanesType"] = Api.prototype.GetFreezePanesType;
	Api.prototype["GetDocumentInfo"] = Api.prototype.GetDocumentInfo;

	Api.prototype["AddCustomFunction"] = Api.prototype.AddCustomFunction;
	Api.prototype["RemoveCustomFunction"] = Api.prototype.RemoveCustomFunction;
	Api.prototype["ClearCustomFunctions"] = Api.prototype.ClearCustomFunctions;
	Api.prototype["AddCustomFunctionLibrary"] = Api.prototype.AddCustomFunctionLibrary;

	Api.prototype["GetReferenceStyle"] = Api.prototype.GetReferenceStyle;
	Api.prototype["SetReferenceStyle"] = Api.prototype.SetReferenceStyle;

	Api.prototype["GetWorksheetFunction"] = Api.prototype.GetWorksheetFunction;
	Api.prototype["InsertPivotExistingWorksheet"] = Api.prototype.InsertPivotExistingWorksheet;
	Api.prototype["InsertPivotNewWorksheet"] = Api.prototype.InsertPivotNewWorksheet;
	Api.prototype["GetPivotByName"] = Api.prototype.GetPivotByName;
	Api.prototype["RefreshAllPivots"] = Api.prototype.RefreshAllPivots;
	Api.prototype["GetAllPivotTables"] = Api.prototype.GetAllPivotTables;

	ApiWorksheet.prototype["GetVisible"] = ApiWorksheet.prototype.GetVisible;
	ApiWorksheet.prototype["SetVisible"] = ApiWorksheet.prototype.SetVisible;
	ApiWorksheet.prototype["SetActive"] = ApiWorksheet.prototype.SetActive;		
	ApiWorksheet.prototype["GetActiveCell"] = ApiWorksheet.prototype.GetActiveCell;
	ApiWorksheet.prototype["GetSelection"] = ApiWorksheet.prototype.GetSelection;
	ApiWorksheet.prototype["GetCells"] = ApiWorksheet.prototype.GetCells;
	ApiWorksheet.prototype["GetCols"] = ApiWorksheet.prototype.GetCols;
	ApiWorksheet.prototype["GetRows"] = ApiWorksheet.prototype.GetRows;
	ApiWorksheet.prototype["GetUsedRange"] = ApiWorksheet.prototype.GetUsedRange;
	ApiWorksheet.prototype["GetName"] = ApiWorksheet.prototype.GetName;
	ApiWorksheet.prototype["SetName"] = ApiWorksheet.prototype.SetName;
	ApiWorksheet.prototype["GetIndex"] = ApiWorksheet.prototype.GetIndex;
	ApiWorksheet.prototype["GetRange"] = ApiWorksheet.prototype.GetRange;
	ApiWorksheet.prototype["GetRangeByNumber"] = ApiWorksheet.prototype.GetRangeByNumber;
	ApiWorksheet.prototype["FormatAsTable"] = ApiWorksheet.prototype.FormatAsTable;
	ApiWorksheet.prototype["SetColumnWidth"] = ApiWorksheet.prototype.SetColumnWidth;
	ApiWorksheet.prototype["SetRowHeight"] = ApiWorksheet.prototype.SetRowHeight;
	ApiWorksheet.prototype["SetDisplayGridlines"] = ApiWorksheet.prototype.SetDisplayGridlines;
	ApiWorksheet.prototype["SetDisplayHeadings"] = ApiWorksheet.prototype.SetDisplayHeadings;
	ApiWorksheet.prototype["SetLeftMargin"] = ApiWorksheet.prototype.SetLeftMargin;
	ApiWorksheet.prototype["GetLeftMargin"] = ApiWorksheet.prototype.GetLeftMargin;	
	ApiWorksheet.prototype["SetRightMargin"] = ApiWorksheet.prototype.SetRightMargin;
	ApiWorksheet.prototype["GetRightMargin"] = ApiWorksheet.prototype.GetRightMargin;
	ApiWorksheet.prototype["SetTopMargin"] = ApiWorksheet.prototype.SetTopMargin;
	ApiWorksheet.prototype["GetTopMargin"] = ApiWorksheet.prototype.GetTopMargin;	
	ApiWorksheet.prototype["SetBottomMargin"] = ApiWorksheet.prototype.SetBottomMargin;
	ApiWorksheet.prototype["GetBottomMargin"] = ApiWorksheet.prototype.GetBottomMargin;		
	ApiWorksheet.prototype["SetPageOrientation"] = ApiWorksheet.prototype.SetPageOrientation;
	ApiWorksheet.prototype["GetPageOrientation"] = ApiWorksheet.prototype.GetPageOrientation;
	ApiWorksheet.prototype["GetPrintHeadings"] = ApiWorksheet.prototype.GetPrintHeadings;
	ApiWorksheet.prototype["SetPrintHeadings"] = ApiWorksheet.prototype.SetPrintHeadings;
	ApiWorksheet.prototype["GetPrintGridlines"] = ApiWorksheet.prototype.GetPrintGridlines;
	ApiWorksheet.prototype["SetPrintGridlines"] = ApiWorksheet.prototype.SetPrintGridlines;
	ApiWorksheet.prototype["GetDefNames"] = ApiWorksheet.prototype.GetDefNames;
	ApiWorksheet.prototype["GetDefName"] = ApiWorksheet.prototype.GetDefName;
	ApiWorksheet.prototype["AddDefName"] = ApiWorksheet.prototype.AddDefName;
	ApiWorksheet.prototype["GetComments"] = ApiWorksheet.prototype.GetComments;
	ApiWorksheet.prototype["Delete"] = ApiWorksheet.prototype.Delete;
	ApiWorksheet.prototype["SetHyperlink"] = ApiWorksheet.prototype.SetHyperlink;
	ApiWorksheet.prototype["AddChart"] = ApiWorksheet.prototype.AddChart;
	ApiWorksheet.prototype["AddShape"] = ApiWorksheet.prototype.AddShape;
	ApiWorksheet.prototype["AddImage"] = ApiWorksheet.prototype.AddImage;
	ApiWorksheet.prototype["AddOleObject"] = ApiWorksheet.prototype.AddOleObject;
	ApiWorksheet.prototype["ReplaceCurrentImage"] = ApiWorksheet.prototype.ReplaceCurrentImage;
	ApiWorksheet.prototype["AddWordArt"] = ApiWorksheet.prototype.AddWordArt;
	ApiWorksheet.prototype["GetAllDrawings"] = ApiWorksheet.prototype.GetAllDrawings;
	ApiWorksheet.prototype["GetAllImages"] = ApiWorksheet.prototype.GetAllImages;
	ApiWorksheet.prototype["GetAllShapes"] = ApiWorksheet.prototype.GetAllShapes;
	ApiWorksheet.prototype["GetAllCharts"] = ApiWorksheet.prototype.GetAllCharts;
	ApiWorksheet.prototype["GetAllOleObjects"] = ApiWorksheet.prototype.GetAllOleObjects;
	ApiWorksheet.prototype["Move"] = ApiWorksheet.prototype.Move;
	ApiWorksheet.prototype["GetFreezePanes"] = ApiWorksheet.prototype.GetFreezePanes;
	ApiWorksheet.prototype["AddProtectedRange"] = ApiWorksheet.prototype.AddProtectedRange;
	ApiWorksheet.prototype["GetProtectedRange"] = ApiWorksheet.prototype.GetProtectedRange;
	ApiWorksheet.prototype["GetAllProtectedRanges"] = ApiWorksheet.prototype.GetAllProtectedRanges;
	ApiWorksheet.prototype["Paste"] = ApiWorksheet.prototype.Paste;
	ApiWorksheet.prototype["GetPivotByName"] = ApiWorksheet.prototype.GetPivotByName;
	ApiWorksheet.prototype["GetAllPivotTables"] = ApiWorksheet.prototype.GetAllPivotTables;
	ApiWorksheet.prototype["RefreshAllPivots"] = ApiWorksheet.prototype.RefreshAllPivots;

	ApiRange.prototype["GetClassType"] = ApiRange.prototype.GetClassType;
	ApiRange.prototype["GetRow"] = ApiRange.prototype.GetRow;
	ApiRange.prototype["GetCol"] = ApiRange.prototype.GetCol;
	ApiRange.prototype["Clear"] = ApiRange.prototype.Clear;
	ApiRange.prototype["GetRows"] = ApiRange.prototype.GetRows;
	ApiRange.prototype["GetCols"] = ApiRange.prototype.GetCols;
	ApiRange.prototype["End"] = ApiRange.prototype.End;
	ApiRange.prototype["GetCells"] = ApiRange.prototype.GetCells;
	ApiRange.prototype["SetOffset"] = ApiRange.prototype.SetOffset;
	ApiRange.prototype["GetAddress"] = ApiRange.prototype.GetAddress;	
	ApiRange.prototype["GetCount"] = ApiRange.prototype.GetCount;
	ApiRange.prototype["GetValue"] = ApiRange.prototype.GetValue;
	ApiRange.prototype["SetValue"] = ApiRange.prototype.SetValue;
	ApiRange.prototype["GetFormula"] = ApiRange.prototype.GetFormula;
	ApiRange.prototype["GetValue2"] = ApiRange.prototype.GetValue2;
	ApiRange.prototype["GetText"] = ApiRange.prototype.GetText;
	ApiRange.prototype["SetFontColor"] = ApiRange.prototype.SetFontColor;
	ApiRange.prototype["GetHidden"] = ApiRange.prototype.GetHidden;
	ApiRange.prototype["SetHidden"] = ApiRange.prototype.SetHidden;	
	ApiRange.prototype["GetColumnWidth"] = ApiRange.prototype.GetColumnWidth;	
	ApiRange.prototype["SetColumnWidth"] = ApiRange.prototype.SetColumnWidth;	
	ApiRange.prototype["GetRowHeight"] = ApiRange.prototype.GetRowHeight;
	ApiRange.prototype["SetRowHeight"] = ApiRange.prototype.SetRowHeight;
	ApiRange.prototype["SetFontSize"] = ApiRange.prototype.SetFontSize;
	ApiRange.prototype["SetFontName"] = ApiRange.prototype.SetFontName;
	ApiRange.prototype["SetAlignVertical"] = ApiRange.prototype.SetAlignVertical;
	ApiRange.prototype["SetAlignHorizontal"] = ApiRange.prototype.SetAlignHorizontal;
	ApiRange.prototype["SetBold"] = ApiRange.prototype.SetBold;
	ApiRange.prototype["SetItalic"] = ApiRange.prototype.SetItalic;
	ApiRange.prototype["SetUnderline"] = ApiRange.prototype.SetUnderline;
	ApiRange.prototype["SetStrikeout"] = ApiRange.prototype.SetStrikeout;
	ApiRange.prototype["SetWrap"] = ApiRange.prototype.SetWrap;
	ApiRange.prototype["SetWrapText"] = ApiRange.prototype.SetWrap;	
	ApiRange.prototype["GetWrapText"] = ApiRange.prototype.GetWrapText;
	ApiRange.prototype["SetFillColor"] = ApiRange.prototype.SetFillColor;
	ApiRange.prototype["GetFillColor"] = ApiRange.prototype.GetFillColor;
	ApiRange.prototype["GetNumberFormat"] = ApiRange.prototype.GetNumberFormat;
	ApiRange.prototype["SetNumberFormat"] = ApiRange.prototype.SetNumberFormat;
	ApiRange.prototype["SetBorders"] = ApiRange.prototype.SetBorders;
	ApiRange.prototype["Merge"] = ApiRange.prototype.Merge;
	ApiRange.prototype["UnMerge"] = ApiRange.prototype.UnMerge;
	ApiRange.prototype["ForEach"] = ApiRange.prototype.ForEach;
	ApiRange.prototype["AddComment"] = ApiRange.prototype.AddComment;
	ApiRange.prototype["GetWorksheet"] = ApiRange.prototype.GetWorksheet;
	ApiRange.prototype["GetDefName"] = ApiRange.prototype.GetDefName;
	ApiRange.prototype["GetComment"] = ApiRange.prototype.GetComment;
	ApiRange.prototype["Select"] = ApiRange.prototype.Select;
	ApiRange.prototype["SetOrientation"] = ApiRange.prototype.SetOrientation;
	ApiRange.prototype["GetOrientation"] = ApiRange.prototype.GetOrientation;
	ApiRange.prototype["SetSort"] = ApiRange.prototype.SetSort;
	ApiRange.prototype["Delete"] = ApiRange.prototype.Delete;
	ApiRange.prototype["Insert"] = ApiRange.prototype.Insert;
	ApiRange.prototype["AutoFit"] = ApiRange.prototype.AutoFit;
	ApiRange.prototype["GetAreas"] = ApiRange.prototype.GetAreas;
	ApiRange.prototype["Copy"] = ApiRange.prototype.Copy;
	ApiRange.prototype["Cut"] = ApiRange.prototype.Cut;
	ApiRange.prototype["Paste"] = ApiRange.prototype.Paste;
	ApiRange.prototype["Find"] = ApiRange.prototype.Find;
	ApiRange.prototype["FindNext"] = ApiRange.prototype.FindNext;
	ApiRange.prototype["FindPrevious"] = ApiRange.prototype.FindPrevious;
	ApiRange.prototype["Replace"] = ApiRange.prototype.Replace;
	ApiRange.prototype["GetCharacters"] = ApiRange.prototype.GetCharacters;
	ApiRange.prototype["PasteSpecial"] = ApiRange.prototype.PasteSpecial;
	ApiRange.prototype["GetPivotTable"] = ApiRange.prototype.GetPivotTable;


	ApiDrawing.prototype["GetClassType"]               =  ApiDrawing.prototype.GetClassType;
	ApiDrawing.prototype["SetSize"]                    =  ApiDrawing.prototype.SetSize;
	ApiDrawing.prototype["SetPosition"]                =  ApiDrawing.prototype.SetPosition;
	ApiDrawing.prototype["GetWidth"]                   =  ApiDrawing.prototype.GetWidth;
	ApiDrawing.prototype["GetHeight"]                  =  ApiDrawing.prototype.GetHeight;
	ApiDrawing.prototype["GetLockValue"]               =  ApiDrawing.prototype.GetLockValue;
	ApiDrawing.prototype["SetLockValue"]               =  ApiDrawing.prototype.SetLockValue;

	ApiImage.prototype["GetClassType"]                 =  ApiImage.prototype.GetClassType;

	ApiShape.prototype["GetClassType"]                 =  ApiShape.prototype.GetClassType;
	ApiShape.prototype["GetDocContent"]                =  ApiShape.prototype.GetDocContent;
	ApiShape.prototype["GetContent"]                   =  ApiShape.prototype.GetContent;
	ApiShape.prototype["SetVerticalTextAlign"]         =  ApiShape.prototype.SetVerticalTextAlign;



	ApiChart.prototype["SetSeriaValues"]              =  ApiChart.prototype.SetSeriaValues;
	ApiChart.prototype["SetSeriaXValues"]             =  ApiChart.prototype.SetSeriaXValues;
	ApiChart.prototype["SetSeriaName"]                =  ApiChart.prototype.SetSeriaName;
	ApiChart.prototype["SetCatFormula"]               =  ApiChart.prototype.SetCatFormula;
	ApiChart.prototype["AddSeria"]                    =  ApiChart.prototype.AddSeria;
	ApiChart.prototype["SetSize"]                     =  ApiChart.prototype.SetSize      = ApiDrawing.prototype.SetSize;
	ApiChart.prototype["SetPosition"]                 =  ApiChart.prototype.SetPosition  = ApiDrawing.prototype.SetPosition;
	ApiChart.prototype["GetWidth"]                    =  ApiChart.prototype.GetWidth     = ApiDrawing.prototype.GetWidth;
	ApiChart.prototype["GetHeight"]                   =  ApiChart.prototype.GetHeight    = ApiDrawing.prototype.GetHeight;
	ApiChart.prototype["GetLockValue"]                =  ApiChart.prototype.GetLockValue = ApiDrawing.prototype.GetLockValue;
	ApiChart.prototype["SetLockValue"]                =  ApiChart.prototype.SetLockValue = ApiDrawing.prototype.SetLockValue;

	ApiOleObject.prototype["GetClassType"]            = ApiOleObject.prototype.GetClassType;
	ApiOleObject.prototype["SetData"]                 = ApiOleObject.prototype.SetData;
	ApiOleObject.prototype["GetData"]                 = ApiOleObject.prototype.GetData;
	ApiOleObject.prototype["SetApplicationId"]        = ApiOleObject.prototype.SetApplicationId;
	ApiOleObject.prototype["GetApplicationId"]        = ApiOleObject.prototype.GetApplicationId;

	ApiColor.prototype["GetClassType"]                =  ApiColor.prototype.GetClassType;
	ApiColor.prototype["GetRGB"]                      =  ApiColor.prototype.GetRGB;


	ApiName.prototype["GetName"]                 =  ApiName.prototype.GetName;
	ApiName.prototype["SetName"]                 =  ApiName.prototype.SetName;
	ApiName.prototype["Delete"]                  =  ApiName.prototype.Delete;
	ApiName.prototype["GetRefersTo"]             =  ApiName.prototype.GetRefersTo;
	ApiName.prototype["SetRefersTo"]             =  ApiName.prototype.SetRefersTo;
	ApiName.prototype["GetRefersToRange"]        =  ApiName.prototype.GetRefersToRange;


	ApiComment.prototype["GetClassType"]         =  ApiComment.prototype.GetClassType;
	ApiComment.prototype["GetText"]              =  ApiComment.prototype.GetText;
	ApiComment.prototype["SetText"]              =  ApiComment.prototype.SetText;
	ApiComment.prototype["GetId"]                =  ApiComment.prototype.GetId;
	ApiComment.prototype["GetAuthorName"]        =  ApiComment.prototype.GetAuthorName;
	ApiComment.prototype["SetAuthorName"]        =  ApiComment.prototype.SetAuthorName;
	ApiComment.prototype["GetUserId"]            =  ApiComment.prototype.GetUserId;
	ApiComment.prototype["SetUserId"]            =  ApiComment.prototype.SetUserId;
	ApiComment.prototype["IsSolved"]             =  ApiComment.prototype.IsSolved;
	ApiComment.prototype["SetSolved"]            =  ApiComment.prototype.SetSolved;
	ApiComment.prototype["GetTimeUTC"]           =  ApiComment.prototype.GetTimeUTC;
	ApiComment.prototype["SetTimeUTC"]           =  ApiComment.prototype.SetTimeUTC;
	ApiComment.prototype["GetTime"]              =  ApiComment.prototype.GetTime;
	ApiComment.prototype["SetTime"]              =  ApiComment.prototype.SetTime;
	ApiComment.prototype["GetQuoteText"]         =  ApiComment.prototype.GetQuoteText;
	ApiComment.prototype["GetRepliesCount"]      =  ApiComment.prototype.GetRepliesCount;
	ApiComment.prototype["GetReply"]             =  ApiComment.prototype.GetReply;
	ApiComment.prototype["AddReply"]             =  ApiComment.prototype.AddReply;
	ApiComment.prototype["RemoveReplies"]        =  ApiComment.prototype.RemoveReplies;
	ApiComment.prototype["Delete"]               =  ApiComment.prototype.Delete;


	ApiCommentReply.prototype["GetClassType"]         =  ApiCommentReply.prototype.GetClassType;
	ApiCommentReply.prototype["GetText"]              =  ApiCommentReply.prototype.GetText;
	ApiCommentReply.prototype["SetText"]              =  ApiCommentReply.prototype.SetText;
	ApiCommentReply.prototype["SetTextGetAuthorName"] =  ApiCommentReply.prototype.SetTextGetAuthorName;
	ApiCommentReply.prototype["GetAuthorName"]        =  ApiCommentReply.prototype.GetAuthorName;
	ApiCommentReply.prototype["SetAuthorName"]        =  ApiCommentReply.prototype.SetAuthorName;
	ApiCommentReply.prototype["GetUserId"]            =  ApiCommentReply.prototype.GetUserId;
	ApiCommentReply.prototype["SetUserId"]            =  ApiCommentReply.prototype.SetUserId;
	ApiCommentReply.prototype["GetTimeUTC"]           =  ApiCommentReply.prototype.GetTimeUTC;
	ApiCommentReply.prototype["SetTimeUTC"]           =  ApiCommentReply.prototype.SetTimeUTC;
	ApiCommentReply.prototype["GetTime"]              =  ApiCommentReply.prototype.GetTime;
	ApiCommentReply.prototype["SetTime"]              =  ApiCommentReply.prototype.SetTime;


	ApiAreas.prototype["GetCount"]               = ApiAreas.prototype.GetCount;
	ApiAreas.prototype["GetItem"]                = ApiAreas.prototype.GetItem;
	ApiAreas.prototype["GetParent"]              = ApiAreas.prototype.GetParent;


	ApiCharacters.prototype["GetCount"]          = ApiCharacters.prototype.GetCount;
	ApiCharacters.prototype["GetParent"]         = ApiCharacters.prototype.GetParent;
	ApiCharacters.prototype["Delete"]            = ApiCharacters.prototype.Delete;
	ApiCharacters.prototype["Insert"]            = ApiCharacters.prototype.Insert;
	ApiCharacters.prototype["SetCaption"]        = ApiCharacters.prototype.SetCaption;
	ApiCharacters.prototype["GetCaption"]        = ApiCharacters.prototype.GetCaption;
	ApiCharacters.prototype["SetText"]           = ApiCharacters.prototype.SetText;
	ApiCharacters.prototype["GetText"]           = ApiCharacters.prototype.GetText;
	ApiCharacters.prototype["GetFont"]           = ApiCharacters.prototype.GetFont;

	
	ApiFont.prototype["GetParent"]               = ApiFont.prototype.GetParent;
	ApiFont.prototype["GetBold"]                 = ApiFont.prototype.GetBold;
	ApiFont.prototype["SetBold"]                 = ApiFont.prototype.SetBold;
	ApiFont.prototype["GetItalic"]               = ApiFont.prototype.GetItalic;
	ApiFont.prototype["SetItalic"]               = ApiFont.prototype.SetItalic;
	ApiFont.prototype["GetSize"]                 = ApiFont.prototype.GetSize;
	ApiFont.prototype["SetSize"]                 = ApiFont.prototype.SetSize;
	ApiFont.prototype["GetStrikethrough"]        = ApiFont.prototype.GetStrikethrough;
	ApiFont.prototype["SetStrikethrough"]        = ApiFont.prototype.SetStrikethrough;
	ApiFont.prototype["GetUnderline"]            = ApiFont.prototype.GetUnderline;
	ApiFont.prototype["SetUnderline"]            = ApiFont.prototype.SetUnderline;
	ApiFont.prototype["GetSubscript"]            = ApiFont.prototype.GetSubscript;
	ApiFont.prototype["SetSubscript"]            = ApiFont.prototype.SetSubscript;
	ApiFont.prototype["GetSuperscript"]          = ApiFont.prototype.GetSuperscript;
	ApiFont.prototype["SetSuperscript"]          = ApiFont.prototype.SetSuperscript;
	ApiFont.prototype["GetName"]                 = ApiFont.prototype.GetName;
	ApiFont.prototype["SetName"]                 = ApiFont.prototype.SetName;
	ApiFont.prototype["GetColor"]                = ApiFont.prototype.GetColor;
	ApiFont.prototype["SetColor"]                = ApiFont.prototype.SetColor;

	ApiFreezePanes.prototype["FreezeAt"]         = ApiFreezePanes.prototype.FreezeAt;
	ApiFreezePanes.prototype["FreezeColumns"]    = ApiFreezePanes.prototype.FreezeColumns;
	ApiFreezePanes.prototype["FreezeRows"]       = ApiFreezePanes.prototype.FreezeRows;
	ApiFreezePanes.prototype["GetLocation"]      = ApiFreezePanes.prototype.GetLocation;
	ApiFreezePanes.prototype["Unfreeze"]         = ApiFreezePanes.prototype.Unfreeze;

	ApiProtectedRange.prototype["SetTitle"]      = ApiProtectedRange.prototype.SetTitle;
	ApiProtectedRange.prototype["SetRange"]      = ApiProtectedRange.prototype.SetRange;
	ApiProtectedRange.prototype["AddUser"]       = ApiProtectedRange.prototype.AddUser;
	ApiProtectedRange.prototype["DeleteUser"]    = ApiProtectedRange.prototype.DeleteUser;
	ApiProtectedRange.prototype["GetAllUsers"]   = ApiProtectedRange.prototype.GetAllUsers;
	ApiProtectedRange.prototype["SetAnyoneType"] = ApiProtectedRange.prototype.SetAnyoneType;
	ApiProtectedRange.prototype["GetUser"]       = ApiProtectedRange.prototype.GetUser;

	ApiProtectedRangeUserInfo.prototype["GetName"]  = ApiProtectedRangeUserInfo.prototype.GetName;
	ApiProtectedRangeUserInfo.prototype["GetType"]  = ApiProtectedRangeUserInfo.prototype.GetType;
	ApiProtectedRangeUserInfo.prototype["GetId"]    = ApiProtectedRangeUserInfo.prototype.GetId;

	ApiWorksheetFunction.prototype["ASC"]             =  ApiWorksheetFunction.prototype.ASC;
	ApiWorksheetFunction.prototype["CHAR"]            =  ApiWorksheetFunction.prototype.CHAR;
	ApiWorksheetFunction.prototype["CLEAN"]           =  ApiWorksheetFunction.prototype.CLEAN;
	ApiWorksheetFunction.prototype["CODE"]            =  ApiWorksheetFunction.prototype.CODE;
	ApiWorksheetFunction.prototype["CONCATENATE"]     =  ApiWorksheetFunction.prototype.CONCATENATE;
	ApiWorksheetFunction.prototype["DOLLAR"]          =  ApiWorksheetFunction.prototype.DOLLAR;
	ApiWorksheetFunction.prototype["EXACT"]           =  ApiWorksheetFunction.prototype.EXACT;
	ApiWorksheetFunction.prototype["FIND"]            =  ApiWorksheetFunction.prototype.FIND;
	ApiWorksheetFunction.prototype["FINDB"]           =  ApiWorksheetFunction.prototype.FINDB;
	ApiWorksheetFunction.prototype["FIXED"]           =  ApiWorksheetFunction.prototype.FIXED;
	ApiWorksheetFunction.prototype["LEFT"]            =  ApiWorksheetFunction.prototype.LEFT;
	ApiWorksheetFunction.prototype["LEFTB"]           =  ApiWorksheetFunction.prototype.LEFTB;
	ApiWorksheetFunction.prototype["LEN"]             =  ApiWorksheetFunction.prototype.LEN;
	ApiWorksheetFunction.prototype["LENB"]            =  ApiWorksheetFunction.prototype.LENB;
	ApiWorksheetFunction.prototype["LOWER"]           =  ApiWorksheetFunction.prototype.LOWER;
	ApiWorksheetFunction.prototype["MID"]             =  ApiWorksheetFunction.prototype.MID;
	ApiWorksheetFunction.prototype["MIDB"]            =  ApiWorksheetFunction.prototype.MIDB;
	ApiWorksheetFunction.prototype["NUMBERVALUE"]     =  ApiWorksheetFunction.prototype.NUMBERVALUE;
	ApiWorksheetFunction.prototype["PROPER"]          =  ApiWorksheetFunction.prototype.PROPER;
	ApiWorksheetFunction.prototype["REPLACE"]         =  ApiWorksheetFunction.prototype.REPLACE;
	ApiWorksheetFunction.prototype["REPLACEB"]        =  ApiWorksheetFunction.prototype.REPLACEB;
	ApiWorksheetFunction.prototype["REPT"]            =  ApiWorksheetFunction.prototype.REPT;
	ApiWorksheetFunction.prototype["RIGHT"]           =  ApiWorksheetFunction.prototype.RIGHT;
	ApiWorksheetFunction.prototype["RIGHTB"]          =  ApiWorksheetFunction.prototype.RIGHTB;
	ApiWorksheetFunction.prototype["SEARCH"]          =  ApiWorksheetFunction.prototype.SEARCH;
	ApiWorksheetFunction.prototype["SEARCHB"]         =  ApiWorksheetFunction.prototype.SEARCHB;
	ApiWorksheetFunction.prototype["SUBSTITUTE"]      =  ApiWorksheetFunction.prototype.SUBSTITUTE;
	ApiWorksheetFunction.prototype["T"]               =  ApiWorksheetFunction.prototype.T;
	ApiWorksheetFunction.prototype["TEXT"]            =  ApiWorksheetFunction.prototype.TEXT;
	ApiWorksheetFunction.prototype["TRIM"]            =  ApiWorksheetFunction.prototype.TRIM;
	ApiWorksheetFunction.prototype["UNICHAR"]         =  ApiWorksheetFunction.prototype.UNICHAR;
	ApiWorksheetFunction.prototype["UNICODE"]         =  ApiWorksheetFunction.prototype.UNICODE;
	ApiWorksheetFunction.prototype["UPPER"]           =  ApiWorksheetFunction.prototype.UPPER;
	ApiWorksheetFunction.prototype["VALUE"]           =  ApiWorksheetFunction.prototype.VALUE;
	ApiWorksheetFunction.prototype["AVEDEV"]          =  ApiWorksheetFunction.prototype.AVEDEV;
	ApiWorksheetFunction.prototype["AVERAGE"]         =  ApiWorksheetFunction.prototype.AVERAGE;
	ApiWorksheetFunction.prototype["AVERAGEA"]        =  ApiWorksheetFunction.prototype.AVERAGEA;
	ApiWorksheetFunction.prototype["AVERAGEIF"]       =  ApiWorksheetFunction.prototype.AVERAGEIF;
	ApiWorksheetFunction.prototype["AVERAGEIFS"]      =  ApiWorksheetFunction.prototype.AVERAGEIFS;
	ApiWorksheetFunction.prototype["BETADIST"]        =  ApiWorksheetFunction.prototype.BETADIST;
	ApiWorksheetFunction.prototype["BETAINV"]         =  ApiWorksheetFunction.prototype.BETAINV;
	ApiWorksheetFunction.prototype["BINOMDIST"]       =  ApiWorksheetFunction.prototype.BINOMDIST;
	ApiWorksheetFunction.prototype["CHIDIST"]         =  ApiWorksheetFunction.prototype.CHIDIST;
	ApiWorksheetFunction.prototype["CHIINV"]          =  ApiWorksheetFunction.prototype.CHIINV;
	ApiWorksheetFunction.prototype["CONFIDENCE"]      =  ApiWorksheetFunction.prototype.CONFIDENCE;
	ApiWorksheetFunction.prototype["CHITEST"]         =  ApiWorksheetFunction.prototype.CHITEST;
	ApiWorksheetFunction.prototype["COUNT"]           =  ApiWorksheetFunction.prototype.COUNT;
	ApiWorksheetFunction.prototype["COUNTA"]          =  ApiWorksheetFunction.prototype.COUNTA;
	ApiWorksheetFunction.prototype["COUNTBLANK"]      =  ApiWorksheetFunction.prototype.COUNTBLANK;
	ApiWorksheetFunction.prototype["COUNTIF"]         =  ApiWorksheetFunction.prototype.COUNTIF;
	ApiWorksheetFunction.prototype["COUNTIFS"]        =  ApiWorksheetFunction.prototype.COUNTIFS;
	ApiWorksheetFunction.prototype["CRITBINOM"]       =  ApiWorksheetFunction.prototype.CRITBINOM;
	ApiWorksheetFunction.prototype["DEVSQ"]           =  ApiWorksheetFunction.prototype.DEVSQ;
	ApiWorksheetFunction.prototype["EXPONDIST"]       =  ApiWorksheetFunction.prototype.EXPONDIST;
	ApiWorksheetFunction.prototype["FDIST"]           =  ApiWorksheetFunction.prototype.FDIST;
	ApiWorksheetFunction.prototype["FINV"]            =  ApiWorksheetFunction.prototype.FINV;
	ApiWorksheetFunction.prototype["FISHER"]          =  ApiWorksheetFunction.prototype.FISHER;
	ApiWorksheetFunction.prototype["FISHERINV"]       =  ApiWorksheetFunction.prototype.FISHERINV;
	ApiWorksheetFunction.prototype["FREQUENCY"]       =  ApiWorksheetFunction.prototype.FREQUENCY;
	ApiWorksheetFunction.prototype["GAMMA"]           =  ApiWorksheetFunction.prototype.GAMMA;
	ApiWorksheetFunction.prototype["GAMMADIST"]       =  ApiWorksheetFunction.prototype.GAMMADIST;
	ApiWorksheetFunction.prototype["GAMMAINV"]        =  ApiWorksheetFunction.prototype.GAMMAINV;
	ApiWorksheetFunction.prototype["GAMMALN"]         =  ApiWorksheetFunction.prototype.GAMMALN;
	ApiWorksheetFunction.prototype["GAUSS"]           =  ApiWorksheetFunction.prototype.GAUSS;
	ApiWorksheetFunction.prototype["GEOMEAN"]         =  ApiWorksheetFunction.prototype.GEOMEAN;
	ApiWorksheetFunction.prototype["GROWTH"]          =  ApiWorksheetFunction.prototype.GROWTH;
	ApiWorksheetFunction.prototype["HARMEAN"]         =  ApiWorksheetFunction.prototype.HARMEAN;
	ApiWorksheetFunction.prototype["HYPGEOMDIST"]     =  ApiWorksheetFunction.prototype.HYPGEOMDIST;
	ApiWorksheetFunction.prototype["KURT"]            =  ApiWorksheetFunction.prototype.KURT;
	ApiWorksheetFunction.prototype["LARGE"]           =  ApiWorksheetFunction.prototype.LARGE;
	ApiWorksheetFunction.prototype["LINEST"]          =  ApiWorksheetFunction.prototype.LINEST;
	ApiWorksheetFunction.prototype["LOGEST"]          =  ApiWorksheetFunction.prototype.LOGEST;
	ApiWorksheetFunction.prototype["LOGINV"]          =  ApiWorksheetFunction.prototype.LOGINV;
	ApiWorksheetFunction.prototype["LOGNORMDIST"]     =  ApiWorksheetFunction.prototype.LOGNORMDIST;
	ApiWorksheetFunction.prototype["MAX"]             =  ApiWorksheetFunction.prototype.MAX;
	ApiWorksheetFunction.prototype["MAXA"]            =  ApiWorksheetFunction.prototype.MAXA;
	ApiWorksheetFunction.prototype["MEDIAN"]          =  ApiWorksheetFunction.prototype.MEDIAN;
	ApiWorksheetFunction.prototype["MIN"]             =  ApiWorksheetFunction.prototype.MIN;
	ApiWorksheetFunction.prototype["MINA"]            =  ApiWorksheetFunction.prototype.MINA;
	ApiWorksheetFunction.prototype["NEGBINOMDIST"]    =  ApiWorksheetFunction.prototype.NEGBINOMDIST;
	ApiWorksheetFunction.prototype["NORMDIST"]        =  ApiWorksheetFunction.prototype.NORMDIST;
	ApiWorksheetFunction.prototype["NORMINV"]         =  ApiWorksheetFunction.prototype.NORMINV;
	ApiWorksheetFunction.prototype["NORMSDIST"]       =  ApiWorksheetFunction.prototype.NORMSDIST;
	ApiWorksheetFunction.prototype["NORMSINV"]        =  ApiWorksheetFunction.prototype.NORMSINV;
	ApiWorksheetFunction.prototype["PERCENTILE"]      =  ApiWorksheetFunction.prototype.PERCENTILE;
	ApiWorksheetFunction.prototype["PERCENTRANK"]     =  ApiWorksheetFunction.prototype.PERCENTRANK;
	ApiWorksheetFunction.prototype["PERMUT"]          =  ApiWorksheetFunction.prototype.PERMUT;
	ApiWorksheetFunction.prototype["PERMUTATIONA"]    =  ApiWorksheetFunction.prototype.PERMUTATIONA;
	ApiWorksheetFunction.prototype["PHI"]             =  ApiWorksheetFunction.prototype.PHI;
	ApiWorksheetFunction.prototype["POISSON"]         =  ApiWorksheetFunction.prototype.POISSON;
	ApiWorksheetFunction.prototype["QUARTILE"]        =  ApiWorksheetFunction.prototype.QUARTILE;
	ApiWorksheetFunction.prototype["RANK"]            =  ApiWorksheetFunction.prototype.RANK;
	ApiWorksheetFunction.prototype["SKEW"]            =  ApiWorksheetFunction.prototype.SKEW;
	ApiWorksheetFunction.prototype["SMALL"]           =  ApiWorksheetFunction.prototype.SMALL;
	ApiWorksheetFunction.prototype["STANDARDIZE"]     =  ApiWorksheetFunction.prototype.STANDARDIZE;
	ApiWorksheetFunction.prototype["STDEV"]           =  ApiWorksheetFunction.prototype.STDEV;
	ApiWorksheetFunction.prototype["STDEVA"]          =  ApiWorksheetFunction.prototype.STDEVA;
	ApiWorksheetFunction.prototype["STDEVP"]          =  ApiWorksheetFunction.prototype.STDEVP;
	ApiWorksheetFunction.prototype["STDEVPA"]         =  ApiWorksheetFunction.prototype.STDEVPA;
	ApiWorksheetFunction.prototype["TDIST"]           =  ApiWorksheetFunction.prototype.TDIST;
	ApiWorksheetFunction.prototype["TINV"]            =  ApiWorksheetFunction.prototype.TINV;
	ApiWorksheetFunction.prototype["TREND"]           =  ApiWorksheetFunction.prototype.TREND;
	ApiWorksheetFunction.prototype["TRIMMEAN"]        =  ApiWorksheetFunction.prototype.TRIMMEAN;
	ApiWorksheetFunction.prototype["VAR"]             =  ApiWorksheetFunction.prototype.VAR;
	ApiWorksheetFunction.prototype["VARA"]            =  ApiWorksheetFunction.prototype.VARA;
	ApiWorksheetFunction.prototype["VARP"]            =  ApiWorksheetFunction.prototype.VARP;
	ApiWorksheetFunction.prototype["VARPA"]           =  ApiWorksheetFunction.prototype.VARPA;
	ApiWorksheetFunction.prototype["WEIBULL"]         =  ApiWorksheetFunction.prototype.WEIBULL;
	ApiWorksheetFunction.prototype["ZTEST"]           =  ApiWorksheetFunction.prototype.ZTEST;
	ApiWorksheetFunction.prototype["DATE"]            =  ApiWorksheetFunction.prototype.DATE;
	ApiWorksheetFunction.prototype["DATEVALUE"]       =  ApiWorksheetFunction.prototype.DATEVALUE;
	ApiWorksheetFunction.prototype["DAY"]             =  ApiWorksheetFunction.prototype.DAY;
	ApiWorksheetFunction.prototype["DAYS"]            =  ApiWorksheetFunction.prototype.DAYS;
	ApiWorksheetFunction.prototype["DAYS360"]         =  ApiWorksheetFunction.prototype.DAYS360;
	ApiWorksheetFunction.prototype["EDATE"]           =  ApiWorksheetFunction.prototype.EDATE;
	ApiWorksheetFunction.prototype["EOMONTH"]         =  ApiWorksheetFunction.prototype.EOMONTH;
	ApiWorksheetFunction.prototype["HOUR"]            =  ApiWorksheetFunction.prototype.HOUR;
	ApiWorksheetFunction.prototype["ISOWEEKNUM"]      =  ApiWorksheetFunction.prototype.ISOWEEKNUM;
	ApiWorksheetFunction.prototype["MINUTE"]          =  ApiWorksheetFunction.prototype.MINUTE;
	ApiWorksheetFunction.prototype["MONTH"]           =  ApiWorksheetFunction.prototype.MONTH;
	ApiWorksheetFunction.prototype["NETWORKDAYS"]     =  ApiWorksheetFunction.prototype.NETWORKDAYS;
	ApiWorksheetFunction.prototype["NOW"]             =  ApiWorksheetFunction.prototype.NOW;
	ApiWorksheetFunction.prototype["SECOND"]          =  ApiWorksheetFunction.prototype.SECOND;
	ApiWorksheetFunction.prototype["TIME"]            =  ApiWorksheetFunction.prototype.TIME;
	ApiWorksheetFunction.prototype["TIMEVALUE"]       =  ApiWorksheetFunction.prototype.TIMEVALUE;
	ApiWorksheetFunction.prototype["TODAY"]           =  ApiWorksheetFunction.prototype.TODAY;
	ApiWorksheetFunction.prototype["WEEKDAY"]         =  ApiWorksheetFunction.prototype.WEEKDAY;
	ApiWorksheetFunction.prototype["WEEKNUM"]         =  ApiWorksheetFunction.prototype.WEEKNUM;
	ApiWorksheetFunction.prototype["WORKDAY"]         =  ApiWorksheetFunction.prototype.WORKDAY;
	ApiWorksheetFunction.prototype["YEAR"]            =  ApiWorksheetFunction.prototype.YEAR;
	ApiWorksheetFunction.prototype["YEARFRAC"]        =  ApiWorksheetFunction.prototype.YEARFRAC;
	ApiWorksheetFunction.prototype["BESSELI"]         =  ApiWorksheetFunction.prototype.BESSELI;
	ApiWorksheetFunction.prototype["BESSELJ"]         =  ApiWorksheetFunction.prototype.BESSELJ;
	ApiWorksheetFunction.prototype["BESSELK"]         =  ApiWorksheetFunction.prototype.BESSELK;
	ApiWorksheetFunction.prototype["BESSELY"]         =  ApiWorksheetFunction.prototype.BESSELY;
	ApiWorksheetFunction.prototype["BIN2DEC"]         =  ApiWorksheetFunction.prototype.BIN2DEC;
	ApiWorksheetFunction.prototype["BIN2HEX"]         =  ApiWorksheetFunction.prototype.BIN2HEX;
	ApiWorksheetFunction.prototype["BIN2OCT"]         =  ApiWorksheetFunction.prototype.BIN2OCT;
	ApiWorksheetFunction.prototype["BITAND"]          =  ApiWorksheetFunction.prototype.BITAND;
	ApiWorksheetFunction.prototype["BITLSHIFT"]       =  ApiWorksheetFunction.prototype.BITLSHIFT;
	ApiWorksheetFunction.prototype["BITOR"]           =  ApiWorksheetFunction.prototype.BITOR;
	ApiWorksheetFunction.prototype["BITRSHIFT"]       =  ApiWorksheetFunction.prototype.BITRSHIFT;
	ApiWorksheetFunction.prototype["BITXOR"]          =  ApiWorksheetFunction.prototype.BITXOR;
	ApiWorksheetFunction.prototype["COMPLEX"]         =  ApiWorksheetFunction.prototype.COMPLEX;
	ApiWorksheetFunction.prototype["CONVERT"]         =  ApiWorksheetFunction.prototype.CONVERT;
	ApiWorksheetFunction.prototype["DEC2BIN"]         =  ApiWorksheetFunction.prototype.DEC2BIN;
	ApiWorksheetFunction.prototype["DEC2HEX"]         =  ApiWorksheetFunction.prototype.DEC2HEX;
	ApiWorksheetFunction.prototype["DEC2OCT"]         =  ApiWorksheetFunction.prototype.DEC2OCT;
	ApiWorksheetFunction.prototype["DELTA"]           =  ApiWorksheetFunction.prototype.DELTA;
	ApiWorksheetFunction.prototype["ERF"]             =  ApiWorksheetFunction.prototype.ERF;
	ApiWorksheetFunction.prototype["ERFC"]            =  ApiWorksheetFunction.prototype.ERFC;
	ApiWorksheetFunction.prototype["GESTEP"]          =  ApiWorksheetFunction.prototype.GESTEP;
	ApiWorksheetFunction.prototype["HEX2BIN"]         =  ApiWorksheetFunction.prototype.HEX2BIN;
	ApiWorksheetFunction.prototype["HEX2DEC"]         =  ApiWorksheetFunction.prototype.HEX2DEC;
	ApiWorksheetFunction.prototype["HEX2OCT"]         =  ApiWorksheetFunction.prototype.HEX2OCT;
	ApiWorksheetFunction.prototype["IMABS"]           =  ApiWorksheetFunction.prototype.IMABS;
	ApiWorksheetFunction.prototype["IMAGINARY"]       =  ApiWorksheetFunction.prototype.IMAGINARY;
	ApiWorksheetFunction.prototype["IMARGUMENT"]      =  ApiWorksheetFunction.prototype.IMARGUMENT;
	ApiWorksheetFunction.prototype["IMCONJUGATE"]     =  ApiWorksheetFunction.prototype.IMCONJUGATE;
	ApiWorksheetFunction.prototype["IMCOS"]           =  ApiWorksheetFunction.prototype.IMCOS;
	ApiWorksheetFunction.prototype["IMCOSH"]          =  ApiWorksheetFunction.prototype.IMCOSH;
	ApiWorksheetFunction.prototype["IMCOT"]           =  ApiWorksheetFunction.prototype.IMCOT;
	ApiWorksheetFunction.prototype["IMCSC"]           =  ApiWorksheetFunction.prototype.IMCSC;
	ApiWorksheetFunction.prototype["IMCSCH"]          =  ApiWorksheetFunction.prototype.IMCSCH;
	ApiWorksheetFunction.prototype["IMDIV"]           =  ApiWorksheetFunction.prototype.IMDIV;
	ApiWorksheetFunction.prototype["IMEXP"]           =  ApiWorksheetFunction.prototype.IMEXP;
	ApiWorksheetFunction.prototype["IMLN"]            =  ApiWorksheetFunction.prototype.IMLN;
	ApiWorksheetFunction.prototype["IMLOG10"]         =  ApiWorksheetFunction.prototype.IMLOG10;
	ApiWorksheetFunction.prototype["IMLOG2"]          =  ApiWorksheetFunction.prototype.IMLOG2;
	ApiWorksheetFunction.prototype["IMPOWER"]         =  ApiWorksheetFunction.prototype.IMPOWER;
	ApiWorksheetFunction.prototype["IMPRODUCT"]       =  ApiWorksheetFunction.prototype.IMPRODUCT;
	ApiWorksheetFunction.prototype["IMREAL"]          =  ApiWorksheetFunction.prototype.IMREAL;
	ApiWorksheetFunction.prototype["IMSEC"]           =  ApiWorksheetFunction.prototype.IMSEC;
	ApiWorksheetFunction.prototype["IMSECH"]          =  ApiWorksheetFunction.prototype.IMSECH;
	ApiWorksheetFunction.prototype["IMSIN"]           =  ApiWorksheetFunction.prototype.IMSIN;
	ApiWorksheetFunction.prototype["IMSINH"]          =  ApiWorksheetFunction.prototype.IMSINH;
	ApiWorksheetFunction.prototype["IMSQRT"]          =  ApiWorksheetFunction.prototype.IMSQRT;
	ApiWorksheetFunction.prototype["IMSUB"]           =  ApiWorksheetFunction.prototype.IMSUB;
	ApiWorksheetFunction.prototype["IMSUM"]           =  ApiWorksheetFunction.prototype.IMSUM;
	ApiWorksheetFunction.prototype["IMTAN"]           =  ApiWorksheetFunction.prototype.IMTAN;
	ApiWorksheetFunction.prototype["OCT2BIN"]         =  ApiWorksheetFunction.prototype.OCT2BIN;
	ApiWorksheetFunction.prototype["OCT2DEC"]         =  ApiWorksheetFunction.prototype.OCT2DEC;
	ApiWorksheetFunction.prototype["OCT2HEX"]         =  ApiWorksheetFunction.prototype.OCT2HEX;
	ApiWorksheetFunction.prototype["DAVERAGE"]        =  ApiWorksheetFunction.prototype.DAVERAGE;
	ApiWorksheetFunction.prototype["DCOUNT"]          =  ApiWorksheetFunction.prototype.DCOUNT;
	ApiWorksheetFunction.prototype["DCOUNTA"]         =  ApiWorksheetFunction.prototype.DCOUNTA;
	ApiWorksheetFunction.prototype["DGET"]            =  ApiWorksheetFunction.prototype.DGET;
	ApiWorksheetFunction.prototype["DMAX"]            =  ApiWorksheetFunction.prototype.DMAX;
	ApiWorksheetFunction.prototype["DMIN"]            =  ApiWorksheetFunction.prototype.DMIN;
	ApiWorksheetFunction.prototype["DPRODUCT"]        =  ApiWorksheetFunction.prototype.DPRODUCT;
	ApiWorksheetFunction.prototype["DSTDEV"]          =  ApiWorksheetFunction.prototype.DSTDEV;
	ApiWorksheetFunction.prototype["DSTDEVP"]         =  ApiWorksheetFunction.prototype.DSTDEVP;
	ApiWorksheetFunction.prototype["DSUM"]            =  ApiWorksheetFunction.prototype.DSUM;
	ApiWorksheetFunction.prototype["DVAR"]            =  ApiWorksheetFunction.prototype.DVAR;
	ApiWorksheetFunction.prototype["DVARP"]           =  ApiWorksheetFunction.prototype.DVARP;
	ApiWorksheetFunction.prototype["ACCRINT"]         =  ApiWorksheetFunction.prototype.ACCRINT;
	ApiWorksheetFunction.prototype["ACCRINTM"]        =  ApiWorksheetFunction.prototype.ACCRINTM;
	ApiWorksheetFunction.prototype["AMORDEGRC"]       =  ApiWorksheetFunction.prototype.AMORDEGRC;
	ApiWorksheetFunction.prototype["AMORLINC"]        =  ApiWorksheetFunction.prototype.AMORLINC;
	ApiWorksheetFunction.prototype["COUPDAYBS"]       =  ApiWorksheetFunction.prototype.COUPDAYBS;
	ApiWorksheetFunction.prototype["COUPDAYS"]        =  ApiWorksheetFunction.prototype.COUPDAYS;
	ApiWorksheetFunction.prototype["COUPDAYSNC"]      =  ApiWorksheetFunction.prototype.COUPDAYSNC;
	ApiWorksheetFunction.prototype["COUPNCD"]         =  ApiWorksheetFunction.prototype.COUPNCD;
	ApiWorksheetFunction.prototype["COUPNUM"]         =  ApiWorksheetFunction.prototype.COUPNUM;
	ApiWorksheetFunction.prototype["COUPPCD"]         =  ApiWorksheetFunction.prototype.COUPPCD;
	ApiWorksheetFunction.prototype["CUMIPMT"]         =  ApiWorksheetFunction.prototype.CUMIPMT;
	ApiWorksheetFunction.prototype["CUMPRINC"]        =  ApiWorksheetFunction.prototype.CUMPRINC;
	ApiWorksheetFunction.prototype["DB"]              =  ApiWorksheetFunction.prototype.DB;
	ApiWorksheetFunction.prototype["DDB"]             =  ApiWorksheetFunction.prototype.DDB;
	ApiWorksheetFunction.prototype["DISC"]            =  ApiWorksheetFunction.prototype.DISC;
	ApiWorksheetFunction.prototype["DOLLARDE"]        =  ApiWorksheetFunction.prototype.DOLLARDE;
	ApiWorksheetFunction.prototype["DOLLARFR"]        =  ApiWorksheetFunction.prototype.DOLLARFR;
	ApiWorksheetFunction.prototype["DURATION"]        =  ApiWorksheetFunction.prototype.DURATION;
	ApiWorksheetFunction.prototype["EFFECT"]          =  ApiWorksheetFunction.prototype.EFFECT;
	ApiWorksheetFunction.prototype["FV"]              =  ApiWorksheetFunction.prototype.FV;
	ApiWorksheetFunction.prototype["FVSCHEDULE"]      =  ApiWorksheetFunction.prototype.FVSCHEDULE;
	ApiWorksheetFunction.prototype["INTRATE"]         =  ApiWorksheetFunction.prototype.INTRATE;
	ApiWorksheetFunction.prototype["IPMT"]            =  ApiWorksheetFunction.prototype.IPMT;
	ApiWorksheetFunction.prototype["IRR"]             =  ApiWorksheetFunction.prototype.IRR;
	ApiWorksheetFunction.prototype["ISPMT"]           =  ApiWorksheetFunction.prototype.ISPMT;
	ApiWorksheetFunction.prototype["MDURATION"]       =  ApiWorksheetFunction.prototype.MDURATION;
	ApiWorksheetFunction.prototype["MIRR"]            =  ApiWorksheetFunction.prototype.MIRR;
	ApiWorksheetFunction.prototype["NOMINAL"]         =  ApiWorksheetFunction.prototype.NOMINAL;
	ApiWorksheetFunction.prototype["NPER"]            =  ApiWorksheetFunction.prototype.NPER;
	ApiWorksheetFunction.prototype["NPV"]             =  ApiWorksheetFunction.prototype.NPV;
	ApiWorksheetFunction.prototype["ODDFPRICE"]       =  ApiWorksheetFunction.prototype.ODDFPRICE;
	ApiWorksheetFunction.prototype["ODDFYIELD"]       =  ApiWorksheetFunction.prototype.ODDFYIELD;
	ApiWorksheetFunction.prototype["ODDLPRICE"]       =  ApiWorksheetFunction.prototype.ODDLPRICE;
	ApiWorksheetFunction.prototype["ODDLYIELD"]       =  ApiWorksheetFunction.prototype.ODDLYIELD;
	ApiWorksheetFunction.prototype["PDURATION"]       =  ApiWorksheetFunction.prototype.PDURATION;
	ApiWorksheetFunction.prototype["PMT"]             =  ApiWorksheetFunction.prototype.PMT;
	ApiWorksheetFunction.prototype["PPMT"]            =  ApiWorksheetFunction.prototype.PPMT;
	ApiWorksheetFunction.prototype["PRICE"]           =  ApiWorksheetFunction.prototype.PRICE;
	ApiWorksheetFunction.prototype["PRICEDISC"]       =  ApiWorksheetFunction.prototype.PRICEDISC;
	ApiWorksheetFunction.prototype["PRICEMAT"]        =  ApiWorksheetFunction.prototype.PRICEMAT;
	ApiWorksheetFunction.prototype["PV"]              =  ApiWorksheetFunction.prototype.PV;
	ApiWorksheetFunction.prototype["RATE"]            =  ApiWorksheetFunction.prototype.RATE;
	ApiWorksheetFunction.prototype["RECEIVED"]        =  ApiWorksheetFunction.prototype.RECEIVED;
	ApiWorksheetFunction.prototype["RRI"]             =  ApiWorksheetFunction.prototype.RRI;
	ApiWorksheetFunction.prototype["SLN"]             =  ApiWorksheetFunction.prototype.SLN;
	ApiWorksheetFunction.prototype["SYD"]             =  ApiWorksheetFunction.prototype.SYD;
	ApiWorksheetFunction.prototype["TBILLEQ"]         =  ApiWorksheetFunction.prototype.TBILLEQ;
	ApiWorksheetFunction.prototype["TBILLPRICE"]      =  ApiWorksheetFunction.prototype.TBILLPRICE;
	ApiWorksheetFunction.prototype["TBILLYIELD"]      =  ApiWorksheetFunction.prototype.TBILLYIELD;
	ApiWorksheetFunction.prototype["VDB"]             =  ApiWorksheetFunction.prototype.VDB;
	ApiWorksheetFunction.prototype["XIRR"]            =  ApiWorksheetFunction.prototype.XIRR;
	ApiWorksheetFunction.prototype["XNPV"]            =  ApiWorksheetFunction.prototype.XNPV;
	ApiWorksheetFunction.prototype["YIELD"]           =  ApiWorksheetFunction.prototype.YIELD;
	ApiWorksheetFunction.prototype["YIELDDISC"]       =  ApiWorksheetFunction.prototype.YIELDDISC;
	ApiWorksheetFunction.prototype["YIELDMAT"]        =  ApiWorksheetFunction.prototype.YIELDMAT;
	ApiWorksheetFunction.prototype["ABS"]             =  ApiWorksheetFunction.prototype.ABS;
	ApiWorksheetFunction.prototype["ACOS"]            =  ApiWorksheetFunction.prototype.ACOS;
	ApiWorksheetFunction.prototype["ACOSH"]           =  ApiWorksheetFunction.prototype.ACOSH;
	ApiWorksheetFunction.prototype["ACOT"]            =  ApiWorksheetFunction.prototype.ACOT;
	ApiWorksheetFunction.prototype["ACOTH"]           =  ApiWorksheetFunction.prototype.ACOTH;
	ApiWorksheetFunction.prototype["AGGREGATE"]       =  ApiWorksheetFunction.prototype.AGGREGATE;
	ApiWorksheetFunction.prototype["ARABIC"]          =  ApiWorksheetFunction.prototype.ARABIC;
	ApiWorksheetFunction.prototype["ASIN"]            =  ApiWorksheetFunction.prototype.ASIN;
	ApiWorksheetFunction.prototype["ASINH"]           =  ApiWorksheetFunction.prototype.ASINH;
	ApiWorksheetFunction.prototype["ATAN"]            =  ApiWorksheetFunction.prototype.ATAN;
	ApiWorksheetFunction.prototype["ATAN2"]           =  ApiWorksheetFunction.prototype.ATAN2;
	ApiWorksheetFunction.prototype["ATANH"]           =  ApiWorksheetFunction.prototype.ATANH;
	ApiWorksheetFunction.prototype["BASE"]            =  ApiWorksheetFunction.prototype.BASE;
	ApiWorksheetFunction.prototype["CEILING"]         =  ApiWorksheetFunction.prototype.CEILING;
	ApiWorksheetFunction.prototype["COMBIN"]          =  ApiWorksheetFunction.prototype.COMBIN;
	ApiWorksheetFunction.prototype["COMBINA"]         =  ApiWorksheetFunction.prototype.COMBINA;
	ApiWorksheetFunction.prototype["COS"]             =  ApiWorksheetFunction.prototype.COS;
	ApiWorksheetFunction.prototype["COSH"]            =  ApiWorksheetFunction.prototype.COSH;
	ApiWorksheetFunction.prototype["COT"]             =  ApiWorksheetFunction.prototype.COT;
	ApiWorksheetFunction.prototype["COTH"]            =  ApiWorksheetFunction.prototype.COTH;
	ApiWorksheetFunction.prototype["CSC"]             =  ApiWorksheetFunction.prototype.CSC;
	ApiWorksheetFunction.prototype["CSCH"]            =  ApiWorksheetFunction.prototype.CSCH;
	ApiWorksheetFunction.prototype["DECIMAL"]         =  ApiWorksheetFunction.prototype.DECIMAL;
	ApiWorksheetFunction.prototype["DEGREES"]         =  ApiWorksheetFunction.prototype.DEGREES;
	ApiWorksheetFunction.prototype["EVEN"]            =  ApiWorksheetFunction.prototype.EVEN;
	ApiWorksheetFunction.prototype["EXP"]             =  ApiWorksheetFunction.prototype.EXP;
	ApiWorksheetFunction.prototype["FACT"]            =  ApiWorksheetFunction.prototype.FACT;
	ApiWorksheetFunction.prototype["FACTDOUBLE"]      =  ApiWorksheetFunction.prototype.FACTDOUBLE;
	ApiWorksheetFunction.prototype["FLOOR"]           =  ApiWorksheetFunction.prototype.FLOOR;
	ApiWorksheetFunction.prototype["GCD"]             =  ApiWorksheetFunction.prototype.GCD;
	ApiWorksheetFunction.prototype["INT"]             =  ApiWorksheetFunction.prototype.INT;
	ApiWorksheetFunction.prototype["LCM"]             =  ApiWorksheetFunction.prototype.LCM;
	ApiWorksheetFunction.prototype["LN"]              =  ApiWorksheetFunction.prototype.LN;
	ApiWorksheetFunction.prototype["LOG"]             =  ApiWorksheetFunction.prototype.LOG;
	ApiWorksheetFunction.prototype["LOG10"]           =  ApiWorksheetFunction.prototype.LOG10;
	ApiWorksheetFunction.prototype["MOD"]             =  ApiWorksheetFunction.prototype.MOD;
	ApiWorksheetFunction.prototype["MROUND"]          =  ApiWorksheetFunction.prototype.MROUND;
	ApiWorksheetFunction.prototype["MULTINOMIAL"]     =  ApiWorksheetFunction.prototype.MULTINOMIAL;
	ApiWorksheetFunction.prototype["MUNIT"]           =  ApiWorksheetFunction.prototype.MUNIT;
	ApiWorksheetFunction.prototype["ODD"]             =  ApiWorksheetFunction.prototype.ODD;
	ApiWorksheetFunction.prototype["PI"]              =  ApiWorksheetFunction.prototype.PI;
	ApiWorksheetFunction.prototype["POWER"]           =  ApiWorksheetFunction.prototype.POWER;
	ApiWorksheetFunction.prototype["PRODUCT"]         =  ApiWorksheetFunction.prototype.PRODUCT;
	ApiWorksheetFunction.prototype["QUOTIENT"]        =  ApiWorksheetFunction.prototype.QUOTIENT;
	ApiWorksheetFunction.prototype["RADIANS"]         =  ApiWorksheetFunction.prototype.RADIANS;
	ApiWorksheetFunction.prototype["RAND"]            =  ApiWorksheetFunction.prototype.RAND;
	ApiWorksheetFunction.prototype["RANDBETWEEN"]     =  ApiWorksheetFunction.prototype.RANDBETWEEN;
	ApiWorksheetFunction.prototype["ROMAN"]           =  ApiWorksheetFunction.prototype.ROMAN;
	ApiWorksheetFunction.prototype["ROUND"]           =  ApiWorksheetFunction.prototype.ROUND;
	ApiWorksheetFunction.prototype["ROUNDDOWN"]       =  ApiWorksheetFunction.prototype.ROUNDDOWN;
	ApiWorksheetFunction.prototype["ROUNDUP"]         =  ApiWorksheetFunction.prototype.ROUNDUP;
	ApiWorksheetFunction.prototype["SEC"]             =  ApiWorksheetFunction.prototype.SEC;
	ApiWorksheetFunction.prototype["SECH"]            =  ApiWorksheetFunction.prototype.SECH;
	ApiWorksheetFunction.prototype["SERIESSUM"]       =  ApiWorksheetFunction.prototype.SERIESSUM;
	ApiWorksheetFunction.prototype["SIGN"]            =  ApiWorksheetFunction.prototype.SIGN;
	ApiWorksheetFunction.prototype["SIN"]             =  ApiWorksheetFunction.prototype.SIN;
	ApiWorksheetFunction.prototype["SINH"]            =  ApiWorksheetFunction.prototype.SINH;
	ApiWorksheetFunction.prototype["SQRT"]            =  ApiWorksheetFunction.prototype.SQRT;
	ApiWorksheetFunction.prototype["SQRTPI"]          =  ApiWorksheetFunction.prototype.SQRTPI;
	ApiWorksheetFunction.prototype["SUBTOTAL"]        =  ApiWorksheetFunction.prototype.SUBTOTAL;
	ApiWorksheetFunction.prototype["SUM"]             =  ApiWorksheetFunction.prototype.SUM;
	ApiWorksheetFunction.prototype["SUMIF"]           =  ApiWorksheetFunction.prototype.SUMIF;
	ApiWorksheetFunction.prototype["SUMIFS"]          =  ApiWorksheetFunction.prototype.SUMIFS;
	ApiWorksheetFunction.prototype["SUMSQ"]           =  ApiWorksheetFunction.prototype.SUMSQ;
	ApiWorksheetFunction.prototype["TAN"]             =  ApiWorksheetFunction.prototype.TAN;
	ApiWorksheetFunction.prototype["TANH"]            =  ApiWorksheetFunction.prototype.TANH;
	ApiWorksheetFunction.prototype["TRUNC"]           =  ApiWorksheetFunction.prototype.TRUNC;
	ApiWorksheetFunction.prototype["CHOOSE"]          =  ApiWorksheetFunction.prototype.CHOOSE;
	ApiWorksheetFunction.prototype["COLUMNS"]         =  ApiWorksheetFunction.prototype.COLUMNS;
	ApiWorksheetFunction.prototype["HLOOKUP"]         =  ApiWorksheetFunction.prototype.HLOOKUP;
	ApiWorksheetFunction.prototype["HYPERLINK"]       =  ApiWorksheetFunction.prototype.HYPERLINK;
	ApiWorksheetFunction.prototype["INDEX"]           =  ApiWorksheetFunction.prototype.INDEX;
	ApiWorksheetFunction.prototype["LOOKUP"]          =  ApiWorksheetFunction.prototype.LOOKUP;
	ApiWorksheetFunction.prototype["MATCH"]           =  ApiWorksheetFunction.prototype.MATCH;
	ApiWorksheetFunction.prototype["ROWS"]            =  ApiWorksheetFunction.prototype.ROWS;
	ApiWorksheetFunction.prototype["TRANSPOSE"]       =  ApiWorksheetFunction.prototype.TRANSPOSE;
	ApiWorksheetFunction.prototype["VLOOKUP"]         =  ApiWorksheetFunction.prototype.VLOOKUP;
	ApiWorksheetFunction.prototype["ISERR"]           =  ApiWorksheetFunction.prototype.ISERR;
	ApiWorksheetFunction.prototype["ISERROR"]         =  ApiWorksheetFunction.prototype.ISERROR;
	ApiWorksheetFunction.prototype["ISEVEN"]          =  ApiWorksheetFunction.prototype.ISEVEN;
	ApiWorksheetFunction.prototype["ISFORMULA"]       =  ApiWorksheetFunction.prototype.ISFORMULA;
	ApiWorksheetFunction.prototype["ISLOGICAL"]       =  ApiWorksheetFunction.prototype.ISLOGICAL;
	ApiWorksheetFunction.prototype["ISNA"]            =  ApiWorksheetFunction.prototype.ISNA;
	ApiWorksheetFunction.prototype["ISNONTEXT"]       =  ApiWorksheetFunction.prototype.ISNONTEXT;
	ApiWorksheetFunction.prototype["ISNUMBER"]        =  ApiWorksheetFunction.prototype.ISNUMBER;
	ApiWorksheetFunction.prototype["ISODD"]           =  ApiWorksheetFunction.prototype.ISODD;
	ApiWorksheetFunction.prototype["ISREF"]           =  ApiWorksheetFunction.prototype.ISREF;
	ApiWorksheetFunction.prototype["ISTEXT"]          =  ApiWorksheetFunction.prototype.ISTEXT;
	ApiWorksheetFunction.prototype["N"]               =  ApiWorksheetFunction.prototype.N;
	ApiWorksheetFunction.prototype["NA"]              =  ApiWorksheetFunction.prototype.NA;
	ApiWorksheetFunction.prototype["SHEET"]           =  ApiWorksheetFunction.prototype.SHEET;
	ApiWorksheetFunction.prototype["SHEETS"]          =  ApiWorksheetFunction.prototype.SHEETS;
	ApiWorksheetFunction.prototype["TYPE"]            =  ApiWorksheetFunction.prototype.TYPE;
	ApiWorksheetFunction.prototype["AND"]             =  ApiWorksheetFunction.prototype.AND;
	ApiWorksheetFunction.prototype["FALSE"]           =  ApiWorksheetFunction.prototype.FALSE;
	ApiWorksheetFunction.prototype["IF"]              =  ApiWorksheetFunction.prototype.IF;
	ApiWorksheetFunction.prototype["IFERROR"]         =  ApiWorksheetFunction.prototype.IFERROR;
	ApiWorksheetFunction.prototype["IFNA"]            =  ApiWorksheetFunction.prototype.IFNA;
	ApiWorksheetFunction.prototype["NOT"]             =  ApiWorksheetFunction.prototype.NOT;
	ApiWorksheetFunction.prototype["OR"]              =  ApiWorksheetFunction.prototype.OR;
	ApiWorksheetFunction.prototype["TRUE"]            =  ApiWorksheetFunction.prototype.TRUE;
	ApiWorksheetFunction.prototype["XOR"]             =  ApiWorksheetFunction.prototype.XOR;

	ApiWorksheetFunction.prototype["BETA_DIST"]       =  ApiWorksheetFunction.prototype.BETA_DIST;
	ApiWorksheetFunction.prototype["BETA_INV"]        =  ApiWorksheetFunction.prototype.BETA_INV;
	ApiWorksheetFunction.prototype["BINOM_DIST"]      =  ApiWorksheetFunction.prototype.BINOM_DIST;
	ApiWorksheetFunction.prototype["BINOM_DIST_RANGE"]=  ApiWorksheetFunction.prototype.BINOM_DIST_RANGE;
	ApiWorksheetFunction.prototype["BINOM_INV"]       =  ApiWorksheetFunction.prototype.BINOM_INV;
	ApiWorksheetFunction.prototype["CHISQ_DIST"]      =  ApiWorksheetFunction.prototype.CHISQ_DIST;
	ApiWorksheetFunction.prototype["CHISQ_DIST_RT"]   =  ApiWorksheetFunction.prototype.CHISQ_DIST_RT;
	ApiWorksheetFunction.prototype["CHISQ_INV"]       =  ApiWorksheetFunction.prototype.CHISQ_INV;
	ApiWorksheetFunction.prototype["CHISQ_INV_RT"]    =  ApiWorksheetFunction.prototype.CHISQ_INV_RT;
	ApiWorksheetFunction.prototype["CONFIDENCE_NORM"] =  ApiWorksheetFunction.prototype.CONFIDENCE_NORM;
	ApiWorksheetFunction.prototype["CONFIDENCE_T"]    =  ApiWorksheetFunction.prototype.CONFIDENCE_T;
	ApiWorksheetFunction.prototype["EXPON_DIST"]      =  ApiWorksheetFunction.prototype.EXPON_DIST;
	ApiWorksheetFunction.prototype["F_DIST"]          =  ApiWorksheetFunction.prototype.F_DIST;
	ApiWorksheetFunction.prototype["F_INV"]           =  ApiWorksheetFunction.prototype.F_INV;
	ApiWorksheetFunction.prototype["FORECAST_ETS"]    =  ApiWorksheetFunction.prototype.FORECAST_ETS;
	ApiWorksheetFunction.prototype["GAMMA_DIST"]      =  ApiWorksheetFunction.prototype.GAMMA_DIST;
	ApiWorksheetFunction.prototype["GAMMA_INV"]       =  ApiWorksheetFunction.prototype.GAMMA_INV;
	ApiWorksheetFunction.prototype["GAMMALN_PRECISE"] =  ApiWorksheetFunction.prototype.GAMMALN_PRECISE;
	ApiWorksheetFunction.prototype["HYPGEOM_DIST"]    =  ApiWorksheetFunction.prototype.HYPGEOM_DIST;
	ApiWorksheetFunction.prototype["LOGNORM_DIST"]    =  ApiWorksheetFunction.prototype.LOGNORM_DIST;
	ApiWorksheetFunction.prototype["LOGNORM_INV"]     =  ApiWorksheetFunction.prototype.LOGNORM_INV;
	ApiWorksheetFunction.prototype["NEGBINOM_DIST"]   =  ApiWorksheetFunction.prototype.NEGBINOM_DIST;
	ApiWorksheetFunction.prototype["NORM_DIST"]       =  ApiWorksheetFunction.prototype.NORM_DIST;
	ApiWorksheetFunction.prototype["NORM_INV"]        =  ApiWorksheetFunction.prototype.NORM_INV;
	ApiWorksheetFunction.prototype["PERCENTILE_EXC"]  =  ApiWorksheetFunction.prototype.PERCENTILE_EXC;
	ApiWorksheetFunction.prototype["PERCENTILE_INC"]  =  ApiWorksheetFunction.prototype.PERCENTILE_INC;
	ApiWorksheetFunction.prototype["PERCENTRANK_EXC"] =  ApiWorksheetFunction.prototype.PERCENTRANK_EXC;
	ApiWorksheetFunction.prototype["PERCENTRANK_INC"] =  ApiWorksheetFunction.prototype.PERCENTRANK_INC;
	ApiWorksheetFunction.prototype["POISSON_DIST"]    =  ApiWorksheetFunction.prototype.POISSON_DIST;
	ApiWorksheetFunction.prototype["QUARTILE_EXC"]    =  ApiWorksheetFunction.prototype.QUARTILE_EXC;
	ApiWorksheetFunction.prototype["QUARTILE_INC"]    =  ApiWorksheetFunction.prototype.QUARTILE_INC;
	ApiWorksheetFunction.prototype["RANK_AVG"]        =  ApiWorksheetFunction.prototype.RANK_AVG;
	ApiWorksheetFunction.prototype["RANK_EQ"]         =  ApiWorksheetFunction.prototype.RANK_EQ;
	ApiWorksheetFunction.prototype["SKEW_P"]          =  ApiWorksheetFunction.prototype.SKEW_P;
	ApiWorksheetFunction.prototype["STDEV_S"]         =  ApiWorksheetFunction.prototype.STDEV_S;
	ApiWorksheetFunction.prototype["STDEV_P"]         =  ApiWorksheetFunction.prototype.STDEV_P;
	ApiWorksheetFunction.prototype["T_DIST"]          =  ApiWorksheetFunction.prototype.T_DIST;
	ApiWorksheetFunction.prototype["T_INV"]           =  ApiWorksheetFunction.prototype.T_INV;
	ApiWorksheetFunction.prototype["VAR_P"]           =  ApiWorksheetFunction.prototype.VAR_P;
	ApiWorksheetFunction.prototype["VAR_S"]           =  ApiWorksheetFunction.prototype.VAR_S;
	ApiWorksheetFunction.prototype["WEIBULL_DIST"]    =  ApiWorksheetFunction.prototype.WEIBULL_DIST;
	ApiWorksheetFunction.prototype["Z_TEST"]          =  ApiWorksheetFunction.prototype.Z_TEST;
	ApiWorksheetFunction.prototype["NETWORKDAYS_INTL"]=  ApiWorksheetFunction.prototype.NETWORKDAYS_INTL;
	ApiWorksheetFunction.prototype["WORKDAY_INTL"]    =  ApiWorksheetFunction.prototype.WORKDAY_INTL;
	ApiWorksheetFunction.prototype["ERF_PRECISE"]     =  ApiWorksheetFunction.prototype.ERF_PRECISE;
	ApiWorksheetFunction.prototype["ERFC_PRECISE"]    =  ApiWorksheetFunction.prototype.ERFC_PRECISE;
	ApiWorksheetFunction.prototype["CEILING_MATH"]    =  ApiWorksheetFunction.prototype.CEILING_MATH;
	ApiWorksheetFunction.prototype["CEILING_PRECISE"] =  ApiWorksheetFunction.prototype.CEILING_PRECISE;
	ApiWorksheetFunction.prototype["ECMA_CEILING"]    =  ApiWorksheetFunction.prototype.ECMA_CEILING;
	ApiWorksheetFunction.prototype["FLOOR_PRECISE"]   =  ApiWorksheetFunction.prototype.FLOOR_PRECISE;
	ApiWorksheetFunction.prototype["FLOOR_MATH"]      =  ApiWorksheetFunction.prototype.FLOOR_MATH;
	ApiWorksheetFunction.prototype["ISO_CEILING"]     =  ApiWorksheetFunction.prototype.ISO_CEILING;
	ApiWorksheetFunction.prototype["ERROR_TYPE"]      =  ApiWorksheetFunction.prototype.ERROR_TYPE;
	ApiWorksheetFunction.prototype["FORECAST_ETS_CONFINT"] =  ApiWorksheetFunction.prototype.FORECAST_ETS_CONFINT;
	ApiWorksheetFunction.prototype["FORECAST_ETS_SEASONALITY"] =  ApiWorksheetFunction.prototype.FORECAST_ETS_SEASONALITY;
	ApiWorksheetFunction.prototype["FORECAST_ETS_STAT"] =  ApiWorksheetFunction.prototype.FORECAST_ETS_STAT;
	ApiWorksheetFunction.prototype["F_DIST_RT"]      =  ApiWorksheetFunction.prototype.F_DIST_RT;
	ApiWorksheetFunction.prototype["F_INV_RT"]       =  ApiWorksheetFunction.prototype.F_INV_RT;
	ApiWorksheetFunction.prototype["NORM_S_DIST"]    =  ApiWorksheetFunction.prototype.NORM_S_DIST;
	ApiWorksheetFunction.prototype["NORM_S_INV"]     =  ApiWorksheetFunction.prototype.NORM_S_INV;
	ApiWorksheetFunction.prototype["T_DIST_2T"]      =  ApiWorksheetFunction.prototype.T_DIST_2T;
	ApiWorksheetFunction.prototype["T_DIST_RT"]      =  ApiWorksheetFunction.prototype.T_DIST_RT;
	ApiWorksheetFunction.prototype["T_INV_2T"]       =  ApiWorksheetFunction.prototype.T_INV_2T;


	ApiPivotTable.prototype["AddDataField"]                       = ApiPivotTable.prototype.AddDataField;
	ApiPivotTable.prototype["AddFields"]                          = ApiPivotTable.prototype.AddFields;
	ApiPivotTable.prototype["ClearAllFilters"]                    = ApiPivotTable.prototype.ClearAllFilters;
	ApiPivotTable.prototype["ClearTable"]                         = ApiPivotTable.prototype.ClearTable
	ApiPivotTable.prototype["GetData"]                            = ApiPivotTable.prototype.GetData;
	ApiPivotTable.prototype["GetPivotData"]                       = ApiPivotTable.prototype.GetPivotData;
	ApiPivotTable.prototype["GetPivotFields"]                     = ApiPivotTable.prototype.GetPivotFields;
	ApiPivotTable.prototype["PivotValueCell"]                     = ApiPivotTable.prototype.PivotValueCell;
	ApiPivotTable.prototype["ShowDetails"]                        = ApiPivotTable.prototype.ShowDetails;
	ApiPivotTable.prototype["RefreshTable"]                       = ApiPivotTable.prototype.RefreshTable;
	ApiPivotTable.prototype["Update"]                             = ApiPivotTable.prototype.Update;
	ApiPivotTable.prototype["SetRepeatAllLabels"]                 = ApiPivotTable.prototype.SetRepeatAllLabels;
	ApiPivotTable.prototype["SetRowAxisLayout"]                   = ApiPivotTable.prototype.SetRowAxisLayout;
	ApiPivotTable.prototype["SetSubtotalLocation"]                = ApiPivotTable.prototype.SetSubtotalLocation;
	ApiPivotTable.prototype["RemoveField"]                        = ApiPivotTable.prototype.RemoveField;
	ApiPivotTable.prototype["MoveField"]                          = ApiPivotTable.prototype.MoveField;
	ApiPivotTable.prototype["Select"]                             = ApiPivotTable.prototype.Select;
	ApiPivotTable.prototype["GetColumnFields"]                    = ApiPivotTable.prototype.GetColumnFields;
	ApiPivotTable.prototype["GetDataFields"]                      = ApiPivotTable.prototype.GetDataFields;
	ApiPivotTable.prototype["GetHiddenFields"]                    = ApiPivotTable.prototype.GetHiddenFields;
	ApiPivotTable.prototype["GetVisibleFields"]                   = ApiPivotTable.prototype.GetVisibleFields;
	ApiPivotTable.prototype["GetPageFields"]                      = ApiPivotTable.prototype.GetPageFields;
	ApiPivotTable.prototype["GetRowFields"]                       = ApiPivotTable.prototype.GetRowFields;
	ApiPivotTable.prototype["GetName"]                            = ApiPivotTable.prototype.GetName;
	ApiPivotTable.prototype["SetName"]                            = ApiPivotTable.prototype.SetName;
	ApiPivotTable.prototype["GetColumnGrand"]                     = ApiPivotTable.prototype.GetColumnGrand;
	ApiPivotTable.prototype["SetColumnGrand"]                     = ApiPivotTable.prototype.SetColumnGrand;
	ApiPivotTable.prototype["GetRowGrand"]                        = ApiPivotTable.prototype.GetRowGrand;
	ApiPivotTable.prototype["SetRowGrand"]                        = ApiPivotTable.prototype.SetRowGrand;
	ApiPivotTable.prototype["GetDisplayFieldsInReportFilterArea"] = ApiPivotTable.prototype.GetDisplayFieldsInReportFilterArea;
	ApiPivotTable.prototype["SetDisplayFieldsInReportFilterArea"] = ApiPivotTable.prototype.SetDisplayFieldsInReportFilterArea;
	ApiPivotTable.prototype["GetDisplayFieldCaptions"]            = ApiPivotTable.prototype.GetDisplayFieldCaptions;
	ApiPivotTable.prototype["SetDisplayFieldCaptions"]            = ApiPivotTable.prototype.SetDisplayFieldCaptions;
	ApiPivotTable.prototype["GetTitle"]                           = ApiPivotTable.prototype.GetTitle;
	ApiPivotTable.prototype["SetTitle"]                           = ApiPivotTable.prototype.SetTitle;
	ApiPivotTable.prototype["GetDescription"]                    = ApiPivotTable.prototype.GetDescription;
	ApiPivotTable.prototype["SetDescription"]                     = ApiPivotTable.prototype.SetDescription;
	ApiPivotTable.prototype["GetStyleName"]                       = ApiPivotTable.prototype.GetStyleName;
	ApiPivotTable.prototype["SetStyleName"]                       = ApiPivotTable.prototype.SetStyleName;
	ApiPivotTable.prototype["GetTableStyleRowHeaders"]            = ApiPivotTable.prototype.GetTableStyleRowHeaders;
	ApiPivotTable.prototype["SetTableStyleRowHeaders"]            = ApiPivotTable.prototype.SetTableStyleRowHeaders;
	ApiPivotTable.prototype["GetTableStyleColumnHeaders"]         = ApiPivotTable.prototype.GetTableStyleColumnHeaders;
	ApiPivotTable.prototype["SetTableStyleColumnHeaders"]         = ApiPivotTable.prototype.SetTableStyleColumnHeaders;
	ApiPivotTable.prototype["GetTableStyleRowStripes"]            = ApiPivotTable.prototype.GetTableStyleRowStripes;
	ApiPivotTable.prototype["SetTableStyleRowStripes"]            = ApiPivotTable.prototype.SetTableStyleRowStripes;
	ApiPivotTable.prototype["GetTableStyleColumnStripes"]         = ApiPivotTable.prototype.GetTableStyleColumnStripes;
	ApiPivotTable.prototype["SetTableStyleColumnStripes"]         = ApiPivotTable.prototype.SetTableStyleColumnStripes;
	ApiPivotTable.prototype["GetSource"]                          = ApiPivotTable.prototype.GetSource;
	ApiPivotTable.prototype["SetSource"]                          = ApiPivotTable.prototype.SetSource;
	ApiPivotTable.prototype["GetColumnRange"]                     = ApiPivotTable.prototype.GetColumnRange;
	ApiPivotTable.prototype["GetRowRange"]                        = ApiPivotTable.prototype.GetRowRange;
	ApiPivotTable.prototype["GetDataBodyRange"]                   = ApiPivotTable.prototype.GetDataBodyRange;
	ApiPivotTable.prototype["GetTableRange1"]                     = ApiPivotTable.prototype.GetTableRange1;
	ApiPivotTable.prototype["GetTableRange2"]                     = ApiPivotTable.prototype.GetTableRange2;
	ApiPivotTable.prototype["GetGrandTotalName"]                  = ApiPivotTable.prototype.GetGrandTotalName;
	ApiPivotTable.prototype["SetGrandTotalName"]                  = ApiPivotTable.prototype.SetGrandTotalName;
	ApiPivotTable.prototype["SetLayoutBlankLine"]                 = ApiPivotTable.prototype.SetLayoutBlankLine;
	ApiPivotTable.prototype["SetLayoutSubtotals"]                 = ApiPivotTable.prototype.SetLayoutSubtotals;
	ApiPivotTable.prototype["GetParent"]                          = ApiPivotTable.prototype.GetParent;

	ApiPivotDataField.prototype["Remove"]               = ApiPivotDataField.prototype.Remove;
	ApiPivotDataField.prototype["Move"]                 = ApiPivotDataField.prototype.Move;
	ApiPivotDataField.prototype["SetFunction"]          = ApiPivotDataField.prototype.SetFunction;
	ApiPivotDataField.prototype["GetFunction"]          = ApiPivotDataField.prototype.GetFunction;
	ApiPivotDataField.prototype["GetPosition"]          = ApiPivotDataField.prototype.GetPosition;
	ApiPivotDataField.prototype["SetPosition"]          = ApiPivotDataField.prototype.SetPosition;
	ApiPivotDataField.prototype["GetOrientation"]       = ApiPivotDataField.prototype.GetOrientation;
	ApiPivotDataField.prototype["GetValue"]             = ApiPivotDataField.prototype.GetValue;
	ApiPivotDataField.prototype["SetValue"]             = ApiPivotDataField.prototype.SetValue;
	ApiPivotDataField.prototype["GetCaption"]           = ApiPivotDataField.prototype.GetCaption;
	ApiPivotDataField.prototype["SetCaption"]           = ApiPivotDataField.prototype.SetCaption;
	ApiPivotDataField.prototype["GetName"]              = ApiPivotDataField.prototype.GetName;
	ApiPivotDataField.prototype["SetName"]              = ApiPivotDataField.prototype.SetName;
	ApiPivotDataField.prototype["GetNumberFormat"]      = ApiPivotDataField.prototype.GetNumberFormat;
	ApiPivotDataField.prototype["SetNumberFormat"]      = ApiPivotDataField.prototype.SetNumberFormat;
	ApiPivotDataField.prototype["GetIndex"]             = ApiPivotDataField.prototype.GetIndex;
	ApiPivotDataField.prototype["GetPivotField"]        = ApiPivotDataField.prototype.GetPivotField;

	ApiPivotField.prototype["ClearAllFilters"]           = ApiPivotField.prototype.ClearAllFilters;
	ApiPivotField.prototype["ClearLabelFilters"]         = ApiPivotField.prototype.ClearLabelFilters;
	ApiPivotField.prototype["ClearManualFilters"]        = ApiPivotField.prototype.ClearManualFilters;
	ApiPivotField.prototype["ClearValueFilters"]         = ApiPivotField.prototype.ClearValueFilters;
	ApiPivotField.prototype["GetPivotItems"]             = ApiPivotField.prototype.GetPivotItems;
	ApiPivotField.prototype["Move"]                      = ApiPivotField.prototype.Move;
	ApiPivotField.prototype["Remove"]                    = ApiPivotField.prototype.Remove;
	ApiPivotField.prototype["GetPosition"]               = ApiPivotField.prototype.GetPosition;
	ApiPivotField.prototype["SetPosition"]               = ApiPivotField.prototype.SetPosition;
	ApiPivotField.prototype["GetOrientation"]            = ApiPivotField.prototype.GetOrientation;
	ApiPivotField.prototype["SetOrientation"]            = ApiPivotField.prototype.SetOrientation;
	ApiPivotField.prototype["GetValue"]                  = ApiPivotField.prototype.GetValue;
	ApiPivotField.prototype["SetValue"]                  = ApiPivotField.prototype.SetValue;
	ApiPivotField.prototype["GetCaption"]                = ApiPivotField.prototype.GetCaption;
	ApiPivotField.prototype["SetCaption"]                = ApiPivotField.prototype.SetCaption;
	ApiPivotField.prototype["GetName"]                   = ApiPivotField.prototype.GetName;
	ApiPivotField.prototype["SetName"]                   = ApiPivotField.prototype.SetName;
	ApiPivotField.prototype["GetSourceName"]             = ApiPivotField.prototype.GetSourceName;
	ApiPivotField.prototype["GetIndex"]                  = ApiPivotField.prototype.GetIndex;
	ApiPivotField.prototype["GetTable"]                  = ApiPivotField.prototype.GetTable;
	ApiPivotField.prototype["GetParent"]                 = ApiPivotField.prototype.GetParent;
	ApiPivotField.prototype["GetLayoutCompactRow"]       = ApiPivotField.prototype.GetLayoutCompactRow;
	ApiPivotField.prototype["SetLayoutCompactRow"]       = ApiPivotField.prototype.SetLayoutCompactRow;
	ApiPivotField.prototype["GetLayoutForm"]             = ApiPivotField.prototype.GetLayoutForm;
	ApiPivotField.prototype["SetLayoutForm"]             = ApiPivotField.prototype.SetLayoutForm;
	ApiPivotField.prototype["GetLayoutPageBreak"]        = ApiPivotField.prototype.GetLayoutPageBreak;
	ApiPivotField.prototype["SetLayoutPageBreak"]        = ApiPivotField.prototype.SetLayoutPageBreak;
	ApiPivotField.prototype["GetShowingInAxis"]          = ApiPivotField.prototype.GetShowingInAxis;
	ApiPivotField.prototype["GetRepeatLabels"]           = ApiPivotField.prototype.GetRepeatLabels;
	ApiPivotField.prototype["SetRepeatLabels"]           = ApiPivotField.prototype.SetRepeatLabels;
	ApiPivotField.prototype["GetLayoutBlankLine"]        = ApiPivotField.prototype.GetLayoutBlankLine;
	ApiPivotField.prototype["SetLayoutBlankLine"]        = ApiPivotField.prototype.SetLayoutBlankLine;
	ApiPivotField.prototype["GetShowAllItems"]           = ApiPivotField.prototype.GetShowAllItems;
	ApiPivotField.prototype["SetShowAllItems"]           = ApiPivotField.prototype.SetShowAllItems;
	ApiPivotField.prototype["GetLayoutSubtotals"]        = ApiPivotField.prototype.GetLayoutSubtotals;
	ApiPivotField.prototype["SetLayoutSubtotals"]        = ApiPivotField.prototype.SetLayoutSubtotals;
	ApiPivotField.prototype["GetLayoutSubtotalLocation"] = ApiPivotField.prototype.GetLayoutSubtotalLocation;
	ApiPivotField.prototype["SetLayoutSubtotalLocation"] = ApiPivotField.prototype.SetLayoutSubtotalLocation;
	ApiPivotField.prototype["GetSubtotalName"]           = ApiPivotField.prototype.GetSubtotalName;
	ApiPivotField.prototype["SetSubtotalName"]           = ApiPivotField.prototype.SetSubtotalName;
	ApiPivotField.prototype["GetSubtotals"]              = ApiPivotField.prototype.GetSubtotals;
	ApiPivotField.prototype["SetSubtotals"]              = ApiPivotField.prototype.SetSubtotals;
	ApiPivotField.prototype["GetDragToColumn"]           = ApiPivotField.prototype.GetDragToColumn;
	ApiPivotField.prototype["SetDragToColumn"]           = ApiPivotField.prototype.SetDragToColumn;
	ApiPivotField.prototype["GetDragToRow"]              = ApiPivotField.prototype.GetDragToRow;
	ApiPivotField.prototype["SetDragToRow"]              = ApiPivotField.prototype.SetDragToRow;
	ApiPivotField.prototype["GetDragToData"]             = ApiPivotField.prototype.GetDragToData;
	ApiPivotField.prototype["SetDragToData"]             = ApiPivotField.prototype.SetDragToData;
	ApiPivotField.prototype["GetDragToPage"]             = ApiPivotField.prototype.GetDragToPage;
	ApiPivotField.prototype["SetDragToPage"]             = ApiPivotField.prototype.SetDragToPage;
	ApiPivotField.prototype["GetCurrentPage"]            = ApiPivotField.prototype.GetCurrentPage;
	ApiPivotField.prototype["GetNumberFormat"]           = ApiPivotField.prototype.GetNumberFormat;
	ApiPivotField.prototype["SetNumberFormat"]           = ApiPivotField.prototype.SetNumberFormat;
	ApiPivotField.prototype["SetFunction"]               = ApiPivotField.prototype.SetFunction;
	ApiPivotField.prototype["GetFunction"]               = ApiPivotField.prototype.GetFunction;

	ApiPivotItem.prototype["GetName"]    = ApiPivotItem.prototype.GetName;
	ApiPivotItem.prototype["GetCaption"] = ApiPivotItem.prototype.GetCaption;
	ApiPivotItem.prototype["GetValue"]   = ApiPivotItem.prototype.GetValue;
	ApiPivotItem.prototype["GetParent"]  = ApiPivotItem.prototype.GetParent;

	function private_SetCoords(oDrawing, oWorksheet, nExtX, nExtY, nFromCol, nColOffset, nFromRow, nRowOffset, pos) {
		oDrawing.x = 0;
		oDrawing.y = 0;
		oDrawing.extX = 0;
		oDrawing.extY = 0;
		AscFormat.CheckSpPrXfrm(oDrawing);
		oDrawing.spPr.xfrm.setExtX(nExtX / 36000.0);
		oDrawing.spPr.xfrm.setExtY(nExtY / 36000.0);
		oDrawing.setBDeleted(false);
		oDrawing.setWorksheet(oWorksheet);
		oDrawing.addToDrawingObjects(pos);
		oDrawing.setDrawingBaseType(AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
		oDrawing.setDrawingBaseCoords(nFromCol, nColOffset / 36000.0, nFromRow, nRowOffset / 36000.0, 0, 0, 0, 0, 0, 0, 0, 0);
		oDrawing.setDrawingBaseExt(nExtX / 36000.0, nExtY / 36000.0);
	}

	function private_MakeBorder(lineStyle, color) {
		var border = new AscCommonExcel.BorderProp();
		switch (lineStyle) {
			case 'Double':
				border.setStyle(Asc.c_oAscBorderStyles.Double);
				break;
			case 'Hair':
				border.setStyle(Asc.c_oAscBorderStyles.Hair);
				break;
			case 'DashDotDot':
				border.setStyle(Asc.c_oAscBorderStyles.DashDotDot);
				break;
			case 'DashDot':
				border.setStyle(Asc.c_oAscBorderStyles.DashDot);
				break;
			case 'Dotted':
				border.setStyle(Asc.c_oAscBorderStyles.Dotted);
				break;
			case 'Dashed':
				border.setStyle(Asc.c_oAscBorderStyles.Dashed);
				break;
			case 'Thin':
				border.setStyle(Asc.c_oAscBorderStyles.Thin);
				break;
			case 'MediumDashDotDot':
				border.setStyle(Asc.c_oAscBorderStyles.MediumDashDotDot);
				break;
			case 'SlantDashDot':
				border.setStyle(Asc.c_oAscBorderStyles.SlantDashDot);
				break;
			case 'MediumDashDot':
				border.setStyle(Asc.c_oAscBorderStyles.MediumDashDot);
				break;
			case 'MediumDashed':
				border.setStyle(Asc.c_oAscBorderStyles.MediumDashed);
				break;
			case 'Medium':
				border.setStyle(Asc.c_oAscBorderStyles.Medium);
				break;
			case 'Thick':
				border.setStyle(Asc.c_oAscBorderStyles.Thick);
				break;
			case 'None':
			default:
				border.setStyle(Asc.c_oAscBorderStyles.None);
				break;
		}

		if (color) {
			border.c = color.color;
		}
		return border;
	}

	function private_AddDefName(wb, name, ref, sheetInd, hidden) {
		let res = wb.checkDefName(name);
		if (!res.status) {
			logError(new Error('Invalid name.'));
			return false;
		}
		res = wb.oApi.asc_checkDataRange(Asc.c_oAscSelectionDialogType.Chart, ref, false);
		if (res === Asc.c_oAscError.ID.DataRangeError) {
			logError(new Error('Invalid range.'));
			return false;
		}
		if (sheetInd) {
			sheetInd = (wb.getWorksheet(sheetInd)) ? sheetInd : undefined;
		}
		let defName = new Asc.asc_CDefName(name, ref, sheetInd, undefined, hidden, undefined, undefined, true);
		wb.oApi.asc_setDefinedNames(defName);

		return true;
	}

	function private_MM2EMU(mm) {
		return mm * 36000.0;
	}

	function private_GetDrawingLockType(sType) {
		var nLockType = -1;
		switch (sType) {
			case "noGrp":
				nLockType = AscFormat.LOCKS_MASKS.noGrp;
				break;
			case "noUngrp":
				nLockType = AscFormat.LOCKS_MASKS.noUngrp;
				break;
			case "noSelect":
				nLockType = AscFormat.LOCKS_MASKS.noSelect;
				break;
			case "noRot":
				nLockType = AscFormat.LOCKS_MASKS.noRot;
				break;
			case "noChangeAspect":
				nLockType = AscFormat.LOCKS_MASKS.noChangeAspect;
				break;
			case "noMove":
				nLockType = AscFormat.LOCKS_MASKS.noMove;
				break;
			case "noResize":
				nLockType = AscFormat.LOCKS_MASKS.noResize;
				break;
			case "noEditPoints":
				nLockType = AscFormat.LOCKS_MASKS.noEditPoints;
				break;
			case "noAdjustHandles":
				nLockType = AscFormat.LOCKS_MASKS.noAdjustHandles;
				break;
			case "noChangeArrowheads":
				nLockType = AscFormat.LOCKS_MASKS.noChangeArrowheads;
				break;
			case "noChangeShapeType":
				nLockType = AscFormat.LOCKS_MASKS.noChangeShapeType;
				break;
			case "noDrilldown":
				nLockType = AscFormat.LOCKS_MASKS.noDrilldown;
				break;
			case "noTextEdit":
				nLockType = AscFormat.LOCKS_MASKS.noTextEdit;
				break;
			case "noCrop":
				nLockType = AscFormat.LOCKS_MASKS.noCrop;
				break;
			case "txBox":
				nLockType = AscFormat.LOCKS_MASKS.txBox;
				break;
		}

		return nLockType;
	}

	/**
	 * Validates the parsed JSDoc or options for custom functions.
	 * @param {object} jsdoc - Parsed JSDoc object.
	 * @returns {boolean} - Returns false if JSDoc isn't valid
	 */
	function private_ValidateParamsForCustomFunction(jsdoc) {
		let result = true;
		const types = [
			'number', 'string', 'boolean', 'bool', 'any',
			'number[]', 'string[]', 'boolean[]', 'bool[]', 'any[]',
			'number[][]', 'string[][]', 'boolean[][]', 'bool[][]', 'any[][]'
		];

		if (jsdoc.returnInfo && !types.includes(jsdoc.returnInfo.type)) {
			result = false;
		}

		for (let index = 0; result && index < jsdoc.params.length; index++) {
			const param = jsdoc.params[index];
			if (!types.includes(param.type)) {
				result = false;
				break;
			}
		}

		return result;
	}

	function logError(err) {
		if (console.error)
			console.error(err);
		else
			console.log(err);
	}

	function throwException(err) {
		if (!console.error)
			logError(err);
		throw err;
	}
	function private_executeOtherActiveSheet(ws, ranges, func) {
		let oldActiveSheet = ws && ws.workbook && ws.workbook.getActive();
		let isChangedActiveSheet;
		if (oldActiveSheet != null && ws && oldActiveSheet !== ws.index) {
			ws.workbook.setActive(ws.index);
			isChangedActiveSheet = true;
		}

		let oldSelection;
		if (ranges) {
			oldSelection = ws.selectionRange.clone();
			let newSelection = new AscCommonExcel.SelectionRange(ws);
			if (ranges.length === 1) {
				let bbox = ranges[0].bbox;
				newSelection.assign2(bbox);
			} else {
				for (let i = 0; i < ranges.length; i++) {
					if (i !== 0) {
						newSelection.addRange();
					}
					newSelection.getLast().assign2(ranges[i].bbox);
				}
			}
			newSelection.Select(true);
		}

		func();

		isChangedActiveSheet && ws.workbook.setActive(oldActiveSheet);
		oldSelection && oldSelection.Select(true);
	}

	function private_MakeError(message) {
		console.error(new Error(message) );
	};

}(window, null));

