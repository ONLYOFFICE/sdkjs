/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	 * Class representing a chart.
	 * @constructor
	 */
	function ApiChart(oChart) {
		ApiDrawing.call(this, oChart);
		this.Chart = oChart;
	}

	ApiChart.prototype = Object.create(ApiDrawing.prototype);
	ApiChart.prototype.constructor = ApiChart;

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
	 * */

	/**
	 * Possible values for the position of chart tick labels (either horizontal or vertical).
	 * * <b>"none"</b> - does not display the selected tick labels.
	 * * <b>"nextTo"</b> - sets the position of the selected tick labels next to the main label.
	 * * <b>"low"</b> - sets the position of the selected tick labels in the part of the chart with lower values.
	 * * <b>"high"</b> - sets the position of the selected tick labels in the part of the chart with higher values.
	 * @typedef {("none" | "nextTo" | "low" | "high")} TickLabelPosition
	 * **/

	/**
	 * The page orientation type.
	 * @typedef {("xlLandscape" | "xlPortrait")} PageOrientation
	 * */

	/**
	 * The type of tick mark appearance.
	 * @typedef {("cross" | "in" | "none" | "out")} TickMark
	 * */

	/**
	 * Text transform type.
	 * @typedef {("textArchDown" | "textArchDownPour" | "textArchUp" | "textArchUpPour" | "textButton" | "textButtonPour" | "textCanDown"
	 * | "textCanUp" | "textCascadeDown" | "textCascadeUp" | "textChevron" | "textChevronInverted" | "textCircle" | "textCirclePour"
	 * | "textCurveDown" | "textCurveUp" | "textDeflate" | "textDeflateBottom" | "textDeflateInflate" | "textDeflateInflateDeflate" | "textDeflateTop"
	 * | "textDoubleWave1" | "textFadeDown" | "textFadeLeft" | "textFadeRight" | "textFadeUp" | "textInflate" | "textInflateBottom" | "textInflateTop"
	 * | "textPlain" | "textRingInside" | "textRingOutside" | "textSlantDown" | "textSlantUp" | "textStop" | "textTriangle" | "textTriangleInverted"
	 * | "textWave1" | "textWave2" | "textWave4" | "textNoShape")} TextTransform
	 * */

	/**
	 * Axis position in the chart.
	 * @typedef {("top" | "bottom" | "right" | "left")} AxisPos
	 */

	/**
	 * Standard numeric format.
	 * @typedef {("General" | "0" | "0.00" | "#,##0" | "#,##0.00" | "0%" | "0.00%" |
	 * "0.00E+00" | "# ?/?" | "# ??/??" | "m/d/yyyy" | "d-mmm-yy" | "d-mmm" | "mmm-yy" | "h:mm AM/PM" |
	 * "h:mm:ss AM/PM" | "h:mm" | "h:mm:ss" | "m/d/yyyy h:mm" | "#,##0_);(#,##0)" | "#,##0_);[Red](#,##0)" | 
	 * "#,##0.00_);(#,##0.00)" | "#,##0.00_);[Red](#,##0.00)" | "mm:ss" | "[h]:mm:ss" | "mm:ss.0" | "##0.0E+0" | "@")} NumFormat
	 */

	/**
	 * The cell references type.
	 * @typedef {('xlA1' | 'xlR1C1')} ReferenceStyle
	 * */

	//TODO not support "xlPasteAllMergingConditionalFormats" / "xlPasteAllUsingSourceTheme" / "xlPasteValidation"
	/**
	 * Specifies the part of the range to be pasted.
	 * @typedef {("xlPasteAll" | "xlPasteAllExceptBorders" |
	 *  | "xlPasteColumnWidths" | "xlPasteComments"
	 * | "xlPasteFormats" | "xlPasteFormulas" | "xlPasteFormulasAndNumberFormats"
	 * | "xlPasteValues" | "xlPasteValuesAndNumberFormats" )} PasteType
	 * */

	/**
	 * Specifies how numeric data will be calculated with the destinations cells on the worksheet.
	 * @typedef {("xlPasteSpecialOperationAdd" | "xlPasteSpecialOperationDivide" | "xlPasteSpecialOperationMultiply"|
	 * "xlPasteSpecialOperationNone" | "xlPasteSpecialOperationSubtract" )} PasteSpecialOperation
	 * */

	/**
	 * Class representing a base class for the color types.
	 * @constructor
	 */
	function ApiColor(color) {
		this.color = color;
	}

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
	 * @param {string} expression - Any valid expression.
	 * @param {string} [format] - A valid named or user-defined format expression.
	 * @returns {string}
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
		const parsedJSDoc = this.parsedJSDoc.shift();
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
	 * Creates a new worksheet. The new worksheet becomes the active sheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The name of a new worksheet.
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
	 */
	Api.prototype.SetLocale = function (LCID) {
		this.asc_setLocale(LCID, null, null);
	};

	/**
	 * Returns the current locale ID.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 */
	Api.prototype.GetLocale = function () {
		return this.asc_getLocale();
	};

	/**
	 * Returns an object that represents the active sheet.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiWorksheet}
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

	//vba
	/*let functionsMap = {"ACCRINT": "AccrInt","ACCRINTM": "AccrIntM","ACOS": "Acos","ACOSH": "Acosh","ACOT": "Acot","ACOTH": "Acoth","AGGREGATE": "Aggregate","AMORDEGRC": "AmorDegrc","AMORLINC": "AmorLinc","AND": "And","ARABIC": "Arabic","ASC": "Asc","ASIN": "Asin","ASINH": "Asinh","ATAN2": "Atan2","ATANH": "Atanh","AVEDEV": "AveDev","AVERAGE": "Average","AVERAGEIF": "AverageIf","AVERAGEIFS": "AverageIfs","BAHTTEXT": "BahtText","BASE": "Base","BESSELI": "BesselI","BESSELJ": "BesselJ","BESSELK": "BesselK","BESSELY": "BesselY","BETA.DIST": "Beta.Dist","BETA.INV": "Beta.Inv","BETADIST": "BetaDist","BETAINV": "BetaInv","BIN2DEC": "Bin2Dec","BIN2HEX": "Bin2Hex","BIN2OCT": "Bin2Oct","BINOM.DIST": "Binom.Dist","BINOM.DIST.RANGE": "Binom.Dist.Range","BINOM.INV": "Binom.Inv","BINOMDIST": "BinomDist","BITAND": "Bitand","BITLSHIFT": "Bitlshift","BITOR": "Bitor","BITRSHIFT": "Bitrshift","BITXOR": "Bitxor","CEILING": "Ceiling","CEILING.MATH": "Ceiling.Math","CEILING.PRECISE": "Ceiling.Precise","CHIDIST": "ChiDist","CHIINV": "ChiInv","CHISQ.DIST": "ChiSq.Dist","CHISQ.DIST.RT": "ChiSq.Dist.RT","CHISQ.INV": "ChiSq.Inv","CHISQ.INV.RT": "ChiSq.Inv.RT","CHISQ.TEST": "ChiSq.Test","CHITEST": "ChiTest","CHOOSE": "Choose","CLEAN": "Clean","COMBIN": "Combin","COMBINA": "Combina","COMPLEX": "Complex","CONFIDENCE": "Confidence","CONFIDENCE.NORM": "Confidence.Norm","CONFIDENCE.T": "Confidence.T","CONVERT": "Convert","CORREL": "Correl","COSH": "Cosh","COT": "Cot","COTH": "Coth","COUNT": "Count","COUNTA": "CountA","COUNTBLANK": "CountBlank","COUNTIF": "CountIf","COUNTIFS": "CountIfs","COUPDAYBS": "CoupDayBs","COUPDAYS": "CoupDays","COUPDAYSNC": "CoupDaysNc","COUPNCD": "CoupNcd","COUPNUM": "CoupNum","COUPPCD": "CoupPcd","COVAR": "Covar","COVARIANCE.P": "Covariance.P","COVARIANCE.S": "Covariance.S","CRITBINOM": "CritBinom","CSC": "Csc","CSCH": "Csch","CUMIPMT": "CumIPmt","CUMPRINC": "CumPrinc","DAVERAGE": "DAverage","DAYS": "Days","DAYS360": "Days360","DB": "Db","DBCS": "Dbcs","DCOUNT": "DCount","DCOUNTA": "DCountA","DDB": "Ddb","DEC2BIN": "Dec2Bin","DEC2HEX": "Dec2Hex","DEC2OCT": "Dec2Oct","DECIMAL": "Decimal","DEGREES": "Degrees","DELTA": "Delta","DEVSQ": "DevSq","DGET": "DGet","DISC": "Disc","DMAX": "DMax","DMIN": "DMin","DOLLAR": "Dollar","DOLLARDE": "DollarDe","DOLLARFR": "DollarFr","DPRODUCT": "DProduct","DSTDEV": "DStDev","DSTDEVP": "DStDevP","DSUM": "DSum","DURATION": "Duration","DVAR": "DVar","DVARP": "DVarP","EDATE": "EDate","EFFECT": "Effect","ENCODEURL": "EncodeUrl","EOMONTH": "EoMonth","ERF": "Erf","ERF.PRECISE": "Erf.Precise","ERFC": "ErfC","ERFC.PRECISE": "ErfC.Precise","EVEN": "Even","EXPON.DIST": "Expon.Dist","EXPONDIST": "ExponDist","F.DIST": "F.Dist","F.DIST.RT": "F.Dist.RT","F.INV": "F.Inv","F.INV.RT": "F.Inv.RT","F.TEST": "F.Test","FACT": "Fact","FACTDOUBLE": "FactDouble","FDIST": "FDist","FILTERXML": "FilterXML","FIND": "Find","FINDB": "FindB","FINV": "FInv","FISHER": "Fisher","FISHERINV": "FisherInv","FIXED": "Fixed","FLOOR": "Floor","FLOOR.MATH": "Floor.Math","FLOOR.PRECISE": "Floor.Precise","FORECAST": "Forecast","FORECAST.ETS": "Forecast.ETS","FORECAST.ETS.CONFINT": "Forecast.ETS.ConfInt","FORECAST.ETS.SEASONALITY": "Forecast.ETS.Seasonality","FORECAST.ETS.STAT": "Forecast.ETS.STAT","FORECAST.LINEAR": "Forecast.Linear","FREQUENCY": "Frequency","FTEST": "FTest","FV": "Fv","FVSCHEDULE": "FVSchedule","GAMMA": "Gamma","GAMMA.DIST": "Gamma.Dist","GAMMA.INV": "Gamma.Inv","GAMMADIST": "GammaDist","GAMMAINV": "GammaInv","GAMMALN": "GammaLn","GAMMALN.PRECISE": "GammaLn.Precise","GAUSS": "Gauss","GCD": "Gcd","GEOMEAN": "GeoMean","GESTEP": "GeStep","GROWTH": "Growth","HARMEAN": "HarMean","HEX2BIN": "Hex2Bin","HEX2DEC": "Hex2Dec","HEX2OCT": "Hex2Oct","HLOOKUP": "HLookup","HYPGEOM.DIST": "HypGeom.Dist","HYPGEOMDIST": "HypGeomDist","IFERROR": "IfError","IFNA": "IfNa","IMABS": "ImAbs","IMAGINARY": "Imaginary","IMARGUMENT": "ImArgument","IMCONJUGATE": "ImConjugate","IMCOS": "ImCos","IMCOSH": "ImCosh","IMCOT": "ImCot","IMCSC": "ImCsc","IMCSCH": "ImCsch","IMDIV": "ImDiv","IMEXP": "ImExp","IMLN": "ImLn","IMLOG10": "ImLog10","IMLOG2": "ImLog2","IMPOWER": "ImPower","IMPRODUCT": "ImProduct","IMREAL": "ImReal","IMSEC": "ImSec","IMSECH": "ImSech","IMSIN": "ImSin","IMSINH": "ImSinh","IMSQRT": "ImSqrt","IMSUB": "ImSub","IMSUM": "ImSum","IMTAN": "ImTan","INDEX": "Index","INTERCEPT": "Intercept","INTRATE": "IntRate","IPMT": "Ipmt","IRR": "Irr","ISERR": "IsErr","ISERROR": "IsError","ISEVEN": "IsEven","ISFORMULA": "IsFormula","ISLOGICAL": "IsLogical","ISNA": "IsNA","ISNONTEXT": "IsNonText","ISNUMBER": "IsNumber","ISO.CEILING": "ISO.Ceiling","ISODD": "IsOdd","ISOWEEKNUM": "IsoWeekNum","ISPMT": "Ispmt","ISTEXT": "IsText","KURT": "Kurt","LARGE": "Large","LCM": "Lcm","LINEST": "LinEst","LN": "Ln","LOG": "Log","LOG10": "Log10","LOGEST": "LogEst","LOGINV": "LogInv","LOGNORM.DIST": "LogNorm.Dist","LOGNORM.INV": "LogNorm.Inv","LOGNORMDIST": "LogNormDist","LOOKUP": "Lookup","MATCH": "Match","MAX": "Max","MDETERM": "MDeterm","MDURATION": "MDuration","MEDIAN": "Median","MIN": "Min","MINVERSE": "MInverse","MIRR": "MIrr","MMULT": "MMult","MODE": "Mode","MODE.MULT": "Mode.Mult","MODE.SNGL": "Mode.Sngl","MROUND": "MRound","MULTINOMIAL": "MultiNomial","MUNIT": "Munit","NEGBINOM.DIST": "NegBinom.Dist","NEGBINOMDIST": "NegBinomDist","NETWORKDAYS": "NetworkDays","NETWORKDAYS.INTL": "NetworkDays.Intl","NOMINAL": "Nominal","NORM.DIST": "Norm.Dist","NORM.INV": "Norm.Inv","NORM.S.DIST": "Norm.S.Dist","NORM.S.INV": "Norm.S.Inv","NORMDIST": "NormDist","NORMINV": "NormInv","NORMSDIST": "NormSDist","NORMSINV": "NormSInv","NPER": "NPer","NPV": "Npv","NUMBERVALUE": "NumberValue","OCT2BIN": "Oct2Bin","OCT2DEC": "Oct2Dec","OCT2HEX": "Oct2Hex","ODD": "Odd","ODDFPRICE": "OddFPrice","ODDFYIELD": "OddFYield","ODDLPRICE": "OddLPrice","ODDLYIELD": "OddLYield","OR": "Or","PDURATION": "PDuration","PEARSON": "Pearson","PERCENTILE": "Percentile","PERCENTILE.EXC": "Percentile.Exc","PERCENTILE.INC": "Percentile.Inc","PERCENTRANK": "PercentRank","PERCENTRANK.EXC": "PercentRank.Exc","PERCENTRANK.INC": "PercentRank.Inc","PERMUT": "Permut","PERMUTATIONA": "Permutationa","PHI": "Phi","PHONETIC": "Phonetic","PI": "Pi","PMT": "Pmt","POISSON": "Poisson","POISSON.DIST": "Poisson.Dist","POWER": "Power","PPMT": "Ppmt","PRICE": "Price","PRICEDISC": "PriceDisc","PRICEMAT": "PriceMat","PROB": "Prob","PRODUCT": "Product","PROPER": "Proper","PV": "Pv","QUARTILE": "Quartile","QUARTILE.EXC": "Quartile.Exc","QUARTILE.INC": "Quartile.Inc","QUOTIENT": "Quotient","RADIANS": "Radians","RANDBETWEEN": "RandBetween","RANK": "Rank","RANK.AVG": "Rank.Avg","RANK.EQ": "Rank.Eq","RATE": "Rate","RECEIVED": "Received","REPLACE": "Replace","REPLACEB": "ReplaceB","REPT": "Rept","ROMAN": "Roman","ROUND": "Round","ROUNDDOWN": "RoundDown","ROUNDUP": "RoundUp","RRI": "Rri","RSQ": "RSq","RTD": "RTD","SEARCH": "Search","SEARCHB": "SearchB","SEC": "Sec","SECH": "Sech","SERIESSUM": "SeriesSum","SINH": "Sinh","SKEW": "Skew","SKEW.P": "Skew.p","SLN": "Sln","SLOPE": "Slope","SMALL": "Small","SQRTPI": "SqrtPi","STANDARDIZE": "Standardize","STDEV": "StDev","STDEV.P": "StDev.P","STDEV.S": "StDev.S","STDEVP": "StDevP","STEYX": "StEyx","SUBSTITUTE": "Substitute","SUBTOTAL": "Subtotal","SUM": "Sum","SUMIF": "SumIf","SUMIFS": "SumIfs","SUMPRODUCT": "SumProduct","SUMSQ": "SumSq","SUMX2MY2": "SumX2MY2","SUMX2PY2": "SumX2PY2","SUMXMY2": "SumXMY2","SYD": "Syd","T.DIST": "T.Dist","T.DIST.2T": "T.Dist.2T","T.DIST.RT": "T.Dist.RT","T.INV": "T.Inv","T.INV.2T": "T.Inv.2T","T.TEST": "T.Test","TANH": "Tanh","TBILLEQ": "TBillEq","TBILLPRICE": "TBillPrice","TBILLYIELD": "TBillYield","TDIST": "TDist","TEXT": "Text","TINV": "TInv","TRANSPOSE": "Transpose","TREND": "Trend","TRIM": "Trim","TRIMMEAN": "TrimMean","TTEST": "TTest","UNICHAR": "Unichar","UNICODE": "Unicode","USDOLLAR": "USDollar","VAR": "Var","VAR.P": "Var.P","VAR.S": "Var.S","VARP": "VarP","VDB": "Vdb","VLOOKUP": "VLookup","WEBSERVICE": "WebService","WEEKDAY": "Weekday","WEEKNUM": "WeekNum","WEIBULL": "Weibull","WEIBULL.DIST": "Weibull.Dist","WORKDAY": "WorkDay","WORKDAY.INTL": "WorkDay.Intl","XIRR": "Xirr","XNPV": "Xnpv","XOR": "Xor","YEARFRAC": "YearFrac","YIELDDISC": "YieldDisc","YIELDMAT": "YieldMat","Z.TEST": "Z.Test","ZTEST": "ZTest"
	};*/

	//js
	/*let functionsMap2 = {"INIT": "init","ACCRINT": "accrInt","ACCRINTM": "accrIntM","ACOS": "acos","ACOSH": "acosh","ACOT": "acot","ACOTH": "acoth","AMORDEGRC": "amorDegrc","AMORLINC": "amorLinc","AND": "and","ARABIC": "arabic","AREAS": "areas","ASC": "asc","ASIN": "asin","ASINH": "asinh","ATAN": "atan","ATAN2": "atan2","ATANH": "atanh","AVEDEV": "aveDev","AVERAGE": "average","AVERAGEA": "averageA","AVERAGEIF": "averageIf","AVERAGEIFS": "averageIfs",//"BAHTTEXT": "bahtText","BASE": "base","BESSELI": "besselI","BESSELJ": "besselJ","BESSELK": "besselK","BESSELY": "besselY","BETA.DIST": "beta.Dist","BETA.INV": "beta.Inv","BIN2DEC": "bin2Dec","BIN2HEX": "bin2Hex","BIN2OCT": "bin2Oct","BINOM.DIST": "binom.Dist","BINOM.DIST.RANGE": "binom.Dist.Range","BINOM.INV": "binom.Inv","BITAND": "bitand","BITLSHIFT": "bitlshift","BITOR": "bitor","BITRSHIFT": "bitrshift","BITXOR": "bitxor","CEILING.MATH": "ceiling.Math","CEILING.PRECISE": "ceiling.Precise","CHAR": "char","CHISQ.DIST": "chiSq.Dist","CHISQ.DIST.RT": "chiSq.Dist.RT","CHISQ.INV": "chiSq.Inv","CHISQ.INV.RT": "chiSq.Inv.RT","CHOOSE": "choose","CLEAN": "clean","CODE": "code","COLUMNS": "columns","COMBIN": "combin","COMBINA": "combina","COMPLEX": "complex","CONCATENATE": "concatenate","CONFIDENCE.NORM": "confidence.Norm","CONFIDENCE.T": "confidence.T","CONVERT": "convert","COS": "cos","COSH": "cosh","COT": "cot","COTH": "coth","COUNT": "count","COUNTA": "countA","COUNTBLANK": "countBlank","COUNTIF": "countIf","COUNTIFS": "countIfs","COUPDAYBS": "coupDayBs","COUPDAYS": "coupDays","COUPDAYSNC": "coupDaysNc","COUPNCD": "coupNcd","COUPNUM": "coupNum","COUPPCD": "coupPcd","CSC": "csc","CSCH": "csch","CUMIPMT": "cumIPmt","CUMPRINC": "cumPrinc","DAVERAGE": "daverage","DCOUNT": "dcount","DCOUNTA": "dcountA","DGET": "dget","DMAX": "dmax","DMIN": "dmin","DPRODUCT": "dproduct","DSTDEV": "dstDev","DSTDEVP": "dstDevP","DSUM": "dsum","DVAR": "dvar","DVARP": "dvarP","DATE": "date","DATEVALUE": "datevalue","DAY": "day","DAYS": "days","DAYS360": "days360","DB": "db",//"DBCS": "dbcs","DDB": "ddb","DEC2BIN": "dec2Bin","DEC2HEX": "dec2Hex","DEC2OCT": "dec2Oct","DECIMAL": "decimal","DEGREES": "degrees","DELTA": "delta","DEVSQ": "devSq","DISC": "disc","DOLLAR": "dollar","DOLLARDE": "dollarDe","DOLLARFR": "dollarFr","DURATION": "duration","ECMA.CEILING": "ecma.Ceiling","EDATE": "edate","EFFECT": "effect","EOMONTH": "eoMonth","ERF": "erf","ERFC": "erfC","ERFC.PRECISE": "erfC.Precise","ERF.PRECISE": "erf.Precise","ERROR.TYPE": "error.Type","EVEN": "even","EXACT": "exact","EXP": "exp","EXPON.DIST": "expon.Dist","FVSCHEDULE": "fvschedule","F.DIST": "f.Dist","F.DIST.RT": "f.Dist.RT","F.INV": "f.Inv","F.INV.RT": "f.Inv.RT","FACT": "fact","FACTDOUBLE": "factDouble","FALSE": "false","FIND": "find","FINDB": "findB","FISHER": "fisher","FISHERINV": "fisherInv","FIXED": "fixed","FLOOR.MATH": "floor.Math","FLOOR.PRECISE": "floor.Precise","FV": "fv","GAMMA": "gamma","GAMMALN": "gammaLn","GAMMALN.PRECISE": "gammaLn.Precise","GAMMA.DIST": "gamma.Dist","GAMMA.INV": "gamma.Inv","GAUSS": "gauss","GCD": "gcd","GESTEP": "geStep","GEOMEAN": "geoMean","HLOOKUP": "hlookup","HARMEAN": "harMean","HEX2BIN": "hex2Bin","HEX2DEC": "hex2Dec","HEX2OCT": "hex2Oct","HOUR": "hour","HYPGEOM.DIST": "hypGeom.Dist","HYPERLINK": "hyperlink","ISO.CEILING": "iso.Ceiling","IF": "if","IMABS": "imAbs","IMARGUMENT": "imArgument","IMCONJUGATE": "imConjugate","IMCOS": "imCos","IMCOSH": "imCosh","IMCOT": "imCot","IMCSC": "imCsc","IMCSCH": "imCsch","IMDIV": "imDiv","IMEXP": "imExp","IMLN": "imLn","IMLOG10": "imLog10","IMLOG2": "imLog2","IMPOWER": "imPower","IMPRODUCT": "imProduct","IMREAL": "imReal","IMSEC": "imSec","IMSECH": "imSech","IMSIN": "imSin","IMSINH": "imSinh","IMSQRT": "imSqrt","IMSUB": "imSub","IMSUM": "imSum","IMTAN": "imTan","IMAGINARY": "imaginary","INT": "int","INTRATE": "intRate","IPMT": "ipmt","IRR": "irr","ISERR": "isErr","ISERROR": "isError","ISEVEN": "isEven","ISFORMULA": "isFormula","ISLOGICAL": "isLogical","ISNA": "isNA","ISNONTEXT": "isNonText","ISNUMBER": "isNumber","ISODD": "isOdd","ISTEXT": "isText","ISOWEEKNUM": "isoWeekNum","ISPMT": "ispmt","ISREF": "isref","KURT": "kurt","LARGE": "large","LCM": "lcm","LEFT": "left","LEFTB": "leftb","LEN": "len","LENB": "lenb","LN": "ln","LOG": "log","LOG10": "log10","LOGNORM.DIST": "logNorm.Dist","LOGNORM.INV": "logNorm.Inv","LOOKUP": "lookup","LOWER": "lower","MDURATION": "mduration","MIRR": "mirr","MROUND": "mround","MATCH": "match","MAX": "max","MAXA": "maxA","MEDIAN": "median","MID": "mid","MIDB": "midb","MIN": "min","MINA": "minA","MINUTE": "minute","MOD": "mod","MONTH": "month","MULTINOMIAL": "multiNomial","N": "n","NPER": "nper","NA": "na","NEGBINOM.DIST": "negBinom.Dist","NETWORKDAYS": "networkDays","NETWORKDAYS.INTL": "networkDays.Intl","NOMINAL": "nominal","NORM.DIST": "norm.Dist","NORM.INV": "norm.Inv","NORM.S.DIST": "norm.S.Dist","NORM.S.INV": "norm.S.Inv","NOT": "not","NOW": "now","NPV": "npv","NUMBERVALUE": "numberValue","OCT2BIN": "oct2Bin","OCT2DEC": "oct2Dec","OCT2HEX": "oct2Hex","ODD": "odd","ODDFPRICE": "oddFPrice","ODDFYIELD": "oddFYield","ODDLPRICE": "oddLPrice","ODDLYIELD": "oddLYield","OR": "or","PDURATION": "pduration","PERCENTRANK.EXC": "percentRank.Exc","PERCENTRANK.INC": "percentRank.Inc","PERCENTILE.EXC": "percentile.Exc","PERCENTILE.INC": "percentile.Inc","PERMUT": "permut","PERMUTATIONA": "permutationa","PHI": "phi","PI": "pi","PMT": "pmt","POISSON.DIST": "poisson.Dist","POWER": "power","PPMT": "ppmt","PRICE": "price","PRICEDISC": "priceDisc","PRICEMAT": "priceMat","PRODUCT": "product","PROPER": "proper","PV": "pv","QUARTILE.EXC": "quartile.Exc","QUARTILE.INC": "quartile.Inc","QUOTIENT": "quotient","RADIANS": "radians","RAND": "rand","RANDBETWEEN": "randBetween","RANK.AVG": "rank.Avg","RANK.EQ": "rank.Eq","RATE": "rate","RECEIVED": "received","REPLACE": "replace","REPLACEB": "replaceB","REPT": "rept","RIGHT": "right","RIGHTB": "rightb","ROMAN": "roman","ROUND": "round","ROUNDDOWN": "roundDown","ROUNDUP": "roundUp","ROWS": "rows","RRI": "rri","SEC": "sec","SECH": "sech","SECOND": "second","SERIESSUM": "seriesSum","SHEET": "sheet","SHEETS": "sheets","SIGN": "sign","SIN": "sin","SINH": "sinh","SKEW": "skew","SKEW.P": "skew.p","SLN": "sln","SMALL": "small","SQRT": "sqrt","SQRTPI": "sqrtPi","STDEVA": "stDevA","STDEVPA": "stDevPA","STDEV.P": "stDev.P","STDEV.S": "stDev.S","STANDARDIZE": "standardize","SUBSTITUTE": "substitute","SUBTOTAL": "subtotal","SUM": "sum","SUMIF": "sumIf","SUMIFS": "sumIfs","SUMSQ": "sumSq","SYD": "syd","T": "t","TBILLEQ": "tbillEq","TBILLPRICE": "tbillPrice","TBILLYIELD": "tbillYield","T.DIST": "t.Dist","T.DIST.2T": "t.Dist.2T","T.DIST.RT": "t.Dist.RT","T.INV": "t.Inv","T.INV.2T": "t.Inv.2T","TAN": "tan","TANH": "tanh","TEXT": "text","TIME": "time","TIMEVALUE": "timevalue","TODAY": "today","TRIM": "trim","TRIMMEAN": "trimMean","TRUE": "true","TRUNC": "trunc","TYPE": "type","USDOLLAR": "usdollar","UNICHAR": "unichar","UNICODE": "unicode","UPPER": "upper","VLOOKUP": "vlookup","VALUE": "value","VARA": "varA","VARPA": "varPA","VAR.P": "var.P","VAR.S": "var.S","VDB": "vdb","WEEKNUM": "weekNum","WEEKDAY": "weekday","WEIBULL.DIST": "weibull.Dist","WORKDAY": "workDay","WORKDAY.INTL": "workDay.Intl","XIRR": "xirr","XNPV": "xnpv","XOR": "xor","YEAR": "year","YEARFRAC": "yearFrac","YIELD": "yield","YIELDDISC": "yieldDisc","YIELDMAT": "yieldMat","Z.TEST": "z.Test"
	};*/

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
	// 				let maxArg = AscCommonExcel.cFormulaFunction[i].prototype.argumentsMax
	// 				let minArg = AscCommonExcel.cFormulaFunction[i].prototype.argumentsMin
	// 				if (maxArg < 10) {
	// 					for (let j = 1; j <= maxArg; j++) {
	// 						let test12 = "arg" + (j)
	//
	// 						test1 += "\t * @param {" + getArgType(AscCommonExcel.cFormulaFunction[i].prototype.argumentsType && AscCommonExcel.cFormulaFunction[i].prototype.argumentsType[j-1]) + "} ";
	// 						if (j <= minArg) {
	// 							test1 += test12;
	// 						} else {
	// 							test1 += "[" + test12 + "]";
	// 						}
	// 						test1 +=  ".\n";
	//
	// 						test2 += j === maxArg ? test12 : (test12 + ",")
	// 					}
	// 				}
	//
	//
	// 				test += "/**\n" + "\t * Returns the result of calculating the function.\n" + "\t * @memberof ApiWorksheetFunction\n" + "\t * @typeofeditors [\"CSE\"]\n" + test1
	// 					+ "\t * @returns {number | string | boolean}\n" + "\t */\n" + "\tApiWorksheetFunction.prototype." + i.replaceAll(".","_")  + "= function (" + test2 + ") {\n" + "\t\tthis.private_calculateFunction(\"" + i + "\", arguments);\n" + "\t};"
	// 			//}
	//
	// 			test += "\n"
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
			if (cElementType.empty === result.type) {
				result = new cNumber(0);
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
	}

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
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ASC = function (arg1) {
		return this.private_calculateFunction("ASC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHAR = function (arg1) {
		return this.private_calculateFunction("CHAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CLEAN = function (arg1) {
		return this.private_calculateFunction("CLEAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CODE = function (arg1) {
		return this.private_calculateFunction("CODE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CONCATENATE = function () {
		return this.private_calculateFunction("CONCATENATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DOLLAR = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EXACT = function (arg1, arg2) {
		return this.private_calculateFunction("EXACT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FIND = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FIND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FINDB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FINDB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @param {boolean} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FIXED = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FIXED", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LEFT = function (arg1, arg2) {
		return this.private_calculateFunction("LEFT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LEFTB = function (arg1, arg2) {
		return this.private_calculateFunction("LEFTB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LEN = function (arg1) {
		return this.private_calculateFunction("LEN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LENB = function (arg1) {
		return this.private_calculateFunction("LENB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOWER = function (arg1) {
		return this.private_calculateFunction("LOWER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MID = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MID", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MIDB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MIDB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} [arg2].
	 * @param {string} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NUMBERVALUE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NUMBERVALUE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PROPER = function (arg1) {
		return this.private_calculateFunction("PROPER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {string} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.REPLACE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("REPLACE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {string} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.REPLACEB = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("REPLACEB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.REPT = function (arg1, arg2) {
		return this.private_calculateFunction("REPT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RIGHT = function (arg1, arg2) {
		return this.private_calculateFunction("RIGHT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RIGHTB = function (arg1, arg2) {
		return this.private_calculateFunction("RIGHTB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SEARCH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SEARCH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SEARCHB = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SEARCHB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {string} arg2.
	 * @param {string} arg3.
	 * @param {string} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUBSTITUTE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SUBSTITUTE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T = function (arg1) {
		return this.private_calculateFunction("T", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {string} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TEXT = function (arg1, arg2) {
		return this.private_calculateFunction("TEXT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TRIM = function (arg1) {
		return this.private_calculateFunction("TRIM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.UNICHAR = function (arg1) {
		return this.private_calculateFunction("UNICHAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.UNICODE = function (arg1) {
		return this.private_calculateFunction("UNICODE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.UPPER = function (arg1) {
		return this.private_calculateFunction("UPPER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VALUE = function (arg1) {
		return this.private_calculateFunction("VALUE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AVEDEV = function () {
		return this.private_calculateFunction("AVEDEV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AVERAGE = function () {
		return this.private_calculateFunction("AVERAGE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AVERAGEA = function () {
		return this.private_calculateFunction("AVERAGEA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {any} arg2.
	 * @param {ApiRange} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AVERAGEIF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("AVERAGEIF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AVERAGEIFS = function () {
		return this.private_calculateFunction("AVERAGEIFS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BETADIST = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETADIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BETA_DIST = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("BETA.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BETA_INV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETA.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BETAINV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("BETAINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BINOMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOMDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BINOM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOM.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BINOM_DIST.RANGE = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("BINOM.DIST.RANGE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BINOM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("BINOM.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHIDIST = function (arg1, arg2) {
		return this.private_calculateFunction("CHIDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHIINV = function (arg1, arg2) {
		return this.private_calculateFunction("CHIINV", arguments);
	};

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHISQ_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CHISQ.DIST", arguments);
	};

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHISQ_DIST_RT = function (arg1, arg2) {
		return this.private_calculateFunction("CHISQ.DIST.RT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHISQ_INV = function (arg1, arg2) {
		return this.private_calculateFunction("CHISQ.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHISQ_INV.RT = function (arg1, arg2) {
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

	//todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.CHISQ_TEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("CHISQ.TEST", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE_NORM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE.NORM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CONFIDENCE_T = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONFIDENCE.T", arguments);
	};

	// //todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.CORREL = function (arg1, arg2) {
	// 	return this.private_calculateFunction("CORREL", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUNT = function () {
		return this.private_calculateFunction("COUNT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUNTA = function () {
		return this.private_calculateFunction("COUNTA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUNTBLANK = function (arg1) {
		return this.private_calculateFunction("COUNTBLANK", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUNTIF = function (arg1, arg2) {
		return this.private_calculateFunction("COUNTIF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUNTIFS = function () {
		return this.private_calculateFunction("COUNTIFS", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVAR = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVAR", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVARIANCE_P = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVARIANCE.P", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.COVARIANCE_S = function (arg1, arg2) {
	// 	return this.private_calculateFunction("COVARIANCE.S", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CRITBINOM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CRITBINOM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DEVSQ = function () {
		return this.private_calculateFunction("DEVSQ", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EXPON_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("EXPON.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EXPONDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("EXPONDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.F_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("F.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.F_DIST.RT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.DIST.RT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.F_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.F_INV.RT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("F.INV.RT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FISHER = function (arg1) {
		return this.private_calculateFunction("FISHER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FISHERINV = function (arg1) {
		return this.private_calculateFunction("FISHERINV", arguments);
	};

	//todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {number} arg1.
	//  * @param {any} arg2.
	//  * @param {any} arg3.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FORECAST = function (arg1, arg2, arg3) {
	// 	return this.private_calculateFunction("FORECAST", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {ApiRange} arg2.
	 * @param {ApiRange} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("FORECAST.ETS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {ApiRange} arg2.
	 * @param {ApiRange} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @param {number} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS.CONFINT = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("FORECAST.ETS.CONFINT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} arg2.
	 * @param {number} [arg3].
	 * @param {number} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS.SEASONALITY = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("FORECAST.ETS.SEASONALITY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FORECAST_ETS.STAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("FORECAST.ETS.STAT", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {number} arg1.
	//  * @param {any} arg2.
	//  * @param {any} arg3.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FORECAST_LINEAR = function (arg1, arg2, arg3) {
	// 	return this.private_calculateFunction("FORECAST.LINEAR", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FREQUENCY = function (arg1, arg2) {
		return this.private_calculateFunction("FREQUENCY", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.FTEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("FTEST", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.F_TEST = function (arg1, arg2) {
	// 	return this.private_calculateFunction("F.TEST", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMA = function (arg1) {
		return this.private_calculateFunction("GAMMA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMA_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GAMMA.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMADIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GAMMADIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMA_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("GAMMA.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMAINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("GAMMAINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMALN = function (arg1) {
		return this.private_calculateFunction("GAMMALN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAMMALN_PRECISE = function (arg1) {
		return this.private_calculateFunction("GAMMALN.PRECISE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GAUSS = function (arg1) {
		return this.private_calculateFunction("GAUSS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GEOMEAN = function () {
		return this.private_calculateFunction("GEOMEAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} [arg2].
	 * @param {ApiRange} [arg3].
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GROWTH = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("GROWTH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HARMEAN = function () {
		return this.private_calculateFunction("HARMEAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HYPGEOMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("HYPGEOMDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {boolean} arg5.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HYPGEOM_DIST = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("HYPGEOM.DIST", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.INTERCEPT = function (arg1, arg2) {
	// 	return this.private_calculateFunction("INTERCEPT", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.KURT = function () {
		return this.private_calculateFunction("KURT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LARGE = function (arg1, arg2) {
		return this.private_calculateFunction("LARGE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} [arg2].
	 * @param {boolean} [arg3].
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LINEST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LINEST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} [arg2].
	 * @param {boolean} [arg3].
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOGEST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LOGEST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOGINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOGNORM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("LOGNORM.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOGNORM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGNORM.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOGNORMDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOGNORMDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MAX = function () {
		return this.private_calculateFunction("MAX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MAXA = function () {
		return this.private_calculateFunction("MAXA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MEDIAN = function () {
		return this.private_calculateFunction("MEDIAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MIN = function () {
		return this.private_calculateFunction("MIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MINA = function () {
		return this.private_calculateFunction("MINA", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE = function () {
	// 	return this.private_calculateFunction("MODE", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE_MULT = function () {
	// 	return this.private_calculateFunction("MODE.MULT", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MODE_SNGL = function () {
	// 	return this.private_calculateFunction("MODE.SNGL", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NEGBINOMDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NEGBINOMDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NEGBINOM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NEGBINOM.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORMDIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NORMDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORM_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NORM.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORMINV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NORMINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORM_INV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NORM.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORMSDIST = function (arg1) {
		return this.private_calculateFunction("NORMSDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {boolean} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORM_S_DIST = function (arg1, arg2) {
		return this.private_calculateFunction("NORM.S.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORMSINV = function (arg1) {
		return this.private_calculateFunction("NORMSINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NORM_S_INV = function (arg1) {
		return this.private_calculateFunction("NORM.S.INV", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.PEARSON = function (arg1, arg2) {
	// 	return this.private_calculateFunction("PEARSON", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTILE = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTILE_EXC = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE.EXC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTILE_INC = function (arg1, arg2) {
		return this.private_calculateFunction("PERCENTILE.INC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK_EXC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK.EXC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERCENTRANK_INC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PERCENTRANK.INC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERMUT = function (arg1, arg2) {
		return this.private_calculateFunction("PERMUT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PERMUTATIONA = function (arg1, arg2) {
		return this.private_calculateFunction("PERMUTATIONA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PHI = function (arg1) {
		return this.private_calculateFunction("PHI", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.POISSON = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("POISSON", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.POISSON_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("POISSON.DIST", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @param {number} arg3.
	//  * @param {number} [arg4].
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.PROB = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("PROB", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.QUARTILE = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.QUARTILE_EXC = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE.EXC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.QUARTILE_INC = function (arg1, arg2) {
		return this.private_calculateFunction("QUARTILE.INC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {ApiRange} arg2.
	 * @param {boolean} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RANK = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {ApiRange} arg2.
	 * @param {boolean} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RANK_AVG = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK.AVG", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {ApiRange} arg2.
	 * @param {boolean} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RANK_EQ = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RANK.EQ", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.RSQ = function (arg1, arg2) {
	// 	return this.private_calculateFunction("RSQ", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SKEW = function () {
		return this.private_calculateFunction("SKEW", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SKEW_P = function () {
		return this.private_calculateFunction("SKEW.P", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SLOPE = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SLOPE", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SMALL = function (arg1, arg2) {
		return this.private_calculateFunction("SMALL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STANDARDIZE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("STANDARDIZE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEV = function () {
		return this.private_calculateFunction("STDEV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEV_S = function () {
		return this.private_calculateFunction("STDEV.S", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEVA = function () {
		return this.private_calculateFunction("STDEVA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEVP = function () {
		return this.private_calculateFunction("STDEVP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEV_P = function () {
		return this.private_calculateFunction("STDEV.P", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.STDEVPA = function () {
		return this.private_calculateFunction("STDEVPA", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.STEYX = function (arg1, arg2) {
	// 	return this.private_calculateFunction("STEYX", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TDIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TDIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T_DIST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("T.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T_DIST_2T = function (arg1, arg2) {
		return this.private_calculateFunction("T.DIST.2T", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T_DIST_RT = function (arg1, arg2) {
		return this.private_calculateFunction("T.DIST.RT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T_INV = function (arg1, arg2) {
		return this.private_calculateFunction("T.INV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.T_INV_2T = function (arg1, arg2) {
		return this.private_calculateFunction("T.INV.2T", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TINV = function (arg1, arg2) {
		return this.private_calculateFunction("TINV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {ApiRange} [arg2].
	 * @param {ApiRange} [arg3].
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TREND = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("TREND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TRIMMEAN = function (arg1, arg2) {
		return this.private_calculateFunction("TRIMMEAN", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @param {number} arg3.
	//  * @param {number} arg4.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.TTEST = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("TTEST", arguments);
	// };

	//todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @param {number} arg3.
	//  * @param {number} arg4.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.T_TEST = function (arg1, arg2, arg3, arg4) {
	// 	return this.private_calculateFunction("T.TEST", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VAR = function () {
		return this.private_calculateFunction("VAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VARA = function () {
		return this.private_calculateFunction("VARA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VARP = function () {
		return this.private_calculateFunction("VARP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VAR_P = function () {
		return this.private_calculateFunction("VAR.P", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VAR_S = function () {
		return this.private_calculateFunction("VAR.S", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VARPA = function () {
		return this.private_calculateFunction("VARPA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WEIBULL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WEIBULL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WEIBULL_DIST = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WEIBULL.DIST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ZTEST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("ZTEST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.Z_TEST = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("Z.TEST", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DATE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DATEVALUE = function (arg1) {
		return this.private_calculateFunction("DATEVALUE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DAY = function (arg1) {
		return this.private_calculateFunction("DAY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DAYS = function (arg1, arg2) {
		return this.private_calculateFunction("DAYS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {boolean} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DAYS360 = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DAYS360", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EDATE = function (arg1, arg2) {
		return this.private_calculateFunction("EDATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EOMONTH = function (arg1, arg2) {
		return this.private_calculateFunction("EOMONTH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HOUR = function (arg1) {
		return this.private_calculateFunction("HOUR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISOWEEKNUM = function (arg1) {
		return this.private_calculateFunction("ISOWEEKNUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MINUTE = function (arg1) {
		return this.private_calculateFunction("MINUTE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MONTH = function (arg1) {
		return this.private_calculateFunction("MONTH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NETWORKDAYS = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("NETWORKDAYS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {number} [arg3].
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NETWORKDAYS_INTL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("NETWORKDAYS.INTL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NOW = function () {
		return this.private_calculateFunction("NOW", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SECOND = function (arg1) {
		return this.private_calculateFunction("SECOND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TIME = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TIME", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TIMEVALUE = function (arg1) {
		return this.private_calculateFunction("TIMEVALUE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TODAY = function () {
		return this.private_calculateFunction("TODAY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WEEKDAY = function (arg1, arg2) {
		return this.private_calculateFunction("WEEKDAY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WEEKNUM = function (arg1, arg2) {
		return this.private_calculateFunction("WEEKNUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WORKDAY = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("WORKDAY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {number} [arg3].
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.WORKDAY_INTL = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("WORKDAY.INTL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.YEAR = function (arg1) {
		return this.private_calculateFunction("YEAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.YEARFRAC = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("YEARFRAC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BESSELI = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELI", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BESSELJ = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELJ", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BESSELK = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELK", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BESSELY = function (arg1, arg2) {
		return this.private_calculateFunction("BESSELY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BIN2DEC = function (arg1) {
		return this.private_calculateFunction("BIN2DEC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BIN2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("BIN2HEX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BIN2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("BIN2OCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BITAND = function (arg1, arg2) {
		return this.private_calculateFunction("BITAND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BITLSHIFT = function (arg1, arg2) {
		return this.private_calculateFunction("BITLSHIFT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BITOR = function (arg1, arg2) {
		return this.private_calculateFunction("BITOR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BITRSHIFT = function (arg1, arg2) {
		return this.private_calculateFunction("BITRSHIFT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BITXOR = function (arg1, arg2) {
		return this.private_calculateFunction("BITXOR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COMPLEX = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("COMPLEX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CONVERT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CONVERT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DEC2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2BIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DEC2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2HEX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DEC2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("DEC2OCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DELTA = function (arg1, arg2) {
		return this.private_calculateFunction("DELTA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ERF = function (arg1, arg2) {
		return this.private_calculateFunction("ERF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ERF_PRECISE = function (arg1) {
		return this.private_calculateFunction("ERF.PRECISE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ERFC = function (arg1) {
		return this.private_calculateFunction("ERFC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ERFC_PRECISE = function (arg1) {
		return this.private_calculateFunction("ERFC.PRECISE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GESTEP = function (arg1, arg2) {
		return this.private_calculateFunction("GESTEP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HEX2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("HEX2BIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HEX2DEC = function (arg1) {
		return this.private_calculateFunction("HEX2DEC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HEX2OCT = function (arg1, arg2) {
		return this.private_calculateFunction("HEX2OCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMABS = function (arg1) {
		return this.private_calculateFunction("IMABS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMAGINARY = function (arg1) {
		return this.private_calculateFunction("IMAGINARY", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMARGUMENT = function (arg1) {
		return this.private_calculateFunction("IMARGUMENT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCONJUGATE = function (arg1) {
		return this.private_calculateFunction("IMCONJUGATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCOS = function (arg1) {
		return this.private_calculateFunction("IMCOS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCOSH = function (arg1) {
		return this.private_calculateFunction("IMCOSH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCOT = function (arg1) {
		return this.private_calculateFunction("IMCOT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCSC = function (arg1) {
		return this.private_calculateFunction("IMCSC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMCSCH = function (arg1) {
		return this.private_calculateFunction("IMCSCH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMDIV = function (arg1, arg2) {
		return this.private_calculateFunction("IMDIV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMEXP = function (arg1) {
		return this.private_calculateFunction("IMEXP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMLN = function (arg1) {
		return this.private_calculateFunction("IMLN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMLOG10 = function (arg1) {
		return this.private_calculateFunction("IMLOG10", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMLOG2 = function (arg1) {
		return this.private_calculateFunction("IMLOG2", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMPOWER = function (arg1, arg2) {
		return this.private_calculateFunction("IMPOWER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMPRODUCT = function () {
		return this.private_calculateFunction("IMPRODUCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMREAL = function (arg1) {
		return this.private_calculateFunction("IMREAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSEC = function (arg1) {
		return this.private_calculateFunction("IMSEC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSECH = function (arg1) {
		return this.private_calculateFunction("IMSECH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSIN = function (arg1) {
		return this.private_calculateFunction("IMSIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSINH = function (arg1) {
		return this.private_calculateFunction("IMSINH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSQRT = function (arg1) {
		return this.private_calculateFunction("IMSQRT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSUB = function (arg1, arg2) {
		return this.private_calculateFunction("IMSUB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMSUM = function () {
		return this.private_calculateFunction("IMSUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IMTAN = function (arg1) {
		return this.private_calculateFunction("IMTAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.OCT2BIN = function (arg1, arg2) {
		return this.private_calculateFunction("OCT2BIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.OCT2DEC = function (arg1) {
		return this.private_calculateFunction("OCT2DEC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.OCT2HEX = function (arg1, arg2) {
		return this.private_calculateFunction("OCT2HEX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DAVERAGE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DAVERAGE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DCOUNT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DCOUNT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DCOUNTA = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DCOUNTA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DGET = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DGET", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DMAX = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DMAX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DMIN = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DMIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DPRODUCT = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DPRODUCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DSTDEV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSTDEV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DSTDEVP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSTDEVP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DSUM = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DSUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DVAR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DVAR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {string} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DVARP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("DVARP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} [arg7].
	 * @param {any} [arg8].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACCRINT = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ACCRINT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACCRINTM = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("ACCRINTM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AMORDEGRC = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("AMORDEGRC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AMORLINC = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("AMORLINC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPDAYBS = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYBS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPDAYS = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPDAYSNC = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPDAYSNC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPNCD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPNCD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPNUM = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPNUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COUPPCD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("COUPPCD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CUMIPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("CUMIPMT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CUMPRINC = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("CUMPRINC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DB = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DDB = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DDB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("DISC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DOLLARDE = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLARDE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DOLLARFR = function (arg1, arg2) {
		return this.private_calculateFunction("DOLLARFR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DURATION = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("DURATION", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EFFECT = function (arg1, arg2) {
		return this.private_calculateFunction("EFFECT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("FV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FVSCHEDULE = function (arg1, arg2) {
		return this.private_calculateFunction("FVSCHEDULE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.INTRATE = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("INTRATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("IPMT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IRR = function (arg1, arg2) {
		return this.private_calculateFunction("IRR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISPMT = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("ISPMT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MDURATION = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("MDURATION", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MIRR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MIRR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NOMINAL = function (arg1, arg2) {
		return this.private_calculateFunction("NOMINAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NPER = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("NPER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NPV = function () {
		return this.private_calculateFunction("NPV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} arg7.
	 * @param {any} arg8.
	 * @param {any} [arg9].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ODDFPRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9) {
		return this.private_calculateFunction("ODDFPRICE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} arg7.
	 * @param {any} arg8.
	 * @param {any} [arg9].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ODDFYIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9) {
		return this.private_calculateFunction("ODDFYIELD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} arg7.
	 * @param {any} [arg8].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ODDLPRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ODDLPRICE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} arg7.
	 * @param {any} [arg8].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ODDLYIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) {
		return this.private_calculateFunction("ODDLYIELD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PDURATION = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("PDURATION", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PMT = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PMT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PPMT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("PPMT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PRICE = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("PRICE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PRICEDISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PRICEDISC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PRICEMAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("PRICEMAT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PV = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("PV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} [arg4].
	 * @param {number} [arg5].
	 * @param {number} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RATE = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("RATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RECEIVED = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("RECEIVED", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RRI = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("RRI", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SLN = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SLN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SYD = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SYD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TBILLEQ = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLEQ", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TBILLPRICE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLPRICE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TBILLYIELD = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("TBILLYIELD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {number} arg4.
	 * @param {number} arg5.
	 * @param {number} [arg6].
	 * @param {boolean} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VDB = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("VDB", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.XIRR = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("XIRR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.XNPV = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("XNPV", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} arg6.
	 * @param {any} [arg7].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.YIELD = function (arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		return this.private_calculateFunction("YIELD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} [arg5].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.YIELDDISC = function (arg1, arg2, arg3, arg4, arg5) {
		return this.private_calculateFunction("YIELDDISC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @param {any} arg5.
	 * @param {any} [arg6].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.YIELDMAT = function (arg1, arg2, arg3, arg4, arg5, arg6) {
		return this.private_calculateFunction("YIELDMAT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ABS = function (arg1) {
		return this.private_calculateFunction("ABS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACOS = function (arg1) {
		return this.private_calculateFunction("ACOS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACOSH = function (arg1) {
		return this.private_calculateFunction("ACOSH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACOT = function (arg1) {
		return this.private_calculateFunction("ACOT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ACOTH = function (arg1) {
		return this.private_calculateFunction("ACOTH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AGGREGATE = function () {
		return this.private_calculateFunction("AGGREGATE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ARABIC = function (arg1) {
		return this.private_calculateFunction("ARABIC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ASIN = function (arg1) {
		return this.private_calculateFunction("ASIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ASINH = function (arg1) {
		return this.private_calculateFunction("ASINH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ATAN = function (arg1) {
		return this.private_calculateFunction("ATAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ATAN2 = function (arg1, arg2) {
		return this.private_calculateFunction("ATAN2", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ATANH = function (arg1) {
		return this.private_calculateFunction("ATANH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.BASE = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("BASE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("CEILING", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CEILING_MATH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("CEILING.MATH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CEILING_PRECISE = function (arg1, arg2) {
		return this.private_calculateFunction("CEILING.PRECISE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COMBIN = function (arg1, arg2) {
		return this.private_calculateFunction("COMBIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COMBINA = function (arg1, arg2) {
		return this.private_calculateFunction("COMBINA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COS = function (arg1) {
		return this.private_calculateFunction("COS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COSH = function (arg1) {
		return this.private_calculateFunction("COSH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COT = function (arg1) {
		return this.private_calculateFunction("COT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COTH = function (arg1) {
		return this.private_calculateFunction("COTH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CSC = function (arg1) {
		return this.private_calculateFunction("CSC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CSCH = function (arg1) {
		return this.private_calculateFunction("CSCH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DECIMAL = function (arg1, arg2) {
		return this.private_calculateFunction("DECIMAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.DEGREES = function (arg1) {
		return this.private_calculateFunction("DEGREES", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ECMA_CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("ECMA.CEILING", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EVEN = function (arg1) {
		return this.private_calculateFunction("EVEN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.EXP = function (arg1) {
		return this.private_calculateFunction("EXP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FACT = function (arg1) {
		return this.private_calculateFunction("FACT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FACTDOUBLE = function (arg1) {
		return this.private_calculateFunction("FACTDOUBLE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FLOOR = function (arg1, arg2) {
		return this.private_calculateFunction("FLOOR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FLOOR_PRECISE = function (arg1, arg2) {
		return this.private_calculateFunction("FLOOR.PRECISE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FLOOR_MATH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("FLOOR.MATH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.GCD = function () {
		return this.private_calculateFunction("GCD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.INT = function (arg1) {
		return this.private_calculateFunction("INT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISO_CEILING = function (arg1, arg2) {
		return this.private_calculateFunction("ISO.CEILING", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LCM = function () {
		return this.private_calculateFunction("LCM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LN = function (arg1) {
		return this.private_calculateFunction("LN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOG = function (arg1, arg2) {
		return this.private_calculateFunction("LOG", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOG10 = function (arg1) {
		return this.private_calculateFunction("LOG10", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MDETERM = function (arg1) {
	// 	return this.private_calculateFunction("MDETERM", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MINVERSE = function (arg1) {
	// 	return this.private_calculateFunction("MINVERSE", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.MMULT = function (arg1, arg2) {
	// 	return this.private_calculateFunction("MMULT", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MOD = function (arg1, arg2) {
		return this.private_calculateFunction("MOD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MROUND = function (arg1, arg2) {
		return this.private_calculateFunction("MROUND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MULTINOMIAL = function () {
		return this.private_calculateFunction("MULTINOMIAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MUNIT = function (arg1) {
		return this.private_calculateFunction("MUNIT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ODD = function (arg1) {
		return this.private_calculateFunction("ODD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PI = function () {
		return this.private_calculateFunction("PI", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.POWER = function (arg1, arg2) {
		return this.private_calculateFunction("POWER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.PRODUCT = function () {
		return this.private_calculateFunction("PRODUCT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.QUOTIENT = function (arg1, arg2) {
		return this.private_calculateFunction("QUOTIENT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RADIANS = function (arg1) {
		return this.private_calculateFunction("RADIANS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RAND = function () {
		return this.private_calculateFunction("RAND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.RANDBETWEEN = function (arg1, arg2) {
		return this.private_calculateFunction("RANDBETWEEN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ROMAN = function (arg1, arg2) {
		return this.private_calculateFunction("ROMAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ROUND = function (arg1, arg2) {
		return this.private_calculateFunction("ROUND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ROUNDDOWN = function (arg1, arg2) {
		return this.private_calculateFunction("ROUNDDOWN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ROUNDUP = function (arg1, arg2) {
		return this.private_calculateFunction("ROUNDUP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SEC = function (arg1) {
		return this.private_calculateFunction("SEC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SECH = function (arg1) {
		return this.private_calculateFunction("SECH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @param {any} arg3.
	 * @param {any} arg4.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SERIESSUM = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("SERIESSUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SIGN = function (arg1) {
		return this.private_calculateFunction("SIGN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SIN = function (arg1) {
		return this.private_calculateFunction("SIN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SINH = function (arg1) {
		return this.private_calculateFunction("SINH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SQRT = function (arg1) {
		return this.private_calculateFunction("SQRT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SQRTPI = function (arg1) {
		return this.private_calculateFunction("SQRTPI", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUBTOTAL = function () {
		return this.private_calculateFunction("SUBTOTAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUM = function () {
		return this.private_calculateFunction("SUM", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {any} arg2.
	 * @param {ApiRange} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUMIF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("SUMIF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUMIFS = function () {
		return this.private_calculateFunction("SUMIFS", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMPRODUCT = function () {
	// 	return this.private_calculateFunction("SUMPRODUCT", arguments);
	// };

	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SUMSQ = function () {
		return this.private_calculateFunction("SUMSQ", arguments);
	};

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMX2MY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMX2MY2", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMX2PY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMX2PY2", arguments);
	// };

	// todo need array
	// /**
	//  * Returns the result of calculating the function.
	//  * @memberof ApiWorksheetFunction
	//  * @typeofeditors ["CSE"]
	//  * @param {any} arg1.
	//  * @param {any} arg2.
	//  * @returns {number | string | boolean}
	//  */
	// ApiWorksheetFunction.prototype.SUMXMY2 = function (arg1, arg2) {
	// 	return this.private_calculateFunction("SUMXMY2", arguments);
	// };


	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TAN = function (arg1) {
		return this.private_calculateFunction("TAN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TANH = function (arg1) {
		return this.private_calculateFunction("TANH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {number} arg1.
	 * @param {number} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TRUNC = function (arg1, arg2) {
		return this.private_calculateFunction("TRUNC", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.CHOOSE = function () {
		return this.private_calculateFunction("CHOOSE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.COLUMNS = function (arg1) {
		return this.private_calculateFunction("COLUMNS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HLOOKUP = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("HLOOKUP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} arg1.
	 * @param {any} [arg2].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.HYPERLINK = function (arg1, arg2) {
		return this.private_calculateFunction("HYPERLINK", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @param {any} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.INDEX = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("INDEX", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {ApiRange} arg2.
	 * @param {ApiRange} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.LOOKUP = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("LOOKUP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {number} arg2.
	 * @param {number} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.MATCH = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("MATCH", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ROWS = function (arg1) {
		return this.private_calculateFunction("ROWS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TRANSPOSE = function (arg1) {
		return this.private_calculateFunction("TRANSPOSE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {number} arg2.
	 * @param {number} arg3.
	 * @param {boolean} [arg4].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.VLOOKUP = function (arg1, arg2, arg3, arg4) {
		return this.private_calculateFunction("VLOOKUP", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ERROR_TYPE = function (arg1) {
		return this.private_calculateFunction("ERROR.TYPE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISERR = function (arg1) {
		return this.private_calculateFunction("ISERR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISERROR = function (arg1) {
		return this.private_calculateFunction("ISERROR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISEVEN = function (arg1) {
		return this.private_calculateFunction("ISEVEN", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISFORMULA = function (arg1) {
		return this.private_calculateFunction("ISFORMULA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISLOGICAL = function (arg1) {
		return this.private_calculateFunction("ISLOGICAL", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISNA = function (arg1) {
		return this.private_calculateFunction("ISNA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISNONTEXT = function (arg1) {
		return this.private_calculateFunction("ISNONTEXT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISNUMBER = function (arg1) {
		return this.private_calculateFunction("ISNUMBER", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISODD = function (arg1) {
		return this.private_calculateFunction("ISODD", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISREF = function (arg1) {
		return this.private_calculateFunction("ISREF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.ISTEXT = function (arg1) {
		return this.private_calculateFunction("ISTEXT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.N = function (arg1) {
		return this.private_calculateFunction("N", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NA = function () {
		return this.private_calculateFunction("NA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {string} [arg1].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SHEET = function (arg1) {
		return this.private_calculateFunction("SHEET", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} [arg1].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.SHEETS = function (arg1) {
		return this.private_calculateFunction("SHEETS", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TYPE = function (arg1) {
		return this.private_calculateFunction("TYPE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.AND = function () {
		return this.private_calculateFunction("AND", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.FALSE = function () {
		return this.private_calculateFunction("FALSE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {boolean} arg1.
	 * @param {any} arg2.
	 * @param {any} [arg3].
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IF = function (arg1, arg2, arg3) {
		return this.private_calculateFunction("IF", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IFERROR = function (arg1, arg2) {
		return this.private_calculateFunction("IFERROR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {any} arg1.
	 * @param {any} arg2.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.IFNA = function (arg1, arg2) {
		return this.private_calculateFunction("IFNA", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @param {boolean} arg1.
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.NOT = function (arg1) {
		return this.private_calculateFunction("NOT", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.OR = function () {
		return this.private_calculateFunction("OR", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
	 */
	ApiWorksheetFunction.prototype.TRUE = function () {
		return this.private_calculateFunction("TRUE", arguments);
	};
	/**
	 * Returns the result of calculating the function.
	 * @memberof ApiWorksheetFunction
	 * @typeofeditors ["CSE"]
	 * @returns {number | string | boolean}
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
				formula.setFormula(formula.getFormula());
				formula.parse();
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
	 * Subscribes to the specified event and calls the callback function when the event fires.
	 * @function
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} eventName - The event name.
	 * @param {function} callback - Function to be called when the event fires.
	 * @fires Api#onWorksheetChange
	 */
	Api.prototype["attachEvent"] = Api.prototype.attachEvent;

	/**
	 * Unsubscribes from the specified event.
	 * @function
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {string} eventName - The event name.
	 * @fires Api#onWorksheetChange
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
	 */

	/**
	 * Sets a type to the freeze panes.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {FreezePaneType} FreezePaneType - The freeze panes type ("null" to unfreeze).
	 * @since 8.0.0
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
	 * Returns the cell references style.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ReferenceStyle}
	 * */
	Api.prototype.GetReferenceStyle = function () {
		let bReferenceStyle = this.asc_getR1C1Mode();
		return bReferenceStyle ? "xlR1C1" : "xlA1";
	};

	/**
	 * Sets the cell references style.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {ReferenceStyle} sReferenceStyle - Type of reference style
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
	 * Returns the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
	 */
	ApiWorksheet.prototype.GetVisible = function () {
		return !this.worksheet.getHidden();
	};

	/**
	 * Sets the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isVisible - Specifies if the sheet is visible or not.
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
	 */
	ApiWorksheet.prototype.GetName = function () {
		return this.worksheet.getName();
	};

	/**
	 * Sets a name to the current active sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sName - The name which will be displayed for the current sheet at the sheet tab.
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
	 */
	ApiWorksheet.prototype.SetRowHeight = function (nRow, nHeight) {
		this.worksheet.setRowHeight(nHeight, nRow, nRow, true);
	};

	/**
	 * Specifies whether the current sheet gridlines must be displayed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isDisplayed - Specifies whether the current sheet gridlines must be displayed or not. The default value is <b>true</b>.
	 */
	ApiWorksheet.prototype.SetDisplayGridlines = function (isDisplayed) {
		this.worksheet.setDisplayGridlines(!!isDisplayed);
	};

	/**
	 * Specifies whether the current sheet row/column headers must be displayed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} isDisplayed - Specifies whether the current sheet row/column headers must be displayed or not. The default value is <b>true</b>.
	 */
	ApiWorksheet.prototype.SetDisplayHeadings = function (isDisplayed) {
		this.worksheet.setDisplayHeadings(!!isDisplayed);
	};

	/**
	 * Sets the left margin of the sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {number} nPoints - The left margin size measured in points.
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
	 * */
	ApiWorksheet.prototype.SetPageOrientation = function (sPageOrientation) {
		this.worksheet.PagePrintOptions.pageSetup.asc_setOrientation('xlLandscape' === sPageOrientation ? 1 : 0);
	};

	/**
	 * Returns the page orientation.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {PageOrientation}
	 * */
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
	 * */
	ApiWorksheet.prototype.GetPrintHeadings = function () {
		return this.worksheet.PagePrintOptions.asc_getHeadings();
	};

	/**
	 * Specifies whether the current sheet row/column headers must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bPrint - Specifies whether the current sheet row/column headers must be printed or not.
	 * */
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
	 * */
	ApiWorksheet.prototype.GetPrintGridlines = function () {
		return this.worksheet.PagePrintOptions.asc_getGridLines();
	};

	/**
	 * Specifies whether the current sheet gridlines must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bPrint - Defines if cell gridlines are printed on this page or not.
	 * */
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
	 * */
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
			return new ApiChart(oChart);
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
	 * */
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
	 */
	ApiWorksheet.prototype.GetAllCharts = function () {
		var allDrawings = this.worksheet.Drawings;
		var allApiDrawings = [];

		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++) {
			if (allDrawings[nDrawing].graphicObject && allDrawings[nDrawing].isChart()) {
				allApiDrawings.push(new ApiChart(allDrawings[nDrawing].graphicObject));
			}
		}
		return allApiDrawings;
	};

	/**
	 * Returns all OLE objects from the current sheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiOleObject[]}.
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
	 * Returns the freeze panes from the current worksheet.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiFreezePanes}
	 * @since 8.0.0
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
			if (this.worksheet.editUserProtectedRanges(null, settings, true)) {
				result = new ApiProtectedRange(settings);
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
	 * Pastes the contents of the Clipboard onto the sheet.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - Object that specifies where the Clipboard contents should be pasted. If this argument is omitted, the current selection is used.
	 * @since 8.1.0
	 */
	ApiWorksheet.prototype.Paste = function (destination) {
		var oApi = Asc["editor"];
		if (destination) {
			if (destination instanceof ApiRange) {
				AscCommon.g_specialPasteHelper && AscCommon.g_specialPasteHelper.Special_Paste_Hide_Button();
				let ws =  destination.range.worksheet;
				private_executeOtherActiveSheet(ws, destination.range, function () {
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
	 */

	/**
	 * Specifies the line style used to form the cell border.
	 * @typedef {("None" | "Double" | "Hair" | "DashDotDot" | "DashDot" | "Dotted" | "Dashed" | "Thin" | "MediumDashDotDot" | "SlantDashDot" | "MediumDashDot" | "MediumDashed" | "Medium" | "Thick")} LineStyle
	 */

	//TODO xlManual param
	/**
	 * Specifies the sort order.
	 * @typedef {("xlAscending" | "xlDescending")}  SortOrder
	 * */

	//TODO xlGuess param
	/**
	 * Specifies whether the first row of the sort range contains the header information.
	 * @typedef {("xlNo" | "xlYes")} SortHeader
	 * */

	/**
	 * Specifies if the sort should be by row or column.
	 * @typedef {("xlSortColumns" | "xlSortRows")} SortOrientation
	 * */

	/**
	 * Specifies the range angle.
	 * @typedef {("xlDownward" | "xlHorizontal" | "xlUpward" | "xlVertical")} Angle
	 */

	/**
	 * Specifies the direction of end in the specified range.
	 * @typedef {("xlUp" | "xlDown" | "xlToRight" | "xlToLeft")} Direction
	 */

	/**
	 * Returns a type of the ApiRange class.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {"range"}
	 */
	ApiRange.prototype.GetClassType = function () {
		return "range";
	};

	/**
	 * Returns a row number for the selected cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
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
	 * @return {boolean} - returns false if such a range does not exist.
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
	 * @return {string | string[][]} - return Value2 property (value without format) if formula doesn't exist.
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
	 * @return {string | string[][]}
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
	 * @return {string | string[][]}
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
	 */
	ApiRange.prototype.GetRowHeight = function () {
		return this.range.worksheet.getRowHeight(this.range.bbox.r1);
	};

	/**
	 * Sets the row height value.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {pt} nHeight - The row height in the current range measured in points.
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
	 * * <b>"none"</b> - for no underlining;
	 * * <b>"single"</b> - for a single line underlining the cell contents;
	 * * <b>"singleAccounting"</b> - for a single line underlining the cell contents but not protruding beyond the cell borders;
	 * * <b>"double"</b> - for a double line underlining the cell contents;
	 * * <b>"doubleAccounting"</b> - for a double line underlining the cell contents but not protruding beyond the cell borders.
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
	 */
	ApiRange.prototype.SetWrap = function (isWrap) {
		this.range.setWrap(!!isWrap);
	};

	/**
	 * Returns the information about the wrapping cell style.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {boolean}
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
	 */
	ApiRange.prototype.SetFillColor = function (oColor) {
		this.range.setFillColor('No Fill' === oColor ? null : oColor.color);
	};
	/**
	 * Returns the background color for the current cell range. Returns 'No Fill' when the color of the background in the cell / cell range is null.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiColor|'No Fill'} - return 'No Fill' when the color to the background in the cell / cell range is null.
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
	 * @return {Angle}
	 */
	ApiRange.prototype.GetOrientation = function () {
		return this.range.getAngle();
	};

	/**
	 * Sets an angle to the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {Angle} angle - Specifies the range angle.
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
	 * @param {?string} shift - Specifies how to shift cells to replace the deleted cells ("up", "left").
	 */
	ApiRange.prototype.Delete = function (shift) {
		let preDeleteAction = function() {
			cellCommentator.updateCommentsDependencies(false, val, checkRange);
			wsView.shiftCellWatches(false, val, bbox);
			wsView.model.shiftDataValidation(false, val, checkRange, true);
			wsView._cleanCache(lockRange);
		};
		let val;
		let ws = this.Worksheet.worksheet;
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
		if (shift == "up") {
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
	 * @return {ApiAreas}
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
	 * Copies the range to the specified range or to the Clipboard.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - Specifies the new range to which the specified range will be copied. If this argument is omitted, Onlyoffice copies the range to the Clipboard.
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
				AscCommon.g_clipboardBase && AscCommon.g_clipboardBase.ClearBuffer();
			} else {
				logError(new Error('Invalid destination'));
			}
		} else {
			let ws =  this.range.worksheet;
			private_executeOtherActiveSheet(ws, this.range, function () {
				oApi && oApi.asc_Copy();
			});
			oApi && oApi.wb.cleanCopyData();
		}
	};

	/**
	 * Cuts the range to the specified range or to the Clipboard.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange?} [destination] - Specifies the new range to which the specified range will be cuted. If this argument is omitted, Onlyoffice copies the range to the Clipboard.
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
				AscCommon.g_clipboardBase && AscCommon.g_clipboardBase.ClearBuffer();
			} else {
				logError(new Error('Invalid destination'));
			}
		} else {
			let ws =  this.range.worksheet;
			private_executeOtherActiveSheet(ws, this.range, function () {
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
	 * Pastes the Range object to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {PasteType} [sPasteType="xlPasteAll"]  - Type of special paste
	 * @param {PasteSpecialOperation} [sPasteSpecialOperation="xlPasteSpecialOperationNone"] - Operation of special paste
	 * @param {boolean} bSkipBlanks [bSkipBlanks=false] - Case sensitive or not. The default value is "false".
	 * @param {boolean} bTranspose [bTranspose=false] - Case sensitive or not. The default value is "false".
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
		private_executeOtherActiveSheet(ws, this.range, function () {
			oApi && oApi.asc_Paste();
		});
	};

	/**
	 * Search data type (formulas or values).
	 * @typedef {("xlFormulas" | "xlValues")} XlFindLookIn
	 */

	/**
	 * Specifies whether the whole search text or any part of the search text is matched.
	 * @typedef {("xlWhole" | "xlPart")} XlLookAt
	 */

	/**
	 * Range search order - by rows or by columns.
	 * @typedef {("xlByRows" | "xlByColumns")} XlSearchOrder
	 */

	/**
	 * Range search direction - next match or previous match.
	 * @typedef {("xlNext" | "xlPrevious")} XlSearchDirection
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
	 * @return {ApiCharacters}
	 * @since 7.4.0
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
	 * */
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
	 */
	ApiDrawing.prototype.GetWidth = function () {
		return private_MM2EMU(this.Drawing.GetWidth());
	};
	/**
	 * Returns the height of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {EMU}
	 */
	ApiDrawing.prototype.GetHeight = function () {
		return private_MM2EMU(this.Drawing.GetHeight());
	};
	/**
	 * Returns the lock value for the specified lock type of the current drawing.
	 * @typeofeditors ["CSE"]
	 * @param {"noGrp" | "noUngrp" | "noSelect" | "noRot" | "noChangeAspect" | "noMove" | "noResize" | "noEditPoints" | "noAdjustHandles"
	 * 	| "noChangeArrowheads" | "noChangeShapeType" | "noDrilldown" | "noTextEdit" | "noCrop" | "txBox"} sType - Lock type in the string format.
	 * @returns {bool}
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
	 * @param {"noGrp" | "noUngrp" | "noSelect" | "noRot" | "noChangeAspect" | "noMove" | "noResize" | "noEditPoints" | "noAdjustHandles"
	 * 	| "noChangeArrowheads" | "noChangeShapeType" | "noDrilldown" | "noTextEdit" | "noCrop" | "txBox"} sType - Lock type in the string format.
	 * @param {bool} bValue - Specifies if the specified lock is applied to the current drawing.
	 * @returns {bool}
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
	 */
	ApiShape.prototype.GetClassType = function () {
		return "shape";
	};

	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @returns {?ApiDocumentContent}
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

	/**
	 * Returns a type of the ApiChart class.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @returns {"chart"}
	 */
	ApiChart.prototype.GetClassType = function () {
		return "chart";
	};

	/**
	 *  Specifies the chart title with the specified parameters.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CSE"]
	 *  @param {string} sTitle - The title which will be displayed for the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the chart title is written in bold font or not.
	 */
	ApiChart.prototype.SetTitle = function (sTitle, nFontSize, bIsBold) {
		AscFormat.builder_SetChartTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};

	/**
	 *  Specifies the chart horizontal axis title.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CSE"]
	 *  @param {string} sTitle - The title which will be displayed for the horizontal axis of the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the horizontal axis title is written in bold font or not.
	 * */
	ApiChart.prototype.SetHorAxisTitle = function (sTitle, nFontSize, bIsBold) {
		AscFormat.builder_SetChartHorAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};

	/**
	 *  Specifies the chart vertical axis title.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CSE"]
	 *  @param {string} sTitle - The title which will be displayed for the vertical axis of the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the vertical axis title is written in bold font or not.
	 * */
	ApiChart.prototype.SetVerAxisTitle = function (sTitle, nFontSize, bIsBold) {
		AscFormat.builder_SetChartVertAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};


	/**
	 * Specifies the direction of the data displayed on the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bIsMinMax - The <code>true</code> value sets the normal data direction for the vertical axis (from minimum to maximum).
	 * The <code>false</code> value sets the inverted data direction for the vertical axis (from maximum to minimum).
	 * */
	ApiChart.prototype.SetVerAxisOrientation = function (bIsMinMax) {
		AscFormat.builder_SetChartVertAxisOrientation(this.Chart, bIsMinMax);
	};


	/**
	 * Specifies the major tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMajorTickMark = function (sTickMark) {
		AscFormat.builder_SetChartHorAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the minor tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMinorTickMark = function (sTickMark) {
		AscFormat.builder_SetChartHorAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the major tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMajorTickMark = function (sTickMark) {
		AscFormat.builder_SetChartVerAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the minor tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMinorTickMark = function (sTickMark) {
		AscFormat.builder_SetChartVerAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the direction of the data displayed on the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bIsMinMax - The <code>true</code> value sets the normal data direction for the horizontal axis
	 * (from minimum to maximum). The <code>false</code> value sets the inverted data direction for the horizontal axis (from maximum to minimum).
	 * */
	ApiChart.prototype.SetHorAxisOrientation = function (bIsMinMax) {
		AscFormat.builder_SetChartHorAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specifies the chart legend position.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {"left" | "top" | "right" | "bottom" | "none"} sLegendPos - The position of the chart legend inside the chart window.
	 * */
	ApiChart.prototype.SetLegendPos = function (sLegendPos) {
		if (sLegendPos === "left" || sLegendPos === "top" || sLegendPos === "right" || sLegendPos === "bottom" || sLegendPos === "none")
			AscFormat.builder_SetChartLegendPos(this.Chart, sLegendPos);
		else
			AscFormat.builder_SetChartLegendPos(this.Chart, "none");
	};

	/**
	 * Specifies the legend font size.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 * */
	ApiChart.prototype.SetLegendFontSize = function (nFontSize) {
		AscFormat.builder_SetLegendFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies which chart data labels are shown for the chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {boolean} bShowVal - Whether to show or hide the chart data values.
	 * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowDataLabels = function (bShowSerName, bShowCatName, bShowVal, bShowPercent) {
		AscFormat.builder_SetShowDataLabels(this.Chart, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};

	/**
	 * Spicifies the show options for the data labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {number} nSeriesIndex - The series index from the array of the data used to build the chart from.
	 * @param {number} nPointIndex - The point index from this series.
	 * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {boolean} bShowVal - Whether to show or hide the chart data values.
	 * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowPointDataLabel = function (nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent) {
		AscFormat.builder_SetShowPointDataLabel(this.Chart, nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};

	/**
	 * Sets the possible values for the position of the chart tick labels in relation to the main vertical label or the chart data values.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart vertical tick labels.
	 * */
	ApiChart.prototype.SetVertAxisTickLabelPosition = function (sTickLabelPosition) {
		AscFormat.builder_SetChartVertAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Sets the possible values for the position of the chart tick labels in relation to the main horizontal label or the chart data values.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart horizontal tick labels.
	 * */
	ApiChart.prototype.SetHorAxisTickLabelPosition = function (sTickLabelPosition) {
		AscFormat.builder_SetChartHorAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Specifies the visual properties of the major vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorVerticalGridlines = function (oStroke) {
		AscFormat.builder_SetVerAxisMajorGridlines(this.Chart, oStroke ? oStroke.Ln : null);
	};

	/**
	 * Specifies the visual properties of the minor vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorVerticalGridlines = function (oStroke) {
		AscFormat.builder_SetVerAxisMinorGridlines(this.Chart, oStroke ? oStroke.Ln : null);
	};


	/**
	 * Specifies the visual properties of the major horizontal gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorHorizontalGridlines = function (oStroke) {
		AscFormat.builder_SetHorAxisMajorGridlines(this.Chart, oStroke ? oStroke.Ln : null);
	};

	/**
	 * Specifies the visual properties of the minor vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 */
	ApiChart.prototype.SetMinorHorizontalGridlines = function (oStroke) {
		AscFormat.builder_SetHorAxisMinorGridlines(this.Chart, oStroke ? oStroke.Ln : null);
	};


	/**
	 * Specifies the font size to the horizontal axis labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetHorAxisLablesFontSize = function (nFontSize) {
		AscFormat.builder_SetHorAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies the font size to the vertical axis labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetVertAxisLablesFontSize = function (nFontSize) {
		AscFormat.builder_SetVerAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Sets a style to the current chart by style ID.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param nStyleId - One of the styles available in the editor.
	 * @returns {boolean}
	 */
	ApiChart.prototype.ApplyChartStyle = function (nStyleId) {
		if (typeof (nStyleId) !== "number" || nStyleId < 0)
			return false;

		var nChartType = this.Chart.getChartType();
		var aStyle = AscCommon.g_oChartStyles[nChartType] && AscCommon.g_oChartStyles[nChartType][nStyleId];

		if (aStyle) {
			this.Chart.applyChartStyleByIds(aStyle);
			return true;
		}

		return false;
	};

	/**
	 * Sets values from the specified range to the specified series.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with series values. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column,
	 * * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaValues = function (sRange, nSeria) {
		return this.Chart.SetSeriaValues(sRange, nSeria);
	};

	/**
	 * Sets the x-axis values from the specified range to the specified series. It is used with the scatter charts only.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with series x-axis values. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column,
	 * * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaXValues = function (sRange, nSeria) {
		return this.Chart.SetSeriaXValues(sRange, nSeria);
	};

	/**
	 * Sets a name to the specified series.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sNameRange - The series name. Can be a range of cells or usual text. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column,
	 * * "Example series".
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaName = function (sNameRange, nSeria) {
		return this.Chart.SetSeriaName(sNameRange, nSeria);
	};

	/**
	 * Sets a range with the category values to the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sRange - A range of cells from the sheet with the category names. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column.
	 */
	ApiChart.prototype.SetCatFormula = function (sRange) {
		return this.Chart.SetCatFormula(sRange);
	};

	/**
	 * Adds a new series to the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {string} sNameRange - The series name. Can be a range of cells or usual text. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column,
	 * * "Example series".
	 * @param {string} sValuesRange - A range of cells from the sheet with series values. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column.
	 * @param {string} [sXValuesRange=undefined] - A range of cells from the sheet with series x-axis values. It is used with the scatter charts only. For example:
	 * * "'sheet 1'!$A$2:$A$5" - must be a single cell, row or column,
	 * * "A1:A5" - must be a single cell, row or column.
	 */
	ApiChart.prototype.AddSeria = function (sNameRange, sValuesRange, sXValuesRange) {
		if (this.Chart.isScatterChartType() && typeof (sXValuesRange) === "string" && sXValuesRange !== "") {
			this.Chart.addScatterSeries(sNameRange, sXValuesRange, sValuesRange);
		} else
			this.Chart.addSeries(sNameRange, sValuesRange);
	};

	/**
	 * Removes the specified series from the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.RemoveSeria = function (nSeria) {
		return this.Chart.RemoveSeria(nSeria);
	};

	/**
	 * Sets the fill to the chart plot area.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the plot area.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetPlotAreaFill = function (oFill) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		this.Chart.SetPlotAreaFill(oFill.UniFill);
		return true;
	};

	/**
	 * Sets the outline to the chart plot area.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the plot area outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetPlotAreaOutLine = function (oStroke) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		this.Chart.SetPlotAreaOutLine(oStroke.Ln);
		return true;
	};

	/**
	 * Sets the fill to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the series.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {boolean} [bAll=false] - Specifies if the fill will be applied to all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriesFill = function (oFill, nSeries, bAll) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetSeriesFill(oFill.UniFill, nSeries, bAll);
	};

	/**
	 * Sets the outline to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the series outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {boolean} [bAll=false] - Specifies if the outline will be applied to all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriesOutLine = function (oStroke, nSeries, bAll) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetSeriesOutLine(oStroke.Ln, nSeries, bAll);
	};

	/**
	 * Sets the fill to the data point in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the data point.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nDataPoint - The index of the data point in the specified chart series.
	 * @param {boolean} [bAllSeries=false] - Specifies if the fill will be applied to the specified data point in all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetDataPointFill = function (oFill, nSeries, nDataPoint, bAllSeries) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetDataPointFill(oFill.UniFill, nSeries, nDataPoint, bAllSeries);
	};

	/**
	 * Sets the outline to the data point in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the data point outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nDataPoint - The index of the data point in the specified chart series.
	 * @param {boolean} bAllSeries - Specifies if the outline will be applied to the specified data point in all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetDataPointOutLine = function (oStroke, nSeries, nDataPoint, bAllSeries) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetDataPointOutLine(oStroke.Ln, nSeries, nDataPoint, bAllSeries);
	};

	/**
	 * Sets the fill to the marker in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the marker.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nMarker - The index of the marker in the specified chart series.
	 * @param {boolean} [bAllMarkers=false] - Specifies if the fill will be applied to all markers in the specified chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetMarkerFill = function (oFill, nSeries, nMarker, bAllMarkers) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetMarkerFill(oFill.UniFill, nSeries, nMarker, bAllMarkers);
	};

	/**
	 * Sets the outline to the marker in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the marker outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nMarker - The index of the marker in the specified chart series.
	 * @param {boolean} [bAllMarkers=false] - Specifies if the outline will be applied to all markers in the specified chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetMarkerOutLine = function (oStroke, nSeries, nMarker, bAllMarkers) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetMarkerOutLine(oStroke.Ln, nSeries, nMarker, bAllMarkers);
	};

	/**
	 * Sets the fill to the chart title.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the title.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetTitleFill = function (oFill) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetTitleFill(oFill.UniFill);
	};

	/**
	 * Sets the outline to the chart title.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the title outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetTitleOutLine = function (oStroke) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetTitleOutLine(oStroke.Ln);
	};

	/**
	 * Sets the fill to the chart legend.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the legend.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetLegendFill = function (oFill) {
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetLegendFill(oFill.UniFill);
	};

	/**
	 * Sets the outline to the chart legend.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the legend outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetLegendOutLine = function (oStroke) {
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetLegendOutLine(oStroke.Ln);
	};

	/**
	 * Sets the specified numeric format to the axis values.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {NumFormat | String} sFormat - Numeric format (can be custom format).
	 * @param {AxisPos} - Axis position.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetAxieNumFormat = function (sFormat, sAxiePos) {
		var nAxiePos = -1;
		switch (sAxiePos) {
			case "bottom":
				nAxiePos = AscFormat.AX_POS_B;
				break;
			case "left":
				nAxiePos = AscFormat.AX_POS_L;
				break;
			case "right":
				nAxiePos = AscFormat.AX_POS_R;
				break;
			case "top":
				nAxiePos = AscFormat.AX_POS_T;
				break;
			default:
				return false;
		}

		return this.Chart.SetAxieNumFormat(sFormat, nAxiePos);
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
	 */
	ApiName.prototype.SetRefersTo = function (sRef) {
		this.DefName.setRef(sRef);
	};

	/**
	 * Returns a formula that the name is defined to refer to.
	 * @memberof ApiName
	 * @typeofeditors ["CSE"]
	 * @returns {string}
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
	 */
	ApiComment.prototype.GetClassType = function () {
		return "comment";
	};

	/**
	 * Returns the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
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
	 */
	ApiAreas.prototype.GetItem = function (ind) {
		return this.Items[ind - 1] || null;
	};

	/**
	 * Returns the parent object for the specified collection.
	 * @memberof ApiAreas
	 * @typeofeditors ["CSE"]
	 * @returns {number}
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
	 */

	/**
	 * Returns the type of underline applied to the specified font.
	 * @memberof ApiFont
	 * @typeofeditors ["CSE"]
	 * @returns {XlUnderlineStyle | null}
	 * @since 7.4.0
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
	 * @return {boolean} - Returns false if a user doesn't have permission to modify the protected range.
	 * @since 8.1.0
	 */
	ApiProtectedRange.prototype.SetTitle = function (sTitle) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let result = false;
		if (isValidTitle && sTitle !== this.protectedRange.asc_getName()) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
				newProtectedRange.asc_setName(sTitle);
				if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
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
	 * @return {boolean} - Returns false if a user doesn't have permission to modify the protected range.
	 * @since 8.1.0
	 */
	ApiProtectedRange.prototype.SetRange = function (sRange) {
		let isValidRange = typeof (sRange) === 'string' && sRange.trim() !== '';
		let result = false;
		if (isValidRange /*asc_getRef !==*/) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone(this.protectedRange._ws, true);
				newProtectedRange.asc_setRef(sRange);
				if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Specifies the user type of the protected range.
	 * @typedef {("CanEdit" | "CanView" | "NotView")} ProtectedRangeUserType
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
				users.push(newUser);
				newProtectedRange.asc_setUsers(users);
				worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true);
				result = new ApiProtectedRangeUserInfo(result, this.protectedRange);
			}
		}
		return result;
	};

	/**
	 * Removes a user from the current protected range.
	 * @memberof ApiProtectedRange
	 * @param {string} sId - The user ID.
	 * @returns {bool}
	 * @since 8.1.0
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
				if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Returns an object that represents a user from the current protected range.
	 * @memberof ApiProtectedRange
	 * @param {string} sId - The user ID.
	 * @returns {ApiProtectedRangeUserInfo | null}
	 * @since 8.1.0
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
	 */
	ApiProtectedRangeUserInfo.prototype.GetId = function () {
		return this.userInfo.asc_getId();
	};


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

	Api.prototype["AddCustomFunction"] = Api.prototype.AddCustomFunction;

	Api.prototype["GetReferenceStyle"] = Api.prototype.GetReferenceStyle;
	Api.prototype["SetReferenceStyle"] = Api.prototype.SetReferenceStyle;

	Api.prototype["GetWorksheetFunction"] = Api.prototype.GetWorksheetFunction;
	
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

	ApiChart.prototype["GetClassType"]                 =  ApiChart.prototype.GetClassType;
	ApiChart.prototype["SetTitle"]                     =  ApiChart.prototype.SetTitle;
	ApiChart.prototype["SetHorAxisTitle"]              =  ApiChart.prototype.SetHorAxisTitle;
	ApiChart.prototype["SetVerAxisTitle"]              =  ApiChart.prototype.SetVerAxisTitle;
	ApiChart.prototype["SetVerAxisOrientation"]        =  ApiChart.prototype.SetVerAxisOrientation;
	ApiChart.prototype["SetHorAxisOrientation"]        =  ApiChart.prototype.SetHorAxisOrientation;
	ApiChart.prototype["SetLegendPos"]                 =  ApiChart.prototype.SetLegendPos;
	ApiChart.prototype["SetLegendFontSize"]            =  ApiChart.prototype.SetLegendFontSize;
	ApiChart.prototype["SetShowDataLabels"]            =  ApiChart.prototype.SetShowDataLabels;
	ApiChart.prototype["SetShowPointDataLabel"]        =  ApiChart.prototype.SetShowPointDataLabel;
	ApiChart.prototype["SetVertAxisTickLabelPosition"] =  ApiChart.prototype.SetVertAxisTickLabelPosition;
	ApiChart.prototype["SetHorAxisTickLabelPosition"]  =  ApiChart.prototype.SetHorAxisTickLabelPosition;

	ApiChart.prototype["SetHorAxisMajorTickMark"]  =  ApiChart.prototype.SetHorAxisMajorTickMark;
	ApiChart.prototype["SetHorAxisMinorTickMark"]  =  ApiChart.prototype.SetHorAxisMinorTickMark;
	ApiChart.prototype["SetVertAxisMajorTickMark"]  =  ApiChart.prototype.SetVertAxisMajorTickMark;
	ApiChart.prototype["SetVertAxisMinorTickMark"]  =  ApiChart.prototype.SetVertAxisMinorTickMark;



	ApiChart.prototype["SetMajorVerticalGridlines"]   =  ApiChart.prototype.SetMajorVerticalGridlines;
	ApiChart.prototype["SetMinorVerticalGridlines"]   =  ApiChart.prototype.SetMinorVerticalGridlines;
	ApiChart.prototype["SetMajorHorizontalGridlines"] =  ApiChart.prototype.SetMajorHorizontalGridlines;
	ApiChart.prototype["SetMinorHorizontalGridlines"] =  ApiChart.prototype.SetMinorHorizontalGridlines;
	ApiChart.prototype["SetHorAxisLablesFontSize"]    =  ApiChart.prototype.SetHorAxisLablesFontSize;
	ApiChart.prototype["SetVertAxisLablesFontSize"]   =  ApiChart.prototype.SetVertAxisLablesFontSize;
	ApiChart.prototype["ApplyChartStyle"]             =  ApiChart.prototype.ApplyChartStyle;
	ApiChart.prototype["SetSeriaValues"]              =  ApiChart.prototype.SetSeriaValues;
	ApiChart.prototype["SetSeriaXValues"]             =  ApiChart.prototype.SetSeriaXValues;
	ApiChart.prototype["SetSeriaName"]                =  ApiChart.prototype.SetSeriaName;
	ApiChart.prototype["SetCatFormula"]               =  ApiChart.prototype.SetCatFormula;
	ApiChart.prototype["AddSeria"]                    =  ApiChart.prototype.AddSeria;
	ApiChart.prototype["RemoveSeria"]                 =  ApiChart.prototype.RemoveSeria;
	ApiChart.prototype["SetPlotAreaFill"]             =  ApiChart.prototype.SetPlotAreaFill;
	ApiChart.prototype["SetPlotAreaOutLine"]          =  ApiChart.prototype.SetPlotAreaOutLine;
	ApiChart.prototype["SetSeriesFill"]               =  ApiChart.prototype.SetSeriesFill;
	ApiChart.prototype["SetSeriesOutLine"]            =  ApiChart.prototype.SetSeriesOutLine;
	ApiChart.prototype["SetDataPointFill"]            =  ApiChart.prototype.SetDataPointFill;
	ApiChart.prototype["SetDataPointOutLine"]         =  ApiChart.prototype.SetDataPointOutLine;
	ApiChart.prototype["SetMarkerFill"]               =  ApiChart.prototype.SetMarkerFill;
	ApiChart.prototype["SetMarkerOutLine"]            =  ApiChart.prototype.SetMarkerOutLine;
	ApiChart.prototype["SetTitleFill"]                =  ApiChart.prototype.SetTitleFill;
	ApiChart.prototype["SetTitleOutLine"]             =  ApiChart.prototype.SetTitleOutLine;
	ApiChart.prototype["SetLegendFill"]               =  ApiChart.prototype.SetLegendFill;
	ApiChart.prototype["SetLegendOutLine"]            =  ApiChart.prototype.SetLegendOutLine;
	ApiChart.prototype["SetAxieNumFormat"]            =  ApiChart.prototype.SetAxieNumFormat;

	ApiOleObject.prototype["GetClassType"]            = ApiOleObject.prototype.GetClassType;
	ApiOleObject.prototype["SetData"]              = ApiOleObject.prototype.SetData;
	ApiOleObject.prototype["GetData"]              = ApiOleObject.prototype.GetData;
	ApiOleObject.prototype["SetApplicationId"]        = ApiOleObject.prototype.SetApplicationId;
	ApiOleObject.prototype["GetApplicationId"]        = ApiOleObject.prototype.GetApplicationId;

	ApiColor.prototype["GetClassType"]                 =  ApiColor.prototype.GetClassType;


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
	ApiWorksheetFunction.prototype["BINOM_INV"]       =  ApiWorksheetFunction.prototype.BINOM_INV;
	ApiWorksheetFunction.prototype["CHISQ_INV"]       =  ApiWorksheetFunction.prototype.CHISQ_INV;
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
	};

	function logError(err) {
		if (console.error)
			console.error(err);
		else
			console.log(err);
	};

	function throwException(err) {
		if (!console.error)
			logError(err);
		throw err;
	};
	function private_executeOtherActiveSheet(ws, range, func) {
		let oldActiveSheet = ws && ws.workbook && ws.workbook.getActive();
		let isChangedActiveSheet;
		if (oldActiveSheet != null && ws && oldActiveSheet !== ws.index) {
			ws.workbook.setActive(ws.index);
			isChangedActiveSheet = true;
		}

		let oldSelection;
		if (range) {
			oldSelection = ws.selectionRange.clone();
			let newSelection = new AscCommonExcel.SelectionRange(ws);
			let bbox = range.bbox;
			newSelection.assign2(bbox);
			newSelection.Select(true);
		}

		func();

		isChangedActiveSheet && ws.workbook.setActive(oldActiveSheet);
		oldSelection && oldSelection.Select(true);
	};

}(window, null));
