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
			logError(new Error('Invalid parametr "FreezePaneType".'));
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
					user.asc_setName(userInfo.asc_getName());

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
	 * Returns a protected range object.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sTitle - The title which will be displayed for the current protected range.
	 * @returns {ApiProtectedRange | null}
	 */
	ApiWorksheet.prototype.GetProtectedRange = function (sTitle) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let result = null;
		if (isValidTitle) {
			let protectedRange = this.worksheet.getProtectedRangeByName(sTitle);
			result = protectedRange && protectedRange.val ? new ApiProtectedRange(protectedRange.val.clone()) : null;
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
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {ApiProtectedRange[] | null}
	 */
	ApiWorksheet.prototype.GetAllProtectedRanges = function () {
		let protectedRanges = this.worksheet && this.worksheet.workbook && this.worksheet.workbook.oApi.asc_getProtectedRanges();
		let result = null;
		if (protectedRanges) {
			result = [];
			for (let i  = 0; i < protectedRanges.length; i++) {
				result.push(new ApiProtectedRange(protectedRanges[i].clone()));
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
			col--
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
	 * Copies a range to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} destination - Specifies a new range to which the specified range will be copied.
	 */
	ApiRange.prototype.Copy = function (destination) {
		if (destination && destination instanceof ApiRange) {
			let bboxFrom = this.range.bbox;
			let cols = bboxFrom.c2 - bboxFrom.c1;
			let rows = bboxFrom.r2 - bboxFrom.r1;
			let bbox = destination.range.bbox;
			let range = destination.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols));
			this.range.move(range.bbox, true, destination.range.worksheet);
		} else {
			logError(new Error('Invalid destination'));
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
			options.asc_setSpecificRange(this.Address);
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
			options.asc_setSpecificRange(this.Address);
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
	 * Class representing user protected range.
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
	 * @return {boolean} - returns false if user don't have rules on change protected range.
	 */
	ApiProtectedRange.prototype.SetTitle = function (sTitle) {
		let isValidTitle = typeof (sTitle) === 'string' && sTitle.trim() !== '';
		let result = false;
		if (isValidTitle && sTitle !== this.protectedRange.asc_getName()) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone();
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
	 * @param {string} sRange - The range of cells from the current protected range.
	 * @return {boolean} - returns false if user don't have rules on change protected range.
	 */
	ApiProtectedRange.prototype.SetRange = function (sRange) {
		let isValidRange = typeof (sRange) === 'string' && sRange.trim() !== '';
		let result = false;
		if (isValidRange /*asc_getRef !==*/) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone();
				newProtectedRange.asc_setRef(sRange);
				if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Specifies the user protected range type.
	 * @typedef {("CanEdit" | "CanView" | "NotView")} ProtectedRangeUserType
	 */


	/**
	 * Sets the user to the current protected range.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sId - The user id.
	 * @param {string} sName - The user name.
	 * @param {ProtectedRangeUserType} protectedRangeUserType - Specifies the protected range user type.
	 * @returns {ApiProtectedRangeUserInfo | null} - returns null if user don't have rules on change protected range.
	 */
	ApiProtectedRange.prototype.AddUser = function (sId, sName, protectedRangeUserType) {
		let isValidIdTitle = typeof (sId) === 'string' && sId.trim() !== '';
		let isValidTitle = typeof (sName) === 'string' && sName.trim() !== '';
		let result = null;
		if (isValidTitle && isValidIdTitle) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let newProtectedRange = this.protectedRange.clone();

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
	 * Remove the user current protected range.
	 * @memberof ApiProtectedRange
	 * @param {string} sId - The user id.
	 * @returns {bool}
	 */
	ApiProtectedRange.prototype.DeleteUser = function (sId) {
		let isValidId = typeof (sId) === 'string' && sId.trim() !== '';
		let result = false;
		if (isValidId) {
			let worksheet = this.protectedRange._ws;
			if (worksheet) {
				let userInfo = this.protectedRange.getUserById(sId);
				if (userInfo) {
					let newProtectedRange = this.protectedRange.clone();
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
	 * Returns all protected range users from the current worksheet.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @returns {ApiProtectedRangeUserInfo[] | null}
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
	 * Sets protected type for anyone user.
	 * @memberof ApiProtectedRange
	 * @typeofeditors ["CSE"]
	 * @param {ProtectedRangeUserType} protectedRangeUserType - Specifies the protected range user type.
	 * @returns {bool}
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
				let newProtectedRange = this.protectedRange.clone();
				newProtectedRange.asc_setType(nType);
				if (worksheet.editUserProtectedRanges(this.protectedRange, newProtectedRange, true)) {
					result = true;
				}
			}
		}
		return result;
	};

	/**
	 * Returns an object that represents the user protected range.
	 * @memberof ApiProtectedRange
	 * @param {string} sId - The user id.
	 * @returns {ApiProtectedRangeUserInfo | null}
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
	 * Class representing user protected range.
	 * @constructor
	 */
	function ApiProtectedRangeUserInfo(userInfo, protectedRange) {
		this.userInfo = userInfo;
		this.protectedRange = protectedRange;
	}

	/**
	 * Returns the name property of the current user's info.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
	 */
	ApiProtectedRangeUserInfo.prototype.GetName = function () {
		//the sets methods are available from the parent
		// not adding by ApiProtectedRangeUserInfo because need change id/name together
		return this.userInfo.asc_getName();
	};

	/**
	 * Returns the type property of the current user's info.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {ProtectedRangeUserType}
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
	 * Returns the id property of the current user's info.
	 * @memberof ApiProtectedRangeUserInfo
	 * @typeofeditors ["CSE"]
	 * @returns {string | null}
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

	ApiRange.prototype["GetClassType"] = ApiRange.prototype.GetClassType
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
	ApiRange.prototype["Paste"] = ApiRange.prototype.Paste;
	ApiRange.prototype["Find"] = ApiRange.prototype.Find;
	ApiRange.prototype["FindNext"] = ApiRange.prototype.FindNext;
	ApiRange.prototype["FindPrevious"] = ApiRange.prototype.FindPrevious;
	ApiRange.prototype["Replace"] = ApiRange.prototype.Replace;
	ApiRange.prototype["GetCharacters"] = ApiRange.prototype.GetCharacters;


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

}(window, null));
