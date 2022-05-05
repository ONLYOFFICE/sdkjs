/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
		if (value.getTime){
			return new AscCommonExcel.cNumber(new Asc.cDate(value.getTime()).getExcelDate());
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
	 */
	var Api = window["Asc"]["spreadsheet_api"];

	/**
	 * Class representing a sheet.
	 * @constructor
	 * @property {bool} Visible - Returns or sets the state of sheet visibility.
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
	 * @property {bool} PrintHeadings - Returns or sets the page PrintHeadings property.
	 * @property {bool} PrintGridlines - Returns or sets the page PrintGridlines property.
	 * @property {Array} Defnames - Returns an array of the ApiName objects.
	 * @property {Array} Comments - Returns an array of the ApiComment objects.
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
	 * @property {number} Count - Returns the rows or columns count.
	 * @property {string} Value - Returns a value from the first cell of the specified range or sets it to this cell.
	 * @property {string} Formula - Returns a formula from the first cell of the specified range or sets it to this cell.
	 * @property {string} Value2 - Returns the value2 (value without format) from the first cell of the specified range or sets it to this cell.
	 * @property {string} Text - Returns the text from the first cell of the specified range or sets it to this cell.
	 * @property {ApiColor} FontColor - Sets the text color to the current cell range with the previously created color object.
	 * @property {bool} Hidden - Returns or sets the value hiding property.
	 * @property {number} ColumnWidth - Returns or sets the width of all the columns in the specified range measured in points.
	 * @property {number} Width - Returns a value that represents the range width measured in points.
	 * @property {number} RowHeight - Returns or sets the height of the first row in the specified range measured in points.
	 * @property {number} Height - Returns a value that represents the range height measured in points.
	 * @property {number} FontSize - Sets the font size to the characters of the current cell range.
	 * @property {string} FontName - Sets the specified font family as the font name for the current cell range.
	 * @property {'center' | 'bottom' | 'top' | 'distributed' | 'justify'} AlignVertical - Sets the text vertical alignment to the current cell range.
	 * @property {'left' | 'right' | 'center' | 'justify'} AlignHorizontal - Sets the text horizontal alignment to the current cell range.
	 * @property {bool} Bold - Sets the bold property to the text characters from the current cell or cell range.
	 * @property {'none' | 'single' | 'singleAccounting' | 'double' | 'doubleAccounting'} Underline - Sets the type of underline applied to the font.
	 * @property {bool} Strikeout - Sets a value that indicates whether the contents of the current cell or cell range are displayed struck through.
	 * @property {ApiColor|'No Fill'} FillColor - Returns or sets the background color of the current cell range.
	 * @property {string} NumberFormat - Sets a value that represents the format code for the object.
	 * @property {ApiRange} MergeArea - Returns the cell or cell range from the merge area.
	 * @property {ApiWorksheet} Worksheet - Returns the ApiWorksheet object that represents the worksheet containing the specified range.
	 * @property {ApiName} DefName - Returns the ApiName object.
	 * @property {ApiComment | null} Comments - Returns the ApiComment collection that represents all the comments from the specified worksheet.
	 * @property {'xlDownward' | 'xlHorizontal' | 'xlUpward' | 'xlVertical'} Orientation - Sets an angle to the current cell range.
	 * @property {ApiAreas} Areas - Returns a collection of the areas.
	 */
	function ApiRange(range, areas) {
		this.range = range;
		this.areas = areas || null;
	}


	/**
	 * Class representing a graphical object.
	 * @constructor
	 */
	function ApiDrawing(Drawing)
	{
		this.Drawing = Drawing;
	}

	/**
	 * Class representing a shape.
	 * @constructor
	 */
	function ApiShape(oShape){
		ApiDrawing.call(this, oShape);
		this.Shape = oShape;
	}
	ApiShape.prototype = Object.create(ApiDrawing.prototype);
	ApiShape.prototype.constructor = ApiShape;

	/**
	 * Class representing an image.
	 * @constructor
	 */
	function ApiImage(oImage){
		ApiDrawing.call(this, oImage);
	}
	ApiImage.prototype = Object.create(ApiDrawing.prototype);
	ApiImage.prototype.constructor = ApiImage;

	/**
	 * Class representing a chart.
	 * @constructor
	 */
	function ApiChart(oChart){
		ApiDrawing.call(this, oChart);
		this.Chart = oChart;
	}
	ApiChart.prototype = Object.create(ApiDrawing.prototype);
	ApiChart.prototype.constructor = ApiChart;

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
     * * **"none"** - does not display the selected tick labels.
     * * **"nextTo"** - sets the position of the selected tick labels next to the main label.
     * * **"low"** - sets the position of the selected tick labels in the part of the chart with lower values.
     * * **"high"** - sets the position of the selected tick labels in the part of the chart with higher values.
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
	 * Class representing a base class for color types.
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
	 * @property {string} Text - Returns the text from the first cell in range.
	 */
	function ApiComment(comment, ws) {
		this.Comment = comment;
		this.WS = ws;
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
	Api.prototype.SetLocale = function(LCID) {
		this.asc_setLocale(LCID, null, null);
	};
	
	/**
	 * Returns the current locale ID.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 */
	Api.prototype.GetLocale = function() {
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
	 * @returns {bool} - returns false if sTheme isn't a string.
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
	Api.prototype.CreateNewHistoryPoint = function(){
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
	 * @returns {ApiRange | Error}
	 */
	Api.prototype.Intersect  = function (Range1, Range2) {
		if (Range1.GetWorksheet().Id === Range2.GetWorksheet().Id) {
			var res = Range1.range.bbox.intersection(Range2.range.bbox);
			if (!res) {
				return "Ranges do not intersect.";
			} else {
				return new ApiRange(this.GetActiveSheet().worksheet.getRange3(res.r1, res.c1, res.r2, res.c2));
			}
		} else {
			return new Error('Ranges should be from one worksheet.');
		}
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
	 * @param {bool} isHidden - Defines if the range name is hidden or not.
	 * @returns {Error | true} - returns error if sName or sRef are invalid.
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
	Api.prototype.GetRange = function(sRange) {
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
	 * Returns the mail merge fields.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @param {number} nSheet - The sheet index.
	 * @returns {string[]}
	 */
	Api.prototype.private_GetMailMergeFields = function (nSheet) {
		var oSheet     = this.GetSheet(nSheet);
		var arrFields  = [];
		var colIndex   = 0;
		var colsCount  = 0;
		var oRange     = oSheet.GetRangeByNumber(1, colIndex);
		var fieldValue = undefined;

		while (oRange.GetValue() !== "") {
			colsCount++;
			colIndex++;
			oRange = oSheet.GetRangeByNumber(1, colIndex);
		}
			
		for (var nCol = 0; nCol < colsCount; nCol++) {
			oRange     = oSheet.GetRangeByNumber(0, nCol);
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
		var oSheet           = this.GetSheet(nSheet);
		var arrMailMergeMap  = [];
		var valuesInRow      = null;

		var rowIndex         = 1;
		var rowsCount        = 0;
		var colIndex         = 0;
		var colsCount        = 0;

		var mergeValue       = undefined;
		
		var oRange           = oSheet.GetRangeByNumber(rowIndex, 0);

		// определяем количество строк с данными
		while (oRange.GetValue() !== "") {
			rowsCount++;
			rowIndex++;
			oRange = oSheet.GetRangeByNumber(rowIndex, 0);
		}

		oRange     = oSheet.GetRangeByNumber(1, colIndex);
		// определяем количество столбцов с данными
		while (oRange.GetValue() !== "") {
			colsCount++;
			colIndex++;
			oRange = oSheet.GetRangeByNumber(1, colIndex);
		}	
		
		for (var nRow = 1; nRow < rowsCount + 1; nRow++) {
			valuesInRow = [];

			for (var nCol = 0; nCol < colsCount; nCol++) {
				oRange     = oSheet.GetRangeByNumber(nRow, nCol);
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
	Api.prototype.GetMailMergeData = function(nSheet, bWithFormat) {
		if (bWithFormat !== true)
			bWithFormat = false;

		var arrFields       = this.private_GetMailMergeFields(nSheet);
		var arrMailMergeMap = this.private_GetMailMergeMap(nSheet, arrFields, bWithFormat);
		var resultList      = [arrFields];

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
	 * @returns {bool}
	 */
	Api.prototype.RecalculateAllFormulas = function(fLogger) {
		var formulas = this.wbModel.getAllFormulas(true);
		var _compare = function(_val1, _val2) {
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
	 * Returns the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {bool}
	 */
	ApiWorksheet.prototype.GetVisible = function () {
		return !this.worksheet.getHidden();
	};

	/**
	 * Sets the state of sheet visibility.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {bool} isVisible - Specifies if the sheet is visible or not.
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
		var cell = this.worksheet.selectionRange.activeCell;
		return new ApiRange(this.worksheet.getCell3(cell.row, cell.col));
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
	 * @returns {ApiRange}
	 */
	ApiWorksheet.prototype.GetCells = function (row, col) {
		if (row) row--;
		if (typeof col !== "undefined" && typeof row !== "undefined") {
			return new ApiRange(this.worksheet.getRange3(row, col, row, col));
		} else if (typeof row !== "undefined") {
			var r = (row) ?  (row / AscCommon.gc_nMaxCol0) >> 0 : row;
			var c = (row) ? row % AscCommon.gc_nMaxCol0 : row;
			if (r && c) c--;
			return new ApiRange(this.worksheet.getRange3(r, c, r, c));
		}
		else {
			return new ApiRange(this.worksheet.getRange3(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0));
		}
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
	 * @returns {ApiRange}
	 */
	ApiWorksheet.prototype.GetRows = function (value) {
		if (typeof  value === "undefined") {
			return this.GetCells();
		} else if (typeof value == "number" || value.indexOf(':') == -1) {
			value = parseInt(value);
			if (value > 0 && value <=  AscCommon.gc_nMaxRow0 + 1 && value[0] !== NaN) {
				value --;
			} else {
				return new Error('The nRow must be greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1));
			}
			return new ApiRange(this.worksheet.getRange3(value, 0, value, AscCommon.gc_nMaxCol0));
		} else {
			value = value.split(':');
			var isError = false;
			for (var i = 0; i < value.length; ++i) {
				value[i] = parseInt(value[i]);
				if (value[i] > 0 && value[i] <= AscCommon.gc_nMaxRow0 + 1 && value[0] !== NaN) {
					value[i] --;
				} else {
					isError = true;
				}
			}
			if (isError) {
				return new Error('The nRow must be greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1));
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
		var sOldName = this.worksheet.getName();
		this.worksheet.setName(sName);
		var oWorkbook = this.worksheet.workbook;
		if(oWorkbook) {
			oWorkbook.handleChartsOnChangeSheetName(this.worksheet, sOldName, sName)
		}
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
	 * @param {string} sRange - The range of cells from the current sheet.
	 * @returns {ApiRange | null} - returns null if such a range does not exist.
	 */
	ApiWorksheet.prototype.GetRange = function (sRange) {
		var Range = this.worksheet.getRange2(sRange);

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
	 */
	ApiWorksheet.prototype.SetColumnWidth = function (nColumn, nWidth) {
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
	 * @param {bool} isDisplayed - Specifies whether the current sheet gridlines must be displayed or not. The default value is <b>true</b>.
	 */
	ApiWorksheet.prototype.SetDisplayGridlines = function (isDisplayed) {
		this.worksheet.setDisplayGridlines(!!isDisplayed);
	};

	/**
	 * Specifies whether the current sheet row/column headers must be displayed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {bool} isDisplayed - Specifies whether the current sheet row/column headers must be displayed or not. The default value is <b>true</b>.
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
	ApiWorksheet.prototype.GetPageOrientation = function ()	{
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
	 * @returns {bool} - Specifies whether the current sheet row/column headings must be printed or not.
	 * */
	ApiWorksheet.prototype.GetPrintHeadings = function ()	{
		return this.worksheet.PagePrintOptions.asc_getHeadings();
	};

	/**
	 * Specifies whether the current sheet row/column headers must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {bool} bPrint - Specifies whether the current sheet row/column headers must be printed or not.
	 * */
	ApiWorksheet.prototype.SetPrintHeadings = function (bPrint)	{
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
	 * @returns {bool} - True if cell gridlines are printed on this page.
	 * */
	ApiWorksheet.prototype.GetPrintGridlines = function ()	{
		return this.worksheet.PagePrintOptions.asc_getGridLines();
	};

	/**
	 * Specifies whether the current sheet gridlines must be printed or not.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {bool} bPrint - Defines if cell gridlines are printed on this page or not.
	 * */
	ApiWorksheet.prototype.SetPrintGridlines = function (bPrint)	{
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
		var res =  this.worksheet.workbook.getDefinedNamesWS(this.worksheet.getId());
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
	 * @param {bool} isHidden - Defines if the range name is hidden or not.
	 * @returns {Error | true} - returns error if sName or sRef are invalid.
	 */
	ApiWorksheet.prototype.AddDefName = function (sName, sRef, isHidden) {
		return private_AddDefName(this.worksheet.workbook, sName, sRef, this.worksheet.getId(), isHidden);
	};

	Object.defineProperty(ApiWorksheet.prototype, "DefNames", {
		get: function () {
			return this.GetDefNames();
		}
	});

	/**
	 * Returns an array of ApiComment objects.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @returns {ApiComment[]}
	 */
	ApiWorksheet.prototype.GetComments = function () {
		var comments = [];
		for (var i = 0; i < this.worksheet.aComments.length; i++) {
			comments.push(new ApiComment(this.worksheet.aComments[i], this.worksheet.workbook.getWorksheet(this.Index)));
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
	 * @param {string} ScreenTip - The screen tip text.
	 * @param {string} sTextToDisplay - The link text that will be displayed on the sheet.
	 * */
	ApiWorksheet.prototype.SetHyperlink = function (sRange, sAddress, sScreenTip, sTextToDisplay)	{
		var range = new ApiRange(this.worksheet.getRange2(sRange));
		var p = /^(?:http:\/\/|https:\/\/)/;
		if (range && range.range.isOneCell() && sAddress) {
			var externalLink = sAddress.match(p) || sAddress.search(/mailto:/i) !== -1;
			if (externalLink && AscCommonExcel.getFullHyperlinkLength(sAddress) > Asc.c_nMaxHyperlinkLength) {
				return null;
			}

			this.worksheet.selectionRange.assign2(range.range.bbox);
			var  Hyperlink = new Asc.asc_CHyperlink();
			if (sScreenTip) {
				Hyperlink.asc_setText(sScreenTip);
			} else {
				Hyperlink.asc_setText(sAddress);
			}
			if (sTextToDisplay) {
				Hyperlink.asc_setTooltip(sTextToDisplay);
			}
			if (externalLink) {
				Hyperlink.asc_setHyperlinkUrl(sAddress);
			} else {
				Hyperlink.asc_setRange(sAddress);
				Hyperlink.asc_setSheet(this.GetName());
			}
			this.worksheet.workbook.oApi.wb.insertHyperlink(Hyperlink);
		}
	};

	/**
	 * Creates a chart of the specified type from the selected data range of the current sheet.
	 * <note>Please note that the horizontal and vertical offsets are calculated within the limits of the specified column and
	 * row cells only. If this value exceeds the cell width or height, another vertical/horizontal position will be set.</note>
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sDataRange - The selected cell range which will be used to get the data for the chart, formed specifically and including the sheet name.
	 * @param {bool} bInRows - Specifies whether to take the data from the rows or from the columns. If true, the data from the rows will be used.
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
		function (sDataRange, bInRows, sType, nStyleIndex, nExtX, nExtY, nFromCol, nColOffset,  nFromRow, nRowOffset) {
			var settings = new Asc.asc_ChartSettings();
			settings.type = AscFormat.ChartBuilderTypeToInternal(sType);
			settings.style = nStyleIndex;
			settings.inColumns = !bInRows;
			settings.putRange(sDataRange);
			var oChart = AscFormat.DrawingObjectsController.prototype.getChartSpace(settings);
			if(arguments.length === 8){//support old variant
				oChart.setBDeleted(false);
				oChart.setWorksheet(this.worksheet);
				oChart.addToDrawingObjects();
				oChart.setDrawingBaseCoords(arguments[4], 0, arguments[5], 0, arguments[6], 0, arguments[7], 0, 0, 0, 0, 0);
			}
			else{
				private_SetCoords(oChart, this.worksheet, nExtX, nExtY, nFromCol, nColOffset,  nFromRow, nRowOffset);
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
	ApiWorksheet.prototype.AddShape = function(sType, nWidth, nHeight, oFill, oStroke, nFromCol, nColOffset, nFromRow, nRowOffset){
		var oShape = AscFormat.builder_CreateShape(sType, nWidth/36000, nHeight/36000, oFill.UniFill, oStroke.Ln, null, this.worksheet.workbook.theme, this.worksheet.getDrawingDocument(), false, this.worksheet);
		private_SetCoords(oShape, this.worksheet, nWidth, nHeight, nFromCol, nColOffset,  nFromRow, nRowOffset);
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
	ApiWorksheet.prototype.AddImage = function(sImageSrc, nWidth, nHeight, nFromCol, nColOffset, nFromRow, nRowOffset){
		var oImage = AscFormat.DrawingObjectsController.prototype.createImage(sImageSrc, 0, 0, nWidth/36000, nHeight/36000);
		private_SetCoords(oImage, this.worksheet, nWidth, nHeight, nFromCol, nColOffset,  nFromRow, nRowOffset);
		return new ApiImage(oImage);
	};

	/**
	 * Replaces the current image with a new one.
	 * @memberof ApiWorksheet
	 * @typeofeditors ["CSE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The image width in English measure units.
	 * @param {EMU} nHeight - The image height in English measure units.
	 */
	ApiWorksheet.prototype.ReplaceCurrentImage = function(sImageUrl, Width, Height){

		var oWorksheet = Asc['editor'].wb.getWorksheet();
		if(oWorksheet && oWorksheet.objectRender && oWorksheet.objectRender.controller){

			var oController = oWorksheet.objectRender.controller;
			var _w = Width/36000.0;
			var _h = Height/36000.0;
			var oImage = oController.createImage(sImageUrl, 0, 0, _w, _h);
			oImage.setWorksheet(oWorksheet.model);
			var selectedObjects, spTree;
			if(oController.selection.groupSelection){
				selectedObjects = oController.selection.groupSelection.selectedObjects;
			}
			else{
				selectedObjects = oController.selectedObjects;
			}
			if(selectedObjects.length > 0){
				if(selectedObjects[0].group){
					spTree = selectedObjects[0].group.spTree;
				}
				else{
					spTree = oController.getDrawingArray();
				}

				for(var i = 0; i < spTree.length; ++i){
					if(spTree[i] === selectedObjects[0]){
						if(spTree[i].getObjectType() === AscDFH.historyitem_type_ImageShape){
							spTree[i].setBlipFill(AscFormat.CreateBlipFillRasterImageId(sImageUrl));
							if(selectedObjects[0].group){
								oController.selection.groupSelection.resetInternalSelection();
								selectedObjects[0].group.selectObject(spTree[i], 0);
							}
							else{
								oController.resetSelection();
								oController.selectObject(spTree[i], 0);
							}
						}
						else{
							var _xfrm = spTree[i].spPr && spTree[i].spPr.xfrm;
							var _xfrm2 = oImage.spPr.xfrm;
							if(_xfrm){
								_xfrm2.setOffX(_xfrm.offX);
								_xfrm2.setOffY(_xfrm.offY);
							}
							else{
								if(AscFormat.isRealNumber(spTree[i].x) && AscFormat.isRealNumber(spTree[i].y)){
									_xfrm2.setOffX(spTree[i].x);
									_xfrm2.setOffY(spTree[i].y);
								}
							}
							if(selectedObjects[0].group){
								var _group = selectedObjects[0].group;
								_group.removeFromSpTreeByPos(i);
								_group.addToSpTree(i, oImage);
								oImage.setGroup(_group);
								oController.selection.groupSelection.resetInternalSelection();
								_group.selectObject(oImage, 0);
							}
							else{
								var _object = spTree[i];
								_object.deleteDrawingBase();
								oImage.setBDeleted(false);
								oImage.setWorksheet(oWorksheet.model);
								oImage.addToDrawingObjects(i);
								oImage.setDrawingBaseType(AscCommon.c_oAscCellAnchorType.cellanchorAbsolute);
								oImage.setDrawingBaseCoords(0, 0, 0, 0, 0, 0, 0, 0, _object.x, _object.y, oImage.spPr.xfrm.extX, oImage.spPr.xfrm.extY);
								oImage.setDrawingBaseExt(oImage.spPr.xfrm.extX, oImage.spPr.xfrm.extY);
								oController.resetSelection();
								oController.selectObject(oImage, 0);
							}
						}
						return;
					}
				}
			}
			var cell = this.worksheet.selectionRange.activeCell;
			private_SetCoords(oImage, oWorksheet.model, Width, Height, cell ? cell.col : 0, 0,  cell ? cell.row : 0, 0, undefined);
			oController.resetSelection();
			oController.selectObject(oImage, 0);
			oWorksheet.isSelectOnShape = true;
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
	ApiRange.prototype.GetClassType = function()
	{
		return "range";
	};

	/**
	 * Returns a row number for the selected cell.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 */
	ApiRange.prototype.GetRow = function () {
		return this.range.bbox.r1;
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
		return this.range.bbox.c1;
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
		this.range.cleanAll();
	};

	/**
	 * Returns a Range object that represents the rows in the specified range. If the specified row is outside the Range object, a new Range will be returned that represents the cells between the columns of the original range in the specified row.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {number} nRow - The row number (starts counting from 1, the 0 value returns an error).
	 * @returns {ApiRange}
	 */
	ApiRange.prototype.GetRows = function (nRow) {
		if (typeof nRow === "undefined") {
			return new ApiRange(this.range.worksheet.getRange3(this.range.bbox.r1, 0, this.range.bbox.r2, AscCommon.gc_nMaxCol0));
			// return new ApiWorksheet(this.range.worksheet).GetRows();	// return all rows from current sheet
		} else {
			if (typeof nRow === "number" && nRow > 0 && nRow <= AscCommon.gc_nMaxRow0 + 1) {
				nRow--;
				if ( (nRow >= this.range.bbox.r1) && (nRow <= this.range.bbox.r2) ) {
					return new ApiRange(this.range.worksheet.getRange3(nRow, 0, nRow, AscCommon.gc_nMaxCol0));
				} else {
					return new ApiRange(this.range.worksheet.getRange3(nRow, this.range.bbox.c1, nRow, this.range.bbox.c2));
				}
			} else {
				return new Error('The nRow must be greater than 0 and less then ' + (AscCommon.gc_nMaxRow0 + 1));
			}
		}
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
	 * @returns {ApiRange}
	 */
	 ApiRange.prototype.GetCols = function (nCol) {
		if (typeof nCol === "undefined") {
			return new ApiRange(this.range.worksheet.getRange3(0, this.range.bbox.c1, AscCommon.gc_nMaxRow0, this.range.bbox.c2));
		} else {
			if (typeof nCol === "number" && nCol > 0 && nCol <= AscCommon.gc_nMaxCol0 + 1)
			{
				nCol--;
				if ( (nCol >= this.range.bbox.c1) && (nCol <= this.range.bbox.c2) ) {
					return new ApiRange(this.range.worksheet.getRange3(0, nCol, AscCommon.gc_nMaxRow0, nCol));
				}
				else {
					return new ApiRange(this.range.worksheet.getRange3(this.range.bbox.r1, nCol, this.range.bbox.r2, nCol));
				}
			} else {
				return new Error('The nCol must be greater than 0 and less then ' + (AscCommon.gc_nMaxCol0 + 1));
			}
		} 
		
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
		var bbox = this.range.bbox;
		var r1, c1, r2, c2;
		switch (direction) {
			case "xlUp":
				r1 = r2 = 0;
				c1 = c2 = bbox.c1;
				break;
			case "xlDown":
				r1 = r2 = AscCommon.gc_nMaxRow0;
				c1 = c2 = bbox.c1;
				break;
			case "xlToRight":
				r1 = r2 = bbox.r1;
				c1 = c2 = AscCommon.gc_nMaxCol0;
				break;
			case "xlToLeft":
				r1 = r2 = bbox.r1;
				c1 = c2 = 0;
				break;
			default:
				r1 = bbox.r1;
				c1 = bbox.c1;
				r2 = bbox.r2;
				c2 = bbox.c2;
				break;
		}
		return new ApiRange(this.range.worksheet.getRange3(r1, c1, r2, c2));
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
		if (row) row--;
		var bbox = this.range.bbox;
		if (typeof col !== "undefined" && typeof row !== "undefined") {
			return new ApiRange(this.range.worksheet.getRange3(row, col, row, col));
		} else if (typeof row !== "undefined") {
			var cellCount = bbox.c2 - bbox.c1 + 1; 
			var r = bbox.r1 + ((row) ?  (row / cellCount) >> 0 : row);
			var c = bbox.c1 + ((r) ? 1 : 0) + ((row) ? row % cellCount : row);
			if (r && c) c--;
			return new ApiRange(this.range.worksheet.getRange3(r, c, r, c));
		} else {
			return new ApiRange(this.range.worksheet.getRange3(bbox.r1, bbox.c1, bbox.r2, bbox.c2));
		}
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
	 * Returns the cell address.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {bool} RowAbs - Defines if the link to the row is absolute or not.
	 * @param {bool} ColAbs - Defines if the link to the column is absolute or not.
	 * @param {string} RefStyle - The reference style.
	 * @param {bool} External - Defines if the range is in the current file or not.
	 * @param {range} RelativeTo - The range which the current range is relative to.
	 * @returns {string | null} - returns address of range as string. 
	 */
	 ApiRange.prototype.GetAddress = function (RowAbs, ColAbs, RefStyle, External, RelativeTo) {
		var range = this.range.bbox;
		var isOneCell = this.range.isOneCell();
		var ws = this.range.worksheet;
		var value = ""
		var row1 = range.r1 + ( (RowAbs || RefStyle != "xlR1C1") ? 1 : 0),
			col1 = range.c1 + ( (ColAbs || RefStyle != "xlR1C1") ? 1 : 0),
			row2 = range.r2 + ( (RowAbs || RefStyle != "xlR1C1") ? 1 : 0),
			col2 = range.c2 + ( (ColAbs || RefStyle != "xlR1C1") ? 1 : 0);
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
			value = row1 + col1 + row2 + col2;
		} else {
			// xlA1 - default
			row1 = (RowAbs ? "$" : "") + row1;
			col1 = (ColAbs ? "$" : "") + AscCommon.g_oCellAddressUtils.colnumToColstr(col1);
			row2 = isOneCell ? "" : ( (RowAbs ? "$" : "") + row2);
			col2 = isOneCell ? "" : ( (ColAbs ? ":$" : ":") + AscCommon.g_oCellAddressUtils.colnumToColstr(col2) );
			value = col1 + row1 + col2 + row2;
		}
		return (External) ? '[' + ws.workbook.oApi.DocInfo.Title + ']' + AscCommon.parserHelp.get3DRef(ws.sName, value) : value;
};

	/**
	 * Returns the rows or columns count.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 */
	ApiRange.prototype.GetCount = function () {
		var range = this.range.bbox;
		var	count;
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
	 * Returns a value of the first cell in the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 */
	ApiRange.prototype.GetValue = function () {
		return this.range.getValue();
	};

	/**
	 * Sets a value to the current cell or cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string | bool | number | Array[] | Array[][]} data - The general value for the cell or cell range.
	 * @return {bool} - returns false if such a range does not exist.
	 */
	ApiRange.prototype.SetValue = function (data) {
		if (!this.range)
			return false;
		
		if (Array.isArray(data)) {
			var checkDepth = function(x) { return Array.isArray(x) ? 1 + Math.max.apply(this, x.map(checkDepth)) : 0;};
			var maxDepth = checkDepth(data);
			if (maxDepth <= 2) {
				if (this.range.isOneCell()) {
					data = maxDepth == 1 ? data[0] : data[0][0];
				} else {
					var bbox = this.range.bbox;
					var nCol = Math.min( (bbox.c2 - bbox.c1 + 1), data.length);
					var nRow = bbox.r2 - bbox.r1 + 1;
					for (var i = 0; i < nCol; i++) {
						var cRow = (nRow > data[i].length) ? data[i].length : nRow;
						for (var k = 0; k < cRow; k++) {
							var cell = this.range.worksheet.getRange3( (bbox.r1 + k), (bbox.c1 + i), (bbox.r1 + k), (bbox.c1 + i) );
							var value = checkFormat( ( (maxDepth == 1) ? data[i] : data[i][k] ) );
							cell.setValue(value.toString());
							if (value.type === AscCommonExcel.cElementType.number)
								cell.setNumFormat(AscCommon.getShortDateFormat());
						}
					}
					// ToDo update range in setValue
					var worksheet = this.range.worksheet;
					worksheet.workbook.handlers.trigger("cleanCellCache", worksheet.getId(), [this.range.bbox], true);
					return true;
				}
			}
		}
		data = checkFormat(data);
		this.range.setValue(data.toString());
		if (data.type === AscCommonExcel.cElementType.number)
			this.SetNumberFormat(AscCommon.getShortDateFormat());

		// ToDo update range in setValue
		var worksheet = this.range.worksheet;
		worksheet.workbook.handlers.trigger("cleanCellCache", worksheet.getId(), [this.range.bbox], true);
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
	 * Returns a formula of the first cell in the current range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @return {string} - return Value2 property (value without format) if formula doesn't exist.  
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
	 * Returns the Value2 property (value without format) of the first cell in the current range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @return {string} 
	 */
	ApiRange.prototype.GetValue2 = function () {
		return this.range.getValueWithoutFormat();
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
	 * Returns the text of the first cell in the current range.
	 * @typeofeditors ["CSE"]
	 * @memberof ApiRange
	 * @return {string} 
	 */
	ApiRange.prototype.GetText = function () {
		return this.range.getValueWithFormat();
	};

	Object.defineProperty(ApiRange.prototype, "Text", {
		get: function () {
			return this.range.getValueWithFormat();
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
	 * Returns the value hiding property.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {bool} - returns true if the values in the range specified are hidden.
	 */
	ApiRange.prototype.GetHidden = function () {
		var range = this.range;
		var worksheet = range.worksheet;
		var bbox = range.bbox;
		switch (bbox.getType()) {
			case 2:		
				return worksheet.getColHidden(bbox.c1);	

			case 3:
				return worksheet.getRowHidden(bbox.r1);				

			default:
				return false;
		}
	};
	/**
	 * Sets the value hiding property.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {bool} isHidden - Specifies if the values in the current range are hidden or not.
	 */
	ApiRange.prototype.SetHidden = function (isHidden) {
		var range = this.range;
		var worksheet = range.worksheet;
		var bbox = range.bbox;
		switch (bbox.getType()) {
			case 2:		
				worksheet.setColHidden(isHidden, bbox.c1, bbox.c2);	
				break;

			case 3:
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
		this.range.worksheet.setRowHeight(nHeight, this.range.bbox.r1, this.range.bbox.r2, false);
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
	 * @returns {bool} - return false if sAligment doesn't exist.
	 */
	ApiRange.prototype.SetAlignVertical = function (sAligment) {
		switch(sAligment)
		{
			case "center":
			{
				this.range.setAlignVertical(Asc.c_oAscVAlign.Center);
				break;
			}
			case "bottom":
			{
				this.range.setAlignVertical(Asc.c_oAscVAlign.Bottom);
				break;
			}
			case "top":
			{
				this.range.setAlignVertical(Asc.c_oAscVAlign.Top);
				break;
			}
			case "distributed":
			{
				this.range.setAlignVertical(Asc.c_oAscVAlign.Dist);
				break;
			}
			case "justify":
			{
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
	 * @returns {bool} - return false if sAligment doesn't exist.
	 */
	ApiRange.prototype.SetAlignHorizontal = function (sAlignment) {
		switch(sAlignment)
		{
			case "left":
			{
				this.range.setAlignHorizontal(AscCommon.align_Left);
				break;
			}
			case "right":
			{
				this.range.setAlignHorizontal(AscCommon.align_Right);
				break;
			}
			case "justify":
			{
				this.range.setAlignHorizontal(AscCommon.align_Justify);
				break;
			}
			case "center":
			{
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
	 * @param {bool} isBold - Specifies that the contents of the current cell / cell range are displayed bold.
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
	 * @param {bool} isItalic - Specifies that the contents of the current cell / cell range are displayed italicized.
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
	 * @param {bool} isStrikeout - Specifies if the contents of the current cell / cell range are displayed struck through.
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
	 * @param {bool} isWrap - Specifies if the words in the cell will be wrapped to fit the cell size.
	 */
	ApiRange.prototype.SetWrap = function (isWrap) {
		this.range.setWrap(!!isWrap);
	};

	/**
	 * Returns the information about the wrapping cell style.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @returns {bool}
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
	 * Specifies whether a number in the cell should be treated like number, currency, date, time, etc. or just like text.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {string} sFormat - Specifies the mask applied to the number in the cell.
	 */
	ApiRange.prototype.SetNumberFormat = function (sFormat) {
		this.range.setNumFormat(sFormat);
	};
	Object.defineProperty(ApiRange.prototype, "NumberFormat", {
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
	 * @param {bool} isAcross - When set to <b>true</b>, the cells within the selected range will be merged along the rows,
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
	 * @returns {ApiRange}
	 */
	Object.defineProperty(ApiRange.prototype, "MergeArea", {
		get: function () {
			if (this.range.isOneCell()) {
				var bb = this.range.hasMerged();
				return new ApiRange((bb) ? AscCommonExcel.Range.prototype.createFromBBox(this.range.worksheet, bb) : this.range);
			} else {
				return new Error('Range must be is one cell.');
			}
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
	 * @returns {bool} - returns false if comment can't be added.
	 */
	ApiRange.prototype.AddComment = function (sText) {
		var ws = Asc['editor'].wb.getWorksheet(this.range.getWorksheet().getIndex());
		if (ws) {
			var comment = new Asc.asc_CCommentData();
			comment.sText = sText;
			comment.nCol = this.range.bbox.c1;
			comment.nRow = this.range.bbox.r1;
			comment.bDocument = false;
			ws.cellCommentator.addComment(comment, true);

			return true;
		}

		return false;
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
		return new ApiComment(ws.cellCommentator.getComment(this.range.bbox.c1, this.range.bbox.r1, false), ws);
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
			newSelection.assign2(this.range.bbox);
			newSelection.Select();
		}
	};

	/**
	 * Returns the current range angle.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @return {Angle}
	 */
	ApiRange.prototype.GetOrientation = function() {
	  return this.range.getAngle();
	};

	/**
	 * Sets an angle to the current cell range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {Angle} angle - Specifies the range angle.
	 */
	ApiRange.prototype.SetOrientation = function(angle) {
        switch(angle) {
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

		var getSortLevel = function(_key, _order) {
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
		if(tables && tables.length) {
			obj = tables[0];
		} else if(ws.AutoFilter && ws.AutoFilter.Ref && ws.AutoFilter.Ref.intersection(range)) {
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
	 * @param {?String} shift - Specifies how to shift cells to replace the deleted cells ("up", "left").
	 */
	ApiRange.prototype.Delete = function(shift) {
		if (shift && typeof Shift == "string") {
			shift = shift.toLocaleLowerCase();
		} else {
			var bbox = this.range.bbox;
			var rows = bbox.r2 - bbox.r1 + 1;
			var cols = bbox.c2 - bbox.c1 + 1;
			shift = (rows <= cols) ? "up" : "left";
		}
		if (shift == "up")
			this.range.deleteCellsShiftUp();
		else
			this.range.deleteCellsShiftLeft()
	};

	/**
	 * Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {?String} shift - Specifies which way to shift the cells ("right", "down").
	 */
	 ApiRange.prototype.Insert = function(shift) {
		if (shift && typeof Shift == "string") {
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
	ApiRange.prototype.AutoFit = function(bRows, bCols) {
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
	ApiRange.prototype.GetAreas = function() {
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
	ApiRange.prototype.Copy = function(destination) {
		if (destination && destination instanceof ApiRange) {
			var cols = this.GetCols().Count - 1;
			var rows = this.GetRows().Count - 1;
			var bbox = destination.range.bbox;
			var range = destination.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols) );
			this.range.move(range.bbox, true, destination.range.worksheet);
		} else {
			return new Error ("Invalid destination");
		}
	};
	
	/**
	 * Pastes the Range object to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CSE"]
	 * @param {ApiRange} rangeFrom - Specifies the range to be pasted to the current range
	 */
	ApiRange.prototype.Paste = function(rangeFrom) {
		if (rangeFrom && rangeFrom instanceof ApiRange) {
			var cols = rangeFrom.GetCols().Count - 1;
			var rows = rangeFrom.GetRows().Count - 1;
			var bbox = this.range.bbox;
			var range = this.range.worksheet.getRange3(bbox.r1, bbox.c1, (bbox.r1 + rows), (bbox.c1 + cols) );
			rangeFrom.range.move(range.bbox, true, range.worksheet);
		} else {
			return new Error ("Invalid range");
		}
	};

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
	ApiDrawing.prototype.GetClassType = function()
	{
		return "drawing";
	};

	/**
	 * Sets a size of the object (image, shape, chart) bounding box.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CSE"]
	 * @param {EMU} nWidth - The object width measured in English measure units.
	 * @param {EMU} nHeight - The object height measured in English measure units.
	 */
	ApiDrawing.prototype.SetSize = function(nWidth, nHeight)
	{
		var fWidth = nWidth/36000.0;
		var fHeight = nHeight/36000.0;
		if(this.Drawing && this.Drawing.spPr && this.Drawing.spPr.xfrm)
		{
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
	ApiDrawing.prototype.SetPosition = function(nFromCol, nColOffset, nFromRow, nRowOffset){
		var extX = null, extY = null;
		if(this.Drawing.drawingBase){
			if(this.Drawing.drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorOneCell ||
				this.Drawing.drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorAbsolute){
				extX = this.Drawing.drawingBase.ext.cx;
				extY = this.Drawing.drawingBase.ext.cy;
			}
		}
		if(!AscFormat.isRealNumber(extX) || !AscFormat.isRealNumber(extY)){
			if(this.Drawing.spPr && this.Drawing.spPr.xfrm){
				extX = this.Drawing.spPr.xfrm.extX;
				extY = this.Drawing.spPr.xfrm.extY;
			}
			else{
				extX = 5;
				extY = 5;
			}
		}
		this.Drawing.setDrawingBaseType(AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
		this.Drawing.setDrawingBaseCoords(nFromCol, nColOffset/36000.0, nFromRow, nRowOffset/36000.0, 0, 0, 0, 0, 0, 0, extX, extY);
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
	ApiImage.prototype.GetClassType = function()
	{
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
	ApiShape.prototype.GetClassType = function()
	{
		return "shape";
	};

	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted. 
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiShape.prototype.GetContent = function()
	{
		var oApi = Asc["editor"];
		if(oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content)
		{
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
	ApiShape.prototype.GetDocContent = function()
	{
		var oApi = Asc["editor"];
		if(oApi && this.Drawing && this.Drawing.txBody && this.Drawing.txBody.content)
		{
			return oApi.private_CreateApiDocContent(this.Drawing.txBody.content);
		}
		return null;
	};

	/**
	 * Sets the vertical alignment to the shape content where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CSE"]
	 * @param {"top" | "center" | "bottom" } sVerticalAlign - The vertical alignment type for the shape inner contents.
	 * @returns {bool} - returns false if shape or aligment doesn't exist. 
	 */
	ApiShape.prototype.SetVerticalTextAlign = function(sVerticalAlign)
	{
		if(this.Shape)
		{
			switch(sVerticalAlign)
			{
				case "top":
				{
					this.Shape.setVerticalAlign(4);
					break;
				}
				case "center":
				{
					this.Shape.setVerticalAlign(1);
					break;
				}
				case "bottom":
				{
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
	ApiChart.prototype.GetClassType = function()
	{
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
	ApiChart.prototype.SetTitle = function (sTitle, nFontSize, bIsBold)
	{
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
	ApiChart.prototype.SetHorAxisTitle = function (sTitle, nFontSize, bIsBold)
	{
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
	ApiChart.prototype.SetVerAxisTitle = function (sTitle, nFontSize, bIsBold)
	{
		AscFormat.builder_SetChartVertAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};


	/**
	 * Specifies the direction of the data displayed on the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {bool} bIsMinMax - The <code>true</code> value sets the normal data direction for the vertical axis (from minimum to maximum).
	 * The <code>false</code> value sets the inverted data direction for the vertical axis (from maximum to minimum).
	 * */
	ApiChart.prototype.SetVerAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartVertAxisOrientation(this.Chart, bIsMinMax);
	};


	/**
	 * Specifies the major tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the minor tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the major tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the minor tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies the direction of the data displayed on the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {bool} bIsMinMax - The <code>true</code> value sets the normal data direction for the horizontal axis
	 * (from minimum to maximum). The <code>false</code> value sets the inverted data direction for the horizontal axis (from maximum to minimum).
	 * */
	ApiChart.prototype.SetHorAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartHorAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specifies the chart legend position.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {"left" | "top" | "right" | "bottom" | "none"} sLegendPos - The position of the chart legend inside the chart window.
	 * */
	ApiChart.prototype.SetLegendPos = function(sLegendPos)
	{
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
	ApiChart.prototype.SetLegendFontSize = function(nFontSize)
	{
		AscFormat.builder_SetLegendFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies which chart data labels are shown for the chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {bool} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {bool} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {bool} bShowVal - Whether to show or hide the chart data values.
	 * @param {bool} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowDataLabels = function(bShowSerName, bShowCatName, bShowVal, bShowPercent)
	{
		AscFormat.builder_SetShowDataLabels(this.Chart, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};

	/**
	 * Spicifies the show options for the data labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {number} nSeriesIndex - The series index from the array of the data used to build the chart from.
	 * @param {number} nPointIndex - The point index from this series.
	 * @param {bool} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {bool} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {bool} bShowVal - Whether to show or hide the chart data values.
	 * @param {bool} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowPointDataLabel = function(nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent)
	{
		AscFormat.builder_SetShowPointDataLabel(this.Chart, nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};

	/**
	 * Sets the possible values for the position of the chart tick labels in relation to the main vertical label or the chart data values.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart vertical tick labels.
	 * */
	ApiChart.prototype.SetVertAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartVertAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Sets the possible values for the position of the chart tick labels in relation to the main horizontal label or the chart data values.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart horizontal tick labels.
	 * */
	ApiChart.prototype.SetHorAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartHorAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Specifies the visual properties of the major vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specifies the visual properties of the minor vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specifies the visual properties of the major horizontal gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specifies the visual properties of the minor vertical gridline.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 */
	ApiChart.prototype.SetMinorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specifies the font size to the horizontal axis labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	*/
	ApiChart.prototype.SetHorAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetHorAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies the font size to the vertical axis labels.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	*/
	ApiChart.prototype.SetVertAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetVerAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Applies a set of visual settings to the chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CSE"]
	 * @param {number} nStyleIndex - The style index that will be set to the current chart.
	*/
	ApiChart.prototype.ApplyChartStyle = function(nStyleIndex){
		if(this.Chart){
			var chart = this.Chart.chart;
			var plot_area = chart.plotArea;
			var oCurChartSettings = AscFormat.DrawingObjectsController.prototype.getPropsFromChart.call(AscFormat.DrawingObjectsController.prototype, this.Chart);
			var _cur_type = oCurChartSettings.type;
			if(AscCommon.g_oChartStyles[_cur_type] && AscCommon.g_oChartStyles[_cur_type][nStyleIndex]){
				this.Chart.applyChartStyleByIds(AscCommon.g_oChartStyles[_cur_type][nStyleIndex])
			}
		}
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
	 * @returns {Error | true} - returns error if sName is invalid.
	 */
	ApiName.prototype.SetName = function (sName) {
		if (!sName || typeof sName !== 'string' || !this.DefName) {
			return new Error('Invalid name or Defname is undefined.');
		}
		var res = this.DefName.wb.checkDefName(sName);
		if (!res.status) {
			return new Error('Invalid name.'); // invalid name
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
	 * @param {string} sRef	- The range reference which must contain the sheet name, followed by sign ! and a range of cells. 
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
	 * Returns the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {string}
	 */
	ApiComment.prototype.GetText = function () {
		return this.Comment.asc_getText();
	};
	Object.defineProperty(ApiComment.prototype, "Text", {
		get: function () {
			return this.GetText();
		}
	});

	/**
	 * Deletes the ApiComment object.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 */
	ApiComment.prototype.Delete = function () {
		this.WS.cellCommentator.removeComment(this.Comment.asc_getId());
	};

	/**
	 * Returns a type of the ApiComment class.
	 * @memberof ApiComment
	 * @typeofeditors ["CSE"]
	 * @returns {number}
	 */
	 ApiComment.prototype.GetClassType = function () {
		return this.Comment.getType();
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
	ApiWorksheet.prototype["GetDefNames"] = ApiWorksheet.prototype.GetDefNames;
	ApiWorksheet.prototype["GetDefName"] = ApiWorksheet.prototype.GetDefName;
	ApiWorksheet.prototype["AddDefName"] = ApiWorksheet.prototype.AddDefName;
	ApiWorksheet.prototype["GetComments"] = ApiWorksheet.prototype.GetComments;
	ApiWorksheet.prototype["Delete"] = ApiWorksheet.prototype.Delete;
	ApiWorksheet.prototype["SetHyperlink"] = ApiWorksheet.prototype.SetHyperlink;
	ApiWorksheet.prototype["AddChart"] = ApiWorksheet.prototype.AddChart;
	ApiWorksheet.prototype["AddShape"] = ApiWorksheet.prototype.AddShape;
	ApiWorksheet.prototype["AddImage"] = ApiWorksheet.prototype.AddImage;
	ApiWorksheet.prototype["ReplaceCurrentImage"] = ApiWorksheet.prototype.ReplaceCurrentImage;

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


	ApiDrawing.prototype["GetClassType"]               =  ApiDrawing.prototype.GetClassType;
	ApiDrawing.prototype["SetSize"]                    =  ApiDrawing.prototype.SetSize;
	ApiDrawing.prototype["SetPosition"]                =  ApiDrawing.prototype.SetPosition;

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



	ApiChart.prototype["SetMajorVerticalGridlines"]  =  ApiChart.prototype.SetMajorVerticalGridlines;
	ApiChart.prototype["SetMinorVerticalGridlines"]  =  ApiChart.prototype.SetMinorVerticalGridlines;
	ApiChart.prototype["SetMajorHorizontalGridlines"]  =  ApiChart.prototype.SetMajorHorizontalGridlines;
	ApiChart.prototype["SetMinorHorizontalGridlines"]  =  ApiChart.prototype.SetMinorHorizontalGridlines;
	ApiChart.prototype["SetHorAxisLablesFontSize"]   =  ApiChart.prototype.SetHorAxisLablesFontSize;
	ApiChart.prototype["SetVertAxisLablesFontSize"]  =  ApiChart.prototype.SetVertAxisLablesFontSize;
	ApiChart.prototype["ApplyChartStyle"]            =  ApiChart.prototype.ApplyChartStyle;


	ApiColor.prototype["GetClassType"]                 =  ApiColor.prototype.GetClassType;


	ApiName.prototype["GetName"]                 =  ApiName.prototype.GetName;
	ApiName.prototype["SetName"]                 =  ApiName.prototype.SetName;
	ApiName.prototype["Delete"]                  =  ApiName.prototype.Delete;
	ApiName.prototype["GetRefersTo"]             =  ApiName.prototype.GetRefersTo;
	ApiName.prototype["SetRefersTo"]             =  ApiName.prototype.SetRefersTo;
	ApiName.prototype["GetRefersToRange"]        =  ApiName.prototype.GetRefersToRange;


	ApiComment.prototype["GetText"]              =  ApiComment.prototype.GetText;
	ApiComment.prototype["Delete"]               =  ApiComment.prototype.Delete;
	ApiComment.prototype["GetClassType"]         =  ApiComment.prototype.GetClassType;
	

	ApiAreas.prototype["GetCount"]               = ApiAreas.prototype.GetCount;
	ApiAreas.prototype["GetItem"]                = ApiAreas.prototype.GetItem;
	ApiAreas.prototype["GetParent"]              = ApiAreas.prototype.GetParent;


	function private_SetCoords(oDrawing, oWorksheet, nExtX, nExtY, nFromCol, nColOffset,  nFromRow, nRowOffset, pos){
		oDrawing.x = 0;
		oDrawing.y = 0;
		oDrawing.extX = 0;
		oDrawing.extY = 0;
		AscFormat.CheckSpPrXfrm(oDrawing);
		oDrawing.spPr.xfrm.setExtX(nExtX/36000.0);
		oDrawing.spPr.xfrm.setExtY(nExtY/36000.0);
		oDrawing.setBDeleted(false);
		oDrawing.setWorksheet(oWorksheet);
		oDrawing.addToDrawingObjects(pos);
		oDrawing.setDrawingBaseType(AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
		oDrawing.setDrawingBaseCoords(nFromCol, nColOffset/36000.0, nFromRow, nRowOffset/36000.0, 0, 0, 0, 0, 0, 0, 0, 0);
		oDrawing.setDrawingBaseExt(nExtX/36000.0, nExtY/36000.0);
	}

	function private_MakeBorder(lineStyle, color) {
		var border = new AscCommonExcel.BorderProp();
		switch (lineStyle) {
			case 'Double':
				border.setStyle(AscCommon.c_oAscBorderStyles.Double);
				break;
			case 'Hair':
				border.setStyle(AscCommon.c_oAscBorderStyles.Hair);
				break;
			case 'DashDotDot':
				border.setStyle(AscCommon.c_oAscBorderStyles.DashDotDot);
				break;
			case 'DashDot':
				border.setStyle(AscCommon.c_oAscBorderStyles.DashDot);
				break;
			case 'Dotted':
				border.setStyle(AscCommon.c_oAscBorderStyles.Dotted);
				break;
			case 'Dashed':
				border.setStyle(AscCommon.c_oAscBorderStyles.Dashed);
				break;
			case 'Thin':
				border.setStyle(AscCommon.c_oAscBorderStyles.Thin);
				break;
			case 'MediumDashDotDot':
				border.setStyle(AscCommon.c_oAscBorderStyles.MediumDashDotDot);
				break;
			case 'SlantDashDot':
				border.setStyle(AscCommon.c_oAscBorderStyles.SlantDashDot);
				break;
			case 'MediumDashDot':
				border.setStyle(AscCommon.c_oAscBorderStyles.MediumDashDot);
				break;
			case 'MediumDashed':
				border.setStyle(AscCommon.c_oAscBorderStyles.MediumDashed);
				break;
			case 'Medium':
				border.setStyle(AscCommon.c_oAscBorderStyles.Medium);
				break;
			case 'Thick':
				border.setStyle(AscCommon.c_oAscBorderStyles.Thick);
				break;
			case 'None':
			default:
				border.setStyle(AscCommon.c_oAscBorderStyles.None);
				break;
		}

		if (color) {
			border.c = color.color;
		}
		return border;
	}

	function private_AddDefName(wb, name, ref, sheetId, hidden) {
		var res = wb.checkDefName(name);
		if (!res.status) {
			return new Error('Invalid name.');
		}
		res = wb.oApi.asc_checkDataRange(Asc.c_oAscSelectionDialogType.Chart, ref, false);
		if (res === Asc.c_oAscError.ID.DataRangeError) {
			return new Error('Invalid range.');
		}
		if (sheetId) {
			sheetId = (wb.getWorksheetById(sheetId)) ? sheetId : undefined;
		}
		wb.addDefName(name, ref, sheetId, hidden, false)

		return true;
	}

}(window, null));
