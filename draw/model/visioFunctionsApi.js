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

(function(window, document)
{

	/**
	 * accepts visio shadow arguments and common arguments, return OnlyOffice api objects of different types.
	 * So for foreground color it return Unifill and for stroke too. May cause problems
	 * https://learn.microsoft.com/ru-ru/office/client-developer/visio/themeval-function.
	 * For font color return CUniColor.
	 * themeValue - if no cell passed (cell is ignored)
	 * @param {Cell_Type} cell
	 * @param {Shape_Type} shape
	 * @param {Page_Type} pageInfo
	 * @param {CTheme[]} themes
	 * @param {string?} themeValue
	 * @param {string?} defaultValue
	 * @return {CUniFill | CUniColor | any}
	 */
	function themeval(cell, shape, pageInfo, themes, themeValue, defaultValue) {
		// https://visualsignals.typepad.co.uk/vislog/2013/05/visio-2013-themes-in-the-shapesheet-part-2.html

		// if cell value is not 'Themed' waiting number representing color otherwise
		// if QuickStyle cell is from 100 to 106 or from 200 to 206 using VariationColorIndex cell value and
		// 	QuickStyle cell value % 100 get color from theme
		// if QuickStyle cell value is not in that range it is smt like from 0 to 7 representing theme colors like:
		//  dk1, lt1, accent1, ...

		// TODO handle multiple themes by schemeEnum tags from theme xml and theme properties cells
		// https://disk.yandex.ru/d/YZEevC0lUeUBfQ

		/** @type {CUniColor} */
		let calculatedColor = null;
		/** @type {CUniFill | CUniColor} */
		let result = null;

		let cellValue = cell && cell.v; // formula?
		let cellName = cell && cell.n;

		let quickStyleCellName;
		let quickStyleModifiersCellName;
		let getModifiersMethod;
		let variationStyleIndexVariable;

		if (themeValue === "LineColor") {
			cellName = "LineColor";
			cellValue = "Themed";
		} else if (themeValue === "FillColor") {
			// cellName = "FillForegnd";
			cellValue = "Themed";
		} else if (themeValue === "TextColor") {
			cellName = "Color";
			cellValue = "Themed";
		}

		let initialDefaultValue = null;

		if (cellName === "LineColor") {
			quickStyleCellName = "QuickStyleLineColor";
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLnStyle;
			variationStyleIndexVariable = "lineIdx";

			initialDefaultValue = AscFormat.CreateUnfilFromRGB(0,0,0);
		} else if (cellName === "Color") {
			// Text color
			quickStyleCellName = "QuickStyleFontColor";
			quickStyleModifiersCellName = "QuickStyleFontMatrix";
			getModifiersMethod = themes[0].getFontStyle;
			variationStyleIndexVariable = "fontIdx";

			initialDefaultValue =  AscFormat.CreateUnfilFromRGB(0,0,0).fill.color;
		} else if (cellName === "FillForegnd" || cellName === "FillBkgnd") {
			quickStyleCellName = "QuickStyleFillColor";
			quickStyleModifiersCellName = "QuickStyleFillMatrix";
			getModifiersMethod = themes[0].getFillStyle;
			variationStyleIndexVariable = "fillIdx";

			if (cellName === "FillForegnd") {
				initialDefaultValue =  AscFormat.CreateUnfilFromRGB(255,255,255);
			} else if (cellName === "FillBkgnd") {
				initialDefaultValue =  AscFormat.CreateUnfilFromRGB(0,0,0);
			}
		} else {
			console.log("themeval argument error. cell name is unknown. return null.");
			return null;
		}

		// lets define if shape is connector
		let isConnectorShape = shape.getCellNumberValue("EndArrow") !== 0
			|| shape.getCellNumberValue("BeginArrow") !== 0;

		// find theme index
		// ! Because now we only calculate colors lets find theme by
		// ext uri="{2703A3B3-D2E1-43D9-8057-6E9D74E0F44A}" clrScheme extension schemeEnum
		// which is sometimes different from ext uri="{D75FF423-9257-4291-A4FE-1B2448832E17} - themeSchemeSchemeEnum
		// and pick ColorSchemeIndex instead of ThemeIndex cell
		// upd: if connector styles are used lets use ConnectorSchemeIndex cell and
		// ext uri="{D75FF423-9257-4291-A4FE-1B2448832E17} - themeSchemeSchemeEnum to find theme
		let themeIndex = 0; // zero index means no theme - use default values
		let themeScopeCellName = isConnectorShape ? "ConnectorSchemeIndex" : "ColorSchemeIndex";
		let shapeColorSchemeThemeIndex = shape.getCellNumberValue(themeScopeCellName);
		if (isNaN(shapeColorSchemeThemeIndex) || shapeColorSchemeThemeIndex === null) {
			shapeColorSchemeThemeIndex = 0; // zero index means no theme
		}
		if (shapeColorSchemeThemeIndex === 65534) {
			let pageThemeIndexCell = pageInfo.pageSheet.elements.find(function (el) {
				return el.n === themeScopeCellName;
			});
			if (pageThemeIndexCell !== undefined) {
				let pageThemeIndex = Number(pageThemeIndexCell.v);
				if (!isNaN(pageThemeIndex)) {
					themeIndex = pageThemeIndex;
				} else {
					console.log("pageThemeIndex was not parsed");
				}
			} else {
				// it's ok sometimes
				// console.log("pageThemeIndexCell not found");
			}
		} else {
			themeIndex = shapeColorSchemeThemeIndex;
		}


		// if THEMEVAL was called with themeValue (argument like "FillColor") even if themeIndex is 0 we should
		// use any theme otherwise if no themeValue argument was passed and 0 themeIndex is used we should return
		// default value
		// see colored rectangle in that file https://disk.yandex.ru/d/IzxVtx0a7GqbQA
		let theme = themes[0];
		if ((themeValue === null || themeValue === undefined) && themeIndex === 0) {
			return initialDefaultValue;
		}
		if (themeIndex === 0) {
			// use themes[0] for THEMEVAL()
			theme = themes[0];
		} else {
			// find theme by themeIndex
			// theme = themes.find(function (theme) {
			// 	let themeEnum = Number(theme.themeElements.themeExt.themeSchemeSchemeEnum);
			// 	return themeEnum === themeIndex;
			// });
			theme = themes.find(function (theme) {
				// if search by theme index - theme.themeElements.themeExt.themeSchemeSchemeEnum
				let findThemeByElement = isConnectorShape?
					theme.themeElements.themeExt.themeSchemeSchemeEnum :
					theme.themeElements.clrScheme.clrSchemeExtLst.schemeEnum;
				let clrSchemeThemeEnum = Number(findThemeByElement);
				return clrSchemeThemeEnum === themeIndex;
			});
			if (theme === undefined) {
				console.log("Theme was not found by theme enum in themes. using themes[0]");
				theme = themes[0];
			}
		}


		let quickStyleColor = shape.getCellNumberValue(quickStyleCellName);
		let quickStyleMatrix = shape.getCellNumberValue(quickStyleModifiersCellName);
		// get color using "VariationColorIndex" cell and quickStyleColor cell
		if (!isNaN(quickStyleColor)) {
			if (100 <= quickStyleColor && quickStyleColor <= 106 ||
				(200 <= quickStyleColor && quickStyleColor <= 206)) {
				//todo 200-206?
				let variationColorIndexCell = shape.getCell("VariationColorIndex");
				let variationColorIndex = 0;
				if (variationColorIndexCell) {
					variationColorIndex = parseInt(variationColorIndexCell.v);
				}
				if (!isNaN(variationColorIndex)) {
					if (65534 === variationColorIndex) {
						variationColorIndex = 0;
						let pageVariationColorIndexCell = pageInfo.pageSheet.elements.find(function (el) {
							return el.n === "VariationColorIndex";
						});
						if (pageVariationColorIndexCell !== undefined) {
							let pageVariationColorIndexIndex = Number(pageVariationColorIndexCell.v);
							if (!isNaN(pageVariationColorIndexIndex)) {
								variationColorIndex = pageVariationColorIndexIndex;
							} else {
								console.log("pageVariationColorIndexIndex was not parsed");
							}
						} else {
							// console.log("pageVariationColorIndexIndex not found");
						}
					}
					calculatedColor = theme.getVariationClrSchemeColor(variationColorIndex,
						quickStyleColor % 100);
				}
			} else {
				switch(quickStyleColor) {
					case 0:
						calculatedColor = AscFormat.builder_CreateSchemeColor("dk1");
						break;
					case 1:
						calculatedColor = AscFormat.builder_CreateSchemeColor("lt1");
						break;
					case 2:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent1");
						break;
					case 3:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent2");
						break;
					case 4:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent3");
						break;
					case 5:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent4");
						break;
					case 6:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent5");
						break;
					case 7:
						calculatedColor = AscFormat.builder_CreateSchemeColor("accent6");
						break;
					case 8:
						//todo
						break;
					default:
						break;
				}
			}
		}
		// add matrix modifiers consider color and cells: "VariationStyleIndex" and quickStyleModifiersCellName
		if (!isNaN(quickStyleMatrix)) {
			let getMedifiersResult = null;
			if (0 === quickStyleMatrix) {
				//todo
			} else if (1 <= quickStyleMatrix && quickStyleMatrix <= 6) {
				getMedifiersResult = getModifiersMethod.call(theme, quickStyleMatrix, calculatedColor, isConnectorShape);
			} else if (100 <= quickStyleMatrix && quickStyleMatrix <= 103) {
				let variationStyleIndexCell = shape.getCell("VariationStyleIndex");
				let variationStyleIndex = 0;
				if (variationStyleIndexCell) {
					variationStyleIndex = parseInt(variationStyleIndexCell.v);
				}
				if (!isNaN(variationStyleIndex)) {
					if (65534 === variationStyleIndex) {
						variationStyleIndex = 0;
						let pageVariationStyleIndexCell = pageInfo.pageSheet.elements.find(function (el) {
							return el.n === "VariationStyleIndex";
						});
						if (pageVariationStyleIndexCell !== undefined) {
							let pageVariationStyleIndexIndex = Number(pageVariationStyleIndexCell.v);
							if (!isNaN(pageVariationStyleIndexIndex)) {
								variationStyleIndex = pageVariationStyleIndexIndex;
							} else {
								console.log("pageVariationStyleIndexIndex was not parsed");
							}
						} else {
							// console.log("pageVariationStyleIndexIndex not found");
						}
					}
					let varStyle = theme.getVariationStyleScheme(variationStyleIndex,
						quickStyleMatrix % 100);
					if (varStyle && null !== varStyle[variationStyleIndexVariable]) {
						let styleId = varStyle[variationStyleIndexVariable];
						getMedifiersResult = getModifiersMethod.call(theme, styleId, calculatedColor, isConnectorShape);

					}
				}
			}

			// getModifiersMethod return not only
			// uniFill, so we narrow down the range of returns
			if (quickStyleCellName === "QuickStyleLineColor") {
				result = getMedifiersResult && getMedifiersResult.Fill;
			} else if (quickStyleCellName === "QuickStyleFontColor") {
				// and it is color
				result = getMedifiersResult && getMedifiersResult.fontPropsObject.color;
			} else if (quickStyleCellName === "QuickStyleFillColor") {
				//leave result because it is fill
				result = getMedifiersResult;
			} else {
				console.log("Error in themeval. result is not changed to appropriate type or quickStyleCellName is not set.");
			}
		}

		/**
		 * calculate color on theme. for quickStyleVariation to consider real color
		 * @param {CUniColor | CUniFill} color
		 * @param {CTheme} theme
		 */
		function calculateOnTheme(color, theme) {
			if (color.constructor.name === "CUniColor") {
				color.Calculate(theme);
			} else if (color.constructor.name === "CUniFill") {
				color.calculate(theme);
			}

			 // otherwise it is a link to theme color
			return color.createDuplicate();
		}

		if (result !== null) {
			// result have appropriate type for cell already
			return calculateOnTheme(result, theme);
		}
		if (calculatedColor !== null) {
			let fromColorResult = null;
			if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
				fromColorResult = AscFormat.CreateUniFillByUniColor(calculatedColor);
			} else if (cellName === "Color") {
				fromColorResult = calculatedColor;
			}

			return calculateOnTheme(result, theme);
		} else {
			if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
				console.log("no color found. so painting lt1.");
				calculatedColor = AscFormat.CreateUniFillByUniColor(AscFormat.builder_CreateSchemeColor("lt1"));
			} else if (cellName === "Color") {
				// for text color
				console.log("no text color found. so painting dk1.");
				calculatedColor = AscFormat.builder_CreateSchemeColor("dk1");
			}

			return calculateOnTheme(result, theme);
		}
	}

	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonDraw']  = window['AscCommonDraw'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

	window['AscCommonDraw'].themeval = themeval;

})(window, window.document);
