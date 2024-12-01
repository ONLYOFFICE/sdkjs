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
	 * @param {string?} themeValue - value to calculate if cell is not considered
	 * @param {string?} defaultValue
	 * @param {boolean?} gradientEnabled
	 * @param {number?}  themedColorsRow
	 * @return {CUniFill | CUniColor | number | any}
	 */
	function themeval(cell, shape, pageInfo, themes, themeValue,
					  defaultValue, gradientEnabled, themedColorsRow) {
		// https://visualsignals.typepad.co.uk/vislog/2013/05/visio-2013-themes-in-the-shapesheet-part-2.html

		// if cell value is not 'Themed' waiting number representing color otherwise
		// if QuickStyle cell is from 100 to 106 or from 200 to 206 using VariationColorIndex cell value and
		// 	QuickStyle cell value % 100 get color from theme
		// if QuickStyle cell value is not in that range it is smt like from 0 to 7 representing theme colors like:
		//  dk1, lt1, accent1, ...

		/** @type {CUniColor} */
		let calculatedColor = null;
		/** @type {CUniFill | CUniColor | number | any} */
		let result = null;

		let cellValue = cell && cell.v; // formula?
		let cellName = cell && cell.n;

		let quickStyleCellName;
		let quickStyleModifiersCellName;
		let getModifiersMethod;
		let isGetLineEndStyle = false;
		let isFillIdx = false;
		let isLineIdx = false;
		let isFontIdx = false;

		if (themeValue === "LineColor") {
			cellName = "LineColor";
			cellValue = "Themed";
		} else if (themeValue === "FillColor") {
			cellName = "FillForegnd";
			cellValue = "Themed";
		} else if (themeValue === "TextColor") {
			cellName = "Color";
			cellValue = "Themed";
		}

		let initialDefaultValue = null;

		if (gradientEnabled === undefined) {
			gradientEnabled = true;
		}

		if (cellName === "LineColor") {
			quickStyleCellName = "QuickStyleLineColor";
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLnStyle;
			isLineIdx = true;

			initialDefaultValue = AscFormat.CreateUnfilFromRGB(0,0,0);
		} else if (cellName === "Color") {
			// Text color
			quickStyleCellName = "QuickStyleFontColor";
			quickStyleModifiersCellName = "QuickStyleFontMatrix";
			getModifiersMethod = themes[0].getFontStyle;
			isFontIdx = true;

			initialDefaultValue =  AscFormat.CreateUnfilFromRGB(0,0,0).fill.color;
		} else if (cellName === "FillForegnd" || cellName === "FillBkgnd" ||
			cellName === "GradientStopColor" || cellName === "GradientStopPosition" ||
			cellName === "FillGradientAngle" ) {
			quickStyleCellName = "QuickStyleFillColor";
			quickStyleModifiersCellName = "QuickStyleFillMatrix";
			getModifiersMethod = themes[0].getFillStyle;
			isFillIdx = true;

			if (cellName === "FillForegnd") {
				initialDefaultValue =  AscFormat.CreateUnfilFromRGB(255,255,255);
			} else if (cellName === "FillBkgnd") {
				initialDefaultValue =  AscFormat.CreateUnfilFromRGB(0,0,0);
			} else if (cellName === "GradientStopColor") {
				initialDefaultValue =  AscFormat.CreateUniColorRGB(255,255,255);
			} else if (cellName === "GradientStopPosition") {
				initialDefaultValue =  0;
			} else if (cellName === "FillGradientAngle") {
				initialDefaultValue =  5400000;
			}
		} else if (cellName === "LinePattern") {
			// dash dot or smth. get from a:ln from a:lnStyleLst
			// use QuickStyleLineColor to calculate all line params
			quickStyleCellName = "QuickStyleLineColor";
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLnStyle;
			isLineIdx = true;

			initialDefaultValue = 1; // visio solid
		} else if (cellName === "LineWeight") {
			// line weight in inches
			quickStyleCellName = "QuickStyleLineColor";
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLnStyle;
			isLineIdx = true;

			// // 9255 emus = 0.01041666666666667 inches is document.xml StyleSheet ID=0 LineWeight e. g. default value
			initialDefaultValue = 0.01041666666666667;
		} else if (cellName === "FillGradientEnabled") {
			quickStyleCellName = "QuickStyleFillColor";
			quickStyleModifiersCellName = "QuickStyleFillMatrix";
			getModifiersMethod = themes[0].getFillStyle;
			isFillIdx = true;

			initialDefaultValue = false;
		} else if (cellName === "GradientStopColorTrans") {
			// quickStyleCellName = "QuickStyleFillColor";
			// quickStyleModifiersCellName = "QuickStyleFillMatrix";
			// getModifiersMethod = themes[0].getFillStyle;
			// isFillIdx = true;
			//
			// initialDefaultValue = 0;
			AscCommon.consoleLog("Themed GradientStopColorTrans is unhandled");
			return 0;
		} else if (cellName === "EndArrow" || cellName === "BeginArrow") {
			quickStyleCellName = null; // so color will not be calculated
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLineEndStyle;
			isGetLineEndStyle = true;
			isLineIdx = true;

			initialDefaultValue = "0"; // string is return type in calculateValue
		} else if (cellName === "EndArrowSize" || cellName === "BeginArrowSize") {
			quickStyleCellName = null; // so color will not be calculated
			quickStyleModifiersCellName = "QuickStyleLineMatrix";
			getModifiersMethod = themes[0].getLineEndStyle;
			isGetLineEndStyle = true;
			isLineIdx = true;

			initialDefaultValue = 2; // number is return type in calculateValue
		} else if (cellName === "FillPattern") {
			quickStyleCellName = null; // so color will not be calculated
			quickStyleModifiersCellName = "QuickStyleFillMatrix";
			getModifiersMethod = themes[0].getFillProp;
			isFillIdx = true;

			initialDefaultValue = 1; // number is return type in calculateValue. Solid fill.
		} else {
			AscCommon.consoleLog("themeval argument error. cell name: " + cellName + " is unknown. return undefined.");
			return undefined;
		}

		// lets define if shape is connector
		// TODO rewrite consider 2.2.7.4.9	Connector
		let isConnectorShape = shape.getCellNumberValue("EndArrow") !== 0
			|| shape.getCellNumberValue("BeginArrow") !== 0;

		// TODO rewrite themeScopeCellName choose consider 2.2.7.4.2	Dynamic Theme Identification
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
		if (isNaN(shapeColorSchemeThemeIndex)) {
			shapeColorSchemeThemeIndex = 0; // zero index means no theme
		}
		if (shapeColorSchemeThemeIndex === 65534) {
			let pageThemeIndex = pageInfo.pageSheet.getCellNumberValue(themeScopeCellName);
			if (!isNaN(pageThemeIndex)) {
				themeIndex = pageThemeIndex;
			} else {
				// it's ok sometimes
				// AscCommon.consoleLog("pageThemeIndexCell not found");
				themeIndex = 0;
			}
		} else {
			themeIndex = shapeColorSchemeThemeIndex;
		}


		// if THEMEVAL was called with themeValue (argument like "FillColor") even if themeIndex is 0 we should
		// use any theme otherwise if no themeValue argument was passed and 0 themeIndex is used we should return
		// default value
		// see colored rectangle in that file https://disk.yandex.ru/d/IzxVtx0a7GqbQA
		let theme = themes[0];
		if (themeIndex === 0) {
			return initialDefaultValue;
		} else {
			// find theme by themeIndex
			theme = themes.find(function (theme) {
				// if search by theme index - theme.themeElements.themeExt.themeSchemeSchemeEnum
				let findThemeByElement = isConnectorShape?
					theme.themeElements.themeExt.themeSchemeSchemeEnum :
					theme.themeElements.clrScheme.clrSchemeExtLst.schemeEnum;
				let themeEnum = Number(findThemeByElement);
				return themeEnum === themeIndex;
			});
			if (theme === undefined) {
				AscCommon.consoleLog("Theme was not found by theme enum in themes. using themes[0]");
				theme = themes[0];
			}
		}


		let quickStyleColor = shape.getCellNumberValue(quickStyleCellName);
		let quickStyleMatrix = shape.getCellNumberValue(quickStyleModifiersCellName);
		// get color using "VariationColorIndex" cell and quickStyleColor cell
		if (!isNaN(quickStyleColor)) {
			if (100 <= quickStyleColor && quickStyleColor <= 106 ||
				(200 <= quickStyleColor && quickStyleColor <= 206)) {
				if (theme.isVariationClrSchemeLstExists()) {
					let variationColorIndex = shape.getCellNumberValue("VariationColorIndex");
					if (!isNaN(variationColorIndex)) {
						if (65534 === variationColorIndex) {
							let pageVariationColorIndexIndex = pageInfo.pageSheet.getCellNumberValue("VariationColorIndex");
							if (!isNaN(pageVariationColorIndexIndex)) {
								variationColorIndex = pageVariationColorIndexIndex;
							} else {
								variationColorIndex = 0;
								// AscCommon.consoleLog("pageVariationColorIndexIndex not found");
							}
						}
						calculatedColor = theme.getVariationClrSchemeColor(variationColorIndex,
							quickStyleColor % 100);
					}
				} else {
					let index = quickStyleColor % 100;
					switch (index) {
						case 0:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent1");
							break;
						case 1:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent2");
							break;
						case 2:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent3");
							break;
						case 3:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent4");
							break;
						case 4:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent5");
							break;
						case 5:
							calculatedColor = AscFormat.builder_CreateSchemeColor("accent6");
							break;
						case 6:
							calculatedColor = AscFormat.builder_CreateSchemeColor("dk1");
							break;
					}
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
				// consider 2.4.4.275	QuickStyleLineMatrix return root stylesheet value (default value)
				return initialDefaultValue;
			} else if (1 <= quickStyleMatrix && quickStyleMatrix <= 6) {
				if (cellName === "EndArrow" || cellName === "EndArrowSize" ||
					cellName === "BeginArrow" || cellName === "BeginArrowSize") {
					// getLineEndStyle is method
					getMedifiersResult = theme.getLineEndStyle(quickStyleMatrix, isConnectorShape);
				} else {
					getMedifiersResult = getModifiersMethod.call(theme, quickStyleMatrix, calculatedColor, isConnectorShape);
				}
			} else if (100 <= quickStyleMatrix && quickStyleMatrix <= 103) {
				let variationStyleIndex = shape.getCellNumberValue("VariationStyleIndex");
				if (!isNaN(variationStyleIndex)) {
					if (65534 === variationStyleIndex) {
						let pageVariationStyleIndexIndex = pageInfo.pageSheet.getCellNumberValue("VariationStyleIndex");
						if (!isNaN(pageVariationStyleIndexIndex)) {
							variationStyleIndex = pageVariationStyleIndexIndex;
						} else {
							variationStyleIndex = 0;
							// AscCommon.consoleLog("pageVariationStyleIndexIndex not found");
						}
					}
					let varStyle = theme.getVariationStyleScheme(variationStyleIndex,
						quickStyleMatrix % 100);
					let styleId = null;
					if (varStyle) {
						if (isFillIdx) {
							styleId = varStyle.fillIdx;
						} else if (isLineIdx) {
							styleId = varStyle.lineIdx;
						} else if (isFontIdx) {
							styleId = varStyle.fontIdx;
						}
					}
					if (null !== styleId) {
						if (isGetLineEndStyle) {
							// getLineEndStyle is method.
							// When we get lineEnd/lineStart type or its size we don't need
							// calculatedColor and so arguments are different
							// getModifiersMethod.call(theme, styleId, isConnectorShape);
							getMedifiersResult = theme.getLineEndStyle(styleId, isConnectorShape);
						} else {
							getMedifiersResult = getModifiersMethod.call(theme, styleId, calculatedColor, isConnectorShape);
						}
					}
				}
			}

			if (!gradientEnabled && getMedifiersResult && getMedifiersResult.fill instanceof AscFormat.CGradFill) {
				// disable gradient
				getMedifiersResult = AscFormat.CreateUniFillByUniColor(calculatedColor);
			}

			// getModifiersMethod return not only
			// uniFill, so we narrow down the range of returns
			if (cellName === "LineColor") {
				result = getMedifiersResult && getMedifiersResult.Fill;
			} else if (cellName === "LinePattern") {
				// simple type - number
				let originalDashType = getMedifiersResult && getMedifiersResult.prstDash;
				// map c_oDashType number to c_oDashType visio number
				let toVisioDashMap = {
					0: 2,
					1: 4,
					2: 3,
					3: 16,
					4: 18,
					5: 19,
					6: 1,
					7: 9,
					8: 22,
					9: 12,
					10: 10
				};
				result = toVisioDashMap[originalDashType];
				if (result === undefined) {
					AscCommon.consoleLog("Dash map was not realized. Unknown dash type in theme");
					result = 1; // visio solid
				}
			} else if (cellName === "Color") {
				// and it is color
				result = getMedifiersResult && getMedifiersResult.fontPropsObject.color;
			} else if (cellName === "FillForegnd" || cellName === "FillBkgnd") {
				//leave result because it is fill
				result = getMedifiersResult;
			} else if (cellName === "LineWeight") {
				// let's map standart ooxml result in emus to visio result - inches
				result = getMedifiersResult && getMedifiersResult.w /
					(AscCommonWord.g_dKoef_in_to_mm * AscCommonWord.g_dKoef_mm_to_emu);
			} else if (cellName === "FillGradientEnabled") {
				result = getMedifiersResult && getMedifiersResult.fill instanceof AscFormat.CGradFill;
			} else if (cellName === "GradientStopColor") {
				// and it is color
				let themedFillIsGradient = getMedifiersResult && getMedifiersResult.fill instanceof AscFormat.CGradFill;
				if (themedFillIsGradient && getMedifiersResult.fill.colors[themedColorsRow]) {
					result = getMedifiersResult.fill.colors[themedColorsRow].color;
				} else {
					return initialDefaultValue;
				}
			} else if (cellName === "GradientStopPosition") {
				let themedFillIsGradient = getMedifiersResult && getMedifiersResult.fill instanceof AscFormat.CGradFill;
				if (themedFillIsGradient && getMedifiersResult.fill.colors[themedColorsRow]) {
					result = getMedifiersResult.fill.colors[themedColorsRow].pos;
				} else {
					return initialDefaultValue;
				}
			} else if (cellName === "FillGradientAngle") {
				let themedFillIsGradient = getMedifiersResult && getMedifiersResult.fill instanceof AscFormat.CGradFill;
				if (themedFillIsGradient && getMedifiersResult.fill.lin) {
					result = getMedifiersResult.fill.lin.angle;
				} else {
					return initialDefaultValue;
				}
			} else if (cellName === "EndArrow") {
				let endArrowType = getMedifiersResult && getMedifiersResult.lineEx.end;
				result = String(endArrowType);
			} else if (cellName === "BeginArrow") {
				let beginArrowType = getMedifiersResult && getMedifiersResult.lineEx.start;
				result = String(beginArrowType);
			} else if (cellName === "EndArrowSize") {
				let endArrowSize = getMedifiersResult && getMedifiersResult.lineEx.endSize;
				result = Number(endArrowSize);
			} else if (cellName === "BeginArrowSize") {
				let beginArrowSize = getMedifiersResult && getMedifiersResult.lineEx.startSize;
				result = Number(beginArrowSize);
			} else if (cellName === "FillPattern") {
				let fillPattern = getMedifiersResult && getMedifiersResult.pattern;
				result = Number(fillPattern);
			} else {
				AscCommon.consoleLog("Error in themeval. result is not changed to appropriate type or quickStyleCellName is not set.");
			}
		}

		/**
		 * calculate exact color on theme. for quickStyleVariation to consider real color
		 * @param {CUniColor | CUniFill} color
		 * @param {CTheme} theme
		 */
		function calculateOnTheme(color, theme) {
			let RGBA = {R:0, G:0, B:0, A:255};
			if (color instanceof AscFormat.CUniColor ) {
				color.Calculate(theme, undefined, undefined, undefined, RGBA);
			} else if (color instanceof AscFormat.CUniFill) {
				color.calculate(theme, undefined, undefined, undefined, RGBA);
			}

			// otherwise it is a link to theme color
			return color.createDuplicate();
		}

		// typeof NaN === "number", isNaN({...}) === true so NaN check is typeof result === "number" && isNaN(result)
		if (result !== null && !(typeof result === "number" && isNaN(result)) && result !== "null" &&
			result !== undefined && result !== "undefined") {
			// result have appropriate type for cell already
			if (result instanceof AscFormat.CUniColor || result instanceof AscFormat.CUniFill) {
				return calculateOnTheme(result, theme);
			} else {
				// simple type like string or number - so we don't use clone
				return result;
			}
		} else {
			AscCommon.consoleLog("Unknown themeval error. Return initialDefaultValue");
			return initialDefaultValue;
		}
		// code below never calls. result is always calculated. Default values are set above are returned if
		// default theme is used
		// if (calculatedColor !== null) {
		// 	let fromColorResult = null;
		// 	if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
		// 		fromColorResult = AscFormat.CreateUniFillByUniColor(calculatedColor);
		// 	} else if (cellName === "Color") {
		// 		fromColorResult = calculatedColor;
		// 	}
		//
		// 	return calculateOnTheme(result, theme);
		// } else {
		// 	if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
		// 		AscCommon.consoleLog("no color found. so painting lt1.");
		// 		calculatedColor = AscFormat.CreateUniFillByUniColor(AscFormat.builder_CreateSchemeColor("lt1"));
		// 	} else if (cellName === "Color") {
		// 		// for text color
		// 		AscCommon.consoleLog("no text color found. so painting dk1.");
		// 		calculatedColor = AscFormat.builder_CreateSchemeColor("dk1");
		// 	}
		//
		// 	return calculateOnTheme(result, theme);
		// }
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
