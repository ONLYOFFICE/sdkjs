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
	// Import
	let Shape_Type = window['AscCommonDraw'].Shape_Type;

	/**
	 * calculateShapeParamsAndConvertToCShape
	 * @memberof Shape_Type
	 * @param {CVisioDocument} visioDocument
	 * @param {Page_Type} pageInfo
	 * @return {{geometryCShape: CShape, textCShape: ?CShape}} cShapesObjects
	 */
	Shape_Type.prototype.toGeometryAndTextCShapes = function (visioDocument, pageInfo) {
		/**
		 * Can parse themeval
		 * @param {Cell_Type} cell
		 * @param {Shape_Type} shape
		 * @param {Page_Type} pageInfo
		 * @param {CTheme[]} themes
		 * @param {{fontColor?: boolean, lineUniFill?: boolean, uniFillForegnd?: boolean}} themeValWasUsedFor - changes during function
		 * @return {(CUniFill | CUniColor | *)}
		 */
		function calculateCellValue(cell, shape, pageInfo, themes, themeValWasUsedFor) {
			let cellValue = cell && cell.v;
			let cellName = cell && cell.n;

			let returnValue;

			if (/#\w{6}/.test(cellValue)) {
				// check if hex
				let rgba = AscCommon.RgbaHexToRGBA(cellValue);
				if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
					returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B);
				} else if (cellName === "Color") {
					// for text color
					returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B).fill.color;
				} else {
					console.log("wrong calculateCellValue argument cell. Cell unsupported. return null");
					return null;
				}
			} else if (cellValue === 'Themed') {
				// equal to THEMEVAL() call
				returnValue = AscCommonDraw.themeval(cell, shape, pageInfo, themes);

				if (cellName === "LineColor") {
					themeValWasUsedFor.lineUniFill = true;
				} else if (cellName === "FillForegnd") {
					themeValWasUsedFor.uniFillForegnd = true;
				} else if (cellName === "Color") {
					// for text color
					themeValWasUsedFor.fontColor = true;
				}
			} else {
				let colorIndex = parseInt(cellValue);
				if (!isNaN(colorIndex)) {
					let rgba = AscCommon.RgbaHexToRGBA(cellValue);
					switch (colorIndex) {
						case 0:
							rgba = AscCommon.RgbaHexToRGBA('#000000');
							break;
						case 1:
							rgba = AscCommon.RgbaHexToRGBA('#FFFFFF');
							break;
						case 2:
							rgba = AscCommon.RgbaHexToRGBA('#FF0000');
							break;
						case 3:
							rgba = AscCommon.RgbaHexToRGBA('#00FF00');
							break;
						case 4:
							rgba = AscCommon.RgbaHexToRGBA('#0000FF');
							break;
						case 5:
							rgba = AscCommon.RgbaHexToRGBA('#FFFF00');
							break;
						case 6:
							rgba = AscCommon.RgbaHexToRGBA('#FF00FF');
							break;
						case 7:
							rgba = AscCommon.RgbaHexToRGBA('#00FFFF');
							break;
						case 8:
							rgba = AscCommon.RgbaHexToRGBA('#800000');
							break;
						case 9:
							rgba = AscCommon.RgbaHexToRGBA('#008000');
							break;
						case 10:
							rgba = AscCommon.RgbaHexToRGBA('#000080');
							break;
						case 11:
							rgba = AscCommon.RgbaHexToRGBA('#808000');
							break;
						case 12:
							rgba = AscCommon.RgbaHexToRGBA('#800080');
							break;
						case 13:
							rgba = AscCommon.RgbaHexToRGBA('#008080');
							break;
						case 14:
							rgba = AscCommon.RgbaHexToRGBA('#C0C0C0');
							break;
						case 15:
							rgba = AscCommon.RgbaHexToRGBA('#E6E6E6');
							break;
						case 16:
							rgba = AscCommon.RgbaHexToRGBA('#CDCDCD');
							break;
						case 17:
							rgba = AscCommon.RgbaHexToRGBA('#B3B3B3');
							break;
						case 18:
							rgba = AscCommon.RgbaHexToRGBA('#9A9A9A');
							break;
						case 19:
							rgba = AscCommon.RgbaHexToRGBA('#808080');
							break;
						case 20:
							rgba = AscCommon.RgbaHexToRGBA('#666666');
							break;
						case 21:
							rgba = AscCommon.RgbaHexToRGBA('#4D4D4D');
							break;
						case 22:
							rgba = AscCommon.RgbaHexToRGBA('#333333');
							break;
						case 23:
							rgba = AscCommon.RgbaHexToRGBA('#1A1A1A');
							break;
					}
					if (rgba) {
						if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
							returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B);
						} else if (cellName === "Color") {
							returnValue = AscFormat.CreateUnfilFromRGB(rgba.R, rgba.G, rgba.B).fill.color;
						} else {
							console.log("wrong calculateCellValue argument cell. Cell unsupported. return null");
							return null;
						}
					}
				}
			}
			if (!returnValue) {
				if (cellName === "LineColor" || cellName === "FillForegnd" || cellName === "FillBkgnd") {
					console.log("no color found. so painting lt1.");
					returnValue = AscFormat.CreateUniFillByUniColor(AscFormat.builder_CreateSchemeColor("lt1"));
				} else if (cellName === "Color") {
					// for text color
					console.log("no text color found. so painting dk1.");
					returnValue = AscFormat.builder_CreateSchemeColor("dk1");
				} else {
					console.log("no calculateCellValue result return undefined");
				}
			}
			return returnValue;
		}

		/**
		 * handle QuickStyleVariation cell which can change color (but only if color is a result of ThemeVal)
		 * cant be separated for unifill and stroke
		 * @param {CUniFill} oStrokeUniFill
		 * @param {CUniFill} uniFill
		 * @param {Shape_Type} shape
		 * @param {{lineUniFill: boolean, uniFillForegnd: boolean}} themeValWasUsedFor
		 */
		function handleQuickStyleVariation(oStrokeUniFill, uniFill, shape, themeValWasUsedFor) {
			// https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/68bb0221-d8a1-476e-a132-8c60a49cea63?redirectedfrom=MSDN
			// consider "QuickStyleVariation" cell
			// https://visualsignals.typepad.co.uk/vislog/2013/05/visio-2013-themes-in-the-shapesheet-part-2.html
			let backgroundColorHSL = {H: undefined, S: undefined, L: undefined};
			let lineColorHSL = {H: undefined, S: undefined, L: undefined};
			let fillColorHSL = {H: undefined, S: undefined, L: undefined};
			let lineColor = oStrokeUniFill.fill && oStrokeUniFill.fill.color && oStrokeUniFill.fill.color.color.RGBA;
			let fillColor = uniFill.fill && uniFill.fill.color && uniFill.fill.color.color.RGBA;

			if (lineColor !== undefined && fillColor !== undefined) {
				AscFormat.CColorModifiers.prototype.RGB2HSL(255, 255, 255, backgroundColorHSL);
				AscFormat.CColorModifiers.prototype.RGB2HSL(lineColor.R, lineColor.G, lineColor.B, lineColorHSL);
				AscFormat.CColorModifiers.prototype.RGB2HSL(fillColor.R, fillColor.G, fillColor.B, fillColorHSL);

				// covert L to percents
				backgroundColorHSL.L = backgroundColorHSL.L / 255 * 100;
				lineColorHSL.L = lineColorHSL.L / 255 * 100;
				fillColorHSL.L = fillColorHSL.L / 255 * 100;

				let quickStyleVariationCell = shape.getCell("QuickStyleVariation");
				if (quickStyleVariationCell) {
					let quickStyleVariationCellValue = Number(quickStyleVariationCell.v);
					if ((quickStyleVariationCellValue & 4) === 4 && themeValWasUsedFor.lineUniFill) {
						// line color variation enabled (bit mask used)
						if (Math.abs(backgroundColorHSL.L - lineColorHSL.L) < 16.66) {
							if (backgroundColorHSL.L <= 72.92) {
								// if background is dark set stroke to white
								lineColor.R = 255;
								lineColor.G = 255;
								lineColor.B = 255;
							} else {
								if (Math.abs(backgroundColorHSL.L - fillColorHSL.L) >
									Math.abs(backgroundColorHSL.L - lineColorHSL.L)) {
									// evaluation = THEMEVAL("FillColor")
									// get theme shape fill color despite cell
									// line below will give unifill with pattern maybe or gradient
									// lineUniFill = AscCommonDraw.themeval(this.theme, shape, null, "FillColor");
									lineColor.R = fillColor.R;
									lineColor.G = fillColor.G;
									lineColor.B = fillColor.B;
								} else {
									// evaluation = THEMEVAL("LineColor") or not affected I guess
									// get theme line color despite cell
									// lineUniFill = AscCommonDraw.themeval(this.theme, shape, null, "LineColor");
								}
							}
						}
					}

					if ((quickStyleVariationCellValue & 8) === 8 && themeValWasUsedFor.uniFillForegnd) {
						// fill color variation enabled (bit mask used)
						if (Math.abs(backgroundColorHSL.L - fillColorHSL.L) < 16.66) {
							if (backgroundColorHSL.L <= 72.92) {
								// if background is dark set stroke to white
								fillColor.R = 255;
								fillColor.G = 255;
								fillColor.B = 255;
							} else {
								if (Math.abs(backgroundColorHSL.L - lineColorHSL.L) >
									Math.abs(backgroundColorHSL.L - fillColorHSL.L)) {
									// evaluation = THEMEVAL("FillColor")
									// get theme shape fill color despite cell
									// line below will give unifill with pattern maybe or gradient
									// lineUniFill = AscCommonDraw.themeval(this.theme, shape, null, "FillColor");
									fillColor.R = lineColor.R;
									fillColor.G = lineColor.G;
									fillColor.B = lineColor.B;
								}
							}
						}
					}

					// if ((quickStyleVariationCellValue & 2) === 2) {
					// 	// text color variation enabled (bit mask used)
					// 	// Text color variation is realized in getTextCShape function handleTextQuickStyleVariation
					// }
				}
			}
		}

		//TODO import
		/**
		 * afin rotate clockwise
		 * @param {number} x
		 * @param {number} y
		 * @param {number} radiansRotateAngle radians Rotate AntiClockWise Angle. E.g. 30 degrees rotates does DOWN.
		 * @returns {{x: number, y: number}} point
		 */
		function rotatePointAroundCordsStartClockWise(x, y, radiansRotateAngle) {
			let newX = x * Math.cos(radiansRotateAngle) + y * Math.sin(radiansRotateAngle);
			let newY = x * (-1) * Math.sin(radiansRotateAngle) + y * Math.cos(radiansRotateAngle);
			return {x : newX, y: newY};
		}

		/**
		 * @param {CTheme} theme
		 * @param {Shape_Type} shape
		 * @param {CShape} cShape
		 * @param {CUniFill} lineUniFill
		 * @param {CUniFill} fillUniFill
		 * @return {CShape} textCShape
		 */
		function getTextCShape(theme, shape, cShape, lineUniFill, fillUniFill) {
			// see 2.2.8	Text [MS-VSDX]-220215
			/**
			 * handle QuickStyleVariation cell which can change color (but only if color is a result of ThemeVal)
			 * @param textUniColor
			 * @param lineUniFill
			 * @param fillUniFill
			 * @param {{fontColor:boolean}} themeValWasUsedFor - sets during calculateCellValue
			 */
			function handleTextQuickStyleVariation(textUniColor, lineUniFill, fillUniFill, themeValWasUsedFor) {
				// https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/68bb0221-d8a1-476e-a132-8c60a49cea63?redirectedfrom=MSDN
				// consider "QuickStyleVariation" cell
				// https://visualsignals.typepad.co.uk/vislog/2013/05/visio-2013-themes-in-the-shapesheet-part-2.html

				// line and fill QuickStyleVariation are handled in handleQuickStyleVariation

				if (!themeValWasUsedFor.fontColor) {
					return;
				}

				let backgroundColorHSL = {H: undefined, S: undefined, L: undefined};
				let textColorHSL = {H: undefined, S: undefined, L: undefined};
				let lineColorHSL = {H: undefined, S: undefined, L: undefined};
				let fillColorHSL = {H: undefined, S: undefined, L: undefined};

				let textColorRGBA = textUniColor.color && textUniColor.color.RGBA;
				let lineColorRGBA = lineUniFill.fill && lineUniFill.fill.color && lineUniFill.fill.color.color.RGBA;
				let fillColorRGBA = fillUniFill.fill && fillUniFill.fill.color && fillUniFill.fill.color.color.RGBA;

				if (lineColorRGBA !== undefined && fillColorRGBA !== undefined && textColorRGBA !== undefined) {
					AscFormat.CColorModifiers.prototype.RGB2HSL(255, 255, 255, backgroundColorHSL);
					AscFormat.CColorModifiers.prototype.RGB2HSL(lineColorRGBA.R, lineColorRGBA.G, lineColorRGBA.B, lineColorHSL);
					AscFormat.CColorModifiers.prototype.RGB2HSL(fillColorRGBA.R, fillColorRGBA.G, fillColorRGBA.B, fillColorHSL);
					AscFormat.CColorModifiers.prototype.RGB2HSL(textColorRGBA.R, textColorRGBA.G, textColorRGBA.B, textColorHSL);

					// covert L to percents
					backgroundColorHSL.L = backgroundColorHSL.L / 255 * 100;
					lineColorHSL.L = lineColorHSL.L / 255 * 100;
					fillColorHSL.L = fillColorHSL.L / 255 * 100;
					textColorHSL.L = textColorHSL.L / 255 * 100;


					let quickStyleVariationCell = shape.getCell("QuickStyleVariation");
					if (quickStyleVariationCell) {
						let quickStyleVariationCellValue = Number(quickStyleVariationCell.v);

						if ((quickStyleVariationCellValue & 2) === 2) {
							// text color variation enabled (bit mask used)

							// let fillPattern = shape.getCellNumberValue("FillPattern");
							// if (fillPattern !== 0) {
							// 	console.log("TextQuickStyleVariation for shapes with FillPattern !== 0 is disabled");
							// 	// consider example https://disk.yandex.ru/d/2fbgXRrCBThlCw
							// 	return;
							// }

							if (Math.abs(backgroundColorHSL.L - textColorHSL.L) < 16.66) {
								if (backgroundColorHSL.L <= 72.92) {
									// if background is dark set stroke to white
									textColorRGBA.R = 255;
									textColorRGBA.G = 255;
									textColorRGBA.B = 255;
								} else {
									// return the color with the largest absolute difference in luminance from the
									// formula evaluation of the "TextColor", "FillColor", and "LineColor"
									let fillDifferenceIsTheLargest =
										Math.abs(backgroundColorHSL.L - fillColorHSL.L) >
										Math.abs(backgroundColorHSL.L - lineColorHSL.L) &&
										Math.abs(backgroundColorHSL.L - fillColorHSL.L) >
											Math.abs(backgroundColorHSL.L - textColorHSL.L);
									if (fillDifferenceIsTheLargest) {
										textColorRGBA.R = fillColorRGBA.R;
										textColorRGBA.G = fillColorRGBA.G;
										textColorRGBA.B = fillColorRGBA.B;
									} else {
										if (Math.abs(backgroundColorHSL.L - lineColorHSL.L) >
											Math.abs(backgroundColorHSL.L - textColorHSL.L)) {
											textColorRGBA.R = lineColorRGBA.R;
											textColorRGBA.G = lineColorRGBA.G;
											textColorRGBA.B = lineColorRGBA.B;
										} // else leave text color
									}
								}
							}
						}
					}
				}
			}

			/**
			 * @param propsRowNum
			 * @param {?Section_Type} paragraphPropsCommon
			 * @param textCShape
			 */
			function parseParagraphAndAddToShapeContent(propsRowNum, paragraphPropsCommon, textCShape) {
				if (paragraphPropsCommon === null) {
					console.log("paragraphPropsCommon is null. Creating default paragraph");
					// create new paragraph to hold new properties
					let oContent = textCShape.getDocContent();
					let paragraph = new Paragraph(textCShape.getDrawingDocument(), null, true);
					// Set defaultParagraph justify/align text - center
					paragraph.Pr.SetJc(AscCommon.align_Left);
					oContent.Content.push(paragraph);
					paragraph.SetParent(oContent);
					return;
				}
				let paragraphPropsFinal = propsRowNum !== null && paragraphPropsCommon.getRow(propsRowNum);

				// handle horizontal align

				// 0 Specifies that the defaultParagraph is left aligned.
				// 1 Specifies that the defaultParagraph is centered.
				// 2 Specifies that the defaultParagraph is right aligned.
				// 3 Specifies that the defaultParagraph is justified.
				// 4 Specifies that the defaultParagraph is distributed.
				let hAlignCell = paragraphPropsFinal && paragraphPropsFinal.getCell("HorzAlign");

				let horizontalAlign = AscCommon.align_Left;
				if (hAlignCell && hAlignCell.constructor.name === "Cell_Type") {
					// omit calculateCellValue here
					// let fontColor = calculateCellValue(theme, shape, characterColorCell);
					let horAlignTryParse = Number(hAlignCell.v);
					if (!isNaN(horAlignTryParse)) {
						switch (horAlignTryParse) {
							case 0:
								horizontalAlign = AscCommon.align_Left;
								break;
							case 1:
								horizontalAlign = AscCommon.align_Center;
								break;
							case 2:
								horizontalAlign = AscCommon.align_Right;
								break;
							case 3:
								horizontalAlign = AscCommon.align_Justify;
								break;
							case 4:
								horizontalAlign = AscCommon.align_Distributed;
								break;
						}
					} else {
						console.log("horizontal align was not parsed so default is set (left)");
					}
				} else {
					console.log("horizontal align cell was not found so default is set (left)");
				}


				// create new paragraph to hold new properties
				let oContent = textCShape.getDocContent();
				let paragraph = new Paragraph(textCShape.getDrawingDocument(), null, true);
				// Set defaultParagraph justify/align text - center
				paragraph.Pr.SetJc(horizontalAlign);
				oContent.Content.push(paragraph);
				paragraph.SetParent(oContent);

				// paragraph.Pr.Spacing.Before = 0;
				// paragraph.Pr.Spacing.After = 0;
			}


			/**
			 * Parses run props and adds run to paragraph
			 * @param characterRowNum
			 * @param characterPropsCommon
			 * @param oRun
			 * @param paragraph
			 * @param lineUniFill
			 * @param fillUniFill
			 * @param theme
			 * @param shape
			 * @param visioDocument
			 */
			function parseRunAndAddToParagraph(characterRowNum, characterPropsCommon,  oRun, paragraph, lineUniFill, fillUniFill, theme, shape, visioDocument) {
				let characterPropsFinal = characterRowNum !== null && characterPropsCommon.getRow(characterRowNum);

				/**
				 * Let's memorize what color properties used themeVal because quickStyleVariation can change only those
				 * color props that used themeVal function.
				 * @type {{fontColor:boolean}}
				 */
				let themeValWasUsedFor = {
					fontColor: false
				}

				// handle Color
				let characterColorCell = characterPropsFinal && characterPropsFinal.getCell("Color");
				let fontColor;
				if (characterColorCell && characterColorCell.constructor.name === "Cell_Type") {
					fontColor = calculateCellValue(characterColorCell, shape, pageInfo,
						visioDocument.themes, themeValWasUsedFor);
				} else {
					console.log("text color cell not found! set text color as themed");
					fontColor = AscCommonDraw.themeval(null, shape, pageInfo, visioDocument.themes, "TextColor");
					themeValWasUsedFor.fontColor = true;
				}

				handleTextQuickStyleVariation(fontColor, lineUniFill, fillUniFill, themeValWasUsedFor);
				const textColor1 = new CDocumentColor(fontColor.color.RGBA.R, fontColor.color.RGBA.G,
					fontColor.color.RGBA.B, false);
				oRun.Set_Color(textColor1);

				// handle fontSize (doesn't work - see comment below)
				let fontSizeCell = characterPropsFinal && characterPropsFinal.getCell("Size");
				if (fontSizeCell && fontSizeCell.constructor.name === "Cell_Type") {
					// omit calculateCellValue here
					// let fontColor = calculateCellValue(theme, shape, characterColorCell);
					let fontSize = Number(fontSizeCell.v);
					if (!isNaN(fontSize)) {
						nFontSize = fontSize * 72; // convert from in to pt
					} else {
						console.log("font size was not parsed so default is set (9 pt)");
					}
				} else {
					console.log("font size was not found so default is set (9 pt)");
				}
				// i dont know why but when i set font size not for overall shape but for runs text shifts to the top
				// oRun.SetFontSize(nFontSize);

				// handle font
				let cRFonts = new CRFonts();
				cRFonts.Ascii = {Name: "Calibri", Index: 1};
				cRFonts.HAnsi = {Name: "Calibri", Index: 1};
				cRFonts.CS = {Name: "Calibri", Index: 1};
				cRFonts.EastAsia = {Name: "Calibri", Index: 1};
				let fontCell = characterPropsFinal && characterPropsFinal.getCell("Font");
				if (fontCell && fontCell.constructor.name === "Cell_Type") {
					// TODO support Themed values
					// let fontColor = calculateCellValue(theme, shape, characterColorCell);

					// all document fonts all loaded already in CVisioDocument.prototype.loadFonts

					let fontName = fontCell.v;
					if (fontName !== "Themed") {
						let loadedCFont = visioDocument.loadedFonts.find(function (cFont) {
							return cFont.name === fontName;
						});
						if (loadedCFont !== undefined) {
							cRFonts.Ascii = {Name: fontName, Index: -1};
							cRFonts.HAnsi = {Name: fontName, Index: -1};
							cRFonts.CS = {Name: fontName, Index: -1};
							cRFonts.EastAsia = {Name: fontName, Index: -1};
						} else {
							console.log("Font was not found for Run. So default is set (Calibri).");
						}
					} else {
						console.log("Font themed is unhandeled, Calibri is used.");
					}
				} else {
					console.log("fontCell was not found so default is set (Calibri). Check mb AsianFont or ScriptFont");
				}
				oRun.Set_RFonts(cRFonts);

				// add run to last paragraph
				paragraph.Add_ToContent(paragraph.Content.length - 1, oRun);
			}



			let textElement = shape.getTextElement();
			if (!textElement) {
				return null;
			}

			/**
			 * text shape saves text only no fill and no line. it goes along with CShape with the same id + "Text".
			 * It has not local coordinates but the same cord system like shape.
			 * @type {CShape}
			 */
			let textCShape = new AscFormat.CShape();
			textCShape.Id = cShape.Id + "Text";
			textCShape.setParent(visioDocument);

			// set default settings
			// see sdkjs/common/Drawings/CommonController.js createTextArt: function (nStyle, bWord, wsModel, sStartString)
			// for examples
			// https://api.onlyoffice.com/docbuilder/textdocumentapi just some related info
			let bWord = false;
			textCShape.setWordShape(bWord);
			textCShape.setBDeleted(false);
			if (bWord) {
				textCShape.createTextBoxContent();
			} else {
				textCShape.createTextBody();
			}
			textCShape.setVerticalAlign(1); // sets text vert align center equal to anchor set to txBody bodyPr
			textCShape.bSelectedText = false;

			// instead of AscFormat.AddToContentFromString(oContent, sText);
			// use https://api.onlyoffice.com/docbuilder/presentationapi/apishape api implementation code
			// to work with text separated into ParaRuns to split properties use

			// read propsCommonObjects
			let characterPropsCommon = shape.getSection("Character");
			let paragraphPropsCommon = shape.getSection("Paragraph");
			let fieldPropsCommon = shape.getSection("Field");

			// to store last entries of cp/pp/tp like
			//					<cp IX='0'/> or
			//         <tp IX='0'/>
			// character properties are used until another element specifies new character properties.
			// TODO tp_Type is not parsed?
			let propsRunsObjects = {
				"cp_Type": null,
				"tp_Type": null
			};

			let oContent = textCShape.getDocContent();
			oContent.Content = [];

			let nFontSize = 9;

			// read text
			textElement.elements.forEach(function(textElementPart, i) {
				if (typeof textElementPart === "string" || textElementPart.constructor.name === "fld_Type") {

					// create defaultParagraph
					if (oContent.Content.length === 0) {
						parseParagraphAndAddToShapeContent(0, paragraphPropsCommon, textCShape);
					}
					let paragraph = oContent.Content.slice(-1)[0];

					// create paraRun using propsObjects

					// equal to ApiParagraph.prototype.AddText method
					let oRun = new ParaRun(paragraph, false);
					if (typeof textElementPart === "string") {
						oRun.AddText(textElementPart);
					} else if (textElementPart.constructor.name === "fld_Type") {
						// text field

						let optionalValue = textElementPart.value;

						let fieldRowNum = textElementPart.iX;
						let fieldPropsFinal = fieldRowNum !== null && fieldPropsCommon.getRow(fieldRowNum);

						// handle Value
						let fieldValueCell = fieldPropsFinal && fieldPropsFinal.getCell("Value");

						if (fieldValueCell.v || optionalValue) {
							oRun.AddText(fieldValueCell.v || optionalValue);
						} else {
							console.log("field_Type was not parsed");
						}
					}

					// setup Run
					// check character properties: get cp_Type object and in characterPropsCommon get needed Row
					let characterRowNum = propsRunsObjects.cp_Type && propsRunsObjects.cp_Type.iX;
					if (propsRunsObjects.cp_Type === null) {
						characterRowNum = 0;
					}
					parseRunAndAddToParagraph(characterRowNum, characterPropsCommon,
						oRun, paragraph, lineUniFill, fillUniFill, theme, shape,
						visioDocument);
				} else if (textElementPart.constructor.name === "pp_Type") {
					// setup Paragraph

					// check defaultParagraph properties: get pp_Type object and in paragraphPropsCommon get needed Row
					let paragraphRowNum = textElementPart.iX;
					parseParagraphAndAddToShapeContent(paragraphRowNum, paragraphPropsCommon, textCShape);

				} else if (textElementPart.constructor.name === "cp_Type" || textElementPart.constructor.name === "tp_Type") {
					propsRunsObjects[textElementPart.constructor.name] = textElementPart;
				} else {
					console.log("undkown type in text tag");
				}
			});

			// create defaultParagraph if no strings found
			if (oContent.Content.length === 0) {
				// create defaultParagraph
				parseParagraphAndAddToShapeContent(0, paragraphPropsCommon, textCShape);
			}

			// handle horizontal align i. e. defaultParagraph align

			// handle vertical align
			let verticalAlignCell = shape.getCell("VerticalAlign");
			if (verticalAlignCell) {
				// 0 - top, 1 - middle, 2 - bottom
				let verticalAlign = Number(verticalAlignCell.v);
				if (!isNaN(verticalAlign)) {
					//  0 - bottom, 1, 2, 3 - ctr, 4, - top
					// but global_MatrixTransformer transformations changes values to
					// 0 - top, 1, 2, 3 - center, 4 - bottom
					if (verticalAlign === 0) {
						textCShape.setVerticalAlign(0); // sets text vert align center equal to anchor set to txBody bodyPr
					} else if (verticalAlign === 2) {
						textCShape.setVerticalAlign(4); // sets text vert align center equal to anchor set to txBody bodyPr
					}
					// else leave center align
				} else {
					console.log("vertical align cell was not parsed for shape. align set to center. Shape:", shape);
				}
			} else {
				console.log("vertical align cell was not found for shape. align set to center. Shape:", shape);
			}


			// setup text properties
			let oTextPr;
			oTextPr = new CTextPr();
			// i dont know why but when i set font size not for overall shape but for runs text shifts to the top
			oTextPr.FontSize = nFontSize;
			// oTextPr.RFonts.Ascii = {Name: "Calibri", Index: -1};
			// oTextPr.RFonts.HAnsi = {Name: "Calibri", Index: -1};
			// oTextPr.RFonts.CS = {Name: "Calibri", Index: -1};
			// oTextPr.RFonts.EastAsia = {Name: "Calibri", Index: -1};

			// apply text properties
			oContent.SetApplyToAll(true);
			oContent.AddToParagraph(new ParaTextPr(oTextPr));
			// oContent.SetParagraphAlign(AscCommon.align_Center);
			oContent.SetApplyToAll(false);

			let oBodyPr = textCShape.getBodyPr().createDuplicate();
			oBodyPr.rot = 0;
			// oBodyPr.spcFirstLastPara = false;
			// oBodyPr.vertOverflow = AscFormat.nVOTOverflow;
			// oBodyPr.horzOverflow = AscFormat.nHOTOverflow;
			// oBodyPr.vert = AscFormat.nVertTThorz; // default //( ( Horizontal ))
			oBodyPr.wrap = AscFormat.nTWTSquare; // default
			// oBodyPr.setDefaultInsets();
			// oBodyPr.numCol = 1;
			// oBodyPr.spcCol = 0;
			// oBodyPr.rtlCol = 0;
			// oBodyPr.fromWordArt = false;
			// oBodyPr.anchor = 1; // 4 - bottom, 1,2,3 - center
			// oBodyPr.anchorCtr = false;
			// oBodyPr.forceAA = false;
			// oBodyPr.compatLnSpc = true;
			// // oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry("textNoShape");

			// cShape.bCheckAutoFitFlag = true;
			// oBodyPr.textFit = new AscFormat.CTextFit();
			// oBodyPr.textFit.type = AscFormat.text_fit_Auto;

			// oBodyPr.upright = false; // default

			let leftMarginInch = shape.getCellNumberValue("LeftMargin");
			let topMarginInch = shape.getCellNumberValue("TopMargin");
			let rightMarginInch = shape.getCellNumberValue("RightMargin");
			let bottomMarginInch = shape.getCellNumberValue("BottomMargin");


			// CHECKS SIGN but positive tIns gives bottom inset. Check https://disk.yandex.ru/d/IU1vdjzcF9p3IQ
			// it is may global graphics transform issue so set bottom inset as top and opposite
			oBodyPr.tIns = bottomMarginInch * g_dKoef_in_to_mm;
			oBodyPr.bIns = topMarginInch * g_dKoef_in_to_mm;
			oBodyPr.lIns = leftMarginInch * g_dKoef_in_to_mm;
			oBodyPr.rIns = rightMarginInch * g_dKoef_in_to_mm;

			if (bWord) {
				textCShape.setBodyPr(oBodyPr);
			} else {
				textCShape.txBody.setBodyPr(oBodyPr);
			}


			// handle cords

			// to rotate around point we 1) add one more offset 2) rotate around center
			// could be refactored maybe
			// https://www.figma.com/file/jr1stjGUa3gKUBWxNAR80T/locPinHandle?type=design&node-id=0%3A1&mode=design&t=raXzFFsssqSexysi-1
			let txtPinX_inch = shape.getCellNumberValue("TxtPinX");
			let txtPinY_inch = shape.getCellNumberValue("TxtPinY");


			// consider https://disk.yandex.ru/d/2XzRaPTKzKHFjA
			// where TxtHeight and TxtWidth get all shape height and width and txtPinX_inch and txtPinY_inch are not defined
			// also check for {}, undefined, NaN, null
			let oSpPr = new AscFormat.CSpPr();
			let oXfrm = new AscFormat.CXfrm();
			if (!(isNaN(txtPinX_inch) || txtPinX_inch === null)  && !(isNaN(txtPinY_inch) || txtPinY_inch === null)) {
				// https://www.figma.com/file/WiAC4sxQuJaq65h6xppMYC/cloudFare?type=design&node-id=0%3A1&mode=design&t=SZbio0yIyxq0YnMa-1s

				let shapeWidth = shape.getCellNumberValue("Width");
				let shapeHeight = shape.getCellNumberValue("Height");
				let shapeLocPinX = shape.getCellNumberValue("LocPinX");
				let shapeLocPinY = shape.getCellNumberValue("LocPinY");
				let txtWidth_inch = shape.getCellNumberValue("TxtWidth");
				let txtHeight_inch = shape.getCellNumberValue("TxtHeight");
				let txtLocPinX_inch = shape.getCellNumberValue("TxtLocPinX");
				let txtLocPinY_inch = shape.getCellNumberValue("TxtLocPinY");

				let textAngle = shape.getCellNumberValue("TxtAngle");

				// defaultParagraph.Pr.SetJc(AscCommon.align_Left);
				let oBodyPr = textCShape.getBodyPr().createDuplicate();
				// oBodyPr.anchor = 4; // 4 - bottom, 1,2,3 - center

				let globalXmm = cShape.spPr.xfrm.offX;
				let localXmm = (txtPinX_inch - txtLocPinX_inch) * g_dKoef_in_to_mm;
				oXfrm.setOffX(globalXmm + localXmm); // mm

				let flipYCell = shape.getCell("FlipY");
				let flipVertically = flipYCell ?  flipYCell.v === "1" : false;
				if (flipVertically) {
					// if we flip figure we flip text pinY around shape pinY

					if (txtPinY_inch > 0) {
						// y cord of text block start. when cord system starts in left bottom corner on shape
						let blockCord = txtPinY_inch - txtLocPinY_inch;
						// (y part of vector) from shape center to txt block start
						let fromShapeCenterToBlockStart = blockCord - shapeLocPinY;

						let globalYmm = cShape.spPr.xfrm.offY;

						// mirror distance fromBlock start ToShapeCenter then add text block height to it
						// + shapeLocPinY made shift from shape center to shape bottom bcs we calculate
						// localYmm starting from bottom of shape not from center
						let localYmm = (-fromShapeCenterToBlockStart - txtHeight_inch + shapeLocPinY) * g_dKoef_in_to_mm;
						oXfrm.setOffY(globalYmm + localYmm);
					} else {
						// negative, y part of vector. y cord of text block start. when cord system starts in left bottom corner on shape
						let blockCord = txtPinY_inch + (txtHeight_inch - txtLocPinY_inch);

						// lets make it negative like y part of vector. It comes from top to bottom.
						// It is vector that comes from shape center to text block start.
						let fromBlockToShapeCenter = blockCord - shapeLocPinY;

						let globalYmm = cShape.spPr.xfrm.offY;
						// Finally we mirror fromBlockToShapeCenter by multiplying by -1 and add shapeLocPinY to move its
						// start to bottom on shape
						let localYmm = (-fromBlockToShapeCenter + shapeLocPinY) * g_dKoef_in_to_mm;
						oXfrm.setOffY(globalYmm + localYmm);
					}
				} else {
					let globalYmm = cShape.spPr.xfrm.offY;
					let localYmm = (txtPinY_inch - txtLocPinY_inch) * g_dKoef_in_to_mm;
					oXfrm.setOffY(globalYmm + localYmm);
				}

				oXfrm.setExtX(txtWidth_inch * g_dKoef_in_to_mm);
				oXfrm.setExtY(txtHeight_inch * g_dKoef_in_to_mm);
				oXfrm.setRot( 0);
			} else {
				// create text block with shape sizes
				let globalXmm = cShape.spPr.xfrm.offX;
				let globalYmm = cShape.spPr.xfrm.offY;
				oXfrm.setOffX(globalXmm); // mm
				oXfrm.setOffY(globalYmm);
				oXfrm.setExtX(shapeWidth_inch * g_dKoef_in_to_mm);
				oXfrm.setExtY(shapeHeight_inch * g_dKoef_in_to_mm);
				oXfrm.setRot(0);
			}
			oSpPr.setXfrm(oXfrm);
			oXfrm.setParent(oSpPr);
			oSpPr.setFill(AscFormat.CreateNoFillUniFill());
			oSpPr.setLn(AscFormat.CreateNoFillLine());

			textCShape.setSpPr(oSpPr);
			oSpPr.setParent(textCShape);

			// just trash below
			//
			// // placeholder
			// let oUniNvPr = new AscFormat.UniNvPr();
			// oUniNvPr.nvPr.ph = Asc.asc_docs_api.prototype.CreatePlaceholder("object");
			//
			// // cShape.txBody.content2 = cShape.txBody.content;
			//
			// cShape.setNvSpPr(oUniNvPr);

			// copy settings from presentations debug
			// cShape.txBody.content.CurPos.TableMove = 1;
			// cShape.txBody.content.ReindexStartPos = 0;
			// cShape.txBody.content.Content[0].Content[1].CollPrChangeMine = false;
			// cShape.txBody.content.Content[0].Content[1].State.ContentPos = 10;
			// cShape.txBody.content.Content[0].Index = -1;
			// cShape.txBody.compiledBodyPr = null;

			// Set Paragraph (the only one defaultParagraph exist) justify/align text - center
			// cShape.txBody.content.Content[0].Pr.SetJc(AscCommon.align_Center);

			// cShape.recalculateTextStyles();
			// cShape.recalculateTransformText(); // recalculates text position (i. e. transformText objects);
			// cShape.recalculateContent();
			// cShape.recalculateContent2();
			// cShape.recalculateContentWitCompiledPr();

			// use ParaRun.prototype.Set_Color
			// cShape.txBody.content.Content[0].Content[1].Pr.Color = TextColor1;
			// cShape.txBody.content.Content[0].Content[0].Pr.Color = TextColor1;

			return textCShape;
		}

		// there was case with shape type group with no PinX and PinY
		// https://disk.yandex.ru/d/tl877cuzcRcZYg
		let pinX_inch = this.getCellNumberValue("PinX");
		let pinY_inch = this.getCellNumberValue("PinY");

		// also check for {}, undefined, NaN, null
		if (isNaN(pinX_inch) || pinX_inch === null || isNaN(pinY_inch) || pinY_inch === null) {
			console.log('pinX_inch or pinY_inch is NaN for Shape. Its ok sometimes. ' +
				'Empty CShape is returned. See original shape: ', this);
			// let's use empty shape
			let emptyCShape = new AscFormat.CShape();
			emptyCShape.setWordShape(false);
			emptyCShape.setBDeleted(false);

			var oSpPr = new AscFormat.CSpPr();
			var oXfrm = new AscFormat.CXfrm();
			// oXfrm.setOffX(0);
			// oXfrm.setOffY(0);
			// oXfrm.setExtX(0);
			// oXfrm.setExtY(0);

			oSpPr.setXfrm(oXfrm);
			oXfrm.setParent(oSpPr);
			// oSpPr.setFill(AscFormat.CreateNoFillUniFill());
			// oSpPr.setLn(AscFormat.CreateNoFillLine());

			emptyCShape.setSpPr(oSpPr);
			oSpPr.setParent(emptyCShape)
			return {geometryCShape: emptyCShape, textCShape: null};
		}

		let shapeAngle = this.getCellNumberValue("Angle");
		let locPinX_inch = this.getCellNumberValue("LocPinX");
		let locPinY_inch = this.getCellNumberValue("LocPinY");
		let shapeWidth_inch = this.getCellNumberValue("Width");
		let shapeHeight_inch = this.getCellNumberValue("Height");

		// to rotate around point we 1) add one more offset 2) rotate around center
		// could be refactored maybe
		// https://www.figma.com/file/jr1stjGUa3gKUBWxNAR80T/locPinHandle?type=design&node-id=0%3A1&mode=design&t=raXzFFsssqSexysi-1
		let redVector = {x: -(locPinX_inch - shapeWidth_inch/2), y: -(locPinY_inch - shapeHeight_inch/2)};
		// rotate antiClockWise by shapeAngle
		let purpleVector = rotatePointAroundCordsStartClockWise(redVector.x, redVector.y, -shapeAngle);
		let rotatedCenter = {x: pinX_inch - redVector.x + purpleVector.x, y: pinY_inch - redVector.y + purpleVector.y};
		let turquoiseVector = {x: -shapeWidth_inch/2, y: -shapeHeight_inch/2};
		let x_inch = rotatedCenter.x + turquoiseVector.x + redVector.x;
		let y_inch = rotatedCenter.y + turquoiseVector.y + redVector.y;

		let x_mm = x_inch * g_dKoef_in_to_mm;
		let y_mm = y_inch * g_dKoef_in_to_mm;
		let shapeWidth_mm = shapeWidth_inch * g_dKoef_in_to_mm;
		let shapeHeight_mm = shapeHeight_inch * g_dKoef_in_to_mm;

		// TODO check if gradient enabled
		// let gradientEnabled = shape.getCell("FillGradientEnabled");
		// console.log("Gradient enabled:", gradientEnabled);

		/** @type CUniFill */
		let uniFillForegndWithPattern = null;
		/**
		 * Fill without pattern applied.We need fill without pattern applied bcs pattern applied can set
		 * NoSolidFill object without color, so we will not be able to calculate handleVariationColor function result.
		 * @type CUniFill */
		let uniFillForegnd = null;

		/** @type CUniFill */
		let	uniFillBkgnd = null;

		let	nPatternType = null;

		/** @type CUniFill */
		let lineUniFillWithPattern = null;
		/**
		 * We need fill without pattern applied bcs pattern applied can set NoSolidFill object without color,
		 * so we will not be able to calculate handleVariationColor function result
		 * @type CUniFill */
		let lineUniFill = null;

		/**
		 * Let's memorize what color properties used themeVal because quickStyleVariation can change only those
		 * color props that used themeVal function.
		 * @type {{lineUniFill: boolean, uniFillForegnd: boolean}}
		 */
		let themeValWasUsedFor = {
			lineUniFill : false,
			uniFillForegnd: false
		}


		let fillForegndCell = this.getCell("FillForegnd");
		if (fillForegndCell) {
			// console.log("FillForegnd was found:", fillForegndCell);
			uniFillForegnd = calculateCellValue(fillForegndCell, this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);

			let fillForegndTrans = this.getCell("FillForegndTrans");
			if (fillForegndTrans) {
				let fillForegndTransValue = Number(fillForegndTrans.v);
				if (!isNaN(fillForegndTransValue)) {
					let fillObj = uniFillForegnd.fill;
					if (fillObj.constructor.name === "CPattFill") {
						// pattern fill
						fillObj.fgClr.color.RGBA.A = fillObj.fgClr.color.RGBA.A * (1 - fillForegndTransValue);
					} else {
						fillObj.color.color.RGBA.A = fillObj.color.color.RGBA.A * (1 - fillForegndTransValue);
					}
				} else {
					// console.log("fillForegndTrans value is themed or something. Not calculated for", shape);
				}
			}
		}
		let fillBkgndCell = this.getCell("FillBkgnd");
		if (fillBkgndCell) {
			// console.log("FillBkgnd was found:", fillBkgndCell);
			uniFillBkgnd = calculateCellValue(fillBkgndCell, this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);

			let fillBkgndTrans = this.getCell("FillBkgndTrans");
			if (fillBkgndTrans) {
				let fillBkgndTransValue = Number(fillBkgndTrans.v);
				if (!isNaN(fillBkgndTransValue)) {
					let fillObj = uniFillBkgnd.fill;
					if (fillObj.constructor.name === "CPattFill") {
						// pattern fill
						fillObj.bgClr.color.RGBA.A = fillObj.fgClr.color.RGBA.A * (1 - fillBkgndTransValue);
					} else {
						fillObj.color.color.RGBA.A = fillObj.color.color.RGBA.A * (1 - fillBkgndTransValue);
					}
				} else {
					// console.log("fillBkgndTrans value is themed or something. Not calculated for", this);
				}
			}
		}

		let lineColorCell = this.getCell("LineColor");
		if (lineColorCell) {
			// console.log("LineColor was found for shape", lineColorCell);
			lineUniFill = calculateCellValue(lineColorCell, this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);
		} else {
			console.log("LineColor cell for line stroke (border) was not found painting red");
			lineUniFill = AscFormat.CreateUnfilFromRGB(255,0,0);
		}

		// calculate variation before pattern bcs pattern can make NoFillUniFill object without color
		handleQuickStyleVariation(lineUniFill, uniFillForegnd, this, themeValWasUsedFor);

		// add read matrix modifier width?
		// + handle line pattens?
		let linePattern = this.getCell("LinePattern");
		if (linePattern) {
			if (linePattern.v === "0") {
				lineUniFillWithPattern = AscFormat.CreateNoFillUniFill();
			} else {
				//todo types
				lineUniFillWithPattern = lineUniFill;
			}
		}

		let fillPattern = this.getCell("FillPattern");
		/** @type ?number */
		let fillPatternType = null;
		if (fillPattern) {
			// console.log("fillPattern was found:", fillPattern);
			let fillPatternTypeTryParse = parseInt(fillPattern.v);
			if (!isNaN(fillPatternTypeTryParse)) {
				fillPatternType = fillPatternTypeTryParse;
			}
		}

		if (null !== fillPatternType && uniFillBkgnd && uniFillForegnd) {
			// https://learn.microsoft.com/ru-ru/office/client-developer/visio/fillpattern-cell-fill-format-section
			if (fillPatternType === 0) {
				uniFillForegndWithPattern = AscFormat.CreateNoFillUniFill();
			} else if (fillPatternType === 1) {
				uniFillForegndWithPattern = uniFillForegnd;
			} else if (fillPatternType > 1) {
				//todo types
				nPatternType = 0;//"cross";
				uniFillForegndWithPattern = AscFormat.CreatePatternFillUniFill(nPatternType,
					uniFillBkgnd.fill.color, uniFillForegnd.fill.color);
				// uniFill = AscFormat.builder_CreatePatternFill(nPatternType, uniFillBkgnd.fill.color, uniFillForegnd.fill.color);
			}
		} else if (uniFillForegnd) {
			uniFillForegndWithPattern = uniFillForegnd;
		} else {
			console.log("FillForegnd not found for shape", this);
			uniFillForegndWithPattern = AscFormat.CreateNoFillUniFill();
		}

		let lineWidthEmu = null;
		let lineWeightCell = this.getCell("LineWeight");
		if (lineWeightCell && lineWeightCell.v && lineWeightCell.v !== "Themed") {
			// to cell.v visio always saves inches
			let lineWeightInches = Number(lineWeightCell.v);
			if (!isNaN(lineWeightInches)) {
				lineWidthEmu = lineWeightInches * AscCommonWord.g_dKoef_in_to_mm * AscCommonWord.g_dKoef_mm_to_emu;
			} else {
				console.log("caught unknown error. line will be painted 9525 emus");
				// 9255 emus = 0.01041666666666667 inches is document.xml StyleSheet ID=0 LineWeight e. g. default value
				lineWidthEmu = 9525;
			}
		} else {
			console.log("LineWeight cell was not calculated. line will be painted 9525 emus");
			lineWidthEmu = 9525;
		}

		// console.log("Calculated lineUniFill unifill", lineUniFill, "for shape", this);
		// console.log("Calculated fill UniFill", uniFillForegndWithPattern, "for shape", this);

		var oStroke = AscFormat.builder_CreateLine(lineWidthEmu, {UniFill: lineUniFillWithPattern});
		// var oStroke = AscFormat.builder_CreateLine(12700, {UniFill: AscFormat.CreateUnfilFromRGB(255,0,0)});

		if (this.type === "Foreign") {
			console.log("Shape has type Foreign and may not be displayed. " +
				"Check shape.elements --> ForeignData_Type obj. See shape:", this);
		}

		let flipXCell = this.getCell("FlipX");
		let flipHorizontally = flipXCell ? flipXCell.v === "1" : false;

		let flipYCell = this.getCell("FlipY");
		let flipVertically = flipYCell ?  flipYCell.v === "1" : false;

		let cShape = this.convertToCShapeUsingParamsObj({
			x_mm: x_mm, y_mm: y_mm,
			w_mm: shapeWidth_mm, h_mm: shapeHeight_mm,
			rot: shapeAngle,
			oFill: uniFillForegndWithPattern, oStroke: oStroke,
			flipHorizontally: flipHorizontally, flipVertically: flipVertically,
			cVisioDocument: visioDocument
		});

		cShape.Id = String(this.iD); // it was string in cShape

		let textCShape = getTextCShape(visioDocument.themes[0], this, cShape, lineUniFill, uniFillForegnd);

		cShape.recalculate();
		cShape.recalculateLocalTransform(cShape.transform);

		if (textCShape !== null) {
			textCShape.recalculate();
			textCShape.recalculateLocalTransform(textCShape.transform);
		}

		return {geometryCShape: cShape, textCShape: textCShape};
	}

	/**
	 * converts !Shape TypeGroup! To CGroupShape Recursively.
	 * let's say shape can only have subshapes if its Type='Group'.
	 * @memberOf Shape_Type
	 * @param {CVisioDocument} visioDocument
	 * @param {Page_Type} pageInfo
	 * @param {CGroupShape?} currentGroupHandling
	 * @return {{cGroupShape: CGroupShape, textCShape: CShape}}
	 */
	Shape_Type.prototype.toCGroupShapeRecursively = function (visioDocument, pageInfo, currentGroupHandling) {
		// if we need to create CGroupShape create CShape first then copy its properties to CGroupShape object
		// so anyway create CShapes
		let cShapes = this.toGeometryAndTextCShapes(visioDocument, pageInfo);

		if (this.type === "Group") {
			// CGroupShape cant support text. So cShape will represent everything related to Shape Type="Group".
			// Let's push cShape into CGroupShape object.

			let groupShape = new AscFormat.CGroupShape();
			// this.graphicObjectsController = new AscFormat.DrawingObjectsController();
			// let groupShape = AscFormat.builder_CreateGroup();

			groupShape.setLocks(0);

			groupShape.setBDeleted(false);

			// Create CGroupShape with SpPr from cShape but with no fill and line
			let noLineFillSpPr = cShapes.geometryCShape.spPr.createDuplicate();
			noLineFillSpPr.setFill(AscFormat.CreateNoFillUniFill());
			noLineFillSpPr.setLn(AscFormat.CreateNoFillLine());

			groupShape.setSpPr(noLineFillSpPr);
			groupShape.spPr.setParent(groupShape);
			groupShape.rot = cShapes.geometryCShape.rot;
			groupShape.brush = cShapes.geometryCShape.brush;
			groupShape.bounds = cShapes.geometryCShape.bounds;
			groupShape.flipH = cShapes.geometryCShape.flipH;
			groupShape.flipV = cShapes.geometryCShape.flipV;
			groupShape.localTransform = cShapes.geometryCShape.localTransform;
			groupShape.pen = cShapes.geometryCShape.pen;
			groupShape.Id = cShapes.geometryCShape.Id + "Group";

			groupShape.addToSpTree(groupShape.spTree.length, cShapes.geometryCShape);
			groupShape.spTree[groupShape.spTree.length-1].setGroup(groupShape);

			cShapes.geometryCShape.spPr.xfrm.setOffX(0);
			cShapes.geometryCShape.spPr.xfrm.setOffY(0);
			
			// cShape.setLocks(1)?;

			groupShape.setParent2(visioDocument);

			if (!currentGroupHandling) {

				currentGroupHandling = groupShape;
				let subShapes = this.getSubshapes();
				for (let i = 0; i < subShapes.length; i++) {
					const subShape = subShapes[i];
					subShape.toCGroupShapeRecursively(visioDocument, pageInfo, currentGroupHandling);
				}

				// textCShape is returned from this function

			} else {

				// if currentGroupHandling add groupShape (withShape in it) and textCShape to it

				currentGroupHandling.addToSpTree(currentGroupHandling.spTree.length, groupShape);
				currentGroupHandling.spTree[currentGroupHandling.spTree.length-1].setGroup(currentGroupHandling);
				groupShape.recalculateLocalTransform(groupShape.transform);

				if (cShapes.textCShape !== null) {
					currentGroupHandling.addToSpTree(currentGroupHandling.spTree.length, cShapes.textCShape);
					currentGroupHandling.spTree[currentGroupHandling.spTree.length-1].setGroup(currentGroupHandling);
					// cShapes.textCShape.recalculateLocalTransform(cShapes.textCShape.transform); // exists below
				}

				currentGroupHandling = groupShape;
				let subShapes = this.getSubshapes();
				for (let i = 0; i < subShapes.length; i++) {
					const subShape = subShapes[i];
					subShape.toCGroupShapeRecursively(visioDocument, pageInfo, currentGroupHandling);
				}
			}
			// recalculate text other positions to local (group) coordinates
			cShapes.geometryCShape.recalculateLocalTransform(cShapes.geometryCShape.transform);
			cShapes.geometryCShape.recalculateTransformText();
			cShapes.geometryCShape.recalculateContent();
			// cShape.recalculate(); // doesnt work here

			if (cShapes.textCShape !== null) {
				// even if not add textCShape to currentGroupHandling above do recalculate just in case
				cShapes.textCShape.recalculateLocalTransform(cShapes.textCShape.transform);
				cShapes.textCShape.recalculateTransformText();
				cShapes.textCShape.recalculateContent();
			}

		} else {
			// if read cShape not CGroupShape
			if (!currentGroupHandling) {
				throw new Error("Group handler was called on simple shape");
			} else {
				// add shape and text to currentGroupHandling

				currentGroupHandling.addToSpTree(currentGroupHandling.spTree.length, cShapes.geometryCShape);
				currentGroupHandling.spTree[currentGroupHandling.spTree.length-1].setGroup(currentGroupHandling);

				// recalculate text other positions to local (group) coordinates
				cShapes.geometryCShape.recalculateLocalTransform(cShapes.geometryCShape.transform);
				cShapes.geometryCShape.recalculateTransformText();
				cShapes.geometryCShape.recalculateContent();
				// cShape.recalculate(); // doesnt work here

				if (cShapes.textCShape !== null) {
					currentGroupHandling.addToSpTree(currentGroupHandling.spTree.length, cShapes.textCShape);
					currentGroupHandling.spTree[currentGroupHandling.spTree.length-1].setGroup(currentGroupHandling);

					cShapes.textCShape.recalculateLocalTransform(cShapes.textCShape.transform);
					cShapes.textCShape.recalculateTransformText();
					cShapes.textCShape.recalculateContent();
				}
			}
		}

		if (currentGroupHandling) {
			currentGroupHandling.recalculate();
		}

		return {cGroupShape: currentGroupHandling, textCShape: cShapes.textCShape};
	}

	/**
	 * @memberOf Shape_Type
	 * @param {{x_mm, y_mm, w_mm, h_mm, rot, oFill, oStroke, flipHorizontally, flipVertically, cVisioDocument}} paramsObj
	 * @return {CShape} CShape
	 */
	Shape_Type.prototype.convertToCShapeUsingParamsObj = function(paramsObj) {
		let x = paramsObj.x_mm;
		let y = paramsObj.y_mm;
		let w_mm = paramsObj.w_mm;
		let h_mm = paramsObj.h_mm;
		let rot = paramsObj.rot;
		let oFill = paramsObj.oFill;
		let oStroke = paramsObj.oStroke;
		let cVisioDocument = paramsObj.cVisioDocument;
		let flipHorizontally = paramsObj.flipHorizontally;
		let flipVertically = paramsObj.flipVertically;

		let shapeGeom = AscCommonDraw.getGeometryFromShape(this);

		let sType   = "rect";
		let nWidth_mm  = Math.round(w_mm);
		let nHeight_mm = Math.round(h_mm);
		//let oDrawingDocument = new AscCommon.CDrawingDocument();
		let shape = AscFormat.builder_CreateShape(sType, nWidth_mm, nHeight_mm,
			oFill, oStroke, cVisioDocument, cVisioDocument.theme, null, false);
		shape.spPr.xfrm.setOffX(x);
		shape.spPr.xfrm.setOffY(y);
		shape.spPr.xfrm.setRot(rot);
		shape.spPr.xfrm.setFlipH(flipHorizontally);
		shape.spPr.xfrm.setFlipV(flipVertically);

		shape.spPr.setGeometry(shapeGeom);
		// shape.recalculate();
		return shape;
	};


	//-------------------------------------------------------------export-------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonDraw']  = window['AscCommonDraw'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

})(window, window.document);
