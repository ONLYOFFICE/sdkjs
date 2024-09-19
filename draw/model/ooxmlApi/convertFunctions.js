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
	 * @param {Number} drawingPageScale
	 * @return {{geometryCShape: CShape | CImageShape, textCShape: ?CShape}} cShapesObjects
	 */
	Shape_Type.prototype.toGeometryAndTextCShapes = function (visioDocument, pageInfo, drawingPageScale) {

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
		 * @param {number} drawingPageScale
		 * @return {CShape} textCShape
		 */
		function getTextCShape(theme, shape, cShape, lineUniFill,
							   fillUniFill, drawingPageScale) {
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
					let paragraph = new Paragraph(textCShape.getDrawingDocument(), true);
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
				let paragraph = new Paragraph(textCShape.getDrawingDocument(), true);
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
					fontColor = characterColorCell.calculateValue(shape, pageInfo,
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

			// see https://disk.yandex.ru/d/xy2yxhAQHlUHsA shape with number-text has HideText cell v=1 and when we
			// open file and display all unit numbers
			// like number-text in that file these numbers can overlap like there https://disk.yandex.ru/d/G_GaAB2yH9OMDg
			if (shape.getCellNumberValue("HideText") === 1) {
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
			textCShape.setVerticalAlign(1); // sets text vert align center. equal to anchor set to txBody bodyPr
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
					// but baseMatrix transformations changes values to
					// 0 - top, 1, 2, 3 - center, 4 - bottom
					// NO REVERT NOW TOP IS TOP BOTTOM IS BOTTOM
					if (verticalAlign === 0) {
						textCShape.setVerticalAlign(4); // sets text vert align center equal to anchor set to txBody bodyPr
					} else if (verticalAlign === 2) {
						textCShape.setVerticalAlign(0); // sets text vert align center equal to anchor set to txBody bodyPr
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
			// oTextPr.FontSize = nFontSize * scale;
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

			let leftMarginInch = shape.getCellNumberValueWithScale("LeftMargin", drawingPageScale);
			let topMarginInch = shape.getCellNumberValueWithScale("TopMargin", drawingPageScale);
			let rightMarginInch = shape.getCellNumberValueWithScale("RightMargin", drawingPageScale);
			let bottomMarginInch = shape.getCellNumberValueWithScale("BottomMargin", drawingPageScale);


			// CHECKS SIGN but positive tIns gives bottom inset. Check https://disk.yandex.ru/d/IU1vdjzcF9p3IQ
			// it is may global graphics transform issue so set bottom inset as top and opposite
			oBodyPr.tIns = topMarginInch * g_dKoef_in_to_mm;
			oBodyPr.bIns = bottomMarginInch * g_dKoef_in_to_mm;
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
			let txtPinX_inch = shape.getCellNumberValueWithScale("TxtPinX", drawingPageScale);
			let txtPinY_inch = shape.getCellNumberValueWithScale("TxtPinY", drawingPageScale);


			// consider https://disk.yandex.ru/d/2XzRaPTKzKHFjA
			// where TxtHeight and TxtWidth get all shape height and width and txtPinX_inch and txtPinY_inch are not defined
			// also check for {}, undefined, NaN, null
			let oSpPr = new AscFormat.CSpPr();
			let oXfrm = new AscFormat.CXfrm();
			if (!(isNaN(txtPinX_inch) || txtPinX_inch === null)  && !(isNaN(txtPinY_inch) || txtPinY_inch === null)) {
				// https://www.figma.com/file/WiAC4sxQuJaq65h6xppMYC/cloudFare?type=design&node-id=0%3A1&mode=design&t=SZbio0yIyxq0YnMa-1s

				let shapeWidth = shape.getCellNumberValueWithScale("Width", drawingPageScale);
				let shapeHeight = shape.getCellNumberValueWithScale("Height", drawingPageScale);
				let shapeLocPinX = shape.getCellNumberValueWithScale("LocPinX", drawingPageScale);
				let shapeLocPinY = shape.getCellNumberValueWithScale("LocPinY", drawingPageScale);
				let txtWidth_inch = shape.getCellNumberValueWithScale("TxtWidth", drawingPageScale);
				let txtHeight_inch = shape.getCellNumberValueWithScale("TxtHeight", drawingPageScale);
				let txtLocPinX_inch = shape.getCellNumberValueWithScale("TxtLocPinX", drawingPageScale);
				let txtLocPinY_inch = shape.getCellNumberValueWithScale("TxtLocPinY", drawingPageScale);

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

		/**
		 * endArrow can be beginning or ending
		 * @param {string} arrowType
		 * @param {number}  arrowSize
		 * @return {AscFormat.EndArrow} endArrowObject
		 */
		function getEndArrow(arrowType, arrowSize) {
			// 2.4.4.20	BeginArrow in MS-VSDX and 20.1.10.33 ST_LineEndType (Line End Type) in ECMA
			let endArrow = new AscFormat.EndArrow();

			switch (arrowType) {
				case "0":
					endArrow.type = AscFormat.LineEndType.vsdxNone;
					break;
				case "1":
				case "3":
				case "12":
					endArrow.type = AscFormat.LineEndType.vsdxArrow;
					break;
				case "22":
					endArrow.type = endArrow.GetTypeCode("diamond");
					break;
				case "20":
				case "41":
				case "10":
				case "42":
					endArrow.type = endArrow.GetTypeCode("oval");
					break;
				case "5":
				case "8":
				case "17":
				case "19":
					endArrow.type = AscFormat.LineEndType.vsdxStealth;
					break;
				case "9":
					endArrow.type = AscFormat.LineEndType.vsdxDimensionLine;
					break;
				case "2":
				case "4":
				case "6":
				case "13":
				case "14":
				case "15":
				case "16":
				case "18":
					endArrow.type = AscFormat.LineEndType.vsdxTriangle;
					break;
				case "Themed":
					endArrow.type = AscFormat.LineEndType.vsdxNone;
					break;
				case !isNaN(Number(arrowType)) && arrowType:
					// is unhandled number
					endArrow.type = AscFormat.LineEndType.vsdxArrow;
					break;
				default:
					endArrow.type = AscFormat.LineEndType.vsdxNone;
			}

			if (arrowSize >= 0 && arrowSize <= 6 ) {
				let sizeEnumVsdxShift = 3;
				// see AscFormat.LineEndSize
				endArrow.len = arrowSize + sizeEnumVsdxShift;
				endArrow.w = arrowSize + sizeEnumVsdxShift;
			} else {
				console.log("arrowSize unknown:", arrowSize);
				endArrow.len = AscFormat.LineEndSize.vsdxMedium;
				endArrow.w = AscFormat.LineEndSize.vsdxMedium;
			}

			return endArrow;
		}


		// Method start

		// Refact:
		// 1) I guess any cell can be = THEMEVAL() so better to always
		// use Cell_Type.calculateValue method
		// consider sometimes = THEMEVAL() can be replaced not to Themed but
		// to concrete value on save for Cell_Type.v
		// 2) May be create methods on rows sections and shape -
		// this.calculateCellValue("FillBkgnd",this, pageInfo,
		// 			visioDocument.themes, themeValWasUsedFor, true);
		// 3) May be bind arguments to calculateValue function
		// 4) May be move getTextCShape to other file

		// there was case with shape type group with no PinX and PinY
		// https://disk.yandex.ru/d/tl877cuzcRcZYg
		let pinX_inch = this.getCellNumberValueWithScale("PinX", drawingPageScale);
		let pinY_inch = this.getCellNumberValueWithScale("PinY", drawingPageScale);

		let layerProperties = this.getLayerProperties(pageInfo);
		// only if all shape layers are invisible shape is invisible
		let areShapeLayersInvisible = layerProperties["Visible"] === "0";

		let isShapeDeleted = this.del === "1";


		// also check for {}, undefined, NaN, null
		if (isNaN(pinX_inch) || pinX_inch === null || isNaN(pinY_inch) || pinY_inch === null ||
			areShapeLayersInvisible || isShapeDeleted) {
			// console.log('pinX_inch or pinY_inch is NaN for Shape or areShapeLayersInvisible. Its ok sometimes. ' +
				// 'Empty CShape is returned. See original shape: ', this);
			// let's use empty shape
			let emptyCShape = new AscFormat.CShape();
			emptyCShape.setWordShape(false);
			emptyCShape.setBDeleted(false);

			var oSpPr = new AscFormat.CSpPr();
			var oXfrm = new AscFormat.CXfrm();

			oSpPr.setXfrm(oXfrm);
			oXfrm.setParent(oSpPr);

			emptyCShape.setSpPr(oSpPr);
			oSpPr.setParent(emptyCShape);
			emptyCShape.setParent2(visioDocument);
			return {geometryCShape: emptyCShape, textCShape: null};
		}

		let shapeAngle = this.getCellNumberValue("Angle");
		let locPinX_inch = this.getCellNumberValueWithScale("LocPinX", drawingPageScale);
		let locPinY_inch = this.getCellNumberValueWithScale("LocPinY", drawingPageScale);
		let shapeWidth_inch = this.getCellNumberValueWithScale("Width", drawingPageScale);
		let shapeHeight_inch = this.getCellNumberValueWithScale("Height", drawingPageScale);

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

		/** @type CUniFill */
		let uniFillForegndWithPattern = null;
		/**
		 * Fill without pattern applied.We need fill without pattern applied bcs pattern applied can set
		 * NoSolidFill object without color, so we will not be able to calculate handleVariationColor function result.
		 * @type CUniFill */
		let uniFillForegnd = null;

		/** @type CUniFill */
		let	uniFillBkgnd = null;

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

		let gradientEnabledCell = this.getCell("FillGradientEnabled");
		let gradientEnabled;
		if (gradientEnabledCell !== undefined) {
			gradientEnabled = gradientEnabledCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor, true);
		} else {
			gradientEnabled = false;
		}


		// FillGradientDir and FillPattern can tell about gradient type
		// if FillGradient Enabled
		// FillGradientDir defines gradient type (shape) and clolors define colors. If gradient is linear gradient type is complemented with angle.
		// 13 FillGradientDir is path. path cant be set in interface. also like some radial gradient types witch cant be set in interface.
		// FillGradientDir > 13 is linear like FillGradientDir = 0
		//
		// if FillGradientEnabled Disabled
		// FillPattern defines gradient type and colors define. There linear types with different predefined angles, rectandulat and radial gradient types.
		// Rectangular and radial gradient types differs. There are two colors when i set three colors for gradient. Also FillPattern gradients are not listed in
		// interface. Only true patterns.
		//
		// Its better to convert linear FillPattern gradients there. But FillPattern radial gradients seems to be
		// not like FillGradientDir radial gradients but with different colors
		if (gradientEnabled) {
			let fillGradientDir = this.getCellNumberValue("FillGradientDir");

			// global matrix transform: invert Y axis causes 0 is bottom of gradient and 100000 is top
			let invertGradient = true;
			if (fillGradientDir === 3) {
				// radial gradient seems to be handled in another way
				invertGradient = false;
			}

			// now let's come through gradient stops
			let fillGradientStopsSection = this.getSection("FillGradient");
			let rows = fillGradientStopsSection.getElements();
			let fillGradientStops = [];
			for (const rowKey in rows) {
				let row = rows[rowKey];
				if (row.del) {
					continue;
				}

				// has color (CUniColor) and pos from 0 to 100000
				let colorStop = new AscFormat.CGs();

				// calculate color (CUniColor)
				let color = new AscFormat.CUniColor();
				let gradientStopColorCell = row.getCell("GradientStopColor");
				color = gradientStopColorCell.calculateValue(this, pageInfo,
					visioDocument.themes, themeValWasUsedFor, gradientEnabled, rowKey);

				let gradientStopColorTransCell = row.getCell("GradientStopColorTrans");
				let gradientStopColorTransValue = gradientStopColorTransCell.calculateValue(this, pageInfo,
					visioDocument.themes, themeValWasUsedFor, gradientEnabled, rowKey);
				color.RGBA.A = color.RGBA.A * (1 - gradientStopColorTransValue);

				// now let's get pos
				let gradientStopPositionCell = row.getCell("GradientStopPosition");
				let pos = gradientStopPositionCell.calculateValue(this, pageInfo,
					visioDocument.themes, themeValWasUsedFor, gradientEnabled, rowKey);
				pos = invertGradient ? 100000 - pos : pos;

				colorStop.setColor(color);
				colorStop.setPos(pos);

				fillGradientStops.push({Gs : colorStop});
			}

			if (fillGradientDir === 3) {
				// radial
				uniFillForegnd = AscFormat.builder_CreateRadialGradient(fillGradientStops);
			} else {
				// also if fillGradientDir === 0 - linear
				let fillGradientAngleCell = this.getCell("FillGradientAngle");
				// TODO handle multiple gradient types
				let fillGradientAngle = fillGradientAngleCell.calculateValue(this, pageInfo,
					visioDocument.themes, themeValWasUsedFor, gradientEnabled);

				uniFillForegnd = AscFormat.builder_CreateLinearGradient(fillGradientStops, fillGradientAngle);
			}
		} else {
			let fillForegndCell = this.getCell("FillForegnd");
			if (fillForegndCell) {
				// console.log("FillForegnd was found:", fillForegndCell);
				uniFillForegnd = fillForegndCell.calculateValue(this, pageInfo,
					visioDocument.themes, themeValWasUsedFor, gradientEnabled);

				let fillForegndTransValue = this.getCellNumberValue("FillForegndTrans");
				if (!isNaN(fillForegndTransValue)) {
					let fillObj = uniFillForegnd.fill;
					if (fillObj.constructor.name === "CPattFill") {
						// pattern fill
						fillObj.fgClr.color.RGBA.A = fillObj.fgClr.color.RGBA.A * (1 - fillForegndTransValue);
					} else {
						fillObj.color.color.RGBA.A = fillObj.color.color.RGBA.A * (1 - fillForegndTransValue);
					}
				} else {
					console.log("fillForegndTrans value is themed or something. Not calculated for", this);
				}
			}
		}


		let fillBkgndCell = this.getCell("FillBkgnd");
		if (fillBkgndCell) {
			// console.log("FillBkgnd was found:", fillBkgndCell);
			uniFillBkgnd = fillBkgndCell.calculateValue(this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);

			let fillBkgndTransValue = this.getCellNumberValue("FillBkgndTrans");
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

		let lineColorCell = this.getCell("LineColor");
		if (lineColorCell) {
			// console.log("LineColor was found for shape", lineColorCell);
			lineUniFill = lineColorCell.calculateValue(this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);
		} else {
			console.log("LineColor cell for line stroke (border) was not found painting red");
			lineUniFill = AscFormat.CreateUnfilFromRGB(255,0,0);
		}

		// calculate variation before pattern bcs pattern can make NoFillUniFill object without color
		handleQuickStyleVariation(lineUniFill, uniFillForegnd, this, themeValWasUsedFor);

		let lineWidthEmu = null;
		let lineWeightCell = this.getCell("LineWeight");
		if (lineWeightCell) {
			// to cell.v visio always saves inches
			// let lineWeightInches = Number(lineWeightCell.v);
			let lineWeightInches = lineWeightCell.calculateValue(this, pageInfo,
				visioDocument.themes, themeValWasUsedFor);
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

		// not scaling lineWidth
		/**	 * @type {CLn}	 */
		let oStroke = AscFormat.builder_CreateLine(lineWidthEmu, {UniFill: lineUniFill});

		let linePattern = this.getCell("LinePattern");
		if (linePattern) {
			// see ECMA-376-1 - L.4.8.5.2 Line Dash Properties and [MS-VSDX]-220215 (1) - 2.4.4.180	LinePattern
			let linePatternNumber = linePattern.calculateValue(this, pageInfo, visioDocument.themes, themeValWasUsedFor);
			if (isNaN(linePatternNumber)) {
				oStroke.setPrstDash(oStroke.GetDashCode("vsdxSolid"));
			} else {
				let shift = 11;
				let dashTypeName = oStroke.GetDashByCode(linePatternNumber + shift);
				if (dashTypeName !== null) {
					oStroke.setPrstDash(linePatternNumber + shift);
				} else {
					oStroke.setPrstDash(oStroke.GetDashCode("vsdxDash"));
				}
			}
		}

		let endArrowTypeCell = this.getCell("EndArrow");
		let endArrowSizeCell = this.getCell("EndArrowSize");
		let endArrowType = endArrowTypeCell ? endArrowTypeCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor) : 0;
		let endArrowSize = endArrowSizeCell ? endArrowSizeCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor) : 1;
		let endArrow = getEndArrow(endArrowType, endArrowSize);
		oStroke.setTailEnd(endArrow);

		let beginArrowTypeCell = this.getCell("BeginArrow");
		let beginArrowSizeCell = this.getCell("BeginArrowSize");
		let beginArrowType = beginArrowTypeCell ? beginArrowTypeCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor) : 0;
		let beginArrowSize = beginArrowSizeCell ? beginArrowSizeCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor) : 1;
		let beginArrow = getEndArrow(beginArrowType, beginArrowSize);
		oStroke.setHeadEnd(beginArrow);


		let fillPatternTypeCell = this.getCell("FillPattern");
		let fillPatternType = fillPatternTypeCell ? fillPatternTypeCell.calculateValue(this, pageInfo,
			visioDocument.themes, themeValWasUsedFor) : 1;

		if (!isNaN(fillPatternType) && uniFillBkgnd && uniFillForegnd) {
			// https://learn.microsoft.com/ru-ru/office/client-developer/visio/fillpattern-cell-fill-format-section
			let isfillPatternTypeGradient = fillPatternType >= 25 && fillPatternType <= 40;
			if (fillPatternType === 0) {
				uniFillForegndWithPattern = AscFormat.CreateNoFillUniFill();
			} else if (fillPatternType === 1 || isfillPatternTypeGradient) {
				uniFillForegndWithPattern = uniFillForegnd;
			} else if (fillPatternType > 1) {

				function mapVisioFillPatternToOOXML(fillPatternType) {
					// change down to up and up to down bcs of Global matrix inverted
					let upSideDownPatterns = true;
					switch (fillPatternType) {
						case 2:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.dnDiag :
								AscCommon.global_hatch_offsets.upDiag;
						case 3:
							return AscCommon.global_hatch_offsets.cross;
						case 4:
							return AscCommon.global_hatch_offsets.diagCross;
						case 5:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.upDiag :
								AscCommon.global_hatch_offsets.dnDiag;
						case 6:
							return AscCommon.global_hatch_offsets.horz;
						case 7:
							return AscCommon.global_hatch_offsets.vert;
						case 8:
							return AscCommon.global_hatch_offsets.pct60;
						case 9:
							return AscCommon.global_hatch_offsets.pct40;
						case 10:
							return AscCommon.global_hatch_offsets.pct25;
						case 11:
							return AscCommon.global_hatch_offsets.pct20;
						case 12:
							return AscCommon.global_hatch_offsets.pct10;
						case 13:
							return AscCommon.global_hatch_offsets.dkHorz;
						case 14:
							return AscCommon.global_hatch_offsets.dkVert;
						case 15:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.dkUpDiag :
								AscCommon.global_hatch_offsets.dkDnDiag;
						case 16:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.dkDnDiag :
								AscCommon.global_hatch_offsets.dkUpDiag;
						case 17:
							return AscCommon.global_hatch_offsets.smCheck;
						case 18:
							return AscCommon.global_hatch_offsets.trellis;
						case 19:
							return AscCommon.global_hatch_offsets.ltHorz;
						case 20:
							return AscCommon.global_hatch_offsets.ltVert;
						case 21:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.ltUpDiag :
								AscCommon.global_hatch_offsets.ltDnDiag;
						case 22:
							return upSideDownPatterns ? AscCommon.global_hatch_offsets.ltDnDiag :
								AscCommon.global_hatch_offsets.ltUpDiag;
						case 23:
							return AscCommon.global_hatch_offsets.smGrid;
						case 24:
							return AscCommon.global_hatch_offsets.pct50;
						default:
							console.log("patten fill unhandled");
							return AscCommon.global_hatch_offsets.cross;
					}
				}

				let ooxmlFillPatternType = mapVisioFillPatternToOOXML(fillPatternType);
				if (uniFillForegnd.fill instanceof AscFormat.CPattFill) {
					uniFillForegndWithPattern = AscFormat.CreatePatternFillUniFill(ooxmlFillPatternType,
						uniFillForegnd.fill.bgClr, uniFillForegnd.fill.fgClr);
				} else {
					uniFillForegndWithPattern = AscFormat.CreatePatternFillUniFill(ooxmlFillPatternType,
						uniFillBkgnd.fill.color, uniFillForegnd.fill.color);
				}
			}
		} else if (uniFillForegnd) {
			uniFillForegndWithPattern = uniFillForegnd;
		} else {
			console.log("FillForegnd not found for shape", this);
			uniFillForegndWithPattern = AscFormat.CreateNoFillUniFill();
		}

		let flipHorizontally = this.getCellNumberValue("FlipX") === 1;

		let flipVertically = this.getCellNumberValue("FlipY") === 1;

		let cShape = this.convertToCShapeUsingParamsObj({
			x_mm: x_mm, y_mm: y_mm,
			w_mm: shapeWidth_mm, h_mm: shapeHeight_mm,
			rot: shapeAngle,
			oFill: uniFillForegndWithPattern, oStroke: oStroke,
			flipHorizontally: flipHorizontally, flipVertically: flipVertically,
			pageInfo: pageInfo,
			cVisioDocument: visioDocument,
			drawingPageScale : drawingPageScale
		});

		cShape.Id = String(this.iD); // it was string in cShape

		// not scaling fontSize
		let textCShape = getTextCShape(visioDocument.themes[0], this, cShape, lineUniFill, uniFillForegnd, drawingPageScale);

		cShape.recalculate();
		cShape.recalculateLocalTransform(cShape.transform);

		if (textCShape !== null) {
			textCShape.recalculate();
			textCShape.recalculateLocalTransform(textCShape.transform);
		}

		if (this.type === "Foreign") {
			// console.log("Shape has type Foreign and may not be displayed. " +
			// 	"Check shape.elements --> ForeignData_Type obj. See shape:", this);

			let foreignDataObject = this.getForeignDataObject();
			if (foreignDataObject) {
				if (this.cImageShape !== null) {
					this.cImageShape.setLocks(0);
					this.cImageShape.setBDeleted(false);
					this.cImageShape.setSpPr(cShape.spPr.createDuplicate());
					this.cImageShape.spPr.setParent(this.cImageShape);
					this.cImageShape.rot = cShape.rot;
					// this.cImageShape.brush = cShape.brush;
					this.cImageShape.bounds = cShape.bounds;
					this.cImageShape.flipH = cShape.flipH;
					this.cImageShape.flipV = cShape.flipV;
					this.cImageShape.localTransform = cShape.localTransform;
					// this.cImageShape.pen = cShape.pen;
					this.cImageShape.Id = cShape.Id;

					this.cImageShape.setParent2(visioDocument);
					this.cImageShape.recalculate();

					cShape = this.cImageShape;
				} else {
					console.log("Unknown error: cImageShape was not initialized on ooxml parse");
				}
			}
		}

		return {geometryCShape: cShape, textCShape: textCShape};
	}

	/**
	 * converts !Shape TypeGroup! To CGroupShape Recursively.
	 * let's say shape can only have subshapes if its Type='Group'.
	 * @memberOf Shape_Type
	 * @param {CVisioDocument} visioDocument
	 * @param {Page_Type} pageInfo
	 * @param {Number} drawingPageScale
	 * @param {CGroupShape?} currentGroupHandling
	 * @return {{cGroupShape: CGroupShape, textCShape: CShape}}
	 */
	Shape_Type.prototype.toCGroupShapeRecursively = function (visioDocument, pageInfo,
															  drawingPageScale, currentGroupHandling) {
		// if we need to create CGroupShape create CShape first then copy its properties to CGroupShape object
		// so anyway create CShapes
		let cShapes = this.toGeometryAndTextCShapes(visioDocument, pageInfo, drawingPageScale);

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
					subShape.toCGroupShapeRecursively(visioDocument, pageInfo, drawingPageScale, currentGroupHandling);
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
					subShape.toCGroupShapeRecursively(visioDocument, pageInfo, drawingPageScale, currentGroupHandling);
				}
			}
			// recalculate positions to local (group) coordinates
			cShapes.geometryCShape.recalculateLocalTransform(cShapes.geometryCShape.transform);
			// cShapes.geometryCShape.recalculateTransformText();
			// cShapes.geometryCShape.recalculateContent();
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

				// recalculate positions to local (group) coordinates
				cShapes.geometryCShape.recalculateLocalTransform(cShapes.geometryCShape.transform);
				// cShapes.geometryCShape.recalculateTransformText();
				// cShapes.geometryCShape.recalculateContent();
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
	 * @param {{x_mm, y_mm, w_mm, h_mm, rot, oFill, oStroke, flipHorizontally, flipVertically, cVisioDocument, drawingPageScale}} paramsObj
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
		let drawingPageScale = paramsObj.drawingPageScale;

		let shapeGeom = AscCommonDraw.getGeometryFromShape(this, drawingPageScale);

		let sType   = "rect";
		let nWidth_mm  = Math.round(w_mm);
		let nHeight_mm = Math.round(h_mm);
		//let oDrawingDocument = new AscCommon.CDrawingDocument();
		let shape = AscFormat.builder_CreateShape(sType, nWidth_mm, nHeight_mm,
			oFill, oStroke, cVisioDocument, cVisioDocument.themes[0], null, false);
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
