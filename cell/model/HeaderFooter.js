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
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	/*
	 * Import
	 * -----------------------------------------------------------------------------
	 */
	const AscBrowser = AscCommon.AscBrowser;
	const History = AscCommon.History;
	const asc = window["Asc"];
	const c_oAscError = asc.c_oAscError;

	//HEADER/FOOTER
	function HeaderFooterField(val, format, text) {
		this.field = val;
		this.format = format;

		this.textField = text;
		this._calculatedText = null;
	}

	HeaderFooterField.prototype.clone = function () {
		let res = new HeaderFooterField();

		res.filed = this.filed;
		res.format = this.format.clone();

		res.textField = this.textField;
		res._calculatedText = this._calculatedText;

		return res;
	};
	HeaderFooterField.prototype.calculateText = function (ws, indexPrintPage, countPrintPages) {
		let res = "";
		let api = window["Asc"]["editor"];
		let printPreviewState = ws && ws.workbook && ws.workbook.printPreviewState;
		let pageSetup;
		if (printPreviewState && printPreviewState.isStart()) {
			pageSetup = printPreviewState.getActivePageSetup();
		}
		if (!pageSetup) {
			pageSetup = ws && ws.model && ws.model.PagePrintOptions && ws.model.PagePrintOptions.pageSetup;
		}

		switch (this.field) {
			case asc.c_oAscHeaderFooterField.pageNumber: {
				let firstPageNumber = (pageSetup && pageSetup.useFirstPageNumber && pageSetup.firstPageNumber) ? pageSetup.firstPageNumber : 1;
				res = indexPrintPage + firstPageNumber + "";
				break;
			}
			case asc.c_oAscHeaderFooterField.pageCount: {
				res = countPrintPages + "";
				break;
			}
			case asc.c_oAscHeaderFooterField.sheetName: {
				res = ws && ws.model && ws.model.sName;
				break;
			}
			case asc.c_oAscHeaderFooterField.fileName: {
				res = api.DocInfo ? api.DocInfo.Title : "";
				break;
			}
			case asc.c_oAscHeaderFooterField.filePath: {

				break;
			}
			case asc.c_oAscHeaderFooterField.date: {
				res = (new Asc.cDate()).getDateString(api);
				break;
			}
			case asc.c_oAscHeaderFooterField.time: {
				res = (new Asc.cDate()).getTimeString(api);
				break;
			}
			case asc.c_oAscHeaderFooterField.lineBreak: {
				//TODO возможно стоит добавлять символ переноса строки к предыдущему параграфу
				res = "\n";
				break;
			}
			case asc.c_oAscHeaderFooterField.text: {
				res = this.textField;
				break;
			}
		}
		this._calculatedText = res;
		return res;
	};

	HeaderFooterField.prototype.pushFormat = function (val) {
		this.format = val;
	};

	HeaderFooterField.prototype.getText = function (ws, indexPrintPage, countPrintPages) {
		return this._calculatedText ? this._calculatedText : this.calculateText(ws, indexPrintPage, countPrintPages);
	};

	HeaderFooterField.prototype.getFormat = function (val) {
		return this.format;
	};

	HeaderFooterField.prototype.getField = function () {
		return this.field;
	};


	function HeaderFooterParser() {
		this.tokens = [];
		this.curTokenPosition = null;
		this.str = null;
		this.font = null;

		this.date = null;

		this.allFontsMap = [];

		this.isCalc = null;
	}

	let c_oPortionPosition = {
		left: 0,
		center: 1,
		right: 2
	};

	let c_nPortionLeftHeader = 0;
	let c_nPortionCenterHeader = 1;
	let c_nPortionRightHeader = 2;
	let c_nPortionLeftFooter = 3;
	let c_nPortionCenterFooter = 4;
	let c_nPortionRightFooter = 5;

	/*Documentation:
		There is no required order in which these codes need to appear.
		The first occurrence of the following codes turns the formatting ON, the second occurrence turns it OFF again:
		strikethrough
		superscript
		subscript
		Superscript and subscript cannot both be ON at same time. Whichever comes first wins and the other is ignored,
		while the first is ON.

		&L - code for "left section" (there are three header / footer locations, "left", "center", and "right"). When two or
		more occurrences of this section marker exist, the contents from all markers are concatenated, in the order of
		appearance, and placed into the left section.
		&P - code for "current page #"
					  &N - code for "total pages"
									&font size - code for "text font size", where font size is a font size in points.
		&K - code for "text font color"
		RGB Color is specified as RRGGBB
		Theme Color is specified as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade
		value, NN is the tint/shade value.

		&S - code for "text strikethrough" on / off
		&X - code for "text super script" on / off
		&Y - code for "text subscript" on / off
		&C - code for "center section". When two or more occurrences of this section marker exist, the contents from all
		markers are concatenated, in the order of appearance, and placed into the center section.
		&D - code for "date"
		&T - code for "time"
		&G - code for "picture as background"
		&U - code for "text single underline"
		&E - code for "double underline"
		&R - code for "right section". When two or more occurrences of this section marker exist, the contents from all
		markers are concatenated, in the order of appearance, and placed into the right section.
		&Z - code for "this workbook's file path"
		&F - code for "this workbook's file name"
		&A - code for "sheet tab name"
		&+ - code for add to page #.
		&- - code for subtract from page #.
		&"font name,font type" - code for "text font name" and "text font type", where font name and font type are
		strings specifying the name and type of the font, separated by a comma. When a hyphen appears in font name,
		it means "none specified". Both of font name and font type can be localized values.
		&"-,Bold" - code for "bold font style"
		&B - also means "bold font style".
		&"-,Regular" - code for "regular font style"
		&"-,Italic" - code for "italic font style"

		&I - also means "italic font style"
		&"-,Bold Italic" code for "bold italic font style"
		&O - code for "outline style"
		&H - code for "shadow style"
	*/
	HeaderFooterParser.prototype.parse = function (date) {
		let c_nText = 0, c_nToken = 1, c_nFontName = 2, c_nFontStyle = 3, c_nFontHeight = 4;

		this.date = date;

		this.font = new AscCommonExcel.Font();
		this.curTokenPosition = c_oPortionPosition.center;
		this.str = "";

		let nState = c_nText;
		let nFontHeight = 0;
		let sFontName = "";
		let sFontStyle = "";

		for (let i = 0; i < date.length; i++) {
			let cChar = date[i];
			switch (nState) {
				case c_nText: {
					switch (cChar) {
						case '&':
							this.pushText();
							nState = c_nToken;
							break;
						case '\n':
							this.pushText();
							this.pushLineBreak();
							break;
						default:
							this.str += cChar;
					}
					break;
				}

				case c_nToken: {
					nState = c_nText;


					switch (cChar) {
						case '&':
							this.str += cChar;
							break;
						case 'L':
							this.setPortion(c_oPortionPosition.left);
							this.font = new AscCommonExcel.Font();
							break;
						case 'C':
							this.setPortion(c_oPortionPosition.center);
							this.font = new AscCommonExcel.Font();
							break;
						case 'R':
							this.setPortion(c_oPortionPosition.right);
							this.font = new AscCommonExcel.Font();
							break;
						case 'P':   //page number
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageNumber));
							break;
						case 'N':   //total page count
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageCount));
							break;
						case 'A':   //current sheet name
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.sheetName));
							break;
						case 'F':   //file name
						{
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.fileName));
							break;
						}
						case 'Z':   //file path
						{
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.filePath));
							break;
						}
						case 'D':   //date
						{
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.date));
							break;
						}
						case 'T':   //time
						{
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.time));
							break;
						}
						case 'B':   //bold
							this.font.b = !this.font.b;
							break;
						case 'I':
							this.font.i = !this.font.i;
							break;
						case 'U':   //underline
							this.font.u = Asc.EUnderline.underlineSingle;
							break;
						case 'E':   //double underline
							this.font.u = Asc.EUnderline.underlineDouble;
							break;
						case 'S':   //strikeout
							this.font.s = !this.font.s;
							break;
						case 'X':   //superscript
							if (this.font.va === AscCommon.vertalign_SuperScript) {
								this.font.va = AscCommon.vertalign_Baseline;
							} else {
								this.font.va = AscCommon.vertalign_SuperScript;
							}
							break;
						case 'Y':   //subsrcipt
							if (this.font.va === AscCommon.vertalign_SubScript) {
								this.font.va = AscCommon.vertalign_Baseline;
							} else {
								this.font.va = AscCommon.vertalign_SubScript;
							}
							break;
						case 'O':   //outlined

							break;
						case 'H':   //shadow

							break;
						case 'K':   //text color
							if (i + 6 < date.length) {
								// eat the following 6 characters
								this.font.c = this.convertFontColor(date.substr(i + 1, 6));
								i += 6;
							}
							break;
						case '\"':  //font name
							sFontName = "";
							sFontStyle = "";
							nState = c_nFontName;
							break;
						case 'G':   //picture
							this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.picture));
							break;
						default:
							if (('0' <= cChar) && (cChar <= '9'))    // font size
							{
								nFontHeight = cChar - '0';
								nState = c_nFontHeight;
							}
					}
					break;
				}
				case c_nFontName: {
					switch (cChar) {
						case '\"':
							this.convertFontName(sFontName);
							sFontName = "";
							nState = c_nText;
							break;
						case ',':
							nState = c_nFontStyle;
							break;
						default:
							sFontName += cChar;
					}
					break;
				}
				case c_nFontStyle: {
					switch (cChar) {
						case '\"':
							this.convertFontName(sFontName);
							sFontName = "";
							this.convertFontStyle(sFontStyle);
							sFontStyle = "";

							nState = c_nText;
							break;
						default:
							sFontStyle += cChar;
					}
					break;
				}
				case c_nFontHeight: {
					if (('0' <= cChar) && (cChar <= '9')) {
						if (nFontHeight >= 0) {
							nFontHeight *= 10;
							nFontHeight += (cChar - '0');
							if (nFontHeight > 1000) {
								nFontHeight = -1;
							}
						}
					} else {
						if (nFontHeight > 0) {
							this.font.fs = nFontHeight;
						}
						i--;
						nState = c_nText;
					}
					break;
				}
			}
		}

		this.endPortion();
	};

	HeaderFooterParser.prototype.calculateTokens = function (ws, indexPrintPage, countPrintPages, forceCalc) {
		if (this.isCalc && !forceCalc) {
			return;
		}
		for (let i = 0; i < this.tokens.length; i++) {
			if (this.tokens[i]) {
				for (let j = 0; j < this.tokens[i].length; j++) {
					let token = this.tokens[i][j];
					token && token.calculateText(ws, indexPrintPage, countPrintPages);
				}
			}
		}
		this.isCalc = true;
	};

	HeaderFooterParser.prototype.convertFontColor = function (rColor) {
		let color;
		if ((rColor[2] === '+') || (rColor[2] === '-')) {
			let theme = rColor.substr(0, 2) - 0;
			let tint = rColor.substr(2) - 0;
			color = AscCommonExcel.g_oColorManager.getThemeColor(theme, tint / 100);

		} else {
			color = new AscCommonExcel.RgbColor(AscCommonExcel.g_clipboardExcel.pasteProcessor._getBinaryColor(rColor));
		}
		return color;
	};

	HeaderFooterParser.prototype.convertFontColorFromObj = function (obj) {
		let color = null;

		if (obj instanceof AscCommonExcel.ThemeColor) {
			let theme = obj.theme.toString();
			if (theme.length === 1) {
				theme = "0" + theme;
			}
			let tint = (obj.tint * 100).toFixed(0);
			if (1 === tint.length) {
				tint = "00" + tint;
			} else if (2 === tint.length) {
				tint = "0" + tint;
			}
			color = theme + "+" + tint;
		} else if (obj instanceof AscCommonExcel.RgbColor) {

			let toHex = function componentToHex(c) {
				let res = c.toString(16);
				return res.length === 1 ? "0" + res : res;
			};

			color = toHex(obj.getR()) + toHex(obj.getG()) + toHex(obj.getB());
		} else if (obj === null) {
			color = "01+000";
		}

		return color;
	};

	HeaderFooterParser.prototype.pushText = function () {
		if (0 !== this.str.length) {
			let newTextField = new HeaderFooterField(asc.c_oAscHeaderFooterField.text, this.font.clone(), this.str);
			if (!this.tokens[this.curTokenPosition]) {
				this.tokens[this.curTokenPosition] = [newTextField];
			} else {
				this.tokens[this.curTokenPosition].push(newTextField);
			}

			this.str = [];
		}
	};

	HeaderFooterParser.prototype.pushField = function (field) {
		if (field) {
			field.pushFormat(this.font.clone());
		}
		if (!this.tokens[this.curTokenPosition]) {
			this.tokens[this.curTokenPosition] = [field];
		} else {
			this.tokens[this.curTokenPosition].push(field);
		}
	};

	HeaderFooterParser.prototype.pushLineBreak = function () {
		this.pushField(new HeaderFooterField(asc.c_oAscHeaderFooterField.lineBreak));
	};

	HeaderFooterParser.prototype.convertFontName = function (rName) {
		if ("" !== rName) {
			// single dash is document default font
			if ((rName.length === 1) && (rName[0] === '-')) {
				//пересмотреть
				this.font.fn = null;
			} else {
				this.font.fn = rName;
				this.allFontsMap[rName] = 1;
			}
		}
	};

	HeaderFooterParser.prototype.convertFontStyle = function (rStyle) {
		//в ms жесткая завязка на font style. в lo - ддопускаются следующие строчки - "bold italic bold"  и тп
		this.font.b = this.font.i = false;

		let fontStyleArr = rStyle.split(" ");
		for (let i = 0; i < fontStyleArr.length; i++) {
			if ("italic" === fontStyleArr[i].toLowerCase()) {
				this.font.i = true;
			} else if ("bold" === fontStyleArr[i].toLowerCase()) {
				this.font.b = true;
			}
		}
	};

	HeaderFooterParser.prototype.endPortion = function () {
		this.pushText();
	};

	HeaderFooterParser.prototype.setPortion = function (val) {
		if (val !== this.curTokenPosition) {
			this.endPortion();
			this.curTokenPosition = val;
		}
	};

	HeaderFooterParser.prototype.getTokensByPosition = function (val) {
		return this.tokens && this.tokens[val];
	};

	HeaderFooterParser.prototype.getAllFonts = function (oFontMap) {
		for (let i in this.allFontsMap) {
			if (!oFontMap[i]) {
				oFontMap[i] = 1;
			}
		}
	};

	HeaderFooterParser.prototype.assembleText = function () {
		let newStr = "";
		let curPortionLeft = this.assemblePortionText(c_oPortionPosition.left);
		if (curPortionLeft) {
			newStr += curPortionLeft;
		}
		let curPortionCenter = this.assemblePortionText(c_oPortionPosition.center);
		if (curPortionCenter) {
			newStr += curPortionCenter;
		}
		let curPortionRight = this.assemblePortionText(c_oPortionPosition.right);
		if (curPortionRight) {
			newStr += curPortionRight;
		}
		this.date = newStr;
		return {str: newStr, left: curPortionLeft, center: curPortionCenter, right: curPortionRight};
	};

	HeaderFooterParser.prototype.splitByParagraph = function (cPortionCode) {
		let res = [];

		if (this.tokens[cPortionCode]) {
			let index = 0;
			let curPortion = this.tokens[cPortionCode];
			for (let i = 0; i < curPortion.length; i++) {
				if (!res[index]) {
					res[index] = [];
				}
				res[index].push(curPortion[i]);
			}
		}

		return res;
	};

	HeaderFooterParser.prototype.assemblePortionText = function (cPortion) {
		let symbolPortion;
		switch (cPortion) {
			case c_oPortionPosition.left: {
				symbolPortion = "L";
				break;
			}
			case c_oPortionPosition.center: {
				symbolPortion = "C";
				break;
			}
			case c_oPortionPosition.right: {
				symbolPortion = "R";
				break;
			}
		}

		let compareColors = function (color1, color2) {
			let isEqual = true;

			if (color1 !== color2 || (color1 && color2 && color1.rgb !== color2.rgb)) {
				isEqual = false;
			}

			return isEqual;
		};
		let res = "";
		let fontList = true;

		let aText = "";
		let prevFont = new AscCommonExcel.Font();
		let paragraphs = this.splitByParagraph(cPortion);
		for (let j = 0; j < paragraphs.length; ++j) {
			let aParaText = "";
			let aPosList = paragraphs[j];

			for (let i = 0; i < aPosList.length; ++i) {

				let aFont = aPosList[i].format;

				// font name and style
				let newFont = aPosList[i].format;
				let bNewFontName = !(prevFont.fn === newFont.fn);
				let bNewStyle = (prevFont.b !== newFont.b) || (prevFont.i !== newFont.i);

				if (bNewFontName || (bNewStyle && fontList)) {
					if (null === newFont.fn) {
						aParaText += "&\"" + "-";
					} else {
						aParaText += "&\"" + newFont.fn;
					}

					//TODO пересмотреть. MS каждый раз прописывает новый font style:
					// сли у предыдущего фрагмента был bold, у нового bold и italic - то у нового будет прописаны и bold и italic
					let fontStyleStr = "";
					if (prevFont.b !== newFont.b) {
						fontStyleStr = ",";
						if (newFont.b === true) {
							fontStyleStr += "Bold";
						} else {
							fontStyleStr += "Regular";
						}
					}
					if (prevFont.i !== newFont.i) {
						if ("" === fontStyleStr) {
							fontStyleStr = ",";
						} else {
							fontStyleStr += " ";
						}

						if (newFont.i === true) {
							fontStyleStr += "Italic";
						} else if (-1 === fontStyleStr.indexOf("Regular")) {
							fontStyleStr += "Regular";
						}
					}

					aParaText += fontStyleStr;
					aParaText += "\"";
				}

				//font size
				newFont.fs = aFont.fs;
				let bFontHtChanged = (prevFont.fs !== newFont.fs);
				if (bFontHtChanged) {
					aParaText += "&" + newFont.fs;
				}

				// underline
				if (prevFont.u !== newFont.u) {
					let underline = (newFont.u === Asc.EUnderline.u) ? prevFont.u : newFont.u;
					(underline === Asc.EUnderline.underlineSingle) ? aParaText += "&U" : aParaText += "&E";
				}

				// strikeout
				if (prevFont.s !== newFont.s) {
					aParaText += "&S";
				}

				// super/sub script
				if (prevFont.va !== newFont.va) {
					//aParaText += "&S";

					switch (newFont.va) {
						// close the previous super/sub script.
						case AscCommon.vertalign_SuperScript:
							aParaText += "&X";
							break;
						case AscCommon.vertalign_SubScript:
							aParaText += "&Y";
							break;
						default:
							(prevFont.va === AscCommon.vertalign_SuperScript) ? aParaText += "&X" : aParaText += "&Y";
							break;
					}
				}


				if (!compareColors(prevFont.c, newFont.c)) {
					let newColor = this.convertFontColorFromObj(newFont.c);
					if (null !== newColor) {
						aParaText += "&K";
						aParaText += newColor;
					}
				}

				prevFont = newFont;

				if (aPosList[i] instanceof HeaderFooterField) {
					if (aPosList[i].field !== undefined) {
						switch (aPosList[i].field) {
							case asc.c_oAscHeaderFooterField.pageNumber: {
								aParaText += "&P";
								break;
							}
							case asc.c_oAscHeaderFooterField.pageCount: {
								aParaText += "&N";
								break;
							}
							case asc.c_oAscHeaderFooterField.date: {
								aParaText += "&D";
								break;
							}
							case asc.c_oAscHeaderFooterField.time: {
								aParaText += "&T";
								break;
							}
							case asc.c_oAscHeaderFooterField.sheetName: {
								aParaText += "&A";
								break;
							}
							case asc.c_oAscHeaderFooterField.fileName: {
								aParaText += "&F";
								break;
							}
							case asc.c_oAscHeaderFooterField.filePath: {

								break;
							}
							case asc.c_oAscHeaderFooterField.picture: {
								aParaText += "&G";
								break;
							}
							case asc.c_oAscHeaderFooterField.text: {
								let aPortionText = aPosList[i].getText();
								if (bFontHtChanged && aParaText.length && "" !== aPortionText) {
									let cLast = aParaText[aParaText.length - 1];
									let cFirst = aPortionText[0];
									if (('0' <= cLast) && (cLast <= '9') && ('0' <= cFirst) && (cFirst <= '9')) {
										aParaText += " ";
									}
								}
								aParaText += aPortionText;
								break;
							}
						}
					}
				}
			}

			if (j !== paragraphs.length - 1) {
				aParaText += "\n";
			}
			aText += aParaText;
		}

		if ("" !== aText) {
			res += "&" + symbolPortion + aText;
		}

		return res;
	};


	function CHeaderFooterEditorSection(type, portion, canvasObj) {
		this.type = type;
		this.portion = portion;
		this.canvasObj = canvasObj;
		this.fragments = null;

		this.pictures = null;
		this.loadPictureInfo = null;

		this.changed = false;
	}

	CHeaderFooterEditorSection.prototype.clone = function () {
		let res = new CHeaderFooterEditorSection();

		res.type = this.type;
		res.portion = this.portion;
		res.canvasObj = this.canvasObj;
		res.pictures = this.pictures;
		if (this.fragments) {
			res.fragments = [];
			for (let i = 0; i < this.fragments.length; i++) {
				res.fragments.push(this.fragments[i].clone());
			}
		}

		return res;
	};
	CHeaderFooterEditorSection.prototype.setFragments = function (val) {
		this.fragments = this.isEmptyFragments(val) ? null : val;
	};
	CHeaderFooterEditorSection.prototype.isEmptyFragments = function (val) {
		let res = false;
		if (val && val.length === 1 && val[0].getFragmentText() === "") {
			res = true;
		}
		return res;
	};
	CHeaderFooterEditorSection.prototype.getFragments = function () {
		return this.fragments;
	};
	CHeaderFooterEditorSection.prototype.drawText = function () {
		let t = this;

		if (!this.canvasObj || !this.canvasObj.drawingCtx) {
			return;
		}

		this.canvasObj.drawingCtx.clear();

		let drawBackground = function () {
			t.canvasObj.drawingCtx.setFillStyle(new AscCommon.CColor(255, 255, 255))
				.fillRect(0, 0, t.canvasObj.canvas.width, t.canvasObj.canvas.height);
		};

		if (!this.fragments) {
			//возможно стоит очищать канву в данном случае
			drawBackground();
			return;
		}

		let canvas = this.canvasObj.canvas;
		let width = this.canvasObj.width;
		let drawingCtx = this.canvasObj.drawingCtx;

		//draw
		//добавляю флаги для учета переноса строки
		let wb = window["Asc"]["editor"].wb;
		let ws = window["Asc"]["editor"].wb.getWorksheet();
		let cellFlags = new AscCommonExcel.CellFlags();
		cellFlags.wrapText = true;
		cellFlags.textAlign = this.getAlign();

		//не зависит от зума страницы
		let realZoom = ws.stringRender.drawingCtx.getZoom();
		ws.stringRender.drawingCtx.changeZoom(1);

		let cellEditorWidth = width - 2 * wb.defaults.worksheetView.cells.padding + 1 + 2 * correctCanvasDiff;
		ws.stringRender.setString(this.fragments, cellFlags);
		let textMetrics = ws.stringRender._measureChars(cellEditorWidth);
		let parentHeight = document.getElementById(this.canvasObj.idParent).clientHeight;
		canvas.height = textMetrics.height > parentHeight ? textMetrics.height : AscCommon.AscBrowser.convertToRetinaValue(parentHeight + 1, true);

		drawBackground();
		ws.stringRender.render(drawingCtx, wb.defaults.worksheetView.cells.padding, 0, cellEditorWidth, ws.settings.activeCellBorderColor);

		ws.stringRender.drawingCtx.changeZoom(realZoom)
	};
	CHeaderFooterEditorSection.prototype.getElem = function () {
		return document.getElementById(this.canvasObj.idParent);
	};
	CHeaderFooterEditorSection.prototype.appendEditor = function (editorElemId) {
		let curElem = this.getElem();
		let editorElem = document.getElementById(editorElemId);
		curElem.appendChild(editorElem);
	};
	CHeaderFooterEditorSection.prototype.getAlign = function (portion) {
		portion = undefined !== portion ? portion : this.portion;

		let res = AscCommon.align_Left;
		if (portion === c_nPortionCenterHeader || portion === c_nPortionCenterFooter) {
			res = AscCommon.align_Center;
		} else if (portion === c_nPortionRightHeader || portion === c_nPortionRightFooter) {
			res = AscCommon.align_Right;
		}
		return res;
	};
	CHeaderFooterEditorSection.prototype.getPictures = function () {
		return this.pictures;
	};
	CHeaderFooterEditorSection.prototype.getStringName = function (portion, type) {
		let sPortionPrefix = this.getStringPortion(portion);
		if (sPortionPrefix) {
			let sType = this.getStringType(type);
			if (sType) {
				return sPortionPrefix + sType;
			}
		}
		return null;
	};

	CHeaderFooterEditorSection.prototype.getStringPortion = function (portion) {
		let sPortion = null;
		if (portion == null) {
			portion = this.portion;
		}
		switch (portion) {
			case c_nPortionLeftHeader:
			case c_nPortionLeftFooter: {
				sPortion = "L";
				break;
			}
			case c_nPortionCenterHeader:
			case c_nPortionCenterFooter: {
				sPortion = "C";
				break;
			}
			case c_nPortionRightHeader:
			case c_nPortionRightFooter: {
				sPortion = "R";
				break;
			}
		}

		return sPortion;
	};

	CHeaderFooterEditorSection.prototype.getStringType = function (type) {
		//"LH", "CH", "RH", "LF", "CF", "RF", "LHEVEN",..., "LHFIRST"
		let sType = null;
		if (type == null) {
			type = this.type;
		}
		switch (type) {
			case asc.c_oAscPageHFType.oddFooter: {
				sType = "F";
				break;
			}
			case asc.c_oAscPageHFType.oddHeader: {
				sType = "H";
				break;
			}
			case asc.c_oAscPageHFType.evenHeader: {
				sType = "HEVEN";
				break;
			}
			case asc.c_oAscPageHFType.evenFooter: {
				sType = "FEVEN";
				break;
			}
			case asc.c_oAscPageHFType.firstHeader: {
				sType = "HFIRST";
				break;
			}
			case asc.c_oAscPageHFType.firstFooter: {
				sType = "FFIRST";
				break;
			}
		}

		return sType;
	};


	function convertFieldToMenuText(val, _text) {
		let textField = null;
		let tM = AscCommon.translateManager;
		let pageTag = "&[" + tM.getValue("Page") + "]";
		let pagesTag = "&[" + tM.getValue("Pages") + "]";
		let tabTag = "&[" + tM.getValue("Tab") + "]";
		let dateTag = "&[" + tM.getValue("Date") + "]";
		let fileTag = "&[" + tM.getValue("File") + "]";
		let timeTag = "&[" + tM.getValue("Time") + "]";
		let pictureTag = "&[" + tM.getValue("Picture") + "]";

		switch (val) {
			case asc.c_oAscHeaderFooterField.pageNumber: {
				textField = pageTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.pageCount: {
				textField = pagesTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.date: {
				textField = dateTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.time: {
				textField = timeTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.sheetName: {
				textField = tabTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.fileName: {
				textField = fileTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.filePath: {

				break;
			}
			case asc.c_oAscHeaderFooterField.lineBreak: {
				textField = "\n";
				break;
			}
			case asc.c_oAscHeaderFooterField.picture: {
				//TODO translate?
				textField = pictureTag;
				break;
			}
			case asc.c_oAscHeaderFooterField.text: {
				//TODO translate?
				textField = _text;
				break;
			}
		}
		return textField;
	}

	let correctCanvasDiff = 0;
	window.Asc.g_header_footer_editor = null;

	function CHeaderFooterEditor(idArr, width, pageType, opt_objForSave) {
		this.parentWidth = AscCommon.AscBrowser.convertToRetinaValue(this.correctCanvasWidth(width), true);
		this.parentHeight = AscCommon.AscBrowser.convertToRetinaValue(90, true);
		this.pageType = undefined === pageType ? asc.c_oAscHeaderFooterType.odd : pageType;//odd, even, first
		this.canvas = [];
		this.sections = [];

		this.curParentFocusId = null;
		this.cellEditor = null;
		this.wbCellEditor = null;
		this.editorElemId = "ce-canvas-outer-menu";


		this.api = window["Asc"]["editor"];
		this.wb = this.api.wb;

		this.presets = null;
		this.menuPresets = null;

		this.alignWithMargins = null;
		this.differentFirst = null;
		this.differentOddEven = null;
		this.scaleWithDoc = null;

		this.needAddPicturesMap = null;

		this.startAddImageAction = null;

		if (idArr) {
			window.Asc.g_header_footer_editor = this;
			this.init(idArr, opt_objForSave);
		}
	}

	CHeaderFooterEditor.prototype.init = function (idArr, opt_objForSave) {
		//создаем 6 канвы(+ добавляем их в дом структуру внутрь элемента от меню) + 3 drawingCtx, необходимые для отрисовки 3 поля
		//делается это только 1 раз при инициализации класса
		//потом эти 6 канвы используются для отрисовки всех first/odd/even
		let t = this;
		let createAndPushCanvasObj = function (id) {
			let obj = {};
			obj.idParent = id;
			obj.id = id + "-canvas";
			obj.width = t.parentWidth;
			obj.height = t.parentHeight;
			obj.canvas = document.createElement('canvas');
			obj.canvas.id = obj.id;
			//TODO перепроверить код ниже. оставляю как было раньше
			/*obj.canvas.style.width = t.parentWidth + "px";
			 obj.canvas.style.height = t.parentHeight + "px";
			 AscCommon.calculateCanvasSize(obj.canvas);*/
			obj.canvas.width = t.parentWidth;
			obj.canvas.height = t.parentHeight;
			obj.canvas.style.width = AscCommon.AscBrowser.convertToRetinaValue(t.parentWidth) + "px";

			let curElem = document.getElementById(id);
			curElem.appendChild(obj.canvas);

			obj.drawingCtx = new asc.DrawingContext({
				canvas: obj.canvas, units: 0/*px*/, fmgrGraphics: t.wb.fmgrGraphics, font: t.wb.m_oFont
			});
			return obj;
		};

		this.parentHeight = AscCommon.AscBrowser.convertToRetinaValue(document.getElementById(idArr[0]).clientHeight + 1, true);

		this.canvas[c_nPortionLeftHeader] = createAndPushCanvasObj(idArr[0]);
		this.canvas[c_nPortionCenterHeader] = createAndPushCanvasObj(idArr[1]);
		this.canvas[c_nPortionRightHeader] = createAndPushCanvasObj(idArr[2]);
		this.canvas[c_nPortionLeftFooter] = createAndPushCanvasObj(idArr[3]);
		this.canvas[c_nPortionCenterFooter] = createAndPushCanvasObj(idArr[4]);
		this.canvas[c_nPortionRightFooter] = createAndPushCanvasObj(idArr[5]);


		//add common options
		let optHeaderFooterProps = opt_objForSave && opt_objForSave.headerFooter;
		let ws = this.wb.getWorksheet();
		this.alignWithMargins = optHeaderFooterProps ? optHeaderFooterProps.alignWithMargins : ws.model.headerFooter.alignWithMargins;
		this.differentFirst = optHeaderFooterProps ? optHeaderFooterProps.differentFirst : ws.model.headerFooter.differentFirst;
		this.differentOddEven = optHeaderFooterProps ? optHeaderFooterProps.differentOddEven : ws.model.headerFooter.differentOddEven;
		this.scaleWithDoc = optHeaderFooterProps ? optHeaderFooterProps.scaleWithDoc : ws.model.headerFooter.scaleWithDoc;

		//сохраняем редактор ячейки
		this.wbCellEditor = this.wb.cellEditor;

		//далее создаем классы, где будем хранить fragments всех типов колонтитулов + выполнять отрисовку
		//хранить будем в следующем виде: [c_nPageHFType.firstHeader/.../][c_nPortionLeft/.../c_nPortionRight]
		this._createAndDrawSections(null, optHeaderFooterProps);
		this._generatePresetsArr();

		//лочим
		ws._isLockedHeaderFooter();
	};

	CHeaderFooterEditor.prototype.switchHeaderFooterType = function (type) {
		if (type === this.pageType) {
			return;
		}
		let isError = this._checkSave();
		if (null !== isError) {
			return isError;
		}

		if (this.cellEditor) {
			//save
			let prevField = this._getSectionById(this.curParentFocusId);
			let prevFragments = this.cellEditor.options.fragments;
			prevField.setFragments(prevFragments);
			prevField.drawText();

			prevField.canvasObj.canvas.style.display = "block";

			this.cellEditor.close();
			let elem = document.getElementById(this.editorElemId);
			if (elem) {
				elem.parentNode.removeChild(elem);
			}
		}

		this.curParentFocusId = null;
		this.cellEditor = null;
		this.pageType = type;

		//ещё возможно нужно будет заново добавлять в parent созданную канву(reinit)
		this._createAndDrawSections(type);
	};

	CHeaderFooterEditor.prototype.click = function (id, x, y) {
		let api = this.api;
		let wb = this.wb;
		let t = this;

		let editLockCallback = function () {
			id = id.replace("#", "");

			//если находимся в том же элементе
			if (t.curParentFocusId === id) {
				api.asc_enableKeyEvents(true);
				return;
			}

			//TODO ещё нужно учитывать, что находимся в той же вкладке - odd/even/...
			//если перед этим редактировали другое поле, сохраняем данные
			if (null !== t.curParentFocusId) {
				let prevField = t._getSectionById(t.curParentFocusId);
				let prevFragments = t.cellEditor.options.fragments;
				prevField.setFragments(prevFragments);
				prevField.drawText();

				prevField.canvasObj.canvas.style.display = "block";
			}

			t.curParentFocusId = id;


			let cSection = t._getSectionById(id);
			if (cSection) {
				let sectionElem = cSection.getElem();
				let fragments = cSection.getFragments();
				let self = wb;
				if (!t.cellEditor) {
					t.cellEditor =
						new AscCommonExcel.CellEditor(sectionElem, wb.input, wb.fmgrGraphics, wb.m_oFont, /*handlers*/{
							"closed": function () {
								self.setCellEditMode(false);
							}, "updated": function () {
								self.Api.checkLastWork();
								self._onUpdateCellEditor.apply(self, arguments);
							}, "updateEditorState": function (state) {
								self.handlers.trigger("asc_onEditCell", state);
							}, "updateEditorSelectionInfo": function (xfs) {
								self.handlers.trigger("asc_onEditorSelectionChanged", xfs);
							}, "onContextMenu": function (event) {
								self.handlers.trigger("asc_onContextMenu", event);
							}, "updateMenuEditorCursorPosition": function (pos, height) {
								self.handlers.trigger("asc_updateEditorCursorPosition", pos, height);
							}, "resizeEditorHeight": function () {
								self.handlers.trigger("asc_resizeEditorHeight");
							}, "isActive": function () {
								return wb.isActive();
							}
						}, AscCommon.AscBrowser.convertToRetinaValue(2, true), true);

					//временно меняем cellEditor у wb
					wb.cellEditor = t.cellEditor;

					//удаляем z-index для интерфейса
					t.cellEditor.canvasOuter.style.zIndex = "";
					t.cellEditor.canvas.style.zIndex = "";
					t.cellEditor.canvasOverlay.style.zIndex = "";
					t.cellEditor.cursor.style.zIndex = "";
				} else {
					t.cellEditor.close();
					cSection.appendEditor(t.editorElemId);
				}

				t._openCellEditor(t.cellEditor, fragments, x, y);
				t.cellEditor.canvasOuter.style.zIndex = "";
				cSection.canvasObj.canvas.style.display = "none";

				api.asc_enableKeyEvents(true);
			}
		};

		editLockCallback();
	};

	CHeaderFooterEditor.prototype.correctCanvasWidth = function (val) {
		if (!val) {
			return val;
		}

		correctCanvasDiff = 0;
		if (AscBrowser.retinaPixelRatio === 1.5 && 0 !== val % 4) {
			correctCanvasDiff = val % 4;
			val -= correctCanvasDiff;
		}

		return val;
	};

	CHeaderFooterEditor.prototype._openCellEditor = function (editor, fragments, x, y) {
		let t = this;

		let wb = this.wb;
		let ws = wb.getWorksheet();

		if (!fragments) {
			fragments = [];
			let tempFragment = new AscCommonExcel.Fragment();
			tempFragment.setFragmentText("");
			tempFragment.format = new AscCommonExcel.Font();
			fragments.push(tempFragment);
		}

		let curSection = this._getSectionById(this.curParentFocusId);
		curSection.changed = true;
		let flags = new window["AscCommonExcel"].CellFlags();
		flags.wrapText = true;
		flags.textAlign = curSection.getAlign();

		let enterOptions = new AscCommonExcel.CEditorEnterOptions();
		enterOptions.focus = true;
		if (undefined !== x && undefined !== y) {
			enterOptions.eventPos = {pageX: x, pageY: y};
		}

		let options = {
			enterOptions: enterOptions,
			fragments: fragments,
			flags: flags,
			font: window['AscCommonExcel'].g_oDefaultFormat.Font,
			background: ws.settings.cells.defaultState.background,
			textColor: new window['AscCommonExcel'].RgbColor(0),
			//zoom: this.getZoom(),
			autoComplete: [],
			autoCompleteLC: [],
			saveValueCallback: function (val, flags) {
				//TODO добавил для того, чтобы при нажатии на стрелки не было падения
			},
			getSides: function () {
				let bottomArr = [];
				for (let i = 0; i < 30; i++) {
					bottomArr.push(t.parentHeight + i * 19);
				}
				return {
					l: [0],
					r: [t.parentWidth + 2 * correctCanvasDiff],
					b: bottomArr,
					cellX: 0,
					cellY: 0,
					ri: 0,
					bi: 0
				};
			},
			checkVisible: function () {
				return true;
			},
			menuEditor: true
		};

		wb.setCellEditMode(true);
		editor.open(options);
		wb.input.disabled = false;

		return true;
	};

	CHeaderFooterEditor.prototype.destroy = function (bSave, opt_objForSave) {
		//возвращаем cellEditor у wb
		let t = this;
		let api = window["Asc"]["editor"];
		let wb = api.wb;
		let ws = wb.getWorksheet();

		if (bSave /*&& bChanged*/ && !ws.collaborativeEditing.getGlobalLock() && wb.canEdit()) {
			let checkError = this._checkSave();
			if (null === checkError) {
				wb.cellEditor.close();
				wb.cellEditor = this.wbCellEditor;
				let saveCallback = function (isSuccess) {
					if (false === isSuccess) {
						ws.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.LockedAllError, c_oAscError.Level.NoCritical);
						return;
					}
					t._saveToModel();
				};
				if (opt_objForSave) {
					opt_objForSave.headerFooter = this.getPropsToInterface(opt_objForSave.headerFooter);
				} else {
					ws._isLockedHeaderFooter(saveCallback);
				}
			} else {
				return checkError;
			}
		} else {
			wb.cellEditor.close();
			wb.cellEditor = this.wbCellEditor;
		}
		delete window.Asc.g_header_footer_editor;

		if (t.startAddImageAction) {
			t.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
			t.startAddImageAction = null;
		}

		return null;
	};

	CHeaderFooterEditor.prototype._checkSave = function () {
		let t = this;

		if (null !== this.curParentFocusId) {
			let prevField = this._getSectionById(this.curParentFocusId);
			let prevFragments = this.cellEditor.options.fragments;
			prevField.setFragments(prevFragments);

			prevField.canvasObj.canvas.style.display = "block";
		}

		let checkError = function (type) {
			let prevHeaderFooter = t._getCurPageHF(type);
			let curHeaderFooter = new Asc.CHeaderFooterData();
			curHeaderFooter.parser = new window["AscCommonExcel"].HeaderFooterParser();
			if (prevHeaderFooter && prevHeaderFooter.parser) {
				let newTokens = [];
				for (let i in prevHeaderFooter.parser.tokens) {
					if (prevHeaderFooter.parser.tokens[i]) {
						newTokens[i] = [];
						for (let j in prevHeaderFooter.parser.tokens[i]) {
							let curPortion = prevHeaderFooter.parser.tokens[i][j];
							if (curPortion) {
								newTokens[i][j] = curPortion.clone();
							}
						}
					}
				}
				curHeaderFooter.parser.tokens = newTokens;
			}

			if (t.sections[type][c_oPortionPosition.left] && t.sections[type][c_oPortionPosition.left].changed) {
				curHeaderFooter.parser.tokens[c_oPortionPosition.left] = t._convertFragments(t.sections[type][c_oPortionPosition.left].fragments);
			}
			if (t.sections[type][c_oPortionPosition.center] && t.sections[type][c_oPortionPosition.center].changed) {
				curHeaderFooter.parser.tokens[c_oPortionPosition.center] = t._convertFragments(t.sections[type][c_oPortionPosition.center].fragments);
			}
			if (t.sections[type][c_oPortionPosition.right] && t.sections[type][c_oPortionPosition.right].changed) {
				curHeaderFooter.parser.tokens[c_oPortionPosition.right] = t._convertFragments(t.sections[type][c_oPortionPosition.right].fragments);
			}

			let oData = curHeaderFooter.parser.assembleText();
			if (oData.str && oData.str.length > Asc.c_oAscMaxHeaderFooterLength) {
				let maxLength = oData.left.length;
				let section = c_oPortionPosition.left;
				if (oData.right.length > oData.left.length && oData.right.length > oData.center.length) {
					section = c_oPortionPosition.right;
					maxLength = oData.right.length;
				} else if (oData.center.length > oData.left.length && oData.center.length > oData.right.length) {
					section = c_oPortionPosition.center;
					maxLength = oData.center.length;
				}

				if (t.sections[type] && t.sections[type][section] && t.sections[type][section].canvasObj) {
					return {id: "#" + t.sections[type][section].canvasObj.idParent, max: maxLength};
				}
			}
			return false
		};

		let pageHeaderType = this._getHeaderFooterType(this.pageType);
		let pageFooterType = this._getHeaderFooterType(this.pageType, true);
		let headerCheck = checkError(pageHeaderType);
		let footerCheck = checkError(pageFooterType);
		if (headerCheck && footerCheck) {
			return headerCheck.max > footerCheck.max ? headerCheck.id : footerCheck.id;
		} else if (headerCheck) {
			return headerCheck.id
		} else if (footerCheck) {
			return footerCheck.id
		}

		return null;
	};

	CHeaderFooterEditor.prototype.setPropsFromInterface = function (props, doChanged) {
		let sections = props.sections;
		if (doChanged) {
			let newSections = [];
			for (let i in sections) {
				newSections[i] = [];
				for (let j in sections[i]) {
					newSections[i][j] = sections[i][j] && sections[i][j].clone();
					if (newSections[i][j]) {
						newSections[i][j].changed = true;
					}
				}
			}
			this.sections = newSections;
		} else {
			this.sections = sections;
		}

		this.alignWithMargins = props.alignWithMargins;
		this.differentFirst = props.differentFirst;
		this.differentOddEven = props.differentOddEven;
		this.scaleWithDoc = props.scaleWithDoc;

		this.needAddPicturesMap = props.needAddPicturesMap;
	};

	CHeaderFooterEditor.prototype.getPropsToInterface = function (savedHeaderFooter) {
		let res = savedHeaderFooter ? savedHeaderFooter : {};

		//merge
		if (this.sections) {
			for (let i in this.sections) {
				if (this.sections[i]) {
					for (let j in this.sections[i]) {
						if (!res.sections) {
							res.sections = [];
						}
						if (!res.sections[i]) {
							res.sections[i] = [];
						}
						res.sections[i][j] = this.sections[i][j];
					}
				}
			}
		}

		res.alignWithMargins = this.alignWithMargins;
		res.differentFirst = this.differentFirst;
		res.differentOddEven = this.differentOddEven;
		res.scaleWithDoc = this.scaleWithDoc;
		res.needAddPicturesMap = this.needAddPicturesMap;

		return res;
	};

	CHeaderFooterEditor.prototype._saveToModel = function (ws, opt_headerFooter, reWrite) {
		if (!ws) {
			ws = this.wb && this.wb.getWorksheet();
		}

		let hF = opt_headerFooter ? opt_headerFooter : ws.model.headerFooter;
		let isAddHistory = false;

		if (reWrite) {
			History.Create_NewPoint();
			History.StartTransaction();
			isAddHistory = true;
			hF.clean();
		}

		let removedPictures = [];
		for (let i = 0; i < this.sections.length; i++) {
			if (!this.sections[i]) {
				continue;
			}

			//сначала формируем новый объект, затем доблавляем в модель и записываем в историю полученную строку
			//возможно стоит пересмотреть(получать вначале строку) - создаём вначале парсер,
			//добавляем туда полученные при редактировании фрагменты, затем получаем строку
			let curHeaderFooter = this._getCurPageHF(i, opt_headerFooter, ws);
			if (null === curHeaderFooter) {
				curHeaderFooter = new Asc.CHeaderFooterData();
			}
			if (!curHeaderFooter.parser) {
				curHeaderFooter.parser = new window["AscCommonExcel"].HeaderFooterParser();
			}

			let beforePictures;
			let curSection = this.sections[i][c_oPortionPosition.left];
			let isChanged = false;
			if (curSection && (curSection.changed || reWrite)) {
				beforePictures = this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.left]);
				curHeaderFooter.parser.tokens[c_oPortionPosition.left] = this._convertFragments(curSection.fragments);
				if (beforePictures && !this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.left])) {
					removedPictures.push(curSection.getStringName());
				}
				isChanged = true;
			}
			curSection = this.sections[i][c_oPortionPosition.center];
			if (curSection && (curSection.changed || reWrite)) {
				beforePictures = this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.center]);
				curHeaderFooter.parser.tokens[c_oPortionPosition.center] = this._convertFragments(curSection.fragments);
				if (beforePictures && !this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.center])) {
					removedPictures.push(curSection.getStringName());
				}
				isChanged = true;
			}
			curSection = this.sections[i][c_oPortionPosition.right];
			if (curSection && (curSection.changed || reWrite)) {
				beforePictures = this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.right]);
				curHeaderFooter.parser.tokens[c_oPortionPosition.right] = this._convertFragments(curSection.fragments);
				if (beforePictures && !this.checkPictures(curHeaderFooter.parser.tokens[c_oPortionPosition.right])) {
					removedPictures.push(curSection.getStringName());
				}
				isChanged = true;
			}

			//нужно добавлять в историю
			if (isChanged) {
				if (!isAddHistory && !opt_headerFooter) {
					History.Create_NewPoint();
					History.StartTransaction();
					isAddHistory = true;
				}

				curHeaderFooter.parser.assembleText();
				//curHeaderFooter.setStr(curHeaderFooter.parser.date);
				hF.setHeaderFooterData(curHeaderFooter.parser.date, i);
			}
		}

		//common options
		hF.setAlignWithMargins(this.alignWithMargins);
		hF.setDifferentFirst(this.differentFirst);
		hF.setDifferentOddEven(this.differentOddEven);
		hF.setScaleWithDoc(this.scaleWithDoc);

		//remove pictures
		if (removedPictures.length) {
			ws.removeLegacyDrawingHFPictures(removedPictures);
		}

		//save pictures
		if (ws && ws.changeLegacyDrawingHFPictures && this.needAddPicturesMap) {
			ws.changeLegacyDrawingHFPictures(this.needAddPicturesMap);
		} else if (opt_headerFooter && this.needAddPicturesMap) {
			if (!opt_headerFooter.legacyDrawingHF) {
				opt_headerFooter.legacyDrawingHF = new AscCommonExcel.CLegacyDrawingHF();
			}
			opt_headerFooter.legacyDrawingHF.addPictures(this.needAddPicturesMap, true);
		}

		if (isAddHistory) {
			History.EndTransaction();
		}
	};

	CHeaderFooterEditor.prototype.setFontName = function (fontName) {
		if (null === this.cellEditor) {
			return;
		}

		let t = this, fonts = {};
		fonts[fontName] = 1;
		t.api._loadFonts(fonts, function () {
			t.cellEditor.setTextStyle("fn", fontName);
			t.wb.restoreFocus();
		});
	};

	CHeaderFooterEditor.prototype.checkPictures = function (aFr) {
		if (!aFr) {
			return false;
		}
		for (let i = 0; i < aFr.length; i++) {
			if (aFr[i].field === asc.c_oAscHeaderFooterField.picture) {
				return true;
			}
		}
		return false;
	};

	CHeaderFooterEditor.prototype.setFontSize = function (fontSize) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("fs", fontSize);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setBold = function (isBold) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("b", isBold);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setItalic = function (isItalic) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("i", isItalic);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setUnderline = function (isUnderline) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("u", isUnderline ? Asc.EUnderline.underlineSingle : Asc.EUnderline.underlineNone);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setStrikeout = function (isStrikeout) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("s", isStrikeout);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setSubscript = function (isSubscript) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("fa", isSubscript ? AscCommon.vertalign_SubScript : null);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setSuperscript = function (isSuperscript) {
		if (null === this.cellEditor) {
			return;
		}

		this.cellEditor.setTextStyle("fa", isSuperscript ? AscCommon.vertalign_SuperScript : null);
		this.wb.restoreFocus();
	};

	CHeaderFooterEditor.prototype.setTextColor = function (color) {
		if (null === this.cellEditor) {
			return;
		}

		if (color instanceof Asc.asc_CColor) {
			color = AscCommonExcel.CorrectAscColor(color);
			this.cellEditor.setTextStyle("c", color);
			this.wb.restoreFocus();
		}
	};

	CHeaderFooterEditor.prototype.addField = function (val) {
		if (null === this.cellEditor) {
			return;
		}

		if (val === asc.c_oAscHeaderFooterField.picture) {
			this.addPictureField();
		} else {
			let textField = convertFieldToMenuText(val);
			if (null !== textField) {
				this.cellEditor.pasteText(textField);
			}
		}
	};

	CHeaderFooterEditor.prototype.addPictureField = function () {
		let t = this;
		let showFileDialog = function (needPushField) {
			t.api.asc_addImage({
				callback: function (oImage) {
					if (oImage) {
						if (!t.needAddPicturesMap) {
							t.needAddPicturesMap = {};
						}
						t.needAddPicturesMap[curSection.getStringName()] = oImage;
						if (needPushField) {
							let textField = convertFieldToMenuText(asc.c_oAscHeaderFooterField.picture);
							if (null !== textField) {
								t.cellEditor.pasteText(textField);
							}
						}
					}
					if (t.startAddImageAction) {
						t.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
						t.startAddImageAction = null;
					}
				}, fErrorCallback: function (error) {
					if (t.startAddImageAction) {
						t.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
						t.startAddImageAction = null;
					}
				}, fStartUploadImageCallback: function () {
					t.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
					t.startAddImageAction = true;
				}
			});
		};

		let ws = this.wb && this.wb.getWorksheet();
		let curSection = this._getSectionById(this.curParentFocusId);
		let sectionId = curSection && curSection.getStringName();
		let bSectionContainsPictures = sectionId && (ws.model.getLegacyDrawingHFById(sectionId) || (this.needAddPicturesMap && this.needAddPicturesMap[sectionId]));
		if (bSectionContainsPictures) {
			let thisFragments = this._convertFragments(this.cellEditor.options && this.cellEditor.options.fragments);
			if (!this.checkPictures(thisFragments)) {
				let textField = convertFieldToMenuText(asc.c_oAscHeaderFooterField.picture);
				if (null !== textField) {
					t.cellEditor.pasteText(textField);
				}
			} else {
				//confirm dialog
				//replace/keep options
				t.api.handlers.trigger("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmReplaceHeaderFooterPicture, function (can) {
					if (can) {
						showFileDialog();
					}
				});
			}
		} else {
			//add new pictures
			showFileDialog(true);
		}
	};

	CHeaderFooterEditor.prototype.getTextPresetsArr = function () {
		let wb = this.wb;
		let ws = wb.getWorksheet();

		let arrPresets = this.menuPresets;
		if (!arrPresets) {
			return [];
		}

		let getFragmentText = function (val) {
			return val.getText(ws, 0, 1);
		};

		let getFragmentsText = function (fragments) {
			let res = "";
			for (let n = 0; n < fragments.length; n++) {
				res += getFragmentText(fragments[n]);
			}
			return res;
		};

		let textPresetsArr = [];
		for (let i = 0; i < arrPresets.length; i++) {
			if (!arrPresets[i]) {
				continue;
			}
			textPresetsArr[i] = "";
			for (let j = 0; j < arrPresets[i].length; j++) {
				if (arrPresets[i][j]) {
					let fragments = this._convertFragments([this._getFragments(arrPresets[i][j])]);
					if ("" !== textPresetsArr[i]) {
						textPresetsArr[i] += ", ";
					}
					textPresetsArr[i] += getFragmentsText(fragments);
				}
			}
			if ("" === textPresetsArr[i]) {
				textPresetsArr[i] = AscCommon.translateManager.getValue("None");
			}
		}

		return textPresetsArr;
	};

	CHeaderFooterEditor.prototype.applyPreset = function (type, bFooter) {

		let curType = this._getHeaderFooterType(this.pageType, bFooter);
		let section = this.sections[curType];

		if (this.cellEditor) {
			if (section[c_oPortionPosition.left] && section[c_oPortionPosition.left].canvasObj) {
				this.click(section[c_oPortionPosition.left].canvasObj.idParent);
			}
		}

		this.curParentFocusId = null;

		let fragments;
		for (let i = 0; i < section.length; i++) {
			if (!this.presets[type][i]) {
				section[i].setFragments(null);
			} else {
				fragments = [this._getFragments(this.presets[type][i], new AscCommonExcel.Font())];
				section[i].setFragments(fragments);
			}
			section[i].drawText();
			section[i].changed = true;
			section[i].canvasObj.canvas.style.display = "block";
		}
	};

	CHeaderFooterEditor.prototype.getAppliedPreset = function (type, bFooter) {
		let res = Asc.c_oAscHeaderFooterPresets.none;
		type = undefined !== type ? type : this.pageType;
		let curType = this._getHeaderFooterType(type, bFooter);
		let section = this.sections[curType];

		for (let i = 0; i < section.length; i++) {

			if (null !== section[i].fragments) {
				res = Asc.c_oAscHeaderFooterPresets.custom;
				break;
			}
		}

		return res;
	};

	CHeaderFooterEditor.prototype.setAlignWithMargins = function (val) {
		this.alignWithMargins = val;
	};

	CHeaderFooterEditor.prototype.setDifferentFirst = function (val) {
		let checkError;
		if (!val && (checkError = this._checkSave()) !== null) {
			return checkError;
		}
		this.differentFirst = val;

		return null;
	};

	CHeaderFooterEditor.prototype.setDifferentOddEven = function (val) {
		let checkError;
		if (!val && (checkError = this._checkSave()) !== null) {
			return checkError;
		}
		this.differentOddEven = val;

		return null;
	};

	CHeaderFooterEditor.prototype.setScaleWithDoc = function (val) {
		this.scaleWithDoc = val;
	};

	CHeaderFooterEditor.prototype.getAlignWithMargins = function () {
		return true === this.alignWithMargins || null === this.alignWithMargins;
	};

	CHeaderFooterEditor.prototype.getDifferentFirst = function () {
		return true === this.differentFirst;
	};

	CHeaderFooterEditor.prototype.getDifferentOddEven = function () {
		return true === this.differentOddEven;
	};

	CHeaderFooterEditor.prototype.getScaleWithDoc = function () {
		return true === this.scaleWithDoc || null === this.scaleWithDoc;
	};

	CHeaderFooterEditor.prototype._createAndDrawSections = function (pageCommonType, opt_objForSave, opt_headerFooter) {
		let pageHeaderType = this._getHeaderFooterType(pageCommonType);
		let pageFooterType = this._getHeaderFooterType(pageCommonType, true);

		let getFragments = function (textPropsArr) {
			if (!textPropsArr) {
				return null;
			}
			let _fragments = [];
			let pictures;
			for (let i = 0; i < textPropsArr.length; i++) {
				let curProps = textPropsArr[i];
				let text = convertFieldToMenuText(curProps.field, curProps.textField);
				if (null !== text) {
					let tempFragment = new AscCommonExcel.Fragment();
					tempFragment.setFragmentText(text);
					tempFragment.format = curProps.format;
					_fragments.push(tempFragment);
				}
				if (curProps.field === asc.c_oAscHeaderFooterField.picture) {
					pictures = true;
				}
			}
			return {fragments: _fragments, pictures: pictures};
		};

		let _addFragments = function (_curParser, curSection) {
			let _fragments = getFragments(_curParser);
			if (_fragments && null !== _fragments.fragments) {
				curSection.fragments = _fragments.fragments;
				curSection.pictures = _fragments.pictures;
			}
		};

		//header
		let curPageHF, parser;
		if (!this.sections[pageHeaderType]) {
			this.sections[pageHeaderType] = [];

			//создаём секции, если они уже не созданы
			this.sections[pageHeaderType][c_oPortionPosition.left] = new CHeaderFooterEditorSection(pageHeaderType, c_nPortionLeftHeader, this.canvas[c_nPortionLeftHeader]);
			this.sections[pageHeaderType][c_oPortionPosition.center] = new CHeaderFooterEditorSection(pageHeaderType, c_nPortionCenterHeader, this.canvas[c_nPortionCenterHeader]);
			this.sections[pageHeaderType][c_oPortionPosition.right] = new CHeaderFooterEditorSection(pageHeaderType, c_nPortionRightHeader, this.canvas[c_nPortionRightHeader]);

			//в случае print preview храним временно опции в этом объекте
			if (opt_objForSave) {
				if (opt_objForSave.sections[pageHeaderType][c_oPortionPosition.left]) {
					this.sections[pageHeaderType][c_oPortionPosition.left].fragments = opt_objForSave.sections[pageHeaderType][c_oPortionPosition.left].fragments;
				}
				if (opt_objForSave.sections[pageHeaderType][c_oPortionPosition.center]) {
					this.sections[pageHeaderType][c_oPortionPosition.center].fragments = opt_objForSave.sections[pageHeaderType][c_oPortionPosition.center].fragments;
				}
				if (opt_objForSave.sections[pageHeaderType][c_oPortionPosition.right]) {
					this.sections[pageHeaderType][c_oPortionPosition.right].fragments = opt_objForSave.sections[pageHeaderType][c_oPortionPosition.right].fragments;
				}
			} else {
				//получаем из модели необходимый нам элемент
				curPageHF = this._getCurPageHF(pageHeaderType, opt_headerFooter);
				if (curPageHF && curPageHF.str) {
					if (!curPageHF.parser) {
						curPageHF.parse();
					}
					parser = curPageHF.parser.tokens;
					_addFragments(parser[0], this.sections[pageHeaderType][c_oPortionPosition.left]);
					_addFragments(parser[1], this.sections[pageHeaderType][c_oPortionPosition.center]);
					_addFragments(parser[2], this.sections[pageHeaderType][c_oPortionPosition.right]);
				}
			}
		}

		//footer
		if (!this.sections[pageFooterType]) {
			this.sections[pageFooterType] = [];

			//создаём секции, если они уже не созданы
			this.sections[pageFooterType][c_oPortionPosition.left] = new CHeaderFooterEditorSection(pageFooterType, c_nPortionLeftFooter, this.canvas[c_nPortionLeftFooter]);
			this.sections[pageFooterType][c_oPortionPosition.center] = new CHeaderFooterEditorSection(pageFooterType, c_nPortionCenterFooter, this.canvas[c_nPortionCenterFooter]);
			this.sections[pageFooterType][c_oPortionPosition.right] = new CHeaderFooterEditorSection(pageFooterType, c_nPortionRightFooter, this.canvas[c_nPortionRightFooter]);

			//в случае print preview храним временно опции в этом объекте
			if (opt_objForSave) {
				if (opt_objForSave.sections[pageFooterType][c_oPortionPosition.left]) {
					this.sections[pageFooterType][c_oPortionPosition.left].fragments = opt_objForSave.sections[pageFooterType][c_oPortionPosition.left].fragments;
				}
				if (opt_objForSave.sections[pageFooterType][c_oPortionPosition.center]) {
					this.sections[pageFooterType][c_oPortionPosition.center].fragments = opt_objForSave.sections[pageFooterType][c_oPortionPosition.center].fragments;
				}
				if (opt_objForSave.sections[pageFooterType][c_oPortionPosition.right]) {
					this.sections[pageFooterType][c_oPortionPosition.right].fragments = opt_objForSave.sections[pageFooterType][c_oPortionPosition.right].fragments;
				}
			} else {
				//получаем из модели необходимый нам элемент
				curPageHF = this._getCurPageHF(pageFooterType, opt_headerFooter);
				if (curPageHF && curPageHF.str) {
					if (!curPageHF.parser) {
						curPageHF.parse();
					}
					parser = curPageHF.parser.tokens;
					_addFragments(parser[0], this.sections[pageFooterType][c_oPortionPosition.left]);
					_addFragments(parser[1], this.sections[pageFooterType][c_oPortionPosition.center]);
					_addFragments(parser[2], this.sections[pageFooterType][c_oPortionPosition.right]);
				}
			}
		}

		//DRAW AFTER OPEN MENU
		this.sections[pageHeaderType][c_oPortionPosition.left].drawText();
		this.sections[pageHeaderType][c_oPortionPosition.center].drawText();
		this.sections[pageHeaderType][c_oPortionPosition.right].drawText();
		this.sections[pageFooterType][c_oPortionPosition.left].drawText();
		this.sections[pageFooterType][c_oPortionPosition.center].drawText();
		this.sections[pageFooterType][c_oPortionPosition.right].drawText();
	};

	CHeaderFooterEditor.prototype._getHeaderFooterType = function (type, bFooter) {
		let res = bFooter ? asc.c_oAscPageHFType.oddFooter : asc.c_oAscPageHFType.oddHeader;

		if (type === asc.c_oAscHeaderFooterType.first) {
			res = bFooter ? asc.c_oAscPageHFType.firstFooter : asc.c_oAscPageHFType.firstHeader;
		} else if (type === asc.c_oAscHeaderFooterType.even) {
			res = bFooter ? asc.c_oAscPageHFType.evenFooter : asc.c_oAscPageHFType.evenHeader;
		}

		return res;
	};

	CHeaderFooterEditor.prototype._getCurPageHF = function (type, opt_headerFooter, ws) {
		let res = null;
		if (!ws) {
			ws = this.wb.getWorksheet();
		}
		let hF = opt_headerFooter ? opt_headerFooter : ws.model.headerFooter;

		//TODO можно у класса CHeaderFooter реализовать данную функцию
		if (hF) {
			switch (type) {
				case asc.c_oAscPageHFType.firstHeader: {
					res = hF.firstHeader;
					break;
				}
				case asc.c_oAscPageHFType.oddHeader: {
					res = hF.oddHeader;
					break;
				}
				case asc.c_oAscPageHFType.evenHeader: {
					res = hF.evenHeader;
					break;
				}
				case asc.c_oAscPageHFType.firstFooter: {
					res = hF.firstFooter;
					break;
				}
				case asc.c_oAscPageHFType.oddFooter: {
					res = hF.oddFooter;
					break;
				}
				case asc.c_oAscPageHFType.evenFooter: {
					res = hF.evenFooter;
					break;
				}
			}
		}
		return res;
	};

	CHeaderFooterEditor.prototype._getSectionById = function (id) {
		let res = null;
		let type = this._getHeaderFooterType(this.pageType);
		let i;
		if (this.sections && this.sections[type]) {
			for (i = 0; i < this.sections[type].length; i++) {
				if (id === this.sections[type][i].canvasObj.idParent) {
					return this.sections[type][i];
				}
			}
		}
		type = this._getHeaderFooterType(this.pageType, true);
		if (this.sections && this.sections[type]) {
			for (i = 0; i < this.sections[type].length; i++) {
				if (id === this.sections[type][i].canvasObj.idParent) {
					return this.sections[type][i];
				}
			}
		}
		return res;
	};

	CHeaderFooterEditor.prototype._convertFragments = function (fragments) {
		if (!fragments) {
			return null;
		}

		//TODO возможно стоит созадавать tokens внутри парсера с элементами Fragments
		let res = [];

		let tM = AscCommon.translateManager;

		let bToken, text, symbol, startToken, tokenText, tokenFormat;
		for (let j = 0; j < fragments.length; j++) {
			text = "";
			for (let n = 0; n < fragments[j].getFragmentText().length; n++) {
				symbol = fragments[j].getFragmentText()[n];
				if (symbol !== "&") {
					text += symbol;
				}

				//если несколько таких символов подряд, ms оставляет 1 как текст
				//пока игнорируем данную ситуацию
				if (symbol === "&") {
					if ("" !== text) {
						res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.text, fragments[j].format, text));
						text = "";
					}

					bToken = true;
					tokenFormat = fragments[j].format;
				} else if (startToken) {
					if (symbol === "]") {
						switch (tokenText.toLowerCase()) {
							case tM.getValue("Page").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageNumber, tokenFormat));
								break;
							}
							case tM.getValue("Pages").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageCount, tokenFormat));
								break;
							}
							case tM.getValue("Date").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.date, tokenFormat));
								break;
							}
							case tM.getValue("Time").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.time, tokenFormat));
								break;
							}
							case tM.getValue("Tab").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.sheetName, tokenFormat));
								break;
							}
							case tM.getValue("File").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.fileName, tokenFormat));
								break;
							}
							case "&[Path]&[File]": {
								text = "";
								break;
							}
							case tM.getValue("Picture").toLowerCase(): {
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.picture, tokenFormat));
								break;
							}
							default: {
								if ("" !== text && j === fragments.length - 1 && n === fragments[j].getFragmentText().length - 1) {
									res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.text, fragments[j].format, text));
									text = "";
								}
								break;
							}
						}
						bToken = false;
						startToken = false;
					} else {
						tokenText += symbol;
					}

					if ("" !== text && j === fragments.length - 1 && n === fragments[j].getFragmentText().length - 1) {
						res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.text, fragments[j].format, text));
					}
				} else if (bToken) {
					//начинаем просматривать аргумент
					if (symbol === "[") {
						startToken = true;
						tokenText = "";
					} else {
						//если за "&" следует спецсимвол
						switch (symbol) {
							case 'l':
							case 'c':
							case 'r':
							case 'b':   //bold
							case 'i':
							case 'u':   //underline
							case 'e':   //double underline
							case 's':   //strikeout
							case 'x':   //superscript
							case 'y':   //subsrcipt
							case 'o':   //outlined
							case 'h':   //shadow
							case 'k':   //text color
							case '\"':  //font name
								break;
							case 'p':   //page number
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageNumber, tokenFormat));
								break;
							}
							case 'n':   //total page count
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.pageCount, tokenFormat));
								break;
							}
							case 'a':   //current sheet name
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.sheetName, tokenFormat));
								break;
							}
							case 'f':   //file name
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.fileName, tokenFormat));
								break;
							}
							case 'z':   //file path
							{
								text = "";
								//res.push((new HeaderFooterField(asc.c_oAscHeaderFooterField.filePath)));
								break;
							}
							case 'd':   //date
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.date, tokenFormat));
								break;
							}
							case 't':   //time
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.time, tokenFormat));
								break;
							}
							case 'g':   //picture
							{
								text = "";
								res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.picture, tokenFormat));
								break;
							}
							default: {
								if ("" !== text && j === fragments.length - 1 && n === fragments[j].getFragmentText().length - 1) {
									res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.text, fragments[j].format, text));
									text = "";
								}
								break;
							}
						}
						bToken = false;
					}
				} else if ("" !== text && n === fragments[j].getFragmentText().length - 1) {
					res.push(new HeaderFooterField(asc.c_oAscHeaderFooterField.text, fragments[j].format, text));
				}
			}
		}
		return res;
	};

	CHeaderFooterEditor.prototype._getFragments = function (text, format) {
		let tempFragment = new AscCommonExcel.Fragment();
		tempFragment.setFragmentText(text);
		tempFragment.format = format;
		return tempFragment;
	};

	CHeaderFooterEditor.prototype._generatePresetsArr = function () {
		let docInfo = window["Asc"]["editor"].DocInfo;
		let userInfo = docInfo ? docInfo.get_UserInfo() : null;
		let userName = userInfo ? userInfo.get_FullName() : "";
		let fileName = docInfo ? docInfo.get_Title() : "";

		let tM = AscCommon.translateManager;
		let confidential = tM.getValue("Confidential");
		let preparedBy = tM.getValue("Prepared by ");
		let page = tM.getValue("Page");
		let pageOf = tM.getValue("Page %1 of %2");

		let pageTag = "&[" + page + "]";
		let pagesTag = "&[" + tM.getValue("Pages") + "]";
		let tabTag = "&[" + tM.getValue("Tab") + "]";
		let dateTag = "&[" + tM.getValue("Date") + "]";
		let fileTag = "&[" + tM.getValue("File") + "]";

		let arrPresets = [];
		let arrPresetsMenu = [];
		arrPresets[0] = arrPresetsMenu[0] = [null, null, null];
		arrPresets[1] = arrPresetsMenu[1] = [null, page + " " + pageTag, null];
		arrPresets[2] = [null, pageOf.replace("%1", pageTag).replace("%2", pagesTag), null];
		arrPresetsMenu[2] = [null, pageOf.replace("%1", pageTag).replace("%2", "?"), null];
		arrPresets[3] = arrPresetsMenu[3] = [null, tabTag, null];
		arrPresets[4] = arrPresetsMenu[4] = [confidential, dateTag, page + " " + pageTag];
		arrPresets[5] = arrPresetsMenu[5] = [null, fileTag, null];
		//arrPresets[6] = [null, "&[Path]&[File]", null];
		arrPresets[6] = arrPresetsMenu[6] = [null, tabTag, page + " " + pageTag];
		arrPresets[7] = arrPresetsMenu[7] = [tabTag, confidential, page + " " + pageTag];
		arrPresets[8] = arrPresetsMenu[8] = [null, fileTag, page + " " + pageTag];
		//arrPresets[10] = [null,"&[Path]&[File]","Page &[Page]"];
		arrPresets[9] = arrPresetsMenu[9] = [null, page + " " + pageTag, tabTag];
		arrPresets[10] = arrPresetsMenu[10] = [null, page + " " + pageTag, fileName];
		arrPresets[11] = arrPresetsMenu[11] = [null, page + " " + pageTag, fileTag];
		//arrPresets[12] = [null,"Page &[Page]","&[Path]&[File]"];
		arrPresets[12] = arrPresetsMenu[12] = [userName, page + " " + pageTag, dateTag];
		arrPresets[13] = arrPresetsMenu[13] = [null, preparedBy + userName + " " + dateTag, page + " " + pageTag];

		this.presets = arrPresets;
		this.menuPresets = arrPresetsMenu;
	};


	CHeaderFooterEditor.prototype.getPageType = function () {
		return this.pageType;
	};


	function CLegacyDrawingHF(ws) {
		this.drawings = [];
		this.cfe = null;
		this.cff = null;
		this.cfo = null;
		this.che = null;
		this.chf = null;
		this.cho = null;
		this.lfe = null;
		this.lff = null;
		this.lfo = null;
		this.lhe = null;
		this.lhf = null;
		this.lho = null;
		this.rfe = null;
		this.rff = null;
		this.rfo = null;
		this.rhe = null;
		this.rhf = null;
		this.rho = null;

		this.id = null;
		this.ws = ws;
	}

	CLegacyDrawingHF.prototype.init = function () {

	};

	CLegacyDrawingHF.prototype.addPictures = function (picturesMap, notDeletePictures) {
		let t = this;
		let api = window["Asc"]["editor"];

		for (let i in picturesMap) {
			if (picturesMap.hasOwnProperty(i)) {
				let sSectionId = i;
				let oldHFDrawing = t.getDrawingById(sSectionId);
				let newHFDrawing = new CLegacyDrawingHFDrawing();
				newHFDrawing.id = sSectionId;

				const ws = api.wb.getWorksheet();

				let url = picturesMap[i].src;
				let image = picturesMap[i].Image;

				var __w = Math.max(parseInt(image.width * AscCommon.g_dKoef_pix_to_mm), 1);
				var __h = Math.max(parseInt(image.height * AscCommon.g_dKoef_pix_to_mm), 1);

				newHFDrawing.graphicObject = ws.objectRender.controller.createImage(url, 0, 0, __w, __h);
				t.changePicture(oldHFDrawing && oldHFDrawing.obj, newHFDrawing, true);

				if (!notDeletePictures) {
					delete picturesMap[i];
				}
			}
		}
	};

	CLegacyDrawingHF.prototype.removePictures = function (aPictures) {
		let t = this;
		let api = window["Asc"]["editor"];

		for (let i = 0; i < aPictures.length; i++) {
			let oLegacyDrawingHFDrawing = t.getDrawingById(aPictures[i]);
			if (oLegacyDrawingHFDrawing) {
				t.changePicture(oLegacyDrawingHFDrawing.obj, null, true);
			}
		}
	};

	CLegacyDrawingHF.prototype.changePicture = function (from, to, addToHistory) {
		if (from) {
			this.removePicture(from, addToHistory);
		}
		if (to) {
			this.addPicture(to, addToHistory);
		}

		if ((from || to) && addToHistory && this.ws) {
			let fromData = from && new AscCommonExcel.UndoRedoData_LegacyDrawingHFDrawing(from.id, from.graphicObject.Id);
			let toData = to && new AscCommonExcel.UndoRedoData_LegacyDrawingHFDrawing(to.id, to.graphicObject.Id);
			History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeLegacyDrawingHFDrawing, this.ws.getId(),
				null, new AscCommonExcel.UndoRedoData_FromTo(fromData, toData));
		}
	};

	CLegacyDrawingHF.prototype.removePicture = function (picture) {
		for (let i = 0; i < this.drawings.length; i++) {
			if (this.drawings[i].id === picture.id) {

				this.drawings.splice(i);
				break;
			}
		}
	};

	CLegacyDrawingHF.prototype.addPicture = function (picture) {
		this.drawings.unshift(picture);
	};

	CLegacyDrawingHF.prototype.getDrawingById = function (id) {
		let res = null;

		for (let i = 0; i < this.drawings.length; i++) {
			if (this.drawings[i].id === id) {
				res = {obj: this.drawings[i], index: i};
				break;
			}
		}

		return res;
	};


	function CLegacyDrawingHFDrawing() {
		this.id = null;//"LH", "CH", "RH", "LF", "CF", "RF", "LHEVEN",..., "LHFIRST"
		this.graphicObject = null;
	}

	CLegacyDrawingHFDrawing.prototype.init = function () {

	};


	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].HeaderFooterParser = HeaderFooterParser;
	window["AscCommonExcel"].CHeaderFooterEditorSection = CHeaderFooterEditorSection;

	window["Asc"]["asc_CHeaderFooterEditor"] = window["AscCommonExcel"].CHeaderFooterEditor = CHeaderFooterEditor;
	let prot = CHeaderFooterEditor.prototype;
	prot["click"] = prot.click;
	prot["destroy"] = prot.destroy;
	prot["setFontName"] = prot.setFontName;
	prot["setFontSize"] = prot.setFontSize;
	prot["setBold"] = prot.setBold;
	prot["setItalic"] = prot.setItalic;
	prot["setUnderline"] = prot.setUnderline;
	prot["setStrikeout"] = prot.setStrikeout;
	prot["setSubscript"] = prot.setSubscript;
	prot["setSuperscript"] = prot.setSuperscript;
	prot["setTextColor"] = prot.setTextColor;
	prot["addField"] = prot.addField;
	prot["switchHeaderFooterType"] = prot.switchHeaderFooterType;
	prot["getTextPresetsArr"] = prot.getTextPresetsArr;
	prot["applyPreset"] = prot.applyPreset;
	prot["getAppliedPreset"] = prot.getAppliedPreset;

	prot["setAlignWithMargins"] = prot.setAlignWithMargins;
	prot["setDifferentFirst"] = prot.setDifferentFirst;
	prot["setDifferentOddEven"] = prot.setDifferentOddEven;
	prot["setScaleWithDoc"] = prot.setScaleWithDoc;
	prot["getAlignWithMargins"] = prot.getAlignWithMargins;
	prot["getDifferentFirst"] = prot.getDifferentFirst;
	prot["getDifferentOddEven"] = prot.getDifferentOddEven;
	prot["getScaleWithDoc"] = prot.getScaleWithDoc;

	prot["getPageType"] = prot.getPageType;

	window["AscCommonExcel"].CLegacyDrawingHF = CLegacyDrawingHF;
	window["AscCommonExcel"].CLegacyDrawingHFDrawing = CLegacyDrawingHFDrawing;

	window['AscCommonExcel'].c_oPortionPosition = c_oPortionPosition;

})(window);
