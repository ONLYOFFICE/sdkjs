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
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function(window, undefined) {



	const MAX_LABELS_COUNT = 300;
	const MAX_SERIES_COUNT = 255;
	const MAX_POINTS_COUNT = 4096;
	const MIN_STOCK_COUNT = 4;
// Import
	const c_oAscChartType = AscCommon.c_oAscChartType;
	const c_oAscChartSubType = AscCommon.c_oAscChartSubType;
	const g_oIdCounter = AscCommon.g_oIdCounter;
	const g_oTableId = AscCommon.g_oTableId;
	const oNumFormatCache = AscCommon.oNumFormatCache;
	const isRealObject = AscCommon.isRealObject;
	const History = AscCommon.History;
	const global_MatrixTransformer = AscCommon.global_MatrixTransformer;

	const CShape = AscFormat.CShape;
	const Ax_Counter = AscFormat.Ax_Counter;
	const checkTxBodyDefFonts = AscFormat.checkTxBodyDefFonts;

	const c_oAscTickLabelsPos = Asc.c_oAscTickLabelsPos;
	const c_oAscChartLegendShowSettings = Asc.c_oAscChartLegendShowSettings;
	const c_oAscTickMark = Asc.c_oAscTickMark;

	const EFFECT_NONE = 0;
	const EFFECT_SUBTLE = 1;
	const EFFECT_MODERATE = 2;
	const EFFECT_INTENSE = 3;

	let CHART_STYLE_MANAGER = null;
	const SKIP_LBL_LIMIT = 100;

	const BAR_SHAPE_CONE = 0;
	const BAR_SHAPE_CONETOMAX = 1;
	const BAR_SHAPE_BOX = 2;
	const BAR_SHAPE_CYLINDER = 3;
	const BAR_SHAPE_PYRAMID = 4;
	const BAR_SHAPE_PYRAMIDTOMAX = 5;

	const DISP_BLANKS_AS_GAP = 0;
	const DISP_BLANKS_AS_SPAN = 1;
	const DISP_BLANKS_AS_ZERO = 2;


	const DEFAULT_LBLS_DISTANCE = 10.0 * (25.4 / 72);///TODO

	function checkVerticalTitle(title) {
		return false;
	}

	function checkNoFillMarkers(symbol) {
		return symbol === AscFormat.SYMBOL_X || symbol === AscFormat.SYMBOL_STAR || symbol === AscFormat.SYMBOL_PLUS;
	}

	function GetTextPrFormArrObjects(aObjects, bFirstBreak, bLbl) {
		var oResultTextPr;
		for (var i = 0; i < aObjects.length; ++i) {
			var oContent = aObjects[i];
			if (!oContent)
				continue;
			oContent = bLbl ? oContent.compiledDlb && oContent.compiledDlb.txBody && oContent.compiledDlb.txBody.content : oContent.txBody && oContent.txBody.content;
			if (!oContent)
				continue;
			oContent.SetApplyToAll(true);
			var oTextPr = oContent.GetCalculatedTextPr();
			oContent.SetApplyToAll(false);
			if (!oResultTextPr) {
				oResultTextPr = oTextPr;
				if (bFirstBreak) {
					return oResultTextPr;
				}
			} else {
				oResultTextPr.Compare(oTextPr);
			}
		}
		return oResultTextPr;
	}


	function checkBlackUnifill(unifill, bLines) {
		if (unifill && unifill.fill && unifill.fill.color) {
			var RGBA = unifill.fill.color.RGBA;
			if (RGBA.R === 0 && RGBA.G === 0 && RGBA.B === 0) {
				if (bLines) {
					RGBA.R = 134;
					RGBA.G = 134;
					RGBA.B = 134;
				} else {
					RGBA.R = 255;
					RGBA.G = 255;
					RGBA.B = 255;
				}
			}
		}
	}

	function CRect(x, y, w, h) {
		this.x = x;
		this.y = y;
		this.w = w;
		this.h = h;
		this.fHorPadding = 0.0;
		this.fVertPadding = 0.0;
	}

	CRect.prototype.copy = function () {
		var ret = new CRect(this.x, this.y, this.w, this.h);
		ret.fHorPadding = this.fHorPadding;
		ret.fVertPadding = this.fVertPadding;
		return ret;
	};
	CRect.prototype.intersection = function (oRect) {
		if (this.x + this.w < oRect.x || oRect.x + oRect.w < this.x
			|| this.y + this.h < oRect.y || oRect.y + oRect.h < this.y) {
			return false;
		}

		var x0, y0, x1, y1;
		var bResetHorPadding = true, bResetVertPadding = true;
		if (this.fHorPadding > 0.0 && oRect.fHorPadding > 0.0) {
			x0 = this.x + oRect.fHorPadding;
			bResetHorPadding = false;
		} else {
			x0 = Math.max(this.x, oRect.x);
		}
		if (this.fVertPadding > 0.0 && oRect.fVertPadding > 0.0) {
			y0 = this.y + oRect.fVertPadding;
			bResetVertPadding = false;
		} else {
			y0 = Math.max(this.y, oRect.y);
		}
		if (this.fHorPadding < 0.0 && oRect.fHorPadding < 0.0) {
			x1 = this.x + this.w + oRect.fHorPadding;
			bResetHorPadding = false;
		} else {
			x1 = Math.min(this.x + this.w, oRect.x + oRect.w);
		}
		if (this.fVertPadding < 0.0 && oRect.fVertPadding < 0.0) {
			y1 = this.y + this.h + oRect.fVertPadding;
			bResetVertPadding = false;
		} else {
			y1 = Math.min(this.y + this.h, oRect.y + oRect.h);
		}
		if (bResetHorPadding) {
			this.fHorPadding = 0.0;
		}
		if (bResetVertPadding) {
			this.fVertPadding = 0.0;
		}
		this.x = x0;
		this.y = y0;
		this.w = x1 - x0;
		this.h = y1 - y0;
		return true;

	};

	function CreateUnifillSolidFillSchemeColorByIndex(index) {
		var ret = new AscFormat.CUniFill();
		ret.setFill(new AscFormat.CSolidFill());
		ret.fill.setColor(new AscFormat.CUniColor());
		ret.fill.color.setColor(new AscFormat.CSchemeColor());
		ret.fill.color.color.setId(index);
		return ret;
	}

	function CChartStyleManager() {
		this.styles = [];
	}

	CChartStyleManager.prototype =
		{
			init: function () {
				AscFormat.ExecuteNoHistory(
					function () {
						var DefaultDataPointPerDataPoint =
							[
								[
									CreateUniFillSchemeColorWidthTint(8, 0.885),
									CreateUniFillSchemeColorWidthTint(8, 0.55),
									CreateUniFillSchemeColorWidthTint(8, 0.78),
									CreateUniFillSchemeColorWidthTint(8, 0.925),
									CreateUniFillSchemeColorWidthTint(8, 0.7),
									CreateUniFillSchemeColorWidthTint(8, 0.3)
								],
								[
									CreateUniFillSchemeColorWidthTint(0, 0),
									CreateUniFillSchemeColorWidthTint(1, 0),
									CreateUniFillSchemeColorWidthTint(2, 0),
									CreateUniFillSchemeColorWidthTint(3, 0),
									CreateUniFillSchemeColorWidthTint(4, 0),
									CreateUniFillSchemeColorWidthTint(5, 0)
								],
								[
									CreateUniFillSchemeColorWidthTint(0, -0.5),
									CreateUniFillSchemeColorWidthTint(1, -0.5),
									CreateUniFillSchemeColorWidthTint(2, -0.5),
									CreateUniFillSchemeColorWidthTint(3, -0.5),
									CreateUniFillSchemeColorWidthTint(4, -0.5),
									CreateUniFillSchemeColorWidthTint(5, -0.5)
								],
								[
									CreateUniFillSchemeColorWidthTint(8, 0.05),
									CreateUniFillSchemeColorWidthTint(8, 0.55),
									CreateUniFillSchemeColorWidthTint(8, 0.78),
									CreateUniFillSchemeColorWidthTint(8, 0.15),
									CreateUniFillSchemeColorWidthTint(8, 0.7),
									CreateUniFillSchemeColorWidthTint(8, 0.3)
								]
							];
						var s = DefaultDataPointPerDataPoint;
						var f = CreateUniFillSchemeColorWidthTint;
						this.styles[0] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, s[0], EFFECT_SUBTLE, EFFECT_NONE, [], 3, s[0], 7);
						this.styles[1] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, s[1], EFFECT_SUBTLE, EFFECT_NONE, [], 3, s[1], 7);
						for (var i = 2; i < 8; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, [f(i - 2, 0)], EFFECT_SUBTLE, EFFECT_NONE, [], 3, [f(i - 2, 0)], 7);
						}
						this.styles[8] = new CChartStyle(EFFECT_SUBTLE, EFFECT_SUBTLE, s[0], EFFECT_SUBTLE, EFFECT_SUBTLE, [f(12, 0)], 5, s[0], 9);
						this.styles[9] = new CChartStyle(EFFECT_SUBTLE, EFFECT_SUBTLE, s[1], EFFECT_SUBTLE, EFFECT_SUBTLE, [f(12, 0)], 5, s[1], 9);
						for (i = 10; i < 16; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_SUBTLE, EFFECT_SUBTLE, [f(i - 10, 0)], EFFECT_SUBTLE, EFFECT_SUBTLE, [f(12, 0)], 5, [f(i - 10, 0)], 9);
						}
						this.styles[16] = new CChartStyle(EFFECT_MODERATE, EFFECT_INTENSE, s[0], EFFECT_SUBTLE, EFFECT_NONE, [], 5, s[0], 9);
						this.styles[17] = new CChartStyle(EFFECT_MODERATE, EFFECT_INTENSE, s[1], EFFECT_INTENSE, EFFECT_NONE, [], 5, s[1], 9);
						for (i = 18; i < 24; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_MODERATE, EFFECT_INTENSE, [f(i - 18, 0)], EFFECT_SUBTLE, EFFECT_NONE, [], 5, [f(i - 18, 0)], 9);
						}
						this.styles[24] = new CChartStyle(EFFECT_INTENSE, EFFECT_INTENSE, s[0], EFFECT_SUBTLE, EFFECT_NONE, [], 7, s[0], 13);
						this.styles[25] = new CChartStyle(EFFECT_MODERATE, EFFECT_INTENSE, s[1], EFFECT_SUBTLE, EFFECT_NONE, [], 7, s[1], 13);
						for (i = 26; i < 32; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_MODERATE, EFFECT_INTENSE, [f(i - 26, 0)], EFFECT_SUBTLE, EFFECT_NONE, [], 7, [f(i - 26, 0)], 13);
						}
						this.styles[32] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, s[0], EFFECT_SUBTLE, EFFECT_SUBTLE, [f(8, -0.5)], 5, s[0], 9);
						this.styles[33] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, s[1], EFFECT_SUBTLE, EFFECT_SUBTLE, s[2], 5, s[1], 9);
						for (i = 34; i < 40; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_NONE, EFFECT_SUBTLE, [f(i - 34, 0)], EFFECT_SUBTLE, EFFECT_SUBTLE, [f(i - 34, -0.5)], 5, [f(i - 34, 0)], 9);
						}
						this.styles[40] = new CChartStyle(EFFECT_INTENSE, EFFECT_INTENSE, s[3], EFFECT_SUBTLE, EFFECT_NONE, [], 5, s[3], 9);
						this.styles[41] = new CChartStyle(EFFECT_INTENSE, EFFECT_INTENSE, s[1], EFFECT_INTENSE, EFFECT_NONE, [], 5, s[1], 9);
						for (i = 42; i < 48; ++i) {
							this.styles[i] = new CChartStyle(EFFECT_INTENSE, EFFECT_INTENSE, [f(i - 42, 0)], EFFECT_SUBTLE, EFFECT_NONE, [], 5, [f(i - 42, 0)], 9);
						}

						this.defaultLineStyles = [];
						this.defaultLineStyles[0] = new ChartLineStyle(f(15, 0.75), f(15, 0.5), f(15, 0.75), f(15, 0), EFFECT_SUBTLE);
						for (i = 0; i < 32; ++i) {
							this.defaultLineStyles[i] = this.defaultLineStyles[0];
						}
						this.defaultLineStyles[32] = new ChartLineStyle(f(8, 0.75), f(8, 0.5), f(8, 0.75), f(8, 0), EFFECT_SUBTLE);
						this.defaultLineStyles[33] = this.defaultLineStyles[32];
						this.defaultLineStyles[34] = new ChartLineStyle(f(8, 0.75), f(8, 0.5), f(8, 0.75), f(8, 0), EFFECT_SUBTLE);
						for (i = 35; i < 40; ++i) {
							this.defaultLineStyles[i] = this.defaultLineStyles[34];
						}
						this.defaultLineStyles[40] = new ChartLineStyle(f(8, 0.75), f(8, 0.9), f(12, 0), f(12, 0), EFFECT_NONE);
						for (i = 41; i < 48; ++i) {
							this.defaultLineStyles[i] = this.defaultLineStyles[40];
						}
					},
					this, []);
			},
			getStyleByIndex: function (index) {
				if (AscFormat.isRealNumber(index)) {
					return this.styles[(index - 1) % 48];
				}
				return this.styles[1];
			},

			getDefaultLineStyleByIndex: function (index) {
				if (AscFormat.isRealNumber(index)) {
					return this.defaultLineStyles[(index - 1) % 48];
				}
				return this.defaultLineStyles[2];
			}
		};
	CHART_STYLE_MANAGER = new CChartStyleManager();

	function ChartLineStyle(axisAndMajorGridLines, minorGridlines, chartArea, otherLines, floorChartArea) {
		this.axisAndMajorGridLines = axisAndMajorGridLines;
		this.minorGridlines = minorGridlines;
		this.chartArea = chartArea;
		this.otherLines = otherLines;
		this.floorChartArea = floorChartArea;
	}

	function CChartStyle(effect, fill1, fill2, fill3, line1, line2, line3, line4, markerSize) {
		this.effect = effect;
		this.fill1 = fill1;
		this.fill2 = fill2;
		this.fill3 = fill3;

		this.line1 = line1;
		this.line2 = line2;
		this.line3 = line3;
		this.line4 = line4;

		this.markerSize = markerSize;
	}

	function CreateUniFillSchemeColorWidthTint(schemeColorId, tintVal) {
		return AscFormat.ExecuteNoHistory(
			function (schemeColorId, tintVal) {
				return CreateUniFillSolidFillWidthTintOrShade(CreateUnifillSolidFillSchemeColorByIndex(schemeColorId), tintVal);
			},
			this, [schemeColorId, tintVal]);
	}

	function checkFiniteNumber(num) {
		if (AscFormat.isRealNumber(num) && isFinite(num) && num > 0) {
			return num;
		}
		return 0;
	}

	var G_O_VISITED_HLINK_COLOR = CreateUniFillSolidFillWidthTintOrShade(CreateUnifillSolidFillSchemeColorByIndex(10), 0);
	var G_O_HLINK_COLOR = CreateUniFillSolidFillWidthTintOrShade(CreateUnifillSolidFillSchemeColorByIndex(11), 0);
	var G_O_NO_ACTIVE_COMMENT_BRUSH = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(248, 231, 195));
	var G_O_ACTIVE_COMMENT_BRUSH = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(240, 200, 120));


	var CChangesDrawingsBool = AscDFH.CChangesDrawingsBool;
	var CChangesDrawingsLong = AscDFH.CChangesDrawingsLong;
	var CChangesDrawingsDouble = AscDFH.CChangesDrawingsDouble;
	var CChangesDrawingsString = AscDFH.CChangesDrawingsString;
	var CChangesDrawingsObject = AscDFH.CChangesDrawingsObject;
	var CChangesDrawingsObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
	var CChangesDrawingsContent = AscDFH.CChangesDrawingsContent;


	var drawingsChangesMap = window['AscDFH'].drawingsChangesMap;


	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetNvGrFrProps] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetThemeOverride] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ShapeSetBDeleted] = CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetParent] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetChart] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetClrMapOvr] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetDate1904] = CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetExternalData] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetLang] = CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetPivotSource] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetPrintSettings] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetProtection] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetRoundedCorners] = CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetSpPr] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetStyle] = CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetTxPr] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_SetGroup] = CChangesDrawingsObject;

	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_AddUserShape] = CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_RemoveUserShape] = CChangesDrawingsContent;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_ChartStyle] = CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_ChartSpace_ChartColors] = CChangesDrawingsObjectNoId;


	AscDFH.drawingContentChanges[AscDFH.historyitem_ChartSpace_RemoveUserShape] =
		AscDFH.drawingContentChanges[AscDFH.historyitem_ChartSpace_AddUserShape] = function (oClass) {
			return oClass.userShapes;
		};

	function CheckParagraphTextPr(oParagraph, oTextPr) {
		var oParaPr = oParagraph.Pr.Copy();
		var oParaPr2 = new CParaPr();
		var oCopyTextPr = oTextPr.Copy();
		if (oCopyTextPr.FontFamily) {
			oCopyTextPr.RFonts.Set_FromObject(
				{
					Ascii: {
						Name: oCopyTextPr.FontFamily.Name,
						Index: -1
					},
					EastAsia: {
						Name: oCopyTextPr.FontFamily.Name,
						Index: -1
					},
					HAnsi: {
						Name: oCopyTextPr.FontFamily.Name,
						Index: -1
					},
					CS: {
						Name: oCopyTextPr.FontFamily.Name,
						Index: -1
					}
				}
			);
		}
		oParaPr2.DefaultRunPr = oCopyTextPr;
		oParaPr.Merge(oParaPr2);
		oParagraph.Set_Pr(oParaPr);
	}

	function CheckObjectTextPr(oElement, oTextPr, oDrawingDocument) {
		if (oElement) {
			if (!oElement.txPr) {
				oElement.setTxPr(AscFormat.CreateTextBodyFromString("", oDrawingDocument, oElement));
			}
			oElement.txPr.content.Content[0].Set_DocumentIndex(0);

			CheckParagraphTextPr(oElement.txPr.content.Content[0], oTextPr);
			if (oElement.tx && oElement.tx.rich) {
				var aContent = oElement.tx.rich.content.Content;
				for (var i = 0; i < aContent.length; ++i) {
					CheckParagraphTextPr(aContent[i], oTextPr);
				}

				oElement.tx.rich.content.SetApplyToAll(true);
				var oParTextPr = new AscCommonWord.ParaTextPr(oTextPr);
				oElement.tx.rich.content.AddToParagraph(oParTextPr);
				oElement.tx.rich.content.SetApplyToAll(false);
			}
			CheckParagraphTextPr(oElement.txPr.content.Content[0], oTextPr);
		}
	}

	function CheckIncDecFontSize(oElement, bIncrease, oDrawingDocument, nDefaultSize) {
		if (oElement) {
			if (!oElement.txPr) {
				oElement.setTxPr(AscFormat.CreateTextBodyFromString("", oDrawingDocument, oElement));
			}
			var oParaPr = oElement.txPr.content.Content[0].Pr.Copy();
			oElement.txPr.content.Content[0].Set_DocumentIndex(0);
			var oCopyTextPr;
			if (oParaPr.DefaultRunPr) {
				oCopyTextPr = oParaPr.DefaultRunPr.Copy();
			} else {
				oCopyTextPr = new CTextPr();
			}
			if (!AscFormat.isRealNumber(oCopyTextPr.FontSize)) {
				oCopyTextPr.FontSize = nDefaultSize;
			}
			oCopyTextPr.FontSize = oCopyTextPr.GetIncDecFontSize(bIncrease);
			oParaPr.DefaultRunPr = oCopyTextPr;
			oElement.txPr.content.Content[0].Set_Pr(oParaPr);
			if (oElement.tx && oElement.tx.rich) {
				oElement.tx.rich.content.SetApplyToAll(true);
				oElement.tx.rich.content.IncreaseDecreaseFontSize(bIncrease);
				oElement.tx.rich.content.SetApplyToAll(false);

			}
		}
	}


	function CPathMemory() {
		this.size = 1000;
		this.ArrPathCommand = new Float64Array(this.size);
		this.curPos = -1;

		this.path = new AscFormat.Path2(this);
	}

	CPathMemory.prototype.AllocPath = function () {

		if (this.curPos + 1 >= this.ArrPathCommand.length) {
			var aNewArray = new Float64Array((((3 / 2) * (this.curPos + 1)) >> 0) + 1);
			for (var i = 0; i < this.ArrPathCommand.length; ++i) {
				aNewArray[i] = this.ArrPathCommand[i];
			}
			this.ArrPathCommand = aNewArray;
			this.path.ArrPathCommand = aNewArray;
		}

		this.path.startPos = ++this.curPos;
		this.path.curLen = 0;
		this.ArrPathCommand[this.curPos] = 0;
		return this.path;
	};
	CPathMemory.prototype.AllocPath2 = function () {
		this.path = new AscFormat.Path2(this);
		return this.AllocPath();
	};
	CPathMemory.prototype.GetPath = function (index) {
		this.path.startPos = index;
		this.path.curLen = 0;
		return this.path;
	};

	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetNvGrFrProps] = function (oClass, value) {
		oClass.nvGraphicFramePr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetThemeOverride] = function (oClass, value) {
		oClass.themeOverride = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ShapeSetBDeleted] = function (oClass, value) {
		oClass.bDeleted = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetParent] = function (oClass, value) {
		oClass.oldParent = oClass.parent;
		oClass.parent = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetChart] = function (oClass, value) {
		oClass.chart = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetClrMapOvr] = function (oClass, value) {
		oClass.clrMapOvr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetDate1904] = function (oClass, value) {
		oClass.date1904 = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetExternalData] = function (oClass, value) {
		oClass.externalData = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetLang] = function (oClass, value) {
		oClass.lang = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetPivotSource] = function (oClass, value) {
		oClass.pivotSource = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetPrintSettings] = function (oClass, value) {
		oClass.printSettings = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetProtection] = function (oClass, value) {
		oClass.protection = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetRoundedCorners] = function (oClass, value) {
		oClass.roundedCorners = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetStyle] = function (oClass, value) {
		oClass.style = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetTxPr] = function (oClass, value) {
		oClass.txPr = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_SetGroup] = function (oClass, value) {
		oClass.group = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_ChartStyle] = function (oClass, value) {
		oClass.chartStyle = value;
	};
	drawingsChangesMap[AscDFH.historyitem_ChartSpace_ChartColors] = function (oClass, value) {
		oClass.chartColors = value;
	};

	function CLabelsBox(aStrings, oAxis, oChartSpace) {
		this.x = 0.0;
		this.y = 0.0;
		this.extX = 0.0;
		this.extY = 0.0;
		this.aLabels = [];
		this.maxMinWidth = -1.0;

		this.bRotated = false;
		this.align = true;

		this.chartSpace = oChartSpace;
		this.axis = oAxis;
		this.count = 0;

		var oStyle = null, oLbl, fMinW;
		var oFirstTextPr = null;
		for (var i = 0; i < aStrings.length; ++i) {
			if (typeof aStrings[i] === "string") {
				oLbl = fCreateLabel(aStrings[i], i, oAxis, oChartSpace, oAxis.txPr, oAxis.spPr, oChartSpace.getDrawingDocument());
				if (oStyle) {
					oLbl.lastStyleObject = oStyle;
				}
				if (oFirstTextPr) {
					var aRuns = oLbl.tx.rich.content.Content[0] && oLbl.tx.rich.content.Content[0].Content;
					if (aRuns) {
						for (var j = 0; j < aRuns.length; ++j) {
							var oRun = aRuns[j];
							if (oRun.RecalcInfo && true === oRun.RecalcInfo.TextPr) {
								oRun.RecalcInfo.TextPr = false;
								oRun.CompiledPr = oFirstTextPr;
							}
						}
					}
				}
				fMinW = oLbl.tx.rich.content.RecalculateMinMaxContentWidth().Min;
				if (!oFirstTextPr) {
					if (oLbl.tx.rich.content.Content[0]) {
						oFirstTextPr = oLbl.tx.rich.content.Content[0].Get_FirstTextPr2();
					}
				}
				if (fMinW > this.maxMinWidth) {
					this.maxMinWidth = fMinW;
				}
				this.aLabels.push(oLbl);
				if (!oStyle) {
					oStyle = oLbl.lastStyleObject;
				}
				++this.count;
			} else {
				this.aLabels.push(null);
			}
		}

	}

	CLabelsBox.prototype.draw = function (graphics) {
		for (var i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i])
				this.aLabels[i].draw(graphics);
		}
		// graphics.p_width(70);
		// graphics.p_color(0, 0, 0, 255);
		// graphics._s();
		// graphics._m(this.x, this.y);
		// graphics._l(this.x + this.extX, this.y + 0);
		// graphics._l(this.x + this.extX, this.y + this.extY);
		// graphics._l(this.x + 0, this.y + this.extY);
		// graphics._z();
		// graphics.ds();
	};
	CLabelsBox.prototype.checkMaxMinWidth = function () {
		if (this.maxMinWidth < 0.0) {
			var oStyle = null, oLbl, fMinW;
			for (var i = 0; i < this.aLabels.length; ++i) {
				oLbl = this.aLabels[i];
				if (oLbl) {
					if (oStyle) {
						oLbl.lastStyleObject = oStyle;
					}
					fMinW = oLbl.tx.rich.content.RecalculateMinMaxContentWidth().Min;
					if (fMinW > this.maxMinWidth) {
						this.maxMinWidth = fMinW;
					}
					if (!oStyle) {
						oStyle = oLbl.lastStyleObject;
					}
				}
			}
		}
		return this.maxMinWidth >= 0.0 ? this.maxMinWidth : 0.0;
	};
	CLabelsBox.prototype.hit = function (x, y) {
		if(this.bRadarCat) {
			for(let nLbl = 0; nLbl < this.aLabels.length; ++nLbl) {
				let oLbl = this.aLabels[nLbl];
				if(oLbl && oLbl.hit(x, y)) {
					return true;
				}
			}
			return false;
		}
		var tx, ty;
		if (this.chartSpace && this.chartSpace.invertTransform) {
			tx = this.chartSpace.invertTransform.TransformPointX(x, y);
			ty = this.chartSpace.invertTransform.TransformPointY(x, y);
			return tx >= this.x && ty >= this.y && tx <= this.x + this.extX && ty <= this.y + this.extY;
		}
		return false;
	};
	CLabelsBox.prototype.updatePosition = function (x, y) {
		// this.posX = x;
		// this.posY = y;
		// this.transform = this.localTransform.CreateDublicate();
		// global_MatrixTransformer.TranslateAppend(this.transform, x, y);
		// this.invertTransform = global_MatrixTransformer.Invert(this.transform);
		for (var i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i])
				this.aLabels[i].updatePosition(x, y);
		}
	};
	CLabelsBox.prototype.layoutHorNormal = function (fAxisY, fDistance, fXStart, fInterval, bOnTickMark, fForceContentWidth) {
		this.bRotated = false;
		this.align = (fDistance >= 0);
		var fMaxHeight = 0.0;
		var fCurX = bOnTickMark ? fXStart - fInterval / 2.0 : fXStart;
		if (fInterval < 0.0) {
			fCurX += fInterval;
		}
		var oFirstLabel = null, fFirstLabelCenterX = null, oLastLabel = null, fLastLabelCenterX = null;
		var fContentWidth = fForceContentWidth ? fForceContentWidth : Math.abs(fInterval);
		var fHorShift = Math.abs(fInterval) / 2.0 - fContentWidth / 2.0;
		for (var i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i]) {
				var oLabel = this.aLabels[i];
				var oContent = oLabel.tx.rich.content;
				oContent.Reset(0, 0, fContentWidth, 20000.0);
				oContent.Recalculate_Page(0, true);
				var fCurHeight = oContent.GetSummaryHeight();
				if (fCurHeight > fMaxHeight) {
					fMaxHeight = fCurHeight;
				}
				var fX, fY;
				fX = fCurX + fHorShift;
				if (fDistance >= 0.0) {
					fY = fAxisY + fDistance;
				} else {
					fY = fAxisY + fDistance - fCurHeight;
				}
				var oTransform = oLabel.transformText;
				oTransform.Reset();
				global_MatrixTransformer.TranslateAppend(oTransform, fX, fY);
				oTransform = oLabel.localTransformText;
				oTransform.Reset();
				global_MatrixTransformer.TranslateAppend(oTransform, fX, fY);


				if (oFirstLabel === null) {
					oFirstLabel = oLabel;
					fFirstLabelCenterX = fCurX + Math.abs(fInterval) / 2.0;
				}
				oLastLabel = oLabel;
				fLastLabelCenterX = fCurX + Math.abs(fInterval) / 2.0;
			}
			fCurX += fInterval;
		}

		var x0, x1;
		if (bOnTickMark && oFirstLabel && oLastLabel) {
			var fFirstLabelContentWidth = oFirstLabel.tx.rich.getMaxContentWidth(fContentWidth);
			var fLastLabelContentWidth = oLastLabel.tx.rich.getMaxContentWidth(fContentWidth);
			x0 = Math.min(fFirstLabelCenterX - fFirstLabelContentWidth / 2.0,
				fLastLabelCenterX - fLastLabelContentWidth / 2.0, fXStart, fXStart + fInterval * (this.aLabels.length - 1));
			x1 = Math.max(fFirstLabelCenterX + fFirstLabelContentWidth / 2.0,
				fLastLabelCenterX + fLastLabelContentWidth / 2.0, fXStart, fXStart + fInterval * (this.aLabels.length - 1));
		} else {
			x0 = Math.min(fXStart, fXStart + fInterval * (this.aLabels.length));
			x1 = Math.max(fXStart, fXStart + fInterval * (this.aLabels.length));
		}
		this.x = x0;
		this.extX = x1 - x0;

		if (fDistance >= 0.0) {
			this.y = fAxisY;
			this.extY = fDistance + fMaxHeight;
		} else {
			this.y = fAxisY + fDistance - fMaxHeight;
			this.extY = fMaxHeight - fDistance;
		}
	};
	CLabelsBox.prototype.layoutHorRotated = function (fAxisY, fDistance, fXStart, fXEnd, fInterval, bOnTickMark) {
		this.bRotated = true;
		this.align = (fDistance >= 0);
		var bTickLblSkip = AscFormat.isRealNumber(this.axis.tickLblSkip);
		if (bTickLblSkip) {
			this.layoutHorRotated2(this.aLabels, fAxisY, fDistance, fXStart, fInterval, bOnTickMark);
		} else {

			var fAngle = Math.PI / 4.0, fMultiplier = Math.sin(fAngle);
			var aLabelsSource = [].concat(this.aLabels);
			var oLabel = aLabelsSource[0];
			var i = 1;
			while (!oLabel && i < aLabelsSource.length) {
				oLabel = aLabelsSource[i];
			}
			if (oLabel) {
				var oContent = oLabel.tx.rich.content;
				oContent.SetApplyToAll(true);
				oContent.SetParagraphAlign(AscCommon.align_Left);
				oContent.SetParagraphIndent({FirstLine: 0.0, Left: 0.0});
				oContent.SetApplyToAll(false);
				var oSize = oLabel.tx.rich.getContentOneStringSizes();
				var fInset = fMultiplier * (oSize.h);
				fInset *= 2;
				if (fInset <= fInterval) {
					this.layoutHorRotated2(this.aLabels, fAxisY, fDistance, fXStart, fInterval, bOnTickMark);
				} else {


					var nIntervalCount = bOnTickMark ? this.count - 1 : this.count;
					var fInterval_ = Math.abs(fXEnd - fXStart) / nIntervalCount;
					var nLblTickSkip = (fInset / fInterval_ + 0.5) >> 0;
					var aLabels = [].concat(aLabelsSource);
					var index = 0;
					if (nLblTickSkip > 1) {
						for (i = 0; i < aLabels.length; ++i) {
							if (aLabels[i]) {
								if ((index % nLblTickSkip) !== 0) {
									aLabels[i] = null;
								}
								index++;
							}
						}
					}
					this.layoutHorRotated2(aLabels, fAxisY, fDistance, fXStart, fInterval, bOnTickMark);
				}
			}
		}

	};
	CLabelsBox.prototype.layoutHorRotated2 = function (aLabels, fAxisY, fDistance, fXStart, fInterval, bOnTickMark) {
		this.bRotated = true;
		this.align = (fDistance >= 0);
		var i;
		var fMaxHeight = 0.0;
		var fCurX = bOnTickMark ? fXStart : fXStart + fInterval / 2.0;
		var fAngle = Math.PI / 4.0, fMultiplier = Math.sin(fAngle);
		var fMinLeft = null, fMaxRight = null;
		for (i = 0; i < aLabels.length; ++i) {
			if (aLabels[i]) {
				var oLabel = aLabels[i];
				var oContent = oLabel.tx.rich.content;
				oContent.SetApplyToAll(true);
				oContent.SetParagraphAlign(AscCommon.align_Left);
				oContent.SetParagraphIndent({FirstLine: 0.0, Left: 0.0});
				oContent.SetApplyToAll(false);
				var oSize = oLabel.tx.rich.getContentOneStringSizes();
				var fBoxW = fMultiplier * (oSize.w + oSize.h);
				var fBoxH = fBoxW;
				if (fBoxH > fMaxHeight) {
					fMaxHeight = fBoxH;
				}
				var fX1, fY0, fXC, fYC;
				fY0 = fAxisY + fDistance;
				if (fDistance >= 0.0) {
					fXC = fCurX - oSize.w * fMultiplier / 2.0;
					fYC = fY0 + fBoxH / 2.0;
				} else {
					//fX1 = fCurX - oSize.h*fMultiplier;
					fXC = fCurX + oSize.w * fMultiplier / 2.0;
					fYC = fY0 - fBoxH / 2.0;
				}
				var oTransform = oLabel.localTransformText;
				oTransform.Reset();
				global_MatrixTransformer.TranslateAppend(oTransform, -oSize.w / 2.0, -oSize.h / 2.0);
				global_MatrixTransformer.RotateRadAppend(oTransform, fAngle);
				global_MatrixTransformer.TranslateAppend(oTransform, fXC, fYC);
				oLabel.transformText = oTransform.CreateDublicate();
				if (null === fMinLeft || (fXC - fBoxW / 2.0) < fMinLeft) {
					fMinLeft = fXC - fBoxW / 2.0;
				}
				if (null === fMaxRight || (fXC + fBoxW / 2.0) > fMaxRight) {
					fMaxRight = fXC + fBoxW / 2.0;
				}
			}
			fCurX += fInterval;
		}
		this.aLabels = aLabels;
		var aPoints = [];
		aPoints.push(fXStart);
		var nIntervalCount = bOnTickMark ? aLabels.length - 1 : aLabels.length;
		aPoints.push(fXStart + fInterval * nIntervalCount);
		if (null !== fMinLeft) {
			aPoints.push(fMinLeft);
		}
		if (null !== fMaxRight) {
			aPoints.push(fMaxRight);
		}
		this.x = Math.min.apply(Math, aPoints);
		this.extX = Math.max.apply(Math, aPoints) - this.x;
		if (fDistance >= 0.0) {
			this.y = fAxisY;
			this.extY = fDistance + fMaxHeight;
		} else {
			this.y = fAxisY + fDistance - fMaxHeight;
			this.extY = fMaxHeight - fDistance;
		}
	};
	CLabelsBox.prototype.layoutVertNormal = function (fAxisX, fDistance, fYStart, fInterval, bOnTickMark, fMaxBlockWidth) {
		var fCurY = bOnTickMark ? fYStart : fYStart + fInterval / 2.0;
		this.bRotated = false;
		this.align = (fDistance < 0);
		var fDistance_ = Math.abs(fDistance);
		var oTransform, oContent, oLabel, fMinY = fYStart, fMaxY = fYStart + fInterval * (this.aLabels.length - 1), fY,
			i;
		var fMaxContentWidth = 0.0, oSize;
		var fLabelHeight = 0.0;
		var oCompiledPr = null;
		for (i = 0; i < this.aLabels.length; ++i) {
			oLabel = this.aLabels[i];
			if (oLabel) {
				if (oCompiledPr) {
					oLabel.tx.rich.content.Content[0].CompiledPr = oCompiledPr;
				}
				oSize = oLabel.tx.rich.getContentOneStringSizes();
				if (!oCompiledPr) {
					oCompiledPr = oLabel.tx.rich.content.Content[0].CompiledPr;
				}
				fLabelHeight = oSize.h;
				break;
			}
		}

		var nIntervalCount = bOnTickMark ? this.count - 1 : this.count;
		var fInterval_ = Math.abs(fMinY - fMaxY) / nIntervalCount;

		var nLblTickSkip = (fLabelHeight / fInterval_) >> 0, index = 0;
		if (nLblTickSkip > 1) {
			for (i = 0; i < this.aLabels.length; ++i) {
				if (this.aLabels[i]) {
					if ((index % nLblTickSkip) !== 0) {
						this.aLabels[i] = null;
					}
					index++;
				}
			}
		}
		for (i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i]) {
				oLabel = this.aLabels[i];
				oContent = oLabel.tx.rich.content;
				oContent.SetApplyToAll(true);
				oContent.SetParagraphAlign(AscCommon.align_Left);
				oContent.SetApplyToAll(false);
				if (oCompiledPr) {
					oLabel.tx.rich.content.Content[0].CompiledPr = oCompiledPr;
				}
				oSize = oLabel.tx.rich.getContentOneStringSizes();
				if (!oCompiledPr) {
					oCompiledPr = oLabel.tx.rich.content.Content[0].CompiledPr;
				}
				if (oSize.w + fDistance_ > fMaxBlockWidth) {
					break;
				}

				if (oSize.w > fMaxContentWidth) {
					fMaxContentWidth = oSize.w;
				}
				oTransform = oLabel.localTransformText;
				oTransform.Reset();
				fY = fCurY - oSize.h / 2.0;
				if (fDistance > 0.0) {
					global_MatrixTransformer.TranslateAppend(oTransform, fAxisX + fDistance, fY);
				} else {
					global_MatrixTransformer.TranslateAppend(oTransform, fAxisX + fDistance - oSize.w, fY);
				}
				oLabel.transformText = oTransform.CreateDublicate();
				if (fY < fMinY) {
					fMinY = fY;
				}
				if (fY + oSize.h > fMaxY) {
					fMaxY = fY + oSize.h;
				}
			}
			fCurY += fInterval;
		}
		oCompiledPr = null;
		if (i < this.aLabels.length) {
			var fMaxMinWidth = this.checkMaxMinWidth();
			fMaxContentWidth = 0.0;
			for (i = 0; i < this.aLabels.length; ++i) {
				oLabel = this.aLabels[i];
				if (oLabel) {
					oContent = oLabel.tx.rich.content;
					oContent.SetApplyToAll(true);
					oContent.SetParagraphAlign(AscCommon.align_Center);
					oContent.SetApplyToAll(false);
					var fContentWidth;
					if (fMaxMinWidth + fDistance_ < fMaxBlockWidth) {
						fContentWidth = oContent.RecalculateMinMaxContentWidth().Min + 0.1;
					} else {
						fContentWidth = fMaxBlockWidth - fDistance_;
					}
					if (fContentWidth > fMaxContentWidth) {
						fMaxContentWidth = fContentWidth;
					}
					if (oCompiledPr) {
						oContent.Content[0].CompiledPr = oCompiledPr;
					}
					oSize = oLabel.tx.rich.getContentOneStringSizes();
					if (!oCompiledPr) {
						oCompiledPr = oContent.Content[0].CompiledPr;
					}
					oContent.Reset(0, 0, fContentWidth, 20000);//выставляем большую ширину чтобы текст расчитался в одну строку.
					oContent.Recalculate_Page(0, true);
					var fContentHeight = oContent.GetSummaryHeight();
					oTransform = oLabel.localTransformText;
					oTransform.Reset();
					fY = fCurY - fContentHeight / 2.0;
					if (fDistance > 0.0) {
						global_MatrixTransformer.TranslateAppend(oTransform, fAxisX + fDistance, fY);
					} else {
						global_MatrixTransformer.TranslateAppend(oTransform, fAxisX + fDistance - fContentWidth, fY);
					}
					oLabel.transformText = oTransform.CreateDublicate();
					if (fY < fMinY) {
						fMinY = fY;
					}
					if (fY + fContentHeight > fMaxY) {
						fMaxY = fY + fContentHeight;
					}
				}
				fCurY += fInterval;
			}
		}

		if (fDistance > 0.0) {
			this.x = fAxisX;
		} else {
			this.x = fAxisX + fDistance - fMaxContentWidth;
		}
		this.extX = fMaxContentWidth + fDistance_;

		this.y = fMinY;
		this.extY = fMaxY - fMinY;
	};
	CLabelsBox.prototype.setPosition = function (x, y) {
		this.x = x;
		this.y = y;
		for (var i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i]) {
				var lbl = this.aLabels[i];
				lbl.setPosition(lbl.relPosX + x, lbl.relPosY + y);
			}
		}
	};
	CLabelsBox.prototype.checkShapeChildTransform = function (t) {
		// this.transform = this.localTransform.CreateDublicate();
		// global_MatrixTransformer.TranslateAppend(this.transform, this.posX, this.posY);
		// this.invertTransform = global_MatrixTransformer.Invert(this.transform);
		for (var i = 0; i < this.aLabels.length; ++i) {
			if (this.aLabels[i])
				this.aLabels[i].checkShapeChildTransform(t);
		}
	};
	CLabelsBox.prototype.getLabelsOffset = function() {
		let dStakeOffset = 1;
		if(this.axis) {
			if(AscFormat.isRealNumber(this.axis.lblOffset)) {
				dStakeOffset = this.axis.lblOffset / 100;
			}
		}
		let dFontSize = 11;
		for(let nLbl = 0; nLbl < this.aLabels.length; ++nLbl) {
			let oLbl = this.aLabels[nLbl];
			if(oLbl) {
				let oTextPr = oLbl.tx.rich.content.Content[0].Get_CompiledPr2(false).TextPr;
				AscCommon.g_oTextMeasurer.SetTextPr(oTextPr, oLbl.Get_Theme());
				AscCommon.g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, 1);
				let oInfo = AscCommon.g_oTextMeasurer.Measure2Code("A".charCodeAt(0));
				let dHeight = 0.8*oInfo.Height;
				return dHeight + (dHeight / 2 ) * dStakeOffset;
			}
		}
		return dFontSize * (25.4 / 72) * dStakeOffset;
	};
	CLabelsBox.prototype.drawSelect = function(drawingDocument, isDrawHandles) {
		if(this.bRadarCat) {
			for(let nLbl = 0; nLbl < this.aLabels.length; ++nLbl) {
				let oLbl = this.aLabels[nLbl];
				if(oLbl) {
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.chartSpace.transform, oLbl.x, oLbl.y, oLbl.extX, oLbl.extY, false, false, undefined, isDrawHandles);
				}
			}
		}
		else {
			drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.chartSpace.transform, this.x, this.y, this.extX, this.extY, false, false, undefined, isDrawHandles);
		}
	};

	function fCreateLabel(sText, idx, oParent, oChart, oTxPr, oSpPr, oDrawingDocument) {
		var dlbl = new AscFormat.CDLbl();
		dlbl.parent = oParent;
		dlbl.chart = oChart;
		dlbl.spPr = oSpPr;
		dlbl.txPr = oTxPr;
		dlbl.idx = idx;
		dlbl.tx = new AscFormat.CChartText();
		dlbl.tx.setChart(oChart);
		dlbl.tx.rich = AscFormat.CreateTextBodyFromString(sText, oDrawingDocument, dlbl);
		var content = dlbl.tx.rich.content;
		content.SetApplyToAll(true);
		content.SetParagraphAlign(AscCommon.align_Center);
		content.SetApplyToAll(false);
		dlbl.txBody = dlbl.tx.rich;
		dlbl.oneStringWidth = -1.0;
		return dlbl;
	}

	function fLayoutHorLabelsBox(oLabelsBox, fY, fXStart, fXEnd, bOnTickMark, fDistance, bForceVertical, bNumbers, fForceContentWidth) {
		let fAxisLength = fXEnd - fXStart;
		let nLabelsCount = oLabelsBox.aLabels.length;

		let bOnTickMark_ = bOnTickMark && nLabelsCount > 1;
		let nIntervalCount = bOnTickMark_ ? nLabelsCount - 1 : nLabelsCount;
		let fInterval = fAxisLength / nIntervalCount;
		let fForceContentWidth_ = fForceContentWidth;
		if (!bForceVertical || true) {//TODO: implement for vertical labels
			let fMaxMinWidth = oLabelsBox.checkMaxMinWidth();
			let fCheckInterval;
			if (AscFormat.isRealNumber(fForceContentWidth)) {
				fCheckInterval = fForceContentWidth;
			} else {
				fCheckInterval = Math.abs(fInterval);
				let nMinInterval = -1;
				let nLastIdx = -1;
				for (let nLbl = 0; nLbl < nLabelsCount; ++nLbl) {
					let oLbl = oLabelsBox.aLabels[nLbl];
					if (oLbl) {
						if (nLastIdx === -1) {
							nLastIdx = nLbl;
						} else {
							let nIdxInterval = nLbl - nLastIdx;
							if (nMinInterval === -1) {
								nMinInterval = nIdxInterval;
							} else {
								nMinInterval = Math.min(nMinInterval, nIdxInterval);
							}
							if (nMinInterval === 1) {
								break;
							}
						}
					}
				}
				if (nMinInterval > 0) {
					fCheckInterval *= nMinInterval;
					if (!fForceContentWidth_) {
						fForceContentWidth_ = fCheckInterval;
					} else {
						fForceContentWidth_ = Math.max(fCheckInterval, fForceContentWidth_);
					}
				}
			}
			if (fMaxMinWidth <= fCheckInterval) {
				oLabelsBox.layoutHorNormal(fY, fDistance, fXStart, fInterval, bOnTickMark_, fForceContentWidth_);
			} else {
				oLabelsBox.layoutHorRotated(fY, fDistance, fXStart, fXEnd, fInterval, bOnTickMark_);

			}
		}
	}

	function fLayoutVertLabelsBox(oLabelsBox, fX, fYStart, fYEnd, bOnTickMark, fDistance, bForceVertical) {
		var fAxisLength = fYEnd - fYStart;
		var nLabelsCount = oLabelsBox.aLabels.length;

		var bOnTickMark_ = bOnTickMark && nLabelsCount > 1;
		var nIntervalCount = bOnTickMark_ ? nLabelsCount - 1 : nLabelsCount;
		var fInterval = fAxisLength / nIntervalCount;

		if (!bForceVertical || true) {
			oLabelsBox.layoutVertNormal(fX, fDistance, fYStart, fInterval, bOnTickMark_);
		} else {
			//TODO: vertical text
		}
	}

	function fLayoutRadarCatLabelsBox(oLabelsBox, oRect, aAlphaPoints) {
		if(oLabelsBox && aAlphaPoints) {
			const nLblsLength = oLabelsBox.aLabels.length;
			const fAngleInterval = 2*Math.PI / nLblsLength;
			let dMinX = 10000, dMinY = 10000, dMaxX = -dMinX, dMaxY = -dMinY;

			let dCatLabelsHeight = 0;

			for(let nLabel = 0; nLabel < oLabelsBox.aLabels.length; ++nLabel) {
				var oLabel = oLabelsBox.aLabels[nLabel];
				if(oLabel) {
					let oSize = oLabel.tx.rich.getContentOneStringSizes();
					if(oSize.h > dCatLabelsHeight) {
						dCatLabelsHeight = oSize.h;
					}
				}
			}
			let dX = oRect.x;
			let dY = oRect.y;
			let extX = oRect.w;
			let extY = oRect.h;
			let dCenterX = dX + extX / 2;
			let dCenterY = dY + extY / 2;
			let dLength = Math.min(extX/2, extY/2)

			for(let nLblIdx = 0; nLblIdx < nLblsLength; ++nLblIdx) {
				let oLbl = oLabelsBox.aLabels[nLblIdx];
				let oAlphaPt = aAlphaPoints[nLblIdx];
				if(oLbl && oAlphaPt) {
					let oLblTxBody = oLbl.tx.rich;
					let oSize = oLblTxBody.getContentOneStringSizes();
					let fAngle = oAlphaPt.pos;
					let dQuad = fAngle / (Math.PI/2);
					let nQuad = dQuad >> 0;
					let oTransform = oLbl.localTransformText;
					oTransform.Reset();
					let dNewX = 0, dNewY = 0;
					if(AscFormat.fApproxEqual(dQuad, nQuad, 0.01)) {
						if(AscFormat.fApproxEqual(fAngle, 0, 0.5) || AscFormat.fApproxEqual(fAngle, 2*Math.PI, 0.5)) {
							dNewX = dCenterX - oSize.w / 2;
							dNewY = dCenterY - dLength - dCatLabelsHeight;
						}
						else if(AscFormat.fApproxEqual(fAngle, Math.PI / 2, 0.5)) {
							dNewX = dCenterX + dLength;
							dNewY = dCenterY - oSize.h / 2;
						}
						else if(AscFormat.fApproxEqual(fAngle, Math.PI, 0.5)) {
							dNewX = dCenterX - oSize.w/ 2;
							dNewY = dCenterY + dLength;
						}
						else if(AscFormat.fApproxEqual(fAngle, 3 * Math.PI / 2, 0.5)) {
							dNewX = dCenterX - dLength - oSize.w;
							dNewY = dCenterY - oSize.h / 2;
						}
					}
					else {
						if(nQuad === 0) {
							dNewX = dCenterX + Math.sin(fAngle) * dLength;
							dNewY = dCenterY - Math.cos(fAngle) * dLength - oSize.h;
						}
						else if(nQuad === 1) {
							dNewX = dCenterX + Math.sin(fAngle) * dLength;
							dNewY = dCenterY - Math.cos(fAngle) * dLength;
						}
						else if(nQuad === 2) {
							dNewX = dCenterX + Math.sin(fAngle) * dLength - oSize.w;
							dNewY = dCenterY - Math.cos(fAngle) * dLength;
						}
						else if(nQuad === 3) {
							dNewX = dCenterX + Math.sin(fAngle) * dLength - oSize.w;
							dNewY = dCenterY - Math.cos(fAngle) * dLength - oSize.h;
						}
					}
					let oLblTxContent = oLblTxBody.content;
					let oFirstPara = oLblTxContent.Content[0];
					dMinX = Math.min(dMinX, dNewX);
					dMinY = Math.min(dMinY, dNewY);
					dMaxX = Math.max(dMaxX, dNewX + oSize.w);
					dMaxY = Math.max(dMaxY, dNewY + oSize.h);
					oLbl.extX = oSize.w;
					oLbl.extY = oSize.h;
					oLbl.x = dNewX;
					oLbl.y = dNewY;
					global_MatrixTransformer.TranslateAppend(oLbl.localTransform, dNewX, dNewY);
					if(oFirstPara.CompiledPr.Pr.ParaPr.Jc === AscCommon.align_Center) {
						dNewX = dNewX + oSize.w/2 - oLblTxContent.XLimit/2;
					}
					global_MatrixTransformer.TranslateAppend(oTransform, dNewX, dNewY);
				}
			}
			oLabelsBox.x = dMinX;
			oLabelsBox.y = dMinY;
			oLabelsBox.extX = dMaxX - dMinX;
			oLabelsBox.extY = dMaxY - dMinY;
			oLabelsBox.bRadarCat = true;
		}
	}

	function CAxisGrid() {
		this.fStart = 0.0;
		this.fStride = 0.0;
		this.bOnTickMark = true;
		this.nCount = 0;
		this.minVal = 0.0;
		this.maxVal = 0.0;
		this.aStrings = [];
	}

	CAxisGrid.prototype.calculatePoints = function(aPoints, bOnTickMark, aScale) {
		if(!Array.isArray(aPoints)) {
			return;
		}
		let fStartSeriesPos = 0.0;
		if(!this.bOnTickMark) {
			fStartSeriesPos = this.fStride / 2.0;
		}
		for(let j = 0; j < this.aStrings.length; ++j) {
			aPoints.push({
				val: aScale[j],
				pos: this.fStart + j * this.fStride + fStartSeriesPos
			})
		}
	};

	function CChartSpace() {
		AscFormat.CGraphicObjectBase.call(this);
		this.nvGraphicFramePr = null;
		this.chart = null;
		this.clrMapOvr = null;
		this.date1904 = false;
		this.externalData = null;
		this.lang = null;
		this.pivotSource = null;
		this.printSettings = null;
		this.protection = null;
		this.roundedCorners = false;
		this.spPr = null;
		this.style = 2;
		this.txPr = null;
		this.themeOverride = null;
		this.userShapes = [];//userShapes
		this.chartStyle = null;
		this.chartColors = null;

		this.dataRefs = null;
		this.pathMemory = new CPathMemory();
		this.cachedCanvas = null;
		this.ptsCount = 0;
		this.isSparkline = false;
		this.selection =
			{
				title: null,
				legend: null,
				legendEntry: null,
				axisLbls: null,
				dataLbls: null,
				dataLbl: null,
				plotArea: null,
				rotatePlotArea: null,
				axis: null,
				minorGridlines: null,
				majorGridlines: null,
				gridLines: null,
				chart: null,
				series: null,
				datPoint: null,
				hiLowLines: null,
				upBars: null,
				downBars: null,
				markers: null,
				textSelection: null
			};
	}

	AscFormat.InitClass(CChartSpace, AscFormat.CGraphicObjectBase, AscDFH.historyitem_type_ChartSpace);

	CChartSpace.prototype.changeSize = CShape.prototype.changeSize;
	CChartSpace.prototype.getDataRefs = function () {
		if (!this.dataRefs) {
			this.dataRefs = new AscFormat.CChartDataRefs(this);
		}
		return this.dataRefs;
	};
	CChartSpace.prototype.clearDataRefs = function () {
		if (this.dataRefs) {
			this.dataRefs = null;
		}
	};
	CChartSpace.prototype.AllocPath = function () {
		return this.pathMemory.AllocPath().startPos;
	};
	CChartSpace.prototype.GetPath = function (index) {
		return this.pathMemory.GetPath(index);
	};
	CChartSpace.prototype.checkTypeCorrect = function () {
		if (!this.chart) {
			return false;
		}
		if (!this.chart.plotArea) {
			return false
		}
		if (this.chart.plotArea.charts.length === 0) {
			return false;
		}
		var allSeries = this.getAllSeries();
		if (allSeries.length === 0) {
			return false;
		}
		return true;
	};
	CChartSpace.prototype.getSortedChartsId = function () {
		if (this.chartObj) {
			return this.chartObj._sortChartsForDrawing(this);
		}
		return [];
	};
	CChartSpace.prototype._getSeriesArrayIdx = function (oChart, nSeriesIdx) {
		if (oChart.series[nSeriesIdx] && oChart.series[nSeriesIdx].idx === nSeriesIdx) {
			return nSeriesIdx;
		}
		for (var i = 0; i < oChart.series.length; ++i) {
			if (oChart.series[i].idx === nSeriesIdx) {
				return i;
			}
		}
		return -1;
	};
	CChartSpace.prototype._getPtArrayIdx = function (oChart, nSeriesIdx, nPtIdx) {
		var oSeries = oChart.series[this._getSeriesArrayIdx(nSeriesIdx)];
		if (oSeries) {
			var aPoints = oSeries.getNumPts();
			if (aPoints[nPtIdx] && aPoints[nPtIdx].idx === nPtIdx) {
				return nPtIdx;
			}
			for (var i = 0; i < aPoints.length; ++i) {
				if (aPoints[i].idx === nPtIdx) {
					return i;
				}
			}
		}
		return -1;
	};
	CChartSpace.prototype.getMarkerPaths = function (sChartId, nSeriesIdx, nPtIdx) {
		var oChartObj = this.chartObj;
		if (!oChartObj) {
			return [];
		}
		var oChart = AscCommon.g_oTableId.Get_ById(sChartId);
		if (!AscCommon.isRealObject(oChart)) {
			return [];
		}

		var oDrawChart = oChartObj.charts[sChartId];
		if (!AscCommon.isRealObject(oDrawChart)) {
			return [];
		}
		if (!AscCommon.isRealObject(oDrawChart.paths)) {
			return [];
		}
		if (!Array.isArray(oDrawChart.paths.points)) {
			return [];
		}
		var nArrayIdx = this._getSeriesArrayIdx(oChart, nSeriesIdx);
		if (!AscCommon.isRealObject(oDrawChart.paths.points[nArrayIdx])) {
			return [];
		}
		if (!Array.isArray(oDrawChart.paths.points[nArrayIdx])) {
			return [];
		}
		var aMarkers = oDrawChart.paths.points[nArrayIdx];
		var nPtIdx = this._getPtArrayIdx(oChart, nSeriesIdx, nPtIdx);
		if (!AscFormat.isRealNumber(aMarkers[nPtIdx])) {
			return [];
		}
		return [aMarkers[nPtIdx]];
	};
	CChartSpace.prototype.drawSelect = function (drawingDocument, nPageIndex) {
		var i, oPath;

		var oApi = Asc.editor || editor;
		var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
		if (this.selectStartPage === nPageIndex) {
			drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.SHAPE, this.getTransformMatrix(), 0, 0, this.extX, this.extY, false, this.canRotate(), undefined, isDrawHandles);
			if (window["NATIVE_EDITOR_ENJINE"]) {
				return;
			}
			if (this.selection.textSelection) {
				drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.selection.textSelection.transform, 0, 0, this.selection.textSelection.extX, this.selection.textSelection.extY, false, false, false, isDrawHandles);
			} else if (this.selection.title) {
				drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.selection.title.transform, 0, 0, this.selection.title.extX, this.selection.title.extY, false, false, false, isDrawHandles);
			} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
				var series = this.getAllSeries();
				var ser = series[this.selection.dataLbls];
				if (ser) {
					var pts = ser.getNumPts();
					if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
						for (i = 0; i < pts.length; ++i) {
							if (pts[i] && pts[i].compiledDlb && !pts[i].compiledDlb.bDelete) {
								drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, pts[i].compiledDlb.transform, 0, 0, pts[i].compiledDlb.extX, pts[i].compiledDlb.extY, false, false, undefined, isDrawHandles);
							}
						}
					} else {
						if (pts[this.selection.dataLbl] && pts[this.selection.dataLbl].compiledDlb) {
							drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, pts[this.selection.dataLbl].compiledDlb.transform, 0, 0, pts[this.selection.dataLbl].compiledDlb.extX, pts[this.selection.dataLbl].compiledDlb.extY, false, false, undefined, isDrawHandles);
						}
					}
				}

			} else if (this.selection.dataLbl) {
				drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.selection.dataLbl.transform, 0, 0, this.selection.dataLbl.extX, this.selection.dataLbl.extY, false, false, undefined, isDrawHandles);
			} else if (this.selection.legend) {
				if (AscFormat.isRealNumber(this.selection.legendEntry)) {
					var oEntry = this.chart.legend.findCalcEntryByIdx(this.selection.legendEntry);
					if (oEntry) {
						drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, oEntry.transformText, 0, 0, oEntry.contentWidth, oEntry.contentHeight, false, false, undefined, isDrawHandles);
					}
				} else {
					drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.SHAPE, this.selection.legend.transform, 0, 0, this.selection.legend.extX, this.selection.legend.extY, false, false, undefined, isDrawHandles);
				}
			} else if (this.selection.axisLbls) {
				var oLabels = this.selection.axisLbls.labels;
				oLabels.drawSelect(drawingDocument, isDrawHandles);
			} else if (this.selection.hiLowLines) {
				if (this.chartObj) {
					var oDrawChart = this.chartObj.charts[this.selection.chart];
					if (oDrawChart) {
						var seriesPaths = oDrawChart.paths.values;
						if (Array.isArray(seriesPaths)) {
							for (var i = 0; i < seriesPaths.length; ++i) {
								if (AscFormat.isRealNumber(seriesPaths[i].lowLines)) {
									oPath = this.GetPath(seriesPaths[i].lowLines);
									oPath.drawTracks(drawingDocument, this.transform);
								}
								if (AscFormat.isRealNumber(seriesPaths[i].highLines)) {
									oPath = this.GetPath(seriesPaths[i].highLines);
									oPath.drawTracks(drawingDocument, this.transform);
								}
							}
						}
					}
				}
			} else if (this.selection.upBars) {
				if (this.chartObj) {
					var oDrawChart = this.chartObj.charts[this.selection.chart];
					if (oDrawChart) {
						var seriesPaths = oDrawChart.paths.values;
						if (Array.isArray(seriesPaths)) {
							for (var i = 0; i < seriesPaths.length; ++i) {
								if (AscFormat.isRealNumber(seriesPaths[i].upBars)) {
									oPath = this.GetPath(seriesPaths[i].upBars);
									oPath.drawTracks(drawingDocument, this.transform);
								}
							}
						}
					}
				}
			} else if (this.selection.downBars) {
				if (this.chartObj) {
					var oDrawChart = this.chartObj.charts[this.selection.chart];
					if (oDrawChart) {
						var seriesPaths = oDrawChart.paths.values;
						if (Array.isArray(seriesPaths)) {
							for (var i = 0; i < seriesPaths.length; ++i) {
								if (AscFormat.isRealNumber(seriesPaths[i].downBars)) {
									oPath = this.GetPath(seriesPaths[i].downBars);
									oPath.drawTracks(drawingDocument, this.transform);
								}
							}
						}
					}
				}
			} else if (this.selection.plotArea) {

				var oChartSize = this.getChartSizes(true);
				drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.SHAPE, this.transform, oChartSize.startX, oChartSize.startY, oChartSize.w, oChartSize.h, false, false, undefined, isDrawHandles);
				/*if(!this.selection.rotatePlotArea)
                 {
                 drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, this.transform, oChartSize.startX, oChartSize.startY, oChartSize.w, oChartSize.h, false, false);
                 }
                 else
                 {
                 var arr = [
                 {x: oChartSize.startX, y: oChartSize.startY},
                 {x: oChartSize.startX + oChartSize.w/2, y: oChartSize.startY},
                 {x: oChartSize.startX + oChartSize.w, y: oChartSize.startY},
                 {x: oChartSize.startX + oChartSize.w, y: oChartSize.startY + oChartSize.h/2},
                 {x: oChartSize.startX + oChartSize.w, y: oChartSize.startY + oChartSize.h},
                 {x: oChartSize.startX + oChartSize.w/2, y: oChartSize.startY + oChartSize.h},
                 {x: oChartSize.startX, y: oChartSize.startY + oChartSize.h},
                 {x: oChartSize.startX, y: oChartSize.startY + oChartSize.h/2},
                 {x: oChartSize.startX, y: oChartSize.startY}
                 ];
                 drawingDocument.AutoShapesTrack.DrawEditWrapPointsPolygon(arr, this.transform);
                 }*/
			} else if (AscFormat.isRealNumber(this.selection.series)) {
				if (this.chartObj) {
					var oDrawChart = this.chartObj.charts[this.selection.chart];
					if (oDrawChart) {
						var seriesPaths = this.selection.markers ? oDrawChart.paths.points : oDrawChart.paths.series;
						var Paths;
						var b3dPie = oDrawChart.chart.getObjectType() === AscDFH.historyitem_type_PieChart;
						if (b3dPie) {
							Paths = seriesPaths;
						} else {
							if (Array.isArray(seriesPaths)) {
								Paths = seriesPaths[this.selection.series];
							}
						}

						if (Array.isArray(Paths)) {
							var aPointsPaths = Paths;
							if (AscFormat.isRealNumber(this.selection.datPoint)) {
								if (AscFormat.isRealNumber(aPointsPaths[this.selection.datPoint])) {
									oPath = this.GetPath(aPointsPaths[this.selection.datPoint]);
									oPath.drawTracks(drawingDocument, this.transform);
								} else if (Array.isArray(aPointsPaths[this.selection.datPoint])) {
									var aPointsPaths2 = aPointsPaths[this.selection.datPoint];
									for (var z = 0; z < aPointsPaths2.length; ++z) {
										if (AscCommon.isRealObject(aPointsPaths2[z])) {
											// downPath: 1230
											// frontPath: []
											// insidePath: 1188
											// upPath: 1213
											if (AscFormat.isRealNumber(aPointsPaths2[z].downPath) && !b3dPie) {
												oPath = this.GetPath(aPointsPaths2[z].downPath);
												oPath.drawTracks(drawingDocument, this.transform);
											}
											if (AscFormat.isRealNumber(aPointsPaths2[z].upPath)) {
												oPath = this.GetPath(aPointsPaths2[z].upPath);
												oPath.drawTracks(drawingDocument, this.transform);
											}

											var aFrontPaths = aPointsPaths2[z].frontPath || aPointsPaths2[z].frontPaths;
											if (Array.isArray(aFrontPaths) && !b3dPie) {
												for (var s = 0; s < aFrontPaths.length; ++s) {
													if (AscFormat.isRealNumber(aFrontPaths[s])) {
														oPath = this.GetPath(aFrontPaths[s]);
														oPath.drawTracks(drawingDocument, this.transform);
													}
												}
											}
											aFrontPaths = aPointsPaths2[z].darkPaths;
											if (Array.isArray(aFrontPaths)) {
												for (var s = 0; s < aFrontPaths.length; ++s) {
													if (AscFormat.isRealNumber(aFrontPaths[s])) {
														oPath = this.GetPath(aFrontPaths[s]);
														oPath.drawTracks(drawingDocument, this.transform);
													}
												}
											}

										} else if (AscFormat.isRealNumber(aPointsPaths2[z])) {
											oPath = this.GetPath(aPointsPaths2[z]);
											oPath.drawTracks(drawingDocument, this.transform);
										}
									}
								} else if (AscCommon.isRealObject(aPointsPaths[this.selection.datPoint])) {
									if (Array.isArray(aPointsPaths[this.selection.datPoint].frontPaths)) {
										for (var l = 0; l < aPointsPaths[this.selection.datPoint].frontPaths.length; ++l) {
											if (AscFormat.isRealNumber(aPointsPaths[this.selection.datPoint].frontPaths[l])) {
												oPath = this.GetPath(aPointsPaths[this.selection.datPoint].frontPaths[l]);
												oPath.drawTracks(drawingDocument, this.transform);
											}
										}
									}
									if (Array.isArray(aPointsPaths[this.selection.datPoint].darkPaths)) {
										for (var l = 0; l < aPointsPaths[this.selection.datPoint].darkPaths.length; ++l) {
											if (AscFormat.isRealNumber(aPointsPaths[this.selection.datPoint].darkPaths[l])) {
												oPath = this.GetPath(aPointsPaths[this.selection.datPoint].darkPaths[l]);
												oPath.drawTracks(drawingDocument, this.transform);
											}
										}
									}
									if (AscFormat.isRealNumber(aPointsPaths[this.selection.datPoint].path)) {
										oPath = this.GetPath(aPointsPaths[this.selection.datPoint].path);
										oPath.drawTracks(drawingDocument, this.transform);
									}
								}
							} else {
								for (var l = 0; l < aPointsPaths.length; ++l) {
									if (AscFormat.isRealNumber(aPointsPaths[l])) {
										oPath = this.GetPath(aPointsPaths[l]);
										oPath.drawTracks(drawingDocument, this.transform);
									} else if (Array.isArray(aPointsPaths[l])) {
										var aPointsPaths2 = aPointsPaths[l];
										for (var z = 0; z < aPointsPaths2.length; ++z) {
											if (AscFormat.isRealNumber(aPointsPaths2[z])) {
												oPath = this.GetPath(aPointsPaths2[z]);
												oPath.drawTracks(drawingDocument, this.transform);
											} else if (AscCommon.isRealObject(aPointsPaths2[z])) {
												// downPath: 1230
												// frontPath: []
												// insidePath: 1188
												// upPath: 1213
												if (AscFormat.isRealNumber(aPointsPaths2[z].downPath) && !b3dPie) {
													oPath = this.GetPath(aPointsPaths2[z].downPath);
													oPath.drawTracks(drawingDocument, this.transform);
												}
												if (AscFormat.isRealNumber(aPointsPaths2[z].upPath)) {
													oPath = this.GetPath(aPointsPaths2[z].upPath);
													oPath.drawTracks(drawingDocument, this.transform);
												}

												var aFrontPaths = aPointsPaths2[z].frontPath || aPointsPaths2[z].frontPaths;
												if (Array.isArray(aFrontPaths) && !b3dPie) {
													for (var s = 0; s < aFrontPaths.length; ++s) {
														if (AscFormat.isRealNumber(aFrontPaths[s])) {
															oPath = this.GetPath(aFrontPaths[s]);
															oPath.drawTracks(drawingDocument, this.transform);
														}
													}
												}
												aFrontPaths = aPointsPaths2[z].darkPaths;
												if (Array.isArray(aFrontPaths)) {
													for (var s = 0; s < aFrontPaths.length; ++s) {
														if (AscFormat.isRealNumber(aFrontPaths[s])) {
															oPath = this.GetPath(aFrontPaths[s]);
															oPath.drawTracks(drawingDocument, this.transform);
														}
													}
												}

											}

										}
									} else {
										if (AscCommon.isRealObject(aPointsPaths[l])) {
											if (Array.isArray(aPointsPaths[l].frontPaths)) {
												for (var p = 0; p < aPointsPaths[l].frontPaths.length; ++p) {
													if (AscFormat.isRealNumber(aPointsPaths[l].frontPaths[p])) {
														oPath = this.GetPath(aPointsPaths[l].frontPaths[p]);
														oPath.drawTracks(drawingDocument, this.transform);
													}
												}
											}
											if (Array.isArray(aPointsPaths[l].darkPaths)) {
												for (var p = 0; p < aPointsPaths[l].darkPaths.length; ++p) {
													if (AscFormat.isRealNumber(aPointsPaths[l].darkPaths[p])) {
														oPath = this.GetPath(aPointsPaths[l].darkPaths[p]);
														oPath.drawTracks(drawingDocument, this.transform);
													}
												}
											}
											if (AscFormat.isRealNumber(aPointsPaths[l].path)) {
												oPath = this.GetPath(aPointsPaths[l].path);
												oPath.drawTracks(drawingDocument, this.transform);
											}
										}
									}
								}
							}
						} else {
							if (AscFormat.isRealNumber(Paths)) {
								if (AscFormat.isRealNumber(this.selection.datPoint)) {
									Paths = seriesPaths[this.selection.datPoint];
								}
								oPath = this.GetPath(Paths);
								oPath.drawTracks(drawingDocument, this.transform);
							}
						}

					}
				}
			} else if (this.selection.axis) {
				var oAxObj;
				var oChartObj = this.chartObj;
				if (Array.isArray(oChartObj.axesChart)) {
					for (i = 0; i < oChartObj.axesChart.length; ++i) {
						if (oChartObj.axesChart[i] && oChartObj.axesChart[i].axis === this.selection.axis) {
							oAxObj = oChartObj.axesChart[i];
							break;
						}
					}
				}
				if (this.selection.majorGridlines) {
					if (oAxObj && oAxObj.paths && AscFormat.isRealNumber(oAxObj.paths.gridLines)) {
						oPath = this.GetPath(oAxObj.paths.gridLines);
						oPath.drawTracks(drawingDocument, this.transform);
					}
				} else if (this.selection.minorGridlines) {
					if (oAxObj && oAxObj.paths && AscFormat.isRealNumber(oAxObj.paths.minorGridLines)) {
						oPath = this.GetPath(oAxObj.paths.minorGridLines);
						oPath.drawTracks(drawingDocument, this.transform);
					}
				}
			}
		}
	};
	CChartSpace.prototype.recalculateTextPr = function () {
		if (this.txPr && this.txPr.content) {
			this.txPr.content.Reset(0, 0, 10, 10);
			this.txPr.content.Recalculate_Page(0, true);
		}
	};
	CChartSpace.prototype.getSelectionState = function () {
		var content_selection = null, content = null;
		if (this.selection.textSelection) {
			content = this.selection.textSelection.getDocContent();
			if (content)
				content_selection = content.GetSelectionState();
		}
		return {
			content: content,
			title: this.selection.title,
			legend: this.selection.legend,
			legendEntry: this.selection.legendEntry,
			axisLbls: this.selection.axisLbls,
			dataLbls: this.selection.dataLbls,
			dataLbl: this.selection.dataLbl,
			textSelection: this.selection.textSelection,
			plotArea: this.selection.plotArea,
			rotatePlotArea: this.selection.rotatePlotArea,
			contentSelection: content_selection,
			recalcTitle: this.recalcInfo.recalcTitle,
			bRecalculatedTitle: this.recalcInfo.bRecalculatedTitle,
			axis: this.selection.axis,
			minorGridlines: this.selection.minorGridlines,
			majorGridlines: this.selection.majorGridlines,
			gridLines: this.selection.gridLines,
			chart: this.selection.chart,
			series: this.selection.series,
			datPoint: this.selection.datPoint,
			markers: this.selection.markers,
			hiLowLines: this.selection.hiLowLines,
			upBars: this.selection.upBars,
			downBars: this.selection.downBars

		}
	};
	CChartSpace.prototype.setSelectionState = function (state) {
		this.selection.title = state.title;
		this.selection.legend = state.legend;
		this.selection.legendEntry = state.legendEntry;
		this.selection.axisLbls = state.axisLbls;
		this.selection.dataLbls = state.dataLbls;
		this.selection.dataLbl = state.dataLbl;
		this.selection.rotatePlotArea = state.rotatePlotArea;
		this.selection.textSelection = state.textSelection;
		this.selection.plotArea = state.plotArea;
		this.selection.axis = state.axis;
		this.selection.minorGridlines = state.minorGridlines;
		this.selection.majorGridlines = state.majorGridlines;
		this.selection.gridLines = state.gridLines;
		this.selection.chart = state.chart;
		this.selection.datPoint = state.datPoint;
		this.selection.markers = state.markers;
		this.selection.series = state.series;
		this.selection.hiLowLines = state.hiLowLines;
		this.selection.downBars = state.downBars;
		this.selection.upBars = state.upBars;
		if (isRealObject(state.recalcTitle)) {
			this.recalcInfo.recalcTitle = state.recalcTitle;
			this.recalcInfo.bRecalculatedTitle = state.bRecalculatedTitle;
		}
		if (state.contentSelection) {
			if (this.selection.textSelection) {
				var content = this.selection.textSelection.getDocContent();
				if (content)
					content.SetSelectionState(state.contentSelection, state.contentSelection.length - 1);
			}
		}
	};
	CChartSpace.prototype.loadDocumentStateAfterLoadChanges = function (state) {
		this.selection.title = null;
		this.selection.legend = null;
		this.selection.legendEntry = null;
		this.selection.axisLbls = null;
		this.selection.dataLbls = null;
		this.selection.dataLbl = null;
		this.selection.plotArea = null;
		this.selection.rotatePlotArea = null;
		this.selection.gridLine = null;
		this.selection.series = null;
		this.selection.datPoint = null;
		this.selection.markers = null;
		this.selection.textSelection = null;
		this.selection.axis = null;
		this.selection.minorGridlines = null;
		this.selection.majorGridlines = null;
		this.selection.hiLowLines = null;
		this.selection.upBars = null;
		this.selection.downBars = null;
		var bRet = false, aTitles, i;
		if (state.DrawingsSelectionState) {
			var chartSelection = state.DrawingsSelectionState.chartSelection;
			if (chartSelection) {
				if (this.chart) {
					if (chartSelection.title) {
						aTitles = this.getAllTitles();
						for (i = 0; i < aTitles.length; ++i) {
							if (aTitles[i] === chartSelection.title) {
								this.selection.title = aTitles[i];
								bRet = true;
								break;
							}
						}
						if (this.selection.title) {
							if (this.selection.title === chartSelection.textSelection) {
								var oTitleContent = this.selection.title.getDocContent();
								if (oTitleContent && oTitleContent === chartSelection.content) {
									this.selection.textSelection = this.selection.title;
									if (true === state.DrawingSelection) {
										oTitleContent.SetContentPosition(state.StartPos, 0, 0);
										oTitleContent.SetContentSelection(state.StartPos, state.EndPos, 0, 0, 0);
									} else {
										oTitleContent.SetContentPosition(state.Pos, 0, 0);
									}
								}
							}
						}
					}
				}
			}
		}
		return bRet;
	};
	CChartSpace.prototype.resetInternalSelection = function (noResetContentSelect) {
		if (this.selection.title) {
			this.selection.title.selected = false;
		}
		if (this.selection.textSelection) {
			if (!(noResetContentSelect === true)) {
				var content = this.selection.textSelection.getDocContent();
				content && content.RemoveSelection();
			}
			this.selection.textSelection = null;
		}
	};
	CChartSpace.prototype.getDocContent = function () {
		return null;
	};
	CChartSpace.prototype.resetSelection = function (noResetContentSelect) {
		this.resetInternalSelection(noResetContentSelect);
		this.selection.title = null;
		this.selection.legend = null;
		this.selection.legendEntry = null;
		this.selection.axisLbls = null;
		this.selection.dataLbls = null;
		this.selection.dataLbl = null;
		this.selection.textSelection = null;
		this.selection.plotArea = null;
		this.selection.rotatePlotArea = null;
		this.selection.series = null;
		this.selection.chart = null;
		this.selection.datPoint = null;
		this.selection.markers = null;
		this.selection.axis = null;
		this.selection.minorGridlines = null;
		this.selection.majorGridlines = null;
		this.selection.hiLowLines = null;
		this.selection.upBars = null;
		this.selection.downBars = null;
	};
	CChartSpace.prototype.getStyles = function () {
		return AscFormat.ExecuteNoHistory(function () {

			var styles = new CStyles(false);
			var style = new CStyle("dataLblStyle", null, null, null, true);
			var text_pr = new CTextPr();
			text_pr.FontSize = 10;
			text_pr.Unifill = CreateUnfilFromRGB(0, 0, 0);
			var parent_objects = this.getParentObjects();
			var theme = parent_objects.theme;

			var para_pr = new CParaPr();
			para_pr.Jc = AscCommon.align_Center;
			para_pr.Spacing.Before = 0.0;
			para_pr.Spacing.After = 0.0;
			para_pr.Spacing.Line = 1;
			para_pr.Spacing.LineRule = Asc.linerule_Auto;
			style.ParaPr = para_pr;

			var minor_font = theme.themeElements.fontScheme.minorFont;


			if (minor_font) {
				if (typeof minor_font.latin === "string" && minor_font.latin.length > 0) {
					text_pr.RFonts.Ascii = {Name: minor_font.latin, Index: -1};
				}
				if (typeof minor_font.ea === "string" && minor_font.ea.length > 0) {
					text_pr.RFonts.EastAsia = {Name: minor_font.ea, Index: -1};
				}
				if (typeof minor_font.cs === "string" && minor_font.cs.length > 0) {
					text_pr.RFonts.CS = {Name: minor_font.cs, Index: -1};
				}

				if (typeof minor_font.sym === "string" && minor_font.sym.length > 0) {
					text_pr.RFonts.HAnsi = {Name: minor_font.sym, Index: -1};
				}
			}
			style.TextPr = text_pr;

			var chart_text_pr;

			if (this.txPr
				&& this.txPr.content
				&& this.txPr.content.Content[0]
				&& this.txPr.content.Content[0].Pr) {
				style.ParaPr.Merge(this.txPr.content.Content[0].Pr);
				if (this.txPr.content.Content[0].Pr.DefaultRunPr) {
					chart_text_pr = this.txPr.content.Content[0].Pr.DefaultRunPr;
					style.TextPr.Merge(chart_text_pr);
				}

			}

			if (this.txPr
				&& this.txPr.content
				&& this.txPr.content.Content[0]
				&& this.txPr.content.Content[0].Pr) {
				style.ParaPr.Merge(this.txPr.content.Content[0].Pr);
				if (this.txPr.content.Content[0].Pr.DefaultRunPr)
					style.TextPr.Merge(this.txPr.content.Content[0].Pr.DefaultRunPr);
			}
			styles.Add(style);
			return {lastId: style.Id, styles: styles};
		}, this, []);
	};
	CChartSpace.prototype.getParagraphTextPr = function () {
		if (this.selection.title && !this.selection.textSelection) {
			return GetTextPrFormArrObjects([this.selection.title]);
		} else if (this.selection.legend) {
			if (!AscFormat.isRealNumber(this.selection.legendEntry)) {
				if (AscFormat.isRealNumber(this.legendLength)) {
					var arrForProps = [];
					for (var i = 0; i < this.legendLength; ++i) {
						arrForProps.push(this.chart.legend.getCalcEntryByIdx(i, this.getDrawingDocument()))
					}
					return GetTextPrFormArrObjects(arrForProps);
				}
				return GetTextPrFormArrObjects(this.chart.legend.calcEntryes);
			} else {
				var calcLegendEntry = this.chart.legend.getCalcEntryByIdx(this.selection.legendEntry, this.getDrawingDocument());
				if (calcLegendEntry) {
					return GetTextPrFormArrObjects([calcLegendEntry]);
				}
			}
		} else if (this.selection.textSelection) {
			return this.selection.textSelection.txBody.content.GetCalculatedTextPr();
		} else if (this.selection.axisLbls && this.selection.axisLbls.labels) {
			return GetTextPrFormArrObjects(this.selection.axisLbls.labels.aLabels, true);
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var pts = ser.getNumPts();
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					return GetTextPrFormArrObjects(pts, undefined, true);
				} else {
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						return GetTextPrFormArrObjects([pt], undefined, true);
					}
				}
			}
		}
		var text_pr = new CTextPr();
		text_pr.InitDefault();
		text_pr.FontSize = 10;
		if (AscFormat.isRealNumber(this.style)) {
			if (this.style > 40) {
				text_pr.Unifill = AscFormat.CreateUnfilFromRGB(255, 255, 255);
			} else {
				var default_style = AscFormat.CHART_STYLE_MANAGER.getDefaultLineStyleByIndex(this.style);
				var oUnifill = default_style.axisAndMajorGridLines.createDuplicate();
				if (oUnifill && oUnifill.fill && oUnifill.fill.color && oUnifill.fill.color.Mods) {
					oUnifill.fill.color.Mods.Mods.length = 0;
				}
				text_pr.Unifill = oUnifill;
			}
		} else {
			text_pr.Unifill = AscFormat.CreateUnfilFromRGB(0, 0, 0);
		}
		text_pr.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);
		if (this.txPr && this.txPr.content && this.txPr.content.Content[0] && this.txPr.content.Content[0].Pr.DefaultRunPr) {
			text_pr.Merge(this.txPr.content.Content[0].Pr.DefaultRunPr.Copy());
		}
		if (text_pr.RFonts && text_pr.RFonts.Ascii && text_pr.RFonts.Ascii.Name) {
			text_pr.FontFamily.Name = text_pr.RFonts.Ascii.Name;
		}
		return text_pr;
	};
	CChartSpace.prototype.applyLabelsFunction = function (fCallback, value) {
		if (this.selection.title) {
			var DefaultFontSize = 18;
			if (this.selection.title !== this.chart.title) {
				DefaultFontSize = 10;
			}
			fCallback(this.selection.title, value, this.getDrawingDocument(), DefaultFontSize);
		} else if (this.selection.legend) {
			if (!AscFormat.isRealNumber(this.selection.legendEntry)) {
				fCallback(this.selection.legend, value, this.getDrawingDocument(), 10);
				for (var i = 0; i < this.selection.legend.legendEntryes.length; ++i) {
					fCallback(this.selection.legend.legendEntryes[i], value, this.getDrawingDocument(), 10);
				}
			} else {
				var entry = this.selection.legend.findLegendEntryByIndex(this.selection.legendEntry);
				if (!entry) {
					entry = new AscFormat.CLegendEntry();
					entry.setIdx(this.selection.legendEntry);
					if (this.selection.legend.txPr) {
						entry.setTxPr(this.selection.legend.txPr.createDuplicate());
					}

					this.selection.legend.addLegendEntry(entry);
				}
				if (entry) {
					fCallback(entry, value, this.getDrawingDocument(), 10);
				}
			}
		} else if (this.selection.axisLbls) {
			fCallback(this.selection.axisLbls, value, this.getDrawingDocument(), 10);
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var pts = ser.getNumPts();
				if (!ser.dLbls) {
					var oDlbls;
					var oChart = ser.parent;
					if (oChart && oChart.dLbls) {
						oDlbls = oChart.dLbls.createDuplicate();
					} else {
						oDlbls = new AscFormat.CDLbls();
					}
					ser.setDLbls(oDlbls);
				}
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					fCallback(ser.dLbls, value, this.getDrawingDocument(), 10);
					for (var i = 0; i < pts.length; ++i) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pts[i].idx);
						if (dLbl) {
							if (ser.dLbls.txPr && !dLbl.txPr) {
								dLbl.setTxPr(ser.dLbls.txPr.createDuplicate());
							}
							fCallback(dLbl, value, this.getDrawingDocument(), 10);
						}
					}
				} else {
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pt.idx);
						if (!dLbl) {
							dLbl = new AscFormat.CDLbl();
							dLbl.setIdx(pt.idx);
							if (ser.dLbls.txPr) {
								dLbl.merge(ser.dLbls);
							}
							ser.dLbls.addDLbl(dLbl);

						}
						fCallback(dLbl, value, this.getDrawingDocument(), 10);
					}
				}
			}
		} else {
			var oDD = this.getDrawingDocument();
			fCallback(this, value, oDD, 10);
			var chart = this.chart, i;
			if (chart) {
				for (i = 0; i < chart.pivotFmts.length; ++i) {
					chart.pivotFmts[i] && fCallback(chart.pivotFmts[i], value, oDD, 10);
				}
				if (chart.legend) {
					fCallback(chart.legend, value, oDD, 10);
					for (i = 0; i < chart.legend.legendEntryes.length; ++i) {
						chart.legend.legendEntryes[i] && fCallback(chart.legend.legendEntryes[i], value, oDD, 10);
					}
				}
				if (chart.title) {
					fCallback(chart.title, value, oDD, 18);
				}
				var plot_area = chart.plotArea;
				if (plot_area) {
					for (i = 0; i < plot_area.charts.length; ++i) {
						plot_area.charts[i] && plot_area.charts[i].applyLabelsFunction(fCallback, value, oDD);
						/*TODO надо бы этот метод переименовать чтоб название не вводило в заблуждение*/
					}
					var cur_axis;
					for (i = 0; i < plot_area.axId.length; ++i) {
						cur_axis = plot_area.axId[i];
						fCallback(cur_axis, value, oDD, 10);
						if (cur_axis.title) {
							fCallback(cur_axis.title, value, oDD, 10);
						}
					}
				}
			}
		}
	};
	CChartSpace.prototype.paragraphIncDecFontSize = function (bIncrease) {
		if (this.selection.textSelection) {
			this.selection.textSelection.checkDocContent();
			var content = this.selection.textSelection.getDocContent();
			if (content) {
				content.IncreaseDecreaseFontSize(bIncrease);
			}
			return;
		}
		this.applyLabelsFunction(CheckIncDecFontSize, bIncrease);
	};
	CChartSpace.prototype.paragraphAdd = function (paraItem, bRecalculate) {
		if (paraItem.Type === para_TextPr) {
			var _paraItem;
			if (paraItem.Value.Unifill && paraItem.Value.Unifill.checkWordMods()) {
				_paraItem = paraItem.Copy();
				_paraItem.Value.Unifill.convertToPPTXMods();
			} else {
				_paraItem = paraItem;
			}
			if (this.selection.textSelection) {
				this.selection.textSelection.checkDocContent();
				this.selection.textSelection.paragraphAdd(paraItem, bRecalculate);
				return;
			}
			/*if(this.selection.title)
             {
             this.selection.title.checkDocContent();
             CheckObjectTextPr(this.selection.title, _paraItem.Value, this.getDrawingDocument(), 18);
             if(this.selection.title.tx && this.selection.title.tx.rich && this.selection.title.tx.rich.content)
             {
             this.selection.title.tx.rich.content.SetApplyToAll(true);
             this.selection.title.tx.rich.content.AddToParagraph(_paraItem);
             this.selection.title.tx.rich.content.SetApplyToAll(false);
             }
             return;
             }*/
			this.applyLabelsFunction(CheckObjectTextPr, _paraItem.Value);
		} else if (paraItem.IsText() || paraItem.IsSpace()) {
			if (this.selection.title) {
				this.selection.textSelection = this.selection.title;
				this.selection.textSelection.checkDocContent();
				this.selection.textSelection.txBody.content.SelectAll();
				this.selection.textSelection.paragraphAdd(paraItem, bRecalculate);
			} else if (AscFormat.isRealNumber(this.selection.dataLbl)) {
				var oDlbl = this.getCompiledDlblBySelect();
				if (oDlbl) {
					this.selection.textSelection = oDlbl;
					this.selection.textSelection.checkDocContent();
					this.selection.textSelection.txBody.content.SelectAll();
					this.selection.textSelection.applyTextFunction(CDocumentContent.prototype.AddToParagraph, CTable.prototype.AddToParagraph, [paraItem, bRecalculate]);
				}
			}
		}
	};
	CChartSpace.prototype.applyTextFunction = function (docContentFunction, tableFunction, args) {
		if (docContentFunction === CDocumentContent.prototype.AddToParagraph && !this.selection.textSelection) {
			this.paragraphAdd(args[0], args[1]);
			return;
		}
		if (this.selection.textSelection) {
			this.selection.textSelection.checkDocContent();

			var bTrackRevision = false;
			if (typeof editor !== "undefined" && editor && editor.WordControl && editor.WordControl.m_oLogicDocument.IsTrackRevisions()) {
				bTrackRevision = editor.WordControl.m_oLogicDocument.GetLocalTrackRevisions();
				editor.WordControl.m_oLogicDocument.SetLocalTrackRevisions(false);
			}
			this.selection.textSelection.applyTextFunction(docContentFunction, tableFunction, args);

			if (false !== bTrackRevision) {
				editor.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTrackRevision);
			}
		}
	};
	CChartSpace.prototype.selectTitle = function (title, pageIndex) {
		title.select(this, pageIndex);
		this.resetInternalSelection();
		this.selection.legend = null;
		this.selection.legendEntry = null;
		this.selection.axisLbls = null;
		this.selection.dataLbls = null;
		this.selection.dataLbl = null;
		this.selection.textSelection = null;
		this.selection.plotArea = null;
		this.selection.rotatePlotArea = null;
		this.selection.axis = null;
		this.selection.minorGridlines = null;
		this.selection.majorGridlines = null,
			this.selection.gridLines = null;
		this.selection.chart = null;
		this.selection.series = null;
		this.selection.datPoint = null;
		this.selection.markers = null;
		this.selection.hiLowLines = null;
		this.selection.upBars = null;
		this.selection.downBars = null;

		this.selection.hiLowLines = null;
		this.selection.upBars = null;
		this.selection.downBars = null;
	};
	CChartSpace.prototype.recalculateCurPos = AscFormat.DrawingObjectsController.prototype.recalculateCurPos;
	CChartSpace.prototype.documentUpdateSelectionState = function () {
		if (this.selection.textSelection) {
			this.selection.textSelection.updateSelectionState();
		}
	};
	CChartSpace.prototype.getLegend = function () {
		if (this.chart) {
			return this.chart.legend;
		}
		return null;
	};
	CChartSpace.prototype.getChartTitle = function () {
		if (this.chart && this.chart.title) {
			return this.chart.title;
		}
		return null;
	};
	CChartSpace.prototype.getAllTitles = function () {
		var ret = [];
		if (this.chart) {
			if (this.chart.title) {
				ret.push(this.chart.title);
			}
			if (this.chart.plotArea) {
				var aAxes = this.chart.plotArea.axId;
				for (var i = 0; i < aAxes.length; ++i) {
					if (aAxes[i] && aAxes[i].title) {
						ret.push(aAxes[i].title);
					}
				}
			}
		}
		return ret;
	};
	CChartSpace.prototype.getSelectedChart = function () {
		return AscCommon.g_oTableId.Get_ById(this.selection.chart);
	};
	CChartSpace.prototype.getSelectedSeries = function () {
		if (AscFormat.isRealNumber(this.selection.series)) {
			var oTypedChart = this.getSelectedChart();
			if (oTypedChart) {
				//TODO: Rework this. Take the series only by index.
				var aSeries, oSeries;
				if (oTypedChart.getObjectType() === AscDFH.historyitem_type_BarChart) {
					aSeries = this.getAllSeries();
					for (var nIdx = 0; nIdx < aSeries.length; ++nIdx) {
						oSeries = aSeries[nIdx];
						if (oSeries.idx === this.selection.series) {
							return oSeries;
						}
					}
				} else {
					aSeries = oTypedChart.series;
					oSeries = aSeries[this.selection.series];
					if (oSeries) {
						return oSeries;
					}
				}
			}
		}
		return null;
	};
	CChartSpace.prototype.getCompiledDlblBySelect = function () {
		var ser = this.getAllSeries()[this.selection.dataLbls];
		if (ser) {
			var oDlbls = ser.dLbls;
			if (AscFormat.isRealNumber(this.selection.dataLbl)) {

				var pts = ser.getNumPts();
				var pt = pts[this.selection.dataLbl];
				if (pt) {
					return pt.compiledDlb;
				}
			}
		}
		return null;
	};
	CChartSpace.prototype.getFill = function () {
		var ret = null;

		var oChart = this.getSelectedChart();
		if (this.selection.plotArea) {
			if (this.chart.plotArea.brush && this.chart.plotArea.brush.fill) {
				ret = this.chart.plotArea.brush;
			}
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var oDlbls = ser.dLbls;
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					if (oDlbls && oDlbls.spPr && oDlbls.spPr.Fill && oDlbls.spPr.Fill.fill) {
						ret = oDlbls.spPr.Fill;
					}
				} else {
					var pts = ser.getNumPts();
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pt.idx);
						if (dLbl && dLbl.spPr && dLbl.spPr.Fill && dLbl.spPr.Fill.fill) {
							ret = dLbl.spPr.Fill;
						}
					}
				}
			}
		} else if (this.selection.axisLbls) {
			if (this.selection.axisLbls.spPr && this.selection.axisLbls.spPr.Fill && this.selection.axisLbls.spPr.Fill.fill) {
				ret = this.selection.axisLbls.spPr.Fill;
			}
		} else if (this.selection.title) {
			if (this.selection.title.spPr && this.selection.title.spPr.Fill && this.selection.title.spPr.Fill.fill) {
				ret = this.selection.title.spPr.Fill;
			}
		} else if (this.selection.legend) {
			if (this.selection.legend.spPr && this.selection.legend.spPr.Fill && this.selection.legend.spPr.Fill.fill) {
				ret = this.selection.legend.spPr.Fill;
			}
		} else if (AscFormat.isRealNumber(this.selection.series)) {
			var oSeries = this.getSelectedSeries();
			if (oSeries) {
				if (AscFormat.isRealNumber(this.selection.datPoint)) {
					var pts = oSeries.getNumPts();
					var datPoint = this.selection.datPoint;
					if (oSeries.getObjectType() === AscDFH.historyitem_type_LineSeries ||
						oSeries.getObjectType() === AscDFH.historyitem_type_ScatterSer ||
						oSeries.getObjectType() === AscDFH.historyitem_type_RadarSeries) {
						if (!this.selection.markers) {
							datPoint++;
						}
					}
					for (var j = 0; j < pts.length; ++j) {
						var pt = pts[j];
						if (pt.idx === datPoint) {
							if (this.selection.markers) {
								if (pt.compiledMarker && pt.compiledMarker.brush && pt.compiledMarker.brush.fill) {
									ret = pt.compiledMarker.brush;
								}
							} else {
								if (pt.brush && pt.brush.fill) {
									ret = pt.brush;
								}
							}
							break;
						}
					}
				} else {
					if (this.selection.markers) {
						if (oSeries.compiledSeriesMarker && oSeries.compiledSeriesMarker.brush && oSeries.compiledSeriesMarker.brush.fill) {
							ret = oSeries.compiledSeriesMarker.brush;
						}
					} else {
						if (oSeries.compiledSeriesBrush && oSeries.compiledSeriesBrush.fill) {
							ret = oSeries.compiledSeriesBrush;
						}
					}
				}
			}
		} else if (AscCommon.isRealObject(this.selection.hiLowLines)) {
			ret = AscFormat.CreateNoFillUniFill();
		} else if (this.selection.upBars) {
			if (oChart && oChart.upDownBars
				&& oChart.upDownBars.upBarsBrush && oChart.upDownBars.upBarsBrush.fill) {
				ret = oChart.upDownBars.upBarsBrush;
			}
		} else if (this.selection.downBars) {
			if (oChart && oChart.upDownBars
				&& oChart.upDownBars.downBarsBrush && oChart.upDownBars.downBarsBrush.fill) {
				ret = oChart.upDownBars.downBarsBrush;
			}
		} else if (this.selection.axis) {

		} else {
			if (this.brush && this.brush.fill) {
				ret = this.brush;
			}
		}
		if (!ret) {
			ret = AscFormat.CreateNoFillUniFill();
		}
		var parents = this.getParentObjects();
		var RGBA = {R: 255, G: 255, B: 255, A: 255};
		ret.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
		return ret;
	};
	CChartSpace.prototype.getStroke = function () {
		var ret = null;
		var oChart = this.getSelectedChart();
		if (this.selection.plotArea) {
			if (this.chart.plotArea.pen && this.chart.plotArea.pen.Fill) {
				ret = this.chart.plotArea.pen;
			}
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var oDlbls = ser.dLbls;
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					if (oDlbls && oDlbls.spPr && oDlbls.spPr.ln && oDlbls.spPr.ln.Fill) {
						ret = oDlbls.spPr.ln;
					}
				} else {
					var pts = ser.getNumPts();
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pt.idx);
						if (dLbl && dLbl.spPr && dLbl.spPr.ln && dLbl.spPr.ln.Fill) {
							ret = dLbl.spPr.ln;
						}
					}
				}
			}
		} else if (this.selection.axisLbls) {
			if (this.selection.axisLbls.spPr && this.selection.axisLbls.spPr.ln && this.selection.axisLbls.spPr.ln.Fill) {
				ret = this.selection.axisLbls.spPr.ln;
			}
		} else if (this.selection.title) {
			if (this.selection.title.spPr && this.selection.title.spPr.ln && this.selection.title.spPr.ln.Fill) {
				ret = this.selection.title.spPr.ln;
			}
		} else if (this.selection.legend) {
			if (this.selection.legend.spPr && this.selection.legend.spPr.ln && this.selection.legend.spPr.ln.Fill) {
				ret = this.selection.legend.spPr.ln;
			}
		} else if (AscFormat.isRealNumber(this.selection.series)) {
			var oSeries = this.getSelectedSeries();
			if (oSeries) {
				if (AscFormat.isRealNumber(this.selection.datPoint)) {
					var pts = oSeries.getNumPts();
					var datPoint = this.selection.datPoint;
					if (oSeries.getObjectType() === AscDFH.historyitem_type_LineSeries ||
						oSeries.getObjectType() === AscDFH.historyitem_type_ScatterSer ||
						oSeries.getObjectType() === AscDFH.historyitem_type_RadarSeries) {
						if (!this.selection.markers) {
							datPoint++;
						}
					}
					for (var j = 0; j < pts.length; ++j) {
						var pt = pts[j];
						if (pt.idx === datPoint) {
							if (this.selection.markers) {
								if (pt && pt.compiledMarker && pt.compiledMarker.pen && pt.compiledMarker.pen.Fill) {
									ret = pt.compiledMarker.pen;
								}
							} else {
								if (pt.pen && pt.pen.Fill) {
									ret = pt.pen;
								}
							}
							break;
						}
					}
				} else {
					if (this.selection.markers) {
						if (oSeries.compiledSeriesMarker && oSeries.compiledSeriesMarker.pen && oSeries.compiledSeriesMarker.pen.Fill) {
							ret = oSeries.compiledSeriesMarker.pen;
						}
					} else {
						if (oSeries.compiledSeriesPen && oSeries.compiledSeriesPen.Fill) {
							ret = oSeries.compiledSeriesPen;
						}
					}
				}
			}
		} else if (AscCommon.isRealObject(this.selection.hiLowLines)) {
			if (oChart && oChart.calculatedHiLowLines
				&& oChart.calculatedHiLowLines.Fill) {
				ret = oChart.calculatedHiLowLines;
			}
		} else if (this.selection.upBars) {
			if (oChart && oChart.upDownBars
				&& oChart.upDownBars.upBarsPen && oChart.upDownBars.upBarsPen.Fill) {
				ret = oChart.upDownBars.upBarsPen;
			}
		} else if (this.selection.downBars) {
			if (oChart && oChart.upDownBars
				&& oChart.upDownBars.downBarsPen && oChart.upDownBars.downBarsPen.Fill) {
				ret = oChart.upDownBars.downBarsPen;
			}
		} else if (this.selection.axis) {
			if (this.selection.majorGridlines) {
				if (this.selection.axis.majorGridlines && this.selection.axis.majorGridlines.ln && this.selection.axis.majorGridlines.ln.Fill) {
					ret = this.selection.axis.majorGridlines.ln;
				}
			}
			if (this.selection.minorGridlines) {
				if (this.selection.axis.minorGridlines && this.selection.axis.minorGridlines.ln && this.selection.axis.minorGridlines.ln.Fill) {
					ret = this.selection.axis.minorGridlines.ln;
				}
			}
		} else {
			if (this.pen && this.pen.Fill) {
				ret = this.pen;
			}
		}
		if (!ret) {
			ret = AscFormat.CreateNoFillLine();
		}
		var parents = this.getParentObjects();
		var RGBA = {R: 255, G: 255, B: 255, A: 255};
		ret.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
		return ret;
	};
	CChartSpace.prototype.canFill = function () {
		if (this.selection.axis) {
			return false;
		}
		if (this.selection.axisLbls) {
			return false;//TODO
		}
		if (AscCommon.isRealObject(this.selection.hiLowLines)) {
			return false;
		}
		if (AscFormat.isRealNumber(this.selection.series) && !this.selection.markers) {
			var oChart = AscCommon.g_oTableId.Get_ById(this.selection.chart);
			if (oChart) {
				if (oChart.getObjectType() === AscDFH.historyitem_type_LineChart
					|| oChart.getObjectType() === AscDFH.historyitem_type_ScatterChart) {
					return false;
				}
			}
		}
		return true;
	};
	CChartSpace.prototype.changeFill = function (unifill) {
		var oChart = AscCommon.g_oTableId.Get_ById(this.selection.chart);

		var unifill2;
		if (this.selection.plotArea) {
			if (!this.chart.plotArea.spPr) {
				this.chart.plotArea.setSpPr(new AscFormat.CSpPr());
				this.chart.plotArea.spPr.setParent(this.chart.plotArea);
			}
			unifill2 = AscFormat.CorrectUniFill(unifill, this.chart.plotArea.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.chart.plotArea.spPr.setFill(unifill2);
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var oDlbls = ser.dLbls;
				if (!ser.dLbls) {

					if (ser.parent && ser.parent.dLbls) {
						oDlbls = ser.parent.dLbls.createDuplicate();
					} else {
						oDlbls = new AscFormat.CDLbls();
					}

					ser.setDLbls(oDlbls);
				}
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					if (!oDlbls.spPr) {
						oDlbls.setSpPr(new AscFormat.CSpPr());
						oDlbls.spPr.setParent(oDlbls);
					}
					var brush = null;
					unifill2 = AscFormat.CorrectUniFill(unifill, (oDlbls.spPr && oDlbls.spPr.Fill) ? oDlbls.spPr.Fill.createDuplicate() : null, this.getEditorType());
					unifill2.convertToPPTXMods();
					oDlbls.spPr.setFill(unifill2);
				} else {
					var pts = ser.getNumPts();
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pt.idx);
						if (!dLbl) {
							dLbl = new AscFormat.CDLbl();
							dLbl.setIdx(pt.idx);
							if (ser.dLbls.txPr) {
								dLbl.merge(ser.dLbls);
							}
							ser.dLbls.addDLbl(dLbl);

						}
						var brush = null;
						if (pt.compiledDlb) {
							brush = pt.compiledDlb.brush;
						}
						if (!dLbl.spPr) {
							dLbl.setSpPr(new AscFormat.CSpPr());
							dLbl.spPr.setParent(dLbl);
						}
						var brush = null;
						unifill2 = AscFormat.CorrectUniFill(unifill, (dLbl.spPr && dLbl.spPr.Fill) ? dLbl.spPr.Fill.createDuplicate() : null, this.getEditorType());
						unifill2.convertToPPTXMods();
						dLbl.spPr.setFill(unifill2);
					}
				}
			}
		} else if (this.selection.axisLbls) {
			if (!this.selection.axisLbls.spPr) {
				this.selection.axisLbls.setSpPr(new AscFormat.CSpPr());
				this.selection.axisLbls.spPr.setParent(this.selection.axisLbls);
			}
			unifill2 = AscFormat.CorrectUniFill(unifill, (this.selection.axisLbls.spPr && this.selection.axisLbls.spPr.Fill) ? this.selection.axisLbls.spPr.Fill.createDuplicate() : null, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.selection.axisLbls.spPr.setFill(unifill2);
		} else if (this.selection.title) {
			if (!this.selection.title.spPr) {
				this.selection.title.setSpPr(new AscFormat.CSpPr());
				this.selection.title.spPr.setParent(this.selection.title);
			}
			unifill2 = AscFormat.CorrectUniFill(unifill, this.selection.title.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.selection.title.spPr.setFill(unifill2);
		} else if (this.selection.legend) {
			if (!this.chart.legend.spPr) {
				this.chart.legend.setSpPr(new AscFormat.CSpPr());
				this.chart.legend.spPr.setParent(this.chart.legend);
			}
			unifill2 = AscFormat.CorrectUniFill(unifill, this.chart.legend.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.chart.legend.spPr.setFill(unifill2);
		} else if (AscFormat.isRealNumber(this.selection.series)) {
			var oSeries = this.getSelectedSeries();
			if (oSeries) {
				if (AscFormat.isRealNumber(this.selection.datPoint)) {
					var pts = oSeries.getNumPts();
					var datPoint = this.selection.datPoint;
					if (oSeries.getObjectType() === AscDFH.historyitem_type_LineSeries ||
						oSeries.getObjectType() === AscDFH.historyitem_type_ScatterSer ||
						oSeries.getObjectType() === AscDFH.historyitem_type_RadarSeries) {
						if (!this.selection.markers) {
							datPoint++;
						}
					}
					for (var j = 0; j < pts.length; ++j) {
						var pt = pts[j];
						if (pt.idx === datPoint) {
							var oDataPoint = null;

							if (Array.isArray(oSeries.dPt)) {
								for (var k = 0; k < oSeries.dPt.length; ++k) {
									if (oSeries.dPt[k].idx === pts[j].idx) {
										oDataPoint = oSeries.dPt[k];
										break;
									}
								}
							}
							if (!oDataPoint) {
								oDataPoint = new AscFormat.CDPt();
								oDataPoint.setIdx(pt.idx);
								oSeries.addDPt(oDataPoint);
							}
							if (this.selection.markers) {
								if (pt.compiledMarker) {
									oDataPoint.setMarker(pt.compiledMarker.createDuplicate());
									if (!oDataPoint.marker.spPr) {
										oDataPoint.marker.setSpPr(new AscFormat.CSpPr());
										oDataPoint.marker.spPr.setParent(oDataPoint.marker);
									}
									unifill2 = AscFormat.CorrectUniFill(unifill, pt.compiledMarker.brush, this.getEditorType());
									unifill2.convertToPPTXMods();
									oDataPoint.marker.spPr.setFill(unifill2);
								}
							} else {
								if (!oDataPoint.spPr) {
									oDataPoint.setSpPr(new AscFormat.CSpPr());
									oDataPoint.spPr.setParent(oDataPoint);
								}
								unifill2 = AscFormat.CorrectUniFill(unifill, pt.brush, this.getEditorType());
								unifill2.convertToPPTXMods();
								oDataPoint.spPr.setFill(unifill2);
							}
							break;
						}
					}
				} else {
					if (this.selection.markers) {
						if (oSeries.setMarker && oSeries.compiledSeriesMarker) {
							if (!oSeries.marker) {
								oSeries.setMarker(oSeries.compiledSeriesMarker.createDuplicate());
							}
							if (!oSeries.marker.spPr) {
								oSeries.marker.setSpPr(new AscFormat.CSpPr());
								oSeries.marker.spPr.setParent(oSeries.marker);
							}
							unifill2 = AscFormat.CorrectUniFill(unifill, oSeries.compiledSeriesMarker.brush, this.getEditorType());
							unifill2.convertToPPTXMods();
							oSeries.marker.spPr.setFill(unifill2);
							if (Array.isArray(oSeries.dPt)) {
								for (var i = 0; i < oSeries.dPt.length; ++i) {
									if (oSeries.dPt[i].marker && oSeries.dPt[i].marker.spPr) {
										oSeries.dPt[i].marker.spPr.setFill(unifill2.createDuplicate());
									}
								}
							}
						}
					} else {
						if (!oSeries.spPr) {
							oSeries.setSpPr(new AscFormat.CSpPr());
							oSeries.spPr.setParent(oSeries);
						}
						unifill2 = AscFormat.CorrectUniFill(unifill, oSeries.compiledSeriesBrush, this.getEditorType());
						unifill2.convertToPPTXMods();
						oSeries.spPr.setFill(unifill2);
						if (Array.isArray(oSeries.dPt)) {
							for (var i = 0; i < oSeries.dPt.length; ++i) {
								if (oSeries.dPt[i].spPr) {
									oSeries.dPt[i].spPr.setFill(unifill2.createDuplicate());
								}
							}
						}
					}
				}
			}
		}
			// else if(AscFormat.isRealNumber(this.selection.hiLowLines))
			// {
		// }
		else if (this.selection.upBars) {
			if (oChart && oChart.upDownBars) {
				if (!oChart.upDownBars.upBars) {
					oChart.upDownBars.setUpBars(new AscFormat.CSpPr());
					oChart.upDownBars.upBars.setParent(oChart.upDownBars);
				}
				unifill2 = AscFormat.CorrectUniFill(unifill, oChart.upDownBars.upBarsBrush, this.getEditorType());
				unifill2.convertToPPTXMods();
				oChart.upDownBars.upBars.setFill(unifill2);
			}
		} else if (this.selection.downBars) {
			if (oChart && oChart.upDownBars) {
				if (!oChart.upDownBars.downBars) {
					oChart.upDownBars.setDownBars(new AscFormat.CSpPr());
					oChart.upDownBars.downBars.setParent(oChart.upDownBars);
				}
				unifill2 = AscFormat.CorrectUniFill(unifill, oChart.upDownBars.downBarsBrush, this.getEditorType());
				unifill2.convertToPPTXMods();
				oChart.upDownBars.downBars.setFill(unifill2);
			}
		}
			// else if(this.selection.axis)
			// {
			//     var oAxis = this.selection.axis;
			//     if(this.selection.majorGridlines)
			//     {
			//
			//     }
		// }
		else {
			if (this.recalcInfo.recalculatePenBrush) {
				this.recalculatePenBrush();
			}
			unifill2 = AscFormat.CorrectUniFill(unifill, this.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			if (!this.spPr) {
				this.setSpPr(new AscFormat.CSpPr());
				this.spPr.setParent(this);
			}
			this.spPr.setFill(unifill2);
		}
		if (!this.spPr) {
			this.setSpPr(new AscFormat.CSpPr());
			this.spPr.setParent(this);
		}
		this.spPr.setFill(this.spPr.Fill ? this.spPr.Fill.createDuplicate() : this.spPr.Fill);
	};
	CChartSpace.prototype.setFill = function (fill) {
		if (!this.spPr) {
			this.setSpPr(new AscFormat.CSpPr());
			this.spPr.setParent(this);
		}
		this.spPr.setFill(fill);
	};
	CChartSpace.prototype.setNvSpPr = function (pr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetNvGrFrProps, this.nvGraphicFramePr, pr));
		this.nvGraphicFramePr = pr;
	};

	CChartSpace.prototype.changeLine = function (line) {
		var stroke;
		var getPenForCorrect = function (oPen) {
			if (oPen) {
				if (oPen.createDuplicate) {
					return oPen.createDuplicate();
				}
				return oPen;
			}
			return null;
		};
		var oChart = AscCommon.g_oTableId.Get_ById(this.selection.chart);
		if (this.selection.plotArea) {
			if (!this.chart.plotArea.spPr) {
				this.chart.plotArea.setSpPr(new AscFormat.CSpPr());
				this.chart.plotArea.spPr.setParent(this.chart.plotArea);
			}
			stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.chart.plotArea.pen));
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.chart.plotArea.spPr.setLn(stroke);
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var oDlbls = ser.dLbls;
				if (!ser.dLbls) {

					if (ser.parent && ser.parent.dLbls) {
						oDlbls = ser.parent.dLbls.createDuplicate();
					} else {
						oDlbls = new AscFormat.CDLbls();
					}

					ser.setDLbls(oDlbls);
				}
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					if (!oDlbls.spPr) {
						oDlbls.setSpPr(new AscFormat.CSpPr());
						oDlbls.spPr.setParent(oDlbls);
					}
					stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oDlbls.spPr.ln), this.getEditorType());
					if (stroke.Fill) {
						stroke.Fill.convertToPPTXMods();
					}
					oDlbls.spPr.setLn(stroke);
				} else {
					var pts = ser.getNumPts();
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = ser.dLbls && ser.dLbls.findDLblByIdx(pt.idx);
						if (!dLbl) {
							dLbl = new AscFormat.CDLbl();
							dLbl.setIdx(pt.idx);
							if (ser.dLbls.txPr) {
								dLbl.merge(ser.dLbls);
							}
							ser.dLbls.addDLbl(dLbl);

						}
						if (!dLbl.spPr) {
							dLbl.setSpPr(new AscFormat.CSpPr());
							dLbl.spPr.setParent(dLbl);
						}
						stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(dLbl.spPr.ln), this.getEditorType());
						if (stroke.Fill) {
							stroke.Fill.convertToPPTXMods();
						}
						dLbl.spPr.setLn(stroke);
					}
				}
			}
		} else if (this.selection.axisLbls) {
			if (!this.selection.axisLbls.spPr) {
				this.selection.axisLbls.setSpPr(new AscFormat.CSpPr());
				this.selection.axisLbls.spPr.setParent(this.selection.axisLbls);
			}
			stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.selection.axisLbls.spPr.ln), this.getEditorType());
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.selection.axisLbls.spPr.setLn(stroke);
		} else if (this.selection.legend) {
			if (!this.chart.legend.spPr) {
				this.chart.legend.setSpPr(new AscFormat.CSpPr());
				this.chart.legend.spPr.setParent(this.chart.legend);
			}
			stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.chart.legend.pen));
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.chart.legend.spPr.setLn(stroke);
		} else if (this.selection.title) {
			if (!this.selection.title.spPr) {
				this.selection.title.setSpPr(new AscFormat.CSpPr());
				this.selection.title.spPr.setParent(this.selection.title);
			}
			stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.selection.title.pen));
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.selection.title.spPr.setLn(stroke);
		} else if (AscFormat.isRealNumber(this.selection.series)) {
			var oSeries = this.getSelectedSeries();
			if (oSeries) {
				if (AscFormat.isRealNumber(this.selection.datPoint)) {
					var pts = oSeries.getNumPts();
					var datPoint = this.selection.datPoint;
					if (oSeries.getObjectType() === AscDFH.historyitem_type_LineSeries ||
						oSeries.getObjectType() === AscDFH.historyitem_type_ScatterSer ||
						oSeries.getObjectType() === AscDFH.historyitem_type_RadarSeries) {
						if (!this.selection.markers) {
							datPoint++;
						}
					}
					for (var j = 0; j < pts.length; ++j) {
						var pt = pts[j];
						if (pt.idx === datPoint) {
							var oDataPoint = null;

							if (Array.isArray(oSeries.dPt)) {
								for (var k = 0; k < oSeries.dPt.length; ++k) {
									if (oSeries.dPt[k].idx === pts[j].idx) {
										oDataPoint = oSeries.dPt[k];
										break;
									}
								}
							}
							if (!oDataPoint) {
								oDataPoint = new AscFormat.CDPt();
								oDataPoint.setIdx(pt.idx);
								oSeries.addDPt(oDataPoint);
							}
							if (this.selection.markers) {
								if (pt.compiledMarker) {
									oDataPoint.setMarker(pt.compiledMarker.createDuplicate());
									if (!oDataPoint.marker.spPr) {
										oDataPoint.marker.setSpPr(new AscFormat.CSpPr());
										oDataPoint.marker.spPr.setParent(oDataPoint.marker);
									}
									stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(pt.compiledMarker.pen));
									if (stroke.Fill) {
										stroke.Fill.convertToPPTXMods();
									}
									oDataPoint.marker.spPr.setLn(stroke);
								}
							} else {
								if (!oDataPoint.spPr) {
									oDataPoint.setSpPr(new AscFormat.CSpPr());
									oDataPoint.spPr.setParent(oDataPoint);
								}
								stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(pt.pen));
								if (stroke.Fill) {
									stroke.Fill.convertToPPTXMods();
								}
								oDataPoint.spPr.setLn(stroke);
							}
							break;
						}
					}
				} else {
					if (this.selection.markers) {
						if (oSeries.setMarker && oSeries.compiledSeriesMarker) {
							if (!oSeries.marker) {
								oSeries.setMarker(oSeries.compiledSeriesMarker.createDuplicate());
							}
							if (!oSeries.marker.spPr) {
								oSeries.marker.setSpPr(new AscFormat.CSpPr());
								oSeries.marker.spPr.setParent(oSeries.marker);
							}
							stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oSeries.compiledSeriesMarker.pen));
							if (stroke.Fill) {
								stroke.Fill.convertToPPTXMods();
							}
							oSeries.marker.spPr.setLn(stroke);


							if (Array.isArray(oSeries.dPt)) {
								for (var i = 0; i < oSeries.dPt.length; ++i) {
									if (oSeries.dPt[i].marker && oSeries.dPt[i].marker.spPr) {
										oSeries.dPt[i].marker.spPr.setLn(stroke.createDuplicate());
									}
								}
							}
						}
					} else {
						if (!oSeries.spPr) {
							oSeries.setSpPr(new AscFormat.CSpPr());
							oSeries.spPr.setParent(oSeries);
						}
						stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oSeries.compiledSeriesPen));
						if (stroke.Fill) {
							stroke.Fill.convertToPPTXMods();
						}
						oSeries.spPr.setLn(stroke);
						if (Array.isArray(oSeries.dPt)) {
							for (var i = 0; i < oSeries.dPt.length; ++i) {
								if (oSeries.dPt[i].spPr) {
									oSeries.dPt[i].spPr.setLn(stroke.createDuplicate());
								}
							}
						}
					}
				}
			}
		} else if (AscCommon.isRealObject(this.selection.hiLowLines)) {
			if (oChart) {
				stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oChart.calculatedHiLowLines));
				if (stroke.Fill) {
					stroke.Fill.convertToPPTXMods();
				}
				this.selection.hiLowLines.setLn(stroke);
			}
		} else if (this.selection.upBars) {
			if (oChart && oChart.upDownBars) {
				if (!oChart.upDownBars.upBars) {
					oChart.upDownBars.setUpBars(new AscFormat.CSpPr());
					oChart.upDownBars.upBars.setParent(oChart.upDownBars);
				}
				stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oChart.upDownBars.upBarsPen));
				if (stroke.Fill) {
					stroke.Fill.convertToPPTXMods();
				}
				oChart.upDownBars.upBars.setLn(stroke);
			}
		} else if (this.selection.downBars) {
			if (oChart && oChart.upDownBars) {
				if (!oChart.upDownBars.downBars) {
					oChart.upDownBars.setDownBars(new AscFormat.CSpPr());
					oChart.upDownBars.downBars.setParent(oChart.upDownBars);
				}
				stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(oChart.upDownBars.downBarsPen));
				if (stroke.Fill) {
					stroke.Fill.convertToPPTXMods();
				}
				oChart.upDownBars.downBars.setLn(stroke);
			}
		} else if (this.selection.axis) {
			var oAxis = this.selection.axis;
			if (this.selection.majorGridlines) {
				if (!this.selection.axis.majorGridlines) {
					this.selection.axis.setMajorGridlines(new AscFormat.CSpPr());
					this.selection.axis.majorGridlines.setParent(this.selection.axis);
				}
				stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.selection.axis.majorGridlines.ln), this.getEditorType());
				if (stroke.Fill) {
					stroke.Fill.convertToPPTXMods();
				}
				this.selection.axis.majorGridlines.setLn(stroke);
			}
			if (this.selection.minorGridlines) {
				if (!this.selection.axis.minorGridlines) {
					this.selection.axis.setMinorGridlines(new AscFormat.CSpPr());
					this.selection.axis.minorGridlines.setParent(this.selection.axis);
				}
				stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.selection.axis.minorGridlines.ln), this.getEditorType());
				if (stroke.Fill) {
					stroke.Fill.convertToPPTXMods();
				}
				this.selection.axis.minorGridlines.setLn(stroke);
			}
		} else {
			if (this.recalcInfo.recalculatePenBrush) {
				this.recalculatePenBrush();
			}
			stroke = AscFormat.CorrectUniStroke(line, getPenForCorrect(this.pen));
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			if (!this.spPr) {
				this.setSpPr(new AscFormat.CSpPr());
				this.spPr.setParent(this);
			}
			this.spPr.setLn(stroke);
		}
		if (!this.spPr) {
			this.setSpPr(new AscFormat.CSpPr());
			this.spPr.setParent(this);
		}
		this.spPr.setFill(this.spPr.Fill ? this.spPr.Fill.createDuplicate() : this.spPr.Fill);
	};
	CChartSpace.prototype.remove = function () {
		if (this.selection.title) {
			if (this.selection.title.parent) {
				this.selection.title.parent.setTitle(null);
			}
		} else if (this.selection.legend) {
			if (!AscFormat.isRealNumber(this.selection.legendEntry)) {
				if (this.selection.legend.parent && this.selection.legend.parent.setLegend) {
					this.selection.legend.parent.setLegend(null);
				}
			} else {
				var entry = this.selection.legend.findLegendEntryByIndex(this.selection.legendEntry);
				if (!entry) {
					entry = new AscFormat.CLegendEntry();
					entry.setIdx(this.selection.legendEntry);
					this.selection.legend.addLegendEntry(entry);
				}
				entry.setDelete(true);
			}
		} else if (this.selection.axisLbls) {
			if (this.selection.axisLbls && this.selection.axisLbls.setDelete) {
				this.selection.axisLbls.setDelete(true);
			}
		} else if (AscFormat.isRealNumber(this.selection.dataLbls)) {
			var ser = this.getAllSeries()[this.selection.dataLbls];
			if (ser) {
				var oDlbls = ser.dLbls;
				if (!ser.dLbls) {

					if (ser.parent && ser.parent.dLbls) {
						oDlbls = ser.parent.dLbls.createDuplicate();
					} else {
						oDlbls = new AscFormat.CDLbls();
					}

					ser.setDLbls(oDlbls);
				}
				if (!AscFormat.isRealNumber(this.selection.dataLbl)) {
					oDlbls.setDeleteValue(true);
				} else {
					var pts = ser.getNumPts();
					var pt = pts[this.selection.dataLbl];
					if (pt) {
						var dLbl = new AscFormat.CDLbl();
						dLbl.setIdx(pt.idx);
						dLbl.setDelete(true);
						ser.dLbls.addDLbl(dLbl);
					}
				}
			}
		}

		this.selection.title = null;
		this.selection.legend = null;
		this.selection.legendEntry = null;
		this.selection.axisLbls = null;
		this.selection.dataLbls = null;
		this.selection.dataLbl = null;
		this.selection.plotArea = null;
		this.selection.rotatePlotArea = null;
		this.selection.gridLine = null;
		this.selection.series = null;
		this.selection.datPoint = null;
		this.selection.markers = null;
		this.selection.textSelection = null;
		this.selection.axis = null;
		this.selection.minorGridlines = null;
		this.selection.majorGridlines = null;
		this.selection.hiLowLines = null;
		this.selection.downBars = null;
		this.selection.upBars = null;
	};
	CChartSpace.prototype.copy = function (oPr) {
		var drawingDocument = oPr && oPr.drawingDocument;
		var copy = new CChartSpace();
		if (this.chart) {
			copy.setChart(this.chart.createDuplicate(drawingDocument));
		}
		if (this.clrMapOvr) {
			copy.setClrMapOvr(this.clrMapOvr.createDuplicate());
		}
		copy.setDate1904(this.date1904);
		if (this.externalData) {
			copy.setExternalData(this.externalData.createDuplicate());
		}
		copy.setLang(this.lang);
		if (this.pivotSource) {
			copy.setPivotSource(this.pivotSource.createDuplicate());
		}
		if (this.printSettings) {
			copy.setPrintSettings(this.printSettings.createDuplicate());
		}
		if (this.protection) {
			copy.setProtection(this.protection.createDuplicate());
		}
		copy.setRoundedCorners(this.roundedCorners);
		if (this.spPr) {
			copy.setSpPr(this.spPr.createDuplicate());
			copy.spPr.setParent(copy);
		}
		copy.setStyle(this.style);
		if (this.txPr) {
			copy.setTxPr(this.txPr.createDuplicate(oPr))
		}
		for (var i = 0; i < this.userShapes.length; ++i) {
			copy.addUserShape(undefined, this.userShapes[i].copy(oPr));
		}
		copy.setThemeOverride(this.themeOverride);
		copy.setBDeleted(this.bDeleted);
		copy.setLocks(this.locks);
		if (this.chartStyle && this.chartColors) {
			copy.setChartStyle(this.chartStyle.createDuplicate());
			copy.setChartColors(this.chartColors.createDuplicate());
		}
		if (this.macro !== null) {
			copy.setMacro(this.macro);
		}
		if (this.textLink !== null) {
			copy.setTextLink(this.textLink);
		}
		if (!oPr || false !== oPr.cacheImage) {
			copy.cachedImage = this.getBase64Img();
			copy.cachedPixH = this.cachedPixH;
			copy.cachedPixW = this.cachedPixW;
		}
		return copy;
	};
	CChartSpace.prototype.convertToWord = function (document) {
		this.setBDeleted(true);
		var oPr = undefined;
		if (document && document.DrawingDocument) {
			oPr = new AscFormat.CCopyObjectProperties();
			oPr.drawingDocument = document.DrawingDocument;
		}
		var oCopy = this.copy(oPr);
		oCopy.setBDeleted(false);
		return oCopy;
	};
	CChartSpace.prototype.convertToPPTX = function (drawingDocument, worksheet) {
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.drawingDocument = drawingDocument;
		var copy = this.copy(oPr);
		copy.setBDeleted(false);
		copy.setWorksheet(worksheet);
		copy.setParent(null);
		return copy;
	};
	CChartSpace.prototype.handleUpdateType = function () {
		this.recalcInfo.recalculateChart = true;
		this.recalcInfo.recalculateSeriesColors = true;
		this.recalcInfo.recalculatePenBrush = true;
		this.recalcInfo.recalculatePlotAreaBrush = true;
		this.recalcInfo.recalculatePlotAreaPen = true;
		this.recalcInfo.recalculateMarkers = true;
		this.recalcInfo.recalculateGridLines = true;
		this.recalcInfo.recalculateDLbls = true;
		this.recalcInfo.recalculateAxisLabels = true;
		//this.recalcInfo.dataLbls.length = 0;
		//this.recalcInfo.axisLabels.length = 0;
		this.recalcInfo.recalculateAxisVal = true;
		this.recalcInfo.recalculateAxisTickMark = true;
		this.recalcInfo.recalculateHiLowLines = true;
		this.recalcInfo.recalculateUpDownBars = true;
		this.recalcInfo.recalculateLegend = true;
		this.recalcInfo.recalculateReferences = true;
		this.recalcInfo.recalculateFormulas = true;
		this.chartObj = null;
		this.addToRecalculate();
	};
	CChartSpace.prototype.handleUpdateExtents = function (bExtX) {
		var oXfrm = this.spPr && this.spPr.xfrm;
		if (undefined === bExtX || !oXfrm || bExtX && !AscFormat.fApproxEqual(this.extX, oXfrm.extX, 0.01) || false === bExtX && !AscFormat.fApproxEqual(this.extY, oXfrm.extY, 0.01)) {
			this.recalcChart();
			this.recalcBounds();
			this.recalcTransform();
			if (this.recalcWrapPolygon) {
				this.recalcWrapPolygon();
			}
			this.recalcTitles();
			this.handleUpdateInternalChart(false);
		}
	};
	CChartSpace.prototype.handleUpdateInternalChart = function (bColors) {
		this.recalcInfo.recalculateChart = true;
		if (bColors !== false) {
			this.recalcInfo.recalculateSeriesColors = true;
			this.recalcInfo.recalculatePenBrush = true;
		}
		this.recalcInfo.recalculateDLbls = true;
		this.recalcInfo.recalculateAxisLabels = true;
		this.recalcInfo.recalculateMarkers = true;
		this.recalcInfo.recalculateGridLines = true;
		this.recalcInfo.recalculateHiLowLines = true;
		this.recalcInfo.recalculatePlotAreaBrush = true;
		this.recalcInfo.recalculatePlotAreaPen = true;
		this.recalcInfo.recalculateAxisTickMark = true;
		//this.recalcInfo.dataLbls.length = 0;
		//this.recalcInfo.axisLabels.length = 0;
		this.recalcInfo.recalculateAxisVal = true;
		this.recalcInfo.recalculateLegend = true;
		this.chartObj = null;

		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object) {
				this.userShapes[i].object.handleUpdateExtents();
			}
		}
		this.addToRecalculate();

	};
	CChartSpace.prototype.handleUpdateGridlines = function () {
		this.recalcInfo.recalculateGridLines = true;
		this.addToRecalculate();
	};
	CChartSpace.prototype.handleUpdateDataLabels = function () {
		this.recalcInfo.recalculateDLbls = true;
		this.addToRecalculate();
	};
	CChartSpace.prototype.updateChildLabelsTransform = function (posX, posY) {

		if (this.localTransformText) {
			this.transformText = this.localTransformText.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(this.transformText, posX, posY);
			this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);
		}

		if (this.chart) {
			if (this.chart.plotArea) {
				this.chart.plotArea.updatePosition(posX, posY);
				var aCharts = this.chart.plotArea.charts;
				for (var t = 0; t < aCharts.length; ++t) {
					var oChart = aCharts[t];
					var series = oChart.series;
					for (var i = 0; i < series.length; ++i) {
						var ser = series[i];
						var pts = ser.getNumPts();
						for (var j = 0; j < pts.length; ++j) {
							if (pts[j].compiledDlb) {
								pts[j].compiledDlb.updatePosition(posX, posY);
							}
						}
					}
				}
				var aAxes = this.chart.plotArea.axId;
				for (var i = 0; i < aAxes.length; ++i) {
					var oAxis = aAxes[i];
					if (oAxis) {
						if (oAxis.title) {
							oAxis.title.updatePosition(posX, posY);
						}
						if (oAxis.labels) {
							oAxis.labels.updatePosition(posX, posY);
						}
					}
				}

			}
			if (this.chart.title) {
				this.chart.title.updatePosition(posX, posY);
			}
			if (this.chart.legend) {
				this.chart.legend.updatePosition(posX, posY);
			}
		}
		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object && this.userShapes[i].object.updatePosition) {
				this.userShapes[i].object.updatePosition(posX, posY);
			}
		}
	};
	CChartSpace.prototype.recalcTitles = function () {
		let aTitles = this.getAllTitles();
		for (let nTitle = 0; nTitle < aTitles.length; ++nTitle) {
			let oTitle = aTitles[nTitle];
			let oRecalcInfo = oTitle.recalcInfo;
			oRecalcInfo.recalculateContent = true;
			oRecalcInfo.recalcTransform = true;
			oRecalcInfo.recalculateTransformText = true;
			oRecalcInfo.recalculateGeometry = true;
		}
	};
	CChartSpace.prototype.recalcTitles2 = function () {
		let aTitles = this.getAllTitles();
		for (let nTitle = 0; nTitle < aTitles.length; ++nTitle) {
			let oTitle = aTitles[nTitle];
			let oRecalcInfo = oTitle.recalcInfo;
			oRecalcInfo.recalculateContent = true;
			oRecalcInfo.recalcTransform = true;
			oRecalcInfo.recalculateTransformText = true;
			oRecalcInfo.recalculateTxBody = true;
			oRecalcInfo.recalculateGeometry = true;
		}
	};
	CChartSpace.prototype.refreshRecalcData2 = function (pageIndex, object) {
		var nObjectType = null;
		if (object && object.getObjectType) {
			nObjectType = object.getObjectType();
		}
		if (nObjectType === AscDFH.historyitem_type_Title && this.selection.title === object && this.selection.textSelection === object) {
			this.recalcInfo.recalcTitle = object;
		} else if (nObjectType === AscDFH.historyitem_type_DLbl) {
			var oSelectedDLbl = this.getCompiledDlblBySelect();

			if (this.selection.textSelection) {
				this.selection.textSelection = oSelectedDLbl;
			}
			this.recalcInfo.recalcTitle = oSelectedDLbl;
		} else {
			var bOldRecalculateRef = this.recalcInfo.recalculateReferences;
			this.setRecalculateInfo();
			this.handleTitlesAfterChangeTheme();
			this.recalcInfo.recalculateReferences = bOldRecalculateRef;
		}
		this.addToRecalculate();
	};
	CChartSpace.prototype.Refresh_RecalcData2 = function (pageIndex, object) {
		this.refreshRecalcData2(pageIndex, object);
	};
	CChartSpace.prototype.Refresh_RecalcData = function (data) {
		switch (data.Type) {
			case AscDFH.historyitem_ChartSpace_SetStyle: {
				this.handleUpdateStyle();
				break;
			}
			case AscDFH.historyitem_ChartSpace_SetTxPr: {
				this.recalcInfo.recalculateChart = true;
				this.recalcInfo.recalculateLegend = true;
				this.recalcInfo.recalculateDLbls = true;
				this.recalcInfo.recalculateAxisVal = true;
				this.recalcInfo.recalculateAxisCat = true;
				this.recalcInfo.recalculateAxisLabels = true;
				this.addToRecalculate();
				break;
			}
			case AscDFH.historyitem_ChartSpace_SetChart: {
				this.handleUpdateType();
				break;
			}
			case AscDFH.historyitem_ShapeSetBDeleted: {
				if (!this.bDeleted) {
					this.handleUpdateType();
				}
				break;
			}
		}
	};
	CChartSpace.prototype.getAllRasterImages = function (images) {
		if (this.spPr) {
			this.spPr.checkBlipFillRasterImage(images);
		}
		var chart = this.chart;
		if (chart) {
			chart.backWall && chart.backWall.spPr && chart.backWall.spPr.checkBlipFillRasterImage(images);
			chart.floor && chart.floor.spPr && chart.floor.spPr.checkBlipFillRasterImage(images);
			chart.legend && chart.legend.spPr && chart.legend.spPr.checkBlipFillRasterImage(images);
			chart.sideWall && chart.sideWall.spPr && chart.sideWall.spPr.checkBlipFillRasterImage(images);
			chart.title && chart.title.spPr && chart.title.spPr.checkBlipFillRasterImage(images);
			//plotArea
			var plot_area = this.chart.plotArea;
			if (plot_area) {
				plot_area.spPr && plot_area.spPr.checkBlipFillRasterImage(images);
				var i;
				for (i = 0; i < plot_area.axId.length; ++i) {
					var axis = plot_area.axId[i];
					if (axis) {
						axis.spPr && axis.spPr.checkBlipFillRasterImage(images);
						axis.title && axis.title && axis.title.spPr && axis.title.spPr.checkBlipFillRasterImage(images);
					}
				}
				for (i = 0; i < plot_area.charts.length; ++i) {
					plot_area.charts[i].getAllRasterImages(images);
				}
			}
		}
		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object && this.userShapes[i].object.getAllRasterImages) {
				this.userShapes[i].object.getAllRasterImages(images);
			}
		}
	};
	CChartSpace.prototype.getAllContents = function () {

	};
	CChartSpace.prototype.getAllFonts = function (allFonts) {
		this.documentGetAllFontNames(allFonts);
	};
	CChartSpace.prototype.documentGetAllFontNames = function (allFonts) {
		allFonts["+mn-lt"] = 1;
		allFonts["+mn-ea"] = 1;
		allFonts["+mn-cs"] = 1;
		checkTxBodyDefFonts(allFonts, this.txPr);
		var chart = this.chart, i;
		if (chart) {
			for (i = 0; i < chart.pivotFmts.length; ++i) {
				chart.pivotFmts[i] && checkTxBodyDefFonts(allFonts, chart.pivotFmts[i].txPr);
			}
			if (chart.legend) {
				checkTxBodyDefFonts(allFonts, chart.legend.txPr);
				for (i = 0; i < chart.legend.legendEntryes.length; ++i) {
					chart.legend.legendEntryes[i] && checkTxBodyDefFonts(allFonts, chart.legend.legendEntryes[i].txPr);
				}
			}
			if (chart.title) {
				checkTxBodyDefFonts(allFonts, chart.title.txPr);
				if (chart.title.tx && chart.title.tx.rich) {
					checkTxBodyDefFonts(allFonts, chart.title.tx.rich);
					chart.title.tx.rich.content && chart.title.tx.rich.content.Document_Get_AllFontNames(allFonts);
				}
			}
			var plot_area = chart.plotArea;
			if (plot_area) {
				for (i = 0; i < plot_area.charts.length; ++i) {
					plot_area.charts[i] && plot_area.charts[i].documentCreateFontMap(allFonts);
					/*TODO надо бы этот метод переименовать чтоб название не вводило в заблуждение*/
				}
				var cur_axis;
				for (i = 0; i < plot_area.axId.length; ++i) {
					cur_axis = plot_area.axId[i];
					checkTxBodyDefFonts(allFonts, cur_axis.txPr);
					if (cur_axis.title) {
						checkTxBodyDefFonts(allFonts, cur_axis.title.txPr);
						if (cur_axis.title.tx && cur_axis.title.tx.rich) {
							checkTxBodyDefFonts(allFonts, cur_axis.title.tx.rich);
							cur_axis.title.tx.rich.content && cur_axis.title.tx.rich.content.Document_Get_AllFontNames(allFonts);
						}
					}
				}
			}
		}

		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object && this.userShapes[i].object.documentGetAllFontNames) {
				this.userShapes[i].object.documentGetAllFontNames(allFonts);
			}
		}
		if (this.themeOverride && this.themeOverride.themeElements && this.themeOverride.themeElements.fontScheme) {
			AscFormat.checkThemeFonts(allFonts, this.themeOverride.themeElements.fontScheme);
		}
	};
	CChartSpace.prototype.documentCreateFontMap = function (allFonts) {
		if (this.chart) {
			this.chart.title && this.chart.title.txBody && this.chart.title.txBody.content.Document_CreateFontMap(allFonts);
			var i, j, k;
			if (this.chart.legend) {
				var calc_entryes = this.chart.legend.calcEntryes;
				for (i = 0; i < calc_entryes.length; ++i) {
					calc_entryes[i].txBody.content.Document_CreateFontMap(allFonts);
				}
			}
			var axis = this.chart.plotArea.axId, cur_axis;
			for (i = axis.length - 1; i > -1; --i) {
				cur_axis = axis[i];
				if (cur_axis) {
					cur_axis && cur_axis.title && cur_axis.title.txBody && cur_axis.title.txBody.content.Document_CreateFontMap(allFonts);
					if (cur_axis.labels) {
						for (j = cur_axis.labels.aLabels.length - 1; j > -1; --j) {
							cur_axis.labels.aLabels[j] && cur_axis.labels.aLabels[j].txBody && cur_axis.labels.aLabels[j].txBody.content.Document_CreateFontMap(allFonts);
						}
					}
				}
			}

			var series, pts;
			for (i = this.chart.plotArea.charts.length - 1; i > -1; --i) {
				series = this.chart.plotArea.charts[i].series;
				for (j = series.length - 1; j > -1; --j) {
					pts = series[j].getNumPts();
					if (Array.isArray(pts)) {
						for (k = pts.length - 1; k > -1; --k) {
							pts[k].compiledDlb && pts[k].compiledDlb.txBody && pts[k].compiledDlb.txBody.content.Document_CreateFontMap(allFonts);
						}
					}
				}
			}
		}

		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object && this.userShapes[i].object.documentCreateFontMap) {
				this.userShapes[i].object.documentCreateFontMap(allFonts);
			}
		}
	};
	CChartSpace.prototype.setThemeOverride = function (pr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetThemeOverride, this.themeOverride, pr));
		this.themeOverride = pr;
	};
	CChartSpace.prototype.addUserShape = function (pos, item) {
		if (!AscFormat.isRealNumber(pos))
			pos = this.userShapes.length;
		History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_ChartSpace_AddUserShape, pos, [item], true));
		this.userShapes.splice(pos, 0, item);
		item.setParent(this);
	};
	CChartSpace.prototype.removeUserShape = function (pos) {
		var aSplicedShape = this.userShapes.splice(pos, 1);
		History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_ChartSpace_RemoveUserShape, pos, aSplicedShape, false));
		return aSplicedShape[0];
	};
	CChartSpace.prototype.setChartStyle = function (pr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_ChartStyle, this.chartStyle, pr));
		this.chartStyle = pr;
	};
	CChartSpace.prototype.setChartColors = function (pr) {
		History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartSpace_ChartColors, this.chartColors, pr));
		this.chartColors = pr;
	};
	CChartSpace.prototype.recalculateUserShapes = function () {
		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object) {
				this.userShapes[i].object.recalculate();
			}
		}
	};
	CChartSpace.prototype.setParent = function (parent) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetParent, this.parent, parent));
		this.parent = parent;
	};
	CChartSpace.prototype.setChart = function (chart) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetChart, this.chart, chart));
		if (this.chart) {
			this.chart.setParent(null);
		}
		this.chart = chart;
		if (chart) {
			chart.setParent(this);
		}
	};
	CChartSpace.prototype.setClrMapOvr = function (clrMapOvr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetClrMapOvr, this.clrMapOvr, clrMapOvr));
		this.clrMapOvr = clrMapOvr;
	};
	CChartSpace.prototype.setDate1904 = function (date1904) {
		History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ChartSpace_SetDate1904, this.date1904, date1904));
		this.date1904 = date1904;
	};
	CChartSpace.prototype.setExternalData = function (externalData) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetExternalData, this.externalData, externalData));
		this.externalData = externalData;
	};
	CChartSpace.prototype.setLang = function (lang) {
		History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ChartSpace_SetLang, this.lang, lang));
		this.lang = lang;
	};
	CChartSpace.prototype.setPivotSource = function (pivotSource) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetPivotSource, this.pivotSource, pivotSource));
		this.pivotSource = pivotSource;
	};
	CChartSpace.prototype.setPrintSettings = function (printSettings) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetPrintSettings, this.printSettings, printSettings));
		this.printSettings = printSettings;
	};
	CChartSpace.prototype.setProtection = function (protection) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetProtection, this.protection, protection));
		this.protection = protection;
	};
	CChartSpace.prototype.setRoundedCorners = function (roundedCorners) {
		History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ChartSpace_SetRoundedCorners, this.roundedCorners, roundedCorners));
		this.roundedCorners = roundedCorners;
	};
	CChartSpace.prototype.setSpPr = function (spPr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetSpPr, this.spPr, spPr));
		this.spPr = spPr;

		if (spPr) {
			spPr.setParent(this);
		}
		this.recalcInfo.recalculateBrush = true;
		this.recalcInfo.recalculatePen = true;
		this.addToRecalculate();
	};
	CChartSpace.prototype.setStyle = function (style) {
		History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ChartSpace_SetStyle, this.style, style));
		this.style = style;
		this.handleUpdateStyle();
	};
	CChartSpace.prototype.setTxPr = function (txPr) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetTxPr, this.txPr, txPr));
		this.txPr = txPr;
		if (txPr) {
			txPr.setParent(this);
		}
	};
	CChartSpace.prototype.getTransformMatrix = function () {
		return this.transform;
	};
	CChartSpace.prototype.canRotate = function () {
		return false;
	};
	CChartSpace.prototype.drawAdjustments = function () {
	};
	CChartSpace.prototype.setGroup = function (group) {
		History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartSpace_SetGroup, this.group, group));
		this.group = group;
	};
	CChartSpace.prototype.hasCharts = function () {
		if (this.chart && this.chart.plotArea && this.chart.plotArea.charts.length > 0) {
			return true;
		}
		return false;
	};
	CChartSpace.prototype.getRangeObjectStr = function () {
		var oDataRange = this.getDataRefs();
		var sRange = oDataRange.getRange();
		var nInfo = oDataRange.getInfo();
		var bVert = (nInfo & AscFormat.SERIES_FLAG_HOR_VALUE) !== 0;
		return {range: sRange, bVert: bVert};
	};

	CChartSpace.prototype.clearDataCache = function () {
		let aSeries = this.getAllSeries();
		for (let nSer = 0; nSer < aSeries.length; ++nSer) {
			aSeries[nSer].clearDataCache();
		}
	};
	CChartSpace.prototype.clearChartDataCache = function () {
		this.clearDataCache();
	};
	CChartSpace.prototype.recalculateReferences = function () {
		var oSelectedSeries = this.getSelectedSeries();
		if (AscFormat.isRealNumber(this.selection.series)) {
			if (!oSelectedSeries) {
				this.selection.series = null;
				this.selection.datPoint = null;
				this.selection.markers = null;
			}
		}
		var worksheet = this.worksheet;
		if (!worksheet)
			return;
		var charts, series, i, j, ser;
		charts = this.chart.plotArea.charts;
		for (i = 0; i < charts.length; ++i) {
			series = charts[i].series;
			for (j = 0; j < series.length; ++j) {
				series[j].updateData(this.displayEmptyCellsAs, this.displayHidden);
			}
		}
		var aTitles = this.getAllTitles();
		for (i = 0; i < aTitles.length; ++i) {
			var oTitle = aTitles[i];
			if (oTitle.tx) {
				oTitle.tx.update();
			}
		}
		var aAxis = this.chart.plotArea.axId;
		for (i = 0; i < aAxis.length; ++i) {
			aAxis[i].updateNumFormat();
		}
	};
	CChartSpace.prototype.checkEmptyVal = function (val) {
		if (val.numRef) {
			if (!val.numRef.numCache)
				return true;
			if (val.numRef.numCache.pts.length === 0)
				return true;
		} else if (val.numLit) {
			if (val.numLit.pts.length === 0)
				return true;
		} else {
			return true;
		}
		return false;
	};
	CChartSpace.prototype.isEmptySeries = function (series, nSeriesLength) {
		for (var i = 0; i < series.length && i < nSeriesLength; ++i) {
			var ser = series[i];
			if (ser.val) {
				if (!this.checkEmptyVal(ser.val))
					return false;
			}
			if (ser.yVal) {
				if (!this.checkEmptyVal(ser.yVal))
					return false;
			}
		}
		return true;
	};
	CChartSpace.prototype.checkEmptySeries = function () {
		for (var t = 0; t < this.chart.plotArea.charts.length; ++t) {
			var chart_type = this.chart.plotArea.charts[t];
			var series = chart_type.series;
			var nChartType = chart_type.getObjectType();
			var nSeriesLength = (nChartType === AscDFH.historyitem_type_PieChart || nChartType === AscDFH.historyitem_type_DoughnutChart) && this.chart.plotArea.charts.length === 1 ? Math.min(1, series.length) : series.length;
			if (this.isEmptySeries(series, nSeriesLength)) {
				return true;
			}
		}
		return t < 1;
	};
	CChartSpace.prototype.getNeedReflect = function () {
		if (!this.chartObj) {
			this.chartObj = new AscFormat.CChartsDrawer();
		}
		return this.chartObj.calculatePositionLabelsCatAxFromAngle(this);
	};
	CChartSpace.prototype.isChart3D = function (oChart) {
		var oOldFirstChart = this.chart.plotArea.charts[0];
		this.chart.plotArea.charts[0] = oChart;
		var ret = AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this);
		this.chart.plotArea.charts[0] = oOldFirstChart;
		return ret;
	};
	CChartSpace.prototype.getAxisCrossType = function (oAxis) {
		if (oAxis.getObjectType() === AscDFH.historyitem_type_ValAx) {
			return AscFormat.CROSS_BETWEEN_MID_CAT;
		}
		if (oAxis.getObjectType() === AscDFH.historyitem_type_SerAx) {
			var oChart = oAxis.getAllCharts()[0];
			if (oChart) {
				if (oChart.getObjectType() === AscDFH.historyitem_type_SurfaceChart) {
					return AscFormat.CROSS_BETWEEN_MID_CAT;
				} else {
					return AscFormat.CROSS_BETWEEN_BETWEEN;
				}
			}
			return AscFormat.CROSS_BETWEEN_MID_CAT;
		}
		var oCrossAxis = oAxis.crossAx;
		if (oCrossAxis && oCrossAxis.getObjectType() === AscDFH.historyitem_type_ValAx) {
			var oChart = oCrossAxis.getAllCharts()[0];
			if (oChart) {
				var nChartType = oChart.getObjectType();
				if (nChartType === AscDFH.historyitem_type_ScatterChart) {
					return null;
				}
				else if(nChartType === AscDFH.historyitem_type_RadarChart) {
					return AscFormat.CROSS_BETWEEN_MID_CAT;
				}
				else if (nChartType !== AscDFH.historyitem_type_BarChart && (nChartType !== AscDFH.historyitem_type_PieChart && nChartType !== AscDFH.historyitem_type_DoughnutChart)
					|| (nChartType === AscDFH.historyitem_type_BarChart && oChart.barDir !== AscFormat.BAR_DIR_BAR)) {
					if (oCrossAxis) {
						if (this.isChart3D(oChart)) {
							if (nChartType === AscDFH.historyitem_type_AreaChart || nChartType === AscDFH.historyitem_type_SurfaceChart) {
								return AscFormat.isRealNumber(oCrossAxis.crossBetween) ? oCrossAxis.crossBetween : AscFormat.CROSS_BETWEEN_MID_CAT;
							} else if (nChartType === AscDFH.historyitem_type_LineChart) {
								return AscFormat.isRealNumber(oCrossAxis.crossBetween) ? oCrossAxis.crossBetween : AscFormat.CROSS_BETWEEN_BETWEEN;
							} else {
								return AscFormat.CROSS_BETWEEN_BETWEEN;
							}
						} else {
							return AscFormat.isRealNumber(oCrossAxis.crossBetween) ? oCrossAxis.crossBetween : ((nChartType === AscDFH.historyitem_type_AreaChart || nChartType === AscDFH.historyitem_type_SurfaceChart) ? AscFormat.CROSS_BETWEEN_MID_CAT : AscFormat.CROSS_BETWEEN_BETWEEN);
						}
					}
				} else if (nChartType === AscDFH.historyitem_type_BarChart && oChart.barDir === AscFormat.BAR_DIR_BAR) {
					return AscFormat.isRealNumber(oCrossAxis.crossBetween) && !this.isChart3D(oChart) ? oCrossAxis.crossBetween : AscFormat.CROSS_BETWEEN_BETWEEN;
				}
			}
		}
		switch (oAxis.getObjectType()) {
			case AscDFH.historyitem_type_ValAx: {
				return AscFormat.CROSS_BETWEEN_MID_CAT;
			}
			case AscDFH.historyitem_type_CatAx:
			case AscDFH.historyitem_type_DateAx: {
				return AscFormat.CROSS_BETWEEN_BETWEEN;
			}
			default: {
				return AscFormat.CROSS_BETWEEN_BETWEEN;
			}
		}
		return AscFormat.CROSS_BETWEEN_BETWEEN;
	};
	CChartSpace.prototype.getValAxisCrossType = function () {
		if (this.chart && this.chart.plotArea && this.chart.plotArea.charts[0]) {
			var chartType = this.chart.plotArea.charts[0].getObjectType();
			var valAx = this.chart.plotArea.valAx;
			if (chartType === AscDFH.historyitem_type_ScatterChart) {
				return null;
			} else if (chartType !== AscDFH.historyitem_type_BarChart && (chartType !== AscDFH.historyitem_type_PieChart && chartType !== AscDFH.historyitem_type_DoughnutChart)
				|| (chartType === AscDFH.historyitem_type_BarChart && this.chart.plotArea.charts[0].barDir !== AscFormat.BAR_DIR_BAR)) {
				if (valAx) {
					if (AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this)) {
						if (chartType === AscDFH.historyitem_type_AreaChart || chartType === AscDFH.historyitem_type_SurfaceChart) {
							return AscFormat.isRealNumber(valAx.crossBetween) ? valAx.crossBetween : AscFormat.CROSS_BETWEEN_MID_CAT;
						} else if (chartType === AscDFH.historyitem_type_LineChart) {
							return AscFormat.isRealNumber(valAx.crossBetween) ? valAx.crossBetween : AscFormat.CROSS_BETWEEN_BETWEEN;
						} else {
							return AscFormat.CROSS_BETWEEN_BETWEEN;
						}
					} else {
						return AscFormat.isRealNumber(valAx.crossBetween) ? valAx.crossBetween : ((chartType === AscDFH.historyitem_type_AreaChart || chartType === AscDFH.historyitem_type_SurfaceChart) ? AscFormat.CROSS_BETWEEN_MID_CAT : AscFormat.CROSS_BETWEEN_BETWEEN);
					}
				}
			} else if (chartType === AscDFH.historyitem_type_BarChart && this.chart.plotArea.charts[0].barDir === AscFormat.BAR_DIR_BAR) {
				if (valAx) {
					return AscFormat.isRealNumber(valAx.crossBetween) && !AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this) ? valAx.crossBetween : AscFormat.CROSS_BETWEEN_BETWEEN;
				}
			}
		}
		return null;
	};
	CChartSpace.prototype.calculatePosByLayout = function (fPos, nLayoutMode, fLayoutValue, fSize, fChartSize) {
		if (!AscFormat.isRealNumber(fLayoutValue)) {
			return fPos;
		}
		var fRetPos = 0;
		if (nLayoutMode === AscFormat.LAYOUT_MODE_EDGE) {
			fRetPos = fChartSize * fLayoutValue;
		} else {
			fRetPos = fPos + fChartSize * fLayoutValue;
		}
		if (fRetPos + fSize > fChartSize) {
			fRetPos -= (fRetPos + fSize - fChartSize);
		}

		if (fRetPos < 0) {
			fRetPos = 0;
		}
		return fRetPos;
	};
	CChartSpace.prototype.calculateSizeByLayout = function (fPos, fChartSize, fLayoutSize, nSizeMode) {
		if (!AscFormat.isRealNumber(fLayoutSize)) {
			return -1;
		}
		var fRetSize = Math.min(fChartSize * fLayoutSize, fChartSize);
		if (nSizeMode === AscFormat.LAYOUT_MODE_EDGE) {
			fRetSize = fRetSize - fPos;
		}
		return fRetSize;
	};
	CChartSpace.prototype.calculateLayoutByPos = function (fPos, nLayoutMode, fPosValue, fChartSize) {
		var fRetLayoutValue = 0;
		if (nLayoutMode === AscFormat.LAYOUT_MODE_EDGE) {
			fRetLayoutValue = fPosValue / fChartSize;
		} else {

			fRetLayoutValue = (fPosValue - fPos) / fChartSize;
		}
		return fRetLayoutValue;
	};
	CChartSpace.prototype.calculateLayoutBySize = function (fPos, nSizeMode, fChartSize, fSize) {
		var fRetLayout;
		if (nSizeMode === AscFormat.LAYOUT_MODE_EDGE) {
			fRetLayout = (fSize - fPos) / fChartSize;
		} else {
			fRetLayout = fSize / fChartSize;
		}
		return fRetLayout;
	};
	CChartSpace.prototype.calculateLabelsPositions = function (b_recalc_labels, b_recalc_legend) {
		var layout;
		for (var i = 0; i < this.recalcInfo.dataLbls.length; ++i) {
			var series = this.getAllSeries();
			if (this.recalcInfo.dataLbls[i].series && this.recalcInfo.dataLbls[i].pt) {

				var ser_idx = this.recalcInfo.dataLbls[i].series.idx; //сделаем проверку лежит ли серия с индексом this.recalcInfo.dataLbls[i].series.idx в сериях первой диаграммы
				for (var j = 0; j < series.length; ++j) {
					if (series[j].idx === this.recalcInfo.dataLbls[i].series.idx) {
						var bLayout = AscCommon.isRealObject(this.recalcInfo.dataLbls[i].layout) && (AscFormat.isRealNumber(this.recalcInfo.dataLbls[i].layout.x) || AscFormat.isRealNumber(this.recalcInfo.dataLbls[i].layout.y));
						var pos = this.chartObj.recalculatePositionText(this.recalcInfo.dataLbls[i]);
						var oLbl = this.recalcInfo.dataLbls[i];
						if (oLbl.layout) {
							layout = oLbl.layout;
							if (AscFormat.isRealNumber(layout.x)) {
								pos.x = this.calculatePosByLayout(pos.x, layout.xMode, layout.x, this.recalcInfo.dataLbls[i].extX, this.extX);
							}
							if (AscFormat.isRealNumber(layout.y)) {
								pos.y = this.calculatePosByLayout(pos.y, layout.yMode, layout.y, this.recalcInfo.dataLbls[i].extY, this.extY);
							}
						}
						if (pos.x + oLbl.extX > this.extX) {
							pos.x -= (pos.x + oLbl.extX - this.extX);
						}
						if (pos.y + oLbl.extY > this.extY) {
							pos.y -= (pos.y + oLbl.extY - this.extY);
						}
						if (pos.x < 0) {
							pos.x = 0;
						}
						if (pos.y < 0) {
							pos.y = 0;
						}
						oLbl.setPosition(pos.x, pos.y);
						break;
					}
				}
			}
		}
		this.recalcInfo.dataLbls.length = 0;

		if (b_recalc_labels) {
			if (this.chart && this.chart.title) {
				var pos = this.chartObj.recalculatePositionText(this.chart.title);
				if (this.chart.title.layout) {
					layout = this.chart.title.layout;
					if (AscFormat.isRealNumber(layout.x)) {
						pos.x = this.calculatePosByLayout(pos.x, layout.xMode, layout.x, this.chart.title.extX, this.extX);
					}
					if (AscFormat.isRealNumber(layout.y)) {
						pos.y = this.calculatePosByLayout(pos.y, layout.yMode, layout.y, this.chart.title.extY, this.extY);
					}
				}
				this.chart.title.setPosition(pos.x, pos.y);
			}

			if (this.chart && this.chart.plotArea) {
				var aAxes = this.chart.plotArea.axId;
				for (var i = 0; i < aAxes.length; ++i) {
					var oAxis = aAxes[i];
					if (oAxis && oAxis.title) {
						var pos = this.chartObj.recalculatePositionText(oAxis.title);

						if (oAxis.title.layout) {
							layout = oAxis.title.layout;
							if (AscFormat.isRealNumber(layout.x)) {
								pos.x = this.calculatePosByLayout(pos.x, layout.xMode, layout.x, oAxis.title.extX, this.extX);
							}
							if (AscFormat.isRealNumber(layout.y)) {
								pos.y = this.calculatePosByLayout(pos.y, layout.yMode, layout.y, oAxis.title.extY, this.extY);
							}
						}
						oAxis.title.setPosition(pos.x, pos.y);
					}
				}


			}
		}

		if (b_recalc_legend && this.chart && this.chart.legend) {
			var bResetLegendPos = false;
			if (!AscFormat.isRealNumber(this.chart.legend.legendPos)) {
				this.chart.legend.legendPos = Asc.c_oAscChartLegendShowSettings.bottom;
				bResetLegendPos = true;
			}
			var pos = this.chartObj.recalculatePositionText(this.chart.legend);
			if (this.chart.legend.layout) {
				layout = this.chart.legend.layout;
				if (AscFormat.isRealNumber(layout.x)) {
					pos.x = this.calculatePosByLayout(pos.x, layout.xMode, layout.x, this.chart.legend.extX, this.extX);
				}
				if (AscFormat.isRealNumber(layout.y)) {
					pos.y = this.calculatePosByLayout(pos.y, layout.yMode, layout.y, this.chart.legend.extY, this.extY);
				}
			}
			this.chart.legend.setPosition(pos.x, pos.y);
			if (bResetLegendPos) {
				this.chart.legend.legendPos = null;
			}
		}
	};
	CChartSpace.prototype.getCatValues = function () {
		var ret = [];
		var aAllSeries = this.getAllSeries(), oFirstSeries;
		if (aAllSeries.length === 0) {
			return [];
		}
		if (aAllSeries.length > 0) {
			oFirstSeries = aAllSeries[0];
			if (oFirstSeries.getObjectType() !== AscDFH.historyitem_type_ScatterSer) {
				return oFirstSeries.getValues(oFirstSeries.getValuesCount());
			} else {
				if (oFirstSeries.xVal) {
					return oFirstSeries.xVal.getValues();
				}
				var oAxes = this.chart.plotArea.getAxisByTypes();
				var oValAxis = null;
				for (var nAxis = 0; nAxis < oAxes.valAx.length; ++nAxis) {
					if (oAxes.valAx[nAxis].isHorizontal()) {
						oValAxis = oAxes.valAx[nAxis];
						break;
					}
				}
				if (oValAxis) {
					if (Array.isArray(oValAxis.scale)) {
						for (var nVal = 0; nVal < oValAxis.scale.length; ++nVal) {
							ret.push(oValAxis.scale[nVal] + "");
						}
					}
				}
			}
		}
		return ret;
	};
	CChartSpace.prototype.getLabelsForAxis = function (oAxis) {
		var aStrings = [];
		var oPlotArea = this.chart.plotArea, i;
		var nAxisType = oAxis.getObjectType();
		var oSeries = oPlotArea.getSeriesWithSmallestIndexForAxis(oAxis);
		var bCat = false;
		switch (nAxisType) {
			case AscDFH.historyitem_type_DateAx:
			case AscDFH.historyitem_type_CatAx: {
				//расчитаем подписи для горизонтальной оси
				var nPtsLen = 0;
				var aScale = [];
				if (Array.isArray(oAxis.scale)) {
					aScale = aScale.concat(oAxis.scale);

				}
				if (oSeries && oSeries.cat) {
					var oCat = oSeries.cat;
					var oLit = oCat.getLit();
					if (oLit) {
						bCat = true;
						var oLitFormat = null, oPtFormat = null;
						if (typeof oLit.formatCode === "string" && oLit.formatCode.length > 0) {
							oLitFormat = oNumFormatCache.get(oLit.formatCode);
						}
						if (/*nAxisType === AscDFH.historyitem_type_DateAx && */oAxis.numFmt && typeof oAxis.numFmt.formatCode === "string" && oAxis.numFmt.formatCode.length > 0) {
							oLitFormat = oNumFormatCache.get(oAxis.numFmt.formatCode);
						}
						nPtsLen = oLit.ptCount;

						var bTickSkip = AscFormat.isRealNumber(oAxis.tickLblSkip) || nPtsLen >= SKIP_LBL_LIMIT;
						var nTickLblSkip = AscFormat.isRealNumber(oAxis.tickLblSkip) ? oAxis.tickLblSkip : (nPtsLen < SKIP_LBL_LIMIT ? 1 : (Math.floor(nPtsLen / SKIP_LBL_LIMIT) + 1));
						let nLastNoEmptyLblIdx = -1;
						for (i = 0; i < nPtsLen; ++i) {
							if (!bTickSkip ||
								nLastNoEmptyLblIdx === -1 || ((i - nLastNoEmptyLblIdx) >= nTickLblSkip)) {
								var oPt = oLit.getPtByIndex(i);
								if (oPt) {
									var sPt;
									if (typeof oPt.formatCode === "string" && oPt.formatCode.length > 0) {
										oPtFormat = oNumFormatCache.get(oPt.formatCode);
										if (oPtFormat) {
											sPt = oPtFormat.formatToChart(oPt.val);
										} else {
											sPt = oPt.val + "";
										}
									} else if (oLitFormat) {
										sPt = oLitFormat.formatToChart(oPt.val);
									} else {
										sPt = oPt.val + "";
									}
									aStrings.push(sPt);
									if (typeof sPt === "string" && sPt.length > 0) {
										nLastNoEmptyLblIdx = i;
									}
								} else {
									aStrings.push(null);
								}
							} else {
								aStrings.push(null);
							}
						}
					}
				}
				var nPtsLength = 0;
				var aChartsForAxis = oAxis.getAllCharts();
				for (i = 0; i < aChartsForAxis.length; ++i) {
					var oChart = aChartsForAxis[i];
					for (var j = 0; j < oChart.series.length; ++j) {
						var oCurPts = null;
						oSeries = oChart.series[j];
						if (oSeries.val) {
							if (oSeries.val.numRef && oSeries.val.numRef.numCache) {
								oCurPts = oSeries.val.numRef.numCache;
							} else if (oSeries.val.numLit) {
								oCurPts = oSeries.val.numLit;
							}
							if (oCurPts) {
								nPtsLength = Math.max(nPtsLength, oCurPts.ptCount);
							}
						}
					}
				}
				var nCrossBetween = this.getAxisCrossType(oAxis);
				if (nCrossBetween === AscFormat.CROSS_BETWEEN_MID_CAT && nPtsLength < 2) {
					nPtsLength = 2;
				}
				var oLitFormatDate = null;
				if (nAxisType === AscDFH.historyitem_type_DateAx && oAxis.numFmt && typeof oAxis.numFmt.formatCode === "string" && oAxis.numFmt.formatCode.length > 0) {
					oLitFormatDate = oNumFormatCache.get(oAxis.numFmt.formatCode);
				}

				if (nPtsLength > aStrings.length) {
					bTickSkip = AscFormat.isRealNumber(oAxis.tickLblSkip) || nPtsLength >= SKIP_LBL_LIMIT;
					nTickLblSkip = AscFormat.isRealNumber(oAxis.tickLblSkip) ? oAxis.tickLblSkip : (nPtsLength < SKIP_LBL_LIMIT ? 1 : (Math.floor(nPtsLength / SKIP_LBL_LIMIT) + 1));
					var nStartLength = aStrings.length;
					for (i = aStrings.length; i < nPtsLength; ++i) {
						if (!bCat && (!bTickSkip || (((nStartLength + i) % nTickLblSkip) === 0))) {
							if (oLitFormatDate) {
								aStrings.push(oLitFormatDate.formatToChart(i + 1));
							} else {
								aStrings.push(i + 1 + "");
							}
						} else {
							aStrings.push(null);
						}
					}
				} else {
					aStrings.splice(nPtsLength, aStrings.length - nPtsLength);
				}
				if (aScale.length > 0) {
					while (aStrings.length < aScale[aScale.length - 1]) {
						aStrings.push("");
					}
					for (i = 0; i < aScale.length; ++i) {
						if (aScale[i] > 0) {
							break;
						}
						aStrings.splice(0, 0, "");
					}
				}

				break;
			}
			case AscDFH.historyitem_type_ValAx: {
				var aVal = [].concat(oAxis.scale);
				var fMultiplier;
				if (oAxis.dispUnits) {
					fMultiplier = oAxis.dispUnits.getMultiplier();
				} else {
					fMultiplier = 1.0;
				}
				var oNumFormat = null;
				var sFormatCode = oAxis.getFormatCode();
				if (typeof sFormatCode === "string") {
					oNumFormat = oNumFormatCache.get(sFormatCode);
				}
				if (!oNumFormat) {
					oNumFormat = oNumFormatCache.get("General");
				}
				for (var t = 0; t < aVal.length; ++t) {
					var fCalcValue = aVal[t] * fMultiplier;
					var sRichValue;
					if (oNumFormat) {
						sRichValue = oNumFormat.formatToChart(fCalcValue);
					} else {
						sRichValue = fCalcValue + "";
					}
					aStrings.push(sRichValue);
				}
				break;
			}
			case AscDFH.historyitem_type_SerAx: {
				var aAllSeries = this.getAllSeries();
				oAxis.scale = [];
				for (var nSer = 0; nSer < aAllSeries.length; ++nSer) {
					aStrings.push(aAllSeries[nSer].getSeriesName());
					oAxis.scale.push(aAllSeries[nSer].idx);
				}

				break;
			}
		}
		return aStrings;
	};
	CChartSpace.prototype.calculateAxisGrid = function (oAxis, oRect) {
		if (!oAxis) {
			return;
		}

		let oAxisGrid = new CAxisGrid();
		oAxis.grid = oAxisGrid;
		let aStrings = this.getLabelsForAxis(oAxis);
		let nOrientation = oAxis.scaling && AscFormat.isRealNumber(oAxis.scaling.orientation) ? oAxis.scaling.orientation : AscFormat.ORIENTATION_MIN_MAX;
		if(oAxis.isRadarCategories()) {
			let nIntervalsCount = aStrings.length;
			oAxisGrid.nCount = nIntervalsCount;
			oAxisGrid.bOnTickMark = true;
			oAxisGrid.aStrings = aStrings;
			if (nOrientation === AscFormat.ORIENTATION_MIN_MAX) {
				oAxisGrid.fStart = 0;
				oAxisGrid.fStride = 2*Math.PI / nIntervalsCount;
			} else {
				oAxisGrid.fStart = 2*Math.PI;
				oAxisGrid.fStride = -2*Math.PI / nIntervalsCount;
			}
		} else {

			let nCrossType = this.getAxisCrossType(oAxis);
			let bOnTickMark = ((nCrossType === AscFormat.CROSS_BETWEEN_MID_CAT) && (aStrings.length > 1));
			let nIntervalsCount = bOnTickMark ? (aStrings.length - 1) : (aStrings.length);
			let fInterval;
			oAxisGrid.nCount = nIntervalsCount;
			oAxisGrid.bOnTickMark = bOnTickMark;
			oAxisGrid.aStrings = aStrings;
			if (oAxis.getObjectType() === AscDFH.historyitem_type_SerAx) {
				this.checkPrecalculateChartObject();
				var dDepth = this.getDepthPerspective();
				fInterval = dDepth / nIntervalsCount;
				if (nOrientation === AscFormat.ORIENTATION_MIN_MAX) {
					oAxisGrid.fStart = 0;
					oAxisGrid.fStride = fInterval;
				} else {
					oAxisGrid.fStart = dDepth;
					oAxisGrid.fStride = -fInterval;
				}
			} else if (oAxis.isHorizontal()) {
				fInterval = oRect.w / nIntervalsCount;
				if (nOrientation === AscFormat.ORIENTATION_MIN_MAX) {
					oAxisGrid.fStart = oRect.x;
					oAxisGrid.fStride = fInterval;
				} else {
					oAxisGrid.fStart = oRect.x + oRect.w;
					oAxisGrid.fStride = -fInterval;
				}
			} else {
				oAxis.yPoints = [];
				let fRectH = oRect.h;
				let fRectY = oRect.y;
				if(oAxis.isRadarValues()) {
					fRectH /= 2.0;
				}
				fInterval = fRectH / nIntervalsCount;
				if (nOrientation === AscFormat.ORIENTATION_MIN_MAX) {
					oAxisGrid.fStart = fRectY + fRectH;
					oAxisGrid.fStride = -fInterval;
				} else {
					oAxisGrid.fStart = fRectY;
					oAxisGrid.fStride = fInterval;
				}
			}
		}


	};
	CChartSpace.prototype.getView3d = function () {
		return this.chart.getView3d();
	};
	CChartSpace.prototype.changeView3d = function (oView3D) {
		let oCurView3D = this.getView3d();
		if (!oCurView3D && !oView3D) {
			return;
		}
		if (oCurView3D && oView3D && oView3D.isEqual(oCurView3D)) {
			return;
		}
		this.chart.setView3D(oView3D);
		this.setRecalculateInfo();
	};
	CChartSpace.prototype.getDepthPerspective = function () {
		var oProcessor3D = this.chartObj && this.chartObj.processor3D;
		var dDepth = 0;
		if (oProcessor3D && AscFormat.isRealNumber(oProcessor3D.depthPerspective)) {
			dDepth = oProcessor3D.depthPerspective;
		}
		return dDepth;
	};
	CChartSpace.prototype.recalculateAxesSet = function(aAxesSet, oRect, oBaseRect, nIndex, fForceContentWidth) {
		let oCorrectedRect = null;

		let bWithoutLabels = false;
		if(this.chart.plotArea.layout && this.chart.plotArea.layout.layoutTarget === AscFormat.LAYOUT_TARGET_INNER) {
			bWithoutLabels = true;
		}
		let bCorrected = false;
		let fL = oRect.x, fT = oRect.y, fR = oRect.x + oRect.w, fB = oRect.y + oRect.h;
		let fHorPadding = 0.0;
		let fVertPadding = 0.0;
		let fHorInterval = null;
		let oCalcMap = {};
		for(let nAxesSet = 0; nAxesSet < aAxesSet.length; ++nAxesSet) {
			let oCurAxis = aAxesSet[nAxesSet];
			let oCrossAxis = oCurAxis.crossAx;
			if(!oCalcMap[oCurAxis.Id]) {
				this.calculateAxisGrid(oCurAxis, oRect);
				oCalcMap[oCurAxis.Id] = true;
			}
			if(!oCalcMap[oCrossAxis.Id]) {
				this.calculateAxisGrid(oCrossAxis, oRect);
				oCalcMap[oCrossAxis.Id] = true;
			}
			let fCrossValue;
			let fAxisPos;
			let fDistance = DEFAULT_LBLS_DISTANCE;
			let fDistanceSign = 1;
			let nLabelsPos;
			let bLabelsExtremePosition = false;
			let bOnTickMark = oCurAxis.grid.bOnTickMark;
			if(oCurAxis.bDelete) {
				nLabelsPos = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE;
			}
			else {
				if(null !== oCurAxis.tickLblPos) {
					nLabelsPos = oCurAxis.tickLblPos;
				}
				else {
					nLabelsPos = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
				}
			}

			let bCrossAt = false;

			let oCrossGrid = oCrossAxis.grid;
			if(AscFormat.isRealNumber(oCurAxis.crossesAt) && oCrossAxis.scale[0] <= oCurAxis.crossesAt && oCrossAxis.scale[oCrossAxis.scale.length - 1] >= oCurAxis.crossesAt) {
				fCrossValue = oCurAxis.crossesAt;
				bCrossAt = true;
			}
			else {
				switch(oCurAxis.crosses) {
					case AscFormat.CROSSES_MAX:
					{
						fCrossValue = oCrossAxis.scale[oCrossAxis.scale.length - 1];
						if(!oCrossGrid.bOnTickMark) {
							fCrossValue += 1;
						}
						if(nLabelsPos === c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO) {
							fDistanceSign = -fDistanceSign;
							bLabelsExtremePosition = true;
						}
						break;
					}
					case AscFormat.CROSSES_MIN:
					{
						fCrossValue = oCrossAxis.scale[0];
						if(nLabelsPos === c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO) {
							bLabelsExtremePosition = true;
						}
						if(!oCrossGrid.bOnTickMark) {
							fCrossValue -= 1;
						}
						break;
					}
					default:
					{ //includes AutoZero
						if(oCrossAxis.scale[0] <= 0 && oCrossAxis.scale[oCrossAxis.scale.length - 1] >= 0) {
							fCrossValue = 0;
						}
						else if(oCrossAxis.scale[0] > 0) {
							fCrossValue = oCrossAxis.scale[0];
						}
						else {
							fCrossValue = oCrossAxis.scale[oCrossAxis.scale.length - 1];
						}
					}
				}
			}

			if(AscFormat.fApproxEqual(fCrossValue, oCrossAxis.scale[0]) || AscFormat.fApproxEqual(fCrossValue, oCrossAxis.scale[oCrossAxis.scale.length - 1])) {
				bLabelsExtremePosition = true;
			}
			if(oCurAxis.isRadarAxis()) {
				bLabelsExtremePosition = false;
			}
			let bKoeff = 1.0;
			if(oCrossAxis.scale.length > 1) {
				bKoeff = oCrossAxis.scale[1] - oCrossAxis.scale[0];
			}
			fAxisPos = oCrossGrid.fStart;
			if(oCurAxis.isRadarValues()) {
				fAxisPos = oRect.x + oRect.w / 2;
			}
			else {
				if(oCrossAxis.scale.length > 0) {
					fAxisPos += (fCrossValue - oCrossAxis.scale[0]) * (oCrossGrid.fStride) / bKoeff;
				}
			}

			let nOrientation = isRealObject(oCrossAxis.scaling) && AscFormat.isRealNumber(oCrossAxis.scaling.orientation) ? oCrossAxis.scaling.orientation : AscFormat.ORIENTATION_MIN_MAX;
			if(nOrientation === AscFormat.ORIENTATION_MAX_MIN) {
				fDistanceSign = -fDistanceSign;
			}
			let oLabelsBox = null, fPos;
			let fPosStart = oCurAxis.grid.fStart;
			let fPosEnd = oCurAxis.grid.fStart + oCurAxis.grid.nCount * oCurAxis.grid.fStride;

			let bForceVertical = false;
			let bNumbers = false;//TODO

			if(nLabelsPos !== c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE) {
				oLabelsBox = new CLabelsBox(oCurAxis.grid.aStrings, oCurAxis, this);
				switch(nLabelsPos) {
					case c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO:
					{
						fPos = fAxisPos;
						break;
					}
					case c_oAscTickLabelsPos.TICK_LABEL_POSITION_HIGH:
					{
						fPos = oCrossGrid.fStart + oCrossGrid.nCount * oCrossGrid.fStride;
						fDistanceSign = -fDistanceSign;
						break;
					}
					case c_oAscTickLabelsPos.TICK_LABEL_POSITION_LOW:
					{
						fPos = oCrossGrid.fStart;
						break;
					}
				}
			}



			oCurAxis.labels = oLabelsBox;
			oCurAxis.posX = null;
			oCurAxis.posY = null;
			oCurAxis.nullPos = null;
			oCurAxis.xPoints = null;
			oCurAxis.yPoints = null;
			oCurAxis.zPoints = null;
			oCurAxis.alphaPoints = null;
			let aPoints = null;
			if(oCurAxis.isSeriesAxis()) {
				//TODO
			}
			else if(oCurAxis.isRadarCategories()) {
				oCurAxis.alphaPoints = [];
				aPoints = oCurAxis.alphaPoints;
				if (null !== aPoints) {
					oCurAxis.grid.calculatePoints(aPoints, bOnTickMark, oCurAxis.scale);
				}
				fLayoutRadarCatLabelsBox(oLabelsBox, oRect, aPoints);
			}
			else if(oCurAxis.isHorizontal()) {
				oCurAxis.posY = fAxisPos;
				oCurAxis.xPoints = [];
				aPoints = oCurAxis.xPoints;
				if(oLabelsBox) {
					if(!AscFormat.fApproxEqual(oRect.fVertPadding, 0)) {
						fPos -= oRect.fVertPadding;
						if(nLabelsPos === c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO) {
							oCurAxis.posY -= oRect.fVertPadding;
						}
					}


					let nTickLblSkip = AscFormat.isRealNumber(oCurAxis.tickLblSkip) ? oCurAxis.tickLblSkip : 1;
					let bTickSkip = nTickLblSkip > 1;
					let fAxisLength = fPosEnd - fPosStart;
					let nLabelsCount = oLabelsBox.aLabels.length;

					let bOnTickMark_ = bOnTickMark && nLabelsCount > 1;
					let nIntervalCount = bOnTickMark_ ? nLabelsCount - 1 : nLabelsCount;
					fHorInterval = Math.abs(fAxisLength / nIntervalCount);
					if(bTickSkip && !AscFormat.isRealNumber(fForceContentWidth)) {
						fForceContentWidth = Math.abs(fHorInterval) + fHorInterval / nTickLblSkip;
					}
					fDistance = fDistanceSign * oLabelsBox.getLabelsOffset();
					fLayoutHorLabelsBox(oLabelsBox, fPos, fPosStart, fPosEnd, bOnTickMark, fDistance, bForceVertical, bNumbers, fForceContentWidth);
					if(bLabelsExtremePosition) {
						if(fDistance > 0) {
							fVertPadding = -oLabelsBox.extY;
						}
						else {
							fVertPadding = oLabelsBox.extY;
						}
					}
				}

				if (null !== aPoints) {
					oCurAxis.grid.calculatePoints(aPoints, bOnTickMark, oCurAxis.scale);
				}
			}
			else {//vertical axis
				fDistanceSign = -fDistanceSign;
				let oView3D = this.chart.getView3d();
				if(oView3D) {
					let nRotY = oView3D.rotY;
					if(AscFormat.isRealNumber(nRotY)) {
						if(nRotY >= 90 && nRotY < 270) {
							fDistanceSign = -fDistanceSign;
						}
					}
				}
				if(oCurAxis.isRadarAxis()) {
					fDistanceSign = -1;
					if(nOrientation === AscFormat.ORIENTATION_MAX_MIN) {
						fDistanceSign = 1;
					}
				}
				oCurAxis.posX = fAxisPos;
				oCurAxis.yPoints = [];
				aPoints = oCurAxis.yPoints;
				if(oLabelsBox) {
					if(!AscFormat.fApproxEqual(oRect.fHorPadding, 0)) {
						fPos -= oRect.fHorPadding;
						if(nLabelsPos === c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO) {
							oCurAxis.posX -= oRect.fHorPadding;
						}
					}
					fDistance = fDistanceSign * oLabelsBox.getLabelsOffset();
					fLayoutVertLabelsBox(oLabelsBox, fPos, fPosStart, fPosEnd, bOnTickMark, fDistance, bForceVertical);
					if(bLabelsExtremePosition) {
						if(fDistance > 0) {
							fHorPadding = -oLabelsBox.extX;
						}
						else {
							fHorPadding = oLabelsBox.extX;
						}
					}
				}

				if (null !== aPoints) {
					oCurAxis.grid.calculatePoints(aPoints, bOnTickMark, oCurAxis.scale);
				}
			}
			if (null !== aPoints) {
				if(aPoints.length > 1) {
					let fNullPos;
					let oP1 = aPoints[0];
					let oP2 = aPoints[aPoints.length - 1];
					let fAK = (oP1.pos - oP2.pos) / (oP1.val - oP2.val);
					let fBK = oP1.pos - fAK*oP1.val;
					oCurAxis.nullPos = fBK;
				}
			}
			if(oLabelsBox) {
				if(oLabelsBox.x < fL) {
					fL = oLabelsBox.x;
				}
				if(oLabelsBox.x + oLabelsBox.extX > fR) {
					fR = oLabelsBox.x + oLabelsBox.extX;
				}
				if(oLabelsBox.y < fT) {
					fT = oLabelsBox.y;
				}
				if(oLabelsBox.y + oLabelsBox.extY > fB) {
					fB = oLabelsBox.y + oLabelsBox.extY;
				}
			}
		}
		if(nIndex < 2) {
			let fDiff;
			let fPrecision = 0.01;
			oCorrectedRect = new CRect(oRect.x, oRect.y, oRect.w, oRect.h);
			let bWEdge = false;
			let bHEge = false;
			let oPALayout = this.chart.plotArea.layout;
			if(oPALayout) {
				let iN = AscFormat.isRealNumber;
				if(oPALayout.wMode === AscFormat.LAYOUT_MODE_EDGE && iN(oPALayout.w)) {
					bWEdge = true;
				}
				if(oPALayout.hMode === AscFormat.LAYOUT_MODE_EDGE && iN(oPALayout.h)) {
					bHEge = true;
				}
			}
			if(bWithoutLabels) {
				fDiff = fL;
				if(fDiff < 0.0 && !AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.x -= fDiff;

					if(bWEdge) {
						oCorrectedRect.w += fDiff;
					}
					bCorrected = true;
				}
				fDiff = fR - this.extX;
				if(fDiff > 0.0 && !AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.w -= fDiff;
					bCorrected = true;
				}
				fDiff = fT;
				if(fDiff < 0.0 && !AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.y -= fDiff;
					if(bHEge) {
						oCorrectedRect.h += fDiff;
					}
					bCorrected = true;
				}
				fDiff = fB - this.extY;
				if(fDiff > 0.0 && !AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.h -= fDiff;
					bCorrected = true;
				}
			}
			else {
				fDiff = oBaseRect.x - fL;
				if(/*fDiff > 0.0 && */!AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.x += fDiff;
					if(bWEdge) {
						oCorrectedRect.w -= fDiff;
					}
					bCorrected = true;
				}
				fDiff = oBaseRect.x + oBaseRect.w - fR;
				if(/*fDiff < 0.0 && */!AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.w += fDiff;
					bCorrected = true;
				}
				fDiff = oBaseRect.y - fT;
				if(/*fDiff > 0.0 &&*/ !AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.y += fDiff;
					if(bHEge) {
						oCorrectedRect.h -= fDiff;
					}
					bCorrected = true;
				}
				fDiff = oBaseRect.y + oBaseRect.h - fB;
				if(/*fDiff < 0.0 && */!AscFormat.fApproxEqual(fDiff, 0.0, fPrecision)) {
					oCorrectedRect.h += fDiff;
					bCorrected = true;
				}
			}
			if(oCorrectedRect && bCorrected) {
				if(oCorrectedRect.w > oRect.w) {
					return this.recalculateAxesSet(aAxesSet, oCorrectedRect, oBaseRect, ++nIndex, fHorInterval);
				}
				else {
					return this.recalculateAxesSet(aAxesSet, oCorrectedRect, oBaseRect, ++nIndex);
				}
			}
		}
		let _ret = oRect.copy();
		let oAx = aAxesSet[0];
		if(oAx && oAx.isRadarAxis()) {
			let dXC = _ret.x + _ret.w / 2;
			let dYC = _ret.y + _ret.h / 2;
			let dSize = Math.min(_ret.w / 2, _ret.h / 2);
			let dL = dXC - dSize;
			let dT = dYC - dSize;
			_ret.x = dL;
			_ret.y = dT;
			_ret.w = 2*dSize;
			_ret.h = 2*dSize;
		}
		_ret.fHorPadding = fHorPadding;
		_ret.fVertPadding = fVertPadding;
		return _ret;
	};
	CChartSpace.prototype.getReplaceAxis = function (oAxis) {
		var aAxes = this.chart.plotArea.axId;
		if (oAxis.bDelete && oAxis.getObjectType() === AscDFH.historyitem_type_ValAx) {
			var bHorizontal = oAxis.isHorizontal();
			for (var j = 0; j < aAxes.length; ++j) {
				var oCheckAxis = aAxes[j];
				if (!oCheckAxis.bDelete) {
					if (bHorizontal) {
						if (oCheckAxis.isHorizontal()) {
							if (!oAxis.crossAx || oAxis.crossAx && oCheckAxis.crossAx && oAxis.crossAx.getObjectType() === oCheckAxis.crossAx.getObjectType()) {
								return oCheckAxis;
							}
						}
					} else {
						if (oCheckAxis.axPos === AscFormat.AX_POS_R || oCheckAxis.axPos === AscFormat.AX_POS_L) {
							if (!oAxis.crossAx || oAxis.crossAx && oCheckAxis.crossAx && oAxis.crossAx.getObjectType() === oCheckAxis.crossAx.getObjectType()) {
								return oCheckAxis;
							}
						}
					}
				}
			}
		}
		return null;
	};
	CChartSpace.prototype.recalculateAxes = function () {
		this.cachedCanvas = null;
		this.plotAreaRect = null;
		this.bEmptySeries = this.checkEmptySeries();
		if (this.chart && this.chart.plotArea) {
			var oPlotArea = this.chart.plotArea;
			for (i = 0; i < oPlotArea.axId.length; ++i) {
				oCurAxis = oPlotArea.axId[i];
				oCurAxis.posY = null;
				oCurAxis.posX = null;
				oCurAxis.xPoints = null;
				oCurAxis.yPoints = null;
				oCurAxis.zPoints = null;
				oCurAxis.alphaPoints = null;
			}
			if (this.bEmptySeries) {
				return;
			}
			let oSize = this.getChartSizes();
			let oRect = new CRect(oSize.startX, oSize.startY, oSize.w, oSize.h);
			if (!this.chartObj) {
				this.chartObj = new AscFormat.CChartsDrawer()
			}
			var i, j;
			var aCharts = oPlotArea.charts, oChart;
			var oCurAxis, oCurAxis2, aCurAxesSet;
			//temporary add axes to charts with deleted axes
			var oChartsToAxesCount = {};
			var bAdd, nAxesCount;
			var oReplaceAxis;

			var oAxesMap = {};
			for (i = 0; i < aCharts.length; ++i) {
				bAdd = false;
				oChart = aCharts[i];
				if (oChart.axId) {
					nAxesCount = oChart.axId.length;
					for (j = nAxesCount - 1; j > -1; --j) {
						oCurAxis = oChart.axId[j];
						oReplaceAxis = this.getReplaceAxis(oCurAxis);
						if (oReplaceAxis) {
							bAdd = true;
							oChart.axId.push(oReplaceAxis);
						}
					}
					if (bAdd) {
						oChartsToAxesCount[oChart.Id] = nAxesCount;
					}
					for (j = 0; j < oChart.axId.length; ++j) {
						oAxesMap[oChart.axId[j].Id] = oChart.axId[j];
					}
				}
			}
			this.chartObj.preCalculateData(this);
			for (i in oChartsToAxesCount) {
				if (oChartsToAxesCount.hasOwnProperty(i)) {
					oChart = AscCommon.g_oTableId.Get_ById(i);
					if (oChart) {
						oChart.axId.length = oChartsToAxesCount[i]
					}
				}
			}

			var aAxes = [];
			for (var sAxId in oAxesMap) {
				if (oAxesMap.hasOwnProperty(sAxId)) {
					aAxes.push(oAxesMap[sAxId]);
				}
			}
			var aAllAxes = [];//array of axes sets
			var aSeriesAxes = [];
			var dSeriesLabelsWidth = 0;
			var oSetAxis;
			while (aAxes.length > 0) {
				oCurAxis = aAxes.splice(0, 1)[0];
				if (oCurAxis.getObjectType() !== AscDFH.historyitem_type_SerAx) {
					aCurAxesSet = [];
					aCurAxesSet.push(oCurAxis);
					for (i = aAxes.length - 1; i > -1; --i) {
						oCurAxis2 = aAxes[i];
						if (oCurAxis2.getObjectType() !== AscDFH.historyitem_type_SerAx) {
							for (j = aCurAxesSet.length - 1; j > -1; --j) {
								oSetAxis = aCurAxesSet[j];
								if (oSetAxis.crossAx !== oSetAxis && oSetAxis.crossAx === oCurAxis2 ||
									oCurAxis2.crossAx !== oCurAxis2 && oCurAxis2.crossAx === oSetAxis) {
									aCurAxesSet.push(oCurAxis2);
								}
							}
						}
					}
					if (aCurAxesSet.length > 1) {
						aAllAxes.push(aCurAxesSet);
					}
				} else {
					aSeriesAxes.push(oCurAxis);
					this.calculateAxisGrid(oCurAxis, oRect);
					oCurAxis.labels = null;
					var nLabelsPos = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
					if (oCurAxis.bDelete) {
						nLabelsPos = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE;
					} else {
						if (null !== oCurAxis.tickLblPos) {
							nLabelsPos = oCurAxis.tickLblPos;
						} else {
							nLabelsPos = c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO;
						}
					}
					if (nLabelsPos !== c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE) {
						var oLabelsBox = new CLabelsBox(oCurAxis.grid.aStrings, oCurAxis, this);
						oCurAxis.labels = oLabelsBox;
						dSeriesLabelsWidth = oLabelsBox.maxMinWidth;
					}
				}
			}

			var oBaseRect = oRect;
			var aRects = [];
			var fForceContentWidth = undefined;
			for(i = 0; i < aAllAxes.length; ++i) {
				aCurAxesSet = aAllAxes[i];
				aRects.push(this.recalculateAxesSet(aCurAxesSet, oRect, oBaseRect, 0, fForceContentWidth));
			}
			if (aRects.length > 1) {
				oRect = aRects[0].copy();
				for (i = 1; i < aRects.length; ++i) {
					if (!oRect.intersection(aRects[i])) {
						break;
					}
				}
				var fOldHorPadding = 0.0, fOldVertPadding = 0.0;
				if (i === aRects.length) {
					var aRects2 = [];
					for (i = 0; i < aAllAxes.length; ++i) {
						aCurAxesSet = aAllAxes[i];
						if (i === 0) {
							fOldHorPadding = oRect.fHorPadding;
							fOldVertPadding = oRect.fVertPadding;
							oRect.fHorPadding = 0.0;
							oRect.fVertPadding = 0.0;
						}
						aRects2.push(this.recalculateAxesSet(aCurAxesSet, oRect, oBaseRect, 2));
						if (i === 0) {
							oRect.fHorPadding = fOldHorPadding;
							oRect.fVertPadding = fOldVertPadding;
						}
					}

					var bCheckPaddings = false;
					for (i = 0; i < aRects.length; ++i) {
						if (Math.abs(aRects2[i].fVertPadding) > Math.abs(aRects[i].fVertPadding)) {
							if (aRects2[i].fVertPadding > 0) {
								aRects[i].y += (aRects2[i].fVertPadding - aRects[i].fVertPadding);
								aRects[i].h -= (aRects2[i].fVertPadding - aRects[i].fVertPadding);
							} else {
								aRects[i].h -= Math.abs(aRects2[i].fVertPadding - aRects[i].fVertPadding);
							}
							aRects[i].fVertPadding = aRects2[i].fVertPadding;
							bCheckPaddings = true;
						}
						if (Math.abs(aRects2[i].fHorPadding) > Math.abs(aRects[i].fHorPadding)) {
							if (aRects2[i].fHorPadding > 0) {
								aRects[i].x += (aRects2[i].fHorPadding - aRects[i].fHorPadding);
								aRects[i].w -= (aRects2[i].fHorPadding - aRects[i].fHorPadding);
							} else {
								aRects[i].w -= Math.abs(aRects2[i].fHorPadding - aRects[i].fHorPadding);
							}
							aRects[i].fHorPadding = aRects2[i].fHorPadding;
							bCheckPaddings = true;
						}
					}
					if (bCheckPaddings) {
						oRect = aRects[0].copy();
						for (i = 1; i < aRects.length; ++i) {
							if (!oRect.intersection(aRects[i])) {
								break;
							}
						}
						if (i === aRects.length) {
							var aRects2 = [];
							for (i = 0; i < aAllAxes.length; ++i) {
								aCurAxesSet = aAllAxes[i];
								if (i === 0) {
									fOldHorPadding = oRect.fHorPadding;
									fOldVertPadding = oRect.fVertPadding;
									oRect.fHorPadding = 0.0;
									oRect.fVertPadding = 0.0;
								}
								aRects2.push(this.recalculateAxesSet(aCurAxesSet, oRect, oBaseRect, 2));
								if (i === 0) {
									oRect.fHorPadding = fOldHorPadding;
									oRect.fVertPadding = fOldVertPadding;
								}
							}
						}
					}
				}
				this.plotAreaRect = oRect.copy();
			} else {
				if (aRects[0]) {
					this.plotAreaRect = aRects[0].copy();
				}
			}
			aAxes = oPlotArea.axId;
			var oHorAxis;
			for (i = 0; i < aAxes.length; ++i) {
				oCurAxis = aAxes[i];
				var bHorizontal = oCurAxis.isHorizontal();
				oReplaceAxis = this.getReplaceAxis(oCurAxis);
				if (oReplaceAxis) {
					if (bHorizontal) {
						oCurAxis.xPoints = oReplaceAxis.xPoints;
					} else {
						oCurAxis.yPoints = oReplaceAxis.yPoints;
					}
				}
				if (bHorizontal && oCurAxis.getObjectType() !== AscDFH.historyitem_type_SerAx) {
					oHorAxis = oCurAxis;
				}
			}
			if (oHorAxis) {
				for (var nSerAx = 0; nSerAx < aSeriesAxes.length; ++nSerAx) {
					var oSerAx = aSeriesAxes[nSerAx];
					var oGrid = oHorAxis.grid;
					oSerAx.posX = oGrid.fStart + oGrid.fStride * (oGrid.nCount);
					oSerAx.posY = oHorAxis.posY;
					oSerAx.zPoints = [];
					var aAllSeries = this.getAllSeries();
					this.recalculateChart();
					this.calculateAxisGrid(oSerAx);


					var nOrientation = oSerAx.scaling && AscFormat.isRealNumber(oSerAx.scaling.orientation) ? oSerAx.scaling.orientation : AscFormat.ORIENTATION_MIN_MAX;
					var nCrossType = this.getAxisCrossType(oSerAx);
					var bOnTickMark = ((nCrossType === AscFormat.CROSS_BETWEEN_MID_CAT) && (aAllSeries.length > 1));
					var nIntervalsCount = bOnTickMark ? (aAllSeries.length - 1) : (aAllSeries.length);
					oGrid = oSerAx.grid;
					var dDepth = this.getDepthPerspective();
					var fStart, fStride;
					if (nOrientation === AscFormat.ORIENTATION_MIN_MAX) {
						fStart = 0;
						fStride = dDepth / nIntervalsCount;
					} else {
						fStart = dDepth;
						fStride = -dDepth / nIntervalsCount;
					}
					var nSer;
					var fAdd = bOnTickMark ? 0 : fStride / 2;
					for (nSer = 0; nSer < aAllSeries.length; ++nSer) {
						oSerAx.zPoints.push({val: aAllSeries[nSer].idx, pos: fStart + fStride * nSer + fAdd});
					}
					if (oSerAx.labels) {
						let oLabelsBox = oSerAx.labels;
						oLabelsBox.layoutHorNormal(oSerAx.posY, oLabelsBox.getLabelsOffset(), oSerAx.posX, 0, oSerAx.grid.bOnTickMark, 2000);
					}
				}
			}

			let oChartSize = this.getChartSizes(true);
			this.chart.plotArea.x = oChartSize.startX;
			this.chart.plotArea.y = oChartSize.startY;
			this.chart.plotArea.extX = oChartSize.w;
			this.chart.plotArea.extY = oChartSize.h;
			this.chart.plotArea.localTransform.Reset();
			AscCommon.global_MatrixTransformer.TranslateAppend(this.chart.plotArea.localTransform, oChartSize.startX, oChartSize.startY);
		}
	};
	CChartSpace.prototype.distributeValues = function (aVals) {
		if (Array.isArray(aVals) && aVals.length > 2) {
			var fStride = (aVals[aVals.length - 1] - aVals[0]) / (aVals.length - 1);
			for (var nVal = 1; nVal < aVals.length - 1; ++nVal) {
				aVals[nVal] = aVals[0] + fStride * nVal;
			}
		}
	};
	CChartSpace.prototype.checkAxisLabelsTransform = function () {
		if (this.chart && this.chart.plotArea && this.chart.plotArea.charts[0] && this.chart.plotArea.charts[0].getAxisByTypes) {
			var oAxisByTypes = this.chart.plotArea.charts[0].getAxisByTypes();
			var oCatAx = oAxisByTypes.catAx[0], oValAx = oAxisByTypes.valAx[0], deltaX, deltaY, i, oAxisLabels, oLabel,
				oNewPos;
			var oSerAx = oAxisByTypes.serAx[0];
			var oProcessor3D = this.chartObj && this.chartObj.processor3D;
			var aXPoints = [], aYPoints = [];
			if (oCatAx && oValAx && oProcessor3D) {
				if ((oCatAx.isHorizontal() && oCatAx.xPoints) &&
					(oValAx.isVertical() && oValAx.yPoints)) {
					oAxisLabels = oCatAx.labels;

					if (oAxisLabels) {
						var dZPositionCatAxis = oProcessor3D.calculateZPositionCatAxis();
						var dPosY, dPosY2
						if (oAxisLabels.align) {
							dPosY = oAxisLabels.y * this.chartObj.calcProp.pxToMM;
							dPosY2 = oAxisLabels.y;
						} else {

							dPosY = (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM;
							dPosY2 = oAxisLabels.y + oAxisLabels.extY;
						}
						var fBottomLabels = -100;
						var fHalfWidth;
						if (!oAxisLabels.bRotated) {
							for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
								oLabel = oAxisLabels.aLabels[i];
								if (oLabel) {
									var oCPosLabelX;
									fHalfWidth = oLabel.tx.rich.content.XLimit / 2;
									oCPosLabelX = oLabel.localTransformText.TransformPointX(fHalfWidth, 0);
									oNewPos = oProcessor3D.convertAndTurnPoint(oCPosLabelX * this.chartObj.calcProp.pxToMM, dPosY, dZPositionCatAxis);
									oLabel.setPosition2(oNewPos.x / this.chartObj.calcProp.pxToMM + oLabel.localTransformText.tx - oCPosLabelX, oLabel.localTransformText.ty - dPosY2 + oNewPos.y / this.chartObj.calcProp.pxToMM);
									var fBottomContent = oLabel.y + oLabel.tx.rich.content.GetSummaryHeight();
									if (fBottomContent > fBottomLabels) {
										fBottomLabels = fBottomContent;
									}

								}
							}
						} else {
							if (oAxisLabels.align) {
								var labels_offset = oCatAx.labels.getLabelsOffset();
								for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
									if (oAxisLabels.aLabels[i]) {
										oLabel = oAxisLabels.aLabels[i];
										var wh = {
											w: oLabel.tx.rich.getContentWidth(),
											h: oLabel.tx.rich.content.GetSummaryHeight()
										}, w2, h2, x1, y0, xc, yc;
										w2 = wh.w * Math.cos(Math.PI / 4) + wh.h * Math.sin(Math.PI / 4);
										h2 = wh.w * Math.sin(Math.PI / 4) + wh.h * Math.cos(Math.PI / 4);
										x1 = oCatAx.xPoints[i].pos + wh.h * Math.sin(Math.PI / 4);
										y0 = oAxisLabels.y + labels_offset;
										var x1t, y0t;
										var oRes = oProcessor3D.convertAndTurnPoint(x1 * this.chartObj.calcProp.pxToMM, y0 * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
										x1t = oRes.x / this.chartObj.calcProp.pxToMM;
										y0t = oRes.y / this.chartObj.calcProp.pxToMM;
										xc = x1t - w2 / 2;
										yc = y0t + h2 / 2;
										var local_text_transform = oLabel.localTransformText;
										local_text_transform.Reset();
										global_MatrixTransformer.TranslateAppend(local_text_transform, -wh.w / 2, -wh.h / 2);
										global_MatrixTransformer.RotateRadAppend(local_text_transform, Math.PI / 4);
										global_MatrixTransformer.TranslateAppend(local_text_transform, xc, yc);

										var fBottomContent = y0t + h2;
										if (fBottomContent > fBottomLabels) {
											fBottomLabels = fBottomContent;
										}
									}
								}
							} else {
								var labels_offset = oCatAx.labels.getLabelsOffset();
								for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
									if (oAxisLabels.aLabels[i]) {
										oLabel = oAxisLabels.aLabels[i];
										var wh = {
											w: oLabel.tx.rich.getContentWidth(),
											h: oLabel.tx.rich.content.GetSummaryHeight()
										}, w2, h2, x1, y0, xc, yc;
										w2 = wh.w * Math.cos(Math.PI / 4) + wh.h * Math.sin(Math.PI / 4);
										h2 = wh.w * Math.sin(Math.PI / 4) + wh.h * Math.cos(Math.PI / 4);
										x1 = oCatAx.xPoints[i].pos - wh.h * Math.sin(Math.PI / 4);
										y0 = oAxisLabels.y + oAxisLabels.extY - labels_offset;
										var x1t, y0t;
										var oRes = oProcessor3D.convertAndTurnPoint(x1 * this.chartObj.calcProp.pxToMM, y0 * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
										x1t = oRes.x / this.chartObj.calcProp.pxToMM;
										y0t = oRes.y / this.chartObj.calcProp.pxToMM;
										xc = x1t + w2 / 2;
										yc = y0t - h2 / 2;
										local_text_transform = oLabel.localTransformText;
										local_text_transform.Reset();
										global_MatrixTransformer.TranslateAppend(local_text_transform, -wh.w / 2, -wh.h / 2);
										global_MatrixTransformer.RotateRadAppend(local_text_transform, Math.PI / 4);//TODO
										global_MatrixTransformer.TranslateAppend(local_text_transform, xc, yc);
									}
								}
							}
						}

						oNewPos = oProcessor3D.convertAndTurnPoint(oAxisLabels.x * this.chartObj.calcProp.pxToMM, oAxisLabels.y * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x + oAxisLabels.extX) * this.chartObj.calcProp.pxToMM, oAxisLabels.y * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x + oAxisLabels.extX) * this.chartObj.calcProp.pxToMM, (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x) * this.chartObj.calcProp.pxToMM, (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oAxisLabels.x = Math.min.apply(Math, aXPoints);
						oAxisLabels.y = Math.min.apply(Math, aYPoints);
						oAxisLabels.extX = Math.max.apply(Math, aXPoints) - oAxisLabels.x;
						oAxisLabels.extY = Math.max(Math.max.apply(Math, aYPoints), fBottomLabels) - oAxisLabels.y;
					}
					oAxisLabels = oValAx.labels;
					if (oAxisLabels) {
						var dZPositionCatAxis = oProcessor3D.calculateZPositionValAxis();

						var dPosX, dPosX2;
						if (!oAxisLabels.align) {
							dPosX2 = oAxisLabels.x;
						} else {
							dPosX2 = oAxisLabels.x + oAxisLabels.extX;
						}
						dPosX = dPosX2 * this.chartObj.calcProp.pxToMM;
						aXPoints.length = 0;
						aYPoints.length = 0;
						for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
							oLabel = oAxisLabels.aLabels[i];
							if (oLabel) {
								oNewPos = oProcessor3D.convertAndTurnPoint(dPosX, oLabel.localTransformText.ty * this.chartObj.calcProp.pxToMM, dZPositionCatAxis);
								oLabel.setPosition2(oLabel.localTransformText.tx - dPosX2 + oNewPos.x / this.chartObj.calcProp.pxToMM, oNewPos.y / this.chartObj.calcProp.pxToMM);
								aXPoints.push(oLabel.x);
								aYPoints.push(oLabel.y);
								aXPoints.push(oLabel.x + oLabel.tx.rich.getContentWidth());
								aYPoints.push(oLabel.y + oLabel.txBody.content.GetSummaryHeight());
							}
						}
						if (aXPoints.length > 0 && aYPoints.length > 0) {
							oAxisLabels.x = Math.min.apply(Math, aXPoints);
							oAxisLabels.y = Math.min.apply(Math, aYPoints);
							oAxisLabels.extX = Math.max.apply(Math, aXPoints) - oAxisLabels.x;
							oAxisLabels.extY = Math.max.apply(Math, aYPoints) - oAxisLabels.y;
						}
					}
					if (oSerAx) {
						var dPosX, dPosX2;
						dPosX2 = oSerAx.posX;
						dPosX = dPosX2 * this.chartObj.calcProp.pxToMM;

						var dPosY, dPosY2

						dPosY2 = oSerAx.posY;
						dPosY = dPosY2 * this.chartObj.calcProp.pxToMM;

						oAxisLabels = oSerAx.labels;
						if (oAxisLabels) {
							aXPoints.length = 0;
							aYPoints.length = 0;
							for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
								oLabel = oAxisLabels.aLabels[i];
								if (oLabel) {
									var oPt = oSerAx.zPoints[i];
									if (oPt) {
										oNewPos = oProcessor3D.convertAndTurnPoint(dPosX, dPosY, oPt.pos);
										oLabel.setPosition2(oLabel.localTransformText.tx - dPosX2 + oNewPos.x / this.chartObj.calcProp.pxToMM + oLabel.tx.rich.getContentWidth() / 2 + DEFAULT_LBLS_DISTANCE,
											oNewPos.y / this.chartObj.calcProp.pxToMM - oLabel.txBody.content.GetSummaryHeight() / 2);
										aXPoints.push(oLabel.x);
										aYPoints.push(oLabel.y);
										aXPoints.push(oLabel.x + oLabel.tx.rich.getContentWidth());
										aYPoints.push(oLabel.y + oLabel.txBody.content.GetSummaryHeight());
									}
								}
							}
							if (aXPoints.length > 0 && aYPoints.length > 0) {
								oAxisLabels.x = Math.min.apply(Math, aXPoints);
								oAxisLabels.y = Math.min.apply(Math, aYPoints);
								oAxisLabels.extX = Math.max.apply(Math, aXPoints) - oAxisLabels.x;
								oAxisLabels.extY = Math.max.apply(Math, aYPoints) - oAxisLabels.y;
							}
						}
					}
				} else if (((oCatAx.isVertical()) && oCatAx.yPoints) &&
					((oValAx.isHorizontal()) && oValAx.xPoints)) {
					oAxisLabels = oValAx.labels;


					if (oAxisLabels) {
						var dZPositionValAxis = oProcessor3D.calculateZPositionValAxis();
						var dPosY, dPosY2
						if (oAxisLabels.align) {
							dPosY = oAxisLabels.y * this.chartObj.calcProp.pxToMM;
							dPosY2 = oAxisLabels.y;
						} else {
							dPosY = (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM;
							dPosY2 = oAxisLabels.y + oAxisLabels.extY;
						}

						aXPoints.length = 0;
						aYPoints.length = 0;
						for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
							oLabel = oAxisLabels.aLabels[i];
							if (oLabel) {
								var oCPosLabelX;
								fHalfWidth = oLabel.tx.rich.content.XLimit / 2;
								oCPosLabelX = oLabel.localTransformText.TransformPointX(fHalfWidth, 0);
								oNewPos = oProcessor3D.convertAndTurnPoint(oCPosLabelX * this.chartObj.calcProp.pxToMM, dPosY, dZPositionValAxis);
								oLabel.setPosition2(oNewPos.x / this.chartObj.calcProp.pxToMM + oLabel.localTransformText.tx - oCPosLabelX, oLabel.localTransformText.ty - dPosY2 + oNewPos.y / this.chartObj.calcProp.pxToMM);
								aXPoints.push(oLabel.x);
								aYPoints.push(oLabel.y);
								aXPoints.push(oLabel.x + oLabel.tx.rich.getContentWidth());
								aYPoints.push(oLabel.y + oLabel.txBody.content.GetSummaryHeight());
							}
						}

						oNewPos = oProcessor3D.convertAndTurnPoint(oAxisLabels.x * this.chartObj.calcProp.pxToMM, oAxisLabels.y * this.chartObj.calcProp.pxToMM, dZPositionValAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x + oAxisLabels.extX) * this.chartObj.calcProp.pxToMM, oAxisLabels.y * this.chartObj.calcProp.pxToMM, dZPositionValAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x + oAxisLabels.extX) * this.chartObj.calcProp.pxToMM, (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM, dZPositionValAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oNewPos = oProcessor3D.convertAndTurnPoint((oAxisLabels.x) * this.chartObj.calcProp.pxToMM, (oAxisLabels.y + oAxisLabels.extY) * this.chartObj.calcProp.pxToMM, dZPositionValAxis);
						aXPoints.push(oNewPos.x / this.chartObj.calcProp.pxToMM);
						aYPoints.push(oNewPos.y / this.chartObj.calcProp.pxToMM);
						oAxisLabels.x = Math.min.apply(Math, aXPoints);
						oAxisLabels.y = Math.min.apply(Math, aYPoints);
						oAxisLabels.extX = Math.max.apply(Math, aXPoints) - oAxisLabels.x;
						oAxisLabels.extY = Math.max.apply(Math, aYPoints) - oAxisLabels.y;
					}


					oAxisLabels = oCatAx.labels;

					aXPoints.length = 0;
					aYPoints.length = 0;
					if (oAxisLabels) {
						var dZPositionValAxis = oProcessor3D.calculateZPositionValAxis();
						var dPosX, dPosX2;
						if (oAxisLabels.align) {
							dPosX2 = oAxisLabels.x;
						} else {
							dPosX2 = oAxisLabels.x + oAxisLabels.extX;
						}
						dPosX = dPosX2 * this.chartObj.calcProp.pxToMM;
						for (i = 0; i < oAxisLabels.aLabels.length; ++i) {
							oLabel = oAxisLabels.aLabels[i];

							if (oLabel) {
								oNewPos = oProcessor3D.convertAndTurnPoint(dPosX, oLabel.localTransformText.ty * this.chartObj.calcProp.pxToMM, dZPositionValAxis);
								oLabel.setPosition2(oLabel.localTransformText.tx - dPosX2 + oNewPos.x / this.chartObj.calcProp.pxToMM, oNewPos.y / this.chartObj.calcProp.pxToMM);
								aXPoints.push(oLabel.x);
								aYPoints.push(oLabel.y);
								aXPoints.push(oLabel.x + oLabel.txBody.getContentWidth());
								aYPoints.push(oLabel.y + oLabel.txBody.content.GetSummaryHeight());
							}
						}
						if (aXPoints.length > 0 && aYPoints.length > 0) {
							oAxisLabels.x = Math.min.apply(Math, aXPoints);
							oAxisLabels.y = Math.min.apply(Math, aYPoints);
							oAxisLabels.extX = Math.max.apply(Math, aXPoints) - oAxisLabels.x;
							oAxisLabels.extY = Math.max.apply(Math, aYPoints) - oAxisLabels.y;
						}
					}
				}
			}
		}
	};
	CChartSpace.prototype.layoutLegendEntry = function (oEntry, fX, fY, fDistance) {
		oEntry.localY = fY;
		var oFirstLine = oEntry.txBody.content.Content[0].Lines[0];
		var fYFirstLineMiddle = oEntry.localY + oFirstLine.Top + oFirstLine.Metrics.Descent / 2.0 + oFirstLine.Metrics.TextAscent - oFirstLine.Metrics.TextAscent2 / 2.0;// (oFirstLine.Bottom - oFirstLine.Top - oFirstLine.Metrics.Descent - oFirstLine.Metrics.Ascent)/2.0;
		var fPenWidth;
		var oUnionMarker = oEntry.calcMarkerUnion;
		var oLineMarker = oUnionMarker.lineMarker;
		var oMarker = oUnionMarker.marker;
		if (oLineMarker) {
			oLineMarker.localX = fX;
			if (oLineMarker.pen) {
				if (AscFormat.isRealNumber(oLineMarker.pen.w)) {
					fPenWidth = oLineMarker.pen.w / 36000.0;
				} else {
					fPenWidth = 12700.0 / 36000.0;
				}
			} else {
				fPenWidth = 0.0;
			}
			oLineMarker.localY = fYFirstLineMiddle - fPenWidth / 2.0;
			if (oMarker) {
				oMarker.localX = oLineMarker.localX + oLineMarker.spPr.geometry.gdLst['w'] / 2.0 - oMarker.spPr.geometry.gdLst['w'] / 2.0;
				oMarker.localY = fYFirstLineMiddle - oMarker.spPr.geometry.gdLst['h'] / 2.0;
			}
			oEntry.localX = oLineMarker.localX + oLineMarker.spPr.geometry.gdLst['w'] + fDistance;
		} else {
			if (oMarker) {
				oMarker.localX = fX;
				var fMarkerWidth = oMarker.spPr.geometry.gdLst['w'];
				oMarker.localY = fYFirstLineMiddle - oMarker.spPr.geometry.gdLst['h'] / 2.0;
				oEntry.localX = fX + fMarkerWidth + fDistance;
				if (this.chartStyle && this.chartStyle.isSpecialStyle()) {
					var fSpecialSize = oFirstLine.Bottom - oFirstLine.Top;
					oMarker.spPr.geometry.Recalculate(fSpecialSize, fSpecialSize);
					oMarker.localX = fX + (fMarkerWidth - fSpecialSize) / 2.0;
					oMarker.localY = fYFirstLineMiddle - fSpecialSize / 2.0;
				}
			} else {
				oEntry.localX = fX + fDistance;
			}
		}
	};
	CChartSpace.prototype.hitInTextRect = function () {
		return false;
	};
	CChartSpace.prototype.recalculateLegend = function () {
		if (this.chart && this.chart.legend) {
			var parents = this.getParentObjects();
			var RGBA = {R: 0, G: 0, B: 0, A: 255};
			var legend = this.chart.legend;
			var arr_str_labels = [], i, j;
			var calc_entryes = legend.calcEntryes;
			calc_entryes.length = 0;
			var series;
			var legend_pos = c_oAscChartLegendShowSettings.right;
			if (AscFormat.isRealNumber(legend.legendPos)) {
				legend_pos = legend.legendPos;
			}
			var aCharts = this.chart.plotArea.charts;
			var oTypedChart;
			//Order series for the legend
			var aOrderedSeries = [];
			var aChartSeries;
			var nChart;
			var bStraightOrder;
			var bVericalLegend = false;
			if (aCharts.length === 1 && (legend_pos === c_oAscChartLegendShowSettings.left ||
				legend_pos === c_oAscChartLegendShowSettings.right ||
				legend_pos === c_oAscChartLegendShowSettings.leftOverlay ||
				legend_pos === c_oAscChartLegendShowSettings.rightOverlay ||
				legend_pos === c_oAscChartLegendShowSettings.topRight)) {
				bVericalLegend = true;
			}
			for (nChart = 0; nChart < aCharts.length; ++nChart) {
				oTypedChart = aCharts[nChart];
				aChartSeries = [].concat(oTypedChart.series);
				bStraightOrder = true;
				if (bVericalLegend) {
					if (oTypedChart.getObjectType() === AscDFH.historyitem_type_BarChart) {
						if (oTypedChart.grouping === AscFormat.BAR_GROUPING_STACKED
							|| oTypedChart.grouping === AscFormat.BAR_GROUPING_PERCENT_STACKED) {
							bStraightOrder = false;
						}
					} else {
						if (oTypedChart.grouping === AscFormat.GROUPING_STACKED
							|| oTypedChart.grouping === AscFormat.GROUPING_PERCENT_STACKED) {
							bStraightOrder = false;
						}
					}
				}
				if (bStraightOrder) {
					aChartSeries.sort(function (a, b) {
						return a.order - b.order;
					});
				} else {
					aChartSeries.sort(function (a, b) {
						return b.order - a.order;
					});
				}
				if (oTypedChart.getObjectType() === AscDFH.historyitem_type_DoughnutChart ||
					oTypedChart.getObjectType() === AscDFH.historyitem_type_PieChart ||
					oTypedChart.getObjectType() === AscDFH.historyitem_type_OfPieChart ||
					oTypedChart.getObjectType() === AscDFH.historyitem_type_AreaChart) {
					aOrderedSeries = aChartSeries.concat(aOrderedSeries);
				} else {
					aOrderedSeries = aOrderedSeries.concat(aChartSeries);
				}
			}
			series = aOrderedSeries;
			var calc_entry, union_marker, entry;
			var max_width = 0, cur_width, max_font_size = 0, cur_font_size, ser, b_line_series;
			var b_no_line_series = false;
			this.chart.legend.chart = this;
			var b_scatter_no_line = false;
			/*(this.chart.plotArea.charts[0].getObjectType() === AscDFH.historyitem_type_ScatterChart &&
             (this.chart.plotArea.charts[0].scatterStyle === AscFormat.SCATTER_STYLE_MARKER || this.chart.plotArea.charts[0].scatterStyle === AscFormat.SCATTER_STYLE_NONE));  */
			this.legendLength = null;

			var oFirstChart = aCharts[0];
			var bNoPieChart = (oFirstChart.getObjectType() !== AscDFH.historyitem_type_PieChart && oFirstChart.getObjectType() !== AscDFH.historyitem_type_DoughnutChart);
			var bSurfaceChart = (oFirstChart.getObjectType() === AscDFH.historyitem_type_SurfaceChart);

			var bSeriesLegend = aCharts.length > 1 || (bNoPieChart && (!(oFirstChart.varyColors && series.length === 1) || bSurfaceChart));
			if (bSeriesLegend) {
				if (bSurfaceChart) {
					this.legendLength = this.chart.plotArea.charts[0].compiledBandFormats.length;
					ser = series[0];
				} else {
					this.legendLength = series.length;
				}
				for(i = 0; i < this.legendLength; ++i) {
					if(!bSurfaceChart) {
						ser = series[i];

						if(ser.isHidden)
							continue;
						entry = legend.findLegendEntryByIndex(i);
						if(entry && entry.bDelete)
							continue;
						arr_str_labels.push(ser.getSeriesName());
					}
					else {
						entry = legend.findLegendEntryByIndex(i);
						if(entry && entry.bDelete)
							continue;
						var oBandFmt = this.chart.plotArea.charts[0].compiledBandFormats[i];
						arr_str_labels.push(oBandFmt.startValue + "-" + oBandFmt.endValue);
					}
					calc_entry = new AscFormat.CalcLegendEntry(legend, this, i);
					calc_entry.series = ser;
					calc_entry.txBody = AscFormat.CreateTextBodyFromString(arr_str_labels[arr_str_labels.length - 1], this.getDrawingDocument(), calc_entry);

					//if(entry)
					//    calc_entry.txPr = entry.txPr;

					/*if(calc_entryes[0])
					 {
					 calc_entry.lastStyleObject = calc_entryes[0].lastStyleObject;
					 }*/
					calc_entryes.push(calc_entry);
					cur_width = calc_entry.txBody.getRectWidth(2000);
					if(cur_width > max_width)
						max_width = cur_width;

					cur_font_size = calc_entry.txBody.content.Content[0].CompiledPr.Pr.TextPr.FontSize;
					if(cur_font_size > max_font_size)
						max_font_size = cur_font_size;

					calc_entry.calcMarkerUnion = new AscFormat.CUnionMarker();
					union_marker = calc_entry.calcMarkerUnion;
					var pts = ser.getNumPts();
					var nSerType = ser.getObjectType();
					if(nSerType === AscDFH.historyitem_type_BarSeries ||
						nSerType === AscDFH.historyitem_type_BubbleSeries ||
						nSerType === AscDFH.historyitem_type_AreaSeries ||
						nSerType === AscDFH.historyitem_type_PieSeries ||
						(nSerType === AscDFH.historyitem_type_RadarSeries &&
							ser.parent && ser.parent.isFilled())) {

						union_marker.marker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_SQUARE);
						union_marker.marker.pen = ser.compiledSeriesPen;
						union_marker.marker.brush = ser.compiledSeriesBrush;
					}
					else if(nSerType === AscDFH.historyitem_type_SurfaceSeries) {
						var oBandFmt = this.chart.plotArea.charts[0].compiledBandFormats[i];
						union_marker.marker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_SQUARE);
						union_marker.marker.pen = oBandFmt.spPr.ln;
						union_marker.marker.brush = oBandFmt.spPr.Fill;
					}
					else if(nSerType === AscDFH.historyitem_type_LineSeries ||
						nSerType === AscDFH.historyitem_type_ScatterSer ||
						nSerType === AscDFH.historyitem_type_RadarSeries) {

						if(AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this)) {
							union_marker.marker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_SQUARE);
							union_marker.marker.pen = ser.compiledSeriesPen;
							union_marker.marker.brush = ser.compiledSeriesBrush;
						}
						else {
							if(ser.compiledSeriesMarker) {
								var pts = ser.getNumPts();
								union_marker.marker = AscFormat.CreateMarkerGeometryByType(ser.compiledSeriesMarker.symbol);
								if(pts[0] && pts[0].compiledMarker) {
									union_marker.marker.brush = pts[0].compiledMarker.brush;
									union_marker.marker.pen = pts[0].compiledMarker.pen;

								}
							}
							if(ser.compiledSeriesPen && !b_scatter_no_line) {
								union_marker.lineMarker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_DASH);
								union_marker.lineMarker.pen = ser.compiledSeriesPen.createDuplicate(); //Копируем, так как потом возможно придется изменять толщину линии;
							}
							if(!b_scatter_no_line && !AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this)) {
								b_line_series = true;
							}
						}
					}
					if(union_marker.marker) {
						union_marker.marker.pen && union_marker.marker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						union_marker.marker.brush && union_marker.marker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
					}
					union_marker.lineMarker && union_marker.lineMarker.pen && union_marker.lineMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
			} else {
				ser = series[0];
				i = 1;
				while (ser && ser.isHidden) {
					ser = series[i];
					++i;
				}
				ser = ser || series[0];
				this.legendLength = 0;
				if (ser) {
					var pts = ser.getNumPts(), pt;
					var oLitForLegend;
					if (ser && ser.cat) {
						oLitForLegend = ser.cat.getLit();
					}
					var oLitFormat = null, oPtFormat = null;
					if (oLitForLegend && typeof oLitForLegend.formatCode === "string" && oLitForLegend.formatCode.length > 0) {
						oLitFormat = oNumFormatCache.get(oLitForLegend.formatCode);
					}
					this.legendLength = pts.length;

					var oNumLit = ser.getNumLit();
					var nEndIndex = pts.length;
					if (oNumLit) {
						nEndIndex = oNumLit.ptCount;
					}
					for (i = 0; i < nEndIndex; ++i) {
						entry = legend.findLegendEntryByIndex(i);
						if (entry && entry.bDelete)
							continue;
						if (oNumLit) {
							pt = oNumLit.getPtByIndex(i);
						} else {
							pt = pts[i];
						}
						var str_pt = oLitForLegend ? oLitForLegend.getPtByIndex(i) : null;
						if (str_pt) {
							var sPt;
							if (typeof str_pt.formatCode === "string" && str_pt.formatCode.length > 0) {
								oPtFormat = oNumFormatCache.get(str_pt.formatCode);
								if (oPtFormat) {
									sPt = oPtFormat.formatToChart(str_pt.val);
								} else {
									sPt = str_pt.val + "";
								}
							} else if (oLitFormat) {
								sPt = oLitFormat.formatToChart(str_pt.val);
							} else {
								sPt = str_pt.val + "";
							}
							arr_str_labels.push(sPt);
						} else {
							arr_str_labels.push((pt ? (pt.idx + 1) : "") + "");
						}

						calc_entry = new AscFormat.CalcLegendEntry(legend, this, i);
						calc_entry.txBody = AscFormat.CreateTextBodyFromString(arr_str_labels[arr_str_labels.length - 1], this.getDrawingDocument(), calc_entry);

						//if(entry)
						//    calc_entry.txPr = entry.txPr;
						//if(calc_entryes[0])
						//{
						//    calc_entry.lastStyleObject = calc_entryes[0].lastStyleObject;
						//}
						calc_entryes.push(calc_entry);

						cur_width = calc_entry.txBody.getRectWidth(2000);
						if (cur_width > max_width)
							max_width = cur_width;

						cur_font_size = calc_entry.txBody.content.Content[0].CompiledPr.Pr.TextPr.FontSize;
						if (cur_font_size > max_font_size)
							max_font_size = cur_font_size;

						calc_entry.calcMarkerUnion = new AscFormat.CUnionMarker();
						union_marker = calc_entry.calcMarkerUnion;
						if (ser.getObjectType() === AscDFH.historyitem_type_LineSeries && !AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this) ||
							ser.getObjectType() === AscDFH.historyitem_type_ScatterSer ||
							ser.getObjectType() === AscDFH.historyitem_type_RadarSeries) {
							if (pt) {
								if (pt.compiledMarker) {
									union_marker.marker = AscFormat.CreateMarkerGeometryByType(pt.compiledMarker.symbol);
									union_marker.marker.brush = pt.compiledMarker.pen.Fill;
									union_marker.marker.pen = pt.compiledMarker.pen;
								}
								if (pt.pen) {
									union_marker.lineMarker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_DASH);
									union_marker.lineMarker.pen = pt.pen;
								}
							}
							if (!b_scatter_no_line && !AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this))
								b_line_series = true;
						} else {
							b_no_line_series = false;
							union_marker.marker = AscFormat.CreateMarkerGeometryByType(AscFormat.SYMBOL_SQUARE);
							if (pt) {
								union_marker.marker.pen = pt.pen;
								union_marker.marker.brush = pt.brush;
							} else {
								var style = AscFormat.CHART_STYLE_MANAGER.getStyleByIndex(this.style);
								var base_fills = AscFormat.getArrayFillsFromBase(style.fill2, nEndIndex);
								union_marker.marker.brush = base_fills[i];
							}
						}
						if (union_marker.marker) {
							union_marker.marker.pen && union_marker.marker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							union_marker.marker.brush && union_marker.marker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
						union_marker.lineMarker && union_marker.lineMarker.pen && union_marker.lineMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
					}
				}
			}
			var marker_size;
			var distance_to_text;
			var line_marker_width;
			if (b_line_series) {
				marker_size = 2.5;
				line_marker_width = 7.7;//Пока так
				for (i = 0; i < calc_entryes.length; ++i) {
					calc_entry = calc_entryes[i];
					if (calc_entry.calcMarkerUnion.lineMarker) {
						calc_entry.calcMarkerUnion.lineMarker.spPr.geometry.Recalculate(line_marker_width, 1);
						/*Excel не дает сделать толщину линии для маркера легенды больше определенной. Считаем, что это толщина равна 133000emu*/
						if (calc_entry.calcMarkerUnion.lineMarker.pen
							&& AscFormat.isRealNumber(calc_entry.calcMarkerUnion.lineMarker.pen.w) && calc_entry.calcMarkerUnion.lineMarker.pen.w > 133000) {
							calc_entry.calcMarkerUnion.lineMarker.pen.w = 133000;
						}
						calc_entry.calcMarkerUnion.lineMarker.penWidth = calc_entry.calcMarkerUnion.lineMarker.pen && AscFormat.isRealNumber(calc_entry.calcMarkerUnion.lineMarker.pen.w) ? calc_entry.calcMarkerUnion.lineMarker.pen.w / 36000 : 0;
					}
					if (calc_entryes[i].calcMarkerUnion.marker) {
						var marker_width = marker_size;
						if (calc_entryes[i].series) {
							if (calc_entryes[i].series.getObjectType() !== AscDFH.historyitem_type_LineSeries &&
								calc_entryes[i].series.getObjectType() !== AscDFH.historyitem_type_ScatterSer &&
								calc_entryes[i].series.getObjectType() !== AscDFH.historyitem_type_RadarSeries) {
								marker_width = line_marker_width;
							}
						}
						calc_entryes[i].calcMarkerUnion.marker.spPr.geometry.Recalculate(marker_width, marker_size);
						calc_entryes[i].calcMarkerUnion.marker.extX = marker_width;
						calc_entryes[i].calcMarkerUnion.marker.extY = marker_size;
					}
				}
				distance_to_text = 0.5;
			} else {
				marker_size = 0.2 * max_font_size;
				for (i = 0; i < calc_entryes.length; ++i) {
					calc_entryes[i].calcMarkerUnion.marker.spPr.geometry.Recalculate(marker_size, marker_size);
				}
				distance_to_text = marker_size * 0.7;
			}
			var left_inset = marker_size + 3 * distance_to_text;

			var legend_width, legend_height;
			var fFixedWidth = null, fFixedHeight = null;
			var bFixedSize = false;
			if (legend.layout) {
				fFixedWidth = this.calculateSizeByLayout(0, this.extX, legend.layout.w, legend.layout.wMode);
				fFixedHeight = this.calculateSizeByLayout(0, this.extY, legend.layout.h, legend.layout.hMode);
				bFixedSize = AscFormat.isRealNumber(fFixedWidth) && fFixedWidth > 0 && AscFormat.isRealNumber(fFixedHeight) && fFixedHeight > 0;
				if (bFixedSize) {
					var oOldLayout = legend.layout;
					legend.layout = null;
					calc_entryes = [].concat(calc_entryes);
					this.recalculateLegend();
					legend.calcEntryes = calc_entryes;

					legend.naturalWidth = legend.extX;
					legend.naturalHeight = legend.extY;
					legend.layout = oOldLayout;
					var bResetLegendPos = false;
					if (!AscFormat.isRealNumber(this.chart.legend.legendPos)) {
						bResetLegendPos = true;
						this.chart.legend.legendPos = Asc.c_oAscChartLegendShowSettings.bottom;
					}

					this.checkPrecalculateChartObject();
					var pos = this.chartObj.recalculatePositionText(this.chart.legend);
					if (this.chart.legend.layout) {
						if (AscFormat.isRealNumber(legend.layout.x)) {
							pos.x = this.calculatePosByLayout(pos.x, legend.layout.xMode, legend.layout.x, this.chart.legend.extX, this.extX);
						}
						if (AscFormat.isRealNumber(legend.layout.y)) {
							pos.y = this.calculatePosByLayout(pos.y, legend.layout.yMode, legend.layout.y, this.chart.legend.extY, this.extY);
						}
					}
					if (bResetLegendPos) {
						this.chart.legend.legendPos = null;
					}

					fFixedWidth = this.calculateSizeByLayout(pos.x, this.extX, legend.layout.w, legend.layout.wMode);
					fFixedHeight = this.calculateSizeByLayout(pos.y, this.extY, legend.layout.h, legend.layout.hMode);

				}
			}
			if (AscFormat.isRealNumber(legend_pos)) {
				var max_legend_width, max_legend_height;
				var cut_index;
				if ((legend_pos === c_oAscChartLegendShowSettings.left ||
					legend_pos === c_oAscChartLegendShowSettings.leftOverlay ||
					legend_pos === c_oAscChartLegendShowSettings.right ||
					legend_pos === c_oAscChartLegendShowSettings.rightOverlay || legend_pos === c_oAscChartLegendShowSettings.topRight) && !bFixedSize) {
					max_legend_width = this.extX / 3;//Считаем, что ширина легенды не больше трети ширины всей диаграммы;
					var sizes = this.getChartSizes();
					max_legend_height = sizes.h;
					if (b_line_series) {
						left_inset = line_marker_width + 3 * distance_to_text;
						var content_width = max_legend_width - left_inset;
						if (content_width <= 0)
							content_width = 0.01;
						var cur_content_width, max_content_width = 0;
						var arr_heights = [];
						var arr_heights2 = [];
						for (i = 0; i < calc_entryes.length; ++i) {
							calc_entry = calc_entryes[i];
							cur_content_width = calc_entry.txBody.getMaxContentWidth(content_width, true);
							if (cur_content_width > max_content_width)
								max_content_width = cur_content_width;
							arr_heights.push(calc_entry.txBody.getSummaryHeight());
							arr_heights2.push(calc_entry.txBody.getSummaryHeight3());
						}
						if (max_content_width < max_legend_width - left_inset) {
							legend_width = max_content_width + left_inset;
						} else {
							legend_width = max_legend_width;
						}

						var max_entry_height2 = Math.max(0, Math.max.apply(Math, arr_heights));
						for (i = 0; i < arr_heights.length; ++i)
							arr_heights[i] = max_entry_height2;
						var max_entry_height2 = Math.max(0, Math.max.apply(Math, arr_heights2));
						for (i = 0; i < arr_heights2.length; ++i)
							arr_heights2[i] = max_entry_height2;

						var height_summ = 0;
						var height_summ2 = 0;
						for (i = 0; i < arr_heights2.length; ++i) {
							height_summ += arr_heights2[i];
							height_summ2 += arr_heights[i];
							if (height_summ > max_legend_height) {
								cut_index = i;
								break;
							}
						}
						if (AscFormat.isRealNumber(cut_index)) {
							if (cut_index > 0) {
								legend_height = height_summ2 - arr_heights[cut_index];
							} else {
								legend_height = max_legend_height;
							}
						} else {
							cut_index = arr_heights.length;
							legend_height = height_summ2;
						}
						legend.x = 0;
						legend.y = 0;
						legend.extX = legend_width;
						legend.extY = legend_height;
						var summ_h = 0;
						calc_entryes.splice(cut_index, calc_entryes.length - cut_index);
						for (i = 0; i < cut_index && i < calc_entryes.length; ++i) {
							calc_entry = calc_entryes[i];
							this.layoutLegendEntry(calc_entry, distance_to_text, summ_h, distance_to_text);
							summ_h += arr_heights[i];
						}
						legend.setPosition(0, 0);
					} else {
						var content_width = max_legend_width - left_inset;
						if (content_width <= 0)
							content_width = 0.01;
						var cur_content_width, max_content_width = 0;
						var arr_heights = [];
						var arr_heights2 = [];
						for (i = 0; i < calc_entryes.length; ++i) {
							calc_entry = calc_entryes[i];
							cur_content_width = calc_entry.txBody.getMaxContentWidth(content_width, true);
							if (cur_content_width > max_content_width)
								max_content_width = cur_content_width;
							arr_heights.push(calc_entry.txBody.getSummaryHeight());
							arr_heights2.push(calc_entry.txBody.getSummaryHeight3());
						}


						var chart_object;
						if (this.chart && this.chart.plotArea && this.chart.plotArea.charts[0]) {
							chart_object = this.chart.plotArea.charts[0];
						}
						var b_reverse_order = false;
						if (chart_object && chart_object.getObjectType() === AscDFH.historyitem_type_BarChart && chart_object.barDir === AscFormat.BAR_DIR_BAR &&
							(cat_ax && cat_ax.scaling && AscFormat.isRealNumber(cat_ax.scaling.orientation) ? cat_ax.scaling.orientation : AscFormat.ORIENTATION_MIN_MAX) === AscFormat.ORIENTATION_MIN_MAX
							|| chart_object && chart_object.getObjectType() === AscDFH.historyitem_type_SurfaceChart) {
							b_reverse_order = true;
						}

						var max_entry_height2 = Math.max(0, Math.max.apply(Math, arr_heights));
						for (i = 0; i < arr_heights.length; ++i)
							arr_heights[i] = max_entry_height2;
						var max_entry_height2 = Math.max(0, Math.max.apply(Math, arr_heights2));
						for (i = 0; i < arr_heights2.length; ++i)
							arr_heights2[i] = max_entry_height2;
						if (max_content_width < max_legend_width - left_inset) {
							legend_width = max_content_width + left_inset;
						} else {
							legend_width = max_legend_width;
						}
						var height_summ = 0;
						var height_summ2 = 0;
						for (i = 0; i < arr_heights.length; ++i) {
							height_summ += arr_heights2[i];
							height_summ2 += arr_heights[i];
							if (height_summ > max_legend_height) {
								cut_index = i;
								break;
							}
						}


						if (AscFormat.isRealNumber(cut_index)) {
							if (cut_index > 0) {
								legend_height = height_summ2 - arr_heights[cut_index];
							} else {
								legend_height = max_legend_height;
							}
						} else {
							cut_index = arr_heights.length;
							legend_height = height_summ2;
						}
						legend.x = 0;
						legend.y = 0;
						legend.extX = legend_width;
						legend.extY = legend_height;
						var summ_h = 0;


						var b_reverse_order = false;
						var chart_object, cat_ax, start_index, end_index;
						if (this.chart && this.chart.plotArea && this.chart.plotArea.charts[0]) {
							chart_object = this.chart.plotArea.charts[0];
							if (chart_object && chart_object.getAxisByTypes) {
								var axis_by_types = chart_object.getAxisByTypes();
								cat_ax = axis_by_types.catAx[0];
							}
						}
						if (chart_object && chart_object.getObjectType() === AscDFH.historyitem_type_BarChart && chart_object.barDir === AscFormat.BAR_DIR_BAR &&
							(cat_ax && cat_ax.scaling && AscFormat.isRealNumber(cat_ax.scaling.orientation) ? cat_ax.scaling.orientation : AscFormat.ORIENTATION_MIN_MAX) === AscFormat.ORIENTATION_MIN_MAX) {
							b_reverse_order = true;
						}


						if (!b_reverse_order) {
							calc_entryes.splice(cut_index, calc_entryes.length - cut_index);
							for (i = 0; i < cut_index && i < calc_entryes.length; ++i) {
								calc_entry = calc_entryes[i];
								this.layoutLegendEntry(calc_entry, distance_to_text, summ_h, distance_to_text);
								summ_h += arr_heights[i];
							}
						} else {
							calc_entryes.splice(0, calc_entryes.length - cut_index);
							for (i = calc_entryes.length - 1; i > -1; --i) {
								calc_entry = calc_entryes[i];
								this.layoutLegendEntry(calc_entry, distance_to_text, summ_h, distance_to_text);
								summ_h += arr_heights[i];
							}
						}

						legend.setPosition(0, 0);
					}
				} else {
					/*пока сделаем так: максимальная ширимна 0.9 от ширины дмаграммы
                     без заголовка  максимальная высота легенды 0.6 от высоты диаграммы,
                     с заголовком 0.6 от высоты за вычетом высоты заголовка*/
					var fMaxLegendHeight, fMaxLegendWidth;
					if (bFixedSize) {
						fMaxLegendWidth = fFixedWidth;
						fMaxLegendHeight = fFixedHeight;
					} else {
						fMaxLegendWidth = 0.9 * this.extX;
						fMaxLegendHeight = (this.extY - (this.chart.title ? this.chart.title.extY : 0)) * 0.6;
					}
					var fMarkerWidth;
					if (b_line_series) {
						fMarkerWidth = line_marker_width;
					} else {
						fMarkerWidth = marker_size;
					}
					var oUnionMarker;
					var fMaxEntryWidth = 0.0;
					var fSummWidth = 0.0;
					var fCurEntryWidth, oCurEntry;
					var fLegendWidth, fLegendHeight;
					var aHeights = [];
					var aHeights2 = [];
					var fMaxEntryHeight = 0.0;
					var fCurEntryHeight = 0.0;
					var fCurEntryHeight2 = 0.0;
					var aWidths = [];
					var fCurPosX;
					var fCurPosY;
					var fDistanceBetweenLabels;
					var fVertDistanceBetweenLabels;
					var oLineMarker, fPenWidth, oMarker;
					for (i = 0; i < calc_entryes.length; ++i) {
						oCurEntry = calc_entryes[i];
						var fWidth = oCurEntry.txBody.getMaxContentWidth(20000, true);
						aWidths.push(fWidth);
						fCurEntryWidth = distance_to_text + fMarkerWidth + distance_to_text + fWidth;
						if (fMaxEntryWidth < fCurEntryWidth) {
							fMaxEntryWidth = fCurEntryWidth;
						}
						fCurEntryHeight = oCurEntry.txBody.getSummaryHeight();
						fCurEntryHeight2 = oCurEntry.txBody.getSummaryHeight3();
						aHeights.push(fCurEntryHeight);
						aHeights2.push(fCurEntryHeight2);
						if (fMaxEntryHeight < fCurEntryHeight2) {
							fMaxEntryHeight = fCurEntryHeight2;
						}
						fSummWidth += fCurEntryWidth;
					}
					if (fSummWidth <= fMaxLegendWidth) {
						if (bFixedSize) {
							fLegendWidth = fFixedWidth;
							fLegendHeight = fFixedHeight;
						} else {
							fLegendWidth = fSummWidth;
							fLegendHeight = Math.max.apply(Math, aHeights);
						}
						fDistanceBetweenLabels = (fLegendWidth - fSummWidth) / (calc_entryes.length + 1);
						fCurPosX = 0.0;
						legend.extX = fLegendWidth;
						legend.extY = fLegendHeight;
						for (i = 0; i < calc_entryes.length; ++i) {
							fCurPosX += fDistanceBetweenLabels;
							oCurEntry = calc_entryes[i];
							this.layoutLegendEntry(oCurEntry, fCurPosX + distance_to_text, Math.max(0, fLegendHeight / 2.0 - aHeights[i] / 2.0), distance_to_text);
							fCurPosX += distance_to_text + fMarkerWidth + distance_to_text + aWidths[i];
						}

						legend.setPosition(0, 0);
					} else {
						if (fMaxLegendWidth < fMaxEntryWidth) {
							var fTextWidth = Math.max(1.0, fMaxLegendWidth - distance_to_text - fMarkerWidth - distance_to_text);
							fMaxEntryWidth = 0.0;
							fMaxEntryHeight = 0.0;
							aWidths.length = 0;
							aHeights.length = 0;
							for (i = 0; i < calc_entryes.length; ++i) {
								oCurEntry = calc_entryes[i];
								var fWidth = oCurEntry.txBody.getMaxContentWidth(fTextWidth, true);
								aWidths.push(fWidth);
								fCurEntryWidth = distance_to_text + fMarkerWidth + distance_to_text + fWidth;
								if (fMaxEntryWidth < fCurEntryWidth) {
									fMaxEntryWidth = fCurEntryWidth;
								}
								fCurEntryHeight = oCurEntry.txBody.getSummaryHeight();
								fCurEntryHeight2 = oCurEntry.txBody.getSummaryHeight3();
								aHeights.push(fCurEntryHeight);
								aHeights2.push(fCurEntryHeight2);
								if (fMaxEntryHeight < fCurEntryHeight2) {
									fMaxEntryHeight = fCurEntryHeight2;
								}
								fSummWidth += fCurEntryWidth;
							}
						}
						var nColsCount = Math.max(1, (fMaxLegendWidth / fMaxEntryWidth) >> 0);
						var nMaxRowsCount = Math.max(1, (fMaxLegendHeight / fMaxEntryHeight) >> 0);
						var nMaxEntriesCount = nColsCount * nMaxRowsCount;
						if (calc_entryes.length > nMaxEntriesCount) {
							calc_entryes.splice(nMaxEntriesCount, calc_entryes.length - nMaxEntriesCount);
							aWidths.splice(nMaxEntriesCount, aWidths.length - nMaxEntriesCount);
							aHeights.splice(nMaxEntriesCount, aHeights.length - nMaxEntriesCount);
							aHeights2.splice(nMaxEntriesCount, aHeights2.length - nMaxEntriesCount);
							fMaxEntryWidth = distance_to_text + fMarkerWidth + distance_to_text + Math.max.apply(Math, aWidths);
							fMaxEntryHeight = Math.max.apply(Math, aHeights2);
						}
						var nRowsCount = Math.ceil(calc_entryes.length / nColsCount);
						if (bFixedSize) {
							fLegendWidth = fFixedWidth;
							fLegendHeight = fFixedHeight;
						} else {
							fLegendWidth = fMaxEntryWidth * nColsCount;
							fLegendHeight = nRowsCount * Math.max.apply(Math, aHeights);
						}
						fDistanceBetweenLabels = (fLegendWidth - nColsCount * fMaxEntryWidth) / nColsCount;
						fVertDistanceBetweenLabels = (fLegendHeight - nRowsCount * Math.max.apply(Math, aHeights)) / nRowsCount;
						legend.extX = fLegendWidth;
						legend.extY = fLegendHeight;
						fCurPosY = fVertDistanceBetweenLabels / 2.0;
						for (i = 0; i < nRowsCount; ++i) {
							fCurPosX = fDistanceBetweenLabels / 2.0;
							for (j = 0; j < nColsCount && (i * nColsCount + j) < calc_entryes.length; ++j) {
								var nEntryIndex = i * nColsCount + j;
								oCurEntry = calc_entryes[nEntryIndex];
								this.layoutLegendEntry(oCurEntry, fCurPosX + distance_to_text, Math.max(0, fCurPosY + Math.max.apply(Math, aHeights) / 2.0 - aHeights[nEntryIndex] / 2.0), distance_to_text);
								fCurPosX += (fMaxEntryWidth + fDistanceBetweenLabels);
							}
							fCurPosY += (Math.max.apply(Math, aHeights) + fVertDistanceBetweenLabels);
						}
						legend.setPosition(0, 0);
					}
				}
			} else {
				//TODO
			}
			legend.recalcInfo = {recalculateLine: true, recalculateFill: true, recalculateTransparent: true};
			legend.recalculatePen();
			legend.recalculateBrush();
			legend.calcGeometry = null;
			if (legend.spPr) {
				legend.calcGeometry = legend.spPr.geometry;
			}
			if (legend.calcGeometry) {
				legend.calcGeometry.Recalculate(legend.extX, legend.extY);
			}
			for (i = 0; i < calc_entryes.length; ++i) {
				calc_entryes[i].checkWidhtContent();
			}
		}
	};
	CChartSpace.prototype.internalCalculatePenBrushFloorWall = function (oSide, nSideType) {
		if (!oSide) {
			return;
		}

		var parent_objects = this.getParentObjects();
		if (oSide.spPr && oSide.spPr.ln) {
			oSide.pen = oSide.spPr.ln.createDuplicate();
		} else {
			var oCompiledPen = null;
			if (this.style >= 1 && this.style <= 40 && 2 === nSideType) {
				if (parent_objects.theme && parent_objects.theme.themeElements
					&& parent_objects.theme.themeElements.fmtScheme
					&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst
					&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0]) {
					oCompiledPen = parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0].createDuplicate();
					if (this.style >= 1 && this.style <= 32) {
						oCompiledPen.Fill = CreateUnifillSolidFillSchemeColor(15, 0.75);
					} else {
						oCompiledPen.Fill = CreateUnifillSolidFillSchemeColor(8, 0.75);
					}
				}
			}
			oSide.pen = oCompiledPen;
		}
		if (this.style >= 1 && this.style <= 32) {
			if (oSide.spPr && oSide.spPr.Fill) {
				oSide.brush = oSide.spPr.Fill.createDuplicate();
				if (nSideType === 0 || nSideType === 2) {
					var cColorMod = new AscFormat.CColorMod;
					if (nSideType === 2)
						cColorMod.val = 45000;
					else
						cColorMod.val = 35000;
					cColorMod.name = "shade";
					oSide.brush.addColorMod(cColorMod);
				}
			} else {
				oSide.brush = null;
			}
		} else {
			var oSubtleFill;
			if (parent_objects.theme && parent_objects.theme.themeElements
				&& parent_objects.theme.themeElements.fmtScheme
				&& parent_objects.theme.themeElements.fmtScheme.fillStyleLst) {
				oSubtleFill = parent_objects.theme.themeElements.fmtScheme.fillStyleLst[0];
			}

			var oDefaultBrush;
			var tint = 0.20000;
			if (this.style >= 33 && this.style <= 34)
				oDefaultBrush = CreateUnifillSolidFillSchemeColor(8, 0.20000);
			else if (this.style >= 35 && this.style <= 40)
				oDefaultBrush = CreateUnifillSolidFillSchemeColor(this.style - 35, tint);
			else
				oDefaultBrush = CreateUnifillSolidFillSchemeColor(8, 0.95000);

			if (oSide.spPr) {
				oDefaultBrush.merge(oSide.spPr.Fill);
			}

			if (nSideType === 0 || nSideType === 2) {
				var cColorMod = new AscFormat.CColorMod;
				if (nSideType === 0)
					cColorMod.val = 45000;
				else
					cColorMod.val = 35000;
				cColorMod.name = "shade";
				oDefaultBrush.addColorMod(cColorMod);
			}
			oSide.brush = oDefaultBrush;
		}

		if (oSide.brush) {
			oSide.brush.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, {
				R: 0,
				G: 0,
				B: 0,
				A: 255
			}, this.clrMapOvr);
		}
		if (oSide.pen) {
			oSide.pen.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, {
				R: 0,
				G: 0,
				B: 0,
				A: 255
			}, this.clrMapOvr);
		}
	};
	CChartSpace.prototype.recalculateWalls = function () {
		if (this.chart) {
			this.internalCalculatePenBrushFloorWall(this.chart.sideWall, 0);
			this.internalCalculatePenBrushFloorWall(this.chart.backWall, 1);
			this.internalCalculatePenBrushFloorWall(this.chart.floor, 2);
		}
	};
	CChartSpace.prototype.recalculateUpDownBars = function () {
		if (this.chart && this.chart.plotArea) {
			var aCharts = this.chart.plotArea.charts;
			for (var t = 0; t < aCharts.length; ++t) {
				var oChart = aCharts[t];
				if (oChart && oChart.upDownBars) {
					var bars = oChart.upDownBars;
					var up_bars = bars.upBars;
					var down_bars = bars.downBars;
					var parents = this.getParentObjects();
					bars.upBarsBrush = null;
					bars.upBarsPen = null;
					bars.downBarsBrush = null;
					bars.downBarsPen = null;
					if (up_bars || down_bars) {
						var default_bar_line = new AscFormat.CLn();
						if (parents.theme && parents.theme.themeElements
							&& parents.theme.themeElements.fmtScheme
							&& parents.theme.themeElements.fmtScheme.lnStyleLst) {
							default_bar_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
						}
						if (this.style >= 1 && this.style <= 16)
							default_bar_line.setFill(CreateUnifillSolidFillSchemeColor(15, 0));
						else if (this.style >= 17 && this.style <= 32 ||
							this.style >= 41 && this.style <= 48)
							default_bar_line = CreateNoFillLine();
						else if (this.style === 33 || this.style === 34)
							default_bar_line.setFill(CreateUnifillSolidFillSchemeColor(8, 0));
						else if (this.style >= 35 && this.style <= 40)
							default_bar_line.setFill(CreateUnifillSolidFillSchemeColor(this.style - 35, -0.25000));
					}
					if (up_bars) {
						var default_up_bars_fill;
						if (this.style === 1 || this.style === 9 || this.style === 17 || this.style === 25 || this.style === 41) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(8, 0.25000);
						} else if (this.style === 2 || this.style === 10 || this.style === 18 || this.style === 26) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(8, 0.05000);
						} else if (this.style >= 3 && this.style <= 8) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 3, 0.25000);
						} else if (this.style >= 11 && this.style <= 16) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 11, 0.25000);
						} else if (this.style >= 19 && this.style <= 24) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 19, 0.25000);
						} else if (this.style >= 27 && this.style <= 32) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 27, 0.25000);
						} else if (this.style >= 33 && this.style <= 40 || this.style === 42) {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(12, 0);
						} else {
							default_up_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 43, 0.25000);
						}
						if (up_bars.Fill) {
							default_up_bars_fill.merge(up_bars.Fill);
						}
						default_up_bars_fill.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
							R: 0,
							G: 0,
							B: 0,
							A: 255
						}, this.clrMapOvr);
						oChart.upDownBars.upBarsBrush = default_up_bars_fill;
						var up_bars_line = default_bar_line.createDuplicate();
						if (up_bars.ln)
							up_bars_line.merge(up_bars.ln);
						up_bars_line.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
							R: 0,
							G: 0,
							B: 0,
							A: 255
						}, this.clrMapOvr);
						oChart.upDownBars.upBarsPen = up_bars_line;

					}
					if (down_bars) {
						var default_down_bars_fill;
						if (this.style === 1 || this.style === 9 || this.style === 17 || this.style === 25 || this.style === 41 || this.style === 33) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(8, 0.85000);
						} else if (this.style === 2 || this.style === 10 || this.style === 18 || this.style === 26 || this.style === 34) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(8, 0.95000);
						} else if (this.style >= 3 && this.style <= 8) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 3, -0.25000);
						} else if (this.style >= 11 && this.style <= 16) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 11, -0.25000);
						} else if (this.style >= 19 && this.style <= 24) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 19, -0.25000);
						} else if (this.style >= 27 && this.style <= 32) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 27, -0.25000);
						} else if (this.style >= 35 && this.style <= 40) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 35, -0.25000);
						} else if (this.style === 42) {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(8, 0);
						} else {
							default_down_bars_fill = CreateUnifillSolidFillSchemeColor(this.style - 43, -0.25000);
						}
						if (down_bars.Fill) {
							default_down_bars_fill.merge(down_bars.Fill);
						}
						default_down_bars_fill.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
							R: 0,
							G: 0,
							B: 0,
							A: 255
						}, this.clrMapOvr);
						oChart.upDownBars.downBarsBrush = default_down_bars_fill;
						var down_bars_line = default_bar_line.createDuplicate();
						if (down_bars.ln)
							down_bars_line.merge(down_bars.ln);
						down_bars_line.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
							R: 0,
							G: 0,
							B: 0,
							A: 255
						}, this.clrMapOvr);
						oChart.upDownBars.downBarsPen = down_bars_line;
					}
				}
			}
		}
	};
	CChartSpace.prototype.recalculatePlotAreaChartPen = function () {
		if (this.chart && this.chart.plotArea) {
			if (this.chart.plotArea.spPr && this.chart.plotArea.spPr.ln) {
				this.chart.plotArea.pen = this.chart.plotArea.spPr.ln;
				var parents = this.getParentObjects();
				this.chart.plotArea.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
					R: 0,
					G: 0,
					B: 0,
					A: 255
				}, this.clrMapOvr);
			} else {
				this.chart.plotArea.pen = null;
			}
		}
	};
	CChartSpace.prototype.recalculatePenBrush = function () {
		var parents = this.getParentObjects(), RGBA = {R: 0, G: 0, B: 0, A: 255};
		if (this.brush) {
			this.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
		}
		if (this.pen) {
			this.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
			checkBlackUnifill(this.pen.Fill, true);
		}
		if (this.chart) {
			if (this.chart.title) {
				if (this.chart.title.brush) {
					this.chart.title.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.title.pen) {
					this.chart.title.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
			}
			if (this.chart.plotArea) {
				if (this.chart.plotArea.brush) {
					this.chart.plotArea.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.plotArea.pen) {
					this.chart.plotArea.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				for (var t = 0; t < this.chart.plotArea.axId.length; ++t) {
					var oAxis = this.chart.plotArea.axId[t];
					if (oAxis.compiledTickMarkLn) {
						oAxis.compiledTickMarkLn.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						checkBlackUnifill(oAxis.compiledTickMarkLn.Fill, true);
					}
					if (oAxis.compiledMajorGridLines) {
						oAxis.compiledMajorGridLines.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						checkBlackUnifill(oAxis.compiledMajorGridLines.Fill, true);
					}
					if (oAxis.compiledMinorGridLines) {
						oAxis.compiledMinorGridLines.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						checkBlackUnifill(oAxis.compiledMinorGridLines.Fill, true);
					}
					if (oAxis.title) {
						if (oAxis.title.brush) {
							oAxis.title.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
						if (oAxis.title.pen) {
							oAxis.title.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
					}
				}

				for (t = 0; t < this.chart.plotArea.charts.length; ++t) {
					var oChart = this.chart.plotArea.charts[t];
					var series = oChart.series;
					for (var i = 0; i < series.length; ++i) {
						var pts = series[i].getNumPts();
						var oCalcObjects = [], nCalcId = 0;
						for (var j = 0; j < pts.length; ++j) {
							var pt = pts[j];

							if (pt.brush) {
								if (!oCalcObjects[pt.brush.calcId]) {
									pt.brush.calcId = ++nCalcId;
									oCalcObjects[pt.brush.calcId] = pt.brush;
									pt.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
								}

							}
							if (pt.pen) {
								if (!oCalcObjects[pt.pen.calcId]) {
									pt.pen.calcId = ++nCalcId;
									oCalcObjects[pt.pen.calcId] = pt.pen;
									pt.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
								}
							}
							if (pt.compiledMarker) {
								if (pt.compiledMarker.brush) {
									if (!oCalcObjects[pt.compiledMarker.brush.calcId]) {
										pt.compiledMarker.brush.calcId = ++nCalcId;
										oCalcObjects[pt.compiledMarker.brush.calcId] = pt.compiledMarker.brush;
										pt.compiledMarker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
									}
								}
								if (pt.compiledMarker.pen) {
									if (!oCalcObjects[pt.compiledMarker.pen.calcId]) {
										pt.compiledMarker.pen.calcId = ++nCalcId;
										oCalcObjects[pt.compiledMarker.pen.calcId] = pt.compiledMarker.pen;
										pt.compiledMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
									}
								}
							}
							if (pt.compiledDlb) {
								if (pt.compiledDlb.brush) {
									if (!oCalcObjects[pt.compiledDlb.brush.calcId]) {
										pt.compiledDlb.brush.calcId = ++nCalcId;
										oCalcObjects[pt.compiledDlb.brush.calcId] = pt.compiledDlb.brush;
										pt.compiledDlb.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
									}
								}
								if (pt.compiledDlb.pen) {
									if (!oCalcObjects[pt.compiledDlb.pen.calcId]) {
										pt.compiledDlb.pen.calcId = ++nCalcId;
										oCalcObjects[pt.compiledDlb.pen.calcId] = pt.compiledDlb.pen;
										pt.compiledDlb.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
									}
								}

							}
						}
					}
					if (oChart.calculatedHiLowLines) {
						oChart.calculatedHiLowLines.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
					}
					if (oChart.upDownBars) {
						if (oChart.upDownBars.upBarsBrush) {
							oChart.upDownBars.upBarsBrush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
						if (oChart.upDownBars.upBarsPen) {
							oChart.upDownBars.upBarsPen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
						if (oChart.upDownBars.downBarsBrush) {
							oChart.upDownBars.downBarsBrush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
						if (oChart.upDownBars.downBarsPen) {
							oChart.upDownBars.downBarsPen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
						}
					}
				}
			}
			if (this.chart.legend) {
				if (this.chart.legend.brush) {
					this.chart.legend.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.legend.pen) {
					this.chart.legend.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}

				var legend = this.chart.legend;
				for (var i = 0; i < legend.calcEntryes.length; ++i) {
					var union_marker = legend.calcEntryes[i].calcMarkerUnion;
					if (union_marker) {
						if (union_marker.marker) {
							if (union_marker.marker.pen) {
								union_marker.marker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
							if (union_marker.marker.brush) {
								union_marker.marker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
						}
						if (union_marker.lineMarker) {
							if (union_marker.lineMarker.pen) {
								union_marker.lineMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
							if (union_marker.lineMarker.brush) {
								union_marker.lineMarker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
						}
					}
				}
			}
			if (this.chart.floor) {
				if (this.chart.floor.brush) {
					this.chart.floor.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.floor.pen) {
					this.chart.floor.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
			}
			if (this.chart.sideWall) {
				if (this.chart.sideWall.brush) {
					this.chart.sideWall.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.sideWall.pen) {
					this.chart.sideWall.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
			}
			if (this.chart.backWall) {
				if (this.chart.backWall.brush) {
					this.chart.backWall.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
				if (this.chart.backWall.pen) {
					this.chart.backWall.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
				}
			}
		}
	};
	CChartSpace.prototype.getCopyWithSourceFormatting = function (oIdMap) {
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.idMap = oIdMap;
		oPr.drawingDocument = this.getDrawingDocument();
		var oCopy = this.copy(oPr);
		oCopy.updateLinks();
		if (oIdMap) {
			oIdMap[this.Id] = oCopy.Id;
		}
		var oTheme = this.Get_Theme();
		var oColorMap = this.Get_ColorMap();


		fSaveChartObjectSourceFormatting(this, oCopy, oTheme, oColorMap);

		if (!oCopy.txPr || !oCopy.txPr.content || !oCopy.txPr.content.Content[0] || !oCopy.txPr.content.Content[0].Pr) {
			oCopy.setTxPr(AscFormat.CreateTextBodyFromString("", this.getDrawingDocument(), oCopy));
		}
		var bMerge = false;
		var oTextPr = new CTextPr();
		if (!oCopy.txPr.content.Content[0].Pr.DefaultRunPr || !oCopy.txPr.content.Content[0].Pr.DefaultRunPr.RFonts || !oCopy.txPr.content.Content[0].Pr.DefaultRunPr.RFonts.Ascii
			|| !oCopy.txPr.content.Content[0].Pr.DefaultRunPr.RFonts.Ascii.Name) {
			bMerge = true;
			oTextPr.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);
		}

		if (this.txPr && this.txPr.content && this.txPr.content.Content[0] && this.txPr.content.Content[0].Pr.DefaultRunPr && this.txPr.content.Content[0].Pr.DefaultRunPr.Unifill) {
			bMerge = true;
			var oUnifill = this.txPr.content.Content[0].Pr.DefaultRunPr.Unifill.createDuplicate();
			oUnifill.check(this.Get_Theme(), this.Get_ColorMap());
			oTextPr.Unifill = oUnifill.saveSourceFormatting();
		} else if (!AscFormat.isRealNumber(this.style) || this.style < 33) {
			bMerge = true;
			var oUnifill = CreateUnifillSolidFillSchemeColor(15, 0.0);
			oUnifill.check(this.Get_Theme(), this.Get_ColorMap());
			oTextPr.Unifill = oUnifill.saveSourceFormatting();
		} else {
			bMerge = true;
			var oUnifill = CreateUnifillSolidFillSchemeColor(6, 0.0);
			oUnifill.check(this.Get_Theme(), this.Get_ColorMap());
			oTextPr.Unifill = oUnifill.saveSourceFormatting();
		}

		if (bMerge) {

			var oParaPr = oCopy.txPr.content.Content[0].Pr.Copy();
			var oParaPr2 = new CParaPr();
			var oCopyTextPr = oTextPr.Copy();
			oParaPr2.DefaultRunPr = oCopyTextPr;
			oParaPr.Merge(oParaPr2);
			oCopy.txPr.content.Content[0].Set_Pr(oParaPr);
			//CheckObjectTextPr(oCopy, oTextPr, this.getDrawingDocument());
			// fSaveChartObjectSourceFormatting(oCopy, oCopy, oTheme, oColorMap);
		}
		if (this.chart) {
			if (this.chart.title) {
				fSaveChartObjectSourceFormatting(this.chart.title, oCopy.chart.title, oTheme, oColorMap);
			}
			if (this.chart.plotArea) {
				fSaveChartObjectSourceFormatting(this.chart.plotArea, oCopy.chart.plotArea, oTheme, oColorMap);
				var aAxes = oCopy.chart.plotArea.axId;
				if (aAxes) {
					for (var i = 0; i < aAxes.length; ++i) {
						if (aAxes[i] && this.chart.plotArea.axId[i]) {
							fSaveChartObjectSourceFormatting(this.chart.plotArea.axId[i], aAxes[i], oTheme, oColorMap);
							if (this.chart.plotArea.axId[i].compiledLn) {
								if (!aAxes[i].spPr) {
									aAxes[i].setSpPr(new AscFormat.CSpPr());
								}
								aAxes[i].spPr.setLn(this.chart.plotArea.axId[i].compiledLn.createDuplicate(true));
							}
							if (this.chart.plotArea.axId[i].compiledMajorGridLines) {
								if (!aAxes[i].majorGridlines) {
									aAxes[i].setMajorGridlines(new AscFormat.CSpPr());
								}
								aAxes[i].majorGridlines.setLn(this.chart.plotArea.axId[i].compiledMajorGridLines.createDuplicate(true));
							}
							if (this.chart.plotArea.axId[i].compiledMinorGridLines) {
								if (!aAxes[i].minorGridlines) {
									aAxes[i].setMinorGridlines(new AscFormat.CSpPr());
								}
								aAxes[i].minorGridlines.setLn(this.chart.plotArea.axId[i].compiledMinorGridLines.createDuplicate(true));
							}
							if (aAxes[i].title) {
								fSaveChartObjectSourceFormatting(this.chart.plotArea.axId[i].title, aAxes[i].title, oTheme, oColorMap);
							}
						}
					}
				}

				for (var t = 0; t < this.chart.plotArea.charts.length; ++t) {
					if (this.chart.plotArea.charts[t]) {
						var oChartOrig = this.chart.plotArea.charts[t];
						var oChartCopy = oCopy.chart.plotArea.charts[t];
						var series = oChartOrig.series;
						var seriesCopy = oChartCopy.series;

						var oDataPoint;
						for (var i = 0; i < series.length; ++i) {
							series[i].brush = series[i].compiledSeriesBrush;
							series[i].pen = series[i].compiledSeriesPen;
							fSaveChartObjectSourceFormatting(series[i], seriesCopy[i], oTheme, oColorMap);
							var pts = series[i].getNumPts();
							for (var j = 0; j < pts.length; ++j) {
								var pt = pts[j];
								oDataPoint = null;
								if (Array.isArray(seriesCopy[i].dPt)) {
									for (var k = 0; k < seriesCopy[i].dPt.length; ++k) {
										if (seriesCopy[i].dPt[k].idx === pts[j].idx) {
											oDataPoint = seriesCopy[i].dPt[k];
											break;
										}
									}
								}
								if (!oDataPoint) {
									oDataPoint = new AscFormat.CDPt();
									oDataPoint.setIdx(pt.idx);
									seriesCopy[i].addDPt(oDataPoint);
								}
								fSaveChartObjectSourceFormatting(pt, oDataPoint, oTheme, oColorMap);
								if (pt.compiledMarker) {
									var oMarker = pt.compiledMarker.createDuplicate();
									oDataPoint.setMarker(oMarker);
									fSaveChartObjectSourceFormatting(pt.compiledMarker, oMarker, oTheme, oColorMap);
								}
							}
						}
						if (oChartOrig.calculatedHiLowLines) {
							if (!oChartCopy.hiLowLines) {
								oChartCopy.setHiLowLines(new AscFormat.CSpPr());
							}
							oChartCopy.hiLowLines.setLn(oChartOrig.calculatedHiLowLines.createDuplicate(true));
						}
						if (oChartOrig.upDownBars) {
							if (oChartOrig.upDownBars.upBarsBrush) {
								if (!oChartCopy.upDownBars.upBars) {
									oChartCopy.upDownBars.setUpBars(new AscFormat.CSpPr());
								}
								oChartCopy.upDownBars.upBars.setFill(oChartOrig.upDownBars.upBarsBrush.saveSourceFormatting());
							}
							if (oChartOrig.upDownBars.upBarsPen) {
								if (!oChartCopy.upDownBars.upBars) {
									oChartCopy.upDownBars.setUpBars(new AscFormat.CSpPr());
								}
								oChartCopy.upDownBars.upBars.setLn(oChartOrig.upDownBars.upBarsPen.createDuplicate(true));
							}
							if (oChartOrig.upDownBars.downBarsBrush) {
								if (!oChartCopy.upDownBars.downBars) {
									oChartCopy.upDownBars.setDownBars(new AscFormat.CSpPr());
								}
								oChartCopy.upDownBars.downBars.setFill(oChartOrig.upDownBars.downBarsBrush.saveSourceFormatting());
							}
							if (oChartOrig.upDownBars.downBarsPen) {
								if (!oChartCopy.upDownBars.downBars) {
									oChartCopy.upDownBars.setDownBars(new AscFormat.CSpPr());
								}
								oChartCopy.upDownBars.downBars.setLn(oChartOrig.upDownBars.downBarsPen.createDuplicate(true));
							}
						}
					}
				}
			}
			if (this.chart.legend) {
				fSaveChartObjectSourceFormatting(this.chart.legend, oCopy.chart.legend, oTheme, oColorMap);
				var legend = this.chart.legend;
				for (var i = 0; i < legend.legendEntryes.length; ++i) {
					fSaveChartObjectSourceFormatting(legend.legendEntryes[i], oCopy.chart.legend.legendEntryes[i], oTheme, oColorMap);
				}
			}
			if (this.chart.floor) {
				fSaveChartObjectSourceFormatting(this.chart.floor, oCopy.chart.floor, oTheme, oColorMap);
			}
			if (this.chart.sideWall) {

				fSaveChartObjectSourceFormatting(this.chart.sideWall, oCopy.chart.sideWall, oTheme, oColorMap);
			}
			if (this.chart.backWall) {
				fSaveChartObjectSourceFormatting(this.chart.backWall, oCopy.chart.backWall, oTheme, oColorMap);
			}
		}
		return oCopy;
	};
	CChartSpace.prototype.getChartSizes = function (bNotRecalculate) {
		if (this.plotAreaRect && !this.recalcInfo.recalculateAxisVal) {
			return {
				startX: this.plotAreaRect.x,
				startY: this.plotAreaRect.y,
				w: this.plotAreaRect.w,
				h: this.plotAreaRect.h
			};
		}
		if (!this.chartObj)
			this.chartObj = new AscFormat.CChartsDrawer();
		var oChartSize = this.chartObj.calculateSizePlotArea(this, bNotRecalculate);
		var oLayout = this.chart.plotArea.layout;
		if (oLayout) {

			oChartSize.startX = this.calculatePosByLayout(oChartSize.startX, oLayout.xMode, oLayout.x, oChartSize.w, this.extX);
			oChartSize.startY = this.calculatePosByLayout(oChartSize.startY, oLayout.yMode, oLayout.y, oChartSize.h, this.extY);
			var fSize = this.calculateSizeByLayout(oChartSize.startX, this.extX, oLayout.w, oLayout.wMode);
			if (AscFormat.isRealNumber(fSize) && fSize > 0) {
				var fSize2 = this.calculateSizeByLayout(oChartSize.startY, this.extY, oLayout.h, oLayout.hMode);
				if (AscFormat.isRealNumber(fSize2) && fSize2 > 0) {
					oChartSize.w = fSize;
					oChartSize.h = fSize2;
					oChartSize.startX = this.calculatePosByLayout(oChartSize.startX, oLayout.xMode, oLayout.x, oChartSize.w, this.extX);
					oChartSize.startY = this.calculatePosByLayout(oChartSize.startY, oLayout.yMode, oLayout.y, oChartSize.h, this.extY);
					var aCharts = this.chart.plotArea.charts;
					for (var i = 0; i < aCharts.length; ++i) {
						var nChartType = aCharts[i].getObjectType();
						if (nChartType === AscDFH.historyitem_type_PieChart || nChartType === AscDFH.historyitem_type_DoughnutChart) {
							var fCX = oChartSize.startX + oChartSize.w / 2.0;
							var fCY = oChartSize.startY + oChartSize.h / 2.0;
							var fPieSize = Math.min(oChartSize.w, oChartSize.h);
							oChartSize.startX = fCX - fPieSize / 2.0;
							oChartSize.startY = fCY - fPieSize / 2.0;
							oChartSize.w = fPieSize;
							oChartSize.h = fPieSize;
							break;
						}
					}
				}
			}
		}
		if (oChartSize.w <= 0) {
			oChartSize.w = 1;
		}
		if (oChartSize.h <= 0) {
			oChartSize.h = 1;
		}

		return oChartSize;
	};
	CChartSpace.prototype.reindexSeries = function () {
		var aAllSeries = this.getAllSeries();
		for (var nIndex = 0; nIndex < aAllSeries.length; ++nIndex) {
			aAllSeries[nIndex].setIdx(nIndex);
		}
		this.reorderSeries();
	};
	CChartSpace.prototype.reorderSeries = function () {
		var aAllSeries = this.getAllSeries();
		aAllSeries.sort(function (a, b) {
			return a.order - b.order;
		});
		var oSeries;
		for (var nIndex = 0; nIndex < aAllSeries.length; ++nIndex) {
			oSeries = aAllSeries[nIndex];
			oSeries.setOrder(nIndex);
		}
	};
	CChartSpace.prototype.sortSeries = function () {
		if (this.chart && this.chart.plotArea) {
			var aCharts = this.chart.plotArea.charts;
			for (var i = 0; i < aCharts.length; ++i) {
				aCharts[i].sortSeries();
			}
		}
	};
	CChartSpace.prototype.moveSeriesUp = function (oSeries) {
		var aAllSeries = this.getAllSeries();
		aAllSeries.sort(function (a, b) {
			return a.order - b.order;
		});
		for (var i = 0; i < aAllSeries.length; ++i) {
			if (oSeries === aAllSeries[i]) {
				if (i > 0) {
					var nOrder = oSeries.order;
					oSeries.setOrder(aAllSeries[i - 1].order);
					aAllSeries[i - 1].setOrder(nOrder);
				}
				break;
			}
		}
		this.sortSeries();
	};
	CChartSpace.prototype.moveSeriesDown = function (oSeries) {
		var aAllSeries = this.getAllSeries();
		aAllSeries.sort(function (a, b) {
			return a.order - b.order;
		});
		for (var i = 0; i < aAllSeries.length; ++i) {
			if (oSeries === aAllSeries[i]) {
				if (i < aAllSeries.length - 1) {
					var nOrder = oSeries.order;
					oSeries.setOrder(aAllSeries[i + 1].order);
					aAllSeries[i + 1].setOrder(nOrder);
				}
				break;
			}
		}
		this.sortSeries();
	};
	CChartSpace.prototype.getAllSeries = function () {
		if (this.chart && this.chart.plotArea) {
			return this.chart.plotArea.getAllSeries();
		}
		return [];
	};
	CChartSpace.prototype.recalculatePlotAreaChartBrush = function () {
		if (this.chart && this.chart.plotArea) {
			var plot_area = this.chart.plotArea;
			var default_brush;
			if (AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this)) {
				default_brush = CreateNoFillUniFill();
			} else {
				if (this.chart.plotArea && this.chart.plotArea.charts.length === 1 && this.chart.plotArea.charts[0] &&
					(this.chart.plotArea.charts[0].getObjectType() === AscDFH.historyitem_type_PieChart
						|| this.chart.plotArea.charts[0].getObjectType() === AscDFH.historyitem_type_DoughnutChart)) {
					default_brush = CreateNoFillUniFill();
				} else {
					var tint = 0.20000;
					if (this.style >= 1 && this.style <= 32)
						default_brush = CreateUnifillSolidFillSchemeColor(6, tint);
					else if (this.style >= 33 && this.style <= 34)
						default_brush = CreateUnifillSolidFillSchemeColor(8, 0.20000);
					else if (this.style >= 35 && this.style <= 40)
						default_brush = CreateUnifillSolidFillSchemeColor(this.style - 35, tint);
					else
						default_brush = CreateUnifillSolidFillSchemeColor(8, 0.95000);
				}
			}
			if (plot_area.spPr && plot_area.spPr.Fill) {
				default_brush.merge(plot_area.spPr.Fill);
			}
			var parents = this.getParentObjects();
			default_brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
				R: 0,
				G: 0,
				B: 0,
				A: 255
			}, this.clrMapOvr);
			plot_area.brush = default_brush;
		}
	};
	CChartSpace.prototype.recalculateChartPen = function () {
		var parent_objects = this.getParentObjects();
		var default_line = new AscFormat.CLn();
		if (parent_objects.theme && parent_objects.theme.themeElements
			&& parent_objects.theme.themeElements.fmtScheme
			&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst) {
			default_line.merge(parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0]);
		}

		var fill;
		if (this.style >= 1 && this.style <= 32)
			fill = CreateUnifillSolidFillSchemeColor(15, 0.75000);
		else if (this.style >= 33 && this.style <= 40)
			fill = CreateUnifillSolidFillSchemeColor(8, 0.75000);
		else
			fill = CreateUnifillSolidFillSchemeColor(12, 0);
		default_line.setFill(fill);
		if (this.spPr && this.spPr.ln)
			default_line.merge(this.spPr.ln);
		var parents = this.getParentObjects();
		default_line.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
			R: 0,
			G: 0,
			B: 0,
			A: 255
		}, this.clrMapOvr);
		this.pen = default_line;
		checkBlackUnifill(this.pen.Fill, true);
	};
	CChartSpace.prototype.recalculateChartBrush = function () {
		var default_brush;
		if (this.style >= 1 && this.style <= 32)
			default_brush = CreateUnifillSolidFillSchemeColor(6, 0);
		else if (this.style >= 33 && this.style <= 40)
			default_brush = CreateUnifillSolidFillSchemeColor(12, 0);
		else
			default_brush = CreateUnifillSolidFillSchemeColor(8, 0);

		if (this.spPr && this.spPr.Fill) {
			default_brush.merge(this.spPr.Fill);
		}
		var parents = this.getParentObjects();
		default_brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
			R: 0,
			G: 0,
			B: 0,
			A: 255
		}, this.clrMapOvr);
		this.brush = default_brush;

	};
	CChartSpace.prototype.recalculateAxisTickMark = function () {
		if (this.chart && this.chart.plotArea) {
			var oThis = this;
			var calcMajorMinorGridLines = function (axis, defaultStyle, subtleLine, parents) {
				function calcGridLine(defaultStyle, spPr, subtleLine, parents) {
					var compiled_grid_lines = new AscFormat.CLn();
					compiled_grid_lines.merge(subtleLine);
					// if(compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color && compiled_grid_lines.Fill.fill.color.Mods)
					// {
					//     compiled_grid_lines.Fill.fill.color.Mods.Mods.length = 0;
					// }
					if (!compiled_grid_lines.Fill) {
						compiled_grid_lines.setFill(new AscFormat.CUniFill());
					}
					//if(compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color && compiled_grid_lines.Fill.fill.color.Mods)
					//{
					//    compiled_grid_lines.Fill.fill.color.Mods.Mods.length = 0;
					//}
					compiled_grid_lines.Fill.merge(defaultStyle);

					if (subtleLine && subtleLine.Fill && subtleLine.Fill.fill && subtleLine.Fill.fill.color && subtleLine.Fill.fill.color.Mods
						&& compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color) {
						compiled_grid_lines.Fill.fill.color.Mods = subtleLine.Fill.fill.color.Mods.createDuplicate();
					}

					compiled_grid_lines.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
						R: 0,
						G: 0,
						B: 0,
						A: 255
					}, oThis.clrMapOvr);
					checkBlackUnifill(compiled_grid_lines.Fill, true);
					if (spPr && spPr.ln) {
						compiled_grid_lines.merge(spPr.ln);
						compiled_grid_lines.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
							R: 0,
							G: 0,
							B: 0,
							A: 255
						}, oThis.clrMapOvr);
					}

					return compiled_grid_lines;
				}

				axis.compiledLn = calcGridLine(defaultStyle.axisAndMajorGridLines, axis.spPr, subtleLine, parents);
				axis.compiledTickMarkLn = axis.compiledLn.createDuplicate();
				axis.compiledTickMarkLn.headEnd = null;
				axis.compiledTickMarkLn.tailEnd = null;
				axis.compiledTickMarkLn.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
					R: 0,
					G: 0,
					B: 0,
					A: 255
				}, oThis.clrMapOvr);
				checkBlackUnifill(axis.compiledTickMarkLn.Fill, true);
			};
			var default_style = CHART_STYLE_MANAGER.getDefaultLineStyleByIndex(this.style);
			var parent_objects = this.getParentObjects();
			var subtle_line;
			if (parent_objects.theme && parent_objects.theme.themeElements
				&& parent_objects.theme.themeElements.fmtScheme
				&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst) {
				subtle_line = parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0];
			}
			var aAxes = this.chart.plotArea.axId;
			for (var i = 0; i < aAxes.length; ++i) {
				calcMajorMinorGridLines(aAxes[i], default_style, subtle_line, parent_objects);
			}
		}
	};
	CChartSpace.prototype.getXValAxisValues = function () {
		this.checkPrecalculateChartObject();
		return [].concat(this.chart.plotArea.catAx.scale)
	};
	CChartSpace.prototype.getValAxisValues = function () {
		this.checkPrecalculateChartObject();
		return [].concat(this.chart.plotArea.valAx.scale);
	};
	CChartSpace.prototype.checkPrecalculateChartObject = function () {
		if (!this.chartObj || this.recalcInfo.recalculateChart) {
			if (!this.chartObj) {
				this.chartObj = new AscFormat.CChartsDrawer();
			}
			this.chartObj.preCalculateData(this);
		}
	};
	CChartSpace.prototype.recalculateDLbls = function () {
		if (this.chart && this.chart.plotArea) {
			this.cachedCanvas = null;
			var aCharts = this.chart.plotArea.charts;
			for (var t = 0; t < aCharts.length; ++t) {

				var series = aCharts[t].series;
				var nDefaultPosition;
				if (aCharts[t].getDefaultDataLabelsPosition) {
					nDefaultPosition = aCharts[t].getDefaultDataLabelsPosition();
				}

				var default_lbl = new AscFormat.CDLbl();
				default_lbl.initDefault(nDefaultPosition);
				var bSkip = false;
				if (this.ptsCount > MAX_LABELS_COUNT) {
					bSkip = true;
				}
				var nCount = 0;
				var nLblCount = 0;
				for (var i = 0; i < series.length; ++i) {
					var ser = series[i];
					var pts = ser.getNumPts();


					var series_dlb = new AscFormat.CDLbl();
					series_dlb.merge(default_lbl);
					series_dlb.merge(aCharts[t].dLbls);
					series_dlb.merge(ser.dLbls);


					var bCheckAll = true, j;
					if (series_dlb.checkNoLbl()) {
						if (!aCharts[t].dLbls || aCharts[t].dLbls.dLbl.length === 0
							&& !ser.dLbls || ser.dLbls.dLbl.length === 0) {
							for (j = 0; j < pts.length; ++j) {
								pts[j].compiledDlb = null;
							}
							bCheckAll = false;
						}
					}
					if (bCheckAll) {
						for (j = 0; j < pts.length; ++j) {
							var pt = pts[j];
							if (bSkip) {

								if (nLblCount > (MAX_LABELS_COUNT * (nCount / this.ptsCount))) {
									pt.compiledDlb = null;
									nCount++;
									continue;
								}

							}
							var compiled_dlb = new AscFormat.CDLbl();
							compiled_dlb.merge(default_lbl);
							compiled_dlb.merge(aCharts[t].dLbls);
							if (aCharts[t].dLbls)
								compiled_dlb.merge(aCharts[t].dLbls.findDLblByIdx(pt.idx), false);
							compiled_dlb.merge(ser.dLbls);
							if (ser.dLbls) {
								var oSeriesDLbl = ser.dLbls.findDLblByIdx(pt.idx);
								if (oSeriesDLbl) {
									oSeriesDLbl.chart = this;
									compiled_dlb.merge(oSeriesDLbl);
									if (oSeriesDLbl.tx) {
										compiled_dlb.tx = oSeriesDLbl.tx;
									}
								}
							}

							if (compiled_dlb.checkNoLbl()) {
								pt.compiledDlb = null;
							} else {
								pt.compiledDlb = compiled_dlb;
								pt.compiledDlb.chart = this;
								pt.compiledDlb.series = ser;
								pt.compiledDlb.pt = pt;
								pt.compiledDlb.recalculate();
								nLblCount++;
							}
							++nCount;
						}
					}
				}
			}
		}
	};
	CChartSpace.prototype.recalculateHiLowLines = function () {
		if (this.chart && this.chart.plotArea) {
			var aCharts = this.chart.plotArea.charts;
			var parents = this.getParentObjects();
			for (var i = 0; i < aCharts.length; ++i) {
				var oCurChart = aCharts[i];
				if ((oCurChart instanceof AscFormat.CStockChart || oCurChart instanceof AscFormat.CLineChart) && oCurChart.hiLowLines) {
					var default_line = parents.theme.themeElements.fmtScheme.lnStyleLst[0].createDuplicate();
					if (this.style >= 1 && this.style <= 32)
						default_line.setFill(CreateUnifillSolidFillSchemeColor(15, 0));
					else if (this.style >= 33 && this.style <= 34)
						default_line.setFill(CreateUnifillSolidFillSchemeColor(8, 0));
					else if (this.style >= 35 && this.style <= 40)
						default_line.setFill(CreateUnifillSolidFillSchemeColor(8, -0.25000));
					else
						default_line.setFill(CreateUnifillSolidFillSchemeColor(12, 0));
					default_line.merge(oCurChart.hiLowLines.ln);
					oCurChart.calculatedHiLowLines = default_line;
					default_line.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
						R: 0,
						G: 0,
						B: 0,
						A: 255
					}, this.clrMapOvr);
				} else {
					oCurChart.calculatedHiLowLines = null;
				}
			}
		}
	};
	CChartSpace.prototype.recalculateSeriesColors = function () {
		this.cachedCanvas = null;
		this.ptsCount = 0;
		if (this.chart && this.chart.plotArea) {
			var style = CHART_STYLE_MANAGER.getStyleByIndex(this.style);
			var parents = this.getParentObjects();
			var RGBA = {R: 0, G: 0, B: 0, A: 255};
			var aCharts = this.chart.plotArea.charts;
			var aAllSeries = [];
			for (var t = 0; t < aCharts.length; ++t) {
				aAllSeries = aAllSeries.concat(aCharts[t].series);
			}
			var nMaxSeriesIdx = getMaxIdx(aAllSeries);
			for (t = 0; t < aCharts.length; ++t) {
				var oChart = aCharts[t];
				var series = oChart.series;
				if (oChart.varyColors
					&& (series.length === 1 || oChart.getObjectType() === AscDFH.historyitem_type_PieChart || oChart.getObjectType() === AscDFH.historyitem_type_DoughnutChart)) {
					var base_fills2 = getArrayFillsFromBase(style.fill2, getMaxIdx(series));
					for (var ii = 0; ii < series.length; ++ii) {
						var ser = series[ii];
						var pts = ser.getNumPts();
						this.ptsCount += pts.length;

						ser.compiledSeriesBrush = new AscFormat.CUniFill();
						ser.compiledSeriesBrush.merge(base_fills2[ser.idx]);
						if (ser.spPr && ser.spPr.Fill) {
							ser.compiledSeriesBrush.merge(ser.spPr.Fill);
						}


						ser.compiledSeriesPen = new AscFormat.CLn();
						if (style.line1 === EFFECT_NONE) {
							ser.compiledSeriesPen.w = 0;
						} else if (style.line1 === EFFECT_SUBTLE) {
							ser.compiledSeriesPen.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
						} else if (style.line1 === EFFECT_MODERATE) {
							ser.compiledSeriesPen.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[1]);
						} else if (style.line1 === EFFECT_INTENSE) {
							ser.compiledSeriesPen.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[2]);
						}
						var base_line_fills2;
						if (this.style === 34)
							base_line_fills2 = getArrayFillsFromBase(style.line2, getMaxIdx(series));

						ser.compiledSeriesPen.Fill = new AscFormat.CUniFill();
						if (this.style !== 34) {
							ser.compiledSeriesPen.Fill.merge(style.line2[0]);
						} else {
							ser.compiledSeriesPen.Fill.merge(base_line_fills2[ser.idx]);
						}
						if (ser.spPr && ser.spPr.ln)
							ser.compiledSeriesPen.merge(ser.spPr.ln);


						if (!(oChart.getObjectType() === AscDFH.historyitem_type_LineChart || oChart.getObjectType() === AscDFH.historyitem_type_ScatterChart)) {
							var base_fills = getArrayFillsFromBase(style.fill2, getMaxIdx(pts));
							for (var i = 0; i < pts.length; ++i) {
								var compiled_brush = new AscFormat.CUniFill();
								compiled_brush.merge(base_fills[pts[i].idx]);
								if (ser.spPr && ser.spPr.Fill) {
									compiled_brush.merge(ser.spPr.Fill);
								}
								if (Array.isArray(ser.dPt)) {
									for (var j = 0; j < ser.dPt.length; ++j) {
										if (ser.dPt[j].idx === pts[i].idx) {
											if (ser.dPt[j].spPr) {
												compiled_brush.merge(ser.dPt[j].spPr.Fill);
											}
											break;
										}
									}
								}
								pts[i].brush = compiled_brush;
								pts[i].brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}

							default_line = new AscFormat.CLn();
							if (style.line1 === EFFECT_NONE) {
								default_line.w = 0;
							} else if (style.line1 === EFFECT_SUBTLE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
							} else if (style.line1 === EFFECT_MODERATE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[1]);
							} else if (style.line1 === EFFECT_INTENSE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[2]);
							}
							var base_line_fills;
							if (this.style === 34)
								base_line_fills = getArrayFillsFromBase(style.line2, getMaxIdx(pts));
							for (i = 0; i < pts.length; ++i) {
								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								compiled_line.Fill = new AscFormat.CUniFill();
								if (this.style !== 34) {
									compiled_line.Fill.merge(style.line2[0]);
								} else {
									compiled_line.Fill.merge(base_line_fills[pts[i].idx]);
								}
								if (ser.spPr && ser.spPr.ln)
									compiled_line.merge(ser.spPr.ln);
								if (Array.isArray(ser.dPt) && !(ser.getObjectType && ser.getObjectType() === AscDFH.historyitem_type_AreaSeries)) {
									for (var j = 0; j < ser.dPt.length; ++j) {
										if (ser.dPt[j].idx === pts[i].idx) {
											if (ser.dPt[j].spPr) {
												compiled_line.merge(ser.dPt[j].spPr.ln);
											}
											break;
										}
									}
								}
								pts[i].pen = compiled_line;
								pts[i].pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
						} else {
							var default_line;
							if (oChart.getObjectType() === AscDFH.historyitem_type_ScatterChart && oChart.scatterStyle === AscFormat.SCATTER_STYLE_MARKER || oChart.scatterStyle === AscFormat.SCATTER_STYLE_NONE) {
								default_line = new AscFormat.CLn();
								default_line.setFill(new AscFormat.CUniFill());
								default_line.Fill.setFill(new AscFormat.CNoFill());
							} else {
								default_line = parents.theme.themeElements.fmtScheme.lnStyleLst[0];
							}
							var base_line_fills = getArrayFillsFromBase(style.line4, getMaxIdx(pts));
							for (var i = 0; i < pts.length; ++i) {
								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								if (!(oChart.getObjectType() === AscDFH.historyitem_type_ScatterChart && oChart.scatterStyle === AscFormat.SCATTER_STYLE_MARKER || oChart.scatterStyle === AscFormat.SCATTER_STYLE_NONE))
									compiled_line.Fill.merge(base_line_fills[pts[i].idx]);
								compiled_line.w *= style.line3;
								if (ser.spPr && ser.spPr.ln) {
									compiled_line.merge(ser.spPr.ln);
								}
								if (Array.isArray(ser.dPt)) {
									for (var j = 0; j < ser.dPt.length; ++j) {
										if (ser.dPt[j].idx === pts[i].idx) {
											if (ser.dPt[j].spPr) {
												compiled_line.merge(ser.dPt[j].spPr.ln);
											}
											break;
										}
									}
								}
								pts[i].brush = null;
								pts[i].pen = compiled_line;
								pts[i].pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
							}
						}
						for (var j = 0; j < pts.length; ++j) {
							if (pts[j].compiledMarker) {
								pts[j].compiledMarker.pen && pts[j].compiledMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
								pts[j].compiledMarker.brush && pts[j].compiledMarker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);

							}
						}

					}
				}
				else {
					var nChartType = oChart.getObjectType();
					if(nChartType === AscDFH.historyitem_type_LineChart ||
						(nChartType === AscDFH.historyitem_type_RadarChart && (oChart.radarStyle === AscFormat.RADAR_STYLE_STANDARD || oChart.radarStyle === AscFormat.RADAR_STYLE_MARKER))) {
						var base_line_fills = getArrayFillsFromBase(style.line4, nMaxSeriesIdx);
						if(!AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this)) {
							for(var i = 0; i < series.length; ++i) {
								var default_line = parents.theme.themeElements.fmtScheme.lnStyleLst[0];
								var ser = series[i];
								var pts = ser.getNumPts();
								this.ptsCount += pts.length;
								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								compiled_line.Fill && compiled_line.Fill.merge(base_line_fills[ser.idx]);
								compiled_line.w *= style.line3;
								if(ser.spPr && ser.spPr.ln)
									compiled_line.merge(ser.spPr.ln);
								ser.compiledSeriesPen = compiled_line;
								for(var j = 0; j < pts.length; ++j) {
									var oPointLine = null;
									if(Array.isArray(ser.dPt)) {
										for(var k = 0; k < ser.dPt.length; ++k) {
											if(ser.dPt[k].idx === pts[j].idx) {
												if(ser.dPt[k].spPr) {
													oPointLine = ser.dPt[k].spPr.ln;
												}
												break;
											}
										}
									}
									if(oPointLine) {
										compiled_line = ser.compiledSeriesPen.createDuplicate();
										compiled_line.merge(oPointLine);
									}
									else {
										compiled_line = ser.compiledSeriesPen;
									}
									pts[j].brush = null;
									pts[j].pen = compiled_line;
								}
							}
						}
						else {
							var base_fills = getArrayFillsFromBase(style.fill2, nMaxSeriesIdx);
							var base_line_fills = null;
							if(style.line1 === EFFECT_SUBTLE && this.style === 34)
								base_line_fills = getArrayFillsFromBase(style.line2, nMaxSeriesIdx);
							for(var i = 0; i < series.length; ++i) {
								var ser = series[i];
								var compiled_brush = new AscFormat.CUniFill();
								compiled_brush.merge(base_fills[ser.idx]);
								if(ser.spPr && ser.spPr.Fill) {
									compiled_brush.merge(ser.spPr.Fill);
								}
								ser.compiledSeriesBrush = compiled_brush;
								ser.compiledSeriesBrush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
								var pts = ser.getNumPts();
								for(var j = 0; j < pts.length; ++j) {
									pts[j].brush = ser.compiledSeriesBrush;
									if(Array.isArray(ser.dPt) && !(ser.getObjectType && ser.getObjectType() === AscDFH.historyitem_type_AreaSeries)) {
										for(var k = 0; k < ser.dPt.length; ++k) {
											if(ser.dPt[k].idx === pts[j].idx) {
												if(ser.dPt[k].spPr) {
													if(ser.dPt[k].spPr.Fill && ser.dPt[k].spPr.Fill.fill) {
														pts[j].brush = ser.dPt[k].spPr.Fill.createDuplicate();
														pts[j].brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
													}
												}
												break;
											}
										}
									}
								}


								//
								{
									default_line = new AscFormat.CLn();
									if(style.line1 === EFFECT_NONE) {
										default_line.w = 0;
									}
									else if(style.line1 === EFFECT_SUBTLE) {
										default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
									}
									else if(style.line1 === EFFECT_MODERATE) {
										default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[1]);
									}
									else if(style.line1 === EFFECT_INTENSE) {
										default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[2]);
									}
									var base_line_fills;
									if(this.style === 34)
										base_line_fills = getArrayFillsFromBase(style.line2, getMaxIdx(pts));


									var compiled_line = new AscFormat.CLn();
									compiled_line.merge(default_line);
									compiled_line.Fill = new AscFormat.CUniFill();
									if(this.style !== 34)
										compiled_line.Fill.merge(style.line2[0]);
									else
										compiled_line.Fill.merge(base_line_fills[ser.idx]);
									if(ser.spPr && ser.spPr.ln) {
										compiled_line.merge(ser.spPr.ln);
									}
									ser.compiledSeriesPen = compiled_line;
									ser.compiledSeriesPen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
									for(var j = 0; j < pts.length; ++j) {
										if(Array.isArray(ser.dPt) && !(ser.getObjectType && ser.getObjectType() === AscDFH.historyitem_type_AreaSeries)) {
											for(var k = 0; k < ser.dPt.length; ++k) {
												if(ser.dPt[k].idx === pts[j].idx) {
													if(ser.dPt[k].spPr) {
														compiled_line = ser.compiledSeriesPen.createDuplicate();
														compiled_line.merge(ser.dPt[k].spPr.ln);
														pts[j].pen = compiled_line;
														pts[j].pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
													}
													break;
												}
											}
										}
										if(pts[j].compiledMarker) {
											pts[j].compiledMarker.pen && pts[j].compiledMarker.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
											pts[j].compiledMarker.brush && pts[j].compiledMarker.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
										}
									}
								}
							}
						}
					}
					else if(nChartType === AscDFH.historyitem_type_ScatterChart) {
						var base_line_fills = getArrayFillsFromBase(style.line4, nMaxSeriesIdx);
						for(var i = 0; i < series.length; ++i) {
							var default_line = parents.theme.themeElements.fmtScheme.lnStyleLst[0];
							var ser = series[i];
							var pts = ser.getNumPts();
							this.ptsCount += pts.length;
							if(oChart.scatterStyle === AscFormat.SCATTER_STYLE_SMOOTH || oChart.scatterStyle === AscFormat.SCATTER_STYLE_SMOOTH_MARKER) {
								if(!AscFormat.isRealBool(ser.smooth)) {
									ser.smooth = true;
								}
							}
							if(oChart.scatterStyle === AscFormat.SCATTER_STYLE_MARKER || oChart.scatterStyle === AscFormat.SCATTER_STYLE_NONE) {
								default_line = new AscFormat.CLn();
								default_line.setFill(new AscFormat.CUniFill());
								default_line.Fill.setFill(new AscFormat.CNoFill());
							}
							var compiled_line = new AscFormat.CLn();
							compiled_line.merge(default_line);
							if(!(oChart.scatterStyle === AscFormat.SCATTER_STYLE_MARKER || oChart.scatterStyle === AscFormat.SCATTER_STYLE_NONE)) {
								compiled_line.Fill && compiled_line.Fill.merge(base_line_fills[ser.idx]);
							}
							compiled_line.w *= style.line3;
							if(ser.spPr && ser.spPr.ln)
								compiled_line.merge(ser.spPr.ln);
							ser.compiledSeriesPen = compiled_line;
							for(var j = 0; j < pts.length; ++j) {
								pts[j].brush = null;
								pts[j].pen = ser.compiledSeriesPen;
								if(Array.isArray(ser.dPt)) {
									for(var k = 0; k < ser.dPt.length; ++k) {
										if(ser.dPt[k].idx === pts[j].idx) {
											if(ser.dPt[k].spPr) {
												pts[j].pen = ser.compiledSeriesPen.createDuplicate();
												pts[j].pen.merge(ser.dPt[k].spPr.ln);
												pts[j].pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
											}
											break;
										}
									}
								}
							}
						}
					}
					else if(nChartType === AscDFH.historyitem_type_SurfaceChart) {
						var oSurfaceChart = oChart;
						var aValAxArray = this.getValAxisValues();
						var nFmtsCount = aValAxArray.length - 1;
						var oSpPr, oBandFmt, oCompiledBandFmt;
						oSurfaceChart.compiledBandFormats.length = 0;
						var multiplier;
						var axis_by_types = oSurfaceChart.getAxisByTypes();
						var val_ax = axis_by_types.valAx[0];
						if(val_ax.dispUnits)
							multiplier = val_ax.dispUnits.getMultiplier();
						else
							multiplier = 1;
						var num_fmt = val_ax.numFmt, num_format = null, calc_value, rich_value;
						if(num_fmt && typeof num_fmt.formatCode === "string" /*&& !(num_fmt.formatCode === "General")*/) {
							num_format = oNumFormatCache.get(num_fmt.formatCode);
						}
						var oParentObjects = this.getParentObjects();
						var RGBA = {R: 255, G: 255, B: 255, A: 255};
						if(oSurfaceChart.isWireframe()) {
							var base_line_fills = getArrayFillsFromBase(style.line4, nFmtsCount);
							var default_line = parents.theme.themeElements.fmtScheme.lnStyleLst[0];
							for(var i = 0; i < nFmtsCount; ++i) {
								oBandFmt = oSurfaceChart.getBandFmtByIndex(i);
								oSpPr = new AscFormat.CSpPr();
								oSpPr.setFill(AscFormat.CreateNoFillUniFill());
								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								compiled_line.Fill.merge(base_line_fills[i]);
								//compiled_line.w *= style.line3;
								compiled_line.Join = new AscFormat.LineJoin();
								compiled_line.Join.type = AscFormat.LineJoinType.Bevel;
								if(oBandFmt && oBandFmt.spPr) {
									compiled_line.merge(oBandFmt.spPr.ln);
								}
								compiled_line.calculate(oParentObjects.theme, oParentObjects.slide, oParentObjects.layout, oParentObjects.master, RGBA, this.clrMapOvr);
								oSpPr.setLn(compiled_line);
								oCompiledBandFmt = new AscFormat.CBandFmt();
								oCompiledBandFmt.setIdx(i);
								oCompiledBandFmt.setSpPr(oSpPr);


								if(num_format) {
									oCompiledBandFmt.startValue = num_format.formatToChart(aValAxArray[i] * multiplier);
									oCompiledBandFmt.endValue = num_format.formatToChart(aValAxArray[i + 1] * multiplier);

								}
								else {
									oCompiledBandFmt.startValue = '' + (aValAxArray[i] * multiplier);
									oCompiledBandFmt.endValue = '' + (aValAxArray[i + 1] * multiplier);
								}
								oCompiledBandFmt.setSpPr(oSpPr);
								oSurfaceChart.compiledBandFormats.push(oCompiledBandFmt);
							}
						}
						else {
							var base_fills = getArrayFillsFromBase(style.fill2, nFmtsCount);
							var base_line_fills = null;
							if(style.line1 === EFFECT_SUBTLE && this.style === 34)
								base_line_fills = getArrayFillsFromBase(style.line2, nFmtsCount);

							var default_line = new AscFormat.CLn();
							if(style.line1 === EFFECT_NONE) {
								default_line.w = 0;
							}
							else if(style.line1 === EFFECT_SUBTLE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
							}
							else if(style.line1 === EFFECT_MODERATE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[1]);
							}
							else if(style.line1 === EFFECT_INTENSE) {
								default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[2]);
							}

							for(var i = 0; i < nFmtsCount; ++i) {
								oBandFmt = oSurfaceChart.getBandFmtByIndex(i);
								var compiled_brush = new AscFormat.CUniFill();
								oSpPr = new AscFormat.CSpPr();
								compiled_brush.merge(base_fills[i]);
								if(oBandFmt && oBandFmt.spPr) {
									compiled_brush.merge(oBandFmt.spPr.Fill);
								}
								oSpPr.setFill(compiled_brush);

								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								compiled_line.Fill = new AscFormat.CUniFill();
								if(this.style !== 34)
									compiled_line.Fill.merge(style.line2[0]);
								else
									compiled_line.Fill.merge(base_line_fills[i]);
								if(oBandFmt && oBandFmt.spPr && oBandFmt.spPr.ln) {
									compiled_line.merge(oBandFmt.spPr.ln);
								}
								oSpPr.setLn(compiled_line);
								compiled_line.calculate(oParentObjects.theme, oParentObjects.slide, oParentObjects.layout, oParentObjects.master, RGBA, this.clrMapOvr);
								compiled_brush.calculate(oParentObjects.theme, oParentObjects.slide, oParentObjects.layout, oParentObjects.master, RGBA, this.clrMapOvr);
								oCompiledBandFmt = new AscFormat.CBandFmt();
								oCompiledBandFmt.setIdx(i);
								oCompiledBandFmt.setSpPr(oSpPr);
								if(num_format) {
									oCompiledBandFmt.startValue = num_format.formatToChart(aValAxArray[i] * multiplier);
									oCompiledBandFmt.endValue = num_format.formatToChart(aValAxArray[i + 1] * multiplier);

								}
								else {
									oCompiledBandFmt.startValue = '' + (aValAxArray[i] * multiplier);
									oCompiledBandFmt.endValue = '' + (aValAxArray[i + 1] * multiplier);
								}
								oSurfaceChart.compiledBandFormats.push(oCompiledBandFmt);
							}
						}
					}
					else {

						var base_fills = getArrayFillsFromBase(style.fill2, nMaxSeriesIdx);
						var base_line_fills = null;
						if(style.line1 === EFFECT_SUBTLE && this.style === 34)
							base_line_fills = getArrayFillsFromBase(style.line2, nMaxSeriesIdx);
						for(var i = 0; i < series.length; ++i) {
							var ser = series[i];
							var compiled_brush = new AscFormat.CUniFill();
							compiled_brush.merge(base_fills[ser.idx]);
							if(ser.spPr && ser.spPr.Fill) {
								compiled_brush.merge(ser.spPr.Fill);
							}
							ser.compiledSeriesBrush = compiled_brush.createDuplicate();
							var pts = ser.getNumPts();
							this.ptsCount += pts.length;
							for(var j = 0; j < pts.length; ++j) {
								pts[j].brush = ser.compiledSeriesBrush;
								if(Array.isArray(ser.dPt) && !(ser.getObjectType && ser.getObjectType() === AscDFH.historyitem_type_AreaSeries)) {
									for(var k = 0; k < ser.dPt.length; ++k) {
										if(ser.dPt[k].idx === pts[j].idx) {
											if(ser.dPt[k].spPr) {
												if(ser.dPt[k].spPr.Fill && ser.dPt[k].spPr.Fill.fill) {
													compiled_brush = ser.dPt[k].spPr.Fill.createDuplicate();
													pts[j].brush = compiled_brush;
												}
											}
											break;
										}
									}
								}
							}


							//
							{
								default_line = new AscFormat.CLn();
								if(style.line1 === EFFECT_NONE) {
									default_line.w = 0;
								}
								else if(style.line1 === EFFECT_SUBTLE) {
									default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[0]);
								}
								else if(style.line1 === EFFECT_MODERATE) {
									default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[1]);
								}
								else if(style.line1 === EFFECT_INTENSE) {
									default_line.merge(parents.theme.themeElements.fmtScheme.lnStyleLst[2]);
								}
								var base_line_fills;
								if(this.style === 34)
									base_line_fills = getArrayFillsFromBase(style.line2, getMaxIdx(pts));


								var compiled_line = new AscFormat.CLn();
								compiled_line.merge(default_line);
								compiled_line.Fill = new AscFormat.CUniFill();
								if(this.style !== 34)
									compiled_line.Fill.merge(style.line2[0]);
								else
									compiled_line.Fill.merge(base_line_fills[ser.idx]);
								if(ser.spPr && ser.spPr.ln) {
									compiled_line.merge(ser.spPr.ln);
								}
								ser.compiledSeriesPen = compiled_line.createDuplicate();
								ser.compiledSeriesPen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA, this.clrMapOvr);
								for(var j = 0; j < pts.length; ++j) {
									pts[j].pen = ser.compiledSeriesPen;
									if(Array.isArray(ser.dPt) && !(ser.getObjectType && ser.getObjectType() === AscDFH.historyitem_type_AreaSeries)) {
										for(var k = 0; k < ser.dPt.length; ++k) {
											if(ser.dPt[k].idx === pts[j].idx) {
												if(ser.dPt[k].spPr) {
													compiled_line = ser.compiledSeriesPen.createDuplicate();
													compiled_line.merge(ser.dPt[k].spPr.ln);
													pts[j].pen = compiled_line;
												}
												break;
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
	};
	CChartSpace.prototype.recalculateChartTitleEditMode = function (bWord) {
		var old_pos_x, old_pos_y, old_pos_cx, old_pos_cy;
		var pos_x, pos_y;
		old_pos_x = this.recalcInfo.recalcTitle.x;
		old_pos_y = this.recalcInfo.recalcTitle.y;
		old_pos_cx = this.recalcInfo.recalcTitle.x + this.recalcInfo.recalcTitle.extX / 2;
		old_pos_cy = this.recalcInfo.recalcTitle.y + this.recalcInfo.recalcTitle.extY / 2;
		this.cachedCanvas = null;
		this.recalculateAxisLabels();
		this.recalculateDLbls();
		var oDlbl = this.getCompiledDlblBySelect();
		if (oDlbl && this.selection.textSelection instanceof AscFormat.CDLbl) {
			this.selection.textSelection = oDlbl;
		}
		if (checkVerticalTitle(this.recalcInfo.recalcTitle)) {
			pos_x = old_pos_x;
			pos_y = old_pos_cy - this.recalcInfo.recalcTitle.extY / 2;
		} else {
			pos_x = old_pos_cx - this.recalcInfo.recalcTitle.extX / 2;
			pos_y = old_pos_y;
		}
		this.recalcInfo.recalcTitle.setPosition(pos_x, pos_y);
		if (bWord) {
			this.recalcInfo.recalcTitle.updatePosition(this.posX, this.posY);
		}
	};
	CChartSpace.prototype.recalculateMarkers = function () {
		if (this.chart && this.chart.plotArea) {
			var aCharts = this.chart.plotArea.charts;
			var oCurChart;
			var aAllsSeries = [];
			for (var t = 0; t < aCharts.length; ++t) {
				oCurChart = aCharts[t];
				var series = oCurChart.series, pts;
				aAllsSeries = aAllsSeries.concat(series);
				for (var i = 0; i < series.length; ++i) {
					var ser = series[i];
					ser.compiledSeriesMarker = null;
					pts = ser.getNumPts();
					for (var j = 0; j < pts.length; ++j) {
						pts[j].compiledMarker = null;
					}
				}

				var oThis = this;
				var recalculateMarkers2 = function () {
					var chart_style = CHART_STYLE_MANAGER.getStyleByIndex(oThis.style);
					var effect_fill = chart_style.fill1;
					var fill = chart_style.fill2;
					var line = chart_style.line4;
					var masrker_default_size = AscFormat.isRealNumber(oThis.style) ? chart_style.markerSize : 5;
					var default_marker = new AscFormat.CMarker();
					default_marker.setSize(masrker_default_size);
					var parent_objects = oThis.getParentObjects();

					if (parent_objects.theme && parent_objects.theme.themeElements
						&& parent_objects.theme.themeElements.fmtScheme
						&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst) {
						default_marker.setSpPr(new AscFormat.CSpPr());
						default_marker.spPr.setLn(new AscFormat.CLn());
						default_marker.spPr.ln.merge(parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0]);
					}
					var RGBA = {R: 0, G: 0, B: 0, A: 255};
					if (oCurChart.varyColors && (oCurChart.series.length === 1 || oCurChart.getObjectType() === AscDFH.historyitem_type_PieChart || oCurChart.getObjectType() === AscDFH.historyitem_type_DoughnutChart)) {
						if (oCurChart.series.length < 1) {
							return;
						}
						var ser = oCurChart.series[0], pts;
						if (ser.marker && ser.marker.symbol === AscFormat.SYMBOL_NONE && (!Array.isArray(ser.dPt) || ser.dPt.length === 0))
							return;
						pts = ser.getNumPts();
						var series_marker = ser.marker;
						var brushes = getArrayFillsFromBase(fill, getMaxIdx(pts));

						for (var i = 0; i < pts.length; ++i) {
							var d_pt = null;
							if (Array.isArray(ser.dPt)) {
								for (var j = 0; j < ser.dPt.length; ++j) {
									if (ser.dPt[j].idx === pts[i].idx) {
										d_pt = ser.dPt[j];
										break;
									}
								}
							}
							var d_pt_marker = null;
							if (d_pt) {
								d_pt_marker = d_pt.marker;
							}
							var symbol = GetTypeMarkerByIndex(i);
							if (series_marker && AscFormat.isRealNumber(series_marker.symbol)) {
								symbol = series_marker.symbol;
							}
							if (d_pt_marker && AscFormat.isRealNumber(d_pt_marker.symbol)) {
								symbol = d_pt_marker.symbol;
							}

							var compiled_marker = new AscFormat.CMarker();
							compiled_marker.merge(default_marker);
							if (!compiled_marker.spPr) {
								compiled_marker.setSpPr(new AscFormat.CSpPr());
							}
							compiled_marker.setSymbol(symbol);
							if (!checkNoFillMarkers(symbol)) {
								compiled_marker.spPr.setFill(brushes[pts[i].idx]);
							} else {
								compiled_marker.spPr.setFill(AscFormat.CreateNoFillUniFill());
							}
							compiled_marker.spPr.Fill.merge(pts[i].brush);
							if (!compiled_marker.spPr.ln)
								compiled_marker.spPr.setLn(new AscFormat.CLn());
							compiled_marker.spPr.ln.merge(pts[i].pen);
							compiled_marker.merge(series_marker);

							if (d_pt) {
								if (d_pt.spPr) {
									if (d_pt.spPr.Fill) {
										compiled_marker.spPr.setFill(d_pt.spPr.Fill.createDuplicate());
									}
									if (d_pt.spPr.ln) {
										if (!compiled_marker.spPr.ln) {
											compiled_marker.spPr.setLn(new AscFormat.CLn());
										}
										compiled_marker.spPr.ln.merge(d_pt.spPr.ln);
									}
								}
								compiled_marker.merge(d_pt.marker);
							}

							pts[i].compiledMarker = compiled_marker;
							pts[i].compiledMarker.pen = compiled_marker.spPr.ln;
							pts[i].compiledMarker.brush = compiled_marker.spPr.Fill;
							pts[i].compiledMarker.brush.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
							pts[i].compiledMarker.pen.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
						}
					} else {
						var series = oCurChart.series;
						var brushes = getArrayFillsFromBase(fill, getMaxIdx(aAllsSeries));
						var pens_fills = getArrayFillsFromBase(line, getMaxIdx(aAllsSeries));


						for (var i = 0; i < series.length; ++i) {
							var ser = series[i];
							if (ser.marker && ser.marker.symbol === AscFormat.SYMBOL_NONE && (!Array.isArray(ser.dPt) || ser.dPt.length === 0))
								continue;


							var compiled_marker = new AscFormat.CMarker();
							var symbol;
							if (ser.marker && AscFormat.isRealNumber(ser.marker.symbol)) {
								symbol = ser.marker.symbol;
							} else {
								symbol = GetTypeMarkerByIndex(ser.idx);
							}
							compiled_marker.merge(default_marker);
							if (!compiled_marker.spPr) {
								compiled_marker.setSpPr(new AscFormat.CSpPr());
							}
							compiled_marker.setSymbol(symbol);
							if (!checkNoFillMarkers(compiled_marker.symbol)) {
								compiled_marker.spPr.setFill(brushes[ser.idx]);
							} else {
								compiled_marker.spPr.setFill(AscFormat.CreateNoFillUniFill());
							}
							if (!compiled_marker.spPr.ln)
								compiled_marker.spPr.setLn(new AscFormat.CLn());
							compiled_marker.spPr.ln.setFill(pens_fills[ser.idx]);
							compiled_marker.merge(ser.marker);
							ser.compiledSeriesMarker = compiled_marker.createDuplicate();
							ser.compiledSeriesMarker.pen = compiled_marker.spPr.ln;
							ser.compiledSeriesMarker.brush = compiled_marker.spPr.Fill;
							ser.compiledSeriesMarker.brush.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
							ser.compiledSeriesMarker.pen.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
							pts = ser.getNumPts();
							for (var j = 0; j < pts.length; ++j) {
								var d_pt = null;
								if (Array.isArray(ser.dPt)) {
									for (var k = 0; k < ser.dPt.length; ++k) {
										if (ser.dPt[k].idx === pts[j].idx) {
											d_pt = ser.dPt[k];
											break;
										}
									}
								}

								if (d_pt) {
									compiled_marker = ser.compiledSeriesMarker.createDuplicate();
									compiled_marker.merge(d_pt.marker);
									compiled_marker.pen = compiled_marker.spPr.ln;
									compiled_marker.brush = compiled_marker.spPr.Fill;
									compiled_marker.brush.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
									compiled_marker.pen.calculate(parent_objects.theme, parent_objects.slide, parent_objects.layout, parent_objects.master, RGBA, oThis.clrMapOvr);
									pts[j].compiledMarker = compiled_marker;
								} else {
									pts[j].compiledMarker = ser.compiledSeriesMarker;
								}
							}
						}
					}
				};

				switch (oCurChart.getObjectType()) {
					case AscDFH.historyitem_type_LineChart:
					case AscDFH.historyitem_type_RadarChart: {
						if (oCurChart.isMarkerChart()) {
							recalculateMarkers2();
						}
						break;
					}
					case AscDFH.historyitem_type_ScatterChart: {
						if (oCurChart.scatterStyle === AscFormat.SCATTER_STYLE_MARKER || oCurChart.scatterStyle === AscFormat.SCATTER_STYLE_LINE_MARKER || oCurChart.scatterStyle === AscFormat.SCATTER_STYLE_SMOOTH_MARKER) {
							recalculateMarkers2();
						}
						break;
					}
					default: {
						recalculateMarkers2();
						break;
					}
				}
			}
		}
	};
	CChartSpace.prototype.calcGridLine = function (defaultStyle, spPr, subtleLine, parents) {
		if (spPr) {
			var compiled_grid_lines = new AscFormat.CLn();
			compiled_grid_lines.merge(subtleLine);
			// if(compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color && compiled_grid_lines.Fill.fill.color.Mods)
			// {
			//     compiled_grid_lines.Fill.fill.color.Mods.Mods.length = 0;
			// }
			if (!compiled_grid_lines.Fill) {
				compiled_grid_lines.setFill(new AscFormat.CUniFill());
			}
			//if(compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color && compiled_grid_lines.Fill.fill.color.Mods)
			//{
			//    compiled_grid_lines.Fill.fill.color.Mods.Mods.length = 0;
			//}
			compiled_grid_lines.Fill.merge(defaultStyle);

			if (subtleLine && subtleLine.Fill && subtleLine.Fill.fill && subtleLine.Fill.fill.color && subtleLine.Fill.fill.color.Mods
				&& compiled_grid_lines.Fill && compiled_grid_lines.Fill.fill && compiled_grid_lines.Fill.fill.color) {
				compiled_grid_lines.Fill.fill.color.Mods = subtleLine.Fill.fill.color.Mods.createDuplicate();
			}
			compiled_grid_lines.merge(spPr.ln);
			compiled_grid_lines.calculate(parents.theme, parents.slide, parents.layout, parents.master, {
				R: 0,
				G: 0,
				B: 0,
				A: 255
			}, this.clrMapOvr);
			checkBlackUnifill(compiled_grid_lines.Fill, true);
			return compiled_grid_lines;
		}
		return null;
	};
	CChartSpace.prototype.calcMajorMinorGridLines = function (axis, defaultStyle, subtleLine, parents) {
		if (!axis)
			return;
		axis.compiledMajorGridLines = this.calcGridLine(defaultStyle.axisAndMajorGridLines, axis.majorGridlines, subtleLine, parents);
		axis.compiledMinorGridLines = this.calcGridLine(defaultStyle.minorGridlines, axis.minorGridlines, subtleLine, parents);
	};
	CChartSpace.prototype.handleTitlesAfterChangeTheme = function () {
		if (this.chart && this.chart.title) {
			this.chart.title.checkAfterChangeTheme();
		}
		if (this.chart && this.chart.plotArea) {
			var hor_axis = this.chart.plotArea.getHorizontalAxis();
			if (hor_axis && hor_axis.title) {
				hor_axis.title.checkAfterChangeTheme();
			}
			var vert_axis = this.chart.plotArea.getVerticalAxis();
			if (vert_axis && vert_axis.title) {
				vert_axis.title.checkAfterChangeTheme();
			}
		}
	};
	CChartSpace.prototype.recalculateGridLines = function () {
		if (this.chart && this.chart.plotArea) {
			var default_style = CHART_STYLE_MANAGER.getDefaultLineStyleByIndex(this.style);
			var parent_objects = this.getParentObjects();
			var subtle_line;
			if (parent_objects.theme && parent_objects.theme.themeElements
				&& parent_objects.theme.themeElements.fmtScheme
				&& parent_objects.theme.themeElements.fmtScheme.lnStyleLst) {
				subtle_line = parent_objects.theme.themeElements.fmtScheme.lnStyleLst[0];
			}
			var aAxes = this.chart.plotArea.axId;
			for (var i = 0; i < aAxes.length; ++i) {
				var oCurAxis = aAxes[i];
				this.calcMajorMinorGridLines(oCurAxis, default_style, subtle_line, parent_objects);
				this.calcMajorMinorGridLines(oCurAxis, default_style, subtle_line, parent_objects);
			}
		}
	};
	CChartSpace.prototype.recalculateAxisLabels = function () {
		if (this.chart && this.chart.title) {
			var title = this.chart.title;
			//title.parent = this.chart;
			title.chart = this;
			title.recalculate();
		}
		if (this.chart && this.chart.plotArea) {
			var aAxis = this.chart.plotArea.axId;
			for (var i = 0; i < aAxis.length; ++i) {
				var title = aAxis[i].title;
				if (title) {
					title.chart = this;
					title.recalculate();
				}
			}
		}
	};
	CChartSpace.prototype.updateLinks = function () {
		//Этот метод нужен, т. к. мы не полностью поддерживаем формат в котором в одном plotArea может быть несколько разных диаграмм(конкретных типов).
		// Здесь мы берем первую из диаграмм лежащих в массиве plotArea.charts, а также выставляем ссылки для осей ;
		if (this.chart && this.chart.plotArea) {
			var oCheckChart;
			var oPlotArea = this.chart.plotArea;
			var aCharts = oPlotArea.charts;

			oPlotArea.valAx = null;
			oPlotArea.catAx = null;
			oPlotArea.serAx = null;
			oPlotArea.chart = aCharts[0];
			for (var i = 0; i < aCharts.length; ++i) {
				if (aCharts[i].getObjectType() !== AscDFH.historyitem_type_PieChart && aCharts[i].getObjectType() !== AscDFH.historyitem_type_DoughnutChart) {
					oCheckChart = aCharts[i];
					break;
				}
			}
			if (oCheckChart) {
				oPlotArea.chart = oCheckChart;
				if (oCheckChart.getAxisByTypes) {
					var axis_by_types = oCheckChart.getAxisByTypes();
					if (axis_by_types.valAx.length > 0 && axis_by_types.catAx.length > 0) {
						for (var i = 0; i < axis_by_types.valAx.length; ++i) {
							if (axis_by_types.valAx[i].crossAx) {
								for (var j = 0; j < axis_by_types.catAx.length; ++j) {
									if (axis_by_types.catAx[j] === axis_by_types.valAx[i].crossAx) {
										oPlotArea.valAx = axis_by_types.valAx[i];
										oPlotArea.catAx = axis_by_types.catAx[j];
										break;
									}
								}
								if (j < axis_by_types.catAx.length) {
									break;
								}
							}
						}
						if (i === axis_by_types.valAx.length) {
							oPlotArea.valAx = axis_by_types.valAx[0];
							oPlotArea.catAx = axis_by_types.catAx[0];
						}
						if (oPlotArea.valAx && oPlotArea.catAx) {
							for (i = 0; i < axis_by_types.serAx.length; ++i) {
								if (axis_by_types.serAx[i].crossAx === oPlotArea.valAx) {
									oPlotArea.serAx = axis_by_types.serAx[i];
									break;
								}
							}
						}
					} else {
						if (axis_by_types.valAx.length > 1) {//TODO: выставлять оси исходя из настроек
							oPlotArea.valAx = axis_by_types.valAx[1];
							oPlotArea.catAx = axis_by_types.valAx[0];
						}
					}
				}
			}
		}
	};
	CChartSpace.prototype.checkDrawingCache = function (graphics) {
		if (window["NATIVE_EDITOR_ENJINE"] || graphics.RENDERER_PDF_FLAG || this.isSparkline || this.bPreview || graphics.PrintPreview) {
			return false;
		}
		if (graphics.IsSlideBoundsCheckerType) {
			return false;
		}
		if (!this.transform.IsIdentity2()) {
			return false;
		}

		var dBorderW = this.getBorderWidth();
		var dWidth = this.extX + 2 * dBorderW;
		var dHeight = this.extY + 2 * dBorderW;
		var nWidth = (graphics.m_oCoordTransform.sx * dWidth + 0.5) >> 0;
		var nHeight = (graphics.m_oCoordTransform.sy * dHeight + 0.5) >> 0;
		if (nWidth === 0 || nHeight === 0) {
			return false;
		}
		if (this.cachedCanvas) {
			if (this.cachedCanvas.width !== nWidth || this.cachedCanvas.height !== nHeight) {
				this.cachedCanvas = null;
			}
		}
		var ctx;
		if (!this.cachedCanvas) {
			//check loading images
			var aImages = [];
			this.getAllRasterImages(aImages);
			var oApi = window["Asc"]["editor"] || editor;
			if (oApi && oApi.ImageLoader) {
				for (var i = 0; i < aImages.length; ++i) {
					var _image = oApi.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(aImages[i])];
					if (_image != undefined && _image.Image != null && _image.Status === window['AscFonts'].ImageLoadStatus.Loading) {
						return false;
					}
				}
			}

			this.cachedCanvas = document.createElement('canvas');
			this.cachedCanvas.width = nWidth;
			this.cachedCanvas.height = nHeight;
			ctx = this.cachedCanvas.getContext('2d');
			var g = new AscCommon.CGraphics();
			g.init(ctx, nWidth, nHeight, nWidth / graphics.m_oCoordTransform.sx, nHeight / graphics.m_oCoordTransform.sy);
			g.m_oFontManager = AscCommon.g_fontManager;
			g.m_oCoordTransform.tx = -((this.transform.tx - dBorderW) * graphics.m_oCoordTransform.sx + 0.5) >> 0;
			g.m_oCoordTransform.ty = -((this.transform.ty - dBorderW) * graphics.m_oCoordTransform.sy + 0.5) >> 0;
			g.transform(1, 0, 0, 1, 0, 0);
			this.drawInternal(g);
		}

		var _x1 = (graphics.m_oCoordTransform.TransformPointX(this.transform.tx - dBorderW, this.transform.ty - dBorderW) + 0.5) >> 0;
		var _y1 = (graphics.m_oCoordTransform.TransformPointY(this.transform.tx - dBorderW, this.transform.ty - dBorderW) + 0.5) >> 0;

		graphics.SaveGrState();
		graphics.SetIntegerGrid(true);
		graphics.m_oContext.drawImage(this.cachedCanvas, _x1, _y1, nWidth, nHeight);
		graphics.RestoreGrState();
		graphics.FreeFont();
		return true;
	};
	CChartSpace.prototype.getBorderWidth = function () {
		var ln_width;
		if (this.pen && AscFormat.isRealNumber(this.pen.w)) {
			ln_width = this.pen.w / 36000;
		} else {
			ln_width = 0;
		}
		return ln_width;
	};
	CChartSpace.prototype.drawInternal = function (graphics) {
		var oClipRect = this.getClipRect();
		if (oClipRect) {
			graphics.SaveGrState();
			graphics.AddClipRect(oClipRect.x, oClipRect.y, oClipRect.w, oClipRect.h);
		}
		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.transform3(this.transform, false);

		var ln_width = this.getBorderWidth();
		graphics.AddClipRect(-ln_width, -ln_width, this.extX + 2 * ln_width, this.extY + 2 * ln_width);

		if (this.chartObj) {
			this.chartObj.draw(this, graphics);
		}
		if (this.chart && !this.bEmptySeries) {
			if (this.chart.plotArea) {
				// var oChartSize = this.getChartSizes();
				// graphics.p_width(70);
				// graphics.p_color(0, 0, 0, 255);
				// graphics._s();
				// graphics._m(oChartSize.startX, oChartSize.startY);
				// graphics._l(oChartSize.startX + oChartSize.w, oChartSize.startY + 0);
				// graphics._l(oChartSize.startX + oChartSize.w, oChartSize.startY + oChartSize.h);
				// graphics._l(oChartSize.startX + 0, oChartSize.startY + oChartSize.h);
				// graphics._z();
				// graphics.ds();
				var aCharts = this.chart.plotArea.charts;
				for (var t = 0; t < aCharts.length; ++t) {
					var oChart = aCharts[t];
					if (oChart && oChart.series) {
						var series = oChart.series;
						var _len = oChart.getObjectType() === AscDFH.historyitem_type_PieChart ? 1 : series.length;
						for (var i = 0; i < _len; ++i) {
							var ser = series[i];
							var pts = ser.getNumPts();
							for (var j = 0; j < pts.length; ++j) {
								if (pts[j].compiledDlb)
									pts[j].compiledDlb.draw(graphics);
							}
						}
					}
				}
				for (var i = 0; i < this.chart.plotArea.axId.length; ++i) {
					var oAxis = this.chart.plotArea.axId[i];
					if (oAxis.title) {
						oAxis.title.draw(graphics);
					}
					if (oAxis.labels) {
						oAxis.labels.draw(graphics);
					}
				}
			}
			if (this.chart.title) {
				this.chart.title.draw(graphics);
			}
			if (this.chart.legend) {
				this.chart.legend.draw(graphics);
			}
		}
		for (var i = 0; i < this.userShapes.length; ++i) {
			if (this.userShapes[i].object) {
				this.userShapes[i].object.draw(graphics);
			}
		}
		graphics.RestoreGrState();
		if (oClipRect) {
			graphics.RestoreGrState();
		}
	};
	CChartSpace.prototype.draw = function (graphics) {
		if (this.checkNeedRecalculate && this.checkNeedRecalculate()) {
			return;
		}
		if (graphics.IsSlideBoundsCheckerType) {
			graphics.transform3(this.transform);
			graphics._s();
			graphics._m(0, 0);
			graphics._l(this.extX, 0);
			graphics._l(this.extX, this.extY);
			graphics._l(0, this.extY);
			graphics._e();
			return;
		}
		var oUR = graphics.updatedRect;
		if (oUR && this.bounds) {
			if (!oUR.isIntersectOther(this.bounds)) {
				return;
			}
		}
		if (graphics.animationDrawer) {
			graphics.animationDrawer.drawObject(this, graphics);
			return;
		}
		var oldShowParaMarks;
		if (editor) {
			oldShowParaMarks = editor.ShowParaMarks;
			editor.ShowParaMarks = false;
		}
		if (!this.checkDrawingCache(graphics)) {
			this.drawInternal(graphics);
		}
		if (editor) {
			editor.ShowParaMarks = oldShowParaMarks;
		}

		if (this.drawLocks(this.transform, graphics)) {
			graphics.RestoreGrState();
		}
	};
	CChartSpace.prototype.addToSetPosition = function (dLbl) {
		if (dLbl instanceof AscFormat.CDLbl)
			this.recalcInfo.dataLbls.push(dLbl);
		else if (dLbl instanceof AscFormat.CTitle)
			this.recalcInfo.axisLabels.push(dLbl);
	};
	CChartSpace.prototype.recalculateChart = function () {
		this.pathMemory.curPos = -1;
		if (this.chartObj == null)
			this.chartObj = new AscFormat.CChartsDrawer();
		this.chartObj.recalculate(this);


		var oChartSize = this.getChartSizes(true);
		this.chart.plotArea.x = oChartSize.startX;
		this.chart.plotArea.y = oChartSize.startY;
		this.chart.plotArea.extX = oChartSize.w;
		this.chart.plotArea.extY = oChartSize.h;
		this.chart.plotArea.localTransform.Reset();
		AscCommon.global_MatrixTransformer.TranslateAppend(this.chart.plotArea.localTransform, oChartSize.startX, oChartSize.startY);
	};
	CChartSpace.prototype.GetRevisionsChangeElement = function (SearchEngine) {
		var titles = this.getAllTitles(), i;
		if (titles.length === 0) {
			return;
		}
		var oSelectedTitle = this.selection.title || this.selection.textSelection;
		if (oSelectedTitle) {
			for (i = 0; i < titles.length; ++i) {
				if (oSelectedTitle === titles[i]) {
					break;
				}
			}
			if (i === titles.length) {
				return;
			}
		} else {
			if (SearchEngine.GetDirection() > 0) {
				i = 0;
			} else {
				i = titles.length - 1;
			}
		}
		while (!SearchEngine.IsFound()) {
			titles[i].GetRevisionsChangeElement(SearchEngine);
			if (SearchEngine.GetDirection() > 0) {
				if (i === titles.length - 1) {
					break;
				}
				++i;
			} else {
				if (i === 0) {
					break;
				}
				--i;
			}
		}
	};
	CChartSpace.prototype.Search = function (SearchEngine, Type) {
		var titles = this.getAllTitles();
		for (var i = 0; i < titles.length; ++i) {
			titles[i].Search(SearchEngine, Type)
		}
	};
	CChartSpace.prototype.GetSearchElementId = function (bNext, bCurrent) {
		var Current = -1;
		var titles = this.getAllTitles();
		var Len = titles.length;

		var Id = null;
		if (true === bCurrent) {
			for (var i = 0; i < Len; ++i) {
				if (titles[i] === this.selection.textSelection) {
					Current = i;
					break;
				}
			}
		}

		if (true === bNext) {
			var Start = (-1 !== Current ? Current : 0);

			for (var i = Start; i < Len; i++) {
				if (titles[i].GetSearchElementId) {
					Id = titles[i].GetSearchElementId(true, i === Current ? true : false);
					if (null !== Id)
						return Id;
				}
			}
		} else {
			var Start = (-1 !== Current ? Current : Len - 1);

			for (var i = Start; i >= 0; i--) {
				if (titles[i].GetSearchElementId) {
					Id = titles[i].GetSearchElementId(false, i === Current ? true : false);
					if (null !== Id)
						return Id;
				}
			}
		}
		return null;
	};
	CChartSpace.prototype.FindNextFillingForm = function (isNext, isCurrent) {
		return null;
	};
	CChartSpace.prototype.isAccent1Background = function () {
		return this.spPr && this.spPr.Fill && this.spPr.Fill.isAccent1();
	};
	CChartSpace.prototype.addNewSeries = function () {
		var oLastChart = this.chart.plotArea.charts[this.chart.plotArea.charts.length - 1];
		if (!oLastChart) {
			return null;
		}
		if (oLastChart.getObjectType() === AscDFH.historyitem_type_ScatterChart) {
			return this.addScatterSeries(null, null, "={1}");
		} else {
			return this.addSeries(null, "={1}");
		}
	};
	CChartSpace.prototype.setNewSeriesIdxAndOrder = function (oSeries) {
		var aAllSeries = this.getAllSeries();
		var nSeries;
		var oFirstSeries = aAllSeries[0];
		var nIdx = 0, nOrder = 0;
		if (aAllSeries.length > 0) {
			if (oFirstSeries.idx > 0) {
				nIdx = 0;
			} else {
				//Find a gap between series indexes
				var oCurSeries, oNextSeries;
				for (nSeries = 0; nSeries < aAllSeries.length - 1; ++nSeries) {
					oCurSeries = aAllSeries[nSeries];
					oNextSeries = aAllSeries[nSeries + 1];
					if (oNextSeries.idx - oCurSeries.idx > 1) {
						nIdx = oCurSeries.idx + 1;
						break;
					}
				}
				if (nSeries === aAllSeries.length - 1) {
					//There is no gap in the series indexes
					nIdx = aAllSeries[aAllSeries.length - 1].idx + 1;
				}
			}
			aAllSeries.sort(function (a, b) {
				return a.order - b.order;
			});
			nOrder = aAllSeries[aAllSeries.length - 1].order + 1;
		}
		oSeries.setIdx(nIdx);
		oSeries.setOrder(nOrder);
	};
	CChartSpace.prototype.addSeries = function (sName, sValues) {
		var oLastChart = this.chart.plotArea.charts[this.chart.plotArea.charts.length - 1];
		if (!oLastChart) {
			return null;
		}
		History.Create_NewPoint(0);
		var oSeries;
		oSeries = oLastChart.series[0] ? oLastChart.series[0].createDuplicate() : oLastChart.getEmptySeries();
		oSeries.setName(sName);
		oSeries.setValues(sValues);
		this.setNewSeriesIdxAndOrder(oSeries);
		oLastChart.addSer(oSeries);
		this.reorderSeries();
		if (this.chartStyle && this.chartColors) {
			oSeries.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), true);
		} else {
			oSeries.resetFormatting();
		}
		return oSeries;
	};
	CChartSpace.prototype.addScatterSeries = function (sName, sXValues, sYValues) {
		var oLastChart = this.chart.plotArea.charts[this.chart.plotArea.charts.length - 1];
		if (!oLastChart || oLastChart.getObjectType() !== AscDFH.historyitem_type_ScatterChart) {
			return;
		}
		History.Create_NewPoint(0);
		var oSeries;
		oSeries = oLastChart.series[0] ? oLastChart.series[0].createDuplicate() : oLastChart.getEmptySeries();
		oSeries.setName(sName);
		oSeries.setYValues(sYValues);
		oSeries.setXValues(sXValues);
		this.setNewSeriesIdxAndOrder(oSeries);
		oLastChart.addSer(oSeries);

		if (oSeries.spPr) {
			oSeries.setSpPr(null);
		}
		if (this.chartStyle && this.chartColors) {
			oSeries.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), true);
		} else {
			oSeries.resetFormatting();
		}
		return oSeries;
	};
	CChartSpace.prototype.collectRefsInsideRange = function (oRange, aRefs) {
		var oDataRange = this.getDataRefs();
		oDataRange.collectRefsInsideRange(oRange, aRefs);
	};
	CChartSpace.prototype.collectIntersectionRefs = function (aRanges, aRefs) {
		if (!Array.isArray(aRanges) || aRanges.length === 0) {
			return;
		}
		var oDataRange = this.getDataRefs();
		oDataRange.collectIntersectionRefs(aRanges, aRefs);
	};
	CChartSpace.prototype.getCommonRange = function () {
		var oDataRange = this.getDataRefs();
		return oDataRange.getRange();
	};
	CChartSpace.prototype.fillSelectedRanges = function (oWSView) {
		var oSelectedSeries = this.getSelectedSeries();
		if (oSelectedSeries) {
			oSelectedSeries.fillSelectedRanges(oWSView);
			return;
		}
		var oDataRange = this.getDataRefs();
		oDataRange.fillSelectedRanges(oWSView);
	};
	CChartSpace.prototype.canResetToStyle = function () {
		if (this.chartStyle && this.chartColors) {
			return true;
		}
	};
	CChartSpace.prototype.resetToStyle = function () {
		if (!this.chartStyle || !this.chartColors) {
			return;
		}
		this.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), true);

	};
	CChartSpace.prototype.getChartStyleIdx = function () {
		if (!this.chartStyle || !this.chartColors) {
			return null;
		}
		return oChartStyleCache.getStyleIdx(this.getChartType(), this.chartStyle.id);
	};
	CChartSpace.prototype.buildSeries = function (aRefs) {
		if (!Array.isArray(aRefs)) {
			return Asc.c_oAscError.ID.No;
		}
		if (aRefs.length > MAX_SERIES_COUNT) {
			return Asc.c_oAscError.ID.MaxDataSeriesError;
		}
		this.reindexSeries();
		var aAllSeries = this.getAllSeries();
		aAllSeries.sort(function (a, b) {
			return a.order - b.order;
		});

		var nRef, oRef;
		var nSeries, oSeries = null;
		var aAllCharts = this.chart.plotArea.charts;
		var oFirstChart = aAllCharts[0];
		var oLastChart = aAllCharts[aAllCharts.length - 1];
		var oLastSeries;
		for (nRef = 0; nRef < aRefs.length; ++nRef) {
			oRef = aRefs[nRef];
			if (nRef < aAllSeries.length) {
				oSeries = aAllSeries[nRef];
			} else {
				oLastSeries = oLastChart.series[oLastChart.series.length - 1];
				if (oLastSeries) {
					oSeries = oLastSeries.createDuplicate();
					oLastSeries.parent.addSerSilent(oSeries);
				} else {
					oSeries = oLastChart.getEmptySeries();
					oLastChart.addSerSilent(oSeries);
				}
				oSeries.setIdx(nRef);
				oSeries.setOrder(nRef);
			}
			oSeries.setData(oRef);

			var bNewSeries = nRef >= aAllSeries.length;
			if (this.chartStyle && this.chartColors) {
				oSeries.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), bNewSeries);
			} else {
				if (bNewSeries) {
					oSeries.resetFormatting();
				}
			}
		}
		for (nSeries = aAllSeries.length - 1; nSeries >= aRefs.length; --nSeries) {
			aAllSeries[nSeries].remove();
		}
		return Asc.c_oAscError.ID.No;
	};
	CChartSpace.prototype.setRange = function (sRange) {
		if (sRange === this.getCommonRange()) {
			return;
		}
		var oDataRange = this.getDataRefs();
		var aRefs = oDataRange.getSeriesRefsFromUnionRefs(AscFormat.fParseChartFormulaExternal(sRange), undefined, AscFormat.isScatterChartType(this.getChartType()));
		if (!Array.isArray(aRefs)) {
			this.buildSeries([]);
		} else {
			this.buildSeries(aRefs);
		}
		this.recalculate();
		if (this.pivotSource) {
			this.setPivotSource(null);
		}
	};
	CChartSpace.prototype.switchRowCol = function () {
		var oDataRange = this.getDataRefs();
		var aRefs = oDataRange.getSwitchedRefs(this.isScatterChartType());
		if (!aRefs) {
			return Asc.c_oAscError.ID.No;
		}
		var nResult = this.buildSeries(aRefs);
		if (Asc.c_oAscError.ID.No === nResult) {
			this.recalculate();
			this.checkLegendLayoutSize();
		}
		return nResult;
	};
	CChartSpace.prototype.fillDataFromTrack = function (oSelectedRange) {
		let oSelectedSeries = this.getSelectedSeries();
		if (oSelectedSeries) {
			oSelectedSeries.fillFromSelectedRange(oSelectedRange);
			this.recalculate();
			return;
		}
		let oDataRange = this.getDataRefs();
		let nResult = this.buildSeries(oDataRange.getSeriesRefsFromSelectedRange(oSelectedRange, this.isScatterChartType()));
		if (Asc.c_oAscError.ID.No === nResult) {
			this.recalculate();
		}
	};
	CChartSpace.prototype.checkLegendLayoutSize = function () {
		var oChart = this.chart;
		if (!oChart) {
			return;
		}
		var oLegend = oChart.legend;
		if (oLegend) {
			var oLayout = oLegend.layout;
			if (oLayout) {
				if (AscFormat.isRealNumber(oLayout.w) || AscFormat.isRealNumber(oLayout.h)) {
					oLegend.setLayout(null);
					this.recalcInfo.recalculateLegend = true;
					this.recalculate();
					if (AscFormat.isRealNumber(oLayout.w)) {
						oLayout.setW(this.calculateLayoutBySize(oLegend.x, oLayout.wMode, this.extX, oLegend.extX));
					}
					if (AscFormat.isRealNumber(oLayout.h)) {
						oLayout.setH(this.calculateLayoutBySize(oLegend.y, oLayout.hMode, this.extY, oLegend.extY));
					}
					oLegend.setLayout(oLayout);
					this.recalcInfo.recalculateLegend = true;
					this.recalculate();
				}
			}
		}
	};
	CChartSpace.prototype.getCatFormula = function () {
		var aAllSeries = this.getAllSeries();
		var oFirstSeries = aAllSeries[0];
		if (oFirstSeries) {
			return oFirstSeries.asc_getCatValues();
		}
		return "";
	};
	CChartSpace.prototype.setCatFormula = function (sFormula) {
		var aAllSeries = this.getAllSeries();
		for (var i = 0; i < aAllSeries.length; ++i) {
			aAllSeries[i].setCategories(sFormula);
		}
	};
	CChartSpace.prototype.onDataUpdateRecalc = function () {
		this.handleUpdateInternalChart();
		this.recalcInfo.recalculateReferences = true;
	};
	CChartSpace.prototype.onDataUpdate = function () {
		this.onDataUpdateRecalc();
		this.recalculate();
		this.onUpdate(null);
	};
	CChartSpace.prototype.setDLblsDeleteValue = function (bVal) {
		if (this.chart) {
			this.chart.setDLblsDeleteValue(bVal);
		}
	};
	CChartSpace.prototype.getChartType = function () {
		if (this.chart) {
			return this.chart.getChartType();
		}
		return Asc.c_oAscChartTypeSettings.unknown;
	};
	CChartSpace.prototype.isScatterChartType = function () {
		return AscFormat.isScatterChartType(this.getChartType());
	};
	CChartSpace.prototype.changeChartType = function (nType) {
		if (this.chart) {
			this.chart.changeChartType(nType);
			this.checkDlblsPosition();
			this.resetToChartStyleSoft();
		}
	};
	CChartSpace.prototype.canChangeToStockChart = function () {
		return this.chart.plotArea.canChangeToStockChart();
	};
	CChartSpace.prototype.canChangeToComboChart = function () {
		return this.chart.plotArea.canChangeToComboChart();
	};
	CChartSpace.prototype.checkDlblsPosition = function () {
		this.chart.checkDlblsPosition();
	};
	CChartSpace.prototype.getPossibleDLblsPosition = function () {
		var aPositions = this.chart.getPossibleDLblsPosition();
		if (aPositions.length === 0) {
			aPositions.push(Asc.c_oAscChartDataLabelsPos.show);
		}
		return aPositions;
	};
	CChartSpace.prototype.setDlblsProps = function (oProps) {
		this.chart.setDlblsProps(oProps);
	};
	CChartSpace.prototype.getOrderedAxes = function () {
		return this.chart.getOrderedAxes();
	};
	CChartSpace.prototype.is3dChart = function () {
		return AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(this);
	};
	CChartSpace.prototype.getTxPrParaPr = function () {
		if (this.txPr) {
			return this.txPr.getFirstParaParaPr();
		}
		return null;
	};
	CChartSpace.prototype.getTheme = function () {
		return this.getParentObjects().theme;
	};

	CChartSpace.prototype.getSpPrFormStyleEntry = function (oStyleEntry, aColors, nIdx) {
		return AscFormat.CBaseChartObject.prototype.getSpPrFormStyleEntry.call(this, oStyleEntry, aColors, nIdx);
	};
	CChartSpace.prototype.getTxPrFormStyleEntry = function (oStyleEntry, aColors, nIdx) {
		return AscFormat.CBaseChartObject.prototype.getTxPrFormStyleEntry.call(this, oStyleEntry, aColors, nIdx);
	};
	CChartSpace.prototype.applyStyleEntry = function (oStyleEntry, aColors, nIdx, bReset) {
		AscFormat.CBaseChartObject.prototype.applyStyleEntry.call(this, oStyleEntry, aColors, nIdx, bReset);
	};
	CChartSpace.prototype.resetFormatting = function () {
		AscFormat.CBaseChartObject.prototype.resetFormatting.call(this);
	};
	CChartSpace.prototype.resetOwnFormatting = function () {
		AscFormat.CBaseChartObject.prototype.resetOwnFormatting.call(this);
	};
	CChartSpace.prototype.applyChartStyle = function (oChartStyle, oColors, oAdditionalData, bReset) {
		this.applyStyleEntry(oChartStyle.chartArea, oColors.generateColors(1), 0, bReset);
		this.chart.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
	};
	CChartSpace.prototype.resetToChartStyle = function () {
		if (this.chartStyle && this.chartColors) {
			this.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), true);
		}
	};
	CChartSpace.prototype.resetToChartStyleSoft = function () {
		if (this.chartStyle && this.chartColors) {
			this.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), false);
		}
	};
	CChartSpace.prototype.checkElementChartStyle = function (oElement) {
		if (oElement && this.chartStyle && this.chartColors) {
			oElement.applyChartStyle(this.chartStyle, this.chartColors, oChartStyleCache.getAdditionalData(this.getChartType(), this.chartStyle.id), false);
		}
	};
	CChartSpace.prototype.getChildren = function () {
		return [this.chart];
	};
	CChartSpace.prototype.applyChartStyleByIds = function (aId) {
		var aChartStyle = oChartStyleCache.getStyleObject(aId);
		if (Array.isArray(aChartStyle)) {
			this.setChartStyle(aChartStyle[0].createDuplicate());
			this.setChartColors(aChartStyle[1].createDuplicate());
			this.resetToChartStyle();
		}
	};
	CChartSpace.prototype.getTypeName = function () {
		return AscCommon.translateManager.getValue("Chart");
	};
	CChartSpace.prototype.getAllAxes = function () {
		var oChart = this.chart;
		if (oChart) {
			var oPlotArea = oChart.plotArea;
			if (oPlotArea) {
				return oPlotArea.axId;
			}
		}
		return [];
	};
	CChartSpace.prototype.correctAxes = function () {
		var aAxes = this.getAllAxes();
		for (var nAx1 = 0; nAx1 < aAxes.length; ++nAx1) {
			var oAx1 = aAxes[nAx1];
			if (oAx1.crossAx === oAx1) {
				for (var nAx2 = 0; nAx2 < aAxes.length; ++nAx2) {
					var oAx2 = aAxes[nAx2];
					if (oAx2 !== oAx1 && oAx2.crossAx === oAx1) {
						oAx1.setCrossAx(oAx2);
						break;
					}
				}
			}
		}
	};
	CChartSpace.prototype.SetValuesToDataPoints = function (aData, nSeria) {
		if (editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			return false;

		var oSeria, isOkey;
		var allSeries = this.getAllSeries();

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria)
			return false;

		var oVal = this.isScatterChartType() ? oSeria.yVal : oSeria.val;
		if (oVal && oVal.numRef && oVal.numRef.numCache) {
			var oPt, nValue;
			for (var nPt = 0; nPt < aData.length; nPt++) {
				nValue = aData[nPt];
				if (typeof (nValue) !== "number")
					continue;

				oPt = oVal.numRef.numCache.pts[nPt];
				if (oPt) {
					oPt.setVal(nValue);
					isOkey = true;
				} else
					break;
			}
		}

		if (isOkey) {
			this.onDataUpdate();
			return true;
		}

		return false;
	};
	CChartSpace.prototype.SetXValuesToDataPoints = function (aData) {
		if (editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			return false;

		var oSeria, isOkey;
		var allSeries = this.getAllSeries();

		for (var nSeria = 0; nSeria < allSeries.length; nSeria++) {
			oSeria = allSeries[nSeria];
			var oVal = oSeria.xVal;
			if (oVal && oVal.strRef && oVal.strRef.strCache) {
				var oPt, value;
				for (var nPt = 0; nPt < aData.length; nPt++) {
					value = aData[nPt];
					if (typeof (value) === "number")
						value = String(value);

					if (typeof (value) !== "string")
						continue;

					oPt = oVal.strRef.strCache.pts[nPt];
					if (oPt) {
						oPt.setVal(value);
						isOkey = true;
					} else
						break;
				}
			}
		}

		if (isOkey) {
			this.onDataUpdate();
			return true;
		}

		return false;
	};
	CChartSpace.prototype.SetSeriaValues = function (sRange, nSeria) {
		var oSeria;
		var allSeries = this.getAllSeries();

		if (typeof (sRange) !== "string" || sRange === "")
			return false;

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria)
			return false;

		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet) {
			if (oSeria.asc_IsValidValues(sRange) !== 0)
				return false;
			oSeria.asc_setValues(sRange);
			return true;
		}
		// пока только для Spreadsheet
		return false;
	};
	CChartSpace.prototype.SetSeriaXValues = function (sRange, nSeria) {
		var oSeria;
		var allSeries = this.getAllSeries();

		if (typeof (sRange) !== "string" || sRange === "")
			return false;

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria)
			return false;

		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet) {
			if (oSeria.asc_IsValidXValues(sRange) !== 0)
				return false;
			oSeria.asc_setXValues(sRange);
			return true;
		}
		// пока только для Spreadsheet
		return false;
	};
	CChartSpace.prototype.SetSeriaName = function (sName, nSeria) {
		var allSeries = this.getAllSeries();
		var oSeria;

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria || typeof (sName) !== "string")
			return false;

		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet) {
			if (oSeria.asc_IsValidName(sName) !== 0)
				return false;
			oSeria.asc_setName(sName);
			return true;
		}

		var oTx = oSeria.tx;
		if (oTx && oTx.strRef && oTx.strRef.strCache) {
			var oPt;
			oPt = oTx.strRef.strCache.pts[0];
			if (!oPt)
				return false;
			oPt.setVal(sName);
		}
		return true;
	};
	CChartSpace.prototype.SetCatName = function (sName, nCategory) {
		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			return false;

		var oSeria, oCat, oPt, isOkey;
		if (typeof (sName) !== "string")
			return false;

		var allSeries = this.getAllSeries();

		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeria = allSeries[nCurSeries];
			oCat = oSeria.cat;
			if (oCat && oCat.strRef && oCat.strRef.strCache) {
				for (var nPt = 0; nPt < oCat.strRef.strCache.pts.length; nPt++) {
					oPt = oCat.strRef.strCache.pts[nPt];
					if (oPt.idx === nCategory) {
						oPt.setVal(sName);
						isOkey = true;
					}
				}
			}
		}

		if (isOkey)
			return true;

		return false;
	};
	CChartSpace.prototype.SetCatFormula = function (sRange) {
		if (typeof (sRange) !== "string")
			return false;

		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet) {
			var isValid = AscFormat.ExecuteNoHistory(function () {
				var oCat = new AscFormat.CCat();
				return oCat.setValues(sRange).getError();
			}, this, []);

			if (isValid === 0) {
				this.setCatFormula(sRange);
				return true;
			}
		}

		// пока только для Spreadsheet
		return false;
	};
	CChartSpace.prototype.RemoveSeria = function (nSeria) {
		var oSeria;
		var allSeries = this.getAllSeries();

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria)
			return false;

		oSeria.asc_Remove();
		return true;
	};

	CChartSpace.prototype.SetPlotAreaFill = function (oFill) {
		if (!this.chart.plotArea.spPr) {
			this.chart.plotArea.setSpPr(new AscFormat.CSpPr());
			this.chart.plotArea.spPr.setParent(this.chart.plotArea);
		}

		this.chart.plotArea.spPr.setFill(oFill);
	};
	CChartSpace.prototype.SetPlotAreaOutLine = function (oLn) {
		if (!this.chart.plotArea.spPr) {
			this.chart.plotArea.setSpPr(new AscFormat.CSpPr());
			this.chart.plotArea.spPr.setParent(this.chart.plotArea);
		}

		this.chart.plotArea.spPr.setLn(oLn);
	};
	CChartSpace.prototype.SetSeriesFill = function (oFill, nSeries, bAll) {
		var allSeries = this.getAllSeries();
		var oSeries, isOk;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeries = allSeries[nCurSeries];
			if (!bAll && oSeries.idx !== nSeries)
				continue;

			if (!oSeries.spPr) {
				oSeries.setSpPr(new AscFormat.CSpPr());
				oSeries.spPr.setParent(oSeries);
			}

			oSeries.spPr.setFill(oFill);
			if (Array.isArray(oSeries.dPt)) {
				for (var i = 0; i < oSeries.dPt.length; ++i) {
					if (oSeries.dPt[i].spPr) {
						oSeries.dPt[i].spPr.setFill(oFill.createDuplicate());
					}
				}
			}

			isOk = true;
			if (isOk && !bAll)
				return true;
		}

		if (isOk)
			return true;

		return false;
	};
	CChartSpace.prototype.SetSeriesOutLine = function (oLn, nSeries, bAll) {
		var allSeries = this.getAllSeries();
		var oSeries, isOk;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeries = allSeries[nCurSeries];
			if (!bAll && oSeries.idx !== nSeries)
				continue;

			if (!oSeries.spPr) {
				oSeries.setSpPr(new AscFormat.CSpPr());
				oSeries.spPr.setParent(oSeries);
			}

			oSeries.spPr.setLn(oLn);
			if (Array.isArray(oSeries.dPt)) {
				for (var i = 0; i < oSeries.dPt.length; ++i) {
					if (oSeries.dPt[i].spPr) {
						oSeries.dPt[i].spPr.setLn(oLn.createDuplicate());
					}
				}
			}

			isOk = true;
			if (isOk && !bAll)
				return true;
		}


		if (isOk)
			return true;

		return false;
	};
	CChartSpace.prototype.SetDataPointFill = function (oFill, nSeries, nDataPoint, bAllSeries) {
		var allSeries = this.getAllSeries();
		var oDataPoint, oSeries, isOk;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeries = allSeries[nCurSeries];
			if (!bAllSeries && oSeries.idx !== nSeries)
				continue;

			var pts = oSeries.getNumPts();
			if (nDataPoint >= pts.length || nDataPoint < 0)
				continue;

			oDataPoint = oSeries.dPt[nDataPoint];
			if (!oDataPoint) {
				oDataPoint = new AscFormat.CDPt();
				oDataPoint.setIdx(nDataPoint);
				oSeries.addDPt(oDataPoint);
				if (!oDataPoint.spPr) {
					oDataPoint.setSpPr(new AscFormat.CSpPr());
					oDataPoint.spPr.setParent(oDataPoint);
				}
			}
			oDataPoint.spPr.setFill(oFill);

			isOk = true;
			if (isOk && !bAllSeries)
				return true;
		}

		if (isOk)
			return true;

		return false;
	};
	CChartSpace.prototype.SetDataPointOutLine = function (oLn, nSeries, nDataPoint, bAllSeries) {
		var allSeries = this.getAllSeries();
		var oDataPoint, oSeries, isOk;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeries = allSeries[nCurSeries];
			if (!bAllSeries && oSeries.idx !== nSeries)
				continue;

			var pts = oSeries.getNumPts();
			if (nDataPoint >= pts.length || nDataPoint < 0)
				continue;

			oDataPoint = oSeries.dPt[nDataPoint];
			if (!oDataPoint) {
				oDataPoint = new AscFormat.CDPt();
				oDataPoint.setIdx(nDataPoint);
				oSeries.addDPt(oDataPoint);
				if (!oDataPoint.spPr) {
					oDataPoint.setSpPr(new AscFormat.CSpPr());
					oDataPoint.spPr.setParent(oDataPoint);
				}
			}
			oDataPoint.spPr.setLn(oLn);

			isOk = true;
			if (isOk && !bAllSeries)
				return true;
		}

		if (isOk)
			return true;

		return false;
	};
	CChartSpace.prototype.SetMarkerFill = function (oFill, nSeries, nMarker, bAllMarkers) {
		var allSeries = this.getAllSeries();
		var oDataPoint, oSeries;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			if (allSeries[nCurSeries].idx === nSeries) {
				oSeries = allSeries[nCurSeries];
				break;
			}
		}

		if (!oSeries)
			return false;

		if (bAllMarkers) {
			if (oSeries.marker) {
				if (!oSeries.marker.spPr) {
					oSeries.marker.setSpPr(new AscFormat.CSpPr());
					oSeries.marker.spPr.setParent(oSeries.marker);
				}
				oSeries.marker.spPr.setFill(oFill);
				if (Array.isArray(oSeries.dPt)) {
					for (var i = 0; i < oSeries.dPt.length; ++i) {
						if (oSeries.dPt[i].marker && oSeries.dPt[i].marker.spPr)
							oSeries.dPt[i].marker.spPr.setFill(oFill.createDuplicate());
					}
				}
				return true;
			}
		} else {
			var pts = oSeries.getNumPts();
			var oPt;
			for (var nPt = 0; nPt < pts.length; nPt++) {
				if (pts[nPt].idx === nMarker) {
					oPt = pts[nPt];
					break;
				}
			}

			if (!oPt)
				return false;

			if (Array.isArray(oSeries.dPt)) {
				for (var k = 0; k < oSeries.dPt.length; ++k) {
					if (oSeries.dPt[k].idx === nMarker) {
						oDataPoint = oSeries.dPt[k];
						break;
					}
				}
			}

			if (!oDataPoint) {
				oDataPoint = new AscFormat.CDPt();
				oDataPoint.setIdx(nMarker);
				oSeries.addDPt(oDataPoint);
			}

			if (oPt.compiledMarker) {
				oDataPoint.setMarker(oPt.compiledMarker.createDuplicate());
				if (!oDataPoint.marker.spPr) {
					oDataPoint.marker.setSpPr(new AscFormat.CSpPr());
					oDataPoint.marker.spPr.setParent(oDataPoint.marker);
				}
				oDataPoint.marker.spPr.setFill(oFill);
				return true;
			}
		}

		return false;
	};
	CChartSpace.prototype.SetMarkerOutLine = function (oLn, nSeries, nMarker, bAllMarkers) {
		var allSeries = this.getAllSeries();
		var oDataPoint, oSeries;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			if (allSeries[nCurSeries].idx === nSeries) {
				oSeries = allSeries[nCurSeries];
				break;
			}
		}

		if (!oSeries)
			return false;

		if (bAllMarkers) {
			if (oSeries.marker) {
				if (!oSeries.marker.spPr) {
					oSeries.marker.setSpPr(new AscFormat.CSpPr());
					oSeries.marker.spPr.setParent(oSeries.marker);
				}
				oSeries.marker.spPr.setLn(oLn);
				if (Array.isArray(oSeries.dPt)) {
					for (var i = 0; i < oSeries.dPt.length; ++i) {
						if (oSeries.dPt[i].marker && oSeries.dPt[i].marker.spPr)
							oSeries.dPt[i].marker.spPr.setLn(oLn.createDuplicate());
					}
				}
				return true;
			}
		} else {
			var pts = oSeries.getNumPts();
			var oPt;
			for (var nPt = 0; nPt < pts.length; nPt++) {
				if (pts[nPt].idx === nMarker) {
					oPt = pts[nPt];
					break;
				}
			}

			if (!oPt)
				return false;

			if (Array.isArray(oSeries.dPt)) {
				for (var k = 0; k < oSeries.dPt.length; ++k) {
					if (oSeries.dPt[k].idx === nMarker) {
						oDataPoint = oSeries.dPt[k];
						break;
					}
				}
			}

			if (!oDataPoint) {
				oDataPoint = new AscFormat.CDPt();
				oDataPoint.setIdx(nMarker);
				oSeries.addDPt(oDataPoint);
			}

			if (oPt.compiledMarker) {
				oDataPoint.setMarker(oPt.compiledMarker.createDuplicate());
				if (!oDataPoint.marker.spPr) {
					oDataPoint.marker.setSpPr(new AscFormat.CSpPr());
					oDataPoint.marker.spPr.setParent(oDataPoint.marker);
				}
				oDataPoint.marker.spPr.setLn(oLn);
				return true;
			}
		}

		return false;
	};
	CChartSpace.prototype.SetTitleFill = function (oFill) {
		if (this.chart.title) {
			if (!this.chart.title.spPr) {
				this.chart.title.setSpPr(new AscFormat.CSpPr());
				this.chart.title.spPr.setParent(this.chart.title);
			}

			this.chart.title.spPr.setFill(oFill);
			return true;
		}

		return false;
	};
	CChartSpace.prototype.SetTitleOutLine = function (oLn) {
		if (this.chart.title) {
			if (!this.chart.title.spPr) {
				this.chart.title.setSpPr(new AscFormat.CSpPr());
				this.chart.title.spPr.setParent(this.chart.title);
			}

			this.chart.title.spPr.setLn(oLn);
			return true;
		}

		return false;
	};
	CChartSpace.prototype.SetLegendFill = function (oFill) {
		if (this.chart.legend) {
			if (!this.chart.legend.spPr) {
				this.chart.legend.setSpPr(new AscFormat.CSpPr());
				this.chart.legend.spPr.setParent(this.chart.legend);
			}

			this.chart.legend.spPr.setFill(oFill);
			return true;
		}

		return false;
	};
	CChartSpace.prototype.SetLegendOutLine = function (oLn) {
		if (this.chart.legend) {
			if (!this.chart.legend.spPr) {
				this.chart.legend.setSpPr(new AscFormat.CSpPr());
				this.chart.legend.spPr.setParent(this.chart.legend);
			}

			this.chart.legend.spPr.setLn(oLn);
			return true;
		}

		return false;
	};

	CChartSpace.prototype.SetAxieNumFormat = function (sNumFormat, nAxiePos) {
		if (this.chart.plotArea.axId) {
			var oAxie;
			for (var nAxie = 0; nAxie < this.chart.plotArea.axId.length; nAxie++) {
				oAxie = this.chart.plotArea.axId[nAxie];
				if (oAxie instanceof AscFormat.CValAx && oAxie.axPos === nAxiePos) {
					oAxie.numFmt.setFormatCode(sNumFormat);
					if (sNumFormat !== "General")
						oAxie.numFmt.setSourceLinked(false);
					else
						oAxie.numFmt.setSourceLinked(true);

					return true;
				}
			}

		}
		return false;
	};

	CChartSpace.prototype.SetSeriaNumFormat = function (sNumFormat, nSeria) {
		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			return false;

		var allSeries = this.getAllSeries();
		var oSeria;

		for (var nCurSeria = 0; nCurSeria < allSeries.length; nCurSeria++) {
			if (allSeries[nCurSeria].idx === nSeria) {
				oSeria = allSeries[nCurSeria];
				break;
			}
		}

		if (!oSeria || typeof (sNumFormat) !== "string")
			return false;

		var oNumLit = oSeria.getNumLit();
		oNumLit.setFormatCode(sNumFormat);

		for (var nPt = 0; nPt < oNumLit.pts.length; nPt++)
			oNumLit.pts[nPt].setFormatCode(null);

		return true;
	};

	CChartSpace.prototype.SetDataPointNumFormat = function (sFormat, nSeria, nDataPoint, bAllSeries) {
		if (Asc.editor && Asc.editor.editorId === AscCommon.c_oEditorId.Spreadsheet)
			return false;

		var allSeries = this.getAllSeries();
		var oSeries, isOk;
		for (var nCurSeries = 0; nCurSeries < allSeries.length; nCurSeries++) {
			oSeries = allSeries[nCurSeries];
			if (!bAllSeries && oSeries.idx !== nSeria)
				continue;

			var pts = oSeries.getNumPts();
			if (nDataPoint >= pts.length || nDataPoint < 0)
				continue;

			pts[nDataPoint].setFormatCode(sFormat);
			isOk = true;

			if (isOk && !bAllSeries)
				return true;
		}

		if (isOk)
			return true;

		return false;
	};

	CChartSpace.prototype.pasteFormatting = function (oFormatData) {
		if(!oFormatData)
			return;
		let oTextPr = oFormatData.TextPr;
		if(!oTextPr) {
			return;
		}
		let oTitle = this.getChartTitle();
		if(oTitle) {
			oTitle.checkDocContent();
			let fDocContentMethod = AscCommonWord.CDocumentContent.prototype.PasteFormatting;
			let fTableMethod = AscCommonWord.CTable.prototype.PasteFormatting;
			oTitle.applyTextFunction(fDocContentMethod, fTableMethod, [oFormatData]);
		}
	};

	function CAdditionalStyleData() {
		this.dLbls = null;
		this.catAx = null;
		this.valAx = null;
		this.view3D = null;
		this.legend = null;
		this.gapWidth = null;
		this.overlap = null;
		this.gapDepth = null;
	}

	function fSaveChartObjectSourceFormatting(oObject, oObjectCopy, oTheme, oColorMap) {
		if (oObject === oObjectCopy || !oObjectCopy || !oObject) {
			return;
		}

		if (oObject.pen || oObject.brush) {
			if (oObject.pen || oObject.brush) {
				if (!oObjectCopy.spPr) {
					if (oObjectCopy.setSpPr) {
						oObjectCopy.setSpPr(new AscFormat.CSpPr());
						oObjectCopy.spPr.setParent(oObjectCopy);
					}
				}
				if (oObject.brush) {
					if (oTheme && oColorMap) {
						oObject.brush.check(oTheme, oColorMap);
					}
					oObjectCopy.spPr.setFill(oObject.brush.saveSourceFormatting());
				}
				if (oObject.pen) {
					if (oObject.pen.Fill && oTheme && oColorMap) {
						oObject.pen.Fill.check(oTheme, oColorMap);
					}
					oObjectCopy.spPr.setLn(oObject.pen.createDuplicate(true));
				}
			}
		}
		if (oObject.txPr && oObject.txPr.content) {
			AscFormat.SaveContentSourceFormatting(oObject.txPr.content.Content, oObjectCopy.txPr.content.Content, oTheme, oColorMap);
		}
		if (oObject.tx && oObject.tx.rich && oObject.tx.rich.content) {
			AscFormat.SaveContentSourceFormatting(oObject.tx.rich.content.Content, oObjectCopy.tx.rich.content.Content, oTheme, oColorMap);
		}
	}

	function getMaxIdx(arr) {
		var max_idx = 0;
		for (var i = 0; i < arr.length; ++i)
			arr[i].idx > max_idx && (max_idx = arr[i].idx);
		return max_idx + 1;
	}

	function getArrayFillsFromBase(arrBaseFills, needFillsCount) {
		var ret = [];
		var nMaxCycleIdx = parseInt((needFillsCount - 1) / arrBaseFills.length);
		for (var i = 0; i < needFillsCount; ++i) {
			var nCycleIdx = (i / arrBaseFills.length) >> 0;
			var fShadeTint = (nCycleIdx + 1) / (nMaxCycleIdx + 2) * 1.4 - 0.7;

			if (fShadeTint < 0) {
				fShadeTint = -(1 + fShadeTint);
			} else {
				fShadeTint = (1 - fShadeTint);
			}
			var color = CreateUniFillSolidFillWidthTintOrShade(arrBaseFills[i % arrBaseFills.length], fShadeTint);
			ret.push(color);
		}
		return ret;
	}

	function GetTypeMarkerByIndex(index) {
		return AscFormat.MARKER_SYMBOL_TYPE[index % 9];
	}

	function CreateUnfilFromRGB(r, g, b) {
		var ret = new AscFormat.CUniFill();
		ret.setFill(new AscFormat.CSolidFill());
		ret.fill.setColor(new AscFormat.CUniColor());
		ret.fill.color.setColor(new AscFormat.CRGBColor());
		ret.fill.color.color.setColor(r, g, b);
		return ret;
	}

	function CreateColorMapByIndex(index) {
		var ret = [];
		switch (index) {
			case 1: {
				ret.push(CreateUnfilFromRGB(85, 85, 85));
				ret.push(CreateUnfilFromRGB(158, 158, 158));
				ret.push(CreateUnfilFromRGB(114, 114, 114));
				ret.push(CreateUnfilFromRGB(70, 70, 70));
				ret.push(CreateUnfilFromRGB(131, 131, 131));
				ret.push(CreateUnfilFromRGB(193, 193, 193));
				break;
			}
			case 2: {
				for (var i = 0; i < 6; ++i) {
					ret.push(CreateUnifillSolidFillSchemeColorByIndex(i));
				}
				break;
			}
			default: {
				ret.push(CreateUnifillSolidFillSchemeColorByIndex(index - 3));
				break;
			}
		}
		return ret;
	}


	function CreateUniFillSolidFillWidthTintOrShade(unifill, effectVal) {
		var ret = unifill.createDuplicate();
		var unicolor = ret.fill.color;
		if (effectVal !== 0) {
			effectVal *= 100000.0;
			if (!unicolor.Mods)
				unicolor.setMods(new AscFormat.CColorModifiers());
			var mod = new AscFormat.CColorMod();
			if (effectVal > 0) {
				mod.setName("tint");
				mod.setVal(effectVal);
			} else {
				mod.setName("shade");
				mod.setVal(Math.abs(effectVal));
			}
			unicolor.Mods.addMod(mod);
		}
		return ret;
	}

	function CreateUnifillSolidFillSchemeColor(colorId, tintOrShade) {
		var unifill = new AscFormat.CUniFill();
		unifill.setFill(new AscFormat.CSolidFill());
		unifill.fill.setColor(new AscFormat.CUniColor());
		unifill.fill.color.setColor(new AscFormat.CSchemeColor());
		unifill.fill.color.color.setId(colorId);
		return CreateUniFillSolidFillWidthTintOrShade(unifill, tintOrShade);
	}

	function CreateNoFillLine() {
		var ret = new AscFormat.CLn();
		ret.setFill(CreateNoFillUniFill());
		return ret;
	}

	function CreateNoFillUniFill() {
		var ret = new AscFormat.CUniFill();
		ret.setFill(new AscFormat.CNoFill());
		return ret;
	}


	function CreateView3d(nRotX, nRotY, bRAngAx, nDepthPercent) {
		var oView3d = new AscFormat.CView3d();
		AscFormat.isRealNumber(nRotX) && oView3d.setRotX(nRotX);
		AscFormat.isRealNumber(nRotY) && oView3d.setRotY(nRotY);
		AscFormat.isRealBool(bRAngAx) && oView3d.setRAngAx(bRAngAx);
		AscFormat.isRealNumber(nDepthPercent) && oView3d.setDepthPercent(nDepthPercent);
		return oView3d;
	}

	function GetNumFormatFromSeries(aAscSeries) {
		if (aAscSeries &&
			aAscSeries[0] &&
			aAscSeries[0].Val &&
			aAscSeries[0].Val.NumCache &&
			aAscSeries[0].Val.NumCache[0]) {
			if (typeof (aAscSeries[0].Val.NumCache[0].numFormatStr) === "string" && aAscSeries[0].Val.NumCache[0].numFormatStr.length > 0) {
				return aAscSeries[0].Val.NumCache[0].numFormatStr;
			}
		}
		return "General";
	}


	function FillCatAxNumFormat(oCatAxis, aAscSeries) {
		if (!oCatAxis) {
			return;
		}
		var sFormatCode = null;
		if (aAscSeries &&
			aAscSeries[0] &&
			aAscSeries[0].Cat &&
			aAscSeries[0].Cat.NumCache &&
			aAscSeries[0].Cat.NumCache[0]) {
			if (typeof (aAscSeries[0].Cat.NumCache[0].numFormatStr) === "string" && aAscSeries[0].Cat.NumCache[0].numFormatStr.length > 0) {
				sFormatCode = aAscSeries[0].Cat.NumCache[0].numFormatStr;
			}
		}
		if (sFormatCode) {
			oCatAxis.setNumFmt(new AscFormat.CNumFmt());
			oCatAxis.numFmt.setFormatCode(sFormatCode ? sFormatCode : "General");
			oCatAxis.numFmt.setSourceLinked(true);
		}
	}


	function CreateTypedLineChart(type) {
		var oLineChart = new AscFormat.CLineChart();
		oLineChart.setGrouping(type);
		oLineChart.setVaryColors(false);
		oLineChart.setMarker(true);
		oLineChart.setSmooth(false);
		return oLineChart;
	}

	function CreateLineChart(chartSeries, type, bUseCache, oOptions, b3D) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		if (b3D) {
			chart.setView3D(CreateView3d(15, 20, true, 100));
			chart.setDefaultWalls();
		}
		chart.setAutoTitleDeleted(false);
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setPlotVisOnly(true);
		var disp_blanks_as;
		if (type === AscFormat.GROUPING_STANDARD) {
			disp_blanks_as = DISP_BLANKS_AS_GAP;
		} else {
			disp_blanks_as = DISP_BLANKS_AS_ZERO;
		}
		chart.setDispBlanksAs(disp_blanks_as);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(CreateTypedLineChart(type));

		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);

		var line_chart = plot_area.charts[0];
		line_chart.addAxId(cat_ax);
		line_chart.addAxId(val_ax);
		val_ax.setCrosses(2);

		var bMarker = false;
		var nChartType = oOptions.type;
		if (nChartType === Asc.c_oAscChartTypeSettings.lineNormalMarker ||
			nChartType === Asc.c_oAscChartTypeSettings.lineStackedMarker ||
			nChartType === Asc.c_oAscChartTypeSettings.lineStackedPerMarker) {
			bMarker = true;
		}
		bMarker = false;
		line_chart.setMarker(bMarker);
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CLineSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.setMarker(new AscFormat.CMarker());
			if (bMarker) {
				series.marker.setSymbol(GetTypeMarkerByIndex(i));
			} else {
				series.marker.setSymbol(AscFormat.SYMBOL_NONE);
			}
			series.setSmooth(false);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			line_chart.addSer(series);
		}
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		FillCatAxNumFormat(cat_ax, asc_series);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code;
		if (type === AscFormat.GROUPING_PERCENT_STACKED) {
			format_code = "0%";
		} else {
			format_code = GetNumFormatFromSeries(asc_series);
		}
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);


		if (asc_series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}

		chart_space.printSettings.setDefault();
		line_chart.tryChangeType(oOptions.type);
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateTypedBarChart(type, b3D) {
		var oBarChart = new AscFormat.CBarChart();
		if (b3D) {
			oBarChart.set3D(true);
		}
		oBarChart.setBarDir(AscFormat.BAR_DIR_COL);
		oBarChart.setGrouping(type);
		oBarChart.setVaryColors(false);
		oBarChart.setGapWidth(150);
		if (AscFormat.BAR_GROUPING_PERCENT_STACKED === type || AscFormat.BAR_GROUPING_STACKED === type) {
			oBarChart.setOverlap(100);
		}
		if (b3D) {
			oBarChart.setShape(BAR_SHAPE_BOX);
		}
		return oBarChart;
	}

	function CreateBarChart(chartSeries, type, bUseCache, oOptions, b3D, bDepth) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		if (b3D) {
			chart.setView3D(CreateView3d(15, 20, true, bDepth ? 100 : undefined));
			chart.setDefaultWalls();
		}
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setPlotVisOnly(true);
		chart.setDispBlanksAs(DISP_BLANKS_AS_GAP);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(CreateTypedBarChart(type, b3D));

		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);

		var bar_chart = plot_area.charts[0];
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CBarSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.setInvertIfNegative(false);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			bar_chart.addSer(series);
		}
		bar_chart.addAxId(cat_ax);
		bar_chart.addAxId(val_ax);
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		FillCatAxNumFormat(cat_ax, asc_series);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		var num_fmt = val_ax.numFmt;
		var format_code;
		if (type === AscFormat.BAR_GROUPING_PERCENT_STACKED) {
			format_code = "0%";
		} else {
			format_code = GetNumFormatFromSeries(asc_series);
		}
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		if (bDepth) {
			var oSerAx = new AscFormat.CSerAx();
			oSerAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
			var oSerScaling = new AscFormat.CScaling();
			oSerScaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
			oSerAx.setScaling(oSerScaling);
			oSerAx.setDelete(false);
			oSerAx.setAxPos(AscFormat.AX_POS_B);
			oSerAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
			oSerAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
			oSerAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
			oSerAx.setCrossAx(cat_ax);
			oSerAx.setCrosses(AscFormat.CROSSES_AUTO_ZERO);

			plot_area.addAxis(oSerAx);
			bar_chart.addAxId(oSerAx);

		}
		if (asc_series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}
		chart_space.printSettings.setDefault();
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateHBarChart(chartSeries, type, bUseCache, oOptions, b3D) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		if (b3D) {
			chart.setView3D(CreateView3d(15, 20, true, undefined));
			chart.setDefaultWalls();
		}
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setPlotVisOnly(true);
		chart.setDispBlanksAs(DISP_BLANKS_AS_GAP);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		var oBarChart = new AscFormat.CBarChart();
		plot_area.addChart(oBarChart);
		if (b3D) {
			oBarChart.set3D(true);
		}

		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);

		var bar_chart = plot_area.charts[0];
		bar_chart.setBarDir(AscFormat.BAR_DIR_BAR);
		bar_chart.setGrouping(type);
		bar_chart.setVaryColors(false);

		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CBarSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.setInvertIfNegative(false);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			bar_chart.addSer(series);
			if (b3D) {
				bar_chart.setShape(BAR_SHAPE_BOX);
			}
		}
		//bar_chart.setDLbls(new CDLbls());
		//var d_lbls = bar_chart.dLbls;
		//d_lbls.setShowLegendKey(false);
		//d_lbls.setShowVal(true);
		bar_chart.setGapWidth(150);
		if (AscFormat.BAR_GROUPING_PERCENT_STACKED === type || AscFormat.BAR_GROUPING_STACKED === type)
			bar_chart.setOverlap(100);
		bar_chart.addAxId(cat_ax);
		bar_chart.addAxId(val_ax);
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_L);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		FillCatAxNumFormat(cat_ax, asc_series);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_B);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code;
		/*if(type === GROUPING_PERCENT_STACKED)
         {
         format_code = "0%";
         }
         else */
		{
			format_code = GetNumFormatFromSeries(asc_series);
		}
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);

		if (asc_series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}
		chart_space.printSettings.setDefault();
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateAreaChart(chartSeries, type, bUseCache, oOptions) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setPlotVisOnly(true);
		chart.setDispBlanksAs(DISP_BLANKS_AS_ZERO);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(new AscFormat.CAreaChart());

		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);
		var area_chart = plot_area.charts[0];
		area_chart.setGrouping(type);
		area_chart.setVaryColors(false);
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CAreaSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			area_chart.addSer(series);
		}
		//area_chart.setDLbls(new CDLbls());
		area_chart.addAxId(cat_ax);
		area_chart.addAxId(val_ax);
		//var d_lbls = area_chart.dLbls;
		//d_lbls.setShowLegendKey(false);
		//d_lbls.setShowVal(true);
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		cat_ax.scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		FillCatAxNumFormat(cat_ax, asc_series);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_MID_CAT);
		var scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code;
		if (type === AscFormat.GROUPING_PERCENT_STACKED) {
			format_code = "0%";
		} else {
			format_code = GetNumFormatFromSeries(asc_series);
		}
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);

		if (asc_series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}
		chart_space.printSettings.setDefault();

		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreatePieChart(chartSeries, bDoughnut, bUseCache, oOptions, b3D) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setStyle(2);
		chart_space.setChart(new AscFormat.CChart());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		if (b3D) {
			chart.setView3D(CreateView3d(30, 0, true, 100));
			chart.setDefaultWalls();
		}
		chart.setPlotArea(new AscFormat.CPlotArea());
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(bDoughnut ? new AscFormat.CDoughnutChart() : new AscFormat.CPieChart());
		var pie_chart = plot_area.charts[0];
		pie_chart.setVaryColors(true);
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CPieSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			pie_chart.addSer(series);
		}
		pie_chart.setFirstSliceAng(0);
		if (bDoughnut)
			pie_chart.setHoleSize(75);
		chart.setLegend(new AscFormat.CLegend());
		var legend = chart.legend;
		legend.setLegendPos(c_oAscChartLegendShowSettings.right);
		legend.setLayout(new AscFormat.CLayout());
		legend.setOverlay(false);
		chart.setPlotVisOnly(true);
		chart.setDispBlanksAs(DISP_BLANKS_AS_GAP);
		chart.setShowDLblsOverMax(false);
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		chart_space.printSettings.setDefault();
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateScatterChart(chartSeries, bUseCache, oOptions) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setStyle(2);
		chart_space.setChart(new AscFormat.CChart());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		chart.setPlotArea(new AscFormat.CPlotArea());
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(new AscFormat.CScatterChart());
		var scatter_chart = plot_area.charts[0];
		scatter_chart.setScatterStyle(AscFormat.SCATTER_STYLE_MARKER);
		scatter_chart.setVaryColors(false);
		var cat_ax = new AscFormat.CValAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);
		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);
		var oXVal;
		var first_series;
		var start_index = 0;
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CScatterSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			scatter_chart.addSer(series);
		}
		scatter_chart.addAxId(cat_ax);
		scatter_chart.addAxId(val_ax);
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		var scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code = GetNumFormatFromSeries(asc_series);
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);
		if (scatter_chart.series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		chart_space.printSettings.setDefault();
		if (oOptions && oOptions.type !== Asc.c_oAscChartTypeSettings) {
			scatter_chart.tryChangeType(oOptions.type);
		}
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateStockChart(chartSeries, bUseCache, oOptions) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setLegend(new AscFormat.CLegend());
		chart.setPlotVisOnly(true);
		var disp_blanks_as;
		disp_blanks_as = DISP_BLANKS_AS_GAP;
		chart.setDispBlanksAs(disp_blanks_as);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());
		plot_area.addChart(new AscFormat.CStockChart());
		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);

		var oStockChart = plot_area.charts[0];
		oStockChart.addAxId(cat_ax);
		oStockChart.addAxId(val_ax);
		oStockChart.setHiLowLines(new AscFormat.CSpPr());
		oStockChart.setUpDownBars(new AscFormat.CUpDownBars());
		oStockChart.upDownBars.setGapWidth(150);
		oStockChart.upDownBars.setUpBars(new AscFormat.CSpPr());
		oStockChart.upDownBars.setDownBars(new AscFormat.CSpPr());
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CLineSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.setMarker(new AscFormat.CMarker());
			series.setSpPr(new AscFormat.CSpPr());
			series.spPr.setLn(new AscFormat.CLn());
			series.spPr.ln.setW(28575);
			series.spPr.ln.setFill(CreateNoFillUniFill());
			series.marker.setSymbol(AscFormat.SYMBOL_NONE);
			series.setSmooth(false);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			oStockChart.addSer(series);
		}
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code;
		format_code = GetNumFormatFromSeries(asc_series);
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);
		var legend = chart.legend;
		legend.setLegendPos(c_oAscChartLegendShowSettings.right);
		legend.setLayout(new AscFormat.CLayout());
		legend.setOverlay(false);
		chart_space.printSettings.setDefault();
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateSurfaceAxes(sNumFormat) {
		var oCatAx = new AscFormat.CCatAx();
		oCatAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		var oScaling = new AscFormat.CScaling();
		oScaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		oCatAx.setScaling(oScaling);
		oCatAx.setDelete(false);
		oCatAx.setAxPos(AscFormat.AX_POS_B);
		oCatAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		oCatAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		oCatAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		oCatAx.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		oCatAx.setAuto(true);
		oCatAx.setLblAlgn(AscFormat.LBL_ALG_CTR);
		oCatAx.setLblOffset(100);
		oCatAx.setNoMultiLvlLbl(false);
		var oValAx = new AscFormat.CValAx();
		oValAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		var oValScaling = new AscFormat.CScaling();
		oValScaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		oValAx.setScaling(oValScaling);
		oValAx.setDelete(false);
		oValAx.setAxPos(AscFormat.AX_POS_L);
		oValAx.setMajorGridlines(new AscFormat.CSpPr());
		var oNumFmt = new AscFormat.CNumFmt();
		oNumFmt.setFormatCode(sNumFormat);
		oNumFmt.setSourceLinked(true);
		oValAx.setNumFmt(oNumFmt);
		oValAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		oValAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		oValAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		oCatAx.setCrossAx(oValAx);
		oValAx.setCrossAx(oCatAx);
		oValAx.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		oValAx.setCrossBetween(AscFormat.CROSS_BETWEEN_MID_CAT);
		var oSerAx = new AscFormat.CSerAx();
		oSerAx.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		var oSerScaling = new AscFormat.CScaling();
		oSerScaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		oSerAx.setScaling(oSerScaling);
		oSerAx.setDelete(false);
		oSerAx.setAxPos(AscFormat.AX_POS_B);
		oSerAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		oSerAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		oSerAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		oSerAx.setCrossAx(oCatAx);
		oSerAx.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		return {valAx: oValAx, catAx: oCatAx, serAx: oSerAx};
	}

	function CreateSurfaceChart(chartSeries, bUseCache, oOptions, bContour, bWireFrame) {
		var asc_series = chartSeries;
		var oChartSpace = new AscFormat.CChartSpace();
		oChartSpace.setDate1904(false);
		oChartSpace.setLang("en-US");
		oChartSpace.setRoundedCorners(false);
		oChartSpace.setStyle(2);
		oChartSpace.setChart(new AscFormat.CChart());
		var oChart = oChartSpace.chart;
		oChart.setAutoTitleDeleted(false);
		var oView3D = new AscFormat.CView3d();
		oChart.setView3D(oView3D);
		if (!bContour) {
			oView3D.setRotX(15);
			oView3D.setRotY(20);
			oView3D.setRAngAx(false);
			oView3D.setPerspective(30);
		} else {
			oView3D.setRotX(90);
			oView3D.setRotY(0);
			oView3D.setRAngAx(false);
			oView3D.setPerspective(0);
		}
		oChart.setFloor(new AscFormat.CChartWall());
		oChart.floor.setThickness(0);
		oChart.setSideWall(new AscFormat.CChartWall());
		oChart.sideWall.setThickness(0);
		oChart.setBackWall(new AscFormat.CChartWall());
		oChart.backWall.setThickness(0);
		oChart.setPlotArea(new AscFormat.CPlotArea());
		oChart.plotArea.setLayout(new AscFormat.CLayout());
		var oSurfaceChart;
		//if(bContour){
		oSurfaceChart = new AscFormat.CSurfaceChart();
		//}
		if (bWireFrame) {
			oSurfaceChart.setWireframe(true);
		} else {
			oSurfaceChart.setWireframe(false);
		}

		oChart.plotArea.addChart(oSurfaceChart);
		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CSurfaceSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.fillFromAscSeries(asc_series[i], bUseCache);
			oSurfaceChart.addSer(series);
		}
		var oAxes = CreateSurfaceAxes(GetNumFormatFromSeries(asc_series));
		var oCatAx = oAxes.catAx, oValAx = oAxes.valAx, oSerAx = oAxes.serAx;
		oChart.plotArea.addAxis(oCatAx);
		oChart.plotArea.addAxis(oValAx);
		oChart.plotArea.addAxis(oSerAx);
		oSurfaceChart.addAxId(oCatAx);
		oSurfaceChart.addAxId(oValAx);
		oSurfaceChart.addAxId(oSerAx);
		var oLegend = new AscFormat.CLegend();
		oLegend.setLegendPos(c_oAscChartLegendShowSettings.right);
		oLegend.setLayout(new AscFormat.CLayout());
		oLegend.setOverlay(false);
		//oLegend.setTxPr(AscFormat.CreateTextBodyFromString("", oDrawingDocument, oElement));
		oChart.setLegend(oLegend);
		oChart.setPlotVisOnly(true);
		oChart.setDispBlanksAs(DISP_BLANKS_AS_ZERO);
		oChart.setShowDLblsOverMax(false);
		var oPrintSettings = new AscFormat.CPrintSettings();
		oPrintSettings.setDefault();
		oChartSpace.setPrintSettings(oPrintSettings);
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			oChartSpace.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return oChartSpace;
	}


	function CreateRadarChart(chartSeries, bUseCache, oOptions, bMarker, bFilled) {
		var asc_series = chartSeries;
		var chart_space = new CChartSpace();
		chart_space.setDate1904(false);
		chart_space.setLang("en-US");
		chart_space.setRoundedCorners(false);
		chart_space.setChart(new AscFormat.CChart());
		chart_space.setPrintSettings(new AscFormat.CPrintSettings());
		var chart = chart_space.chart;
		chart.setAutoTitleDeleted(false);
		chart.setPlotArea(new AscFormat.CPlotArea());
		chart.setPlotVisOnly(true);
		var disp_blanks_as;
		disp_blanks_as = DISP_BLANKS_AS_GAP;
		chart.setDispBlanksAs(disp_blanks_as);
		chart.setShowDLblsOverMax(false);
		var plot_area = chart.plotArea;
		plot_area.setLayout(new AscFormat.CLayout());

		let oRadarChart = new AscFormat.CRadarChart();
		oRadarChart.setRadarStyle(bFilled ? AscFormat.RADAR_STYLE_FILLED : AscFormat.RADAR_STYLE_MARKER);
		plot_area.addChart(oRadarChart);

		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		plot_area.addAxis(cat_ax);
		plot_area.addAxis(val_ax);

		oRadarChart.addAxId(cat_ax);
		oRadarChart.addAxId(val_ax);
		val_ax.setCrosses(2);

		for (var i = 0; i < asc_series.length; ++i) {
			var series = new AscFormat.CRadarSeries();
			series.setIdx(i);
			series.setOrder(i);
			series.setMarker(new AscFormat.CMarker());
			if (bMarker) {
				series.marker.setSymbol(GetTypeMarkerByIndex(i));
			} else {
				series.marker.setSymbol(AscFormat.SYMBOL_NONE);
			}
			series.fillFromAscSeries(asc_series[i], bUseCache);
			oRadarChart.addSer(series);
		}
		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		FillCatAxNumFormat(cat_ax, asc_series);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code;
		format_code = GetNumFormatFromSeries(asc_series);
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);
		if (asc_series.length > 1) {
			chart.setLegend(new AscFormat.CLegend());
			var legend = chart.legend;
			legend.setLegendPos(c_oAscChartLegendShowSettings.right);
			legend.setLayout(new AscFormat.CLayout());
			legend.setOverlay(false);
		}
		chart_space.printSettings.setDefault();
		if (AscCommon.g_oChartStyles[oOptions.type]) {
			chart_space.applyChartStyleByIds(AscCommon.g_oChartStyles[oOptions.type][0]);
		}
		return chart_space;
	}

	function CreateDefaultAxes(valFormatCode) {
		var cat_ax = new AscFormat.CCatAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);


		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.setAuto(true);
		cat_ax.setLblAlgn(AscFormat.LBL_ALG_CTR);
		cat_ax.setLblOffset(100);
		cat_ax.setNoMultiLvlLbl(false);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		var scaling = cat_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		var num_fmt = val_ax.numFmt;
		num_fmt.setFormatCode(valFormatCode);
		num_fmt.setSourceLinked(true);
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		//cat_ax.setTitle(new CTitle());
		//val_ax.setTitle(new CTitle());
		// var title = val_ax.title;
		// title.setTxPr(new CTextBody());
		// title.txPr.setBodyPr(new AscFormat.CBodyPr());
		// title.txPr.bodyPr.setVert(AscFormat.nVertTTvert);
		return {valAx: val_ax, catAx: cat_ax};
	}

	function CreateScatterAxis() {
		var cat_ax = new AscFormat.CValAx();
		var val_ax = new AscFormat.CValAx();
		cat_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		val_ax.setAxId(++Ax_Counter.GLOBAL_AX_ID_COUNTER);
		cat_ax.setCrossAx(val_ax);
		val_ax.setCrossAx(cat_ax);

		cat_ax.setScaling(new AscFormat.CScaling());
		cat_ax.setDelete(false);
		cat_ax.setAxPos(AscFormat.AX_POS_B);
		cat_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		cat_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		cat_ax.scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		val_ax.setScaling(new AscFormat.CScaling());
		val_ax.setDelete(false);
		val_ax.setAxPos(AscFormat.AX_POS_L);
		val_ax.setMajorGridlines(new AscFormat.CSpPr());
		val_ax.setNumFmt(new AscFormat.CNumFmt());
		val_ax.setMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		val_ax.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax.setCrosses(AscFormat.CROSSES_AUTO_ZERO);
		val_ax.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
		var scaling = val_ax.scaling;
		scaling.setOrientation(AscFormat.ORIENTATION_MIN_MAX);
		var num_fmt = val_ax.numFmt;
		var format_code = "General";
		num_fmt.setFormatCode(format_code);
		num_fmt.setSourceLinked(true);
		return {valAx: val_ax, catAx: cat_ax};
	}

	function getChartSeries(options) {
		var sRange = options.getRange();
		if (typeof sRange !== "string") {
			return [];
		}
		var bInColumns = options.getInColumns();
		var bHorValues = null;
		if (bInColumns === true || bInColumns === false) {
			bHorValues = !bInColumns;
		}
		var bScatter = AscFormat.isScatterChartType(options.getType());
		var oDataRange = new AscFormat.CChartDataRefs(null);
		var aSeriesRefs = oDataRange.getSeriesRefsFromUnionRefs(AscFormat.fParseChartFormulaExternal(sRange), bHorValues, bScatter);
		if (!Array.isArray(aSeriesRefs)) {
			return [];
		}
		var aSeries = [];
		for (var nRef = 0; nRef < aSeriesRefs.length; ++nRef) {
			aSeries.push(aSeriesRefs[nRef].getAscSeries());
		}
		return aSeries;
	}


	function checkSpPrRasterImages(spPr) {
		if (spPr && spPr.Fill && spPr.Fill && spPr.Fill.fill && spPr.Fill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP) {
			var copyBlipFill = spPr.Fill.createDuplicate();
			copyBlipFill.fill.setRasterImageId(spPr.Fill.fill.RasterImageId);
			spPr.setFill(copyBlipFill);
		}
	}

	function checkBlipFillRasterImages(sp) {
		switch (sp.getObjectType()) {
			case AscDFH.historyitem_type_Shape: {
				checkSpPrRasterImages(sp.spPr);
				break;
			}
			case AscDFH.historyitem_type_ImageShape:
			case AscDFH.historyitem_type_OleObject: {
				if (sp.blipFill) {
					var newBlipFill = sp.blipFill.createDuplicate();
					newBlipFill.setRasterImageId(sp.blipFill.RasterImageId);
					sp.setBlipFill(newBlipFill);
				}
				break;
			}
			case AscDFH.historyitem_type_ChartSpace: {
				checkSpPrRasterImages(sp.spPr);
				var chart = sp.chart;
				if (chart) {
					chart.backWall && checkSpPrRasterImages(chart.backWall.spPr);
					chart.floor && checkSpPrRasterImages(chart.floor.spPr);
					chart.legend && checkSpPrRasterImages(chart.legend.spPr);
					chart.sideWall && checkSpPrRasterImages(chart.sideWall.spPr);
					chart.title && checkSpPrRasterImages(chart.title.spPr);
					//plotArea
					var plot_area = sp.chart.plotArea;
					if (plot_area) {
						checkSpPrRasterImages(plot_area.spPr);
						for (var j = 0; j < plot_area.axId.length; ++j) {
							var axis = plot_area.axId[j];
							if (axis) {
								checkSpPrRasterImages(axis.spPr);
								axis.title && axis.title && checkSpPrRasterImages(axis.title.spPr);
							}
						}
						for (j = 0; j < plot_area.charts.length; ++j) {
							plot_area.charts[j].checkSpPrRasterImages();
						}
					}
				}
				break;
			}
			case AscDFH.historyitem_type_GroupShape:
			case AscDFH.historyitem_type_SmartArt:
			case AscDFH.historyitem_type_Drawing: {
				for (var i = 0; i < sp.spTree.length; ++i) {
					checkBlipFillRasterImages(sp.spTree[i]);
				}
				break;
			}
			case AscDFH.historyitem_type_GraphicFrame: {
				break;
			}
		}
	}

	function initStyleManager() {
		CHART_STYLE_MANAGER.init();
	}


	function CChartStyleCache() {
		this.cachedStyles = {};
		this.cachedColors = {};
		this.cachedDLbls = {};
		this.cachedCatAx = {};
		this.cachedValAx = {};
		this.cachedView3d = {};
		this.cachedLegend = {};
	}

	CChartStyleCache.prototype.createBinaryReader = function (sBinary) {
		var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		var oStream = oBinaryFileReader.getbase64DecodedData(sBinary);
		AscCommon.pptx_content_loader.Clear(true);
		return new AscCommon.BinaryChartReader(oStream);
	};
	CChartStyleCache.prototype.checkStyle = function (nId) {
		if (this.cachedStyles[nId] && !this.cachedStyles[nId].dataPoint) {
			this.cachedStyles[nId] = undefined;
		}
		if (!this.cachedStyles[nId]) {
			var sStyleBinary = AscCommon.g_oStylesBinaries[nId];
			if (sStyleBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sStyleBinary);
					var oNewVal = new AscFormat.CChartStyle();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_ChartStyle(t, l, oNewVal);
					});
					this.cachedStyles[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedStyles[nId];
	};
	CChartStyleCache.prototype.checkColors = function (nId) {
		if (!this.cachedStyles[nId]) {
			var sColorsBinary = AscCommon.g_oColorsBinaries[nId];
			if (sColorsBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sColorsBinary);
					var oNewVal = new AscFormat.CChartColors();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_ChartColors(t, l, oNewVal);
					});
					this.cachedColors[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedColors[nId];
	};
	CChartStyleCache.prototype.checkDataLabels = function (nId) {
		if (!this.cachedDLbls[nId]) {
			var sBinary = AscCommon.g_oDataLabelsBinaries[nId];
			if (sBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sBinary);
					var oNewVal = new AscFormat.CDLbls();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_DLbls(t, l, oNewVal);
					});
					this.cachedDLbls[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedDLbls[nId];
	};
	CChartStyleCache.prototype.checkCatAx = function (nId) {
		if (!this.cachedCatAx[nId]) {
			var sBinary = AscCommon.g_oCatBinaries[nId];
			if (sBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sBinary);
					var oNewVal = new AscFormat.CCatAx();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_CatAx(t, l, oNewVal);
					});
					this.cachedCatAx[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedCatAx[nId];
	};
	CChartStyleCache.prototype.checkValAx = function (nId) {
		if (!this.cachedValAx[nId]) {
			var sBinary = AscCommon.g_oValBinaries[nId];
			if (sBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sBinary);
					var oNewVal = new AscFormat.CValAx();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_ValAx(t, l, oNewVal);
					});
					this.cachedValAx[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedValAx[nId];
	};
	CChartStyleCache.prototype.checkView3d = function (nId) {
		if (!this.cachedView3d[nId]) {
			var sBinary = AscCommon.g_oView3dBinaries[nId];
			if (sBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sBinary);
					var oNewVal = new AscFormat.CView3d();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_View3D(t, l, oNewVal);
					});
					this.cachedView3d[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedView3d[nId];
	};
	CChartStyleCache.prototype.checkLegend = function (nId) {
		if (!this.cachedLegend[nId]) {
			var sBinary = AscCommon.g_oLegendBinaries[nId];
			if (sBinary) {
				AscFormat.ExecuteNoHistory(function () {
					var oReader = this.createBinaryReader(sBinary);
					var oNewVal = new AscFormat.CLegend();
					oReader.bcr.Read1(oReader.stream.size, function (t, l) {
						return oReader.ReadCT_Legend(t, l, oNewVal);
					});
					this.cachedLegend[nId] = oNewVal;
				}, this, []);
			}
		}
		return this.cachedLegend[nId];
	};
	CChartStyleCache.prototype.getStyleObject = function (aId) {
		if (!Array.isArray(aId)) {
			return null;
		}
		var nStyle = aId[0];
		var nColor = aId[1];
		var nDataLables = aId[2];
		var nCatAx = aId[3];
		var nValAx = aId[4];
		var nView3D = aId[5];
		var nLegend = aId[6];
		var oChartStyle = this.checkStyle(nStyle);
		if (!oChartStyle) {
			return null;
		}
		var oChartColors = this.checkColors(nColor);
		if (!oChartColors) {
			return null;
		}
		//val.oDataLablesData && val.oDataLablesData.crc32 || null,
		//val.oCatAxData && val.oCatAxData.crc32 || null,
		//val.oValAxData && val.oValAxData.crc32 || null,
		//val.oView3DData && val.oView3DData.crc32 || null,
		//val.oLegendData && val.oLegendData.crc32 || null
		return [
			oChartStyle,
			oChartColors,
			this.checkDataLabels(nDataLables),
			this.checkCatAx(nCatAx),
			this.checkValAx(nValAx),
			this.checkView3d(nView3D),
			this.checkLegend(nLegend)
		];
	};

	CChartStyleCache.prototype.getAdditionalData = function (nType, nStyleId) {
		var nStyleCrc = AscCommon.g_oChartStylesIdMap[nStyleId];
		if (!AscFormat.isRealNumber(nStyleCrc)) {
			return null;
		}
		var aStyles = AscCommon.g_oChartStyles[nType];
		if (!Array.isArray(aStyles)) {
			return null;
		}
		for (var nStyle = 0; nStyle < aStyles.length; ++nStyle) {
			var aStyle = aStyles[nStyle];
			if (aStyle[0] === nStyleCrc) {
				var nDataLables = aStyle[2];
				var nCatAx = aStyle[3];
				var nValAx = aStyle[4];
				var nView3D = aStyle[5];
				var nLegend = aStyle[6];
				var nBarData = aStyle[7];
				var oAdditionalData = new CAdditionalStyleData();
				oAdditionalData.dLbls = this.checkDataLabels(nDataLables);
				oAdditionalData.catAx = this.checkCatAx(nCatAx);
				oAdditionalData.valAx = this.checkValAx(nValAx);
				oAdditionalData.view3D = this.checkView3d(nView3D);
				oAdditionalData.legend = this.checkLegend(nLegend);
				var aBarData;
				if (Array.isArray(AscCommon.g_oBarParams[nBarData])) {
					aBarData = AscCommon.g_oBarParams[nBarData];
					oAdditionalData.gapWidth = aBarData[0];
					oAdditionalData.overlap = aBarData[1];
					oAdditionalData.gapDepth = aBarData[2];
				}
				return oAdditionalData;
			}
		}
		return null;
	};
	CChartStyleCache.prototype.getStyleIdx = function (nType, nStyleId) {
		var nStyleCrc = AscCommon.g_oChartStylesIdMap[nStyleId];
		if (!AscFormat.isRealNumber(nStyleCrc)) {
			return null;
		}
		var aStyles = AscCommon.g_oChartStyles[nType];
		if (!Array.isArray(aStyles)) {
			return null;
		}
		for (var nStyle = 0; nStyle < aStyles.length; ++nStyle) {
			var aStyle = aStyles[nStyle];
			if (aStyle[0] === nStyleCrc) {
				return nStyle + 1;
			}
		}
		return null;
	};
	var oChartStyleCache = new CChartStyleCache();
	//--------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].BAR_SHAPE_CONE = BAR_SHAPE_CONE;
	window['AscFormat'].BAR_SHAPE_CONETOMAX = BAR_SHAPE_CONETOMAX;
	window['AscFormat'].BAR_SHAPE_BOX = BAR_SHAPE_BOX;
	window['AscFormat'].BAR_SHAPE_CYLINDER = BAR_SHAPE_CYLINDER;
	window['AscFormat'].BAR_SHAPE_PYRAMID = BAR_SHAPE_PYRAMID;
	window['AscFormat'].BAR_SHAPE_PYRAMIDTOMAX = BAR_SHAPE_PYRAMIDTOMAX;
	window['AscFormat'].DISP_BLANKS_AS_GAP = DISP_BLANKS_AS_GAP;
	window['AscFormat'].DISP_BLANKS_AS_SPAN = DISP_BLANKS_AS_SPAN;
	window['AscFormat'].DISP_BLANKS_AS_ZERO = DISP_BLANKS_AS_ZERO;
	window['AscFormat'].checkBlackUnifill = checkBlackUnifill;
	window['AscFormat'].CreateUnifillSolidFillSchemeColorByIndex = CreateUnifillSolidFillSchemeColorByIndex;
	window['AscFormat'].CreateUniFillSchemeColorWidthTint = CreateUniFillSchemeColorWidthTint;
	window['AscFormat'].G_O_VISITED_HLINK_COLOR = G_O_VISITED_HLINK_COLOR;
	window['AscFormat'].G_O_HLINK_COLOR = G_O_HLINK_COLOR;
	window['AscFormat'].G_O_NO_ACTIVE_COMMENT_BRUSH = G_O_NO_ACTIVE_COMMENT_BRUSH;
	window['AscFormat'].G_O_ACTIVE_COMMENT_BRUSH = G_O_ACTIVE_COMMENT_BRUSH;
	window['AscFormat'].CChartSpace = CChartSpace;
	window['AscFormat'].CreateUnfilFromRGB = CreateUnfilFromRGB;
	window['AscFormat'].CreateUniFillSolidFillWidthTintOrShade = CreateUniFillSolidFillWidthTintOrShade;
	window['AscFormat'].CreateUnifillSolidFillSchemeColor = CreateUnifillSolidFillSchemeColor;
	window['AscFormat'].CreateNoFillLine = CreateNoFillLine;
	window['AscFormat'].CreateNoFillUniFill = CreateNoFillUniFill;
	window['AscFormat'].CreateView3d = CreateView3d;
	window['AscFormat'].CreateLineChart = CreateLineChart;
	window['AscFormat'].CreateBarChart = CreateBarChart;
	window['AscFormat'].CreateHBarChart = CreateHBarChart;
	window['AscFormat'].CreateAreaChart = CreateAreaChart;
	window['AscFormat'].CreatePieChart = CreatePieChart;
	window['AscFormat'].CreateScatterChart = CreateScatterChart;
	window['AscFormat'].CreateStockChart = CreateStockChart;
	window['AscFormat'].CreateDefaultAxes = CreateDefaultAxes;
	window['AscFormat'].CreateScatterAxis = CreateScatterAxis;
	window['AscFormat'].getChartSeries = getChartSeries;
	window['AscFormat'].checkSpPrRasterImages = checkSpPrRasterImages;
	window['AscFormat'].checkBlipFillRasterImages = checkBlipFillRasterImages;

	window['AscFormat'].PAGE_SETUP_ORIENTATION_DEFAULT = 0;
	window['AscFormat'].PAGE_SETUP_ORIENTATION_LANDSCAPE = 1;
	window['AscFormat'].PAGE_SETUP_ORIENTATION_PORTRAIT = 2;

	window['AscFormat'].MAX_SERIES_COUNT = MAX_SERIES_COUNT;
	window['AscFormat'].MAX_POINTS_COUNT = MAX_POINTS_COUNT;
	window['AscFormat'].MIN_STOCK_COUNT = MIN_STOCK_COUNT;

	window['AscFormat'].initStyleManager = initStyleManager;
	window['AscFormat'].CHART_STYLE_MANAGER = CHART_STYLE_MANAGER;
	window['AscFormat'].g_oChartStyleCache = oChartStyleCache;
	window['AscFormat'].CheckParagraphTextPr = CheckParagraphTextPr;
	window['AscFormat'].CheckObjectTextPr = CheckObjectTextPr;
	window['AscFormat'].CreateColorMapByIndex = CreateColorMapByIndex;
	window['AscFormat'].getArrayFillsFromBase = getArrayFillsFromBase;
	window['AscFormat'].getMaxIdx = getMaxIdx;
	window['AscFormat'].CreateSurfaceChart = CreateSurfaceChart;
	window['AscFormat'].CreateRadarChart = CreateRadarChart;
	window['AscFormat'].CPathMemory = CPathMemory;
	window['AscFormat'].CreateTypedBarChart = CreateTypedBarChart;
	window['AscFormat'].CreateTypedLineChart = CreateTypedLineChart;
	window['AscFormat'].CreateSurfaceAxes = CreateSurfaceAxes;
	window['AscFormat'].GetTypeMarkerByIndex = GetTypeMarkerByIndex;
})(window);
