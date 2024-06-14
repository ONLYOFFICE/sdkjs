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

(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
// Import

	var CreateNoFillLine = AscFormat.CreateNoFillLine;
	var CreateNoFillUniFill = AscFormat.CreateNoFillUniFill;
	
var c_oAscChartTypeSettings = Asc.c_oAscChartTypeSettings;
var c_oAscTickMark = Asc.c_oAscTickMark;
var c_oAscTickLabelsPos = Asc.c_oAscTickLabelsPos;
var fChartSize = 75;

function ChartPreviewManager() {
	AscCommon.CActionOnTimerBase.call(this);
	this.previewGroups = [];
	this.chartsByTypes = [];

	this.CHART_PREVIEW_WIDTH_PIX = 50;
	this.CHART_PREVIEW_HEIGHT_PIX = 50;

	this._canvas_charts = null;


	this.FirstActionOnTimer = true;
	this.Index = -1;
	this.ChartType = -1;
	this.StylesIndexes = [];
	this.Buffer = [];
}
ChartPreviewManager.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
ChartPreviewManager.prototype.constructor = ChartPreviewManager;

	ChartPreviewManager.prototype.GetApi = function() {
		return Asc.editor || editor;
	};
	ChartPreviewManager.prototype.OnBegin = function(nChartType, arrId) {
		var aStyles = AscCommon.g_oChartStyles[nChartType];
		if(!Array.isArray(aStyles)) {
			this.GetApi().sendEvent("asc_onBeginChartStylesPreview", 0);
			return;
		}
		this.ChartType = nChartType;
		this.Index = 0;
		var nStyle;

		this.StylesIndexes.length = 0;
		if(!Array.isArray(arrId)) {
			for(nStyle = 0; nStyle < aStyles.length; ++nStyle) {
				this.StylesIndexes.push(nStyle);
			}
		}
		else {
			var nId;
			for(nId = 0; nId < arrId.length; ++nId) {
				this.StylesIndexes.push(arrId[nId] - 1);
			}
			for(nStyle = 0; nStyle < aStyles.length; ++nStyle) {
				for(nId = 0; nId < arrId.length; ++nId) {
					if(nStyle === (arrId[nId] - 1)) {
						break;
					}
				}
				if(nId === arrId.length) {
					this.StylesIndexes.push(nStyle);
				}
			}
		}
		this.GetApi().sendEvent("asc_onBeginChartStylesPreview", aStyles.length);
	};
	ChartPreviewManager.prototype.OnEnd = function() {
		this.GetApi().sendEvent("asc_onEndChartStylesPreview");
	};
	ChartPreviewManager.prototype.IsContinue = function() {
		return (this.Index < this.StylesIndexes.length);
	};
	ChartPreviewManager.prototype.DoAction = function() {
		let aStyles = AscCommon.g_oChartStyles[this.ChartType];
		if(aStyles) {
			var oStyle = aStyles[this.StylesIndexes[this.Index]]
			if(oStyle) {
				let graphics = this._getGraphics();
				this.createChartPreview(graphics, this.ChartType, oStyle);
				let oPreview = new AscCommon.CStyleImage();
				oPreview.name = this.StylesIndexes[this.Index] + 1;
				oPreview.image = this._canvas_charts.toDataURL("image/png");
				this.Buffer.push(oPreview);
			}
		}
		this.Index++;
	};
	ChartPreviewManager.prototype.OnEndTimer = function() {
		this.GetApi().sendEvent("asc_onAddChartStylesPreview", this.Buffer);
		this.Buffer = [];
	};

ChartPreviewManager.prototype.getAscChartSeriesDefault = function(type) {
	function createItem(value) {
		return { numFormatStr: "General", isDateTimeFormat: false, val: value, isHidden: false };
	}

	// Set data
	var series = [], ser;
	switch(type)
	{
		case c_oAscChartTypeSettings.lineNormal:
		case c_oAscChartTypeSettings.lineNormalMarker:
		case c_oAscChartTypeSettings.scatterLine:
		case c_oAscChartTypeSettings.scatterLineMarker:
		case c_oAscChartTypeSettings.scatterSmooth:
		case c_oAscChartTypeSettings.scatterSmoothMarker:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2), createItem(3), createItem(2), createItem(3) ];
			series.push(ser);
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(2), createItem(3), createItem(2) ];
			series.push(ser);
			break;
		}
        case c_oAscChartTypeSettings.line3d:
        {
            ser = new AscFormat.asc_CChartSeria();
            ser.Val.NumCache = [ createItem(1), createItem(2), createItem(1), createItem(2) ];
            series.push(ser);
            ser = new AscFormat.asc_CChartSeria();
            ser.Val.NumCache = [ createItem(3), createItem(2.5), createItem(3), createItem(3.5) ];
            series.push(ser);
            break;
        }
		case c_oAscChartTypeSettings.lineStacked:
		case c_oAscChartTypeSettings.lineStackedMarker:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(6), createItem(2), createItem(8) ];
			series.push(ser);
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4), createItem(4), createItem(4), createItem(5) ];
			series.push(ser);
			break;
		}
		case c_oAscChartTypeSettings.lineStackedPer:
		case c_oAscChartTypeSettings.lineStackedPerMarker:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2), createItem(4), createItem(2), createItem(4) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2), createItem(2), createItem(2), createItem(2) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.hBarNormal:
		case c_oAscChartTypeSettings.hBarNormal3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(3) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.hBarStacked:
		case c_oAscChartTypeSettings.hBarStacked3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4), createItem(3), createItem(2), createItem(1) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(5), createItem(4), createItem(3), createItem(2) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.hBarStackedPer:
		case c_oAscChartTypeSettings.hBarStackedPer3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(7), createItem(5), createItem(3), createItem(1) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(7), createItem(6), createItem(5), createItem(4) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.barNormal:
		case c_oAscChartTypeSettings.barNormal3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(3) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.barStacked:
		case c_oAscChartTypeSettings.barStacked3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(2), createItem(3), createItem(4) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2), createItem(3), createItem(4), createItem(5) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.barStackedPer:
		case c_oAscChartTypeSettings.barStackedPer3d:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(3), createItem(5), createItem(7) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4), createItem(5), createItem(6), createItem(7) ];
			series.push(ser);
			break;
		}
        case c_oAscChartTypeSettings.barNormal3dPerspective:
        {
            ser = new AscFormat.asc_CChartSeria();
            ser.Val.NumCache = [ createItem(1), createItem(2), createItem(3), createItem(4) ];
            series.push(ser);

            ser = new AscFormat.asc_CChartSeria();
            ser.Val.NumCache = [ createItem(2), createItem(3), createItem(4), createItem(5) ];
            series.push(ser);
            break;
        }
		case c_oAscChartTypeSettings.pie:
		case c_oAscChartTypeSettings.doughnut:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(3), createItem(1) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.areaNormal:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(0), createItem(8), createItem(5), createItem(6) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(0), createItem(4), createItem(2), createItem(9) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.areaStacked:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(0), createItem(8), createItem(5), createItem(11) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4), createItem(4), createItem(4), createItem(4) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.areaStackedPer:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(0), createItem(4), createItem(1), createItem(16) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(4), createItem(4), createItem(4), createItem(4) ];
			series.push(ser);
			break;
		}

		case c_oAscChartTypeSettings.scatter:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(5) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(2), createItem(6) ];
			series.push(ser);
			break;
		}

		default:
		{
			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(3), createItem(5), createItem(7) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(10), createItem(12), createItem(14) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(1), createItem(3), createItem(5) ];
			series.push(ser);

			ser = new AscFormat.asc_CChartSeria();
			ser.Val.NumCache = [ createItem(8), createItem(10), createItem(12) ];
			series.push(ser);
			break;
		}
	}
	return series;
};
ChartPreviewManager.prototype.getChartByType = function(type)
{
	return AscFormat.ExecuteNoHistory(function()
	{
		var settings = new Asc.asc_ChartSettings();
		settings.type = type;
		var chartSeries = this.getAscChartSeriesDefault(type);
		var chart_space = AscFormat.DrawingObjectsController.prototype._getChartSpace(chartSeries, settings, true);
        chart_space.bPreview = true;
		if (Asc['editor'] && AscCommon.c_oEditorId.Spreadsheet === Asc['editor'].getEditorId()) {
			var api_sheet = Asc['editor'];
			chart_space.setWorksheet(api_sheet.wb.getWorksheet().model);
		} else {
			if (editor && editor.WordControl && editor.WordControl.m_oLogicDocument.Slides &&
				editor.WordControl.m_oLogicDocument.Slides[editor.WordControl.m_oLogicDocument.CurPage]) {
				chart_space.setParent(editor.WordControl.m_oLogicDocument.Slides[editor.WordControl.m_oLogicDocument.CurPage]);
			}
		}
		AscFormat.CheckSpPrXfrm(chart_space);
		chart_space.spPr.xfrm.setOffX(0);
		chart_space.spPr.xfrm.setOffY(0);
		chart_space.spPr.xfrm.setExtX(fChartSize);
		chart_space.spPr.xfrm.setExtY(fChartSize);
		settings.putTitle(Asc.c_oAscChartTitleShowSettings.noOverlay);


		var val_ax_props = new AscCommon.asc_ValAxisSettings();
		val_ax_props.putMinValRule(Asc.c_oAscValAxisRule.auto);
		val_ax_props.putMaxValRule(Asc.c_oAscValAxisRule.auto);
		val_ax_props.putTickLabelsPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		val_ax_props.putInvertValOrder(false);
		val_ax_props.putDispUnitsRule(Asc.c_oAscValAxUnits.none);
		val_ax_props.putMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax_props.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		val_ax_props.putCrossesRule(Asc.c_oAscCrossesRule.auto);


		var cat_ax_props = new AscCommon.asc_CatAxisSettings();
		cat_ax_props.putIntervalBetweenLabelsRule(Asc.c_oAscBetweenLabelsRule.auto);
		cat_ax_props.putLabelsPosition(Asc.c_oAscLabelsPosition.betweenDivisions);
		cat_ax_props.putTickLabelsPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		cat_ax_props.putLabelsAxisDistance(100);
		cat_ax_props.putMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax_props.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		cat_ax_props.putIntervalBetweenTick(1);
		cat_ax_props.putCrossesRule(Asc.c_oAscCrossesRule.auto);
		var vert_axis_settings, hor_axis_settings, isScatter;
		switch(type)
		{
			case c_oAscChartTypeSettings.hBarNormal:
			case c_oAscChartTypeSettings.hBarStacked:
			case c_oAscChartTypeSettings.hBarStackedPer:
			{
				vert_axis_settings = cat_ax_props;
				hor_axis_settings = val_ax_props;
				break;
			}
			case c_oAscChartTypeSettings.scatter:
			case c_oAscChartTypeSettings.scatterLine:
			case c_oAscChartTypeSettings.scatterLineMarker:
			case c_oAscChartTypeSettings.scatterMarker:
			case c_oAscChartTypeSettings.scatterNone:
			case c_oAscChartTypeSettings.scatterSmooth:
			case c_oAscChartTypeSettings.scatterSmoothMarker:
			{
				vert_axis_settings = val_ax_props;
				hor_axis_settings = val_ax_props;
				isScatter = true;
				break;
			}
            case c_oAscChartTypeSettings.areaNormal:
            case c_oAscChartTypeSettings.areaStacked:
            case c_oAscChartTypeSettings.areaStackedPer:
            {
                cat_ax_props.putLabelsPosition(AscFormat.CROSS_BETWEEN_BETWEEN);
                vert_axis_settings = val_ax_props;
                hor_axis_settings = cat_ax_props;

                break;
            }

			default :
			{
				vert_axis_settings = val_ax_props;
				hor_axis_settings = cat_ax_props;
				break;
			}
		}

		settings.addVertAxesProps(vert_axis_settings);
		settings.addHorAxesProps(hor_axis_settings);

		AscFormat.DrawingObjectsController.prototype.applyPropsToChartSpace(settings, chart_space);
		chart_space.setBDeleted(false);
		chart_space.updateLinks();
		if(!(isScatter || type === c_oAscChartTypeSettings.stock))
		{
			if(chart_space.chart.plotArea.valAx)
				chart_space.chart.plotArea.valAx.setDelete(true);
			if(chart_space.chart.plotArea.catAx)
				chart_space.chart.plotArea.catAx.setDelete(true);
		}
		else
		{
			if(chart_space.chart.plotArea.valAx)
			{
				chart_space.chart.plotArea.valAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE);
				chart_space.chart.plotArea.valAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);
				chart_space.chart.plotArea.valAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
			}
			if(chart_space.chart.plotArea.catAx)
			{
				chart_space.chart.plotArea.catAx.setTickLblPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE);
				chart_space.chart.plotArea.catAx.setMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);
				chart_space.chart.plotArea.catAx.setMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
			}
		}
		if(!chart_space.spPr)
			chart_space.setSpPr(new AscFormat.CSpPr());

		var new_line = new AscFormat.CLn();
		new_line.setFill(new AscFormat.CUniFill());
		new_line.Fill.setFill(new AscFormat.CNoFill());
		chart_space.spPr.setLn(new_line);

        chart_space.recalcInfo.recalculateReferences = false;
		chart_space.recalculate();

		return chart_space;
	}, this, []);
};

ChartPreviewManager.prototype.clearPreviews = function()
{
	this.previewGroups.length = 0;
};

ChartPreviewManager.prototype.checkChartForPreview = function (type, aStyle) {
	return AscFormat.ExecuteNoHistory(function() {
		if (!this.chartsByTypes[type]) {
			this.chartsByTypes[type] = this.getChartByType(type);
		}
		var chart_space = this.chartsByTypes[type];
		chart_space.applyChartStyleByIds(aStyle);
		chart_space.recalcInfo.recalculateReferences = false;
		chart_space.recalculate();
		return chart_space;
	}, this, []);
};
ChartPreviewManager.prototype.createChartPreview = function(graphics, type, aStyle) {
	var chart_space = this.checkChartForPreview(type, aStyle);
	graphics.save();
	chart_space.draw(graphics);
	graphics.restore();
};

ChartPreviewManager.prototype._isCachedChartStyles = function(chartType) {
	var res = this.previewGroups.hasOwnProperty(chartType);
	if(!res) {
		this.previewGroups[chartType] = [];
	}
	return res;
};
ChartPreviewManager.prototype._getGraphics = function() {
	if (null === this._canvas_charts) {
		this._canvas_charts = document.createElement('canvas');
		this._canvas_charts.width = AscCommon.AscBrowser.convertToRetinaValue(this.CHART_PREVIEW_WIDTH_PIX, true);
		this._canvas_charts.height = AscCommon.AscBrowser.convertToRetinaValue(this.CHART_PREVIEW_HEIGHT_PIX, true);
	}

	var _canvas = this._canvas_charts;
	var ctx = _canvas.getContext('2d');
	var graphics = new AscCommon.CGraphics();
	graphics.init(ctx, _canvas.width, _canvas.height, fChartSize, fChartSize);
	graphics.m_oFontManager = AscCommon.g_fontManager;
	graphics.transform(1,0,0,1,0,0);
	return graphics;
};

ChartPreviewManager.prototype.getChartPreviews = function(chartType, arrId, bEmpty) {

	var aStyles;
	var nIdx, chartStyle, graphics;
	var arrPreviews;
	if(bEmpty) {
		arrPreviews = [];
		aStyles = AscCommon.g_oChartStyles[chartType];
		if(Array.isArray(aStyles)) {
			graphics = this._getGraphics();
			for (nIdx = 0; nIdx < aStyles.length; ++nIdx) {
				chartStyle = new AscCommon.CStyleImage();
				chartStyle.name = nIdx + 1;
				chartStyle.image = null;
				arrPreviews.push(chartStyle);
			}
		}
		return arrPreviews;
	}
	if(Array.isArray(arrId)) {
		arrPreviews = [];
		aStyles = AscCommon.g_oChartStyles[chartType];
		if(Array.isArray(aStyles)) {
			graphics = this._getGraphics();
			for (nIdx = 0; nIdx < arrId.length; ++nIdx) {
				var oStyle = aStyles[arrId[nIdx] - 1];
				if(oStyle) {
					this.createChartPreview(graphics, chartType, oStyle);

					chartStyle = new AscCommon.CStyleImage();
					chartStyle.name = arrId[nIdx];
					chartStyle.image = this._canvas_charts.toDataURL("image/png");
					arrPreviews.push(chartStyle);
				}
			}
		}
		return arrPreviews;
	}
	if (AscFormat.isRealNumber(chartType)) {


		if (!this._isCachedChartStyles(chartType)) {
			aStyles = AscCommon.g_oChartStyles[chartType];
			if (Array.isArray(aStyles)) {
				AscFormat.ExecuteNoHistory(function () {
					graphics = this._getGraphics();
					if(Array.isArray(arrId)) {
						var arrPreviews = [];
						for (nIdx = 0; nIdx < arrId.length; ++nIdx) {
							if(aStyles[arrId[nIdx]]) {
								this.createChartPreview(graphics, chartType, aStyles[arrId[nIdx]]);

								chartStyle = new AscCommon.CStyleImage();
								chartStyle.name = nIdx + 1;
								chartStyle.image = this._canvas_charts.toDataURL("image/png");
								arrPreviews.push(chartStyle);
							}
						}
						return arrPreviews;
					}
					else {
						for (nIdx = 0; nIdx < aStyles.length; ++nIdx) {
							this.createChartPreview(graphics, chartType, aStyles[nIdx]);
							if (!window["IS_NATIVE_EDITOR"]) {
								chartStyle = new AscCommon.CStyleImage();
								chartStyle.name = nIdx + 1;
								chartStyle.image = this._canvas_charts.toDataURL("image/png");
								this.previewGroups[chartType].push(chartStyle);
							}
						}
					}
				}, this, []);

				var api = Asc['editor'];
				if (api && AscCommon.c_oEditorId.Spreadsheet === api.getEditorId()) {
					var _graphics = api.wb.shapeCtx;
					if (_graphics.ClearLastFont) {
						_graphics.ClearLastFont();
					}
				}
			}

		}
	}
	return this.previewGroups[chartType];
};

	function CSmartArtPreviewInfo(nSmartArtId, nSectionId, oImage) {
		this.m_nSectionId = AscFormat.isRealNumber(nSectionId) ? nSectionId : null;
		this.m_nSmartArtId = AscFormat.isRealNumber(nSmartArtId) ? nSmartArtId : null;
		this.m_oImage = oImage ? oImage : null;
	}

	CSmartArtPreviewInfo.prototype.asc_getImage = function () {
		return this.m_oImage;
	};
	CSmartArtPreviewInfo.prototype["asc_getImage"] = CSmartArtPreviewInfo.prototype.asc_getImage;

	CSmartArtPreviewInfo.prototype.asc_getSectionId = function () {
		return this.m_nSectionId;
	};
	CSmartArtPreviewInfo.prototype["asc_getSectionId"] = CSmartArtPreviewInfo.prototype.asc_getSectionId;

	CSmartArtPreviewInfo.prototype.asc_getSmartArtId = function () {
		return this.m_nSmartArtId;
	};
	CSmartArtPreviewInfo.prototype["asc_getSmartArtId"] = CSmartArtPreviewInfo.prototype.asc_getSmartArtId;

	CSmartArtPreviewInfo.prototype.asc_setImage = function (oImage) {
		this.m_oImage = oImage;
	};
	CSmartArtPreviewInfo.prototype["asc_setImage"] = CSmartArtPreviewInfo.prototype.asc_setImage;

	CSmartArtPreviewInfo.prototype.asc_setSectionId = function (nSectionId) {
		this.m_nSectionId = nSectionId;
	};
	CSmartArtPreviewInfo.prototype["asc_setSectionId"] = CSmartArtPreviewInfo.prototype.asc_setSectionId;

	CSmartArtPreviewInfo.prototype.asc_setSmartArtId = function (nSmartArtId) {
		this.m_nSmartArtId = nSmartArtId;
	};
	CSmartArtPreviewInfo.prototype["asc_setSmartArtId"] = CSmartArtPreviewInfo.prototype.asc_setSmartArtId;


	function SmartArtPreviewDrawer() {
		AscCommon.CActionOnTimerBase.call(this);
		this.SMARTART_PREVIEW_SIZE_MM = 8128000 * AscCommonWord.g_dKoef_emu_to_mm;
		this.CANVAS_SIZE = 70;
		this.canvas = null;
		this.imageType = "image/png";
		this.imageBuffer = [];
		this.index = 0;
		this.cache = {};
		this.queue = [];
		this.typeOfSectionLoad = null;
	}
	SmartArtPreviewDrawer.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
	SmartArtPreviewDrawer.prototype.constructor = SmartArtPreviewDrawer;

	SmartArtPreviewDrawer.prototype.Begin = function (nTypeOfSectionLoad) {

		const oApi = Asc.editor || editor;
		if(!oApi) {
			return;
		}
		this.typeOfSectionLoad = nTypeOfSectionLoad;
		const oThis = this;
		AscCommon.g_oBinarySmartArts.checkLoadDrawing().then(function ()
		{
			if (AscFormat.isRealNumber(oThis.typeOfSectionLoad)) {
				const arrPreviewObjects = Asc.c_oAscSmartArtSections[oThis.typeOfSectionLoad].map(function (nTypeOfSmartArt) {
					return new CSmartArtPreviewInfo(nTypeOfSmartArt, oThis.typeOfSectionLoad);
				});
				oThis.queue = oThis.queue.concat(arrPreviewObjects);
				AscCommon.CActionOnTimerBase.prototype.Begin.call(oThis);
			}
		});
	};
	SmartArtPreviewDrawer.prototype.OnBegin = function () {
		const oApi = Asc.editor || editor;
		this.index = 0;
		if (oApi)
			oApi.sendEvent("asc_onBeginSmartArtPreview", this.typeOfSectionLoad);
	}

	SmartArtPreviewDrawer.prototype.OnEnd = function() {
		const oApi = Asc.editor || editor;
		if (oApi) oApi.sendEvent("asc_onEndSmartArtPreview");

	};

	SmartArtPreviewDrawer.prototype.IsContinue = function() {
		return !!this.queue.length;
	};

	SmartArtPreviewDrawer.prototype.DoAction = function() {
		const oApi = Asc.editor || editor;
		oApi.isSkipAddIdToBaseObject = true;
		const oSmartArtPreviewInfo = this.queue.pop();
		if (oSmartArtPreviewInfo) {
			const nTypeOfSmartArt = oSmartArtPreviewInfo.asc_getSmartArtId();
			if (!this.cache[nTypeOfSmartArt]) {
				const oContext = this.createSmartArtPreview(nTypeOfSmartArt);
				const oPreview = new AscCommon.CStyleImage();
				oPreview.name = nTypeOfSmartArt;
				oPreview.image = oContext.canvas.toDataURL(this.imageType, 1);
				this.cache[nTypeOfSmartArt] = oPreview;
			}
			oSmartArtPreviewInfo.asc_setImage(this.cache[nTypeOfSmartArt]);
			this.imageBuffer.push(oSmartArtPreviewInfo);
		}
		delete oApi.isSkipAddIdToBaseObject;
	};

	SmartArtPreviewDrawer.prototype.OnEndTimer = function() {
		const oApi = Asc.editor || editor;
		oApi.sendEvent("asc_onAddSmartArtPreview", this.imageBuffer);
		this.imageBuffer = [];
	};

	SmartArtPreviewDrawer.prototype.getGraphics = function () {
		if (null === this.canvas) {
			this.canvas = document.createElement('canvas');
			this.canvas.width = AscCommon.AscBrowser.convertToRetinaValue(this.CANVAS_SIZE, true);
			this.canvas.height = AscCommon.AscBrowser.convertToRetinaValue(this.CANVAS_SIZE, true);
		}

		const oCanvas = this.canvas;
		const oContext = oCanvas.getContext('2d');
		oContext.fillStyle = 'white';
		oContext.fillRect(0, 0, oCanvas.width, oCanvas.height);
		const oGraphics = new AscCommon.CGraphics();
		oGraphics.isSmartArtPreviewDrawer = true;
		oGraphics.init(oContext, oCanvas.width, oCanvas.height, this.SMARTART_PREVIEW_SIZE_MM, this.SMARTART_PREVIEW_SIZE_MM);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.transform(1,0,0,1,0,0);
		if (AscCommon.AscBrowser.retinaPixelRatio < 2) {
			oGraphics.bDrawRectWithLines = true;
		}
		return oGraphics;
	}

	// SmartArtPreviewDrawer.prototype.createPreviews = function () {
	// const arrPreviewObjects = Asc.c_oAscSmartArtSections[Asc.c_oAscSmartArtSectionNames.OfficeCom].map(function (nTypeOfSmartArt) {
	// 	return new CSmartArtPreviewInfo(nTypeOfSmartArt, Asc.c_oAscSmartArtSectionNames.OfficeCom);
	// });
	// this.queue = this.queue.concat(arrPreviewObjects);
	// 	while (this.IsContinue()) {
	// 		this.DoAction();
	// 	}
	// 	console.log(this.imageBuffer);
	// };

	SmartArtPreviewDrawer.prototype.start = function () {
		this.loadImagePlaceholder(this.createPreviews.bind(this));
	};

	SmartArtPreviewDrawer.prototype.createSmartArtPreview = function (nType) {
		const oSmartArt = this.getSmartArt(nType);
		const oGraphics = this.getGraphics();
		oGraphics.save();
		oSmartArt.draw(oGraphics);
		oGraphics.restore();
		return oGraphics.m_oContext;
	};

	SmartArtPreviewDrawer.prototype.fitSmartArtForPreview = function (oSmartArt) {
		AscFormat.ExecuteNoHistory(function () {
			if (oSmartArt.spTree[0] && oSmartArt.spTree[0].spTree) {
				const nWidth = this.SMARTART_PREVIEW_SIZE_MM;
				const nHeight = this.SMARTART_PREVIEW_SIZE_MM;

				oSmartArt.spTree[0].recalcBounds();
				oSmartArt.spTree[0].bForceGroupBounds = true;
				oSmartArt.spTree[0].spTree.forEach(function (shape) {
					shape.recalcBounds();
				});
				oSmartArt.recalculateBounds();
				let oBounds = oSmartArt.spTree[0].bounds;
				let nSmartArtWidth = oBounds.w;
				let nSmartArtHeight = oBounds.h;

				const nCoefficientHeight = Math.abs(nWidth / nSmartArtWidth);
				const nCoefficientWith = Math.abs(nHeight / nSmartArtHeight);
				const nMinCoefficient = Math.min(nCoefficientHeight, nCoefficientWith);
				if (nMinCoefficient > 1) {
					oSmartArt.changeSize(nMinCoefficient - 0.1, nMinCoefficient - 0.1);
					oSmartArt.spTree[0].recalcBounds();
					oSmartArt.spTree[0].spTree.forEach(function (shape) {
						shape.recalcBounds();
					});
					oSmartArt.recalculateBounds();
					oBounds = oSmartArt.spTree[0].bounds;
					nSmartArtWidth = oBounds.w;
					nSmartArtHeight = oBounds.h;
					const nX = oBounds.x;
					const nY = oBounds.y;
					const nCX = nX + nSmartArtWidth / 2;
					const nCY = nY + nSmartArtHeight / 2;
					const nDX = nWidth / 2 - nCX;
					const nDY = nHeight / 2 - nCY;
					const oXfrm = oSmartArt.spPr.xfrm;
					oXfrm.setOffX(oXfrm.offX + nDX);
					oXfrm.setOffY(oXfrm.offY + nDY);
				}
				delete oSmartArt.spTree[0].bForceGroupBounds;
			}

		}, this, []);
	}

	SmartArtPreviewDrawer.prototype.getSmartArt = function(nSmartArtType) {
		return AscFormat.ExecuteNoHistory(function () {
			const oSmartArt = new AscFormat.SmartArt();
			oSmartArt.bNeedUpdatePosition = false;
			oSmartArt.bFirstRecalculate = false;
			const oApi = Asc.editor || editor;
				oSmartArt.bForceSlideTransform = true;
				oSmartArt.fillByPreset(nSmartArtType, true);
				oSmartArt.getContrastDrawing();
				oSmartArt.setBDeleted2(false);
				const oXfrm = oSmartArt.spPr.xfrm;
				const oDrawingObjects = oApi.getDrawingObjects();
				oXfrm.setOffX(0);
				oXfrm.setOffY((this.SMARTART_PREVIEW_SIZE_MM - oXfrm.extY) / 2);
				if (oDrawingObjects) {
					oSmartArt.setDrawingObjects(oDrawingObjects);
					if (oDrawingObjects.cSld) {
						oSmartArt.setParent2(oDrawingObjects);
						oSmartArt.setRecalculateInfo();
					}

					if (oDrawingObjects.getWorksheetModel) {
						oSmartArt.setWorksheet(oDrawingObjects.getWorksheetModel());
					}
				}
				this.fitSmartArtForPreview(oSmartArt);
				oSmartArt.recalcTransformText();
				oSmartArt.recalculate();

			return oSmartArt;
		}, this, []);
	};

function CreateAscColorByIndex(nIndex)
{
	var oColor = new Asc.asc_CColor();
	oColor.type = Asc.c_oAscColor.COLOR_TYPE_SCHEME;
	oColor.value = nIndex;
	return oColor;
}

function CreateAscFillByIndex(nIndex)
{
	var oAscFill = new Asc.asc_CShapeFill();
	oAscFill.type = Asc.c_oAscFill.FILL_TYPE_SOLID;
	oAscFill.fill = new Asc.asc_CFillSolid();
	oAscFill.fill.color = CreateAscColorByIndex(nIndex);
	return oAscFill;
}

function CreateAscGradFillByIndex(nIndex1, nIndex2, nAngle)
{
	var oAscFill = new Asc.asc_CShapeFill();
	oAscFill.type = Asc.c_oAscFill.FILL_TYPE_GRAD;
	oAscFill.fill = new Asc.asc_CFillGrad();
	oAscFill.fill.GradType = Asc.c_oAscFillGradType.GRAD_LINEAR;
	oAscFill.fill.LinearAngle = nAngle;
	oAscFill.fill.LinearScale = true;
	oAscFill.fill.Colors = [CreateAscColorByIndex(nIndex1), CreateAscColorByIndex(nIndex2)];
	oAscFill.fill.Positions = [0, 100000];
	oAscFill.fill.LinearAngle = nAngle;
	oAscFill.fill.LinearScale = true;
	return oAscFill;
}
function TextArtPreviewManager()
{
	this.canvas = null;
	this.canvasWidth = 50;
	this.canvasHeight = 50;
	this.shapeWidth = 50;
	this.shapeHeight = 50;
	this.TAShape = null;
	this.TextArtStyles = [];

	this.aStylesByIndex = [];
	this.aStylesByIndexToApply = [];

	this.dKoeff = 4;
	//if (AscBrowser.isRetina) {
	//	this.dKoeff <<= 1;
	//}
}
TextArtPreviewManager.prototype.initStyles = function()
{

	var oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(24), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[0] = oTextPr;
	this.aStylesByIndexToApply[0] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscGradFillByIndex(52, 24, 5400000), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[4] = oTextPr;
	this.aStylesByIndexToApply[4] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscGradFillByIndex(44, 42, 5400000), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[8] = oTextPr;
	this.aStylesByIndexToApply[8] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(34), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000)*this.dKoeff);
	this.aStylesByIndex[1] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(34), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000));
	this.aStylesByIndexToApply[1] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(59), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000)*this.dKoeff);
	this.aStylesByIndex[5] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(59), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000));
	this.aStylesByIndexToApply[5] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(52), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000)*this.dKoeff);
	this.aStylesByIndex[9] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = CreateNoFillUniFill();
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(52), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (15773/36000));
	this.aStylesByIndexToApply[9] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(27), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(52), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000)*this.dKoeff);
	this.aStylesByIndex[2] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(27), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(52), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000));
	this.aStylesByIndexToApply[2] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(42), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(46), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000)*this.dKoeff);
	this.aStylesByIndex[6] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(42), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(46), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000));
	this.aStylesByIndexToApply[6] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(57), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(54), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000)*this.dKoeff);
	this.aStylesByIndex[10] = oTextPr;
	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscFillByIndex(57), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = AscFormat.CreatePenFromParams(AscFormat.CorrectUniFill(CreateAscFillByIndex(54), new AscFormat.CUniFill(), 0), undefined, undefined, undefined, undefined, (12700/36000));
	this.aStylesByIndexToApply[10] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscGradFillByIndex(45, 57, 0), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[3] = oTextPr;
	this.aStylesByIndexToApply[3] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscGradFillByIndex(52, 33, 0), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[7] = oTextPr;
	this.aStylesByIndexToApply[7] = oTextPr;

	oTextPr = new CTextPr();
	oTextPr.Bold = true;
	oTextPr.TextFill = AscFormat.CorrectUniFill(CreateAscGradFillByIndex(27, 45, 5400000), new AscFormat.CUniFill(), 0);
	oTextPr.TextOutline = CreateNoFillLine();
	this.aStylesByIndex[11] = oTextPr;
	this.aStylesByIndexToApply[11] = oTextPr;
};

TextArtPreviewManager.prototype.getStylesToApply = function()
{
	if(this.aStylesByIndex.length === 0)
	{
		this.initStyles();
	}
	return this.aStylesByIndexToApply;
};

TextArtPreviewManager.prototype.clear = function()
{
	this.TextArtStyles.length = 0;
};

TextArtPreviewManager.prototype.getWordArtStyles = function()
{
	if(this.TextArtStyles.length === 0)
	{
		this.generateTextArtStyles();
	}
	return this.TextArtStyles;
};

TextArtPreviewManager.prototype.getCanvas = function()
{
	if (null === this.canvas)
	{
		this.canvas = this.createCanvas();
	}
	return this.canvas;
};
TextArtPreviewManager.prototype.createCanvas = function()
{
	var oCanvas = document.createElement('canvas');
	oCanvas.width = AscCommon.AscBrowser.convertToRetinaValue(this.canvasWidth, true);
	oCanvas.height = AscCommon.AscBrowser.convertToRetinaValue(this.canvasHeight, true);
	return oCanvas;
};

TextArtPreviewManager.prototype.getShapeByPrst = function(prst)
{
	var oShape = this.getShape();
    if(!oShape)
    {
        return null;
    }
	var oContent = oShape.getDocContent();

	var TextSpacing = undefined;
	switch(prst)
	{
		case "textButton":
		{
			TextSpacing = 4;
			oContent.AddText("abcde");
			oContent.AddNewParagraph();
			oContent.AddText("Fghi");
			oContent.AddNewParagraph();
			oContent.AddText("Jklmn");
			break;
		}
		case "textArchUp":
		case "textArchDown":
		{
			TextSpacing = 4;
			oContent.AddText("abcdefg");
			break;
		}

		case "textCircle":
		{
			TextSpacing = 4;
			oContent.AddText("abcdefghijklmnop");
			break;
		}
        case "textButtonPour":
        {
			oContent.AddText("abcde");
            oContent.AddNewParagraph();
			oContent.AddText("abc");
            oContent.AddNewParagraph();
			oContent.AddText("abcde");
            break;
        }
        case "textDeflateInflate":
        {
			oContent.AddText("abcde");
            oContent.AddNewParagraph();
			oContent.AddText("abcde");
            break;
        }
        case "textDeflateInflateDeflate":
        {
			oContent.AddText("abcde");
            oContent.AddNewParagraph();
			oContent.AddText("abcde");
            oContent.AddNewParagraph();
			oContent.AddText("abcde");
            break;
        }
		default:
		{
			oContent.AddText("abcde");
			break;
		}
	}
	oContent.SetApplyToAll(true);
	oContent.SetParagraphAlign(AscCommon.align_Center);
	oContent.AddToParagraph(new ParaTextPr({FontSize: 36, Spacing: TextSpacing, Unifill: AscFormat.CreateUnfilFromRGB(0, 0, 0)}));
	oContent.SetApplyToAll(false);

	var oBodypr = oShape.getBodyPr().createDuplicate();
	oBodypr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(prst);
	oBodypr.setDefaultInsets();
	if(!oShape.bWordShape)
	{
		oShape.txBody.setBodyPr(oBodypr);
	}
	else
	{
		oShape.setBodyPr(oBodypr);
	}
	oShape.setBDeleted(false);
	oShape.recalcText();
	oShape.recalculate();
	if(oShape.bWordShape)
	{
		oShape.recalculateText();
	}
	return oShape;
};
TextArtPreviewManager.prototype.getShape =  function()
{
	var oShape = new AscFormat.CShape();
	var oParent = null, oWorkSheet = null;
	var bWord = true;
	let nEditorId = Asc['editor'].getEditorId();
	if (Asc.editor.isPdfEditor()) {
		bWord = false;
	}
	else if(nEditorId === AscCommon.c_oEditorId.Spreadsheet) {
		var api_sheet = Asc['editor'];
		oShape.setWorksheet(api_sheet.wb.getWorksheet().model);
		oWorkSheet = api_sheet.wb.getWorksheet().model;
		bWord = false;
	}
	else if(nEditorId === AscCommon.c_oEditorId.Presentation) {
		let oPres = Asc.editor.private_GetLogicDocument();
		let oSlide = oPres.GetCurrentSlide();
		oShape.setParent(oSlide);
		oParent = oSlide;
		bWord = false;
	}


	var oParentObjects = oShape.getParentObjects();
	var oTrack = new AscFormat.NewShapeTrack("textRect", 0, 0, oParentObjects.theme, oParentObjects.master, oParentObjects.layout, oParentObjects.slide, 0);
	oTrack.track({}, oShape.convertPixToMM(this.canvasWidth), oShape.convertPixToMM(this.canvasHeight));
	oShape = oTrack.getShape(bWord, oShape.getDrawingDocument(), oShape.drawingObjects);
    oShape.setStyle(null);
    oShape.spPr.setFill(AscFormat.CreateUnfilFromRGB(255, 255, 255));
	var oBodypr = oShape.getBodyPr().createDuplicate();
	oBodypr.lIns = 0;
	oBodypr.tIns = 0;
	oBodypr.rIns = 0;
	oBodypr.bIns = 0;
	oBodypr.anchor = 1;
	if(!bWord)
	{
		oShape.txBody.setBodyPr(oBodypr);
	}
	else
	{
		oShape.setBodyPr(oBodypr);
	}
	oShape.spPr.setLn(AscFormat.CreatePenFromParams(CreateNoFillUniFill(), null, null, null, 2, null));
	if(oWorkSheet)
	{
		oShape.setWorksheet(oWorkSheet);
	}
	if(oParent)
	{
		oShape.setParent(oParent);
	}
	oShape.spPr.xfrm.setOffX(0);
	oShape.spPr.xfrm.setOffY(0);
	oShape.spPr.xfrm.setExtX(this.shapeWidth);
	oShape.spPr.xfrm.setExtY(this.shapeHeight);
	oShape._page = 0;
	return oShape;
};

TextArtPreviewManager.prototype.getTAShape = function()
{
	if (!this.TAShape)
	{
		var MainLogicDocument = (editor && editor.WordControl && editor.WordControl.m_oLogicDocument ? editor && editor.WordControl && editor.WordControl.m_oLogicDocument : null);
		var TrackRevisions    = false;
		if (MainLogicDocument && MainLogicDocument.IsTrackRevisions && MainLogicDocument.IsTrackRevisions())
		{
			TrackRevisions = MainLogicDocument.GetLocalTrackRevisions();
			MainLogicDocument.SetLocalTrackRevisions(false);
		}

		var oShape = this.getShape();
		if (!oShape)
		{
			if (false !== TrackRevisions)
				MainLogicDocument.SetLocalTrackRevisions(TrackRevisions);

			return null;
		}

		var oContent = oShape.getDocContent();
		if (oContent)
		{
			if (oContent.MoveCursorToStartPos)
			{
				oContent.MoveCursorToStartPos();
			}
			oContent.AddText("Ta");
			oContent.SetApplyToAll(true);
			oContent.AddToParagraph(new ParaTextPr({FontSize : 109, RFonts : {Ascii : {Name : "Arial", Index : -1}}}));
			oContent.SetParagraphAlign(AscCommon.align_Center);
			oContent.SetParagraphIndent({FirstLine : 0, Left : 0, Right : 0});
			oContent.SetApplyToAll(false);
		}

		if (false !== TrackRevisions)
			MainLogicDocument.SetLocalTrackRevisions(TrackRevisions);

		this.TAShape = oShape;
	}
	return this.TAShape;
};

TextArtPreviewManager.prototype.getWordArtPreview = function(prst)
{
	return this.getWordArtPreviewCanvas(prst).toDataURL("image/png");
};
TextArtPreviewManager.prototype.getWordArtPreviews = function()
{
	return AscFormat.ExecuteNoHistory(function(){
		var aRet = [];
		for(var nIdx = 0; nIdx < AscCommon.g_aTextArtPresets.length; ++nIdx)
		{
			var sPreset = AscCommon.g_aTextArtPresets[nIdx];
			var oPreview = {};
			oPreview["Type"] = sPreset;
			oPreview["Image"] = this.getWordArtPreview(sPreset);
			aRet.push(oPreview);
		}
		return aRet;
	}, this, []);
};
TextArtPreviewManager.prototype.getWordArtPreviewCanvas = function(prst)
{
	var _canvas = this.createCanvas();
	var ctx = _canvas.getContext('2d');
	var graphics = new AscCommon.CGraphics();
	var oShape = this.getShapeByPrst(prst);
    if(!oShape)
    {
        return "";
    }
	graphics.init(ctx, _canvas.width, _canvas.height, oShape.extX, oShape.extY);
	graphics.m_oFontManager = AscCommon.g_fontManager;
	graphics.transform(1,0,0,1,0,0);

	var oldShowParaMarks;
	if(editor)
	{
		oldShowParaMarks = editor.ShowParaMarks;
		editor.ShowParaMarks = false;
	}
	oShape.draw(graphics);

	if(editor)
	{
		editor.ShowParaMarks = oldShowParaMarks;
	}
	return _canvas;
};

TextArtPreviewManager.prototype.generateTextArtStyles = function()
{
    AscFormat.ExecuteNoHistory(function(){

        if(this.aStylesByIndex.length === 0)
        {
            this.initStyles();
        }
        var _canvas = this.getCanvas();
        var ctx = _canvas.getContext('2d');
        var graphics = new AscCommon.CGraphics();
        var oShape = this.getTAShape();
        if(!oShape)
        {
            this.TextArtStyles.length = 0;
            return;
        }
        oShape.recalculate();

        graphics.m_oFontManager = AscCommon.g_fontManager;

        var oldShowParaMarks;
        if(editor)
        {
            oldShowParaMarks = editor.ShowParaMarks;
            editor.ShowParaMarks = false;
        }
        var oContent = oShape.getDocContent();
        oContent.SetApplyToAll(true);
        for(var i = 0; i < this.aStylesByIndex.length; ++i)
        {
            oContent.AddToParagraph(new ParaTextPr(this.aStylesByIndex[i]));
            graphics.init(ctx, _canvas.width, _canvas.height, oShape.extX, oShape.extY);
            graphics.transform(1,0,0,1,0,0);
            oShape.recalcText();
            if(!oShape.bWordShape)
            {
                oShape.recalculate();
            }
            else
            {
                oShape.recalculateText();
            }
            oShape.draw(graphics);
            this.TextArtStyles[i] = _canvas.toDataURL("image/png");
        }
        oContent.SetApplyToAll(false);

        if(editor)
        {
            editor.ShowParaMarks = oldShowParaMarks;
        }
    }, this, []);
};


	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].ChartPreviewManager = ChartPreviewManager;
	window['AscCommon'].SmartArtPreviewDrawer = SmartArtPreviewDrawer;
	window['AscCommon'].TextArtPreviewManager = TextArtPreviewManager;
	// window['AscCommon'].createPreviewSmartArt = function () {
	// 	const SmartArtDrawer = new SmartArtPreviewDrawer();
	// 	SmartArtDrawer.start();
	// }
})(window);
