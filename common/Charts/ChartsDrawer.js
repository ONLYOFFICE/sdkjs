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
var cToDeg = AscFormat.cToDeg;
var ORIENTATION_MIN_MAX = AscFormat.ORIENTATION_MIN_MAX;
var ORIENTATION_MAX_MIN = AscFormat.ORIENTATION_MAX_MIN;
var Point3D = AscFormat.Point3D;

var c_oAscTickMark = Asc.c_oAscTickMark;
var c_oAscChartDataLabelsPos = Asc.c_oAscChartDataLabelsPos;
var c_oAscChartLegendShowSettings = Asc.c_oAscChartLegendShowSettings;

var test_compare_paths = false;
var test_compare_paths_arr;

var c_oChartTypes =
{
	Bar: 0,
	Line: 1,
	HBar: 2,
	Pie: 3,
	Scatter: 4,
	Area: 5,
	Stock: 6,
	DoughnutChart: 7,
	Radar: 8,
	BubbleChart: 9,
	Surface: 10
};

var c_oChartBar3dFaces =
{
	front: 0,
	up: 1,
	left: 2,
	right: 3,
	down: 4,
	back: 5
};


var globalGapDepth = 150;
var isTurnOn3DCharts = true;
var standartMarginForCharts = 13;

function arrReverse(arr) {
	if(!arr || !arr.length)
		return;
	var newarr = [];
	for (var i = 0; i < arr[0].length; ++i) {
		newarr[i] = [];
		for (var j = 0; j < arr.length; ++j) {
			newarr[i][j] = arr[j][i];
		}
	}
	return newarr;
}

//*****MAIN*****
function CChartsDrawer()
{
	this.calcProp = {};

	this.areaChart = null;
	this.chart = null;
	this.cChartSpace = null;
	this.cShapeDrawer = null;
	this.processor3D = null;

	this.areaChart = null;
	this.axesChart = null;

	this.floor3DChart = null;
	this.sideWall3DChart = null;
	this.backWall3DChart = null;

	this.errBars = new CErrBarsDraw(this);
	this.trendline = new CTrendline(this);

	this.changeAxisMap = null;

	this.calcShapesHelper = new calcShapesHelper(this);
}

CChartsDrawer.prototype =
{
    constructor: CChartsDrawer,

	//****draw and recalculate functions****
	recalculate : function (chartSpace) {
		this.cChartSpace = chartSpace;

		this.calcProp = {};

		//nDimensionCount - flag for 3d/2d charts
		if (this._isSwitchCurrent3DChart(chartSpace)) {
			standartMarginForCharts = 16;
			this.nDimensionCount = 3;
		} else {
			this.nDimensionCount = 2;
		}

		if (!chartSpace.bEmptySeries) {
			this._calculateProperties(chartSpace);
		}

		if (this.calcProp.widthCanvas == undefined && this.calcProp.pxToMM == undefined) {
			this.calcProp.pathH = 1000000000;
			this.calcProp.pathW = 1000000000;
			this.calcProp.pxToMM = 1 / AscCommon.g_dKoef_pix_to_mm;
			this.calcProp.widthCanvas = chartSpace.extX * this.calcProp.pxToMM;
			this.calcProp.heightCanvas = chartSpace.extY * this.calcProp.pxToMM;
		}

		this.init(chartSpace);

		//****recalculate****
		//PLOTAREA
		if (!chartSpace.bEmptySeries) {
			if (this.nDimensionCount === 3) {
				this._calaculate3DProperties(chartSpace);
			}

			this.plotAreaChart.recalculate(this);
		}

		//AREA
		this.areaChart.recalculate(this);

		//AXIS
		if (!chartSpace.bEmptySeries) {

			//оси значений и категорий
			this.axesChart = [];
			for (var i = 0; i < chartSpace.chart.plotArea.axId.length; i++) {
				var axId = chartSpace.chart.plotArea.axId[i];
				var axis = new axisChart();
				axis.recalculate(this, axId);
				this.axesChart.push(axis);
			}

			if (this.nDimensionCount === 3) {
				this.floor3DChart.recalculate(this);
				this.sideWall3DChart.recalculate(this);
				this.backWall3DChart.recalculate(this);
			}

		}

		//CHARTS
		if (!chartSpace.bEmptySeries) {
			for(var i in this.charts) {
				if (this.charts.hasOwnProperty(i) && this.charts[i]) {
					this.charts[i].recalculate();
				}
			}
		}

		if (!chartSpace.bEmptySeries) {
			this.errBars.recalculate(this.charts);
		}

		if(!chartSpace.bEmptySeries){
			this.trendline.recalculate(this.charts);
		}
		//for test
		//this._testChartsPaths();
	},

	getStandartMargin: function () {
		return this.calcProp && AscFormat.isRealNumber(this.calcProp.pxToMM) ? standartMarginForCharts / this.calcProp.pxToMM : 0;
	},

	createChartEx: function (seria) {
		this.charts = {};
		switch (seria.layoutId) {
			case AscFormat.SERIES_LAYOUT_CLUSTERED_COLUMN :
				this.charts.chartEx = new drawClusteredColumn(seria, this);
				break;
			case AscFormat.SERIES_LAYOUT_WATERFALL :
				this.charts.chartEx = new drawWaterfallChart(seria, this);
				break;
			case AscFormat.SERIES_LAYOUT_FUNNEL :
				this.charts.chartEx = new drawFunnelChart(seria, this);
				break;
			case AscFormat.SERIES_LAYOUT_PARETO_LINE :
				this.charts.chartEx = new drawParetoChart(seria, this);
				break;
			default :
				this.charts.chartEx = null;
		}
	},

	init: function(chartSpace) {
		//создаём область
		this.areaChart = new areaChart();
		//создаём область
		this.plotAreaChart = new plotAreaChart();
		//Floor This element specifies the floor of a 3D chart.
		this.floor3DChart = new floor3DChart();
		this.sideWall3DChart = new sideWall3DChart();
		this.backWall3DChart = new backWall3DChart();

		//draw chart
		let plotArea = chartSpace.chart.plotArea;
		let newChart;
		if (plotArea.isChartEx()) {
			const series = plotArea.plotAreaRegion.series;
			let maxSeria = null;
			// find the maximum seria, that will indicate the resulting chartExFormat
			for (let i = 0; i < series.length; i++) {
				if (maxSeria === null || maxSeria.layoutId < series[i].layoutId) {
					maxSeria = series[i];
				}
			}
			this.createChartEx(maxSeria);
		} else {
			for (let i = 0; i < plotArea.charts.length; i++) {
				let chart = plotArea.charts[i];
				switch (this._getChartType(chart)) {
					case c_oChartTypes.Bar: {
						newChart = new drawBarChart(chart, this);
						break;
					}
					case c_oChartTypes.Line: {
						newChart = new drawLineChart(chart, this);
						break;
					}
					case c_oChartTypes.HBar: {
						newChart = new drawHBarChart(chart, this);
						break;
					}
					case c_oChartTypes.Pie: {
						newChart = new drawPieChart(chart, this);
						break;
					}
					case c_oChartTypes.Scatter: {
						newChart = new drawScatterChart(chart, this);
						break;
					}
					case c_oChartTypes.Area: {
						newChart = new drawAreaChart(chart, this);
						break;
					}
					case c_oChartTypes.Stock: {
						newChart = new drawStockChart(chart, this);
						break;
					}
					case c_oChartTypes.DoughnutChart: {
						newChart = new drawDoughnutChart(chart, this);
						break;
					}
					case c_oChartTypes.Radar: {
						newChart = new drawRadarChart(chart, this);
						break;
					}
					case c_oChartTypes.BubbleChart: {
						newChart = new drawBubbleChart(chart, this);
						break;
					}
					case c_oChartTypes.Surface: {
						newChart = new drawSurfaceChart(chart, this);
						break;
					}
				}
				if (i === 0) {
					this.chart = newChart;
					this.charts = {};
				}
				this.charts[chart.Id] = newChart;
			}
		}
	},

	draw: function (chartSpace, graphics) {
		var t = this;

		this.cChartSpace = chartSpace;

		var cShapeDrawer = new AscCommon.CShapeDrawer();
		cShapeDrawer.Graphics = graphics;

		this.cShapeDrawer = cShapeDrawer;

		//отрисовываем без пересчёта
		this.areaChart.draw(this);

		var drawCharts = function(bBeforeAxes) {
			//рисуем 3d диаграммы только до отрисовки сетки
			if(!bBeforeAxes && t.nDimensionCount === 3) {
				return;
			}
			if (!chartSpace || !chartSpace.chart || !chartSpace.chart.plotArea) {
				return;
			}

			if (chartSpace.isChartEx()) {
				for (i in t.charts) {
					if (t.charts.hasOwnProperty(i) && t.charts[i]) {
						t.charts[i].draw();
					}
				}
			} else {
				//для начала нужно отсортировать
				var sortCharts = t._sortChartsForDrawing(chartSpace);
				for(var i = 0; i < sortCharts.length; i++) {
					var id = sortCharts[i];
					var chartModel = t._getChartModelById(chartSpace.chart.plotArea, id);
					if(!chartModel) {
						continue;
					}
	
					var type = chartModel.getObjectType();
					var isLinesChart = type === AscDFH.historyitem_type_LineChart || type === AscDFH.historyitem_type_ScatterChart;
					//рисуем линейные диаграммы после отрисовки сетки
					if(t.nDimensionCount !== 3 && ((isLinesChart && bBeforeAxes) || (!isLinesChart && !bBeforeAxes))) {
						continue;
					}
	
					var bIsNoSmartAttack = false;
					if(t.nDimensionCount === 3 || isLinesChart) {
						bIsNoSmartAttack = true;
					}
	
					if(bIsNoSmartAttack) {
						t.cShapeDrawer.bIsNoSmartAttack = true;
					}
	
					t.charts[id].draw();
					if(bIsNoSmartAttack) {
						t.cShapeDrawer.bIsNoSmartAttack = false;
					}
				}
			}
		};

		var i;
		if (!chartSpace.bEmptySeries) {
			this.plotAreaChart.draw(this, true);

			if (this.calcProp.type !== c_oChartTypes.Pie && this.calcProp.type !== c_oChartTypes.DoughnutChart) {
				if (this.nDimensionCount === 3) {
					this.floor3DChart.draw(this);
					this.sideWall3DChart.draw(this);
					this.backWall3DChart.draw(this);
				}
				// GRID
				for(i = 0; i < this.axesChart.length; i++) {
					this.axesChart[i].draw(this, true);
				}
				this.plotAreaChart.draw(this, null, true);
			}

			// DRAW 3D CHARTS
			// рисуем оси поверх 3d-диаграмм и линейных/точечных
			drawCharts(true);

			for(i = 0; i < this.axesChart.length; i++) {
				this.axesChart[i].draw(this);
			}

			//DRAW CHARTS
			drawCharts();

			//err bars
			this.errBars.draw();

			//draw trendline
			this.trendline.draw();
		}
	},

	_testChartsPaths: function() {
		//чтобы сгенерировать все paths нужно выставить buildAllPaths в true
		//добавляем данные из консоли(console.log(JSON.stringify(test_compare_paths_arr));) в буфер обмена
		//далее открыаем файл, запускаем макрос - получаем готовый файл с текстом наших paths в колонке с индексом col

		var buildAllPaths = false;
		var col = 41;
		if(test_compare_paths) {
			if(undefined === window.test_compare_paths_count) {
				window.test_compare_paths_count = 0;
			}
			if (!test_compare_paths_arr) {
				var row = 0;
				var str = "";
				while (true) {
					var val = this.cChartSpace.worksheet.getCell3(row, col).getValue();
					if ("" !== val) {
						str += val;
						row++;
					} else {
						break;
					}
				}
				if("" !== str) {
					test_compare_paths_arr = JSON.parse(str);
				}
			}

			if(!test_compare_paths_arr) {
				test_compare_paths_arr = [];
			}
			if(test_compare_paths_arr[window.test_compare_paths_count] !== JSON.stringify(this.cChartSpace.GetPath().ArrPathCommand)) {
				console.log("error drawing charts" + window.test_compare_paths_count);
			}
			if(buildAllPaths) {
				test_compare_paths_arr[window.test_compare_paths_count] = JSON.stringify(this.cChartSpace.GetPath().ArrPathCommand);
				console.log(JSON.stringify(test_compare_paths_arr));
			}
			window.test_compare_paths_count++;
		}

		//с помощью данного макроса добавляю paths в файл
		/*Private Sub CommandButton1_Click()

		Dim DataObj As MSForms.DataObject
		Set DataObj = New MSForms.DataObject
		DataObj.GetFromClipboard

		SText = DataObj.GetText(1)

		Dim columnNameStart As String
		Dim rowNameStart As String

		Dim test As String, row As Integer
		For i = 1 To Len(SText) Step 1000
		test = Mid(SText, i, 1000)
		ActiveSheet.Cells(row + 1, 42).Value = test
		ActiveSheet.Cells(row + 1, 42).WrapText = False

		row = row + 1
		Next

		End Sub*/

	},

	_sortChartsForDrawing: function(chartSpace) {
		var arr = [];

		var pushIndex = function(index) {
			if(!arr[index]) {
				arr[index] = [];
			}
			arr[index].push(i);
		};

		for(var i in this.charts) {
			var chartModel = this._getChartModelById(chartSpace.chart.plotArea, i);
			if(!chartModel) {
				continue;
			}

			var type = chartModel.getObjectType();
			switch(type) {
				case AscDFH.historyitem_type_DoughnutChart: {
					pushIndex(0);
					break;
				}
				case AscDFH.historyitem_type_PieChart: {
					pushIndex(1);
					break;
				}
				case AscDFH.historyitem_type_AreaChart: {
					pushIndex(2);
					break;
				}
				case AscDFH.historyitem_type_BarChart: {
					pushIndex(3);
					break;
				}
				case AscDFH.historyitem_type_StockChart: {
					pushIndex(4);
					break;
				}
				case AscDFH.historyitem_type_LineChart: {
					pushIndex(5);
					break;
				}
				case AscDFH.historyitem_type_ScatterChart: {
					pushIndex(6);
					break;
				}
				case AscDFH.historyitem_type_RadarChart: {
					pushIndex(7);
					break;
				}
				default: {
					//bubble, stock
					pushIndex(8);
					break;
				}
			}
		}
		var sortArr = [];
		for(var j = 0; j < arr.length; j++) {
			if(arr[j]) {
				for(var k = 0; k < arr[j].length; k++) {
					sortArr.push(arr[j][k]);
				}
			}
		}

		return sortArr;
	},

	recalculateOnly3dProps: function (chartSpace) {
		//this.cChartSpace = chartSpace;
		this.calcProp = {};

		if (this._isSwitchCurrent3DChart(chartSpace)) {
			standartMarginForCharts = 16;
			this.nDimensionCount = 3;

			if (!chartSpace.bEmptySeries) {
				this._calculateProperties(chartSpace);
			}

			if (this.calcProp.widthCanvas == undefined && this.calcProp.pxToMM == undefined) {
				this.calcProp.pathH = 1000000000;
				this.calcProp.pathW = 1000000000;
				this.calcProp.pxToMM = 1 / AscCommon.g_dKoef_pix_to_mm;
				this.calcProp.widthCanvas = chartSpace.extX * this.calcProp.pxToMM;
				this.calcProp.heightCanvas = chartSpace.extY * this.calcProp.pxToMM;
			}

			if (!chartSpace.bEmptySeries) {
				if (this.nDimensionCount === 3) {
					this._calaculate3DProperties(chartSpace);
				}
			}
		}
	},

	_getChartModelById: function(plotArea, id) {
		var res = null;

		var charts = plotArea.charts;
		if(charts) {
			for(var i = 0; i < charts.length; i ++) {
				if(id === charts[i].Id) {
					res = charts[i];
					break;
				}
			}
		}

		return res;
	},

	_getSeriesById: function(oChart, id) {
		var res = null;

		var series = oChart.series;
		if(series) {
			for(var i = 0; i < series.length; i ++) {
				if(id === series[i].Id) {
					res = series[i];
					break;
				}
			}
		}

		return res;
	},

	_getChartModelIdBySerIdx: function(plotArea, idx) {
		var res = null;

		var charts = plotArea.charts;
		if(charts) {
			for(var i = 0; i < charts.length; i ++) {
				for(var j = 0; j < charts[i].series.length; j++) {
					if(idx === charts[i].series[j].idx) {
						res = charts[i].Id;
						break;
					}
				}
			}
		}

		return res;
	},

	_getIndexByIdxSeria: function(series, val) {
		for(var i = 0; i < series.length; i++) {
			if(series[i].idx === val)
				return i;
		}
	},

	//****positions text labels****
	recalculatePositionText: function (obj, idx) {
		var pos;

		if (!this.cChartSpace.bEmptySeries) {
			var type = obj.getObjectType();

			switch (type) {
				case AscDFH.historyitem_type_DLbl: {
					pos = this._calculatePositionDlbl(obj);
					break;
				}
				case AscDFH.historyitem_type_Title: {
					var parentType = obj.parent ? obj.parent.getObjectType() : null;
					if (parentType === AscDFH.historyitem_type_Chart) {
						pos = this._calculatePositionTitle(obj);
					} else if (parentType === AscDFH.historyitem_type_ValAx || parentType === AscDFH.historyitem_type_CatAx || parentType === AscDFH.historyitem_type_DateAx || parentType === AscDFH.historyitem_type_Axis) {
						pos = this._calculatePositionAxisTitle(obj.parent);
					}
					break;
				}
				case AscDFH.historyitem_type_CatAx: {
					if ((obj.chart && obj.chart.chart && obj.chart.chart.plotArea && obj.chart.chart.plotArea.chart) &&
						(obj.chart.chart.plotArea.chart.getObjectType() === AscDFH.historyitem_type_RadarChart)) {
						pos = this._calculatePositionLabelsCatAxForRadar(idx);
					}
					break;
				}
				case AscDFH.historyitem_type_ValAx:
				case AscDFH.historyitem_type_DateAx: {
					break;
				}
				case AscDFH.historyitem_type_Legend: {
					pos = this._calculatePositionLegend(obj);
					break;
				}
				case AscDFH.historyitem_type_TrendLine: {
					/*{
						coefficients : coefficients,
							rSquared : rSquared,
						coordinate : lastPoint
					}*/
					if (this.trendline) {
						//CTrendLine
						return this.trendline && this.trendline.getAdditionalInfo(obj.parent.parent.Id, obj.parent.Id);
					}
					return null;
				}
				default: {
					pos = { x: 0, y: 0 };
					break;
				}
			}
		}

		return {x: pos ? pos.x : undefined, y: pos ? pos.y : undefined};
	},
	
	_calculatePositionDlbl: function(obj/*chartSpace, ser, val, bLayout*/)
	{
		let res = null;
		if (obj.showChartExVal) {
			if (this.charts && this.charts.chartEx) {
				res = this.charts.chartEx._calculateDLbl(obj);
			}
		} else {
			var chartSpace = obj.chart;
			var bLayout = AscCommon.isRealObject(obj.layout) && (AscFormat.isRealNumber(obj.layout.x) || AscFormat.isRealNumber(obj.layout.y));
			var serIdx = obj.series.idx;
			var valIdx = obj.pt.idx;
	
			var chartId = this._getChartModelIdBySerIdx(chartSpace.chart.plotArea, serIdx);
			if(null !== chartId && this.charts[chartId] && this.charts[chartId].chart && this.charts[chartId].chart.series) {
				//TODO нужно переделать все массивы с патами по idx
				var serIndex = this._getIndexByIdxSeria(this.charts[chartId].chart.series, serIdx);
				res = this.charts[chartId]._calculateDLbl(chartSpace, serIndex, valIdx, bLayout, serIdx);
			}
		}

		return res;
	},
	
	_calculatePositionTitle: function(title)
	{	
		var widthGraph = title.chart.extX;

		var widthTitle = title.extX;
		var standartMargin = 7;
		
		var y = standartMargin / this.calcProp.pxToMM;
		var x = widthGraph / 2 - widthTitle / 2;
		
		return {x: x, y: y}
	},

	_calculatePositionAxisTitle: function(axis)
	{
		var chart = axis.parent && axis.parent.parent ? axis.parent.parent : null;
		var legend = chart ? chart.legend : null;
		var title = chart ? chart.title : null;
		var x = 0, y = 0, pxToMM = this.calcProp.pxToMM;

		if(axis && axis.title) {
			var widthAxisTitle = axis.title.extX;
			var heightAxisTitle = axis.title.extY;

			var layout = this.cChartSpace.chart.plotArea.layout;
			if(layout && AscFormat.isRealNumber(layout.y) && AscFormat.isRealNumber(layout.h)) {
				var x1 = layout.x * this.calcProp.widthCanvas;
				var y1 = layout.y * this.calcProp.heightCanvas;
				var w = layout.w * this.calcProp.widthCanvas;
				var h = layout.h * this.calcProp.heightCanvas;
				var hLabels = axis.labels ? axis.labels.extY : 0;
				switch (axis.axPos) {
					case window['AscFormat'].AX_POS_B: {
						y = (y1 + h) / pxToMM + hLabels;
						if(y + heightAxisTitle > this.calcProp.heightCanvas / pxToMM) {
							y = this.calcProp.heightCanvas / pxToMM - heightAxisTitle;
						}
						x = (this.calcProp.chartGutter._left + this.calcProp.trueWidth / 2) / pxToMM - widthAxisTitle / 2;

						break;
					}
					case window['AscFormat'].AX_POS_T: {
						y = standartMarginForCharts / pxToMM;
						x = (this.calcProp.chartGutter._left + this.calcProp.trueWidth / 2) / pxToMM - widthAxisTitle / 2;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.top) {
							y += legend.extY + (standartMarginForCharts / 2) / pxToMM;
						}
						if (title !== null && !title.overlay) {
							y += title.extY + (standartMarginForCharts / 2) / pxToMM;
						}
						break;
					}
					case window['AscFormat'].AX_POS_L: {
						y = (this.calcProp.chartGutter._top + this.calcProp.trueHeight / 2) / pxToMM - heightAxisTitle / 2;
						x = standartMarginForCharts / pxToMM;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.left) {
							x += legend.extX;
						}
						break;
					}
					case window['AscFormat'].AX_POS_R: {
						y = (this.calcProp.chartGutter._top + this.calcProp.trueHeight / 2) / pxToMM - heightAxisTitle / 2;
						x = (this.calcProp.widthCanvas - standartMarginForCharts) / pxToMM - widthAxisTitle;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.right) {
							x -= legend.extX;
						}
						break;
					}
				}
			} else {
				switch (axis.axPos) {
					case window['AscFormat'].AX_POS_B: {
						y = (this.calcProp.heightCanvas - standartMarginForCharts) / pxToMM - heightAxisTitle;
						x = (this.calcProp.chartGutter._left + this.calcProp.trueWidth / 2) / pxToMM - widthAxisTitle / 2;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.bottom) {
							y -= legend.extY + (standartMarginForCharts / 2) / pxToMM;
						}
						break;
					}
					case window['AscFormat'].AX_POS_T: {
						y = standartMarginForCharts / pxToMM;
						x = (this.calcProp.chartGutter._left + this.calcProp.trueWidth / 2) / pxToMM - widthAxisTitle / 2;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.top) {
							y += legend.extY + (standartMarginForCharts / 2) / pxToMM;
						}
						if (title !== null && !title.overlay) {
							y += title.extY + (standartMarginForCharts / 2) / pxToMM;
						}
						break;
					}
					case window['AscFormat'].AX_POS_L: {
						y = (this.calcProp.chartGutter._top + this.calcProp.trueHeight / 2) / pxToMM - heightAxisTitle / 2;
						x = standartMarginForCharts / pxToMM;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.left) {
							x += legend.extX;
						}
						break;
					}
					case window['AscFormat'].AX_POS_R: {
						y = (this.calcProp.chartGutter._top + this.calcProp.trueHeight / 2) / pxToMM - heightAxisTitle / 2;
						x = (this.calcProp.widthCanvas - standartMarginForCharts) / pxToMM - widthAxisTitle;

						if(legend && legend.legendPos === c_oAscChartLegendShowSettings.right) {
							x -= legend.extX;
						}
						break;
					}
				}
			}

			if(this.nDimensionCount === 3)
			{
				var convertResult;
				if(axis instanceof AscFormat.CCatAx) {
					convertResult = this._convertAndTurnPoint((x + widthAxisTitle / 2) * pxToMM, y * pxToMM, this.processor3D.calculateZPositionCatAxis());
					x = convertResult.x / pxToMM - widthAxisTitle / 2;
				} else {
					//TODO избавиться от привзяки к типу диаграммы
					if(!this.processor3D.view3D.getRAngAx() && (this.calcProp.type === c_oChartTypes.Bar || this.calcProp.type === c_oChartTypes.Line))
					{
						var posX = axis.posX;
						var widthLabels = axis.labels && axis.labels.extX ? axis.labels.extX : 0;
						var yPoints = axis.yPoints;
						var upYPoint = yPoints && yPoints[0] ? yPoints[0].pos : 0;
						var downYPoint = yPoints && yPoints[yPoints.length - 1] ? yPoints[yPoints.length - 1].pos : 0;

						var tempX = posX - widthLabels;
						var convertResultX = this._convertAndTurnPoint(tempX * pxToMM, y * pxToMM, 0);
						x = convertResultX.x / pxToMM - widthAxisTitle;

						var convertResultY1 = this._convertAndTurnPoint(posX * pxToMM, upYPoint * pxToMM, 0);
						var convertResultY2 = this._convertAndTurnPoint(posX * pxToMM, downYPoint * pxToMM, 0);
						var heightPerspectiveChart = convertResultY1.y - convertResultY2.y;

						y = (convertResultY2.y + heightPerspectiveChart / 2) / pxToMM - heightAxisTitle / 2;
					}
					else
					{
						convertResult = this._convertAndTurnPoint(x * pxToMM, y * pxToMM, 0);
						y = convertResult.y / pxToMM;
					}
				}

			}
		}

		return {x: x , y: y};
	},

	_calculatePositionLegend: function (legend) {
		var widthLegend = legend.extX;
		var heightLegend = legend.extY;
		var x, y;

		var nLegendPos = legend.legendPos !== null ? legend.legendPos : c_oAscChartLegendShowSettings.right;
		switch (nLegendPos) {
			case c_oAscChartLegendShowSettings.left:
			case c_oAscChartLegendShowSettings.leftOverlay: {
				x = standartMarginForCharts / 2 / this.calcProp.pxToMM;
				y = this.calcProp.heightCanvas / 2 / this.calcProp.pxToMM - heightLegend / 2;
				break;
			}
			case c_oAscChartLegendShowSettings.top: {
				x = this.calcProp.widthCanvas / 2 / this.calcProp.pxToMM - widthLegend / 2;
				y = standartMarginForCharts / 2 / this.calcProp.pxToMM;

				if (legend.parent.title !== null && !legend.parent.title.overlay) {
					y += legend.parent.title.extY + standartMarginForCharts / 2 / this.calcProp.pxToMM;
				}
				break;
			}
			case c_oAscChartLegendShowSettings.right:
			case c_oAscChartLegendShowSettings.rightOverlay: {
				x = (this.calcProp.widthCanvas - standartMarginForCharts / 2) / this.calcProp.pxToMM - widthLegend;
				y = (this.calcProp.heightCanvas / 2) / this.calcProp.pxToMM - heightLegend / 2;
				break;
			}
			case c_oAscChartLegendShowSettings.bottom: {
				x = this.calcProp.widthCanvas / 2 / this.calcProp.pxToMM - widthLegend / 2;
				y = (this.calcProp.heightCanvas - standartMarginForCharts / 2) / this.calcProp.pxToMM - heightLegend;
				break;
			}
			case c_oAscChartLegendShowSettings.topRight: {
				x = (this.calcProp.widthCanvas - standartMarginForCharts / 2) / this.calcProp.pxToMM - widthLegend;
				y = standartMarginForCharts / 2 / this.calcProp.pxToMM;

				if (legend.parent.title !== null && !legend.parent.title.overlay) {
					y += legend.parent.title.extY + standartMarginForCharts / 2 / this.calcProp.pxToMM;
				}
				break;
			}
			default: {
				x = (this.calcProp.widthCanvas - standartMarginForCharts / 2) / this.calcProp.pxToMM - widthLegend;
				y = (this.calcProp.heightCanvas) / this.calcProp.pxToMM - heightLegend / 2;
				break;
			}
		}

		return {x: x, y: y};
	},

	_calculatePositionLabelsCatAxForRadar: function (idx) {
		let chartProp = this.calcProp;

		let trueWidth = chartProp.trueWidth;
		let trueHeight = chartProp.trueHeight;
		let xCenter = (chartProp.chartGutter._left + trueWidth / 2) / chartProp.pxToMM;
		let yCenter = (chartProp.chartGutter._top + trueHeight / 2) / chartProp.pxToMM;

		let radarChart = this.getChartByType(AscDFH.historyitem_type_RadarChart);
		let ptCount = radarChart && radarChart.series && this.cChartDrawer.getPtCount(radarChart.series);

		let angle = (Math.PI * 2 / ptCount) * idx;
		let radius = (trueHeight / 2) / chartProp.pxToMM;
		let startAngle = (Math.PI / 2) * 3 + angle;

		let x = xCenter + radius * Math.cos(startAngle);
		let y = yCenter + radius * Math.sin(startAngle);

		return { x: x, y: y };
	},
	//****calculate margins****
	_calculateMarginsChart: function (chartSpace, dNotPutResult) {
		this.calcProp.chartGutter = {};

		if (this._isSwitchCurrent3DChart(chartSpace)) {
			standartMarginForCharts = 16;
		}
		if (!this.calcProp.pxToMM) {
			this.calcProp.pxToMM = 1 / AscCommon.g_dKoef_pix_to_mm;
		}

		var pxToMM = this.calcProp.pxToMM;
		var plotArea = chartSpace.chart.plotArea;

		//если точки рассчитаны - ставим маргин в зависимости от них
		var marginOnPoints = this._calculateMarginOnPoints(chartSpace/*, isHBar*/);
		var calculateLeft = marginOnPoints.calculateLeft, calculateRight = marginOnPoints.calculateRight, calculateTop = marginOnPoints.calculateTop, calculateBottom = marginOnPoints.calculateBottom;

		//высчитываем выходящие за пределы подписи осей
		var labelsMargin = this._calculateMarginLabels(chartSpace);
		var left = labelsMargin.left, right = labelsMargin.right, top = labelsMargin.top, bottom = labelsMargin.bottom;


		var leftTextLabels = 0;
		var rightTextLabels = 0;
		var topTextLabels = 0;
		var bottomTextLabels = 0;

		//добавляем размеры подписей осей + размеры названия
		//TODO генерировать extX для всех осей
		var axId = chartSpace.chart.plotArea.axId;
		let isLeftAxis = false;
		let isRightAxis = false;
		let isTopExist = chartSpace && chartSpace.chart ? !!chartSpace.chart.title || (!!chartSpace.chart.legend && chartSpace.chart.legend.legendPos === Asc.c_oAscChartLegendShowSettings.top) : false;
		let isVertAxisExist = false;
		if(axId) {
			for(var i = 0; i < axId.length; i++) {
				switch (axId[i].axPos) {
					case window['AscFormat'].AX_POS_B: {
						if (null !== axId[i].title) {
							bottomTextLabels += axId[i].title.extY;
						}
						break;
					}
					case window['AscFormat'].AX_POS_T: {
						if (null !== axId[i].title) {
							topTextLabels += axId[i].title.extY;
						}
						break;
					}
					case window['AscFormat'].AX_POS_L: {
						if (null !== axId[i].title) {
							leftTextLabels += axId[i].title.extX;
						}
						break;
					}
					case window['AscFormat'].AX_POS_R: {
						if (null !== axId[i].title) {
							rightTextLabels += axId[i].title.extX;
						}
						break;
					}
				}
				const bLeftVert = axId[i].axPos === AscFormat.AX_POS_L;
				const bRigthVert = axId[i].axPos === AscFormat.AX_POS_R;
				if (!axId[i].bDelete) {
					if (bLeftVert) {
						isLeftAxis = axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_NEXT_TO || axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_LOW;
						isRightAxis = axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_HIGH;
					}
					if (bRigthVert) {
						isLeftAxis = axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_HIGH;
						isRightAxis = axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_NEXT_TO || axId[i].tickLblPos === AscFormat.TICK_LABEL_POSITION_LOW;
					}
				}
				if (!axId[i].bDelete && axId[i].tickLblPos !== Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE && (bLeftVert || bRigthVert) && !isVertAxisExist) {
					isVertAxisExist = true;
				}
			}
		}

		// if no vetical axis when top is not exist, then should work as if top exist
		isTopExist = !isTopExist && !isVertAxisExist ? true : isTopExist;

		//TITLE
		var topMainTitle = 0;
		if (chartSpace.chart.title !== null && !chartSpace.chart.title.overlay) {
			topMainTitle += chartSpace.chart.title.extY;
		}

		//KEY
		var leftKey = 0, rightKey = 0, topKey = 0, bottomKey = 0;
		if (chartSpace.chart.legend && !chartSpace.chart.legend.overlay) {
			var fLegendExtX = chartSpace.chart.legend.extX;
			var fLegendExtY = chartSpace.chart.legend.extY;
			if (chartSpace.chart.legend.layout) {
				if (AscFormat.isRealNumber(chartSpace.chart.legend.naturalWidth) && AscFormat.isRealNumber(chartSpace.chart.legend.naturalHeight)) {
					fLegendExtX = chartSpace.chart.legend.naturalWidth;
					fLegendExtY = chartSpace.chart.legend.naturalHeight;
				}
			}
			var nLegendPos = chartSpace.chart.legend.legendPos !== null ? chartSpace.chart.legend.legendPos : c_oAscChartLegendShowSettings.right;
			switch (nLegendPos) {
				case c_oAscChartLegendShowSettings.left:
				case c_oAscChartLegendShowSettings.leftOverlay: {
					leftKey += fLegendExtX;
					break;
				}
				case c_oAscChartLegendShowSettings.top: {
					topKey += fLegendExtY;
					break;
				}
				case c_oAscChartLegendShowSettings.right:
				case c_oAscChartLegendShowSettings.rightOverlay: {
					rightKey += fLegendExtX;
					break;
				}
				case c_oAscChartLegendShowSettings.bottom: {
					bottomKey += fLegendExtY;
					break;
				}
				case c_oAscChartLegendShowSettings.topRight: {
					rightKey += fLegendExtX;
					break;
				}
			}
		}

		//исключение - когда среди диаграмм есть груговая
		var pieChart = null;
		var radarChart = null;
		var charts = plotArea.charts;
		for(i = 0; i < charts.length; i++) {
			var chartType = this._getChartType(charts[i]);
			if(c_oChartTypes.Pie === chartType || c_oChartTypes.DoughnutChart === chartType) {
				pieChart = charts[i];
				break;
			} else if (c_oChartTypes.Radar === chartType) {
				radarChart = charts[i];
			}
		}
		var is3dChart = this._isSwitchCurrent3DChart(chartSpace);
		if(!is3dChart && null !== pieChart) {
			//вычисляем истинную(первоначальную) ширину и высоту диаграммы
			left = this._getStandartMargin(left, leftKey, leftTextLabels, 0) + leftKey + leftTextLabels;
			bottom = this._getStandartMargin(bottom, bottomKey, bottomTextLabels, 0) + bottomKey + bottomTextLabels;
			top = this._getStandartMargin(top, topKey, topTextLabels, topMainTitle) + topKey + topTextLabels + topMainTitle;
			right = this._getStandartMargin(right, rightKey, rightTextLabels, 0) + rightKey + rightTextLabels;

			var width = chartSpace.extX - left - right;
			var height = chartSpace.extY - top - bottom;
			var pieSize = width > height ? height : width;

			//размещаем по центру относительно width/height
			left += (width - pieSize)/2;
			right += (width - pieSize)/2;
			top += (height - pieSize)/2;
			bottom += (height - pieSize)/2;
		}

		if (radarChart && (calculateTop || calculateBottom)) {
			let radius, centerX, centerY;
			centerY = chartSpace.extY - calculateBottom;
			radius = centerY - calculateTop;

			calculateBottom = chartSpace.extY - calculateTop - 2*radius;

			let valAx = this.getAxisFromAxId(radarChart.axId, AscDFH.historyitem_type_ValAx);
			centerX = valAx && valAx.posX;
			if (centerX) {
				calculateLeft = centerX - radius;
				calculateRight = chartSpace.extX - calculateLeft - radius * 2;
			}
		}

		if((null === pieChart) || is3dChart) {
			left += this._getStandartMargin(isLeftAxis ? 1 : left, leftKey, leftTextLabels, 0) + leftKey + leftTextLabels;
			bottom += this._getStandartMargin(bottom, bottomKey, bottomTextLabels, 0) + bottomKey + bottomTextLabels;
			top += this._getStandartMargin(isTopExist ? top : 1, topKey, topTextLabels, topMainTitle) + topKey + topTextLabels + topMainTitle;
			right += this._getStandartMargin(isRightAxis ? 1 : right, rightKey, rightTextLabels, 0) + rightKey + rightTextLabels;
		}

		var pxLeft = calculateLeft ? calculateLeft * pxToMM : left * pxToMM;
		var pxRight = calculateRight ? calculateRight * pxToMM : right * pxToMM;
		var pxTop = calculateTop ? calculateTop * pxToMM : top * pxToMM;
		var pxBottom = calculateBottom ? calculateBottom * pxToMM : bottom * pxToMM;

		//TODO позже пересмотреть правку
		if(topMainTitle && topMainTitle * pxToMM > pxTop) {
			pxTop = (this._getStandartMargin(top, topKey, topTextLabels, topMainTitle) / 2 + topMainTitle) * pxToMM;
		}

		//TODO пересмотреть!!!
		if(pieChart && plotArea.charts.length === 1) {
			if (plotArea.layout) {
				var oLayout = plotArea.layout;
				pxLeft = chartSpace.calculatePosByLayout(pxLeft / pxToMM, oLayout.xMode, oLayout.x, (pxRight - pxLeft) / pxToMM, chartSpace.extX) * pxToMM;
				pxTop = chartSpace.calculatePosByLayout(pxTop / pxToMM, oLayout.yMode, oLayout.y, (pxBottom - pxTop) / pxToMM, chartSpace.extY) * pxToMM;

				var fWidthPlotArea = chartSpace.calculateSizeByLayout(pxLeft / pxToMM, chartSpace.extX, oLayout.w, oLayout.wMode);
				if (fWidthPlotArea > 0) {
					pxRight = chartSpace.extX * pxToMM - (pxLeft + fWidthPlotArea * pxToMM);
				}
				var fHeightPlotArea = chartSpace.calculateSizeByLayout(pxTop / pxToMM, chartSpace.extY, oLayout.h, oLayout.hMode);
				if (fHeightPlotArea > 0) {
					pxBottom = chartSpace.extY * pxToMM - (pxTop + fHeightPlotArea * pxToMM);
				}
			}
		}

		if (dNotPutResult) {
			return {left: pxLeft, right: pxRight, top: pxTop, bottom: pxBottom};
		} else {
			this.calcProp.chartGutter._left = pxLeft;
			this.calcProp.chartGutter._right = pxRight;
			this.calcProp.chartGutter._top = pxTop;
			this.calcProp.chartGutter._bottom = pxBottom;
		}

		this._checkMargins();
	},

	_checkMargins: function () {
		if (this.calcProp.chartGutter._left < 0) {
			this.calcProp.chartGutter._left = standartMarginForCharts;
		}
		if (this.calcProp.chartGutter._right < 0) {
			this.calcProp.chartGutter._right = standartMarginForCharts;
		}
		if (this.calcProp.chartGutter._top < 0) {
			this.calcProp.chartGutter._top = standartMarginForCharts;
		}
		if (this.calcProp.chartGutter._bottom < 0) {
			this.calcProp.chartGutter._bottom = standartMarginForCharts;
		}

		if ((this.calcProp.chartGutter._left + this.calcProp.chartGutter._right) >
			this.calcProp.widthCanvas) {
			this.calcProp.chartGutter._left = standartMarginForCharts;
		}
		if (this.calcProp.chartGutter._right > this.calcProp.widthCanvas) {
			this.calcProp.chartGutter._right = standartMarginForCharts;
		}

		if ((this.calcProp.chartGutter._top + this.calcProp.chartGutter._bottom) >
			this.calcProp.heightCanvas) {
			this.calcProp.chartGutter._top = 0;
		}
		if (this.calcProp.chartGutter._bottom > this.calcProp.heightCanvas) {
			this.calcProp.chartGutter._bottom = 0;
		}
	},

	//включаю новую функцию. если будут проблемы с отступами - использовать функцию _calculateMarginOnPoints2
	_calculateMarginOnPoints: function (chartSpace) {
		var calculateLeft = 0, calculateRight = 0, calculateTop = 0, calculateBottom = 0, diffPoints, curBetween;
		var pxToMM = this.calcProp.pxToMM;

		var horizontalAxes = this._getHorizontalAxes(chartSpace);
		var verticalAxes = this._getVerticalAxes(chartSpace);
		var horizontalAxis = horizontalAxes ? horizontalAxes[0] : null;
		var verticalAxis = verticalAxes ? verticalAxes[0] : null;
		var crossBetween = null;
		const plotArea = chartSpace && chartSpace.chart ? chartSpace.chart.plotArea : null;
		const seria = plotArea && plotArea.plotAreaRegion ? plotArea.plotAreaRegion.series[0] : null;
		const type = seria ? seria.layoutId : null;
		const isChartEx = AscFormat.isRealNumber(type) && chartSpace ? chartSpace.isChartEx() : false;
		if(verticalAxis instanceof AscFormat.CValAx) {
			crossBetween = verticalAxis.crossBetween;
		} else if(horizontalAxis instanceof AscFormat.CValAx) {
			crossBetween = horizontalAxis.crossBetween;
		}
		if (horizontalAxis && horizontalAxis.xPoints && horizontalAxis.xPoints.length && this.calcProp.widthCanvas != undefined) {
			if (horizontalAxis instanceof AscFormat.CValAx) {
				if (!horizontalAxis.isReversed()) {
					calculateLeft = horizontalAxis.xPoints[0].pos;
					calculateRight = this.calcProp.widthCanvas / pxToMM - horizontalAxis.xPoints[horizontalAxis.xPoints.length - 1].pos;
				} else {
					calculateLeft = horizontalAxis.xPoints[horizontalAxis.xPoints.length - 1].pos;
					calculateRight = this.calcProp.widthCanvas / pxToMM - horizontalAxis.xPoints[0].pos;
				}
			} else if (((horizontalAxis instanceof AscFormat.CCatAx || horizontalAxis instanceof AscFormat.CDateAx) && verticalAxis && !isNaN(verticalAxis.posX)) || isChartEx) {
				diffPoints = horizontalAxis.xPoints[1] ? Math.abs(horizontalAxis.xPoints[1].pos - horizontalAxis.xPoints[0].pos) : Math.abs(horizontalAxis.xPoints[0].pos - verticalAxis.posX) * 2;

				curBetween = (crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN || isChartEx) ? diffPoints / 2 :  0;
				// for reversed cases, start and end indexes are reversed!
				const isReversed = horizontalAxis.isReversed();
				const startIndex = (!isReversed || isChartEx) ? 0 : horizontalAxis.xPoints.length - 1;
				const endIndex = (!isReversed || isChartEx) ? horizontalAxis.xPoints.length - 1 : 0;
				calculateLeft = horizontalAxis.xPoints[startIndex].pos - curBetween;
				calculateRight = this.calcProp.widthCanvas / pxToMM - (horizontalAxis.xPoints[endIndex].pos + curBetween);
			}
		}

		if (verticalAxis && verticalAxis.yPoints && verticalAxis.yPoints.length && this.calcProp.heightCanvas != undefined) {
			if (verticalAxis instanceof AscFormat.CValAx || isChartEx) {
				if (!verticalAxis.isReversed() || isChartEx) {
					if (isChartEx && type === AscFormat.SERIES_LAYOUT_FUNNEL && verticalAxis.yPoints.length === 1) {
						// for diagrams with single axis and single data point
						let oSize = chartSpace ? chartSpace.getChartSizes() : {startY : 0, h : 0};
						calculateTop = oSize.startY;
						calculateBottom = this.calcProp.heightCanvas / pxToMM - (oSize.h + oSize.startY);
					} else {
						calculateTop = verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos;
						calculateBottom = this.calcProp.heightCanvas / pxToMM - verticalAxis.yPoints[0].pos;
					}
				} else {
					calculateTop = verticalAxis.yPoints[0].pos;
					calculateBottom = this.calcProp.heightCanvas / pxToMM - verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos;
				}
			} else if ((verticalAxis instanceof AscFormat.CCatAx || verticalAxis instanceof AscFormat.CDateAx) && horizontalAxis && !isNaN(horizontalAxis.posY)) {

				diffPoints = verticalAxis.yPoints[1] ? Math.abs(verticalAxis.yPoints[1].pos - verticalAxis.yPoints[0].pos) : Math.abs(verticalAxis.yPoints[0].pos - horizontalAxis.posY) * 2;

				curBetween = 0;
				if (!verticalAxis.isReversed()) {
					if (crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
						curBetween = diffPoints / 2;
					}

					calculateTop = verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos - curBetween;
					calculateBottom = this.calcProp.heightCanvas / pxToMM - (verticalAxis.yPoints[0].pos + curBetween);
				} else {
					if (crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
						curBetween = diffPoints / 2;
					}

					calculateTop = verticalAxis.yPoints[0].pos - curBetween;
					calculateBottom = this.calcProp.heightCanvas / pxToMM - (verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos + curBetween);
				}
			}
		}

		// for funnel chart
		if (!horizontalAxis && isChartEx && verticalAxis && verticalAxis.labels && verticalAxis.yPoints) {
			calculateLeft = this.cChartSpace && this.cChartSpace.plotAreaRect ? this.cChartSpace.plotAreaRect.x : 0;
			let curBetween = verticalAxis.yPoints[1] ? Math.abs(verticalAxis.yPoints[1].pos - verticalAxis.yPoints[0].pos) / 2 : 0;
			// curBetween = !curBetween && this.cChartSpace && this.cChartSpace.plotAreaRect ? verticalAxis.yPoints[0].pos - this.cChartSpace.plotAreaRect.y : curBetween;
			calculateTop -= curBetween;
			calculateBottom -= curBetween;
		}

		return {calculateLeft: calculateLeft, calculateRight : calculateRight, calculateTop: calculateTop, calculateBottom: calculateBottom};
	},
	
	_getStandartMargin: function(labelsMargin, keyMargin, textMargin, topMainTitleMargin)
	{
		var defMargin = standartMarginForCharts / this.calcProp.pxToMM;
		var result;

		if (labelsMargin === 0 && keyMargin === 0 && textMargin === 0 && topMainTitleMargin === 0) {
			result = defMargin;
		} else if (labelsMargin !== 0 && keyMargin === 0 && textMargin === 0 && topMainTitleMargin === 0) {
			result = defMargin / 2;
		} else if (labelsMargin !== 0 && keyMargin === 0 && textMargin !== 0 && topMainTitleMargin === 0) {
			result = defMargin;
		} else if (labelsMargin !== 0 && keyMargin !== 0 && textMargin !== 0 && topMainTitleMargin === 0) {
			result = defMargin + defMargin / 2;
		} else if (labelsMargin === 0 && keyMargin !== 0 && textMargin === 0 && topMainTitleMargin === 0) {
			result = defMargin;
		} else if (labelsMargin === 0 && keyMargin === 0 && textMargin !== 0 && topMainTitleMargin === 0) {
			result = defMargin;
		} else if (labelsMargin === 0 && keyMargin !== 0 && textMargin !== 0 && topMainTitleMargin === 0) {
			result = defMargin + defMargin / 2;
		} else if (labelsMargin !== 0 && keyMargin !== 0 && textMargin === 0 && topMainTitleMargin === 0) {
			result = defMargin;
		} else if (labelsMargin === 0 && keyMargin !== 0 && textMargin !== 0 && topMainTitleMargin === 0) {
			result = defMargin + defMargin / 2;
		} else if (labelsMargin === 0 && keyMargin === 0 && topMainTitleMargin !== 0) {
			result = defMargin + defMargin / 2;
		} else if (labelsMargin === 0 && keyMargin !== 0 && topMainTitleMargin !== 0) {
			result = 2 * defMargin;
		} else if (labelsMargin !== 0 && keyMargin === 0 && topMainTitleMargin !== 0) {
			result = defMargin;
		} else if (labelsMargin !== 0 && keyMargin !== 0 && topMainTitleMargin !== 0) {
			result = 2 * defMargin;
		}

		return result;
	},

	//TEST and change _calculateMarginLabels->_calculateMarginLabels2
	_calculateMarginLabels: function (chartSpace) {
		let left = 0, right = 0, bottom = 0, top = 0;
		let leftDownPointX, leftDownPointY, rightUpPointX, rightUpPointY;

		let crossBetween = chartSpace.getValAxisCrossType();
		let horizontalAxes = this._getHorizontalAxes(chartSpace);
		let verticalAxes = this._getVerticalAxes(chartSpace);
		let horizontalAxis = horizontalAxes ? horizontalAxes[0] : null;
		let verticalAxis = verticalAxes ? verticalAxes[0] : null;

		let getInterval = function (_axis) {
			if (_axis.interval) {
				return Math.abs((horizontalAxis.interval) / 2);
			}
			if (_axis.scale && _axis.scale.length > 1) {
				return Math.abs(_axis.scale[1] - _axis.scale[0]);
			}
			return 1;
		};

		let diffPoints;
		if(horizontalAxis && horizontalAxis.xPoints &&  horizontalAxis.xPoints.length) {
			let orientationHorAxis = !horizontalAxis.isReversed();
			diffPoints = 0;
			if((horizontalAxis instanceof AscFormat.CDateAx || horizontalAxis instanceof AscFormat.CCatAx) && crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
				diffPoints = getInterval(horizontalAxis);
			}
			if(orientationHorAxis) {
				leftDownPointX = horizontalAxis.xPoints[0].pos - diffPoints;
				rightUpPointX = horizontalAxis.xPoints[horizontalAxis.xPoints.length - 1].pos + diffPoints;
			} else {
				leftDownPointX = horizontalAxis.xPoints[horizontalAxis.xPoints.length - 1].pos - diffPoints;
				rightUpPointX = horizontalAxis.xPoints[0].pos + diffPoints;
			}

			if (verticalAxis && verticalAxis.labels && !verticalAxis.bDelete) {
				//подпись оси OY находится левее крайней левой точки
				if (leftDownPointX >= verticalAxis.labels.x) {
					left = leftDownPointX - verticalAxis.labels.x;
				} else if ((verticalAxis.labels.x + verticalAxis.labels.extX) >= rightUpPointX)//правее крайней правой точки
				{
					right = verticalAxis.labels.x + verticalAxis.labels.extX - rightUpPointX;
				}
			}
		}

		if(verticalAxis && verticalAxis.yPoints && verticalAxis.yPoints.length) {
			let orientationVerAxis = !verticalAxis.isReversed();
			diffPoints = 0;
			const isChartEx = chartSpace.isChartEx();
			if((verticalAxis instanceof AscFormat.CDateAx || verticalAxis instanceof AscFormat.CCatAx)&& crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
				diffPoints = getInterval(verticalAxis);
			}
			if(orientationVerAxis || isChartEx) {
				leftDownPointY = verticalAxis.yPoints[0].pos + diffPoints;
				rightUpPointY = verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos - diffPoints;
			} else {
				leftDownPointY = verticalAxis.yPoints[verticalAxis.yPoints.length - 1].pos + diffPoints;
				rightUpPointY = verticalAxis.yPoints[0].pos - diffPoints;
			}

			if (horizontalAxis && horizontalAxis.labels && !horizontalAxis.bDelete) {
				//подпись оси OX находится ниже крайней нижней точки
				if ((horizontalAxis.labels.y + horizontalAxis.labels.extY) >= leftDownPointY) {
					bottom = (horizontalAxis.labels.y + horizontalAxis.labels.extY) - leftDownPointY;
				} else if (horizontalAxis.labels.y <= rightUpPointY)//выше верхней
				{
					top = rightUpPointY - horizontalAxis.labels.y;
				}
			}
		}

		return {left: left, right: right, top: top, bottom: bottom};
	},

	_getHorizontalAxes: function(chartSpace) {
		var res = null;

		var axId = chartSpace.chart.plotArea.axId;
		for(var i = 0; i < axId.length; i++) {
			var axis = chartSpace.chart.plotArea.axId[i];
			if(axis.axPos === window['AscFormat'].AX_POS_B || axis.axPos === window['AscFormat'].AX_POS_T) {
				if(!res) {
					res = [];
				}
				res.push(axis);
			}
		}

		return res;
	},

	_getVerticalAxes: function(chartSpace) {
		var res = [];
		var priorityAxis = [];

		let charts = chartSpace.chart.plotArea.charts;
		const isChartEx = chartSpace.isChartEx();
		if (isChartEx) {
			var axId = chartSpace.chart.plotArea.axId;
			for(var i = 0; i < axId.length; i++) {
				var axis = chartSpace.chart.plotArea.axId[i];
				if(axis.axPos === window['AscFormat'].AX_POS_R || axis.axPos === window['AscFormat'].AX_POS_L) {
					if(!res) {
						res = [];
					}
					res.push(axis);
				}
			}
			return res;
		} else {
			for (let i = 0; i < charts.length; i++) {
				for (let j = 0; j < charts[i].axId.length; j++) {
					var axis = charts[i].axId[j];
					if(axis.axPos === window['AscFormat'].AX_POS_R || axis.axPos === window['AscFormat'].AX_POS_L) {
						if (charts[i].getObjectType() === AscDFH.historyitem_type_RadarChart) {
							priorityAxis.push(axis);
						} else {
							res.push(axis);
						}
					}
				}
			}
			return res.length || priorityAxis.length ? priorityAxis.concat(res) : null;
		}

	},

	isRadarChart: function(axis) {
		let charts = this.cChartSpace && this.cChartSpace.chart && this.cChartSpace.chart.plotArea && this.cChartSpace.chart.plotArea.charts;
		if (charts) {
			for (let i = 0; i < charts.length; i++) {
				for (let j = 0; j < charts[i].axId.length; j++) {
					if (charts[i].getObjectType() === AscDFH.historyitem_type_RadarChart && axis === charts[i].axId[j]) {
						return true;
					}
				}
			}
		}
		return false;
	},

	//****calculate properties****
	_calculateProperties: function (chartSpace) {
		if (!this.calcProp.scale) {
			this.preCalculateData(chartSpace, true);
		}

		//считаем маргины
		this._calculateMarginsChart(chartSpace);

		// check if diagmar size affected by layout
		const isLayoutSizes = this.cChartSpace.isLayoutSizes();

		// layour should not affect for circular charts
		let isCircular = false;
		if (isLayoutSizes && chartSpace && chartSpace.chart && chartSpace.chart.plotArea && Array.isArray(chartSpace.chart.plotArea.charts)) {
			const charts = chartSpace.chart.plotArea.charts;
			for (let i = 0; i < charts.length; i++) {
				const typeChart = charts[0].getObjectType();
				isCircular = !isCircular ? (typeChart === AscDFH.historyitem_type_PieChart || typeChart === AscDFH.historyitem_type_DoughnutChart) : isCircular;
			}
		}
		
		if (isLayoutSizes && !isCircular) {
			this.calcProp.trueWidth = this.cChartSpace.chart.plotArea.extX * this.calcProp.pxToMM;
			this.calcProp.trueHeight = this.cChartSpace.chart.plotArea.extY * this.calcProp.pxToMM;
			this.calcProp.chartGutter._top = this.cChartSpace.chart.plotArea.y * this.calcProp.pxToMM;
		} else {
			this.calcProp.trueWidth = this.calcProp.widthCanvas - this.calcProp.chartGutter._left - this.calcProp.chartGutter._right;
			this.calcProp.trueHeight = this.calcProp.heightCanvas - this.calcProp.chartGutter._top - this.calcProp.chartGutter._bottom;
		}
	},
	
	//****new calculate data****
	_calculateStackedData2: function (data, chart) {
		if(!data) {
			return;
		}

		let maxMinObj;
		let grouping = this.getChartGrouping(chart);
		let chartType = this._getChartType(chart);
		let t = this;

		//TODO стоит сделать общую обработку для всех диаграмм
		let calculateStacked = function(sum) {
			if(c_oChartTypes.HBar === chartType || c_oChartTypes.Bar === chartType) {
				let originalData = $.extend(true, [], data);
				for (let j = 0; j < data.length; j++) {
					for (let i = 0; i < data[j].length; i++) {
						data[j][i] = t._findPrevValue(originalData, j, i)
					}
				}

				if(sum) {
					for (let j = 0; j < data.length; j++) {
						for (let i = 0; i < data[j].length; i++) {
							if (sum[j] === 0) {
								break;
							}
							data[j][i] = data[j][i] / sum[j];
						}
					}
				}
			} else {
				for (let j = 0; j < (data.length - 1); j++) {
					for (let i = 0; i < data[j].length; i++) {
						if (!data[j + 1]) {
							data[j + 1] = [];
						}
						data[j + 1][i] = data[j + 1][i] + data[j][i];
					}
				}

				if(sum && data[0]) {
					for (let j = 0; j < (data[0].length); j++) {
						for (let i = 0; i < data.length; i++) {
							if (sum[j] == 0) {
								data[i][j] = 0;
							} else {
								data[i][j] = (data[i][j]) / (sum[j]);
							}
						}
					}
				}
			}
		};

		let calculateSum = function() {
			//вычисляем сумму
			//для разных диаграмм она вычисляется по-разному
			let res = [];
			if(c_oChartTypes.HBar === chartType || c_oChartTypes.Bar === chartType) {
				for (let j = 0; j < (data.length); j++) {
					res[j] = 0;
					for (let i = 0; i < data[j].length; i++) {
						if (data[j][i] != null) {
							res[j] += Math.abs(data[j][i]);
						}
					}
				}
			} else {
				if(data[0]) {
					for (let j = 0; j < (data[0].length); j++) {
						res[j] = 0;
						for (let i = 0; i < data.length; i++) {
							if (data[i][j] != null) {
								res[j] += Math.abs(data[i][j]);
							}
						}
					}
				}
			}
			return res;
		};

		if (grouping === 'stacked') {
			calculateStacked();
		} else if (grouping === 'stackedPer') {
			calculateStacked(calculateSum());
		}

		maxMinObj = this._getMaxMinValueArray(data);
		return maxMinObj;
	},

	_calculateExtremumAllCharts: function (chartSpace, isFirstChart) {
		//возвращает массив, где первый элемент - для основной оси, второй - для вспомогательной
		//максимальные/минимальные значения среди всех графиков
		var t = this, isStackedType, isCombinationChartTypes;
		var plotArea = chartSpace.chart.plotArea;
		var charts = plotArea.charts;
		if(!charts || isFirstChart) {
			charts = [plotArea.chart];
		}
		var isScatterType = false; //входит в условие расчета шага

		var getMinMaxCurCharts = function(axisCharts, axis) {

			//предварительно проходимся по всем диаграммам и ищем 100% stacked тип
			isStackedType = false;
			for (var i = 0; i < axisCharts.length; i++) {
				grouping = t.getChartGrouping(axisCharts[i]);
				if (AscDFH.historyitem_type_ScatterChart === axisCharts[i].getObjectType()) {
					isScatterType = true;
				}
				if ("stackedPer" === grouping && !isStackedType) {
					isStackedType = true;
					//break; // пока что заменяем на !isStackedType (т.к. другие типы могут быть нужны для расчета шага)
				}
			}

			var minMaxData, min, max, chart, grouping;
			for (var i = 0; i < axisCharts.length; i++) {
				chart = axisCharts[i];
				grouping = t.getChartGrouping(chart);
				minMaxData = t._calculateData2(chart, grouping, axis);

				//условие добавлено для случая когда комбинированная диаграмма содержит stacked и stackedPer
				if (grouping === "stacked" && isStackedType) {
					isCombinationChartTypes = true;
				}
				/*if("stackedPer" !== grouping && isStackedType) {
					minMaxData.min = minMaxData.min*100;
					minMaxData.max = minMaxData.max*100;
					minMaxData.ymin = minMaxData.ymin*100;
					minMaxData.ymax = minMaxData.ymax*100;
				}*/

				if(AscDFH.historyitem_type_ScatterChart === chart.getObjectType()) {
					if(axis.axPos === window['AscFormat'].AX_POS_B || axis.axPos === window['AscFormat'].AX_POS_T) {
						if(i === 0 || minMaxData.min < min) {
							min = minMaxData.min;
						}
						if(i === 0 || minMaxData.max > max) {
							max = minMaxData.max;
						}
					} else {
						if(i === 0 || minMaxData.ymin < min) {
							min = minMaxData.ymin;
						}
						if(i === 0 || minMaxData.ymax > max) {
							max = minMaxData.ymax;
						}
					}
				} else {
					if(i === 0 || minMaxData.min < min) {
						min = minMaxData.min;
					}
					if(i === 0 || minMaxData.max > max) {
						max = minMaxData.max;
					}
				}
			}
			return {min: min, max: max};
		};

		for(var j = 0; j < plotArea.axId.length; j++) {
			var axObj = plotArea.axId[j];
			var axisCharts = this._getChartsByAxisId(charts, axObj.axId);
			var minMaxAxis = getMinMaxCurCharts(axisCharts, axObj);
			axObj.min = minMaxAxis.min;
			axObj.max = minMaxAxis.max;
		}

		let boundaries = {};

		// add Trendline coordinates and precalculate all the necessary results
		if (chartSpace && chartSpace.chart && chartSpace.chart.plotArea && chartSpace.chart.plotArea.charts) {
			const charts = chartSpace.chart.plotArea.charts;
			const dispBlanksAs = chartSpace.chart.dispBlanksAs;
			for (let i = 0; i < charts.length; i++) {
				if (charts[i].series) {
					const subType = this.getChartGrouping(charts[i]);
					const series = charts[i].series;
					if (subType === 'normal') {
						for (let j = 0; j < series.length; j++) {
							const seria = series[j];
							if (seria.trendline && this.nDimensionCount !== 3) {
								const val = seria.val ? seria.val : seria.yVal;
								const valNumCache = this.getNumCache(val);
								const valPts = valNumCache ? valNumCache.pts : null;
								if (!valPts || !valPts.length || seria.isHidden === true || valPts.length < 2) {
									continue;
								}

								// if xVal is given
								const catNumCache = seria.xVal ? this.getNumCache(seria.xVal) : null;
								const catPts = catNumCache ? catNumCache.pts : null;

								// obtain reference of targetPts either catpts or valpts;
								const targetPts = catPts ? catPts : valPts;
								let valIterator = 0;
								let catIterator = 0;

								const decideCatValue = function (index) {
									++catIterator;
									return catPts ? targetPts[index].val : targetPts[index].idx + 1;
								}

								const decideValValue = function () {
									return valPts[valIterator++].val
								}

								if (dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
									const ptCount = catNumCache ? catNumCache.ptCount : valNumCache.ptCount;
									for (let k = 0; k < ptCount; k++) {
										const statement1 = catIterator < targetPts.length;
										const statement2 = valIterator < valPts.length;
										const catVal = statement1 && k === targetPts[catIterator].idx ? decideCatValue(catIterator) : 0;
										const valVal = statement2 && k === valPts[valIterator].idx ? decideValValue() : 0;
										this.trendline.addCoordinate(catVal, valVal, charts[i].Id, seria.Id);
									}
								} else {
									for (let k = 0; k < targetPts.length && valIterator < valPts.length; k++) {
										if ( targetPts[k].idx === valPts[valIterator].idx) {
											const catVal = decideCatValue(k);
											const valVal = decideValValue();
											this.trendline.addCoordinate(catVal, valVal, charts[i].Id, seria.Id);
										} else if (targetPts[k].idx > valPts[valIterator].idx) {
											valIterator++;
											k--;
										}
									}
								}
							}
						}
					}
				}
			}
			if(!chartSpace.bEmptySeries){
				this.trendline.preCalculate(charts, boundaries);
			}
		}

		for(var j = 0; j < plotArea.axId.length; j++) {
			var axObj = plotArea.axId[j];

			if (axObj.scaling.logBase && axObj.min <= 0) {
				axObj.min = 1;
			}

			// check for new boundaries
			if (boundaries[axObj.Id]){
				axObj.min = Math.min(axObj.min, boundaries[axObj.Id].min) ;
				axObj.max = Math.max(axObj.max, boundaries[axObj.Id].max) ;
			}
			
			//если будут проблемы, протестить со старой функцией -> this._getAxisValues(false, minMaxAxis.min, minMaxAxis.max, chartSpace)
			axObj.scale = this._roundValues(this._getAxisValues2(axObj, chartSpace, isStackedType && !isCombinationChartTypes, isScatterType));

			if(isStackedType && !axObj.scaling.logBase && !isCombinationChartTypes) {
				//для случая 100% stacked - если макс/мин равно определенному делению, большие/меньшие - убираем, логарифмическая шкала - исключение
				if(axObj.scale[0] !== 0 && axObj.min === axObj.scale[1]) {
					axObj.scale.splice(0, 1);
				}
				if(axObj.max === axObj.scale[axObj.scale.length - 2]) {
					axObj.scale.splice(axObj.scale.length - 1, 1);
				}
			}
		}
	},

	_getChartsByAxisId: function(charts, id) {
		var res = [];
		for(var i = 0; i < charts.length; i++) {
			if(!charts[i].axId) {
				continue;
			}

			for(var j = 0; j < charts[i].axId.length; j++) {
				if(id === this._searchChangedAxisId(charts[i].axId[j].axId)) {
					res.push(charts[i]);
					break;
				}
			}
		}
		return res;
	},

	_searchChangedAxisId: function(id) {
		var res = id;

		if(this.changeAxisMap && this.changeAxisMap[id]) {
			res = this.changeAxisMap[id];
		}

		return res;
	},

	_searchChangedAxis: function(axis) {
		var res = axis;

		if(this.changeAxisMap && axis.axId && this.changeAxisMap[axis.axId]) {
			var newId = this.changeAxisMap[axis.axId];
			var newAxis = this._searchAxisById(newId);
			if(null !== newAxis) {
				res = newAxis;
			}
		}

		return res;
	},

	_searchAxisById: function(id, chartSpace) {
		var res = null;
		chartSpace = chartSpace || this.cChartSpace;
		if(chartSpace) {
			var axId = chartSpace.chart.plotArea.axId;
			for(var i = 0; i < axId.length; i++) {
				if(id === axId[i].axId) {
					res = axId[i];
					break;
				}
			}
		}

		return res;
	},
	
	_calculateChangeAxisMap: function (chartSpace) {
		//ms рисует по-разному диаграммы со скрытымми/не скрытыми осями
		//если ось скрыта - ищем замену среди основных открытых
		this.changeAxisMap = {};
		var axId = chartSpace.chart.plotArea.axId;

		var searchNeedAxis = function(excludeAxis) {
			var res = null;

			var needPos = excludeAxis.axPos === window['AscFormat'].AX_POS_L || excludeAxis.axPos === window['AscFormat'].AX_POS_R;
			for (var j = 0; j < axId.length; j++) {
				var curAxis = axId[j];
				var curPos = curAxis.axPos === window['AscFormat'].AX_POS_L || curAxis.axPos === window['AscFormat'].AX_POS_R;
				if(excludeAxis.axId !== axId[j].axId && needPos === curPos) {
					if(excludeAxis.getObjectType() === curAxis.getObjectType()) {
						res = curAxis;
						break;
					}
				}
			}

			return res;
		};

		if(axId) {
			for (var i = 0; i < axId.length; i++) {
				if(axId[i].bDelete && AscDFH.historyitem_type_ValAx === axId[i].getObjectType()) {
					var needAxis = searchNeedAxis(axId[i]);
					if(needAxis) {
						this.changeAxisMap[axId[i].axId] = needAxis.axId;
					}
				}
			}
		}
	},

	_calculateData2: function(chart, grouping, axis)
	{
		var xNumCache, yNumCache, newArr, arrValues = [], max = 0, min = 0, minY = 0, maxY = 0;
		var series = chart.series;
		var chartType = this._getChartType(chart);
		var t = this;
		var bFirst = true;
		var maxErr = null, minErr = null;
		const isHorAxis = axis.axPos === window['AscFormat'].AX_POS_B || axis.axPos === window['AscFormat'].AX_POS_T ? true : false;

		var addValues = function(tempValX, tempValY) {
			if(bFirst) {
				min = tempValX;
				max = tempValX;
				minY = tempValY;
				maxY = tempValY;
				bFirst = false;
			}

			if (tempValX < min) {
				min = tempValX;
			}
			if (tempValX > max) {
				max = tempValX;
			}
			if (tempValY < minY) {
				minY = tempValY;
			}
			if (tempValY > maxY) {
				maxY = tempValY;
			}
		};

		const findErrBarIdx = function (_seria, isHorAxis) {
			const errBars = Array.isArray(_seria.errBars) ? _seria.errBars : null;
			if (errBars) {
				for (let i = 0; i < errBars.length; i++) {
					if (isHorAxis && (errBars[i].errDir === AscFormat.st_errdirX || errBars[i].errDir === null)) {
						return i;
					}
					if (!isHorAxis && (errBars[i].errDir === AscFormat.st_errdirY || errBars[i].errDir === null)) {
						return i;
					}
				}
			}
			return null;
		}

		let calcErrValues = function (_seria, _l, _col, isHorAxis) {
			if (!_seria.errBars) {
				return;
			}

			// either horAxis with horErrbars or vertAxis with vertErrBars
			const errBarIdx = findErrBarIdx(_seria, isHorAxis);

			if (errBarIdx === null) {
				return;
			}

			const errBar = _seria.errBars[errBarIdx];
			var oErrVal = t.errBars.calculateErrVal(chart, _l, errBarIdx, _col);
			if (oErrVal) {
				var pointVal;
				if (null !== oErrVal.startVal) {
					pointVal = oErrVal.startVal;
				} else {
					pointVal = oErrVal.val;
				}

				const remainder = oErrVal.plusErrVal % 1;
				var plusErrVal = remainder <= 0.5 ? Math.floor(oErrVal.plusErrVal) : Math.ceil(oErrVal.plusErrVal);
				var minusErrVal = Math.floor(oErrVal.minusErrVal);
				if (!AscFormat.isRealNumber(plusErrVal)) {
					plusErrVal = 0;
				}
				if (!AscFormat.isRealNumber(minusErrVal)) {
					plusErrVal = 0;
				}
				switch (errBar.errBarType) {
					case AscFormat.st_errbartypeBOTH: {
						if (maxErr === null || maxErr < pointVal + plusErrVal) {
							maxErr = pointVal + plusErrVal;
						}
						if (minErr === null || minErr > pointVal - minusErrVal) {
							minErr = pointVal - minusErrVal;
						}
						break;
					}
					case AscFormat.st_errbartypePLUS: {
						if (maxErr === null || maxErr < pointVal + plusErrVal) {
							maxErr = pointVal + plusErrVal;
						}
						break;
					}
					case AscFormat.st_errbartypeMINUS: {
						if (minErr === null || minErr > pointVal - minusErrVal) {
							minErr = pointVal - minusErrVal;
						}
						break;
					}
				}
			}
		};

		var generateArrValues = function (isHorAxis) {

			var seria, numCache, pts;
			if(AscDFH.historyitem_type_ValAx === axis.getObjectType()) {
				var isEn = false, numSeries = 0;

				for (var l = 0; l < series.length; l++) {
					seria = series[l];
					numCache = t.getNumCache(seria.val);
					pts = numCache ? numCache.pts : null;

					if (!pts || !pts.length || seria.isHidden === true) {
						continue;
					}

					var n = 0;
					arrValues[numSeries] = [];
					for (var col = 0; col < numCache.ptCount; col++) {
						var curPoint = numCache.pts[col];

						//условие дбавлено для того, чтобы диаграммы, данные которых имеют мин/макс и пустые ячейки, рисовались грамотно
						if(!curPoint && (t.calcProp.subType === 'stackedPer' || t.calcProp.subType === 'stacked')) {
							curPoint = {val: 0};
						} else if (!curPoint || curPoint.isHidden === true) {
							continue;
						}

						var val = curPoint.val;
						var value = parseFloat(val);

						if (!isEn && !isNaN(value)) {
							min = value;
							max = value;
							isEn = true;
						}

						calcErrValues(seria, l, col, isHorAxis);

						if (!isNaN(value) && value > max) {
							max = value;
						}
						if (!isNaN(value) && value < min) {
							min = value;
						}

						if (isNaN(value) && val == '' && (((chartType === c_oChartTypes.Line ) && grouping === 'normal'))) {
							value = '';
						} else if (isNaN(value)) {
							value = 0;
						}

						if (chartType === c_oChartTypes.Pie || chartType === c_oChartTypes.DoughnutChart) {
							value = Math.abs(value);
						}

						arrValues[numSeries][n] = value;
						n++;
					}
					numSeries++;
				}

				if(min === max) {
					if(min < 0) {
						max = 0;
					} else {
						min = 0;
					}
				}
			} else {
				if(series.length > 0) {
					// every chart except scatter should start from 1
					min = 1;
					max = 1;
					// find max value across each seria
					for (let i = 0; i < series.length; i++) {
						const seria = series[i];
						if (seria) {
							numCache = t.getNumCache(seria.val);
							const ptCount = numCache && AscFormat.isRealNumber(numCache.ptCount) ? numCache.ptCount : 0;
							// trendline can affect max value
							const newMax = seria.trendline && seria.trendline.forward && ptCount > 1 ? ptCount + seria.trendline.forward : ptCount;
							max = Math.max(max, newMax);
						}
					}
				}
			}
		};

		var generateArrValuesScatter = function (isHorAxis) {
			newArr = [];
			for (var l = 0; l < series.length; ++l) {
				newArr[l] = [];

				yNumCache = t.getNumCache(series[l].yVal);
				if(!yNumCache) {
					continue;
				}

				for (var j = 0; j < yNumCache.ptCount; ++j) {
					var val = t._getScatterPointVal(series[l], j);
					calcErrValues(series[l], l, j, isHorAxis);
					if(val) {
						addValues(val.x, val.y);
						newArr[l][j] = [val.x, val.y];
					}
				}

				// check the impact of trendline on scatter chart
				if (series[l].trendline) {
					min = series[l].trendline.backward ? min - series[l].trendline.backward : min;
					max = series[l].trendline.forward ? max + series[l].trendline.forward : max;
				}
			}
		};

		if (chartType !== c_oChartTypes.Scatter) {
			generateArrValues(isHorAxis);
		} else {
			generateArrValuesScatter(isHorAxis);
		}

		if(chartType === c_oChartTypes.Bar || chartType === c_oChartTypes.HBar) {
			arrValues = arrReverse(arrValues);
		}

		//пересчёт данных для накопительных диаграмм
		if (AscDFH.historyitem_type_ValAx === axis.getObjectType() && ("stackedPer" === grouping || "stacked" === grouping)) {
			if (newArr) {
				arrValues = newArr;
			}
			var stackedExtremum = this._calculateStackedData2(arrValues, chart);
			if(stackedExtremum) {
				min = stackedExtremum.min;
				max = stackedExtremum.max;
			}
		}

		if (chartType !== c_oChartTypes.Scatter || isHorAxis) {
			min = (minErr !== null && min > minErr) ? minErr : min;
			max = (maxErr !== null && max < maxErr) ? maxErr : max;
		} else {
			minY = (minErr !== null && minY > minErr) ? minErr : minY;
			maxY = (maxErr !== null && maxY < maxErr) ? maxErr : maxY;
		}

		return {min: min, max: max, ymin: minY, ymax: maxY};
	},

	_getScatterPointVal: function(seria, idx) {
		var yNumCache = this.getNumCache(seria.yVal);

		if (!yNumCache) {
			return null;
		}
		var xNumCache = seria.xVal ? this.getNumCache(seria.xVal) : null;
		var yPoint, xPoint, xVal, yVal;
		var res = null;
		if(yNumCache && xNumCache) {
			yPoint = yNumCache.getPtByIndex(idx);
			xPoint = xNumCache.getPtByIndex(idx);
			if(xPoint) {
				yVal = yPoint ? parseFloat(yPoint.val) : 0;
				xVal = parseFloat(xPoint.val);
				res = {x: xVal, y: yVal};
			}
		} else if(yNumCache) {
			yPoint = yNumCache.getPtByIndex(idx);

			var dispBlanksAs =  this.cChartSpace.chart.dispBlanksAs;
			if(yPoint) {
				yVal = parseFloat(yPoint.val);
			} else if(dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
				yVal = 0;
			} else {
				yVal = null;
			}

			xVal = idx + 1;
			res = {x: xVal, y: yVal, xPoint: xPoint, yPoint: yPoint};

		}
		return res;
	},

	_getAxisValues2: function (axis, chartSpace, isStackedType, isScatter) {
		let isOx = axis.axPos === window['AscFormat'].AX_POS_B || axis.axPos === window['AscFormat'].AX_POS_T;
		//для оси категорий берем интервал 1
		let arrayValues;
		const axisType = axis.getObjectType();
		if(AscDFH.historyitem_type_CatAx === axisType || AscDFH.historyitem_type_DateAx === axisType) {
			arrayValues = [];
			let max = axis.max;
			for(let i = axis.min; i <= max; i++) {
				arrayValues.push(i);
			}
			return arrayValues;
		}

		let yMin = axis.min;
		let yMax = axis.max;
		let logBase = axis.scaling && axis.scaling.logBase;

		let manualMin = axis.scaling && axis.scaling.min !== null ? axis.scaling.min : null;
		let manualMax = axis.scaling && axis.scaling.max !== null ? axis.scaling.max : null;

		if (logBase) {
			yMax = yMax <= 0 ? 10 : yMax;
			yMin = yMin <= 0 ? 1 : yMin;
			arrayValues = this._getLogArray(yMin, yMax, logBase, axis);
			return arrayValues;
		}

		var bIsManualStep = false;
		let t = this;
		let calcAxisMinMax = function (isDefaultMinMax) {
			let trueMinMax = t._getTrueMinMax((manualMin !== null && manualMin > yMin) ? manualMin : yMin, (manualMax !== null && manualMax < yMax) ? manualMax : yMax, isDefaultMinMax, isScatter, manualMax);
			let _axisMin, _axisMax, _step, firstDegree;
			//TODO временная проверка для некорректных минимальных и максимальных значений
			if (manualMax && manualMin && manualMax < manualMin) {
				if (manualMax < 0) {
					manualMax = 0;
				} else {
					manualMin = 0;
				}
			}

			_axisMin = manualMin !== null && manualMin !== undefined ? manualMin : trueMinMax.min;
			_axisMax = manualMax !== null && manualMax !== undefined ? manualMax : trueMinMax.max;

			if (_axisMax < _axisMin) {
				if(_axisMax > 0) {
					manualMax = 2 * _axisMin;
					_axisMax = manualMax;
				} else {
					_axisMin = 2 * _axisMax;
				}
			}

			//приводим к первому порядку
			firstDegree = t._getFirstDegree((Math.abs(_axisMax - _axisMin)) / 10);

			//находим шаг
			if (axis && axis.majorUnit != null) {
				_step = axis.majorUnit;
				bIsManualStep = true;
			} else {
				//было следующее условие - isOx || c_oChartTypes.HBar === this.calcProp.type
				if (isOx /*&& !isScatter && axisMin !== 0 && axisMax !== 0*/) {
					_step = t._getStep(firstDegree.val + (firstDegree.val / 10) * 1.1111);
					// _step = t._getStep(firstDegree.val);
				} else {
					_step = t._getStep(firstDegree.val);
				}

				_step = _step * firstDegree.numPow;
			}
			return {step: _step, axisMin: _axisMin, axisMax: _axisMax};
		};

		//максимальное и минимальное значение(по документации excel)
		let isRadarChart = this.isRadarChart(axis);
		let axisMin, axisMax, step;
		let minMaxCalc = calcAxisMinMax(isStackedType);
		if (isRadarChart) {
			//step get from common calculation
			step = minMaxCalc.step;
			//max/min get from default min/max
			minMaxCalc = calcAxisMinMax(true);
			axisMin = minMaxCalc.axisMin;
			axisMax = minMaxCalc.axisMax;
		} else {
			step = minMaxCalc.step;
			axisMin = minMaxCalc.axisMin;
			axisMax = minMaxCalc.axisMax;
		}
		
		if (isNaN(step) || step === 0) {
			if (AscFormat.isRealNumber(manualMax) && manualMax <= 0) {
				arrayValues = [-1, -0.9, -0.8, -0.7, -0.6, -0.5, -0.4, -0.3, -0.2, -0.1, 0];
			} else {
				arrayValues = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1];
			}
		} else {
			if (false && isRadarChart) {
				arrayValues = this._getRadarAxisValues(axisMin, axisMax, step);
			} else {
				arrayValues = this._getArrayDataValues(step, axisMin, axisMax, manualMin, manualMax, isRadarChart);
			}
		}

		//проверка на переход в другой диапазон из-за ограничения по высоте
		if (!bIsManualStep && !chartSpace.isSparkline) {
			let props = {
				arrayValues: arrayValues,
				step: step,
				axisMin: axisMin,
				axisMax: axisMax,
				manualMin: manualMin,
				manualMax: manualMax
			};
			arrayValues = this._correctDataValuesFromHeight(props, chartSpace);
		}

		//TODO для 3d диаграмм. пересмотреть!
		if (this._isSwitchCurrent3DChart(chartSpace)) {
			if (yMax > 0 && yMin < 0) {
				if (manualMax == null && yMax <= arrayValues[arrayValues.length - 2]) {
					arrayValues.splice(arrayValues.length - 1, 1);
				}
				if (manualMin == null && yMin >= arrayValues[1]) {
					arrayValues.splice(0, 1);
				}
			} else if (yMax > 0 && yMin >= 0) {
				if (manualMax == null && yMax <= arrayValues[arrayValues.length - 2]) {
					arrayValues.splice(arrayValues.length - 1, 1);
				}
			} else if (yMin < 0 && yMax <= 0) {
				if (manualMin == null && yMin >= arrayValues[1]) {
					arrayValues.splice(0, 1);
				}
			}
		}

		return arrayValues;
	},

	_correctDataValuesFromHeight: function (props, chartSpace, isOxAxis) {
		var res = props.arrayValues;
		var heightCanvas = chartSpace.extY * this.calcProp.pxToMM;
		var margins = this._calculateMarginsChart(chartSpace, true);
		var trueHeight = heightCanvas - margins.top - margins.bottom;

		var axisMin = props.axisMin;
		var axisMax = props.axisMax;
		var manualMin = props.manualMin;
		var manualMax = props.manualMax;
		var newStep = props.step;


		if (isOxAxis) {

		} else {
			if (axisMin < 0 && axisMax > 0) {
				/*var limitArr = [0, 0, 0, 28, 20, 20, 18, 20, 18, 18, 17, 16];
				 var limit = limitArr[res.length - 1];
				 var heightGrid = Math.round((trueHeight / (res.length - 1)));
				 while(heightGrid <= limit)
				 {
				 var firstDegreeStep = this._getFirstDegree(newStep);
				 var tempStep = this._getNextStep(firstDegreeStep.val);
				 newStep = tempStep * firstDegreeStep.numPow;
				 res = this._getArrayDataValues(newStep, axisMin, axisMax, manualMin, manualMax);

				 if(res.length <= 3)
				 {
				 break;
				 }

				 limit = limitArr[res.length - 1];
				 heightGrid = Math.round((trueHeight / (res.length - 1)));
				 }*/
			} else {
				var limitArr = [0, 0, 32, 26, 24, 22, 21, 19, 18, 17, 16];
				var limit = limitArr[res.length - 1];
				const num = (trueHeight / (res.length - 1));
				var heightGrid = Math.round(num);
				while (heightGrid <= limit) {
					var firstDegreeStep = this._getFirstDegree(newStep);
					var tempStep = this._getNextStep(firstDegreeStep.val);
					newStep = tempStep * firstDegreeStep.numPow;
					const newRes = this._getArrayDataValues(newStep, axisMin, axisMax, manualMin, manualMax);

					//new array cannot be generated from broken data, example: step is NaN
					if (!newRes) {
						break;
					}

					//after we checked that newRes is working array substitute res
					res = newRes;
					if (res.length <= 2) {
						break;
					}

					limit = limitArr[res.length - 1];
					heightGrid = Math.round((trueHeight / (res.length - 1)));
				}
			}
		}


		return res;
	},

	_getNextStep: function (step) {
		if (step === 1) {
			step = 2;
		} else if (step === 2) {
			step = 5;
		} else if (step === 5) {
			step = 10;
		}

		return step;
	},

	_getArrayDataValues: function (step, axisMin, axisMax, manualMin, manualMax, isRadarChart) {
		if (!AscFormat.isRealNumber(axisMin) || !AscFormat.isRealNumber(axisMax) || !AscFormat.isRealNumber(step)) {
			return null;
		}

		var arrayValues;
		//минимальное значение оси
		//TODO use axisMin
		let minUnit = 0;

		if (AscFormat.isRealNumber(manualMin)) {
			minUnit = manualMin;
		} else {
			if (AscFormat.isRealNumber(manualMax) && manualMax > axisMin && manualMax >= 0) {
				const stepCount = Math.ceil((manualMax - axisMin) / step);
				minUnit = manualMax - (stepCount * step);
				// minUnit = Math.min(minUnit, manualMax - upperLimit);
			} else if (AscFormat.isRealNumber(manualMax)) {
				if (manualMax > 0) {
					minUnit = 0;
				} else if (manualMax === 0) {
					minUnit = -1;
				} else {
					minUnit = manualMax * 2;
				}
			} else {
				// possible problems if axisMin < 0 should be checked
				// const statement1 = !isRadarChart && minUnit >= axisMin;
				// const statement2 = isRadarChart && minUnit > axisMin;
				minUnit = Math.floor(axisMin / step) * step;
			}
		}

		//массив подписей
		arrayValues = this._getArrayAxisValues(minUnit, axisMin, axisMax, step, manualMin, manualMax);

		return arrayValues;
	},

	_getLogArray: function (yMin, yMax, logBase, axis) {	
		// first check dependencies to be numeric and valid
		if((yMin !== 0 && !yMin) || (yMax !== 0 && !yMax)) {
			return [1, logBase];
		}
		logBase = (logBase < 2) ? 2 : logBase;

		/**
		 * axis contains min and max values and logBase
		 * if they are null just use yMin lowerBound and yMax upperBound
		*/
		let trueMin = (axis && axis.scaling && axis.scaling.min) !== null ? _roundValue(axis.scaling.min) : null;
		let trueMax = (axis && axis.scaling && axis.scaling.max) !== null ? _roundValue(axis.scaling.max) : null;
		yMin = (yMin <= 0 || yMin >= 1) ? 1 : yMin;
		yMax = (yMax <= 0) ? logBase : yMax;
		trueMin = (!trueMin || trueMin <= 0) ? yMin : trueMin;
		trueMax = (!trueMax || trueMax <= 0) ? yMax : trueMax;

		//for the cases when min > max ;
		let realTrueMax = Math.max(trueMin, trueMax);
		let realTrueMin = Math.min(trueMin, trueMax);

		// normalize min and max on the scale of logBase;
		// create loop from min to max, multiply each time by logBase
		const trueMaxPower = Math.ceil(Math.log(realTrueMax) / Math.log(logBase));
		const trueMinPower = Math.floor(Math.log(realTrueMin) / Math.log(logBase));
		realTrueMax = Math.pow(logBase, trueMaxPower);
		realTrueMin = Math.pow(logBase, trueMinPower);
		const result = [];
		while (realTrueMin != 0 && realTrueMin <= realTrueMax) {
			result.push(realTrueMin);
			realTrueMin *= logBase;
		}

		// fo the cases where last number was a little bit greater than target
		// 105000000001 <= 10500000000
		const prevVal = realTrueMin / logBase;
		if (Math.abs(realTrueMax - prevVal) > Math.abs(realTrueMax - realTrueMin)){
			result.push(realTrueMin)
		}

		if (result.length === 1) {
			result.push(result[0] * logBase);
		}
		return result;

		// var result = [];

		// var temp;
		// var pow = 0;
		// var tempPow = yMin;

		// var kF = 1000000000;
		// var manualMin = axis.scaling && axis.scaling.min !== null ? Math.round(axis.scaling.min * kF) / kF : null;
		// var manualMax = axis.scaling && axis.scaling.max !== null ? Math.round(axis.scaling.max * kF) / kF : null;

		// if(manualMin !== null) {
		// 	yMin = manualMin;
		// }

		// if (yMin < 1 && yMin > 0) {
		// 	temp = this._getFirstDegree(yMin).numPow;

		// 	tempPow = temp;
		// 	while (tempPow < 1) {
		// 		pow--;
		// 		tempPow = tempPow * 10;
		// 	}
		// } else {
		// 	temp = Math.pow(logBase, 0);
		// }

		// if (logBase < 1) {
		// 	logBase = 2;
		// }

		// var step = 0;
		// var lMax = 1;
		// if (yMin < 1 && yMin > 0) {
		// 	if (lMax < yMax) {
		// 		lMax = yMax;
		// 	}
		// 	if(manualMax !== null && manualMax > lMax) {
		// 		lMax = manualMax;
		// 	}

		// 	while (temp < lMax) {
		// 		temp = Math.pow(logBase, pow);
		// 		if(manualMin !== null && manualMin > temp) {
		// 			pow++;
		// 			continue;
		// 		}
		// 		if(manualMax !== null && manualMax < temp) {
		// 			result[step] = manualMax;
		// 			break;
		// 		}
		// 		result[step] = temp;
		// 		pow++;
		// 		step++;
		// 	}
		// } else {
		// 	if(manualMax !== null && manualMax > yMax) {
		// 		yMax = manualMax;
		// 	}

		// 	while (temp <= yMax) {
		// 		temp = Math.pow(logBase, pow);
		// 		if(manualMin !== null && manualMin > temp) {
		// 			pow++;
		// 			continue;
		// 		}
		// 		if(manualMax !== null && manualMax < temp) {
		// 			result[step] = manualMax;
		// 			break;
		// 		}
		// 		result[step] = temp;
		// 		pow++;
		// 		step++;
		// 	}
		// }

		// return result;
	},

	_getRadarAxisValues: function (axisMin, axisMax, step) {
		var arrayValues = [];

		var maxPointsCount = 40;
		for (var i = 0; i < maxPointsCount; i++) {
			arrayValues[i] = axisMin + step * i;
			if (arrayValues[i] >= axisMax) {
				break;
			}
		}

		if (!arrayValues.length) {
			arrayValues = [0.2, 0.4, 0.6, 0.8, 1, 1.2];
		}

		return arrayValues;
	},

	_getArrayAxisValues: function (minUnit, axisMin, axisMax, step, manualMin, manualMax) {
		var arrayValues = [];
		var stackedPerMax = null !== manualMax ? manualMax : 100;

		if (this.calcProp.subType === 'stackedPer' && step > stackedPerMax) {
			stackedPerMax = step;
		}

		var maxPointsCount = 600;
		for (var i = 0; i < maxPointsCount; i++) {
			if (this.calcProp.subType === 'stackedPer' && (minUnit + step * i) > stackedPerMax) {
				break;
			}

			arrayValues[i] = minUnit + step * i;

			if (axisMax === 0 && axisMin < 0 && arrayValues[i] === axisMax || (this.calcProp.type === c_oChartTypes.Radar && arrayValues[i] === axisMax)) {
				break;
			} else if ((manualMax != null && arrayValues[i] >= axisMax) || (manualMax == null && arrayValues[i] > axisMax)) {
				if (this.calcProp.subType === 'stackedPer') {
					arrayValues[i] = arrayValues[i];
				}

				break;
			} else if (this.calcProp.subType === 'stackedPer') {
				arrayValues[i] = arrayValues[i];
			}
		}

		/*if(this.calcProp.subType == 'stackedPer')
		 {
		 //TODO пересмотреть все ситуации, когда заданы фиксированные максимальные и минимальные значение выше 100%
		 if(step > axisMax)
		 arrayValues = [axisMin, axisMax];
		 }*/

		if (!arrayValues.length) {
			arrayValues = [0.2, 0.4, 0.6, 0.8, 1, 1.2];
		}

		return arrayValues;
	},

	_getStep: function (step) {
		if (step > 1 && step <= 2) {
			step = 2;
		} else if (step > 2 && step <= 5) {
			step = 5;
		} else if (step > 5 &&
			step <= 10) {
			step = 10;
		} else if (step > 10 && step <= 20) {
			step = 20;
		}

		return step;
	},

	_getTrueMinMax: function (yMin, yMax, isStackedType, isScatter, isMaxSet) {

		var axisMax, axisMin, diffMaxMin;
		var cDiff = 1/6;
		//добавил правку в первую ветку: если минмальное значение оказывается равно 0, то максимальное высчитываем
		//с учётом минимального, равного 0.
		// TODO пересмотреть все остальные ситуации!
		if (yMin >= 0 && yMax >= 0) {
			diffMaxMin = (yMax - yMin) / yMax;
			if (isMaxSet === null && cDiff > diffMaxMin) {
				if (isScatter) {
					axisMin = isStackedType ? yMin : yMin - 0.05 * (yMax - yMin);
					axisMax = isStackedType ? yMax : yMax + 0.05 * (yMax - yMin);
				} else {
					axisMin = yMin - ((yMax - yMin) / 2);
					axisMax = isStackedType ? yMax : yMax + 0.05 * (yMax - yMin);
				}
			} else {
				axisMin = 0;
				axisMax = isStackedType ? yMax : yMax + 0.05 * ( yMax - 0 );
			}
		} else if (yMin <= 0 && yMax <= 0) {
			diffMaxMin = (yMin - yMax) / yMin;
			axisMin = isStackedType ? yMin : yMin + 0.05 * (yMin - yMax);

			if (cDiff < diffMaxMin) {
				axisMax = 0;
			} else {
				axisMax = yMax - ((yMin - yMax) / 2)
			}
		} else if (yMax > 0 && yMin < 0) {
			axisMax = isStackedType ? yMax : yMax + 0.05 * (yMax - yMin);
			axisMin = isStackedType ? yMin : yMin + 0.05 * (yMin - yMax);
		}

		if (axisMin === axisMax) {
			if (axisMin < 0) {
				axisMax = 0;
			} else {
				axisMin = 0;
			}
		}

		return {min: axisMin, max: axisMax};
	},


	//****functions for UP Functions****
	preCalculateData: function (chartSpace, notCalcExtremum) {
		this._calculateChangeAxisMap(chartSpace);
		this.cChartSpace = chartSpace;
		this.calcProp.pxToMM = 1 / AscCommon.g_dKoef_pix_to_mm;

		this.calcProp.pathH = 1000000000;
		this.calcProp.pathW = 1000000000;

		this.calcProp.xaxispos = null;
		this.calcProp.yaxispos = null;

		const isChartEx = chartSpace.isChartEx();

		if (!isChartEx) {
			this.calcProp.type = this._getChartType(chartSpace.chart.plotArea.chart);
			this.calcProp.subType = this.getChartGrouping(chartSpace.chart.plotArea.chart);
		}

		if (!notCalcExtremum) {
			if (!isChartEx) {
				//calculate calcProp -> /min/max/ymax/ymin/
				this._calculateExtremumAllCharts(chartSpace);
			} else {
				this._prepChartExData(chartSpace)
			}
		}

		this.calcProp.widthCanvas = chartSpace.extX * this.calcProp.pxToMM;
		this.calcProp.heightCanvas = chartSpace.extY * this.calcProp.pxToMM;
	},

	_handleUserTypedChartExValMinMax: function (axis) {
		if (!axis) {
			return;
		}
		//check for user typed max and min properties
		let trueMin = (axis.scaling && axis.scaling.min != null) ? _roundValue(axis.scaling.min) : null;
		let trueMax = (axis.scaling && axis.scaling.max != null) ? _roundValue(axis.scaling.max) : null;
		const isTrueMin = AscFormat.isRealNumber(trueMin);
		const isTreuMax = AscFormat.isRealNumber(trueMax);

		if (isTrueMin && isTreuMax && trueMin >= trueMax) {
			axis.max = 0;
			axis.min = 0;
		} else {
			trueMin = AscFormat.isRealNumber(trueMin) ? trueMin : axis.min;
			trueMax = AscFormat.isRealNumber(trueMax) ? trueMax : axis.max;

			//for special cases, when min greater than max, or max is negative
			trueMax = (trueMax === 0 && trueMin >= 0) ? 1 : trueMax;
			trueMax = (trueMin > 1 && trueMin > trueMax - 1) ? trueMax * 2 : trueMax;
			trueMin = (trueMax < 0 && trueMax < trueMin) ? trueMax * 2 : trueMin;

			axis.max = trueMax;
			axis.min = trueMin;
		}
	},

	_chartExHandleAxesConfigurations: function (axes, axisProperties) {
		if (!axes || !Array.isArray(axes) || !axisProperties ) {
			return;
		}

		for (let i = 0; i < axes.length; i++) {
			const customAxis = axes[i].axPos === window['AscFormat'].AX_POS_T || axes[i].axPos === window['AscFormat'].AX_POS_B ? axisProperties.cat : axisProperties.val;
			axes[i].min = customAxis.min;
			axes[i].max = customAxis.max;
			this._handleUserTypedChartExValMinMax(axes[i]);
			axes[i].scale = customAxis.scale.length > 0 ? customAxis.scale : this._roundValues(this._getAxisValues2(axes[i], this.cChartSpace, false, false));
		}
	},

	_prepChartExData: function (chartSpace) {
		// plotArea: CPlotArea
		// data: CNumericPoint
		if (!chartSpace || !chartSpace.chart || !chartSpace.chart.plotArea || !chartSpace.chart.plotArea.plotAreaRegion || !chartSpace.chart.plotArea.plotAreaRegion.series) {
			return;
		}

		const plotArea = chartSpace.chart.plotArea;
		const seria = plotArea.plotAreaRegion.series[0];
		if (!seria) {
			return;
		}
		const type = seria.layoutId;
		const secondType = 1 < plotArea.plotAreaRegion.series.length ? plotArea.plotAreaRegion.series[1].layoutId : null;
		const strLit = seria.getCatLit(type);
		const numLit = seria.getValLit();

		if (numLit && numLit.pts && numLit.ptCount > 0) {

			const axisProperties = {
				cat : {max: null, min:null, scale : []},
				val : {max: null, min:null, scale : []},
			}

			const calculateExtremums = function (type, axisProperties, numLit) {
				// ClusteredColumn has unique labels. Example: [1, 3], (3, 7], ... etc.
				const numArr = numLit.pts;
				if (type === AscFormat.SERIES_LAYOUT_FUNNEL || !axisProperties || !numArr || axisProperties.mean || axisProperties.catMax || axisProperties.catMin) {
					return;
				}
				const isUniqueLabels = type === AscFormat.SERIES_LAYOUT_CLUSTERED_COLUMN;
				if (isUniqueLabels) {
					for (let i = 0; i < numArr.length; i++) {
						CCachedChartExData.prototype._chartExSetAxisMinAndMax.call(this, axisProperties.cat, numArr[i].val);
					}
				} else {
					axisProperties.cat.min = numArr.length > 0 ? numArr[0].idx + 1 : 0;
					axisProperties.cat.max = numArr.length > 0 ? numArr[numArr.length - 1].idx + 1 : 0;
				}
			}

			calculateExtremums(type, axisProperties, numLit);
			plotArea.plotAreaRegion.setCachedData(this._createCachedData(type, seria, numLit, strLit, axisProperties, secondType));
			this._chartExHandleAxesConfigurations(plotArea.axId, axisProperties);
		}
	},

	_createCachedData: function (type, seria, numLit, strLit, axisProperties, secondType) {
		switch (type) {
			case AscFormat.SERIES_LAYOUT_CLUSTERED_COLUMN:
				return new CCachedClusteredColumn(type, seria, numLit, strLit, axisProperties, secondType);
			case AscFormat.SERIES_LAYOUT_WATERFALL:
				return new CCachedWaterfall(type, seria, numLit, axisProperties);
			case AscFormat.SERIES_LAYOUT_FUNNEL:
				return new CCachedFunnel(type, numLit, axisProperties);
			default:
				return null;
		}
	},

	_getChartType: function (chart) {
		var res;
		var typeChart = chart.getObjectType();
		switch (typeChart) {
			case AscDFH.historyitem_type_LineChart: {
				res = c_oChartTypes.Line;
				break;
			}
			case AscDFH.historyitem_type_BarChart: {
				if (chart.barDir !== AscFormat.BAR_DIR_BAR) {
					res = c_oChartTypes.Bar;
				} else {
					res = c_oChartTypes.HBar;
				}
				break;
			}
			case AscDFH.historyitem_type_PieChart: {
				res = c_oChartTypes.Pie;
				break;
			}
			case AscDFH.historyitem_type_AreaChart: {
				res = c_oChartTypes.Area;
				break;
			}
			case AscDFH.historyitem_type_ScatterChart: {
				res = c_oChartTypes.Scatter;
				break;
			}
			case AscDFH.historyitem_type_StockChart: {
				res = c_oChartTypes.Stock;
				break;
			}
			case AscDFH.historyitem_type_DoughnutChart: {
				res = c_oChartTypes.DoughnutChart;
				break;
			}
			case AscDFH.historyitem_type_RadarChart: {
				res = c_oChartTypes.Radar;
				break;
			}
			case AscDFH.historyitem_type_BubbleChart: {
				res = c_oChartTypes.BubbleChart;
				break;
			}
			case AscDFH.historyitem_type_SurfaceChart: {
				res = c_oChartTypes.Surface;
				break;
			}
		}
		return res;
	},

	getChartGrouping: function(chart) {
		var res = null;

		var grouping = chart.grouping;
		var typeChart = chart.getObjectType();

		if (typeChart ===  AscDFH.historyitem_type_LineChart || typeChart === AscDFH.historyitem_type_AreaChart) {
			res = (grouping === AscFormat.GROUPING_PERCENT_STACKED) ? "stackedPer" : (grouping === AscFormat.GROUPING_STACKED) ? "stacked" : "normal";
		} else if (this.nDimensionCount === 3 && grouping === AscFormat.BAR_GROUPING_STANDARD) {
			res = "standard";
		} else {
			res = (grouping === AscFormat.BAR_GROUPING_PERCENT_STACKED) ? "stackedPer" : (grouping === AscFormat.BAR_GROUPING_STACKED) ? "stacked" : "normal";
		}

		return res;
	},

	calculateSizePlotArea : function(chartSpace, bNotRecalculate)
	{
		if(!bNotRecalculate || undefined === this.calcProp.chartGutter) {
			this._calculateMarginsChart(chartSpace);
		}

		if (!this.calcProp.chartGutter.left && !this.calcProp.chartGutter.right && !this.calcProp.chartGutter._bottom && !this.calcProp.chartGutter.top) {
			this._calculateMarginsChart(chartSpace);
		}

		var widthCanvas = chartSpace.extX;
		var heightCanvas = chartSpace.extY;
		
		var w = widthCanvas - (this.calcProp.chartGutter._left + this.calcProp.chartGutter._right) / this.calcProp.pxToMM;
		var h = heightCanvas - (this.calcProp.chartGutter._top + this.calcProp.chartGutter._bottom) / this.calcProp.pxToMM;

        return {w: w , h: h , startX: this.calcProp.chartGutter._left / this.calcProp.pxToMM, startY: this.calcProp.chartGutter._top / this.calcProp.pxToMM};
	},

	drawPaths: function (paths, series, useNextPoint, bIsYVal, byIdx, ignoreFill) {

		var seria, brush, pen, numCache, point;
		var seriesPaths = paths.series;
		var pointDiff = useNextPoint ? 1 : 0;

		if(!seriesPaths) {
			return;
		}

		var getSerByIdx = function(_idx) {
			for(var k = 0; k < series.length; k++) {
				if(series[k] && series[k].idx === _idx) {
					return series[k];
				}
			}
			return null;
		};

		for (var i = 0; i < seriesPaths.length; i++) {

			if (!seriesPaths[i]) {
				continue;
			}

			if(byIdx) {
				seria = getSerByIdx(i);
			} else {
				seria = series[i];
			}

			if(!seria) {
				continue;
			}

			brush = seria.brush;
			pen = seria.pen;

			for (var j = 0; j < seriesPaths[i].length; j++) {

				if (bIsYVal) {
					numCache = this.getNumCache(seria.yVal);
				} else {
					numCache = this.getNumCache(seria.val);
				}

				if(numCache){
					if(byIdx) {
						point = numCache.getPtByIndex(j + pointDiff);
					} else {
						point = numCache.pts[j + pointDiff];
					}

					if (point && point.pen) {
						pen = point.pen;
					}
					if (point && point.brush) {
						brush = point.brush;
					}
				}
				if (seriesPaths[i][j]) {
					this.drawPath(seriesPaths[i][j], pen, ignoreFill ? null : brush);
				}
			}
		}
	},

	drawPathsPoints: function (paths, series, bIsYVal) {

		var seria, brush, pen, dataSeries, markerBrush, markerPen, numCache;

		if(!paths.series) {
			return;
		}

		for (var i = 0; i < paths.series.length; i++) {
			seria = series[i];
			brush = seria.brush;
			pen = seria.pen;

			if (bIsYVal) {
				numCache = this.getNumCache(seria.yVal);
			} else {
				numCache = this.getNumCache(seria.val);
			}

			dataSeries = paths.series[i];

			if (!dataSeries || !paths.points[i]) {
				continue;
			}

			//draw point
			for (var k = 0; k < paths.points[i].length; k++) {
				var numPoint = numCache ? numCache.getPtByIndex(k) : null;
				if (numPoint) {
					markerBrush = numPoint.compiledMarker ? numPoint.compiledMarker.brush : null;
					markerPen = numPoint.compiledMarker ? numPoint.compiledMarker.pen : null;
				}

				//frame of point
				if (paths.points[i][k] && paths.points[i][k].framePaths) {
					this.drawPath(paths.points[i][k].framePaths, null, markerBrush, false);
				}
				//point
				if (paths.points[i][k]) {
					this.drawPath(paths.points[i][k].path, markerPen, markerBrush, true);
				}
			}
		}
	},

	drawPathsByIdx: function (paths, series, useNextPoint, bIsYVal) {

		var seria, brush, pen, numCache, point;
		var seriesPaths = paths.series;

		if(!seriesPaths) {
			return;
		}

		for (var i = 0; i < seriesPaths.length; i++) {

			if (!seriesPaths[i]) {
				continue;
			}

			seria = series[i];
			brush = seria.brush;
			pen = seria.pen;

			for (var j = 0; j < seriesPaths[i].length; j++) {

				if (bIsYVal) {
					numCache = this.getNumCache(seria.yVal);
				} else {
					numCache = this.getNumCache(seria.val);
				}

				var idx = seriesPaths[i][j].idx;
				if(numCache){
					if(useNextPoint) {
						idx = seriesPaths[i][j + 1] ? seriesPaths[i][j + 1].idx : seriesPaths[i][j].idx;
						point = numCache.getPtByIndex(idx);
					} else {
						point = numCache.getPtByIndex(idx);
					}
				}

				if (point && point.pen) {
					pen = point.pen;
				}
				if (point && point.brush) {
					brush = point.brush;
				}

				if (seriesPaths[i][j]) {
					this.drawPath(seriesPaths[i][j].path, pen, brush);
				}
			}
		}
	},

	drawPath: function(path, pen, brush)
	{
		if(!AscFormat.isRealNumber(path)){
			return;
		}
		var oPath = this.cChartSpace.GetPath(path);
		if(!oPath)
			return;
		
		if(pen)
            oPath.stroke = true;

		var cGeometry = new CGeometry2();
		this.cShapeDrawer.Clear();

        cGeometry.AddPath(oPath);
		this.cShapeDrawer.fromShape2(new CColorObj(pen, brush, cGeometry) ,this.cShapeDrawer.Graphics, cGeometry);

		this.cShapeDrawer.draw(cGeometry);
	},
	
	//****functions for chart classes****
	calculatePoint: function (x, y, size, symbol) {
		size = size / 2.69;
		var halfSize = size / 2;
		var dashDotHeight = size / 5;
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		var framePaths = null, framePathsId = null;

		/*
		 var AscFormat.SYMBOL_PICTURE = 5;*/

		path.moveTo(x * pathW, y * pathW);

		switch (symbol) {
			case AscFormat.SYMBOL_DASH: {
				path.moveTo((x - halfSize) * pathW, (y - dashDotHeight / 2) * pathW);
				path.lnTo((x + halfSize) * pathW, (y - dashDotHeight / 2) * pathW);
				path.lnTo((x + halfSize) * pathW, (y + dashDotHeight / 2) * pathW);
				path.lnTo((x - halfSize) * pathW, (y + dashDotHeight / 2) * pathW);
				break;
			}
			case AscFormat.SYMBOL_DOT: {
				path.moveTo((x) * pathW, (y - dashDotHeight / 2) * pathW);
				path.lnTo((x + halfSize) * pathW, (y - dashDotHeight / 2) * pathW);
				path.lnTo((x + halfSize) * pathW, (y + dashDotHeight / 2) * pathW);
				path.lnTo((x) * pathW, (y + dashDotHeight / 2) * pathW);
				break;
			}

			case AscFormat.SYMBOL_PLUS: {
				path.moveTo(x * pathW, (y + halfSize) * pathW);
				path.lnTo(x * pathW, (y - halfSize) * pathW);
				path.moveTo((x - halfSize) * pathW, y * pathW);
				path.lnTo((x + halfSize) * pathW, y * pathW);
				break;
			}

			case AscFormat.SYMBOL_CIRCLE: {
				path.moveTo((x + halfSize) * pathW, y * pathW);
				path.arcTo(halfSize * pathW, halfSize * pathW, 0, Math.PI * 2 * cToDeg);
				break;
			}

			case AscFormat.SYMBOL_STAR: {
				path.moveTo((x - halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x + halfSize) * pathW, (y - halfSize) * pathW);
				path.moveTo((x + halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x - halfSize) * pathW, (y - halfSize) * pathW);
				path.moveTo(x * pathW, (y + halfSize) * pathW);
				path.lnTo(x * pathW, (y - halfSize) * pathW);
				break;
			}

			case AscFormat.SYMBOL_X: {
				path.moveTo((x - halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x + halfSize) * pathW, (y - halfSize) * pathW);
				path.moveTo((x + halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x - halfSize) * pathW, (y - halfSize) * pathW);
				break;
			}

			case AscFormat.SYMBOL_TRIANGLE: {
				path.moveTo((x - size / Math.sqrt(3)) * pathW, (y + size / 3) * pathW);
				path.lnTo(x * pathW, (y - (2 / 3) * size) * pathW);
				path.lnTo((x + size / Math.sqrt(3)) * pathW, (y + size / 3) * pathW);
				path.lnTo((x - size / Math.sqrt(3)) * pathW, (y + size / 3) * pathW);
				break;
			}

			case AscFormat.SYMBOL_SQUARE: {
				path.moveTo((x - halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x - halfSize) * pathW, (y - halfSize) * pathW);
				path.lnTo((x + halfSize) * pathW, (y - halfSize) * pathW);
				path.lnTo((x + halfSize) * pathW, (y + halfSize) * pathW);
				path.lnTo((x - halfSize) * pathW, (y + halfSize) * pathW);
				break;
			}

			case AscFormat.SYMBOL_DIAMOND: {
				path.moveTo((x - halfSize) * pathW, y * pathW);
				path.lnTo(x * pathW, (y - halfSize) * pathW);
				path.lnTo((x + halfSize) * pathW, y * pathW);
				path.lnTo(x * pathW, (y + halfSize) * pathW);
				path.lnTo((x - halfSize) * pathW, y * pathW);
				break;
			}
		}

		if (symbol === AscFormat.SYMBOL_PLUS || symbol === AscFormat.SYMBOL_STAR || symbol === AscFormat.SYMBOL_X) {
			framePathsId = this.cChartSpace.AllocPath();
			framePaths = this.cChartSpace.GetPath(framePathsId);
			framePaths.moveTo((x - halfSize) * pathW, (y + halfSize) * pathW);
			framePaths.lnTo((x - halfSize) * pathW, (y - halfSize) * pathW);
			framePaths.lnTo((x + halfSize) * pathW, (y - halfSize) * pathW);
			framePaths.lnTo((x + halfSize) * pathW, (y + halfSize) * pathW);
			framePaths.lnTo((x - halfSize) * pathW, (y + halfSize) * pathW);
		}

		return {framePaths: framePathsId, path: pathId};
	},

	getYPosition: function (val, axis, ignoreAxisLimits, isRect) {
		var yPoints = axis.yPoints ? axis.yPoints : axis.xPoints;
		var isOx = axis.axPos === window['AscFormat'].AX_POS_T || axis.axPos === window['AscFormat'].AX_POS_B;
		var logBase = axis.scaling.logBase;
		if (logBase) {
			return this._getYPositionLogBase(val, yPoints, isOx, logBase);
		}

		//позиция в заисимости от положения точек на оси OY
		var result, resPos, resVal, diffVal;

		if (!yPoints[1] && val === yPoints[0].val) {
			result = yPoints[0].pos;
		} else if (val < yPoints[0].val) {
			let firstVal = yPoints[1] ? yPoints[1] : yPoints[0];
			resPos = Math.abs(firstVal.pos - yPoints[0].pos);
			resVal = firstVal.val - yPoints[0].val;
			diffVal = yPoints[0].val - val;
			if (isOx || isRect) {
				result = yPoints[0].pos - Math.abs((diffVal / resVal) * resPos);
			} else {
				result = yPoints[0].pos + Math.abs((diffVal / resVal) * resPos);
			}

			if (!ignoreAxisLimits && (result > yPoints[yPoints.length - 1].pos || result < yPoints[0].pos)) {
				result = yPoints[0].pos;
			}
		} else if (val > yPoints[yPoints.length - 1].val) {
			var test = yPoints[1] ? yPoints[1] : yPoints[0];
			resPos = Math.abs(test.pos - yPoints[0].pos);
			resVal = test.val - yPoints[0].val;
			diffVal = Math.abs(yPoints[yPoints.length - 1].val - val);

			if (!axis.isReversed()) {
				if (isOx) {
					result = yPoints[yPoints.length - 1].pos + (diffVal / resVal) * resPos;
				} else {
					result = yPoints[yPoints.length - 1].pos - (diffVal / resVal) * resPos;
				}
			} else {
				if (isOx) {
					result = yPoints[yPoints.length - 1].pos - (diffVal / resVal) * resPos;
				} else {
					result = yPoints[yPoints.length - 1].pos + (diffVal / resVal) * resPos;
				}
			}

			if (!ignoreAxisLimits && (result > yPoints[yPoints.length - 1].pos || result < yPoints[0].pos)) {
				result = yPoints[yPoints.length - 1].pos;
			}
		} else {
			/*for (var s = 0; s < yPoints.length; s++) {
				if (val >= yPoints[s].val && val <= yPoints[s + 1].val) {
					resPos = Math.abs(yPoints[s + 1].pos - yPoints[s].pos);
					resVal = yPoints[s + 1].val - yPoints[s].val;

					if(resVal === 0) {
						resVal = 1;
					}

					var startPos = yPoints[s].pos;

					if (!isOx) {
						if (!axis.isReversed()) {
							result = -(resPos / resVal) * (Math.abs(val - yPoints[s].val)) + startPos;
						} else {
							result = (resPos / resVal) * (Math.abs(val - yPoints[s].val)) + startPos;
						}
					} else {
						if (axis.isReversed()) {
							result = -(resPos / resVal) * (Math.abs(val - yPoints[s].val)) + startPos;
						} else {
							result = (resPos / resVal) * (Math.abs(val - yPoints[s].val)) + startPos;
						}
					}
					break;
				}
			}*/


			var getResult = function(index, isChartEx) {
				resPos = Math.abs(yPoints[index + 1].pos - yPoints[index].pos);
				resVal = yPoints[index + 1].val - yPoints[index].val;

				if(resVal === 0) {
					resVal = 1;
				}

				var res;
				var startPos = yPoints[index].pos;

				// for the cases of reversed axis, direction is calculated 
				let firstDirection = !isOx && (!axis.isReversed() || isChartEx) ? -1 : 1;
				let secondDirection = isOx && (axis.isReversed() || isChartEx) ? -1 : 1;

				res = (firstDirection * secondDirection * (resPos / resVal) * (Math.abs(val - yPoints[index].val))) + startPos;

				return res;
			};

			var i = 0, j = yPoints.length - 1, k;
			const isChartEx = this.cChartSpace.isChartEx();
			while (i <= j) {
				k = Math.floor((i + j) / 2);

				if (val >= yPoints[k].val && yPoints[k + 1] && val <= yPoints[k + 1].val) {
					result = getResult(k, isChartEx);
					break;
				} else if (val < yPoints[k].val) {
					j = k - 1;
				} else {
					i = k + 1;
				}
			}

		}

		return result;
	},

	_getYPositionLogBase: function (val, yPoints, isOx, logBase) {
		// first check dependencies to be numeric and valid
		if (!val && val <= 0) {
			return this.axesChart[0].axis.xPoints ? (this.calcProp.heightCanvas / this.calcProp.pxToMM) : 0;
		}
		if (!yPoints || yPoints.length < 2) {
			return null;
		}
		logBase = (logBase < 2) ? 2 : logBase;

		const startingPoint = Math.log(yPoints[yPoints.length - 1].val) / Math.log(logBase);
		const currentPoint = Math.log(val) / Math.log(logBase);
		const pointsDiff = startingPoint - currentPoint;
		const startingPos = yPoints[yPoints.length - 1].pos;
		const stepDistance = yPoints[yPoints.length - 2].pos - yPoints[yPoints.length - 1].pos;
		return (stepDistance * pointsDiff) + startingPos;

		// var logVal = Math.log(val) / Math.log(logBase);
		// var result;

		// //TODO переписать функцию!
		// var parseVal, maxVal, minVal, startPos = 0, diffPos;

		// if (yPoints.length < 2) {
		// 	return parseFloat("0." + logVal.toString().split(".")[1]) * yPoints[0].pos;
		// }
		// if (logVal < 0) {
		// 	parseVal = logVal.toString().split(".");
		// 	maxVal = Math.pow(logBase, parseVal[0]);
		// 	minVal = Math.pow(logBase, parseFloat(parseVal[0]) - 1);
		// 	for (var i = 0; i < yPoints.length; i++) {
		// 		if (yPoints[i].val < maxVal && yPoints[i].val >= minVal) {
		// 			if(yPoints[i + 1]) {
		// 				startPos = yPoints[i + 1].pos;
		// 				diffPos = yPoints[i].pos - yPoints[i + 1].pos;
		// 			} else {
		// 				startPos = yPoints[i].pos;
		// 				diffPos = yPoints[i - 1].pos - yPoints[i].pos;
		// 			}

		// 			break;
		// 		}
		// 	}
		// 	result = startPos + parseFloat("0." + parseVal[1]) * diffPos;
		// } else {
		// 	parseVal = logVal.toString().split(".");
		// 	minVal = Math.pow(logBase, parseVal[0]);
		// 	maxVal = Math.pow(logBase, parseFloat(parseVal[0]) + 1);
		// 	for (var i = 0; i < yPoints.length; i++) {
		// 		if (yPoints[i].val < maxVal && yPoints[i].val >= minVal) {
		// 			if(yPoints[i + 1]) {
		// 				startPos = yPoints[i].pos;
		// 				diffPos = yPoints[i].pos - yPoints[i + 1].pos;
		// 			} else {
		// 				startPos = yPoints[i].pos;
		// 				diffPos = yPoints[i - 1].pos - yPoints[i].pos;
		// 			}
		// 			break;
		// 		}
		// 	}
		// 	result = startPos - parseFloat("0." + parseVal[1]) * diffPos;
		// }

		// return result;
	},

	getLogarithmicValue: function (val, logBase) {
		if (val <= 0) {
			return 0;
		}

		var logVal = Math.log(val) / Math.log(logBase);

		var temp = 0;
		if (logVal > 0 && logVal < 1) {
			for (var l = 0; l < logVal; l++) {
				if (l !== 0) {
					temp += Math.pow(logBase, l);
				}
				if (l + 1 > logVal) {
					temp += (logVal - l) * Math.pow(logBase, l + 1);
					break;
				}
			}
		} else {
			var parseTemp = logVal.toString().split(".");
			temp = Math.pow(logBase, parseFloat(parseTemp[0]));

			temp = temp - temp * parseFloat("0." + parseTemp[1]);
		}

		return temp;
	},
	
	
	
	//****for 3D****
	_calaculate3DProperties: function (chartSpace) {
		var widthCanvas = this.calcProp.widthCanvas;
		var heightCanvas = this.calcProp.heightCanvas;
		var left = this.calcProp.chartGutter._left;
		var right = this.calcProp.chartGutter._right;
		var bottom = this.calcProp.chartGutter._bottom;
		var top = this.calcProp.chartGutter._top;

		standartMarginForCharts = 17;


		this.processor3D =
			new AscFormat.Processor3D(widthCanvas, heightCanvas, left, right, bottom, top, chartSpace, this);
		this.processor3D.calaculate3DProperties();
		this.processor3D.correctPointsPosition(chartSpace);
	},

	_convertAndTurnPoint: function (x, y, z, isNScale, isNRotate, isNProject) {
		return this.processor3D.convertAndTurnPoint(x, y, z, isNScale, isNRotate, isNProject);
	},

	//position of catAx labels(left or right) - returns false(left of axis)/true(right of axis) or null(standard position)
	calculatePositionLabelsCatAxFromAngle: function (chartSpace) {
		var res = null;
		var oView3D = chartSpace.chart.getView3d();
		var angleOy = oView3D && oView3D.rotY ? (-oView3D.rotY / 360) * (Math.PI * 2) : 0;
		if (oView3D && !oView3D.getRAngAx() && angleOy !== 0) {
			angleOy = Math.abs(angleOy);
			res = angleOy >= Math.PI / 2 && angleOy < 3 * Math.PI / 2;
		}

		return res;
	},
	
	
	//****accessory functions****
	_getSumArray: function (arr, isAbs) {
		if (typeof(arr) === 'number') {
			return arr;
		} else if (typeof(arr) === 'string') {
			return parseFloat(arr);
		}

		var i, sum;
		for (i = 0, sum = 0; i < arr.length; i++) {
			if (typeof(arr[i]) === 'object' && arr[i].val != null && arr[i].val != undefined) {
				sum += parseFloat(isAbs ? Math.abs(arr[i].val) : arr[i].val);
			} else if (arr[i]) {
				sum += isAbs ? Math.abs(arr[i]) : arr[i];
			}
		}
		return sum;
	},

	_getMaxMinValueArray: function (array) {
		var max = 0, min = 0;
		for (var i = 0; i < array.length; i++) {
			for (var j = 0; j < array[i].length; j++) {
				if (i === 0 && j === 0) {
					min = array[i][j];
					max = array[i][j];
				}

				if (array[i][j] > max) {
					max = array[i][j];
				}

				if (array[i][j] < min) {
					min = array[i][j];
				}
			}
		}
		return {max: max, min: min};
	},
	
	_findPrevValue: function(originalData, num, max) {
		var summ = 0;
		for (var i = 0; i <= max; i++) {
			if (originalData[num][max] >= 0) {
				if (originalData[num][i] >= 0)
					summ += originalData[num][i];
			}

			else {
				if (originalData[num][i] < 0)
					summ += originalData[num][i];
			}
		}
		return summ;
	},

	_getFirstDegree: function (val) {
		var secPart = val.toString().split('.');
		var numPow = 1, tempMax, expNum;

		if (secPart[1] && secPart[1].toString().indexOf('e+') != -1 && secPart[0] && secPart[0].toString().length == 1) {
			expNum = secPart[1].toString().split('e+');
			numPow = Math.pow(10, expNum[1]);
		} else if (secPart[1] && secPart[1].toString().indexOf('e-') != -1 && secPart[0] && secPart[0].toString().length == 1) {
			expNum = secPart[1].toString().split('e');
			numPow = Math.pow(10, expNum[1]);
		} else if (0 != secPart[0]) {
			numPow = Math.pow(10, secPart[0].toString().length - 1);
		} else if (0 == secPart[0] && secPart[1] != undefined) {
			tempMax = val;
			var num = 0;
			while (1 > tempMax) {
				tempMax = tempMax * 10;
				num--;
			}
			numPow = Math.pow(10, num);
			val = tempMax;
		}

		if (tempMax == undefined) {
			val = val / numPow;
		}

		return {val: val, numPow: numPow};
	},

	getIdxPoint: function (seria, index, bXVal) {
		var res = null;

		var ser;
		if (bXVal) {
			ser = seria.val ? seria.val : seria.xVal;
		} else {
			ser = seria.val ? seria.val : seria.yVal;
		}

		if (!ser) {
			return null;
		}

		var numCache = this.getNumCache(ser);
		var pts = numCache ? numCache.pts : null;

		if (pts == null) {
			return null;
		}

		if (pts[index] && pts[index].idx === index) {
			return pts[index];
		}

		//TODO need binary search! start with index
		for (var i = 0; i < pts.length; i++) {
			if (pts[i].idx === index) {
				res = pts[i];
				break;
			}
		}

		return res;
	},

	getPointByIndex: function (seria, index, bXVal) {
		var ser;
		if (bXVal) {
			ser = seria.val ? seria.val : seria.xVal;
		} else {
			ser = seria.val ? seria.val : seria.yVal;
		}

		if (!ser) {
			return null;
		}

		//todo use getNumCache
		var oCache = (ser.numRef && ser.numRef.numCache) ||  ser.numLit;
		if(oCache) {
			return oCache.getPtByIndex(index);
		}

		return null;
	},

	getPtCount: function (series) {
		let numCache, oVal;
		for (var i = 0; i < series.length; i++) {
			if (series[i].val) {
				oVal = series[i].val;
			} else if (series[i].xVal) {
				oVal = series[i].xVal;
			} else if (series[i].yVal) {
				oVal = series[i].yVal;
			}

			numCache = oVal && oVal.getNumCache && oVal.getNumCache();
			if (numCache && numCache.ptCount) {
				return numCache.ptCount;
			}
		}

		return 0;
	},

	_roundValues: function (values) {
		if (values.length) {
		  for (let i = 0; i < values.length; i++) {
			values[i] = _roundValue(values[i]);
		  }
		}
	  
		return values;
	},
	
	
	//***spline functions***
	//TODO пока включаю calculate_Bezier. проверить корретность calculate_Bezier2!
	calculate_Bezier2: function (x, y, x1, y1, x2, y2, x3, y3) {
		var pts = [], bz = [];

		pts[0] = {x: x, y: y};
		pts[1] = {x: x1, y: y1};
		pts[2] = {x: x2, y: y2};
		pts[3] = {x: x3, y: y3};

		var d01 = this.XYZDist(pts[0], pts[1]);
		var d12 = this.XYZDist(pts[1], pts[2]);
		var d23 = this.XYZDist(pts[2], pts[3]);
		var d02 = this.XYZDist(pts[0], pts[2]);
		var d13 = this.XYZDist(pts[1], pts[3]);

		//start point
		bz[0] = pts[1];


		//control points
		if ((d02 / 6 < d12 / 2) && (d13 / 6 < d12 / 2)) {
			var f;
			if (x !== x1) {
				f = 1 / 6;
			} else {
				f = 1 / 3;
			}

			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), f));

			if (x2 !== x3) {
				f = 1 / 6;
			} else {
				f = 1 / 3;
			}

			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), f))
		} else if ((d02 / 6 >= d12 / 2) && (d13 / 6 >= d12 / 2)) {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13));
		} else if ((d02 / 6 >= d12 / 2)) {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13 * (d13 / d02)));
		} else {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02 * (d02 / d13)));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13));
		}

		//end point
		bz[3] = pts[2];

		return bz;
	},

	//***spline functions***
	calculate_Bezier: function (x, y, x1, y1, x2, y2, x3, y3, t) {
		var pts = [], bz = [];

		pts[0] = {x: x, y: y};
		pts[1] = {x: x1, y: y1};
		pts[2] = {x: x2, y: y2};
		pts[3] = {x: x3, y: y3};

		var d01 = this.XYZDist(pts[0], pts[1]);
		var d12 = this.XYZDist(pts[1], pts[2]);
		var d23 = this.XYZDist(pts[2], pts[3]);
		var d02 = this.XYZDist(pts[0], pts[2]);
		var d13 = this.XYZDist(pts[1], pts[3]);

		bz[0] = pts[1];

		if ((d02 / 6 < d12 / 2) && (d13 / 6 < d12 / 2)) {
			var f;
			if (x !== x1) {
				f = 1 / 6;
			} else {
				f = 1 / 3;
			}

			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), f));

			if (x2 !== x3) {
				f = 1 / 6;
			} else {
				f = 1 / 3;
			}

			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), f))
		}

		else if ((d02 / 6 >= d12 / 2) && (d13 / 6 >= d12 / 2)) {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13));
		} else if ((d02 / 6 >= d12 / 2)) {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13 * (d13 / d02)));
		} else {
			bz[1] = this.XYZAdd(pts[1], this.XYZMult(this.XYZSub(pts[2], pts[0]), d12 / 2 / d02 * (d02 / d13)));
			bz[2] = this.XYZAdd(pts[2], this.XYZMult(this.XYZSub(pts[1], pts[3]), d12 / 2 / d13));
		}

		bz[3] = pts[2];

		var pt = this._bezier4(bz[0], bz[1], bz[2], bz[3], t);

		return [pt.x, pt.y];
	},

	_bezier4: function (p1, p2, p3, p4, t) {
		var mum1, mum13, t3, mum12, t2;
		var p = {};

		mum1 = 1 - t;
		mum13 = mum1 * mum1 * mum1;
		mum12 = mum1 * mum1;
		t2 = t * t;
		t3 = t * t * t;

		p.x = mum13 * p1.x + 3 * t * mum12 * p2.x + 3 * t2 * mum1 * p3.x + t3 * p4.x;
		p.y = mum13 * p1.y + 3 * t * mum12 * p2.y + 3 * t2 * mum1 * p3.y + t3 * p4.y;

		return p;
	},

	XYZAdd: function(a, b)
	{
		return {x: a.x + b.x, y: a.y + b.y};
	},

	XYZSub: function(a, b)
	{
		return {x: a.x - b.x, y: a.y - b.y};
	},

	XYZMult: function(a, b)
	{
		return {x: a.x * b, y: a.y * b};
	},

	XYZDist: function(a, b)
	{
		return Math.pow((Math.pow((a.x - b.x), 2) + Math.pow((a.y - b.y), 2)), 0.5);
	},

	calculateCountSeries: function (chart) {
		var series = chart.series;
		var typeChart = chart.getObjectType();
		var counter = 0, numCache, seriaVal, ptCount;
		for (var i = 0; i < series.length; i++) {
			seriaVal = series[i].val ? series[i].val : series[i].yVal;
			numCache = this.getNumCache(seriaVal);
			if (!series[i].isHidden) {
				if (AscDFH.historyitem_type_PieChart === typeChart) {
					ptCount = 1;
				} else {
					if (numCache && (ptCount === undefined || ptCount < numCache.ptCount)) {
						ptCount = numCache.ptCount;
					}
				}

				counter++;
			} else if (3 === this.nDimensionCount) {
				counter++;
			}
		}

		if (AscDFH.historyitem_type_PieChart === typeChart) {
			counter = 1;
		}

		return {series: counter, points: ptCount};
	},

	calculateFirstChartCountSeries: function () {
		let chart = this.cChartSpace && this.cChartSpace.chart && this.cChartSpace.chart.plotArea && this.cChartSpace.chart.plotArea.charts &&
			this.cChartSpace.chart.plotArea.charts[0];
		return chart ? this.calculateCountSeries(chart) : null;
	},
	
	//вспомогательные функции работающие с тремя координатами
	//получаем к-ты уравнения прямой по 2 точкам
	getLineEquation: function (point1, point2) {
		var x0 = point1.x, y0 = point1.y, z0 = point1.z;
		var x1 = point2.x, y1 = point2.y, z1 = point2.z;


		/*x - x0 	 =  	y - y0 	 =  	z - z0
		 x1 - x0 			y1 - y0 		z1 - z0

		 l 					m 				n
		 */

		var l = x1 - x0;
		var m = y1 - y0;
		var n = z1 - z0;

		//check line
		/*var x123 = (point1.x - x0) / (x1 - x0);
		 var y123 = (point1.y - y0) / (y1 - y0);
		 var z123 = (point1.z - z0) / (z1 - z0);

		 var x321 = (point2.x - x0) / (x1 - x0);
		 var y321 = (point2.y - y0) / (y1 - y0);
		 var z321 = (point2.z - z0) / (z1 - z0);*/

		return {l: l, m: m, n: n, x1: x0, y1: y0, z1: z0};
	},

	getLineEquation2d: function (point1, point2) {
		var x1 = point1.x, y1 = point1.y;
		var x2 = point2.x, y2 = point2.y;

		//y = kx + b
		var k = (y2 - y1) / (x2 - x1);
		var b = y1 - k * x1;

		return {k: k, b: b};
	},

	isIntersectionLineAndLine2d: function (x1, y1, x2, y2, x3, y3, x4, y4) {
	
		if ((x1 === x2 && y1 === y2) || (x3 === x4 && y3 === y4)) {
			return false;
		}

		var k = ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1));

		if (k === 0) {
			return false;
		}

		var ua = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / k;
		var ub = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / k;

		if (ua < 0 || ua > 1 || ub < 0 || ub > 1) {
			return false;
		}

		var x = x1 + ua * (x2 - x1);
		var y = y1 + ua * (y2 - y1);

		return { x: x, y: y };
	},

	isIntersectionLineAndLine: function (equation1, equation2) {
		var xo = equation1.x1;
		var yo = equation1.y1;
		var zo = equation1.z1;
		var p = equation1.l;
		var q = equation1.m;
		var r = equation1.n;

		var x1 = equation2.x1;
		var y1 = equation2.y1;
		var z1 = equation2.z1;
		var p1 = equation2.l;
		var q1 = equation2.m;
		var r1 = equation2.n;

		var x = (xo * q * p1 - x1 * q1 * p - yo * p * p1 + y1 * p * p1) / (q * p1 - q1 * p);
		var y = (yo * p * q1 - y1 * p1 * q - xo * q * q1 + x1 * q * q1) / (p * q1 - p1 * q);
		var z = (zo * q * r1 - z1 * q1 * r - yo * r * r1 + y1 * r * r1) / (q * r1 - q1 * r);

		return {x: x, y: y, z: z};
	},

	//поиск точки пересечения плоскости и прямой
	isIntersectionPlainAndLine: function (plainEquation, lineEquation) {
		var A = plainEquation.a;
		var B = plainEquation.b;
		var C = plainEquation.c;
		var D = plainEquation.d;

		var l = lineEquation.l;
		var m = lineEquation.m;
		var n = lineEquation.n;
		var x1 = lineEquation.x1;
		var y1 = lineEquation.y1;
		var z1 = lineEquation.z1;


		//x - x1		y - y1		z - z1
		//			=			=			t
		//  l			  m		 	  n

		/*x = t * l + x1
		 y = t * m + y1
		 z = t * n + z1*/


		/*A * x + B * y + C * z + D = 0

		 A * (t * l + x1) + B * (t * m + y1) + C * (t * n + z1) + D = 0;

		 A * t * l + A * x1 + B * t * m + B * y1 + C * t * n + C * z1 + D

		 A * t * l + B * t * m + C * t * n       + A * x1 + B * y1 + C * z1 + D*/

		var t = -(A * x1 + B * y1 + C * z1 + D) / (A * l + B * m + C * n);

		var x = t * l + x1;
		var y = t * m + y1;
		var z = t * n + z1;

		return {x: x, y: y, z: z};
	},
	
	isIntersectionPlainAndLineSegment: function(plainEquation, point1, point2, projPoint1, projPoint2)
	{
		var res = null;
		var lineEquation = this.getLineEquation(point1, point2);
		
		var intersection = this.isIntersectionPlainAndLine(plainEquation, lineEquation);
		var convertResult = this._convertAndTurnPoint(intersection.x, intersection.y, intersection.z);
		var isBetweenX = (convertResult.x >= projPoint1.x && convertResult.x <= projPoint2.x) || (convertResult.x <= projPoint1.x && convertResult.x >= projPoint2.x);
		var isBetweenY = (convertResult.y >= projPoint1.y && convertResult.y <= projPoint2.y) || (convertResult.y <= projPoint1.y && convertResult.y >= projPoint2.y);
		
		//принадлежит ли даная точка отрезку
		if(isBetweenX && isBetweenY)
		{
			var vectorMultiplication = ((convertResult.x - projPoint1.x) * (projPoint2.y - projPoint1.y)) - ((convertResult.y - projPoint1.y) * (projPoint2.x - projPoint1.x));
			if(Math.round(vectorMultiplication) === 0)
			{
				res = convertResult;
			}
		}
		
		return res;
	},
	
	isPoint2DLieOnLine: function(lineEquation, point)
	{
		return Math.round(point.y * 1000) / 1000 === Math.round((lineEquation.k * point.x + lineEquation.b) * 1000) / 1000;
	},
	
	isPointLieIntoPlane: function(planeEquation, point)
	{
		var resEquation = planeEquation.a * point.x + planeEquation.b * point.y + planeEquation.c * point.z + planeEquation.d;
		return Math.round(resEquation) === 0;
	},
	
	isPointsLieIntoOnePlane2: function(point1, point2, point3, point4)
	{
		var bRes = false;
		
		var plain1 = this.getPlainEquation(point1, point2, point3);
		var plain2 = this.getPlainEquation(point3, point4, point1);
		
		//todo пересмотреть округление
		if(Math.round(plain1.a) === Math.round(plain2.a) && Math.round(plain1.b) === Math.round(plain2.b) && Math.round(plain1.c) === Math.round(plain2.c) && Math.round(plain1.d) === Math.round(plain2.d))
		{
			bRes = true;
		}
		
		return bRes;
	},
	
	isPointsLieIntoOnePlane: function(point1, point2, point3, point4)
	{
		var bRes = false;
		
		var plain1 = this.getPlainEquation(point1, point2, point3);
		
		if(this.isPointLieIntoPlane(plain1, point4))
		{
			bRes = true;
		}
		
		return bRes;
	},
	
	isPointsLieIntoOnePlane3: function(point1, point2, point3, point4)
	{
		/*var vector1 = Math.sqrt(Math.pow(point4.x - point1.x, 2) + Math.pow(point4.y - point1.y, 2) + Math.pow(point4.z - point1.z, 2));//DA
		var vector2 = Math.sqrt(Math.pow(point4.x - point2.x, 2) + Math.pow(point4.y - point2.y, 2) + Math.pow(point4.z - point2.z, 2));//DB
		var vector2 = Math.sqrt(Math.pow(point4.x - point3.x, 2) + Math.pow(point4.y - point3.y, 2) + Math.pow(point4.z - point3.z, 2));//DC*/
		
		var vector1 = {x: point4.x - point1.x, y: point4.y - point1.y, z: point4.z - point1.z};//DA
		var vector2 = {x: point4.x - point2.x, y: point4.y - point2.y, z: point4.z - point2.z};//DB
		var vector3 = {x: point4.x - point3.x, y: point4.y - point3.y, z: point4.z - point3.z};//DC
		
		var a1 = vector1.x;
		var b1 = vector1.y;
		var c1 = vector1.z;
		var a2 = vector2.x;
		var b2 = vector2.y;
		var c2 = vector2.z;
		var a3 = vector3.x;
		var b3 = vector3.y;
		var c3 = vector3.z;
		
		var res = a1 * b2 * c3 + a3 * b1 * c2 + a2 * b3 * c1 - a3 * b2 * c1 - a1 * b3 * c2 - a2 * b1 * c3;
		
		return res;
	},
	
	//получаем площадь произвольного выпуклого четырехугольника
	getAreaQuadrilateral: function(point0, point1, point2, point3)
	{
		//длины сторон
		var a = Math.sqrt(Math.pow(point3.x - point0.x, 2) + Math.pow(point3.y - point0.y, 2));
		var b = Math.sqrt(Math.pow(point1.x - point0.x, 2) + Math.pow(point1.y - point0.y, 2));
		var c = Math.sqrt(Math.pow(point2.x - point1.x, 2) + Math.pow(point2.y - point1.y, 2));
		var d = Math.sqrt(Math.pow(point3.x - point2.x, 2) + Math.pow(point3.y - point2.y, 2));
		
		//длины диагоналей
		var e = Math.sqrt(Math.pow(point3.x - point1.x, 2) + Math.pow(point3.y - point1.y, 2));
		var f = Math.sqrt(Math.pow(point0.x - point2.x, 2) + Math.pow(point0.y - point2.y, 2));
		
		//полупериметр
		var p = (a + b + c + d) / 2;
		
		
		var res = Math.sqrt((p - a) * (p - b) * (p - c) * (p - d) + (1/4) * ((e * f + a * c + b * d) * (e * f - a * c - b * d)));
		
		return res;
	},
	
	//получаем площадь произвольного трехугольника
	getAreaTriangle: function(point0, point1, point2)
	{
		//длины сторон
		var a = Math.sqrt(Math.pow(point1.x - point0.x, 2) + Math.pow(point1.y - point0.y, 2));
		var b = Math.sqrt(Math.pow(point2.x - point1.x, 2) + Math.pow(point2.y - point1.y, 2));
		var c = Math.sqrt(Math.pow(point2.x - point0.x, 2) + Math.pow(point2.y - point0.y, 2));
		
		//полупериметр
		var p = (a + b + c ) / 2;

		return Math.sqrt(p * (p - a) * (p - b) * (p - c));
	},
	
	//из массива точек получаем минимальные/максимальные x,y,z
	getMinMaxPoints: function(points)
	{
		var minX, maxX, minY, maxY, minZ, maxZ;
		
		for(var n = 0; n < points.length; n++)
		{
			if(0 === n)
			{
				minX = points[0].x;
				maxX = points[0].x;
				minY = points[0].y;
				maxY = points[0].y;
				minZ = points[0].z;
				maxZ = points[0].z;
			}
			else
			{
				if(points[n].x < minX)
				{
					minX = points[n].x;
				}
				
				if(points[n].x > maxX)
				{
					maxX = points[n].x;
				}
				
				if(points[n].y < minY)
				{
					minY = points[n].y;
				}
				
				if(points[n].y > maxY)
				{
					maxY = points[n].y;
				}
				
				if(points[n].z < minZ)
				{
					minZ = points[n].z;
				}
				
				if(points[n].z > maxZ)
				{
					maxZ = points[n].z;
				}
			}
		}
		
		return {minX: minX, maxX: maxX, minY : minY, maxY: maxY, minZ: minZ, maxZ: maxZ};
	},
	
	//TODO временная функция для теста результата значений
	getPlainEquation3: function(point1, point2, point3)
	{
		var a1 = point1.x, b1 = point1.y, c1 = point1.z, a2 = point2.x, b2 = point2.y, c2 = point2.z, a3 = point3.x, b3 = point3.y, c3 = point3.z;
		
		var subMatrix = function(oldmat, row, col) 
		{
			var i;
			var j;
			var m = 0;
			var k = 0;
			var len = oldmat.length;
			var retmat = [];

			if ((len <= row) || (len <= col)) return 0;

			for (j = 0; j < len; j++) {
				if (j !== row) {
					retmat[k] = [];
					for (i = 0; i < len; i++) {
						if (i !== col) {
							retmat[k][m] = oldmat[j][i];
							m++;
						}
					}
					k++;
				}
				m = 0;
			}
			return retmat;
		};
		
		var determinant = function(mat) 
		{
			var i;
			var tmpVal = 0;
			var row = mat.length;
			var newDet = [];
			newDet[0] = [];

			switch (row) {
				case 1:
					return mat[0][0];
				case 2:
					return (mat[0][0] * mat[1][1] - mat[0][1] * mat[1][0]);
				default:
					for (i = 0; i < row; i++) {
						if (mat[0][i] === 0) i++;
						if (i < row) {
							newDet = subMatrix(mat, 0, i);
							if (!isNaN(mat[0][i]))
								tmpVal += parseFloat(mat[0][i]) * Math.pow(-1, i) * determinant(newDet);
						}
					}
					return tmpVal;
			}
		};
		
		var replaceCol = function(mat, col, vec) 
		{
			var i = 0;
			var j = 0;
			var tmp = [];

			for (i = 0; i < mat.length; i++) {
				tmp[i] = [];

				for (j = 0; j < mat.length; j++) {
					if (col === j)
						tmp[i][col] = vec[i];
					else
						tmp[i][j] = mat[i][j]; 

				}
			}
			return tmp;
		};

		
		var mat = [];
		mat[0] = []; mat[1] = []; mat[2] = [];
		mat[0][0] = a1; mat[0][1] = b1; mat[0][2] = c1;
		mat[1][0] = a2; mat[1][1] = b2; mat[1][2] = c2;
		mat[2][0] = a3; mat[2][1] = b3; mat[2][2] = c3;

		var det = determinant(mat);
		
		var vec = [-1, -1, -1];
		var factor;

		var a = determinant(replaceCol(mat, 0, vec));
		if (a < 0) 
		{
			factor = -1;
		}
		else 
		{
			factor = 1;
		}
		a *= factor;
		
		var b = determinant(replaceCol(mat, 1, vec)) * factor;
		var c = determinant(replaceCol(mat, 2, vec)) * factor;
		var d = det * factor;

		return {a: a, b: b, c: c, d: d};
	},
	
	//TODO если будет проблема при отрисовке граней(перекрытие друг другом), вернуть протестированную функцию  - getPlainEquation3
	getPlainEquation: function(point1, point2, point3)
	{
		var x1 = point1.x, y1 = point1.y, z1 = point1.z;
		var x2 = point2.x, y2 = point2.y, z2 = point2.z;
		var x3 = point3.x, y3 = point3.y, z3 = point3.z;
		
		var a = x2 - x1;
		var b = y2 - y1;
		var c = z2 - z1;
		var d = x3 - x1;
		var e = y3 - y1;
		var f = z3 - z1;
		
		var xK = b*f - e*c;
		var yK = -(a*f - d*c);
		var zK = a*e - d*b;
		
		var a1 = -xK;
		var b1 = -yK;
		var c1 = -zK;
		var d1 = -(-y1 * yK - x1*xK - z1 * zK);

		// var a1 = y1 *(z2 - z3) + y2 *(z3 - z1) + y3* (z1 - z2) ;
		// var b1 = z1 *(x2 - x3) + z2 *(x3 - x1) + z3 *(x1 - x2); 
		// var c1 = x1 *(y2 - y3) + x2 *(y3 - y1) + x3 *(y1 - y2) ;
		// var d1 = -(x1 * (y2 * z3 - y3 * z2) + x2 * (y3 * z1 - y1 * z3) + x3 * (y1 * z2 - y2 * z1));

		return {a: a1, b: b1, c: c1, d: d1};
	},
	
	//уравнение плоскости
	getPlainEquation2: function(point1, point2, point3)
	{
		var x1 = point1.x, y1 = point1.y, z1 = point1.z;
		var x2 = point2.x, y2 = point2.y, z2 = point2.z;
		var x3 = point3.x, y3 = point3.y, z3 = point3.z;
		
		var x21 = x2 - x1;
		var y21 = y2 - y1;
		var z21 = z2 - z1;
		
		var x31 = x3 - x1;
		var y31 = y3 - y1;
		var z31 = z3 - z1;
			
		//(x - x1)*(y21 * z31 - x21 * y31) - (y - y1)*(x21 * z31 - z21 * x31) + (z - z1)(x21 * y31 - y21 * x31) 
		
		var tempA = y21 * z31 - z21 * y31;
		
		var tempB = x21 * z31 - z21 * x31;
		
		var tempC = x21 * y31 - y21 * x31;
		
		//(x - x1)*(tempA) - (y - y1)*(tempB) + (z - z1)(tempC)
		
		//x * tempA - x1 * tempA - y * tempB + y1 * tempB + z * tempC - z1 * tempC
		
		var a = tempA;
		var b = tempB;
		var c = tempC;
		var d =  y1 * tempB - x1 * tempA - z1 * tempC;
			
		return {a: a, b: b, c: c, d: d};
	},


	//******calculate graphic objects for 3d*******
	calculateRect3D: function (points, val, isNotDrawDownVerge, isNotOnlyFrontFaces, points2) {
		var res;
		var point12, point22, point32, point42, point52, point62, point72, point82;

		var point1 = points[0];
		var point2 = points[1];
		var point3 = points[2];
		var point4 = points[3];
		var point5 = points[4];
		var point6 = points[5];
		var point7 = points[6];
		var point8 = points[7];

		if (points2) {
			point12 = points2[0];
			point22 = points2[1];
			point32 = points2[2];
			point42 = points2[3];
			point52 = points2[4];
			point62 = points2[5];
			point72 = points2[6];
			point82 = points2[7];
		} else if (points) {
			point12 = point1;
			point22 = point2;
			point32 = point3;
			point42 = point4;
			point52 = point5;
			point62 = point6;
			point72 = point7;
			point82 = point8;
		}

		var frontPaths = [];
		var darkPaths = [];

		var addPathToArr = function (isFront, face, index) {
			frontPaths[index] = null;
			darkPaths[index] = null;

			if (isFront) {
				frontPaths[index] = face;
			} else {
				darkPaths[index] = face;
			}
		};

		var face;
		//front
		face = this._calculatePathFace(point1, point5, point8, point4, true);
		addPathToArr(this._isVisibleVerge3D(point52, point12, point42, val), face, 0);

		//down
		if (val === 0 && this.calcProp.type === c_oChartTypes.Bar) {
			face = this._calculatePathFace(point1, point2, point3, point4, true);
			addPathToArr(true, face, 1);
		} else {
			face = this._calculatePathFace(point1, point2, point3, point4, true);
			addPathToArr(this._isVisibleVerge3D(point42, point12, point22, val), face, 1);
		}


		//left
		if (val === 0 && this.calcProp.type === c_oChartTypes.HBar) {
			face = this._calculatePathFace(point1, point5, point6, point2, true);
			addPathToArr(!isNotDrawDownVerge, face, 2);
		} else {
			face = this._calculatePathFace(point1, point5, point6, point2, true);
			addPathToArr((!isNotDrawDownVerge && this._isVisibleVerge3D(point22, point12, point52, val)), face, 2);
		}

		//right
		if (val === 0 && this.calcProp.type === c_oChartTypes.HBar) {
			face = this._calculatePathFace(point4, point8, point7, point3, true);
			addPathToArr(true, face, 3);
		} else {
			face = this._calculatePathFace(point4, point8, point7, point3, true);
			addPathToArr(this._isVisibleVerge3D(point82, point42, point32, val), face, 3);
		}

		//up
		if (val === 0 && this.calcProp.type === c_oChartTypes.Bar) {
			face = this._calculatePathFace(point5, point6, point7, point8, true);
			addPathToArr(true, face, 4);
		} else {
			face = this._calculatePathFace(point5, point6, point7, point8, true);
			addPathToArr(this._isVisibleVerge3D(point62, point52, point82, val), face, 4);
		}

		//unfront
		face = this._calculatePathFace(point2, point6, point7, point3, true);
		addPathToArr(this._isVisibleVerge3D(point32, point22, point62, val), face, 5);

		if (!isNotOnlyFrontFaces) {
			res = frontPaths;
		} else {
			res = {frontPaths: frontPaths, darkPaths: darkPaths};
		}

		return res;
	},

	_calculatePathFace: function(point1, point2, point3, point4, isConvertPxToMM)
	{
		var pxToMm = 1;
		if(isConvertPxToMM)
		{
			pxToMm = this.calcProp.pxToMM;
		} 

		var pathId = this.cChartSpace.AllocPath();
		var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		path.moveTo(point1.x / pxToMm * pathW, point1.y / pxToMm * pathH);
		path.lnTo(point2.x / pxToMm * pathW, point2.y / pxToMm * pathH);
		path.lnTo(point3.x / pxToMm * pathW, point3.y / pxToMm * pathH);
		path.lnTo(point4.x / pxToMm * pathW, point4.y / pxToMm * pathH);
		path.lnTo(point1.x / pxToMm * pathW, point1.y / pxToMm * pathH);

		return pathId;
	},
	
	calculatePathFacesArray: function(faceArr, isConvertPxToMM)
	{
		var pxToMm = 1;
		if(isConvertPxToMM)
		{
			pxToMm = this.calcProp.pxToMM;
		} 

		var pathId = this.cChartSpace.AllocPath();
		var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;
		
		var firstPoint = null;
		for(var i = 0; i < faceArr.length; i++)
		{
			if(faceArr[i] && faceArr[i].point)
			{
				if(null === firstPoint)
				{
					path.moveTo(faceArr[i].point.x / pxToMm * pathW, faceArr[i].point.y / pxToMm * pathH);
					firstPoint = faceArr[i].point;
				}
				else
				{
					path.lnTo(faceArr[i].point.x / pxToMm * pathW, faceArr[i].point.y / pxToMm * pathH);
				}
			}
		}
		
		if(null !== firstPoint)
		{
			path.lnTo(firstPoint.x / pxToMm * pathW, firstPoint.y / pxToMm * pathH);
		}

		return pathId;
	},
	
	_isVisibleVerge3D: function(k, n, m, val, cylinder)
	{
		//roberts method - calculate normal
		var aX = n.x * m.y - m.x * n.y;
		var bY = - (k.x * m.y - m.x * k.y);
		var cZ = k.x * n.y - n.x * k.y;
		var visible = aX + bY + cZ;
		var _valAx = this.cChartSpace.chart.plotArea.valAx;

		var result;
		if(this.calcProp.type === c_oChartTypes.Bar)
		{
			result = val > 0 && visible < 0 || val < 0 && visible > 0;
			if(!(this.calcProp.subType === "stacked") && !(this.calcProp.subType === "stackedPer") && _valAx.isReversed())
				result = !result;
		}
		else if(this.calcProp.type === c_oChartTypes.Line)
		{
			result = visible < 0;
		}
		else if(this.calcProp.type === c_oChartTypes.HBar)
		{
			result = val < 0 && visible < 0 || val > 0 && visible > 0;
			
			if(!(this.calcProp.subType === "stacked") && !(this.calcProp.subType === "stackedPer") && _valAx.isReversed())
				result = !result;
		}

		if (!(this.calcProp.subType === "stacked") && !(this.calcProp.subType === "stackedPer")) {
			if ((_valAx.crossAx.crosses === AscFormat.CROSSES_MAX && val > 0) ||
				(_valAx.crossAx.crosses === AscFormat.CROSSES_MIN && val < 0)) {
				result = !result;
			} else if (_valAx.crossAx.crossesAt) {
				if ((val > 0 && val < _valAx.crossAx.crossesAt) || (val < 0 && val > _valAx.crossAx.crossesAt)) {
					result = !result;
				}
			}
		}

		if (cylinder) {
			result = !result;
		}
		
		return result;
	},

	calculateCylinder: function (points, val, isNotOnlyFrontFaces, notDraw, isNotAllPointsVisible, cone) {
		var res;
		var segmentPoints = points[0];
		var segmentPoints2 = points[1];
		var point1 = points[2];
		var point2 = points[3];
		var point4 = points[4];
		var point5 = points[5];
		var point6 = points[6];
		var point8 = points[7];
		var sortCylinderPoints1 = points[8];
		var sortCylinderPoints2 = points[9];

		var frontPaths = [];
		var darkPaths = [];

		// условие для конусов
		var isVisibleReverse = cone ? isNotAllPointsVisible : true;

		var addPathToArr = function (isFront, face, index) {
			frontPaths[index] = null;
			darkPaths[index] = null;

			if (isFront) {
				frontPaths[index] = face;
			} else {
				darkPaths[index] = face;
			}
		};

		var face, faceFront;

		//front
		faceFront = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, false, true, isNotAllPointsVisible);
		addPathToArr(this._isVisibleVerge3D(sortCylinderPoints1[0], sortCylinderPoints1[sortCylinderPoints1.length - 1],
			sortCylinderPoints2[0], val, !isVisibleReverse), faceFront, 0);

		//down
		if (val === 0) {
			face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
			addPathToArr(true, face, 1);
		} else {
			if (cone) {
				if (this._isVisibleVerge3D(point4, point1, point2, val)) {
					face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
					addPathToArr(true, face, 1);
				} else {
					face = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, true, true, true);
					addPathToArr(true, face, 1);
				}
			} else {
				face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
				addPathToArr(this._isVisibleVerge3D(point4, point1, point2, val), face, 1);
			}
		}

		//up
		if (!notDraw) {
			if (val === 0) {
				face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
				addPathToArr(true, face, 4);
			} else {
				if (cone) {
					if (this._isVisibleVerge3D(point6, point5, point8, val)) {
						face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
						addPathToArr(true, face, 4);
					} else {
						face = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, true, false, true, true);
						addPathToArr(true, face, 4);
					}
				} else {
					face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
					addPathToArr(this._isVisibleVerge3D(point6, point5, point8, val), face, 4);
				}
			}
		}
		
		//unfront
		faceFront = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, false, true, isNotAllPointsVisible);
		addPathToArr(this._isVisibleVerge3D(sortCylinderPoints1[0], sortCylinderPoints1[sortCylinderPoints1.length - 1],
			sortCylinderPoints2[0], val, isVisibleReverse), faceFront, 5);

		if (!isNotOnlyFrontFaces) {
			res = frontPaths;
		} else {
			res = { frontPaths: frontPaths, darkPaths: darkPaths };
		}

		return res;
	},

	// Draws a rectangle, 
	// params are start x and y, width, height and isCOnverted
	_calculateRect: function (x, y, w, h, isPxToMmConverted) {
		if (!this.cChartSpace || !this.calcProp) {
			return null;
		}
		const pathId = this.cChartSpace.AllocPath();
		const path = this.cChartSpace.GetPath(pathId);

		const pathH = this.calcProp.pathH;
		const pathW = this.calcProp.pathW;

		if (!isPxToMmConverted && AscFormat.isRealNumber(this.calcProp.pxToMM)) {
			const pxToMm = this.calcProp.pxToMM;
			x = x / pxToMm;
			y = y / pxToMm;
			w = w / pxToMm;
			h = h / pxToMm;
		}

		path.moveTo(x * pathW, y * pathH);
		path.lnTo(x * pathW, (y - h) * pathH);
		path.lnTo((x + w) * pathW, (y - h) * pathH);
		path.lnTo((x + w) * pathW, y * pathH);
		path.lnTo(x * pathW, y * pathH);

		return pathId;
	},

	// Draws a line, 
	// params are point1, length, constPoint, isVertical, and isPxToMmConverted
	_calculateLine: function (p1, l, constP, isVertical, isPxToMmConverted) {
		if (!this.cChartSpace || !this.calcProp) {
			return null;
		}

		const pathId = this.cChartSpace.AllocPath();
		const path = this.cChartSpace.GetPath(pathId);

		const pathH = this.calcProp.pathH;
		const pathW = this.calcProp.pathW;

		if (!isPxToMmConverted) {
			const pxToMm = this.calcProp.pxToMM;
			p1 = p1 / pxToMm;
			l = l / pxToMm;
			constP = constP / pxToMm;
		}

		if (isVertical) {
			path.moveTo(constP * pathW, p1 * pathH);
			path.lnTo(constP * pathW, (p1 - l) * pathH);
		} else {
			path.moveTo(p1 * pathW, constP * pathH);
			path.lnTo(p1 + l * pathW, constP * pathH);
		}

		return pathId;
	},

	_calculatePathFaceCylinder: function(segmentPoints, segmentPoints2, up, down, isConvertPxToMM, check)
	{
		var pxToMm = 1;
		if(isConvertPxToMM)
		{
			pxToMm = this.calcProp.pxToMM;
		} 

		var pathId = this.cChartSpace.AllocPath();
		var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		if (up) {
			for (var i = 0; i < segmentPoints2.length; i++) {
				if (i % 2 === 0) {
					path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
				}
			}
			if (check) {
				for (var i = segmentPoints2.length - 1; i >= 0; i--) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
					}
				}
			} else {
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH)
			}
		} else if (down) {
			for (var i = 0; i < segmentPoints.length; i++) {
				if (i % 2 === 0) {
					path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
				}
			}
			if (check) {
				for (var i = segmentPoints.length - 1; i >= 0; i--) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
					}
				}
			} else {
				path.lnTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)
			}
		} else {
			var endIndex = 0;
			var startIndex = segmentPoints.length - 1;

			if (!check) {
				path.moveTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH)
				for (var i = 0; i < segmentPoints2.length; i++) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

				path.moveTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)
				for (var i = 0; i < segmentPoints.length; i++) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)

			} else if (segmentPoints2.length === 1) {
				path.lnTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH);
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

				for (var k = endIndex; k <= startIndex; k++) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints[k].x / pxToMm * pathW, segmentPoints[k].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

			} else {

				path.moveTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH)
			
				for (var k = endIndex; k <= startIndex; k++) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints[k].x / pxToMm * pathW, segmentPoints[k].y / pxToMm * pathH);
					}
				}
		
				for (k = startIndex; endIndex <= k; k--) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints2[k].x / pxToMm * pathW, segmentPoints2[k].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH);
			}

		}

		return pathId;
	},
	
	calculatePolygon: function(array)
	{
		if(!array)
			return null;

		var pathId = this.cChartSpace.AllocPath();
		var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		
		for(var i = 0; i < array.length; i++)
		{
			var point = array[i];
			if(i === 0)
			{
				path.moveTo(point.x * pathW, point.y * pathH);
			}
			else
			{
				path.lnTo(point.x * pathW, point.y * pathH);
			}
			
			if(i === array.length - 1)
			{
				path.lnTo(array[0].x * pathW, array[0].y * pathH);
			}
		}
		

		return pathId;
	},

	_isSwitchCurrent3DChart: function (chartSpace) {
		var res = false;

		var chart = chartSpace && chartSpace.chart ? chartSpace.chart.plotArea.charts[0] : null;
		var typeChart = chart ? chart.getObjectType() : null;
		var oView3D = chartSpace && chartSpace.chart && chartSpace.chart.getView3d();
		if (isTurnOn3DCharts && oView3D) {
			var isPerspective = !oView3D.getRAngAx();

			var isBar = typeChart === AscDFH.historyitem_type_BarChart && chart && chart.barDir !== AscFormat.BAR_DIR_BAR;
			var isHBar = typeChart === AscDFH.historyitem_type_BarChart && chart && chart.barDir === AscFormat.BAR_DIR_BAR;
			var isLine = typeChart === AscDFH.historyitem_type_LineChart;
			var isPie = typeChart === AscDFH.historyitem_type_PieChart;
			var isArea = typeChart === AscDFH.historyitem_type_AreaChart;
			var isSurface = typeChart === AscDFH.historyitem_type_SurfaceChart;

			if (!isPerspective && (isBar || isLine || isHBar || isPie || isArea || isSurface)) {
				res = true;
			} else if (isPerspective && (isBar || isLine || isHBar || isArea || isPie || isSurface)) {
				res = true;
			}
		}

		return res;
	},

	getNumCache: function (val) {
		var res = null;

		if (val) {
			res = val.getNumCache();
		}

		return res;
	},

	getHorizontalPoints: function (chartSpace) {
		var res = null;

		if(!chartSpace) {
			chartSpace = this.cChartSpace;
		}

		var plotArea = chartSpace.chart.plotArea;
		if (plotArea.valAx && plotArea.valAx.xPoints) {
			res = plotArea.valAx.xPoints;
		} else if (plotArea.catAx && plotArea.catAx.xPoints) {
			res = plotArea.catAx.xPoints;
		}

		return res;
	},

	getVerticalPoints: function (chartSpace) {
		var res = null;

		if(!chartSpace) {
			chartSpace = this.cChartSpace;
		}

		var plotArea = chartSpace.chart.plotArea;
		if (plotArea.valAx && plotArea.valAx.yPoints) {
			res = plotArea.valAx.yPoints;
		} else if (plotArea.catAx && plotArea.catAx.yPoints) {
			res = plotArea.catAx.yPoints;
		}

		return res;
	},

	getPlotAreaPoints: function (chartSpace) {
		if(!chartSpace) {
			chartSpace = this.cChartSpace;
		}

		var xPoints = this.getHorizontalPoints(chartSpace);
		var yPoints = this.getVerticalPoints(chartSpace);
		var pxToMm = this.calcProp.pxToMM;
		var plotArea = chartSpace.chart.plotArea;
		var left, right, bottom, top;
		var t = this;

		var defaultCalculate = function() {
			var widthGraph = t.calcProp.widthCanvas;
			var heightGraph = t.calcProp.heightCanvas;
			var leftMargin = t.calcProp.chartGutter._left;
			var rightMargin = t.calcProp.chartGutter._right;
			var topMargin = t.calcProp.chartGutter._top;
			var bottomMargin = t.calcProp.chartGutter._bottom;

			left = leftMargin / pxToMm;
			right = (widthGraph - rightMargin) / pxToMm;
			bottom = (heightGraph - bottomMargin) / pxToMm;
			top = (topMargin) / pxToMm;
		};

		if (xPoints && yPoints) {
			var crossBetweenX = chartSpace.getValAxisCrossType();
			var crossDiffX = 0;
			if (crossBetweenX === AscFormat.CROSS_BETWEEN_BETWEEN && plotArea.valAx.posX && this.calcProp.type !== c_oChartTypes.HBar) {
				crossDiffX = xPoints[1] ? Math.abs((xPoints[1].pos - xPoints[0].pos) / 2) : Math.abs(xPoints[0].pos - plotArea.valAx.posX);
			}
			left = xPoints[0].pos - crossDiffX;
			right = xPoints[xPoints.length - 1].pos + crossDiffX;

			var crossBetweenY = chartSpace.getValAxisCrossType();
			var crossDiffY = 0;
			if (crossBetweenY === AscFormat.CROSS_BETWEEN_BETWEEN && plotArea.valAx.posY) {
				crossDiffY = yPoints[1] ? Math.abs((yPoints[1].pos - yPoints[0].pos) / 2) : Math.abs(yPoints[0].pos - plotArea.valAx.posY);
			}
			bottom = yPoints[0].pos + crossDiffY;
			top = yPoints[yPoints.length - 1].pos - crossDiffY;
		} else {
			defaultCalculate();
		}

		if(isNaN(left) || isNaN(top) || isNaN(bottom) || isNaN(right)) {
			defaultCalculate();
		}

		return {left: left, right: right, bottom: bottom, top: top};
	},

	//common functions for grid
	getHorizontalGridLines: function (axis, isCatAxis) {
		let t = this;
		let gridLines, minorGridLines;
		let crossBetween = this.cChartSpace.getValAxisCrossType();
		let points = axis.yPoints;

		if(!points) {
			return;
		}
		if(!axis.majorGridlines && !axis.minorGridlines) {
			return;
		}

		let widthLine = this.calcProp.trueWidth;
		let bottomMargin = this.calcProp.heightCanvas - this.calcProp.chartGutter._bottom;
		let posX = this.calcProp.chartGutter._left;
		let posMinorY, posY, crossDiff;

		if (crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN && isCatAxis) {
			let posAxis = (this.calcProp.heightCanvas - this.calcProp.chartGutter._bottom)/this.calcProp.pxToMM;
			crossDiff = points[1] ? Math.abs((points[1].pos - points[0].pos) / 2) : Math.abs(points[0].pos - posAxis);
		}

		let isRadarValues = axis.isRadarValues();
		
		let ptCount, tempAngle, trueHeight, trueWidth, xDiff, xCenter, yCenter;
		if(isRadarValues) {
			let radarChart = this.getChartByType(AscDFH.historyitem_type_RadarChart);
			let ptCount = radarChart && radarChart.series && this.getPtCount(radarChart.series);
			
			if(ptCount) {
				tempAngle = 2 * Math.PI / ptCount;
				trueHeight = this.calcProp.trueHeight;
				trueWidth = this.calcProp.trueWidth;
				xDiff = ((trueHeight / 2) / points.length) / this.calcProp.pxToMM;
				xCenter = (this.calcProp.chartGutter._left + trueWidth/2) / this.calcProp.pxToMM;
				yCenter = (this.calcProp.chartGutter._top + trueHeight/2) / this.calcProp.pxToMM;
			}
		}

		let calculateRadarGridLines = function () {
			let y, x, radius, xFirst, yFirst;

			for (let k = 0; k < ptCount; k++) {
				y = i * xDiff;
				x = xCenter;

				radius = y;

				y = yCenter - radius * Math.cos(k * tempAngle);
				x = x + radius * Math.sin(k * tempAngle);

				let pathH = t.calcProp.pathH;
				let pathW = t.calcProp.pathW;

				path.stroke = true;
				if (k === 0) {
					xFirst = x;
					yFirst = y;
					path.moveTo(x * pathW, y * pathH);
				} else {
					if (k === ptCount - 1) {
						path.lnTo(x * pathW, y * pathH);
						path.lnTo(xFirst * pathW, yFirst * pathH);
					} else {
						path.lnTo(x * pathW, y * pathH);
					}
				}
			}

			if (!gridLines) {
				gridLines = pathId;
			}
		};

		let minorLinesCount = isCatAxis ? 2 : 5;
		let stepY = points[1] ? Math.abs(points[1].pos - points[0].pos) : Math.abs(points[0].pos - axis.posY) * 2;
		let minorStep = (stepY / minorLinesCount) * this.calcProp.pxToMM;

		let pathId = t.cChartSpace.AllocPath();
		let path = t.cChartSpace.GetPath(pathId);
		let i;
		for (i = 0; i < points.length; i++) {
			if(isRadarValues) {
				if(ptCount) {
					calculateRadarGridLines();
				}
			} else {
				if(isCatAxis && points[i].val < 0) {
					continue;
				}

				if (crossDiff) {
					posY = (points[i].pos + crossDiff) * this.calcProp.pxToMM;
				} else {
					posY = points[i].pos * this.calcProp.pxToMM;
				}

				if (!gridLines) {
					gridLines = pathId;
				}

				this._calculateGridLine(posX, posY, posX + widthLine, posY, path);

				if (crossDiff && i === points.length - 1) {
					if (crossDiff) {
						posY = (points[i].pos - crossDiff) * this.calcProp.pxToMM;
					}

					i++;
					if (!gridLines) {
						gridLines = pathId;
					}
					this._calculateGridLine(posX, posY, posX + widthLine, posY, path);
				}
			}
		}

		let pathIdMinor = t.cChartSpace.AllocPath();
		let pathMinor = t.cChartSpace.GetPath(pathIdMinor);
		for (i = 0; i < points.length; i++) {
			if(!isRadarValues) {
				if(isCatAxis && points[i].val < 0) {
					continue;
				}

				if (crossDiff) {
					posY = (points[i].pos + crossDiff) * this.calcProp.pxToMM;
				} else {
					posY = points[i].pos * this.calcProp.pxToMM;
				}

				//промежуточные линии
				for (let n = 0; n < minorLinesCount; n++) {
					posMinorY = posY + n * minorStep;

					if (posMinorY < this.calcProp.chartGutter._top || posMinorY > bottomMargin) {
						break;
					}

					if (!minorGridLines) {
						minorGridLines = pathIdMinor;
					}

					this._calculateGridLine(posX, posMinorY, posX + widthLine, posMinorY, pathMinor);
				}
			}
		}

		return {gridLines: gridLines, minorGridLines: minorGridLines};
	},

	getVerticalGridLines: function (axis, isCatAxis) {
		var gridLines, minorGridLines;
		const isChartEx = this.cChartSpace.isChartEx();

		var crossBetween = this.cChartSpace.getValAxisCrossType();
		if(null === crossBetween && isCatAxis) {
			crossBetween = axis.crossAx ? axis.crossAx.crossBetween : null;
		}

		if(null === crossBetween && isChartEx) {
			crossBetween = 0;
		}

		var heightLine = this.calcProp.trueHeight;
		var rightMargin = this.calcProp.widthCanvas - this.calcProp.chartGutter._right;
		var posY = this.calcProp.chartGutter._top;
		var posMinorX;
		var points = axis.xPoints;

		if (!points) {
			return;
		}
		if(!axis.majorGridlines && !axis.minorGridlines) {
			return;
		}
		var minorLinesCount = isCatAxis || isChartEx ? 2 : 5;

		var posAxis = this.calcProp.chartGutter._left / this.calcProp.pxToMM;
		var stepX = points[1] ? Math.abs((points[1].pos - points[0].pos)) : (Math.abs(points[0].pos - posAxis) * 2);
		var minorStep = (stepX * this.calcProp.pxToMM) / minorLinesCount;
		var posX, crossDiff;

		if (crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN && (isCatAxis || isChartEx)) {
			crossDiff = points[1] ? Math.abs((points[1].pos - points[0].pos) / 2) : Math.abs(points[0].pos - posAxis);
		}

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);
		var i;
		for (i = 0; i < points.length; i++) {
			if((isCatAxis && points[i].val < 0) && !isChartEx) {
				continue;
			}

			if (crossDiff) {
				posX = (points[i].pos - crossDiff) * this.calcProp.pxToMM;
			} else {
				posX = points[i].pos * this.calcProp.pxToMM;
			}

			if (!gridLines) {
				gridLines = pathId;
			}
			this._calculateGridLine(posX, posY, posX, posY + heightLine, path);

			if (crossDiff && i === points.length - 1) {
				if (crossDiff) {
					posX = (points[i].pos + crossDiff) * this.calcProp.pxToMM;
				}

				i++;
				if (!gridLines) {
					gridLines = pathId;
				}
				this._calculateGridLine(posX, posY, posX, posY + heightLine, path);
			}
		}

		var pathIdMinor = this.cChartSpace.AllocPath();
		var pathMinor = this.cChartSpace.GetPath(pathIdMinor);
		for (i = 0; i < points.length; i++) {
			if((isCatAxis && points[i].val < 0) && !isChartEx) {
				continue;
			}

			if (crossDiff) {
				posX = (points[i].pos - crossDiff) * this.calcProp.pxToMM;
			} else {
				posX = points[i].pos * this.calcProp.pxToMM;
			}

			//промежуточные линии
			for (var n = 0; n <= minorLinesCount; n++) {
				posMinorX = posX + n * minorStep;

				if (posMinorX < this.calcProp.chartGutter._left || posMinorX > rightMargin) {
					break;
				}

				if (!minorGridLines) {
					minorGridLines = pathIdMinor;
				}

				this._calculateGridLine(posMinorX, posY, posMinorX, posY + heightLine, pathMinor);
			}
		}

		return {gridLines: gridLines, minorGridLines: minorGridLines};
	},

	_calculateGridLine: function (x, y, x1, y1, path) {
		var t = this;

		var calculate2DLine = function (x, y, x1, y1) {

			if(!path) {
				var pathId = t.cChartSpace.AllocPath();
				path = t.cChartSpace.GetPath(pathId);
			}

			var pathH = t.calcProp.pathH;
			var pathW = t.calcProp.pathW;

			path.stroke = true;
			var pxToMm = t.calcProp.pxToMM;
			path.moveTo(x / pxToMm * pathW, y / pxToMm * pathH);
			path.lnTo(x1 / pxToMm * pathW, y1 / pxToMm * pathH);

			return pathId;
		};

		var calculate3DLine = function (x, y, x1, y1, x2, y2, x3, y3) {

			if(!path) {
				var pathId = t.cChartSpace.AllocPath();
				path = t.cChartSpace.GetPath(pathId);
			}

			var pathH = t.calcProp.pathH;
			var pathW = t.calcProp.pathW;

			path.stroke = true;
			var pxToMm = t.calcProp.pxToMM;
			path.moveTo(x / pxToMm * pathW, y / pxToMm * pathH);
			path.lnTo(x1 / pxToMm * pathW, y1 / pxToMm * pathH);
			path.lnTo(x2 / pxToMm * pathW, y2 / pxToMm * pathH);
			if (x3 !== undefined && y3 !== undefined) {
				path.lnTo(x3 / pxToMm * pathW, y3 / pxToMm * pathH);
				path.lnTo(x / pxToMm * pathW, y / pxToMm * pathH);
			}

			return pathId;
		};

		if (this.nDimensionCount === 3) {
			var view3DProp = this.cChartSpace.chart.getView3d();
			var angleOx = view3DProp && view3DProp.rotX ? (-view3DProp.rotX / 360) * (Math.PI * 2) : 0;
			var angleOy = view3DProp && view3DProp.rotY ? (-view3DProp.rotY / 360) * (Math.PI * 2) : 0;
			var perspectiveDepth = this.processor3D.depthPerspective;
			var angleOz = 0;

			var rAngAx = this.processor3D.view3D.getRAngAx();
			var isVertLine = x === x1;

			var convertResult = this._convertAndTurnPoint(x, y, 0);
			var x1n = convertResult.x;
			var y1n = convertResult.y;
			convertResult = this._convertAndTurnPoint(x, y, perspectiveDepth);
			var x2n = convertResult.x;
			var y2n = convertResult.y;
			convertResult = this._convertAndTurnPoint(x1, y1, perspectiveDepth);
			var x3n = convertResult.x;
			var y3n = convertResult.y;
			convertResult = this._convertAndTurnPoint(x1, y1, 0);
			var x4n = convertResult.x;
			var y4n = convertResult.y;

			if (!isVertLine) {
				if (rAngAx) {
					path = calculate3DLine(x1n, y1n, x2n, y2n, x3n, y3n);
				} else {
					var angleOyAbs = Math.abs(angleOy);
					if (angleOyAbs >= 0 && angleOyAbs < Math.PI / 2) {
						path = calculate3DLine(x1n, y1n, x2n, y2n, x3n, y3n);
					} else if (angleOyAbs >= Math.PI / 2 && angleOyAbs < Math.PI) {
						path = calculate3DLine(x4n, y4n, x1n, y1n, x2n, y2n);
					} else if (angleOyAbs >= Math.PI && angleOyAbs < 3 * Math.PI / 2) {
						path = calculate3DLine(x1n, y1n, x4n, y4n, x3n, y3n);
					} else {
						path = calculate3DLine(x2n, y2n, x3n, y3n, x4n, y4n);
					}
				}
			} else {
				if (rAngAx) {
					path = calculate3DLine(x2n, y2n, x3n, y3n, x4n, y4n);
				} else {
					if (this.calcProp.type === c_oChartTypes.HBar) {
						var angleOyAbs = Math.abs(angleOy);
						if (angleOyAbs >= 3 * Math.PI / 2 || angleOyAbs < Math.PI / 2) {
							path = calculate3DLine(x2n, y2n, x3n, y3n, x4n, y4n);
						} else if (angleOyAbs >= Math.PI / 2 && angleOyAbs < 3 * Math.PI / 2) {
							path = calculate3DLine(x1n, y1n, x4n, y4n, x3n, y3n);
						}
					} else {
						path = calculate3DLine(x2n, y2n, x3n, y3n, x4n, y4n);
					}
				}
			}

		} else {
			path = calculate2DLine(x, y, x1, y1);
		}

		return path;
	},

	_getRadarGridLines: function (axis) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);
		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		if (axis.getObjectType() === AscDFH.historyitem_type_ValAx) {
			var yPoints = axis.yPoints;
			var orientation = !axis.isReversed();
			var trueWidth = this.calcProp.trueWidth;
			var xCenter = axis.posX;

			var yCenter, trueHeight;
			if (yPoints.length > 0) {
				yCenter = orientation ? yPoints[0].pos : yPoints[yPoints.length - 1].pos;
			} else {
				trueHeight = this.calcProp.trueHeight;
				yCenter = (this.calcProp.chartGutter._top + trueHeight / 2) / this.calcProp.pxToMM;
			}

			let radarChart = this.getChartByType(AscDFH.historyitem_type_RadarChart);
			let ptCount = radarChart && radarChart.series && this.getPtCount(radarChart.series);

			if (!ptCount) {
				return null;
			}
			var tempAngle = 2 * Math.PI / ptCount;

			var radius, x, y, firstX, firstY;
			for (var i = 0; i < yPoints.length; i++) {
				radius = yPoints[0].pos - yPoints[yPoints.length - (i + 1)].pos;
				for (var n = 0; n < ptCount; n++) {

					x = xCenter + radius * Math.sin(n * tempAngle);
					if (orientation) {
						y = yCenter - radius * Math.cos(n * tempAngle);
					} else {
						y = yCenter + radius * Math.cos(n * tempAngle);
					}

					if (n === 0) {
						path.moveTo(x * pathW, y * pathH);
						firstX = x;
						firstY = y;
					} else {
						path.lnTo(x * pathW, y * pathH);
					}
				}
				path.lnTo(firstX * pathW, firstY * pathH);
			}
		}
		return { gridLines: pathId };
	},

	getAxisFromAxId: function(axId, type) {
		var res = null;
		for(var i = 0; i < axId.length; i++) {
			if(axId[i].getObjectType() === type) {
				res = this._searchChangedAxis(axId[i]);
				break;
			}
		}
		return res;
	},

	getAxisByPos: function(axId, isHor) {
		var res = null;
		for(var i = 0; i < axId.length; i++) {
			if((!isHor && (axId[i].axPos === window['AscFormat'].AX_POS_L || axId[i].axPos === window['AscFormat'].AX_POS_R)) || (isHor && ((axId[i].axPos === window['AscFormat'].AX_POS_T || axId[i].axPos === window['AscFormat'].AX_POS_B)))) {
				res = this._searchChangedAxis(axId[i]);
				break;
			}
		}
		return res;
	},

	getPositionZero: function(axis) {
		var res = null;

		var points = axis.xPoints ? axis.xPoints : axis.yPoints;
		if(points) {
			for(var i = 0; i < points.length; i++) {
				if(0 === points[i].val) {
					res = points[i].pos * this.calcProp.pxToMM;
					break;
				}
			}

			if(null === res) {
				if (points[0].val < 0 && points[points.length - 1].val > 0) {
					res = axis.nullPos * this.calcProp.pxToMM;
				} else if (points[0].val < 0) {
					res = points[points.length - 1].pos * this.calcProp.pxToMM;
				} else {
					res = points[0].pos * this.calcProp.pxToMM;
				}
			}
		}

		return res;
	},

	calculateSplineLine: function (x, y, x1, y1, x2, y2, x3, y3, catAx, valAx) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		var startCoords;
		var endCoords;

		for (var i = 0; i <= 1;) {
			var splineCoords = this.calculate_Bezier(x, y, x1, y1, x2, y2, x3, y3, i);

			if (i === 0) {
				startCoords = {
					x: this.getYPosition(splineCoords[0], catAx),
					y: this.getYPosition(splineCoords[1], valAx)
				};
			}

			endCoords = {
				x: this.getYPosition(splineCoords[0], catAx),
				y: this.getYPosition(splineCoords[1], valAx)
			};

			if (i === 0) {
				path.moveTo(startCoords.x * pathW, startCoords.y * pathH);
			}
			path.lnTo(endCoords.x * pathW, endCoords.y * pathH);

			i = i + 0.1;
		}

		return pathId;
	},

	calculatePyramid: function (hBar, subType, startX, startY, height, gapDepth, individualBarValue, perspectiveDepth, 
		val, nullPositionOX, maxH, minH) {
		//todo оптимизировать под вертикальные bar
		var x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, x5, y5, z5, x6, y6, z6, x7, y7, z7, x8, y8, z8;
		var point1, point2, point3, point4, point5, point6, point7, point8;
		var value = height;
		var points2;
		perspectiveDepth = perspectiveDepth === 0 ? individualBarValue : perspectiveDepth;

		var downPlainPointX1, downPlainPointX2, downPlainPointX3, downPlainPointX4, upPlainPointX1, upPlainPointX2, upPlainPointX3, upPlainPointX4;
		var downPlainPointY1, downPlainPointY2, downPlainPointY3, downPlainPointY4, upPlainPointY1, upPlainPointY2, upPlainPointY3, upPlainPointY4;
		var downPlainPointZ1, downPlainPointZ2, downPlainPointZ3, downPlainPointZ4, upPlainPointZ1, upPlainPointZ2, upPlainPointZ3, upPlainPointZ4;

		var pointPlainUp1, pointPlainUp2, pointPlainUp3, pointPlainUp4;
		var pointPlainDown1, pointPlainDown2, pointPlainDown3, pointPlainDown4;
		var lineEquation1, lineEquation2, lineEquation3, lineEquation4;
		var plainEquationDown;
		var plainEquationUp;

		if (hBar) {
			downPlainPointX1 = startX, downPlainPointY1 = startY, downPlainPointZ1 = 0 + gapDepth;
			downPlainPointX2 = startX, downPlainPointY2 = startY, downPlainPointZ2 = perspectiveDepth + gapDepth;
			downPlainPointX3 = startX, downPlainPointY3 = startY + individualBarValue, downPlainPointZ3 = gapDepth;
			downPlainPointX4 = startX, downPlainPointY4 = startY + individualBarValue, downPlainPointZ4 = perspectiveDepth + gapDepth;

			upPlainPointX1 = startX + height, upPlainPointY1 = startY, upPlainPointZ1 = perspectiveDepth + gapDepth;
			upPlainPointX2 = startX + height, upPlainPointY2 = startY, upPlainPointZ2 = gapDepth;
			upPlainPointX3 = startX + height, upPlainPointY3 = startY + individualBarValue, upPlainPointZ3 = perspectiveDepth + gapDepth;
			upPlainPointX4 = startX + height, upPlainPointY4 = startY + individualBarValue, upPlainPointZ4 = gapDepth;
		} else {
			downPlainPointX1 = startX, downPlainPointY1 = startY, downPlainPointZ1 = 0 + gapDepth;
			downPlainPointX2 = startX, downPlainPointY2 = startY, downPlainPointZ2 = perspectiveDepth + gapDepth;
			downPlainPointX3 = startX + individualBarValue, downPlainPointY3 = startY, downPlainPointZ3 = perspectiveDepth + gapDepth;
			downPlainPointX4 = startX + individualBarValue, downPlainPointY4 = startY, downPlainPointZ4 = 0 + gapDepth;

			upPlainPointX1 = startX, upPlainPointY1 = startY - height, upPlainPointZ1 = 0 + gapDepth;
			upPlainPointX2 = startX, upPlainPointY2 = startY - height, upPlainPointZ2 = perspectiveDepth + gapDepth;
			upPlainPointX3 = startX + individualBarValue, upPlainPointY3 = startY - height, upPlainPointZ3 = perspectiveDepth + gapDepth;
			upPlainPointX4 = startX + individualBarValue, upPlainPointY4 = startY - height, upPlainPointZ4 = 0 + gapDepth;
		}

		if (subType === "stackedPer" || subType === "stacked") {
			value = maxH;
			if (val < 0 && minH) {
				value = minH;
			}		
		
			if (hBar) {
				x1 = nullPositionOX, y1 = startY, z1 = gapDepth;
				x2 = nullPositionOX, y2 = startY, z2 = perspectiveDepth + gapDepth;
				x5 = nullPositionOX, y5 = startY + individualBarValue, z5 = gapDepth;
				x6 = nullPositionOX, y6 = startY + individualBarValue, z6 = perspectiveDepth + gapDepth;
				x3 = nullPositionOX - value + 1, y3 = startY + individualBarValue / 2, z3 = perspectiveDepth / 2 + gapDepth;
				x4 = nullPositionOX - value + 1, y4 = startY + individualBarValue / 2, z4 = perspectiveDepth / 2 + gapDepth;
				x7 = nullPositionOX - value + 1, y7 = startY + individualBarValue / 2, z7 = perspectiveDepth / 2 + gapDepth;
				x8 = nullPositionOX - value + 1, y8 = startY + individualBarValue / 2, z8 = perspectiveDepth / 2 + gapDepth;
			} else {
				x1 = startX, y1 = nullPositionOX, z1 = 0 + gapDepth;
				x2 = startX, y2 = nullPositionOX, z2 = perspectiveDepth + gapDepth;
				x3 = startX + individualBarValue, y3 = nullPositionOX, z3 = perspectiveDepth + gapDepth;
				x4 = startX + individualBarValue, y4 = nullPositionOX, z4 = 0 + gapDepth;
				x5 = startX + individualBarValue / 2, y5 = nullPositionOX - value - 1, z5 = 0 + gapDepth + perspectiveDepth / 2;
				x6 = startX + individualBarValue / 2, y6 = nullPositionOX - value - 1, z6 = perspectiveDepth + gapDepth - perspectiveDepth / 2;
				x7 = startX + individualBarValue / 2, y7 = nullPositionOX - value - 1, z7 = perspectiveDepth + gapDepth - perspectiveDepth / 2;
				x8 = startX + individualBarValue / 2, y8 = nullPositionOX - value - 1, z8 = 0 + gapDepth + perspectiveDepth / 2;
			}

			//расчет верхней и нижней пересекающих плоскостей
			pointPlainDown1 = this._convertAndTurnPoint(downPlainPointX1, downPlainPointY1, downPlainPointZ1);
			pointPlainDown2 = this._convertAndTurnPoint(downPlainPointX2, downPlainPointY2, downPlainPointZ2);
			pointPlainDown3 = this._convertAndTurnPoint(downPlainPointX3, downPlainPointY3, downPlainPointZ3);
			pointPlainDown4 = this._convertAndTurnPoint(downPlainPointX4, downPlainPointY4, downPlainPointZ4);

			pointPlainUp1 = this._convertAndTurnPoint(upPlainPointX1, upPlainPointY1, upPlainPointZ1);
			pointPlainUp2 = this._convertAndTurnPoint(upPlainPointX2, upPlainPointY2, upPlainPointZ2);
			pointPlainUp3 = this._convertAndTurnPoint(upPlainPointX3, upPlainPointY3, upPlainPointZ3);
			pointPlainUp4 = this._convertAndTurnPoint(upPlainPointX4, upPlainPointY4, upPlainPointZ4);


			point1 = this._convertAndTurnPoint(x1, y1, z1);
			point2 = this._convertAndTurnPoint(x2, y2, z2);
			point3 = this._convertAndTurnPoint(x3, y3, z3);
			point4 = this._convertAndTurnPoint(x4, y4, z4);
			point5 = this._convertAndTurnPoint(x5, y5, z5);
			point6 = this._convertAndTurnPoint(x6, y6, z6);
			point7 = this._convertAndTurnPoint(x7, y7, z7);
			point8 = this._convertAndTurnPoint(x8, y8, z8);

			//расчет уравнений прямых для ребер пирамиды
			if (hBar) {
				lineEquation1 = this.getLineEquation(point1, point4);
				lineEquation2 = this.getLineEquation(point2, point3);
				lineEquation3 = this.getLineEquation(point6, point8);
				lineEquation4 = this.getLineEquation(point5, point7);
			} else {
				lineEquation1 = this.getLineEquation(point1, point5);
				lineEquation2 = this.getLineEquation(point4, point8);
				lineEquation3 = this.getLineEquation(point3, point7);
				lineEquation4 = this.getLineEquation(point2, point6);
			}

			//расчет уравнений плоскостей
			plainEquationDown =
				this.getPlainEquation(pointPlainDown1, pointPlainDown2, pointPlainDown3, pointPlainDown4);
			plainEquationUp = this.getPlainEquation(pointPlainUp1, pointPlainUp2, pointPlainUp3, pointPlainUp4);

			//находим точки пересечения плоскостей с пирамидой для расчета усечения

			if (hBar) {
				point1 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation1);
				point2 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation2);
				point5 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation4);
				point6 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation3);

				if (val !== 0) {
					point3 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation2);
					point4 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation1);
					point7 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation3);
					point8 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation4);
				} else {
					point3 = point1;
					point4 = point2;
					point7 = point5;
					point8 = point6;
				}
			} else {
				point1 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation1);
				point2 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation4);
				point3 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation3);
				point4 = this.isIntersectionPlainAndLine(plainEquationDown, lineEquation2);
				if (val !== 0) {
					point5 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation1);
					point6 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation4);
					point7 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation3);
					point8 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation2);
				} else {
					point1 = point5;
					point2 = point6;
					point3 = point7;
					point4 = point8;
				}
				if (this.cChartSpace.chart.plotArea.valAx.isReversed()) {
					x5 = startX, y5 = startY - height / 2, z5 = 0 + gapDepth;
					x6 = startX, y6 = startY - height / 2, z6 = perspectiveDepth + gapDepth;
					x7 = startX + individualBarValue, y7 = startY - height / 2, z7 = perspectiveDepth + gapDepth;
					x8 = startX + individualBarValue, y8 = startY - height / 2, z8 = 0 + gapDepth;
	
					pointPlainUp1 = this._convertAndTurnPoint(x5, y5, z5);
					pointPlainUp2 = this._convertAndTurnPoint(x6, y6, z6);
					pointPlainUp3 = this._convertAndTurnPoint(x7, y7, z7);
					pointPlainUp4 = this._convertAndTurnPoint(x8, y8, z8);
					plainEquationUp = this.getPlainEquation(pointPlainUp1, pointPlainUp2, pointPlainUp3, pointPlainUp4)
			
					var point12 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation1);
					var point22 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation4);
					var point32 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation3);
					var point42 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation2);
					points2 = [point12, point22, point32, point42, point5, point6, point7, point8];
				}
			}
		} else {

			if (maxH) {
				height = maxH;
			} 
			if (minH && val < 0) {
				height = minH;
			}

			if (hBar) {
				x1 = startX, y1 = startY, z1 = gapDepth;
				x2 = startX, y2 = startY, z2 = perspectiveDepth + gapDepth;
				x5 = startX, y5 = startY + individualBarValue, z5 = gapDepth;
				x6 = startX, y6 = startY + individualBarValue, z6 = perspectiveDepth + gapDepth;

				x3 = startX + height, y3 = startY + individualBarValue / 2, z3 = perspectiveDepth / 2 + gapDepth;
				x4 = startX + height, y4 = startY + individualBarValue / 2, z4 = perspectiveDepth / 2 + gapDepth;
				x7 = startX + height, y7 = startY + individualBarValue / 2, z7 = perspectiveDepth / 2 + gapDepth;
				x8 = startX + height, y8 = startY + individualBarValue / 2, z8 = perspectiveDepth / 2 + gapDepth;
			} else {
				x1 = startX, y1 = startY, z1 = 0 + gapDepth;
				x2 = startX, y2 = startY, z2 = perspectiveDepth + gapDepth;
				x3 = startX + individualBarValue, y3 = startY, z3 = perspectiveDepth + gapDepth;
				x4 = startX + individualBarValue, y4 = startY, z4 = 0 + gapDepth;

				x5 = startX + individualBarValue / 2, y5 = startY - height, z5 = 0 + gapDepth + perspectiveDepth / 2;
				x6 = startX + individualBarValue / 2, y6 = startY - height, z6 = perspectiveDepth + gapDepth - perspectiveDepth / 2;
				x7 = startX + individualBarValue / 2, y7 = startY - height, z7 = perspectiveDepth + gapDepth - perspectiveDepth / 2;
				x8 = startX + individualBarValue / 2, y8 = startY - height, z8 = 0 + gapDepth + perspectiveDepth / 2;
			}

			point1 = this._convertAndTurnPoint(x1, y1, z1);
			point2 = this._convertAndTurnPoint(x2, y2, z2);
			point3 = this._convertAndTurnPoint(x3, y3, z3);
			point4 = this._convertAndTurnPoint(x4, y4, z4);
			point5 = this._convertAndTurnPoint(x5, y5, z5);
			point6 = this._convertAndTurnPoint(x6, y6, z6);
			point7 = this._convertAndTurnPoint(x7, y7, z7);
			point8 = this._convertAndTurnPoint(x8, y8, z8);

			if (maxH || minH) {

				pointPlainUp1 = this._convertAndTurnPoint(upPlainPointX1, upPlainPointY1, upPlainPointZ1);
				pointPlainUp2 = this._convertAndTurnPoint(upPlainPointX2, upPlainPointY2, upPlainPointZ2);
				pointPlainUp3 = this._convertAndTurnPoint(upPlainPointX3, upPlainPointY3, upPlainPointZ3);
				pointPlainUp4 = this._convertAndTurnPoint(upPlainPointX4, upPlainPointY4, upPlainPointZ4);

				plainEquationUp = this.getPlainEquation(pointPlainUp1, pointPlainUp2, pointPlainUp3, pointPlainUp4);

				if (hBar) {
					lineEquation1 = this.getLineEquation(point1, point4);
					lineEquation2 = this.getLineEquation(point2, point3);
					lineEquation3 = this.getLineEquation(point6, point8);
					lineEquation4 = this.getLineEquation(point5, point7);
					point3 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation2);
					point4 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation1);
					point7 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation3);
					point8 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation4);
				} else {
					lineEquation1 = this.getLineEquation(point1, point5);
					lineEquation2 = this.getLineEquation(point4, point8);
					lineEquation3 = this.getLineEquation(point3, point7);
					lineEquation4 = this.getLineEquation(point2, point6);
					point5 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation1);
					point6 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation4);
					point7 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation3);
					point8 = this.isIntersectionPlainAndLine(plainEquationUp, lineEquation2);
				}
			}
		}

		var points = [point1, point2, point3, point4, point5, point6, point7, point8];
		var paths;
		if (!hBar) {
			paths = this.calculateRect3D(points, val, null, subType !== "standard", points2);
			return { paths: paths, points: points }
		} else {
			paths = this.calculateRect3D(points, val, null, true);
			return paths;
		}
	},

	isConeIntersection: function (hbar, subType, startX, startY, height, gapDepth, individualBarValue, perspectiveDepth,
		val, nullPositionOX, maxH, minH) {
		var wUp, lUp, wDown, lDown;
		var value, l1, l2, l3, l4;
		var pyramidX1, pyramidX2, pyramidY1, pyramidY2, rectX1, rectX2, rectY1, rectY2;
		if (subType === "stacked" || subType === "stackedPer") {
			value = maxH;
			if (val < 0 && minH) {
				value = minH;
			}
		} else {
			value = height;
			if (maxH) {
				value = hbar ? -maxH : maxH;
			} 
			if (minH && val < 0) {
				value = hbar ? -minH: minH;
			}
		}
		if (this.cChartSpace.chart.plotArea.valAx.isReversed()) {
			if (subType === "stacked" || subType === "stackedPer") {
				if (hbar) {
					startX = val < 0 ? --startX : ++startX;
				} else {
					startY = val < 0 ? ++startY : --startY;
				}
			}
		}

		// рассчитываем большую и малую полуось оснований усеченного конуса через нахождение точек пересечения линий пирамиды и параллелепипеда
		if (hbar) {
			pyramidX1 = startY + individualBarValue / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = startY, pyramidY2 = nullPositionOX;
			rectX1 = startY, rectY1 = startX + height;
			rectX2 = startY + individualBarValue, rectY2 = startX + height;

			l1 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = startY + individualBarValue;
			l2 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);

			pyramidX1 = gapDepth + perspectiveDepth / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = gapDepth, pyramidY2 = nullPositionOX;
			rectX1 = gapDepth, rectY1 = startX + height;
			rectX2 = gapDepth + perspectiveDepth, rectY2 = startX + height;

			l3 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = gapDepth + perspectiveDepth;
			l4 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
		
			wUp = l1.x && l2.x ? (l2.x - l1.x) / 2 : 0;
			lUp = l3.x && l4.x ? (l4.x - l3.x) / 2 : 0;

			pyramidX1 = startY + individualBarValue / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = startY, pyramidY2 = nullPositionOX;
			rectX1 = startY + individualBarValue, rectY1 = startX;
			rectX2 = startY, rectY2 = startX;

			l1 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = startY + individualBarValue;
			l2 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);

			pyramidX1 = gapDepth + perspectiveDepth / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = gapDepth, pyramidY2 = nullPositionOX;
			rectX1 = gapDepth, rectY1 = startX;
			rectX2 = gapDepth + perspectiveDepth, rectY2 = startX;

			l3 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = gapDepth + perspectiveDepth;
			l4 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
		
			wDown = l1.x && l2.x ? (l2.x - l1.x) / 2 : 0;
			lDown = l3.x && l4.x ? (l4.x - l3.x) / 2 : 0;
		} else {
			// координаты точек ребра пирамиды в плоскости x, y
			pyramidX1 = startX + individualBarValue / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = startX, pyramidY2 = nullPositionOX;
			// координаты точек стороны пересекающего прямоугольника
			rectX1 = startX, rectY1 = startY - height;
			rectX2 = startX + individualBarValue, rectY2 = startY - height;

			l1 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = startX + individualBarValue;
			pyramidY2 = nullPositionOX;
			l2 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);

			pyramidX1 = gapDepth + perspectiveDepth / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = gapDepth, pyramidY2 = nullPositionOX;
			rectX1 = gapDepth, rectY1 = startY - height;
			rectX2 = gapDepth + perspectiveDepth, rectY2 = startY - height;

			l3 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = gapDepth + perspectiveDepth;
			pyramidY2 = nullPositionOX;
			l4 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
		
			wUp = l1.x && l2.x ? (l2.x - l1.x) / 2 : 0;
			lUp = l3.x && l4.x ? (l4.x - l3.x) / 2 : 0;

			pyramidX1 = startX + individualBarValue / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = startX, pyramidY2 = nullPositionOX;
			rectX1 = startX, rectY1 = startY;
			rectX2 = startX + individualBarValue, rectY2 = startY;

			l1 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = startX + individualBarValue;
			pyramidY2 = nullPositionOX;
			l2 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);

			pyramidX1 = gapDepth + perspectiveDepth / 2, pyramidY1 = nullPositionOX - value;
			pyramidX2 = gapDepth, pyramidY2 = nullPositionOX;
			rectX1 = gapDepth, rectY1 = startY;
			rectX2 = gapDepth + perspectiveDepth, rectY2 = startY;

			l3 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
			pyramidX2 = gapDepth + perspectiveDepth;
			pyramidY2 = nullPositionOX;
			l4 = this.isIntersectionLineAndLine2d(pyramidX1, pyramidY1, pyramidX2, pyramidY2, rectX1, rectY1, rectX2, rectY2);
		
			wDown = l1.x && l2.x ? (l2.x - l1.x) / 2 : 0;
			lDown = l3.x && l4.x ? (l4.x - l3.x) / 2 : 0;
		}

		return { wUp: wUp, lUp: lUp, wDown: wDown, lDown: lDown };		
	},

	getStartStackedPyramidPosition: function (val, chart, valAx, j, chartProp, summBarVal, maxVal, nullPositionOX1, subType, 
		ptCount, axisMax, axisMin, testMaxHeight, testMinHeight, type, hBar) {

		var endBlockPositionMax, endBlockPositionMin, maxH, minH, h;
		var maxPer, indexMax, valueMax;
		var minPer, indexMin, valueMin;

		if (subType === "stacked") {

			if (type === AscFormat.BAR_SHAPE_PYRAMIDTOMAX || type === AscFormat.BAR_SHAPE_CONETOMAX) {
				h = this.getStackedMaxHeight(ptCount, chart, subType, j).heightPer;
				valueMax = h.maxH;
				valueMin = h.minH;

				minH = nullPositionOX1 - this.getYPosition(valueMin, valAx, null, true) * chartProp.pxToMM;
				maxH = nullPositionOX1 - this.getYPosition(valueMax, valAx, null, true) * chartProp.pxToMM;

				if (valueMax > axisMax) {
					maxH = -testMaxHeight / (axisMax / valueMax);
				} 
				if (valueMin < axisMin) {
					minH = -testMinHeight / (axisMin / valueMin);
				}

			} else if (type === AscFormat.BAR_SHAPE_PYRAMID || type === AscFormat.BAR_SHAPE_CONE) {
				h = this.getStackedMaxHeight(ptCount, chart, null, j).height;
				valueMax = h.valueMax;
				valueMin = h.valueMin;

				endBlockPositionMax = this.getYPosition(maxVal, valAx, null, true) * chartProp.pxToMM;
				maxH = nullPositionOX1 - endBlockPositionMax;
				minH = nullPositionOX1 - endBlockPositionMax;

				if (valueMax > axisMax) {
					maxH = maxH / (axisMax / valueMax);
				}

				if (valueMin < axisMin) {
					minH = maxH / (axisMin / valueMin);
				}
			}
			
		} else if (subType === "stackedPer") {
			h = this.getStackedMaxHeight(ptCount, chart, null, j).heightPer;
			//получаем пропорционально наибольшее значение высоты и его индекс
			indexMax = h.indexMax;
			indexMin = h.indexMin;
			minH = h.minH;
			maxH = h.maxH;
			maxPer = h.maxPer;
			minPer = h.minPer;
			
			//значения для проверки условия усечения
			valueMax = maxPer;
			valueMin = minPer;
			//по индексу определяем когда получаем наибольшую высоту для расчета усечения
			//делится на this.summBarVal, чтобы получить пропорциональную величину для высоты
			if (indexMax === j) {
				endBlockPositionMax = this.getYPosition(maxH / summBarVal[j], valAx, null, true) * chartProp.pxToMM;
			} else {
				endBlockPositionMax = this.getYPosition(maxPer, valAx, null, true) * chartProp.pxToMM;
			}
		
			if (indexMin === j) {
				endBlockPositionMin = this.getYPosition(minH / summBarVal[j], valAx, null, true) * chartProp.pxToMM;
			} else {
				endBlockPositionMin = this.getYPosition(minPer, valAx, null, true) * chartProp.pxToMM;
			}
		
			maxH = nullPositionOX1 - endBlockPositionMax;
			minH = nullPositionOX1 - endBlockPositionMin;

			if (valueMax > axisMax) {
				maxH = maxH + (maxH - maxH * axisMax);
			}
			if (valueMin < axisMin) {
				minH = minH + (minH + minH * axisMin);
			}

		} else {
			h = this.getStackedMaxHeight(ptCount, chart, subType, j).heightPer;
			valueMax = h.maxH;
			valueMin = h.minH;

			if (type === AscFormat.BAR_SHAPE_PYRAMIDTOMAX || type === AscFormat.BAR_SHAPE_CONETOMAX) {
				maxH = hBar === true ? testMaxHeight: -testMaxHeight;
				minH = hBar === true ? testMinHeight: -testMinHeight;
			}

			if (val > axisMax) {
				maxH = hBar === true ? testMaxHeight / (axisMax / val) : -testMaxHeight / (axisMax / val);
			}
			if (val < axisMin) {
				minH = hBar === true ? testMinHeight / (axisMin / val) : -testMinHeight / (axisMin / val);
			}

			if (type === AscFormat.BAR_SHAPE_PYRAMIDTOMAX || type === AscFormat.BAR_SHAPE_CONETOMAX) {
				maxH = hBar === true ? maxH / (axisMax / valueMax) : -testMaxHeight / (axisMax / valueMax);
				minH = hBar === true ? minH / (axisMin / valueMin) : -testMinHeight / (axisMin / valueMin);
			}
		}

		return {maxH: maxH, minH: minH};
	},

	getStackedMaxHeight: function (ptCount, chart, subType, j) {

		var cChartDrawer = this;
		var calculateSummStacked = function (j) {
			var curVal;
			var temp = 0;
			var idxPoint;
			for (var k = 0; k < chart.series.length; k++) {
				idxPoint = cChartDrawer.getIdxPoint(chart.series[k], j);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;
	
				if (curVal) {
					temp += Math.abs(curVal);
				}
			}
	
			return temp;
		};

		var getMaxHeightStacked = function (ptCount, chart, cChartDrawer, subType) {
			var curVal;
			var tempMax = 0;
			var tempMin = 0;
			var idxPoint;
			var maxH1 = [];
			var minH1 = [];
			var maxH2 = [];
			var minH2 = [];
			var countMax = 0;
			var countMin = 0;
	
			//нахождение пропорционально наибольшей и наименьшей высоты для накопительных процентных пирамид
			for (var i = 0; i < ptCount; i++) {
				var sumBar = calculateSummStacked(i);
				for (var k = 0; k < chart.series.length; k++) {
					idxPoint = cChartDrawer.getIdxPoint(chart.series[k], i);
					curVal = idxPoint ? parseFloat(idxPoint.val) : 0;
	
					if (curVal > 0) {
						tempMax += curVal;
					} else {
						tempMin += curVal;
					}

					if (subType === "normal" || subType === "standard") {
						if (countMax < curVal) {
							countMax = curVal;
						}
						if (countMin > curVal) {
							countMin = curVal;
						}
					}
				}

				//пропорциональная высота
				maxH1[i] = tempMax / sumBar;
				minH1[i] = tempMin / sumBar;
	
				//обычная высота по значениям
				maxH2[i] = tempMax;
				minH2[i] = tempMin;

				if (subType === "stacked") {
					if (countMax < tempMax) {
						countMax = tempMax;
					}
					if (countMin > tempMin) {
						countMin = tempMin;
					}
				}
	
				tempMax = 0;
				tempMin = 0;
			}
			if (subType) {
				return { maxH: countMax, minH: countMin };
			}

			var indexMax = 0;
			var indexMin = 0;

			for (var k = 0; k < maxH1.length; k++) {
				if (tempMax < maxH1[k]) {
					tempMax = maxH1[k];
					countMax = maxH2[k];
					indexMax = k;
				}
			}
	
			for (k = 0; k < minH1.length; k++) {
				if (tempMin > minH1[k]) {
					tempMin = minH1[k];
					countMin = minH2[k];
					indexMin = k;
				}
			}
	
			return {
				maxH: countMax,
				minH: countMin,
				indexMax: indexMax,
				indexMin: indexMin,
				maxPer: tempMax,
				minPer: tempMin
			};
		};
	
		var getMaxHeight = function (j, cChartDrawer) {
			var curVal;
			var tempMax = 0;
			var tempMin = 0;
			var idxPoint;
			for (var k = 0; k < chart.series.length; k++) {
				idxPoint = cChartDrawer.getIdxPoint(chart.series[k], j);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;
			
				if (curVal > 0) {
					tempMax += curVal;
				} else {
					tempMin += curVal;
				}
			}
			return { valueMax: tempMax, valueMin: tempMin };
		};

		return {
				heightPer: getMaxHeightStacked(ptCount, chart, cChartDrawer, subType),
				height: getMaxHeight(j, cChartDrawer) 
			};
	},

	_calculateCylinder: function (startX, startY, individualBarWidth, height, val, gapDepth, perspectiveDepth,
		checkPathMethod, hbar, subType, nullPositionOX, maxH, minH) {
		var centerUpX, centerUpY, centerUpZ, centerDownX, centerDownY, centerDownZ;
		var points;
		var sizes1, sizes2, cone, sizes12, sizes22;

		if (hbar) {
			centerDownX = startX, centerDownY = startY + individualBarWidth / 2, centerDownZ = 0 + gapDepth + perspectiveDepth / 2;
			centerUpX = startX + height, centerUpY = startY + individualBarWidth / 2, centerUpZ = 0 + gapDepth + perspectiveDepth / 2;
		} else {
			centerDownX = startX + individualBarWidth / 2, centerDownY = startY, centerDownZ = 0 + gapDepth + perspectiveDepth / 2;
			centerUpX = startX + individualBarWidth / 2, centerUpY = startY - height, centerUpZ = 0 + gapDepth + perspectiveDepth / 2;
		}

		if (subType) {
			cone = true;
			// за оси эллипса берем 1/2 длин ребер оснований усеченной пирамиды
			points = this.isConeIntersection(hbar, subType, startX, startY, height, gapDepth, individualBarWidth, perspectiveDepth,
				val, nullPositionOX, maxH, minH);

			if ((subType === "stacked" || subType === "stackedPer") && this.cChartSpace.chart.plotArea.valAx.isReversed()) {
				sizes12 = points.wUp !== 0 ? points.wUp : individualBarWidth / 2;
				sizes22 = points.lUp !== 0 ? points.lUp : perspectiveDepth / 2;
				sizes1 = points.wDown;
				sizes2 = points.lDown;
			} else {
				sizes12 = points.wUp;
				sizes22 = points.lUp;
				sizes1 = points.wDown !== 0 ? points.wDown : individualBarWidth / 2;
				sizes2 = points.lDown !== 0 ? points.lDown : perspectiveDepth / 2;
			}

		} else {
			//большая и малая полуось эллипса
			sizes1 = individualBarWidth / 2;
			sizes2 = perspectiveDepth / 2;
		}

		var segmentPoints = [];
		var segmentPoints2 = [];
		var segmentPoint1, segmentPoint2;

		var A, B, A1, B1;
		// в ms 180 градусов составляют 17 сегментов
		// todo пока что рассчитываем дополнительные точки, нужно будет оптимизировать
		var dt = Math.PI / 34; //Math.PI / 17;

		//рассчитываем стартовый угол
		var angel = Math.abs(this.processor3D.angleOy);
		var k = Math.PI / 2;
		k += angel;

		// получаем точки основания цилиндра через парамметрические уравнения эллиптического цилиндра
		for (var t = k; t <= Math.PI * 2 + k; t += dt) {
			A = sizes1 * Math.cos(t);
			B = sizes2 * Math.sin(t);

			if (cone) {
				A1 = sizes12 * Math.cos(t);
				B1 = sizes22 * Math.sin(t);
				if (hbar) {
					segmentPoint1 = this._convertAndTurnPoint(centerDownX, centerDownY + A, centerDownZ - B);
					segmentPoint2 = this._convertAndTurnPoint(centerUpX, centerUpY + A1, centerUpZ - B1);
				} else {
					segmentPoint1 = this._convertAndTurnPoint(centerDownX + A, centerDownY, centerDownZ + B);
					segmentPoint2 = this._convertAndTurnPoint(centerUpX + A1, centerUpY, centerUpZ + B1);		
				}
			} else {
				if (hbar) {
					segmentPoint1 = this._convertAndTurnPoint(centerDownX, centerDownY + A, centerDownZ - B);
					segmentPoint2 = this._convertAndTurnPoint(centerUpX, centerUpY + A, centerUpZ - B);
				} else {
					segmentPoint1 = this._convertAndTurnPoint(centerDownX + A, centerDownY, centerDownZ + B);
					segmentPoint2 = this._convertAndTurnPoint(centerUpX + A, centerUpY, centerUpZ + B);
				}
			}

			segmentPoints.push(segmentPoint1);
			segmentPoints2.push(segmentPoint2);	
		}

		var sortCylinderPoints1 = [];
		var sortCylinderPoints2 = [];
		var invisible = false;

		// сортируем точки по видимости для построения плоскости цилиндра
		for (var i = 1; i < segmentPoints.length; i++) {
			if (this._isVisibleVerge3D(segmentPoints[i], segmentPoints[i - 1], segmentPoints2[i - 1], val, true)) {
				if (!invisible) {
					sortCylinderPoints1.push(segmentPoints[i - 1]);
					sortCylinderPoints2.push(segmentPoints2[i - 1]);
				} else {
					break;
				}
			} else {
				invisible = true;
			}
		}		
		if (invisible) {
			for (var k = segmentPoints.length - 1; i <= k; k--) {
				if (this._isVisibleVerge3D(segmentPoints[k], segmentPoints[k - 1], segmentPoints2[k - 1], val, true)) {
					sortCylinderPoints1.unshift(segmentPoints[k - 1]);
					sortCylinderPoints2.unshift(segmentPoints2[k - 1]);
				}
			}
		}
		var isNotAllPointsVisible = invisible;

		// проверяем если все точки поверхности цилиндра(конуса) либо видимы либо невидимы
		// если уловие выполняется, то для отрисовки цилиндра(конуса) достаточно отрисовать эллипс (т.е вид сверху или снизу)
		if (sortCylinderPoints1.length === 0 || sortCylinderPoints2.length === 0 || val === 0) {
			sortCylinderPoints1 = segmentPoints;
			sortCylinderPoints2 = segmentPoints2;
			isNotAllPointsVisible = false;
		}

		var x12, y12, z12, x22, y22, z22, x32, y32, z32, x42, y42, z42, x52, y52, z52, x62, y62, z62, x72, y72, z72, x82, y82, z82;
		var point1, point2, point4, point5, point6, point8;

		if (hbar) {
			x12 = startX, y12 = startY, z12 = 0 + gapDepth;
			x42 = startX + height, y42 = startY, z42 = perspectiveDepth + gapDepth;
			x52 = startX, y52 = startY + individualBarWidth, z52 = 0 + gapDepth;
			x62 = startX, y62 = startY + individualBarWidth, z62 = perspectiveDepth + gapDepth;
			x72 = startX + height, y72 = startY + individualBarWidth, z72 =  perspectiveDepth + gapDepth;
			x82 = startX + height, y82 = startY + individualBarWidth, z82 = 0 + gapDepth;

			point1 = this._convertAndTurnPoint(x12, y12, z12);
			point2 = this._convertAndTurnPoint(x52, y52, z52);
			point4 = this._convertAndTurnPoint(x62, y62, z62);
			point5 = this._convertAndTurnPoint(x72, y72, z72);
			point6 = this._convertAndTurnPoint(x42, y42, z42);
			point8 = this._convertAndTurnPoint(x82, y82, z82);
		} else {
			x12 = startX, y12 = startY, z12 = 0 + gapDepth;
			x22 = startX, y22 = startY, z22 = perspectiveDepth + gapDepth;
			x42 = startX + individualBarWidth, y42 = startY, z42 = 0 + gapDepth;
			x52 = startX, y52 = startY - height, z52 = 0 + gapDepth;
			x62 = startX, y62 = startY - height, z62 = perspectiveDepth + gapDepth;
			x82 = startX + individualBarWidth, y82 = startY - height, z82 = 0 + gapDepth;

			point1 = this._convertAndTurnPoint(x12, y12, z12);
			point2 = this._convertAndTurnPoint(x22, y22, z22);
			point4 = this._convertAndTurnPoint(x42, y42, z42);
			point5 = this._convertAndTurnPoint(x52, y52, z52);
			point6 = this._convertAndTurnPoint(x62, y62, z62);
			point8 = this._convertAndTurnPoint(x82, y82, z82);
		}

		var points = [segmentPoints, segmentPoints2, point1, point2, point4, point5, point6, point8,
			sortCylinderPoints1, sortCylinderPoints2];
		var paths = this.calculateCylinder(points, val, checkPathMethod, false, isNotAllPointsVisible, cone);

		return paths;
	},

	checkingPenForDrawing: function (shape, verge) {
		var isCone = shape === AscFormat.BAR_SHAPE_CONE || shape === AscFormat.BAR_SHAPE_CONETOMAX;
		return isCone && (verge === 0 || verge === 5);
	},

	getAutoSum: function (startM, endM, startN, endN, calcFunc) {
		var res = 0;
		for (var i = startM; i <= endM; i++) {
			for (var j = startN; j <= endN; j++) {
				res += calcFunc(i, j);
			}
		}
		return res;
	},

	getValWithStacked: function (ser, val, oChart) {
		var grouping = this.getChartGrouping(oChart);
		var res = null;

		var point, oSer, i;
		if (grouping === "stacked") {
			for (i = 0; i <= ser; i++) {
				oSer = oChart.series[i];
				if (oSer) {
					point = this.getPointByIndex(oSer, val);
					if (point) {
						if (!res) {
							res = 0;
						}
						res += point.val;
					}
				}
			}
		} else if (grouping === "stackedPer") {
			var sumVal = 0;
			for (i = 0; i <= oChart.series.length; i++) {
				oSer = oChart.series[i];
				if (oSer) {
					point = this.getPointByIndex(oSer, val);
					if (point) {
						if (!res) {
							res = 0;
						}
						sumVal += Math.abs(point.val);
					}
				}
			}

			for (i = 0; i <= ser; i++) {
				oSer = oChart.series[i];
				if (oSer) {
					point = this.getPointByIndex(oSer, val);
					if (point) {
						if (!res) {
							res = 0;
						}
						res += point.val;
					}
				}
			}
			if(sumVal === 0) {
				res = 0;
			} else {
				res = res / sumVal;
			}
		} else {
			oSer = oChart.series[ser];
			point = this.getPointByIndex(oSer, val);
			if (point) {
				res = point.val;
			}
		}

		return res;
	},

	getValWithoutStacked: function (ser, val, oChart, bXVal) {
		var res = null;

		var oSer = oChart.series[ser];
		var point = this.getPointByIndex(oSer, val, bXVal);
		if (point) {
			res = point.val;
		}

		return res;
	},

	getValWithPercent: function (ser, val, oChart) {
		var grouping = this.getChartGrouping(oChart);
		var res = null;

		var point, oSer, i;
		if (grouping === "stackedPer") {
			var sumVal = 0;
			for (i = 0; i <= oChart.series.length; i++) {
				oSer = oChart.series[i];
				if (oSer) {
					point = this.getPointByIndex(oSer, val);
					if (point) {
						sumVal += Math.abs(point.val);
					}
				}
			}

			oSer = oChart.series[ser];
			if (oSer) {
				point = this.getPointByIndex(oSer, val);
				if (point) {
					if (!res) {
						res = 0;
					}
					res += point.val;
				}
			}

			if(sumVal === 0) {
				res = 0;
			} else {
				res = res / sumVal;
			}
		}

		return res;
	},

	getErrBarsPosition: function (errBars, oChart) {
		//true -> horizontal/false -vertical
		if (!errBars) {
			return null;
		}

		var res = null;
		if (errBars.errDir === AscFormat.st_errdirX) {
			res = true;
		} else if (errBars.errDir === AscFormat.st_errdirY) {
			res = false;
		} else if (errBars.errDir === null) {
			var chartType = this._getChartType(oChart.chart);
			if (chartType === c_oChartTypes.HBar) {
				res = true;
			} else {
				res = false;
			}
		}
		return res;
	},

	getNextArrElem: function (arr, index) {
		let res = null;
		for (let i = index + 1; i < arr.length; i++) {
			if (arr[i]) {
				res = arr[i];
				break;
			}
		}
		return res;
	},

	getNextCachePoint: function (oNumCache, index, dispBlanksAs) {
		let pt = oNumCache.getPtByIndex(index + 1);
		if (!pt) {
			if (dispBlanksAs === AscFormat.DISP_BLANKS_AS_SPAN) {
				//search next real pt
				for (let i = index + 2; i < oNumCache.ptCount; i++) {
					pt = oNumCache.getPtByIndex(i);
					if (pt) {
						break;
					}
				}
			} else if (dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
				pt = {val: 0, idx: index + 1};
			}
		}
		return pt;
	},

	calculateLine: function (x, y, x1, y1, path, isContinuePath) {
		let pathId;
		if (!path) {
			pathId = this.cChartSpace.AllocPath();
			path = this.cChartSpace.GetPath(pathId);
		}

		var pathH = this.calcProp.pathH;
		var pathW = this.calcProp.pathW;

		if (!isContinuePath) {
			path.moveTo(x * pathW, y * pathH);
		}
		path.lnTo(x1 * pathW, y1 * pathH);

		return pathId;
	},

	getChartByType: function (type) {
		for (let i in this.charts) {
			if (this.charts[i] && this.charts[i].chart.getObjectType() === type) {
				return this.charts[i].chart;
			}
		}
		return null;
	}
};


/** @constructor */
function drawBarChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;
	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
	this.sortZIndexPaths = [];
	this.summBarVal = [];

	this.temp = [];
	this.temp2 = [];
}

drawBarChart.prototype = {
	constructor: drawBarChart,

	recalculate: function () {
		this.paths = {};
		this.summBarVal = [];

		this.sortZIndexPaths = [];

		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.seriesCount = countSeries.series;
		this.ptCount = countSeries.points;
		this.subType = this.cChartDrawer.getChartGrouping(this.chart);
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);
		if (!this.catAx) {
			this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_DateAx);
		}
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);
		this.serAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_SerAx);

		this._recalculateBars();
	},

	draw: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			this._DrawBars3D();
		} else {
			this._DrawBars();
		}
	},

	_DrawBars: function () {
		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();

		var left = (this.chartProp.chartGutter._left - 1) / this.chartProp.pxToMM;
		var top = (this.chartProp.chartGutter._top - 1) / this.chartProp.pxToMM;
		var right = this.chartProp.trueWidth / this.chartProp.pxToMM;
		var bottom = this.chartProp.trueHeight / this.chartProp.pxToMM;
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(left, top, right, bottom);

		this.cChartDrawer.drawPaths(this.paths, this.chart.series, null, null, true);
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	_recalculateBars: function (/*isSkip*/) {

		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;


		var scaleAxis = this.valAx.scale;
		var axisMin = scaleAxis[0] < scaleAxis[scaleAxis.length - 1] ? scaleAxis[0] :scaleAxis[scaleAxis.length - 1];
		var axisMax = scaleAxis[0] < scaleAxis[scaleAxis.length - 1] ? scaleAxis[scaleAxis.length - 1] :scaleAxis[0];

		if (!xPoints || !yPoints) {
			return;
		}

		var widthGraph = this.chartProp.widthCanvas - this.chartProp.chartGutter._left - this.chartProp.chartGutter._right;

		var defaultOverlap = (this.subType === "stacked" || this.subType === "stackedPer" || this.subType === "standard") ? 100 : 0;
		var overlap = AscFormat.isRealNumber(this.chart.overlap) ? this.chart.overlap : defaultOverlap;
		var numCache = this.cChartDrawer.getNumCache(this.chart.series[0].val);
		var width = widthGraph / xPoints.length;
		if (this.cChartSpace.getValAxisCrossType() && numCache) {
			width = widthGraph / (xPoints.length - 1);
		}

		var gapWidth = AscFormat.isRealNumber(this.chart.gapWidth) ? this.chart.gapWidth : 150;

		var individualBarWidth = width / (this.seriesCount - (this.seriesCount - 1) * (overlap / 100) + gapWidth / 100);
		var widthOverLap = individualBarWidth * (overlap / 100);
		var hmargin = (gapWidth / 100 * individualBarWidth) / 2;

		var nullPositionOX = this.catAx.posY * this.chartProp.pxToMM;

		var height, startX, startY, val, paths, seriesHeight = [], tempValues = [], seria, startYColumnPosition, startXPosition, prevVal, idx, seriesCounter = 0;
		var cubeCount = 0, k;

		for (var i = 0; i < this.chart.series.length; i++) {
			numCache = this.cChartDrawer.getNumCache(this.chart.series[i].val);
			seria = numCache ? numCache.pts : [];
			seriesHeight[i] = [];
			tempValues[i] = [];

			if (numCache == null || this.chart.series[i].isHidden) {
				continue;
			}

			var isValMoreZero = false;
			var isValLessZero = 0;

			var shapeType = null !== this.chart.series[i].shape ? this.chart.series[i].shape : this.chart.shape;
			for (var j = 0; j < seria.length; j++) {

				//for 3d charts
				if (val > 0) {
					isValMoreZero = true;
				} else if (val < 0) {
					isValLessZero++;
				}

				//стартовая позиция колонки Y(+ высота с учётом поправок на накопительные диаграммы)
				val = parseFloat(seria[j].val);
				idx = seria[j].idx != null ? seria[j].idx : j;
				/*if (this.valAx && this.valAx.scaling.logBase) {
					val = this.cChartDrawer.getLogarithmicValue(val, this.valAx.scaling.logBase);
				}*/
				prevVal = 0;
				if (this.subType === "stacked" || this.subType === "stackedPer") {
					for (k = 0; k < tempValues.length; k++) {
						if (tempValues[k][idx] && tempValues[k][idx] > 0) {
							prevVal += tempValues[k][idx];
						}
					}
				}

				//shapeType = 5; //раскоментировать для теста пирамид
				tempValues[i][idx] = val;

				startYColumnPosition = this._getStartYColumnPosition(seriesHeight, i, idx, val, yPoints, prevVal, shapeType, axisMax, axisMin);
				startY = startYColumnPosition.startY;
				height = startYColumnPosition.height;

				var maxHeight = startYColumnPosition.maxH;
				var minHeight = startYColumnPosition.minH;

				seriesHeight[i][idx] = height;

				//стартовая позиция колонки X
				if (!this.catAx.isReversed()) {
					if (xPoints[1] && xPoints[1].pos && xPoints[idx]) {
						startXPosition = xPoints[idx].pos - Math.abs((xPoints[1].pos - xPoints[0].pos) / 2);
					} else if (xPoints[idx]) {
						startXPosition = xPoints[idx].pos - Math.abs(xPoints[0].pos - this.valAx.posX);
					} else {
						startXPosition = xPoints[0].pos - Math.abs(xPoints[0].pos - this.valAx.posX);
					}
				} else {
					if (xPoints[1] && xPoints[1].pos && xPoints[idx]) {
						startXPosition = xPoints[idx].pos + Math.abs((xPoints[1].pos - xPoints[0].pos) / 2);
					} else if (xPoints[idx]) {
						startXPosition = xPoints[idx].pos + Math.abs(xPoints[0].pos - this.valAx.posX);
					} else {
						startXPosition = xPoints[0].pos + Math.abs(xPoints[0].pos - this.valAx.posX);
					}
				}

				if (!this.catAx.isReversed()) {
					if (seriesCounter === 0) {
						startX =
							startXPosition * this.chartProp.pxToMM + hmargin + seriesCounter * (individualBarWidth);
					} else {
						startX = startXPosition * this.chartProp.pxToMM + hmargin +
							(seriesCounter * individualBarWidth - seriesCounter * widthOverLap);
					}
				} else {
					if (i === 0) {
						startX =
							startXPosition * this.chartProp.pxToMM - hmargin - seriesCounter * (individualBarWidth);
					} else {
						startX = startXPosition * this.chartProp.pxToMM - hmargin -
							(seriesCounter * individualBarWidth - seriesCounter * widthOverLap);
					}
				}


				if (this.catAx.isReversed()) {
					startX = startX - individualBarWidth;
				}

				if (this.valAx.isReversed() &&
					(this.subType === "stackedPer" || this.subType === "stacked")) {
					startY = startY + height;
				}

				//for 3d charts
				if (this.cChartDrawer.nDimensionCount === 3) {
					//расскомментируем, чтобы включить старую схему отрисовки(+ переименовать функции _DrawBars3D -> _DrawBars3D2)
					//this.sortZIndexPaths.push({seria: i, point: idx, paths: paths.paths, x: paths.x, y: paths.y, zIndex: paths.zIndex});

					var testHeight;
					if ((axisMax > 0 && axisMin > 0) || (axisMax < 0 && axisMin < 0)) {
						testHeight = Math.abs(yPoints[0].pos - yPoints[yPoints.length - 1].pos) * this.chartProp.pxToMM;
					} else {
						var endBlockPosition, startBlockPosition;
						if (val > 0) {
							endBlockPosition = this.cChartDrawer.getYPosition(axisMax, this.valAx) * this.chartProp.pxToMM;
							startBlockPosition = prevVal ? this.cChartDrawer.getYPosition(0, this.valAx) * this.chartProp.pxToMM : nullPositionOX;
							testHeight = startBlockPosition - endBlockPosition;
						} else {
							endBlockPosition = this.cChartDrawer.getYPosition(axisMin, this.valAx) * this.chartProp.pxToMM;
							startBlockPosition = prevVal ? this.cChartDrawer.getYPosition(0, this.valAx) * this.chartProp.pxToMM : nullPositionOX;
							testHeight = startBlockPosition - endBlockPosition;
						}
					}

					testHeight = this.subType === "normal" ? testHeight : height;
					paths = this._calculateRect3D(startX, startY, individualBarWidth, height, val, i, shapeType, maxHeight, minHeight, this.temp, cubeCount, idx, testHeight);

					if (this.subType === "standard") {
						for (k = 0; k < paths.paths.length; k++) {
							this.sortZIndexPaths.push({
								seria: i,
								point: idx,
								verge: k,
								paths: paths.paths[k],
								facePoints: paths.facePoints[k],
								shapeType: shapeType
							});
						}
					}
					paths = paths.paths2;

					cubeCount++;
				} else {
					paths = this.cChartDrawer._calculateRect(startX, startY, individualBarWidth, height);
				}

				var serIdx = this.chart.series[i].idx;
				if (!this.paths.series) {
					this.paths.series = [];
				}
				if (!this.paths.series[serIdx]) {
					this.paths.series[serIdx] = [];
				}
				this.paths.series[serIdx][idx] = paths;

				let errBars = this.chart.series[i].errBars[0];
				if (errBars) {
					var _pointX = (startX + individualBarWidth / 2) / this.chartProp.pxToMM;
					var pointY = (startY + height) / this.chartProp.pxToMM;
					var _pointVal = this.cChartDrawer.getValWithStacked(i, idx, this.chart);
					// var _pointVal = this.subType === "stacked" || this.subType === "stackedPer" ? this._getStackedValue(this.chart.series, i, j, val) : val;
					if (errBars) {
						this.cChartDrawer.errBars.putPoint(_pointX, pointY, _pointVal, _pointVal,  serIdx, idx);
					}
				}
			}

			seriesCounter++;
		}

		if (this.cChartDrawer.nDimensionCount === 3) {

			var getMinZ = function (_paths, checkRotation, length, verges, angelY, angelX, _gapDepth) {
				var arr = _paths.facePoints, verge = _paths.verge, seria = _paths.seria;
				var zIndex = 0;
				if (!arr) {
					arr = []
				}
				for (var i = 0; i < arr.length; i++) {
					if (i === 0) {
						zIndex = arr[i].z;
					} else if (arr[i].z < zIndex) {
						zIndex = arr[i].z;
					}
				}
				//проверяем что оси не под прямым углов и отступ в глубину OZ = 0
				//в зависимости от угла увеличиваем zIndex фронтальной грани не первой к точке наблюдения колонки
				if (!checkRotation && _gapDepth === 0) {
					if (verge === verges.front && seria > 0 && angelX > 0) {
						zIndex += angelX * 2;
					}
					if (verge === verges.unfront && seria < length && angelX > 0 && angelY > 90 && angelY <
						270) {
						zIndex += angelX * 2;
					}
				}
				return zIndex;
			};

			var angelX = this.cChartDrawer.processor3D.view3D.rotX;
			var angelY = this.cChartDrawer.processor3D.view3D.rotY;
			var cSortFaces;

			if (this.subType !== "standard") {
				this.sortParallelepipeds = [];
				//когда оси под прямым углом можно отсортировать без поиска пересечений в sortParallelepipeds
				if (this.cChartDrawer.processor3D.view3D.getRAngAx()) {
					var angle;
					if (this.subType !== "normal") {
						angle = -this.cChartDrawer.processor3D.angleOx;
						this.temp.sort(function sortArr(a, b) {
							if (b.z === a.z) {
								if (angle > 0) {
									return b.y - a.y;
								} else {
									return a.y - b.y;
								}
							} else {
								return b.z - a.z;
							}
						});
						for (var i = 0; i < this.temp.length; i++) {
							this.sortParallelepipeds.push({ nextIndex: i });
						}
					} else {
						angle = Math.abs(this.cChartDrawer.processor3D.angleOy);
						this.temp.sort(function sortArr(a, b) {
							if (b.z === a.z) {
								if (angle < Math.PI) {
									return a.x - b.x;
								} else {
									return b.x - a.x;
								}
							} else {
								return b.z - a.z;
							}
						});
						for (var i = 0; i < this.temp.length; i++) {
							this.sortParallelepipeds.push({ nextIndex: i });
						}
					}
				} else {
					cSortFaces = new window['AscFormat'].CSortFaces(this.cChartDrawer);
					this.sortParallelepipeds = cSortFaces.sortParallelepipeds(this.temp);
				}

			} else {
				var checkRotation = this.cChartDrawer.processor3D.view3D.getRAngAx();
				var length = this.chart.series.length - 1;
				var _gapDepth = this.chart.gapDepth;
				var verges = {
					front: 0, down: 1, left: 2, right: 3, up: 4, unfront: 5
				};

				this.sortZIndexPaths.sort(function sortArr(a, b) {
					var ZB = getMinZ(b, checkRotation, length, verges, angelY, angelX, _gapDepth);
					var ZA = getMinZ(a, checkRotation, length, verges, angelY, angelX, _gapDepth);
					if ((ZB === ZA) && (b.facePoints[0] && a.facePoints[0]))
						return b.facePoints[0].y - a.facePoints[0].y;

					return ZB - ZA;
				});
			}
		}
	},

	_getStartYColumnPosition: function (seriesHeight, i, j, val, yPoints, prevValue, type, axisMax, axisMin) {
		var startY, height, curVal, prevVal, endBlockPosition, startBlockPosition, maxH, minH, h;
		var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
			this.catAx.posY * this.chartProp.pxToMM;

		if (this.subType === "stacked" || this.subType === "stackedPer") {
			curVal = this._getStackedValue(this.chart.series, i, j, val);
			prevVal = this._getStackedValue(this.chart.series, i - 1, j, val);

			if (this.subType === "stacked") {
				//если максимальное значение задано вручную, и присутвуют точки, которые больше этого значения
				if (curVal > axisMax) {
					curVal = axisMax;
				}
				if (curVal < axisMin) {
					curVal = axisMin;
				}

				endBlockPosition = this.cChartDrawer.getYPosition(curVal, this.valAx, null, true) * this.chartProp.pxToMM;
				startBlockPosition = prevVal ? this.cChartDrawer.getYPosition(prevVal, this.valAx, null, true) * this.chartProp.pxToMM : nullPositionOX;
			} else {
				this._calculateSummStacked(j);

				var test = this.summBarVal[j];

				//если максимальное значение задано вручную, и присутвуют точки, которые больше этого значения
				if (curVal / test > axisMax) {
					curVal = axisMax * test;
				}
				if (curVal / test < axisMin) {
					curVal = axisMin * test;
				}

				if (prevVal / test > axisMax) {
					prevVal = axisMax * test;
				}
				if (prevVal / test < axisMin) {
					prevVal = axisMin * test;
				}

				endBlockPosition = this.cChartDrawer.getYPosition((curVal / test), this.valAx, null, true) * this.chartProp.pxToMM;
				startBlockPosition = test ? this.cChartDrawer.getYPosition((prevVal / test), this.valAx, null, true) * this.chartProp.pxToMM : nullPositionOX;
			}

			if (this.valAx.scaling.logBase) {
				startY = val === 0 ? nullPositionOX : startBlockPosition;
				height = val === 0 ? 0 : startBlockPosition - endBlockPosition;
			} else {
				startY = startBlockPosition;
				height = startBlockPosition - endBlockPosition;
			}

			if (this.valAx.isReversed()) {
				height = -height;
			}
		} else {
			if (this.valAx.scaling.logBase) {
				height = val === 0 ? 0 : nullPositionOX - this.cChartDrawer.getYPosition(val, this.valAx) * this.chartProp.pxToMM;
				startY = val === 0 ? nullPositionOX : nullPositionOX;
			} else {
				height = nullPositionOX - this.cChartDrawer.getYPosition(val, this.valAx) * this.chartProp.pxToMM;
				startY = nullPositionOX;
			}
		}
		if (type === AscFormat.BAR_SHAPE_PYRAMID || type === AscFormat.BAR_SHAPE_PYRAMIDTOMAX ||
			type === AscFormat.BAR_SHAPE_CONE || type === AscFormat.BAR_SHAPE_CONETOMAX) {
			var testMaxHeight = this.cChartDrawer.getYPosition(axisMax, this.valAx) * this.chartProp.pxToMM - nullPositionOX;
			var testMinHeight = this.cChartDrawer.getYPosition(axisMin, this.valAx) * this.chartProp.pxToMM - nullPositionOX;

			var maxVal = this._getStackedValue(this.chart.series, this.chart.series.length - 1, j, val);
			h = this.cChartDrawer.getStartStackedPyramidPosition(val, this.chart, this.valAx, j, this.chartProp,
				this.summBarVal, maxVal, nullPositionOX, this.subType, this.ptCount, axisMax, axisMin, testMaxHeight, testMinHeight, type, false);

			maxH = h.maxH;
			minH = h.minH;
		}

		return {startY: startY, height: height, maxH: maxH, minH: minH };
	},

	_calculateSummStacked: function (j) {
		if (!this.summBarVal[j]) {
			var curVal;
			var temp = 0;
			var idxPoint;
			for (var k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], j);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;

				if (curVal) {
					temp += Math.abs(curVal);
				}
			}

			this.summBarVal[j] = temp;
		}
	},

	_getStandardMaxHeight: function () {
		var curVal;
		var idxPoint;
		var mainMaxH = 0;
		var mainMinH = 0;
		for (var i = 0; i < this.ptCount; i++) {
			this._calculateSummStacked(i);
			for (var k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], i);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;

				if (curVal > mainMaxH) {
					mainMaxH = curVal;
				}
				if (curVal < mainMinH) {
					mainMinH = curVal;
				}
			}
		}

		return {mainMaxH: mainMaxH, mainMinH: mainMinH}
	},

	_getMaxHeightStacked: function () {
		var curVal;
		var tempMax = 0;
		var tempMin = 0;
		var idxPoint;
		var maxH1 = [];
		var minH1 = [];
		var maxH2 = [];
		var minH2 = [];
		var countMax = 0;
		var countMin = 0;

		//нахождение пропорционально наибольшей и наименьшей высоты для накопительных процентных пирамид
		for (var i = 0; i < this.ptCount; i++) {
			this._calculateSummStacked(i);
			for (var k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], i);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;

				if (curVal > 0) {
					tempMax += curVal;
				} else {
					tempMin += curVal;
				}
			}
			//пропорциональная высота
			maxH1[i] = tempMax / this.summBarVal[i];
			minH1[i] = tempMin / this.summBarVal[i];

			//обычная высота по значениям
			maxH2[i] = tempMax;
			minH2[i] = tempMin;
			tempMax = 0;
			tempMin = 0;
		}

		var indexMax = 0;
		var indexMin = 0;
		//сортировка высоты по пропорциональным положительным и отрицательным значениям
		for (var k = 0; k < maxH1.length; k++) {
			if (tempMax < maxH1[k]) {
				tempMax = maxH1[k];
				countMax = maxH2[k];
				indexMax = k;
			}
		}

		for (k = 0; k < minH1.length; k++) {
			if (tempMin > minH1[k]) {
				tempMin = minH1[k];
				countMin = minH2[k];
				indexMin = k;
			}
		}

		return {
			maxH: countMax,
			minH: countMin,
			indexMax: indexMax,
			indexMin: indexMin,
			maxPer: tempMax,
			minPer: tempMin
		};
	},

	_getMaxHeight: function (j) {
		var curVal;
		var tempMax = 0;
		var tempMin = 0;
		var idxPoint;
		for (var k = 0; k < this.chart.series.length; k++) {
			idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], j);
			curVal = idxPoint ? parseFloat(idxPoint.val) : 0;

			if (curVal > 0) {
				tempMax += curVal;
			} else {
				tempMin += curVal;
			}
		}

		return {maxH: tempMax, minH: tempMin};
	},

	_getStackedValue: function (series, i, j, val) {
		var result = 0, curVal, idxPoint;
		for (var k = 0; k <= i; k++) {
			idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], j);
			curVal = idxPoint ? idxPoint.val : 0;

			if (idxPoint && val >= 0 && curVal > 0) {
				result += parseFloat(curVal);
			} else if (idxPoint && val <= 0 && curVal < 0) {
				result += parseFloat(curVal);
			}
		}

		return result;
	},

	_calculateDLbl: function (chartSpace, ser, val, bLayout, serIdx) {
		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);
		if (!this.paths || !this.paths.series || !this.paths.series[serIdx] || !this.paths.series[serIdx][val] || !point || !point.compiledDlb) {
			return;
		}

		var path = this.paths.series[serIdx][val];
		//ToDo пересмотреть для 3d диаграмм
		var isZeroH = false;
		if (this.cChartDrawer.nDimensionCount === 3) {
			if (AscFormat.isRealNumber(path[0])) {
				path = path[0];
			} else if (AscFormat.isRealNumber(path[5])) {
				path = path[5];
			} else if (AscFormat.isRealNumber(path[2])) {
				path = path[2];
			} else if (AscFormat.isRealNumber(path[3])) {
				path = path[3];
			} else if (AscFormat.isRealNumber(path[1])) {
				//TODO добавлено для случая нулевой точки. возможно в данном случае сдвиги нужно считать иначе
				path = path[1];
				isZeroH = true;
			} else if (AscFormat.isRealNumber(path[4])) {
				path = path[4];
				isZeroH = true;
			}
		}

		if (!AscFormat.isRealNumber(path)) {
			return;
		}
		var oPath = this.cChartSpace.GetPath(path);
		var oCommand0 = oPath.getCommandByIndex(0);
		var oCommand1 = oPath.getCommandByIndex(1);
		var oCommand2 = oPath.getCommandByIndex(2);
		var oCommand3 = oPath.getCommandByIndex(3);

		var x = oCommand0.X;
		var y = isZeroH && point.val !== 0 ? oCommand3.Y : oCommand0.Y;

		var h = oCommand0.Y - oCommand1.Y;
		var w = oCommand2.X - oCommand1.X;

		var pxToMm = this.chartProp.pxToMM;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX, centerY;
		//TODO check revert valAx orientation for other charts labels
		var maxMinAxis = this.valAx && this.valAx.isReversed();

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				centerX = x + w / 2 - width / 2;
				if (isZeroH && point.val !== 0) {
					centerY = y - height / 2;
				} else {
					centerY = y - h / 2 - height / 2;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.inBase: {
				centerX = x + w / 2 - width / 2;
				centerY = y;
				if ((point.val > 0 && !maxMinAxis) || (point.val < 0 && maxMinAxis)) {
					centerY = y - height;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.inEnd: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h;
				if ((point.val < 0 && !maxMinAxis) || (point.val > 0 && maxMinAxis)) {
					centerY = centerY - height;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.outEnd: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h - height;
				if ((point.val < 0 && !maxMinAxis) || (point.val > 0 && maxMinAxis)) {
					centerY = centerY + height;
				}
				break;
			}
		}
		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},


	//TODO delete after test
	_DrawBars3D2: function () {
		var t = this;
		var processor3D = this.cChartDrawer.processor3D;

		var verges = {
			front: 0, down: 1, left: 2, right: 3, up: 4, unfront: 5
		};

		var drawVerges = function (i, j, paths, onlyLessNull, start, stop) {
			var brush, pen, options;
			options = t._getOptionsForDrawing(i, j, onlyLessNull);
			if (options !== null) {
				pen = options.pen;
				brush = options.brush;

				for (var k = start; k <= stop; k++) {
					t._drawBar3D(paths[k], pen, brush, k);
				}
			}
		};

		var draw = function (onlyLessNull, start, stop) {
			for (var i = 0; i < t.sortZIndexPaths.length; i++) {
				drawVerges(t.sortZIndexPaths[i].seria, t.sortZIndexPaths[i].point, t.sortZIndexPaths[i].paths,
					onlyLessNull, start, stop);
			}
		};

		if (this.subType === "standard") {
			draw(true, verges.front, verges.unfront);
			draw(false, verges.front, verges.unfront);
		} else {
			draw(true, verges.down, verges.up);
			draw(false, verges.down, verges.up);

			draw(true, verges.unfront, verges.unfront);
			draw(false, verges.unfront, verges.unfront);

			draw(true, verges.front, verges.front);
			draw(false, verges.front, verges.front);
		}
	},

	_DrawBars3D: function () {
		var t = this;

		var drawVerges = function (i, j, paths, onlyLessNull, k, isNotPen, isNotBrush) {
			var brush = null, pen = null, options;
			options = t._getOptionsForDrawing(i, j, onlyLessNull);
			if (paths !== null && options !== null) {
				if (!isNotPen) {
					pen = options.pen;
				}
				if (!isNotBrush) {
					brush = options.brush;
				}

				if (pen && pen.Join) {
					pen = pen.createDuplicate();
					pen.Join = new AscFormat.LineJoin();
					pen.Join.type = Asc['c_oAscLineJoinType'].Round;
				}

				t._drawBar3D(paths, pen, brush, k, options.val);
			}
		};

		var index, faces, face, isNotPen;
		if (this.subType !== "standard") {
			for (var i = 0; i < this.sortParallelepipeds.length; i++) {
				index = this.sortParallelepipeds[i].nextIndex;
				faces = this.temp[index].faces;
				for (var j = 0; j < faces.length; j++) {
					face = faces[j];
					isNotPen = this.cChartDrawer.checkingPenForDrawing(face.shapeType, face.verge);
					drawVerges(face.seria, face.point, face.frontPaths, null, face.verge, isNotPen);
				}
			}
		} else {
			for (var i = 0; i < this.sortZIndexPaths.length; i++) {
				isNotPen = this.cChartDrawer.checkingPenForDrawing(this.sortZIndexPaths[i].shapeType, this.sortZIndexPaths[i].verge);
				drawVerges(this.sortZIndexPaths[i].seria, this.sortZIndexPaths[i].point,
					this.sortZIndexPaths[i].paths, null, this.sortZIndexPaths[i].verge, isNotPen);
			}
		}

	},

	_getOptionsForDrawing: function (ser, point, onlyLessNull) {
		var seria = this.chart.series[ser];
		var numCache = this.cChartDrawer.getNumCache(seria.val);
		if (!numCache) {
			return null;
		}

		var pt = numCache.getPtByIndex(point);
		if (!seria || !this.paths.series[ser] || !this.paths.series[ser][point] || !pt) {
			return null;
		}

		var brush = seria.brush;
		var pen = seria.pen;

		if ((pt.val > 0 && onlyLessNull === true) || (pt.val < 0 && onlyLessNull === false)) {
			return null;
		}

		if (pt.pen) {
			pen = pt.pen;
		}
		if (pt.brush) {
			brush = pt.brush;
		}

		return {pen: pen, brush: brush, val: pt.val}
	},

	_drawBar3D: function (path, pen, brush, k, val) {
		//затемнение боковых сторон
		//в excel всегда темные боковые стороны, лицевая и задняя стороны светлые
		//TODO пересмотреть получения pen
		if (null === pen || (null !== pen && null === pen.Fill) ||
			(null !== pen && null !== pen.Fill && null === pen.Fill.fill)) {
			pen = AscFormat.CreatePenFromParams(brush, undefined, undefined, undefined, undefined, 0.1);
		}

		if (k !== 5 && k !== 0) {
			var props = this.cChartSpace.getParentObjects();
			var duplicateBrush = brush;
			if (null !== brush) {
				duplicateBrush = brush.createDuplicate();
				var cColorMod = new AscFormat.CColorMod;

				if (k === 1 || k === 4) {
					//для градиентной заливки верхнюю и нижнюю грань закрашиываем первым и последним цветом соотвенственно
					if (duplicateBrush.fill && AscDFH.historyitem_type_GradFill === duplicateBrush.fill.getObjectType()) {
						var colors = duplicateBrush.fill.colors;
						//ToDo проверить stacked charts!
						var color;
						var valAxOrientation = this.valAx.getOrientation();
						if ((val > 0 && valAxOrientation === ORIENTATION_MIN_MAX) ||
							(val < 0 && valAxOrientation !== ORIENTATION_MIN_MAX)) {
							if (k === 4 && colors && colors[0] && colors[0].color) {
								color = colors[0].color;
							} else if (k === 1 && colors[colors.length - 1] && colors[colors.length - 1].color) {
								color = colors[colors.length - 1].color;
							}
						} else {
							if (k === 4 && colors && colors[0] && colors[0].color) {
								color = colors[colors.length - 1].color;
							} else if (k === 1 && colors[colors.length - 1] && colors[colors.length - 1].color) {
								color = colors[0].color;
							}
						}

						var tempColor = new AscFormat.CUniFill();
						tempColor.setFill(new AscFormat.CSolidFill());
						tempColor.fill.setColor(color);
						duplicateBrush = tempColor;
					}

					cColorMod.val = 45000;
				} else {
					cColorMod.val = 35000;
				}

				cColorMod.name = "shade";
				duplicateBrush.addColorMod(cColorMod);
				duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master, new AscFormat.CUniColor().RGBA, this.cChartSpace.clrMapOvr);

				if (null === pen) {
					pen.setFill(duplicateBrush);
				}
			}

			this.cChartDrawer.drawPath(path, pen, duplicateBrush);
		} else {
			this.cChartDrawer.drawPath(path, pen, brush);
		}
	},

	_calculateRect3D: function (startX, startY, individualBarWidth, height, val, serNum, type, maxH, minH, arr, cubeCount, idx, testHeight) {
		var depths = this.getDepths(serNum);
		var gapDepth = depths.gapDepth, perspectiveDepth = depths.perspectiveDepth;

		var point1, point2, point3, point4, point5, point6, point7, point8;
		var isNotDrawDownVerge, paths2, paths, points, points, paths, nullPositionOX;

		var x12 = startX, y12 = startY, z12 = 0 + gapDepth;
		var x22 = startX, y22 = startY, z22 = perspectiveDepth + gapDepth;
		var x32 = startX + individualBarWidth, y32 = startY, z32 = perspectiveDepth + gapDepth;
		var x42 = startX + individualBarWidth, y42 = startY, z42 = 0 + gapDepth;
		var x52 = startX, y52 = startY - height, z52 = 0 + gapDepth;
		var x62 = startX, y62 = startY - height, z62 = perspectiveDepth + gapDepth;
		var x72 = startX + individualBarWidth, y72 = startY - height, z72 = perspectiveDepth + gapDepth;
		var x82 = startX + individualBarWidth, y82 = startY - height, z82 = 0 + gapDepth;

		point1 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12);
		point2 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22);
		point3 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32);
		point4 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42);
		point5 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52);
		point6 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62);
		point7 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72);
		point8 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82);

		var facePoints;
		switch (type) {
			case AscFormat.BAR_SHAPE_PYRAMID:
			case AscFormat.BAR_SHAPE_PYRAMIDTOMAX: {
				var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
					this.catAx.posY * this.chartProp.pxToMM;
				paths = this.cChartDrawer.calculatePyramid(false, this.subType, startX, startY, height, gapDepth,
					individualBarWidth, perspectiveDepth, val, nullPositionOX, maxH, minH);
				
				paths = paths.paths;
				break
			}
			case AscFormat.BAR_SHAPE_CYLINDER: {
				paths = this.cChartDrawer._calculateCylinder(startX, startY, individualBarWidth, height, val, gapDepth,
					perspectiveDepth, this.subType !== "standard", false, false);
				break;
			}
			case AscFormat.BAR_SHAPE_CONE:
			case AscFormat.BAR_SHAPE_CONETOMAX: {
				var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
					this.catAx.posY * this.chartProp.pxToMM;
				paths = this.cChartDrawer._calculateCylinder(startX, startY, individualBarWidth, height, val, gapDepth,
					perspectiveDepth, this.subType !== "standard", false, this.subType, nullPositionOX, maxH, minH);
				break;
			}
			default: {
				points = [point1, point2, point3, point4, point5, point6, point7, point8];
				paths = this.cChartDrawer.calculateRect3D(points, val, null, this.subType !== "standard", false);
				break;
			}
		}

		facePoints = [[point1, point4, point8, point5], [point1, point2, point3, point4],
			[point1, point2, point6, point5], [point4, point8, point7, point3], [point5, point6, point7, point8],
			[point2, point3, point7, point6]];
		
		//paths2 для расчета подписей, в дальнейшем нужно оптимизировать этот процесс
		points = [point1, point2, point3, point4, point5, point6, point7, point8];
		paths2 = this.cChartDrawer.calculateRect3D(points, val, isNotDrawDownVerge);
	
		if (this.subType !== "standard") {

			height = testHeight;
			x12 = startX, y12 = startY, z12 = 0 + gapDepth;
			x22 = startX, y22 = startY, z22 = perspectiveDepth + gapDepth;
			x32 = startX + individualBarWidth, y32 = startY, z32 = perspectiveDepth + gapDepth;
			x42 = startX + individualBarWidth, y42 = startY, z42 = 0 + gapDepth;
			x52 = startX, y52 = startY - height, z52 = 0 + gapDepth;
			x62 = startX, y62 = startY - height, z62 = perspectiveDepth + gapDepth;
			x72 = startX + individualBarWidth, y72 = startY - height, z72 = perspectiveDepth + gapDepth;
			x82 = startX + individualBarWidth, y82 = startY - height, z82 = 0 + gapDepth;
			
			point1 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12);
			point2 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22);
			point3 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32);
			point4 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42);
			point5 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52);
			point6 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62);
			point7 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72);
			point8 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82);

			var point11 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12, null, null, true);
			var point22 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22, null, null, true);
			var point33 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32, null, null, true);
			var point44 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42, null, null, true);
			var point55 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52, null, null, true);
			var point66 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62, null, null, true);
			var point77 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72, null, null, true);
			var point88 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82, null, null, true);

			var controlPoint = this.cChartDrawer._convertAndTurnPoint(x12 + individualBarWidth / 2, y12 - height / 2, z12);

			var arrPoints = [[point1, point4, point8, point5], [point1, point2, point3, point4],
				[point1, point2, point6, point5], [point4, point8, point7, point3], [point5, point6, point7, point8],
				[point6, point2, point3, point7]];

			var arrPoints2 = [[point11, point44, point88, point55], [point11, point22, point33, point44],
				[point11, point22, point66, point55], [point44, point88, point77, point33],
				[point55, point66, point77, point88], [point66, point22, point33, point77]];

			if (!arr) {
				arr = [];
			}
			if (!arr[cubeCount]) {
				arr[cubeCount] = {};
			}
			if (!arr[cubeCount].faces) {
				arr[cubeCount].faces = [];
				arr[cubeCount].arrPoints = [point11, point22, point33, point44, point55, point66, point77, point88];
				arr[cubeCount].z = controlPoint.z;
				arr[cubeCount].y = controlPoint.y;
				arr[cubeCount].x = controlPoint.x;
				arr[cubeCount].isValZero = val === 0;
			}
	
			for (var k = 0; k < paths.frontPaths.length; k++) {
				if (null === paths.frontPaths[k] && null === paths.darkPaths[k]) {
					continue;
				}
				//this.sortZIndexPaths.push({seria: i, point: idx, verge: k, paths: paths[k], points: arrPoints2[k], points2: arrPoints[k], plainEquation: plainEquation});

				var plainEquation = this.cChartDrawer.getPlainEquation(arrPoints2[k][0], arrPoints2[k][1], arrPoints2[k][2], arrPoints2[k][3]);
				var plainArea = this.cChartDrawer.getAreaQuadrilateral(arrPoints[k][0], arrPoints[k][1], arrPoints[k][2], arrPoints[k][3]);
				arr[cubeCount].faces.push({
					seria: serNum,
					point: idx,
					verge: k,
					frontPaths: paths.frontPaths[k],
					darkPaths: paths.darkPaths[k],
					points: arrPoints2[k],
					points2: arrPoints[k],
					plainEquation: plainEquation,
					plainArea: plainArea,
					shapeType: type,
				});
			}
		}

		return {
			paths: paths,
			paths2: paths2,
			facePoints: facePoints,
		};
	},

	_calculateRect3DNew: function (startX, startY, individualBarWidth, height, val, serNum, type, maxH, minH, arr, cubeCount, idx, testHeight) {
		var depths = this.getDepths(serNum);
		var gapDepth = depths.gapDepth, perspectiveDepth = depths.perspectiveDepth;

		var point1, point2, point3, point4, point5, point6, point7, point8;
		var isNotDrawDownVerge, paths2, paths, points, points, paths;

		var x12 = startX, y12 = startY, z12 = 0 + gapDepth;
		var x22 = startX, y22 = startY, z22 = perspectiveDepth + gapDepth;
		var x32 = startX + individualBarWidth, y32 = startY, z32 = perspectiveDepth + gapDepth;
		var x42 = startX + individualBarWidth, y42 = startY, z42 = 0 + gapDepth;
		var x52 = startX, y52 = startY - height, z52 = 0 + gapDepth;
		var x62 = startX, y62 = startY - height, z62 = perspectiveDepth + gapDepth;
		var x72 = startX + individualBarWidth, y72 = startY - height, z72 = perspectiveDepth + gapDepth;
		var x82 = startX + individualBarWidth, y82 = startY - height, z82 = 0 + gapDepth;

		point1 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12);
		point2 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22);
		point3 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32);
		point4 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42);
		point5 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52);
		point6 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62);
		point7 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72);
		point8 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82);


		var needAddOnlyFrontFace = this.subType === "standard";
		var nullPositionOX = this.catAx.posY * this.chartProp.pxToMM;

		var t = this;
		var dRadius1, dRadius2, uRadius1, uRadius2;
		var calcCylinderProps = function (bCone) {
			//большая и малая полуось оснований усеченного конуса
			if (bCone) {
				// за оси эллипса берем 1/2 длин ребер оснований усеченной пирамиды
				points = t.cChartDrawer.isConeIntersection(false, t.subType, startX, startY, height, gapDepth, individualBarWidth, perspectiveDepth, val, nullPositionOX, maxH, minH);

				if (t.val === 0) {
					uRadius1 = 0;
					uRadius2 = 0;
					dRadius1 = points.wDown;
					dRadius2 = points.lDown;
				} else if ((t.subType === "stacked" || t.subType === "stackedPer") && t.cChartSpace.chart.plotArea.valAx.isReversed()) {
					uRadius1 = points.wUp !== 0 ? points.wUp : individualBarWidth / 2;
					uRadius2 = points.lUp !== 0 ? points.lUp : perspectiveDepth / 2;
					dRadius1 = points.wDown;
					dRadius2 = points.lDown;
				} else {
					uRadius1 = points.wUp;
					uRadius2 = points.lUp;
					dRadius1 = points.wDown !== 0 ? points.wDown : individualBarWidth / 2;
					dRadius2 = points.lDown !== 0 ? points.lDown : perspectiveDepth/ 2;
				}

			} else {
				//большая и малая полуось эллипса
				dRadius1 = individualBarWidth / 2;
				dRadius2 = perspectiveDepth / 2;
			}
		};


		var facePoints;
		switch (type) {
			case AscFormat.BAR_SHAPE_PYRAMID:
			case AscFormat.BAR_SHAPE_PYRAMIDTOMAX: {
				var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
					this.catAx.posY * this.chartProp.pxToMM;
				paths = this.cChartDrawer.calculatePyramid(false, this.subType, startX, startY, height, gapDepth,
					individualBarWidth, perspectiveDepth, val, nullPositionOX, maxH, minH);

				paths = paths.paths;
				break
			}
			case AscFormat.BAR_SHAPE_CYLINDER: {
				//paths = this.cChartDrawer._calculateCylinder(startX, startY, individualBarWidth, height, val, gapDepth, perspectiveDepth, this.subType !== "standard", false, false);

				calcCylinderProps();
				this.cChartDrawer.calcShapesHelper.setProps(startX, startY, gapDepth, individualBarWidth, height, perspectiveDepth, needAddOnlyFrontFace, val, false, dRadius1, dRadius2);
				paths = this.cChartDrawer.calcShapesHelper.getCylinder();

				break;
			}
			case AscFormat.BAR_SHAPE_CONE:
			case AscFormat.BAR_SHAPE_CONETOMAX: {
				//paths = this.cChartDrawer._calculateCylinder(startX, startY, individualBarWidth, height, val, gapDepth, perspectiveDepth, this.subType !== "standard", false, this.subType, nullPositionOX, maxH, minH);

				calcCylinderProps(true);
				this.cChartDrawer.calcShapesHelper.setProps(startX, startY, gapDepth, individualBarWidth, height, perspectiveDepth, needAddOnlyFrontFace, val, false, dRadius1, dRadius2, uRadius1, uRadius2);
				paths = this.cChartDrawer.calcShapesHelper.getCone();

				break;
			}
			default: {
				points = [point1, point2, point3, point4, point5, point6, point7, point8];
				paths = this.cChartDrawer.calculateRect3D(points, val, null, this.subType !== "standard", false);
				break;
			}
		}

		facePoints = [[point1, point4, point8, point5], [point1, point2, point3, point4],
			[point1, point2, point6, point5], [point4, point8, point7, point3], [point5, point6, point7, point8],
			[point2, point3, point7, point6]];

		//paths2 для расчета подписей, в дальнейшем нужно оптимизировать этот процесс
		points = [point1, point2, point3, point4, point5, point6, point7, point8];
		paths2 = this.cChartDrawer.calculateRect3D(points, val, isNotDrawDownVerge);

		if (this.subType !== "standard") {

			height = testHeight;
			x12 = startX, y12 = startY, z12 = 0 + gapDepth;
			x22 = startX, y22 = startY, z22 = perspectiveDepth + gapDepth;
			x32 = startX + individualBarWidth, y32 = startY, z32 = perspectiveDepth + gapDepth;
			x42 = startX + individualBarWidth, y42 = startY, z42 = 0 + gapDepth;
			x52 = startX, y52 = startY - height, z52 = 0 + gapDepth;
			x62 = startX, y62 = startY - height, z62 = perspectiveDepth + gapDepth;
			x72 = startX + individualBarWidth, y72 = startY - height, z72 = perspectiveDepth + gapDepth;
			x82 = startX + individualBarWidth, y82 = startY - height, z82 = 0 + gapDepth;

			point1 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12);
			point2 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22);
			point3 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32);
			point4 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42);
			point5 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52);
			point6 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62);
			point7 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72);
			point8 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82);

			var point11 = this.cChartDrawer._convertAndTurnPoint(x12, y12, z12, null, null, true);
			var point22 = this.cChartDrawer._convertAndTurnPoint(x22, y22, z22, null, null, true);
			var point33 = this.cChartDrawer._convertAndTurnPoint(x32, y32, z32, null, null, true);
			var point44 = this.cChartDrawer._convertAndTurnPoint(x42, y42, z42, null, null, true);
			var point55 = this.cChartDrawer._convertAndTurnPoint(x52, y52, z52, null, null, true);
			var point66 = this.cChartDrawer._convertAndTurnPoint(x62, y62, z62, null, null, true);
			var point77 = this.cChartDrawer._convertAndTurnPoint(x72, y72, z72, null, null, true);
			var point88 = this.cChartDrawer._convertAndTurnPoint(x82, y82, z82, null, null, true);

			var controlPoint = this.cChartDrawer._convertAndTurnPoint(x12 + individualBarWidth / 2, y12 - height / 2, z12);

			var arrPoints = [[point1, point4, point8, point5], [point1, point2, point3, point4],
				[point1, point2, point6, point5], [point4, point8, point7, point3], [point5, point6, point7, point8],
				[point6, point2, point3, point7]];

			var arrPoints2 = [[point11, point44, point88, point55], [point11, point22, point33, point44],
				[point11, point22, point66, point55], [point44, point88, point77, point33],
				[point55, point66, point77, point88], [point66, point22, point33, point77]];

			if (!arr) {
				arr = [];
			}
			if (!arr[cubeCount]) {
				arr[cubeCount] = {};
			}
			if (!arr[cubeCount].faces) {
				arr[cubeCount].faces = [];
				arr[cubeCount].arrPoints = [point11, point22, point33, point44, point55, point66, point77, point88];
				arr[cubeCount].z = controlPoint.z;
				arr[cubeCount].y = controlPoint.y;
				arr[cubeCount].x = controlPoint.x;
				arr[cubeCount].isValZero = val === 0;
			}

			for (var k = 0; k < paths.frontPaths.length; k++) {
				if (null === paths.frontPaths[k] && null === paths.darkPaths[k]) {
					continue;
				}
				//this.sortZIndexPaths.push({seria: i, point: idx, verge: k, paths: paths[k], points: arrPoints2[k], points2: arrPoints[k], plainEquation: plainEquation});

				var plainEquation = this.cChartDrawer.getPlainEquation(arrPoints2[k][0], arrPoints2[k][1], arrPoints2[k][2], arrPoints2[k][3]);
				var plainArea = this.cChartDrawer.getAreaQuadrilateral(arrPoints[k][0], arrPoints[k][1], arrPoints[k][2], arrPoints[k][3]);
				arr[cubeCount].faces.push({
					seria: serNum,
					point: idx,
					verge: k,
					frontPaths: paths.frontPaths[k],
					darkPaths: paths.darkPaths[k],
					points: arrPoints2[k],
					points2: arrPoints[k],
					plainEquation: plainEquation,
					plainArea: plainArea,
					shapeType: type,
				});
			}
		}

		return {
			paths: paths,
			paths2: paths2,
			facePoints: facePoints,
		};
	},

	getDepths: function (serNum) {
		//параметр r и глубина по OZ
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;

		//сдвиг по OZ в глубину
		var gapDepth = this.chart.gapDepth != null ? this.chart.gapDepth : globalGapDepth;
		if (this.subType === "standard") {
			perspectiveDepth = (perspectiveDepth / (gapDepth / 100 + 1)) / this.seriesCount;
		} else {
			perspectiveDepth = perspectiveDepth / (gapDepth / 100 + 1);
		}

		//если есть ось z берем pos из zPoints
		if (this.serAx && this.serAx.zPoints[serNum]) {
			return { gapDepth: this.serAx.zPoints[serNum].pos - (perspectiveDepth / 2), perspectiveDepth: perspectiveDepth };
		}

		var DiffGapDepth = perspectiveDepth * (gapDepth / 2) / 100;
		if (this.subType === "standard") {
			gapDepth = (perspectiveDepth + DiffGapDepth + DiffGapDepth) * serNum + DiffGapDepth;
		} else {
			gapDepth = DiffGapDepth;
		}

		return { gapDepth: gapDepth, perspectiveDepth: perspectiveDepth };
	}
};



/** @constructor */
function drawClusteredColumn(seria, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.seria = seria;

	this.catAx = null;
	this.valAx = null;

	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
}

drawClusteredColumn.prototype = {
	constructor: drawClusteredColumn,

	recalculate: function () {
		const cachedData = this.cChartSpace ? this.cChartSpace.getCachedData() : null;
		this.recalculateClusteredColumn(cachedData);
	},

	recalculateClusteredColumn: function (cachedData) {
		// need data, at least two axes, chartProp and chartGutter
		if (!cachedData || !this.cChartSpace.chart.plotArea.axId || !Array.isArray(this.cChartSpace.chart.plotArea.axId) || this.cChartSpace.chart.plotArea.axId.length < 2 || !this.chartProp || !this.chartProp.chartGutter) {
			return;
		}

		const valAxis = this.cChartSpace.chart.plotArea.axId[1];
		const catAxis = this.cChartSpace.chart.plotArea.axId[0];

		const catStart = this.chartProp.chartGutter._left;
		let valStart = this.cChartSpace.chart.plotArea.axId ? this.cChartSpace.chart.plotArea.axId[0].posY * this.chartProp.pxToMM : this.chartProp.trueHeight + this.chartProp.chartGutter._top;
		const coeff = catAxis.scaling.gapWidth;

		// 1 px gap for each section length
		const gapWidth = 0.5 / this.chartProp.pxToMM;
		const gapNumber = cachedData.data.length;
		//Each bar will have 2 gapWidth and 2 margins , on left and right sides
		const initialBarWidth = (this.chartProp.trueWidth - (2 * gapWidth * gapNumber)) / cachedData.data.length;
		const barWidth = (initialBarWidth / (1 + coeff));
		const margin = (initialBarWidth - barWidth) / 2;

		let start = (catStart + margin + gapWidth);
		// margin and gapwidth in the left and in the right;
		const gap = 2 * (margin + gapWidth);

		// save pareto line its beginning, and interval;
		if (cachedData.paretoLine) {
			cachedData.paretoLine.push(start + barWidth / 2);
			cachedData.paretoLine.push(gap + barWidth);
		}

		const cChartDrawer = this.cChartDrawer;
		const chartProp = this.chartProp;
		const paths = this.paths;

		const recalculateBar = function (fData, i) {
			const startY = cChartDrawer.getYPosition(fData.val, valAxis, true);
			const bW = i === 0 ? barWidth : barWidth - gapWidth;
			const height = valStart - (startY * chartProp.pxToMM);
			paths[i] = cChartDrawer._calculateRect(start, valStart, bW, height);
			start += (bW + gap);
		}


		cachedData.data.forEach(recalculateBar);
	},

	draw: function () {
		if (!this.cChartDrawer || !this.cChartDrawer.cShapeDrawer || !this.cChartDrawer.cShapeDrawer.Graphics || !this.chartProp || !this.chartProp.chartGutter) {
			return;
		}
		this.startRect();
		this.drawChart();
		this.endRect();
	},

	startRect: function () {
		// find chart starting coordinates, width and height;
		let leftRect = this.chartProp.chartGutter._left / this.chartProp.pxToMM;
		let topRect = (this.chartProp.chartGutter._top) / this.chartProp.pxToMM;
		let rightRect = this.chartProp.trueWidth / this.chartProp.pxToMM;
		let bottomRect = (this.chartProp.trueHeight) / this.chartProp.pxToMM;

		if (!AscFormat.isRealNumber(leftRect) || !AscFormat.isRealNumber(topRect) || !AscFormat.isRealNumber(rightRect) || !AscFormat.isRealNumber(bottomRect) ) {
			return
		}

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);
	},

	drawChart: function () {
		//TODO !!!
		//series color
		/*<cx:plotArea>
		<cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="60000"/>
			<a:lumOff val="40000"/>
			</a:srgbClr>
			</a:solidFill>*/

		//point color
		/*cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:dataPt idx="2">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="40000"/>
			<a:lumOff val="60000"/>
			</a:srgbClr>
			</a:solidFill>
			</cx:spPr>*/

		//this.cChartSpace.chart.plotArea.plotAreaRegion.series[0].spPr.Fill.ln

		let series = this.cChartSpace.chart.plotArea.plotAreaRegion.series;

		let oSeries = series[0];
		if(oSeries) {
			for (let i in this.paths) {
				if (this.paths.hasOwnProperty(i) && this.paths[i]) {
					let nPtIdx = parseInt(i);
					let pen = oSeries.getPtPen(nPtIdx);
					let brush = oSeries.getPtBrush(nPtIdx);
					this.cChartDrawer.drawPath(this.paths[i], pen, brush);
				}
			}
		}
	},

	endRect: function () {
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	_calculateDLbl: function (compiledDlb) {
		if (!this.paths || !this.chartProp) {
			return;
		}
		const path = this.paths[compiledDlb.idx];

		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		const oPath = this.cChartSpace.GetPath(path);
		const oCommand0 = oPath.getCommandByIndex(0);
		const oCommand1 = oPath.getCommandByIndex(1);
		const oCommand2 = oPath.getCommandByIndex(2);

		const x = oCommand0.X;
		const y = oCommand0.Y;

		const h = oCommand0.Y - oCommand1.Y;
		const w = oCommand2.X - oCommand1.X;

		const pxToMm = this.chartProp.pxToMM;

		const width = compiledDlb.extX;
		const height = compiledDlb.extY;

		let centerX, centerY;
		// DATA_LABEL_POS_IN_END
		switch (compiledDlb.dLblPos) {
			case AscFormat.DATA_LABEL_POS_CTR: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h / 2 - height / 2;
				break;
			}
			case AscFormat.DATA_LABEL_POS_IN_BASE: {
				centerX = x + w / 2 - width / 2;
				centerY = y - height;
				break;
			}
			case AscFormat.DATA_LABEL_POS_IN_END: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h;
				break;
			}
			case AscFormat.DATA_LABEL_POS_OUT_END: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h - height;
				break;
			}
		}

		const top = this.chartProp.chartGutter._top / pxToMm;
		if (centerY < top) {
			centerY = top;
		}

		const bottom = ((this.chartProp.trueHeight + this.chartProp.chartGutter._top) / pxToMm) - height ;
		if (centerY >= bottom) {
			centerY = bottom;
		}

		return {x: centerX, y: centerY};
	}
};

function drawParetoChart(seria, chartsDrawer) {
	drawClusteredColumn.call(this, seria, chartsDrawer);
	this.linePath = null;
}

AscFormat.InitClassWithoutType(drawParetoChart, drawClusteredColumn);

drawParetoChart.prototype.recalculate = function () {
		const cachedData = this.cChartSpace ? this.cChartSpace.getCachedData() : null;
		cachedData.paretoLine = [];
		this.recalculateClusteredColumn(cachedData);
		const nTotal = this.calculateTotal(cachedData.data);
		this.recalculateLinePath(cachedData, nTotal);
};

drawParetoChart.prototype.recalculateLinePath = function (cachedData, nTotal) {
	if (!cachedData || !this.cChartSpace.chart.plotArea.axId || !Array.isArray(this.cChartSpace.chart.plotArea.axId) || this.cChartSpace.chart.plotArea.axId.length < 3 || !this.chartProp && !this.chartProp.chartGutter) {
		return null;
	}
	const valAxis = this.cChartSpace.chart.plotArea.axId[2];
	const pathId = this.cChartSpace.AllocPath();
	const path = this.cChartSpace.GetPath(pathId);

	const pathH = this.chartProp.pathH;
	const pathW = this.chartProp.pathW;
	const pxToMm = this.chartProp.pxToMM;

	// starting point
	// interval
	let percentage = 0;
	for (let i = 0; i < cachedData.data.length; i++) {
		const val = cachedData.data[i].val;
		percentage += (val / nTotal);
		const yPos = this.cChartDrawer.getYPosition(percentage, valAxis, true);
		const xPos = cachedData.paretoLine[0] + i * cachedData.paretoLine[1];
		if (i === 0) {
			path.moveTo((xPos / pxToMm) * pathW, yPos * pathH);
		} else {
			path.lnTo((xPos / pxToMm) * pathW, yPos * pathH);
		}
	}
	this.linePath = pathId;
};

drawParetoChart.prototype.calculateTotal = function (aData) {
	if (!aData) {
		return null;
	}

	// calculate max number of occurrences
	let nTotal = 0;
	for (let i = 0; i < aData.length; i++) {
		// aggregation object does not have field occurrence;
		const val = aData[i].val;
		nTotal += val;
	}
	return nTotal;

};

drawParetoChart.prototype.draw = function () {
		if (!this.cChartDrawer || !this.cChartDrawer.cShapeDrawer || !this.cChartDrawer.cShapeDrawer.Graphics || !this.chartProp || !this.chartProp.chartGutter) {
			return;
		}
		this.startRect();
		this.drawChart();
		this.drawParetoLine();
		this.endRect();
};

drawParetoChart.prototype.drawParetoLine = function () {
	if (!this.linePath) {
		return;
	}
	const pen = this.cChartSpace && this.cChartSpace.chart && this.cChartSpace.chart.plotArea && this.cChartSpace.chart.plotArea.axId && this.cChartSpace.chart.plotArea.axId[1] ? this.cChartSpace.chart.plotArea.axId[1].compiledMajorGridLines : null;
	if (pen) {
		this.cChartDrawer.drawPath(this.linePath, pen);
	}
};

/** @constructor */
function drawWaterfallChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;

	this.catAx = null;
	this.valAx = null;

	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
}

drawWaterfallChart.prototype = {
	constructor: drawWaterfallChart,

	recalculate: function () {
		const cachedData = this.cChartSpace ? this.cChartSpace.getCachedData() : null;
		if (!cachedData || !this.cChartSpace.chart.plotArea.axId) {
			return;
		}

		if (cachedData.data.length > 0 && this.chartProp && this.chartProp.chartGutter) {
			const valAxis = this.cChartSpace.chart.plotArea.axId[1];
			const catAxis = this.cChartSpace.chart.plotArea.axId[0];

			const catStart = this.chartProp.chartGutter._left;
			let valStart = this.cChartSpace.chart.plotArea.axId ? this.cChartSpace.chart.plotArea.axId[0].posY * this.chartProp.pxToMM : this.chartProp.trueHeight + this.chartProp.chartGutter._top;
			const zeroValStart = valStart;

			if (AscFormat.isRealNumber(valStart) && AscFormat.isRealNumber(this.chartProp.pxToMM)) {
				const coeff = catAxis.scaling.gapWidth;
				// 1 px gap for each section length
				const gapWidth = 0.5 / this.chartProp.pxToMM;
				const gapNumber = cachedData.data.length;
				//Each bar will have 2 gapWidth and 2 margins , on left and right sides
				const initialBarWidth = (this.chartProp.trueWidth - (2 * gapWidth * gapNumber))/ cachedData.data.length;
				const barWidth = (initialBarWidth / (1 + coeff));
				const margin = (initialBarWidth - barWidth) / 2;

				let start = (catStart + margin + gapWidth);

				const cChartDrawer = this.cChartDrawer;
				const cChartSpace = this.cChartSpace;
				const chartProp = this.chartProp;
				const paths = this.paths;
				const calculateConnectorLine = this._calculateConnectorLine;

				const recalculateBar = function (fVal, i, isTotal) {
					const valPos = cChartDrawer.getYPosition(fVal, valAxis, true) * cChartDrawer.calcProp.pxToMM; // calc roof of the current bar
					const next = start + (barWidth + margin + gapWidth + gapWidth + margin); // start of the next bar
					const height = isTotal && i !== 0 ? zeroValStart - valPos : valStart - valPos;
					paths[i] = [];
					paths[i].push(cChartDrawer._calculateRect(start, isTotal ? zeroValStart : valStart, barWidth, height));
					//dont need last connectorLine
					if (i !== cachedData.data.length - 1 && margin !== 0) {
						//exclude gapWidth from connector line
						paths[i].push(calculateConnectorLine(valPos, start + barWidth + gapWidth, next - gapWidth, cChartSpace, chartProp))
					}
					valStart = valPos;	// go up in y direction
					start = next;	// go right in x direction
				}

				cachedData.forEach(recalculateBar);
			}
		}
	},

	_calculateConnectorLine: function (y, x1, x2, cChartSpace, chartProp) {
		if (!cChartSpace || !chartProp) {
			return null;
		}
		const pathId = cChartSpace.AllocPath();
		const path = cChartSpace.GetPath(pathId);

		const pathH = chartProp.pathH;
		const pathW = chartProp.pathW;
		const pxToMm = chartProp.pxToMM;

		path.moveTo((x1 / pxToMm) * pathW, (y / pxToMm)  * pathH);
		path.lnTo((x2 / pxToMm) * pathW, (y / pxToMm)  * pathH);

		return pathId;
	},

	draw: function () {
		if (!this.cChartDrawer || !this.cChartDrawer.calcProp || !this.cChartDrawer.cShapeDrawer || !this.cChartDrawer.cShapeDrawer.Graphics || !this.cChartDrawer.calcProp.chartGutter) {
			return;
		}

		// find chart starting coordinates, width and height;
		let leftRect = this.cChartDrawer.calcProp.chartGutter._left / this.cChartDrawer.calcProp.pxToMM;
		let topRect = (this.cChartDrawer.calcProp.chartGutter._top) / this.cChartDrawer.calcProp.pxToMM;
		let rightRect = this.cChartDrawer.calcProp.trueWidth / this.cChartDrawer.calcProp.pxToMM;
		let bottomRect = (this.cChartDrawer.calcProp.trueHeight) / this.cChartDrawer.calcProp.pxToMM;

		if (!AscFormat.isRealNumber(leftRect) || !AscFormat.isRealNumber(topRect) || !AscFormat.isRealNumber(rightRect) || !AscFormat.isRealNumber(bottomRect) ) {
			return;
		}

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);


		//TODO !!!
		//series color
		/*<cx:plotArea>
		<cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="60000"/>
			<a:lumOff val="40000"/>
			</a:srgbClr>
			</a:solidFill>*/

		//point color
		/*cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:dataPt idx="2">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="40000"/>
			<a:lumOff val="60000"/>
			</a:srgbClr>
			</a:solidFill>
			</cx:spPr>*/

		//this.cChartSpace.chart.plotArea.plotAreaRegion.series[0].spPr.Fill.ln

		let series = this.cChartSpace.chart.plotArea.plotAreaRegion.series;

		let oSeries = series[0];
		if(oSeries) {
			for (let i in this.paths) {
				if (this.paths.hasOwnProperty(i) && this.paths[i]) {
					for (let j in this.paths[i]) {
						if (this.paths[i].hasOwnProperty(j) && this.paths[i][j]) {
							let nPtIdx = parseInt(i);
							let pen = (j === '0' ) ? oSeries.getPtPen(nPtIdx) : this.cChartSpace.chart.plotArea.axId[1].compiledMajorGridLines;
							let brush = oSeries.getPtBrush(nPtIdx);
							this.cChartDrawer.drawPath(this.paths[i][j], pen, brush);
						}
					}
				}
			}
		}
		
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	}, 

	_calculateDLbl: function (compiledDlb) {
		if (!this.paths || !this.chartProp) {
			return;
		}
		const path = this.paths[compiledDlb.idx] && Array.isArray(this.paths[compiledDlb.idx]) && this.paths[compiledDlb.idx].length > 0 ? this.paths[compiledDlb.idx][0] : null;

		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		const oPath = this.cChartSpace.GetPath(path);
		const oCommand0 = oPath.getCommandByIndex(0);
		const oCommand1 = oPath.getCommandByIndex(1);
		const oCommand2 = oPath.getCommandByIndex(2);

		const x = oCommand0.X;
		const y = oCommand0.Y;

		const h = oCommand0.Y - oCommand1.Y;
		const w = oCommand2.X - oCommand1.X;

		const pxToMm = this.chartProp.pxToMM;

		const width = compiledDlb.extX;
		const height = compiledDlb.extY;

		let centerX, centerY;
		// DATA_LABEL_POS_IN_END
		switch (compiledDlb.dLblPos) {
			case AscFormat.DATA_LABEL_POS_CTR: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h / 2 - height / 2;
				break;
			}
			case AscFormat.DATA_LABEL_POS_IN_BASE: {
				centerX = x + w / 2 - width / 2;
				centerY = y - height;
				break;
			}
			case AscFormat.DATA_LABEL_POS_IN_END: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h;
				break;
			}
			case AscFormat.DATA_LABEL_POS_OUT_END: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h - height;
				break;
			}
		}

		const top = this.chartProp.chartGutter._top / pxToMm;
		if (centerY < top) {
			centerY = top;
		}

		const bottom = ((this.chartProp.trueHeight + this.chartProp.chartGutter._top) / pxToMm) - height ;
		if (centerY >= bottom) {
			centerY = bottom;
		}

		return {x: centerX, y: centerY};
	}
}

/** @constructor */
function drawFunnelChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;

	this.catAx = null;
	this.valAx = null;

	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
	this.linePath = null;
}

drawFunnelChart.prototype = {
	constructor: drawFunnelChart,

	recalculate: function () {
		const cachedData = this.cChartSpace ? this.cChartSpace.getCachedData() : null;
		if (!cachedData || !this.cChartSpace.chart.plotArea.axId) {
			return;
		}
		if (cachedData.data.length > 0 && this.chartProp && this.chartProp.chartGutter) {
			const valAxis = this.cChartSpace.chart.plotArea.axId[0];

			const catMiddle = (this.chartProp.trueWidth / 2 + this.chartProp.chartGutter._left);
			let valStart = this.chartProp.chartGutter._top;
			let chartHeight = this.chartProp.trueHeight;

			if (AscFormat.isRealNumber(valStart) && AscFormat.isRealNumber(catMiddle)) {
				const coeff = valAxis.scaling.gapWidth;
				// 1 px gap for each section length
				const gapWidth = 0.5 / this.chartProp.pxToMM;
				const gapNumber = cachedData.data.length;
				//Each bar will have 2 gapWidth and 2 margins , on top and bottom sides
				const initialBarHeight = (chartHeight - (2 * gapWidth * gapNumber))/ cachedData.data.length;
				const barHeight = (initialBarHeight / (1 + coeff));
				const margin = (initialBarHeight - barHeight) / 2;

				// because calculate rect accepts bottom left point as starting y, we add barHeight to the calculation of starting point 
				let startVertical = (valStart + margin + gapWidth + barHeight);

				const cChartDrawer = this.cChartDrawer;
				const chartProp = this.chartProp;
				const paths = this.paths;

				const recalculateBar = function (fVal, i) {
					const barWidth = valAxis.max && fVal > 0 ? (fVal / valAxis.max) * chartProp.trueWidth: 0;
					paths[i] = cChartDrawer._calculateRect(catMiddle - (barWidth / 2), startVertical, barWidth, barHeight);
					startVertical += margin + gapWidth + gapWidth + margin + barHeight;
				}

				cachedData.forEach(recalculateBar);

				// vertical line with 1 px from left 
				this.linePath = this.cChartDrawer._calculateLine(startVertical, (startVertical - valStart), this.chartProp.chartGutter._left + 1, true);
			}
		}
	},

	draw: function () {
		if (!this.cChartDrawer || !this.cChartDrawer.calcProp || !this.cChartDrawer.cShapeDrawer || !this.cChartDrawer.cShapeDrawer.Graphics || !this.cChartDrawer.calcProp.chartGutter) {
			return;
		}

		// find chart starting coordinates, width and height;
		let leftRect = this.cChartDrawer.calcProp.chartGutter._left / this.cChartDrawer.calcProp.pxToMM;
		let topRect = (this.cChartDrawer.calcProp.chartGutter._top) / this.cChartDrawer.calcProp.pxToMM;
		let rightRect = this.cChartDrawer.calcProp.trueWidth / this.cChartDrawer.calcProp.pxToMM;
		let bottomRect = (this.cChartDrawer.calcProp.trueHeight) / this.cChartDrawer.calcProp.pxToMM;

		if (!AscFormat.isRealNumber(leftRect) || !AscFormat.isRealNumber(topRect) || !AscFormat.isRealNumber(rightRect) || !AscFormat.isRealNumber(bottomRect) ) {
			return
		}

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);


		//TODO !!!
		//series color
		/*<cx:plotArea>
		<cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="60000"/>
			<a:lumOff val="40000"/>
			</a:srgbClr>
			</a:solidFill>*/

		//point color
		/*cx:plotAreaRegion>
		<cx:series layoutId="clusteredColumn" uniqueId="{F41B4FED-57DF-4014-AD1A-6DD63A203692}">
			<cx:dataPt idx="2">
			<cx:spPr>
		<a:solidFill>
		<a:srgbClr val="ED7D31">
			<a:lumMod val="40000"/>
			<a:lumOff val="60000"/>
			</a:srgbClr>
			</a:solidFill>
			</cx:spPr>*/

		//this.cChartSpace.chart.plotArea.plotAreaRegion.series[0].spPr.Fill.ln

		let series = this.cChartSpace.chart.plotArea.plotAreaRegion.series;

		let oSeries = series[0];
		if(oSeries) {
			if (this.linePath) {
				let pen = this.cChartSpace.pen;
				this.cChartDrawer.drawPath(this.linePath, pen);
			}
			for (let i in this.paths) {
				if (this.paths.hasOwnProperty(i) && this.paths[i]) {
					let nPtIdx = parseInt(i);
					let pen = oSeries.getPtPen(nPtIdx);
					let brush = oSeries.getPtBrush(nPtIdx);
					this.cChartDrawer.drawPath(this.paths[i], pen, brush);
				}
			}
		}
		
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	}, 

	_calculateDLbl: function (compiledDlb) {
		if (!this.paths || !this.chartProp) {
			return;
		}
		const path = this.paths[compiledDlb.idx];

		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		const oPath = this.cChartSpace.GetPath(path);
		const oCommand0 = oPath.getCommandByIndex(0);
		const oCommand1 = oPath.getCommandByIndex(1);
		const oCommand2 = oPath.getCommandByIndex(2);

		const x = oCommand0.X;
		const y = oCommand0.Y;

		const h = oCommand0.Y - oCommand1.Y;
		const w = oCommand2.X - oCommand1.X;

		const pxToMm = this.chartProp.pxToMM;

		const width = compiledDlb.extX;
		const height = compiledDlb.extY;

		let centerX, centerY;

		centerX = x + w / 2 - width / 2;
		centerY = y - h / 2 - height / 2;
		const top = this.chartProp.chartGutter._top / pxToMm;
		if (centerY < top) {
			centerY = top;
		}

		const bottom = ((this.chartProp.trueHeight + this.chartProp.chartGutter._top) / pxToMm) - height ;
		if (centerY >= bottom) {
			centerY = bottom;
		}

		return {x: centerX, y: centerY};
	}
}

/** @constructor */
function drawLineChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;
	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
}

drawLineChart.prototype = {
	constructor: drawLineChart,

	draw: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			this._drawLines3D();
		} else {
			this._drawLines();
		}
	},

	recalculate: function () {
		this.paths = {};

		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.seriesCount = countSeries.series;
		this.ptCount = countSeries.points;
		this.subType = this.cChartDrawer.getChartGrouping(this.chart);
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);
		if(!this.catAx) {
			this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_DateAx);
		}
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);

		this._calculateLines();
	},

	_calculateLines: function () {
		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;
		let isLog;

		if(!xPoints || !yPoints) {
			return;
		}

		var points, y, x, val, seria, dataSeries, compiledMarkerSize, compiledMarkerSymbol, idx, numCache, idxPoint;
		for (var i = 0; i < this.chart.series.length; i++) {

			seria = this.chart.series[i];
			numCache = this.cChartDrawer.getNumCache(seria.val);
			isLog = this.valAx && this.valAx.scaling ? this.valAx.scaling.logBase : false;

			if (!numCache) {
				continue;
			}

			dataSeries = numCache.pts;

			for (var n = 0; n < numCache.ptCount; n++) {
				//рассчитываем значения
				//используем для поиска n - idx с 0 индексом может не существовать, а точку в нулевой позиции необходимо отрисовать
				val = this._getYVal(n, i);

				x = this.catAx ? this.cChartDrawer.getYPosition(n + 1, this.catAx) : xPoints[n].pos;
				y = this.cChartDrawer.getYPosition(val, this.valAx);

				if (!this.paths.points) {
					this.paths.points = [];
				}
				if (!this.paths.points[i]) {
					this.paths.points[i] = [];
				}

				if (!points) {
					points = [];
				}
				if (!points[i]) {
					points[i] = [];
				}

				idxPoint = this.cChartDrawer.getIdxPoint(seria, n);
				compiledMarkerSize = idxPoint && idxPoint.compiledMarker && idxPoint.compiledMarker.size ? idxPoint.compiledMarker.size : null;
				compiledMarkerSymbol = idxPoint && idxPoint.compiledMarker && AscFormat.isRealNumber(idxPoint.compiledMarker.symbol) ? idxPoint.compiledMarker.symbol : null;

				if (val != null && ((isLog && val !== 0) || !isLog)) {
					this.paths.points[i][n] = this.cChartDrawer.calculatePoint(x, y, compiledMarkerSize, compiledMarkerSymbol);
					let errBars = this.chart.series[i].errBars[0];
					if (errBars) {
						this.cChartDrawer.errBars.putPoint(x, y, val, null,  seria.idx, n);
					}
					points[i][n] = {x: x, y: y};
				} else {
					this.paths.points[i][n] = null;
					points[i][n] = null;
				}
			}
		}

		this._calculateAllLines(points);
	},

	_calculateAllLines: function (points) {
		if(!points) {
			return;
		}

		let chart3d = this.cChartDrawer.nDimensionCount === 3;
		let perspectiveDepth, serDiff, gapDepth, depthSer, diffGapDepth;
		if (chart3d) {
			//сдвиг по OZ в глубину
			perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
			serDiff = perspectiveDepth / this.seriesCount;
			gapDepth = this.chart.gapDepth != null ? this.chart.gapDepth : globalGapDepth;
			depthSer = serDiff / ((gapDepth / 100) + 1);
			diffGapDepth = (depthSer * (gapDepth / 100)) / 2;
			depthSer = (perspectiveDepth / this.seriesCount - 2 * diffGapDepth);
		}

		let x, y, x1, y1, x2, y2, x3, y3, isSplineLine;
		for (let i = 0; i < points.length; i++) {
			isSplineLine = this.chart.series[i].smooth !== false;

			if (!points[i]) {
				continue;
			}

			for (let n = 0; n < points[i].length; n++) {
				if (!this.paths.series) {
					this.paths.series = [];
				}
				if (!this.paths.series[i]) {
					this.paths.series[i] = [];
				}

				let curPoint = points[i][n];
				let nextPoint = AscFormat.DISP_BLANKS_AS_SPAN === this.cChartSpace.chart.dispBlanksAs && !chart3d ? this.cChartDrawer.getNextArrElem(points[i], n) : points[i][n+1];
				if (curPoint != null && nextPoint != null) {
					if (chart3d) {
						x = curPoint.x * this.chartProp.pxToMM;
						y = curPoint.y * this.chartProp.pxToMM;

						x1 = nextPoint.x * this.chartProp.pxToMM;
						y1 = nextPoint.y * this.chartProp.pxToMM;

						if (!this.paths.series) {
							this.paths.series = [];
						}
						if (!this.paths.series[i]) {
							this.paths.series[i] = [];
						}

						let point1, point2, point3, point4, point5, point6, point7, point8, widthLine = 0.5;
						point1 = this.cChartDrawer._convertAndTurnPoint(x, y - widthLine, diffGapDepth + serDiff * i);
						point2 = this.cChartDrawer._convertAndTurnPoint(x1, y1 - widthLine, diffGapDepth + serDiff * i);
						point3 = this.cChartDrawer._convertAndTurnPoint(x1, y1 - widthLine, diffGapDepth + depthSer + serDiff * i);
						point4 = this.cChartDrawer._convertAndTurnPoint(x, y - widthLine, diffGapDepth + depthSer + serDiff * i);


						point5 = this.cChartDrawer._convertAndTurnPoint(x, y + widthLine, diffGapDepth + serDiff * i);
						point6 = this.cChartDrawer._convertAndTurnPoint(x1, y1 + widthLine, diffGapDepth + serDiff * i);
						point7 = this.cChartDrawer._convertAndTurnPoint(x1, y1 + widthLine, diffGapDepth + depthSer + serDiff * i);
						point8 = this.cChartDrawer._convertAndTurnPoint(x, y + widthLine, diffGapDepth + depthSer + serDiff * i);

						this.paths.series[i][n] = this.cChartDrawer.calculateRect3D([point1, point2, point3, point4, point5, point6, point7, point8]);
					} else if (isSplineLine) {
						x = points[i][n - 1] ? n - 1 : 0;
						y = this._getYVal(x, i);

						x1 = n;
						y1 = this._getYVal(x1, i);

						x2 = points[i][n + 1] ? n + 1 : n;
						y2 = this._getYVal(x2, i);

						x3 = points[i][n + 2] ? n + 2 : points[i][n + 1] ? n + 1 : n;
						y3 = this._getYVal(x3, i);

						this.paths.series[i][n] = this.cChartDrawer.calculateSplineLine(x + 1, y, x1 + 1, y1, x2 + 1, y2, x3 + 1, y3, this.catAx, this.valAx);
					} else {
						this.paths.series[i][n] = this._calculateLine(curPoint.x, curPoint.y, nextPoint.x, nextPoint.y);
					}
				}
			}
		}
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);
		var path;

		if(!point) {
			return;
		}

		var commandIndex = 0;
		if (this.cChartDrawer.nDimensionCount === 3) {
			var curSer = this.paths.series[ser];
			if (val === curSer.length && curSer[val - 1]) {
				if (AscFormat.isRealNumber(curSer[val - 1][5])) {
					path = curSer[val - 1][5];
				} else if (AscFormat.isRealNumber(curSer[val - 1][2])) {
					path = curSer[val - 1][2];
				}
				commandIndex = 3;
			} else if (curSer[val] && AscFormat.isRealNumber(curSer[val][3]))//reverse
			{
				path = curSer[val][3];
			} else if (curSer[val] && AscFormat.isRealNumber(curSer[val][2])) {
				path = curSer[val][2];
			}
		} else {
			if(this.paths.points[ser] && this.paths.points[ser][val]){
				path = this.paths.points[ser][val].path;
			}
		}

		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		var oPath = this.cChartSpace.GetPath(path);
		var oCommand0 = oPath.getCommandByIndex(commandIndex);
		var x = oCommand0.X;
		var y = oCommand0.Y;

		var pxToMm = this.chartProp.pxToMM;
		var constMargin = 5 / pxToMm;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x - width / 2;
		var centerY = y - height / 2;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_drawLines: function (/*isSkip*/) {
		//TODO для того, чтобы верхняя линия рисовалась. пересмотреть!
		var diffPen = 3;
		var leftRect = this.chartProp.chartGutter._left / this.chartProp.pxToMM;
		var topRect = (this.chartProp.chartGutter._top - diffPen) / this.chartProp.pxToMM;
		var rightRect = this.chartProp.trueWidth / this.chartProp.pxToMM;
		var bottomRect = (this.chartProp.trueHeight + diffPen) / this.chartProp.pxToMM;

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);
		this.cChartDrawer.drawPaths(this.paths, this.chart.series, true);
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();

		this.cChartDrawer.drawPathsPoints(this.paths, this.chart.series);
	},

	_getYVal: function (n, i) {
		var tempVal;
		var val = 0;
		var idxPoint;
		var k;
		if (this.subType === "stacked") {
			for (k = 0; k <= i; k++) {
				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[k], n);
				tempVal = idxPoint ? parseFloat(idxPoint.val) : 0;
				if (tempVal) {
					val += tempVal;
				}
			}
		} else if (this.subType === "stackedPer") {
			var sumVal = 0;
			for (k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[k], n);
				//TODO сейчас рисуем непрерывную линию, если нужно разорваться - не нужно 0 подставлять
				tempVal = idxPoint ? parseFloat(idxPoint.val) : 0;
				if (tempVal) {
					if (k <= i) {
						val += tempVal;
					}
					sumVal += Math.abs(tempVal);
				}
			}
			if(sumVal === 0) {
				val = 0;
			} else {
				val = val / sumVal;
			}
		} else {
			idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[i], n);
			//TODO blank SPAN option
			var dispBlanksAs =  this.cChartSpace.chart.dispBlanksAs;
			if(idxPoint) {
				val = parseFloat(idxPoint.val);
			} else if(dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
				val = 0;
			} else {
				val = null;
			}
		}

		return val;
	},

	_calculateLine: function (x, y, x1, y1) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		path.moveTo(x * pathW, y * pathH);
		path.lnTo(x1 * pathW, y1 * pathH);


		return pathId;
	},

	//TODO пока включаю функцию _calculateSplineLine. с _calculateSplineLine2 отрисовается неверно. проверить!
	_calculateSplineLine2: function (x, y, x1, y1, x2, y2, x3, y3, xPoints, yPoints) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		var splineCoords = this.cChartDrawer.calculate_Bezier(x, y, x1, y1, x2, y2, x3, y3);

		x = this.cChartDrawer.getYPosition(splineCoords[0].x, this.catAx);
		y = this.cChartDrawer.getYPosition(splineCoords[0].y, this.valAx);

		x1 = this.cChartDrawer.getYPosition(splineCoords[1].x, this.catAx);
		y1 = this.cChartDrawer.getYPosition(splineCoords[1].y, this.valAx);

		x2 = this.cChartDrawer.getYPosition(splineCoords[2].x, this.catAx);
		y2 = this.cChartDrawer.getYPosition(splineCoords[2].y, this.valAx);

		x3 = this.cChartDrawer.getYPosition(splineCoords[3].x, this.catAx);
		y3 = this.cChartDrawer.getYPosition(splineCoords[3].y, this.valAx);

		path.moveTo(x * pathW, y * pathH);
		path.cubicBezTo(x1 * pathW, y1 * pathH, x2 * pathW, y2 * pathH, x3 * pathW, y3 * pathH);

		return pathId;
	},

	_drawLines3D: function () {
		var t = this;

		var drawVerges = function (j, i) {
			var brush, pen, seria;

			seria = t.chart.series[j];
			brush = seria.brush;
			pen = seria.pen;

			var numCache = seria.val && seria.val.getNumCache();
			if (t.paths.series[j] && t.paths.series[j][i] && numCache && numCache.pts[i]) {
				if (numCache.pts && numCache.pts[i] && numCache.pts[i].pen) {
					pen = numCache.pts[i].pen;
				}
				if (numCache.pts && numCache.pts[i] && numCache.pts[i].brush) {
					brush = numCache.pts[i].brush;
				}

				for (var k = 0; k < t.paths.series[j][i].length; k++) {
					t._drawLine3D(t.paths.series[j][i][k], pen, brush, k);
				}
			}
		};


		//рисуем по сериям
		var onSeries = function (onlyLessNull) {
			var drawNeedVerge = function () {
				for (var j = 0; j < t.paths.series.length; j++) {
					for (var i = 0; i < t.ptCount; i++) {
						drawVerges(j, i, onlyLessNull);
					}
				}
			};

			drawNeedVerge();
		};


		var reverseSeriesOnSeries = function (onlyLessNull) {
			var drawNeedVerge = function () {
				for (var j = t.paths.series.length - 1; j >= 0; j--) {
					if (!t.paths.series) {
						return;
					}

					for (var i = 0; i < t.ptCount; i++) {
						drawVerges(j, i, onlyLessNull);
					}
				}
			};

			drawNeedVerge();
		};


		if (!this.cChartDrawer.processor3D.view3D.getRAngAx()) {
			var angle = Math.abs(this.cChartDrawer.processor3D.angleOy);
			if (angle > Math.PI / 2 && angle < 3 * Math.PI / 2) {
				onSeries();
			} else {
				reverseSeriesOnSeries();
			}
		} else {
			reverseSeriesOnSeries();
		}
	},

	_drawLine3D: function (path, pen, brush, k) {
		//затемнение боковых сторон
		//в excel всегда темные боковые стороны, лицевая и задняя стороны светлые

		//todo возможно стоит проверить fill.type на FILL_TYPE_NOFILL и рисовать отдельно границы, если они заданы!
		//brush = pen.Fill;

		if (k !== 2) {
			var props = this.cChartSpace.getParentObjects();
			var duplicateBrush = brush.createDuplicate();
			var cColorMod = new AscFormat.CColorMod;

			cColorMod.name = "shade";
			if (k === 1 || k === 4) {
				cColorMod.val = 45000;
			} else {
				cColorMod.val = 35000;
			}

			this._addColorMods(cColorMod, duplicateBrush);

			duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master, new AscFormat.CUniColor().RGBA, this.cChartSpace.clrMapOvr);
			pen = AscFormat.CreatePenFromParams(duplicateBrush, undefined, undefined, undefined, undefined, 0.1);

			this.cChartDrawer.drawPath(path, pen, duplicateBrush);
		} else {
			pen = AscFormat.CreatePenFromParams(brush, undefined, undefined, undefined, undefined, 0.1);
			this.cChartDrawer.drawPath(path, pen, brush);
		}
	},

	_addColorMods: function (cColorMod, duplicateBrush) {
		if (duplicateBrush) {
			duplicateBrush.addColorMod(cColorMod);
		}
	}
};


/** @constructor */
function drawAreaChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;
	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.points = null;
	this.paths = {};
	this.upFaces = [];
	this.downFaces = [];

	this.sortZIndexPaths = [];
	this.sortZIndexPathsFront = [];
	this.sortZIndexPathsBack = [];
	this.sortZIndexPathsRight = [];
	this.sortZIndexPathsLeft = [];

	this.intersections = [];
	this.intersectionsFar = [];

	this.xPoints = null;
	this.yPoints = null;

	//for 3d
	this.darkFaces = null;
	this.gapDepth = null;
	this.perspectiveDepth = null;
}

drawAreaChart.prototype = {
	constructor: drawAreaChart,

	draw: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			this._drawBars3D();
		} else {
			this._drawLines();
		}
	},

	recalculate: function () {
		this.paths = {};
		this.points = null;

		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.seriesCount = countSeries.series;
		this.ptCount = countSeries.points;
		this.subType = this.cChartDrawer.getChartGrouping(this.chart);
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);
		if(!this.catAx) {
			this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_DateAx);
		}
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);
		this.serAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_SerAx);

		this._calculateProps();
		this._calculate();
	},

	_calculateProps: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			var gapDepth = this.chart.gapDepth != null ? this.chart.gapDepth : globalGapDepth;
			var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
			perspectiveDepth = this.subType === "normal" ? (perspectiveDepth / (gapDepth / 100 + 1)) / this.seriesCount : perspectiveDepth / (gapDepth / 100 + 1);
			var DiffGapDepth = perspectiveDepth * (gapDepth / 2) / 100;

			this.perspectiveDepth = perspectiveDepth;
			if (this.subType === "normal") {
				this.gapDepth = [];

				for (var i = 0; i < this.seriesCount; i++) {
					if (this.serAx && this.serAx.zPoints[i]) {
						this.gapDepth[i] = this.serAx.zPoints[i].pos - this.perspectiveDepth / 2;
					} else {
						this.gapDepth[i] = (this.perspectiveDepth + DiffGapDepth + DiffGapDepth) * i + DiffGapDepth;
					}
				}
			} else {
				this.gapDepth = DiffGapDepth;
			}
		}

		this.xPoints = this.catAx.xPoints;
		this.yPoints = this.valAx.yPoints;
	},

	_calculate: function () {
		var y, x, val, seria, dataSeries, numCache;
		var pxToMm = this.chartProp.pxToMM;
		var nullPositionOX = this.catAx.posY;
		let isLog;

		for (var i = 0; i < this.chart.series.length; i++) {

			seria = this.chart.series[i];
			numCache = this.cChartDrawer.getNumCache(seria.val);
			isLog = this.valAx && this.valAx.scaling ? this.valAx.scaling.logBase : false;

			if (!numCache) {
				continue;
			}

			dataSeries = numCache.pts;

			for (var n = 0; n < numCache.ptCount; n++) {
				//рассчитываем значения
				val = this._getYVal(n, i);

				if((null === val && this.cChartDrawer.nDimensionCount !== 3) || (isLog && val === 0)) {
					continue;
				}

				x = this.xPoints[n].pos;
				y = this.cChartDrawer.getYPosition(val, this.valAx);

				if (!this.points) {
					this.points = [];
				}
				if (!this.points[i]) {
					this.points[i] = [];
				}

				if (val != null) {
					this.points[i][n] = {x: x, y: y, val: val};
				} else {
					this.points[i][n] = {x: x, y: nullPositionOX, val: val};
				}

				let errBars = seria.errBars[0];
				if (errBars && this.chartProp && this.chartProp.pxToMM) {
					const idx = AscFormat.isRealNumber(dataSeries[n].idx) ? dataSeries[n].idx : n;
					var _pointVal = this.cChartDrawer.getValWithStacked(i, idx, this.chart);
					this.cChartDrawer.errBars.putPoint(x, y, _pointVal, _pointVal,  seria.idx, idx);
				}

			}

			if (this.cChartDrawer.nDimensionCount === 3) {
				//для normal рассчитываем видимые/невидимые грани для каждой серии
				if (this.subType === "normal") {
					this._calculateDarkSideOfTheFace(i);
				} else if (this.darkFaces === null) {
					this._calculateDarkSideOfTheFace(null);
				}
			}
		}


		if (this.cChartDrawer.nDimensionCount === 3) {
			this._calculatePaths3D();

			//var cSortFaces = new CSortFaces(this.cChartDrawer);
			//this.upFaces = cSortFaces.sortFaces(this.upFaces);

			var upDownFaces = this.upFaces.concat(this.downFaces);

			//более быстрая сортировка
			var angle = this.cChartDrawer.processor3D.angleOx;
			if (angle < 0) {
				upDownFaces.sort(function sortArr(a, b) {
					if (b.midY === a.midY) {
						return a.seria - b.seria;
					} else {
						return b.midY - a.midY;
					}
				});
			} else {
				upDownFaces.sort(function sortArr(a, b) {
					if (b.midY === a.midY) {
						return a.seria - b.seria;
					} else {
						return a.midY - b.midY;
					}
				});
			}

			var anotherFaces = this.sortZIndexPathsFront.concat(this.sortZIndexPathsBack).concat(this.sortZIndexPathsLeft).concat(this.sortZIndexPathsRight);
			this.sortZIndexPaths = upDownFaces.concat(anotherFaces);

			//медленная, но более качественный сортировка
			//var anotherFaces = this.sortZIndexPathsFront.concat(this.sortZIndexPathsBack).concat(this.sortZIndexPathsLeft).concat(this.sortZIndexPathsRight);
			//this.sortZIndexPaths = this.upFaces.concat(anotherFaces);
			//this.sortZIndexPaths = this.downFaces.concat(this.sortZIndexPaths)
			//this.sortZIndexPaths = cSortFaces.sortFaces(this.sortZIndexPaths);
		} else {
			this._calculatePaths();
		}
	},

	_calculatePaths: function () {
		let points = this.points;
		let isStacked = this.subType === "stackedPer" || this.subType === "stacked";

		for (let i = 0; i < points.length; i++) {
			if (!this.paths.series) {
				this.paths.series = [];
			}
			if (points[i]) {
				this.paths.series[i] = this._calculateLine(points[i], isStacked ? i : null);
			}
		}
	},

	_calculatePaths3D: function () {
		var points = this.points;
		var prevPoints;
		var isStacked = this.subType === "stackedPer" || this.subType === "stacked";

		if (isStacked) {
			this._calculateAllIntersection();
		}

		for (var i = 0; i < points.length; i++) {
			if (!this.paths.series) {
				this.paths.series = [];
			}

			prevPoints = isStacked ? this._getPrevSeriesPoints(points, i) : null;

			if (points[i]) {
				if (isStacked) {
					this.paths.series[i] = this._calculateStacked3D(prevPoints, i, points);
				} else {
					this.paths.series[i] = this._calculateLine3D(points[i], i, points);
				}
			}
		}
	},

	_calculateLine: function (points, seriesIndex) {
		if (!points) {
			return;
		}
		let allPoints = this.points;

		let pathId = this.cChartSpace.AllocPath();
		let path = this.cChartSpace.GetPath(pathId);

		if (!path) {
			return;
		}

		let pathH = this.chartProp.pathH;
		let pathW = this.chartProp.pathW;


		let point;
		let pxToMm = this.chartProp.pxToMM;

		let nullPositionOX = this.catAx && this.catAx.posY * this.chartProp.pxToMM;

		let isPrevPoints = seriesIndex !== null && seriesIndex !== 0;
		let t = this;
		let searchPreviousPoint = function (pointIndex) {
			for (let i = seriesIndex; i >= 0; i--) {
				let prevPoints = t._getPrevSeriesPoints(allPoints, i);
				if (prevPoints && prevPoints[pointIndex]) {
					return prevPoints[pointIndex];
				}
			}
			return null;
		};

		//точки данной серии
		if(this.subType === "stacked" || this.subType === "stackedPer") {
			let pointsLength = points.length;
			for (let i = 0; i < pointsLength; i++) {
				point = points[i];
				if(!point) {

				} else if (i === 0) {
					path.moveTo(point.x * pathW, point.y * pathH);
				} else {
					path.lnTo(point.x * pathW, point.y * pathH);
				}
			}

			//точки предыдущей серии
			if (isPrevPoints && pointsLength) {
				for (let i = pointsLength - 1; i >= 0; i--) {
					point = searchPreviousPoint(i);
					if (!point) {
						continue;
					}
					path.lnTo(point.x * pathW, point.y * pathH);

					if (i === 0) {
						path.lnTo(points[0].x * pathW, points[0].y * pathH);
					}
				}
			} else {
				path.lnTo(points[points.length - 1].x * pathW, nullPositionOX / pxToMm * pathH);
				path.lnTo(points[0].x * pathW, nullPositionOX / pxToMm * pathH);
				path.lnTo(points[0].x * pathW, points[0].y * pathH);
			}
		} else {
			let startSegmentPoint;
			for (let i = 0; i < points.length; i++) {
				point = points[i];
				if(!point) {
					if(startSegmentPoint) {
						//возвращаемся к оси, далее к начальной точки сегмента
						path.lnTo(points[i - 1].x * pathW, nullPositionOX / pxToMm * pathH);
						path.lnTo(startSegmentPoint.x * pathW, nullPositionOX / pxToMm * pathH);
						path.lnTo(startSegmentPoint.x * pathW, startSegmentPoint.y * pathH);
					}
					startSegmentPoint = null;
				} else {
					if(!startSegmentPoint) {
						startSegmentPoint = point;
						path.moveTo(point.x * pathW, nullPositionOX / pxToMm * pathH);
					}
					path.lnTo(point.x * pathW, point.y * pathH);
				}

				if(i === points.length - 1 && point) {
					if(startSegmentPoint) {
						//возвращаемся к оси, далее к начальной точки сегмента
						path.lnTo(point.x * pathW, nullPositionOX / pxToMm * pathH);
						path.lnTo(startSegmentPoint.x * pathW, nullPositionOX / pxToMm * pathH);
						path.lnTo(startSegmentPoint.x * pathW, startSegmentPoint.y * pathH);
					}
				}
			}
		}

		return pathId;
	},

	_calculateLine3D: function (points, seriaNum) {
		//pointsIn3D[0] - верхняя кривая ближней стороны, pointsIn3D[1] - нижняя кривая ближней стороны, pointsIn3D[2] - верхняя кривая дальней стороны, pointsIn3D[3] - нижняя кривая дальней стороны
		var pointsIn3D = [], t = this, pxToMm = this.chartProp.pxToMM;
		var nullPosition = this.catAx.posY * this.chartProp.pxToMM;

		//сдвиг по OZ в глубину
		var gapDepth = this.gapDepth[seriaNum];

		var getProjectPoints = function (currentZ, startN) {
			pointsIn3D[startN] = [];
			pointsIn3D[startN + 1] = [];
			for (var i = 0; i < points.length; i++) {
				pointsIn3D[startN][i] = t.cChartDrawer._convertAndTurnPoint(points[i].x * pxToMm, points[i].y * pxToMm, currentZ + gapDepth);
				pointsIn3D[startN + 1][i] = t.cChartDrawer._convertAndTurnPoint(points[i].x * pxToMm, nullPosition, currentZ + gapDepth);
			}
		};

		var zNear = 0;
		var zFar = this.perspectiveDepth;
		//рассчитываем ближние и дальние точки конкретной серии
		getProjectPoints(zNear, 0);
		getProjectPoints(zFar, 2);

		return this._calculateRect3D(pointsIn3D, seriaNum);
	},

	_calculateStacked3D: function (prevPoints, seria, allPoints) {
		var points = allPoints[seria];
		var t = this, nullPositionOX = this.catAx.posY * this.chartProp.pxToMM, pxToMm = this.chartProp.pxToMM, res = [];

		for (var i = 0; i < points.length - 1; i++) {
			var x = points[i].x * pxToMm;
			var y = points[i].y * pxToMm;
			var x1 = points[i + 1].x * pxToMm;
			var y1 = points[i + 1].y * pxToMm;

			//рассчитываем 8 точек для каждого ректа
			var prevX = prevPoints ? prevPoints[i].x * pxToMm : x;
			var prevY = prevPoints ? prevPoints[i].y * pxToMm : nullPositionOX;
			var prevX1 = prevPoints ? prevPoints[i + 1].x * pxToMm : x1;
			var prevY1 = prevPoints ? prevPoints[i + 1].y * pxToMm : nullPositionOX;


			if (prevY < y && prevY1 < y1) {
				var temp = y;
				y = prevY;
				prevY = temp;

				var temp1 = y1;
				y1 = prevY1;
				prevY1 = temp1;

				temp = x;
				x = prevX;
				prevX = temp;

				temp1 = x1;
				x1 = prevX1;
				prevX1 = temp1;
			}

			if (!this.prevPoints) {
				this.prevPoints = [];
			}

			if (!this.prevPoints[i]) {
				this.prevPoints[i] = [];
			}


			//начальные точки, без проекции и без поворота
			var p1 = {x: x, y: y, z: this.gapDepth};
			var p2 = {x: x, y: y, z: this.gapDepth + this.perspectiveDepth};
			var p3 = {x: x1, y: y1, z: this.gapDepth + this.perspectiveDepth};
			var p4 = {x: x1, y: y1, z: this.gapDepth};
			var p5 = {x: prevX, y: prevY, z: this.gapDepth};
			var p6 = {x: prevX, y: prevY, z: this.gapDepth + this.perspectiveDepth};
			var p7 = {x: prevX1, y: prevY1, z: this.gapDepth + this.perspectiveDepth};
			var p8 = {x: prevX1, y: prevY1, z: this.gapDepth};
			var arrNotRotatePoints = [p1, p2, p3, p4, p5, p6, p7, p8];

			//повернутые, но без проекции
			var point11 = t.cChartDrawer._convertAndTurnPoint(p1.x, p1.y, p1.z, null, null, true);
			var point22 = t.cChartDrawer._convertAndTurnPoint(p2.x, p2.y, p2.z, null, null, true);
			var point33 = t.cChartDrawer._convertAndTurnPoint(p3.x, p3.y, p3.z, null, null, true);
			var point44 = t.cChartDrawer._convertAndTurnPoint(p4.x, p4.y, p4.z, null, null, true);
			var point55 = t.cChartDrawer._convertAndTurnPoint(p5.x, p5.y, p5.z, null, null, true);
			var point66 = t.cChartDrawer._convertAndTurnPoint(p6.x, p6.y, p6.z, null, null, true);
			var point77 = t.cChartDrawer._convertAndTurnPoint(p7.x, p7.y, p7.z, null, null, true);
			var point88 = t.cChartDrawer._convertAndTurnPoint(p8.x, p8.y, p8.z, null, null, true);
			var arrPoints = [point11, point22, point33, point44, point55, point66, point77, point88];

			//спроецированные и повернутые точки
			var point1 = t.cChartDrawer._convertAndTurnPoint(p1.x, p1.y, p1.z);
			var point2 = t.cChartDrawer._convertAndTurnPoint(p2.x, p2.y, p2.z);
			var point3 = t.cChartDrawer._convertAndTurnPoint(p3.x, p3.y, p3.z);
			var point4 = t.cChartDrawer._convertAndTurnPoint(p4.x, p4.y, p4.z);
			var point5 = t.cChartDrawer._convertAndTurnPoint(p5.x, p5.y, p5.z);
			var point6 = t.cChartDrawer._convertAndTurnPoint(p6.x, p6.y, p6.z);
			var point7 = t.cChartDrawer._convertAndTurnPoint(p7.x, p7.y, p7.z);
			var point8 = t.cChartDrawer._convertAndTurnPoint(p8.x, p8.y, p8.z);
			var arrPointsProject = [point1, point2, point3, point4, point5, point6, point7, point8];

			var paths = this._calculateRect3DStacked(arrPoints, arrPointsProject, arrNotRotatePoints, i, seria);
			res.push(paths);
		}

		return res;
	},

	_calculateDarkSideOfTheFace: function (seria) {
		var pxToMm = this.chartProp.pxToMM;

		var minX = this.xPoints[0].pos < this.xPoints[this.xPoints.length - 1].pos ? this.xPoints[0].pos * pxToMm : this.xPoints[this.xPoints.length - 1].pos * pxToMm;
		var maxX = this.xPoints[0].pos < this.xPoints[this.xPoints.length - 1].pos ? this.xPoints[this.xPoints.length - 1].pos * pxToMm : this.xPoints[0].pos * pxToMm;

		var minValue = this.valAx.min;
		var maxValue = this.valAx.max;
		if (this.subType === "stackedPer") {
			minValue = minValue / 100;
			maxValue = maxValue / 100;
		}

		var maxY = this.cChartDrawer.getYPosition(minValue, this.valAx) * pxToMm;
		var minY = this.cChartDrawer.getYPosition(maxValue, this.valAx) * pxToMm;

		var gapDepth = seria === null ? this.gapDepth : this.gapDepth[seria];

		var point1 = this.cChartDrawer._convertAndTurnPoint(minX, maxY, gapDepth);
		var point2 = this.cChartDrawer._convertAndTurnPoint(minX, maxY, gapDepth + this.perspectiveDepth);
		var point3 = this.cChartDrawer._convertAndTurnPoint(maxX, maxY, gapDepth + this.perspectiveDepth);
		var point4 = this.cChartDrawer._convertAndTurnPoint(maxX, maxY, gapDepth);

		var point5 = this.cChartDrawer._convertAndTurnPoint(minX, minY, gapDepth);
		var point6 = this.cChartDrawer._convertAndTurnPoint(minX, minY, gapDepth + this.perspectiveDepth);
		var point7 = this.cChartDrawer._convertAndTurnPoint(maxX, minY, gapDepth + this.perspectiveDepth);
		var point8 = this.cChartDrawer._convertAndTurnPoint(maxX, minY, gapDepth);

		var darkFaces;
		if (seria === null) {
			this.darkFaces = {};
			darkFaces = this.darkFaces;
		} else {
			if (null === this.darkFaces) {
				this.darkFaces = [];
			}
			this.darkFaces[seria] = {};
			darkFaces = this.darkFaces[seria];
		}


		//front
		darkFaces["front"] = null;
		if (this._isVisibleVerge3D(point5, point1, point4)) {
			darkFaces["front"] = 1;
		}

		//down
		darkFaces["down"] = null;
		if (this._isVisibleVerge3D(point4, point1, point2)) {
			darkFaces["down"] = 1;
		}

		//left
		darkFaces["left"] = null;
		if (this._isVisibleVerge3D(point2, point1, point5, true)) {
			darkFaces["left"] = 1;
		}

		//right
		darkFaces["right"] = null;
		if (this._isVisibleVerge3D(point8, point4, point3, true)) {
			darkFaces["right"] = 1;
		}

		//up
		darkFaces["up"] = null;
		if (this._isVisibleVerge3D(point6, point5, point8)) {
			darkFaces["up"] = 1;
		}

		//unfront
		darkFaces["unfront"] = null;
		if (this._isVisibleVerge3D(point3, point2, point6)) {
			darkFaces["unfront"] = 1;
		}
	},

	_calculateRect3D: function (pointsIn3D, seriaNum) {
		var path, pathId;
		var pxToMm = this.chartProp.pxToMM;
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var oThis = this;

		var calculateFace = function (number, isReverse, isFirstPoint) {
			if (!isReverse) {
				for (var i = 0; i < pointsIn3D[number].length; i++) {
					if (i === 0 && isFirstPoint) {
						path.moveTo(pointsIn3D[number][i].x / pxToMm * pathW,
							pointsIn3D[number][i].y / pxToMm * pathH);
					} else {
						path.lnTo(pointsIn3D[number][i].x / pxToMm * pathW,
							pointsIn3D[number][i].y / pxToMm * pathH);
					}
				}
			} else {
				for (var i = pointsIn3D[number].length - 1; i >= 0; i--) {
					if (i === pointsIn3D[number].length - 1 && isFirstPoint) {
						path.moveTo(pointsIn3D[number][i].x / pxToMm * pathW,
							pointsIn3D[number][i].y / pxToMm * pathH);
					} else {
						path.lnTo(pointsIn3D[number][i].x / pxToMm * pathW,
							pointsIn3D[number][i].y / pxToMm * pathH);
					}
				}
			}
		};

		var calculateRect = function (p1, p2, p3, p4) {
			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			path.moveTo(p1.x / pxToMm * pathW, p1.y / pxToMm * pathH);
			path.lnTo(p2.x / pxToMm * pathW, p2.y / pxToMm * pathH);
			path.lnTo(p3.x / pxToMm * pathW, p3.y / pxToMm * pathH);
			path.lnTo(p4.x / pxToMm * pathW, p4.y / pxToMm * pathH);
			path.lnTo(p1.x / pxToMm * pathW, p1.y / pxToMm * pathH);

			return pathId;
		};

		var calculateUpDownFace = function (isDown) {
			var arrayPaths = [];
			for (var i = 0; i < pointsIn3D[0].length - 1; i++) {
				var point1Up = pointsIn3D[0][i];
				var point2Up = pointsIn3D[0][i + 1];

				var point1UpFar = pointsIn3D[2][i];
				var point2UpFar = pointsIn3D[2][i + 1];

				var point1Down = pointsIn3D[1][i];
				var point2Down = pointsIn3D[1][i + 1];

				var point1DownFar = pointsIn3D[3][i];
				var point2DownFar = pointsIn3D[3][i + 1];

				var path = null;

				if (!isDown) {
					if (point1Up.y > point1Down.y && point2Up.y < point2Down.y) {
						path = calculateRect(point1Up, point1UpFar, point2UpFar, point2Up);
						arrayPaths.push(path);
						path = calculateRect(point1Down, point1DownFar, point2DownFar, point2Down);
						arrayPaths.push(path);
					} else if (point1Up.y < point1Down.y && point2Up.y > point2Down.y) {
						path = calculateRect(point1Down, point1DownFar, point2DownFar, point2Down);
						arrayPaths.push(path);
						path = calculateRect(point1Up, point1UpFar, point2UpFar, point2Up);
						arrayPaths.push(path);
					} else {
						path = calculateRect(point1Up, point1UpFar, point2UpFar, point2Up);
						arrayPaths.push(path);
					}
				} else {
					if (point1Up.y > point1Down.y && point2Up.y < point2Down.y) {
						path = calculateRect(point1Up, point1UpFar, point2UpFar, point2Up);
						arrayPaths.push(path);
						path = calculateRect(point1Down, point1DownFar, point2DownFar, point2Down);
						arrayPaths.push(path);
					} else if (point1Up.y < point1Down.y && point2Up.y > point2Down.y) {
						path = calculateRect(point1Up, point1UpFar, point2UpFar, point2Up);
						arrayPaths.push(path);
						path = calculateRect(point1Down, point1DownFar, point2DownFar, point2Down);
						arrayPaths.push(path);
					} else {
						path = calculateRect(point1Down, point1DownFar, point2DownFar, point2Down);
						arrayPaths.push(path);
					}
				}
			}

			return arrayPaths;
		};

		var paths = [];
		var upNear = 0, downNear = 1, upFar = 2, downFar = 3;

		//for define VisibleVerge, as in bar charts
		var point1 = pointsIn3D[downNear][0];
		var point2 = pointsIn3D[downFar][0];
		var point3 = pointsIn3D[downFar][pointsIn3D[downFar].length - 1];
		var point4 = pointsIn3D[downNear][pointsIn3D[downNear].length - 1];

		var point5 = pointsIn3D[upNear][0];
		var point6 = pointsIn3D[upFar][0];
		var point7 = pointsIn3D[upFar][pointsIn3D[upFar].length - 1];
		var point8 = pointsIn3D[upNear][pointsIn3D[upNear].length - 1];


		//front
		paths[0] = null;
		if (this.darkFaces[seriaNum]["front"]) {
			pathId = this.cChartSpace.AllocPath();
			path = this.cChartSpace.GetPath(pathId);


			calculateFace(upNear, false, true);
			calculateFace(downNear, true);
			path.lnTo(point5.x / pxToMm * pathW, point5.y / pxToMm * pathH);

			paths[0] = pathId;
		}

		//down
		paths[1] = null;
		var arrayPaths = calculateUpDownFace(true);
		paths[1] = arrayPaths;

		//left
		paths[2] = null;
		if (this.darkFaces[seriaNum]["left"]) {
			arrayPaths = calculateRect(pointsIn3D[0][0], pointsIn3D[1][0], pointsIn3D[3][0], pointsIn3D[2][0]);
			paths[2] = arrayPaths;
		}

		//right
		paths[3] = null;
		if (this.darkFaces[seriaNum]["right"]) {
			arrayPaths = calculateRect(pointsIn3D[0][pointsIn3D[0].length - 1], pointsIn3D[1][pointsIn3D[1].length - 1], pointsIn3D[3][pointsIn3D[3].length - 1], pointsIn3D[2][pointsIn3D[2].length - 1]);
			paths[3] = arrayPaths;
		}

		//up
		paths[4] = null;
		//if(this.darkFaces[seriaNum]["up"])
		//{
		arrayPaths = calculateUpDownFace();
		paths[4] = arrayPaths;
		//}

		//unfront
		paths[5] = null;
		if (this.darkFaces[seriaNum]["unfront"]) {
			pathId = this.cChartSpace.AllocPath();
			path = this.cChartSpace.GetPath(pathId);
			calculateFace(upFar, false, true);
			calculateFace(downFar, true);
			path.lnTo(pointsIn3D[upFar][0].x / pxToMm * pathW, pointsIn3D[upFar][0].y / pxToMm * pathH);
			paths[5] = pathId;
		}

		return paths;
	},

	_calculateRect3DStacked: function (arrPoints, arrPointsProject, arrNotRotatePoints, point, seria) {
		var pxToMm = this.chartProp.pxToMM, t = this;

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		//for define VisibleVerge, as in bar charts
		var p1 = arrNotRotatePoints[0];
		var p2 = arrNotRotatePoints[1];
		var p3 = arrNotRotatePoints[2];
		var p4 = arrNotRotatePoints[3];
		var p5 = arrNotRotatePoints[4];
		var p6 = arrNotRotatePoints[5];
		var p7 = arrNotRotatePoints[6];
		var p8 = arrNotRotatePoints[7];

		var point1 = arrPointsProject[0];
		var point2 = arrPointsProject[1];
		var point3 = arrPointsProject[2];
		var point4 = arrPointsProject[3];
		var point5 = arrPointsProject[4];
		var point6 = arrPointsProject[5];
		var point7 = arrPointsProject[6];
		var point8 = arrPointsProject[7];

		var point11 = arrPoints[0];
		var point22 = arrPoints[1];
		var point33 = arrPoints[2];
		var point44 = arrPoints[3];
		var point55 = arrPoints[4];
		var point66 = arrPoints[5];
		var point77 = arrPoints[6];
		var point88 = arrPoints[7];

		var oThis = this;
		var generateFace = function (p11, p22, p33, p44, p111, p222, p333, p444, p1111, p2222, p3333, p4444,
									 faceIndex, ser) {
			if (ser === undefined) {
				ser = seria;
			}

			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			path.moveTo(p11.x / pxToMm * pathW, p11.y / pxToMm * pathH);
			path.lnTo(p22.x / pxToMm * pathW, p22.y / pxToMm * pathH);
			path.lnTo(p33.x / pxToMm * pathW, p33.y / pxToMm * pathH);
			path.lnTo(p44.x / pxToMm * pathW, p44.y / pxToMm * pathH);
			path.lnTo(p11.x / pxToMm * pathW, p11.y / pxToMm * pathH);

			var arrPoints = [p11, p22, p33, p44];
			var arrPoints2 = [p111, p222, p333, p444];
			var plainEquation = t.cChartDrawer.getPlainEquation(p111, p222, p333);
			var plainArea = t.cChartDrawer.getAreaQuadrilateral(p11, p22, p33, p44);

			var arrPointsNotRotate = [p1111, p2222, p3333, p4444];
			var midY = (p1111.y + p2222.y + p3333.y + p4444.y) / 4;

			return {
				seria: ser,
				point: 0,
				verge: faceIndex,
				paths: pathId,
				points: arrPoints2,
				points2: arrPoints,
				plainEquation: plainEquation,
				plainArea: plainArea,
				arrPointsNotRotate: arrPointsNotRotate,
				midY: midY
			};
		};

		//рассчитываем все грани, кроме верхних и нижних
		this._calculateSimpleRect(arrPoints, arrPointsProject, point, seria);

		//находим пересечение грани с предыдущими гранями. если есть, то делим грани
		var breakFaces = this.intersections[point] && this.intersections[point][seria] ? this.intersections[point][seria] : null;

		if (breakFaces && breakFaces.up && breakFaces.up.length) {
			//сортируем грани одной точки
			breakFaces.up = breakFaces.up.sort(function sortArr(a, b) {
				return a.x - b.x;
			});

			var prevNear, prevFar, prevNearProject, prevFarProject, prevNotRotateNear, prevNotRotateFar, prevPoint;
			var near, far, nearProject, farProject, notRotateNear, notRotateFar, face;
			for (var i = 0; i < breakFaces.up.length - 1; i++) {

				prevPoint = breakFaces.up[i];
				prevNearProject = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth /*+ DiffGapDepth*/);
				prevFarProject = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth + this.perspectiveDepth);

				prevNear = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth /*+ DiffGapDepth*/, null, null, true);
				prevFar = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth + this.perspectiveDepth, null, null, true);

				prevNotRotateNear = {x: prevPoint.x, y: prevPoint.y, z: this.gapDepth};
				prevNotRotateFar = {x: prevPoint.x, y: prevPoint.y, z: this.gapDepth + this.perspectiveDepth};

				point = breakFaces.up[i + 1];
				nearProject = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth /*+ DiffGapDepth*/);
				farProject = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth + this.perspectiveDepth);

				near = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth /*+ DiffGapDepth*/, null, null, true);
				far = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth + this.perspectiveDepth, null, null, true);

				notRotateNear = {x: point.x, y: point.y, z: this.gapDepth};
				notRotateFar = {x: point.x, y: point.y, z: this.gapDepth + this.perspectiveDepth};


				face = generateFace(prevNearProject, prevFarProject, farProject, nearProject, prevNear, prevFar, far, near, prevNotRotateNear, prevNotRotateFar, notRotateNear, notRotateFar, 1);
				this.upFaces.push(face);
			}
		}

		if (breakFaces && breakFaces.down && breakFaces.down.length) {
			//сортируем грани одной точки
			breakFaces.down = breakFaces.down.sort(function sortArr(a, b) {
				return a.x - b.x;
			});

			for (var i = 0; i < breakFaces.down.length - 1; i++) {
				prevPoint = breakFaces.down[i];
				prevNearProject =
					t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth /*+ DiffGapDepth*/);
				prevFarProject = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y,
					this.gapDepth + this.perspectiveDepth);

				prevNear =
					t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y, this.gapDepth /*+ DiffGapDepth*/,
						null, null, true);
				prevFar = t.cChartDrawer._convertAndTurnPoint(prevPoint.x, prevPoint.y,
					this.gapDepth + this.perspectiveDepth, null, null, true);

				prevNotRotateNear = {x: prevPoint.x, y: prevPoint.y, z: this.gapDepth};
				prevNotRotateFar = {x: prevPoint.x, y: prevPoint.y, z: this.gapDepth + this.perspectiveDepth};

				point = breakFaces.down[i + 1];
				nearProject =
					t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth /*+ DiffGapDepth*/);
				farProject =
					t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth + this.perspectiveDepth);

				near = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth /*+ DiffGapDepth*/, null,
					null, true);
				far = t.cChartDrawer._convertAndTurnPoint(point.x, point.y, this.gapDepth + this.perspectiveDepth,
					null, null, true);

				notRotateNear = {x: point.x, y: point.y, z: this.gapDepth};
				notRotateFar = {x: point.x, y: point.y, z: this.gapDepth + this.perspectiveDepth};


				face = generateFace(prevNearProject, prevFarProject, farProject, nearProject, prevNear, prevFar, far, near, prevNotRotateNear, prevNotRotateFar, notRotateNear, notRotateFar, 1);
				this.downFaces.push(face);
			}
		} else {
			//TODO проверить и убрать
			face = generateFace(point1, point2, point3, point4, point11, point22, point33, point44, p1, p2, p3, p4, 1);
			this.downFaces.push(face);
		}

		return [];
	},

	_calculateAllIntersection: function () {
		var allPoints = this.points;
		var prevPoints;
		var nullPositionOX = this.catAx.posY * this.chartProp.pxToMM;
		var pxToMm = this.chartProp.pxToMM;

		for (var seria = 0; seria < allPoints.length; seria++) {
			var points = allPoints[seria];

			if(!points) {
				continue;
			}

			for (var i = 0; i < points.length - 1; i++) {
				var x = points[i].x * pxToMm;
				var y = points[i].y * pxToMm;
				var x1 = points[i + 1].x * pxToMm;
				var y1 = points[i + 1].y * pxToMm;

				prevPoints = this._getPrevSeriesPoints(allPoints, seria);

				//рассчитываем 8 точек для каждого ректа
				var prevX = prevPoints ? prevPoints[i].x * pxToMm : x;
				var prevY = prevPoints ? prevPoints[i].y * pxToMm : nullPositionOX;
				var prevX1 = prevPoints ? prevPoints[i + 1].x * pxToMm : x1;
				var prevY1 = prevPoints ? prevPoints[i + 1].y * pxToMm : nullPositionOX;

				if (!this.prevPoints) {
					this.prevPoints = [];
				}

				if (!this.prevPoints[i]) {
					this.prevPoints[i] = [];
				}


				var curRect = {
					point1: {x: x, y: y, z: this.gapDepth},
					point2: {x: x1, y: y1, z: this.gapDepth},
					prevPoint1: {x: prevX, y: prevY, z: this.gapDepth},
					prevPoint2: {x: prevX1, y: prevY1, z: this.gapDepth}
				};

				this._checkIntersection(curRect, i, seria);
				this.prevPoints[i][seria] = curRect;
			}
		}
	},

	_checkIntersection: function (curRect, pointIndex, seriaIndex) {
		var t = this;

		var curPoint1 = curRect.point1;
		var curPoint2 = curRect.point2;
		var curPoint3 = curRect.prevPoint1;
		var curPoint4 = curRect.prevPoint2;

		if (curPoint1.y > curPoint3.y) {
			curPoint1 = curRect.prevPoint1;
			curPoint2 = curRect.prevPoint2;
			curPoint3 = curRect.point1;
			curPoint4 = curRect.point2;
		}

		var addToArr = function (point, seria, isDown, elem) {
			if (!t.intersections[point]) {
				t.intersections[point] = [];
			}

			if (!t.intersections[point][seria]) {
				t.intersections[point][seria] = {};
			}

			if (isDown && !t.intersections[point][seria].down) {
				t.intersections[point][seria].down = [];
			}
			if (!isDown && !t.intersections[point][seria].up) {
				t.intersections[point][seria].up = [];
			}

			var arr = isDown ? t.intersections[point][seria].down : t.intersections[point][seria].up;
			arr.push(elem);
		};

		//текущая верхняя и нижняя прямая и их пересечение
		var curLine1 = this.cChartDrawer.getLineEquation(curPoint1, curPoint2);
		var curLine2 = this.cChartDrawer.getLineEquation(curPoint3, curPoint4);
		var curTempIntersection = this.cChartDrawer.isIntersectionLineAndLine(curLine1, curLine2);

		var curDown = [{start: curPoint3, end: curPoint4, eq: curLine2}];
		//если пересечения текущих прямых вписывается в границы диаграммы
		var curIntersection = null;
		if (curTempIntersection && curTempIntersection.x > curPoint3.x && curTempIntersection.x < curPoint4.x) {
			curIntersection = curTempIntersection;

			curDown[0] = {start: curPoint3, end: curIntersection, eq: curLine2};
			curDown[1] = {start: curIntersection, end: curPoint4, eq: curLine1};
		}

		var curUp = [{start: curPoint1, end: curPoint2, eq: curLine1}];
		//если пересечения текущих прямых вписывается в границы диаграммы
		curIntersection = null;
		if (curTempIntersection && curTempIntersection.x > curPoint1.x && curTempIntersection.x < curPoint2.x) {
			curIntersection = curTempIntersection;

			curUp[0] = {start: curPoint1, end: curIntersection, eq: curLine1};
			curUp[1] = {start: curIntersection, end: curPoint2, eq: curLine2};
		}


		//первая/последняя точка текущей грани, пересечение текущей грани
		if (!this.prevPoints[pointIndex].length ||
			(this.prevPoints[pointIndex].length && !this.prevPoints[pointIndex][seriaIndex])) {
			//заносим первую точку грани - ПЕРЕПРОВеРИТЬ
			addToArr(pointIndex, seriaIndex, true, curPoint3);
			addToArr(pointIndex, seriaIndex, null, curPoint1);

			//заносим текущее пересечение
			if (curIntersection) {
				addToArr(pointIndex, seriaIndex, true, curIntersection);
				addToArr(pointIndex, seriaIndex, null, curIntersection);

				//сразу заносим последнюю точку грани
				addToArr(pointIndex, seriaIndex, true, curPoint2);
				addToArr(pointIndex, seriaIndex, null, curPoint4);
			} else {
				//сразу заносим последнюю точку грани
				addToArr(pointIndex, seriaIndex, true, curPoint4);
				addToArr(pointIndex, seriaIndex, null, curPoint2);
			}
		}

		//пересечения с прямыми предыдущих серий
		var line1, line2, intersection;
		if (this.prevDown && this.prevDown[pointIndex]) {
			for (var i = 0; i < this.prevDown[pointIndex].length; i++) {
				var prevDown = this.prevDown[pointIndex][i];

				for (var j = 0; j < this.prevDown[pointIndex][i].length; j++) {
					for (var k = 0; k < curDown.length; k++) {
						line1 = this.prevDown[pointIndex][i][j];
						line2 = curDown[k];
						intersection = this.cChartDrawer.isIntersectionLineAndLine(line1.eq, line2.eq);

						if (intersection) {
							//предыдущая - i
							if (intersection.x > line1.start.x && intersection.x < line1.end.x) {
								addToArr(pointIndex, i, true, intersection);
							}

							//текущая серия
							if (intersection.x > line2.start.x && intersection.x < line2.end.x) {
								addToArr(pointIndex, seriaIndex, true, intersection);
							}
						}

					}
				}
			}
		}

		//пересечения с прямыми предыдущих серий
		if (this.prevUp && this.prevUp[pointIndex]) {
			for (var i = 0; i < this.prevUp[pointIndex].length; i++) {
				var prevUp = this.prevUp[pointIndex][i];

				for (var j = 0; j < this.prevUp[pointIndex][i].length; j++) {
					for (var k = 0; k < curUp.length; k++) {
						line1 = this.prevUp[pointIndex][i][j];
						line2 = curUp[k];
						intersection = this.cChartDrawer.isIntersectionLineAndLine(line1.eq, line2.eq);

						if (intersection) {
							//предыдущая - i
							if (intersection.x > line1.start.x && intersection.x < line1.end.x) {
								addToArr(pointIndex, i, null, intersection);
							}

							//текущая серия
							if (intersection.x > line2.start.x && intersection.x < line2.end.x) {
								addToArr(pointIndex, seriaIndex, null, intersection);
							}
						}

					}
				}
			}
		}

		if (!this.prevDown) {
			this.prevDown = [];
		}
		if (!this.prevDown[pointIndex]) {
			this.prevDown[pointIndex] = [];
		}

		this.prevDown[pointIndex][seriaIndex] = curDown;

		if (!this.prevUp) {
			this.prevUp = [];
		}
		if (!this.prevUp[pointIndex]) {
			this.prevUp[pointIndex] = [];
		}

		this.prevUp[pointIndex][seriaIndex] = curUp;
	},

	_calculateSimpleRect: function (arrPoints, arrPointsProject, point, seria) {
		var pxToMm = this.chartProp.pxToMM, t = this, paths = [];
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		var point1 = arrPointsProject[0];
		var point2 = arrPointsProject[1];
		var point3 = arrPointsProject[2];
		var point4 = arrPointsProject[3];
		var point5 = arrPointsProject[4];
		var point6 = arrPointsProject[5];
		var point7 = arrPointsProject[6];
		var point8 = arrPointsProject[7];

		var point11 = arrPoints[0];
		var point22 = arrPoints[1];
		var point33 = arrPoints[2];
		var point44 = arrPoints[3];
		var point55 = arrPoints[4];
		var point66 = arrPoints[5];
		var point77 = arrPoints[6];
		var point88 = arrPoints[7];

		var insidePoint = {};
		insidePoint.x =
			(point11.x + point22.x + point33.x + point44.x + point55.x + point66.x + point77.x + point88.x) / 8;
		insidePoint.y =
			(point11.y + point22.y + point33.y + point44.y + point55.y + point66.y + point77.y + point88.y) / 8;
		insidePoint.z =
			(point11.z + point22.z + point33.z + point44.z + point55.z + point66.z + point77.z + point88.z) / 8;

		var oThis = this;
		var calculateSimpleFace = function (p1, p2, p3, p4, p11, p22, p33, p44, faceIndex) {
			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			path.pathH = pathH;
			path.pathW = pathW;

			path.moveTo(p1.x / pxToMm * pathW, p1.y / pxToMm * pathH);
			path.lnTo(p2.x / pxToMm * pathW, p2.y / pxToMm * pathH);
			path.lnTo(p3.x / pxToMm * pathW, p3.y / pxToMm * pathH);
			path.lnTo(p4.x / pxToMm * pathW, p4.y / pxToMm * pathH);
			path.lnTo(p1.x / pxToMm * pathW, p1.y / pxToMm * pathH);


			var arrPoints = [p1, p2, p3, p4];
			var arrPoints2 = [p11, p22, p33, p44];
			var plainEquation = t.cChartDrawer.getPlainEquation(p11, p22, p33);
			var plainArea = t.cChartDrawer.getAreaQuadrilateral(p1, p2, p3, p4);

			//TODO POINT = 0!!!!
			if (faceIndex === 0) {
				t.sortZIndexPathsFront.push({
					seria: seria,
					point: 0,
					verge: faceIndex,
					paths: pathId,
					points: arrPoints2,
					points2: arrPoints,
					plainEquation: plainEquation,
					plainArea: plainArea,
					z: insidePoint.z
				});
			} else if (faceIndex === 5) {
				t.sortZIndexPathsBack.push({
					seria: seria,
					point: 0,
					verge: faceIndex,
					paths: pathId,
					points: arrPoints2,
					points2: arrPoints,
					plainEquation: plainEquation,
					plainArea: plainArea,
					z: insidePoint.z
				});
			} else if (faceIndex === 2) {
				t.sortZIndexPathsLeft.push({
					seria: seria,
					point: 0,
					verge: faceIndex,
					paths: pathId,
					points: arrPoints2,
					points2: arrPoints,
					plainEquation: plainEquation,
					plainArea: plainArea,
					z: insidePoint.z
				});
			} else if (faceIndex === 3) {
				t.sortZIndexPathsRight.push({
					seria: seria,
					point: 0,
					verge: faceIndex,
					paths: pathId,
					points: arrPoints2,
					points2: arrPoints,
					plainEquation: plainEquation,
					plainArea: plainArea,
					z: insidePoint.z
				});
			}

			return pathId;
		};

		let countSeries = oThis.cChartDrawer.calculateCountSeries(oThis.chart);
		let ptCount = countSeries.points;

		//left
		paths[2] = null;
		if (this.darkFaces["left"] && point === 0) {
			paths[2] = calculateSimpleFace(point1, point2, point6, point5, point11, point22, point66, point55, 2);
		}

		//right
		paths[3] = null;
		if (this.darkFaces["right"] && point === ptCount - 2) {
			paths[3] = calculateSimpleFace(point4, point8, point7, point3, point44, point88, point77, point33, 3);
		}

		//front
		if (this.darkFaces["front"]) {
			paths[0] = calculateSimpleFace(point1, point5, point8, point4, point11, point55, point88, point44, 0);
		}

		if (this.darkFaces["unfront"]) {
			paths[5] = calculateSimpleFace(point2, point6, point7, point3, point22, point66, point77, point33, 5);
		}
	},

	_getPrevSeriesPoints: function (points, i) {
		var prevPoints = null;

		for (var p = i - 1; p >= 0; p--) {
			if (points[p]) {
				prevPoints = points[p];
				break;
			}
		}

		return prevPoints;
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var pxToMm = this.chartProp.pxToMM;

		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);
		var path;

		if(!point) {
			return;
		}

		path = this.points[ser][val];

		if (!path) {
			return;
		}

		var x = path.x;
		var y = path.y;

		if (this.cChartDrawer.nDimensionCount === 3) {
			var gapDepth = this.gapDepth;
			if (this.subType === "normal") {
				gapDepth = this.gapDepth[ser];
			}

			var turnPoint = this.cChartDrawer._convertAndTurnPoint(x * pxToMm, y * pxToMm, gapDepth);
			x = turnPoint.x / pxToMm;
			y = turnPoint.y / pxToMm;
		}


		var constMargin = 5 / pxToMm;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x - width / 2;
		var centerY = y - height / 2;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_drawLines: function (/*isSkip*/) {
		//ширина линии
		var brush;
		var pen;
		var seria, dataSeries, numCache;

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(
			this.chartProp.chartGutter._left / this.chartProp.pxToMM,
			(this.chartProp.chartGutter._top - 1) / this.chartProp.pxToMM,
			this.chartProp.trueWidth / this.chartProp.pxToMM, this.chartProp.trueHeight / this.chartProp.pxToMM);

		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			numCache = this.cChartDrawer.getNumCache(seria.val);
			dataSeries = numCache && numCache.pts ? numCache.pts : null;

			if (!dataSeries) {
				continue;
			}

			if (dataSeries[0] && dataSeries[0].pen) {
				pen = dataSeries[0].pen;
			}
			if (dataSeries[0] && dataSeries[0].brush) {
				brush = dataSeries[0].brush;
			}

			if (this.cChartDrawer.nDimensionCount === 3) {
				for (var j = 0; j < this.paths.series[i].length; j++) {
					this.cChartDrawer.drawPath(this.paths.series[i][j], pen, brush);
				}
			} else {
				this.cChartDrawer.drawPath(this.paths.series[i], pen, brush);
			}
		}

		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	_getYVal: function (n, i) {
		//TODO сделать общую функцию для line/area!
		var tempVal;
		var val = 0;
		var idxPoint;
		var k;
		if (this.subType === "stacked") {
			for (k = 0; k <= i; k++) {
				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[k], n);
				tempVal = idxPoint ? parseFloat(idxPoint.val) : 0;
				if (tempVal) {
					val += tempVal;
				}
			}
		} else if (this.subType === "stackedPer") {
			var sumVal = 0;
			for (k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[k], n);
				tempVal = idxPoint ? parseFloat(idxPoint.val) : 0;
				if (tempVal) {
					if (k <= i) {
						val += tempVal;
					}
					sumVal += Math.abs(tempVal);
				}
			}
			if(sumVal === 0) {
				val = 0;
			} else {
				val = val / sumVal;
			}
		} else {
			idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[i], n);
			//TODO blank SPAN option
			var dispBlanksAs =  this.cChartSpace.chart.dispBlanksAs;
			if(idxPoint) {
				val = parseFloat(idxPoint.val);
			} else if(dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
				val = 0;
			} else {
				val = null;
			}
		}

		return val;
	},

	_drawLines3D: function (/*isSkip*/) {
		//ширина линии
		var brush;
		var pen;
		var seria, dataSeries, numCache;

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(
			this.chartProp.chartGutter._left / this.chartProp.pxToMM,
			(this.chartProp.chartGutter._top - 1) / this.chartProp.pxToMM,
			this.chartProp.trueWidth / this.chartProp.pxToMM, this.chartProp.trueHeight / this.chartProp.pxToMM);

		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			numCache = this.cChartDrawer.getNumCache(seria.val);
			dataSeries = numCache && numCache.pts ? numCache.pts : null;

			if (!dataSeries) {
				continue;
			}

			if (dataSeries[0] && dataSeries[0].pen) {
				pen = dataSeries[0].pen;
			}
			if (dataSeries[0] && dataSeries[0].brush) {
				brush = dataSeries[0].brush;
			}

			for (var j = 0; j < this.paths.series[i].length; j++) {
				this.cChartDrawer.drawPath(this.paths.series[i][j], pen, brush);
			}
		}

		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	_drawBar3D: function (path, pen, brush, k) {
		//затемнение боковых сторон
		//в excel всегда темные боковые стороны, лицевая и задняя стороны светлые
		//pen = this.cChartSpace.chart.plotArea.valAx.compiledMajorGridLines;
		//pen.setFill(brush);
		pen = AscFormat.CreatePenFromParams(brush, undefined, undefined, undefined, undefined, 0.2);
		if (k !== 5 && k !== 0) {
			var props = this.cChartSpace.getParentObjects();

			if (brush.isNoFill()) {
				return;
			}

			var duplicateBrush = brush.createDuplicate();
			var cColorMod = new AscFormat.CColorMod;
			if (k === 1 || k === 4) {
				cColorMod.val = 45000;
			} else {
				cColorMod.val = 35000;
			}
			cColorMod.name = "shade";
			duplicateBrush.addColorMod(cColorMod);
			duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master,
				new AscFormat.CUniColor().RGBA, this.cChartSpace.clrMapOvr);

			pen.setFill(duplicateBrush);
			if (path && path.length) {
				for (var i = 0; i < path.length; i++) {
					this.cChartDrawer.drawPath(path[i], pen, duplicateBrush);
				}
			} else {
				this.cChartDrawer.drawPath(path, pen, duplicateBrush);
			}
		} else {
			if (path && path.length) {
				for (var i = 0; i < path.length; i++) {
					this.cChartDrawer.drawPath(path[i], pen, brush);
				}
			} else {
				this.cChartDrawer.drawPath(path, pen, brush);
			}
		}
	},

	_standardDraw: function (isReverse) {
		var j = isReverse ? 0 : this.paths.series.length - 1;
		var i = isReverse ? this.paths.series.length : 0;

		var seria, brush, pen, numCache;
		while (((j < i) && isReverse) || ((j >= i) && !isReverse)) {
			seria = this.chart.series[j];
			brush = seria.brush;
			pen = seria.pen;

			numCache = this.cChartDrawer.getNumCache(seria.val);
			if (numCache && numCache.pts[0].pen) {
				pen = numCache.pts[0].pen;
			}
			if (numCache && numCache.pts[0].brush) {
				brush = numCache.pts[0].brush;
			}

			this._drawBar3D(this.paths.series[j][1], pen, brush, 1);
			this._drawBar3D(this.paths.series[j][4], pen, brush, 4);
			this._drawBar3D(this.paths.series[j][2], pen, brush, 2);
			this._drawBar3D(this.paths.series[j][3], pen, brush, 3);
			this._drawBar3D(this.paths.series[j][5], pen, brush, 5);
			this._drawBar3D(this.paths.series[j][0], pen, brush, 0);

			j = isReverse ? j + 1 : j - 1;
		}

	},

	_drawBars3D: function () {
		var t = this;
		var isStacked = this.subType === "stackedPer" || this.subType === "stacked";

		if (!isStacked) {
			var angle = Math.abs(this.cChartDrawer.processor3D.angleOy);
			if (!this.cChartDrawer.processor3D.view3D.getRAngAx() && angle > Math.PI / 2 &&
				angle < 3 * Math.PI / 2) {
				if (this.serAx && this.serAx.isReversed()) {
					this._standardDraw(false);
				} else {
					this._standardDraw(true);
				}
			} else {
				if (this.serAx && this.serAx.isReversed()) {
					this._standardDraw(true);
				} else {
					this._standardDraw(false);
				}
			}


		} else {
			var drawVerges = function (i, j, paths, onlyLessNull, k) {
				var brush, pen, options;
				options = t._getOptionsForDrawing(i, j, onlyLessNull);
				if (paths !== null && options !== null) {
					pen = options.pen;
					brush = options.brush;

					t._drawBar3D(paths, pen, brush, k);
				}
			};

			//if(this.cChartDrawer.processor3D.view3D.rAngAx)
			//{
			for (var i = 0; i < this.sortZIndexPaths.length; i++) {
				drawVerges(this.sortZIndexPaths[i].seria, this.sortZIndexPaths[i].point, this.sortZIndexPaths[i].paths, null, this.sortZIndexPaths[i].verge);
			}
			//}
		}
	},

	_getOptionsForDrawing: function (ser, point, onlyLessNull) {
		var seria = this.chart.series[ser];
		var numCache = this.cChartDrawer.getNumCache(seria.val);
		if(!numCache) {
			return null;
		}

		var pt = numCache.getPtByIndex(point);
		if (!seria || !this.paths.series[ser] || !this.paths.series[ser][point] || !pt) {
			return null;
		}

		var brush = seria.brush;
		var pen = seria.pen;

		if ((pt.val > 0 && onlyLessNull === true) || (pt.val < 0 && onlyLessNull === false)) {
			return null;
		}

		if (pt.pen) {
			pen = pt.pen;
		}
		if (pt.brush) {
			brush = pt.brush;
		}

		return {pen: pen, brush: brush}
	},

	_isVisibleVerge3D: function (k, n, m, isSideFace) {
		//roberts method - calculate normal
		var aX = n.x * m.y - m.x * n.y;
		var bY = -(k.x * m.y - m.x * k.y);
		var cZ = k.x * n.y - n.x * k.y;
		var visible = (aX + bY + cZ) < 0;

		var valOrientation = !this.valAx.isReversed();
		var catOrientation = !this.catAx.isReversed();

		if ((!valOrientation && !isSideFace) || (!valOrientation && catOrientation)) {
			return !visible;
		} else if (!catOrientation && isSideFace && valOrientation) {
			return !visible;
		}
		return visible;
	},

	_getIntersectionLines: function (line1Point1, line1Point2, line2Point1, line2Point2) {
		var chartHeight = this.chartProp.trueHeight;
		var top = this.chartProp.chartGutter._top;

		var x1 = line1Point1.x;
		var x2 = line1Point2.x;
		var y1 = line1Point1.y;
		var y2 = line1Point2.y;

		var x3 = line2Point1.x;
		var x4 = line2Point2.x;
		var y3 = line2Point1.y;
		var y4 = line2Point2.y;

		var x = ((x1 * y2 - x2 * y1) * (x4 - x3) - (x3 * y4 - x4 * y3) * (x2 - x1)) /
			((y1 - y2) * (x4 - x3) - (y3 - y4) * (x2 - x1));
		var y = ((y3 - y4) * x - (x3 * y4 - x4 * y3)) / (x4 - x3);

		x = -x;

		var res = null;

		if (y < top + chartHeight && y > top && x > line1Point1.x && x < line1Point2.x) {
			res = {x: x, y: y};
		}

		return res;
	}
};


/** @constructor */
function drawHBarChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;
	this.ptCount = null;
	this.seriesCount = null;
	this.subType = null;

	this.paths = {};
	this.sortZIndexPaths = [];

	this.summBarVal = [];
}

drawHBarChart.prototype = {
	constructor: drawHBarChart,

	recalculate: function () {
		this.paths = {};
		this.summBarVal = [];

		this.sortZIndexPaths = [];

		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.seriesCount = countSeries.series;
		this.ptCount = countSeries.points;
		this.subType = this.cChartDrawer.getChartGrouping(this.chart);
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);
		if(!this.catAx) {
			this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_DateAx);
		}
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);

		this._recalculateBars();
	},

	draw: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			this._DrawBars3D();
		} else {
			this._drawBars();
		}
	},

	_recalculateBars: function (/*isSkip*/) {
		//соответствует подписям оси категорий(OX)
		var xPoints = this.valAx.xPoints;
		var yPoints = this.catAx.yPoints;

		var heightGraph = this.chartProp.heightCanvas - this.chartProp.chartGutter._top - this.chartProp.chartGutter._bottom;

		var defaultOverlap = (this.subType === "stacked" || this.subType === "stackedPer") ? 100 : 0;
		var overlap = AscFormat.isRealNumber(this.chart.overlap) ? this.chart.overlap : defaultOverlap;
		var ptCount = this.cChartDrawer.getPtCount(this.chart.series);
		var height = heightGraph / ptCount;
		var crossBetween = this.cChartSpace.getValAxisCrossType();
		if (crossBetween) {
			height = heightGraph / (ptCount - 1);
		}

		var gapWidth = AscFormat.isRealNumber(this.chart.gapWidth) ? this.chart.gapWidth : 150;

		var individualBarHeight = height /
			(this.seriesCount - (this.seriesCount - 1) * (overlap / 100) + gapWidth / 100);
		var widthOverLap = individualBarHeight * (overlap / 100);
		var hmargin = (gapWidth / 100 * individualBarHeight) / 2;

		var width, startX, startY, val, paths, seriesHeight = [], seria, startXColumnPosition, startYPosition, newStartX, newStartY, idx, seriesCounter = 0, numCache;

		//for 3d
		var perspectiveDepth, gapDepth, DiffGapDepth;
		if (this.cChartDrawer.nDimensionCount === 3) {
			perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
			//сдвиг по OZ в глубину
			gapDepth = this.chart.gapDepth != null ? this.chart.gapDepth : globalGapDepth;
			perspectiveDepth = perspectiveDepth / (gapDepth / 100 + 1);
			DiffGapDepth = perspectiveDepth * (gapDepth / 2) / 100;
		}

		var cubeCount = 0;
		for (var i = 0; i < this.chart.series.length; i++) {
			numCache = this.cChartDrawer.getNumCache(this.chart.series[i].val);

			if (!numCache || this.chart.series[i].isHidden) {
				continue;
			}

			seria = numCache.pts;
			seriesHeight[i] = [];


			var isValMoreZero = false;
			var isValLessZero = 0;
			var shapeType = null !== this.chart.series[i].shape ? this.chart.series[i].shape : this.chart.shape;
			//shapeType = 0;
			for (var j = 0; j < seria.length; j++) {
				//стартовая позиция колонки Y(+ высота с учётом поправок на накопительные диаграммы)
				val = parseFloat(seria[j].val);

				if (val > 0) {
					isValMoreZero = true;
				} else if (val < 0) {
					isValLessZero++;
				}

				/*if (this.valAx && this.valAx.scaling.logBase) {
					val = this.cChartDrawer.getLogarithmicValue(val, this.valAx.scaling.logBase, xPoints);
				}*/
				idx = seria[j].idx != null ? seria[j].idx : j;


				startXColumnPosition = this._getStartYColumnPosition(seriesHeight, idx, i, val, xPoints, shapeType);
				startX = startXColumnPosition.startY / this.chartProp.pxToMM;
				width = startXColumnPosition.width / this.chartProp.pxToMM;
				var maxH = startXColumnPosition.maxH;
				var minH = startXColumnPosition.minH;

				seriesHeight[i][idx] = startXColumnPosition.width;


				//стартовая позиция колонки Y
				if (!this.catAx.isReversed()) {
					if (yPoints[1] && yPoints[1].pos && yPoints[idx]) {
						startYPosition = yPoints[idx].pos + Math.abs((yPoints[1].pos - yPoints[0].pos) / 2);
					} else if(yPoints[idx]){
						startYPosition = yPoints[idx].pos + Math.abs(yPoints[0].pos - this.valAx.posY);
					} else {
						startYPosition = yPoints[0].pos + Math.abs(yPoints[0].pos - this.valAx.posY);
					}
				} else {
					if (yPoints[1] && yPoints[1].pos && yPoints[idx]) {
						startYPosition = yPoints[idx].pos - Math.abs((yPoints[1].pos - yPoints[0].pos) / 2);
					} else if(yPoints[idx]){
						startYPosition = yPoints[idx].pos - Math.abs(yPoints[0].pos - this.valAx.posY);
					} else {
						startYPosition = yPoints[0].pos - Math.abs(yPoints[0].pos - this.valAx.posY);
					}
				}


				if (!this.catAx.isReversed()) {
					if (seriesCounter === 0) {
						startY = startYPosition * this.chartProp.pxToMM - hmargin - seriesCounter * (individualBarHeight);
					} else {
						startY = startYPosition * this.chartProp.pxToMM - hmargin - (seriesCounter * individualBarHeight - seriesCounter * widthOverLap);
					}
				} else {
					if (i === 0) {
						startY = startYPosition * this.chartProp.pxToMM + hmargin + seriesCounter * (individualBarHeight);
					} else {
						startY = startYPosition * this.chartProp.pxToMM + hmargin + (seriesCounter * individualBarHeight - seriesCounter * widthOverLap);
					}
				}

				newStartY = startY;
				if (this.catAx.isReversed()) {
					newStartY = startY + individualBarHeight;
				}

				newStartX = startX;
				if (this.valAx.isReversed() &&
					(this.subType === "stackedPer" || this.subType === "stacked")) {
					newStartX = startX - width;
				}

				if (this.cChartDrawer.nDimensionCount === 3) {
					paths = this.calculateParallalepiped(newStartX, newStartY, val, width, DiffGapDepth, perspectiveDepth, individualBarHeight, 
						seriesHeight, i, idx, cubeCount, shapeType, maxH, minH);
					cubeCount++;
				} else {
					paths = this.cChartDrawer._calculateRect(newStartX, newStartY / this.chartProp.pxToMM, width, individualBarHeight / this.chartProp.pxToMM, true);
				}

				var serIdx = this.chart.series[i].idx;
				if (!this.paths.series) {
					this.paths.series = [];
				}
				if (!this.paths.series[serIdx]) {
					this.paths.series[serIdx] = [];
				}

				if (height !== 0) {
					this.paths.series[serIdx][idx] = paths;
				}

				let errBars = this.chart.series[i].errBars[0];
				if (errBars) {
					var _pointX = (startX) / this.chartProp.pxToMM;
					var pointY = (startY - individualBarHeight / 2) / this.chartProp.pxToMM;
					//var _pointVal = this.subType === "stacked" || this.subType === "stackedPer" ? this._getStackedValue(this.chart.series, i, j, val) : val;
					var _pointVal = this.cChartDrawer.getValWithStacked(i, idx, this.chart);
					this.cChartDrawer.errBars.putPoint(_pointX, pointY, _pointVal, _pointVal,  serIdx, idx);
					this.cChartDrawer.errBars.putPoint(_pointX, pointY, _pointVal, _pointVal,  serIdx, idx);
				}

			}

			//if (seria.length) {
				seriesCounter++;
			//}
		}

		if (this.cChartDrawer.nDimensionCount === 3) {
			if (this.cChartDrawer.processor3D.view3D.getRAngAx()) {
				var angle;
				if (this.subType !== "normal") {
					angle = Math.abs(this.cChartDrawer.processor3D.angleOy);
					this.sortZIndexPaths.sort(function sortArr(a, b) {
						if (b.zIndex === a.zIndex) {
							if (angle < Math.PI) {
								return a.x - b.x;
							} else {
								return b.x - a.x;
							}
						} else {
							return b.zIndex - a.zIndex;
						}
					});
				} else {
					angle = -this.cChartDrawer.processor3D.angleOx;
					this.sortZIndexPaths.sort(function sortArr(a, b) {
						if (b.zIndex === a.zIndex) {
							if (angle > 0) {
								return b.y - a.y;
							} else {
								return a.y - b.y;
							}
						} else {
							return b.zIndex - a.zIndex;
						}
					});
				}
			} else {
				var cSortFaces = new window['AscFormat'].CSortFaces(this.cChartDrawer);
				//this.sortZIndexPaths = cSortFaces.sortFaces(this.sortZIndexPaths);
				this.sortParallelepipeds = cSortFaces.sortParallelepipeds(this.temp, this.sortZIndexPaths);
			}
		}
	},

	_calculateLine: function (x, y, x1, y1) {
		var pxToMm = this.chartProp.pxToMM;

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		path.moveTo(x / pxToMm * pathW, y / pxToMm * pathH);
		path.lnTo(x1 / pxToMm * pathW, y1 / pxToMm * pathH);

		return pathId;
	},

	_getOptionsForDrawing: function (ser, point, onlyLessNull) {
		var seria = this.chart.series[ser];
		var numCache = this.cChartDrawer.getNumCache(seria.val);
		if(!numCache) {
			return null;
		}

		var pt = numCache.getPtByIndex(point);
		if (!seria || !this.paths.series[ser] || !this.paths.series[ser][point] || !pt) {
			return null;
		}

		var brush = seria.brush;
		var pen = seria.pen;

		if ((pt.val > 0 && onlyLessNull === true) || (pt.val < 0 && onlyLessNull === false)) {
			return null;
		}

		if (pt.pen) {
			pen = pt.pen;
		}
		if (pt.brush) {
			brush = pt.brush;
		}

		return {pen: pen, brush: brush}
	},

	_getStartYColumnPosition: function (seriesHeight, j, i, val, xPoints, type) {
		var startY, width, curVal, prevVal, endBlockPosition, startBlockPosition;
		var catAx = this.catAx;
		var h, maxH, minH;

		var axisMin = xPoints[0].val < xPoints[xPoints.length - 1].val ? xPoints[0].val : xPoints[xPoints.length - 1].val;
		var axisMax = xPoints[0].val < xPoints[xPoints.length - 1].val ? xPoints[xPoints.length - 1].val : xPoints[0].val;

		//в ms отрисовка сделана следующим образом: если диаграмма типа normal, то стартовую точку отрисовки столбцов берем позицию X оси категорий(posX)
		//если диаграмма типа stacked то рисуем от позиции X ноля оси категорий - getPositionZero(позиция ноля и оси могут отличиться в зависимости от настроек)
		var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) : catAx.posX * this.chartProp.pxToMM;

		var testMaxHeight = this.cChartDrawer.getYPosition(axisMax, this.valAx) * this.chartProp.pxToMM - nullPositionOX;
		var testMinHeight = this.cChartDrawer.getYPosition(axisMin, this.valAx) * this.chartProp.pxToMM - nullPositionOX;

		if (this.subType === "stacked" || this.subType === "stackedPer") {
			curVal = this._getStackedValue(this.chart.series, i, j, val);
			prevVal = this._getStackedValue(this.chart.series, i - 1, j, val);

			if (this.subType === "stacked") {
				//если максимальное значение задано вручную, и присутвуют точки, которые больше этого значения
				if (curVal > axisMax) {
					curVal = axisMax;
				}
				if (curVal < axisMin) {
					curVal = axisMin;
				}

				endBlockPosition = this.cChartDrawer.getYPosition(curVal, this.valAx) * this.chartProp.pxToMM;
				startBlockPosition = prevVal ? this.cChartDrawer.getYPosition(prevVal, this.valAx) * this.chartProp.pxToMM : nullPositionOX;
			} else {
				this._calculateSummStacked(j);

				var test = this.summBarVal[j];

				//если максимальное значение задано вручную, и присутвуют точки, которые больше этого значения
				if (curVal / test > axisMax) {
					curVal = axisMax * test;
				}
				if (curVal / test < axisMin) {
					curVal = axisMin * test;
				}

				if (prevVal / test > axisMax) {
					prevVal = axisMax * test;
				}
				if (prevVal / test < axisMin) {
					prevVal = axisMin * test;
				}

				endBlockPosition = this.cChartDrawer.getYPosition((curVal / test), this.valAx) * this.chartProp.pxToMM;
				startBlockPosition = test ? this.cChartDrawer.getYPosition((prevVal / test), this.valAx) * this.chartProp.pxToMM : nullPositionOX;
			}

			startY = startBlockPosition;
			width = endBlockPosition - startBlockPosition;

			if (this.valAx.isReversed()) {
				width = -width;
			}
		} else {
			width = this.cChartDrawer.getYPosition(val, this.valAx) * this.chartProp.pxToMM - nullPositionOX;
			startY = nullPositionOX;
		}

		if (type === AscFormat.BAR_SHAPE_PYRAMID || type === AscFormat.BAR_SHAPE_PYRAMIDTOMAX ||
			type === AscFormat.BAR_SHAPE_CONE || type === AscFormat.BAR_SHAPE_CONETOMAX) {
			var maxVal = this._getStackedValue(this.chart.series, this.chart.series.length - 1, j, val);
			h = this.cChartDrawer.getStartStackedPyramidPosition(val, this.chart, this.valAx, j, this.chartProp, 
				this.summBarVal, maxVal, nullPositionOX, this.subType, this.ptCount, axisMax, axisMin, testMaxHeight, testMinHeight, type, true);
			maxH = h.maxH;
			minH = h.minH;
		}

		return {startY: startY, width: width, maxH: maxH, minH: minH};
	},

	_calculateSummStacked: function (j) {
		if (!this.summBarVal[j]) {
			var curVal;
			var temp = 0, idxPoint;
			for (var k = 0; k < this.chart.series.length; k++) {
				idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[k], j);
				curVal = idxPoint ? parseFloat(idxPoint.val) : 0;

				if (curVal) {
					temp += Math.abs(curVal);
				}
			}
			this.summBarVal[j] = temp;
		}
	},

	_getStackedValue: function (series, i, j, val) {
		var result = 0, curVal, idxPoint;
		for (var k = 0; k <= i; k++) {
			idxPoint = this.cChartDrawer.getIdxPoint(series[k], j);
			curVal = idxPoint ? idxPoint.val : 0;

			if (idxPoint && val > 0 && curVal > 0) {
				result += parseFloat(curVal);
			} else if (idxPoint && val < 0 &&
				curVal < 0) {
				result += parseFloat(curVal);
			}
		}

		return result;
	},

	_drawBars: function () {
		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect((this.chartProp.chartGutter._left - 1) / this.chartProp.pxToMM, (this.chartProp.chartGutter._top - 1) / this.chartProp.pxToMM, this.chartProp.trueWidth / this.chartProp.pxToMM, this.chartProp.trueHeight / this.chartProp.pxToMM);
		this.cChartDrawer.drawPaths(this.paths, this.chart.series, null, null, true);
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	_calculateDLbl: function (chartSpace, ser, val, bLayout, serIdx) {
		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);
		var path = this.paths && this.paths.series && this.paths.series[serIdx] && this.paths.series[serIdx][val];

		if(!point || !path) {
			return;
		}

		if (this.cChartDrawer.nDimensionCount === 3 && path.frontPaths) {
			var frontPaths = path.frontPaths;
			if (this.cChartDrawer.nDimensionCount === 3) {
				if (AscFormat.isRealNumber(frontPaths[0])) {
					path = frontPaths[0];
				} else if (AscFormat.isRealNumber(frontPaths[5])) {
					path = frontPaths[5];
				} else if (AscFormat.isRealNumber(frontPaths[2])) {
					path = frontPaths[2];
				} else if (AscFormat.isRealNumber(frontPaths[3])) {
					path = frontPaths[3];
				}
			}
		}

		if (!AscFormat.isRealNumber(path)) {
			return;
		}
		var oPath = this.cChartSpace.GetPath(path);
		var oCommand0 = oPath.getCommandByIndex(0);
		var oCommand1 = oPath.getCommandByIndex(1);
		var oCommand2 = oPath.getCommandByIndex(2);


		var x = oCommand0.X;
		var y = oCommand0.Y;

		var h = oCommand0.Y - oCommand1.Y;
		var w = oCommand2.X - oCommand1.X;

		var pxToMm = this.chartProp.pxToMM;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX, centerY;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				centerX = x + w / 2 - width / 2;
				centerY = y - h / 2 - height / 2;
				break;
			}
			case c_oAscChartDataLabelsPos.inBase: {
				centerX = x;
				centerY = y - h / 2 - height / 2;
				if (point.val < 0) {
					centerX = x - width;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.inEnd: {
				centerX = x + w - width;
				centerY = y - h / 2 - height / 2;
				if (point.val < 0) {
					centerX = x + w;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.outEnd: {
				centerX = x + w;
				centerY = y - h / 2 - height / 2;
				if (point.val < 0) {
					centerX = x + w - width;
				}
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	calculateParallalepiped: function (newStartX, newStartY, val, width, DiffGapDepth, perspectiveDepth,
									   individualBarHeight, seriesHeight, i, idx, cubeCount, type, maxH, minH) {
		var paths, paths2;
		var point1, point2, point3, point4, point5, point6, point7, point8;
		var x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, x5, y5, z5, x6, y6, z6, x7, y7, z7, x8, y8, z8;
		var xPoints = this.valAx.xPoints;

		var scaleAxis = this.valAx.scale;
		var axisMax = scaleAxis[0] < scaleAxis[scaleAxis.length - 1] ? scaleAxis[scaleAxis.length - 1] : scaleAxis[0];

		width = width * this.chartProp.pxToMM;
		newStartX = newStartX * this.chartProp.pxToMM;
		newStartY = newStartY - individualBarHeight;

		//рассчитываем 8 точек для каждого столбца
		x1 = newStartX, y1 = newStartY, z1 = DiffGapDepth;
		x2 = newStartX, y2 = newStartY, z2 = perspectiveDepth + DiffGapDepth;
		x3 = newStartX + width, y3 = newStartY, z3 = perspectiveDepth + DiffGapDepth;
		x4 = newStartX + width, y4 = newStartY, z4 = DiffGapDepth;
		x5 = newStartX, y5 = newStartY + individualBarHeight, z5 = DiffGapDepth;
		x6 = newStartX, y6 = newStartY + individualBarHeight, z6 = perspectiveDepth + DiffGapDepth;
		x7 = newStartX + width, y7 = newStartY + individualBarHeight, z7 = perspectiveDepth + DiffGapDepth;
		x8 = newStartX + width, y8 = newStartY + individualBarHeight, z8 = DiffGapDepth;

		//поворот относительно осей
		point1 = this.cChartDrawer._convertAndTurnPoint(x1, y1, z1);
		point2 = this.cChartDrawer._convertAndTurnPoint(x2, y2, z2);
		point3 = this.cChartDrawer._convertAndTurnPoint(x3, y3, z3);
		point4 = this.cChartDrawer._convertAndTurnPoint(x4, y4, z4);
		point5 = this.cChartDrawer._convertAndTurnPoint(x5, y5, z5);
		point6 = this.cChartDrawer._convertAndTurnPoint(x6, y6, z6);
		point7 = this.cChartDrawer._convertAndTurnPoint(x7, y7, z7);
		point8 = this.cChartDrawer._convertAndTurnPoint(x8, y8, z8);

		var points = [point1, point2, point3, point4, point5, point6, point7, point8];

		switch (type) {
			case AscFormat.BAR_SHAPE_PYRAMID:
			case AscFormat.BAR_SHAPE_PYRAMIDTOMAX: {
				var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
					this.catAx.posX * this.chartProp.pxToMM;
				paths = this.cChartDrawer.calculatePyramid(true, this.subType, newStartX, newStartY, width, DiffGapDepth, 
					individualBarHeight, perspectiveDepth, val, nullPositionOX, maxH, minH);
				paths2 = this.cChartDrawer.calculateRect3D(points, val, null, true);
				break
			}
			case AscFormat.BAR_SHAPE_CYLINDER: {
				paths = this.cChartDrawer._calculateCylinder(newStartX, newStartY, individualBarHeight, width, val, DiffGapDepth, perspectiveDepth, true, true);
				paths2 = this.cChartDrawer.calculateRect3D(points, val, null, true);
				break;
			}
			case AscFormat.BAR_SHAPE_CONE:
			case AscFormat.BAR_SHAPE_CONETOMAX: {
				var nullPositionOX = this.subType === "stacked" ? this.cChartDrawer.getPositionZero(this.valAx) :
					this.catAx.posX * this.chartProp.pxToMM;
				paths = this.cChartDrawer._calculateCylinder(newStartX, newStartY, individualBarHeight, width, val, DiffGapDepth, perspectiveDepth, true, true,
					this.subType, nullPositionOX, maxH, minH);
				paths2 = this.cChartDrawer.calculateRect3D(points, val, null, true);
				break;
			}
			default: {
				paths2 = this.cChartDrawer.calculateRect3D(points, val, null, true);
				paths = paths2;
				break;
			}
		}

		if (this.cChartDrawer.processor3D.view3D.getRAngAx()) {
			var controlPoint = this.cChartDrawer._convertAndTurnPoint(x5 + width / 2, y5 - individualBarHeight / 2,
				z5);

			for (var k = 0; k < paths.frontPaths.length; k++) {
				this.sortZIndexPaths.push({
					seria: i,
					point: idx,
					verge: k,
					frontPaths: paths.frontPaths[k],
					darkPaths: paths.darkPaths[k],
					x: controlPoint.x,
					y: controlPoint.y,
					zIndex: controlPoint.z,
					shapeType: type,
				});
			}
		} else {
			//рассчитываем 8 точек для каждого столбца одинакового размера для рассчета положения столбцов
			if (this.subType === "normal") {
				var startXColumnPosition = this._getStartYColumnPosition(seriesHeight, idx, i, axisMax, xPoints);
				width = startXColumnPosition.width / this.chartProp.pxToMM;

				x3 = newStartX + width, y3 = newStartY, z3 = perspectiveDepth + DiffGapDepth;
				x4 = newStartX + width, y4 = newStartY, z4 = DiffGapDepth;
				x7 = newStartX + width, y7 = newStartY + individualBarHeight, z7 = perspectiveDepth + DiffGapDepth;
				x8 = newStartX + width, y8 = newStartY + individualBarHeight, z8 = DiffGapDepth;

				point3 = this.cChartDrawer._convertAndTurnPoint(x3, y3, z3);
				point4 = this.cChartDrawer._convertAndTurnPoint(x4, y4, z4);
				point7 = this.cChartDrawer._convertAndTurnPoint(x7, y7, z7);
				point8 = this.cChartDrawer._convertAndTurnPoint(x8, y8, z8);
			}

			//не проецируем на плоскость
			var point11 = this.cChartDrawer._convertAndTurnPoint(x1, y1, z1, null, null, true);
			var point22 = this.cChartDrawer._convertAndTurnPoint(x2, y2, z2, null, null, true);
			var point33 = this.cChartDrawer._convertAndTurnPoint(x3, y3, z3, null, null, true);
			var point44 = this.cChartDrawer._convertAndTurnPoint(x4, y4, z4, null, null, true);
			var point55 = this.cChartDrawer._convertAndTurnPoint(x5, y5, z5, null, null, true);
			var point66 = this.cChartDrawer._convertAndTurnPoint(x6, y6, z6, null, null, true);
			var point77 = this.cChartDrawer._convertAndTurnPoint(x7, y7, z7, null, null, true);
			var point88 = this.cChartDrawer._convertAndTurnPoint(x8, y8, z8, null, null, true);


			var arrPoints = [[point1, point4, point8, point5], [point1, point2, point3, point4],
				[point1, point2, point6, point5], [point4, point8, point7, point3],
				[point5, point6, point7, point8], [point6, point2, point3, point7]];

			var arrPoints2 = [[point11, point44, point88, point55], [point11, point22, point33, point44],
				[point11, point22, point66, point55], [point44, point88, point77, point33],
				[point55, point66, point77, point88], [point66, point22, point33, point77]];


			if (!this.temp) {
				this.temp = [];
			}
			if (!this.temp[cubeCount]) {
				this.temp[cubeCount] = {};
			}
			if (!this.temp[cubeCount].faces) {
				this.temp[cubeCount].faces = [];
				this.temp[cubeCount].arrPoints =
					[point11, point22, point33, point44, point55, point66, point77, point88];
				this.temp[cubeCount].z = point11.z;
				this.temp[cubeCount].isValZero = val === 0 ? true : false;
			}

			for (var k = 0; k < paths.frontPaths.length; k++) {
				if (null === paths.frontPaths[k] && null === paths.darkPaths[k]) {
					continue;
				}

				//this.sortZIndexPaths.push({seria: i, point: idx, verge: k, paths: paths[k], points: arrPoints2[k], points2: arrPoints[k], plainEquation: plainEquation});

				var plainEquation = this.cChartDrawer.getPlainEquation(arrPoints2[k][0], arrPoints2[k][1], arrPoints2[k][2], arrPoints2[k][3]);
				var plainArea = this.cChartDrawer.getAreaQuadrilateral(arrPoints[k][0], arrPoints[k][1], arrPoints[k][2], arrPoints[k][3]);
				this.temp[cubeCount].faces.push({
					seria: i,
					point: idx,
					verge: k,
					frontPaths: paths.frontPaths[k],
					darkPaths: paths.darkPaths[k],
					points: arrPoints2[k],
					points2: arrPoints[k],
					plainEquation: plainEquation,
					plainArea: plainArea,
					shapeType: type,
				});
			}
		}

		return paths2;
	},

	//TODO delete after test
	_DrawBars3D2: function () {
		var t = this;
		var draw = function (onlyLessNull) {
			var brush, pen, options;
			for (var i = 0; i < t.ptCount; i++) {
				if (!t.paths.series) {
					return;
				}

				for (var j = 0; j < t.paths.series.length; ++j) {
					options = t._getOptionsForDrawing(j, i, onlyLessNull);
					if (options !== null) {
						pen = options.pen;
						brush = options.brush;
					} else {
						continue;
					}

					for (var k = 0; k < t.paths.series[j][i].length; k++) {
						t._drawBar3D(t.paths.series[j][i][k], pen, brush, k);
					}
				}
			}
		};

		var drawReverse = function (onlyLessNull) {
			var brush, pen, options;
			for (var i = 0; i < t.ptCount; i++) {
				if (!t.paths.series) {
					return;
				}

				for (var j = t.paths.series.length - 1; j >= 0; --j) {
					options = t._getOptionsForDrawing(j, i, onlyLessNull);
					if (options !== null) {
						pen = options.pen;
						brush = options.brush;
					} else {
						continue;
					}

					for (var k = 0; k < t.paths.series[j][i].length; k++) {
						t._drawBar3D(t.paths.series[j][i][k], pen, brush, k);
					}
				}
			}
		};

		if (this.subType === "stacked" || this.subType === "stackedPer") {
			if (!this.valAx.isReversed()) {
				drawReverse(true);
				draw(false);
			} else {
				drawReverse(false);
				draw(true);
			}
		} else {
			if (!this.catAx.isReversed()) {
				draw();
			} else {
				drawReverse();
			}
		}
	},


	_DrawBars3D3: function () {
		var t = this;

		var verges = {
			front: 0, down: 1, left: 2, right: 3, up: 4, unfront: 5
		};

		var drawVerges = function (i, j, paths, onlyLessNull, start, stop) {
			var brush, pen, options;
			options = t._getOptionsForDrawing(i, j, onlyLessNull);
			if (options !== null) {
				pen = options.pen;
				brush = options.brush;

				for (var k = start; k <= stop; k++) {
					t._drawBar3D(paths[k], pen, brush, k);
				}
			}
		};

		var draw = function (onlyLessNull, start, stop) {
			for (var i = 0; i < t.sortZIndexPaths.length; i++) {
				drawVerges(t.sortZIndexPaths[i].seria, t.sortZIndexPaths[i].point, t.sortZIndexPaths[i].paths,
					onlyLessNull, start, stop);
			}
		};

		if (this.subType === "standard") {
			draw(true, verges.front, verges.unfront);
			draw(false, verges.front, verges.unfront);
		} else {
			draw(true, verges.down, verges.up);
			draw(false, verges.down, verges.up);

			draw(true, verges.unfront, verges.unfront);
			draw(false, verges.unfront, verges.unfront);

			draw(true, verges.front, verges.front);
			draw(false, verges.front, verges.front);
		}
	},

	_DrawBars3D1: function () {
		var t = this;
		var processor3D = this.cChartDrawer.processor3D;

		var drawVerges = function (i, j, paths, onlyLessNull, k) {
			var brush, pen, options;
			options = t._getOptionsForDrawing(i, j, onlyLessNull);
			if (paths !== null && options !== null) {
				pen = options.pen;
				brush = options.brush;

				t._drawBar3D(paths, pen, brush, k);
			}
		};

		for (var i = 0; i < this.sortZIndexPaths.length; i++) {
			if (this.sortZIndexPaths[i].seria !== undefined) {
				drawVerges(this.sortZIndexPaths[i].seria, this.sortZIndexPaths[i].point,
					this.sortZIndexPaths[i].paths, null, this.sortZIndexPaths[i].verge);
			} else {
				//if(options !== null)
				//{
				var pen = t.valAx.compiledMajorGridLines;
				var brush = null;

				this.cChartDrawer.drawPath(this.sortZIndexPaths[i].paths, pen, brush);
				//}
			}

		}
	},

	_DrawBars3D: function () {
		var t = this;

		var drawVerges = function (i, j, paths, onlyLessNull, k, isNotPen, isNotBrush) {
			var brush, pen, options;
			options = t._getOptionsForDrawing(i, j, onlyLessNull);
			if (paths !== null && options !== null) {
				if (!isNotPen) {
					pen = options.pen;
				}
				if (!isNotBrush) {
					brush = options.brush;
				}

				t._drawBar3D(paths, pen, brush, k);
			}
		};

		var index, faces, face, isNotPen;
		if (this.cChartDrawer.processor3D.view3D.getRAngAx()) {
			for (var i = 0; i < this.sortZIndexPaths.length; i++) {
				isNotPen = this.cChartDrawer.checkingPenForDrawing(this.sortZIndexPaths[i].shapeType, this.sortZIndexPaths[i].verge);
				drawVerges(this.sortZIndexPaths[i].seria, this.sortZIndexPaths[i].point,
					this.sortZIndexPaths[i].darkPaths, null, this.sortZIndexPaths[i].verge, isNotPen, true);
			}

			for (var i = 0; i < this.sortZIndexPaths.length; i++) {
				isNotPen = this.cChartDrawer.checkingPenForDrawing(this.sortZIndexPaths[i].shapeType, this.sortZIndexPaths[i].verge);
				drawVerges(this.sortZIndexPaths[i].seria, this.sortZIndexPaths[i].point,
					this.sortZIndexPaths[i].frontPaths, null, this.sortZIndexPaths[i].verge, isNotPen);
			}
		} else {
			for (var i = 0; i < this.sortParallelepipeds.length; i++) {
				index = this.sortParallelepipeds[i].nextIndex;
				faces = this.temp[index].faces;
				for (var j = 0; j < faces.length; j++) {
					face = faces[j];
					isNotPen = this.cChartDrawer.checkingPenForDrawing(face.shapeType, face.verge);
					drawVerges(face.seria, face.point, face.darkPaths, null, face.verge, isNotPen, true);
				}
			}

			for (var i = 0; i < this.sortParallelepipeds.length; i++) {
				index = this.sortParallelepipeds[i].nextIndex;
				faces = this.temp[index].faces;
				for (var j = 0; j < faces.length; j++) {
					face = faces[j];
					isNotPen = this.cChartDrawer.checkingPenForDrawing(face.shapeType, face.verge);
					drawVerges(face.seria, face.point, face.frontPaths, null, face.verge, isNotPen);
				}
			}
		}
	},

	_drawBar3D: function (path, pen, brush, k) {
		//затемнение боковых сторон
		var fill = this._getFill(pen, brush, k);
		var newBrush = fill.brush;
		var newPen = fill.pen;
		this.cChartDrawer.drawPath(path, newPen, newBrush);
	},

	_getFill: function (pen, brush, face) {
		//k: 0 - передняя, 1 - верхняя, 2 - левая, 3 - правая, 4 - нижняя, 5 - задняя
		var shade = "shade";
		var shadeValue1 = 35000;
		var shadeValue2 = 45000;

		var newBrush = brush;
		var newPen = pen;
		var t = this;

		/*if(null === pen || (null !== pen && null === pen.Fill) || (null !== pen && null !== pen.Fill && null === pen.Fill.fill))
		 {
		 pen = AscFormat.CreatePenFromParams(brush, undefined, undefined, undefined, undefined, 0.1);
		 }*/

		//TODO будет время - сделать градиентную заливку в зависимости от угла!!!!
		var color;
		if (brush && brush.fill && AscDFH.historyitem_type_GradFill === brush.fill.getObjectType()) {
			switch (face) {
				case c_oChartBar3dFaces.front:
				case c_oChartBar3dFaces.back: {
					break;
				}
				case c_oChartBar3dFaces.up: {
					color = this._getGradFill(brush, pen, c_oChartBar3dFaces.up);
					newBrush = color.brush;
					newPen = color.pen;

					break;
				}
				case c_oChartBar3dFaces.left: {
					color = this._getGradFill(brush, pen, c_oChartBar3dFaces.left);
					newBrush = color.brush;
					newPen = color.pen;

					break;
				}
				case c_oChartBar3dFaces.right: {
					color = this._getGradFill(brush, pen, c_oChartBar3dFaces.right);
					newBrush = color.brush;
					newPen = color.pen;

					break;
				}
				case c_oChartBar3dFaces.down: {
					color = this._getGradFill(brush, pen, c_oChartBar3dFaces.down);
					newBrush = color.brush;
					newPen = color.pen;

					break;
				}
			}
		} else if (brush && brush.fill) {
			switch (face) {
				case c_oChartBar3dFaces.front:
				case c_oChartBar3dFaces.back: {
					break;
				}
				case c_oChartBar3dFaces.up:
				case c_oChartBar3dFaces.down: {
					newBrush = this._applyColorModeByBrush(brush, shadeValue1);
					if (null === pen) {
						newPen = pen.setFill(newBrush);
					}
					break;
				}
				case c_oChartBar3dFaces.left:
				case c_oChartBar3dFaces.right: {
					newBrush = this._applyColorModeByBrush(brush, shadeValue2);
					if (null === pen) {
						newPen = pen.setFill(newBrush);
					}
					break;
				}
			}
		}

		return {brush: newBrush, pen: newPen};
	},

	_getGradFill: function (brushFill, penFill, faceIndex) {
		var gradientPen = penFill;
		var gradientBrush = brushFill;

		var angleKf = 60000;
		var shade = "shade";
		var shadeValue1 = 35000;
		var t = this;

		if (brushFill.fill.lin && null !== brushFill.fill.lin.angle) {
			var getCSolidColor = function (color, colorMod) {
				var duplicateBrush = brushFill.createDuplicate();
				var tempColor = new AscFormat.CUniFill();
				tempColor.setFill(new AscFormat.CSolidFill());
				tempColor.fill.setColor(color);
				if (colorMod) {
					tempColor = t._applyColorModeByBrush(tempColor, colorMod);
				}

				return tempColor;
			};

			var angle = brushFill.fill.lin.angle / angleKf;
			var colors = brushFill.fill.colors;

			if (angle >= 0 && angle < 45) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.left) {
					gradientBrush = getCSolidColor(colors[0].color);
				} else if (faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = getCSolidColor(colors[colors.length - 1].color);
				}
			} else if (angle >= 45 && angle < 90) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.left) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = getCSolidColor(colors[colors.length - 1].color);
				}
			} else if (angle >= 90 && angle < 135) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.left) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = getCSolidColor(colors[colors.length - 1].color);
				}
			} else if (angle >= 135 && angle < 180) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.left ||
					faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				}
			} else if (angle >= 180 && angle < 225) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.left) {
					gradientBrush = getCSolidColor(colors[colors.length - 1].color, shadeValue1);
				}
			} else if (angle >= 225 && angle < 270) {
				if (faceIndex === c_oChartBar3dFaces.up) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.left || faceIndex === c_oChartBar3dFaces.down ||
					faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				}
			} else if (angle >= 270 && angle < 315) {
				if (faceIndex === c_oChartBar3dFaces.up) {
					gradientBrush = getCSolidColor(colors[colors.length - 1].color);
				} else if (faceIndex === c_oChartBar3dFaces.left || faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				}
			} else if (angle >= 315 && angle <= 360) {
				if (faceIndex === c_oChartBar3dFaces.up || faceIndex === c_oChartBar3dFaces.right) {
					gradientBrush = this._applyColorModeByBrush(brushFill, shadeValue1);
				} else if (faceIndex === c_oChartBar3dFaces.left || faceIndex === c_oChartBar3dFaces.down) {
					gradientBrush = getCSolidColor(colors[0].color, shadeValue1);
				}
			}
		}

		return {pen: gradientPen, brush: gradientBrush};
	},

	_applyColorModeByBrush: function (brushFill, val) {
		var props = this.cChartSpace.getParentObjects();
		var duplicateBrush = brushFill.createDuplicate();
		var cColorMod = new AscFormat.CColorMod("shade", val);
		duplicateBrush.addColorMod(cColorMod);
		duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master, new AscFormat.CUniColor().RGBA, this.cChartSpace.clrMapOvr);

		return duplicateBrush;
	}

};

/** @constructor */
function drawPieChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;

	this.tempAngle = null;
	this.paths = {};
	this.cX = null;
	this.cY = null;
	this.angleFor3D = null;
	this.properties3d = null;
	this.usually3dPropsCalc = [];
	this.empty3DSegment = null;

	this.tempDrawOrder = null;
}

drawPieChart.prototype = {
	constructor: drawPieChart,

	draw: function () {
		if (this.cChartDrawer.nDimensionCount === 3) {
			this._drawPie3D();
		} else {
			this._drawPie();
		}
	},

	recalculate: function () {
		this.tempAngle = null;
		this.paths = {};

		if (this.cChartDrawer.nDimensionCount === 3) {
			if (this.cChartDrawer.processor3D.view3D.getRAngAx()) {
				this.properties3d = this.cChartDrawer.processor3D.calculatePropertiesForPieCharts();
				this._recalculatePie3D();
			} else {
				this._recalculatePie3DPerspective();
			}
		} else {
			this._recalculatePie();
		}
	},

	_drawPie: function () {
		var numCache = this._getFirstRealNumCache(true);
		if(!numCache) {
			return;
		}

		var brush, pen, val, path;
		for (var i = numCache.ptCount - 1; i >= 0; i--) {
			var point = numCache.getPtByIndex(i);
			brush = point ? point.brush : null;
			pen = point ? point.pen : null;
			path = this.paths.series[i];

			this.cChartDrawer.drawPath(path, pen, brush);
		}
	},

	_recalculatePie: function () {
		var trueWidth = this.chartProp.trueWidth;
		var trueHeight = this.chartProp.trueHeight;

		var numCache = this._getFirstRealNumCache(true);
		if(!numCache) {
			return;
		}

		var sumData = this.cChartDrawer._getSumArray(numCache.pts, true);

		var radius = Math.min(trueHeight, trueWidth) / 2;
		var xCenter = this.chartProp.chartGutter._left + trueWidth / 2;
		var yCenter = this.chartProp.chartGutter._top + trueHeight / 2;

		var firstSliceAng = this.chart && this.chart.firstSliceAng ? this.chart.firstSliceAng : 0;
		this.tempAngle = Math.PI / 2 - (firstSliceAng / 180) * Math.PI;
		//рисуем против часовой стрелки, поэтому цикл с конца
		var angle;
		for (var i = numCache.ptCount - 1; i >= 0; i--) {
			var point = numCache.getPtByIndex(i);
			var val = point ? point.val : 0;
			angle = Math.abs((parseFloat(val / sumData)) * (Math.PI * 2));
			//правка связана с реализацией arcTo, где swAng зануляется и приравнивается к значению
			if(angle < 10e-16) {
				angle = 0;
			}

			if (!this.paths.series) {
				this.paths.series = [];
			}
			if (sumData === 0)//TODO стоит пересмотреть
			{
				this.paths.series[i] = this._calculateEmptySegment(radius, xCenter, yCenter);
			} else {
				this.paths.series[i] = this._calculateSegment(angle, radius, xCenter, yCenter);
			}
		}
	},

	_getFirstRealNumCache: function (returnCache) {
		var series = this.chart.series;

		//todo use getNumCache
		var numCache;
		for (var i = 0; i < series.length; i++) {
			if(returnCache) {
				numCache = series[i].val.numRef && series[i].val.numRef.numCache ? series[i].val.numRef.numCache : series[i].val.numLit;
				if (numCache) {
					return numCache;
				}
			} else {
				numCache = series[i].val.numRef && series[i].val.numRef.numCache ? series[i].val.numRef.numCache.pts : series[i].val.numLit.pts;
				if (numCache && numCache.length) {
					return numCache;
				}
			}
		}

		if(returnCache) {
			numCache = series[0].val.numRef && series[0].val.numRef.numCache ? series[0].val.numRef.numCache : series[0].val.numLit;
		} else {
			numCache = series[0].val.numRef && series[0].val.numRef.numCache ? series[0].val.numRef.numCache.pts : series[0].val.numLit.pts;
		}
		return numCache;
	},

	_calculateSegment: function (angle, radius, xCenter, yCenter) {
		if (isNaN(angle)) {
			return null;
		}

		var startAngle = (this.tempAngle);
		var endAngle = angle;

		if (radius < 0) {
			radius = 0;
		}
		var path = this._calculateArc(radius, startAngle, endAngle, xCenter, yCenter);

		this.tempAngle += angle;

		return path;
	},

	_calculateEmptySegment: function (radius, xCenter, yCenter) {

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var pxToMm = this.chartProp.pxToMM;

		var x0 = xCenter + radius * Math.cos(this.tempAngle);
		var y0 = yCenter - radius * Math.sin(this.tempAngle);

		path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
		path.lnTo(x0 / pxToMm * pathW, y0 / pxToMm * pathH);
		path.arcTo(radius / pxToMm * pathW, radius / pxToMm * pathH, this.tempAngle, this.tempAngle);
		path.lnTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);

		return pathId;
	},

	_calculateArc: function (radius, stAng, swAng, xCenter, yCenter) {

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		var pxToMm = this.chartProp.pxToMM;

		var x0 = xCenter + radius * Math.cos(stAng);
		var y0 = yCenter - radius * Math.sin(stAng);

		path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
		path.lnTo(x0 / pxToMm * pathW, y0 / pxToMm * pathH);
		path.arcTo(radius / pxToMm * pathW, radius / pxToMm * pathH, -1 * stAng * cToDeg, -1 * swAng * cToDeg);
		path.lnTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);

		return pathId;
	},

	_changeAngle: function (radius, stAng, swAng, xCenter, yCenter, depth, radius1, radius2) {
		//корректируем центр
		yCenter = yCenter - depth / 2;

		var x0 = xCenter + radius * Math.cos(stAng);
		var y0 = yCenter - radius * Math.sin(stAng);
		var kFX = radius / radius1;
		var kFY = radius / radius2;

		var cX;
		//if(this.cX !== null)
		//cX = this.cX;
		if (x0 < xCenter) {
			cX = xCenter - (xCenter - x0) / kFX;
		} else if (x0 > xCenter) {
			cX = xCenter + (x0 - xCenter) / kFX;
		} else {
			cX = xCenter;
		}

		var cY;
		//if(this.cY !== null)
		//cY = this.cY;
		if (y0 < yCenter) {
			cY = yCenter - (yCenter - y0) / kFY;
		} else if (y0 > yCenter) {
			cY = yCenter + (y0 - yCenter) / kFY;
		} else {
			cY = yCenter;
		}


		var x01 = xCenter + radius * Math.cos(stAng + swAng);
		var y01 = yCenter - radius * Math.sin(stAng + swAng);

		var aX;
		if (x01 < xCenter) {
			aX = xCenter - (xCenter - x01) / kFX;
		} else if (x01 > xCenter) {
			aX = xCenter + (x01 - xCenter) / kFX;
		} else {
			aX = xCenter;
		}


		var aY;
		if (y01 < yCenter) {
			aY = yCenter - (yCenter - y01) / kFY;
		} else if (y01 > yCenter) {
			aY = yCenter + (y01 - yCenter) / kFY;
		} else {
			aY = yCenter;
		}

		this.cX = aX;
		this.cY = aY;

		var a = Math.sqrt(Math.pow(cX - xCenter, 2) + Math.pow(cY - yCenter, 2));
		var b = Math.sqrt(Math.pow(aX - cX, 2) + Math.pow(aY - cY, 2));
		var c = Math.sqrt(Math.pow(aX - xCenter, 2) + Math.pow(aY - yCenter, 2));

		var cosNewAngle = (Math.pow(c, 2) + Math.pow(a, 2) - Math.pow(b, 2)) / (2 * c * a);
		if (cosNewAngle > 1) {
			cosNewAngle = 1;
		}

		if (cosNewAngle < -1) {
			cosNewAngle = -1;
		}

		var res;
		if (Math.abs(swAng) > Math.PI) {
			res = 2 * Math.PI - Math.acos(cosNewAngle);
		} else {
			res = Math.acos(cosNewAngle);
		}

		return res;
	},

	_calculateDLbl: function (chartSpace, ser, val, bLayout) {
		var pxToMm = this.chartProp.pxToMM;

		//TODO сделать через idx как у drawDoughnutChart!!!
		if (!this.paths.series[val]) {
			var numCache = this._getFirstRealNumCache();
			if(numCache) {
				for (var i = 0; i < numCache.length; i++) {
					if(val === numCache[i].idx) {
						val = i;
						break;
					}
				}
			}

			if(!this.paths.series[val]) {
				return;
			}
		}

		var path;
		if (this.cChartDrawer.nDimensionCount === 3) {
			if (this.paths.series[val][ser] && this.paths.series[val][ser].upPath) {
				path = this.paths.series[val][ser].upPath;
			} else if (this.paths.series[val][ser] && this.paths.series[val][ser].insidePath) {
				// функция для случая когда все значения серии равны нулю, рассчитывает позицию исходя из insidePath
				return this._calculateDLblOnlyZeroValue(val, this.paths.series[val][ser].insidePath);
			}
		} else {
			path = this.paths.series[val];
		}

		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		var getEllipseRadius = function (radius1, radius2, alpha) {
			var a = radius1 * radius2;
			var b = Math.sqrt(Math.pow(radius2, 2) * Math.pow(Math.cos(alpha), 2) +
				Math.pow(radius1, 2) * Math.pow(Math.sin(alpha), 2));

			return  a / b;
		};

		var oPath = this.cChartSpace.GetPath(path);
		var oCommand0 = oPath.getCommandByIndex(0);
		var oCommand1 = oPath.getCommandByIndex(1);
		var oCommand2 = oPath.getCommandByIndex(2);

		var centerX = oCommand0.X;
		var centerY = oCommand0.Y;

		var radius = oCommand2.hR;
		var stAng = oCommand2.stAng;
		var swAng = oCommand2.swAng;

		if (this.cChartDrawer.nDimensionCount === 3 && oCommand2.wR) {
			radius = getEllipseRadius(oCommand2.hR, oCommand2.wR, -1 * stAng - swAng / 2 - Math.PI / 2);
		}

		var _numCache = this.chart.series[0].val.numRef ? this.chart.series[0].val.numRef.numCache : this.chart.series[0].val.numLit;
		var point = _numCache ? _numCache.getPtByIndex(val) : null;

		if (!point || !point.compiledDlb) {
			return;
		}

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var tempCenterX, tempCenterY;

		//TODO высчитать позиции, как в екселе +  ограничения
		var oPos;
		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.bestFit: {
				oPos = this._calculateBestFitPosition(stAng, swAng, radius, width, height, centerX, centerY, bLayout);
				if (!oPos.bError) {
					centerX = oPos.fX;
					centerY = oPos.fY;
				} else {
					centerX = centerX + (radius / 2) * Math.cos(-1 * stAng - swAng / 2) - width / 2;
					centerY = centerY - (radius / 2) * Math.sin(-1 * stAng - swAng / 2) - height / 2;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				centerX = centerX + (radius / 2) * Math.cos(-1 * stAng - swAng / 2) - width / 2;
				centerY = centerY - (radius / 2) * Math.sin(-1 * stAng - swAng / 2) - height / 2;
				break;
			}
			case c_oAscChartDataLabelsPos.inBase: {
				centerX = centerX + (radius / 2) * Math.cos(-1 * stAng - swAng / 2) - width / 2;
				centerY = centerY - (radius / 2) * Math.sin(-1 * stAng - swAng / 2) - height / 2;
				break;
			}
			case c_oAscChartDataLabelsPos.inEnd: {
				oPos = this._calculateInEndDLblPosition(stAng, swAng, radius, width, height, centerX, centerY);
				if (!oPos.bError) {
					centerX = oPos.fX;
					centerY = oPos.fY;
					break;
				}
				tempCenterX = centerX + (radius) * Math.cos(-1 * stAng - swAng / 2);
				tempCenterY = centerY - (radius) * Math.sin(-1 * stAng - swAng / 2);

				if (tempCenterX < centerX && tempCenterY < centerY) {
					centerX = tempCenterX;
					centerY = tempCenterY;
				} else if (tempCenterX > centerX && tempCenterY < centerY) {
					centerX = tempCenterX - width;
					centerY = tempCenterY;
				} else if (tempCenterX < centerX && tempCenterY > centerY) {
					centerX = tempCenterX;
					centerY = tempCenterY - height;
				} else {
					centerX = tempCenterX - width;
					centerY = tempCenterY - height;
				}
				break;
			}
			case c_oAscChartDataLabelsPos.outEnd: {
				tempCenterX = centerX + (radius) * Math.cos(-1 * stAng - swAng / 2);
				tempCenterY = centerY - (radius) * Math.sin(-1 * stAng - swAng / 2);

				if (tempCenterX < centerX && tempCenterY < centerY) {
					centerX = tempCenterX - width;
					centerY = tempCenterY - height;
				} else if (tempCenterX > centerX && tempCenterY < centerY) {
					centerX = tempCenterX;
					centerY = tempCenterY - height;
				} else if (tempCenterX < centerX && tempCenterY > centerY) {
					centerX = tempCenterX - width;
					centerY = tempCenterY;
				} else {
					centerX = tempCenterX;
					centerY = tempCenterY;
				}
				break;
			}
		}
		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_calculateDLblOnlyZeroValue: function (val, path) {
		var numCache = this._getFirstRealNumCache();
		var oCommand1, calcPath, oCommand0;
		// циклом находим крайнюю точку
		for (var i = 0; i < this.paths.series.length; i++) {
			calcPath = null;
			if (this.paths.series[i] && numCache[i] && null != numCache[i].val && this.paths.series[i][numCache[i].val]) {
				calcPath = this.paths.series[i][numCache[i].val].insidePath;
				calcPath = null != calcPath && this.cChartSpace.GetPath(calcPath).getCommandByIndex(1);
				if (calcPath) {
					oCommand1 = calcPath;
				}
			}
		}

		if (!oCommand1) {
			return;
		}
		if (!AscFormat.isRealNumber(path)) {
			return;
		}

		path = this.cChartSpace.GetPath(path);
		oCommand0 = path.getCommandByIndex(0);

		var x = oCommand0.X + (oCommand1.X - oCommand0.X) / 2;
		var y = oCommand0.Y + (oCommand1.Y - oCommand0.Y) / 2;

		var _numCache = this.chart.series[0].val.numRef ? this.chart.series[0].val.numRef.numCache : this.chart.series[0].val.numLit;
		var point = _numCache ? _numCache.getPtByIndex(val) : null;

		if (!point || !point.compiledDlb) {
			return;
		}

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		return { x: x - width / 2, y: y - height / 2 }
	},

	//****fast calulate and drawing(for switch on slow drawing: change name function _Slow)
	_recalculatePie3D: function () {
		var trueWidth = this.chartProp.trueWidth;
		var trueHeight = this.chartProp.trueHeight;

		var numCache = this._getFirstRealNumCache(true);
		if(!numCache || !numCache.pts) {
			return;
		}
		this.empty3DSegment = numCache.pts.length === 1;

		var sumData = this.cChartDrawer._getSumArray(numCache.pts, true);

		var radius = Math.min(trueHeight, trueWidth) / 2;
		if (radius < 0) {
			radius = 0;
		}

		var xCenter = this.chartProp.chartGutter._left + trueWidth / 2;
		var yCenter = this.chartProp.chartGutter._top + trueHeight / 2;

		var startAngle = this.cChartDrawer.processor3D.angleOy ? this.cChartDrawer.processor3D.angleOy : 0;
		var startAngle3D = startAngle !== 0 && startAngle !== undefined ? this._changeAngle(radius, Math.PI / 2, startAngle, xCenter, yCenter, this.properties3d.depth, this.properties3d.radius1, this.properties3d.radius2) : 0;

		this.angleFor3D = Math.PI / 2 - startAngle3D;
		startAngle = startAngle + Math.PI / 2;

		for (var i = numCache.ptCount - 1; i >= 0; i--) {
			var point = numCache.getPtByIndex(i);
			var val = point ? point.val : 0;
			var partOfSum = val / sumData;
			var swapAngle = Math.abs((parseFloat(partOfSum)) * (Math.PI * 2));

			if (!this.paths.series) {
				this.paths.series = [];
			}

			this.paths.series[i] = this._calculateSegment3D(startAngle, swapAngle, radius, xCenter, yCenter);

			startAngle += swapAngle;
		}
	},

	_recalculatePie3DPerspective: function () {
		var left = this.chartProp.chartGutter._left;
		var right = this.chartProp.chartGutter._right;
		var top = this.chartProp.chartGutter._top;
		var bottom = this.chartProp.chartGutter._bottom;

		var trueWidth = this.chartProp.trueWidth;

		var widthCanvas = this.chartProp.widthCanvas;
		var heightCanvas = this.chartProp.heightCanvas;

		var height = heightCanvas - (top + bottom);
		var width = widthCanvas - (left + right);

		var tempDepth = this.cChartDrawer.processor3D.depthPerspective;

		var x1 = left, y1 = top + height, z1 = 0;
		var x2 = left, y2 = top, z2 = 0;
		var x3 = left + width, y3 = top, z3 = 0;
		var x4 = left + width, y4 = top + height, z4 = 0;

		var x5 = left, y5 = top + height, z5 = tempDepth;
		var x6 = left, y6 = top, z6 = tempDepth;
		var x7 = left + width, y7 = top, z7 = tempDepth;
		var x8 = left + width, y8 = top + height, z8 = tempDepth;

		var point1 = this.cChartDrawer._convertAndTurnPoint(x1, y1, z1);
		var point2 = this.cChartDrawer._convertAndTurnPoint(x2, y2, z2);
		//var point3 = this.cChartDrawer._convertAndTurnPoint(x3, y3, z3);
		//var point4 = this.cChartDrawer._convertAndTurnPoint(x4, y4, z4);

		var point5 = this.cChartDrawer._convertAndTurnPoint(x5, y5, z5);
		var point6 = this.cChartDrawer._convertAndTurnPoint(x6, y6, z6);
		//var point7 = this.cChartDrawer._convertAndTurnPoint(x7, y7, z7);
		//var point8 = this.cChartDrawer._convertAndTurnPoint(x8, y8, z8);


		var radius3D1 = (z6 - z2) / 2;
		var radius3D2 = (z5 - z1) / 2;

		var center3D1 = new Point3D(x2 + ((x3 - x2) / 2), y2, z2 + (radius3D1));
		var center3D2 = new Point3D(x1 + ((x4 - x1) / 2), y1, z1 + (radius3D2));

		var pointCenter1 = this.cChartDrawer._convertAndTurnPoint(center3D1.x, center3D1.y, center3D1.z);
		var pointCenter2 = this.cChartDrawer._convertAndTurnPoint(center3D2.x, center3D2.y, center3D2.z);

		//TEST DRAW FRAME
		//this._calculateTestFrame(point1, point2, point3, point4, point5, point6, point7, point8);

		var test1 = this.cChartDrawer._convertAndTurnPoint(x2, y2, center3D1.z);
		var test2 = this.cChartDrawer._convertAndTurnPoint(x3, y3, center3D1.z);

		var test11 = this.cChartDrawer._convertAndTurnPoint(x4, y4, center3D2.z);
		var test22 = this.cChartDrawer._convertAndTurnPoint(x1, y1, center3D2.z);

		var radius11 = Math.abs((test2.x - test1.x) / 2);
		var radius12 = Math.abs((point6.y - point2.y) / 2);
		var radius21 = Math.abs((test22.x - test11.x) / 2);
		var radius22 = Math.abs((point5.y - point1.y) / 2);

		var center1 = {
			x: this.chartProp.chartGutter._left + trueWidth / 2,
			y: (point2.y - point6.y) / 2 + point6.y
		};
		var center2 = {
			x: this.chartProp.chartGutter._left + trueWidth / 2,
			y: (point1.y - point5.y) / 2 + point5.y
		};

		var angles1 = this._calculateAngles3DPerspective(center1.x, center1.y, radius11, radius12, radius3D1,
			center3D1);
		var angles2 = this._calculateAngles3DPerspective(center2.x, center2.y, radius21, radius22, radius3D2,
			center3D2);

		if (!this.paths.series) {
			this.paths.series = [];
		}

		for (var i = 0; i < angles1.length; i++) {
			var start = angles1[i].start;
			var start1 = angles2[i].start;

			if (i === angles1.length - 1) {
				var end = angles1[0].start + 2 * Math.PI;
				angles1[i].swap = end - start;

				var end1 = angles2[0].start + 2 * Math.PI;
				angles2[i].swap = end1 - start1;
			}

			var paths = this._calculateSegment3DPerspective(radius11, radius12, radius21, radius22, angles1[i],
				angles2[i], center1, center2, pointCenter1, pointCenter2, Math.sign(point2.y - point6.y));

			if (null === this.tempDrawOrder) {
				this.tempDrawOrder = Math.sign(point2.y - point6.y) < 0 ? true : null;
			}

			if (!this.paths.series[angles1.length - i - 1]) {
				this.paths.series[angles1.length - i - 1] = [];
			}

			this.paths.series[angles1.length - i - 1].push(paths);
		}
	},

	_calculateAngles3DPerspective: function (xCenter, yCenter, radius1, radius2, radius3D1, center3D1) {
		var t = this;
		var widthCanvas = this.chartProp.widthCanvas;

		var numCache = this._getFirstRealNumCache(true);
		if(!numCache) {
			return;
		}
		this.empty3DSegment = numCache.pts.length === 1;

		var sumData = this.cChartDrawer._getSumArray(numCache.pts, true);

		var startAngle = Math.PI / 2;
		var newStartAngle = startAngle;

		var oView3D = this.cChartSpace.chart.getView3d();
		var firstAngle = oView3D && oView3D.rotY ? (-oView3D.rotY / 360) * (Math.PI * 2) : 0;


		var getAngleByCoordsSidesTriangle = function (aC, bC, cC) {
			var res;

			var a = Math.sqrt(Math.pow(aC.x, 2) + Math.pow(aC.y, 2));
			var b = Math.sqrt(Math.pow(bC.x, 2) + Math.pow(bC.y, 2));
			var c = Math.sqrt(Math.pow(cC.x, 2) + Math.pow(cC.y, 2));
			res = Math.acos((Math.pow(b, 2) + Math.pow(c, 2) - Math.pow(a, 2)) / (2 * b * c));

			return res;
		};

		var getPointsByAngle = function (angle) {
			var point1 = t.cChartDrawer._convertAndTurnPoint(center3D1.x + radius3D1 * Math.cos(angle), center3D1.y,
				center3D1.z + radius3D1 * Math.sin(angle));

			var y11 = point1.y;
			var x11 = Math.sqrt(
				Math.abs(Math.pow(radius1, 2) * (1 - (Math.pow(y11 - (yCenter), 2) / Math.pow(radius2, 2)))));

			var x111;
			if ((angle <= 3 * Math.PI / 2 && angle >= Math.PI / 2) || (angle >= -3 * Math.PI / 2 && angle <= -Math.PI / 2)) {
				x111 = xCenter - x11;
			} else {
				x111 = widthCanvas - (xCenter - x11)
			}

			return {x: x111, y: y11};
		};

		var angles = [];
		for (var i = numCache.ptCount; i >= 0; i--) {
			//рассчитываем угол
			var swapAngle;
			if (i === numCache.ptCount) {
				swapAngle = firstAngle;
			} else {
				var point = numCache.getPtByIndex(i);
				var val = point ? point.val : 0;
				var partOfSum = val / sumData;
				swapAngle = Math.abs((parseFloat(partOfSum)) * (Math.PI * 2));
			}


			var tempSwapAngle = 0, newSwapAngle = 0, tempStartAngle = startAngle;
			while (true) {
				if (i === numCache.length && swapAngle < 0) {
					if (tempStartAngle - Math.PI / 2 > startAngle + swapAngle) {
						tempSwapAngle = -Math.PI / 2;
					} else {
						tempSwapAngle = (startAngle + swapAngle) - tempStartAngle;
					}
				} else {
					if (tempStartAngle + Math.PI / 2 < startAngle + swapAngle) {
						tempSwapAngle = Math.PI / 2;
					} else {
						tempSwapAngle = (startAngle + swapAngle) - tempStartAngle;
					}
				}


				var p1 = getPointsByAngle(tempStartAngle);
				var p2 = getPointsByAngle(tempStartAngle + tempSwapAngle);
				newSwapAngle += getAngleByCoordsSidesTriangle({x: p2.x - p1.x, y: p2.y - p1.y},
					{x: xCenter - p1.x, y: yCenter - p1.y}, {x: xCenter - p2.x, y: yCenter - p2.y});


				if (i === numCache.ptCount && swapAngle < 0) {
					if (tempStartAngle - Math.PI / 2 > startAngle + swapAngle) {
						tempStartAngle -= Math.PI / 2;
					} else {
						if (i !== numCache.ptCount) {
							angles.push(
								{start: newStartAngle, swap: newSwapAngle, end: newStartAngle + newSwapAngle});
						}

						break;
					}
				} else {
					if (tempStartAngle + Math.PI / 2 < startAngle + swapAngle) {
						tempStartAngle += Math.PI / 2;
					} else {
						if (i !== numCache.ptCount) {
							angles.push(
								{start: newStartAngle, swap: newSwapAngle, end: newStartAngle + newSwapAngle});
						}

						break;
					}
				}

			}

			startAngle += swapAngle;
			if (i === numCache.ptCount) {
				if (swapAngle < 0) {
					newStartAngle -= newSwapAngle;
				} else {
					newStartAngle += newSwapAngle;
				}

			} else {
				newStartAngle += newSwapAngle;
			}
		}

		return angles;
	},

	_calculateArc3D: function (radius, stAng, swAng, xCenter, yCenter, bIsNotDrawFrontFace, depth, radius1, radius2) {
		var properties = this.cChartDrawer.processor3D.calculatePropertiesForPieCharts();
		var oChartSpace = this.cChartSpace;

		depth = !depth ? properties.depth : depth;
		radius1 = !radius1 ? properties.radius1 : radius1;
		radius2 = !radius2 ? properties.radius2 : radius2;


		var pxToMm = this.chartProp.pxToMM;
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		swAng = this._changeAngle(radius, stAng, swAng, xCenter, yCenter, depth, radius1, radius2);
		stAng = this.angleFor3D;
		//корректируем центр
		yCenter = yCenter - depth / 2;


		var getSegmentPoints = function (startAng, endAng) {
			var radiusSpec = (radius1 * radius2) / Math.sqrt(
					Math.pow(radius2, 2) * Math.pow((Math.cos(startAng)), 2) +
					Math.pow(radius1, 2) * Math.pow(Math.sin(startAng), 2));
			var radiusSpec2 = (radius1 * radius2) / Math.sqrt(
					Math.pow(radius2, 2) * Math.pow((Math.cos(endAng)), 2) +
					Math.pow(radius1, 2) * Math.pow(Math.sin(endAng), 2));

			var x0 = (xCenter + radiusSpec * Math.cos(startAng));
			var y0 = (yCenter - radiusSpec * Math.sin(startAng));

			var x1 = (xCenter + radiusSpec * Math.cos(startAng));
			var y1 = ((yCenter + depth) - radiusSpec * Math.sin(startAng));

			var x2 = (xCenter + radiusSpec2 * Math.cos(endAng));
			var y2 = (yCenter - radiusSpec2 * Math.sin(endAng));

			var x3 = (xCenter + radiusSpec2 * Math.cos(endAng));
			var y3 = ((yCenter + depth) - radiusSpec2 * Math.sin(endAng));

			return {x0: x0, y0: y0, x1: x1, y1: y1, x2: x2, y2: y2, x3: x3, y3: y3};
		};

		var breakAng = function (startAng, swapAng) {
			var res = [];
			var endAng = startAng + swapAng;
			var _compare = function (_ang) {
				var deltaAng = 0.00000000000001;
				return startAng < _ang && endAng > _ang && _ang - startAng > deltaAng && endAng -_ang > deltaAng;
			};

			res.push({angle: startAng});
			if (_compare(-2 * Math.PI)) {
				res.push({angle: -2 * Math.PI});
			}
			if (_compare(-Math.PI)) {
				res.push({angle: -Math.PI});
			}
			if (startAng < 0 && endAng > 0) {
				res.push({angle: 0});
			}
			if (_compare(Math.PI)) {
				res.push({angle: Math.PI});
			}
			if (_compare(2 * Math.PI)) {
				res.push({angle: 2 * Math.PI});
			}
			res.push({angle: endAng});

			return res;
		};

		var calculateInsideFaces = function (startAng, swapAng) {

			var pathId = oChartSpace.AllocPath();
			var path = oChartSpace.GetPath(pathId);

			var endAng = startAng + swapAng;
			var p = getSegmentPoints(startAng, endAng);

			path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
			path.lnTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
			path.lnTo(p.x1 / pxToMm * pathW, p.y1 / pxToMm * pathH);
			path.lnTo(xCenter / pxToMm * pathW, (yCenter + depth) / pxToMm * pathH);

			path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
			path.lnTo(p.x2 / pxToMm * pathW, p.y2 / pxToMm * pathH);
			path.lnTo(p.x3 / pxToMm * pathW, p.y3 / pxToMm * pathH);
			path.lnTo(xCenter / pxToMm * pathW, (yCenter + depth) / pxToMm * pathH);


			return pathId;
		};

		var calculateFrontFace = function (startAng, swapAng) {

			var pathId = oChartSpace.AllocPath();
			var path = oChartSpace.GetPath(pathId);

			var endAng = startAng + swapAng;
			var p = getSegmentPoints(startAng, endAng);

			path.moveTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
			path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * startAng * cToDeg, -1 * swapAng * cToDeg);
			path.lnTo(p.x3 / pxToMm * pathW, p.y3 / pxToMm * pathH);
			path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * startAng * cToDeg - 1 * swapAng * cToDeg, 1 * swapAng * cToDeg);
			path.lnTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);

			return pathId;
		};

		var calculateUpFace = function (startAng, swapAng, empty) {

			var pathId = oChartSpace.AllocPath();
			var path = oChartSpace.GetPath(pathId);

			var endAng = startAng + swapAng;
			var p = getSegmentPoints(startAng, endAng);

			path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
			if (empty) {
				path.moveTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
				path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * stAng * cToDeg, -1 * swapAng * cToDeg);
				path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
			} else {
				path.lnTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
				path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * stAng * cToDeg, -1 * swapAng * cToDeg);
				path.lnTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
			}

			return pathId;
		};

		var calculateDownFace = function (startAng, swapAng) {
			var pathId = oChartSpace.AllocPath();
			var path = oChartSpace.GetPath(pathId);

			var endAng = startAng + swapAng;
			var p = getSegmentPoints(startAng, endAng);

			path.moveTo(xCenter / pxToMm * pathW, (yCenter + depth) / pxToMm * pathH);
			path.lnTo(p.x0 / pxToMm * pathW, (p.y0 + depth) / pxToMm * pathH);
			path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * stAng * cToDeg,
				-1 * swapAng * cToDeg);
			path.lnTo(xCenter / pxToMm * pathW, (yCenter + depth) / pxToMm * pathH);


			return pathId;
		};

		//FRONT FACES
		//break front faces
		var arrAngles = breakAng(stAng, swAng);
		var frontPath = [];
		for (var i = 1; i < arrAngles.length; i++) {
			var start = arrAngles[i - 1].angle;
			var end = arrAngles[i].angle;
			var swap = end - start;

			if ((start >= 0 && start >= Math.PI && start <= 2 * Math.PI) || (start < 0 && start >= -Math.PI && start <= 0)) {
				frontPath.push(calculateFrontFace(start, swap));
			}
		}

		//INSIDE FACES
		var insidePath = calculateInsideFaces(stAng, swAng);

		//UP FACE
		var upPath = calculateUpFace(stAng, swAng, this.empty3DSegment);

		//DOWN FACE
		var downPath = calculateDownFace(stAng, swAng);

		this.angleFor3D += swAng;

		return {frontPath: frontPath, upPath: upPath, insidePath: insidePath, downPath: downPath};
	},

	_calculateSegment3D: function (startAngle, swapAngle, radius, xCenter, yCenter) {
		if (isNaN(swapAngle)) {
			return null;
		}

		if (radius < 0) {
			radius = 0;
		}

		var path = [];
		path.push(this._calculateArc3D(radius, startAngle, swapAngle, xCenter, yCenter));

		return path;
	},

	_calculateSegment3DPerspective: function (radiusUp1, radiusUp2, radiusDown1, radiusDown2, angles1, angles2, center1, center2, pointCenter1, pointCenter2, upFaceSign) {
		var xCenter = center1.x, yCenter = center1.y, xCenter1 = center2.x, yCenter1 = center2.y;
		var startAngle1 = angles1.start, swapAngle1 = angles1.swap, startAngle2 = angles2.start, swapAngle2 = angles2.swap;

		var pxToMm = this.chartProp.pxToMM;
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var oThis = this;


		var getSegmentPoints = function (startAng, endAng, startAng2, endAng2) {
			var radiusSpec = (radiusUp1 * radiusUp2) / Math.sqrt(Math.pow(radiusUp2, 2) * Math.pow((Math.cos(startAng)), 2) + Math.pow(radiusUp1, 2) * Math.pow(Math.sin(startAng), 2));
			var radiusSpec2 = (radiusUp1 * radiusUp2) / Math.sqrt(Math.pow(radiusUp2, 2) * Math.pow((Math.cos(endAng)), 2) + Math.pow(radiusUp1, 2) * Math.pow(Math.sin(endAng), 2));

			var radiusSpec11 = (radiusDown1 * radiusDown2) / Math.sqrt(Math.pow(radiusDown2, 2) * Math.pow((Math.cos(startAng2)), 2) + Math.pow(radiusDown1, 2) * Math.pow(Math.sin(startAng2), 2));
			var radiusSpec12 = (radiusDown1 * radiusDown2) / Math.sqrt(Math.pow(radiusDown2, 2) * Math.pow((Math.cos(endAng2)), 2) + Math.pow(radiusDown1, 2) * Math.pow(Math.sin(endAng2), 2));

			var x0 = (xCenter + radiusSpec * Math.cos(startAng));
			var y0 = (yCenter - radiusSpec * Math.sin(startAng));

			var x1 = (xCenter1 + radiusSpec11 * Math.cos(startAng2));
			var y1 = ((yCenter1) - radiusSpec11 * Math.sin(startAng2));

			var x2 = (xCenter + radiusSpec2 * Math.cos(endAng));
			var y2 = (yCenter - radiusSpec2 * Math.sin(endAng));

			var x3 = (xCenter1 + radiusSpec12 * Math.cos(endAng2));
			var y3 = ((yCenter1) - radiusSpec12 * Math.sin(endAng2));

			return {x0: x0, y0: y0, x1: x1, y1: y1, x2: x2, y2: y2, x3: x3, y3: y3};
		};

		var breakAng = function (startAng, swapAng) {
			var res = [];
			var endAng = startAng + swapAng;

			res.push({angle: startAng});

			var tempStartAng = startAng;
			var tempEndAng = endAng;
			var tempPI = Math.PI;

			if (tempStartAng <= -2 * tempPI && tempEndAng > -2 * tempPI) {
				res.push({angle: -2 * Math.PI});
			}
			if (tempStartAng <= -tempPI && tempEndAng > -tempPI) {
				res.push({angle: -Math.PI});
			}
			if (tempStartAng <= 0 && tempEndAng > 0) {
				res.push({angle: 0});
			}
			if (tempStartAng <= tempPI && tempEndAng > tempPI) {
				res.push({angle: Math.PI});
			}
			if (tempStartAng <= 2 * tempPI && tempEndAng > 2 * tempPI) {
				res.push({angle: 2 * Math.PI});
			}
			res.push({angle: endAng});

			return res;
		};

		var calculateInsideFaces = function (startAng1, swapAng1, startAng2, swapAng2) {

			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			var endAng1 = startAng1 + swapAng1;
			var endAng2 = startAng2 + swapAng2;
			var p = getSegmentPoints(startAng1, endAng1, startAng2, endAng2);

			path.moveTo(pointCenter1.x / pxToMm * pathW, pointCenter1.y / pxToMm * pathH);
			path.lnTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
			path.lnTo(p.x1 / pxToMm * pathW, p.y1 / pxToMm * pathH);
			path.lnTo(pointCenter2.x / pxToMm * pathW, (pointCenter2.y) / pxToMm * pathH);

			path.moveTo(pointCenter1.x / pxToMm * pathW, pointCenter1.y / pxToMm * pathH);
			path.lnTo(p.x2 / pxToMm * pathW, p.y2 / pxToMm * pathH);
			path.lnTo(p.x3 / pxToMm * pathW, p.y3 / pxToMm * pathH);
			path.lnTo(pointCenter2.x / pxToMm * pathW, (pointCenter2.y) / pxToMm * pathH);

			return pathId;
		};

		var calculateFrontFace = function (startAng, swapAng, startAng2, swapAng2) {

			if(isNaN(startAng) || isNaN(swapAng)) {
				return null;
			}

			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			var endAng = startAng + swapAng;
			var endAng2 = startAng2 + swapAng2;

			var p = getSegmentPoints(startAng, endAng, startAng2, endAng2);
			//var p2 = getSegmentPoints(startAng2, endAng2);

			path.moveTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);
			path.arcTo(radiusUp1 / pxToMm * pathW, radiusUp2 / pxToMm * pathH, -1 * startAng * cToDeg, -1 * swapAng * cToDeg);
			path.lnTo(p.x3 / pxToMm * pathW, p.y3 / pxToMm * pathH);
			path.arcTo(radiusDown1 / pxToMm * pathW, radiusDown2 / pxToMm * pathH, -1 * startAng2 * cToDeg - 1 * swapAng2 * cToDeg, 1 * swapAng2 * cToDeg);
			path.lnTo(p.x0 / pxToMm * pathW, p.y0 / pxToMm * pathH);


			return pathId;
		};

		var calculateUpFace = function (startAng, swapAng, empty) {

			if(isNaN(startAng) ||  isNaN(swapAng)) {
				return null;
			}

			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			var radiusSpec = (radiusUp1 * radiusUp2) / Math.sqrt(Math.pow(radiusUp2, 2) * Math.pow((Math.cos(startAng)), 2) + Math.pow(radiusUp1, 2) * Math.pow(Math.sin(startAng), 2));

			var x0 = (xCenter + radiusSpec * Math.cos(startAng));
			var y0 = (yCenter - radiusSpec * Math.sin(startAng));

			path.moveTo(pointCenter1.x / pxToMm * pathW, pointCenter1.y / pxToMm * pathH);
			if (empty) {
				path.moveTo(x0 / pxToMm * pathW, y0 / pxToMm * pathH);
				path.arcTo(radiusUp1 / pxToMm * pathW, radiusUp2 / pxToMm * pathH, -1 * startAng * cToDeg, -1 * swapAng * cToDeg);
				path.moveTo(pointCenter1.x / pxToMm * pathW, pointCenter1.y / pxToMm * pathH);
			} else {
				path.lnTo(x0 / pxToMm * pathW, y0 / pxToMm * pathH);
				path.arcTo(radiusUp1 / pxToMm * pathW, radiusUp2 / pxToMm * pathH, -1 * startAng * cToDeg, -1 * swapAng * cToDeg);
				path.lnTo(pointCenter1.x / pxToMm * pathW, pointCenter1.y / pxToMm * pathH);
			}


			return pathId;
		};

		var calculateDownFace = function (startAng, swapAng) {

			if(isNaN(startAng) ||  isNaN(swapAng)) {
				return null;
			}

			var pathId = oThis.cChartSpace.AllocPath();
			var path = oThis.cChartSpace.GetPath(pathId);

			var radiusSpec = (radiusDown1 * radiusDown2) / Math.sqrt(Math.pow(radiusDown2, 2) * Math.pow((Math.cos(startAng)), 2) + Math.pow(radiusDown1, 2) * Math.pow(Math.sin(startAng), 2));
			//var radiusSpec2 = (radius11 * radius2) /  Math.sqrt(Math.pow(radius2, 2) * Math.pow((Math.cos(endAng)), 2) + Math.pow(radius11, 2) * Math.pow(Math.sin(endAng),2))

			var x = (xCenter1 + radiusSpec * Math.cos(startAng));
			var y = (yCenter1 - radiusSpec * Math.sin(startAng));

			path.moveTo(pointCenter2.x / pxToMm * pathW, (pointCenter2.y) / pxToMm * pathH);
			path.lnTo(x / pxToMm * pathW, y / pxToMm * pathH);
			path.arcTo(radiusDown1 / pxToMm * pathW, radiusDown2 / pxToMm * pathH, -1 * startAng * cToDeg, -1 * swapAng * cToDeg);
			path.lnTo(pointCenter2.x / pxToMm * pathW, (pointCenter2.y) / pxToMm * pathH);


			return pathId;
		};

		//FRONT FACES
		//break front faces
		var arrAngles = breakAng(startAngle1, swapAngle1);
		var arrAngles2 = breakAng(startAngle2, swapAngle2);
		var frontPath = [];
		for (var i = 1; i < arrAngles.length; i++) {
			var start = arrAngles[i - 1].angle;
			var end = arrAngles[i].angle;
			var swap = end - start;

			var start2 = arrAngles2[i - 1].angle;
			var end2;
			if (!arrAngles2[i]) {
				end2 = arrAngles2[i - 1].angle;
			} else {
				end2 = arrAngles2[i].angle;
			}

			var swap2 = end2 - start2;

			if ((start >= 0 && start >= Math.PI && start <= 2 * Math.PI) || (start < 0 && start >= -Math.PI && start <= 0)) {
				frontPath.push(calculateFrontFace(upFaceSign * start, upFaceSign * swap, start2, swap2));
			}
		}

		//INSIDE FACES
		var insidePath = calculateInsideFaces(upFaceSign * startAngle1, upFaceSign * swapAngle1, startAngle2,
			swapAngle2);

		//UP FACE
		var upPath = calculateUpFace(upFaceSign * startAngle1, upFaceSign * swapAngle1, this.empty3DSegment);

		//DOWN FACE
		var downPath = calculateDownFace(startAngle2, swapAngle2);

		return {frontPath: frontPath, upPath: upPath, insidePath: insidePath, downPath: downPath};
	},


	_drawPie3D: function () {
		var numCache = this._getFirstRealNumCache(true);
		var t = this;
		var shade = "shade";
		var shadeValue = 35000;

		var drawPath = function (path, pen, brush, isShadePen, isShadeBrush) {
			if (path) {
				if (brush) {
					var props = t.cChartSpace.getParentObjects();
					var duplicateBrush = brush.createDuplicate();
					var cColorMod = new AscFormat.CColorMod(shade, shadeValue);
					if (duplicateBrush) {
						duplicateBrush.addColorMod(cColorMod);
						duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master, new AscFormat.CUniColor().RGBA, t.cChartSpace.clrMapOvr);
					}
					if (isShadePen) {
						pen = AscFormat.CreatePenFromParams(duplicateBrush, undefined, undefined, undefined,
							undefined, 0);
					}
					if (isShadeBrush) {
						brush = duplicateBrush;
					}
				}

				t.cChartDrawer.drawPath(path, pen, brush);
			}

		};

		var _firstPoint = numCache.getPtByIndex(0);
		var pen = _firstPoint ? _firstPoint.pen : null;
		drawPath(this.paths.test, pen, null);

		var sides = {down: 0, inside: 1, up: 2, front: 3};
		var drawPaths = function (side) {
			for (var i = 0, len = numCache.ptCount; i < len; i++) {
				var point = numCache.getPtByIndex(i);
				var brush = point ? point.brush : null;
				var pen = point ? point.pen : null;
				var path = t.paths.series[i];

				if (path) {
					for (var j = path.length - 1; j >= 0; j--) {
						if (side === sides.down) {
							drawPath(path[j].downPath, pen, null);
						} else if (side === sides.inside) {
							//выставляю закругленные соединения
							if (pen && pen.Join) {
								pen = pen.createDuplicate();
								pen.Join.type = Asc['c_oAscLineJoinType'].Round;
							}

							drawPath(path[j].insidePath, pen, brush, null, true);
						} else if (side === sides.up) {
							drawPath(path[j].upPath, pen, brush);
						} else if (side === sides.front) {
							for (var k = 0; k < path[j].frontPath.length; k++) {
								drawPath(path[j].frontPath[k], pen, brush, true, true);
							}
						}
					}
				}
			}
		};

		drawPaths(sides.down);
		drawPaths(sides.inside);
		if (this.tempDrawOrder !== null) {
			drawPaths(sides.up);
			drawPaths(sides.front);
		} else {
			drawPaths(sides.front);
			drawPaths(sides.up);
		}
	},

	//best fit DLbl
	_calculateBestFitPosition: function (fStartAngle, fSweepAngle, fRadius, fWidth, fHeight, fCenterX, fCenterY, bLayout) {
		var fStartAngle_ = fStartAngle;
		var fEndAngle = fStartAngle + fSweepAngle;
		if (bLayout) {
			return this._calculateBestFitPositionOuter(fStartAngle_, fEndAngle, fRadius, fWidth, fHeight, fCenterX, fCenterY);
		}
		var oRet = this._calculateBestFitPositionInner(fStartAngle_, fEndAngle, fRadius, fWidth, fHeight, fCenterX, fCenterY);
		if (!oRet.bError) {
			if (AscFormat.fCheckBoxIntersectionSegment(oRet.fX, oRet.fY, fWidth, fHeight, fCenterX, fCenterY, fCenterX + fRadius * Math.cos(fStartAngle_), fCenterY + fRadius * Math.sin(fStartAngle_)) || AscFormat.fCheckBoxIntersectionSegment(oRet.fX, oRet.fY, fWidth, fHeight, fCenterX, fCenterY, fCenterX + fRadius * Math.cos(fEndAngle), fCenterY + fRadius * Math.sin(fEndAngle))) {
				oRet.bError = true;
			}
		}
		if (oRet.bError) {
			return this._calculateBestFitPositionOuter(fStartAngle_, fEndAngle, fRadius, fWidth, fHeight, fCenterX, fCenterY);
		}
		return oRet;
	},

	_calculateBestFitPositionInner: function(fStartAngle, fEndAngle, fPieRadius, fLabelWidth, fLabelHeight, fCenterX, fCenterY){
		var oResult = {bError: true};
		var fBisectAngle = AscFormat.normalizeRotate((fStartAngle + fEndAngle)/2.0);
		
		if(AscFormat.fApproxEqual(fBisectAngle, 0) || AscFormat.fApproxEqual(fBisectAngle, Math.PI/2) || AscFormat.fApproxEqual(fBisectAngle, Math.PI) || AscFormat.fApproxEqual(fBisectAngle, 3*Math.PI/2)){
			return this._calculateInEndDLblPosition(fStartAngle, fEndAngle - fStartAngle, fPieRadius, fLabelWidth, fLabelHeight, fCenterX, fCenterY);
		}
		var fBisectAngle2 = AscFormat.normalizeRotate(fBisectAngle + Math.PI/4) - Math.PI/4; 
		var nIndexArea = ((fBisectAngle2 + Math.PI/4)/(Math.PI/2)) >> 0;
				
		
		var fLengthCoeff =  ((fBisectAngle2 + Math.PI/4) - (Math.PI/2)*nIndexArea)/(Math.PI/2);		
				
		var fXs = fCenterX + fPieRadius*Math.cos(fBisectAngle);
		var fYs = fCenterY + fPieRadius*Math.sin(fBisectAngle);
		var fDeltaX, fDeltaY, oSolvation;

        switch(nIndexArea){
            case 0:{
                if(fBisectAngle2 < 0){
                    fDeltaX = fLabelWidth;
                    fDeltaY = -(1 - fLengthCoeff)*fLabelHeight;
                }
                else{
                    fDeltaX = fLabelWidth;
                    fDeltaY = fLabelHeight*fLengthCoeff;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 > 0 && oSolvation.x1 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX);
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
                    }
                    else if(oSolvation.x2 > 0 && oSolvation.x2 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX);
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
                    }
                }
                break;
            }
            case 1:{
                if(fBisectAngle < Math.PI/2){
                    fDeltaX = (1 - fLengthCoeff)*fLabelWidth;
                    fDeltaY = fLabelHeight;
                }
                else{
                    fDeltaX = - fLengthCoeff*fLabelWidth;
                    fDeltaY = fLabelHeight;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 > 0 && oSolvation.x1 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY);
                    }
                    else if(oSolvation.x2 > 0 && oSolvation.x2 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY);
                    }
                }
                break;
            }
            case 2:{
                if(fBisectAngle < Math.PI){
                    fDeltaX = -fLabelWidth;
                    fDeltaY = (1 - fLengthCoeff)*fLabelHeight;
                }
                else{
                    fDeltaX = -fLabelWidth;
                    fDeltaY = - fLengthCoeff*fLabelHeight;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 > 0 && oSolvation.x1 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;

                    }
                    else if(oSolvation.x2 > 0 && oSolvation.x2 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;
                    }
                }
                break;
            }
            case 3:{
                fLengthCoeff = 1 - fLengthCoeff;
                if(fBisectAngle < 3*Math.PI/2){
                    fDeltaX = -fLabelWidth*fLengthCoeff;
                    fDeltaY = -fLabelHeight;
                }
                else{
                    fDeltaX = (1 - fLengthCoeff)*fLabelWidth;
                    fDeltaY = -fLabelHeight;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 > 0 && oSolvation.x1 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight;
                    }
                    else if(oSolvation.x2 > 0 && oSolvation.x2 < 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight;
                    }
                }
                break;
            }
        }
		return oResult;
	},

	_calculateBestFitPositionOuter: function(fStartAngle, fEndAngle, fPieRadius, fLabelWidth, fLabelHeight, fCenterX, fCenterY){
        var oResult = {bError: true};
        var fBisectAngle = AscFormat.normalizeRotate((fStartAngle + fEndAngle)/2.0);
        var fBisectAngle2 = AscFormat.normalizeRotate(fBisectAngle + Math.PI/4) - Math.PI/4;
        var nIndexArea = ((fBisectAngle2 + Math.PI/4)/(Math.PI/2)) >> 0;


        var fLengthCoeff =  ((fBisectAngle2 + Math.PI/4) - (Math.PI/2)*nIndexArea)/(Math.PI/2);

        var fXs = fCenterX + fPieRadius*Math.cos(fBisectAngle);
        var fYs = fCenterY + fPieRadius*Math.sin(fBisectAngle);
        var fDeltaX, fDeltaY, oSolvation;

        var fAngleApproxDelta = 1e-4;
        switch(nIndexArea){
            case 0:{
                if(AscFormat.fApproxEqual(fBisectAngle2, 0, fAngleApproxDelta)){
                    fDeltaX = 0;
                    fDeltaY = 0;
                }
                else if(fBisectAngle2 < 0){
                    fDeltaX = 0;
                    fDeltaY = fLengthCoeff*fLabelHeight;
                }
                else{
                    fDeltaX = 0;
                    fDeltaY = -(1 - fLengthCoeff)*fLabelHeight;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX);
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
                    }
                    else if(oSolvation.x2 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX);
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
                    }
                    else if(oSolvation.x1 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX);
						oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
					}
					else if(oSolvation.x2 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX);
						oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - (1 - fLengthCoeff)*fLabelHeight;
					}
                }
                break;
            }
            case 1:{
                if(AscFormat.fApproxEqual(fBisectAngle, Math.PI/2, fAngleApproxDelta)){
                    fDeltaX = 0;
                    fDeltaY = 0;
                }
                else if(fBisectAngle < Math.PI/2){
                    fDeltaX = -fLengthCoeff*fLabelWidth;
                    fDeltaY = 0;
                }
                else{
                    fDeltaX = (1 - fLengthCoeff)*fLabelWidth;
                    fDeltaY = 0;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY);
                    }
                    else if(oSolvation.x2 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY);
                    }
                    else if(oSolvation.x1 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
						oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY);
					}
					else if(oSolvation.x2 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth*fLengthCoeff;
						oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY);
					}
                }
                break;
            }
            case 2:{
                if(AscFormat.fApproxEqual(fBisectAngle, Math.PI, fAngleApproxDelta)){
                    fDeltaX = 0;
                    fDeltaY = 0;
                }
                else if(fBisectAngle < Math.PI){
                    fDeltaX = 0;
                    fDeltaY = -fLengthCoeff*fLabelHeight;
                }
                else{
                    fDeltaX = 0;
                    fDeltaY = (1 - fLengthCoeff)*fLabelHeight;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;

                    }
                    else if(oSolvation.x2 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;
                    }
                    else if(oSolvation.x1 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - fLabelWidth;
						oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;

					}
					else if(oSolvation.x2 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - fLabelWidth;
						oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight*fLengthCoeff;
					}
                }
                break;
            }
            case 3:{
                if(fBisectAngle < 3*Math.PI/2){
                    fDeltaX = fLabelWidth*fLengthCoeff;
                    fDeltaY = 0;
                }
                else{
                    fDeltaX = -(1 - fLengthCoeff)*fLabelWidth;
                    fDeltaY = 0;
                }
                oSolvation = AscFormat.fSolveQuadraticEquation(fPieRadius*fPieRadius, 2*(fDeltaX*(fXs - fCenterX) + fDeltaY*(fYs - fCenterY)), fDeltaX*fDeltaX + fDeltaY*fDeltaY - fPieRadius*fPieRadius);
                if(!oSolvation.bError){
                    if(oSolvation.x1 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - (1 - fLengthCoeff)*fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight;
                    }
                    else if(oSolvation.x2 >= 1){
                        oResult.bError = false;
                        oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - (1 - fLengthCoeff)*fLabelWidth;
                        oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight;
                    }
					else if(oSolvation.x1 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x1*(fXs - fCenterX) - (1 - fLengthCoeff)*fLabelWidth;
						oResult.fY = fCenterY + oSolvation.x1*(fYs - fCenterY) - fLabelHeight;
					}
					else if(oSolvation.x2 >= 0){
						oResult.bError = false;
						oResult.fX = fCenterX + oSolvation.x2*(fXs - fCenterX) - (1 - fLengthCoeff)*fLabelWidth;
						oResult.fY = fCenterY + oSolvation.x2*(fYs - fCenterY) - fLabelHeight;
					}
                }
                break;
            }
        }
        return oResult;
	},

    _calculateInEndDLblPosition: function(fStartAngle, fSweepAngle, fPieRadius, fLabelWidth, fLabelHeight, fCenterX, fCenterY){
        var fEndAngle = fStartAngle + fSweepAngle;
        var oResult = {bError: true, fX: 0.0, fY: 0.0};
        var fBisectAngle = AscFormat.normalizeRotate((fStartAngle + fEndAngle)/2);
        var nQuadrantIndex = (2.0*fBisectAngle/Math.PI) >> 0;
        var fHalfRectWidthVector = fLabelWidth/ 2, fHalfRectHeightVector = fLabelHeight/2;
        if(nQuadrantIndex === 1 || nQuadrantIndex == 2){
            fHalfRectWidthVector = -fHalfRectWidthVector;
        }
        if(nQuadrantIndex === 2 || nQuadrantIndex == 3){
            fHalfRectHeightVector = -fHalfRectHeightVector;
        }

        var fXs = fCenterX + fPieRadius*Math.cos(fBisectAngle), fYs = fCenterY + fPieRadius*Math.sin(fBisectAngle);
        var a = fPieRadius*fPieRadius, b = 2*( (fXs - fCenterX)*fHalfRectWidthVector + (fYs - fCenterY)*fHalfRectHeightVector), c = fHalfRectWidthVector*fHalfRectWidthVector + fHalfRectHeightVector*fHalfRectHeightVector - fPieRadius*fPieRadius;
        var oSolution = AscFormat.fSolveQuadraticEquation(a, b, c);
		if(oSolution.bError){
			return oResult;
		}
		var D = b*b - 4*a*c;
        if(D < 0){
            return oResult;
        }
        var t1 = oSolution.x1, t2 = oSolution.x2;
        if(t1 > 0 && t1 < 1){
            oResult.bError = false;
            oResult.fX = fCenterX + t1*(fXs - fCenterX) - fLabelWidth/2;
            oResult.fY = fCenterY + t1*(fYs - fCenterY) - fLabelHeight/2;
            return oResult
        }
        if(t2 > 0 && t2 < 1){
            oResult.bError = false;
            oResult.fX = fCenterX + t2*(fXs - fCenterX) - fLabelWidth/2;
            oResult.fY = fCenterY + t2*(fYs - fCenterY) - fLabelHeight/2;
            return oResult
        }
        return oResult;
    },

	//For test
	_calculateTestFrame: function (point1, point2, point3, point4, point5, point6, point7, point8) {
		var pxToMm = this.chartProp.pxToMM;
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		path.moveTo(point1.x / pxToMm * pathW, point1.y / pxToMm * pathH);
		path.lnTo(point2.x / pxToMm * pathW, point2.y / pxToMm * pathH);
		path.lnTo(point3.x / pxToMm * pathW, point3.y / pxToMm * pathH);
		path.lnTo(point4.x / pxToMm * pathW, point4.y / pxToMm * pathH);
		path.lnTo(point1.x / pxToMm * pathW, point1.y / pxToMm * pathH);

		path.moveTo(point5.x / pxToMm * pathW, point5.y / pxToMm * pathH);
		path.lnTo(point6.x / pxToMm * pathW, point6.y / pxToMm * pathH);
		path.lnTo(point7.x / pxToMm * pathW, point7.y / pxToMm * pathH);
		path.lnTo(point8.x / pxToMm * pathW, point8.y / pxToMm * pathH);
		path.lnTo(point5.x / pxToMm * pathW, point5.y / pxToMm * pathH);

		path.moveTo(point1.x / pxToMm * pathW, point1.y / pxToMm * pathH);
		path.lnTo(point5.x / pxToMm * pathW, point5.y / pxToMm * pathH);

		path.moveTo(point2.x / pxToMm * pathW, point2.y / pxToMm * pathH);
		path.lnTo(point6.x / pxToMm * pathW, point6.y / pxToMm * pathH);

		path.moveTo(point3.x / pxToMm * pathW, point3.y / pxToMm * pathH);
		path.lnTo(point7.x / pxToMm * pathW, point7.y / pxToMm * pathH);

		path.moveTo(point4.x / pxToMm * pathW, point4.y / pxToMm * pathH);
		path.lnTo(point8.x / pxToMm * pathW, point8.y / pxToMm * pathH);


		this.paths.test = pathId;
	},

	//TODO delete after test
	_recalculatePie3D_Slow: function () {
		var trueWidth = this.chartProp.trueWidth;
		var trueHeight = this.chartProp.trueHeight;

		var numCache = this._getFirstRealNumCache();
		if(!numCache) {
			return;
		}

		var sumData = this.cChartDrawer._getSumArray(numCache, true);

		var radius = Math.min(trueHeight, trueWidth) / 2;
		if (radius < 0) {
			radius = 0;
		}

		var xCenter = this.chartProp.chartGutter._left + trueWidth / 2;
		var yCenter = this.chartProp.chartGutter._top + trueHeight / 2;

		var startAngle = this.cChartDrawer.processor3D.angleOy ? this.cChartDrawer.processor3D.angleOy : 0;
		var startAngle3D = startAngle !== 0 && startAngle !== undefined ? this._changeAngle(radius, Math.PI / 2, startAngle, xCenter, yCenter, this.properties3d) : 0;

		this.tempAngle = Math.PI / 2 + startAngle;
		this.angleFor3D = Math.PI / 2 - startAngle3D;

		//рисуем против часовой стрелки, поэтому цикл с конца
		var depth = this.properties3d.depth;
		for (var n = 0; n < depth; n++) {
			if (!this.paths.series) {
				this.paths.series = [];
			}

			for (var i = numCache.length - 1; i >= 0; i--) {
				var angle = Math.abs((parseFloat(numCache[i].val / sumData)) * (Math.PI * 2));
				if (!this.paths.series[n]) {
					this.paths.series[n] = [];
				}

				if (sumData === 0)//TODO стоит пересмотреть
				{
					this.paths.series[n][i] = this._calculateEmptySegment(radius, xCenter, yCenter);
				} else {
					this.paths.series[n][i] = this._calculateSegment3D(angle, radius, xCenter, yCenter, n, i);
				}
			}
		}

	},

	_calculateSegment3D_Slow: function (angle, radius, xCenter, yCenter, depth, i) {
		if (isNaN(angle)) {
			return null;
		}

		var startAngle = (this.tempAngle);
		var path = this._calculateArc3D(radius, startAngle, angle, xCenter, yCenter, depth, i);
		this.tempAngle += angle;

		return path;
	},

	_calculateArc3D_Slow: function (radius, stAng, swAng, xCenter, yCenter, depth, seriaNum) {
		var radius1 = this.properties3d.radius1;
		var radius2 = this.properties3d.radius2;
		var pxToMm = this.chartProp.pxToMM;
		var t = this;

		var x0, y0, radiusSpec;
		var calculateProps = function () {
			if (t.usually3dPropsCalc && t.usually3dPropsCalc[seriaNum]) {
				swAng = t.usually3dPropsCalc[seriaNum].swAng;
				stAng = t.usually3dPropsCalc[seriaNum].stAng;
				radiusSpec = t.usually3dPropsCalc[seriaNum].radiusSpec;
				x0 = t.usually3dPropsCalc[seriaNum].x0;

				yCenter = yCenter + t.properties3d.depth / 2 - depth;
				y0 = (yCenter - radiusSpec * Math.sin(stAng));
			} else {
				swAng = t._changeAngle(radius, stAng, swAng, xCenter, yCenter, t.properties3d);
				stAng = t.angleFor3D;

				//корректируем центр
				yCenter = yCenter + t.properties3d.depth / 2 - depth;

				radiusSpec = (radius1 * radius2) / Math.sqrt(Math.pow(radius2, 2) * Math.pow((Math.cos(stAng)), 2) + Math.pow(radius1, 2) * Math.pow(Math.sin(stAng), 2));

				x0 = (xCenter + radiusSpec * Math.cos(stAng));
				y0 = (yCenter - radiusSpec * Math.sin(stAng));
			}
		};


		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		calculateProps();

		path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);
		path.lnTo(x0 / pxToMm * pathW, y0 / pxToMm * pathH);
		path.arcTo(radius1 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * stAng * cToDeg, -1 * swAng * cToDeg);
		path.lnTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);


		this.angleFor3D += swAng;
		if (!this.usually3dPropsCalc[seriaNum]) {
			this.usually3dPropsCalc[seriaNum] = {swAng: swAng, stAng: stAng, xCenter: xCenter, x0: x0, radiusSpec: radiusSpec};
		}

		return pathId;
	},

	_drawPie3D_Slow: function () {
		var numCache = this._getFirstRealNumCache();
		if(!numCache) {
			return;
		}

		var props = this.cChartSpace.getParentObjects();
		var brush, pen, val;
		var path;
		for (var n = 0; n < this.paths.series.length; n++) {
			for (var i = 0, len = numCache.length; i < len; i++) {
				val = numCache[i];
				brush = val.brush;

				if (n === 0 || n === this.paths.series.length - 1) {
					pen = val.pen;
				} else {
					pen = null;
				}

				path = this.paths.series[n][i];

				var duplicateBrush = brush;
				if (n !== this.paths.series.length - 1) {
					duplicateBrush = brush.createDuplicate();
					var cColorMod = new AscFormat.CColorMod("shade", 35000);

					duplicateBrush.addColorMod(cColorMod);
					duplicateBrush.calculate(props.theme, props.slide, props.layout, props.master, new AscFormat.CUniColor().RGBA, this.cChartSpace.clrMapOvr);
				}

				this.cChartDrawer.drawPath(path, pen, duplicateBrush);
			}
		}

	}
};


	/** @constructor */
function drawDoughnutChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;

	this.tempAngle = null;
	this.paths = {};
}

drawDoughnutChart.prototype = {
	constructor: drawDoughnutChart,

	draw: function () {
		this._drawPie();
	},

	recalculate: function () {
		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.seriesCount = countSeries.series;

		this.tempAngle = null;
		this.paths = {};
		this._reCalculatePie();
	},

	_drawPie: function () {
		var brush, pen;
		var path;
		var idxPoint, numCache;

		for (var n = 0; n < this.chart.series.length; n++) {
			numCache = this.cChartDrawer.getNumCache(this.chart.series[n].val);

			if (!numCache) {
				continue;
			}

			for (var k = 0; k < numCache.ptCount; k++) {
				
				if(!this.paths.series[n] || !this.paths.series[n][k]) {
					continue;
				}

				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[n], k);

				brush = idxPoint ? idxPoint.brush : null;
				pen = idxPoint ? idxPoint.pen : null;
				path = this.paths.series[n][k];

				this.cChartDrawer.drawPath(path, pen, brush);
			}
		}
	},

	_reCalculatePie: function () {
		var trueWidth = this.chartProp.trueWidth;
		var trueHeight = this.chartProp.trueHeight;

		var sumData;
		var outRadius = Math.min(trueHeight, trueWidth) / 2;

		//% from out radius
		var defaultSize = 50;
		var holeSize = this.chart.holeSize ? this.chart.holeSize : defaultSize;

		//first ang
		var firstSliceAng = this.chart.firstSliceAng ? this.chart.firstSliceAng : 0;
		firstSliceAng = (firstSliceAng / 360) * (Math.PI * 2);

		//inner radius
		var radius = outRadius * (holeSize / 100);
		var step = (outRadius - radius) / this.seriesCount;

		var xCenter = this.chartProp.chartGutter._left + trueWidth / 2;
		var yCenter = this.chartProp.chartGutter._top + trueHeight / 2;

		var numCache, idxPoint, angle, curVal, seriesCounter = 0;
		for (var n = 0; n < this.chart.series.length; n++) {
			this.tempAngle = Math.PI / 2;
			numCache = this.cChartDrawer.getNumCache(this.chart.series[n].val);
			if (!numCache || this.chart.series[n].isHidden) {
				continue;
			}

			sumData = this.cChartDrawer._getSumArray(numCache.pts, true);

			//рисуем против часовой стрелки, поэтому цикл с конца
			for (var k = numCache.ptCount - 1; k >= 0; k--) {

				idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[n], k);
				curVal = idxPoint ? idxPoint.val : 0;
				angle = Math.abs((parseFloat(curVal / sumData)) * (Math.PI * 2));

				//правка связана с реализацией arcTo, где swAng зануляется и приравнивается к значению
				if(angle < 10e-16) {
					angle = 0;
				}

				if (!this.paths.series) {
					this.paths.series = [];
				}
				if (!this.paths.series[n]) {
					this.paths.series[n] = [];
				}

				if (angle) {
					this.paths.series[n][k] =
						this._calculateSegment(angle, radius, xCenter, yCenter, radius + step * (seriesCounter + 1), radius + step * seriesCounter, firstSliceAng);
				} else {
					this.paths.series[n][k] = null;
				}
			}

			if (numCache.pts.length) {
				seriesCounter++;
			}

		}
	},

	_calculateSegment: function (angle, radius, xCenter, yCenter, radius1, radius2, firstSliceAng) {
		var startAngle = this.tempAngle - firstSliceAng;
		var endAngle = angle;

		if (radius < 0) {
			radius = 0;
		}
		var path = this._calculateArc(radius, startAngle, endAngle, xCenter, yCenter, radius1, radius2);

		this.tempAngle += angle;

		return path;
	},

	_calculateArc: function (radius, stAng, swAng, xCenter, yCenter, radius1, radius2) {

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		var pxToMm = this.chartProp.pxToMM;

		var x2 = xCenter + radius1 * Math.cos(stAng);
		var y2 = yCenter - radius1 * Math.sin(stAng);

		var x1 = xCenter + radius2 * Math.cos(stAng);
		var y1 = yCenter - radius2 * Math.sin(stAng);

		//var x3 = xCenter + radius1*Math.cos(stAng + swAng);
		//var y3 = yCenter - radius1*Math.sin(stAng + swAng);

		var x4 = xCenter + radius2 * Math.cos(stAng + swAng);
		var y4 = yCenter - radius2 * Math.sin(stAng + swAng);

		path.moveTo(x1 / pxToMm * pathW, y1 / pxToMm * pathH);
		path.lnTo(x2 / pxToMm * pathW, y2 / pxToMm * pathH);
		path.arcTo(radius1 / pxToMm * pathW, radius1 / pxToMm * pathH, -1 * stAng * cToDeg, -1 * swAng * cToDeg);
		path.lnTo(x4 / pxToMm * pathW, y4 / pxToMm * pathH);
		path.arcTo(radius2 / pxToMm * pathW, radius2 / pxToMm * pathH, -1 * stAng * cToDeg - swAng * cToDeg, swAng * cToDeg);
		path.moveTo(xCenter / pxToMm * pathW, yCenter / pxToMm * pathH);


		return pathId;
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		if (!this.paths || !this.paths.series || !this.paths.series[ser] || !AscFormat.isRealNumber(this.paths.series[ser][val])) {
			return;
		}

		var path = this.paths.series[ser][val];
		var oPath = this.cChartSpace.GetPath(path);

		var oCommand2 = oPath.getCommandByIndex(2);
		var oCommand4 = oPath.getCommandByIndex(4);
		var oCommand5 = oPath.getCommandByIndex(5);

		var radius1 = oCommand2.hR;
		var stAng = oCommand2.stAng;
		var swAng = oCommand2.swAng;

		var radius2 = oCommand4.hR;
		var xCenter = oCommand5.X;
		var yCenter = oCommand5.Y;


		var newRadius = radius2 + (radius1 - radius2) / 2;
		var centerX = xCenter + newRadius * Math.cos(-1 * stAng - swAng / 2);
		var centerY = yCenter - newRadius * Math.sin(-1 * stAng - swAng / 2);

		var numCache = this.cChartDrawer.getNumCache(this.chart.series[ser].val);
		var point = null;
		if(numCache){
			point = numCache.getPtByIndex(val);
		}
		if (!point) {
			return;
		}

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.ctr: {
				centerX = centerX - width / 2;
				centerY = centerY - height / 2;
				break;
			}
			case c_oAscChartDataLabelsPos.inBase: {
				centerX = centerX - width / 2;
				centerY = centerY - height / 2;
				break;
			}
		}
		if (centerX < 0) {
			centerX = 0;
		}
		if (centerY < 0) {
			centerY = 0;
		}

		return {x: centerX, y: centerY};
	}
};


/** @constructor */
function drawRadarChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.radarStyle = null;
	this.valAx = null;
	this.catAx = null;

	this._minMaxValue = null;

	this.paths = {};
	this.fillPaths = [];
}

drawRadarChart.prototype = {
	constructor: drawRadarChart,

	draw: function () {
		this._drawLines();
	},

	recalculate: function () {
		this.paths = {};
		this.radarStyle = this.chart.radarStyle;
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);

		this._calculateLines();
		if (this.radarStyle === AscFormat.RADAR_STYLE_FILLED) {
			this._calculateLines(true);
		}
	},

	_calculateLines: function (fillPath) {
		let dispBlanksAs = this.cChartSpace.chart.dispBlanksAs;
		if (this.radarStyle === AscFormat.RADAR_STYLE_FILLED && dispBlanksAs === AscFormat.DISP_BLANKS_AS_GAP) {
			dispBlanksAs = AscFormat.DISP_BLANKS_AS_ZERO;
		}
		let yPoints = this.valAx.yPoints;

		let xCenter = this.valAx.posX;
		let yCenter = !this.valAx.isReversed() ? yPoints[0].pos : yPoints[yPoints.length - 1].pos;

		let valueMinMax;
		if (yPoints && yPoints.length > 1) {
			valueMinMax = this._getMinMaxValue();
			this._minMaxValue = valueMinMax;
		}

		//TODO
		/*if (this.valAx.scaling.logBase && (val < 0)) {
			break;
		}*/

		let isLastPoint = function (_number, _oNumCache) {
			if (dispBlanksAs === AscFormat.DISP_BLANKS_AS_GAP) {
				return _number === _oNumCache.ptCount - 1;
			} else {
				for (let k = _number + 1; k < _oNumCache.ptCount; k++) {
					let _pt = _oNumCache.getPtByIndex(k);
					if (_pt) {
						return false;
					}
				}
				return true;
			}
		};

		let y, y1, x, x1, curSeries, dataSeries;
		let radius1, radius2, xFirst, yFirst, calcPath, points;
		for (let i = 0; i < this.chart.series.length; i++) {

			curSeries = this.chart.series[i];
			let oNumCache = this.cChartDrawer.getNumCache(curSeries.val);
			if (!oNumCache) {
				continue;
			}
			dataSeries = oNumCache.pts;
			if (!dataSeries) {
				continue;
			}

			if (fillPath) {
				calcPath = this._calculateFillPath();
			}

			let pt, nextPt, alpha1, alpha2;
			let isOnePoint = oNumCache.ptCount === 1;
			for (let n = 0; n < oNumCache.ptCount; n++) {
				//first point
				pt = oNumCache.getPtByIndex(n);
				x = y = null;
				if (!pt && dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO) {
					pt = {val: 0, idx: n};
				}
				if (pt) {
					radius1 = this._getRadius(pt.val, valueMinMax);
					alpha1 = this._getAlpha(pt.idx);

					y = yCenter - radius1 * Math.cos(alpha1);
					x = xCenter + radius1 * Math.sin(alpha1);

					if (n === 0) {
						xFirst = x;
						yFirst = y;
					}
				}

				//second point
				if (!isOnePoint) {
					nextPt = this.cChartDrawer.getNextCachePoint(oNumCache, n, dispBlanksAs);
					if (nextPt) {
						radius2 = this._getRadius(nextPt.val, valueMinMax);
						alpha2 = this._getAlpha(nextPt.idx);
					}
				}

				x1 = y1 = null;
				if (nextPt) {
					y1 = !isOnePoint ? yCenter - radius2 * Math.cos(alpha2) : null;
					x1 = !isOnePoint ? xCenter + radius2 * Math.sin(alpha2) : null;
				}

				if (isOnePoint) {
					this._addPointToPaths(x, y, pt, i, n);
				} else if (fillPath) {
					//AscFormat.RADAR_STYLE_MARKER AscFormat.RADAR_STYLE_FILLED  AscFormat.RADAR_STYLE_STANDARD
					if (x1 !== null && x !== null && y !== null && y1 !== null) {
						points = {x: x, y: y, x1: x1, y1: y1};
						calcPath(pt, n, points, oNumCache);
					}
					if (isLastPoint(n + 1, oNumCache) && xFirst != null && x1 !== null) {
						points = {x: x1, y: y1, x1: xFirst, y1: yFirst};
						calcPath(pt || nextPt, n + 1, points, oNumCache, true);
						break;
					}
				} else {
					//1. draw first point marker
					if (n === 0 && x !== null) {
						this._addPointToPaths(x, y, pt, i, n);
					}
					//2. draw next point marker
					if (x1 !== null) {
						this._addPointToPaths(x1, y1, nextPt, i, n + 1);
					}
					//3. draw main line
					if (x1 !== null && x !== null) {
						this._addLineToTemporary(x, y, x1, y1, i, n);
					}
					//4. draw last line(return by first point)
					if ((isLastPoint(n + 1, oNumCache)) && xFirst != null && x1 !== null) {
						this._addLineToTemporary(x1, y1, xFirst, yFirst, i, n + 1);
						break;
					}
				}
			}
		}
		this._addLineToPathsFromTemporary();
	},

	_addPointToPaths: function (x, y, pt, indexSer, indexPt) {
		if (!this.paths.points) {
			this.paths.points = [];
		}
		if (!this.paths.points[indexSer]) {
			this.paths.points[indexSer] = [];
		}

		let size = pt && pt.compiledMarker && pt.compiledMarker.size;
		let symbol = pt && pt.compiledMarker && pt.compiledMarker.symbol;
		this.paths.points[indexSer][indexPt] = this.cChartDrawer.calculatePoint(x, y, size, symbol);
	},

	_addLineToTemporary: function (x1, y1, x2, y2, indexSer, indexPt) {
		//made for unification lines with identical settings
		if (!this._tempLines) {
			this._tempLines = [];
		}
		if (!this._tempLines[indexSer]) {
			this._tempLines[indexSer] = [];
		}

		this._tempLines[indexSer][indexPt] = {x1: x1, y1: y1, x2: x2, y2: y2};
	},

	_addLineToPathsFromTemporary: function () {
		let t = this;
		let path, pathId;
		let generateNewPath = function () {
			pathId = t.cChartSpace.AllocPath();
			path = t.cChartSpace.GetPath(pathId);
		};

		let dispBlanksAs = this.cChartSpace.chart.dispBlanksAs;
		if (this._tempLines) {
			if (!this.paths.series) {
				this.paths.series = [];
			}
			for (let ser = 0; ser < this._tempLines.length; ser++) {
				if (!this.paths.series[ser]) {
					this.paths.series[ser] = [];
				}
				let uninterruptedLine = true;
				if (this._tempLines[ser]) {
					for (let point = 0, length = this._tempLines[ser].length; point < length; point++) {
						let curTempSer = this._tempLines[ser];
						let containerForDrawSer = this.paths.series[ser];
						if (curTempSer && curTempSer[point]) {
							if (!containerForDrawSer.length) {
								//add new start path
								generateNewPath();
								this.cChartDrawer.calculateLine(curTempSer[point].x1, curTempSer[point].y1, curTempSer[point].x2, curTempSer[point].y2, path);
								containerForDrawSer[point] = pathId;
							} else {
								if (this._comparePointsSettings(ser, point - 1, point, dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO)) {
									//use previous path
									if (uninterruptedLine && point === length - 1 && path && curTempSer[0] &&
										curTempSer[point].x2 === curTempSer[0].x1 && curTempSer[point].y2 === curTempSer[0].y1) {
										path.close();
									} else {
										this.cChartDrawer.calculateLine(curTempSer[point].x1, curTempSer[point].y1, curTempSer[point].x2, curTempSer[point].y2, path, true);
									}
								} else {
									generateNewPath();
									this.cChartDrawer.calculateLine(curTempSer[point].x1, curTempSer[point].y1, curTempSer[point].x2, curTempSer[point].y2, path);
									containerForDrawSer[point] = pathId;
								}
							}
						} else {
							uninterruptedLine = false;
						}
					}
				}
			}
		}
	},

	_comparePointsSettings: function (serIndex, pt1Index, pt2index, useNullPoints) {
		let ser = this.chart.series[serIndex];
		let oNumCache = ser && this.cChartDrawer.getNumCache(ser.val);
		if (oNumCache) {
			let pt1 = oNumCache.getPtByIndex(pt1Index);
			let pt2 = oNumCache.getPtByIndex(pt2index);
			if ((useNullPoints && (!pt1 || !pt2)) || (pt1 && pt2 && pt1.pen && pt2.pen && pt1.pen.IsIdentical(pt2.pen))) {
				return true;
			}
		}
		return false;
	},

	_addLineToPaths: function (x1, y1, x2, y2, indexSer, indexPt) {
		if (!this.paths.series) {
			this.paths.series = [];
		}
		if (!this.paths.series[indexSer]) {
			this.paths.series[indexSer] = [];
		}

		this.paths.series[indexSer][indexPt] = this.cChartDrawer.calculateLine(x1, y1, x2, y2);
	},

	_getAlpha: function (numPoint) {
		let res = null;
		if (this.catAx && this.catAx.alphaPoints) {
			let oPt = this.catAx.alphaPoints[numPoint];
			if (oPt) {
				return oPt.pos;
			}
		}
		return res;
	},

	_calculateFillPath: function () {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;
		var t = this;
		return function (pt, n, points, oNumCache, endOfPath) {
			var pen, brush;
			var y = points.y, y1 = points.y1, x = points.x, x1 = points.x1;

			if (n === 0) {
				path.moveTo(x * pathW, y * pathH);
				path.lnTo(x1 * pathW, y1 * pathH);
			} else {
				path.lnTo(x * pathW, y * pathH);
				path.lnTo(x1 * pathW, y1 * pathH);
			}

			if ((oNumCache.ptCount - 2 === n || endOfPath) && pt) {
				pen = pt.pen;
				brush = pt.brush;
				t.fillPaths.push({path: pathId, pen: pen, brush: brush});
			}
			return pathId;
		};
	},

	_getMinMaxValue: function () {
		var temp = 0, resMax = 0, resMin = 0, numCache;
		for (var i = 0; i < this.chart.series.length; i++) {
			numCache = this.cChartDrawer.getNumCache(this.chart.series[i].val);
			if (!numCache) {
				continue;
			}

			for (var n = 0; n < numCache.pts.length; n++) {
				if (!numCache.pts[n]) {
					continue;
				}
				temp = parseFloat(numCache.pts[n].val);

				resMin = temp < resMin ? temp : resMin;
				resMax = temp > resMax ? temp : resMax;
			}
		}
		return {min: resMin, max: resMax};
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var numCache = this.cChartDrawer.getNumCache(this.chart.series[ser].val);
		if (!numCache) {
			return;
		}

		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);

		var oCommand, oPath;
		if (this.paths.points) {
			if (this.paths.points[ser] && this.paths.points[ser][val]) {
				oPath = this.cChartSpace.GetPath(this.paths.points[ser][val].path);
				oCommand = oPath.getCommandByIndex(0);
			}
		}

		if (!oCommand) {
			return;
		}

		var x = oCommand.X;
		var y = oCommand.Y;

		var pxToMm = this.chartProp.pxToMM;
		var constMargin = 5 / pxToMm;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x - width / 2;
		var centerY = y - height / 2;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.outEnd: {

				let yPoints = this.valAx.yPoints;
				if (yPoints && yPoints.length > 1) {
					let xCenter = this.valAx.posX;
					let yCenter = !this.valAx.isReversed() ? yPoints[0].pos : yPoints[yPoints.length - 1].pos;

					let radius1 = this._getRadius(point.val, this._minMaxValue);
					let alpha1 = this._getAlpha(point.idx);

					//TODO margin differ from ms
					centerY = yCenter - (radius1 + constMargin) * Math.cos(alpha1) - height / 2;
					centerX = xCenter + (radius1 + constMargin) * Math.sin(alpha1) - width / 2;
				}

				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_getRadius: function (val) {
		var yPoints = this.valAx.yPoints;
		var orientation = !this.valAx.isReversed();
		return (orientation ? yPoints[0].pos : yPoints[yPoints.length - 1].pos) - this.cChartDrawer.getYPosition(val, this.valAx);
	},

	_drawLines: function () {
		//this.cShapeDrawer.Graphics.SaveGrState();
		//this.cShapeDrawer.Graphics.AddClipRect(this.chartProp.chartGutter._left / this.chartProp.pxToMM, this.chartProp.chartGutter._top / this.chartProp.pxToMM, this.chartProp.trueWidth / this.chartProp.pxToMM, this.chartProp.trueHeight / this.chartProp.pxToMM);
		if (this.radarStyle === AscFormat.RADAR_STYLE_FILLED) {
			for (var i = 0; i < this.fillPaths.length; i++) {
				this.cChartDrawer.drawPath(this.fillPaths[i].path, this.fillPaths[i].pen, this.fillPaths[i].brush);
			}
		} else {
			this.cChartDrawer.drawPaths(this.paths, this.chart.series, true, null, null, true);
			this.cChartDrawer.drawPathsPoints(this.paths, this.chart.series);
		}
	}
};


/** @constructor */
function drawScatterChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;

	this.paths = {};
}

drawScatterChart.prototype = {
	constructor: drawScatterChart,

	recalculate: function () {
		this.paths = {};

		this.catAx = this.chart.axId[0].xPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.catAx = this.cChartDrawer._searchChangedAxis(this.catAx);
		this.valAx = this.chart.axId[0].yPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.valAx = this.cChartDrawer._searchChangedAxis(this.valAx);

		this._recalculateScatter();
	},

	draw: function () {
		this._drawScatter();
	},

	_recalculateScatter: function () {
		let seria, yVal, xVal, points, yNumCache, compiledMarkerSize, compiledMarkerSymbol, yPoint, idx, xPoint;
		let dispBlanksAs =  this.cChartSpace.chart.dispBlanksAs;
		let isLog;

		let t = this;
		let _initObjs = function (_index) {
			if (!t.paths.points) {
				t.paths.points = [];
			}
			if (!t.paths.points[_index]) {
				t.paths.points[_index] = [];
			}
			if (!points) {
				points = [];
			}
			if (!points[_index]) {
				points[_index] = [];
			}
		};

		for (let i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			yNumCache = this.cChartDrawer.getNumCache(seria.yVal);
			isLog = this.valAx && this.valAx.scaling ? this.valAx.scaling.logBase : false;

			if (!yNumCache || !yNumCache.pts || yNumCache.pts.length === 0) {
				continue;
			}

			compiledMarkerSize = seria && seria.compiledSeriesMarker ? seria.compiledSeriesMarker.size : null;
			compiledMarkerSymbol = seria && seria.compiledSeriesMarker ? seria.compiledSeriesMarker.symbol : null;

			//idx не всегда начинается с нуля
			for (let n = 0; n < yNumCache.ptCount; n++) {
				idx = yNumCache.pts && undefined !== yNumCache.pts[n] ? yNumCache.pts[n].idx : null;

				if (dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO || dispBlanksAs === AscFormat.DISP_BLANKS_AS_GAP) {
					if (yNumCache.pts[n-1] && yNumCache.pts[n] && yNumCache.pts[n].idx - yNumCache.pts[n-1].idx > 1) {
						for (let k = 0; k < yNumCache.pts[n].idx - yNumCache.pts[n-1].idx - 1; k++) {
							_initObjs(i);
							this.paths.points[i].push(null);
							points[i].push(dispBlanksAs === AscFormat.DISP_BLANKS_AS_ZERO ? {x: 0, y: 0} : null);
						}
					}
				}

				let values = this.cChartDrawer._getScatterPointVal(seria, idx);
				if(values){
					yVal = values.y;
					xVal = values.x;
					xPoint = values.xPoint;
					yPoint = values.yPoint;

					if(yPoint && yPoint.compiledMarker) {
						compiledMarkerSize = yPoint.compiledMarker.size;
					}
					if(yPoint && yPoint.compiledMarker) {
						compiledMarkerSymbol = yPoint.compiledMarker.symbol;
					}

					_initObjs(i);

					if (yVal != null && ((isLog && yVal !== 0) || !isLog)) {
						let x = this.cChartDrawer.getYPosition(xVal, this.catAx, true);
						let y = this.cChartDrawer.getYPosition(yVal, this.valAx, true);
						this.paths.points[i].push(this.cChartDrawer.calculatePoint(x, y, compiledMarkerSize, compiledMarkerSymbol));

						let errBars = this.chart.series[i].errBars[0];
						if (errBars) {
							this.cChartDrawer.errBars.putPoint(x, y, xVal, yVal, seria.idx, idx);
						}

						points[i].push({x: xVal, y: yVal});
					} else {
						this.paths.points[i].push(null);
						points[i].push(null);
					}
				}

				//idx - индекс точки по оси OY
				/*idx = yNumCache.pts && undefined !== yNumCache.pts[n] ? yNumCache.pts[n].idx : null;
				if(null === idx) {
					continue;
				}

				//вычисляем yVal
				//пытаемся вычислить xVal  в зависимости от idx точки по OY
				yVal = this._getYVal(n, i);

				xPoint = this.cChartDrawer.getIdxPoint(seria, idx, true);
				if(undefined === xPoint) {
					continue;
				}

				if (xPoint) {
					xVal = xPoint.val;
					if (!isNaN(parseFloat(xVal))) {
						xVal = parseFloat(xVal);
					} else {
						xVal = n + 1;
					}
				} else {
					//xVal = betweenAxisCross ? n : n + 1;
					//xVal = this.catAx instanceof AscFormat.CCatAx ? n : n + 1;
					xVal = n + 1;
				}

				yPoint = this.cChartDrawer.getIdxPoint(seria, idx);
				compiledMarkerSize = yPoint && yPoint.compiledMarker ? yPoint.compiledMarker.size : null;
				compiledMarkerSymbol = yPoint && yPoint.compiledMarker ? yPoint.compiledMarker.symbol : null;


				if (!this.paths.points) {
					this.paths.points = [];
				}
				if (!this.paths.points[i]) {
					this.paths.points[i] = [];
				}

				if (!points) {
					points = [];
				}
				if (!points[i]) {
					points[i] = [];
				}

				if (yVal != null) {
					this.paths.points[i][n] = this.cChartDrawer.calculatePoint(this.cChartDrawer.getYPosition(xVal, this.catAx), this.cChartDrawer.getYPosition(yVal, this.valAx, true), compiledMarkerSize, compiledMarkerSymbol);
					points[i][n] = {x: xVal, y: yVal};
				} else {
					this.paths.points[i][n] = null;
					points[i][n] = null;
				}*/
			}
		}

		this._calculateAllLines(points);
	},

	_recalculateScatter2: function () {
		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;
		var betweenAxisCross = this.valAx.crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN;

		var seria, yVal, xVal, points, yNumCache, compiledMarkerSize, compiledMarkerSymbol, yPoint, idx, xPoint;
		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			yNumCache = this.cChartDrawer.getNumCache(seria.yVal);

			if (!yNumCache) {
				continue;
			}

			for (var n = 0; n < yNumCache.ptCount; n++) {
				//idx - индекс точки по оси OY
				idx = yNumCache.pts && undefined !== yNumCache.pts[n] ? yNumCache.pts[n].idx : null;
				if(null === idx) {
					continue;
				}

				//вычисляем yVal
				//пытаемся вычислить xVal  в зависимости от idx точки по OY
				yVal = this._getYVal(n, i);
				xPoint = this.cChartDrawer.getIdxPoint(seria, idx, true);
				if (xPoint) {
					xVal = xPoint.val;
					if (!isNaN(parseFloat(xVal))) {
						xVal = parseFloat(xVal);
					} else {
						xVal = n + 1;
					}
				} else {
					//xVal = betweenAxisCross ? n : n + 1;
					//xVal = this.catAx instanceof AscFormat.CCatAx ? n : n + 1;
					xVal = n + 1;
				}


				yPoint = this.cChartDrawer.getIdxPoint(seria, idx);
				compiledMarkerSize = yPoint && yPoint.compiledMarker ? yPoint.compiledMarker.size : null;
				compiledMarkerSymbol = yPoint && yPoint.compiledMarker ? yPoint.compiledMarker.symbol : null;


				if (!this.paths.points) {
					this.paths.points = [];
				}
				if (!this.paths.points[i]) {
					this.paths.points[i] = [];
				}

				if (!points) {
					points = [];
				}
				if (!points[i]) {
					points[i] = [];
				}

				if (yVal != null) {
					this.paths.points[i][n] = this.cChartDrawer.calculatePoint(this.cChartDrawer.getYPosition(xVal, this.catAx, true), this.cChartDrawer.getYPosition(yVal, this.valAx), compiledMarkerSize, compiledMarkerSymbol);
					points[i][n] = {x: xVal, y: yVal};
				} else {
					this.paths.points[i][n] = null;
					points[i][n] = null;
				}
			}
		}

		this._calculateAllLines(points);
	},

	_calculateAllLines: function (points) {
		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;

		//пока smooth для логарифмической шкалы не выставлю - рисуется не верно
		var allAxisLog = this.catAx.scaling && this.catAx.scaling.logBase && this.valAx.scaling && this.valAx.scaling.logBase;
		var x, y, x1, y1, x2, y2, x3, y3, isSplineLine;

		if(!points) {
			return;
		}

		for (var i = 0; i < points.length; i++) {
			isSplineLine = !allAxisLog && this.chart.series[i].smooth !== false;

			if (!points[i]) {
				continue;
			}

			for (var n = 0; n < points[i].length; n++) {
				if (!this.paths.series) {
					this.paths.series = [];
				}
				if (!this.paths.series[i]) {
					this.paths.series[i] = [];
				}

				if (points[i][n] != null && points[i][n + 1] != null) {
					if (isSplineLine) {

						x = points[i][n - 1] ? points[i][n - 1].x : points[i][n].x;
						y = points[i][n - 1] ? points[i][n - 1].y : points[i][n].y;

						x1 = points[i][n].x;
						y1 = points[i][n].y;

						x2 = points[i][n + 1] ? points[i][n + 1].x : points[i][n].x;
						y2 = points[i][n + 1] ? points[i][n + 1].y : points[i][n].y;

						x3 = points[i][n + 2] ? points[i][n + 2].x : points[i][n + 1] ? points[i][n + 1].x : points[i][n].x;
						y3 = points[i][n + 2] ? points[i][n + 2].y : points[i][n + 1] ? points[i][n + 1].y : points[i][n].y;

						//this.paths.series[i][n] = {path: this.cChartDrawer.calculateSplineLine(x, y, x1, y1, x2, y2, x3, y3, this.catAx, this.valAx), idx: points[i][n].idx};
						this.paths.series[i][n] = this.cChartDrawer.calculateSplineLine(x, y, x1, y1, x2, y2, x3, y3, this.catAx, this.valAx);
					} else {
						x = this.cChartDrawer.getYPosition(points[i][n].x, this.catAx, true);
						y = this.cChartDrawer.getYPosition(points[i][n].y, this.valAx, true);

						x1 = this.cChartDrawer.getYPosition(points[i][n + 1].x, this.catAx, true);
						y1 = this.cChartDrawer.getYPosition(points[i][n + 1].y, this.valAx, true);

						//this.paths.series[i][n] = {path: this._calculateLine(x, y, x1, y1), idx: points[i][n].idx};
						this.paths.series[i][n] = this._calculateLine(x, y, x1, y1);
					}
				}
			}
		}
	},

	_getYVal: function (n, i) {
		var idxPoint = this.cChartDrawer.getPointByIndex(this.chart.series[i], n);
		return idxPoint ? parseFloat(idxPoint.val) : null;
	},

	_drawScatter: function () {
		//TODO 2 раза проходимся по сериям!
		//add clip rect
		var diffPen = 2;
		var leftRect = this.chartProp.chartGutter._left / this.chartProp.pxToMM;
		var topRect = (this.chartProp.chartGutter._top - diffPen) / this.chartProp.pxToMM;
		var rightRect = this.chartProp.trueWidth / this.chartProp.pxToMM;
		var bottomRect = (this.chartProp.trueHeight + diffPen) / this.chartProp.pxToMM;

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);
		//draw lines
		//this.cChartDrawer.drawPathsByIdx(this.paths, this.chart.series, true, true);
		this.cChartDrawer.drawPaths(this.paths, this.chart.series, true, true);
		//end clip rect
		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();

		this.cChartDrawer.drawPathsPoints(this.paths, this.chart.series, true);
	},

	_calculateLine: function (x, y, x1, y1) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		path.moveTo(x * pathH, y * pathW);
		path.lnTo(x1 * pathH, y1 * pathW);

		return pathId;
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var point = this.cChartDrawer.getIdxPoint(this.chart.series[ser], val);
		var path;

		if(!point) {
			return;
		}

		if (this.paths && this.paths.points) {
			if (this.paths.points[ser] && this.paths.points[ser][val]) {
				var oPath = this.cChartSpace.GetPath(this.paths.points[ser][val].path);
				path = oPath.getCommandByIndex(0);
			}
		}

		if (!path) {
			return;
		}

		var x = path.X;
		var y = path.Y;

		var pxToMm = this.chartProp.pxToMM;
		var constMargin = 5 / pxToMm;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x - width / 2;
		var centerY = y - height / 2;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	//TODO пока включаю функцию _calculateSplineLine. с _calculateSplineLine2 отрисовается неверно. проверить!
	_calculateSplineLine2: function (points, k, xPoints, yPoints) {

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var x = points[k - 1] ? points[k - 1].x : points[k].x;
		var y = points[k - 1] ? points[k - 1].y : points[k].y;

		var x1 = points[k].x;
		var y1 = points[k].y;

		var x2 = points[k + 1] ? points[k + 1].x : points[k].x;
		var y2 = points[k + 1] ? points[k + 1].y : points[k].y;

		var x3 = points[k + 2] ? points[k + 2].x : points[k + 1] ? points[k + 1].x : points[k].x;
		var y3 = points[k + 2] ? points[k + 2].y : points[k + 1] ? points[k + 1].y : points[k].y;

		var splineCoords = this.cChartDrawer.calculate_Bezier(x, y, x1, y1, x2, y2, x3, y3);

		x = this.cChartDrawer.getYPosition(splineCoords[0].x, this.catAx);
		y = this.cChartDrawer.getYPosition(splineCoords[0].y, this.valAx);

		x1 = this.cChartDrawer.getYPosition(splineCoords[1].x, this.catAx);
		y1 = this.cChartDrawer.getYPosition(splineCoords[1].y, this.valAx);

		x2 = this.cChartDrawer.getYPosition(splineCoords[2].x, this.catAx);
		y2 = this.cChartDrawer.getYPosition(splineCoords[2].y, this.valAx);

		x3 = this.cChartDrawer.getYPosition(splineCoords[3].x, this.catAx);
		y3 = this.cChartDrawer.getYPosition(splineCoords[3].y, this.valAx);

		path.moveTo(x * pathW, y * pathH);
		path.cubicBezTo(x1 * pathW, y1 * pathH, x2 * pathW, y2 * pathH, x3 * pathW, y3 * pathH);


		return pathId;
	}
};


/** @constructor */
function drawStockChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;

	this.paths = {};
}

drawStockChart.prototype = {
	constructor: drawStockChart,

	draw: function () {
		this._drawLines();
	},

	recalculate: function () {
		this.paths = {};

		this.catAx = this.chart.axId[0].xPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.catAx = this.cChartDrawer._searchChangedAxis(this.catAx);
		this.valAx = this.chart.axId[0].yPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.valAx = this.cChartDrawer._searchChangedAxis(this.valAx);

		this._calculateLines();
	},

	_calculateLines: function () {
		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;

		var trueWidth = this.chartProp.trueWidth;

		var numCache = this.cChartDrawer.getNumCache(this.chart.series[0].val);
		if(!numCache) {
			return;
		}

		var koffX = trueWidth / numCache.pts.length;

		var gapWidth = this.chart.upDownBars && AscFormat.isRealNumber(this.chart.upDownBars.gapWidth) ? this.chart.upDownBars.gapWidth : 150;

		var widthBar = koffX / (1 + gapWidth / 100);

		var val1, val2, val3, val4, xVal, yVal1, yVal2, yVal3, yVal4, curNumCache, lastNamCache;
		for (var i = 0; i < numCache.pts.length; i++) {

			val1 = null, val2 = null, val3 = null, val4 = null;
			val1 = numCache.pts[i].val;

			lastNamCache = this.cChartDrawer.getNumCache(this.chart.series[this.chart.series.length - 1].val);
			lastNamCache = lastNamCache ? lastNamCache.pts : null;
			val4 = lastNamCache && lastNamCache[i] ? lastNamCache[i].val : null;

			for (var k = 1; k < this.chart.series.length - 1; k++) {
				curNumCache = this.cChartDrawer.getNumCache(this.chart.series[k].val);
				if (curNumCache && curNumCache.pts[i]) {
					if (k === 1) {
						val2 = curNumCache.pts[i].val;
						val3 = curNumCache.pts[i].val;
					} else {
						if (parseFloat(val2) > parseFloat(curNumCache.pts[i].val)) {
							val2 = curNumCache.pts[i].val;
						}
						if (parseFloat(val3) < parseFloat(curNumCache.pts[i].val)) {
							val3 = curNumCache.pts[i].val;
						}
					}
				}
			}

			if (!this.paths.values) {
				this.paths.values = [];
			}
			if (!this.paths.values[i]) {
				this.paths.values[i] = {};
			}

			xVal = this.cChartDrawer.getYPosition(i + 1, this.catAx);
			yVal1 = this.cChartDrawer.getYPosition(val1, this.valAx);
			yVal2 = this.cChartDrawer.getYPosition(val2, this.valAx);
			yVal3 = this.cChartDrawer.getYPosition(val3, this.valAx);
			yVal4 = this.cChartDrawer.getYPosition(val4, this.valAx);

			if (val2 !== null && val1 !== null) {
				this.paths.values[i].lowLines = this._calculateLine(xVal, yVal2, xVal, yVal1);
			}
			if (val3 !== null && val4 !== null) {
				this.paths.values[i].highLines = this._calculateLine(xVal, yVal4, xVal, yVal3);
			}

			if (val1 !== null && val4 !== null) {
				if (parseFloat(val1) > parseFloat(val4)) {
					this.paths.values[i].downBars = this._calculateUpDownBars(xVal, yVal1, xVal, yVal4, widthBar / this.chartProp.pxToMM);
				} else {
					this.paths.values[i].upBars = this._calculateUpDownBars(xVal, yVal1, xVal, yVal4, widthBar / this.chartProp.pxToMM);
				}
			}
		}
	},

	_drawLines: function () {
		var brush;
		var pen;
		var numCache = this.cChartDrawer.getNumCache(this.chart.series[0].val);
		if(!numCache) {
			return;
		}

		for (var i = 0; i < numCache.pts.length; i++) {

			pen = this.chart.calculatedHiLowLines;

			this.cChartDrawer.drawPath(this.paths.values[i].lowLines, pen, brush);
			this.cChartDrawer.drawPath(this.paths.values[i].highLines, pen, brush);

			if (this.paths.values[i].downBars) {
				brush = this.chart.upDownBars ? this.chart.upDownBars.downBarsBrush : null;
				pen = this.chart.upDownBars ? this.chart.upDownBars.downBarsPen : null;
				this.cChartDrawer.drawPath(this.paths.values[i].downBars, pen, brush);
			} else {
				brush = this.chart.upDownBars ? this.chart.upDownBars.upBarsBrush : null;
				pen = this.chart.upDownBars ? this.chart.upDownBars.upBarsPen : null;
				this.cChartDrawer.drawPath(this.paths.values[i].upBars, pen, brush);
			}
		}
	},

	_calculateLine: function (x, y, x1, y1) {

		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;


		path.moveTo(x * pathW, y * pathH);
		path.lnTo(x1 * pathW, y1 * pathH);

		return pathId;
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var pxToMm = this.chartProp.pxToMM;
		var min = this.valAx.scale[0];
		var max = this.valAx.scale[this.valAx.scale.length - 1];

		var digHeight = Math.abs(max - min);

		if (/*this.chartProp.min < 0 && this.chartProp.max <= 0*/min < 0 && max <= 0) {
			min = -1 * max;
		}

		var numCache = this.cChartDrawer.getNumCache(this.chart.series[0].val);
		if(!numCache) {
			return {x: null, y: null};
		}

		var koffX = this.chartProp.trueWidth / numCache.pts.length;
		var koffY = this.chartProp.trueHeight / digHeight;

		var point = this.chart.series[ser].val.numRef ? this.chart.series[ser].val.numRef.numCache.pts[val] : this.chart.series[ser].val.numLit.pts[val];

		var x = this.chartProp.chartGutter._left + (val) * koffX + koffX / 2;
		var y = this.chartProp.trueHeight - (point.val - min) * koffY + this.chartProp.chartGutter._top;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x / pxToMm - width / 2;
		var centerY = y / pxToMm - height / 2;
		var constMargin = 5 / pxToMm;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_calculateUpDownBars: function (x, y, x1, y1, width) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		path.moveTo((x - width / 2) * pathW, y * pathH);
		path.lnTo((x - width / 2) * pathW, y1 * pathH);
		path.lnTo((x + width / 2) * pathW, y1 * pathH);
		path.lnTo((x + width / 2) * pathW, y * pathH);
		path.lnTo((x - width / 2) * pathW, y * pathH);

		return pathId;
	}
};

/** @constructor */
function drawBubbleChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;

	this.paths = {};
}

drawBubbleChart.prototype = {
	constructor: drawBubbleChart,

	recalculate: function () {
		this.paths = {};

		this.catAx = this.chart.axId[0].xPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.catAx = this.cChartDrawer._searchChangedAxis(this.catAx);
		this.valAx = this.chart.axId[0].yPoints ? this.chart.axId[0] : this.chart.axId[1];
		this.valAx = this.cChartDrawer._searchChangedAxis(this.valAx);

		this._recalculateScatter();
	},

	draw: function () {
		this._drawScatter();
	},

	_recalculateScatter: function () {
		var xPoints = this.catAx.xPoints;
		var yPoints = this.valAx.yPoints;

		var seria, yVal, xVal, points, x, y, yNumCache, xNumCache;
		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			points = [];
			yNumCache = this.cChartDrawer.getNumCache(seria.yVal);
			if(!yNumCache) {
				continue;
			}

			for (var n = 0; n < yNumCache.pts.length; n++) {
				yVal = parseFloat(yNumCache.pts[n].val);

				xNumCache = this.cChartDrawer.getNumCache(seria.xVal);
				if (xNumCache && xNumCache.pts[n] && xNumCache.pts[n].val) {
					if (!isNaN(parseFloat(xNumCache.pts[n].val))) {
						xVal = parseFloat(xNumCache.pts[n].val);
					} else {
						xVal = n + 1;
					}
				} else {
					xVal = n + 1;
				}

				points[n] = {x: xVal, y: yVal}
			}

			for (var k = 0; k < points.length; k++) {
				y = this.cChartDrawer.getYPosition(points[k].y, this.valAx);
				x = this.cChartDrawer.getYPosition(points[k].x, this.catAx);


				if (!this.paths.points) {
					this.paths.points = [];
				}
				if (!this.paths.points[i]) {
					this.paths.points[i] = [];
				}

				this.paths.points[i][k] = this._calculateBubble(x, y, seria.bubbleSize, k);
			}
		}
	},

	_drawScatter: function () {
		var seria, brush, pen, markerBrush, markerPen, yNumCache;
		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			brush = seria.brush;
			pen = seria.pen;

			//draw bubble
			if (this.paths.points && this.paths.points[i]) {
				for (var k = 0; k < this.paths.points[i].length; k++) {
					yNumCache = this.cChartDrawer.getNumCache(this.chart.series[i].yVal);
					if(!yNumCache) {
						continue;
					}

					markerBrush = yNumCache.pts[k].compiledMarker.brush;
					markerPen = yNumCache.pts[k].compiledMarker.pen;

					//point
					this.cChartDrawer.drawPath(this.paths.points[i][k], markerPen, markerBrush, true);
				}
			}
		}
	},

	_calculateDLbl: function (chartSpace, ser, val) {
		var point;
		if (this.chart.series[ser - 1]) {
			//point = this.chart.series[ser - 1].yVal.numRef ? this.chart.series[ser - 1].yVal.numRef.numCache.pts[val] : this.chart.series[ser - 1].yVal.numLit.pts[val]
			point = this.cChartDrawer.getNumCache(this.chart.series[ser - 1].yVal);
		} else {
			//point = this.chart.series[ser].yVal.numRef ? this.chart.series[ser].yVal.numRef.numCache.pts[val] : this.chart.series[ser].yVal.numLit.pts[val];
			point = this.cChartDrawer.getNumCache(this.chart.series[ser].yVal);
		}
		if(!point) {
			return {x: null, y: null};
		}

		var path;

		if (this.paths && this.paths.points) {
			if (this.paths.points[ser] && this.paths.points[ser][val]) {
				var oPath = this.cChartSpace.GetPath(this.paths.points[ser][val].path);
				path = oPath.getCommandByIndex(0);
			}

		}

		if (!path) {
			return;
		}

		var x = path.X;
		var y = path.Y;

		var pxToMm = this.chartProp.pxToMM;
		var constMargin = 5 / pxToMm;

		var width = point.compiledDlb.extX;
		var height = point.compiledDlb.extY;

		var centerX = x - width / 2;
		var centerY = y - height / 2;

		switch (point.compiledDlb.dLblPos) {
			case c_oAscChartDataLabelsPos.b: {
				centerY = centerY + height / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.bestFit: {
				break;
			}
			case c_oAscChartDataLabelsPos.ctr: {
				break;
			}
			case c_oAscChartDataLabelsPos.l: {
				centerX = centerX - width / 2 - constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.r: {
				centerX = centerX + width / 2 + constMargin;
				break;
			}
			case c_oAscChartDataLabelsPos.t: {
				centerY = centerY - height / 2 - constMargin;
				break;
			}
		}

		if (centerX < 0) {
			centerX = 0;
		}
		if (centerX + width > this.chartProp.widthCanvas / pxToMm) {
			centerX = this.chartProp.widthCanvas / pxToMm - width;
		}

		if (centerY < 0) {
			centerY = 0;
		}
		if (centerY + height > this.chartProp.heightCanvas / pxToMm) {
			centerY = this.chartProp.heightCanvas / pxToMm - height;
		}

		return {x: centerX, y: centerY};
	},

	_calculateBubble: function (x, y, bubbleSize, k) {
		var defaultSize = 4;

		if (bubbleSize) {
			var maxSize, curSize, yPoints, maxDiamBubble, diffSize, maxArea;


			maxSize = this.cChartDrawer._getMaxMinValueArray(bubbleSize).max;
			curSize = bubbleSize[k].val;

			yPoints = this.valAx.yPoints ? this.valAx.yPoints : this.catAx.yPoints;
			maxDiamBubble = Math.abs(yPoints[1].pos - yPoints[0].pos) * 2;

			diffSize = maxSize / curSize;

			var isDiam = false;

			if (isDiam) {
				defaultSize = (maxDiamBubble / diffSize) / 2;
			} else {
				maxArea = 1 / 4 * (Math.PI * (maxDiamBubble * maxDiamBubble));
				defaultSize = Math.sqrt((maxArea / diffSize) / Math.PI);
			}
		}


		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		path.moveTo((x + defaultSize) * pathW, y * pathH);
		path.arcTo(defaultSize * pathW, defaultSize * pathW, 0, Math.PI * 2 * cToDeg);

		return pathId;
	}
};


/** @constructor */
function drawSurfaceChart(chart, chartsDrawer) {
	this.chartProp = chartsDrawer.calcProp;
	this.cChartDrawer = chartsDrawer;
	this.cChartSpace = chartsDrawer.cChartSpace;

	this.chart = chart;
	this.catAx = null;
	this.valAx = null;

	this.paths = {};
}

drawSurfaceChart.prototype = {
	constructor: drawSurfaceChart,

	recalculate: function () {
		this.paths = {};

		var countSeries = this.cChartDrawer.calculateCountSeries(this.chart);
		this.ptCount = countSeries.points;
		this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_CatAx);
		if(!this.catAx) {
			this.catAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_DateAx);
		}
		this.valAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_ValAx);
		this.serAx = this.cChartDrawer.getAxisFromAxId(this.chart.axId, AscDFH.historyitem_type_SerAx);

		this._recalculate();
	},

	draw: function () {
		this._draw();
	},

	_recalculate: function () {
		var yPoints = this.valAx.yPoints;
		var xPoints = this.catAx.xPoints;
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;

		var y, x, z, val, seria, dataSeries, idx, numCache, idxPoint, points = [], points3d = [], idx2, val2;
		for (var i = 0; i < this.chart.series.length; i++) {
			seria = this.chart.series[i];
			numCache = this.cChartDrawer.getNumCache(seria.val);

			if (!numCache) {
				continue;
			}

			dataSeries = numCache.pts;
			for (var n = 0; n < this.ptCount; n++) {
				//рассчитываем значения
				idx = dataSeries[n] && dataSeries[n].idx != null ? dataSeries[n].idx : n;

				//TODO временно заменил idx на n. позже нужно использовать idx  -> val = this._getYVal(idx, i);
				val = this._getYVal(n, i);
				if (null === val) {
					val = 0;
				}

				x = xPoints[n].pos * this.chartProp.pxToMM;
				y = this.cChartDrawer.getYPosition(val, this.valAx) * this.chartProp.pxToMM;
				if (this.serAx && this.serAx.zPoints[i]) {
					z = this.serAx.zPoints[i].pos;
				} else {
					z = (perspectiveDepth / (this.chart.series.length - 1)) * (i);
				}

				//рассчитываем значения
				idx2 = dataSeries[n + 1] && dataSeries[n + 1].idx != null ? dataSeries[n + 1].idx : null;

				//TODO временно заменил idx на n. позже нужно использовать idx
				//var val2 = Math.round(this._getYVal(n + 1, i) * roundInt) / roundInt;
				val2 = this._getYVal(n + 1, i);
				if (null === val2) {
					val2 = 0;
				}

				if (!points) {
					points = [];
				}
				if (!points[i]) {
					points[i] = [];
				}

				if (!points3d) {
					points3d = [];
				}
				if (!points3d[i]) {
					points3d[i] = [];
				}

				if (val != null) {
					points3d[i][n] = {x: x, y: y, z: z, val: val};
					var convertResult = this.cChartDrawer._convertAndTurnPoint(x, y, z);
					var x1 = convertResult.x;
					var y1 = convertResult.y;
					points[i][n] = {x: x1, y: y1, val: val};
				} else {
					points[i][n] = null;
				}
			}
		}

		this._calculateAllFaces(points, points3d);
	},


	_calculateAllFaces: function (points, points3d) {
		if (!this.paths.test) {
			this.paths.test = [];
		}
		if (!this.paths.test2) {
			this.paths.test2 = [];
		}

		for (var i = 0; i < points.length - 1; i++) {
			for (var j = 0; j < points[i].length - 1; j++) {
				var p1 = points[i][j];
				var p2 = points[i + 1][j];
				var p3 = points[i][j + 1];
				var p4 = points[i + 1][j + 1];

				this.paths.test.push(this.cChartDrawer._calculatePathFace(p1, p3, p4, p2, true));

				var p13d = points3d[i][j];
				var p23d = points3d[i + 1][j];
				var p33d = points3d[i][j + 1];
				var p43d = points3d[i + 1][j + 1];

				//рассчитываем отдельный сегмент - смотрим, может ли он располагаться в одной плоскости + делим его плоскостями сетки
				//p, p2, p21, p1 - точки данного сегмента
				//  p1-----p21
				//  |       |
				//  |       |
				//  p-------p2
				this._calculateFace(p1, p2, p3, p4, p13d, p23d, p33d, p43d);
			}
		}
	},

	_calculateFace: function (p, p1, p2, p21, p3d, p13d, p23d, p213d) {
		var t = this;

		//все стороны сегмента
		var lines = [];
		lines[0] = {p1: p3d, p2: p23d, p111: p, p222: p2};
		lines[1] = {p1: p13d, p2: p213d, p111: p1, p222: p21};
		lines[2] = {p1: p3d, p2: p13d, p111: p, p222: p1};
		lines[3] = {p1: p23d, p2: p213d, p111: p2, p222: p21};


		var pointsValue = [p1, p2, p21, p];
		if (this.cChartDrawer.isPointsLieIntoOnePlane(p3d, p13d, p213d, p23d))//не делим диагональю данный сегмент
		{
			this._getIntersectionPlanesAndLines(lines, pointsValue, true);
		} else//делим диагональю данный сегмент
		{
			var max = p.val;
			var maxIndex = 0;
			var arrVal = [p, p1, p2, p21];
			for (var m = 0; m < arrVal.length; m++) {
				if (arrVal[m].val > max) {
					max = arrVal[m].val;
					maxIndex = m;
				}
			}

			//разбиваем диагональю данный сегмент на два сегмента
			var lines1, pointsValue1, lines2, pointsValue2;
			if (p1.val + p2.val < p21.val + p.val) {
				//добавляем диагональ
				lines.push({p1: p213d, p2: p3d, p111: p21, p222: p});
				//this.paths.test.push(this._calculatePath(p21.x, p21.y, p.x, p.y));

				lines1 = [lines[0], lines[3], lines[4]];
				pointsValue1 = [p, p21, p2];
				lines2 = [lines[2], lines[1], lines[4]];
				pointsValue2 = [p, p1, p21];
			} else {
				//добавляем диагональ
				lines.push({p1: p13d, p2: p23d, p111: p1, p222: p2});
				//this.paths.test.push(this._calculatePath(p2.x, p2.y, p1.x, p1.y));

				lines1 = [lines[2], lines[0], lines[4]];
				pointsValue1 = [p, p1, p2];
				lines2 = [lines[4], lines[1], lines[3]];
				pointsValue2 = [p2, p1, p21];
			}

			//для поверхностных диаграмм без заливки
			var bIsWireframeChart = false;

			//находим пересечение двух сегментов с плоскостями сетки
			var pointsFace1 = this._getIntersectionPlanesAndLines(lines1, pointsValue1, !bIsWireframeChart);
			var pointsFace2 = this._getIntersectionPlanesAndLines(lines2, pointsValue2, !bIsWireframeChart);

			if (bIsWireframeChart) {
				var lengthFaces = Math.max(pointsFace1.length, pointsFace2.length);
				for (var l = 0; l < lengthFaces; l++) {
					var newPath = null;
					if (pointsFace1[l] && pointsFace2[l]) {
						var cleanPoints1 = this._getArrayWithoutRepeatePoints(pointsFace1[l]);
						var cleanPoints2 = this._getArrayWithoutRepeatePoints(pointsFace2[l]);

						if (cleanPoints1.length >= 3 && cleanPoints2.length >= 3) {
							newPath = cleanPoints1.concat(cleanPoints2);
						}
					} else if (pointsFace1[l]) {
						newPath = pointsFace1[l];
					} else if (pointsFace2[l]) {
						newPath = pointsFace2[l];
					}


					if (newPath) {
						if (!t.paths.test2[l]) {
							t.paths.test2[l] = [];
						}

						var path2 = t._calculateTempFace(newPath);
						t.paths.test2[l].push(path2);
					}
				}
			}

			//TODO временно убираю. если будут проблемы в отрисовке - раскомментировать!
			/*var lengthFaces = Math.max(pointsFace1.length, pointsFace2.length);
			 for(var l = 0; l < lengthFaces; l++)
			 {
			 if(pointsFace1[l] && pointsFace2[l])
			 {
			 //находим две точки, принадлежащие диагональной прямой. у обоих сегментов они должны быть
			 var lineEquation = t.cChartDrawer.getLineEquation2d(lines[4].p111, lines[4].p222);

			 var points1 = [];
			 for(var s = 0; s < pointsFace1[l].length; s++)
			 {
			 if(null === this._isEqualPoints(points1, pointsFace1[l][s]) && t.cChartDrawer.isPoint2DLieOnLine(lineEquation, pointsFace1[l][s]))
			 {
			 points1.push(pointsFace1[l][s]);
			 }
			 }

			 var points2 = []
			 for(var s = 0; s < pointsFace2[l].length; s++)
			 {
			 if(null === this._isEqualPoints(points2, pointsFace2[l][s]) && t.cChartDrawer.isPoint2DLieOnLine(lineEquation, pointsFace2[l][s]))
			 {
			 points2.push(pointsFace2[l][s]);
			 }
			 }

			 if(points1.length < 2 && points2.length === 2)
			 {
			 pointsFace1[l].push(points2[0]);
			 pointsFace1[l].push(points2[1]);
			 }
			 else if(points2.length < 2 && points1.length === 2)
			 {
			 pointsFace2[l].push(points1[0]);
			 pointsFace2[l].push(points1[1]);
			 }
			 }


			 if(!t.paths.test2[l])
			 {
			 t.paths.test2[l] = [];
			 }
			 if(pointsFace1[l])
			 {
			 var path1 = t._calculateTempFace(pointsFace1[l]);
			 t.paths.test2[l].push(path1);
			 }
			 if(pointsFace2[l])
			 {
			 var path2 = t._calculateTempFace(pointsFace2[l]);
			 t.paths.test2[l].push(path2);
			 }
			 }*/
		}
	},


	_getIntersectionPlanesAndLines: function (lines, pointsValue, bIsAddIntoPaths) {
		var t = this;
		var yPoints = this.valAx.yPoints;
		var xPoints = this.catAx.xPoints;
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;

		var getGridPlain = function (index) {
			var gridX1 = xPoints[0].pos * t.chartProp.pxToMM;
			var gridX2 = xPoints[xPoints.length - 1].pos * t.chartProp.pxToMM;

			var gridY1 = yPoints[index].pos * t.chartProp.pxToMM;
			var gridY2 = yPoints[index].pos * t.chartProp.pxToMM;
			var gridPlain = t.cChartDrawer.getPlainEquation({x: gridX1, y: gridY1, z: 0}, {x: gridX2, y: gridY2, z: 0}, {x: gridX2, y: gridY2, z: perspectiveDepth});

			return gridPlain;
		};

		var getMinMaxValArray = function (pointsValue) {
			var min, max;
			if (pointsValue.length === 4) {
				min = Math.min(pointsValue[0].val, pointsValue[1].val, pointsValue[2].val, pointsValue[3].val);
				max = Math.max(pointsValue[0].val, pointsValue[1].val, pointsValue[2].val, pointsValue[3].val);
			} else {
				min = Math.min(pointsValue[0].val, pointsValue[1].val, pointsValue[2].val);
				max = Math.max(pointsValue[0].val, pointsValue[1].val, pointsValue[2].val);
			}
			return {min: min, max: max}
		};

		var calculateFaceBetween2GridLines = function (minVal, maxVal, k, pointsValue, res) {
			var result = false;

			if (yPoints[k - 1] && minVal >= yPoints[k - 1].val && maxVal <= yPoints[k].val) {
				var p1 = pointsValue[0];
				var p2 = pointsValue[1];
				var p3 = pointsValue[2];
				var p4 = pointsValue[3] ? pointsValue[3] : pointsValue[2];

				var path = t._calculateTempFace([p1, p2, p3, p4]);

				var addIndex = k;
				if (minVal === maxVal && yPoints[k] && minVal === yPoints[k].val) {
					addIndex = k + 1;
				}

				if (bIsAddIntoPaths) {
					if (!t.paths.test2[addIndex]) {
						t.paths.test2[addIndex] = [];
					}
					t.paths.test2[addIndex].push(path);
				}

				res[k] = [p1, p2, p3, p4];
				result = true;
			}

			return result;
		};

		var minMaxVal = getMinMaxValArray(pointsValue);
		var minVal = minMaxVal.min;
		var maxVal = minMaxVal.max;

		var res = [];

		var prevPoints = null;
		for (var k = 0; k < yPoints.length; k++) {
			//если сегмент весь находится между двумя соседними плоскостями сетки, то есть ни с одной из них не имеет пересечений
			if (calculateFaceBetween2GridLines(minVal, maxVal, k, pointsValue, res)) {
				break;
			}
			//если значение сетки больше максимального значения сегмента
			if (yPoints[k - 1] && yPoints[k].val > maxVal && yPoints[k - 1].val >= maxVal) {
				break;
			}

			//точки, которые находятся между данными плоскостями сетки(или лежат на них), обязательно должны войти в сегмент
			var pointNeedAddIntoFace = null;
			if (yPoints[k - 1]) {
				for (var i = 0; i < pointsValue.length; i++) {
					if (yPoints[k - 1].val <= pointsValue[i].val && yPoints[k].val >= pointsValue[i].val) {
						if (null === pointNeedAddIntoFace) {
							pointNeedAddIntoFace = [];
						}

						pointNeedAddIntoFace.push(pointsValue[i]);
					}
				}
			}


			var isCalculatePrevPoints = false;
			if (null === prevPoints) {
				for (var j = 0; j < pointsValue.length; j++) {
					if (pointsValue[j].val <= yPoints[k].val) {
						if (!prevPoints) {
							prevPoints = [];
						}
						prevPoints.push(pointsValue[j]);
						isCalculatePrevPoints = true;
					}
				}
			}

			//находим точки пересечения с текущей плоскостью сетки
			var gridPlane = getGridPlain(k);
			var points = this._getIntersectionPlaneAndLines(gridPlane, lines, pointsValue);

			if (!isCalculatePrevPoints && null === points && prevPoints) {
				for (var j = 0; j < pointsValue.length; j++) {
					if (pointsValue[j].val >= yPoints[k - 1].val) {
						if (!points) {
							points = [];
						}
						points.push(pointsValue[j]);
						isCalculatePrevPoints = true;
					}
				}
			}

			var arrPoints = null, p1, p2, p3, p4;
			if (null !== points && prevPoints) {
				p1 = prevPoints[0];
				p2 = prevPoints[1] ? prevPoints[1] : prevPoints[0];
				p3 = points[0];
				p4 = points[1] ? points[1] : points[0];

				arrPoints = [p1, p2, p3, p4];
				if (points[2]) {
					arrPoints.push(points[2]);
				}
				res[k] = arrPoints;
			} else if (prevPoints && prevPoints.length === 3 && !points && isCalculatePrevPoints) {
				p1 = prevPoints[0];
				p2 = prevPoints[1];
				p3 = prevPoints[2];
				p4 = prevPoints[3] ? prevPoints[3] : prevPoints[2];

				arrPoints = [p1, p2, p3, p4];
				res[k] = arrPoints;
			}

			//добавляем точки, которые обязательно должны присутвовать в сегменте
			if (arrPoints && null !== pointNeedAddIntoFace) {
				for (var i = 0; i < pointNeedAddIntoFace.length; i++) {
					arrPoints.push(pointNeedAddIntoFace[i]);
				}
			}

			//add to path array
			if (arrPoints && bIsAddIntoPaths) {
				var path = t._calculateTempFace(arrPoints);
				if (!t.paths.test2[k]) {
					t.paths.test2[k] = [];
				}
				t.paths.test2[k].push(path);
			}

			if (points !== null) {
				prevPoints = points;
			}
		}

		return res;
	},

	_getIntersectionPlaneAndLines: function (gridPlain, lines, pointsValue) {
		var res = null;

		var clearIntersectionPoints = [];
		var segmentIntersectionPoints = [];

		//n --> lines --> 0 - down, 1 - up, 2 - left, 3 - right, 4 - diagonal
		for (var n = 0; n < lines.length; n++) {
			var convertResult = this.cChartDrawer.isIntersectionPlainAndLineSegment(gridPlain, lines[n].p1, lines[n].p2, lines[n].p111, lines[n].p222);
			if (!convertResult) {
				continue;
			}

			if (null === this._isEqualPoints(pointsValue, convertResult)) {
				clearIntersectionPoints.push(convertResult);
			} else {
				if (null === this._isEqualPoints(segmentIntersectionPoints, convertResult)) {
					segmentIntersectionPoints.push(convertResult);
				}
			}
		}

		var p1, p2;
		if (!segmentIntersectionPoints.length) {
			if (clearIntersectionPoints.length === 2)//две точки, не равняющиеся ни одной точке сегмента
			{
				p1 = clearIntersectionPoints[0];
				p2 = clearIntersectionPoints[1];

				res = [p1, p2];
			} else if (clearIntersectionPoints.length === 1)//одна точка, не равняющиеся ни одной точке сегмента
			{
				p1 = clearIntersectionPoints[0];

				res = [p1];
			}
		} else if (segmentIntersectionPoints.length && clearIntersectionPoints.length) {
			if (1 === segmentIntersectionPoints.length && 1 === clearIntersectionPoints.length) {
				p1 = segmentIntersectionPoints[0];
				p2 = clearIntersectionPoints[0];

				res = [p1, p2];
			}
		} else if (segmentIntersectionPoints.length) {
			if (2 === segmentIntersectionPoints.length) {
				p1 = segmentIntersectionPoints[0];
				p2 = segmentIntersectionPoints[1];

				res = [p1, p2];
			} else if (1 === segmentIntersectionPoints.length) {
				p1 = segmentIntersectionPoints[0];

				res = [p1];
			}
		}

		return res;
	},

	_isEqualPoints: function (arr, point) {
		var res = null;

		for (var p = 0; p < arr.length; p++) {
			if (arr[p] && parseInt(point.x) === parseInt(arr[p].x) && parseInt(point.y) === parseInt(arr[p].y)) {
				res = p;
				break;
			}
		}

		return res;
	},

	_getArrayWithoutRepeatePoints: function (arr) {
		var newArray = [];

		for (var i = 0; i < arr.length; i++) {
			if (null === this._isEqualPoints(newArray, arr[i])) {
				newArray.push(arr[i]);
			}
		}

		return newArray;
	},

	_calculateTempFace: function (points) {
		var summX = 0;
		var summY = 0;
		for (var i = 0; i < points.length; i++) {
			summX += points[i].x;
			summY += points[i].y;
		}
		var x = 1 / (points.length) * (summX);
		var y = 1 / (points.length) * (summY);

		var sortArray = [];
		var repeatePoint = [];
		var nIndividualPoints = 0;
		for (var i = 0; i < points.length; i++) {
			var tan = Math.atan2(points[i].x - x, points[i].y - y);
			if (!repeatePoint[tan]) {
				sortArray[i] = {tan: tan, point: points[i]};
				repeatePoint[tan] = 1;
				nIndividualPoints++;
			}
		}

		var path = null
		if (nIndividualPoints > 2) {
			sortArray.sort(function sortArr(a, b) {
				return b.tan - a.tan;
			});

			path = this.cChartDrawer.calculatePathFacesArray(sortArray, true);
		}

		return path
	},


	_getYVal: function (n, i) {
		var idxPoint = this.cChartDrawer.getIdxPoint(this.chart.series[i], n);
		var val = idxPoint ? parseFloat(idxPoint.val) : null;

		return val;
	},

	_calculatePath: function (x, y, x1, y1) {
		var pathId = this.cChartSpace.AllocPath();
		var path = this.cChartSpace.GetPath(pathId);

		var pxToMm = this.chartProp.pxToMM;

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;
		var gdLst = [];

		path.pathH = pathH;
		path.pathW = pathW;
		gdLst["w"] = 1;
		gdLst["h"] = 1;

		path.moveTo(x / pxToMm * pathW, y / pxToMm * pathH);
		path.lnTo(x1 / pxToMm * pathW, y1 / pxToMm * pathH);

		return pathId;
	},

	_draw: function () {
		var style = AscFormat.CHART_STYLE_MANAGER.getStyleByIndex(this.cChartSpace.style);

		for (var i = 0; i < this.paths.test2.length; i++) {
			if (!this.paths.test2[i]) {
				continue;
			}

			for (var j = 0; j < this.paths.test2[i].length; j++) {
				style = this.chart.compiledBandFormats[i - 1];
				var brush = style && style.spPr ? style.spPr.Fill : null;
				var pen = style && style.spPr ? style.spPr.ln : null;

				//линии пока делаю по цвету как и заливку
				if (!pen || (pen && 0 === pen.w)) {
					pen = AscFormat.CreatePenFromParams(brush, undefined, undefined, undefined, undefined, 0.13);
				}

				this.cChartDrawer.drawPath(this.paths.test2[i][j], pen, brush);
			}
		}
	}
};

function axisChart() {
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	this.axis = null;
	this.paths = {};

	this.axisType = null;
}

axisChart.prototype = {
	constructor: axisChart,

	draw: function (chartsDrawer, isDrawGrid) {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;

		if(isDrawGrid) {
			this._drawGridLines();
		} else {
			if (!this.axis.isRadarCategories()) {
				this._drawAxis();
				this._drawTickMark();
			}
		}
	},

	recalculate: function (chartsDrawer, axisId) {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		this.axis = axisId;

		this.paths = {};

		this.axisType = this.axis.getObjectType();
		if (!(this.axisType === AscDFH.historyitem_type_SerAx)) {
			this._calculateGridLines();
		}

		if (this.axis.bDelete !== true) {
			if (this.axisType === AscDFH.historyitem_type_ValAx) {
				this._calculateValAxis();
				this._calculateValTickMark();
			} else if (this.axisType === AscDFH.historyitem_type_CatAx) {
				this._calculateCatAxis();
				this._calculateCatTickMark();
			} else if (this.axisType === AscDFH.historyitem_type_SerAx) {
				this._calculateSerAxis();
				this._calculateSerTickMark();
			}
		}
	},

	_calculateGridLines: function () {

		var paths;
		if (this.axis.isRadarValues()) {
			paths = this.cChartDrawer._getRadarGridLines(this.axis);
		} else {
			if (this.axis.axPos === window['AscFormat'].AX_POS_L || this.axis.axPos === window['AscFormat'].AX_POS_R) {
				//ось слева или справа, линии горизонтальные
				paths = this.cChartDrawer.getHorizontalGridLines(this.axis, this.axisType === AscDFH.historyitem_type_CatAx);
			} else {
				paths = this.cChartDrawer.getVerticalGridLines(this.axis, this.axisType === AscDFH.historyitem_type_CatAx);
			}
		}
		this.paths.gridLines = paths ? paths.gridLines : null;
		this.paths.minorGridLines = paths ? paths.minorGridLines : null;
	},

	_calculateValAxis: function () {
		let axisPos;
		let left = this.chartProp.chartGutter._left / this.chartProp.pxToMM;
		let right = (this.chartProp.widthCanvas - this.chartProp.chartGutter._right) / this.chartProp.pxToMM;
		let top = this.chartProp.chartGutter._top / this.chartProp.pxToMM;
		let bottom = (this.chartProp.heightCanvas - this.chartProp.chartGutter._bottom) / this.chartProp.pxToMM;

		if (this.axis && this.axis.isRadarValues()) {

			var axis = this.axis;
			var yPoints = axis.yPoints;
			var orientation = !axis.isReversed();
			var xCenter = axis.posX;

			let yCenter, trueHeight;
			if (yPoints.length > 0) {
				yCenter = orientation ? yPoints[0].pos : yPoints[yPoints.length - 1].pos;
			} else {
				trueHeight = this.chartProp.trueHeight;
				yCenter = (this.chartProp.chartGutter._top + trueHeight / 2) / this.chartProp.pxToMM;
			}

			let radarChart = this.cChartDrawer.getChartByType(AscDFH.historyitem_type_RadarChart);
			let ptCount = radarChart && radarChart.series && this.cChartDrawer.getPtCount(radarChart.series);
			if (!ptCount) {
				return;
			}
			let tempAngle = 2 * Math.PI / ptCount;

			let radius, x, y;
			radius = Math.abs(yPoints[yPoints.length - 1].pos - yPoints[0].pos);
			for (let n = 0; n < ptCount; n++) {
				x = xCenter + radius * Math.sin(n * tempAngle);
				y = yCenter - radius * Math.cos(n * tempAngle);

				if (!this.paths.axisLine) {
					this.paths.axisLine = [];
				}
				this.paths.axisLine.push(this._calculateLine(xCenter, yCenter, x, y));
			}

		} else if (this.axis.axPos === window['AscFormat'].AX_POS_T || this.axis.axPos === window['AscFormat'].AX_POS_B) {
			axisPos = this.axis.posY;
			this.paths.axisLine = this._calculateLine(left, axisPos, right, axisPos);
		} else {
			axisPos = this.axis.posX;
			this.paths.axisLine = this._calculateLine(axisPos, top, axisPos, bottom);
		}
	},

	_calculateCatAxis: function () {
		var axisPos;
		var left = this.chartProp.chartGutter._left / this.chartProp.pxToMM;
		var right = (this.chartProp.widthCanvas - this.chartProp.chartGutter._right) / this.chartProp.pxToMM;
		var top = this.chartProp.chartGutter._top / this.chartProp.pxToMM;
		var bottom = (this.chartProp.heightCanvas - this.chartProp.chartGutter._bottom) / this.chartProp.pxToMM;

		if (this.axis.axPos === window['AscFormat'].AX_POS_R || this.axis.axPos === window['AscFormat'].AX_POS_L) {
			axisPos = this.axis.posX;
			this.paths.axisLine = this._calculateLine(axisPos, top, axisPos, bottom);
		} else {
			axisPos = this.axis.posY;
			this.paths.axisLine = this._calculateLine(left, axisPos, right, axisPos);
		}
	},

	_calculateSerAxis: function () {
		var nullPositionOx = this.axis.posY * this.chartProp.pxToMM;
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		var positionX = this.cChartDrawer.processor3D.calculateXPositionSerAxis();

		var x, y, x1, y1;

		if (positionX) {
			x = this.chartProp.chartGutter._left;
			x1 = this.chartProp.chartGutter._left;
		} else {
			x = (this.chartProp.widthCanvas - this.chartProp.chartGutter._right);
			x1 = (this.chartProp.widthCanvas - this.chartProp.chartGutter._right);
		}
		y = nullPositionOx;
		var convertResult = this.cChartDrawer._convertAndTurnPoint(x, y, 0);
		x = convertResult.x / this.chartProp.pxToMM;
		y = convertResult.y / this.chartProp.pxToMM;

		y1 = nullPositionOx;
		convertResult = this.cChartDrawer._convertAndTurnPoint(x1, y1, perspectiveDepth);
		x1 = convertResult.x / this.chartProp.pxToMM;
		y1 = convertResult.y / this.chartProp.pxToMM;


		this.paths.axisLine = this._calculateLine(x, y, x1, y1);
	},

	_getTickmarksPropsSer: function () {
		var widthLine = 0, crossMajorStep = 0;
		switch (this.axis.majorTickMark) {
			case c_oAscTickMark.TICK_MARK_CROSS: {
				widthLine = 5;
				crossMajorStep = 5;
				break;
			}
			case c_oAscTickMark.TICK_MARK_IN: {
				widthLine = 5;
				break;
			}
			case c_oAscTickMark.TICK_MARK_NONE: {
				widthLine = 0;
				break;
			}
			case c_oAscTickMark.TICK_MARK_OUT: {
				widthLine = -5;
				break;
			}
		}

		if (this.axis.axPos === window['AscFormat'].AX_POS_B) {
			widthLine = -widthLine;
			crossMajorStep = -crossMajorStep;
		}

		return {widthLine: widthLine, crossMajorStep: crossMajorStep};
	},

	_getTickmarksProps: function () {
		var widthLine = 0, widthMinorLine = 0;
		var crossMajorStep = 0;
		var crossMinorStep = 0;
		switch (this.axis.majorTickMark) {
			case c_oAscTickMark.TICK_MARK_CROSS: {
				widthLine = 5;
				crossMajorStep = 5;
				break;
			}
			case c_oAscTickMark.TICK_MARK_IN: {
				widthLine = 5;
				break;
			}
			case c_oAscTickMark.TICK_MARK_NONE: {
				widthLine = 0;
				break;
			}
			case c_oAscTickMark.TICK_MARK_OUT: {
				widthLine = -5;
				break;
			}
		}

		switch (this.axis.minorTickMark) {
			case c_oAscTickMark.TICK_MARK_CROSS: {
				widthMinorLine = 3;
				crossMinorStep = 3;
				break;
			}
			case c_oAscTickMark.TICK_MARK_IN: {
				widthMinorLine = 3;
				break;
			}
			case c_oAscTickMark.TICK_MARK_NONE: {
				widthMinorLine = 0;
				break;
			}
			case c_oAscTickMark.TICK_MARK_OUT: {
				widthMinorLine = -3;
				break;
			}
		}

		var axPos = this.axis.axPos;
		if (axPos !== window['AscFormat'].AX_POS_L && axPos !== window['AscFormat'].AX_POS_T) {
			widthMinorLine = -widthMinorLine;
			widthLine = -widthLine;
			crossMajorStep = -crossMajorStep;
			crossMinorStep = -crossMinorStep;
		}
		return {
			widthLine: widthLine, crossMajorStep: crossMajorStep,
			widthMinorLine: widthMinorLine, crossMinorStep: crossMinorStep
		};
	},

	_calculateCatTickMark: function () {
		var tickmarksProps = this._getTickmarksProps();
		var widthLine = tickmarksProps.widthLine, widthMinorLine = tickmarksProps.widthMinorLine;
		var crossMajorStep = tickmarksProps.crossMajorStep, crossMinorStep = tickmarksProps.crossMinorStep;

		var orientation = this.axis ? this.axis.getOrientation() : ORIENTATION_MIN_MAX;
		var minorStep, posX, posY, k, firstDiff = 0;
		var tickMarkSkip = this.axis.tickMarkSkip ? this.axis.tickMarkSkip : 1;

		var pathId = this.cChartSpace.AllocPath(), path = this.cChartSpace.GetPath(pathId);

		var i, n;
		var minorLinesCount = 2;
		if (this.axis.axPos === window['AscFormat'].AX_POS_R || this.axis.axPos === window['AscFormat'].AX_POS_L) {
			var yPoints = this.axis.yPoints;

			if(!yPoints) {
				return;
			}

			var stepY = yPoints[1] ? Math.abs(yPoints[1].pos - yPoints[0].pos) : Math.abs(yPoints[0].pos - this.chartProp.chartGutter._bottom / this.chartProp.pxToMM);
			minorStep = stepY / minorLinesCount;
			posX = this.axis.posX;

			//сдвиг, если положение оси - между делениями
			if (this.axis.crossAx.crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
				//TODO избавиться от использовения параметров лругой оси!!!
				firstDiff = yPoints[1] ? Math.abs(yPoints[1].pos - yPoints[0].pos) : Math.abs(yPoints[0].pos - this.cChartSpace.chart.plotArea.valAx.posY) * 2;
			}

			if (orientation !== ORIENTATION_MIN_MAX) {
				minorStep = -minorStep;
				firstDiff = -firstDiff;
			}

			for (i = 0; i < yPoints.length; i++) {
				k = i * tickMarkSkip;
				if (k >= yPoints.length) {
					break;
				}

				if(yPoints[k].val < 0) {
					continue;
				}
				//основные линии
				posY = yPoints[k].pos + firstDiff / 2;

				if (!this.paths.tickMarks) {
					this.paths.tickMarks = pathId;
				}
				this._calculateLine(posX, posY, posX + widthLine / this.chartProp.pxToMM, posY, path);

				if (((i + 1) * tickMarkSkip) === yPoints.length)//если последняя основная линия, то рисуем её
				{
					var posYtemp = yPoints[yPoints.length - 1].pos - firstDiff / 2;
					this._calculateLine(posX - crossMajorStep / this.chartProp.pxToMM, posYtemp, posX + widthLine / this.chartProp.pxToMM, posYtemp, path);
				}


				//промежуточные линии
				if (widthMinorLine !== 0) {
					for (n = 1; n < minorLinesCount; n++) {
						var posMinorY = posY - n * minorStep * tickMarkSkip;

						if (((posMinorY < yPoints[yPoints.length - 1].pos - firstDiff / 2) &&
							orientation === ORIENTATION_MIN_MAX) ||
							((posMinorY > yPoints[yPoints.length - 1].pos - firstDiff / 2) &&
							orientation !== ORIENTATION_MIN_MAX)) {
							break;
						}

						this._calculateLine(posX - crossMinorStep / this.chartProp.pxToMM, posMinorY, posX + widthMinorLine / this.chartProp.pxToMM, posMinorY, path);
					}
				}
			}
		} else {
			var xPoints = this.axis.xPoints;

			if(!xPoints) {
				return;
			}

			var stepX = xPoints[1] ? Math.abs(xPoints[1].pos - xPoints[0].pos) : Math.abs(xPoints[0].pos - this.axis.posX) * 2;
			minorStep = stepX / minorLinesCount;
			posY = this.axis.posY;

			var posMinorX;
			if (this.axis.crossAx.crossBetween === AscFormat.CROSS_BETWEEN_BETWEEN) {
				if (xPoints[1]) {
					firstDiff = Math.abs(xPoints[1].pos - xPoints[0].pos);
				} else if (this.cChartSpace.chart.plotArea.valAx.posX) {
					firstDiff = Math.abs(this.cChartSpace.chart.plotArea.valAx.posX - xPoints[0].pos) * 2;
				}
			}

			if (orientation !== ORIENTATION_MIN_MAX) {
				minorStep = -minorStep;
				firstDiff = -firstDiff;
			}

			//сам рассчёт основных и промежуточных линий
			for (i = 0; i < xPoints.length; i++) {
				k = i * tickMarkSkip;
				if (k >= xPoints.length) {
					break;
				}

				if(xPoints[k].val < 0) {
					continue;
				}

				posX = xPoints[k].pos - firstDiff / 2;
				if (!this.paths.tickMarks) {
					this.paths.tickMarks = pathId;
				}
				this._calculateLine(posX, posY - crossMajorStep / this.chartProp.pxToMM, posX, posY + widthLine / this.chartProp.pxToMM, path);

				if (((i + 1) * tickMarkSkip) === xPoints.length)//если последняя основная линия, то рисуем её
				{
					var posXtemp = xPoints[xPoints.length - 1].pos + firstDiff / 2;
					this._calculateLine(posXtemp, posY - crossMajorStep / this.chartProp.pxToMM, posXtemp, posY + widthLine / this.chartProp.pxToMM, path);
				}

				//промежуточные линии
				if (widthMinorLine !== 0) {
					for (n = 1; n < minorLinesCount; n++) {
						posMinorX = posX + n * minorStep * tickMarkSkip;

						if (((posMinorX > xPoints[xPoints.length - 1].pos + firstDiff / 2) &&
							orientation === ORIENTATION_MIN_MAX) ||
							((posMinorX < xPoints[xPoints.length - 1].pos + firstDiff / 2) &&
							orientation !== ORIENTATION_MIN_MAX)) {
							break;
						}

						this._calculateLine(posMinorX, posY - crossMinorStep / this.chartProp.pxToMM, posMinorX, posY + widthMinorLine / this.chartProp.pxToMM, path);
					}
				}
			}
		}
	},

	_calculateValTickMark: function () {
		let orientation = this.axis ? this.axis.getOrientation() : ORIENTATION_MIN_MAX;
		let tickmarksProps = this._getTickmarksProps();
		let widthLine = tickmarksProps.widthLine;
		let widthMinorLine = tickmarksProps.widthMinorLine / this.chartProp.pxToMM;
		let crossMajorStep = tickmarksProps.crossMajorStep, crossMinorStep = tickmarksProps.crossMinorStep;

		let minorLinesCount = 5;
		let points, minorStep, posY, posX;
		let axPos = this.axis.axPos;

		let pathId = this.cChartSpace.AllocPath();
		let path = this.cChartSpace.GetPath(pathId);

		if (this.axis && this.axis.isRadarValues()) {
			points = this.axis.yPoints;
			if(!points || !widthMinorLine) {
				return;
			}

			let axis = this.axis;
			let xCenter = axis.posX;
			let yCenter = orientation ? points[0].pos : points[points.length - 1].pos;

			let getLineXY = function (_posY, _angle, _width) {
				let _radius = Math.abs(yCenter - _posY);

				//first point on ray from center
				x1 = xCenter + _radius * Math.sin(_angle);
				y1 = yCenter - _radius * Math.cos(_angle);

				x2 = x1 + _width * Math.sin(Math.PI / 2 - _angle);
				y2 = y1 + _width * Math.cos(Math.PI / 2 - _angle);
			};

			let stepY = points[1] ? Math.abs(points[1].pos - points[0].pos) : Math.abs(points[0].pos - this.chartProp.chartGutter._bottom / this.chartProp.pxToMM);
			minorStep = stepY / minorLinesCount;

			let radarChart = this.cChartDrawer.getChartByType(AscDFH.historyitem_type_RadarChart);
			let ptCount = radarChart && radarChart.series && this.cChartDrawer.getPtCount(radarChart.series);

			let tempAngle = 2 * Math.PI / ptCount;
			let x1, y1, x2, y2;
			for (let nDataPt = 0; nDataPt < ptCount; nDataPt++) {
				for (let nAxisPt = 0; nAxisPt < points.length; nAxisPt++) {
					if (!this.paths.tickMarks) {
						this.paths.tickMarks = pathId;
					}

					//main lines
					let angle = nDataPt * tempAngle;

					posY = points[nAxisPt].pos;
					getLineXY(posY, angle, widthMinorLine);
					this._calculateLine(x1, y1, x2, y2, path);

					//minor lines
					if (!((!orientation && nAxisPt === 0) || (orientation && nAxisPt === points.length - 1))) {
						for (let nMinor = 0; nMinor < minorLinesCount; nMinor++) {
							let posMinorY = posY - nMinor * minorStep;
							getLineXY(posMinorY, angle, widthMinorLine);
							this._calculateLine(x1, y1, x2, y2, path);
						}
					}
				}
			}

		} else if (axPos === window['AscFormat'].AX_POS_T || axPos === window['AscFormat'].AX_POS_B) {
			points = this.axis.xPoints;
			if(!points) {
				return;
			}

			let stepX = points[1] ? Math.abs(points[1].pos - points[0].pos) : Math.abs(points[0].pos - this.chartProp.chartGutter._bottom / this.chartProp.pxToMM);
			minorStep = stepX / minorLinesCount;
			posY = this.axis.posY;

			let posMinorX;
			for (let i = 0; i < points.length; i++) {
				posX = points[i].pos;
				if (!this.paths.tickMarks) {
					this.paths.tickMarks = pathId;
				}
				this._calculateLine(posX, posY - crossMajorStep / this.chartProp.pxToMM, posX, posY + widthLine / this.chartProp.pxToMM, path);

				//промежуточные линии
				if (widthMinorLine !== 0 && !((orientation === ORIENTATION_MIN_MAX && i === points.length - 1) || (orientation !== ORIENTATION_MIN_MAX && i === 0))) {
					for (let n = 0; n < minorLinesCount; n++) {
						posMinorX = posX + n * minorStep;

						this._calculateLine(posMinorX, posY - crossMinorStep / this.chartProp.pxToMM, posMinorX, posY + widthMinorLine, path);
					}
				}
			}
		} else {
			points = this.axis.yPoints;
			if(!points) {
				return;
			}

			let stepY = points[1] ? Math.abs(points[1].pos - points[0].pos) : Math.abs(points[0].pos - this.chartProp.chartGutter._bottom / this.chartProp.pxToMM);
			minorStep = stepY / minorLinesCount;
			posX = this.axis.posX;

			let posMinorY;
			for (let i = 0; i < points.length; i++) {
				//основные линии
				posY = points[i].pos;

				if (!this.paths.tickMarks) {
					this.paths.tickMarks = pathId;
				}
				this._calculateLine(posX - crossMajorStep / this.chartProp.pxToMM, posY, posX + widthLine / this.chartProp.pxToMM, posY, path);

				//промежуточные линии
				if (widthMinorLine !== 0 && !((orientation === ORIENTATION_MIN_MAX && i === points.length - 1) || (orientation !== ORIENTATION_MIN_MAX && i === 0))) {
					for (let n = 0; n < minorLinesCount; n++) {
						posMinorY = posY - n * minorStep;

						this._calculateLine(posX - crossMinorStep / this.chartProp.pxToMM, posMinorY, posX + widthMinorLine, posMinorY, path);
					}
				}
			}
		}
	},

	_calculateSerTickMark: function () {
		let perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		let tickmarksProps = this._getTickmarksPropsSer();
		let widthLine = tickmarksProps.widthLine;

		let seriesCount = this.axis.labels && this.axis.labels.count;
		if (seriesCount == null) {
			let _chart = this.axis.parent && this.axis.parent.chart;
			let countSeries = this.cChartDrawer.calculateCountSeries(_chart);
			seriesCount = countSeries.series;
		}
		
		if (widthLine !== 0) {
			let positionX = this.cChartDrawer.processor3D.calculateXPositionSerAxis();
			let stepY = perspectiveDepth / seriesCount;
			let startX;
			if (positionX) {
				startX = this.chartProp.chartGutter._left;
			} else {
				startX = this.chartProp.widthCanvas - this.chartProp.chartGutter._right;
			}

			let startY = this.axis.posY * this.chartProp.pxToMM;

			for (let i = 0; i <= seriesCount; i++) {
				//основные линии
				if (!this.paths.tickMarks) {
					this.paths.tickMarks = [];
				}

				let convertResult = this.cChartDrawer._convertAndTurnPoint(startX, startY, i * stepY);
				let x = convertResult.x / this.chartProp.pxToMM;
				let y = convertResult.y / this.chartProp.pxToMM;

				this.paths.tickMarks[i] = this._calculateLine(x, y, x + widthLine / this.chartProp.pxToMM, y);
			}
		}
	},

	_calculateLine: function (x, y, x1, y1, path) {
		var pathId;
		if(!path) {
			pathId = this.cChartSpace.AllocPath();
			path = this.cChartSpace.GetPath(pathId);
		}

		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		if (this.axisType === AscDFH.historyitem_type_SerAx) {
			path.moveTo(x * pathW, y * pathH);
			path.lnTo(x1 * pathW, y1 * pathH);

			return pathId;
		}

		if (this.cChartDrawer.nDimensionCount === 3) {
			var z;
			if (this.axisType === AscDFH.historyitem_type_ValAx) {
				z = this.cChartDrawer.processor3D.calculateZPositionValAxis();
			} else if (this.axisType === AscDFH.historyitem_type_CatAx) {
				z = this.cChartDrawer.processor3D.calculateZPositionCatAxis();
			}
			var convertResult = this.cChartDrawer._convertAndTurnPoint(x * this.chartProp.pxToMM, y * this.chartProp.pxToMM, z);
			x = convertResult.x / this.chartProp.pxToMM;
			y = convertResult.y / this.chartProp.pxToMM;

			convertResult = this.cChartDrawer._convertAndTurnPoint(x1 * this.chartProp.pxToMM, y1 * this.chartProp.pxToMM, z);
			x1 = convertResult.x / this.chartProp.pxToMM;
			y1 = convertResult.y / this.chartProp.pxToMM;

		}

		if (this.axis.axPos === window['AscFormat'].AX_POS_L || this.axis.axPos === window['AscFormat'].AX_POS_R) {
			path.moveTo(x1 * pathW, y1 * pathH);
			path.lnTo(x * pathW, y * pathH);
		} else {
			path.moveTo(x * pathW, y * pathH);
			path.lnTo(x1 * pathW, y1 * pathH);
		}

		return pathId;
	},

	_drawGridLines: function () {
		var pen;
		var path;
		if (!this.paths.gridLines) {
			return;
		}
		if(!this.axis.compiledMajorGridLines && !this.axis.compiledMinorGridLines) {
			return;
		}
		this.cChartDrawer.cShapeDrawer.bDrawSmartAttack = true;
		if (this.paths.minorGridLines) {
			path = this.paths.minorGridLines;
			pen = this.axis.compiledMinorGridLines;
			this.cChartDrawer.drawPath(path, pen);
		}
		if(this.paths.gridLines) {
			pen = this.axis.compiledMajorGridLines;
			path = this.paths.gridLines;
			this.cChartDrawer.drawPath(path, pen);
		}
		this.cChartDrawer.cShapeDrawer.bDrawSmartAttack = false;
	},

	_drawAxis: function () {
		let pen = this.axis.compiledLn;
		let path = this.paths.axisLine;
		if (path && path.length) {
			for (let i = 0; i < path.length; i++) {
				this.cChartDrawer.drawPath(path[i], pen);
			}
		} else {
			this.cChartDrawer.drawPath(path, pen);
		}
	},

	_drawTickMark: function () {
		var pen, path;
		if (this.axisType === AscDFH.historyitem_type_SerAx) {
			if (!this.paths.tickMarks) {
				return;
			}

			for (var i = 0; i < this.paths.tickMarks.length; i++) {
				pen = this.axis ? this.axis.compiledTickMarkLn : null;
				path = this.paths.tickMarks[i];
				this.cChartDrawer.drawPath(path, pen);
			}
		} else {
			if (this.paths.tickMarks) {
				this.cChartDrawer.cShapeDrawer.bDrawSmartAttack = true;
				pen = this.axis.compiledTickMarkLn;
				path = this.paths.tickMarks;
				this.cChartDrawer.drawPath(path, pen);
				this.cChartDrawer.cShapeDrawer.bDrawSmartAttack = false;
			}
		}
	}
};

	/** @constructor */
function floor3DChart()
{
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	
	this.paths = {};
}

floor3DChart.prototype =
{
    constructor: floor3DChart,
	
	draw : function(chartsDrawer)
    {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this._draw();
	},
	
	recalculate: function(chartsDrawer)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this.paths = {};
		this._calculate();
	},
	
	_calculate : function()
	{
		var nullPositionOy = this.chartProp.heightCanvas - this.chartProp.chartGutter._bottom;	
		var maxPositionOy = this.chartProp.chartGutter._top;
		var yPoints = this.cChartSpace.chart.plotArea.valAx ? this.cChartSpace.chart.plotArea.valAx.yPoints : null;
		if(yPoints && yPoints[0] && yPoints[yPoints.length - 1])
		{
			nullPositionOy = yPoints[0].pos > yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
			maxPositionOy = yPoints[0].pos < yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
		}
		
		var pxToMm = this.chartProp.pxToMM;
		var poisition = this.cChartDrawer.processor3D.calculateFloorPosition();
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		var point1, point2, point3, point4;
		
		switch(poisition)
		{
			case AscCommon.c_oChartFloorPosition.Left:
			{
				point1 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, 0);
				point2 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, maxPositionOy, 0);
				point3 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, maxPositionOy, perspectiveDepth);
				point4 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);

				break;
			}
			case AscCommon.c_oChartFloorPosition.Right:
			{
				point1 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, 0);
				point2 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, maxPositionOy, 0);
				point3 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, maxPositionOy, perspectiveDepth);
				point4 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, perspectiveDepth);
				
				break;
			}
			case AscCommon.c_oChartFloorPosition.Bottom:
			{
				point1 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, 0);
				point2 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);
				point3 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, perspectiveDepth);
				point4 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, 0);
			
				break;
			}
			case AscCommon.c_oChartFloorPosition.Top:
			{
				point1 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, maxPositionOy, 0);
				point2 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, maxPositionOy, perspectiveDepth);
				point3 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, maxPositionOy, perspectiveDepth);
				point4 = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, maxPositionOy, 0);
				
				break;
			}
		}
		
		if(point1)
		{
			this.paths.chartFloor = this.cChartDrawer.calculatePolygon([{x: point1.x / pxToMm, y: point1.y / pxToMm}, {x: point2.x / pxToMm, y: point2.y / pxToMm}, {x: point3.x / pxToMm, y: point3.y / pxToMm}, {x: point4.x / pxToMm, y: point4.y / pxToMm}]);
		}
	},
		
	_draw: function()
	{
		//TODO цвет заливки неправильно выставляется при чтении. поэтому использую пока цвет сетки
		var brush = this.cChartSpace.chart.floor ? this.cChartSpace.chart.floor.brush : null;
		var pen = this.cChartSpace.chart.floor ? this.cChartSpace.chart.floor.pen : null;
		var path = this.paths.chartFloor;
		
		this.cChartDrawer.drawPath(path, pen, brush);
	}

};

	/** @constructor */
function sideWall3DChart()
{
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	
	this.paths = {};
}

sideWall3DChart.prototype =
{
	constructor: sideWall3DChart,
	
	draw : function(chartsDrawer)
    {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this._draw();
	},
	
	recalculate: function(chartsDrawer)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this.paths = {};
		this._calculate();
	},
	
	_calculate : function()
	{
		var nullPositionOy = this.chartProp.heightCanvas - this.chartProp.chartGutter._bottom;	
		var maxPositionOy = this.chartProp.chartGutter._top;
		var yPoints = this.cChartSpace.chart.plotArea.valAx ? this.cChartSpace.chart.plotArea.valAx.yPoints : null;
		if(yPoints && yPoints[0] && yPoints[yPoints.length - 1])
		{
			nullPositionOy = yPoints[0].pos > yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
			maxPositionOy = yPoints[0].pos < yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
		}
		
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		var convertResult, x1n, y1n, x2n, y2n, x3n, y3n, x4n, y4n;
		if(this.chartProp.type === c_oChartTypes.HBar)
		{
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, 0);
			x1n = convertResult.x / this.chartProp.pxToMM;
			y1n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);
			x2n = convertResult.x / this.chartProp.pxToMM;
			y2n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, perspectiveDepth);
			x3n = convertResult.x / this.chartProp.pxToMM;
			y3n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, 0);
			x4n = convertResult.x / this.chartProp.pxToMM;
			y4n = convertResult.y / this.chartProp.pxToMM;
		}
		else
		{
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, 0);
			x1n = convertResult.x / this.chartProp.pxToMM;
			y1n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, this.chartProp.chartGutter._top, 0);
			x2n = convertResult.x / this.chartProp.pxToMM;
			y2n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, this.chartProp.chartGutter._top, perspectiveDepth);
			x3n = convertResult.x / this.chartProp.pxToMM;
			y3n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);
			x4n = convertResult.x / this.chartProp.pxToMM;
			y4n = convertResult.y / this.chartProp.pxToMM;
		}
		
		this.paths = this.cChartDrawer.calculatePolygon([{x: x1n, y: y1n}, {x: x2n, y: y2n}, {x: x3n, y: y3n}, {x: x4n, y: y4n}]);
	},
		
	_draw: function()
	{
		//TODO цвет заливки неправильно выставляется при чтении. поэтому использую пока цвет сетки
		var brush = this.cChartSpace.chart.sideWall ? this.cChartSpace.chart.sideWall.brush : null;
		var pen = this.cChartSpace.chart.sideWall ? this.cChartSpace.chart.sideWall.pen : null;
		var path = this.paths;
		
		this.cChartDrawer.drawPath(path, pen, brush);
	}
};

	/** @constructor */
function backWall3DChart()
{
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	
	this.paths = {};
}

backWall3DChart.prototype =
{
	constructor: backWall3DChart,
	
	draw : function(chartsDrawer)
    {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this._draw();
	},
	
	recalculate: function(chartsDrawer)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this.paths = {};
		this._calculate();
	},
	
	_calculate : function()
	{
		var nullPositionOy = this.chartProp.heightCanvas - this.chartProp.chartGutter._bottom;	
		var maxPositionOy = this.chartProp.chartGutter._top;
		var yPoints = this.cChartSpace.chart.plotArea.valAx ? this.cChartSpace.chart.plotArea.valAx.yPoints : null;
		if(yPoints && yPoints[0] && yPoints[yPoints.length - 1])
		{
			nullPositionOy = yPoints[0].pos > yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
			maxPositionOy = yPoints[0].pos < yPoints[yPoints.length - 1].pos ? yPoints[0].pos * this.chartProp.pxToMM : yPoints[yPoints.length - 1].pos * this.chartProp.pxToMM;
		}
		
		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		var convertResult, x1n, y1n, x2n, y2n, x3n, y3n, x4n, y4n;
		if(this.chartProp.type === c_oChartTypes.HBar)
		{
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);
			x1n = convertResult.x / this.chartProp.pxToMM;
			y1n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, this.chartProp.chartGutter._top, perspectiveDepth);
			x2n = convertResult.x / this.chartProp.pxToMM;
			y2n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, this.chartProp.chartGutter._top, perspectiveDepth);
			x3n = convertResult.x / this.chartProp.pxToMM;
			y3n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, perspectiveDepth);
			x4n = convertResult.x / this.chartProp.pxToMM;
			y4n = convertResult.y / this.chartProp.pxToMM;
		}
		else
		{
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, nullPositionOy, perspectiveDepth);
			x1n = convertResult.x / this.chartProp.pxToMM;
			y1n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.chartGutter._left, this.chartProp.chartGutter._top, perspectiveDepth);
			x2n = convertResult.x / this.chartProp.pxToMM;
			y2n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, this.chartProp.chartGutter._top, perspectiveDepth);
			x3n = convertResult.x / this.chartProp.pxToMM;
			y3n = convertResult.y / this.chartProp.pxToMM;
			convertResult = this.cChartDrawer._convertAndTurnPoint(this.chartProp.widthCanvas - this.chartProp.chartGutter._right, nullPositionOy, perspectiveDepth);
			x4n = convertResult.x / this.chartProp.pxToMM;
			y4n = convertResult.y / this.chartProp.pxToMM;
		}
		
		this.paths = this.cChartDrawer.calculatePolygon([{x: x1n, y: y1n}, {x: x2n, y: y2n}, {x: x3n, y: y3n}, {x: x4n, y: y4n}]);
	},
		
	_draw: function()
	{
		//TODO цвет заливки неправильно выставляется при чтении. поэтому использую пока цвет сетки
		var brush = this.cChartSpace.chart.backWall ? this.cChartSpace.chart.backWall.brush : null;
		var pen = this.cChartSpace.chart.backWall ? this.cChartSpace.chart.backWall.pen : null;
		var path = this.paths;
		
		this.cChartDrawer.drawPath(path, pen, brush);
	}
};


	/** @constructor */
function areaChart()
{
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	
	this.paths = null;
}

areaChart.prototype =
{
    constructor: areaChart,
	
	draw : function(chartsDrawer)
    {
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this._drawArea();
	},
	
	recalculate: function(chartsDrawer)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this.paths = null;
		this._calculateArea();
	},
	
	_calculateArea: function()
	{
        var pathId = this.cChartSpace.AllocPath();
        var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		var pxToMm = this.chartProp.pxToMM;
		
		path.moveTo(0, 0);
		path.lnTo(0 / pxToMm * pathW, this.chartProp.heightCanvas / pxToMm * pathH);
		path.lnTo(this.chartProp.widthCanvas / pxToMm * pathW, this.chartProp.heightCanvas / pxToMm * pathH);
		path.lnTo(this.chartProp.widthCanvas / pxToMm * pathW, 0 / pxToMm * pathH);
		path.lnTo(0, 0);

		this.paths = pathId;
	},
	
	_drawArea: function()
	{
		const pen = this.cChartSpace.pen;
		const brush = this.cChartSpace.brush;
		const oShapeDrawer = this.cChartDrawer.cShapeDrawer;
		let bNoSmartAttackVal;
		if(oShapeDrawer)
		{
			bNoSmartAttackVal = oShapeDrawer.bIsNoSmartAttack;
			oShapeDrawer.bIsNoSmartAttack = true;
		}
		this.cChartDrawer.drawPath(this.paths, pen, brush);
		if(oShapeDrawer)
		{
			oShapeDrawer.bIsNoSmartAttack = bNoSmartAttackVal;
		}
	}
};

	/** @constructor */
function plotAreaChart()
{
	this.chartProp = null;
	this.cChartSpace = null;
	this.cChartDrawer = null;
	
	this.paths = null;
}

plotAreaChart.prototype =
{
    constructor: plotAreaChart,
	
	draw: function(chartsDrawer, ignorePen, ignoreBrush)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this._drawArea(ignorePen, ignoreBrush);
	},
	
	recalculate: function(chartsDrawer)
	{
		this.chartProp = chartsDrawer.calcProp;
		this.cChartSpace = chartsDrawer.cChartSpace;
		this.cChartDrawer = chartsDrawer;
		
		this.paths = null;
		
		if(this.cChartDrawer.nDimensionCount === 3 && this.chartProp.type !== c_oChartTypes.Pie)
			this._calculateArea3D();
		else
			this._calculateArea();
	},
	
	_calculateArea: function()
	{
        var pathId = this.cChartSpace.AllocPath();
        var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;

		//смещаем на px все точки, посольку рисуем прямоугольную область
		var px = 1/this.chartProp.pxToMM;
		var plotAreaPoints = this.cChartDrawer.getPlotAreaPoints();
		var left = plotAreaPoints.left - px;
		var right = plotAreaPoints.right;
		var top = plotAreaPoints.top - px;
		var bottom = plotAreaPoints.bottom;
		
		path.moveTo(left * pathW, bottom * pathH);
		path.lnTo(right * pathW, bottom * pathH);
		path.lnTo(right * pathW, top * pathH);
		path.lnTo(left * pathW, top * pathH);
		path.lnTo(left * pathW, bottom * pathH);

		this.paths = pathId;
	},
	
	_calculateArea3D: function()
	{
        var pathId = this.cChartSpace.AllocPath();
        var path  = this.cChartSpace.GetPath(pathId);
		
		var pathH = this.chartProp.pathH;
		var pathW = this.chartProp.pathW;
		
		var pxToMm = this.chartProp.pxToMM;

		var plotAreaPoints = this.cChartDrawer.getPlotAreaPoints();
		var left = plotAreaPoints.left * pxToMm - 1;
		var right = plotAreaPoints.right * pxToMm - 1;
		var top = plotAreaPoints.top * pxToMm - 1;
		var bottom = plotAreaPoints.bottom * pxToMm - 1;

		var perspectiveDepth = this.cChartDrawer.processor3D.depthPerspective;
		
		var convertResult = this.cChartDrawer._convertAndTurnPoint(left, bottom, perspectiveDepth);
		var x1n = convertResult.x;
		var y1n = convertResult.y;
		convertResult = this.cChartDrawer._convertAndTurnPoint(right, bottom, perspectiveDepth);
		var x2n = convertResult.x;
		var y2n = convertResult.y;
		convertResult = this.cChartDrawer._convertAndTurnPoint(right, top, perspectiveDepth);
		var x3n = convertResult.x;
		var y3n = convertResult.y;
		convertResult = this.cChartDrawer._convertAndTurnPoint(left, top, perspectiveDepth);
		var x4n = convertResult.x;
		var y4n = convertResult.y;
		convertResult = this.cChartDrawer._convertAndTurnPoint(left, bottom, perspectiveDepth);
		var x5n = convertResult.x;
		var y5n = convertResult.y;

		
		path.moveTo(x1n / pxToMm * pathW, y1n / pxToMm * pathH);
		path.lnTo(x2n  / pxToMm * pathW, y2n / pxToMm * pathH);
		path.lnTo(x3n / pxToMm * pathW, y3n / pxToMm * pathH);
		path.lnTo(x4n / pxToMm * pathW, y4n / pxToMm * pathH);
		path.moveTo(x5n / pxToMm * pathW, y5n / pxToMm * pathH);

		this.paths = pathId;
	},
	
	_drawArea: function(ignorePen, ignoreBrush)
	{
		var pen = ignorePen ? null : this.cChartSpace.chart.plotArea.pen;
		var brush = ignoreBrush ? null : this.cChartSpace.chart.plotArea.brush;
		this.cChartDrawer.drawPath(this.paths, pen, brush);
	}
};

function CErrBarsDraw(chartsDrawer) {
	this.cChartDrawer = chartsDrawer;
	this.points = null;
	this.paths = [];
}

CErrBarsDraw.prototype = {
	constructor: CErrBarsDraw,

	draw: function () {
		var diffPen = 0;
		var leftRect = this.cChartDrawer.calcProp.chartGutter._left / this.cChartDrawer.calcProp.pxToMM;
		var topRect = (this.cChartDrawer.calcProp.chartGutter._top - diffPen) / this.cChartDrawer.calcProp.pxToMM;
		var rightRect = this.cChartDrawer.calcProp.trueWidth / this.cChartDrawer.calcProp.pxToMM;
		var bottomRect = (this.cChartDrawer.calcProp.trueHeight + diffPen) / this.cChartDrawer.calcProp.pxToMM;

		this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
		this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);

		for (let chartId in this.paths) {
			if (this.paths.hasOwnProperty(chartId) && this.paths[chartId]) {
				for (let serId in this.paths[chartId]) {
					if (this.paths[chartId].hasOwnProperty(serId) && this.paths[chartId][serId]) {
						for (let errBarIdx in this.paths[chartId][serId]) {
							if (this.paths[chartId][serId].hasOwnProperty(errBarIdx) && this.paths[chartId][serId][errBarIdx]) {
								const seria = this.charts[chartId].chart.series[serId];
								for (let i = 0; i < this.paths[chartId][serId][errBarIdx].length; i++) {
									const errBar = seria.getErrBarById(errBarIdx);
									const pen = errBar && errBar.getPen();
									if (pen) {
										this.cChartDrawer.drawPath(this.paths[chartId][serId][errBarIdx][i], pen);
									}
								}
							}
						}
					}
				}
			}
		}

		this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
	},

	recalculate: function (charts) {
		this.paths = {};
		this.charts = charts;
		for (var i in charts) {
			if (charts.hasOwnProperty(i) && charts[i]) {
				this.paths[i] = this.recalculateChart(charts[i]);
			}
		}
	},

	putPoint: function (x, y, xVal, yVal, nSeries, nPoint) {
		if (!this.points) {
			this.points = [];
		}
		if (!this.points[nSeries]) {
			this.points[nSeries] = [];
		}

		this.points[nSeries][nPoint] = {x: x, y: y, xVal: xVal, yVal: yVal};
	},

	recalculateChart: function (oChart) {
		if (!this.points) {
			return;
		}

		var res = [];
		var t = this;
		var pxToMm = this.cChartDrawer.calcProp.pxToMM;
		var calcErrLine = function (errBar, _start, _end, pos) {
			var pathId = t.cChartDrawer.cChartSpace.AllocPath();
			var path = t.cChartDrawer.cChartSpace.GetPath(pathId);

			var pathH = t.cChartDrawer.calcProp.pathH;
			var pathW = t.cChartDrawer.calcProp.pathW;

			if (t.cChartDrawer.getErrBarsPosition(errBar, oChart)) {
				path.moveTo(_start * pathW, pos * pathH);
				path.lnTo(_end * pathW, pos * pathH);
			} else {
				path.moveTo(pos * pathW, _start * pathH);
				path.lnTo(pos * pathW, _end * pathH);
			}

			if (!errBar.noEndCap) {
				var sizeCap = 5 / pxToMm;

				if (isHorPos) {
					path.moveTo(_start * pathW, (pos - sizeCap / 2) * pathH);
					path.lnTo(_start * pathW, (pos + sizeCap / 2) * pathH);

					path.moveTo(_end * pathW, (pos - sizeCap / 2) * pathH);
					path.lnTo(_end * pathW, (pos + sizeCap / 2) * pathH);
				} else {
					path.moveTo((pos - sizeCap / 2) * pathW, _start * pathH);
					path.lnTo((pos + sizeCap / 2) * pathW, _start * pathH);

					path.moveTo((pos - sizeCap / 2) * pathW, _end * pathH);
					path.lnTo((pos + sizeCap / 2) * pathW, _end * pathH);
				}
			}

			return pathId;
		};

		var errBars;
		var isHorPos;
		for (var i = 0; i < this.points.length; i++) {
			if (this.points[i]) {
				var serIndex = t.cChartDrawer._getIndexByIdxSeria(oChart.chart.series, i);
				if (!oChart.chart.series[serIndex] || !oChart.chart.series[serIndex].errBars[0]) {
					continue;
				}

				errBars = oChart.chart.series[serIndex].errBars;

				if (!errBars) {
					return;
				}

				for (let errBarIdx = 0; errBarIdx < errBars.length; errBarIdx ++) {
					const errBar = errBars[errBarIdx];
					for (var j = 0; j < this.points[i].length; j++) {
						if (this.points[i][j]) {

							isHorPos = t.cChartDrawer.getErrBarsPosition(errBar, oChart);
							var axis = this.cChartDrawer.getAxisByPos(oChart.chart.axId, isHorPos);
							// if categoryAxis in use, then we need other axis to calculate the position
							var otherAxis = this.cChartDrawer.getAxisByPos(oChart.chart.axId, !isHorPos);
							var isCatAx = axis.getObjectType() === AscDFH.historyitem_type_CatAx;

							// calculate the magnitude of error in one direction 
							var oErrVal = this.calculateErrVal(oChart.chart, serIndex, errBarIdx, j, isCatAx);
							if (oErrVal !== null) {
								var plusErrVal = oErrVal.plusErrVal;
								var minusErrVal = oErrVal.minusErrVal;
								var x = this.points[i][j].x;
								var y = this.points[i][j].y;

								//TODO, for now supports only lines
								var pointVal;
								if (null !== oErrVal.startVal) {
									pointVal = oErrVal.startVal;
								} else {
									switch (this.cChartDrawer._getChartType(oChart.chart)) {
										case c_oChartTypes.Scatter: {
											pointVal = isCatAx ? j + 1 : (isHorPos ? this.points[i][j].xVal : this.points[i][j].yVal);
											break;
										}
										default: {
											pointVal = isCatAx ? j + 1 : this.points[i][j].xVal;
											break;
										}
									}
								}

								var start, end;
								switch (errBar.errBarType) {
									case AscFormat.st_errbartypeBOTH: {
										start = this.cChartDrawer.getYPosition(pointVal - minusErrVal, axis, true);
										end = this.cChartDrawer.getYPosition(pointVal + plusErrVal, axis, true);
										break;
									}
									case AscFormat.st_errbartypePLUS: {
										start = this.cChartDrawer.getYPosition(pointVal, axis, true);
										end = this.cChartDrawer.getYPosition(pointVal + plusErrVal, axis, true);
										break;
									}
									case AscFormat.st_errbartypeMINUS: {
										start = this.cChartDrawer.getYPosition(pointVal - minusErrVal, axis, true);
										end = this.cChartDrawer.getYPosition(pointVal, axis, true);
										break;
									}
								}

								if (!res[serIndex]) {
									res[serIndex] = [];
								}

								if (!res[serIndex][errBar.Id]) {
									res[serIndex][errBar.Id] = [];
								}

								var fixPosition = x;
								var isCatOtherAxis = otherAxis.getObjectType() === AscDFH.historyitem_type_CatAx;
								if (isHorPos) {
									if (isCatOtherAxis && y) {
										fixPosition = y;
									} else {
										fixPosition = this.cChartDrawer.getYPosition(isCatOtherAxis ? j + 1 : this.points[i][j].yVal, otherAxis, true);
									}
								}
								res[serIndex][errBar.Id].push(calcErrLine(errBar, start, end, fixPosition));
							}
						}
					}
				}
			}
		}

		return res;
	},

	getPointPos: function (oChart, path) {
		var res = null;
		var type = oChart.chart.getObjectType();
		if (type === AscDFH.historyitem_type_LineChart) {
			var commandIndex = 0;
			var oPath = this.cChartDrawer.cChartSpace.GetPath(path);
			var oCommand0 = oPath.getCommandByIndex(commandIndex);
			res = {x: oCommand0.X, y: oCommand0.Y};
		}

		return res;
	},

	calculateErrVal: function (oChart, ser, errBarIdx, val, isCatAx) {
		var t = this;
		var seria = oChart.series[ser];

		var startVal = null;
		var plusErrVal = null;
		var minusErrVal = null;
		const errBars = seria && Array.isArray(seria.errBars) && errBarIdx < seria.errBars.length ? seria.errBars[errBarIdx] : null;

		if (seria && errBars) {
			const isHorErrBar = errBars.errDir === AscFormat.st_errdirX;
			var grouping = t.cChartDrawer.getChartGrouping(oChart);
			var chartType = t.cChartDrawer._getChartType(oChart);
			var pointVal = null;
			if (chartType === c_oChartTypes.Scatter) {
				pointVal = t.cChartDrawer.getValWithoutStacked(ser, val, oChart, isHorErrBar);
				if (!pointVal && isHorErrBar) {
					pointVal = val + 1;
				}
			} else {
				pointVal = errBars.errValType === AscFormat.st_errvaltypePERCENTAGE ? (grouping === "stackedPer" ? t.cChartDrawer.getValWithPercent(ser, val, oChart) : t.cChartDrawer.getValWithoutStacked(ser, val, oChart)) : t.cChartDrawer.getValWithStacked(ser, val, oChart);
			}

			if (pointVal == null) {
				return pointVal;
			}

			//TODO J
			var calculateSerSum = function (_ser, _startVal, endVal) {
				return t.cChartDrawer.getAutoSum(_ser, _ser, _startVal, endVal, function (_ser, _val) {
					const numPoint = seria ? t.cChartDrawer.getPointByIndex(seria, _val, true) : null;
					const currPointVal = numPoint ? numPoint.val : _val + 1;
					var _pointVal = (isCatAx || isHorErrBar) ? currPointVal : t.cChartDrawer.getValWithStacked(_ser, _val, oChart);
					return _pointVal != null ? _pointVal : 0;
				});
			};

			var calculateStDev = function () {
				var ny, mSumm, aSumm;
				ny = pointsCount;

				var startSer = ser;
				var endSer = ser;
				var startPoint = 0;
				var endPoint = pointsCount - 1;
				mSumm = calculateSerSum(startSer, startPoint, endPoint);
				var m = mSumm / ny;

				aSumm = t.cChartDrawer.getAutoSum(startSer, endSer, startPoint, endPoint, function (_ser, _val) {
					const numPoint = seria ? t.cChartDrawer.getPointByIndex(seria, _val, true) : null;
					const currPointVal = numPoint ? numPoint.val : _val + 1;
					var _pointVal = (isCatAx || isHorErrBar) ? currPointVal : t.cChartDrawer.getValWithStacked(_ser, _val, oChart);
					if (_pointVal == null) {
						return 0;
					}
					var av = _pointVal - m;
					return av * av;
				});

				return Math.sqrt(aSumm / (ny - 1));
			};

			var pointsCount = seria.getValuesCount();
			if (!pointsCount) {
				if (t.cChartDrawer.getErrBarsPosition(errBars, oChart)) {
					if (seria.xVal) {
						pointsCount = seria.xVal.getValuesCount ?  seria.xVal.getValuesCount() : (seria.xVal.getValues ? seria.xVal.getValues().length : 0);
					} else if (seria.yVal) {
						pointsCount = seria.yVal.getValuesCount ? seria.yVal.getValuesCount(): (seria.yVal.getValues ? seria.yVal.getValues().length : 0);
					}
				} else {
					if (seria.yVal) {
						pointsCount = seria.yVal.getValuesCount ? seria.yVal.getValuesCount(): (seria.yVal.getValues ? seria.yVal.getValues().length : 0);
					} else if (seria.xVal) {
						pointsCount = seria.xVal.getValuesCount ?  seria.xVal.getValuesCount() : (seria.xVal.getValues ? seria.xVal.getValues().length : 0);
					}
				}
			}
			switch (errBars.errValType) {
				case AscFormat.st_errvaltypeCUST: {
					//TODO numRef ?
					var numLit = errBars.plus && (errBars.plus.numLit || (errBars.plus.numRef && errBars.plus.numRef.numCache));
					if (errBars.plus && numLit) {
						plusErrVal = numLit.getPtByIndex(numLit.ptCount === 1 ? 0 : val);
						if (plusErrVal) {
							plusErrVal = plusErrVal.val;
						}
					}
					numLit = errBars.minus && (errBars.minus.numLit || (errBars.minus.numRef && errBars.minus.numRef.numCache));
					if (errBars.minus && numLit) {
						minusErrVal = numLit.getPtByIndex(numLit.ptCount === 1 ? 0 : val);
						if (minusErrVal) {
							minusErrVal = minusErrVal.val;
						}
					}
					break;
				}
				case AscFormat.st_errvaltypeFIXEDVAL: {
					plusErrVal = errBars.val;
					break;
				}
				case AscFormat.st_errvaltypePERCENTAGE: {
					plusErrVal = pointVal * errBars.val / 100;
					// this number is used to calculate an axis, in this case the magnitude of error is calculated from the inital value
					// but the inital value is - stacked value
					pointVal = t.cChartDrawer.getValWithStacked(ser, val, oChart);
					break;
				}
				case AscFormat.st_errvaltypeSTDDEV: {
					//calculate Standard Deviation +  change start val
					plusErrVal = calculateStDev();
					startVal = (calculateSerSum(ser, 0, pointsCount - 1)) / pointsCount;
					break;
				}
				case AscFormat.st_errvaltypeSTDERR: {
					//official excel information. formula for calculating the error
					//here the gormula is STDEV(currentSerData)/SQRT(COUNT(currentSerData))
					plusErrVal = calculateStDev() / Math.sqrt(pointsCount);
					break;
				}
			}
		}

		if (null == minusErrVal) {
			minusErrVal = plusErrVal;
		}

		return plusErrVal !== null ? {startVal: startVal, plusErrVal: Math.abs(plusErrVal), minusErrVal: Math.abs(minusErrVal), val: pointVal} : null;
	}
};

/** @constructor */
function CGeometry2()
{
    this.pathLst = [];
	this.isLine = false;
	this.gdLst = [];
}

CGeometry2.prototype =
{
    constructor: CGeometry2,
	
	canFill: function()
    {
        if(this.preset === "line")
            return false;
        for(var i = 0; i < this.pathLst.length; ++i)
        {
            if(this.pathLst[i].fill !== "none")
                return true;
        }
        return  false;
    },

    AddPath: function(path)
    {
        this.pathLst.push(path);
    },
	
    AddRect: function(l, t, r, b)
    {
        this.rectS = {};
        this.rectS.l = l;
        this.rectS.t = t;
        this.rectS.r = r;
        this.rectS.b = b;
    },

    draw: function(shape_drawer)
    {
        for (var i=0, n=this.pathLst.length; i<n;++i)
            this.pathLst[i].drawSmart(shape_drawer);
    },
	
	check_bounds: function(checker)
    {

        for(var i=0, n=this.pathLst.length; i<n;++i)

            this.pathLst[i].check_bounds(checker);

    }
};

	/** @constructor */
function CColorObj(pen, brush, geometry)
{
    this.pen = pen;
	this.brush = brush;
	this.geometry = geometry;
}

CColorObj.prototype =
{
	constructor: CColorObj,
	
	check_bounds: function (checker) {
		if (this.geometry) {
			this.geometry.check_bounds(checker);
		}
	}
};

	/** @constructor */
	function calcShapesHelper(chartsDrawer)
	{
		this.chartsDrawer = chartsDrawer;

		this.x = null;//startX
		this.y = null;//startY
		this.z = null;//gapDepth

		this.w = null;//individualBarWidth
		this.h = null;//height
		this.d = null;//perspectiveDepth

		this.onlyFrontFaces = null;
		this.val = null;
		this.rotated = null;//hbar

		this.dRadius1 = null;
		this.dRadius2 = null;
		//нужны только для конусов uRadius1/uRadius2
		this.uRadius1 = null;
		this.uRadius2 = null;
	}

	calcShapesHelper.prototype.setProps = function (x, y, z, w, h, d, onlyFrontFaces, val, rotated, dRadius1, dRadius2, uRadius1, uRadius2, cone) {
		this.x = x;//startX
		this.y = y;//startY
		this.z = z;//gapDepth

		this.w = w;//individualBarWidth
		this.h = h;//height
		this.d = d;//perspectiveDepth

		this.onlyFrontFaces = onlyFrontFaces;
		this.val = val;
		this.rotated = rotated;//hbar

		this.dRadius1 = dRadius1;
		this.dRadius2 = dRadius2;
		this.uRadius1 = uRadius1;
		this.uRadius2 = uRadius2;
	};

	calcShapesHelper.prototype.getCone = function () {
		return this._getCylinder(true);
	};

	calcShapesHelper.prototype.getCylinder = function () {
		return this._getCylinder();
	};

	calcShapesHelper.prototype._getCylinder = function (bCone) {
		var centerUpX, centerUpY, centerUpZ, centerDownX, centerDownY, centerDownZ;

		if (this.rotated) {
			centerDownX = this.x, centerDownY = this.y + this.w / 2, centerDownZ = 0 + this.z + this.d / 2;
			centerUpX = this.x + this.h, centerUpY = this.y + this.w / 2, centerUpZ = 0 + this.z + this.d / 2;
		} else {
			centerDownX = this.x + this.w / 2, centerDownY = this.y, centerDownZ = 0 + this.z + this.d / 2;
			centerUpX = this.x + this.w / 2, centerUpY = this.y - this.h, centerUpZ = 0 + this.z + this.d / 2;
		}

		var segmentPoints = [];
		var segmentPoints2 = [];
		var segmentPoint1, segmentPoint2;

		var A, B, A1, B1;
		// в ms 180 градусов составляют 17 сегментов
		// todo пока что рассчитываем дополнительные точки, нужно будет оптимизировать
		var dt = Math.PI / 34; //Math.PI / 17;

		//рассчитываем стартовый угол
		var angel = Math.abs(this.chartsDrawer.processor3D.angleOy);
		var k = Math.PI / 2;
		k += angel;

		// получаем точки основания цилиндра через парамметрические уравнения эллиптического цилиндра
		for (var t = k; t <= Math.PI * 2 + k; t += dt) {
			A = this.dRadius1 * Math.cos(t);
			B = this.dRadius2 * Math.sin(t);

			if (bCone) {
				A1 = this.uRadius1 * Math.cos(t);
				B1 = this.uRadius2 * Math.sin(t);
				if (this.rotated) {
					segmentPoint1 = this.chartsDrawer._convertAndTurnPoint(centerDownX, centerDownY + A, centerDownZ - B);
					segmentPoint2 = this.chartsDrawer._convertAndTurnPoint(centerUpX, centerUpY + A1, centerUpZ - B1);
				} else {
					segmentPoint1 = this.chartsDrawer._convertAndTurnPoint(centerDownX + A, centerDownY, centerDownZ + B);
					segmentPoint2 = this.chartsDrawer._convertAndTurnPoint(centerUpX + A1, centerUpY, centerUpZ + B1);
				}
			} else {
				if (this.rotated) {
					segmentPoint1 = this.chartsDrawer._convertAndTurnPoint(centerDownX, centerDownY + A, centerDownZ - B);
					segmentPoint2 = this.chartsDrawer._convertAndTurnPoint(centerUpX, centerUpY + A, centerUpZ - B);
				} else {
					segmentPoint1 = this.chartsDrawer._convertAndTurnPoint(centerDownX + A, centerDownY, centerDownZ + B);
					segmentPoint2 = this.chartsDrawer._convertAndTurnPoint(centerUpX + A, centerUpY, centerUpZ + B);
				}
			}

			segmentPoints.push(segmentPoint1);
			segmentPoints2.push(segmentPoint2);
		}

		var sortCylinderPoints1 = [];
		var sortCylinderPoints2 = [];
		var invisible = false;

		// сортируем точки по видимости для построения плоскости цилиндра
		for (var i = 1; i < segmentPoints.length; i++) {
			if (this.chartsDrawer._isVisibleVerge3D(segmentPoints[i], segmentPoints[i - 1], segmentPoints2[i - 1], this.val, true)) {
				if (!invisible) {
					sortCylinderPoints1.push(segmentPoints[i - 1]);
					sortCylinderPoints2.push(segmentPoints2[i - 1]);
				}
				else {
					break;
				}
			} else {
				invisible = true;
			}
		}
		if (invisible) {
			for (var k = segmentPoints.length - 1; i <= k; k--) {
				if (this.chartsDrawer._isVisibleVerge3D(segmentPoints[k], segmentPoints[k - 1], segmentPoints2[k - 1], this.val, true)) {
					sortCylinderPoints1.unshift(segmentPoints[k - 1]);
					sortCylinderPoints2.unshift(segmentPoints2[k - 1]);
				}
			}
		}
		var isNotAllPointsVisible = invisible;

		// проверяем если все точки поверхности цилиндра(конуса) либо видимы либо невидимы
		// если уловие выполняется, то для отрисовки цилиндра(конуса) достаточно отрисовать эллипс (т.е вид сверху или снизу)
		if (sortCylinderPoints1.length === 0 || sortCylinderPoints2.length === 0) {
			sortCylinderPoints1 = segmentPoints;
			sortCylinderPoints2 = segmentPoints2;
			isNotAllPointsVisible = false;
		}

		var x12, y12, z12, x22, y22, z22, x32, y32, z32, x42, y42, z42, x52, y52, z52, x62, y62, z62, x72, y72, z72, x82, y82, z82;
		var point1, point2, point4, point5, point6, point8;

		if (this.rotated) {
			x12 = this.x, y12 = this.y, z12 = 0 + this.z;
			x42 = this.x + this.h, y42 = this.y, z42 = this.d + this.z;
			x52 = this.x, y52 = this.y + this.w, z52 = 0 + this.z;
			x62 = this.x, y62 = this.y + this.w, z62 = this.d + this.z;
			x72 = this.x + this.h, y72 = this.y + this.w, z72 =  this.d + this.z;
			x82 = this.x + this.h, y82 = this.y + this.w, z82 = 0 + this.z;

			point1 = this.chartsDrawer._convertAndTurnPoint(x12, y12, z12);
			point2 = this.chartsDrawer._convertAndTurnPoint(x52, y52, z52);
			point4 = this.chartsDrawer._convertAndTurnPoint(x62, y62, z62);
			point5 = this.chartsDrawer._convertAndTurnPoint(x72, y72, z72);
			point6 = this.chartsDrawer._convertAndTurnPoint(x42, y42, z42);
			point8 = this.chartsDrawer._convertAndTurnPoint(x82, y82, z82);
		} else {
			x12 = this.x, y12 = this.y, z12 = 0 + this.z;

			x22 = this.x, y22 = this.y, z22 = this.d + this.z;
			x42 = this.x + this.w, y42 = this.y, z42 = 0 + this.z;
			x52 = this.x, y52 = this.y - this.h, z52 = 0 + this.z;
			x62 = this.x, y62 = this.y - this.h, z62 = this.d + this.z;
			x82 = this.x + this.w, y82 = this.y - this.h, z82 = 0 + this.z;

			point1 = this.chartsDrawer._convertAndTurnPoint(x12, y12, z12);
			point2 = this.chartsDrawer._convertAndTurnPoint(x22, y22, z22);
			point4 = this.chartsDrawer._convertAndTurnPoint(x42, y42, z42);
			point5 = this.chartsDrawer._convertAndTurnPoint(x52, y52, z52);
			point6 = this.chartsDrawer._convertAndTurnPoint(x62, y62, z62);
			point8 = this.chartsDrawer._convertAndTurnPoint(x82, y82, z82);
		}

		var points = [segmentPoints, segmentPoints2, point1, point2, point4, point5, point6, point8,
			sortCylinderPoints1, sortCylinderPoints2];

		return this.calculateCylinder(points, false, isNotAllPointsVisible, bCone);
	};

	calcShapesHelper.prototype.calculateCylinder = function (points, notDraw, isNotAllPointsVisible, cone) {
		var res;
		var segmentPoints = points[0];
		var segmentPoints2 = points[1];
		var point1 = points[2];
		var point2 = points[3];
		var point4 = points[4];
		var point5 = points[5];
		var point6 = points[6];
		var point8 = points[7];
		var sortCylinderPoints1 = points[8];
		var sortCylinderPoints2 = points[9];

		var frontPaths = [];
		var darkPaths = [];

		// условие для конусов
		var isVisibleReverse = cone ? isNotAllPointsVisible : true;

		var addPathToArr = function (isFront, face, index) {
			frontPaths[index] = null;
			darkPaths[index] = null;

			if (isFront) {
				frontPaths[index] = face;
			} else {
				darkPaths[index] = face;
			}
		};

		var face, faceFront;

		//front
		faceFront = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, false, true, isNotAllPointsVisible);
		addPathToArr(this._isVisibleVerge3D(sortCylinderPoints1[0], sortCylinderPoints1[sortCylinderPoints1.length - 1],
			sortCylinderPoints2[0], this.val, !isVisibleReverse), faceFront, 0);

		//down
		if (this.val === 0) {
			face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
			addPathToArr(true, face, 1);
		} else {
			if (cone) {
				if (this._isVisibleVerge3D(point4, point1, point2, this.val)) {
					face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
					addPathToArr(true, face, 1);
				} else {
					face = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, true, true, true);
					addPathToArr(true, face, 1);
				}
			} else {
				face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, false, true, true);
				addPathToArr(this._isVisibleVerge3D(point4, point1, point2, this.val), face, 1);
			}
		}

		//up
		if (!notDraw) {
			if (this.val === 0) {
				face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
				addPathToArr(true, face, 4);
			} else {
				if (cone) {
					if (this._isVisibleVerge3D(point6, point5, point8, this.val)) {
						face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
						addPathToArr(true, face, 4);
					} else {
						face = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, true, false, true, true);
						addPathToArr(true, face, 4);
					}
				} else {
					face = this._calculatePathFaceCylinder(segmentPoints, segmentPoints2, true, false, true);
					addPathToArr(this._isVisibleVerge3D(point6, point5, point8, this.val), face, 4);
				}
			}
		}

		//unfront
		faceFront = this._calculatePathFaceCylinder(sortCylinderPoints1, sortCylinderPoints2, false, false, true, isNotAllPointsVisible);
		addPathToArr(this._isVisibleVerge3D(sortCylinderPoints1[0], sortCylinderPoints1[sortCylinderPoints1.length - 1],
			sortCylinderPoints2[0], this.val, isVisibleReverse), faceFront, 5);

		if (this.onlyFrontFaces) {
			res = frontPaths;
		} else {
			res = { frontPaths: frontPaths, darkPaths: darkPaths };
		}

		return res;
	};

	calcShapesHelper.prototype._calculatePathFaceCylinder = function(segmentPoints, segmentPoints2, up, down, isConvertPxToMM, check)
	{
		var pxToMm = 1;
		if(isConvertPxToMM)
		{
			pxToMm = this.chartsDrawer.calcProp.pxToMM;
		}

		var pathId = this.chartsDrawer.cChartSpace.AllocPath();
		var path  = this.chartsDrawer.cChartSpace.GetPath(pathId);

		var pathH = this.chartsDrawer.calcProp.pathH;
		var pathW = this.chartsDrawer.calcProp.pathW;

		if (up) {
			for (var i = 0; i < segmentPoints2.length; i++) {
				if (i % 2 === 0) {
					path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
				}
			}
			if (check) {
				for (var i = segmentPoints2.length - 1; i >= 0; i--) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
					}
				}
			} else {
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH)
			}
		} else if (down) {
			for (var i = 0; i < segmentPoints.length; i++) {
				if (i % 2 === 0) {
					path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
				}
			}
			if (check) {
				for (var i = segmentPoints.length - 1; i >= 0; i--) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
					}
				}
			} else {
				path.lnTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)
			}
		} else {
			var endIndex = 0;
			var startIndex = segmentPoints.length - 1;

			if (!check) {
				path.moveTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH)
				for (var i = 0; i < segmentPoints2.length; i++) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints2[i].x / pxToMm * pathW, segmentPoints2[i].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

				path.moveTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)
				for (var i = 0; i < segmentPoints.length; i++) {
					if (i % 2 === 0) {
						path.lnTo(segmentPoints[i].x / pxToMm * pathW, segmentPoints[i].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints[0].x / pxToMm * pathW, segmentPoints[0].y / pxToMm * pathH)

			} else if (segmentPoints2.length === 1) {
				path.lnTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH);
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

				for (var k = endIndex; k <= startIndex; k++) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints[k].x / pxToMm * pathW, segmentPoints[k].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints2[0].x / pxToMm * pathW, segmentPoints2[0].y / pxToMm * pathH);

			} else {

				path.moveTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH)

				for (var k = endIndex; k <= startIndex; k++) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints[k].x / pxToMm * pathW, segmentPoints[k].y / pxToMm * pathH);
					}
				}

				for (k = startIndex; endIndex <= k; k--) {
					if (k % 2 === 0) {
						path.lnTo(segmentPoints2[k].x / pxToMm * pathW, segmentPoints2[k].y / pxToMm * pathH);
					}
				}
				path.lnTo(segmentPoints[endIndex].x / pxToMm * pathW, segmentPoints[endIndex].y / pxToMm * pathH);
			}

		}

		return pathId;
	};

	calcShapesHelper.prototype._isVisibleVerge3D = function(k, n, m, val, cylinder)
	{
		//roberts method - calculate normal
		var aX = n.x * m.y - m.x * n.y;
		var bY = - (k.x * m.y - m.x * k.y);
		var cZ = k.x * n.y - n.x * k.y;
		var visible = aX + bY + cZ;

		var result;
		if(this.chartsDrawer.calcProp.type === c_oChartTypes.Bar)
		{
			result = val > 0 && visible < 0 || val < 0 && visible > 0;
			if(!(this.chartsDrawer.calcProp.subType === "stacked") && !(this.chartsDrawer.calcProp.subType === "stackedPer") && this.chartsDrawer.cChartSpace.chart.plotArea.valAx.isReversed())
				result = !result;
		}
		else if(this.chartsDrawer.calcProp.type === c_oChartTypes.Line)
		{
			result = visible < 0;
		}
		else if(this.chartsDrawer.calcProp.type === c_oChartTypes.HBar)
		{
			result = val < 0 && visible < 0 || val > 0 && visible > 0;

			if(!(this.chartsDrawer.calcProp.subType === "stacked") && !(this.chartsDrawer.calcProp.subType === "stackedPer") && this.chartsDrawer.cChartSpace.chart.plotArea.valAx.isReversed())
				result = !result;
		}
		if (cylinder) {
			result = !result;
		}

		return result;
	};

	calcShapesHelper.prototype.isIntersectionLineAndLine2d = function (x1, y1, x2, y2, x3, y3, x4, y4) {

		if ((x1 === x2 && y1 === y2) || (x3 === x4 && y3 === y4)) {
			return false;
		}

		var k = ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1));

		if (k === 0) {
			return false;
		}

		var ua = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / k;
		var ub = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / k;

		if (ua < 0 || ua > 1 || ub < 0 || ub > 1) {
			return false;
		}

		var x = x1 + ua * (x2 - x1);
		var y = y1 + ua * (y2 - y1);

		return { x: x, y: y };
	};

	function CTrendline(chartsDrawer) {
		this.cChartDrawer = chartsDrawer;
		this.storage = {};//[chartId][seriaId]
		this.ptCount = 0;

		//control trend calculate type
		this.bAllowDrawByBezier = true;
		this.bAllowDrawByPoints = false;
		this.stopAdding = false;
	}

	CTrendline.prototype = {

		constructor: CTrendline,

		addCoordinate: function (xVal, yVal, chartId, seriaId) {
			if (!this.storage[chartId]) {
				this.storage[chartId] = {};
			}
			if (!this.storage[chartId][seriaId]) {
				this.storage[chartId][seriaId] = new CTrendData();
			}
			if (!this.storage[chartId][seriaId].isEmpty() && xVal === this.storage[chartId][seriaId].coords.catVals[0]) {
				this.stopAdding = true;
			}

			// in the case of duplicated data, no further adding should be allowed
			if (!this.stopAdding) {
				this.ptCount+=1;
				this.storage[chartId][seriaId].addCatVal(xVal);
				this.storage[chartId][seriaId].addValVal(yVal);
	
				if (yVal > 0) {
					this.storage[chartId][seriaId].setMinLogVal(yVal);
				}
			}
		},

		getAdditionalInfo: function (chartId, seriesId) {
			if (this.storage[chartId] && this.storage[chartId][seriesId]) {
				const coefficients = this.storage[chartId][seriesId].getCoefficients();
				const rSquared = this.storage[chartId][seriesId].getRSquared();
				const lastPoint = this.storage[chartId][seriesId].getLastPoint();

				let startCat = this.cChartDrawer.calcProp.chartGutter._left / this.cChartDrawer.calcProp.pxToMM;
				let startVal = (this.cChartDrawer.calcProp.chartGutter._top) / this.cChartDrawer.calcProp.pxToMM;
				let endCat = (this.cChartDrawer.calcProp.trueWidth / this.cChartDrawer.calcProp.pxToMM) + startCat;
				let endVal = (this.cChartDrawer.calcProp.trueHeight / this.cChartDrawer.calcProp.pxToMM) + startVal;
				
				if (lastPoint.catVal < startCat || lastPoint.catVal > endCat) {
					lastPoint.catVal = 0;
				}

				if (lastPoint.valVal < startVal || lastPoint.valVal > endVal) {
					lastPoint.valVal = 0;
				}

				if (coefficients || rSquared) {
					const additionalInfo = {
						coefficients : coefficients,
						rSquared : rSquared,
						coordinate : lastPoint
					}
					return additionalInfo;
				}
			}
			return null;
		},

		preCalculate: function (charts, boundaries) {
			for (let i in charts) {
				const index1 = charts[i].Id;
				if (charts.hasOwnProperty(i) && this.storage[index1]) {
					for (let j in charts[i].series) {
						const index2 = charts[i].series[j].Id;
						if (charts[i].series.hasOwnProperty(j) && this.storage[index1][index2]) {
							this._calculateLine(charts[i].series[j].parent, this.storage[index1][index2], charts[i].series[j].trendline);
							this._setNewBoundaries(charts[i].series[j].parent, this.storage[index1][index2], boundaries);
						}
					}
				}
			}
		},

		_calculateLine: function (oChart, storageElement, attributes) {
			const type = attributes.trendlineType;
			const catAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[0]);
			const valAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[1]);
			const coords = storageElement.getCoords();
			// moving average differs much from other trendlines
			if (type === AscFormat.TRENDLINE_TYPE_MOVING_AVG) {
				const period = attributes.period ? attributes.period : 2;
				const pointVals = this._getMAline(coords.catVals, coords.valVals, this.ptCount, period, valAxis.scaling.logBase);
				storageElement.setPointVals(pointVals);
			} else {
				let order = attributes.order ? attributes.order + 1 : 2;
				// calculate constants of predicted equation and get its possible mappings such as calcYVal, calcXval, calcSlope
				const coefficients = this._getEquationCoefficients(coords.catVals, coords.valVals, type, order, attributes.intercept);
				const equationStorage = this._obtainEquationStorage(type);
				if (coefficients && equationStorage) {
					let catMax = catAxis.max;
					let catMin = catAxis.min;
					const midPointsNum = 100;
					const lineBuilder = new CLineBuilder(coefficients, catMin, catMax, valAxis.scaling.min, valAxis.scaling.max, valAxis.scaling.logBase);
					lineBuilder.setCalcYVal(equationStorage.calcYVal);
					let bezierFailed = false;
					//poly lines can not be drawn with approximated bezier
					if (this.bAllowDrawByBezier && type !== AscFormat.TRENDLINE_TYPE_POLY) {
						lineBuilder.setCalcXVal(equationStorage.calcXVal);
						lineBuilder.setCalcSlope(equationStorage.calcSlope);
						const cutPoint = valAxis.scaling.logBase ? (Math.log(1000) / Math.log(valAxis.scaling.logBase)) : 1000;
						const lineCoords = lineBuilder.drawWithApproximatedBezier(0.01, 1.56, cutPoint);
						if (!lineCoords.failed) {
							storageElement.setBezierApproximationResults(lineCoords.mainLine, lineCoords.startPoint);
							storageElement.setLineReversed(lineCoords.isReversed);
						} else {
							bezierFailed = true;
						}
					}
					if (this.bAllowDrawByPoints || type === AscFormat.TRENDLINE_TYPE_POLY || bezierFailed) {
						const minLogVal = storageElement.getMinLogVal();
						const pointVals = lineBuilder.drawWithPoints(midPointsNum, minLogVal);
						storageElement.setPointVals(pointVals);
					}
					storageElement.setBoundary(lineBuilder.getBoundary());

					if (attributes.dispEq || attributes.dispEq === null) {
						storageElement.setCoefficients(coefficients);
					}
					if (attributes.dispRSqr || attributes.dispRSqr === null) {
						const rSquared = this._dispRSquared(coords.catVals, coords.valVals, coefficients, type);
						storageElement.setRSquared(rSquared);
					}
				}
			}
		},

		// compare boundaries[chartId][seriaId] with local coordinates[chartId][seriaId].boundary
		_setNewBoundaries: function (oChart, storageElement, boundaries) {
			const boundary = storageElement.getBoundary();
			if (boundary) {
				const catAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[0]);
				const valAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[1]);
				boundary.catMax = Math.ceil(boundary.catMax);
				boundary.catMin = Math.floor(boundary.catMin);

				if (!boundaries[catAxis.Id]) {
					boundaries[catAxis.Id] = {min: null, max: null};
				}

				if (!boundaries[valAxis.Id]) {
					boundaries[valAxis.Id] = {min: null, max: null};
				}

				if (!boundaries[catAxis.Id].min && !boundaries[catAxis.Id].max && !boundaries[valAxis.Id].min && !boundaries[valAxis.Id].max) {
					boundaries[catAxis.Id].min = boundary.catMin;
					boundaries[catAxis.Id].max = boundary.catMax;
					boundaries[valAxis.Id].min = boundary.valMin;
					boundaries[valAxis.Id].max = boundary.valMax;
				} else {
					boundaries[catAxis.Id].min = Math.min(boundary.catMin, boundaries[catAxis.Id].min);
					boundaries[catAxis.Id].max = Math.max(boundary.catMax, boundaries[catAxis.Id].max);
					boundaries[valAxis.Id].min = Math.min(boundary.valMin, boundaries[valAxis.Id].min);
					boundaries[valAxis.Id].max = Math.max(boundary.valMax, boundaries[valAxis.Id].max);
				}
			}
		},

		// After precalculate is finished, just transform all the data into the pixel format
		recalculate: function (charts) {
			for (let i in charts) {
				if (charts.hasOwnProperty(i) && this.storage[i]) {
					for (let j in charts[i].chart.series) {
						if (charts[i].chart.series.hasOwnProperty(j) && this.storage[i][charts[i].chart.series[j].Id]) {
							this._calculateRelativeLine(charts[i].chart.series[j].parent, this.storage[i][charts[i].chart.series[j].Id]);
							// console.log(this.getAdditionalInfo(i, charts[i].chart.series[j].Id));
						}
					}
				}
			}
		},

		_calculateRelativeLine: function (oChart, storageElement) {
			const catAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[0]);
			const valAxis = this.cChartDrawer._searchChangedAxis(oChart.axId[1]);
			const horOrientation = !!oChart.axId[0].xPoints;
			const bezierVals = storageElement.getBezierVals();
			const startBezierVal = storageElement.getStartBezierVal();
			const pointVals = storageElement.getPointVals();

			//Convert given values into pixel values
			const _valsToPos = function (arr, cChartDrawer) {
				const posArr = {xPos: [], yPos: []}
				for (let i = 0; i < arr.catVals.length; i++) {
					let pos1 = cChartDrawer.getYPosition(arr.catVals[i], catAxis, true);
					let pos2 = cChartDrawer.getYPosition(arr.valVals[i], valAxis, true);

					let catPos = horOrientation ? pos1 : pos2;
					let valPos = horOrientation ? pos2 : pos1;

					posArr.xPos.push(catPos);
					posArr.yPos.push(valPos);
				}
				return posArr
			};

			const pathH = this.cChartDrawer.calcProp.pathH;
			const pathW = this.cChartDrawer.calcProp.pathW;

			if (bezierVals) {

				const pathId = this.cChartDrawer.cChartSpace.AllocPath();
				let path = this.cChartDrawer.cChartSpace.GetPath(pathId);
				const positions = _valsToPos(bezierVals, this.cChartDrawer);

				// find the starting point
				if (startBezierVal) {
					const startPosiontion = _valsToPos(startBezierVal, this.cChartDrawer);
					path.moveTo(startPosiontion.xPos[0] * pathW, startPosiontion.yPos[0] * pathH);
					path.lnTo(positions.xPos[0] * pathW, positions.yPos[0] * pathH);
				} else {
					path.moveTo(positions.xPos[0] * pathW, positions.yPos[0] * pathH);
				}


				if (positions.xPos.length === 4) {
					path.cubicBezTo(positions.xPos[1] * pathW, positions.yPos[1] * pathH, positions.xPos[2] * pathW, positions.yPos[2] * pathH, positions.xPos[3] * pathW,
						positions.yPos[3] * pathH);
				} else if (positions.xPos.length === 3) {
					path.quadBezTo(positions.xPos[1] * pathW, positions.yPos[1] * pathH, positions.xPos[2] * pathW, positions.yPos[2] * pathH);
				} else {
					for (let i = 1; i < positions.xPos.length; i++) {
						path.lnTo(positions.xPos[i] * pathW, positions.yPos[i] * pathH);
					}
				}

				storageElement.setLastPoint(positions.xPos, positions.yPos);
				storageElement.setBezierPath(pathId);
			}

			if (pointVals) {
				const pathId = this.cChartDrawer.cChartSpace.AllocPath();
				let path = this.cChartDrawer.cChartSpace.GetPath(pathId);

				const positions = _valsToPos(pointVals, this.cChartDrawer);

				path.moveTo(positions.xPos[0] * pathW, positions.yPos[0] * pathH);

				for (let i = 0; i < positions.xPos.length; i++) {
					path.lnTo(positions.xPos[i] * pathW, positions.yPos[i] * pathH);
				}
				storageElement.setLastPoint(positions.xPos, positions.yPos);
				storageElement.setPointPath(pathId);
			}
		},

		// find moving average
		// idea is to create sliding window with slight changes
		// sum stands for the sum of all elements in the window
		// counter stands for the amount of the elements in the window
		_getMAline: function (catVals, valVals, ptCount, period, allow) {
			const result = {catVals: [], valVals: []};
			if (ptCount > period) {
				let sum = 0;
				let counter = 0;
				let j = 0;
				for (let i = 0; i < ptCount; i++) {
					//check if past existed and remove it
					if (i >= period && catVals[j - counter] === i + 1 - period) {
						sum -= valVals[j - counter];
						counter--;
					}
					//check if next exists and add it
					if (catVals[j] === i + 1) {
						sum += valVals[j];
						counter++;
						j++;
					}
					//sum/counter after sliding window will be created
					if (i >= (period - 1) && counter !== 0) {
						let yVal = sum / counter;
						if (!allow || yVal > 0) {
							result.valVals.push(yVal);
							result.catVals.push(i + 1);
						}
					}
				}
			}

			return result
		},

		_obtainEquationStorage: function (type) {
			switch (type) {
				case AscFormat.TRENDLINE_TYPE_EXP :
					return {calcYVal: function (val, supps, isLog) {
						const res = supps[1] * Math.exp(val * supps[0]);
						if (isLog && res <= 0) {
							return NaN;
						} else {
							return isLog ? (Math.log(res) / Math.log(isLog)) : res;
						}
					}, calcXVal: function (val, supps, isLog) {
						val = isLog ? Math.pow(isLog, val) : val;
						return Math.log(val / supps[1]) / supps[0];
					}, calcSlope: function (val, supps, isLog) {
						if (!isLog) {
							return supps[1] * supps[0] * Math.exp(val * supps[0]);
						} else {
							return (Math.log(Math.exp(1)) / Math.log(isLog)) * supps[0];
						}
					}};
				case AscFormat.TRENDLINE_TYPE_LINEAR :
					return {calcYVal: function (val, supps, isLog) {
						const res = supps[0] * val + supps[1];
						if (isLog && res <= 0) {
							return NaN;
						} else {
							return isLog ? (Math.log(res) / Math.log(isLog)) : res;
						}
					}, calcXVal: function (val, supps, isLog) {
						val = isLog ? Math.pow(isLog, val) : val;
						return (val - supps[1]) / supps[0];
					}, calcSlope: function (val, supps, isLog) {
						if (!isLog) {
							return supps[0];
						} else {
							return supps[0] / (Math.log(isLog) * (supps[0] * val + supps[1]));
						}
					}};
				case AscFormat.TRENDLINE_TYPE_LOG :
					return {calcYVal: function (val, supps, isLog) {
						if (val > 0) {
							const res = supps[0] * Math.log(val) + supps[1];
							if (isLog && res <= 0) {
								return NaN;
							} else {
								return isLog ? (Math.log(res) / Math.log(isLog)) : res;
							}
						} else {
							return NaN;
						}
					}, calcXVal: function (val, supps, isLog) {
						val = isLog ? Math.pow(isLog, val) : val;
						return Math.exp((val - supps[1]) / supps[0]);
					}, calcSlope: function (val, supps, isLog) {
						if (!isLog) {
							return supps[0] / val;
						} else {
							return supps[0] / (Math.log(isLog) * val * (supps[0] * Math.log(val) + supps[1]));
						}
					}};
				case AscFormat.TRENDLINE_TYPE_POLY :
					return {calcYVal: function (val, supps) {
						let result = 0;
						let power = 0;
						for (let i = supps.length - 1; i >= 0; i--) {
							result += (Math.pow(val, power) * supps[i]);
							power++;
						}
						return result;
					}};
				case AscFormat.TRENDLINE_TYPE_POWER :
					return {calcYVal: function (val, supps, isLog) {
						const res = supps[1] * Math.pow(val, supps[0]);
						if (isLog && res <= 0) {
							return NaN;
						} else {
							return isLog ? (Math.log(res) / Math.log(isLog)) : res;
						}
					}, calcXVal: function (val, supps, isLog) {
						val = isLog ? Math.pow(isLog, val) : val;
						return Math.pow((val / supps[1]), 1 / supps[0]);
					}, calcSlope: function (val, supps, isLog) {
						if (!isLog) {
							return supps[1] * supps[0] * Math.pow(val, supps[0] - 1);
						} else {
							return supps[0] / (Math.log(isLog) * val);
						}
					}};
				default : 
					return null;

			}
		},

		_getEquationCoefficients: function (catVals, valVals, type, pow, intercept) {
			if (catVals.length === valVals.length) {

				if (type === AscFormat.TRENDLINE_TYPE_LOG || type === AscFormat.TRENDLINE_TYPE_POWER) {
					intercept = null;
				}

				const size = catVals.length;
				pow = size < pow ? size : pow;

				const mappingStorage = this._obtainMappingStorage(type);

				const _mapCoordinates = function () {
					for (let i = 0; i < size; i++) {
						if (mappingStorage.xVal) {
							catVals[i] = mappingStorage.xVal(catVals[i]);
						}
						if (mappingStorage.yVal) {
							valVals[i] = mappingStorage.yVal(valVals[i]);
						}
						if (isNaN(catVals[i]) || isNaN(valVals[i])) {
							return false;
						}
					}
					return true;
				};

				const mapped = this.stopAdding ? true : _mapCoordinates();

				if (mapped) {

					/*
					X = createMatrix(coordsX)
					X_T = findTranspose(X)
					P = matMul(X, X_T)
					P_I = findInv(M)
					S = matMul(X_T, Y)
					result = matMul(P_I, S)
					*/

					const _createMatrix = function (arr, size, starting, shift) {
						let result = [];
						shift = shift ? shift : 0;
						if (shift && mappingStorage.yVal) {
							shift = mappingStorage.yVal(shift);
						}
						for (let i = 0; i < arr.length; i++) {
							let power = starting;
							let row = [];
							for (let j = 0; j < size; j++) {
								row.push(Math.pow((arr[i] - shift), power));
								power++;
							}
							result.push(row);
						}
						return result;
					};

					const _findTranspose = function (A, N, M) {
						const B = Array(N);
						for (let i = 0; i < N; i++) {
							B[i] = Array(M).fill(1);
						}
						for (let i = 0; i < N; i++) {
							for (let j = 0; j < M; j++) {
								B[i][j] = A[j][i];
							}
						}
						return B;
					};

					const _matMul = function (A, B, R, C, D) {
						const rslt = Array(R);
						for (let i = 0; i < R; i++) {
							rslt[i] = Array(C).fill(0);
						}
						for (let i = 0; i < R; i++) {
							for (let j = 0; j < C; j++) {
								for (let k = 0; k < D; k++) {
									rslt[i][j] += A[i][k] * B[k][j];
								}
							}
						}

						return rslt
					};

					const _findInverse = function (M) {
						// I use Guassian Elimination to calculate the inverse:
						// (1) 'augment' the matrix (left) by the identity (on the right)
						// (2) Turn the matrix on the left into the identity by elemetry row ops
						// (3) The matrix on the right is the inverse (was the identity matrix)
						// There are 3 elemtary row ops: (I combine b and c in my code)
						// (a) Swap 2 rows
						// (b) Multiply a row by a scalar
						// (c) Add 2 rows

						//if the matrix isn't square: exit (error)
						if (M.length !== M[0].length) {
							return;
						}

						//create the identity matrix (I), and a copy (C) of the original
						let i = 0, ii = 0, j = 0, dim = M.length, e = 0, t = 0;
						let I = [], C = [];
						for (let i = 0; i < dim; i += 1) {
							// Create the row
							I[I.length] = [];
							C[C.length] = [];
							for (let j = 0; j < dim; j += 1) {

								//if we're on the diagonal, put a 1 (for identity)
								if (i === j) {
									I[i][j] = 1;
								} else {
									I[i][j] = 0;
								}

								// Also, make the copy of the original
								C[i][j] = M[i][j];
							}
						}

						// Perform elementary row operations
						for (let i = 0; i < dim; i += 1) {
							// get the element e on the diagonal
							e = C[i][i];

							// if we have a 0 on the diagonal (we'll need to swap with a lower row)
							if (e === 0) {
								//look through every row below the i'th row
								for (ii = i + 1; ii < dim; ii += 1) {
									//if the ii'th row has a non-0 in the i'th col
									if (C[ii][i] !== 0) {
										//it would make the diagonal have a non-0 so swap it
										for (let j = 0; j < dim; j++) {
											e = C[i][j];       //temp store i'th row
											C[i][j] = C[ii][j];//replace i'th row by ii'th
											C[ii][j] = e;      //repace ii'th by temp
											e = I[i][j];       //temp store i'th row
											I[i][j] = I[ii][j];//replace i'th row by ii'th
											I[ii][j] = e;      //repace ii'th by temp
										}
										//don't bother checking other rows since we've swapped
										break;
									}
								}
								//get the new diagonal
								e = C[i][i];
								//if it's still 0, not invertable (error)
								if (e === 0) {
									return
								}
							}

							// Scale this row down by e (so we have a 1 on the diagonal)
							for (let j = 0; j < dim; j++) {
								C[i][j] = C[i][j] / e; //apply to original matrix
								I[i][j] = I[i][j] / e; //apply to identity
							}

							// Subtract this row (scaled appropriately for each row) from ALL of
							// the other rows so that there will be 0's in this column in the
							// rows above and below this one
							for (let ii = 0; ii < dim; ii++) {
								// Only apply to other rows (we want a 1 on the diagonal)
								if (ii === i) {
									continue;
								}

								// We want to change this element to 0
								e = C[ii][i];

								// Subtract (the row above(or below) scaled by e) from (the
								// current row) but start at the i'th column and assume all the
								// stuff left of diagonal is 0 (which it should be if we made this
								// algorithm correctly)
								for (let j = 0; j < dim; j++) {
									C[ii][j] -= e * C[i][j]; //apply to original matrix
									I[ii][j] -= e * I[i][j]; //apply to identity
								}
							}
						}

						//we've done all operations, C should be the identity
						//matrix I should be the inverse:
						return I;
					};

					const flatten = function (arr, intercept, value) {
						const res = Array(arr.length);
						for (let i = 0; i < arr.length; i++) {
							res[res.length - 1 - i] = arr[i][0];
						}
						if (intercept !== 0) {
							res.push(value);
						} else {
							if (mappingStorage.bValBackward) {
								res[res.length - 1] = mappingStorage.bValBackward(res[res.length - 1]);
							}
						}
						return res;
					};

					// if intercept is number then one column shoud be dropped, therefor y_intercept stands for number of columns that should be dropped
					const y_intercept = (typeof intercept === 'number') ? 1 : 0;

					// Letiables = ((X_T@X)^-1)@X_T@Y
					const X = _createMatrix(catVals, pow - y_intercept, y_intercept, 0);
					const X_T = _findTranspose(X, pow - y_intercept, catVals.length);
					const P = _matMul(X_T, X, pow - y_intercept, pow - y_intercept, catVals.length);
					const P_I = _findInverse(P);
					if (P_I) {
						const Y = _createMatrix(valVals, 1, 1, intercept);
						const S = _matMul(X_T, Y, pow - y_intercept, 1, catVals.length);
						const letiables = _matMul(P_I, S, pow - y_intercept, 1, pow - y_intercept);
						return flatten(letiables, y_intercept, intercept);
					}
				}
			}
		},

		_obtainMappingStorage: function (type) {
			switch (type) {
				case AscFormat.TRENDLINE_TYPE_EXP :
					return {
						yVal: function (val) {
							return val > 0 ? Math.log(val) : NaN;
						}, bValBackward: function (val) {
							return Math.exp(val);
						}, bValForward: function (val) {
							return Math.log(val);
						}, yValBackward: function (val) {
							return Math.pow(Math.exp(1), val);
						}};
				case AscFormat.TRENDLINE_TYPE_LOG :
					return {
						xVal: function (val) {
							return val > 0 ? Math.log(val) : NaN;
						}};
				case AscFormat.TRENDLINE_TYPE_POWER :
					return {
						xVal: function (val) {
							return val > 0 ? Math.log(val) : NaN;
						}, yVal: function (val) {
							return val > 0 ? Math.log(val) : NaN;
						}, bValBackward: function (val) {
							return Math.exp(val);
						}, bValForward: function (val) {
							return Math.log(val);
						}, yValBackward: function (val) {
							return Math.pow(Math.exp(1), val);
						}};
				default : 
						return {};
			}
		},

		_dispRSquared: function (catVals, valVals, coefficients, type) {
			const mappingStorage = this._obtainMappingStorage(type);
			const predictY = function (xVal) {
				let result = mappingStorage.bValForward ? mappingStorage.bValForward(coefficients[coefficients.length - 1]) : coefficients[coefficients.length - 1];
				let power = 1;
				for (let i = coefficients.length - 2; i >= 0; i--) {
					result += (Math.pow(xVal, power) * coefficients[i]);
					power++;
				}
				result = mappingStorage.yValBackward ? mappingStorage.yValBackward(result) : result;
				return result;
			};

			// R squared = Correlation * Correlation
			// Correlation = (N∑xy-∑x∑y) / (sqrt([N∑x^2-(∑x)^2][N∑y^2-(∑y)^2]))

			const N = valVals.length;
			let sumSquared = 0;
			let sumRegression = 0;
			let XY = 0;
			let X = 0;
			let Y = 0;
			let XSquared = 0;
			let YSquared = 0;
			for (let i = 0; i < N; i++) {
				const yVal = mappingStorage.yValBackward ? mappingStorage.yValBackward(valVals[i]) : valVals[i];
				const yValPred = predictY(catVals[i]);
				XY += (yVal * yValPred);
				X += yVal;
				Y += yValPred;
				XSquared += (yVal * yVal);
				YSquared += (yValPred * yValPred);
			}
			let divisor = (Math.sqrt((N * XSquared - Math.pow(X, 2)) * (N * YSquared - Math.pow(Y, 2))));
			if (divisor != 0) {
				const correlation = ((N * XY) - (X * Y)) / divisor;
				return correlation * correlation;
			}
			return null;
		},

		draw: function () {
			let plotArea = this.cChartDrawer.cChartSpace.chart.plotArea;
			for (let i in this.storage) {
				let oChart = this.cChartDrawer._getChartModelById(plotArea, i);
				if (!this.storage.hasOwnProperty(i) || !oChart || !this.storage[i]) {
					continue;
				}
				for (let j in this.storage[i]) {
					let oSeries = this.cChartDrawer._getSeriesById(oChart, j);
					if (!this.storage[i].hasOwnProperty(j) || !oSeries || !this.storage[i][j]) {
						continue;
					}
					let pen = oSeries.trendline.getPen();
					if (!pen) {
						pen = this.cChartDrawer.cChartSpace.chart.plotArea.axId[1].compiledMajorGridLines;
					}
					const bezierPath = this.storage[i][j].getBezierPath();
					const pointPath = this.storage[i][j].getPointPath();
					if (pen && (bezierPath || pointPath)) {
						let leftRect = this.cChartDrawer.calcProp.chartGutter._left / this.cChartDrawer.calcProp.pxToMM;
						let topRect = (this.cChartDrawer.calcProp.chartGutter._top) / this.cChartDrawer.calcProp.pxToMM;
						let rightRect = this.cChartDrawer.calcProp.trueWidth / this.cChartDrawer.calcProp.pxToMM;
						let bottomRect = (this.cChartDrawer.calcProp.trueHeight) / this.cChartDrawer.calcProp.pxToMM;

						this.cChartDrawer.cShapeDrawer.Graphics.SaveGrState();
						this.cChartDrawer.cShapeDrawer.Graphics.AddClipRect(leftRect, topRect, rightRect, bottomRect);

						let bezierExist = false;
						if (bezierPath) {
							this.cChartDrawer.drawPath(bezierPath, pen);
							bezierExist = true;
						}
						if (pointPath) {
							if (bezierExist) {
								pen.Fill.fill.color.RGBA.R = 255;
								pen.Fill.fill.color.RGBA.G = 0;
								pen.Fill.fill.color.RGBA.B = 0;
							}
							this.cChartDrawer.drawPath(pointPath, pen);
						}

						this.cChartDrawer.cShapeDrawer.Graphics.RestoreGrState();
					}
				}
			}
		}
	};

	function CLineBuilder(coefficients, catMin, catMax, valMin, valMax, isLog) {
		this.coefficients = coefficients;
		this.valMin = valMin;
		this.valMax = valMax;
		this.catMin = catMin;
		this.catMax = catMax;
		this.isLog = isLog;
		this.calcYVal = null;
		this.calcXVal = null;
		this.calcSlope = null;
		this.boundary = null;
	}

	// set calcYVal! calcXVal! calcSlope!
	CLineBuilder.prototype = {

		constructor: CLineBuilder,

		setCalcYVal: function (func) {
			this.calcYVal = func;
		},

		setCalcXVal: function (func) {
			this.calcXVal = func;
		},

		setCalcSlope: function (func) {
			this.calcSlope = func;
		},

		getBoundary: function () {
			return this.boundary;
		},

		// calculate approximated bezier
		// normalize if axis in log scale
		// find boundary
		drawWithApproximatedBezier: function (error, tailLimit, cutPoint) {
			error = error ? error : 0.01;
			tailLimit = tailLimit ? tailLimit : 1.56;
			if (this.isLog && !cutPoint) {
				cutPoint = Math.log(1000) / Math.log(this.isLog);
			}
			if (!this.isLog && !cutPoint) {
				cutPoint = 1000;
			}
			const results = this._calcApproximatedBezier(error, tailLimit, cutPoint);

			if (!results.failed) {
				if (this.isLog) {
					this._normalize(results.mainLine);
					this._normalize(results.startPoint);
				}
	
				let catMinStart = results.startPoint.catVals.length > 0 ? results.startPoint.catVals[0] : null;
				let catMaxStart = results.startPoint.catVals.length > 0 ? results.startPoint.catVals[0] : null;
				let valMinStart = results.startPoint.valVals.length > 0 ? results.startPoint.valVals[0] : null;
				let valMaxStart = results.startPoint.valVals.length > 0 ? results.startPoint.valVals[0] : null;
	
				let size = results.mainLine.catVals.length;
				let catMin = Math.min(results.mainLine.catVals[0], results.mainLine.catVals[size - 1]);
				let catMax = Math.max(results.mainLine.catVals[0], results.mainLine.catVals[size - 1]);
				let valMin = Math.min(results.mainLine.valVals[0], results.mainLine.valVals[size - 1]);
				let valMax = Math.max(results.mainLine.valVals[0], results.mainLine.valVals[size - 1]);
	
				if (results.startPoint.valVals.length > 0) {
					catMin = Math.min(catMin, catMinStart);
					catMax = Math.max(catMax, catMaxStart);
					valMin = Math.min(valMin, valMinStart);
					valMax = Math.max(valMax, valMaxStart);
				}
	
				if (!this.boundary) {
					this.boundary = {catMin: catMin, catMax: catMax, valMin: valMin, valMax: valMax};
				} else {
					this.boundary.catMin = Math.min(this.boundary.catMin, catMin);
					this.boundary.catMax = Math.max(this.boundary.catMax, catMax);
					this.boundary.valMin = Math.min(this.boundary.valMin, valMin);
					this.boundary.valMax = Math.max(this.boundary.valMax, valMax);
				}
			}

			return results;
		},

		// this function can return two lines 
		// secondaryLine which is just a straight line 
		// primaryLine which is bezier driven line
		_calcApproximatedBezier: function (error, tailLimit, cutPoint) {
			// error: the threshold to check the accuracy of the approximated line 
			// tailLimit: the threshold at which the slope of the point becomes too high leading to its arctangent to become high
			// cutPoint: the threshold at which line should be cutted. For lines with a long tails, approximation does not work, therefore they must be cutted
			const controlPoints = {catVals: [], valVals: []};
			if (!this.calcYVal || !this.calcXVal || !this.calcSlope) {
				return
			}

			// Step 1 find edges of line 
			// --------------------------------------

			let start = {catVal: null, valVal: null};
			const end = {catVal: null, valVal: null};
			this._lineCoordinate(start, this.catMin, true);
			this._lineCoordinate(end, this.catMax, true);
			// yValue for log scale can not be <= 0
			// therefore use close to 0 val 
			if (this.isLog) {
				const lowerBoundary = (Math.log(0.01) / Math.log(this.isLog));
				if (isNaN(start.valVal)) {
					this._lineCoordinate(start, lowerBoundary, false);
				}
				if (isNaN(end.valVal)) {
					this._lineCoordinate(end, lowerBoundary, false);
				}
			}

			this._checkBoundaries(start);
			this._checkBoundaries(end);

			//start point cant be changed with end point
			let isReversed = false;
			//start should always be lower than end
			if (start.valVal > end.valVal) {
				isReversed = true;
				const temp = {catVal: end.catVal, valVal: end.valVal};
				end.valVal = start.valVal;
				end.catVal = start.catVal;
				start = temp;
			}

			//sometimes lines can point in opposite direction, we need to track directions
			let direction = true;
			const line2Letiables = this._findLine(end.catVal, end.valVal);
			let line1Letiables = this._findLine(start.catVal, start.valVal);

			// if the slope is to low at the end point, then
			// starting point should be in the range of three order of magnitude. 
			const lowestPoint = this.isLog ? end.valVal - cutPoint : end.valVal / cutPoint;

			// sometimes graph can have huge tails, to detect them use slope of starting and end point of line
			// if the slope of the starting point is to high, then cut the line up to certain point (lowestPoint)
			if (Math.atan(Math.abs(line2Letiables[1])) < tailLimit && Math.atan(Math.abs(line1Letiables[1])) > tailLimit && start.valVal < lowestPoint) {
				this._lineCoordinate(start, lowestPoint, false);
				line1Letiables = this._findLine(start.catVal, start.valVal);
			}
			// if the slope of the ending point is to high, then cut the line into 2/3 of its original size
			const startPoint = {catVals: [], valVals: []};
			if (Math.atan(Math.abs(line1Letiables[1])) < tailLimit && Math.atan(Math.abs(line2Letiables[1])) > tailLimit) {
				startPoint.catVals.push(start.catVal);
				startPoint.valVals.push(start.valVal);
				start = this._cutLine(start, end);
				line1Letiables = this._findLine(start.catVal, start.valVal);
				direction = false;
			}

			//------------------------------------
			//Step 2 check slopes at the edges to continue
			controlPoints.catVals.push(start.catVal);
			controlPoints.valVals.push(start.valVal);

			if (line2Letiables[1] === line1Letiables[1]) {
				controlPoints.catVals.push(end.catVal);
				controlPoints.valVals.push(end.valVal);
				return {isReversed: isReversed, startPoint: startPoint, mainLine: controlPoints}
			}

			// --------------------------------------
			//Step 3 calculate mid control point
			const midCatVal = (line2Letiables[0] - line1Letiables[0]) / (line1Letiables[1] - line2Letiables[1]);
			const midValVal = (midCatVal * line1Letiables[1]) + line1Letiables[0];
			controlPoints.catVals.push(midCatVal);
			controlPoints.valVals.push(midValVal);
			controlPoints.catVals.push(end.catVal);
			controlPoints.valVals.push(end.valVal);

			// Based on the midpoint derive the direction of the line either long curve to right or long curve to bottom
			const diff1 = Math.abs(start.catVal - midCatVal);
			const diff2 = Math.abs(midCatVal - end.catVal);
			direction = diff1 > diff2 ? false : direction;

			// --------------------------------------

			//Step 4 check mid control point
			const val1 = this._predictY(controlPoints.valVals, 1);
			const val2 = this._predictY(controlPoints.valVals, 0);
			const valMax = Math.max(val1, val2);
			const valMin = Math.min(val1, val2);
			const predictedErr = this._check(controlPoints, valMin, valMax);
			if (predictedErr <= error && predictedErr >= -error) {
				return {isReversed: isReversed, startPoint: startPoint, mainLine: controlPoints}
			}
			//--------------------------------------------
			//Step 5 find approximation of cubic bezier;
			//Step 6 for each approximation find the error, if error is huge repeat step 5 with new approximation
			controlPoints.catVals.splice(1, 0, null);
			controlPoints.valVals.splice(1, 0, null);

			let lineStart = start.valVal;
			let lineEnd = midValVal;
			const storage = {minError: null, cpY: null, cpX: null};
			let failed = true;
			for (let i = 0; i < 10; i++) {
				controlPoints.valVals[1] = (0.5 * lineEnd) + (0.5 * lineStart);
				controlPoints.catVals[1] = (controlPoints.valVals[1] - line1Letiables[0]) / line1Letiables[1];
				let predictedError = this._check(controlPoints, valMin, valMax);
				if (storage.minError === null || Math.abs(predictedError) <= storage.minError) {
					storage.minError = Math.abs(predictedError);
					storage.cpY = controlPoints.valVals[1];
					storage.cpX = controlPoints.catVals[1];
				}

				if ((predictedError > error && direction) || (predictedError < -error && !direction)) {
					lineEnd = controlPoints.valVals[1];
				} else if ((predictedError > error && !direction) || (predictedError < -error && direction)) {
					lineStart = controlPoints.valVals[1];
				} else {
					failed = false;
					break;
				}
			}

			if (failed) {
				return {failed : true};
			}

			if (controlPoints.catVals[1] !== storage.cpX || controlPoints.valVals[1] !== storage.cpY) {
				controlPoints.catVals[1] = storage.cpX;
				controlPoints.valVals[1] = storage.cpY;
			}

			return {isReversed: isReversed, startPoint: startPoint, mainLine: controlPoints};
		},

		// find catVal and valVal of line at given point
		_lineCoordinate: function (point, val, calcValAxis) {
			if (calcValAxis) {
				point.catVal = val;
				point.valVal = this.calcYVal(val, this.coefficients, this.isLog);
			} else {
				point.catVal = this.calcXVal(val, this.coefficients, this.isLog);
				point.valVal = val;
			}
		},

		// Check for val boundary
		_checkBoundaries: function (pos) {
			if (this.valMin) {
				let lowerBoundary = this.isLog ? (Math.log(this.valMin) / Math.log(this.isLog)) : this.valMin;
				if (pos.valVal < lowerBoundary) {
					this._lineCoordinate(pos, lowerBoundary, false);
				}
			}
			if (this.valMax) {
				let upperBoundary = this.isLog ? (Math.log(this.valMax) / Math.log(this.isLog)) : this.valMax;
				if (pos.valVal > upperBoundary) {
					this._lineCoordinate(pos, upperBoundary, false);
				}
			}
		},

		// find slope of the line at given point
		_findLine: function (catVal, valVal) {
			const m = this.calcSlope(catVal, this.coefficients, this.isLog);
			const b = valVal - (m * catVal);
			return [b, m];
		},

		//we need to cut a tail of the line
		//up to 1 percent of its Val (Y) length 
		_cutLine: function (start, end) {
			let treshold = 0.05;
			let obtainedError = 0;
			const maxIterations = 10;
			let i = 0;
			let startingPoint =  {catVal: start.catVal, valVal: start.valVal};
			let prev = null;
			let newObtainedError = null;
			while (obtainedError < treshold && i < maxIterations) {
				prev = {catVal: start.catVal, valVal: start.valVal};
				this._lineCoordinate(start, (end.catVal + start.catVal) / 2, true);
				newObtainedError = 1 - (end.valVal - start.valVal)/ (end.valVal - startingPoint.valVal);
				if (newObtainedError >= obtainedError) {
					obtainedError = newObtainedError;
				} else {
					break;
				}
				i++;
			}
			return prev ? prev : start;
		},

		// predict catVal and valVal from bezier line using bernstein polynomials
		// find actual valVal at given predicted catVal
		// compare both valVals (actual and predicted) on normalized scale
		_check: function (controlPoints, valMin, valMax) {

			const t = 0.5;
			const predictedYVal = this._predictY(controlPoints.valVals, t);
			const predictedXVal = this._predictY(controlPoints.catVals, t);
			const actualYVal = this.calcYVal(predictedXVal, this.coefficients, this.isLog);
			const normalizedPredictedYVal = (predictedYVal - valMin) / (valMax - valMin);
			const normalizedActualYVal = (actualYVal - valMin) / (valMax - valMin);
			return normalizedPredictedYVal - normalizedActualYVal;
		},

		//https://en.wikipedia.org/wiki/B%C3%A9zier_curve
		_predictY: function (cP, t) {
			const len = cP.length;
			const a = this._obtainBernsteinPolynomials(len - 1, t);
			let result = 0;
			for (let i = 0; i < len; i++) {
				result += a[i] * cP[i];
			}
			return result;
		},

		_obtainBernsteinPolynomials: function (len, t) {
			const result = [];
			result.push(Math.pow((1 - t), len));
			result.push(len * t * Math.pow((1 - t), (len - 1)));
			if (len === 3) {
				result.push(len * Math.pow(t, (len - 1)) * (1 - t));
			}
			result.push(Math.pow(t, len));
			return result
		},

		_normalize: function (arr) {
			for (let i = 0; i < arr.valVals.length; i++) {
				arr.valVals[i] = Math.pow(this.isLog, arr.valVals[i]);
			}
		},

		// set calcYVal!
		drawWithPoints: function (pointsNumber, valMin) {

			const results = {catVals: [], valVals: []};

			if (!this.calcYVal) {
				return results;
			}

			let catNet = this.catMax - this.catMin;

			// find positive edges for log axis scenario
			let k = 0;
			let j = pointsNumber - 1;
			let remainder = pointsNumber;

			if (this.isLog) {
				// take first positive point along the line
				while (k < pointsNumber) {
					const catValMin = (catNet / (pointsNumber - 1)) * k + this.catMin;
					let valValMin = this.calcYVal(catValMin, this.coefficients);
					remainder--;
					if (valValMin >= 0) {
						this.catMin = catValMin;
						break;
					}
					k++;
				}

				// if there exist positive point along the line, take last positive point
				if (k !== pointsNumber - 1) {
					while (j >= 0) {
						const catValMax = (catNet / (pointsNumber - 1)) * j + this.catMin;
						let valValMax = this.calcYVal(catValMax, this.coefficients);
						remainder--;
						if (valValMax >= 0) {
							this.catMax = catValMax;
							break;
						}
						j--;
					}
				}

				// compare if first or last points actually changed 
				if (k !== 0 || j !== pointsNumber - 1) {
					catNet = this.catMax - this.catMin;
				}
			}

			// find everything between edges
			// k is index of first point and j is index oj last point
			if (k !== j) {
				for (let i = 0; i < remainder; i++) {
					const catVal = (catNet / (remainder - 1)) * i + this.catMin;
					let valVal = this.calcYVal(catVal, this.coefficients);
					//in the log axis, yVal can not be negative
					if (this.isLog && valVal <= 0) {
						valVal = valMin ? Math.pow(this.isLog, Math.floor(Math.log(valMin) / Math.log(this.isLog))) : 0.01;
					}

					results.catVals.push(catVal);
					results.valVals.push(valVal);

					if (!this.boundary) {
						this.boundary = {catMin: catVal, catMax: catVal, valMin: valVal, valMax: valVal};
					} else {
						this.boundary.catMin = Math.min(this.boundary.catMin, catVal);
						this.boundary.catMax = Math.max(this.boundary.catMax, catVal);
						this.boundary.valMin = Math.min(this.boundary.valMin, valVal);
						this.boundary.valMax = Math.max(this.boundary.valMax, valVal);
					}
				}
			}
			return results;
		}
	}

	function CTrendData() {
		this.coords = {catVals: [], valVals: []};
		this.bezierPath = null;
		this.pointPath = null;
		this.coefficients = null;
		this.rSquared = null;
		this.bezierVals = null;
		this.pointVals = null;
		this.boundary = null;
		this.minLogVal = null;
		this.startBezierVal = null;
		this.isReversed = false;
		this.lastPoint = {catVal: null, valVal: null};
	}

	// set calcYVal! calcXVal! calcSlope!
	CTrendData.prototype = {
		constructor: CTrendData,

		isEmpty: function () {
			if (this.coords.catVals.length === 0 || this.coords.valVals.length === 0){
				return true;
			}
			return false;
		},

		getBezierPath: function () {
			return this.bezierPath;
		},

		getPointPath: function () {
			return this.pointPath;
		},

		getMinLogVal: function () {
			return this.minLogVal;
		},

		getCoefficients: function () {
			return this.coefficients;
		},

		getRSquared: function () {
			return this.rSquared;
		},

		getLastPoint: function () {
			return this.lastPoint;
		},

		getCoords: function () {
			return this.coords;
		},

		getBoundary: function () {
			if (this.boundary && isFinite(this.boundary.catMin) && isFinite(this.boundary.catMax) && isFinite(this.boundary.valMin) && isFinite(this.boundary.valMax) &&
				this.boundary.catMin !== null && this.boundary.catMax !== null && this.boundary.valMin !== null && this.boundary.valMax !== null) {
				return this.boundary;
			}
			return null;
		},

		getBezierVals: function () {
			return this.bezierVals;
		},

		getStartBezierVal: function () {
			if (this.startBezierVal && this.startBezierVal.catVals.length > 0 && this.startBezierVal.valVals.length > 0) {
				return this.startBezierVal;
			}
			return null
		},

		getPointVals: function () {
			return this.pointVals
		},

		setBezierApproximationResults: function (mainLine, startingPoint) {
			this.bezierVals = mainLine;
			this.startBezierVal = startingPoint;
		},

		setLineReversed: function (isReversed) {
			this.isReversed = isReversed;
		},

		addCatVal: function (val) {
			this.coords.catVals.push(val);
		},

		addValVal: function (val) {
			this.coords.valVals.push(val);
		},

		setMinLogVal: function (val) {
			if (!this.minLogVal) {
				this.minLogVal = val;
			} else {
				this.minLogVal = Math.min(this.minLogVal, val);
			}
		},

		setPointVals: function (pointVals) {
			this.pointVals = pointVals;
		},

		setCoefficients: function (coefficients) {
			this.coefficients = coefficients;
		},

		setRSquared: function (rSquared) {
			this.rSquared = rSquared;
		},

		setBoundary: function (boundary) {
			this.boundary = boundary;
		},

		setLastPoint: function (catArr, valArr) {
			this.lastPoint.catVal = this.isReversed ? catArr[0] : catArr[catArr.length - 1];
			this.lastPoint.valVal = this.isReversed ? valArr[0] : valArr[valArr.length - 1];
		},

		setBezierPath: function (bezierPath) {
			this.bezierPath = bezierPath;
		},

		setPointPath: function (pointPath) {
			this.pointPath = pointPath;
		}
	}

	function CCachedChartExData (type, data) {
		this.type = type;
		this.data = data;
	}

	CCachedChartExData.prototype = {
		constructor: CCachedChartExData,
		forEach: function (callBack) {
			if (this.data) {
				this.data.forEach(callBack);
			}
		},
		_chartExSetAxisMinAndMax: function (axis, num) {
			if (!axis || (num !== 0 && !num)) {
				return;
			}
			axis.max = (axis.max === null || num > axis.max) ? num : axis.max;
			axis.min = (axis.min === null || num < axis.min) ? num : axis.min;
		}
	}

	function CCachedClusteredColumn(type, seria, numLit, strLit, axisProperties, secondType) {
		// if subType is true then aggregation else binning
		this.subTypeAggr = seria && seria.layoutPr && seria.layoutPr.aggregation;
		this.binning = null;
		CCachedChartExData.call(this, type, []);
		this._calculate(seria, numLit, strLit, axisProperties);
		if (secondType === AscFormat.SERIES_LAYOUT_PARETO_LINE) {
			this.data.sort(function (a, b) {
				return b.val - a.val;
			});
		}
	}

	AscFormat.InitClassWithoutType(CCachedClusteredColumn, CCachedChartExData);

	CCachedClusteredColumn.prototype._calculate = function (seria, numLit, strLit, axisProperties) {
		if (!seria || !seria.layoutPr || !numLit || !axisProperties) {
			return;
		}

        const numArr = numLit.pts;

        // set axis min to zero
		this._chartExSetAxisMinAndMax(axisProperties.val, 0);

		if (this.subTypeAggr) {
            const strArr = strLit ? strLit.pts : null;
			this.calculateAggregation(numArr, strArr, axisProperties);
		} else {
			this.binning = seria.layoutPr.binning;
			this.calculateBinning(numArr, axisProperties);
		}
	}

	CCachedClusteredColumn.prototype.calculateAggregation = function (numArr, strArr, axisProperties) {
        if (!numArr) {
            return;
        }

		axisProperties.cat.scale.push(1);
		const dictSet = {}
		if (strArr && strArr.length !== 0) {
			// create object of key and values
			for (let i = 0; i < numArr.length; i++) {
				const key = strArr[i].val;
				if (AscFormat.isRealNumber(dictSet[key])) {
					this.data[dictSet[key]].val += numArr[i].val;
				} else {
					this.data.push({lblName: key, val: numArr[i].val});
					dictSet[key] = this.data.length - 1;
				}
				this._chartExSetAxisMinAndMax(axisProperties.val, this.data[dictSet[key]].val);
			}
		} else {
			// Cases when labels do not exist
			let sum = 0;
			for (let i = 0; i < numArr.length; i++) {
				sum += numArr[i].val;
			}
			this.data.push({lblName: '', val: sum});
			this._chartExSetAxisMinAndMax(axisProperties.val, sum);
		}
	}

	CCachedClusteredColumn.prototype.calculateBinning = function (numArr, axisProperties) {
		if (!this.binning) {
			return;
		}
		const addRangesAndFillCatScale = function (localResults, localBinning, catLimits, axisProperties) {
			if (!localResults || !localBinning || !catLimits || !axisProperties) {
				return;
			}
			let prev = catLimits.trueMin;
			const addRange = function (minVal, maxVal) {
				localResults.push({ min: minVal, max: maxVal, val: 0 });
			}
			if (catLimits.isUnderflowExist) {
				addRange(null, prev);
			}

			axisProperties.cat.scale.push(prev);
			for (let i = 0; i < localBinning.compiledBinCount; i++) {
				let curr = prev + localBinning.compiledBinSize;

				if (catLimits.isOverflowExist) {
					curr = Math.min(curr, catLimits.trueMax);
				}

				addRange(prev, curr);
				prev = curr;
				axisProperties.cat.scale.push(curr);
			}

			if (catLimits.isOverflowExist) {
				addRange(catLimits.trueMax, null);
			}
		}

		const countOccurrencesAndValExtremum = function (localResults, localBinning, numArr, axisProperties, chartsDrawer) {
			if (!localResults || !localBinning || !numArr || !axisProperties) {
				return ;
			}
			if (numArr.length === 1) {
				localResults[0].val++;
				chartsDrawer._chartExSetAxisMinAndMax(axisProperties.val, 1);
			} else {
				for (let i = 0; i < numArr.length; i++) {
					// sometimes when bincount is fixed, values shifts into the last bean.
					// if this issue will be fixed remove isFound and last if case
					let isFound = false;
					for (let j = 0; j < localResults.length; j++) {
						const min = localResults[j].min;
						const max = localResults[j].max;
						const statement1 = (j === 0 && numArr[i].val === min);
						const statement2 = localBinning.intervalClosed !== AscFormat.INTERVAL_CLOSED_SIDE_L ?
							((!min || numArr[i].val > min) && (!max || numArr[i].val <= max)) :
							((!min || numArr[i].val >= min) && (!max || numArr[i].val < max));
						if (statement1 || statement2) {
							isFound = true;
							localResults[j].val++;
							chartsDrawer._chartExSetAxisMinAndMax(axisProperties.val, localResults[j].val);
						}
					}
					if (!isFound && localResults.length > 1) {
						const lastIndex = localResults.length - 1;
						localResults[lastIndex].val++;
						chartsDrawer._chartExSetAxisMinAndMax(axisProperties.val, localResults[lastIndex].val);
					}
				}
			}
		}

		const createSubChars = function (localResults, localBinning) {
			let start = '[';
			let end = localBinning.intervalClosed === AscFormat.INTERVAL_CLOSED_SIDE_L ? ')' : ']';
			// user can manually set minimum and maximum, therefore alternative start and end needed
			const startByExpression = localBinning.intervalClosed === AscFormat.INTERVAL_CLOSED_SIDE_L ? '<' : '≤';
			const endByExpression = localBinning.intervalClosed === AscFormat.INTERVAL_CLOSED_SIDE_L ? '≥' : '>';

			for (let i = 0; i < localResults.length; i++) {
				if (localResults[i].min === null) {
					localResults[i].subChars = [startByExpression];
				} else if (localResults[i].max === null) {
					localResults[i].subChars = [endByExpression];
				} else {
					if (i === 1 && localBinning.intervalClosed !== AscFormat.INTERVAL_CLOSED_SIDE_L) {
						start = '(';
					}

					if (i === (localResults.length - 1) && localBinning.intervalClosed === AscFormat.INTERVAL_CLOSED_SIDE_L) {
						end = ']';
					}

					localResults[i].subChars = [start, end];
				}
			}
		}


		const catLimits = this.binning.recalculate(axisProperties);
		addRangesAndFillCatScale(this.data, this.binning, catLimits, axisProperties);
		countOccurrencesAndValExtremum(this.data, this.binning, numArr, axisProperties, this);
		createSubChars(this.data, this.binning, axisProperties);
	}

	function CCachedWaterfall(type, seria, numLit, axisProperties) {
		CCachedChartExData.call(this, type, []);
		this.subtotalsIndex = 0;
		this.subtotals = seria.layoutPr && seria.layoutPr.subtotals ? seria.layoutPr.subtotals.idx : [];
		this._calculate(seria, numLit, axisProperties);
	}

	AscFormat.InitClassWithoutType(CCachedWaterfall, CCachedChartExData);

	CCachedWaterfall.prototype._calculate = function (seria, numLit, axisProperties) {
		if (!seria || !seria.layoutPr || !numLit || !axisProperties) {
			return;
		}

		const ptCount = numLit.ptCount;
		const numArr = numLit.pts;

		axisProperties.val.min = 0;
		axisProperties.val.max = 0;
		let sum = 0;

		// Check if any total values exist is in the array
		this.subtotalsIndex = 0;

		let j = 0;
		for (let i = 0; i < ptCount; i++) {
			if (j < numArr.length && numArr[j].idx === i) {
				const isTotal = this.checkIsTotal(numArr[j].idx);
				sum = isTotal ? numArr[j].val : sum + numArr[j].val;
				j++;
			}
			this.data.push(sum);
			this._chartExSetAxisMinAndMax(axisProperties.val, sum);
			axisProperties.cat.scale.push(i + 1);
		}
	}

	CCachedWaterfall.prototype.checkIsTotal = function (idx) {
		if (this.subtotalsIndex < this.subtotals.length) {
			if (this.subtotals[this.subtotalsIndex] === idx) {
				this.subtotalsIndex += 1;
				return true;
			}
		}
		return false;
	}

	CCachedWaterfall.prototype.forEach = function (callback) {
		this.subtotalsIndex = 0;
		for (let i = 0; i < this.data.length ; i++) {
			const isTotal = this.checkIsTotal(i);
			callback(this.data[i], i, isTotal);
		}
	}

	function CCachedFunnel(type, numLit, axisProperties) {
		CCachedChartExData.call(this, type, []);
		this._calculate(numLit, axisProperties);
	}

	AscFormat.InitClassWithoutType(CCachedFunnel, CCachedChartExData);

	CCachedFunnel.prototype._calculate = function (numLit, axisProperties) {
		if (!numLit || !axisProperties) {
			return;
		}

		const ptCount = numLit.ptCount;
		const numArr = numLit.pts;

		//if there is at least one element in the array then min is 1, however if there is no elements then min is 0
		axisProperties.val.min = 0;
		axisProperties.val.max = 0;

		let j = 0;
		for (let i = 0; i < ptCount; i++) {
			let val = 0;
			if (j < numArr.length && numArr[j].idx === i) {
				val = numArr[j].val;
				j++;
			}
			this._chartExSetAxisMinAndMax(axisProperties.val, val);
			this.data.push(val);
			axisProperties.val.scale.push(ptCount - i);
		}
	}

	// if rounding is strong it affects whole number. Example 106.82 -> 107, for the precision 2
	// if weak, then only decimal places. Example 106.82 ->106.8, for the precision 1
	function _roundValue (num, isStrong, precision) {
		if (num !== 0 && (!num || !isFinite(num))) {
			return 1;
		}

		if (num === 0) {
			return num;
		}

		// if num is negative
		let isNegative = false;
		if (num < 0) {
			isNegative = true;
			num = -num;
		}

		if (!precision || precision < 0) {
			//default precision is 9!
			precision = 9;
		}

		let count = 0;

		// Normalize the number by adjusting its scale
		if (isStrong) {
			while (num >= 10) {
				num /= 10;
				count++;
			}
		}

		while (num < 1) {
			num *= 10;
			count--;
		}

		// Round the number to two decimal places
		const kF = Math.pow(10, precision);
		const roundedNum = Math.round(num * kF) / kF;

		// Return the normalized number with the appropriate scale
		num = (count >= 0) ? roundedNum * Math.pow(10, count) : roundedNum / Math.pow(10, -count);
		return isNegative ? -num : num;
	}


	//----------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].CChartsDrawer = CChartsDrawer;
	window['AscFormat'].CTrendline = CTrendline;
	window['AscFormat'].CLineBuilder = CLineBuilder;
	window['AscFormat'].CColorObj = CColorObj;
	window["AscFormat"].c_oChartTypes = c_oChartTypes;
	window["AscCommon"]._roundValue = _roundValue;
	window["AscFormat"].CachedClusteredColumn = CCachedClusteredColumn;
})(window);
