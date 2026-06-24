/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window, undefined){

// Import
var CShape = AscFormat.CShape;
var CChartSpace = AscFormat.CChartSpace;

    CChartSpace.prototype.getEditorType = function()
    {
        return 0;
    };
CChartSpace.prototype.recalculateTransform = function()
{
    CShape.prototype.recalculateTransform.call(this);
    this.localTransform.Reset();
};
CChartSpace.prototype.recalcText = function()
{
    this.recalcInfo.recalculateAxisLabels = true;
    this.recalcTitles2();
    this.handleUpdateInternalChart(false);
};

CChartSpace.prototype.recalculateBounds = CShape.prototype.recalculateBounds;
CChartSpace.prototype.getRotateAngle = CShape.prototype.getRotateAngle;
CChartSpace.prototype.getInvertTransform = CShape.prototype.getInvertTransform;
CChartSpace.prototype.hit = CShape.prototype.hit;
CChartSpace.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
CChartSpace.prototype.hitInPath = CShape.prototype.hitInPath;
CChartSpace.prototype.check_bounds = CShape.prototype.check_bounds;
CChartSpace.prototype.handleUpdateLn = function()
{
    this.recalcInfo.recalculatePenBrush = true;
    this.recalcInfo.recalculatePlotAreaPen = true;
    this.addToRecalculate();
};
CChartSpace.prototype.setRecalculateInfo = function()
{
    this.recalcInfo =
    {
        recalcTitle: null,
        bRecalculatedTitle: false,
        recalculateTransform: true,
        recalculateBounds:    true,
        recalculateChart:     true,
        recalculateSeriesColors: true,
        recalculateMarkers: true,
        recalculateGridLines: true,
        recalculateDLbls: true,
        recalculateAxisLabels: true,
        dataLbls:[],
        axisLabels: [],
        recalculateAxisVal: true,
        recalculateAxisTickMark: true,
        recalculateBrush: true,
        recalculatePen: true,
        recalculatePlotAreaBrush: true,
        recalculatePlotAreaPen: true,
        recalculateHiLowLines: true,
        recalculateUpDownBars: true,
        recalculateLegend: true,
        recalculateReferences: true,
        recalculatePenBrush: true,
        recalculateTextPr : true
		};
    this.chartObj = null;
    this.rectGeometry = AscFormat.ExecuteNoHistory(function(){return  AscFormat.CreateGeometry("rect");},  this, []);
    this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
};

CChartSpace.prototype.checkNeedRecalculate = function(){
    return this.recalcInfo.recalculateChart === true;
};

CChartSpace.prototype.recalcTransform = function()
{
    this.recalcInfo.recalculateTransform = true;
};
CChartSpace.prototype.recalcBounds = function()
{
    this.recalcInfo.recalculateBounds = true;
};
CChartSpace.prototype.recalcChart = function()
{
    this.recalcInfo.recalculateChart = true;
};
CChartSpace.prototype.recalcSeriesColors = function()
{
    this.recalcInfo.recalculateSeriesColors = true;
    this.recalcInfo.recalculatePenBrush = true;
    this.recalcInfo.recalculatePlotAreaBrush = true;
};

CChartSpace.prototype.recalcDLbls = function()
{
    this.recalcInfo.recalculateDLbls = true;
};

CChartSpace.prototype.addToRecalculate = CShape.prototype.addToRecalculate;

CChartSpace.prototype.handleUpdatePosition = function()
{
    this.recalcTransform();
    this.recalcBounds();
  //  this.recalcDLbls();
    //this.setRecalculateInfo();
    for(var i = 0; i < this.userShapes.length; ++i)
    {
        if(this.userShapes[i].object)
        {
            this.userShapes[i].object.handleUpdateExtents();
        }
    }
    this.addToRecalculate();
};
CChartSpace.prototype.handleUpdateFlip = function()
{
    this.recalcTransform();
    this.addToRecalculate();
};
CChartSpace.prototype.handleUpdateStyle = function()
{
    this.recalcInfo.recalculateSeriesColors = true;
    this.recalcInfo.recalculatePenBrush = true;
    this.recalcInfo.recalculateLegend = true;
    this.recalcInfo.recalculatePlotAreaBrush = true;
    this.recalcInfo.recalculatePlotAreaPen = true;
    this.recalcInfo.recalculateBrush = true;
    this.recalcInfo.recalculatePen = true;
    this.recalcInfo.recalculateHiLowLines = true;
    this.recalcInfo.recalculateUpDownBars = true;
    this.handleTitlesAfterChangeTheme();
    this.recalcInfo.recalculateAxisLabels = true;
    this.recalcInfo.recalculateAxisVal = true;
    this.addToRecalculate();
};

CChartSpace.prototype.handleUpdateFill = function()
{
    this.recalcInfo.recalculatePenBrush = true;
    this.recalcInfo.recalculatePlotAreaBrush = true;
    this.recalcInfo.recalculateBrush = true;
    this.recalcInfo.recalculateChart = true;
    this.recalcInfo.recalculateSeriesColors = true;
	this.recalcInfo.recalculateMarkers = true;
    this.addToRecalculate();
};
CChartSpace.prototype.handleUpdateLn = function()
{
    this.recalcInfo.recalculatePenBrush = true;
    this.recalcInfo.recalculatePlotAreaPen = true;
    this.recalcInfo.recalculatePen = true;
    this.recalcInfo.recalculateChart = true;
    this.recalcInfo.recalculateSeriesColors = true;
	this.recalcInfo.recalculateMarkers = true;
    this.addToRecalculate();
};
CChartSpace.prototype.convertPixToMM = CShape.prototype.convertPixToMM;
CChartSpace.prototype.getHierarchy = CShape.prototype.getHierarchy;
CChartSpace.prototype.getParentObjects = function()
{
    var parents = { slide: null, layout: null, master: null, theme: this.themeOverride ? this.themeOverride : window["Asc"]["editor"].wbModel.theme };
    if(this.clrMapOvr)
    {
        parents.slide = {clrMap: this.clrMapOvr};
    }
    return parents;
};
CChartSpace.prototype.recalculateTransform = CShape.prototype.recalculateTransform;

CChartSpace.prototype.createResizeTrack = CShape.prototype.createResizeTrack;
CChartSpace.prototype.createMoveTrack = CShape.prototype.createMoveTrack;

CChartSpace.prototype.recalculateBounds = function()
{
    var transform = this.transform;
    var a_x = [];
    var a_y = [];
    a_x.push(transform.TransformPointX(0, 0));
    a_y.push(transform.TransformPointY(0, 0));
    a_x.push(transform.TransformPointX(this.extX, 0));
    a_y.push(transform.TransformPointY(this.extX, 0));
    a_x.push(transform.TransformPointX(this.extX, this.extY));
    a_y.push(transform.TransformPointY(this.extX, this.extY));
    a_x.push(transform.TransformPointX(0, this.extY));
    a_y.push(transform.TransformPointY(0, this.extY));
    this.bounds.l = Math.min.apply(Math, a_x);
    this.bounds.t = Math.min.apply(Math, a_y);
    this.bounds.r = Math.max.apply(Math, a_x);
    this.bounds.b = Math.max.apply(Math, a_y);
    this.bounds.w = this.bounds.r - this.bounds.l;
    this.bounds.h = this.bounds.b - this.bounds.t;
    this.bounds.x = this.bounds.l;
    this.bounds.y = this.bounds.t;
    if(this.drawingBase && !this.group)
    {
        this.drawingBase.checkBoundsFromTo();
    }
};




CChartSpace.prototype.recalculate = function()
{
    if(this.bDeleted && !this.isFrameChart)
        return;

    var oController = this.getDrawingObjectsController();
    //Use this check to prevent charts recalculation on not initialized sheets bug 50467
    if(!oController && !this.isFrameChart) {
        return;
    }
    //---------------------------------------------------
    AscFormat.ExecuteNoHistory(function()
    {
			const bSendUpdateDiagram = this.isNeedUpdateFrameChart();
        this.updateLinks();

        if(this.recalcInfo.recalcTitle)
        {
            this.recalculateChartTitleEditMode();
            this.recalcInfo.recalcTitle.updatePosition(this.transform.tx, this.transform.ty);
            this.recalcInfo.recalcTitle = null;
            this.recalcInfo.bRecalculatedTitle = true;
        }
        var b_transform = false;
        var bCheckLabels = false;
        if(this.recalcInfo.recalculateTransform)
        {
            this.recalculateTransform();
            this.recalculateSnapArrays();
            this.rectGeometry.Recalculate(this.extX, this.extY);
            this.recalcInfo.recalculateTransform = false;
            b_transform = true;
        }
        if(this.recalcInfo.recalculateReferences)
        {
            this.recalculateReferences();
            this.recalcInfo.recalculateReferences = false;
        }
        if(this.recalcInfo.recalculateMarkers)
        {
            this.recalculateMarkers();
            this.recalcInfo.recalculateMarkers = false;
        }
        if(this.recalcInfo.recalculateSeriesColors)
        {
            this.recalculateSeriesColors();
            this.recalcInfo.recalculateSeriesColors = false;
            this.recalcInfo.recalculatePenBrush = true;
        }
        if(this.recalcInfo.recalculateGridLines)
        {
            this.recalculateGridLines();
            this.recalcInfo.recalculateGridLines = false;
        }
        if(this.recalcInfo.recalculateAxisTickMark)
        {
            this.recalculateAxisTickMark();
            this.recalcInfo.recalculateAxisTickMark = false;
        }
        if(this.recalcInfo.recalculateDLbls)
        {
            this.recalculateDLbls();
            this.recalcInfo.recalculateDLbls = false;
        }

        if(this.recalcInfo.recalculateBrush)
        {
            this.recalculateChartBrush();
            this.recalcInfo.recalculateBrush = false;
        }

        if(this.recalcInfo.recalculatePen)
        {
            this.recalculateChartPen();
            this.recalcInfo.recalculatePen = false;
        }

        if(this.recalcInfo.recalculateHiLowLines)
        {
            this.recalculateHiLowLines();
            this.recalcInfo.recalculateHiLowLines = false;
        }
        if(this.recalcInfo.recalculatePlotAreaBrush)
        {
            this.recalculatePlotAreaChartBrush();
            this.recalculateWalls();
            this.recalcInfo.recalculatePlotAreaBrush = false;
        }
        if(this.recalcInfo.recalculatePlotAreaPen)
        {
            this.recalculatePlotAreaChartPen();
            this.recalcInfo.recalculatePlotAreaPen = false;
        }
        if(this.recalcInfo.recalculateUpDownBars)
        {
            this.recalculateUpDownBars();
            this.recalcInfo.recalculateUpDownBars = false;
        }

        var b_recalc_labels = false;
        if(this.recalcInfo.recalculateAxisLabels)
        {
            this.recalculateAxisLabels();
            this.recalcInfo.recalculateAxisLabels = false;
            b_recalc_labels = true;
        }


        var b_recalc_legend = false;
        if(this.recalcInfo.recalculateLegend)
        {
            this.recalculateLegend();
            this.recalcInfo.recalculateLegend = false;
            b_recalc_legend = true;
        }

        if(this.recalcInfo.recalculateAxisVal)
        {
            this.recalculateAxes();
            this.recalcInfo.recalculateAxisVal = false;
            bCheckLabels = true;
        }



        if(this.recalcInfo.recalculatePenBrush)
        {
            this.recalculatePenBrush();
            this.recalcInfo.recalculatePenBrush = false;
        }

        if(this.recalcInfo.recalculateChart)
        {
            this.recalculateChart();
            this.recalcInfo.recalculateChart = false;
            if(bCheckLabels && this.chartObj.nDimensionCount === 3)
            {
                this.checkAxisLabelsTransform();
            }
        }


        this.calculateLabelsPositions(b_recalc_labels, b_recalc_legend);

        if(this.recalcInfo.recalculateBounds)
        {
            this.recalculateBounds();
            this.recalcInfo.recalculateBounds = false;
        }

        if(this.recalcInfo.recalculateTextPr)
        {
            this.recalculateTextPr();
            this.recalcInfo.recalculateTextPr = false;
        }

        this.recalculateUserShapes();
       // if(b_transform)
        {
            this.updateChildLabelsTransform(this.transform.tx, this.transform.ty);
        }
        this.recalcInfo.dataLbls.length = 0;
        this.recalcInfo.axisLabels.length = 0;
        this.bNeedUpdatePosition = true;
			if (bSendUpdateDiagram) {
				const oApi = Asc.editor;
				oApi && oApi.frameManager.setIsNeedUpdateMainDiagram(true);
			}
    }, this, []);
};


CChartSpace.prototype.getDrawingDocument = CShape.prototype.getDrawingDocument;
CChartSpace.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;

CChartSpace.prototype.Get_Theme = CShape.prototype.Get_Theme;
CChartSpace.prototype.Get_ColorMap = CShape.prototype.Get_ColorMap;

    CChartSpace.prototype.Clear_ContentChanges = function(){
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Clear();
        }
    };

    CChartSpace.prototype.Add_ContentChanges = function(Changes){
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Add( Changes );
        }
    };

    CChartSpace.prototype.Refresh_ContentChanges = function(){
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Refresh();
        }
    };

	CChartSpace.prototype.addExternalReferenceToEditor = function (oPastedWb)
	{
		if (!oPastedWb)
			return;

		const oApi = Asc.editor;
		const oWbModel = oApi.wbModel;

		const oMockWb = new AscCommonExcel.Workbook(undefined, undefined, false);
		oMockWb.externalReferences = oPastedWb.externalReferences;
		oMockWb.dependencyFormulas = oPastedWb.dependencyFormulas;
		const oCachedWorksheets = this.getWorksheetsFromCache(oMockWb);
		const oMainExternalReference = oWbModel.addExternalReferenceFromWorksheets(oCachedWorksheets, oPastedWb, oMockWb);
		if (oMainExternalReference) {
			this.convertRefsToExternal(oMainExternalReference, oPastedWb.externalReferences, oMockWb);
		} else if (oPastedWb.externalReferences.length) {
			this.convertRefsToExternal(oMainExternalReference, oPastedWb.externalReferences, oPastedWb);
		}
	}
	CChartSpace.prototype.convertRefsToExternal = function (oMainExternalReference, arrPastedExternalReferences, oMockWb)
	{
		const aSeries = this.getAllSeries();
		if (this.isChartEx()) {
			for (let i = 0; i < aSeries.length; i += 1) {
				const oSeria = aSeries[i];
				let oData = oSeria.getData();
				if(oData) {
					let aDims = oData.dimension;
					for(let nDim = 0; nDim < aDims.length; ++nDim) {
						let oDim = aDims[nDim];
						oDim.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
					}
				}
				if(oSeria.tx && oSeria.tx.txData) {
					oSeria.tx.txData.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
				}
			}
		} else {
			for (let i = 0; i < aSeries.length; i += 1)
			{
				const oSeria = aSeries[i];
				const oVal = oSeria.val || oSeria.yVal;
				if (oVal && oVal.numRef)
				{
					oVal.numRef.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
				}
				const oCat = oSeria.cat || oSeria.xVal;
				if (oCat)
				{
					if (oCat.numRef)
					{
						oCat.numRef.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
					}
					if (oCat.strRef)
					{
						oCat.strRef.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
					}
				}
				if (oSeria.tx && oSeria.tx.strRef)
				{
					oSeria.tx.strRef.updateToExternal(oMainExternalReference, arrPastedExternalReferences, oMockWb);
				}
			}
		}
		this.onDataUpdateRecalc();
	};
	CChartSpace.prototype.applySpecialPasteProps = function (oExternalWb)
	{
		if (oExternalWb && this.canPasteExternal(oExternalWb))
		{
			this.addExternalReferenceToEditor(oExternalWb);
			this.checkChartExRefs();
		}
		if (!this.XLSX || this.XLSX.length)
		{
			this.setXLSX(new Uint8Array(0));
		}
		this.setExternalReference(null);
	};
	CChartSpace.prototype.canPasteExternal = function(oExternalWb) {
		return AscCommonExcel.isAllowPasteLink(oExternalWb);
	};
	CChartSpace.prototype.isWorkbookChart = function () {
		return true;
	};
	CChartSpace.prototype.checkChartExRefs = function() {
		const wbModel = Asc.editor && Asc.editor.wbModel;
		const ws = wbModel && wbModel.getWorksheet(0);
		if (this.isChartEx() && ws) {
			function checkFormula(formula) {
				const formulaText = formula && formula.content;
				if (formulaText) {
					if (!(AscCommon.parserHelp.isName3D(formulaText, 0) || AscCommon.parserHelp.isDefName(formulaText, 0))) {
						const xlchartDefName = wbModel.getNewXlChartId();
						const defName = wbModel.addDefName(xlchartDefName, formulaText, undefined, true);
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_DefinedNamesChange, null, null,
							new AscCommonExcel.UndoRedoData_FromTo(null, defName.getUndoDefName()));
						formula.setContent(xlchartDefName);
					}
				}
			}
			const series = this.getAllSeries();
			for (let i = 0; i < series.length; i += 1) {
				const seria = series[i];
				const data = seria.getData();
				if (data) {
					const dimensions = data.dimension;
					for (let j = 0; j < dimensions.length; j++) {
						const dimension = dimensions[j];
						checkFormula(dimension.f);
					}
				}
				if (seria.tx && seria.tx.txData) {
					checkFormula(seria.tx.txData.f);
				}
			}
		}
	};
	CChartSpace.prototype.isNeedUpdateFrameChart = function() {
		if (this.isFrameChart) {
			for (let recalcProp in this.recalcInfo) {
				if (typeof this.recalcInfo[recalcProp] === "boolean" && this.recalcInfo[recalcProp] ||
					Array.isArray(this.recalcInfo[recalcProp]) && this.recalcInfo[recalcProp].length) {
					return true;
				}
			}
		}
		return false;
	};

})(window);
