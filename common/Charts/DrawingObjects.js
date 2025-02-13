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
var c_oAscCellAnchorType = AscCommon.c_oAscCellAnchorType;
var c_oAscLockTypes = AscCommon.c_oAscLockTypes;
var parserHelp = AscCommon.parserHelp;
var gc_nMaxRow = AscCommon.gc_nMaxRow;
var gc_nMaxCol = AscCommon.gc_nMaxCol;
    var History = AscCommon.History;

var MOVE_DELTA = AscFormat.MOVE_DELTA;

var c_oAscError = Asc.c_oAscError;
var c_oAscChartTypeSettings = Asc.c_oAscChartTypeSettings;
var c_oAscChartTitleShowSettings = Asc.c_oAscChartTitleShowSettings;
var c_oAscGridLinesSettings = Asc.c_oAscGridLinesSettings;
var c_oAscValAxisRule = Asc.c_oAscValAxisRule;
var c_oAscInsertOptions = Asc.c_oAscInsertOptions;
var c_oAscDeleteOptions = Asc.c_oAscDeleteOptions;
var c_oAscSelectionType = Asc.c_oAscSelectionType;
var global_mouseEvent = AscCommon.global_mouseEvent;

var aSparklinesStyles =
[
    [
        [4, -0.499984740745262],
        [5, 0],
        [4, -0.499984740745262],
        [4,  0.39997558519241921],
        [4, 0.39997558519241921],
        [4, 0],
        [4, 0]
    ],
    [
        [5, -0.499984740745262],
        [6, 0],
        [5,  -0.499984740745262],
        [5, 0.39997558519241921],
        [5, 0.39997558519241921],
        [5, 0],
        [5, 0]
    ],
    [
        [6, -0.499984740745262],
        [7, 0],
        [6, -0.499984740745262],
        [6, 0.39997558519241921],
        [6, 0.39997558519241921],
        [6, 0],
        [6, 0]
    ],
    [
        [7, -0.499984740745262],
        [8, 0],
        [7, -0.499984740745262],
        [7, 0.39997558519241921],
        [7, 0.39997558519241921],
        [7, 0],
        [7, 0]
    ],
    [
        [8, -0.499984740745262],
        [9, 0],
        [8, -0.499984740745262],
        [8, 0.39997558519241921],
        [8, 0.39997558519241921],
        [8, 0],
        [8, 0]
    ],
    [
        [9, -0.499984740745262],
        [4, 0],

        [9, -0.499984740745262],
        [9, 0.39997558519241921],
        [9, 0.39997558519241921],
        [9, 0],
        [9, 0]
    ],
    [
        [4, -0.249977111117893],
        [5, 0],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893]
    ],
    [
        [5, -0.249977111117893],
        [6, 0],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893]
    ],
    [
        [6, -0.249977111117893],
        [7, 0],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893]
    ],
    [
        [7, -0.249977111117893],
        [8, 0],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893]
    ],
    [
        [8, -0.249977111117893],
        [9, 0],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893]
    ],
    [
        [9, -0.249977111117893],
        [4, 0],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893]
    ],
    [
        [4, 0],
        [5, 0],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.249977111117893]
    ],
    [
        [5, 0],
        [6, 0],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.249977111117893]
    ],
    [
        [6, 0],
        [7, 0],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.249977111117893]
    ],
    [
        [7, 0],
        [8, 0],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.249977111117893]
    ],
    [
        [8, 0],
        [9, 0],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.249977111117893]
    ],
    [
        [9, 0],
        [4, 0],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.249977111117893]
    ],
    [
        [4, 0.39997558519241921],
        [0, -0.499984740745262],
        [4, 0.79998168889431442],
        [4, -0.249977111117893],
        [4, -0.249977111117893],
        [4, -0.499984740745262],
        [4, -0.499984740745262]
    ],
    [
        [5, 0.39997558519241921],
        [0, -0.499984740745262],
        [5, 0.79998168889431442],
        [5, -0.249977111117893],
        [5, -0.249977111117893],
        [5, -0.499984740745262],
        [5, -0.499984740745262]
    ],
    [
        [6, 0.39997558519241921],
        [0, -0.499984740745262],
        [6, 0.79998168889431442],
        [6, -0.249977111117893],
        [6, -0.249977111117893],
        [6, -0.499984740745262],
        [6, -0.499984740745262]
    ],
    [
        [7, 0.39997558519241921],
        [0, -0.499984740745262],
        [7, 0.79998168889431442],
        [7, -0.249977111117893],
        [7, -0.249977111117893],
        [7, -0.499984740745262],
        [7, -0.499984740745262]
    ],
    [
        [8, 0.39997558519241921],
        [0, -0.499984740745262],
        [8, 0.79998168889431442],
        [8, -0.249977111117893],
        [8, -0.249977111117893],
        [8, -0.499984740745262],
        [8, -0.499984740745262]
    ],
    [
        [9, 0.39997558519241921],
        [0, -0.499984740745262],
        [9, 0.79998168889431442],
        [9, -0.249977111117893],
        [9, -0.249977111117893],
        [9, -0.499984740745262],
        [9, -0.499984740745262]
    ],
    [
        [1, 0.499984740745262],
        [1, 0.249977111117893],
        [1, 0.249977111117893],
        [1, 0.249977111117893],
        [1, 0.249977111117893],
        [1, 0.249977111117893],
        [1, 0.249977111117893]
    ],
    [
        [1, 0.34998626667073579],
        [0, -0.249977111117893],
        [0, -0.249977111117893],
        [0, -0.249977111117893],
        [0, -0.249977111117893],
        [0, -0.249977111117893],
        [0, -0.249977111117893]
    ],
    [
        [0xFF323232],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000]
    ],
    [
        [0xFF000000],
        [0xFF0070C0],
        [0xFF0070C0],
        [0xFF0070C0],
        [0xFF0070C0],
        [0xFF0070C0],
        [0xFF0070C0]
    ],
    [
        [0xFF376092],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000],
        [0xFFD00000]
    ],
    [
        [0xFF0070C0],
        [0xFF000000],
        [0xFF000000],
        [0xFF000000],
        [0xFF000000],
        [0xFF000000],
        [0xFF000000]
    ],
    [
        [0xFF5F5F5F],
        [0xFFFFB620],
        [0xFFD70077],
        [0xFF5687C2],
        [0xFF359CEB],
        [0xFF56BE79],
        [0xFFFF5055]
    ],
    [
        [0xFF5687C2],
        [0xFFFFB620],
        [0xFFD70077],
        [0xFF777777],
        [0xFF359CEB],
        [0xFF56BE79],
        [0xFFFF5055]
    ],
    [
        [0xFFC6EFCE],
        [0xFFFFC7CE],
        [0xFF8CADD6],
        [0xFFFFDC47],
        [0xFFFFEB9C],
        [0xFF60D276],
        [0xFFFF5367]
    ],
    [
        [0xFF00B050],
        [0xFFFF0000],
        [0xFF0070C0],
        [0xFFFFC000],
        [0xFFFFC000],
        [0xFF00B050],
        [0xFFFF0000]
    ],
    [
        [3, 0],
        [9, 0],
        [8, 0],
        [4, 0],
        [5, 0],
        [6, 0],
        [7, 0]
    ],
    [
        [1, 0],
        [9, 0],
        [8, 0],
        [4, 0],
        [5, 0],
        [6, 0],
        [7, 0]
    ]
];

function isObject(what) {
    return ( (what != null) && (typeof(what) == "object") );
}

function DrawingBounds(minX, maxX, minY, maxY)
{
    this.minX = minX;
    this.maxX = maxX;
    this.minY = minY;
    this.maxY = maxY;
}

function getCurrentTime() {
    var currDate = new Date();
    return currDate.getTime();
}

function roundPlus(x, n) { //x - число, n - количество знаков
    if ( isNaN(x) || isNaN(n) ) return false;
    var m = Math.pow(10,n);
    return Math.round(x * m) / m;
}

function CCellObjectInfo () {
	this.col = 0;
	this.row = 0;
	this.colOff = 0;
	this.rowOff = 0;
}
CCellObjectInfo.prototype.initAfterSerialize = function() {
	this.row = Math.max(0, this.row);
	this.col = Math.max(0, this.col);
};

/** @constructor */
function asc_CChartBinary(chart) {

    this["binary"] = null;
    if (chart && chart.getObjectType() === AscDFH.historyitem_type_ChartSpace)
    {
        this["IsChartEx"] = chart.isChartEx();
        var writer = new AscCommon.BinaryChartWriter(new AscCommon.CMemory(false)), pptx_writer;
        if(this["IsChartEx"])
        {
            writer.WriteCT_ChartExSpace(chart);
        }
        else
        {
            writer.WriteCT_ChartSpace(chart);
        }
        this["binary"] = writer.memory.pos + ";" + writer.memory.GetBase64Memory();
        if(chart.theme)
        {
            pptx_writer = new AscCommon.CBinaryFileWriter();
            pptx_writer.WriteTheme(chart.theme);
            this["themeBinary"] = pptx_writer.pos + ";" + pptx_writer.GetBase64Memory();
        }
        if(chart.colorMapOverride)
        {
            pptx_writer = new AscCommon.CBinaryFileWriter();
            pptx_writer.WriteRecord1(1, chart.colorMapOverride, pptx_writer.WriteClrMap);
            this["colorMapBinary"] = pptx_writer.pos + ";" + pptx_writer.GetBase64Memory();
        }
        this["urls"] = JSON.stringify(AscCommon.g_oDocumentUrls.getUrls());
        if(chart.parent && chart.parent.docPr){
            this["cTitle"] = chart.parent.docPr.title;
            this["cDescription"] = chart.parent.docPr.descr;
        }
        else{
            this["cTitle"] = chart.getTitle();
            this["cDescription"] = chart.getDescription();
        }
    }
}

asc_CChartBinary.prototype = {

    asc_getBinary: function() { return this["binary"]; },
    asc_setBinary: function(val) { this["binary"] = val; },
    asc_getThemeBinary: function() { return this["themeBinary"]; },
    asc_setThemeBinary: function(val) { this["themeBinary"] = val; },
    asc_setColorMapBinary: function(val){this["colorMapBinary"] = val;},
    asc_getColorMapBinary: function(){return this["colorMapBinary"];},
    asc_setIsChartEx: function(val){this["IsChartEx"] = val;},
    asc_getIsChartEx: function(){return this["IsChartEx"];},
    getChartSpace: function(workSheet)
    {
        var binary = this["binary"];
        var stream = AscFormat.CreateBinaryReader(this["binary"], 0, this["binary"].length);
        //надо сбросить то, что остался после открытия документа
        AscCommon.pptx_content_loader.Clear();
        var oNewChartSpace = Asc.editor.isPdfEditor() ? new AscPDF.CPdfChart() : new AscFormat.CChartSpace();
        var oBinaryChartReader = new AscCommon.BinaryChartReader(stream);
        if(this["IsChartEx"]) {
            oBinaryChartReader.ExternalReadCT_ChartExSpace(stream.size , oNewChartSpace, workSheet);
        }
        else {
            oBinaryChartReader.ExternalReadCT_ChartSpace(stream.size , oNewChartSpace, workSheet);
        }
        return oNewChartSpace;
    },

    getTheme: function()
    {
        var binary = this["themeBinary"];
        if(binary)
        {
            let oBinaryReader = AscFormat.CreatePPTYLoader(binary, 0, binary.length);
            return oBinaryReader.ReadTheme();
        }
        return null;
    },

    getColorMap: function()
    {
        var binary = this["colorMapBinary"];
        if(binary)
        {
            let oBinaryReader = AscFormat.CreatePPTYLoader(binary, 0, binary.length);
            let _rec = oBinaryReader.stream.GetUChar();
            let ret = new AscFormat.ClrMap();
            oBinaryReader.ReadClrMap(ret);
            return ret;
        }
        return null;
    }

};

/** @constructor */
function asc_CChartSeria() {
    this.Val = { Formula: "", NumCache: [] };
    this.xVal = { Formula: "", NumCache: [] };
    this.Cat = { Formula: "", NumCache: [] };
    this.TxCache = { Formula: "", NumCache: [] };
    this.Marker = { Size: 0, Symbol: "" };
    this.FormatCode = "";
    this.isHidden = false;
}

asc_CChartSeria.prototype = {

    asc_getValFormula: function() { return this.Val.Formula; },
    asc_setValFormula: function(formula) { this.Val.Formula = formula; },

    asc_getxValFormula: function() { return this.xVal.Formula; },
    asc_setxValFormula: function(formula) { this.xVal.Formula = formula; },

    asc_getCatFormula: function() { return this.Cat.Formula; },
    asc_setCatFormula: function(formula) { this.Cat.Formula = formula; },

    asc_getTitle: function() { return this.TxCache.NumCache.length > 0 ? this.TxCache.NumCache[0].val : ""; },
    asc_setTitle: function(title) { this.TxCache.NumCache = [{ numFormatStr: "General", isDateTimeFormat: false, val: title, isHidden: false }] ; },

    asc_getTitleFormula: function() { return this.TxCache.Formula; },
    asc_setTitleFormula: function(val) { this.TxCache.Formula = val; },

    asc_getMarkerSize: function() { return this.Marker.Size; },
    asc_setMarkerSize: function(size) { this.Marker.Size = size; },

    asc_getMarkerSymbol: function() { return this.Marker.Symbol; },
    asc_setMarkerSymbol: function(symbol) { this.Marker.Symbol = symbol; },

    asc_getFormatCode: function() { return this.FormatCode; },
    asc_setFormatCode: function(format) { this.FormatCode = format; }
};

var nSparklineMultiplier = 3.75;

function CSparklineView()
{
    this.col = null;
    this.row = null;
    this.ws = null;
    this.extX = null;
    this.extY = null;
    this.chartSpace = null;
}

function CreateSparklineMarker(oUniFill, bPreview)
{
    var oMarker = new AscFormat.CMarker();
    oMarker.symbol = AscFormat.SYMBOL_DIAMOND;
    oMarker.size = 10;
    if(bPreview){
        oMarker.size = 30;
    }
    oMarker.spPr = new AscFormat.CSpPr();
    oMarker.spPr.Fill = oUniFill;
    oMarker.spPr.ln = AscFormat.CreateNoFillLine();
    return oMarker;
}

function CreateUniFillFromExcelColor(oExcelColor)
{
    var oUnifill;
    /*if(oExcelColor instanceof AscCommonExcel.ThemeColor)
    {

        oUnifill = AscFormat.CreateUnifillSolidFillSchemeColorByIndex(AscCommonExcel.map_themeExcel_to_themePresentation[oExcelColor.theme]);
        if(oExcelColor.tint != null)
        {
            var unicolor = oUnifill.fill.color;
            if(!unicolor.Mods)
                unicolor.setMods(new AscFormat.CColorModifiers());
            var mod = new AscFormat.CColorMod();
            if(oExcelColor.tint > 0)
            {
                mod.setName("wordTint");
                mod.setVal(Math.round(oExcelColor.tint*255));
            }
            else
            {
                mod.setName("wordShade");
                mod.setVal(Math.round(255 + oExcelColor.tint*255));
            }
            unicolor.Mods.addMod(mod);
        }

        //oUnifill = AscFormat.CreateUniFillSchemeColorWidthTint(map_themeExcel_to_themePresentation[oExcelColor.theme], oExcelColor.tint != null ? oExcelColor.tint : 0);
    }
    else*/
    {
        oUnifill = AscFormat.CreateUnfilFromRGB(oExcelColor.getR(), oExcelColor.getG(), oExcelColor.getB())
    }
    return oUnifill;
}

CSparklineView.prototype.initFromSparkline = function(oSparkline, oSparklineGroup, worksheetView, bForPreview)
{
    AscFormat.ExecuteNoHistory(function(){
        this.ws = worksheetView;
        var settings = new Asc.asc_ChartSettings();
        var nSparklineType = oSparklineGroup.asc_getType();
        switch(nSparklineType)
        {
            case Asc.c_oAscSparklineType.Column:
            {
                settings.type = c_oAscChartTypeSettings.barNormal;
                break;
            }
            case Asc.c_oAscSparklineType.Stacked:
            {
                settings.type = c_oAscChartTypeSettings.barStackedPer;
                break;
            }
            default:
            {
                settings.type = c_oAscChartTypeSettings.lineNormal;
                break;
            }
        }
        var ser = new asc_CChartSeria();
        ser.Val.Formula = oSparkline.f;
        if(oSparkline.oCache){
            ser.Val.NumCache = oSparkline.oCache;
        }
        var chartSeries = [ser];
        var chart_space = AscFormat.DrawingObjectsController.prototype._getChartSpace(chartSeries, settings, true);
        chart_space.isSparkline = true;
        chart_space.setBDeleted(false);
        if(worksheetView){
            chart_space.setWorksheet(worksheetView.model);
        }

        chart_space.displayHidden = oSparklineGroup.asc_getDisplayHidden();
        chart_space.displayEmptyCellsAs = oSparklineGroup.asc_getDisplayEmpty();
        settings.putTitle(c_oAscChartTitleShowSettings.none);
        settings.putLegendPos(Asc.c_oAscChartLegendShowSettings.none);


        chart_space.recalculateReferences();
        chart_space.recalcInfo.recalculateReferences = false;
        var oSerie = chart_space.chart.plotArea.charts[0].series[0];
        var aSeriesPoints = oSerie.getNumPts();

        var val_ax_props = new AscCommon.asc_ValAxisSettings();
        var i, fMinVal = null, fMaxVal = null;
        if(settings.type !== c_oAscChartTypeSettings.barStackedPer) {
            if (oSparklineGroup.asc_getMinAxisType() === Asc.c_oAscSparklineAxisMinMax.Custom && oSparklineGroup.asc_getManualMin() !== null) {
                val_ax_props.putMinValRule(c_oAscValAxisRule.fixed);
                val_ax_props.putMinVal(oSparklineGroup.asc_getManualMin());
            }
            else {
                val_ax_props.putMinValRule(c_oAscValAxisRule.auto);

                for (i = 0; i < aSeriesPoints.length; ++i) {
                    if (fMinVal === null) {
                        fMinVal = aSeriesPoints[i].val;
                    }
                    else {
                        if (fMinVal > aSeriesPoints[i].val) {
                            fMinVal = aSeriesPoints[i].val;
                        }
                    }
                }
            }
            if (oSparklineGroup.asc_getMaxAxisType() === Asc.c_oAscSparklineAxisMinMax.Custom && oSparklineGroup.asc_getManualMax() !== null) {
                val_ax_props.putMinValRule(c_oAscValAxisRule.fixed);
                val_ax_props.putMinVal(oSparklineGroup.asc_getManualMax());
            }
            else {
                val_ax_props.putMaxValRule(c_oAscValAxisRule.auto);
                for (i = 0; i < aSeriesPoints.length; ++i) {
                    if (fMaxVal === null) {
                        fMaxVal = aSeriesPoints[i].val;
                    } else {
                        if (fMaxVal < aSeriesPoints[i].val) {
                            fMaxVal = aSeriesPoints[i].val;
                        }
                    }
                }
            }
            if (fMinVal !== null && fMaxVal !== null) {
                if (fMinVal !== fMaxVal) {
                    val_ax_props.putMinValRule(c_oAscValAxisRule.fixed);
                    val_ax_props.putMinVal(fMinVal);
                    val_ax_props.putMaxValRule(c_oAscValAxisRule.fixed);
                    val_ax_props.putMaxVal(fMaxVal);
                }
            }
            else if (fMinVal !== null) {
                val_ax_props.putMinValRule(c_oAscValAxisRule.fixed);
                val_ax_props.putMinVal(fMinVal);
            }
            else if (fMaxVal !== null) {
                val_ax_props.putMaxValRule(c_oAscValAxisRule.fixed);
                val_ax_props.putMaxVal(fMaxVal);
            }
        }
        else
        {
            val_ax_props.putMinValRule(c_oAscValAxisRule.fixed);
            val_ax_props.putMinVal(-1);

            val_ax_props.putMaxValRule(c_oAscValAxisRule.fixed);
            val_ax_props.putMaxVal(1);
        }

        val_ax_props.putTickLabelsPos(Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE);
        val_ax_props.putInvertValOrder(false);
        val_ax_props.putDispUnitsRule(Asc.c_oAscValAxUnits.none);
        val_ax_props.putMajorTickMark(Asc.c_oAscTickMark.TICK_MARK_NONE);
        val_ax_props.putMinorTickMark(Asc.c_oAscTickMark.TICK_MARK_NONE);
        val_ax_props.putCrossesRule(Asc.c_oAscCrossesRule.auto);
        val_ax_props.putLabel(c_oAscChartTitleShowSettings.none);
        val_ax_props.putGridlines(c_oAscGridLinesSettings.none);

        var cat_ax_props = new AscCommon.asc_CatAxisSettings();
        cat_ax_props.putIntervalBetweenLabelsRule(Asc.c_oAscBetweenLabelsRule.auto);
        cat_ax_props.putLabelsPosition(Asc.c_oAscLabelsPosition.betweenDivisions);
        cat_ax_props.putTickLabelsPos(Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NONE);
        cat_ax_props.putLabelsAxisDistance(100);
        cat_ax_props.putMajorTickMark(Asc.c_oAscTickMark.TICK_MARK_NONE);
        cat_ax_props.putMinorTickMark(Asc.c_oAscTickMark.TICK_MARK_NONE);
        cat_ax_props.putIntervalBetweenTick(1);
        cat_ax_props.putCrossesRule(Asc.c_oAscCrossesRule.auto);
        cat_ax_props.putLabel(c_oAscChartTitleShowSettings.none);
        cat_ax_props.putGridlines(c_oAscGridLinesSettings.none);
        if(oSparklineGroup.rightToLeft)
        {
            cat_ax_props.putInvertCatOrder(true);
        }
        settings.addVertAxesProps(val_ax_props);
        settings.addHorAxesProps(cat_ax_props);

        AscFormat.DrawingObjectsController.prototype.applyPropsToChartSpace(settings, chart_space);

        oSerie = chart_space.chart.plotArea.charts[0].series[0];
        aSeriesPoints = oSerie.getNumPts();
        if(!chart_space.spPr)
            chart_space.setSpPr(new AscFormat.CSpPr());

        var new_line = new AscFormat.CLn();
        new_line.setFill(AscFormat.CreateNoFillUniFill());
        chart_space.spPr.setLn(new_line);
        chart_space.spPr.setFill(AscFormat.CreateNoFillUniFill());
        var dLineWidthSpaces = 500;
        if(!chart_space.chart.plotArea.spPr)
        {
            chart_space.chart.plotArea.setSpPr(new AscFormat.CSpPr());
            chart_space.chart.plotArea.spPr.setFill(AscFormat.CreateNoFillUniFill());
        }
        var oAxis = chart_space.chart.plotArea.getAxisByTypes();
        oAxis.valAx[0].setDelete(true);

        if(!oSerie.spPr)
        {
            oSerie.setSpPr(new AscFormat.CSpPr());
        }
        var fCallbackSeries = null;
        if(nSparklineType === Asc.c_oAscSparklineType.Line)
        {
            var oLn = new AscFormat.CLn();
            oLn.setW(36000*nSparklineMultiplier*25.4*(bForPreview ? 2.25 : oSparklineGroup.asc_getLineWeight())/72);
            oSerie.spPr.setLn(oLn);
            if(oSparklineGroup.markers && oSparklineGroup.colorMarkers)
            {
                chart_space.chart.plotArea.charts[0].setMarker(true);
                chart_space.recalcInfo.recalculateReferences = false;
                oSerie.marker = CreateSparklineMarker(CreateUniFillFromExcelColor(oSparklineGroup.colorMarkers), bForPreview);
            }

            fCallbackSeries = function(oSeries, nIdx, oExcelColor)
            {
                for(var t = 0; t < oSeries.dPt.length; ++t)
                {
                    if(oSeries.dPt[t].idx === nIdx)
                    {
                        if(oSeries.dPt[t].marker && oSeries.dPt[t].marker.spPr)
                        {
                            oSeries.dPt[t].marker.spPr.Fill = CreateUniFillFromExcelColor(oExcelColor);
                        }
                        return;
                    }
                }
                var oDPt = new AscFormat.CDPt();
                oDPt.idx = nIdx;
                oDPt.marker = CreateSparklineMarker(CreateUniFillFromExcelColor(oExcelColor), bForPreview);
                oSeries.addDPt(oDPt);
            }
        }
        else
        {
            chart_space.chart.plotArea.charts[0].setGapWidth(30);
            chart_space.chart.plotArea.charts[0].setOverlap(50);
            fCallbackSeries = function(oSeries, nIdx, oExcelColor)
            {
                for(var t = 0; t < oSeries.dPt.length; ++t)
                {
                    if(oSeries.dPt[t].idx === nIdx)
                    {
                        if(oSeries.dPt[t].spPr)
                        {
                            if(oExcelColor){
                                oSeries.dPt[t].spPr.Fill = CreateUniFillFromExcelColor(oExcelColor);
                                oSeries.dPt[t].spPr.ln = new AscFormat.CLn();
                                oSeries.dPt[t].spPr.ln.Fill = oSeries.dPt[t].spPr.Fill.createDuplicate();
                                oSeries.dPt[t].spPr.ln.w = dLineWidthSpaces;
                            }
                            else{
                                oSeries.dPt[t].spPr.Fill = AscFormat.CreateNoFillUniFill();
                                oSeries.dPt[t].spPr.ln = AscFormat.CreateNoFillLine();
                            }
                        }
                        return;
                    }
                }

                var oDPt = new AscFormat.CDPt();
                oDPt.idx = nIdx;
                oDPt.spPr = new AscFormat.CSpPr();
                if(oExcelColor ) {
                    oDPt.spPr.Fill = CreateUniFillFromExcelColor(oExcelColor);
                    oDPt.spPr.ln = new AscFormat.CLn();
                    oDPt.spPr.ln.Fill = oDPt.spPr.Fill.createDuplicate();
                    oDPt.spPr.ln.w = dLineWidthSpaces;
                }
                else{
                    oDPt.spPr.Fill = AscFormat.CreateNoFillUniFill();
                    oDPt.spPr.ln = AscFormat.CreateNoFillLine();
                }
                oSeries.addDPt(oDPt);
            }
        }
        var aMaxPoints = null, aMinPoints = null;
        if(aSeriesPoints.length > 0)
        {
            if(fCallbackSeries)
            {
                if(oSparklineGroup.negative && oSparklineGroup.colorNegative)
                {
                    for(i = 0; i < aSeriesPoints.length; ++i)
                    {
                        if(aSeriesPoints[i].val < 0)
                        {
                            fCallbackSeries(oSerie, aSeriesPoints[i].idx, oSparklineGroup.colorNegative);
                        }
                    }
                }
                if(oSparklineGroup.last && oSparklineGroup.colorLast)
                {
                    fCallbackSeries(oSerie, aSeriesPoints[aSeriesPoints.length - 1].idx, oSparklineGroup.colorLast);
                }
                if(oSparklineGroup.first && oSparklineGroup.colorFirst)
                {
                    fCallbackSeries(oSerie, aSeriesPoints[0].idx, oSparklineGroup.colorFirst);
                }
                if(oSparklineGroup.high && oSparklineGroup.colorHigh)
                {
                    aMaxPoints = [aSeriesPoints[0]];
                    for(i = 1; i < aSeriesPoints.length; ++i)
                    {
                        if(aSeriesPoints[i].val > aMaxPoints[0].val)
                        {
                            aMaxPoints.length = 0;
                            aMaxPoints.push(aSeriesPoints[i]);
                        }
                        else if(aSeriesPoints[i].val === aMaxPoints[0].val)
                        {
                            aMaxPoints.push(aSeriesPoints[i]);
                        }
                    }
                    for(i = 0; i < aMaxPoints.length; ++i)
                    {
                        fCallbackSeries(oSerie, aMaxPoints[i].idx, oSparklineGroup.colorHigh);
                    }
                }
                if(oSparklineGroup.low && oSparklineGroup.colorLow)
                {
                    aMinPoints = [aSeriesPoints[0]];
                    for(i = 1; i < aSeriesPoints.length; ++i)
                    {
                        if(aSeriesPoints[i].val < aMinPoints[0].val)
                        {
                            aMinPoints.length = 0;
                            aMinPoints.push(aSeriesPoints[i]);
                        }
                        else if(aSeriesPoints[i].val === aMinPoints[0].val)
                        {
                            aMinPoints.push(aSeriesPoints[i]);
                        }
                    }
                    for(i = 0; i < aMinPoints.length; ++i)
                    {
                        fCallbackSeries(oSerie, aMinPoints[i].idx, oSparklineGroup.colorLow);
                    }
                }


                if(nSparklineType !== Asc.c_oAscSparklineType.Line){
                    for(i = 0; i < aSeriesPoints.length; ++i)
                    {
                        if(AscFormat.fApproxEqual(aSeriesPoints[i].val,  0))
                        {
                            fCallbackSeries(oSerie, aSeriesPoints[i].idx, null);
                        }
                    }
                }
            }
        }
        if(!oSparklineGroup.displayXAxis)
        {
            oAxis.catAx[0].setDelete(true);
        }
        else if(aSeriesPoints.length > 1)
        {
            aSeriesPoints = [].concat(aSeriesPoints);
            var dMinVal, dMaxVal, bSorted = false;
            if(val_ax_props.minVal == null)
            {
                if(aMinPoints)
                {
                    dMinVal = aMinPoints[0].val;
                }
                else
                {

                    aSeriesPoints.sort(function(a, b){return a.val - b.val;});
                    bSorted = true;
                    dMinVal = aSeriesPoints[0].val;
                }
            }
            else
            {
                dMinVal = val_ax_props.minVal;
            }

            if(val_ax_props.maxVal == null)
            {
                if(aMaxPoints)
                {
                    dMaxVal = aMaxPoints[0].val;
                }
                else
                {
                    if(!bSorted)
                    {
                        aSeriesPoints.sort(function(a, b){return a.val - b.val;});
                        bSorted = true;
                    }
                    dMaxVal = aSeriesPoints[aSeriesPoints.length - 1].val;
                }
            }
            else
            {
                dMaxVal = val_ax_props.maxVal;
            }
            if((dMaxVal !== dMinVal) && (dMaxVal < 0 || dMinVal > 0))
            {
                oAxis.catAx[0].setDelete(true);
            }
        }
        if(oSparklineGroup.colorSeries)
        {
            var oUnifill = CreateUniFillFromExcelColor(oSparklineGroup.colorSeries);
            var oSerie = chart_space.chart.plotArea.charts[0].series[0];

            //oSerie.setSpPr(new AscFormat.CSpPr());
            if(nSparklineType === Asc.c_oAscSparklineType.Line)
            {
                var oLn = oSerie.spPr.ln;
                oLn.setFill(oUnifill);
                oSerie.spPr.setLn(oLn);
            }
            else
            {
                    oSerie.spPr.setFill(oUnifill);
                    oSerie.spPr.ln = new AscFormat.CLn();
                    oSerie.spPr.ln.Fill = oSerie.spPr.Fill.createDuplicate();
                    oSerie.spPr.ln.w = dLineWidthSpaces;
            }
        }
        this.chartSpace = chart_space;
        if(worksheetView)
        {
            var oBBox = oSparkline.sqRef;
            this.col = oBBox.c1;
            this.row = oBBox.r1;
            this.x = worksheetView.getCellLeft(oBBox.c1, 3);
            this.y = worksheetView.getCellTop(oBBox.r1, 3);
            var oMergeInfo = worksheetView.model.getMergedByCell( oBBox.r1, oBBox.c1 );
            if(oMergeInfo)
            {
                this.extX = 0;
                for(i = oMergeInfo.c1; i <= oMergeInfo.c2; ++i)
                {
                    this.extX += worksheetView.getColumnWidth(i, 3)
                }
                this.extY = 0;
                for(i = oMergeInfo.r1; i <= oMergeInfo.r2; ++i)
                {
                    this.extY = worksheetView.getRowHeight(i, 3);
                }
            }
            else
            {
                this.extX = worksheetView.getColumnWidth(oBBox.c1, 3);
                this.extY = worksheetView.getRowHeight(oBBox.r1, 3);
            }
            AscFormat.CheckSpPrXfrm(this.chartSpace);
            this.updatePlotAreaLayout();
            this.chartSpace.recalculate();
        }
    }, this, []);
};

CSparklineView.prototype.updatePlotAreaLayout = function()
{
    if(!this.chartSpace)
    {
        return;
    }
    var offX = this.x*nSparklineMultiplier;
    var offY = this.y*nSparklineMultiplier;
    var extX = this.extX*nSparklineMultiplier;
    var extY = this.extY*nSparklineMultiplier;
    this.chartSpace.spPr.xfrm.setOffX(offX);
    this.chartSpace.spPr.xfrm.setOffY(offY);
    this.chartSpace.spPr.xfrm.setExtX(extX);
    this.chartSpace.spPr.xfrm.setExtY(extY);
    var oLayout = new AscFormat.CLayout();
    oLayout.setXMode(AscFormat.LAYOUT_MODE_EDGE);
    oLayout.setYMode(AscFormat.LAYOUT_MODE_EDGE);
    oLayout.setLayoutTarget(AscFormat.LAYOUT_TARGET_INNER);
    var fInset = 2.0;
    var fPosX, fPosY, fExtX, fExtY;
    fExtX = extX - 2*fInset;
    fExtY = extY - 2*fInset;
    this.chartSpace.bEmptySeries = false;
    if(fExtX <= 0.0 || fExtY <= 0.0)
    {
        this.chartSpace.bEmptySeries = true;
        return;
    }
    fPosX = (extX - fExtX) / 2.0;
    fPosY = (extY - fExtY) / 2.0;
    var fLayoutX = this.chartSpace.calculateLayoutByPos(0, oLayout.xMode, fPosX, extX);
    var fLayoutY = this.chartSpace.calculateLayoutByPos(0, oLayout.yMode, fPosY, extY);
    var fLayoutW = this.chartSpace.calculateLayoutBySize(fPosX, oLayout.wMode, extX, fExtX);
    var fLayoutH = this.chartSpace.calculateLayoutBySize(fPosY, oLayout.hMode, extY, fExtY);
    oLayout.setX(fLayoutX);
    oLayout.setY(fLayoutY);
    oLayout.setW(fLayoutW);
    oLayout.setH(fLayoutH);
    this.chartSpace.chart.plotArea.setLayout(oLayout);
};

CSparklineView.prototype.draw = function(graphics, offX, offY)
{
    var x = this.ws.getCellLeft(this.col, 3) - offX;
    var y = this.ws.getCellTop(this.row, 3) - offY;

    x = this.ws.checkRtl(x, undefined, 3);

    var i;

    var extX;
    var extY;
    var oMergeInfo = this.ws.model.getMergedByCell( this.row, this.col );
    if(oMergeInfo){
        extX = 0;
        for(i = oMergeInfo.c1; i <= oMergeInfo.c2; ++i){
            extX += this.ws.getColumnWidth(i, 3);
        }
        extY = 0;
        for(i = oMergeInfo.r1; i <= oMergeInfo.r2; ++i){
            extY = this.ws.getRowHeight(i, 3);
        }
    }
    else{
        extX = this.ws.getColumnWidth(this.col, 3);
        extY = this.ws.getRowHeight(this.row, 3);
    }

    if (this.ws.getRightToLeft()) {
        x -= extX;
    }

    var bExtent = Math.abs(this.extX - extX) > 0.01 || Math.abs(this.extY - extY) > 0.01;
    var bPosition = Math.abs(this.x - x) > 0.01 || Math.abs(this.y - y) > 0.01;
    if(bExtent || bPosition)
    {
        this.x = x;
        this.y = y;
        this.extX = extX;
        this.extY = extY;
        AscFormat.ExecuteNoHistory(function(){
            if(bPosition)
            {
                this.chartSpace.spPr.xfrm.setOffX(x*nSparklineMultiplier);
                this.chartSpace.spPr.xfrm.setOffY(y*nSparklineMultiplier);
            }
            if(bExtent)
            {
                this.chartSpace.spPr.xfrm.setExtX(extX*nSparklineMultiplier);
                this.chartSpace.spPr.xfrm.setExtY(extY*nSparklineMultiplier);
                this.updatePlotAreaLayout();
            }
        }, this, []);
        if(bExtent)
        {
            this.chartSpace.handleUpdateExtents();
            this.chartSpace.recalculate();
        }
        else
        {
            this.chartSpace.x = x*nSparklineMultiplier;
            this.chartSpace.y = y*nSparklineMultiplier;
            this.chartSpace.transform.tx = this.chartSpace.x;
            this.chartSpace.transform.ty = this.chartSpace.y;
        }
    }

    var _true_height = this.chartSpace.chartObj.calcProp.trueHeight;
    var _true_width = this.chartSpace.chartObj.calcProp.trueWidth;

	this.chartSpace.chartObj.calcProp.trueWidth = this.chartSpace.extX * this.chartSpace.chartObj.calcProp.pxToMM;
	this.chartSpace.chartObj.calcProp.trueHeight = this.chartSpace.extY * this.chartSpace.chartObj.calcProp.pxToMM;

    this.chartSpace.draw(graphics);

	this.chartSpace.chartObj.calcProp.trueWidth = _true_width;
	this.chartSpace.chartObj.calcProp.trueHeight = _true_height;
};


CSparklineView.prototype.setMinMaxValAx = function(minVal, maxVal, oSparklineGroup)
{
    var oAxis = this.chartSpace.chart.plotArea.getAxisByTypes();
    var oValAx = oAxis.valAx[0];
    if(oValAx)
    {
        if(minVal !== null)
        {
            if(!oValAx.scaling)
            {
                oValAx.setScaling(new AscFormat.CScaling());
            }
            if(oValAx.scaling.min === null || !AscFormat.fApproxEqual(oValAx.scaling.min, minVal))
            {
                oValAx.scaling.setMin(minVal);
            }
        }
        if(maxVal !== null)
        {
            if(!oValAx.scaling)
            {
                oValAx.setScaling(new AscFormat.CScaling());
            }
            if(oValAx.scaling.max === null || !AscFormat.fApproxEqual(oValAx.scaling.max, maxVal))
            {
                oValAx.scaling.setMax(maxVal);
            }
        }


        if(oSparklineGroup.displayXAxis)
        {
             var  aSeriesPoints = this.chartSpace.chart.plotArea.charts[0].series[0].getNumPts();
            if(aSeriesPoints.length > 1)
            {
                aSeriesPoints = [].concat(aSeriesPoints);
                var dMinVal, dMaxVal, bSorted = false;
                if(minVal == null)
                {
                    aSeriesPoints.sort(function(a, b){return a.val - b.val;});
                    bSorted = true;
                    dMinVal = aSeriesPoints[0].val;
                }
                else
                {
                    dMinVal = minVal;
                }

                if(maxVal == null)
                {
                    if(!bSorted)
                    {
                        aSeriesPoints.sort(function(a, b){return a.val - b.val;});
                        bSorted = true;
                    }
                    dMaxVal = aSeriesPoints[aSeriesPoints.length - 1].val;
                }
                else
                {
                    dMaxVal = maxVal;
                }
                if(dMaxVal < 0 || dMinVal > 0)
                {
                    oAxis.catAx[0].setDelete(true);
                }
                else
                {
                    oAxis.catAx[0].setDelete(false);
                }
            }
        }
        this.chartSpace.recalculate();
    }
};
//-----------------------------------------------------------------------------------
// Manager
//-----------------------------------------------------------------------------------



    var rAF = (function() {
        return window.requestAnimationFrame ||
            window.webkitRequestAnimationFrame ||
            window.mozRequestAnimationFrame ||
            window.oRequestAnimationFrame ||
            window.msRequestAnimationFrame ||
            function(callback) { return setTimeout(callback, 1000/ 60); };
    })();

    var cAF = (function () {
        return window.cancelAnimationFrame ||
            window.cancelRequestAnimationFrame ||
            window.webkitCancelAnimationFrame ||
            window.webkitCancelRequestAnimationFrame ||
            window.mozCancelRequestAnimationFrame ||
            window.oCancelRequestAnimationFrame ||
            window.msCancelRequestAnimationFrame ||
            clearTimeout;
    })();


    function DrawingBase(ws) {
        this.worksheet = ws;

        this.Type = c_oAscCellAnchorType.cellanchorTwoCell;
        this.Pos = { X: 0, Y: 0 };

        this.editAs = c_oAscCellAnchorType.cellanchorTwoCell;
        this.from = new CCellObjectInfo();
        this.to = new CCellObjectInfo();
        this.ext = { cx: 0, cy: 0 };

        this.graphicObject = null; // CImage, CShape, GroupShape or CChartAsGroup

        this.boundsFromTo =
            {
                from: new CCellObjectInfo(),
                to  : new CCellObjectInfo()
            };
    }

    //{ prototype


    DrawingBase.prototype.isUseInDocument = function() {
        if(this.worksheet && this.worksheet.model){
            var aDrawings = this.worksheet.model.Drawings;
            for(var i = 0; i < aDrawings.length; ++i){
                if(aDrawings[i] === this){
                    return true;
                }
            }
        }
        return false;
    };

    DrawingBase.prototype.getAllFonts = function(AllFonts) {
        var _t = this;
        _t.graphicObject && _t.graphicObject.documentGetAllFontNames && _t.graphicObject.documentGetAllFontNames(AllFonts);
    };

    DrawingBase.prototype.isImage = function() {
        var _t = this;
        return _t.graphicObject ? _t.graphicObject.isImage() : false;
    };

    DrawingBase.prototype.isShape = function() {
        var _t = this;
        return _t.graphicObject ? _t.graphicObject.isShape() : false;
    };

    DrawingBase.prototype.isGroup = function() {
        var _t = this;
        return _t.graphicObject ? _t.graphicObject.isGroup() : false;
    };

    DrawingBase.prototype.isChart = function() {
        var _t = this;
        return _t.graphicObject ? _t.graphicObject.isChart() : false;
    };

    DrawingBase.prototype.isGraphicObject = function() {
        var _t = this;
        return _t.graphicObject != null;
    };

    DrawingBase.prototype.isLocked = function() {
        var _t = this;
        return ( (_t.graphicObject.lockType != c_oAscLockTypes.kLockTypeNone) && (_t.graphicObject.lockType != c_oAscLockTypes.kLockTypeMine) )
    };

    DrawingBase.prototype.getCanvasContext = function() {
        return this.getDrawingObjects().drawingDocument.CanvasHitContext;
    };

    DrawingBase.prototype.pxToMm = function(val) {
        let oDrawingObjects = this.getDrawingObjects();
        if(oDrawingObjects) {
            return oDrawingObjects.pxToMm(val);
        }
        return 0;
    };
    DrawingBase.prototype.mmToPx = function(val) {
        let oDrawingObjects = this.getDrawingObjects();
        if(oDrawingObjects) {
            return oDrawingObjects.mmToPx(val);
        }
        return 0;
    };

    // GraphicObject: x, y, extX, extY
    DrawingBase.prototype.getGraphicObjectMetrics = function() {
        var _t = this;
        var metrics = { x: 0, y: 0, extX: 0, extY: 0 };

        var coordsFrom, coordsTo;
        switch(_t.Type)
        {
            case c_oAscCellAnchorType.cellanchorAbsolute:
            {
                metrics.x = this.Pos.X;
                metrics.y = this.Pos.Y;
                metrics.extX = this.ext.cx;
                metrics.extY = this.ext.cy;
                break;
            }
            case c_oAscCellAnchorType.cellanchorOneCell:
            {
                if (this.worksheet) {
                    coordsFrom = this.getDrawingObjects().calculateCoords(_t.from);
                    metrics.x = this.pxToMm( coordsFrom.x );
                    metrics.y = this.pxToMm( coordsFrom.y );
                    metrics.extX = this.ext.cx;
                    metrics.extY = this.ext.cy;
                    if (this.worksheet.getRightToLeft()) {
                        metrics.x -= metrics.extX;
                    }
                }
                break;
            }
            case c_oAscCellAnchorType.cellanchorTwoCell:
            {
                if (this.worksheet) {
                    coordsFrom = this.getDrawingObjects().calculateCoords(_t.from);
                    if (!this.worksheet.getRightToLeft()) {
                        metrics.x = this.pxToMm( coordsFrom.x );
                        metrics.y = this.pxToMm( coordsFrom.y );
                        coordsTo = this.getDrawingObjects().calculateCoords(_t.to);
                        metrics.extX = this.pxToMm( coordsTo.x - coordsFrom.x );
                        metrics.extY = this.pxToMm( coordsTo.y - coordsFrom.y );
                    } else {
                        metrics.y = this.pxToMm( coordsFrom.y );

                        coordsTo = this.getDrawingObjects().calculateCoords(_t.to);
                        metrics.extX = this.pxToMm(  coordsFrom.x - coordsTo.x );
                        metrics.extY = this.pxToMm( coordsTo.y - coordsFrom.y );

                        metrics.x = this.pxToMm( coordsTo.x );
                    }
                }
                break;
            }
        }
        if(metrics.extX < 0)
        {
            metrics.extX = 0;
        }
        if(metrics.extY < 0)
        {
            metrics.extY = 0;
        }
        return metrics;
    };

    // Считаем From/To исходя из graphicObject


    DrawingBase.prototype._getGraphicObjectCoords = function()
    {
        var _t = this;

        if ( _t.isGraphicObject() ) {
            var ret = {Pos:{}, ext: {}, from: {}, to: {}};
            var rot = AscFormat.isRealNumber(_t.graphicObject.rot) ? _t.graphicObject.rot : 0;
            rot = AscFormat.normalizeRotate(rot);

            var fromX, fromY, toX, toY;
            if (AscFormat.checkNormalRotate(rot))
            {
                fromX = this.mmToPx(_t.graphicObject.x);
                fromY = this.mmToPx(_t.graphicObject.y);
                toX = this.mmToPx(_t.graphicObject.x + _t.graphicObject.extX);
                toY = this.mmToPx(_t.graphicObject.y + _t.graphicObject.extY);
                ret.Pos.X = _t.graphicObject.x;
                ret.Pos.Y = _t.graphicObject.y;
                ret.ext.cx = _t.graphicObject.extX;
                ret.ext.cy = _t.graphicObject.extY;
            }
            else
            {
                var _xc, _yc;
                _xc = _t.graphicObject.x + _t.graphicObject.extX/2;
                _yc = _t.graphicObject.y + _t.graphicObject.extY/2;
                fromX =  this.mmToPx(_xc - _t.graphicObject.extY/2);
                fromY =  this.mmToPx(_yc - _t.graphicObject.extX/2);
                toX = this.mmToPx(_xc + _t.graphicObject.extY/2);
                toY = this.mmToPx(_yc + _t.graphicObject.extX/2);
                ret.Pos.X = _xc - _t.graphicObject.extY/2;
                ret.Pos.Y = _yc - _t.graphicObject.extX/2;
                ret.ext.cx = _t.graphicObject.extY;
                ret.ext.cy = _t.graphicObject.extX;
            }

            if (this.worksheet.getRightToLeft()) {
                let temp = fromX;
                fromX = toX;
                toX = temp;
            }

            var fromColCell = this.worksheet.findCellByXY(fromX, fromY, true, false, true);
            var fromRowCell = this.worksheet.findCellByXY(fromX, fromY, true, true, false);
            var toColCell = this.worksheet.findCellByXY(toX, toY, true, false, true);
            var toRowCell = this.worksheet.findCellByXY(toX, toY, true, true, false);

            ret.from.col = fromColCell.col;
            ret.from.colOff = this.pxToMm(fromColCell.colOff);
            ret.from.row = fromRowCell.row;
            ret.from.rowOff = this.pxToMm(fromRowCell.rowOff);

            ret.to.col = toColCell.col;
            ret.to.colOff = this.pxToMm(toColCell.colOff);
            ret.to.row = toRowCell.row;
            ret.to.rowOff = this.pxToMm(toRowCell.rowOff);
            return ret;
        }
        return null;
    };

    DrawingBase.prototype.setGraphicObjectCoords = function() {
        var _t = this;
        var oCoords = this._getGraphicObjectCoords();
        if(oCoords)
        {
            this.Pos.X = oCoords.Pos.X;
            this.Pos.Y = oCoords.Pos.Y;
            this.ext.cx = oCoords.ext.cx;
            this.ext.cy = oCoords.ext.cy;
            this.from.col = oCoords.from.col;
            this.from.colOff = oCoords.from.colOff;
            this.from.row = oCoords.from.row;
            this.from.rowOff = oCoords.from.rowOff;
            this.to.col = oCoords.to.col;
            this.to.colOff = oCoords.to.colOff;
            this.to.row = oCoords.to.row;
            this.to.rowOff = oCoords.to.rowOff;
        }
        if ( _t.isGraphicObject() ) {

            var rot = AscFormat.isRealNumber(_t.graphicObject.rot) ? _t.graphicObject.rot : 0;
            rot = AscFormat.normalizeRotate(rot);

            var fromX, fromY, toX, toY;
            if (AscFormat.checkNormalRotate(rot))
            {
                fromX = this.mmToPx(_t.graphicObject.x);
                fromY = this.mmToPx(_t.graphicObject.y);
                toX = this.mmToPx(_t.graphicObject.x + _t.graphicObject.extX);
                toY = this.mmToPx(_t.graphicObject.y + _t.graphicObject.extY);
                this.Pos.X = _t.graphicObject.x;
                this.Pos.Y = _t.graphicObject.y;
                this.ext.cx = _t.graphicObject.extX;
                this.ext.cy = _t.graphicObject.extY;
            }
            else
            {
                var _xc, _yc;
                _xc = _t.graphicObject.x + _t.graphicObject.extX/2;
                _yc = _t.graphicObject.y + _t.graphicObject.extY/2;
                fromX = this.mmToPx(_xc - _t.graphicObject.extY/2);
                fromY = this.mmToPx(_yc - _t.graphicObject.extX/2);
                toX = this.mmToPx(_xc + _t.graphicObject.extY/2);
                toY = this.mmToPx(_yc + _t.graphicObject.extX/2);
                this.Pos.X = _xc - _t.graphicObject.extY/2;
                this.Pos.Y = _yc - _t.graphicObject.extX/2;
                this.ext.cx = _t.graphicObject.extY;
                this.ext.cy = _t.graphicObject.extX;
            }

            if (this.worksheet.getRightToLeft()) {
                let temp = fromX;
                fromX = toX;
                toX = temp;
            }

            var fromColCell = this.worksheet.findCellByXY(fromX, fromY, true, false, true);
            var fromRowCell = this.worksheet.findCellByXY(fromX, fromY, true, true, false);
            var toColCell = this.worksheet.findCellByXY(toX, toY, true, false, true);
            var toRowCell = this.worksheet.findCellByXY(toX, toY, true, true, false);

            _t.from.col = fromColCell.col;
            _t.from.colOff = this.pxToMm(fromColCell.colOff);
            _t.from.row = fromRowCell.row;
            _t.from.rowOff = this.pxToMm(fromRowCell.rowOff);

            _t.to.col = toColCell.col;
            _t.to.colOff = this.pxToMm(toColCell.colOff);
            _t.to.row = toRowCell.row;
            _t.to.rowOff = this.pxToMm(toRowCell.rowOff);
        }
    };

    DrawingBase.prototype.checkBoundsFromTo = function() {
        var _t = this;

        if ( _t.isGraphicObject() && _t.graphicObject.bounds) {


            var bounds = _t.graphicObject.bounds;


            var fromX =  this.mmToPx(bounds.x > 0 ? bounds.x : 0), fromY =  this.mmToPx(bounds.y > 0 ? bounds.y : 0),
                toX = this.mmToPx(bounds.x + bounds.w), toY = this.mmToPx(bounds.y + bounds.h);
            if(toX < 0)
            {
                toX = 0;
            }
            if(toY < 0)
            {
                toY = 0;
            }

            if (this.worksheet.getRightToLeft()) {
                let temp = fromX;
                fromX = toX;
                toX = temp;
            }

            var fromColCell = this.worksheet.findCellByXY(fromX, fromY, true, false, true);
            var fromRowCell = this.worksheet.findCellByXY(fromX, fromY, true, true, false);
            var toColCell = this.worksheet.findCellByXY(toX, toY, true, false, true);
            var toRowCell = this.worksheet.findCellByXY(toX, toY, true, true, false);

            _t.boundsFromTo.from.col = fromColCell.col;
            _t.boundsFromTo.from.colOff = this.pxToMm(fromColCell.colOff);
            _t.boundsFromTo.from.row = fromRowCell.row;
            _t.boundsFromTo.from.rowOff = this.pxToMm(fromRowCell.rowOff);

            _t.boundsFromTo.to.col = toColCell.col;
            _t.boundsFromTo.to.colOff = this.pxToMm(toColCell.colOff);
            _t.boundsFromTo.to.row = toRowCell.row;
            _t.boundsFromTo.to.rowOff = this.pxToMm(toRowCell.rowOff);
        }
    };

    // Реальное смещение по высоте
    DrawingBase.prototype.getRealTopOffset = function() {
        var _t = this;
        var val = _t.worksheet._getRowTop(_t.from.row) + this.mmToPx(_t.from.rowOff);
        return window["Asc"].round(val);
    };

    // Реальное смещение по ширине
    DrawingBase.prototype.getRealLeftOffset = function() {
        var _t = this;
        var val = _t.worksheet._getColLeft(_t.from.col) + this.mmToPx(_t.from.colOff);
        return window["Asc"].round(val);
    };

    // Ширина по координатам
    DrawingBase.prototype.getWidthFromTo = function() {
        return (this.worksheet._getColLeft(this.to.col) + this.mmToPx(this.to.colOff) -
            this.worksheet._getColLeft(this.from.col) - this.mmToPx(this.from.colOff));
    };

    // Высота по координатам
    DrawingBase.prototype.getHeightFromTo = function() {
        return this.worksheet._getRowTop(this.to.row) + this.mmToPx(this.to.rowOff) -
            this.worksheet._getRowTop(this.from.row) - this.mmToPx(this.from.rowOff);
    };

    // Видимое смещение объекта от первой видимой строки
    DrawingBase.prototype.getVisibleTopOffset = function(withHeader) {
        var _t = this;
        var headerRowOff = _t.worksheet._getRowTop(0);
        var fvr = _t.worksheet._getRowTop(_t.worksheet.getFirstVisibleRow(true));
        var off = _t.getRealTopOffset() - fvr;
        off = (off > 0) ? off : 0;
        return withHeader ? headerRowOff + off : off;
    };

    // Видимое смещение объекта от первой видимой колонки
    DrawingBase.prototype.getVisibleLeftOffset = function(withHeader) {
        var _t = this;
        var headerColOff = _t.worksheet._getColLeft(0);
        var fvc = _t.worksheet._getColLeft(_t.worksheet.getFirstVisibleCol(true));
        var off = _t.getRealLeftOffset() - fvc;
        off = (off > 0) ? off : 0;
        return withHeader ? headerColOff + off : off;
    };
    DrawingBase.prototype.getDrawingObjects = function() {
        return this.worksheet && this.worksheet.objectRender;
    };
    DrawingBase.prototype.checkTarget = function(target, bEdit) {
        if(!this.graphicObject) {
            return false;
        }
        if(AscCommon.isFileBuild()) {
            return false;
        }
        var bUpdateExtents = false;
        var nType = bEdit ? this.graphicObject.getDrawingBaseType() : this.Type;
        if(target.target === AscCommonExcel.c_oTargetType.RowResize) {
            if(nType === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell ||
                nType === AscCommon.c_oAscCellAnchorType.cellanchorOneCell) {
                if(this.from.row >= target.row) {
                    bUpdateExtents = true;
                }
                else if(this.to.row >= target.row &&
                    nType === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell) {
                    bUpdateExtents = true;
                }
            }
            else {
                this.checkBoundsFromTo();
                if(this.boundsFromTo.to.row >= target.row) {
                    bUpdateExtents = true;
                }
            }
        }
        else {
            if(nType === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell ||
                nType === AscCommon.c_oAscCellAnchorType.cellanchorOneCell) {
                if(this.from.col >= target.col) {
                    bUpdateExtents = true;
                }
                else if(this.to.col >= target.col &&
                    nType === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell) {
                    bUpdateExtents = true;
                }
            }
            else {
                this.checkBoundsFromTo();
                if(this.boundsFromTo.to.col >= target.col) {
                    bUpdateExtents = true;
                }
            }
        }
        return bUpdateExtents;
    };
    DrawingBase.prototype.draw = function (graphics) {
        if(this.graphicObject) {
            this.graphicObject.draw(graphics);
        }
    };
    DrawingBase.prototype.getBoundsFromTo = function() {
        return this.boundsFromTo;
    };
    DrawingBase.prototype.onUpdate = function (oRect) {
        if(AscCommon.isFileBuild()) {
            return;
        }
        var oDO = this.getDrawingObjects();
        if(!oDO) {
            return;
        }
        var oRange, oClipRect = null;
        if(this.isUseInDocument()) {
            if(!oRect) {
                var oB = this.getBoundsFromTo();
                var c1 = oB.from.col;
                var r1 = oB.from.row;
                var c2 = oB.to.col;
                var r2 = oB.to.row;
                oRange = new Asc.Range(c1, r1, c2, r2, true);
                oClipRect =  this.worksheet.rangeToRectAbs(oRange, 3);
            }
            else {
                oClipRect = oRect;
            }
        }
        oDO.showDrawingObjects(new AscCommon.CDrawTask(oClipRect));
    };
    DrawingBase.prototype.onSlicerUpdate = function (sName) {
        if(!this.graphicObject) {
            return false;
        }
        return this.graphicObject.onSlicerUpdate(sName);
    };
    DrawingBase.prototype.onSlicerDelete = function (sName) {
        if(!this.graphicObject) {
            return false;
        }
        return this.graphicObject.onSlicerDelete(sName);
    };
    DrawingBase.prototype.onTimeSlicerDelete = function (sName) {
        if(!this.graphicObject) {
            return false;
        }
        return this.graphicObject.onTimeSlicerDelete(sName);
    };
    DrawingBase.prototype.onSlicerLock = function (sName, bLock) {
        if(!this.graphicObject) {
            return;
        }
        this.graphicObject.onSlicerLock(sName, bLock);
    };
    DrawingBase.prototype.onSlicerChangeName = function (sName, sNewName) {
        if(!this.graphicObject) {
            return;
        }
        this.graphicObject.onSlicerChangeName(sName, sNewName);
    };
    DrawingBase.prototype.getSlicerViewByName = function (name) {
        if(!this.graphicObject) {
            return;
        }
        return this.graphicObject.getSlicerViewByName(name);
    };
    DrawingBase.prototype.handleObject = function (fCallback) {
        if(!this.graphicObject) {
            return;
        }
        this.graphicObject.handleObject(fCallback);
    };
    DrawingBase.prototype.initAfterSerialize = function(ws) {
        if(!this.graphicObject) {
            return;
        }
        let bIsShape = this.graphicObject.isShape();
        let bIsImage = this.graphicObject.isImage();
        if((bIsShape || bIsImage) && !this.graphicObject.spPr) {
            return;
        }
        if(AscCommon.IsHiddenObj(this.graphicObject)) {
            return;
        }
        this.graphicObject.setBDeleted(false);
        this.from.initAfterSerialize();
        this.to.initAfterSerialize();
        this.graphicObject.setDrawingBase(this);
        this.graphicObject.setWorksheet(ws);
        let oXfrm = this.graphicObject.spPr && this.graphicObject.spPr.xfrm;
        this.graphicObject.checkEmptySpPrAndXfrm(oXfrm);
        if(this.clientData) {
            this.graphicObject.setClientData(this.clientData);
        }
        ws.Drawings.push(this);
    };

    function DrawingObjects() {

    //-----------------------------------------------------------------------------------
    // Scroll offset
    //-----------------------------------------------------------------------------------

    var ScrollOffset = function() {

        this.getX = function() {
            return 2 * worksheet._getColLeft(0) - worksheet._getColLeft(worksheet.getFirstVisibleCol(true));
        };

        this.getY = function() {
            return 2 * worksheet._getRowTop(0) - worksheet._getRowTop(worksheet.getFirstVisibleRow(true));
        }
    };

    //-----------------------------------------------------------------------------------
    // Private
    //-----------------------------------------------------------------------------------

    var _this = this;
    var asc = window["Asc"];
    var api = asc["editor"];
    var worksheet = null;

    var drawingCtx = null;
    var overlayCtx = null;

    var scrollOffset = new ScrollOffset();

    var aObjects = [];
    var aImagesSync = [];


    var oStateBeforeLoadChanges = null;

    _this.zoom = { last: 1, current: 1 };
    _this.canEdit = null;
    _this.drawingArea = null;
    _this.drawingDocument = null;
    _this.asyncImageEndLoaded = null;
    _this.CompositeInput = null;

    _this.lastX = 0;
    _this.lastY = 0;

    _this.nCurPointItemsLength = -1;

    _this.bUpdateMetrics = true;
    _this.shiftMap = {};

    // Task timer
    _this.animId = null;
    _this.drawTask = null;

    function drawTaskFunction() {
        _this.drawingDocument.CheckTargetShow();
        if(_this.drawTask) {
            _this.showDrawingObjectsEx(_this.drawTask.getRect());
            _this.drawTask = null;
        }
        _this.animId = null;
    }

    //-----------------------------------------------------------------------------------
    // Create drawing
    //-----------------------------------------------------------------------------------


    //}

    //-----------------------------------------------------------------------------------
    // Constructor
    //-----------------------------------------------------------------------------------

    _this.addShapeOnSheet = function(sType){
        if(this.controller){
            if (_this.canEdit()) {
                _this.controller.resetSelection();
                var activeCell = worksheet.model.selectionRange.activeCell;
                var metrics = {};
                metrics.col = activeCell.col;
                metrics.colOff = 0;
                metrics.row = activeCell.row;
                metrics.rowOff = 0;
                var coordsFrom = _this.calculateCoords(metrics);
                var ext_x, ext_y;
                var oExt = AscFormat.fGetDefaultShapeExtents(sType);
                ext_x = oExt.x;
                ext_y = oExt.y;
                History.Create_NewPoint();

                var posX = pxToMm(coordsFrom.x) + MOVE_DELTA;
                var posY = pxToMm(coordsFrom.y) + MOVE_DELTA;
                var oTrack = new AscFormat.NewShapeTrack(sType, posX, posY, _this.controller.getTheme(), null, null, null, 0);
                oTrack.track({}, posX + ext_x, posY + ext_y);
                var oShape = oTrack.getShape(false, _this.drawingDocument, null);
                oShape.setWorksheet(worksheet.model);
                oShape.addToDrawingObjects();
                oShape.checkDrawingBaseCoords();
                oShape.select(_this.controller, 0);
                _this.controller.startRecalculate();
                worksheet.setSelectionShape(true);
            }
        }
    };

    _this.getScrollOffset = function()
    {
        return scrollOffset;
    };

    _this.saveStateBeforeLoadChanges = function(){
        if(this.controller){
            oStateBeforeLoadChanges = {};
            this.controller.Save_DocumentStateBeforeLoadChanges(oStateBeforeLoadChanges);
        }
        else{
            oStateBeforeLoadChanges = null;
        }
        return oStateBeforeLoadChanges;
    };

    _this.loadStateAfterLoadChanges = function(){
        if(_this.controller){
            _this.controller.clearPreTrackObjects();
            _this.controller.clearTrackObjects();
            _this.controller.resetSelection();
            _this.controller.changeCurrentState(new AscFormat.NullState(this.controller));
            if(oStateBeforeLoadChanges){
                _this.controller.loadDocumentStateAfterLoadChanges(oStateBeforeLoadChanges);
            }
        }
        oStateBeforeLoadChanges = null;
        return oStateBeforeLoadChanges;
    };

    _this.getStateBeforeLoadChanges = function(){
        return oStateBeforeLoadChanges;
    };

    _this.createDrawingObject = function(type) {
        var drawingBase = new DrawingBase(worksheet);
        if(AscFormat.isRealNumber(type))
        {
            drawingBase.Type = type;
        }
        return drawingBase;
    };

    _this.cloneDrawingObject = function(object) {

        var copyObject = _this.createDrawingObject();

        copyObject.Type = object.Type;
        copyObject.editAs = object.editAs;
        copyObject.Pos.X = object.Pos.X;
        copyObject.Pos.Y = object.Pos.Y;
        copyObject.ext.cx = object.ext.cx;
        copyObject.ext.cy = object.ext.cy;

        copyObject.from.col = object.from.col;
        copyObject.from.colOff = object.from.colOff;
        copyObject.from.row = object.from.row;
        copyObject.from.rowOff = object.from.rowOff;

        copyObject.to.col = object.to.col;
        copyObject.to.colOff = object.to.colOff;
        copyObject.to.row = object.to.row;
        copyObject.to.rowOff = object.to.rowOff;


        copyObject.boundsFromTo.from.col =  object.boundsFromTo.from.col;
        copyObject.boundsFromTo.from.colOff = object.boundsFromTo.from.colOff;
        copyObject.boundsFromTo.from.row =  object.boundsFromTo.from.row;
        copyObject.boundsFromTo.from.rowOff = object.boundsFromTo.from.rowOff;
        copyObject.boundsFromTo.to.col =  object.boundsFromTo.to.col;
        copyObject.boundsFromTo.to.colOff = object.boundsFromTo.to.colOff;
        copyObject.boundsFromTo.to.row =  object.boundsFromTo.to.row;
        copyObject.boundsFromTo.to.rowOff = object.boundsFromTo.to.rowOff;

        copyObject.graphicObject = object.graphicObject;
        return copyObject;
    };

    _this.pxToMm = function(val) {
        return pxToMm(val);
    };

    _this.mmToPx = function(val) {
        return mmToPx(val);
    };

    _this.createShapeAndInsertContent = function(oParaContent){
        var track_object = new AscFormat.NewShapeTrack("textRect", 0, 0, Asc['editor'].wbModel.theme, null, null, null, 0);
        track_object.track({}, 0, 0);
        var shape = track_object.getShape(false, _this.drawingDocument, this);
        shape.spPr.setFill(AscFormat.CreateNoFillUniFill());
        //shape.setParent(this);
        shape.txBody.content.Content[0].Add_ToContent(0, oParaContent.Copy());
        var body_pr = shape.getBodyPr();
        var w = shape.txBody.getMaxContentWidth(150, true) + body_pr.lIns + body_pr.rIns;
        var h = shape.txBody.content.GetSummaryHeight() + body_pr.tIns + body_pr.bIns;
        shape.spPr.xfrm.setExtX(w);
        shape.spPr.xfrm.setExtY(h);
        shape.spPr.xfrm.setOffX(0);
        shape.spPr.xfrm.setOffY(0);
        shape.spPr.setLn(AscFormat.CreateNoFillLine());
        shape.setWorksheet(worksheet.model);
        //shape.addToDrawingObjects();
        return shape;
    };
    //-----------------------------------------------------------------------------------
    // Public methods
    //-----------------------------------------------------------------------------------

    _this.init = function(currentSheet) {

        const api = window["Asc"]["editor"];
        worksheet = currentSheet;

        drawingCtx = currentSheet.drawingGraphicCtx;
        overlayCtx = currentSheet.overlayGraphicCtx;

        _this.drawingArea = currentSheet.drawingArea;
        _this.drawingArea.init();
        _this.drawingDocument = currentSheet.getDrawingDocument();
        _this.drawingDocument.drawingObjects = this;
        _this.drawingDocument.AutoShapesTrack = api.wb.autoShapeTrack;
        _this.drawingDocument.TargetHtmlElement = document.getElementById('id_target_cursor');
        _this.drawingDocument.InitGuiCanvasShape(api.shapeElementId);
        _this.controller = new AscFormat.DrawingObjectsController(_this);
        _this.canEdit = function() { return api.canEdit(); };

        aImagesSync = [];

        const oWS = currentSheet.model;
        aObjects = oWS.Drawings;
        let oGraphic;
        for (let i = 0; aObjects && (i < aObjects.length); i++)
        {
            aObjects[i] = _this.cloneDrawingObject(aObjects[i]);
            const drawingObject = aObjects[i];
            // Check drawing area
            drawingObject.drawingArea = _this.drawingArea;
            drawingObject.worksheet = currentSheet;
            oGraphic = drawingObject.graphicObject;
            oGraphic.setDrawingBase(drawingObject);
            oGraphic.setDrawingObjects(_this);
						oGraphic.generateLocalDrawingPart();
            oGraphic.getAllRasterImages(aImagesSync);
        }
        aImagesSync = _this.checkImageBullets(currentSheet, aImagesSync);
        const oLegacyDrawing = oWS.legacyDrawingHF;
        if(oLegacyDrawing)
        {
            const aLegacyDrawings = oLegacyDrawing.drawings;
            for(let nDrawing = 0; nDrawing < aLegacyDrawings.length; ++nDrawing)
            {
                let oLegacyDrawing = aLegacyDrawings[nDrawing];
                let oGraphic = oLegacyDrawing.graphicObject;
                oGraphic.getAllRasterImages(aImagesSync);
            }
        }



        aImagesSync = _this.checkImageBullets(currentSheet, aImagesSync);

        for(let i = 0; i < aImagesSync.length; ++i)
        {
            const localUrl = aImagesSync[i];
            if(api.DocInfo && api.DocInfo.get_OfflineApp()) {
                const urlWithMedia = AscCommon.g_oDocumentUrls.mediaPrefix + localUrl;
                if (api.imagesFromGeneralEditor && api.imagesFromGeneralEditor[urlWithMedia]) {
                    AscCommon.g_oDocumentUrls.addImageUrl(localUrl, api.imagesFromGeneralEditor[urlWithMedia]);
                } else {
                    AscCommon.g_oDocumentUrls.addImageUrl(localUrl, api.documentUrl + urlWithMedia);
                }
            }
            aImagesSync[i] = AscCommon.getFullImageSrc2(localUrl);
        }

        if(aImagesSync.length > 0)
        {
            const old_val = api.ImageLoader.bIsAsyncLoadDocumentImages;
            api.ImageLoader.bIsAsyncLoadDocumentImages = true;
            api.ImageLoader.LoadDocumentImages(aImagesSync);
            api.ImageLoader.bIsAsyncLoadDocumentImages = old_val;
        }
		_this.recalculate(true);
        worksheet.model.Drawings = aObjects;
    };

    _this.checkImageBullets = function (currentSheet, arrImages) {
        const aObjects = currentSheet.model.Drawings;
        const arrContentsWithImageBullet = [];
        const oBulletImages = {};
        const arrBulletImagesAsync = [];

        for (let i = 0; aObjects && (i < aObjects.length); i += 1) {
            const drawingObject = aObjects[i];
            drawingObject.graphicObject.getDocContentsWithImageBullets(arrContentsWithImageBullet);
            drawingObject.graphicObject.getImageFromBulletsMap(oBulletImages);
        }

        for (let localUrl in oBulletImages) {
            if(api.DocInfo && api.DocInfo.get_OfflineApp()) {
                const urlWithMedia = AscCommon.g_oDocumentUrls.mediaPrefix + localUrl;
                if (api.imagesFromGeneralEditor && api.imagesFromGeneralEditor[urlWithMedia]) {
                    AscCommon.g_oDocumentUrls.addImageUrl(localUrl, api.imagesFromGeneralEditor[urlWithMedia]);
                } else {
                    AscCommon.g_oDocumentUrls.addImageUrl(localUrl, api.documentUrl + urlWithMedia);
                }
            }
            const fullUrl = AscCommon.getFullImageSrc2(localUrl);
            arrBulletImagesAsync.push(fullUrl);
        }

        api.ImageLoader.LoadImagesWithCallback(arrBulletImagesAsync, function () {
            for (let i = 0; i < arrContentsWithImageBullet.length; i += 1) {
                const oContent = arrContentsWithImageBullet[i];
                oContent.Recalculate();
                _this.showDrawingObjects();
            }
        });

        const arrImagesWithoutImageBullets = arrImages.filter(function (sImageId) {
           return !oBulletImages[sImageId];
        });

        return arrImagesWithoutImageBullets;
    }


    _this.getSelectedDrawingsRange = function()
    {
        var i, rmin=gc_nMaxRow, rmax = 0, cmin = gc_nMaxCol, cmax = 0, selectedObjects = this.controller.selectedObjects, drawingBase;
        for(i = 0; i < selectedObjects.length; ++i)
        {

            drawingBase = selectedObjects[i].drawingBase;
            if(drawingBase)
            {
                if(drawingBase.from.col < cmin)
                {
                    cmin = drawingBase.from.col;
                }
                if(drawingBase.from.row < rmin)
                {
                    rmin = drawingBase.from.row;
                }
                if(drawingBase.to.col > cmax)
                {
                    cmax = drawingBase.to.col;
                }
                if(drawingBase.to.row > rmax)
                {
                    rmax = drawingBase.to.row;
                }
            }
        }
        return new Asc.Range(cmin, rmin, cmax, rmax, true);
    };

	_this.getScreenPosition = function(X, Y) {
		var _x = X * Asc.getCvtRatio(3, 0, worksheet._getPPIX()) + scrollOffset.getX();
		var _y = Y * Asc.getCvtRatio(3, 0, worksheet._getPPIY()) + scrollOffset.getY();
		return new AscCommon.asc_CRect(_x, _y, 0, 0 );
	};
	_this.convertCoordsToCursorWR = function(X, Y) {
		return this.drawingArea.convertCoordsToCursorWR(X, Y);
	};

    _this.getContextMenuPosition = function(){
        if(!worksheet){
            return new AscCommon.asc_CRect( 0, 0, 5, 5 );
        }
        let oPos = this.controller.getContextMenuPosition(0);
		return _this.getScreenPosition(oPos.X, oPos.Y);
    };

    _this.recalculate =  function(all)
    {
        _this.controller.recalculate2(all);
    };

    _this.preCopy = function() {
        _this.shiftMap = {};
        var selected_objects = _this.controller.selectedObjects;
        if(selected_objects.length > 0)
        {
            var min_x, min_y, i;
            min_x = selected_objects[0].x;
            min_y = selected_objects[0].y;
            for(i = 1; i < selected_objects.length; ++i)
            {
                if(selected_objects[i].x < min_x)
                    min_x = selected_objects[i].x;

                if(selected_objects[i].y < min_y)
                    min_y = selected_objects[i].y;
            }
            for(i = 0; i < selected_objects.length; ++i)
            {
                _this.shiftMap[selected_objects[i].Get_Id()] = {x: selected_objects[i].x - min_x, y: selected_objects[i].y - min_y};
            }
        }

    };

    _this.getAllFonts = function(AllFonts) {

    };

    _this.getOverlay = function() {
        return api.wb.trackOverlay;
    };

    _this.OnUpdateOverlay = function() {
        _this.drawingArea.drawSelection(_this.drawingDocument);
    };

    _this.changeZoom = function(factor) {

        _this.zoom.last = _this.zoom.current;
        _this.zoom.current = factor;

        _this.resizeCanvas();
    };

    _this.resizeCanvas = function() {
		_this.drawingArea.init();
    };

    _this.getCanvasContext = function() {
        return _this.drawingDocument.CanvasHitContext;
    };

    _this.getDrawingObjects = function() {
        return aObjects;
    };

    _this.getWorksheet = function() {
        return worksheet;
    };

	_this.getContextWidth = function () {
		return drawingCtx.getWidth();
	};
	_this.getContextHeight = function () {
		return drawingCtx.getHeight();
	};
    _this.getContext = function () {
        return drawingCtx;
    };

    _this.getWorksheetModel = function() {
        return worksheet.model;
    };

    _this.callTrigger = function(triggerName, param) {
        if ( triggerName )
            worksheet.model.workbook.handlers.trigger(triggerName, param);
    };

    _this.getDrawingObjectsBounds = function()
    {
        var arr_x = [], arr_y = [], bounds;
        for(var i = 0; i < aObjects.length; ++i)
        {
            bounds = aObjects[i].graphicObject.bounds;
            arr_x.push(bounds.l);
            arr_x.push(bounds.r);
            arr_y.push(bounds.t);
            arr_y.push(bounds.b);
        }
        return new DrawingBounds(Math.min.apply(Math, arr_x), Math.max.apply(Math, arr_x), Math.min.apply(Math, arr_y), Math.max.apply(Math, arr_y));
    };

    //-----------------------------------------------------------------------------------
    // Drawing objects
    //-----------------------------------------------------------------------------------

    _this.showDrawingObjects = function(graphicOption) {
        if(!worksheet || !worksheet.model ||
            !api || !api.wb || !api.wb.model) {
            return
        }
        if ( (worksheet.model.index !== api.wb.model.getActive()))
            return;
        var oNewTask;
        if(graphicOption) {
            oNewTask = graphicOption;
        }
        else {
            oNewTask = new AscCommon.CDrawTask(null);
                }
        if(window["IS_NATIVE_EDITOR"]) {
            _this.showDrawingObjectsEx(oNewTask.getRect());
            return;
        }
        if(_this.drawTask === null) {
            _this.drawTask = oNewTask;
                }
            else {
            _this.drawTask = _this.drawTask.union(oNewTask);
            }
        if(_this.animId === null) {
            _this.animId = rAF(drawTaskFunction);
        }
    };

    _this.showDrawingObjectsEx = function(oUpdateRect) {
        if(!drawingCtx) {
            return;
        }
        if (worksheet.model !== api.wb.model.getActiveWs()) {
            return;
        }
        if (!oUpdateRect) {
            if(!window['IS_NATIVE_EDITOR']) {
                _this.drawingArea.clear();
            }
                }
        for (var nDrawing = 0; nDrawing < aObjects.length; nDrawing++) {
            _this.drawingArea.drawObject(aObjects[nDrawing], oUpdateRect);
                }
        _this.OnUpdateOverlay();
        _this.controller.updateSelectionState(true);
        AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
    };

    _this.updateRange = function(oRange) {
        if(!drawingCtx) {
            return;
        }
        if (worksheet.model.index !== api.wb.model.getActive()) {
            return;
        }
        for (var nDrawing = 0; nDrawing < aObjects.length; nDrawing++) {
            _this.drawingArea.updateRange(aObjects[nDrawing], oRange);
        }
        _this.OnUpdateOverlay();
        _this.controller.updateSelectionState(true);
        AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
		Asc.editor.wbModel.mathTrackHandler.Update();
    };

    _this.print = function(oOptions) {
        for(var nObject = 0; nObject < aObjects.length; ++nObject) {
            var oDrawing = aObjects[nObject];
            oDrawing.draw(oOptions.ctx.DocumentRenderer);
        }
    };

    _this.getDrawingDocument = function()
    {
        return _this.drawingDocument;
    };

    _this.printGraphicObject = function(graphicObject, ctx) {

        if ( graphicObject && ctx ) {
            // Image
            if ( graphicObject instanceof AscFormat.CImageShape )
                printImage(graphicObject, ctx);
            // Shape
            else if ( graphicObject instanceof AscFormat.CShape )
                printShape(graphicObject, ctx);
            // Chart
            else if (graphicObject instanceof AscFormat.CChartSpace)
                printChart(graphicObject, ctx);
            // Group
            else if ( graphicObject instanceof AscFormat.CGroupShape )
                printGroup(graphicObject, ctx);
        }

        // Print functions
        function printImage(graphicObject, ctx) {

            if ( (graphicObject instanceof AscFormat.CImageShape) && graphicObject && ctx ) {
                // Save
                graphicObject.draw( ctx );
            }
        }

        function printShape(graphicObject, ctx) {

            if ( (graphicObject instanceof AscFormat.CShape) && graphicObject && ctx ) {
                graphicObject.draw( ctx );
            }
        }

        function printChart(graphicObject, ctx) {

            if ( (graphicObject instanceof AscFormat.CChartSpace) && graphicObject && ctx ) {

                graphicObject.draw( ctx );
            }
        }

        function printGroup(graphicObject, ctx) {

            if ( (graphicObject instanceof AscFormat.CGroupShape) && graphicObject && ctx ) {
                for ( var i = 0; i < graphicObject.arrGraphicObjects.length; i++ ) {
                    var graphicItem = graphicObject.arrGraphicObjects[i];

                    if ( graphicItem instanceof AscFormat.CImageShape )
                        printImage(graphicItem, ctx);

                    else if ( graphicItem instanceof AscFormat.CShape )
                        printShape(graphicItem, ctx);

                    else if (graphicItem instanceof AscFormat.CChartSpace )
                        printChart(graphicItem, ctx);
                }
            }
        }
    };

    _this.getMaxColRow = function() {
        var r = -1, c = -1;
        aObjects.forEach(function (item) {
            r = Math.max(r, item.boundsFromTo.to.row);
            c = Math.max(c, item.boundsFromTo.to.col);
        });

        return new AscCommon.CellBase(r, c);
    };

    //-----------------------------------------------------------------------------------
    // For object type
    //-----------------------------------------------------------------------------------


    _this.calculateObjectMetrics = function (object, width, height, opt_checkRtl) {
        // Обработка картинок большого разрешения
        var bCorrect = false;
        var metricCoeff = 1;

        var coordsFrom = _this.calculateCoords(object.from);
        var realTopOffset = coordsFrom.y;
        var realLeftOffset = coordsFrom.x;

        var areaWidth = worksheet._getColLeft(worksheet.getLastVisibleCol()) - worksheet._getColLeft(worksheet.getFirstVisibleCol(true)); 	// по ширине
        if (areaWidth < width) {
            metricCoeff = width / areaWidth;

            width = areaWidth;
            height /= metricCoeff;
            bCorrect = true;
        }

        var areaHeight = worksheet._getRowTop(worksheet.getLastVisibleRow()) - worksheet._getRowTop(worksheet.getFirstVisibleRow(true)); 	// по высоте
        if (areaHeight < height) {
            metricCoeff = height / areaHeight;

            height = areaHeight;
            width /= metricCoeff;
            bCorrect = true;
        }

        if (opt_checkRtl && worksheet.getRightToLeft()) {
            width *= -1;
        }
        var toCell = worksheet.findCellByXY(realLeftOffset + width, realTopOffset + height, true, false, false);
        object.to.col = toCell.col;
        object.to.colOff = pxToMm(toCell.colOff);
        object.to.row = toCell.row;
        object.to.rowOff = pxToMm(toCell.rowOff);
        return bCorrect;
    };



    _this.addImageObjectCallback = function (_image, options) {
        var isOption = options && options.cell;
        if (!_image.Image) {
            worksheet.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.UplImageUrl, c_oAscError.Level.NoCritical);
        } else {
            var drawingObject = _this.createDrawingObject();
            drawingObject.worksheet = worksheet;

            var activeCell = worksheet.model.selectionRange.activeCell;
            drawingObject.from.col = isOption ? options.cell.col : activeCell.col;
            drawingObject.from.row = isOption ? options.cell.row : activeCell.row;

            var oSize;
            if(!isOption) {
                var oImgP = new Asc.asc_CImgProperty();
                oImgP.ImageUrl = _image.src;
                oSize = oImgP.asc_getOriginSize(api);
            }
            else {
                oSize = new AscCommon.asc_CImageSize(Math.max((options.width * AscCommon.g_dKoef_pix_to_mm), 1),
                    Math.max((options.height * AscCommon.g_dKoef_pix_to_mm), 1), true);
            }
            var bCorrect = _this.calculateObjectMetrics(drawingObject, mmToPx(oSize.asc_getImageWidth()), mmToPx(oSize.asc_getImageHeight()), true);

            var coordsFrom = _this.calculateCoords(drawingObject.from);
            var coordsTo = _this.calculateCoords(drawingObject.to);
            if (worksheet.getRightToLeft()) {
                let temp = coordsTo.x;
                coordsTo.x = coordsFrom.x;
                coordsFrom.x = temp;
            }
            if(bCorrect) {
                _this.controller.addImageFromParams(_image.src, pxToMm(coordsFrom.x) + MOVE_DELTA, pxToMm(coordsFrom.y) + MOVE_DELTA, pxToMm(coordsTo.x - coordsFrom.x), pxToMm(coordsTo.y - coordsFrom.y));
            }
            else {
                _this.controller.addImageFromParams(_image.src, pxToMm(coordsFrom.x) + MOVE_DELTA, pxToMm(coordsFrom.y) + MOVE_DELTA, oSize.asc_getImageWidth(), oSize.asc_getImageHeight());
            }
        }
    };



    _this.addImageDrawingObject = function(imageUrls, options) {
        if (imageUrls && _this.canEdit()) {
            api.ImageLoader.LoadImagesWithCallback(imageUrls, function(){
                // CImage
                _this.controller.resetSelection();
                History.Create_NewPoint();
                for(var i = 0; i < imageUrls.length; ++i){
                    var sImageUrl = imageUrls[i];
                    var _image = api.ImageLoader.LoadImage(sImageUrl, 1);
                    if (null != _image) {
                        _this.addImageObjectCallback(_image, options);
                    } else {
                        worksheet.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.UplImageUrl, c_oAscError.Level.NoCritical);
                        break;
                    }
                }
                _this.controller.startRecalculate();
                worksheet.setSelectionShape(true);
            }, []);
        }
    };

    _this.addTextArt = function(nStyle)
    {
        if (_this.canEdit()) {

            var oVisibleRange = worksheet.getVisibleRange();
            _this.controller.resetSelection();
            var dLeft = worksheet.getCellLeft(oVisibleRange.c1, 3);
            var dTop = worksheet.getCellTop(oVisibleRange.r1, 3);
            var dRight = worksheet.getCellLeft(oVisibleRange.c2, 3) + worksheet.getColumnWidth(oVisibleRange.c2, 3);
            var dBottom = worksheet.getCellTop(oVisibleRange.r2, 3) + worksheet.getRowHeight(oVisibleRange.r2, 3);
            _this.controller.addTextArtFromParams(nStyle, dLeft, dTop, dRight - dLeft, dBottom - dTop, worksheet.model);
            worksheet.setSelectionShape(true);
        }

    };

    _this.addChartSpace = function (oChartSpace) {
        if (!_this.canEdit())
            return;

        let aSelectedDrawings = _this.controller.getSelectedArray();
        if(aSelectedDrawings.length === 1 && aSelectedDrawings[0].isChart()) {
            _this.controller.checkSelectedObjectsAndCallback(function () {
                let oSelectedChartSpace = aSelectedDrawings[0];
                oSelectedChartSpace.fromOther(oChartSpace);
                _this.controller.startRecalculate();
                _this.sendGraphicObjectProps();
            }, [], false, 0, [], false);
            return;
        }

        History.Create_NewPoint();
        worksheet.setSelectionShape(true);
        let oCSForAdd = oChartSpace.copy();
        let w, h, chartLeft, chartTop;
        w = this.convertMetric(AscCommon.AscBrowser.convertToRetinaValue(AscCommon.c_oAscChartDefines.defaultChartWidth, true), 0, 3);
        h = this.convertMetric(AscCommon.AscBrowser.convertToRetinaValue(AscCommon.c_oAscChartDefines.defaultChartHeight, true), 0, 3);
        chartLeft =  -this.convertMetric(this.getScrollOffset().getX(), 0, 3) + (this.convertMetric(this.getContextWidth(), 0, 3) - w) / 2;
        if(chartLeft < 0)
        {
            chartLeft = 0;
        }
        chartTop =  -this.convertMetric(this.getScrollOffset().getY(), 0, 3) + (this.convertMetric(this.getContextHeight(), 0, 3) - h) / 2;
        if(chartTop < 0)
        {
            chartTop = 0;
        }
        oCSForAdd.setTransformParams(chartLeft, chartTop, w, h, 0, false, false);
        oCSForAdd.setWorksheet(this.getWorksheetModel());
        oCSForAdd.setBDeleted(false);
        oCSForAdd.setDrawingObjects(this);
        oCSForAdd.addToDrawingObjects();
        _this.controller.resetSelection();
        _this.controller.selectObject(oCSForAdd, 0);
        oCSForAdd.addToRecalculate();
        oCSForAdd.checkDrawingBaseCoords();
        _this.controller.startRecalculate();
        this.sendGraphicObjectProps();

    };

    _this.addSlicers = function(aNames) {
        if (_this.canEdit()) {
            if(Array.isArray(aNames) && aNames.length > 0) {
                var oVisibleRange = worksheet.getVisibleRange();
                var nSlicerCount = aNames.length;
                var dSlicerWidth = 2 * 25.4;//
                var dSlicerHeight = 2.76 * 25.4;
                var dSlicerInset = 10;
                var dTotalWidth = dSlicerWidth + nSlicerCount * dSlicerInset;
                var dTotalHeight = dSlicerHeight + nSlicerCount * dSlicerInset;
                var dLeft = worksheet.getCellLeft(oVisibleRange.c1, 3);
                var dTop = worksheet.getCellTop(oVisibleRange.r1, 3);
                var dRight = worksheet.getCellLeft(oVisibleRange.c2, 3) + worksheet.getColumnWidth(oVisibleRange.c2, 3);
                var dBottom = worksheet.getCellTop(oVisibleRange.r2, 3) + worksheet.getRowHeight(oVisibleRange.r2, 3);
                _this.controller.resetSelection();
                var dStartPosX = Math.max(0, (dLeft + dRight) / 2.0 - dTotalWidth / 2);
                var dStartPosY = Math.max(0, (dTop + dBottom) / 2.0 - dTotalHeight / 2);
                var oSlicer, x, y;
                for(var nSlicerIndex = 0; nSlicerIndex < nSlicerCount; ++nSlicerIndex) {
                    oSlicer = new AscFormat.CSlicer();
                    oSlicer.setName(aNames[nSlicerIndex]);
                    x = dStartPosX + dSlicerInset * nSlicerIndex;
                    y = dStartPosY + dSlicerInset * nSlicerIndex;
                    oSlicer.setBDeleted(false);
                    oSlicer.setWorksheet(worksheet.model);
                    oSlicer.setTransformParams(x, y, dSlicerWidth, dSlicerHeight, 0, false, false);
                    oSlicer.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorAbsolute);
                    oSlicer.checkDrawingBaseCoords();
                }
                _this.controller.startRecalculate();
                oSlicer.select(_this.controller, 0);
                worksheet.setSelectionShape(true);

            }
        }
    };
    _this.addSignatureLine = function(oPr, Width, Height, sImgUrl)
    {
        var oApi = window["Asc"]["editor"];
        var sGuid = oPr.asc_getGuid();
        var oSpToEdit = null;
        if(sGuid)
        {
            var oDrawingObjects = this.controller;
            var ret = [], allSpr = [];
            allSpr = allSpr.concat(allSpr.concat(oDrawingObjects.getAllSignatures2(ret, oDrawingObjects.getDrawingArray())));
            for(var i = 0; i < allSpr.length; ++i)
            {
                if(allSpr[i].getSignatureLineGuid() === sGuid)
                {
                    oSpToEdit = allSpr[i];
                    break;
                }
            }
        }
        History.Create_NewPoint();
        if(!oSpToEdit)
        {
            _this.controller.resetSelection();
            var dLeft = worksheet.getCellLeft(worksheet.model.selectionRange.activeCell.col, 3);
            var dTop = worksheet.getCellTop(worksheet.model.selectionRange.activeCell.row, 3);
            var oSignatureLine = AscFormat.fCreateSignatureShape(oPr, false, worksheet.model, Width, Height, sImgUrl);
            oSignatureLine.spPr.xfrm.setOffX(dLeft);
            oSignatureLine.spPr.xfrm.setOffY(dTop);
            oSignatureLine.addToDrawingObjects();
            oSignatureLine.checkDrawingBaseCoords();
            _this.controller.selectObject(oSignatureLine, 0);
            worksheet.setSelectionShape(true);
            oApi.sendEvent("asc_onAddSignature", oSignatureLine.signatureLine.id);
        }
        else
        {
            oSpToEdit.setSignaturePr(oPr, sImgUrl);
            oApi.sendEvent("asc_onAddSignature", sGuid);
        }

        _this.controller.startRecalculate();
    };

    _this.addMath = function(Type){
        if (_this.canEdit() && _this.controller) {

            var oTargetContent = _this.controller.getTargetDocContent();
            if(oTargetContent){

                _this.controller.checkSelectedObjectsAndCallback(function(){
                    var MathElement = new AscCommonWord.MathMenu(Type);
                    _this.controller.paragraphAdd(MathElement, false);
	                _this.controller.recalculate();
					_this.controller.updateSelectionState();
                }, [], false, AscDFH.historydescription_Spreadsheet_CreateGroup);
                return;
            }


            _this.controller.resetSelection();

            var activeCell = worksheet.model.selectionRange.activeCell;
            var coordsFrom = _this.calculateCoords({col: activeCell.col, row: activeCell.row, colOff: 0, rowOff: 0});

            History.Create_NewPoint();
            var oSp = _this.controller.createTextArt(0, false, worksheet.model, "");
            _this.controller.resetSelection();
            oSp.setWorksheet(_this.controller.drawingObjects.getWorksheetModel());
            oSp.setDrawingObjects(_this.controller.drawingObjects);
            oSp.addToDrawingObjects();

            var oContent = oSp.getDocContent();
            if(oContent){
                oContent.MoveCursorToStartPos(false);
                oContent.AddToParagraph(new AscCommonWord.MathMenu(Type), false);
            }
            oSp.checkExtentsByDocContent();
            oSp.spPr.xfrm.setOffX(pxToMm(coordsFrom.x) + MOVE_DELTA - ((worksheet.getRightToLeft() && oSp.spPr.xfrm.extX) ? oSp.spPr.xfrm.extX : 0));
            oSp.spPr.xfrm.setOffY(pxToMm(coordsFrom.y) + MOVE_DELTA);

            oSp.checkDrawingBaseCoords();
            _this.controller.selectObject(oSp, 0);
            _this.controller.selection.textSelection = oSp;
            oSp.addToRecalculate();
            _this.controller.startRecalculate();
            worksheet.setSelectionShape(true);
        }
    };
    _this.setMathProps = function(MathProps)
    {
        _this.controller.setMathProps(MathProps);
    }
    _this.convertMathView = function(isToLinear, isAll)
    {
        _this.controller.convertMathView(isToLinear, isAll);
        _this.controller.updateSelectionState();
    }
    _this.setListType = function(type, subtype, custom)
    {
        if(_this.controller.checkSelectedObjectsProtectionText())
        {
            return;
        }
        var NumberInfo =
            {
                Type    : 0,
                SubType : -1
            };

        NumberInfo.Type    = type;
        NumberInfo.SubType = subtype;
        NumberInfo.Custom = custom;
        _this.controller.checkSelectedObjectsAndCallback(_this.controller.setParagraphNumbering, [AscFormat.fGetPresentationBulletByNumInfo(NumberInfo)], false, AscDFH.historydescription_Presentation_SetParagraphNumbering);
    };

    _this.editImageDrawingObject = function(imageUrl, obj) {

        if ( imageUrl ) {
            var _image = api.ImageLoader.LoadImage(imageUrl, 1);

            var addImageObject = function (_image) {

                if ( !_image.Image ) {
                    worksheet.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.UplImageUrl, c_oAscError.Level.NoCritical);
                }
                else {
                    if ( obj && obj.isImageChangeUrl ) {
                        var imageProp = new Asc.asc_CImgProperty();
                        imageProp.ImageUrl = _image.src;
                        _this.setGraphicObjectProps(imageProp);
                    }
                    else if ( obj && obj.isShapeImageChangeUrl ) {
                        var imgProps = new Asc.asc_CImgProperty();
                        var shapeProp = new Asc.asc_CShapeProperty();
                        imgProps.ShapeProperties = shapeProp;
                        shapeProp.fill = new Asc.asc_CShapeFill();
                        shapeProp.fill.type = Asc.c_oAscFill.FILL_TYPE_BLIP;
                        shapeProp.fill.fill = new Asc.asc_CFillBlip();
                        shapeProp.fill.fill.asc_putUrl(_image.src);
                        if(obj.textureType !== null && obj.textureType !== undefined){
                            shapeProp.fill.fill.asc_putType(obj.textureType);
                        }
                        _this.setGraphicObjectProps(imgProps);
                    }
                    else if(obj && obj.isTextArtChangeUrl)
                    {
                        var imgProps = new Asc.asc_CImgProperty();
                        var AscShapeProp = new Asc.asc_CShapeProperty();
                        imgProps.ShapeProperties = AscShapeProp;
                        var oFill = new Asc.asc_CShapeFill();
                        oFill.type = Asc.c_oAscFill.FILL_TYPE_BLIP;
                        oFill.fill = new Asc.asc_CFillBlip();
                        oFill.fill.asc_putUrl(imageUrl);
                        if(obj.textureType !== null && obj.textureType !== undefined){
                            oFill.fill.asc_putType(obj.textureType);
                        }
                        AscShapeProp.textArtProperties = new Asc.asc_TextArtProperties();
                        AscShapeProp.textArtProperties.asc_putFill(oFill);

                        _this.setGraphicObjectProps(imgProps);
                    }
                    else if (obj && obj.fAfterUploadOleObjectImage) {
                        obj.fAfterUploadOleObjectImage(_image.src);
                    }

                    _this.showDrawingObjects();
                }
            };

            if (null != _image) {
                addImageObject(_image);
            }
            else {
                _this.asyncImageEndLoaded = function(_image) {
                    addImageObject(_image);
                    _this.asyncImageEndLoaded = null;
                }
            }
        }
    };

    _this.addChartDrawingObject = function(chart) {

        if (!_this.canEdit())
            return;

        worksheet.setSelectionShape(true);

        if ( chart instanceof Asc.asc_ChartSettings )
        {
            if(api.isChartEditor)
            {
				_this.controller.selectObject(aObjects[0].graphicObject, 0);
				_this.controller.editChartDrawingObjects(chart);
                return;
            }
            _this.controller.addChartDrawingObject(chart);
        }
        else if ( isObject(chart) && chart["binary"] )
        {
            var model = worksheet.model;
			History.Clear();
            for (var i = 0; i < aObjects.length; i++) {
                aObjects[i].graphicObject.deleteDrawingBase();
            }
			aObjects.length = 0;

			var oAllRange = model.getRange3(0, 0, model.getRowsCount(), model.getColsCount());
			oAllRange.cleanAll();

			worksheet.endEditChart();

			//worksheet._clean();

            var asc_chart_binary = new Asc.asc_CChartBinary();
            asc_chart_binary.asc_setBinary(chart["binary"]);
            asc_chart_binary.asc_setThemeBinary(chart["themeBinary"]);
            asc_chart_binary.asc_setColorMapBinary(chart["colorMapBinary"]);
            asc_chart_binary.asc_setIsChartEx(chart["IsChartEx"]);
            var oNewChartSpace = asc_chart_binary.getChartSpace(model);
            var theme = asc_chart_binary.getTheme();
            if(theme)
            {
                model.workbook.theme = theme;
            }
            if(chart["cTitle"] || chart["cDescription"]){
                if(!oNewChartSpace.nvGraphicFramePr){
                    var nv_sp_pr = new AscFormat.UniNvPr();
                    nv_sp_pr.cNvPr.setId(1024);
                    oNewChartSpace.setNvSpPr(nv_sp_pr);
                }
                oNewChartSpace.setTitle(chart["cTitle"]);
                oNewChartSpace.setDescription(chart["cDescription"]);
            }
            var colorMapOverride = asc_chart_binary.getColorMap();
            if(colorMapOverride)
            {
                AscFormat.DEFAULT_COLOR_MAP = colorMapOverride;
            }

            if(typeof chart["urls"] === "string") {
                AscCommon.g_oDocumentUrls.addUrls(JSON.parse(chart["urls"]));
            }
            var font_map = {};
            oNewChartSpace.documentGetAllFontNames(font_map);
            AscFormat.checkThemeFonts(font_map, model.workbook.theme.themeElements.fontScheme);
            window["Asc"]["editor"]._loadFonts(font_map,
                function() {
                    AscCommonExcel.executeInR1C1Mode(false,
                        function()
                        {
                            var max_r = 0, max_c = 0;

                            var series = oNewChartSpace.getAllSeries(), ser;

                            function fFillCell(oCell, sNumFormat, value)
                            {
                                var oCellValue = new AscCommonExcel.CCellValue();
                                if(AscFormat.isRealNumber(value))
                                {
                                    oCellValue.number = value;
                                    oCellValue.type = AscCommon.CellValueType.Number;
                                }
                                else
                                {
                                    oCellValue.text = value;
                                    oCellValue.type = AscCommon.CellValueType.String;
                                }
                                oCell.setNumFormat(sNumFormat);
                                oCell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
                            }

                            function fillTableFromRef(ref)
                            {
                                var cache;
                                if(ref.numCache) {
                                    cache =ref.numCache;
                                }
                                else if(ref.strCache) {
                                    cache = ref.strCache;
                                }
                                else {
                                    if(Array.isArray(ref.levelData)) {
                                        cache = ref.levelData[0];
                                    }
                                    else if(typeof ref.v === "string") {
                                        cache = AscFormat.ExecuteNoHistory(
                                            function() {
                                                let oStrCache = new AscFormat.CStrCache();
                                                let oPt = new AscFormat.CStringPoint();
                                                oPt.idx = 0;
                                                oPt.val = ref.v;
                                                oStrCache.addPt(oPt);
                                                return oStrCache;
                                            }, this, []
                                        );
                                    }
                                }
                                var sFormula = "";
                                if(typeof ref.f === "string") {
                                    sFormula = ref.f;
                                }
                                else if(ref.f && ref.f.content) {
                                    sFormula = ref.f.content;
                                }
                                var lit_format_code;
                                if(cache)
                                {
                                    lit_format_code = (typeof cache.formatCode === "string" && cache.formatCode.length > 0) ? cache.formatCode : "General";

                                    if(sFormula[0] === '(')
                                        sFormula = sFormula.slice(1);
                                    if(sFormula[sFormula.length-1] === ')')
                                        sFormula = sFormula.slice(0, -1);
                                    var f1 = sFormula;

                                    var arr_f = f1.split(",");
                                    var pt_index = 0, i, j, pt, nPtCount, k;
                                    for(i = 0; i < arr_f.length; ++i)
                                    {
                                        var parsed_ref = parserHelp.parse3DRef(arr_f[i]);
                                        if(parsed_ref)
                                        {
                                            var source_worksheet = model.workbook.getWorksheetByName(parsed_ref.sheet);
                                            if(source_worksheet === model)
                                            {
                                                var range = source_worksheet.getRange2(parsed_ref.range);
                                                if(range)
                                                {
                                                    range = range.bbox;

                                                    if(range.r1 > max_r)
                                                        max_r = range.r1;
                                                    if(range.r2 > max_r)
                                                        max_r = range.r2;

                                                    if(range.c1 > max_c)
                                                        max_c = range.c1;
                                                    if(range.c2 > max_c)
                                                        max_c = range.c2;

                                                    if(i === arr_f.length - 1)
                                                    {
                                                        nPtCount = cache.getPtCount();
                                                        if((nPtCount - pt_index) <=(range.r2 - range.r1 + 1))
                                                        {
                                                            for(k = range.c1; k <= range.c2; ++k)
                                                            {
                                                                for(j = range.r1; j <= range.r2; ++j)
                                                                {
                                                                    source_worksheet._getCell(j, k, function(cell) {
                                                                        pt = cache.getPtByIndex(pt_index + j - range.r1);
                                                                        if(pt)
                                                                        {
                                                                            fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
                                                                        }
                                                                    });
                                                                }
                                                            }
                                                            pt_index += (range.r2 - range.r1 + 1);
                                                        }
                                                        else if((nPtCount - pt_index) <= (range.c2 - range.c1 + 1))
                                                        {
                                                            for(k = range.r1; k <= range.r2; ++k)
                                                            {
                                                                for(j = range.c1;  j <= range.c2; ++j)
                                                                {
                                                                    source_worksheet._getCell(k, j, function(cell) {
                                                                        pt = cache.getPtByIndex(pt_index + j - range.c1);
                                                                        if(pt)
                                                                        {
                                                                            fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
                                                                        }
                                                                    });
                                                                }
                                                            }
                                                            pt_index += (range.c2 - range.c1 + 1);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if(range.r1 === range.r2)
                                                        {
                                                            for(j = range.c1;  j <= range.c2; ++j)
                                                            {
                                                                source_worksheet._getCell(range.r1, j, function(cell) {
                                                                    pt = cache.getPtByIndex(pt_index);
                                                                    if(pt)
                                                                    {
                                                                        fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
                                                                    }
                                                                    ++pt_index;
                                                                });
                                                            }
                                                        }
                                                        else
                                                        {
                                                            for(j = range.r1; j <= range.r2; ++j)
                                                            {
                                                                source_worksheet._getCell(j, range.c1, function(cell) {
                                                                    pt = cache.getPtByIndex(pt_index);
                                                                    if(pt)
                                                                    {
                                                                        fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
                                                                    }
                                                                    ++pt_index;
                                                                });
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            var first_num_ref;
                            if(series[0])
                            {
                                if(series[0].val)
                                    first_num_ref = series[0].val.numRef;
                                else if(series[0].yVal)
                                    first_num_ref = series[0].yVal.numRef;
                            }
                            if(first_num_ref)
                            {
                                var resultRef = parserHelp.parse3DRef(first_num_ref.f);
                                if(resultRef)
                                {
                                    model.workbook.aWorksheets[0].sName = resultRef.sheet;
                                    var oCat, oVal;
                                    for(var i = 0; i < series.length; ++i)
                                    {
                                        ser = series[i];
                                        oVal = ser.val || ser.yVal;
                                        if(oVal && oVal.numRef)
                                        {
                                            fillTableFromRef(oVal.numRef);
                                        }
                                        oCat = ser.cat || ser.xVal;
                                        if(oCat)
                                        {
                                            if(oCat.numRef)
                                            {
                                                fillTableFromRef(oCat.numRef);
                                            }
                                            if(oCat.strRef)
                                            {
                                                fillTableFromRef(oCat.strRef);
                                            }
                                        }
                                        if(ser.tx && ser.tx.strRef)
                                        {
                                            fillTableFromRef(ser.tx.strRef);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if(oNewChartSpace.isChartEx()) {
                                    let aSeries = oNewChartSpace.getAllSeries();

                                    for(let nS = 0; nS < aSeries.length; ++nS) {
                                        let oSeries = aSeries[nS];
                                        let oData = oSeries.getData();
                                        if(oData) {
                                            let aDims = oData.dimension;
                                            for(let nDim = 0; nDim < aDims.length; ++nDim) {
                                                let oDim = aDims[nDim];
                                                fillTableFromRef(oDim);
                                            }
                                        }
                                        if(oSeries.tx && oSeries.tx.txData) {
                                            fillTableFromRef(oSeries.tx.txData);
                                        }
                                    }
                                }
                            }
                            oAllRange = oAllRange.bbox;
                            oAllRange.r2 = Math.max(oAllRange.r2, max_r);
                            oAllRange.c2 = Math.max(oAllRange.c2, max_c);
                            worksheet._updateRange(oAllRange);
                            worksheet.draw();
                            aImagesSync.length = 0;
                            oNewChartSpace.getAllRasterImages(aImagesSync);
                            oNewChartSpace.setBDeleted(false);
                            oNewChartSpace.setWorksheet(model);
                            oNewChartSpace.addToDrawingObjects();
                            oNewChartSpace.recalcInfo.recalculateReferences = false;
                            var oDrawingBase_ = oNewChartSpace.drawingBase;
                            oNewChartSpace.drawingBase = null;
                            oNewChartSpace.recalculate();
                            AscFormat.CheckSpPrXfrm(oNewChartSpace);
                            oNewChartSpace.drawingBase = oDrawingBase_;

                            var canvas_height = worksheet.drawingCtx.getHeight(3);
                            var pos_y = (canvas_height - oNewChartSpace.spPr.xfrm.extY)/2;
                            if(pos_y < 0)
                            {
                                pos_y = 0;
                            }

                            var canvas_width = worksheet.drawingCtx.getWidth(3);
                            var pos_x = (canvas_width - oNewChartSpace.spPr.xfrm.extX)/2;
                            if(pos_x < 0)
                            {
                                pos_x = 0;
                            }
                            oNewChartSpace.spPr.xfrm.setOffX(pos_x);
                            oNewChartSpace.spPr.xfrm.setOffY(pos_y);
                            oNewChartSpace.checkDrawingBaseCoords();
                            oNewChartSpace.recalculate();
                            worksheet._scrollToRange(_this.getSelectedDrawingsRange());
                            _this.showDrawingObjects();
                            _this.controller.resetSelection();
                            _this.controller.selectObject(oNewChartSpace, 0);
                            _this.controller.updateSelectionState();
                            _this.sendGraphicObjectProps();
                            if(aImagesSync.length > 0)
                            {
                                window["Asc"]["editor"].ImageLoader.LoadDocumentImages(aImagesSync);
                            }
                            History.Clear();
                        });
                }
            );


        }
    };

    _this.editChartDrawingObject = function(chart)
    {
        if ( chart )
        {
            if(api.isChartEditor)
            {
				_this.controller.selectObject(aObjects[0].graphicObject, 0);
            }
            _this.controller.editChartDrawingObjects(chart);
            //_this.showDrawingObjects();
        }
    };

    _this.getRecommendedChartData = function () {

    };

    _this.checkSparklineGroupMinMaxVal = function(oSparklineGroup)
    {
        var maxVal = null, minVal = null, i, j, sparkline, nPtCount = 0;
        if(oSparklineGroup.type !== Asc.c_oAscSparklineType.Stacked &&
            (Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.minAxisType || Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.maxAxisType))
        {
            for(i = 0; i < oSparklineGroup.arrSparklines.length; ++i)
            {
				sparkline = oSparklineGroup.arrSparklines[i];
                if (!sparkline.oCacheView) {
					sparkline.oCacheView = new CSparklineView();
					sparkline.oCacheView.initFromSparkline(sparkline, oSparklineGroup, worksheet);
                }
                var aPoints = sparkline.oCacheView.chartSpace.chart.plotArea.charts[0].series[0].getNumPts();
                for(j = 0; j < aPoints.length; ++j)
                {
                    ++nPtCount;
                    if(Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.maxAxisType)
                    {
                        if(maxVal === null)
                        {
                            maxVal = aPoints[j].val;
                        }
                        else
                        {
                            if(maxVal < aPoints[j].val)
                            {
                                maxVal = aPoints[j].val;
                            }
                        }
                    }
                    if(Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.minAxisType)
                    {
                        if(minVal === null)
                        {
                            minVal = aPoints[j].val;
                        }
                        else
                        {
                            if(minVal > aPoints[j].val)
                            {
                                minVal = aPoints[j].val;
                            }
                        }
                    }
                }
            }
            if((maxVal !== null || minVal !== null) )
            {
                if(maxVal !== null && minVal !== null && AscFormat.fApproxEqual(minVal, maxVal))
                {
                    if(nPtCount > 1)
                    {
                        minVal -= 0.1;
                        maxVal += 0.1;
                    }
                    else
                    {
                        if(maxVal >= 0)
                        {
                            minVal = null;
                        }
                        else
                        {
                            maxVal = null;
                        }
                    }
                }
                if(maxVal !== null)
                {
                    maxVal -= 0.01;
                }
                if(minVal !== null)
                {
                    minVal += 0.01;
                }
                for(i = 0; i < oSparklineGroup.arrSparklines.length; ++i)
                {
					oSparklineGroup.arrSparklines[i].oCacheView.setMinMaxValAx(minVal, maxVal, oSparklineGroup);
                }
            }
        }
    };

    _this.drawSparkLineGroups = function(oDrawingContext, aSparklineGroups, range, offsetX, offsetY)
    {

        var graphics, i, j, sparkline;
        if(oDrawingContext instanceof AscCommonExcel.CPdfPrinter)
        {
            graphics = oDrawingContext.DocumentRenderer;
        }
        else
        {
            graphics = new AscCommon.CGraphics();
            graphics.init(oDrawingContext.ctx, oDrawingContext.getWidth(0), oDrawingContext.getHeight(0),
                oDrawingContext.getWidth(3)*nSparklineMultiplier, oDrawingContext.getHeight(3)*nSparklineMultiplier);
            graphics.m_oFontManager = AscCommon.g_fontManager;
        }

        for(i = 0; i < aSparklineGroups.length; ++i) {
            var oSparklineGroup = aSparklineGroups[i];

            if(oSparklineGroup.type !== Asc.c_oAscSparklineType.Stacked &&
                (Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.minAxisType || Asc.c_oAscSparklineAxisMinMax.Group === oSparklineGroup.maxAxisType))
            {
                AscFormat.ExecuteNoHistory(function(){
                    _this.checkSparklineGroupMinMaxVal(oSparklineGroup);
                }, _this, []);
            }

            if(oDrawingContext instanceof AscCommonExcel.CPdfPrinter)
            {
                graphics.SaveGrState();
                var _baseTransform;
                if(!oDrawingContext.Transform)
                {
                    _baseTransform = new AscCommon.CMatrix();
                }
                else
                {
                    _baseTransform = oDrawingContext.Transform.CreateDublicate();
                }
                _baseTransform.sx /= nSparklineMultiplier;
                _baseTransform.sy /= nSparklineMultiplier;
                graphics.SetBaseTransform(_baseTransform);
            }
            for(j = 0; j < oSparklineGroup.arrSparklines.length; ++j) {
				sparkline = oSparklineGroup.arrSparklines[j];
				if (!sparkline.checkInRange(range)) {
					continue;
				}

                if (!sparkline.oCacheView) {
					sparkline.oCacheView = new CSparklineView();
					sparkline.oCacheView.initFromSparkline(sparkline, oSparklineGroup, worksheet);
                }


				sparkline.oCacheView.draw(graphics, offsetX, offsetY);


            }
            if(oDrawingContext instanceof AscCommonExcel.CPdfPrinter)
            {
                graphics.SetBaseTransform(null);
                graphics.RestoreGrState();
            }
        }
        if(oDrawingContext instanceof AscCommonExcel.CPdfPrinter)
        {
        }
        else
        {
            oDrawingContext.restore();
        }
    };

    _this.pushToAObjects = function(aDrawing)
    {
        aObjects = [];
        for(var i = 0; i < aDrawing.length; ++i)
        {
            aObjects.push(aDrawing[i]);
        }
    };

    _this.updateDrawingObject = function(bInsert, operType, updateRange) {

        if(!History.CanAddChanges() || History.CanNotAddChanges)
            return;

        var metrics = null;
		var count, bNeedRedraw = false, offset;
            for (var i = 0; i < aObjects.length; i++)
            {
                var obj = aObjects[i];
                    metrics = { from: {}, to: {} };
                    metrics.from.col = obj.from.col; metrics.to.col = obj.to.col;
                    metrics.from.colOff = obj.from.colOff; metrics.to.colOff = obj.to.colOff;
                    metrics.from.row = obj.from.row; metrics.to.row = obj.to.row;
                    metrics.from.rowOff = obj.from.rowOff; metrics.to.rowOff = obj.to.rowOff;

                var bCanMove = false;
                var bCanResize = false;
                var oGraphicObject = obj.graphicObject;
                var nDrawingBaseType = oGraphicObject.getDrawingBaseType();
                switch (nDrawingBaseType)
                {
                    case AscCommon.c_oAscCellAnchorType.cellanchorTwoCell:
                    {
                        bCanMove = true;
                        bCanResize = true;
                        break;
                    }
                    case AscCommon.c_oAscCellAnchorType.cellanchorOneCell:
                    {
                        bCanMove = true;
                        break;
                    }
                }
                if (bInsert)
                {		// Insert
                    switch (operType)
                    {
                        case c_oAscInsertOptions.InsertColumns:
                        {
                            count = updateRange.c2 - updateRange.c1 + 1;
                            // Position
                            if (updateRange.c1 <= obj.from.col)
                            {
                                if(bCanMove)
                                {
                                    metrics.from.col += count;
                                    metrics.to.col += count;
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else if ((updateRange.c1 > obj.from.col) && (updateRange.c1 <= obj.to.col))
                            {
                                if(bCanResize)
                                {
                                    metrics.to.col += count;
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else
                            {
                                metrics = null;
                            }
                            break;
                        }
                        case c_oAscInsertOptions.InsertCellsAndShiftRight:
                        {
                            break;
                        }
                        case c_oAscInsertOptions.InsertRows:
                        {
                            // Position
                            count = updateRange.r2 - updateRange.r1 + 1;

                            if (updateRange.r1 <= obj.from.row)
                            {
                                if(bCanMove)
                                {
                                    metrics.from.row += count;
                                    metrics.to.row += count;
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else if ((updateRange.r1 > obj.from.row) && (updateRange.r1 <= obj.to.row))
                            {
                                if(bCanResize)
                                {
                                    metrics.to.row += count;
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else
                            {
                                metrics = null;
                            }
                            break;
                        }
                        case c_oAscInsertOptions.InsertCellsAndShiftDown:
                        {
                            break;
                        }
                    }
                }
                else
                {				// Delete
                    switch (operType)
                    {
                        case c_oAscDeleteOptions.DeleteColumns:
                        {

                            // Position
                            count = updateRange.c2 - updateRange.c1 + 1;

                            if (updateRange.c1 <= obj.from.col)
                            {
                                // outside
                                if (updateRange.c2 < obj.from.col)
                                {
                                    if(bCanMove)
                                    {
                                        metrics.from.col -= count;
                                        metrics.to.col -= count;
                                    }
                                    else
                                    {
                                        metrics = null;
                                    }
                                }
                                // inside
                                else
                                {
                                    if(bCanResize)
                                    {
                                        metrics.from.col = updateRange.c1;
                                        metrics.from.colOff = 0;

                                        offset = 0;
                                        if (obj.to.col - updateRange.c2 - 1 > 0)
                                        {
                                            offset = obj.to.col - updateRange.c2 - 1;
                                        }
                                        else
                                        {
                                            offset = 1;
                                            metrics.to.colOff = 0;
                                        }
                                        metrics.to.col = metrics.from.col + offset;
                                    }
                                    else
                                    {
                                        metrics = null;
                                    }
                                }
                            }
                            else if ((updateRange.c1 > obj.from.col) && (updateRange.c1 <= obj.to.col))
                            {
                                if(bCanResize)
                                {
                                    // outside
                                    if (updateRange.c2 >= obj.to.col)
                                    {
                                        metrics.to.col = updateRange.c1;
                                        metrics.to.colOff = 0;
                                    }
                                    else
                                    {
                                        metrics.to.col -= count;
                                    }
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else
                            {
                                metrics = null;
                            }
                            break;
                        }
                        case c_oAscDeleteOptions.DeleteCellsAndShiftLeft:
                        {
                            // Range
                            break;
                        }
                        case c_oAscDeleteOptions.DeleteRows:
                        {
                            // Position
                            count = updateRange.r2 - updateRange.r1 + 1;

                            if (updateRange.r1 <= obj.from.row)
                            {
                                // outside
                                if (updateRange.r2 < obj.from.row)
                                {
                                    if(bCanMove)
                                    {
                                        metrics.from.row -= count;
                                        metrics.to.row -= count;
                                    }
                                    else
                                    {
                                        metrics = null;
                                    }
                                }
                                // inside
                                else
                                {
                                    if(bCanResize)
                                    {
                                        metrics.from.row = updateRange.r1;
                                        metrics.from.colOff = 0;

                                        offset = 0;
                                        if (obj.to.row - updateRange.r2 - 1 > 0)
                                            offset = obj.to.row - updateRange.r2 - 1;
                                        else {
                                            offset = 1;
                                            metrics.to.colOff = 0;
                                        }
                                        metrics.to.row = metrics.from.row + offset;
                                    }
                                    else
                                    {
                                        metrics = null;
                                    }
                                }
                            }
                            else if ((updateRange.r1 > obj.from.row) && (updateRange.r1 <= obj.to.row))
                            {
                                if(bCanResize)
                                {
                                    // outside
                                    if (updateRange.r2 >= obj.to.row)
                                    {
                                        metrics.to.row = updateRange.r1;
                                        metrics.to.colOff = 0;
                                    }
                                    else
                                    {
                                        metrics.to.row -= count;
                                    }
                                }
                                else
                                {
                                    metrics = null;
                                }
                            }
                            else
                            {
                                metrics = null;
                            }
                            break;
                        }
                        case c_oAscDeleteOptions.DeleteCellsAndShiftTop:
                        {
                            // Range
                            break;
                        }
                    }
                }

                // Normalize position
                if (metrics)
                {
                    var oFrom = metrics.from;
                    var oTo = metrics.to;
                    if (oFrom.col < 0)
                    {
                        oFrom.col = 0;
                        oFrom.colOff = 0;
                    }

                    if (oTo.col <= 0) {
                        oTo.col = 1;
                        oTo.colOff = 0;
                    }

                    if (oFrom.row < 0) {
                        oFrom.row = 0;
                        oFrom.rowOff = 0;
                    }

                    if (oTo.row <= 0) {
                        oTo.row = 1;
                        oTo.rowOff = 0;
                    }

                    if (oFrom.col === oTo.col) {
                        if(oFrom.colOff > oTo.colOff) {
                            oTo.colOff = oFrom.colOff;
                        }
                    }
                    if (oFrom.row === oTo.row) {
                        if(oFrom.rowOff > oTo.rowOff) {
                            oTo.rowOff = oFrom.rowOff;
                        }
                    }

                    oGraphicObject.setDrawingBaseCoords(oFrom.col, oFrom.colOff, oFrom.row, oFrom.rowOff,
                        oTo.col, oTo.colOff, oTo.row, oTo.rowOff, obj.Pos.X, obj.Pos.Y, obj.ext.cx, obj.ext.cy);
                    oGraphicObject.recalculate();

                    if(oGraphicObject.spPr && oGraphicObject.spPr.xfrm)
                    {
                        var oXfrm = oGraphicObject.spPr.xfrm;
                        oXfrm.setOffX(oGraphicObject.x);
                        oXfrm.setOffY(oGraphicObject.y);
                        oXfrm.setExtX(oGraphicObject.extX);
                        oXfrm.setExtY(oGraphicObject.extY);
                    }
                    oGraphicObject.recalculate();
                    bNeedRedraw = true;
                }
            }
        bNeedRedraw && _this.showDrawingObjects();
    };

    _this.updateDrawingsTransform = function(target) {
        if(false === _this.bUpdateMetrics) {
            return;
        }
        if(!AscCommon.isRealObject(target)) {
            return;
        }
        if (target.target === AscCommonExcel.c_oTargetType.RowResize ||
            target.target === AscCommonExcel.c_oTargetType.ColumnResize) {
            for(var i = 0; i < aObjects.length; ++i) {
                var oDrawingBase = aObjects[i];
                var oGraphicObject = oDrawingBase.graphicObject;
                if(oDrawingBase.checkTarget(target, false)) {
                    oGraphicObject.handleUpdateExtents();
                    oGraphicObject.recalculate();
                }
            }
        }
    };

    _this.updateSizeDrawingObjects = function(target) {
        if(AscCommon.isFileBuild()) {
            return;
        }
        var oGraphicObject;
        var bCheck, bRecalculate;
        if(!History.CanAddChanges() || History.CanNotAddChanges) {
            _this.updateDrawingsTransform(target);
            return;
        }
        var aObjectsForCheck = [], i, drawingObject, bDraw = false;
        if (target.target === AscCommonExcel.c_oTargetType.RowResize ||
            target.target === AscCommonExcel.c_oTargetType.ColumnResize) {
            for (i = 0; i < aObjects.length; i++) {
                drawingObject = aObjects[i];
                bCheck = drawingObject.checkTarget(target, false);
                if (bCheck) {
                    oGraphicObject = drawingObject.graphicObject;
                    bRecalculate = true;
                    if(oGraphicObject){
                        var oldX = oGraphicObject.x;
                        var oldY = oGraphicObject.y;
                        var oldExtX = oGraphicObject.extX;
                        var oldExtY = oGraphicObject.extY;
                        if(drawingObject.Type === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell
                        && drawingObject.editAs !== AscCommon.c_oAscCellAnchorType.cellanchorTwoCell) {
                            if(drawingObject.editAs === AscCommon.c_oAscCellAnchorType.cellanchorAbsolute) {
                                aObjectsForCheck.push({object: drawingObject, coords:  drawingObject._getGraphicObjectCoords()});
                            }
                            else {
                                    oGraphicObject.recalculateTransform();
                                    oGraphicObject.extX = oldExtX;
                                    oGraphicObject.extY = oldExtY;
                                    aObjectsForCheck.push({object: drawingObject, coords:  drawingObject._getGraphicObjectCoords()});
                            }
                        } else if (oGraphicObject.hasSmartArt && oGraphicObject.hasSmartArt()) { // fixme: this is a crutch for resizing columns and rows, fix after normal recalculate smartart
                            if (oGraphicObject.spPr.xfrm.isZeroCh()) {
                                oGraphicObject.updateCoordinatesAfterInternalResize();
                            }
                            oGraphicObject.recalculateTransform();
                            var coefficient = 1;
                            if (oldExtX !== oGraphicObject.extX) {
                                coefficient = oGraphicObject.extX / oldExtX;
                                oGraphicObject.extY *= oGraphicObject.extX / oldExtX;
                            } else if (oldExtY !== oGraphicObject.extY) {
                                coefficient = oGraphicObject.extY / oldExtY;
                                oGraphicObject.extX *= coefficient;
                            }
                            if (coefficient !== 1) {
                                oGraphicObject.changeSize(coefficient, coefficient);
                                oGraphicObject.spPr.xfrm.setChExtX(oGraphicObject.spPr.xfrm.extX);
                                oGraphicObject.spPr.xfrm.setChExtY(oGraphicObject.spPr.xfrm.extY);
                            }
                            aObjectsForCheck.push({object: drawingObject, coords:  drawingObject._getGraphicObjectCoords()});
                        } else {
                            if(oGraphicObject.recalculateTransform) {
                                oGraphicObject.recalculateTransform();
                                var fDelta = 0.01;
                                bRecalculate = false;
                                if(!AscFormat.fApproxEqual(oldX, oGraphicObject.x, fDelta) || !AscFormat.fApproxEqual(oldY, oGraphicObject.y, fDelta)
                                    || !AscFormat.fApproxEqual(oldExtX, oGraphicObject.extX, fDelta) || !AscFormat.fApproxEqual(oldExtY, oGraphicObject.extY, fDelta)){
                                    bRecalculate = true;
                                }
                            }
                            if(bRecalculate) {
                                oGraphicObject.handleUpdateExtents();
                                oGraphicObject.recalculate();
                                bDraw = true;
                            }
                            else {
                                drawingObject.checkBoundsFromTo();
                            }
                        }
                    }
                }
                else {
                    drawingObject.checkBoundsFromTo();
                }
            }
        }
        if(aObjectsForCheck.length > 0) {
            var aId = [];
            for(i = 0; i < aObjectsForCheck.length; ++i) {
                aId.push(aObjectsForCheck[i].object.graphicObject.Get_Id());
            }
            Asc.editor.checkObjectsLock(aId, function (bLock) {
                var i, oObjectToCheck, oGraphicObject;
                var bUpdateDrawingBaseCoords = (bLock === true) && History.CanAddChanges() && !History.CanNotAddChanges;
                for(i = 0; i < aObjectsForCheck.length; ++i) {
                    oObjectToCheck = aObjectsForCheck[i];
                    oGraphicObject = oObjectToCheck.object.graphicObject;
                    if(bUpdateDrawingBaseCoords) {
                        var oC = oObjectToCheck.coords;
                        oGraphicObject.setDrawingBaseCoords(oC.from.col, oC.from.colOff, oC.from.row, oC.from.rowOff,
                            oC.to.col, oC.to.colOff, oC.to.row, oC.to.rowOff,
                            oC.Pos.X, oC.Pos.Y, oC.ext.cx, oC.ext.cy);
                    }
                    oGraphicObject.handleUpdateExtents();
                    if (oGraphicObject.hasSmartArt && oGraphicObject.hasSmartArt()) {
                        oGraphicObject.checkExtentsByDocContent(true, true);
                    }
                    oGraphicObject.recalculate();
                }
                _this.showDrawingObjects();
            });
        }
        else {
            if(bDraw) {
                _this.showDrawingObjects();
            }
        }
    };

    _this.applyMoveResizeRange = function(oRanges) {

        var oChart = null;
        var aSelectedObjects = _this.controller.selection.groupSelection ? _this.controller.selection.groupSelection.selectedObjects : _this.controller.selectedObjects;
        if(aSelectedObjects.length === 1
            && aSelectedObjects[0].getObjectType() === AscDFH.historyitem_type_ChartSpace) {
            oChart = aSelectedObjects[0];
        }
        else {
            return;
        }
        _this.controller.checkSelectedObjectsAndCallback(function () {
            oChart.fillDataFromTrack(oRanges);
        }, [], false, AscDFH.historydescription_ChartDrawingObjects);
    };

    //-----------------------------------------------------------------------------------
    // Graphic object
    //-----------------------------------------------------------------------------------

    _this.addGraphicObject = function(graphic, position, lockByDefault) {

        worksheet.cleanSelection();
        var drawingObject = _this.createDrawingObject();
        drawingObject.graphicObject = graphic;
        graphic.setDrawingBase(drawingObject);

        var ret;
        if (AscFormat.isRealNumber(position)) {
            aObjects.splice(position, 0, drawingObject);
            ret = position;
        }
        else {
            ret = aObjects.length;
            aObjects.push(drawingObject);
        }

        if ( lockByDefault ) {
            Asc.editor.checkObjectsLock([drawingObject.graphicObject.Id], function(result) {});
        }
        worksheet.setSelectionShape(true);

        return ret;
    };


    _this.addOleObject = function(fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, sData, sApplicationId, bSelect, arrImagesForAddToHistory){
        var drawingObject = _this.createDrawingObject();
        drawingObject.worksheet = worksheet;

        var activeCell = worksheet.model.selectionRange.activeCell;
        drawingObject.from.col = activeCell.col;
        drawingObject.from.row = activeCell.row;

        _this.calculateObjectMetrics(drawingObject, nWidthPix, nHeightPix);

        var coordsFrom = _this.calculateCoords(drawingObject.from);
        _this.controller.resetSelection();
        _this.controller.addOleObjectFromParams(pxToMm(coordsFrom.x), pxToMm(coordsFrom.y), fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, sData, sApplicationId, bSelect, arrImagesForAddToHistory);
        worksheet.setSelectionShape(true);
    };

    _this.editOleObject = function(oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, bResize, arrImagesForAddToHistory){
        this.controller.editOleObjectFromParams(oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, bResize, arrImagesForAddToHistory);
    };

    _this.startEditCurrentOleObject = function(){
        this.controller.startEditCurrentOleObject();
    };

    _this.groupGraphicObjects = function() {

        if ( _this.controller.canGroup() ) {

            if(this.controller.checkSelectedObjectsProtection())
            {
                return;
            }
            _this.controller.checkSelectedObjectsAndCallback(_this.controller.createGroup, [], false, AscDFH.historydescription_Spreadsheet_CreateGroup);
            worksheet.setSelectionShape(true);
        }
    };

    _this.unGroupGraphicObjects = function() {

        if ( _this.controller.canUnGroup() ) {
            _this.controller.unGroup();
            worksheet.setSelectionShape(true);
            api.isStartAddShape = false;
        }
    };

    _this.getDrawingBase = function(graphicId) {
        var oDrawing = AscCommon.g_oTableId.Get_ById(graphicId);

        if(oDrawing) {
            if(oDrawing.chart
                && oDrawing.chart.getObjectType
                && oDrawing.chart.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                oDrawing = oDrawing.chart;
            }
            while(oDrawing.group) {
                oDrawing = oDrawing.group;
            }
            if(oDrawing.drawingBase) {
                for (var i = 0; i < aObjects.length; i++) {
                    if ( aObjects[i] === oDrawing.drawingBase )
                        return aObjects[i];
                }
            }
        }
        return null;
    };

    _this.deleteDrawingBase = function(graphicId) {

        var position = null;
        var bRedraw = false;
        for (var i = 0; i < aObjects.length; i++) {
            if ( aObjects[i].graphicObject.Id == graphicId ) {
                aObjects[i].graphicObject.deselect(_this.controller);
                if ( aObjects[i].isChart() )
                    worksheet.endEditChart();
                aObjects.splice(i, 1);
                bRedraw = true;
                position = i;
                break;
            }
        }

        /*if ( bRedraw ) {
         worksheet._endSelectionShape();
         _this.sendGraphicObjectProps();
         _this.showDrawingObjects();
         }*/

        return position;
    };

    _this.checkGraphicObjectPosition = function(x, y, w, h) {

        /*	Принцип:
         true - если перемещение в области или требуется увеличить лист вправо/вниз
         false - наезд на хидеры
         */

        var response = { result: true, x: 0, y: 0 };


        // выход за границу слева или сверху
        if ( y < 0 ) {
            response.result = false;
            response.y = Math.abs(y);
        }
        if ( x < 0 ) {
            response.result = false;
            response.x = Math.abs(x);
        }

        return response;
    };

    _this.resetLockedGraphicObjects = function() {

        for (var i = 0; i < aObjects.length; i++) {
            aObjects[i].graphicObject.lockType = c_oAscLockTypes.kLockTypeNone;
        }
    };

    _this.tryResetLockedGraphicObject = function(id) {

        var bObjectFound = false;
        for (var i = 0; i < aObjects.length; i++) {
            if ( aObjects[i].graphicObject.Id == id ) {
                aObjects[i].graphicObject.lockType = c_oAscLockTypes.kLockTypeNone;
                bObjectFound = true;
                break;
            }
        }
        return bObjectFound;
    };

    _this.getDrawingCanvas = function() {
        return { shapeCtx: api.wb.shapeCtx, shapeOverlayCtx: api.wb.shapeOverlayCtx, autoShapeTrack: api.wb.autoShapeTrack, trackOverlay: api.wb.trackOverlay };
    };

    _this.convertMetric = function(val, from, to) {
        /* Параметры конвертирования (from/to)
         0 - px, 1 - pt, 2 - in, 3 - mm
         */
        return val * ascCvtRatio(from, to);
    };

    _this.convertPixToMM = function(pix)
    {
        return _this.convertMetric(pix, 0, 3);
    };
    _this.getSelectedGraphicObjects = function() {
        return _this.controller.selectedObjects;
    };

    _this.selectedGraphicObjectsExists = function() {
        return _this.controller && _this.controller.selectedObjects.length > 0;
    };

    _this.loadImageRedraw = function(imageUrl) {

        var _image = api.ImageLoader.LoadImage(imageUrl, 1);

        if (null != _image) {
            imageLoaded(_image);
        }
        else {
            _this.asyncImageEndLoaded = function(_image) {
                imageLoaded(_image);
                _this.asyncImageEndLoaded = null;
            }
        }

        function imageLoaded(_image) {
            if ( !_image.Image ) {
                worksheet.model.workbook.handlers.trigger("asc_onError", c_oAscError.ID.UplImageUrl, c_oAscError.Level.NoCritical);
            }
            else
                _this.showDrawingObjects();
        }
    };

    _this.getOriginalImageSize = function() {

        var selectedObjects = _this.controller.selectedObjects;
        if ( (selectedObjects.length == 1) ) {


            if(selectedObjects[0].isImage()){
                var imageUrl = selectedObjects[0].getImageUrl();

                var oImagePr = new Asc.asc_CImgProperty();
                oImagePr.asc_putImageUrl(imageUrl);
                var oSize = oImagePr.asc_getOriginSize(api);
                if(oSize.IsCorrect) {
                    return oSize;
                }
            }
        }
        return new AscCommon.asc_CImageSize( 50, 50, false );
    };

    _this.getSelectionImg = function() {
        return _this.controller.getSelectionImage().asc_getImageUrl();
    };

    _this.sendGraphicObjectProps = function () {
        if (worksheet) {
            worksheet.handlers.trigger("selectionChanged");
        }
    };

    _this.setGraphicObjectProps = function(props) {

        var objectProperties = props;

		var _img;
        if ( !AscCommon.isNullOrEmptyString(objectProperties.ImageUrl) ) {
            if (window['IS_NATIVE_EDITOR']) {
                _this.controller.setGraphicObjectProps( objectProperties );
            } else {
                _img = api.ImageLoader.LoadImage(objectProperties.ImageUrl, 1);

                if (null != _img) {
                    _this.controller.setGraphicObjectProps( objectProperties );
                }
                else {
                    _this.asyncImageEndLoaded = function(_image) {
                        _this.controller.setGraphicObjectProps( objectProperties );
                        _this.asyncImageEndLoaded = null;
                    }
                }
            }
        }
        else if ( objectProperties.ShapeProperties && objectProperties.ShapeProperties.fill && objectProperties.ShapeProperties.fill.fill &&
            !AscCommon.isNullOrEmptyString(objectProperties.ShapeProperties.fill.fill.url) ) {

            if (window['IS_NATIVE_EDITOR']) {
                _this.controller.setGraphicObjectProps( objectProperties );
            } else {
                _img = api.ImageLoader.LoadImage(objectProperties.ShapeProperties.fill.fill.url, 1);
                if ( null != _img ) {
                    _this.controller.setGraphicObjectProps( objectProperties );
                }
                else {
                    _this.asyncImageEndLoaded = function(_image) {
                        _this.controller.setGraphicObjectProps( objectProperties );
                        _this.asyncImageEndLoaded = null;
                    }
                }
            }
        }
        else if ( objectProperties.ShapeProperties && objectProperties.ShapeProperties.textArtProperties &&
            objectProperties.ShapeProperties.textArtProperties.Fill && objectProperties.ShapeProperties.textArtProperties.Fill.fill &&
            !AscCommon.isNullOrEmptyString(objectProperties.ShapeProperties.textArtProperties.Fill.fill.url) ) {

            if (window['IS_NATIVE_EDITOR']) {
                _this.controller.setGraphicObjectProps( objectProperties );
            } else {
                _img = api.ImageLoader.LoadImage(objectProperties.ShapeProperties.textArtProperties.Fill.fill.url, 1);
                if ( null != _img ) {
                    _this.controller.setGraphicObjectProps( objectProperties );
                }
                else {
                    _this.asyncImageEndLoaded = function(_image) {
                        _this.controller.setGraphicObjectProps( objectProperties );
                        _this.asyncImageEndLoaded = null;
                    }
                }
            }
        }
        else if(AscFormat.isRealNumber(objectProperties.ChangeLevel)) {
            this.setGraphicObjectLayer(objectProperties.ChangeLevel);
        }
        else {
            objectProperties.ImageUrl = null;
            _this.applySliderCallbacks(
                function(){
                    _this.controller.setGraphicObjectProps(props);
                },
                function() {
                    _this.controller.setGraphicObjectPropsCallBack(props);
                    _this.controller.startRecalculate();
                    _this.sendGraphicObjectProps();
                }
            );
        }
    };

    _this.applySliderCallbacks = function(fStartCallback, fApplyCallback) {

        var Point =  History.Points[History.Index];
        if(!api.noCreatePoint || api.exucuteHistory) {
            if( !api.noCreatePoint && !api.exucuteHistory && api.exucuteHistoryEnd) {
                if(_this.nCurPointItemsLength === -1 && !(Point && Point.Items.length === 0)) {
                    fStartCallback();
                }
                else{
                    fApplyCallback();
                }
                api.exucuteHistoryEnd = false;
                _this.nCurPointItemsLength = -1;
            }
            else {
                fStartCallback();
            }
            if(api.exucuteHistory) {
                Point =  History.Points[History.Index];
                if(Point) {
                    _this.nCurPointItemsLength = Point.Items.length;
                }
                else {
                    _this.nCurPointItemsLength = -1;
                }
                api.exucuteHistory = false;
            }
            if(api.exucuteHistoryEnd) {
                api.exucuteHistoryEnd = false;
            }
        }
        else {
            if(Point && Point.Items.length > 0) {
                if(this.nCurPointItemsLength > -1) {
                    var nBottomIndex = - 1;
                    for ( var Index = Point.Items.length - 1; Index > nBottomIndex; Index-- ) {
                        var Item = Point.Items[Index];
                        if(!Item.Class.Read_FromBinary2)
                            Item.Class.Undo( Item.Type, Item.Data, Item.SheetId );
                        else
                            Item.Class.Undo(Item.Data);
                    }
                    Point.Items.length = 0;
                    _this.controller.setSelectionState(Point.SelectionState);
                    fApplyCallback();
                }
                else {
                    fStartCallback();
                    var Point =  History.Points[History.Index];
                    if(Point) {
                        _this.nCurPointItemsLength = Point.Items.length;
                    }
                    else {
                        _this.nCurPointItemsLength = -1;
                    }
                }
            }
        }
        api.exucuteHistoryEnd = false;
    };

    _this.showChartSettings = function() {
        api.wb.handlers.trigger("asc_onShowChartDialog", true);
    };

    _this.setDrawImagePlaceParagraph = function(element_id, props) {
        _this.drawingDocument.InitGuiCanvasTextProps(element_id);
        _this.drawingDocument.DrawGuiCanvasTextProps(props);
    };

    //-----------------------------------------------------------------------------------
    // Graphic object mouse & keyboard events
    //-----------------------------------------------------------------------------------


    _this.onStartUserAction = function() {
        if(!AscCommon.SpeechWorker.isEnabled) {
            return;
        }
        this.HistoryIndex = AscCommon.History.Index;
        this.BeforeActionSelectionState = this.controller.getSelectionState();
    };

    _this.onEndUserAction = function() {
        if(!this.BeforeActionSelectionState) {
            return;
        }
        if(this.HistoryIndex !== AscCommon.History.Index) {
            return;
        }

        const oEndSelectionState = this.controller.getSelectionState();
        const oBeforeSelectionState = this.BeforeActionSelectionState;
        this.BeforeActionSelectionState = null;

        const oSpeechData = AscCommon.getSpeechDescription(oBeforeSelectionState, oEndSelectionState);
        if(oSpeechData) {
            AscCommon.SpeechWorker.speech(
                oSpeechData.type,
                oSpeechData.obj
            );
        }
    };

    _this.graphicObjectMouseDown = function(e, x, y) {
        var offsets = _this.drawingArea.getOffsets(x, y, true);
        if ( offsets )
            _this.controller.onMouseDown( e, pxToMm(x - offsets.x), pxToMm(y - offsets.y) );

        //_this.private_UpdateCursorXY();
    };

    _this.graphicObjectMouseMove = function(e, x, y) {
        e.IsLocked = e.isLocked;

        _this.lastX = x;
        _this.lastY = y;
        var offsets = _this.drawingArea.getOffsets(x, y, true);
        if ( offsets ) {

            var fX = pxToMm(x - offsets.x);
            var fY = pxToMm(y - offsets.y);
            _this.controller.onMouseMove( e, fX, fY );
            if(worksheet && worksheet.model) {
                AscCommon.CollaborativeEditing.Check_ForeignCursorsLabels(fX, fY, worksheet.model.Id);
            }
        }
    };

    _this.graphicObjectMouseUp = function(e, x, y) {
        var offsets = _this.drawingArea.getOffsets(x, y, true);
        if ( offsets )
            _this.controller.onMouseUp( e, pxToMm(x - offsets.x), pxToMm(y - offsets.y) );
        //_this.private_UpdateCursorXY();
    };

    _this.isPointInDrawingObjects3 = function(x, y, page, bSelected, bText) {
        var offsets = _this.drawingArea.getOffsets(x, y, true);
        if ( offsets )
            return _this.controller.isPointInDrawingObjects3(pxToMm(x - offsets.x), pxToMm(y - offsets.y), page, bSelected, bText );
        return false;
    };

    // keyboard
    _this.private_UpdateCursorXY = function (bUpdateX, bUpdateY) {
        let oController = _this.controller;
        if(oController) {
            let oContent = oController.getTargetDocContent();
            if(oContent) {
                if (true === oContent.Selection.Use && true !== oContent.Selection.Start)
                    Asc.editor.sendEvent("asc_onSelectionEnd");
                else if (!oContent.Selection.Use)
                    Asc.editor.sendEvent("asc_onCursorMove");
                return;
            }
        }
        Asc.editor.sendEvent("asc_onSelectionEnd");
    };
    _this.graphicObjectKeyDown = function(e) {
        Asc.editor.sendEvent("asc_onBeforeKeyDown", e);
        let nHistoryIndex = AscCommon.History.Index;
        let ret = _this.controller.onKeyDown( e );
        // if(nHistoryIndex === AscCommon.History.Index) {
        //     _this.private_UpdateCursorXY();
        // }
        Asc.editor.sendEvent("asc_onKeyDown", e);
        return ret;
    };
    _this.graphicObjectKeyUp = function(e) {
        return _this.controller.onKeyUp( e );
    };

    _this.graphicObjectKeyPress = function(e) {

        e.KeyCode = e.keyCode;
        e.CtrlKey = e.metaKey || e.ctrlKey;
        e.AltKey = e.altKey;
        e.ShiftKey = e.shiftKey;
        e.Which = e.which;

        this.onStartUserAction();
        let ret = _this.controller.onKeyPress( e );
        this.onEndUserAction();
        return ret;
    };

    //-----------------------------------------------------------------------------------
    // Asc
    //-----------------------------------------------------------------------------------

    _this.cleanWorksheet = function() {
        if (worksheet) {
            var model = worksheet.model;
            History.Clear();
            for (var i = 0; i < aObjects.length; i++) {
                aObjects[i].graphicObject.deleteDrawingBase();
            }
            aObjects.length = 0;

            var oAllRange = model.getRange3(0, 0, model.getRowsCount(), model.getColsCount());
            oAllRange.cleanAll();

            worksheet.endEditChart();
        }
    };

    _this.getWordChartObject = function() {
        for (var i = 0; i < aObjects.length; i++) {
            var drawingObject = aObjects[i];

            if ( drawingObject.isChart() ) {
                var oChartSpace = drawingObject.graphicObject;
                if(!oChartSpace.recalcInfo.recalculateTransform) {
                    if(oChartSpace.spPr && oChartSpace.spPr.xfrm) {
                        oChartSpace.spPr.xfrm.extX = oChartSpace.extX;
                        oChartSpace.spPr.xfrm.extY = oChartSpace.extY;
                    }
                }
                return new asc_CChartBinary(oChartSpace);
            }
        }
        return null;
    };

    _this.getAscChartObject = function(bNoLock) {

        var settings;
        if(api.isChartEditor)
        {
            if (!aObjects.length) {
                return null;
            }
            return _this.controller.getPropsFromChart(aObjects[0].graphicObject);
        }
        settings = _this.controller.getChartProps();
        if ( !settings )
        {
            settings = new Asc.asc_ChartSettings();
            settings.putRanges(worksheet.getSelectionRangeValues(true, true));

            settings.putStyle(2);
            settings.putType(Asc.c_oAscChartTypeSettings.lineNormal);
            settings.putTitle(Asc.c_oAscChartTitleShowSettings.noOverlay);

            var vert_axis_settings = new AscCommon.asc_ValAxisSettings();
            settings.addVertAxesProps(vert_axis_settings);
            vert_axis_settings.setDefault();
            vert_axis_settings.putLabel(Asc.c_oAscChartVertAxisLabelShowSettings.none);
            vert_axis_settings.putGridlines(Asc.c_oAscGridLinesSettings.major);
            var hor_axis_settings = new AscCommon.asc_CatAxisSettings();
            settings.addHorAxesProps(hor_axis_settings);
            hor_axis_settings.setDefault();
            hor_axis_settings.putLabel(Asc.c_oAscChartHorAxisLabelShowSettings.none);
            hor_axis_settings.putGridlines(Asc.c_oAscGridLinesSettings.none);
            settings.putDataLabelsPos(Asc.c_oAscChartDataLabelsPos.none);
            settings.putSeparator(",");
            settings.putLine(true);
            settings.putShowMarker(false);
            settings.putView3d(null);
        }
        else{
            if(true !== bNoLock){
                this.controller.checkSelectedObjectsAndFireCallback(function(){});
            }
        }
        return settings;
    };

    //-----------------------------------------------------------------------------------
    // Selection
    //-----------------------------------------------------------------------------------

    _this.selectDrawingObjectRange = function(drawing) {
		worksheet.cleanSelection();
        worksheet.endEditChart();
        drawing.fillSelectedRanges(worksheet);
        if (worksheet.isChartAreaEditMode || (api && api.isShowVisibleAreaOleEditor)) {
            worksheet._drawSelection();
        }
    };

    _this.unselectDrawingObjects = function() {

        worksheet.endEditChart();
        _this.controller.resetSelectionState();
    };

    _this.getDrawingObject = function(id) {

        for (var i = 0; i < aObjects.length; i++) {
            if (aObjects[i].graphicObject.Id == id)
                return aObjects[i];
        }
        return null;
    };

    _this.getGraphicSelectionType = function(id) {

        var selected_objects, selection, controller = _this.controller;
        if(controller.selection.groupSelection)
        {
            selected_objects = controller.selection.groupSelection.selectedObjects;
            selection = controller.selection.groupSelection.selection;
        }
        else
        {
            selected_objects = controller.selectedObjects;
            selection = controller.selection;
        }
        if(selection.chartSelection && selection.chartSelection.selection.textSelection)
        {
            return c_oAscSelectionType.RangeChartText;
        }
        if(selection.textSelection)
        {
            return c_oAscSelectionType.RangeShapeText;
        }
        if(selected_objects[0] )
        {
            if(selected_objects[0].getObjectType() === AscDFH.historyitem_type_ChartSpace && selected_objects.length === 1)
                return c_oAscSelectionType.RangeChart;

            if(selected_objects[0].getObjectType() === AscDFH.historyitem_type_ImageShape)
                return c_oAscSelectionType.RangeImage;
            if(selected_objects[0].getObjectType() === AscDFH.historyitem_type_SlicerView)
                return c_oAscSelectionType.RangeSlicer;

            return c_oAscSelectionType.RangeShape;

        }
        return undefined;
    };

    //-----------------------------------------------------------------------------------
    // Position
    //-----------------------------------------------------------------------------------

    _this.getCurrentDrawingMacrosName = function() {
        return _this.controller.getCurrentDrawingMacrosName();
    };
    _this.assignMacrosToCurrentDrawing = function(sGuid) {
        _this.controller.assignMacrosToCurrentDrawing(sGuid);
    };
    _this.setGraphicObjectLayer = function(layerType) {
        _this.controller.setGraphicObjectLayer(layerType);
    };

    _this.setGraphicObjectAlign = function(alignType) {
        _this.controller.setGraphicObjectAlign(alignType);
    };
    _this.distributeGraphicObjectHor = function() {
        _this.controller.distributeGraphicObjectHor();
    };

    _this.distributeGraphicObjectVer = function() {
        _this.controller.distributeGraphicObjectVer();
    };

    _this.getSelectedDrawingObjectsCount = function() {
        var selectedObjects = _this.controller.selection.groupSelection ? this.controller.selection.groupSelection.selectedObjects : this.controller.selectedObjects;
        return selectedObjects.length;
    };

    _this.checkCursorPlaceholder = function (x, y) {
        if (window["IS_NATIVE_EDITOR"]) {
            return null;
        }

        const oWS = this.getWorksheet();
        const oDrawingDocument = oWS.getDrawingDocument();
        const nPage = oWS.workbook.model.nActive;
        const oContext = oWS.workbook.trackOverlay.m_oContext;

        const oRect = {};
        const nLeft = 2 * oWS.cellsLeft - oWS._getColLeft(oWS.visibleRange.c1);
        const nTop = 2 * oWS.cellsTop - oWS._getRowTop(oWS.visibleRange.r1);
        oRect.left   = nLeft;
        oRect.right  = nLeft + AscCommon.AscBrowser.convertToRetinaValue(oContext.canvas.width);
        oRect.top    = nTop;
        oRect.bottom = nTop + AscCommon.AscBrowser.convertToRetinaValue(oContext.canvas.height);

        const nPXtoMM = Asc.getCvtRatio(0/*mm*/, 3/*px*/, oWS._getPPIX());
        const oOffsets = oWS.objectRender.drawingArea.getOffsets(x, y);
        let nX = x;
        let nY = y;
        if (oOffsets) {
            nX -= oOffsets.x;
            nY -= oOffsets.y;
        }
        nX *= nPXtoMM;
        nY *= nPXtoMM;

        return oDrawingDocument.placeholders.onPointerMove({X: nX, Y: nY, Page: nPage}, oRect, oContext.canvas.width * nPXtoMM, oContext.canvas.height * nPXtoMM);
    };
    
    _this.checkCursorDrawingObject = function(x, y) {

        let offsets = _this.drawingArea.getOffsets(x, y);
	    let objectInfo = null;
	    let oApi = Asc.editor || editor;
        if ( offsets ) {
            let graphicObjectInfo = _this.controller.isPointInDrawingObjects( pxToMm(x - offsets.x), pxToMm(y - offsets.y) );

            if ( graphicObjectInfo && graphicObjectInfo.objectId ) {
	            objectInfo = { cursor: null, id: null, object: null };
                objectInfo.object = _this.getDrawingBase(graphicObjectInfo.objectId);
                if(objectInfo.object) {
                    objectInfo.id = graphicObjectInfo.objectId;
                    let sCursorType = graphicObjectInfo.cursorType;
                    if(oApi) {
                        if(!oApi.isShowShapeAdjustments()) {
                            if(sCursorType !== "text") {
                                sCursorType = "default";
                            }
                        }
                    }
                    objectInfo.cursor = sCursorType;
                    objectInfo.hyperlink = graphicObjectInfo.hyperlink;
                    objectInfo.macro = graphicObjectInfo.macro;
                    objectInfo.tooltip = graphicObjectInfo.tooltip;
                }
            }
        }
		if(oApi.isFormatPainterOn()) {
			if(objectInfo) {
				let oData = oApi.getFormatPainterData();
				if(oData && oData.isDrawingData()) {
					objectInfo.cursor = AscCommon.Cursors.ShapeCopy;
				}
			}
		}

	    if(oApi.isInkDrawerOn()) {
		    if(objectInfo) {
			    objectInfo.cursor = "default";
		    }
	    }
        return objectInfo;
    };

    _this.getPositionInfo = function(x, y) {

        var info = new CCellObjectInfo();

        var tmp = worksheet._findColUnderCursor(x, true);
        if (tmp) {
            info.col = tmp.col;
            info.colOff = pxToMm(x - worksheet._getColLeft(info.col));
        }
        tmp = worksheet._findRowUnderCursor(y, true);
        if (tmp) {
            info.row = tmp.row;
            info.rowOff = pxToMm(y - worksheet._getRowTop(info.row));
        }

        return info;
    };


    _this.checkCurrentTextObjectExtends = function()
    {
        var oController = this.controller;
        if(oController)
        {
            var oTargetTextObject = AscFormat.getTargetTextObject(oController);
            if(oTargetTextObject.checkExtentsByDocContent)
            {
                oTargetTextObject.checkExtentsByDocContent(true, true);
            }
        }
    };

    _this.beginCompositeInput = function(){

        _this.controller.CreateDocContent();
        _this.drawingDocument.TargetStart();
        _this.drawingDocument.TargetShow();
        var oContent = _this.controller.getTargetDocContent(true);
        if (!oContent) {
            return false;
        }
        var oPara = oContent.GetCurrentParagraph();
        if (!oPara) {
            return false;
        }
        if (true === oContent.IsSelectionUse())
            oContent.Remove(1, true, false, true);
        var oRun = oPara.Get_ElementByPos(oPara.Get_ParaContentPos(false, false));
        if (!oRun || !(oRun instanceof ParaRun)) {
            return false;
        }

        _this.CompositeInput = new AscWord.RunCompositeInput_Old(oRun);

        oRun.Set_CompositeInput(_this.CompositeInput);
        _this.controller.startRecalculate();
        _this.sendGraphicObjectProps();
    };

    _this.Begin_CompositeInput = function(){
        if(!_this.controller.canEdit()) {
            return;
        }
        if(_this.controller.checkSelectedObjectsProtectionText()) {
            return;
        }
        Asc.editor.sendEvent("asc_onUserActionStart");
        History.Create_NewPoint(AscDFH.historydescription_Document_CompositeInput);
        _this.beginCompositeInput();
        _this.controller.recalculateCurPos(true, true);
        Asc.editor.sendEvent("asc_onUserActionEnd");
    };


    _this.addCompositeText = function(nCharCode){

        if (null === _this.CompositeInput)
            return;

        var oRun = _this.CompositeInput.Run;
        var nPos = _this.CompositeInput.Pos + _this.CompositeInput.Length;
        var oChar;

        if (para_Math_Run === oRun.Type)
        {
            oChar = new CMathText();
            oChar.add(nCharCode);
        }
        else
        {
            if (32 == nCharCode || 12288 == nCharCode)
                oChar = new AscWord.CRunSpace();
            else
                oChar = new AscWord.CRunText(nCharCode);
        }

        oRun.Add_ToContent(nPos, oChar, true);
        _this.CompositeInput.Length++;
    };
    _this.Add_CompositeText = function(nCharCode)
    {
        if (null === _this.CompositeInput)
            return;
        Asc.editor.sendEvent("asc_onUserActionStart");
        History.Create_NewPoint(AscDFH.historydescription_Document_CompositeInputReplace);
        _this.addCompositeText(nCharCode);
        _this.checkCurrentTextObjectExtends();
        _this.controller.recalculate();
        _this.controller.recalculateCurPos(true, true);
        _this.controller.updateSelectionState();
        Asc.editor.sendEvent("asc_onUserActionEnd");
    };

    _this.removeCompositeText = function(nCount){
        if (null === _this.CompositeInput)
            return;

        var oRun = _this.CompositeInput.Run;
        var nPos = _this.CompositeInput.Pos + _this.CompositeInput.Length;

        var nDelCount = Math.max(0, Math.min(nCount, _this.CompositeInput.Length, oRun.Content.length, nPos));
        oRun.Remove_FromContent(nPos - nDelCount, nDelCount, true);
        _this.CompositeInput.Length -= nDelCount;
    };

    _this.Remove_CompositeText = function(nCount){
        _this.removeCompositeText(nCount);
        _this.checkCurrentTextObjectExtends();
        _this.controller.recalculate();
        _this.controller.recalculateCurPos(true, true);
        _this.controller.updateSelectionState();
    };
    _this.Replace_CompositeText = function(arrCharCodes)
    {
        if (null === _this.CompositeInput)
            return;
        Asc.editor.sendEvent("asc_onUserActionStart");
        History.Create_NewPoint(AscDFH.historydescription_Document_CompositeInputReplace);
        _this.removeCompositeText(_this.CompositeInput.Length);
        for (var nIndex = 0, nCount = arrCharCodes.length; nIndex < nCount; ++nIndex)
        {
            _this.addCompositeText(arrCharCodes[nIndex]);
        }
        _this.checkCurrentTextObjectExtends();
		_this.controller.startRecalculate();
        _this.controller.recalculateCurPos(true, true);
        _this.controller.updateSelectionState();
        Asc.editor.sendEvent("asc_onUserActionEnd");
    };
    _this.Set_CursorPosInCompositeText = function(nPos)
    {
        if (null === _this.CompositeInput)
            return;
        var oRun = _this.CompositeInput.Run;

        var nInRunPos = Math.max(Math.min(_this.CompositeInput.Pos + nPos, _this.CompositeInput.Pos + _this.CompositeInput.Length, oRun.Content.length), _this.CompositeInput.Pos);
        oRun.State.ContentPos = nInRunPos;
        _this.controller.recalculateCurPos(true, true);
        _this.controller.updateSelectionState();
    };
    _this.Get_CursorPosInCompositeText = function()
    {
        if (null === _this.CompositeInput)
            return 0;

        var oRun = _this.CompositeInput.Run;
        var nInRunPos = oRun.State.ContentPos;
        var nPos = Math.min(_this.CompositeInput.Length, Math.max(0, nInRunPos - _this.CompositeInput.Pos));
        return nPos;
    };
    _this.End_CompositeInput = function()
    {
        if (null === _this.CompositeInput)
            return;

        var oRun = _this.CompositeInput.Run;
        oRun.Set_CompositeInput(null);
        _this.CompositeInput = null;
        var oController = _this.controller;
        if(oController)
        {
            var oTargetTextObject = AscFormat.getTargetTextObject(oController);
            if(oTargetTextObject && oTargetTextObject.txWarpStructNoTransform)
            {
                oTargetTextObject.recalculateContent();
            }
        }

        _this.controller.recalculateCurPos(true, true);
        _this.sendGraphicObjectProps();
        _this.showDrawingObjects();
    };
    _this.Get_MaxCursorPosInCompositeText = function()
    {
        if (null === _this.CompositeInput)
            return 0;

        return _this.CompositeInput.Length;
    };


    //-----------------------------------------------------------------------------------
    // Private Misc Methods
    //-----------------------------------------------------------------------------------

    function ascCvtRatio(fromUnits, toUnits) {
        return asc.getCvtRatio( fromUnits, toUnits, drawingCtx.getPPIX() );
    }

    function mmToPx(val) {
        return val * ascCvtRatio(3, 0);
    }

    function pxToMm(val) {
        return val * ascCvtRatio(0, 3);
    }
}

	DrawingObjects.prototype.calculateCoords = function (cell) {
        var ws = this.getWorksheet();
		var coords = {x: 0, y: 0};
		//0 - px, 1 - pt, 2 - in, 3 - mm
		if (cell && ws) {
			var rowHeight = ws.getRowHeight(cell.row, 3);
			var colWidth = ws.getColumnWidth(cell.col, 3);
			var resultRowOff = cell.rowOff > rowHeight ? rowHeight : cell.rowOff;
			var resultColOff = cell.colOff > colWidth ? colWidth : cell.colOff;
			coords.y = ws._getRowTop(cell.row) + ws.objectRender.convertMetric(resultRowOff, 3, 0) - ws._getRowTop(0);
			coords.x = ws._getColLeft(cell.col) + ws.objectRender.convertMetric(resultColOff, 3, 0) - ws._getColLeft(0);
            coords.x = ws.checkRtl(coords.x, undefined, undefined, true);
		}
		return coords;
	};
    DrawingObjects.prototype.getDocumentPositionBinary = function() {
        if(this.controller) {
            var oPosition = this.controller.getDocumentPositionForCollaborative();
            if(!oPosition) {
                return "";
            }
            //console.log("POSITION: " + oPosition.Position);
            var oWriter = new AscCommon.CMemory(true);
            oWriter.CheckSize(50);
            return AscCommon.CollaborativeEditing.GetDocumentPositionBinary(oWriter, oPosition);
        }
        return "";
    };
function ClickCounter() {
    this.x = 0;
    this.y = 0;

    this.lastX = -1000;
    this.lastY = -1000;

    this.button = 0;
    this.time = 0;
    this.clickCount = 0;
    this.log = false;
}
ClickCounter.prototype.mouseDownEvent = function(x, y, button) {
    var currTime = getCurrentTime();

    if (undefined === button)
        button = 0;

	var _eps = 3 * global_mouseEvent.KoefPixToMM;

    var oApi = Asc && Asc.editor;
    if(oApi && oApi.isMobileVersion && (!window["NATIVE_EDITOR_ENJINE"]))
    {
        _eps *= 2;
    }
	if ((Math.abs(global_mouseEvent.X - global_mouseEvent.LastX) > _eps) || (Math.abs(global_mouseEvent.Y - global_mouseEvent.LastY) > _eps))
	{
		// not only move!!! (touch - fast click in different places)
		global_mouseEvent.LastClickTime = -1;
		global_mouseEvent.ClickCount    = 0;
	}

	this.x = x;
	this.y = y;

	if (this.button === button && Math.abs(this.x - this.lastX) <= _eps && Math.abs(this.y - this.lastY) <= _eps && (currTime - this.time < 500)) {
        ++this.clickCount;
    } else {
        this.clickCount = 1;
    }

    this.lastX = this.x;
	this.lastY = this.y;

    if (this.log) {
        console.log("-----");
        console.log("x-> " + this.x + " : " + x);
        console.log("y-> " + this.y + " : " + y);
        console.log("Time: " + (currTime - this.time));
        console.log("Count: " + this.clickCount);
        console.log("");
    }
    this.time = currTime;
};
ClickCounter.prototype.mouseMoveEvent = function(x, y) {
    if (this.x !== x || this.y !== y) {
        this.x = x;
        this.y = y;
        this.clickCount = 0;

        if (this.log) {
            console.log("Reset counter");
        }
    }
};
ClickCounter.prototype.getClickCount = function() {
    return this.clickCount;
};

//--------------------------------------------------------export----------------------------------------------------
    var prot;
    window['AscFormat'] = window['AscFormat'] || {};
    window['Asc'] = window['Asc'] || {};
    window['AscFormat'].isObject = isObject;
    window['AscFormat'].CCellObjectInfo = CCellObjectInfo;

    window["Asc"]["asc_CChartBinary"] = window["Asc"].asc_CChartBinary = asc_CChartBinary;
    prot = asc_CChartBinary.prototype;
    prot["asc_getBinary"] = prot.asc_getBinary;
    prot["asc_setBinary"] = prot.asc_setBinary;
    prot["asc_getThemeBinary"] = prot.asc_getThemeBinary;
    prot["asc_setThemeBinary"] = prot.asc_setThemeBinary;
    prot["asc_setColorMapBinary"] = prot.asc_setColorMapBinary;
    prot["asc_getColorMapBinary"] = prot.asc_getColorMapBinary;

    window["AscFormat"].asc_CChartSeria = asc_CChartSeria;
    prot = asc_CChartSeria.prototype;
    prot["asc_getValFormula"] = prot.asc_getValFormula;
    prot["asc_setValFormula"] = prot.asc_setValFormula;
    prot["asc_getxValFormula"] = prot.asc_getxValFormula;
    prot["asc_setxValFormula"] = prot.asc_setxValFormula;
    prot["asc_getCatFormula"] = prot.asc_getCatFormula;
    prot["asc_setCatFormula"] = prot.asc_setCatFormula;
    prot["asc_getTitle"] = prot.asc_getTitle;
    prot["asc_setTitle"] = prot.asc_setTitle;
    prot["asc_getTitleFormula"] = prot.asc_getTitleFormula;
    prot["asc_setTitleFormula"] = prot.asc_setTitleFormula;
    prot["asc_getMarkerSize"] = prot.asc_getMarkerSize;
    prot["asc_setMarkerSize"] = prot.asc_setMarkerSize;
    prot["asc_getMarkerSymbol"] = prot.asc_getMarkerSymbol;
    prot["asc_setMarkerSymbol"] = prot.asc_setMarkerSymbol;
    prot["asc_getFormatCode"] = prot.asc_getFormatCode;
    prot["asc_setFormatCode"] = prot.asc_setFormatCode;

    window["AscFormat"].DrawingObjects = DrawingObjects;
    window["AscFormat"].DrawingBase = DrawingBase;
    window["AscFormat"].ClickCounter = ClickCounter;
    window["AscFormat"].aSparklinesStyles = aSparklinesStyles;
    window["AscFormat"].CSparklineView = CSparklineView;
    window["AscFormat"].CCellObjectInfo = CCellObjectInfo;
})(window);
