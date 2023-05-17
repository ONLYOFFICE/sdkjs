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

    var drawingsChangesMap = window['AscDFH'].drawingsChangesMap;
    var drawingContentChanges = window['AscDFH'].drawingContentChanges;

    drawingsChangesMap[AscDFH.historyitem_DLbl_SetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetDLblPos] = function(oClass, value) {
        oClass.dLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetLayout] = function(oClass, value) {
        oClass.layout = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetSeparator] = function(oClass, value) {
        oClass.separator = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowBubbleSize] = function(oClass, value) {
        oClass.showBubbleSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowCatName] = function(oClass, value) {
        oClass.showCatName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowLegendKey] = function(oClass, value) {
        oClass.showLegendKey = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowPercent] = function(oClass, value) {
        oClass.showPercent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowSerName] = function(oClass, value) {
        oClass.showSerName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowVal] = function(oClass, value) {
        oClass.showVal = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetShowDLblsRange] = function(oClass, value) {
        oClass.showDlblsRange = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetTx] = function(oClass, value) {
        if(oClass.tx && oClass.tx.strRef || value && value.strRef) {
            oClass.onChangeDataRefs();
        }
        oClass.tx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbl_SetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PlotArea_SetCatAx] = function(oClass, value) {
        oClass.catAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PlotArea_SetDateAx] = function(oClass, value) {
        oClass.dateAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PlotArea_SetDTable] = function(oClass, value) {
        oClass.dTable = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PlotArea_SetLayout] = function(oClass, value) {
        oClass.layout = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PlotArea_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CommonChartFormat_SetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CommonChart_SetDlbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_Set3D] = function(oClass, value) {
        oClass.b3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetGapDepth] = function(oClass, value) {
        oClass.gapDepth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetShape] = function(oClass, value) {
        oClass.shape = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetBarDir] = function(oClass, value) {
        oClass.barDir = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetGapWidth] = function(oClass, value) {
        oClass.gapWidth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetGrouping] = function(oClass, value) {
        oClass.grouping = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetOverlap] = function(oClass, value) {
        oClass.overlap = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetSerLines] = function(oClass, value) {
        oClass.serLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CommonChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaChart_SetDropLines] = function(oClass, value) {
        oClass.dropLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaChart_SetGrouping] = function(oClass, value) {
        oClass.grouping = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetErrBars] = function(oClass, value) {
        oClass.errBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetPictureOptions] = function(oClass, value) {
        oClass.pictureOptions = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AreaSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetAuto] = function(oClass, value) {
        oClass.auto = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetAxId] = function(oClass, value) {
        oClass.internalSetAxId(value);
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetAxPos] = function(oClass, value) {
        oClass.axPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetCrossAx] = function(oClass, value) {
        oClass.crossAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetCrosses] = function(oClass, value) {
        oClass.crosses = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetCrossesAt] = function(oClass, value) {
        oClass.crossesAt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetLblAlgn] = function(oClass, value) {
        oClass.lblAlgn = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetLblOffset] = function(oClass, value) {
        oClass.lblOffset = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetMajorGridlines] = function(oClass, value) {
        oClass.majorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetMajorTickMark] = function(oClass, value) {
        oClass.majorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetMinorGridlines] = function(oClass, value) {
        oClass.minorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetMinorTickMark] = function(oClass, value) {
        oClass.minorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetNoMultiLvlLbl] = function(oClass, value) {
        oClass.noMultiLvlLbl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetScaling] = function(oClass, value) {
        oClass.scaling = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetTickLblPos] = function(oClass, value) {
        oClass.tickLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetTickLblSkip] = function(oClass, value) {
        oClass.tickLblSkip = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetTickMarkSkip] = function(oClass, value) {
        oClass.tickMarkSkip = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetTitle] = function(oClass, value) {
        oClass.title = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_CatAxSetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxAuto] = function(oClass, value) {
        oClass.auto = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxAxId] = function(oClass, value) {
        oClass.internalSetAxId(value);
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxAxPos] = function(oClass, value) {
        oClass.axPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxBaseTimeUnit] = function(oClass, value) {
        oClass.baseTimeUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxCrossAx] = function(oClass, value) {
        oClass.crossAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxCrosses] = function(oClass, value) {
        oClass.crosses = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxCrossesAt] = function(oClass, value) {
        oClass.crossesAt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxLblOffset] = function(oClass, value) {
        oClass.lblOffset = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMajorGridlines] = function(oClass, value) {
        oClass.majorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMajorTickMark] = function(oClass, value) {
        oClass.majorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMajorTimeUnit] = function(oClass, value) {
        oClass.majorTimeUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMajorUnit] = function(oClass, value) {
        oClass.majorUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMajorGridlines] = function(oClass, value) {
        oClass.majorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMinorTickMark] = function(oClass, value) {
        oClass.minorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMinorTimeUnit] = function(oClass, value) {
        oClass.minorTimeUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxMinorUnit] = function(oClass, value) {
        oClass.minorUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxScaling] = function(oClass, value) {
        oClass.scaling = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxTickLblPos] = function(oClass, value) {
        oClass.tickLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxTitle] = function(oClass, value) {
        oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DateAxTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetAxId] = function(oClass, value) {
        oClass.internalSetAxId(value);
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetAxPos] = function(oClass, value) {
        oClass.axPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetCrossAx] = function(oClass, value) {
        oClass.crossAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetCrosses] = function(oClass, value) {
        oClass.crosses = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetCrossesAt] = function(oClass, value) {
        oClass.crossesAt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetMajorGridlines] = function(oClass, value) {
        oClass.majorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetMajorTickMark] = function(oClass, value) {
        oClass.majorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetMinorGridlines] = function(oClass, value) {
        oClass.minorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetMinorTickMark] = function(oClass, value) {
        oClass.minorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetScaling] = function(oClass, value) {
        oClass.scaling = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetTickLblPos] = function(oClass, value) {
        oClass.tickLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetTickLblSkip] = function(oClass, value) {
        oClass.tickLblSkip = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetTickMarkSkip] = function(oClass, value) {
        oClass.tickMarkSkip = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetTitle] = function(oClass, value) {
        oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SerAxSetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetAxId] = function(oClass, value) {
        oClass.internalSetAxId(value);
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetAxPos] = function(oClass, value) {
        oClass.axPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetCrossAx] = function(oClass, value) {
        oClass.crossAx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetCrossBetween] = function(oClass, value) {
        oClass.crossBetween = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetCrosses] = function(oClass, value) {
        oClass.crosses = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetCrossesAt] = function(oClass, value) {
        oClass.crossesAt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetDispUnits] = function(oClass, value) {
        oClass.dispUnits = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMajorGridlines] = function(oClass, value) {
        oClass.majorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMajorTickMark] = function(oClass, value) {
        oClass.majorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMajorUnit] = function(oClass, value) {
        oClass.majorUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMinorGridlines] = function(oClass, value) {
        oClass.minorGridlines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMinorTickMark] = function(oClass, value) {
        oClass.minorTickMark = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetMinorUnit] = function(oClass, value) {
        oClass.minorUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetScaling] = function(oClass, value) {
        oClass.scaling = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetTickLblPos] = function(oClass, value) {
        oClass.tickLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetTitle] = function(oClass, value) {
        oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ValAxSetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BandFmt_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BandFmt_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetErrBars] = function(oClass, value) {
        oClass.errBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetInvertIfNegative] = function(oClass, value) {
        oClass.invertIfNegative = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetPictureOptions] = function(oClass, value) {
        oClass.pictureOptions = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetShape] = function(oClass, value) {
        oClass.shape = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BarSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetBubble3D] = function(oClass, value) {
        oClass.bubble3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetBubbleScale] = function(oClass, value) {
        oClass.bubbleScale = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetShowNegBubbles] = function(oClass, value) {
        oClass.showNegBubbles = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetSizeRepresents] = function(oClass, value) {
        oClass.sizeRepresents = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetBubble3D] = function(oClass, value) {
        oClass.bubble3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetBubbleSize] = function(oClass, value) {
        oClass.bubbleSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetErrBars] = function(oClass, value) {
        oClass.errBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetInvertIfNegative] = function(oClass, value) {
        oClass.invertIfNegative = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetXVal] = function(oClass, value) {
        oClass.xVal = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_BubbleSeries_SetYVal] = function(oClass, value) {
        oClass.yVal = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Cat_SetMultiLvlStrRef] = function(oClass, value) {
        oClass.multiLvlStrRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Cat_SetNumLit] = function(oClass, value) {
        oClass.numLit = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Cat_SetNumRef] = function(oClass, value) {
        oClass.numRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Cat_SetStrLit] = function(oClass, value) {
        oClass.strLit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Cat_SetStrRef] = function(oClass, value) {
        oClass.strRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_ChartText_SetRich] = function(oClass, value) {
        oClass.rich = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartText_SetStrRef] = function(oClass, value) {
        oClass.strRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetDLblPos] = function(oClass, value) {
        oClass.dLblPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetLeaderLines] = function(oClass, value) {
        oClass.leaderLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetNumFmt] = function(oClass, value) {
        oClass.numFmt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetSeparator] = function(oClass, value) {
        oClass.separator = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowBubbleSize] = function(oClass, value) {
        oClass.showBubbleSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowCatName] = function(oClass, value) {
        oClass.showCatName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowLeaderLines] = function(oClass, value) {
        oClass.showLeaderLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowLegendKey] = function(oClass, value) {
        oClass.showLegendKey = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowPercent] = function(oClass, value) {
        oClass.showPercent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowSerName] = function(oClass, value) {
        oClass.showSerName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetShowVal] = function(oClass, value) {
        oClass.showVal = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DLbls_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetBubble3D] = function(oClass, value) {
        oClass.bubble3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetExplosion] = function(oClass, value) {
        oClass.explosion = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetInvertIfNegative] = function(oClass, value) {
        oClass.invertIfNegative = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetPictureOptions] = function(oClass, value) {
        oClass.pictureOptions = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DPt_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetShowHorzBorder] = function(oClass, value) {
        oClass.showHorzBorder = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetShowKeys] = function(oClass, value) {
        oClass.showKeys = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetShowOutline] = function(oClass, value) {
        oClass.showOutline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetShowVertBorder] = function(oClass, value) {
        oClass.showVertBorder = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DTable_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DispUnitsSetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DispUnitsSetBuiltInUnit] = function(oClass, value) {
        oClass.builtInUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DispUnitsSetCustUnit] = function(oClass, value) {
        oClass.custUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DispUnitsSetDispUnitsLbl] = function(oClass, value) {
        oClass.dispUnitsLbl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DoughnutChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DoughnutChart_SetFirstSliceAng] = function(oClass, value) {
        oClass.firstSliceAng = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DoughnutChart_SetHoleSize] = function(oClass, value) {
        oClass.holeSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DoughnutChart_SetVaryColor] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetErrBarType] = function(oClass, value) {
        oClass.errBarType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetErrDir] = function(oClass, value) {
        oClass.errDir = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetErrValType] = function(oClass, value) {
        oClass.errValType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetMinus] = function(oClass, value) {
        oClass.minus = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetNoEndCap] = function(oClass, value) {
        oClass.noEndCap = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetPlus] = function(oClass, value) {
        oClass.plus = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
        oClass.pen = null;
    };
    drawingsChangesMap[AscDFH.historyitem_ErrBars_SetVal] = function(oClass, value) {
        oClass.val = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetH] = function(oClass, value) {
        oClass.h = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetHMode] = function(oClass, value) {
        oClass.hMode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetLayoutTarget] = function(oClass, value) {
        oClass.layoutTarget = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetW] = function(oClass, value) {
        oClass.w = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetWMode] = function(oClass, value) {
        oClass.wMode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetX] = function(oClass, value) {
        oClass.x = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetXMode] = function(oClass, value) {
        oClass.xMode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetY] = function(oClass, value) {
        oClass.y = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetYMode] = function(oClass, value) {
        oClass.yMode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Layout_SetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Legend_SetLayout] = function(oClass, value) {
        oClass.layout = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Legend_SetLegendPos] = function(oClass, value) {
        oClass.legendPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Legend_SetOverlay] = function(oClass, value) {
        oClass.overlay = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Legend_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Legend_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LegendEntry_SetDelete] = function(oClass, value) {
        oClass.bDelete = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LegendEntry_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LegendEntry_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetDropLines] = function(oClass, value) {
        oClass.dropLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetGrouping] = function(oClass, value) {
        oClass.grouping = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetHiLowLines] = function(oClass, value) {
        oClass.hiLowLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetSmooth] = function(oClass, value) {
        oClass.smooth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetUpDownBars] = function(oClass, value) {
        oClass.upDownBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetErrBars] = function(oClass, value) {
        oClass.errBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetSmooth] = function(oClass, value) {
        oClass.smooth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LineSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Marker_SetSize] = function(oClass, value) {
        oClass.size = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Marker_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Marker_SetSymbol] = function(oClass, value) {
        oClass.symbol = value;
    };
    drawingsChangesMap[AscDFH.historyitem_MinusPlus_SetNumLit] = function(oClass, value) {
        oClass.numLit = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_MinusPlus_SetNumRef] = function(oClass, value) {
        oClass.numRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_MultiLvlStrRef_SetF] = function(oClass, value) {
        oClass.internalSetFormula(value);
    };
    drawingsChangesMap[AscDFH.historyitem_MultiLvlStrRef_SetMultiLvlStrCache] = function(oClass, value) {
        oClass.multiLvlStrCache = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumRef_SetF] = function(oClass, value) {
        oClass.internalSetFormula(value);
    };
    drawingsChangesMap[AscDFH.historyitem_NumRef_SetNumCache] = function(oClass, value) {
        oClass.numCache = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumericPoint_SetFormatCode] = function(oClass, value) {
        oClass.formatCode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumericPoint_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumericPoint_SetVal] = function(oClass, value) {
        oClass.val = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumFmt_SetFormatCode] = function(oClass, value) {
        oClass.formatCode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumFmt_SetSourceLinked] = function(oClass, value) {
        oClass.sourceLinked = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumLit_SetFormatCode] = function(oClass, value) {
        oClass.formatCode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_NumLit_SetPtCount] = function(oClass, value) {
        oClass.ptCount = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetGapWidth] = function(oClass, value) {
        oClass.gapWidth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetOfPieType] = function(oClass, value) {
        oClass.ofPieType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetSecondPieSize] = function(oClass, value) {
        oClass.secondPieSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetSerLines] = function(oClass, value) {
        oClass.serLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetSplitPos] = function(oClass, value) {
        oClass.splitPos = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetSplitType] = function(oClass, value) {
        oClass.splitType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_OfPieChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PictureOptions_SetApplyToEnd] = function(oClass, value) {
        oClass.applyToEnd = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PictureOptions_SetApplyToFront] = function(oClass, value) {
        oClass.applyToFront = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PictureOptions_SetApplyToSides] = function(oClass, value) {
        oClass.applyToSides = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PictureOptions_SetPictureFormat] = function(oClass, value) {
        oClass.pictureFormat = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PictureOptions_SetPictureStackUnit] = function(oClass, value) {
        oClass.pictureStackUnit = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieChart_SetFirstSliceAng] = function(oClass, value) {
        oClass.firstSliceAng = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieChart_3D] = function(oClass, value) {
        oClass.b3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetExplosion] = function(oClass, value) {
        oClass.explosion = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PieSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_PivotFmt_SetDLbl] = function(oClass, value) {
        oClass.dLbl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotFmt_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotFmt_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotFmt_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotFmt_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarChart_SetRadarStyle] = function(oClass, value) {
        oClass.radarStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RadarSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Scaling_SetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scaling_SetLogBase] = function(oClass, value) {
        oClass.logBase = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scaling_SetMax] = function(oClass, value) {
        oClass.max = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scaling_SetMin] = function(oClass, value) {
        oClass.min = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scaling_SetOrientation] = function(oClass, value) {
        oClass.orientation = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterChart_SetScatterStyle] = function(oClass, value) {
        oClass.scatterStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterChart_SetVaryColors] = function(oClass, value) {
        oClass.varyColors = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetErrBars] = function(oClass, value) {
        oClass.errBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetMarker] = function(oClass, value) {
        oClass.marker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetSmooth] = function(oClass, value) {
        oClass.smooth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetXVal] = function(oClass, value) {
        oClass.xVal = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_ScatterSer_SetYVal] = function(oClass, value) {
        oClass.yVal = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Tx_SetStrRef] = function(oClass, value) {
        oClass.strRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Tx_SetVal] = function(oClass, value) {
        oClass.val = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StockChart_SetDLbls] = function(oClass, value) {
        oClass.dLbls = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StockChart_SetDropLines] = function(oClass, value) {
        oClass.dropLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StockChart_SetHiLowLines] = function(oClass, value) {
        oClass.hiLowLines = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StockChart_SetUpDownBars] = function(oClass, value) {
        oClass.upDownBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StrCache_SetPtCount] = function(oClass, value) {
        oClass.ptCount = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StrPoint_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StrPoint_SetVal] = function(oClass, value) {
        oClass.val = value;
        if(AscFonts.IsCheckSymbols) {
            if(typeof value === "string" && value.length > 0) {
                AscFonts.FontPickerByCharacter.getFontsByString(value);
            }
        }
    };
    drawingsChangesMap[AscDFH.historyitem_StrRef_SetF] = function(oClass, value) {
        oClass.internalSetFormula(value);
    };
    drawingsChangesMap[AscDFH.historyitem_StrRef_SetStrCache] = function(oClass, value) {
        oClass.strCache = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceChart_SetWireframe] = function(oClass, value) {
        oClass.wireframe = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceSeries_SetCat] = function(oClass, value) {
        oClass.cat = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SurfaceSeries_SetVal] = function(oClass, value) {
        oClass.val = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Title_SetLayout] = function(oClass, value) {
        oClass.layout = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Title_SetOverlay] = function(oClass, value) {
        oClass.overlay = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Title_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Title_SetTx] = function(oClass, value) {
        if(oClass.tx && oClass.tx.strRef || value && value.strRef) {
            oClass.onChangeDataRefs();
        }
        oClass.tx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Title_SetTxPr] = function(oClass, value) {
        oClass.txPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetBackward] = function(oClass, value) {
        oClass.backward = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetDispEq] = function(oClass, value) {
        oClass.dispEq = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetDispRSqr] = function(oClass, value) {
        oClass.dispRSqr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetForward] = function(oClass, value) {
        oClass.forward = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetIntercept] = function(oClass, value) {
        oClass.intercept = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetName] = function(oClass, value) {
        oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetOrder] = function(oClass, value) {
        oClass.order = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetPeriod] = function(oClass, value) {
        oClass.period = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetTrendlineLbl] = function(oClass, value) {
        oClass.trendlineLbl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Trendline_SetTrendlineType] = function(oClass, value) {
        oClass.trendlineType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_UpDownBars_SetDownBars] = function(oClass, value) {
        oClass.downBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_UpDownBars_SetGapWidth] = function(oClass, value) {
        oClass.gapWidth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_UpDownBars_SetUpBars] = function(oClass, value) {
        oClass.upBars = value;
    };
    drawingsChangesMap[AscDFH.historyitem_YVal_SetNumLit] = function(oClass, value) {
        oClass.numLit = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_YVal_SetNumRef] = function(oClass, value) {
        oClass.numRef = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetAutoTitleDeleted] = function(oClass, value) {
        oClass.autoTitleDeleted = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetBackWall] = function(oClass, value) {
        oClass.backWall = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetDispBlanksAs] = function(oClass, value) {
        oClass.dispBlanksAs = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetFloor] = function(oClass, value) {
        oClass.floor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetLegend] = function(oClass, value) {
        oClass.legend = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetPlotArea] = function(oClass, value) {
        oClass.plotArea = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetPlotVisOnly] = function(oClass, value) {
        oClass.plotVisOnly = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetShowDLblsOverMax] = function(oClass, value) {
        oClass.showDLblsOverMax = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetSideWall] = function(oClass, value) {
        oClass.sideWall = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetTitle] = function(oClass, value) {
        oClass.title = value;
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_Chart_SetView3D] = function(oClass, value) {
        oClass.view3D = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_ChartWall_SetPictureOptions] = function(oClass, value) {
        oClass.pictureOptions = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartWall_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartWall_SetThickness] = function(oClass, value) {
        oClass.thickness = value;
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetDepthPercent] = function(oClass, value) {
        oClass.depthPercent = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetHPercent] = function(oClass, value) {
        oClass.hPercent = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetPerspective] = function(oClass, value) {
        oClass.perspective = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetRAngAx] = function(oClass, value) {
        oClass.rAngAx = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetRotX] = function(oClass, value) {
        oClass.rotX = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_View3d_SetRotY] = function(oClass, value) {
        oClass.rotY = value;
        oClass.Refresh_RecalcData();
    };
    drawingsChangesMap[AscDFH.historyitem_ExternalData_SetAutoUpdate] = function(oClass, value) {
        oClass.autoUpdate = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ExternalData_SetId] = function(oClass, value) {
        oClass.id = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotSource_SetFmtId] = function(oClass, value) {
        oClass.fmtId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PivotSource_SetName] = function(oClass, value) {
        oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Protection_SetChartObject] = function(oClass, value) {
        oClass.chartObject = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Protection_SetData] = function(oClass, value) {
        oClass.data = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Protection_SetFormatting] = function(oClass, value) {
        oClass.formatting = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Protection_SetSelection] = function(oClass, value) {
        oClass.selection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Protection_SetUserInterface] = function(oClass, value) {
        oClass.userInterface = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrintSettingsSetHeaderFooter] = function(oClass, value) {
        oClass.headerFooter = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrintSettingsSetPageMargins] = function(oClass, value) {
        oClass.pageMargins = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrintSettingsSetPageSetup] = function(oClass, value) {
        oClass.pageSetup = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetAlignWithMargins] = function(oClass, value) {
        oClass.alignWithMargins = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetDifferentFirst] = function(oClass, value) {
        oClass.differentFirst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetDifferentOddEven] = function(oClass, value) {
        oClass.differentOddEven = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetEvenFooter] = function(oClass, value) {
        oClass.evenFooter = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetEvenHeader] = function(oClass, value) {
        oClass.evenHeader = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetFirstFooter] = function(oClass, value) {
        oClass.firstFooter = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetFirstHeader] = function(oClass, value) {
        oClass.firstHeader = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetOddFooter] = function(oClass, value) {
        oClass.oddFooter = value;
    };
    drawingsChangesMap[AscDFH.historyitem_HeaderFooterChartSetOddHeader] = function(oClass, value) {
        oClass.oddHeader = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetB] = function(oClass, value) {
        oClass.b = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetFooter] = function(oClass, value) {
        oClass.footer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetHeader] = function(oClass, value) {
        oClass.header = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetL] = function(oClass, value) {
        oClass.l = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetR] = function(oClass, value) {
        oClass.r = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageMarginsSetT] = function(oClass, value) {
        oClass.t = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetBlackAndWhite] = function(oClass, value) {
        oClass.blackAndWhite = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetCopies] = function(oClass, value) {
        oClass.copies = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetDraft] = function(oClass, value) {
        oClass.draft = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetFirstPageNumber] = function(oClass, value) {
        oClass.firstPageNumber = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetHorizontalDpi] = function(oClass, value) {
        oClass.horizontalDpi = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetOrientation] = function(oClass, value) {
        oClass.orientation = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetPaperHeight] = function(oClass, value) {
        oClass.paperHeight = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetPaperSize] = function(oClass, value) {
        oClass.paperSize = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetPaperWidth] = function(oClass, value) {
        oClass.paperWidth = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetUseFirstPageNumb] = function(oClass, value) {
        oClass.useFirstPageNumb = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PageSetupSetVerticalDpi] = function(oClass, value) {
        oClass.verticalDpi = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CommonSeries_SetIdx] = function(oClass, value) {
        oClass.idx = value;
        oClass.Refresh_RecalcData({Type: AscDFH.historyitem_CommonSeries_SetIdx});
    };
    drawingsChangesMap[AscDFH.historyitem_CommonSeries_SetOrder] = function(oClass, value) {
        oClass.order = value;
        oClass.Refresh_RecalcData({Type: AscDFH.historyitem_CommonSeries_SetOrder});
    };
    drawingsChangesMap[AscDFH.historyitem_CommonSeries_SetTx] = function(oClass, value) {
        oClass.tx = value;
        oClass.Refresh_RecalcData({Type: AscDFH.historyitem_CommonSeries_SetTx});
        oClass.onChangeDataRefs();
    };
    drawingsChangesMap[AscDFH.historyitem_CommonSeries_SetSpPr] = function(oClass, value) {
        oClass.spPr = value;
        oClass.Refresh_RecalcData({Type: AscDFH.historyitem_CommonSeries_SetSpPr});
    };
    drawingsChangesMap[AscDFH.historyitem_CommonChart_DataLabelsRange] = function(oClass, value) {
        oClass.dataLablesRange = value;
    };

    drawingsChangesMap[AscDFH.historyitem_ChartStyleAxisTitle] = function(oClass, value) {
        oClass.axisTitle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleCategoryAxis] = function(oClass, value) {
        oClass.categoryAxis = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleChartArea] = function(oClass, value) {
        oClass.chartArea = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataLabel] = function(oClass, value) {
        oClass.dataLabel = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataLabelCallout] = function(oClass, value) {
        oClass.dataLabelCallout = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataPoint] = function(oClass, value) {
        oClass.dataPoint = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataPoint3D] = function(oClass, value) {
        oClass.dataPoint3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataPointLine] = function(oClass, value) {
        oClass.dataPointLine = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataPointMarker] = function(oClass, value) {
        oClass.dataPointMarker = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataPointWireframe] = function(oClass, value) {
        oClass.dataPointWireframe = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDataTable] = function(oClass, value) {
        oClass.dataTable = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDownBar] = function(oClass, value) {
        oClass.downBar = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleDropLine] = function(oClass, value) {
        oClass.dropLine = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleErrorBar] = function(oClass, value) {
        oClass.errorBar = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleFloor] = function(oClass, value) {
        oClass.floor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleGridlineMajor] = function(oClass, value) {
        oClass.gridlineMajor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleGridlineMinor] = function(oClass, value) {
        oClass.gridlineMinor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleHiLoLine] = function(oClass, value) {
        oClass.hiLoLine = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleLeaderLine] = function(oClass, value) {
        oClass.leaderLine = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleLegend] = function(oClass, value) {
        oClass.legend = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStylePlotArea] = function(oClass, value) {
        oClass.plotArea = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStylePlotArea3D] = function(oClass, value) {
        oClass.plotArea3D = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleSeriesAxis] = function(oClass, value) {
        oClass.seriesAxis = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleSeriesLine] = function(oClass, value) {
        oClass.seriesLine = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleTitle] = function(oClass, value) {
        oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleTrendline] = function(oClass, value) {
        oClass.trendline = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleTrendlineLabel] = function(oClass, value) {
        oClass.trendlineLabel = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleUpBar] = function(oClass, value) {
        oClass.upBar = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleValueAxis] = function(oClass, value) {
        oClass.valueAxis = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleWall] = function(oClass, value) {
        oClass.wall = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleMarkerLayout] = function(oClass, value) {
        oClass.markerLayout = value;
    };

    drawingsChangesMap[AscDFH.historyitem_ChartStyleMarkerId] = function(oClass, value) {
        oClass.id = value;
    };

    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryType] = function(oClass, value) {
        oClass.type = value;
    };

    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryLineWidthScale] = function(oClass, value) {
        oClass.lineWidthScale = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryLnRef] = function(oClass, value) {
        oClass.lnRef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryFillRef] = function(oClass, value) {
        oClass.fillRef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryEffectRef] = function(oClass, value) {
        oClass.effectRef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryFontRef] = function(oClass, value) {
        oClass.fontRef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryDefRPr] = function(oClass, value) {
        oClass.defRPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntryBodyPr] = function(oClass, value) {
        oClass.bodyPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChartStyleEntrySpPr] = function(oClass, value) {
        oClass.spPr = value;
    };

    drawingsChangesMap[AscDFH.historyitem_MarkerLayoutSymbol] = function(oClass, value) {
        oClass.symbol = value;
    };

    drawingsChangesMap[AscDFH.historyitem_MarkerLayoutSize] = function(oClass, value) {
        oClass.size = value;
    };


    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowBubbleSize] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowCatName] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowLegendKey] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowPercent] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowSerName] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowVal] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetShowDLblsRange] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_Set3D] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_AreaChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetAuto] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetNoMultiLvlLbl] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxAuto] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetInvertIfNegative] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetBubble3D] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetShowNegBubbles] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetBubble3D] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetInvertIfNegative] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowBubbleSize] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowCatName] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowLeaderLines] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowLegendKey] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowPercent] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowSerName] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetShowVal] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetBubble3D] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetInvertIfNegative] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetShowHorzBorder] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetShowKeys] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetShowOutline] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetShowVertBorder] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_DoughnutChart_SetVaryColor] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetNoEndCap] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_SetOverlay] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_LegendEntry_SetDelete] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetMarker] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetSmooth] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetSmooth] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_NumFmt_SetSourceLinked] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PieChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PieChart_3D] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_RadarChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterChart_SetVaryColors] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetSmooth] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceChart_SetWireframe] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Title_SetOverlay] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetDispEq] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetDispRSqr] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetAutoTitleDeleted] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetPlotVisOnly] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetShowDLblsOverMax] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetRAngAx] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_ExternalData_SetAutoUpdate] = window['AscDFH'].CChangesDrawingsBool;

    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetDLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetGapDepth] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetShape] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetBarDir] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetGapWidth] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetGrouping] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetOverlap] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AreaChart_SetGrouping] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetAxId] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetAxPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetCrosses] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetLblAlgn] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetLblOffset] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetMajorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetMinorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetTickLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetTickLblSkip] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetTickMarkSkip] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxAxPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxCrosses] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxLblOffset] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMajorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMinorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxTickLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetAxId] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetAxPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetMajorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetMinorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetTickLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetTickLblSkip] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetTickMarkSkip] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetAxId] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetAxPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetCrossBetween] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMajorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMinorTickMark] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetTickLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BandFmt_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetBubbleScale] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetSizeRepresents] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetDLblPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetExplosion] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DispUnitsSetParent] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DoughnutChart_SetFirstSliceAng] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_DoughnutChart_SetHoleSize] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetErrBarType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetErrDir] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetErrValType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetHMode] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetLayoutTarget] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetWMode] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetXMode] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetYMode] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_SetLegendPos] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_LegendEntry_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetGrouping] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Marker_SetSize] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Marker_SetSymbol] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_MultiLvlStrCache_SetPtCount] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_NumLit_SetPtCount] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetGapWidth] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetOfPieType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetSecondPieSize] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetSplitType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PictureOptions_SetPictureFormat] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PieChart_SetFirstSliceAng] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetExplosion] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PivotFmt_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Scaling_SetOrientation] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterChart_SetScatterStyle] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_StrCache_SetPtCount] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_StringLiteral_SetPtCount] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_StrPoint_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetPeriod] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetTrendlineType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_UpDownBars_SetGapWidth] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetDispBlanksAs] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ChartWall_SetThickness] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetDepthPercent] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetHPercent] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetPerspective] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetRotX] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_View3d_SetRotY] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_NumericPoint_SetIdx] = window['AscDFH'].CChangesDrawingsLong;

    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetCrossesAt] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxBaseTimeUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxCrossesAt] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMajorTimeUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMajorUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMinorTimeUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMinorUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetCrosses] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetCrossesAt] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetCrosses] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetCrossesAt] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMajorUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMinorUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_DispUnitsSetBuiltInUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetVal] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetH] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetW] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetX] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetY] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Layout_SetParent] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetSplitPos] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PictureOptions_SetPictureStackUnit] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Scaling_SetLogBase] = window['AscDFH'].CChangesDrawingsDouble2;
    AscDFH.changesFactory[AscDFH.historyitem_Scaling_SetMax] = window['AscDFH'].CChangesDrawingsDouble2;
    AscDFH.changesFactory[AscDFH.historyitem_Scaling_SetMin] = window['AscDFH'].CChangesDrawingsDouble2;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetBackward] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetForward] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetIntercept] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_NumericPoint_SetVal] = window['AscDFH'].CChangesDrawingsDouble2;

    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetSeparator] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxAxId] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetSeparator] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_MultiLvlStrRef_SetF] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_NumRef_SetF] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_NumericPoint_SetFormatCode] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_NumFmt_SetFormatCode] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_NumLit_SetFormatCode] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_Tx_SetVal] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_StrPoint_SetVal] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_StrRef_SetF] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetName] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_ExternalData_SetId] = window['AscDFH'].CChangesDrawingsString;

    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetLayout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetTx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbl_SetParent] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_SetDTable] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_SetLayout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChartFormat_SetParent] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_SetDlbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_SetSerLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaChart_SetDropLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetErrBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetPictureOptions] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetCrossAx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetMajorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetMinorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetScaling] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CatAxSetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxCrossAx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMajorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxMajorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxScaling] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DateAxTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetCrossAx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetMajorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetMinorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetScaling] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SerAxSetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetCrossAx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetDispUnits] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMajorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetMinorGridlines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetScaling] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ValAxSetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BandFmt_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetErrBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetPictureOptions] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetShape] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetBubbleSize] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetErrBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetXVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleSeries_SetYVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Cat_SetMultiLvlStrRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Cat_SetNumLit] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Cat_SetNumRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Cat_SetStrLit] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Cat_SetStrRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartFormatSetChart] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartText_SetRich] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartText_SetStrRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetLeaderLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetNumFmt] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetMarker] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetPictureOptions] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DPt_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DTable_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DispUnitsSetCustUnit] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DispUnitsSetDispUnitsLbl] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_DoughnutChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetMinus] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetPlus] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ErrBars_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_SetLayout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LegendEntry_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetDropLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetHiLowLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_SetUpDownBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetErrBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetMarker] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Marker_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_MinusPlus_SetNumLit] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_MinusPlus_SetNumRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_MultiLvlStrRef_SetMultiLvlStrCache] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_NumRef_SetNumCache] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_SetSerLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PictureOptions_SetApplyToEnd] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PictureOptions_SetApplyToFront] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PictureOptions_SetApplyToSides] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PieChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PivotFmt_SetDLbl] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PivotFmt_SetMarker] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PivotFmt_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PivotFmt_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarChart_SetRadarStyle] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Scaling_SetParent] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetErrBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetMarker] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetXVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterSer_SetYVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Tx_SetStrRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_StockChart_SetDLbls] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_StockChart_SetDropLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_StockChart_SetHiLowLines] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_StockChart_SetUpDownBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_StrRef_SetStrCache] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceSeries_SetCat] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceSeries_SetVal] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Title_SetLayout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Title_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Title_SetTx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Title_SetTxPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Trendline_SetTrendlineLbl] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_UpDownBars_SetDownBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_UpDownBars_SetUpBars] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_YVal_SetNumLit] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_YVal_SetNumRef] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetBackWall] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetFloor] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetLegend] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetPlotArea] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetSideWall] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_SetView3D] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartWall_SetPictureOptions] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartWall_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;

    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_AddAxis] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_AddChart] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_RemoveChart] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_PlotArea_RemoveAxis] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_RemoveSeries] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_AddSeries] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_AddFilteredSeries] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_RemoveFilteredSeries] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_BarChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_AreaChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_AreaSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonSeries_RemoveDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_BarSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_BubbleChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_DLbls_SetDLbl] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_Legend_AddLegendEntry] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_LineChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_LineSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_MultiLvlStrCache_SetLvl] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonLit_RemoveDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_NumLit_AddPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_OfPieChart_AddCustSplit] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_PieSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_RadarChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_RadarSeries_SetDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_ScatterChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_StockChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonLit_RemoveDPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_StrCache_AddPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_StringLiteral_SetPt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceChart_AddAxId] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_SurfaceChart_AddBandFmt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_Chart_AddPivotFmt] = window['AscDFH'].CChangesDrawingsContent;
    AscDFH.changesFactory[AscDFH.historyitem_CommonSeries_SetIdx] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CommonSeries_SetOrder] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_CommonSeries_SetTx] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CommonSeries_SetSpPr] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_CommonChart_DataLabelsRange] = window['AscDFH'].CChangesDrawingsObject;

    AscDFH.changesFactory[AscDFH.historyitem_PivotSource_SetFmtId] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PivotSource_SetName] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_Protection_SetChartObject] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Protection_SetData] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Protection_SetFormatting] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Protection_SetSelection] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_Protection_SetUserInterface] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PrintSettingsSetHeaderFooter] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PrintSettingsSetPageMargins] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_PrintSettingsSetPageSetup] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetAlignWithMargins] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetDifferentFirst] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetDifferentOddEven] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetEvenFooter] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetEvenHeader] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetFirstFooter] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetFirstHeader] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetOddFooter] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_HeaderFooterChartSetOddHeader] = window['AscDFH'].CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetB] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetFooter] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetHeader] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetL] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetR] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageMarginsSetT] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetBlackAndWhite] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetCopies] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetDraft] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetFirstPageNumber] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetHorizontalDpi] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetOrientation] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetPaperHeight] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetPaperSize] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetPaperWidth] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetUseFirstPageNumb] = window['AscDFH'].CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PageSetupSetVerticalDpi] = window['AscDFH'].CChangesDrawingsLong;

    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleAxisTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleCategoryAxis] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleChartArea] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataLabel] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataLabelCallout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataPoint] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataPoint3D] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataPointLine] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataPointMarker] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataPointWireframe] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDataTable] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDownBar] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleDropLine] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleErrorBar] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleFloor] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleGridlineMajor] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleGridlineMinor] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleHiLoLine] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleLeaderLine] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleLegend] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStylePlotArea] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStylePlotArea3D] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleSeriesAxis] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleSeriesLine] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleTitle] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleTrendline] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleTrendlineLabel] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleUpBar] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleValueAxis] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleWall] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleMarkerLayout] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleMarkerId] = window['AscDFH'].CChangesDrawingsLong;

    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryType] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryLineWidthScale] = window['AscDFH'].CChangesDrawingsDouble2;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryLnRef] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryFillRef] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryEffectRef] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryFontRef] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryDefRPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntryBodyPr] = window['AscDFH'].CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ChartStyleEntrySpPr] = window['AscDFH'].CChangesDrawingsObject;

    AscDFH.changesFactory[AscDFH.historyitem_MarkerLayoutSymbol] = window['AscDFH'].CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_MarkerLayoutSize] = window['AscDFH'].CChangesDrawingsLong;

    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryLnRef] = AscFormat.StyleRef;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryFillRef] = AscFormat.StyleRef;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryEffectRef] = AscFormat.StyleRef;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryFontRef] = AscFormat.FontRef;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartStyleEntryBodyPr] = AscFormat.CBodyPr;


    drawingContentChanges[AscDFH.historyitem_PlotArea_AddAxis] =
        drawingContentChanges[AscDFH.historyitem_BarChart_AddAxId] =
            drawingContentChanges[AscDFH.historyitem_AreaChart_AddAxId] =
                drawingContentChanges[AscDFH.historyitem_CommonChart_AddAxId] =
                    drawingContentChanges[AscDFH.historyitem_BubbleChart_AddAxId] =
                        drawingContentChanges[AscDFH.historyitem_LineChart_AddAxId] =
                            drawingContentChanges[AscDFH.historyitem_RadarChart_AddAxId] =
                                drawingContentChanges[AscDFH.historyitem_ScatterChart_AddAxId] =
                                    drawingContentChanges[AscDFH.historyitem_StockChart_AddAxId] =
                                        drawingContentChanges[AscDFH.historyitem_SurfaceChart_AddAxId] =
                                            drawingContentChanges[AscDFH.historyitem_PlotArea_RemoveAxis] = function(oClass) {
                                                return oClass.axId;
                                            };

    drawingContentChanges[AscDFH.historyitem_PlotArea_AddChart] =
        drawingContentChanges[AscDFH.historyitem_PlotArea_RemoveChart] = function(oClass) {
            oClass.onChangeDataRefs();
            return oClass.charts;
        };

    drawingContentChanges[AscDFH.historyitem_CommonChart_RemoveSeries] =
        drawingContentChanges[AscDFH.historyitem_CommonChart_AddSeries] = function(oClass) {
            oClass.onChangeDataRefs();
            return oClass.series;
        };

    drawingContentChanges[AscDFH.historyitem_CommonChart_AddFilteredSeries] =
        drawingContentChanges[AscDFH.historyitem_CommonChart_RemoveFilteredSeries] = function(oClass) {
            return oClass.filteredSeries;
        };

    drawingContentChanges[AscDFH.historyitem_AreaSeries_SetDPt] =
        drawingContentChanges[AscDFH.historyitem_CommonSeries_RemoveDPt] =
            drawingContentChanges[AscDFH.historyitem_BarSeries_SetDPt] =
                drawingContentChanges[AscDFH.historyitem_LineSeries_SetDPt] =
                    drawingContentChanges[AscDFH.historyitem_PieSeries_SetDPt] =
                        drawingContentChanges[AscDFH.historyitem_ScatterSer_SetDPt] =
                            drawingContentChanges[AscDFH.historyitem_BubbleSeries_SetDPt] =
                                drawingContentChanges[AscDFH.historyitem_RadarSeries_SetDPt] =
                                    drawingContentChanges[AscDFH.historyitem_PieSeries_SetDPt] =
                                        drawingContentChanges[AscDFH.historyitem_LineSeries_SetDPt] =
                                            drawingContentChanges[AscDFH.historyitem_RadarSeries_SetDPt] = function(oClass) {
                                                return oClass.dPt;
                                            };

    drawingContentChanges[AscDFH.historyitem_DLbls_SetDLbl] = function(oClass) {
        oClass.onChangeDataRefs();
        return oClass.dLbl;
    };

    drawingContentChanges[AscDFH.historyitem_Legend_AddLegendEntry] = function(oClass) {
        return oClass.legendEntryes;
    };

    drawingContentChanges[AscDFH.historyitem_MultiLvlStrCache_SetLvl] = function(oClass) {
        return oClass.lvl;
    };

    drawingContentChanges[AscDFH.historyitem_CommonLit_RemoveDPt] =
        drawingContentChanges[AscDFH.historyitem_NumLit_AddPt] =
            drawingContentChanges[AscDFH.historyitem_StrCache_AddPt] =
                drawingContentChanges[AscDFH.historyitem_StringLiteral_SetPt] = function(oClass) {
                    return oClass.pts;
                };

    drawingContentChanges[AscDFH.historyitem_OfPieChart_AddCustSplit] = function(oClass) {
        return oClass.custSplit
    };

    drawingContentChanges[AscDFH.historyitem_SurfaceChart_AddBandFmt] = function(oClass) {
        return oClass.bandFmts;
    };

    drawingContentChanges[AscDFH.historyitem_Chart_AddPivotFmt] = function(oClass) {
        return oClass.pivotFmts;
    };


// Import
    var CMatrix = AscCommon.CMatrix;
    var isRealObject = AscCommon.isRealObject;
    var History = AscCommon.History;
    var global_MatrixTransformer = AscCommon.global_MatrixTransformer;

    var CShape = AscFormat.CShape;
    var checkSpPrRasterImages = AscFormat.checkSpPrRasterImages;

    var c_oAscChartLegendShowSettings = Asc.c_oAscChartLegendShowSettings;
    var c_oAscChartDataLabelsPos = Asc.c_oAscChartDataLabelsPos;
    var c_oAscValAxisRule = Asc.c_oAscValAxisRule;
    var c_oAscValAxUnits = Asc.c_oAscValAxUnits;
    var c_oAscTickMark = Asc.c_oAscTickMark;
    var c_oAscTickLabelsPos = Asc.c_oAscTickLabelsPos;
    var c_oAscCrossesRule = Asc.c_oAscCrossesRule;
    var c_oAscBetweenLabelsRule = Asc.c_oAscBetweenLabelsRule;
    var c_oAscLabelsPosition = Asc.c_oAscLabelsPosition;
    var c_oAscAxisType = Asc.c_oAscAxisType;

    var CChangesDrawingsBool = AscDFH.CChangesDrawingsBool;
    var CChangesDrawingsLong = AscDFH.CChangesDrawingsLong;
    var CChangesDrawingsDouble = AscDFH.CChangesDrawingsDouble;
    var CChangesDrawingsString = AscDFH.CChangesDrawingsString;
    var CChangesDrawingsObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
    var CChangesDrawingsObject = AscDFH.CChangesDrawingsObject;
    var CChangesDrawingsContent = AscDFH.CChangesDrawingsContent;
    var CChangesDrawingsDouble2 = AscDFH.CChangesDrawingsDouble2;

    var InitClass = AscFormat.InitClass;
    var CBaseFormatObject = AscFormat.CBaseFormatObject;

    function CBaseChartObject() {
        CBaseFormatObject.call(this);
    }

    InitClass(CBaseChartObject, CBaseFormatObject, AscDFH.historyitem_type_Unknown);
    CBaseChartObject.prototype.notAllowedWithoutId = function() {
        return true;
    };
    CBaseChartObject.prototype.getChartSpace = function() {
        var oCurElement = this;
        while(oCurElement) {
            if(oCurElement.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                return oCurElement;
            }
            oCurElement = oCurElement.parent;
        }
        return null;
    };
    CBaseChartObject.prototype.getDrawingDocument = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getDrawingDocument();
        }
        return null;
    };
    CBaseChartObject.prototype.onChartInternalUpdate = function(bColors) {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.handleUpdateInternalChart(bColors)
        }
    };
    CBaseChartObject.prototype.onChartUpdateType = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.handleUpdateType()
        }
    };
    CBaseChartObject.prototype.onChartUpdateDataLabels = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.handleUpdateDataLabels()
        }
    };
    CBaseChartObject.prototype.getTxPrParaPr = function() {
        if(this.txPr) {
            return this.txPr.getFirstParaParaPr();
        }
        return null;
    };
    CBaseChartObject.prototype.onChangeDataRefs = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.clearDataRefs();
        }
    };
    CBaseChartObject.prototype.getTheme = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getTheme();
        }
        return null;
    };
    CBaseChartObject.prototype.getSpPrFormStyleEntry = function(oStyleEntry, aColors, nIdx) {
        var oTheme = this.getTheme();
        var oSpPr = oStyleEntry.spPr;
        //fill
        var oFill;
        var oFillRef = oStyleEntry.fillRef;
        var oFillRefUnicolor = oFillRef.getNoStyleUnicolor(nIdx, aColors);
        oFill = oTheme.getFillStyle(oFillRef.idx, oFillRefUnicolor || aColors[nIdx]);
        if(oSpPr && oSpPr.Fill) {
            oFill = oSpPr.Fill.createDuplicate();
            var bIsSpecialStyle = oStyleEntry.isSpecialStyle();
            oFill.checkPhColor(oFillRefUnicolor || aColors[nIdx], bIsSpecialStyle);
            if(bIsSpecialStyle) {
                if(AscFormat.isRealNumber(nIdx)) {
                    var nPatternType = oStyleEntry.getSpecialPatternType(nIdx);
                    oFill.checkPatternType(nPatternType);
                }
            }
        }
        //line
        var oLn;
        var oLineRef = oStyleEntry.lnRef;
        var oLineRefUnicolor = oLineRef.getNoStyleUnicolor(nIdx, aColors);
        oLn = oTheme.getLnStyle(oLineRef.idx, oLineRefUnicolor);
        if(oSpPr && oSpPr.ln) {
            oLn = oSpPr.ln.createDuplicate();
            oLn.Fill.checkPhColor(oLineRefUnicolor, false);
        }
        if(AscFormat.isRealNumber(oLn.w) && AscFormat.isRealNumber(oStyleEntry.lineWidthScale)) {
            oLn.w *= oStyleEntry.lineWidthScale;
        }
        var oResultSpPr = new AscFormat.CSpPr();
        oResultSpPr.setFill(oFill);
        oResultSpPr.setLn(oLn);
        return oResultSpPr;
    };
    CBaseChartObject.prototype.getDocContentsWithImageBullets = function (arrContents) {};
    CBaseChartObject.prototype.getImageFromBulletsMap = function(oImages) {};
    CBaseChartObject.prototype.getTxPrFormStyleEntry = function(oStyleEntry, aColors, nIdx) {
        var oFontRef = oStyleEntry.fontRef;
        var oParaPr = new AscCommonWord.CParaPr();
        var oTextPr = new AscCommonWord.CTextPr();
        var oRFonts = oTextPr.RFonts;
        oRFonts.SetFontStyle(oFontRef.idx);
        var oFontUnicolor = oFontRef.getNoStyleUnicolor(nIdx, aColors);
        if(oFontUnicolor) {
            oTextPr.SetUnifill(AscFormat.CreateUniFillByUniColor(oFontUnicolor))
        }
        if(oStyleEntry.defRPr) {
            oTextPr.Merge(oStyleEntry.defRPr);
            if(oTextPr.Unifill) {
                oTextPr.Unifill.checkPhColor(oFontUnicolor, false);
            }
        }
        oParaPr.DefaultRunPr = oTextPr;
        var oTxPr;
        if(this.txPr && this.txPr.content && this.txPr.content.Content[0]) {
            oTxPr = this.txPr;
        }
        else {
            oTxPr = AscFormat.CreateTextBodyFromString("", this.getDrawingDocument(), this);
        }
        if(oStyleEntry.bodyPr) {
            oTxPr.setBodyPr(oStyleEntry.bodyPr.createDuplicate())
        }
        oTxPr.content.Content[0].Set_Pr(oParaPr);
        return oTxPr;
    };
    CBaseChartObject.prototype.applyStyleEntry = function(oStyleEntry, aColors, nIdx, bReset) {
        if(!this.setSpPr && !this.setTxPr || !oStyleEntry) {
            return;
        }
        if(this.setSpPr) {
            var oSpPr = this.getSpPrFormStyleEntry(oStyleEntry, aColors, nIdx);
            if(this.spPr) {
                if(bReset !== false || !this.spPr.Fill) {
                    this.spPr.setFill(oSpPr.Fill);
                }
                if(bReset !== false || !this.spPr.ln) {
                    this.spPr.setLn(oSpPr.ln);
                }
            }
            else {
                this.setSpPr(oSpPr);
            }
        }
        if(this.setTxPr) {
            this.setTxPr(this.getTxPrFormStyleEntry(oStyleEntry, aColors, nIdx));
            if(this.tx && this.tx.rich) {
                let oTxBody = this.tx.rich;
                if(oTxBody.content) {
                    let aParagraphs = oTxBody.content.Content;
                    for(let nPara = 0; nPara < aParagraphs.length; ++nPara) {
                        let oParagraph = aParagraphs[nPara];
                        oParagraph.Clear_TextFormatting(true);
                        if(oParagraph.Pr && oParagraph.Pr.DefaultRunPr) {
                            let oCopyPr = oParagraph.Pr.Copy();
                            oCopyPr.DefaultRunPr = undefined;
                            oParagraph.Set_Pr(oCopyPr);
                        }
                    }
                }
            }
        }
    };
    CBaseChartObject.prototype.resetFormatting = function() {
        this.resetOwnFormatting();
        var aChildren = this.getChildren();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            var oChild = aChildren[nChild];
            if(oChild && oChild.resetFormatting) {
                oChild.resetFormatting();
            }
        }
    };
    CBaseChartObject.prototype.resetOwnFormatting = function() {
        if(this.setTxPr && this.txPr) {
            this.setTxPr(null);
        }
        if(this.setSpPr && this.spPr) {
            if(!this.spPr.xfrm) {
                this.setSpPr(null);
            }
            else {
                if(this.spPr.Fill) {
                    this.spPr.setFill(null);
                }
                if(this.spPr.ln) {
                    this.spPr.setLn(null);
                }
            }
        }
        if(this.setSymbol && this.symbol !== SYMBOL_NONE && this.symbol !== null) {
            this.setSymbol(null);
        }
    };
    CBaseChartObject.prototype.isForm = function() {
        return false;
    };
    CBaseChartObject.prototype.GetParaDrawing = function() {
        return null;
    };
    CBaseChartObject.prototype.isObjectInSmartArt = function() {
        return false;
    };

    function getMinMaxFromArrPoints(aPoints) {
        if(Array.isArray(aPoints) && aPoints.length > 0) {
            if(isRealObject(aPoints[0]) && AscFormat.isRealNumber(aPoints[0].val) && isRealObject(aPoints[aPoints.length - 1]) && AscFormat.isRealNumber(aPoints[aPoints.length - 1].val)) {
                if(aPoints[0].val - aPoints[aPoints.length - 1].val <= 0) {
                    return {min: aPoints[0].val, max: aPoints[aPoints.length - 1].val};
                }
                else {
                    return {min: aPoints[aPoints.length - 1].val, max: aPoints[0].val};
                }
            }
        }
        return {min: null, max: null};
    }

    var SCALE_INSET_COEFF = 1.016;//Возможно придется уточнять
    function CDLbl() {
        CBaseChartObject.call(this);
        this.bDelete = null;
        this.dLblPos = null;
        this.idx = null;
        this.layout = null;
        this.numFmt = null;
        this.separator = null;
        this.showBubbleSize = null;
        this.showCatName = null;
        this.showLegendKey = null;
        this.showPercent = null;
        this.showSerName = null;
        this.showVal = null;
		this.showDlblsRange = null;
        this.spPr = null;
        this.tx = null;
        this.txPr = null;

        this.recalcInfo =
        {
            recalcTransform: true,
            recalculateTransformText: true,
            recalcStyle: true,
            recalculateTxBody: true,
            recalculateBrush: true,
            recalculatePen: true,
            recalculateContent: true
        };

        this.chart = null;
        this.series = null;

        this.x = 0;
        this.y = 0;
        this.calcX = null;
        this.calcY = null;

        this.relPosX = null;
        this.relPosY = null;
        this.txBody = null;
        this.transform = new CMatrix();
        this.transformText = new CMatrix();
        this.ownTransform = new CMatrix();
        this.ownTransformText = new CMatrix();
        this.localTransform = new CMatrix();
        this.localTransformText = new CMatrix();
        this.compiledStyles = null;
    }

    InitClass(CDLbl, CBaseChartObject, AscDFH.historyitem_type_DLbl);
    CDLbl.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
	CDLbl.prototype.notAllowedWithoutId = function() {
		return false;
	};
    CDLbl.prototype.Check_AutoFit = function() {
        return true;
    };
    CDLbl.prototype.getChildren = function() {
        return [this.layout, this.numFmt, this.spPr, this.tx, this.txPr];
    };
    CDLbl.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setDelete(this.bDelete);
        oCopy.setDLblPos(this.dLblPos);
        oCopy.setIdx(this.idx);
        if(this.layout) {
            oCopy.setLayout(this.layout.createDuplicate());
        }
        if(this.numFmt) {
            oCopy.setNumFmt(this.numFmt.createDuplicate());
        }
        oCopy.setSeparator(this.separator);
        oCopy.setShowBubbleSize(this.showBubbleSize);
        oCopy.setShowCatName(this.showCatName);
        oCopy.setShowLegendKey(this.showLegendKey);
        oCopy.setShowPercent(this.showPercent);
        oCopy.setShowSerName(this.showSerName);
        oCopy.setShowVal(this.showVal);
        oCopy.setShowDlblsRange(this.showDlblsRange);
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(this.tx) {
            oCopy.setTx(this.tx.createDuplicate());
        }
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CDLbl.prototype.checkShapeChildTransform = function(transform) {
        this.updatePosition(this.posX, this.posY);
        global_MatrixTransformer.MultiplyAppend(this.transform, transform);
        global_MatrixTransformer.MultiplyAppend(this.transformText, transform);
        if((this instanceof CTitle) || (this instanceof CDLbl)) {
            this.invertTransform = global_MatrixTransformer.Invert(this.transform);
            this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);
        }
    };
    CDLbl.prototype.getCompiledFill = function() {
        return this.spPr && this.spPr.Fill ? this.spPr.Fill : null;
    };
    CDLbl.prototype.getCompiledLine = function() {
        return this.spPr && this.spPr.ln ? this.spPr.ln : null;
    };
    CDLbl.prototype.getCompiledTransparent = function() {
        return this.spPr && this.spPr.Fill ? this.spPr.Fill.transparent : null;
    };
    CDLbl.prototype.recalculate = function() {
        if(this.bDelete) {
            return;
        }
        AscFormat.ExecuteNoHistory(function() {
            if(this.recalcInfo.recalculateBrush) {
                this.recalculateBrush();
                this.recalcInfo.recalculateBrush = false;
            }
            if(this.recalcInfo.recalculatePen) {
                this.recalculatePen();
                this.recalcInfo.recalculatePen = false;
            }
            if(this.recalcInfo.recalcStyle) {
                this.recalculateStyle();
                //this.recalcInfo.recalcStyle = false;
            }
            if(this.recalcInfo.recalculateTxBody) {
                this.recalculateTxBody();
                this.recalcInfo.recalculateTxBody = false;
            }
            if(this.recalcInfo.recalculateContent) {
                this.recalculateContent();
                //this.recalcInfo.recalculateContent = false;
            }
            if(this.recalcInfo.recalcTransform) {
                this.recalculateTransform();
                //this.recalcInfo.recalcTransform = false;
            }
            if(this.recalcInfo.recalculateTransformText) {
                this.recalculateTransformText();
                //this.recalcInfo.recalcTransformText = false;
            }
            if(this.chart) {
                this.chart.addToSetPosition(this);
            }
        }, this, []);
    };
    CDLbl.prototype.recalculateBrush = CShape.prototype.recalculateBrush;
    CDLbl.prototype.recalculatePen = CShape.prototype.recalculatePen;
    CDLbl.prototype.check_bounds = CShape.prototype.check_bounds;
    CDLbl.prototype.selectionCheck = CShape.prototype.selectionCheck;
    CDLbl.prototype.getInvertTransform = CShape.prototype.getInvertTransform;
    CDLbl.prototype.getDocContent = CShape.prototype.getDocContent;
    CDLbl.prototype.updateSelectionState = function() {
        if(this.txBody && this.txBody.content) {
            if(!this.txBody.content.DrawingDocument) {
                if(this.chart) {
                    var oDrawingDocument = this.chart.getDrawingDocument();
                    if(oDrawingDocument) {
                        this.txBody.content.DrawingDocument = oDrawingDocument;
                        var aContent = this.txBody.content.Content;
                        for(var i = 0; i < aContent.length; ++i) {
                            aContent[i].DrawingDocument = oDrawingDocument;
                        }
                    }
                }

            }
        }
        CShape.prototype.updateSelectionState.call(this);
    };
    CDLbl.prototype.selectionSetStart = CShape.prototype.selectionSetStart;
    CDLbl.prototype.selectionSetEnd = CShape.prototype.selectionSetEnd;
    CDLbl.prototype.getDrawingDocument = function() {
        return this.chart && this.chart.getDrawingDocument && this.chart.getDrawingDocument();
    };
    CDLbl.prototype.checkHitToBounds = function(x, y) {
        var oInvertTransform = this.getInvertTransform();
        var _x, _y;
        if(oInvertTransform) {
            _x = oInvertTransform.TransformPointX(x, y);
            _y = oInvertTransform.TransformPointY(x, y);
        }
        else {
            _x = x - this.transform.tx;
            _y = y - this.transform.ty;
        }

        return _x >= 0 && _x <= this.extX && _y >= 0 && _y < this.extY;
    };
    CDLbl.prototype.getCanvasContext = function() {
        return this.chart && this.chart.getCanvasContext();
    };
    CDLbl.prototype.convertPixToMM = function(pix) {
        return this.chart && this.chart.convertPixToMM(pix);
    };
    CDLbl.prototype.checkDlbl = function() {
        if(this.series && this.pt) {
            var oSeries = this.series;
            if(oSeries) {
                var oDlbls;
                if(!oSeries.dLbls) {
                    var oChart = oSeries.parent;
                    if(oChart && oChart.dLbls) {
                        oDlbls = oChart.dLbls.createDuplicate();
                    }
                    else {
                        oDlbls = new AscFormat.CDLbls();
                    }
                    oSeries.setDLbls(oDlbls);
                }
                else {
                    oDlbls = oSeries.dLbls;
                }
                var dLbl = oDlbls.findDLblByIdx(this.pt.idx);
                if(!dLbl) {
                    dLbl = this.createDuplicate();
                    dLbl.setDelete(undefined);
                    dLbl.setIdx(this.pt.idx);
                    dLbl.setParent(oSeries.dLbls);
                    oSeries.dLbls.addDLbl(dLbl);
                    dLbl.series = oSeries;
                }
                dLbl.series = this.series;
                dLbl.chart = this.chart;
                return dLbl;
            }
        }
        return null;
    };
    CDLbl.prototype.checkDocContent = function() {
        var oDlbl = this.checkDlbl();
        if(oDlbl) {
            oDlbl.txBody = this.txBody;
            CTitle.prototype.checkDocContent.call(oDlbl);
            this.txBody = oDlbl.tx.rich;
            if(oDlbl.tx && oDlbl.tx.rich && oDlbl.tx.rich.content) {
                if(!oDlbl.tx.rich.content.DrawingDocument) {
                    if(this.chart) {
                        var oDrawingDocument = this.chart.getDrawingDocument();
                        if(oDrawingDocument) {
                            oDlbl.tx.rich.content.DrawingDocument = oDrawingDocument;
                            var aContent = oDlbl.tx.rich.content.Content;
                            for(var i = 0; i < aContent.length; ++i) {
                                aContent[i].DrawingDocument = oDrawingDocument;
                            }
                        }
                    }

                }
            }
        }
    };
    CDLbl.prototype.applyTextFunction = function(docContentFunction, tableFunction, args) {
        var oDlbl = this.checkDlbl();
        if(!oDlbl) {
            return;
        }

        if(oDlbl.tx && oDlbl.tx.rich && oDlbl.tx.rich.content) {
            docContentFunction.apply(oDlbl.tx.rich.content, args);
        }
    };
    CDLbl.prototype.hit = function(x, y) {
        var tx = this.invertTransform.TransformPointX(x, y);
        var ty = this.invertTransform.TransformPointY(x, y);
        return tx >= 0 && tx <= this.extX && ty >= 0 && ty <= this.extY;
    };
    CDLbl.prototype.hitInPath = CShape.prototype.hitInPath;
    CDLbl.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
    CDLbl.prototype.hitInBoundingRect = CShape.prototype.hitInBoundingRect;
    CDLbl.prototype.hitInTextRect = function(x, y) {
        var content = this.getDocContent && this.getDocContent();
        if (content && this.invertTransformText) {
            return AscFormat.HitToRect(x, y, this.invertTransformText, 0, 0, this.contentWidth, this.contentHeight);
        }
    };
    CDLbl.prototype.getCompiledStyle = function() {
        return null;
    };
    CDLbl.prototype.getChartSpace = function() {
        if(this.chart) {
            return this.chart;
        }
        return CBaseChartObject.prototype.getChartSpace.call(this);
    };
    CDLbl.prototype.getParentObjects = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getParentObjects();
        }
    };
    CDLbl.prototype.recalculateTransform = function() {
    };
    CDLbl.prototype.recalculateTransformText = function() {
        if(this.txBody === null)
            return;
        this.ownTransformText.Reset();
        var _text_transform = this.ownTransformText;
        var _shape_transform = this.ownTransform;
        var _body_pr = this.getBodyPr();
        var _content_height = this.txBody.content.GetSummaryHeight();
        var _l, _t, _r, _b;

        var _t_x_lt, _t_y_lt, _t_x_rt, _t_y_rt, _t_x_lb, _t_y_lb, _t_x_rb, _t_y_rb;
        if(isRealObject(this.spPr) && isRealObject(this.spPr.geometry) && isRealObject(this.spPr.geometry.rect)) {
            var _rect = this.spPr.geometry.rect;
            _l = _rect.l + _body_pr.lIns;
            _t = _rect.t + _body_pr.tIns;
            _r = _rect.r - _body_pr.rIns;
            _b = _rect.b - _body_pr.bIns;
        }
        else {
            _l = _body_pr.lIns;
            _t = _body_pr.tIns;
            _r = this.extX - _body_pr.rIns;
            _b = this.extY - _body_pr.bIns;
        }

        if(_l >= _r) {
            var _c = (_l + _r) * 0.5;
            _l = _c - 0.01;
            _r = _c + 0.01;
        }

        if(_t >= _b) {
            _c = (_t + _b) * 0.5;
            _t = _c - 0.01;
            _b = _c + 0.01;
        }

        _t_x_lt = _shape_transform.TransformPointX(_l, _t);
        _t_y_lt = _shape_transform.TransformPointY(_l, _t);

        _t_x_rt = _shape_transform.TransformPointX(_r, _t);
        _t_y_rt = _shape_transform.TransformPointY(_r, _t);

        _t_x_lb = _shape_transform.TransformPointX(_l, _b);
        _t_y_lb = _shape_transform.TransformPointY(_l, _b);

        _t_x_rb = _shape_transform.TransformPointX(_r, _b);
        _t_y_rb = _shape_transform.TransformPointY(_r, _b);

        var _dx_t, _dy_t;
        _dx_t = _t_x_rt - _t_x_lt;
        _dy_t = _t_y_rt - _t_y_lt;

        var _dx_lt_rb, _dy_lt_rb;
        _dx_lt_rb = _t_x_rb - _t_x_lt;
        _dy_lt_rb = _t_y_rb - _t_y_lt;

        var _vertical_shift;
        var _text_rect_height = _b - _t;
        var _text_rect_width = _r - _l;
        var nVert = _body_pr.vert;


        if(!_body_pr.upright) {
            if(!(nVert === AscFormat.nVertTTvert || nVert === AscFormat.nVertTTvert270)) {
                switch(_body_pr.anchor) {
                    case 0: //b
                    { // (Text Anchor Enum ( Bottom ))
                        _vertical_shift = _text_rect_height - _content_height;
                        break;
                    }
                    case 1:    //ctr
                    {// (Text Anchor Enum ( Center ))
                        _vertical_shift = (_text_rect_height - _content_height) * 0.5;
                        break;
                    }
                    case 2: //dist
                    {// (Text Anchor Enum ( Distributed )) TODO: пока выравнивание  по центру. Переделать!
                        _vertical_shift = (_text_rect_height - _content_height) * 0.5;
                        break;
                    }
                    case 3: //just
                    {// (Text Anchor Enum ( Justified )) TODO: пока выравнивание  по центру. Переделать!
                        _vertical_shift = (_text_rect_height - _content_height) * 0.5;
                        break;
                    }
                    case 4: //t
                    {//Top
                        _vertical_shift = 0;
                        break;
                    }
                }
                global_MatrixTransformer.TranslateAppend(_text_transform, 0, _vertical_shift);
                if(_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
                    var alpha = Math.atan2(_dy_t, _dx_t);
                    global_MatrixTransformer.RotateRadAppend(_text_transform, -alpha);
                    global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_lt, _t_y_lt);
                }
                else {
                    alpha = Math.atan2(_dy_t, _dx_t);
                    global_MatrixTransformer.RotateRadAppend(_text_transform, Math.PI - alpha);
                    global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_rt, _t_y_rt);
                }
            }
            else {
                switch(_body_pr.anchor) {
                    case 0: //b
                    { // (Text Anchor Enum ( Bottom ))
                        _vertical_shift = _text_rect_width - _content_height;
                        break;
                    }
                    case 1:    //ctr
                    {// (Text Anchor Enum ( Center ))
                        _vertical_shift = (_text_rect_width - _content_height) * 0.5;
                        break;
                    }
                    case 2: //dist
                    {// (Text Anchor Enum ( Distributed ))
                        _vertical_shift = (_text_rect_width - _content_height) * 0.5;
                        break;
                    }
                    case 3: //just
                    {// (Text Anchor Enum ( Justified ))
                        _vertical_shift = (_text_rect_width - _content_height) * 0.5;
                        break;
                    }
                    case 4: //t
                    {//Top
                        _vertical_shift = 0;
                        break;
                    }
                }
                global_MatrixTransformer.TranslateAppend(_text_transform, 0, _vertical_shift);
                var _alpha;
                _alpha = Math.atan2(_dy_t, _dx_t);
                if(nVert === AscFormat.nVertTTvert) {
                    if(_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
                        global_MatrixTransformer.RotateRadAppend(_text_transform, -_alpha - Math.PI * 0.5);
                        global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_rt, _t_y_rt);
                    }
                    else {
                        global_MatrixTransformer.RotateRadAppend(_text_transform, Math.PI * 0.5 - _alpha);
                        global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_lt, _t_y_lt);
                    }
                }
                else {
                    if(_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
                        global_MatrixTransformer.RotateRadAppend(_text_transform, -_alpha - Math.PI * 1.5);
                        global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_lb, _t_y_lb);
                    }
                    else {
                        global_MatrixTransformer.RotateRadAppend(_text_transform, -Math.PI * 0.5 - _alpha);
                        global_MatrixTransformer.TranslateAppend(_text_transform, _t_x_rb, _t_y_rb);
                    }
                }
            }
        }
        else {

            var _full_flip = {flipH: false, flipV: false};

            var _hc = this.extX * 0.5;
            var _vc = this.extY * 0.5;
            var _transformed_shape_xc = this.transform.TransformPointX(_hc, _vc);
            var _transformed_shape_yc = this.transform.TransformPointY(_hc, _vc);


            var _content_width, content_height2;
            if(!(nVert === AscFormat.nVertTTvert || nVert === AscFormat.nVertTTvert270)) {
                _content_width = _r - _l;
                content_height2 = _b - _t;
            }
            else {
                _content_width = _b - _t;
                content_height2 = _r - _l;
            }

            switch(_body_pr.anchor) {
                case 0: //b
                { // (Text Anchor Enum ( Bottom ))
                    _vertical_shift = content_height2 - _content_height;
                    break;
                }
                case 1:    //ctr
                {// (Text Anchor Enum ( Center ))
                    _vertical_shift = (content_height2 - _content_height) * 0.5;
                    break;
                }
                case 2: //dist
                {// (Text Anchor Enum ( Distributed ))
                    _vertical_shift = (content_height2 - _content_height) * 0.5;
                    break;
                }
                case 3: //just
                {// (Text Anchor Enum ( Justified ))
                    _vertical_shift = (content_height2 - _content_height) * 0.5;
                    break;
                }
                case 4: //t
                {//Top
                    _vertical_shift = 0;
                    break;
                }
            }

            var _text_rect_xc = _l + (_r - _l) * 0.5;
            var _text_rect_yc = _t + (_b - _t) * 0.5;

            var _vx = _text_rect_xc - _hc;
            var _vy = _text_rect_yc - _vc;

            var _transformed_text_xc, _transformed_text_yc;
            if(!_full_flip.flipH) {
                _transformed_text_xc = _transformed_shape_xc + _vx;
            }
            else {
                _transformed_text_xc = _transformed_shape_xc - _vx;
            }

            if(!_full_flip.flipV) {
                _transformed_text_yc = _transformed_shape_yc + _vy;
            }
            else {
                _transformed_text_yc = _transformed_shape_yc - _vy;
            }

            global_MatrixTransformer.TranslateAppend(_text_transform, 0, _vertical_shift);
            if(nVert === AscFormat.nVertTTvert) {
                global_MatrixTransformer.TranslateAppend(_text_transform, -_content_width * 0.5, -content_height2 * 0.5);
                global_MatrixTransformer.RotateRadAppend(_text_transform, -Math.PI * 1.5);
                global_MatrixTransformer.TranslateAppend(_text_transform, _content_width * 0.5, content_height2 * 0.5);

            }
            if(nVert === AscFormat.nVertTTvert270) {
                global_MatrixTransformer.TranslateAppend(_text_transform, -_content_width * 0.5, -content_height2 * 0.5);
                global_MatrixTransformer.RotateRadAppend(_text_transform, -Math.PI * 1.5);
                global_MatrixTransformer.TranslateAppend(_text_transform, _content_width * 0.5, content_height2 * 0.5);
            }
            global_MatrixTransformer.TranslateAppend(_text_transform, _transformed_text_xc - _content_width * 0.5, _transformed_text_yc - content_height2 * 0.5);

            this.clipRect =
            {
                x: -_body_pr.lIns,
                y: -_vertical_shift - _body_pr.tIns,
                w: this.contentWidth + (_body_pr.rIns + _body_pr.lIns),
                h: this.contentHeight + (_body_pr.bIns + _body_pr.tIns)
            };
        }

        this.transformText = this.ownTransformText.CreateDublicate();
    };
    CDLbl.prototype.getStyles = function() {
        if(this.lastStyleObject)
            return this.lastStyleObject;
        return AscFormat.ExecuteNoHistory(function() {
            var styles = new CStyles(false);
            var style = new CStyle("dataLblStyle", null, null, null, true);
            var text_pr = new CTextPr();
            text_pr.FontSize = 10;
            var oChartSpace = this.getChartSpace();
            if(oChartSpace && AscFormat.isRealNumber(oChartSpace.style)) {

                if(oChartSpace.style > 40) {
                    text_pr.Unifill = AscFormat.CreateUnfilFromRGB(255, 255, 255);
                }
                else {
                    var default_style = AscFormat.CHART_STYLE_MANAGER.getDefaultLineStyleByIndex(oChartSpace.style);
                    var oUnifill = default_style.axisAndMajorGridLines.createDuplicate();
                    if(oUnifill && oUnifill.fill && oUnifill.fill.color && oUnifill.fill.color.Mods) {
                        oUnifill.fill.color.Mods.Mods.length = 0;
                    }
                    text_pr.Unifill = oUnifill;
                }
            }
            else {
                text_pr.Unifill = AscFormat.CreateUnfilFromRGB(0, 0, 0);
            }

            var para_pr = new CParaPr();
            para_pr.Jc = AscCommon.align_Center;
            para_pr.Spacing.Before = 0.0;
            para_pr.Spacing.After = 0.0;
            para_pr.Spacing.Line = 1;
            para_pr.Spacing.LineRule = Asc.linerule_Auto;
            style.ParaPr = para_pr;
            text_pr.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);
            style.TextPr = text_pr;
            var chart_text_pr;

            var oParaPr = oChartSpace && oChartSpace.getTxPrParaPr();
            if(oParaPr) {
                style.ParaPr.Merge(oParaPr);
                if(oParaPr.DefaultRunPr) {
                    chart_text_pr = oParaPr.DefaultRunPr;
                    style.TextPr.Merge(chart_text_pr);
                }
            }
            if(this instanceof CTitle || this.parent instanceof CTitle) {
                style.TextPr.Bold = true;
                if(this.parent instanceof CChart || (this.parent && (this.parent.parent instanceof CChart))) {
                    if(chart_text_pr && typeof chart_text_pr.FontSize === "number")
                        style.TextPr.FontSize = (chart_text_pr.FontSize * 1.2) >> 0;
                    else
                        style.TextPr.FontSize = 18;
                }
            }
            if(this instanceof CalcLegendEntry && this.legend) {
                oParaPr = this.legend.getTxPrParaPr();
                if(oParaPr) {
                    style.ParaPr.Merge(oParaPr);
                    if(oParaPr.DefaultRunPr)
                        style.TextPr.Merge(oParaPr.DefaultRunPr);
                }
                if(AscFormat.isRealNumber(this.idx)) {
                    var aLegendEntries = this.legend.legendEntryes;
                    for(var i = 0; i < aLegendEntries.length; ++i) {
                        if(this.idx === aLegendEntries[i].idx) {
                            var oLegendEntry = aLegendEntries[i];
                            oParaPr = oLegendEntry.getTxPrParaPr();
                            if(oParaPr) {
                                style.ParaPr.Merge(oParaPr);
                                if(oParaPr.DefaultRunPr)
                                    style.TextPr.Merge(oParaPr.DefaultRunPr);
                            }
                            break;
                        }
                    }
                }

            }
            if(!(this instanceof CTitle)) {
                if(this.parent) {
                    oParaPr = this.parent.getTxPrParaPr();
                    if(oParaPr) {
                        style.ParaPr.Merge(oParaPr);
                        if(oParaPr.DefaultRunPr)
                            style.TextPr.Merge(oParaPr.DefaultRunPr);
                    }
                }
            }
            oParaPr = this.getTxPrParaPr();
            if(oParaPr) {
                style.ParaPr.Merge(oParaPr);
                if(oParaPr.DefaultRunPr)
                    style.TextPr.Merge(oParaPr.DefaultRunPr);
            }
            styles.Add(style);
            if(!(this instanceof CTitle))
                this.lastStyleObject = {lastId: style.Id, styles: styles, shape: this, slide: null};
            return {lastId: style.Id, styles: styles, shape: this, slide: null};
        }, this, []);
    };
    CDLbl.prototype.Get_Theme = function() {
        if(this.chart) {
            return this.chart.Get_Theme();
        }
        else {
            if(this.series && this.series.Get_Theme) {
                return this.series.Get_Theme();
            }
        }
        return null;
    };
    CDLbl.prototype.Get_ColorMap = function() {
        if(this.chart) {
            return this.chart.Get_ColorMap();
        }
        else {
            if(this.series && this.series.Get_ColorMap) {
                return this.series.Get_ColorMap();
            }
        }
        return AscFormat.GetDefaultColorMap();
    };
    CDLbl.prototype.Get_AbsolutePage = function() {
        if(this.chart && this.chart.Get_AbsolutePage) {
            return this.chart.Get_AbsolutePage();
        }
        return 0;
    };
    CDLbl.prototype.recalculateStyle = function() {
        AscFormat.ExecuteNoHistory(function() {
                this.compiledStyles = this.getStyles();
            },
            this, []);
    };
    CDLbl.prototype.Get_Styles = function(lvl) {
        if(this.recalcInfo.recalcStyle) {
            this.recalculateStyle();
            this.recalcInfo.recalcStyle = false;
        }
        return this.compiledStyles;
    };
    CDLbl.prototype.checkNoLbl = function() {
        if(this.tx && this.tx.rich)
            return false;
        else {
            return !(this.showSerName || this.showCatName || this.showVal || this.showPercent);
        }
    };
    CDLbl.prototype.getPercentageString = function() {
        if(this.series && this.pt) {
            return this.series.getStrPercentageValByIndex(this.pt.idx)
        }
        return "";
    };
    CDLbl.prototype.getValueString = function() {
        var sFormatCode;
        if(this.pt && this.series) {
            if(this.numFmt && typeof this.numFmt.formatCode === "string" && this.numFmt.formatCode.length > 0) {
                sFormatCode = this.numFmt.formatCode;
            }
            else if(typeof this.pt.formatCode === "string" && this.pt.formatCode.length > 0) {
                sFormatCode = this.pt.formatCode;
            }
            else {
                sFormatCode = this.series.getValObjectSourceNumFormat(this.pt.idx);
            }

            var num_format = AscCommon.oNumFormatCache.get(sFormatCode);
            return num_format.formatToChart(this.pt.val)
        }
        return "";
    };
    CDLbl.prototype.getDefaultTextForTxBody = function() {
        var compiled_string = "";
        var separator;
        if(typeof this.separator === "string") {
            separator = this.separator + " ";
        }
        else if(this.series.getObjectType() === AscDFH.historyitem_type_PieSeries) {
            if(this.showPercent && this.showCatName && !this.showSerName && !this.showVal) {
                separator = "\n";
            }
            else {
                separator = ", "
            }
        }
        else {
            separator = ", "
        }
        if(this.showSerName) {
            compiled_string += this.series.getSeriesName();
        }
        if(this.showCatName) {
            if(compiled_string.length > 0)
                compiled_string += separator;
            compiled_string += this.series.getCatName(this.pt.idx);
        }
        if(this.showVal) {
            if(compiled_string.length > 0)
                compiled_string += separator;
            compiled_string += this.getValueString();
        }
        if(this.showPercent) {
            if(compiled_string.length > 0)
                compiled_string += separator;
            compiled_string += this.getPercentageString();
        }
        return compiled_string;
    };
    CDLbl.prototype.getMaxWidth = function(bodyPr) {
        var oChartSpace = this.getChartSpace();
        if(!(this.parent && (this.parent.axPos === AX_POS_L || this.parent.axPos === AX_POS_R))) {
            if(!oChartSpace) {
                return 20000;
            }
            switch(bodyPr.vert) {
                case AscFormat.nVertTTeaVert:
                case AscFormat.nVertTTmongolianVert:
                case AscFormat.nVertTTvert:
                case AscFormat.nVertTTwordArtVert:
                case AscFormat.nVertTTwordArtVertRtl:
                case AscFormat.nVertTTvert270:
                {
                    return oChartSpace.extY / 2;
                }
                case AscFormat.nVertTThorz:
                {
                    return oChartSpace.extX / 5
                }
            }
            return oChartSpace.extX / 5;
        }
        else {
            return 20000;//надписи для осей значений не переносятся поэтому выставляем большую ширину.
        }
    };
    CDLbl.prototype.getBodyPr = function() {
        var ret = new AscFormat.CBodyPr();
        ret.setDefault();
        ret.anchor = 1;
        var oBaseBodyPr = new AscFormat.CBodyPr();

        if(this.txPr && this.txPr.bodyPr) {
            oBaseBodyPr.merge(this.txPr.bodyPr);
        }
        if(this.tx && this.tx.rich) {
            oBaseBodyPr.merge(this.tx.rich.bodyPr);
        }
        if(this.parent && (this.parent.axPos === AX_POS_L || this.parent.axPos === AX_POS_R)
            && (oBaseBodyPr.vert === null && oBaseBodyPr.rot === null)) {
            ret.vert = AscFormat.nVertTTvert270;
        }
        ret.merge(oBaseBodyPr);
        var nVert = ret.vert;
        //Пока не поддерживаем bodyPr.rot. Костыль под эффект_штурмовика.docx.
        if(AscFormat.isRealNumber(ret.rot) && 0 !== ret.rot) {
            if(Math.abs(ret.rot - 5400000) < 1000) {
                if(ret.vert === AscFormat.nVertTTvert270) {
                    nVert = AscFormat.nVertTThorz;
                }
                else if(ret.vert === AscFormat.nVertTThorz) {
                    nVert = AscFormat.nVertTTvert;
                }
            }
            else if(Math.abs(ret.rot + 5400000) < 1000) {
                if(ret.vert === AscFormat.nVertTTvert) {
                    nVert = AscFormat.nVertTThorz;
                }
                else if(ret.vert === AscFormat.nVertTThorz) {
                    nVert = AscFormat.nVertTTvert270;
                }
            }
        }
        //

        switch(nVert) {
            case AscFormat.nVertTTeaVert:
            case AscFormat.nVertTTmongolianVert:
            case AscFormat.nVertTTvert:
            case AscFormat.nVertTTwordArtVert:
            case AscFormat.nVertTTwordArtVertRtl:
            case AscFormat.nVertTTvert270:
            {
                ret.lIns = SCALE_INSET_COEFF;
                ret.rIns = SCALE_INSET_COEFF;
                ret.tIns = SCALE_INSET_COEFF * 0.5;
                ret.bIns = SCALE_INSET_COEFF * 0.5;
                break;
            }
            case AscFormat.nVertTThorz:
            {
                ret.lIns = SCALE_INSET_COEFF;
                ret.rIns = SCALE_INSET_COEFF;
                ret.tIns = SCALE_INSET_COEFF * 0.5;
                ret.bIns = SCALE_INSET_COEFF * 0.5;
                break;
            }
        }
        ret.vert = nVert;
        return ret;
    };
    CDLbl.prototype.recalculateContent = function() {
        if(this.txBody) {
            var bodyPr = this.getBodyPr();
            var max_box_width = this.getMaxWidth(bodyPr);
            /*получено экспериментальным путем нужно уточнить*/
            var max_content_width = max_box_width - 2 * SCALE_INSET_COEFF;

            var content = this.txBody.content;


            var sParPasteId = null;
            if(window['AscCommon'].g_specialPasteHelper && window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph) {
                sParPasteId = window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph;
                window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph = null;
            }

            content.RecalculateContent(max_content_width, 20000, 0);
            if(sParPasteId) {
                window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph = sParPasteId;
            }
            // content.Reset(0, 0, max_content_width, 20000);
            //
            // content.Recalculate_Page(0, true);
            var pargs = content.Content;
            var max_width = 0;
            for(var i = 0; i < pargs.length; ++i) {
                var par = pargs[i];
                for(var j = 0; j < par.Lines.length; ++j) {
                    if(par.Lines[j].Ranges[0].W > max_width) {
                        max_width = par.Lines[j].Ranges[0].W;
                    }
                }
            }
            max_width += 1;
            // content.Reset(0, 0, max_width, 20000);
            // content.Recalculate_Page(0, true);

            content.RecalculateContent(max_width, 20000, 0);
            switch(bodyPr.vert) {
                case AscFormat.nVertTTeaVert:
                case AscFormat.nVertTTmongolianVert:
                case AscFormat.nVertTTvert:
                case AscFormat.nVertTTwordArtVert:
                case AscFormat.nVertTTwordArtVertRtl:
                case AscFormat.nVertTTvert270:
                {
                    this.extX = Math.min(content.GetSummaryHeight() + 4.4 * SCALE_INSET_COEFF, max_box_width);
                    this.extY = max_width + 2 * SCALE_INSET_COEFF;
                    this.x = 0;
                    this.y = 0;

                    this.txBody.contentWidth = this.extY;
                    this.txBody.contentHeight = this.extX;
                    this.contentWidth = this.extY;
                    this.contentHeight = this.extX;
                    break;
                }
                default:
                {
                    var _rot = AscFormat.isRealNumber(bodyPr.rot) ? bodyPr.rot * AscFormat.cToRad2 : 0;
                    var t = new CMatrix();
                    global_MatrixTransformer.RotateRadAppend(t, -_rot);
                    var w, h, x0, y0, x1, y1, x2, y2, x3, y3;
                    w = max_width;
                    h = this.txBody.content.GetSummaryHeight();
                    x0 = 0;
                    y0 = 0;
                    x1 = t.TransformPointX(w, 0);
                    y1 = t.TransformPointY(w, 0);
                    x2 = t.TransformPointX(w, h);
                    y2 = t.TransformPointY(w, h);
                    x3 = t.TransformPointX(0, h);
                    y3 = t.TransformPointY(0, h);

                    this.extX = Math.max(x0, x1, x2, x3) - Math.min(x0, x1, x2, x3) + 1.25;
                    this.extY = Math.max(y0, y1, y2, y3) - Math.min(y0, y1, y2, y3) + SCALE_INSET_COEFF;
                    this.x = 0;
                    this.y = 0;

                    this.txBody.contentWidth = this.extX;
                    this.txBody.contentHeight = this.extY;
                    this.contentWidth = this.extX;
                    this.contentHeight = this.extY;
                    break;
                }
            }

        }
    };
    CDLbl.prototype.recalculateTxBody = function() {
		if(this.showDlblsRange) {
			let sText;
			if(this.series && this.pt) {
				let oLblsRange = this.series.datalabelsRange;
				if(oLblsRange) {
					let oCache = oLblsRange.strCache;
					if(oCache) {
						let oPt = oCache.getPtByIndex(this.pt.idx);
						sText = oPt && oPt.val;
					}

				}
			}
			if(sText) {
				this.txBody = AscFormat.CreateTextBodyFromString(sText, this.getDrawingDocument(), this);
				return;
			}

		}
        if(this.tx && this.tx.rich) {
            this.txBody = this.tx.rich;
            this.txBody.parent = this;
        }
        else {
            this.txBody = AscFormat.CreateTextBodyFromString(this.getDefaultTextForTxBody(), this.getDrawingDocument(), this);
        }
    };
    CDLbl.prototype.initDefault = function(nDefaultPosition) {
        this.setDelete(false);
        this.setDLblPos(AscFormat.isRealNumber(nDefaultPosition) ? nDefaultPosition : c_oAscChartDataLabelsPos.inBase);
        this.setIdx(null);
        this.setLayout(null);
        this.setNumFmt(null);
        this.setSeparator(null);
        this.setShowBubbleSize(false);
        this.setShowCatName(false);
        this.setShowLegendKey(false);
        this.setShowPercent(false);
        this.setShowSerName(false);
        this.setShowVal(false);
        this.setSpPr(null);
        this.setTx(null);
        this.setTxPr(null);
    };
    CDLbl.prototype.merge = function(dLbl, noCopyTxBody) {
        if(!dLbl)
            return;
        if(dLbl.bDelete != null)
            this.setDelete(dLbl.bDelete);
        if(dLbl.dLblPos != null)
            this.setDLblPos(dLbl.dLblPos);

        if(dLbl.idx != null)
            this.setIdx(dLbl.idx);

        if(dLbl.layout != null) {
            this.setLayout(dLbl.layout.createDuplicate());
        }

        if(dLbl.numFmt != null)
            this.setNumFmt(dLbl.numFmt);

        if(dLbl.separator != null)
            this.setSeparator(dLbl.separator);

        if(dLbl.showBubbleSize != null)
            this.setShowBubbleSize(dLbl.showBubbleSize);

        if(dLbl.showCatName != null)
            this.setShowCatName(dLbl.showCatName);

        if(dLbl.showLegendKey != null)
            this.setShowLegendKey(dLbl.showLegendKey);

        if(dLbl.showPercent != null)
            this.setShowPercent(dLbl.showPercent);

        if(dLbl.showSerName != null)
            this.setShowSerName(dLbl.showSerName);


        if(dLbl.showVal != null)
            this.setShowVal(dLbl.showVal);
        if(dLbl.showDlblsRange != null)
            this.setShowDlblsRange(dLbl.showDlblsRange);

        if(dLbl.spPr != null) {
            if(this.spPr == null) {
                this.setSpPr(new AscFormat.CSpPr());
            }
            if(dLbl.spPr.Fill) {
                if(this.spPr.Fill == null) {
                    this.spPr.setFill(new AscFormat.CUniFill());
                }
                this.spPr.Fill.merge(dLbl.spPr.Fill);
            }
            if(dLbl.spPr.ln) {
                if(this.spPr.ln == null) {
                    this.spPr.setLn(new AscFormat.CLn());
                }
                this.spPr.ln.merge(dLbl.spPr.ln);
            }
        }
        if(dLbl.tx) {
            if(this.tx == null) {
                this.setTx(new CChartText());
            }
            this.tx.merge(dLbl.tx, noCopyTxBody);
        }
        if(dLbl.txPr) {
            if(noCopyTxBody === true) {
                var oldParent = dLbl.txPr.parent;
                this.setTxPr(dLbl.txPr);
                dLbl.txPr.parent = oldParent;
            }
            else {
                this.setTxPr(dLbl.txPr.createDuplicate());
            }
            this.txPr.setParent(this);
        }
    };
    CDLbl.prototype.draw = CShape.prototype.draw;
    CDLbl.prototype.isEmptyPlaceholder = function() {
        return false;
    };
    CDLbl.prototype.setPosition = function(x, y) {
        this.x = x;
        this.y = y;


        this.calcX = this.x;
        this.calcY = this.y;

        this.localTransform.Reset();
        global_MatrixTransformer.TranslateAppend(this.localTransform, this.calcX, this.calcY);
        //  if (isRealObject(this.chart))
        //  {
        //      global_MatrixTransformer.MultiplyAppend(this.localTransform, this.chart.localTransform);
        //  }

        this.transform = this.localTransform.CreateDublicate();
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);


        this.localTransformText = this.ownTransformText.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.localTransformText, this.calcX, this.calcY);
        //  if (isRealObject(this.chart))
        //  {
        //      global_MatrixTransformer.MultiplyAppend(this.localTransformText, this.chart.localTransform);
        //  }

        this.transformText = this.localTransformText.CreateDublicate();
        this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);
    };
    CDLbl.prototype.setPosition2 = function(x, y) {
        this.x = x;
        this.y = y;

        this.calcX = this.x;
        this.calcY = this.y;


        this.localTransform.tx = x;
        this.localTransform.ty = y;

        this.transform = this.localTransform.CreateDublicate();
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);


        this.localTransformText.tx = x;
        this.localTransformText.ty = y;
        this.transformText = this.localTransformText.CreateDublicate();
        this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);
    };
    CDLbl.prototype.updateTransformMatrix = function() {
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, this.posX, this.posY);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);

        if(this.localTransformText) {
            this.transformText = this.localTransformText.CreateDublicate();
            global_MatrixTransformer.TranslateAppend(this.transformText, this.posX, this.posY);
            this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);
        }

    };
    CDLbl.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, x, y);

        this.transformText = this.localTransformText.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transformText, x, y);


        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);

    };
    CDLbl.prototype.setDelete = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetDelete, this.bDelete, pr));
        this.bDelete = pr;
        this.Refresh_RecalcData2();
    };
    CDLbl.prototype.setDLblPos = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DLbl_SetDLblPos, this.dLblPos, pr));
        this.dLblPos = pr;
    };
    CDLbl.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DLbl_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CDLbl.prototype.setLayout = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbl_SetLayout, this.layout, pr));
        this.layout = pr;
        this.setParentToChild(pr);
    };
    CDLbl.prototype.setNumFmt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbl_SetNumFmt, this.numFmt, pr));
        this.numFmt = pr;
        this.setParentToChild(pr);
    };
    CDLbl.prototype.setSeparator = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_DLbl_SetSeparator, this.separator, pr));
        this.separator = pr;
    };
    CDLbl.prototype.setShowBubbleSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowBubbleSize, this.showBubbleSize, pr));
        this.showBubbleSize = pr;
    };
    CDLbl.prototype.setShowCatName = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowCatName, this.showCatName, pr));
        this.showCatName = pr;
    };
    CDLbl.prototype.setShowLegendKey = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowLegendKey, this.showLegendKey, pr));
        this.showLegendKey = pr;
    };
    CDLbl.prototype.setShowPercent = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowPercent, this.showPercent, pr));
        this.showPercent = pr;
    };
    CDLbl.prototype.setShowSerName = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowSerName, this.showSerName, pr));
        this.showSerName = pr;
    };
    CDLbl.prototype.setShowVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowVal, this.showVal, pr));
        this.showVal = pr;
    };
    CDLbl.prototype.setShowDlblsRange = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbl_SetShowDLblsRange, this.showDlblsRange, pr));
        this.showDlblsRange = pr;
    };
    CDLbl.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbl_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CDLbl.prototype.setTx = function(pr) {
        if(this.tx && this.tx.strRef || pr && pr.strRef) {
            this.onChangeDataRefs();
        }
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbl_SetTx, this.tx, pr));
        this.tx = pr;
        this.setParentToChild(pr);
    };
    CDLbl.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbl_SetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
    };
    CDLbl.prototype.handleUpdateFill = function() {
        this.Refresh_RecalcData2();
    };
    CDLbl.prototype.handleUpdateLn = function() {
        this.Refresh_RecalcData2();
    };
    CDLbl.prototype.Refresh_RecalcData2 = function(pageIndex) {
        if(this.parent && this.parent.Refresh_RecalcData2) {
            this.parent.Refresh_RecalcData2(pageIndex, this);
        }
        else {
            if(this.chart) {
                this.chart.Refresh_RecalcData2(pageIndex, this);
            }
        }
    };
    CDLbl.prototype.checkPosition = function(aPositions) {
        fCheckDLblPosition(this, aPositions);
    };
    CDLbl.prototype.setSettings = function(nPos, oProps) {
        fCheckDLblSettings(this, nPos, oProps)
    };
    CDLbl.prototype.correctValues = function() {
        if(this.bDelete !== true){
            if(null === this.showLegendKey){
                this.setShowLegendKey(false);
            }
            if(null === this.showVal){
                this.setShowVal(false);
            }
            if(null === this.showCatName){
                this.setShowCatName(false);
            }
            if(null === this.showSerName){
                this.setShowSerName(false);
            }
            if(null === this.showPercent){
                this.setShowPercent(false);
            }
            if(null === this.showBubbleSize){
                this.setShowBubbleSize(false);
            }
            if(this.setShowLeaderLines && null === this.showLeaderLines){
                this.setShowLeaderLines(false);
            }
        }
    };


    function CSeriesBase() {
        CBaseChartObject.call(this);
        this.idx = null;
        this.order = null;
        this.tx = null;
        this.spPr = null;
		this.datalabelsRange = null;
    }

    InitClass(CSeriesBase, CBaseChartObject, AscDFH.historyitem_type_Unknown);
    CSeriesBase.prototype.clearDataCache = function() {
        if(this.val) {
            this.val.clearDataCache();
        }
        if(this.yVal) {
            this.yVal.clearDataCache();
        }
        if(this.cat) {
            this.cat.clearDataCache();
        }
        if(this.xVal) {
            this.xVal.clearDataCache();
        }
        if(this.tx) {
            this.tx.clearDataCache();
        }
        if(this.errBars) {
            this.errBars.clearDataCache();
        }
    };
    CSeriesBase.prototype.updateData = function(displayEmptyCellsAs, displayHidden) {
        if(this.val) {
            this.val.update(displayEmptyCellsAs, displayHidden, this);
        }
        if(this.yVal) {
            this.yVal.update(displayEmptyCellsAs, displayHidden, this);
        }
        if(this.cat) {
            this.cat.update(this);
        }
        if(this.xVal) {
            this.xVal.update(this);
        }
        if(this.tx) {
            this.tx.update();
        }
        if(this.errBars) {
            this.errBars.update();
        }
    };
    CSeriesBase.prototype.Refresh_RecalcData = function(oData) {
        this.onChartUpdateType();
    };
    CSeriesBase.prototype.setIdx = function(val) {
        if(this.idx !== val) {
            if(History.CanAddChanges()) {
                History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CommonSeries_SetIdx, this.idx, val));
            }
            this.idx = val;
        }
    };
    CSeriesBase.prototype.setOrder = function(val) {
        if(this.order !== val) {
            if(History.CanAddChanges()) {
                History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CommonSeries_SetOrder, this.order, val));
            }
            this.order = val;
        }
    };
    CSeriesBase.prototype.setTx = function(val) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CommonSeries_SetTx, this.tx, val));
        this.tx = val;
        this.setParentToChild(val);
        this.onChangeDataRefs();
    };
    CSeriesBase.prototype.setSpPr = function(val) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CommonSeries_SetSpPr, this.spPr, val));
        this.spPr = val;
        this.setParentToChild(val);
    };
    CSeriesBase.prototype.setDataLabelsRange = function(val) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CommonChart_DataLabelsRange, this.spPr, val));
        this.datalabelsRange = val;
        this.setParentToChild(val);
    };
    CSeriesBase.prototype.getChildren = function() {
        return [this.spPr, this.tx, this.cat || this.xVal, this.val || this.yVal];
    };
    CSeriesBase.prototype.fillObject = function(oCopy, oIdMap) {
        if(AscFormat.isRealNumber(this.idx)) {
            oCopy.setIdx(this.idx);
        }
        if(AscFormat.isRealNumber(this.order)) {
            oCopy.setOrder(this.order);
        }
        if(this.dLbls && oCopy.setDLbls) {
            oCopy.setDLbls(this.dLbls.createDuplicate());
        }
        if(Array.isArray(this.dPt) && this.dPt.length > 0 && oCopy.addDPt) {
            for(var nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
                oCopy.addDPt(this.dPt[nDpt].createDuplicate());
            }
        }
        if(AscCommon.isRealObject(this.spPr)) {
            oCopy.setSpPr(this.spPr.createDuplicate());
            var nCopyType = oCopy.getObjectType();
            var nThisType = this.getObjectType();
            if(!(nCopyType === AscDFH.historyitem_type_LineSeries || nCopyType === AscDFH.historyitem_type_ScatterSer)
                && (nThisType === AscDFH.historyitem_type_LineSeries || nThisType === AscDFH.historyitem_type_ScatterSer)) {
                if(oCopy.hasNoFill()) {
                    oCopy.spPr.setFill(null);
                }
            }
            if((nCopyType === AscDFH.historyitem_type_LineSeries || nCopyType === AscDFH.historyitem_type_ScatterSer)
                && !(nThisType === AscDFH.historyitem_type_LineSeries || nThisType === AscDFH.historyitem_type_ScatterSer)) {
                if(oCopy.hasNoFillLine()) {
                    oCopy.spPr.setLn(null);
                }
            }
        }
        if(AscCommon.isRealObject(this.tx)) {
            oCopy.setTx(this.tx.createDuplicate());
        }
        var oCat = this.cat || this.xVal;
        if(AscCommon.isRealObject(oCat)) {
            oCopy.setCat(oCat.createDuplicate());
        }
        var oVal = this.val || this.yVal;
        if(AscCommon.isRealObject(oVal)) {
            oCopy.setVal(oVal.createDuplicate());
        }
    };
    CSeriesBase.prototype.getSeriesName = function() {
        if(this.tx) {
            if(typeof this.tx.val === "string") {
                return this.tx.val;
            }
            if(this.tx.strRef
                && this.tx.strRef.strCache
                && AscFormat.isRealNumber(this.tx.strRef.strCache.ptCount)
                && this.tx.strRef.strCache.ptCount > 0) {
                if(this.tx.strRef.strCache.pts.length > 0) {
                    return this.tx.getText(false);
                }
                return "";
            }
        }
        return AscCommon.translateManager.getValue('Series') + " " + (this.idx + 1);
    };
    CSeriesBase.prototype.handleUpdateFill = function() {
        this.onChartInternalUpdate();
    };
    CSeriesBase.prototype.handleUpdateLn = function() {
        this.onChartInternalUpdate();
    };
    CSeriesBase.prototype.documentCreateFontMap = function(allFonts) {
        this.dLbls && this.dLbls.documentCreateFontMap(allFonts);
    };
    CSeriesBase.prototype.applyLabelsFunction = function(fCallback, value, oDD) {
        this.dLbls && this.dLbls.applyLabelsFunction(fCallback, value, oDD);
    };
    CSeriesBase.prototype.getAllRasterImages = function(images) {
        this.spPr && this.spPr.checkBlipFillRasterImage(images);
        this.dLbls && this.dLbls.getAllRasterImages(images);
        this.marker && this.marker.spPr && this.marker.spPr.checkBlipFillRasterImage(images);
        if(this.dPt) {
            for(var i = 0; i < this.dPt.length; ++i) {
                this.dPt[i].spPr && this.dPt[i].spPr.checkBlipFillRasterImage(images);
                this.dPt[i].marker && this.dPt[i].marker.spPr && this.dPt[i].marker.spPr.checkBlipFillRasterImage(images);
            }
        }
    };
    CSeriesBase.prototype.checkSpPrRasterImages = function(images) {
        checkSpPrRasterImages(this.spPr);
        checkSpPrRasterImages(this.dLbls);
        if(Array.isArray(this.dPt)) {
            for(var i = 0; i < this.dPt.length; ++i) {
                checkSpPrRasterImages(this.dPt[i].spPr);
                this.dPt[i].marker && checkSpPrRasterImages(this.dPt[i].marker.spPr);
            }
        }
    };
    CSeriesBase.prototype.getValRefFormula = function() {
        var oVal = this.val || this.yVal;
        if(oVal) {
            return oVal.getRefFormula();
        }
        return null;
    };
    CSeriesBase.prototype.getCatName = function(idx) {
        var pts;
        var cat;
        if(this.cat) {
            cat = this.cat;
        }
        else if(this.xVal) {
            cat = this.xVal;
        }

        if(cat) {
            if(cat && cat.strRef && cat.strRef.strCache) {
                pts = cat.strRef.strCache.pts;
            }
            else if(cat.numRef && cat.numRef.numCache) {
                pts = cat.numRef.numCache.pts;
            }
            if(Array.isArray(pts)) {
                for(var i = 0; i < pts.length; ++i) {
                    if(pts[i].idx === idx) {
                        return pts[i].val + "";
                    }
                }
            }
        }
        return (idx + 1) + "";
    };
    CSeriesBase.prototype.getStrPercentageValByIndex = function(idx) {
        var pts = this.getNumPts();
        if(Array.isArray(pts)) {
            var i;
            var summ = 0, value;
            for(i = 0; i < pts.length; ++i) {
                if(AscFormat.isRealNumber(pts[i].val))
                    summ += Math.abs(pts[i].val);

                if(pts[i].idx === idx) {
                    value = pts[i].val;
                }
            }

            if(summ > 0 && AscFormat.isRealNumber(value))
                return Math.round(100 * (value / summ)) + "%";
        }
        return "";
    };
    CSeriesBase.prototype.checkDlblsPosition = function(aPossiblePositions) {
        if(this.dLbls) {
            this.dLbls.checkPosition(aPossiblePositions);
        }
    };
    CSeriesBase.prototype.isFiltered = function() {
        return !this.parent.isVisible(this);
    };
    CSeriesBase.prototype.isVisible = function() {
        if(!this.parent) {
            return true;
        }
        return this.parent.getSeriesArrayIdx(this) > -1;
    };
    CSeriesBase.prototype.setVisible = function(bVal) {
        var oChart = this.parent;
        if(!oChart) {
            return;
        }
        if(bVal) {
            if(this.isVisible()) {
                return;
            }
            else {
                oChart.removeFilteredSeries(oChart.getFilteredSeriesArrayIdx(this));
                oChart.addSer(this);
            }
        }
        else {
            if(this.isVisible()) {
                oChart.removeSeriesInternal(this.getSeriesArrayIdx(this));
                oChart.addFilteredSeries(this);
            }
            else {
                return;
            }
        }
    };
    CSeriesBase.prototype.getOrder = function() {
        return this.order;
    };
    CSeriesBase.prototype.setName = function(sName) {
        var oResult;
        if(typeof sName === "string" && sName.length > 0) {
            var oTx = new CTx();
            oResult = oTx.setValues(sName);
            if(oResult.isSuccessful() && oTx.isValid()) {
                this.setTx(oTx);
            }
            else {
                oResult = new CParseResult();
                this.setTx(null);
            }
        }
        else {
            oResult = new CParseResult();
            this.setTx(null);
        }
        return oResult.getError();
    };
    CSeriesBase.prototype.getName = function() {
        if(this.tx) {
            return this.tx.getFormula();
        }
        return "";
    };
    CSeriesBase.prototype.setValues = function(sValues) {
        var oResult;
        if(sValues === "" || sValues === null) {
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.NoValues);
            return oResult;
        }
        var oVal = new CYVal();
        oResult = oVal.setValues(sValues);
        if(oVal.isValid()) {
            this.setVal(oVal);
        }
        else {
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.DataRangeError);
        }
        return oResult;
    };
    CSeriesBase.prototype.setCategories = function(sValues) {
        if(!this.setCat) {
            return;
        }
        var oResult;
        if(sValues === null || typeof sValues === "string" && sValues.length === 0) {
            if(this.cat) {
                this.setCat(null);
            }
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.No);
        }
        else {
            var oCat = new CCat();
            oResult = oCat.setValues(sValues);
            if(oCat.isValid()) {
                this.setCat(oCat);
            }
            else {
                if(this.cat) {
                    this.setCat(null);
                }
                oResult = new CParseResult();
                oResult.setError(Asc.c_oAscError.ID.DataRangeError);
            }
        }
        return oResult;
    };
    CSeriesBase.prototype.getValues = function(nMaxValues) {
        if(this.cat) {
            return this.cat.getValues(nMaxValues);
        }
        if(this.val) {
            var ret = [];
            for(var nIndex = 0; nIndex < nMaxValues; ++nIndex) {
                ret.push((nIndex + 1) + "");
            }
            return ret;
        }
        return [];
    };
    CSeriesBase.prototype.getCatDataRefs = function() {
        var oCat = this.cat || this.xVal;
        if(oCat) {
            return oCat.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CSeriesBase.prototype.getTxDataRefs = function() {
        if(this.tx) {
            return this.tx.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CSeriesBase.prototype.getValDataRefs = function() {
        var oVal = this.val || this.yVal;
        if(oVal) {
            return oVal.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CSeriesBase.prototype.getErrBarsMinusDataRefs = function() {
        var oVal = this.errBars;
        if(oVal) {
            return oVal.getMinusDataRefs();
        }
        return new CDataRefs([]);
    };
    CSeriesBase.prototype.getErrBarsPlusDataRefs = function() {
        var oVal = this.errBars;
        if(oVal) {
            return oVal.getPlusDataRefs();
        }
        return new CDataRefs([]);
    };
    CSeriesBase.prototype.collectRefs = function(oSource, aRefs) {
        if(oSource) {
            oSource.collectRefs(aRefs);
        }
    };
    CSeriesBase.prototype.getValuesCount = function() {
        if(this.val) {
            return this.val.getValuesCount()
        }
        return 0;
    };
    CSeriesBase.prototype.setXValues = function(sValues) {
        var oResult;
        if(sValues === "" || sValues === null) {
            this.setXVal(null);
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.No);
        }
        else {
            var oCat = new CCat();
            oResult = oCat.setValues(sValues);
            if(oCat.isValid()) {
                this.setXVal(oCat);
            }
            else {
                if(this.xVal) {
                    this.setXVal(null);
                }
                oResult = new CParseResult();
                oResult.setError(Asc.c_oAscError.ID.DataRangeError);
            }
        }
        return oResult;
    };
    CSeriesBase.prototype.setYValues = function(sValues) {
        var oResult;
        if(sValues === "" || sValues === null) {
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.NoValues);
            return oResult;
        }
        var oVal = new CYVal();
        oResult = oVal.setValues(sValues);
        if(oVal.isValid()) {
            this.setYVal(oVal);
        }
        else {
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.ID.DataRangeError);
        }
        return oResult;
    };
    CSeriesBase.prototype.remove = function() {
        var oChart = this.parent;
        if(!oChart) {
            return;
        }
        oChart.removeSeries(oChart.getSeriesArrayIdx(this));
        oChart.reorderSeries();
    };
    CSeriesBase.prototype.setData = function(oSeriesData) {
        var sValues = oSeriesData.val.getFormula();
        var sTx = oSeriesData.tx.getFormula();
        var sCat = oSeriesData.cat.getFormula();
        this.setName(sTx);
        if(this.getObjectType() === AscDFH.historyitem_type_ScatterSer) {
            this.setYValues(sValues);
            this.setXValues(sCat);
        }
        else {
            this.setValues(sValues);
            this.setCategories(sCat);
        }
    };
    CSeriesBase.prototype.fillSelectedRanges = function(oWSView) {
        var oData = new CSeriesDataRefs(this);
        oData.fillSelectedRanges(oWSView);
    };
    CSeriesBase.prototype.fillFromSelectedRange = function(oSelectedRange) {
        var oData = new CSeriesDataRefs(this);
        oData.fillFromSelectedRange(oSelectedRange);
        this.setData(oData);
    };
    CSeriesBase.prototype.setDLblsDeleteValue = function(bVal) {
        if(this.dLbls) {
            this.dLbls.setDeleteValue(bVal)
        }
    };
    CSeriesBase.prototype.getFirstPointFormatCode = function() {
        var sDefaultValAxFormatCode = null;
        var aPoints = this.getNumPts();
        if(aPoints[0] && typeof aPoints[0].formatCode === "string" && aPoints[0].formatCode.length > 0) {
            sDefaultValAxFormatCode = aPoints[0].formatCode;
        }
        return sDefaultValAxFormatCode;
    };
    CSeriesBase.prototype.setDlblsProps = function(oProps) {
        if(!this.parent) {
            return;
        }
        var nPos = oProps.getDataLabelsPos();
        nPos = fCheckDLPostion(nPos, this.parent.getPossibleDLblsPosition());
        if(!this.dLbls) {
            this.setDLbls(new AscFormat.CDLbls());
        }
        this.dLbls.setSettings(nPos, oProps);
        this.dLbls.checkChartStyle();
    };
    CSeriesBase.prototype.getCatSourceNumFormat = function() {
        var oCat = this.cat || this.xVal;
        if(!oCat) {
            return "General";
        }
        return oCat.getSourceNumFormat();
    };
    CSeriesBase.prototype.getValSourceNumFormat = function() {
        var oChart = this.parent;
        if(!oChart) {
            return "General";
        }
        if(oChart.getObjectType() === AscDFH.historyitem_type_BarChart) {
            if(oChart.grouping === AscFormat.BAR_GROUPING_PERCENT_STACKED) {
                return "0%";
            }
        }
        else {
            if(oChart.grouping === AscFormat.GROUPING_PERCENT_STACKED) {
                return "0%";
            }
        }
        return this.getValObjectSourceNumFormat(0);
    };
    CSeriesBase.prototype.getValObjectSourceNumFormat = function(nPtIdx) {
        var oVal = this.val || this.yVal;
        if(!oVal) {
            return "General";
        }
        return oVal.getSourceNumFormat(nPtIdx);
    };
    CSeriesBase.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.val) {
            this.val.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.yVal) {
            this.yVal.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.cat) {
            this.cat.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.xVal) {
            this.xVal.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.tx) {
            this.tx.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.errBars) {
            this.errBars.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CSeriesBase.prototype.fillFromAscSeries = function(oAscSeries, bUseCache) {
        var bIsScatter = this.isScatter();
        var oVal = new AscFormat.CYVal();
        if(bIsScatter) {
            this.setYVal(oVal);
        }
        else {
            this.setVal(oVal);
        }
        oVal.fillFromAsc(oAscSeries.Val, bUseCache);
        if (oAscSeries.FormatCode !== "")
            this.getNumLit().setFormatCode(oAscSeries.FormatCode);
            
        var oAscCat = oAscSeries.Cat;
        if(oAscCat && typeof oAscCat.Formula === "string" && oAscCat.Formula.length > 0) {
            var oCat = new CCat();
            if(bIsScatter) {
                this.setXVal(oCat);
            }
            else {
                this.setCat(oCat);
            }
            oCat.fillFromAsc(oAscCat, bUseCache);
        }
        var oAscTx = oAscSeries.TxCache;
        if(oAscTx && typeof oAscTx.Formula === "string" && oAscTx.Formula.length > 0) {
            this.setTx(new CTx());
            this.tx.fillFromAsc(oAscTx, bUseCache);
        }
    };
    CSeriesBase.prototype.getNumLit = function() {
        if(this.val) {
            if(this.val.numRef && this.val.numRef.numCache)
                return this.val.numRef.numCache;
            else if(this.val.numLit)
                return this.val.numLit;
        }
        else if(this.yVal) {
            if(this.yVal.numRef && this.yVal.numRef.numCache)
                return this.yVal.numRef.numCache;
            else if(this.yVal.numLit)
                return this.yVal.numLit;
        }
        return null;
    };
    CSeriesBase.prototype.getNumPts = function() {
        var oNumLit = this.getNumLit();
        if(oNumLit) {
            return oNumLit.pts
        }
        return [];
    };
    CSeriesBase.prototype.hasNoFillLine = function() {
        if(this.spPr) {
            return this.spPr.hasNoFillLine();
        }
        return false;
    };
    CSeriesBase.prototype.hasNoFill = function() {
        if(this.spPr) {
            return this.spPr.hasNoFill();
        }
        return false;
    };
    CSeriesBase.prototype.checkR1C1ModeForExternal = function(sFormula) {
        var bR1C1Mode = false;
        if(typeof AscCommonExcel === "object" && AscCommonExcel !== null) {
            bR1C1Mode = AscCommonExcel.g_R1C1Mode;
        }
        if(!bR1C1Mode) {
            return sFormula;
        }
        else {
            var aRefs = fParseChartFormula(sFormula);
            if(!Array.isArray(aRefs) || aRefs.length === 0) {
                return sFormula;
            }
            else {
                var oDataRefs = new CDataRefs(aRefs);
                return oDataRefs.getDataRange();
            }
        }
        return sFormula;
    };
    CSeriesBase.prototype.isVaryColors = function() {
        if(!this.parent) {
            return false;
        }
        return this.parent.varyColors === true;
    };
    CSeriesBase.prototype.getDataStyleEntry = function(oChartStyle) {
        return oChartStyle.getDataEntry(this);
    };
    CSeriesBase.prototype.getDptByIdx = function(idx) {
        if(Array.isArray(this.dPt)) {
            for(var i = 0; i < this.dPt.length; ++i) {
                if(this.dPt[i].idx === idx) {
                    return this.dPt[i];
                }
            }
        }
        return null;
    };
    CSeriesBase.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData,  bReset) {
        if(!this.parent) {
            return;
        }

        var bResetLine = false;
        var nType = this.getObjectType();
        if(nType === AscDFH.historyitem_type_ScatterSer) {//TODO: radar chart
            if(this.parent.isNoLine()) {
                bResetLine = true;
            }
        }
        var bMarkerChart = false;
        if(nType === AscDFH.historyitem_type_LineSeries
            || nType === AscDFH.historyitem_type_RadarSeries
            || nType === AscDFH.historyitem_type_ScatterSer) {
            if(this.parent.isMarkerChart()) {
                bMarkerChart = true;
            }
        }
        var nColorsCount;
        var aColors;
        var oDataStyleEntry = this.getDataStyleEntry(oChartStyle);
        var oMarker;
        if(bReset && oAdditionalData) {
            if(!oAdditionalData.dLbls) {
                this.setDLbls(null);
            }
            else {
                this.setDLbls(oAdditionalData.dLbls.createDuplicate());
            }
        }
        if(this.dLbls) {
            this.dLbls.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.errBars) {
            this.errBars.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.trendline) {
            this.trendline.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.isVaryColors() && this.addDPt) {
            var oVal = (this.val || this.yVal);
            nColorsCount = oVal.getValuesCount();
            aColors = oColors.generateColors(nColorsCount);
            var nDPtCount = nColorsCount;
            var nDPt, oDPt;
            if(bReset) {
                this.removeAllDPts();
                for(nDPt = 0; nDPt < nDPtCount; ++nDPt) {
                    oDPt = new CDPt();
                    oDPt.setIdx(nDPt);
                    this.addDPt(oDPt);
                    if(this.marker) {
                        oMarker = this.marker.createDuplicate();
                        oDPt.setMarker(oMarker);
                        oMarker.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
                    }
                    oDPt.applyStyleEntry(oDataStyleEntry, aColors, nDPt, bReset);
                }
            }
            else {
                for(nDPt = 0; nDPt < nDPtCount; ++nDPt) {
                    oDPt = this.getDptByIdx(nDPt);
                    if(!oDPt) {
                        oDPt = new CDPt();
                        oDPt.setIdx(nDPt);
                        this.addDPt(oDPt);
                        if(this.marker) {
                            oMarker = this.marker.createDuplicate();
                            oDPt.setMarker(oMarker);
                            oMarker.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
                        }
                        oDPt.applyStyleEntry(oDataStyleEntry, aColors, nDPt, bReset);
                    }
                }
                if(Array.isArray(this.dPt)) {
                    for(nDPt = this.dPt.length - 1; nDPt > -1 ; --nDPt) {
                        oDPt = this.dPt[nDPt];
                        if(oDPt.idx >= nDPtCount) {
                            this.removeDPt(nDPt)
                        }
                    }
                }
            }
            this.resetOwnFormatting();
        }
        else {
            nColorsCount = this.getMaxSeriesIdx() + 1;
            aColors = oColors.generateColors(nColorsCount);
            this.removeAllDPts();
            this.applyStyleEntry(oDataStyleEntry, aColors, this.idx, bReset);
        }

        if(bMarkerChart) {
            this.setMarker(new CMarker());
            if(this.marker) {
                this.marker.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
            }
        }
        if(bResetLine) {
            if(!this.spPr) {
                this.setSpPr(new AscFormat.CSpPr());
            }
            this.spPr.setLn(AscFormat.CreateNoFillLine());
        }

		if(nType === AscDFH.historyitem_type_RadarSeries) {
			if(this.parent.radarStyle === AscFormat.RADAR_STYLE_FILLED) {
				if(this.spPr && this.spPr.hasNoFill()) {
					this.spPr.setFill(null);
				}
			}
		}
    };
    CSeriesBase.prototype.getMaxSeriesIdx = function() {
        if(!this.parent) {
            return -1;
        }
        return this.parent.getMaxSeriesIdx();
    };
    CSeriesBase.prototype.removeAllDPts = function() {
        if(Array.isArray(this.dPt)) {
            for(var nDPt = this.dPt.length - 1; nDPt > -1; --nDPt) {
                this.removeDPt(nDPt);
            }
        }
    };
    CSeriesBase.prototype.removeDPt = function(idx) {
        if(!Array.isArray(this.dPt)) {
            return;
        }
        if(this.dPt[idx]) {
            var arrPt = this.dPt.splice(idx, 1);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonSeries_RemoveDPt, idx, arrPt, false));
            arrPt[0].setParent(null);
        }
    };
    CSeriesBase.prototype.getErrBars = function() {
        return this.errBars || null;
    };
	CSeriesBase.prototype.checkSeriesAfterChangeType = function() {

	};
    CSeriesBase.prototype.asc_getName = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function() {
                return oThis.checkR1C1ModeForExternal(oThis.getName());
            }, this, []);
    };
    CSeriesBase.prototype["asc_getName"] = CSeriesBase.prototype.asc_getName;
    CSeriesBase.prototype.asc_getNameVal = function() {
        return AscFormat.ExecuteNoHistory(function() {
            if(this.tx) {
                return this.tx.getText(false);
            }
            return "";
        }, this, []);
    };
    CSeriesBase.prototype["asc_getNameVal"] = CSeriesBase.prototype.asc_getNameVal;
    CSeriesBase.prototype.asc_getSeriesName = function() {
        return this.getSeriesName();
    };
    CSeriesBase.prototype["asc_getSeriesName"] = CSeriesBase.prototype.asc_getSeriesName;
    CSeriesBase.prototype.asc_setName = function(sName) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        this.setName(sName);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_setName"] = CSeriesBase.prototype.asc_setName;
    CSeriesBase.prototype.asc_IsValidName = function(sName) {
        return AscFormat.ExecuteNoHistory(function() {
            if(sName === "" || sName === null) {
                return Asc.c_oAscError.ID.No;
            }
            var oTx = new CTx();
            return oTx.setValues(sName).getError();
        }, this, [])
    };
    CSeriesBase.prototype["asc_IsValidName"] = CSeriesBase.prototype.asc_IsValidName;
    CSeriesBase.prototype.asc_setValues = function(sValues) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        this.setValues(sValues);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_setValues"] = CSeriesBase.prototype.asc_setValues;
    CSeriesBase.prototype.asc_IsValidValues = function(sValues) {
        return AscFormat.ExecuteNoHistory(function() {
            if(sValues === "" || sValues === null) {
                return Asc.c_oAscError.ID.NoValues;
            }
            var oVal = new CYVal();
            return oVal.setValues(sValues).getError();
        }, this, []);
    };
    CSeriesBase.prototype["asc_IsValidValues"] = CSeriesBase.prototype.asc_IsValidValues;
    CSeriesBase.prototype.asc_getValues = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function() {
            return oThis.checkR1C1ModeForExternal(oThis.val.getFormula());
        }, this, []);
    };
    CSeriesBase.prototype["asc_getValues"] = CSeriesBase.prototype.asc_getValues;
    CSeriesBase.prototype.asc_getValuesArr = function() {
        return AscFormat.ExecuteNoHistory(function() {
            return this.val.getValues(this.val.getValuesCount());
        }, this, []);
    };
    CSeriesBase.prototype["asc_getValuesArr"] = CSeriesBase.prototype.asc_getValuesArr;
    CSeriesBase.prototype.asc_setXValues = function(sValues) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        this.setXValues(sValues);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_setXValues"] = CSeriesBase.prototype.asc_setXValues;
    CSeriesBase.prototype.asc_IsValidXValues = function(sValues) {
        return AscFormat.ExecuteNoHistory(function() {
            if(sValues === "" || sValues === null) {
                return Asc.c_oAscError.ID.No;
            }
            else {
                var oCat = new CCat();
                return oCat.setValues(sValues).getError();
            }
        }, this, []);
    };
    CSeriesBase.prototype["asc_IsValidXValues"] = CSeriesBase.prototype.asc_IsValidXValues;
    CSeriesBase.prototype.asc_getCatValues = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function() {
            if(oThis.cat) {
                return oThis.checkR1C1ModeForExternal(oThis.cat.getFormula());
            }
            else {
                return "";
            }
        }, this, []);
    };
    CSeriesBase.prototype["asc_getCatValues"] = CSeriesBase.prototype.asc_getCatValues;
    CSeriesBase.prototype.asc_getXValues = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function() {
            if(oThis.xVal) {
                return oThis.checkR1C1ModeForExternal(oThis.xVal.getFormula());
            }
            else {
                return "";
            }
        }, this, []);
    };
    CSeriesBase.prototype["asc_getXValues"] = CSeriesBase.prototype.asc_getXValues;
    CSeriesBase.prototype.asc_getXValuesArr = function() {
        return AscFormat.ExecuteNoHistory(function() {
            if(this.xVal) {
                return this.xVal.getValues();
            }
            return [];
        }, this, []);
    };
    CSeriesBase.prototype["asc_getXValuesArr"] = CSeriesBase.prototype.asc_getXValuesArr;
    CSeriesBase.prototype.asc_setYValues = function(sValues) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        this.setYValues(sValues);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_setYValues"] = CSeriesBase.prototype.asc_setYValues;
    CSeriesBase.prototype.asc_IsValidYValues = function(sValues) {
        return this.asc_IsValidValues(sValues);
    };
    CSeriesBase.prototype["asc_IsValidYValues"] = CSeriesBase.prototype.asc_IsValidYValues;
    CSeriesBase.prototype.asc_getYValues = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function() {
            if(oThis.yVal) {
                return oThis.checkR1C1ModeForExternal(oThis.yVal.getFormula());
            }
            else {
                return "";
            }
        }, this, []);
    };
    CSeriesBase.prototype["asc_getYValues"] = CSeriesBase.prototype.asc_getYValues;
    CSeriesBase.prototype.asc_getYValuesArr = function() {
        return AscFormat.ExecuteNoHistory(function() {
            if(this.yVal) {
                return this.yVal.getValues();
            }
            return [];
        }, this, []);
    };
    CSeriesBase.prototype["asc_getYValuesArr"] = CSeriesBase.prototype.asc_getYValuesArr;
    CSeriesBase.prototype.asc_Remove = function() {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        History.Create_NewPoint(0);
        this.remove();
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_Remove"] = CSeriesBase.prototype.asc_Remove;
    CSeriesBase.prototype.isScatter = function() {
        return this.getObjectType() === AscDFH.historyitem_type_ScatterSer;
    };
    CSeriesBase.prototype.asc_IsScatter = function() {
        return this.isScatter();
    };
    CSeriesBase.prototype["asc_IsScatter"] = CSeriesBase.prototype.asc_IsScatter;
    CSeriesBase.prototype.asc_getOrder = function() {
        return this.getOrder();
    };
    CSeriesBase.prototype["asc_getOrder"] = CSeriesBase.prototype.asc_getOrder;
    CSeriesBase.prototype.asc_setOrder = function(nVal) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        History.Create_NewPoint(0);
        this.setOrder(nVal);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_setOrder"] = CSeriesBase.prototype.asc_setOrder;
    CSeriesBase.prototype.asc_getIdx = function() {
        return this.idx;
    };
    CSeriesBase.prototype["asc_getIdx"] = CSeriesBase.prototype.asc_getIdx;
    CSeriesBase.prototype.asc_MoveUp = function() {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        if(!this.parent) {
            return;
        }
        History.Create_NewPoint(0);
        this.parent.moveSeriesUp(this);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_MoveUp"] = CSeriesBase.prototype.asc_MoveUp;
    CSeriesBase.prototype.asc_MoveDown = function() {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        if(!this.parent) {
            return;
        }
        History.Create_NewPoint(0);
        this.parent.moveSeriesDown(this);
        oChartSpace.onDataUpdate();
    };
    CSeriesBase.prototype["asc_MoveDown"] = CSeriesBase.prototype.asc_MoveDown;
    CSeriesBase.prototype.onDataUpdate = function() {
        if(!this.parent) {
            return;
        }
        this.parent.onDataUpdate();
    };
    CSeriesBase.prototype.asc_getChartType = function() {
        if(this.parent) {
            return this.parent.getChartType();
        }
        return Asc.c_oAscChartTypeSettings.unknown;
    };
    CSeriesBase.prototype["asc_getChartType"] = CSeriesBase.prototype.asc_getChartType;
    CSeriesBase.prototype.getPreviewBrush = function() {
        if(this.compiledSeriesBrush) {
            return this.compiledSeriesBrush;
        }
        return null;
    };
    CSeriesBase.prototype.drawPreviewRect = function(sDivId) {
        var oCanvas = AscCommon.checkCanvasInDiv(sDivId);
        if(!oCanvas) {
            return;
        }
        var oContext = oCanvas.getContext("2d");
        if(!oContext) {
            return;
        }
        var dMMW = oCanvas.width / 96 * 25.4;
        var dMMH = oCanvas.height / 96 * 25.4;
        oContext.clearRect(0, 0, oCanvas.width, oCanvas.height);
        var oGraphics = new AscCommon.CGraphics();
        oGraphics.init(oContext, oCanvas.width, oCanvas.height, dMMW, dMMH);
        oGraphics.m_oFontManager = AscCommon.g_fontManager;
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.SaveGrState();
        oGraphics.SetIntegerGrid(false);
        var ShapeDrawer = new AscCommon.CShapeDrawer();
        ShapeDrawer.fromShape2(new AscFormat.ObjectToDraw(this.getPreviewBrush(), null, dMMW, dMMH, null, new AscCommon.CMatrix()), oGraphics, null);
        ShapeDrawer.draw(null);
        oGraphics.RestoreGrState();
    };
    CSeriesBase.prototype.asc_drawPreviewRect = function(sDivId) {
        this.drawPreviewRect(sDivId);
    };
    CSeriesBase.prototype["asc_drawPreviewRect"] = CSeriesBase.prototype.asc_drawPreviewRect;
    CSeriesBase.prototype.isSecondaryAxis = function() {
        if(this.parent) {
            return this.parent.isSecondaryAxis();
        }
        return false;
    };
    CSeriesBase.prototype.asc_getIsSecondaryAxis = function() {
        return this.isSecondaryAxis();
    };
    CSeriesBase.prototype["asc_getIsSecondaryAxis"] = CSeriesBase.prototype.asc_getIsSecondaryAxis;
    CSeriesBase.prototype.canChangeAxisType = function() {
        if(!this.parent) {
            return false;
        }
        return this.parent.canChangeAxisType();
    };
    CSeriesBase.prototype.asc_canChangeAxisType = function() {
        return this.canChangeAxisType();
    };
    CSeriesBase.prototype["asc_canChangeAxisType"] = CSeriesBase.prototype.asc_canChangeAxisType;
    CSeriesBase.prototype.tryChangeSeriesAxisType = function(bIsSecondary) {
        if(!this.parent) {
            return Asc.c_oAscError.ID.No;
        }
        if(!this.canChangeAxisType()) {
            return Asc.c_oAscError.ID.No;
        }
        return this.parent.tryChangeSeriesAxisType(this, bIsSecondary);
    };
    CSeriesBase.prototype.asc_TryChangeAxisType = function(bIsSecondary) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        History.Create_NewPoint(0);
        var nRes = this.tryChangeSeriesAxisType(bIsSecondary);
        if(nRes === Asc.c_oAscError.ID.No) {
            oChartSpace.onDataUpdate();
        }
        return nRes;
    };
    CSeriesBase.prototype["asc_TryChangeAxisType"] = CSeriesBase.prototype.asc_TryChangeAxisType;
    CSeriesBase.prototype.tryChangeChartType = function(nType) {
        if(!this.parent) {
            return Asc.c_oAscError.ID.No;
        }
        return this.parent.tryChangeSeriesChartType(this, nType);
    };
    CSeriesBase.prototype.asc_TryChangeChartType = function(nType) {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        History.Create_NewPoint(0);
        var nRes = this.tryChangeChartType(nType);
        if(nRes === Asc.c_oAscError.ID.No) {
            oChartSpace.onDataUpdate();
        }
        return nRes;
    };
    CSeriesBase.prototype["asc_TryChangeChartType"] = CSeriesBase.prototype.asc_TryChangeChartType;

    function CPlotArea() {
        CBaseChartObject.call(this);
        this.charts = [];
        this.dTable = null;
        this.layout = null;
        this.spPr = null;
        this.axId = [];

        //ТоDo
        this.valAx = null;
        this.catAx = null;
        this.serAx = null;
        this.dateAx = null;
        this.chart = null;
        //

        this.posX = 0;
        this.posY = 0;
        this.x = 0;
        this.y = 0;
        this.rot = 0;
        this.extX = 5;
        this.extY = 5;

        this.localTransform = new AscCommon.CMatrix();
        this.transform = new AscCommon.CMatrix();
        this.invertTransform = new AscCommon.CMatrix();
    }

    InitClass(CPlotArea, CBaseChartObject, AscDFH.historyitem_type_PlotArea);
    CPlotArea.prototype.Refresh_RecalcData = function(data) {
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_RemoveAxis:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_RemoveChart:
            case AscDFH.historyitem_PlotArea_AddChart:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_PlotArea_SetCatAx:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetDateAx:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetDTable:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetLayout:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetSerAx:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetSpPr:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_SetValAx:
            {
                break;
            }
            case AscDFH.historyitem_PlotArea_AddAxis:
            {
                break;
            }
        }
    };
    CPlotArea.prototype.Refresh_RecalcData2 = function(data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(data);
        }
    };
    CPlotArea.prototype.checkShapeChildTransform = function(t) {
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.MultiplyAppend(this.transform, t);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
    };
    CPlotArea.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, x, y);

        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
    };
    CPlotArea.prototype.getChildren = function() {
        var aRet = [this.dTable, this.layout, this.spPr];
        for(var nAx = 0; nAx < this.axId.length; ++nAx) {
            aRet.push(this.axId[nAx]);
        }
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            aRet.push(this.charts[nChart]);
        }
        return aRet;
    };
    CPlotArea.prototype.fillObject = function(oCopy, oIdMap) {
        var i, j, k;
        if(this.dTable) {
            oCopy.setDTable(this.dTable.createDuplicate());
        }
        if(this.layout) {
            oCopy.setLayout(this.layout.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }

        var len = this.axId.length;
        for(i = 0; i < len; i++) {
            var oAxis = this.axId[i].createDuplicate();
            oAxis.setAxId(++AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
            oCopy.addAxis(oAxis);
        }
        var cur_chart, cur_axis;
        for(i = 0; i < this.charts.length; ++i) {
            cur_chart = this.charts[i];
            oCopy.addChart(cur_chart.createDuplicate(), oCopy.charts.length);
            if(Array.isArray(cur_chart.axId)) {
                for(j = 0; j < cur_chart.axId.length; ++j) {
                    cur_axis = cur_chart.axId[j];
                    for(k = 0; k < this.axId.length; ++k) {
                        if(cur_axis === this.axId[k]) {
                            oCopy.charts[i].addAxId(oCopy.axId[k]);
                        }
                    }
                }
            }
        }

        //выставим пересечения осей в копии

        for(i = 0; i < this.axId.length; ++i) {
            cur_axis = this.axId[i];
            for(j = 0; j < this.axId.length; ++j) {
                if(cur_axis.crossAx === this.axId[j]) {
                    oCopy.axId[i].setCrossAx(oCopy.axId[j]);
                    break;
                }
            }
        }
    };
    CPlotArea.prototype.getSeriesWithSmallestIndexForAxis = function(oAxis) {
        var aCharts = this.charts;
        var oRet = null;
        var oChart, aSeries;
        for(var i = 0; i < aCharts.length; ++i) {
            oChart = aCharts[i];
            var aAxes = aCharts[i].axId;
            if(!aAxes) {
                continue;
            }
            for(var j = 0; j < aAxes.length; ++j) {
                if(aAxes[j] === oAxis) {
                    aSeries = oChart.series;
                    for(var k = 0; k < aSeries.length; ++k) {
                        if(oRet === null
                            || aSeries[k].cat && !oRet.cat
                            || aSeries[k].idx < oRet.idx && aSeries[k].cat && oRet.cat) {
                            oRet = aSeries[k];
                        }
                    }
                }
            }
        }
        return oRet;
    };
    CPlotArea.prototype.getChartsForAxis = function(oAxis) {
        var aCharts = this.charts;
        var oChart;
        var aRet = [];
        for(var i = 0; i < aCharts.length; ++i) {
            oChart = aCharts[i];
            var aAxes = aCharts[i].axId;
            if(!aAxes) {
                continue;
            }
            for(var j = 0; j < aAxes.length; ++j) {
                if(aAxes[j] === oAxis) {
                    aRet.push(oChart);
                    break;
                }
            }
        }
        return aRet;
    };
    CPlotArea.prototype.getChartWithSuitableAxes = function(nType, bIsSecondary) {
        var aCharts = this.charts;
        var nChart, oChart, oAxes;
        var nCurChartType;
        for(nChart = 0; nChart < aCharts.length; ++nChart) {
            oChart = aCharts[nChart];
            if(oChart.isSecondaryAxis() === bIsSecondary) {
                nCurChartType = oChart.getChartType();
                if(this.isPieType(nCurChartType) || this.isDoughnutType(nCurChartType)) {
                    continue;
                }
                if(this.isHBarType(nCurChartType)) {
                    if(this.isHBarType(nType)) {
                        return oChart;
                    }
                    else {
                        return null;
                    }
                }
                else {
                    if(this.isHBarType(nType)) {
                        return null;
                    }
	                if(this.isRadarChart(nType)) {
		                if(this.isRadarChart(nCurChartType)) {
			                return oChart;
		                }
		                else {
			                return null;
		                }
	                }
	                if(this.isRadarChart(nCurChartType)) {
		                if(!this.isRadarChart(nType)) {
			                return null;
		                }
	                }
                    if(this.isScatterType(nCurChartType)) {
                        if(this.isScatterType(nType)) {
                            return oChart;
                        }
                        else {
                            return null;
                        }
                    }
                    oAxes = oChart.getAxisByTypes();
                    if(oAxes.valAx.length > 0 && (oAxes.catAx.length > 0 || oAxes.dateAx.length > 0)) {
                        return oChart;
                    }
                    else {
                        return null;
                    }
                }
            }
        }
        return null;
    };
    CPlotArea.prototype.addAxis = function(axis) {
        //сначала проверим не лежит ли ось уже в plotArea
        if(!axis)
            return;
        var i;
        for(i = 0; i < this.axId.length; ++i) {
            if(this.axId[i] === axis)
                return;
        }
        //если такой оси нет, можно добавлять.
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotArea_AddAxis, this.axId.length, [axis], true));
        this.axId.push(axis);
        this.setParentToChild(axis);
    };
    CPlotArea.prototype.addChart = function(pr, idx) {
        var pos;
        if(AscFormat.isRealNumber(idx))
            pos = idx;
        else
            pos = this.charts.length;
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotArea_AddChart, pos, [pr], true));
        this.charts.splice(pos, 0, pr);
        this.setParentToChild(pr);
        this.onChartUpdateType();
        this.onChangeDataRefs();
    };
    CPlotArea.prototype.removeChartByPos = function(pos) {
        if(this.charts[pos]) {
            var chart = this.charts.splice(pos, 1)[0];
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotArea_RemoveChart, pos, [chart], false));
            this.onChangeDataRefs();
            //удалим все оси этой диаграммы, проверив прежде нет ли ссылок на данные оси в других диаграммах
            if(Array.isArray(chart.axId)) {
                var chart_axis = chart.axId;
                for(var i = 0; i < chart_axis.length; ++i) {
                    var axis = chart_axis[i];
                    for(var j = 0; j < this.charts.length; ++j) {
                        var other_chart = this.charts[j];
                        if(Array.isArray(other_chart.axId)) {
                            for(var k = 0; k < other_chart.axId.length; ++k) {
                                if(other_chart.axId[k] === axis)
                                    break;
                            }
                            if(k < other_chart.axId.length)
                                break;
                        }
                    }
                    if(j === this.charts.length)
                        this.removeAxis(axis);
                }
            }
            chart.setParent(null);
        }
    };
    CPlotArea.prototype.removeChart = function(oChart) {
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            if(this.charts[nChart] === oChart) {
                this.removeChartByPos(nChart);
                return;
            }
        }
    };
    CPlotArea.prototype.removeCharts = function(startPos, endPos) {
        for(var i = endPos; i >= startPos; --i) {
            this.removeChartByPos(i);
        }
    };
    CPlotArea.prototype.removeAxisByPos = function(nPos) {
        if(nPos > -1 && nPos < this.axId.length) {
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PlotArea_RemoveAxis, nPos, [this.axId[nPos]], false));
            this.axId.splice(nPos, 1)[0].setParent(null);
        }
    };
    CPlotArea.prototype.removeAxis = function(axis) {
        for(var i = this.axId.length - 1; i > -1; --i) {
            if(this.axId[i] === axis) {
                this.removeAxisByPos(i);
            }
        }
    };
    CPlotArea.prototype.setDTable = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PlotArea_SetDTable, this.dTable, pr));
        this.dTable = pr;
        this.setParentToChild(pr);
    };
    CPlotArea.prototype.setLayout = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PlotArea_SetLayout, this.layout, pr));
        this.layout = pr;
        this.setParentToChild(pr);
    };
    CPlotArea.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PlotArea_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CPlotArea.prototype.getHorizontalAxis = function() {
        if(this.charts[0]) {
            var aAxes = this.charts[0].axId;
            if(aAxes) {
                for(var i = 0; i < aAxes.length; ++i) {
                    if(aAxes[i].axPos === AX_POS_B || aAxes[i].axPos === AX_POS_T)
                        return aAxes[i];
                }
            }
        }
        return null;
    };
    CPlotArea.prototype.getVerticalAxis = function() {
        if(this.charts[0]) {
            var aAxes = this.charts[0].axId;
            if(aAxes) {
                for(var i = 0; i < aAxes.length; ++i) {
                    if(aAxes[i].axPos === AX_POS_L || aAxes[i].axPos === AX_POS_R)
                        return aAxes[i];
                }
            }
        }
        return null;
    };
    CPlotArea.prototype.getAxisByTypes = function() {
        var ret = {valAx: [], catAx: [], dateAx: [], serAx: []};
        for(var i = 0; i < this.axId.length; ++i) {
            var axis = this.axId[i];
            switch(axis.getObjectType()) {
                case AscDFH.historyitem_type_CatAx:
                case AscDFH.historyitem_type_DateAx:
                {
                    ret.catAx.push(axis);
                    break;
                }
                case AscDFH.historyitem_type_ValAx:
                {
                    ret.valAx.push(axis);
                    break;
                }
                case AscDFH.historyitem_type_SerAx:
                {
                    ret.serAx.push(axis);
                    break;
                }
            }
        }
        return ret;
    };
    CPlotArea.prototype.handleUpdateFill = function() {
        if(this.parent && this.parent.handleUpdateFill) {
            this.parent.handleUpdateFill();
        }
    };
    CPlotArea.prototype.handleUpdateLn = function() {
        if(this.parent && this.parent.handleUpdateLn) {
            this.parent.handleUpdateLn();
        }
    };
    CPlotArea.prototype.getCompiledLine = CShape.prototype.getCompiledLine;
    CPlotArea.prototype.getCompiledFill = CShape.prototype.getCompiledFill;
    CPlotArea.prototype.getCompiledTransparent = CShape.prototype.getCompiledTransparent;
    CPlotArea.prototype.check_bounds = CShape.prototype.check_bounds;
    CPlotArea.prototype.getCardDirectionByNum = CShape.prototype.getCardDirectionByNum;
    CPlotArea.prototype.getNumByCardDirection = CShape.prototype.getNumByCardDirection;
    CPlotArea.prototype.getResizeCoefficients = CShape.prototype.getResizeCoefficients;
    CPlotArea.prototype.getInvertTransform = CShape.prototype.getInvertTransform;
    CPlotArea.prototype.getTransformMatrix = CShape.prototype.getTransformMatrix;
    CPlotArea.prototype.hitToHandles = CShape.prototype.hitToHandles;
    CPlotArea.prototype.getFullFlipH = CShape.prototype.getFullFlipH;
    CPlotArea.prototype.getFullFlipV = CShape.prototype.getFullFlipV;
    CPlotArea.prototype.getAspect = CShape.prototype.getAspect;
    CPlotArea.prototype.getGeom = CShape.prototype.getGeom;
    CPlotArea.prototype.convertPixToMM = function(pix) {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.convertPixToMM(pix);
        }
        return 1;
    };
    CPlotArea.prototype.hitInBoundingRect = CShape.prototype.hitInBoundingRect;
    CPlotArea.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
    CPlotArea.prototype.hitInPath = CShape.prototype.hitInPath;
    CPlotArea.prototype.checkHitToBounds = function(x, y) {
        CDLbl.prototype.checkHitToBounds.call(this, x, y);
    };
    CPlotArea.prototype.getCanvasContext = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getCanvasContext();
        }
        return null;
    };
    CPlotArea.prototype.canRotate = function() {
        return false;
    };
    CPlotArea.prototype.getNoChangeAspect = function() {
        return false;
    };
    CPlotArea.prototype.isChart = function() {
        return false;
    };
    CPlotArea.prototype.canMove = function() {
        return true;
    };
    CPlotArea.prototype.reindexSeries = function() {
        if(this.parent) {
            this.parent.reindexSeries();
        }
    };
    CPlotArea.prototype.reorderSeries = function() {
        if(this.parent) {
            this.parent.reorderSeries();
        }
    };
    CPlotArea.prototype.moveSeriesUp = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesUp(oSeries);
        }
    };
    CPlotArea.prototype.moveSeriesDown = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesDown(oSeries);
        }
    };
    CPlotArea.prototype.onDataUpdate = function() {
        if(this.parent) {
            this.parent.onDataUpdate();
        }
    };
    CPlotArea.prototype.setDLblsDeleteValue = function(bVal) {
        var nChart, nChartsCount = this.charts.length;
        for(nChart = 0; nChart < nChartsCount; ++nChart) {
            this.charts[nChart].setDLblsDeleteValue(bVal);
        }
    };
    CPlotArea.prototype.getChartType = function() {
        if(this.charts.length > 0) {
            if(this.charts.length > 1) {
                if(this.charts.length === 2) {
                    var oFirstChart = this.charts[0];
                    var oSecondChart = this.charts[1];
                    if(oFirstChart.getChartType() === Asc.c_oAscChartTypeSettings.barNormal
                        && oSecondChart.getChartType() === Asc.c_oAscChartTypeSettings.lineNormal) {
                        if(!oFirstChart.isSecondaryAxis()) {
                            if(!oSecondChart.isSecondaryAxis()) {
                                return Asc.c_oAscChartTypeSettings.comboBarLine;
                            }
                            else {
                                return Asc.c_oAscChartTypeSettings.comboBarLineSecondary;
                            }
                        }
                    }
                    else if(oFirstChart.getChartType() === Asc.c_oAscChartTypeSettings.areaNormal
                        && oSecondChart.getChartType() === Asc.c_oAscChartTypeSettings.barNormal) {
                        if(!oFirstChart.isSecondaryAxis()) {
                            if(!oSecondChart.isSecondaryAxis()) {
                                return Asc.c_oAscChartTypeSettings.comboAreaBar;
                            }
                        }
                    }
                }
                return Asc.c_oAscChartTypeSettings.comboCustom;
            }
            return this.charts[0].getChartType();
        }
        return Asc.c_oAscChartTypeSettings.unknown;
    };
    CPlotArea.prototype.addChartWithAxes = function(oTypedChart) {
        this.removeAllCharts();
        this.removeAllAxes();
        this.addChart(oTypedChart, 0);
        var aAxes = oTypedChart.axId;
        if(Array.isArray(aAxes)) {
            for(var nAx = 0; nAx < aAxes.length; ++nAx) {
                var oAxis = aAxes[nAx];
                this.addAxis(oAxis);
            }
        }
    };
    CPlotArea.prototype.getAxisNumFormatByType = function(nType, aSeries) {
        var oCT = Asc.c_oAscChartTypeSettings;
        var sNewNumFormat;
        if(nType === oCT.barStackedPer
            || nType === oCT.barStackedPer3d
            || nType === oCT.hBarStackedPer
            || nType === oCT.hBarStackedPer3d
            || nType === oCT.lineStackedPer
            || nType === oCT.lineStackedPerMarker
            || nType === oCT.areaStackedPer) {
            sNewNumFormat = "0%";
        }
        else {
            sNewNumFormat = aSeries[0] && aSeries[0].getFirstPointFormatCode() || "General"
        }
        return sNewNumFormat;
    };
    CPlotArea.prototype.getBarGroupingByType = function(nType) {
        var oCT = Asc.c_oAscChartTypeSettings;
        var nGrouping = null;
        switch(nType) {
            case oCT.barNormal:
            case oCT.hBarNormal:
            case oCT.barNormal3d:
            case oCT.hBarNormal3d:
            {
                nGrouping = AscFormat.BAR_GROUPING_CLUSTERED;
                break;
            }
            case oCT.barStacked:
            case oCT.hBarStacked:
            case oCT.barStacked3d:
            case oCT.hBarStacked3d:
            {
                nGrouping = AscFormat.BAR_GROUPING_STACKED;
                break;
            }
            case oCT.barNormal3dPerspective:
            {
                nGrouping = AscFormat.BAR_GROUPING_STANDARD;
                break;
            }
            case oCT.hBarStackedPer3d:
            case oCT.hBarStackedPer:
            case oCT.barStackedPer3d:
            case oCT.barStackedPer:
            {
                nGrouping = AscFormat.BAR_GROUPING_PERCENT_STACKED;
                break;
            }
        }
        return nGrouping;
    };
    CPlotArea.prototype.getBarDirByType = function(nType) {
        if(this.isBarType(nType)) {
            return AscFormat.BAR_DIR_COL;
        }
        else if(this.isHBarType(nType)) {
            return AscFormat.BAR_DIR_BAR;
        }
    };
    CPlotArea.prototype.getIs3dByType = function(nType) {
        var oCT = Asc.c_oAscChartTypeSettings;
        return nType === oCT.barNormal3d ||
            nType === oCT.barStacked3d ||
            nType === oCT.barStackedPer3d ||
            nType === oCT.barNormal3dPerspective ||
            nType === oCT.hBarNormal3d ||
            nType === oCT.hBarStacked3d ||
            nType === oCT.hBarStackedPer3d ||
            nType === oCT.line3d ||
            nType === oCT.pie3d;
    };
    CPlotArea.prototype.getIsPerspective = function(nType) {
        var oCT = Asc.c_oAscChartTypeSettings;
        return oCT.barNormal3dPerspective === nType;
    };
    CPlotArea.prototype.check3DOptions = function(nType) {
        this.parent.check3DOptions(this.getIs3dByType(nType), this.getIsPerspective(nType));
    };
    CPlotArea.prototype.createBarChart = function(nType, aSeries, aAxes, oOldChart) {
        var nNewGrouping;
        var isNew3D;
        var nSeries, oSeries;
        var nNewBarDir;
        var oBarChart = new AscFormat.CBarChart();
        nNewGrouping = this.getBarGroupingByType(nType);
        nNewBarDir = this.getBarDirByType(nType);
        oBarChart.mergeWithoutSeries(oOldChart);
        oBarChart.setVaryColors(false);
        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            oSeries = new AscFormat.CBarSeries();
            aSeries[nSeries].fillObject(oSeries);
            oBarChart.addSer(oSeries);
	        oSeries.checkSeriesAfterChangeType();
        }
        oBarChart.setBarDir(nNewBarDir);
        oBarChart.setGrouping(nNewGrouping);
        oBarChart.setGapWidth(150);
        if(AscFormat.BAR_GROUPING_PERCENT_STACKED === nNewGrouping
            || AscFormat.BAR_GROUPING_STACKED === nNewGrouping) {
            oBarChart.setOverlap(100);
        }
        isNew3D = this.getIs3dByType(nType);
        this.parent.check3DOptions(isNew3D, this.getIsPerspective(nType));
        if(isNew3D) {
            if(oBarChart.b3D !== true) {
                oBarChart.set3D(true);
            }
        }
        else {
            if(oBarChart.b3D !== false) {
                oBarChart.set3D(false);
            }
        }
        oBarChart.addAxes(aAxes);
        return oBarChart;

    };
    CPlotArea.prototype.canChangeToComboChart = function() {
        var aSeries = this.getAllSeries();
        if(aSeries.length < 2) {
            return false;
        }
        return true;
    }
    CPlotArea.prototype.switchToCombo = function(nType) {
        if(this.charts.length < 1) {
            return;
        }
        var nCurType = this.getChartType();
        if(nCurType === nType) {
            return;
        }
        //TODO: Use type
        if(!this.canChangeToComboChart()) {
            return;
        }
        var aSeries = this.getAllSeries();
        var oTypedChart;
        var nAx, oAxis, aAllAxes, aFirstAxes, aSecondAxes, aFirstChartSeries, aSecondChartSeries;
        var aAllCharts = [], nChart;
        var oBarChart, oLineChart, oAreaChart;
        var nLength = (aSeries.length / 2 + 0.5) >> 0;
        aFirstChartSeries = aSeries.slice(0, nLength);
        aSecondChartSeries = aSeries.slice(nLength);
        if(nType === Asc.c_oAscChartTypeSettings.comboBarLine
            || nType === Asc.c_oAscChartTypeSettings.comboBarLineSecondary
            || nType === Asc.c_oAscChartTypeSettings.comboCustom) {
            aAllAxes = aFirstAxes = aSecondAxes = this.createRegularAxes(this.getAxisNumFormatByType(Asc.c_oAscChartTypeSettings.barNormal, aFirstChartSeries), false);
            if(nType === Asc.c_oAscChartTypeSettings.comboBarLineSecondary) {
                aSecondAxes = this.createRegularAxes(this.getAxisNumFormatByType(Asc.c_oAscChartTypeSettings.lineNormal, aSecondChartSeries), true);
                aAllAxes = aAllAxes.concat(aSecondAxes);
            }
            oTypedChart = this.charts[0];
            oBarChart = this.createBarChart(Asc.c_oAscChartTypeSettings.barNormal, aFirstChartSeries, aFirstAxes, oTypedChart);
            oLineChart = this.createLineChart(Asc.c_oAscChartTypeSettings.lineNormal, aSecondChartSeries, aSecondAxes, oTypedChart);
            aAllCharts.push(oBarChart);
            aAllCharts.push(oLineChart);
        }
        else if(nType === Asc.c_oAscChartTypeSettings.comboAreaBar) {
            aAllAxes = aFirstAxes = aSecondAxes = this.createRegularAxes(this.getAxisNumFormatByType(Asc.c_oAscChartTypeSettings.barNormal, aFirstChartSeries), false);
            oTypedChart = this.charts[0];
            oAreaChart = this.createAreaChart(Asc.c_oAscChartTypeSettings.areaStacked, aFirstChartSeries, aFirstAxes, oTypedChart);
            oBarChart = this.createBarChart(Asc.c_oAscChartTypeSettings.barNormal, aSecondChartSeries, aSecondAxes, oTypedChart);
            aAllCharts.push(oAreaChart);
            aAllCharts.push(oBarChart);
        }
        if(aAllCharts.length > 0) {
            this.removeAllCharts();
            this.removeAllAxes();
            for(nChart = 0; nChart < aAllCharts.length; ++nChart) {
                this.addChart(aAllCharts[nChart], nChart);
            }
            for(nAx = 0; nAx < aAllAxes.length; ++nAx) {
                oAxis = aAllAxes[nAx];
                this.addAxis(oAxis);
            }
        }
    };
    CPlotArea.prototype.switchToBarChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var aAxes;
        var oBarChart;
        if(this.getBarDirByType(nType) === AscFormat.BAR_DIR_COL) {
            aAxes = this.createRegularAxes(this.getAxisNumFormatByType(nType, this.getAllSeries()), false);
        }
        else {
            aAxes = this.createHBarAxes(this.getAxisNumFormatByType(nType, this.getAllSeries()), false);
        }
        oBarChart = this.createBarChart(nType, this.getAllSeries(), aAxes, this.charts[0]);
        this.addChartWithAxes(oBarChart);
    };
    CPlotArea.prototype.getGroupingByType = function(nType) {
        var oCT = Asc.c_oAscChartTypeSettings;
        var nNewGrouping = AscFormat.GROUPING_STANDARD;
        if(nType === oCT.lineNormal
            || nType === oCT.lineNormalMarker
            || nType === oCT.line3d
            || nType === oCT.areaNormal) {
            nNewGrouping = AscFormat.GROUPING_STANDARD;
        }
        else if(nType === oCT.lineStacked
            || nType === oCT.lineStackedMarker
            || nType === oCT.areaStacked) {
            nNewGrouping = AscFormat.GROUPING_STACKED;
        }
        else if(nType === oCT.lineStackedPer
            || nType === oCT.lineStackedPerMarker
            || nType === oCT.areaStackedPer) {
            nNewGrouping = AscFormat.GROUPING_PERCENT_STACKED;
        }
        return nNewGrouping;
    };
    CPlotArea.prototype.getIsMarkerByType = function(nType) {
        return getIsMarkerByType(nType);
    };
    CPlotArea.prototype.getIsSmoothByType = function(nType) {
        return getIsSmoothByType(nType);
    };
    CPlotArea.prototype.getIsLineByType = function(nType) {
        return getIsLineByType(nType);
    };
    CPlotArea.prototype.createLineChart = function(nType, aSeries, aAxes, oOldChart) {
        var nNewGrouping;
        nNewGrouping = this.getGroupingByType(nType);
        var nSeries, oSeries;
        var oLineChart = new AscFormat.CLineChart();
        oLineChart.mergeWithoutSeries(oOldChart);
        oLineChart.setGrouping(nNewGrouping);
        oLineChart.setVaryColors(false);
        oLineChart.setMarker(new AscFormat.CMarker());
        if(oLineChart.hiLowLines) {
            oLineChart.setHiLowLines(null);
        }
        if(oLineChart.upDownBars) {
            oLineChart.setUpDownBars(null);
        }
        oLineChart.marker.setSymbol(AscFormat.SYMBOL_NONE);
        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            oSeries = new AscFormat.CLineSeries();
            aSeries[nSeries].fillObject(oSeries);
            oLineChart.addSer(oSeries);
            oSeries.checkSeriesAfterChangeType();
        }
        oLineChart.addAxes(aAxes);
        var bMarker = this.getIsMarkerByType(nType);
        if(oLineChart.marker !== bMarker) {
            oLineChart.setMarkerValue(bMarker);
        }
        this.parent.check3DOptions(this.getIs3dByType(nType), false);
        return oLineChart;
    };
    CPlotArea.prototype.switchToLineChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var aSeries = this.getAllSeries();
        var aAxes = this.createRegularAxes(this.getAxisNumFormatByType(nType, aSeries), false);
        var oLineChart = this.createLineChart(nType, aSeries, aAxes, this.charts[0]);
        oLineChart.addAxes(aAxes);
        this.addChartWithAxes(oLineChart);
    };
	CPlotArea.prototype.createRadarChart = function(nType, aSeries, aAxes, oOldChart) {
		let oRadarChart = new AscFormat.CRadarChart();
		let nRadarStyle;
		if(nType === Asc.c_oAscChartTypeSettings.radarFilled) {
			nRadarStyle = AscFormat.RADAR_STYLE_FILLED;
		}
		else {
			nRadarStyle = AscFormat.RADAR_STYLE_MARKER;
		}
		oRadarChart.mergeWithoutSeries(oOldChart);
		oRadarChart.setRadarStyle(nRadarStyle);
        oRadarChart.setVaryColors(false);
		var nSeries, oSeries;
		for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
			oSeries = new AscFormat.CRadarSeries();
			aSeries[nSeries].fillObject(oSeries);
			oSeries.changeChartType(nType);
			oRadarChart.addSer(oSeries);
			oSeries.checkSeriesAfterChangeType(nType);
		}
		oRadarChart.addAxes(aAxes);
		this.parent.check3DOptions(false, false);
		return oRadarChart;
	};
	CPlotArea.prototype.switchToRadar = function(nType) {
		if(!this.parent) {
			return;
		}
		const aCharts = this.charts;
		if(aCharts.length < 1) {
			return;
		}
		if(aCharts.length === 1) {
			if(aCharts[0].tryChangeType(nType)) {
				return;
			}
		}
		const aSeries = this.getAllSeries();
		const aAxes = this.createRadarAxes(this.getAxisNumFormatByType(nType, aSeries), false);
		const oRadarChart = this.createRadarChart(nType, aSeries, aAxes, this.charts[0]);
		oRadarChart.addAxes(aAxes);
		this.addChartWithAxes(oRadarChart);
	};
    CPlotArea.prototype.createPieChart = function(nType, aSeries, oOldChart) {
        var oPieChart = new AscFormat.CPieChart();
        oPieChart.mergeWithoutSeries(oOldChart);
        oPieChart.setVaryColors(true);
        for(var nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            var oSeries = new AscFormat.CPieSeries();
            aSeries[nSeries].fillObject(oSeries);
            //remove this----------
            //----------------------
            oPieChart.addSer(oSeries);
	        oSeries.checkSeriesAfterChangeType();
        }
        if(nType === Asc.c_oAscChartTypeSettings.pie3d) {
            if(!this.parent.view3D) {
                this.parent.setView3D(AscFormat.CreateView3d(30, 0, true, 100));
                this.parent.setDefaultWalls();
            }
            if(!this.parent.view3D.rAngAx) {
                this.parent.view3D.setRAngAx(true);
            }
            if(this.parent.view3D.rotX < 0) {
                this.parent.view3D.rotX = 30;
            }
        }
        else {
            this.parent.check3DOptions(false, false);
        }
        return oPieChart;
    };
    CPlotArea.prototype.switchToPieChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var aSeries = this.getAllSeries();
        var oPieChart = this.createPieChart(nType, aSeries, this.charts[0]);
        this.addChartWithAxes(oPieChart);
    };
    CPlotArea.prototype.createDoughnutChart = function(nType, aSeries, oOldChart) {
        var oDoughnutChart = new AscFormat.CDoughnutChart();
        oDoughnutChart.mergeWithoutSeries(oOldChart);
        oDoughnutChart.setVaryColors(true);
        for(var nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            var oSeries = new AscFormat.CPieSeries();
            aSeries[nSeries].fillObject(oSeries);
            oDoughnutChart.addSer(oSeries);
	        oSeries.checkSeriesAfterChangeType();
        }
        this.parent.check3DOptions(false, false);
        return oDoughnutChart;
    };
    CPlotArea.prototype.switchToDoughnutChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var oDoughnutChart = this.createDoughnutChart(nType, this.getAllSeries(), this.charts[0]);
        this.addChartWithAxes(oDoughnutChart);
    };
    CPlotArea.prototype.createAreaChart = function(nType, aSeries, aAxes, oOldChart) {
        var nNewGrouping = this.getGroupingByType(nType);
        var nSeries, oSeries;
        var oAreaChart = new AscFormat.CAreaChart();
        oAreaChart.mergeWithoutSeries(oOldChart);
        oAreaChart.setGrouping(nNewGrouping);
        oAreaChart.setVaryColors(false);
        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            oSeries = new AscFormat.CAreaSeries();
            aSeries[nSeries].fillObject(oSeries);
            oAreaChart.addSer(oSeries);
	        oSeries.checkSeriesAfterChangeType();
        }
        oAreaChart.addAxes(aAxes);
        this.parent.check3DOptions(false, false);
        return oAreaChart;
    };
    CPlotArea.prototype.switchToAreaChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var aSeries = this.getAllSeries();
        var aAxes = this.createRegularAxes(this.getAxisNumFormatByType(nType, aSeries), false, true);
        var oAreaChart = this.createAreaChart(nType, aSeries, aAxes, this.charts[0]);
        oAreaChart.addAxes(aAxes);
        this.addChartWithAxes(oAreaChart);
    };
    CPlotArea.prototype.createScatterChart = function(nType, aSeries, aAxes, oOldChart) {
        var oScatterChart = new AscFormat.CScatterChart();
        oScatterChart.mergeWithoutSeries(oOldChart);
        oScatterChart.setVaryColors(false);
        var nSeries, oSeries;
        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            oSeries = new AscFormat.CScatterSeries();
            aSeries[nSeries].fillObject(oSeries);
            oScatterChart.addSer(oSeries);
	        oSeries.checkSeriesAfterChangeType();
        }
        oScatterChart.setScatterStyle(AscFormat.SCATTER_STYLE_MARKER);
        oScatterChart.addAxes(aAxes);
        this.parent.check3DOptions(false, false);
        if(nType !== Asc.c_oAscChartTypeSettings.scatter) {
            oScatterChart.parent = this;
            oScatterChart.tryChangeType(nType);
            oScatterChart.parent = null;
        }
        return oScatterChart;
    };
    CPlotArea.prototype.switchToScatterChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        var aSeries = this.getAllSeries();
        var aAxes = this.createScatterAxes(false);
        var oScatterChart = this.createScatterChart(nType, aSeries, aAxes, this.charts[0]);
        this.addChartWithAxes(oScatterChart);
    };
    CPlotArea.prototype.createStockChart = function(nType, aSeries, aAxes, oOldChart) {
        var oStockChart = new AscFormat.CStockChart();
        oStockChart.mergeWithoutSeries(oOldChart);
        oStockChart.setVaryColors(false);
        var nSeries, oSeries;
        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            oSeries = new AscFormat.CLineSeries();
            aSeries[nSeries].fillObject(oSeries);
            oStockChart.addSer(oSeries);
        }
        oStockChart.setHiLowLines(new AscFormat.CSpPr());
        oStockChart.setUpDownBars(new AscFormat.CUpDownBars());
        oStockChart.upDownBars.setGapWidth(150);
        oStockChart.upDownBars.setUpBars(new AscFormat.CSpPr());
        oStockChart.upDownBars.setDownBars(new AscFormat.CSpPr());
        oStockChart.addAxes(aAxes);
        this.parent.check3DOptions(false, false);
        return oStockChart;
    };
    CPlotArea.prototype.switchToStockChart = function(nType) {
        if(!this.parent) {
            return;
        }
        var aCharts = this.charts;
        if(aCharts.length < 1) {
            return;
        }
        if(aCharts.length === 1) {
            if(aCharts[0].tryChangeType(nType)) {
                return;
            }
        }
        if(!this.canChangeToStockChart()) {
            return;
        }
        var aAxes = this.createRegularAxes("General", false);
        var oStockChart = this.createStockChart(nType, this.getAllSeries(), aAxes, this.charts[0]);
        this.addChartWithAxes(oStockChart);
    };
    CPlotArea.prototype.canChangeToStockChart = function() {
        var oChartSpace = this.getChartSpace();
        var nType = oChartSpace.getChartType();
        if(nType === Asc.c_oAscChartTypeSettings.stock) {
            return true;
        }
        return (this.getAllSeries().length === AscFormat.MIN_STOCK_COUNT);
    };
    CPlotArea.prototype.changeChartType = function(nType) {
        if(!this.parent) {
            return;
        }
        if(this.charts.length < 1) {
            return;
        }
        var nCurType = this.getChartType();
        if(nCurType === nType) {
            return;
        }
        if(nType === Asc.c_oAscChartTypeSettings.comboCustom
            || nType === Asc.c_oAscChartTypeSettings.comboAreaBar
            || nType === Asc.c_oAscChartTypeSettings.comboBarLine
            || nType === Asc.c_oAscChartTypeSettings.comboBarLineSecondary) {
            this.switchToCombo(nType);
        }
        else if(this.isBarType(nType) || this.isHBarType(nType)) {
            this.switchToBarChart(nType);
        }
        else if(this.isLineType(nType)) {
            this.switchToLineChart(nType);
        }
        else if(this.isPieType(nType)) {
            this.switchToPieChart(nType);
        }
        else if(this.isDoughnutType(nType)) {
            this.switchToDoughnutChart(nType);
        }
        else if(this.isAreaType(nType)) {
            this.switchToAreaChart(nType);
        }
        else if(this.isScatterType(nType)) {
            this.switchToScatterChart(nType);
        }
        else if(this.isStockChart(nType)) {
            this.switchToStockChart(nType);
        }
		else if(this.isRadarChart(nType)) {
			this.switchToRadar(nType);
        }
    };
    CPlotArea.prototype.getAllSeries = function() {
        var _ret = [];
        var aCharts = this.charts;
        for(var i = 0; i < aCharts.length; ++i) {
            _ret = _ret.concat(aCharts[i].series);
        }
        _ret.sort(function(a, b) {
            return a.idx - b.idx;
        });
        return _ret;
    };
    CPlotArea.prototype.getMaxSeriesIdx = function() {
        var aAllSeries = this.getAllSeries();
        if(aAllSeries.length === 0) {
            return -1;
        }
        return aAllSeries[aAllSeries.length - 1].idx;
    };
    CPlotArea.prototype.removeAllCharts = function() {
        this.removeCharts(0, this.charts.length);
    };
    CPlotArea.prototype.removeAllAxes = function() {
        var nStart = this.axId.length - 1;
        var nEnd = 0;
        for(var nAxIdx = nStart; nAxIdx >= nEnd; --nAxIdx) {
            this.removeAxisByPos(nAxIdx);
        }
    };
    CPlotArea.prototype.checkDlblsPosition = function() {
        var aCharts = this.charts;
        for(var nChart = 0; nChart < aCharts.length; ++nChart) {
            aCharts[nChart].checkDlblsPosition();
        }
    };
    CPlotArea.prototype.getPossibleDLblsPosition = function() {
        var aCharts = this.charts;
        if(aCharts.length === 0) {
            return [];
        }
        var aPositions = aCharts[0].getPossibleDLblsPosition(), aCurPositions, nLblPos;
        if(aPositions.length > 0) {
            for(var nChart = 1; nChart < aCharts.length; ++nChart) {
                aCurPositions = aCharts[nChart].getPossibleDLblsPosition();
                if(aCurPositions.length === 0) {
                    return [];
                }
                for(var nPos = aPositions.length - 1; nPos > -1; --nPos) {
                    nLblPos = aPositions[nPos];
                    for(var nCurPos = 0; nCurPos < aCurPositions.length; ++nCurPos) {
                        if(nLblPos === aCurPositions[nCurPos]) {
                            break;
                        }
                    }
                    if(nCurPos === aCurPositions.length) {
                        aPositions.splice(nPos, 1);
                    }
                }
                if(aPositions.length === 0) {
                    return aPositions;
                }
            }
        }
        return aPositions;
    };
    CPlotArea.prototype.getAxesCrosses = function() {
        var aRet = [];
        var oCheckedMap = {};
        var oCurAxis;
        var aCross;
        for(var nAx = 0; nAx < this.axId.length; ++nAx) {
            oCurAxis = this.axId[nAx];
            aCross = [];
            while(oCurAxis && !oCheckedMap[oCurAxis.Id]) {
                aCross.push(oCurAxis);
                oCheckedMap[oCurAxis.Id] = true;
                oCurAxis = oCurAxis.crossAx;
            }
            if(aCross.length > 1) {
                aRet.push(aCross);
            }
        }
        return aRet;
    };
    CPlotArea.prototype.getCatValMergeAxes = function() {
        var aCrosses = this.getAxesCrosses();
        var aCross = aCrosses[0];
        var oCatAxMerge = null, oValAxMerge = null, oAxis, nAx;
        if(aCross) {
            for(nAx = 0; nAx < aCross.length; ++nAx) {
                oAxis = aCross[nAx];
                if(oAxis.getObjectType() === AscDFH.historyitem_type_CatAx ||
                    oAxis.getObjectType() === AscDFH.historyitem_type_DateAx) {
                    oCatAxMerge = oAxis;
                }
                else if(oAxis.getObjectType() === AscDFH.historyitem_type_ValAx) {
                    oValAxMerge = oAxis;
                }
            }
            if(!oCatAxMerge || !oValAxMerge) {
                for(nAx = 0; nAx < aCross.length; ++nAx) {
                    oAxis = aCross[nAx];
                    if(oAxis.getObjectType() === AscDFH.historyitem_type_ValAx) {
                        if(oAxis.axPos === AX_POS_L || oAxis.axPos === AX_POS_R) {
                            oValAxMerge = oAxis;
                        }
                        else {
                            oCatAxMerge = oAxis;
                        }
                    }
                }
            }
        }
        if(oCatAxMerge && oValAxMerge) {
            return [oCatAxMerge, oValAxMerge];
        }
        return null;
    };
    CPlotArea.prototype.createCatValAxes = function(sNewNumFormat) {
        var oAxes = AscFormat.CreateDefaultAxes(sNewNumFormat);
        var oCatAx = oAxes.catAx;
        var oValAx = oAxes.valAx;
        var aMerge = this.getCatValMergeAxes();
        if(aMerge) {
            oCatAx.merge(aMerge[0]);
            oValAx.merge(aMerge[1]);
        }
        oValAx.checkNumFormat(sNewNumFormat);
        return [oCatAx, oValAx];
    };
    CPlotArea.prototype.createHBarAxes = function(sNewNumFormat, bSecondary) {
        var aHAxes = this.createCatValAxes(sNewNumFormat);
        var oCatAx = aHAxes[0], oValAx = aHAxes[1];
        var aAxes = this.axId;
        var nAx, oAx;
        oValAx.setAxPos(AX_POS_B);
        oCatAx.setAxPos(AX_POS_L);
        if(bSecondary) {
            oCatAx.setDelete(true);
            for(nAx = 0; nAx < aAxes.length; ++nAx) {
                oAx = aAxes[nAx];
                if(oAx.axPos === AX_POS_B || oAx.axPos === AX_POS_T) {
                    if(oAx.axPos === AX_POS_B) {
                        oValAx.setAxPos(AX_POS_T);
                    }
                    else {
                        oValAx.setAxPos(AX_POS_B);
                    }
                    if(oAx.crosses === null || oAx.crosses === CROSSES_AUTO_ZERO || oAx.crosses === CROSSES_MIN) {
                        oValAx.setCrosses(CROSSES_MAX);
                    }
                    else {
                        oValAx.setCrosses(CROSSES_AUTO_ZERO);
                    }
                }
            }

            oValAx.setMajorGridlines(null);
            oValAx.setMinorGridlines(null);
        }
        return [oCatAx, oValAx];
    };
	CPlotArea.prototype.createRadarAxes = function(sNewNumFormat, bSecondary) {

        const oCurAxes = this.getAxisByTypes();
		const aAxes = this.createRegularAxes(sNewNumFormat, false);
        const oValAx = aAxes[1];
        if(oValAx) {
            if(oValAx.crossBetween !== AscFormat.CROSS_BETWEEN_BETWEEN) {
                oValAx.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
            }
            if(bSecondary) {
                const oOldValAx = oCurAxes.valAx[0];
                if(oOldValAx) {
                    const aCharts = this.getChartsForAxis(oOldValAx);
                    const oChart = aCharts[0];
                    if(oChart && oChart.getObjectType() === AscDFH.historyitem_type_RadarChart) {
                        oValAx.setMajorGridlines(null);
                        oValAx.setMinorGridlines(null);
                    }
                }
            }
        }
        return aAxes;
	};
    CPlotArea.prototype.createRegularAxes = function(sNewNumFormat, bSecondary, bArea) {
        var aRegAxes = this.createCatValAxes(sNewNumFormat);
        var oCatAx = aRegAxes[0], oValAx = aRegAxes[1];
        var aAxes = this.axId;
        var nAx, oAx;
        oValAx.setAxPos(AX_POS_L);
        oCatAx.setAxPos(AX_POS_B);
		if(bArea) {
			oValAx.setCrossBetween(AscFormat.CROSS_BETWEEN_MID_CAT);
		}
        if(bSecondary) {
            oCatAx.setDelete(true);
            for(nAx = 0; nAx < aAxes.length; ++nAx) {
                oAx = aAxes[nAx];
                if(oAx.axPos === AX_POS_L || oAx.axPos === AX_POS_R) {
                    if(oAx.axPos === AX_POS_L) {
                        oValAx.setAxPos(AX_POS_R);
                    }
                    else {
                        oValAx.setAxPos(AX_POS_L);
                    }
                    if(oAx.crosses === null || oAx.crosses === CROSSES_AUTO_ZERO || oAx.crosses === CROSSES_MIN) {
                        oValAx.setCrosses(CROSSES_MAX);
                        oCatAx.setCrosses(CROSSES_MAX);
                    }
                    else {
                        oValAx.setCrosses(CROSSES_AUTO_ZERO);
                        oCatAx.setCrosses(CROSSES_AUTO_ZERO);
                    }
                }
            }
            oValAx.setMajorGridlines(null);
            oValAx.setMinorGridlines(null);
        }
        return [oCatAx, oValAx];
    };
    CPlotArea.prototype.createScatterAxes = function(bSecondary) {
        var oAxes = AscFormat.CreateScatterAxis();
        var aMergeAxes = this.getCatValMergeAxes();
        if(aMergeAxes) {
            oAxes.catAx.merge(aMergeAxes[0]);
            oAxes.valAx.merge(aMergeAxes[1]);
        }
        var aAxes = this.axId;
        var nAx, oAx;
        oAxes.valAx.setAxPos(AX_POS_L);
        oAxes.catAx.setAxPos(AX_POS_B);
        if(bSecondary) {
            for(nAx = 0; nAx < aAxes.length; ++nAx) {
                oAx = aAxes[nAx];
                if(oAx.axPos === AX_POS_L || oAx.axPos === AX_POS_R) {
                    if(oAx.axPos === AX_POS_L) {
                        oAxes.valAx.setAxPos(AX_POS_R);
                    }
                    else {
                        oAxes.valAx.setAxPos(AX_POS_L);
                    }
                    if(oAx.crosses === null || oAx.crosses === CROSSES_AUTO_ZERO || oAx.crosses === CROSSES_MIN) {
                        oAxes.valAx.setCrosses(CROSSES_MAX);
                    }
                    else {
                        oAxes.valAx.setCrosses(CROSSES_AUTO_ZERO);
                    }
                }
                else {
                    if(oAx.axPos === AX_POS_T) {
                        oAxes.catAx.setAxPos(AX_POS_B);
                    }
                    else {
                        oAxes.catAx.setAxPos(AX_POS_T);
                    }
                    if(oAx.crosses === null || oAx.crosses === CROSSES_AUTO_ZERO || oAx.crosses === CROSSES_MIN) {
                        oAxes.catAx.setCrosses(CROSSES_MAX);
                    }
                    else {
                        oAxes.catAx.setCrosses(CROSSES_AUTO_ZERO);
                    }
                }
            }
            oAxes.valAx.setMajorGridlines(null);
            oAxes.valAx.setMinorGridlines(null);
            oAxes.catAx.setMajorGridlines(null);
            oAxes.catAx.setMinorGridlines(null);
        }
        return [oAxes.catAx, oAxes.valAx];
    };
    CPlotArea.prototype.createSurfaceAxes = function(sNewNumFormat) {
        var oAxes = AscFormat.CreateSurfaceAxes(sNewNumFormat);

        var aMerge = this.getCatValMergeAxes();
        if(aMerge) {
            oAxes.catAx.merge(aMerge[0]);
            oAxes.valAx.merge(aMerge[1]);
        }
        oAxes.valAx.checkNumFormat(sNewNumFormat);
        return [oAxes.catAx, oAxes.valAx, oAxes.serAx];
    };
    CPlotArea.prototype.getOrderedAxes = function() {
        return new COrderedAxes(this);
    };
    CPlotArea.prototype.isBarType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.barNormal === nType
            || Asc.c_oAscChartTypeSettings.barStacked === nType
            || Asc.c_oAscChartTypeSettings.barStackedPer === nType
            || Asc.c_oAscChartTypeSettings.barNormal3d === nType
            || Asc.c_oAscChartTypeSettings.barStacked3d === nType
            || Asc.c_oAscChartTypeSettings.barStackedPer3d === nType
            || Asc.c_oAscChartTypeSettings.barNormal3dPerspective === nType) {
            return true;
        }
        return false;
    };
    CPlotArea.prototype.isHBarType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.hBarNormal === nType
            || Asc.c_oAscChartTypeSettings.hBarStacked === nType
            || Asc.c_oAscChartTypeSettings.hBarStackedPer === nType
            || Asc.c_oAscChartTypeSettings.hBarNormal3d === nType
            || Asc.c_oAscChartTypeSettings.hBarStacked3d === nType
            || Asc.c_oAscChartTypeSettings.hBarStackedPer3d === nType) {
            return true;
        }
        return false;
    };
    CPlotArea.prototype.isLineType = function(nType) {
        return getIsLineType(nType);
    };
    CPlotArea.prototype.isPieType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.pie === nType
            || Asc.c_oAscChartTypeSettings.pie3d === nType) {
            return true
        }
        return false;
    };
    CPlotArea.prototype.isDoughnutType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.doughnut === nType) {
            return true
        }
        return false;
    };
    CPlotArea.prototype.isAreaType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.areaNormal === nType
            || Asc.c_oAscChartTypeSettings.areaStacked === nType
            || Asc.c_oAscChartTypeSettings.areaStackedPer === nType) {
            return true
        }
        return false;
    };
    CPlotArea.prototype.isRadarType = function(nType) {
        if(Asc.c_oAscChartTypeSettings.radar === nType
            || Asc.c_oAscChartTypeSettings.radarMarker === nType
            || Asc.c_oAscChartTypeSettings.radarFilled === nType) {
            return true
        }
        return false;
    };
    CPlotArea.prototype.isScatterType = function(nType) {
        return isScatterChartType(nType);
    };
    CPlotArea.prototype.isStockChart = function(nType) {
        if(Asc.c_oAscChartTypeSettings.stock === nType) {
            return true
        }
        return false;
    };
    CPlotArea.prototype.isRadarChart = function(nType) {
      return this.isRadarType(nType);
    };
    CPlotArea.prototype.setDlblsProps = function(oProps) {
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            this.charts[nChart].setDlblsProps(oProps);
        }
    };
    CPlotArea.prototype.is3dChart = function() {
        if(!this.parent) {
            return false;
        }
        return this.parent.is3dChart();
    };
    CPlotArea.prototype.hasChartWithSecondaryAxis = function() {
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            if(this.charts[nChart].isSecondaryAxis()) {
                return this.charts[nChart];
            }
        }
        return null;
    };
    CPlotArea.prototype.getChartsWithSecondaryAxis = function() {
        var aRet = [];
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            if(this.charts[nChart].isSecondaryAxis()) {
                aRet.push(this.charts[nChart]);
            }
        }
        return aRet;
    };
    CPlotArea.prototype.addAxes = function(aAxes) {
        for(var nAx = 0; nAx < aAxes.length; ++nAx) {
            this.addAxis(aAxes[nAx]);
        }
    };
    CPlotArea.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(!this.is3dChart()) {
            this.applyStyleEntry(oChartStyle.plotArea, oColors.generateColors(1), 0, bReset);
        }
        else {
            this.applyStyleEntry(oChartStyle.plotArea3D, oColors.generateColors(1), 0, bReset);
        }
        for(var nChart = 0; nChart < this.charts.length; ++nChart) {
            this.charts[nChart].applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.dTable) {
            this.dTable.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        for(var nAx = 0; nAx < this.axId.length; ++nAx) {
            this.axId[nAx].applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
    };
    CPlotArea.prototype.initPostOpen = function(aChartWithAxis) {
        // выставляем axis в chart
        // TODO: 1. Диаграмм может быть больше, но мы пока работаем только с одной
        // TODO: 2. Избавиться от oIdToAxisMap, aChartWithAxis, т.к. они здесь больше не нужны
        ///  var oZeroChart = this.charts[0];
        ///  if ( oZeroChart )
        ///  {
        ///      var len = this.axId.length
        ///      for ( var i = 0; i < len; i++ )
        ///          oZeroChart.addAxId(this.axId[i]);
        ///  }
        let oIdToAxisMap = {};
        for(let nAxIndex = 0; nAxIndex < this.axId.length; ++nAxIndex)
        {
            let oCurAxis = this.axId[nAxIndex];
            if (null != oCurAxis.axId)
                oIdToAxisMap[oCurAxis.axId] = oCurAxis;
        }
        for(let nAxIndex = 0; nAxIndex < this.axId.length; ++nAxIndex)
        {
            let oCurAxis = this.axId[nAxIndex];
            oCurAxis.setCrossAx(oIdToAxisMap[oCurAxis.crossAxId]);
            delete oCurAxis.crossAxId;
        }
        for(var nChartIndex = 0; nChartIndex < aChartWithAxis.length; ++nChartIndex)
        {
            var oCurChartWithAxis = aChartWithAxis[nChartIndex];
            var axis = oIdToAxisMap[oCurChartWithAxis.axisId];
            oCurChartWithAxis.chart.addAxId(axis);
            if(axis && axis.getObjectType() === AscDFH.historyitem_type_ValAx && !AscFormat.isRealNumber(axis.crossBetween))
            {
                if(oCurChartWithAxis.chart.getObjectType() === AscDFH.historyitem_type_AreaChart)
                {
                    axis.setCrossBetween(AscFormat.CROSS_BETWEEN_MID_CAT);
                }
                else
                {
                    axis.setCrossBetween(AscFormat.CROSS_BETWEEN_BETWEEN);
                }
            }
        }


        //check: does category axis exist
        for(var _i = this.charts.length - 1; _i > -1; --_i){
            var oChart = this.charts[_i];
            if(oChart)
            {
                if(oChart.getObjectType() !== AscDFH.historyitem_type_ScatterChart &&
                    oChart.getObjectType() !== AscDFH.historyitem_type_PieChart &&
                    oChart.getObjectType() !== AscDFH.historyitem_type_DoughnutChart)
                {
                    var axis_by_types = oChart.getAxisByTypes();
                    if(axis_by_types.valAx.length === 0 || axis_by_types.catAx.length === 0)
                    {
                        this.removeCharts(_i, _i);
                        if(oChart.axId){
                            oChart.axId.length = 0;
                            oChart = oChart.createDuplicate();
                            if(oChart.setParent){
                                oChart.setParent(this);
                            }
                        }
                        var sDefaultValAxFormatCode = null;
                        if(oChart && oChart.series[0]){
                            var aPoints = oChart.series[0].getNumPts();
                            if(aPoints[0] && typeof aPoints[0].formatCode === "string" && aPoints[0].formatCode.length > 0){
                                sDefaultValAxFormatCode = aPoints[0].formatCode;
                            }
                        }
                        var need_num_fmt = sDefaultValAxFormatCode;
                        var axis_obj = AscFormat.CreateDefaultAxes(need_num_fmt ? need_num_fmt : "General");
                        var cat_ax = axis_obj.catAx;
                        var val_ax = axis_obj.valAx;
                        if(oChart.getObjectType() === AscDFH.historyitem_type_BarChart && oChart.barDir === AscFormat.BAR_DIR_BAR)
                        {
                            if(cat_ax.axPos !== AscFormat.AX_POS_L)
                            {
                                cat_ax.setAxPos(AscFormat.AX_POS_L);
                            }
                            if(val_ax.axPos !== AscFormat.AX_POS_B)
                            {
                                val_ax.setAxPos(AscFormat.AX_POS_B);
                            }
                        }
                        else
                        {
                            if(cat_ax.axPos !== AscFormat.AX_POS_B)
                            {
                                cat_ax.setAxPos(AscFormat.AX_POS_B);
                            }
                            if(val_ax.axPos !== AscFormat.AX_POS_L)
                            {
                                val_ax.setAxPos(AscFormat.AX_POS_L);
                            }
                        }

                        this.addChart(oChart);
                        oChart.addAxId(cat_ax);
                        oChart.addAxId(val_ax);
                        this.addAxis(cat_ax);
                        this.addAxis(val_ax);
                    }
                    else
                    {
                        if(oChart.getObjectType() === AscDFH.historyitem_type_BarChart && oChart.barDir === AscFormat.BAR_DIR_BAR)
                        {
                            for(var _c = 0; _c < axis_by_types.valAx.length; ++_c)
                            {
                                var val_ax = axis_by_types.valAx[_c];
                                if(val_ax.axPos !== AscFormat.AX_POS_B && val_ax.axPos !== AscFormat.AX_POS_T )
                                {
                                    val_ax.setAxPos(AscFormat.AX_POS_B);
                                }
                            }
                            for(var _c = 0; _c < axis_by_types.catAx.length; ++_c)
                            {
                                var cat_ax = axis_by_types.catAx[_c];
                                if(cat_ax.axPos !== AscFormat.AX_POS_L && cat_ax.axPos !== AscFormat.AX_POS_R )
                                {
                                    cat_ax.setAxPos(AscFormat.AX_POS_L);
                                }
                            }
                        }
                    }
                }
                if(oChart.setVaryColors && oChart.varyColors === null){
                    oChart.setVaryColors(false);
                }
                if(oChart.setSmooth && oChart.smooth === null){
                    //oChart.setSmooth(false);
                }
                if(oChart.setGapWidth && oChart.gapWidth === null){
                    oChart.setGapWidth(150);
                }
                var oDlbls;
                if(oChart.setDLbls && oChart.dLbls === null){
                    oDlbls = new AscFormat.CDLbls();
                    oDlbls.setShowLegendKey(false);
                    oDlbls.setShowVal(false);
                    oDlbls.setShowCatName(false);
                    oDlbls.setShowSerName(false);
                    oDlbls.setShowPercent(false);
                    oDlbls.setShowBubbleSize(false);
                    oChart.setDLbls(oDlbls);
                }
            }
        }
    };

    function getIsMarkerByType(nType) {
        if(nType === Asc.c_oAscChartTypeSettings.scatter ||
            nType === Asc.c_oAscChartTypeSettings.scatterLineMarker ||
            nType === Asc.c_oAscChartTypeSettings.scatterMarker ||
            nType === Asc.c_oAscChartTypeSettings.scatterSmoothMarker ||
            nType === Asc.c_oAscChartTypeSettings.lineNormalMarker ||
            nType === Asc.c_oAscChartTypeSettings.lineStackedMarker ||
            nType === Asc.c_oAscChartTypeSettings.lineStackedPerMarker ||
            nType === Asc.c_oAscChartTypeSettings.radarMarker) {
            return true;
        }
        return false;
    }


    function getIsSmoothByType(nType) {
        if(nType === Asc.c_oAscChartTypeSettings.scatterSmoothMarker ||
            nType === Asc.c_oAscChartTypeSettings.scatterSmooth) {
            return true;
        }
        return false;
    }

    function getIsLineByType(nType) {
        if(nType === Asc.c_oAscChartTypeSettings.scatterSmoothMarker ||
            nType === Asc.c_oAscChartTypeSettings.scatterSmooth ||
            nType === Asc.c_oAscChartTypeSettings.scatterLineMarker ||
            nType === Asc.c_oAscChartTypeSettings.scatterLine ||
            nType === Asc.c_oAscChartTypeSettings.lineNormal ||
            nType === Asc.c_oAscChartTypeSettings.lineStacked ||
            nType === Asc.c_oAscChartTypeSettings.lineStackedPer ||
            nType === Asc.c_oAscChartTypeSettings.lineNormalMarker ||
            nType === Asc.c_oAscChartTypeSettings.lineStackedMarker ||
            nType === Asc.c_oAscChartTypeSettings.lineStackedPerMarker ||
            nType === Asc.c_oAscChartTypeSettings.line3d ||
            nType === Asc.c_oAscChartTypeSettings.radar ||
            nType === Asc.c_oAscChartTypeSettings.radarMarker) {
            return true;
        }
        return false;
    }

    function getIsLineType(nType) {
        if(Asc.c_oAscChartTypeSettings.lineNormal === nType
            || Asc.c_oAscChartTypeSettings.lineStacked === nType
            || Asc.c_oAscChartTypeSettings.lineStackedPer === nType
            || Asc.c_oAscChartTypeSettings.lineNormalMarker === nType
            || Asc.c_oAscChartTypeSettings.lineStackedMarker === nType
            || Asc.c_oAscChartTypeSettings.lineStackedPerMarker === nType
            || Asc.c_oAscChartTypeSettings.line3d === nType) {
            return true;
        }
        return false;
    }

    function COrderedAxes(oPlotArea) {
        this.verticalAxes = [];
        this.horizontalAxes = [];
        this.depthAxes = [];
        var aTypedCharts = oPlotArea.charts;
        var oCheckedAxes = {};
        var nChart;
        var oTypedChart;
        var aAxes, oAxis, nAx;
        for(nChart = 0; nChart < aTypedCharts.length; ++nChart) {
            oTypedChart = aTypedCharts[nChart];
            aAxes = oTypedChart.axId;
            for(nAx = 0; nAx < aAxes.length; ++nAx) {
                oAxis = aAxes[nAx];
                if(!oCheckedAxes[oAxis.Id]) {
                    if(oAxis.getObjectType() === AscDFH.historyitem_type_SerAx) {
                        this.depthAxes.push(oAxis);
                    }
                    else if(oAxis.axPos === AX_POS_L || oAxis.axPos === AX_POS_R) {
                        this.verticalAxes.push(oAxis);
                    }
                    else {
                        this.horizontalAxes.push(oAxis);
                    }
                    oCheckedAxes[oAxis.Id] = true;
                }
            }
        }
    }

    COrderedAxes.prototype.getVerticalAxes = function() {
        return this.verticalAxes;
    };
    COrderedAxes.prototype.getHorizontalAxes = function() {
        return this.horizontalAxes;
    };
    COrderedAxes.prototype.getDepthAxes = function() {
        return this.depthAxes;
    };

    function CChartBase() {
        CBaseChartObject.call(this);
        this.series = [];
        this.filteredSeries = [];
        this.axId = [];
        this.dLbls = null;
        this.varyColors = null;
    }
    InitClass(CChartBase, CBaseChartObject, AscDFH.historyitem_type_Unknown);
    CChartBase.prototype.getChildren = function() {
        var aRet = [this.dLbls];
        for(var nSeries = 0; nSeries < this.series.length; ++nSeries) {
            aRet.push(this.series[nSeries]);
        }
        return aRet;
    };
    CChartBase.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.dLbls)
            oCopy.setDLbls(this.dLbls.createDuplicate());
        if(AscFormat.isRealBool(this.varyColors))
            oCopy.setVaryColors(this.varyColors);
        for(var nSeries = 0; nSeries < this.series.length; ++nSeries) {
            var ser = oCopy.getEmptySeries();
            this.series[nSeries].fillObject(ser);
            oCopy.addSer(ser);
        }
    };
    CChartBase.prototype.removeSeriesInternal = function(idx) {
        if(this.series[idx]) {
            var arrSeries = this.series.splice(idx, 1);
            arrSeries[0].setParent(null);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonChart_RemoveSeries, idx, arrSeries, false));
            this.onChangeDataRefs();
        }
    };
    CChartBase.prototype.removeSeries = function(idx) {
        this.removeSeriesInternal(idx);
        if(this.series.length === 0) {
            if(this.parent) {
                if(this.parent.charts.length > 1) {
                    this.parent.removeChart(this);
                }
            }
        }
    };
    CChartBase.prototype.removeAllSeries = function() {
        for(var nSeries = this.series.length - 1; nSeries > -1; --nSeries) {
            this.removeSeriesInternal(nSeries);
        }
    };
    CChartBase.prototype.reindexSeries = function() {
        if(this.parent) {
            this.parent.reindexSeries();
        }
    };
    CChartBase.prototype.reorderSeries = function() {
        if(this.parent) {
            this.parent.reorderSeries();
        }
    };
    CChartBase.prototype.moveSeriesUp = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesUp(oSeries);
        }
    };
    CChartBase.prototype.moveSeriesDown = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesDown(oSeries);
        }
    };
    CChartBase.prototype.sortSeries = function() {
        var aAllSeries = [].concat(this.series);
        this.removeAllSeries();
        aAllSeries.sort(function(a, b) {
            return a.order - b.order;
        });
        for(var nSeries = 0; nSeries < aAllSeries.length; ++nSeries) {
            this.addSer(aAllSeries[nSeries]);
        }
    };
    CChartBase.prototype.onDataUpdate = function() {
        if(this.parent) {
            this.parent.onDataUpdate();
        }
    };
    CChartBase.prototype.addSerSilent = function(ser) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonChart_AddSeries, this.series.length, [ser], true));
        this.series.push(ser);
        this.onChangeDataRefs();
        this.setParentToChild(ser);
    };
    CChartBase.prototype.addSer = function(ser) {
        this.addSerSilent(ser);
        this.onChartUpdateType();
    };
    CChartBase.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CommonChart_SetDlbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
        this.onChartUpdateDataLabels();
    };
    CChartBase.prototype.documentCreateFontMap = function(allFonts) {
        this.dLbls && this.dLbls.documentCreateFontMap(allFonts);
        for(var i = 0; i < this.series.length; ++i) {
            this.series[i].documentCreateFontMap(allFonts);
        }
    };
    CChartBase.prototype.applyLabelsFunction = function(fCallback, value, oDD) {
        this.dLbls && this.dLbls.applyLabelsFunction(fCallback, value, oDD);
        for(var i = 0; i < this.series.length; ++i) {
            this.series[i].applyLabelsFunction(fCallback, value, oDD);
        }
    };
    CChartBase.prototype.getAllRasterImages = function(images) {
        this.dLbls && this.dLbls.getAllRasterImages(images);
        for(var i = 0; i < this.series.length; ++i) {
            this.series[i].getAllRasterImages(images);
        }
    };
    CChartBase.prototype.checkSpPrRasterImages = function(images) {
        this.dLbls && this.dLbls.checkSpPrRasterImages(images);
        for(var i = 0; i < this.series.length; ++i) {
            this.series[i].checkSpPrRasterImages(images);
        }
    };
    CChartBase.prototype.removeDataLabels = function() {
        var i;
        for(i = 0; i < this.series.length; ++i) {
            if(typeof this.series[i].setDLbls === "function" && this.series[i].dLbls)
                this.series[i].setDLbls(null);
        }
    };
    CChartBase.prototype.addAxes = function(aAxes) {
        for(var nAx = 0; nAx < aAxes.length; ++nAx) {
            this.addAxId(aAxes[nAx]);
        }
    };
    CChartBase.prototype.addAxId = function(pr) {
        if(!pr)
            return;
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonChart_AddAxId, this.axId.length, [pr], true));
        this.axId.push(pr);
    };
    CChartBase.prototype.getAxisByTypes = CPlotArea.prototype.getAxisByTypes;
    CChartBase.prototype.setVaryColors = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_CommonChart_SetVaryColors, this.varyColors, pr));
        this.varyColors = pr;
    };
    CChartBase.prototype.getSeriesArrayIdx = function(oSeries) {
        for(var i = 0; i < this.series.length; ++i) {
            if(this.series[i] === oSeries) {
                return i;
            }
        }
        return -1;
    };
    CChartBase.prototype.getFilteredSeriesArrayIdx = function(oSeries) {
        for(var i = 0; i < this.filteredSeries.length; ++i) {
            if(this.filteredSeries[i] === oSeries) {
                return i;
            }
        }
        return -1;
    };
    CChartBase.prototype.addFilteredSeries = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonChart_AddFilteredSeries, this.filteredSeries.length, [pr], true));
        this.filteredSeries.push(pr);
        this.setParentToChild(pr);
    };
    CChartBase.prototype.removeFilteredSeries = function(idx) {
        if(this.filteredSeries[idx]) {
            this.filteredSeries[idx].setParent(null);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonChart_AddFilteredSeries, idx, this.filteredSeries.splice(idx, 1), false));
        }
    };
    CChartBase.prototype.getChartType = function() {
        return Asc.c_oAscChartTypeSettings.unknown;
    };
    CChartBase.prototype.isSecondaryAxis = function() {
        if(!Array.isArray(this.axId) || this.axId.length === 0) {
            return undefined;
        }
        var oPlotArea = this.parent;
        if(!oPlotArea) {
            return false;
        }
        var aAllAxes = oPlotArea.axId;
        var oFirstAxis = aAllAxes[0];
        if(!oFirstAxis) {
            return false;
        }
        var nAxis;
        for(nAxis = 0; nAxis < this.axId.length; ++nAxis) {
            if(this.axId[nAxis] === oFirstAxis) {
                return false;
            }
        }
        return true;
    };
    CChartBase.prototype.setDLblsDeleteValue = function(bVal) {
        if(this.dLbls) {
            this.dLbls.setDeleteValue(bVal)
        }
        var nSeries, nSeriesCount = this.series.length;
        for(nSeries = 0; nSeries < nSeriesCount; ++nSeries) {
            this.series[nSeries].setDLblsDeleteValue(bVal);
        }
    };
    CChartBase.prototype.getPossibleDLblsPosition = function() {
        var aPositions = [];
        if(!this.parent) {
            return aPositions;
        }
        var b3D = this.parent.is3dChart();
        switch(this.getObjectType()) {
            case AscDFH.historyitem_type_BarChart:
            {
                if(!b3D) {
                    aPositions.push(c_oAscChartDataLabelsPos.ctr);
                    aPositions.push(c_oAscChartDataLabelsPos.inEnd);
                    aPositions.push(c_oAscChartDataLabelsPos.inBase);
                    if(this.grouping === AscFormat.BAR_GROUPING_CLUSTERED) {
                        aPositions.push(c_oAscChartDataLabelsPos.outEnd);
                    }
                }
                break;
            }
            case AscDFH.historyitem_type_LineChart:
            case AscDFH.historyitem_type_ScatterChart:
            {
                if(!b3D) {
                    aPositions.push(c_oAscChartDataLabelsPos.ctr);
                    aPositions.push(c_oAscChartDataLabelsPos.l);
                    aPositions.push(c_oAscChartDataLabelsPos.t);
                    aPositions.push(c_oAscChartDataLabelsPos.r);
                    aPositions.push(c_oAscChartDataLabelsPos.b);
                }
                break;
            }
            case AscDFH.historyitem_type_PieChart:
            {
                aPositions.push(c_oAscChartDataLabelsPos.ctr);
                aPositions.push(c_oAscChartDataLabelsPos.inEnd);
                aPositions.push(c_oAscChartDataLabelsPos.outEnd);
                aPositions.push(c_oAscChartDataLabelsPos.bestFit);
                break;
            }
        }
        return aPositions;
    };
    CChartBase.prototype.checkDlblsPosition = function() {
        var aPossiblePositions = this.getPossibleDLblsPosition();
        if(this.dLbls) {
            this.dLbls.checkPosition(aPossiblePositions);
        }
        for(var nSeries = 0; nSeries < this.series.length; ++nSeries) {
            this.series[nSeries].checkDlblsPosition(aPossiblePositions);
        }
    };
    CChartBase.prototype.mergeWithoutSeries = function(oTypedChart) {
        if(!oTypedChart) {
            return;
        }
        var aSeries = oTypedChart.series;
        oTypedChart.series = [];
        oTypedChart.fillObject(this);
        oTypedChart.series = aSeries;
    };
    CChartBase.prototype.setDlblsProps = function(oProps) {
        var nPos = oProps.getDataLabelsPos();
        if(nPos === Asc.c_oAscChartDataLabelsPos.none) {
            if(this.dLbls) {
                this.setDLbls(null);
            }
            this.removeDataLabels();
        }
        else {
            nPos = fCheckDLPostion(nPos, this.getPossibleDLblsPosition());
            if(!this.dLbls) {
                this.setDLbls(new AscFormat.CDLbls());
            }
            this.dLbls.setSettings(nPos, oProps);
            this.dLbls.checkChartStyle();
            for(var nSer = 0; nSer < this.series.length; ++nSer) {
                this.series[nSer].setDlblsProps(oProps);
            }
        }
    };
    CChartBase.prototype.canChangeAxisType = function() {
        if(this.getObjectType() === AscDFH.historyitem_type_PieChart ||
            this.getObjectType() === AscDFH.historyitem_type_DoughnutChart) {
            //TODO: implement
            //<c:extLst>
            // <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
            //    <c16:uniqueId val="{0000000D-5343-412C-A86B-69BCE2BFEEEB}"/>
            // </c:ext>
            //</c:extLst>
            return false;
        }
        var bHBarChart = this.isHBar();
        var aCharts = this.parent.charts;
        var nChart, oChart;
        if(this.getObjectType() === AscDFH.historyitem_type_BarChart &&
            this.barDir === AscFormat.BAR_DIR_BAR) {
            bHBarChart = true;
        }
        for(nChart = 0; nChart < aCharts.length; ++nChart) {
            oChart = aCharts[nChart];
            if(bHBarChart !== oChart.isHBar()) {
                return false;
            }
        }
        return true;
    };
    CChartBase.prototype.isHBar = function() {
        return (this.getObjectType() === AscDFH.historyitem_type_BarChart &&
        this.barDir === AscFormat.BAR_DIR_BAR);
    };
    CChartBase.prototype.checkValAxesFormatByType = function(nType) {
        if(!this.parent) {
            return;
        }
        if(Array.isArray(this.axId)) {
            var oAxes = this.getAxisByTypes();
            var aAxes = oAxes.valAx;
            for(var nAx = 0; nAx < aAxes.length; ++nAx) {
                var oAxis = aAxes[nAx];
                oAxis.checkNumFormat(this.parent.getAxisNumFormatByType(nType, this.series));
            }
        }
    };
    CChartBase.prototype.tryMoveSeries = function(oSeries, nType, bIsSecondaryAxis, bExactType) {
        var aCharts = this.parent.charts;
        var nChart, oChart, oNewSeries;
        var bCanMove;
        for(nChart = 0; nChart < aCharts.length; ++nChart) {
            oChart = aCharts[nChart];
            if(oChart !== this) {
                oChart = aCharts[nChart];
                bCanMove = false;
                if(bIsSecondaryAxis === oChart.isSecondaryAxis()) {
                    if(bExactType) {
                        if(nType === oChart.getChartType()) {
                            bCanMove = true;
                        }
                    }
                    else {
                        if(oChart.tryChangeType(nType)) {
                            bCanMove = true;
                        }
                    }
                }
                if(bCanMove) {
                    oNewSeries = oChart.getEmptySeries();
                    oSeries.fillObject(oNewSeries);
                    this.removeSeries(this.getSeriesArrayIdx(oSeries));
                    oChart.addSer(oNewSeries);
	                oNewSeries.checkSeriesAfterChangeType(nType);
                    return true;
                }
            }
        }
        return false;
    };
    CChartBase.prototype.tryCreateNewChartFormSeries = function(oSeries, nType, bIsSecondaryAxis) {
        var oPlotArea = this.parent;
        if(!oPlotArea) {
            return Asc.c_oAscError.ID.No;
        }
        var oNewChart = null;
        var aNewAxes = [];
        var nResult = Asc.c_oAscError.ID.No;
        var oChartForAxes = null;
        if(oPlotArea.isPieType(nType) || oPlotArea.isDoughnutType(nType)) {
            if(oPlotArea.isPieType(nType)) {
                oNewChart = oPlotArea.createPieChart(nType, [oSeries], this);
            }
            else {
                oNewChart = oPlotArea.createDoughnutChart(nType, [oSeries], this);
            }
        }
        else {
            oChartForAxes = oPlotArea.getChartWithSuitableAxes(nType, bIsSecondaryAxis);
            if(oChartForAxes) {
                aNewAxes = oChartForAxes.axId;
            }
            else {
                var bCreateAxes = false;
                var aChartsSAxes = oPlotArea.getChartsWithSecondaryAxis();
                if(aChartsSAxes.length > 0) {
                    if(aChartsSAxes.length === 1 && aChartsSAxes[0] === this && this.series.length === 1) {
                        this.removeSeries(this.getSeriesArrayIdx(oSeries));
                        bCreateAxes = true;
                    }
                    else {
                        nResult = Asc.c_oAscError.ID.SecondaryAxis;
                    }
                }
                else {
                    bCreateAxes = true;
                }
                if(bCreateAxes) {
                    if(oPlotArea.isScatterType(nType)) {
                        aNewAxes = oPlotArea.createScatterAxes(true);
                    }
                    else if(oPlotArea.isHBarType(nType)) {
                        aNewAxes = oPlotArea.createHBarAxes(oPlotArea.getAxisNumFormatByType(nType, [oSeries]), true);
                    }
					else if(oPlotArea.isRadarType(nType)) {
	                    aNewAxes = oPlotArea.createRadarAxes(oPlotArea.getAxisNumFormatByType(nType, [oSeries]), true)
                    }
                    else {
						let aCharts = oPlotArea.charts;
						let bAsSecondary = true;
						for(let nChart = 0; nChart < aCharts.length; ++ nChart) {
							if(aCharts[nChart].getObjectType() === AscDFH.historyitem_type_RadarChart) {
								bAsSecondary = false;
								break;
							}
						}
						const bIsArea = oPlotArea.isAreaType(nType)
	                    aNewAxes = oPlotArea.createRegularAxes(oPlotArea.getAxisNumFormatByType(nType, [oSeries]), bAsSecondary, bIsArea);
					}
					oPlotArea.addAxes(aNewAxes);
                }
            }
            if(aNewAxes.length > 0) {
                if(oPlotArea.isHBarType(nType) || oPlotArea.isBarType(nType)) {
                    oNewChart = oPlotArea.createBarChart(nType, [oSeries], aNewAxes, this);
                }
                else if(oPlotArea.isLineType(nType)) {
                    oNewChart = oPlotArea.createLineChart(nType, [oSeries], aNewAxes, this);
                }
                else if(oPlotArea.isAreaType(nType)) {
                    oNewChart = oPlotArea.createAreaChart(nType, [oSeries], aNewAxes, this);
                }
                else if(oPlotArea.isScatterType(nType)) {
                    oNewChart = oPlotArea.createScatterChart(nType, [oSeries], aNewAxes, this);
                }
                else if(oPlotArea.isStockChart(nType)) {
                    oNewChart = oPlotArea.createStockChart(nType, [oSeries], aNewAxes, this);
                }
                else if(oPlotArea.isRadarChart(nType)) {
                    oNewChart = oPlotArea.createRadarChart(nType, [oSeries], aNewAxes, this);
                }
            }
        }
        if(oNewChart) {
            this.removeSeries(this.getSeriesArrayIdx(oSeries));
            oPlotArea.addChart(oNewChart, null);
            nResult = Asc.c_oAscError.ID.No;
        }
        return nResult;
    };
    CChartBase.prototype.tryChangeSeriesChartType = function(oSeries, nType) {
        if(!this.parent) {
            return Asc.c_oAscError.ID.No;
        }
        if(this.tryChangeType(nType)) {
            return Asc.c_oAscError.ID.No;
        }
        var bIsSecondaryAxis = this.isSecondaryAxis();
        if(this.tryMoveSeries(oSeries, nType, bIsSecondaryAxis, false)) {
            return Asc.c_oAscError.ID.No;
        }
        var nRes = this.tryCreateNewChartFormSeries(oSeries, nType, bIsSecondaryAxis);
        if(nRes === Asc.c_oAscError.ID.No) {
            return nRes;
        }
        if(this.tryMoveSeries(oSeries, nType, !bIsSecondaryAxis, true)) {
            return Asc.c_oAscError.ID.No;
        }
        return this.tryCreateNewChartFormSeries(oSeries, nType, !bIsSecondaryAxis);
    };
    CChartBase.prototype.tryChangeSeriesAxisType = function(oSeries, bIsSecondaryAxis) {
        if(this.isSecondaryAxis() === bIsSecondaryAxis) {
            return Asc.c_oAscError.ID.No;
        }
        if(!this.parent) {
            return Asc.c_oAscError.ID.No;
        }
        var nChartType = this.getChartType();
        if(this.tryMoveSeries(oSeries, nChartType, bIsSecondaryAxis, false)) {
            return Asc.c_oAscError.ID.No;
        }
        return this.tryCreateNewChartFormSeries(oSeries, this.getChartType(), bIsSecondaryAxis);
    };
    CChartBase.prototype.tryChangeType = function(nNewType) {
        if(!this.parent) {
            return false;
        }
        return this.getChartType() === nNewType;
    };
    CChartBase.prototype.remove = function() {
        if(!this.parent) {
            return;
        }
        var nChart;
        var aCharts = this.parent.charts;
        for(nChart = 0; nChart < aCharts.length; ++nChart) {
            if(aCharts[nChart] === this) {
                this.parent.removeChartByPos(nChart);
                return;
            }
        }
    };
    CChartBase.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(oAdditionalData && bReset) {
            if(!oAdditionalData.dLbls) {
                this.setDLbls(null);
            }
            else {
                this.setDLbls(oAdditionalData.dLbls.createDuplicate());
            }
			if(this.getObjectType() === AscDFH.historyitem_type_BarChart) {
				this.setGapWidth(oAdditionalData.gapWidth);
				this.setOverlap(oAdditionalData.overlap);
				this.setGapDepth(oAdditionalData.gapDepth);
			}
        }
        for(var nSeries = 0; nSeries < this.series.length; ++nSeries) {
            this.series[nSeries].applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.dLbls) {
            this.dLbls.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
    };
    CChartBase.prototype.getMaxSeriesIdx = function() {
        if(!this.parent) {
            return -1;
        }
        return this.parent.getMaxSeriesIdx();
    };
    function CBarChart() {
        CChartBase.call(this);
        this.barDir = null;
        this.gapWidth = null;
        this.grouping = null;
        this.overlap = null;
        this.serLines = null;

        this.b3D = null;
        this.gapDepth = null;
        this.shape = null;
    }

    InitClass(CBarChart, CChartBase, AscDFH.historyitem_type_BarChart);
    CBarChart.prototype.set3D = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BarChart_Set3D, this.b3D, pr));
        this.b3D = pr;
    };
    CBarChart.prototype.setGapDepth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetGapDepth, this.gapDepth, pr));
        this.gapDepth = pr;
    };
    CBarChart.prototype.setShape = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetShape, this.shape, pr));
        this.shape = pr;
    };
    CBarChart.prototype.getEmptySeries = function() {
        return new CBarSeries();
    };
    CBarChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(AscFormat.isRealNumber(this.barDir) && oCopy.setBarDir)
            oCopy.setBarDir(this.barDir);
        if(AscFormat.isRealNumber(this.gapWidth) && oCopy.setGapWidth)
            oCopy.setGapWidth(this.gapWidth);
        if(AscFormat.isRealNumber(this.gapDepth) && oCopy.setGapDepth)
            oCopy.setGapDepth(this.gapDepth);
        if(AscFormat.isRealNumber(this.shape) && oCopy.setShape)
            oCopy.setShape(this.shape);
        if(AscFormat.isRealNumber(this.grouping) && oCopy.setGrouping)
            oCopy.setGrouping(this.grouping);
        if(AscFormat.isRealNumber(this.overlap) && oCopy.setOverlap)
            oCopy.setOverlap(this.overlap);
        if(AscFormat.isRealBool(this.b3D) && oCopy.getObjectType() === AscDFH.historyitem_type_BarChart) {
            oCopy.set3D(this.b3D);
        }
    };
    CBarChart.prototype.getDefaultDataLabelsPosition = function() {
        if(!AscFormat.isRealNumber(this.grouping) || this.grouping === AscFormat.BAR_GROUPING_CLUSTERED || this.grouping === AscFormat.BAR_GROUPING_STANDARD) {
            return c_oAscChartDataLabelsPos.outEnd;
        }
        return c_oAscChartDataLabelsPos.ctr;
    };
    CBarChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_BarChart_AddAxId:
            case AscDFH.historyitem_CommonChart_AddAxId:
            {
                break;
            }
            case AscDFH.historyitem_BarChart_SetBarDir:
            case AscDFH.historyitem_BarChart_Set3D:
            case AscDFH.historyitem_BarChart_SetGapDepth:
            case AscDFH.historyitem_BarChart_SetShape:
            {
                this.onChartInternalUpdate();
                break;
            }
            case AscDFH.historyitem_BarChart_SetDLbls:
            case AscDFH.historyitem_CommonChart_SetDlbls:
            {
                this.onChartUpdateDataLabels();
                break;
            }
            case AscDFH.historyitem_BarChart_SetGapWidth:
            {
                break;
            }
            case AscDFH.historyitem_BarChart_SetGrouping:
            {
                this.onChartInternalUpdate();
                break;
            }
            case AscDFH.historyitem_BarChart_SetOverlap:
            {
                break;
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_BarChart_SetSerLines:
            {
                break;
            }
            case AscDFH.historyitem_BarChart_SetVaryColors:
            case AscDFH.historyitem_CommonChart_SetVaryColors:
            {
                break;
            }
        }

    };
    CBarChart.prototype.setBarDir = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetBarDir, this.barDir, pr));
        this.barDir = pr;
        this.onChartInternalUpdate();
    };
    CBarChart.prototype.setGapWidth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetGapWidth, this.gapWidth, pr));
        this.gapWidth = pr;
    };
    CBarChart.prototype.setGrouping = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetGrouping, this.grouping, pr));
        this.grouping = pr;
        this.onChartInternalUpdate();
    };
    CBarChart.prototype.setOverlap = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarChart_SetOverlap, this.overlap, pr));
        this.overlap = pr;
    };
    CBarChart.prototype.setSerLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarChart_SetSerLines, this.serLines, pr));
        this.serLines = pr;
        this.setParentToChild(pr);
    };
    CBarChart.prototype.getChartType = function() {
        var nType = Asc.c_oAscChartTypeSettings.unknown;
        var bHBar = (this.barDir === AscFormat.BAR_DIR_BAR);
        var oCS = this.getChartSpace();
        var b3D = false;
        if(oCS && AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(oCS)) {
            b3D = true;
        }
        if(bHBar) {
            switch(this.grouping) {
                case AscFormat.BAR_GROUPING_CLUSTERED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.hBarNormal3d : Asc.c_oAscChartTypeSettings.hBarNormal;
                    break;
                }
                case AscFormat.BAR_GROUPING_STACKED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.hBarStacked3d : Asc.c_oAscChartTypeSettings.hBarStacked;
                    break;
                }
                case AscFormat.BAR_GROUPING_PERCENT_STACKED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.hBarStackedPer3d : Asc.c_oAscChartTypeSettings.hBarStackedPer;
                    break;
                }
                default:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.hBarNormal3d : Asc.c_oAscChartTypeSettings.hBarNormal;
                    break;
                }
            }
        }
        else {
            switch(this.grouping) {
                case AscFormat.BAR_GROUPING_CLUSTERED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.barNormal3d : Asc.c_oAscChartTypeSettings.barNormal;
                    break;
                }
                case AscFormat.BAR_GROUPING_STACKED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.barStacked3d : Asc.c_oAscChartTypeSettings.barStacked;
                    break;
                }
                case AscFormat.BAR_GROUPING_PERCENT_STACKED:
                {
                    nType = b3D ? Asc.c_oAscChartTypeSettings.barStackedPer3d : Asc.c_oAscChartTypeSettings.barStackedPer;
                    break;
                }
                default:
                {
                    if(AscFormat.BAR_GROUPING_STANDARD && b3D) {
                        nType = Asc.c_oAscChartTypeSettings.barNormal3dPerspective;
                    }
                    else {
                        nType = Asc.c_oAscChartTypeSettings.barNormal;
                    }
                    break;
                }
            }
        }
        return nType;
    };
    CBarChart.prototype.tryChangeType = function(nType) {
        if(this.getChartType() === nType) {
            return true;
        }
        if(!this.parent) {
            return false;
        }
        if(this.isHBar() && !this.parent.isHBarType(nType)
            || !this.isHBar() && !this.parent.isBarType(nType)) {
            return false;
        }
        var nNewBarDir = this.parent.getBarDirByType(nType);
        var bChangedGrouping = false;
        var nOldGrouping = this.grouping;
        var nNewGrouping = this.parent.getBarGroupingByType(nType);
        if(this.grouping !== nNewGrouping) {
            this.setGrouping(nNewGrouping);
            bChangedGrouping = true;
        }
        if(!AscFormat.isRealNumber(this.gapWidth)) {
            this.setGapWidth(150);
        }
        if(AscFormat.BAR_GROUPING_PERCENT_STACKED === nNewGrouping
            || AscFormat.BAR_GROUPING_STACKED === nNewGrouping) {
            if(!AscFormat.isRealNumber(this.overlap)
                || nOldGrouping !== AscFormat.BAR_GROUPING_PERCENT_STACKED
                || nOldGrouping !== AscFormat.BAR_GROUPING_STACKED) {
                this.setOverlap(100);
            }
        }
        else {
            if(bChangedGrouping && this.overlap !== null) {
                this.setOverlap(null);
            }
        }
        var isNew3D = this.parent.getIs3dByType(nType);
        this.parent.check3DOptions(nType);
        if(isNew3D) {
            if(this.b3D !== true) {
                this.set3D(true);
            }
			this.setGapDepth(null);  
			this.setOverlap(null);
        }
        else {
            if(this.b3D !== false) {
                this.set3D(false);
            }
        }
        if(this.barDir !== nNewBarDir) {
            this.setBarDir(nNewBarDir);
        }
        this.checkValAxesFormatByType(nType);
        return true;
    };

    function CAreaChart() {
        CChartBase.call(this);
        this.dropLines = null;
        this.grouping = null;
    }

    InitClass(CAreaChart, CChartBase, AscDFH.historyitem_type_AreaChart);
    CAreaChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.ctr;
    };
    CAreaChart.prototype.getEmptySeries = function() {
        return new CAreaSeries();
    };
    CAreaChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_AreaChart_AddAxId:
            case AscDFH.historyitem_CommonChart_AddAxId:
            {
                break
            }
            case AscDFH.historyitem_AreaChart_SetDLbls:
            case AscDFH.historyitem_CommonChart_SetDlbls:
            {
                this.onChartUpdateDataLabels();
                break
            }
            case AscDFH.historyitem_AreaChart_SetDropLines:
            {
                break
            }
            case AscDFH.historyitem_AreaChart_SetGrouping:
            {
                this.onChartInternalUpdate();
                break
            }
            case AscDFH.historyitem_AreaChart_SetVaryColors:
            case AscDFH.historyitem_CommonChart_SetVaryColors:
            {
                break
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
        }
    };
    CAreaChart.prototype.getChildren = function() {
        var aRet = CChartBase.prototype.getChildren.call(this);
        aRet.push(this.dropLines);
        return aRet;
    };
    CAreaChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.dropLines && oCopy.setDropLines)
            oCopy.setDropLines(this.dropLines.createDuplicate());
        if(AscFormat.isRealNumber(this.grouping) && oCopy.setGrouping)
            oCopy.setGrouping(this.grouping);
    };
    CAreaChart.prototype.setDropLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaChart_SetDropLines, this.dropLines, pr));
        this.dropLines = pr;
        this.setParentToChild(pr);
    };
    CAreaChart.prototype.setGrouping = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_AreaChart_SetGrouping, this.grouping, pr));
        this.grouping = pr;
        this.onChartInternalUpdate();
    };
    CAreaChart.prototype.getChartType = function() {
        var nType = Asc.c_oAscChartTypeSettings.unknown;
        switch(this.grouping) {
            case AscFormat.GROUPING_PERCENT_STACKED:
            {
                nType = Asc.c_oAscChartTypeSettings.areaStackedPer;
                break;
            }
            case AscFormat.GROUPING_STACKED:
            {
                nType = Asc.c_oAscChartTypeSettings.areaStacked;
                break;
            }
            default:
            {
                nType = Asc.c_oAscChartTypeSettings.areaNormal;
                break;
            }
        }
        return nType;
    };
    CAreaChart.prototype.tryChangeType = function(nType) {
        if(!this.parent) {
            return false;
        }
        if(!this.parent.isAreaType(nType)) {
            return false;
        }
        if(nType === this.getChartType()) {
            return true;
        }
        var nNewGrouping;
        nNewGrouping = this.parent.getGroupingByType(nType);
        if(this.grouping !== nNewGrouping) {
            this.setGrouping(nNewGrouping);
        }
        this.checkValAxesFormatByType(nType);
        this.parent.check3DOptions(nType);
        return true;
    };

    function CAreaSeries() {
        CSeriesBase.call(this);
        this.cat = null;
        this.dLbls = null;
        this.dPt = [];
        this.errBars = null;
        this.pictureOptions = null;
        this.trendline = null;
        this.val = null;
    }

    InitClass(CAreaSeries, CSeriesBase, AscDFH.historyitem_type_AreaSeries);
    CAreaSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        aRet.push(this.dLbls);
        aRet.push(this.errBars);
        aRet.push(this.pictureOptions);
        aRet.push(this.trendline);
        for(var nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
            aRet.push(this.dPt[nDpt]);
        }
        return aRet;
    };
    CAreaSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.errBars && oCopy.setErrBars)
            oCopy.setErrBars(this.errBars.createDuplicate());
        if(this.pictureOptions && oCopy.setPictureOptions)
            oCopy.setPictureOptions(this.pictureOptions.createDuplicate());
        if(this.trendline && oCopy.setTrendline)
            oCopy.setTrendline(this.trendline.createDuplicate());
    };
    CAreaSeries.prototype.setCat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetCat, this.cat, pr));
        this.cat = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_AreaSeries_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.setErrBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetErrBars, this.errBars, pr));
        this.errBars = pr;
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.setPictureOptions = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetPictureOptions, this.pictureOptions, pr));
        this.pictureOptions = pr;
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.setTrendline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CAreaSeries.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_AreaSeries_SetVal, this.val, pr));
        this.val = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };

    var AX_POS_L = 0;
    var AX_POS_T = 1;
    var AX_POS_R = 2;
    var AX_POS_B = 3;

    var CROSSES_AUTO_ZERO = 0;
    var CROSSES_MAX = 1;
    var CROSSES_MIN = 2;

    var LBL_ALG_CTR = 0;
    var LBL_ALG_L = 1;
    var LBL_ALG_R = 2;

    var TIME_UNIT_DAYS = 0;
    var TIME_UNIT_MONTHS = 1;
    var TIME_UNIT_YEARS = 2;

    var CROSS_BETWEEN_BETWEEN = 0;
    var CROSS_BETWEEN_MID_CAT = 1;

    function CAxisBase() {
        CBaseChartObject.call(this);
        this.axId = null;
        this.scaling = null;
        this.bDelete = null;
        this.axPos = null;
        this.majorGridlines = null;
        this.minorGridlines = null;
        this.title = null;
        this.numFmt = null;
        this.majorTickMark = null;
        this.minorTickMark = null;
        this.tickLblPos = null;
        this.spPr = null;
        this.txPr = null;
        this.crossAx = null;
        this.crosses = null;
        this.crossesAt = null;


        this.bDelete = false;
        this.majorTickMark = c_oAscTickMark.TICK_MARK_OUT;
        this.minorTickMark = c_oAscTickMark.TICK_MARK_NONE;
        this.crosses = AscFormat.CROSSES_AUTO_ZERO;
    }

    InitClass(CAxisBase, CBaseChartObject, AscDFH.historyitem_type_Unknown);
    CAxisBase.prototype.getPlotArea = function() {
        return this.parent;
    };
    CAxisBase.prototype.onUpdate = function() {
        this.onChartInternalUpdate();
    };
    CAxisBase.prototype.Refresh_RecalcData = function() {
        this.onUpdate();
    };
    CAxisBase.prototype.Refresh_RecalcData2 = function(pageIndex, object) {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.Refresh_RecalcData2(pageIndex, object);
        }
    };
    CAxisBase.prototype.getDrawingDocument = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getDrawingDocument();
        }
        return null;
    };
    CAxisBase.prototype.handleUpdateFill = function() {
        this.onUpdate();
    };
    CAxisBase.prototype.handleUpdateLn = function() {
        this.onUpdate();
    };
    CAxisBase.prototype.getParentObjects = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            return oChartSpace.getParentObjects();
        }
        return null;
    };
    CAxisBase.prototype.internalSetAxId = function(pr) {
        this.axId = pr;
        if(AscFormat.isRealNumber(pr)) {
            AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER = Math.max(pr, AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
        }
    };
    CAxisBase.prototype.setAxId = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetAxId, this.axId, pr));
        this.internalSetAxId(pr);
    };
    CAxisBase.prototype.setScaling = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetScaling, this.scaling, pr));
        this.scaling = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setDelete = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_CatAxSetDelete, this.bDelete, pr));
        this.bDelete = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setAxPos = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetAxPos, this.axPos, pr));
        this.axPos = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setMajorGridlines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetMajorGridlines, this.majorGridlines, pr));
        this.majorGridlines = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setMinorGridlines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetMinorGridlines, this.minorGridlines, pr));
        this.minorGridlines = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setTitle = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetTitle, this.title, pr));
        this.title = pr;
        this.setParentToChild(pr);
        this.onUpdate();
        this.onChangeDataRefs();
    };
    CAxisBase.prototype.setNumFmt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetNumFmt, this.numFmt, pr));
        this.numFmt = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setMajorTickMark = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetMajorTickMark, this.majorTickMark, pr));
        this.majorTickMark = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setMinorTickMark = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetMinorTickMark, this.minorTickMark, pr));
        this.minorTickMark = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CAxisBase.prototype.setTickLblPos = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetTickLblPos, this.tickLblPos, pr));
        this.tickLblPos = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setCrossAx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_CatAxSetCrossAx, this.crossAx, pr));
        this.crossAx = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setCrosses = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetCrosses, this.crosses, pr));
        this.crosses = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.setCrossesAt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_CatAxSetCrossesAt, this.crossesAt, pr));
        this.crossesAt = pr;
        this.onUpdate();
    };
    CAxisBase.prototype.getChildren = function() {
        return [this.majorGridlines,
            this.minorGridlines,
            this.numFmt,
            this.scaling,
            this.spPr,
            this.title,
            this.txPr];
    };
    CAxisBase.prototype.mergeBase = function(oAxis) {
        if(AscFormat.isRealNumber(oAxis.axPos)) {
            this.setAxPos(oAxis.axPos);
        }
        if(AscFormat.isRealNumber(oAxis.crosses)) {
            this.setCrosses(oAxis.crosses);
        }
        if(AscFormat.isRealNumber(oAxis.crossesAt)) {
            this.setCrossesAt(oAxis.crossesAt);
        }
        if(AscFormat.isRealBool(oAxis.bDelete)) {
            this.setDelete(oAxis.bDelete);
        }
        if(AscCommon.isRealObject(oAxis.majorGridlines)) {
            this.setMajorGridlines(oAxis.majorGridlines.createDuplicate());
        }
        if(AscFormat.isRealNumber(oAxis.majorTickMark)) {
            this.setMajorTickMark(oAxis.majorTickMark);
        }
        if(AscCommon.isRealObject(oAxis.minorGridlines)) {
            this.setMinorGridlines(oAxis.minorGridlines.createDuplicate());
        }
        if(AscFormat.isRealNumber(oAxis.minorTickMark)) {
            this.setMinorTickMark(oAxis.minorTickMark);
        }
        if(AscCommon.isRealObject(oAxis.numFmt)) {
            this.setNumFmt(oAxis.numFmt.createDuplicate());
        }
        if(AscCommon.isRealObject(oAxis.scaling)) {
            this.setScaling(oAxis.scaling.createDuplicate());
        }
        if(AscCommon.isRealObject(oAxis.spPr)) {
            this.setSpPr(oAxis.spPr.createDuplicate());
        }
        if(AscFormat.isRealNumber(oAxis.tickLblPos)) {
            this.setTickLblPos(oAxis.tickLblPos);
        }
        if(AscCommon.isRealObject(oAxis.title)) {
            this.setTitle(oAxis.title.createDuplicate());
        }
        if(AscCommon.isRealObject(oAxis.txPr)) {
            this.setTxPr(oAxis.txPr.createDuplicate());
        }
    };
    CAxisBase.prototype.merge = function(oAxis) {
        this.mergeBase(oAxis);
    };
    CAxisBase.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.merge(this);
    };
    CAxisBase.prototype.getSourceFormatCode = function() {
        return "General";
    };
    CAxisBase.prototype.updateNumFormat = function() {
        var oNumFmt = this.numFmt;
        if(!oNumFmt) {
            oNumFmt = new CNumFmt();
            oNumFmt.setSourceLinked(true);
            oNumFmt.setFormatCode("General");
            this.setNumFmt(oNumFmt)
        }
        if(oNumFmt.sourceLinked) {
            oNumFmt.setFormatCode(this.getSourceFormatCode());
        }
    };
    CAxisBase.prototype.isHorizontal = function() {
        return this.axPos === AX_POS_B || this.axPos === AX_POS_T;
    };
    CAxisBase.prototype.isVertical = function() {
        return this.axPos === AX_POS_L || this.axPos === AX_POS_R;
    };
    CAxisBase.prototype.getGridlinesSetting = function() {
        if(!this.majorGridlines && !this.minorGridlines)
            return Asc.c_oAscGridLinesSettings.none;
        if(this.majorGridlines && !this.minorGridlines)
            return Asc.c_oAscGridLinesSettings.major;
        if(this.minorGridlines && !this.majorGridlines)
            return Asc.c_oAscGridLinesSettings.minor;
        return Asc.c_oAscGridLinesSettings.majorMinor;
    };
    CAxisBase.prototype.getLabelSetting = function() {
        if(this.isHorizontal()) {
            if(this.title) {
                return Asc.c_oAscChartHorAxisLabelShowSettings.noOverlay;
            }
            else {
                return Asc.c_oAscChartHorAxisLabelShowSettings.none;
            }
        }
        else {
            if(this.title) {
                var tx_body;
                if(this.title.tx && this.title.tx.rich) {
                    tx_body = this.title.tx.rich;
                }
                else if(this.title.txPr) {
                    tx_body = this.title.txPr;
                }
                if(tx_body) {
                    var oBodyPr = this.title.getBodyPr();
                    if(oBodyPr && oBodyPr.vert === AscFormat.nVertTThorz) {
                        return Asc.c_oAscChartVertAxisLabelShowSettings.horizontal;
                    }
                    else {
                        return Asc.c_oAscChartVertAxisLabelShowSettings.rotated;
                    }
                }
                else {
                    return Asc.c_oAscChartVertAxisLabelShowSettings.none;
                }
            }
            else {
                return Asc.c_oAscChartVertAxisLabelShowSettings.none;
            }
        }
    };
    CAxisBase.prototype.setGridlinesSetting = function(gridinesSettings) {
        let bSetMajor = false;
        let bSetMinor = false;
        switch(gridinesSettings) {
            case Asc.c_oAscGridLinesSettings.none:
            {
                if(this.majorGridlines) {
                    this.setMajorGridlines(null);
                }
                if(this.minorGridlines) {
                    this.setMinorGridlines(null);
                }
                break;
            }
            case Asc.c_oAscGridLinesSettings.major:
            {
                if(!this.majorGridlines) {
                    this.setMajorGridlines(new AscFormat.CSpPr());
                    bSetMajor = true;
                }
                if(this.minorGridlines) {
                    this.setMinorGridlines(null);
                }
                break;
            }
            case Asc.c_oAscGridLinesSettings.minor:
            {
                if(!this.minorGridlines) {
                    this.setMinorGridlines(new AscFormat.CSpPr());
                    bSetMinor = true;
                }
                if(this.majorGridlines) {
                    this.setMajorGridlines(null);
                }
                break;
            }
            case Asc.c_oAscGridLinesSettings.majorMinor:
            {
                if(!this.minorGridlines) {
                    this.setMinorGridlines(new AscFormat.CSpPr());
                    bSetMinor = true;
                }
                if(!this.majorGridlines) {
                    this.setMajorGridlines(new AscFormat.CSpPr());
                    bSetMajor = true;
                }
                break;
            }
        }
        if(bSetMajor || bSetMinor) {
            const oChartSpace = this.getChartSpace();
            if(oChartSpace) {
                let oSpPr = this.spPr;
                oChartSpace.checkElementChartStyle(this);
                this.setSpPr(oSpPr);
            }
        }
    };
    CAxisBase.prototype.setLabelSetting = function(labelSetting) {
        var _text_body;
        if(labelSetting !== null) {
            if(this.isHorizontal()) {
                switch(labelSetting) {
                    case Asc.c_oAscChartHorAxisLabelShowSettings.none:
                    {
                        if(this.title)
                            this.setTitle(null);
                        break;
                    }
                    case Asc.c_oAscChartHorAxisLabelShowSettings.noOverlay:
                    {
                        if(this.title && this.title.tx && this.title.tx.rich) {
                            _text_body = this.title.tx.rich;
                        }
                        else {
                            if(!this.title) {
                                this.setTitle(new AscFormat.CTitle());
                            }
                            if(!this.title.txPr) {
                                this.title.setTxPr(new AscFormat.CTextBody());
                            }
                            if(!this.title.txPr.bodyPr) {
                                this.title.txPr.setBodyPr(new AscFormat.CBodyPr());
                            }
                            if(!this.title.txPr.content) {
                                this.title.txPr.setContent(new AscFormat.CDrawingDocContent(this.title.txPr, this.getDrawingDocument(), 0, 0, 100, 500, false, false, true));
                            }
                            _text_body = this.title.txPr;
                        }
                        if(this.title.overlay !== false) {
                            this.title.setOverlay(false);
                        }

                        if(!_text_body.bodyPr || _text_body.bodyPr.isNotNull()) {
                            _text_body.setBodyPr(new AscFormat.CBodyPr());
                        }
                        break;
                    }
                }
            }
            else {
                switch(labelSetting) {
                    case Asc.c_oAscChartVertAxisLabelShowSettings.none:
                    {
                        if(this.title) {
                            this.setTitle(null);
                        }
                        break;
                    }
                    case Asc.c_oAscChartVertAxisLabelShowSettings.vertical:
                    {
                        //TODO: Set this setting when the vertical text will be implemented
                        break;
                    }
                    default:
                    {
                        if(labelSetting === Asc.c_oAscChartVertAxisLabelShowSettings.rotated
                            || labelSetting === Asc.c_oAscChartVertAxisLabelShowSettings.horizontal) {
                            if(this.title && this.title.tx && this.title.tx.rich) {
                                _text_body = this.title.tx.rich;
                            }
                            else {
                                if(!this.title) {
                                    this.setTitle(new AscFormat.CTitle());
                                }
                                if(!this.title.txPr) {
                                    this.title.setTxPr(new AscFormat.CTextBody());
                                }
                                _text_body = this.title.txPr;
                            }
                            if(!_text_body.bodyPr) {
                                _text_body.setBodyPr(new AscFormat.CBodyPr());
                            }
                            var _body_pr = _text_body.bodyPr.createDuplicate();
                            if(!_text_body.content) {
                                _text_body.setContent(new AscFormat.CDrawingDocContent(_text_body, this.getDrawingDocument(), 0, 0, 100, 500, false, false, true));
                            }
                            if(labelSetting === Asc.c_oAscChartVertAxisLabelShowSettings.rotated) {
                                _body_pr.reset();
                            }
                            else {
                                _body_pr.setVert(AscFormat.nVertTThorz);
                                _body_pr.setRot(0);
                            }
                            _text_body.setBodyPr(_body_pr);
                            if(this.title.overlay !== false) {
                                this.title.setOverlay(false);
                            }
                        }
                    }
                }
            }
        }
    };
    CAxisBase.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        var aColors = oColors.generateColors(1);
        if(this.majorGridlines) {
            this.setMajorGridlines(this.getSpPrFormStyleEntry(oChartStyle.gridlineMajor, aColors, 0));
        }
        if(this.minorGridlines) {
            this.setMinorGridlines(this.getSpPrFormStyleEntry(oChartStyle.gridlineMinor, aColors, 0));
        }
        if(this.title) {
            this.title.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
    };
    CAxisBase.prototype.applyAdditionalSettings = function(oAdditionalDataAxis) {
        if(!oAdditionalDataAxis) {
            return;
        }
        this.setDelete(oAdditionalDataAxis.bDelete);
        //this.setAxPos(oAdditionalDataAxis.axPos);
        if(!oAdditionalDataAxis.majorGridlines) {
            this.setMajorGridlines(null);
        }
        else {
            this.setMajorGridlines(oAdditionalDataAxis.majorGridlines.createDuplicate());
        }
        if(!oAdditionalDataAxis.minorGridlines) {
            this.setMinorGridlines(null);
        }
        else {
            this.setMinorGridlines(oAdditionalDataAxis.minorGridlines.createDuplicate());
        }
        if(!oAdditionalDataAxis.title) {
            this.setTitle(null);
        }
        else {
            this.setTitle(oAdditionalDataAxis.title.createDuplicate());
        }
        this.setMajorTickMark(oAdditionalDataAxis.majorTickMark);
        this.setMajorTickMark(oAdditionalDataAxis.minorTickMark);
        this.setTickLblPos(oAdditionalDataAxis.tickLblPos);
        this.setCrosses(oAdditionalDataAxis.crosses);
        this.setCrossesAt(oAdditionalDataAxis.crossesAt);
    };
    CAxisBase.prototype.checkLogScale = function() {
        if(this.scaling && AscFormat.isRealNumber(this.scaling.logBase)) {
            var aPoints = this.yPoints || this.xPoints;
            if(Array.isArray(aPoints) && aPoints.length > 2) {
                aPoints.sort(function(a, b) {
                    return a.val - b.val
                });
                if(aPoints[0].val > 0) {
                    var dLog = this.scaling.logBase;
                    var oFirstPt = aPoints[0];
                    var oLastPt = aPoints[aPoints.length - 1];
                    var oPt;
                    var dMinLog = getBaseLog(dLog, oFirstPt.val);
                    var dMaxLog = getBaseLog(dLog, oLastPt.val);
                    var dDiff = dMaxLog - dMinLog;
                    if(dDiff > 0) {
                        var dScale = (oLastPt.pos - oFirstPt.pos) / dDiff;
                        if(!AscFormat.fApproxEqual(0, dScale)) {
                            for(var nPt = 1; nPt < aPoints.length - 1; ++nPt) {
                                oPt = aPoints[nPt];
                                oPt.pos = oFirstPt.pos + (getBaseLog(dLog, oPt.val) - dMinLog) * dScale;
                            }
                        }
                    }
                }
            }
        }
    };
	CAxisBase.prototype.getAllCharts = function() {
		let oPlotArea = this.parent;
		if(!oPlotArea)
			return [];
		return oPlotArea.getChartsForAxis(this);
	};
	CAxisBase.prototype.isRadarAxis = function() {
		let aCharts = this.getAllCharts();
		for(let nChart = 0; nChart < aCharts.length; ++nChart) {
			if(aCharts[nChart].getObjectType() === AscDFH.historyitem_type_RadarChart) {
				return true;
			}
		}
		return false;
	};
	CAxisBase.prototype.isRadarCategories = function() {
		if(this.isRadarAxis()) {
			let nType = this.getObjectType();
			if(nType === AscDFH.historyitem_type_CatAx ||
				nType === AscDFH.historyitem_type_DateAx) {
				return true;
			}
		}
		return false;
	};

	CAxisBase.prototype.isRadarValues = function() {
		if(this.isRadarAxis()) {
			let nType = this.getObjectType();
			if(nType === AscDFH.historyitem_type_ValAx) {
				return true;
			}
		}
		return false;
	};
	CAxisBase.prototype.isSeriesAxis = function() {
		return this.getObjectType() === AscDFH.historyitem_type_SerAx;
	};

    function getBaseLog(x, y) {
        return Math.log(y) / Math.log(x);
    }

    function CCatAx() {
        CAxisBase.call(this);
        this.auto = null;
        this.extLst = null;
        this.lblAlgn = null;
        this.lblOffset = null;
        this.noMultiLvlLbl = false;
        this.tickLblSkip = null;
        this.tickMarkSkip = null;
    }

    InitClass(CCatAx, CAxisBase, AscDFH.historyitem_type_CatAx);
    CCatAx.prototype.getMenuProps = function() {
        var ret = new AscCommon.asc_CatAxisSettings();

        if(AscFormat.isRealNumber(this.tickMarkSkip))
            ret.putIntervalBetweenTick(this.tickMarkSkip);
        else
            ret.putIntervalBetweenTick(1);


        if(!AscFormat.isRealNumber(this.tickLblSkip))
            ret.putIntervalBetweenLabelsRule(c_oAscBetweenLabelsRule.auto);
        else {
            ret.putIntervalBetweenLabelsRule(c_oAscBetweenLabelsRule.manual);
            ret.putIntervalBetweenLabels(this.tickLblSkip);
        }

        var scaling = this.scaling;
        if(!scaling || scaling.orientation !== ORIENTATION_MAX_MIN)
            ret.putInvertCatOrder(false);
        else
            ret.putInvertCatOrder(true);

        //настройки пересечения с другой осью

        var crossAx = this.crossAx;

        if(crossAx) {
            if(AscFormat.isRealNumber(this.maxCatVal)) {
                ret.putCrossMaxVal(this.maxCatVal);
            }
            ret.putCrossMinVal(1);
            if(AscFormat.isRealNumber(crossAx.crossesAt)) {
                ret.putCrossesRule(c_oAscCrossesRule.value);
                ret.putCrosses(crossAx.crossesAt);
            }
            else if(crossAx.crosses === CROSSES_MAX) {
                ret.putCrossesRule(c_oAscCrossesRule.maxValue);
                if(AscFormat.isRealNumber(this.maxCatVal)) {
                    ret.putCrosses(this.maxCatVal);
                }
            }
            else if(crossAx.crosses === CROSSES_MIN) {
                ret.putCrossesRule(c_oAscCrossesRule.minValue);
                ret.putCrosses(1);
            }
            else {
                ret.putCrossesRule(c_oAscCrossesRule.auto);
                ret.putCrosses(1);
            }
        }

        if(AscFormat.isRealNumber(this.lblOffset))
            ret.putLabelsAxisDistance(this.lblOffset);
        else
            ret.putLabelsAxisDistance(100);

        if(this.crossAx) {
            ret.putLabelsPosition(this.crossAx.crossBetween === CROSS_BETWEEN_MID_CAT ? c_oAscLabelsPosition.byDivisions : c_oAscLabelsPosition.betweenDivisions);
        }
        else {
            ret.putLabelsPosition(c_oAscLabelsPosition.betweenDivisions);
        }
        if(AscFormat.isRealNumber(this.tickLblPos))
            ret.putTickLabelsPos(this.tickLblPos);
        else
            ret.putTickLabelsPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);

        //настройки засечек на оси
        if(AscFormat.isRealNumber(this.majorTickMark))
            ret.putMajorTickMark(this.majorTickMark);
        else
            ret.putMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);

        if(AscFormat.isRealNumber(this.minorTickMark))
            ret.putMinorTickMark(this.minorTickMark);
        else
            ret.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);

        ret.putShow(!this.bDelete);
        ret.putGridlines(this.getGridlinesSetting());
        ret.putLabel(this.getLabelSetting());
        ret.putNumFmt(new AscCommon.asc_CAxNumFmt(this));
        ret.putAuto(this.auto !== false);
	    ret.putIsRadarAxis(this.isRadarAxis());
        return ret;
    };
    CCatAx.prototype.setMenuProps = function(props) {
        if(!(isRealObject(props)
            && typeof props.getAxisType === "function"
            && (props.getAxisType() === c_oAscAxisType.cat || props.getAxisType() === c_oAscAxisType.date)))
            return;

        var intervalBetweenTick = props.getIntervalBetweenTick();
        var intervalBetweenLabelsRule = props.getIntervalBetweenLabelsRule();
        var intervalBetweenLabels = props.getIntervalBetweenLabels();
        var invertCatOrder = props.getInvertCatOrder();
        var labelsAxisDistance = props.getLabelsAxisDistance();
        var axisType = props.getAxisType();
        var majorTickMark = props.getMajorTickMark();
        var minorTickMark = props.getMinorTickMark();
        var tickLabelsPos = props.getTickLabelsPos();
        var crossesRule = props.getCrossesRule();
        var crosses = props.getCrosses();
        var labelsPosition = props.getLabelsPosition();


        var bChanged = false;
        if(AscFormat.isRealNumber(intervalBetweenTick) && this.tickMarkSkip !== intervalBetweenTick && this.setTickMarkSkip) {
            this.setTickMarkSkip(intervalBetweenTick);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(intervalBetweenLabelsRule) && this.setTickLblSkip) {
            if(intervalBetweenLabelsRule === c_oAscBetweenLabelsRule.auto) {
                if(AscFormat.isRealNumber(this.tickLblSkip)) {
                    this.setTickLblSkip(null);
                    bChanged = true;
                }
            }
            else if(intervalBetweenLabelsRule === c_oAscBetweenLabelsRule.manual && AscFormat.isRealNumber(intervalBetweenLabels) && this.tickLblSkip !== intervalBetweenLabels) {
                this.setTickLblSkip(intervalBetweenLabels);
                bChanged = true;
            }
        }

        if(!this.scaling)
            this.setScaling(new CScaling());
        var scaling = this.scaling;
        if(AscFormat.isRealBool(invertCatOrder)) {
            var new_orientation = invertCatOrder ? ORIENTATION_MAX_MIN : ORIENTATION_MIN_MAX;
            if(scaling.orientation !== new_orientation) {
                scaling.setOrientation(invertCatOrder ? ORIENTATION_MAX_MIN : ORIENTATION_MIN_MAX);
                bChanged = true;
            }
        }
        if(AscFormat.isRealNumber(labelsAxisDistance) && this.lblOffset !== labelsAxisDistance && this.setLblOffset) {
            this.setLblOffset(labelsAxisDistance);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(axisType)) {
            //TODO
        }

        if(AscFormat.isRealNumber(majorTickMark) && this.majorTickMark !== majorTickMark) {
            this.setMajorTickMark(majorTickMark);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(minorTickMark) && this.minorTickMark !== minorTickMark) {
            this.setMinorTickMark(minorTickMark);
            bChanged = true;
        }
        if(bChanged) {
            if(this.spPr && this.spPr.hasNoFillLine()
                && (this.majorTickMark !== c_oAscTickMark.TICK_MARK_NONE || this.minorTickMark !== c_oAscTickMark.TICK_MARK_NONE)) {
                this.spPr.setLn(null);
            }
        }

        if(AscFormat.isRealNumber(tickLabelsPos) && this.tickLblPos !== tickLabelsPos) {
            this.setTickLblPos(tickLabelsPos);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(crossesRule) && isRealObject(this.crossAx)) {
            if(crossesRule === c_oAscCrossesRule.auto) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_AUTO_ZERO) {
                    this.crossAx.setCrosses(CROSSES_AUTO_ZERO);
                    bChanged = true;
                }
            }
            else if(crossesRule === c_oAscCrossesRule.value) {
                if(AscFormat.isRealNumber(crosses)) {
                    if(this.crossAx.crossesAt !== crosses) {
                        this.crossAx.setCrossesAt(crosses >> 0);
                        bChanged = true;
                    }
                    if(this.crossAx !== null) {
                        this.crossAx.setCrosses(null);
                        bChanged = true;
                    }
                }
            }
            else if(crossesRule === c_oAscCrossesRule.maxValue) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_MAX) {
                    this.crossAx.setCrosses(CROSSES_MAX);
                    bChanged = true;
                }
            }
            else if(crossesRule === c_oAscCrossesRule.minValue) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_MIN) {
                    this.crossAx.setCrosses(CROSSES_MIN);
                    bChanged = true;
                }
            }
        }

        if(AscFormat.isRealNumber(labelsPosition) && isRealObject(this.crossAx) && this.crossAx.setCrossBetween) {
            var new_lbl_position = labelsPosition === c_oAscLabelsPosition.byDivisions ? CROSS_BETWEEN_MID_CAT : CROSS_BETWEEN_BETWEEN;
            if(this.crossAx.crossBetween !== new_lbl_position) {
                this.crossAx.setCrossBetween(new_lbl_position);
                bChanged = true;
            }
        }
        if(props.getShow() !== null) {
            var bDelete = !props.getShow();
            if(bDelete !== this.bDelete) {
                this.setDelete(bDelete);
            }
        }
        var nGridlines = props.getGridlines();
        if(nGridlines !== null) {
            this.setGridlinesSetting(nGridlines);
        }
        this.setLabelSetting(props.getLabel());
        var oAscFmt = props.getNumFmt();
        if(oAscFmt && oAscFmt.isCorrect()) {
            var oNumFmt = this.numFmt || new AscFormat.CNumFmt();
            oNumFmt.setFromAscObject(oAscFmt);
            if(!this.numFmt) {
                this.setNumFmt(oNumFmt);
            }
        }
        this.setAuto(props.getAuto());
    };
    CCatAx.prototype.setAuto = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_CatAxSetAuto, this.auto, pr));
        this.auto = pr;
        this.onUpdate();
    };
    CCatAx.prototype.setLblAlgn = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetLblAlgn, this.lblAlgn, pr));
        this.lblAlgn = pr;
        this.onUpdate();
    };
    CCatAx.prototype.setLblOffset = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetLblOffset, this.lblOffset, pr));
        this.lblOffset = pr;
        this.onUpdate();
    };
    CCatAx.prototype.setNoMultiLvlLbl = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_CatAxSetNoMultiLvlLbl, this.noMultiLvlLbl, pr));
        this.noMultiLvlLbl = pr;
        this.onUpdate();
    };
    CCatAx.prototype.setTickLblSkip = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetTickLblSkip, this.tickLblSkip, pr));
        this.tickLblSkip = pr;
        this.onUpdate();
    };
    CCatAx.prototype.setTickMarkSkip = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_CatAxSetTickMarkSkip, this.tickMarkSkip, pr));
        this.tickMarkSkip = pr;
        this.onUpdate();
    };
    CCatAx.prototype.merge = function(oAxis) {
        this.mergeBase(oAxis);
        if(AscFormat.isRealBool(oAxis.auto)) {
            this.setAuto(oAxis.auto);
        }
        if(AscFormat.isRealNumber(oAxis.lblAlgn)) {
            this.setLblAlgn(oAxis.lblAlgn);
        }
        if(AscFormat.isRealNumber(oAxis.lblOffset)) {
            this.setLblOffset(oAxis.lblOffset);
        }
        if(AscFormat.isRealBool(oAxis.noMultiLvlLbl)) {
            this.setNoMultiLvlLbl(oAxis.noMultiLvlLbl);
        }
        if(AscFormat.isRealNumber(oAxis.tickLblSkip)) {
            this.setTickLblSkip(oAxis.tickLblSkip);
        }
        if(AscFormat.isRealNumber(oAxis.tickMarkSkip)) {
            this.setTickMarkSkip(oAxis.tickMarkSkip);
        }
    };
    CCatAx.prototype.getSourceFormatCode = function() {
        var oPlotArea = this.getPlotArea();
        var sDefault = "General";
        if(!oPlotArea) {
            return sDefault;
        }
        var oSeries = oPlotArea.getSeriesWithSmallestIndexForAxis(this);
        if(!oSeries) {
            return sDefault;
        }
        return oSeries.getCatSourceNumFormat();
    };
    CCatAx.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(bReset && oAdditionalData) {
            this.applyAdditionalSettings(oAdditionalData.catAx);
        }
        this.applyStyleEntry(oChartStyle.categoryAxis, oColors.generateColors(1), 0, bReset);
        CAxisBase.prototype.applyChartStyle.call(this, oChartStyle, oColors, oAdditionalData, bReset);
    };

    function CDateAx() {
        CAxisBase.call(this);
        this.auto = null;
        this.baseTimeUnit = null;
        this.extLst = null;
        this.lblOffset = null;
        this.majorTimeUnit = null;
        this.majorUnit = null;
        this.minorTimeUnit = null;
        this.minorUnit = null;
    }

    InitClass(CDateAx, CAxisBase, AscDFH.historyitem_type_DateAx);
    CDateAx.prototype.getMenuProps = function() {
        var oProps = CCatAx.prototype.getMenuProps.call(this);
        oProps.putAxisType(c_oAscAxisType.date);
        return oProps;
    };
    CDateAx.prototype.setMenuProps = CCatAx.prototype.setMenuProps;
    CDateAx.prototype.setAuto = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DateAxAuto, this.auto, pr));
        this.auto = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setBaseTimeUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DateAxBaseTimeUnit, this.baseTimeUnit, pr));
        this.baseTimeUnit = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setLblOffset = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DateAxLblOffset, this.lblOffset, pr));
        this.lblOffset = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setMajorTimeUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DateAxMajorTimeUnit, this.majorTimeUnit, pr));
        this.majorTimeUnit = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setMajorUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DateAxMajorUnit, this.majorUnit, pr));
        this.majorUnit = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setMinorTimeUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DateAxMinorTimeUnit, this.minorTimeUnit, pr));
        this.minorTimeUnit = pr;
        this.onUpdate();
    };
    CDateAx.prototype.setMinorUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DateAxMinorUnit, this.minorUnit, pr));
        this.minorUnit = pr;
        this.onUpdate();
    };
    CDateAx.prototype.merge = function(oAxis) {
        this.mergeBase(oAxis);
        if(AscFormat.isRealNumber(oAxis.baseTimeUnit)) {
            this.setBaseTimeUnit(oAxis.baseTimeUnit);
        }
        if(AscFormat.isRealNumber(oAxis.majorTimeUnit)) {
            this.setMajorTimeUnit(oAxis.majorTimeUnit);
        }
        if(AscFormat.isRealNumber(oAxis.majorUnit)) {
            this.setMajorUnit(oAxis.majorUnit);
        }
        if(AscFormat.isRealNumber(oAxis.minorTimeUnit)) {
            this.setMinorTimeUnit(oAxis.minorTimeUnit);
        }
        if(AscFormat.isRealNumber(oAxis.minorUnit)) {
            this.setMinorUnit(oAxis.minorUnit);
        }
        if(AscFormat.isRealBool(oAxis.auto)) {
            this.setAuto(oAxis.auto);
        }
        if(AscFormat.isRealNumber(oAxis.lblOffset)) {
            this.setLblOffset(oAxis.lblOffset);
        }
    };
    CDateAx.prototype.getSourceFormatCode = function() {
        return CCatAx.prototype.getSourceFormatCode.call(this);
    };
    CDateAx.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(bReset && oAdditionalData) {
            this.applyAdditionalSettings(oAdditionalData.catAx);
        }
        this.applyStyleEntry(oChartStyle.categoryAxis, oColors.generateColors(1), 0, bReset);
        CAxisBase.prototype.applyChartStyle.call(this, oChartStyle, oColors, oAdditionalData, bReset);
    };

    function CSerAx() {
        CAxisBase.call(this);
        this.extLst = null;
        this.tickLblSkip = null;
        this.tickMarkSkip = null;
    }

    InitClass(CSerAx, CAxisBase, AscDFH.historyitem_type_SerAx);
    CSerAx.prototype.getMenuProps = CCatAx.prototype.getMenuProps;
    CSerAx.prototype.setMenuProps = CCatAx.prototype.setMenuProps;
    CSerAx.prototype.setTickLblSkip = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_SerAxSetTickLblSkip, this.tickLblSkip, pr));
        this.tickLblSkip = pr;
        this.onUpdate();
    };
    CSerAx.prototype.setTickMarkSkip = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_SerAxSetTickMarkSkip, this.tickMarkSkip, pr));
        this.tickMarkSkip = pr;
        this.onUpdate();
    };
    CSerAx.prototype.merge = function(oAxis) {
        this.mergeBase(oAxis);
        if(AscFormat.isRealNumber(oAxis.tickLblSkip)) {
            this.setTickLblSkip(oAxis.tickLblSkip);
        }
        if(AscFormat.isRealNumber(oAxis.tickMarkSkip)) {
            this.setTickMarkSkip(oAxis.tickMarkSkip);
        }
    };
    CSerAx.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.seriesAxis, oColors.generateColors(1), 0, bReset);
        CAxisBase.prototype.applyChartStyle.call(this, oChartStyle, oColors, oAdditionalData, bReset);
    };

    function CValAx() {
        CAxisBase.call(this);
        this.crossBetween = null;
        this.majorUnit = null;
        this.minorUnit = null;
        this.dispUnits = null;
        this.extLst = null;
    }

    InitClass(CValAx, CAxisBase, AscDFH.historyitem_type_ValAx);
    CValAx.prototype.setCrossBetween = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ValAxSetCrossBetween, this.crossBetween, pr));
        this.crossBetween = pr;
        this.onUpdate();
    };
    CValAx.prototype.setDispUnits = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ValAxSetDispUnits, this.dispUnits, pr));
        this.dispUnits = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CValAx.prototype.setMajorUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_ValAxSetMajorUnit, this.majorUnit, pr));
        this.majorUnit = pr;
        this.onUpdate();
    };
    CValAx.prototype.setMinorUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_ValAxSetMinorUnit, this.minorUnit, pr));
        this.minorUnit = pr;
        this.onUpdate();
    };
    CValAx.prototype.getMenuProps = function() {
        var ret = new AscCommon.asc_ValAxisSettings();
        var scaling = this.scaling;

        //настройки логарифмической шкалы
        if(scaling && AscFormat.isRealNumber(scaling.logBase)) {
            ret.putLogScale(true);
            ret.putLogBase(scaling.logBase);
        }
        else {
            ret.putLogScale(false);
        }

        var oMinMaxOnAxis;
        if(this.axPos === AX_POS_L || this.axPos === AX_POS_R) {
            oMinMaxOnAxis = getMinMaxFromArrPoints(this.yPoints);
        }
        else {
            oMinMaxOnAxis = getMinMaxFromArrPoints(this.xPoints);
        }
        //настроки максимального значения по оси
        if(scaling && AscFormat.isRealNumber(scaling.max)) {
            ret.putMaxValRule(c_oAscValAxisRule.fixed);
            ret.putMaxVal(scaling.max);
        }
        else {
            ret.putMaxValRule(c_oAscValAxisRule.auto);
            ret.putMaxVal(oMinMaxOnAxis.max);
        }

        //настройки минимального значения по оси
        if(scaling && AscFormat.isRealNumber(scaling.min)) {
            ret.putMinValRule(c_oAscValAxisRule.fixed);
            ret.putMinVal(scaling.min);
        }
        else {
            ret.putMinValRule(c_oAscValAxisRule.auto);
            ret.putMinVal(oMinMaxOnAxis.min);
        }

        //настройка ориентации оси
        ret.putInvertValOrder(scaling && scaling.orientation === ORIENTATION_MAX_MIN);

        //настройка множителя единиц на оси
        if(isRealObject(this.dispUnits)) {
            var disp_units = this.dispUnits;
            if(AscFormat.isRealNumber(disp_units.builtInUnit)) {
                ret.putDispUnitsRule(disp_units.builtInUnit);
                ret.putShowUnitsOnChart(isRealObject(disp_units.dispUnitsLbl));
            }
            else if(AscFormat.isRealNumber(disp_units.custUnit)) {
                ret.putDispUnitsRule(c_oAscValAxUnits.CUSTOM);
                ret.putUnits(disp_units.custUnit);
                ret.putShowUnitsOnChart(isRealObject(disp_units.dispUnitsLbl));
            }
            else {
                ret.putDispUnitsRule(c_oAscValAxUnits.none);
                ret.putShowUnitsOnChart(false);
            }
        }
        else {
            ret.putDispUnitsRule(c_oAscValAxUnits.none);
            ret.putShowUnitsOnChart(false);
        }

        //настройки засечек на оси
        if(AscFormat.isRealNumber(this.majorTickMark))
            ret.putMajorTickMark(this.majorTickMark);
        else
            ret.putMajorTickMark(c_oAscTickMark.TICK_MARK_NONE);

        if(AscFormat.isRealNumber(this.minorTickMark))
            ret.putMinorTickMark(this.minorTickMark);
        else
            ret.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);

        if(AscFormat.isRealNumber(this.tickLblPos))
            ret.putTickLabelsPos(this.tickLblPos);
        else
            ret.putTickLabelsPos(c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);

        var crossAx = this.crossAx;
        if(crossAx) {
            //настройки пересечения с другой осью
            if(AscFormat.isRealNumber(crossAx.crossesAt)) {
                ret.putCrossesRule(c_oAscCrossesRule.value);
                ret.putCrosses(crossAx.crossesAt);
            }
            else if(crossAx.crosses === CROSSES_MAX) {
                ret.putCrossesRule(c_oAscCrossesRule.maxValue);
                ret.putCrosses(oMinMaxOnAxis.max);
            }
            else if(crossAx.crosses === CROSSES_MIN) {
                ret.putCrossesRule(c_oAscCrossesRule.minValue);
                ret.putCrosses(oMinMaxOnAxis.min);
            }
            else {
                ret.putCrossesRule(c_oAscCrossesRule.auto);
                if(AscFormat.isRealNumber(oMinMaxOnAxis.min) && AscFormat.isRealNumber(oMinMaxOnAxis.max)) {
                    if(0 >= oMinMaxOnAxis.min && 0 <= oMinMaxOnAxis.max) {
                        ret.putCrosses(0);
                    }
                    else if(oMinMaxOnAxis.min > 0) {
                        ret.putCrosses(oMinMaxOnAxis.min);
                    }
                    else {
                        ret.putCrosses(oMinMaxOnAxis.max);
                    }
                }
            }
        }
        ret.putShow(!this.bDelete);
        ret.putGridlines(this.getGridlinesSetting());
        ret.putLabel(this.getLabelSetting());
        ret.putNumFmt(new AscCommon.asc_CAxNumFmt(this));
	    ret.putIsRadarAxis(this.isRadarAxis());
        return ret;
    };
    CValAx.prototype.setMenuProps = function(props) {
        var bChanged = false;
        if(!(isRealObject(props) && typeof props.getAxisType === "function" && props.getAxisType() === c_oAscAxisType.val))
            return;
        if(!this.scaling)
            this.setScaling(new CScaling());

        var scaling = this.scaling;
        if(AscFormat.isRealNumber(props.minValRule)) {
            if(props.minValRule === c_oAscValAxisRule.auto) {
                if(AscFormat.isRealNumber(scaling.min)) {
                    scaling.setMin(null);
                    bChanged = true;
                }
            }
            else {
                if(AscFormat.isRealNumber(props.minVal)) {
                    if(!(props.maxValRule === c_oAscValAxisRule.fixed && props.maxVal < props.minVal) && scaling.min !== props.minVal) {
                        scaling.setMin(props.minVal);
                        bChanged = true;
                    }
                }
            }
        }

        if(AscFormat.isRealNumber(props.maxValRule)) {
            if(props.maxValRule === c_oAscValAxisRule.auto) {
                if(AscFormat.isRealNumber(scaling.max)) {
                    scaling.setMax(null);
                    bChanged = true;
                }
            }
            else {
                if(AscFormat.isRealNumber(props.maxVal)) {
                    if(!AscFormat.isRealNumber(scaling.min) || scaling.min < props.maxVal) {
                        if(scaling.max !== props.maxVal) {
                            scaling.setMax(props.maxVal);
                            bChanged = true;
                        }
                    }
                }
            }
        }

        if(AscFormat.isRealBool(props.invertValOrder)) {
            var new_or = props.invertValOrder ? ORIENTATION_MAX_MIN : ORIENTATION_MIN_MAX;
            if(scaling.orientation !== new_or) {
                scaling.setOrientation(new_or);
                bChanged = true;
            }
        }


        if(AscFormat.isRealBool(props.logScale)) {
            if(props.logScale && AscFormat.isRealNumber(props.logBase) && props.logBase >= 2 && props.logBase <= 1000) {
                if(scaling.logBase !== props.logBase) {
                    scaling.setLogBase(props.logBase);
                    bChanged = true;
                }
            }
            else if(!props.logScale && scaling.logBase !== null) {
                scaling.setLogBase(null);
                bChanged = true;
            }
        }

        if(AscFormat.isRealNumber(props.dispUnitsRule)) {
            if(props.dispUnitsRule === c_oAscValAxUnits.none) {
                if(this.dispUnits) {
                    this.setDispUnits(null);
                    bChanged = true;
                }
            }
            else {
                if(!this.dispUnits) {
                    this.setDispUnits(new CDispUnits());
                    bChanged = true;
                }
                if(this.dispUnits.builtInUnit !== props.dispUnitsRule) {
                    this.dispUnits.setBuiltInUnit(props.dispUnitsRule);
                    bChanged = true;
                }
                if(AscFormat.isRealBool(this.showUnitsOnChart)) {
                    this.dispUnits.setDispUnitsLbl(new CDLbl());
                    bChanged = true;
                }
            }
        }

        if(AscFormat.isRealNumber(props.majorTickMark) && this.majorTickMark !== props.majorTickMark) {
            this.setMajorTickMark(props.majorTickMark);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(props.minorTickMark) && this.minorTickMark !== props.minorTickMark) {
            this.setMinorTickMark(props.minorTickMark);
            bChanged = true;
        }
        if(bChanged) {
            if(this.spPr && this.spPr.hasNoFillLine()
                && (this.majorTickMark !== c_oAscTickMark.TICK_MARK_NONE || this.minorTickMark !== c_oAscTickMark.TICK_MARK_NONE)) {
                this.spPr.setLn(null);
            }
        }

        if(AscFormat.isRealNumber(props.tickLabelsPos) && this.tickLblPos !== props.tickLabelsPos) {
            this.setTickLblPos(props.tickLabelsPos);
            bChanged = true;
        }

        if(AscFormat.isRealNumber(props.crossesRule) && isRealObject(this.crossAx)) {
            if(props.crossesRule === c_oAscCrossesRule.auto) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_AUTO_ZERO) {
                    this.crossAx.setCrosses(CROSSES_AUTO_ZERO);
                    bChanged = true;
                }
            }
            else if(props.crossesRule === c_oAscCrossesRule.value) {
                if(AscFormat.isRealNumber(props.crosses)) {
                    if(this.crossAx.crossesAt !== props.crosses) {
                        this.crossAx.setCrossesAt(props.crosses);
                        bChanged = true;
                    }
                    if(this.crossAx.crosses !== null) {
                        this.crossAx.setCrosses(null);
                        bChanged = true;
                    }
                }
            }
            else if(props.crossesRule === c_oAscCrossesRule.maxValue) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_MAX) {
                    this.crossAx.setCrosses(CROSSES_MAX);
                    bChanged = true;
                }
            }
            else if(props.crossesRule === c_oAscCrossesRule.minValue) {
                if(this.crossAx.crossesAt !== null) {
                    this.crossAx.setCrossesAt(null);
                    bChanged = true;
                }
                if(this.crossAx.crosses !== CROSSES_MIN) {
                    this.crossAx.setCrosses(CROSSES_MIN);
                    bChanged = true;
                }
            }
        }
        if(bChanged) {
            if(this.bDelete === true) {
                this.setDelete(false);
            }
        }
        if(props.getShow() !== null) {
            var bDelete = !props.getShow();
            if(bDelete !== this.bDelete) {
                this.setDelete(bDelete);
            }
        }
        var nGridlines = props.getGridlines();
        if(nGridlines !== null) {
            this.setGridlinesSetting(nGridlines);
        }
        this.setLabelSetting(props.getLabel());
        var oAscFmt = props.getNumFmt();
        if(oAscFmt && oAscFmt.isCorrect()) {
            var oNumFmt = this.numFmt || new AscFormat.CNumFmt();
            oNumFmt.setFromAscObject(oAscFmt);
            if(!this.numFmt) {
                this.setNumFmt(oNumFmt);
            }
        }
    };
    CValAx.prototype.getFormatCode = function(oChartSpace, oSeries) {
        var oNumFmt = this.numFmt;
        var sFormatCode = null;

        if(oNumFmt) {
            if(oNumFmt.sourceLinked) {
                return this.getSourceFormatCode();
            }
            sFormatCode = oNumFmt.formatCode;
            if(typeof sFormatCode === "string" && sFormatCode.length > 0) {
                return sFormatCode;
            }
        }
        return "General";
    };
    CValAx.prototype.getSourceFormatCode = function() {
        var oPlotArea = this.getPlotArea();
        var sDefault = "General";
        if(!oPlotArea) {
            return sDefault;
        }
        var oSeries = oPlotArea.getSeriesWithSmallestIndexForAxis(this);
        if(!oSeries) {
            return sDefault;
        }
        if(oSeries.getObjectType() === AscDFH.historyitem_type_ScatterSer && this.isHorizontal()) {
            return oSeries.getCatSourceNumFormat();
        }
        return oSeries.getValSourceNumFormat();
    };
    CValAx.prototype.merge = function(oAxis) {
        if(!oAxis) {
            return;
        }
        this.mergeBase(oAxis);
        if(AscFormat.isRealNumber(oAxis.axPos)) {
            this.setAxPos(oAxis.axPos);
        }
        if(AscFormat.isRealNumber(oAxis.crossBetween)) {
            this.setCrossBetween(oAxis.crossBetween);
        }
        if(AscFormat.isRealNumber(oAxis.majorUnit)) {
            this.setMajorUnit(oAxis.majorUnit);
        }

        if(AscFormat.isRealNumber(oAxis.minorUnit)) {
            this.setMinorUnit(oAxis.minorUnit);
        }
        if(AscCommon.isRealObject(oAxis.dispUnits)) {
            this.setDispUnits(oAxis.dispUnits.createDuplicate())
        }
    };
    CValAx.prototype.checkNumFormat = function(sNewFormat) {
        if(typeof sNewFormat === "string" && sNewFormat.length > 0) {
            if(!this.numFmt) {
                this.setNumFmt(new AscFormat.CNumFmt());
            }
            var oNumFmt = this.numFmt;
            if(oNumFmt.formatCode !== sNewFormat) {
                oNumFmt.setFormatCode(sNewFormat);
            }
            if(oNumFmt.sourceLinked !== true) {
                oNumFmt.setSourceLinked(true);
            }
        }
    };
    CValAx.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(bReset && oAdditionalData) {
            this.applyAdditionalSettings(oAdditionalData.valAx);
        }
        this.applyStyleEntry(oChartStyle.valueAxis, oColors.generateColors(1), 0, bReset);
        CAxisBase.prototype.applyChartStyle.call(this, oChartStyle, oColors, oAdditionalData, bReset);
    };

    function CBandFmt() {
        CBaseChartObject.call(this);
        this.idx = null;
        this.spPr = null;
    }

    InitClass(CBandFmt, CBaseChartObject, AscDFH.historyitem_type_BandFmt);
    CBandFmt.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setIdx(this.idx);
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
    };
    CBandFmt.prototype.getChildren = function() {
        return [this.spPr];
    };
    CBandFmt.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BandFmt_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CBandFmt.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BandFmt_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };

    function CBarSeries() {
        CSeriesBase.call(this);
        this.cat = null;
        this.dLbls = null;
        this.dPt = [];
        this.errBars = null;
        this.invertIfNegative = null;
        this.pictureOptions = null;
        this.shape = null;
        this.trendline = null;
        this.val = null;
    }

    InitClass(CBarSeries, CSeriesBase, AscDFH.historyitem_type_BarSeries);
    CBarSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        aRet.push(this.dLbls);
        for(var nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
            aRet.push(this.dPt[nDpt]);
        }
        aRet.push(this.errBars);
        aRet.push(this.pictureOptions);
        aRet.push(this.trendline);
        return aRet;
    };
    CBarSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(oCopy.setErrBars && this.errBars)
            oCopy.setErrBars(this.errBars.createDuplicate());
        if(oCopy.setInvertIfNegative && AscFormat.isRealBool(this.invertIfNegative))
            oCopy.setInvertIfNegative(this.invertIfNegative);
        if(oCopy.setPictureOptions && this.pictureOptions)
            oCopy.setPictureOptions(this.pictureOptions.createDuplicate());
        if(oCopy.setShape && AscFormat.isRealNumber(this.shape))
            oCopy.setShape(this.shape);
        if(oCopy.setTrendline && this.trendline)
            oCopy.setTrendline(this.trendline.createDuplicate());
    };
    CBarSeries.prototype.setCat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetCat, this.cat, pr));
        this.cat = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_BarSeries_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.setErrBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetErrBars, this.errBars, pr));
        this.errBars = pr;
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.setInvertIfNegative = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BarSeries_SetInvertIfNegative, this.invertIfNegative, pr));
        this.invertIfNegative = pr;
    };
    CBarSeries.prototype.setPictureOptions = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetPictureOptions, this.pictureOptions, pr));
        this.pictureOptions = pr;
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.setShape = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BarSeries_SetShape, this.shape, pr));
        this.shape = pr;
    };
    CBarSeries.prototype.setTrendline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CBarSeries.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BarSeries_SetVal, this.val, pr));
        this.val = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };

    function CBubbleChart() {
        CChartBase.call(this);
        this.bubble3D = null;
        this.bubbleScale = null;
        this.showNegBubbles = null;
        this.sizeRepresents = null;
    }

    InitClass(CBubbleChart, CChartBase, AscDFH.historyitem_type_BubbleChart);
    CBubbleChart.prototype.Refresh_RecalcData = function() {
    };
    CBubbleChart.prototype.getEmptySeries = function() {
        return new CBubbleSeries();
    };
    CBubbleChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.r;
    };
    CBubbleChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.bubble3D !== null && oCopy.setBubble3D) {
            oCopy.setBubble3D(this.bubble3D);
        }
        if(this.bubbleScale !== null && oCopy.setBubbleScale) {
            oCopy.setBubbleScale(this.bubbleScale);
        }
        if(this.showNegBubbles !== null && oCopy.setShowNegBubbles) {
            oCopy.setShowNegBubbles(this.showNegBubbles);
        }
        if(this.sizeRepresents !== null && oCopy.setSizeRepresents) {
            oCopy.setSizeRepresents(this.sizeRepresents);
        }
    };
    CBubbleChart.prototype.setBubble3D = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BubbleChart_SetBubble3D, this.bubble3D, pr));
        this.bubble3D = pr;
    };
    CBubbleChart.prototype.setBubbleScale = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BubbleChart_SetBubbleScale, this.bubbleScale, pr));
        this.bubbleScale = pr;
    };
    CBubbleChart.prototype.setShowNegBubbles = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BubbleChart_SetShowNegBubbles, this.showNegBubbles, pr));
        this.showNegBubbles = pr;
    };
    CBubbleChart.prototype.setSizeRepresents = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_BubbleChart_SetSizeRepresents, this.sizeRepresents, pr));
        this.sizeRepresents = pr;
    };
    CBubbleChart.prototype.convertToScutterChart = function() {
        var oScatter = new AscFormat.CScatterChart();
        oScatter.setScatterStyle(AscFormat.SCATTER_STYLE_LINE_MARKER);
        if (null != this.varyColors)
            oScatter.setVaryColors(this.varyColors);
        if (null != this.dLbls)
            oScatter.setDLbls(this.dLbls);
        for (var i = 0, length = this.series.length; i < length; ++i) {
            var bubbleSer = this.series[i];
            var scatterSer = new AscFormat.CScatterSeries();
            if (null != bubbleSer.idx)
                scatterSer.setIdx(bubbleSer.idx);
            if (null != bubbleSer.order)
                scatterSer.setOrder(bubbleSer.order);
            if (null != bubbleSer.tx)
                scatterSer.setTx(bubbleSer.tx);
            //if (null != bubbleSer.spPr)
            //    scatterSer.setSpPr(bubbleSer.spPr);
            for (var j = 0, length2 = bubbleSer.dPt.length; j < length2; ++j) {
                scatterSer.addDPt(bubbleSer.dPt[j]);
            }
            if (null != bubbleSer.dLbls)
                scatterSer.setDLbls(bubbleSer.dLbls);
            if (null != bubbleSer.trendline)
                scatterSer.setTrendline(bubbleSer.trendline);
            if (null != bubbleSer.errBars)
                scatterSer.setErrBars(bubbleSer.errBars);
            if (null != bubbleSer.xVal)
                scatterSer.setXVal(bubbleSer.xVal);
            if (null != bubbleSer.yVal)
                scatterSer.setYVal(bubbleSer.yVal);
            var spPr = new AscFormat.CSpPr();
            var ln = new AscFormat.CLn();
            ln.setW(28575);
            var uni_fill = new AscFormat.CUniFill();
            uni_fill.setFill(new AscFormat.CNoFill());
            ln.setFill(uni_fill);
            spPr.setLn(ln);
            scatterSer.setSpPr(spPr);
            scatterSer.setSmooth(false);
            oScatter.addSer(scatterSer);
        }
        return oScatter;
    };

    function CBubbleSeries() {
        CSeriesBase.call(this);
        this.bubble3D = null;
        this.bubbleSize = null;
        this.dLbls = null;
        this.dPt = [];
        this.errBars = null;
        this.invertIfNegative = null;
        this.trendline = null;
        this.xVal = null;
        this.yVal = null;
    }

    InitClass(CBubbleSeries, CSeriesBase, AscDFH.historyitem_type_BubbleSeries);
    CBubbleSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        aRet.push(this.bubbleSize);
        aRet.push(this.dLbls);
        for(var nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
            aRet.push(this.dPt[nDpt]);
        }
        aRet.push(this.errBars);
        aRet.push(this.trendline);
        return aRet;
    };
    CBubbleSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(oCopy.setBubble3D && AscFormat.isRealBool(this.bubble3D)) {
            oCopy.setBubble3D(this.bubble3D);
        }
        if(oCopy.setBubbleSize && this.bubbleSize) {
            oCopy.setBubbleSize(this.bubbleSize.createDuplicate());
        }
        if(oCopy.setErrBars && this.errBars) {
            oCopy.setErrBars(this.errBars.createDuplicate());
        }
        if(oCopy.setInvertIfNegative && AscFormat.isRealBool(this.invertIfNegative)) {
            oCopy.setInvertIfNegative(this.invertIfNegative);
        }
        if(oCopy.setTrendline && this.trendline) {
            oCopy.setTrendline(this.trendline.createDuplicate());
        }
    };
    CBubbleSeries.prototype.setBubble3D = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BubbleSeries_SetBubble3D, this.bubble3D, pr));
        this.bubble3D = pr;
    };
    CBubbleSeries.prototype.setBubbleSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetBubbleSize, this.bubbleSize, pr));
        this.bubbleSize = pr;
        this.setParentToChild(pr);
    };
    CBubbleSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CBubbleSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_BubbleSeries_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CBubbleSeries.prototype.setErrBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetErrBars, this.errBars, pr));
        this.errBars = pr;
        this.setParentToChild(pr);
    };
    CBubbleSeries.prototype.setInvertIfNegative = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_BubbleSeries_SetInvertIfNegative, this.invertIfNegative, pr));
        this.invertIfNegative = pr;
    };
    CBubbleSeries.prototype.setTrendline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CBubbleSeries.prototype.setXVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetXVal, this.xVal, pr));
        this.xVal = pr;
        this.setParentToChild(pr);
        this.onChangeDataRefs();
    };
    CBubbleSeries.prototype.setYVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_BubbleSeries_SetYVal, this.yVal, pr));
        this.yVal = pr;
        this.setParentToChild(pr);
        this.onChangeDataRefs();
    };
    CBubbleSeries.prototype.setCat = function(pr) {
        this.setXVal(pr);
    };
    CBubbleSeries.prototype.setVal = function(pr) {
        this.setYVal(pr);
    };

    function CChartText() {
        CBaseChartObject.call(this);
        this.rich = null;
        this.strRef = null;
        this.chart = null;
    }

    InitClass(CChartText, CBaseChartObject, AscDFH.historyitem_type_ChartText);
    CChartText.prototype.getChildren = function() {
        return [this.strRef, this.rich];
    };
    CChartText.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.rich) {
            oCopy.setRich(this.rich.createDuplicate());
        }
        if(this.strRef) {
            oCopy.setStrRef(this.strRef.createDuplicate());
        }
    };
    CChartText.prototype.getStyles = CDLbl.prototype.getStyles;
    CChartText.prototype.Get_Theme = CDLbl.prototype.Get_Theme;
    CChartText.prototype.Get_ColorMap = CDLbl.prototype.Get_ColorMap;
    CChartText.prototype.getChartSpace = CDLbl.prototype.getChartSpace;
    CChartText.prototype.setChart = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartFormatSetChart, this.chart, pr));
        this.chart = pr;
    };
    CChartText.prototype.merge = function(tx, noCopyTextBody) {
        if(tx.rich) {
            if(noCopyTextBody === true) {
                this.setRich(tx.rich);
            }
            else {
                this.setRich(tx.rich.createDuplicate());
            }
            this.rich.setParent(this);
        }
        if(tx.strRef) {
            this.strRef = tx.strRef;
        }
    };
    CChartText.prototype.setRich = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartText_SetRich, this.rich, pr));
        this.rich = pr;
        this.setParentToChild(pr);
    };
    CChartText.prototype.setStrRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartText_SetStrRef, this.strRef, pr));
        this.strRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CChartText.prototype.update = function() {
        if(this.strRef) {
            this.strRef.updateCache();
        }
    };
    CChartText.prototype.clearDataCache = function() {
        if(this.strRef) {
            this.strRef.clearDataCache();
        }
    };

    function CDLbls() {
        CBaseChartObject.call(this);
        this.bDelete = null;
        this.dLbl = [];
        this.dLblPos = null;
        this.leaderLines = null;
        this.numFmt = null;
        this.separator = null;
        this.showBubbleSize = null;
        this.showCatName = null;
        this.showLeaderLines = null;
        this.showLegendKey = null;
        this.showPercent = null;
        this.showSerName = null;
        this.showVal = null;
        this.spPr = null;
        this.txPr = null;
    }

    InitClass(CDLbls, CBaseChartObject, AscDFH.historyitem_type_DLbls);
    CDLbls.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
    CDLbls.prototype.Refresh_RecalcData2 = function() {
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.handleUpdateFill = function() {
        this.Refresh_RecalcData2();
    };
    CDLbls.prototype.handleUpdateLn = function() {
        this.Refresh_RecalcData2();
    };
    CDLbls.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setDelete(this.bDelete);
        oCopy.removeAllDLbls();
        for(var i = 0; i < this.dLbl.length; ++i) {
            oCopy.addDLbl(this.dLbl[i].createDuplicate());
        }
        this.dLblPos !== null && oCopy.setDLblPos(this.dLblPos);
        if(this.leaderLines) {
            oCopy.setLeaderLines(this.leaderLines.createDuplicate());
        }
        if(this.numFmt) {
            oCopy.setNumFmt(this.numFmt.createDuplicate());
        }
        this.separator !== null && oCopy.setSeparator(this.separator);
        this.showBubbleSize !== null && oCopy.setShowBubbleSize(this.showBubbleSize);
        this.showCatName !== null && oCopy.setShowCatName(this.showCatName);
        this.showLeaderLines !== null && oCopy.setShowLeaderLines(this.showLeaderLines);
        this.showLegendKey !== null && oCopy.setShowLegendKey(this.showLegendKey);
        this.showPercent !== null && oCopy.setShowPercent(this.showPercent);
        this.showSerName !== null && oCopy.setShowSerName(this.showSerName);
        this.showVal !== null && oCopy.setShowVal(this.showVal);
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CDLbls.prototype.getChildren = function() {
        var aRet = [];
        for(var i = 0; i < this.dLbl.length; ++i) {
            aRet.push(this.dLbl[i]);
        }
        aRet.push(this.leaderLines);
        aRet.push(this.numFmt);
        aRet.push(this.spPr);
        aRet.push(this.txPr);
        return aRet;
    };
    CDLbls.prototype.documentCreateFontMap = function(allFonts) {
        AscFormat.checkTxBodyDefFonts(allFonts, this.txPr);
        for(var i = 0; i < this.dLbl.length; ++i) {
            this.dLbl[i] && AscFormat.checkTxBodyDefFonts(allFonts, this.dLbl[i].txPr);
            this.dLbl[i].tx && this.dLbl[i].tx.rich && this.dLbl[i].tx.rich.content && this.dLbl[i].tx.rich.content.Document_Get_AllFontNames(allFonts);
        }
    };
    CDLbls.prototype.applyLabelsFunction = function(fCallback, value, oDD) {
        fCallback(this, value, oDD, 10);
        for(var i = 0; i < this.dLbl.length; ++i) {
            this.dLbl[i] && fCallback(this.dLbl[i], value, oDD, 10);
        }
    };
    CDLbls.prototype.findDLblByIdx = function(idx) {
        for(var i = 0; i < this.dLbl.length; ++i) {
            if(this.dLbl[i].idx === idx)
                return this.dLbl[i];
        }
        return null;
    };
    CDLbls.prototype.getAllRasterImages = function(images) {
        this.spPr && this.spPr.checkBlipFillRasterImage(images);
        for(var i = 0; i < this.dLbl.length; ++i) {
            this.dLbl[i] && this.dLbl[i].spPr && this.dLbl[i].spPr.checkBlipFillRasterImage(images);
        }
    };
    CDLbls.prototype.checkSpPrRasterImages = function() {
        checkSpPrRasterImages(this.spPr);
        for(var i = 0; i < this.dLbl.length; ++i) {
            this.dLbl[i] && checkSpPrRasterImages(this.dLbl[i].spPr);
        }
    };
    CDLbls.prototype.setDelete = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetDelete, this.bDelete, pr));
        this.bDelete = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.addDLbl = function(pr) {
        for(var nDLbl = this.dLbl.length - 1; nDLbl > -1; --nDLbl) {
            if(this.dLbl[nDLbl].idx === pr.idx) {
                this.removeDLbl(nDLbl);
            }
        }
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DLbls_SetDLbl, this.dLbl.length, [pr], true));
        this.dLbl.push(pr);
        this.setParentToChild(pr);
        this.onChangeDataRefs();

    };
    CDLbls.prototype.removeDLbl = function(nIndex) {
        if(nIndex > -1 && nIndex < this.dLbl.length) {
            this.dLbl[nIndex].setParent(null);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_DLbls_SetDLbl, nIndex, this.dLbl.splice(nIndex, 1), false));
            this.onChangeDataRefs();
        }
    };
    CDLbls.prototype.removeAllDLbls = function() {
        for(var i = this.dLbl.length - 1; i > -1; --i) {
            this.removeDLbl(i);
        }
    };
    CDLbls.prototype.setDLblPos = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DLbls_SetDLblPos, this.dLblPos, pr));
        this.dLblPos = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setLeaderLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbls_SetLeaderLines, this.leaderLines, pr));
        this.leaderLines = pr;
        this.setParentToChild(pr);
    };
    CDLbls.prototype.setNumFmt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbls_SetNumFmt, this.numFmt, pr));
        this.numFmt = pr;
        this.setParentToChild(pr);
    };
    CDLbls.prototype.setSeparator = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_DLbls_SetSeparator, this.separator, pr));
        this.separator = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setShowBubbleSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowBubbleSize, this.showBubbleSize, pr));
        this.showBubbleSize = pr;
    };
    CDLbls.prototype.setShowCatName = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowCatName, this.showCatName, pr));
        this.showCatName = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setShowLeaderLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowLeaderLines, this.showLeaderLines, pr));
        this.showLeaderLines = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setShowLegendKey = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowLegendKey, this.showLegendKey, pr));
        this.showLegendKey = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setShowPercent = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowPercent, this.showPercent, pr));
        this.showPercent = pr;
    };
    CDLbls.prototype.setShowSerName = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowSerName, this.showSerName, pr));
        this.showSerName = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setShowVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DLbls_SetShowVal, this.showVal, pr));
        this.showVal = pr;
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbls_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DLbls_SetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
        this.onChartUpdateDataLabels();
    };
    CDLbls.prototype.setDeleteValue = function(bValue) {
        if(bValue === false) {
            if(this.bDelete !== null) {
                this.setDelete(null);
            }
            for(var nDlbl = this.dLbl.length - 1; nDlbl > -1; --nDlbl) {
                var oDlbl = this.dLbl[nDlbl];
                if(oDlbl.bDelete) {
                    this.removeDLbl(nDlbl);
                }
                if(oDlbl.bDelete === false) {
                    oDlbl.setDelete(null);
                }
            }
        }
        else {
            if(this.bDelete !== true) {
                this.setDelete(true);
                this.removeAllDLbls();
            }
        }
    };
    CDLbls.prototype.checkPosition = function(aPositions) {
        fCheckDLblPosition(this, aPositions);
    };
    CDLbls.prototype.setSettings = function(nPos, oProps) {
        fCheckDLblSettings(this, nPos, oProps)
    };
    CDLbls.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        var aColors;

        //check chart or series
        var nParentType = this.parent.getObjectType();
        if(nParentType === AscDFH.historyitem_type_BarChart ||
            nParentType === AscDFH.historyitem_type_AreaChart ||
            nParentType === AscDFH.historyitem_type_BubbleChart ||
            nParentType === AscDFH.historyitem_type_DoughnutChart ||
            nParentType === AscDFH.historyitem_type_LineChart ||
            nParentType === AscDFH.historyitem_type_OfPieChart ||
            nParentType === AscDFH.historyitem_type_PieChart ||
            nParentType === AscDFH.historyitem_type_RadarChart ||
            nParentType === AscDFH.historyitem_type_ScatterChart ||
            nParentType === AscDFH.historyitem_type_StockChart ||
            nParentType === AscDFH.historyitem_type_SurfaceChart) {
            this.removeAllDLbls();
            this.applyStyleEntry(oChartStyle.dataLabel, oColors.generateColors(1), 0, bReset);
        }
        else if( nParentType === AscDFH.historyitem_type_AreaSeries ||
            nParentType === AscDFH.historyitem_type_BarSeries ||
            nParentType === AscDFH.historyitem_type_BubbleSeries ||
            nParentType === AscDFH.historyitem_type_LineSeries ||
            nParentType === AscDFH.historyitem_type_PieSeries ||
            nParentType === AscDFH.historyitem_type_RadarSeries ||
            nParentType === AscDFH.historyitem_type_ScatterSer ||
            nParentType === AscDFH.historyitem_type_SurfaceSeries) {
            var nMaxSeriesIdx = this.parent.getMaxSeriesIdx() + 1;
            aColors = oColors.generateColors(nMaxSeriesIdx + 1);
            this.applyStyleEntry(oChartStyle.dataLabel, aColors, this.parent.idx, bReset);
            this.removeAllDLbls();
            if(this.bDelete !== true) {
                if(this.bDelete !== null) {
                    this.setDelete(null);
                }
                if(this.parent.isVaryColors()) {
                    var nPtsCount = this.parent.getValuesCount();
                    aColors = oColors.generateColors(nPtsCount);
                    for(var nDLbl = 0; nDLbl < nPtsCount; ++nDLbl) {
                        var oDLbl = new CDLbl();
                        CDLbl.prototype.fillObject.call(this, oDLbl);
                        oDLbl.setIdx(nDLbl);
                        this.addDLbl(oDLbl);
                        oDLbl.applyStyleEntry(oChartStyle.dataLabel, aColors, nDLbl, bReset);
                    }
                }
            }
        }
    };
    CDLbls.prototype.checkChartStyle = function() {
        var oChartSpace = this.getChartSpace();
        if(!oChartSpace) {
            return;
        }
        oChartSpace.checkElementChartStyle(this);
    }
    CDLbls.prototype.correctValues = CDLbl.prototype.correctValues;

    function fCheckDLblSettings(oLbls, nPos, oProps) {
        if(oLbls.dLblPos !== nPos) {
            oLbls.setDLblPos(nPos);
        }
        if(AscFormat.isRealBool(oProps.showCatName) ||
            AscFormat.isRealBool(oProps.showSerName) ||
            AscFormat.isRealBool(oProps.showVal)) {
            if(oLbls.setDelete && oLbls.bDelete) {
                oLbls.setDelete(false);
            }
            if(AscFormat.isRealBool(oProps.showCatName)) {
                oLbls.setShowCatName(oProps.showCatName);
            }
            if(AscFormat.isRealBool(oProps.showSerName)) {
                oLbls.setShowSerName(oProps.showSerName);
            }
            if(AscFormat.isRealBool(oProps.showVal)) {
                oLbls.setShowVal(oProps.showVal);
            }
            if(!AscFormat.isRealBool(oLbls.showLegendKey) || oLbls.showLegendKey === true) {
                oLbls.setShowLegendKey(false);
            }
            if(!AscFormat.isRealBool(oLbls.showPercent) || oLbls.showPercent === true) {
                oLbls.setShowPercent(false);
            }
            if(!AscFormat.isRealBool(oLbls.showBubbleSize) || oLbls.showBubbleSize === true) {
                oLbls.setShowBubbleSize(false);
            }
            if(typeof oProps.separator === "string" && oProps.separator.length > 0) {
                oLbls.setSeparator(oProps.separator);
            }
        }

        if(Array.isArray(oLbls.dLbl)) {
            for(var nLbl = 0; nLbl < oLbls.dLbl.length; ++nLbl) {
                fCheckDLblSettings(oLbls.dLbl[nLbl], nPos, oProps)
            }
        }
    }

    function fCheckDLPostion(nPostion, aPositions) {
        if(aPositions.length === 0) {
            return null;
        }
        for(var nIndex = 0; nIndex < aPositions.length; ++nIndex) {
            if(aPositions[nIndex] === nPostion) {
                return nPostion;
            }
        }
        return aPositions[0];
    }

    function fCheckDLblPosition(oLbls, aPositions) {
        var nPos = fCheckDLPostion(oLbls.dLblPos, aPositions);
        if(oLbls.dLblPos !== nPos) {
            oLbls.setDLblPos(nPos);
        }
        if(Array.isArray(oLbls.dLbl)) {
            for(nPos = 0; nPos < oLbls.dLbl.length; ++nPos) {
                oLbls.dLbl[nPos].checkPosition(aPositions);
            }
        }
    }

    function CDPt() {
        CBaseChartObject.call(this);
        this.bubble3D = null;
        this.explosion = null;
        this.idx = null;
        this.invertIfNegative = null;
        this.marker = null;
        this.pictureOptions = null;
        this.spPr = null;

        this.recalcInfo =
        {
            recalcLbl: true
        };
    }

    InitClass(CDPt, CBaseChartObject, AscDFH.historyitem_type_DPt);
    CDPt.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.marker);
        aRet.push(this.pictureOptions);
        aRet.push(this.spPr);
        return aRet;
    };
    CDPt.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setBubble3D(this.bubble3D);
        oCopy.setExplosion(this.explosion);
        oCopy.setIdx(this.idx);
        oCopy.setInvertIfNegative(this.invertIfNegative);
        if(this.marker) {
            oCopy.setMarker(this.marker.createDuplicate());
        }
        if(this.pictureOptions) {
            oCopy.setPictureOptions(this.pictureOptions.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
    };
    CDPt.prototype.setBubble3D = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DPt_SetBubble3D, this.bubble3D, pr));
        this.bubble3D = pr;
    };
    CDPt.prototype.setExplosion = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DPt_SetExplosion, this.explosion, pr));
        this.explosion = pr;
    };
    CDPt.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DPt_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CDPt.prototype.setInvertIfNegative = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DPt_SetInvertIfNegative, this.invertIfNegative, pr));
        this.invertIfNegative = pr;
    };
    CDPt.prototype.setMarker = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DPt_SetMarker, this.marker, pr));
        this.marker = pr;
        this.setParentToChild(pr);
    };
    CDPt.prototype.setPictureOptions = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DPt_SetPictureOptions, this.pictureOptions, pr));
        this.pictureOptions = pr;
        this.setParentToChild(pr);
    };
    CDPt.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DPt_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };

    function CDTable() {
        CBaseChartObject.call(this);
        this.showHorzBorder = null;
        this.showKeys = null;
        this.showOutline = null;
        this.showVertBorder = null;
        this.spPr = null;
        this.txPr = null;
    }

    InitClass(CDTable, CBaseChartObject, AscDFH.historyitem_type_DTable);
    CDTable.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.spPr);
        aRet.push(this.txPr);
        return aRet;
    };
    CDTable.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setShowHorzBorder(this.showHorzBorder);
        oCopy.setShowKeys(this.showKeys);
        oCopy.setShowOutline(this.showOutline);
        oCopy.setShowVertBorder(this.showVertBorder);
        if(this.spPr)
            oCopy.setSpPr(this.spPr.createDuplicate());
        if(this.txPr)
            oCopy.setTxPr(this.txPr.createDuplicate());
    };
    CDTable.prototype.setShowHorzBorder = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DTable_SetShowHorzBorder, this.showHorzBorder, pr));
        this.showHorzBorder = pr;
    };
    CDTable.prototype.setShowKeys = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DTable_SetShowKeys, this.showHorzBorder, pr));
        this.showKeys = pr;
    };
    CDTable.prototype.setShowOutline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DTable_SetShowOutline, this.showHorzBorder, pr));
        this.showOutline = pr;
    };
    CDTable.prototype.setShowVertBorder = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_DTable_SetShowVertBorder, this.showHorzBorder, pr));
        this.showVertBorder = pr;
    };
    CDTable.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DTable_SetSpPr, this.showHorzBorder, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CDTable.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DTable_SetTxPr, this.showHorzBorder, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
    };
    CDTable.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.chartArea, oColors.generateColors(1), 0, bReset);
    };

    var UNIT_MULTIPLIERS = [];
    UNIT_MULTIPLIERS[c_oAscValAxUnits.BILLIONS] = 1.0 / 1000000000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.HUNDRED_MILLIONS] = 1.0 / 100000000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.HUNDREDS] = 1.0 / 100.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.HUNDRED_THOUSANDS] = 1.0 / 100000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.MILLIONS] = 1.0 / 1000000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.TEN_MILLIONS] = 1.0 / 10000000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.TEN_THOUSANDS] = 1.0 / 10000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.TRILLIONS] = 1.0 / 1000000000000.0;
    UNIT_MULTIPLIERS[c_oAscValAxUnits.THOUSANDS] = 1.0 / 1000.0;

    function CDispUnits() {
        CBaseChartObject.call(this);
        this.builtInUnit = null;
        this.custUnit = null;
        this.dispUnitsLbl = null;
    }

    InitClass(CDispUnits, CBaseChartObject, AscDFH.historyitem_type_DispUnits);
    CDispUnits.prototype.getChildren = function() {
        return [this.dispUnitsLbl];
    };
    CDispUnits.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setBuiltInUnit(this.builtInUnit);
        oCopy.setCustUnit(this.custUnit);
        if(this.dispUnitsLbl) {
            oCopy.setDispUnitsLbl(this.dispUnitsLbl.createDuplicate());
        }
    };
    CDispUnits.prototype.getMultiplier = function() {
        if(AscFormat.isRealNumber(this.builtInUnit)) {
            if(AscFormat.isRealNumber(UNIT_MULTIPLIERS[this.builtInUnit]))
                return UNIT_MULTIPLIERS[this.builtInUnit];
        }
        else if(AscFormat.isRealNumber(this.custUnit) && this.custUnit > 0.0)
            return 1.0 / this.custUnit;
        return 1;
    };
    CDispUnits.prototype.onUpdate = function() {
        if(this.parent) {
            this.parent.onUpdate();
        }
    };
    CDispUnits.prototype.setBuiltInUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_DispUnitsSetBuiltInUnit, this.builtInUnit, pr));
        this.builtInUnit = pr;
        this.onUpdate();
    };
    CDispUnits.prototype.setCustUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DispUnitsSetCustUnit, this.custUnit, pr));
        this.custUnit = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CDispUnits.prototype.setDispUnitsLbl = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_DispUnitsSetDispUnitsLbl, this.dispUnitsLbl, pr));
        this.dispUnitsLbl = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };

    function CDoughnutChart() {
        CChartBase.call(this);
        this.firstSliceAng = null;
        this.holeSize = null;
    }

    InitClass(CDoughnutChart, CChartBase, AscDFH.historyitem_type_DoughnutChart);
    CDoughnutChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;

        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_DoughnutChart_SetDLbls :
            {
                this.onChartUpdateDataLabels();
                break;
            }
            case AscDFH.historyitem_DoughnutChart_SetFirstSliceAng :
            {
                break;
            }
            case AscDFH.historyitem_DoughnutChart_SetHoleSize :
            {
                break;
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_DoughnutChart_SetVaryColor :
            {
                break;
            }
        }
    };
    CDoughnutChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.ctr;
    };
    CDoughnutChart.prototype.getEmptySeries = function() {
        return new CPieSeries();
    };
    CDoughnutChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(oCopy.setFirstSliceAng && this.firstSliceAng !== null) {
            oCopy.setFirstSliceAng(this.firstSliceAng);
        }
        if(oCopy.setHoleSize && this.holeSize !== null) {
            oCopy.setHoleSize(this.holeSize);
        }
    };
    CDoughnutChart.prototype.setFirstSliceAng = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DoughnutChart_SetFirstSliceAng, this.firstSliceAng, pr));
        this.firstSliceAng = pr;
    };
    CDoughnutChart.prototype.setHoleSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_DoughnutChart_SetHoleSize, this.holeSize, pr));
        this.holeSize = pr;
    };
    CDoughnutChart.prototype.getChartType = function() {
        return Asc.c_oAscChartTypeSettings.doughnut;
    };

    function CErrBars() {
        CBaseChartObject.call(this);
        this.errBarType = null;
        this.errDir = null;
        this.errValType = null;
        this.minus = null;
        this.noEndCap = null;
        this.plus = null;
        this.spPr = null;
        this.val = null;

        this.pen = null;
    }

    InitClass(CErrBars, CBaseChartObject, AscDFH.historyitem_type_ErrBars);
    CErrBars.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.minus);
        aRet.push(this.plus);
        aRet.push(this.spPr);
        return aRet;
    };
    CErrBars.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setErrBarType(this.errBarType);
        oCopy.setErrDir(this.errDir);
        oCopy.setErrValType(this.errValType);
        if(this.minus) {
            oCopy.setMinus(this.minus.createDuplicate());
        }
        oCopy.setNoEndCap(this.noEndCap);
        if(this.plus) {
            oCopy.setPlus(this.plus.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        oCopy.setVal(this.val);
    };
    CErrBars.prototype.setErrBarType = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ErrBars_SetErrBarType, this.errBarType, pr));
        this.errBarType = pr;
    };
    CErrBars.prototype.setErrDir = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ErrBars_SetErrDir, this.errDir, pr));
        this.errDir = pr;
    };
    CErrBars.prototype.setErrValType = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ErrBars_SetErrValType, this.errDir, pr));
        this.errValType = pr;
    };
    CErrBars.prototype.setMinus = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ErrBars_SetMinus, this.minus, pr));
        this.minus = pr;
        this.setParentToChild(pr);
    };
    CErrBars.prototype.setNoEndCap = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ErrBars_SetNoEndCap, this.noEndCap, pr));
        this.noEndCap = pr;
    };
    CErrBars.prototype.setPlus = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ErrBars_SetPlus, this.plus, pr));
        this.plus = pr;
        this.setParentToChild(pr);
    };
    CErrBars.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ErrBars_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
        this.pen = null;
    };
    CErrBars.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_ErrBars_SetVal, this.val, pr));
        this.val = pr;
    };
    CErrBars.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.errorBar, oColors.generateColors(1), 0, bReset);
    };
    CErrBars.prototype.handleUpdateLn = function() {
        this.pen = null;
    };
    CErrBars.prototype.update = function() {
        if(this.plus) {
            this.plus.update();
        }
        if(this.minus) {
            this.minus.update();
        }
    };
    CErrBars.prototype.clearDataCache = function() {
        if(this.plus) {
            this.plus.clearDataCache();
        }
        if(this.minus) {
            this.minus.clearDataCache();
        }
    };
    CErrBars.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.plus) {
            this.plus.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.minus) {
            this.minus.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CErrBars.prototype.getMinusDataRefs = function() {
        if(this.minus) {
            this.minus.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CErrBars.prototype.getPlusDataRefs = function() {
        if(this.plus) {
            this.plus.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CErrBars.prototype.getPen = function() {
        if(this.pen === null) {
            let oPen = AscFormat.builder_CreateLine(12700, {UniFill: AscFormat.CreateUnfilFromRGB(0, 0, 0)});
            let oLn = this.spPr && this.spPr.ln;
            if(oLn) {
                oPen.merge(oLn);
            }
            this.pen = oPen;
            if(this.pen.Fill) {
                let oChartSpace = this.getChartSpace();
                this.pen.Fill.check(oChartSpace.Get_Theme(), oChartSpace.Get_ColorMap());
            }
        }
        return this.pen;
    };

    function CLayout() {
        CBaseChartObject.call(this);
        this.h = null;
        this.hMode = null;
        this.layoutTarget = null;
        this.w = null;
        this.wMode = null;
        this.x = null;
        this.xMode = null;
        this.y = null;
        this.yMode = null;
    }

    InitClass(CLayout, CBaseChartObject, AscDFH.historyitem_type_Layout);
    CLayout.prototype.Refresh_RecalcData = function(Data) {
        if(this.parent && this.parent.Refresh_RecalcData2) {
            this.parent.Refresh_RecalcData2();
        }
    };
    CLayout.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setH(this.h);
        oCopy.setHMode(this.hMode);
        oCopy.setLayoutTarget(this.layoutTarget);
        oCopy.setW(this.w);
        oCopy.setWMode(this.wMode);
        oCopy.setX(this.x);
        oCopy.setXMode(this.xMode);
        oCopy.setY(this.y);
        oCopy.setYMode(this.yMode);
    };
    CLayout.prototype.setH = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Layout_SetH, this.h, pr));
        this.h = pr;
    };
    CLayout.prototype.setHMode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Layout_SetHMode, this.hMode, pr));
        this.hMode = pr;
    };
    CLayout.prototype.setLayoutTarget = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Layout_SetLayoutTarget, this.layoutTarget, pr));
        this.layoutTarget = pr;
    };
    CLayout.prototype.setW = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Layout_SetW, this.w, pr));
        this.w = pr;
    };
    CLayout.prototype.setWMode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Layout_SetWMode, this.wMode, pr));
        this.wMode = pr;
    };
    CLayout.prototype.setX = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Layout_SetX, this.x, pr));
        this.x = pr;
    };
    CLayout.prototype.setXMode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Layout_SetXMode, this.xMode, pr));
        this.xMode = pr;
    };
    CLayout.prototype.setY = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Layout_SetY, this.y, pr));
        this.y = pr;
    };
    CLayout.prototype.setYMode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Layout_SetYMode, this.yMode, pr));
        this.yMode = pr;
    };

    function CLegend() {
        CBaseChartObject.call(this);
        this.layout = null;
        this.legendEntryes = [];
        this.legendPos = null;
        this.overlay = false;
        this.spPr = null;
        this.txPr = null;

        this.rot = 0;
        this.flipH = false;
        this.flipV = false;
        this.x = null;
        this.y = null;
        this.extX = null;
        this.extY = null;
        this.calcEntryes = [];
        this.transform = new CMatrix();
        this.invertTransform = new CMatrix();
        this.localTransform = new CMatrix();
    }

    InitClass(CLegend, CBaseChartObject, AscDFH.historyitem_type_Legend);
    CLegend.prototype.recalculatePen = CShape.prototype.recalculatePen;
    CLegend.prototype.recalculateBrush = CShape.prototype.recalculateBrush;
    CLegend.prototype.getCompiledLine = CShape.prototype.getCompiledLine;
    CLegend.prototype.getCompiledFill = CShape.prototype.getCompiledFill;
    CLegend.prototype.getCompiledTransparent = CShape.prototype.getCompiledTransparent;
    CLegend.prototype.check_bounds = CShape.prototype.check_bounds;
    CLegend.prototype.getCardDirectionByNum = CShape.prototype.getCardDirectionByNum;
    CLegend.prototype.getNumByCardDirection = CShape.prototype.getNumByCardDirection;
    CLegend.prototype.getResizeCoefficients = CShape.prototype.getResizeCoefficients;
    CLegend.prototype.getInvertTransform = CShape.prototype.getInvertTransform;
    CLegend.prototype.getTransformMatrix = CShape.prototype.getTransformMatrix;
    CLegend.prototype.hitToHandles = CShape.prototype.hitToHandles;
    CLegend.prototype.getFullFlipH = CShape.prototype.getFullFlipH;
    CLegend.prototype.getFullFlipV = CShape.prototype.getFullFlipV;
    CLegend.prototype.getAspect = CShape.prototype.getAspect;
    CLegend.prototype.getGeom = CShape.prototype.getGeom;
    CLegend.prototype.hitInBoundingRect = CShape.prototype.hitInBoundingRect;
    CLegend.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
    CLegend.prototype.hitInPath = CShape.prototype.hitInPath;
    CLegend.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
    CLegend.prototype.convertPixToMM = function(pix) {
        return this.parent && this.parent.parent && this.parent.parent.convertPixToMM && this.parent.parent.convertPixToMM(pix);
    };
    CLegend.prototype.findCalcEntryByIdx = function(idx) {
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            if(this.calcEntryes[i] && this.calcEntryes[i].idx === idx) {
                return this.calcEntryes[i];
            }
        }
        return null;
    };
    CLegend.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.layout);
        for(var i = 0; i < this.legendEntryes.length; ++i) {
            aRet.push(this.legendEntryes[i]);
        }
        aRet.push(this.spPr);
        aRet.push(this.txPr);
        return aRet;
    };
    CLegend.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.layout) {
            oCopy.setLayout(this.layout.createDuplicate());
        }
        for(var i = 0; i < this.legendEntryes.length; ++i) {
            oCopy.addLegendEntry(this.legendEntryes[i].createDuplicate());
        }
        oCopy.setLegendPos(this.legendPos);
        oCopy.setOverlay(this.overlay);
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CLegend.prototype.getCalcEntryByIdx = function(idx, drawingDocument) {
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            if(this.calcEntryes[i] && this.calcEntryes[i].idx == idx) {
                return this.calcEntryes[i];
            }
        }
        return AscFormat.ExecuteNoHistory(function() {

            var calcEntry = new AscFormat.CalcLegendEntry(this, this.chart, idx);
            calcEntry.txBody = AscFormat.CreateTextBodyFromString("" + idx, drawingDocument, calcEntry);
            calcEntry.txBody.getRectWidth(2000);
            return calcEntry;
        }, this, []);
    };
    CLegend.prototype.updateLegendPos = function() {
        if(this.overlay) {
            if(c_oAscChartLegendShowSettings.left === this.legendPos) {
                this.legendPos = c_oAscChartLegendShowSettings.leftOverlay;
            } else if(c_oAscChartLegendShowSettings.right === this.legendPos) {
                this.legendPos = c_oAscChartLegendShowSettings.rightOverlay;
            }
        }
    };
    CLegend.prototype.getCompiledStyle = function() {
        return null;
    };
    CLegend.prototype.getParentObjects = function() {
        if(this.chart) {
            return this.chart.getParentObjects();
        }
        else {
            return {};
        }
    };
    CLegend.prototype.checkHitToBounds = function(x, y) {
        CDLbl.prototype.checkHitToBounds.call(this, x, y);
    };
    CLegend.prototype.getCanvasContext = function() {
        return CDLbl.prototype.getCanvasContext.call(this);
    };
    CLegend.prototype.canRotate = function() {
        return false;
    };
    CLegend.prototype.getNoChangeAspect = function() {
        return false;
    };
    CLegend.prototype.isChart = function() {
        return false;
    };
    CLegend.prototype.canMove = function() {
        return true;
    };
    CLegend.prototype.selectObject = function() {

    };
    CLegend.prototype.getHierarchy = function() {
        return this.chart ? this.chart.getHierarchy() : []
    };
    CLegend.prototype.isEmptyPlaceholder = function() {
        return false;
    };
    CLegend.prototype.isPlaceholder = function() {
        return false;
    };
    CLegend.prototype.draw = function(g) {
        g.bDrawSmart = true;
        CShape.prototype.draw.call(this, g);
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            this.calcEntryes[i].draw(g);
        }
        g.bDrawSmart = false;
    };
    CLegend.prototype.hit = function(x, y) {
        var t_x = this.invertTransform.TransformPointX(x, y);
        var t_y = this.invertTransform.TransformPointY(x, y);
        return t_x >= 0 && t_y >= 0 && t_x <= this.extX && t_y <= this.extY;
    };
    CLegend.prototype.setPosition = function(x, y) {
        this.x = x;
        this.y = y;
        this.localTransform.Reset();
        global_MatrixTransformer.TranslateAppend(this.localTransform, this.x, this.y);
        //if(this.chart)
        //    global_MatrixTransformer.MultiplyAppend(this.localTransform, this.chart.localTransform);
        this.transform = this.localTransform.CreateDublicate();
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        var entry;
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            entry = this.calcEntryes[i];
            entry.localTransformText.Reset();
            global_MatrixTransformer.TranslateAppend(entry.localTransformText, entry.localX, entry.localY);

            global_MatrixTransformer.MultiplyAppend(entry.localTransformText, this.localTransform);
            entry.transformText = entry.localTransformText.CreateDublicate();
            entry.invertTransformText = global_MatrixTransformer.Invert(entry.transformText);
            if(entry.calcMarkerUnion.marker) {
                entry.calcMarkerUnion.marker.localTransform.Reset();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.marker.localTransform, entry.calcMarkerUnion.marker.localX, entry.calcMarkerUnion.marker.localY);

                global_MatrixTransformer.MultiplyAppend(entry.calcMarkerUnion.marker.localTransform, this.localTransform);
                entry.calcMarkerUnion.marker.transform = entry.calcMarkerUnion.marker.localTransform.CreateDublicate();
                entry.calcMarkerUnion.marker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.marker.transform);
            }
            if(entry.calcMarkerUnion.lineMarker) {
                entry.calcMarkerUnion.lineMarker.localTransform.Reset();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.lineMarker.localTransform, entry.calcMarkerUnion.lineMarker.localX, entry.calcMarkerUnion.lineMarker.localY);

                global_MatrixTransformer.MultiplyAppend(entry.calcMarkerUnion.lineMarker.localTransform, this.localTransform);
                entry.calcMarkerUnion.lineMarker.transform = entry.calcMarkerUnion.lineMarker.localTransform.CreateDublicate();
                entry.calcMarkerUnion.lineMarker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.lineMarker.transform);
            }
        }
    };
    CLegend.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, x, y);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        var entry;
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            entry = this.calcEntryes[i];
            entry.transformText = entry.localTransformText.CreateDublicate();
            global_MatrixTransformer.TranslateAppend(entry.transformText, x, y);
            entry.invertTransformText = global_MatrixTransformer.Invert(entry.transformText);

            if(entry.calcMarkerUnion.marker) {
                entry.calcMarkerUnion.marker.transform = entry.calcMarkerUnion.marker.localTransform.CreateDublicate();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.marker.transform, x, y);

                entry.calcMarkerUnion.marker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.marker.transform);
            }

            if(entry.calcMarkerUnion.lineMarker) {
                entry.calcMarkerUnion.lineMarker.transform = entry.calcMarkerUnion.lineMarker.localTransform.CreateDublicate();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.lineMarker.transform, x, y);

                entry.calcMarkerUnion.lineMarker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.lineMarker.transform);

            }
        }
    };
    CLegend.prototype.checkShapeChildTransform = function(t) {
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, this.posX, this.posY);
        global_MatrixTransformer.MultiplyAppend(this.transform, t);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        var entry;
        for(var i = 0; i < this.calcEntryes.length; ++i) {
            entry = this.calcEntryes[i];
            entry.transformText = entry.localTransformText.CreateDublicate();
            global_MatrixTransformer.TranslateAppend(entry.transformText, this.posX, this.posY);


            global_MatrixTransformer.MultiplyAppend(entry.transformText, t);

            entry.invertTransformText = global_MatrixTransformer.Invert(entry.transformText);


            if(entry.calcMarkerUnion.marker) {
                entry.calcMarkerUnion.marker.transform = entry.calcMarkerUnion.marker.localTransform.CreateDublicate();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.marker.transform, this.posX, this.posY);

                global_MatrixTransformer.MultiplyAppend(entry.calcMarkerUnion.marker.transform, t);

                entry.calcMarkerUnion.marker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.marker.transform);
            }

            if(entry.calcMarkerUnion.lineMarker) {
                entry.calcMarkerUnion.lineMarker.transform = entry.calcMarkerUnion.lineMarker.localTransform.CreateDublicate();
                global_MatrixTransformer.TranslateAppend(entry.calcMarkerUnion.lineMarker.transform, this.posX, this.posY);
                global_MatrixTransformer.MultiplyAppend(entry.calcMarkerUnion.lineMarker.transform, t);

                entry.calcMarkerUnion.lineMarker.invertTransform = global_MatrixTransformer.Invert(entry.calcMarkerUnion.lineMarker.transform);
            }
        }
    };
    CLegend.prototype.setLayout = function(layout) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Legend_SetLayout, this.layout, layout));
        this.layout = layout;
        this.setParentToChild(layout);
    };
    CLegend.prototype.addLegendEntry = function(legendEntry) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Legend_AddLegendEntry, this.legendEntryes.length, [legendEntry], true));
        this.legendEntryes.push(legendEntry);
        this.setParentToChild(legendEntry);
    };
    CLegend.prototype.setLegendPos = function(legendPos) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Legend_SetLegendPos, this.legendPos, legendPos));
        this.legendPos = legendPos;
        this.onChartInternalUpdate();
    };
    CLegend.prototype.setOverlay = function(overlay) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Legend_SetOverlay, this.overlay, overlay));
        this.overlay = overlay;
        this.onChartInternalUpdate();
    };
    CLegend.prototype.setSpPr = function(spPr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Legend_SetSpPr, this.spPr, spPr));
        this.spPr = spPr;
        this.setParentToChild(spPr);
    };
    CLegend.prototype.setTxPr = function(txPr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Legend_SetTxPr, this.txPr, txPr));
        this.txPr = txPr;
        this.setParentToChild(txPr);
    };
    CLegend.prototype.Refresh_RecalcData2 = function() {
        this.onChartInternalUpdate();
    };
    CLegend.prototype.findLegendEntryByIndex = function(idx) {
        for(var i = 0; i < this.legendEntryes.length; ++i) {
            if(this.legendEntryes[i].idx === idx)
                return this.legendEntryes[i];
        }
        return null;
    };
    CLegend.prototype.getPropsPos = function() {
        var nLegendPos = c_oAscChartLegendShowSettings.bottom;
        if(AscFormat.isRealNumber(this.legendPos)) {
            nLegendPos = this.legendPos;
        }
        if(this.isOverlay()) {
            if(c_oAscChartLegendShowSettings.left === nLegendPos) {
                nLegendPos = c_oAscChartLegendShowSettings.leftOverlay;
            }
            if(c_oAscChartLegendShowSettings.right === nLegendPos) {
                nLegendPos = c_oAscChartLegendShowSettings.rightOverlay;
            }
        }
        return nLegendPos;
    };
    CLegend.prototype.isOverlay = function() {
        return (this.overlay === true);
    };
    CLegend.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.legend, oColors.generateColors(1), 0, bReset);
    };

    function CLegendEntry() {
        CBaseChartObject.call(this);
        this.bDelete = null;
        this.idx = null;
        this.txPr = null;
    }

    InitClass(CLegendEntry, CBaseChartObject, AscDFH.historyitem_type_LegendEntry);
    CLegendEntry.prototype.getChildren = function() {
        return [this.txPr];
    };
    CLegendEntry.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setDelete(this.bDelete);
        oCopy.setIdx(this.idx);
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CLegendEntry.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
    CLegendEntry.prototype.setDelete = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_LegendEntry_SetDelete, this.bDelete, pr));
        this.bDelete = pr;
        this.Refresh_RecalcData2();
    };
    CLegendEntry.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_LegendEntry_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CLegendEntry.prototype.Refresh_RecalcData2 = function() {
        if(this.parent && this.parent.Refresh_RecalcData2) {
            this.parent.Refresh_RecalcData2();
        }
    };
    CLegendEntry.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LegendEntry_SetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
    };

    function CLineChart() {
        CChartBase.call(this);
        this.dropLines = null;
        this.grouping = null;
        this.hiLowLines = null;
        this.marker = null;
        this.smooth = false;
        this.upDownBars = null;
    }

    InitClass(CLineChart, CChartBase, AscDFH.historyitem_type_LineChart);
    CLineChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_LineChart_AddAxId:
            case AscDFH.historyitem_CommonChart_AddAxId:
            {
                break
            }
            case AscDFH.historyitem_LineChart_SetDLbls:
            case AscDFH.historyitem_CommonChart_SetDlbls:
            {
                this.onChartUpdateDataLabels();
                break
            }
            case AscDFH.historyitem_LineChart_SetDropLines:
            {
                break
            }
            case AscDFH.historyitem_LineChart_SetGrouping:
            {
                this.onChartInternalUpdate();
                break
            }
            case AscDFH.historyitem_LineChart_SetHiLowLines:
            {
                break
            }
            case AscDFH.historyitem_LineChart_SetMarker:
            {
                this.onChartUpdateType();
                break
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break
            }
            case AscDFH.historyitem_LineChart_SetSmooth:
            {
                this.onChartUpdateType();
                break
            }
            case AscDFH.historyitem_LineChart_SetUpDownBars:
            {
                break
            }
            case AscDFH.historyitem_LineChart_SetVaryColors:
            case AscDFH.historyitem_CommonChart_SetVaryColors:
            {
                break
            }
        }
    };
    CLineChart.prototype.getEmptySeries = function() {
        return new CLineSeries();
    };
    CLineChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.r;
    };
    CLineChart.prototype.getAllRasterImages = function(images) {
        CChartBase.prototype.getAllRasterImages.call(this, images);
        if(this.upDownBars) {
            this.upDownBars.upBars && this.upDownBars.upBars.checkBlipFillRasterImage(images);
            this.upDownBars.downBars && this.upDownBars.downBars.checkBlipFillRasterImage(images);
        }
    };
    CLineChart.prototype.checkSpPrRasterImages = function(images) {
        CChartBase.prototype.checkSpPrRasterImages.call(this, images);
        if(this.upDownBars) {
            checkSpPrRasterImages(this.upDownBars.upBars);
            checkSpPrRasterImages(this.upDownBars.downBars);
        }
    };
    CLineChart.prototype.getChildren = function() {
        var aRet = CChartBase.prototype.getChildren.call(this);
        aRet.push(this.dropLines);
        aRet.push(this.hiLowLines);
        aRet.push(this.upDownBars);
        return aRet;
    };
    CLineChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.dropLines && oCopy.setDropLines)
            oCopy.setDropLines(this.dropLines.createDuplicate());
        if(AscFormat.isRealNumber(this.grouping) && oCopy.setGrouping)
            oCopy.setGrouping(this.grouping);
        if(this.hiLowLines && oCopy.setHiLowLines)
            oCopy.setHiLowLines(this.hiLowLines.createDuplicate());
        if(this.upDownBars && oCopy.setUpDownBars)
            oCopy.setUpDownBars(this.upDownBars.createDuplicate());
        if(AscFormat.isRealBool(this.marker) && oCopy.setMarker)
            oCopy.setMarker(this.marker);
        if(AscFormat.isRealBool(this.smooth) && oCopy.setSmooth)
            oCopy.setSmooth(this.smooth);
    };
    CLineChart.prototype.setDropLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineChart_SetDropLines, this.dropLines, pr));
        this.dropLines = pr;
        this.setParentToChild(pr);
    };
    CLineChart.prototype.setGrouping = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_LineChart_SetGrouping, this.grouping, pr));
        this.grouping = pr;
        this.onChartInternalUpdate();
    };
    CLineChart.prototype.setHiLowLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineChart_SetHiLowLines, this.hiLowLines, pr));
        this.hiLowLines = pr;
        this.setParentToChild(pr);
    };
    CLineChart.prototype.setMarker = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_LineChart_SetMarker, this.marker, pr));
        this.marker = pr;
        this.onChartUpdateType();
    };
    CLineChart.prototype.setSmooth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_LineChart_SetSmooth, this.smooth, pr));
        this.smooth = pr;
        this.onChartUpdateType();
    };
    CLineChart.prototype.setUpDownBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineChart_SetUpDownBars, this.upDownBars, pr));
        this.upDownBars = pr;
        this.setParentToChild(pr);
    };
    CLineChart.prototype.getLineParams = function() {
        var nFlag = 0;//
        var bShowMarker = false;
        var bNoLine = true;
        var bSmooth = true;
        var aSeries = this.series, oSeries;
        var nSer;
        if(this.marker !== false) {
            for(nSer = 0; nSer < aSeries.length; ++nSer) {
                oSeries = aSeries[nSer];
                if(!oSeries.marker) {
                    bShowMarker = true;
                    break;
                }
                if(oSeries.marker.symbol !== AscFormat.SYMBOL_NONE) {
                    bShowMarker = true;
                    break;
                }
            }
        }
        for(nSer = 0; nSer < aSeries.length; ++nSer) {
            oSeries = aSeries[nSer];
            if(!oSeries.hasNoFillLine()) {
                bNoLine = false;
                break;
            }
        }
        for(nSer = 0; nSer < aSeries.length; ++nSer) {
            oSeries = aSeries[nSer];
            if(oSeries.smooth === false) {
                bSmooth = false;
                break;
            }
        }
        if(bShowMarker) {
            nFlag |= 1;
        }
        if(bNoLine) {
            nFlag |= 2;
        }
        if(bSmooth) {
            nFlag |= 4;
        }
        return nFlag;
    };
    CLineChart.prototype.getChartType = function() {

        var bFlags = this.getLineParams();
        var bShowMarker = (bFlags & 1) === 1;
        var bNoLine = (bFlags & 2) === 2;
        var bSmooth = (bFlags & 4) === 4;
        var nType = Asc.c_oAscChartTypeSettings.unknown;
        switch(this.grouping) {
            case AscFormat.GROUPING_PERCENT_STACKED:
            {
                if(!bShowMarker) {
                    nType = Asc.c_oAscChartTypeSettings.lineStackedPer;
                }
                else {
                    nType = Asc.c_oAscChartTypeSettings.lineStackedPerMarker;
                }
                break;
            }
            case AscFormat.GROUPING_STACKED:
            {
                if(!bShowMarker) {
                    nType = Asc.c_oAscChartTypeSettings.lineStacked;
                }
                else {
                    nType = Asc.c_oAscChartTypeSettings.lineStackedMarker;
                }
                break;
            }
            default:
            {
                var oCS = this.getChartSpace();
                if(oCS && AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(oCS)) {
                    nType = Asc.c_oAscChartTypeSettings.line3d;
                }
                else {
                    if(!bShowMarker) {
                        nType = Asc.c_oAscChartTypeSettings.lineNormal;
                    }
                    else {
                        nType = Asc.c_oAscChartTypeSettings.lineNormalMarker;
                    }
                }
                break;
            }
        }
        return nType;
    };
    CLineChart.prototype.setMarkerValue = function(bVal) {
        var aSeries = this.series, nSeries, oSeries;
        if(bVal) {
            if(!this.marker) {
                this.setMarker(true);
            }
            for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
                oSeries = aSeries[nSeries];
                if(oSeries.marker && oSeries.marker.symbol === AscFormat.SYMBOL_NONE) {
                    oSeries.setMarker(null);
                }
            }
        }
        else {
            if(this.marker) {
                this.setMarker(false);
            }
            for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
                oSeries = aSeries[nSeries];
                if(!oSeries.marker) {
                    oSeries.setMarker(new AscFormat.CMarker());
                }
                if(oSeries.marker.symbol !== AscFormat.SYMBOL_NONE) {
                    oSeries.marker.setSymbol(AscFormat.SYMBOL_NONE);
                }
            }
        }
    };
    CLineChart.prototype.isMarkerChart = function() {
        var bFlags = this.getLineParams();
        var bShowMarker = (bFlags & 1) === 1;
        return bShowMarker;
    };
    CLineChart.prototype.isNoLine = function() {
        var bFlags = this.getLineParams();
        var bNoLine = (bFlags & 2) === 2;
        return bNoLine;
    };
    CLineChart.prototype.isSmooth = function() {
        var bFlags = this.getLineParams();
        var bSmooth = (bFlags & 4) === 4;
        return bSmooth;
    };

    CLineChart.prototype.setLineParams = function(bMarker, bLine, bSmooth) {
        if(!AscFormat.isRealBool(bMarker) || !AscFormat.isRealBool(bLine) || !AscFormat.isRealBool(bSmooth)) {
            return;
        }
        if(bMarker === this.isMarkerChart() && bLine === !this.isNoLine() && bSmooth === this.isSmooth()) {
            return;
        }
        this.setMarkerValue(bMarker);
        var nSer, oSeries;
        if(!bLine)
        {
            for(nSer = 0; nSer < this.series.length; ++nSer)
            {
                oSeries = this.series[nSer];
                oSeries.removeAllDPts();
                if(!oSeries.spPr)
                {
                    oSeries.setSpPr(new AscFormat.CSpPr());
                }

                if(AscFormat.isRealBool(oSeries.smooth))
                {
                    oSeries.setSmooth(null);
                }
                oSeries.spPr.setLn(AscFormat.CreateNoFillLine());
            }
        }
        else
        {
            for(nSer = 0; nSer < this.series.length; ++nSer)
            {
                oSeries = this.series[nSer];
                oSeries.removeAllDPts();
                if(oSeries.smooth !== (bSmooth === true))
                {
                    oSeries.setSmooth(bSmooth === true);
                }
                if(oSeries.spPr && oSeries.spPr.ln)
                {
                    oSeries.spPr.setLn(null);
                }
            }
        }
        if(this.smooth !== (bSmooth === true))
        {
            this.setSmooth(bSmooth === true);
        }
        for(nSer = 0; nSer < this.series.length; ++nSer)
        {
            oSeries = this.series[nSer];
            if(oSeries.smooth !== (bSmooth === true))
            {
                oSeries.setSmooth(bSmooth === true);
            }
        }
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            var oChartStyle = oChartSpace.chartStyle;
            var oChartColors = oChartSpace.chartColors;
            if(oChartStyle && oChartColors) {
                this.applyChartStyle(oChartStyle, oChartColors, null, false);
            }
        }
    };
    CLineChart.prototype.tryChangeType = function(nType) {
        if(!this.parent) {
            return false;
        }
        if(!this.parent.isLineType(nType)) {
            return false;
        }
        if(nType === this.getChartType()) {
            return true;
        }
        var nNewGrouping;
        nNewGrouping = this.parent.getGroupingByType(nType);
        if(this.grouping !== nNewGrouping) {
            this.setGrouping(nNewGrouping);
        }
        var bMarker = this.parent.getIsMarkerByType(nType);
        if(this.marker !== bMarker) {
            this.setMarkerValue(bMarker);
        }
        this.checkValAxesFormatByType(nType);
        this.parent.check3DOptions(nType);
        return true;
    };
	CLineChart.prototype.convert3Dto2D = function() {
		this.setMarker(true);
		this.setSmooth(false);
		for (var i = 0, length = this.series.length; i < length; ++i) {
			var seria = this.series[i];
			if (null == seria.marker) {
				var marker = new AscFormat.CMarker();
				marker.setSymbol(AscFormat.SYMBOL_NONE);
				seria.setMarker(marker);
			}
		}
	};

    function CLineSeries() {
        CSeriesBase.call(this);
        this.cat = null;
        this.dLbls = null;
        this.dPt = [];
        this.errBars = null;
        this.marker = null;
        this.smooth = null;
        this.trendline = null;
        this.val = null;
    }

    InitClass(CLineSeries, CSeriesBase, AscDFH.historyitem_type_LineSeries);
    CLineSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        for(var nDpt = 0; nDpt < this.dPt.length; ++nDpt) {
            aRet.push(this.dPt[nDpt]);
        }
        aRet.push(this.dLbls);
        aRet.push(this.errBars);
        aRet.push(this.marker);
        aRet.push(this.trendline);
        return aRet;
    };
    CLineSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(oCopy.setErrBars && this.errBars)
            oCopy.setErrBars(this.errBars.createDuplicate());
        if(oCopy.setMarker && this.marker)
            oCopy.setMarker(this.marker.createDuplicate());
        if(oCopy.setSmooth && AscFormat.isRealBool(this.smooth))
            oCopy.setSmooth(this.smooth);
        if(oCopy.setTrendline && this.trendline)
            oCopy.setTrendline(this.trendline.createDuplicate());
    };
    CLineSeries.prototype.recalculateBrush = function() {
    };
    CLineSeries.prototype.setCat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetCat, this.cat, pr));
        this.cat = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_LineSeries_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.setErrBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetErrBars, this.errBars, pr));
        this.errBars = pr;
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.setMarker = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetMarker, this.marker, pr));
        this.marker = pr;
        this.setParentToChild(pr);
        this.onChartInternalUpdate();
    };
    CLineSeries.prototype.setSmooth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_LineSeries_SetSmooth, this.smooth, pr));
        this.smooth = pr;
        this.onChartInternalUpdate();
    };
    CLineSeries.prototype.setTrendline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_LineSeries_SetVal, this.val, pr));
        this.val = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CLineSeries.prototype.getPreviewBrush = function() {
        if(this.compiledSeriesPen && this.compiledSeriesPen.Fill) {
            return this.compiledSeriesPen.Fill;
        }
        return null;
    };
    CLineSeries.prototype.checkSeriesAfterChangeType = function() {
	    this.setMarker(new AscFormat.CMarker());
	    this.marker.setSymbol(AscFormat.SYMBOL_NONE);
	    this.setSmooth(false);
	    if(this.spPr && this.spPr.hasNoFillLine()) {
		    this.spPr.setLn(null);
	    }
    };

    var SYMBOL_CIRCLE = 0;
    var SYMBOL_DASH = 1;
    var SYMBOL_DIAMOND = 2;
    var SYMBOL_DOT = 3;
    var SYMBOL_NONE = 4;
    var SYMBOL_PICTURE = 5;
    var SYMBOL_PLUS = 6;
    var SYMBOL_SQUARE = 7;
    var SYMBOL_STAR = 8;
    var SYMBOL_TRIANGLE = 9;
    var SYMBOL_X = 10;

    var MARKER_SYMBOL_TYPE = [];
    MARKER_SYMBOL_TYPE[0] = SYMBOL_DIAMOND;
    MARKER_SYMBOL_TYPE[1] = SYMBOL_SQUARE;
    MARKER_SYMBOL_TYPE[2] = SYMBOL_TRIANGLE;
    MARKER_SYMBOL_TYPE[3] = SYMBOL_X;
    MARKER_SYMBOL_TYPE[4] = SYMBOL_STAR;
    MARKER_SYMBOL_TYPE[5] = SYMBOL_CIRCLE;
    MARKER_SYMBOL_TYPE[6] = SYMBOL_PLUS;
    MARKER_SYMBOL_TYPE[7] = SYMBOL_DOT;
    MARKER_SYMBOL_TYPE[8] = SYMBOL_DASH;

    function CMarker() {
        CBaseChartObject.call(this);
        this.size = null; //2 <= size <= 72
        this.spPr = null;
        this.symbol = null;
    }

    InitClass(CMarker, CBaseChartObject, AscDFH.historyitem_type_Marker);
    CMarker.prototype.getChildren = function() {
        return [this.spPr];
    };
    CMarker.prototype.merge = function(otherMarker) {
        if(isRealObject(otherMarker)) {
            if(AscFormat.isRealNumber(otherMarker.size)) {
                this.setSize(otherMarker.size);
            }
            if(AscFormat.isRealNumber(otherMarker.symbol)) {
                this.setSymbol(otherMarker.symbol);
            }
            if(otherMarker.spPr && (otherMarker.spPr.Fill || otherMarker.spPr.ln)) {
                if(!this.spPr) {
                    this.setSpPr(new AscFormat.CSpPr());
                }
                if(otherMarker.spPr.Fill) {
                    this.spPr.setFill(otherMarker.spPr.Fill.createDuplicate());
                }
                if(otherMarker.spPr.ln) {
                    if(!this.spPr.ln) {
                        this.spPr.setLn(new AscFormat.CLn());
                    }
                    this.spPr.ln.merge(otherMarker.spPr.ln);
                }
            }
        }
        return this;
    };
    CMarker.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.merge(this);
    };
    CMarker.prototype.setSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Marker_SetSize, this.size, pr));
        this.size = pr;
    };
    CMarker.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Marker_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CMarker.prototype.setSymbol = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Marker_SetSymbol, this.symbol, pr));
        this.symbol = pr;
    };
    CMarker.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        var nParentType = this.parent.getObjectType();
        var aColors;
        var nMaxSeriesIdx;
        if(nParentType === AscDFH.historyitem_type_AreaSeries ||
           nParentType === AscDFH.historyitem_type_BarSeries ||
           nParentType === AscDFH.historyitem_type_BubbleSeries ||
           nParentType === AscDFH.historyitem_type_LineSeries ||
           nParentType === AscDFH.historyitem_type_PieSeries ||
           nParentType === AscDFH.historyitem_type_RadarSeries ||
           nParentType === AscDFH.historyitem_type_ScatterSer ||
           nParentType === AscDFH.historyitem_type_SurfaceSeries) {
            nMaxSeriesIdx = this.parent.getMaxSeriesIdx() + 1;
            aColors = oColors.generateColors(nMaxSeriesIdx + 1);
            this.applyStyleEntry(oChartStyle.dataPointMarker, aColors, this.parent.idx, bReset);
        }
        else if(nParentType === AscDFH.historyitem_type_DPt) {
            var oSeries = this.parent.parent;
            if(!oSeries) {
                return;
            }
            if(oSeries.isVaryColors()) {
                var nPtsCount = oSeries.getValuesCount();
                aColors = oColors.generateColors(nPtsCount);
                this.applyStyleEntry(oChartStyle.dataPointMarker, aColors, this.parent.idx, bReset);
            }
            else {
                nMaxSeriesIdx = this.parent.getMaxSeriesIdx() + 1;
                aColors = oColors.generateColors(nMaxSeriesIdx + 1);
                this.applyStyleEntry(oChartStyle.dataPointMarker, aColors, oSeries.idx, bReset);
            }
        }

        var oMarkerLayout = oChartStyle.markerLayout;
        if(oMarkerLayout) {
            if(AscFormat.isRealNumber(oMarkerLayout.symbol)) {
                this.setSymbol(oMarkerLayout.symbol);
            }
            if(AscFormat.isRealNumber(oMarkerLayout.size)) {
                this.setSize(oMarkerLayout.size);
            }
        }
    };

    function CMinusPlus() {
        CBaseChartObject.call(this);
        this.numLit = null;
        this.numRef = null;
    }
    InitClass(CMinusPlus, CBaseChartObject, AscDFH.historyitem_type_MinusPlus);
    CMinusPlus.prototype.getChildren = function() {
        return [this.numRef, this.numLit];
    };
    CMinusPlus.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.numRef) {
            oCopy.setNumRef(this.numRef.createDuplicate());
        }
        if(this.numLit) {
            oCopy.setNumLit(this.numLit.createDuplicate());
        }
    };
    CMinusPlus.prototype.setNumLit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_MinusPlus_SetNumLit, this.numLit, pr));
        this.numLit = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CMinusPlus.prototype.setNumRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_MinusPlus_SetNumRef, this.numRef, pr));
        this.numRef = pr;
        this.setParentToChild(pr);
        this.onChangeDataRefs();
    };
    CMinusPlus.prototype.update = function() {
        if(this.numRef) {
            this.numRef.updateCache();
        }
    };
    CMinusPlus.prototype.clearDataCache = function() {
        if(this.numRef) {
            this.numRef.clearDataCache();
        }
    };
    CMinusPlus.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.numRef) {
            this.numRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CMinusPlus.prototype.getDataRefs = function() {
        if(this.numRef) {
            this.numRef.getDataRefs();
        }
        return new CDataRefs([]);
    };

    function CMultiLvlStrCache() {
        CBaseChartObject.call(this);
        this.lvl = [];
        this.ptCount = null;
    }

    InitClass(CMultiLvlStrCache, CBaseChartObject, AscDFH.historyitem_type_MultiLvlStrCache);
    CMultiLvlStrCache.prototype.getChildren = function() {
        return this.lvl;
    };
    CMultiLvlStrCache.prototype.fillObject = function(oCopy, oIdMap) {
        for(var i = 0; i < this.lvl.length; ++i) {
            oCopy.addLvl(this.lvl[i].createDuplicate());
        }
        oCopy.setPtCount(this.ptCount);
    };
    CMultiLvlStrCache.prototype.addLvl = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_MultiLvlStrCache_SetLvl, this.lvl.length, [pr], true));
        this.lvl.push(pr);
        this.setParentToChild(pr);
    };
    CMultiLvlStrCache.prototype.setPtCount = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_MultiLvlStrCache_SetPtCount, this.ptCount, pr));
        this.ptCount = pr;
    };
    CMultiLvlStrCache.prototype.update = function(sFormula, oSeries) {
        if(!(typeof sFormula === "string" && sFormula.length > 0)) {
            return false;
        }
        var nPtCount = 0;
        var aParsedRef = AscFormat.fParseChartFormula(sFormula);
        if(!Array.isArray(aParsedRef) || aParsedRef.length === 0) {
            return false;
        }
        if(aParsedRef.length > 0) {
            var nRows = 0, nRef, oRef, oBBox, nPtIdx, nCol, oWS, oCell, sVal, nCols = 0, nRow;
            var nLvl, oLvl;
            var bLvlsByRows = true;
            if(oSeries) {
                var sSeriesFormula = oSeries.getValRefFormula();

                if(sSeriesFormula) {
                    if(sSeriesFormula.charAt(0) === '=') {
                        sSeriesFormula = sSeriesFormula.slice(1);
                    }
                    var aParsedSeriesRef = AscFormat.fParseChartFormula(sSeriesFormula);
                    if(aParsedSeriesRef.length === aParsedRef.length) {
                        var oSeriesRef;
                        for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                            oRef = aParsedRef[nRef];
                            oSeriesRef = aParsedSeriesRef[nRef];
							let oSerBB = oSeriesRef.bbox;
							let oRefBB = oSeriesRef.bbox;
							if(oSerBB.r1 > oRefBB.r2 || oRefBB.r1 > oSerBB.r2) {// check empty intersection (bug 59334)
								break;
							}
                        }
                        if(nRef === aParsedRef.length) {
                            bLvlsByRows = false;
                        }
                    }
                }
            }
            if(bLvlsByRows) {
                for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                    oRef = aParsedRef[nRef];
                    oBBox = oRef.bbox;
                    nPtCount += (oBBox.c2 - oBBox.c1 + 1);
                    nRows = Math.max(nRows, oBBox.r2 - oBBox.r1 + 1);
                }
                for(nLvl = 0; nLvl < nRows; ++nLvl) {
                    oLvl = new CStrCache();
                    nPtIdx = 0;
                    for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                        oRef = aParsedRef[nRef];
                        oBBox = oRef.bbox;
                        oWS = oRef.worksheet;
                        if(nLvl < (oBBox.r2 - oBBox.r1 + 1)) {
                            for(nCol = oBBox.c1; nCol <= oBBox.c2; ++nCol) {
                                oCell = oWS.getCell3(nLvl + oBBox.r1, nCol);
                                sVal = oCell.getValueWithFormat();
                                if(typeof sVal === "string" && sVal.length > 0) {
                                    oLvl.addStringPoint(nPtIdx, sVal);
                                }
                                ++nPtIdx;
                            }
                        }
                        else {
                            nPtIdx += (oBBox.c2 - oBBox.c1 + 1);
                        }
                    }
                    nPtCount = Math.max(nPtCount, nPtIdx);
                    oLvl.setPtCount(nPtIdx);
                    this.addLvl(oLvl);
                }
            }
            else {
                for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                    oRef = aParsedRef[nRef];
                    oBBox = oRef.bbox;
                    nPtCount += (oBBox.r2 - oBBox.r1 + 1);
                    nCols = Math.max(nCols, oBBox.c2 - oBBox.c1 + 1);
                }
                for(nLvl = 0; nLvl < nCols; ++nLvl) {
                    oLvl = new CStrCache();
                    nPtIdx = 0;
                    for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                        oRef = aParsedRef[nRef];
                        oBBox = oRef.bbox;
                        oWS = oRef.worksheet;
                        if(nLvl < (oBBox.c2 - oBBox.c1 + 1)) {
                            for(nRow = oBBox.r1; nRow <= oBBox.r2; ++nRow) {
                                oCell = oWS.getCell3(nRow, nLvl + oBBox.c1);
                                sVal = oCell.getValueWithFormat();
                                if(typeof sVal === "string" && sVal.length > 0) {
                                    oLvl.addStringPoint(nPtIdx, sVal);
                                }
                                ++nPtIdx;
                            }
                        }
                        else {
                            nPtIdx += (oBBox.r2 - oBBox.r1 + 1);
                        }
                    }
                    nPtCount = Math.max(nPtCount, nPtIdx);
                    oLvl.setPtCount(nPtIdx);
                    this.addLvl(oLvl);
                }
            }
        }
        this.setPtCount(nPtCount);
        return true;
    };
    CMultiLvlStrCache.prototype.getValues = function(nMaxValues) {
        var ret = [];
        var nEnd = nMaxValues || this.ptCount;
        for(var nIndex = 0; nIndex < nEnd; ++nIndex) {
            var sValue = "";
            for(var nLvl = 0; nLvl < this.lvl.length; ++nLvl) {
                var oLvl = this.lvl[nLvl];
                var oPt = oLvl.getPtByIndex(nIndex);
                if(oPt) {
                    sValue += (oPt.val + " ");
                }
            }
            ret.push(sValue);
        }
        return ret;
    };

    function CChartRefBase() {
        CBaseChartObject.call(this);
        this.f = null;
    }

    InitClass(CChartRefBase, CBaseChartObject, AscDFH.historyitem_type_Unknown);
    CChartRefBase.prototype.onUpdate = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            if(!oChartSpace.recalcInfo.recalculateReferences) {
                oChartSpace.recalcInfo.recalculateReferences = true;
                oChartSpace.handleUpdateInternalChart();
            }
        }
    };
    CChartRefBase.prototype.onUpdateCache = function() {
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.handleUpdateInternalChart();
        }
    };
    CChartRefBase.prototype.Refresh_RecalcData = function(data) {
        if(data && data.Type === AscDFH.historyitem_NumRef_SetF) {
            this.onUpdate();
        }
    };
    CChartRefBase.prototype.getFormula = function() {
        return "=" + this.f;
    };
    CChartRefBase.prototype.getParsedRefs = function() {
        return AscFormat.fParseChartFormula(this.f);
    };
    CChartRefBase.prototype.internalSetFormula = function(pr) {
        this.f = pr;
        this.onChangeDataRefs();
    };
    CChartRefBase.prototype.setF = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_NumRef_SetF, this.f, pr));
        this.internalSetFormula(pr);
        this.onUpdate();
    };
    CChartRefBase.prototype.getDataRefs = function() {
        var oDataRefs = new CDataRefs(this.getParsedRefs());
        oDataRefs.setRef(this);
        return oDataRefs;
    };
    CChartRefBase.prototype.handleOnMoveRange = function(oRangeFrom, oRangeTo) {
        var oDataRefs = this.getDataRefs();
        oDataRefs.move(oRangeFrom, oRangeTo);
        var sFormula = oDataRefs.getFormula();
        if(typeof sFormula === "string" && sFormula.length > 0) {
            this.setF(sFormula);
        }
    };
    CChartRefBase.prototype.handleRemoveWorksheets = function(aNames) {
        var sFormula = this.f;
        var sName;
        if(sFormula !== null) {
            for(var nName = 0; nName < aNames.length; ++nName) {
                sName = aNames[nName];
                sFormula = sFormula.replace(new RegExp(RegExp.escape(sName), 'g'), "#REF");
            }
            if(this.f !== sFormula) {
                this.setF(sFormula);
            }
        }
    };
    CChartRefBase.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(typeof this.f === "string" && this.f.length > 0) {
            var sFormula = this.f;
            if(sFormula.indexOf(sOldSheetName) > -1) {
                this.setF(sFormula.replace(new RegExp(RegExp.escape(sOldSheetName), 'g'), sNewSheetName));
            }
        }
    };
    CChartRefBase.prototype.updateCache = function() {
    };
    CChartRefBase.prototype.clearDataCache = function() {
    };
    CChartRefBase.prototype.updateCacheAndCat = function() {
        if(this.parent && this.parent.getObjectType() === AscDFH.historyitem_type_Cat) {
            this.parent.update();
        }
        else {
            this.updateCache();
        }
    };
    CChartRefBase.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setF(this.f);
    };

    function CMultiLvlStrRef() {
        CChartRefBase.call(this);
        this.multiLvlStrCache = null;
    }

    InitClass(CMultiLvlStrRef, CChartRefBase, AscDFH.historyitem_type_MultiLvlStrRef);
    CMultiLvlStrRef.prototype.getChildren = function() {
        return [this.multiLvlStrCache];
    };
    CMultiLvlStrRef.prototype.fillObject = function(oCopy, oIdMap) {
        CChartRefBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.multiLvlStrCache) {
            oCopy.setMultiLvlStrCache(this.multiLvlStrCache.createDuplicate());
        }
    };
    CMultiLvlStrRef.prototype.setMultiLvlStrCache = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_MultiLvlStrRef_SetMultiLvlStrCache, this.multiLvlStrCache, pr));
        this.multiLvlStrCache = pr;
        this.setParentToChild(pr);
    };
    CMultiLvlStrRef.prototype.clearDataCache = function() {
        if(this.multiLvlStrCache) {
            this.setMultiLvlStrCache(null);
        }
    };
    CMultiLvlStrRef.prototype.updateCache = function(oSeries) {
        AscFormat.ExecuteNoHistory(function() {
            let oCache = new CMultiLvlStrCache();
            if(oCache.update(this.f, oSeries) || !this.multiLvlStrCache) {
                this.setMultiLvlStrCache(oCache);
            }
            this.onUpdateCache();
        }, this, []);
    };
    CMultiLvlStrRef.prototype.getValues = function(nMaxValues) {
        if(!this.multiLvlStrCache) {
            this.updateCache();
        }
        return this.multiLvlStrCache.getValues(nMaxValues);
    };
    CMultiLvlStrRef.prototype.getFirstLvlCache = function() {
        var oCache = this.multiLvlStrCache;
        if(oCache && oCache.lvl[0]) {
            let oFirst = oCache.lvl[0];
            let oThis = this;
            return AscFormat.ExecuteNoHistory(function () {
                let oCopy = new CStrCache();
                oCopy.ptCount = oCache.ptCount;
                oCopy.pts = oFirst.pts;
                return oCopy;
            });
        }
        return null;
    };

    function isObjectSeries(oObject) {
        return oObject && (oObject.superclass === CSeriesBase);
    }

    function CNumRef() {
        CChartRefBase.call(this);
        this.numCache = null;
    }

    InitClass(CNumRef, CChartRefBase, AscDFH.historyitem_type_NumRef);
    CNumRef.prototype.getChildren = function() {
        return [this.numCache];
    };
    CNumRef.prototype.fillObject = function(oCopy, oIdMap) {
        CChartRefBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.numCache) {
            oCopy.setNumCache(this.numCache.createDuplicate());
        }
    };
    CNumRef.prototype.setNumCache = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_NumRef_SetNumCache, this.numCache, pr));
        this.numCache = pr;
        this.setParentToChild(pr);
    };
    CNumRef.prototype.clearDataCache = function() {
        if(this.numCache) {
            this.setNumCache(null);
        }
    };
    CNumRef.prototype.updateCache = function(displayEmptyCellsAs, displayHidden, ser) {
        AscFormat.ExecuteNoHistory(function() {
            if(!this.numCache) {
                this.setNumCache(new CNumLit());
                this.numCache.setFormatCode("General");
                this.numCache.setPtCount(0);
            }
            let oSeries = null;
            if(ser) {
                oSeries = ser;
            }
            else {
                let oCurObjectForCheck = this;
                while (!oSeries && oCurObjectForCheck) {
                    if(isObjectSeries(oCurObjectForCheck)) {
                        oSeries = oCurObjectForCheck;
                    }
                    else {
                        oCurObjectForCheck = oCurObjectForCheck.parent;
                    }
                }
            }
            if(oSeries) {
                oSeries.isHidden = true;
            }
            this.numCache.update(this.f, displayEmptyCellsAs, displayHidden, oSeries);
            this.onUpdateCache();
        }, this, []);
    };
    CNumRef.prototype.getValuesCount = function() {
        if(!this.numCache) {
            this.updateCache();
        }
        return this.numCache.ptCount;
    };
    CNumRef.prototype.getValues = function(nMaxValues) {
        if(!this.numCache) {
            this.updateCache();
        }
        return this.numCache.getValues(nMaxValues);
    };
    CNumRef.prototype.getNumFormat = function(nPtIdx) {
        if(!this.numCache) {
            return "General";
        }
        return this.numCache.getNumFormat(nPtIdx);
    };
    CNumRef.prototype.fillFromAsc = function(oValCache, bUseCache) {
        this.setF(oValCache.Formula);
        if(bUseCache) {
            this.setNumCache(new AscFormat.CNumLit());
            var oNumCache = this.numCache;
            oNumCache.setPtCount(oValCache.NumCache.length);
            for(var nPt = 0; nPt < oValCache.NumCache.length; ++nPt) {
                var oAscPt = oValCache.NumCache[nPt];
                var oPt = new AscFormat.CNumericPoint();
                oPt.setIdx(nPt);
                oPt.setFormatCode(oAscPt.numFormatStr);
                oPt.setVal(parseFloat(oAscPt.val));
                oNumCache.addPt(oPt);
            }
        }
    };

    function CStrRef() {
        CChartRefBase.call(this);
        this.strCache = null;
    }

    InitClass(CStrRef, CChartRefBase, AscDFH.historyitem_type_StrRef);
    CStrRef.prototype.getChildren = function() {
        return [this.strCache];
    };
    CStrRef.prototype.fillObject = function(oCopy, oIdMap) {
        CChartRefBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.strCache) {
            oCopy.setStrCache(this.strCache.createDuplicate());
        }
    };
    CStrRef.prototype.setStrCache = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_StrRef_SetStrCache, this.strCache, pr));
        this.strCache = pr;
        this.setParentToChild(pr);
    };
    CStrRef.prototype.clearDataCache = function() {
        if(this.strCache) {
            this.setStrCache(null);
        }
    };
    CStrRef.prototype.updateCache = function() {
        AscFormat.ExecuteNoHistory(function() {
            if(!this.strCache) {
                this.setStrCache(new CStrCache());
            }
            this.strCache.update(this.f);
            this.onUpdateCache();
        }, this, []);
    };
    CStrRef.prototype.getText = function(bNoUpdate) {
        if(!this.strCache) {
            if(bNoUpdate !== true) {
                this.updateCache();
            }
        }
        if(!this.strCache) {
            return "";
        }
        var aValues = this.strCache.getValues(null);
        var sRet = "";
        for(var i = 0; i < aValues.length; ++i) {
            if(i > 0) {
                sRet += " ";
            }
            sRet += aValues[i];
        }
        return sRet;
    };
    CStrRef.prototype.getValues = function(nMaxCount, bAddEmpty) {
        if(!this.strCache) {
            this.updateCache();
        }
        return this.strCache.getValues(nMaxCount, bAddEmpty);
    };
    CStrRef.prototype.fillFromAsc = function(oCatCache, bUseCache) {
        this.setF(oCatCache.Formula);
        if(bUseCache) {
            this.setStrCache(new AscFormat.CStrCache());
            var oStrCache = this.strCache;
            oStrCache.setPtCount(oCatCache.NumCache.length);
            for(var nPt = 0; nPt < oCatCache.NumCache.length; ++nPt) {
                var oAscPt = oCatCache.NumCache[nPt];
                var oPt = new AscFormat.CStringPoint();
                oPt.setIdx(nPt);
                oPt.setVal(oAscPt.val + "");
                oStrCache.addPt(oPt);
            }
        }
    };

    function CNumericPoint() {
        CBaseChartObject.call(this);
        this.formatCode = null;
        this.idx = null;
        this.val = null;
    }

    InitClass(CNumericPoint, CBaseChartObject, AscDFH.historyitem_type_NumericPoint);
    CNumericPoint.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.formatCode !== null) {
            oCopy.setFormatCode(this.formatCode);
        }
        if(this.idx !== null) {
            oCopy.setIdx(this.idx);
        }
        if(this.val !== null) {
            oCopy.setVal(this.val);
        }
    };
    CNumericPoint.prototype.setFormatCode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_NumericPoint_SetFormatCode, this.formatCode, pr));
        this.formatCode = pr;
    };
    CNumericPoint.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_NumericPoint_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CNumericPoint.prototype.setVal = function(pr) {
        var _pr = parseFloat(pr);
        if(isNaN(_pr)) {
            _pr = 0;
        }
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_NumericPoint_SetVal, this.val, _pr));
        this.val = _pr;
    };
    CNumericPoint.prototype.getFormatCode = function() {
        if(typeof this.formatCode === "string" && this.formatCode.length > 0) {
            return this.formatCode;
        }
        return "General";
    };

    function CNumFmt() {
        CBaseChartObject.call(this);
        this.formatCode = null;
        this.sourceLinked = null;
    }

    InitClass(CNumFmt, CBaseChartObject, AscDFH.historyitem_type_NumFmt);
    CNumFmt.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.formatCode !== null) {
            oCopy.setFormatCode(this.formatCode);
        }
        if(this.sourceLinked !== null) {
            oCopy.setSourceLinked(this.sourceLinked);
        }
    };
    CNumFmt.prototype.setFormatCode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_NumFmt_SetFormatCode, this.formatCode, pr));
        this.formatCode = pr;
    };
    CNumFmt.prototype.setSourceLinked = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_NumFmt_SetSourceLinked, this.sourceLinked, pr));
        this.sourceLinked = pr;
    };
    CNumFmt.prototype.setFromAscObject = function(o) {
        if(!o || !o.isCorrect()) {
            return;
        }
        if(o.formatCode !== this.formatCode) {
            this.setFormatCode(o.formatCode);
        }
        if(o.sourceLinked !== this.sourceLinked) {
            this.setSourceLinked(o.sourceLinked);
        }
    };

    function CNumLit() {
        CBaseChartObject.call(this);
        this.formatCode = null;
        this.pts = [];
        this.ptCount = null;
    }

    InitClass(CNumLit, CBaseChartObject, AscDFH.historyitem_type_NumLit);
    CNumLit.prototype.removeDPt = function(idx) {
        if(this.pts[idx]) {
            this.pts[idx].setParent(null);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonLit_RemoveDPt, idx, [this.pts[idx]], false));
            this.pts.splice(idx, 1);
        }
    };
    //CNumLit.prototype.getChildren = function() {
    //    return this.pts;
    //};
    CNumLit.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setFormatCode(this.formatCode);
        for(var i = 0; i < this.pts.length; ++i) {
            oCopy.addPt(this.pts[i].createDuplicate());
        }
        oCopy.setPtCount(this.ptCount);
    };
    CNumLit.prototype.getPtByIndex = function(idx) {
        if(this.pts[idx] && this.pts[idx].idx === idx)
            return this.pts[idx];
        for(var i = 0; i < this.pts.length; ++i) {
            if(this.pts[i].idx === idx)
                return this.pts[i];
        }
        return null;
    };
    CNumLit.prototype.getPtCount = function() {
        if(AscFormat.isRealNumber(this.ptCount)) {
            return this.ptCount;
        }
        return this.pts.length;
    };
    CNumLit.prototype.setFormatCode = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_NumLit_SetFormatCode, this.formatCode, pr));
        this.formatCode = pr;
    };
    CNumLit.prototype.addPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_NumLit_AddPt, this.pts.length, [pr], true));
        this.pts.push(pr);
        this.setParentToChild(pr);
    };
    CNumLit.prototype.setPtCount = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_NumLit_SetPtCount, this.ptCount, pr));
        this.ptCount = pr;
    };
    CNumLit.prototype.removeAllPts = function() {
        var start_idx = this.pts.length - 1;
        for(var i = start_idx; i > -1; --i) {
            this.removeDPt(i);
        }
        this.setPtCount(0);
    };
    CNumLit.prototype.addNumericPoint = function(idx, dNumber) {
        var oPt = new CNumericPoint();
        oPt.setIdx(idx);
        oPt.setVal(dNumber);
        this.addPt(oPt);
        this.setPtCount(Math.max(this.pts.length, idx + 1));
    };
    CNumLit.prototype.update = function(sFormula, displayEmptyCellsAs, displayHidden, ser) {
        AscFormat.ExecuteNoHistory(function() {
            if(!(typeof sFormula === "string" && sFormula.length > 0)) {
                this.setPtCount(0);
                this.setFormatCode("General");
                return;
            }
            var aParsedRef = AscFormat.fParseChartFormula(sFormula);
            if(!Array.isArray(aParsedRef) || aParsedRef.length === 0) {
                if(ser) {
                    ser.isHidden = (this.pts.length === 0);
                }
                return;
            }
            this.removeAllPts();
            var nRef, oRef, oMinRef, oWS, oBB, nPtIdx, nPtCount, nCount;
            var oHM, nHidden, nR, nC, oMinBB, oCell, aSpanPoints = [], oStartPoint;
            var dVal, sVal, oPt, sCellFC, nSpan, nLastNoEmptyIndex, nSpliceIndex;
            var sCommonFormatCode = null;
            var bHaveTotalRow;
            nPtCount = 0;
            nPtIdx = 0;
            for(nRef = 0; nRef < aParsedRef.length; ++nRef) {
                oRef = aParsedRef[nRef];
                oWS = oRef.worksheet;
                if(!oWS) {
                    continue;
                }
                oHM = oWS.hiddenManager;
                oBB = oRef.bbox;
                if(!oBB) {
                    continue;
                }
                if(displayHidden) {
                    nHidden = 0;
                }
                else {
                    nHidden = oHM.getHiddenRowsCount(oBB.r1, oBB.r2 + 1) * oBB.getWidth() + oHM.getHiddenColsCount(oBB.c1, oBB.c2 + 1) * oBB.getHeight();
                }
                bHaveTotalRow = oWS.isTableTotalRow(new Asc.Range(oBB.c1, oBB.r2, oBB.c1, oBB.r2)) && oBB.r1 !== oBB.r2;
                if(bHaveTotalRow) {
                    nHidden += 1;
                }
                nCount = Math.max(0, oBB.getHeight() * oBB.getWidth() - nHidden);
                nPtCount += nCount;
                oMinRef = oRef.getMinimalCellsRange();
                if(!oMinRef) {
                    continue;
                }
                oMinBB = oMinRef.bbox;
                nPtIdx += ((oMinBB.r1 - oBB.r1) * oBB.getWidth() + (oMinBB.c1 - oBB.c1) * oMinBB.getHeight());
                for(nR = oMinBB.r1; nR <= oMinBB.r2; ++nR) {
                    if(!displayHidden && oWS.getRowHidden(nR)) {
                        continue;
                    }
                    if(nR === oBB.r2 && bHaveTotalRow) {
                        continue;
                    }
                    for(nC = oMinBB.c1; nC <= oMinBB.c2; ++nC) {
                        if(!displayHidden && oWS.getColHidden(nC)) {
                            continue;
                        }
                        if(ser) {
                            ser.isHidden = false;
                        }
                        oCell = oWS.getCell3(nR, nC);
                        dVal = oCell.getNumberValue();
                        if(!AscFormat.isRealNumber(dVal) && (!AscFormat.isRealNumber(displayEmptyCellsAs) || displayEmptyCellsAs === 1)) {
                            sVal = oCell.getValueForEdit();
                            if((typeof sVal === "string") && sVal.length > 0) {
                                dVal = 0;
                            }
                        }
                        if(AscFormat.isRealNumber(dVal)) {
                            oPt = new AscFormat.CNumericPoint();
                            oPt.setIdx(nPtIdx);
                            oPt.setVal(dVal);
                            sCellFC = oCell.getNumFormatStr();
                            oPt.setFormatCode(sCellFC);
                            if(sCommonFormatCode === null) {
                                sCommonFormatCode = sCellFC;
                            }
                            else {
                                if(sCommonFormatCode !== undefined) {
                                    if(sCommonFormatCode !== sCellFC) {
                                        sCommonFormatCode = undefined;
                                    }
                                }
                            }
                            this.addPt(oPt);
                            if(aSpanPoints.length > 0) {
                                if(AscFormat.isRealNumber(nLastNoEmptyIndex)) {
                                    oStartPoint = this.getPtByIndex(nLastNoEmptyIndex);
                                    if(oStartPoint) {
                                        for(nSpan = 0; nSpan < aSpanPoints.length; ++nSpan) {
                                            aSpanPoints[nSpan].setVal(oStartPoint.val + ((oPt.val - oStartPoint.val) / (aSpanPoints.length + 1)) * (nSpan + 1));
                                            this.pts.splice(nSpliceIndex + nSpan, 0, aSpanPoints[nSpan]);
                                        }
                                    }
                                }
                                aSpanPoints.length = 0;
                            }
                            nLastNoEmptyIndex = oPt.idx;
                            nSpliceIndex = this.pts.length;
                        }
                        else {
                            if(AscFormat.isRealNumber(displayEmptyCellsAs) && displayEmptyCellsAs !== 1) {
                                sVal = oCell.getValue();
                                if(displayEmptyCellsAs === 2 || ((typeof sVal === "string") && sVal.length > 0)) {
                                    oPt = new AscFormat.CNumericPoint();
                                    oPt.setIdx(nPtIdx);
                                    oPt.setVal(0);
                                    this.addPt(oPt);
                                    if(aSpanPoints.length > 0) {
                                        if(AscFormat.isRealNumber(nLastNoEmptyIndex)) {
                                            oStartPoint = this.getPtByIndex(nLastNoEmptyIndex);
                                            if(oStartPoint) {
                                                for(nSpan = 0; nSpan < aSpanPoints.length; ++nSpan) {
                                                    aSpanPoints[nSpan].setVal(oStartPoint.val + ((oPt.val - oStartPoint.val) / (aSpanPoints.length + 1)) * (nSpan + 1));
                                                    this.pts.splice(nSpliceIndex + nSpan, 0, aSpanPoints[nSpan]);
                                                }
                                            }
                                        }
                                        aSpanPoints.length = 0;
                                    }
                                    nLastNoEmptyIndex = oPt.idx;
                                    nSpliceIndex = this.pts.length;
                                }
                                else if(displayEmptyCellsAs === 0 && ser && ser.getObjectType() === AscDFH.historyitem_type_LineSeries) {
                                    oPt = new AscFormat.CNumericPoint();
                                    oPt.setIdx(nPtIdx);
                                    oPt.setVal(0);
                                    aSpanPoints.push(oPt);
                                }
                            }
                        }
                        nPtIdx++;
                    }
                }
            }
            this.setPtCount(nPtCount);
            //check format
            if(sCommonFormatCode) {
                this.setFormatCode(sCommonFormatCode);
                for(var nPt = 0; nPt < this.pts.length; ++nPt) {
                    oPt = this.pts[nPt];
                    oPt.setFormatCode(null);
                }
            }
            else {
                this.setFormatCode("General");
            }
        }, this, []);
    };
    CNumLit.prototype.getValues = function(nMaxCount) {
        var ret = [];

        var nEnd = nMaxCount || this.ptCount;
        for(var nIndex = 0; nIndex < nEnd; ++nIndex) {
            var oPt = this.getPtByIndex(nIndex);
            if(oPt) {
                var sFormatCode = oPt.formatCode || this.formatCode;
                if(sFormatCode) {
                    var oNumFmt = AscCommon.oNumFormatCache.get(sFormatCode);
                    ret.push(oNumFmt.formatToChart(oPt.val));
                }
                else {
                    ret.push(oPt.val + "");
                }

            }
            else {
                ret.push("");
            }
        }
        return ret;
    };
    CNumLit.prototype.getFormula = function() {
        var sRet = "={";
        for(var nIndex = 0; nIndex < this.pts.length; ++nIndex) {
            sRet += this.pts[nIndex].val;
            if(nIndex < this.pts.length - 1) {
                sRet += ", ";
            }
        }
        sRet += "}";
        return sRet;
    };
    CNumLit.prototype.getNumFormat = function(nPtIdx) {
        var nIdx = AscFormat.isRealNumber(nPtIdx) ? nPtIdx : 0;
        var oPt = this.pts[nIdx];
        if(oPt) {
            if(typeof oPt.formatCode === "string" && oPt.formatCode.length > 0) {
                return oPt.formatCode;
            }
        }
        if(typeof this.formatCode === "string" && this.formatCode.length > 0) {
            return this.formatCode;
        }
        return "General";
    };

    function COfPieChart() {
        CChartBase.call(this);
        this.custSplit = [];
        this.gapWidth = null;
        this.ofPieType = null;
        this.secondPieSize = null;
        this.serLines = null;
        this.splitPos = null;
        this.splitType = null;
    }

    InitClass(COfPieChart, CChartBase, AscDFH.historyitem_type_OfPieChart);
    COfPieChart.prototype.getEmptySeries = function() {
        return new CPieSeries();
    };
    COfPieChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.bestFit;
    };
    COfPieChart.prototype.getChildren = function() {
        var aRet = CChartBase.prototype.getChildren.call(this);
        for(var i = 0; i < this.custSplit.length; ++i) {
            aRet.push(this.custSplit[i]);
        }
        aRet.push(this.serLines);
        return aRet;
    };
    COfPieChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        var i;
        if(oCopy.addCustSplit) {
            for(i = 0; i < this.custSplit.length; ++i) {
                oCopy.addCustSplit(this.custSplit[i].createDuplicate());
            }
        }
        if(this.gapWidth !== null && oCopy.setGapWidth) {
            oCopy.setGapWidth(this.gapWidth);
        }
        if(this.ofPieType !== null && oCopy.setOfPieType) {
            oCopy.setOfPieType(this.ofPieType);
        }
        if(this.secondPieSize !== null && oCopy.setSecondPieSize) {
            oCopy.setSecondPieSize(this.secondPieSize);
        }
        if(this.serLines !== null && oCopy.setSerLines) {
            oCopy.setSerLines(this.serLines.createDuplicate());
        }
        if(this.splitPos !== null && oCopy.setSplitPos) {
            oCopy.setSplitPos(this.splitPos);
        }
        if(this.splitType !== null && oCopy.setSplitType) {
            oCopy.setSplitType(this.splitType);
        }
    };
    COfPieChart.prototype.getAllRasterImages = function(images) {
        CChartBase.prototype.getAllRasterImages.call(this, images);
        if(this.serLines && this.serLines.spPr) {
            this.serLines.spPr.checkBlipFillRasterImage(images);
        }
    };
    COfPieChart.prototype.checkSpPrRasterImages = function(images) {
        CChartBase.prototype.checkSpPrRasterImages.call(this, images);
        if(this.serLines && this.serLines.spPr) {
            checkSpPrRasterImages(this.serLines.spPr);
        }
    };
    COfPieChart.prototype.addCustSplit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_OfPieChart_AddCustSplit, this.custSplit.length, [pr], true));
        this.custSplit.push(pr);
        this.setParentToChild(pr);
    };
    COfPieChart.prototype.setGapWidth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_OfPieChart_SetGapWidth, this.gapWidth, pr));
        this.gapWidth = pr;
    };
    COfPieChart.prototype.setOfPieType = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_OfPieChart_SetOfPieType, this.ofPieType, pr));
        this.ofPieType = pr;
    };
    COfPieChart.prototype.setSecondPieSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_OfPieChart_SetSecondPieSize, this.secondPieSize, pr));
        this.secondPieSize = pr;
    };
    COfPieChart.prototype.setSerLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_OfPieChart_SetSerLines, this.serLines, pr));
        this.serLines = pr;
        this.setParentToChild(pr);
    };
    COfPieChart.prototype.setSplitPos = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_OfPieChart_SetSplitPos, this.splitPos, pr));
        this.splitPos = pr;
    };
    COfPieChart.prototype.setSplitType = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_OfPieChart_SetSplitType, this.splitType, pr));
        this.splitType = pr;
    };
    COfPieChart.prototype.convertToPieChart = function() {
        var oPie = new AscFormat.CPieChart();
        if (null != this.varyColors)
            oPie.setVaryColors(this.varyColors);
        if (null != this.dLbls)
            oPie.setDLbls(this.dLbls);
        for (var i = 0, length = this.series.length; i < length; ++i) {
            oPie.addSer(this.series[i]);
        }
        oPie.setFirstSliceAng(0);
        return oPie;
    };

    function CPictureOptions() {
        CBaseChartObject.call(this);
        this.applyToEnd = null;
        this.applyToFront = null;
        this.applyToSides = null;
        this.pictureFormat = null;
        this.pictureStackUnit = null;
    }

    InitClass(CPictureOptions, CBaseChartObject, AscDFH.historyitem_type_PictureOptions);
    CPictureOptions.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.applyToEnd !== null) {
            oCopy.setApplyToEnd(this.applyToEnd);
        }
        if(this.applyToFront !== null) {
            oCopy.setApplyToFront(this.applyToFront);
        }
        if(this.applyToSides !== null) {
            oCopy.setApplyToSides(this.applyToSides);
        }
        if(this.pictureFormat !== null) {
            oCopy.setPictureFormat(this.pictureFormat);
        }
        if(this.pictureStackUnit !== null) {
            oCopy.setPictureStackUnit(this.pictureStackUnit);
        }
    };
    CPictureOptions.prototype.setApplyToEnd = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PictureOptions_SetApplyToEnd, this.applyToEnd, pr));
        this.applyToEnd = pr;
    };
    CPictureOptions.prototype.setApplyToFront = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PictureOptions_SetApplyToFront, this.applyToFront, pr));
        this.applyToFront = pr;
    };
    CPictureOptions.prototype.setApplyToSides = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PictureOptions_SetApplyToSides, this.applyToSides, pr));
        this.applyToSides = pr;
    };
    CPictureOptions.prototype.setPictureFormat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PictureOptions_SetPictureFormat, this.pictureFormat, pr));
        this.pictureFormat = pr;
    };
    CPictureOptions.prototype.setPictureStackUnit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PictureOptions_SetPictureStackUnit, this.pictureStackUnit, pr));
        this.pictureStackUnit = pr;
    };

    function CPieChart() {
        CChartBase.call(this);
        this.firstSliceAng = null;
        this.b3D = null;
    }

    InitClass(CPieChart, CChartBase, AscDFH.historyitem_type_PieChart);
    CPieChart.prototype.set3D = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PieChart_3D, this.b3D, pr));
        this.b3D = pr;
    };
    CPieChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;

        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_PieChart_SetDLbls:
            case AscDFH.historyitem_CommonChart_SetDlbls:
            {
                this.onChartUpdateDataLabels();
                break;
            }
            case AscDFH.historyitem_PieChart_SetFirstSliceAng:
            {
                break;
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_PieChart_SetVaryColors:
            case AscDFH.historyitem_CommonChart_SetVaryColors:
            {
                break;
            }
        }
    };
    CPieChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.bestFit;
    };
    CPieChart.prototype.getEmptySeries = function() {
        return new CPieSeries();
    };
    CPieChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.firstSliceAng !== null && oCopy.setFirstSliceAng) {
            oCopy.setFirstSliceAng(this.firstSliceAng);
        }
        if(this.b3D && oCopy.set3D) {
            oCopy.set3D(this.b3D);
        }
    };
    CPieChart.prototype.setFirstSliceAng = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PieChart_SetFirstSliceAng, this.firstSliceAng, pr));
        this.firstSliceAng = pr;
    };
    CPieChart.prototype.getChartType = function() {
        var oCS = this.getChartSpace();
        if(oCS && AscFormat.CChartsDrawer.prototype._isSwitchCurrent3DChart(oCS)) {
            return Asc.c_oAscChartTypeSettings.pie3d;
        }
        else {
            return Asc.c_oAscChartTypeSettings.pie;
        }
    };
    CPieChart.prototype.tryChangeType = function(nType) {
        if(!this.parent) {
            return false;
        }
        if(!this.parent.isPieType(nType)) {
            return false;
        }
        if(nType === this.getChartType()) {
            return true;
        }
        this.parent.check3DOptions(nType);
        return true;
    };

    function CPieSeries() {
        CSeriesBase.call(this);
        this.cat = null;
        this.dLbls = null;
        this.dPt = [];
        this.explosion = null;
        this.val = null;
    }

    InitClass(CPieSeries, CSeriesBase, AscDFH.historyitem_type_PieSeries);
    CPieSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        for(var i = 0; i < this.dPt.length; ++i) {
            aRet.push(this.dPt[i]);
        }
        return aRet;
    };
    CPieSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.explosion !== null && oCopy.setExplosion) {
            oCopy.setExplosion(this.explosion);
        }
    };
    CPieSeries.prototype.setCat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PieSeries_SetCat, this.cat, pr));
        this.cat = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CPieSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PieSeries_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CPieSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_PieSeries_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CPieSeries.prototype.setExplosion = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PieSeries_SetExplosion, this.explosion, pr));
        this.explosion = pr;
    };
    CPieSeries.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PieSeries_SetVal, this.val, pr));
        this.val = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
	CPieSeries.prototype.checkSeriesAfterChangeType = function() {
		if(this.spPr) {
			var oSpPr = this.spPr;
			if(oSpPr.Fill) {
				oSpPr.setFill(null);
			}
			if(oSpPr.ln) {
				oSpPr.setLn(null);
			}
		}
	};
    function CPivotFmt() {
        CBaseChartObject.call(this);
        this.dLbl = null;
        this.idx = null;
        this.marker = null;
        this.spPr = null;
        this.txPr = null;
    }

    InitClass(CPivotFmt, CBaseChartObject, AscDFH.historyitem_type_PivotFmt);
    CPivotFmt.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        aRet.push(this.dLbl);
        aRet.push(this.marker);
        aRet.push(this.spPr);
        aRet.push(this.txPr);
        return aRet;
    };
    CPivotFmt.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.dLbl) {
            oCopy.setLbl(this.dLbl.createDuplicate());
        }
        if(this.idx !== null) {
            oCopy.setIdx(this.idx);
        }
        if(this.marker) {
            oCopy.setMarker(this.marker.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CPivotFmt.prototype.setLbl = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PivotFmt_SetDLbl, this.dLbl, pr));
        this.dLbl = pr;
        this.setParentToChild(pr);
    };
    CPivotFmt.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PivotFmt_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CPivotFmt.prototype.setMarker = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PivotFmt_SetMarker, this.marker, pr));
        this.marker = pr;
        this.setParentToChild(pr);
    };
    CPivotFmt.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PivotFmt_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CPivotFmt.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PivotFmt_SetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
    };

    function CRadarChart() {
        CLineChart.call(this);
        this.radarStyle = null;
    }

    InitClass(CRadarChart, CLineChart, AscDFH.historyitem_type_RadarChart);
    CRadarChart.prototype.getEmptySeries = function() {
        return new CRadarSeries();
    };
    CRadarChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.outEnd;
    };
    CRadarChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_RadarChart_AddAxId:
            case AscDFH.historyitem_CommonChart_AddAxId:
            {
                break;
            }
            case AscDFH.historyitem_RadarChart_SetDLbls:
            case AscDFH.historyitem_CommonChart_SetDlbls:
            {
                this.onChartUpdateDataLabels();
                break;
            }
            case AscDFH.historyitem_RadarChart_SetRadarStyle:
            {
                break;
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_RadarChart_SetVaryColors:
            case AscDFH.historyitem_CommonChart_SetVaryColors:
            {
                break;
            }
        }
    };
    CRadarChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(oCopy.setRadarStyle && this.radarStyle !== null) {
            oCopy.setRadarStyle(this.radarStyle);
        }
    };
    CRadarChart.prototype.setRadarStyle = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_RadarChart_SetRadarStyle, this.radarStyle, pr));
        this.radarStyle = pr;
    };
	CRadarChart.prototype.isFilled = function() {
		return AscFormat.RADAR_STYLE_FILLED === this.radarStyle;
	};
    CRadarChart.prototype.convertToLineChart = function() {
        var bMarkerNull = this.isFilled();
        var oLine = new AscFormat.CLineChart();
        oLine.setGrouping(AscFormat.GROUPING_STANDARD);
        if (null != this.varyColors)
            oLine.setVaryColors(this.varyColors);
        if (null != this.dLbls)
            oLine.setDLbls(this.dLbls);
        for (var i = 0, length = this.series.length; i < length; ++i) {
            var radarSer = this.series[i];
            var lineSer = new AscFormat.CLineSeries();
            if (null != radarSer.idx)
                lineSer.setIdx(radarSer.idx);
            if (null != radarSer.order)
                lineSer.setOrder(radarSer.order);
            if (null != radarSer.tx)
                lineSer.setTx(radarSer.tx);
            if (null != radarSer.spPr)
                lineSer.setSpPr(radarSer.spPr);
            if (null != radarSer.marker)
                lineSer.setMarker(radarSer.marker);
            else if(bMarkerNull){
                var marker = new AscFormat.CMarker();
                marker.setSymbol(AscFormat.SYMBOL_NONE);
                lineSer.setMarker(marker);
            }

            for (var j = 0, length2 = radarSer.dPt.length; j < length2; ++j) {
                lineSer.addDPt(radarSer.dPt[j]);
            }
            if (null != radarSer.dLbls)
                lineSer.setDLbls(radarSer.dLbls);
            if (null != radarSer.cat)
                lineSer.setCat(radarSer.cat);
            if (null != radarSer.val)
                lineSer.setVal(radarSer.val);
            lineSer.setSmooth(false);
            oLine.addSer(lineSer);
        }
        oLine.setMarker(true);
        oLine.setSmooth(false);
        return oLine;
    };
	CRadarChart.prototype.getChartType = function() {
		if(this.isFilled()) {
			return Asc.c_oAscChartTypeSettings.radarFilled;
		}
		if(this.isMarkerChart()) {
			return Asc.c_oAscChartTypeSettings.radarMarker;
		}
		return Asc.c_oAscChartTypeSettings.radar;
	};
	CRadarChart.prototype.setDLbls = function(oPr) {
		if(oPr) {
			if(oPr.dLblPos !== null) {
				oPr.setDLblPos(null);
			}
		}
		CChartBase.prototype.setDLbls.call(this, oPr);
	};
	CRadarChart.prototype.tryChangeType = function(nType) {
		if(!this.parent) {
			return false;
		}
		if(!this.parent.isRadarType(nType)) {
			return false;
		}
		if(nType === this.getChartType()) {
			return true;
		}
		let nRadarStyle;
		if(nType === Asc.c_oAscChartTypeSettings.radarFilled) {
			nRadarStyle = AscFormat.RADAR_STYLE_FILLED;
		}
		else {
			nRadarStyle = AscFormat.RADAR_STYLE_MARKER;//excel uses this style for no marker chart
		}
		if(this.radarStyle !== nRadarStyle) {
			this.setRadarStyle(nRadarStyle);
		}
		const aSeries = this.series;

		for(let nSer = 0; nSer < aSeries.length; ++nSer) {
			aSeries[nSer].changeChartType(nType);
		}
		this.checkValAxesFormatByType(nType);
		this.parent.check3DOptions(nType);
		return true;
	};
	CRadarChart.prototype.getLineParams = function() {
		let nFlag = 0;//
		let bShowMarker = false;
		let bNoLine = true;
		let aSeries = this.series, oSeries;
		let nSer;
		if(this.radarStyle === AscFormat.RADAR_STYLE_MARKER) {
			for(nSer = 0; nSer < aSeries.length; ++nSer) {
				oSeries = aSeries[nSer];
				if(!oSeries.marker) {
					bShowMarker = true;
					break;
				}
				if(oSeries.marker.symbol !== AscFormat.SYMBOL_NONE) {
					bShowMarker = true;
					break;
				}
			}
		}
		for(nSer = 0; nSer < aSeries.length; ++nSer) {
			oSeries = aSeries[nSer];
			if(!oSeries.hasNoFillLine()) {
				bNoLine = false;
				break;
			}
		}
		if(bShowMarker) {
			nFlag |= 1;
		}
		if(bNoLine) {
			nFlag |= 2;
		}
		return nFlag;
	};
	CRadarChart.prototype.setGrouping = function(val) {

	};

    function CRadarSeries() {
        CLineSeries.call(this);
        this.cat = null;
        this.dLbls = null;
        this.dPt = [];
        this.marker = null;
        this.val = null;

    }

    InitClass(CRadarSeries, CLineSeries, AscDFH.historyitem_type_RadarSeries);
    CRadarSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren(this);
        aRet.push(this.dLbls);
        for(var i = 0; i < this.dPt.length; ++i) {
            aRet.push(this.dPt[i]);
        }
        aRet.push(this.marker);
        return aRet;
    };
    CRadarSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.marker && oCopy.setMarker) {
            oCopy.setMarker(this.marker.createDuplicate());
        }
    };
    CRadarSeries.prototype.setDLbls = function(pr) {
	    if(pr) {
		    if(pr.dLblPos !== null) {
			    pr.setDLblPos(null);
		    }
	    }
	    CLineSeries.prototype.setDLbls.call(this, pr);
    };
	CRadarSeries.prototype.changeChartType = function(nType) {
		let oSpPr = this.spPr;
		if(nType === Asc.c_oAscChartTypeSettings.radarFilled) {
			if(!oSpPr) {
				oSpPr = new AscFormat.CSpPr();
				this.setSpPr(oSpPr);
			}
			oSpPr.setLn(AscFormat.CreateNoFillLine());
			if(oSpPr.hasNoFill()) {
				oSpPr.setFill(null);
			}
		}
		else {
			if(oSpPr) {
				if(this.hasNoFillLine()) {
					if(oSpPr.ln) {
						oSpPr.setLn(null);
					}
				}
				if(oSpPr.Fill) {
					oSpPr.setFill(null);
				}
			}
		}
		if(nType === Asc.c_oAscChartTypeSettings.radarMarker) {
			if(this.marker) {
				this.setMarker(null);
			}
		}
		else {
			if(!this.marker) {
				this.setMarker(new CMarker());
			}
			let oMarker = this.marker;
			if(oMarker.symbol !== AscFormat.SYMBOL_NONE) {
				oMarker.setSymbol(AscFormat.SYMBOL_NONE);
			}
		}
	};
	CRadarSeries.prototype.checkSeriesAfterChangeType = function(nType) {
		if(!this.parent) {
			return;
		}
		const nRadarStyle = this.parent.radarStyle;
		if(nRadarStyle === AscFormat.RADAR_STYLE_FILLED) {
			if(this.spPr && this.spPr.hasNoFill()) {
				this.spPr.setFill(null);
			}
		}
		else {
			if(this.spPr && this.spPr.hasNoFillLine()) {
				this.spPr.setLn(null);
			}
		}
		const bMarker = (nType === Asc.c_oAscChartTypeSettings.radarMarker);
		if(!this.marker) {
			this.setMarker(new AscFormat.CMarker());
		}
		if (bMarker) {
			if(this.marker.symbol === null || this.marker.symbol === AscFormat.SYMBOL_NONE) {
				this.marker.setSymbol(AscFormat.GetTypeMarkerByIndex(this.idx));
			}
		} else {
			this.marker.setSymbol(AscFormat.SYMBOL_NONE);
		}
	};

	CRadarSeries.prototype.getPreviewBrush = function() {
		if(this.parent && this.parent.radarStyle === AscFormat.RADAR_STYLE_FILLED) {
			return CSeriesBase.prototype.getPreviewBrush.call(this);
		}
		return CLineSeries.prototype.getPreviewBrush.call(this);
	};

	var ORIENTATION_MAX_MIN = 0;
    var ORIENTATION_MIN_MAX = 1;

    function CScaling() {
        CBaseChartObject.call(this);
        this.logBase = null;
        this.max = null;
        this.min = null;
        this.orientation = null;
    }

    InitClass(CScaling, CBaseChartObject, AscDFH.historyitem_type_Scaling);
    CScaling.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.logBase !== null) {
            oCopy.setLogBase(this.logBase);
        }
        if(this.max !== null) {
            oCopy.setMax(this.max);
        }
        if(this.min !== null) {
            oCopy.setMin(this.min);
        }
        if(this.orientation !== null) {
            oCopy.setOrientation(this.orientation);
        }
    };
    CScaling.prototype.setLogBase = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_Scaling_SetLogBase, this.logBase, pr));
        this.logBase = pr;
        this.onChartInternalUpdate();
    };
    CScaling.prototype.setMax = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_Scaling_SetMax, this.max, pr));
        this.max = pr;
        this.onChartInternalUpdate();
    };
    CScaling.prototype.setMin = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_Scaling_SetMin, this.min, pr));
        this.min = pr;
        this.onChartInternalUpdate();
    };
    CScaling.prototype.setOrientation = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Scaling_SetOrientation, this.orientation, pr));
        this.orientation = pr;
        this.onChartInternalUpdate();
    };

    function CScatterChart() {
        CChartBase.call(this);
        this.scatterStyle = null;
    }

    InitClass(CScatterChart, CChartBase, AscDFH.historyitem_type_ScatterChart);
    CScatterChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.r;
    };
    CScatterChart.prototype.getEmptySeries = function() {
        return new CScatterSeries();
    };
    CScatterChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.scatterStyle !== null && oCopy.setScatterStyle) {
            oCopy.setScatterStyle(this.scatterStyle);
        }
    };
    CScatterChart.prototype.setScatterStyle = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ScatterChart_SetScatterStyle, this.scatterStyle, pr));
        this.scatterStyle = pr;
        this.onChartUpdateType();
    };
    CScatterChart.prototype.getLineParams = function() {
        var nFlag = 0;//
        var bShowMarker = false;
        var bNoLine = false;
        var bSmooth = false;
        var aSeries = this.series, nSer, oSeries;
        switch(this.scatterStyle) {
            case AscFormat.SCATTER_STYLE_LINE:
            {
                bNoLine = false;
                bSmooth = false;
                bShowMarker = false;
                break;
            }
            case AscFormat.SCATTER_STYLE_LINE_MARKER:
            {
                bNoLine = false;
                bSmooth = false;
                bShowMarker = true;
                break;
            }
            case AscFormat.SCATTER_STYLE_MARKER:
            {
                bNoLine = true;
                bShowMarker = false;
                for(nSer = 0; nSer < aSeries.length; ++nSer) {
                    oSeries = aSeries[nSer];
                    if(!(oSeries.marker && oSeries.marker.symbol === AscFormat.SYMBOL_NONE)) {
                        bShowMarker = true;
                        break;
                    }
                }
                break;
            }
            case AscFormat.SCATTER_STYLE_NONE:
            {
                bNoLine = true;
                bShowMarker = false;
                break;
            }
            case AscFormat.SCATTER_STYLE_SMOOTH:
            {
                bNoLine = false;
                bSmooth = true;
                bShowMarker = false;
                break;
            }
            case AscFormat.SCATTER_STYLE_SMOOTH_MARKER:
            {
                bNoLine = false;
                bSmooth = true;
                bShowMarker = true;
                break;
            }
        }
        if(!bNoLine) {
            for(nSer = 0; nSer < aSeries.length; ++nSer) {
                oSeries = aSeries[nSer];
                if(!oSeries.hasNoFillLine()) {
                    break;
                }
            }
            if(nSer === aSeries.length) {
                bNoLine = true;
            }
            if(bSmooth) {
                for(nSer = 0; nSer < aSeries.length; ++nSer) {
                    oSeries = aSeries[nSer];
                    if(!oSeries.smooth) {
                        bSmooth = false;
                        break;
                    }
                }
            }
        }
        if(bShowMarker) {
            nFlag |= 1;
        }
        if(bNoLine) {
            nFlag |= 2;
        }
        if(bSmooth) {
            nFlag |= 4;
        }
        return nFlag;
    };
    CScatterChart.prototype.isMarkerChart = function() {
        var bFlags = this.getLineParams();
        var bShowMarker = (bFlags & 1) === 1;
        return bShowMarker;
    };
    CScatterChart.prototype.isNoLine = function() {
        var bFlags = this.getLineParams();
        var bNoLine = (bFlags & 2) === 2;
        return bNoLine;
    };
    CScatterChart.prototype.isSmooth = function() {
        var bFlags = this.getLineParams();
        var bSmooth = (bFlags & 4) === 4;
        return bSmooth;
    };
    CScatterChart.prototype.getChartType = function() {
        var bIsMarker = this.isMarkerChart();
        var bIsLine = !this.isNoLine();
        var bSmooth = this.isSmooth();
        var nType = Asc.c_oAscChartTypeSettings.scatter;
        if(bIsLine && !bIsMarker && !bSmooth) {
            nType = Asc.c_oAscChartTypeSettings.scatterLine;
        }
        else if(bIsLine && bIsMarker && !bSmooth) {
            nType = Asc.c_oAscChartTypeSettings.scatterLineMarker;
        }
        else if(!bIsLine && bIsMarker && !bSmooth) {
            nType = Asc.c_oAscChartTypeSettings.scatter;
        }
        else if(bIsLine && !bIsMarker && bSmooth) {
            nType = Asc.c_oAscChartTypeSettings.scatterSmooth;
        }
        else if(bIsLine && bIsMarker && bSmooth) {
            nType = Asc.c_oAscChartTypeSettings.scatterSmoothMarker;
        }
        return nType;
    };
    CScatterChart.prototype.tryChangeType = function(nNewType) {
        if(!this.parent) {
            return false;
        }
        if(this.getChartType() === nNewType) {
            return true;
        }
        if(!this.parent.isScatterType(nNewType)) {
            return false;
        }
        var bMarker = getIsMarkerByType(nNewType);
        var bLine = getIsLineByType(nNewType);
        var bSmooth = getIsSmoothByType(nNewType);
        this.setLineParams(bMarker, bLine, bSmooth);
        return true;
    };

    CScatterChart.prototype.setLineParams = function(bMarker, bLine, bSmooth) {
        if(!AscFormat.isRealBool(bMarker) || !AscFormat.isRealBool(bLine) || !AscFormat.isRealBool(bSmooth)) {
            return;
        }
        if(bMarker === this.isMarkerChart() && bLine === !this.isNoLine() && bSmooth === this.isSmooth()) {
            return;
        }
        var nSer, oSeries;
        for(nSer = 0; nSer < this.series.length; ++nSer) {
            oSeries = this.series[nSer];
            if(oSeries.marker) {
                oSeries.setMarker(null);
            }
            if(AscFormat.isRealBool(oSeries.smooth)) {
                oSeries.setSmooth(null);
            }
        }
        var new_scatter_style;
        if(bLine) {
            for(nSer = 0; nSer < this.series.length; ++nSer) {
                oSeries = this.series[nSer];
                oSeries.removeAllDPts();
                if(oSeries.spPr && oSeries.spPr.ln) {
                    oSeries.spPr.setLn(null);
                }
            }
            if(bSmooth) {
                if(bMarker) {
                    new_scatter_style = AscFormat.SCATTER_STYLE_SMOOTH_MARKER;
                    for(nSer = 0; nSer < this.series.length; ++nSer) {
                        oSeries = this.series[nSer];
                        if(oSeries.marker) {
                            oSeries.setMarker(null);
                        }
                        if(oSeries.smooth !== true) {
                            oSeries.setSmooth(true);
                        }
                    }
                }
                else {
                    new_scatter_style = AscFormat.SCATTER_STYLE_SMOOTH;
                    for(nSer = 0; nSer < this.series.length; ++nSer) {
                        oSeries = this.series[nSer];
                        if(!oSeries.marker) {
                            oSeries.setMarker(new AscFormat.CMarker());
                        }
                        oSeries.marker.setSymbol(AscFormat.SYMBOL_NONE);
                        oSeries.setSmooth(true);
                    }
                }
            }
            else {
                if(bMarker) {
                    new_scatter_style = AscFormat.SCATTER_STYLE_LINE_MARKER;
                    for(nSer = 0; nSer < this.series.length; ++nSer) {
                        oSeries = this.series[nSer];
                        if(oSeries.marker) {
                            oSeries.setMarker(null);
                        }
                        oSeries.setSmooth(false);
                    }
                }
                else {
                    new_scatter_style = AscFormat.SCATTER_STYLE_LINE;
                    for(nSer = 0; nSer < this.series.length; ++nSer) {
                        oSeries = this.series[nSer];
                        if(!oSeries.marker) {
                            oSeries.setMarker(new AscFormat.CMarker());
                        }
                        oSeries.marker.setSymbol(AscFormat.SYMBOL_NONE);
                        oSeries.setSmooth(false);
                    }
                }
            }
        }
        else {
            for(nSer = 0; nSer < this.series.length; ++nSer) {
                oSeries = this.series[nSer];
                oSeries.removeAllDPts();
                if(!oSeries.spPr) {
                    oSeries.setSpPr(new AscFormat.CSpPr());
                }
                oSeries.spPr.setLn(AscFormat.CreateNoFillLine());
            }
            if(bMarker) {
                new_scatter_style = AscFormat.SCATTER_STYLE_MARKER;
                for(nSer = 0; nSer < this.series.length; ++nSer) {
                    oSeries = this.series[nSer];
                    if(oSeries.marker) {
                        oSeries.setMarker(null);
                    }
                    oSeries.setSmooth(false);
                }
            }
            else {
                new_scatter_style = AscFormat.SCATTER_STYLE_MARKER;
                for(nSer = 0; nSer < this.series.length; ++nSer) {
                    oSeries = this.series[nSer];
                    if(!oSeries.marker) {
                        oSeries.setMarker(new AscFormat.CMarker());
                    }
                    oSeries.marker.setSymbol(AscFormat.SYMBOL_NONE);
                }
            }
        }
        this.setScatterStyle(new_scatter_style);
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            var oChartStyle = oChartSpace.chartStyle;
            var oChartColors = oChartSpace.chartColors;
            if(oChartStyle && oChartColors) {
                this.applyChartStyle(oChartStyle, oChartColors, null, false);
            }
        }
    };

    function CScatterSeries() {
        CSeriesBase.call(this);
        this.dLbls = null;
        this.dPt = [];
        this.errBars = null;
        this.marker = null;
        this.smooth = null;
        this.trendline = null;
        this.xVal = null;
        this.yVal = null;
    }

    InitClass(CScatterSeries, CSeriesBase, AscDFH.historyitem_type_ScatterSer);
    CScatterSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.errBars && oCopy.setErrBars) {
            oCopy.setErrBars(this.errBars.createDuplicate());
        }
        if(this.marker && oCopy.setMarker) {
            oCopy.setMarker(this.marker.createDuplicate());
        }
        if(this.smooth !== null && oCopy.setSmooth) {
            oCopy.setSmooth(this.smooth);
        }
        if(this.trendline && oCopy.setTrendline) {
            oCopy.setTrendline(this.trendline.createDuplicate());
        }
    };
    CScatterSeries.prototype.getChildren = function() {
        var aRet = CSeriesBase.prototype.getChildren.call(this);
        aRet.push(this.dLbls);
        for(var i = 0; i < this.dPt.length; ++i) {
            aRet.push(this.dPt[i]);
        }
        aRet.push(this.errBars);
        aRet.push(this.marker);
        aRet.push(this.trendline);
        return aRet;
    };
    CScatterSeries.prototype.setDLbls = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetDLbls, this.dLbls, pr));
        this.dLbls = pr;
        this.setParentToChild(pr);
    };
    CScatterSeries.prototype.addDPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_ScatterSer_SetDPt, this.dPt.length, [pr], true));
        this.dPt.push(pr);
        this.setParentToChild(pr);
    };
    CScatterSeries.prototype.setErrBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetErrBars, this.errBars, pr));
        this.errBars = pr;
        this.setParentToChild(pr);
    };
    CScatterSeries.prototype.setMarker = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetMarker, this.marker, pr));
        this.marker = pr;
        this.setParentToChild(pr);
    };
    CScatterSeries.prototype.setSmooth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ScatterSer_SetSmooth, this.smooth, pr));
        this.smooth = pr;
    };
    CScatterSeries.prototype.setTrendline = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CScatterSeries.prototype.setXVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetXVal, this.xVal, pr));
        this.xVal = pr;
        this.setParentToChild(pr);
        this.onChangeDataRefs();
    };
    CScatterSeries.prototype.setYVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ScatterSer_SetYVal, this.yVal, pr));
        this.yVal = pr;
        this.setParentToChild(pr);
        this.onChangeDataRefs();
    };
    CScatterSeries.prototype.setCat = function(pr) {
        this.setXVal(pr);
    };
    CScatterSeries.prototype.setVal = function(pr) {
        this.setYVal(pr);
    };
    CScatterSeries.prototype.checkSeriesAfterChangeType = function(pr) {
	    this.setMarker(null);
    };

    function CTx() {
        CBaseChartObject.call(this);
        this.strRef = null;
        this.val = null;
    }

    InitClass(CTx, CBaseChartObject, AscDFH.historyitem_type_Tx);
    CTx.prototype.getChildren = function() {
        return [this.strRef];
    };
    CTx.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.strRef) {
            oCopy.setStrRef(this.strRef.createDuplicate());
        }
        if(this.val !== null) {
            oCopy.setVal(this.val);
        }
    };
    CTx.prototype.setStrRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Tx_SetStrRef, this.strRef, pr));
        this.strRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CTx.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Tx_SetVal, this.strRef, pr));
        this.val = pr;
    };
    CTx.prototype.getText = function(bNoUpdate) {
        var sRet = "";
        if(typeof this.val === "string" && this.val.length > 0) {
            sRet = this.val;
        }
        else if(this.strRef) {
            sRet = this.strRef.getText(bNoUpdate);
        }
        return sRet;
    };
    CTx.prototype.getFormula = function() {
        var sRet = "";
        if(typeof this.val === "string" && this.val.length > 0) {
            sRet = "=\"" + this.val + "\"";
        }
        else if(this.strRef) {
            sRet = this.strRef.getFormula();
        }
        return sRet;
    };
    CTx.prototype.getDataRefs = function() {
        if(this.strRef) {
            return this.strRef.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CTx.prototype.collectRefs = function(aRefs) {
        if(this.strRef) {
            aRefs.push(this.strRef);
        }
    };
    CTx.prototype.setValues = function(sName) {
        var oResult = new CParseResult();
        var oParser, bResult, oLastElem;
        if(typeof sName === "string" && sName.length > 0) {
            fParseStrRef(sName, false, oResult);
            if(oResult.isSuccessful()) {
                this.setStrRef(oResult.getObject());
                this.setVal(null);
            }
            else {
                var aParsed = AscFormat.fParseChartFormula(sName);
                if(aParsed.length > 0) {
                    return oResult;
                }
                if(sName[0] === "=") {
                    oParser = new AscCommonExcel.parserFormula(sName.slice(1), null, Asc.editor.wbModel.aWorksheets[0]);
                    bResult = oParser.parse(true, true);
                    if(bResult && oParser.outStack.length === 1 &&
                        (oLastElem = oParser.outStack[0]) &&
                        oLastElem.type === AscCommonExcel.cElementType.string) {
                        if(this.strRef) {
                            this.setStrRef(null);
                        }
                        this.setVal(oLastElem.getValue());
                        oResult.setError(Asc.c_oAscError.ID.No);
                        oResult.setObject(oLastElem.getValue());
                    }
                    else {
                        oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
                    }
                }
                else {
                    if(this.strRef) {
                        this.setStrRef(null);
                    }
                    this.setVal(sName);
                    oResult.setError(Asc.c_oAscError.ID.No);
                    oResult.setObject(sName);
                }
            }
        }
        else {
            oResult.setError(Asc.c_oAscError.ID.DataRangeError);
        }
        return oResult;
    };
    CTx.prototype.isValid = function() {
        if(this.strRef || (typeof this.val === "string")) {
            return true;
        }
        return false;
    };
    CTx.prototype.update = function() {
        if(this.strRef) {
            this.strRef.updateCache();
        }
    };
    CTx.prototype.clearDataCache = function() {
        if(this.strRef) {
            this.strRef.clearDataCache();
        }
    };
    CTx.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.strRef) {
            this.strRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CTx.prototype.fillFromAsc = function(oTxCache, bUseCache) {
        this.setStrRef(new AscFormat.CStrRef());
        this.strRef.fillFromAsc(oTxCache, bUseCache);
    };

    function CStockChart() {
        CChartBase.call(this);
        this.dropLines = null;
        this.hiLowLines = null;
        this.upDownBars = null;
    }

    InitClass(CStockChart, CChartBase, AscDFH.historyitem_type_StockChart);
    CStockChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.r;
    };
    CStockChart.prototype.Refresh_RecalcData = function() {
        this.onChartUpdateType();
    };
    CStockChart.prototype.getEmptySeries = function() {
        return new CLineSeries();
    };
    CStockChart.prototype.getChildren = function() {
        var aRet = CChartBase.prototype.getChildren.call(this);
        aRet.push(this.dLbls);
        aRet.push(this.dropLines);
        aRet.push(this.hiLowLines);
        aRet.push(this.upDownBars);
        return aRet;
    };
    CStockChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.dLbls && oCopy.setDLbls) {
            oCopy.setDLbls(this.dLbls.createDuplicate());
        }
        if(this.dropLines && oCopy.setDropLines) {
            oCopy.setDropLines(this.dropLines.createDuplicate());
        }
        if(this.hiLowLines && oCopy.setHiLowLines) {
            oCopy.setHiLowLines(this.hiLowLines.createDuplicate());
        }
        if(this.upDownBars && oCopy.setUpDownBars) {
            oCopy.setUpDownBars(this.upDownBars.createDuplicate());
        }
    };
    CStockChart.prototype.setDropLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_StockChart_SetDropLines, this.dropLines, pr));
        this.dropLines = pr;
        this.setParentToChild(pr);
    };
    CStockChart.prototype.setHiLowLines = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_StockChart_SetHiLowLines, this.hiLowLines, pr));
        this.hiLowLines = pr;
        this.setParentToChild(pr);
    };
    CStockChart.prototype.setUpDownBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_StockChart_SetUpDownBars, this.upDownBars, pr));
        this.upDownBars = pr;
        this.setParentToChild(pr);
    };
    CStockChart.prototype.getChartType = function() {
        return Asc.c_oAscChartTypeSettings.stock;
    };

    CStockChart.prototype.isMarkerChart = function() {
        return false;
    };
    CStockChart.prototype.isNoLine = function() {
        return false;
    };
    CStockChart.prototype.isSmooth = function() {
        return false;
    };

    function CStrCache() {
        CBaseChartObject.call(this);
        this.pts = [];
        this.ptCount = null;
    }

    InitClass(CStrCache, CBaseChartObject, AscDFH.historyitem_type_StrCache);
    CStrCache.prototype.removeDPt = function(idx) {
        if(this.pts[idx]) {
            this.pts[idx].setParent(null);
            History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_CommonLit_RemoveDPt, idx, [this.pts[idx]], false));
            this.pts.splice(idx, 1);
        }
    };
    CStrCache.prototype.removeAllPts = function() {
        var start_idx = this.pts.length - 1;
        for(var i = start_idx; i > -1; --i) {
            this.removeDPt(i);
        }
        this.setPtCount(0);
    };
    //CStrCache.prototype.getChildren = function() {
    //    return this.pts;
    //};
    CStrCache.prototype.fillObject = function(oCopy, oIdMap) {
        for(var i = 0; i < this.pts.length; ++i) {
            oCopy.addPt(this.pts[i].createDuplicate());
        }
        if(this.ptCount !== null) {
            oCopy.setPtCount(this.ptCount);
        }
    };
    CStrCache.prototype.getPtByIndex = function(idx) {
        if(this.pts[idx] && this.pts[idx].idx === idx)
            return this.pts[idx];
        for(var i = 0; i < this.pts.length; ++i) {
            if(this.pts[i].idx === idx)
                return this.pts[i];
        }
        return null;
    };
    CStrCache.prototype.getPtCount = function() {
        if(AscFormat.isRealNumber(this.ptCount)) {
            return this.ptCount;
        }
        return this.pts.length;
    };
    CStrCache.prototype.addPt = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_StrCache_AddPt, this.pts.length, [pr], true));
        this.pts.push(pr);
        this.setParentToChild(pr);
    };
    CStrCache.prototype.setPtCount = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_StrCache_SetPtCount, this.ptCount, pr));
        this.ptCount = pr;
    };
    CStrCache.prototype.update = function(sFormula) {
        AscFormat.ExecuteNoHistory(function() {
            var str_cache = this;
            if(!(typeof sFormula === "string" && sFormula.length > 0)) {
                str_cache.removeAllPts();
                return;
            }
            var pt_index = 0, i, j, cell, pt, value_width_format, row_hidden, col_hidden, nPtCount = 0;
            var aParsedRef = AscFormat.fParseChartFormula(sFormula);

            if(!Array.isArray(aParsedRef) || aParsedRef.length === 0) {
                return;
            }
            str_cache.removeAllPts();
            var fParseTableDataString = function(oRef, oCache) {
                if(Array.isArray(oRef)) {
                    for(var i = 0; i < oRef.length; ++i) {
                        if(Array.isArray(oRef[i])) {
                            fParseTableDataString(oRef, oCache);
                        }
                        else {
                            pt = new AscFormat.CStringPoint();
                            pt.setIdx(pt_index);
                            pt.setVal(oRef[i].value);
                            str_cache.addPt(pt);
                            ++pt_index;
                            ++nPtCount;
                        }
                    }
                }
            };
            for(i = 0; i < aParsedRef.length; ++i) {
                var oCurRef = aParsedRef[i];
                var source_worksheet = oCurRef.worksheet;
                if(source_worksheet) {
                    var range = oCurRef.bbox;
                    if(range.r1 === range.r2) {
                        row_hidden = source_worksheet.getRowHidden(range.r1);
                        j = range.c1;
                        while(i === 0 && source_worksheet.getColHidden(j) && j <= range.c2) {
                            ++j;
                        }
                        for(; j <= range.c2; ++j) {
                            if(!row_hidden && !source_worksheet.getColHidden(j)) {
                                cell = source_worksheet.getCell3(range.r1, j);
                                value_width_format = cell.getValueWithFormat();
                                if(typeof value_width_format === "string" && value_width_format.length > 0) {
                                    pt = new AscFormat.CStringPoint();
                                    pt.setIdx(nPtCount);
                                    pt.setVal(value_width_format);

                                    if(str_cache.pts.length === 0) {
                                        pt.formatCode = cell.getNumFormatStr()
                                    }
                                    str_cache.addPt(pt);
                                    //addPointToMap(oThis.pointsMap, source_worksheet, range.r1, j, pt);
                                }
                                ++nPtCount;
                            }
                            pt_index++;
                        }
                    }
                    else {
                        col_hidden = source_worksheet.getColHidden(range.c1);
                        j = range.r1;
                        while(i === 0 && source_worksheet.getRowHidden(j) && j <= range.r2) {
                            ++j;
                        }
                        for(; j <= range.r2; ++j) {
                            if(!col_hidden && !source_worksheet.getRowHidden(j)) {
                                cell = source_worksheet.getCell3(j, range.c1);
                                value_width_format = cell.getValueWithFormat();
                                if(typeof value_width_format === "string" && value_width_format.length > 0) {
                                    pt = new AscFormat.CStringPoint();
                                    pt.setIdx(nPtCount);
                                    pt.setVal(cell.getValueWithFormat());

                                    if(str_cache.pts.length === 0) {
                                        pt.formatCode = cell.getNumFormatStr()
                                    }
                                    str_cache.addPt(pt);
                                    //addPointToMap(oThis.pointsMap, source_worksheet, j, range.c1,  pt);
                                }
                                ++nPtCount;
                            }
                            pt_index++;
                        }
                    }
                }
                else {
                    fParseTableDataString(oCurRef);
                }
            }
            this.setPtCount(nPtCount);
        }, this, []);
    };
    CStrCache.prototype.addStringPoint = function(idx, sStr) {
        var oPt = new CStringPoint();
        oPt.setIdx(idx);
        oPt.setVal(sStr);
        this.addPt(oPt);
        this.setPtCount(Math.max(this.pts.length, idx + 1));
    };
    CStrCache.prototype.getValues = function(nMaxCount, bAddEmpty) {
        var ret = [];
        var nEnd = nMaxCount || this.ptCount;
        for(var nIndex = 0; nIndex < nEnd; ++nIndex) {
            var oPt = this.getPtByIndex(nIndex);
            if(oPt) {
                ret.push(oPt.val + "");
            }
            else {
                if(bAddEmpty) {
                    ret.push("");
                }
            }
        }
        return ret;
    };
    CStrCache.prototype.getFormula = function() {
        var sRet = "={";
        for(var nIndex = 0; nIndex < this.pts.length; ++nIndex) {
            sRet += ("\"" + this.pts[nIndex].val + "\"");
            if(nIndex < this.pts.length - 1) {
                sRet += ", ";
            }
        }
        sRet += "}";
        return sRet;
    };

    function CStringPoint() {
        CBaseChartObject.call(this);
        this.idx = null;
        this.val = null;
    }

    InitClass(CStringPoint, CBaseChartObject, AscDFH.historyitem_type_StrPoint);
    CStringPoint.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.idx !== null) {
            oCopy.setIdx(this.idx);
        }
        if(this.val !== null) {
            oCopy.setVal(this.val);
        }
    };
    CStringPoint.prototype.setIdx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_StrPoint_SetIdx, this.idx, pr));
        this.idx = pr;
    };
    CStringPoint.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_StrPoint_SetVal, this.val, pr));
        this.val = pr;
        if(AscFonts.IsCheckSymbols) {
            if(typeof pr === "string" && pr.length > 0) {
                AscFonts.FontPickerByCharacter.getFontsByString(pr);
            }
        }
    };

    function CSurfaceChart() {
        CChartBase.call(this);
        this.bandFmts = [];
        this.wireframe = null;
        this.compiledBandFormats = [];
    }

    InitClass(CSurfaceChart, CChartBase, AscDFH.historyitem_type_SurfaceChart);
    CSurfaceChart.prototype.Refresh_RecalcData = function(data) {
        if(!isRealObject(data))
            return;
        switch(data.Type) {
            case AscDFH.historyitem_CommonChartFormat_SetParent:
            {
                break;
            }
            case AscDFH.historyitem_SurfaceChart_AddAxId:
            case AscDFH.historyitem_CommonChart_AddAxId:
            {
                break;
            }
            case AscDFH.historyitem_SurfaceChart_AddBandFmt:
            {
                break;
            }
            case AscDFH.historyitem_CommonChart_AddSeries:
            case AscDFH.historyitem_CommonChart_AddFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveFilteredSeries:
            case AscDFH.historyitem_CommonChart_RemoveSeries:
            {
                this.onChartUpdateType();
                this.onChangeDataRefs();
                break;
            }
            case AscDFH.historyitem_SurfaceChart_SetWireframe:
            {
                break;
            }
        }
    };
    CSurfaceChart.prototype.isWireframe = function() {
        return this.wireframe === true;
    };
    CSurfaceChart.prototype.getBandFmtByIndex = function(idx) {
        for(var i = 0; i < this.bandFmts.length; ++i) {
            if(this.bandFmts[i].idx === idx) {
                return this.bandFmts[i];
            }
        }
        return null;
    };
    CSurfaceChart.prototype.getEmptySeries = function() {
        return new CSurfaceSeries();
    };
    CSurfaceChart.prototype.getDefaultDataLabelsPosition = function() {
        return c_oAscChartDataLabelsPos.ctr;
    };
    CSurfaceChart.prototype.getChildren = function() {
        var aRet = CChartBase.prototype.getChildren.call(this);
        for(var i = 0; i < this.bandFmts.length; ++i) {
            aRet.push(this.bandFmts[i]);
        }
        return aRet;
    };
    CSurfaceChart.prototype.fillObject = function(oCopy, oIdMap) {
        CChartBase.prototype.fillObject.call(this, oCopy, oIdMap);
        var i;
        if(oCopy.addBandFmt) {
            for(i = 0; i < this.bandFmts.length; ++i) {
                oCopy.addBandFmt(this.bandFmts[i].createDuplicate());
            }
        }
        if(this.wireframe !== null && oCopy.setWireframe) {
            oCopy.setWireframe(this.wireframe);
        }
    };
    CSurfaceChart.prototype.getAllRasterImages = function(images) {
        CChartBase.prototype.getAllRasterImages.call(this, images);
        for(var i = 0; i < this.bandFmts.length; ++i) {
            this.bandFmts[i] && this.bandFmts[i].spPr && this.bandFmts[i].spPr.checkBlipFillRasterImage(images);
        }
    };
    CSurfaceChart.prototype.checkSpPrRasterImages = function(images) {
        CChartBase.prototype.checkSpPrRasterImages.call(this, images);
        for(var i = 0; i < this.bandFmts.length; ++i) {
            this.bandFmts[i] && checkSpPrRasterImages(this.bandFmts[i].spPr);
        }
    };
    CSurfaceChart.prototype.addBandFmt = function(fmt) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_SurfaceChart_AddBandFmt, this.bandFmts.length, [fmt], true));
        this.bandFmts.push(fmt);
        this.setParentToChild(fmt);
    };
    CSurfaceChart.prototype.setWireframe = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_SurfaceChart_SetWireframe, this.wireframe, pr));
        this.wireframe = pr;
    };
    CSurfaceChart.prototype.getChartType = function() {
        var oCS = this.getChartSpace();
        var oView3D = oCS && oCS.chart.view3D;
        var nType = Asc.c_oAscChartTypeSettings.unknown;
        if(this.isWireframe()) {
            if(oView3D && oView3D.rotX === 90 && oView3D.rotY === 0 && oView3D.rAngAx === false && oView3D.perspective === 0) {
                nType = Asc.c_oAscChartTypeSettings.contourWireframe;
            }
            else {
                nType = Asc.c_oAscChartTypeSettings.surfaceWireframe;
            }
        }
        else {
            if(oView3D && oView3D.rotX === 90 && oView3D.rotY === 0 && oView3D.rAngAx === false && oView3D.perspective === 0) {
                nType = Asc.c_oAscChartTypeSettings.contourNormal;
            }
            else {
                nType = Asc.c_oAscChartTypeSettings.surfaceNormal;
            }
        }
        return nType;
    };
    CSurfaceChart.prototype.setDlblsProps = function(oProps) {
    };

    function CSurfaceSeries() {
        CSeriesBase.call(this);
        this.cat = null;
        this.val = null;
    }

    InitClass(CSurfaceSeries, CSeriesBase, AscDFH.historyitem_type_SurfaceSeries);
    CSurfaceSeries.prototype.fillObject = function(oCopy, oIdMap) {
        CSeriesBase.prototype.fillObject.call(this, oCopy, oIdMap);
    };
    CSurfaceSeries.prototype.setCat = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_SurfaceSeries_SetCat, this.cat, pr));
        this.cat = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CSurfaceSeries.prototype.setVal = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_SurfaceSeries_SetVal, this.val, pr));
        this.val = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CSurfaceSeries.prototype.isWireframe = function(oChartStyle, oColors) {
        if(!this.parent) {
            return false;
        }
        return this.parent.isWireframe();
    };
    CSurfaceSeries.prototype.setDlblsProps = function(oProps) {
    };

    function CTitle() {
        CBaseChartObject.call(this);
        this.layout = null;
        this.overlay = false;
        this.spPr = null;
        this.tx = null;
        this.txPr = null;

        this.txBody = null;
        this.x = null;
        this.y = null;
        this.calcX = null;
        this.calcY = null;
        this.extX = null;
        this.extY = null;

        this.transform = new CMatrix();
        this.transformText = new CMatrix();
        this.ownTransform = new CMatrix();
        this.ownTransformText = new CMatrix();
        this.localTransform = new CMatrix();
        this.localTransformText = new CMatrix();

        this.recalcInfo =
        {
            recalculateTxBody: true,
            recalcTransform: true,
            recalculateTransformText: true,
            recalculateContent: true,
            recalculateBrush: true,
            recalculatePen: true,
            recalcStyle: true,
            recalculateGeometry: true
        };
    }

    InitClass(CTitle, CBaseChartObject, AscDFH.historyitem_type_Title);
    CTitle.prototype.hitInPath = CShape.prototype.hitInPath;
    CTitle.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;
    CTitle.prototype.hitInBoundingRect = CShape.prototype.hitInBoundingRect;
    CTitle.prototype.hitInTextRect = CDLbl.prototype.hitInTextRect;
    CTitle.prototype.recalculateGeometry = CShape.prototype.recalculateGeometry;
    CTitle.prototype.getTransform = CShape.prototype.getTransform;
    CTitle.prototype.check_bounds = CShape.prototype.check_bounds;
    CTitle.prototype.selectionCheck = CShape.prototype.selectionCheck;
    CTitle.prototype.getInvertTransform = CShape.prototype.getInvertTransform;
    CTitle.prototype.recalculatePen = CShape.prototype.recalculatePen;
    CTitle.prototype.recalculateBrush = CShape.prototype.recalculateBrush;
    CTitle.prototype.updateSelectionState = CShape.prototype.updateSelectionState;
    CTitle.prototype.selectionSetStart = CShape.prototype.selectionSetStart;
    CTitle.prototype.selectionSetEnd = CShape.prototype.selectionSetEnd;
    CTitle.prototype.getStyles = CDLbl.prototype.getStyles;
    CTitle.prototype.Get_Theme = CDLbl.prototype.Get_Theme;
    CTitle.prototype.Get_ColorMap = CDLbl.prototype.Get_ColorMap;
    CTitle.prototype.recalculateStyle = CDLbl.prototype.recalculateStyle;
    CTitle.prototype.recalculateTxBody = CDLbl.prototype.recalculateTxBody;
    CTitle.prototype.recalculateTransform = CDLbl.prototype.recalculateTransform;
    CTitle.prototype.recalculateTransformText = CDLbl.prototype.recalculateTransformText;
    CTitle.prototype.recalculateContent = CDLbl.prototype.recalculateContent;
    CTitle.prototype.setPosition = CDLbl.prototype.setPosition;
    CTitle.prototype.checkHitToBounds = CDLbl.prototype.checkHitToBounds;
    CTitle.prototype.getBodyPr = CDLbl.prototype.getBodyPr;
    CTitle.prototype.getCompiledStyle = CDLbl.prototype.getCompiledStyle;
    CTitle.prototype.getCompiledFill = CDLbl.prototype.getCompiledFill;
    CTitle.prototype.getCompiledLine = CDLbl.prototype.getCompiledLine;
    CTitle.prototype.getCompiledTransparent = CDLbl.prototype.getCompiledTransparent;
    CTitle.prototype.Get_Styles = CDLbl.prototype.Get_Styles;
    CTitle.prototype.getCanvasContext = CDLbl.prototype.getCanvasContext;
    CTitle.prototype.convertPixToMM = CDLbl.prototype.convertPixToMM;
    CTitle.prototype.isEmptyPlaceholder = CDLbl.prototype.isEmptyPlaceholder;
    CTitle.prototype.checkShapeChildTransform = CDLbl.prototype.checkShapeChildTransform;
    CTitle.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
    CTitle.prototype.chekBodyPrTransform = function() {
        return false;
    };
    CTitle.prototype.checkContentWordArt = function() {
        return false;
    };
    CTitle.prototype.Get_AbsolutePage = function() {
        if(this.chart && this.chart.Get_AbsolutePage) {
            return this.chart.Get_AbsolutePage();
        }
        return 0;
    };
    CTitle.prototype.Refresh_RecalcData2 = function(pageIndex) {
        this.recalcInfo.recalculateTxBody = true;
        this.recalcInfo.recalcTransform = true;
        this.recalcInfo.recalculateTransformText = true;
        this.recalcInfo.recalculateContent = true;
        this.recalcInfo.recalculateGeometry = true;

        this.parent && this.parent.Refresh_RecalcData2 && this.parent.Refresh_RecalcData2(pageIndex, this);
    };
    CTitle.prototype.checkAfterChangeTheme = function() {
        this.resetRecalcFlags();
    };
    CTitle.prototype.GetRevisionsChangeElement = function(SearchEngine) {
        var oContent = this.getDocContent();
        if(oContent) {
            oContent.GetRevisionsChangeElement(SearchEngine);
        }
    };
    CTitle.prototype.handleUpdateFill = function() {

        this.recalcInfo.recalculateBrush = true;
        this.Refresh_RecalcData();
    };
    CTitle.prototype.handleUpdateLn = function() {
        this.recalcInfo.recalculatePen = true;
        this.Refresh_RecalcData();
    };
    CTitle.prototype.Search = function(SearchEngine, Type) {
        var content = this.getDocContent();
        if(content && this.tx && this.tx.rich) {
            var dd = this.getDrawingDocument();
            dd.StartSearchTransform && dd.StartSearchTransform(this.transformText);
            content.Search(SearchEngine, Type);
            dd.EndSearchTransform && dd.EndSearchTransform();
        }
    };
    CTitle.prototype.GetSearchElementId = function(bNext, bCurrent) {
        var content = this.getDocContent();
        if(content && this.tx && this.tx.rich) {
            return content.GetSearchElementId(bNext, bCurrent);
        }
        return null;
    };
    CTitle.prototype.FindNextFillingForm = function(isNext, isCurrent) {
		return null;
	};
    CTitle.prototype.Set_CurrentElement = function(bUpdate, pageIndex) {

        var chart = this.chart;
        if(chart && typeof editor !== "undefined" && editor && editor.WordControl && editor.WordControl.m_oLogicDocument) {
            var bDocument = false, drawing_objects;
            if(editor.WordControl.m_oLogicDocument instanceof CDocument) {
                bDocument = true;
                drawing_objects = editor.WordControl.m_oLogicDocument.DrawingObjects;
            }
            else if(editor.WordControl.m_oLogicDocument instanceof CPresentation) {
                if(chart.parent) {
                    drawing_objects = chart.parent.graphicObject;
                }
            }
            if(drawing_objects) {
                drawing_objects.resetSelection(true);
                var para_drawing;
                if(chart.group) {
                    var main_group = chart.group.getMainGroup();
                    drawing_objects.selectObject(main_group, pageIndex);
                    main_group.selectObject(chart, pageIndex);
                    main_group.selection.chartSelection = chart;

                    chart.selection.textSelection = this;
                    chart.selection.title = this;
                    drawing_objects.selection.groupSelection = main_group;
                    para_drawing = main_group.parent;
                }
                else {
                    drawing_objects.selectObject(chart, pageIndex);
                    drawing_objects.selection.chartSelection = chart;
                    chart.selection.textSelection = this;
                    chart.selection.title = this;
                    para_drawing = chart.parent;
                }
                if(bDocument && para_drawing instanceof AscCommonWord.ParaDrawing) {

                    var hdr_ftr = para_drawing.DocumentContent.IsHdrFtr(true);
                    if(hdr_ftr) {
                        hdr_ftr.Content.SetDocPosType(docpostype_DrawingObjects);
                        hdr_ftr.Set_CurrentElement(bUpdate);
                    }
                    else {
                        drawing_objects.document.SetDocPosType(docpostype_DrawingObjects);
                    }
                }
            }

        }
    };
    CTitle.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.layout);
        aRet.push(this.spPr);
        aRet.push(this.tx);
        aRet.push(this.txPr);
        return aRet;
    };
    CTitle.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.layout) {
            oCopy.setLayout(this.layout.createDuplicate());
        }
        oCopy.setOverlay(this.overlay);
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(this.tx) {
            oCopy.setTx(this.tx.createDuplicate());
        }
        if(this.txPr) {
            oCopy.setTxPr(this.txPr.createDuplicate());
        }
    };
    CTitle.prototype.paragraphAdd = function(paraItem, bRecalculate) {
        var content = this.getDocContent();
        if(content) {
            content.AddToParagraph(paraItem, bRecalculate);
        }
    };
    CTitle.prototype.applyTextFunction = function(docContentFunction, tableFunction, args) {
        var content = this.getDocContent();
        if(content) {
            docContentFunction.apply(content, args);
        }
    };
    CTitle.prototype.checkDocContent = function() {
        if(this.tx && this.tx.rich && this.tx.rich.content && !this.showDlblsRange) {
            return;
        }
        else if(this.txBody && this.txBody.content) {
            var StartPage = this.txBody.content.StartPage;
            if(!this.tx) {
                this.setTx(new CChartText());
            }
            this.tx.setRich(this.txBody.createDuplicate2());
            this.tx.rich.setParent(this);
            if(this.txPr) {
                if(this.txPr.content && this.txPr.content.Content[0] && this.txPr.content.Content[0].Pr.DefaultRunPr) {
                    for(var i = 0; i < this.tx.rich.content.Content.length; ++i) {
                        AscFormat.CheckParagraphTextPr(this.tx.rich.content.Content[i], this.txPr.content.Content[0].Pr.DefaultRunPr);
                    }
                }
                if(this.txPr.bodyPr) {
                    this.tx.rich.setBodyPr(this.txPr.bodyPr.createDuplicate());
                }
            }
            var selection_state = this.txBody.content.GetSelectionState();
            this.txBody = this.tx.rich;
            this.txBody.content.SetSelectionState(selection_state, selection_state.length - 1);
            if(AscFormat.isRealNumber(StartPage)) {
                this.txBody.content.Set_StartPage(StartPage);
            }
			if(this.showDlblsRange && this.setShowDlblsRange) {
				this.setShowDlblsRange(false);
			}
            //if(editor && editor.isDocumentEditor)
            //{
            //    this.recalculateContent();
            //}
        }
    };
    CTitle.prototype.getDocContent = function() {
        if(this.recalcInfo.recalculateTxBody) {
            AscFormat.ExecuteNoHistory(this.recalculateTxBody, this, []);
            this.recalcInfo.recalculateTxBody = false;
        }
        if(this.txBody && this.txBody.content) {
            return this.txBody.content;
        }
    };
    CTitle.prototype.select = function(chartSpace, pageIndex) {
        this.selected = true;
        this.selectStartPage = pageIndex;
        var content = this.getDocContent && this.getDocContent();
        if(content)
            content.Set_StartPage(pageIndex);
        chartSpace.selection.title = this;
    };
    CTitle.prototype.getMaxWidth = function(bodyPr) {
        switch(bodyPr.vert) {
            case AscFormat.nVertTTeaVert:
            case AscFormat.nVertTTmongolianVert:
            case AscFormat.nVertTTvert:
            case AscFormat.nVertTTwordArtVert:
            case AscFormat.nVertTTwordArtVertRtl:
            case AscFormat.nVertTTvert270:
            {
                var vert_axis = this.chart.chart.plotArea.getVerticalAxis();
                if(vert_axis && vert_axis.title === this) {
                    var hor_axis = this.chart.chart.plotArea.getHorizontalAxis();
                    return this.chart.extY - (hor_axis && hor_axis.title ? hor_axis.title.extY : 0 )
                }
                return this.chart.extY / 2;
            }
            case AscFormat.nVertTThorz:
            {
                return this.chart.extX * 0.8;
            }
        }
        return this.chart.extX * 0.5;
    };
    CTitle.prototype.IsUseInDocument = function() {
        if(this.parent && this.parent.title === this &&
            this.chart && this.chart.IsUseInDocument) {
            return this.chart.IsUseInDocument();
        }
        return false;
    };
    CTitle.prototype.Check_AutoFit = function() {
        return true;
    };
    CTitle.prototype.getDrawingDocument = function() {
        if(this.chart && this.chart.getDrawingDocument)
            return this.chart && this.chart.getDrawingDocument();
        return this.parent && this.parent.getDrawingDocument && this.parent.getDrawingDocument();
    };
    CTitle.prototype.draw = function(graphics) {
        CDLbl.prototype.draw.call(this, graphics);
    };
    CTitle.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, x, y);

        this.transformText = this.localTransformText.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transformText, x, y);

        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);

    };
    CTitle.prototype.getParentObjects = function() {
        if(this.chart) {
            return this.chart.getParentObjects();
        }
        else if(this.parent) {
            return this.parent.getParentObjects();
        }
        return null;
    };
    CTitle.prototype.resetRecalcFlags = function() {
        this.recalcInfo.recalculateTxBody = true;
        this.recalcInfo.recalcTransform = true;
        this.recalcInfo.recalculateTransformText = true;
        this.recalcInfo.recalculateContent = true;
        this.recalcInfo.recalculateGeometry = true;
        if(this.tx && this.tx.rich && this.tx.rich.content) {
            this.tx.rich.content.Recalc_AllParagraphs_CompiledPr();
        }
    };
    CTitle.prototype.onUpdate = function() {
       this.resetRecalcFlags();
        var oChartSpace = this.getChartSpace();
        if(oChartSpace) {
            oChartSpace.recalcInfo.recalculateAxisLabels = true;
        }
    };
    CTitle.prototype.getDefaultTextForTxBody = function() {
        var sText;
        if(this.tx && this.tx.strRef) {
            sText = this.tx.strRef.getText(false);
            if(typeof sText === "string" && sText.length > 0) {
                return sText;
            }
        }
        var key = 'Axis Title';
        if(this.parent) {
            if(this.parent.getObjectType() === AscDFH.historyitem_type_Chart) {
                if(this.parent.plotArea && this.parent.plotArea.charts.length === 1 && Array.isArray(this.parent.plotArea.charts[0].series)
                    && this.parent.plotArea.charts[0].series.length === 1 && this.parent.plotArea.charts[0].series[0].tx) {
                    var oTx = this.parent.plotArea.charts[0].series[0].tx;
                    sText = oTx.getText(false);
                    if(typeof sText === "string" && sText.length > 0) {
                        return sText;
                    }
                }
                key = 'Diagram Title';
            }
            else {
                if(this.parent.axPos === AX_POS_B || this.parent.axPos === AX_POS_T) {
                    key = 'X Axis';
                }
                else {
                    key = 'Y Axis';
                }
            }
        }
        return AscCommon.translateManager.getValue(key);
    };
    CTitle.prototype.recalculate = function() {
        AscFormat.ExecuteNoHistory(function() {
            if(this.recalcInfo.recalculateBrush) {
                this.recalculateBrush();
                this.recalcInfo.recalculateBrush = false;
            }
            if(this.recalcInfo.recalculatePen) {
                this.recalculatePen();
                this.recalcInfo.recalculatePen = false;
            }
            if(this.recalcInfo.recalcStyle) {
                this.recalculateStyle();
                this.recalcInfo.recalcStyle = false;
            }
            if(this.recalcInfo.recalculateTxBody) {
                this.recalculateTxBody();
                this.recalcInfo.recalculateTxBody = false;
            }
            if(this.recalcInfo.recalculateContent) {
                this.recalculateContent();
                this.recalcInfo.recalculateContent = false;
            }
            if(this.recalcInfo.recalcTransform) {
                this.recalculateTransform();
                this.recalcInfo.recalcTransform = false;
            }
            if(this.recalcInfo.recalculateGeometry) {
                this.recalculateGeometry && this.recalculateGeometry();
                this.recalcInfo.recalculateGeometry = false;
            }
            if(this.recalcInfo.recalculateTransformText) {
                this.recalculateTransformText();
                this.recalcInfo.recalculateTransformText = false;
            }
            if(this.chart) {
                this.chart.addToSetPosition(this);
            }
        }, this, []);
    };
    CTitle.prototype.setLayout = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Title_SetLayout, this.layout, pr));
        this.layout = pr;
        this.setParentToChild(pr);
    };
    CTitle.prototype.setOverlay = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Title_SetOverlay, this.overlay, pr));
        this.overlay = pr;
        this.onUpdate();
    };
    CTitle.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Title_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CTitle.prototype.setTx = function(pr) {
        if(this.tx && this.tx.strRef || pr && pr.strRef) {
            this.onChangeDataRefs();
        }
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Title_SetTx, this.tx, pr));
        this.tx = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CTitle.prototype.setTxPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Title_SetTxPr, this.txPr, pr));
        this.txPr = pr;
        this.setParentToChild(pr);
        this.onUpdate();
    };
    CTitle.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        var nParentType = this.parent.getObjectType();
        if(nParentType === AscDFH.historyitem_type_Chart) {
            this.applyStyleEntry(oChartStyle.title, oColors.generateColors(1), 0, bReset);
        }
        else if(nParentType === AscDFH.historyitem_type_ValAx ||
            nParentType === AscDFH.historyitem_type_CatAx ||
            nParentType === AscDFH.historyitem_type_DateAx ||
            nParentType === AscDFH.historyitem_type_SerAx) {
            this.applyStyleEntry(oChartStyle.axisTitle, oColors.generateColors(1), 0, bReset);
        }
    };

    function CTrendLine() {
        CBaseChartObject.call(this);
        this.backward = null;
        this.dispEq = null;
        this.dispRSqr = null;
        this.forward = null;
        this.intercept = null;
        this.name = null;
        this.order = null;
        this.period = null;
        this.spPr = null;
        this.trendlineLbl = null;
        this.trendlineType = null;
    }

    InitClass(CTrendLine, CBaseChartObject, AscDFH.historyitem_type_TrendLine);
    CTrendLine.prototype.setBackward = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Trendline_SetBackward, this.backward, pr));
        this.backward = pr;
    };
    CTrendLine.prototype.setDispEq = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Trendline_SetDispEq, this.dispEq, pr));
        this.dispEq = pr;
    };
    CTrendLine.prototype.setDispRSqr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Trendline_SetDispRSqr, this.dispRSqr, pr));
        this.dispRSqr = pr;
    };
    CTrendLine.prototype.setForward = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Trendline_SetForward, this.forward, pr));
        this.forward = pr;
    };
    CTrendLine.prototype.setIntercept = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_Trendline_SetIntercept, this.intercept, pr));
        this.intercept = pr;
    };
    CTrendLine.prototype.setName = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_Trendline_SetName, this.name, pr));
        this.name = pr;
    };
    CTrendLine.prototype.setOrder = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Trendline_SetOrder, this.order, pr));
        this.order = pr;
    };
    CTrendLine.prototype.setPeriod = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Trendline_SetPeriod, this.period, pr));
        this.period = pr;
    };
    CTrendLine.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Trendline_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CTrendLine.prototype.setTrendlineLbl = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Trendline_SetTrendlineLbl, this.trendlineLbl, pr));
        this.trendlineLbl = pr;
        this.setParentToChild(pr);
    };
    CTrendLine.prototype.setTrendlineType = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Trendline_SetTrendlineType, this.trendlineType, pr));
        this.trendlineType = pr;
    };
    CTrendLine.prototype.getChildren = function() {
        return [this.spPr];
    };
    CTrendLine.prototype.fillObject = function(oCopy, oIdMap) {
        if(AscFormat.isRealNumber(this.backward)) {
            oCopy.setBackward(this.backward);
        }
        if(AscFormat.isRealBool(this.dispEq)) {
            oCopy.setDispEq(this.dispEq);
        }
        if(AscFormat.isRealBool(this.dispRSqr)) {
            oCopy.setDispRSqr(this.dispRSqr);
        }
        if(AscFormat.isRealNumber(this.forward)) {
            oCopy.setForward(this.forward);
        }
        if(AscFormat.isRealNumber(this.intercept)) {
            oCopy.setIntercept(this.intercept);
        }
        if(typeof this.name === "string") {
            oCopy.setName(this.name);
        }
        if(AscFormat.isRealNumber(this.order)) {
            oCopy.setOrder(this.order);
        }
        if(AscFormat.isRealNumber(this.period)) {
            oCopy.setPeriod(this.period);
        }
        if(isRealObject(this.spPr)) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(AscFormat.isRealNumber(this.trendlineType)) {
            oCopy.setTrendlineType(this.trendlineType);
        }
    };
    CTrendLine.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.trendline, oColors.generateColors(1), 0, bReset);
    };

    function CUpDownBars() {
        CBaseChartObject.call(this);
        this.downBars = null;
        this.gapWidth = null;
        this.upBars = null;
    }

    InitClass(CUpDownBars, CBaseChartObject, AscDFH.historyitem_type_UpDownBars);
    CUpDownBars.prototype.Refresh_RecalcData = function() {
        if(this.parent) {
            this.parent.Refresh_RecalcData && this.parent.Refresh_RecalcData();
        }
    };
    CUpDownBars.prototype.setDownBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_UpDownBars_SetDownBars, this.downBars, pr));
        this.downBars = pr;
        this.setParentToChild(pr);
    };
    CUpDownBars.prototype.setGapWidth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_UpDownBars_SetGapWidth, this.downBars, pr));
        this.gapWidth = pr;
    };
    CUpDownBars.prototype.setUpBars = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_UpDownBars_SetUpBars, this.downBars, pr));
        this.upBars = pr;
        this.setParentToChild(pr);
    };
    CUpDownBars.prototype.handleUpdateFill = function() {
        this.Refresh_RecalcData();
    };
    CUpDownBars.prototype.handleUpdateLn = function() {
        this.Refresh_RecalcData();
    };
    CUpDownBars.prototype.getChildren = function() {
        return [this.upBars, this.downBars];
    };
    CUpDownBars.prototype.fillObject = function(oCopy, oIdMap) {
        if(AscFormat.isRealNumber(this.gapWidth)) {
            oCopy.setGapWidth(this.gapWidth);
        }
        if(isRealObject(this.upBars)) {
            oCopy.setUpBars(this.upBars.createDuplicate());
        }
        if(isRealObject(this.downBars)) {
            oCopy.setDownBars(this.downBars.createDuplicate());
        }
    };

    function CYVal() {
        CBaseChartObject.call(this);
        this.numLit = null;
        this.numRef = null;
    }

    InitClass(CYVal, CBaseChartObject, AscDFH.historyitem_type_YVal);
    CYVal.prototype.getChildren = function() {
        return [this.numLit, this.numRef];
    };
    CYVal.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.numLit) {
            oCopy.setNumLit(this.numLit.createDuplicate());
        }
        if(this.numRef) {
            oCopy.setNumRef(this.numRef.createDuplicate());
        }
    };
    CYVal.prototype.setNumLit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_YVal_SetNumLit, this.numLit, pr));
        this.numLit = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CYVal.prototype.setNumRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_YVal_SetNumRef, this.numRef, pr));
        this.numRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CYVal.prototype.setValues = function(sValues) {
        var oResult;
        if(typeof sValues === "string" && sValues.length > 0) {
            oResult = new CParseResult();
            fParseNumLit(sValues, true, oResult);
            if(oResult.isSuccessful()) {
                if(this.numRef) {
                    this.setNumRef(null);
                }
                this.setNumLit(oResult.getObject());
            }
            else {
                oResult = new CParseResult();
                fParseNumRef(sValues, true, oResult);
                if(oResult.isSuccessful()) {
                    if(this.numLit) {
                        this.setNumLit(null);
                    }
                    this.setNumRef(oResult.getObject());
                }
            }
        }
        else {
            oResult = new CParseResult();
            oResult.setError(Asc.c_oAscError.NoValues);
        }
        return oResult;
    };
    CYVal.prototype.getValuesCount = function() {
        if(this.numLit) {
            return this.numLit.ptCount;
        }
        if(this.numRef) {
            return this.numRef.getValuesCount();
        }
        return 0;
    };
    CYVal.prototype.getValues = function(nMaxValues) {
        if(this.numLit) {
            return this.numLit.getValues(nMaxValues);
        }
        if(this.numRef) {
            return this.numRef.getValues(nMaxValues);
        }
    };
    CYVal.prototype.getFormula = function() {
        if(this.numLit) {
            return this.numLit.getFormula();
        }
        if(this.numRef) {
            return this.numRef.getFormula();
        }
    };
    CYVal.prototype.getRefFormula = function() {
        if(this.numRef) {
            return this.numRef.getFormula();
        }
        return null;
    };
    CYVal.prototype.getDataRefs = function() {
        if(this.numRef) {
            return this.numRef.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CYVal.prototype.collectRefs = function(aRefs) {
        if(this.numRef) {
            aRefs.push(this.numRef);
        }
    };
    CYVal.prototype.isValid = function() {
        if(this.numRef || this.numLit) {
            return true;
        }
        return false;
    };
    CYVal.prototype.clearDataCache = function() {
        if(this.numRef) {
            this.numRef.clearDataCache();
        }
    };
    CYVal.prototype.update = function(displayEmptyCellsAs, displayHidden, ser) {
        if(this.numRef) {
            this.numRef.updateCache(displayEmptyCellsAs, displayHidden, ser);
        }
    };
    CYVal.prototype.getSourceNumFormat = function(nPtIdx) {
        if(this.numRef) {
            return this.numRef.getNumFormat(nPtIdx);
        }
        if(this.numLit) {
            return this.numLit.getNumFormat(nPtIdx);
        }
        return "General";
    };
    CYVal.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.numRef) {
            this.numRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CYVal.prototype.fillFromAsc = function(oValCache, bUseCache) {
        this.setNumRef(new AscFormat.CNumRef());
        var oNumRef = this.numRef;
        oNumRef.fillFromAsc(oValCache, bUseCache);
    };
    CYVal.prototype.getNumCache = function() {
        if(this.numRef && this.numRef.numCache) {
            return this.numRef.numCache;
        }
        else if(this.numLit) {
            return this.numLit;
        }
        return null;
    };

    function CCat() {
        CBaseChartObject.call(this);
        this.multiLvlStrRef = null;
        this.numLit = null;
        this.numRef = null;
        this.strLit = null;
        this.strRef = null;

        this.calculatedRef = null;
    }

    InitClass(CCat, CBaseChartObject, AscDFH.historyitem_type_Cat);
    CCat.prototype.getChildren = function() {
        return [this.multiLvlStrRef, this.numLit, this.numRef, this.strLit, this.strRef];
    };
    CCat.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.multiLvlStrRef) {
            oCopy.setMultiLvlStrRef(this.multiLvlStrRef.createDuplicate());
        }
        if(this.numLit) {
            oCopy.setNumLit(this.numLit.createDuplicate());
        }
        if(this.numRef) {
            oCopy.setNumRef(this.numRef.createDuplicate());
        }
        if(this.strLit) {
            oCopy.setStrLit(this.strLit.createDuplicate());
        }
        if(this.strRef) {
            oCopy.setStrRef(this.strRef.createDuplicate());
        }
    };
    CCat.prototype.setMultiLvlStrRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Cat_SetMultiLvlStrRef, this.multiLvlStrRef, pr));
        this.multiLvlStrRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CCat.prototype.setNumLit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Cat_SetNumLit, this.multiLvlStrRef, pr));
        this.numLit = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CCat.prototype.setNumRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Cat_SetNumRef, this.multiLvlStrRef, pr));
        this.numRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CCat.prototype.setStrLit = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Cat_SetStrLit, this.multiLvlStrRef, pr));
        this.strLit = pr;
        this.setParentToChild(pr);
    };
    CCat.prototype.setStrRef = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Cat_SetStrRef, this.multiLvlStrRef, pr));
        this.strRef = pr;
        this.onChangeDataRefs();
        this.setParentToChild(pr);
    };
    CCat.prototype.setValues = function(sValues) {
        this.calculatedRef = null;
        var oNumRef, oNumLit, oStrRef, oStrLit, oMultiLvl, oRef, oResult;
        oResult = new CParseResult();
        fParseNumRef(sValues, false, oResult);
        oNumRef = oResult.getObject();
        if(!oNumRef) {
            fParseStrRef(sValues, true, oResult);
            oRef = oResult.getObject();
            if(oRef) {
                if(oRef.getObjectType() === AscDFH.historyitem_type_StrRef) {
                    oStrRef = oRef;
                }
                else if(oRef.getObjectType() === AscDFH.historyitem_type_MultiLvlStrRef) {
                    oMultiLvl = oRef;
                }
            }
            if(!oStrRef && !oMultiLvl) {
                fParseNumLit(sValues, false, oResult);
                oNumLit = oResult.getObject();
                if(!oNumLit) {
                    fParseStrLit(sValues, oResult);
                    oStrLit = oResult.getObject();
                }
            }
        }
        if(oNumRef || oNumLit || oStrRef || oStrLit || oMultiLvl) {
            if(oNumRef) {
                this.setNumRef(oNumRef);
                this.setNumLit(null);
                this.setStrRef(null);
                this.setStrLit(null);
                this.setMultiLvlStrRef(null);
            }
            else if(oStrRef) {
                this.setStrRef(oStrRef);
                this.setNumRef(null);
                this.setNumLit(null);
                this.setStrLit(null);
                this.setMultiLvlStrRef(null);
            }
            else if(oMultiLvl) {
                this.setMultiLvlStrRef(oMultiLvl);
                this.setNumRef(null);
                this.setNumLit(null);
                this.setStrRef(null);
                this.setStrLit(null);
            }
            else if(oNumLit) {
                this.setNumLit(oNumLit);
                this.setNumRef(null);
                this.setStrRef(null);
                this.setStrLit(null);
                this.setMultiLvlStrRef(null);
            }
            else if(oStrLit) {
                this.setStrLit(oStrLit);
                this.setNumRef(null);
                this.setNumLit(null);
                this.setStrRef(null);
                this.setMultiLvlStrRef(null);
            }
        }
        return oResult;
    };
    CCat.prototype.getValues = function(nMaxCount) {
        if(this.numLit) {
            return this.numLit.getValues(nMaxCount);
        }
        if(this.numRef) {
            return this.numRef.getValues(nMaxCount);
        }
        if(this.strLit) {
            return this.strLit.getValues(nMaxCount);
        }
        if(this.strRef) {
            return this.strRef.getValues(nMaxCount, true);
        }
        if(this.multiLvlStrRef) {
            return this.multiLvlStrRef.getValues(nMaxCount);
        }
    };
    CCat.prototype.getFormula = function() {
        if(this.numLit) {
            return this.numLit.getFormula();
        }
        if(this.numRef) {
            return this.numRef.getFormula();
        }
        if(this.strLit) {
            return this.strLit.getFormula();
        }
        if(this.strRef) {
            return this.strRef.getFormula();
        }
        if(this.multiLvlStrRef) {
            return this.multiLvlStrRef.getFormula();
        }
        return "";
    };
    CCat.prototype.getDataRefs = function() {
        if(this.numRef) {
            return this.numRef.getDataRefs();
        }
        if(this.strRef) {
            return this.strRef.getDataRefs();
        }
        if(this.multiLvlStrRef) {
            return this.multiLvlStrRef.getDataRefs();
        }
        return new CDataRefs([]);
    };
    CCat.prototype.collectRefs = function(aRefs) {
        if(this.numRef) {
            aRefs.push(this.numRef);
        }
        if(this.strRef) {
            aRefs.push(this.strRef);
        }
        if(this.multiLvlStrRef) {
            aRefs.push(this.multiLvlStrRef);
        }
    };
    CCat.prototype.isValid = function() {
        if(this.multiLvlStrRef ||
            this.numLit ||
            this.numRef ||
            this.strLit ||
            this.strRef) {
            return true;
        }
        return false;
    };
    CCat.prototype.getStringPointsLit = function() {
        if(this.calculatedRef) {
            if(this.calculatedRef.strCache) {
                return this.calculatedRef.strCache;
            }
            return null;
        }
        if(this.strRef && this.strRef.strCache) {
            return this.strRef.strCache;
        }
        else if(this.strLit) {
            return this.strLit;
        }
        else if(this.multiLvlStrRef) {
            return this.multiLvlStrRef.getFirstLvlCache();
        }
        return null;
    };
    CCat.prototype.getLit = function() {
        if(this.calculatedRef) {
            if(this.calculatedRef.strCache) {
                return this.calculatedRef.strCache;
            }
            if(this.calculatedRef.numCache) {
                return this.calculatedRef.numCache;
            }
            if(this.calculatedRef.getFirstLvlCache) {
                return this.calculatedRef.getFirstLvlCache();
            }
            return null;
        }
        var oLit = null;
        if(this.strRef && this.strRef.strCache) {
            oLit = this.strRef.strCache;
        }
        else if(this.strLit) {
            oLit = this.strLit;
        }
        else if(this.numRef && this.numRef.numCache) {
            oLit = this.numRef.numCache;
        }
        else if(this.numLit) {
            oLit = this.numLit;
        }
        else if(this.multiLvlStrRef) {
            //TODO: implement multiLvlStrCache and remove this
            oLit = this.multiLvlStrRef.getFirstLvlCache();
        }
        return oLit;
    };
    CCat.prototype.clearDataCache = function() {
        if(this.numRef) {
            this.numRef.clearDataCache();
        }
        if(this.strRef) {
            this.strRef.clearDataCache();
        }
        if(this.multiLvlStrRef) {
            this.multiLvlStrRef.clearDataCache();
        }
    };
    CCat.prototype.update = function(oSeries) {
        return AscFormat.ExecuteNoHistory(function(){
            this.calculatedRef = null;
            if(this.numRef || this.strRef || this.multiLvlStrRef) {
                var sFormula = this.getFormula();
                if(typeof sFormula === "string" && sFormula.length > 0) {
                    var oTestCat = new CCat();
                    var oRes = oTestCat.setValues(sFormula);
                    var oNumRef = oTestCat.numRef;
                    var oStrRef = oTestCat.strRef;
                    var oMultiLvlStrRef = oTestCat.multiLvlStrRef;
                    if(oRes && oRes.error === Asc.c_oAscError.ID.No && (oNumRef || oStrRef || oMultiLvlStrRef)) {
                        this.calculatedRef = oNumRef || oStrRef || oMultiLvlStrRef;
                        if(this.calculatedRef) {
                            if(this.calculatedRef.getObjectType() === AscDFH.historyitem_type_MultiLvlStrRef) {
                                this.calculatedRef.updateCache(oSeries);
                            }
                            else {
                                this.calculatedRef.updateCache();
                            }
                        }
                    }
                }
            }
            if(this.multiLvlStrRef) {
                this.multiLvlStrRef.updateCache(oSeries);
            }
            if(this.numRef) {
                this.numRef.updateCache();
            }
            if(this.strRef) {
                this.strRef.updateCache();
            }
            if(!this.calculatedRef) {
                this.calculatedRef = (this.multiLvlStrRef || this.numRef || this.strRef);
            }
        }, this, []);
    };
    CCat.prototype.getSourceNumFormat = function() {
        if(this.calculatedRef) {
            if(this.calculatedRef.getObjectType() === AscDFH.historyitem_type_NumRef) {
                return this.calculatedRef.getNumFormat();
            }
            return "General";
        }
        if(this.numRef) {
            return this.numRef.getNumFormat();
        }
        if(this.numLit) {
            return this.numLit.getNumFormat();
        }
        return "General";
    };
    CCat.prototype.handleOnChangeSheetName = function(sOldSheetName, sNewSheetName) {
        if(this.multiLvlStrRef) {
            this.multiLvlStrRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.numRef) {
            this.numRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
        if(this.strRef) {
            this.strRef.handleOnChangeSheetName(sOldSheetName, sNewSheetName);
        }
    };
    CCat.prototype.fillFromAsc = function(oCatCache, bUseCache) {
        var bVal = false;
        var sFormatCode = oCatCache.formatCode;
        var oNumFormat = null;
        if(typeof sFormatCode === "string" && sFormatCode.length > 0) {
            oNumFormat = AscCommon.oNumFormatCache.get(oCatCache.formatCode);
        }
        if(oNumFormat && oNumFormat.isDateTimeFormat()) {
            var aPts = oCatCache.NumCache, oPt, nPt;
            for(nPt = 0; nPt < aPts.length; ++nPt) {
                oPt = aPts[nPt];
                if(oPt) {
                    sFormatCode = oPt.numFormatStr;
                    if(typeof sFormatCode === "string" && sFormatCode.length > 0) {
                        oNumFormat = AscCommon.oNumFormatCache.get(sFormatCode);
                        if(!oNumFormat.isDateTimeFormat() || !AscFormat.isRealNumber(parseFloat(oPt.val))) {
                            break;
                        }
                    }
                }
            }
            if(nPt === aPts.length) {
                bVal = true;
            }
        }
        if(bVal) {
            this.setNumRef(new CNumRef());
            this.numRef.fillFromAsc(oCatCache, bUseCache);
        }
        else {
            this.setStrRef(new CStrRef());
            this.strRef.fillFromAsc(oCatCache, bUseCache);
        }
    };
    CCat.prototype.getNumCache = function() {
        if(this.calculatedRef) {
            if(this.calculatedRef.getObjectType() === AscDFH.historyitem_type_NumRef) {
                return this.calculatedRef.numCache;
            }
            return null;
        }
        if(this.numRef && this.numRef.numCache) {
            return this.numRef.numCache;
        }
        else if(this.numLit) {
            return this.numLit;
        }
        return null;
    };

    function CChart() {
        CBaseChartObject.call(this);
        this.autoTitleDeleted = null;
        this.backWall = null;
        this.dispBlanksAs = null;
        this.floor = null;
        this.legend = null;
        this.pivotFmts = [];
        this.plotArea = null;
        this.plotVisOnly = null;
        this.showDLblsOverMax = false;
        this.sideWall = null;
        this.title = null;
        this.view3D = null;
    }

    InitClass(CChart, CBaseChartObject, AscDFH.historyitem_type_Chart);
    CChart.prototype.Refresh_RecalcData = function() {
        this.onChartInternalUpdate();
    };
    CChart.prototype.CheckCorrect = function() {
        if(!this.plotArea) {
            return false;
        }
        return true;
    };
    CChart.prototype.getView3d = function() {
        return AscFormat.ExecuteNoHistory(function() {
            var _ret;
            var oChart = this.plotArea && this.plotArea.charts[0];
            if(oChart) {
                if(this.view3D) {
                    _ret = this.view3D.createDuplicate();
                    if(oChart.getObjectType() === AscDFH.historyitem_type_SurfaceChart) {
                        if(!AscFormat.isRealNumber(_ret.rotX)) {
                            _ret.rotX = 15;
                        }
                        if(!AscFormat.isRealNumber(_ret.rotY)) {
                            _ret.rotY = 20;
                        }
                    }
                    else {
                        if(!AscFormat.isRealNumber(_ret.rotX)) {
                            _ret.rotX = 0;
                        }
                        if(!AscFormat.isRealNumber(_ret.rotY)) {
                            _ret.rotY = 0;
                        }
                    }
                    return _ret;
                }
                if(oChart.b3D) {
                    _ret = new CView3d();
                    _ret.setRotX(30);
                    _ret.setRotY(0);
                    _ret.setRAngAx(false);
                    _ret.setDepthPercent(100);
                    return _ret;
                }
            }
            return null;
        }, this, []);
    };
    CChart.prototype.getParentObjects = function() {
        return this.parent && this.parent.getParentObjects();
    };
    CChart.prototype.handleUpdateDataLabels = function() {
        this.onChartUpdateDataLabels();
    };
    CChart.prototype.handleUpdateFill = function() {
        if(this.parent && this.parent.handleUpdateFill) {
            this.parent.handleUpdateFill();
        }
    };
    CChart.prototype.handleUpdateLn = function() {
        if(this.parent && this.parent.handleUpdateLn) {
            this.parent.handleUpdateLn();
        }
    };
    CChart.prototype.setDefaultWalls = function() {
        var oFloor = new CChartWall();
        oFloor.setThickness(0);
        this.setFloor(oFloor);
        var oSideWall = new CChartWall();
        oSideWall.setThickness(0);
        this.setSideWall(oSideWall);
        var oBackWall = new CChartWall();
        oBackWall.setThickness(0);
        this.setBackWall(oBackWall);
    };
    CChart.prototype.getChildren = function() {
        var aRet = [];
        aRet.push(this.backWall);
        aRet.push(this.floor);
        aRet.push(this.legend);
        var Count = this.pivotFmts.length;
        for(var i = 0; i < Count; i++) {
            aRet.push(this.pivotFmts[i]);
        }
        aRet.push(this.plotArea);
        aRet.push(this.sideWall);
        aRet.push(this.title);
        aRet.push(this.view3D);
        return aRet;
    };
    CChart.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setAutoTitleDeleted(this.autoTitleDeleted);
        if(this.backWall) {
            oCopy.setBackWall(this.backWall.createDuplicate());
        }
        if(this.dispBlanksAs !== null) {
            oCopy.setDispBlanksAs(this.dispBlanksAs);
        }
        if(this.floor) {
            oCopy.setFloor(this.floor.createDuplicate());
        }
        if(this.legend) {
            oCopy.setLegend(this.legend.createDuplicate());
        }
        var Count = this.pivotFmts.length;
        for(var i = 0; i < Count; i++) {
            oCopy.setPivotFmts(this.pivotFmts[i].createDuplicate());
        }
        if(this.plotArea) {
            oCopy.setPlotArea(this.plotArea.createDuplicate());
        }
        if(this.plotVisOnly !== null) {
            oCopy.setPlotVisOnly(this.plotVisOnly);
        }
        if(this.showDLblsOverMax !== null) {
            oCopy.setShowDLblsOverMax(this.showDLblsOverMax);
        }
        if(this.sideWall) {
            oCopy.setSideWall(this.sideWall.createDuplicate());
        }
        if(this.title) {
            oCopy.setTitle(this.title.createDuplicate());
        }
        if(this.view3D) {
            oCopy.setView3D(this.view3D.createDuplicate());
        }
    };
    CChart.prototype.Refresh_RecalcData2 = function(pageIndex, object) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(pageIndex, object);
        }
    };
    CChart.prototype.getDrawingDocument = function() {
        if(this.parent) {
            return this.parent.getDrawingDocument();
        }
        return null;
    };
    CChart.prototype.setAutoTitleDeleted = function(autoTitleDeleted) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Chart_SetAutoTitleDeleted, this.autoTitleDeleted, autoTitleDeleted));
        this.autoTitleDeleted = autoTitleDeleted;
    };
    CChart.prototype.setBackWall = function(backWall) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetBackWall, this.backWall, backWall));
        this.backWall = backWall;
        this.setParentToChild(backWall);
    };
    CChart.prototype.setDispBlanksAs = function(dispBlanksAs) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_Chart_SetDispBlanksAs, this.dispBlanksAs, dispBlanksAs));
        this.dispBlanksAs = dispBlanksAs;
    };
    CChart.prototype.setFloor = function(floor) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetFloor, this.floor, floor));
        this.floor = floor;
        this.setParentToChild(floor);
    };
    CChart.prototype.setLegend = function(legend) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetLegend, this.legend, legend));
        this.legend = legend;
        this.setParentToChild(legend);
    };
    CChart.prototype.setPivotFmts = function(pivotFmt) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsContent(this, AscDFH.historyitem_Chart_AddPivotFmt, this.pivotFmts.length, [pivotFmt], true));
        this.pivotFmts.push(pivotFmt);
        this.setParentToChild(pivotFmt);
    };
    CChart.prototype.setPlotArea = function(plotArea) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetPlotArea, this.plotArea, plotArea));
        this.plotArea = plotArea;
        this.setParentToChild(plotArea);
    };
    CChart.prototype.setPlotVisOnly = function(plotVisOnly) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Chart_SetPlotVisOnly, this.plotVisOnly, plotVisOnly));
        this.plotVisOnly = plotVisOnly;
    };
    CChart.prototype.setShowDLblsOverMax = function(showDLblsOverMax) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Chart_SetShowDLblsOverMax, this.showDLblsOverMax, showDLblsOverMax));
        this.showDLblsOverMax = showDLblsOverMax;
    };
    CChart.prototype.setSideWall = function(sideWall) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetSideWall, this.sideWall, sideWall));
        this.sideWall = sideWall;
        this.setParentToChild(sideWall);
    };
    CChart.prototype.setTitle = function(title) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetTitle, this.title, title));
        this.title = title;
        this.setParentToChild(title);
        this.onChartInternalUpdate();
        this.onChangeDataRefs();
    };
    CChart.prototype.setView3D = function(view3D) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_Chart_SetView3D, this.view3D, view3D));
        this.view3D = view3D;
        this.setParentToChild(view3D);
        this.onChartInternalUpdate();
    };
    CChart.prototype.reindexSeries = function() {
        if(this.parent) {
            this.parent.reindexSeries();
        }
    };
    CChart.prototype.reorderSeries = function() {
        if(this.parent) {
            this.parent.reorderSeries();
        }
    };
    CChart.prototype.moveSeriesUp = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesUp(oSeries);
        }
    };
    CChart.prototype.moveSeriesDown = function(oSeries) {
        if(this.parent) {
            this.parent.moveSeriesDown(oSeries);
        }
    };
    CChart.prototype.onDataUpdate = function() {
        if(this.parent) {
            this.parent.onDataUpdate();
        }
    };
    CChart.prototype.setDLblsDeleteValue = function(bVal) {
        if(this.plotArea) {
            this.plotArea.setDLblsDeleteValue(bVal);
        }
    };
    CChart.prototype.getChartType = function() {
        if(this.plotArea) {
            return this.plotArea.getChartType();
        }
        return Asc.c_oAscChartTypeSettings.unknown;
    };
    CChart.prototype.changeChartType = function(nType) {
        if(this.plotArea) {
            this.plotArea.changeChartType(nType);
        }
    };
    CChart.prototype.checkDlblsPosition = function() {
        this.plotArea.checkDlblsPosition();
    };
    CChart.prototype.getPossibleDLblsPosition = function() {
        return this.plotArea.getPossibleDLblsPosition();
    };
    CChart.prototype.check3DOptions = function(is3D, bPerspective) {
        if(is3D) {
            if(!this.view3D) {
                this.setView3D(new AscFormat.CView3d());
            }
            var oView3d = this.view3D;
            if(!AscFormat.isRealNumber(oView3d.rotX)) {
                oView3d.setRotX(15);
            }
            if(!AscFormat.isRealNumber(oView3d.rotY)) {
                oView3d.setRotY(20);
            }
            if(!AscFormat.isRealBool(oView3d.rAngAx)) {
                oView3d.setRAngAx(true);
            }
            if(bPerspective) {
                if(!AscFormat.isRealNumber(oView3d.depthPercent)) {
                    oView3d.setDepthPercent(100);
                }
            }
            else {
                if(null !== oView3d.depthPercent) {
                    oView3d.setDepthPercent(null);
                }
            }
            this.setDefaultWalls();
        }
        else {
            if(this.view3D) {
                this.setView3D(null);
            }
            if(this.floor) {
                this.setFloor(null);
            }
            if(this.sideWall) {
                this.setSideWall(null);
            }
            if(this.backWall) {
                this.setBackWall(null);
            }
        }
    };
    CChart.prototype.getOrderedAxes = function() {
        return this.plotArea.getOrderedAxes();
    };
    CChart.prototype.setDlblsProps = function(oProps) {
        return this.plotArea.setDlblsProps(oProps);
    };
    CChart.prototype.is3dChart = function() {
        if(this.parent) {
            return this.parent.is3dChart();
        }
        return false;
    };
    CChart.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        if(bReset && oAdditionalData) {
            if(!oAdditionalData.view3D) {
                this.setView3D(null);
            }
            else {
                this.setView3D(oAdditionalData.view3D.createDuplicate());
            }
            if(!oAdditionalData.legend) {
                this.setLegend(null);
            }
            else {
                this.setLegend(oAdditionalData.legend.createDuplicate());
            }
        }
        if(this.backWall) {
            this.backWall.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.floor) {
            this.floor.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.legend) {
            this.legend.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.plotArea) {
            this.plotArea.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.sideWall) {
            this.sideWall.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
        if(this.title) {
            this.title.applyChartStyle(oChartStyle, oColors, oAdditionalData, bReset);
        }
    };

    function CChartWall() {
        CBaseChartObject.call(this);
        this.pictureOptions = null;
        this.spPr = null;
        this.thickness = null;
    }

    InitClass(CChartWall, CBaseChartObject, AscDFH.historyitem_type_ChartWall);
    CChartWall.prototype.getChildren = function() {
        return [this.pictureOptions, this.spPr];
    };
    CChartWall.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.pictureOptions) {
            oCopy.setPictureOptions(this.pictureOptions.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
        if(AscFormat.isRealNumber(this.thickness)) {
            oCopy.setThickness(this.thickness);
        }
    };
    CChartWall.prototype.setPictureOptions = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartWall_SetPictureOptions, this.pictureOptions, pr));
        this.pictureOptions = pr;
        this.setParentToChild(pr);
    };
    CChartWall.prototype.setSpPr = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartWall_SetSpPr, this.spPr, pr));
        this.spPr = pr;
        this.setParentToChild(pr);
    };
    CChartWall.prototype.setThickness = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ChartWall_SetThickness, this.thickness, pr));
        this.thickness = pr;
    };
    CChartWall.prototype.applyChartStyle = function(oChartStyle, oColors, oAdditionalData, bReset) {
        if(!this.parent) {
            return;
        }
        this.applyStyleEntry(oChartStyle.wall, oColors.generateColors(1), 0, bReset);
    };

    function CView3d() {
        CBaseChartObject.call(this);
        this.depthPercent = null;
        this.hPercent = null;
        this.perspective = null;
        this.rAngAx = null;
        this.rotX = null;
        this.rotY = null;
    }

    InitClass(CView3d, CBaseChartObject, AscDFH.historyitem_type_View3d);
    CView3d.prototype.fillObject = function(oCopy, oIdMap) {
        AscFormat.isRealNumber(this.depthPercent) && oCopy.setDepthPercent(this.depthPercent);
        AscFormat.isRealNumber(this.hPercent) && oCopy.setHPercent(this.hPercent);
        AscFormat.isRealNumber(this.perspective) && oCopy.setPerspective(this.perspective);
        AscFormat.isRealBool(this.rAngAx) && oCopy.setRAngAx(this.rAngAx);
        AscFormat.isRealNumber(this.rotX) && oCopy.setRotX(this.rotX);
        AscFormat.isRealNumber(this.rotY) && oCopy.setRotY(this.rotY);
    };
    CView3d.prototype.setDepthPercent = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_View3d_SetDepthPercent, this.depthPercent, pr));
        this.depthPercent = pr;
    };
    CView3d.prototype.setHPercent = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_View3d_SetHPercent, this.hPercent, pr));
        this.hPercent = pr;
    };
    CView3d.prototype.setPerspective = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_View3d_SetPerspective, this.perspective, pr));
        this.perspective = pr;
    };
    CView3d.prototype.setRAngAx = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_View3d_SetRAngAx, this.rAngAx, pr));
        this.rAngAx = pr;
    };
    CView3d.prototype.getRAngAx = function() {

        if(AscFormat.isRealBool(this.rAngAx)) {
            return this.rAngAx;
        }
        if(AscFormat.isRealNumber(this.perspective)) {
            return false;
        }
        return this.rAngAx !== false;
    };
    CView3d.prototype.setRotX = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_View3d_SetRotX, this.rotX, pr));
        this.rotX = pr;
    };
    CView3d.prototype.setRotY = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_View3d_SetRotY, this.rotY, pr));
        this.rotY = pr;
    };
    CView3d.prototype.isEqual = function(pr) {
        return this.depthPercent === pr.depthPercent &&
                this.hPercent === pr.hPercent &&
                this.perspective === pr.perspective &&
                this.rAngAx === pr.rAngAx &&
                this.rotX === pr.rotX &&
                this.rotY === pr.rotY;
    };
    CView3d.prototype.asc_getRotX = function() {
        return this.rotY; //ms changes x and y in interface
    };
    CView3d.prototype["asc_getRotX"] = CView3d.prototype.asc_getRotX;
    CView3d.prototype.asc_getRotY = function() {
        return this.rotX; //ms changes x and y in interface
    };
    CView3d.prototype["asc_getRotY"] = CView3d.prototype.asc_getRotY;
    CView3d.prototype.asc_getPerspective = function() {
        if(this.asc_getRightAngleAxes()) {
            return null;
        }
        if(AscFormat.isRealNumber(this.perspective)) {
            return this.perspective;
        }
        return AscFormat.global3DPersperctive;
    };
    CView3d.prototype["asc_getPerspective"] = CView3d.prototype.asc_getPerspective;
    CView3d.prototype.asc_getRightAngleAxes = function() {
        return this.getRAngAx();
    };
    CView3d.prototype["asc_getRightAngleAxes"] = CView3d.prototype.asc_getRightAngleAxes;
    CView3d.prototype.asc_getDepth = function() {
        return AscFormat.isRealNumber(this.depthPercent) ? this.depthPercent : AscFormat.globalBasePercent;
    };
    CView3d.prototype["asc_getDepth"] = CView3d.prototype.asc_getDepth;

    CView3d.prototype.asc_getHeight = function() {
        return this.hPercent || null;
    };
    CView3d.prototype["asc_getHeight"] = CView3d.prototype.asc_getHeight;

    CView3d.prototype.asc_setRotX = function(pr) {
        this.rotY = pr;//ms changes x and y in interface
    };
    CView3d.prototype["asc_setRotX"] = CView3d.prototype.asc_setRotX;
    CView3d.prototype.asc_setRotY = function(pr) {
        this.rotX = pr;//ms changes x and y in interface
    };
    CView3d.prototype["asc_setRotY"] = CView3d.prototype.asc_setRotY;
    CView3d.prototype.asc_setPerspective = function(pr) {
        this.perspective = pr;
    };
    CView3d.prototype["asc_setPerspective"] = CView3d.prototype.asc_setPerspective;
    CView3d.prototype.asc_setRightAngleAxes = function(pr) {
        this.rAngAx = pr;
        if(pr) {
            this.asc_setPerspective(null);
        }
    };
    CView3d.prototype["asc_setRightAngleAxes"] = CView3d.prototype.asc_setRightAngleAxes;
    CView3d.prototype.asc_setDepth = function(v) {
        this.depthPercent = v;
    };
    CView3d.prototype["asc_setDepth"] = CView3d.prototype.asc_setDepth;
    CView3d.prototype.asc_setHeight = function(v) {
        this.hPercent = v;
    };
    CView3d.prototype["asc_setHeight"] = CView3d.prototype.asc_setHeight;


    function CExternalData() {
        CBaseChartObject.call(this);
        this.autoUpdate = null;
        this.id = null;
    }

    InitClass(CExternalData, CBaseChartObject, AscDFH.historyitem_type_ExternalData);
    CExternalData.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.autoUpdate !== null) {
            oCopy.setAutoUpdate(this.autoUpdate);
        }
        if(this.id !== null) {
            oCopy.setId(this.id);
        }
    };
    CExternalData.prototype.setAutoUpdate = function(pr) {
        History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_ExternalData_SetAutoUpdate, this.autoUpdate, pr));
        this.autoUpdate = pr;
    };
    CExternalData.prototype.setId = function(pr) {
        History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_ExternalData_SetId, this.id, pr));
        this.id = pr;
    };

    function CPivotSource() {
        CBaseChartObject.call(this);
        this.fmtId = null;
        this.name = null;
    }

    InitClass(CPivotSource, CBaseChartObject, AscDFH.historyitem_type_PivotSource);
    CPivotSource.prototype.setFmtId = function(pr) {
        History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PivotSource_SetFmtId, this.fmtId, pr));
        this.fmtId = pr;
    };
    CPivotSource.prototype.setName = function(pr) {
        History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_PivotSource_SetName, this.name, pr));
        this.name = pr;
    };
    CPivotSource.prototype.fillObject = function(oCopy, oIdMap) {
        if(AscFormat.isRealNumber(this.fmtId)) {
            oCopy.setFmtId(this.fmtId);
        }
        if(typeof this.name === "string") {
            oCopy.setName(this.name);
        }
    };

    function CProtection() {
        CBaseChartObject.call(this);
        this.chartObject = null;
        this.data = null;
        this.formatting = null;
        this.selection = null;
        this.userInterface = null;
    }

    InitClass(CProtection, CBaseChartObject, AscDFH.historyitem_type_Protection);
    CProtection.prototype.setChartObject = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Protection_SetChartObject, this.chartObject, pr));
        this.chartObject = pr;
    };
    CProtection.prototype.setData = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Protection_SetData, this.data, pr));
        this.data = pr;
    };
    CProtection.prototype.setFormatting = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Protection_SetFormatting, this.formatting, pr));
        this.formatting = pr;
    };
    CProtection.prototype.setSelection = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Protection_SetSelection, this.selection, pr));
        this.selection = pr;
    };
    CProtection.prototype.setUserInterface = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_Protection_SetUserInterface, this.userInterface, pr));
        this.userInterface = pr;
    };
    CProtection.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.chartObject !== null)
            oCopy.setChartObject(this.chartObject);
        if(this.data !== null)
            oCopy.setData(this.data);
        if(this.formatting !== null)
            oCopy.setFormatting(this.formatting);
        if(this.selection !== null)
            oCopy.setSelection(this.selection);
        if(this.userInterface !== null)
            oCopy.setUserInterface(this.userInterface);
    };


    function CPrintSettings() {
        CBaseChartObject.call(this);
        this.headerFooter = null;
        this.pageMargins = null;
        this.pageSetup = null;
    }

    InitClass(CPrintSettings, CBaseChartObject, AscDFH.historyitem_type_PrintSettings);
    CPrintSettings.prototype.getChildren = function() {
        return [this.headerFooter, this.pageMargins, this.pageSetup];
    };
    CPrintSettings.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.headerFooter)
            oCopy.setHeaderFooter(this.headerFooter.createDuplicate());

        if(this.pageMargins)
            oCopy.setPageMargins(this.pageMargins.createDuplicate());

        if(this.pageSetup)
            oCopy.setPageSetup(this.pageSetup.createDuplicate());
    };
    CPrintSettings.prototype.setHeaderFooter = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PrintSettingsSetHeaderFooter, this.headerFooter, pr));
        this.headerFooter = pr;
        this.setParentToChild(pr);
    };
    CPrintSettings.prototype.setPageMargins = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PrintSettingsSetPageMargins, this.pageMargins, pr));
        this.pageMargins = pr;
        this.setParentToChild(pr);
    };
    CPrintSettings.prototype.setPageSetup = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_PrintSettingsSetPageSetup, this.pageSetup, pr));
        this.pageSetup = pr;
        this.setParentToChild(pr);
    };
    CPrintSettings.prototype.setDefault = function() {
        this.setHeaderFooter(new CHeaderFooterChart());
        this.setPageMargins(new CPageMarginsChart());
        this.setPageSetup(new CPageSetup());
        this.pageMargins.setB(0.75);
        this.pageMargins.setL(0.7);
        this.pageMargins.setR(0.7);
        this.pageMargins.setT(0.75);
        this.pageMargins.setHeader(0.3);
        this.pageMargins.setFooter(0.3);
    };

    function CHeaderFooterChart() {
        CBaseChartObject.call(this);
        this.alignWithMargins = null;
        this.differentFirst = null;
        this.differentOddEven = null;
        this.evenFooter = null;
        this.evenHeader = null;
        this.firstFooter = null;
        this.firstHeader = null;
        this.oddFooter = null;
        this.oddHeader = null;
    }

    InitClass(CHeaderFooterChart, CBaseChartObject, AscDFH.historyitem_type_HeaderFooterChart);
    CHeaderFooterChart.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.alignWithMargins !== null)
            oCopy.setAlignWithMargins(this.alignWithMargins);
        if(this.differentFirst !== null)
            oCopy.setDifferentFirst(this.differentFirst);
        if(this.differentOddEven !== null)
            oCopy.setDifferentOddEven(this.differentOddEven);
        if(this.evenFooter !== null)
            oCopy.setEvenFooter(this.evenFooter);
        if(this.evenHeader !== null)
            oCopy.setEvenHeader(this.evenHeader);
        if(this.firstFooter !== null)
            oCopy.setFirstFooter(this.firstFooter);
        if(this.firstHeader !== null)
            oCopy.setFirstHeader(this.firstHeader);
        if(this.oddFooter !== null)
            oCopy.setOddFooter(this.oddFooter);
        if(this.oddHeader !== null)
            oCopy.setOddHeader(this.oddHeader);
    };
    CHeaderFooterChart.prototype.setAlignWithMargins = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooterChartSetAlignWithMargins, this.alignWithMargins, pr));
        this.alignWithMargins = pr;
    };
    CHeaderFooterChart.prototype.setDifferentFirst = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooterChartSetDifferentFirst, this.differentFirst, pr));
        this.differentFirst = pr;
    };
    CHeaderFooterChart.prototype.setDifferentOddEven = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_HeaderFooterChartSetDifferentOddEven, this.differentOddEven, pr));
        this.differentOddEven = pr;
    };
    CHeaderFooterChart.prototype.setEvenFooter = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetEvenFooter, this.evenFooter, pr));
        this.evenFooter = pr;
    };
    CHeaderFooterChart.prototype.setEvenHeader = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetEvenHeader, this.evenHeader, pr));
        this.evenHeader = pr;
    };
    CHeaderFooterChart.prototype.setFirstFooter = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetFirstFooter, this.firstFooter, pr));
        this.firstFooter = pr;
    };
    CHeaderFooterChart.prototype.setFirstHeader = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetFirstHeader, this.firstHeader, pr));
        this.firstHeader = pr;
    };
    CHeaderFooterChart.prototype.setOddFooter = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetOddFooter, this.oddFooter, pr));
        this.oddFooter = pr;
    };
    CHeaderFooterChart.prototype.setOddHeader = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsString(this, AscDFH.historyitem_HeaderFooterChartSetOddHeader, this.oddHeader, pr));
        this.oddHeader = pr;
    };

    function CPageMarginsChart() {
        CBaseChartObject.call(this);
        this.b = null;
        this.footer = null;
        this.header = null;
        this.l = null;
        this.r = null;
        this.t = null;
    }

    InitClass(CPageMarginsChart, CBaseChartObject, AscDFH.historyitem_type_PageMarginsChart);
    CPageMarginsChart.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.b !== null)
            oCopy.setB(this.b);
        if(this.footer !== null)
            oCopy.setFooter(this.footer);
        if(this.header !== null)
            oCopy.setHeader(this.header);
        if(this.l !== null)
            oCopy.setL(this.l);
        if(this.r !== null)
            oCopy.setR(this.r);
        if(this.t !== null)
            oCopy.setT(this.t);
    };
    CPageMarginsChart.prototype.setB = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetB, this.b, pr));
        this.b = pr;
    };
    CPageMarginsChart.prototype.setFooter = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetFooter, this.footer, pr));
        this.footer = pr;
    };
    CPageMarginsChart.prototype.setHeader = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetHeader, this.header, pr));
        this.header = pr;
    };
    CPageMarginsChart.prototype.setL = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetL, this.l, pr));
        this.l = pr;
    };
    CPageMarginsChart.prototype.setR = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetR, this.r, pr));
        this.r = pr;
    };
    CPageMarginsChart.prototype.setT = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageMarginsSetT, this.t, pr));
        this.t = pr;
    };

    function CPageSetup() {
        CBaseChartObject.call(this);
        this.blackAndWhite = null;
        this.copies = null;
        this.draft = null;
        this.firstPageNumber = null;
        this.horizontalDpi = null;
        this.orientation = null;
        this.paperHeight = null;
        this.paperSize = null;
        this.paperWidth = null;
        this.useFirstPageNumb = null;
        this.verticalDpi = null;
    }

    InitClass(CPageSetup, CBaseChartObject, AscDFH.historyitem_type_PageSetup);
    CPageSetup.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.blackAndWhite !== null)
            oCopy.setBlackAndWhite(this.blackAndWhite);
        if(this.copies !== null)
            oCopy.setCopies(this.copies);
        if(this.draft !== null)
            oCopy.setDraft(this.draft);
        if(this.firstPageNumber !== null)
            oCopy.setFirstPageNumber(this.firstPageNumber);
        if(this.horizontalDpi !== null)
            oCopy.setHorizontalDpi(this.horizontalDpi);
        if(this.orientation !== null)
            oCopy.setOrientation(this.orientation);
        if(this.paperHeight !== null)
            oCopy.setPaperHeight(this.paperHeight);
        if(this.paperSize !== null)
            oCopy.setPaperSize(this.paperSize);
        if(this.paperWidth !== null)
            oCopy.setPaperWidth(this.paperWidth);
        if(this.useFirstPageNumb !== null)
            oCopy.setUseFirstPageNumb(this.useFirstPageNumb);
        if(this.verticalDpi !== null)
            oCopy.setVerticalDpi(this.verticalDpi);
    };
    CPageSetup.prototype.setBlackAndWhite = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetupSetBlackAndWhite, this.blackAndWhite, pr));
        this.blackAndWhite = pr;
    };
    CPageSetup.prototype.setCopies = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetCopies, this.copies, pr));
        this.copies = pr;
    };
    CPageSetup.prototype.setDraft = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetupSetDraft, this.draft, pr));
        this.draft = pr;
    };
    CPageSetup.prototype.setFirstPageNumber = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetFirstPageNumber, this.firstPageNumber, pr));
        this.firstPageNumber = pr;
    };
    CPageSetup.prototype.setHorizontalDpi = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetHorizontalDpi, this.horizontalDpi, pr));
        this.horizontalDpi = pr;
    };
    CPageSetup.prototype.setOrientation = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetOrientation, this.orientation, pr));
        this.orientation = pr;
    };
    CPageSetup.prototype.setPaperHeight = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageSetupSetPaperHeight, this.paperHeight, pr));
        this.paperHeight = pr;
    };
    CPageSetup.prototype.setPaperSize = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetPaperSize, this.paperSize, pr));
        this.paperSize = pr;
    };
    CPageSetup.prototype.setPaperWidth = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsDouble(this, AscDFH.historyitem_PageSetupSetPaperWidth, this.paperWidth, pr));
        this.paperWidth = pr;
    };
    CPageSetup.prototype.setUseFirstPageNumb = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsBool(this, AscDFH.historyitem_PageSetupSetUseFirstPageNumb, this.useFirstPageNumb, pr));
        this.useFirstPageNumb = pr;
    };
    CPageSetup.prototype.setVerticalDpi = function(pr) {
        History.CanAddChanges() && History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_PageSetupSetVerticalDpi, this.verticalDpi, pr));
        this.verticalDpi = pr;
    };

    function CreateTextBodyFromString(str, drawingDocument, parent) {
        var tx_body = new AscFormat.CTextBody();
        tx_body.setParent(parent);
        tx_body.setBodyPr(new AscFormat.CBodyPr());
        var old_is_doc_editor = false;
        if(typeof editor !== "undefined" && editor && editor.isDocumentEditor) {
            editor.isDocumentEditor = false;
            old_is_doc_editor = true;
        }
        tx_body.setContent(CreateDocContentFromString(str, drawingDocument, tx_body));

        if(typeof editor !== "undefined" && editor && old_is_doc_editor) {
            editor.isDocumentEditor = true;
        }
        return tx_body;
    }

    function CreateDocContentFromString(str, drawingDocument, parent) {
        var content = new AscFormat.CDrawingDocContent(parent, drawingDocument, 0, 0, 0, 0, false, false, true);
        AddToContentFromString(content, str);
        return content;
    }

    function CheckContentTextAndAdd(oContent, sText) {
        oContent.SetApplyToAll(true);
        var sContentText = oContent.GetSelectedText(false, {NewLine: true, NewParagraph: true});
        oContent.SetApplyToAll(false);
        if(sContentText !== sText) {
            oContent.ClearContent(true);
            AddToContentFromString(oContent, sText);
        }
    }

    function AddToContentFromString(content, str) {
        content.MoveCursorToStartPos(false);
        content.AddText(str);
    }

    function CValAxisLabels(chart, axis) {
        this.x = null;
        this.y = null;
        this.extX = null;
        this.extY = null;
        this.transform = new CMatrix();
        this.localTransform = new CMatrix();
        this.aLabels = [];
        this.chart = chart;
        this.posX = null;
        this.posY = null;
        this.axis = axis;
    }

    CValAxisLabels.prototype.hit = function(x, y) {
        var tx, ty;
        if(this.chart && this.chart.invertTransform) {
            tx = this.chart.invertTransform.TransformPointX(x, y);
            ty = this.chart.invertTransform.TransformPointY(x, y);
            return tx >= this.x && ty >= this.y && tx <= this.x + this.extX && ty <= this.y + this.extY;
        }
        return false;
    };
    CValAxisLabels.prototype.draw = function(g) {
        if(this.chart) {
            //g.SetIntegerGrid(false);
            //g.p_width(70);
            //g.transform3(this.chart.transform, false);
            //g.p_color(0, 0, 0, 255);
            //g._s();
            //g._m(this.x, this.y);
            //g._l(this.x + this.extX, this.y + 0);
            //g._l(this.x + this.extX, this.y + this.extY);
            //g._l(this.x + 0, this.y + this.extY);
            //g._z();
            //g.ds();
            //g.SetIntegerGrid(true);
        }
        for(var i = 0; i < this.aLabels.length; ++i) {
            if(this.aLabels[i])
                this.aLabels[i].draw(g);
        }
    };
    CValAxisLabels.prototype.setPosition = function(x, y) {
        this.x = x;
        this.y = y;
        for(var i = 0; i < this.aLabels.length; ++i) {
            if(this.aLabels[i]) {
                var lbl = this.aLabels[i];
                lbl.setPosition(lbl.relPosX + x, lbl.relPosY + y);
            }
        }
    };
    CValAxisLabels.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, x, y);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        for(var i = 0; i < this.aLabels.length; ++i) {
            if(this.aLabels[i])
                this.aLabels[i].updatePosition(x, y);
        }
    };
    CValAxisLabels.prototype.checkShapeChildTransform = function(t) {
        this.transform = this.localTransform.CreateDublicate();
        global_MatrixTransformer.TranslateAppend(this.transform, this.posX, this.posY);
        this.invertTransform = global_MatrixTransformer.Invert(this.transform);
        for(var i = 0; i < this.aLabels.length; ++i) {
            if(this.aLabels[i])
                this.aLabels[i].checkShapeChildTransform(t);
        }
    };

    function CalcLegendEntry(legend, chart, idx) {
        this.chart = chart;
        this.legend = legend;
        this.idx = idx;
        this.x = null;
        this.y = null;
        this.extX = null;
        this.extY = null;
        this.calcMarkerUnion = null;
        this.txBody = null;
        this.txPr = null;
        this.spPr = new AscFormat.CSpPr();
        this.transform = new CMatrix();
        this.transformText = new CMatrix();

        this.contentWidth = 0;
        this.contentHeight = 0;

        this.localTransform = new CMatrix();
        this.localTransformText = new CMatrix();

        this.localX = null;
        this.localY = null;

        this.recalcInfo =
        {
            recalcStyle: true
        };
    }

    CalcLegendEntry.prototype.updatePosition = CShape.prototype.updatePosition;
    CalcLegendEntry.prototype.getChartSpace = function() {
        return this.chart;
    };
    CalcLegendEntry.prototype.getStyles = CDLbl.prototype.getStyles;
    CalcLegendEntry.prototype.recalculateStyle = CDLbl.prototype.recalculateStyle;
    CalcLegendEntry.prototype.Get_Styles = CDLbl.prototype.Get_Styles;
    CalcLegendEntry.prototype.Get_Theme = CDLbl.prototype.Get_Theme;
    CalcLegendEntry.prototype.Get_ColorMap = CDLbl.prototype.Get_ColorMap;
    CalcLegendEntry.prototype.recalculate = function() {
    };
    CalcLegendEntry.prototype.draw = function(g) {

        CShape.prototype.draw.call(this, g);
        if(this.calcMarkerUnion)
            this.calcMarkerUnion.draw(g);
    };
    CalcLegendEntry.prototype.isEmptyPlaceholder = function() {
        return false;
    };
    CalcLegendEntry.prototype.checkWidhtContent = function() {
        var par = this.txBody.content.Content[0];
        var max_width = 0;
        for(var j = 0; j < par.Lines.length; ++j) {
            if(par.Lines[j].Ranges[0].W > max_width) {
                max_width = par.Lines[j].Ranges[0].W;
            }
        }
        this.contentWidth = max_width;
        this.contentHeight = this.txBody.getSummaryHeight();
    };
    CalcLegendEntry.prototype.hit = function(x, y) {
        var tx, ty, oGeometry;
        if(this.invertTransformText) {
            tx = this.invertTransformText.TransformPointX(x, y);
            ty = this.invertTransformText.TransformPointY(x, y);
            if(tx >= 0 && tx <= this.contentWidth && ty >= 0 && ty <= this.contentHeight) {
                return true;
            }
        }
        if(this.chart) {
            var oCanvasHit = this.chart.getCanvasContext();
            var calcMarkerUnion = this.calcMarkerUnion;
            if(calcMarkerUnion.marker && calcMarkerUnion.marker.invertTransform) {
                tx = calcMarkerUnion.marker.invertTransform.TransformPointX(x, y);
                ty = calcMarkerUnion.marker.invertTransform.TransformPointY(x, y);
                oGeometry = calcMarkerUnion.marker.spPr.geometry;
                if(oGeometry.hitInInnerArea(oCanvasHit, tx, ty) || oGeometry.hitInPath(oCanvasHit, tx, ty)) {
                    return true;
                }
            }
            if(calcMarkerUnion.lineMarker && calcMarkerUnion.lineMarker.invertTransform) {
                tx = calcMarkerUnion.lineMarker.invertTransform.TransformPointX(x, y);
                ty = calcMarkerUnion.lineMarker.invertTransform.TransformPointY(x, y);
                oGeometry = calcMarkerUnion.lineMarker.spPr.geometry;
                if(oGeometry.hitInInnerArea(oCanvasHit, tx, ty) || oGeometry.hitInPath(oCanvasHit, tx, ty)) {
                    return true;
                }
            }
        }
        return false;
    };
    CalcLegendEntry.prototype.getTxPrParaPr = function() {
        if(this.txPr) {
            return this.txPr.getFirstParaParaPr();
        }
        return null;
    };
    CalcLegendEntry.prototype.isForm = function() {
        return false;
    };
    CalcLegendEntry.prototype.GetParaDrawing = function() {
        return null;
    };

    function CompiledMarker() {
        this.spPr = new AscFormat.CSpPr();
        this.x = null;
        this.y = null;
        this.extX = null;
        this.extY = null;
        this.localX = null;
        this.localY = null;
        this.transform = new CMatrix();
        this.localTransform = new CMatrix();
        this.pen = null;
        this.brush = null;
    }

    CompiledMarker.prototype.draw = CShape.prototype.draw;
    CompiledMarker.prototype.check_bounds = CShape.prototype.check_bounds;
    CompiledMarker.prototype.isEmptyPlaceholder = function() {
        return false;
    };
    CompiledMarker.prototype.isForm = function() {
        return false;
    };
    CompiledMarker.prototype.GetParaDrawing = function() {
        return null;
    };

    function CUnionMarker() {
        this.lineMarker = null;
        this.marker = null;
    }

    CUnionMarker.prototype.draw = function(g) {
        this.lineMarker && this.lineMarker.draw(g);
        this.marker && this.marker.draw(g);
    };

    function CreateMarkerGeometryByType(type) {
        var ret = new AscFormat.Geometry();
        var w = 43200, h = 43200;

        function AddRect(geom, w, h) {
            geom.AddPathCommand(1, "0", "0");
            geom.AddPathCommand(2, w + "", "0");
            geom.AddPathCommand(2, w + "", h + "");
            geom.AddPathCommand(2, "0", h + "");
            geom.AddPathCommand(6);
        }

        function AddPlus(geom, w, h) {
            geom.AddPathCommand(0, undefined, "none", undefined, w, h);
            geom.AddPathCommand(1, w / 2 + "", "0");
            geom.AddPathCommand(2, w / 2 + "", h + "");
            geom.AddPathCommand(1, "0", h / 2 + "");
            geom.AddPathCommand(2, w + "", h / 2 + "");
        }

        function AddX(geom, w, h) {
            geom.AddPathCommand(0, undefined, "none", undefined, w, h);
            geom.AddPathCommand(1, "0", "0");
            geom.AddPathCommand(2, w + "", h + "");
            geom.AddPathCommand(1, w + "", "0");
            geom.AddPathCommand(2, "0", h + "");
        }

        switch(type) {
            case SYMBOL_CIRCLE:
            {
                ret.AddPathCommand(0, undefined, undefined, undefined, w, h);
                ret.AddPathCommand(1, "0", h / 2 + "");
                ret.AddPathCommand(3, w / 2 + "", h / 2 + "", "cd2", "cd4");
                ret.AddPathCommand(3, w / 2 + "", h / 2 + "", "_3cd4", "cd4");
                ret.AddPathCommand(3, w / 2 + "", h / 2 + "", "0", "cd4");
                ret.AddPathCommand(3, w / 2 + "", h / 2 + "", "cd4", "cd4");
                ret.AddPathCommand(6);
                break;
            }
            case SYMBOL_DASH:
            case SYMBOL_DOT:
            {
                ret.AddPathCommand(0, undefined, "none", undefined, w, h);
                ret.AddPathCommand(1, type === SYMBOL_DASH ? "0" : w / 2 + "", h / 2 + "");
                ret.AddPathCommand(2, w + "", h / 2 + "");
                break;
            }
            case SYMBOL_DIAMOND:
            {
                ret.AddPathCommand(0, undefined, undefined, undefined, w, h);
                ret.AddPathCommand(1, w / 2 + "", "0");
                ret.AddPathCommand(2, w + "", h / 2 + "");
                ret.AddPathCommand(2, w / 2 + "", h + "");
                ret.AddPathCommand(2, "0", h / 2 + "");
                ret.AddPathCommand(6);
                break;
            }
            case SYMBOL_NONE:
            {
                break;
            }
            case SYMBOL_PICTURE:
            case SYMBOL_SQUARE:
            {
                ret.AddPathCommand(0, undefined, undefined, undefined, w, h);
                AddRect(ret, w, h);
                break;
            }
            case SYMBOL_PLUS:
            {
                /* extrusionOk, fill, stroke, w, h*/
                ret.AddPathCommand(0, undefined, undefined, false, w, h);
                AddRect(ret, w, h);
                ret.AddPathCommand(0, undefined, "none", false, w, h);
                AddPlus(ret, w, h);
                break;
            }
            case SYMBOL_STAR:
            {
                ret.AddPathCommand(0, undefined, undefined, false, w, h);
                AddRect(ret, w, h);
                ret.AddPathCommand(0, undefined, "none", false, w, h);
                AddPlus(ret, w, h);
                AddX(ret, w, h);
                break;
            }
            case SYMBOL_TRIANGLE:
            {
                ret.AddPathCommand(0, undefined, undefined, undefined, w, h);
                ret.AddPathCommand(1, w / 2 + "", "0");
                ret.AddPathCommand(2, w + "", h + "");
                ret.AddPathCommand(2, "0", h + "");
                ret.AddPathCommand(6);
                break;
            }
            case SYMBOL_X:
            {
                ret.AddPathCommand(0, undefined, undefined, false, w, h);
                AddRect(ret, w, h);
                ret.AddPathCommand(0, undefined, "none", false, w, h);
                AddX(ret, w, h);
                break;
            }
        }
        var ret2 = new CompiledMarker();
        ret2.spPr.geometry = ret;
        return ret2;
    }

    function isScatterChartType(nType) {
        return (Asc.c_oAscChartTypeSettings.scatter === nType
        || Asc.c_oAscChartTypeSettings.scatterLine === nType
        || Asc.c_oAscChartTypeSettings.scatterLineMarker === nType
        || Asc.c_oAscChartTypeSettings.scatterMarker === nType
        || Asc.c_oAscChartTypeSettings.scatterNone === nType
        || Asc.c_oAscChartTypeSettings.scatterSmooth === nType
        || Asc.c_oAscChartTypeSettings.scatterSmoothMarker === nType)
    }

    function isStockChartType(nType) {
        return (Asc.c_oAscChartTypeSettings.stock === nType)
    }

    function isComboChartType(nType) {
        return (Asc.c_oAscChartTypeSettings.comboAreaBar === nType
            || Asc.c_oAscChartTypeSettings.comboBarLine === nType
            || Asc.c_oAscChartTypeSettings.comboBarLineSecondary === nType
            || Asc.c_oAscChartTypeSettings.comboCustom === nType)
    }

    function CParseResult() {
        this.error = Asc.c_oAscError.ID.No;
        this.obj = null;
    }

    CParseResult.prototype.setError = function(val) {
        this.error = val;
        if(val !== Asc.c_oAscError.ID.No) {
            this.obj = null;
        }
    };
    CParseResult.prototype.setObject = function(val) {
        this.obj = val;
    };
    CParseResult.prototype.isSuccessful = function() {
        return this.error === Asc.c_oAscError.ID.No && this.obj !== null;
    };
    CParseResult.prototype.getObject = function() {
        return this.obj;
    };
    CParseResult.prototype.getError = function() {
        return this.error;
    };

    function fParseChartFormula(sFormula) {
        var res;
        AscCommonExcel.executeInR1C1Mode(false, function() {
            res = fParseChartFormulaInternal(sFormula);
        });
        return res;
    }
    function fParseChartFormulaExternal(sFormula) {
        var res;
        AscCommonExcel.executeInR1C1Mode(false, function() {
            res = fParseChartFormulaInternal(sFormula);
        });
        if(!Array.isArray(res) || res.length === 0) {
            AscCommonExcel.executeInR1C1Mode(true, function() {
                res = fParseChartFormulaInternal(sFormula);
            });
        }
        return res;
    }
    function fParseChartFormulaInternal(sFormula) {
        if(!(typeof sFormula === "string" && sFormula.length > 0)) {
            return [];
        }
        var oWB = Asc.editor && Asc.editor.wbModel;
        if(!oWB) {
            return [];
        }
        var _sFormula = sFormula;
        if(_sFormula.charAt(0) === '=') {
            _sFormula = _sFormula.slice(1);
        }
        var oWS = oWB.getWorksheet(0);
        if(!oWS) {
            return [];
        }
        var aParsed = AscCommonExcel.getRangeByRef(_sFormula, oWS);
        for(var nParsed = aParsed.length - 1; nParsed > -1; nParsed--) {
            if(!aParsed[nParsed].bbox || !aParsed[nParsed].worksheet) {
                aParsed.splice(nParsed, 1);
            }
        }
        return aParsed;
    }

    function fCreateRef(oBBoxInfo) {
        if(oBBoxInfo) {
            return AscCommon.parserHelp.getEscapeSheetName(oBBoxInfo.worksheet.getName()) + "!" + oBBoxInfo.bbox.getAbsName();
        }
        return null;
    }

    function fParseSingleRow(sVal, fCallback) {
        if(sVal[0] !== "{" || sVal[sVal.length - 1] !== "}") {
            return null;
        }
        var oParser, bResult, result = null;
        oParser = new AscCommonExcel.parserFormula(sVal, null, Asc.editor.wbModel.aWorksheets[0]);
        bResult = oParser.parse(true, true);
        if(bResult && oParser.outStack.length === 1) {
            var oLastElem = oParser.outStack[0];
            if(oLastElem.type === AscCommonExcel.cElementType.array &&
                oLastElem.getRowCount() === 1) {
                var aRow = oLastElem.getRow(0);
                result = fCallback(aRow);
            }
        }
        return result;
    }

    function fParseNumArray(sVal, bForce) {
        return fParseSingleRow(sVal, (function(bForce) {
            return function(aRow) {
                var result = null, oToken, nIndex;
                if(aRow.length > 0) {
                    result = [];
                    for(nIndex = 0; nIndex < aRow.length; ++nIndex) {
                        oToken = aRow[nIndex];
                        if(oToken.type === AscCommonExcel.cElementType.number) {
                            result.push(oToken.toNumber());
                        }
                        else {
                            if(bForce) {
                                result.push(oToken.getValue());
                            }
                            else {
                                return null;
                            }
                        }
                    }
                }
                return result;
            }
        })(bForce));
    }

    function fParseStrArray(sVal) {
        return fParseSingleRow(sVal, (function() {
            return function(aRow) {
                var result = null, oToken, nIndex;
                if(aRow.length > 0) {
                    result = [];
                    for(nIndex = 0; nIndex < aRow.length; ++nIndex) {
                        oToken = aRow[nIndex];
                        result.push(oToken.toString());
                    }
                }
                return result;
            }
        })());
    }

    function fParseNumLit(sVal, bForce, oResult) {
        var result = null;
        var sParsed, aNumbers, nIndex;
        if(typeof sVal === "string" && sVal.length > 0) {
            if(sVal[0] === "=") {
                sParsed = sVal.slice(1);
                aNumbers = fParseNumArray(sParsed, bForce);
            }
            else {
                sParsed = sVal;
                aNumbers = fParseNumArray(sParsed, bForce);
                if(!Array.isArray(aNumbers)) {
                    sParsed = "{" + sParsed + "}";
                    aNumbers = fParseNumArray(sParsed, bForce);
                }
            }
            if(Array.isArray(aNumbers) && aNumbers.length > 0) {
                result = new CNumLit();
                result.setFormatCode("General");
                for(nIndex = 0; nIndex < aNumbers.length; ++nIndex) {
                    result.addNumericPoint(nIndex, aNumbers[nIndex]);
                }
                oResult.setError(Asc.c_oAscError.ID.No);
                oResult.setObject(result);
            }
            else {
                oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
            }
        }
        else {
            oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
        }
    }

    function fParseStrLit(sVal, oResult) {
        var result = null;
        var sParsed, aStr, nIndex;
        if(typeof sVal === "string" && sVal.length > 0) {
            if(sVal[0] === "=") {
                sParsed = sVal.slice(1);
                aStr = fParseStrArray(sParsed);
            }
            else {
                sParsed = sVal;
                aStr = fParseStrArray(sParsed);
                if(!Array.isArray(aStr)) {
                    sParsed = "{" + sParsed + "}";
                    aStr = fParseStrArray(sParsed);
                }
            }
            if(Array.isArray(aStr) && aStr.length > 0) {
                result = new CStrCache();
                for(nIndex = 0; nIndex < aStr.length; ++nIndex) {
                    result.addStringPoint(nIndex, aStr[nIndex]);
                }
                oResult.setError(Asc.c_oAscError.ID.No);
                oResult.setObject(result);
            }
            else {
                oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
            }
        }
        else {
            oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
        }
    }

    function fGetParsedArray(sVal) {
        var aParsed;
        if(sVal[0] === "=") {
            aParsed = sVal.slice(1).split(",");
        }
        else {
            aParsed = sVal.split(",");
        }
        return aParsed;
    }
    function fCheckParseRefsError(aParsed, oResult) {
        var bR1C1;
        for(var nIndex = 0; nIndex < aParsed.length; ++nIndex) {
            var sRef = aParsed[nIndex];
            var oParsedRef = null;
            bR1C1 = false;
            AscCommonExcel.executeInR1C1Mode(false, function() {
                oParsedRef = AscCommon.parserHelp.parse3DRef(sRef);
            });
            if(!oParsedRef) {
                bR1C1 = true;
                AscCommonExcel.executeInR1C1Mode(true, function() {
                    oParsedRef = AscCommon.parserHelp.parse3DRef(sRef);
                });
            }
            if(!oParsedRef) {
                oResult.setError(Asc.c_oAscError.ID.DataRangeError);
                return;
            }
            var oWS = Asc.editor.wbModel.getWorksheetByName(oParsedRef.sheet);
            if(!oWS) {
                oResult.setError(Asc.c_oAscError.ID.InvalidReference);
                return;
            }
            var oRange = null;
            AscCommonExcel.executeInR1C1Mode(bR1C1, function() {
                oRange = oWS.getRange2(oParsedRef.range);
            });
            if(!oRange) {
                oResult.setError(Asc.c_oAscError.ID.DataRangeError);
                return;
            }
        }
        oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
    }

    function fParseNumRef(sVal, bForce, oResult) {
        var result = null, aParsed, nIndex, oWS, oRange, nRow, nCol, oCell;
        if(typeof sVal === "string" && sVal.length > 0) {
            aParsed = fGetParsedArray(sVal);
            if(Array.isArray(aParsed) && aParsed.length > 0) {
                var aRanges = fParseChartFormulaExternal(sVal);
                if(aRanges.length === aParsed.length) {
                    var sFormula;
                    if(aParsed.length > 1) {
                        sFormula = "(";
                    }
                    else {
                        sFormula = "";
                    }
                    for(nIndex = 0; nIndex < aRanges.length; ++nIndex) {
                        oRange = aRanges[nIndex];
                        oWS = oRange.worksheet;
                        if(Math.abs(oRange.bbox.r2 - oRange.bbox.r1) !== 0 && Math.abs(oRange.bbox.c2 - oRange.bbox.c1) !== 0) {
                            oResult.setError(Asc.c_oAscError.ID.NoSingleRowCol);
                            return;
                        }
                        if(bForce === false) {
                            //check strings in cells
                            for(nRow = oRange.bbox.r1; nRow <= oRange.bbox.r2; ++nRow) {
                                for(nCol = oRange.bbox.c1; nCol <= oRange.bbox.c2; ++nCol) {
                                    oCell = oWS.getCell3(nRow, nCol);
                                    if(!CChartDataRefs.prototype.privateCheckCellValueNumberOrEmpty(oCell)) {
                                        oResult.setError(Asc.c_oAscError.ID.DataRangeError);
                                        return;
                                    }
                                }
                            }
                        }
                        if(nIndex > 0) {
                            sFormula += ",";
                        }
                        AscCommonExcel.executeInR1C1Mode(false, function() {
                            sFormula += fCreateRef(oRange);
                        });
                    }
                    if(aParsed.length > 1) {
                        sFormula += ")";
                    }
                    result = new CNumRef();
                    result.setF(sFormula);
                    oResult.setError(Asc.c_oAscError.ID.No);
                    oResult.setObject(result);
                }
                else {
                    fCheckParseRefsError(aParsed, oResult);
                }
            }
            else {
                oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
            }
        }
        else {
            oResult.setError(Asc.c_oAscError.ID.ErrorInFormula);
        }
    }

    function fParseStrRef(sVal, bMultiLvl, oResult) {
        var result = null, aParsed, nIndex, oRange;
        var bMultyRange = false;
        if(typeof sVal === "string" && sVal.length > 0) {
            aParsed = fGetParsedArray(sVal);
            if(Array.isArray(aParsed) && aParsed.length > 0) {
                var aRanges = fParseChartFormulaExternal(sVal);
                if(aRanges.length === aParsed.length) {
                    var sFormula;
                    if(aRanges.length > 1) {
                        sFormula = "(";
                    }
                    else {
                        sFormula = "";
                    }
                    for(nIndex = 0; nIndex < aRanges.length; ++nIndex) {
                        oRange = aRanges[nIndex];
                        if(Math.abs(oRange.bbox.r2 - oRange.bbox.r1) !== 0 && Math.abs(oRange.bbox.c2 - oRange.bbox.c1) !== 0) {
                            if(bMultiLvl !== true) {
                                oResult.setError(Asc.c_oAscError.ID.NoSingleRowCol);
                                return;
                            }
                            bMultyRange = true;
                        }
                        if(nIndex > 0) {
                            sFormula += ",";
                        }
                        AscCommonExcel.executeInR1C1Mode(false, function() {
                            sFormula += fCreateRef(oRange);
                        });
                    }
                    if(aParsed.length > 1) {
                        sFormula += ")";
                    }
                    if(bMultyRange) {
                        result = new CMultiLvlStrRef();
                        result.setF(sFormula);
                    }
                    else {
                        result = new CStrRef();
                        result.setF(sFormula);
                    }
                    oResult.setError(Asc.c_oAscError.ID.No);
                    oResult.setObject(result);
                }
                else {
                    fCheckParseRefsError(aParsed, oResult);
                }
            }
        }
    }

    var SERIES_FLAG_HOR_VALUE = 1;
    var SERIES_FLAG_VERT_VALUE = 2;
    var SERIES_FLAG_CAT = 4;
    var SERIES_FLAG_TX = 8;
    var SERIES_FLAG_CONTINUOUS = 16;

    function CDataRefs(aRefs) {
        this.aRefs = [];
        this.ref = null;
        for(var nRef = 0; nRef < aRefs.length; ++nRef) {
            this.addInternal(aRefs[nRef]);
        }
    }

    CDataRefs.prototype.setRef = function(oRef) {
        this.ref = oRef;
    };
    CDataRefs.prototype.isOneCell = function() {
        if(this.aRefs.length === 1) {
            return this.aRefs[0].bbox.isOneCell();
        }
        return false;
    };
    CDataRefs.prototype.isEmpty = function() {
        return this.aRefs.length === 0;
    };
    CDataRefs.prototype.getEqualWorksheet = function() {
        if(this.isEmpty()) {
            return null;
        }
        var oWorksheet = this.aRefs[0].worksheet;
        for(var nRef = 1; nRef < this.aRefs.length; ++nRef) {
            if(oWorksheet !== this.aRefs[nRef].worksheet) {
                return null;
            }
        }
        return oWorksheet;
    };
    CDataRefs.prototype.isCorrect = function() {
        if(this.isEmpty()) {
            return false;
        }
        if(!this.getEqualWorksheet()) {
            return false;
        }
        return true;
    };
    CDataRefs.prototype.isCorrectForVal = function() {
        if(this.isCorrect() && (this.isInOneCol() || this.isInOneRow())) {
            return true;
        }
        return false;
    };
    CDataRefs.prototype.isContinuous = function() {
        return this.aRefs.length === 1;
    };
    CDataRefs.prototype.getMaxRow = function() {
        if(this.isCorrect()) {
            var aSorted = [].concat(this.aRefs);
            aSorted.sort(function(a, b) {
                return a.bbox.r2 - b.bbox.r2;
            });
            return aSorted[aSorted.length - 1].bbox.r2;
        }
        return null;
    };
    CDataRefs.prototype.getMinRow = function() {
        if(this.isCorrect()) {
            var aSorted = [].concat(this.aRefs);
            aSorted.sort(function(a, b) {
                return a.bbox.r1 - b.bbox.r1;
            });
            return aSorted[0].bbox.r1;
        }
        return null;
    };
    CDataRefs.prototype.getMaxCol = function() {
        if(this.isCorrect()) {
            var aSorted = [].concat(this.aRefs);
            aSorted.sort(function(a, b) {
                return a.bbox.c2 - b.bbox.c2;
            });
            return aSorted[aSorted.length - 1].bbox.c2;
        }
        return null;
    };
    CDataRefs.prototype.getMinCol = function() {
        if(this.isCorrect()) {
            var aSorted = [].concat(this.aRefs);
            aSorted.sort(function(a, b) {
                return a.bbox.c1 - b.bbox.c1;
            });
            return aSorted[0].bbox.c1;
        }
        return null;
    };
    CDataRefs.prototype.isInRow = function() {
        if(!this.isCorrect()) {
            return false;
        }
        var oRef = this.aRefs[0];
        for(var nRef = 1; nRef < this.aRefs.length; ++nRef) {
            if(!oRef.bbox.isEqualRows(this.aRefs[nRef].bbox)) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.isInOneRow = function() {
        if(this.isInRow()) {
            var oRef = this.aRefs[0];
            if(oRef.bbox.r1 === oRef.bbox.r2) {
                return true;
            }
        }
        return false;
    };
    CDataRefs.prototype.isInCol = function() {
        if(!this.isCorrect()) {
            return false;
        }
        var oRef = this.aRefs[0];
        for(var nRef = 1; nRef < this.aRefs.length; ++nRef) {
            if(!oRef.bbox.isEqualCols(this.aRefs[nRef].bbox)) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.isInOneCol = function() {
        if(this.isInCol()) {
            var oRef = this.aRefs[0];
            if(oRef.bbox.c1 === oRef.bbox.c2) {
                return true;
            }
        }
        return false;
    };
    CDataRefs.prototype.isEqualCols = function(oOther) {
        if(this.getEqualWorksheet() !== oOther.getEqualWorksheet()) {
            return false;
        }
        if(this.aRefs.length !== oOther.aRefs.length) {
            return false;
        }
        if(!this.isInRow() || !oOther.isInRow()) {
            return false;
        }
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            if(!this.aRefs[nRef].bbox.isEqualCols(oOther.aRefs[nRef].bbox)) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.isEqualRows = function(oOther) {
        if(this.getEqualWorksheet() !== oOther.getEqualWorksheet()) {
            return false;
        }
        if(this.aRefs.length !== oOther.aRefs.length) {
            return false;
        }
        if(!this.isInCol() || !oOther.isInCol()) {
            return false;
        }
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            if(!this.aRefs[nRef].bbox.isEqualRows(oOther.aRefs[nRef].bbox)) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.isAboveInRows = function(oOther) {
        if(this.isEqualCols(oOther)) {
            return this.aRefs[0].bbox.r2 < oOther.aRefs[0].bbox.r1;
        }
        return false;
    };
    CDataRefs.prototype.isToTheLeftInCols = function(oOther) {
        if(this.isEqualRows(oOther)) {
            return this.aRefs[0].bbox.c2 < oOther.aRefs[0].bbox.c1;
        }
        return false;
    };
    CDataRefs.prototype.isSameCol = function(oOther) {
        var oData = new CDataRefs(this.aRefs.concat(oOther.aRefs));
        if(oData.isInCol() && oData.aRefs[0].bbox.c1 === oData.aRefs[0].bbox.c2) {
            return true;
        }
        return false;
    };
    CDataRefs.prototype.isSameRow = function(oOther) {
        var oData = new CDataRefs(this.aRefs.concat(oOther.aRefs));
        if(oData.isInRow() && oData.aRefs[0].bbox.r1 === oData.aRefs[0].bbox.r2) {
            return true;
        }
        return false;
    };
    CDataRefs.prototype.isToTheLeftInSameRow = function(oOther) {
        if(this.isSameRow(oOther)) {
            if(this.getMaxCol() < oOther.getMinCol()) {
                return true;
            }
        }
        return false;
    };
    CDataRefs.prototype.isAboveInSameCol = function(oOther) {
        if(this.isSameCol(oOther)) {
            if(this.getMaxRow() < oOther.getMinRow()) {
                return true;
            }
        }
        return false;
    };
    CDataRefs.prototype.isEqual = function(oOther) {
        if(this.aRefs.length !== oOther.aRefs.length) {
            return false;
        }
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            var oRef = this.aRefs[nRef];
            var oOtherRef = oOther.aRefs[nRef];
            if(oRef.worksheet !== oOtherRef.worksheet) {
                return false;
            }
            if(!oRef.bbox.isEqual(oOtherRef.bbox)) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.union = function(oOther) {
        var oClone = oOther.clone();
        this.aRefs = this.aRefs.concat(oClone.aRefs);
        this.checkUnion();
    };
    CDataRefs.prototype.add = function(oRef) {
        this.addInternal(oRef);
        this.checkUnion();
    };
    CDataRefs.prototype.addInternal = function(oRef) {
        var oMinRef = oRef.getMinimalCellsRange();
        if(oMinRef) {
            this.aRefs.push(oMinRef);
        }
    };
    CDataRefs.prototype.checkUnion = function() {
        var nRef, nOtherRef, oRef, oOtherRef, oBBox, oOtherBBox;
        do {
            for(nRef = this.aRefs.length - 1; nRef > -1; --nRef) {
                oRef = this.aRefs[nRef];
                oBBox = oRef.bbox;
                for(nOtherRef = nRef - 1; nOtherRef > -1; --nOtherRef) {
                    oOtherRef = this.aRefs[nOtherRef];
                    oOtherBBox = oOtherRef.bbox;
                    if(oRef.worksheet === oOtherRef.worksheet) {
                        if(oBBox.isNeighbor(oOtherBBox)) {
                            oBBox.union2(oOtherBBox);
                            this.aRefs.splice(nOtherRef, 1);
                            break;
                        }
                    }
                }
                if(nOtherRef > -1) {
                    break;
                }
            }
        } while(nRef > -1)
    };
    CDataRefs.prototype.clone = function() {
        var oCopy = new CDataRefs([]);
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            var oRef = this.aRefs[nRef];
            var oCopyRef = oRef.createFromBBox(oRef.worksheet, oRef.bbox.clone(true));
            oCopy.aRefs.push(oCopyRef);
        }
        return oCopy;
    };
    CDataRefs.prototype.intersection = function(oOther) {
        var oIntersectRefs = new CDataRefs([]);
        var nRef, nOtherRef, oRef, oOtherRef, oBBox, oOtherBBox, oIntersectBBox;
        for(nOtherRef = 0; nOtherRef < oOther.aRefs.length; ++nOtherRef) {
            oOtherRef = oOther.aRefs[nOtherRef];
            for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
                oRef = this.aRefs[nRef];
                if(oRef.worksheet === oOtherRef.worksheet) {
                    oBBox = oRef.bbox;
                    oOtherBBox = oOtherRef.bbox;
                    oIntersectBBox = oBBox.intersection(oOtherBBox);
                    if(oIntersectBBox) {
                        oIntersectRefs.add(oRef.createFromBBox(oRef.worksheet, oIntersectBBox));
                    }
                }
            }
        }
        return oIntersectRefs;
    };
    CDataRefs.prototype.getDataRange = function() {
        var sResult = this.getFormulaWithCurrentMode();
        if(sResult.length > 0) {
            sResult = "=" + sResult;
        }
        return sResult;
    };
    CDataRefs.prototype.getFormula = function() {
        var sRes;
        var oThis = this;
        AscCommonExcel.executeInR1C1Mode(false, function() {
            sRes = oThis.getFormulaWithCurrentMode();
        });
        return sRes;
    };
    CDataRefs.prototype.getFormulaWithCurrentMode = function() {
        var sResult = "";
        var sCurRef;
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            sCurRef = fCreateRef(this.aRefs[nRef]);
            if(sCurRef === null) {
                return "";
            }
            sResult += sCurRef;
            if(nRef < this.aRefs.length - 1) {
                sResult += ",";
            }
        }
        return sResult;
    };
    CDataRefs.prototype.getHorRefs = function() {
        var aRet = [];
        var oRef, oRefInRet, nRef, nRefInRet;
        for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
            oRef = this.aRefs[nRef];
            for(nRefInRet = 0; nRefInRet < aRet.length; ++nRefInRet) {
                oRefInRet = aRet[nRefInRet];
                if(oRefInRet.worksheet === oRef.worksheet &&
                    oRefInRet.bbox.isEqualCols(oRef.bbox)) {
                    break;
                }
            }
            if(nRefInRet === aRet.length) {
                aRet.push(oRef);
            }
        }
        aRet.sort(function(a, b) {
            return a.bbox.c1 - b.bbox.c1;
        });
        return aRet;
    };
    CDataRefs.prototype.getVertRefs = function() {
        var aRet = [];
        var oRef, oRefInRet, nRef, nRefInRet;
        for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
            oRef = this.aRefs[nRef];
            for(nRefInRet = 0; nRefInRet < aRet.length; ++nRefInRet) {
                oRefInRet = aRet[nRefInRet];
                if(oRefInRet.worksheet === oRef.worksheet &&
                    oRefInRet.bbox.isEqualRows(oRef.bbox)) {
                    break;
                }
            }
            if(nRefInRet === aRet.length) {
                aRet.push(oRef);
            }
        }
        aRet.sort(function(a, b) {
            return a.bbox.r1 - b.bbox.r1;
        });
        return aRet;
    };
    CDataRefs.prototype.getCellsCount = function() {
        var nCount = 0;
        var oRef, nRef;
        for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
            oRef = this.aRefs[nRef];
            nCount += ((oRef.c2 - oRef.c1 + 1) * (oRef.r2 - oRef.r1 + 1));
        }
        return nCount;
    };
    CDataRefs.prototype.getGrid = function() {
        if(!this.getEqualWorksheet()) {
            return null;
        }
        var nRef, oRef;
        var aHorRefs = this.getHorRefs();
        for(nRef = aHorRefs.length - 1; nRef > 0; --nRef) {
            if(aHorRefs[nRef].bbox.c1 < aHorRefs[nRef - 1].bbox.c2) {
                return null;
            }
        }
        var aVertRefs = this.getVertRefs();
        for(nRef = aVertRefs.length - 1; nRef > 0; --nRef) {
            if(aVertRefs[nRef].bbox.r1 < aVertRefs[nRef - 1].bbox.r2) {
                return null;
            }
        }
        var nR1, nR2, nC1, nC2;
        var nCount;
        var nHorRef, nVertRef;
        var oHorBBox, oVertBBox;
        var aRows = [], aRow;
        for(nVertRef = 0; nVertRef < aVertRefs.length; ++nVertRef) {
            oVertBBox = aVertRefs[nVertRef].bbox;
            nR1 = oVertBBox.r1;
            nR2 = oVertBBox.r2;
            aRow = [];
            for(nHorRef = 0; nHorRef < aHorRefs.length; ++nHorRef) {
                oHorBBox = aHorRefs[nHorRef].bbox;
                nC1 = oHorBBox.c1;
                nC2 = oHorBBox.c2;
                nCount = 0;
                for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
                    oRef = this.aRefs[nRef];
                    if(oRef.bbox.r1 === nR1 && oRef.bbox.r2 === nR2 &&
                        oRef.bbox.c1 === nC1 && oRef.bbox.c2 === nC2) {
                        if(nCount === 0) {
                            ++nCount;
                            aRow.push(oRef);
                        }
                        else {
                            return null;
                        }
                    }
                }
                if(nCount !== 1) {
                    return null;
                }
            }
            aRows.push(aRow);
        }
        return aRows;
    };
    CDataRefs.prototype.hasIntersection = function(oRange) {
        for(var nRange = 0; nRange < this.aRefs.length; ++nRange) {
            if(this.aRefs[nRange].isIntersect(oRange)) {
                return true;
            }
        }
        return false;
    };
    CDataRefs.prototype.isInside = function(oRange) {
        if(this.aRefs.length === 0) {
            return false;
        }
        for(var nRange = 0; nRange < this.aRefs.length; ++nRange) {
            if(!oRange.containsRange(this.aRefs[nRange])) {
                return false;
            }
        }
        return true;
    };
    CDataRefs.prototype.clear = function() {
        this.aRefs.length = 0;
    };
    CDataRefs.prototype.collectBoundsByWS = function(oBounds) {
        var oRef, oBBox;
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            oRef = this.aRefs[nRef];
            oBBox = oRef.bbox;
            var sWSName = oRef.worksheet.getName();
            var oCurBounds = oBounds[sWSName];
            if(oCurBounds) {
                oCurBounds.bbox.union2(oBBox);
            }
            else {
                oBounds[sWSName] = new AscCommonExcel.Range(oRef.worksheet, oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2);
            }
        }
    };
    CDataRefs.prototype.move = function(oRangeFrom, oRangeTo) {
        var nDiffR = oRangeTo.bbox.r1 - oRangeFrom.bbox.r1;
        var nDiffC = oRangeTo.bbox.c1 - oRangeFrom.bbox.c1;
        for(var nRef = 0; nRef < this.aRefs.length; ++nRef) {
            var oRef = this.aRefs[nRef];
            if(oRef.worksheet === oRangeFrom.worksheet) {
                oRef.bbox.r1 += nDiffR;
                oRef.bbox.r2 += nDiffR;
                oRef.bbox.c1 += nDiffC;
                oRef.bbox.c2 += nDiffC;
            }
        }
    };
    CDataRefs.prototype.getCellsCount = function() {
        var nCount = 0;
        var oRef, nRef, oBBox;
        for(nRef = 0; nRef < this.aRefs.length; ++nRef) {
            oRef = this.aRefs[nRef];
            oBBox = oRef.bbox;
            nCount += (oBBox.r2 - oBBox.r1 + 1) * (oBBox.c2 - oBBox.c1 + 1);
        }
        return nCount;
    };
    CDataRefs.prototype.collectRefsInsideRange = function(oRange, aRefs) {
        if(this.ref) {
            if(this.isInside(oRange)) {
                aRefs.push(this.ref);
            }
        }
    };
    CDataRefs.prototype.collectIntersectionRefs = function(aRanges, aCollectedRefs) {
        if(this.ref) {
            for(var nRange = 0; nRange < aRanges.length; ++nRange) {
                if(this.hasIntersection(aRanges[nRange])) {
                    aCollectedRefs.push(this.ref);
                    break;
                }
            }
        }
    };

    var SERIES_COMPARE_RESULT_NONE = 0;
    var SERIES_COMPARE_RESULT_RIGHT = 1;
    var SERIES_COMPARE_RESULT_LEFT = 2;
    var SERIES_COMPARE_RESULT_ABOVE = 3;
    var SERIES_COMPARE_RESULT_BELOW = 4;

    function CSeriesDataRefs(oSeries) {
        if(oSeries) {
            this.val = oSeries.getValDataRefs();
            this.cat = oSeries.getCatDataRefs();
            this.tx = oSeries.getTxDataRefs();
            this.errBarsMinus = oSeries.getErrBarsMinusDataRefs();
            this.errBarsPlus = oSeries.getErrBarsPlusDataRefs();
        }
        else {
            this.val = new CDataRefs([]);
            this.cat = new CDataRefs([]);
            this.tx = new CDataRefs([]);
            this.errBarsMinus = new CDataRefs([]);
            this.errBarsPlus = new CDataRefs([]);
        }
        this.series = oSeries;
    }

    CSeriesDataRefs.prototype.getInfo = function() {
        var nInfo = 0;
        if(!this.val.isCorrectForVal()) {
            return nInfo;
        }
        if(this.tx.isEmpty() && this.cat.isEmpty()) {
            if(this.val.isInOneRow()) {
                nInfo |= SERIES_FLAG_HOR_VALUE;
                if(this.val.isContinuous()) {
                    nInfo |= SERIES_FLAG_CONTINUOUS;
                }
            }
            if(this.val.isInOneCol()) {
                nInfo |= SERIES_FLAG_VERT_VALUE;
                if(this.val.isContinuous()) {
                    nInfo |= SERIES_FLAG_CONTINUOUS;
                }
            }
        }
        else if(this.tx.isEmpty() && !this.cat.isEmpty()) {
            if(this.cat.isAboveInRows(this.val)) {
                nInfo |= SERIES_FLAG_HOR_VALUE;
                nInfo |= SERIES_FLAG_CAT;
                if(this.val.isContinuous()) {
                    nInfo |= SERIES_FLAG_CONTINUOUS;
                }
            }
            else if(this.cat.isToTheLeftInCols(this.val)) {
                nInfo |= SERIES_FLAG_VERT_VALUE;
                nInfo |= SERIES_FLAG_CAT;
                if(this.val.isContinuous()) {
                    nInfo |= SERIES_FLAG_CONTINUOUS;
                }
            }
        }
        else if(!this.tx.isEmpty() && this.cat.isEmpty()) {
            if(this.tx.isContinuous()) {
                if(this.tx.isAboveInSameCol(this.val)) {
                    nInfo |= SERIES_FLAG_VERT_VALUE;
                    nInfo |= SERIES_FLAG_TX;
                    if(this.val.isContinuous()) {
                        nInfo |= SERIES_FLAG_CONTINUOUS;
                    }
                }
                else if(this.tx.isToTheLeftInSameRow(this.val)) {
                    nInfo |= SERIES_FLAG_HOR_VALUE;
                    nInfo |= SERIES_FLAG_TX;
                    if(this.val.isContinuous()) {
                        nInfo |= SERIES_FLAG_CONTINUOUS;
                    }
                }
            }
        }
        else if(!this.tx.isEmpty() && !this.cat.isEmpty()) {
            if(this.tx.isContinuous()) {
                if(this.cat.isAboveInRows(this.val) && this.tx.isToTheLeftInSameRow(this.val)) {
                    nInfo |= SERIES_FLAG_HOR_VALUE;
                    nInfo |= SERIES_FLAG_CAT;
                    nInfo |= SERIES_FLAG_TX;
                    if(this.val.isContinuous()) {
                        nInfo |= SERIES_FLAG_CONTINUOUS;
                    }
                }
                else if(this.cat.isToTheLeftInCols(this.val) && this.tx.isAboveInSameCol(this.val)) {
                    nInfo |= SERIES_FLAG_VERT_VALUE;
                    nInfo |= SERIES_FLAG_CAT;
                    nInfo |= SERIES_FLAG_TX;
                    if(this.val.isContinuous()) {
                        nInfo |= SERIES_FLAG_CONTINUOUS;
                    }
                }
            }
        }
        return nInfo;
    };
    CSeriesDataRefs.prototype.compare = function(oOther) {
        var nInfo = this.getInfo();
        if(nInfo === 0) {
            return SERIES_COMPARE_RESULT_NONE;
        }
        var nOtherInfo = oOther.getInfo();
        if((nInfo & nOtherInfo) !== nInfo) {
            return SERIES_COMPARE_RESULT_NONE;
        }
        if(!this.cat.isEqual(oOther.cat)) {
            return SERIES_COMPARE_RESULT_NONE;
        }
        return this.compareValAndTx(oOther);
    };
    CSeriesDataRefs.prototype.compareValAndTx = function(oOther) {
        var nInfo = this.getInfo();
        if(this.val.isAboveInRows(oOther.val)) {
            if((nInfo & SERIES_FLAG_TX) && !this.tx.isAboveInRows(oOther.tx)) {
                return SERIES_COMPARE_RESULT_NONE;
            }
            return SERIES_COMPARE_RESULT_ABOVE;
        }
        else if(oOther.val.isAboveInRows(this.val)) {
            if((nInfo & SERIES_FLAG_TX) && !oOther.tx.isAboveInRows(this.tx)) {
                return SERIES_COMPARE_RESULT_NONE;
            }
            return SERIES_COMPARE_RESULT_BELOW;
        }
        else if(this.val.isToTheLeftInCols(oOther.val)) {
            if((nInfo & SERIES_FLAG_TX) && !this.tx.isToTheLeftInCols(oOther.tx)) {
                return SERIES_COMPARE_RESULT_NONE;
            }
            return SERIES_COMPARE_RESULT_LEFT;
        }
        else if(oOther.val.isToTheLeftInCols(this.val)) {
            if((nInfo & SERIES_FLAG_TX) && !oOther.tx.isToTheLeftInCols(this.tx)) {
                return SERIES_COMPARE_RESULT_NONE;
            }
            return SERIES_COMPARE_RESULT_RIGHT;
        }
        return SERIES_COMPARE_RESULT_NONE
    };
    CSeriesDataRefs.prototype.fillSelectedRanges = function(oWSView) {
        fFillSelectedRanges(this.val, this.cat, this.tx, this.getInfo(), true, oWSView);
    };
    CSeriesDataRefs.prototype.fillFromSelectedRange = function(oSelectedRange) {
        fFillDataFromSelectedRange(this, oSelectedRange);
    };
    CSeriesDataRefs.prototype.hasIntersection = function(oRange) {
        if(this.val.hasIntersection(oRange)) {
            return true;
        }
        if(this.cat.hasIntersection(oRange)) {
            return true;
        }
        if(this.tx.hasIntersection(oRange)) {
            return true;
        }
        if(this.errBarsMinus.hasIntersection(oRange)) {
            return true;
        }
        if(this.errBarsPlus.hasIntersection(oRange)) {
            return true;
        }
        return false;
    };
    CSeriesDataRefs.prototype.collectBoundsByWS = function(oBounds) {
        this.val.collectBoundsByWS(oBounds);
        this.cat.collectBoundsByWS(oBounds);
        this.tx.collectBoundsByWS(oBounds);
        this.errBarsMinus.collectBoundsByWS(oBounds);
        this.errBarsPlus.collectBoundsByWS(oBounds);
    };
    CSeriesDataRefs.prototype.collectRefsInsideRange = function(oRange, aRefs) {
        if(!this.series) {
            return;
        }
        if(!this.hasIntersection(oRange)) {
            return;
        }
        this.val.collectRefsInsideRange(oRange, aRefs);
        this.cat.collectRefsInsideRange(oRange, aRefs);
        this.tx.collectRefsInsideRange(oRange, aRefs);
        this.errBarsMinus.collectRefsInsideRange(oRange, aRefs);
        this.errBarsPlus.collectRefsInsideRange(oRange, aRefs);
    };
    CSeriesDataRefs.prototype.collectIntersectionRefs = function(aRanges, aCollectedRefs) {
        this.val.collectIntersectionRefs(aRanges, aCollectedRefs);
        this.cat.collectIntersectionRefs(aRanges, aCollectedRefs);
        this.tx.collectIntersectionRefs(aRanges, aCollectedRefs);
        this.errBarsMinus.collectIntersectionRefs(aRanges, aCollectedRefs);
        this.errBarsPlus.collectIntersectionRefs(aRanges, aCollectedRefs);
    };
    CSeriesDataRefs.prototype.getAscSeries = function() {
        var oSeria = new AscFormat.asc_CChartSeria();
        oSeria.Val.Formula = this.val.getFormula();
        oSeria.Val.NumCache = [];
        oSeria.TxCache.Formula = this.tx.getFormula();
        oSeria.TxCache.NumCache = [];
        var oCat = {Formula: this.cat.getFormula(), NumCache: []};
        oSeria.Cat = oCat;
        oSeria.xVal = oCat;
        return oSeria;
    };
    CSeriesDataRefs.prototype.getValCellsCount = function() {
        return this.val.getCellsCount();
    };

    function fFillDataFromSelectedRange(oData, oSelectedRange) {
        var ranges = oSelectedRange.ranges;
        for(var i = 0; i < ranges.length; ++i) {
            if(ranges[i].chartRangeIndex === 0) {
                if(oData.val.isContinuous()) {
                    oData.val.aRefs[0].bbox.assign2(ranges[i]);
                }
            }
            else if(ranges[i].chartRangeIndex === 1) {
                if(oData.tx.isContinuous()) {
                    oData.tx.aRefs[0].bbox.assign2(ranges[i]);
                }
            }
            else if(ranges[i].chartRangeIndex === 2) {
                if(oData.cat.isContinuous()) {
                    oData.cat.aRefs[0].bbox.assign2(ranges[i]);
                }
            }
        }
    }

    function fFillSelectedRanges(oVal, oCat, oTx, nInfo, bSeparated, oWSView) {
        if(!oVal) {
            return;
        }
        if(!oVal.isContinuous()) {
            return;
        }
        if(!oTx.isEmpty() && !oTx.isContinuous()) {
            return;
        }
        if(!oCat.isEmpty() && !oCat.isContinuous()) {
            return;
        }
        var oRef = oVal.aRefs[0];
        if(oRef.worksheet !== oWSView.model) {
            return;
        }
        oWSView.isChartAreaEditMode = true;
        var oSelectionRange = new AscCommonExcel.SelectionRange(oWSView);
        oSelectionRange.ranges = [];
        var oBBox, oRange;
        oBBox = oVal.aRefs[0].bbox;
        oSelectionRange.addRange();
        oRange = oSelectionRange.getLast();
        oRange.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2, true);
        oRange.separated = bSeparated;
        oRange.chartRangeIndex = 0;
        oRange.vert = (nInfo & SERIES_FLAG_HOR_VALUE);
        if(!oTx.isEmpty()) {
            oBBox = oTx.aRefs[0].bbox;
            oSelectionRange.addRange();
            oRange = oSelectionRange.getLast();
            oRange.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2, true);
            oRange.separated = bSeparated;
            oRange.chartRangeIndex = 1;
        }
        if(!oCat.isEmpty()) {
            oBBox = oCat.aRefs[0].bbox;
            oSelectionRange.addRange();
            oRange = oSelectionRange.getLast();
            oRange.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2, true);
            oRange.separated = bSeparated;
            oRange.chartRangeIndex = 2;
        }
        oWSView.oOtherRanges = oSelectionRange;
    }

    function CChartDataRefs(oChartSpace) {
        this.chartSpace = oChartSpace;
        this.val = new CDataRefs([]);
        this.cat = new CDataRefs([]);
        this.tx = new CDataRefs([]);
        this.info = 0;
        this.seriesRefs = [];
        this.boundsByWS = {};
        this.labelsRefs = [];
        this.updateDataRefs();
    }

    CChartDataRefs.prototype.updateDataRefs = function() {
        this.val.clear();
        this.cat.clear();
        this.tx.clear();
        this.info = 0;
        this.seriesRefs.length = 0;
        this.labelsRefs.length = 0;
        this.boundsByWS = {};
        if(!this.chartSpace) {
            return;
        }
        var aSeries = this.chartSpace.getAllSeries();
        aSeries.sort(function(a, b) {
            return a.order - b.order;
        });
        var nSeries;
        var oSeriesRefs;
        var nStartIdx = aSeries.length;

        for(nSeries = 0; nSeries < aSeries.length; ++nSeries) {
            this.seriesRefs.push(new CSeriesDataRefs(aSeries[nSeries]));
        }
        for(nSeries = 0; nSeries < this.seriesRefs.length; ++nSeries) {
            oSeriesRefs = this.seriesRefs[nSeries];
            if(oSeriesRefs.getInfo() !== 0) {
                nStartIdx = nSeries;
                break;
            }
        }
        if(nStartIdx < this.seriesRefs.length) {
            var nFirstInfo = oSeriesRefs.getInfo();
            var bAlign = false;
            if(nFirstInfo & SERIES_FLAG_HOR_VALUE) {
                if(this.checkSeries(nStartIdx, SERIES_COMPARE_RESULT_ABOVE)) {
                    bAlign = true;
                }
            }
            if(!bAlign) {
                if(nFirstInfo & SERIES_FLAG_VERT_VALUE) {
                    if(this.checkSeries(nStartIdx, SERIES_COMPARE_RESULT_LEFT)) {
                        bAlign = true;
                    }
                }
            }
        }
        this.checkLabels();
        this.checkBoundsByWs();
    };
    CChartDataRefs.prototype.checkSeries = function(nStartIdx, nCompareResult) {
        var oFirstSeriesRefs = this.seriesRefs[nStartIdx];
        var oValRefs = oFirstSeriesRefs.val.clone();
        var oTxRefs = oFirstSeriesRefs.tx.clone();
        var oCatRefs = oFirstSeriesRefs.cat.clone();
        var oPrevRefs = oFirstSeriesRefs;
        var oSeriesRefs, nSeries;
        var bFirstCatIsEmpty = oCatRefs.isEmpty();
        for(nSeries = nStartIdx + 1; nSeries < this.seriesRefs.length; ++nSeries) {
            oSeriesRefs = this.seriesRefs[nSeries];
            if(bFirstCatIsEmpty) {
                if(oPrevRefs.compare(oSeriesRefs) !== nCompareResult) {
                    break;
                }
            }
            else {
                if(oSeriesRefs.cat.isEmpty()) {
                    if(oPrevRefs.compareValAndTx(oSeriesRefs) !== nCompareResult) {
                        break;
                    }
                }
                else {
                    if(oPrevRefs.compare(oSeriesRefs) !== nCompareResult) {
                        break;
                    }
                }
            }
            oValRefs.union(oSeriesRefs.val);
            oTxRefs.union(oSeriesRefs.tx);
            oPrevRefs = oSeriesRefs;
        }
        if(nSeries === this.seriesRefs.length) {
            this.val = oValRefs;
            this.tx = oTxRefs;
            this.cat = oCatRefs;
            this.info = oFirstSeriesRefs.getInfo();
            if((this.info & SERIES_FLAG_HOR_VALUE) &&
                (this.info & SERIES_FLAG_VERT_VALUE)) {
                if(nStartIdx < this.seriesRefs.length - 1) {
                    if(nCompareResult === SERIES_COMPARE_RESULT_ABOVE) {
                        this.info = (this.info & ~(SERIES_FLAG_VERT_VALUE));
                    }
                    else if(nCompareResult === SERIES_COMPARE_RESULT_LEFT) {
                        this.info = (this.info & ~(SERIES_FLAG_HOR_VALUE));
                    }
                }
            }
            return true;
        }
        return false;
    };
    CChartDataRefs.prototype.checkLabels = function() {
        this.labelsRefs.length = 0;
        if(!this.chartSpace) {
            return;
        }
        var oChart = this.chartSpace.chart;
        if(!oChart) {
            return;
        }
        var oThis = this;
        oChart.traverse(function(oElement) {
            if(oElement.getObjectType() === AscDFH.historyitem_type_ChartText) {
                if(oElement.strRef) {
                    oThis.labelsRefs.push(oElement.strRef.getDataRefs());
                }
            }
        });

    };
    CChartDataRefs.prototype.checkBoundsByWs = function() {
        if(this.info !== 0) {
            this.val.collectBoundsByWS(this.boundsByWS);
            this.cat.collectBoundsByWS(this.boundsByWS);
            this.tx.collectBoundsByWS(this.boundsByWS);
        }
        else {
            for(var nSeries = 0; nSeries < this.seriesRefs.length; ++nSeries) {
                this.seriesRefs[nSeries].collectBoundsByWS(this.boundsByWS);
            }
        }
        for(var nLblRef = 0; nLblRef < this.labelsRefs.length; ++nLblRef) {
            this.labelsRefs[nLblRef].collectBoundsByWS(this.boundsByWS);
        }
    };
    CChartDataRefs.prototype.calculateUnionRefs = function() {
        if(this.info === 0) {
            return null;
        }
        var oTxCatIntersection = new CDataRefs([]);
        var oTx = this.tx;
        var oCat = this.cat;
        if(!oTx.isEmpty() && !oCat.isEmpty()) {
            var oTxRange = oTx.aRefs[0];
            var oTxBBox = oTxRange.bbox;
            var oCatBBox = oCat.aRefs[0].bbox;
            var oIntersectionBBox;
            if(this.info & SERIES_FLAG_HOR_VALUE) {
                oIntersectionBBox = new Asc.Range(oTxBBox.c1, oCatBBox.r1, oTxBBox.c2, oCatBBox.r2, true);
            }
            else if(this.info & SERIES_FLAG_VERT_VALUE) {
                oIntersectionBBox = new Asc.Range(oCatBBox.c1, oTxBBox.r1, oCatBBox.c2, oTxBBox.r2, true);
            }
            if(oIntersectionBBox) {
                oTxCatIntersection.add(oTx.aRefs[0].createFromBBox(oTxRange.worksheet, oIntersectionBBox));
            }
        }
        var oResult = new CDataRefs([]);
        oResult.union(oTxCatIntersection);
        oResult.union(oCat);
        oResult.union(oTx);
        oResult.union(this.val);
        return oResult;
    };
    CChartDataRefs.prototype.getUnionRefs = function() {
        this.updateDataRefs();
        return this.calculateUnionRefs();
    };
    CChartDataRefs.prototype.getRange = function() {
        var oResult = this.getUnionRefs();
        if(oResult) {
            return oResult.getDataRange();
        }
        return "";
    };
    CChartDataRefs.prototype.getInfo = function() {
        this.updateDataRefs();
        return this.info;
    };
    CChartDataRefs.prototype.getSeriesRefsFromUnionRefs = function(aRefs, bHorValue, bScatter) {
        if(aRefs.length === 0) {
            return [];
        }
        var oRefs = new CDataRefs(aRefs);
        if(oRefs.isEmpty()) {
            return [];
        }
        var aGrid = oRefs.getGrid();
        var aGridRow, oRef, nRef, oBBox;
        var nRow, nCol;
        var oCell;
        var nEndCol, nEndRow;
        var nTopHeader = -1, nLeftHeader = -1;
        var nRowsCount, nColsCount;
        var nGridRow, nGridCol;
        var nR1, nR2, nC1, nC2;
        var bHorizontalValues;

        var oVal = null, oTx = null, oCat = null, nInfo = 0;
        if(Array.isArray(aGrid)) {
            aGridRow = aGrid[0];
            if(!Array.isArray(aGridRow)) {
                return [];
            }
            oRef = aGridRow[0];
            if(!oRef) {
                return [];
            }
            //try to find headers
            oBBox = oRef.bbox;
            //check empty range in the left top corner
            nRow = oBBox.r1;
            for(nCol = oBBox.c1; nCol <= oBBox.c2; ++nCol) {
                oCell = oRef.worksheet.getCell3(nRow, nCol);
                if(!oCell.isEmptyTextString()) {
                    break;
                }
            }
            nEndCol = nCol - 1;
            if(nEndCol >= oBBox.c1) {
                for(nRow = oBBox.r1 + 1; nRow <= oBBox.r2; ++nRow) {
                    for(nCol = oBBox.c1; nCol <= nEndCol; ++nCol) {
                        oCell = oRef.worksheet.getCell3(nRow, nCol);
                        if(!oCell.isEmptyTextString()) {
                            break;
                        }
                    }
                    if(nCol <= nEndCol) {
                        break;
                    }
                }
                nEndRow = nRow - 1;
                if(nEndCol === oBBox.c2) {
                    if(aGridRow.length === 1) {
                        if(oBBox.c1 === oBBox.c2) {
                            nEndCol = oBBox.c1 - 1;
                        }
                        else {
                            nEndCol = oBBox.c1;
                        }
                    }
                }
                if(nEndRow === oBBox.r2) {
                    if(aGrid.length === 1) {
                        if(oBBox.r1 === oBBox.r2) {
                            nEndRow = oBBox.r1 - 1;
                        }
                        else {
                            nEndRow = oBBox.r1;
                        }
                    }
                }
                nTopHeader = nEndRow - oBBox.r1;
                nLeftHeader = nEndCol - oBBox.c1;
            }
            if(nTopHeader < 0 && nLeftHeader < 0) {
                //consider the row/col as the title
                //if the last cell in the row/col contains a non empty non numeric value
                aGridRow = aGrid[0];
                oRef = aGridRow[aGridRow.length - 1];
                oBBox = oRef.bbox;
                nCol = oBBox.c2;

                var nStartRow = oBBox.r2;
                if(aGridRow.length === 1) {
                    for(nRow = oBBox.r2 - 1; nRow >= oBBox.r1; --nRow) {
                        if(!this.privateCheckCellDateTimeFormatFull(oRef.worksheet.getCell3(nRow, nCol))) {
                            break;
                        }
                    }
                    nStartRow = nRow;
                }
                for(nRow = nStartRow; nRow >= oBBox.r1; --nRow) {
                    if(this.privateCheckCellValueForHeader(oRef.worksheet.getCell3(nRow, nCol))) {
                        break;
                    }
                }
                nTopHeader = nRow - oBBox.r1;
                if(nRow === oBBox.r2) {
                    if(aGridRow.length === 1) {
                        nTopHeader -= 1;
                    }
                }

                aGridRow = aGrid[aGrid.length - 1];
                oRef = aGridRow[0];
                oBBox = oRef.bbox;
                nRow = oBBox.r2;
                var nStartCol = oBBox.c2;
                if(aGrid.length === 1) {
                    for(nCol = oBBox.c2 - 1; nCol >= oBBox.c1; --nCol) {
                        if(!this.privateCheckCellDateTimeFormatFull(oRef.worksheet.getCell3(nRow, nCol))) {
                            break;
                        }
                    }
                    nStartCol = nCol;
                }
                for(nCol = nStartCol; nCol >= oBBox.c1; --nCol) {
                    if(this.privateCheckCellValueForHeader(oRef.worksheet.getCell3(nRow, nCol))) {
                        break;
                    }
                }
                nLeftHeader = nCol - oBBox.c1;
                if(nCol === oBBox.r2) {
                    if(aGrid.length === 1) {
                        nLeftHeader -= 1;
                    }
                }
            }
            bHorizontalValues = bHorValue;
            nRowsCount = 0;
            nColsCount = 0;
            for(nRow = 0; nRow < aGrid.length; ++nRow) {
                aGridRow = aGrid[nRow];
                oRef = aGridRow[0];
                nRowsCount += (oRef.bbox.getHeight());
            }
            aGridRow = aGrid[0];
            for(nCol = 0; nCol < aGridRow.length; ++nCol) {
                oRef = aGridRow[nCol];
                nColsCount += (oRef.bbox.getWidth());
            }
            nRowsCount -= (nTopHeader + 1);
            nColsCount -= (nLeftHeader + 1);
            if(bHorizontalValues !== true && bHorizontalValues !== false) {
                if(nRowsCount > nColsCount) {
                    bHorizontalValues = false;
                }
                else {
                    bHorizontalValues = true;
                }
            }
            if(bScatter) {
                if(bHorizontalValues) {
                    if(nTopHeader === -1 && nRowsCount > 1) {
                        nTopHeader = 0;
                    }
                }
                else {
                    if(nLeftHeader === -1 && nColsCount > 1) {
                        nLeftHeader = 0;
                    }
                }
            }
            oVal = new CDataRefs([]);
            oCat = new CDataRefs([]);
            oTx = new CDataRefs([]);
            if(nTopHeader > -1) {
                if(bHorizontalValues) {
                    nInfo |= SERIES_FLAG_CAT;
                }
                else {
                    nInfo |= SERIES_FLAG_TX;
                }
                aGridRow = aGrid[0];
                for(nRef = 0; nRef < aGridRow.length; ++nRef) {
                    oRef = aGridRow[nRef];
                    oBBox = oRef.bbox;
                    nR1 = oBBox.r1;
                    nR2 = oBBox.r1 + nTopHeader;
                    nC2 = oBBox.c2;
                    if(nRef === 0) {
                        nC1 = oBBox.c1 + nLeftHeader + 1;
                    }
                    else {
                        nC1 = oBBox.c1;
                    }
                    if(nC1 <= nC2) {
                        if(bHorizontalValues) {
                            oCat.add(oRef.createFromBBox(oRef.worksheet, new Asc.Range(nC1, nR1, nC2, nR2, true)));
                        }
                        else {
                            oTx.add(oRef.createFromBBox(oRef.worksheet, new Asc.Range(nC1, nR1, nC2, nR2, true)));
                        }
                    }
                }
            }
            if(nLeftHeader > -1) {
                if(bHorizontalValues) {
                    nInfo |= SERIES_FLAG_TX;
                }
                else {
                    nInfo |= SERIES_FLAG_CAT;
                }
                for(nGridRow = 0; nGridRow < aGrid.length; ++nGridRow) {
                    aGridRow = aGrid[nGridRow];
                    oRef = aGridRow[0];
                    oBBox = oRef.bbox;
                    nC1 = oBBox.c1;
                    nC2 = nC1 + nLeftHeader;
                    nR2 = oBBox.r2;
                    if(nGridRow === 0) {
                        nR1 = oBBox.r1 + nTopHeader + 1;
                    }
                    else {
                        nR1 = oBBox.r1;
                    }
                    if(nR1 <= nR2) {
                        if(bHorizontalValues) {
                            oTx.add(oRef.createFromBBox(oRef.worksheet, new Asc.Range(nC1, nR1, nC2, nR2, true)));
                        }
                        else {
                            oCat.add(oRef.createFromBBox(oRef.worksheet, new Asc.Range(nC1, nR1, nC2, nR2, true)));
                        }
                    }
                }
            }
            if(bHorizontalValues) {
                nInfo |= SERIES_FLAG_HOR_VALUE;
            }
            else {
                nInfo |= SERIES_FLAG_VERT_VALUE;
            }
            for(nGridRow = 0; nGridRow < aGrid.length; ++nGridRow) {
                aGridRow = aGrid[nGridRow];
                for(nGridCol = 0; nGridCol < aGridRow.length; ++nGridCol) {
                    oRef = aGridRow[nGridCol];
                    oBBox = oRef.bbox;
                    if(nGridRow === 0) {
                        nR1 = oBBox.r1 + nTopHeader + 1;
                    }
                    else {
                        nR1 = oBBox.r1;
                    }
                    nR2 = oBBox.r2;
                    if(nGridCol === 0) {
                        nC1 = oBBox.c1 + nLeftHeader + 1;
                    }
                    else {
                        nC1 = oBBox.c1;
                    }
                    nC2 = oBBox.c2;
                    if(nC1 <= nC2 && nR1 <= nR2) {
                        oVal.add(oRef.createFromBBox(oRef.worksheet, new Asc.Range(nC1, nR1, nC2, nR2, true)));
                    }
                }
            }
            var oData = new CChartDataRefs(null);
            oData.val = oVal;
            oData.tx = oTx;
            oData.cat = oCat;
            oData.info = nInfo;
            return oData.getSeriesFromFixedRefs();
        }
        else {
            var aSeries = [];
            bHorizontalValues = bHorValue;
            if(bHorizontalValues !== true && bHorizontalValues !== false) {
                oBBox = aRefs[0].bbox;
                nRowsCount = oBBox.getHeight();
                nColsCount = oBBox.getWidth();
                if(nRowsCount > nColsCount) {
                    bHorizontalValues = false;
                }
                else {
                    bHorizontalValues = true;
                }
            }
            for(nRef = 0; nRef < aRefs.length; ++nRef) {
                aSeries = aSeries.concat(this.getSeriesRefsFromUnionRefs([aRefs[nRef]], bHorizontalValues, false));
            }
            return aSeries;
        }
    };
    CChartDataRefs.prototype.getSeriesFromFixedRefs = function() {
        if(this.info === 0) {
            return null;
        }
        var aData = null;
        var aValHorRefs = this.val.getHorRefs();
        var aValVertRefs = this.val.getVertRefs();
        var aTxVertRefs, aTxHorRefs;
        var aCatVertRefs, aCatHorRefs;
        var nHorRef, nVertRef, nCol, nRow, nTxRef, nCatRef;
        var oValHorRef, oValVertRef, oCatRef, oTxRef;
        var oSeriesData;
        var oBBox;
        if(this.info & AscFormat.SERIES_FLAG_HOR_VALUE) {
            aData = [];
            aCatVertRefs = this.cat.getVertRefs();
            aTxHorRefs = this.tx.getHorRefs();
            for(nVertRef = 0; nVertRef < aValVertRefs.length; ++nVertRef) {
                oValVertRef = aValVertRefs[nVertRef];
                for(nRow = oValVertRef.bbox.r1; nRow <= oValVertRef.bbox.r2; ++nRow) {
                    oSeriesData = new CSeriesDataRefs(null);
                    aData.push(oSeriesData);
                    for(nHorRef = 0; nHorRef < aValHorRefs.length; ++nHorRef) {
                        oValHorRef = aValHorRefs[nHorRef];
                        oBBox = new Asc.Range(oValHorRef.bbox.c1, nRow, oValHorRef.bbox.c2, nRow, true);
                        oSeriesData.val.add(oValHorRef.createFromBBox(oValHorRef.worksheet, oBBox));
                        if(aCatVertRefs.length > 0) {
                            for(nCatRef = 0; nCatRef < aCatVertRefs.length; ++nCatRef) {
                                oCatRef = aCatVertRefs[nCatRef];
                                oBBox = new Asc.Range(oValHorRef.bbox.c1, oCatRef.bbox.r1, oValHorRef.bbox.c2, oCatRef.bbox.r2, true);
                                oSeriesData.cat.add(oCatRef.createFromBBox(oCatRef.worksheet, oBBox));
                            }
                        }
                    }
                    if(aTxHorRefs.length > 0) {
                        for(nTxRef = 0; nTxRef < aTxHorRefs.length; ++nTxRef) {
                            oTxRef = aTxHorRefs[nTxRef];
                            oBBox = new Asc.Range(oTxRef.bbox.c1, nRow, oTxRef.bbox.c2, nRow, true);
                            oSeriesData.tx.add(oTxRef.createFromBBox(oTxRef.worksheet, oBBox));
                        }
                    }
                }
            }
        }
        else if(this.info & AscFormat.SERIES_FLAG_VERT_VALUE) {
            aData = [];
            aCatHorRefs = this.cat.getHorRefs();
            aTxVertRefs = this.tx.getVertRefs();
            for(nHorRef = 0; nHorRef < aValHorRefs.length; ++nHorRef) {
                oValHorRef = aValHorRefs[nHorRef];
                for(nCol = oValHorRef.bbox.c1; nCol <= oValHorRef.bbox.c2; ++nCol) {
                    oSeriesData = new CSeriesDataRefs(null);
                    aData.push(oSeriesData);
                    for(nVertRef = 0; nVertRef < aValVertRefs.length; ++nVertRef) {
                        oValVertRef = aValVertRefs[nVertRef];
                        oBBox = new Asc.Range(nCol, oValVertRef.bbox.r1, nCol, oValVertRef.bbox.r2, true);
                        oSeriesData.val.add(oValVertRef.createFromBBox(oValVertRef.worksheet, oBBox));
                        if(aCatHorRefs.length > 0) {
                            for(nCatRef = 0; nCatRef < aCatHorRefs.length; ++nCatRef) {
                                oCatRef = aCatHorRefs[nCatRef];
                                oBBox = new Asc.Range(oCatRef.bbox.c1, oValVertRef.bbox.r1, oCatRef.bbox.c2, oValVertRef.bbox.r2, true);
                                oSeriesData.cat.add(oCatRef.createFromBBox(oCatRef.worksheet, oBBox));
                            }
                        }
                    }
                    if(aTxVertRefs.length > 0) {
                        for(nTxRef = 0; nTxRef < aTxVertRefs.length; ++nTxRef) {
                            oTxRef = aTxVertRefs[nTxRef];
                            oBBox = new Asc.Range(nCol, oTxRef.bbox.r1, nCol, oTxRef.bbox.r2, true);
                            oSeriesData.tx.add(oTxRef.createFromBBox(oTxRef.worksheet, oBBox));
                        }
                    }
                }
            }
        }
        return aData;
    };
    CChartDataRefs.prototype.getSeriesRefsFromSelectedRange = function(oSelectedRange, bScatter) {
        this.fillFromSelectedRange(oSelectedRange);
        if(this.info === 0) {
            return null;
        }
        if(this.tx.isEmpty() && this.cat.isEmpty()) {
            var oUnionRefs = this.calculateUnionRefs();
            if(oUnionRefs) {
                var bHor;
                if(this.info & AscFormat.SERIES_FLAG_HOR_VALUE) {
                    bHor = true;
                }
                else {
                    bHor = false;
                }
                return this.getSeriesRefsFromUnionRefs(oUnionRefs.aRefs, bHor, bScatter);
            }
            else {
                return this.getSeriesFromFixedRefs();
            }
        }
        else {
            return this.getSeriesFromFixedRefs();
        }
    };
    CChartDataRefs.prototype.getSwitchedRefs = function(bScatter) {
        this.updateDataRefs();
        if(this.info === 0) {
            return null;
        }
        var aData = null;
        if(!this.tx.isEmpty() && !this.cat.isEmpty()) {
            var nOldInfo = this.info;
            var oOldCat = this.cat;
            var oOldTx = this.tx;
            this.cat = oOldTx;
            this.tx = oOldCat;
            if(this.info & AscFormat.SERIES_FLAG_HOR_VALUE) {
                this.info = (this.info & (~AscFormat.SERIES_FLAG_HOR_VALUE)) | AscFormat.SERIES_FLAG_VERT_VALUE;
            }
            else if(this.info & AscFormat.SERIES_FLAG_VERT_VALUE) {
                this.info = (this.info & (~AscFormat.SERIES_FLAG_VERT_VALUE)) | AscFormat.SERIES_FLAG_HOR_VALUE;
            }
            aData = this.getSeriesFromFixedRefs();
            this.info = nOldInfo;
            this.cat = oOldCat;
            this.tx = oOldTx;
        }
        else {
            var oUnionRefs = this.getUnionRefs();
            if(oUnionRefs) {
                var bHor;
                if(this.info & AscFormat.SERIES_FLAG_HOR_VALUE) {
                    bHor = false;
                }
                else {
                    bHor = true;
                }
                aData = this.getSeriesRefsFromUnionRefs(oUnionRefs.aRefs, bHor, bScatter);
            }
        }
        return aData;
    };
    CChartDataRefs.prototype.checkDataRange = function(sRange, bHorValue, nType) {
        if(typeof  sRange !== "string") {
            return  Asc.c_oAscError.ID.DataRangeError;
        }
        var aSeriesRefs = this.getSeriesRefsFromUnionRefs(AscFormat.fParseChartFormulaExternal(sRange), bHorValue, isScatterChartType(nType));
        if(!Array.isArray(aSeriesRefs)) {
            return  Asc.c_oAscError.ID.DataRangeError;
        }
        var nRef;
        if(isStockChartType(nType)) {
            if(aSeriesRefs.length !== AscFormat.MIN_STOCK_COUNT) {
                return Asc.c_oAscError.ID.StockChartError;
            }
            for(nRef = 0; nRef < aSeriesRefs.length; ++nRef) {
                if(aSeriesRefs[nRef].getValCellsCount() < AscFormat.MIN_STOCK_COUNT) {
                    return Asc.c_oAscError.ID.StockChartError;
                }
            }
            return Asc.c_oAscError.ID.No;
        }
        if(isComboChartType(nType)) {
            if(aSeriesRefs.length < 2) {
                return Asc.c_oAscError.ID.ComboSeriesError;
            }
        }
        if(aSeriesRefs.length > AscFormat.MAX_SERIES_COUNT) {
            return Asc.c_oAscError.ID.MaxDataSeriesError;
        }
        for(nRef = 0; nRef < aSeriesRefs.length; ++nRef) {
            if(aSeriesRefs[nRef].getValCellsCount() > AscFormat.MAX_POINTS_COUNT) {
                return Asc.c_oAscError.ID.MaxDataPointsError;
            }
        }
        return Asc.c_oAscError.ID.No;
    };
    CChartDataRefs.prototype.fillSelectedRanges = function(oWSView) {
        this.updateDataRefs();
        fFillSelectedRanges(this.val, this.cat, this.tx, this.info, false, oWSView);
    };
    CChartDataRefs.prototype.fillFromSelectedRange = function(oSelectedRange) {
        this.updateDataRefs();
        fFillDataFromSelectedRange(this, oSelectedRange);
    };
    CChartDataRefs.prototype.privateCheckCellValueForHeader = function(oCell) {
        if(!this.privateCheckCellValueNumberOrEmpty(oCell)) {
            return true;
        }
        return this.privateCheckCellDateTimeFormat(oCell);
    };
    CChartDataRefs.prototype.privateCheckCellValueNumberOrEmpty = function(oCell) {
        var sValue = oCell.getValue();
	    var dNumValue = oCell.getNumberValue();
        if(AscFormat.isRealNumber(dNumValue) || sValue === "") {
            return true;
        }
        return false;
    };
    CChartDataRefs.prototype.privateCheckCellDateTimeFormat = function(oCell) {
        var nNumFmtType = oCell.getNumFormatType();
        if(Asc.c_oAscNumFormatType.Time === nNumFmtType ||
            Asc.c_oAscNumFormatType.Date === nNumFmtType) {
            return true;
        }
        return false;
    };
    CChartDataRefs.prototype.privateCheckCellDateTimeFormatFull = function(oCell) {
        if(this.privateCheckCellValueNumberOrEmpty(oCell) &&
            this.privateCheckCellDateTimeFormat(oCell)) {
            return true;
        }
        return false;
    };
    CChartDataRefs.prototype.hasIntersection = function(oRange) {
        var sWSName = oRange.worksheet.getName();
        var oSheetBounds = this.boundsByWS[sWSName];
        if(!oSheetBounds) {
            return false;
        }
        if(!oSheetBounds.isIntersect(oRange)) {
            return false;
        }
        if(this.info !== 0) {
            if(this.val.hasIntersection(oRange)) {
                return true;
            }
            if(this.cat.hasIntersection(oRange)) {
                return true;
            }
            if(this.tx.hasIntersection(oRange)) {
                return true;
            }
        }
        else {
            for(var nSeries = 0; nSeries < this.seriesRefs.length; ++nSeries) {
                if(this.seriesRefs[nSeries].hasIntersection(oRange)) {
                    return true;
                }
            }
        }
        for(var nLblRef = 0; nLblRef < this.labelsRefs.length; ++nLblRef) {
            if(this.labelsRefs[nLblRef].hasIntersection(oRange)) {
                return true;
            }
        }
        return false;
    };
    CChartDataRefs.prototype.collectRefsInsideRange = function(oRange, aRefs) {
        if(!this.hasIntersection(oRange)) {
            return;
        }
        for(var nSeries = 0; nSeries < this.seriesRefs.length; ++nSeries) {
            this.seriesRefs[nSeries].collectRefsInsideRange(oRange, aRefs);
        }
        for(var nLblRef = 0; nLblRef < this.labelsRefs.length; ++nLblRef) {
            this.labelsRefs[nLblRef].collectRefsInsideRange(oRange, aRefs);
        }
    };
    CChartDataRefs.prototype.collectIntersectionRefs = function(aRanges, aCollectedRefs) {
        if(!Array.isArray(aRanges) || aRanges.length === 0) {
            return;
        }
        var aIntersectionRanges = [];
        var oRange;
        for(var nRange = 0; nRange < aRanges.length; ++nRange) {
            oRange = aRanges[nRange];
            if(this.hasIntersection(oRange)) {
                aIntersectionRanges.push(oRange);
            }
        }
        if(aIntersectionRanges.length > 0) {
            for(var nSeries = 0; nSeries < this.seriesRefs.length; ++nSeries) {
                this.seriesRefs[nSeries].collectIntersectionRefs(aIntersectionRanges, aCollectedRefs);
            }
            for(var nLblRef = 0; nLblRef < this.labelsRefs.length; ++nLblRef) {
                this.labelsRefs[nLblRef].collectIntersectionRefs(aIntersectionRanges, aCollectedRefs);
            }
        }
    };

    function isValidChartRange(sRange) {
        if(sRange === "") {
            return Asc.c_oAscError.ID.No;
        }
        var sCheck = sRange;
        if(sRange[0] === "=") {
            sCheck = sRange.slice(1);
        }
        var aRanges = AscFormat.fParseChartFormulaExternal(sCheck);
        return (aRanges.length !== 0) ? Asc.c_oAscError.ID.No : Asc.c_oAscError.ID.DataRangeError;
    }

    function CChartStyle() {
        CBaseFormatObject.call(this);
        this.id = null;
        this.axisTitle = null;
        this.categoryAxis = null;
        this.chartArea = null;
        this.dataLabel = null;
        this.dataLabelCallout = null;
        this.dataPoint = null;
        this.dataPoint3D = null;
        this.dataPointLine = null;
        this.dataPointMarker = null;
        this.dataPointWireframe = null;
        this.dataTable = null;
        this.downBar = null;
        this.dropLine = null;
        this.errorBar = null;
        this.floor = null;
        this.gridlineMajor = null;
        this.gridlineMinor = null;
        this.hiLoLine = null;
        this.leaderLine = null;
        this.legend = null;
        this.plotArea = null;
        this.plotArea3D = null;
        this.seriesAxis = null;
        this.seriesLine = null;
        this.title = null;
        this.trendline = null;
        this.trendlineLabel = null;
        this.upBar = null;
        this.valueAxis = null;
        this.wall = null;
        this.markerLayout = null;
    }

    InitClass(CChartStyle, CBaseFormatObject, AscDFH.historyitem_type_ChartStyle);
    CChartStyle.prototype.addEntry = function(oEntry) {
        if(!AscCommon.isRealObject(oEntry)) {
            return;
        }
        switch (oEntry.type) {
            case 1: this.setAxisTitle(oEntry); break;
            case 2: this.setCategoryAxis(oEntry); break;
            case 3: this.setChartArea(oEntry); break;
            case 4: this.setDataLabel(oEntry); break;
            case 5: this.setDataLabelCallout(oEntry); break;
            case 6: this.setDataPoint(oEntry); break;
            case 7: this.setDataPoint3D(oEntry); break;
            case 8: this.setDataPointLine(oEntry); break;
            case 9: this.setDataPointMarker(oEntry); break;
            case 10: this.setDataPointWireframe(oEntry); break;
            case 11: this.setDataTable(oEntry); break;
            case 12: this.setDownBar(oEntry); break;
            case 13: this.setDropLine(oEntry); break;
            case 14: this.setErrorBar(oEntry); break;
            case 15: this.setFloor(oEntry); break;
            case 16: this.setGridlineMajor(oEntry); break;
            case 17: this.setGridlineMinor(oEntry); break;
            case 18: this.setHiLoLine(oEntry); break;
            case 19: this.setLeaderLine(oEntry); break;
            case 20: this.setLegend(oEntry); break;
            case 21: this.setPlotArea(oEntry); break;
            case 22: this.setPlotArea3D(oEntry); break;
            case 23: this.setSeriesAxis(oEntry); break;
            case 24: this.setSeriesLine(oEntry); break;
            case 25: this.setTitle(oEntry); break;
            case 26: this.setTrendline(oEntry); break;
            case 27: this.setTrendlineLabel(oEntry); break;
            case 28: this.setUpBar(oEntry); break;
            case 29: this.setValueAxis(oEntry); break;
            case 30: this.setWall(oEntry); break;
        }
    };
    CChartStyle.prototype.getStyleEntries = function() {
        return [this.axisTitle,
            this.categoryAxis,
            this.chartArea,
            this.dataLabel,
            this.dataLabelCallout,
            this.dataPoint,
            this.dataPoint3D,
            this.dataPointLine,
            this.dataPointMarker,
            this.dataPointWireframe,
            this.dataTable,
            this.downBar,
            this.dropLine,
            this.errorBar,
            this.floor,
            this.gridlineMajor,
            this.gridlineMinor,
            this.hiLoLine,
            this.leaderLine,
            this.legend,
            this.plotArea,
            this.plotArea3D,
            this.seriesAxis,
            this.seriesLine,
            this.title,
            this.trendline,
            this.trendlineLabel,
            this.upBar,
            this.valueAxis,
            this.wall];
    };
    CChartStyle.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.axisTitle) {
            oCopy.setAxisTitle(this.axisTitle.createDuplicate());
        }
        if(this.categoryAxis) {
            oCopy.setCategoryAxis(this.categoryAxis.createDuplicate());
        }
        if(this.chartArea) {
            oCopy.setChartArea(this.chartArea.createDuplicate());
        }
        if(this.dataLabel) {
            oCopy.setDataLabel(this.dataLabel.createDuplicate());
        }
        if(this.dataLabelCallout) {
            oCopy.setDataLabelCallout(this.dataLabelCallout.createDuplicate());
        }
        if(this.dataPoint) {
            oCopy.setDataPoint(this.dataPoint.createDuplicate());
        }
        if(this.dataPoint3D) {
            oCopy.setDataPoint3D(this.dataPoint3D.createDuplicate());
        }
        if(this.dataPointLine) {
            oCopy.setDataPointLine(this.dataPointLine.createDuplicate());
        }
        if(this.dataPointMarker) {
            oCopy.setDataPointMarker(this.dataPointMarker.createDuplicate());
        }
        if(this.dataPointWireframe) {
            oCopy.setDataPointWireframe(this.dataPointWireframe.createDuplicate());
        }
        if(this.dataTable) {
            oCopy.setDataTable(this.dataTable.createDuplicate());
        }
        if(this.downBar) {
            oCopy.setDownBar(this.downBar.createDuplicate());
        }
        if(this.dropLine) {
            oCopy.setDropLine(this.dropLine.createDuplicate());
        }
        if(this.errorBar) {
            oCopy.setErrorBar(this.errorBar.createDuplicate());
        }
        if(this.floor) {
            oCopy.setFloor(this.floor.createDuplicate());
        }
        if(this.gridlineMajor) {
            oCopy.setGridlineMajor(this.gridlineMajor.createDuplicate());
        }
        if(this.gridlineMinor) {
            oCopy.setGridlineMinor(this.gridlineMinor.createDuplicate());
        }
        if(this.hiLoLine) {
            oCopy.setHiLoLine(this.hiLoLine.createDuplicate());
        }
        if(this.leaderLine) {
            oCopy.setLeaderLine(this.leaderLine.createDuplicate());
        }
        if(this.legend) {
            oCopy.setLegend(this.legend.createDuplicate());
        }
        if(this.plotArea) {
            oCopy.setPlotArea(this.plotArea.createDuplicate());
        }
        if(this.plotArea3D) {
            oCopy.setPlotArea3D(this.plotArea3D.createDuplicate());
        }
        if(this.seriesAxis) {
            oCopy.setSeriesAxis(this.seriesAxis.createDuplicate());
        }
        if(this.seriesLine) {
            oCopy.setSeriesLine(this.seriesLine.createDuplicate());
        }
        if(this.title) {
            oCopy.setTitle(this.title.createDuplicate());
        }
        if(this.trendline) {
            oCopy.setTrendline(this.trendline.createDuplicate());
        }
        if(this.trendlineLabel) {
            oCopy.setTrendlineLabel(this.trendlineLabel.createDuplicate());
        }
        if(this.upBar) {
            oCopy.setUpBar(this.upBar.createDuplicate());
        }
        if(this.valueAxis) {
            oCopy.setValueAxis(this.valueAxis.createDuplicate());
        }
        if(this.wall) {
            oCopy.setWall(this.wall.createDuplicate());
        }
        if(this.markerLayout) {
            oCopy.setMarkerLayout(this.markerLayout.createDuplicate());
        }
        if(this.id !== null) {
            oCopy.setId(this.id);
        }
    };
    CChartStyle.prototype.setId = function(pr) {
        History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ChartStyleMarkerId, this.id, pr));
        this.id = pr;
    };
    CChartStyle.prototype.setAxisTitle = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleAxisTitle, this.axisTitle, pr));
        this.axisTitle = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setCategoryAxis = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleCategoryAxis, this.categoryAxis, pr));
        this.categoryAxis = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setChartArea = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleChartArea, this.chartArea, pr));
        this.chartArea = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataLabel = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataLabel, this.dataLabel, pr));
        this.dataLabel = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataLabelCallout = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataLabelCallout, this.dataLabelCallout, pr));
        this.dataLabelCallout = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataPoint = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataPoint, this.dataPoint, pr));
        this.dataPoint = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataPoint3D = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataPoint3D, this.dataPoint3D, pr));
        this.dataPoint3D = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataPointLine = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataPointLine, this.dataPointLine, pr));
        this.dataPointLine = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataPointMarker = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataPointMarker, this.dataPointMarker, pr));
        this.dataPointMarker = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataPointWireframe = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataPointWireframe, this.dataPointWireframe, pr));
        this.dataPointWireframe = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDataTable = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDataTable, this.dataTable, pr));
        this.dataTable = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDownBar = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDownBar, this.downBar, pr));
        this.downBar = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setDropLine = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleDropLine, this.dropLine, pr));
        this.dropLine = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setErrorBar = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleErrorBar, this.errorBar, pr));
        this.errorBar = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setFloor = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleFloor, this.floor, pr));
        this.floor = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setGridlineMajor = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleGridlineMajor, this.gridlineMajor, pr));
        this.gridlineMajor = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setGridlineMinor = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleGridlineMinor, this.gridlineMinor, pr));
        this.gridlineMinor = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setHiLoLine = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleHiLoLine, this.hiLoLine, pr));
        this.hiLoLine = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setLeaderLine = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleLeaderLine, this.leaderLine, pr));
        this.leaderLine = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setLegend = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleLegend, this.legend, pr));
        this.legend = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setPlotArea = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStylePlotArea, this.plotArea, pr));
        this.plotArea = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setPlotArea3D = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStylePlotArea3D, this.plotArea3D, pr));
        this.plotArea3D = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setSeriesAxis = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleSeriesAxis, this.seriesAxis, pr));
        this.seriesAxis = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setSeriesLine = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleSeriesLine, this.seriesLine, pr));
        this.seriesLine = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setTitle = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleTitle, this.title, pr));
        this.title = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setTrendline = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleTrendline, this.trendline, pr));
        this.trendline = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setTrendlineLabel = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleTrendlineLabel, this.trendlineLabel, pr));
        this.trendlineLabel = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setUpBar = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleUpBar, this.upBar, pr));
        this.upBar = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setValueAxis = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleValueAxis, this.valueAxis, pr));
        this.valueAxis = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setWall = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleWall, this.wall, pr));
        this.wall = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.setMarkerLayout = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleMarkerLayout, this.markerLayout, pr));
        this.markerLayout = pr;
        this.setParentToChild(pr);
    };
    CChartStyle.prototype.readAttribute = function(nType, pReader) {
    };
    CChartStyle.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };
    CChartStyle.prototype.privateWriteAttributes = function(pWriter) {
    };
    CChartStyle.prototype.writeChildren = function(pWriter) {
    };
    CChartStyle.prototype.getChildren = function() {
        return [];
    };
    CChartStyle.prototype.getDataEntry = function(oSeries) {
        var oStyleEntry = this.dataPoint;
        var nSeriesType = oSeries.getObjectType();
        switch(nSeriesType) {
            case AscDFH.historyitem_type_AreaSeries:
            case AscDFH.historyitem_type_BubbleSeries:
            case AscDFH.historyitem_type_PieSeries: {
                oStyleEntry = this.dataPoint;
                break;
            }
            case AscDFH.historyitem_type_BarSeries: {
                if(oSeries.b3D) {
                    oStyleEntry = this.dataPoint3D;
                }
                else {
                    oStyleEntry = this.dataPoint;
                }
                break;
            }
            case AscDFH.historyitem_type_LineSeries:
            case AscDFH.historyitem_type_RadarSeries:
            case AscDFH.historyitem_type_ScatterSer: {
                oStyleEntry = this.dataPointLine;
                break;
            }
            case AscDFH.historyitem_type_SurfaceSeries: {
                if(oSeries.isWireframe()) {
                    oStyleEntry = this.dataPointWireframe;
                }
                else {
                    oStyleEntry = this.dataPoint3D;
                }
                break;
            }
        }
        return oStyleEntry;
    };
    CChartStyle.prototype.specilaStyles = {1001: true, 1002: true};
    CChartStyle.prototype.isSpecialStyle = function() {
        return this.specilaStyles[this.id] === true;
    };

    function CStyleEntry() {
        CBaseFormatObject.call(this);
        this.type = null;
        this.lineWidthScale = null;
        this.lnRef = null;
        this.fillRef = null;
        this.effectRef = null;
        this.fontRef = null;
        this.defRPr = null;
        this.bodyPr = null;
        this.spPr = null;
    }

    InitClass(CStyleEntry, CBaseFormatObject, AscDFH.historyitem_type_ChartStyleEntry);
    CStyleEntry.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.type !== null) {
            oCopy.setType(this.type);
        }
        if(this.lineWidthScale !== null) {
            oCopy.setLineWidthScale(this.lineWidthScale.createDuplicate());
        }
        if(this.lnRef) {
            oCopy.setLnRef(this.lnRef.createDuplicate());
        }
        if(this.fillRef) {
            oCopy.setFillRef(this.fillRef.createDuplicate());
        }
        if(this.effectRef) {
            oCopy.setEffectRef(this.effectRef.createDuplicate());
        }
        if(this.fontRef) {
            oCopy.setFontRef(this.fontRef.createDuplicate());
        }
        if(this.defRPr) {
            oCopy.setDefRPr(this.defRPr.Copy());
        }
        if(this.bodyPr) {
            oCopy.setBodyPr(this.bodyPr.createDuplicate());
        }
        if(this.spPr) {
            oCopy.setSpPr(this.spPr.createDuplicate());
        }
    };
    CStyleEntry.prototype.setType = function(pr) {
        History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_ChartStyleEntryType, this.type, pr));
        this.type = pr;
    };
    CStyleEntry.prototype.setLineWidthScale = function(pr) {
        History.Add(new CChangesDrawingsDouble2(this, AscDFH.historyitem_ChartStyleEntryLineWidthScale, this.lineWidthScale, pr));
        this.lineWidthScale = pr;
    };
    CStyleEntry.prototype.setLnRef = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryLnRef, this.lnRef, pr));
        this.lnRef = pr;
    };
    CStyleEntry.prototype.setFillRef = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryFillRef, this.fillRef, pr));
        this.fillRef = pr;
    };
    CStyleEntry.prototype.setEffectRef = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryEffectRef, this.effectRef, pr));
        this.effectRef = pr;
    };
    CStyleEntry.prototype.setFontRef = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryFontRef, this.fontRef, pr));
        this.fontRef = pr;
    };
    CStyleEntry.prototype.setDefRPr = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryDefRPr, this.defRPr, pr));
        this.defRPr = pr;
    };
    CStyleEntry.prototype.setBodyPr = function(pr) {
        History.Add(new CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ChartStyleEntryBodyPr, this.bodyPr, pr));
        this.bodyPr = pr;
    };
    CStyleEntry.prototype.setSpPr = function(pr) {
        History.Add(new CChangesDrawingsObject(this, AscDFH.historyitem_ChartStyleEntrySpPr, this.spPr, pr));
        this.spPr = pr;
    };
    CStyleEntry.prototype.readAttribute = function(nType, pReader) {
    };
    CStyleEntry.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };
    CStyleEntry.prototype.privateWriteAttributes = function(pWriter) {
    };
    CStyleEntry.prototype.writeChildren = function(pWriter) {
    };
    CStyleEntry.prototype.getChildren = function() {
        return [];
    };
    CStyleEntry.prototype.isSpecialStyle = function() {
        if(this.parent) {
            return this.parent.isSpecialStyle();
        }
        return false;
    };
    CStyleEntry.prototype.specialPatterns = ["smGrid", "pct60", "wdDnDiag", "lgCheck", "pct75", "wdUpDiag", "plaid", "pct80", "lgConfetti", "narHorz", "pct90", "diagBrick", "cross", "pct30", "dkDnDiag", "smCheck",
        "pct70", "dkUpDiag", "dkVert", "trellis", "smConfetti", "dkHorz", "sphere", "diagCross", "lgGrid", "pct40", "ltDnDiag", "solidDmnd", "pct50", "ltUpDiag", "ltVert", "pct25",
        "pct5", "horz", "weave", "divot", "dotGrid", "pct10", "dnDiag", "zigZag", "pct20", "dashUpDiag", "narVert", "dashVert", "dotDmnd", "horzBrick", "shingle", "wave", "dashDnDiag",
        "dashVert"];
    CStyleEntry.prototype.getSpecialPatternType = function(nIdx) {
        var nArrayIdx = nIdx % this.specialPatterns.length;
        return AscCommon.global_hatch_offsets[this.specialPatterns[nArrayIdx]];
    };

    function CMarkerLayout() {
        CBaseFormatObject.call(this);
        this.symbol = null;
        this.size = null;
    }
    InitClass(CMarkerLayout, CBaseFormatObject, AscDFH.historyitem_type_MarkerLayout);
    CMarkerLayout.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.symbol !== null) {
            oCopy.setSymbol(this.symbol);
        }
        if(this.size !== null) {
            oCopy.setSize(this.size);
        }
    };
    CMarkerLayout.prototype.setSymbol = function(pr) {
        History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_MarkerLayoutSymbol, this.symbol, pr));
        this.symbol = pr;
    };
    CMarkerLayout.prototype.setSize = function(pr) {
        History.Add(new CChangesDrawingsLong(this, AscDFH.historyitem_MarkerLayoutSize, this.size, pr));
        this.size = pr;
    };
    CMarkerLayout.prototype.readAttribute = function(nType, pReader) {
    };
    CMarkerLayout.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };
    CMarkerLayout.prototype.privateWriteAttributes = function(pWriter) {
    };
    CMarkerLayout.prototype.writeChildren = function(pWriter) {
    };
    CMarkerLayout.prototype.getChildren = function() {
        return [];
    };

    function CChartColors() {
        this.meth = null;
        this.id = null;
        this.items = [];
    }
    CChartColors.prototype.Write_ToBinary = function(writer) {
        this.WriteToBinary(writer);
    };
    CChartColors.prototype.Read_FromBinary = function(writer) {
        this.ReadFromBinary(writer);
    };
    CChartColors.prototype.WriteToBinary = function(writer) {
        AscFormat.writeString(writer, this.meth);
        AscFormat.writeLong(writer, this.id);
        writer.WriteLong(this.items.length);
        for(var nItem = 0; nItem < this.items.length; ++nItem) {
            var oItem = this.items[nItem];
            var bIsUnicolor = (oItem instanceof AscFormat.CUniColor);
            writer.WriteBool(bIsUnicolor);
            oItem.Write_ToBinary(writer);
        }
    };
    CChartColors.prototype.ReadFromBinary = function(reader) {
        this.meth = AscFormat.readString(reader);
        this.id = AscFormat.readLong(reader);
        var nCount = reader.GetLong();
        for(var nItem = 0; nItem < nCount; ++nItem) {
            var bIsUnicolor = reader.GetBool();
            var oItem;
            if(bIsUnicolor) {
                oItem = new AscFormat.CUniColor();
            }
            else {
                oItem = new AscFormat.CColorModifiers();
            }
            oItem.Read_FromBinary(reader);
            this.items.push(oItem);
        }
    };
    CChartColors.prototype.createDuplicate = function() {
        var oCopy = new CChartColors();
        oCopy.meth = this.meth;
        oCopy.id = this.id;
        for(var nItem = 0; nItem < this.items.length; ++nItem) {
            oCopy.items.push(this.items[nItem].createDuplicate());
        }
        return oCopy;
    };
    CChartColors.prototype.setMeth = function(pr) {
        this.meth = pr;
    };
    CChartColors.prototype.setId = function(pr) {
        this.id = pr;
    };
    CChartColors.prototype.addItem = function(pr) {
        this.items.push(pr);
    };
    CChartColors.prototype.addItem = function(pr) {
        this.items.push(pr);
    };
    CChartColors.prototype.getBaseColors = function() {
        var aColors = [];
        for(var nItem = 0; nItem < this.items.length; ++nItem) {
            var oItem = this.items[nItem];
            if(oItem instanceof AscFormat.CUniColor) {
                aColors.push(oItem);
            }
        }
        return aColors;
    };
    CChartColors.prototype.getVariations = function() {
        var aVariations = [];
        for(var nItem = 0; nItem < this.items.length; ++nItem) {
            var oItem = this.items[nItem];
            if(oItem instanceof AscFormat.CColorModifiers) {
                aVariations.push(oItem);
            }
        }
        return aVariations;
    };
    CChartColors.prototype.generateCycleColors = function(aBaseColors, aVariations, nCount) {
        var aColors = [];
        for(var nColor = 0; nColor < nCount; ++nColor) {
            var oBaseColor = aBaseColors[nColor % aBaseColors.length];
            var oVariation = null;
            if(aVariations.length > 0) {
                var nVariation = (nColor / aBaseColors.length) >> 0;
                oVariation = aVariations[nVariation % aVariations.length];
            }
            var oColor = oBaseColor.createDuplicate();
            if(oVariation) {
                oColor.Mods = oVariation.createDuplicate();
            }
            aColors.push(oColor);
        }
        return aColors;
    };
    CChartColors.prototype.generateWithinLinearColors = function(aBaseColors, aVariations, nCount) {
        return this.generateCycleColors(aBaseColors, aVariations, nCount);
    };
    CChartColors.prototype.generateAcrossLinearColors = function(aBaseColors, aVariations, nCount) {
        return this.generateCycleColors(aBaseColors, aVariations, nCount);
    };
    CChartColors.prototype.generateWithinLinearReversedColors = function(aBaseColors, aVariations, nCount) {
        return this.generateCycleColors(aBaseColors, aVariations, nCount);
    };
    CChartColors.prototype.generateAcrossLinearReversedColors = function(aBaseColors, aVariations, nCount) {
        return this.generateCycleColors(aBaseColors, aVariations, nCount);
    };
    CChartColors.prototype.generateColors = function(nCount) {
        var aBaseColors = this.getBaseColors();
        var aVariations = this.getVariations();
        var sMeth = this.meth || "cycle";
        if("cycle" === sMeth) {
            return this.generateCycleColors(aBaseColors, aVariations, nCount);
        }
        else if("withinLinear" === sMeth) {
            return this.generateWithinLinearColors(aBaseColors, aVariations, nCount)
        }
        else if("acrossLinear" === sMeth) {
            return this.generateAcrossLinearColors(aBaseColors, aVariations, nCount)
        }
        else if("withinLinearReversed" === sMeth) {
            return this.generateWithinLinearReversedColors(aBaseColors, aVariations, nCount)
        }
        else if("acrossLinearReversed" === sMeth) {
            return this.generateAcrossLinearReversedColors(aBaseColors, aVariations, nCount)
        }
        return [];
    };

    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ChartSpace_ChartColors] = CChartColors;

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CDLbl = CDLbl;
    window['AscFormat'].CPlotArea = CPlotArea;
    window['AscFormat'].CBarChart = CBarChart;
    window['AscFormat'].CAreaChart = CAreaChart;
    window['AscFormat'].CAreaSeries = CAreaSeries;
    window['AscFormat'].CCatAx = CCatAx;
    window['AscFormat'].CDateAx = CDateAx;
    window['AscFormat'].CSerAx = CSerAx;
    window['AscFormat'].CValAx = CValAx;
    window['AscFormat'].CBandFmt = CBandFmt;
    window['AscFormat'].CBarSeries = CBarSeries;
    window['AscFormat'].CBubbleChart = CBubbleChart;
    window['AscFormat'].CBubbleSeries = CBubbleSeries;
    window['AscFormat'].CCat = CCat;
    window['AscFormat'].CChartText = CChartText;
    window['AscFormat'].CDLbls = CDLbls;
    window['AscFormat'].CDPt = CDPt;
    window['AscFormat'].CDTable = CDTable;
    window['AscFormat'].CDispUnits = CDispUnits;
    window['AscFormat'].CDoughnutChart = CDoughnutChart;
    window['AscFormat'].CErrBars = CErrBars;
    window['AscFormat'].CLayout = CLayout;
    window['AscFormat'].CLegend = CLegend;
    window['AscFormat'].CLegendEntry = CLegendEntry;
    window['AscFormat'].CLineChart = CLineChart;
    window['AscFormat'].CLineSeries = CLineSeries;
    window['AscFormat'].CMarker = CMarker;
    window['AscFormat'].CMinusPlus = CMinusPlus;
    window['AscFormat'].CMultiLvlStrCache = CMultiLvlStrCache;
    window['AscFormat'].CMultiLvlStrRef = CMultiLvlStrRef;
    window['AscFormat'].CNumRef = CNumRef;
    window['AscFormat'].CNumericPoint = CNumericPoint;
    window['AscFormat'].CNumFmt = CNumFmt;
    window['AscFormat'].CNumLit = CNumLit;
    window['AscFormat'].COfPieChart = COfPieChart;
    window['AscFormat'].CPictureOptions = CPictureOptions;
    window['AscFormat'].CPieChart = CPieChart;
    window['AscFormat'].CPieSeries = CPieSeries;
    window['AscFormat'].CPivotFmt = CPivotFmt;
    window['AscFormat'].CRadarChart = CRadarChart;
    window['AscFormat'].CRadarSeries = CRadarSeries;
    window['AscFormat'].CScaling = CScaling;
    window['AscFormat'].CScatterChart = CScatterChart;
    window['AscFormat'].CScatterSeries = CScatterSeries;
    window['AscFormat'].CTx = CTx;
    window['AscFormat'].CStockChart = CStockChart;
    window['AscFormat'].CStrCache = CStrCache;
    window['AscFormat'].CStringPoint = CStringPoint;
    window['AscFormat'].CStrRef = CStrRef;
    window['AscFormat'].CSurfaceChart = CSurfaceChart;
    window['AscFormat'].CSurfaceSeries = CSurfaceSeries;
    window['AscFormat'].CTitle = CTitle;
    window['AscFormat'].CTrendLine = CTrendLine;
    window['AscFormat'].CUpDownBars = CUpDownBars;
    window['AscFormat'].CYVal = CYVal;
    window['AscFormat'].CChart = CChart;
    window['AscFormat'].CChartWall = CChartWall;
    window['AscFormat'].CView3d = CView3d;
    window['AscFormat'].CExternalData = CExternalData;
    window['AscFormat'].CPivotSource = CPivotSource;
    window['AscFormat'].CProtection = CProtection;
    window['AscFormat'].CPrintSettings = CPrintSettings;
    window['AscFormat'].CHeaderFooterChart = CHeaderFooterChart;
    window['AscFormat'].CPageMarginsChart = CPageMarginsChart;
    window['AscFormat'].CPageSetup = CPageSetup;
    window['AscFormat'].CreateTextBodyFromString = CreateTextBodyFromString;
    window['AscFormat'].CreateDocContentFromString = CreateDocContentFromString;
    window['AscFormat'].AddToContentFromString = AddToContentFromString;
    window['AscFormat'].CheckContentTextAndAdd = CheckContentTextAndAdd;
    window['AscFormat'].CValAxisLabels = CValAxisLabels;
    window['AscFormat'].CalcLegendEntry = CalcLegendEntry;
    window['AscFormat'].CUnionMarker = CUnionMarker;
    window['AscFormat'].CreateMarkerGeometryByType = CreateMarkerGeometryByType;
    window['AscFormat'].isScatterChartType = isScatterChartType;
    window['AscFormat'].fParseChartFormula = fParseChartFormula;
    window['AscFormat'].fParseChartFormulaExternal = fParseChartFormulaExternal;
    window['AscFormat'].fCreateRef = fCreateRef;
    window['AscFormat'].CChartDataRefs = CChartDataRefs;
    window['AscFormat'].getIsMarkerByType = getIsMarkerByType;
    window['AscFormat'].getIsSmoothByType = getIsSmoothByType;
    window['AscFormat'].getIsLineByType = getIsLineByType;
    window['AscFormat'].getIsLineType = getIsLineType;
    window['AscFormat'].isValidChartRange = isValidChartRange;
    window['AscFormat'].CChartStyle = CChartStyle;
    window['AscFormat'].CStyleEntry = CStyleEntry;
    window['AscFormat'].CMarkerLayout = CMarkerLayout;
    window['AscFormat'].CChartColors = CChartColors;
    window['AscFormat'].CBaseChartObject = CBaseChartObject;

    window['AscFormat'].AX_POS_L = AX_POS_L;
    window['AscFormat'].AX_POS_T = AX_POS_T;
    window['AscFormat'].AX_POS_R = AX_POS_R;
    window['AscFormat'].AX_POS_B = AX_POS_B;

    window['AscFormat'].CROSSES_AUTO_ZERO = CROSSES_AUTO_ZERO;
    window['AscFormat'].CROSSES_MAX = CROSSES_MAX;
    window['AscFormat'].CROSSES_MIN = CROSSES_MIN;

    window['AscFormat'].LBL_ALG_CTR = LBL_ALG_CTR;
    window['AscFormat'].LBL_ALG_L = LBL_ALG_L;
    window['AscFormat'].LBL_ALG_R = LBL_ALG_R;

    window['AscFormat'].TIME_UNIT_DAYS = TIME_UNIT_DAYS;
    window['AscFormat'].TIME_UNIT_MONTHS = TIME_UNIT_MONTHS;
    window['AscFormat'].TIME_UNIT_YEARS = TIME_UNIT_YEARS;

    window['AscFormat'].CROSS_BETWEEN_BETWEEN = CROSS_BETWEEN_BETWEEN;
    window['AscFormat'].CROSS_BETWEEN_MID_CAT = CROSS_BETWEEN_MID_CAT;

    window['AscFormat'].SYMBOL_CIRCLE = SYMBOL_CIRCLE;
    window['AscFormat'].SYMBOL_DASH = SYMBOL_DASH;
    window['AscFormat'].SYMBOL_DIAMOND = SYMBOL_DIAMOND;
    window['AscFormat'].SYMBOL_DOT = SYMBOL_DOT;
    window['AscFormat'].SYMBOL_NONE = SYMBOL_NONE;
    window['AscFormat'].SYMBOL_PICTURE = SYMBOL_PICTURE;
    window['AscFormat'].SYMBOL_PLUS = SYMBOL_PLUS;
    window['AscFormat'].SYMBOL_SQUARE = SYMBOL_SQUARE;
    window['AscFormat'].SYMBOL_STAR = SYMBOL_STAR;
    window['AscFormat'].SYMBOL_TRIANGLE = SYMBOL_TRIANGLE;
    window['AscFormat'].SYMBOL_X = SYMBOL_X;

    window['AscFormat'].MARKER_SYMBOL_TYPE = MARKER_SYMBOL_TYPE;

    window['AscFormat'].ORIENTATION_MAX_MIN = ORIENTATION_MAX_MIN;
    window['AscFormat'].ORIENTATION_MIN_MAX = ORIENTATION_MIN_MAX;

    window['AscFormat'].SERIES_FLAG_HOR_VALUE = SERIES_FLAG_HOR_VALUE;
    window['AscFormat'].SERIES_FLAG_VERT_VALUE = SERIES_FLAG_VERT_VALUE;
    window['AscFormat'].SERIES_FLAG_CAT = SERIES_FLAG_CAT;
    window['AscFormat'].SERIES_FLAG_TX = SERIES_FLAG_TX;
    window['AscFormat'].SERIES_FLAG_CONTINUOUS = SERIES_FLAG_CONTINUOUS;
})(window);
