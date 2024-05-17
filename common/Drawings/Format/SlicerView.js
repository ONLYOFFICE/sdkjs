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

(function() {

    AscDFH.drawingsChangesMap[AscDFH.historyitem_SlicerViewName] = function(oClass, value){
        oClass.name = value;
        oClass.data.clear();
        oClass.removeAllButtonsTmpState();
        oClass.handleUpdateExtents();
    };
    AscDFH.changesFactory[AscDFH.historyitem_SlicerViewName] = window['AscDFH'].CChangesDrawingsString;

    var LEFT_PADDING = 2.54;
    var RIGHT_PADDING = LEFT_PADDING;
    var BOTTOM_PADDING = LEFT_PADDING;
    var TOP_PADDING = LEFT_PADDING / 2;
    var SPACE_BETWEEN = LEFT_PADDING * 14 / 37;
    var HEADER_BUTTON_WIDTH = 20 * AscCommon.g_dKoef_pix_to_mm;
    var HEADER_TOP_PADDING = LEFT_PADDING / 4;
    var HEADER_BOTTOM_PADDING = HEADER_TOP_PADDING;
    var HEADER_LEFT_PADDING = LEFT_PADDING;
    var HEADER_RIGHT_PADDING = HEADER_TOP_PADDING + RIGHT_PADDING + 2*HEADER_BUTTON_WIDTH;
    var SCROLL_WIDTH = 15 * 25.4 / 96;
    var SCROLLER_WIDTH = 13 * 25.4 / 96;

    var STATE_FLAG_WHOLE = 1;
    var STATE_FLAG_HEADER = 2;
    var STATE_FLAG_SELECTED = 4;
    var STATE_FLAG_DATA = 8;
    var STATE_FLAG_HOVERED = 16;

    var SCROLL_TIMER_INTERVAL = 200;

    var LOCKED_ALPHA = 51;

    var STYLE_TYPE = {};
    STYLE_TYPE.WHOLE = STATE_FLAG_WHOLE;
    STYLE_TYPE.HEADER = STATE_FLAG_HEADER;
    STYLE_TYPE.SELECTED_DATA = STATE_FLAG_SELECTED | STATE_FLAG_DATA | 0;
    STYLE_TYPE.SELECTED_NO_DATA = STATE_FLAG_SELECTED | 0 | 0;
    STYLE_TYPE.UNSELECTED_DATA = 0 | STATE_FLAG_DATA | 0;
    STYLE_TYPE.UNSELECTED_NO_DATA = 0 | 0 | 0;
    STYLE_TYPE.HOVERED_SELECTED_DATA = STATE_FLAG_SELECTED | STATE_FLAG_DATA | STATE_FLAG_HOVERED;
    STYLE_TYPE.HOVERED_SELECTED_NO_DATA = STATE_FLAG_SELECTED | 0 | STATE_FLAG_HOVERED;
    STYLE_TYPE.HOVERED_UNSELECTED_DATA = 0 | STATE_FLAG_DATA | STATE_FLAG_HOVERED;
    STYLE_TYPE.HOVERED_UNSELECTED_NO_DATA = 0 | 0 | STATE_FLAG_HOVERED;

    var SCROLL_COLORS = {};
    SCROLL_COLORS[STYLE_TYPE.WHOLE] = 0xF1F1F1;
    SCROLL_COLORS[STYLE_TYPE.HEADER] = 0xF1F1F1;
    SCROLL_COLORS[STYLE_TYPE.SELECTED_DATA] = 0xADADAD;
    SCROLL_COLORS[STYLE_TYPE.SELECTED_NO_DATA] = 0xADADAD;
    SCROLL_COLORS[STYLE_TYPE.UNSELECTED_DATA] = 0xF1F1F1;
    SCROLL_COLORS[STYLE_TYPE.UNSELECTED_NO_DATA] = 0xF1F1F1;
    SCROLL_COLORS[STYLE_TYPE.HOVERED_SELECTED_DATA] = 0xADADAD;
    SCROLL_COLORS[STYLE_TYPE.HOVERED_SELECTED_NO_DATA] = 0xADADAD;
    SCROLL_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_DATA] = 0xCFCFCF;
    SCROLL_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_NO_DATA] = 0xCFCFCF;

    var SCROLL_BUTTON_TYPE_LEFT = 0;
    var SCROLL_BUTTON_TYPE_TOP = 1;
    var SCROLL_BUTTON_TYPE_RIGHT = 2;
    var SCROLL_BUTTON_TYPE_BOTTOM = 3;

    var SCROLL_ARROW_COLORS = {};
    SCROLL_ARROW_COLORS[STYLE_TYPE.WHOLE] = 0xF1F1F1;
    SCROLL_ARROW_COLORS[STYLE_TYPE.HEADER] = 0xF1F1F1;
    SCROLL_ARROW_COLORS[STYLE_TYPE.SELECTED_DATA] = 0xFFFFFF;
    SCROLL_ARROW_COLORS[STYLE_TYPE.SELECTED_NO_DATA] = 0xFFFFFF;
    SCROLL_ARROW_COLORS[STYLE_TYPE.UNSELECTED_DATA] = 0xADADAD;
    SCROLL_ARROW_COLORS[STYLE_TYPE.UNSELECTED_NO_DATA] = 0xADADAD;
    SCROLL_ARROW_COLORS[STYLE_TYPE.HOVERED_SELECTED_DATA] = 0xFFFFFF;
    SCROLL_ARROW_COLORS[STYLE_TYPE.HOVERED_SELECTED_NO_DATA] = 0xFFFFFF;
    SCROLL_ARROW_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_DATA] = 0xFFFFFF;
    SCROLL_ARROW_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_NO_DATA] = 0xFFFFFF;

    var HEADER_BUTTON_COLORS = {};
    HEADER_BUTTON_COLORS[STYLE_TYPE.WHOLE] = null;
    HEADER_BUTTON_COLORS[STYLE_TYPE.HEADER] = null;
    HEADER_BUTTON_COLORS[STYLE_TYPE.SELECTED_DATA] = 0x7D858C;
    HEADER_BUTTON_COLORS[STYLE_TYPE.SELECTED_NO_DATA] = 0x7D858C;
    HEADER_BUTTON_COLORS[STYLE_TYPE.UNSELECTED_DATA] = null;
    HEADER_BUTTON_COLORS[STYLE_TYPE.UNSELECTED_NO_DATA] = null;
    HEADER_BUTTON_COLORS[STYLE_TYPE.HOVERED_SELECTED_DATA] = 0x7D858C;
    HEADER_BUTTON_COLORS[STYLE_TYPE.HOVERED_SELECTED_NO_DATA] = 0x7D858C;
    HEADER_BUTTON_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_DATA] = 0xD8DADC;
    HEADER_BUTTON_COLORS[STYLE_TYPE.HOVERED_UNSELECTED_NO_DATA] = 0xD8DADC;

    var ICON_MULTISELECT = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4NCjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNCA1QzQgNC40NDc3MiA0LjQ0NzcyIDQgNSA0SDE1QzE1LjU1MjMgNCAxNiA0LjQ0NzcyIDE2IDVWNkMxNiA2LjU1MjI4IDE1LjU1MjMgNyAxNSA3SDVDNC40NDc3MiA3IDQgNi41NTIyOCA0IDZWNVpNMTUgMTNINUw1IDE0SDE1VjEzWk01IDEyQzQuNDQ3NzIgMTIgNCAxMi40NDc3IDQgMTNWMTRDNCAxNC41NTIzIDQuNDQ3NzIgMTUgNSAxNUgxNUMxNS41NTIzIDE1IDE2IDE0LjU1MjMgMTYgMTRWMTNDMTYgMTIuNDQ3NyAxNS41NTIzIDEyIDE1IDEySDVaTTUgOEM0LjQ0NzcyIDggNCA4LjQ0NzcyIDQgOVYxMEM0IDEwLjU1MjMgNC40NDc3MiAxMSA1IDExSDE1QzE1LjU1MjMgMTEgMTYgMTAuNTUyMyAxNiAxMFY5QzE2IDguNDQ3NzIgMTUuNTUyMyA4IDE1IDhINVoiIGZpbGw9IiM0NDQ0NDQiLz4NCjwvc3ZnPg0K";
    var ICON_MULTISELECT_INVERTED = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4NCjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNCA1QzQgNC40NDc3MiA0LjQ0NzcyIDQgNSA0SDE1QzE1LjU1MjMgNCAxNiA0LjQ0NzcyIDE2IDVWNkMxNiA2LjU1MjI4IDE1LjU1MjMgNyAxNSA3SDVDNC40NDc3MiA3IDQgNi41NTIyOCA0IDZWNVpNMTUgMTNINUw1IDE0SDE1VjEzWk01IDEyQzQuNDQ3NzIgMTIgNCAxMi40NDc3IDQgMTNWMTRDNCAxNC41NTIzIDQuNDQ3NzIgMTUgNSAxNUgxNUMxNS41NTIzIDE1IDE2IDE0LjU1MjMgMTYgMTRWMTNDMTYgMTIuNDQ3NyAxNS41NTIzIDEyIDE1IDEySDVaTTUgOEM0LjQ0NzcyIDggNCA4LjQ0NzcyIDQgOVYxMEM0IDEwLjU1MjMgNC40NDc3MiAxMSA1IDExSDE1QzE1LjU1MjMgMTEgMTYgMTAuNTUyMyAxNiAxMFY5QzE2IDguNDQ3NzIgMTUuNTUyMyA4IDE1IDhINVoiIGZpbGw9IndoaXRlIi8+DQo8L3N2Zz4NCg==";
    var ICON_CLEAR_FILTER = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4NCjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNMTAgMTZMOCAxNFYxMEwzIDVIMTVMMTAgMTBWMTZaTTE2LjE0NjQgMTYuODUzNkwxNC41IDE1LjIwNzFMMTIuODUzNiAxNi44NTM2TDEyLjE0NjQgMTYuMTQ2NUwxMy43OTI5IDE0LjVMMTIuMTQ2NCAxMi44NTM2TDEyLjg1MzYgMTIuMTQ2NUwxNC41IDEzLjc5MjlMMTYuMTQ2NCAxMi4xNDY1TDE2Ljg1MzYgMTIuODUzNkwxNS4yMDcxIDE0LjVMMTYuODUzNiAxNi4xNDY1TDE2LjE0NjQgMTYuODUzNloiIGZpbGw9IiM0NDQ0NDQiLz4NCjwvc3ZnPg0K";
    var ICON_CLEAR_FILTER_DISABLED = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4NCjxnIG9wYWNpdHk9IjAuNCI+DQo8cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTEwIDE2TDggMTRWMTBMMyA1SDE1TDEwIDEwVjE2Wk0xNi4xNDY0IDE2Ljg1MzZMMTQuNSAxNS4yMDcxTDEyLjg1MzYgMTYuODUzNkwxMi4xNDY0IDE2LjE0NjVMMTMuNzkyOSAxNC41TDEyLjE0NjQgMTIuODUzNkwxMi44NTM2IDEyLjE0NjVMMTQuNSAxMy43OTI5TDE2LjE0NjQgMTIuMTQ2NUwxNi44NTM2IDEyLjg1MzZMMTUuMjA3MSAxNC41TDE2Ljg1MzYgMTYuMTQ2NUwxNi4xNDY0IDE2Ljg1MzZaIiBmaWxsPSIjNDQ0NDQ0Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";

    var oDefaultWrapObject = {oTxWarpStruct: null, oTxWarpStructParamarks: null, oTxWarpStructNoTransform: null, oTxWarpStructParamarksNoTransform: null};


    function getSlicerIconsForLoad() {
        return [ICON_MULTISELECT, ICON_MULTISELECT_INVERTED, ICON_CLEAR_FILTER, ICON_CLEAR_FILTER_DISABLED];
    }

    function getGraphicsScale(graphics) {
        var vector_koef;
        if(graphics.m_oCoordTransform) {
            vector_koef = 1.0 / graphics.m_oCoordTransform.sx;
        }
        else {
            var zoom = 1, ppiX = 96;
            if (window.Asc && window.Asc.editor) {
                zoom = window.Asc.editor.asc_getZoom();
            }
            vector_koef = 25.4 / (ppiX * zoom);
            vector_koef /= AscCommon.AscBrowser.retinaPixelRatio;
        }
        return vector_koef;
    }
    function setGraphicsSettings(graphics, oBorderPr, oLastBorderPr) {
        var oColor = oBorderPr.c || new AscCommonExcel.RgbColor(0);
        var oLastColor = null;
        if(oLastBorderPr) {
            oLastColor = oLastBorderPr.c || new AscCommonExcel.RgbColor(0);
        }
        if(!oColor.isEqual(oLastColor)) {
            graphics.p_color(oColor.getR(), oColor.getG(), oColor.getB(), 255);
        }
        var nW = oBorderPr.w, dW;
        if(!AscFormat.isRealNumber(nW)) {
            nW = 1;
        }
        var vector_koef;
        vector_koef = getGraphicsScale(graphics);
        dW = nW * vector_koef;
        if(!oLastBorderPr || oBorderPr.s !== oLastBorderPr.s) {
            var aDash = oBorderPr.getDashSegments();
            if(!graphics.m_oCoordTransform) {
                if(!AscFormat.isRealNumber(vector_koef)) {
                    vector_koef = getGraphicsScale(graphics);
                }
                for(var nSegment = 0; nSegment < aDash.length; ++nSegment) {
                    aDash[nSegment] *= vector_koef;
                }
            }
            graphics.p_dash(aDash);
        }
        return dW;
    }
    function drawHorBorder(graphics, oBorderPr, oPrevBorderPr, align, y, x, r) {
        var oLastBorderPr = null;
        if(oBorderPr && oBorderPr.s !== Asc.c_oAscBorderStyles.None) {
            graphics.drawHorLine(align, y, x, r, setGraphicsSettings(graphics, oBorderPr, oPrevBorderPr));
            oLastBorderPr = oBorderPr;
        }
        return oLastBorderPr;
    }
    function drawVerBorder(graphics, oBorderPr, oPrevBorderPr, align, x, y, b) {
        var oLastBorderPr = null;
        if(oBorderPr && oBorderPr.s !== Asc.c_oAscBorderStyles.None) {
            graphics.drawVerLine(align, x, y, b, setGraphicsSettings(graphics, oBorderPr, oPrevBorderPr));
            oLastBorderPr = oBorderPr;
        }
        return oLastBorderPr;
    }

    function CSlicerCache() {
        this.view = null;
        this.values = null;
        this.locked = false;
        this.styleName = null;
    }
    CSlicerCache.prototype.setView = function(oView) {
        this.view = oView;
    };
    CSlicerCache.prototype.setValues = function(oValues) {
        this.values = oValues;
    };
    CSlicerCache.prototype.setLocked = function(bVal) {
        this.locked = bVal;
    };
    CSlicerCache.prototype.setStyleName = function(sVal) {
        this.styleName = sVal;
    };
    CSlicerCache.prototype.getView = function() {
        return this.view;
    };
    CSlicerCache.prototype.getValues = function() {
        if(Array.isArray(this.values)) {
            return this.values;
        }
        return [];
    };
    CSlicerCache.prototype.getLocked = function() {
        return this.locked;
    };
    CSlicerCache.prototype.getStyleName = function() {
        return this.styleName;
    };
    CSlicerCache.prototype.clear = function() {
        this.values = null;
        this.view = null;
        this.styleName = null;
    };
    CSlicerCache.prototype.hasData = function() {
        return this.values !== null && this.view !== null;
    };
    CSlicerCache.prototype.save = function() {
        var oRet = new CSlicerCache();
        if(this.values) {
            oRet.values = this.values;
        }
        if(this.view) {
            oRet.view = this.view;
        }
        if(this.styleName) {
            oRet.styleName = this.styleName;
        }
        return oRet;
    };
    CSlicerCache.prototype.getValues = function() {
        if(Array.isArray(this.values)) {
            return this.values;
        }
        return [];
    };
    CSlicerCache.prototype.getValuesCount = function () {
        return this.getValues().length;
    };
    CSlicerCache.prototype.getCaption = function() {
        if(this.view && typeof this.view.caption === "string") {
            return this.view.caption;
        }
        return "";
    };
    CSlicerCache.prototype.getShowCaption = function() {
        if(this.view) {
            return this.view.showCaption !== false;
        }
        return false;
    };
    CSlicerCache.prototype.getColumnsCount = function() {
        if(this.view && AscFormat.isRealNumber(this.view.columnCount)) {
            return this.view.columnCount;
        }
        return 1;
    };
    CSlicerCache.prototype.getButtonHeight = function() {
        if(this.view && AscFormat.isRealNumber(this.view.rowHeight)) {
            return this.view.rowHeight * g_dKoef_emu_to_mm;
        }
        return 0.26 * 25.4;
    };
    CSlicerCache.prototype.getValue = function (nIndex) {
        if(nIndex > -1 && nIndex < this.getValuesCount()) {
            return this.getValues()[nIndex];
        }
        return null;
    };
    CSlicerCache.prototype.getVal = function(oValue) {
        if(!oValue) {
            return null;
        }
        return oValue.val;
    };
    CSlicerCache.prototype.getHiddenByOther = function(oValue) {
        if(!oValue) {
            return false;
        }
        return oValue.hiddenByOtherColumns === true;
    };
    CSlicerCache.prototype.getVisible = function(oValue) {
        if(!oValue) {
            return null;
        }
        return oValue.visible !== false;
    };
    CSlicerCache.prototype.getString = function (nIndex) {
        var oValue = this.getValue(nIndex);
        if(oValue && typeof oValue.text === "string" && oValue.text.length > 0) {
            return oValue.text;
        }
        return AscCommon.translateManager.getValue("(blank)");
    };

    function CSlicerData(slicer) {
        CSlicerCache.call(this);
        this.slicer = slicer;
    }
    CSlicerData.prototype = Object.create(CSlicerCache.prototype);
    CSlicerData.prototype.constructor = CSlicerData;
    CSlicerData.prototype.updateData = function() {
        this.clear();
        var oData = this.retrieveData();
        this.values = oData.getValues();
        this.view = oData.getView();
        this.styleName = oData.getStyleName();
    };
    CSlicerData.prototype.retrieveData = function() {
        var oData = new CSlicerCache();
        var oWorkbook = this.slicer.getWorkbook();
        if(!oWorkbook) {
            return oData;
        }
        var sName = this.slicer.getName();
        var oView = oWorkbook.getSlicerByName(sName);
        if(!oView) {
            return oData;
        }
        var oCache = oView.getCacheDefinition();
        if(!oCache) {
            return oData;
        }
        var oValues = oCache.getFilterValues();
        if(!oValues || !Array.isArray(oValues.values)) {
            return oData
        }
        oData.setValues(oValues.values);
        oData.setView(oView.clone());
        oData.setStyleName(this.slicer.getStyleName());
        return oData;
    };
    CSlicerData.prototype.checkData = function() {
        if(!this.hasData()) {
            this.updateData();
        }
    };
    CSlicerData.prototype.getValues = function() {
        this.checkData();
        return CSlicerCache.prototype.getValues.call(this);
    };
    CSlicerData.prototype.getCaption = function() {
        this.checkData();
        return CSlicerCache.prototype.getCaption.call(this);
    };
    CSlicerData.prototype.getShowCaption = function() {
        this.checkData();
        return CSlicerCache.prototype.getShowCaption.call(this);
    };
    CSlicerData.prototype.getColumnsCount = function() {
        this.checkData();
        return CSlicerCache.prototype.getColumnsCount.call(this);
    };
    CSlicerData.prototype.getButtonHeight = function() {
        this.checkData();
        return CSlicerCache.prototype.getButtonHeight.call(this);
    };
    CSlicerData.prototype.getButtonState = function (nIndex) {
        var oValue = this.getValue(nIndex);
        if(oValue) {
            var nState = 0;
            if(this.getHiddenByOther(oValue) === false) {
                nState |= STATE_FLAG_DATA;
            }
            if(this.getVisible(oValue) !== false) {
                nState |=STATE_FLAG_SELECTED;
            }

            return nState;
        }
        return STYLE_TYPE.WHOLE;
    };
    CSlicerData.prototype.isAllValuesSelected = function () {
        var nCount = this.getValuesCount();
        for(var nValue = 0; nValue < nCount; ++nValue) {
            var oValue = this.getValue(nValue);
            if(oValue && oValue.visible === false) {
                return false;
            }
        }
        return true;
    };
    CSlicerData.prototype.onViewUpdate = function () {
        var oWorksheet = this.slicer.getWorksheet();
        if(!oWorksheet) {
            this.slicer.removeAllButtonsTmpState();
            return;
        }
        if(this.slicer.getLocked()) {
            this.slicer.removeAllButtonsTmpState();
            return;
        }
        var oData = this.retrieveData();
        var aValues = oData.getValues();
        var nValuesCount = aValues.length, nValue, oValue, oApplyValue, nButtonState;
        var aValuesToApply = [];
        var bNeedUpdate = false, bVisible;
        for(nValue = 0; nValue < nValuesCount; ++nValue) {
            oValue = aValues[nValue];
            if(!oValue) {
                break;
            }
            nButtonState = this.slicer.getViewButtonState(nValue);
            if(nButtonState === null) {
                break;
            }
            oApplyValue = oValue.clone();
            bVisible = (nButtonState & STATE_FLAG_SELECTED) !== 0;
            if(this.getVisible(oValue) !== bVisible) {
                oApplyValue.asc_setVisible(bVisible);
                bNeedUpdate = true;
            }
            aValuesToApply.push(oApplyValue);
        }
        if(bNeedUpdate) {
            if(this.slicer.isSubscribed()) {
                this.values = aValuesToApply;
            }
            else {
                Asc.editor.wb.setFilterValuesFromSlicer(this.slicer.getName(), aValuesToApply);
            }
        }
        else {
            this.slicer.removeAllButtonsTmpState();
        }
    };
    CSlicerData.prototype.onDataUpdate = function() {
        if(AscCommon.isFileBuild()) {
           return;
        }
        var oOldCache = this.save();
        this.clear();
        if(this.needUpdateValues(oOldCache)) {
            this.slicer.handleUpdateExtents();
            this.slicer.handleUpdateFill();
            this.slicer.handleUpdateLn();
            this.slicer.recalculate();
        }
        else {
            this.slicer.removeAllButtonsTmpState();
        }
        this.slicer.unsubscribeFromEvents();
        this.slicer.onUpdate(this.slicer.bounds);
    };
    CSlicerData.prototype.needUpdateValues = function(oOldCache) {
        var nCount = this.getValuesCount();
        if(nCount !== oOldCache.getValuesCount()) {
            return true;
        }
        for(var nValue = 0; nValue < nCount; ++nValue) {
            if(this.getString(nValue) !== oOldCache.getString(nValue)) {
                return true;
            }
        }
        if(!AscFormat.fApproxEqual(this.getButtonHeight(), oOldCache.getButtonHeight())) {
            return true;
        }
        if(this.getColumnsCount() !== oOldCache.getColumnsCount()) {
            return true;
        }
        if(this.getShowCaption() !== oOldCache.getShowCaption()) {
            return true;
        }
        if(this.getCaption() !== oOldCache.getCaption()) {
            return true;
        }
        if(this.getStyleName() !== oOldCache.getStyleName()) {
            return true;
        }
        return false;
    };

    function CSlicer() {
        AscFormat.CShape.call(this);
        this.name = null;

        this.recalcInfo.recalculateHeader = true;
        this.recalcInfo.recalculateButtons = true;
        this.header = null;

        this.data = new CSlicerData(this);

        AscFormat.ExecuteNoHistory(function() {
            this.txStyles = new CStyles(false);
        }, this, []);

        this.buttonsContainer = null;

        this.eventListener = null;
    }
    CSlicer.prototype = Object.create(AscFormat.CShape.prototype);
    CSlicer.prototype.constructor = CSlicer;
    CSlicer.prototype.getObjectType = function () {
        return AscDFH.historyitem_type_SlicerView;
    };
    CSlicer.prototype.toStream = function (s) {
        s.WriteUChar(AscCommon.g_nodeAttributeStart);
        s._WriteString2(0, this.name);
        s.WriteUChar(AscCommon.g_nodeAttributeEnd);
    };
    CSlicer.prototype.fromStream = function (s) {
        var _len = s.GetULong();
        var _start_pos = s.cur;
        var _end_pos = _len + _start_pos;
        var _at;
// attributes
        s.GetUChar();
        while (true) {
            _at = s.GetUChar();
            if (_at === AscCommon.g_nodeAttributeEnd)
                break;
            switch (_at) {
                case 0: {
                    this.setName(s.GetString2());
                    break;
                }
                default: {
                    s.Seek2(_end_pos);
                    return;
                }
            }
        }
        s.Seek2(_end_pos);
    };
    CSlicer.prototype.setName = function(val) {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_SlicerViewName, this.name, val));
        this.name = val;
    };
    CSlicer.prototype.getName = function() {
        return this.name;
    };
    CSlicer.prototype.getViewButtonsCount = function () {
        if(!this.buttonsContainer) {
            return 0;
        }
        return this.buttonsContainer.getViewButtonsCount();
    };
    CSlicer.prototype.getViewButtonState = function(nIndex) {
        if(!this.buttonsContainer) {
            return null;
        }
        return this.buttonsContainer.getViewButtonState(nIndex);
    };
    CSlicer.prototype.getStyleName = function() {
        var oWorkbook = this.getWorkbook();
        if(!oWorkbook) {
            return null;
        }
        var oSlicerStyles = oWorkbook.SlicerStyles;
        var sStyleName = oWorkbook.getSlicerStyle(this.name);
        if(!sStyleName) {
            sStyleName = oSlicerStyles.DefaultStyle;
        }
        return sStyleName;
    };
    CSlicer.prototype.getDXF = function(nType) {
        var oWorkbook = this.getWorkbook();
        if(!oWorkbook) {
            return null;
        }
        var sStyleName = this.getStyleName();
        if(!sStyleName) {
            return null;
        }
        var oSlicerStyles = oWorkbook.SlicerStyles;
        var oTableStyles = oWorkbook.TableStyles;
        var oStyle;
        var oDXF = null;
        if(nType === STYLE_TYPE.WHOLE || nType === STYLE_TYPE.HEADER) {
            oStyle = oTableStyles.AllStyles[sStyleName];
            if(!oStyle) {
                oStyle = oTableStyles.AllStyles[oTableStyles.DefaultTableStyle];
            }
            if(oStyle) {
                if(nType === STYLE_TYPE.WHOLE) {
                    if(oStyle.wholeTable) {
                        oDXF = oStyle.wholeTable.dxf;
                    }
                }
                else {
                    if(oStyle.headerRow) {
                        oDXF = oStyle.headerRow.dxf;
                    }
                }
            }
        }
        else {
            oStyle = oSlicerStyles.AllStyles[sStyleName];
            if(!oStyle) {
                oStyle = oSlicerStyles.AllStyles[oSlicerStyles.DefaultStyle];
            }
            if(oStyle) {
                switch (nType) {
                    case STYLE_TYPE.HOVERED_SELECTED_NO_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.hoveredSelectedItemWithNoData];
                        break;
                    }
                    case STYLE_TYPE.HOVERED_UNSELECTED_NO_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.hoveredUnselectedItemWithNoData];
                        break;
                    }
                    case STYLE_TYPE.HOVERED_SELECTED_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.hoveredSelectedItemWithData];
                        break;
                    }
                    case STYLE_TYPE.HOVERED_UNSELECTED_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.hoveredUnselectedItemWithData];
                        break;
                    }
                    case STYLE_TYPE.SELECTED_NO_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.selectedItemWithNoData];
                        break;
                    }
                    case STYLE_TYPE.SELECTED_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.selectedItemWithData];
                        break;
                    }
                    case STYLE_TYPE.UNSELECTED_NO_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.unselectedItemWithNoData];
                        break;
                    }
                    case STYLE_TYPE.UNSELECTED_DATA: {
                        oDXF = oStyle[Asc.ST_slicerStyleType.unselectedItemWithData];
                        break;
                    }
                }
            }
        }
        return oDXF;
    };
    CSlicer.prototype.getFont = function(nType) {
        var oFont = null;
        var oDXF = this.getDXF(nType);
        if(oDXF) {
            oFont = oDXF.getFont();
            if(oFont) {
                return oFont;
            }
        }
        oDXF = this.getDXF(STYLE_TYPE.WHOLE);
        if(oDXF) {
            return oDXF.getFont2();
        }
        return AscCommonExcel.g_oDefaultFormat.Font;
    };
    CSlicer.prototype.getFill = function(nType) {
        var oFill;
        var oDXF = this.getDXF(nType);
        if(oDXF) {
            oFill = oDXF.getFill();
            if(oFill) {
                return oFill;
            }
        }
        if(nType !== STYLE_TYPE.WHOLE) {
            return null;
        }
        oFill = new AscCommonExcel.Fill();
        oFill.fromColor(new AscCommonExcel.RgbColor(0xFFFFFF));
        return oFill;
    };
    CSlicer.prototype.getBorder = function(nType) {
        var oBorder;
        var oDXF = this.getDXF(nType);
        if(oDXF) {
            return oDXF.getBorder();
        }

        oDXF = this.getDXF(STYLE_TYPE.WHOLE);
        if(oDXF) {
            return oDXF.getBorder();
        }
        var r = 91, g = 155, b = 213;
        if(nType !== STYLE_TYPE.HEADER && nType !== STYLE_TYPE.WHOLE) {
            r = 204;
            g = 204;
            b = 204;
        }
        oBorder = new AscCommonExcel.Border();
        oBorder.initDefault();
        if(nType !== STYLE_TYPE.HEADER) {
            oBorder.l = new AscCommonExcel.BorderProp();
            oBorder.l.setStyle(Asc.c_oAscBorderStyles.Thin);
            oBorder.l.c = AscCommonExcel.createRgbColor(r, g, b);
            oBorder.t = new AscCommonExcel.BorderProp();
            oBorder.t.setStyle(Asc.c_oAscBorderStyles.Thin);
            oBorder.t.c = AscCommonExcel.createRgbColor(r, g, b);
            oBorder.r = new AscCommonExcel.BorderProp();
            oBorder.r.setStyle(Asc.c_oAscBorderStyles.Thin);
            oBorder.r.c = AscCommonExcel.createRgbColor(r, g, b);
        }
        oBorder.b = new AscCommonExcel.BorderProp();
        oBorder.b.setStyle(Asc.c_oAscBorderStyles.Thin);
        oBorder.b.c = AscCommonExcel.createRgbColor(r, g, b);
        return oBorder;
    };
    CSlicer.prototype.recalculateBrush = function() {
        var oFill = this.getFill(STYLE_TYPE.WHOLE);
        var oParents = this.getParentObjects();
        this.brush = AscCommonExcel.convertFillToUnifill(oFill);
        this.brush.calculate(oParents.theme, oParents.slide, oParents.layout, oParents.master, {R:0, G:0, B:0, A: 255});
    };
    CSlicer.prototype.recalculatePen = function() {
        this.pen = null;
    };
    CSlicer.prototype.recalculateGeometry = function() {
        this.calcGeometry = AscFormat.CreateGeometry("rect");
        this.calcGeometry.Recalculate(this.extX, this.extY);
    };
    CSlicer.prototype.canRotate = function() {
        return false;
    };
    CSlicer.prototype.recalculate = function () {
        if(AscCommon.isFileBuild()) {
            return;
        }

        //--------------For bug 46500--------------------
        var bCollaborativeChanges = false;
        var oWorkbook = this.getWorkbook();
        if(oWorkbook) {
            bCollaborativeChanges = oWorkbook.bCollaborativeChanges;
        }
        if(AscFonts.IsCheckSymbols && bCollaborativeChanges) {
            return;
        }
        //-------------------------------------------------
        
        AscFormat.ExecuteNoHistory(function () {
            AscFormat.CShape.prototype.recalculate.call(this);
            if(this.recalcInfo.recalculateHeader) {
                this.recalculateHeader();
                this.recalcInfo.recalculateHeader = false;
            }
            if(this.recalcInfo.recalculateButtons) {
                this.recalculateButtons();
                this.recalcInfo.recalculateButtons = false;
            }

        }, this, []);
    };
    CSlicer.prototype.recalculateHeader = function() {
        var bShowHeader = this.getShowCaption();
        if(!bShowHeader) {
            this.header = null;
            return;
        }
        if(!this.header) {
            this.header = new CHeader(this);
        }
        this.header.worksheet = this.worksheet;
        this.header.setRecalculateInfo();
        this.header.recalculate();
    };
    CSlicer.prototype.getButtonsContainerWidth = function() {
        return Math.max(this.extX - LEFT_PADDING - RIGHT_PADDING, 0);
    };
    CSlicer.prototype.getButtonsContainerHeight = function() {
        var nHeight = this.extY;
        if(this.header) {
            nHeight -= this.header.extY;
        }
        return Math.max(nHeight - TOP_PADDING - BOTTOM_PADDING, 0);
    };
    CSlicer.prototype.getWidthByButton = function(dBWidth) {
        if(this.buttonsContainer) {
            var dBContainerW = this.buttonsContainer.getWidthByButton(dBWidth);
            return dBContainerW + LEFT_PADDING + RIGHT_PADDING;
        }
        else {
            return this.extX;
        }
    };
    CSlicer.prototype.recalculateButtons = function() {
        if(!this.buttonsContainer) {
            this.buttonsContainer = new CButtonsContainer(this);
        }
        this.buttonsContainer.clear();
        var nValuesCount = this.getValuesCount();
        for(var nValue = 0; nValue < nValuesCount; ++nValue) {
            this.buttonsContainer.addButton(new CButton(this.buttonsContainer));
        }
        this.buttonsContainer.x = LEFT_PADDING;
        this.buttonsContainer.y = TOP_PADDING;
        if(this.header) {
            this.buttonsContainer.y += this.header.extY;
        }
        this.buttonsContainer.extX = this.getButtonsContainerWidth();
        this.buttonsContainer.extY = this.getButtonsContainerHeight();
        this.buttonsContainer.recalculate();
    };
    CSlicer.prototype.getColumnsCount = function() {
        return this.data.getColumnsCount();
    };
    CSlicer.prototype.getCaption = function() {
        return this.data.getCaption();
    };
    CSlicer.prototype.getButtonHeight = function() {
        return this.data.getButtonHeight();
    };
    CSlicer.prototype.getShowCaption = function() {
        return this.data.getShowCaption();
    };
    CSlicer.prototype.getTxStyles = function (nType) {
        var oFont = this.getFont(nType);
        if(!oFont) {
            oFont = this.getFont(STYLE_TYPE.WHOLE);
        }
        var oTextPr =  this.txStyles.Default.TextPr;
        oTextPr.InitDefault();
        oTextPr.FillFromExcelFont(oFont);
        var oParaPr = this.txStyles.Default.ParaPr;
        oParaPr.SetSpacing(1, undefined, 0, 0, undefined, undefined);
        return {styles: this.txStyles, lastId: undefined};
    };
    CSlicer.prototype.isMultiSelect = function() {
        if(this.header) {
            return this.header.isMultiSelect();
        }
        return false;
    };
    CSlicer.prototype.internalDraw = function(graphics, transform, transformText, pageIndex) {
        var oUR = graphics.updatedRect;
        if(oUR && this.bounds) {
            if(!oUR.isIntersectOther(this.bounds)) {
                return;
            }
        }
        AscFormat.CShape.prototype.draw.call(this, graphics, transform, transformText, pageIndex);
        if(graphics.isBoundsChecker()) {
            return;
        }
        graphics.SaveGrState();
        graphics.transform3(this.transform, false);
        if(this.header) {
            this.header.draw(graphics, transform, transformText, pageIndex);
        }
        if(this.buttonsContainer) {
            this.buttonsContainer.draw(graphics, transform, transformText, pageIndex);
        }
        graphics.RestoreGrState();
        var oBorder = this.getBorder(STYLE_TYPE.WHOLE);
        if(oBorder) {
            var oTransform = transform || this.transform;
            graphics.transform3(oTransform);
            var oSide, oLastDrawn;
            oSide = oBorder.l;
            oLastDrawn = drawVerBorder(graphics, oSide, oLastDrawn, 1, 0, 0, this.extY) || oLastDrawn;
            oSide = oBorder.t;
            oLastDrawn = drawHorBorder(graphics, oSide, oLastDrawn, 1, 0, 0, this.extX) || oLastDrawn;
            oSide = oBorder.r;
            oLastDrawn = drawVerBorder(graphics, oSide, oLastDrawn, 1, this.extX, 0, this.extY) || oLastDrawn;
            oSide = oBorder.b;
            oLastDrawn = drawHorBorder(graphics, oSide, oLastDrawn, 1, this.extY, 0, this.extX) || oLastDrawn;
            graphics.reset();
        }
        if(!AscCommon.IsShapeToImageConverter && !graphics.isPdf()) {
            if(this.getLocked()) {
                var oOldBrush = this.brush;
                this.brush = AscFormat.CreateSolidFillRGBA(0, 0, 0, LOCKED_ALPHA);
                AscFormat.CShape.prototype.draw.call(this, graphics);
                this.brush = oOldBrush;
            }
        }
    };
    CSlicer.prototype.draw = function (graphics, transform, transformText, pageIndex) {
        AscFormat.ExecuteNoHistory(this.internalDraw, this, [graphics, transform, transformText, pageIndex]);
    };
    CSlicer.prototype.handleUpdateExtents = function () {
        this.recalcInfo.recalculateHeader = true;
        this.recalcInfo.recalculateButtons = true;
        AscFormat.CShape.prototype.handleUpdateExtents.call(this);
    };
    CSlicer.prototype.isEventListener = function (child) {
        return this.eventListener === child;
    };
    CSlicer.prototype.setEventListener = function (child) {
        this.eventListener = child;
    };
    CSlicer.prototype.handleClearButtonClick = function () {
        if(!this.buttonsContainer) {
            return;
        }
        this.buttonsContainer.selectAllButtons();
        this.onViewUpdate();
    };
    CSlicer.prototype.onDataUpdate = function() {
        this.data.onDataUpdate();
    };
    CSlicer.prototype.removeAllButtonsTmpState = function() {
        if(this.buttonsContainer) {
            this.buttonsContainer.removeAllButtonsTmpState();
        }
    };
    CSlicer.prototype.onViewUpdate = function() {
        this.data.onViewUpdate();
    };
    CSlicer.prototype.onMouseMove = function (e, x, y) {
        var bRet = false;
        if(this.eventListener) {
            if(!e.IsLocked) {
                return this.onMouseUp(e, x, y);
            }
            if(this.getLocked()) {
                return this.onMouseUp(e, x, y);
            }
            if(e.CtrlKey) {
                if(!this.isSubscribed()) {
                    this.subscribeToEvents();
                }
            }
            else {
                if(this.isSubscribed()) {
                    this.unsubscribeFromEvents();
                }
            }
            this.eventListener.onMouseMove(e, x, y);
            return true;
        }
        if(this.getLocked()) {
            return false;
        }
        if(!this.data.hasData()) {
            return false;
        }
        if(this.header) {
            bRet = bRet || this.header.onMouseMove(e, x, y);
        }
        if(this.buttonsContainer) {
            bRet = bRet || this.buttonsContainer.onMouseMove(e, x, y);
        }
        if(this.hitInInnerArea(x, y)) {
            bRet = true;
        }
        return bRet;
    };
    CSlicer.prototype.onMouseDown = function (e, x, y) {
        if(this.getLocked()) {
            return false;
        }
        if(!this.data.hasData()) {
            return false;
        }
        if(e.button !== 0) {
            return false;
        }
        var bRet = false, bRes;
        e.IsLocked = true;
        if(this.eventListener) {
            this.eventListener.onMouseUp(e, x, y);
        }
        if(this.header) {
            bRes = this.header.onMouseDown(e, x, y);
            bRet = bRet || bRes;
        }
        if(this.buttonsContainer) {
            bRes = this.buttonsContainer.onMouseDown(e, x, y);
            bRet = bRet || bRes;
            if(bRes && e.CtrlKey) {
                this.subscribeToEvents();
            }
        }
        return bRet;
    };
    CSlicer.prototype.getCursorInfo = function (e, x, y) {
        if(this.getLocked() || !this.hit(x, y) || !this.data.hasData()) {
            return null;
        }
        var oRet = (this.header && this.header.getCursorInfo(e, x, y)) || (this.buttonsContainer && this.buttonsContainer.getCursorInfo(e, x, y));
        if(oRet) {
            oRet.objectId = this.Get_Id();
            oRet.bMarker = false;
            return oRet;
        }
        return oRet;
    };
    CSlicer.prototype.onMouseUp = function (e, x, y) {
        var bRet = false;
        if(this.eventListener) {
            bRet = this.eventListener.onMouseUp(e, x, y);
            this.setEventListener(null);
            this.onUpdate(this.bounds);
            this.onViewUpdate();
            return bRet;
        }
        return bRet;
    };
    CSlicer.prototype.getValues = function () {
        return this.data.getValues();
    };
    CSlicer.prototype.getButtonState = function (nIndex) {
        return this.data.getButtonState(nIndex);
    };
    CSlicer.prototype.getValuesCount = function () {
        return this.data.getValuesCount();
    };
    CSlicer.prototype.getString = function (nIndex) {
        return this.data.getString(nIndex);
    };
    CSlicer.prototype.isAllValuesSelected = function () {
        return this.data.isAllValuesSelected();
    };
    CSlicer.prototype.getInvFullTransformMatrix = function () {
        return this.invertTransform;
    };
    CSlicer.prototype.onWheel = function (deltaX, deltaY) {
        if(this.getLocked() || !this.data.hasData()) {
            return false;
        }
        return this.buttonsContainer.onWheel(deltaX, deltaY);
    };
    CSlicer.prototype.onSlicerUpdate = function (sName) {
        if(AscCommon.isFileBuild()) {
            return;
        }
        if(this.name === sName) {
            this.onDataUpdate();
        }
    };
    CSlicer.prototype.onSlicerLock = function (sName, bLock) {
        if(AscCommon.isFileBuild()) {
            return;
        }
        if(this.name === sName) {
            this.data.setLocked(bLock);
            this.onUpdate(this.bounds);
            this.unsubscribeFromEvents();
        }
    };
    CSlicer.prototype.onSlicerChangeName = function (sName, sNewName) {
        if(AscCommon.isFileBuild()) {
            return;
        }
        if(this.name === sName) {
            this.setName(sNewName);
            this.onDataUpdate();
        }
    };
    CSlicer.prototype.onSlicerDelete = function (sName) {
        if(AscCommon.isFileBuild()) {
            return false;
        }
        var bRet = false;
        var oMainGroup, oCurGroup;
        if(this.name === sName) {
            if(this.group) {
                this.group.removeFromSpTree(this.Id);
                oCurGroup = this.group;
                while (oCurGroup.spTree.length === 0 && oCurGroup.group) {
                    oCurGroup.group.removeFromSpTree(oCurGroup.Get_Id());
                    oCurGroup = oCurGroup.group;
                }
                if(oCurGroup.spTree.length === 0) {
                    if(oCurGroup.drawingBase) {
                        oCurGroup.deleteDrawingBase();
                    }
                }
                else {
                    oMainGroup = this.group.getMainGroup();
                    if(oMainGroup) {
                        oMainGroup.updateCoordinatesAfterInternalResize();
                    }
                }
                bRet = true;
            }
            else {
                if(this.drawingBase) {
                    this.deleteDrawingBase();
                    bRet = true;
                }
            }
        }
        if(bRet) {
            this.onUpdate(this.bounds);
            var oController = this.getDrawingObjectsController();
            if(oController) {
                this.deselect(oController);
            }
        }
        return bRet;
    };
    CSlicer.prototype.deleteSlicer = function () {
        Asc.editor.wb.deleteSlicer(this.name);
    };
    CSlicer.prototype.subscribeToEvents = function () {
        var drawingObjects = this.getDrawingObjectsController();
        if(drawingObjects) {
            drawingObjects.addEventListener(this);
        }
    };
    CSlicer.prototype.unsubscribeFromEvents = function () {
        var drawingObjects = this.getDrawingObjectsController();
        if(drawingObjects) {
            drawingObjects.removeEventListener(this);
        }
    };
    CSlicer.prototype.isSubscribed = function () {
        var drawingObjects = this.getDrawingObjectsController();
        if(drawingObjects) {
            return drawingObjects.isEventListener(this);
        }
        return false;
    };
    CSlicer.prototype.onKeyUp = function (e) {
        if(e.keyCode === 91 /*meta*/|| e.keyCode === 17/*ctrl*/) {
            if(this.isSubscribed()) {
                this.unsubscribeFromEvents();
                if(!this.eventListener) {
                    this.onViewUpdate();
                }
            }
        }
    };
    CSlicer.prototype.getLocked = function () {
        return this.data.getLocked() || (this.lockType !== AscCommon.c_oAscLockTypes.kLockTypeNone && this.lockType !== AscCommon.c_oAscLockTypes.kLockTypeMine);
    };
    CSlicer.prototype.copy = function (oPr) {
        var copy = new CSlicer();
        this.fillObject(copy, oPr);
        if(this.name !== null) {
            copy.setName(this.name);
        }
        if(this.macro !== null) {
            copy.setMacro(this.macro);
        }
        if(this.textLink !== null) {
            copy.setTextLink(this.textLink);
        }
        return copy;
    };
    CSlicer.prototype.invertMultiSelect = function () {
        if(this.header) {
            return this.header.invertMultiSelect();
        }
        return false;
    };
    CSlicer.prototype.getSlicer = function() {
        var oWorkbook = this.getWorkbook();
        if(!oWorkbook) {
            return oWorkbook;
        }
        return oWorkbook.getSlicerByName(this.name);
    };
    CSlicer.prototype.getButtonWidth = function() {
        if(!this.buttonsContainer) {
            return 0;
        }
        return this.buttonsContainer.getButtonWidth() * g_dKoef_mm_to_emu + 0.5 >> 0;
    };
    CSlicer.prototype.setButtonWidth = function(dWidth) {
        var bChange = false;
        if(AscFormat.isRealNumber(dWidth) && dWidth > 0) {
            var dCurWidth = this.getButtonWidth();
            if(!AscFormat.fApproxEqual(dCurWidth, dWidth)) {
                var mmWidth = dWidth * g_dKoef_emu_to_mm;
                var dNewWidth = this.getWidthByButton(mmWidth);
                if(!AscFormat.fApproxEqual(dNewWidth, this.extX)) {
                    AscFormat.CheckSpPrXfrm(this);
                    this.spPr.xfrm.setExtX(dNewWidth);
                    this.checkDrawingBaseCoords();
                    bChange = true;
                }
            }
        }
        return bChange;
    };
    CSlicer.prototype.checkAutofit = function (bIgnoreWordShape) {
        return false;
    };
    CSlicer.prototype.checkTextWarp = function(oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
        return oDefaultWrapObject;
    };
    CSlicer.prototype.getSlicerViewByName = function (name) {
        if(this.name === name) {
            return this;
        }
        return null;
    };
    CSlicer.prototype.getAllSlicerViews = function(aSlicerView) {
        aSlicerView.push(this);
    };

    function CHeader(slicer) {
        AscFormat.CShape.call(this);
        this.slicer = slicer;
        this.worksheet = slicer.worksheet;
        this.txBody = null;
        this.buttons = [];
        this.buttons.push(new CInterfaceButton(this));
        this.buttons.push(new CInterfaceButton(this));
        this.buttons[1].removeTmpState();
        this.setBDeleted(false);
        this.setTransformParams(0, 0, 0, 0, 0, false, false);
        this.createTextBody();
        this.bodyPr = new AscFormat.CBodyPr();
        this.bodyPr.setDefault();
        this.bodyPr.anchor = 1;//vertical align ctr
        this.bodyPr.lIns = HEADER_LEFT_PADDING;
        this.bodyPr.rIns = HEADER_RIGHT_PADDING;
        this.bodyPr.tIns = HEADER_TOP_PADDING;
        this.bodyPr.bIns = HEADER_BOTTOM_PADDING;
        this.bodyPr.horzOverflow = AscFormat.nHOTClip;
        this.bodyPr.vertOverflow = AscFormat.nVOTClip;

        this.eventListener = null;
        this.startButton = null;
    }
    CHeader.prototype = Object.create(AscFormat.CShape.prototype);
    CHeader.prototype.getString = function() {
        return this.slicer.getCaption();
    };
    CHeader.prototype.Get_Styles = function() {
        return this.slicer.getTxStyles(STYLE_TYPE.HEADER);
    };
    CHeader.prototype.getParentObjects = function() {
        return this.slicer.getParentObjects();
    };
    CHeader.prototype.isMultiSelect = function() {
        return this.buttons[0].isSelected();
    };
    CHeader.prototype.invertMultiSelect = function () {
        this.buttons[0].setInvertSelectTmpState();
        return true;
    };
    CHeader.prototype.recalculateBrush = function () {
        var oFill = this.slicer.getFill(STYLE_TYPE.HEADER);
        var oParents = this.slicer.getParentObjects();
        this.brush = AscCommonExcel.convertFillToUnifill(oFill);
        this.brush.calculate(oParents.theme, oParents.slide, oParents.layout, oParents.master, {R:0, G:0, B:0, A: 255});
    };
    CHeader.prototype.recalculatePen = function () {
        this.pen = null;
    };
    CHeader.prototype.recalculateContent = function () {
        if(this.bRecalcContent) {
            return;
        }
        this.setTransformParams(0, 0, this.slicer.extX, HEADER_BUTTON_WIDTH, 0, false, false);
        this.recalculateGeometry();
        this.recalculateTransform();
        this.txBody.content.Recalc_AllParagraphs_CompiledPr();
        this.txBody.recalculateOneString(this.getString());
        var dHeight = this.contentHeight + HEADER_TOP_PADDING + HEADER_BOTTOM_PADDING;
        dHeight = Math.max(dHeight, HEADER_BUTTON_WIDTH + 1);
        this.setTransformParams(0, 0, this.slicer.extX, dHeight, 0, false, false);
        this.recalcInfo.recalculateContent = false;
        this.bRecalcContent = true;
        this.recalculate();
        this.recalculateButtons();
        this.bRecalcContent = false;
    };
    CHeader.prototype.getBodyPr = function () {
        return this.bodyPr;
    };
    CHeader.prototype.recalculateGeometry = function() {
        this.calcGeometry = AscFormat.CreateGeometry("rect");
        this.calcGeometry.Recalculate(this.extX, this.extY);
    };
    CHeader.prototype.recalculateButtons = function() {
        var oButton = this.buttons[1];
        var x, y;
        x = this.extX - RIGHT_PADDING - HEADER_BUTTON_WIDTH;
        y = this.extY / 2 - HEADER_BUTTON_WIDTH / 2;
        oButton.setTransformParams(x, y, HEADER_BUTTON_WIDTH, HEADER_BUTTON_WIDTH, 0, false, false);
        oButton.recalculate();
        oButton = this.buttons[0];
        x = this.extX - HEADER_RIGHT_PADDING;
        oButton.setTransformParams(x, y, HEADER_BUTTON_WIDTH, HEADER_BUTTON_WIDTH, 0, false, false);
        oButton.recalculate();
    };
    CHeader.prototype.draw = function (graphics) {
        var oUR = graphics.updatedRect;
        if(oUR && this.bounds) {
            if(!oUR.isIntersectOther(this.bounds)) {
                return;
            }
        }
        AscFormat.CShape.prototype.draw.call(this, graphics);
        if(graphics.isBoundsChecker()) {
            return;
        }

        graphics.SaveGrState();
        graphics.transform3(this.transform);
        graphics.SaveGrState();
        graphics.AddClipRect(0, 0, this.extX, this.extY);
        this.buttons[0].draw(graphics);
        this.buttons[1].draw(graphics);
        graphics.RestoreGrState();
        var oLastDrawn;
        var oDXF = this.slicer.getDXF(STYLE_TYPE.HEADER);
        var oBorder = oDXF && oDXF.getBorder();
        if(oBorder) {
            oLastDrawn = drawVerBorder(graphics, oBorder.l, oLastDrawn, 1, 0, 0, this.extY) || oLastDrawn;
            oLastDrawn = drawHorBorder(graphics, oBorder.t, oLastDrawn, 1, 0, 0, this.extX) || oLastDrawn;
            oLastDrawn = drawVerBorder(graphics, oBorder.r, oLastDrawn, 1, this.extX, 0, this.extY) || oLastDrawn;
            var bFull = false;
            if(oLastDrawn) {
                bFull = true;
            }
            else {
                var oFill = this.slicer.getFill(STYLE_TYPE.HEADER);
                if(oFill) {
                    var oWholeFill = this.slicer.getFill(STYLE_TYPE.WHOLE);
                    if(!oWholeFill || !oFill.isEqual(oWholeFill)) {
                        bFull = true;
                    }
                }
            }
            if(bFull) {
                oLastDrawn = drawHorBorder(graphics, oBorder.b, oLastDrawn, 1, this.extY, 0, this.extX) || oLastDrawn;
            }
            else {
                oLastDrawn = drawHorBorder(graphics, oBorder.b, oLastDrawn, 1, this.extY, LEFT_PADDING, this.slicer.extX - RIGHT_PADDING) || oLastDrawn;
            }
        }
        else {
            oBorder = this.slicer.getBorder(STYLE_TYPE.HEADER);
            if(oBorder) {
                oLastDrawn = drawHorBorder(graphics, oBorder.b, oLastDrawn, 1, this.extY, LEFT_PADDING, this.slicer.extX - RIGHT_PADDING) || oLastDrawn;
            }
        }
        graphics.RestoreGrState();
    };
    CHeader.prototype.getTxStyles = function (nType) {
        return this.slicer.getTxStyles(nType);
    };
    CHeader.prototype.getBorder = function (nType) {
        return null;
    };
    CHeader.prototype.getFill = function (nType) {
        var nColor = HEADER_BUTTON_COLORS[nType];
        var oFill = null;
        if(nColor !== null) {
            oFill = new AscCommonExcel.Fill();
            oFill.fromColor(new AscCommonExcel.RgbColor(nColor));
        }
        return oFill;
    };
    CHeader.prototype.getIcon = function(nIndex, nType) {
        var sRet;
        if(nIndex === 0) {
            if(nType & STATE_FLAG_SELECTED) {
                sRet = ICON_MULTISELECT_INVERTED;
            }
            else {
                sRet = ICON_MULTISELECT;
            }
        }
        else {
            if(nType & STATE_FLAG_WHOLE) {
                sRet = ICON_CLEAR_FILTER_DISABLED;
            }
            else {
                sRet = ICON_CLEAR_FILTER;
            }
        }
        return sRet;
    };
    CHeader.prototype.getFullTransformMatrix = function () {
        return this.transform;
    };
    CHeader.prototype.getInvFullTransformMatrix = function () {
        return this.invertTransform;
    };
    CHeader.prototype.recalculateTransform = function() {
        AscFormat.CShape.prototype.recalculateTransform.call(this);
        var oMT = AscCommon.global_MatrixTransformer;
        this.transform = this.getFullTransform();
        this.invertTransform = oMT.Invert(this.transform);
    };
    CHeader.prototype.recalculateTransformText = function() {
        AscFormat.CShape.prototype.recalculateTransformText.call(this);
        var oMT = AscCommon.global_MatrixTransformer;
        this.transformText = this.getFullTextTransform();
        this.invertTransformText = oMT.Invert(this.transformText);
    };
    CHeader.prototype.recalculateSnapArrays = function() {
    };
    CHeader.prototype.getFullTransform = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        var oTransform = oMT.CreateDublicateM(this.localTransform);

        var oParentTransform = this.slicer.transform;
        oParentTransform && oMT.MultiplyAppend(oTransform, oParentTransform);
        return oTransform;
    };
    CHeader.prototype.getFullTextTransform = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        var oParentTransform = this.slicer.transform;
        var oTransformText = oMT.CreateDublicateM(this.localTransformText);
        oParentTransform && oMT.MultiplyAppend(oTransformText, oParentTransform);
        return oTransformText;
    };
    CHeader.prototype.getInvFullTransformMatrix = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        return oMT.Invert(this.getFullTransform());
    };
    CHeader.prototype.isEventListener = function (child) {
        return this.eventListener === child;
    };
    CHeader.prototype.onMouseMove = function (e, x, y) {
        if(this.eventListener) {
            return this.eventListener.onMouseMove(e, x, y);
        }
        var bRet = false;
        bRet |= this.buttons[0].onMouseMove(e, x, y);
        bRet |= this.buttons[1].onMouseMove(e, x, y);
        return bRet;
    };
    CHeader.prototype.onMouseDown = function (e, x, y) {
        var bRet = false;
        bRet |= this.buttons[0].onMouseDown(e, x, y);
        bRet |= this.buttons[1].onMouseDown(e, x, y);
        return bRet;
    };
    CHeader.prototype.getCursorInfo = function (e, x, y) {
        if(!this.hit(x, y)) {
            return null;
        }
        var oRet = this.buttons[0].getCursorInfo(e, x, y);
        if(oRet) {
            return  oRet;
        }
        oRet = this.buttons[1].getCursorInfo(e, x, y);
        if(!oRet) {
            oRet = {
                cursorType: "move",
                tooltip: !this.txBody.bFit ? this.getString() : null
            }
        }
        return oRet;
    };
    CHeader.prototype.onMouseUp = function (e, x, y) {
        var bRet = false;
        if(this.eventListener) {
            bRet = this.eventListener.onMouseUp(e, x, y);
            this.eventListener = null;
            return bRet;
        }
        bRet |= this.buttons[0].onMouseUp(e, x, y);
        bRet |= this.buttons[1].onMouseUp(e, x, y);
        this.setEventListener(null);
        return bRet;
    };
    CHeader.prototype.getButtonIndex = function (oButton) {
        for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
            if(this.buttons[nButton] === oButton) {
                return nButton;
            }
        }
        return -1;
    };
    CHeader.prototype.setEventListener = function (child) {
        this.eventListener = child;
        if(child) {
            this.slicer.setEventListener(this);
        }
        else {
            if(this.slicer.isEventListener(this)) {
                this.slicer.setEventListener(null);
            }
        }
    };
    CHeader.prototype.handleMouseUp = function (nIndex) {
        var oButton = this.buttons[nIndex];
        if(!oButton) {
            return;
        }
        if(nIndex === 1) {
            this.slicer.handleClearButtonClick();
        }
    };
    CHeader.prototype.handleMouseDown = function (nIndex) {
        var oButton = this.buttons[nIndex];
        if(!oButton) {
            return;
        }
        if(nIndex === 0) {
            oButton.setInvertSelectTmpState();
            this.slicer.onUpdate(oButton.bounds);
        }
    };
    CHeader.prototype.isButtonDisabled = function (nIndex) {
        if(nIndex === 1) {
            return this.slicer.isAllValuesSelected();
        }
        else {
            return false;
        }
    };
    CHeader.prototype.getButtonState = function (nIndex) {
        var oButton = this.buttons[nIndex];
        var nRet = STYLE_TYPE.WHOLE;
        if(!oButton) {
            return nRet
        }
        if(nIndex === 0) {
            if(oButton.haveTmpState()) {
                nRet = oButton.tmpState;
            }
            else {
                nRet = STYLE_TYPE.WHOLE;
            }
        }
        else {
            if(this.slicer.isAllValuesSelected()) {
                nRet = STYLE_TYPE.WHOLE;
            }
            else {
                nRet = STYLE_TYPE.UNSELECTED_DATA;
            }
        }
        return nRet;
    };
    CHeader.prototype.getParentObjects = function () {
        return this.slicer.getParentObjects();
    };
    CHeader.prototype.getScrollOffsetX = function () {
        return 0;
    };
    CHeader.prototype.getScrollOffsetY = function () {
        return 0;
    };
    CHeader.prototype.clipTextRect = function(graphics, transform, transformText, pageIndex) {
    };
    CHeader.prototype.onUpdate = function(oBounds) {
        var oClipBounds = oBounds;
        if(!oClipBounds) {
            oClipBounds = this.bounds;
        }
        this.slicer.onUpdate(oClipBounds);
    };
    CHeader.prototype.getTooltipText = function (nIndex) {
        var sText = null;
        if(nIndex === 0) {
            sText = AscCommon.translateManager.getValue("Multi-Select (Alt+S)");
        }
        else if(nIndex === 1) {
            sText = AscCommon.translateManager.getValue("Clear Filter (Alt+C)");
        }
        return sText;
    };
    CHeader.prototype.checkAutofit = function (bIgnoreWordShape) {
        return false;
    };
    CHeader.prototype.checkTextWarp = function(oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
        return oDefaultWrapObject;
    };
    CHeader.prototype.isForm = function() {
        return false;
    };

    function CButtonBase(parent) {
        AscFormat.CShape.call(this);
        this.parent = parent;
        this.tmpState = null;
        this.worksheet = parent.worksheet;
        this.setBDeleted(false);
        AscFormat.CheckSpPrXfrm3(this);
        this.isHovered = false;
    }
    CButtonBase.prototype = Object.create(AscFormat.CShape.prototype);
    CButtonBase.prototype.getTxBodyType = function () {
        var nRet = null;
        return nRet;
    };
    CButtonBase.prototype.getString = function() {
        return "";
    };
    CButtonBase.prototype.haveTmpState = function() {
        return this.tmpState !== null;
    };
    CButtonBase.prototype.Get_Styles = function() {
        return this.parent.getTxStyles(this.getTxBodyType());
    };
    CButtonBase.prototype.getBodyPr = function () {
        return this.bodyPr;
    };
    CButtonBase.prototype.getFullTransform = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        var oTransform = oMT.CreateDublicateM(this.localTransform);

        var oScrollMatrix = new AscCommon.CMatrix();
        oScrollMatrix.tx = this.parent.getScrollOffsetX();
        oScrollMatrix.ty = this.parent.getScrollOffsetY();
        oMT.MultiplyAppend(oTransform, oScrollMatrix);
        var oParentTransform = this.parent.getFullTransformMatrix();
        oParentTransform && oMT.MultiplyAppend(oTransform, oParentTransform);
        return oTransform;
    };
    CButtonBase.prototype.getFullTextTransform = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        var oParentTransform = this.parent.getFullTransformMatrix();
        var oTransformText = oMT.CreateDublicateM(this.localTransformText);
        var oScrollMatrix = new AscCommon.CMatrix();
        oScrollMatrix.tx = this.parent.getScrollOffsetX();
        oScrollMatrix.ty = this.parent.getScrollOffsetY();
        oMT.MultiplyAppend(oTransformText, oScrollMatrix);
        oParentTransform && oMT.MultiplyAppend(oTransformText, oParentTransform);
        return oTransformText;
    };
    CButtonBase.prototype.getInvFullTransformMatrix = function() {
        var oMT = AscCommon.global_MatrixTransformer;
        return oMT.Invert(this.getFullTransform());
    };
    CButtonBase.prototype.getOwnState = function() {
        return this.parent.getButtonState(this.parent.getButtonIndex(this));
    };
    CButtonBase.prototype.getState = function() {
        var nState = 0;
        if(this.haveTmpState()) {
            nState = this.tmpState;
        }
        else {
            nState = this.getOwnState();
        }
        if(nState !== STYLE_TYPE.WHOLE && nState !== STYLE_TYPE.HEADER) {
            if(this.isHovered) {
                nState |= STATE_FLAG_HOVERED;
            }
            else {
                nState &= (~STATE_FLAG_HOVERED);
            }
        }
        return nState;
    };
    CButtonBase.prototype.setUnselectTmpState = function() {
        this.setTmpState(this.getOwnState() & (~STATE_FLAG_SELECTED));
    };
    CButtonBase.prototype.setSelectTmpState = function() {
        this.setTmpState(this.getOwnState() | STATE_FLAG_SELECTED);
    };
    CButtonBase.prototype.setHoverState = function() {
        var bOld = this.isHovered;
        this.isHovered = true;
        if(bOld !== this.isHovered) {
            this.onUpdate(this.bounds);
        }
    };
    CButtonBase.prototype.getTooltipText = function() {
        return null;
    };
    CButtonBase.prototype.setNotHoverState = function() {
        var bOld = this.isHovered;
        this.isHovered = false;
        if(bOld !== this.isHovered) {
            this.onUpdate(this.bounds);
        }
    };
    CButtonBase.prototype.setInvertSelectTmpState = function() {
        var nOwnState = this.getOwnState();
        if(nOwnState & STATE_FLAG_SELECTED) {
            this.setTmpState(nOwnState & (~STATE_FLAG_SELECTED));
        }
        else {
            this.setTmpState(nOwnState | STATE_FLAG_SELECTED);
        }
    };
    CButtonBase.prototype.setTmpState = function(state) {
        var nOld = this.tmpState;
        this.tmpState = state;
        if(nOld !== this.tmpState) {
            this.onUpdate(this.bounds);
        }
    };
    CButtonBase.prototype.removeTmpState = function() {
        this.setTmpState(null);
    };
    CButtonBase.prototype.isSelected = function() {
        return (this.getState() & STATE_FLAG_SELECTED) !== 0;
    };
    CButtonBase.prototype.recalculate = function() {
        AscFormat.CShape.prototype.recalculate.call(this);
    };
    CButtonBase.prototype.recalculateBrush = function () {
        //Empty procedure. Set of brushes for all states will be recalculated in CSlicer
    };
    CButtonBase.prototype.recalculatePen = function () {
        this.pen = null;
    };
    CButtonBase.prototype.recalculateContent = function () {
    };
    CButtonBase.prototype.recalculateGeometry = function() {
        this.calcGeometry = AscFormat.CreateGeometry("rect");
        this.calcGeometry.Recalculate(this.extX, this.extY);
    };
    CButtonBase.prototype.recalculateTransform = function() {
        AscFormat.CShape.prototype.recalculateTransform.call(this);
        var oMT = AscCommon.global_MatrixTransformer;
        this.transform = this.getFullTransform();
        this.invertTransform = oMT.Invert(this.transform);
    };
    CButtonBase.prototype.recalculateTransformText = function() {
        AscFormat.CShape.prototype.recalculateTransformText.call(this);
        var oMT = AscCommon.global_MatrixTransformer;
        this.transformText = this.getFullTextTransform();
        this.invertTransformText = oMT.Invert(this.transformText);
    };
    CButtonBase.prototype.recalculateBounds = function() {
        this.bounds.x = this.transform.tx;
        this.bounds.y = this.transform.ty;
        this.bounds.l = this.bounds.x;
        this.bounds.t = this.bounds.y;
        this.bounds.r = this.bounds.x + this.extX;
        this.bounds.b = this.bounds.y + this.extY;
        this.bounds.w = this.bounds.r - this.bounds.l;
        this.bounds.h = this.bounds.b - this.bounds.t;
    };
    CButtonBase.prototype.recalculateSnapArrays = function() {
    };
    CButtonBase.prototype.checkAutofit = function (bIgnoreWordShape) {
        return false;
    };
    CButtonBase.prototype.checkTextWarp = function(oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
        return oDefaultWrapObject;
    };
    CButtonBase.prototype.addToRecalculate = function() {
    };
    CButtonBase.prototype.draw = function (graphics) {
        var oUR = graphics.updatedRect;
        if(oUR && this.bounds) {
            if(!oUR.isIntersectOther(this.bounds)) {
                return;
            }
        }
        var parents = this.getParentObjects();
        this.brush = AscCommonExcel.convertFillToUnifill(this.parent.getFill(this.getState()));
        this.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, {R: 0, G: 0, B: 0, A: 255});
        this.recalculateTransform();
        this.recalculateTransformText();
        if(!graphics.isBoundsChecker()) {
            this.recalculateBounds();
        }
        AscFormat.CShape.prototype.draw.call(this, graphics);
        if(graphics.isBoundsChecker()) {
            return;
        }
        var oBorder = this.parent.getBorder(this.getState());
        if(oBorder) {
            graphics.transform3(this.transform);
            var oLastDrawn = null;
            oLastDrawn = drawVerBorder(graphics, oBorder.l, oLastDrawn, 0, 0, 0, this.extY) || oLastDrawn;
            oLastDrawn = drawHorBorder(graphics, oBorder.t, oLastDrawn, 0, 0, 0, this.extX) || oLastDrawn;
            oLastDrawn = drawVerBorder(graphics, oBorder.r, oLastDrawn, 2, this.extX, 0, this.extY) || oLastDrawn;
            oLastDrawn = drawHorBorder(graphics, oBorder.b, oLastDrawn, 2, this.extY, 0, this.extX) || oLastDrawn;
        }
    };
    CButtonBase.prototype.hit = function(x, y) {
        if(!this.parent.hit(x, y)) {
            return false;
        }
        var oInv = this.invertTransform;
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        return tx >= 0 && tx <= this.extX && ty >= 0 && ty <= this.extY;
    };
    CButtonBase.prototype.onMouseMove = function (e, x, y) {
        if(e.IsLocked) {
            return false;
        }
        var bHover = this.hit(x, y);
        var bRet = bHover !== this.isHovered;
        if(bHover) {
            this.setHoverState();
        }
        else {
            this.setNotHoverState();
        }
        return bRet;
    };
    CButtonBase.prototype.onMouseDown = function (e, x, y) {
        if(this.hit(x, y)) {
            this.parent.setEventListener(this);
            return true;
        }
        return false;
    };
    CButtonBase.prototype.onMouseUp = function (e, x, y) {
        this.parent.setEventListener(null);
        return false;
    };
    CButtonBase.prototype.onUpdate = function(oBounds) {
        this.parent.onUpdate(oBounds);
    };
    CButtonBase.prototype.getCursorInfo = function(e, x, y) {
        if(!this.hit(x, y)) {
            return null;
        }
        else {
            return {
                cursorType: "default",
                tooltip: this.getTooltipText()
            }
        }
    };

    function CButton(parent) {
        CButtonBase.call(this, parent);
        this.textBoxes = {};
        this.bodyPr = new AscFormat.CBodyPr();
        this.bodyPr.setDefault();
        this.bodyPr.anchor = 1;//vertical align ctr
        this.bodyPr.lIns = LEFT_PADDING;
        this.bodyPr.rIns = RIGHT_PADDING;
        this.bodyPr.tIns = 0;
        this.bodyPr.bIns = 0;
        this.bodyPr.bIns = 0;
        this.bodyPr.horzOverflow = AscFormat.nHOTClip;
        this.bodyPr.vertOverflow = AscFormat.nVOTClip;
    }
    CButton.prototype = Object.create(CButtonBase.prototype);
    CButton.prototype.getTxBodyType = function () {
        var nRet = null;
        for(var key in this.textBoxes) {
            if(this.textBoxes.hasOwnProperty(key)) {
                if(this.textBoxes[key] === this.txBody) {
                    nRet = parseInt(key);
                    break;
                }
            }
        }
        return nRet;
    };
    CButton.prototype.getString = function() {
        return this.parent.getString(this.parent.getButtonIndex(this));
    };
    CButton.prototype.recalculateContent = function () {
        var sText = this.getString();
        var oFontMap = {};
        var oFont, oFoundFont, oMapFont;

        var oFirstButton = this.parent.buttons[0];
        var bUseFirstButton = (oFirstButton !== this);
        var nType;
        for(var key in STYLE_TYPE) {
            if(STYLE_TYPE.hasOwnProperty(key)) {
                nType = STYLE_TYPE[key];
                if(nType !== STYLE_TYPE.WHOLE && nType !== STYLE_TYPE.HEADER) {
                    oFont = this.parent.getFont(nType);
                    oFoundFont = null;
                    if(oFont) {
                        for(var fontKey in oFontMap) {
                            if(oFontMap.hasOwnProperty(fontKey)) {
                                oMapFont = oFontMap[fontKey];
                                if(oMapFont && oMapFont.isEqual(oFont)) {
                                    oFoundFont = oMapFont;
                                    break;
                                }
                            }
                        }
                    }
                    if(!oFoundFont) {
                        this.createTextBody();
                        this.textBoxes[nType] = this.txBody;
                        if(bUseFirstButton) {
                            this.txBody.content.Content[0].CompiledPr = oFirstButton.textBoxes[nType].content.Content[0].CompiledPr;
                        }
                        this.txBody.recalculateOneString(sText);
                        oFontMap[key] = oFont;
                    }
                    else {
                        this.textBoxes[nType] = this.textBoxes[STYLE_TYPE[fontKey]];
                    }
                }
            }
        }
    };
    CButton.prototype.clipTextRect = function (graphics, transform, transformText, pageIndex) {
    };
    CButton.prototype.draw = function (graphics) {
        this.txBody = this.textBoxes[this.getState()];
        if(!this.txBody) {
            this.txBody = null;
        }
        CButtonBase.prototype.draw.call(this, graphics);
    };
    CButton.prototype.getTooltipText = function() {
        if(!this.txBody || this.txBody.bFit) {
            return null;
        }
        return this.getString();
    };
    CButton.prototype.isForm = function() {
        return false;
    };

    function CInterfaceButton(parent) {
        CButtonBase.call(this, parent);
        this.setTmpState(STYLE_TYPE.UNSELECTED_DATA);
    }
    CInterfaceButton.prototype = Object.create(CButtonBase.prototype);
    CInterfaceButton.prototype.isDisabled = function () {
        return this.parent.isButtonDisabled(this.parent.getButtonIndex(this));
    };
    CInterfaceButton.prototype.hit = function (x, y) {
        if(this.isDisabled()) {
            return false;
        }
        return CButtonBase.prototype.hit.call(this, x, y);
    };
    CInterfaceButton.prototype.onMouseDown = function (e, x, y) {
        if(this.isDisabled()) {
            return false;
        }
        var bRet = CButtonBase.prototype.onMouseDown.call(this, e, x, y);
        if(bRet) {
            this.parent.handleMouseDown(this.parent.getButtonIndex(this));
        }
        return bRet;
    };
    CInterfaceButton.prototype.onMouseUp = function (e, x, y) {
        var bEventListener = this.parent.isEventListener(this);
        var bRet = CButtonBase.prototype.onMouseUp.call(this, e, x, y);
        if(bEventListener) {
            this.parent.handleMouseUp(this.parent.getButtonIndex(this));
        }
        return bRet;
    };
    CInterfaceButton.prototype.draw = function(graphics) {
        var oUR = graphics.updatedRect;
        if(oUR && this.bounds) {
            if(!oUR.isIntersectOther(this.bounds)) {
                return;
            }
        }
        CButtonBase.prototype.draw.call(this, graphics);
        if(graphics.isBoundsChecker()) {
            return;
        }
        var sIcon = this.parent.getIcon(this.parent.getButtonIndex(this), this.getState());
        if(null !== sIcon) {
            graphics.SaveGrState();
            graphics.transform3(this.transform);

            graphics.SetIntegerGrid(true);
            graphics.drawImage(sIcon, 0, 0, this.extX, this.extY, 255, null, null);
            graphics.RestoreGrState();
        }
    };
    CInterfaceButton.prototype.getTooltipText = function() {
        return this.parent.getTooltipText(this.parent.getButtonIndex(this));
    };

    function CButtonsContainer(slicer) {
        this.slicer = slicer;
        this.worksheet = slicer.worksheet;
        this.buttons = [];
        this.x = 0;
        this.y = 0;
        this.extX = 0;
        this.extY = 0;
        this.contentW = 0;
        this.contentH = 0;
        this.scrollTop = 0;
        this.scrollLeft = 0;
        this.scroll = new CScroll(this);

        this.eventListener = null;
        this.startX = 0;
        this.startY = 0;
        this.startButton = -1;
    }
    CButtonsContainer.prototype.getParentObjects = function() {
        return this.slicer.getParentObjects();
    };
    CButtonsContainer.prototype.clear = function() {
        this.buttons.length = 0;
    };
    CButtonsContainer.prototype.addButton = function (oButton) {
        this.buttons.push(oButton);
    };
    CButtonsContainer.prototype.getButtonIndex = function (oButton) {
        for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
            if(this.buttons[nButton] === oButton) {
                return nButton;
            }
        }
        return -1;
    };
    CButtonsContainer.prototype.getViewButtonsCount = function () {
        return this.buttons.length;
    };
    CButtonsContainer.prototype.getViewButtonState = function(nIndex) {
        var oButton = this.getButton(nIndex);
        if(!oButton) {
            return null;
        }
        return oButton.getState();
    };
    CButtonsContainer.prototype.getButton = function (nIndex) {
        if(nIndex > -1 && nIndex < this.buttons.length) {
            return this.buttons[nIndex];
        }
        return null;
    };
    CButtonsContainer.prototype.findButtonIndex = function (x, y) {
        var oInv = this.getInvFullTransformMatrix();
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        ty -= this.getScrollOffsetY();
        var nRow, nRowsCount = this.getRowsCount();
        for(nRow = 0;nRow < nRowsCount; ++nRow) {
            if(ty < this.getRowStart(nRow)) {
                break;
            }
        }
        --nRow;
        if(nRow === -1) {
            return -1
        }
        var nCol, nColsCount = this.getColumnsCount();
        for(nCol = 0; nCol < nColsCount; ++nCol) {
            if(tx < this.getColumnStart(nCol)) {
                break;
            }
        }
        --nCol;
        if(nCol === -1) {
            --nRow;
            if(nRow === -1) {
                return - 1;
            }
            nCol = nColsCount - 1
        }
        var nIndex = nRow * nColsCount;
        if(nRow < nRowsCount - 1) {
            nIndex += nCol;
        }
        else {
            nIndex += Math.min(nCol, this.buttons.length - (nRowsCount - 1)*nColsCount);
        }
        return nIndex;

    };
    CButtonsContainer.prototype.getTxStyles = function (nType) {
        return this.slicer.getTxStyles(nType);
    };
    CButtonsContainer.prototype.getBorder = function (nType) {
        return this.slicer.getBorder(nType);
    };
    CButtonsContainer.prototype.getFill = function (nType) {
        return this.slicer.getFill(nType);
    };
    CButtonsContainer.prototype.getFont = function (nType) {
        return this.slicer.getFont(nType);
    };
    CButtonsContainer.prototype.getButtonHeight = function () {
        return this.slicer.getButtonHeight();
    };
    CButtonsContainer.prototype.getButtonWidth = function () {
        var nColumnCount = this.getColumnsCount();
        var nSpaceCount = nColumnCount - 1;
        var dTotalHeight = this.getTotalHeight();
        var dButtonWidth;
        if(dTotalHeight <= this.extY) {
            dButtonWidth = Math.max(0, this.slicer.getButtonsContainerWidth() - nSpaceCount * SPACE_BETWEEN) / nColumnCount;
        }
        else {
            dButtonWidth = Math.max(0, this.slicer.getButtonsContainerWidth() - this.scroll.getWidth() - SPACE_BETWEEN - nSpaceCount * SPACE_BETWEEN) / nColumnCount;
        }
        return dButtonWidth;
    };
    CButtonsContainer.prototype.getWidthByButton = function (dBWidth) {
        var nColumnCount = this.getColumnsCount();
        var nSpaceCount = nColumnCount - 1;
        var dTotalHeight = this.getTotalHeight();
        var dWidth;
        if(dTotalHeight <= this.extY) {
            dWidth = dBWidth * nColumnCount + nSpaceCount * SPACE_BETWEEN;
        }
        else {
            dWidth =  dBWidth * nColumnCount + this.scroll.getWidth() + SPACE_BETWEEN + nSpaceCount * SPACE_BETWEEN;
        }
        return dWidth;
    };
    CButtonsContainer.prototype.getColumnsCount = function () {
        return this.slicer.getColumnsCount();
    };
    CButtonsContainer.prototype.getRowsCount = function () {
        return ((this.buttons.length - 1) / this.getColumnsCount() >> 0) + 1;
    };
    CButtonsContainer.prototype.getRowsInFrame = function () {
        if(this.buttons.length === 0) {
            return 0;
        }
        var dCount = (this.extY) / (this.getButtonHeight() + SPACE_BETWEEN);
        var nCeil = (dCount >> 0);
        if(dCount - nCeil > 0) {
            ++nCeil;
        }
        return Math.min(this.buttons.length, nCeil);
    };
    CButtonsContainer.prototype.getRowsInFrameFull = function () {
        if(this.buttons.length === 0) {
            return 0;
        }
        var dCount = (this.extY + SPACE_BETWEEN) / (this.getButtonHeight() + SPACE_BETWEEN);
        if(dCount <= 0) {
            return 0;
        }
        return (dCount >> 0);
    };
    CButtonsContainer.prototype.getScrolledRows = function () {
        return this.getRowsCount() - this.getRowsInFrameFull();
    };
    CButtonsContainer.prototype.getTotalHeight = function () {
        var nRowsCount = this.getRowsCount();
        return  this.getButtonHeight() * nRowsCount + SPACE_BETWEEN * (nRowsCount - 1);
    };
    CButtonsContainer.prototype.getColumnStart = function (nColumn) {
        return this.x + (this.getButtonWidth() + SPACE_BETWEEN) * nColumn;
    };
    CButtonsContainer.prototype.getRowStart = function (nRow) {
        return this.y + (this.getButtonHeight() + SPACE_BETWEEN) * nRow;
    };
    CButtonsContainer.prototype.checkScrollTop = function() {
        this.scrollTop = Math.max(0, Math.min(this.scrollTop, this.getScrolledRows()));
    };
    CButtonsContainer.prototype.recalculate = function() {
        var nColumnCount = this.getColumnsCount();
        var dButtonWidth, dButtonHeight;
        dButtonHeight = this.getButtonHeight();
        dButtonWidth = this.getButtonWidth();
        this.checkScrollTop();
        var nColumn, nRow, nButtonIndex, oButton, x ,y;
        for(nButtonIndex = 0; nButtonIndex < this.buttons.length; ++nButtonIndex) {
            nColumn = nButtonIndex % nColumnCount;
            nRow = nButtonIndex / nColumnCount >> 0;
            oButton = this.buttons[nButtonIndex];
            x = this.getColumnStart(nColumn);
            y = this.getRowStart(nRow);
            oButton.setTransformParams(x, y, dButtonWidth, dButtonHeight, 0, false, false);
            oButton.recalculate();
        }
        this.scroll.bVisible = this.getTotalHeight() > this.extY;
        this.scroll.recalculate();
    };
    CButtonsContainer.prototype.draw = function (graphics) {
        if(this.buttons.length > 0) {
            graphics.SaveGrState();
            graphics.transform3(this.slicer.transform);
            graphics.AddClipRect(0, this.y - SPACE_BETWEEN / 2, this.slicer.extX, this.extY + SPACE_BETWEEN / 2);
            this.checkScrollTop();
            var nColumns = this.getColumnsCount();
            var nStart = this.scrollTop * nColumns;
            var nEnd = Math.min(this.buttons.length - 1, nStart + this.getRowsInFrame() * nColumns - 1);
            for(var nButton = nStart; nButton <= nEnd; ++nButton) {
                this.buttons[nButton].draw(graphics);
            }
            graphics.RestoreGrState();
            this.scroll.draw(graphics);
        }
    };
    CButtonsContainer.prototype.getFullTransformMatrix = function () {
        return AscCommon.global_MatrixTransformer.CreateDublicateM(this.slicer.transform);
    };
    CButtonsContainer.prototype.getInvFullTransformMatrix = function () {
        var oM = this.getFullTransformMatrix();
        return AscCommon.global_MatrixTransformer.Invert(oM);
    };
    CButtonsContainer.prototype.hit = function (x, y) {
        var oInv = this.getInvFullTransformMatrix();
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        var bottom = Math.min(this.y + this.extY, this.getRowStart(this.getRowsCount() - 1) + this.getButtonHeight());
        return tx >= this.x && ty >= this.y &&
            tx <= this.x + this.extX && ty <= bottom;
    };
    CButtonsContainer.prototype.isEventListener = function (child) {
        return this.eventListener === child;
    };
    CButtonsContainer.prototype.onScroll = function () {
        var nOldScroll = this.scrollTop;
        this.scrollTop = Math.max(0, Math.min(this.scroll.getScrollTop(), this.getScrolledRows()));
        this.checkScrollTop();
        if(this.scrollTop !== nOldScroll) {
            for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
                var oButton = this.buttons[nButton];
                oButton.recalculateTransform();
                oButton.recalculateTransformText();
                oButton.recalculateBounds();
            }
            this.onUpdate(this.getBounds());
        }
    };
    CButtonsContainer.prototype.getBounds = function () {
        var l, t, r, b;
        var oTransform = this.slicer.transform;
        l = oTransform.TransformPointX(this.x, this.y) - 0.5;
        t = oTransform.TransformPointY(this.x, this.y) - 0.5;
        r = oTransform.TransformPointX(this.x + this.extX, this.y + this.extY) + 0.5;
        b = oTransform.TransformPointY(this.x + this.extX, this.y + this.extY) + 0.5;
        return new AscFormat.CGraphicBounds(l, t, r, b);
    };
    CButtonsContainer.prototype.onUpdate = function (oBounds) {
        this.slicer.onUpdate(oBounds);
    };
    CButtonsContainer.prototype.onMouseMove = function (e, x, y) {
        if(this.eventListener) {
            return this.eventListener.onMouseMove(e, x, y);
        }
        var bRet = false, nButton, nFindButton, nLast;
        if(e.IsLocked) {
            if(this.slicer.isEventListener(this)) {
                bRet = true;
                if(this.startButton > -1) {
                    var oButton = this.getButton(this.startButton);
                    if(oButton) {
                        nFindButton = this.findButtonIndex(x, y);
                        if(this.slicer.isMultiSelect() || e.CtrlKey) {

                            this.removeAllButtonsTmpState();
                            oButton.setHoverState();
                            if(nFindButton < this.startButton) {
                                for(nButton = Math.max(0, nFindButton); nButton <= this.startButton; ++nButton) {
                                    this.buttons[nButton].setInvertSelectTmpState();
                                }
                            }
                            else {
                                nLast = Math.min(nFindButton, this.buttons.length - 1);
                                for(nButton = this.startButton; nButton <= nLast; ++nButton) {
                                    this.buttons[nButton].setInvertSelectTmpState();
                                }
                            }
                        }
                        else {
                            for(nButton = 0; nButton < this.buttons.length; ++nButton) {
                                this.buttons[nButton].setUnselectTmpState();
                            }
                            oButton.setHoverState();
                            if(nFindButton < this.startButton) {
                                for(nButton = Math.max(0, nFindButton); nButton <= this.startButton; ++nButton) {
                                    this.buttons[nButton].setSelectTmpState();
                                }
                            }
                            else {
                                nLast = Math.min(nFindButton, this.buttons.length - 1);
                                for(nButton = this.startButton; nButton <= nLast; ++nButton) {
                                    this.buttons[nButton].setSelectTmpState();
                                }
                            }
                        }

                    }
                }
            }
            else {
                bRet = false;
            }
        }
        else {
            if(this.slicer.isEventListener(this)) {
                bRet = this.slicer.onMouseUp(e, x, y);
            }
            else {
                for(nButton = 0; nButton < this.buttons.length; ++nButton) {
                    bRet |= this.buttons[nButton].onMouseMove(e, x, y);
                }
                bRet |= this.scroll.onMouseMove(e, x, y);
            }
        }
        return bRet;
    };
    CButtonsContainer.prototype.onMouseDown = function (e, x, y) {
        if(this.eventListener) {
            this.onMouseUp(e, x, y);
            if(!this.eventListener) {
                return this.onMouseDown(e, x, y);
            }
        }
        if(this.scroll.onMouseDown(e, x, y)) {
            return true;
        }
        if(this.hit(x, y)) {
            this.slicer.setEventListener(this);
            var oInv = this.getInvFullTransformMatrix();
            this.startX = oInv.TransformPointX(x, y);
            this.startY = oInv.TransformPointY(x, y);
            this.startButton = -1;
            for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
                if(this.buttons[nButton].hit(x, y)) {
                    this.startButton = nButton;
                    break;
                }
            }
            this.onMouseMove(e, x, y);
            return true;
        }
        return false;
    };
    CButtonsContainer.prototype.onMouseUp = function (e, x, y) {
        var bRet = false;
        if(this.eventListener) {
            bRet = this.eventListener.onMouseUp(e, x, y);
            this.setEventListener(null);
            return bRet;
        }
        if(!this.slicer.isEventListener(this)) {
            for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
                bRet |= this.buttons[nButton].onMouseUp(e, x, y);
            }
            bRet |= this.scroll.onMouseUp(e, x, y);
            this.setEventListener(null);
        }
        return bRet;
    };
    CButtonsContainer.prototype.onWheel = function (deltaX, deltaY) {
        return this.scroll.onWheel(deltaX, deltaY);
    };
    CButtonsContainer.prototype.setEventListener = function (child) {
        this.eventListener = child;
        if(child) {
            this.slicer.setEventListener(this);
        }
        else {
            if(this.slicer.isEventListener(this)) {
                this.slicer.setEventListener(null);
                this.removeAllButtonsTmpState();
            }
        }
    };
    CButtonsContainer.prototype.getButtonState = function(nIndex) {
        return this.slicer.getButtonState(nIndex);
    };
    CButtonsContainer.prototype.getString = function(nIndex) {
        return this.slicer.getString(nIndex);
    };
    CButtonsContainer.prototype.getScrollOffsetX = function () {
        return 0;
    };
    CButtonsContainer.prototype.getScrollOffsetY = function () {
        return -(this.getRowStart(this.scrollTop) - this.y);
    };
    CButtonsContainer.prototype.selectAllButtons = function () {
        for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
            this.buttons[nButton].setSelectTmpState();
        }
    };
    CButtonsContainer.prototype.removeAllButtonsTmpState = function () {
        for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
            this.buttons[nButton].removeTmpState();
        }
    };
    CButtonsContainer.prototype.getIcon = function (nIndex, nType) {
        return null;
    };
    CButtonsContainer.prototype.getCursorInfo = function (e, x, y) {
        if(!this.hit(x, y)) {
            return null;
        }
        var oRet;
        for(var i = 0; i < this.buttons.length; ++i) {
            oRet = this.buttons[i].getCursorInfo(e, x, y);
            if(oRet) {
                return oRet;
            }
        }
        return {
            cursorType: "default",
            tooltip: null
        }
    };
    CButtonsContainer.prototype.getTooltipText = function (nIndex) {
        return this.getString(nIndex)
    };

    function CScrollButton(parent, type) {
        CInterfaceButton.call(this, parent);
        this.type = type;
    }
    CScrollButton.prototype = Object.create(CInterfaceButton.prototype);
    CScrollButton.prototype.draw = function (graphics) {
        CInterfaceButton.prototype.draw.call(this, graphics);
        var dInd = this.extX / 4.0;
        var dMid = this.extX / 2.0;

        var nColor = SCROLL_ARROW_COLORS[this.getState()];
        graphics.b_color1((nColor >> 16) & 0xFF, (nColor >> 8) & 0xFF, nColor & 0xFF, 0xFF);
        if(this.type === SCROLL_BUTTON_TYPE_TOP) {
            graphics.SaveGrState();
            graphics.transform3(this.transform);
            graphics._s();
            graphics._m(dInd, dMid);
            graphics._l(dMid, dInd);
            graphics._l(this.extX - dInd, dMid);
            graphics._z();
            graphics.df();
            graphics.RestoreGrState();
        }
        else if(this.type === SCROLL_BUTTON_TYPE_BOTTOM) {
            graphics.SaveGrState();
            graphics.transform3(this.transform);
            graphics._s();
            graphics._m(dInd, dMid);
            graphics._l(dMid, this.extY - dInd);
            graphics._l(this.extX - dInd, dMid);
            graphics._z();
            graphics.df();
            graphics.RestoreGrState();
        }
    };

    function CScroll(parent) {
        this.parent = parent;
        this.extX = 0;
        this.extY = 0;
        this.bVisible = false;
        this.buttons = [];
        this.buttons[0] = new CScrollButton(this, SCROLL_BUTTON_TYPE_TOP);
        this.buttons[1] = new CScrollButton(this, SCROLL_BUTTON_TYPE_BOTTOM);
        this.state = STYLE_TYPE.UNSELECTED_DATA;
        this.timerId = null;

        this.tmpScrollerPos = null;
        this.startScrollerPos = null;
        this.startScrollTop = null;
    }
    CScroll.prototype.getBounds = function() {
        var oTransform = this.getFullTransformMatrix();
        var x = this.getPosX();
        var y = this.getPosY();
        var extX = this.getWidth();
        var extY = this.getHeight();
        var l = oTransform.TransformPointX(x, y);
        var t = oTransform.TransformPointY(x, y);
        var r = oTransform.TransformPointX(x + extX, y + extY);
        var b = oTransform.TransformPointY(x + extX, y + extY);
        return new AscFormat.CGraphicBounds(l, t, r, b);
    };
    CScroll.prototype.getTxStyles = function () {
        return this.parent.getTxStyles();
    };
    CScroll.prototype.getFullTransformMatrix = function () {
        return this.parent.getFullTransformMatrix();
    };
    CScroll.prototype.getInvFullTransformMatrix = function () {
        return this.parent.getInvFullTransformMatrix();
    };
    CScroll.prototype.getFill = function (nType) {
        var oFill = new AscCommonExcel.Fill();
        oFill.fromColor(new AscCommonExcel.RgbColor(SCROLL_COLORS[nType]));
        return oFill;
    };
    CScroll.prototype.getBorder = function(nType) {
        var r, g, b;
        r = 0xCE;
        g = 0xCE;
        b = 0xCE;
        var oBorder = new AscCommonExcel.Border();
        oBorder.initDefault();
        oBorder.l = new AscCommonExcel.BorderProp();
        oBorder.l.setStyle(Asc.c_oAscBorderStyles.Thin);
        oBorder.l.c = AscCommonExcel.createRgbColor(r, g, b);
        oBorder.t = new AscCommonExcel.BorderProp();
        oBorder.t.setStyle(Asc.c_oAscBorderStyles.Thin);
        oBorder.t.c = AscCommonExcel.createRgbColor(r, g, b);
        oBorder.r = new AscCommonExcel.BorderProp();
        oBorder.r.setStyle(Asc.c_oAscBorderStyles.Thin);
        oBorder.r.c = AscCommonExcel.createRgbColor(r, g, b);
        oBorder.b = new AscCommonExcel.BorderProp();
        oBorder.b.setStyle(Asc.c_oAscBorderStyles.Thin);
        oBorder.b.c = AscCommonExcel.createRgbColor(r, g, b);
        return oBorder;
    };
    CScroll.prototype.getPosX = function () {
        return this.parent.x + this.parent.extX - this.getWidth();
    };
    CScroll.prototype.getPosY = function () {
        return this.parent.y;
    };
    CScroll.prototype.getHeight = function() {
        return this.parent.extY;
    };
    CScroll.prototype.getWidth = function() {
        return SCROLL_WIDTH;
    };
    CScroll.prototype.getButtonContainerPosX = function(nIndex) {
        return this.getPosX();
    };
    CScroll.prototype.getButtonContainerPosY = function(nIndex) {
        var dRet = 0;
        if(nIndex === 0) {
            dRet = this.getPosY();
        }
        else {
            dRet = this.getPosY() + this.getHeight() - this.getButtonContainerSize();
        }
        return dRet;
    };
    CScroll.prototype.getButtonContainerSize = function() {
        return this.getWidth();
    };
    CScroll.prototype.getButtonPosX = function (nIndex) {
        return this.getButtonContainerPosX(nIndex) + this.getButtonContainerSize() / 2 - this.getButtonSize() / 2;
    };
    CScroll.prototype.getButtonPosY = function (nIndex) {
        return this.getButtonContainerPosY(nIndex) + this.getButtonContainerSize() / 2 - this.getButtonSize() / 2;
    };
    CScroll.prototype.getButtonSize = function () {
        return this.getScrollerWidth();
    };
    CScroll.prototype.getRailPosX = function () {
        return this.getPosX() + this.getWidth() / 2 - this.getRailWidth() / 2;
    };
    CScroll.prototype.getRailPosY = function () {
        return this.getPosY() + this.getButtonContainerSize();
    };
    CScroll.prototype.getRailHeight = function() {
        return this.getHeight() - 2 * this.getButtonContainerSize();
    };
    CScroll.prototype.getRailWidth = function() {
        return SCROLLER_WIDTH;
    };
    CScroll.prototype.getScrollerX = function() {
        return this.getRailPosX() +  this.getRailWidth() / 2 - this.getScrollerWidth() / 2;
    };
    CScroll.prototype.internalGetRelScrollerY = function(nPos) {
        return (this.getRailHeight() - this.getScrollerHeight()) * (nPos/ (this.parent.getScrolledRows()));
    };
    CScroll.prototype.getRelScrollerY = function() {
        if(this.tmpScrollerPos !== null) {
            return this.tmpScrollerPos;
        }
        return this.internalGetRelScrollerY(this.parent.scrollTop);
    };
    CScroll.prototype.getMaxRelScrollerY = function() {
        return this.internalGetRelScrollerY(this.parent.getScrolledRows());
    };
    CScroll.prototype.getScrollerY = function() {
        return this.getRailPosY() + this.getRelScrollerY();
    };
    CScroll.prototype.getScrollerWidth = function() {
        return this.getRailWidth();
    };
    CScroll.prototype.getScrollTop = function() {
        return this.parent.getScrolledRows() * (this.getScrollerY() - this.getRailPosY()) / (this.getRailHeight() - this.getScrollerHeight()) + 0.5 >> 0;
    };
    CScroll.prototype.getScrollerHeight = function() {
        var dRailH = this.getRailHeight();
        var dMinRailH = dRailH / 4;
        return Math.max(dMinRailH, dRailH * (dRailH / this.parent.getTotalHeight()));
    };
    CScroll.prototype.getState = function() {
        return this.state;
    };
    CScroll.prototype.getString = function() {
        return "";
    };
    CScroll.prototype.getButtonState = function(nIndex) {
        return this.state;
    };
    CScroll.prototype.hit = function(x, y) {
        if(!this.bVisible) {
            return false;
        }
        var oInv = this.parent.getInvFullTransformMatrix();
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        var l = this.getPosX();
        var t = this.getPosY();
        var r = l + this.getWidth();
        var b = t + this.getHeight();
        return tx >= l && tx <= r && ty >= t && ty <= b;
    };
    CScroll.prototype.hitInScroller = function(x, y) {
        if(!this.bVisible) {
            return false;
        }
        var oInv = this.parent.getInvFullTransformMatrix();
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        var l = this.getScrollerX();
        var t = this.getScrollerY();
        var r = l + this.getScrollerWidth();
        var b = t + this.getScrollerHeight();
        return tx >= l && tx <= r && ty >= t && ty <= b;
    };
    CScroll.prototype.draw = function(graphics) {
        if(!this.bVisible) {
            return;
        }
        var oUR = graphics.updatedRect;
        if(oUR) {
            if(!oUR.isIntersectOther(this.getBounds())) {
                return;
            }
        }
        var x, y, extX, extY, oButton;
        oButton = this.buttons[0];
        x = this.getButtonPosX(0);
        y = this.getButtonPosY(0);
        extX = this.getButtonSize();
        extY = this.getButtonSize();
        oButton.setTransformParams(x, y, extX, extY, 0, false, false);
        oButton.recalculate();
        oButton.draw(graphics);
        oButton = this.buttons[1];
        x = this.getButtonPosX(1);
        y = this.getButtonPosY(1);
        oButton.setTransformParams(x, y, extX, extY, 0, false, false);
        oButton.recalculate();
        oButton.draw(graphics);

        x = this.getScrollerX();
        y = this.getScrollerY();
        extX = this.getScrollerWidth();
        extY = this.getScrollerHeight();
        var nColor = SCROLL_COLORS[this.getState()];

        graphics.SaveGrState();
        graphics.transform3(this.parent.getFullTransformMatrix());
        graphics.p_color(0xCE, 0xCE, 0xCE, 0xFF);
        graphics.b_color1((nColor >> 16) & 0xFF, (nColor >> 8) & 0xFF, nColor & 0xFF, 0xFF);
        graphics.rect(x, y, extX, extY);
        graphics.df();

        graphics.p_width(0);
        graphics.drawHorLine(0, y, x, x + extX, 0);
        graphics.drawHorLine(0, y + extY, x, x + extX, 0);
        graphics.drawVerLine(2, x, y, y + extY, 0);
        graphics.drawVerLine(2, x + extX, y, y + extY, 0);
        graphics.RestoreGrState();
    };
    CScroll.prototype.onMouseMove = function (e, x, y) {
        if(!this.bVisible) {
            return false;
        }
        var bRet = false;
        if(this.eventListener) {
            this.eventListener.onMouseMove(e, x, y);
            return true;
        }
        if(this.parent.isEventListener(this)){
            if(this.startScrollerPos === null) {
                this.startScrollerPos = y;
            }
            if(this.startScrollTop === null) {
                this.startScrollTop = this.parent.scrollTop;
            }
            var dy = y - this.startScrollerPos;
            this.setTmpScroll(dy + this.internalGetRelScrollerY(this.startScrollTop));
            return true;
        }
        bRet |= this.buttons[0].onMouseMove(e, x, y);
        bRet |= this.buttons[1].onMouseMove(e, x, y);

        //TODO: Use object for scroller
        var bHit = this.hitInScroller(x, y);
        var nState = this.getState();
        if(nState & STATE_FLAG_HOVERED) {
            if(!bHit) {
                this.setState(nState & (~STATE_FLAG_HOVERED));
                bRet = true;
            }
        }
        else {
            if(bHit) {
                this.setState(nState | STATE_FLAG_HOVERED);
                bRet = true;
            }
        }
        //-----------------------------
        return bRet;
    };
    CScroll.prototype.onMouseDown = function (e, x, y) {
        var bRet = false;
        if(this.hit(x, y)) {
            bRet |= this.buttons[0].onMouseDown(e, x, y);
            bRet |= this.buttons[1].onMouseDown(e, x, y);
            if(!bRet) {
                if(this.hitInScroller(x, y)) {
                    //TODO: Use object for scroller
                    this.startScrollerPos = y;
                    this.startScrollTop = this.parent.scrollTop;
                    this.setState(this.state | STATE_FLAG_SELECTED);
                    this.parent.setEventListener(this);
                    this.parent.onScroll();
                    //-----------------------------
                }
                else {
                    this.parent.setEventListener(this);
                    var oInv = this.parent.getInvFullTransformMatrix();
                    var ty = oInv.TransformPointY(x, y);
                    if(ty < this.getScrollerY()) {
                        this.startScroll(-this.internalGetRelScrollerY(3));
                    }
                    else {
                        this.startScroll(this.internalGetRelScrollerY(3));
                    }
                }
                return true;
            }
        }
        return bRet;
    };
    CScroll.prototype.onMouseUp = function (e, x, y) {
        this.endScroll();
        var bRet = false;
        if(this.eventListener) {
            bRet = this.eventListener.onMouseUp(e, x, y);
            this.eventListener = null;
            return bRet;
        }
        bRet |= this.buttons[0].onMouseUp(e, x, y);
        bRet |= this.buttons[1].onMouseUp(e, x, y);
        this.setEventListener(null);
        return bRet;
    };
    CScroll.prototype.isButtonDisabled = function (nIndex) {
        return !this.bVisible;
    };
    CScroll.prototype.isEventListener = function (child) {
        return this.eventListener === child;
    };
    CScroll.prototype.getButtonIndex = function (oButton) {
        for(var nButton = 0; nButton < this.buttons.length; ++nButton) {
            if(this.buttons[nButton] === oButton) {
                return nButton;
            }
        }
        return -1;
    };
    CScroll.prototype.setEventListener = function (child) {
        this.eventListener = child;
        if(child) {
            this.parent.setEventListener(this);
        }
        else {
            if(this.parent.isEventListener(this)) {
                this.parent.setEventListener(null);
            }
        }
    };
    CScroll.prototype.getParentObjects = function () {
        return this.parent.getParentObjects();
    };
    CScroll.prototype.handleMouseUp = function (nIndex) {
        var oButton  = this.buttons[nIndex];
        if(!oButton) {
            return;
        }
        oButton.setUnselectTmpState();
    };
    CScroll.prototype.handleMouseDown = function (nIndex) {
        var oButton  = this.buttons[nIndex];
        if(!oButton) {
            return;
        }
        oButton.setSelectTmpState();
        if(nIndex === 0) {
            this.startScroll(-this.internalGetRelScrollerY(1));
        }
        else {
            this.startScroll(this.internalGetRelScrollerY(1));
        }
    };
    CScroll.prototype.startScroll = function (step) {
        this.endScroll();
        var oScroll = this;
        oScroll.addScroll(step);
        this.timerId = setInterval(function () {
            oScroll.addScroll(step);
        }, SCROLL_TIMER_INTERVAL);
    };
    CScroll.prototype.addScroll = function (step) {
        this.setTmpScroll(this.getRelScrollerY() + step);
        this.parent.onScroll();
    };
    CScroll.prototype.setTmpScroll = function (val) {
        this.tmpScrollerPos = Math.max(0, Math.min(this.getMaxRelScrollerY(), val));
        this.parent.onScroll();
        this.onUpdate(this.getBounds());
    };
    CScroll.prototype.clearTmpScroll = function () {
        if(this.tmpScrollerPos !== null) {
            this.tmpScrollerPos = null;
            this.parent.onScroll();
            this.onUpdate(this.getBounds());
        }
    };
    CScroll.prototype.endScroll = function () {
        if(this.timerId !== null) {
            clearInterval(this.timerId);
            this.timerId = null;
        }
        this.clearTmpScroll();
        this.setState(this.state & (~STATE_FLAG_SELECTED));
        this.startScrollerPos = null;
        this.startScrollTop = null;
    };
    CScroll.prototype.getScrollOffsetX = function () {
        return 0;
    };
    CScroll.prototype.getScrollOffsetY = function () {
        return 0;
    };
    CScroll.prototype.recalculate = function () {
        this.buttons[0].recalculate();
        this.buttons[1].recalculate();
    };
    CScroll.prototype.onWheel = function (deltaX, deltaY) {
        if(!this.bVisible) {
            return false;
        }
        var delta = deltaY;
        if(Math.abs(deltaX) > Math.abs(deltaY)) {
            delta = deltaX;
        }
        if(delta > 0) {
            this.addScroll(this.internalGetRelScrollerY(3));
        }
        else if(delta < 0) {
            this.addScroll(-this.internalGetRelScrollerY(3));
        }
        this.endScroll();
        return true;
    };
    CScroll.prototype.getIcon = function(nIndex, nType) {
        return null;
    };
    CScroll.prototype.onUpdate = function(oBounds) {
        this.parent.onUpdate(oBounds);
    };
    CScroll.prototype.setState = function(nState) {
        var nOld = this.state;
        this.state = nState;
        if(this.state !== nOld) {
            this.onUpdate(this.getBounds());
        }
    };


    AscDFH.changesFactory[AscDFH.historyitem_TimelineSlicerViewName] = AscDFH.CChangesDrawingsString;
    AscDFH.drawingsChangesMap[AscDFH.historyitem_TimelineSlicerViewName] = function(oClass, value) {
        oClass.tag = value;
    };

    function CTimeslicer() {
        AscFormat.CShape.call(this);
        this.name = null;
    }
    AscFormat.InitClass(CTimeslicer, AscFormat.CShape, AscDFH.historyitem_type_TimelineSlicerView);
    CTimeslicer.prototype.setName = function (val) {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_TimelineSlicerViewName, this.name, val));
        this.name = val;
    };

    CTimeslicer.prototype.copy = function (oPr) {
        const oCopy = new CTimeslicer();
        oCopy.setTag(this.tag);
        return oCopy;
    };
    CTimeslicer.prototype.fromStream = function (s) {
        var _len = s.GetULong();
        var _start_pos = s.cur;
        var _end_pos = _len + _start_pos;
        var _at;
// attributes
        s.GetUChar();
        while (true) {
            _at = s.GetUChar();
            if (_at === AscCommon.g_nodeAttributeEnd)
                break;
            switch (_at) {
                case 0: {
                    this.setName(s.GetString2());
                    break;
                }
                default: {
                    s.Seek2(_end_pos);
                    return;
                }
            }
        }
        s.Seek2(_end_pos);
    };
    CTimeslicer.prototype.toStream = function (s) {
        s.WriteUChar(AscCommon.g_nodeAttributeStart);
        s._WriteString2(0, this.name);
        s.WriteUChar(AscCommon.g_nodeAttributeEnd);
    };
    CTimeslicer.prototype.canSelect = function () {
        return false;
    };
    CTimeslicer.prototype.hit = function () {
        return false;
    };

    CTimeslicer.prototype.onTimeSlicerDelete = function (sName) {
        return CSlicer.prototype.onSlicerDelete.call(this, sName);
    };
    window["AscFormat"] = window["AscFormat"] || {};
    window["AscFormat"].CSlicer = CSlicer;
    window["AscFormat"].CTimeslicer = CTimeslicer;

    window["AscCommonExcel"] = window["AscCommonExcel"] || {};
    window["AscCommonExcel"].getSlicerIconsForLoad = getSlicerIconsForLoad;

})();
