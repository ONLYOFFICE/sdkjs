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

$(function () {
    Asc.spreadsheet_api.prototype._init = function () {
        this._loadModules();
    };
    Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
        callback();
    };
    AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
    };
    AscCommonExcel.WorkbookView.prototype._init = function () {
    };
    AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
        callback(true);
    };
    AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
    };
    AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
    };
    AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
    };
    AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
    };
    AscCommonExcel.WorksheetView.prototype._init = function () {
    };
    AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
    };
    AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
    };
    AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
    };
    AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
    };
    AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
    };
    AscCommonExcel.WorksheetView.prototype.draw = function () {
    };
    AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
    };
    AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
    };
    AscCommonExcel.WorksheetView.prototype.getZoom = function () {
    };
    AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
    };

    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
    };
    AscCommonExcel.WorksheetView.prototype._isLockedCells = function (oFromRange, subType, callback) {
        callback(true);
        return true;
    };
    AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
        callback(true);
    };
    AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
        callback(true);
    };
    AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
    };
    AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
    };

    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
    };
    Asc.ReadDefTableStyles = function () {
    };

    let api = new Asc.spreadsheet_api({
        'id-view': 'editor_sdk'
    });
    api.FontLoader = {
        LoadDocumentFonts: function () {
        }
    };
    window["Asc"]["editor"] = api;
    AscCommon.g_oTableId.init();
    api._onEndLoadSdk();
    api.isOpenOOXInBrowser = false;
    api._openDocument(AscCommon.getEmpty());
    api._openOnClient();
    api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
    api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
        api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
    let wb = api.wbModel;
    wb.handlers.add("getSelectionState", function () {
        return null;
    });
    wb.handlers.add("getLockDefNameManagerStatus", function () {
        return true;
    });
    wb.handlers.add("asc_onConfirmAction", function (test1, callback) {
        callback(true);
    });
    let wsView = api.wb.getWorksheet(0);
    wsView.handlers = api.handlers;
    wsView.objectRender = new AscFormat.DrawingObjects();
    // Initialize global variables and functions for tests
    let ws = api.wbModel.aWorksheets[0];
    const CSerial = window['AscCommonExcel'].CSerial;
    let settings, cSerial, autofillRange, oFromRange;
    const getRange = function(c1, r1, c2, r2) {
        return new window["Asc"].Range(c1, r1, c2, r2);
    };
    const clearData = function(c1, r1, c2, r2) {
        ws.autoFilters.deleteAutoFilter(getRange(0, 0, 0, 0));
        ws.removeRows(r1, r2, false);
        ws.removeCols(c1, c2);
    };
    const autofillData = function(assert, fromRangeTo, expectedData, description) {
        for (let i = fromRangeTo.r1; i <= fromRangeTo.r2; i++) {
            for (let j = fromRangeTo.c1; j <= fromRangeTo.c2; j++) {
                let fromRangeToVal = ws.getCell3(i, j);
                let dataVal = expectedData[i - fromRangeTo.r1][j - fromRangeTo.c1];
                assert.strictEqual(fromRangeToVal.getValue(), dataVal, `${description} Cell: ${fromRangeToVal.getName()}, Value: ${dataVal}`);
            }
        }
    };
    const reverseAutofillData = function(assert, rangeTo, expectedData, description) {
        for (let i = rangeTo.r1; i >= rangeTo.r2; i--) {
            for (let j = rangeTo.c1; j >= rangeTo.c2; j--) {
                let rangeToVal = ws.getCell3(i, j);
                let dataVal = expectedData[Math.abs(i - rangeTo.r1)][Math.abs(j - rangeTo.c1)];
                assert.strictEqual(rangeToVal.getValue(), dataVal, `${description} Cell: ${rangeToVal.getName()}, Value: ${dataVal}`);
            }
        }
    };
    const getFilledData = function(c1, r1, c2, r2, testData) {
        let oFromRange = ws.getRange4(0, 0);
        oFromRange.fillData(testData);
        oFromRange.worksheet.selectionRange.ranges = [getRange(c1, r1, c2, r2)];
        oFromRange.bbox = getRange(c1, r1, c2, r2);
        
        return oFromRange;
    }

    QUnit.module('Serial');
    QUnit.test('Autofill linear progression', function (assert) {
        // Fill data
        const testData = [
          ['1']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings = {
            'step': '1',
            'seriesIn': 'Row',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        }
        // Run AutoFill -> Serial
        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['2', '3', '4', '5', '6']], 'Autofill one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange
        // Fill data
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['3'], ['4'], ['5'], ['6']], 'Autofill one Column');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Autofill growth progression', function (assert) {
        const testData = [
            ['1']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings = {
            'step': '2',
            'seriesIn': 'Row',
            'type': 'Growth',
            'stopValue': '',
            'trend': false
        }

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['2', '4', '8', '16', '32']], 'Autofill one Row');
        clearData(0, 0, 5, 5);
        // Select vertical oFromRange
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        oFromRange = ws.getRange4(0,0);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['4'], ['8'], ['16'], ['32']], 'Autofill one Column');
        clearData(0, 0, 5, 5);
        // Select vertical oFromRange with stopValue
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.stopValue = '10';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['4'], ['8'], [''], ['']], 'Autofill one Column with stopValue = 10');
        clearData(0, 0, 5, 5);
    });
    QUnit.test('Autofill default mode', function (assert) {
        let testData = [
            ['1', '2']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings = {
            'step': '',
            'seriesIn': 'Row',
            'type': 'AutoFill',
            'stopValue': '',
            'trend': false
        }

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(2, 0, 5, 0);
        autofillData(assert, autofillRange, [['3', '4', '5', '6']], 'Autofill one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange
        testData = [
            ['1'],
            ['2']
        ];
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 2, 0, 5);
        autofillData(assert, autofillRange, [['3'], ['4'], ['5'], ['6']], 'Autofill one Column');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Negative cases', function (assert) {
        const testData = [
           ['']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings = {
            'step': '1',
            'seriesIn': 'Row',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        }

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        let autoFillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill Linear progression - one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill Linear progression - one Column');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange Growth
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings.seriesIn = 'Rows';
        settings.type = 'Growth';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill Growth progression - one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange Growth
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.seriesIn = 'Columns';
        settings.type = 'Growth';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill Growth progression - one Column');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange autofill default mode
        oFromRange = getFilledData(0, 0, 5, 0, testData);
        settings.seriesIn = 'Rows';
        settings.type = 'AutoFill';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill default mode - one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange autofill default mode
        oFromRange = getFilledData(0, 0, 0, 5, testData);
        settings.seriesIn = 'Columns';
        settings.type = 'AutoFill';

        cSerial = new CSerial (settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill default mode - one Column');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Autofill horizontal progression multiple cells', function (assert) {
        const testData = [
            ['1'],
            ['2'],
            ['3'],
            ['4'],
            ['5'],
            ['6']
        ];
        console.log('Autofill horizontal progression - multiple cells');
        oFromRange = getFilledData(0, 0, 5, 5, testData);
        settings = {
            'step': '1',
            'seriesIn': 'Rows',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        };
        let cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
    });
    QUnit.test('Autofill vertical progression multiple cells', function (assert) {
        const testData = [
            ['1', '2', '3', '4', '5', '6']
        ];
        console.log('Autofill vertical progression - multiple cells');
        oFromRange = getFilledData(0, 0, 5, 5, testData);
        settings = {
            'step': '1',
            'seriesIn': 'Columns',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        };
        let cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
    });
});
