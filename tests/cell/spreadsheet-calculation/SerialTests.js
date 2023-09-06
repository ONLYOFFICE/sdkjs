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
    let settings, cSerial, autofillRange, oFromRange, expectedData;
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
    const getFilledData = function(c1, r1, c2, r2, testData, oStartRange) {
        let [row, col] = oStartRange;
        let oFromRange = ws.getRange4(row, col);
        oFromRange.fillData(testData);
        oFromRange.worksheet.selectionRange.ranges = [getRange(c1, r1, c2, r2)];
        oFromRange.bbox = getRange(c1, r1, c2, r2);
        
        return oFromRange;
    }

    QUnit.module('Serial');
    QUnit.test('Autofill linear progression - one filled row/column', function (assert) {
        // Fill data
        const testData = [
          ['1']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings = {
            'step': '2',
            'seriesIn': 'Rows',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        }
        // Run AutoFill -> Serial
        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['3', '5', '7', '9', '11']], 'Autofill one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange
        // Fill data
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['3'], ['5'], ['7'], ['9'], ['11']], 'Autofill one Column');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange with stopValue
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.stopValue = '7';
        settings.seriesIn = 'Rows';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['3', '5', '7', '', '']], 'Autofill one Row with stopValue = 7');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange with stopValue
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.stopValue = '7';
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['3'], ['5'], ['7'], [''], ['']], 'Autofill one Column with stopValue = 7');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange with trend step
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.trend = true;
        settings.seriesIn = 'Rows';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['2', '3', '4', '5', '6']], 'Autofill one Row with trend step');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange with trend step
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['3'], ['4'], ['5'], ['6']], 'Autofill one Column with trend step');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Autofill growth progression - one filled row/column', function (assert) {
        const testData = [
            ['1']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings = {
            'step': '2',
            'seriesIn': 'Rows',
            'type': 'Growth',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['2', '4', '8', '16', '32']], 'Autofill one Row');
        clearData(0, 0, 5, 5);
        // Select vertical oFromRange
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        oFromRange = ws.getRange4(0,0);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['4'], ['8'], ['16'], ['32']], 'Autofill one Column');
        clearData(0, 0, 5, 5);
        // Select vertical oFromRange with stopValue
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.stopValue = '10';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['2'], ['4'], ['8'], [''], ['']], 'Autofill one Column with stopValue = 10');
        clearData(0, 0, 5, 5);
        // Select horizontal oFromRange with stopValue
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.stopValue = '10';
        settings.seriesIn = 'Rows';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['2', '4', '8', '', '']], 'Autofill one Row with stopValue = 10');
        clearData(0, 0, 5, 0);
        // Select horizontal oFromRange with trend step
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.trend = true;
        settings.seriesIn = 'Rows';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['1', '1', '1', '1', '1']], 'Autofill one Row with trend step');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange with trend step
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['1'], ['1'], ['1'], ['1'], ['1']], 'Autofill one Column with trend step');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Autofill default mode', function (assert) {
        let testData = [
            ['1', '2']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings = {
            'step': '',
            'seriesIn': 'Rows',
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
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 2, 0, 5);
        autofillData(assert, autofillRange, [['3'], ['4'], ['5'], ['6']], 'Autofill one Column');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Negative cases', function (assert) {
        const testData = [
           ['']
        ];
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings = {
            'step': '1',
            'seriesIn': 'Rows',
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
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill Linear progression - one Column');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange Growth
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.seriesIn = 'Rows';
        settings.type = 'Growth';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill Growth progression - one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange Growth
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';
        settings.type = 'Growth';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill Growth progression - one Column');
        clearData(0, 0, 0, 5);
        // Select horizontal oFromRange autofill default mode
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.seriesIn = 'Rows';
        settings.type = 'AutoFill';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill default mode - one Row');
        clearData(0, 0, 5, 0);
        // Select vertical oFromRange autofill default mode
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.seriesIn = 'Columns';
        settings.type = 'AutoFill';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autoFillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autoFillRange, [[''], [''], [''], [''], ['']], 'Autofill default mode - one Column');
        clearData(0, 0, 0, 5);
    });
    QUnit.test('Autofill horizontal progression - multiple filled cells', function (assert) {
        const testData = [
            ['1'],
            ['2'],
            ['3'],
            ['4'],
            ['5'],
            ['6']
        ];
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings = {
            'step': '1',
            'seriesIn': 'Rows',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['2', '3', '4', '5', '6'],
            ['3', '4', '5', '6', '7'],
            ['4', '5', '6', '7', '8'],
            ['5', '6', '7', '8', '9'],
            ['6', '7', '8', '9', '10'],
            ['7', '8', '9', '10', '11']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression');
        clearData(0, 0, 5, 5);
        // Growth progression
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.type = 'Growth';
        settings.step = '2';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['2', '4', '8', '16', '32'],
            ['4', '8', '16', '32', '64'],
            ['6', '12', '24', '48', '96'],
            ['8', '16', '32', '64', '128'],
            ['10', '20', '40', '80', '160'],
            ['12', '24', '48', '96', '192']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression');
        clearData(0, 0, 5, 5);
        // Linear progression with stop value
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.stopValue = '7';
        settings.step = '2';
        settings.type = 'Linear';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['3', '5', '7', '', ''],
            ['4', '6', '', '', ''],
            ['5', '7', '', '', ''],
            ['6', '', '', '', ''],
            ['7', '', '', '', ''],
            ['', '', '', '', '']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression with stop value');
        clearData(0, 0, 5, 5);
        // Growth progression with stop value
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.stopValue = '16';
        settings.type = 'Growth';
        settings.step = '2';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['2', '4', '8', '16', ''],
            ['4', '8', '16', '', ''],
            ['6', '12', '', '', ''],
            ['8', '16', '', '', ''],
            ['10', '', '', '', ''],
            ['12', '', '', '', '']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with stop value');
        clearData(0, 0, 5, 5);
        // Linear progression with trend
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.trend = true;
        settings.type = 'Linear';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['2', '3', '4', '5', '6'],
            ['3', '4', '5', '6', '7'],
            ['4', '5', '6', '7', '8'],
            ['5', '6', '7', '8', '9'],
            ['6', '7', '8', '9', '10'],
            ['7', '8', '9', '10', '11']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression with trend');
        clearData(0, 0, 5, 5);
        // Growth progression with trend
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.trend = true;
        settings.type = 'Growth';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 5);
        expectedData = [
            ['1', '1', '1', '1', '1'],
            ['2', '2', '2', '2', '2'],
            ['3', '3', '3', '3', '3'],
            ['4', '4', '4', '4', '4'],
            ['5', '5', '5', '5', '5'],
            ['6', '6', '6', '6', '6']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with trend');
        clearData(0, 0, 5, 5);
    });
    QUnit.test('Autofill vertical progression - multiple filled cells', function (assert) {
        const testData = [
            ['1', '2', '3', '4', '5', '6']
        ];
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings = {
            'step': '1',
            'seriesIn': 'Columns',
            'type': 'Linear',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['2', '3', '4', '5', '6', '7'],
            ['3', '4', '5', '6', '7', '8'],
            ['4', '5', '6', '7', '8', '9'],
            ['5', '6', '7', '8', '9', '10'],
            ['6', '7', '8', '9', '10', '11']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression');
        clearData(0, 0, 5, 5);
        // Growth progression
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.type = 'Growth';
        settings.step = '2';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['2', '4', '6', '8', '10', '12'],
            ['4', '8', '12', '16', '20', '24'],
            ['8', '16', '24', '32', '40', '48'],
            ['16', '32', '48', '64', '80', '96'],
            ['32', '64', '96', '128', '160', '192']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression');
        clearData(0, 0, 5, 5);
        // Linear progression with stop value
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.stopValue = '10';
        settings.step = '2';
        settings.type = 'Linear';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['3', '4', '5', '6', '7', '8'],
            ['5', '6', '7', '8', '9', '10'],
            ['7', '8', '9', '10', '', ''],
            ['9', '10', '', '', '', ''],
            ['', '', '', '', '', '']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression with stop value');
        clearData(0, 0, 5, 5);
        // Growth progression with stop value
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.stopValue = '32';
        settings.type = 'Growth';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['2', '4', '6', '8', '10', '12'],
            ['4', '8', '12', '16', '20', '24'],
            ['8', '16', '24', '32', '', ''],
            ['16', '32', '', '', '', ''],
            ['32', '', '', '', '', '']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with stop value');
        clearData(0, 0, 5, 5);
        // Linear progression with trend
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.trend = true;
        settings.type = 'Linear';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['2', '3', '4', '5', '6', '7'],
            ['3', '4', '5', '6', '7', '8'],
            ['4', '5', '6', '7', '8', '9'],
            ['5', '6', '7', '8', '9', '10'],
            ['6', '7', '8', '9', '10', '11']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression with trend');
        clearData(0, 0, 5, 5);
        // Growth progression with trend
        oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
        settings.trend = true;
        settings.type = 'Growth';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 5, 5);
        expectedData = [
            ['1', '2', '3', '4', '5', '6'],
            ['1', '2', '3', '4', '5', '6'],
            ['1', '2', '3', '4', '5', '6'],
            ['1', '2', '3', '4', '5', '6'],
            ['1', '2', '3', '4', '5', '6']
        ];
        autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with trend');
        clearData(0, 0, 5, 5);
    });
    QUnit.test('Autofill Date type - one filled row/column', function (assert) {
        const testData = [
            ['09/04/2023']
        ];
        // Horizontal dateUnit - Day
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings = {
            'step': '1',
            'seriesIn': 'Rows',
            'type': 'Date',
            'dateUnit': 'Day',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['45174', '45175', '45176', '45177', '45178']], 'Autofill Row. Date progression - Day');
        clearData(0, 0, 5, 0);
        // Horizontal dateUnit - Weekday
        oFromRange = getFilledData(0, 0, 7, 0, testData, [0, 0]);
        settings.dateUnit = 'Weekday';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 7, 0);
        autofillData(assert, autofillRange, [['45174', '45175', '45176', '45177', '45180', '45181', '45182']], 'Autofill Row. Date progression - Weekday');
        clearData(0, 0, 7, 0);
        // Horizontal dateUnit - Month, Step - 2
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.dateUnit = 'Month';
        settings.step = '2';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['45234', '45295', '45355', '45416', '45477']], 'Autofill Row. Date progression - Month, Step - 2');
        clearData(0, 0, 5, 0);
        // Horizontal dateUnit - Year, Step - 1
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.dateUnit = 'Year';
        settings.step = '1';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['45539', '45904', '46269', '46634', '47000']], 'Autofill Row. Date progression - Year, Step - 1');
        clearData(0, 0, 5, 0);
        // Vertical dateUnit - Day
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.dateUnit = 'Day';
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], ['45177'], ['45178']], 'Autofill Column. Date progression - Day');
        clearData(0, 0, 0, 5);
        // Vertical dateUnit - Weekday
        oFromRange = getFilledData(0, 0, 0, 7, testData, [0, 0]);
        settings.dateUnit = 'Weekday';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 7);
        autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], ['45177'], ['45180'], ['45181'], ['45182']], 'Autofill Column. Date progression - Weekday');
        clearData(0, 0, 0, 7);
        // Vertical dateUnit - Month, Step - 2
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.dateUnit = 'Month';
        settings.step = '2';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['45234'], ['45295'], ['45355'], ['45416'], ['45477']], 'Autofill Column. Date progression - Month, Step - 2');
        clearData(0, 0, 0, 5);
        // Vertical dateUnit - Year, Step - 1
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.dateUnit = 'Year';
        settings.step = '1';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['45539'], ['45904'], ['46269'], ['46634'], ['47000']], 'Autofill Column. Date progression - Year, Step - 1');
        clearData(0, 0, 0, 5);
        // Horizontal dateUnit - Day, Stop value - 45176
        oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
        settings.dateUnit = 'Day';
        settings.stopValue = '45176';
        settings.seriesIn = 'Rows';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 0);
        autofillData(assert, autofillRange, [['45174', '45175', '45176', '', '']], 'Autofill Row. Date progression - Day, Stop value - 45176');
        clearData(0, 0, 5, 0);
        // Vertical dateUnit - Day, Stop value - 45176
        oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
        settings.dateUnit = 'Day';
        settings.stopValue = '45176';
        settings.seriesIn = 'Columns';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 0, 5);
        autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], [''], ['']], 'Autofill Column. Date progression - Day, Stop value - 45176');
    });
    QUnit.test('Autofill Date type - Horizontal multiple cells', function (assert) {
        const testData = [
            ['01/01/2023'],
            ['09/04/2023'],
            ['01/12/2023'],
            ['12/12/2023']
        ];
        // DateUnit - Day. Step - 3
        oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
        settings = {
            'step': '3',
            'seriesIn': 'Rows',
            'type': 'Date',
            'dateUnit': 'Day',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 3);
        expectedData = [
           ['44930', '44933', '44936', '44939', '44942', '44945'],
           ['45176', '45179', '45182', '45185', '45188', '45191'],
           ['44941', '44944', '44947', '44950', '44953', '44956'],
           ['45275', '45278', '45281', '45284', '45287', '45290']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Day, Step - 3');
        clearData(0, 0, 5, 3);
        // DateUnit - Weekday
        oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
        settings.dateUnit = 'Weekday';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 3);
        expectedData = [
           ['44930', '44935', '44938', '44942', '44945'],
           ['45176', '45180', '45183', '45187', '45190'],
           ['44942', '44945', '44949', '44952', '44956'],
           ['45275', '45278', '45281', '45285', '45288']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Weekday, Step - 3');
        clearData(0, 0, 5, 3);
        // DateUnit - Month
        oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
        settings.dateUnit = 'Month';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 3);
        expectedData = [
           ['45017', '45108', '45200', '45292', '45383'],
           ['45264', '45355', '45447', '45539', '45630'],
           ['45028', '45119', '45211', '45303', '45394'],
           ['45363', '45455', '45547', '45638', '45728']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Month, Step - 3');
        clearData(0, 0, 5, 3);
        // DateUnit - Year
        oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
        settings.dateUnit = 'Year';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(1, 0, 5, 3);
        expectedData = [
            ['46023', '47119', '48214', '49310', '50406'],
            ['46269', '47365', '48461', '49556', '50652'],
            ['46034', '47130', '48225', '49321', '50417'],
            ['46368', '47464', '48560', '49655', '50751']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Year, Step - 3');
        clearData(0, 0, 5, 3);
    });
    QUnit.test('Autofill Date type - Vertical multiple cells', function (assert) {
       const testData = [
           ['01/01/2023', '09/04/2023', '01/12/2023', '12/12/2023']
       ];
       // DateUnit - Day. Step - 3
        oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
        settings = {
            'step': '3',
            'seriesIn': 'Columns',
            'type': 'Date',
            'dateUnit': 'Day',
            'stopValue': '',
            'trend': false
        };

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 3, 5);
        expectedData = [
            ['44930', '45176', '44941', '45275'],
            ['44933', '45179', '44944', '45278'],
            ['44936', '45182', '44947', '45281'],
            ['44939', '45185', '44950', '45284'],
            ['44942', '45188', '44953', '45287'],
            ['44945', '45191', '44956', '45290']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Day, Step - 3');
        clearData(0, 0, 3, 5);
        // DateUnit - Weekday
        oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
        settings.dateUnit = 'Weekday';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 3, 5);
        expectedData = [
            ['44930', '45176', '44942', '45275'],
            ['44935', '45180', '44945', '45278'],
            ['44938', '45183', '44949', '45281'],
            ['44942', '45187', '44952', '45285'],
            ['44945', '45190', '44956', '45288']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Weekday, Step - 3');
        clearData(0, 0, 3, 5);
        // DateUnit - Month
        oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
        settings.dateUnit = 'Month';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 3, 5);
        expectedData = [
            ['45017', '45264', '45028', '45363'],
            ['45108', '45355', '45119', '45455'],
            ['45200', '45447', '45211', '45547'],
            ['45292', '45539', '45303', '45638'],
            ['45383', '45630', '45394', '45728']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Month, Step - 3');
        clearData(0, 0, 3, 5);
        // DateUnit - Year
        oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
        settings.dateUnit = 'Year';

        cSerial = new CSerial(settings, oFromRange);
        cSerial.exec();
        autofillRange = getRange(0, 1, 3, 5);
        expectedData = [
            ['46023', '46269', '46034', '46368'],
            ['47119', '47365', '47130', '47464'],
            ['48214', '48461', '48225', '48560'],
            ['49310', '49556', '49321', '49655'],
            ['50406', '50652', '50417', '50751']
        ];
        autofillData(assert, autofillRange, expectedData, 'Date progression - Year, Step - 3');
        clearData(0, 0, 3, 5);
    });
});
