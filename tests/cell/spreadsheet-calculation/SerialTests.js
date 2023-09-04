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
            'seriesIn': 'Row',
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
            'seriesIn': 'Row',
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
        let expectedData = [
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
        let expectedData = [
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
});
