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

$(function () {
    // Mocks necessary methods for work api and disable collaborative methods.
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
    //Goal Seek
    let CGoalSeek = AscCommonExcel.CGoalSeek;
    let CParserFormula = AscCommonExcel.parserFormula;
    let g_oIdCounter = AscCommon.g_oIdCounter;
    let sData = AscCommon.getEmpty();
    let wb, ws, oParserFormula, oGoalSeek, nResult, nChangingVal, nExpectedVal, api;
    // Solver
    let CSolver = AscCommonExcel.CSolver;
    let asc_CSolverParams = AscCommonExcel.asc_CSolverParams;
    let c_oAscOperator = AscCommonExcel.c_oAscOperator;
    let c_oAscDerivativeType = AscCommonExcel.c_oAscDerivativeType;
    let c_oAscSolvingMethod = AscCommonExcel.c_oAscSolvingMethod;
    let c_oAscOptimizeTo = AscCommonExcel.c_oAscOptimizeTo;
    let oSolver;

    if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
        Asc.spreadsheet_api.prototype._init = function() {
        };
        api = new Asc.spreadsheet_api({
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
        api.OpenDocumentFromBin(null, AscCommon.getEmpty());
        api.initCollaborativeEditing({});
        api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
            api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
        wb = api.wbModel;
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

        let docInfo = new Asc.asc_CDocInfo();
        docInfo.asc_putTitle("TeSt.xlsx");
        api.DocInfo = docInfo;
        api.initCollaborativeEditing({});
        window["Asc"]["editor"] = api;

        AscCommon.g_oTableId.init();
        /*wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);*/
        AscCommon.History.init(wb);
        wb.maxDigitWidth = 7;
        wb.paddingPlusBorder = 5;

        if (this.User) {
            g_oIdCounter.Set_UserId(this.User.asc_getId());
        }

        g_oIdCounter.Set_Load(false);

        var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
        oBinaryFileReader.Read(sData, wb);
        ws = wb.getWorksheet(wb.getActive());
        AscCommonExcel.getFormulasInfo();
    }
    //wb.dependencyFormulas.lockRecal();
    // Goal Seek
    CGoalSeek.prototype.resume = function() {
        this.setIsPause(false);
        while (true) {
            let bIsFinish = this.calculate();
            if (bIsFinish) {
                break;
            }
        }
    };
    CGoalSeek.prototype.step = function() {
        this.setIsPause(false);
        this.setIsSingleStep(true);
        while (true) {
            let bIsFinish = this.calculate();
            if (bIsFinish) {
                break;
            }
        }
    };

    const getRange = function (c1, r1, c2, r2) {
        return new window["Asc"].Range(c1, r1, c2, r2);
    };
    const clearData = function (c1, r1, c2, r2) {
        ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
        ws.removeRows(r1, r2, false);
        ws.removeCols(c1, c2);
    };
    const getResult = function (nExpectedVal, oChangingCell, sFormula, sFormulaCell) {
        let nResult, nChangingVal;
        // Init objects ParserFormula and GoalSeek
        oParserFormula = new CParserFormula(sFormula, sFormulaCell, ws);
        oGoalSeek = new CGoalSeek(oParserFormula, nExpectedVal, oChangingCell);
        oGoalSeek.init();
        // Run goal seek
        while (true) {
            let bIsFinish = oGoalSeek.calculate();
            if (bIsFinish) {
                break;
            }
        }
        // Update data for formula
        oParserFormula.parse();
        // Get results and changing value
        nResult = Number(oParserFormula.calculate().getValue());
        nChangingVal = Number(oGoalSeek.getChangingCell().getValue());

        return [nResult, nChangingVal];
    };
    const pauseGoalSeek = function (nExpectedVal, oChangingCell, sFormula, sFormulaCell) {
        let oParserFormula, oGoalSeek;
        // Init objects ParserFormula and GoalSeek
        oParserFormula = new CParserFormula(sFormula, sFormulaCell, ws);
        oGoalSeek = new CGoalSeek(oParserFormula, nExpectedVal, ws.getRange4(0, 0));
        oGoalSeek.init();
        // Run goal seek
        while (true) {
            let bIsFinish = oGoalSeek.calculate();
            oGoalSeek.pause();
            if (bIsFinish) {
                break;
            }
        }

        return [oParserFormula, oGoalSeek];
    };
    const getResultOutOfLoop = function(oParserFormula, oGoalSeek) {
        let nResult, nChangingVal;

        oParserFormula.parse()
        nResult = Number(oParserFormula.calculate().getValue());
        nChangingVal = Number(oGoalSeek.getChangingCell().getValue());

        return [nResult, nChangingVal];
    };
    // Solver
    const getHiddenDefinedNamesWS = function(oDependencyFormulas) {
        const oActiveWS = oDependencyFormulas.wb.getActiveWs();
        const aHiddenDefNames = [];

        oDependencyFormulas._foreachDefNameSheet(oActiveWS.getId(), function(oDefName) {
            if (oDefName.hidden) {
                aHiddenDefNames.push(oDefName);
            }
        });

        return aHiddenDefNames;
    };
    // Mock the method that starts the solver calculation. Replaces async loop to sync loop
    AscCommonExcel.WorkbookView.prototype.startSolver = function(oSolverParams) {
        if (!this.model) {
            return;
        }

        const CSolver = AscCommonExcel.CSolver;
        const wbModel = this.model;
        const ws = wbModel.getActiveWs();
        const t = this;
        let oSolver;

        const callback = function (isSuccess) {
            if (!isSuccess) {
                t.handlers.trigger("asc_onError", c_oAscError.ID.LockedCellGoalSeek, c_oAscError.Level.NoCritical);
                return;
            }
            //open history point
            History.Create_NewPoint();
            History.StartTransaction();
            // Init CSolver object
            wbModel.setSolver(new CSolver(oSolverParams, ws))
            oSolver = wbModel.getSolver();
            oSolver.prepare();
            // Run solver
            if (oSolverParams.getOptions().getShowIterResults()) {
                oSolver.step();
            } else {
                while (true) {
                    let bIsFinish = oSolver.calculate();
                    if (bIsFinish) {
                        break;
                    }
                }
            }
        };

        //need lock
        const aLocksInfo = [];
        //cells locks info
        const wsChangingCell = AscCommonExcel.actualWsByRef(oSolverParams.getChangingCells(), ws);
        const sChangingCell = AscCommonExcel.convertToAbsoluteRef(oSolverParams.getChangingCells());
        const oChangingCell = wsChangingCell && wsChangingCell.getRange2(sChangingCell);
        if (oChangingCell) {
            aLocksInfo.push(this.collaborativeEditing.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Range, null, wsChangingCell.getId(),
                new AscCommonExcel.asc_CCollaborativeRange(oChangingCell.bbox.c1, oChangingCell.bbox.r1, oChangingCell.bbox.c2, oChangingCell.bbox.r2)));
        }

        const wsFormula = AscCommonExcel.actualWsByRef(oSolverParams.getObjectiveFunction(), ws);
        const sFormulaCell = AscCommonExcel.convertToAbsoluteRef(oSolverParams.getObjectiveFunction());
        const oFormulaCell = wsFormula && wsFormula.getRange2(sFormulaCell);
        if (oFormulaCell) {
            aLocksInfo.push(this.collaborativeEditing.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Range, null, wsFormula.getId(),
                new AscCommonExcel.asc_CCollaborativeRange(oFormulaCell.bbox.c1, oFormulaCell.bbox.r1, oFormulaCell.bbox.c2, oFormulaCell.bbox.r2)));
        }

        this.collaborativeEditing.lock(aLocksInfo, callback);

        const oChangedCell = oSolver.getChangingCell();
        if (oChangedCell) {
            // update worksheet field
            let ws = this.getWorksheetById(oChangedCell.worksheet.Id);
            ws._updateRange(oChangedCell.bbox);
            ws.draw();
        }
    };
    CSolver.prototype.step = function() {
        let oSolver = this;

        this.setIsPause(false);
        this.setIsSingleStep(true);
        while (true) {
            let bIsFinish = oSolver.calculate();
            if (bIsFinish) {
                break;
            }
        }
    };
    const checkUndoRedo = function(fBefore, fAfter, desc) {
        fAfter("after_" + desc);
        AscCommon.History.Undo();
        fBefore("undo_" + desc);
        AscCommon.History.Redo();
        fAfter("redo_" + desc);
        AscCommon.History.Undo();
    };

    QUnit.module('Goal seek');
    QUnit.test('PMT formula', function (assert) {
        const aTestData = [
            ['', '180', '100000'],
            ['0.072', '1', '100000'],
            ['0.072', '180', '0'],
            ['0.01', '180', '100000'],
            ['0.072', '10', '100000'],
            ['0.072', '180', '100'],
            ['-0.10', '180', '100000'],
            ['0.072', '-10', '100000'],
            ['0.072', '180', '-100'],
            ['1', '180', '100000'],
            ['0.072', '200', '100000'],
            ['0.072', '180', '200000']

        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find "Interest rate" parameter
        let nExpectedVal = -900;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'PMT(A1/12,B1,C1)', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit term in month" parameter
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'PMT(A2/12,B2,C2)', 'D2');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit term in month" for PMT formula. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 180, `Case: Find "Credit term in month" for PMT formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit sum" parameter
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 2), 'PMT(A3/12,B3,C3)', 'D3');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find "Credit sum" for PMT formula. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Trying to find "Interest rate" parameter with 0.01 as changing value
        nExpectedVal = -900;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), 'PMT(A4/12,B4,C4)', 'D4');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula with 0.01 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula with 0.01 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit term in month" parameter with 10 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 1), 'PMT(A5/12,B5,C5)', 'D5');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit term in month" for PMT formula with 10 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 180, `Case: Find "Credit term in month" for PMT formula with 10 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit sum" parameter with 100 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 2), 'PMT(A6/12,B6,C6)', 'D6');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula with 100 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find "Credit sum" for PMT formula with 100 as changing value. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Trying to find "Interest rate" parameter with -0.10 as changing value
        nExpectedVal = -900;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 0), 'PMT(A7/12,B7,C7)', 'D7');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula with -0.10 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula with -0.10 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit term in month" parameter with -10 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 1), 'PMT(A8/12,B8,C8)', 'D8');
        assert.strictEqual(nResult, 0, `Negative case: Find "Credit term in month" for PMT formula with -10 as changing value. Resolve not found. Result PMT: ${nResult}`);
        // Trying to find "Credit sum" parameter with -100 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 2), 'PMT(A9/12,B9,C9)', 'D9');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula with -100 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find "Credit sum" for PMT formula with -100 as changing value. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Trying to find "Interest rate" parameter with 1 as changing value
        nExpectedVal = -900;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), 'PMT(A10/12,B10,C10)', 'D10');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula with 1 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula with 1 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit term in month" parameter with 200 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 1), 'PMT(A11/12,B11,C11)', 'D11');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit term in month" for PMT formula with 200 as changing value. Resolve not found. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 180, `Case: Find "Credit term in month" for PMT formula with 200 as changing value. Resolve not found. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "Credit sum" parameter with 200000 as changing value
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(11, 2), 'PMT(A12/12,B12,C12)', 'D12');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula with 200000 as changing value. Result PMT: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find "Credit sum" for PMT formula with 200000 as changing value. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Clear data
        clearData(0, 0, 3, 11);
    });
    QUnit.test('Custom formula. S = v * t', function (assert) {
        const aTestData = [
            ['', '5'],
            ['2000', ''],
            ['5', '5'],
            ['2000', '-1'],
            ['3000', '5'],
            ['2000', '10'],
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find "time" parameter for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1*B1', 'D1');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "time" for custom formula. Result: ${nResult}`);
        assert.strictEqual(nChangingVal, 2000, `Case: Find "time" for custom formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "speed" parameter for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'A2*B2', 'D2');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "speed" for custom formula. Result: ${nResult}`);
        assert.strictEqual(nChangingVal, 5, `Case: Find "speed" for custom formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "time" parameter with 5 as changing value for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), 'A3*B3', 'D3');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "time" for custom formula with 5 as changing value. Result: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2000, `Case: Find "time" for custom formula with 5 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "speed" parameter with -1 as changing value for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 1), 'A4*B4', 'D4');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "speed" for custom formula with -1 as changing value. Result: ${nResult}`);
        assert.strictEqual(nChangingVal, 5, `Case: Find "speed" for custom formula with -1 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "time" parameter with 3000 as changing value for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 0), 'A5*B5', 'D5');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "time" for custom formula with 3000 as changing value. Result: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2000, `Case: Find "time" for custom formula with 3000 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "speed" parameter with 10 as changing value for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 1), 'A6*B6', 'D6');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "speed" for custom formula with 10 as changing value. Result: ${nResult}`);
        assert.strictEqual(nChangingVal, 5, `Case: Find "speed" for custom formula with 10 as changing value. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 3, 3);
    });
    QUnit.test('Custom formula. Arithmetical operations', function (assert) {
        const aTestData = [
            ['', '5'],
            ['5', ''],
            ['', '2', '4'],
            ['1', '', '4'],
            ['1', '2', ''],
            ['', '10'],
            ['2', ''],
            ['', '2'],
            ['-10', ''],
            ['', '25', '5'],
            ['8', '', '5'],
            ['', ''],
            ['0', '', '123', '1', '-1233'],
            ['2'],
            [''],
            ['', '5'],
            ['3', ''],
            ['', '8'],
            ['4', ''],
            ['', '9'],
            ['5', ''],
            ['', '4'],
            ['6', ''],
            ['5', '5'],
            ['7','5'],
            ['','2'],
            ['8', ''],
            ['', '7'],
            ['9', '-1'],
            ['1', '1'],
            ['10', '5'],
            ['', '3'],
            ['5', '5'],
            ['5', '-3'],
            ['9', '2', '4'],
            ['1', '-2', '4'],
            ['1', '2', '0.1'],
            ['1', '2'],
            ['-10', '-2'],
            ['8', '25', '1'],
            ['8', '25', '12'],
            ['8', '25', '-2'],
            ['0', '2', '123', '1', '-1233'],
            ['0', '-10', '123', '1', '-1233'],
            ['0', '200000', '123', '1', '-1233'],
            ['1', '2']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find first parameter for formula x = a + b
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1+B1', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a+b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find first parameter for a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a + b
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'A2+B2', 'D2');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a+b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find second parameter for a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), '(A3+B3)*C3+3', 'D3');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for (a+b)*c+3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 1, `Case: Find first parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 1), '(A4+B4)*C4+3', 'D4');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for (a+b)*c+3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find third parameter for formula x = (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 2), '(A5+B5)*C5+3', 'D5');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find third parameter for (a+b)*c+3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 4, `Case: Find third parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = a^b
        nExpectedVal = 1024;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 0), 'A6^B6', 'D6');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find first parameter for a^b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a^b
        nExpectedVal = 1024;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 1), 'A7^B7', 'D7');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 10, `Case: Find second parameter for a^b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = -a*b
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 0), 'A8*B8', 'D8');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for -a*b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), -10, `Case: Find first parameter for -a*b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = -a*b
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 1), 'A9*B9', 'D9');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for -a*b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter for -a*b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = a + b / c
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), 'A10+B10/C10', 'D10');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a+b/c. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 8, `Case: Find first parameter for a+b/c. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a + b / c
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 1), 'A11+B11/C11', 'D11');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a+b/c. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 25, `Case: Find second parameter for a+b/c. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula a + 0
        nExpectedVal = 100000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(11, 0), 'A12+0', 'D12');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a+0. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find first parameter for a+0. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula a + b - c + d + e
        nExpectedVal = 100000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(12, 1), 'A13+B13-C13+D13+E13', 'F13');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a+b-c+d+e. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 101355, `Case: Find second parameter for a+b-c+d+e. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter for formula 20*a-20/a
        nExpectedVal = 25;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(13, 0), '20*A14-20/A14', 'D14');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find parameter for 20*a-20/a. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(2)), 1.80, `Case: Find parameter for 20*a-20/a. Result ChangingVal: ${nChangingVal}`);
        //Trying to find parameter for formula SQRT(SQRT(a) + SQRT(a))
        nExpectedVal = 12;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(14, 0), 'SQRT(SQRT(A15) + SQRT(A15))', 'D15');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find parameter for SQRT(SQRT(a) + SQRT(a)). Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 5184, `Case: Find parameter for SQRT(SQRT(a) + SQRT(a)). Result ChangingVal: ${nChangingVal}`);
        //Trying to find first parameter for a^b (3^5) formula
        nExpectedVal = 243;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(15, 0), 'A16^B16', 'D16');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (3^5). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 3, `Case: Find first parameter for a^b (3^5). Result ChangingVal: ${nChangingVal}`);
        //Trying to find second parameter for a^b (3^5) formula
        nExpectedVal = 243;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(16, 1), 'A17^B17', 'D17');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (3^5). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find second parameter for a^b (3^5). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (4^8)
        nExpectedVal = 65536;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(17, 0), 'A18^B18', 'D18');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (4^8). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 4, `Case: Find first parameter for a^b (4^8). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for a^b (4^8)
        nExpectedVal = 65536;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(18, 1), 'A19^B19', 'D19');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (4^8). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 8, `Case: Find second parameter for a^b (4^8). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (5^9)
        nExpectedVal = 1953125;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(19, 0), 'A20^B20', 'D20');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (5^9). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find first parameter for a^b (5^9). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for a^b (5^9)
        nExpectedVal = 1953125;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(20, 1), 'A21^B21', 'D21');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (5^9). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 9, `Case: Find second parameter for a^b (5^9). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (6^4)
        nExpectedVal = 1296;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(21, 0), 'A22^B22', 'D22');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (6^4). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 6, `Case: Find first parameter for a^b (6^4). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for a^b (6^4)
        nExpectedVal = 1296;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(22, 1), 'A23^B23', 'D23');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (6^4). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 4, `Case: Find second parameter for a^b (6^4). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter with 5 as changing value for a^b (7^5)
        nExpectedVal = 16807;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(23, 0), 'A24^B24', 'D24');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter with 5 as changing value for a^b (7^5). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 7, `Case: Find first parameter with 5 as changing value for a^b (7^5). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for a^b (7^5)
        nExpectedVal = 16807;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(24, 1), 'A25^B25', 'D25');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (7^5). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find second parameter for a^b (7^5). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (8^2)
        nExpectedVal = 64;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(25, 0), 'A26^B26', 'D26');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (8^2). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 8, `Case: Find first parameter for a^b (8^2). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for a^b (8^2)
        nExpectedVal = 64;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(26, 1), 'A27^B27', 'D27');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b (8^2). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter for a^b (8^2). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (9^7)
        nExpectedVal = 4782969;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(27, 0), 'A28^B28', 'D28');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (9^7). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 9, `Case: Find first parameter for a^b (9^7). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with -1 as changing value for a^b (9^7`)
        nExpectedVal = 4782969;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(28, 1), 'A29^B29', 'D29');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with -1 as changing value for a^b (9^7). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 7, `Case: Find second parameter with -1 as changing value for a^b (9^7). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (10^1)
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(29, 0), 'A30^B30', 'D30');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (10^1). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 10, `Case: Find first parameter for a^b (10^1). Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with 5 as changing value for a^b (10^1)
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(30, 1), 'A31^B31', 'D31');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with 5 as changing value for a^b (10^1). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 1, `Case: Find second parameter with 5 as changing value for a^b (10^1). Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for a^b (1024^3)
        nExpectedVal = 1073741824;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(31, 0), 'A32^B32', 'D32');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b (1024^3). Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 1024, `Case: Find first parameter for a^b (1024^3). Result ChangingVal: ${nChangingVal}`);
        // Cases with start changed value
        // Trying to find first parameter with 5 as changed value for formula a+b
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(32, 0), 'A33+B33', 'D33');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter with 5 as changed value for formula a+b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find first parameter with 5 as changed value for formula a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with -3 as changed value for formula a+b
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(33, 1), 'A34+B34', 'D34');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with -3 as changed value for formula a+b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find second parameter with -3 as changed value for formula a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter with 9 as changed value for formula x = (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(34, 0), '(A35+B35)*C35+3', 'D35');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter with 9 as changed value for formula (a + b) * c + 3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 1, `Case: Find first parameter with 9 as changed value for formula (a + b) * c + 3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with -2 as changed value for formula (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(35, 1), '(A36+B36)*C36+3', 'D36');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with -2 as changed value for formula (a + b) * c + 3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter with -2 as changed value for formula (a + b) * c + 3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find third parameter with 0.1 as changed value for formula (a + b) * c + 3
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(36, 2), '(A37+B37)*C37+3', 'D37');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find third parameter with 0.1 as changed value for formula (a + b) * c + 3. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 4, `Case: Find third parameter with 0.1 as changed value for formula (a + b) * c + 3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter with 1 as changed value for formula -a*b
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(37, 0), 'A38*B38', 'D38');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter with 1 as changed value for formula -a*b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), -10, `Case: Find first parameter with 1 as changed value for formula -a*b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with -2 as changed value for formula -a*b
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(38, 1), 'A39*B39', 'D39');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with -2 as changed value for formula -a*b. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter with -2 as changed value for formula -a*b. Result ChangingVal: ${nChangingVal}`);
        //Trying to find third parameter with 1 as changed value for formula a + b / c
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(39, 2), 'A40+B40/C40', 'D40');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find third parameter with 1 as changed value for formula a + b / c. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find third parameter with 1 as changed value for formula a + b / c. Result ChangingVal: ${nChangingVal}`);
        // Trying to find third parameter with 12 as changed value for formula a + b / c
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(40, 2), 'A41+B41/C41', 'D41');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find third parameter with 12 as changed value for formula a + b / c. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find third parameter with 12 as changed value for formula a + b / c. Result ChangingVal: ${nChangingVal}`);
        // Trying to find third parameter with -2 as changed value for formula a + b / c
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(41, 2), 'A42+B42/C42', 'D42');
        assert.strictEqual(Math.round(nResult), 8, `Negative case: Find third parameter with -2 as changed value for formula a + b / c. Result not found. Result formula: ${nResult}`);
        // Trying to find second parameter with 2 as changed value for formula a + b - c + d + e
        nExpectedVal = 100000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(42, 1), 'A43+B43-C43+D43+E43', 'F43');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with 2 as changed value for a+b-c+d+e. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 101355, `Case: Find second parameter with 2 as changed value for a+b-c+d+e. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with -10 as changed value for formula a + b - c + d + e
        nExpectedVal = 100000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(43, 1), 'A44+B44-C44+D44+E44', 'F44');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with -10 as changed value for a+b-c+d+e. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 101355, `Case: Find second parameter with -10 as changed value for a+b-c+d+e. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter with 200000 as changed value for formula a + b - c + d + e
        nExpectedVal = 100000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(44, 1), 'A45+B45-C45+D45+E45', 'F45');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter with 200000 as changed value for a+b-c+d+e. Result formula: ${nResult}`);
        assert.strictEqual(Math.round(nChangingVal), 101355, `Case: Find second parameter with 200000 as changed value for a+b-c+d+e. Result ChangingVal: ${nChangingVal}`);
        nExpectedVal = 5.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(45, 0), 'A46*B46', 'D46');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find first parameter with 1 as changed value for formula a*b. Result formula: ${nResult}`);
        assert.strictEqual(nChangingVal, 2.75, `Case: Find first parameter with 1 as changed value for formula a*b. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0,0, 3, 45);
    });
    QUnit.test('Financials calculation', function (assert) {
       const aTestData = [
           ['', '10', '0.1', '2.59374246'],
           ['', '0.1' ],
           ['10', ''],
           ['', '41', '228'],
           ['1000', '0', '228'],
           ['1000', '41', '1'],
           ['5', '10', '0.1', '2.59374246'],
           ['-3', '0.1' ],
           ['10', '1'],
           ['1', '41', '228'],
           ['1000', '1', '228'],
           ['1000', '41', '2'],
           ['-2', '41', '228'],
           ['1000', '-2', '228'],
           ['1000', '41', '50'],
           ['2000', '41', '228'],
           ['1000', '100', '228'],
           ['1000', '41', '320'],
       ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Try to find "start investment" parameter for formula a*d
        nExpectedVal = 500000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1*D1', 'E1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "start investment" parameter for formula a*d. Result formula: ${nResult.toFixed(9)}`);
        assert.strictEqual(Number(nChangingVal.toFixed(1)), 192771.6, `Case: Find "start investment" parameter for formula a*d. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Try to find "term" parameter for formula (1+a)^b
        nExpectedVal = 2.59;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 0), '(1+B2)^A2', 'E2');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "term" parameter for formula (1+a)^b. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 10, `Case: Find "term" parameter for formula (1+a)^b. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Income" parameter for formula (1+a)^b
        nExpectedVal = 2.59;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 1), '(1+B3)^A3', 'E3');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Income" parameter for formula (1+a)^b. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(2)), 0.10, `Case: Find "Income" parameter for formula (1+a)^b. Result ChangingVal: ${nChangingVal}`);
        // Try to find first parameter  for formula (((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), '(((A4*12)*(60-B4))*2)/C4', 'E4');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find first parameter  for formula ((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 1000, `Case: Find first parameter  for formula ((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find second parameter  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 1), '(((A5*12)*(60-B5))*2)/C5', 'E5');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find second parameter  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 41, `Case: Find second parameter  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find third parameter  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 2), '(((A6*12)*(60-B6))*2)/C6', 'E6');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find third parameter  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 228, `Case: Find third parameter  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find "start investment" parameter with 5 as changed value for formula a*d
        nExpectedVal = 500000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 0), 'A7*D7', 'E7');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "start investment" parameter with 5 as changed value for formula a*d. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(1)), 192771.6, `Case: Find "start investment" parameter with 5 as changed value for formula a*d. Result ChangingVal: ${nChangingVal}`);
        // Try to find "term" parameter with -3 as changed value for formula (1+a)^b
        nExpectedVal = 2.59;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 0), '(1+B8)^A8', 'E8');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "term" parameter with -3 as changed value for formula (1+a)^b. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 10, `Case: Find "term" parameter with -3 as changed value for formula (1+a)^b. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Income" parameter with 1 as changed value for formula (1+a)^b
        nExpectedVal = 2.59;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 1), '(1+B9)^A9', 'E9');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Income" parameter with 1 as changed value for formula (1+a)^b. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(2)), 0.10, `Case: Find "Income" parameter with 1 as changed value for formula (1+a)^b. Result ChangingVal: ${nChangingVal}`);
        // Try to find first parameter with 1 as changed value  for formula (((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), '(((A10*12)*(60-B10))*2)/C10', 'E10');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find first parameter with 1 as changed value  for formula ((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 1000, `Case: Find first parameter with 1 as changed value  for formula ((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find second parameter with 1 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 1), '(((A11*12)*(60-B11))*2)/C11', 'E11');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find second parameter with 1 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 41, `Case: Find second parameter with 1 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find third parameter with 2 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(11, 2), '(((A12*12)*(60-B12))*2)/C12', 'E12');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find third parameter with 2 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 228, `Case: Find third parameter with 2 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find first parameter with -2 as changed value  for formula (((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(12, 0), '(((A13*12)*(60-B13))*2)/C13', 'E13');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find first parameter with -2 as changed value  for formula ((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 1000, `Case: Find first parameter with -2 as changed value  for formula ((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find second parameter with -2 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(13, 1), '(((A14*12)*(60-B14))*2)/C14', 'E14');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find second parameter with -2 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 41, `Case: Find second parameter with -2 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find third parameter with 50 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(14, 2), '(((A15*12)*(60-B15))*2)/C15', 'E15');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find third parameter with 50 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 228, `Case: Find third parameter with 50 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find first parameter with 2000 as changed value  for formula (((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(15, 0), '(((A16*12)*(60-B16))*2)/C16', 'E16');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find first parameter with 2000 as changed value  for formula ((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 1000, `Case: Find first parameter with 2000 as changed value  for formula ((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find second parameter with 100 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(16, 1), '(((A17*12)*(60-B17))*2)/C17', 'E17');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find second parameter with 100 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 41, `Case: Find second parameter with 100 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Try to find third parameter with 320 as changed value  for formula ((a * 12) * (60 - b)) * 2) / c
        nExpectedVal = 2000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(17, 2), '(((A18*12)*(60-B18))*2)/C18', 'E18');
        assert.strictEqual(Number(nResult.toFixed()), nExpectedVal, `Case: Find third parameter with 320 as changed value  for formula (((a*12)*(60-b))*2)/c. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 228, `Case: Find third parameter with 320 as changed value  for formula (((a*12)*(60-b))*2)/c. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 5, 17);
    });
    QUnit.test('FV Formula', function (assert) {
        const aTestData = [
            ['', '12', '-1000'],
            ['0.1230', '0', '-1000'],
            ['0.1230', '12', ''],
            ['0.1', '12', '-1000'],
            ['0.1230', '0.5', '-1000'],
            ['0.1230', '12', '1'],
            ['-1', '12', '-1000'],
            ['0.1230', '-1', '-1000'],
            ['0.1230', '12', '-1'],
            ['1', '12', '-1000'],
            ['0.1230', '15', '-1000'],
            ['0.1230', '12', '-2000']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Try to find "Rate" parameter for FV formula
        nExpectedVal = 12700;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'FV(A1/12,B1,C1)', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Rate" parameter for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.1230, `Case: Find "Rate" parameter for FV formula. Result ChangingVal: ${nChangingVal}`);
        //Try to find "Count of payments" parameter for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'FV(A2/12,B2,C2)', 'D2');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Count of payments" parameter for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 12, `Case: Find "Count of payments" parameter for FV formula. Result ChangingVal: ${nChangingVal}`);
        //Try to find "Payment" parameter for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 2), 'FV(A3/12,B3,C3)', 'D3');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Payment" parameter for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), -1000, `Case: Find "Payment" parameter for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Rate" parameter with 0.1 as changed value for FV formula
        nExpectedVal = 12700;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), 'FV(A4/12,B4,C4)', 'D4');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Rate" parameter with 0.1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.1230, `Case: Find "Rate" parameter with 0.1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Count of payments" parameter with 0.5 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 1), 'FV(A5/12,B5,C5)', 'D5');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Count of payments" parameter with 0.5 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 12, `Case: Find "Count of payments" parameter with 0.5 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Payment" parameter with 1 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 2), 'FV(A6/12,B6,C6)', 'D6');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Payment" parameter with 1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), -1000, `Case: Find "Payment" parameter with 1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Rate" parameter with -1 as changed value for FV formula
        nExpectedVal = 12700;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 0), 'FV(A7/12,B7,C7)', 'D7');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Rate" parameter with -1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.1230, `Case: Find "Rate" parameter with -1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Count of payments" parameter with -1 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 1), 'FV(A8/12,B8,C8)', 'D8');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Count of payments" parameter with -1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 12, `Case: Find "Count of payments" parameter with -1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Payment" parameter with -1 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 2), 'FV(A9/12,B9,C9)', 'D9');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Payment" parameter with -1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), -1000, `Case: Find "Payment" parameter with -1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Rate" parameter with 1 as changed value for FV formula
        nExpectedVal = 12700;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), 'FV(A10/12,B10,C10)', 'D10');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Rate" parameter with 1 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.1230, `Case: Find "Rate" parameter with 1 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Count of payments" parameter with 15 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 1), 'FV(A11/12,B11,C11)', 'D11');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Count of payments" parameter with 15 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), 12, `Case: Find "Count of payments" parameter with 15 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Try to find "Payment" parameter with -2000 as changed value for FV formula
        nExpectedVal = 12700.16;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(11, 2), 'FV(A12/12,B12,C12)', 'D12');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Payment" parameter with -2000 as changed value for FV formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed()), -1000, `Case: Find "Payment" parameter with -2000 as changed value for FV formula. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 3, 11);
    });
    QUnit.test('LOOKUP Formula', function (assert) {
        const testData = [
            ['0', '1', '2', '3', '4', '5', '6'],
            ['1', '1', '2', '3', '4', '5', '6'],
            ['5', '1', '2', '3', '4', '5', '6'],
            ['-1', '1', '2', '3', '4', '5', '6'],
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);
        // Trying to find parameter with 0 as changed value for LOOKUP formula
        nExpectedVal = 3;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'LOOKUP(A1,B1:G1)', 'H1');
        assert.notOk(nResult, `Negative case: Trying to find parameter with 0 as changed value for LOOKUP formula. Result formula: ${nResult}`);
        assert.strictEqual(nChangingVal, 0, `Negative case: Trying to find parameter with 0 as changed value for LOOKUP formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 1 as changed value for LOOKUP formula
        nExpectedVal = 3;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 0), 'LOOKUP(A2,B2:G2)', 'H2');
        assert.notOk(nResult, `Negative case: Try to find parameter with 1 as changed value for LOOKUP formula. Result formula: ${nResult}`);
        assert.strictEqual(nChangingVal, 0.99, `Negative case: Try to find parameter with 1 as changed value for LOOKUP formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 5 as changed value for LOOKUP formula
        nExpectedVal = 3;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), 'LOOKUP(A3,B3:G3)', 'H3');
        assert.strictEqual(nResult, nExpectedVal, `Case: Try to find parameter with 5 as changed value for LOOKUP formula. Result formula: ${nResult}`);
        assert.strictEqual(nChangingVal, 3.45, `Case: Try to find parameter with 5 as changed value for LOOKUP formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with -1 as changed value for LOOKUP formula
        nExpectedVal = 3;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), 'LOOKUP(A4,B4:G4)', 'H4');
        assert.notOk(nResult, `Negative case: Try to find parameter with -1 as changed value for LOOKUP formula. Result formula: ${nResult}`);
        assert.strictEqual(nChangingVal, -1, `Negative case: Try to find parameter with -1 as changed value for LOOKUP formula. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 6, 3);
    });
    QUnit.test('Math formulas', function (assert) {
        const testData = [
            [''],
            ['-0.1'],
            ['1'],
            ['2'],
            [''],
            ['-0.1'],
            ['1'],
            ['2'],
            [''],
            ['-0.1'],
            ['1'],
            ['2']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);
        // Trying to find parameter for COS formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'COS(A1)', 'B1');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Trying to find parameter for COS formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 1.047, `Case: Trying to find parameter for COS formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with -0.1 as changed value for COS formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 0), 'COS(A2)', 'B2');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with -0.1 as changed value for COS formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), -1.047, `Case: Try to find parameter with -0.1 as changed value for COS formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 1 as changed value for COS formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), 'COS(A3)', 'B3');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 1 as changed value for COS formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 1.047, `Case: Try to find parameter with 1 as changed value for COS formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 2 as changed value for COS formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), 'COS(A4)', 'B4');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 2 as changed value for COS formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 1.047, `Case: Try to find parameter with 2 as changed value for COS formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter for SIN formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 0), 'SIN(A5)', 'B5');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Trying to find parameter for SIN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.524, `Case: Trying to find parameter for SIN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with -0.1 as changed value for SIN formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 0), 'SIN(A6)', 'B6');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with -0.1 as changed value for SIN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.524, `Case: Try to find parameter with -0.1 as changed vyingalue for SIN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 1 as changed value for SIN formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 0), 'SIN(A7)', 'B7');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 1 as changed value for SIN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.524, `Case: Try to find parameter with 1 as changed value for SIN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 2 as changed value for SIN formula
        nExpectedVal = 0.5;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 0), 'SIN(A8)', 'B8');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 2 as changed value for SIN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 2.618, `Case: Try to find parameter with 2 as changed value for SIN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter  for TAN formula
        nExpectedVal = 1;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 0), 'TAN(A9)', 'B9');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Trying to find parameter for TAN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.785, `Case: Trying to find parameter for TAN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with -0.1 as changed value for TAN formula
        nExpectedVal = 1;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), 'TAN(A10)', 'B10');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with -0.1 as changed value for TAN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.785, `Case: Try to find parameter with -0.1 as changed value for TAN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 1 as changed value for TAN formula
        nExpectedVal = 1;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 0), 'TAN(A11)', 'B11');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 1 as changed value for TAN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 0.785, `Case: Try to find parameter with 1 as changed value for TAN formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find parameter with 2 as changed value for TAN formula
        nExpectedVal = 1;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(11, 0), 'TAN(A12)', 'B12');
        assert.strictEqual(Number(nResult.toFixed(1)), nExpectedVal, `Case: Try to find parameter with 2 as changed value for TAN formula. Result formula: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(3)), 3.927, `Case: Try to find parameter with 2 as changed value for TAN formula. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 1, 11);
    });
    QUnit.test('DEVSQ formula', function (assert) {
       const testData = [
           ['','5', '8', '7', '11', '4', '3'],
           ['1','5', '8', '7', '11', '4', '3'],
           ['-1','5', '8', '7', '11', '4', '3'],
           ['9','5', '8', '7', '11', '4', '3']
       ];
       // Fill data
       let oRange = ws.getRange4(0, 0);
       oRange.fillData(testData);
       // Trying to find parameter for DEVSQ formula
       let nExpectedVal = 48;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'DEVSQ(A1:G1)', 'H1');
       assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Trying to find parameter for DEVSQ formula. Result formula: ${nResult}`);
       assert.strictEqual(Math.round(nChangingVal), 4, `Case: Trying to find parameter for DEVSQ formula. Result ChangingVal: ${nChangingVal}`);
       // Trying to find parameter with 1 as changed value for DEVSQ formula
       nExpectedVal = 48;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 0), 'DEVSQ(A2:G2)', 'H2');
       assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Try to find parameter with 1 as changed value for DEVSQ formula. Result formula: ${nResult}`);
       assert.strictEqual(Math.round(nChangingVal), 4, `Case: Try to find parameter with 1 as changed value for DEVSQ formula. Result ChangingVal: ${nChangingVal}`);
       // Trying to find parameter with -1 as changed value for DEVSQ formula
       nExpectedVal = 48;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), 'DEVSQ(A3:G3)', 'H3');
       assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Try to find parameter with -1 as changed value for DEVSQ formula. Result formula: ${nResult}`);
       assert.strictEqual(Math.round(nChangingVal), 4, `Case: Try to find parameter with -1 as changed value for DEVSQ formula. Result ChangingVal: ${nChangingVal}`);
       // Trying fo find parameter with 9 as changed value for DEVSQ formula
       nExpectedVal = 48;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 0), 'DEVSQ(A4:G4)', 'H4');
       assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Try to find parameter with 9 as changed value for DEVSQ formula. Result formula: ${nResult}`);
       assert.strictEqual(Number(nChangingVal.toFixed(6)), 8.666667, `Case: Try to find parameter with 9 as changed value for DEVSQ formula. Result ChangingVal: ${nChangingVal}`);
       // Clear data
       clearData(0, 0, 7, 3);
    });
    /*
     * TODO: BESSELI works differently than BESSELI in ms with negative numbers
     * Repro:
     * 1. Choose BESSELI formula
     * 2. Input x = -1, n = 1 (=BESSELI(-1,1))
     * Expected result: -0,565159
     * Actual result: 0,565159
     */
    QUnit.test('BESSEL formula', function(assert) {
       const testData = [
           ['', '1'],
           ['1', '1'],
           ['3', '1'],
       ];
       // Fill data
       let oRange = ws.getRange4(0, 0);
       oRange.fillData(testData);
       // Trying to find parameter for BESSELI formula
       nExpectedVal = 0.981;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'BESSELI(A1,B1)', 'D1');
       assert.strictEqual(Number(nResult.toFixed(3)), nExpectedVal, `Case: Trying to find first parameter for BESSELI formula. Result formula: ${nResult}`);
       assert.strictEqual(Number(nChangingVal.toFixed(1)), 1.5, `Case: Trying to find first parameter for BESSELI formula. Result ChangingVal: ${nChangingVal}`);
       // Trying to find parameter with 1 as changed value for BESSELI formula
       nExpectedVal = 0.981;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 0), 'BESSELI(A2,B2)', 'D2');
       assert.strictEqual(Number(nResult.toFixed(3)), nExpectedVal, `Case: Trying to find second parameter for BESSELI formula. Result formula: ${nResult}`);
       assert.strictEqual(Number(nChangingVal.toFixed(1)), 1.5, `Case: Trying to find second parameter for BESSELI formula. Result ChangingVal: ${nChangingVal}`);
       // Trying to find parameter with 3 as changed value for BESSELI formula
       nExpectedVal = 0.981;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), 'BESSELI(A3,B3)', 'D3');
       assert.strictEqual(Number(nResult.toFixed(3)), nExpectedVal, `Case: Try to find parameter with 3 as changed value for BESSELI formula. Result formula: ${nResult}`);
       assert.strictEqual(Number(nChangingVal.toFixed(1)), 1.5, `Case: Try to find parameter with 3 as changed value for BESSELI formula. Result ChangingVal: ${nChangingVal}`);
       // Clear data
       clearData(0, 0, 2, 2);
    });
    QUnit.test('Test: pause, resume, step methods', function(assert) {
        const testData = [
            ['', '180', '100000']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);
        // Method pause test on PMT formula find "Interest Rate"
        nExpectedVal = -900;
        [oParserFormula, oGoalSeek] = pauseGoalSeek(nExpectedVal, ws.getRange4(0, 0), 'PMT(A1/12,B1,C1)', 'D1');
        [nResult, nChangingVal] = getResultOutOfLoop(oParserFormula, oGoalSeek);
        assert.strictEqual(oGoalSeek.getIsPause(), true, `Case: Test pause method. Attribute bIsPause is ${oGoalSeek.getIsPause()}`);
        assert.strictEqual(oGoalSeek.getCurrentAttempt(), 1, `Case: Test pause method. Goal seek is paused. Iteration: ${oGoalSeek.getCurrentAttempt()}, Formula result: ${nResult}, changing val: ${nChangingVal}`);
        // Method resume test on PMT formula find "Interest Rate"
        oGoalSeek.resume();
        [nResult, nChangingVal] = getResultOutOfLoop(oParserFormula, oGoalSeek);
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Test resume method. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Test resume method. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 3, 0);
        // Method step test on PMT formula find "Interest Rate"
        //Fill data
        oRange.fillData(testData);
        // Pause goal seek
        [oParserFormula, oGoalSeek] = pauseGoalSeek(nExpectedVal, ws.getRange4(0, 0), 'PMT(A1/12,B1,C1)', 'D1');
        // Run goal seek step by step until goal is reached
        oGoalSeek.step();
        assert.strictEqual(oGoalSeek.getIsPause(), true, `Case: Test step method. Attribute bIsPause is ${oGoalSeek.getIsPause()}`);
        [nResult, nChangingVal] = getResultOutOfLoop(oParserFormula, oGoalSeek);
        assert.strictEqual(oGoalSeek.getCurrentAttempt(), 2, `Case: Test step method. Goal seek is paused. Iteration: ${oGoalSeek.getCurrentAttempt()}, Formula result: ${nResult}, changing val: ${nChangingVal}`);
        oGoalSeek.step();
        assert.strictEqual(oGoalSeek.getIsPause(), true, `Case: Test step method. Attribute bIsPause is ${oGoalSeek.getIsPause()}`);
        [nResult, nChangingVal] = getResultOutOfLoop(oParserFormula, oGoalSeek);
        assert.strictEqual(oGoalSeek.getCurrentAttempt(), 3, `Case: Test step method. Goal seek is paused. Iteration: ${oGoalSeek.getCurrentAttempt()}, Formula result: ${nResult}, changing val: ${nChangingVal}`);
        // Final step
        oGoalSeek.step();
        assert.strictEqual(oGoalSeek.getIsPause(), false, `Case: Test step method. Final step. Attribute bIsPause is ${oGoalSeek.getIsPause()}`);
        [nResult, nChangingVal] = getResultOutOfLoop(oParserFormula, oGoalSeek);
        assert.strictEqual(oGoalSeek.getCurrentAttempt(), 4, `Case: Test step method. Final step. Goal seek is paused. Iteration: ${oGoalSeek.getCurrentAttempt()}, Formula result: ${nResult}, changing val: ${nChangingVal}`);
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Test step method. Result PMT: ${nResult}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Test step method. Result ChangingVal: ${nChangingVal}`);
        // Clear data
        clearData(0, 0, 3, 0);
    });
    QUnit.test('Test: Financial formula. Bug #65864', function(assert) {
       const testData = [
           ['Q.1', '200000', '0.12', '=B1*C1'],
           ['Q.2', '300000', '0.13', '=B2*C2'],
           ['Q.3', '100000', '0.11', '=B3*C3'],
           ['Q.4', '0', '0.14', '=B4*C4'],
           ['Total', '', '', '=SUM(D1:D4)']
       ];
       // Fill data
       let oRange = ws.getRange4(0, 0);
       oRange.fillData(testData);
       // Trying to find parameter for Financial formula
       nExpectedVal = 100000;
       [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 1), 'SUM(D1:D4)', 'D5');
       assert.strictEqual(Number(nResult.toFixed(0)), nExpectedVal, `Case: Trying to find first parameter for Financial formula Bug #65864. Result formula: ${nResult}`);
       assert.strictEqual(Number(nChangingVal.toFixed(0)), 185714, `Case: Trying to find first parameter for Financial formula Bug #65864. Result ChangingVal: ${nChangingVal}`);
       clearData(0, 0, 3, 4);
    });
    QUnit.module('Solver');
    QUnit.test('Test: Example - Order distribution task', function(assert) {
        // Filling data
        const testData = [
            ['1000'],
            ['30'],
            ['Worker 1', '15', '2500', '0', '=C3*D3', '=D3/B3'],
            ['Worker 2', '7', '900', '0', '=C4*D4', '=D4/B4'],
            ['Worker 3', '10', '1550', '0', '=C5*D5', '=D5/B5'],
            ['Worker 4', '12', '2150', '0', '=C6*D6', '=D6/B6'],
            ['Total', '=SUM(B3:B6)', '', '=SUM(D3:D6)', '=SUM(E3:E6)', '=MAX(F3:F6)']
        ];
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);
        // Imitating of filling dialogue window of Solver tool
        const oParams = new asc_CSolverParams();
        oParams.setObjectiveFunction('Sheet1!$E$7');
        oParams.setOptimizeResultTo(c_oAscOptimizeTo.min);
        oParams.setChangingCells('Sheet1!$D$3:$D$6');
        oParams.addConstraint(0, {cellRef: 'Sheet1!$D$3:$D$6', operator: c_oAscOperator['='], constraint: 'integer'});
        oParams.addConstraint(1, {cellRef: 'Sheet1!$D$3:$D$6', operator: c_oAscOperator['>='], constraint: '0'});
        oParams.addConstraint(2, {cellRef: 'Sheet1!$D$7', operator: c_oAscOperator['='], constraint: '1000'});
        oParams.addConstraint(3, {cellRef: 'Sheet1!$F$7', operator: c_oAscOperator['<='], constraint: 'Sheet1!$A$2'});
        oParams.setVariablesNonNegative(false);
        oParams.setSolvingMethod(c_oAscSolvingMethod.grgNonlinear);
        // Filling options for calculation
        const oOptions = oParams.getOptions();
        oOptions.setConstraintPrecision('0,000001');
        oOptions.setIntOptimal('1');
        oOptions.setConvergence('0,0001');
        oOptions.setDerivatives(c_oAscDerivativeType.forward);
        oOptions.setPopulationSize('0');
        oOptions.setRandomSeed('0');
        oOptions.setMutationRate('0,075');
        oOptions.setEvoMaxTime('30');
        assert.ok(oParams, 'Params is created');
        // Creating Solver object for calculate example.
        oSolver = new CSolver(oParams, ws);
        assert.ok(oSolver, 'Solver is created');
        // Checking attributes of Solver
        let bIsParserFormulaObject = oSolver.getParsedFormula() instanceof CParserFormula;
        assert.strictEqual(bIsParserFormulaObject, true, 'Check Solver attributes: Objective cell is parserFormula object');
        assert.strictEqual(oSolver.getChangingCell().getName(), '$D$3:$D$6', 'Check Solver attributes: Changing cells is D3:D6');
        assert.strictEqual(isNaN(oSolver.getMaxIterations()), false, 'Check Solver attributes: Max iterations is not NaN');
        assert.strictEqual(oSolver.getSolvingMethod(), c_oAscSolvingMethod.grgNonlinear, 'Check Solver attributes: Solving method is grgNonlinear');
        assert.strictEqual(oSolver.getOptimizeResultTo(), c_oAscOptimizeTo.min, 'Check Solver attributes: Optimize result to is min');
        assert.strictEqual(oSolver.getVariablesNonNegative(), false, 'Check Solver attributes: Variables non negative is false');
        const aConstraints = oSolver.getConstraints();
        const aExpectedData = [
            {cellRef: '$D$3:$D$6', operator: c_oAscOperator['='], constraint: 'integer'},
            {cellRef: '$D$3:$D$6', operator: c_oAscOperator['>='],constraint: 0},
            {cellRef: '$D$7', operator: c_oAscOperator['='], constraint: 1000},
            {cellRef: '$F$7', operator: c_oAscOperator['<='], constraint: '$A$2'}
        ];
        aConstraints.forEach(function(oConstraint, index) {
            let constraint = oConstraint.getConstraint() instanceof AscCommonExcel.Range ? oConstraint.getConstraint().getName() : oConstraint.getConstraint();
            assert.strictEqual(oConstraint.getCell().getName(), aExpectedData[index].cellRef, `Check Solver attributes: Constraint element #${index} cell is ${aExpectedData[index].cellRef}`);
            assert.strictEqual(oConstraint.getOperator(), aExpectedData[index].operator, `Check Solver attributes: Constraint element #${index} operator is ${aExpectedData[index].operator}`);
            assert.strictEqual(constraint, aExpectedData[index].constraint, `Check Solver attributes: Constraint element #${index} constraint is ${aExpectedData[index].constraint}`);
        });
        // Checking calculateConstraints method
        const aConstraintsResult = oSolver.calculateConstraints();
        const aExpectedConstraintsResult = [true, true, false, true];
        aConstraintsResult.forEach(function(bResult, index) {
            assert.strictEqual(bResult, aExpectedConstraintsResult[index], `Check Solver calculateConstraints method: Constraint element #${index} result is ${aExpectedConstraintsResult[index]}`);
        })
        // Checking calculate logic
        oSolver.calculate();
        clearData(0, 0, 6, 7);
    });
    QUnit.test('Test: Example - Delivery goods in stores. Linear programming', function(assert) {
        // Filling data
        const testData = [
            // Cost of good delivery in stores
            // Store 1, Store 2, Store 3, Store 4, Store 5
            ['55', '41', '28', '11', '25'], // Warehouse 1
            ['40', '50', '8', '32', '30'],  // Warehouse 2
            ['45', '25', '60', '38', '20'], // Warehouse 3
            ['', '', '', '', '', ''], // Separator
            // Routes of delivery
            // Store 1, Store 2, Store 3, Store 4, Store 5, Total, Warehouse capacity
            ['0', '0', '0', '0', '0', '=SUM(A5:E5)', '400'], // Warehouse 1 Row: 5
            ['0', '0', '0', '0', '0', '=SUM(A6:E6)', '700'], // Warehouse 2
            ['0', '0', '0', '0', '0', '=SUM(A7:E7)', '300'], // Warehouse 3
            ['=SUM(A5:A7)', '=SUM(B5:B7)', '=SUM(C5:C7)', '=SUM(D5:D7)', '=SUM(E5:E7)'], // Total
            ['300', '230', '150', '320', '400'], // Demand
            ['=SUMPRODUCT(A1:E3,A5:E7)'] // Total cost of delivery
        ];
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);
        // Imitating of filling dialogue window of Solver tool
        const oParams = new asc_CSolverParams();
        oParams.setObjectiveFunction('Sheet1!$A$10');
        oParams.setOptimizeResultTo(c_oAscOptimizeTo.min);
        oParams.setChangingCells('Sheet1!$A$5:$E$7');
        oParams.addConstraint(0, {cellRef: 'Sheet1!$A$8:$E$8', operator: c_oAscOperator['='], constraint: 'Sheet1!$A$9:$E$9'});
        oParams.addConstraint(1, {cellRef: 'Sheet1!$F$5:$F$7', operator: c_oAscOperator['<='], constraint: 'Sheet1!$G$5:$G$7'});
        // Todo for mvp non-negative variables make as additional constraint. Later replace to checkbox setting.
        oParams.addConstraint(2, {cellRef: 'Sheet1!$A$5:$E$7', operator: c_oAscOperator['>='], constraint: '0'});
        oParams.setVariablesNonNegative(false);
        oParams.setSolvingMethod(c_oAscSolvingMethod.simplexLP);
        // Filling options for calculation. It's default option
        const oOptions = oParams.getOptions();
        oOptions.setConstraintPrecision('0,000001');
        oOptions.setIntOptimal('1');
        oOptions.setConvergence('0,0001');
        oOptions.setDerivatives(c_oAscDerivativeType.forward);
        oOptions.setPopulationSize('0');
        oOptions.setRandomSeed('0');
        oOptions.setMutationRate('0,075');
        oOptions.setEvoMaxTime('30');
        assert.ok(oParams, 'Params is created');
        // Creating Solver object for calculate example.
        oSolver = new CSolver(oParams, ws);
        assert.ok(oSolver, 'Solver is created');
        // Checking attributes of Solver
        let bIsParserFormulaObject = oSolver.getParsedFormula() instanceof CParserFormula;
        assert.strictEqual(bIsParserFormulaObject, true, 'Check Solver attributes: Objective cell is parserFormula object');
        assert.strictEqual(oSolver.getChangingCell().getName(), '$A$5:$E$7', 'Check Solver attributes: Changing cells is D3:D6');
        assert.strictEqual(isNaN(oSolver.getMaxIterations()), false, 'Check Solver attributes: Max iterations is not NaN');
        assert.strictEqual(oSolver.getSolvingMethod(), c_oAscSolvingMethod.simplexLP, 'Check Solver attributes: Solving method is grgNonlinear');
        assert.strictEqual(oSolver.getOptimizeResultTo(), c_oAscOptimizeTo.min, 'Check Solver attributes: Optimize result to is min');
        assert.strictEqual(oSolver.getVariablesNonNegative(), false, 'Check Solver attributes: Variables non negative is false');
        const aConstraints = oSolver.getConstraints();
        const aExpectedData = [
            {cellRef: '$A$8:$E$8', operator: c_oAscOperator['='], constraint: '$A$9:$E$9'},
            {cellRef: '$F$5:$F$7', operator: c_oAscOperator['<='], constraint: '$G$5:$G$7'},
            {cellRef: '$A$5:$E$7', operator: c_oAscOperator['>='],constraint: 0}
        ];
        aConstraints.forEach(function(oConstraint, index) {
            let constraint = oConstraint.getConstraint() instanceof AscCommonExcel.Range ? oConstraint.getConstraint().getName() : oConstraint.getConstraint();
            assert.strictEqual(oConstraint.getCell().getName(), aExpectedData[index].cellRef, `Check Solver attributes: Constraint element #${index} cell is ${aExpectedData[index].cellRef}`);
            assert.strictEqual(oConstraint.getOperator(), aExpectedData[index].operator, `Check Solver attributes: Constraint element #${index} operator is ${aExpectedData[index].operator}`);
            assert.strictEqual(constraint, aExpectedData[index].constraint, `Check Solver attributes: Constraint element #${index} constraint is ${aExpectedData[index].constraint}`);
        });
        // Checking prepare data for solving logic
        oSolver.prepare();
        // Checking prepare data for calculate by Simplex. Check whole CSimplexTableau class.
       let oSimplexTableau = oSolver.getSimplexTableau();
       assert.ok(oSimplexTableau, 'Check prepare data for calculate by Simplex. CSimplexTableau is created.');
       assert.strictEqual(oSimplexTableau.getBounded(), true, 'Check prepare data for calculate by Simplex. Check bBounded attribute is true');
       assert.strictEqual(oSimplexTableau.nBranchAndCutIters, 0, 'Check prepare data for calculate by Simplex. Check nBranchAndCutIters attribute is 0');
       // Checking colByVarIndex array
        let aColByVarIndex = oSimplexTableau.getColByVarIndex();
        let aColByVarIndexExpected = [
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1 ,-1 ,-1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, 1, 2,
             3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15
        ];
        assert.deepEqual(aColByVarIndex, aColByVarIndexExpected, 'Check prepare data for calculate by Simplex. Check colByVarIndex attribute');
        assert.strictEqual(oSimplexTableau.getObjectiveRowIndex(), 0, 'Check prepare data for calculate by Simplex. Check nCostRowIndex attribute is 0');
        assert.strictEqual(oSimplexTableau.getSolutionIsFound(), false, 'Check prepare data for calculate by Simplex. Check nEvaluation attribute is 0');
        assert.strictEqual(oSimplexTableau.getFeasible(), false, 'Check prepare data for calculate by Simplex. Check bFeasible attribute is true');
        assert.strictEqual(oSimplexTableau.getHeight(), 29, 'Check prepare data for calculate by Simplex. Check height attribute is 29');
        assert.strictEqual(oSimplexTableau.getLastElementIndex(), 43, 'Check prepare data for calculate by Simplex. Check lastElementIndex attribute is 43');
        assert.strictEqual(oSimplexTableau.getVarsCount(), 43, 'Check prepare data for calculate by Simplex. Check varsCount attribute is 43');
        assert.strictEqual(oSimplexTableau.getPrecision(), 1e-6, 'Check prepare data for calculate by Simplex. Check nPrecision attribute is 1e-8');
        assert.strictEqual(oSimplexTableau.getRhsColumn(), 0, 'Check prepare data for calculate by Simplex. Check nRhsColumn attribute is 0');
        assert.strictEqual(oSimplexTableau.getWidth(), 16, 'Check prepare data for calculate by Simplex. Check width attribute is 16');
        // Checking rowByVarIndex
        let aRowByVarIndex = oSimplexTableau.getRowByVarIndex();
        let aRowByVarIndexExpected = [
            1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
            11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
            21, 22, 23, 24, 25, 26, 27, 28, -1, -1,
            -1, -1, -1, -1, -1, -1 ,-1 ,-1 ,-1 ,-1,
            -1, -1, -1
        ];
        assert.deepEqual(aRowByVarIndex, aRowByVarIndexExpected, 'Check prepare data for calculate by Simplex. Check rowByVarIndex attribute');
        // Checking varIndexByCol
        let aVarIndexByCol = oSimplexTableau.getVarIndexByCol();
        let aVarIndexByColExpected = [-1, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42];
        assert.deepEqual(aVarIndexByCol, aVarIndexByColExpected, 'Check prepare data for calculate by Simplex. Check varIndexByCol attribute');
        // Checking varIndexByRow
        let aVarIndexByRow = oSimplexTableau.getVarIndexByRow();
        let aVarIndexByRowExpected = [
            -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
            11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
            21, 22, 23, 24, 25, 26, 27
        ];
        assert.deepEqual(aVarIndexByRow, aVarIndexByRowExpected, 'Check prepare data for calculate by Simplex. Check varIndexByRow attribute');
        // Checking matrix
        let aMatrix = oSimplexTableau.getMatrix();
        let aMatrixExpected = [
            [0, -55, -41, -28, -11, -25, -40, -50, -8, -32, -30, -45, -25, -60, -38, -20],
            [-300, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0],
            [300, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0],
            [-230, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0],
            [230, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0],
            [-150, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0],
            [150, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0],
            [-320, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0],
            [320, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0],
            [-400, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1],
            [400, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1],
            [400, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [700, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0],
            [300, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1],
            [0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1]
        ];
        assert.deepEqual(aMatrix, aMatrixExpected, 'Check prepare data for calculate by Simplex. Check matrix attribute');
        // Check calculate logic
        oSolver.setStartTime(Date.now());
        while (true) {
           let bCompleteCalculation = oSolver.calculate();
           if (bCompleteCalculation) {
               break;
           }
        }
        aMatrix = oSimplexTableau.getMatrix();
        aMatrixExpected = [
            [33370, -20, -11, -25, -16, -10, -40, -15, -8, -16, -5, -15, -35, -62, -32, -30],
            [300, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [230, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, -1, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0],
            [150, 0, 0, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0],
            [320, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0],
            [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [70, 0, -1, 0, 0, 1, 0, -1, 0, 0, 0, 1, 1, 1, 1, 0],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1],
            [250, -1, 0, -1, -1, -1, 0, 1, 0, 1, -1, -1, -1, -1, 0, -1],
            [0, 0, 0, 0, 1, 1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 1],
            [80, 1, 1, 1, 1, 0, 0, 0, 0, -1, 1, 0, 0, 0, -1, 0],
            [-0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [-0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [-0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            [320, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0],
            [80, 1, 1, 1, 1, 0, 0, 0, 0, -1, 1, 0, 0, 0, -1, 0],
            [300, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0],
            [-0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0],
            [150, 0, 0, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0],
            [-0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0],
            [250, -1, 0, -1, -1, -1, 0, 1, 0, 1, -1, -1, -1, -1, 0, -1],
            [-0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0],
            [230, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, -1, 0, 0, 0],
            [-0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0],
            [-0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0],
            [70, 0, -1, 0, 0, 1, 0, -1, 0, 0, 0, 1, 1, 1, 1, 0]
        ];
        assert.deepEqual(aMatrix, aMatrixExpected, 'Check calculate logic. Check updated matrix attribute');
        // Checks fill of cells  from "By Changing Variable Cells"
        const oVarIndexByCellName = oSimplexTableau.getVarIndexByCellName();
        const oExpectedChangingCells = {
            'A5': '0',
            'A6': '300',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '230',
            'C5': '0',
            'C6': '150',
            'C7': '0',
            'D5': '320',
            'D6': '0',
            'D7': '0',
            'E5': '80',
            'E6': '250',
            'E7': '70'

        };
        for (let sCellName in oVarIndexByCellName) {
            const sCellVal = ws.getCell2(sCellName).getValue();
            assert.strictEqual(sCellVal, oExpectedChangingCells[sCellName], `Check fill of cells from "By Changing Variable Cells" Cell: ${sCellName}. Value: ${sCellVal}`);
        }
        // Checks result of objective formula.
        const sObjectiveFormula = ws.getCell2('A10').getValue();
        assert.strictEqual(sObjectiveFormula, '33370', `Check result of objective formula. Cell: A10. Value: ${sObjectiveFormula}`);
        clearData(0, 0, 100, 100);
    });
    QUnit.test('Check API', function (assert) {
        // Filling data
        const testData = [
            // Cost of good delivery in stores
            // Store 1, Store 2, Store 3, Store 4, Store 5
            ['55', '41', '28', '11', '25'], // Warehouse 1
            ['40', '50', '8', '32', '30'],  // Warehouse 2
            ['45', '25', '60', '38', '20'], // Warehouse 3
            ['', '', '', '', '', ''], // Separator
            // Routes of delivery
            // Store 1, Store 2, Store 3, Store 4, Store 5, Total, Warehouse capacity
            ['0', '0', '0', '0', '0', '=SUM(A5:E5)', '400'], // Warehouse 1 Row: 5
            ['0', '0', '0', '0', '0', '=SUM(A6:E6)', '700'], // Warehouse 2
            ['0', '0', '0', '0', '0', '=SUM(A7:E7)', '300'], // Warehouse 3
            ['=SUM(A5:A7)', '=SUM(B5:B7)', '=SUM(C5:C7)', '=SUM(D5:D7)', '=SUM(E5:E7)'], // Total
            ['300', '230', '150', '320', '400'], // Demand
            ['=SUMPRODUCT(A1:E3,A5:E7)'] // Total cost of delivery
        ];
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(testData);

        // Testing api. Test saving Solver params. Test GetSolverParams and CloseSolver api. Options have default values
        let solverParams = api.asc_GetSolverParams();
        assert.ok(solverParams, 'Check API. asc_GetSolverParams is created. Open Solver params dialogue window');
        assert.strictEqual(solverParams.getObjectiveFunction(), null, 'Check API. asc_GetSolverParams. objectiveFunction is null');
        assert.strictEqual(solverParams.getOptimizeResultTo(), c_oAscOptimizeTo.max, 'Check API. asc_GetSolverParams. optimizeResultTo is max');
        assert.strictEqual(solverParams.getValueOf(), '0', 'Check API. asc_GetSolverParams. valueOf is 0');
        assert.strictEqual(solverParams.getChangingCells(), null, 'Check API. asc_GetSolverParams. changingCells is null');
        assert.strictEqual(solverParams.getConstraints().size, 0, 'Check API. asc_GetSolverParams. constraints is empty');
        assert.strictEqual(solverParams.getVariablesNonNegative(), true, 'Check API. asc_GetSolverParams. bVariablesNonNegative is false');
        assert.strictEqual(solverParams.getSolvingMethod(), c_oAscSolvingMethod.simplexLP, 'Check API. asc_GetSolverParams. solvingMethod is simplexLP');
        // Fills field from Solver params
        solverParams.setObjectiveFunction('Sheet1!$A$10');
        solverParams.setOptimizeResultTo(c_oAscOptimizeTo.min);
        solverParams.setChangingCells('Sheet1!$A$5:$E$7');
        solverParams.addConstraint(1, {cellRef: 'Sheet1!$A$8:$E$8', operator: c_oAscOperator['='], constraint: 'Sheet1!$A$9:$E$9'});
        solverParams.addConstraint(2, {cellRef: 'Sheet1!$F$5:$F$7', operator: c_oAscOperator['<='], constraint: 'Sheet1!$G$5:$G$7'});
        solverParams.addConstraint(3, {cellRef: 'Sheet1!$A$5:$E$7', operator: c_oAscOperator['>='], constraint: '0'});
        solverParams.setVariablesNonNegative(false);
        api.asc_CloseSolver(false, solverParams);
        solverParams = null;
        assert.strictEqual(solverParams, null, 'Check API. asc_CloseSolver is closed. Solver params is null');
        solverParams = api.asc_GetSolverParams();
        assert.ok(solverParams, 'Check API. asc_GetSolverParams is created with saved data. Reopen Solver params dialogue window after closing');
        assert.strictEqual(solverParams.getObjectiveFunction(), 'Sheet1!$A$10', 'Check API. asc_GetSolverParams. objectiveFunction is Sheet1!$A$10');
        assert.strictEqual(solverParams.getOptimizeResultTo(), c_oAscOptimizeTo.min, 'Check API. asc_GetSolverParams. optimizeResultTo is min');
        assert.strictEqual(solverParams.getValueOf(), '0', 'Check API. asc_GetSolverParams. valueOf is 0');
        assert.strictEqual(solverParams.getChangingCells(), 'Sheet1!$A$5:$E$7', 'Check API. asc_GetSolverParams. changingCells is Sheet1!$A$5:$E$7');
        assert.strictEqual(solverParams.getConstraints().size, 3, 'Check API. asc_GetSolverParams. constraints has 3 elements');
        assert.strictEqual(solverParams.getVariablesNonNegative(), false, 'Check API. asc_GetSolverParams. bVariablesNonNegative is false');
        assert.strictEqual(solverParams.getSolvingMethod(), c_oAscSolvingMethod.simplexLP, 'Check API. asc_GetSolverParams. solvingMethod is simplexLP');
        // Checks define names
        let aDefNames = getHiddenDefinedNamesWS(wb.dependencyFormulas);
        assert.strictEqual(aDefNames.length, 36, 'Check API. Define names is created. Count of define names is 36');
        // Checks start solve with enabled Show iterative result.
        solverParams = api.asc_GetSolverParams();
        solverParams.getOptions().setShowIterResults(true);
        api.asc_StartSolver(solverParams);

        let oSimplexTableau = wb.getSolver().getSimplexTableau();
        let oVarIndexByCellName = oSimplexTableau.getVarIndexByCellName();
        let oExpectedChangingCells = {
            'A5': '0',
            'A6': '0',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '0',
            'C5': '0',
            'C6': '0',
            'C7': '0',
            'D5': '0',
            'D6': '0',
            'D7': '0',
            'E5': '0',
            'E6': '0',
            'E7': '400'

        };
        for (let sCellName in oVarIndexByCellName) {
            const sCellVal = ws.getCell2(sCellName).getValue();
            assert.strictEqual(sCellVal, oExpectedChangingCells[sCellName], `StartSolver API. Check fill of cells from "By Changing Variable Cells" with enabled "Show iteration result" 1st iteration. Cell: ${sCellName}. Value: ${sCellVal}`);
        }
        api.asc_StepSolver();

        oExpectedChangingCells = {
            'A5': '0',
            'A6': '0',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '0',
            'C5': '0',
            'C6': '0',
            'C7': '0',
            'D5': '320',
            'D6': '0',
            'D7': '0',
            'E5': '0',
            'E6': '0',
            'E7': '400'

        };
        for (let sCellName in oVarIndexByCellName) {
            const sCellVal = ws.getCell2(sCellName).getValue();
            assert.strictEqual(sCellVal, oExpectedChangingCells[sCellName], `StepSolver API. Check fill of cells from "By Changing Variable Cells" with enabled "Show iteration result" 2nd iteration. Cell: ${sCellName}. Value: ${sCellVal}`);
        }
        // Close solver without saving result
        api.asc_CloseSolver(false, solverParams);

        oExpectedChangingCells = {
            'A5': '0',
            'A6': '0',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '0',
            'C5': '0',
            'C6': '0',
            'C7': '0',
            'D5': '0',
            'D6': '0',
            'D7': '0',
            'E5': '0',
            'E6': '0',
            'E7': '0'

        };
        for (let sCellName in oVarIndexByCellName) {
            const sCellVal = ws.getCell2(sCellName).getValue();
            assert.strictEqual(sCellVal, oExpectedChangingCells[sCellName], `CloseSolver API. isSave - false. Check fill of cells from "By Changing Variable Cells" Original values. Cell: ${sCellName}. Value: ${sCellVal}`);
        }
        solverParams = null;
        // Check start solver with final result
        solverParams = api.asc_GetSolverParams();
        solverParams.getOptions().setShowIterResults(false);
        api.asc_StartSolver(solverParams);
        let oUndoExpectedChangingCells = {
            'A5': '0',
            'A6': '0',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '0',
            'C5': '0',
            'C6': '0',
            'C7': '0',
            'D5': '0',
            'D6': '0',
            'D7': '0',
            'E5': '0',
            'E6': '0',
            'E7': '0'
        }
        oExpectedChangingCells = {
            'A5': '0',
            'A6': '300',
            'A7': '0',
            'B5': '0',
            'B6': '0',
            'B7': '230',
            'C5': '0',
            'C6': '150',
            'C7': '0',
            'D5': '320',
            'D6': '0',
            'D7': '0',
            'E5': '80',
            'E6': '250',
            'E7': '70'

        };
        // Close solver with saving result.
        api.asc_CloseSolver(true, solverParams);
        checkUndoRedo(function (_desc) {
            for (let sCellName in oVarIndexByCellName) {
                const sCellVal = ws.getCell2(sCellName).getValue();
                assert.strictEqual(sCellVal, oUndoExpectedChangingCells[sCellName], _desc + `Cell: ${sCellName}. Value: ${sCellVal}`);
            }
        }, function (_desc) {
            for (let sCellName in oVarIndexByCellName) {
                const sCellVal = ws.getCell2(sCellName).getValue();
                assert.strictEqual(sCellVal, oExpectedChangingCells[sCellName], _desc + `Cell: ${sCellName}. Value: ${sCellVal}`);
            }
        },`CloseSolver API with final result. isSave - true. Check fill of cells from "By Changing Variable Cells".`);
    });
});
