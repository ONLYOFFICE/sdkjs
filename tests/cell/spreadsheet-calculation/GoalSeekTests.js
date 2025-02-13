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
    let CGoalSeek = AscCommonExcel.CGoalSeek;
    let CParserFormula = AscCommonExcel.parserFormula;
    let g_oIdCounter = AscCommon.g_oIdCounter;
    let sData = AscCommon.getEmpty();
    let wb, ws, oParserFormula, oGoalSeek, nResult, nChangingVal, nExpectedVal, api;

    if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
        Asc.spreadsheet_api.prototype._init = function() {
        };
        api = new Asc.spreadsheet_api({
            'id-view': 'editor_sdk'
        });

        let docInfo = new Asc.asc_CDocInfo();
        docInfo.asc_putTitle("TeSt.xlsx");
        api.DocInfo = docInfo;
        api.initCollaborativeEditing({});
        window["Asc"]["editor"] = api;

        AscCommon.g_oTableId.init();
        wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);
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
    wb.dependencyFormulas.lockRecal();
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
});
