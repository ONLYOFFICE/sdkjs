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
    let CGoalSeek = AscCommonExcel.CGoalSeek;
    let CParserFormula = AscCommonExcel.parserFormula;
    let g_oIdCounter = AscCommon.g_oIdCounter;
    let sData = AscCommon.getEmpty();
    let wb, ws, oParserFormula, oGoalSeek, nResult, nChangingVal, nExpectedVal;

    if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
        Asc.spreadsheet_api.prototype._init = function() {
        };
        let api = new Asc.spreadsheet_api({
            'id-view': 'editor_sdk'
        });

        let docInfo = new Asc.asc_CDocInfo();
        docInfo.asc_putTitle("TeSt.xlsx");
        api.DocInfo = docInfo;

        window["Asc"]["editor"] = api;

        wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api);
        AscCommon.History.init(wb);
        wb.maxDigitWidth = 7;
        wb.paddingPlusBorder = 5;

        AscCommon.g_oTableId.init();
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
        // Run goal seek
        oGoalSeek.calculate();
        // Update data for formula
        oParserFormula.parse();
        // Get results and changing value
        nResult = Number(oParserFormula.calculate().getValue());
        nChangingVal = Number(oGoalSeek.getChangingCell().getValue());

        return [nResult, nChangingVal];
    };

    QUnit.module('Goal seek');
    QUnit.test('PMT formula', function (assert) {
        const aTestData = [
            ['', '180', '100000'],
            ['0.072', '1', '100000'],
            ['0.072', '180', '']
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
        // Clear data
        clearData(0, 0, 3, 2);
    });
    QUnit.test('Custom formula. S = v * t', function (assert) {
        const aTestData = [
            ['', '5'],
            ['2000', '']
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
        // Clear data
        clearData(0, 0, 3, 1);
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
            ['-10', '2'],
            ['', '25', '5'],
            ['8', '', '5'],
            ['', ''],
            ['0', '', '123', '1', '-1233'],
            ['2'],
            ['']
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
        // Clear data
        clearData(0,0, 3, 13);
    });
    QUnit.test('Financials calculation', function (assert) {
       const aTestData = [
           ['', '10', '0.1', '2.59374246'],
           ['', '0.1' ],
           ['10', ''],
           ['', '41', '228'],
           ['1000', '', '228'],
           ['1000', '41', '1']
       ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Try to find "start investment" parameter for formula a*d
        nExpectedVal = 500000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1*D1', 'E1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "start investment" parameter for formula a*d. Result formula: ${nResult.toFixed(9)}`);
        assert.strictEqual(Number(nChangingVal.toFixed(1)), 192771.6, `Case: Find "start investment" parameter for formula a*d. Result ChangingVal: ${nChangingVal.toFixed(9)}`);
        // Try to find "term" parameter for formula a*d
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
        // Clear data
        clearData(0, 0, 5, 5);
    });
    QUnit.test('FV Formula', function (assert) {
        const aTestData = [
            ['', '12', '-1000'],
            ['0.1230', '', '-1000'],
            ['0.1230', '12', '']
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

    });
});
