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
        console.log('DEBUG: ChangingVal after goal seek', oChangingCell.getValue());
        console.log('DEBUG: Calc formula', oParserFormula.calculate().getValue());
        nResult = Number(oParserFormula.calculate().getValue());
        nChangingVal = Number(oGoalSeek.getChangingCell().getValue());

        return [nResult, nChangingVal];
    };

    QUnit.module('Goal seek');
    QUnit.test('PMT formula', function (assert) {
        let aTestData = [
            ['', '180', '100000'],
            ['0.072', '', '100000'],
            ['0.072', '180', '']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);

        // Trying to find "Interest rate" parameter
        console.log('DEBUG: Trying to find "Interest rate" parameter');
        let nExpectedVal = -900;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'PMT(A1/12,B1,C1)', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Number(nChangingVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula. Result ChangingVal: ${nChangingVal}`);

        // Trying to find "Credit term in month" parameter
        console.log('DEBUG: Trying to find "Credit term in month" parameter');
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'PMT(A2/12,B2,C2)', 'D2');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit term in month" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 180, `Case: Find "Credit term in month" for PMT formula. Result ChangingVal: ${nChangingVal.toFixed()}`);

        // Trying to find "Credit sum" parameter
        console.log('DEBUG: Trying to find "Credit sum" parameter');
        nExpectedVal = -910.05;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 2), 'PMT(A3/12,B3,C3)', 'D3');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 100000, `Case: Find "Credit sum" for PMT formula. Result ChangingVal: ${nChangingVal.toFixed()}`);

        clearData(0, 0, 3, 2);
    });
    QUnit.test('Custom formula. S = v * t', function (assert) {
        let aTestData = [
            ['', '5'],
            ['2000', '']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find "time" parameter for formula S = v*t
        console.log(`DEBUG: Trying to find "time" parameter for formula S = v*t`);
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1*B1', 'D1');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "time" for custom formula. Result: ${nExpectedVal}`);
        assert.strictEqual(nChangingVal, 2000, `Case: Find "time" for custom formula. Result ChangingVal: ${nChangingVal}`);
        // Trying to find "speed" parameter for formula S = v*t
        console.log(`DEBUG: Trying to find "speed" parameter for formula S = v*t`);
        nExpectedVal = 10000;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'A2*B2', 'D2');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "speed" for custom formula. Result: ${nExpectedVal}`);
        assert.strictEqual(nChangingVal, 5, `Case: Find "speed" for custom formula. Result ChangingVal: ${nChangingVal}`);

        clearData(0, 0, 3, 1);
    });
    QUnit.test('Custom formula. Arithmetical operations', function (assert) {
        let aTestData = [
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
            ['8', '', '5']
        ]
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find first parameter for formula x = a + b
        console.log(`DEBUG: Trying to find first parameter for formula x = a + b`);
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1+B1', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a+b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find first parameter for a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a + b
        console.log(`DEBUG: Trying to find second parameter for formula x = a + b`);
        nExpectedVal = 10;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'A2+B2', 'D2');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a+b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 5, `Case: Find second parameter for a+b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = (a + b) * c + 3
        console.log(`DEBUG: Trying to find first parameter for formula x = (a+b)*c+3`);
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(2, 0), '(A3+B3)*C3+3', 'D3');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for (a+b)*c+3. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 1, `Case: Find first parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = (a + b) * c + 3
        console.log(`DEBUG: Trying to find second parameter for formula x = (a+b)*c+3`);
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(3, 1), '(A4+B4)*C4+3', 'D4');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for (a+b)*c+3. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find third parameter for formula x = (a + b) * c + 3
        console.log(`DEBUG: Trying to find third parameter for formula x = (a+b)*c+3`);
        nExpectedVal = 15;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(4, 2), '(A5+B5)*C5+3', 'D5');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find third parameter for (a+b)*c+3. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 4, `Case: Find third parameter for (a+b)*c+3. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = a^b
        console.log(`DEBUG: Trying to find first parameter for formula x = a^b`);
        nExpectedVal = 1024;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(5, 0), 'A6^B6', 'D6');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a^b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find first parameter for a^b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a^b
        console.log(`DEBUG: Trying to find second parameter for formula x = a^b`);
        nExpectedVal = 1024;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(6, 1), 'A7^B7', 'D7');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a^b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 10, `Case: Find second parameter for a^b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = -a*b
        console.log(`DEBUG: Trying to find first parameter for formula x = -a*b`);
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(7, 0), 'A8*B8', 'D8');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for -a*b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), -10, `Case: Find first parameter for -a*b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = -a*b
        console.log(`DEBUG: Trying to find second parameter for formula x = -a*b`);
        nExpectedVal = -20;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(8, 1), 'A9*B9', 'D9');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for -a*b. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 2, `Case: Find second parameter for -a*b. Result ChangingVal: ${nChangingVal}`);
        // Trying to find first parameter for formula x = a + b / c
        console.log(`DEBUG: Trying to find first parameter for formula x = a + b / c`);
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(9, 0), 'A10+B10/C10', 'D10');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find first parameter for a+b/c. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 8, `Case: Find first parameter for a+b/c. Result ChangingVal: ${nChangingVal}`);
        // Trying to find second parameter for formula x = a + b / c
        console.log(`DEBUG: Trying to find second parameter for formula x = a + b / c`);
        nExpectedVal = 13;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange4(10, 1), 'A11+B11/C11', 'D11');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find second parameter for a+b/c. Result formula: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nChangingVal), 25, `Case: Find second parameter for a+b/c. Result ChangingVal: ${nChangingVal}`);

        clearData(0,0, 3, 10);
    });
    QUnit.test('Math formulas', function (assert) {
        // Trying to find  parameter for formula ACOS(a)
        console.log('DEBUG: Trying to find  parameter for formula ACOS(a)');
        nExpectedVal = 1.048;
        [nResult, nChangingVal] = getResult(nExpectedVal, ws.getRange2("A1"), 'ACOS(A1)', 'D1');
        assert.strictEqual(Number(nResult.toFixed(3)), nExpectedVal, `Case: Find  parameter for formula ACOS(a). Result formula: ${nExpectedVal}`);
        assert.strictEqual(nChangingVal, 0.5, `Case: Find  parameter for formula ACOS(a). Result ChangingVal: ${nChangingVal}`);
    });
});
