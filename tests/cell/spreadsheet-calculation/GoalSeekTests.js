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
    let wb, ws, oParserFormula, oGoalSeek, nResult, nDesiredVal, nExpectedVal;

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
    const getResult = function (nExpectedVal, oDesiredVal, sFormula, sFormulaCell) {
        let nResult, nDesiredVal;
        // Init objects ParserFormula and GoalSeek
        oParserFormula = new CParserFormula(sFormula, sFormulaCell, ws);
        oGoalSeek = new CGoalSeek(oParserFormula, nExpectedVal, oDesiredVal);
        // Run goal seek
        oGoalSeek.calculate();
        // Update data for formula
        oParserFormula.parse();
        console.log('DEBUG: DesiredVal after goal seek', oDesiredVal.getValue());
        console.log('DEBUG: Calc formula', oParserFormula.calculate().getValue());
        nResult = Number(oParserFormula.calculate().getValue());
        nDesiredVal = Number(oGoalSeek.getDesiredVal().getValue());

        return [nResult, nDesiredVal];
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
        let nExpectedVal = -900;
        [nResult, nDesiredVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'PMT(A1/12,B1,C1)', 'D1');
        assert.strictEqual(Math.round(nResult), nExpectedVal, `Case: Find "Interest rate" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Number(nDesiredVal.toFixed(4)), 0.0702, `Case: Find "Interest rate" for PMT formula. Result DesiredVal: ${nDesiredVal}`);

        // Trying to find "Credit term in month" parameter
        console.log('DEBUG: Trying to find "Credit term in month" parameter');
        nExpectedVal = -910.05;
        [nResult, nDesiredVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'PMT(A2/12,B2,C2)', 'D2');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit term in month" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nDesiredVal), 180, `Case: Find "Credit term in month" for PMT formula. Result DesiredVal: ${nDesiredVal.toFixed()}`);

        // Trying to find "Credit sum" parameter
        nExpectedVal = -910.05;
        [nResult, nDesiredVal] = getResult(nExpectedVal, ws.getRange4(2, 2), 'PMT(A3/12,B3,C3)', 'D3');
        assert.strictEqual(Number(nResult.toFixed(2)), nExpectedVal, `Case: Find "Credit sum" for PMT formula. Result PMT: ${nExpectedVal}`);
        assert.strictEqual(Math.round(nDesiredVal), 100000, `Case: Find "Credit sum" for PMT formula. Result DesiredVal: ${nDesiredVal.toFixed()}`);

        clearData(0, 0, 3, 2);
    });
    QUnit.test('Custom formula', function (assert) {
        let aTestData = [
            ['', '5'],
            ['2000', '']
        ];
        // Fill data
        let oRange = ws.getRange4(0, 0);
        oRange.fillData(aTestData);
        // Trying to find "time" parameter for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nDesiredVal] = getResult(nExpectedVal, ws.getRange4(0, 0), 'A1*B1', 'D1');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "time" for custom formula. Result: ${nExpectedVal}`);
        assert.strictEqual(nDesiredVal, 2000, `Case: Find "time" for custom formula. Result DesiredVal: ${nDesiredVal}`);
        // Trying to find "speed" parameter for formula S = v*t
        nExpectedVal = 10000;
        [nResult, nDesiredVal] = getResult(nExpectedVal, ws.getRange4(1, 1), 'A2*B2', 'D2');
        assert.strictEqual(nResult, nExpectedVal, `Case: Find "speed" for custom formula. Result: ${nExpectedVal}`);
        assert.strictEqual(nDesiredVal, 5, `Case: Find "speed" for custom formula. Result DesiredVal: ${nDesiredVal}`);

        clearData(0, 0, 3, 1);
    });
});
