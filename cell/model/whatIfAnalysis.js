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
(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
    // Import
    let cBaseOperator = AscCommonExcel.cBaseOperator;
    /**
    * @constructor
    */
    function CGoalSeek(oParsedFormula, nExpectedVal, oChangingCell) {
        this.oParsedFormula = oParsedFormula;
        this.nExpectedVal = nExpectedVal;
        this.oChangingCell = oChangingCell;
        this.nRelativeError = 0.001; // relative error of goal seek. Default value is 1e-12
        this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
    }
    CGoalSeek.prototype.calculate = function() {
        let oParsedFormula = this.getParsedFormula();
        let oChangingCell = this.getChangingCell();
        let nExpectedVal = this.getExpectedVal()
        let sChangingVal = this.getChangingCell().getValue();
        let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
        let nStep = null;
        let nCurAttempt = 0;
        let nPrevValue = null;
        oChangingCell.setValue('0');
        oParsedFormula.parse();
        let nFormulaOutStackLen = oParsedFormula.getOutStackSize();
        let bLastElementIsOperator = oParsedFormula.getOutStackElem(nFormulaOutStackLen - 1) instanceof cBaseOperator;
        if (bLastElementIsOperator && nExpectedVal < 0) {
            nStep = bLastElementIsOperator ? -0.1 : -0.01;
        }
        nStep = bLastElementIsOperator ? 0.1 : 0.01;

        while(true) {
            nCurAttempt++;
            console.log(`-DEBUG: Loop ${nCurAttempt}:  nChangingVal = ${nChangingVal}`);
            oChangingCell.setValue(String(nChangingVal));
            oParsedFormula.parse();
            let nFactValue = oParsedFormula.calculate().getValue();
            let nDiff = Math.abs(nFactValue - nExpectedVal);
            if (nDiff < this.getRelativeError() || nCurAttempt === this.getMaxIterations()) {
                console.log(`-DEBUG: Common attempts for goal seek: ${nCurAttempt}`);
                return;
            } else if (Math.abs(nFactValue) > Math.abs(nExpectedVal)) {
                //If nFactValue > nExpectedVal, then we use Regula Falsi method for calculate final changing value
                // in interval [nPrevValue, nChangingVal]
                console.log(`-DEBUG: Regula Falsi method`);
                let nLow = nPrevValue
                let nHigh = nChangingVal;
                while (true) {
                    nCurAttempt++;
                    oChangingCell.setValue(String(nLow));
                    oParsedFormula.parse();
                    let nLowFx = oParsedFormula.calculate().getValue() - nExpectedVal;
                    oChangingCell.setValue(String(nHigh));
                    oParsedFormula.parse();
                    let nHighFx = oParsedFormula.calculate().getValue() - nExpectedVal;
                    nChangingVal = nLow - nLowFx * (nHigh - nLow) / (nHighFx - nLowFx);
                    console.log(`-DEBUG: Loop ${nCurAttempt}:  nChangingVal = ${nChangingVal}`)
                    oChangingCell.setValue(String(nChangingVal));
                    oParsedFormula.parse();
                    nDiff = Math.abs(oParsedFormula.calculate().getValue() - nExpectedVal);
                    if (nDiff < this.getRelativeError() || nCurAttempt === this.getMaxIterations()) {
                        console.log(`-DEBUG: Common attempts for goal seek: ${nCurAttempt}`);
                        return;
                    } else if (nDiff < 0 === nLowFx < 0) {
                        nLow = nChangingVal;
                    } else if (nDiff < 0 === nHighFx < 0) {
                        nHigh = nChangingVal;
                    }
                }
            } else {
                nPrevValue = nChangingVal;
                nChangingVal += nStep;
                nStep *= 2;
            }
        }
    };
    CGoalSeek.prototype.getParsedFormula = function() {
        return this.oParsedFormula;
    };
    CGoalSeek.prototype.getExpectedVal = function() {
        return this.nExpectedVal;
    };
    CGoalSeek.prototype.getChangingCell = function() {
        return this.oChangingCell
    };
    CGoalSeek.prototype.setChangingCell = function(oChangingCell) {
        this.oChangingCell = oChangingCell;
    };
    CGoalSeek.prototype.getRelativeError = function() {
        return this.nRelativeError;
    };
    CGoalSeek.prototype.setRelativeError = function(nRelativeError) {
        this.nRelativeError = nRelativeError;
    };
    CGoalSeek.prototype.getMaxIterations = function() {
        return this.nMaxIterations;
    };
    CGoalSeek.prototype.setMaxIterations = function(nMaxIterations) {
        this.nMaxIterations = nMaxIterations;
    };

    // Export
    window['AscCommonExcel'] = window['AscCommonExcel'] || {};
    window['AscCommonExcel'].CGoalSeek = CGoalSeek;

})(window);
