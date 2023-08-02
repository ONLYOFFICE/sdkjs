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
    const cBaseOperator = AscCommonExcel.cBaseOperator;
    const sRegNumDecimalSeparator = AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
    /**
    * @constructor
    */
    function CGoalSeek(oParsedFormula, nExpectedVal, oChangingCell) {
        this.oParsedFormula = oParsedFormula;
        this.nExpectedVal = nExpectedVal;
        this.oChangingCell = oChangingCell;
        this.nRelativeError = 1e-6; // relative error of goal seek. Default value is 1e-6
        this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
        this.nStep = null;
    }
    CGoalSeek.prototype.calculate = function() {
        let oParsedFormula = this.getParsedFormula();
        let oChangingCell = this.getChangingCell();
        let nExpectedVal = this.getExpectedVal();
        let sChangingVal = this.getChangingCell().getValue();
        let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
        let nCurAttempt = 0;
        let nPrevValue = null;
        let bReverseCompare = false;
        this.initStep();
        let nStep = this.getStep();
        oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
        oParsedFormula.parse();
        let nFirstFormulaResult = oParsedFormula.calculate().getValue();
        if (nFirstFormulaResult > nExpectedVal) {
            bReverseCompare = true;
        }

        while(true) {
            nCurAttempt++;
            let sChangingValue = String(nChangingVal).replace('.', sRegNumDecimalSeparator);
            oChangingCell.setValue(sChangingValue);
            oParsedFormula.parse();
            let nFactValue = oParsedFormula.calculate().getValue();
            let nDiff = Math.abs(nFactValue - nExpectedVal);
            if (nDiff < this.getRelativeError() || nCurAttempt === this.getMaxIterations()) {
                return;
            } else if ((!bReverseCompare && nFactValue > nExpectedVal) || (bReverseCompare && nFactValue < nExpectedVal)) {
                //If nFactValue > nExpectedVal, then we use Regula Falsi method for calculate final changing value
                // in interval [nPrevValue, nChangingVal]
                let nLow = nPrevValue
                let nHigh = nChangingVal;
                while (true) {
                    nCurAttempt++;
                    oChangingCell.setValue(String(nLow).replace('.', sRegNumDecimalSeparator));
                    oParsedFormula.parse();
                    let nLowFx = oParsedFormula.calculate().getValue() - nExpectedVal;
                    oChangingCell.setValue(String(nHigh).replace('.', sRegNumDecimalSeparator));
                    oParsedFormula.parse();
                    let nHighFx = oParsedFormula.calculate().getValue() - nExpectedVal;
                    nChangingVal = nLow - nLowFx * (nHigh - nLow) / (nHighFx - nLowFx);
                    oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
                    oParsedFormula.parse();
                    nDiff = oParsedFormula.calculate().getValue() - nExpectedVal;
                    if (Math.abs(nDiff) < this.getRelativeError() || nCurAttempt === this.getMaxIterations()) {
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
    CGoalSeek.prototype.initStep = function () {
        let oChangingCell = this.getChangingCell();
        let sChangingVal = this.getChangingCell().getValue();
        let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
        let oParsedFormula = this.getParsedFormula();
        let nExpectedVal = this.getExpectedVal();

        // Find first formula result for init step
        oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
        oParsedFormula.parse();
        let nFirstFormulaResult = oParsedFormula.calculate().getValue();
        // Checking: Formula has arithmetic operator
        let nFormulaOutStackLen = oParsedFormula.getOutStackSize();
        let oLastElement = oParsedFormula.getOutStackElem(nFormulaOutStackLen - 1);
        let bLastElementIsOperator = oLastElement instanceof cBaseOperator;
        if (bLastElementIsOperator && !oLastElement.toString().includes('/') && (nExpectedVal < 0 || nFirstFormulaResult > nExpectedVal)) {
            this.setStep(bLastElementIsOperator && !oLastElement.toString().includes('^') ? -0.1 : -0.01);
        } else if (nExpectedVal > 0 && nFirstFormulaResult === 0) { // Checking next value. If next value is negative number than step must be gone reverse
            oChangingCell.setValue(String('0.1').replace('.', sRegNumDecimalSeparator));
            oParsedFormula.parse();
            let nFormulaResult = oParsedFormula.calculate().getValue();
            if (nFormulaResult < 0) {
                this.setStep(bLastElementIsOperator && !oLastElement.toString().includes('^') ? -0.1 : -0.01);
            } else {
                this.setStep(bLastElementIsOperator && !oLastElement.toString().includes('^') ? 0.1 : 0.01);
            }
        } else {
            this.setStep(bLastElementIsOperator && !oLastElement.toString().includes('^') ? 0.1 : 0.01);
        }
    }
    CGoalSeek.prototype.getStep = function () {
        return this.nStep;
    }
    CGoalSeek.prototype.setStep = function(nStep) {
        this.nStep = nStep;
    }
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
