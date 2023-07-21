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
    /**
    * @constructor
    */
    function CGoalSeek(oParsedFormula, nExpectedVal, oDesiredVal) {
        this.oParsedFormula = oParsedFormula;
        this.nExpectedVal = nExpectedVal;
        this.oDesiredVal = oDesiredVal;
        this.nRelativeError = 1e-12; // relative error of goal seek. Default value is 1e-12
        this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
    }
    CGoalSeek.prototype.calculate = function() {
        function findInterval(nX0) {
            const INTERVAL_MAX_ATTEMPTS = 1500;
            let nCurIntervalAttempt = 0;
            // k is constant number for define step
            const k = bIsSmallScale ? 1 : 0.0001;
            let nF0, nF1, nX1, nStep;

            while(true) {
                nCurIntervalAttempt++;
                oDesiredVal.setValue(String(nX0));
                oParsedFormula.parse()
                nF0 = oParsedFormula.calculate().getValue() - nExpectedVal;
                nStep = k * Math.abs(nF0);
                nX1 = nX0 + nStep;
                oDesiredVal.setValue(String(nX1));
                oParsedFormula.parse()
                nF1 = oParsedFormula.calculate().getValue() - nExpectedVal;

                if(((nF0 < 0) !== (nF1 < 0)) || nCurIntervalAttempt >= INTERVAL_MAX_ATTEMPTS) {
                    console.log(`--DEBUG: Attempts for find interval: ${nCurIntervalAttempt}`);
                    return [nX0, nX1];
                } else {
                    nX0 = nX1;
                    nF0 = nF1;
                }
            }
        }

        let oParsedFormula = this.getParsedFormula();
        let oDesiredVal = this.getDesiredVal();
        let nExpectedVal = this.getExpectedVal()
        let sDesiredVal = this.getDesiredVal().getValue();
        let nCurAttempt = 0;
        let nX0 = null;
        let bIsSmallScale = false;
        // Desired Value is 0 for define is it small scale?
        oDesiredVal.setValue("0");
        oParsedFormula.parse();
        let nFirstFormulaValue = Number(oParsedFormula.calculate().getValue());
        if (nFirstFormulaValue === 0) {
            bIsSmallScale = true;
            nX0 = sDesiredVal && sDesiredVal !== "0" ? Number(sDesiredVal) : 1;
        } else {
            nX0 = sDesiredVal && sDesiredVal !== "0" ? Number(sDesiredVal) : 0.01;
        }
        let [nLow, nHigh] = findInterval(nX0);
        let nDesiredVal = (nLow + nHigh) / 2;

        while(true) {
            oDesiredVal.setValue(String(nDesiredVal));
            oParsedFormula.parse()
            let nFactValue = oParsedFormula.calculate().getValue();
            let nDiff = Math.abs(nFactValue - nExpectedVal);

            if(nDiff <= this.getRelativeError() || nCurAttempt >= this.getMaxIterations()) {
                console.log(`--DEBUG: Common attempts for goal seek: ${nCurAttempt}`);
                return;
            } else if(nFactValue > nExpectedVal) {
                nHigh = nDesiredVal;
            } else {
                nLow = nDesiredVal;
            }

            nDesiredVal = (nLow + nHigh) / 2;
            nCurAttempt++;
        }
    }
    CGoalSeek.prototype.getParsedFormula = function() {
        return this.oParsedFormula;
    }
    CGoalSeek.prototype.getExpectedVal = function() {
        return this.nExpectedVal;
    }
    CGoalSeek.prototype.getDesiredVal = function() {
        return this.oDesiredVal
    }
    CGoalSeek.prototype.setDesiredVal = function(oDesiredVal) {
        this.oDesiredVal = oDesiredVal;
    }
    CGoalSeek.prototype.getRelativeError = function() {
        return this.nRelativeError;
    }
    CGoalSeek.prototype.setRelativeError = function(nRelativeError) {
        this.nRelativeError = nRelativeError;
    }
    CGoalSeek.prototype.getMaxIterations = function() {
        return this.nMaxIterations;
    }
    CGoalSeek.prototype.setMaxIterations = function(nMaxIterations) {
        this.nMaxIterations = nMaxIterations;
    }

    // Export
    window['AscCommonExcel'] = window['AscCommonExcel'] || {};
    window['AscCommonExcel'].CGoalSeek = CGoalSeek;

})(window);
