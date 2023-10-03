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
	const oCNumberType = AscCommonExcel.cNumber;
	/**
	 * Class representing a goal seek feature
	 * @param {parserFormula} oParsedFormula - Formula object.
	 * For a goal seek uses methods: parse - for update values in formula, calculate - for calculate formula result.
	 * @param {number} nExpectedVal - Expected value.
	 * @param {Cell} oChangingCell - Changing cell.
	 * @constructor
	 */
	function CGoalSeek(oParsedFormula, nExpectedVal, oChangingCell) {
		this.oParsedFormula = oParsedFormula;
		this.nExpectedVal = nExpectedVal;
		this.oChangingCell = oChangingCell;
		this.nRelativeError = 1e-4; // relative error of goal seek. Default value is 1e-4
		this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
		this.nStepDirection = null;
		this.sRegNumDecimalSeparator = AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
		this.nCurAttempt = 0;
		this.nPrevValue = null;
		this.nPrevFactValue = null;
		this.nFirstFormulaResult = null;
		this.sFirstChangingCell = null;
		this.nIntervalId = null;
		this.bIsPause = false;
		this.bIsSingleStep = false;
		this.bEnabledRidder = false;
	}

	/**
	 * Main method calculating goal seek.
	 * Takes "Changing cell" and its change according to expected value until result of formula be equal to expected value.
	 * For calculate uses exponent step and Ridder method.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.calculate = function () {
		let oGoalSeek = this;
		let sRegNumDecimalSeparator = this.getRegNumDecimalSeparator();
		let oParsedFormula = this.getParsedFormula();
		let oChangingCell = this.getChangingCell();
		let nExpectedVal = this.getExpectedVal();
		let sChangingVal = oChangingCell.getValue();
		let sFirstChangingCell = this.getFirstChangingCell();
		if (sFirstChangingCell == null) {
			sFirstChangingCell = sChangingVal;
			this.setFirstChangingCell(sChangingVal);
		}
		let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
		let nCurAttempt = this.getCurrentAttempt();
		let nPrevValue = this.getPrevValue();
		let nPrevFactValue = this.getPrevFactValue();
		let bReverseCompare = false;
		if (this.getStepDirection() == null) {
			this.initStepDirection();
		}
		let nStepDirection = this.getStepDirection();
		let nFirstFormulaResult = this.getFirstFormulaResult();
		let nStartChangingVal = sFirstChangingCell ? Number(sFirstChangingCell) : 0;
		if (nFirstFormulaResult == null) {
			oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
			oParsedFormula.parse();
			nFirstFormulaResult = oParsedFormula.calculate().getValue();
			this.setFirstFormulaResult(nFirstFormulaResult);
		}
		if (nFirstFormulaResult > nExpectedVal) {
			bReverseCompare = true;
		}

		/* That algorithm works for most formulas who can solve with help linear or non-linear equations.
		 * Exception is engineering formula like DEC2BIN, DEC2HEX etc. Those formulas can't be solved with that way because
		 * they are not linear or non-linear equations.
		 */
		this.nIntervalId = setInterval(function () {
			if (!oGoalSeek.getEnabledRidder()) nCurAttempt++;
			//.sendEvent("update", nExpectedVal, nChangingVal, nCurAttempt); - update event
			let sChangingValue = String(nChangingVal).replace('.', sRegNumDecimalSeparator);
			oChangingCell.setValue(sChangingValue);
			oParsedFormula.parse();
			let nFactValue = oParsedFormula.calculate().getValue();
			if (nFactValue instanceof oCNumberType) {
				nFactValue = nFactValue.toNumber();
			}
			let nDiff = Math.abs(nFactValue - nExpectedVal);
			if (nDiff < oGoalSeek.getRelativeError()) {
				//.sendEvent("stop", true); - stop event. true/false - isSuccess;
				clearInterval(oGoalSeek.getIntervalId());
			} else if (nCurAttempt === oGoalSeek.getMaxIterations() || isNaN(nDiff)) {
				//.sendEvent("stop", false); - stop event. true/false - isSuccess;
				clearInterval(oGoalSeek.getIntervalId());
			} else if (oGoalSeek.getIsPause()) {
				oGoalSeek.setCurrentAttempt(nCurAttempt);
				oGoalSeek.setPrevValue(nPrevValue);
				oGoalSeek.setPrevFactValue(nPrevFactValue);
				clearInterval(oGoalSeek.getIntervalId());
			} else if (oGoalSeek.getEnabledRidder() || ((!bReverseCompare && (nFactValue > nExpectedVal || (nPrevFactValue && nFactValue < nPrevFactValue))) || (bReverseCompare && (nFactValue < nExpectedVal || (nPrevFactValue && nFactValue > nPrevFactValue))))) {
				//If nFactValue > nExpectedVal, then we use Ridder method for calculate final changing value
				// in interval [nPrevValue, nChangingVal]
				clearInterval(oGoalSeek.getIntervalId());
				let nLow = nPrevValue;
				let nHigh = nChangingVal;
				oGoalSeek.nIntervalId = setInterval(function () {
					nCurAttempt++;
					//.sendEvent("update", nExpectedVal, nChangingVal, nCurAttempt); - update event
					// Search f(low_value) and f(high_value)
					oChangingCell.setValue(String(nLow).replace('.', sRegNumDecimalSeparator));
					oParsedFormula.parse();
					let nLowFx = oParsedFormula.calculate().getValue() - nExpectedVal;
					oChangingCell.setValue(String(nHigh).replace('.', sRegNumDecimalSeparator));
					oParsedFormula.parse();
					let nHighFx = oParsedFormula.calculate().getValue() - nExpectedVal;
					// Search avg value in interval [nLow, nHigh]
					let nMedianVal = (nLow + nHigh) / 2;
					oChangingCell.setValue(String(nMedianVal).replace('.', sRegNumDecimalSeparator));
					oParsedFormula.parse();
					let nMedianFx = oParsedFormula.calculate().getValue() - nExpectedVal;
					// Search changing value via root of exponential function
					nChangingVal = nMedianVal + (nMedianVal - nLow) * Math.sign(nLowFx - nHighFx) * nMedianFx / Math.sqrt(Math.pow(nMedianFx,2) - nLowFx * nHighFx);
					// If result exponential function is NaN then we use nMedianVal as changing value. It may be possible for unlinear function like sin, cos, tg.
					if (isNaN(nChangingVal)) {
						nChangingVal = nMedianVal;
					}
					oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
					oParsedFormula.parse();
					nDiff = oParsedFormula.calculate().getValue() - nExpectedVal;
					if (Math.abs(nDiff) < oGoalSeek.getRelativeError()) {
						//.sendEvent("stop", true); - stop event. true/false - isSuccess;
						clearInterval(oGoalSeek.getIntervalId());
					} else if (nCurAttempt === oGoalSeek.getMaxIterations()) {
						//.sendEvent("stop", false); - stop event. true/false - isSuccess;
						clearInterval(oGoalSeek.getIntervalId());
					} else if (oGoalSeek.getIsPause()) {
						oGoalSeek.setCurrentAttempt(nCurAttempt);
						oGoalSeek.setPrevValue(nLow);
						if (!oGoalSeek.getEnabledRidder()) oGoalSeek.setEnabledRidder(true);
						clearInterval(oGoalSeek.getIntervalId());
					} else if (nMedianFx < 0 !== nDiff < 0) {
						nLow = nMedianVal;
						nHigh = nChangingVal;
					} else if (nDiff < 0 !== nLowFx < 0) {
						nHigh = nChangingVal;
					} else {
						nLow = nChangingVal;
					}
					if (oGoalSeek.getIsSingleStep()) {
						oGoalSeek.setCurrentAttempt(nCurAttempt);
						oGoalSeek.setPrevValue(nLow);
						if (!oGoalSeek.getEnabledRidder()) oGoalSeek.setEnabledRidder(true);
						oGoalSeek.setIsPause(true);
						oGoalSeek.setIsSingleStep(false);
						clearInterval(oGoalSeek.getIntervalId());
					}
				});
			} else {
				nPrevValue = nChangingVal;
				nPrevFactValue = nFactValue;
				if (nStartChangingVal === 0) {
					nChangingVal = (1 / 100 * nStepDirection) + (Math.pow(2, nCurAttempt - 1) - 1) * (1 / 10 * nStepDirection);
				} else {
					nChangingVal = nStartChangingVal + (nStartChangingVal / 100 * nStepDirection) + (Math.pow(2, nCurAttempt - 1) - 1) * (nStartChangingVal / 10 * nStepDirection);
				}
				if (oGoalSeek.getIsSingleStep()) {
					oGoalSeek.setCurrentAttempt(nCurAttempt);
					oGoalSeek.setPrevValue(nPrevValue);
					oGoalSeek.setPrevFactValue(nPrevFactValue);
					oGoalSeek.setIsPause(true);
					oGoalSeek.setIsSingleStep(false);
					clearInterval(oGoalSeek.getIntervalId());
				}
			}
		}, 50);
	};
	/**
	 * Initialize step direction. Reverse direction (-1) or forward direction (+1).
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.initStepDirection = function () {
		let oChangingCell = this.getChangingCell();
		let sChangingVal = this.getChangingCell().getValue();
		let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
		let oParsedFormula = this.getParsedFormula();
		let nExpectedVal = this.getExpectedVal();
		let sRegNumDecimalSeparator = this.getRegNumDecimalSeparator();
		let nFirstChangedVal = null;

		// Init next changed value for find nextFormulaResult
		if (nChangingVal === 0) {
			nFirstChangedVal = 0.01
		} else {
			nFirstChangedVal = nChangingVal + nChangingVal / 100;
		}
		// Find first and next formula result for check step direction
		oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
		oParsedFormula.parse();
		let nFirstFormulaResult = oParsedFormula.calculate().getValue();
		oChangingCell.setValue(String(nFirstChangedVal).replace('.', sRegNumDecimalSeparator));
		oParsedFormula.parse();
		let nNextFormulaResult = oParsedFormula.calculate().getValue();
		// If result of formula returns type cNumber, convert to Number
		if (nFirstFormulaResult instanceof oCNumberType) {
			nFirstFormulaResult = nFirstFormulaResult.toNumber();
		}
		if (nNextFormulaResult instanceof oCNumberType) {
			nNextFormulaResult = nNextFormulaResult.toNumber();
		}
		// Init step direction
		if ((nFirstFormulaResult > nExpectedVal && nNextFormulaResult > nFirstFormulaResult)) {
			this.setStepDirection(-1);
		} else if (nNextFormulaResult < nFirstFormulaResult && nNextFormulaResult < nExpectedVal) {
			this.setStepDirection(-1);
		} else if (nFirstFormulaResult === nNextFormulaResult) {
			this.setStepDirection(-1);
		} else {
			this.setStepDirection(1);
		}
	};
	/**
	 * Returns a step direction.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getStepDirection = function () {
		return this.nStepDirection;
	};
	/**
	 * Sets a step direction.
	 * @memberof CGoalSeek
	 * @param {number} nStepDirection
	 */
	CGoalSeek.prototype.setStepDirection = function (nStepDirection) {
		this.nStepDirection = nStepDirection;
	};
	/**
	 * Returns a formula object.
	 * @memberof CGoalSeek
	 * @returns {parserFormula}
	 */
	CGoalSeek.prototype.getParsedFormula = function () {
		return this.oParsedFormula;
	};
	/**
	 * Returns expected value.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getExpectedVal = function () {
		return this.nExpectedVal;
	};
	/**
	 * Returns changing cell.
	 * @memberof CGoalSeek
	 * @returns {Cell}
	 */
	CGoalSeek.prototype.getChangingCell = function () {
		return this.oChangingCell
	};
	/**
	 * Sets changing cell.
	 * @memberof CGoalSeek
	 * @param {Cell} oChangingCell
	 */
	CGoalSeek.prototype.setChangingCell = function (oChangingCell) {
		this.oChangingCell = oChangingCell;
	};
	/**
	 * Returns relative error.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getRelativeError = function () {
		return this.nRelativeError;
	};
	/**
	 * Sets relative error.
	 * @memberof CGoalSeek
	 * @param {number} nRelativeError
	 */
	CGoalSeek.prototype.setRelativeError = function (nRelativeError) {
		this.nRelativeError = nRelativeError;
	};
	/**
	 * Returns max iterations.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getMaxIterations = function () {
		return this.nMaxIterations;
	};
	/**
	 * Sets max iterations.
	 * @memberof CGoalSeek
	 * @param {number} nMaxIterations
	 */
	CGoalSeek.prototype.setMaxIterations = function (nMaxIterations) {
		this.nMaxIterations = nMaxIterations;
	};
	/**
	 * Returns number decimal separator according chosen region. It may be "." or ",".
	 * @memberof CGoalSeek
	 * @returns {string}
	 */
	CGoalSeek.prototype.getRegNumDecimalSeparator = function () {
		return this.sRegNumDecimalSeparator;
	};
	/**
	 * Returns number of the current attempt goal seek.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getCurrentAttempt = function () {
		return this.nCurAttempt;
	}
	/**
	 * Sets number of the current attempt goal seek.
	 * @memberof CGoalSeek
	 * @param {number} nCurAttempt
	 */
	CGoalSeek.prototype.setCurrentAttempt = function (nCurAttempt) {
		this.nCurAttempt = nCurAttempt;
	}
	/**
	 * Returns previous value of changing cell.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getPrevValue = function () {
		return this.nPrevValue;
	}
	/**
	 * Sets previous value of changing cell.
	 * @memberof CGoalSeek
	 * @param {number} nPrevValue
	 */
	CGoalSeek.prototype.setPrevValue = function (nPrevValue) {
		this.nPrevValue = nPrevValue;
	}
	/**
	 * Returns previous value of formula result.
	 * @returns {number}
	 */
	CGoalSeek.prototype.getPrevFactValue = function () {
		return this.nPrevFactValue;
	}
	/**
	 * Sets previous value of formula result.
	 * @param {number} nPrevFactValue
	 */
	CGoalSeek.prototype.setPrevFactValue = function (nPrevFactValue) {
		this.nPrevFactValue = nPrevFactValue;
	}
	/**
	 * Returns first formula result.
	 * @returns {number}
	 */
	CGoalSeek.prototype.getFirstFormulaResult = function () {
		return this.nFirstFormulaResult
	}
	/**
	 * Sets first formula result.
	 * @param {number} nFirstFormulaResult
	 */
	CGoalSeek.prototype.setFirstFormulaResult = function (nFirstFormulaResult) {
		this.nFirstFormulaResult = nFirstFormulaResult
	}
	/**
	 * Returns first changing cell value. Uses for cancel goal seek result.
	 * @returns {string}
	 */
	CGoalSeek.prototype.getFirstChangingCell = function () {
		return this.sFirstChangingCell;
	}
	/**
	 * Sets first changing cell value. Uses for cancel goal seek result.
	 * @param {string} oFirstChangingCell
	 */
	CGoalSeek.prototype.setFirstChangingCell = function (oFirstChangingCell) {
		this.oFirstChangingCell = oFirstChangingCell
	}
	/**
	 * Returns an interval id. Uses for clear interval.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getIntervalId = function() {
		return this.nIntervalId;
	};
	/**
	 * Returns a flag calculating is pause or not? Uses for "pause" and "continue" methods.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getIsPause = function() {
		return this.bIsPause;
	};
	/**
	 * Sets a flag calculating is pause.
	 * @memberof CGoalSeek
	 * @param {boolean} bIsPause
	 */
	CGoalSeek.prototype.setIsPause = function(bIsPause) {
		this.bIsPause = bIsPause;
	};
	/**
	 * Returns a flag is single step. Uses for "step" method.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getIsSingleStep = function() {
		return this.bIsSingleStep;
	};
	/**
	 * Sets a flag is single step.
	 * @memberof CGoalSeek
	 * @param bIsSingleStep
	 */
	CGoalSeek.prototype.setIsSingleStep = function(bIsSingleStep) {
		this.bIsSingleStep = bIsSingleStep;
	};
	/**
	 * Returns flag is already used Ridder method or not
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getEnabledRidder = function() {
		return this.bEnabledRidder
	}
	/**
	 * Sets flag is already used Ridder method or not
	 * @param {boolean} bEnabledRidder
	 */
	CGoalSeek.prototype.setEnabledRidder = function(bEnabledRidder) {
		this.bEnabledRidder = bEnabledRidder
	}
	/**
	 * Returns error type
	 * @memberof CGoalSeek
	 * @param {AscCommonExcel.Worksheet} ws - checked sheet.
	 * @param {Asc.Range} range - checked range.
	 * @param {Asc.c_oAscSelectionDialogType} type - dialog type.
	 * @returns {Asc.c_oAscError}
	 */
	CGoalSeek.prototype.isValidDataRef = function (ws, range, type) {
		let res = Asc.c_oAscError.ID.No;
		if (range === null) {
			//error text: the formula is missing a range...
			return Asc.c_oAscError.ID.DataRangeError;
		}
		if (!range.isOneCell()) {
			//error text: reference must be to a single cell...
			//TODO check def names
			res = Asc.c_oAscError.ID.MustSingleCell;
		}

		switch (type) {
			case Asc.c_oAscSelectionDialogType.GoalSeek_Cell: {
				//check formula contains
				let isFormula = false;
				ws && ws._getCellNoEmpty(range.r1, range.c1, function (cell) {
					if (cell && cell.isFormula()) {
						isFormula = true;
					}
				});
				if (!isFormula) {
					res = Asc.c_oAscError.ID.MustContainFormula;
				}

				break;
			}
			case Asc.c_oAscSelectionDialogType.GoalSeek_ChangingCell: {

				break;
			}
		}

		return res;
	};
	/**
	 * Pauses a goal seek calculation.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.pause = function() {
		this.setIsPause(true);
	};
	/**
	 * Resumes a goal seek calculation.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.resume = function() {
		this.setIsPause(false);
		this.calculate();
	};
	/**
	 * Resumes goal calculation by one step than pause it again.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.step = function() {
		this.setIsPause(false);
		this.setIsSingleStep(true);
		this.calculate();
	};

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CGoalSeek = CGoalSeek;

})(window);
