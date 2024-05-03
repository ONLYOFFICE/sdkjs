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
	 * @param {Range} oChangingCell - Changing cell.
	 * @constructor
	 */
	function CGoalSeek(oParsedFormula, nExpectedVal, oChangingCell) {
		this.oParsedFormula = oParsedFormula;
		this.nExpectedVal = nExpectedVal;
		this.oChangingCell = oChangingCell;
		this.sFormulaCellName = null;
		this.nStepDirection = null;
		this.sRegNumDecimalSeparator = AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
		this.nFirstChangingVal = null;
		this.nIntervalId = null;
		this.nCurAttempt = 0;
		this.nChangingVal = null;
		this.nPrevValue = null;
		this.nPrevFactValue = null;
		this.bIsPause = false;
		this.bReverseCompare = null;
		this.bEnabledRidder = false;
		this.nLow = null;
		this.nHigh = null;
		this.bIsSingleStep = false;

		this.nRelativeError = 1e-4; // relative error of goal seek. Default value is 1e-4
		this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
		this.nDelay = 70; // in ms for interval.
	}

	/**
	 * Fills attributes for work with goal seek.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.init = function() {
		let oChangingCell = this.getChangingCell();
		let sChangingVal = oChangingCell.getValue();

		this.setFirstChangingValue(sChangingVal);
		this.setFormulaCellName(this.getParsedFormula());
		this.initStepDirection();
		this.initReverseCompare();
		this.setChangingValue(sChangingVal ? Number(sChangingVal) : 0);
	};
	/**
	 * Main logic of calculating goal seek.
	 * Takes "Changing cell" and its change according to expected value until result of formula be equal to expected value.
	 * For calculate uses exponent step and Ridder method.
	 * Notice: That algorithm works for most formulas who can solve with help linear or non-linear equations.
	 * Exception is engineering formula like DEC2BIN, DEC2HEX etc. Those formulas can't be solved with that way because
	 * they are not linear or non-linear equations.
	 * Runs only in sync or async loop.
	 * @memberof CGoalSeek
	 * @return {boolean} The flag who recognizes end a loop of calculation goal seek. True - stop a loop, false - continue a loop.
	 */
	CGoalSeek.prototype.calculate = function() {
		if (this.getIsPause()) {
			return true;
		}

		let nChangingVal = this.getChangingValue();
		let nExpectedVal = this.getExpectedVal();
		let nPrevFactValue = this.getPrevFactValue();
		let nFactValue, nDiff, nMedianFx, nMedianVal, nLowFx;

		this.increaseCurrentAttempt();
		// Exponent step mode
		if (!this.getEnabledRidder()) {
			nFactValue = this.calculateFormula(nChangingVal);
			nDiff = nFactValue - nExpectedVal;
			// Checks should it switch to Ridder algorithm.
			if (this.getReverseCompare()) {
				this.setEnabledRidder(!!(nFactValue < nExpectedVal || (nPrevFactValue && nFactValue > nPrevFactValue)));
			} else {
				this.setEnabledRidder(!!(nFactValue > nExpectedVal || (nPrevFactValue && nFactValue < nPrevFactValue)));
			}
		}
		if (this.getEnabledRidder()) {
			if (this.getLowBorder() == null && this.getHighBorder() == null) {
				this.setLowBorder(this.getPrevValue());
				this.setHighBorder(nChangingVal);
			}
			let nLow = this.getLowBorder();
			let nHigh = this.getHighBorder();
			// Search f(lowBorder_value) and f(highBorder_value)
			nLowFx = this.calculateFormula(nLow) - nExpectedVal;
			let nHighFx = this.calculateFormula(nHigh) - nExpectedVal;
			// Search avg value in interval [nLow, nHigh]
			nMedianVal = (nLow + nHigh) / 2;
			nMedianFx = this.calculateFormula(nMedianVal) - nExpectedVal;
			// Search changing value via root of exponential function
			nChangingVal = nMedianVal + (nMedianVal - nLow) * Math.sign(nLowFx - nHighFx) * nMedianFx / Math.sqrt(Math.pow(nMedianFx,2) - nLowFx * nHighFx);
			// If result exponential function is NaN then we use nMedianVal as changing value. It may be possible for unlinear function like sin, cos, tg.
			if (isNaN(nChangingVal)) {
				nChangingVal = nMedianVal;
			}
			nFactValue = this.calculateFormula(nChangingVal);
			nDiff = nFactValue - nExpectedVal;
			this.setChangingValue(nChangingVal);
		}

		var oApi = Asc.editor;
		oApi.sendEvent("asc_onGoalSeekUpdate", nExpectedVal, nFactValue, this.getCurrentAttempt(), this.getFormulaCellName());

		// Check: Need a finish calculate
		if (Math.abs(nDiff) < this.getRelativeError()) {
			oApi.sendEvent("asc_onGoalSeekStop", true);
			return true;
		}
		if (this.getCurrentAttempt() >= this.getMaxIterations() || isNaN(nDiff)) {
			oApi.sendEvent("asc_onGoalSeekStop", false);
			return true;
		}

		//Calculates next changing value
		if (this.getEnabledRidder()) {
			if (nMedianFx < 0 !== nDiff < 0) {
				this.setLowBorder(nMedianVal);
				this.setHighBorder(nChangingVal);
			} else if (nDiff < 0 !== nLowFx < 0) {
				this.setHighBorder(nChangingVal);
			} else {
				this.setLowBorder(nChangingVal);
			}
		} else { // Exponent step logic
			let nCurAttempt = this.getCurrentAttempt();
			let nFirstChangingVal = this.getFirstChangingValue();
			let nStepDirection = this.getStepDirection();
			this.setPrevValue(nChangingVal);
			this.setPrevFactValue(nFactValue);
			if (!nFirstChangingVal) {
				this.setChangingValue((1 / 100 * nStepDirection) + (Math.pow(2, nCurAttempt - 1) - 1) * (1 / 10 * nStepDirection));
			} else {
				this.setChangingValue(nFirstChangingVal + (nFirstChangingVal / 100 * nStepDirection) + (Math.pow(2, nCurAttempt - 1) - 1) * (nFirstChangingVal / 10 * nStepDirection));
			}
		}
		if (this.getIsSingleStep()) {
			this.setIsPause(true);
			this.setIsSingleStep(false);
			return true;
		}

		return false;
	};
	/**
	 * Initialize step direction. Reverse direction (-1) or forward direction (+1).
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.initStepDirection = function() {
		let sChangingVal = this.getChangingCell().getValue();
		let nChangingVal = sChangingVal ? Number(sChangingVal) : 0;
		let nExpectedVal = this.getExpectedVal();
		let nFirstChangedVal = null;

		// Init next changed value for find nextFormulaResult
		if (nChangingVal === 0) {
			nFirstChangedVal = 0.01
		} else {
			nFirstChangedVal = nChangingVal + nChangingVal / 100;
		}
		// Find first and next formula result for check step direction
		let nFirstFormulaResult = this.calculateFormula(nChangingVal);
		let nNextFormulaResult = this.calculateFormula(nFirstChangedVal);
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
	 * Returns step direction.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getStepDirection = function() {
		return this.nStepDirection;
	};
	/**
	 * Sets step direction.
	 * @memberof CGoalSeek
	 * @param {number} nStepDirection
	 */
	CGoalSeek.prototype.setStepDirection = function(nStepDirection) {
		this.nStepDirection = nStepDirection;
	};
	/**
	 * Returns a formula object.
	 * @memberof CGoalSeek
	 * @returns {parserFormula}
	 */
	CGoalSeek.prototype.getParsedFormula = function() {
		return this.oParsedFormula;
	};
	/**
	 * Returns a result of formula with picked changing value.
	 * @memberof CGoalSeek
	 * @param {number} nChangingVal
	 * @returns {number} nFactValue
	 */
	CGoalSeek.prototype.calculateFormula = function(nChangingVal) {
		let oParsedFormula = this.getParsedFormula();
		let sRegNumDecimalSeparator = this.getRegNumDecimalSeparator();
		let oChangingCell = this.getChangingCell();
		let nFactValue = null;

		oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
		oChangingCell.worksheet.workbook.dependencyFormulas.unlockRecal();
		oParsedFormula.parse();
		nFactValue = oParsedFormula.calculate().getValue();
		// If result of formula returns type cNumber, convert to Number
		if (nFactValue instanceof oCNumberType) {
			nFactValue = nFactValue.toNumber();
		}

		return nFactValue;
	};
	/**
	 * Returns expected value.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getExpectedVal = function() {
		return this.nExpectedVal;
	};
	/**
	 * Returns changing cell.
	 * @memberof CGoalSeek
	 * @returns {Range}
	 */
	CGoalSeek.prototype.getChangingCell = function() {
		return this.oChangingCell
	};
	/**
	 * Sets changing cell.
	 * @memberof CGoalSeek
	 * @param {Range} oChangingCell
	 */
	CGoalSeek.prototype.setChangingCell = function(oChangingCell) {
		this.oChangingCell = oChangingCell;
	};
	/**
	 * Returns relative error.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getRelativeError = function() {
		return this.nRelativeError;
	};
	/**
	 * Sets relative error.
	 * @memberof CGoalSeek
	 * @param {number} nRelativeError
	 */
	CGoalSeek.prototype.setRelativeError = function(nRelativeError) {
		this.nRelativeError = nRelativeError;
	};
	/**
	 * Returns max iterations.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getMaxIterations = function() {
		return this.nMaxIterations;
	};
	/**
	 * Sets max iterations.
	 * @memberof CGoalSeek
	 * @param {number} nMaxIterations
	 */
	CGoalSeek.prototype.setMaxIterations = function(nMaxIterations) {
		this.nMaxIterations = nMaxIterations;
	};
	/**
	 * Returns formula cell name.
	 * @returns {string}
	 */
	CGoalSeek.prototype.getFormulaCellName = function() {
		return this.sFormulaCellName;
	}
	/**
	 * Sets formula cell name.
	 * @param {parserFormula} oParsedFormula
	 */
	CGoalSeek.prototype.setFormulaCellName = function(oParsedFormula) {
		let oCellWithFormula = oParsedFormula.getParent();
		let ws = oParsedFormula.getWs();

		this.sFormulaCellName = ws.getRange4(oCellWithFormula.nRow, oCellWithFormula.nCol).getName();
	}
	/**
	 * Returns a number decimal separator according chosen region. It may be "." or ",".
	 * @memberof CGoalSeek
	 * @returns {string}
	 */
	CGoalSeek.prototype.getRegNumDecimalSeparator = function() {
		return this.sRegNumDecimalSeparator;
	};
	/**
	 * Returns first changing cell value in number type.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getFirstChangingValue = function() {
		return this.nFirstChangingVal;
	}
	/**
	 * Sets first changing cell value in number type.
	 * @memberof CGoalSeek
	 * @param {string} sChangingVal
	 */
	CGoalSeek.prototype.setFirstChangingValue = function(sChangingVal) {
		this.nFirstChangingVal = sChangingVal ? Number(sChangingVal) : null;
	};
	/**
	 * Returns a delay in ms. Using for interval in UI.
	 * @returns {number}
	 */
	CGoalSeek.prototype.getDelay = function() {
		return this.nDelay;
	};
	/**
	 * Returns an id of interval. Uses for clear interval in UI.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getIntervalId = function() {
		return this.nIntervalId;
	};
	/**
	 * Sets an id of interval. Uses for clear interval in UI.
	 * @param {number} nIntervalId
	 */
	CGoalSeek.prototype.setIntervalId = function(nIntervalId) {
		this.nIntervalId = nIntervalId;
	};
	/**
	 * Returns a number of the current attempt goal seek.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getCurrentAttempt = function() {
		return this.nCurAttempt;
	};
	/**
	 * Increases a value of the current attempt goal seek.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.increaseCurrentAttempt = function() {
		this.nCurAttempt += 1;
	};
	/**
	 * Returns a value of changing cell.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getChangingValue = function() {
		return this.nChangingVal
	};
	/**
	 * Sets a value of changing cell.
	 * @memberof CGoalSeek
	 * @param {number} nChangingVal
	 */
	CGoalSeek.prototype.setChangingValue = function(nChangingVal) {
		this.nChangingVal = nChangingVal
	};
	/**
	 * Returns previous value of changing cell.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getPrevValue = function() {
		return this.nPrevValue;
	};
	/**
	 * Sets previous value of changing cell.
	 * @memberof CGoalSeek
	 * @param {number} nPrevValue
	 */
	CGoalSeek.prototype.setPrevValue = function(nPrevValue) {
		this.nPrevValue = nPrevValue;
	};
	/**
	 * Returns previous value of formula result.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getPrevFactValue = function() {
		return this.nPrevFactValue;
	};
	/**
	 * Sets previous value of formula result.
	 * @memberof CGoalSeek
	 * @param {number} nPrevFactValue
	 */
	CGoalSeek.prototype.setPrevFactValue = function(nPrevFactValue) {
		this.nPrevFactValue = nPrevFactValue;
	};
	/**
	 * Returns a flag who recognizes goal seek is paused or not.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getIsPause = function() {
		return this.bIsPause;
	};
	/**
	 * Sets a flag who recognizes goal seek is paused or not.
	 * @memberof CGoalSeek
	 * @param {boolean} bIsPause
	 */
	CGoalSeek.prototype.setIsPause = function(bIsPause) {
		this.bIsPause = bIsPause;
	};
	/**
	 * Returns a flag who recognizes should be use compare reverse (when result of calculation formula is less than expected) or not.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getReverseCompare = function() {
		return this.bReverseCompare;
	};
	/**
	 * Initializes a flag who recognizes should be use compare reverse (when result of calculation formula is less than expected) or not.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.initReverseCompare = function() {
		let nFirstFormulaResult = null;
		let nFirstChangingVal = this.getFirstChangingValue() ? Number(this.getFirstChangingValue()) : 0;

		this.bReverseCompare = false;
		nFirstFormulaResult = this.calculateFormula(nFirstChangingVal);
		if (nFirstFormulaResult > this.getExpectedVal()) {
			this.bReverseCompare = true;
		}
	};
	/**
	 * Returns a flag who recognizes goal seek is already using Ridder method or not.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getEnabledRidder = function() {
		return this.bEnabledRidder
	};
	/**
	 * Sets a flag who recognizes goal seek is already using Ridder method or not.
	 * @memberof CGoalSeek
	 * @param {boolean} bEnabledRidder
	 */
	CGoalSeek.prototype.setEnabledRidder = function(bEnabledRidder) {
		this.bEnabledRidder = bEnabledRidder
	};
	/**
	 * Returns a lower border value for Ridder algorithm.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getLowBorder = function() {
		return this.nLow;
	};
	/**
	 * Sets a lower border value for Ridder algorithm.
	 * @memberof CGoalSeek
	 * @param {number} nLow
	 */
	CGoalSeek.prototype.setLowBorder = function(nLow) {
		this.nLow = nLow;
	};
	/**
	 * Returns an upper border value for Ridder algorithm.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getHighBorder = function() {
		return this.nHigh;
	};
	/**
	 * Sets an upper border value for Ridder algorithm.
	 * @memberof CGoalSeek
	 * @param {number} nHigh
	 */
	CGoalSeek.prototype.setHighBorder = function(nHigh) {
		this.nHigh = nHigh;
	};
	/**
	 * Returns a flag who recognizes goal seek is runs in "single step" mode. Uses for "step" method.
	 * @memberof CGoalSeek
	 * @returns {boolean}
	 */
	CGoalSeek.prototype.getIsSingleStep = function() {
		return this.bIsSingleStep;
	};
	/**
	 * Sets a flag who recognizes goal seek is runs in "single step" mode.
	 * @memberof CGoalSeek
	 * @param bIsSingleStep
	 */
	CGoalSeek.prototype.setIsSingleStep = function(bIsSingleStep) {
		this.bIsSingleStep = bIsSingleStep;
	};
	/**
	 * Returns error type
	 * @memberof CGoalSeek
	 * @param {AscCommonExcel.Worksheet} ws - checked sheet.
	 * @param {Asc.Range} range - checked range.
	 * @param {Asc.c_oAscSelectionDialogType} type - dialog type.
	 * @returns {Asc.c_oAscError}
	 */
	CGoalSeek.prototype.isValidDataRef = function(ws, range, type) {
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
				let isNumberResult = true;
				//MustFormulaResultNumber
				ws && ws._getCellNoEmpty(range.r1, range.c1, function (cell) {
					if (cell && cell.isFormula()) {
						isFormula = true;
						if (cell.number == null) {
							isNumberResult = false;
						}
					}
				});
				if (!isFormula) {
					res = Asc.c_oAscError.ID.MustContainFormula;
				} else if (!isNumberResult) {
					res = Asc.c_oAscError.ID.MustFormulaResultNumber;
				}

				break;
			}
			case Asc.c_oAscSelectionDialogType.GoalSeek_ChangingCell: {
				let isValue = true;
				ws && ws._getCellNoEmpty(range.r1, range.c1, function (cell) {
					if (cell && cell.isFormula()) {
						isValue = false;
					}
				});
				if (!isValue) {
					res = Asc.c_oAscError.ID.MustContainValue;
				}
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
		let oGoalSeek = this;

		this.setIsPause(false);
		this.setIntervalId(setInterval(function() {
			let bIsFinish = oGoalSeek.calculate();
			if (bIsFinish) {
				clearInterval(oGoalSeek.getIntervalId());
			}
		}, this.getDelay()));
	};
	/**
	 * Resumes goal calculation by one step than pause it again.
	 * @memberof CGoalSeek
	 */
	CGoalSeek.prototype.step = function() {
		let oGoalSeek = this;

		this.setIsPause(false);
		this.setIsSingleStep(true);
		this.setIntervalId(setInterval(function() {
			let bIsFinish = oGoalSeek.calculate();
			if (bIsFinish) {
				clearInterval(oGoalSeek.getIntervalId());
			}
		}, this.getDelay()));
	};

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CGoalSeek = CGoalSeek;

})(window);
