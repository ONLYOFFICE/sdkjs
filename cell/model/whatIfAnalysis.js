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
	const CellValueType = AscCommon.CellValueType;

	// Collections for UI Solver feature
	/** @enum {number} */
	const c_oAscOptimizeTo = {
		max: 1,
		min: 2,
		valueOf: 3
	};
	/** @enum {number} */
	const c_oAscSolvingMethod = {
		grgNonlinear: 1,
		simplexLP: 2,
		evolutionary: 3
	};
	/** @enum {number} */
	const c_oAscDerivativeType = {
		forward: 1,
		central: 2
	};
	/** @enum {number} */
	const c_oAscOperator = {
		'<=': 1,
		'=': 2,
		'>=': 3,
		integer: 4,
		bin: 5,
		diff: 6
	};
	/**@enum {number} */
	const c_oAscSolverResult = {
		keepSolverSolution: 0,
		restoreOriginalValues: 1
	};
	/**@enum {number} */
	const c_oResultStatus = {
		foundOptimalSolution: 0,
		solutionHasConverged: 1,
		cannotImproveSolution: 2,
		maxIterationsReached: 3,
		objectiveCellNotConverge: 4,
		notFindFeasibleSolution: 5,
		stoppedByUser: 6,
		linearityConditionsNotSatisfied: 7,
		tooLargeProblem: 8,
		errorValInObjectiveOrConstraintCell: 9,
		maxTimeReached: 10,
		notEnoughMemory: 11,
		errorInModel: 12,
		foundIntegerSolution: 13,
		maxFeasibleSolutionReached: 14,
		maxSubproblemSolutionReached: 15,
	};

	// Common class
	/**
	 * Class representing base attributes and methods for features of analysis.
	 * @param {parserFormula} oParsedFormula - Formula object
	 * @param {Range} oChangingCell - Changing cells.
	 * For Goal Seek feature it's 1-1 object, for Solver 1-*.
	 * @constructor
	 */
	function CBaseAnalysis(oParsedFormula, oChangingCell) {
		this.oParsedFormula = oParsedFormula;
		this.oChangingCell = oChangingCell;
		this.sRegNumDecimalSeparator = AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
		this.nIntervalId = null;
		this.bIsPause = false;
		this.nDelay = 70; // in ms for interval.
		this.nCurAttempt = 0;
		this.bIsSingleStep = false;
		this.nPrevFactValue = null;
		// todo add value from settings
		this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
	}

	/**
	 * Returns a result of formula with picked changing value.
	 * @memberof CBaseAnalysis
	 * @param {number} [nChangingVal]
	 * @param {Range} [oChangingCell]
	 * @returns {number} nFactValue
	 */
	CBaseAnalysis.prototype.calculateFormula = function(nChangingVal, oChangingCell) {
		let oParsedFormula = this.getParsedFormula();
		let nFactValue = null;

		if (nChangingVal !== undefined && oChangingCell) {
			let sRegNumDecimalSeparator = this.getRegNumDecimalSeparator();
			oChangingCell.setValue(String(nChangingVal).replace('.', sRegNumDecimalSeparator));
			oChangingCell.worksheet.workbook.dependencyFormulas.unlockRecal();
		}
		oParsedFormula.parse();
		nFactValue = oParsedFormula.calculate().getValue();
		// If result of formula returns type cNumber, convert to Number
		if (nFactValue instanceof oCNumberType) {
			nFactValue = nFactValue.toNumber();
		}

		return nFactValue;
	};
	/**
	 * Returns a formula object.
	 * @memberof CBaseAnalysis
	 * @returns {parserFormula}
	 */
	CBaseAnalysis.prototype.getParsedFormula = function() {
		return this.oParsedFormula;
	};
	/**
	 * Returns changing cell.
	 * @memberof CBaseAnalysis
	 * @returns {Range}
	 */
	CBaseAnalysis.prototype.getChangingCell = function() {
		return this.oChangingCell;
	};
	/**
	 * Sets changing cell.
	 * @memberof CBaseAnalysis
	 * @param {Range} oChangingCell
	 */
	CBaseAnalysis.prototype.setChangingCell = function(oChangingCell) {
		this.oChangingCell = oChangingCell;
	};
	/**
	 * Returns a number decimal separator according chosen region. It may be "." or ",".
	 * @memberof CBaseAnalysis
	 * @returns {string}
	 */
	CBaseAnalysis.prototype.getRegNumDecimalSeparator = function() {
		return this.sRegNumDecimalSeparator;
	};
	/**
	 * Returns an id of interval. Uses for clear interval in UI.
	 * @memberof CBaseAnalysis
	 * @returns {number}
	 */
	CBaseAnalysis.prototype.getIntervalId = function() {
		return this.nIntervalId;
	};
	/**
	 * Sets an id of interval. Uses for clear interval in UI.
	 * @memberof CBaseAnalysis
	 * @param {number} nIntervalId
	 */
	CBaseAnalysis.prototype.setIntervalId = function(nIntervalId) {
		this.nIntervalId = nIntervalId;
	};
	/**
	 * Returns a flag who recognizes calculation process is paused or not.
	 * @memberof CBaseAnalysis
	 * @returns {boolean}
	 */
	CBaseAnalysis.prototype.getIsPause = function() {
		return this.bIsPause;
	};
	/**
	 * Sets a flag who recognizes calculation process is paused or not.
	 * @memberof CBaseAnalysis
	 * @param bIsPause
	 */
	CBaseAnalysis.prototype.setIsPause = function(bIsPause) {
		this.bIsPause = bIsPause;
	};
	/**
	 * Returns a delay in ms. Using for interval in UI.
	 * @memberof CBaseAnalysis
	 * @returns {number}
	 */
	CBaseAnalysis.prototype.getDelay = function() {
		return this.nDelay;
	};
	/**
	 * Returns a number of the current attempt.
	 * @memberof CBaseAnalysis
	 * @returns {number}
	 */
	CBaseAnalysis.prototype.getCurrentAttempt = function() {
		return this.nCurAttempt;
	};
	/**
	 * Increases a value of the current attempt.
	 * @memberof CBaseAnalysis
	 */
	CBaseAnalysis.prototype.increaseCurrentAttempt = function() {
		this.nCurAttempt += 1;
	};
	/**
	 * Returns a flag who recognizes goal seek is runs in "single step" mode. Uses for "step" method.
	 * @memberof CBaseAnalysis
	 * @returns {boolean}
	 */
	CBaseAnalysis.prototype.getIsSingleStep = function() {
		return this.bIsSingleStep;
	};
	/**
	 * Sets a flag who recognizes goal seek is runs in "single step" mode.
	 * @memberof CBaseAnalysis
	 * @param {boolean} bIsSingleStep
	 */
	CBaseAnalysis.prototype.setIsSingleStep = function(bIsSingleStep) {
		this.bIsSingleStep = bIsSingleStep;
	};
	/**
	 * Returns previous value of formula result.
	 * @memberof CBaseAnalysis
	 * @returns {number}
	 */
	CBaseAnalysis.prototype.getPrevFactValue = function() {
		return this.nPrevFactValue;
	};
	/**
	 * Sets previous value of formula result.
	 * @memberof CBaseAnalysis
	 * @param {number} nPrevFactValue
	 */
	CBaseAnalysis.prototype.setPrevFactValue = function(nPrevFactValue) {
		this.nPrevFactValue = nPrevFactValue;
	};
	/**
	 * Returns max iterations.
	 * @memberof CBaseAnalysis
	 * @returns {number}
	 */
	CBaseAnalysis.prototype.getMaxIterations = function() {
		return this.nMaxIterations;
	};
	/**
	 * Sets max iterations.
	 * @memberof CBaseAnalysis
	 * @param {number} nMaxIterations
	 */
	CBaseAnalysis.prototype.setMaxIterations = function(nMaxIterations) {
		this.nMaxIterations = nMaxIterations;
	};

	// Goal seek
	/**
	 * Class representing a goal seek feature
	 * @param {parserFormula} oParsedFormula - Formula object.
	 * For a goal seek uses methods: parse - for update values in formula, calculate - for calculate formula result.
	 * @param {number} nExpectedVal - Expected value.
	 * @param {Range} oChangingCell - Changing cell.
	 * @constructor
	 */
	function CGoalSeek(oParsedFormula, nExpectedVal, oChangingCell) {
		CBaseAnalysis.call(this, oParsedFormula, oChangingCell);

		this.nExpectedVal = nExpectedVal;
		this.sFormulaCellName = null;
		this.nStepDirection = null;
		this.nFirstChangingVal = null;
		this.nChangingVal = null;
		this.nPrevValue = null;
		this.bReverseCompare = null;
		this.bEnabledRidder = false;
		this.nLow = null;
		this.nHigh = null;

		this.nRelativeError = 1e-4; // relative error of goal seek. Default value is 1e-4
	}

	CGoalSeek.prototype = Object.create(CBaseAnalysis.prototype);
	CGoalSeek.prototype.constructor = CGoalSeek;
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

		const oChangingCell = this.getChangingCell();
		let nChangingVal = this.getChangingValue();
		let nExpectedVal = this.getExpectedVal();
		let nPrevFactValue = this.getPrevFactValue();
		let nFactValue, nDiff, nMedianFx, nMedianVal, nLowFx;

		this.increaseCurrentAttempt();
		// Exponent step mode
		if (!this.getEnabledRidder()) {
			nFactValue = this.calculateFormula(nChangingVal, oChangingCell);
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
			nLowFx = this.calculateFormula(nLow, oChangingCell) - nExpectedVal;
			let nHighFx = this.calculateFormula(nHigh, oChangingCell) - nExpectedVal;
			// Search avg value in interval [nLow, nHigh]
			nMedianVal = (nLow + nHigh) / 2;
			nMedianFx = this.calculateFormula(nMedianVal, oChangingCell) - nExpectedVal;
			// Search changing value via root of exponential function
			nChangingVal = nMedianVal + (nMedianVal - nLow) * Math.sign(nLowFx - nHighFx) * nMedianFx / Math.sqrt(Math.pow(nMedianFx,2) - nLowFx * nHighFx);
			// If result exponential function is NaN then we use nMedianVal as changing value. It may be possible for unlinear function like sin, cos, tg.
			if (isNaN(nChangingVal)) {
				nChangingVal = nMedianVal;
			}
			nFactValue = this.calculateFormula(nChangingVal, oChangingCell);
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
		const oChangingCell = this.getChangingCell();
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
		let nFirstFormulaResult = this.calculateFormula(nChangingVal, oChangingCell);
		let nNextFormulaResult = this.calculateFormula(nFirstChangedVal, oChangingCell);
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
	 * Returns expected value.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getExpectedVal = function() {
		return this.nExpectedVal;
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
	 * Returns formula cell name.
	 * @returns {string}
	 */
	CGoalSeek.prototype.getFormulaCellName = function() {
		return this.sFormulaCellName;
	};
	/**
	 * Sets formula cell name.
	 * @param {parserFormula} oParsedFormula
	 */
	CGoalSeek.prototype.setFormulaCellName = function(oParsedFormula) {
		let oCellWithFormula = oParsedFormula.getParent();
		let ws = oParsedFormula.getWs();

		this.sFormulaCellName = ws.getRange4(oCellWithFormula.nRow, oCellWithFormula.nCol).getName();
	};
	/**
	 * Returns first changing cell value in number type.
	 * @memberof CGoalSeek
	 * @returns {number}
	 */
	CGoalSeek.prototype.getFirstChangingValue = function() {
		return this.nFirstChangingVal;
	};
	/**
	 * Sets first changing cell value in number type.
	 * @memberof CGoalSeek
	 * @param {string} sChangingVal
	 */
	CGoalSeek.prototype.setFirstChangingValue = function(sChangingVal) {
		this.nFirstChangingVal = sChangingVal ? Number(sChangingVal) : null;
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
		const oChangingCell = this.getChangingCell();
		let nFirstFormulaResult = null;
		let nFirstChangingVal = this.getFirstChangingValue() ? Number(this.getFirstChangingValue()) : 0;

		this.bReverseCompare = false;
		nFirstFormulaResult = this.calculateFormula(nFirstChangingVal, oChangingCell);
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

	// Solver
	// Classes for interact with UI.
	/**
	 * Class representing object with input data from dialogue window of Solver tool.
	 * @constructor
	 * @returns {asc_CSolverParams}
	 */
	function asc_CSolverParams () {
		this.sObjectiveFunction = null;
		this.sChangingCells = null;
		this.nOptimizeResultTo = c_oAscOptimizeTo.max;
		this.sValueOf = '0';
		this.aConstraints = new Map();
		this.bVariablesNonNegative = true;
		this.nSolvingMethod = c_oAscSolvingMethod.simplexLP;

		this.oOptions = new asc_COptions();

		this.oDefNameToAttribute = {
			'solver_opt': 'sObjectiveFunction',
			'solver_typ': 'nOptimizeResultTo',
			'solver_val': 'sValueOf',
			'solver_adj': 'sChangingCells',
			'solver_neg': 'bVariablesNonNegative',
			'solver_eng': 'nSolvingMethod'
		};

		return this;
	}

	/**
	 * Returns value of "Set Objective" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {string}
	 */
	asc_CSolverParams.prototype.getObjectiveFunction = function () {
		return this.sObjectiveFunction;
	};
	/**
	 * Sets value of "Set Objective" parameter.
	 * @memberof asc_CSolverParams
	 * @param {string|null} objectiveFunction
	 */
	asc_CSolverParams.prototype.setObjectiveFunction = function (objectiveFunction) {
		this.sObjectiveFunction = objectiveFunction;
	};
	/**
	 * Returns value of "By Changing Variable Cells" parameters.
	 * @memberof asc_CSolverParams
	 * @returns {string}
	 */
	asc_CSolverParams.prototype.getChangingCells = function () {
		return this.sChangingCells;
	};
	/**
	 * Sets value of "By Changing Variable Cells" parameters.
	 * @param {string|null} changingCells
	 */
	asc_CSolverParams.prototype.setChangingCells = function (changingCells) {
		this.sChangingCells = changingCells;
	};
	/**
	 * Returns value of "To" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {c_oAscOptimizeTo}
	 */
	asc_CSolverParams.prototype.getOptimizeResultTo = function () {
		return this.nOptimizeResultTo;
	};

	/**
	 * Sets value of "To" parameter.
	 * @memberof asc_CSolverParams
	 * @param {c_oAscOptimizeTo} optimizeResultTo
	 */
	asc_CSolverParams.prototype.setOptimizeResultTo = function (optimizeResultTo) {
		this.nOptimizeResultTo = optimizeResultTo;
	};
	/**
	 * Returns value of "Value of" input field.
	 * @memberof asc_CSolverParams
	 * @param {string} sValueOf
	 */
	asc_CSolverParams.prototype.setValueOf = function (sValueOf) {
		this.sValueOf = sValueOf;
	};
	/**
	 * Sets value of "Value of" input field.
	 * @memberof {asc_CSolverParams}
	 * @returns {string}
	 */
	asc_CSolverParams.prototype.getValueOf = function () {
		return this.sValueOf;
	};
	/**
	 * Returns value of "Subject to the Constraints" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {Map}
	 */
	asc_CSolverParams.prototype.getConstraints = function () {
		return this.aConstraints;
	};
	/**
	 * Adds the constraint in "Subject to the Constraints" parameter.
	 * @memberof asc_CSolverParams
	 * @param {number} index - key of constraint in Map object. Starts with 1 ... N.
	 * @param {{cellRef:string, operator:c_oAscOperator, constraint:string}} constraint
	 */
	asc_CSolverParams.prototype.addConstraint = function (index, constraint) {
		this.aConstraints.set(index, constraint);
	};
	/**
	 * Edits the chosen constraint in "Subject to the Constraints" parameter.
	 * @memberof asc_CSolverParams
	 * @param {number} index - index of chosen constraint
	 * @param {{cellRef:string, operator:c_oAscOperator, constraint:string}} constraint
	 */
	asc_CSolverParams.prototype.editConstraint = function (index, constraint) {
		this.aConstraints.set(index, constraint);
	};
	/**
	 * Removes the chosen constraint in "Subject to the Constraints" parameter.
	 * @memberof asc_CSolverParams
	 * @param {number} index
	 */
	asc_CSolverParams.prototype.removeConstraint = function (index) {
		this.aConstraints.delete(index);
	};
	/**
	 * Returns value of "Make Unconstrained Variables Non-negative" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {boolean}
	 */
	asc_CSolverParams.prototype.getVariablesNonNegative = function () {
		return this.bVariablesNonNegative;
	};
	/**
	 * Sets value of "Make Unconstrained Variables Non-negative" parameter.
	 * @memberof asc_CSolverParams
	 * @param {boolean} variablesNonNegative
	 */
	asc_CSolverParams.prototype.setVariablesNonNegative = function (variablesNonNegative) {
		this.bVariablesNonNegative = variablesNonNegative;
	};
	/**
	 * Returns value of "Select a Solving Method" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {c_oAscSolvingMethod}
	 */
	asc_CSolverParams.prototype.getSolvingMethod = function () {
		return this.nSolvingMethod;
	};
	/**
	 * Sets value of "Select a Solving Method" parameter.
	 * @memberof asc_CSolverParams
	 * @param {c_oAscSolvingMethod} solvingMethod
	 */
	asc_CSolverParams.prototype.setSolvingMethod = function (solvingMethod) {
		this.nSolvingMethod = solvingMethod;
	};
	/**
	 * Returns the options.
	 * @memberof asc_CSolverParams
	 * @returns {asc_COptions}
	 */
	asc_CSolverParams.prototype.getOptions = function () {
		return this.oOptions;
	};
	/**
	 * Resets all params and options to default value.
	 * @memberof asc_CSolverParams
	 */
	asc_CSolverParams.prototype.resetAll = function () {
		this.setObjectiveFunction(null);
		this.setChangingCells(null);
		this.setOptimizeResultTo(c_oAscOptimizeTo.max);
		this.setValueOf('0');
		this.getConstraints().clear();
		this.setVariablesNonNegative(true);
		this.getOptions().resetAll();

	};
	/**
	 * Returns object which connecting def name to attribute of class.
	 * @memberof  asc_CSolverParams
	 * @returns {{
	 * solver_opt: string, solver_typ: string, solver_val: string,
	 * solver_adj: string, solver_neg: string, solver_eng: string
	 * }}
	 */
	asc_CSolverParams.prototype.getDefNameToAttribute = function() {
		return this.oDefNameToAttribute;
	};

	/**
	 * Returns array of hidden defined names.
	 * @param {DependencyGraph} oDependencyFormulas
	 * @returns {asc_CDefName[]}
	 */
	function getHiddenDefinedNamesWS (oDependencyFormulas) {
		const oActiveWS = oDependencyFormulas.wb.getActiveWs();
		const aHiddenDefNames = [];

		oDependencyFormulas._foreachDefNameSheet(oActiveWS.getId(), function(oDefName) {
			if (oDefName.hidden) {
				aHiddenDefNames.push(oDefName.getAscCDefName(true));
			}
		});

		return aHiddenDefNames;
	}

	/**
	 * Adds Def names ranges according to the values of the Solver params and options.
	 * @memberof asc_CSolverParams
	 * @param {Workbook} oWbModel
	 */
	asc_CSolverParams.prototype.createDefNames = function(oWbModel) {
		if (!oWbModel) {
			return;
		}
		function fillDefName(sDefName, ref) {
			let oOldDefName = null;
			if (mSolverDefNames.size) {
				oOldDefName = mSolverDefNames.get(sDefName);
			}
			if (ref === null) {
				if (oOldDefName) {
					const sSheetId = this.getSheetIdByIndex(oOldDefName.LocalSheetId);
					oWbModel.dependencyFormulas._removeDefName(sSheetId, oOldDefName.Name, null);
				}
				return;
			}
			if (typeof ref === 'boolean') {
				ref = ref ? 1 : 2;
			}
			let oNewDefName = new Asc.asc_CDefName(sDefName, ref, nWsId, undefined, true, undefined, undefined, true);
			oWbModel.editDefinesNames(oOldDefName, oNewDefName);
		}

		const oDependencyFormulas = oWbModel.dependencyFormulas;
		const aWsDefNames = getHiddenDefinedNamesWS(oDependencyFormulas);
		const oDefNameToAttribute = this.getDefNameToAttribute();
		const oOptions = this.getOptions();
		const oOptionsDefNameToAttribute = oOptions.getDefNameToAttribute();
		const mSolverDefNames = new Map();
		const nWsId = oWbModel.getActive();

		// Finds existed defined names for the solver.
		if (aWsDefNames.length) {
			aWsDefNames.forEach(function(oDefName) {
				if (~oDefName.Name.indexOf('solver_')) {
					mSolverDefNames.set(oDefName.Name, oDefName);
				}
			});
		}
		// Fills constant define names
		fillDefName('solver_ver', 3);
		fillDefName('solver_est', 1);
		fillDefName('solver_nwt', 1)
		// Fills Solver parameters
		for (let sDefName in oDefNameToAttribute) {
			fillDefName(sDefName, this[oDefNameToAttribute[sDefName]]);
		}
		// Fills constraints
		const mConstraints = this.getConstraints();
		let nIndex = 1;
		if (!mConstraints.size) {
			return;
		}
		fillDefName('solver_num', mConstraints.size);
		mConstraints.forEach(function(constraint) {
			fillDefName('solver_lhs' + nIndex, constraint.cellRef);
			fillDefName('solver_rhs' + nIndex, constraint.constraint);
			fillDefName('solver_rel' + nIndex, constraint.operator);
			nIndex++;
		});
		// Fills options
		for (let sDefName in oOptionsDefNameToAttribute) {
			fillDefName(sDefName, oOptions[oOptionsDefNameToAttribute[sDefName]]);
		}
	};
	/**
	 * Gets def names ranges for filling fields of solver params and options.
	 * @memberof asc_CSolverParams
	 * @param {Workbook} oWbModel
	 */
	asc_CSolverParams.prototype.getDefNames = function(oWbModel) {
		if (!oWbModel) {
			return;
		}
		function fillAttribute (oAttribute, sDefName) {
			let oDefName = mSolverDefNames.get(sDefName);
			if (oDefName) {
				oAttribute = typeof oAttribute === 'boolean' ? oDefName.Ref === 1 : oDefName.Ref;
			}
			return oAttribute;
		}

		const oDependencyFormulas = oWbModel.dependencyFormulas;
		const aWsDefNames = getHiddenDefinedNamesWS(oDependencyFormulas);
		const oDefNameToAttribute = this.getDefNameToAttribute();
		const oOptions = this.getOptions();
		const oOptionsDefNameToAttribute = oOptions.getDefNameToAttribute();
		const mSolverDefNames = new Map();

		if (!aWsDefNames.length) {
			return;
		}
		aWsDefNames.forEach(function(oDefName) {
			if (~oDefName.Name.indexOf("solver_")) {
				mSolverDefNames.set(oDefName.Name, oDefName);
			}
		});
		if (!mSolverDefNames.size) {
			return;
		}
		// Fills solver params
		for (let sDefName in oDefNameToAttribute) {
			this[oDefNameToAttribute[sDefName]] = fillAttribute(this[oDefNameToAttribute[sDefName]], sDefName);
		}
		// Fills constraints.
		const nConstraintsSize = mSolverDefNames.get('solver_num') && mSolverDefNames.get('solver_num').Ref;
		if (nConstraintsSize === undefined) {
			return;
		}
		for (let i = 1; i <= nConstraintsSize; i++) {
			/** @type {{cellRef:string, operator: c_oAscOperator, constraint:string}} */
			const oConstraint = {};
			oConstraint.cellRef = fillAttribute(oConstraint.cellRef, 'solver_lhs' + i);
			oConstraint.operator = fillAttribute(oConstraint.operator, 'solver_rel' + i);
			oConstraint.constraint = fillAttribute(oConstraint.constraint, 'solver_rhs' + i);
			if (oConstraint.cellRef && oConstraint.operator && oConstraint.constraint) {
				this.addConstraint(i, oConstraint);
			}
		}
		// Fills options
		for (let sDefName in oOptionsDefNameToAttribute) {
			oOptions[oOptionsDefNameToAttribute[sDefName]] = fillAttribute(oOptions[oOptionsDefNameToAttribute[sDefName]], sDefName);
		}
	};


	/**
	 * Class representing options for calculating.
	 * @constructor
	 */
	function asc_COptions () {
		const sNumberDecimalSeparator = AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
		// All methods
		this.sConstraintPrecision = '0.000001'.replace('.', sNumberDecimalSeparator);
		this.bAutomaticScaling = false;
		this.bShowIterResults = false;
		this.bIgnoreIntConstriants = false;
		this.sIntOptimal = '1';
		this.sMaxTime = '2147483647';
		this.sIterations = '2147483647';
		this.sMaxSubproblems = '2147483647';
		this.sMaxFeasibleSolution = '2147483647';
		// GRG Nonlinear and Evolutionary
		this.sConvergence = '0.0001'.replace('.', sNumberDecimalSeparator);
		this.nDerivatives = c_oAscDerivativeType.forward;
		this.bMultistart = false; // ?
		this.sPopulationSize = '100';
		this.sRandomSeed = '0';
		this.bRequireBounds = false;
		this.sMutationRate = '0.075'.replace('.', sNumberDecimalSeparator);
		this.sEvoMaxTime = '30';

		this.oDefNameToAttribute = {
			'solver_pre': 'sConstraintPrecision',
			'solver_scl': 'bAutomaticScaling',
			'solver_sho': 'bShowIterResults',
			'solver_rlx': 'bIgnoreIntConstriants',
			'solver_tol': 'sIntOptimal',
			'solver_tim': 'sMaxTime',
			'solver_itr': 'sIterations',
			'solver_nod': 'sMaxSubproblems',
			'solver_mip': 'sMaxFeasibleSolution',
			'solver_cvg': 'sConvergence',
			'solver_drv': 'nDerivatives',
			'solver_msl': 'bMultistart',
			'solver_ssz': 'sPopulationSize',
			'solver_rsd': 'sRandomSeed',
			'solver_rbv': 'bRequireBounds',
			'solver_mrt': 'sMutationRate',
			'solver_mni': 'sEvoMaxTime'
		};
	}

	/**
	 * Returns value of "Constraint Precision" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getConstraintPrecision = function() {
		return this.sConstraintPrecision;
	};
	/**
	 * Returns value of "Use automatic Scaling" parameter.
	 * @memberof asc_COptions
	 * @returns {boolean}
	 */
	asc_COptions.prototype.getAutomaticScaling = function() {
		return this.bAutomaticScaling;
	};
	/**
	 * Returns value of "Show Iteration Results" parameter.
	 * @memberof asc_COptions
	 * @returns {boolean}
	 */
	asc_COptions.prototype.getShowIterResults = function() {
		return this.bShowIterResults;
	};
	/**
	 * Returns value of "Ignore Integer Constraints" parameter.
	 * @memberof asc_COptions
	 * @returns {boolean}
	 */
	asc_COptions.prototype.getIgnoreIntConstraints = function() {
		return this.bIgnoreIntConstriants;
	};
	/**
	 * Returns value of "Integer Optimality (%)" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getIntOptimal = function () {
		return this.sIntOptimal;
	};
	/**
	 * Returns value of "Max Time (Seconds)" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getMaxTime = function () {
		return this.sMaxTime;
	};
	/**
	 * Returns value of "Iterations"
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getIterations = function () {
		return this.sIterations;
	};
	/**
	 * Returns value of "Max Subproblems" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getMaxSubproblems = function () {
		return this.sMaxSubproblems;
	};
	/**
	 * Returns value of "Max Feasible Solutions" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getMaxFeasibleSolution = function () {
		return this.sMaxFeasibleSolution;
	};
	/**
	 * Returns value of "Convergence" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getConvergence = function () {
		return this.sConvergence;
	};
	/**
	 * Returns value of "Derivatives" parameter.
	 * @memberof asc_COptions
	 * @returns {c_oAscDerivativeType}
	 */
	asc_COptions.prototype.getDerivatives = function () {
		return this.nDerivatives;
	};
	/**
	 * Returns value of "Use multistart" parameter.
	 * @memberof asc_COptions
	 * @returns {boolean}
	 */
	asc_COptions.prototype.getMultistart = function () {
		return this.bMultistart;
	};
	/**
	 * Returns value of "Population Size" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getPopulationSize = function () {
		return this.sPopulationSize;
	};
	/**
	 * Returns value of "Random seed" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getRandomSeed = function () {
		return this.sRandomSeed;
	};
	/**
	 * Returns value of "Require Bounds on Variables" parameter.
	 * @memberof asc_COptions
	 * @returns {boolean}
	 */
	asc_COptions.prototype.getRequireBounds = function () {
		return this.bRequireBounds;
	};
	/**
	 * Returns value of "Mutation Rate" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getMutationRate = function () {
		return this.sMutationRate;
	};
	/**
	 * Returns value of "Maximum Time without improvement" parameter.
	 * @memberof asc_COptions
	 * @returns {string}
	 */
	asc_COptions.prototype.getEvoMaxTime = function () {
		return this.sEvoMaxTime;
	};
	/**
	 * Returns object which connecting def name to attribute of class.
	 * @memberof asc_COptions
	 * @returns {{
	 * solver_pre: string, solver_scl: string, solver_sho: string, solver_rlx: string, solver_tol: string,
	 * solver_tim: string,solver_itr: string, solver_nod: string,solver_mip: string, solver_cvg: string,
	 * solver_drv: string, solver_msl: string, solver_ssz: string, solver_rsd: string, solver_rbv: string,
	 * solver_mrt: string, solver_mni: string
	 * }}
	 */
	asc_COptions.prototype.getDefNameToAttribute = function() {
		return this.oDefNameToAttribute;
	};
	/**
	 * Sets value of "Constraint Precision" parameter.
	 * @memberof asc_COptions
	 * @param {string} constraintPrecision
	 */
	asc_COptions.prototype.setConstraintPrecision = function (constraintPrecision) {
		this.sConstraintPrecision = constraintPrecision;
	};
	/**
	 * Sets value of "Use automatic Scaling" parameter.
	 * @memberof asc_COptions
	 * @param {boolean} automaticScaling
	 */
	asc_COptions.prototype.setAutomaticScaling = function (automaticScaling) {
		this.bAutomaticScaling = automaticScaling;
	};
	/**
	 * Sets value of "Show Iteration Results" parameter.
	 * @memberof asc_COptions
	 * @param {boolean} showIterResults
	 */
	asc_COptions.prototype.setShowIterResults = function (showIterResults) {
		this.bShowIterResults = showIterResults;
	};
	/**
	 * Sets value of "Ignore Integer Constraints" parameter.
	 * @memberof asc_COptions
	 * @param {boolean} ignoreIntConstraints
	 */
	asc_COptions.prototype.setIgnoreIntConstraints = function (ignoreIntConstraints) {
		this.bIgnoreIntConstriants = ignoreIntConstraints;
	};
	/**
	 * Sets value of "Integer Optimality (%)" parameter.
	 * @memberof asc_COptions
	 * @param {string} intOptimality
	 */
	asc_COptions.prototype.setIntOptimal = function (intOptimality) {
		this.sIntOptimal = intOptimality;
	};
	/**
	 * Sets value of "Max Time (Seconds)" parameter.
	 * @memberof asc_COptions
	 * @param {string} maxTime
	 */
	asc_COptions.prototype.setMaxTime = function (maxTime) {
		this.sMaxTime = maxTime;
	};
	/**
	 * Sets value of "Iterations" parameter
	 * @memberof asc_COptions
	 * @param {string} iterations
	 */
	asc_COptions.prototype.setIterations = function (iterations) {
		this.sIterations = iterations;
	};
	/**
	 * Sets value of "Max Subproblems" parameter.
	 * @memberof asc_COptions
	 * @param {string} maxSubproblems
	 */
	asc_COptions.prototype.setMaxSubproblems = function (maxSubproblems) {
		this.sMaxSubproblems = maxSubproblems;
	};
	/**
	 * Sets value of "Max Feasible Solution" parameter.
	 * @memberof asc_COptions
	 * @param {string} maxFeasibleSolution
	 */
	asc_COptions.prototype.setMaxFeasibleSolution = function (maxFeasibleSolution) {
		this.sMaxFeasibleSolution = maxFeasibleSolution;
	};
	/**
	 * Sets value of "Convergence" parameter.
	 * @memberof asc_COptions
	 * @param {string} convergence
	 */
	asc_COptions.prototype.setConvergence = function (convergence) {
		this.sConvergence = convergence;
	};
	/**
	 * Sets value of "Derivatives" parameter.
	 * @memberof asc_COptions
	 * @param {c_oAscDerivativeType} derivatives
	 */
	asc_COptions.prototype.setDerivatives = function (derivatives) {
		this.nDerivatives = derivatives;
	};
	/**
	 * Sets value of "Use Multistart" parameter.
	 * @memberof asc_COptions
	 * @param {boolean} multistart
	 */
	asc_COptions.prototype.setMultistart = function (multistart) {
		this.bMultistart = multistart;
	};
	/**
	 * Sets value of "Population Size" parameter.
	 * @memberof asc_COptions
	 * @param {string} populationSize
	 */
	asc_COptions.prototype.setPopulationSize = function (populationSize) {
		this.sPopulationSize = populationSize;
	};
	/**
	 * Sets value of "Random Seed" parameter.
	 * @memberof asc_COptions
	 * @param {string} randomSeed
	 */
	asc_COptions.prototype.setRandomSeed = function (randomSeed) {
		this.sRandomSeed = randomSeed;
	};
	/**
	 * Sets value of "Require Bounds on Variables" parameter.
	 * @memberof asc_COptions
	 * @param {boolean} requireBounds
	 */
	asc_COptions.prototype.setRequireBounds = function (requireBounds) {
		this.bRequireBounds = requireBounds;
	};
	/**
	 * Sets value of "Mutation Rate" parameter.
	 * @memberof asc_COptions
	 * @param {string} mutationRate
	 */
	asc_COptions.prototype.setMutationRate = function (mutationRate) {
		this.sMutationRate = mutationRate;
	};
	/**
	 * Sets value of "Maximum Time without improvement" parameter.
	 * @memberof asc_COptions
	 * @param {string} evoMaxTime
	 */
	asc_COptions.prototype.setEvoMaxTime = function (evoMaxTime) {
		this.sEvoMaxTime = evoMaxTime;
	};
	/**
	 * Resets all options to default value.
	 * @memberof asc_COptions
	 */
	asc_COptions.prototype.resetAll = function() {
		this.setConstraintPrecision('0.000001');
		this.setAutomaticScaling(false);
		this.setShowIterResults(false);
		this.setIgnoreIntConstraints(false);
		this.setIntOptimal('1');
		this.setMaxTime('2147483647');
		this.setIterations('2147483647');
		this.setMaxSubproblems('2147483647');
		this.setMaxFeasibleSolution('2147483647');
		this.setConvergence('0.0001');
		this.setDerivatives(c_oAscDerivativeType.forward);
		this.setMultistart(false);
		this.setPopulationSize('100');
		this.setRandomSeed('0');
		this.setRequireBounds(false);
		this.setMutationRate('0.075');
		this.setEvoMaxTime('30');
	};

	/**
	 * Class representing Solver Results dialogue window data.
	 * @constructor
	 * @returns {asc_CSolverResults}
	 */
	function asc_CSolverResults () {
		this.sStatus = null;
		this.sDescription = null;
		this.nAction = c_oAscSolverResult.keepSolverSolution;
		this.bReturnToSolverParams = false;

		return this;
	}

	/**
	 * Returns status of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @returns {string}
	 */
	asc_CSolverResults.prototype.getStatus = function() {
		return this.sStatus;
	};
	/**
	 * Returns status description of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @returns {string}
	 */
	asc_CSolverResults.prototype.getDescription = function() {
		return this.sDescription;
	};
	/**
	 * Returns current action of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @returns {c_oAscSolverResult}
	 */
	asc_CSolverResults.prototype.getAction = function() {
		return this.nAction;
	};
	/**
	 * Returns state of "Return to solver params" checkbox of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @returns {boolean}
	 */
	asc_CSolverResults.prototype.getReturnToSolverParams = function() {
		return this.bReturnToSolverParams;
	};
	/**
	 * Sets status of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @param {string} sStatus
	 */
	asc_CSolverResults.prototype.setStatus = function(sStatus) {
		this.sStatus = sStatus;
	};
	/**
	 * Sets status description of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @param {string} sDescription
	 */
	asc_CSolverResults.prototype.setDescription = function(sDescription) {
		this.sDescription = sDescription;
	};
	/**
	 * Sets current action of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @param {c_oAscSolverResult} nAction
	 */
	asc_CSolverResults.prototype.setAction = function(nAction) {
		this.nAction = nAction;
	};
	/**
	 * Sets state of "Return to solver params" checkbox of Solver Results dialogue window.
	 * @memberof asc_CSolverResults
	 * @param {boolean} bReturnToSolverParams
	 */
	asc_CSolverResults.prototype.setReturnToSolverParams = function(bReturnToSolverParams) {
		this.bReturnToSolverParams = bReturnToSolverParams;
	};


	// Classes and functions for main logic of Solver feature
	/**
	 * Returns worksheet that has selected cell(s) from parameters.
	 * @param {string} sRef - selected cell(s) reference from parameters
	 * @param {Worksheet} oWs
	 * @returns {Worksheet}
	 */
	function actualWsByRef (sRef, oWs) {
		let oActualWs = oWs;
		if (~sRef.indexOf("!")) {
			const sSheetName = sRef.split("!")[0].replace(/'/g, "");
			if (sSheetName !== oActualWs.getName()) {
				oActualWs = oWs.workbook.getWorksheetByName(sSheetName);
			}
		}

		return oActualWs;
	}

	/**
	 * Converts absolute reference containing the worksheet name to reference with only selected cell(s) name.
	 * @param {string} sRef - selected cell(s) reference from parameters
	 * @returns {string}
	 */
	function convertToAbsoluteRef (sRef) {
		let sConvertedRef = sRef;
		if (~sRef.indexOf("!")) {
			sConvertedRef = sRef.split("!")[1];
		}

		return sConvertedRef;
	}

	/**
	 * Class representing a constraints option for Solver.
	 * @param {Range|Cell} oCell
	 * @param {c_oAscOperator} nOperator
	 * @param {Range|number|string} constraint
	 * @constructor
	 */
	function CConstraint (oCell, nOperator, constraint) {
		this.oCell = oCell;
		this.nOperator = nOperator;
		this.constraint = constraint;
	}
	
	/**
	 * Returns cell reference.
	 * @memberof CConstraint
	 * @returns {Range|Cell}
	 */
	CConstraint.prototype.getCell = function() {
		return this.oCell;
	};
	/**
	 * Returns comparison operator.
	 * @memberof CConstraint
	 * @returns {c_oAscOperator}
	 */
	CConstraint.prototype.getOperator = function() {
		return this.nOperator;
	};
	/**
	 * Returns constraint. Element that comparisons with reference cell.
	 * @memberof CConstraint
	 * @returns {Range|number|string}
	 */
	CConstraint.prototype.getConstraint = function() {
		return this.constraint;
	};

	/**
	 * Class representing logic of solving linear programming by Simplex method.
	 * @param {CSolver} oModel
	 * @constructor
	 */
	function CSimplexTableau (oModel) {
		this.oModel = oModel;

		this.aMatrix = null;
		this.nWidth = 0;
		this.nHeight = 0;
		this.aVariablesIndexes = null;
		this.aConstraintsIndexes = null;

		this.nObjectiveRowIndex = 0;
		this.nRhsColumn = 0;

		this.oVarIndexByCellName = null;
		this.oUnrestrictedVars = null;

		// Solution attributes
		this.nLastColumnId = null;
		this.nLastRowId = null;
		this.aVarIndexesCycle = [];
		this.mRepeatedVars = new Map();

		this.bFeasible = false;
		this.bSolutionIsFound = false;

		this.aVarIndexByRow = null;
		this.aVarIndexByCol = null;

		this.aRowByVarIndex = null;
		this.aColByVarIndex = null;

		this.nPrecision = 1e-8;

		this.nLastElementIndex = 0;

		this.oVariables = null;
		this.nVars = 0;

		this.bBounded = true;
		this.nUnboundedVarIndex = null;

		this.nBranchAndCutIters = 0;
	}

	/**
	 * Initializes a start data for calculating logic.
	 * @memberof CSimplexTableau
	 */
	CSimplexTableau.prototype.init = function () {
		const oModel = this.getModel();
		const oOptions = oModel.getOptions();
		const oChangingCells = oModel.getChangingCell();
		const oBboxChangingCells = oChangingCells.bbox;
		const nTotalCountVariables = oBboxChangingCells.getWidth() * oBboxChangingCells.getHeight();
		const aConstraints = oModel.getConstraints();
		const aSimplexConstraints = this.getSimplexConstraints(aConstraints);
		const nTotalCountConstraints = aSimplexConstraints.length;
		const nOptionPrecision = Number(oOptions.getConstraintPrecision().replace(/,/g, "."));

		if (!isNaN(nOptionPrecision)) {
			this.setPrecision(nOptionPrecision);
		}
		this.setVariables(oChangingCells);
		// Init width and height of matrix.
		this.setWidth(nTotalCountVariables + 1);
		this.setHeight(nTotalCountConstraints + 1);
		// Init indexes of variables and constraints
		this.setConstraintsIndexes(this.fillConstraintsIndexes(nTotalCountConstraints));
		this.setVariablesIndexes(this.fillVariablesIndexes(nTotalCountVariables, nTotalCountConstraints));

		// Init matrix
		const aMatrix = this.createMatrix();

		const nHeight = this.getHeight();
		const nWidth = this.getWidth();
		this.setVarIndexByRow(new Array(nHeight));
		this.setVarIndexByCol(new Array(nWidth));
		this.addVarIndexByRowElem(-1, 0);
		this.addVarIndexByColElem(-1, 0);

		this.setVarsCount(nWidth + nHeight - 2);
		this.setRowByVarIndex(new Array(this.getVarsCount()));
		this.setColByVarIndex(new Array(this.getVarsCount()));

		this.setLastElementIndex(this.getVarsCount());

		// Fills matrix
		this.fillMatrix(aMatrix, aSimplexConstraints);
		// Init last index row and column of matrix
		this.setLastColumnIndex(nWidth - 1);
		this.setLastRowIndex(nHeight - 1);
	};
	/**
	 * Calculates solution using Simplex LP method.
	 * @memberof CSimplexTableau
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 * @see {@link https://en.wikibooks.org/wiki/Operations_Research/The_Simplex_Method#The_Simplex_method|Theory material}
	 */
	CSimplexTableau.prototype.calculate = function () {
		let bStopCalculating = false;

		if (!this.getFeasible() && !bStopCalculating) {
			bStopCalculating = this.phase1();
		}
		if (this.getFeasible() && !bStopCalculating) {
			// current solution is feasible
			bStopCalculating = this.phase2();
		}

		return bStopCalculating;
	};
	/**
	 * Obtains a Basic Feasible Solution (BFS)
	 * Convert a non-standard form tableau to a standard form tableau by eliminating
	 * all negative values on the Right Hand Side (RHS)
	 * This results in a Basic Feasible Solution (BFS)
	 * @memberof CSimplexTableau
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSimplexTableau.prototype.phase1 = function () {
		const oModel = this.getModel();
		const aMatrix = this.getMatrix();
		const nRhsColumnId = this.getRhsColumn();
		const nLastColumnId = this.getLastColumnIndex();
		const nLastRowId = this.getLastRowIndex();
		const aVarIndexesCycle = this.getVarIndexesCycle();

		let nLeavingRowIndex = 0;
		let nRhsValue = -this.getPrecision();
		// Step 1: Find pivot row. Selecting leaving variable (feasibility condition). Basic variable with most negative value.
		for (let i = 1; i <= nLastRowId; i++) {
			const nValue = aMatrix[i][nRhsColumnId];
			if (nValue < nRhsValue) {
				nRhsValue = nValue;
				nLeavingRowIndex = i;
			}
		}
		// If the leaving row isn't found, for 1st phase, found a feasible solution.  Finished 1st phase.
		if (nLeavingRowIndex === 0) {
			this.setFeasible(true);
			return false;
		}

		// Step 2: Find pivot column. Selecting entering variable.
		const aObjectiveRow = aMatrix[this.getObjectiveRowIndex()];
		const aLeavingRow = aMatrix[nLeavingRowIndex];
		let nEnteringColumnIndex = 0;
		let nMaxQuotient = -Infinity;
		for (let i = 1; i <= nLastColumnId; i++) {
			const nCoefficient = aLeavingRow[i];
			if (nCoefficient < -this.getPrecision()) {
				const nQuotient = -aObjectiveRow[i] / nCoefficient;
				if (nMaxQuotient < nQuotient) {
					nMaxQuotient = nQuotient;
					nEnteringColumnIndex = i;
				}
			}
		}
		// If the entering row isn't found, feasible solution isn't found too.
		if (nEnteringColumnIndex === 0) {
			this.setFeasible(false);
			//TODO call api to show error about solution isn't found
			return true;
		}

		// Check for cycles
		const aVarIndexByRow = this.getVarIndexByRow();
		const aVarIndexByCol = this.getVarIndexByCol();
		aVarIndexesCycle.push([aVarIndexByRow[nLeavingRowIndex], aVarIndexByCol[nEnteringColumnIndex]]);
		if (this.checkForCycles(aVarIndexesCycle)) {
			this.setFeasible(false);
			this.clearVarIndexesCycle();
			this.getRepeatedVars().clear();
			//TODO call api to show error about solution isn't found
			return true
		}

		this.pivot(nLeavingRowIndex, nEnteringColumnIndex);
		oModel.increaseCurrentAttempt();

		return false;
	};
	/**
	 * Runs simplex on Initial Basic Feasible Solution (BFS)
	 * Apply simplex to obtain optimal solution used as phase2 of the simplex
	 * @memberof CSimplexTableau
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSimplexTableau.prototype.phase2 = function () {
		const oModel = this.getModel();
		const aMatrix = this.getMatrix();
		const nRhsColumnId = this.getRhsColumn();
		const nLastColumnId = this.getLastColumnIndex();
		const nLastRowId = this.getLastRowIndex();
		const aVarIndexesCycle = this.getVarIndexesCycle();
		const nPrecision = this.getPrecision();
		const aObjectiveRow = aMatrix[this.getObjectiveRowIndex()];

		let nEnteringColumnIndex = 0;
		let nEnteringValue = nPrecision;
		let bReducedCostNegative = false;

		for (let col = 1; col <= nLastColumnId; col++) {
			const nObjectiveCoefValue = aObjectiveRow[col];
			if (nObjectiveCoefValue > nEnteringValue) {
				nEnteringValue = nObjectiveCoefValue;
				nEnteringColumnIndex = col;
				bReducedCostNegative = false;
			}
		}
		// If no entering column could be found we're done with phase 2. Solution has been found.
		if (nEnteringColumnIndex === 0) {
			this.setSolutionIsFound(true);
			oModel.increaseCurrentFeasibleCount();
			return true;
		}
		// Selecting leaving variable
		let nLeavingRowIndex = 0;
		let nMinQuotient = Infinity;

		for (let row = 1; row <= nLastRowId; row++) {
			const aRow = aMatrix[row];
			const nRhsValue = aRow[nRhsColumnId];
			const nColValue = aRow[nEnteringColumnIndex];

			if (-nPrecision < nColValue && nColValue < nPrecision) {
				continue;
			}
			if (nColValue > 0 && nPrecision > nRhsValue && nRhsValue > -nPrecision) {
				nMinQuotient = 0;
				nLeavingRowIndex = row;
				break;
			}
			const nQuotient = bReducedCostNegative ? -nRhsValue / nColValue : nRhsValue / nColValue;
			if (nQuotient > nPrecision && nMinQuotient > nQuotient) {
				nMinQuotient = nQuotient;
				nLeavingRowIndex = row;
			}
		}

		if (nMinQuotient === Infinity) {
			// Optimal value is -Infinity or solution hasn't been found.
			this.setSolutionIsFound(false);
			this.setBounded(false);
			this.setUnboundVarIndex(this.getVarIndexByCol()[nEnteringColumnIndex]);
			//TODO call api to show error about solution isn't found
			return true;
		}

		// Check for cycles
		const aVarIndexByRow = this.getVarIndexByRow();
		const aVarIndexByCol = this.getVarIndexByCol();
		aVarIndexesCycle.push([aVarIndexByRow[nLeavingRowIndex], aVarIndexByCol[nEnteringColumnIndex]]);
		if (this.checkForCycles(aVarIndexesCycle)) {
			this.setFeasible(false);
			this.clearVarIndexesCycle();
			this.getRepeatedVars().clear();
			//TODO call api to show error about solution isn't found
			return true
		}

		this.pivot(nLeavingRowIndex, nEnteringColumnIndex);
		oModel.increaseCurrentAttempt();

		return false;
	};
	/**
	 * Returns the main model object with the necessary data for solving a task by the Simplex method.
	 * @memberof CSimplexTableau
	 * @returns {CSolver}
	 */
	CSimplexTableau.prototype.getModel = function () {
		return this.oModel;
	};
	/**
	 * Returns constraint precision.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getPrecision = function () {
		return this.nPrecision;
	};
	/**
	 * Sets constraint precision.
	 * @memberof CSimplexTableau
	 * @param {number} nPrecision
	 */
	CSimplexTableau.prototype.setPrecision = function (nPrecision) {
		this.nPrecision = nPrecision;
	};
	/**
	 * Returns variables of model. It's the changing variable cells.
	 * @memberof CSimplexTableau
	 * @returns {Range}
	 */
	CSimplexTableau.prototype.getVariables = function () {
		return this.oVariables;
	};
	/**
	 * Sets variables of model. It's the changing variable cells.
	 * @memberof CSimplexTableau
	 * @param {Range} oVariables
	 */
	CSimplexTableau.prototype.setVariables = function (oVariables) {
		this.oVariables = oVariables;
	};
	/**
	 * Returns width of matrix.
	 * @memberof
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getWidth = function () {
		return this.nWidth;
	};
	/**
	 * Sets width of matrix.
	 * @memberof
	 * @param {number} nWidth
	 */
	CSimplexTableau.prototype.setWidth = function (nWidth) {
		this.nWidth = nWidth;
	};
	/**
	 * Creates new array with constraints suitable for simplex calculate logic.
	 * Method extracts ranges of ref cells and values of constraints for making a flat array.
	 * @memberof CSimplexTableau
	 * @param {CConstraint[]}aModelConstraints
	 * @returns {CConstraint[]}
	 */
	CSimplexTableau.prototype.getSimplexConstraints = function (aModelConstraints) {
		/** @type {CConstraint[]} */
		const aSimplexConstraints = [];

		for (let index = 0, length = aModelConstraints.length; index < length; index++) {
			const oConstraint = aModelConstraints[index];
			const nOperator = oConstraint.getOperator();
			const oRefCellsRange = oConstraint.getCell();
			const constraintData = oConstraint.getConstraint();
			let nConstraintIter = 0;
			let oConstraintBbox, bVertical, oConstraintWs;

			if (typeof constraintData === 'string') {
				continue;
			}
			if (typeof constraintData === 'object') {
				oConstraintBbox = constraintData.bbox;
				bVertical = oConstraintBbox.c1 === oConstraintBbox.c2;
				oConstraintWs = constraintData.worksheet;
			}
			oRefCellsRange._foreachNoEmpty(function (oRefCell) {
				/** @type {Cell} */
				const oCell = oRefCell.clone();
				let oNewConstraint, nConstraintVal;

				if (oRefCell.isFormula()) {
					oCell.formulaParsed = oRefCell.getFormulaParsed().clone();
				}
				if (typeof constraintData === 'object') {
					const nRowConstraint = bVertical ? oConstraintBbox.r1 + nConstraintIter : oConstraintBbox.r1;
					const nColConstraint = bVertical ? oConstraintBbox.c1 : oConstraintBbox.c1 + nConstraintIter;
					if (nRowConstraint <= oConstraintBbox.r2 && nColConstraint <= oConstraintBbox.c2) {
						oConstraintWs._getCell(nRowConstraint, nColConstraint, function (oConstraintCell) {
							nConstraintVal = oConstraintCell.getNumberValue();
						});
						nConstraintIter++;
					}
				} else {
					nConstraintVal = constraintData;
				}
				if (nOperator === c_oAscOperator['=']) { // Changes the equal ("=") constraint to 2 constraints - ">=" and "<=".
					// Firstly adds a constraint with >= operator, then add  <=.
					let nNewOperator = c_oAscOperator['>='];
					oNewConstraint = new CConstraint(oCell, nNewOperator, nConstraintVal);
					aSimplexConstraints.push(oNewConstraint);
					nNewOperator = c_oAscOperator['<='];
					oNewConstraint = new CConstraint(oCell, nNewOperator, nConstraintVal);
					aSimplexConstraints.push(oNewConstraint);
				} else {
					oNewConstraint = new CConstraint(oCell, nOperator, nConstraintVal);
					aSimplexConstraints.push(oNewConstraint);
				}
			});
		}

		return aSimplexConstraints;
	};
	/**
	 * Returns height of matrix.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getHeight = function () {
		return this.nHeight;
	};
	/**
	 * Sets height of matrix.
	 * @memberof CSimplexTableau
	 * @param {number} nHeight
	 */
	CSimplexTableau.prototype.setHeight = function (nHeight) {
		this.nHeight = nHeight;
	};
	/**
	 * Returns array with indexes of constraints.
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getConstraintsIndexes = function () {
		return this.aConstraintsIndexes;
	};
	/**
	 * Sets array of constraints indexes
	 * @memberof CSimplexTableau
	 * @param {number[]} aConstraintsIndexes
	 */
	CSimplexTableau.prototype.setConstraintsIndexes = function (aConstraintsIndexes) {
		this.aConstraintsIndexes = aConstraintsIndexes;
	};
	/**
	 * Fills array with indexes of constraints.
	 * @memberof CSimplexTableau
	 * @param {number} nTotalCountConstraints
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.fillConstraintsIndexes = function (nTotalCountConstraints) {
		const aConstraintsIndexes = [];

		for (let i = 0; i < nTotalCountConstraints; i++) {
			aConstraintsIndexes.push(i);
		}

		return aConstraintsIndexes;
	};
	/**
	 * Returns array with indexes of variables.
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getVariablesIndexes = function () {
		return this.aVariablesIndexes;
	};
	/**
	 * Sets array of constraints indexes
	 * @memberof CSimplexTableau
	 * @param {number[]} aVariablesIndexes
	 */
	CSimplexTableau.prototype.setVariablesIndexes = function (aVariablesIndexes) {
		this.aVariablesIndexes = aVariablesIndexes;
	};
	/**
	 * Fills array with indexes of variables.
	 * @memberof CSimplexTableau
	 * @param {number} nTotalCountVariables
	 * @param {number} nLastIndexElement
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.fillVariablesIndexes = function (nTotalCountVariables, nLastIndexElement) {
		const aVariablesIndexes = [];

		for (let j = 0; j < nTotalCountVariables; j++) {
			aVariablesIndexes.push(nLastIndexElement++);
		}

		return aVariablesIndexes;
	};
	/**
	 * @memberof CSimplexTableau
	 * @returns {number[][]}
	 */
	CSimplexTableau.prototype.getMatrix = function () {
		return this.aMatrix;
	};
	/**
	 * Sets empty matrix that needs to be filled using addRow method.
	 * @memberof CSimplexTableau
	 * @param {[]}aMatrix
	 */
	CSimplexTableau.prototype.setMatrix = function (aMatrix) {
		this.aMatrix = aMatrix;
	};
	/**
	 * Adds rows for filling matrix.
	 * @memberof CSimplexTableau
	 * @param {number[]} aRow
	 * @param {number} nRowIndex
	 */
	CSimplexTableau.prototype.addRow = function (aRow, nRowIndex) {
		const aMatrix = this.getMatrix();
		aMatrix[nRowIndex] = aRow;
	};
	/**
	 * Builds an empty matrix.
	 * @memberof CSimplexTableau
	 * @returns {number[][]}
	 */
	CSimplexTableau.prototype.createMatrix = function () {
		const nWidth = this.getWidth();
		const nHeight = this.getHeight();
		const tmpRow = new Array(nWidth);
		for (let i = 0, length = nWidth; i < length; i++) {
			tmpRow[i] = 0;
		}
		this.setMatrix(new Array(nHeight));
		for (let j = 0, length = nHeight; j < length; j++) {
			this.addRow(tmpRow.slice(), j);
		}

		return this.getMatrix();
	};
	/**
	 * Extracts arguments from parserFormula.
	 * @memberof CSimplexTableau
	 * @param {[]} aOutStackFormula - arguments of formula
	 * @param {boolean} bIsEmpty - flag recognizes whether empty arguments are allowed
	 * @returns {Cell[]}
	 */
	CSimplexTableau.prototype.getArgsFormula = function (aOutStackFormula, bIsEmpty) {
		const aArgsFormula = [];

		AscCommonExcel.foreachRefElements(function (oRange) {
			oRange._foreachNoEmpty(function (oCell) {
				if (oCell.isFormula() || oCell.getNumberValue() || bIsEmpty) {
					const oTempCell = oCell.clone();
					if (oCell.isFormula()) {
						oTempCell.formulaParsed = oCell.getFormulaParsed().clone();
					}
					aArgsFormula.push(oTempCell);
				}
			})
		}, aOutStackFormula);

		return aArgsFormula;
	};
	/**
	 * Returns an object that stores the variable index by cell name key.
	 * @memberof CSimplexTableau
	 * @returns {{}}
	 */
	CSimplexTableau.prototype.getVarIndexByCellName = function () {
		if (!this.oVarIndexByCellName) {
			this.oVarIndexByCellName = {};
		}

		return this.oVarIndexByCellName;
	};
	/**
	 * Fills matrix by variable and constraints data.
	 * @memberof CSimplexTableau
	 * @param {number[][]} aMatrix
	 * @param {CConstraint[]} aConstraints
	 */
	CSimplexTableau.prototype.fillMatrix = function (aMatrix, aConstraints) {
		const oThis = this;
		const aConstraintsIndexes = this.getConstraintsIndexes();
		const aVariablesIndexes = this.getVariablesIndexes();
		const aObjectiveFuncRow = aMatrix[0];
		const oModel = this.getModel();
		const oVariables = this.getVariables();
		const nCoeff = oModel.getOptimizeResultTo() === c_oAscOptimizeTo.min ? -1 : 1;
		// Extracts coefficients from objective function
		const oObjectiveFunc = oModel.getParsedFormula();
		const aArgsObjectiveFunc = this.getArgsFormula(oObjectiveFunc.outStack, false);

		// Links variable with variable's index
		let nIter = 0;
		const oVarIndexByCellName = this.getVarIndexByCellName();
		oVariables._foreachNoEmpty(function (oCell) {
			oVarIndexByCellName[oCell.getName()] = aVariablesIndexes[nIter];
			nIter++;
		});
		// Fills the matrix's first row coefficients of the objective function
		for (let i = 0, length = aArgsObjectiveFunc.length; i < length; i++) {
			function getVal (nValue) {
				aObjectiveFuncRow[i + 1] = nValue * nCoeff;
				oThis.addRowByVarIndexElem(-1, nVarIndex);
				oThis.addColByVarIndexElem(i + 1, nVarIndex);
				oThis.addVarIndexByColElem(nVarIndex, i + 1);
			}
			const oArg = aArgsObjectiveFunc[i];
			let nVarIndex = aVariablesIndexes[i];
			if (oArg.isFormula() && !oArg.getNumberValue()) { // Tries to find the coefficient in linked cells
				const aArgsFormula = this.getArgsFormula(oArg.getFormulaParsed().outStack, false);
				aArgsFormula.forEach(function (oCell) {
					const nValue = oCell.getNumberValue();
					if (nValue) {
						getVal(nValue);
					}
				});
			} else if (oArg.getNumberValue()) {
				getVal(oArg.getNumberValue());
			}
		}
		// Fills remain matrix's rows by constraints
		const FIRST_COLUMN_INDEX = 0;
		const aColByVarIndex = this.getColByVarIndex();
		let nRowIndex = 1;
		let coefficient = 1;
		for (let j = 0, length = aConstraints.length; j < length; j++) {
			function getVariable(oCell) {
				const nColumn = aColByVarIndex[oVarIndexByCellName[oCell.getName()]];
				aRow[nColumn] = nOperator === c_oAscOperator['<='] ? coefficient : -coefficient;
			}
			const oConstraint = aConstraints[j];
			const nOperator = oConstraint.getOperator();
			const oRefCell = oConstraint.getCell();
			/**@type {number} */
			const nConstraintValue = oConstraint.getConstraint();
			const nConstraintIndex = aConstraintsIndexes[j];
			this.addRowByVarIndexElem(nRowIndex, nConstraintIndex);
			this.addColByVarIndexElem(-1, nConstraintIndex);
			this.addVarIndexByRowElem(nConstraintIndex, nRowIndex);

			const aRow = aMatrix[nRowIndex++];
			// Work with terms
			if (oRefCell.isFormula()) {
				const aArgsFormula = this.getArgsFormula(oRefCell.getFormulaParsed().outStack, true);
				aArgsFormula.forEach(function (oLinkedCell) {
					if (oVariables.bbox.contains(oLinkedCell.nCol, oLinkedCell.nRow)) {
						if (oLinkedCell.isFormula()) {// Try to find coefficient for variable
							const aOutStack = oLinkedCell.getFormulaParsed().outStack;
							const aArgsLinkFormula = oThis.getArgsFormula(aOutStack, false);
							aArgsLinkFormula.forEach(function (oCell) {
								if (oCell.getNumberValue()) {
									coefficient = oCell.getNumberValue();
								}
							});
						}
						getVariable(oLinkedCell);
					}
				});
			} else if (oVariables.bbox.contains(oRefCell.nCol, oRefCell.nRow)) {
				getVariable(oRefCell);
			}
			aRow[FIRST_COLUMN_INDEX] = nOperator === c_oAscOperator['<='] ? nConstraintValue : -nConstraintValue;
		}
	};
	/**
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getVarIndexByRow = function () {
		return this.aVarIndexByRow;
	};
	/**
	 * Sets an empty array for attribute aVarIndexByRow
	 * @memberof CSimplexTableau
	 * @param {number[]} aVarIndexByRow
	 */
	CSimplexTableau.prototype.setVarIndexByRow = function (aVarIndexByRow) {
		this.aVarIndexByRow = aVarIndexByRow;
	};
	/**
	 * Adds element to varIndexByRow array.
	 * @memberof CSimplexTableau
	 * @param {number} nValue
	 * @param {number} nIndex
	 */
	CSimplexTableau.prototype.addVarIndexByRowElem = function (nValue, nIndex) {
		this.aVarIndexByRow[nIndex] = nValue;
	};
	/**
	 * Returns array of indexes of variables by matrix's columns.
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getVarIndexByCol = function () {
		return this.aVarIndexByCol;
	};
	/**
	 * Sets an empty array for attribute aVarIndexByCol.
	 * @memberOf CSimplexTableau
	 * @param {[]} aVarIndexByCol
	 */
	CSimplexTableau.prototype.setVarIndexByCol = function (aVarIndexByCol) {
		this.aVarIndexByCol = aVarIndexByCol;
	};
	/**
	 * Adds element to varIndexByCol array.
	 * @memberof CSimplexTableau
	 * @param {number} nValue
	 * @param {number} nIndex
	 */
	CSimplexTableau.prototype.addVarIndexByColElem = function (nValue, nIndex) {
		this.aVarIndexByCol[nIndex] = nValue;
	};
	/**
	 * Returns total count of variables including constraints.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getVarsCount = function () {
		return this.nVars;
	};
	/**
	 * Sets total count of variables including constraints.
	 * @memberof CSimplexTableau
	 * @param {number} nVars
	 */
	CSimplexTableau.prototype.setVarsCount = function (nVars) {
		this.nVars = nVars;
	};
	/**
	 * Returns array of row's indexes by variable' index. Shows which row of matrix located variable.
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getRowByVarIndex = function () {
		return this.aRowByVarIndex;
	};
	/**
	 * Sets empty array of row's indexes by variable' index. Shows which row of matrix located variable.
	 * @memberof CSimplexTableau
	 * @param {[]}aRowByVarIndex
	 */
	CSimplexTableau.prototype.setRowByVarIndex = function (aRowByVarIndex) {
		this.aRowByVarIndex = aRowByVarIndex;
	};
	/**
	 * Adds element to array of row's indexes by variable' index. Shows which row of matrix located variable.
	 * @memberof CSimplexTableau
	 * @param {number} nValue
	 * @param {number} nIndex
	 */
	CSimplexTableau.prototype.addRowByVarIndexElem = function (nValue, nIndex) {
		this.aRowByVarIndex[nIndex] = nValue;
	};
	/**
	 * Returns array of column's indexes by variable' index. Shows which column of matrix located variable.
	 * @memberof CSimplexTableau
	 * @returns {number[]}
	 */
	CSimplexTableau.prototype.getColByVarIndex = function () {
		return this.aColByVarIndex;
	};
	/**
	 * Sets empty array of column's indexes by variable' index. Shows which column of matrix located variable.
	 * @param {[]} aColByVarIndex
	 */
	CSimplexTableau.prototype.setColByVarIndex = function (aColByVarIndex) {
		this.aColByVarIndex = aColByVarIndex;
	};
	/**
	 * Adds element to array of column's indexes by variable' index. Shows which column of matrix located variable.
	 * @memberof CSimplexTableau
	 * @param {number} nValue
	 * @param {number} nIndex
	 */
	CSimplexTableau.prototype.addColByVarIndexElem = function (nValue, nIndex) {
		this.aColByVarIndex[nIndex] = nValue;
	};
	/**
	 * Returns index of the matrix's last element.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getLastElementIndex = function () {
		return this.nLastElementIndex;
	};
	/**
	 * Sets index of the matrix's last element.
	 * @memberof CSimplexTableau
	 * @param {number} nLastElementIndex
	 */
	CSimplexTableau.prototype.setLastElementIndex = function (nLastElementIndex) {
		this.nLastElementIndex = nLastElementIndex;
	};
	/**
	 * Returns the index of last column in matrix.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getLastColumnIndex = function () {
		return this.nLastColumnId
	};
	/**
	 * Sets the index of last column in matrix.
	 * @memberof CSimplexTableau
	 * @param {number} nLastColumnIndex
	 */
	CSimplexTableau.prototype.setLastColumnIndex = function (nLastColumnIndex) {
		this.nLastColumnId = nLastColumnIndex;
	};
	/**
	 * Returns the index of last row in matrix.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getLastRowIndex = function () {
		return this.nLastRowId;
	};
	/**
	 * Sets the index of last row in matrix.
	 * @memberof CSimplexTableau
	 * @param {number} nLastRowIndex
	 */
	CSimplexTableau.prototype.setLastRowIndex = function (nLastRowIndex) {
		this.nLastRowId = nLastRowIndex;
	};
	/**
	 * Returns bounded attribute
	 * @memberof CSimplexTableau
	 * @returns {boolean}
	 */
	CSimplexTableau.prototype.getBounded = function () {
		return this.bBounded;
	};
	/**
	 * Sets bounded attribute
	 * @memberof CSimplexTableau
	 * @param {boolean} bBounded
	 */
	CSimplexTableau.prototype.setBounded = function (bBounded) {
		this.bBounded = bBounded;
	};
	/**
	 * Returns feasible attribute
	 * @memberof CSimplexTableau
	 * @returns {boolean}
	 */
	CSimplexTableau.prototype.getFeasible = function () {
		return this.bFeasible;
	};
	/**
	 * Sets feasible attribute
	 * @memberof CSimplexTableau
	 * @param {boolean} bFeasible
	 */
	CSimplexTableau.prototype.setFeasible = function (bFeasible) {
		this.bFeasible = bFeasible;
	};
	/**
	 * Returns the index's column with the right-hand side value of the constraint or objective function.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getRhsColumn = function () {
		return this.nRhsColumn;
	};
	/**
	 * Returns an array with the index's variables in cycle.
	 * @memberof CSimplexTableau
	 * @returns {[number,number][]}
	 */
	CSimplexTableau.prototype.getVarIndexesCycle = function () {
		return this.aVarIndexesCycle;
	};
	/**
	 * Clears elements of array with the index's variables in cycle
	 */
	CSimplexTableau.prototype.clearVarIndexesCycle = function () {
		this.aVarIndexesCycle = [];
	};
	/**
	 * Returns repeated basic vars.
	 * This Map is used for checking basic variables are repeated.
	 * If the whole basic table is repeated. It's a cycle
	 * The map stores here indexes of repeated basic variables.
	 * @memberof CSimplexTableau
	 * @returns {Map<number, []>}
	 */
	CSimplexTableau.prototype.getRepeatedVars = function () {
		return this.mRepeatedVars;
	};
	/**
	 * Checks whether the same sequence of basic variables occurred twice in a row.
	 * @memberOf CSimplexTableau
	 * @param {[number,number][]} aVarIndexesCycle
	 * @returns {boolean}
	 */
	CSimplexTableau.prototype.checkForCycles = function (aVarIndexesCycle) {
		const mRepeatedVars = this.getRepeatedVars();

		for (let firstIndex = 0, len1 = aVarIndexesCycle.length - 1; firstIndex < len1; firstIndex++) {
			for (let secondIndex = firstIndex + 1, len2 = aVarIndexesCycle.length; secondIndex < len2; secondIndex++) {
				const aFirstVarIndexes = aVarIndexesCycle[firstIndex];
				const aSecondVarIndexes = aVarIndexesCycle[secondIndex];
				if (aFirstVarIndexes[0] === aSecondVarIndexes[0] && aFirstVarIndexes[1] === aSecondVarIndexes[1]) {
					// Continues the logic of recognizing a cycle only when the whole basic variable is repeated the second time in a row.
					if (secondIndex - firstIndex > len2 - secondIndex) {
						// Check and fill (if it's missing) the indexes of repeated basic variables.
						if (!mRepeatedVars.has(firstIndex + secondIndex)) {
							mRepeatedVars.set(firstIndex + secondIndex, [firstIndex, secondIndex]);
						}
						break;
					}
					let bCycleFound = true;
					// Checks whether the last indexes of the basic variable from the first loop (presumably)
					// and the last indexes from the whole array of passed basic variables are equal.
					const aFirstLastIdsVar = aVarIndexesCycle[secondIndex - 1];
					const aSecondLastIdsVar = aVarIndexesCycle[len2 - 1];
					if (aFirstLastIdsVar[0] !== aSecondLastIdsVar[0] && aFirstLastIdsVar[1] !== aSecondLastIdsVar[1]) {
						bCycleFound = false;
						continue;
					}
					// Check and fill (if it's missing) the last  indexed  of repeated basic variables.
					if (!mRepeatedVars.has((secondIndex - 1) - (len2 - 1))) {
						mRepeatedVars.set((secondIndex - 1) - (len2 - 1), [secondIndex - 1, len2 - 1]);
					}
					if (bCycleFound && len2 - secondIndex === mRepeatedVars.size) {
						return true;
					}
				}
			}
		}

		return false;
	};
	/**
	 * Executes pivot operations over a matrix, on a given row and column.
	 * @memberof CSimplexTableau
	 * @param {number} nLeavingRowIndex
	 * @param {number} nEnteringColumnIndex
	 */
	CSimplexTableau.prototype.pivot = function (nLeavingRowIndex, nEnteringColumnIndex) {
		const aMatrix = this.getMatrix();
		const aPivotRow = aMatrix[nLeavingRowIndex];
		const nQuotient = aMatrix[nLeavingRowIndex][nEnteringColumnIndex];

		const nLastRow = this.getLastRowIndex();
		const nLastColumn = this.getLastColumnIndex();

		const aVarIndexByRow = this.getVarIndexByRow();
		const aVarIndexByCol = this.getVarIndexByCol();
		const aRowByVarIndex = this.getRowByVarIndex();
		const aColByVarIndex = this.getColByVarIndex();

		const nLeavingBasicIndex = aVarIndexByRow[nLeavingRowIndex];
		const nEnteringBasicIndex = aVarIndexByCol[nEnteringColumnIndex];

		aVarIndexByRow[nLeavingRowIndex] = nEnteringBasicIndex;
		aVarIndexByCol[nEnteringColumnIndex] = nLeavingBasicIndex;
		aRowByVarIndex[nEnteringBasicIndex] = nLeavingRowIndex;
		aRowByVarIndex[nLeavingBasicIndex] = -1;
		aColByVarIndex[nEnteringBasicIndex] = -1;
		aColByVarIndex[nLeavingBasicIndex] = nEnteringColumnIndex;
		// Step 1: Transforming pivot row. Calculate new pivot row using formula: newPivotRowElem = oldPivotRowElem / quotient (leading element)
		/** @type {number[]} */
		const aNonZeroColumns = [];
		let nNonZeroColumnId = 0;
		for (let col = 0; col <= nLastColumn; col++) {
			// Checks that the value of pivotRow isn't zero, considering variation in calculation.
			if (!(aPivotRow[col] >= -1e-16 && aPivotRow[col] <= 1e-16)) {
				aPivotRow[col] /= nQuotient;
				aNonZeroColumns[nNonZeroColumnId++] = col;
			} else {
				aPivotRow[col] = 0;
			}
		}
		aPivotRow[nEnteringColumnIndex] = 1 / nQuotient;

		// Step 2: Calculate new rows using formula: newRow = oldRow - oldCoeffOfEnteringVar * newPivotRow
		for (let row = 0; row <= nLastRow; row++) {
			if (row === nLeavingRowIndex) {
				continue;
			}
			// Checks that the value of row isn't zero, considering variation in calculation.
			if (aMatrix[row][nEnteringColumnIndex] >= -1e-16 && aMatrix[row][nEnteringColumnIndex] <= 1e-16) {
				continue;
			}
			const aRow = aMatrix[row];
			const nCoefficient = aRow[nEnteringColumnIndex];
			if (!(nCoefficient >= -1e-16 && nCoefficient <= 1e-16)) {
				for (let nonZeroColId = 0; nonZeroColId < nNonZeroColumnId; nonZeroColId++) {
					let nCol = aNonZeroColumns[nonZeroColId];
					const nPivotRowElem = aPivotRow[nCol];
					if (!(nPivotRowElem >= -1e-16 && nPivotRowElem <= 1e-16)) {
						aRow[nCol] -= nCoefficient * nPivotRowElem;
					} else if (nPivotRowElem !== 0) {
						aPivotRow[nCol] = 0;
					}
				}
				aRow[nEnteringColumnIndex] = -nCoefficient / nQuotient;
			} else if (nCoefficient !== 0) {
				aRow[nEnteringColumnIndex] = 0;
			}
		}
	};
	/**
	 * Returns index of row with objective function data.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getObjectiveRowIndex = function () {
		return this.nObjectiveRowIndex;
	};
	/**
	 * Returns a flag whether a solution has been found.
	 * @memberof CSimplexTableau
	 * @returns {boolean}
	 */
	CSimplexTableau.prototype.getSolutionIsFound = function () {
		return this.bSolutionIsFound;
	};
	/**
	 * Sets a flag whether a solution has been found.
	 * @memberof CSimplexTableau
	 * @param {boolean} bSolutionIsFound
	 */
	CSimplexTableau.prototype.setSolutionIsFound = function (bSolutionIsFound) {
		this.bSolutionIsFound = bSolutionIsFound;
	};
	/**
	 * Returns unbounded index of variable.
	 * @memberof CSimplexTableau
	 * @returns {number}
	 */
	CSimplexTableau.prototype.getUnboundVarIndex = function () {
		return this.nUnboundedVarIndex;
	};
	/**
	 * Sets unbounded index of variable.
	 * @memberof CSimplexTableau
	 * @param {number} nUnboundVarIndex
	 */
	CSimplexTableau.prototype.setUnboundVarIndex = function (nUnboundVarIndex) {
		this.nUnboundedVarIndex = nUnboundVarIndex;
	};
	/**
	 * Updates "Changing variables Cells" values.
	 * @memberof CSimplexTableau
	 */
	CSimplexTableau.prototype.updateVariableValues = function () {
		const oVariablesCells = this.getVariables();
		const nRoundingCoefficient = Math.round(1 / this.getPrecision());
		const oVarIndexByCellName = this.getVarIndexByCellName();
		const aRowByVarIndex = this.getRowByVarIndex();
		const nRhsColumn = this.getRhsColumn();
		const aMatrix = this.getMatrix();
		const sRegNumDecimalSeparator = this.getModel().getRegNumDecimalSeparator();

		oVariablesCells._foreachNoEmpty(function (oCell) {
			const sCellName = oCell.getName();
			const nVarIndex = oVarIndexByCellName[sCellName];
			const nRowId = aRowByVarIndex[nVarIndex];
			if (nRowId !== -1) {
				const nVarValue =  aMatrix[nRowId][nRhsColumn];
				const nResult = Math.round((nVarValue + Number.EPSILON) * nRoundingCoefficient) / nRoundingCoefficient;
				oCell.setValue(String(nResult).replace('.', sRegNumDecimalSeparator));
			}
		});
	};

	// Main class of solver feature
	/**
	 * Class representing a solver feature.
	 * @param {asc_CSolverParams} oParams - solver parameters
	 * @param {Worksheet} oWs
	 * @constructor
	 */
	function CSolver (oParams, oWs) {
		const oParsedFormula = this.convertToCell(oParams.getObjectiveFunction(), oWs).getFormulaParsed();
		const oChangingCellsWs = actualWsByRef(oParams.getChangingCells(), oWs);
		const sChangingCells = convertToAbsoluteRef(oParams.getChangingCells());
		CBaseAnalysis.call(this, oParsedFormula, oChangingCellsWs.getRange2(sChangingCells));

		// Solver parameters
		this.nOptimizeResultTo = oParams.getOptimizeResultTo();
		this.nValueOf = oParams.getValueOf() !== null ? parseFloat(oParams.getValueOf().replace(/,/g, ".")) : null;
		this.aConstraints = this.initConstraints(oParams.getConstraints(), oWs);
		this.bIsVarsNonNegative = oParams.getVariablesNonNegative();
		this.nSolvingMethod = oParams.getSolvingMethod();

		// Attributes for calculating logic
		this.nStartTime = null;
		this.nCurrentSubProblem = 0;
		this.nCurrentFeasibleCount = 0;
		this.nGradient = null;
		this.oSimplexTableau = null;
		this.oStartChangingCells = null;

		// Calculating option
		this.oOptions = oParams.getOptions();

		// Updating nMaxIterations attribute of super class to value from calculating options.
		this.setMaxIterations(parseFloat(this.oOptions.getIterations()));
	}

	CSolver.prototype = Object.create(CBaseAnalysis.prototype);
	CSolver.prototype.constructor = CSolver;
	/**
	 * Prepares data for calculating.
	 * @memberof CSolver
	 */
	CSolver.prototype.prepare = function () {
		const oThis = this;
		const aConstraints = this.getConstraints();
		const nSolutionMethod = this.getSolvingMethod();
		const oChangingCells = this.getChangingCell();

		// Fills empty cells to 0 value for changing cells
		oChangingCells._foreachNoEmpty(function (oCell) {
			oThis.setStartChangingCells(oCell.getName(), oCell.getValueWithoutFormat()); // Saves original data
			if (oCell.getNumberValue() === null) {
				if (oCell.getType() !== CellValueType.Number) {
					oCell.setTypeInternal(CellValueType.Number);
				}
				oCell.setValueNumberInternal(0);
			}
		});
		switch (nSolutionMethod) {
			case c_oAscSolvingMethod.grgNonlinear:
				break;
			case c_oAscSolvingMethod.simplexLP:
				this.setSimplexTableau(new CSimplexTableau(this))
				const oSimplexTableau = this.getSimplexTableau();
				oSimplexTableau.init(aConstraints);
				break;
			case  c_oAscSolvingMethod.evolutionary:
				break;
		}
	};
	/**
	 * Main logic of solver calculating.
	 * Runs only in sync or async loop.
	 * @memberof CSolver
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSolver.prototype.calculate = function () {
		if (this.getIsPause()) {
			return true;
		}

		const nSolutionMethod = this.getSolvingMethod();
		let bCompleteCalculation = false;

		if (this.isFinishCalculating()) {
			return true;
		}
		switch (nSolutionMethod) {
			case c_oAscSolvingMethod.grgNonlinear:
				bCompleteCalculation = this.grgOptimization();
				break;
			case c_oAscSolvingMethod.simplexLP:
				bCompleteCalculation = this.simplexOptimization();
				break;
			case c_oAscSolvingMethod.evolutionary:
				bCompleteCalculation = this.evolutionOptimization();
				break;
		}
		if (this.getIsSingleStep()) {
			this.setIsPause(true);
			this.setIsSingleStep(false);
			return true;
		}

		return bCompleteCalculation;
	};
	/**
	 * Converts cell reference from UI to Cell object.
	 * @memberof CSolver
	 * @param {string} sCellRef
	 * @param {Worksheet}oWs
	 * @returns {Cell}
	 */
	CSolver.prototype.convertToCell = function (sCellRef, oWs) {
		const oCellWs = actualWsByRef(sCellRef, oWs);
		const sCellActualRef = convertToAbsoluteRef(sCellRef);
		const oCellRange = oCellWs.getCell2(sCellActualRef);
		let oCell = null;

		oCellWs._getCell(oCellRange.bbox.r1, oCellRange.bbox.c1, function (oElem) {
			oCell = oElem;
		});

		return oCell;
	};
	/**
	 * Initializes the constraints array with necessary for solving data.
	 * @memberof CSolver
	 * @param {Map} oConstraints
	 * @param {Worksheet} oWs
	 * @returns {CConstraint[]}
	 */
	CSolver.prototype.initConstraints = function (oConstraints, oWs) {
		const aExcludeWords = ['integer', 'binary', 'AllDifferent'];
		const aConstraints = [];

		oConstraints.forEach(function (oConstraint) {
			const oConstraintsWs = actualWsByRef(oConstraint.constraint, oWs);
			const sConstraintsActualRef = convertToAbsoluteRef(oConstraint.constraint);
			const oCellRefWs = actualWsByRef(oConstraint.cellRef, oWs);
			const sCellRefActualRef = convertToAbsoluteRef(oConstraint.cellRef);
			let constraintData = oConstraintsWs.getRange2(sConstraintsActualRef);
			if (constraintData == null) {
				constraintData = aExcludeWords.includes(oConstraint.constraint) ? oConstraint.constraint :
					Number(oConstraint.constraint.replace(/,/g, "."));
			}
			aConstraints.push(new CConstraint(oCellRefWs.getRange2(sCellRefActualRef), oConstraint.operator, constraintData));
		});
		if (this.getVariablesNonNegative()) {
			aConstraints.push(new CConstraint(this.getChangingCell(), c_oAscOperator['>='], 0));
		}

		return aConstraints;
	};
	/**
	 * Calculates and returns an array of constraint results.
	 * @memberof CSolver
	 * @returns {boolean[]}
	 */
	CSolver.prototype.calculateConstraints = function () { // todo Need to rework. Is it needed?
		const oThis = this;
		/** @type {boolean[]} */
		const aConstraintsResult = [];
		const aConstraints = this.getConstraints();
		const oOptions = this.getOptions();
		const nConstraintPrecision = Number(oOptions.getConstraintPrecision().replace(/,/g, "."));
		const bIgnoreIntConstraints = oOptions.getIgnoreIntConstraints();

		for (let i = 0, length = aConstraints.length; i < length; i++) {
			const oConstraintCell = aConstraints[i].getCell();
			let bResult = false;
			/** @type {Cell} */
			let oPrevConstraintCell = null;

			oConstraintCell._foreachNoEmpty(function (oElem, nIndex, nCol, nStartRow) {
				if (oElem.getNumberValue() !== null) {
					const constraintData = aConstraints[i].getConstraint();
					const nOperator = aConstraints[i].getOperator();
					const oWs = oElem.ws;
					let constraintValue = typeof constraintData === 'object' ? oThis.convertToCell(constraintData.getName(), oWs).getNumberValue() : constraintData;
					if (constraintValue === 'integer' && !bIgnoreIntConstraints) {
						bResult = oElem.getNumberValue() === parseInt(oElem.getNumberValue());
					} else if (constraintValue === 'binary' && !bIgnoreIntConstraints) {
						bResult = oElem.getNumberValue() === 0 || oElem.getNumberValue() === 1;
					} else if (constraintValue === 'AllDifferent' && !bIgnoreIntConstraints && nIndex !== nStartRow) {
						bResult = oPrevConstraintCell.getNumberValue() !== oElem.getNumberValue();
					} else {
						switch (nOperator) {
							case c_oAscOperator['=']:
								let nDiff = constraintValue - oElem.getNumberValue();
								bResult = nDiff < nConstraintPrecision;
								break;
							case c_oAscOperator['>=']:
								bResult = oElem.getNumberValue() >= constraintValue - nConstraintPrecision;
								break;
							case c_oAscOperator['<=']:
								bResult = oElem.getNumberValue() <= constraintValue + nConstraintPrecision;
								break;
						}
					}
					oPrevConstraintCell = oElem;
					if (!bResult) {
						return true; // break loop
					}
				}
			});
			aConstraintsResult.push(bResult);
		}

		return aConstraintsResult;
	};
	/**
	 * Checks whether limits are exceeded from options.
	 * e.g., exceeds of maximum: iterations, time, subproblems, etc.
	 * @memberof CSolver
	 * @returns {boolean}
	 */
	CSolver.prototype.isFinishCalculating = function () {
		const nCurrentTime = Date.now();
		const nMaxIterations = this.getMaxIterations();
		const oOptions = this.getOptions();
		const nTimeMax = parseFloat(oOptions.getMaxTime());
		const nMaxSubproblems = parseFloat(oOptions.getMaxSubproblems());
		const nMaxFeasibleSolution = parseFloat(oOptions.getMaxFeasibleSolution());

		let bIsTimeMax = false;
		let bIterationIsReached = false;
		let bMaxSubproblems = false;
		let bMaxFeasibleSolution = false;

		if (!isNaN(nTimeMax)) {
			bIsTimeMax = nCurrentTime - this.getStartTime() >= nTimeMax * 1000;
		}
		if (!isNaN(nMaxIterations)) {
			bIterationIsReached = this.getCurrentAttempt() >= nMaxIterations;
		}
		if (!isNaN(nMaxSubproblems)) {
			bMaxSubproblems = this.getCurrentSubProblem() >= nMaxSubproblems;
		}
		if (!isNaN(nMaxFeasibleSolution)) {
			bMaxFeasibleSolution = this.getCurrentFeasibleCount() >= nMaxFeasibleSolution;
		}

		return bIsTimeMax || bIterationIsReached || bMaxSubproblems || bMaxFeasibleSolution;
	};
	/**
	 * Tries to find solution by GRG (Generalized reduced gradient) method.
	 * Uses for non-linear programming tasks.
	 * @memberof CSolver
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSolver.prototype.grgOptimization = function () {
		const oChangingCells = this.getChangingCell();
		const aConstraints = this.getConstraints();
		const nOptimizeResultTo = this.getOptimizeResultTo();
		const oOptions = this.getOptions();
		const nDerivatives = oOptions.getDerivatives();

		oChangingCells._foreachNoEmpty(function (oChangingCell) {

		});
	};
	/**
	 * Tries to find solution by Simplex method.
	 * Uses for LP tasks.
	 * @memberof CSolver
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSolver.prototype.simplexOptimization = function () {
		const oSimplexTableau = this.getSimplexTableau();
		const oOptions = this.getOptions();
		const bShowIterResults = oOptions.getShowIterResults();

		let bCompleteCalculation = oSimplexTableau.calculate();
		if (bShowIterResults && !bCompleteCalculation) {
			this.fillResult();
		}
		if (bCompleteCalculation && oSimplexTableau.getSolutionIsFound()) {
			this.fillResult();
		}

		return bCompleteCalculation;
	};
	/**
	 * Tries to find solution by Evolutionary method.
	 * Uses for non-smooth solver problems.
	 * @memberof CSolver
	 * @returns {boolean} The flag who recognizes end a loop of solver calculation. True - stop a loop, false - continue a loop.
	 */
	CSolver.prototype.evolutionOptimization = function () {

	};
	/**
	 * Resumes calculation by one step than pause it again.
	 * @memberof CSolver
	 */
	CSolver.prototype.step = function () {
		let oSolver = this;

		this.setIsPause(false);
		this.setIsSingleStep(true);
		this.setIntervalId(setInterval(function() {
			let bIsFinish = oSolver.calculate();
			if (bIsFinish) {
				clearInterval(oSolver.getIntervalId());
			}
		}, this.getDelay()));
	};
	/**
	 * Fills "Changing Variable Cells" values from result of calculation.
	 * @memberOf CSolver
	 */
	CSolver.prototype.fillResult = function () {
		const nSolutionMethod = this.getSolvingMethod();

		switch (nSolutionMethod) {
			case c_oAscSolvingMethod.grgNonlinear:
				break;
			case c_oAscSolvingMethod.simplexLP:
				const oSimplexTableau = this.getSimplexTableau();
				oSimplexTableau.updateVariableValues();
				break;
			case  c_oAscSolvingMethod.evolutionary:
				break;
		}
	};
	/**
	 * Returns value of "Optimize to" parameter
	 * @memberof CSolver
	 * @returns {c_oAscOptimizeTo}
	 */
	CSolver.prototype.getOptimizeResultTo = function () {
		return this.nOptimizeResultTo;
	};
	/**
	 * Returns converted to number type value of "Value of" input field from "To" parameter.
	 * @memberof CSolver
	 * @returns {null|number}
	 */
	CSolver.prototype.getValueOf = function () {
		return this.nValueOf;
	};
	/**
	 * Returns the array of constraints.
	 * @memberof CSolver
	 * @returns {CConstraint[]}
	 */
	CSolver.prototype.getConstraints = function () {
		return this.aConstraints;
	};
	/**
	 * Returns value of "Make Unconstrained Variables Non-Negative".
	 * @memberof CSolver
	 * @returns {boolean}
	 */
	CSolver.prototype.getVariablesNonNegative = function () {
		return this.bIsVarsNonNegative;
	};
	/**
	 * Returns the solving method.
	 * @memberof CSolver
	 * @returns {c_oAscSolvingMethod}
	 */
	CSolver.prototype.getSolvingMethod = function () {
		return this.nSolvingMethod;
	};
	/**
	 * Returns solver options.
	 * @memberof CSolver
	 * @returns {asc_COptions}
	 */
	CSolver.prototype.getOptions = function () {
		return this.oOptions;
	};
	/**
	 * Sets time from start calculation process.
	 * @memberof CSolver
	 * @param {number} nStartTime
	 */
	CSolver.prototype.setStartTime = function (nStartTime) {
		this.nStartTime = nStartTime;
	};
	/**
	 * Returns time from start calculation process.
	 * @returns {number}
	 */
	CSolver.prototype.getStartTime = function () {
		return this.nStartTime;
	};
	/**
	 * Sets current number of subproblem.
	 * @memberof CSolver
	 * @param {number} nCurrentSubProblem
	 */
	CSolver.prototype.setCurrentSubProblem = function (nCurrentSubProblem) {
		this.nCurrentSubProblem = nCurrentSubProblem;
	};
	/**
	 * Returns current number of subproblem.
	 * @memberof CSolver
	 * @returns {number}
	 */
	CSolver.prototype.getCurrentSubProblem = function () {
		return this.nCurrentSubProblem;
	};
	/**
	 * Increases number of subproblem.
	 * @memberof CSolver
	 */
	CSolver.prototype.increaseCurrentSubProblem = function () {
		this.nCurrentSubProblem++;
	};
	/**
	 * Return current count of feasible solution.
	 * @memberof CSolver
	 * @returns {number}
	 */
	CSolver.prototype.getCurrentFeasibleCount = function () {
		return this.nCurrentFeasibleCount;
	};
	/**
	 * Increases current count of feasible solution.
	 * @memberof CSolver
	 */
	CSolver.prototype.increaseCurrentFeasibleCount = function () {
		this.nCurrentFeasibleCount++;
	};
	/**
	 * Sets result of gradient for determine direction of motion for nonbasic variables.
	 * @memberof CSolver
	 * @param {number} nGradient
	 */
	CSolver.prototype.setGradient = function (nGradient) {
		this.nGradient = nGradient;
	};
	/**
	 * Returns result of gradient for determine direction of motion for nonbasic variables.
	 * @memberof CSolver
	 * @returns {number}
	 */
	CSolver.prototype.getGradient = function () {
		return this.nGradient;
	};
	/**
	 * Sets simplex tableau object. Using for calculating by simplex method.
	 * @memberof CSolver
	 * @param {CSimplexTableau} oSimplexTableau
	 */
	CSolver.prototype.setSimplexTableau = function (oSimplexTableau) {
		this.oSimplexTableau = oSimplexTableau;
	};
	/**
	 * Gets simplex tableau object. Using for calculating by simplex method.
	 * @memberof CSolver
	 * @returns {CSimplexTableau}
	 */
	CSolver.prototype.getSimplexTableau = function () {
		return this.oSimplexTableau;
	};
	/**
	 * Sets start range with original data
	 * @memberof CSolver
	 * @param {string} sCellName
	 * @param {string} sValue
	 */
	CSolver.prototype.setStartChangingCells = function (sCellName, sValue) {
		if (this.oStartChangingCells == null) {
			this.oStartChangingCells = {};
		}
		this.oStartChangingCells[sCellName] = sValue;
	};
	/**
	 * Returns start range with original data.
	 * @memberof CSolver
	 * @returns {{cellName:string, value:string}}
	 */
	CSolver.prototype.getStartChangingCells = function () {
		return this.oStartChangingCells;
	};
	/**
	 * Returns error type
	 * @memberof CSolver
	 * @param {AscCommonExcel.Worksheet} ws - checked sheet.
	 * @param {Asc.Range} range - checked range.
	 * @param {Asc.c_oAscSelectionDialogType} type - dialog type.
	 * @returns {Asc.c_oAscError}
	 */
	CSolver.prototype.isValidDataRef = function(ws, range, type) {
		let res = Asc.c_oAscError.ID.No;
		if (range && !range.isOneCell()) {
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
		}

		return res;
	};


	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CGoalSeek = CGoalSeek;
	window['AscCommonExcel'].CSolver = CSolver;
	window['AscCommonExcel'].actualWsByRef = actualWsByRef;
	window['AscCommonExcel'].convertToAbsoluteRef = convertToAbsoluteRef;

	// Collections and classes for UI part
	window['AscCommonExcel'].c_oAscDerivativeType = c_oAscDerivativeType;
	window['AscCommonExcel'].c_oAscOperator = c_oAscOperator;
	window['AscCommonExcel'].c_oAscOptimizeTo = c_oAscOptimizeTo;
	window['AscCommonExcel'].c_oAscSolvingMethod = c_oAscSolvingMethod;
	window['AscCommonExcel'].c_oAscSolverResult = c_oAscSolverResult;
	window['AscCommonExcel'].c_oResultStatus = c_oResultStatus;

	window['AscCommonExcel'].asc_CSolverParams = asc_CSolverParams;
	window['AscCommonExcel'].asc_CSolverResults = asc_CSolverResults;

})(window);
