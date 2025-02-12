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

	//Collections for UI
	/** @enum {number} */
	const c_oAscOptimizeTo = {
		max: 0,
		min: 1,
		valueOf: 2
	};
	/** @enum {number} */
	const c_oAscSolvingMethod = {
		grgNonlinear: 0,
		simplexLP: 1,
		evolutionary: 2
	};
	/** @enum {number} */
	const c_oAscDerivativeType = {
		forward: 0,
		central: 1
	};
	/** @enum {number} */
	const c_oAscOperator = {
		'>=': 0,
		'=': 1,
		'<=': 2,
		'int': 3,
		'bin': 4,
		'dif': 5
	};

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
		// todo add value from settings
		this.nMaxIterations = 100; // max iterations of goal seek. Default value is 100
	}

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
		this.nCurAttempt = 0;
		this.nChangingVal = null;
		this.nPrevValue = null;
		this.nPrevFactValue = null;
		this.bReverseCompare = null;
		this.bEnabledRidder = false;
		this.nLow = null;
		this.nHigh = null;
		this.bIsSingleStep = false;

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

	// Solver
	// Classes for interact with UI.
	/**
	 * @constructor
	 * @returns {asc_CSolverParams}
	 */
	function asc_CSolverParams () {
		this.sObjectiveFunction = null;
		this.sChangingCells = null;
		this.nOptimizeResultTo = null;
		this.aConstraints = new Map();
		this.bVariablesNonNegative = false;
		this.nSolvingMethod = null;

		this.oOptions = new asc_COptions();

		return this;
	}


	/**
	 * Returns the objective function.
	 * @memberof asc_CSolverParams
	 * @returns {string}
	 */
	asc_CSolverParams.prototype.getObjectiveFunction = function () {
		return this.sObjectiveFunction;
	};
	/**
	 * Sets the objective function.
	 * @memberof asc_CSolverParams
	 * @param {string} objectiveFunction
	 */
	asc_CSolverParams.prototype.setObjectiveFunction = function (objectiveFunction) {
		this.sObjectiveFunction = objectiveFunction;
	};
	/**
	 * Returns "Optimize result to" parameter.
	 * @memberof asc_CSolverParams
	 * @returns {c_oAscOptimizeTo}
	 */
	asc_CSolverParams.prototype.getOptimizeResultTo = function () {
		return this.nOptimizeResultTo;
	};

	/**
	 * Sets "Optimize result to" parameter.
	 * @memberof asc_CSolverParams
	 * @param {number} optimizeResultTo
	 */
	asc_CSolverParams.prototype.setOptimizeResultTo = function (optimizeResultTo) {
		this.nOptimizeResultTo = optimizeResultTo;
	};
	/**
	 * Returns the constraints.
	 * @memberof asc_CSolverParams
	 * @returns {Map}
	 */
	asc_CSolverParams.prototype.getConstraints = function () {
		return this.aConstraints;
	};
	/**
	 * Sets the constraints.
	 * @memberof asc_CSolverParams
	 * @param {Map} constraints
	 */
	asc_CSolverParams.prototype.setConstraints = function (constraints) {
		this.aConstraints = constraints;
	};
	/**
	 * Returns whether unconstrained variables are non-negative.
	 * @memberof asc_CSolverParams
	 * @returns {boolean}
	 */
	asc_CSolverParams.prototype.getVariablesNonNegative = function () {
		return this.bVariablesNonNegative;
	};
	/**
	 * Sets whether unconstrained variables are non-negative.
	 * @memberof asc_CSolverParams
	 * @param {boolean} variablesNonNegative
	 */
	asc_CSolverParams.prototype.setVariablesNonNegative = function (variablesNonNegative) {
		this.bVariablesNonNegative = variablesNonNegative;
	};
	/**
	 * Returns the solving method.
	 * @memberof asc_CSolverParams
	 * @returns {c_oAscSolvingMethod}
	 */
	asc_CSolverParams.prototype.getSolvingMethod = function () {
		return this.nSolvingMethod;
	};
	/**
	 * Sets the solving method.
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
	 * Class representing options for calculating.
	 * @constructor
	 */
	function asc_COptions () {
		// All methods
		this.nConstraintPrecision = null;
		this.bAutomaticScaling = false;
		this.bShowIterResults = false;
		this.bIgnoreIntConstriants = false;
		this.nIntOptimal = null;
		this.nMaxTime = null;
		this.nMaxSubproblems = null;
		this.nMaxFeasibleSolution = null;
		// GRG Nonlinear and Evolutionary
		this.nConvergence = null;
		this.nDerivatives = null;
		this.bMultistart = false; // ?
		this.nPopulationSize = null;
		this.nRandomSeed = null;
		this.bRequireBounds = false;
		this.nMutationRate = null;
		this.nEvoMaxTime = null;
	}

	/**
	 * Returns the constraint precision.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getConstraintPrecision = function() {
		return this.nConstraintPrecision;
	};
	/**
	 * Returns the automatic scaling option.
	 * @memberof asc_COptions
	 * @returns {boolean} True if automatic scaling is enabled
	 */
	asc_COptions.prototype.getAutomaticScaling = function() {
		return this.bAutomaticScaling;
	};
	/**
	 * Returns whether to show iteration results.
	 * @memberof asc_COptions
	 * @returns {boolean} True if iteration results should be shown.
	 */
	asc_COptions.prototype.getShowIterResults = function() {
		return this.bShowIterResults;
	};
	/**
	 * Returns whether to ignore integer constraints.
	 * @memberof asc_COptions
	 * @returns {boolean} True if integer constraints are ignored
	 */
	asc_COptions.prototype.getIgnoreIntConstraints = function() {
		return this.bIgnoreIntConstriants;
	};
	/**
	 * Returns the integer optimality value.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getIntOptimal = function () {
		return this.nIntOptimal;
	};
	/**
	 * Returns the maximum calculation time.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getMaxTime = function () {
		return this.nMaxTime;
	};
	/**
	 * Returns the maximum subproblems value.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getMaxSubproblems = function () {
		return this.nMaxSubproblems;
	};
	/**
	 * Returns the maximum feasible solution count.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getMaxFeasibleSolution = function () {
		return this.nMaxFeasibleSolution;
	};
	/**
	 * Returns the convergence value.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getConvergence = function () {
		return this.nConvergence;
	};
	/**
	 * Returns the derivatives value.
	 * @memberof asc_COptions
	 * @returns {number} The derivatives value.
	 */
	asc_COptions.prototype.getDerivatives = function () {
		return this.nDerivatives;
	};
	/**
	 * Returns whether multistart is enabled.
	 * @memberof asc_COptions
	 * @returns {boolean} True if multistart is enabled
	 */
	asc_COptions.prototype.getMultistart = function () {
		return this.bMultistart;
	};
	/**
	 * Returns the population size for evolutionary calculations.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getPopulationSize = function () {
		return this.nPopulationSize;
	};
	/**
	 * Returns the random seed for calculations.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getRandomSeed = function () {
		return this.nRandomSeed;
	};
	/**
	 * Returns require bounds on Variables option.
	 * @memberof asc_COptions
	 * @returns {boolean} True if bounds are required.
	 */
	asc_COptions.prototype.getRequireBounds = function () {
		return this.bRequireBounds;
	};
	/**
	 * Returns the mutation rate for evolutionary calculations.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getMutationRate = function () {
		return this.nMutationRate;
	};
	/**
	 * Returns value of "Maximum Time without improvement" option.
	 * @memberof asc_COptions
	 * @returns {number}
	 */
	asc_COptions.prototype.getEvoMaxTime = function () {
		return this.nEvoMaxTime;
	};
	/**
	 * Sets the constraint precision.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setConstraintPrecision = function (value) {
		this.nConstraintPrecision = value;
	};
	/**
	 * Enables or disables automatic scaling.
	 * @memberof asc_COptions
	 * @param {boolean} value
	 */
	asc_COptions.prototype.setAutomaticScaling = function (value) {
		this.bAutomaticScaling = value;
	};
	/**
	 * Sets whether to show iteration results.
	 * @memberof asc_COptions
	 * @param {boolean} value
	 */
	asc_COptions.prototype.setShowIterResults = function (value) {
		this.bShowIterResults = value;
	};
	/**
	 * Sets whether to ignore integer constraints.
	 * @memberof asc_COptions
	 * @param {boolean} value
	 */
	asc_COptions.prototype.setIgnoreIntConstraints = function (value) {
		this.bIgnoreIntConstriants = value;
	};
	/**
	 * Sets the integer optimality value.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setIntOptimal = function (value) {
		this.nIntOptimal = value;
	};
	/**
	 * Sets the maximum calculation time.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setMaxTime = function (value) {
		this.nMaxTime = value;
	};
	/**
	 * Sets the maximum subproblems value.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setMaxSubproblems = function (value) {
		this.nMaxSubproblems = value;
	};
	/**
	 * Sets the maximum feasible solution count.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setMaxFeasibleSolution = function (value) {
		this.nMaxFeasibleSolution = value;
	};
	/**
	 * Sets the convergence value.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setConvergence = function (value) {
		this.nConvergence = value;
	};
	/**
	 * Sets the derivatives value.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setDerivatives = function (value) {
		this.nDerivatives = value;
	};
	/**
	 * Enables or disables multistart.
	 * @memberof asc_COptions
	 * @param {boolean} value
	 */
	asc_COptions.prototype.setMultistart = function (value) {
		this.bMultistart = value;
	};
	/**
	 * Sets the population size for evolutionary calculations.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setPopulationSize = function (value) {
		this.nPopulationSize = value;
	};
	/**
	 * Sets the random seed for calculations.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setRandomSeed = function (value) {
		this.nRandomSeed = value;
	};
	/**
	 * Sets whether bounds are required on Variables option.
	 * @memberof asc_COptions
	 * @param {boolean} value
	 */
	asc_COptions.prototype.setRequireBounds = function (value) {
		this.bRequireBounds = value;
	};
	/**
	 * Sets the mutation rate for evolutionary calculations.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setMutationRate = function (value) {
		this.nMutationRate = value;
	};
	/**
	 * Sets the "Maximum Time without improvement" option value.
	 * @memberof asc_COptions
	 * @param {number} value
	 */
	asc_COptions.prototype.setEvoMaxTime = function (value) {
		this.nEvoMaxTime = value;
	};

	// Classes for main logic of Solver feature
	/**
	 * Class representing a constraints option for Solver.
	 * @param {Range} oCell
	 * @param {c_oAscOperator} nOperator
	 * @param {Range|number} constraint
	 * @constructor
	 */
	function CConstraints (oCell, nOperator, constraint) {
		this.oCell = oCell;
		this.nOperator = nOperator;
		this.constraint = constraint;
	}
	
	/**
	 * Returns cell reference.
	 * @memberof CConstraints
	 * @returns {Range}
	 */
	CConstraints.prototype.getCell = function() {
		return this.oCell;
	};
	/**
	 * Returns comparison operator.
	 * @memberof CConstraints
	 * @returns {number}
	 */
	CConstraints.prototype.getOperator = function() {
		return this.nOperator;
	};
	/**
	 * Returns constraint. Element that comparisons with reference cell.
	 * @memberof CConstraints
	 * @returns {Range|number}
	 */
	CConstraints.prototype.getConstraint = function() {
		return this.constraint;
	};

	// Main class of solver feature
	/**
	 * Class representing a solver feature
	 * @param  {asc_CSolverParams} oParams - solver parameters
	 * @constructor
	 */
	function CSolver (oParams) {
		//TODO logic for converting string from attributes: sObjective and sChangingCells to cell object.
		CBaseAnalysis.call(this, oParams.parsedFormula, oParams.changingCells);

		//Solver parameters
		this.nOptimizeResultTo = oParams.getOptimizeResultTo();
		this.aConstraints = oParams.getConstraints();
		this.bIsVarsNonNegative = oParams.getVariablesNonNegative();
		this.nSolvingMethod = oParams.getSolvingMethod();
		// Calculating option
		this.oOptions = oParams.getOptions();
	}

	CSolver.prototype = Object.create(CBaseAnalysis.prototype);
	CSolver.prototype.constructor = CSolver;

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CGoalSeek = CGoalSeek;
	window['AscCommonExcel'].CSolver = CSolver;

})(window);
