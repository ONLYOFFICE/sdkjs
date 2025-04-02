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
		this.nOptimizeResultTo = null;
		this.sValueOf = null;
		this.aConstraints = new Map();
		this.bVariablesNonNegative = false;
		this.nSolvingMethod = null;

		this.oOptions = new asc_COptions();

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
	 * @param {string} objectiveFunction
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
	 * @param {string} changingCells
	 */
	asc_CSolverParams.prototype.setChangingCells = function (changingCells) {
		this.sChangingCells = changingCells;
	}
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
	}
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
	 * @param {number} index - key of constraint in Map object
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
	 * Class representing options for calculating.
	 * @constructor
	 */
	function asc_COptions () {
		// All methods
		this.sConstraintPrecision = null;
		this.bAutomaticScaling = false;
		this.bShowIterResults = false;
		this.bIgnoreIntConstriants = false;
		this.sIntOptimal = null;
		this.sMaxTime = null;
		this.sIterations = null;
		this.sMaxSubproblems = null;
		this.sMaxFeasibleSolution = null;
		// GRG Nonlinear and Evolutionary
		this.sConvergence = null;
		this.nDerivatives = null;
		this.bMultistart = false; // ?
		this.sPopulationSize = null;
		this.sRandomSeed = null;
		this.bRequireBounds = false;
		this.sMutationRate = null;
		this.sEvoMaxTime = null;
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
	}
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
	}
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

	// Classes and functions for main logic of Solver feature
	/**
	 * Returns worksheet that has selected cell(s) from parameters.
	 * @param {string} sRef - selected cell(s) reference from parameters
	 * @param {Worksheet} oWs
	 * @returns {Worksheet}
	 * @private
	 */
	function _actualWsByRef (sRef, oWs) {
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
	 * @private
	 */
	function _convertToAbsoluteRef (sRef) {
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

		this.nCostRowIndex = 0;
		this.nRhsColumn = 0;

		this.aVariablesPerIndex = []; // ?
		this.oUnrestrictedVars = null;

		// Solution attributes
		this.bFeasible = true; // until proven guilty
		this.nEvaluation = 0;
		this.nSimplexIters = 0; // ?

		this.aVarIndexByRow = null;
		this.aVarIndexByCol = null;

		this.aRowByVarIndex = null;
		this.aColByVarIndex = null;

		this.nPrecision = 1e-8;

		this.aOptionalObjectives = []; // ?
		this.oObjectivesByPriority = {}; // ?

		this.savedState = null;

		this.aAvailableIndexes = [];
		this.nLastElementIndex = 0;

		this.oVariables = null;
		this.nVars = 0;

		this.bBounded = true;
		this.unboundedVarIndex = null;

		this.nBranchAndCutIters = 0;
	}

	/**
	 * Initializes a start data for calculating logic.
	 * @memberof CSimplexTableau
	 */
	CSimplexTableau.prototype.init = function () {
		const oModel = this.getModel();
		const oChangingCells = oModel.getChangingCell();
		const oBboxChangingCells = oChangingCells.bbox;
		const nTotalCountVariables = oBboxChangingCells.getWidth() * oBboxChangingCells.getHeight();
		const aConstraints = oModel.getConstraints();
		const aSimplexConstraints = this.getSimplexConstraints(aConstraints);
		const nTotalCountConstraints = aSimplexConstraints.length;

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
	};
	/**
	 * @returns {CSolver}
	 */
	CSimplexTableau.prototype.getModel = function () {
		return this.oModel;
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
		const oVarIndexByCellName = {};
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
	}
	/**
	 * Returns available indexes.
	 * @memberof CSimplexTableau
	 * @returns {[]}
	 */
	CSimplexTableau.prototype.getAvailableIndexes = function () {
		return this.aAvailableIndexes;
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

	// Main class of solver feature
	/**
	 * Class representing a solver feature.
	 * @param {asc_CSolverParams} oParams - solver parameters
	 * @param {Worksheet} oWs
	 * @constructor
	 */
	function CSolver (oParams, oWs) {
		const oParsedFormula = this.convertToCell(oParams.getObjectiveFunction(), oWs).getFormulaParsed();
		const oChangingCellsWs = _actualWsByRef(oParams.getChangingCells(), oWs);
		const sChangingCells = _convertToAbsoluteRef(oParams.getChangingCells());
		CBaseAnalysis.call(this, oParsedFormula, oChangingCellsWs.getRange2(sChangingCells));

		// Solver parameters
		this.nOptimizeResultTo = oParams.getOptimizeResultTo();
		this.nValueOf = oParams.getValueOf() !== null ? parseFloat(oParams.getValueOf().replace(/,/g, ".")) : null;
		this.aConstraints = this.initConstraints(oParams.getConstraints(), oWs);
		this.bIsVarsNonNegative = oParams.getVariablesNonNegative();
		this.nSolvingMethod = oParams.getSolvingMethod();

		// Attributes for calculating logic
		this.nStartTime = null;
		this.nCurrentSubProblem = null;
		this.nGradient = null;
		this.oSimplexTableau = null;

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
		const aConstraints = this.getConstraints();
		const nSolutionMethod = this.getSolvingMethod();
		const oChangingCells = this.getChangingCell();

		// Fills empty cells to 0 value for changing cells
		oChangingCells._foreachNoEmpty(function (oCell) {
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
		const aConstraintsResult = this.calculateConstraints();
		const nSolutionMethod = this.getSolvingMethod();

		if (this.isFinishCalculating(aConstraintsResult)) {
			return true;
		}
		switch (nSolutionMethod) {
			case c_oAscSolvingMethod.grgNonlinear:
				this.grgOptimization();
				break;
			case c_oAscSolvingMethod.simplexLP:
				this.simplexOptimization();
				break;
			case c_oAscSolvingMethod.evolutionary:
				this.evolutionOptimization();
				break;
		}

		return false;
	};
	/**
	 * Converts cell reference from UI to Cell object.
	 * @memberof CSolver
	 * @param {string} sCellRef
	 * @param {Worksheet}oWs
	 * @returns {Cell}
	 */
	CSolver.prototype.convertToCell = function (sCellRef, oWs) {
		const oCellWs = _actualWsByRef(sCellRef, oWs);
		const sCellActualRef = _convertToAbsoluteRef(sCellRef);
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

		oConstraints.forEach((oConstraint) => {
			const oConstraintsWs = _actualWsByRef(oConstraint.constraint, oWs);
			const sConstraintsActualRef = _convertToAbsoluteRef(oConstraint.constraint);
			const oCellRefWs = _actualWsByRef(oConstraint.cellRef, oWs);
			const sCellRefActualRef = _convertToAbsoluteRef(oConstraint.cellRef);
			let constraintData = oConstraintsWs.getRange2(sConstraintsActualRef);
			if (constraintData == null) {
				constraintData = aExcludeWords.includes(oConstraint.constraint) ? oConstraint.constraint :
					Number(oConstraint.constraint.replace(/,/g, "."));
			}
			aConstraints.push(new CConstraint(oCellRefWs.getRange2(sCellRefActualRef), oConstraint.operator, constraintData));
		});

		return aConstraints;
	};
	/**
	 * Calculates and returns an array of constraint results.
	 * @memberof CSolver
	 * @returns {boolean[]}
	 */
	CSolver.prototype.calculateConstraints = function () { // todo Need to rework.
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
	 * Checks whether finish calculating.
	 * @memberof CSolver
	 * @param {boolean[]} aConstraintsResult
	 * @returns {boolean}
	 */
	CSolver.prototype.isFinishCalculating = function (aConstraintsResult) { //todo: Need to rework
		const nCurrentTime = Date.now();
		const bConstraintsIsSatisfied = !aConstraintsResult.includes(false);
		const nMaxIterations = this.getMaxIterations();
		const oOptions = this.getOptions();
		const nTimeMax = parseFloat(oOptions.getMaxTime());
		const nConvergence = parseFloat(oOptions.getConvergence().replace(/,/g, "."));
		const nValue = this.getValueOf();
		const nMaxSubproblems = parseFloat(oOptions.getMaxSubproblems());
		const nPrevFactValue = this.getPrevFactValue();

		let nFactValue = this.calculateFormula();
		let nDiff = nFactValue - nPrevFactValue;
		let bIsSatisfied = bConstraintsIsSatisfied && (!!nPrevFactValue && nDiff < nConvergence);
		let bIsTimeMax = true;
		if (!isNaN(nTimeMax)) {
			bIsTimeMax = nCurrentTime - this.getStartTime() < nTimeMax * 1000;
		}
		let bIterationIsReached = true;
		if (!isNaN(nMaxIterations)) {
			bIterationIsReached = this.getCurrentAttempt() >= nMaxIterations;
		}
		let bMaxSubproblems = true;
		if (!isNaN(nMaxSubproblems)) {
			bMaxSubproblems = this.getCurrentSubProblem() >= nMaxSubproblems;
		}

		return bIsSatisfied && bIsTimeMax && bIterationIsReached && bMaxSubproblems;
	};
	CSolver.prototype.grgOptimization = function () {
		const oChangingCells = this.getChangingCell();
		const aConstraints = this.getConstraints();
		const nOptimizeResultTo = this.getOptimizeResultTo();
		const oOptions = this.getOptions();
		const nDerivatives = oOptions.getDerivatives();

		oChangingCells._foreachNoEmpty(function (oChangingCell) {

		});


	};
	CSolver.prototype.simplexOptimization = function () {

	};
	CSolver.prototype.evolutionOptimization = function () {

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
	}
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
	}

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CGoalSeek = CGoalSeek;
	window['AscCommonExcel'].CSolver = CSolver;

	window['AscCommonExcel'].c_oAscDerivativeType = c_oAscDerivativeType;
	window['AscCommonExcel'].c_oAscOperator = c_oAscOperator;
	window['AscCommonExcel'].c_oAscOptimizeTo = c_oAscOptimizeTo;
	window['AscCommonExcel'].c_oAscSolvingMethod = c_oAscSolvingMethod;

	window['AscCommonExcel'].asc_CSolverParams = asc_CSolverParams;

})(window);
