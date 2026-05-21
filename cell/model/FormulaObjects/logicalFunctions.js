/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	var cErrorType = AscCommonExcel.cErrorType;
	var cNumber = AscCommonExcel.cNumber;
	var cString = AscCommonExcel.cString;
	var cBool = AscCommonExcel.cBool;
	var cError = AscCommonExcel.cError;
	var cArea = AscCommonExcel.cArea;
	var cArea3D = AscCommonExcel.cArea3D;
	var cEmpty = AscCommonExcel.cEmpty;
	var cArray = AscCommonExcel.cArray;
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var cElementType = AscCommonExcel.cElementType;
	var argType = Asc.c_oAscFormulaArgumentType;

	cFormulaFunctionGroup['Logical'] = cFormulaFunctionGroup['Logical'] || [];
	cFormulaFunctionGroup['Logical'].push(cAND, cFALSE, cIF, cIFERROR, cIFNA, cIFS, cNOT, cOR, cSWITCH, cTRUE, cXOR);

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cAND() {
	}

	//***array-formula***
	cAND.prototype = Object.create(cBaseFunction.prototype);
	cAND.prototype.constructor = cAND;
	cAND.prototype.name = 'AND';
	cAND.prototype.argumentsMin = 1;
	cAND.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cAND.prototype.argumentsType = [[argType.logical]];
	cAND.prototype.enabledToSingle = {"*": true};
	cAND.prototype.Calculate = function (arg) {
		let argResult = null;
		for (let i = 0; i < arg.length; i++) {
			if (arg[i].type === cElementType.cellsRange || arg[i].type === cElementType.cellsRange3D) {
				let argArr = arg[i].getValue();
				for (var j = 0; j < argArr.length; j++) {
					if (argArr[j] instanceof cError) {
						return argArr[j];
					} else if (!(argArr[j] instanceof cString || argArr[j] instanceof cEmpty)) {
						if (argResult === null) {
							argResult = argArr[j].tocBool();
						} else {
							argResult = new cBool(argResult.value && argArr[j].tocBool().value);
						}
						if (argResult.value === false) {
							return new cBool(false);
						}
					}
				}
			} else {
				if (arg[i].type === cElementType.cell || arg[i].type === cElementType.cell3D) {
					arg[i] = arg[i].getValue();
				}
				if (arg[i].type === cElementType.string) {
					// skip string check
					continue
				} else if (arg[i].type === cElementType.error) {
					return arg[i];
				} else if (arg[i].type === cElementType.array) {
					arg[i].foreach(function (elem) {
						if (elem && elem.type === cElementType.error) {
							argResult = elem;
							return true;
						} else if (elem.type === cElementType.string || elem.type === cElementType.empty) {
							return false;
						} else {
							if (argResult === null) {
								argResult = elem.tocBool();
							} else {
								argResult = new cBool(argResult.value && elem.tocBool().value);
							}
							if (argResult.value === false) {
								return true;
							}
						}
					});
				} else {
					if (argResult === null) {
						argResult = arg[i].tocBool();
					} else {
						argResult = new cBool(argResult.value && arg[i].tocBool().value);
					}
					if (argResult.value === false) {
						return new cBool(false);
					}
				}
			}
		}
		if (argResult === null) {
			return new cError(cErrorType.wrong_value_type);
		}
		return argResult;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFALSE() {
	}

	//***array-formula***
	cFALSE.prototype = Object.create(cBaseFunction.prototype);
	cFALSE.prototype.constructor = cFALSE;
	cFALSE.prototype.name = 'FALSE';
	cFALSE.prototype.argumentsMax = 0;
	cFALSE.prototype.argumentsType = null;
	cFALSE.prototype.Calculate = function () {
		return new cBool(false);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cIF() {
	}

	//***array-formula***
	cIF.prototype = Object.create(cBaseFunction.prototype);
	cIF.prototype.constructor = cIF;
	cIF.prototype.name = 'IF';
	cIF.prototype.argumentsMin = 2;
	cIF.prototype.argumentsMax = 3;
	cIF.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cIF.prototype.argumentsType = [argType.logical, argType.any, argType.any];
	cIF.prototype.arrayIndexes = {1: 1, 2: 1};
	cIF.prototype.getArrayIndex = function (index, type, args) {
		let res = false;

		//TODO need recheck
		if (args && args[0].type === cElementType.array) {
			return false;
		}

		if (type === cElementType.array) {
			return false;
		}
		if (this.arrayIndexes) {
			res = this.arrayIndexes[index];
		}
		return res;
	};
	cIF.prototype.enabledToSingle = {"1": true, "2": true};
	cIF.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2];

		if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0);
		}
		if (arg1 instanceof cArray) {
			arg1 = arg1.getElement(0);
		}
		if (arg2 instanceof cArray) {
			arg2 = arg2.getElement(0);
		}

		arg0 = arg0.tocBool();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		} else if (arg0.value) {
			return arg1 ? arg1 instanceof cEmpty ? new cNumber(0) : arg1 : new cBool(true);
		} else {
			return arg2 ? arg2 instanceof cEmpty ? new cNumber(0) : arg2 : new cBool(false);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cIFERROR() {
	}

	//***array-formula***
	cIFERROR.prototype = Object.create(cBaseFunction.prototype);
	cIFERROR.prototype.constructor = cIFERROR;
	cIFERROR.prototype.name = 'IFERROR';
	cIFERROR.prototype.argumentsMin = 2;
	cIFERROR.prototype.argumentsMax = 2;
	cIFERROR.prototype.argumentsType = [argType.any, argType.any];
	cIFERROR.prototype.enabledToSingle = {"1": true};
	cIFERROR.prototype.Calculate = function (arg) {
		let arg0 = arg[0];
		if (arg0.type === cElementType.array) {
			arg0 = arg0.getElement(0);
		}
		if (arg0.type === cElementType.cell || arg0.type === cElementType.cell3D) {
			arg0 = arg0.getValue();
		}
		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]);
		}

		if (arg0.type === cElementType.error) {
			return arg[1].type === cElementType.array ? arg[1].getElement(0) : arg[1];
		} else {
			return arg[0];
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cIFNA() {
	}

	//***array-formula***
	cIFNA.prototype = Object.create(cBaseFunction.prototype);
	cIFNA.prototype.constructor = cIFNA;
	cIFNA.prototype.name = 'IFNA';
	cIFNA.prototype.argumentsMin = 2;
	cIFNA.prototype.argumentsMax = 2;
	cIFNA.prototype.isXLFN = true;
	cIFNA.prototype.argumentsType = [argType.any, argType.any];
	cIFNA.prototype.enabledToSingle = {"1": true};
	cIFNA.prototype.Calculate = function (arg) {

		let arg0 = arg[0];
		if (arg0.type === cElementType.array) {
			arg0 = arg0.getElement(0);
		}
		if (arg0.type === cElementType.cell || arg0.type === cElementType.cell3D) {
			arg0 = arg0.getValue();
		}
		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]);
		}

		if (arg0.type === cElementType.error && cErrorType.not_available === arg0.errorType) {
			return arg[1].type === cElementType.array ? arg[1].getElement(0) : arg[1];
		} else {
			return arg[0];
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cIFS() {
	}

	//***array-formula***
	cIFS.prototype = Object.create(cBaseFunction.prototype);
	cIFS.prototype.constructor = cIFS;
	cIFS.prototype.name = 'IFS';
	cIFS.prototype.argumentsMin = 2;
	cIFS.prototype.isXLFN = true;
	cIFS.prototype.argumentsType = [[argType.logical, argType.any]];
	cIFS.prototype.enabledToSingle = {"odd": true};
	cIFS.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1], true);
		var argClone = oArguments.args;

		var res = null;
		for (var i = 0; i < arg.length; i++) {
			var argN = argClone[i];
			if (cElementType.error === argN.type) {
				res = argN;
				break;
			} else if (cElementType.string === argN.type) {
				res = new cError(cErrorType.wrong_value_type);
				break;
			} else if (cElementType.number === argN.type || cElementType.bool === argN.type) {
				if (!argClone[i + 1]) {
					res = new cError(cErrorType.not_available);
					break;
				}

				argN = argN.tocBool();
				if (true === argN.value) {
					res = argClone[i + 1];
					break;
				}
			}
			if (i === arg.length - 1) {
				res = new cError(cErrorType.not_available);
				break;
			}
			i++;
		}

		if (null === res) {
			return new cError(cErrorType.not_available);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cNOT() {
	}

	//***array-formula***
	cNOT.prototype = Object.create(cBaseFunction.prototype);
	cNOT.prototype.constructor = cNOT;
	cNOT.prototype.name = 'NOT';
	cNOT.prototype.argumentsMin = 1;
	cNOT.prototype.argumentsMax = 1;
	cNOT.prototype.argumentsType = [argType.logical];
	cNOT.prototype.Calculate = function (arg) {

		let arg0 = arg[0];
		if (arg0.type === cElementType.array) {
			arg0 = arg0.getElement(0);
		}
		if (arg0.type === cElementType.cell || arg0.type === cElementType.cell3D) {
			arg0 = arg0.getValue();
		}
		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]);
		}

		if (arg0.type === cElementType.string) {
			let res = arg0.tocBool();
			if (res.type === cElementType.string) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				return new cBool(!res.value);
			}
		} else if (arg0.type === cElementType.error) {
			return arg0;
		} else {
			return new cBool(!arg0.tocBool().value);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cOR() {
	}

	//***array-formula***
	cOR.prototype = Object.create(cBaseFunction.prototype);
	cOR.prototype.constructor = cOR;
	cOR.prototype.name = 'OR';
	cOR.prototype.argumentsMin = 1;
	cOR.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cOR.prototype.argumentsType = [[argType.logical]];
	cOR.prototype.enabledToSingle = {"*": true};
	cOR.prototype.Calculate = function (arg) {
		let argResult = null;
		for (let i = 0; i < arg.length; i++) {
			let argI;
			if (arg[i].type === cElementType.cell || arg[i].type === cElementType.cell3D) {
				argI = arg[i].getValue();
			}

			if (arg[i].type === cElementType.error) {
				return arg[i];
			}

			if (arg[i].type === cElementType.cellsRange || arg[i].type === cElementType.cellsRange3D || arg[i].type === cElementType.array) {
				let dimensions = arg[i].getDimensions();

				for (let row = 0; row < dimensions.row; row++) {
					for (let col = 0; col < dimensions.col; col++) {
						let elemVal = arg[i].getValueByRowCol ? arg[i].getValueByRowCol(row,col,true) : arg[i].getElementRowCol(row,col);
						
						if (elemVal.type === cElementType.error) {
							return elemVal;
						}

						let argToBool = elemVal.tocBool();
						if (argToBool.type === cElementType.string) {
							continue;
						} if (argToBool.type === cElementType.bool) {
							argResult = argToBool;
						} else if (argToBool.type === cElementType.error) {
							return argToBool;
						}

						if (argResult && argResult.value === true) {
							return new cBool(true);
						}
					}
				}

			} else if (arg[i].type === cElementType.string) {
				let stringToBool = arg[i].tocBool();
				if (stringToBool.type === cElementType.bool) {
					argResult = stringToBool;
				}
			} else {
				if (argResult === null) {
					argResult = arg[i].tocBool();
				} else {
					argResult = new cBool(argResult.value || arg[i].tocBool().value);
				}
			}

			if (argResult && argResult.value === true) {
				return new cBool(true);
			}
		}
		if (argResult === null) {
			return new cError(cErrorType.wrong_value_type);
		}
		return argResult;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSWITCH() {
	}

	//***array-formula***
	cSWITCH.prototype = Object.create(cBaseFunction.prototype);
	cSWITCH.prototype.constructor = cSWITCH;
	cSWITCH.prototype.name = 'SWITCH';
	cSWITCH.prototype.argumentsMin = 3;
	cSWITCH.prototype.argumentsMax = 126;
	cSWITCH.prototype.isXLFN = true;
	cSWITCH.prototype.argumentsType = [argType.any, argType.any, argType.any, [argType.any, argType.any]];
	cSWITCH.prototype.enabledToSingle = {"evenFrom2": true};
	cSWITCH.prototype.Calculate = function (arg) {
		let oArguments = this._prepareArguments(arg, arguments[1], true);
		let argClone = oArguments.args;

		let argError;
		let argCloneOdd = argClone.filter(function(item, index) {
			return index === 0 || index % 2 !== 0;
		});

		// check all odd arguments for errors plus the starting one(first arg)
		if (argError = this._checkErrorArg(argCloneOdd)) {
			return argError;
		}

		let arg0 = argClone[0];
		if (cElementType.cell === argClone[0].type || cElementType.cell3D === argClone[0].type) {
			arg0 = arg0.getValue()
		}
		let arg0Type = arg0.type;

		let res = null;
		for (let i = 1; i < argClone.length; i++) {
			let argN = argClone[i];
			if (cElementType.cell === argN.type || cElementType.cell3D === argN.type) {
				argN = argN.getValue();
			}
			let argNType = argN.type;

			if (arg0Type === argNType && arg0.getValue() === argN.getValue()) {
				if (!argClone[i + 1]) {
					return new cError(cErrorType.not_available);
				} else {
					res = argClone[i + 1];
					break;
				}
			}
			if (i === argClone.length - 1) {
				res = argClone[i];
			}
			i++;
		}

		if (null === res) {
			return new cError(cErrorType.not_available);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTRUE() {
	}

	//***array-formula***
	cTRUE.prototype = Object.create(cBaseFunction.prototype);
	cTRUE.prototype.constructor = cTRUE;
	cTRUE.prototype.name = 'TRUE';
	cTRUE.prototype.argumentsMax = 0;
	cTRUE.prototype.argumentsType = null;
	cTRUE.prototype.Calculate = function () {
		return new cBool(true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cXOR() {
	}

	//***array-formula***
	cXOR.prototype = Object.create(cBaseFunction.prototype);
	cXOR.prototype.constructor = cXOR;
	cXOR.prototype.name = 'XOR';
	cXOR.prototype.argumentsMin = 1;
	cXOR.prototype.argumentsMax = 254;
	cXOR.prototype.isXLFN = true;
	cXOR.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cXOR.prototype.argumentsType = [[argType.logical]];
	cXOR.prototype.enabledToSingle = {"*": true};
	// cXOR.prototype.Calculate = function (arg) {
	// 	let argResult = null;
	// 	let nTrueValues = 0;
	// 	for (let i = 0; i < arg.length; i++) {
	// 		if (arg[i] instanceof cArea || arg[i] instanceof cArea3D) {
	// 			let allCellsEmpty = true;
	// 			let argArr = arg[i].getValue();
	// 			for (let j = 0; j < argArr.length; j++) {
	// 				let emptyArg = argArr[j] instanceof cEmpty;
	// 				if (argArr[j] instanceof cError) {
	// 					return argArr[j];
	// 				} else if (argArr[j] instanceof cString || emptyArg || argArr[j] instanceof cBool) {
	// 					argResult = new cBool(true);
	// 					nTrueValues++;
	// 				} else if (argArr.length === 1 && argArr[j] instanceof cNumber) {
	// 					if (argResult == null) {
	// 						argResult = argArr[j].tocBool();
	// 					} else {
	// 						argResult = new cBool(argArr[j].tocBool().value);
	// 					}

	// 					if (argResult.value === true) {
	// 						nTrueValues++;
	// 					}
	// 				}
	// 				if (!emptyArg) {
	// 					allCellsEmpty = false;
	// 				}
	// 			}
	// 			//if the range is empty - return an error
	// 			//if the range contains at least one non-empty cell (without an error) - the result is false
	// 			if (argResult === null && !allCellsEmpty) {
	// 				argResult = new cBool(false);
	// 			} else if (allCellsEmpty) {
	// 				argResult = null;
	// 			}
	// 		} else {
	// 			if (arg[i] instanceof cString) {
	// 				return new cError(cErrorType.wrong_value_type);
	// 			} else if (arg[i] instanceof cError) {
	// 				return arg[i];
	// 			} else if (arg[i] instanceof cArray) {
	// 				arg[i].foreach(function (elem) {
	// 					if (elem instanceof cError) {
	// 						argResult = elem;
	// 						return true;
	// 					} else if (elem instanceof cString || elem instanceof cEmpty) {
	// 						return false;
	// 					} else {
	// 						if (argResult === null) {
	// 							argResult = elem.tocBool();
	// 						} else {
	// 							argResult = new cBool(elem.tocBool().value);
	// 						}
	// 					}

	// 					if (argResult.value === true) {
	// 						nTrueValues++;
	// 					}
	// 				})
	// 			} else {
	// 				if (argResult === null) {
	// 					argResult = arg[i].tocBool();
	// 				} else {
	// 					argResult = new cBool(arg[i].tocBool().value);
	// 				}

	// 				if (argResult.value === true) {
	// 					nTrueValues++;
	// 				}
	// 			}
	// 		}
	// 	}
	// 	if (argResult === null) {
	// 		return new cError(cErrorType.wrong_value_type);
	// 	} else {
	// 		if (nTrueValues % 2) {
	// 			argResult = new cBool(true);
	// 		} else {
	// 			argResult = new cBool(false);
	// 		}
	// 	}

	// 	return argResult;
	// };
	cXOR.prototype.Calculate = function (arg) {
		let argResult = null;
		let nTrueValues = 0;
		for (let i = 0; i < arg.length; i++) {
			if (arg[i].type === cElementType.cell || arg[i].type === cElementType.cell3D) {
				arg[i] = arg[i].getValue();
			}
 
			if (arg[i].type === cElementType.error) {
				return arg[i];
			}
 
			if (arg[i].type === cElementType.cellsRange || arg[i].type === cElementType.cellsRange3D || arg[i].type === cElementType.array) {
				let dimensions = arg[i].getDimensions();
				let hasNonEmpty = false;
 
				for (let row = 0; row < dimensions.row; row++) {
					for (let col = 0; col < dimensions.col; col++) {
						let elemVal = arg[i].getValueByRowCol ? arg[i].getValueByRowCol(row, col, true) : arg[i].getElementRowCol(row, col);
 
						if (elemVal.type === cElementType.error) {
							return elemVal;
						}
 
						if (elemVal.type === cElementType.empty) {
							continue;
						}
 
						let argToBool = elemVal.tocBool();
						if (argToBool.type === cElementType.string) {
							continue;
						} else if (argToBool.type === cElementType.error) {
							return argToBool;
						} else if (argToBool.type === cElementType.bool) {
							hasNonEmpty = true;
							argResult = argToBool;
							if (argResult.value === true) {
								nTrueValues++;
							}
						}
					}
				}
 
				// if the range contains at least one non-empty cell (without an error) - mark result as valid
				if (argResult === null && hasNonEmpty) {
					argResult = new cBool(false);
				}
			} else {
				let argToBool = arg[i].tocBool();
				if (argToBool.type === cElementType.error) {
					return argToBool;
				} else if (argToBool.type === cElementType.bool) {
					argResult = argToBool;
					if (argResult.value === true) {
						nTrueValues++;
					}
				}
			}
		}
 
		if (argResult === null) {
			return new cError(cErrorType.wrong_value_type);
		}
 
		return new cBool(nTrueValues % 2 !== 0);
	};

})(window);
