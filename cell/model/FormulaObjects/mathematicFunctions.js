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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	var cElementType = AscCommonExcel.cElementType;
	var cErrorType = AscCommonExcel.cErrorType;
	var cExcelSignificantDigits = AscCommonExcel.cExcelSignificantDigits;
	var cNumber = AscCommonExcel.cNumber;
	var cString = AscCommonExcel.cString;
	var cBool = AscCommonExcel.cBool;
	var cError = AscCommonExcel.cError;
	var cArea = AscCommonExcel.cArea;
	var cArea3D = AscCommonExcel.cArea3D;
	var cRef = AscCommonExcel.cRef;
	var cRef3D = AscCommonExcel.cRef3D;
	var cEmpty = AscCommonExcel.cEmpty;
	var cArray = AscCommonExcel.cArray;
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var argType = Asc.c_oAscFormulaArgumentType;

	var _func = AscCommonExcel._func;

	var AGGREGATE_FUNC_AVE     = 1;
	var AGGREGATE_FUNC_CNT     = 2;
	var AGGREGATE_FUNC_CNTA    = 3;
	var AGGREGATE_FUNC_MAX     = 4;
	var AGGREGATE_FUNC_MIN     = 5;
	var AGGREGATE_FUNC_PROD    = 6;
	var AGGREGATE_FUNC_STD     = 7;
	var AGGREGATE_FUNC_STDP    = 8;
	var AGGREGATE_FUNC_SUM     = 9;
	var AGGREGATE_FUNC_VAR     = 10;
	var AGGREGATE_FUNC_VARP    = 11;
	var AGGREGATE_FUNC_MEDIAN  = 12;
	var AGGREGATE_FUNC_MODSNGL = 13;
	var AGGREGATE_FUNC_LARGE   = 14;
	var AGGREGATE_FUNC_SMALL   = 15;
	var AGGREGATE_FUNC_PERCINC = 16;
	var AGGREGATE_FUNC_QRTINC  = 17;
	var AGGREGATE_FUNC_PERCEXC = 18;
	var AGGREGATE_FUNC_QRTEXC  = 19;

	cFormulaFunctionGroup['Mathematic'] = cFormulaFunctionGroup['Mathematic'] || [];
	cFormulaFunctionGroup['Mathematic'].push(cABS, cACOS, cACOSH, cACOT, cACOTH, cAGGREGATE, cARABIC, cASIN, cASINH,
		cATAN, cATAN2, cATANH, cBASE, cCEILING, cCEILING_MATH, cCEILING_PRECISE, cCOMBIN, cCOMBINA, cCOS, cCOSH, cCOT,
		cCOTH, cCSC, cCSCH, cDECIMAL, cDEGREES, cECMA_CEILING, cEVEN, cEXP, cFACT, cFACTDOUBLE, cFLOOR, cFLOOR_PRECISE,
		cFLOOR_MATH, cGCD, cINT, cISO_CEILING, cLCM, cLN, cLOG, cLOG10, cMDETERM, cMINVERSE, cMMULT, cMOD, cMROUND,
		cMULTINOMIAL, cMUNIT, cODD, cPI, cPOWER, cPRODUCT, cQUOTIENT, cRADIANS, cRAND, cRANDARRAY, cRANDBETWEEN, cROMAN, cROUND, cROUNDDOWN,
		cROUNDUP, cSEC, cSECH, cSERIESSUM, cSIGN, cSIN, cSINGLE, cSINH, cSQRT, cSQRTPI, cSUBTOTAL, cSUM, cSUMIF, cSUMIFS,
		cSUMPRODUCT, cSUMSQ, cSUMX2MY2, cSUMX2PY2, cSUMXMY2, cTAN, cTANH, cTRUNC, cSEQUENCE);

	var cSubTotalFunctionType = {
		includes: {
			AVERAGE: 1, COUNT: 2, COUNTA: 3, MAX: 4, MIN: 5, PRODUCT: 6, STDEV: 7, STDEVP: 8, SUM: 9, VAR: 10, VARP: 11
		}, excludes: {
			AVERAGE: 101,
			COUNT: 102,
			COUNTA: 103,
			MAX: 104,
			MIN: 105,
			PRODUCT: 106,
			STDEV: 107,
			STDEVP: 108,
			SUM: 109,
			VAR: 110,
			VARP: 111
		}
	};

	function sumxCalc(arg, func) {

		function sumX(a, b, _3d) {
			var sum = 0, i, j;

			if (!_3d) {
				if (a.length == b.length && a[0] && b[0] && a[0].length == b[0].length) {
					for (i = 0; i < a.length; i++) {
						for (j = 0; j < a[0].length; j++) {
							if (a[i][j] instanceof cNumber && b[i][j] instanceof cNumber) {
								sum += func(a[i][j].getValue(), b[i][j].getValue())
							} else {
								return new cError(cErrorType.wrong_value_type);
							}
						}
					}
					return new cNumber(sum);
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else {
				if (a.length == b.length && a[0] && b[0] && a[0].length == b[0].length && a[0][0] && b[0][0] && a[0][0].length == b[0][0].length) {
					for (i = 0; i < a.length; i++) {
						for (j = 0; j < a[0].length; j++) {
							for (var k = 0; k < a[0][0].length; k++) {
								if (a[i][j][k] instanceof cNumber && b[i][j][k] instanceof cNumber) {
									sum += func(a[i][j][k].getValue(), b[i][j][k].getValue())
								} else {
									return new cError(cErrorType.wrong_value_type);
								}
							}
						}
					}
					return new cNumber(sum);
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			}
		}

		var arg0 = arg[0], arg1 = arg[1];

		if (arg0 instanceof cArea3D && arg1 instanceof cArea3D) {
			return sumX(arg0.getMatrix(), arg1.getMatrix(), true);
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		}

		if (arg0 instanceof cArea || arg0 instanceof cArray) {
			arg0 = arg0.getMatrix();
		} else if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cNumber) {
			arg0 = [[arg0]];
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
		}

		if (arg1 instanceof cArea || arg1 instanceof cArray || arg1 instanceof cArea3D) {
			arg1 = arg1.getMatrix();
		} else if (arg1 instanceof cError) {
			return arg1;
		} else if (arg1 instanceof cNumber) {
			arg1 = [[arg1]];
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		return sumX(arg0, arg1, false);
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cABS() {
	}

	//***array-formula***
	cABS.prototype = Object.create(cBaseFunction.prototype);
	cABS.prototype.constructor = cABS;
	cABS.prototype.name = 'ABS';
	cABS.prototype.argumentsMin = 1;
	cABS.prototype.argumentsMax = 1;
	cABS.prototype.argumentsType = [argType.number];
	cABS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = new cNumber(Math.abs(elem.getValue()));
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
		} else {
			return new cNumber(Math.abs(arg0.getValue()));
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cACOS() {
	}

	//***array-formula***
	cACOS.prototype = Object.create(cBaseFunction.prototype);
	cACOS.prototype.constructor = cACOS;
	cACOS.prototype.name = 'ACOS';
	cACOS.prototype.argumentsMin = 1;
	cACOS.prototype.argumentsMax = 1;
	cACOS.prototype.argumentsType = [argType.number];
	cACOS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.acos(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.acos(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cACOSH() {
	}

	//***array-formula***
	cACOSH.prototype = Object.create(cBaseFunction.prototype);
	cACOSH.prototype.constructor = cACOSH;
	cACOSH.prototype.name = 'ACOSH';
	cACOSH.prototype.argumentsMin = 1;
	cACOSH.prototype.argumentsMax = 1;
	cACOSH.prototype.argumentsType = [argType.number];
	cACOSH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.acosh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.acosh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cACOT() {
	}

	//***array-formula***
	cACOT.prototype = Object.create(cBaseFunction.prototype);
	cACOT.prototype.constructor = cACOT;
	cACOT.prototype.name = 'ACOT';
	cACOT.prototype.argumentsMin = 1;
	cACOT.prototype.argumentsMax = 1;
	cACOT.prototype.isXLFN = true;
	cACOT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cACOT.prototype.argumentsType = [argType.number];
	cACOT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					var a = Math.PI / 2 - Math.atan(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			var a = Math.PI / 2 - Math.atan(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cACOTH() {
	}

	//***array-formula***
	cACOTH.prototype = Object.create(cBaseFunction.prototype);
	cACOTH.prototype.constructor = cACOTH;
	cACOTH.prototype.name = 'ACOTH';
	cACOTH.prototype.argumentsMin = 1;
	cACOTH.prototype.argumentsMax = 1;
	cACOTH.prototype.isXLFN = true;
	cACOTH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cACOTH.prototype.argumentsType = [argType.number];
	cACOTH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					var a = Math.atanh(1 / elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			var a = Math.atanh(1 / arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cAGGREGATE() {
	}

	//***array-formula***
	cAGGREGATE.prototype = Object.create(cBaseFunction.prototype);
	cAGGREGATE.prototype.constructor = cAGGREGATE;
	cAGGREGATE.prototype.name = 'AGGREGATE';
	cAGGREGATE.prototype.argumentsMin = 3;
	cAGGREGATE.prototype.argumentsMax = 253;
	cAGGREGATE.prototype.isXLFN = true;
	cAGGREGATE.prototype.argumentsType = [argType.number, argType.number, [argType.reference]];
	cAGGREGATE.prototype.arrayIndexes = {1: 1, 2: 1, 3: 1, 4: 1, 5: 1, 6: 1};
	cAGGREGATE.prototype.getArrayIndex = function (index) {
		if (index === 0) {
			return undefined;
		}
		return 1;
	};
	cAGGREGATE.prototype.Calculate = function (arg) {
		let oArguments = this._prepareArguments([arg[0], arg[1]], arguments[1]);
		let argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();

		let argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		let nFunc = argClone[0].getValue();
		let f = null;
		switch (nFunc) {
			case AGGREGATE_FUNC_AVE:
				f = AscCommonExcel.cAVERAGE.prototype;
				break;
			case AGGREGATE_FUNC_CNT:
				f = AscCommonExcel.cCOUNT.prototype;
				break;
			case AGGREGATE_FUNC_CNTA:
				f = AscCommonExcel.cCOUNTA.prototype;
				break;
			case AGGREGATE_FUNC_MAX:
				f = AscCommonExcel.cMAX.prototype;
				break;
			case AGGREGATE_FUNC_MIN:
				f = AscCommonExcel.cMIN.prototype;
				break;
			case AGGREGATE_FUNC_PROD:
				f = AscCommonExcel.cPRODUCT.prototype;
				break;
			case AGGREGATE_FUNC_STD:
				f = AscCommonExcel.cSTDEV_S.prototype;
				break;
			case AGGREGATE_FUNC_STDP:
				f = AscCommonExcel.cSTDEV_P.prototype;
				break;
			case AGGREGATE_FUNC_SUM:
				f = AscCommonExcel.cSUM.prototype;
				break;
			case AGGREGATE_FUNC_VAR:
				f = AscCommonExcel.cVAR_S.prototype;
				break;
			case AGGREGATE_FUNC_VARP:
				f = AscCommonExcel.cVAR_P.prototype;
				break;
			case AGGREGATE_FUNC_MEDIAN:
				f = AscCommonExcel.cMEDIAN.prototype;
				break;
			case AGGREGATE_FUNC_MODSNGL:
				f = AscCommonExcel.cMODE_SNGL.prototype;
				break;
			case AGGREGATE_FUNC_LARGE:
				if (arg[3]) {
					f = AscCommonExcel.cLARGE.prototype;
				}
				break;
			case AGGREGATE_FUNC_SMALL:
				if (arg[3]) {
					f = AscCommonExcel.cSMALL.prototype;
				}
				break;
			case AGGREGATE_FUNC_PERCINC:
				if (arg[3]) {
					f = AscCommonExcel.cPERCENTILE_INC.prototype;
				}
				break;
			case AGGREGATE_FUNC_QRTINC:
				if (arg[3]) {
					f = AscCommonExcel.cQUARTILE_INC.prototype;
				}
				break;
			case AGGREGATE_FUNC_PERCEXC:
				if (arg[3]) {
					f = AscCommonExcel.cPERCENTILE_EXC.prototype;
				}
				break;
			case AGGREGATE_FUNC_QRTEXC:
				if (arg[3]) {
					f = AscCommonExcel.cQUARTILE_EXC.prototype;
				}
				break;
			default:
				return new cError(cErrorType.not_numeric);
		}

		if (null === f) {
			return new cError(cErrorType.wrong_value_type);
		}

		let nOption = argClone[1].getValue();
		let ignoreHiddenRows = false;
		let ignoreErrorsVal = false;
		let ignoreNestedStAg = false;
		switch (nOption) {
			case 0 : // ignore nested SUBTOTAL and AGGREGATE functions
				ignoreNestedStAg = true;
				break;
			case 1 : // ignore hidden rows, nested SUBTOTAL and AGGREGATE functions
				ignoreNestedStAg = true;
				ignoreHiddenRows = true;
				break;
			case 2 : // ignore error values, nested SUBTOTAL and AGGREGATE functions
				ignoreNestedStAg = true;
				ignoreErrorsVal = true;
				break;
			case 3 : // ignore hidden rows, error values, nested SUBTOTAL and AGGREGATE functions
				ignoreNestedStAg = true;
				ignoreErrorsVal = true;
				ignoreHiddenRows = true;
				break;
			case 4 : // ignore nothing
				break;
			case 5 : // ignore hidden rows
				ignoreHiddenRows = true;
				break;
			case 6 : // ignore error values
				ignoreErrorsVal = true;
				break;
			case 7 : // ignore hidden rows and error values
				ignoreHiddenRows = true;
				ignoreErrorsVal = true;
				break;
			default :
				return new cError(cErrorType.not_numeric);
		}

		let res;
		if (f) {
			let oldExcludeHiddenRows = f.excludeHiddenRows;
			let oldExcludeErrorsVal = f.excludeErrorsVal;
			let oldIgnoreNestedStAg = f.excludeNestedStAg;

			f.excludeHiddenRows = ignoreHiddenRows;
			f.excludeErrorsVal = ignoreErrorsVal;
			f.excludeNestedStAg = ignoreNestedStAg;

			let newArgs = [];
			//14 - 19 особенные функции, требующие второго аргумента
			let doNotCheckRef = nFunc >= 14 && nFunc <= 19;
			for (let i = 2; i < arg.length; i++) {
				//аргумент может быть только ссылка на ячейку или диапазон ячеек
				//в противном случае - ошибка
				if (doNotCheckRef || this.checkRef(arg[i])) {
					newArgs.push(arg[i]);
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			}

			if (f.argumentsMax && newArgs.length > f.argumentsMax) {
				return new cError(cErrorType.wrong_value_type);
			}

			res = f.Calculate(newArgs);

			f.excludeHiddenRows = oldExcludeHiddenRows;
			f.excludeErrorsVal = oldExcludeErrorsVal;
			f.excludeNestedStAg = oldIgnoreNestedStAg;
		}

		return res;
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cARABIC() {
	}

	//***array-formula***
	cARABIC.prototype = Object.create(cBaseFunction.prototype);
	cARABIC.prototype.constructor = cARABIC;
	cARABIC.prototype.name = 'ARABIC';
	cARABIC.prototype.argumentsMin = 1;
	cARABIC.prototype.argumentsMax = 1;
	cARABIC.prototype.isXLFN = true;
	cARABIC.prototype.argumentsType = [argType.text];
	cARABIC.prototype.Calculate = function (arg) {
		var to_arabic = function (roman) {
			roman = roman.toUpperCase();
			if (roman < 1) {
				return 0;
			} else if (!/^M*(?:D?C{0,3}|C[MD])(?:L?X{0,3}|X[CL])(?:V?I{0,3}|I[XV])$/i.test(roman)) {
				return NaN;
			}

			var chars = {
				"M": 1000,
				"CM": 900,
				"D": 500,
				"CD": 400,
				"C": 100,
				"XC": 90,
				"L": 50,
				"XL": 40,
				"X": 10,
				"IX": 9,
				"V": 5,
				"IV": 4,
				"I": 1
			};
			var arabic = 0;
			roman.replace(/[MDLV]|C[MD]?|X[CL]?|I[XV]?/g, function (i) {
				arabic += chars[i];
			});

			return arabic;
		};

		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			return new cError(cErrorType.wrong_value_type);
		}
		arg0 = arg0.tocString();

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				var a = elem;
				if (cElementType.string === a.type) {
					var res = to_arabic(a.getValue());
					this.array[r][c] = isNaN(res) ? new cError(cErrorType.wrong_value_type) : new cNumber(res);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		}

		//TODO проверить возвращение ошибок!
		var res = to_arabic(arg0.getValue());
		return isNaN(res) ? new cError(cErrorType.wrong_value_type) : new cNumber(res);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cASIN() {
	}

	//***array-formula***
	cASIN.prototype = Object.create(cBaseFunction.prototype);
	cASIN.prototype.constructor = cASIN;
	cASIN.prototype.name = 'ASIN';
	cASIN.prototype.argumentsMin = 1;
	cASIN.prototype.argumentsMax = 1;
	cASIN.prototype.argumentsType = [argType.number];
	cASIN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.asin(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.asin(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cASINH() {
	}

	//***array-formula***
	cASINH.prototype = Object.create(cBaseFunction.prototype);
	cASINH.prototype.constructor = cASINH;
	cASINH.prototype.name = 'ASINH';
	cASINH.prototype.argumentsMin = 1;
	cASINH.prototype.argumentsMax = 1;
	cASINH.prototype.argumentsType = [argType.number];
	cASINH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.asinh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.asinh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cATAN() {
	}

	//***array-formula***
	cATAN.prototype = Object.create(cBaseFunction.prototype);
	cATAN.prototype.constructor = cATAN;
	cATAN.prototype.name = 'ATAN';
	cATAN.prototype.argumentsMin = 1;
	cATAN.prototype.argumentsMax = 1;
	cATAN.prototype.argumentsType = [argType.number];
	cATAN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.atan(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			var a = Math.atan(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cATAN2() {
	}

	cATAN2.prototype = Object.create(cBaseFunction.prototype);
	cATAN2.prototype.constructor = cATAN2;
	cATAN2.prototype.name = 'ATAN2';
	cATAN2.prototype.argumentsMin = 2;
	cATAN2.prototype.argumentsMax = 2;
	cATAN2.prototype.argumentsType = [argType.number, argType.number];
	cATAN2.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();
		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem, b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] =
							a.getValue() == 0 && b.getValue() == 0 ? new cError(cErrorType.division_by_zero) :
								new cNumber(Math.atan2(b.getValue(), a.getValue()))
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem, b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] =
						a.getValue() == 0 && b.getValue() == 0 ? new cError(cErrorType.division_by_zero) :
							new cNumber(Math.atan2(b.getValue(), a.getValue()))
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0, b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] =
						a.getValue() == 0 && b.getValue() == 0 ? new cError(cErrorType.division_by_zero) :
							new cNumber(Math.atan2(b.getValue(), a.getValue()))
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		return (arg0 instanceof cError ? arg0 : arg1 instanceof cError ? arg1 :
				arg1.getValue() == 0 && arg0.getValue() == 0 ? new cError(cErrorType.division_by_zero) :
					new cNumber(Math.atan2(arg1.getValue(), arg0.getValue()))
		)
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cATANH() {
	}

	//***array-formula***
	cATANH.prototype = Object.create(cBaseFunction.prototype);
	cATANH.prototype.constructor = cATANH;
	cATANH.prototype.name = 'ATANH';
	cATANH.prototype.argumentsMin = 1;
	cATANH.prototype.argumentsMax = 1;
	cATANH.prototype.argumentsType = [argType.number];
	cATANH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.atanh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.atanh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cBASE() {
	}

	//***array-formula***
	cBASE.prototype = Object.create(cBaseFunction.prototype);
	cBASE.prototype.constructor = cBASE;
	cBASE.prototype.name = 'BASE';
	cBASE.prototype.argumentsMin = 2;
	cBASE.prototype.argumentsMax = 3;
	cBASE.prototype.isXLFN = true;
	cBASE.prototype.argumentsType = [argType.number, argType.number, argType.number];
	cBASE.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();
		argClone[2] = argClone[2] ? argClone[2].tocNumber() : new cNumber(0);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function base_math(argArray) {
			var number = parseInt(argArray[0]);
			var radix = parseInt(argArray[1]);
			var min_length = parseInt(argArray[2]);

			if (radix < 2 || radix > 36 || number < 0 || number > Math.pow(2, 53) || min_length < 0) {
				return new cError(cErrorType.not_numeric);
			}

			var str = number.toString(radix);
			str = str.toUpperCase();
			if (str.length < min_length) {
				var prefix = "";
				for (var i = 0; i < min_length - str.length; i++) {
					prefix += "0";
				}
				str = prefix + str;
			}

			return new cString(str);
		}

		return this._findArrayInNumberArguments(oArguments, base_math);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCEILING() {
	}

	//***array-formula***
	cCEILING.prototype = Object.create(cBaseFunction.prototype);
	cCEILING.prototype.constructor = cCEILING;
	cCEILING.prototype.name = 'CEILING';
	cCEILING.prototype.argumentsMin = 2;
	cCEILING.prototype.argumentsMax = 2;
	cCEILING.prototype.argumentsType = [argType.number, argType.number];
	cCEILING.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg0 = arg[0].tocNumber();
		arg1 = arg[1].tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		function ceilingHelper(number, significance) {
			if (significance == 0) {
				return new cNumber(0.0);
			}
			if (number > 0 && significance < 0) {
				return new cError(cErrorType.not_numeric);
			} else if (number / significance === Infinity) {
				return new cError(cErrorType.not_numeric);
			} else {
				var quotient = number / significance;
				if (quotient == 0) {
					return new cNumber(0.0);
				}
				var quotientTr = Math.floor(quotient);

				var nolpiat = 5 * Math.sign(quotient) *
					Math.pow(10, Math.floor(Math.log10(Math.abs(quotient))) - cExcelSignificantDigits);

				if (Math.abs(quotient - quotientTr) > nolpiat) {
					++quotientTr;
				}
				return new cNumber(quotientTr * significance);
			}
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem, b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = ceilingHelper(a.getValue(), b.getValue())
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem, b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = ceilingHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0, b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = ceilingHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		return ceilingHelper(arg0.getValue(), arg1.getValue());

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCEILING_MATH() {
	}

	//***array-formula***
	cCEILING_MATH.prototype = Object.create(cBaseFunction.prototype);
	cCEILING_MATH.prototype.constructor = cCEILING_MATH;
	cCEILING_MATH.prototype.name = 'CEILING.MATH';
	cCEILING_MATH.prototype.argumentsMin = 1;
	cCEILING_MATH.prototype.argumentsMax = 3;
	cCEILING_MATH.prototype.isXLFN = true;
	cCEILING_MATH.prototype.argumentsType = [argType.number, argType.number, argType.number];
	cCEILING_MATH.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		if (!argClone[1]) {
			argClone[1] = argClone[0] > 0 ? new cNumber(1) : new cNumber(-1);
		}
		argClone[2] = argClone[2] ? argClone[2].tocNumber() : new cNumber(0);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function floor_math(argArray) {
			var number = argArray[0];
			var significance = argArray[1];
			var mod = argArray[2];

			if (significance === 0 || number === 0) {
				return new cNumber(0.0);
			}

			if (number * significance < 0.0) {
				significance = -significance;
			}

			if (mod === 0 && number < 0.0) {
				return new cNumber(Math.floor(number / significance) * significance);
			} else {
				return new cNumber(Math.ceil(number / significance) * significance);
			}
		}

		return this._findArrayInNumberArguments(oArguments, floor_math);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCEILING_PRECISE() {
	}

	cCEILING_PRECISE.prototype = Object.create(cBaseFunction.prototype);
	cCEILING_PRECISE.prototype.constructor = cCEILING_PRECISE;
	cCEILING_PRECISE.prototype.name = 'CEILING.PRECISE';
	cCEILING_PRECISE.prototype.argumentsMin = 1;
	cCEILING_PRECISE.prototype.argumentsMax = 2;
	cCEILING_PRECISE.prototype.isXLFN = true;
	cCEILING_PRECISE.prototype.argumentsType = [argType.number, argType.number];
	cCEILING_PRECISE.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1] ? argClone[1].tocNumber() : new cNumber(1);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function floorHelper(argArray) {
			var number = argArray[0];
			var significance = argArray[1];
			if (significance === 0 || number === 0) {
				return new cNumber(0.0);
			}

			var absSignificance = Math.abs(significance);
			var quotient = number / absSignificance;
			return new cNumber(Math.ceil(quotient) * absSignificance);
		}

		return this._findArrayInNumberArguments(oArguments, floorHelper);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOMBIN() {
	}

	//***array-formula***
	cCOMBIN.prototype = Object.create(cBaseFunction.prototype);
	cCOMBIN.prototype.constructor = cCOMBIN;
	cCOMBIN.prototype.name = 'COMBIN';
	cCOMBIN.prototype.argumentsMin = 2;
	cCOMBIN.prototype.argumentsMax = 2;
	cCOMBIN.prototype.argumentsType = [argType.number, argType.number];
	cCOMBIN.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1];

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.array === arg0.type && cElementType.array === arg1.type) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					let a = elem, b = arg1.getElementRowCol(r, c);
					if (a.type === cElementType.number && b.type === cElementType.number) {
						let resVal = Math.binomCoeff(a.getValue(), b.getValue());
						if (isNaN(resVal)) {
							resVal = arg0.getValue() !== arg1.getValue() ? new cError(cErrorType.not_numeric) : new cNumber(1);
						} else {
							resVal = new cNumber(resVal);
						}
						this.array[r][c] = resVal;
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				let a = elem, b = arg1;
				if (a.type === cElementType.number && b.type === cElementType.number) {
					if (a.getValue() < 0 || b.getValue() < 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					}

					let resVal = Math.binomCoeff(a.getValue(), b.getValue());
					if (isNaN(resVal)) {
						resVal = arg0.getValue() !== arg1.getValue() ? new cError(cErrorType.not_numeric) : new cNumber(1);
					} else {
						resVal = new cNumber(resVal);
					}

					this.array[r][c] = resVal;
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (cElementType.array === arg1.type) {
			arg1.foreach(function (elem, r, c) {
				let a = arg0, b = elem;
				if (a.type === cElementType.number && b.type === cElementType.number) {
					if (a.getValue() < 0 || b.getValue() < 0 || a.getValue() < b.getValue()) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					}

					let resVal = Math.binomCoeff(a.getValue(), b.getValue());
					if (isNaN(resVal)) {
						resVal = arg0.getValue() !== arg1.getValue() ? new cError(cErrorType.not_numeric) : new cNumber(1);
					} else {
						resVal = new cNumber(resVal);
					}

					this.array[r][c] = resVal;
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		if (arg0.getValue() < 0 || arg1.getValue() < 0 || arg0.getValue() < arg1.getValue()) {
			return new cError(cErrorType.not_numeric);
		}

		let res = Math.binomCoeff(arg0.getValue(), arg1.getValue());

		if (isNaN(res)) {
			return arg0.getValue() !== arg1.getValue() ? new cError(cErrorType.not_numeric) : new cNumber(1);
		}

		return new cNumber(res);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOMBINA() {
	}

	//***array-formula***
	cCOMBINA.prototype = Object.create(cBaseFunction.prototype);
	cCOMBINA.prototype.constructor = cCOMBINA;
	cCOMBINA.prototype.name = 'COMBINA';
	cCOMBINA.prototype.argumentsMin = 2;
	cCOMBINA.prototype.argumentsMax = 2;
	cCOMBINA.prototype.isXLFN = true;
	cCOMBINA.prototype.argumentsType = [argType.number, argType.number];
	cCOMBINA.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1].tocNumber();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function combinaCalculate(argArray) {
			var a = argArray[0];
			var b = argArray[1];

			if (a < 0 || b < 0 || a < b) {
				return new cError(cErrorType.not_numeric);
			}

			a = parseInt(a);
			b = parseInt(b);
			return new cNumber(Math.binomCoeff(a + b - 1, b));
		}

		return this._findArrayInNumberArguments(oArguments, combinaCalculate);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOS() {
	}

	//***array-formula***
	cCOS.prototype = Object.create(cBaseFunction.prototype);
	cCOS.prototype.constructor = cCOS;
	cCOS.prototype.name = 'COS';
	cCOS.prototype.argumentsMin = 1;
	cCOS.prototype.argumentsMax = 1;
	cCOS.prototype.argumentsType = [argType.number];
	cCOS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.cos(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.cos(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOSH() {
	}

	//***array-formula***
	cCOSH.prototype = Object.create(cBaseFunction.prototype);
	cCOSH.prototype.constructor = cCOSH;
	cCOSH.prototype.name = 'COSH';
	cCOSH.prototype.argumentsMin = 1;
	cCOSH.prototype.argumentsMax = 1;
	cCOSH.prototype.argumentsType = [argType.number];
	cCOSH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.cosh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.cosh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOT() {
	}

	//***array-formula***
	cCOT.prototype = Object.create(cBaseFunction.prototype);
	cCOT.prototype.constructor = cCOT;
	cCOT.prototype.name = 'COT';
	cCOT.prototype.argumentsMin = 1;
	cCOT.prototype.argumentsMax = 1;
	cCOT.prototype.isXLFN = true;
	cCOT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCOT.prototype.argumentsType = [argType.number];
	cCOT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		var maxVal = Math.pow(2, 27);
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					if (0 === elem.getValue()) {
						this.array[r][c] = new cError(cErrorType.division_by_zero);
					} else if (Math.abs(elem.getValue()) >= maxVal) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						var a = 1 / Math.tan(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			if (0 === arg0.getValue()) {
				return new cError(cErrorType.division_by_zero);
			} else if (Math.abs(arg0.getValue()) >= maxVal) {
				return new cError(cErrorType.not_numeric);
			}

			var a = 1 / Math.tan(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOTH() {
	}

	//***array-formula***
	cCOTH.prototype = Object.create(cBaseFunction.prototype);
	cCOTH.prototype.constructor = cCOTH;
	cCOTH.prototype.name = 'COTH';
	cCOTH.prototype.argumentsMin = 1;
	cCOTH.prototype.argumentsMax = 1;
	cCOTH.prototype.isXLFN = true;
	cCOTH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCOTH.prototype.argumentsType = [argType.number];
	cCOTH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		//TODO в документации к COTH написано максимальное значение - Math.pow(2, 27), но MS EXCEL в данном случае не выдает ошибку
		//проверку на максиимальное значение убрал
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					if (0 === elem.getValue()) {
						this.array[r][c] = new cError(cErrorType.division_by_zero);
					} else {
						var a = 1 / Math.tanh(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			if (0 === arg0.getValue()) {
				return new cError(cErrorType.division_by_zero);
			}

			var a = 1 / Math.tanh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCSC() {
	}

	//***array-formula***
	cCSC.prototype = Object.create(cBaseFunction.prototype);
	cCSC.prototype.constructor = cCSC;
	cCSC.prototype.name = 'CSC';
	cCSC.prototype.argumentsMin = 1;
	cCSC.prototype.argumentsMax = 1;
	cCSC.prototype.isXLFN = true;
	cCSC.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCSC.prototype.argumentsType = [argType.number];
	cCSC.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		var maxVal = Math.pow(2, 27);
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					if (0 === elem.getValue()) {
						this.array[r][c] = new cError(cErrorType.division_by_zero);
					} else if (Math.abs(elem.getValue()) >= maxVal) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						var a = 1 / Math.sin(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			if (0 === arg0.getValue()) {
				return new cError(cErrorType.division_by_zero);
			} else if (Math.abs(arg0.getValue()) >= maxVal) {
				return new cError(cErrorType.not_numeric);
			}

			var a = 1 / Math.sin(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCSCH() {
	}

	//***array-formula***
	cCSCH.prototype = Object.create(cBaseFunction.prototype);
	cCSCH.prototype.constructor = cCSCH;
	cCSCH.prototype.name = 'CSCH';
	cCSCH.prototype.argumentsMin = 1;
	cCSCH.prototype.argumentsMax = 1;
	cCSCH.prototype.isXLFN = true;
	cCSCH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCSCH.prototype.argumentsType = [argType.number];
	cCSCH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		//TODO в документации к COTH написано максимальное значение - Math.pow(2, 27), но MS EXCEL в данном случае не выдает ошибку
		//проверку на максиимальное значение убрал
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					if (0 === elem.getValue()) {
						this.array[r][c] = new cError(cErrorType.division_by_zero);
					} else {
						var a = 1 / Math.sinh(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			if (0 === arg0.getValue()) {
				return new cError(cErrorType.division_by_zero);
			}

			var a = 1 / Math.sinh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDECIMAL() {
	}

	//***array-formula***
	cDECIMAL.prototype = Object.create(cBaseFunction.prototype);
	cDECIMAL.prototype.constructor = cDECIMAL;
	cDECIMAL.prototype.name = 'DECIMAL';
	cDECIMAL.prototype.argumentsMin = 2;
	cDECIMAL.prototype.argumentsMax = 2;
	cDECIMAL.prototype.isXLFN = true;
	cDECIMAL.prototype.argumentsType = [argType.text, argType.number];
	cDECIMAL.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocString();
		argClone[1] = argClone[1].tocNumber();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function decimal_calculate(argArray) {
			var a = argArray[0];
			var b = argArray[1];
			b = parseInt(b);

			if (b < 2 || b > 36) {
				return new cError(cErrorType.not_numeric);
			}

			var fVal = 0;
			var startIndex = 0;
			while (a[startIndex] === ' ') {
				startIndex++;
			}

			if (b === 16) {
				if (a[startIndex] === 'x' || a[startIndex] === 'X') {
					startIndex++;
				} else if (a[startIndex] === '0' && (a[startIndex + 1] === 'x' || a[startIndex + 1] === 'X')) {
					startIndex += 2;
				}
			}

			a = a.toLowerCase();
			var startPos = 'a'.charCodeAt(0);
			for (var i = startIndex; i < a.length; i++) {
				var n;
				if ('0' <= a[i] && a[i] <= '9') {
					n = a[i] - '0';
				} else if ('a' <= a[i] && a[i] <= 'z') {
					var currentPos = a[i].charCodeAt(0);
					n = 10 + parseInt(currentPos - startPos);
				} else {
					n = b;
				}

				if (b <= n) {
					if (i + 1 === a.length && ((b == 2 && (a[i] == 'b' || a[i] == 'B')) || (b == 16 && (a[i] == 'h' || a[i] == 'H')))) {
						;
					} else {
						return new cError(cErrorType.not_numeric);
					}
				} else {
					fVal = fVal * b + n;
				}
			}

			return isNaN(fVal) ? new cError(cErrorType.not_numeric) : new cNumber(fVal);
		}

		return this._findArrayInNumberArguments(oArguments, decimal_calculate, true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDEGREES() {
	}

	//***array-formula***
	cDEGREES.prototype = Object.create(cBaseFunction.prototype);
	cDEGREES.prototype.constructor = cDEGREES;
	cDEGREES.prototype.name = 'DEGREES';
	cDEGREES.prototype.argumentsMin = 1;
	cDEGREES.prototype.argumentsMax = 1;
	cDEGREES.prototype.argumentsType = [argType.number];
	cDEGREES.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = elem.getValue();
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a * 180 / Math.PI);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = arg0.getValue();
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a * 180 / Math.PI);
		}
		return arg0;

	};

	/**
	 * @constructor
	 * @extends {cCEILING}
	 */
	//TODO нигде нет отписания к этой функции! работает так же как и cCEILING на всех примерах.
	function cECMA_CEILING() {
	}

	//***array-formula***
	cECMA_CEILING.prototype = Object.create(cCEILING.prototype);
	cECMA_CEILING.prototype.constructor = cECMA_CEILING;
	cECMA_CEILING.prototype.name = 'ECMA.CEILING';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEVEN() {
	}

	//***array-formula***
	cEVEN.prototype = Object.create(cBaseFunction.prototype);
	cEVEN.prototype.constructor = cEVEN;
	cEVEN.prototype.name = 'EVEN';
	cEVEN.prototype.argumentsMin = 1;
	cEVEN.prototype.argumentsMax = 1;
	cEVEN.prototype.argumentsType = [argType.number];
	cEVEN.prototype.Calculate = function (arg) {

		function evenHelper(arg) {
			var arg0 = arg.getValue();
			if (arg0 >= 0) {
				arg0 = Math.ceil(arg0);
				if ((arg0 & 1) == 0) {
					return new cNumber(arg0);
				} else {
					return new cNumber(arg0 + 1);
				}
			} else {
				arg0 = Math.floor(arg0);
				if ((arg0 & 1) == 0) {
					return new cNumber(arg0);
				} else {
					return new cNumber(arg0 - 1);
				}
			}
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}

		arg0 = arg0.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = evenHelper(elem);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg0 instanceof cNumber) {
			return evenHelper(arg0);
		}
		return new cError(cErrorType.wrong_value_type);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEXP() {
	}

	//***array-formula***
	cEXP.prototype = Object.create(cBaseFunction.prototype);
	cEXP.prototype.constructor = cEXP;
	cEXP.prototype.name = 'EXP';
	cEXP.prototype.argumentsMin = 1;
	cEXP.prototype.argumentsMax = 1;
	cEXP.prototype.argumentsType = [argType.number];
	cEXP.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.exp(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		}
		if (!(arg0 instanceof cNumber)) {
			return new cError(cErrorType.not_numeric);
		} else {
			var a = Math.exp(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFACT() {
	}

	//***array-formula***
	cFACT.prototype = Object.create(cBaseFunction.prototype);
	cFACT.prototype.constructor = cFACT;
	cFACT.prototype.name = 'FACT';
	cFACT.prototype.argumentsMin = 1;
	cFACT.prototype.argumentsMax = 1;
	cFACT.prototype.argumentsType = [argType.number];
	cFACT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					if (elem.getValue() < 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						var a = Math.fact(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			if (arg0.getValue() < 0) {
				return new cError(cErrorType.not_numeric);
			}
			var a = Math.fact(arg0.getValue());
			return isNaN(a) || a == Infinity ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFACTDOUBLE() {
	}

	//***array-formula***
	cFACTDOUBLE.prototype = Object.create(cBaseFunction.prototype);
	cFACTDOUBLE.prototype.constructor = cFACTDOUBLE;
	cFACTDOUBLE.prototype.name = 'FACTDOUBLE';
	cFACTDOUBLE.prototype.argumentsMin = 1;
	cFACTDOUBLE.prototype.argumentsMax = 1;
	cFACTDOUBLE.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cFACTDOUBLE.prototype.argumentsType = [argType.any];
	cFACTDOUBLE.prototype.Calculate = function (arg) {
		function factDouble(n) {
			if (n == 0) {
				return 0;
			} else if (n < 0) {
				return Number.NaN;
			} else if (n > 300) {
				return Number.Infinity;
			}
			n = Math.floor(n);
			var res = n, _n = n, ost = -(_n & 1);
			n -= 2;

			while (n != ost) {
				res *= n;
				n -= 2;
			}
			return res;
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					if (elem.getValue() < 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						var a = factDouble(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			if (arg0.getValue() < 0) {
				return new cError(cErrorType.not_numeric);
			}
			var a = factDouble(arg0.getValue());
			return isNaN(a) || a == Infinity ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFLOOR() {
	}

	cFLOOR.prototype = Object.create(cBaseFunction.prototype);
	cFLOOR.prototype.constructor = cFLOOR;
	cFLOOR.prototype.name = 'FLOOR';
	cFLOOR.prototype.argumentsMin = 2;
	cFLOOR.prototype.argumentsMax = 2;
	cFLOOR.prototype.argumentsType = [argType.number, argType.number];
	cFLOOR.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		arg0 = arg[0].tocNumber();
		arg1 = arg[1].tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		function floorHelper(number, significance) {
			if (significance == 0) {
				return new cNumber(0.0);
			}
			if ((number > 0 && significance < 0) || (number < 0 && significance > 0)) {
				return new cError(cErrorType.not_numeric);
			} else if (number / significance === Infinity) {
				return new cError(cErrorType.not_numeric);
			} else {
				var quotient = number / significance;
				if (quotient == 0) {
					return new cNumber(0.0);
				}

				var nolpiat = 5 * (quotient < 0 ? -1.0 : quotient > 0 ? 1.0 : 0.0) *
					Math.pow(10, Math.floor(Math.log10(Math.abs(quotient))) - cExcelSignificantDigits);

				return new cNumber(Math.floor(quotient + nolpiat) * significance);
			}
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = floorHelper(a.getValue(), b.getValue())
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem;
				var b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = floorHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0;
				var b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = floorHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		if (arg0 instanceof cString || arg1 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		}

		return floorHelper(arg0.getValue(), arg1.getValue());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFLOOR_PRECISE() {
	}

	//***array-formula***
	cFLOOR_PRECISE.prototype = Object.create(cBaseFunction.prototype);
	cFLOOR_PRECISE.prototype.constructor = cFLOOR_PRECISE;
	cFLOOR_PRECISE.prototype.name = 'FLOOR.PRECISE';
	cFLOOR_PRECISE.prototype.argumentsMin = 1;
	cFLOOR_PRECISE.prototype.argumentsMax = 2;
	cFLOOR_PRECISE.prototype.isXLFN = true;
	cFLOOR_PRECISE.prototype.argumentsType = [argType.number, argType.number];
	cFLOOR_PRECISE.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1] ? argClone[1].tocNumber() : new cNumber(1);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function floorHelper(argArray) {
			var number = argArray[0];
			var significance = argArray[1];
			if (significance === 0 || number === 0) {
				return new cNumber(0.0);
			}

			var absSignificance = Math.abs(significance);
			var quotient = number / absSignificance;
			return new cNumber(Math.floor(quotient) * absSignificance);
		}

		return this._findArrayInNumberArguments(oArguments, floorHelper);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFLOOR_MATH() {
	}

	//***array-formula***
	cFLOOR_MATH.prototype = Object.create(cBaseFunction.prototype);
	cFLOOR_MATH.prototype.constructor = cFLOOR_MATH;
	cFLOOR_MATH.prototype.name = 'FLOOR.MATH';
	cFLOOR_MATH.prototype.argumentsMin = 1;
	cFLOOR_MATH.prototype.argumentsMax = 3;
	cFLOOR_MATH.prototype.isXLFN = true;
	cFLOOR_MATH.prototype.argumentsType = [argType.number, argType.number, argType.number];
	cFLOOR_MATH.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		if (!argClone[1]) {
			argClone[1] = argClone[0] > 0 ? new cNumber(1) : new cNumber(-1);
		}
		argClone[2] = argClone[2] ? argClone[2].tocNumber() : new cNumber(0);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function floor_math(argArray) {
			var number = argArray[0];
			var significance = argArray[1];
			var mod = argArray[2];

			if (significance === 0 || number === 0) {
				return new cNumber(0.0);
			}

			if (number * significance < 0.0) {
				significance = -significance;
			}

			if (mod === 0 && number < 0.0) {
				return new cNumber(Math.ceil(number / significance) * significance);
			} else {
				return new cNumber(Math.floor(number / significance) * significance);
			}
		}

		return this._findArrayInNumberArguments(oArguments, floor_math);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cGCD() {
	}

	cGCD.prototype = Object.create(cBaseFunction.prototype);
	cGCD.prototype.constructor = cGCD;
	cGCD.prototype.name = 'GCD';
	cGCD.prototype.argumentsMin = 1;
	cGCD.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cGCD.prototype.argumentsType = [[argType.any]];
	cGCD.prototype.Calculate = function (arg) {

		var _gcd = 0, argArr;

		function gcd(a, b) {
			var _a = parseInt(a), _b = parseInt(b);
			while (_b != 0)
				_b = _a % (_a = _b);
			return _a;
		}

		for (var i = 0; i < arg.length; i++) {
			var argI = arg[i];

			if (argI instanceof cArea || argI instanceof cArea3D) {
				argArr = argI.getValue();
				for (var j = 0; j < argArr.length; j++) {

					if (argArr[j] instanceof cError) {
						return argArr[j];
					}

					if (argArr[j] instanceof cString) {
						continue;
					}

					if (argArr[j] instanceof cBool) {
						argArr[j] = argArr[j].tocNumber();
					}

					if (argArr[j].getValue() < 0) {
						return new cError(cErrorType.not_numeric);
					}

					_gcd = gcd(_gcd, argArr[j].getValue());
				}
			} else if (argI instanceof cArray) {
				argArr = argI.tocNumber();

				if (argArr.foreach(function (arrElem) {

					if (arrElem instanceof cError) {
						_gcd = arrElem;
						return true;
					}

					if (arrElem instanceof cBool) {
						arrElem = arrElem.tocNumber();
					}

					if (arrElem instanceof cString) {
						return;
					}

					if (arrElem.getValue() < 0) {
						_gcd = new cError(cErrorType.not_numeric);
						return true;
					}
					_gcd = gcd(_gcd, arrElem.getValue());

				})) {
					return _gcd;
				}
			} else {
				argI = argI.tocNumber();

				if (argI.getValue() < 0) {
					return new cError(cErrorType.not_numeric);
				}

				if (argI instanceof cError) {
					return argI;
				}

				_gcd = gcd(_gcd, argI.getValue())
			}
		}

		return new cNumber(_gcd);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cINT() {
	}

	//***array-formula***
	cINT.prototype = Object.create(cBaseFunction.prototype);
	cINT.prototype.constructor = cINT;
	cINT.prototype.name = 'INT';
	cINT.prototype.argumentsMin = 1;
	cINT.prototype.argumentsMax = 1;
	cINT.prototype.inheritFormat = true;
	cINT.prototype.argumentsType = [argType.number];
	cINT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg0 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = new cNumber(Math.floor(elem.getValue()))
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			return new cNumber(Math.floor(arg0.getValue()))
		}

		return new cNumber(Math.floor(arg0.getValue()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	//TODO точная копия функции CEILING.PRECISE. зачем excel две одинаковые функции?
	function cISO_CEILING() {
	}

	cISO_CEILING.prototype = Object.create(cBaseFunction.prototype);
	cISO_CEILING.prototype.constructor = cISO_CEILING;
	cISO_CEILING.prototype.name = 'ISO.CEILING';
	cISO_CEILING.prototype.argumentsMin = 1;
	cISO_CEILING.prototype.argumentsMax = 2;
	cISO_CEILING.prototype.argumentsType = [argType.number, argType.number];
	//cISO_CEILING.prototype.isXLFN = true;
	cISO_CEILING.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();
		argClone[1] = argClone[1] ? argClone[1].tocNumber() : new cNumber(1);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function floorHelper(argArray) {
			var number = argArray[0];
			var significance = argArray[1];
			if (significance === 0 || number === 0) {
				return new cNumber(0.0);
			}

			var absSignificance = Math.abs(significance);
			var quotient = number / absSignificance;
			return new cNumber(Math.ceil(quotient) * absSignificance);
		}

		return this._findArrayInNumberArguments(oArguments, floorHelper);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLCM() {
	}

	//***array-formula***
	cLCM.prototype = Object.create(cBaseFunction.prototype);
	cLCM.prototype.constructor = cLCM;
	cLCM.prototype.name = 'LCM';
	cLCM.prototype.argumentsMin = 1;
	cLCM.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cLCM.prototype.argumentsType = [[argType.any]];
	cLCM.prototype.Calculate = function (arg) {

		var _lcm = 1, argArr;

		function gcd(a, b) {
			var _a = parseInt(a), _b = parseInt(b);
			while (_b != 0)
				_b = _a % (_a = _b);
			return _a;
		}

		function lcm(a, b) {
			return Math.abs(parseInt(a) * parseInt(b)) / gcd(a, b);
		}

		for (var i = 0; i < arg.length; i++) {
			var argI = arg[i];

			if (argI instanceof cArea || argI instanceof cArea3D) {
				argArr = argI.getValue();
				for (var j = 0; j < argArr.length; j++) {

					if (argArr[j] instanceof cError) {
						return argArr[j];
					}

					if (argArr[j] instanceof cString) {
						continue;
					}

					if (argArr[j] instanceof cBool) {
						argArr[j] = argArr[j].tocNumber();
					}

					if (argArr[j].getValue() <= 0) {
						return new cError(cErrorType.not_numeric);
					}

					_lcm = lcm(_lcm, argArr[j].getValue());
				}
			} else if (argI instanceof cArray) {
				argArr = argI.tocNumber();

				if (argArr.foreach(function (arrElem) {

					if (arrElem instanceof cError) {
						_lcm = arrElem;
						return true;
					}

					if (arrElem instanceof cBool) {
						arrElem = arrElem.tocNumber();
					}

					if (arrElem instanceof cString) {
						return;
					}

					if (arrElem.getValue() <= 0) {
						_lcm = new cError(cErrorType.not_numeric);
						return true;
					}
					_lcm = lcm(_lcm, arrElem.getValue());

				})) {
					return _lcm;
				}
			} else {
				argI = argI.tocNumber();

				if (argI.getValue() <= 0) {
					return new cError(cErrorType.not_numeric);
				}

				if (argI instanceof cError) {
					return argI;
				}

				_lcm = lcm(_lcm, argI.getValue())
			}
		}

		return new cNumber(_lcm);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLN() {
	}

	//***array-formula***
	cLN.prototype = Object.create(cBaseFunction.prototype);
	cLN.prototype.constructor = cLN;
	cLN.prototype.name = 'LN';
	cLN.prototype.argumentsMin = 1;
	cLN.prototype.argumentsMax = 1;
	cLN.prototype.argumentsType = [argType.number];
	cLN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg0 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					if (elem.getValue() <= 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						this.array[r][c] = new cNumber(Math.log(elem.getValue()));
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			if (arg0.getValue() <= 0) {
				return new cError(cErrorType.not_numeric);
			} else {
				return new cNumber(Math.log(arg0.getValue()));
			}
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLOG() {
	}

	cLOG.prototype = Object.create(cBaseFunction.prototype);
	cLOG.prototype.constructor = cLOG;
	cLOG.prototype.name = 'LOG';
	cLOG.prototype.argumentsMin = 1;
	cLOG.prototype.argumentsMax = 2;
	cLOG.prototype.argumentsType = [argType.number, argType.number];
	cLOG.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(10);
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						if (1 === b.getValue()) {
							return new cError(cErrorType.division_by_zero);
						}

						this.array[r][c] = new cNumber(Math.log(a.getValue()) / Math.log(b.getValue()));
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem, b = arg1 ? arg1 : new cNumber(10);
				if (a instanceof cNumber && b instanceof cNumber) {

					if (a.getValue() <= 0 || a.getValue() <= 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					}

					if (1 === b.getValue()) {
						return new cError(cErrorType.division_by_zero);
					}

					this.array[r][c] = new cNumber(Math.log(a.getValue()) / Math.log(b.getValue()));
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0, b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {

					if (a.getValue() <= 0 || a.getValue() <= 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					}

					if (1 === b.getValue()) {
						return new cError(cErrorType.division_by_zero);
					}

					this.array[r][c] = new cNumber(Math.log(a.getValue()) / Math.log(b.getValue()));
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		if (!(arg0 instanceof cNumber) || (arg1 && !(arg0 instanceof cNumber))) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg0.getValue() <= 0 || (arg1 && arg1.getValue() <= 0)) {
			return new cError(cErrorType.not_numeric);
		}

		if (1 === arg1.getValue()) {
			return new cError(cErrorType.division_by_zero);
		}

		return new cNumber(Math.log(arg0.getValue()) / Math.log(arg1.getValue()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLOG10() {
	}

	//***array-formula***
	cLOG10.prototype = Object.create(cBaseFunction.prototype);
	cLOG10.prototype.constructor = cLOG10;
	cLOG10.prototype.name = 'LOG10';
	cLOG10.prototype.argumentsMin = 1;
	cLOG10.prototype.argumentsMax = 1;
	cLOG10.prototype.argumentsType = [argType.number];
	cLOG10.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg0 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					if (elem.getValue() <= 0) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						this.array[r][c] = new cNumber(Math.log10(elem.getValue()));
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			if (arg0.getValue() <= 0) {
				return new cError(cErrorType.not_numeric);
			} else {
				return new cNumber(Math.log10(arg0.getValue()));
			}
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMDETERM() {
	}

	//***array-formula***
	cMDETERM.prototype = Object.create(cBaseFunction.prototype);
	cMDETERM.prototype.constructor = cMDETERM;
	cMDETERM.prototype.name = 'MDETERM';
	cMDETERM.prototype.argumentsMin = 1;
	cMDETERM.prototype.argumentsMax = 1;
	cMDETERM.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMDETERM.prototype.arrayIndexes = {0: 1};
	cMDETERM.prototype.argumentsType = [argType.array];
	cMDETERM.prototype.Calculate = function (arg) {

		function determ(A) {
			var N = A.length, denom = 1, exchanges = 0, i, j;

			for (i = 0; i < N; i++) {
				for (j = 0; j < A[i].length; j++) {
					if (A[i][j] instanceof cEmpty || A[i][j] instanceof cString) {
						return NaN;
					}
				}
			}

			for (i = 0; i < N - 1; i++) {
				var maxN = i, maxValue = Math.abs(A[i][i] instanceof cEmpty ? NaN : A[i][i]);
				for (j = i + 1; j < N; j++) {
					var value = Math.abs(A[j][i] instanceof cEmpty ? NaN : A[j][i]);
					if (value > maxValue) {
						maxN = j;
						maxValue = value;
					}
				}
				if (maxN > i) {
					var temp = A[i];
					A[i] = A[maxN];
					A[maxN] = temp;
					exchanges++;
				} else {
					if (maxValue == 0) {
						return maxValue;
					}
				}
				var value1 = A[i][i] instanceof cEmpty ? NaN : A[i][i];
				for (j = i + 1; j < N; j++) {
					var value2 = A[j][i] instanceof cEmpty ? NaN : A[j][i];
					A[j][i] = 0;
					for (var k = i + 1; k < N; k++) {
						A[j][k] = (A[j][k] * value1 - A[i][k] * value2) / denom;
					}
				}
				denom = value1;
			}

			if (exchanges % 2) {
				return -A[N - 1][N - 1];
			} else {
				return A[N - 1][N - 1];
			}
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArray) {
			arg0 = arg0.getMatrix();
		} else {
			return new cError(cErrorType.not_available);
		}

		if (arg0[0].length != arg0.length) {
			return new cError(cErrorType.wrong_value_type);
		}

		arg0 = determ(arg0);

		if (!isNaN(arg0)) {
			return new cNumber(arg0);
		} else {
			return new cError(cErrorType.not_available);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMINVERSE() {
	}

	//***array-formula***
	cMINVERSE.prototype = Object.create(cBaseFunction.prototype);
	cMINVERSE.prototype.constructor = cMINVERSE;
	cMINVERSE.prototype.name = 'MINVERSE';
	cMINVERSE.prototype.argumentsMin = 1;
	cMINVERSE.prototype.argumentsMax = 1;
	cMINVERSE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMINVERSE.prototype.arrayIndexes = {0: 1};
	cMINVERSE.prototype.argumentsType = [argType.array];
	cMINVERSE.prototype.Calculate = function (arg) {
		function Determinant(A) {
			let N = A.length, B = [], denom = 1, exchanges = 0, i, j;

			for (i = 0; i < N; ++i) {
				B[i] = [];
				for (j = 0; j < N; ++j) {
					B[i][j] = A[i][j];
				}
			}

			for (i = 0; i < N - 1; ++i) {
				let maxN = i, maxValue = Math.abs(B[i][i]);
				for (j = i + 1; j < N; ++j) {
					let value = Math.abs(B[j][i]);
					if (value > maxValue) {
						maxN = j;
						maxValue = value;
					}
				}
				if (maxN > i) {
					let temp = B[i];
					B[i] = B[maxN];
					B[maxN] = temp;
					++exchanges;
				} else {
					if (maxValue == 0) {
						return maxValue;
					}
				}
				let value1 = B[i][i];
				for (j = i + 1; j < N; ++j) {
					let value2 = B[j][i];
					B[j][i] = 0;
					for (let k = i + 1; k < N; ++k) {
						B[j][k] = (B[j][k] * value1 - B[i][k] * value2) / denom;
					}
				}
				denom = value1;
			}
			if (exchanges % 2) {
				return -B[N - 1][N - 1];
			} else {
				return B[N - 1][N - 1];
			}
		}

		function MatrixCofactor(i, j, __A) {        //Алгебраическое дополнение матрицы
			let N = __A.length, sign = ((i + j) % 2 == 0) ? 1 : -1;

			for (let m = 0; m < N; m++) {
				for (let n = j + 1; n < N; n++) {
					__A[m][n - 1] = __A[m][n];
				}
				__A[m].length--;
			}
			for (let k = (i + 1); k < N; k++) {
				__A[k - 1] = __A[k];
			}
			__A.length--;

			return sign * Determinant(__A);
		}

		function AdjugateMatrix(_A) {             //Союзная (присоединённая) матрица к A. (матрица adj(A), составленная из алгебраических дополнений A).
			let N = _A.length, B = [], adjA = [];

			for (let i = 0; i < N; i++) {
				adjA[i] = [];
				for (let j = 0; j < N; j++) {
					for (let m = 0; m < N; m++) {
						B[m] = [];
						for (let n = 0; n < N; n++) {
							B[m][n] = _A[m][n];
						}
					}
					adjA[i][j] = MatrixCofactor(j, i, B);
				}
			}

			return adjA;
		}

		function InverseMatrix(A) {
			let i, j;
			for (i = 0; i < A.length; i++) {
				for (j = 0; j < A[i].length; j++) {
					if (cElementType.empty === A[i][j].type || cElementType.string === A[i][j].type) {
						return new cError(cErrorType.wrong_value_type);
					} else {
						A[i][j] = A[i][j].getValue();
					}
				}
			}

			let detA = Determinant(A), invertA, res;

			if (detA != 0) {
				invertA = AdjugateMatrix(A);
				let datA = 1 / detA;
				for (i = 0; i < invertA.length; i++) {
					for (j = 0; j < invertA[i].length; j++) {
						invertA[i][j] = new cNumber(datA * invertA[i][j]);
					}
				}
				res = new cArray();
				res.fillFromArray(invertA);
			} else {
				res = new cError(cErrorType.not_numeric);
			}

			return res;
		}

		function _getArrayCopy(arr) {
			var newArray = [];
			for (var i = 0; i < arr.length; i++) {
				newArray[i] = arr[i].slice();
			}
			return newArray
		}

		let arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type || cElementType.array === arg0.type) {
			if (arg0.isOneElement()) {
				return new cNumber(1 / arg0.getFirstElement());
			}
			arg0 = arg0.getMatrix();
			//TODO при мерже релиза, перейти на функцию getArrayCopy/ добавить параметр getMatrix для генерации копии
			arg0 = _getArrayCopy(arg0);
		} else if (cElementType.number === arg0.type) {
			return new cNumber(1 / arg0);
		} else if (cElementType.cell === arg0.type) {
			let arg0Val = arg0.getValue();
			if (cElementType.number === arg0Val.type) {
				return new cNumber(1 / arg0Val);
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg0[0].length != arg0.length) {
			return new cError(cErrorType.wrong_value_type);
		}

		return InverseMatrix(arg0);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMMULT() {
	}

	cMMULT.prototype = Object.create(cBaseFunction.prototype);
	cMMULT.prototype.constructor = cMMULT;
	cMMULT.prototype.name = 'MMULT';
	cMMULT.prototype.argumentsMin = 2;
	cMMULT.prototype.argumentsMax = 2;
	cMMULT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMMULT.prototype.arrayIndexes = {0: 1, 1: 1};
	cMMULT.prototype.argumentsType = [argType.array, argType.array];
	cMMULT.prototype.Calculate = function (arg) {

		function mult(A, B) {
			if ( !((B[0] && A.length === B[0].length) || (A[0] && B.length === A[0].length)) || 
				(A[0].length === 1 && (B[0].length !== 1 && A.length !== B[0].length)) ||
				(A.length === 1 && (B.length !== 1 && A[0].length !== B.length)) ) {
				return new cError(cErrorType.wrong_value_type);
			}

			let i, j;
			for (i = 0; i < A.length; i++) {
				for (j = 0; j < A[i].length; j++) {
					if (A[i][j].type === cElementType.empty || A[i][j].type === cElementType.string) {
						return new cError(cErrorType.wrong_value_type);
					}
				}
			}
			for (i = 0; i < B.length; i++) {
				for (j = 0; j < B[i].length; j++) {
					if (B[i][j].type === cElementType.empty || B[i][j].type === cElementType.string) {
						return new cError(cErrorType.wrong_value_type);
					}
				}
			}


			let C = new Array(A.length);
			for (i = 0; i < A.length; i++) {
				C[i] = new Array(B[0].length);
				for (j = 0; j < B[0].length; j++) {
					C[i][j] = 0;
					for (let k = 0; k < B.length; k++) {
						let fMatrixElem = A[i] && A[i][k] ? A[i][k].getValue() : false;
						let sMatrixElem = B[k] && B[k][j] ? B[k][j].getValue() : false;
						if (fMatrixElem === false || sMatrixElem === false) {
							return new cError(cErrorType.wrong_value_type);
						}
						C[i][j] += fMatrixElem * sMatrixElem;
					}
					C[i][j] = new cNumber(C[i][j]);
				}
			}
			let res = new cArray();
			res.fillFromArray(C);
			return res;
		}

		let arg0 = arg[0], arg1 = arg[1];
		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.array || arg0.type === cElementType.cell || arg0.type === cElementType.cell3D) {
			arg0 = arg0.getMatrix();
		} else if (arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.getMatrix()[0];
		} else if (arg0.type === cElementType.number) {
			arg0 = arg0.toArray();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.array || arg1.type === cElementType.cell || arg1.type === cElementType.cell3D) {
			arg1 = arg1.getMatrix();
		} else if (arg1.type === cElementType.cellsRange3D) {
			arg1 = arg1.getMatrix()[0];
		} else if (arg1.type === cElementType.number) {
			arg1 = arg1.toArray();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		return mult(arg0, arg1);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMOD() {
	}

	cMOD.prototype = Object.create(cBaseFunction.prototype);
	cMOD.prototype.constructor = cMOD;
	cMOD.prototype.name = 'MOD';
	cMOD.prototype.argumentsMin = 2;
	cMOD.prototype.argumentsMax = 2;
	cMOD.prototype.argumentsType = [argType.number, argType.number];
	cMOD.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		var calc = function (n, d) {
			if (d === 0) {
				return new cError(cErrorType.division_by_zero);
			}
			return new cNumber(n - d * Math.floor(n / d));
		};

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = calc(a.getValue(), b.getValue());
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem, b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {

					this.array[r][c] = calc(a.getValue(), b.getValue());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0, b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = calc(a.getValue(), b.getValue());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		if (!(arg0 instanceof cNumber) || (arg1 && !(arg0 instanceof cNumber))) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg1.getValue() == 0) {
			return new cError(cErrorType.division_by_zero);
		}

		return calc(arg0.getValue(), arg1.getValue());

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMROUND() {
	}

	cMROUND.prototype = Object.create(cBaseFunction.prototype);
	cMROUND.prototype.constructor = cMROUND;
	cMROUND.prototype.name = 'MROUND';
	cMROUND.prototype.argumentsMin = 2;
	cMROUND.prototype.argumentsMax = 2;
	cMROUND.prototype.argumentsType = [argType.any, argType.any];
	cMROUND.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cMROUND.prototype.Calculate = function (arg) {

		var multiple;

		function mroundHelper(num) {
			var multiplier = Math.pow(10, Math.floor(Math.log10(Math.abs(num))) - cExcelSignificantDigits + 1);
			var nolpiat = 0.5 * (num > 0 ? 1 : num < 0 ? -1 : 0) * multiplier;
			var y = (num + nolpiat) / multiplier;
			y = y / Math.abs(y) * Math.floor(Math.abs(y));
			var x = y * multiplier / multiple;

			// var x = number / multiple;
			nolpiat =
				5 * (x / Math.abs(x)) * Math.pow(10, Math.floor(Math.log10(Math.abs(x))) - cExcelSignificantDigits);
			x = x + nolpiat;
			x = x | x;

			return x * multiple;
		}

		function f(a, b, r, c) {
			if (a instanceof cNumber && b instanceof cNumber) {
				if (a.getValue() == 0) {
					this.array[r][c] = new cNumber(0);
				} else if (a.getValue() < 0 && b.getValue() > 0 || arg0.getValue() > 0 && b.getValue() < 0) {
					this.array[r][c] = new cError(cErrorType.not_numeric);
				} else {
					multiple = b.getValue();
					this.array[r][c] = new cNumber(mroundHelper(a.getValue() + b.getValue() / 2))
				}
			} else {
				this.array[r][c] = new cError(cErrorType.wrong_value_type);
			}
		}

		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg0 instanceof cString || arg1 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					f.call(this, elem, arg1.getElementRowCol(r, c), r, c)
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				f.call(this, elem, arg1, r, c);
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				f.call(this, arg0, elem, r, c)
			});
			return arg1;
		}

		if (arg1.getValue() == 0) {
			return new cNumber(0);
		}

		if (arg0.getValue() < 0 && arg1.getValue() > 0 || arg0.getValue() > 0 && arg1.getValue() < 0) {
			return new cError(cErrorType.not_numeric);
		}

		multiple = arg1.getValue();
		return new cNumber(mroundHelper(arg0.getValue() + arg1.getValue() / 2));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMULTINOMIAL() {
	}

	//***array-formula***
	cMULTINOMIAL.prototype = Object.create(cBaseFunction.prototype);
	cMULTINOMIAL.prototype.constructor = cMULTINOMIAL;
	cMULTINOMIAL.prototype.name = 'MULTINOMIAL';
	cMULTINOMIAL.prototype.argumentsMin = 1;
	cMULTINOMIAL.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cMULTINOMIAL.prototype.argumentsType = [[argType.any]];
	cMULTINOMIAL.prototype.Calculate = function (arg) {
		var arg0 = new cNumber(0), fact = 1;

		for (var i = 0; i < arg.length; i++) {
			if (arg[i] instanceof cArea || arg[i] instanceof cArea3D) {
				var _arrVal = arg[i].getValue();
				for (var j = 0; j < _arrVal.length; j++) {
					if (_arrVal[j] instanceof cNumber) {
						if (_arrVal[j].getValue() < 0) {
							return new cError(cErrorType.not_numeric);
						}
						arg0 = _func[arg0.type][_arrVal[j].type](arg0, _arrVal[j], "+");
						fact *= Math.fact(_arrVal[j].getValue());
					} else if (_arrVal[j] instanceof cError) {
						return _arrVal[j];
					} else {
						return new cError(cErrorType.wrong_value_type);
					}
				}
			} else if (arg[i] instanceof cArray) {
				if (arg[i].foreach(function (arrElem) {
					if (arrElem instanceof cNumber) {
						if (arrElem.getValue() < 0) {
							return true;
						}

						arg0 = _func[arg0.type][arrElem.type](arg0, arrElem, "+");
						fact *= Math.fact(arrElem.getValue());
					} else {
						return true;
					}
				})) {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (arg[i] instanceof cRef || arg[i] instanceof cRef3D) {
				var _arg = arg[i].getValue();

				if (_arg.getValue() < 0) {
					return new cError(cErrorType.not_numeric);
				}

				if (_arg instanceof cNumber) {
					if (_arg.getValue() < 0) {
						return new cError(cError.not_numeric);
					}
					arg0 = _func[arg0.type][_arg.type](arg0, _arg, "+");
					fact *= Math.fact(_arg.getValue());
				} else if (_arg instanceof cError) {
					return _arg;
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			} else if (arg[i] instanceof cNumber) {

				if (arg[i].getValue() < 0) {
					return new cError(cErrorType.not_numeric);
				}

				arg0 = _func[arg0.type][arg[i].type](arg0, arg[i], "+");
				fact *= Math.fact(arg[i].getValue());
			} else if (arg[i] instanceof cError) {
				return arg[i];
			} else {
				return new cError(cErrorType.wrong_value_type);
			}

			if (arg0 instanceof cError) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (arg0.getValue() > 170) {
			return new cError(cErrorType.wrong_value_type);
		}

		return new cNumber(Math.fact(arg0.getValue()) / fact);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMUNIT() {
	}

	cMUNIT.prototype = Object.create(cBaseFunction.prototype);
	cMUNIT.prototype.constructor = cMUNIT;
	cMUNIT.prototype.name = "MUNIT";
	cMUNIT.prototype.argumentsMin = 1;
	cMUNIT.prototype.argumentsMax = 1;
	cMUNIT.prototype.isXLFN = true;
	cMUNIT.prototype.argumentsType = [argType.number];
	cMUNIT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cMUNIT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cMUNIT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			//по обработке массива есть вопросы
			//в случае если аргуметт функции должен вернуть массив - берётся первый элемента массива
			//в случае формулы массива возвращается результат от каждого значения массива
			//реализовываю второй вариант
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = parseInt(elem.getValue()) > 0 ? new cNumber(1) : new cError(cErrorType.wrong_value_type);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			var num = parseInt(arg0);
			if (num <= 0) {
				return new cError(cErrorType.wrong_value_type);
			}
			var _arr = [];
			for (var i = 0; i < num; i++) {
				for (var j = 0; j < num; j++) {
					if (!_arr[i]) {
						_arr[i] = [];
					}
					_arr[i][j] = i === j ? new cNumber(1) : new cNumber(0);
				}
			}
			var res = new cArray();
			res.fillFromArray(_arr);
			return res;
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cODD() {
	}

	//***array-formula***
	cODD.prototype = Object.create(cBaseFunction.prototype);
	cODD.prototype.constructor = cODD;
	cODD.prototype.name = 'ODD';
	cODD.prototype.argumentsMin = 1;
	cODD.prototype.argumentsMax = 1;
	cODD.prototype.argumentsType = [argType.number];
	cODD.prototype.Calculate = function (arg) {

		function oddHelper(arg) {
			var arg0 = arg.getValue();
			if (arg0 >= 0) {
				arg0 = Math.ceil(arg0);
				if ((arg0 & 1) == 1) {
					return new cNumber(arg0);
				} else {
					return new cNumber(arg0 + 1);
				}
			} else {
				arg0 = Math.floor(arg0);
				if ((arg0 & 1) == 1) {
					return new cNumber(arg0);
				} else {
					return new cNumber(arg0 - 1);
				}
			}
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = oddHelper(elem);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg0 instanceof cNumber) {
			return oddHelper(arg0);
		}
		return new cError(cErrorType.wrong_value_type);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cPI() {
	}

	//***array-formula***
	cPI.prototype = Object.create(cBaseFunction.prototype);
	cPI.prototype.constructor = cPI;
	cPI.prototype.name = 'PI';
	cPI.prototype.argumentsMax = 0;
	cPI.prototype.argumentsType = null;
	cPI.prototype.Calculate = function () {
		return new cNumber(Math.PI);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cPOWER() {
	}

	cPOWER.prototype = Object.create(cBaseFunction.prototype);
	cPOWER.prototype.constructor = cPOWER;
	cPOWER.prototype.name = 'POWER';
	cPOWER.prototype.argumentsMin = 2;
	cPOWER.prototype.argumentsMax = 2;
	cPOWER.prototype.argumentsType = [argType.number, argType.number];
	cPOWER.prototype.Calculate = function (arg) {

		function powerHelper(a, b) {
			if (a === 0 && b < 0) {
				return new cError(cErrorType.division_by_zero);
			}

			if (a >= 0 || Math.round(b) === b) {
				return new cNumber(Math.pow(a, b));
			} else {
				let r = -1 * Math.pow(-a, b);
				if (Math.round(Math.pow(r, 1 / b)) === Math.round(a)) {
					return new cNumber(r);
				} else {
					return new cError(cErrorType.not_numeric);
				}
			}
		}

		function f(a, b, r, c) {
			if (cElementType.number === a.type && cElementType.number === b.type) {
				this.array[r][c] = powerHelper(a.getValue(), b.getValue());
			} else {
				if (cElementType.error === a.type) {
					this.array[r][c] = a;
				} else if (cElementType.error === b.type) {
					this.array[r][c] = b;
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			}
		}

		let arg0 = arg[0], arg1 = arg[1], t = this;
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.getFullArray();
			if (arg0.isOneElement()) {
				arg0 = arg0.getFirstElement();
			}
		}
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.getFullArray();
			if (arg1.isOneElement()) {
				arg1 = arg1.getFirstElement();
			}
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.array === arg0.type && cElementType.array === arg1.type) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				let arg0Dimensions = arg0.getDimensions(),
					arg1Dimensions = arg1.getDimensions();

				let resArr = new cArray();
				let resDimensions = {col: Math.max(arg0Dimensions.col, arg1Dimensions.col), row: Math.max(arg0Dimensions.row, arg1Dimensions.row)};

				for (let resRow = 0; resRow < resDimensions.row; resRow++) {
					resArr.addRow();
					for (let resCol = 0; resCol < resDimensions.col; resCol++) {
						let base, power;
						/* find base */
						if (arg0Dimensions.row === 1 && arg0Dimensions.col > resCol) {
							base = arg0.getElementRowCol(0, resCol);
						} else if (arg0Dimensions.col === 1 && arg0Dimensions.row > resRow) {
							base = arg0.getElementRowCol(resRow, 0);
						} else {
							if (arg0Dimensions.row - 1 < resRow || arg0Dimensions.col - 1 < resCol) {
								base = new cError(cErrorType.not_available);
							}
							base = base ? base : arg0.getElementRowCol(resRow, resCol);
						}

						/* find power */
						if (arg1Dimensions.row === 1 && arg1Dimensions.col > resCol) {
							power = arg1.getElementRowCol(0, resCol);
						} else if (arg1Dimensions.col === 1 && arg1Dimensions.row > resRow) {
							power = arg1.getElementRowCol(resRow, 0);
						} else {
							if (arg1Dimensions.row - 1 < resRow || arg1Dimensions.col - 1 < resCol) {
								power = new cError(cErrorType.not_available);
							}
							power = power ? power : arg1.getElementRowCol(resRow, resCol);
						}

						f.call(resArr, base, power, resRow, resCol);
					}
				}

				if (resArr) {
					resArr.recalculate();
					return resArr;
				}
			} else {
				arg0.foreach(function (elem, r, c) {
					f.call(this, elem, arg1.getElementRowCol(r, c), r, c);
				});
				return arg0;
			}
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				f.call(this, elem, arg1, r, c)
			});
			return arg0;
		} else if (cElementType.array === arg1.type) {
			arg1.foreach(function (elem, r, c) {
				f.call(this, arg0, elem, r, c);
			});
			return arg1;
		}

		if (!(cElementType.number === arg0.type) || (arg1 && !(cElementType.number === arg0.type))) {
			return new cError(cErrorType.wrong_value_type);
		}

		return powerHelper(arg0.getValue(), arg1.getValue());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cPRODUCT() {
	}

	//***array-formula***
	cPRODUCT.prototype = Object.create(cBaseFunction.prototype);
	cPRODUCT.prototype.constructor = cPRODUCT;
	cPRODUCT.prototype.name = 'PRODUCT';
	cPRODUCT.prototype.argumentsMin = 1;
	cPRODUCT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cPRODUCT.prototype.argumentsType = [[argType.number]];
	cPRODUCT.prototype.Calculate = function (arg) {
		var element, arg0 = new cNumber(1);
		for (var i = 0; i < arg.length; i++) {
			element = arg[i];
			if (cElementType.cellsRange === element.type || cElementType.cellsRange3D === element.type) {
				var _arrVal = element.getValue(this.checkExclude, this.excludeHiddenRows, this.excludeErrorsVal,
					this.excludeNestedStAg);
				for (var j = 0; j < _arrVal.length; j++) {
					if (_arrVal[j].type !== cElementType.string) {
						arg0 = _func[arg0.type][_arrVal[j].type](arg0, _arrVal[j], "*");
						if (cElementType.error === arg0.type) {
							return arg0;
						}
					}
				}
			} else if (cElementType.cell === element.type || cElementType.cell3D === element.type) {
				if (!this.checkExclude || !element.isHidden(this.excludeHiddenRows)) {
					var _arg = element.getValue();
					if (_arg.type !== cElementType.string) {
						arg0 = _func[arg0.type][_arg.type](arg0, _arg, "*");
					}
				}
			} else if (cElementType.array === element.type) {
				element.foreach(function (elem) {
					if (cElementType.string === elem.type || cElementType.bool === elem.type ||
						cElementType.empty === elem.type) {
						return;
					}

					arg0 = _func[arg0.type][elem.type](arg0, elem, "*");
				})
			} else {
				arg0 = _func[arg0.type][element.type](arg0, element, "*");
			}
			if (cElementType.error === arg0.type) {
				return arg0;
			}

		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cQUOTIENT() {
	}

	//***array-formula***
	cQUOTIENT.prototype = Object.create(cBaseFunction.prototype);
	cQUOTIENT.prototype.constructor = cQUOTIENT;
	cQUOTIENT.prototype.name = 'QUOTIENT';
	cQUOTIENT.prototype.argumentsMin = 2;
	cQUOTIENT.prototype.argumentsMax = 2;
	cQUOTIENT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cQUOTIENT.prototype.argumentsType = [argType.any, argType.any];
	cQUOTIENT.prototype.Calculate = function (arg) {

		function quotient(a, b) {
			if (b.getValue() != 0) {
				return new cNumber(parseInt(a.getValue() / b.getValue()));
			} else {
				return new cError(cErrorType.division_by_zero);
			}
		}

		function f(a, b, r, c) {
			if (a instanceof cNumber && b instanceof cNumber) {
				this.array[r][c] = quotient(a, b);
			} else {
				this.array[r][c] = new cError(cErrorType.wrong_value_type);
			}
		}

		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg1 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					f.call(this, elem, arg1.getElementRowCol(r, c), r, c);
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				f.call(this, elem, arg1, r, c)
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				f.call(this, arg0, elem, r, c);
			});
			return arg1;
		}

		if (!(arg0 instanceof cNumber) || (arg1 && !(arg0 instanceof cNumber))) {
			return new cError(cErrorType.wrong_value_type);
		}


		return quotient(arg0, arg1);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cRADIANS() {
	}

	//***array-formula***
	cRADIANS.prototype = Object.create(cBaseFunction.prototype);
	cRADIANS.prototype.constructor = cRADIANS;
	cRADIANS.prototype.name = 'RADIANS';
	cRADIANS.prototype.argumentsMin = 1;
	cRADIANS.prototype.argumentsMax = 1;
	cRADIANS.prototype.argumentsType = [argType.number];
	cRADIANS.prototype.Calculate = function (arg) {

		function radiansHelper(ang) {
			return ang * Math.PI / 180
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();

		if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					this.array[r][c] = new cNumber(radiansHelper(elem.getValue()));
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			return (arg0 instanceof cError ? arg0 : new cNumber(radiansHelper(arg0.getValue())));
		}

		return arg0;

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cRAND() {
	}

	//***array-formula***
	cRAND.prototype = Object.create(cBaseFunction.prototype);
	cRAND.prototype.constructor = cRAND;
	cRAND.prototype.name = 'RAND';
	cRAND.prototype.argumentsMax = 0;
	cRAND.prototype.ca = true;
	cRAND.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.area_to_ref;
	cRAND.prototype.argumentsType = null;
	cRAND.prototype.Calculate = function () {
		return new cNumber(Math.random());
	};

	function cRANDARRAY() {
	}

	//***array-formula***
	cRANDARRAY.prototype = Object.create(cBaseFunction.prototype);
	cRANDARRAY.prototype.constructor = cRANDARRAY;
	cRANDARRAY.prototype.name = 'RANDARRAY';
	cRANDARRAY.prototype.argumentsMin = 0;
	cRANDARRAY.prototype.argumentsMax = 5;
	cRANDARRAY.prototype.ca = true;
	cRANDARRAY.prototype.isXLFN = true;
	cRANDARRAY.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cRANDARRAY.prototype.argumentsType = [argType.number, argType.number, argType.number, argType.number, argType.number];
	cRANDARRAY.prototype.Calculate = function (arg) {
		//var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = arg;

		//если какой-то из аргументов массив - обрабатываю здесь
		//если обрабатывать выше и проходиться по массиву, то данная функция всегда будет возвращать массив
		//а нам нужно только значение с индексом 0,0 у возвращаемого массива

		var i, j;
		var matrixRowCount;
		var matrixColCount;
		for (i = 0; i < argClone.length; i++) {
			if (argClone[i].type === cElementType.empty && i !== 4) {
				if (i !== 2) {
					argClone[i] = new cNumber(1);
				} else {
					argClone[i] = new cNumber(0);
				}
			} else if (argClone[i].type === cElementType.array || cElementType.cellsRange === argClone[i].type || cElementType.cellsRange3D === argClone[i].type) {
				argClone[i] = argClone[i].getMatrix();
				if (cElementType.cellsRange3D === argClone[i].type) {
					argClone[i] = argClone[i][0];
				}
				if (matrixRowCount === undefined || matrixRowCount > argClone[i].length) {
					matrixRowCount = argClone[i].length;
				}
				if (matrixColCount === undefined || matrixColCount > argClone[i][0].length) {
					matrixColCount = argClone[i][0].length;
				}
			} else if (i !== 4) {
				argClone[i] = argClone[i].tocNumber()
			}
		}

		if (argClone[4]) {
			if (matrixRowCount === undefined) {
				if (argClone[4].type === cElementType.cell || argClone[4].type === cElementType.cell3D) {
					argClone[4] = argClone[4].getValue();
				}
				if (argClone[4].type === cElementType.string) {
					return new cError(cErrorType.wrong_value_type);
				} else if (argClone[4].type === cElementType.error) {
					return argClone[4];
				}

				argClone[4] = argClone[4].tocBool();
			}
		}

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function randBetween(a, b, _wholeNumber) {
			if (_wholeNumber) {
				return new cNumber(Math.floor(Math.random() * (b - a + 1)) + a);
			} else {
				return new cNumber(Math.random() * (max - min) + min);
			}
		}

		var rowCount;
		var colCount;
		var min;
		var max;
		var wholeNumber;
		var _array;
		if (matrixRowCount !== undefined) {
			_array = new cArray();
			for (i = 0; i < matrixRowCount; i++) {
				_array.addRow();
				for (j = 0; j < matrixColCount; j++) {
					if (argClone[0] && argClone[0][i] && argClone[0][i][j] && argClone[0][i][j].type === cElementType.error) {
						_array.addElement(argClone[0][i][j]);
						continue;
					} else if (argClone[1] && argClone[1][i] && argClone[1][i][j] && argClone[1][i][j].type === cElementType.error) {
						_array.addElement(argClone[1][i][j]);
						continue;
					}

					min = 0;
					if (argClone[2] && argClone[2][i]) {
						if (argClone[2][i][j] && argClone[2][i][j].type === cElementType.error) {
							_array.addElement(argClone[2][i][j]);
							continue;
						}
						min = argClone[2][i][j].getValue();
					} else if (argClone[2]) {
						min = argClone[2].getValue();
					}
					max = 1;
					if (argClone[3] && argClone[3][i]) {
						if (argClone[3][i][j] && argClone[3][i][j].type === cElementType.error) {
							_array.addElement(argClone[3][i][j]);
							continue;
						}
						max = argClone[3][i][j].getValue();
					} else if (argClone[3]) {
						max = argClone[3].getValue();
					}
					wholeNumber = false;
					if (argClone[4] && argClone[4][i]) {
						if (argClone[4][i][j] && (argClone[4][i][j].type === cElementType.error || argClone[4][i][j].type === cElementType.string)) {
							if (argClone[4][i][j].type === cElementType.string) {
								_array.addElement(new cError(cErrorType.wrong_value_type))
							} else {
								_array.addElement(argClone[4][i][j]);
							}
							continue;
						}
						wholeNumber = argClone[4][i][j].getValue();
					} else if (argClone[4]) {
						if (argClone[4][i][j].type === cElementType.string) {
							_array.addElement(new cError(cErrorType.wrong_value_type))
							continue;
						}
						wholeNumber = argClone[4].toBool();
					}

					_array.addElement(randBetween(min, max, wholeNumber));
				}
			}
		} else {
			rowCount = argClone[0] ? parseInt(argClone[0].getValue()) : 1;
			colCount = argClone[1] ? parseInt(argClone[1].getValue()) : 1;
			min = argClone[2] ? argClone[2].getValue() : 0;
			max = argClone[3] ? argClone[3].getValue() : 1;
			wholeNumber = argClone[4] ? argClone[4].toBool() : false;

			if (min > max) {
				return new cError(cErrorType.wrong_value_type);
			}
			if (wholeNumber && (!Number.isInteger(min) || !Number.isInteger(max))) {
				return new cError(cErrorType.wrong_value_type);
			}

			if (rowCount <= 0 || colCount <= 0) {
				return new cError(cErrorType.wrong_value_type);
			}

			_array = new cArray();
			for (i = 0; i < rowCount; i++) {
				_array.addRow();
				for (j = 0; j < colCount; j++) {
					_array.addElement(randBetween(min, max, wholeNumber));
				}
			}
		}

		return _array;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cRANDBETWEEN() {
	}

	//***array-formula***
	cRANDBETWEEN.prototype = Object.create(cBaseFunction.prototype);
	cRANDBETWEEN.prototype.constructor = cRANDBETWEEN;
	cRANDBETWEEN.prototype.name = 'RANDBETWEEN';
	cRANDBETWEEN.prototype.argumentsMin = 2;
	cRANDBETWEEN.prototype.argumentsMax = 2;
	cRANDBETWEEN.prototype.ca = true;
	cRANDBETWEEN.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cRANDBETWEEN.prototype.argumentsType = [argType.any, argType.any];
	cRANDBETWEEN.prototype.Calculate = function (arg) {

		function randBetween(a, b) {
			if ((a > 0 && a < 1) && (b > 0 && b < 1)) {
				return new cNumber(1);
			}

			if ((a < 0 && a > -1) && (b < 0 && b > -1)) {
				return new cNumber(0);
			}

			let firstNumber = Math.ceil(a);
			let secondNumber = Math.floor(b);

			if (firstNumber === secondNumber) {
				return new cNumber(firstNumber);
			}

			return new cNumber(Math.floor(Math.random() * (secondNumber - firstNumber + 1)) + firstNumber);
		}

		function f(a, b, r, c) {
			if (cElementType.number === a.type && cElementType.number === b.type) {
				this.array[r][c] = randBetween(a.getValue(), b.getValue());
			} else {
				this.array[r][c] = new cError(cErrorType.wrong_value_type);
			}
		}

		let arg0 = arg[0] ? arg[0] : new cError(cErrorType.not_available);
		let arg1 = arg[1] ? arg[1] : new cError(cErrorType.not_available);

		if (cElementType.bool === arg0.type || cElementType.bool === arg1.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			let arg0Val = arg0.getValue();
			if (arg0Val && arg0Val.type) {
				if (cElementType.bool === arg0Val.type) {
					return new cError(cErrorType.wrong_value_type);
				}
			}
		}

		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			let arg1Val = arg1.getValue();
			if (arg1Val && arg1Val.type) {
				if (cElementType.bool === arg1Val.type) {
					return new cError(cErrorType.wrong_value_type);
				}
			}
		}

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}

		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.array === arg0.type && cElementType.array === arg1.type) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					f.call(this, elem, arg1.getElementRowCol(r, c), r, c);
				});
				return arg0;
			}
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				f.call(this, elem, arg1, r, c)
			});
			return arg0;
		} else if (cElementType.array === arg1.type) {
			arg1.foreach(function (elem, r, c) {
				f.call(this, arg0, elem, r, c);
			});
			return arg1;
		}

		if (!(cElementType.number === arg0.type) || (arg1 && !(cElementType.number === arg0.type))) {
			return new cError(cErrorType.wrong_value_type);
		} else if (cElementType.number === arg0.type && cElementType.number === arg1.type) {
			if (arg0.getValue() > arg1.getValue()) {
				return new cError(cErrorType.not_numeric);
			}
		}

		return new cNumber(randBetween(arg0.getValue(), arg1.getValue()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROMAN() {
	}

	cROMAN.prototype = Object.create(cBaseFunction.prototype);
	cROMAN.prototype.constructor = cROMAN;
	cROMAN.prototype.name = 'ROMAN';
	cROMAN.prototype.argumentsMin = 1;
	cROMAN.prototype.argumentsMax = 2;
	cROMAN.prototype.argumentsType = [argType.number, argType.number];
	cROMAN.prototype.Calculate = function (arg) {
		function roman(num, mode) {
			if ((mode >= 0) && (mode < 5) && (num >= 0) && (num < 4000)) {
				var chars = ['M', 'D', 'C', 'L', 'X', 'V', 'I'], values = [1000, 500, 100, 50, 10, 5, 1],
					maxIndex = values.length - 1, aRoman = "", index, digit, index2, steps;
				for (var i = 0; i <= maxIndex / 2; i++) {
					index = 2 * i;
					digit = parseInt(num / values[index]);

					if ((digit % 5) == 4) {
						index2 = (digit == 4) ? index - 1 : index - 2;
						steps = 0;
						while ((steps < mode) && (index < maxIndex)) {
							steps++;
							if (values[index2] - values[index + 1] <= num) {
								index++;
							} else {
								steps = mode;
							}
						}
						aRoman += chars[index];
						aRoman += chars[index2];
						num = (num + values[index]);
						num = (num - values[index2]);
					} else {
						if (digit > 4) {
							aRoman += chars[index - 1];
						}
						for (var j = digit % 5; j > 0; j--) {
							aRoman += chars[index];
						}
						num %= values[index];
					}
				}
				return new cString(aRoman);
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		var arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(0);
		if (arg0 instanceof cArea || arg0 instanceof cArea3D || arg1 instanceof cArea || arg1 instanceof cArea3D) {
			return new cError(cErrorType.wrong_value_type);
		}
		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = roman(a.getValue(), b.getValue());
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem, b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = roman(a.getValue(), b.getValue());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0, b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = roman(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		return roman(arg0.getValue(), arg1.getValue());

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROUND() {
	}

	cROUND.prototype = Object.create(cBaseFunction.prototype);
	cROUND.prototype.constructor = cROUND;
	cROUND.prototype.name = 'ROUND';
	cROUND.prototype.argumentsMin = 2;
	cROUND.prototype.argumentsMax = 2;
	cROUND.prototype.inheritFormat = true;
	cROUND.prototype.argumentsType = [argType.number, argType.number];
	cROUND.prototype.Calculate = function (arg) {

		function SignZeroPositive(number) {
			return number < 0 ? -1 : 1;
		}

		function truncate(n) {
			return Math[n > 0 ? "floor" : "ceil"](n);
		}

		function Floor(number, significance) {
			var quotient = number / significance;
			if (quotient == 0) {
				return 0;
			}
			var nolpiat = 5 * Math.sign(quotient) *
				Math.pow(10, Math.floor(Math.log10(Math.abs(quotient))) - cExcelSignificantDigits);
			return truncate(quotient + nolpiat) * significance;
		}

		function roundHelper(number, decimals) {
			if (num_digits > AscCommonExcel.cExcelMaxExponent) {
				if (Math.abs(number) < 1 || num_digits < 1e10) // The values are obtained experimentally
				{
					return new cNumber(number);
				}
				return new cNumber(0);
			} else if (num_digits < AscCommonExcel.cExcelMinExponent) {
				if (Math.abs(number) < 0.01) // The values are obtained experimentally
				{
					return new cNumber(number);
				}
				return new cNumber(0);
			}

			const EPSILON = 1e-14;

			// ->integer
			decimals = decimals >> 0;

			const multiplier = Math.pow(10, decimals);
			const shifted = Math.abs(number) * multiplier;

			// Add epsilon to handle floating point precision issues (1.005 case)
			const compensated = shifted + EPSILON;
			const rounded = Math.floor(compensated + 0.5);

			let result = (Math.sign(number) * rounded) / multiplier;
			return new cNumber(result);
		}

		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
			if (arg0 instanceof cError) {
				return arg0;
			} else if (arg0 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg0 = arg0.tocNumber();
			}
		} else {
			arg0 = arg0.tocNumber();
		}

		if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
			if (arg1 instanceof cError) {
				return arg1;
			} else if (arg1 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg1 = arg1.tocNumber();
			}
		} else {
			arg1 = arg1.tocNumber();
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = roundHelper(a.getValue(), b.getValue())
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem;
				var b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = roundHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0;
				var b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = roundHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		var number = arg0.getValue(), num_digits = arg1.getValue();

		return roundHelper(number, num_digits);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROUNDDOWN() {
	}

	cROUNDDOWN.prototype = Object.create(cBaseFunction.prototype);
	cROUNDDOWN.prototype.constructor = cROUNDDOWN;
	cROUNDDOWN.prototype.name = 'ROUNDDOWN';
	cROUNDDOWN.prototype.argumentsMin = 2;
	cROUNDDOWN.prototype.argumentsMax = 2;
	cROUNDDOWN.prototype.inheritFormat = true;
	cROUNDDOWN.prototype.argumentsType = [argType.number, argType.number];
	cROUNDDOWN.prototype.Calculate = function (arg) {
		function rounddownHelper(number, num_digits) {
			num_digits = Math.trunc(num_digits);
			if (num_digits > AscCommonExcel.cExcelMaxExponent) {
				if (Math.abs(number) >= 1e-100 || num_digits <= 98303) { // The values are obtained experimentally
					return new cNumber(number);
				}
				return new cNumber(0);
			} else if (num_digits < AscCommonExcel.cExcelMinExponent) {
				if (Math.abs(number) >= 1e100) { // The values are obtained experimentally
					return new cNumber(number);
				}
				return new cNumber(0);
			}

			var significance = Math.pow(10, -num_digits);

			if (Number.POSITIVE_INFINITY == Math.abs(number / significance)) {
				return new cNumber(number);
			}
			var x = number * Math.pow(10, num_digits);
			x = Math.trunc(x);
			return new cNumber(x * significance);
		}

		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
			if (arg0 instanceof cError) {
				return arg0;
			} else if (arg0 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg0 = arg0.tocNumber();
			}
		} else {
			arg0 = arg0.tocNumber();
		}

		if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
			if (arg1 instanceof cError) {
				return arg1;
			} else if (arg1 instanceof cString) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				arg1 = arg1.tocNumber();
			}
		} else {
			arg1 = arg1.tocNumber();
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						this.array[r][c] = rounddownHelper(a.getValue(), b.getValue())
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				var a = elem;
				var b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = rounddownHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0;
				var b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					this.array[r][c] = rounddownHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		var number = arg0.getValue(), num_digits = arg1.getValue();
		return rounddownHelper(number, num_digits);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROUNDUP() {
	}

	cROUNDUP.prototype = Object.create(cBaseFunction.prototype);
	cROUNDUP.prototype.constructor = cROUNDUP;
	cROUNDUP.prototype.name = 'ROUNDUP';
	cROUNDUP.prototype.argumentsMin = 2;
	cROUNDUP.prototype.argumentsMax = 2;
	cROUNDUP.prototype.inheritFormat = true;
	cROUNDUP.prototype.argumentsType = [argType.number, argType.number];
	cROUNDUP.prototype.Calculate = function (arg) {
		function roundupHelper(number, num_digits) {
			let fractionalPart = number.toString().split(".")[1];

			if (num_digits > AscCommonExcel.cExcelMaxExponent) {
				if (Math.abs(number) >= 1e-100 || num_digits <= 98303) {
					return new cNumber(number);
				}
				return new cNumber(0);
			} else if (num_digits < AscCommonExcel.cExcelMinExponent) {
				if (Math.abs(number) >= 1e100) {
					return new cNumber(number);
				}
				return new cError(cErrorType.not_numeric);
			} else if (fractionalPart && fractionalPart.length === num_digits) {
				return new cNumber(number);
			}

			let sign = number >= 0 ? 1 : -1,
				absNum = Math.abs(number), significance, roundedAbsNum;

			if (!Number.isInteger(num_digits)) {
				num_digits = num_digits > 0 ? Math.floor(num_digits) : Math.ceil(num_digits);
			}

			significance = Math.pow(10, num_digits);
			roundedAbsNum = Math.ceil(absNum * significance);

			return new cNumber(sign * roundedAbsNum / significance);
		}

		let arg0 = arg[0], arg1 = arg[1];

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		}

		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}

		arg0 = arg0.tocNumber();
		arg1 = arg1.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.array === arg0.type && cElementType.array === arg1.type) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					let a = elem;
					let b = arg1.getElementRowCol(r, c);
					if (cElementType.number === a.type && cElementType.number === b.type) {
						this.array[r][c] = roundupHelper(a.getValue(), b.getValue());
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				let a = elem;
				let b = arg1;
				if (cElementType.number === a.type && cElementType.number === b.type) {
					this.array[r][c] = roundupHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (cElementType.array === arg1.type) {
			arg1.foreach(function (elem, r, c) {
				let a = arg0;
				let b = elem;
				if (cElementType.number === a.type && cElementType.number === b.type) {
					this.array[r][c] = roundupHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		let number = arg0.getValue(), num_digits = arg1.getValue();
		return roundupHelper(number, num_digits);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSEC() {
	}

	//***array-formula***
	cSEC.prototype = Object.create(cBaseFunction.prototype);
	cSEC.prototype.constructor = cSEC;
	cSEC.prototype.name = 'SEC';
	cSEC.prototype.argumentsMin = 1;
	cSEC.prototype.argumentsMax = 1;
	cSEC.prototype.isXLFN = true;
	cSEC.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cSEC.prototype.argumentsType = [argType.number];
	cSEC.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		var maxVal = Math.pow(2, 27);
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					if (Math.abs(elem.getValue()) >= maxVal) {
						this.array[r][c] = new cError(cErrorType.not_numeric);
					} else {
						var a = 1 / Math.cos(elem.getValue());
						this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
					}
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			if (Math.abs(arg0.getValue()) >= maxVal) {
				return new cError(cErrorType.not_numeric);
			}

			var a = 1 / Math.cos(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSECH() {
	}

	//***array-formula***
	cSECH.prototype = Object.create(cBaseFunction.prototype);
	cSECH.prototype.constructor = cSECH;
	cSECH.prototype.name = 'SECH';
	cSECH.prototype.argumentsMin = 1;
	cSECH.prototype.argumentsMax = 1;
	cSECH.prototype.isXLFN = true;
	cSECH.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cSECH.prototype.argumentsType = [argType.number];
	cSECH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		//TODO в документации к COTH написано максимальное значение - Math.pow(2, 27), но MS EXCEL в данном случае не выдает ошибку
		//проверку на максиимальное значение убрал
		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.error === arg0.type) {
			return arg0;
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				if (cElementType.number === elem.type) {
					var a = 1 / Math.cosh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else {
			var a = 1 / Math.cosh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSERIESSUM() {
	}

	//***array-formula***
	cSERIESSUM.prototype = Object.create(cBaseFunction.prototype);
	cSERIESSUM.prototype.constructor = cSERIESSUM;
	cSERIESSUM.prototype.name = 'SERIESSUM';
	cSERIESSUM.prototype.argumentsMin = 4;
	cSERIESSUM.prototype.argumentsMax = 4;
	cSERIESSUM.prototype.arrayIndexes = {3: 1};
	cSERIESSUM.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cSERIESSUM.prototype.argumentsType = [argType.any, argType.any, argType.any, argType.any];
	cSERIESSUM.prototype.Calculate = function (arg) {

		function SERIESSUM(x, n, m, a) {

			x = x.getValue();
			n = n.getValue();
			m = m.getValue();

			for (var i = 0; i < a.length; i++) {
				if (!(a[i] instanceof cNumber)) {
					return new cError(cErrorType.wrong_value_type);
				}
				a[i] = a[i].getValue();
			}

			function sumSeries(x, n, m, a) {
				var sum = 0;
				for (var i = 0; i < a.length; i++) {
					sum += a[i] * Math.pow(x, n + i * m)
				}
				return sum;
			}

			return new cNumber(sumSeries(x, n, m, a));
		}

		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2], arg3 = arg[3];
		if (arg0 instanceof cNumber || arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.tocNumber();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg1 instanceof cNumber || arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.tocNumber();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg2 instanceof cNumber || arg2 instanceof cRef || arg2 instanceof cRef3D) {
			arg2 = arg2.tocNumber();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		if (arg3 instanceof cNumber || arg3 instanceof cRef || arg3 instanceof cRef3D) {
			arg3 = [arg3.tocNumber()];
		} else if (arg3 instanceof cArea || arg3 instanceof cArea3D) {
			arg3 = arg3.getValue();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		return SERIESSUM(arg0, arg1, arg2, arg3);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSIGN() {
	}

	//***array-formula***
	cSIGN.prototype = Object.create(cBaseFunction.prototype);
	cSIGN.prototype.constructor = cSIGN;
	cSIGN.prototype.name = 'SIGN';
	cSIGN.prototype.argumentsMin = 1;
	cSIGN.prototype.argumentsMax = 1;
	cSIGN.prototype.argumentsType = [argType.number];
	cSIGN.prototype.Calculate = function (arg) {

		function signHelper(arg) {
			if (arg < 0) {
				return new cNumber(-1.0);
			} else if (arg == 0) {
				return new cNumber(0.0);
			} else {
				return new cNumber(1.0);
			}
		}

		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}

		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = elem.getValue();
					this.array[r][c] = signHelper(a)
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = arg0.getValue();
			return signHelper(a);
		}
		return arg0;

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSIN() {
	}

	//***array-formula***
	cSIN.prototype = Object.create(cBaseFunction.prototype);
	cSIN.prototype.constructor = cSIN;
	cSIN.prototype.name = 'SIN';
	cSIN.prototype.argumentsMin = 1;
	cSIN.prototype.argumentsMax = 1;
	cSIN.prototype.argumentsType = [argType.number];
	cSIN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.sin(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.sin(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSINGLE() {
	}

	//***array-formula***
	cSINGLE.prototype = Object.create(cBaseFunction.prototype);
	cSINGLE.prototype.constructor = cSINGLE;
	cSINGLE.prototype.name = 'SINGLE';
	cSINGLE.prototype.argumentsMin = 1;
	cSINGLE.prototype.argumentsMax = 1;
	cSINGLE.prototype.argumentsType = [argType.any];
	cSINGLE.prototype.arrayIndexes = {0: 1};
	cSINGLE.prototype.isXLFN = true;
	cSINGLE.prototype.Calculate = function (arg) {
		let arg0 = arg[0];
		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0.type === cElementType.array) {
			arg0 = arg0.getElement(0);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSINH() {
	}

	//***array-formula***
	cSINH.prototype = Object.create(cBaseFunction.prototype);
	cSINH.prototype.constructor = cSINH;
	cSINH.prototype.name = 'SINH';
	cSINH.prototype.argumentsMin = 1;
	cSINH.prototype.argumentsMax = 1;
	cSINH.prototype.argumentsType = [argType.number];
	cSINH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.sinh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.sinh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSQRT() {
	}

	//***array-formula***
	cSQRT.prototype = Object.create(cBaseFunction.prototype);
	cSQRT.prototype.constructor = cSQRT;
	cSQRT.prototype.name = 'SQRT';
	cSQRT.prototype.argumentsMin = 1;
	cSQRT.prototype.argumentsMax = 1;
	cSQRT.prototype.argumentsType = [argType.number];
	cSQRT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.sqrt(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.sqrt(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSQRTPI() {
	}

	//***array-formula***
	cSQRTPI.prototype = Object.create(cBaseFunction.prototype);
	cSQRTPI.prototype.constructor = cSQRTPI;
	cSQRTPI.prototype.name = 'SQRTPI';
	cSQRTPI.prototype.argumentsMin = 1;
	cSQRTPI.prototype.argumentsMax = 1;
	cSQRTPI.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cSQRTPI.prototype.argumentsType = [argType.any];
	cSQRTPI.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.sqrt(elem.getValue() * Math.PI);
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.sqrt(arg0.getValue() * Math.PI);
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUBTOTAL() {
	}

	//***array-formula***
	cSUBTOTAL.prototype = Object.create(cBaseFunction.prototype);
	cSUBTOTAL.prototype.constructor = cSUBTOTAL;
	cSUBTOTAL.prototype.name = 'SUBTOTAL';
	cSUBTOTAL.prototype.argumentsMin = 2;
	cSUBTOTAL.prototype.argumentsType = [argType.number, argType.reference, [argType.reference]];
	cSUBTOTAL.prototype.arrayIndexes = {1: 1, 2: 1, 3: 1, 4: 1, 5: 1, 6: 1, 7: 1};
	cSUBTOTAL.prototype.getArrayIndex = function (index) {
		if (index === 0) {
			return undefined;
		}
		return 1;
	};
	cSUBTOTAL.prototype.Calculate = function (arg) {
		let f, exclude = false, arg0 = arg[0];

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (cElementType.number !== arg0.type) {
			return arg0;
		}

		arg0 = arg0.getValue();

		switch (arg0) {
			case cSubTotalFunctionType.excludes.AVERAGE:
				exclude = true;
			case cSubTotalFunctionType.includes.AVERAGE:
				f = AscCommonExcel.cAVERAGE.prototype;
				break;
			case cSubTotalFunctionType.excludes.COUNT:
				exclude = true;
			case cSubTotalFunctionType.includes.COUNT:
				f = AscCommonExcel.cCOUNT.prototype;
				break;
			case cSubTotalFunctionType.excludes.COUNTA:
				exclude = true;
			case cSubTotalFunctionType.includes.COUNTA:
				f = AscCommonExcel.cCOUNTA.prototype;
				break;
			case cSubTotalFunctionType.excludes.MAX:
				exclude = true;
			case cSubTotalFunctionType.includes.MAX:
				f = AscCommonExcel.cMAX.prototype;
				break;
			case cSubTotalFunctionType.excludes.MIN:
				exclude = true;
			case cSubTotalFunctionType.includes.MIN:
				f = AscCommonExcel.cMIN.prototype;
				break;
			case cSubTotalFunctionType.excludes.PRODUCT:
				exclude = true;
			case cSubTotalFunctionType.includes.PRODUCT:
				f = cPRODUCT.prototype;
				break;
			case cSubTotalFunctionType.excludes.STDEV:
				exclude = true;
			case cSubTotalFunctionType.includes.STDEV:
				f = AscCommonExcel.cSTDEV.prototype;
				break;
			case cSubTotalFunctionType.excludes.STDEVP:
				exclude = true;
			case cSubTotalFunctionType.includes.STDEVP:
				f = AscCommonExcel.cSTDEVP.prototype;
				break;
			case cSubTotalFunctionType.excludes.SUM:
				exclude = true;
			case cSubTotalFunctionType.includes.SUM:
				f = cSUM.prototype;
				break;
			case cSubTotalFunctionType.excludes.VAR:
				exclude = true;
			case cSubTotalFunctionType.includes.VAR:
				f = AscCommonExcel.cVAR.prototype;
				break;
			case cSubTotalFunctionType.excludes.VARP:
				exclude = true;
			case cSubTotalFunctionType.includes.VARP:
				f = AscCommonExcel.cVARP.prototype;
				break;
		}

		// resArr - result for an array arguments
		// res - single cell result for a range arguments 
		let resArr, resDimensions, res;
		if (f) {
			// the inner results are ignored to avoid double summation
			let oldExcludeHiddenRows = f.excludeHiddenRows;
			let oldIgnoreNestedStAg = f.excludeNestedStAg;
			let oldCheckExclude = f.checkExclude;

			f.excludeHiddenRows = exclude;
			f.excludeNestedStAg = true;
			f.checkExclude = true;

			for (let i = 1; i < arg.length; i++) {
				// array may come as a result of computing internal functions (for example, offset or if)
				if (arg[i].type == cElementType.array) {
					resDimensions = arg[i].getDimensions();
					if (!resArr) {
						resArr = new cArray();
					}

					for (let r = 0; r < resDimensions.row; r++) {
						if (!resArr.rowCount || resArr.rowCount < (r + 1)) {
							resArr.addRow();
						}
						for (let c = 0; c < resDimensions.col; c++) {
							let existedArrayElem = resArr.getValue2(r, c);
							let newArrayElem = f.Calculate([arg[i].getElementRowCol(r, c)]);
							if (!existedArrayElem) {
								res ? resArr.addElementInRow(cSUM.prototype.Calculate([res, newArrayElem]), r) : resArr.addElementInRow(newArrayElem, r);
							} else {
								if (c > (resArr.countElementInRow[r] - 1)) {
									resArr.addElementInRow(cSUM.prototype.Calculate([existedArrayElem, newArrayElem]), r);
								} else {
									resArr.array[r][c] = cSUM.prototype.Calculate([existedArrayElem, newArrayElem]);
								}
							}
						}
					}
					// if we found an array, set res to null for further calculation
					res = null;

				} else if (arg[i].type === cElementType.error) {
					return arg[i];
				} else {
					res = res ? cSUM.prototype.Calculate([f.Calculate([arg[i]]), res]) : f.Calculate([arg[i]]);
					if (res && resArr) {
						// add result to each element of the array
						resArr.foreach(function(elem, r, c) {
							// elem = elem + res;
							resArr.array[r][c] = cSUM.prototype.Calculate([elem, res]);
						});

						// if we have an array, set res to null for further calculation
						res = null;
					}
				}
			}

			res = resArr ? resArr : f.Calculate(arg.slice(1));
			f.excludeHiddenRows = oldExcludeHiddenRows;
			f.excludeNestedStAg = oldIgnoreNestedStAg;
			f.checkExclude = oldCheckExclude;
		} else {
			res = new cError(cErrorType.wrong_value_type);
		}

		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUM() {
	}

	//***array-formula***
	cSUM.prototype = Object.create(cBaseFunction.prototype);
	cSUM.prototype.constructor = cSUM;
	cSUM.prototype.name = 'SUM';
	cSUM.prototype.argumentsMin = 1;
	cSUM.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cSUM.prototype.inheritFormat = true;
	cSUM.prototype.argumentsType = [[argType.number]];
	cSUM.prototype.Calculate = function (arg) {
		var element, _arg, arg0 = new cNumber(0);
		for (var i = 0; i < arg.length; i++) {
			element = arg[i];
			if (cElementType.cellsRange === element.type || cElementType.cellsRange3D === element.type) {
				var _arrVal = element.getValue(this.checkExclude, this.excludeHiddenRows, this.excludeErrorsVal,
					this.excludeNestedStAg);
				for (var j = 0; j < _arrVal.length; j++) {
					if (cElementType.bool !== _arrVal[j].type && cElementType.string !== _arrVal[j].type) {
						arg0 = _func[arg0.type][_arrVal[j].type](arg0, _arrVal[j], "+");
					}
					if (cElementType.error === arg0.type) {
						return arg0;
					}
				}
			} else if (cElementType.cell === element.type || cElementType.cell3D === element.type) {
				if (!this.checkExclude || !element.isHidden(this.excludeHiddenRows)) {
					_arg = element.getValue();
					if (cElementType.bool !== _arg.type && cElementType.string !== _arg.type) {
						arg0 = _func[arg0.type][_arg.type](arg0, _arg, "+");
					}
				}
			} else if (cElementType.array === element.type) {
				element.foreach(function (arrElem) {
					if (cElementType.cell === arrElem.type || cElementType.cell3D === arrElem.type) {
						arrElem = arrElem.getValue();
					}
					if (cElementType.bool !== arrElem.type && cElementType.string !== arrElem.type &&
						cElementType.empty !== arrElem.type) {
						arg0 = _func[arg0.type][arrElem.type](arg0, arrElem, "+");
					}
				});
			} else {
				_arg = element.tocNumber();
				arg0 = _func[arg0.type][_arg.type](arg0, _arg, "+");
			}
			if (cElementType.error === arg0.type) {
				return arg0;
			}

		}

		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMIF() {
	}

	//***array-formula***
	cSUMIF.prototype = Object.create(cBaseFunction.prototype);
	cSUMIF.prototype.constructor = cSUMIF;
	cSUMIF.prototype.name = 'SUMIF';
	cSUMIF.prototype.argumentsMin = 2;
	cSUMIF.prototype.argumentsMax = 3;
	cSUMIF.prototype.arrayIndexes = {0: 1, 2: 1};
	cSUMIF.prototype.exactTypes = {0: 1};
	cSUMIF.prototype.argumentsType = [argType.reference, argType.any, argType.reference];
	cSUMIF.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1], arg2 = arg[2] ? arg[2] : arg[0], _sum = 0, matchingInfo;
		if (cElementType.cell !== arg0.type && cElementType.cell3D !== arg0.type &&
			cElementType.cellsRange !== arg0.type) {
			if (cElementType.cellsRange3D === arg0.type) {
				arg0 = arg0.tocArea();
				if (!arg0) {
					return new cError(cErrorType.wrong_value_type);
				}
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (cElementType.cell !== arg2.type && cElementType.cell3D !== arg2.type &&
			cElementType.cellsRange !== arg2.type) {
			if (cElementType.cellsRange3D === arg2.type) {
				arg2 = arg2.tocArea();
				if (!arg2) {
					return new cError(cErrorType.wrong_value_type);
				}
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (cElementType.string !== arg1.type) {
			arg1 = arg1.tocString();
		}

		if (cElementType.string !== arg1.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		matchingInfo = AscCommonExcel.matchingValue(arg1);
		if (cElementType.cellsRange === arg0.type || cElementType.cell === arg0.type) {
			let arg0Matrix = arg0.getMatrix(), arg2Matrix = arg[2] ? arg2.getMatrix() : arg0Matrix, valMatrix2;
			for (let i = 0; i < arg0Matrix.length; i++) {
				for (let j = 0; j < arg0Matrix[i].length; j++) {
					if (arg2Matrix[i] && (valMatrix2 = arg2Matrix[i][j]) && cElementType.number === valMatrix2.type &&
						AscCommonExcel.matching(arg0Matrix[i][j], matchingInfo)) {
						_sum += valMatrix2.getValue();
					}
				}
			}
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		return new cNumber(_sum);
	};

	/**
	 * @constructor
	 */
	function SUMIFSCache() {
		this.cacheRanges = {};
		this.cacheId = {};
	}

	SUMIFSCache.prototype._get = function (range, arg, valueForSearching, parent) {
		var res, _this = this, wsId = range.getWorksheet().getId();
		var sRangeName;
		AscCommonExcel.executeInR1C1Mode(false, function () {
			sRangeName = wsId + AscCommon.g_cCharDelimiter + range.getName();
		});
		var sInputKey = valueForSearching.getValue() + AscCommon.g_cCharDelimiter + sRangeName /*+ g_cCharDelimiter + valueForSearching.type*/;
		var cacheElem;
		if (parent) {
			if (!parent.cacheId) {
				parent.cacheId = [];
			}
			cacheElem = parent.cacheId[sInputKey];
		} else {
			cacheElem = this.cacheId[sInputKey];
		}

		if (!cacheElem) {
			cacheElem = {elems: null, cacheId: null};

			if (parent && parent.cacheId) {
				parent.cacheId[sInputKey] = cacheElem;
			} else {
				this.cacheId[sInputKey] = cacheElem;
			}

			var cacheRange = this.cacheRanges[wsId];
			if (!cacheRange) {
				cacheRange = new AscCommonExcel.RangeDataManager(null);
				this.cacheRanges[wsId] = cacheRange;
			}
			cacheRange.add(range.getBBox0(), cacheElem);
		}

		return cacheElem;
	};
	SUMIFSCache.prototype.clean = function () {
		this.cacheRanges = {};
		this.cacheId = {};
	};
	SUMIFSCache.prototype.remove = function (cell) {
		var wsId = cell.ws.getId();
		var cacheRange = this.cacheRanges[wsId];
		if (cacheRange) {
			var oGetRes = cacheRange.get(new Asc.Range(cell.nCol, cell.nRow, cell.nCol, cell.nRow));
			for (var i = 0, length = oGetRes.all.length; i < length; ++i) {
				var elem = oGetRes.all[i];
				elem.data.cacheId = null;
				elem.data.elems = null;
			}
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMIFS() {
	}

	//TODO есть расхождение с MS - смотри в файле + arrayIndexes - нужно формировать при условии что все нечетные аргумента - массивы(начиная с 3 аргумента)
	//***array-formula***
	cSUMIFS.prototype = Object.create(cBaseFunction.prototype);
	cSUMIFS.prototype.constructor = cSUMIFS;
	cSUMIFS.prototype.name = 'SUMIFS';
	cSUMIFS.prototype.argumentsMin = 3;
	cSUMIFS.prototype.arrayIndexes = {0: 1, 1: 1, 3: 1, 5: 1, 7: 1};
	cSUMIFS.prototype.getArrayIndex = function (index) {
		if (index === 0) {
			return 1;
		}
		return index % 2 !== 0 ? 1 : undefined;
	};
	cSUMIFS.prototype.exactTypes = {0: 1, 1: 1};	// in this function every odd argument is should be checked for type reference
	cSUMIFS.prototype.argumentsType = [argType.reference, [argType.reference, argType.any]];
	cSUMIFS.prototype.Calculate = function (arg) {
		let arg0 = arg[0];
		if (cElementType.cell !== arg0.type && cElementType.cell3D !== arg0.type && cElementType.cellsRange !== arg0.type) {
			if (cElementType.cellsRange3D === arg0.type) {
				arg0 = arg0.tocArea();
				if (!arg0) {
					return new cError(cErrorType.wrong_value_type);
				}
			} else {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		let getRange = function (curArg) {
			let res = null;
			if (cElementType.cellsRange === curArg.type || cElementType.cell === curArg.type) {
				res = curArg.range && curArg.range.bbox ? curArg.range.bbox : null;
			} else if (cElementType.cellsRange3D === curArg.type || cElementType.cell3D === curArg.type) {
				res = curArg.getBBox0();
			}
			return res;
		};

		let getMatrixFromCache = function (area, searchVal, parent) {
			let range = area.getRange();
			let ws = area.getWS();
			let bb = range.getBBox0();
			let oSearchRange = ws.getRange3(bb.r1, bb.c1, bb.r2, bb.c2);
			return g_oSUMIFSCache._get(oSearchRange, range, searchVal, parent);
		};

		let getElems = function (a1, a2, p, calcSum) {
			let matchingInfo = AscCommonExcel.matchingValue(a2);

			let arg1Matrix = a1.getMatrix(), enterMatrix;
			if (p && p.length) {
				enterMatrix = p;
			} else {
				enterMatrix = arg1Matrix;
			}

			let res = undefined;
			for (let i = 0; i < enterMatrix.length; ++i) {
				if (!arg1Matrix[i]) {
					continue;
				}
				if (!enterMatrix[i]) {
					continue;
				}

				for (let j = 0; j < enterMatrix[i].length; ++j) {
					if (!enterMatrix[i][j] || !arg1Matrix[i][j]) {
						continue;
					}

					if (AscCommonExcel.matching(arg1Matrix[i][j], matchingInfo)) {
						if (!res) {
							res = [];
						}
						if (!res[i]) {
							res[i] = [];
						}
						res[i][j] = arg1Matrix[i][j];
						if (calcSum) {
							_calcSum(i, j);
						}
					}
				}
			}

			return res;
		};

		let _calcSum = function (i, j) {
			let arg0Val = arg0.getValueByRowCol ? arg0.getValueByRowCol(i, j) : arg0.tocNumber();
			if (arg0Val && cElementType.number === arg0Val.type) {
				_sum += arg0Val.getValue();
			}
		};

		let _sum = 0;
		let arg0Range = getRange(arg0);
		let arg0C = arg0Range.c2 - arg0Range.c1 + 1;
		let arg0R = arg0Range.r2 - arg0Range.r1 + 1;
		let cacheElem, parent;
		let arg1, arg2, arg1Range;
		for (let k = 1; k < arg.length; k += 2) {
			arg1 = arg[k];
			arg2 = arg[k + 1];


			arg1Range = getRange(arg1);
			if (!arg1Range) {
				return new cError(cErrorType.wrong_value_type);
			}

			let arg1C = arg1Range.c2 - arg1Range.c1 + 1;
			let arg1R = arg1Range.r2 - arg1Range.r1 + 1;
			if (arg0R !== arg1R || arg0C !== arg1C) {
				return new cError(cErrorType.wrong_value_type);
			}

			//в кэш кладём истинное значение для поиска, а не весь диапазон
			if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
				arg2 = arg2.cross(arguments[1]);
			} else if (cElementType.array === arg2.type) {
				arg2 = arg2.getElementRowCol(0, 0);
			}

			arg2 = arg2.tocString();

			parent = cacheElem;
			cacheElem = getMatrixFromCache(arg1, arg2, cacheElem);

			if (null === cacheElem.elems) {
				if (cElementType.cell !== arg1.type && cElementType.cell3D !== arg1.type &&
					cElementType.cellsRange !== arg1.type) {
					if (cElementType.cellsRange3D === arg1.type) {
						arg1 = arg1.tocArea();
						if (!arg1) {
							return new cError(cErrorType.wrong_value_type);
						}
					} else {
						return new cError(cErrorType.wrong_value_type);
					}
				}

				if (cElementType.string !== arg2.type) {
					return new cError(cErrorType.wrong_value_type);
				}

				cacheElem.elems = getElems(arg1, arg2, parent ? parent.elems : null, k + 1 === arg.length - 1);
			} else if (k + 1 === arg.length - 1 && cacheElem.elems) {
				for (let i = 0; i < cacheElem.elems.length; i++) {
					if (cacheElem.elems[i]) {
						for (let j = 0; j < cacheElem.elems[i].length; j++) {
							if (cacheElem.elems[i][j]) {
								_calcSum(i, j);
							}
						}
					}
				}
			}

			if (undefined === cacheElem.elems) {
				return new cNumber(0);
			}
			if (cElementType.error === cacheElem.elems.type) {
				return cacheElem.elems;
			}
		}

		return new cNumber(_sum);
	};
	cSUMIFS.prototype.checkArguments = function (countArguments) {
		return 1 === countArguments % 2 && cBaseFunction.prototype.checkArguments.apply(this, arguments);
	};
	cSUMIFS.prototype.checkArgumentsTypes = function (args) {
		// check first element, then all odd ones
		if (!cBaseFunction.prototype.checkArgumentsTypes.call(this, [args[0]])) {
			return false
		}

		for (let i = 1; i < args.length; i += 2) {
			// check reference type for each odd element
			let oddArgument = args[i];
			if (oddArgument && this.exactTypes[1]) {
				if (oddArgument.type !== cElementType.cellsRange && oddArgument.type !== cElementType.cellsRange3D 
					&& oddArgument.type !== cElementType.cell && oddArgument.type !== cElementType.cell3D) {
						return false
				}
			}
		}
		return true;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMPRODUCT() {
	}

	//***array-formula***
	cSUMPRODUCT.prototype = Object.create(cBaseFunction.prototype);
	cSUMPRODUCT.prototype.constructor = cSUMPRODUCT;
	cSUMPRODUCT.prototype.name = 'SUMPRODUCT';
	cSUMPRODUCT.prototype.argumentsMin = 1;
	cSUMPRODUCT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cSUMPRODUCT.prototype.argumentsType = [[argType.array]];
	cSUMPRODUCT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cSUMPRODUCT.prototype.Calculate = function (arg) {
		let arg0 = new cNumber(0), resArr = [], col = 0, row = 0, res = 1, _res = [], i;

		for (i = 0; i < arg.length; i++) {
			if (cElementType.cellsRange === arg[i].type || cElementType.array === arg[i].type) {
				resArr[i] = arg[i].getMatrix();
			} else if (cElementType.cellsRange3D === arg[i].type) {
				if (arg[i].isSingleSheet()) {
					resArr[i] = arg[i].getMatrix()[0];
				} else {
					return new cError(cErrorType.bad_reference);
				}
			} else if (cElementType.cell === arg[i].type || cElementType.cell3D === arg[i].type) {
				let val = arg[i].getValue();
				if (cElementType.empty === val.type) {
					return new cError(cErrorType.wrong_value_type);
				} else {
					resArr[i] = [[val]];
				}
			} else {
				resArr[i] = [[arg[i]]];
			}

			let matrixSize = arg[i].getDimensions();

			row = Math.max(matrixSize.row, row);
			col = Math.max(matrixSize.col, col);

			if (row !== matrixSize.row || col !== matrixSize.col) {
				return new cError(cErrorType.not_numeric);
			}

			if (cElementType.error === arg[i].type) {
				return arg[i];
			}
		}

		let emptyVal = new cEmpty();
		for (let iRow = 0; iRow < row; iRow++) {
			for (let iCol = 0; iCol < col; iCol++) {
				res = 1;
				for (let iRes = 0; iRes < resArr.length; iRes++) {
					arg0 = resArr[iRes] && resArr[iRes][iRow] && resArr[iRes][iRow][iCol];
					if (!arg0) {
						arg0 = emptyVal;
					}
					if (cElementType.error === arg0.type) {
						return arg0;
					} else if (cElementType.string === arg0.type) {
						if (cElementType.error === arg0.tocNumber().type) {
							res *= 0;
						} else {
							res *= arg0.tocNumber().getValue();
						}
					} else if (cElementType.bool === arg0.type) {
						res *= 0;
					} else {
						res *= arg0.tocNumber().getValue();
					}
				}
				_res.push(res);
			}
		}
		res = 0;
		for (i = 0; i < _res.length; i++) {
			res += _res[i]
		}

		return new cNumber(res);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMSQ() {
	}

	//***array-formula***
	cSUMSQ.prototype = Object.create(cBaseFunction.prototype);
	cSUMSQ.prototype.constructor = cSUMSQ;
	cSUMSQ.prototype.name = 'SUMSQ';
	cSUMSQ.prototype.argumentsMin = 1;
	cSUMSQ.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cSUMSQ.prototype.argumentsType = [[argType.number]];
	cSUMSQ.prototype.Calculate = function (arg) {
		var arg0 = new cNumber(0), _arg;

		function sumsqHelper(a, b) {
			var c = _func[b.type][b.type](b, b, "*");
			return _func[a.type][c.type](a, c, "+");
		}

		for (var i = 0; i < arg.length; i++) {
			if (arg[i] instanceof cArea || arg[i] instanceof cArea3D) {
				var _arrVal = arg[i].getValue();
				for (var j = 0; j < _arrVal.length; j++) {
					if (_arrVal[j] instanceof cNumber) {
						arg0 = sumsqHelper(arg0, _arrVal[j]);
					} else if (_arrVal[j] instanceof cError) {
						return _arrVal[j];
					}
				}
			} else if (arg[i] instanceof cRef || arg[i] instanceof cRef3D) {
				_arg = arg[i].getValue();
				if (_arg instanceof cNumber) {
					arg0 = sumsqHelper(arg0, _arg);
				}
			} else if (arg[i] instanceof cArray) {
				arg[i].foreach(function (arrElem) {
					if (arrElem instanceof cNumber) {
						arg0 = sumsqHelper(arg0, arrElem);
					}
				})
			} else {
				_arg = arg[i].tocNumber();
				arg0 = sumsqHelper(arg0, _arg);
			}
			if (arg0 instanceof cError) {
				return arg0;
			}

		}

		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMX2MY2() {
	}

	cSUMX2MY2.prototype = Object.create(cBaseFunction.prototype);
	cSUMX2MY2.prototype.constructor = cSUMX2MY2;
	cSUMX2MY2.prototype.name = 'SUMX2MY2';
	cSUMX2MY2.prototype.argumentsMin = 2;
	cSUMX2MY2.prototype.argumentsMax = 2;
	cSUMX2MY2.prototype.arrayIndexes = {0: 1, 1: 1};
	cSUMX2MY2.prototype.argumentsType = [argType.array, argType.array];
	cSUMX2MY2.prototype.Calculate = function (arg) {
		var func = function (a, b) {
			return a * a - b * b;
		};

		return sumxCalc(arg, func);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMX2PY2() {
	}

	cSUMX2PY2.prototype = Object.create(cBaseFunction.prototype);
	cSUMX2PY2.prototype.constructor = cSUMX2PY2;
	cSUMX2PY2.prototype.name = 'SUMX2PY2';
	cSUMX2PY2.prototype.argumentsMin = 2;
	cSUMX2PY2.prototype.argumentsMax = 2;
	cSUMX2PY2.prototype.arrayIndexes = {0: 1, 1: 1};
	cSUMX2PY2.prototype.argumentsType = [argType.array, argType.array];
	cSUMX2PY2.prototype.Calculate = function (arg) {

		var func = function (a, b) {
			return a * a + b * b;
		};

		return sumxCalc(arg, func);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUMXMY2() {
	}

	cSUMXMY2.prototype = Object.create(cBaseFunction.prototype);
	cSUMXMY2.prototype.constructor = cSUMXMY2;
	cSUMXMY2.prototype.name = 'SUMXMY2';
	cSUMXMY2.prototype.argumentsMin = 2;
	cSUMXMY2.prototype.argumentsMax = 2;
	cSUMXMY2.prototype.arrayIndexes = {0: 1, 1: 1};
	cSUMXMY2.prototype.argumentsType = [argType.array, argType.array];
	cSUMXMY2.prototype.Calculate = function (arg) {

		var func = function (a, b) {
			return (a - b) * (a - b);
		};

		return sumxCalc(arg, func);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTAN() {
	}

	//***array-formula***
	cTAN.prototype = Object.create(cBaseFunction.prototype);
	cTAN.prototype.constructor = cTAN;
	cTAN.prototype.name = 'TAN';
	cTAN.prototype.argumentsMin = 1;
	cTAN.prototype.argumentsMax = 1;
	cTAN.prototype.argumentsType = [argType.number];
	cTAN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.tan(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.tan(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTANH() {
	}

	//***array-formula***
	cTANH.prototype = Object.create(cBaseFunction.prototype);
	cTANH.prototype.constructor = cTANH;
	cTANH.prototype.name = 'TANH';
	cTANH.prototype.argumentsMin = 1;
	cTANH.prototype.argumentsMax = 1;
	cTANH.prototype.argumentsType = [argType.number];
	cTANH.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();
		if (arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				if (elem instanceof cNumber) {
					var a = Math.tanh(elem.getValue());
					this.array[r][c] = isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			})
		} else {
			var a = Math.tanh(arg0.getValue());
			return isNaN(a) ? new cError(cErrorType.not_numeric) : new cNumber(a);
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTRUNC() {
	}

	cTRUNC.prototype = Object.create(cBaseFunction.prototype);
	cTRUNC.prototype.constructor = cTRUNC;
	cTRUNC.prototype.name = 'TRUNC';
	cTRUNC.prototype.argumentsMin = 1;
	cTRUNC.prototype.argumentsMax = 2;
	cTRUNC.prototype.inheritFormat = true;
	cTRUNC.prototype.argumentsType = [argType.number, argType.number];
	cTRUNC.prototype.Calculate = function (arg) {
		// TODO fix floating point number precision problem (IEEE754)
		// https://0.30000000000000004.com/

		function truncHelper(a, b) {
			//TODO возможно стоит добавить ограничения для коэффициента b(ms не ограничивает; LO - максимальные значения 20/-20)
			if (b > 20) {
				b = 20;
			} else if (!Number.isInteger(b)) {
				b = +b.toFixed();
			}

			let numDegree = Math.pow(10, b);

			return new cNumber(Math.trunc(a * numDegree) / numDegree);
		}

		let arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(0);

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			if (arg0.isOneElement()) {
				arg0 = arg0.getFirstElement();
			} else {
				arg0 = new cError(cErrorType.wrong_value_type);
			}
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0, 0);
		}
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			if (arg1.isOneElement()) {
				arg1 = arg1.getFirstElement();
			} else {
				arg1 = new cError(cErrorType.wrong_value_type);
			}
		} else if (cElementType.array === arg1.type) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue().tocNumber();
		} else {
			arg0 = arg0.tocNumber();
		}

		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue().tocNumber();
		} else {
			arg1 = arg1.tocNumber();
		}


		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.array === arg0.type && cElementType.array === arg1.type) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					let a = elem;
					let b = arg1.getElementRowCol(r, c);
					if (cElementType.number === a.type && cElementType.number === b.type) {
						this.array[r][c] = truncHelper(a.getValue(), b.getValue())
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (cElementType.array === arg0.type) {
			arg0.foreach(function (elem, r, c) {
				let a = elem;
				let b = arg1;
				if (cElementType.number === a.type && cElementType.number === b.type) {
					this.array[r][c] = truncHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (cElementType.array === arg1.type) {
			arg1.foreach(function (elem, r, c) {
				let a = arg0;
				let b = elem;
				if (cElementType.number === a.type && cElementType.number === b.type) {
					this.array[r][c] = truncHelper(a.getValue(), b.getValue())
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		return truncHelper(arg0.getValue(), arg1.getValue());
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSEQUENCE() {
	}

	cSEQUENCE.prototype = Object.create(cBaseFunction.prototype);
	cSEQUENCE.prototype.constructor = cSEQUENCE;
	cSEQUENCE.prototype.name = 'SEQUENCE';
	cSEQUENCE.prototype.argumentsMin = 1;
	cSEQUENCE.prototype.argumentsMax = 4;
	cSEQUENCE.prototype.inheritFormat = true;
	cSEQUENCE.prototype.isXLFN = true;
	cSEQUENCE.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cSEQUENCE.prototype.argumentsType = [argType.number, argType.number, argType.number, argType.number];
	cSEQUENCE.prototype.Calculate = function (arg) {
		// TODO after implementing array autoexpanding, reconsider the behavior of the function when receiving cellsRange as arguments
		function sequenceArray(row, column, start, step) {
			let res = new cArray(),
				startNum = start,
				stepNum = step;

			for (let i = 0; i < row; i++) {
				res.addRow();
				for (let j = 0; j < column; j++) {
					res.addElement(new cNumber(startNum));
					startNum += stepNum;
				}
			}

			return res;
		}

		function sequenceRangeArrayGeneral(args, isRange) {
			const EXPECTED_MAX_ARRAY = 10223960;
			let rowVal = args[0],
				columnVal = args[1],
				startVal = args[2],
				stepVal = args[3];

			// ------------------------- arg0 empty val check -------------------------//
			if (!rowVal) {
				rowVal = new cNumber(0);
			}
			if (cElementType.cell === rowVal.type || cElementType.cell3D === rowVal.type) {
				rowVal = rowVal.getValue();
			}

			// ------------------------- arg1 empty type check -------------------------//
			if (!columnVal) {
				columnVal = new cNumber(0);
			}
			if (cElementType.cell === columnVal.type || cElementType.cell3D === columnVal.type) {
				columnVal = columnVal.getValue();
			}

			// ------------------------- arg2 empty type check -------------------------//
			if (!startVal) {
				startVal = new cNumber(0);
			}
			if (cElementType.cell === startVal.type || cElementType.cell3D === startVal.type) {
				startVal = startVal.getValue();
			}

			// ------------------------- arg3 empty type check -------------------------//
			if (!stepVal) {
				stepVal = new cNumber(0);
			}
			if (cElementType.cell === stepVal.type || cElementType.cell3D === stepVal.type) {
				stepVal = stepVal.getValue();
			}

			rowVal = rowVal.tocNumber();
			columnVal = columnVal.tocNumber();
			startVal = startVal.tocNumber();
			stepVal = stepVal.tocNumber();

			if (cElementType.error === rowVal.type) {
				return rowVal;
			}
			if (cElementType.error === columnVal.type) {
				return columnVal;
			}
			if (cElementType.error === startVal.type) {
				return startVal;
			}
			if (cElementType.error === stepVal.type) {
				return stepVal;
			}

			rowVal = Math.floor(rowVal.getValue());
			columnVal = Math.floor(columnVal.getValue());
			startVal = startVal.getValue();
			stepVal = stepVal.getValue();

			if (rowVal < 1 || columnVal < 1 || (rowVal * columnVal) > EXPECTED_MAX_ARRAY) {
				return new cError(cErrorType.wrong_value_type);
			}

			return isRange ? new cNumber(startVal) : sequenceArray(rowVal, columnVal, startVal, stepVal);
		}

		let arg0 = arg[0],
			arg1 = arg[1] ? arg[1] : new cNumber(1),
			arg2 = arg[2] ? arg[2] : new cNumber(1),
			arg3 = arg[3] ? arg[3] : new cNumber(1),
			res;
		// exceptions = new Map();

		if (arg0.type === cElementType.empty) {
			arg0 = new cNumber(1);
		}
		if (arg1.type === cElementType.empty) {
			arg1 = new cNumber(1);
		}
		if (arg2.type === cElementType.empty) {
			arg2 = new cNumber(1);
		}
		if (arg3.type === cElementType.empty) {
			arg3 = new cNumber(1);
		}

		// if range/array type, write array to map and call arrayHelper
		res = AscCommonExcel.getArrayHelper([arg0, arg1, arg2, arg3], sequenceRangeArrayGeneral);

		if (res) {
			return res;
		}

		return res ? res : sequenceRangeArrayGeneral([arg0, arg1, arg2, arg3], false);
	};


	var g_oSUMIFSCache = new SUMIFSCache();

	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].cAGGREGATE = cAGGREGATE;
	window['AscCommonExcel'].cPRODUCT = cPRODUCT;
	window['AscCommonExcel'].cSUBTOTAL = cSUBTOTAL;
	window['AscCommonExcel'].cSUM = cSUM;
	window['AscCommonExcel'].g_oSUMIFSCache = g_oSUMIFSCache;
})(window);
