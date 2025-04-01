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
	// Import
	var cElementType = AscCommonExcel.cElementType;
	var CellValueType = AscCommon.CellValueType;
	var g_oFormatParser = AscCommon.g_oFormatParser;
	var oNumFormatCache = AscCommon.oNumFormatCache;
	var CellFormat = AscCommon.CellFormat;
	var cErrorType = AscCommonExcel.cErrorType;
	var cNumber = AscCommonExcel.cNumber;
	var cString = AscCommonExcel.cString;
	var cBool = AscCommonExcel.cBool;
	var cError = AscCommonExcel.cError;
	var cEmpty = AscCommonExcel.cEmpty;
	var cArea = AscCommonExcel.cArea;
	var cArea3D = AscCommonExcel.cArea3D;
	var cRef = AscCommonExcel.cRef;
	var cRef3D = AscCommonExcel.cRef3D;
	var cArray = AscCommonExcel.cArray;
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var argType = Asc.c_oAscFormulaArgumentType;

	cFormulaFunctionGroup['TextAndData'] = cFormulaFunctionGroup['TextAndData'] || [];
	cFormulaFunctionGroup['TextAndData'].push(cARRAYTOTEXT, cASC, cBAHTTEXT, cCHAR, cCLEAN, cCODE, cCONCATENATE, cCONCAT, cDOLLAR,
		cEXACT, cFIND, cFINDB, cFIXED, cIMPORTRANGE, cJIS, cLEFT, cLEFTB, cLEN, cLENB, cLOWER, cMID, cMIDB, cNUMBERVALUE, cPHONETIC,
		cPROPER, cREPLACE, cREPLACEB, cREPT, cRIGHT, cRIGHTB, cSEARCH, cSEARCHB, cSUBSTITUTE, cT, cTEXT, cTEXTJOIN,
		cTRIM, cUNICHAR, cUNICODE, cUPPER, cVALUE, cTEXTBEFORE, cTEXTAFTER, cTEXTSPLIT);

	cFormulaFunctionGroup['NotRealised'] = cFormulaFunctionGroup['NotRealised'] || [];
	cFormulaFunctionGroup['NotRealised'].push(cBAHTTEXT, cJIS, cPHONETIC);

	function calcBeforeAfterText(arg, arg1, isAfter) {
		let newArgs = cBaseFunction.prototype._prepareArguments.call(this, arg, arg1, null, null, true).args;
		let text = newArgs[0];
		text = text.tocString();
		if (text.type === cElementType.error) {
			return text;
		}
		text = text.toString();

		let delimiter;
		if (cElementType.cellsRange === arg[1].type || cElementType.array === arg[1].type || cElementType.cellsRange3D === arg[1].type) {
			let isError;
			arg[1].foreach2(function (v) {
				v = v.tocString();
				if (v.type === cElementType.error) {
					isError = v;
				}
				if (!delimiter) {
					delimiter = [];
				}
				delimiter.push(v.toString());
			});
			if (isError) {
				return isError;
			}
			if (!delimiter) {
				delimiter = [""];
			}
		} else {
			delimiter = arg[1].tocString();
			if (delimiter.type === cElementType.error) {
				return delimiter;
			}
			delimiter = [delimiter.toString()];
		}

		let doSearch = function (_text, aDelimiters) {
			let needIndex = -1;
			for (let j = 0; j < aDelimiters.length; j++) {
				let nextDelimiter = match_mode ? aDelimiters[j].toLowerCase() : aDelimiters[j];
				let nextIndex = isReverseSearch ? modifiedText.lastIndexOf(nextDelimiter, startPos) : modifiedText.indexOf(nextDelimiter, startPos);
				if (needIndex === -1 || (((nextIndex < needIndex && !isReverseSearch) || (nextIndex > needIndex && isReverseSearch)) && nextIndex !== -1)) {
					needIndex = nextIndex;
					modifiedDelimiter = nextDelimiter;
				}
			}
			return needIndex;
		};

		//instance_num - при отрицательном вхождении поиск с конца начинается
		let instance_num = newArgs[2] && !(newArgs[2].type === cElementType.empty) ? newArgs[2] : new cNumber(1);
		let match_mode = newArgs[3] && !(newArgs[3].type === cElementType.empty) ? newArgs[3] : new cBool(false);
		let match_end = newArgs[4] && !(newArgs[4].type === cElementType.empty) ? newArgs[4] : new cBool(false);

		match_mode = match_mode.tocBool();
		match_end = match_end.tocBool();

		if (instance_num.type === cElementType.error) {
			return instance_num;
		}
		if (match_mode.type === cElementType.error) {
			return match_mode;
		}
		if (match_end.type === cElementType.error) {
			return match_end;
		}

		instance_num = instance_num.toNumber ? instance_num.toNumber() : 0;
		if (instance_num === 0 || (instance_num > text.length && newArgs[2] && newArgs[2].type !== cElementType.empty)) {
			//Excel returns a #VALUE! error if instance_num = 0 or if instance_num is greater than the length of text.
			return new cError(cErrorType.wrong_value_type);
		}

		match_mode = match_mode.toBool();
		match_end = match_end.toBool();

		let if_not_found = newArgs[5] ? newArgs[5] : new cError(cErrorType.not_available);

		//calculate
		let modifiedText = match_mode ? text.toLowerCase() : text;
		let modifiedDelimiter;

		let isReverseSearch = instance_num < 0;
		let foundIndex = -1;
		let startPos = isReverseSearch ? modifiedText.length : 0;
		let repeatZero = 0;
		let match_end_active = false;
		for (let i = 0; i < Math.abs(instance_num); i++) {
			foundIndex = doSearch(modifiedText, delimiter);
			if (foundIndex === 0) {
				repeatZero++;
			}
			if (foundIndex === -1) {
				if (match_end && i === Math.abs(instance_num) - 1) {
					foundIndex = isReverseSearch ? 0 : text.length;
					match_end_active = true;
				}
				break;
			}
			startPos = isReverseSearch ? foundIndex - modifiedDelimiter.length : foundIndex + modifiedDelimiter.length;
		}

		if (foundIndex === -1) {
			return if_not_found;
		} else {
			return new cString(isAfter ? text.substring(foundIndex + (((repeatZero > 1 || match_end_active) && match_end && isReverseSearch) ? 0 : modifiedDelimiter.length), text.length) : text.substring(0, foundIndex));
		}
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cARRAYTOTEXT() {
	}

	cARRAYTOTEXT.prototype = Object.create(cBaseFunction.prototype);
	cARRAYTOTEXT.prototype.constructor = cARRAYTOTEXT;
	cARRAYTOTEXT.prototype.name = 'ARRAYTOTEXT';
	cARRAYTOTEXT.prototype.isXLFN = true;
	cARRAYTOTEXT.prototype.argumentsMin = 1;
	cARRAYTOTEXT.prototype.argumentsMax = 2;
	cARRAYTOTEXT.prototype.arrayIndexes = {0: 1};
	cARRAYTOTEXT.prototype.argumentsType = [argType.reference, argType.number];
	cARRAYTOTEXT.prototype.Calculate = function (arg) {
		function arrayToTextGeneral(args, isRange) {
			let array = args[0],
				format = args[1];
			let resStr = "", arg0Dimensions;

			if (!format) {
				format = new cNumber(0);
			}

			if (cElementType.error === format.type) {
				return format;
			}

			format = format.tocNumber().getValue();

			if (format !== 0 && format !== 1) {
				return new cError(cErrorType.wrong_value_type);
			}
			// single val check
			if (cElementType.array !== array.type && cElementType.cellsRange !== array.type && cElementType.cellsRange3D !== array.type) {
				let tempArr = new cArray();
				tempArr.addElement(array);
				array = tempArr;
			}

			arg0Dimensions = array.getDimensions();

			for (let i = 0; i < arg0Dimensions.row; i++) {
				for (let j = 0; j < arg0Dimensions.col; j++) {
					let val = array.getValueByRowCol ? array.getValueByRowCol(i, j) : array.getElementRowCol(i, j);
					if (!val) {
						resStr += format === 1 ? "," : ", ";
						continue;
					}
					if (cElementType.string === val.type && format === 1) {
						val = '"' + val.getValue() + '"';
					} else if ((cElementType.cell === val.type || cElementType.cell3D === val.type) && format === 1) {
						let tempVal = val.getValue();
						if (cElementType.string === tempVal.type) {
							val = '"' + tempVal.getValue() + '"';
						} else {
							val = tempVal.getValue().toString();
						}
					} else {
						val = val.getValue().toString();
					}

					if (arg0Dimensions.col - 1 === j && format === 1) {
						resStr += val + ";";
						continue;
					}
					resStr += format === 1 ? val + "," : val + ", ";
				}
			}

			return format === 1 ? new cString("{" + resStr.slice(0, -1) + "}") : new cString(resStr.slice(0, -2));
		}

		let arg0 = arg[0],
			arg1 = arg[1] ? arg[1] : new cNumber(0),
			exceptions = new Map();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (cElementType.empty === arg0.type) {
			return new cError(cErrorType.wrong_value_type);
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			if (cElementType.empty === arg0.getValue().type) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (cElementType.array === arg0.type || cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			// skip checking this argument in helper function
			exceptions.set(0, true);
		}

		if (cElementType.array !== arg1.type && cElementType.cellsRange !== arg1.type && cElementType.cellsRange3D !== arg1.type) {
			// arg1 is not array/cellsRange
			return arrayToTextGeneral([arg0, arg1], false);
		} else {
			return AscCommonExcel.getArrayHelper([arg0, arg1], arrayToTextGeneral, exceptions);
		}
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cASC() {
	}

	cASC.prototype = Object.create(cBaseFunction.prototype);
	cASC.prototype.constructor = cASC;
	cASC.prototype.name = 'ASC';
	cASC.prototype.argumentsMin = 1;
	cASC.prototype.argumentsMax = 1;
	cASC.prototype.argumentsType = [argType.text];
	cASC.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		var calcAsc = function (str) {
			var res = '';
			var fullWidthFrom = 0xFF00;
			var fullWidthTo = 0xFFEF;

			for (var i = 0; i < str.length; i++) {
				var nCh = str[i].charCodeAt(0);
				if (nCh >= fullWidthFrom && nCh <= fullWidthTo) {
					nCh = 0xFF & (nCh + 0x20);
				}
				res += String.fromCharCode(nCh);
			}

			return new cString(res);
		};

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArray) {
			var ret = new cArray();
			arg0.foreach(function (elem, r, c) {
				if (!ret.array[r]) {
					ret.addRow();
				}

				if (elem instanceof cError) {
					ret.addElement(elem);
				} else {
					ret.addElement(calcAsc(elem.toLocaleString()));
				}
			});
			return ret;
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		return calcAsc(arg0.toLocaleString());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cBAHTTEXT() {
	}

	cBAHTTEXT.prototype = Object.create(cBaseFunction.prototype);
	cBAHTTEXT.prototype.constructor = cBAHTTEXT;
	cBAHTTEXT.prototype.name = 'BAHTTEXT';
	cBAHTTEXT.prototype.argumentsType = [argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCHAR() {
	}

	//***array-formula***
	cCHAR.prototype = Object.create(cBaseFunction.prototype);
	cCHAR.prototype.constructor = cCHAR;
	cCHAR.prototype.name = 'CHAR';
	cCHAR.prototype.argumentsMin = 1;
	cCHAR.prototype.argumentsMax = 1;
	cCHAR.prototype.argumentsType = [argType.number];
	cCHAR.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]).tocNumber();
		} else if (arg0 instanceof cArray) {
			var ret = new cArray();
			arg0.foreach(function (elem, r, c) {
				var _elem = elem.tocNumber();
				if (!ret.array[r]) {
					ret.addRow();
				}

				if (_elem instanceof cError) {
					ret.addElement(_elem);
				} else {
					ret.addElement(new cString(String.fromCharCode(_elem.getValue())));
				}
			});
			return ret;
		}

		arg0 = arg0.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}

		return new cString(String.fromCharCode(arg0.getValue()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCLEAN() {
	}

	//***array-formula***
	cCLEAN.prototype = Object.create(cBaseFunction.prototype);
	cCLEAN.prototype.constructor = cCLEAN;
	cCLEAN.prototype.name = 'CLEAN';
	cCLEAN.prototype.argumentsMin = 1;
	cCLEAN.prototype.argumentsMax = 1;
	cCLEAN.prototype.argumentsType = [argType.text];
	cCLEAN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		}
		if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}
		if (arg0 instanceof cError) {
			return arg0;
		}

		var v = arg0.toLocaleString(), l = v.length, res = "";

		for (var i = 0; i < l; i++) {
			if (v.charCodeAt(i) > 0x1f) {
				res += v[i];
			}
		}

		return new cString(res);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCODE() {
	}

	//***array-formula***
	cCODE.prototype = Object.create(cBaseFunction.prototype);
	cCODE.prototype.constructor = cCODE;
	cCODE.prototype.name = 'CODE';
	cCODE.prototype.argumentsMin = 1;
	cCODE.prototype.argumentsMax = 1;
	cCODE.prototype.argumentsType = [argType.text];
	cCODE.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArray) {
			var ret = new cArray();
			arg0.foreach(function (elem, r, c) {
				if (!ret.array[r]) {
					ret.addRow();
				}

				if (elem instanceof cError) {
					ret.addElement(elem);
				} else {
					ret.addElement(new cNumber(elem.toLocaleString().charCodeAt()));
				}
			});
			return ret;
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		return new cNumber(arg0.toLocaleString().charCodeAt());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCONCATENATE() {
	}

	//***array-formula***
	//TODO пересмотреть функцию!!!
	cCONCATENATE.prototype = Object.create(cBaseFunction.prototype);
	cCONCATENATE.prototype.constructor = cCONCATENATE;
	cCONCATENATE.prototype.name = 'CONCATENATE';
	cCONCATENATE.prototype.argumentsMin = 1;
	cCONCATENATE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCONCATENATE.prototype.argumentsType = [[argType.text]];
	cCONCATENATE.prototype.Calculate = function (arg) {
		var arg0 = new cString(""), argI;
		for (var i = 0; i < arg.length; i++) {
			argI = arg[i];
			if (argI instanceof cArea || argI instanceof cArea3D) {
				argI = argI.cross(arguments[1]);
			} else if (argI instanceof cRef || argI instanceof cRef3D) {
				argI = argI.getValue();
			}

			if (argI instanceof cError) {
				return argI;
			} else if (argI instanceof cArray) {
				argI.foreach(function (elem) {
					if (elem instanceof cError) {
						arg0 = elem;
						return true;
					}

					arg0 = new cString(arg0.toString().concat(elem.toLocaleString()));

				});
				if (arg0 instanceof cError) {
					return arg0;
				}
			} else {
				arg0 = new cString(arg0.toString().concat(argI.toLocaleString()));
			}
		}
		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCONCAT() {
	}

	//***array-formula***
	cCONCAT.prototype = Object.create(cBaseFunction.prototype);
	cCONCAT.prototype.constructor = cCONCAT;
	cCONCAT.prototype.name = 'CONCAT';
	cCONCAT.prototype.argumentsMin = 1;
	cCONCAT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cCONCAT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cCONCAT.prototype.argumentsType = [[argType.text]];
	cCONCAT.prototype.isXLFN = true;
	cCONCAT.prototype.Calculate = function (arg) {
		let arg0 = new cString(""), argI;


		let _checkMaxStringLength = function () {
			let maxStringLength = 32767;
			let _str = arg0 && arg0.toString && arg0.toString();
			if (_str && _str.length > maxStringLength) {
				return false;
			}
			return true;
		};

		for (let i = 0; i < arg.length; i++) {
			argI = arg[i];

			if (cElementType.cellsRange === argI.type || cElementType.cellsRange3D === argI.type) {
				let _arrVal = argI.getValue(this.checkExclude, this.excludeHiddenRows);
				for (let j = 0; j < _arrVal.length; j++) {
					let _arrElem = _arrVal[j].toLocaleString();
					if (cElementType.error === _arrElem.type) {
						return _arrVal[j];
					} else {
						arg0 = new cString(arg0.toString().concat(_arrElem));
					}
					if (!_checkMaxStringLength()) {
						return new cError(cErrorType.wrong_value_type)
					}
				}
			} else if (cElementType.cell === argI.type || cElementType.cell3D === argI.type) {
				argI = argI.getValue();
				if (cElementType.error === argI.type) {
					return argI;
				}
				arg0 = new cString(arg0.toString().concat(argI.toLocaleString()));
			} else {
				if (cElementType.error === argI.type) {
					return argI;
				} else if (cElementType.array === argI.type) {
					argI.foreach(function (elem) {
						if (cElementType.error === elem.type) {
							arg0 = elem;
							return true;
						}

						arg0 = new cString(arg0.toString().concat(elem.toLocaleString()));

						if (!_checkMaxStringLength()) {
							arg0 = new cError(cErrorType.wrong_value_type);
							return true;
						}
					});
					if (cElementType.error === arg0.type) {
						return arg0;
					}
				} else {
					arg0 = new cString(arg0.toString().concat(argI.toLocaleString()));
				}
			}
		}

		if (!_checkMaxStringLength()) {
			return new cError(cErrorType.wrong_value_type)
		}

		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDOLLAR() {
	}

	cDOLLAR.prototype = Object.create(cBaseFunction.prototype);
	cDOLLAR.prototype.constructor = cDOLLAR;
	cDOLLAR.prototype.name = 'DOLLAR';
	cDOLLAR.prototype.argumentsMin = 1;
	cDOLLAR.prototype.argumentsMax = 2;
	cDOLLAR.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDOLLAR.prototype.argumentsType = [argType.number, argType.number];
	cDOLLAR.prototype.Calculate = function (arg) {

		function SignZeroPositive(number) {
			return number < 0 ? -1 : 1;
		}

		function truncate(n) {
			return Math[n > 0 ? "floor" : "ceil"](n);
		}

		function Floor(number, significance) {
			let quotient = number / significance;
			if (quotient === 0) {
				return 0;
			}
			let nolpiat = 5 * Math.sign(quotient) *
				Math.pow(10, Math.floor(Math.log10(Math.abs(quotient))) - AscCommonExcel.cExcelSignificantDigits);
			return truncate(quotient + nolpiat) * significance;
		}

		function roundHelper(number, num_digits) {
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

			let significance = SignZeroPositive(number) * Math.pow(10, -truncate(num_digits));

			number += significance / 2;

			if (number / significance == Infinity) {
				return new cNumber(number);
			}

			return new cNumber(Floor(number, significance));
		}

		function toFix(str, skip) {
			let res, _int, _dec, _tmp = "";

			if (skip) {
				return str;
			}

			res = str.split(".");
			_int = res[0];

			if (res.length === 2) {
				_dec = res[1];
			}

			_int = _int.split("").reverse().join("").match(/([^]{1,3})/ig);

			for (let i = _int.length - 1; i >= 0; i--) {
				_tmp += _int[i].split("").reverse().join("");
				if (i != 0) {
					_tmp += ",";
				}
			}

			if (undefined != _dec) {
				while (_dec.length < arg1.getValue()) _dec += "0";
			}

			return "" + _tmp + (res.length === 2 ? "." + _dec + "" : "");
		}

		let arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(2), arg2 = arg[2] ? arg[2] : new cBool(false);

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
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

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() !== arg1.getCountElement() || arg0.getRowCount() !== arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					let a = elem;
					let b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber) {
						let res = roundHelper(a.getValue(), b.getValue());
						this.array[r][c] = toFix(res.toString(), arg2.toBool());
					} else {
						this.array[r][c] = new cError(cErrorType.wrong_value_type);
					}
				});
				return arg0;
			}
		} else if (arg0 instanceof cArray) {
			arg0.foreach(function (elem, r, c) {
				let a = elem;
				let b = arg1;
				if (a instanceof cNumber && b instanceof cNumber) {
					let res = roundHelper(a.getValue(), b.getValue());
					this.array[r][c] = toFix(res.toString(), arg2.toBool());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				let a = arg0;
				let b = elem;
				if (a instanceof cNumber && b instanceof cNumber) {
					let res = roundHelper(a.getValue(), b.getValue());
					this.array[r][c] = toFix(res.toString(), arg2.toBool());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		let number = arg0.getValue(), num_digits = arg1.getValue();

		let res = roundHelper(number, num_digits).getValue();

		let cNull = "";

		if (num_digits > 0) {
			cNull = ".";
			for (let i = 0; i < num_digits; i++, cNull += "0") {
			}
		}


		let format;
		let api = window["Asc"]["editor"];
		let nLocal = api && api.asc_getLocale();
		if (nLocal != null) {
			let info = new Asc.asc_CFormatCellsInfo();
			info.asc_setType(Asc.c_oAscNumFormatType.Currency);
			info.asc_setSymbol(nLocal);
			info.asc_setDecimalPlaces(num_digits);
			let arr = api.asc_getFormatCells(info);
			format = arr && arr[2];
		}
		if (!format) {
			format = "$#,##0" + cNull + ";($#,##0" + cNull + ")";
		}

		res = new cString(oNumFormatCache.get(format)
			.format(roundHelper(number, num_digits).getValue(), CellValueType.Number,
				AscCommon.gc_nMaxDigCount)[0].text);
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEXACT() {
	}

	cEXACT.prototype = Object.create(cBaseFunction.prototype);
	cEXACT.prototype.constructor = cEXACT;
	cEXACT.prototype.name = 'EXACT';
	cEXACT.prototype.argumentsMin = 2;
	cEXACT.prototype.argumentsMax = 2;
	cEXACT.prototype.argumentsType = [argType.text, argType.text];
	cEXACT.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		var arg0val = arg0.toLocaleString(), arg1val = arg1.toLocaleString();
		return new cBool(arg0val === arg1val);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFIND() {
	}

	//***array-formula***
	cFIND.prototype = Object.create(cBaseFunction.prototype);
	cFIND.prototype.constructor = cFIND;
	cFIND.prototype.name = 'FIND';
	cFIND.prototype.argumentsMin = 2;
	cFIND.prototype.argumentsMax = 3;
	cFIND.prototype.argumentsType = [argType.text, argType.text, argType.number];
	cFIND.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1], arg2 = arg.length === 3 ? arg[2] : null, res, str, searchStr,
			pos = -1;

		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0.type === cElementType.cell || arg0.type === cElementType.cell3D) {
			arg0 = arg0.getValue();
		}

		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.cellsRange3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1.type === cElementType.cell || arg1.type === cElementType.cell3D) {
			arg1 = arg1.getValue();
		}

		if (arg2 !== null) {

			if (arg2.type === cElementType.cellsRange || arg2.type === cElementType.cellsRange3D) {
				arg2 = arg2.cross(arguments[1]);
			}

			arg2 = arg2.tocNumber();
			if (arg2.type === cElementType.array) {
				arg2 = arg2.getElementRowCol(0, 0);
			}
			if (arg2.type === cElementType.error) {
				return arg2;
			}

			pos = arg2.getValue();
			pos = pos > 0 ? pos - 1 : pos;
		}

		if (arg0.type === cElementType.array) {
			arg0 = arg0.getElementRowCol(0, 0);
		}
		if (arg1.type === cElementType.array) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg0.type === cElementType.error) {
			return arg0;
		}
		if (arg1.type === cElementType.error) {
			return arg1;
		}

		str = arg1.toLocaleString();
		// searchStr = RegExp.escape(arg0.toLocaleString()); // doesn't work with strings like """ String""" , it's return ""\ String"" instead "" String""
		//TODO need review. bugs 50869; 68343
		searchStr = arg0.toLocaleString().replace(/\"\"/g, "\"");
		searchStr = RegExp.escape(searchStr);

		if (arg2) {

			if (pos > str.length || pos < 0) {
				return new cError(cErrorType.wrong_value_type);
			}

			str = str.substring(pos);
			res = str.search(searchStr);
			if (res >= 0) {
				res += pos;
			}
		} else {
			res = str.search(searchStr);
		}

		if (res < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		return new cNumber(res + 1);
	};

	/**
	 * @constructor
	 * @extends {cFIND}
	 */
	function cFINDB() {
	}

	//***array-formula***
	cFINDB.prototype = Object.create(cFIND.prototype);
	cFINDB.prototype.constructor = cFINDB;
	cFINDB.prototype.name = 'FINDB';
	cFINDB.prototype.argumentsType = [argType.text, argType.text, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFIXED() {
	}

	cFIXED.prototype = Object.create(cBaseFunction.prototype);
	cFIXED.prototype.constructor = cFIXED;
	cFIXED.prototype.name = 'FIXED';
	cFIXED.prototype.argumentsMin = 1;
	cFIXED.prototype.argumentsMax = 3;
	cFIXED.prototype.argumentsType = [argType.number, argType.number, argType.logical];
	cFIXED.prototype.Calculate = function (arg) {

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
				Math.pow(10, Math.floor(Math.log10(Math.abs(quotient))) - AscCommonExcel.cExcelSignificantDigits);
			return truncate(quotient + nolpiat) * significance;
		}

		function roundHelper(number, num_digits) {
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

			var significance = SignZeroPositive(number) * Math.pow(10, -truncate(num_digits));

			number += significance / 2;

			if (number / significance == Infinity) {
				return new cNumber(number);
			}

			return new cNumber(Floor(number, significance));
		}

		function toFix(str, skip) {
			var res, _int, _dec, _tmp = "";

			if (skip) {
				return str;
			}

			res = str.split(".");
			_int = res[0];

			if (res.length == 2) {
				_dec = res[1];
			}

			_int = _int.split("").reverse().join("").match(/([^]{1,3})/ig);

			for (var i = _int.length - 1; i >= 0; i--) {
				_tmp += _int[i].split("").reverse().join("");
				if (i != 0) {
					_tmp += ",";
				}
			}

			if (undefined != _dec) {
				while (_dec.length < arg1.getValue()) _dec += "0";
			}

			return "" + _tmp + (res.length == 2 ? "." + _dec + "" : "");
		}

		var arg0 = arg[0], arg1 = arg[1] ? arg[1] : new cNumber(2), arg2 = arg[2] ? arg[2] : new cBool(false);

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		}

		arg2 = arg2.tocBool();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
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

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			if (arg0.getCountElement() != arg1.getCountElement() || arg0.getRowCount() != arg1.getRowCount()) {
				return new cError(cErrorType.not_available);
			} else {
				arg0.foreach(function (elem, r, c) {
					var a = elem;
					var b = arg1.getElementRowCol(r, c);
					if (a instanceof cNumber && b instanceof cNumber && arg2.toBool) {
						var res = roundHelper(a.getValue(), b.getValue());
						this.array[r][c] = toFix(res.toString(), arg2.toBool());
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
				if (a instanceof cNumber && b instanceof cNumber && arg2.toBool) {
					var res = roundHelper(a.getValue(), b.getValue());
					this.array[r][c] = toFix(res.toString(), arg2.toBool());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg0;
		} else if (arg1 instanceof cArray) {
			arg1.foreach(function (elem, r, c) {
				var a = arg0;
				var b = elem;
				if (a instanceof cNumber && b instanceof cNumber && arg2.toBool) {
					var res = roundHelper(a.getValue(), b.getValue());
					this.array[r][c] = toFix(res.toString(), arg2.toBool());
				} else {
					this.array[r][c] = new cError(cErrorType.wrong_value_type);
				}
			});
			return arg1;
		}

		var number = arg0.getValue(), num_digits = arg1.getValue();

		var cNull = "";

		if (num_digits > 0) {
			cNull = ".";
			for (var i = 0; i < num_digits; i++, cNull += "0") {
			}
		}
		if (!arg2.toBool) {
			return new cError(cErrorType.wrong_value_type);
		}
		return new cString(oNumFormatCache.get("#" + (arg2.toBool() ? "" : ",") + "##0" + cNull)
			.format(roundHelper(number, num_digits).getValue(), CellValueType.Number,
				AscCommon.gc_nMaxDigCount)[0].text)
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cIMPORTRANGE() {
	}

	cIMPORTRANGE.prototype = Object.create(cBaseFunction.prototype);
	cIMPORTRANGE.prototype.constructor = cIMPORTRANGE;
	cIMPORTRANGE.prototype.name = 'IMPORTRANGE';
	cIMPORTRANGE.prototype.argumentsMin = 2;
	cIMPORTRANGE.prototype.argumentsMax = 2;
	cIMPORTRANGE.prototype.isXLUDF = true;
	cIMPORTRANGE.prototype.argumentsType = [argType.text, argType.text];
	cIMPORTRANGE.prototype.Calculate = function (arg) {
		//gs -> allow array(get first element), cRef, cRef3D, cName, cName3d
		//not allow area/area3d
		//if first argument - link, when - text only inside " "
		//if second argument - link, when - text only without " "

		let arg0 = arg[0], arg1 = arg[1];

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type || cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			return new cError(cErrorType.not_available);
		}


		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			arg0 = arg0.getValue();
		}
		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
			arg1 = arg1.getValue();
		}

		arg0 = arg0.tocString();
		arg1 = arg1.tocString();

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		//check valid arguments strings
		let sArg1 = arg1.toString();
		let is3DRef = AscCommon.parserHelp.parse3DRef(sArg1);
		if (!is3DRef) {
			let _range = AscCommonExcel.g_oRangeCache.getAscRange(sArg1);
			if (_range) {
				is3DRef = {sheet: null, sheet2: null, range: sArg1};
			}
		}

		if (!is3DRef) {
			return new cError(cErrorType.not_available);
		}

		if (AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks) {
			if (!AscCommonExcel.importRangeLinksState.importRangeLinks) {
				AscCommonExcel.importRangeLinksState.importRangeLinks = {};
			}
			let linkName = arg0.toString();
			if (!AscCommonExcel.importRangeLinksState.importRangeLinks[linkName]) {
				AscCommonExcel.importRangeLinksState.importRangeLinks[linkName] = [];
			}
			AscCommonExcel.importRangeLinksState.importRangeLinks[linkName].push(is3DRef);
		}

		let checkHyperlink = function (string) {
			if (!string) {
				return false;
			}
			if (typeof string !== 'string') {
				return false;
			}

			// protocol check
			let protocols = ['http://', 'https://', 'ftp://'];
			let hasValidProtocol = false;

			for (let i = 0; i < protocols.length; i++) {
				if (string.indexOf(protocols[i]) === 0) {
					hasValidProtocol = true;
					break;
				}
			}

			// domain check
			let hasDomain = string.indexOf('.') !== -1 && string.charAt(0) !== '.';
			return hasValidProtocol && hasDomain;
		}

		let api = window["Asc"]["editor"];
		let wb = api && api.wbModel;
		let eR = wb && wb.getExternalLinkByName(arg0.toString());
		let ret = new cArray();
		if (eR) {
			if (!is3DRef.sheet) {
				is3DRef.sheet = eR.SheetNames[0];
			}
			let externalWs = wb.getExternalWorksheetByName(eR.Id, is3DRef.sheet);
			if (externalWs) {
				let bbox = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(is3DRef.range)[0];

				if (is3DRef.sheet !== undefined) {
					let index = eR.getSheetByName(is3DRef.sheet);
					if (index != null) {
						let externalSheetDataSet = eR.SheetDataSet[index];
						for (let i = bbox.r1; i <= bbox.r2; i++) {
							let row = externalSheetDataSet.getRow(i + 1);
							if (!row) {
								continue;
							}
							if (!ret.array[i - bbox.r1]) {
								ret.addRow();
							}
							for (let j = bbox.c1; j <= bbox.c2; j++) {
								let externalCell = row.getCell(j);
								if (externalCell) {
									let cellValue = externalCell.getFormulaValue();
									if (cellValue && cellValue.value && checkHyperlink(cellValue.value)) {
										cellValue.hyperlink = cellValue.value;
									}
									ret.addElement(cellValue);
								}
							}
						}
					}
				}
				return ret;
			}
		}
		return new cError(cErrorType.bad_reference);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cJIS() {
	}

	cJIS.prototype = Object.create(cBaseFunction.prototype);
	cJIS.prototype.constructor = cJIS;
	cJIS.prototype.name = 'JIS';
	cJIS.prototype.argumentsMin = 1;
	cJIS.prototype.argumentsMax = 1;
	cJIS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		var calc = function (str) {
			var res = '';
			var fullWidthFrom = 0xFF00;
			var fullWidthTo = 0xFFEF;

			for (var i = 0; i < str.length; i++) {
				var nCh = str[i].charCodeAt(0);
				if (!(nCh >= fullWidthFrom && nCh <= fullWidthTo)) {
					nCh = nCh - 0x20 + 0xff00;
				}
				res += String.fromCharCode(nCh);
			}

			return new cString(res);
		};

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArray) {
			var ret = new cArray();
			arg0.foreach(function (elem, r, c) {
				if (!ret.array[r]) {
					ret.addRow();
				}

				if (elem instanceof cError) {
					ret.addElement(elem.toLocaleString());
				} else {
					ret.addElement(calc(elem));
				}
			});
			return ret;
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		return calc(arg0.toLocaleString());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLEFT() {
	}

	cLEFT.prototype = Object.create(cBaseFunction.prototype);
	cLEFT.prototype.constructor = cLEFT;
	cLEFT.prototype.name = 'LEFT';
	cLEFT.prototype.argumentsMin = 1;
	cLEFT.prototype.argumentsMax = 2;
	cLEFT.prototype.argumentsType = [argType.text, argType.number];
	cLEFT.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg.length == 1 ? new cNumber(1) : arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		arg0 = arg0.tocString();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg1.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		return new cString(arg0.getValue().substring(0, arg1.getValue()))

	};

	/**
	 * @constructor
	 * @extends {cLEFT}
	 */
	function cLEFTB() {
	}

	cLEFTB.prototype = Object.create(cLEFT.prototype);
	cLEFTB.prototype.constructor = cLEFTB;
	cLEFTB.prototype.name = 'LEFTB';
	cLEFTB.prototype.argumentsType = [argType.text, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLEN() {
	}

	//***array-formula***
	cLEN.prototype = Object.create(cBaseFunction.prototype);
	cLEN.prototype.constructor = cLEN;
	cLEN.prototype.name = 'LEN';
	cLEN.prototype.argumentsMin = 1;
	cLEN.prototype.argumentsMax = 1;
	cLEN.prototype.argumentsType = [argType.text];
	cLEN.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		return new cNumber(arg0.toLocaleString().length)
	};

	/**
	 * @constructor
	 * @extends {cLEN}
	 */
	function cLENB() {
	}

	//***array-formula***
	cLENB.prototype = Object.create(cLEN.prototype);
	cLENB.prototype.constructor = cLENB;
	cLENB.prototype.name = 'LENB';
	cLENB.prototype.argumentsType = [argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLOWER() {
	}

	//***array-formula***
	cLOWER.prototype = Object.create(cBaseFunction.prototype);
	cLOWER.prototype.constructor = cLOWER;
	cLOWER.prototype.name = 'LOWER';
	cLOWER.prototype.argumentsMin = 1;
	cLOWER.prototype.argumentsMax = 1;
	cLOWER.prototype.argumentsType = [argType.text];
	cLOWER.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
			if (arg0 instanceof cError) {
				return arg0;
			} else {
				arg0 = arg0.toLocaleString();
			}
		} else {
			arg0 = arg0.toLocaleString();
		}

		return new cString(arg0.toLowerCase());

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMID() {
	}

	//***array-formula***
	cMID.prototype = Object.create(cBaseFunction.prototype);
	cMID.prototype.constructor = cMID;
	cMID.prototype.name = 'MID';
	cMID.prototype.argumentsMin = 3;
	cMID.prototype.argumentsMax = 3;
	cMID.prototype.argumentsType = [argType.text, argType.number, argType.number];
	cMID.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}
		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]);
		}

		arg0 = arg0.toLocaleString();
		arg1 = arg1.tocNumber();
		arg2 = arg2.tocNumber();

		if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}
		if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}
		if (arg2 instanceof cArray) {
			arg2 = arg2.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}
		if (arg1.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}
		if (arg2.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		var l = arg0.length;

		if (arg1.getValue() > l) {
			return new cString("");
		}

		return new cString(arg0.substr(arg1.getValue() == 0 ? 0 : arg1.getValue() - 1, arg2.getValue()));
	};

	/**
	 * @constructor
	 * @extends {cMID}
	 */
	function cMIDB() {
	}

	//***array-formula***
	cMIDB.prototype = Object.create(cMID.prototype);
	cMIDB.prototype.constructor = cMIDB;
	cMIDB.prototype.name = 'MIDB';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cNUMBERVALUE() {
	}

	//***array-formula***
	cNUMBERVALUE.prototype = Object.create(cBaseFunction.prototype);
	cNUMBERVALUE.prototype.constructor = cNUMBERVALUE;
	cNUMBERVALUE.prototype.name = 'NUMBERVALUE';
	cNUMBERVALUE.prototype.argumentsMin = 1;
	cNUMBERVALUE.prototype.argumentsMax = 3;
	cNUMBERVALUE.prototype.isXLFN = true;
	cNUMBERVALUE.prototype.argumentsType = [argType.text, argType.text, argType.text];
	cNUMBERVALUE.prototype.Calculate = function (arg) {

		var oArguments = this._prepareArguments(arg, arguments[1], true);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocString();
		argClone[1] =
			argClone[1] ? argClone[1].tocString() : new cString(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);
		argClone[2] =
			argClone[2] ? argClone[2].tocString() : new cString(AscCommon.g_oDefaultCultureInfo.NumberGroupSeparator);

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		var replaceAt = function (str, index, chr) {
			if (index > str.length - 1) {
				return str;
			} else {
				return str.substr(0, index) + chr + str.substr(index + 1);
			}
		};

		var calcText = function (argArray) {
			var aInputString = argArray[0];
			var aDecimalSeparator = argArray[1], aGroupSeparator = argArray[2];
			var cDecimalSeparator = 0;

			if (aDecimalSeparator) {
				if (aDecimalSeparator.length === 1) {
					cDecimalSeparator = aDecimalSeparator[0];
				} else {
					return new cError(cErrorType.wrong_value_type);
				}
			}

			if (cDecimalSeparator && aGroupSeparator && aGroupSeparator.indexOf(cDecimalSeparator) !== -1) {
				return new cError(cErrorType.wrong_value_type)
			}

			if (aInputString.length === 0) {
				return new cError(cErrorType.wrong_value_type)
			}

			//считаем количество вхождений cDecimalSeparator в строке
			var count = 0;
			for (var i = 0; i < aInputString.length; i++) {
				if (cDecimalSeparator === aInputString[i]) {
					count++;
				}
				if (count > 1) {
					return new cError(cErrorType.wrong_value_type);
				}
			}

			var nDecSep = cDecimalSeparator ? aInputString.indexOf(cDecimalSeparator) : 0;
			if (nDecSep !== 0) {
				var aTemporary = nDecSep >= 0 ? aInputString.substr(0, nDecSep) : aInputString;

				var nIndex = 0;
				while (nIndex < aGroupSeparator.length) {
					var nChar = aGroupSeparator[nIndex];

					aTemporary = aTemporary.replace(new RegExp(RegExp.escape(nChar), "g"), "");
					nIndex++;
				}

				if (nDecSep >= 0) {
					aInputString = aTemporary + aInputString.substr(nDecSep);
				} else {
					aInputString = aTemporary;
				}
			}

			//replace decimal separator
			aInputString =
				aInputString.replace(cDecimalSeparator, AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);

			//delete spaces
			aInputString = aInputString.replace(/(\s|\r|\t|\n)/g, "");
			/*for ( var i = aInputString.length; i >= 0; i--){
			 var c = aInputString.charCodeAt(i);
			 if ( c == 0x0020 || c == 0x0009 || c == 0x000A || c == 0x000D ){
			 aInputString = aInputString.replaceAt( i, 1, "" ); // remove spaces etc.
			 }
			 }*/

			//remove and count '%'
			var nPercentCount = 0;
			for (var i = aInputString.length - 1; i >= 0 && aInputString.charCodeAt(i) === 0x0025; i--) {
				aInputString = replaceAt(aInputString, i, "");
				nPercentCount++;
			}

			var fVal = AscCommon.g_oFormatParser.parse(aInputString, AscCommon.g_oDefaultCultureInfo);
			if (fVal) {
				fVal = fVal.value;
				if (nPercentCount) {
					fVal *= Math.pow(10, -(nPercentCount * 2));
				}

				return new cNumber(fVal);
			}

			return new cError(cErrorType.wrong_value_type)
		};

		return this._findArrayInNumberArguments(oArguments, calcText, true);
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cPHONETIC() {
	}

	cPHONETIC.prototype = Object.create(cBaseFunction.prototype);
	cPHONETIC.prototype.constructor = cPHONETIC;
	cPHONETIC.prototype.name = 'PHONETIC';

	//

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cPROPER() {
	}

	//***array-formula***
	cPROPER.prototype = Object.create(cBaseFunction.prototype);
	cPROPER.prototype.constructor = cPROPER;
	cPROPER.prototype.name = 'PROPER';
	cPROPER.prototype.argumentsMin = 1;
	cPROPER.prototype.argumentsMax = 1;
	cPROPER.prototype.argumentsType = [argType.text];
	cPROPER.prototype.Calculate = function (arg) {
		var reg_PROPER = new RegExp("[-#$+*/^&%<\\[\\]='?_\\@!~`\">: ;.\\)\\(,]|\\d|\\s"), arg0 = arg[0];

		function proper(str) {
			var canUpper = true, retStr = "", regTest;
			for (var i = 0; i < str.length; i++) {
				regTest = reg_PROPER.test(str[i]);

				if (regTest) {
					canUpper = true;
				} else {
					if (canUpper) {
						retStr += str[i].toUpperCase();
						canUpper = false;
						continue;
					}
				}

				retStr += str[i].toLowerCase();

			}
			return retStr;
		}

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArray) {
			var ret = new cArray();
			arg0.foreach(function (elem, r, c) {
				if (!ret.array[r]) {
					ret.addRow();
				}

				if (elem instanceof cError) {
					ret.addElement(elem);
				} else {
					ret.addElement(new cString(proper(elem.toLocaleString())));
				}
			});
			return ret;
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		return new cString(proper(arg0.toLocaleString()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cREPLACE() {
	}

	//***array-formula***
	cREPLACE.prototype = Object.create(cBaseFunction.prototype);
	cREPLACE.prototype.constructor = cREPLACE;
	cREPLACE.prototype.name = 'REPLACE';
	cREPLACE.prototype.argumentsMin = 4;
	cREPLACE.prototype.argumentsMax = 4;
	cREPLACE.prototype.argumentsType = [argType.text, argType.number, argType.number, argType.text];
	cREPLACE.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2], arg3 = arg[3];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]).tocString();
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0).tocString();
		}

		arg0 = arg0.tocString();

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]).tocNumber();
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElement(0).tocNumber();
		}

		arg1 = arg1.tocNumber();

		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]).tocNumber();
		} else if (arg2 instanceof cArray) {
			arg2 = arg2.getElement(0).tocNumber();
		}

		arg2 = arg2.tocNumber();

		if (arg3 instanceof cArea || arg3 instanceof cArea3D) {
			arg3 = arg3.cross(arguments[1]).tocString();
		} else if (arg3 instanceof cArray) {
			arg3 = arg3.getElement(0).tocString();
		}

		arg3 = arg3.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}
		if (arg3 instanceof cError) {
			return arg3;
		}

		if (arg1.getValue() < 1 || arg2.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		var string1 = arg0.getValue(), string2 = arg3.getValue(), res = "";

		string1 = string1.split("");
		string1.splice(arg1.getValue() - 1, arg2.getValue(), string2);
		for (var i = 0; i < string1.length; i++) {
			res += string1[i];
		}

		return new cString(res);

	};

	/**
	 * @constructor
	 * @extends {cREPLACE}
	 */
	function cREPLACEB() {
	}

	//***array-formula***
	cREPLACEB.prototype = Object.create(cREPLACE.prototype);
	cREPLACEB.prototype.constructor = cREPLACEB;
	cREPLACEB.prototype.name = 'REPLACEB';
	cREPLACEB.prototype.argumentsType = [argType.text, argType.number, argType.number, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cREPT() {
	}

	cREPT.prototype = Object.create(cBaseFunction.prototype);
	cREPT.prototype.constructor = cREPT;
	cREPT.prototype.name = 'REPT';
	cREPT.prototype.argumentsMin = 2;
	cREPT.prototype.argumentsMax = 2;
	cREPT.prototype.argumentsType = [argType.text, argType.number];
	cREPT.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], res = "";
		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}


		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocString();
		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]).tocNumber();
		} else if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
		}

		if (arg1 instanceof cError) {
			return arg1;
		} else if (arg1 instanceof cString) {
			return new cError(cErrorType.wrong_value_type);
		} else {
			arg1 = arg1.tocNumber();
		}

		if (arg1.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		return new cString(arg0.getValue().repeat(arg1.getValue()));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cRIGHT() {
	}

	//***array-formula***
	cRIGHT.prototype = Object.create(cBaseFunction.prototype);
	cRIGHT.prototype.constructor = cRIGHT;
	cRIGHT.prototype.name = 'RIGHT';
	cRIGHT.prototype.argumentsMin = 1;
	cRIGHT.prototype.argumentsMax = 2;
	cRIGHT.prototype.argumentsType = [argType.text, argType.number];
	cRIGHT.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg.length === 1 ? new cNumber(1) : arg[1];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		}

		arg0 = arg0.tocString();
		arg1 = arg1.tocNumber();

		if (arg0 instanceof cArray && arg1 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
			arg1 = arg1.getElementRowCol(0, 0);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (arg1.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}
		var l = arg0.getValue().length, _number = l - arg1.getValue();
		return new cString(arg0.getValue().substring(_number < 0 ? 0 : _number, l))

	};

	/**
	 * @constructor
	 * @extends {cRIGHT}
	 */
	function cRIGHTB() {
	}

	//***array-formula***
	cRIGHTB.prototype = Object.create(cRIGHT.prototype);
	cRIGHTB.prototype.constructor = cRIGHTB;
	cRIGHTB.prototype.name = 'RIGHTB';
	cRIGHTB.prototype.argumentsType = [argType.text, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSEARCH() {
	}

	//***array-formula***
	cSEARCH.prototype = Object.create(cBaseFunction.prototype);
	cSEARCH.prototype.constructor = cSEARCH;
	cSEARCH.prototype.name = 'SEARCH';
	cSEARCH.prototype.argumentsMin = 2;
	cSEARCH.prototype.argumentsMax = 3;
	cSEARCH.prototype.arrayIndexes = {0: 1, 1: 1, 2: 1};
	cSEARCH.prototype.argumentsType = [argType.text, argType.text, argType.number];
	cSEARCH.prototype.Calculate = function (arg) {

		const searchString = function (find_text, within_text, start_num) {
			if (start_num < 1 || start_num > within_text.length) {
				return new cError(cErrorType.wrong_value_type);
			}

			let valueForSearching = find_text
				.replace(/(\\)/g, "\\\\")
				.replace(/(\^)/g, "\\^")
				.replace(/(\()/g, "\\(")
				.replace(/(\))/g, "\\)")
				.replace(/(\+)/g, "\\+")
				.replace(/(\[)/g, "\\[")
				.replace(/(\])/g, "\\]")
				.replace(/(\{)/g, "\\{")
				.replace(/(\})/g, "\\}")
				.replace(/(\$)/g, "\\$")
				.replace(/(\.)/g, "\\.")
				.replace(/(~)?\*/g, function ($0, $1) {
					return $1 ? $0 : '(.*)';
				})
				.replace(/(~)?\?/g, function ($0, $1) {
					return $1 ? $0 : '.';
				})
				.replace(/(~\*)/g, "\\*").replace(/(~\?)/g, "\\?");
			valueForSearching = new RegExp(valueForSearching, "ig");
			if ('' === find_text) {
				return new cNumber(start_num);
			}

			let res = within_text.substring(start_num - 1).search(valueForSearching);
			if (res < 0) {
				return new cError(cErrorType.wrong_value_type);
			}

			res += start_num - 1;

			return new cNumber(res + 1);
		}

		const searchInArray = function (arr, findText, withinText, startNum) {
			findText = findText ? findText.tocString() : findText;
			withinText = withinText ? withinText.tocString() : withinText;
			startNum = startNum ? startNum.tocNumber() : startNum;

			arr.foreach(function (elem, r, c) {
				if (!resArr.array[r]) {
					resArr.addRow();
				}

				let item = startNum ? elem.tocString() : elem.tocNumber();
				if (findText && findText.type === cElementType.error) {
					resArr.addElement(findText);
				} else if (withinText && withinText.type === cElementType.error) {
					resArr.addElement(withinText);
				} else if (startNum && startNum.type === cElementType.error) {
					resArr.addElement(startNum);
				} else if (item && item.type === cElementType.error) {
					resArr.addElement(item);
				} else {
					let res = searchString(findText ? findText.getValue() : item.getValue(), withinText ? withinText.getValue() : item.getValue(), startNum ? startNum.getValue() : item.getValue());
					resArr.addElement(res);
				}
			})

			return resArr;
		}

		const t = this;
		let arg0 = arg[0] ? arg[0] : new cEmpty(), arg1 = arg[1] ? arg[1] : new cEmpty(),
			arg2 = arg[2] ? arg[2] : new cNumber(1);

		if (arg0.type === cElementType.cellsRange || arg0.type === cElementType.cellsRange3D) {
			arg0 = arg0.cross(arguments[1]).tocString();
		}

		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.cellsRange3D) {
			arg1 = arg1.cross(arguments[1]).tocString();
		}

		if (arg2.type === cElementType.cellsRange || arg2.type === cElementType.cellsRange3D) {
			arg2 = arg2.cross(arguments[1]).tocNumber();
		}

		let resArr = new cArray();
		if ((arg0.type === cElementType.array && arg1.type === cElementType.array) || (arg0.type === cElementType.array && arg2.type === cElementType.array) || (arg1.type === cElementType.array && arg2.type === cElementType.array)) {
			let findTextArrDimensions = arg0.getDimensions(),
				withinTextArrDimensions = arg1.getDimensions(),
				startNumDimensions = arg2.getDimensions(),
				resCols = Math.max(findTextArrDimensions.col, withinTextArrDimensions.col, startNumDimensions.col),
				resRows = Math.max(findTextArrDimensions.row, withinTextArrDimensions.row, startNumDimensions.row);

			if (arg0.type !== cElementType.array) {
				let tempArg0 = new cArray();
				tempArg0.addElement(arg0);
				arg0 = tempArg0;
			}
			if (arg1.type !== cElementType.array) {
				let tempArg1 = new cArray();
				tempArg1.addElement(arg1);
				arg1 = tempArg1;
			}
			if (arg2.type !== cElementType.array) {
				let tempArg2 = new cArray();
				tempArg2.addElement(arg2);
				arg2 = tempArg2;
			}

			for (let i = 0; i < resRows; i++) {
				resArr.addRow();
				for (let j = 0; j < resCols; j++) {
					let findText, withinText, startNum;
					// get the substring that we will look for
					if ((findTextArrDimensions.col - 1 < j && findTextArrDimensions.col > 1) || (findTextArrDimensions.row - 1 < i && findTextArrDimensions.row > 1)) {
						findText = new cError(cErrorType.not_available);
						resArr.addElement(findText);
						continue;
					} else if (findTextArrDimensions.row === 1 && findTextArrDimensions.col === 1) {
						// get first elem
						findText = arg0.getElementRowCol ? arg0.getElementRowCol(0, 0) : arg0.getValueByRowCol(0, 0);
					} else if (findTextArrDimensions.row === 1) {
						// get elem from first row
						findText = arg0.getElementRowCol ? arg0.getElementRowCol(0, j) : arg0.getValueByRowCol(0, j);
					} else if (findTextArrDimensions.col === 1) {
						// get elem from first col
						findText = arg0.getElementRowCol ? arg0.getElementRowCol(i, 0) : arg0.getValueByRowCol(i, 0);
					} else {
						findText = arg0.getElementRowCol ? arg0.getElementRowCol(i, j) : arg0.getValueByRowCol(i, j);
					}

					// get the string that we will search in
					if ((withinTextArrDimensions.col - 1 < j && withinTextArrDimensions.col > 1) || (withinTextArrDimensions.row - 1 < i && withinTextArrDimensions.row > 1)) {
						withinText = new cError(cErrorType.not_available);
						resArr.addElement(withinText);
						continue;
					} else if (withinTextArrDimensions.row === 1 && withinTextArrDimensions.col === 1) {
						// get first elem
						withinText = arg1.getElementRowCol ? arg1.getElementRowCol(0, 0) : arg1.getValueByRowCol(0, 0);
					} else if (withinTextArrDimensions.row === 1) {
						// get elem from first row
						withinText = arg1.getElementRowCol ? arg1.getElementRowCol(0, j) : arg1.getValueByRowCol(0, j);
					} else if (withinTextArrDimensions.col === 1) {
						// get elem from first col
						withinText = arg1.getElementRowCol ? arg1.getElementRowCol(i, 0) : arg1.getValueByRowCol(i, 0);
					} else {
						withinText = arg1.getElementRowCol ? arg1.getElementRowCol(i, j) : arg1.getValueByRowCol(i, j);
					}

					// get the start num that we will start search
					if ((startNumDimensions.col - 1 < j && startNumDimensions.col > 1) || (startNumDimensions.row - 1 < i && startNumDimensions.row > 1)) {
						startNum = new cError(cErrorType.not_available);
						resArr.addElement(startNum);
						continue;
					} else if (startNumDimensions.row === 1 && startNumDimensions.col === 1) {
						// get first elem
						startNum = arg2.getElementRowCol ? arg2.getElementRowCol(0, 0) : arg2.getValueByRowCol(0, 0);
					} else if (startNumDimensions.row === 1) {
						// get elem from first row
						startNum = arg2.getElementRowCol ? arg2.getElementRowCol(0, j) : arg2.getValueByRowCol(0, j);
					} else if (startNumDimensions.col === 1) {
						// get elem from first col
						startNum = arg2.getElementRowCol ? arg2.getElementRowCol(i, 0) : arg2.getValueByRowCol(i, 0);
					} else {
						startNum = arg2.getElementRowCol ? arg2.getElementRowCol(i, j) : arg2.getValueByRowCol(i, j);
					}

					// check errors
					findText = findText ? findText.tocString() : new cString("");
					withinText = withinText ? withinText.tocString() : new cString("");
					startNum = startNum ? startNum.tocNumber() : new cNumber(0);

					if (findText.type === cElementType.error) {
						resArr.addElement(findText);
						continue
					}
					if (withinText.type === cElementType.error) {
						resArr.addElement(withinText);
						continue
					}
					if (startNum.type === startNum.error) {
						resArr.addElement(startNum);
						continue
					}

					let res = searchString(findText.getValue(), withinText.getValue(), startNum.getValue());
					resArr.addElement(res);
				}
			}

			return resArr;
		} else if (arg0.type === cElementType.array) {
			return searchInArray(arg0, null, arg1, arg2);
		} else if (arg1.type === cElementType.array) {
			return searchInArray(arg1, arg0, null, arg2);
		} else if (arg2.type === cElementType.array) {
			return searchInArray(arg2, arg0, arg1, null);
		}

		arg0 = arg0.tocString();
		arg1 = arg1.tocString();
		arg2 = arg2.tocNumber();

		if (arg0.type === cElementType.error) {
			return arg0;
		}
		if (arg1.type === cElementType.error) {
			return arg1;
		}
		if (arg2.type === cElementType.error) {
			return arg2;
		}

		return searchString(arg0.getValue(), arg1.getValue(), arg2.getValue());

	};

	/**
	 * @constructor
	 * @extends {cRIGHT}
	 */
	function cSEARCHB() {
	}

	//***array-formula***
	cSEARCHB.prototype = Object.create(cSEARCH.prototype);
	cSEARCHB.prototype.constructor = cSEARCHB;
	cSEARCHB.prototype.name = 'SEARCHB';
	cSEARCHB.prototype.argumentsType = [argType.text, argType.text, argType.number];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSUBSTITUTE() {
	}

	//***array-formula***
	cSUBSTITUTE.prototype = Object.create(cBaseFunction.prototype);
	cSUBSTITUTE.prototype.constructor = cSUBSTITUTE;
	cSUBSTITUTE.prototype.name = 'SUBSTITUTE';
	cSUBSTITUTE.prototype.argumentsMin = 3;
	cSUBSTITUTE.prototype.argumentsMax = 4;
	cSUBSTITUTE.prototype.argumentsType = [argType.text, argType.text, argType.text, argType.text];
	cSUBSTITUTE.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2], arg3 = arg[3] ? arg[3] : new cNumber(0);

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]).tocString();
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0).tocString();
		}

		arg0 = arg0.tocString();

		if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]).tocString();
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElement(0).tocString();
		}

		arg1 = arg1.tocString();

		if (arg2 instanceof cArea || arg2 instanceof cArea3D) {
			arg2 = arg2.cross(arguments[1]).tocString();
		} else if (arg2 instanceof cArray) {
			arg2 = arg2.getElement(0).tocString();
		}

		arg2 = arg2.tocString();

		if (arg3 instanceof cArea || arg3 instanceof cArea3D) {
			arg3 = arg3.cross(arguments[1]).tocNumber();
		} else if (arg3 instanceof cArray) {
			arg3 = arg3.getElement(0).tocNumber();
		}

		arg3 = arg3.tocNumber();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}
		if (arg2 instanceof cError) {
			return arg2;
		}
		if (arg3 instanceof cError) {
			return arg3;
		}

		if (arg3.getValue() < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		var string = arg0.getValue(), old_string = arg1.getValue(), new_string = arg2.getValue(),
			occurence = arg3.getValue(), index = 0, res;
		res = string.replace(new RegExp(RegExp.escape(old_string), "g"), function (equal, p1, source) {
			index++;
			if (occurence == 0 || occurence > source.length) {
				return new_string;
			} else if (occurence == index) {
				return new_string;
			}
			return equal;
		});

		return new cString(res);

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cT() {
	}

	//***array-formula***
	cT.prototype = Object.create(cBaseFunction.prototype);
	cT.prototype.constructor = cT;
	cT.prototype.name = 'T';
	cT.prototype.argumentsMin = 1;
	cT.prototype.argumentsMax = 1;
	cT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.replace_only_array;
	cT.prototype.argumentsType = [argType.any];
	cT.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cString || arg0 instanceof cError) {
			return arg0;
		} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			return arg0.getValue2(0, 0);
		} else if (arg[0] instanceof cArray) {
			arg0 = arg[0].getElementRowCol(0, 0);
		}

		if (arg0 instanceof cString || arg0 instanceof cError) {
			return arg[0];
		} else {
			return new cString("");
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTEXT() {
	}

	//***array-formula***
	cTEXT.prototype = Object.create(cBaseFunction.prototype);
	cTEXT.prototype.constructor = cTEXT;
	cTEXT.prototype.name = 'TEXT';
	cTEXT.prototype.argumentsMin = 2;
	cTEXT.prototype.argumentsMax = 2;
	cTEXT.prototype.argumentsType = [argType.any, argType.text];
	cTEXT.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		} else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg1 instanceof cRef || arg1 instanceof cRef3D) {
			arg1 = arg1.getValue();
		} else if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
			arg1 = arg1.cross(arguments[1]);
		} else if (arg1 instanceof cArray) {
			arg1 = arg1.getElementRowCol(0, 0);
		}

		arg1 = arg1.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}
		if (arg1 instanceof cError) {
			return arg1;
		}

		if (!(arg0 instanceof cBool)) {
			var _tmp = arg0.tocNumber();
			if (_tmp instanceof cNumber) {
				arg0 = _tmp;
			}
		}

		var oFormat = new CellFormat(arg1.toString(), undefined, true);
		var a = g_oFormatParser.parse(arg0.toLocaleString(true) + ""), aText;
		aText = oFormat.format(a ? a.value : arg0.toLocaleString(),
			(arg0 instanceof cNumber || a) ? CellValueType.Number : CellValueType.String,
			AscCommon.gc_nMaxDigCountView);
		var text = "";

		for (var i = 0, length = aText.length; i < length; ++i) {

			if (aText[i].format && aText[i].format.getSkip()) {
				text += " ";
				continue;
			}
			if (aText[i].format && aText[i].format.getRepeat()) {
				continue;
			}

			text += aText[i].text;
		}

		var res = new cString(text);
		res.numFormat = AscCommonExcel.cNumFormatNone;
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTEXTJOIN() {
	}

	//***array-formula***
	cTEXTJOIN.prototype = Object.create(cBaseFunction.prototype);
	cTEXTJOIN.prototype.constructor = cTEXTJOIN;
	cTEXTJOIN.prototype.name = 'TEXTJOIN';
	cTEXTJOIN.prototype.argumentsMin = 3;
	cTEXTJOIN.prototype.argumentsMax = 255;
	cTEXTJOIN.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTEXTJOIN.prototype.isXLFN = true;
	cTEXTJOIN.prototype.argumentsType = [argType.text, argType.logical, argType.text, [argType.text]];
	//TODO все, кроме 2 аргумента - массивы
	cTEXTJOIN.prototype.arrayIndexes = {0: 1, 2: 1, 3: 1, 4: 1, 5: 1, 6: 1, 7: 1, 8: 1};
	cTEXTJOIN.prototype.getArrayIndex = function (index) {
		if (index === 1) {
			return undefined;
		}
		return 1;
	};
	cTEXTJOIN.prototype.Calculate = function (arg) {

		var argClone = [arg[0], arg[1]];
		argClone[1] = arg[1].tocBool();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		var ignore_empty = argClone[1].toBool();
		var delimiter = argClone[0];
		var delimiterIter = 0;
		//разделитель может быть в виде массива, где используются все его элементы
		var delimiterArr = this._getOneDimensionalArray(delimiter);
		//если хотя бы один элемент ошибка, то возвращаем ошибку
		if (delimiterArr instanceof cError) {
			return delimiterArr;
		}

		var concatString = function (string1, string2) {
			var res = string1;
			if ("" === string2 && ignore_empty) {
				return res;
			}
			var isStartStr = string1 === "";
			//выбираем разделитель из массива по порядку
			var delimiterStr = isStartStr ? "" : delimiterArr[delimiterIter];
			if (undefined === delimiterStr) {
				delimiterIter = 0;
				delimiterStr = delimiterArr[delimiterIter];
			}
			if (!isStartStr) {
				delimiterIter++;
			}

			res += delimiterStr + string2;
			return res;
		};

		var arg0 = new cString(""), argI;
		for (var i = 2; i < arg.length; i++) {
			argI = arg[i];

			var type = argI.type;
			if (cElementType.cellsRange === type || cElementType.cellsRange3D === type || cElementType.array === type) {
				//получаем одномерный массив
				argI = this._getOneDimensionalArray(argI);

				//если хотя бы один элемент с ошибкой, возвращаем ошибку
				if (argI instanceof cError) {
					return argI;
				}

				for (var n = 0; n < argI.length; n++) {
					arg0 = new cString(concatString(arg0.toString(), argI[n].toLocaleString()));
				}

			} else if (cElementType.cell === type || cElementType.cell3D === type) {
				argI = argI.getValue();

				if (argI instanceof cError) {
					return argI;
				}

				arg0 = new cString(concatString(arg0.toString(), argI.toLocaleString()));
			} else {

				if (argI instanceof cError) {
					return argI;
				}

				arg0 = new cString(concatString(arg0.toString(), argI.toLocaleString()));
			}
		}

		return arg0;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTRIM() {
	}

	//***array-formula***
	cTRIM.prototype = Object.create(cBaseFunction.prototype);
	cTRIM.prototype.constructor = cTRIM;
	cTRIM.prototype.name = 'TRIM';
	cTRIM.prototype.argumentsMin = 1;
	cTRIM.prototype.argumentsMax = 1;
	cTRIM.prototype.argumentsType = [argType.text];
	cTRIM.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]).tocString();
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0).tocString();
		}

		arg0 = arg0.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}

		return new cString(arg0.getValue().replace(AscCommon.rx_space_g, function ($0, $1, $2) {
			var res = " " === $2[$1 + 1] ? "" : $2[$1];
			return res;
		}).replace(/^ | $/g, ""))
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cUNICHAR() {
	}

	//***array-formula***
	cUNICHAR.prototype = Object.create(cBaseFunction.prototype);
	cUNICHAR.prototype.constructor = cUNICHAR;
	cUNICHAR.prototype.name = 'UNICHAR';
	cUNICHAR.prototype.argumentsMin = 1;
	cUNICHAR.prototype.argumentsMax = 1;
	cUNICHAR.prototype.isXLFN = true;
	cUNICHAR.prototype.argumentsType = [argType.number];
	cUNICHAR.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocNumber();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function _func(argArray) {
			var num = parseInt(argArray[0]);
			if (isNaN(num) || num <= 0 || num > 1114111) {
				return new cError(cErrorType.wrong_value_type);
			}

			var res = String.fromCharCode(num);
			if ("" === res) {
				return new cError(cErrorType.wrong_value_type);
			}


			return new cString(res);
		}

		return this._findArrayInNumberArguments(oArguments, _func, true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cUNICODE() {
	}

	//***array-formula***
	cUNICODE.prototype = Object.create(cBaseFunction.prototype);
	cUNICODE.prototype.constructor = cUNICODE;
	cUNICODE.prototype.name = 'UNICODE';
	cUNICODE.prototype.argumentsMin = 1;
	cUNICODE.prototype.argumentsMax = 1;
	cUNICODE.prototype.isXLFN = true;
	cUNICODE.prototype.argumentsType = [argType.text];
	cUNICODE.prototype.Calculate = function (arg) {
		var oArguments = this._prepareArguments(arg, arguments[1]);
		var argClone = oArguments.args;

		argClone[0] = argClone[0].tocString();

		var argError;
		if (argError = this._checkErrorArg(argClone)) {
			return argError;
		}

		function _func(argArray) {
			var str = argArray[0].toLocaleString();
			var res = str.charCodeAt(0);
			return new cNumber(res);
		}

		return this._findArrayInNumberArguments(oArguments, _func, true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cUPPER() {
	}

	//***array-formula***
	cUPPER.prototype = Object.create(cBaseFunction.prototype);
	cUPPER.prototype.constructor = cUPPER;
	cUPPER.prototype.name = 'UPPER';
	cUPPER.prototype.argumentsMin = 1;
	cUPPER.prototype.argumentsMax = 1;
	cUPPER.prototype.argumentsType = [argType.text];
	cUPPER.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		}
		if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		if (arg0 instanceof cError) {
			return arg0;
		}

		if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
			if (arg0 instanceof cError) {
				return arg0;
			} else {
				arg0 = arg0.toLocaleString();
			}
		} else {
			arg0 = arg0.toLocaleString();
		}

		return new cString(arg0.toUpperCase());
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cVALUE() {
	}

	//***array-formula***
	cVALUE.prototype = Object.create(cBaseFunction.prototype);
	cVALUE.prototype.constructor = cVALUE;
	cVALUE.prototype.name = 'VALUE';
	cVALUE.prototype.argumentsMin = 1;
	cVALUE.prototype.argumentsMax = 1;
	cVALUE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cVALUE.prototype.argumentsType = [argType.any];
	cVALUE.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}

		arg0 = arg0.tocString();

		if (arg0 instanceof cError) {
			return arg0;
		}

		var res = g_oFormatParser.parse(arg0.getValue());

		if (res) {
			return new cNumber(res.value);
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTEXTBEFORE() {
	}

	//***array-formula***
	cTEXTBEFORE.prototype = Object.create(cBaseFunction.prototype);
	cTEXTBEFORE.prototype.constructor = cTEXTBEFORE;
	cTEXTBEFORE.prototype.name = 'TEXTBEFORE';
	cTEXTBEFORE.prototype.argumentsMin = 2;
	cTEXTBEFORE.prototype.argumentsMax = 6;
	cTEXTBEFORE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTEXTBEFORE.prototype.argumentsType = [argType.text, argType.text, argType.number, argType.number, argType.number, argType.any];
	cTEXTBEFORE.prototype.isXLFN = true;
	cTEXTBEFORE.prototype.arrayIndexes = {1: 1};
	cTEXTBEFORE.prototype.Calculate = function (arg) {
		return calcBeforeAfterText(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTEXTAFTER() {
	}

	//***array-formula***
	cTEXTAFTER.prototype = Object.create(cBaseFunction.prototype);
	cTEXTAFTER.prototype.constructor = cTEXTAFTER;
	cTEXTAFTER.prototype.name = 'TEXTAFTER';
	cTEXTAFTER.prototype.argumentsMin = 2;
	cTEXTAFTER.prototype.argumentsMax = 6;
	cTEXTAFTER.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTEXTAFTER.prototype.argumentsType = [argType.text, argType.text, argType.number, argType.number, argType.number, argType.any];
	cTEXTAFTER.prototype.isXLFN = true;
	cTEXTAFTER.prototype.arrayIndexes = {1: 1};
	cTEXTAFTER.prototype.Calculate = function (arg) {
		return calcBeforeAfterText(arg, arguments[1], true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTEXTSPLIT() {
	}

	//***array-formula***
	cTEXTSPLIT.prototype = Object.create(cBaseFunction.prototype);
	cTEXTSPLIT.prototype.constructor = cTEXTSPLIT;
	cTEXTSPLIT.prototype.name = 'TEXTSPLIT';
	cTEXTSPLIT.prototype.argumentsMin = 2;
	cTEXTSPLIT.prototype.argumentsMax = 6;
	cTEXTSPLIT.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	//cTEXTSPLIT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cTEXTSPLIT.prototype.arrayIndexes = {1: 1, 2: 1, 4: 1, 5: 1};
	cTEXTSPLIT.prototype.argumentsType = [argType.text, argType.text, argType.text, argType.logical, argType.logical, argType.any];
	cTEXTSPLIT.prototype.isXLFN = true;
	cTEXTSPLIT.prototype.Calculate = function (arg) {

		//функция должна возвращать массив
		let text = arg[0];
		if (text.type === cElementType.error) {
			return text;
		}

		//второй/третий аргумент тоже может быть массивом, каждый из элементов каторого может быть разделителем
		let col_delimiter = arg[1];
		let row_delimiter = arg[2] ? arg[2] : null;

		//если оба empty или хотя бы один из разделителей - пустая строка - ошибка
		if (col_delimiter && row_delimiter && col_delimiter.type === cElementType.empty && row_delimiter.type === cElementType.empty) {
			return new cError(cErrorType.wrong_value_type);
		}
		if (col_delimiter && col_delimiter.type === cElementType.string && col_delimiter.getValue() === "") {
			return new cError(cErrorType.wrong_value_type);
		}
		if (row_delimiter && row_delimiter.type === cElementType.string && row_delimiter.getValue() === "") {
			return new cError(cErrorType.wrong_value_type);
		}

		col_delimiter = col_delimiter.toArray(true, true);
		if (col_delimiter.type === cElementType.error) {
			return col_delimiter;
		}

		if (row_delimiter) {
			row_delimiter = row_delimiter.toArray(true, true);
			if (row_delimiter.type === cElementType.error) {
				return row_delimiter;
			}
		}

		let ignore_empty = arg[3] ? arg[3].tocBool() : new cBool(false);
		if (ignore_empty.type === cElementType.cellsRange3D || ignore_empty.type === cElementType.cellsRange || ignore_empty.type === cElementType.array) {
			ignore_empty = ignore_empty.getValue2(0, 0);
			ignore_empty = ignore_empty.tocBool();
		}
		if (ignore_empty.type === cElementType.error) {
			return ignore_empty;
		}
		ignore_empty = ignore_empty.toBool();

		let match_mode = arg[4] ? arg[4].tocBool() : new cBool(false);
		if (match_mode.type === cElementType.cellsRange3D || match_mode.type === cElementType.cellsRange || match_mode.type === cElementType.array) {
			match_mode = match_mode.getValue2(0, 0);
			match_mode = match_mode.tocBool();
		}
		if (match_mode.type === cElementType.error) {
			return match_mode;
		}
		match_mode = match_mode.toBool();

		//заполняющее_значение. Значение по умолчанию: #Н/Д.
		let pad_with = arg[5] ? arg[5] : new cError(cErrorType.not_available);
		if (pad_with.type === cElementType.cell3D || pad_with.type === cElementType.cell) {
			pad_with = pad_with.getValue();
		}
		if (pad_with.type === cElementType.cellsRange3D || pad_with.type === cElementType.cellsRange) {
			return new cError(cErrorType.wrong_value_type);
		}

		let getRexExpFromArray = function (_array, _match_mode) {
			let sRegExp = "";
			if (Array.isArray(_array)) {
				for (let row = 0; row < _array.length; row++) {
					for (let col = 0; col < _array[row].length; col++) {
						if (sRegExp !== "") {
							sRegExp += "|";
						}

						sRegExp += AscCommon.escapeRegExp(_array[row][col] + "");
					}
				}
			} else {
				sRegExp += "[" + AscCommon.escapeRegExp(_array + "") + "]";
			}

			return _match_mode ? new RegExp(sRegExp, "i") : new RegExp(sRegExp);
		};

		let splitText = function (_text, _rowDelimiter, _colDelimiter) {
			var res;

			if (_rowDelimiter == null || _rowDelimiter === "" || _rowDelimiter && _rowDelimiter[0] === "" || _rowDelimiter && _rowDelimiter[0] && _rowDelimiter[0][0] === "") {
				_rowDelimiter = null;
			} else {
				_rowDelimiter = getRexExpFromArray(_rowDelimiter, match_mode);
			}
			if (_colDelimiter === "" || _colDelimiter && _colDelimiter[0] === "" || _colDelimiter && _colDelimiter[0] && _colDelimiter[0][0] === "") {
				_colDelimiter = null;
			} else {
				_colDelimiter = getRexExpFromArray(_colDelimiter, match_mode);
			}

			var _array = _text.split(_rowDelimiter);
			if (_array) {
				for (let i = 0; i < _array.length; i++) {
					if (!res) {
						res = [];
					}
					res.push(_array[i].split(_colDelimiter));
				}
			}

			return res;
		};

		//обрабатываю первый аргумент - диапазон выше, если сюда он приходит в виже диапазона, то беру первый элемент
		var res;
		if (cElementType.cellsRange3D === text.type || cElementType.cellsRange === text.type) {
			text = text.getValue2(0, 0);
		} else if (cElementType.array === text.type) {
			text = text.getValue2(0, 0);
		} else if (text.type === cElementType.cell3D || text.type === cElementType.cell) {
			text = text.getValue();
		}

		text = text.tocString();
		if (text.type === cElementType.error) {
			return text;
		}
		text = text.toString();
		if (match_mode) {
			text = text.toLowerCase();
		}

		//let array = AscCommon.parseText(text, options);
		let array = splitText(text, row_delimiter, col_delimiter);
		if (array) {
			//проверяем массив на пустые элементы +  дополняем массив pad_with

			let rowCount = array.length;
			let colCount = 0, i, j;
			for (i = 0; i < rowCount; i++) {
				colCount = Math.max(colCount, array[i].length)
			}

			let newArray = [];
			for (i = 0; i < rowCount; i++) {
				let row = [];
				for (j = 0; j < colCount; j++) {
					if (("" === array[i][j] && !ignore_empty) || array[i][j]) {
						row.push(new cString(array[i][j]));
					}
				}

				if (row.length || (!row.length && !ignore_empty)) {
					if (colCount > row.length) {
						while (colCount > row.length) {
							row.push(pad_with);
						}
					}
					newArray.push(row);
				}
			}

			res = new cArray();
			res.fillFromArray(newArray);
		}

		return res && res.isValidArray() ? res : new cError(cErrorType.not_available);
	};

	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].cTEXT = cTEXT;
})(window);
