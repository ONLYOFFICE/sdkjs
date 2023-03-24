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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	var g_cCharDelimiter = AscCommon.g_cCharDelimiter;
	var parserHelp = AscCommon.parserHelp;
	var gc_nMaxRow0 = AscCommon.gc_nMaxRow0;
	var gc_nMaxCol0 = AscCommon.gc_nMaxCol0;
	var g_oCellAddressUtils = AscCommon.g_oCellAddressUtils;
	var CellAddress = AscCommon.CellAddress;

	var cElementType = AscCommonExcel.cElementType;
	var cErrorType = AscCommonExcel.cErrorType;
	var cNumber = AscCommonExcel.cNumber;
	var cString = AscCommonExcel.cString;
	var cBool = AscCommonExcel.cBool;
	var cError = AscCommonExcel.cError;
	var cArea = AscCommonExcel.cArea;
	var cArea3D = AscCommonExcel.cArea3D;
	var cRef = AscCommonExcel.cRef;
	var cRef3D = AscCommonExcel.cRef3D;
	var cName = AscCommonExcel.cName;
	var cName3D = AscCommonExcel.cName3D;
	var cEmpty = AscCommonExcel.cEmpty;
	var cArray = AscCommonExcel.cArray;
	var cBaseFunction = AscCommonExcel.cBaseFunction;

	var checkTypeCell = AscCommonExcel.checkTypeCell;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var argType = Asc.c_oAscFormulaArgumentType;

	var _func = AscCommonExcel._func;

	cFormulaFunctionGroup['LookupAndReference'] = cFormulaFunctionGroup['LookupAndReference'] || [];
	cFormulaFunctionGroup['LookupAndReference'].push(cADDRESS, cAREAS, cCHOOSE, cCHOOSECOLS, cCHOOSEROWS, cCOLUMN, cCOLUMNS, cDROP, cEXPAND, cFILTER, cFORMULATEXT,
		cGETPIVOTDATA, cHLOOKUP, cHYPERLINK, cINDEX, cINDIRECT, cLOOKUP, cMATCH, cOFFSET, cROW, cROWS, cSORT, cRTD, cTRANSPOSE, cTAKE,
		cUNIQUE, cVLOOKUP, cXLOOKUP, cVSTACK, cHSTACK, cTOROW, cTOCOL, cWRAPROWS, cWRAPCOLS, cXMATCH);

	cFormulaFunctionGroup['NotRealised'] = cFormulaFunctionGroup['NotRealised'] || [];
	cFormulaFunctionGroup['NotRealised'].push(cAREAS, cGETPIVOTDATA, cRTD);

	function searchRegExp(str, flags) {
		var vFS = str
			.replace(/(\\)/g, "\\")
			.replace(/(\^)/g, "\\^")
			.replace(/(\()/g, "\\(")
			.replace(/(\))/g, "\\)")
			.replace(/(\+)/g, "\\+")
			.replace(/(\[)/g, "\\[")
			.replace(/(\])/g, "\\]")
			.replace(/(\{)/g, "\\{")
			.replace(/(\})/g, "\\}")
			.replace(/(\$)/g, "\\$")
			.replace(/(~)?\*/g, function ($0, $1) {
				return $1 ? $0 : '(.*)';
			})
			.replace(/(~)?\?/g, function ($0, $1) {
				return $1 ? $0 : '.{1}';
			})
			.replace(/(~\*)/g, "\\*").replace(/(~\?)/g, "\\?");

		return new RegExp(vFS + "$", flags ? flags : "i");
	}

	// TODO переделать поиск: добиться такого же результата как в ms
	function XBinarySearch (target, array, match_mode, isReverse) {
		let mid, item;
		let index = -1;
		let leftPointer = 0,
			rightPointer = array.length - 1;

		function bidirectionalSearch (secondIteration, isDescending) {
			while(leftPointer <= rightPointer) {
				mid = Math.floor((leftPointer + rightPointer) / 2);
				item = undefined !== array[mid].v ? array[mid].v : array[mid];
				if(-1 === match_mode || 0 === match_mode || 1 === match_mode) {
					if (leftPointer === 0 && rightPointer === array.length - 1 && !secondIteration) {
						if ((isDescending && target >= item) || (!isDescending && target <= item)) {
							rightPointer = mid - 1;
						} else {
							leftPointer = mid + 1;
						}
					} else if(!secondIteration) {
						if (target == item) {
							index = mid;
							break;
						} else if ((isDescending && target > item) || (!isDescending && target < item)) {
							rightPointer = mid - 1;
						} else {
							leftPointer = mid + 1;
						}
					}

					if(secondIteration) {
						// exact
						if (0 === match_mode) {
							if(item == target) {
								index = mid;
								break;
							} else if(item < target) {
								if (isDescending) {
									rightPointer = mid - 1;
								} else {
									leftPointer = mid + 1;
								}
							} else {
								if (isDescending) {
									leftPointer = mid + 1
								} else {
									rightPointer = mid - 1;
								}
							}
						}

						// exact or larger
						if (1 === match_mode) {
							if (leftPointer === 0 && rightPointer === array.length - 1) {
								if (item == target) {
									index = mid;
									break;
								} else if (item < target) {
									if (isDescending) {
										rightPointer = mid - 1;
									} else {
										leftPointer = mid + 1;
									}
								} else {
									if (isDescending) {
										leftPointer = mid + 1;
									} else {
										rightPointer = mid - 1;
									}
								}
							} else {
								if (item > target || item == target) {
									index = mid;
									break;
								} else {
									if (isDescending) {
										rightPointer = mid - 1;
									} else {
										leftPointer = mid + 1;
									}
								}
							}
						}

						// exact or smaller
						if (-1 === match_mode) {
							if (leftPointer === 0 && rightPointer === array.length - 1) {
								if(item == target) {
									index = mid;
									break;
								} else if (item < target) {
									if (isDescending) {
										rightPointer = mid - 1;
									} else {
										leftPointer = mid + 1;
									}
								} else {
									if (isDescending) {
										leftPointer = mid + 1;
									} else {
										rightPointer = mid - 1;
									}
								}
							} else {
								if (item < target || item == target) {
									index = mid;
									break;
								} else {
									if (isDescending) {
										leftPointer = mid + 1;
									} else {
										rightPointer = mid - 1;
									}
								}
							}
						}
					}
				}
			}
		}

		bidirectionalSearch(false, isReverse)

		// second iteration
		if(index === -1 && (1 === match_mode || -1 === match_mode || 0 === match_mode)) {
			leftPointer = 0;
			rightPointer = array.length - 1;
			bidirectionalSearch(true, isReverse);
		}

		return index;	
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cADDRESS() {
	}

	//***array-formula***
	cADDRESS.prototype = Object.create(cBaseFunction.prototype);
	cADDRESS.prototype.constructor = cADDRESS;
	cADDRESS.prototype.name = 'ADDRESS';
	cADDRESS.prototype.argumentsMin = 2;
	cADDRESS.prototype.argumentsMax = 5;
	cADDRESS.prototype.argumentsType = [argType.number, argType.number, argType.number, argType.logical, argType.text];
	cADDRESS.prototype.Calculate = function (arg) {
		var rowNumber = arg[0], colNumber = arg[1], refType = arg[2] ? arg[2] : new cNumber(1),
			A1RefType = arg[3] ? arg[3] : new cBool(true), sheetName = arg[4] ? arg[4] : null;

		if (cElementType.cellsRange === rowNumber.type || cElementType.cellsRange3D === rowNumber.type) {
			rowNumber = rowNumber.cross(arguments[1]);
		} else if (cElementType.array === rowNumber.type) {
			rowNumber = rowNumber.getElementRowCol(0, 0);
		}

		if (cElementType.cellsRange === colNumber.type || cElementType.cellsRange3D === colNumber.type) {
			colNumber = colNumber.cross(arguments[1]);
		} else if (cElementType.array === colNumber.type) {
			colNumber = colNumber.getElementRowCol(0, 0);
		}

		if (cElementType.cellsRange === refType.type || cElementType.cellsRange3D === refType.type) {
			refType = refType.cross(arguments[1]);
		} else if (cElementType.array === refType.type) {
			refType = refType.getElementRowCol(0, 0);
		} else if(cElementType.empty === refType.type) {
			refType = new cNumber(1);
		}

		if (cElementType.cellsRange === A1RefType.type || cElementType.cellsRange3D === A1RefType.type) {
			A1RefType = A1RefType.cross(arguments[1]);
		} else if (cElementType.array === A1RefType.type) {
			A1RefType = A1RefType.getElementRowCol(0, 0);
		}  else if(cElementType.empty === A1RefType.type) {
			A1RefType = new cNumber(1);
		}

		if(sheetName){
			if (cElementType.cellsRange === sheetName.type || cElementType.cellsRange3D === sheetName.type) {
				sheetName = sheetName.cross(arguments[1]);
			} else if (cElementType.array === sheetName.type) {
				sheetName = sheetName.getElementRowCol(0, 0);
			} else if (cElementType.cell === sheetName.type || cElementType.cell3D === sheetName.type) {
				sheetName = sheetName.getValue();
			} else if (cElementType.empty === sheetName.type) {
				sheetName = null;
			}
		}

		rowNumber = rowNumber.tocNumber();
		colNumber = colNumber.tocNumber();
		refType = refType.tocNumber();
		A1RefType = A1RefType.tocBool();

		if (cElementType.error === rowNumber.type) {
			return rowNumber;
		}
		if (cElementType.error === colNumber.type) {
			return colNumber;
		}
		if (cElementType.error === refType.type) {
			return refType;
		}
		if (cElementType.error === A1RefType.type) {
			return A1RefType;
		}
		if (sheetName && cElementType.error === sheetName.type) {
			return sheetName;
		}

		rowNumber = rowNumber.getValue();
		colNumber = colNumber.getValue();
		refType = refType.getValue();
		A1RefType = A1RefType.toBool();

		rowNumber = parseInt(rowNumber);
		colNumber = parseInt(colNumber);

		if (refType > 4 || refType < 1 || rowNumber < 1 || rowNumber > AscCommon.gc_nMaxRow || colNumber < 1 ||
			colNumber > AscCommon.gc_nMaxCol) {
			return new cError(cErrorType.wrong_value_type);
		}
		var strRef;
		var absR, absC;
		switch (refType - 1) {
			case AscCommonExcel.referenceType.A:
				absR = true;
				absC = true;
				break;
			case AscCommonExcel.referenceType.ARRC:
				absR = true;
				absC = false;
				break;
			case AscCommonExcel.referenceType.RRAC:
				absR = false;
				absC = true;
				break;
			case AscCommonExcel.referenceType.R:
				absR = false;
				absC = false;
				break;
		}

		strRef = this._getRef(this._absolute(absR, rowNumber, A1RefType),
			this._absolute(absC, A1RefType ? g_oCellAddressUtils.colnumToColstrFromWsView(colNumber) : colNumber,
				A1RefType), A1RefType);

		var res = strRef;
		if(sheetName){
			if("" === sheetName.getValue()){
				res = "!" + strRef;
			} else {
				res = parserHelp.get3DRef(sheetName.toString(), strRef);
			}
		}

		return new cString(res);
	};
	cADDRESS.prototype._getRef = function (row, col, A1RefType) {
		return A1RefType ? col + row : 'R' + row + 'C' + col;
	};
	cADDRESS.prototype._absolute = function (abs, val, A1RefType) {

		return abs ? (A1RefType ? '$' + val : val) : (A1RefType ? val : '[' + val + ']');
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cAREAS() {
	}

	cAREAS.prototype = Object.create(cBaseFunction.prototype);
	cAREAS.prototype.constructor = cAREAS;
	cAREAS.prototype.name = 'AREAS';
	cAREAS.prototype.argumentsType = [argType.reference];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCHOOSE() {
	}

	//***array-formula***
	cCHOOSE.prototype = Object.create(cBaseFunction.prototype);
	cCHOOSE.prototype.constructor = cCHOOSE;
	cCHOOSE.prototype.name = 'CHOOSE';
	cCHOOSE.prototype.argumentsMin = 2;
	cCHOOSE.prototype.argumentsMax = 30;
	cCHOOSE.prototype.argumentsType = [argType.number, [argType.any]];
	cCHOOSE.prototype.Calculate = function (arg) {
		var arg0 = arg[0];

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.cross(arguments[1]);
		}
		arg0 = arg0.tocNumber();

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		if (cElementType.number === arg0.type) {
			if (arg0.getValue() < 1 || arg0.getValue() > arg.length - 1) {
				return new cError(cErrorType.wrong_value_type);
			}

			return arg[Math.floor(arg0.getValue())];
		}

		return new cError(cErrorType.wrong_value_type);
	};

	function chooseRowsCols(arg, argument1, byCol) {
		var argError = cBaseFunction.prototype._checkErrorArg.call(this, arg);
		if (argError) {
			return argError;
		}

		let arg1 = arg[0];
		let matrix;
		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.array || arg1.type === cElementType.cell || arg1.type === cElementType.cell3D) {
			matrix = arg1.getMatrix();
		} else if (arg1.type === cElementType.cellsRange3D) {
			if (arg1.isSingleSheet()) {
				matrix = arg1.getMatrix()[0];
			} else {
				return new cError(cErrorType.bad_reference);
			}
		} else if (arg1.type === cElementType.error) {
			return arg1;
		} else if (arg1.type === cElementType.empty) {
			return new cError(cErrorType.wrong_value_type);
		} else {
			matrix = [[arg1]];
		}

		let pushData = function (_argInside) {
			_argInside = _argInside.tocNumber();
			if (_argInside.type === cElementType.error) {
				error = _argInside;
				return false;
			}
			_argInside = _argInside.toNumber();
			let reverse = _argInside < 0;
			_argInside = Math.abs(_argInside);
			_argInside = parseInt(_argInside);
			if (_argInside < 1 || (_argInside > dimension.col && byCol) || (_argInside > dimension.row && !byCol)) {
				error = new cError(cErrorType.wrong_value_type);
				return false;
			}

			if (!res) {
				res = new cArray();
			}

			if (byCol) {
				res.pushCol(matrix, reverse ? dimension.col - (_argInside - 1) - 1 : _argInside - 1);
			} else {
				res.pushRow(matrix, reverse ? dimension.row - (_argInside - 1) - 1 : _argInside - 1);
			}

			return true;
		};

		let dimension = arg1.getDimensions();
		let res;
		let error;
		for (let i = 1; i < arg.length; i++) {
			let _arg = arg[i];

			if (cElementType.cellsRange === _arg.type || cElementType.cellsRange3D === _arg.type || cElementType.array === _arg.type) {
				let argDimensions = _arg.getDimensions();
				if (argDimensions.col === 1 || argDimensions.row === 1) {
					let byCol = argDimensions.row > 1;
					for (let j = 0; j < Math.max(argDimensions.col, argDimensions.row); j++) {
						if (cElementType.array === _arg.type) {
							if (!pushData(_arg.getElementRowCol(!byCol ? 0 : j, !byCol ? j : 0))) {
								return error;
							}
						} else {
							if (!pushData(_arg.getValue2(!byCol ? 0 : j, !byCol ? j : 0))) {
								return error;
							}
						}
					}
				} else {
					return new cError(cErrorType.wrong_value_type);
				}

				continue;
			}

			if (!pushData(_arg)) {
				return error;
			}
		}

		return res ? res : new cError(cErrorType.wrong_value_type);
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCHOOSECOLS() {
	}

	//***array-formula***
	cCHOOSECOLS.prototype = Object.create(cBaseFunction.prototype);
	cCHOOSECOLS.prototype.constructor = cCHOOSECOLS;
	cCHOOSECOLS.prototype.name = 'CHOOSECOLS';
	cCHOOSECOLS.prototype.argumentsMin = 2;
	cCHOOSECOLS.prototype.argumentsMax = 253;
	cCHOOSECOLS.prototype.argumentsType = [argType.reference, [argType.number]];
	cCHOOSECOLS.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cCHOOSECOLS.prototype.isXLFN = true;
	cCHOOSECOLS.prototype.Calculate = function (arg) {
		return chooseRowsCols(arg, arguments[1], true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCHOOSEROWS() {
	}

	//***array-formula***
	cCHOOSEROWS.prototype = Object.create(cBaseFunction.prototype);
	cCHOOSEROWS.prototype.constructor = cCHOOSEROWS;
	cCHOOSEROWS.prototype.name = 'CHOOSEROWS';
	cCHOOSEROWS.prototype.argumentsMin = 2;
	cCHOOSEROWS.prototype.argumentsMax = 253;
	cCHOOSEROWS.prototype.argumentsType = [argType.reference, [argType.number]];
	cCHOOSEROWS.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cCHOOSEROWS.prototype.isXLFN = true;
	cCHOOSEROWS.prototype.Calculate = function (arg) {
		return chooseRowsCols(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOLUMN() {
	}

	//***array-formula***
	cCOLUMN.prototype = Object.create(cBaseFunction.prototype);
	cCOLUMN.prototype.constructor = cCOLUMN;
	cCOLUMN.prototype.name = 'COLUMN';
	cCOLUMN.prototype.argumentsMax = 1;
	cCOLUMN.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.setArrayRefAsArg;
	cCOLUMN.prototype.argumentsType = [argType.reference];
	cCOLUMN.prototype.Calculate = function (arg) {
		var bbox;
		var res;
		var opt_col = arguments[6];
		if (opt_col !== undefined) {
			return new cNumber(opt_col + 1);
		} else if (0 === arg.length) {
			bbox = arguments[1];
			res = bbox ? new cNumber(bbox.c1 + 1) : null;
		} else {
			var arg0 = arg[0];
			if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
				bbox = arg0.getRange();
				bbox = bbox && bbox.bbox;
				res = bbox ? new cNumber(bbox.c1 + 1) : null;
			} else if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
				bbox = arg0.getRange();
				bbox = bbox && bbox.bbox;

				if (bbox && bbox.c2 > bbox.c1) {
					res = new cArray();
					for (var i = bbox.c1; i <= bbox.c2; i++) {
						res.addElement(new cNumber(i + 1))
					}
				} else {
					res = bbox ? new cNumber(bbox.c1 + 1) : null;
				}
			}
		}
		return res ? res : new cError(cErrorType.bad_reference);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCOLUMNS() {
	}

	//***array-formula***
	cCOLUMNS.prototype = Object.create(cBaseFunction.prototype);
	cCOLUMNS.prototype.constructor = cCOLUMNS;
	cCOLUMNS.prototype.name = 'COLUMNS';
	cCOLUMNS.prototype.argumentsMin = 1;
	cCOLUMNS.prototype.argumentsMax = 1;
	cCOLUMNS.prototype.arrayIndexes = {0: 1};
	cCOLUMNS.prototype.argumentsType = [argType.reference];
	cCOLUMNS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		var range;
		if (cElementType.array === arg0.type) {
			return new cNumber(arg0.getCountElementInRow());
		} else if (cElementType.cellsRange === arg0.type || cElementType.cell === arg0.type ||
			cElementType.cell3D === arg0.type || cElementType.cellsRange3D === arg0.type) {
			range = arg0.getRange();
		}
		return (range ? new cNumber(Math.abs(range.getBBox0().c1 - range.getBBox0().c2) + 1) :
			new cError(cErrorType.wrong_value_type));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cEXPAND() {
	}

	//***array-formula***
	cEXPAND.prototype = Object.create(cBaseFunction.prototype);
	cEXPAND.prototype.constructor = cEXPAND;
	cEXPAND.prototype.name = 'EXPAND';
	cEXPAND.prototype.argumentsMin = 2;
	cEXPAND.prototype.argumentsMax = 4;
	cEXPAND.prototype.arrayIndexes = {0: 1, 3: 1};
	cEXPAND.prototype.argumentsType = [argType.reference, argType.number, argType.number, argType.any];
	cEXPAND.prototype.Calculate = function (arg) {
		const MAX_ARRAY_SIZE = 1048576;
		let array;
		let arg0 = arg[0];
		let arg3 = arg[3] ? arg[3] : new cError(cErrorType.not_available);

		function expandedArray(arr) {
			let res = new cArray();

			for(let i = 0; i < rows; i++) {
				if(!arr[i]) {
					arr[i] = [];
				}
				for(let j = 0; j < columns; j++) {
					if(!arr[i][j]) {
						arr[i][j] = pad_with;
					}
				}
			}
			res.fillFromArray(arr);
			return res;
		}

		// --------------------- arg0(array) type check ----------------------//
		if (cElementType.cellsRange === arg0.type || cElementType.array === arg0.type) {
			array = arg0.getMatrix();
		} else if(cElementType.cellsRange3D === arg0.type) {
			array = arg0.getMatrix()[0];
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			return arg0.getValue();
		} else if (cElementType.number === arg0.type || cElementType.string === arg0.type ||
			cElementType.bool === arg0.type || cElementType.error === arg0.type) {
			return arg0;
		} else {
			return new cError(cErrorType.not_available);
		}
		if(arg0.length === 0){
			return new cError(cErrorType.wrong_value_type);
		}

		// --------------------- arg1(row) type check ----------------------//
		let rows = arg[1];
		let dimension = arg0.getDimensions();
		if(cElementType.empty === rows.type) {
			rows = new cNumber(dimension.row);
		} else if(cElementType.array === rows.type) {
			rows = rows.getElementRowCol(0, 0);
		} else if(cElementType.cellsRange === rows.type || cElementType.cellsRange3D === rows.type) {
			// TODO не получилось точно выяснить поведение функции при передаче в нее cellsRange вторым или третьим аргументом, поэтому пока возвращаем ошибку 
			rows = new cError(cErrorType.wrong_value_type);
		};
		rows = rows.tocNumber();
		
		if(cElementType.error === rows.type) {
			return rows;
		}
		rows = rows.toNumber();

		// --------------------- arg2(column) type check ----------------------//
		let columns = arg[2];
		if(cElementType.empty === columns.type) {
			columns = new cNumber(dimension.row);
		} else if(cElementType.array === columns.type) {
			columns = columns.getElementRowCol(0, 0);
		} else if(cElementType.cellsRange === columns.type || cElementType.cellsRange3D === columns.type) {
			// TODO не получилось точно выяснить поведение функции при передаче в нее cellsRange вторым или третьим аргументом, поэтому пока возвращаем ошибку
			columns = new cError(cErrorType.wrong_value_type);
		}
		columns = columns.tocNumber();

		if(cElementType.error === columns.type) {
			return columns;
		}

		columns = columns.toNumber();

		// --------------------- arg3(pad_with) type check ----------------------//
		let pad_with = arg3;
		if(cElementType.cellsRange === arg3.type || cElementType.cellsRange3D === arg3.type || cElementType.array === arg3.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		// check length and max array size
		if((rows * columns) > MAX_ARRAY_SIZE) {
			return new cError(cErrorType.not_numeric);
		} else if(rows >= array.length && columns >= array[0].length) {
			return expandedArray(array);
		}

		return new cError(cErrorType.wrong_value_type);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFILTER() {
	}

	//***array-formula***
	cFILTER.prototype = Object.create(cBaseFunction.prototype);
	cFILTER.prototype.constructor = cFILTER;
	cFILTER.prototype.name = 'FILTER';
	cFILTER.prototype.argumentsMin = 2;
	cFILTER.prototype.argumentsMax = 3;
	cFILTER.prototype.isXLFN = true;
	cFILTER.prototype.arrayIndexes = {0: 1, 1: 1};
	cFILTER.prototype.argumentsType = [argType.reference, argType.reference, argType.any];
	cFILTER.prototype.Calculate = function (arg) {
		function rangeModeLoop (rows, columns, isColumnMode) {
			let resArr = new cArray();

			for (let i = 0; i < rows; i++) {
				for (let j = 0; j < columns; j++) {
					let val = arg1.getValueByRowCol ? arg1.getValueByRowCol(i, j) : arg1.getElementRowCol(i, j);

					val = val.tocBool();
					val = val.toBool ? val.toBool() : new cError(cErrorType.wrong_value_type);
					if (cElementType.error === val.type) {
						return val;
					}

					if (val) {
						isColumnMode ? resArr.pushCol(arg0._getCol(j), 0) : resArr.pushRow(arg0._getRow(i), 0);
					}
				}
			}
			
			return resArr;
		}
 
		let resultArr = new cArray(),
			arg0 = arg[0],
			arg1 = arg[1],
			arg2 = arg[2] ? arg[2] : new cEmpty(),
			baseMode = false,		// val && range || val && val || range && val
			rangeMode = false;		// range && range

		if (cElementType.empty === arg0.type || cElementType.empty === arg1.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		// 4 options: 1) range && range; 2) range && value; 3) value && range; 4) value && value
		if ((cElementType.array === arg0.type || cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) && (cElementType.array === arg1.type || cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type)) {
			// 1) range && range
			rangeMode = true;
		} else if ((cElementType.array === arg0.type || cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) && (cElementType.array !== arg1.type && cElementType.cellsRange !== arg1.type && cElementType.cellsRange3D !== arg1.type)) {
			// 2) range && value
			// Return array arg0 if arg1 === true and if array arg0 is one-dimensional
			let arg0Dimensons = arg0.getDimensions();
			if ((arg0Dimensons.row > 1 && arg0Dimensons.col > 1)) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				baseMode = true;
			}
		} else if ((cElementType.array !== arg0.type && cElementType.cellsRange !== arg0.type && cElementType.cellsRange3D !== arg0.type) && (cElementType.array === arg1.type || cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type)) {
			// 3) value && range
			baseMode = true;
			arg1 = arg1.isOneElement() ? arg1.getFirstElement() : new cError(cErrorType.wrong_value_type);
		} else {
			// 4) value && value
			baseMode = true;
		}
		
		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}

		if (rangeMode) {
			const initialArrayDimensions = arg0.getDimensions(),
				initRows = initialArrayDimensions.row,
				initColumns = initialArrayDimensions.col,
				lookingArrayDimensions = arg1.getDimensions();

			// check for matching array sizes
			if (lookingArrayDimensions.row === 1 && lookingArrayDimensions.col === initColumns) {
				resultArr = rangeModeLoop(lookingArrayDimensions.row, lookingArrayDimensions.col, true);
			} else if (lookingArrayDimensions.row === initRows && lookingArrayDimensions.col === 1) {
				resultArr = rangeModeLoop(lookingArrayDimensions.row, lookingArrayDimensions.col, false);
			} else {
				// the size of the desired array does not match the initial
				return new cError(cErrorType.wrong_value_type);
			}

			if (resultArr.type === cElementType.error) {
				return resultArr;
			} else {
				resultArr = (resultArr.countElement > 0 || resultArr.rowCount > 0) ? resultArr : ((cElementType.empty !== arg2.type) ? arg2 : new cError(cErrorType.wrong_value_type));
			}
		} else if (baseMode) {
			arg1 = arg1.tocBool();
			arg1 = arg1.toBool ? arg1.toBool() : new cError(cErrorType.wrong_value_type);

			if (cElementType.error === arg1.type) {
				resultArr = arg1;
			} else if (arg1) {
				if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
					arg0 = arg0.getValue();
				}
				resultArr = arg0;
			} else {
				// should be #CALC!
				resultArr = (cElementType.empty !== arg2.type) ? arg2 : new cError(cErrorType.wrong_value_type);
			}
		}

		return resultArr;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cFORMULATEXT() {
	}

	//***array-formula***
	cFORMULATEXT.prototype = Object.create(cBaseFunction.prototype);
	cFORMULATEXT.prototype.constructor = cFORMULATEXT;
	cFORMULATEXT.prototype.name = 'FORMULATEXT';
	cFORMULATEXT.prototype.argumentsMin = 1;
	cFORMULATEXT.prototype.argumentsMax = 1;
	cFORMULATEXT.prototype.isXLFN = true;
	cFORMULATEXT.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.area_to_ref;
	cFORMULATEXT.prototype.argumentsType = [argType.reference];
	cFORMULATEXT.prototype.Calculate = function (arg) {

		var arg0 = arg[0];
		if (cElementType.error === arg0.type) {
			return arg0;
		}

		var res = null;
		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type ||
			cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			var bbox = arg0.getRange();
			var formula = bbox.isFormula();
			if (!formula) {
				return new cError(cErrorType.not_available);
			} else {
				res = new cString(bbox.getValueForEdit(true));
			}
		}

		return (null !== res ? res : new cError(cErrorType.wrong_value_type));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cGETPIVOTDATA() {
	}

	cGETPIVOTDATA.prototype = Object.create(cBaseFunction.prototype);
	cGETPIVOTDATA.prototype.constructor = cGETPIVOTDATA;
	cGETPIVOTDATA.prototype.name = 'GETPIVOTDATA';
	cGETPIVOTDATA.prototype.argumentsType = [argType.text, argType.text, [argType.text, argType.any]];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cHLOOKUP() {
	}

	//***array-formula***
	cHLOOKUP.prototype = Object.create(cBaseFunction.prototype);
	cHLOOKUP.prototype.constructor = cHLOOKUP;
	cHLOOKUP.prototype.name = 'HLOOKUP';
	cHLOOKUP.prototype.argumentsMin = 3;
	cHLOOKUP.prototype.argumentsMax = 4;
	cHLOOKUP.prototype.arrayIndexes = {1: 1, 2: 1};
	cHLOOKUP.prototype.argumentsType = [argType.any, argType.number, argType.number, argType.logical];
	cHLOOKUP.prototype.Calculate = function (arg) {
		//TODO  с excel есть несоостветствие - в тестовом файле - E11:H13
		if(this.bArrayFormula) {
			//исключение, когда в формуле массива берется из одного аргумента только 1 элемент
			if(cElementType.cellsRange3D === arg[2].type || cElementType.cellsRange === arg[2].type) {
				arg[2] = arg[2].getValue2(0,0);
			} else if(cElementType.array === arg[2].type) {
				arg[2] = arg[2].getValue2(0,0);
			}
		}
		return g_oHLOOKUPCache.calculate(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cHYPERLINK() {
	}

	cHYPERLINK.prototype = Object.create(cBaseFunction.prototype);
	cHYPERLINK.prototype.constructor = cHYPERLINK;
	cHYPERLINK.prototype.name = 'HYPERLINK';
	cHYPERLINK.prototype.argumentsMin = 1;
	cHYPERLINK.prototype.argumentsMax = 2;
	cHYPERLINK.prototype.argumentsType = [argType.text, argType.any];
	cHYPERLINK.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg.length === 1 ? null : arg[1];

		if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			arg0 = arg0.cross(arguments[1]);
		} else if (arg0 instanceof cArray) {
			arg0 = arg0.getElementRowCol(0, 0);
		}
		arg0 = arg0.tocString();


		if(arg1) {
			if (arg1 instanceof cArea || arg1 instanceof cArea3D) {
				arg1 = arg1.cross(arguments[1]);
			} else if (arg1 instanceof cArray) {
				arg1 = arg1.getElementRowCol(0, 0);
			}

			if(arg1 instanceof cRef || arg1 instanceof cRef3D) {
				arg1 = arg1.getValue();
			}
			if(arg1 instanceof cEmpty) {
				arg1 = new cNumber(0);
			}
		} else {
			arg1 = arg0.tocString();
		}

		if (arg0 instanceof cError) {
			arg0.hyperlink = "";
			return arg0;
		}
		if (arg1 instanceof cError) {
			arg1.hyperlink = "";
			return arg1;
		}

		var res = arg1;
		res.hyperlink = arg0.getValue();

		return res;
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cINDEX() {
	}

	//***array-formula***
	cINDEX.prototype = Object.create(cBaseFunction.prototype);
	cINDEX.prototype.constructor = cINDEX;
	cINDEX.prototype.name = 'INDEX';
	cINDEX.prototype.argumentsMin = 2;
	cINDEX.prototype.argumentsMax = 4;
	cINDEX.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cINDEX.prototype.arrayIndexes = {0: 1};
	cINDEX.prototype.argumentsType = [argType.reference, argType.number, argType.number];
	cINDEX.prototype.Calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1] && (cElementType.empty !== arg[1].type) ? arg[1] : new cNumber(1),
			arg2 = arg[2] && (cElementType.empty !== arg[2].type) ? arg[2] : new cNumber(1),
			arg3 = arg[3] && (cElementType.empty !== arg[3].type) ? arg[3] : new cNumber(1), res;

		if (cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.tocArea();
			if (!arg0) {
				return new cError(cErrorType.not_available);
			}
		} else if (cElementType.error === arg0.type) {
			return arg0;
		}

		arg1 = arg1.tocNumber();
		arg2 = arg2.tocNumber();
		arg3 = arg3.tocNumber();

		if (cElementType.error === arg1.type || cElementType.error === arg2.type || cElementType.error === arg3.type) {
			return new cError(cErrorType.wrong_value_type);
		}

		//TODO в дальнейшем необходимо продумать преобразования аргументов на основе argumentsType!!!
		if (cElementType.array === arg1.type) {
			arg1 = arg1.getElementRowCol(0,0);
			if (cElementType.error === arg1.type) {
				return new cError(cErrorType.wrong_value_type);
			}
		}
		if (cElementType.array === arg2.type) {
			arg2 = arg2.getElementRowCol(0,0);
			if (cElementType.error === arg2.type) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if(arg[3] && cElementType.empty !== arg[3].type && arg3 > 1) {
			return new cError(cErrorType.bad_reference);
		}

		arg1 = arg1.getValue();
		arg2 = arg2.getValue();

		if (arg1 < 0 || arg2 < 0) {
			return new cError(cErrorType.wrong_value_type);
		}

		var generateArray = function (_from, row, col) {
			var ret = null;
			var _colCount = _from.getCountElementInRow();
			var _rowCount = _from.rowCount;
			var i;
			row = row !== undefined ? Math.ceil(row) : row;
			col = col !== undefined ? Math.ceil(col) : col;
			if (undefined !== row) {
				if (_rowCount < row) {
					if (col === undefined && _rowCount === 1 && _from.array[0] && _from.array[0][row - 1]) {
						ret = new cArray();
						ret.addElement(_from.array[0][row - 1]);
						return ret;
					} else {
						return null;
					}
				}
				ret = new cArray();
				for (i = 0; i < _colCount; i++) {
					ret.addElement(_from.array[row - 1][i])
				}
			} else if (undefined !== col) {
				if (_colCount < col) {
					if (row === undefined && _colCount === 1 && _from.array[col - 1] && _from.array[col - 1][0]) {
						ret = new cArray();
						ret.addElement(_from.array[col - 1][0]);
						return ret;
					} else {
						return null;
					}
				}

				ret = new cArray();
				for (i = 0; i < _rowCount; i++) {
					ret.addRow();
					ret.addElement(_from.array[i][col - 1])
				}
			}
			return ret;
		};

		AscCommonExcel.executeInR1C1Mode(false, function () {
			if (cElementType.array === arg0.type) {
				if ((!arg[1] || arg1 === 0) && (!arg[2] || arg2 === 0)) {
					//возвращаем массив
					res = arg0;
				} else if (!arg[2] || arg2 === 0) {
					//возращаем массив из arg1 строки
					res = generateArray(arg0, arg1);
				} else if (!arg[1] || arg1 === 0) {
					//возращаем массив из arg2 столбца
					res = generateArray(arg0, undefined, arg2);
				} else if(undefined === arg[2] && 1 === arg0.rowCount) {//если последний аргумент опущен, и выделенa 1 строка
					res = arg0.getValue2(0, (0 === arg1) ? 0 : arg1 - 1);
				} else if(undefined === arg[2] && 1 === arg0.getCountElementInRow()) {//если последний аргумент опущен, и выделен 1 столбец
					res = arg0.getValue2((0 === arg1) ? 0 : arg1 - 1, 0);
				} else {
					res = arg0.getValue2((1 === arg0.rowCount || 0 === arg1) ? 0 : arg1 - 1, 0 === arg2 ? 0 : arg2 - 1);
				}
			} else if (cElementType.cellsRange === arg0.type) {
				var ws = arg0.getWS(), bbox = arg0.getBBox0();

				if(cElementType.empty === arg[1].type) {
					arg1 = 0;
				}

				var diffArg1 = arg1 === 0 ? 0 : 1;
				var diffArg2 = arg2 === 0 ? 0 : 1;
				if(undefined === arg[2] && bbox.r1 === bbox.r2) {//если последний аргумент опущен, и выделенa 1 строка
					if (arg1 > Math.abs(bbox.c1 - bbox.c2) + 1) {
						res = new cError(cErrorType.bad_reference);
					} else {
						res = new Asc.Range(bbox.c1 + arg1 - diffArg1, bbox.r1, bbox.c1 + arg1 - diffArg1, bbox.r1);
						res = new cRef(res.getName(), ws);
					}
				} else if(undefined === arg[2] && bbox.c1 === bbox.c2 && arg1 > 0) {//если последний аргумент опущен, и выделен 1 столбец
					if (arg1 > Math.abs(bbox.r1 - bbox.r2) + 1) {
						res = new cError(cErrorType.bad_reference);
					} else {
						res = new Asc.Range(bbox.c1, bbox.r1 + arg1 - diffArg1, bbox.c1, bbox.r1 + arg1 - diffArg1);
						res = new cRef(res.getName(), ws);
					}
				} else if(undefined === arg[2] && Math.abs(bbox.r1 - bbox.r2) + 1 > 1 && Math.abs(bbox.c1 - bbox.c2) + 1 > 1) {//если последний аргумент опущен, и выделен более 1 строки и более 1 столбца
					//так себя ведёт excel в случае с cellsArea
					res = new cError(cErrorType.bad_reference);
				} else if (bbox.r1 === bbox.r2) {/*одна строка*/
					res = new Asc.Range(bbox.c1 + arg2 - 1, bbox.r1, bbox.c1 + arg2 - 1, bbox.r1);
					res = new cRef(res.getName(), ws);
				} else {
					if (0 === arg1 && arg2 > 0) {
						if (arg2 > Math.abs(bbox.c1 - bbox.c2) + 1) {
							res = new cError(cErrorType.bad_reference);
						} else {
							if (0 === arg2 || undefined === arg[2] || cElementType.empty === arg[2].type) {
								res = new Asc.Range(bbox.c1 + arg2 - 1, bbox.r1, bbox.c2 + arg2 - 1, bbox.r2);
							} else {
								res = new Asc.Range(bbox.c1 + arg2 - 1, bbox.r1, bbox.c1 + arg2 - 1, bbox.r2);
							}
							res = res.isOneCell() ? new cRef(res.getName(),ws) : new cArea(res.getName(), ws);
						}
					} else if ((0 === arg2 || undefined === arg[2] || cElementType.empty === arg[2].type) && arg1 > 0) {
						if (arg1 > Math.abs(bbox.r1 - bbox.r2) + 1) {
							res = new cError(cErrorType.bad_reference);
						} else {
							res = new Asc.Range(bbox.c1 + arg2 - diffArg2, bbox.r1 + arg1 - diffArg1, bbox.c2 + arg2 - diffArg2, bbox.r1 + arg1 - diffArg1);
							res = res.isOneCell() ? new cRef(res.getName(),ws) : new cArea(res.getName(), ws);
						}
					} else if ((0 === arg1 || undefined === arg[1] || cElementType.empty === arg[1].type) && (0 === arg2 || undefined === arg[2] || cElementType.empty === arg[2].type)) {
						res = new Asc.Range(bbox.c1 + arg2 - diffArg2, bbox.r1 + arg1 - diffArg1, bbox.c2 + arg2 - diffArg2, bbox.r2 + arg1 - diffArg1);
						res = res.isOneCell() ? new cRef(res.getName(),ws) : new cArea(res.getName(), ws);
					} else {
						if (arg1 > Math.abs(bbox.r1 - bbox.r2) + 1 || arg2 > Math.abs(bbox.c1 - bbox.c2) + 1) {
							res = new cError(cErrorType.bad_reference);
						} else {
							res = new Asc.Range(bbox.c1 + arg2 - diffArg2, bbox.r1 + arg1 - diffArg1, bbox.c1 + arg2 - diffArg2, bbox.r1 + arg1 - diffArg1);
							res = new cRef(res.getName(), ws);
						}
					}
				}
			} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
				if ((0 === arg1 || 1 === arg1) && (0 === arg2 || 1 === arg2)) {
					res = arg0.getValue();
				}
			} else {
				res = new cError(cErrorType.wrong_value_type);
			}
		});

		return res ? res : new cError(cErrorType.bad_reference);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cINDIRECT() {
	}

	//TODO есть разница с MS - в тестовом файле E6
	//***array-formula***
	cINDIRECT.prototype = Object.create(cBaseFunction.prototype);
	cINDIRECT.prototype.constructor = cINDIRECT;
	cINDIRECT.prototype.name = 'INDIRECT';
	cINDIRECT.prototype.argumentsMin = 1;
	cINDIRECT.prototype.argumentsMax = 2;
	cINDIRECT.prototype.ca = true;
	cINDIRECT.prototype.argumentsType = [argType.text, argType.logical];
	cINDIRECT.prototype.Calculate = function (arg) {
		var t = this, arg0 = arg[0].tocString(), arg1 = arg[1] ? arg[1] : new cBool(true), ws = arguments[3],
			wb = ws.workbook, o = {
				Formula: "", pCurrPos: 0
			}, ref, found_operand, ret;

		var _getWorksheetByName = function(name){
			if(!name) {
				return null;
			}
			for(var i = 0; i < wb.aWorksheets.length; i++)
				if(wb.aWorksheets[i].getName().toLowerCase() == name.toLowerCase()){
					return wb.aWorksheets[i];
				}
			return null;
		};

		function parseReference() {
			if ((ref = parserHelp.is3DRef.call(o, o.Formula, o.pCurrPos, true))[0]) {
				var _tableTMP;
				var wsFrom = _getWorksheetByName(ref[1]);
				var wsTo = (null !== ref[2]) ? _getWorksheetByName(ref[2]) : wsFrom;
				if (!(wsFrom && wsTo)) {
					return new cError(cErrorType.bad_reference);
				}
				if (parserHelp.isArea.call(o, o.Formula, o.pCurrPos)) {
					found_operand = new cArea3D(o.real_str ? o.real_str.toUpperCase() : o.operand_str.toUpperCase(), wsFrom, wsTo);
				} else if (parserHelp.isRef.call(o, o.Formula, o.pCurrPos)) {
					if (wsTo !== wsFrom) {
						found_operand = new cArea3D(o.real_str ? o.real_str.toUpperCase() : o.operand_str.toUpperCase(), wsFrom, wsTo);
					} else {
						found_operand = new cRef3D(o.real_str ? o.real_str.toUpperCase() : o.operand_str.toUpperCase(), wsFrom);
					}
				} else if (parserHelp.isName.call(o, o.Formula, o.pCurrPos)) {
					found_operand = new cName3D(o.operand_str, wsFrom);
				}
			} else if (parserHelp.isArea.call(o, o.Formula, o.pCurrPos)) {
				found_operand = new cArea(o.real_str ? o.real_str.toUpperCase() : o.operand_str.toUpperCase(), ws);
			} else if (parserHelp.isRef.call(o, o.Formula, o.pCurrPos, true)) {
				found_operand = new cRef(o.real_str ? o.real_str.toUpperCase() : o.operand_str.toUpperCase(), ws);
			} else if (parserHelp.isName.call(o, o.Formula, o.pCurrPos)) {
				found_operand = new cName(o.operand_str, ws);
			} else if (_tableTMP = parserHelp.isTable.call(o, o.Formula, o.pCurrPos)) {
				found_operand = AscCommonExcel.cStrucTable.prototype.createFromVal(_tableTMP, wb, ws);

				if (found_operand.type === cElementType.error) {
					found_operand = null;
				} else {
					found_operand = found_operand.toRef ? found_operand.toRef() : null;
				}
			}
		}

		if (cElementType.array === arg0.type) {
			ret = new cArray();
			arg0.foreach(function (elem, r) {
				o = {Formula: elem.toString(), pCurrPos: 0};
				AscCommonExcel.executeInR1C1Mode(!!(arg1 && arg1.value === false), parseReference);
				if (!ret.array[r]) {
					ret.addRow();
				}
				ret.addElement(found_operand)
			});
			return ret;
		} else {
			o.Formula = arg0.toString();
			AscCommonExcel.executeInR1C1Mode(!!(arg1 && arg1.value == false), parseReference);
			if (found_operand) {
				if (cElementType.name === found_operand.type || cElementType.name3D === found_operand.type) {
					found_operand = found_operand.toRef(arguments[1]);
					if (found_operand && cElementType.error === found_operand.type) {
						ret = new cError(cErrorType.bad_reference);
					} else {
						ret = found_operand;
					}
				} else {
					ret = found_operand;
				}
			} else {
				ret = new cError(cErrorType.bad_reference);
			}
		}

		return ret;

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cLOOKUP() {
	}

	//***array-formula***
	cLOOKUP.prototype = Object.create(cBaseFunction.prototype);
	cLOOKUP.prototype.constructor = cLOOKUP;
	cLOOKUP.prototype.name = 'LOOKUP';
	cLOOKUP.prototype.argumentsMin = 2;
	cLOOKUP.prototype.argumentsMax = 3;
	cLOOKUP.prototype.arrayIndexes = {1: 1, 2: 1};
	cLOOKUP.prototype.argumentsType = [argType.any, argType.reference, argType.reference];
	cLOOKUP.prototype.Calculate = function (arg) {
		let arg0 = arg[0], arg1 = arg[1], arg2 = 2 === arg.length ? arg1 : arg[2], resC = -1, resR = -1,
			t = this, res;

		if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			if (arg0.isOneElement()) {
				arg0 = arg0.getFirstElement();
			} else {
				arg0 = new cError(cErrorType.wrong_value_type);
			}
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0, 0);
		}


		if (cElementType.cell === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		function arrFinder(arr) {
			if (arr.getRowCount() > arr.getCountElementInRow()) {
				//ищем в первом столбце
				resC = arr.getCountElementInRow() > 1 ? 1 : 0;
				let arrCol = arr.getCol(0);
				resR = _func.binarySearch(arg0, arrCol);
			} else {
				//ищем в первой строке
				resR = arr.getRowCount() > 1 ? 1 : 0;
				let arrRow = arr.getRow(0);
				resC = _func.binarySearch(arg0, arrRow);
			}
		}

		if (!( (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type ||
				cElementType.array === arg1.type) &&
				(cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type ||
					cElementType.array === arg2.type) )) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (cElementType.array === arg1.type && cElementType.array === arg2.type) {
			if (arg1.getRowCount() !== arg2.getRowCount() &&
				arg1.getCountElementInRow() !== arg2.getCountElementInRow()) {
				return new cError(cErrorType.not_available);
			}

			arrFinder(arg1);

			if (resR <= -1 && resC <= -1 || resR <= -2 || resC <= -2) {
				return new cError(cErrorType.not_available);
			}

			return arg2.getElementRowCol(resR, resC);

		} else if (cElementType.array === arg1.type || cElementType.array === arg2.type) {

			let _arg1, _arg2;

			_arg1 = cElementType.array === arg1.type ? arg1 : arg2;

			_arg2 = cElementType.array === arg2.type ? arg1 : arg2;

			let BBox = _arg2.getBBox0();

			if (_arg1.getRowCount() !== (BBox.r2 - BBox.r1) && _arg1.getCountElementInRow() !== (BBox.c2 - BBox.c1)) {
				return new cError(cErrorType.not_available);
			}

			arrFinder(_arg1);

			if (resR <= -1 && resC <= -1 || resR <= -2 || resC <= -2) {
				return new cError(cErrorType.not_available);
			}

			let c = new CellAddress(BBox.r1 + resR, BBox.c1 + resC, 0);
			_arg2.getWS()._getCellNoEmpty(c.getRow0(), c.getCol0(), function (cell) {
				res = checkTypeCell(cell);
			});
			return res;
		} else {
			if (cElementType.cellsRange3D === arg1.type && !arg1.isSingleSheet() ||
				cElementType.cellsRange3D === arg2.type && !arg2.isSingleSheet()) {
				return new cError(cErrorType.not_available);
			}

			//todo test and delete!
			if(false) {
				/*var arg1Range, arg2RowsLength;

				if (cElementType.cellsRange3D === arg1.type) {
					arg1Range = arg1.getMatrix()[0];
				} else if (cElementType.cellsRange === arg1.type) {
					arg1Range = arg1.getMatrix();
				}

				if (cElementType.cellsRange3D === arg2.type) {
					arg2RowsLength = arg2.bbox.r2 - arg2.bbox.r1 + 1;
					//arg2Range = arg2.getMatrix()[0];
				} else if (cElementType.cellsRange === arg2.type) {
					arg2RowsLength = arg2.range.bbox.r2 - arg2.range.bbox.r1 + 1;
					//arg2Range = arg2.getMatrix();
				}

				var bVertical = arg1Range[0].length >= arg1Range.length;//r>=c
				var index;
				var tempArr = [], i;
				if(bVertical) {
					for (i = 0; i < arg1Range[0].length; i++) {
						tempArr.push(arg1Range[0][i]);
					}
				} else {
					for (i = 0; i < arg1Range.length; i++) {
						tempArr.push(arg1Range[i][0]);
					}
				}

				if(tempArr[tempArr.length - 1] && tempArr[tempArr.length - 1].value < arg0.value) {
					//в этом случае фукнция бинарного поиска одаст последний элемент. для конкретного случая это неверно
					//Если функции не удается найти искомое_значение, то в просматриваемом_векторе выбирается наибольшее значение, которое меньше искомого_значения или равно ему.
					var diff = null;
					var endNumber;
					for(i = 0; i < tempArr.length; i++) {
						if(cElementType.number === tempArr[i].type) {
							if(tempArr[i].value <= arg0.value && (null === diff || diff > (arg0.value - tempArr[i].value))) {
								index = i;
								diff = arg0.value - tempArr[i].value;
							}
							endNumber = i;
						}
					}
					if(undefined === index) {
						if(undefined !== endNumber) {
							index = endNumber;
						}
					}
				}
				if(index === undefined) {
					index = _func.binarySearch(arg0, tempArr);

					if (index < 0) {
						return new cError(cErrorType.not_available);
					}
				}*/
			} else {
				var arg2RowsLength;
				var bbox;
				if (cElementType.cellsRange3D === arg1.type) {
					bbox = arg1.bbox;
				} else if (cElementType.cellsRange === arg1.type) {
					bbox = arg1.range.bbox;
				}

				if (cElementType.cellsRange3D === arg2.type) {
					arg2RowsLength = arg2.bbox.r2 - arg2.bbox.r1 + 1;
				} else if (cElementType.cellsRange === arg2.type) {
					arg2RowsLength = arg2.range.bbox.r2 - arg2.range.bbox.r1 + 1;
				}


				var bVertical = bbox.r2 - bbox.r1 >= bbox.c2 - bbox.c1;
				var index;

				const _getValue = function(n) {
					let r, c;
					if(bVertical) {
						r = n;
						c = 0;
					} else {
						r = 0;
						c = n;
					}
					let res = arg1.getValueByRowCol(r, c);
					return res ? res : new cEmpty();
				};

				let length = bVertical ? bbox.r2 - bbox.r1 : bbox.c2 - bbox.c1;
				let lastValue = _getValue(length);
				if(lastValue && lastValue.value < arg0.value) {
					//в этом случае фукнция бинарного поиска одаст последний элемент. для конкретного случая это неверно
					//Если функции не удается найти искомое_значение, то в просматриваемом_векторе выбирается наибольшее значение, которое меньше искомого_значения или равно ему.
					let diff = null;
					let endNumber;
					for(let i = 0; i <= length; i++) {
						let tempValue = _getValue(i);
						if(cElementType.number === tempValue.type) {
							if(tempValue.value <= arg0.value && (null === diff || diff > (arg0.value - tempValue.value))) {
								index = i;
								diff = arg0.value - tempValue.value;
							}
							endNumber = i;
						}
					}
					if(undefined === index) {
						if(undefined !== endNumber) {
							index = endNumber;
						}
					}
				}
				if(index === undefined) {
					index = _func.binarySearchByRange(arg0, arg1);

					if (index === undefined || index < 0) {
						return new cError(cErrorType.not_available);
					}
				}
			} /*else {
				var arg2RowsLength;

				if (cElementType.cellsRange3D === arg2.type) {
					arg2RowsLength = arg2.bbox.r2 - arg2.bbox.r1 + 1;
				} else if (cElementType.cellsRange === arg2.type) {
					arg2RowsLength = arg2.range.bbox.r2 - arg2.range.bbox.r1 + 1;
				}

				index = g_oLOOKUPCache.calculate(arg);

				if (index < 0) {
					return new cError(cErrorType.not_available);
				}
			}*/


			let ws = cElementType.cellsRange3D === arg1.type && arg1.isSingleSheet() ? arg1.getWS() : arg1.ws;

			if (cElementType.cellsRange3D === arg1.type) {
				if (arg1.isSingleSheet()) {
					ws = arg1.getWS();
				} else {
					return new cError(cErrorType.bad_reference);
				}
			} else if (cElementType.cellsRange === arg1.type) {
				ws = arg1.getWS();
			} else {
				return new cError(cErrorType.bad_reference);
			}

			AscCommonExcel.executeInR1C1Mode(false, function () {
				let b = arg2.getBBox0();
				if (2 === arg.length) {
					if (!bVertical) {
						// return the lookup value
						// res = new cRef(ws.getCell3(b.r1 + 0, b.c1 + index).getName(), ws);
						// return the last element in column (like in ms)
						res = new cRef(ws.getCell3(b.r2, b.c1 + index).getName(), ws);
					} else {
						// return the lookup value
						// res = new cRef(ws.getCell3(b.r1 + index, b.c1 + 0).getName(), ws);
						// return the last element in row (like in ms)
						res = new cRef(ws.getCell3(b.r1 + index, b.c2).getName(), ws);
					}
				} else {
					if (1 === arg2RowsLength) {
						// return the lookup value
						// res = new cRef(ws.getCell3(b.r1 + 0, b.c1 + index).getName(), ws);
						// return the last element in column (like in ms)
						res = new cRef(ws.getCell3(b.r2, b.c1 + index).getName(), ws);
					} else {
						// return the lookup value
						// res = new cRef(ws.getCell3(b.r1 + index, b.c1 + 0).getName(), ws);
						// return the last element in row (like in ms)
						res = new cRef(ws.getCell3(b.r1 + index, b.c2).getName(), ws);
					}
				}
			});
			return res;
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cMATCH() {
	}

	//***array-formula***
	cMATCH.prototype = Object.create(cBaseFunction.prototype);
	cMATCH.prototype.constructor = cMATCH;
	cMATCH.prototype.name = 'MATCH';
	cMATCH.prototype.argumentsMin = 2;
	cMATCH.prototype.argumentsMax = 3;
	cMATCH.prototype.arrayIndexes = {1: 1};
	cMATCH.prototype.argumentsType = [argType.any, argType.number, argType.number];
	cMATCH.prototype.Calculate = function (arg) {
		return g_oMatchCache.calculate(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cOFFSET() {
	}

	//***array-formula***
	cOFFSET.prototype = Object.create(cBaseFunction.prototype);
	cOFFSET.prototype.constructor = cOFFSET;
	cOFFSET.prototype.name = 'OFFSET';
	cOFFSET.prototype.argumentsMin = 3;
	cOFFSET.prototype.argumentsMax = 5;
	cOFFSET.prototype.ca = true;
	cOFFSET.prototype.arrayIndexes = {0: 1};
	cOFFSET.prototype.argumentsType = [argType.reference, argType.number, argType.number, argType.number, argType.number];
	cOFFSET.prototype.Calculate = function (arg) {

		function validBBOX(bbox) {
			return 0 <= bbox.r1 && bbox.r1 <= gc_nMaxRow0 && 0 <= bbox.c1 && bbox.c1 <= gc_nMaxCol0 && 0 <= bbox.r2 &&
				bbox.r2 <= gc_nMaxRow0 && 0 <= bbox.c2 && bbox.c2 <= gc_nMaxCol0;
		}

		var arg0 = arg[0], arg1 = arg[1].tocNumber(), arg2 = arg[2].tocNumber();
		var arg3 = 3 < arg.length ? (cElementType.empty === arg[3].type ? new cNumber(1) : arg[3].tocNumber()) : new cNumber(-1);
		var arg4 = 4 < arg.length ? (cElementType.empty === arg[4].type ? new cNumber(1) : arg[4].tocNumber()) : new cNumber(-1);

		var argError;
		if (argError = this._checkErrorArg([arg0, arg1, arg2, arg3, arg4])) {
			return argError;
		}

		arg1 = arg1.getValue();
		arg2 = arg2.getValue();
		arg3 = arg3.getValue();
		arg4 = arg4.getValue();

		if (arg3 == 0 || arg4 == 0) {
			return new cError(cErrorType.bad_reference);
		}

		var res;
		if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type ||
			cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
			var box = arg0.getBBox0();
			if (box) {
				box = box.clone(true);

				//в документации написано, что в отрицательных значений в 4 и 5 аргументах быть не может
				//но на деле ms рассчитывает такие формулы
				//сделал аналогично

				box.c1 = box.c1 + arg2;
				box.r1 = box.r1 + arg1;
				box.c2 = box.c2 + arg2;
				box.r2 = box.r2 + arg1;
				if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
					if (arg.length > 3) {
						if (arg4 < 0) {
							box.c1 = box.c1 + arg4 + 1;
						} else {
							box.c2 = box.c1 + arg4 - 1;
						}

						if (arg3 < 0) {
							box.r1 = box.r1 + arg3 + 1;
						} else {
							box.r2 = box.r1 + arg3 - 1;
						}
					}
				} else {
					if (arg.length > 3) {
						if (arg4 < 0) {
							box.c1 = box.c1 + arg4 + 1;
							box.c2 = box.c1 - arg4 - 1;
						} else {
							box.c2 = box.c1 + arg4 - 1;
						}

						if (arg3 < 0) {
							box.r1 = box.r1 + arg3 + 1;
							box.r2 = box.r1 - arg3 - 1;
						} else {
							box.r2 = box.r1 + arg3 - 1;
						}
					}
				}

				if (!validBBOX(box)) {
					return new cError(cErrorType.bad_reference);
				}

				var name;
				AscCommonExcel.executeInR1C1Mode(false, function () {
					name = box.getName();
				});
				var ws = arg0.getWS();
				var wsCell = arguments[3];
				if (box.isOneCell()) {
					res = wsCell === ws ? new cRef(name, ws) : new cRef3D(name, ws);
				} else {
					res = wsCell === ws ? new cArea(name, ws) : new cArea3D(name, ws, ws);
				}
			}
		}

		if (!res) {
			res = new cError(cErrorType.wrong_value_type);
		}
		return res;

	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROW() {
	}

	//***array-formula***
	cROW.prototype = Object.create(cBaseFunction.prototype);
	cROW.prototype.constructor = cROW;
	cROW.prototype.name = 'ROW';
	cROW.prototype.argumentsMax = 1;
	cROW.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.setArrayRefAsArg;
	cROW.prototype.argumentsType = [argType.reference];
	cROW.prototype.Calculate = function (arg) {
		var bbox;
		var res;
		var opt_row = arguments[5];
		if (opt_row !== undefined) {
			return new cNumber(opt_row + 1);
		} else if (0 === arg.length) {
			bbox = arguments[1];
			res = bbox ? new cNumber(bbox.r1 + 1) : null;
		} else {
			var arg0 = arg[0];
			if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
				bbox = arg0.getRange();
				bbox = bbox && bbox.bbox;
				res = bbox ? new cNumber(bbox.r1 + 1) : null;
			} else if (cElementType.cellsRange === arg0.type || cElementType.cellsRange3D === arg0.type) {
				bbox = arg0.getRange();
				bbox = bbox && bbox.bbox;

				if (bbox && bbox.r2 > bbox.r1) {
					res = new cArray();
					for (var i = bbox.r1; i <= bbox.r2; i++) {
						res.addRow();
						res.addElement(new cNumber(i + 1))
					}
				} else {
					res = bbox ? new cNumber(bbox.r1 + 1) : null;
				}
			}
		}

		return res ? res : new cError(cErrorType.bad_reference);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cROWS() {
	}

	//***array-formula***
	cROWS.prototype = Object.create(cBaseFunction.prototype);
	cROWS.prototype.constructor = cROWS;
	cROWS.prototype.name = 'ROWS';
	cROWS.prototype.argumentsMin = 1;
	cROWS.prototype.argumentsMax = 1;
	cROWS.prototype.arrayIndexes = {0: 1};
	cROWS.prototype.argumentsType = [argType.reference];
	cROWS.prototype.Calculate = function (arg) {
		var arg0 = arg[0];
		var range;
		if (cElementType.array === arg0.type) {
			return new cNumber(arg0.getRowCount());
		} else if (cElementType.cellsRange === arg0.type || cElementType.cell === arg0.type ||
			cElementType.cell3D === arg0.type || cElementType.cellsRange3D === arg0.type) {
			range = arg0.getRange();
		}
		return (range ? new cNumber(Math.abs(range.getBBox0().r1 - range.getBBox0().r2) + 1) :
			new cError(cErrorType.wrong_value_type));
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cRTD() {
	}

	cRTD.prototype = Object.create(cBaseFunction.prototype);
	cRTD.prototype.constructor = cRTD;
	cRTD.prototype.name = 'RTD';
	cRTD.prototype.argumentsType = [argType.text, argType.text, [argType.text]];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cSORT() {
	}

	cSORT.prototype = Object.create(cBaseFunction.prototype);
	cSORT.prototype.constructor = cSORT;
	cSORT.prototype.name = 'SORT';
	cSORT.prototype.argumentsMin = 1;
	cSORT.prototype.argumentsMax = 4;
	cSORT.prototype.isXLFN = true;
	cSORT.prototype.arrayIndexes = {0: 1, 1: 1, 2: 1, 3: 1};
	cSORT.prototype.argumentsType = [argType.reference, argType.number, argType.number, argType.bool];
	cSORT.prototype.Calculate = function (arg) {
		function sortWithIndices(arr) {
			const isRowMode = arr.length === 1 ? true : false;
			const indexedArray = isRowMode
				? arr[0].map(function (item, index) { return { item, index } })
				: arr.map(function (item, index) {
					item = item[0];
					return { item, index };
				});

			indexedArray.sort(function (a,b) {
				const valueA = a.item.value;
				const valueB = b.item.value; 

				if (valueA < valueB) {
					return sort_order === 1 ? -1 : 1;
				} else if (valueA > valueB) {
					return sort_order === 1 ? 1 : -1;
				} else {
					return 0;
				}
			});
			
			return indexedArray;
		}

		function sortArray (array, isByCol) {
			let resultArr = new cArray(),
				tempArrIndicies = [],
				targetElem = isByCol ? array._getRow(sort_index - 1) : array._getCol(sort_index - 1);

			// sorting an array with indices
			tempArrIndicies = sortWithIndices(targetElem);

			for (let i = 0; i < tempArrIndicies.length; i++) {
				let target = isByCol ? array._getCol(tempArrIndicies[i].index) : array._getRow(tempArrIndicies[i].index);
				isByCol ? resultArr.pushCol(target, 0) : resultArr.pushRow(target, 0);
			}

			return resultArr;
		}

		function arrayHelper (byColArray, by_col) {
			let dimensions = byColArray.getDimensions(),
				fElem = sortArray(array, by_col).getFirstElement(),
				resArr = new cArray();
			
			for (let i = 0; i < dimensions.row; i++) {
				resArr.addRow();
				for (let j = 0; j < dimensions.col; j++) {
					if (i === 0 && j === 0) {
						resArr.addElement(fElem);
						continue;
					}
					let el = new cError(cErrorType.wrong_value_type);
					resArr.addElement(el);
				}
			}
			return resArr;
		}

		function isValidArray (array, maxRowCol) {
			let dimensions = array.getDimensions();
			for (let i = 0; i < dimensions.row; i++) {
				for (let j = 0; j < dimensions.col; j++) {
					let elem = array.getValueByRowCol ? array.getValueByRowCol(i, j) : array.getElementRowCol(i, j);
					if (!elem) {
						return false;
					}
					elem = elem.tocNumber();
					if (elem.type === cElementType.error) {
						return false;
					} else if (Math.floor(elem.getValue()) > maxRowCol || Math.floor(elem.getValue()) <= 0) {
						return false;
					}
				}
			}
			return true;
		}

		let arg0 = arg[0],								// array
			arg1 = arg[1] ? arg[1] : new cNumber(1),	// sort_index
			arg2 = arg[2] ? arg[2] : new cNumber(1),	// sort_order
			arg3 = arg[3] ? arg[3] : new cBool(false);	// by_col ?

		// check args err
		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if (cElementType.error === arg1.type) {
			return arg1;
		}
		if (cElementType.error === arg2.type) {
			return arg2;
		}
		if (cElementType.error === arg3.type) {
			return arg3;
		}

		// check args empty
		if (cElementType.empty === arg1.type) {
			arg1 = new cNumber(1);
		}
		if (cElementType.empty === arg2.type) {
			arg2 = new cNumber(1);
		}
		if (cElementType.empty === arg3.type) {
			arg3 = new cBool(false);
		}

		let array, sort_index, sort_order, by_col, isArg1Array = false, isArg3Array = false, maxRows, maxCols;
			
		// check args type:
		// arg0(initial array) check
		if (cElementType.array !== arg0.type && cElementType.cellsRange !== arg0.type && cElementType.cellsRange3D !== arg0.type) {
			let elem;
			if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
				elem = arg0.getValue();
			} else {
				elem = arg0;
			}
			array = new cArray();
			array.addElement(elem);
		} else {
			array = arg0;
		}

		maxRows = array.getDimensions().row;
		maxCols = array.getDimensions().col;

		// arg1(sort_index) check
		if (cElementType.array !== arg1.type && cElementType.cellsRange !== arg1.type && cElementType.cellsRange3D !== arg1.type) {
			sort_index = arg1.tocNumber();
		} else {
			isArg1Array = true;
			let arg1Dimensions = arg1.getDimensions();
			if (arg1Dimensions.row > maxRows || arg1Dimensions.col > maxCols) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				sort_index = arg1.getFirstElement().tocNumber();
			}
		}

		// arg2(sort_order) check
		if (cElementType.array !== arg2.type && cElementType.cellsRange !== arg2.type && cElementType.cellsRange3D !== arg2.type) {
			sort_order = arg2.tocNumber();
		} else if (arg2.isOneElement()){
			sort_order = arg2.getFirstElement();
		} else {
			return new cError(cErrorType.wrong_value_type);
		}

		// arg3(by_col) check
		if (cElementType.array !== arg3.type && cElementType.cellsRange !== arg3.type && cElementType.cellsRange3D !== arg3.type) {
			by_col = arg3.tocBool();
		} else {
			by_col = arg3.getFirstElement();
			if (!by_col) {
				by_col = new cBool(false);
			}	
			isArg3Array = true;
		}

		if (cElementType.error === sort_index.type) {
			return sort_index;
		} else {
			sort_index = Math.floor(sort_index.getValue());
		}

		if (cElementType.error === sort_order.type) {
			return sort_order;
		} else {
			sort_order = Math.floor(sort_order.getValue());
		}

		if (cElementType.error === by_col.type) {
			return by_col;
		} else if (cElementType.bool !== by_col.type) {
			return new cError(cErrorType.wrong_value_type);
		} else {
			by_col = by_col.toBool();
		}

		if (sort_index <= 0 || (sort_order !== -1 && sort_order !== 1)) {
			return new cError(cErrorType.wrong_value_type);
		}

		if (!by_col) {
			if ((sort_index > maxCols) || (isArg1Array && !isValidArray(arg1, maxCols))) {
				return new cError(cErrorType.wrong_value_type);
			}
		} else {
			if ((sort_index > maxRows) || (isArg1Array && !isValidArray(arg1, maxRows))) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (isArg3Array) {
			// TODO it is not completely clear how the function works when receiving an array as the last argument
			return arrayHelper(arg3, by_col);
		}

		return sortArray(array, by_col);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTRANSPOSE() {
	}

	//***array-formula***
	cTRANSPOSE.prototype = Object.create(cBaseFunction.prototype);
	cTRANSPOSE.prototype.constructor = cTRANSPOSE;
	cTRANSPOSE.prototype.name = 'TRANSPOSE';
	cTRANSPOSE.prototype.argumentsMin = 1;
	cTRANSPOSE.prototype.argumentsMax = 1;
	cTRANSPOSE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTRANSPOSE.prototype.arrayIndexes = {0: 1};
	cTRANSPOSE.prototype.argumentsType = [argType.any];
	cTRANSPOSE.prototype.Calculate = function (arg) {

		function TransposeMatrix(A) {

			var tMatrix = [], res = new cArray();

			for (var i = 0; i < A.length; i++) {
				for (var j = 0; j < A[i].length; j++) {
					if (!tMatrix[j]) {
						tMatrix[j] = [];
					}
					tMatrix[j][i] = A[i][j];
				}
			}

			res.fillFromArray(tMatrix);

			return res;
		}

		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type) {
			//TODO возможно стоит на вход функции Calculate в случае применения как формулы массива сразу передавать преобразованный range в array
			if(!this.bArrayFormula) {
				arg0 = arg0.cross(arguments[1]);
				return arg0;
			} else {
				arg0 = arg0.getMatrix();
			}
		} else if(cElementType.cellsRange3D === arg0.type) {
			//TODO возможно стоит на вход функции в случае применения как формулы массива сразу передавать преобразованный range в array
			arg0 = arg0.getMatrix()[0];
		} else if(cElementType.array === arg0.type) {
			arg0 = arg0.getMatrix();
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			return arg0.getValue();
		} else if (cElementType.number === arg0.type || cElementType.string === arg0.type ||
			cElementType.bool === arg0.type || cElementType.error === arg0.type) {
			return arg0;
		} else {
			return new cError(cErrorType.not_available);
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if(0 === arg0.length){
			return new cError(cErrorType.wrong_value_type);
		}

		return TransposeMatrix(arg0);
	};

	function takeDrop(arg, argument1, isDrop) {
		var argError = cBaseFunction.prototype._checkErrorArg.call(this, arg);
		if (argError) {
			return argError;
		}

		let arg1 = arg[0];
		let matrix;
		if (arg1.type === cElementType.cellsRange || arg1.type === cElementType.array || arg1.type === cElementType.cell || arg1.type === cElementType.cell3D) {
			matrix = arg1.getMatrix();
		} else if (arg1.type === cElementType.cellsRange3D) {
			if (arg1.isSingleSheet()) {
				matrix = arg1.getMatrix()[0];
			} else {
				return new cError(cErrorType.bad_reference);
			}
		} else if (arg1.type === cElementType.error) {
			return arg1;
		} else if (arg1.type === cElementType.empty) {
			return new cError(cErrorType.wrong_value_type);
		} else {
			matrix = [[arg1]];
		}

		let array = new cArray();
		array.fillFromArray(matrix);

		let arg2 = arg[1];
		if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
			//_arg = _arg.getValue2(0,0);
			return new cError(cErrorType.wrong_value_type);
		} else if (cElementType.array === arg2.type) {
			//_arg = _arg.getElementRowCol(0, 0);
			return new cError(cErrorType.wrong_value_type);
		}


		if (cElementType.empty === arg2.type) {
			arg2 = null;
		} else {
			arg2 = arg2.tocNumber();
			if (arg2.type === cElementType.error) {
				return arg2;
			}
			arg2 = arg2.toNumber();
			arg2 = parseInt(arg2);
			if (Math.abs(arg2) < 1) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		let arg3 = arg[2] ? arg[2] : new cEmpty();
		if (cElementType.cellsRange === arg3.type || cElementType.cellsRange3D === arg3.type) {
			//_arg = _arg.getValue2(0,0);
			return new cError(cErrorType.wrong_value_type);
		} else if (cElementType.array === arg3.type) {
			//_arg = _arg.getElementRowCol(0, 0);
			return new cError(cErrorType.wrong_value_type);
		}


		if (cElementType.empty === arg3.type) {
			arg3 = null;
		} else {
			arg3 = arg3.tocNumber();
			if (arg3.type === cElementType.error) {
				return arg3;
			}
			arg3 = arg3.toNumber();
			arg3 = parseInt(arg3);
			if (Math.abs(arg3) < 1) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (isDrop) {
			let dimensions = array.getDimensions();

			if (arg2 && dimensions.row <= Math.abs(arg2)) {
				return new cError(cErrorType.wrong_value_type);
			}
			if (arg3 && dimensions.col <= Math.abs(arg3)) {
				return new cError(cErrorType.wrong_value_type);
			}

			if (arg2) {
				if (arg2 < 0) {
					arg2 = dimensions.row - Math.abs(arg2);
				} else {
					arg2 = -1 * (dimensions.row - arg2);
				}
			}
			if (arg3) {
				if (arg3 < 0) {
					arg3 = dimensions.col - Math.abs(arg3);
				} else {
					arg3 = -1 * (dimensions.col - arg3);
				}
			}
		}

		let res = array.crop(arg2, arg3);
		return res ? res : new cError(cErrorType.wrong_value_type);
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTAKE() {
	}

	//***array-formula***
	cTAKE.prototype = Object.create(cBaseFunction.prototype);
	cTAKE.prototype.constructor = cTAKE;
	cTAKE.prototype.name = 'TAKE';
	cTAKE.prototype.argumentsMin = 2;
	cTAKE.prototype.argumentsMax = 3;
	cTAKE.prototype.arrayIndexes = {0: 1};
	cTAKE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTAKE.prototype.isXLFN = true;
	cTAKE.prototype.argumentsType = [argType.reference, argType.number, argType.number];
	cTAKE.prototype.arrayIndexes = {0: 1};
	cTAKE.prototype.Calculate = function (arg) {
		return takeDrop(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cDROP() {
	}

	//***array-formula***
	cDROP.prototype = Object.create(cBaseFunction.prototype);
	cDROP.prototype.constructor = cDROP;
	cDROP.prototype.name = 'DROP';
	cDROP.prototype.argumentsMin = 2;
	cDROP.prototype.argumentsMax = 3;
	cDROP.prototype.arrayIndexes = {0: 1};
	cDROP.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cDROP.prototype.isXLFN = true;
	cDROP.prototype.argumentsType = [argType.reference, argType.number, argType.number];
	cDROP.prototype.arrayIndexes = {0: 1};
	cDROP.prototype.Calculate = function (arg) {
		return takeDrop(arg, arguments[1], true);
	};


	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cUNIQUE() {
	}

	//***array-formula***
	cUNIQUE.prototype = Object.create(cBaseFunction.prototype);
	cUNIQUE.prototype.constructor = cUNIQUE;
	cUNIQUE.prototype.name = 'UNIQUE';
	cUNIQUE.prototype.argumentsMin = 1;
	cUNIQUE.prototype.argumentsMax = 3;
	cUNIQUE.prototype.arrayIndexes = {0: 1};
	cUNIQUE.prototype.argumentsType = [argType.reference, argType.logical, argType.logical];
	cUNIQUE.prototype.isXLFN = true;
	cUNIQUE.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cUNIQUE.prototype.Calculate = function (arg) {

		var _getUniqueArr = function (_arr, _byCol, _exactlyOnce) {
			var rowCount = _arr && _arr.length;
			var colCount = _arr && _arr[0] && _arr[0].length;
			if (!rowCount || !colCount) {
				return cError(cErrorType.wrong_value_type);
			}

			var res = new cArray();
			var repeateArr = [];
			var i, j, n, _value;
			var resArr = [];

			var _key;
			if (!_byCol) {
				var _rowCount = 0;
				for (i = 0; i < rowCount; i++) {
					_key = "";
					for (j = 0; j < colCount; j++) {
						_value = _arr[i][j].getValue();
						_key += _value + ";";
						if (j === colCount - 1) {
							if (!repeateArr[_key]) {
								repeateArr[_key] = {index: _rowCount, count: 1};
								for (n = 0; n < colCount; n++) {
									if (!resArr[_rowCount]) {
										resArr[_rowCount] = [];
									}
									resArr[_rowCount].push(_arr[i][n]);
								}
								_rowCount++;
							}  else {
								repeateArr[_key].count++;
							}
						}
					}
				}
			} else {
				var _colCount = 0;
				for (i = 0; i < colCount; i++) {
					_key = "";
					for (j = 0; j < rowCount; j++) {
						_value = _arr[j][i].getValue();
						_key += _value + ";";
						if (j === rowCount - 1) {
							if (!repeateArr[_key]) {
								repeateArr[_key] = {index: _colCount, count: 1};
								for (n = 0; n < rowCount; n++) {
									if (!resArr[n]) {
										resArr[n] = [];
									}
									resArr[n][_colCount] = _arr[n][i];
								}
								_colCount++;
							} else {
								repeateArr[_key].count++;
							}
						}
					}
				}
			}

			if (_exactlyOnce) {
				var tempArr = [];
				var _counter = 0;
				for (i in repeateArr) {
					var _elem = repeateArr[i];
					if (_elem.count > 1) {
						continue;
					}
					if (!_byCol) {
						tempArr[_counter] = resArr[_elem.index];
					} else {
						for (j = 0; j < rowCount; j++) {
							if (!tempArr[j]) {
								tempArr[j] = [];
							}
							tempArr[j][_counter] = resArr[j][_elem.index];
						}
					}
					_counter++;
				}

				resArr = tempArr;
			}

			if (!resArr.length) {
				return new cError(cErrorType.wrong_value_type);
			}

			res.fillFromArray(resArr);

			return res;
		};

		var arg0 = arg[0];
		if (cElementType.cellsRange === arg0.type) {
			arg0 = arg0.getMatrix();
		} else if(cElementType.cellsRange3D === arg0.type) {
			arg0 = arg0.getMatrix()[0];
		} else if(cElementType.array === arg0.type) {
			arg0 = arg0.getMatrix();
		} else if (cElementType.cell === arg0.type || cElementType.cell3D === arg0.type) {
			return arg0.getValue();
		} else if (cElementType.number === arg0.type || cElementType.string === arg0.type ||
			cElementType.bool === arg0.type || cElementType.error === arg0.type) {
			return arg0;
		} else {
			return new cError(cErrorType.not_available);
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}
		if(0 === arg0.length){
			return new cError(cErrorType.wrong_value_type);
		}

		var arg1 = !arg[1] ? false : arg[1].tocBool();
		if (arg1 && cElementType.error === arg1.type) {
			return arg1;
		} else if (arg1) {
			arg1 = arg1.toBool();
		}

		var arg2 = !arg[2] ? false : arg[2].tocBool();
		if (arg2 && cElementType.error === arg2.type) {
			return arg2;
		} else if (arg2) {
			arg2 = arg2.toBool();
		}

		return _getUniqueArr(arg0, arg1, arg2);
	};

	/**
	 * @constructor
	 */
	function VHLOOKUPCache(bHor) {
		this.cacheId = {};
		this.cacheRanges = {};
		this.bHor = bHor;
	}

	VHLOOKUPCache.prototype.calculate = function (arg, argument1) {
		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2];
		var arg3 = arg[3] ? arg[3].tocBool().value : true;
		var opt_xlookup = arg[4] !== undefined;
		var opt_arg4, opt_arg5;
		if (opt_xlookup) {
			opt_arg4 = arg[4];
			opt_arg5 = arg[5];
		}
		var t = this, number, valueForSearching, r, c, res = -1, min, regexp, count;

		if (!opt_xlookup) {
			number = arg2.getValue() - 1;
			if (cElementType.array === arg2.type) {
				var arg2Val = arg2.getElementRowCol(0, 0);
				number = arg2Val ? arg2Val.getValue() - 1 : number;
			}

			if (isNaN(number)) {
				return new cError(cErrorType.bad_reference);
			}
			if (number < 0) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		if (cElementType.cell3D === arg0.type || cElementType.cell === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		var arg0Val;
		if (cElementType.array === arg0.type) {
			arg0Val = arg0.getElementRowCol(0, 0);
			valueForSearching = ('' + arg0Val.getValue()).toLowerCase();
		} else {
			arg0Val = arg0;
			valueForSearching = ('' + arg0.getValue()).toLowerCase();
		}

		if (cElementType.cellsRange === arg0Val.type) {
			arg0Val = arg0Val.cross(argument1);
		} else if (cElementType.cellsRange3D === arg0Val.type) {
			arg0Val = arg0Val.cross(argument1);
		}

		if (cElementType.error === arg0Val.type) {
			return arg0;
		}
		//TODO не тестировал на hlookup/x - поэтому поставил условия
		if (!opt_xlookup && false === this.bHor && cElementType.empty === arg0Val.type) {
			return new cError(cErrorType.not_available);
		}

		//TODO hlookup не правильно работает если первый агумент массив - раскомментировать тесты для hlookup
		var found = false;
		if (cElementType.array === arg1.type && !opt_xlookup) {
			// ToDo
			if (cElementType.string === arg0.type) {
				regexp = searchRegExp(valueForSearching);
			}
			arg1.foreach(function (elem, r, c) {
				var v = ('' + elem.getValue()).toLowerCase();
				var i = t.bHor ? c : r;
				if (0 === i) {
					min = v;
				}

				if (arg3) {
					if (valueForSearching === v) {
						res = i;
						found = true;
					} else if (valueForSearching > v && !found) {
						res = i;
					}
				} else {
					if (cElementType.string === arg0.type) {
						if (regexp.test(v)) {
							res = i;
						}
					} else if (valueForSearching === v) {
						res = i;
					}
				}

				min = Math.min(min, v);
			});

			if (/*min > valueForSearching ||*/ -1 === res) {
				return new cError(cErrorType.not_available);
			}

			count = this.bHor ? arg1.getRowCount() : arg1.getCountElementInRow();
			if (number > count - 1) {
				return new cError(cErrorType.bad_reference);
			}

			r = this.bHor ? number : res;
			c = this.bHor ? res : number;
			return arg1.getElementRowCol(r, c);
		}

		var range;
		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type ||
			cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			range = arg1.getRange();
		} else if (cElementType.array === arg1.type && opt_xlookup) {
			var _cacheElem = {elements: []};
			arg1.foreach(function (elem, r, c) {
				_cacheElem.elements.push({v: elem, i: (t.bHor ? c : r)});
			});
			return this._calculate(_cacheElem.elements, arg0Val, null, opt_arg4, opt_arg5);
		}

		if (!range) {
			return new cError(cErrorType.bad_reference);
		}

		var bb = range.getBBox0();
		count = this.bHor ? (bb.r2 - bb.r1) : (bb.c2 - bb.c1);
		if (number > count) {
			return new cError(cErrorType.bad_reference);
		}
		var ws = arg1.getWS();
		r = this.bHor ? bb.r1 : bb.r2;
		c = this.bHor ? bb.c2 : bb.c1;
		var oSearchRange = ws.getRange3(bb.r1, bb.c1, r, c);


		res = this._get(oSearchRange, arg0Val, arg3, opt_arg4, opt_arg5);
		if (opt_xlookup) {
			return res;
		}

		if (-1 === res) {
			return new cError(cErrorType.not_available);
		}

		r = this.bHor ? bb.r1 + number : res;
		c = this.bHor ? res : bb.c1 + number;
		var resVal;
		arg1.getWS()._getCellNoEmpty(r, c, function (cell) {
			resVal = checkTypeCell(cell);
		});
		if (cElementType.empty === resVal.type) {
			resVal = new cNumber(0);
		}

		return resVal;
	};
	VHLOOKUPCache.prototype._get = function (range, valueForSearching, arg3Value, opt_arg4, opt_arg5) {
		var res, _this = this, wsId = range.getWorksheet().getId();
		var opt_xlookup = opt_arg4 !== undefined;

		var sRangeName;
		AscCommonExcel.executeInR1C1Mode(false, function () {
			sRangeName = wsId + g_cCharDelimiter + range.getName();
		});

		var cacheElem = this.cacheId[sRangeName];
		if (!cacheElem) {
			cacheElem = {elements: [], results: {}};
			this.generateElements(range, cacheElem);
			this.cacheId[sRangeName] = cacheElem;
			var cacheRange = this.cacheRanges[wsId];
			if (!cacheRange) {
				cacheRange = new AscCommonExcel.RangeDataManager(null);
				this.cacheRanges[wsId] = cacheRange;
			}
			cacheRange.add(range.getBBox0(), cacheElem);
		}

		var sInputKey;
		if (!opt_xlookup) {
			sInputKey =
				valueForSearching.getValue() + g_cCharDelimiter + arg3Value + g_cCharDelimiter + valueForSearching.type;
		} else {
			sInputKey = valueForSearching.getValue() + g_cCharDelimiter + opt_arg4 + g_cCharDelimiter + opt_arg5 +
				g_cCharDelimiter + valueForSearching.type;
		}
		res = cacheElem.results[sInputKey];
		if (!res) {
			cacheElem.results[sInputKey] =
				res = this._calculate(cacheElem.elements, valueForSearching, arg3Value, opt_arg4, opt_arg5);
		}

		return res;
	};
	VHLOOKUPCache.prototype._calculate = function (cacheArray, valueForSearching, lookup, opt_arg4, opt_arg5) {
		var res = -1, i = 0, j, length = cacheArray.length, k, elem, val, nextVal;
		var xlookup = opt_arg4 !== undefined && opt_arg5 !== undefined;

		//TODO неверно работает функция, допустим для случая: VLOOKUP("12",A1:A5,1) 12.00 ; "qwe" ; "3" ; 3.00 ; 4.00

		//ascending order: ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE
		var _compareValues = function (val1, val2, op) {
			if (opt_arg4 === 2 && val2.type === cElementType.string) {
				var matchingInfo = AscCommonExcel.matchingValue(val1);
				return AscCommonExcel.matching(val2, matchingInfo)
			} else {
				var res = _func[val1.type][val2.type](val1, val2, op);
				return res ? res.value : false;
			}
		};

		var addNextOptVal = function (arrayVal, searchVal, isGreater) {
			var _needPush;
			if (opt_arg4 === -1 && (isGreater === false || (isGreater === undefined && _compareValues(arrayVal.v, searchVal, "<")))) {
				_needPush = true;
			} else if (opt_arg4 === 1 && (isGreater || (isGreater === undefined && _compareValues(arrayVal.v, searchVal, ">")))) {
				_needPush = true;
			}
			if (_needPush) {
				if (nextVal === undefined || _compareValues(arrayVal.v, nextVal.v, opt_arg4 === 1 ? "<" : ">")) {
					nextVal = arrayVal;
				}
			}
		};

		var simpleSearch = function (revert) {
			if (revert) {
				for (i = length - 1; i >= 0; i--) {
					elem = cacheArray[i];
					val = elem.v;
					if (_compareValues(valueForSearching, val, "=")) {
						return elem.i;
					}
					opt_arg4 !== undefined && addNextOptVal(elem, valueForSearching);
				}
			} else {
				for (i = 0; i < length; i++) {
					elem = cacheArray[i];
					val = elem.v;
					if (_compareValues(valueForSearching, val, "=")) {
						return elem.i;
					}
					(opt_arg4 === 1 || opt_arg4 === -1) && addNextOptVal(elem, valueForSearching);
				}
			}
			return -1;
		};

		//бинарный поиск для xlookup(так работает ms) бинарный поиск происходит до определенной длины массива
		//как только длина становится меньше n(около 10), начинается линейный поиск
		//так же в случае бинарного поиска когда требуется возвратить меньший или больший элемент(opt_arg4)
		//- возвращается последний обработанный элемент меньший(больший) искомого, между собой элементы не сравниваются

		//мы делаем иначе: бинарный поиск происходит всегда и не зависит от длины массива, при поиске наибольшего(наименьшего)
		//из обработанных элементов выбираем те, которые больше(меньше) -> из них уже ищем наименьший(наибольший)
		//т.е. в итоге получаем следующий наименьший/наибольший элемент
		var _binarySearch = function (revert) {
			i = 0;

			//TODO проверить обратный поиск
			if (revert) {
				j = length - 1;
				while (i <= j) {
					k = Math.ceil((i + j) / 2);
					elem = cacheArray[k];
					val = elem.v;
					if (_compareValues(valueForSearching, val, "=")) {
						return elem.i;
					} else if (_compareValues(valueForSearching, val, "<")) {
						i = k + 1;
						opt_arg4 !== undefined && addNextOptVal(elem, valueForSearching, true);
					} else {
						j = k - 1;
						opt_arg4 !== undefined && addNextOptVal(elem, valueForSearching, false);
					}
				}
			} else {
				j = length - 1;
				while (i <= j) {
					k = Math.floor((i + j) / 2);
					elem = cacheArray[k];
					val = elem.v;
					if (_compareValues(valueForSearching, val, "=")) {
						return elem.i;
					} else if (_compareValues(valueForSearching, val, "<")) {
						j = k - 1;
						opt_arg4 !== undefined && addNextOptVal(elem, valueForSearching, true);
					} else {
						i = k + 1;
						opt_arg4 !== undefined && addNextOptVal(elem, valueForSearching, false);
					}
				}
			}

			if (xlookup) {
				return -1;
			}

			var _res = Math.min(i, j);
			_res = -1 === _res ? _res : cacheArray[_res].i;
			return _res;
		};

		//TODO opt_arg5 - пока не обрабатываю результат == 2( A wildcard match where *, ?, and ~ have)

		if (xlookup) {
			if (Math.abs(opt_arg5) === 1) {
				res = simpleSearch(opt_arg5 < 0);
			} else if (Math.abs(opt_arg5) === 2) {
				res = _binarySearch(opt_arg5 < 0);
			}

			if (res === -1) {
				if ((opt_arg4 === -1 || opt_arg4 === 1) && nextVal) {
					res = nextVal.i;
				}
			}
		} else if (lookup) {
			res = _binarySearch();
			if (res === -1 && cElementType.string === valueForSearching.type) {
				res = simpleSearch();
			}
		} else {
			// Exact value
			res = simpleSearch();
		}

		return res;
	};
	VHLOOKUPCache.prototype.remove = function (cell) {
		var wsId = cell.ws.getId();
		var cacheRange = this.cacheRanges[wsId];
		if (cacheRange) {
			var oGetRes = cacheRange.get(new Asc.Range(cell.nCol, cell.nRow, cell.nCol, cell.nRow));
			for (var i = 0, length = oGetRes.all.length; i < length; ++i) {
				var elem = oGetRes.all[i];
				elem.data.results = {};
			}
		}
	};
	VHLOOKUPCache.prototype.clean = function () {
		this.cacheId = {};
		this.cacheRanges = {};
	};
	VHLOOKUPCache.prototype.generateElements = function (range, cacheElem) {
		var _this = this;

		//сильного прироста не получил, пока оставляю прежнюю обработку, подумать на счёт разбития диапазонов
		range._foreachNoEmpty(function (cell, r, c) {
			cacheElem.elements.push({v: checkTypeCell(cell), i: (_this.bHor ? c : r)});
		});
		return;

		//попытка оптимизации фукнции. если находим диапазон, который полностью перекрывает текущий или пересекаемся с текущим - тогда данные из кэша берём и не обращаемся к модели
		//флаг - получаем из кэша только первый элемент
		var bFast = true;
		if (range && cacheElem) {
			//ищем пересечения с уже имеющимися диапазонами
			var elementsIntervals = [];
			var addByIntervals = function (_elem, isIntersection) {
				if (elementsIntervals && elementsIntervals.length) {
					for (var k = 0; k < elementsIntervals.length; k++) {
						if (elementsIntervals[k].bbox.containsRange(_elem)) {
							return;
						} else if (_elem.bbox.containsRange(elementsIntervals[k].bbox)) {
							elementsIntervals.splice(k, 1);
							elementsIntervals.push(_elem);
							return;
						}
					}
				}

				elementsIntervals.push(_elem);
			};

			var ws = range.getWorksheet();
			if (ws) {
				var rangeData = this.getRangeDataBySheetId(ws.Id);
				if (rangeData) {
					var interval, elem, intersection;
					if (bFast) {
						interval = rangeData.getFirst(range.bbox);
						if (interval) {
							elem = interval;
							if (elem.bbox.containsRange(range.bbox)) {
								//диапазон в кэшэ полностью перекрывает новый
								//формируем новый массив из того, которые в кэше
								if (elem.data && elem.data.elements) {
									cacheElem.elements = elem.data.slice(_this.bHor ? range.bbox.c1 - elem.bbox.c1 : range.bbox.r1 - elem.bbox.r1, _this.bHor ? elem.bbox.c2 - range.bbox.c2 : elem.bbox.r2 - range.bbox.r2);
									return;
								}
							} else /*if (range.bbox.containsRange(elem.bbox))*/{
								//ищем пересечение
								intersection = elem.bbox.intersection(range.bbox);
								addByIntervals({bbox: intersection, elements: elem.data.elements.slice(_this.bHor ? intersection.c1 - elem.bbox.c1 : intersection.r1 - elem.bbox.r1, _this.bHor ? elem.bbox.c2 - intersection.c1 : elem.bbox.r2 - intersection.r1)});
							}
						}
					} else {
						var intervals = rangeData.tree && rangeData.tree.searchNodes(range.bbox);
						if (intervals && intervals.length) {
							for (var i = 0; i < intervals.length; i++) {
								interval = intervals[i];
								elem = interval.data;
								if (elem.bbox.isIntersect(range.bbox)) {
									if (elem.bbox.containsRange(range.bbox)) {
										//диапазон в кэшэ полностью перекрывает новый
										//формируем новый массив из того, которые в кэше
										if (elem.data && elem.data.elements) {
											cacheElem.elements = elem.data.slice(_this.bHor ? range.bbox.c1 - elem.bbox.c1 : range.bbox.r1 - elem.bbox.r1,
												_this.bHor ? elem.bbox.c2 - range.bbox.c2 : elem.bbox.r2 - range.bbox.r2);
											return;
										}
									} else /*if (range.bbox.containsRange(elem.bbox))*/{
										//ищем пересечение
										intersection = elem.bbox.intersection(range.bbox);
										addByIntervals({
											bbox: intersection,
											elements: elem.data.elements.slice(_this.bHor ? intersection.c1 - elem.bbox.c1 : intersection.r1 - elem.bbox.r1,
												_this.bHor ? elem.bbox.c2 - intersection.c1 : elem.bbox.r2 - intersection.r1)
										});
									}
								}
							}
						}

					}
				}
			}

			var addElemsFromWs = function (_range) {
				_range._foreachNoEmpty(function (cell, r, c) {
					cacheElem.elements.push({v: checkTypeCell(cell), i: (_this.bHor ? c : r)});
				});
			};

			if (elementsIntervals && elementsIntervals.length) {
				//сортируем по порядку
				elementsIntervals.sort(function (a, b) {
					return _this.bHor ? b.c1 - a.c1 : b.r1 - a.r1;
				});

				//проходимся по всем диапазонам, заполняем "окна"
				for (var j = 0; j < elementsIntervals.length; j++) {
					var lastCacheElem = cacheElem.elements[cacheElem.elements.length - 1];
					var elemIndex = elementsIntervals[j].elements[0].i;

					//если в cacheElem ещё ничего не добавлено или индекс следующего элемента из интервала не соответвует индексу + 1 элемента из кэша
					var cacheIndex = lastCacheElem && lastCacheElem.i;
					if (!lastCacheElem || elemIndex !== lastCacheElem.i + 1) {
						var r1, c1, r2, c2;
						if (!lastCacheElem) {
							if (!_this.bHor && range.bbox.r1 === elemIndex || _this.bHor && range.bbox.c1 === elemIndex) {
								//начало диапазонов совпадает
								cacheElem.elements = cacheElem.elements.concat(elementsIntervals[j].elements);
								if (j === elementsIntervals.length - 1) {
									lastCacheElem = cacheElem.elements[cacheElem.elements.length - 1];
									cacheIndex = lastCacheElem && lastCacheElem.i;

									if (!(!_this.bHor && range.bbox.r1 === cacheIndex || _this.bHor && range.bbox.c1 === cacheIndex)) {
										//берём последние элементы
										r1 = !_this.bHor ? cacheIndex + 1 : range.bbox.r1;
										c1 = _this.bHor ? cacheIndex + 1 : range.bbox.c1;

										r2 = range.bbox.r2;
										c2 =  range.bbox.c2;

										addElemsFromWs(ws.getRange3(r1, c1, r2, c2));
									}
								}
							} else {
								//берём от начала range до начала первого интервала
								r1 = range.bbox.r1;
								c1 = range.bbox.c1;

								r2 = !_this.bHor ? elemIndex - 1: range.bbox.r2;
								c2 = _this.bHor ? elemIndex - 1 : range.bbox.c2;

								addElemsFromWs(ws.getRange3(r1, c1, r2, c2));
								j--;
							}

						} else {
							//берём от конца последнего элемента из кэша до начала интервала
							r1 = !_this.bHor ? cacheIndex + 1 : range.bbox.r1;
							c1 = _this.bHor ? cacheIndex + 1 : range.bbox.c1;

							r2 = !_this.bHor ? elemIndex - 1: range.bbox.r2;
							c2 = _this.bHor ? elemIndex - 1 : range.bbox.c2;

							addElemsFromWs(ws.getRange3(r1, c1, r2, c2));
							j--;
						}
					} else {
						cacheElem.elements = cacheElem.elements.concat(elementsIntervals[j].elements);
						if (j === elementsIntervals.length - 1) {
							//берём последние элементы
							r1 = !_this.bHor ? cacheIndex + 1 : range.bbox.r1;
							c1 = _this.bHor ? cacheIndex + 1 : range.bbox.c1;

							r2 = range.bbox.r2;
							c2 =  range.bbox.c2;

							addElemsFromWs(ws.getRange3(r1, c1, r2, c2));
						}
					}
				}
			} else {
				addElemsFromWs(range);
			}
		}
	};
	VHLOOKUPCache.prototype.getRangeDataBySheetId = function (id) {
		return this.cacheRanges[id];
	};




	function MatchCache() {
		this.cacheId = {};
		this.cacheRanges = {};
	}

	MatchCache.prototype = Object.create(VHLOOKUPCache.prototype);
	MatchCache.prototype.constructor = MatchCache;
	MatchCache.prototype.calculate = function (arg, _arg1) {
		var arg0 = arg[0], arg1 = arg[1], arg2, arg3;
		var isXMatch = arg[4];

		if(isXMatch) {
			// default values for XMatch
			arg2 = arg[2] ? arg[2] : new cNumber(0);
			arg3 = arg[3] ? arg[3] : new cNumber(1);
		} else {
			// default values for Match
			arg2 = arg[2] ? arg[2] : new cNumber(1);
			arg3 = new cNumber(1);
		}

		if (cElementType.cellsRange3D === arg0.type || cElementType.cellsRange === arg0.type) {
			// arg0 = arg0.cross(_arg1);
			arg0 = arg0.getFullArray().getElementRowCol(0,0);

			if (cElementType.empty === arg0.type) {
				return new cError(cErrorType.not_available);
			}
		} else if (cElementType.array === arg0.type) {
			arg0 = arg0.getElementRowCol(0,0);
		} else if (cElementType.error === arg0.type) {
			return arg0;
		}

		if (cElementType.number === arg2.type || cElementType.bool === arg2.type) {
		} else if (cElementType.error === arg2.type) {
			return arg2;
		} else {
			return new cError(cErrorType.not_available);
		}

		var a2Value = arg2.getValue();
		if (!(-1 === a2Value || 0 === a2Value || 1 === a2Value || 2 === a2Value)) {
			return new cError(cErrorType.not_numeric);
		}

		var a3value = arg3.getValue();
		if(!(-2 === a3value || -1 === a3value || 1 === a3value || 2 === a3value)) {
			return new cError(cErrorType.wrong_value_type);
		}

		if(cElementType.error === arg1.type) {
			return new cError(cErrorType.not_available);
		} else if (cElementType.array === arg1.type) {
			arg1 = arg1.getMatrix();

			var i, a1RowCount = arg1.length, a1ColumnCount = arg1[0].length, arr;

			if (a1RowCount > 1 && a1ColumnCount > 1) {
				return new cError(cErrorType.not_available);
			} else if (a1RowCount === 1 && a1ColumnCount >= 1) {
				arr = arg1[0];
			} else {
				arr = [];
				for (i = 0; i < a1RowCount; i++) {
					arr[i] = arg1[i][0];
				}
			}
			return this._calculate(arr, arg0, arg2);
		}

		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type ||
			cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			// add range.isonecell
			var oSearchRange = arg1.getRange();
			if (!oSearchRange) {
				return new cError(cErrorType.bad_reference);
			}

			var a1RowCount = oSearchRange.bbox.r2 - oSearchRange.bbox.r1 + 1, a1ColumnCount = oSearchRange.bbox.c2 - oSearchRange.bbox.c1 + 1;
			var bHor = false;
			if (a1RowCount > 1 && a1ColumnCount > 1) {
				return new cError(cErrorType.not_available);
			} else if (a1RowCount === 1 && a1ColumnCount >= 1) {
				bHor = true;
			}

			return this._get(oSearchRange, arg0, arg2, arg3, bHor, isXMatch);
		} else {
			return new cError(cErrorType.not_available);
		}
	};
	MatchCache.prototype._get = function (range, arg0, arg2, arg3, bHor, isXMatch) {
		var res, _this = this, wsId = range.getWorksheet().getId(),
			sRangeName = wsId + g_cCharDelimiter + range.getName(), cacheElem = this.cacheId[sRangeName];
		var arg2Value = arg2.getValue();
		var arg3Value = arg3 ? arg3.getValue() : '';
		var valueForSearching = arg0.getValue();
		if (!cacheElem) {
			cacheElem = {elements: [], results: {}};

			range._foreachNoEmpty(function (cell, r, c) {
				cacheElem.elements.push({v: checkTypeCell(cell), i: (bHor ? c - range.bbox.c1 : r -  range.bbox.r1)});
			});
			this.cacheId[sRangeName] = cacheElem;
			var cacheRange = this.cacheRanges[wsId];
			if (!cacheRange) {
				cacheRange = new AscCommonExcel.RangeDataManager(null);
				this.cacheRanges[wsId] = cacheRange;
			}
			cacheRange.add(range.getBBox0(), cacheElem);
		}
		var sInputKey = arg3Value ? valueForSearching + g_cCharDelimiter + arg2Value + g_cCharDelimiter + arg3Value : valueForSearching + g_cCharDelimiter + arg2Value;
		res = cacheElem.results[sInputKey];
		if(!res && isXMatch) {
			cacheElem.results[sInputKey] = res = this._xMatchCalculate(cacheElem.elements, arg0, arg2, arg3);
		} else if (!res) {
			cacheElem.results[sInputKey] = res = this._calculate(cacheElem.elements, arg0, arg2);
		}
		return res;
	};
	MatchCache.prototype._calculate = function (arr, a0, a2) {

		var a2Value = a2.getValue();
		var a0Type = a0.type;
		var a0Value = a0.getValue();
		if (!(cElementType.number === a0Type || cElementType.string === a0Type || cElementType.bool === a0Type ||
			cElementType.error === a0Type || cElementType.empty === a0Type)) {
			if(cElementType.empty === a0Value.type) {
				a0Value = a0Value.tocNumber();
			}
			a0Type = a0Value.type;
			a0Value = a0Value.getValue();
		}

		var item, index = -1, curIndex;
		for (var i = 0; i < arr.length; ++i) {
			item = undefined !== arr[i].v ? arr[i].v : arr[i];
			curIndex = undefined !== arr[i].i ? arr[i].i : i;
			if (item.type === a0Type) {
				if (0 === a2Value) {
					if (cElementType.string === a0Type) {
						if (AscCommonExcel.searchRegExp2(item.toString(), a0Value)) {
							index = curIndex;
							break;
						}
					} else {
						if (item == a0Value) {
							index = curIndex;
							break;
						}
					}
				} else if (1 === a2Value) {
					if (item <= a0Value) {
						index = curIndex;
					} else {
						break;
					}
				} else if (-1 === a2Value) {
					if (item >= a0Value) {
						index = curIndex;
					} else {
						break;
					}
				}
			}
		}

		return (-1 < index) ? new cNumber(index + 1) : new cError(cErrorType.not_available);
	};
	MatchCache.prototype._xMatchCalculate = function (arr, a0, a2, a3) {
		let a0Type = a0.type;
		let	a0Value = a0.getValue(),
			a2Value = a2.getValue(),
			a3value = a3.getValue();

		if (!(cElementType.number === a0Type || cElementType.string === a0Type || cElementType.bool === a0Type ||
			cElementType.error === a0Type || cElementType.empty === a0Type)) {
			if(cElementType.empty === a0Value.type) {
				a0Value = a0Value.tocNumber();
			}
			a0Type = a0Value.type;
			a0Value = a0Value.getValue();
		}

		let item, index = -1, curIndex;

		if(1 === a3value) {
			// first iteration for precise search
			for (let i = 0; i < arr.length; ++i) {
				item = undefined !== arr[i].v ? arr[i].v : arr[i];
				curIndex = undefined !== arr[i].i ? arr[i].i : i;
				if (item.type === a0Type) {
					if (0 === a2Value || 2 === a2Value) {
						if (cElementType.string === a0Type) {
							if (AscCommonExcel.searchRegExp2(item.toString(), a0Value)) {
								index = curIndex;
								break;
							}
						} else {
							if (item == a0Value) {
								index = curIndex;
								break;
							}
						}
					} else if (-1 === a2Value || 1 === a2Value) {
						if (item == a0Value) {
							index = curIndex;
							break;
						} 
					}
				}
			}

			// second iteration for approximate search
			if(index === -1 && (1 === a2Value || -1 === a2Value)) {
				if(-1 === a2Value) {
					const lessEqualArr = arr.filter(function (item) { 
						return item.v <= a0Value;
					}).sort(function (a, b) { return a.v.getValue() < b.v.getValue() ? 1 : -1 });

					if(lessEqualArr.length > 0) {
						let closestVal = lessEqualArr[0];
						for(let i = 0; i < lessEqualArr.length; ++i) {
							if(closestVal.v.getValue() < lessEqualArr[i].v.getValue()) {
								closestVal = lessEqualArr[i];
							} else if(closestVal.v.getValue() == lessEqualArr[i].v.getValue()) {
								if(closestVal.i > lessEqualArr[i].i) {
									closestVal = lessEqualArr[i];
								}
							}
						}
						index = closestVal.i;
					}
				} else if(1 === a2Value) {
					const moreEqualArr = arr.filter(function (item) { 
						return item.v >= a0Value;
					}).sort(function (a, b) { return a.v.getValue() > b.v.getValue() ? 1 : -1 });

					if(moreEqualArr.length > 0) {
						let closestVal = moreEqualArr[0];
						for(let i = 0; i < moreEqualArr.length; ++i) {
							if(closestVal.v.getValue() > moreEqualArr[i].v.getValue()) {
								closestVal = moreEqualArr[i];
							} else if(closestVal.v.getValue() == moreEqualArr[i].v.getValue()) {
								if(closestVal.i > moreEqualArr[i].i) {
									closestVal = moreEqualArr[i];
								}
							}
						}
						index = closestVal.i;
					}
				}
			}
		} else if(-1 === a3value) {
			// reverse search
			// first iteration for precise search
			for (let i = arr.length - 1; i > 0; --i) {
				item = undefined !== arr[i].v ? arr[i].v : arr[i];
				curIndex = undefined !== arr[i].i ? arr[i].i : i;
				if (item.type === a0Type) {
					if (0 === a2Value || 2 === a2Value) {
						if (cElementType.string === a0Type) {
							if (AscCommonExcel.searchRegExp2(item.toString(), a0Value)) {
								index = curIndex;
								break;
							}
						} else {
							if (item == a0Value) {
								index = curIndex;
								break;
							}
						}
					} else if (-1 === a2Value || 1 === a2Value) {
						if (item == a0Value) {
							index = curIndex;
							break;
						} 
					}
				}
			}
			// second iteration for approximate search
			if(index === -1 && (1 === a2Value || -1 === a2Value)) {
				if(-1 === a2Value) {
					const lessEqualArr = arr.filter(function (item) {
						return item.v.getValue() <= a0Value;
					}).sort(function (a, b) { return a.v.getValue() < b.v.getValue() ? 1 : -1 });
					
					if(lessEqualArr.length > 0) {
						let closestVal = lessEqualArr[0];
						for(let i = 0; i < lessEqualArr.length; ++i) {
							if(closestVal.v.getValue() < lessEqualArr[i].v.getValue()) {
								closestVal = lessEqualArr[i];
							} else if(closestVal.v.getValue() == lessEqualArr[i].v.getValue()) {
								if(closestVal.i < lessEqualArr[i].i) {
									closestVal = lessEqualArr[i];
								}
							}
						}
						index = closestVal.i;
					}
				} else if(1 === a2Value) {
					const moreEqualArr = arr.filter(function (item) { 
						return item.v >= a0Value;
					}).sort(function (a, b) { return a.v.getValue() > b.v.getValue() ? 1 : -1 });
					if(moreEqualArr.length > 0) {
						let closestVal = moreEqualArr[0];
						for(let i = 0; i < moreEqualArr.length; ++i) {
							if(closestVal.v.getValue() > moreEqualArr[i].v.getValue()) {
								closestVal = moreEqualArr[i];
							} else if(closestVal.v.getValue() == moreEqualArr[i].v.getValue()) {
								if(closestVal.i < moreEqualArr[i].i) {
									closestVal = moreEqualArr[i];
								}
							}
						}
						index = closestVal.i;
					}
				}
			}

		} else if(2 === a3value) {
			if (2 === a2Value) {
				// wildcard match(err)
				return new cError(cErrorType.wrong_name);
			}
			index = XBinarySearch(a0Value, arr, a2Value, false);
		} else if(-2 === a3value) {
			if (2 === a2Value) {
				// wildcard match(err)
				return new cError(cErrorType.wrong_name);
			}
			index = XBinarySearch(a0Value, arr, a2Value, true);
		}

		return (-1 < index) ? new cNumber(index + 1) : new cError(cErrorType.not_available);
	}

	function LOOKUPCache() {
		this.cacheId = {};
		this.cacheRanges = {};
	}

	LOOKUPCache.prototype = Object.create(VHLOOKUPCache.prototype);
	LOOKUPCache.prototype.constructor = LOOKUPCache;

	LOOKUPCache.prototype.calculate = function (arg) {
		var arg0 = arg[0], arg1 = arg[1];
		var t = this, r, c, count;

		if (cElementType.cell3D === arg0.type || cElementType.cell === arg0.type) {
			arg0 = arg0.getValue();
		}

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		var arg0Val;
		if (cElementType.array === arg0.type) {
			arg0Val = arg0.getElementRowCol(0, 0);
		} else {
			arg0Val = arg0;
		}

		var range;
		if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type ||
			cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type) {
			range = arg1.getRange();
		}
		if (!range) {
			return new cError(cErrorType.bad_reference);
		}

		var bb = range.getBBox0();
		var bHor = bb.r2 - bb.r1 < bb.c2 - bb.c1;
		//count = bHor ? (bb.r2 - bb.r1) : (bb.c2 - bb.c1);

		var ws = arg1.getWS();
		r = bHor ? bb.r1 : bb.r2;
		c = bHor ? bb.c2 : bb.c1;
		var oSearchRange = ws.getRange3(bb.r1, bb.c1, r, c);

		if (cElementType.cellsRange === arg0Val.type) {
			arg0Val = arg0Val.cross(arguments[1]);
		} else if (cElementType.cellsRange3D === arg0Val.type) {
			arg0Val = arg0Val.cross(arguments[1]);
		}

		if (cElementType.error === arg0Val.type) {
			return arg0;
		}

		var res = this._get(oSearchRange, arg0Val, true);
		if (-1 === res) {
			return new cError(cErrorType.not_available);
		}

		return res;
	};
	LOOKUPCache.prototype._calculate = function (cacheArray, valueForSearching, lookup) {
		var res = -1, i = 0, j, length = cacheArray.length, k, elem, val;

		//TODO неверно работает функция, допустим для случая: VLOOKUP("12",A1:A5,1) 12.00 ; "qwe" ; "3" ; 3.00 ; 4.00

		//ascending order: ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE
		var _compareValues = function (val1, val2, op) {
			var res = _func[val1.type][val2.type](val1, val2, op);
			return res ? res.value : false;
		};

		if (lookup) {
			j = length - 1;
			while (i <= j) {
				k = Math.floor((i + j) / 2);
				elem = cacheArray[k];
				val = elem.v;
				if (_compareValues(valueForSearching, val, "=")) {
					return k;
				} else if (_compareValues(valueForSearching, val, "<")) {
					j = k - 1;
				} else {
					i = k + 1;
				}
			}
			res = Math.min(i, j);
		} else {
			// Exact value
			for (; i < length; i++) {
				elem = cacheArray[i];
				val = elem.v;
				if (_compareValues(valueForSearching, val, "=")) {
					return i;
				}
			}
		}
		return res;
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cVLOOKUP() {
	}

	cVLOOKUP.prototype = Object.create(cBaseFunction.prototype);
	cVLOOKUP.prototype.constructor = cVLOOKUP;
	cVLOOKUP.prototype.name = 'VLOOKUP';
	cVLOOKUP.prototype.argumentsMin = 3;
	cVLOOKUP.prototype.argumentsMax = 4;
	cVLOOKUP.prototype.arrayIndexes = {1: 1, 2: {0: 0}};
	cVLOOKUP.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cVLOOKUP.prototype.argumentsType = [argType.any, argType.number, argType.number, argType.logical];
	cVLOOKUP.prototype.Calculate = function (arg) {

		if(this.bArrayFormula) {
		 	//в случае когда первый аргумент - массив
			//исключение, когда в формуле массива берется из одного аргумента только 1 элемент
			if(cElementType.cellsRange3D === arg[2].type || cElementType.cellsRange === arg[2].type) {
				arg[2] = arg[2].getValue2(0,0);
			} else if(cElementType.array === arg[2].type) {
				arg[2] = arg[2].getValue2(0,0);
			}
		}

		return g_oVLOOKUPCache.calculate(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cXLOOKUP() {
	}

	//***array-formula***
	cXLOOKUP.prototype = Object.create(cBaseFunction.prototype);
	cXLOOKUP.prototype.constructor = cXLOOKUP;
	cXLOOKUP.prototype.name = 'XLOOKUP';
	cXLOOKUP.prototype.argumentsMin = 3;
	cXLOOKUP.prototype.argumentsMax = 6;
	cXLOOKUP.prototype.arrayIndexes = {1: 1, 2: 1};
	cXLOOKUP.prototype.argumentsType = [argType.any, argType.reference, argType.reference, argType.any, argType.number, argType.number];
	cXLOOKUP.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.value_replace_area;
	cXLOOKUP.prototype.isXLFN = true;
	cXLOOKUP.prototype.Calculate = function (arg) {

		var arg0 = arg[0], arg1 = arg[1], arg2 = arg[2];
		var arg3 = arg[3], arg4 = arg[4], arg5 = arg[5];

		if (cElementType.error === arg0.type) {
			return arg0;
		}

		if (!((cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type ||
			cElementType.array === arg1.type) &&
			(cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type ||
			cElementType.array === arg2.type))) {
			return new cError(cErrorType.not_available);
		}

		//[if_not_found]
		if (!arg3 || arg3.type === cElementType.empty) {
			arg3 = new cError(cErrorType.not_available);
		}

		//arg4/arg5 - только число
		//[match_mode]
		//0 - If none found, return #N/A. This is the default.
		//-1 - If none found, return the next smaller item.
		//1 - If none found, return the next larger item.
		//2 - A wildcard match where *, ?, and ~
		//TODO если аргумент массив/area - результат становится размером с этот массив
		//TODO либо обрабатывать выше и вызывать эту функцию для каждого элемента массива, либо здесь вычислять этот массив
		if (!arg4) {
			arg4 = new cNumber(0);
		}
		arg4 = arg4.tocNumber();
		if (cElementType.error === arg4.type) {
			return arg4;
		} else {
			arg4 = parseInt(arg4.toNumber());
			if (!(arg4 >= -1 && arg4 <= 2)) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		//TODO если аргумент массив/area - результат становится размером с этот массив
		//TODO либо обрабатывать выше и вызывать эту функцию для каждого элемента массива, либо здесь вычислять этот массив
		//[search_mode]
		//1 - Perform a search starting at the first item. This is the default.
		//-1 - Perform a reverse search starting at the last item.
		//2 - Perform a binary search that relies on lookup_array being sorted in ascending order. If not sorted, invalid results will be returned.
		//-2 - Perform a binary search that relies on lookup_array being sorted in descending order. If not sorted, invalid results will be returned.
		if (!arg5) {
			arg5 = new cNumber(1);
		}
		arg5 = arg5.tocNumber();
		if (cElementType.error === arg5.type) {
			return arg5;
		} else {
			arg5 = parseInt(arg5.toNumber());
			if (!(arg5 >= -2 && arg5 <= 2)) {
				return new cError(cErrorType.wrong_value_type);
			}
		}

		//массив arg1 должен содержать 1 строку или 1 столбец
		var dimensions1 = arg1.getDimensions();
		var dimensions2 = arg2.getDimensions();
		var bVertical = null;
		if (dimensions1 && dimensions2) {
			if (dimensions1.col === 1 && dimensions2.row === dimensions1.row) {
				bVertical = true;
			} else if (dimensions1.row === 1 && dimensions2.col === dimensions1.col) {
				bVertical = false;
			}
		}

		if (bVertical === null) {
			return new cError(cErrorType.wrong_value_type);
		} else {
			var res;
			if (bVertical) {
				res = g_oVLOOKUPCache.calculate([arg0, arg1, null, null, arg4, arg5], arguments[1]);
			} else {
				res = g_oHLOOKUPCache.calculate([arg0, arg1, null, null, arg4, arg5], arguments[1]);
			}

			if (res === -1) {
				return arg3;
			} else {
				//возвращаем из arg2 строку или столбец
				var _startRange = 0;
				if (dimensions2.bbox) {
					_startRange = bVertical ? dimensions2.bbox.r1 : dimensions2.bbox.c1;
				} else if (dimensions1.bbox) {
					_startRange = bVertical ? dimensions1.bbox.r1 : dimensions1.bbox.c1;
				}

				if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
					var _r1 = !bVertical ? dimensions2.bbox.r1 : res - _startRange + dimensions2.bbox.r1;
					var _cl = bVertical ? dimensions2.bbox.c1 : res - _startRange + dimensions2.bbox.c1;
					var _r2 = !bVertical ? dimensions2.bbox.r2 : res - _startRange + dimensions2.bbox.r1;
					var _c2 = bVertical ? dimensions2.bbox.c2 : res - _startRange + dimensions2.bbox.c1;
					var _range = new Asc.Range(_cl, _r1, _c2, _r2);

					var _res;
					var rangeName;
					AscCommonExcel.executeInR1C1Mode(false, function () {
						rangeName = _range.getName();
					});
					if (cElementType.cellsRange === arg2.type) {
						_res = _range.isOneCell() ? new cRef(rangeName, arg2.getWS()) : new cArea(rangeName, arg2.getWS());
					} else {
						_res = _range.isOneCell() ?  new cRef3D(rangeName, arg2.getWS()) : new cArea3D(rangeName, arg2.getWS());
					}
					return _res;
				} else {
					var _length = !bVertical ? dimensions2.row : dimensions2.col;
					var _array = new cArray();
					for (var i = 0; i < _length; i++) {
						var _row = !bVertical ? i : res - _startRange;
						var _col = bVertical ? i : res - _startRange;
						var _elem = arg2.getElementRowCol ? arg2.getElementRowCol(_row, _col) : arg2.getValueByRowCol(_row, _col);
						if (!bVertical) {
							_array.addRow();
							_array.addElement(_elem);
						} else {
							_array.addElement(_elem);
						}
					}
				}

				return _array;
			}
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cVSTACK() {
	}

	//***array-formula***
	cVSTACK.prototype = Object.create(cBaseFunction.prototype);
	cVSTACK.prototype.constructor = cVSTACK;
	cVSTACK.prototype.name = 'VSTACK';
	cVSTACK.prototype.argumentsMin = 1;
	cVSTACK.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cVSTACK.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cVSTACK.prototype.argumentsType = [[argType.reference]];
	cVSTACK.prototype.isXLFN = true;
	cVSTACK.prototype.Calculate = function (arg) {
		let unionArray;
		for (let i = 0; i < arg.length; i++) {
			let matrix;
			if (arg[i].type === cElementType.cellsRange || arg[i].type === cElementType.array || arg[i].type === cElementType.cell || arg[i].type === cElementType.cell3D) {
				matrix = arg[i].getMatrix();
			} else if (arg[i].type === cElementType.cellsRange3D) {
				if (arg[i].isSingleSheet()) {
					matrix = arg[i].getMatrix()[0];
				} else {
					return new cError(cErrorType.bad_reference);
				}
			} else if (arg[i].type === cElementType.error) {
				return arg[i];
			} else if (arg[i].type === cElementType.empty) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				matrix = [[arg[i]]];
			}

			//добавляем по строкам
			for (let j = 0; j < matrix.length; j++) {
				if (matrix[j]) {
					if (!unionArray) {
						unionArray = [];
					}
					unionArray.push(matrix[j]);
				}
			}
		}

		if (unionArray) {
			let res = new cArray();
			res.fillFromArray(unionArray);
			res.fillMatrix(new cError(cErrorType.not_available));
			return res;
		} else {
			return new cError(cErrorType.wrong_value_type);
		}
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cHSTACK() {
	}

	//***array-formula***
	cHSTACK.prototype = Object.create(cBaseFunction.prototype);
	cHSTACK.prototype.constructor = cHSTACK;
	cHSTACK.prototype.name = 'HSTACK';
	cHSTACK.prototype.argumentsMin = 1;
	cHSTACK.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cHSTACK.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cHSTACK.prototype.argumentsType = [[argType.reference]];
	cHSTACK.prototype.isXLFN = true;
	cHSTACK.prototype.Calculate = function (arg) {
		let unionArray;
		let startCol = 0;
		for (let i = 0; i < arg.length; i++) {
			let matrix;
			if (arg[i].type === cElementType.cellsRange || arg[i].type === cElementType.array || arg[i].type === cElementType.cell || arg[i].type === cElementType.cell3D) {
				matrix = arg[i].getMatrix();
			} else if (arg[i].type === cElementType.cellsRange3D) {
				if (arg[i].isSingleSheet()) {
					matrix = arg[i].getMatrix()[0];
				} else {
					return new cError(cErrorType.bad_reference);
				}
			} else if (arg[i].type === cElementType.error) {
				return arg[i];
			} else if (arg[i].type === cElementType.empty) {
				return new cError(cErrorType.wrong_value_type);
			} else {
				matrix = [[arg[i]]];
			}

			let maxColCount = 0;
			for (let j = 0; j < matrix.length; j++) {
				if (matrix[j]) {
					maxColCount = Math.max(maxColCount, matrix[j].length);
					for (let k = 0; k < matrix[j].length; k++) {
						if (matrix[j][k]) {
							if (!unionArray) {
								unionArray = [];
							}
							if (!unionArray[j]) {
								unionArray[j] = [];
							}
							unionArray[j][k + startCol] = matrix[j][k];
						}
					}
				}
			}
			startCol += maxColCount;
		}

		if (unionArray) {
			let res = new cArray();
			res.fillFromArray(unionArray);
			res.fillMatrix(new cError(cErrorType.not_available));
			return res;
		} else {
			return new cError(cErrorType.wrong_value_type);
		}
	};

	function toRowCol(arg, argument1, toCol) {
		var argError = cBaseFunction.prototype._checkErrorArg.call(this, arg);
		if (argError) {
			return argError;
		}

		//из документации:
		//Excel returns a #VALUE! when an array constant contains one or more numbers that are not a whole number.
		//не повторил в мс

		let arg1 = arg[0];
		if (arg1.type === arg1.empty) {
			return new cError(cErrorType.wrong_value_type);
		}
		arg1 = arg1.toArray();

		//Excel returns a #NUM when array is too large.
		let elemCount = arg1.length * arg1[0].length;
		if (elemCount > 1048578) {
			return new cError(cErrorType.not_available);
		}

		//0    Keep all values (default)
		//1    Ignore blanks
		//2    Ignore errors
		//3    Ignore blanks and errors
		let arg2 = arg[1] ? arg[1] : new cNumber(0);
		if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
			arg2 = arg2.cross(argument1);
		} else if (cElementType.array === arg2.type) {
			arg2 = arg2.getElementRowCol(0, 0);
		}
		arg2 = arg2.tocNumber();
		if (arg2.type === cElementType.error) {
			return arg2;
		}
		arg2 = arg2.toNumber();

		//scan_by_column
		let arg3 = arg[2] ? arg[2] : new cBool(false);
		if (cElementType.cellsRange === arg3.type || cElementType.cellsRange3D === arg3.type) {
			arg3 = arg3.cross(argument1);
		} else if (cElementType.array === arg3.type) {
			arg3 = arg3.getElementRowCol(0, 0);
		}
		arg3 = arg3.tocBool();
		if (arg3.type === cElementType.error) {
			return arg3;
		}
		arg3 = arg3.toBool();

		let arg1_array = new cArray();
		arg1_array.fillFromArray(arg1);


		var res = new cArray();
		arg1_array.foreach2(function (elem, r, c) {
			if (elem) {
				let needAdd = true;
				switch (arg2) {
					case 0:
						break;
					case 1:
						if (elem.type === cElementType.empty) {
							needAdd = false;
						}
						break;
					case 2:
						if (elem.type === cElementType.error) {
							needAdd = false;
						}
						break;
					case 3:
						if (elem.type === cElementType.error || elem.type === cElementType.empty) {
							needAdd = false;
						}
						break;
				}
				if (needAdd) {
					if (toCol) {
						res.addRow();
					}
					res.addElement(elem);
				}
			}
		}, arg3);

		return res;
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTOROW() {
	}

	//***array-formula***
	cTOROW.prototype = Object.create(cBaseFunction.prototype);
	cTOROW.prototype.constructor = cTOROW;
	cTOROW.prototype.name = 'TOROW';
	cTOROW.prototype.argumentsMin = 1;
	cTOROW.prototype.argumentsMax = 3;
	cTOROW.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTOROW.prototype.arrayIndexes = {0: 1};
	cTOROW.prototype.argumentsType = [argType.reference, argType.number, argType.bool];
	cTOROW.prototype.isXLFN = true;
	cTOROW.prototype.Calculate = function (arg) {
		return toRowCol(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cTOCOL() {
	}

	//***array-formula***
	cTOCOL.prototype = Object.create(cBaseFunction.prototype);
	cTOCOL.prototype.constructor = cTOCOL;
	cTOCOL.prototype.name = 'TOCOL';
	cTOCOL.prototype.argumentsMin = 1;
	cTOCOL.prototype.argumentsMax = 3;
	cTOCOL.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cTOCOL.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cTOCOL.prototype.argumentsType = [argType.reference, argType.number, argType.bool];
	cTOCOL.prototype.isXLFN = true;
	cTOCOL.prototype.Calculate = function (arg) {
		return toRowCol(arg, arguments[1], true);
	};

	function wrapRowsCols(arg, argument1, toCol) {
		var argError = cBaseFunction.prototype._checkErrorArg.call(this, arg);
		if (argError) {
			return argError;
		}

		let arg1 = arg[0];
		if (arg1.type === cElementType.empty) {
			return new cError(cErrorType.wrong_value_type);
		}
		var dimension = arg1.getDimensions();
		if (dimension.col > 1 && dimension.row > 1) {
			return new cError(cErrorType.wrong_value_type);
		}

		let arg2 = arg[1];
		if (cElementType.cellsRange === arg2.type || cElementType.cellsRange3D === arg2.type) {
			arg2 = arg2.getValueByRowCol(0,0);
		} else if (cElementType.array === arg2.type) {
			arg2 = arg2.getElementRowCol(0, 0);
		} else if (arg2.type === cElementType.empty) {
			return new cError(cErrorType.not_numeric);
		}
		arg2 = arg2.tocNumber();
		if (arg2.type === cElementType.error) {
			return arg2;
		}
		arg2 = arg2.toNumber();

		if (arg2 < 1) {
			return new cError(cErrorType.not_numeric);
		}

		let arg3 = arg[2] ? arg[2] : new cError(cErrorType.not_available);
		if (cElementType.cellsRange === arg3.type || cElementType.cellsRange3D === arg3.type) {
			arg3 = arg3.getValueByRowCol(0,0);
		} else if (cElementType.array === arg3.type) {
			arg3 = arg3.getElementRowCol(0, 0);
		}

		let res = new cArray();
		if (cElementType.cellsRange === arg1.type || cElementType.cellsRange3D === arg1.type || cElementType.array === arg1.type) {
			let rowCounter = 0, colCounter = 0;
			arg1.foreach2(function (val) {
				if (toCol) {
					/*if (res.array.l && res.array[res.array.length - 1].length === arg2) {
						res.addRow();
					}*/
					if (rowCounter === arg2) {
						colCounter++;
						rowCounter = 0;
					}
					if (!res.array[rowCounter]) {
						res.array[rowCounter] = [];
					}
					res.array[rowCounter][colCounter] = val;
					rowCounter++;
				} else {
					if (res.array[res.array.length - 1] && res.array[res.array.length - 1].length === arg2) {
						res.addRow();
					}
					res.addElement(val);
				}
			});
			if (toCol) {
				res.recalculate();
			}
			res.fillMatrix(arg3);
		} else {
			if (cElementType.cell === arg1.type || cElementType.cell3D === arg1.type) {
				arg1 = arg1.getValue();
			}
			res.addElement(arg1);
		}
		return res;
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWRAPROWS() {
	}

	//***array-formula***
	cWRAPROWS.prototype = Object.create(cBaseFunction.prototype);
	cWRAPROWS.prototype.constructor = cWRAPROWS;
	cWRAPROWS.prototype.name = 'WRAPROWS';
	cWRAPROWS.prototype.argumentsMin = 2;
	cWRAPROWS.prototype.argumentsMax = 3;
	cWRAPROWS.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWRAPROWS.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cWRAPROWS.prototype.argumentsType = [argType.any/*vector*/, argType.number, argType.any];
	cWRAPROWS.prototype.isXLFN = true;
	cWRAPROWS.prototype.Calculate = function (arg) {
		return wrapRowsCols(arg, arguments[1]);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cWRAPCOLS() {
	}

	//***array-formula***
	cWRAPCOLS.prototype = Object.create(cBaseFunction.prototype);
	cWRAPCOLS.prototype.constructor = cWRAPCOLS;
	cWRAPCOLS.prototype.name = 'WRAPCOLS';
	cWRAPCOLS.prototype.argumentsMin = 2;
	cWRAPCOLS.prototype.argumentsMax = 3;
	cWRAPCOLS.prototype.numFormat = AscCommonExcel.cNumFormatNone;
	cWRAPCOLS.prototype.returnValueType = AscCommonExcel.cReturnFormulaType.array;
	cWRAPCOLS.prototype.argumentsType = [argType.any/*vector*/, argType.number, argType.any];
	cWRAPCOLS.prototype.isXLFN = true;
	cWRAPCOLS.prototype.Calculate = function (arg) {
		return wrapRowsCols(arg, arguments[1], true);
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cXMATCH() {
	}
		
	//***array-formula***
	cXMATCH.prototype = Object.create(cBaseFunction.prototype);
	cXMATCH.prototype.constructor = cXMATCH;
	cXMATCH.prototype.name = 'XMATCH';
	cXMATCH.prototype.argumentsMin = 2;
	cXMATCH.prototype.argumentsMax = 4;
	cXMATCH.prototype.arrayIndexes = {1: 1};
	cXMATCH.prototype.argumentsType = [argType.any, argType.reference, argType.number, argType.number];
	cXMATCH.prototype.isXLFN = true;
	cXMATCH.prototype.Calculate = function (arg) {
		arg[4] = true;
		return g_oMatchCache.calculate(arg, arguments[1]);
	};

	var g_oVLOOKUPCache = new VHLOOKUPCache(false);
	var g_oHLOOKUPCache = new VHLOOKUPCache(true);
	var g_oMatchCache = new MatchCache();
	var g_oLOOKUPCache = new LOOKUPCache();

//----------------------------------------------------------export----------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].g_oVLOOKUPCache = g_oVLOOKUPCache;
	window['AscCommonExcel'].g_oHLOOKUPCache = g_oHLOOKUPCache;
	window['AscCommonExcel'].g_oMatchCache = g_oMatchCache;
})(window);
