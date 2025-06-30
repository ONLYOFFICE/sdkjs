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

(function (window, undefined) {
	/*
	 * Import
	 * -----------------------------------------------------------------------------
	 */

	var c_oAscInsertOptions = Asc.c_oAscInsertOptions;
	var c_oAscDeleteOptions = Asc.c_oAscDeleteOptions;
	var cElementType = AscCommonExcel.cElementType;

	var EDataValidationType = {
		None: 0,
		Custom: 1,
		Date: 2,
		Decimal: 3,
		List: 4,
		TextLength: 5,
		Time: 6,
		Whole: 7
	};
	var EDataValidationErrorStyle = {
		Stop: 0,
		Warning: 1,
		Information: 2
	};
	var EDataValidationImeMode = {
		NoControl: 0,
		Off: 1,
		On: 2,
		Disabled: 3,
		Hiragana: 4,
		FullKatakana: 5,
		HalfKatakana: 6,
		FullAlpha: 7,
		HalfAlpha: 8,
		FullHangul: 9,
		HalfHangul: 10
	};
	var EDataValidationOperator = {
		Between: 0,
		NotBetween: 1,
		Equal: 2,
		NotEqual: 3,
		LessThan: 4,
		LessThanOrEqual: 5,
		GreaterThan: 6,
		GreaterThanOrEqual: 7
	};

	function checkIntegerType(val) {
		return val && AscCommonExcel.cElementType.number === val.type;
	}

	function isNum(value) {
		return !isNaN(parseFloat(value)) && isFinite(value);
	}

	function CDataFormula(value) {
		this.text = value;
		this._formula = null;
	}

	CDataFormula.prototype._init = function (ws, locale, doNotBuildDependencies) {
		if (this._formula || this.text == null) {
			return;
		}
		var t = this;
		var formulaText = isNum(this.text) ? this.text + "" : this.text;
		this._formula = new AscCommonExcel.parserFormula(formulaText, this, ws);
		if (!locale) {
			AscCommonExcel.executeInR1C1Mode(false, function () {
				t._formula.parse();
			});
		} else {
			this._formula.parse(locale);
		}

		if (!doNotBuildDependencies) {
			this._formula.buildDependencies();
		}
	};
	CDataFormula.prototype.clone = function () {
		var res = new CDataFormula();
		res.text = this.text;
		//this._formula = null;
		return res;
	};
	CDataFormula.prototype.onFormulaEvent = function (type, eventData) {
		if (AscCommon.c_oNotifyParentType.ChangeFormula === type) {
			this.text = eventData.assemble;
		}
	};
	CDataFormula.prototype.getValue = function (ws, returnRaw, local, offset) {
		this._init(ws, local);
		if (offset) {
			this._formula.changeOffset(offset);
		}
		var activeCell = ws.getSelection().activeCell;
		var res = this._formula.calculate(null, new Asc.Range(activeCell.col, activeCell.row, activeCell.col, activeCell.row));
		return returnRaw ? this._formula.simplifyRefType(res) : res;
	};
	CDataFormula.prototype.Write_ToBinary2 = function (writer) {
		if (null !== this.text) {
			writer.WriteBool(true);
			writer.WriteString2(this.text);
		} else {
			writer.WriteBool(false);
		}
	};
	CDataFormula.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.text = reader.GetString2();
		}
	};
	CDataFormula.prototype.asc_getValue = function () {
		return this.text;
	};
	CDataFormula.prototype.asc_setValue = function (val) {
		this.text = val;
	};
	CDataFormula.prototype.setOffset = function (offset) {
		if (this._formula) {
			this.text = this._formula.changeOffset(offset, null, true).assemble(true);
		}
	};


	function CDataValidation() {
		this.ranges = null;

		this.allowBlank = false;
		this.showDropDown = false; // Excel considers this field to be a refusal to display
		this.showErrorMessage = false;
		this.showInputMessage = false;
		this.type = EDataValidationType.None;
		this.errorStyle = EDataValidationErrorStyle.Stop;
		this.imeMode = EDataValidationImeMode.NoControl;
		this.operator = EDataValidationOperator.Between;
		this.error = null;
		this.errorTitle = null;
		this.prompt = null;
		this.promptTitle = null;

		this.formula1 = null;
		this.formula2 = null;

		//while on open
		this.list = null;

		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		this._tempSelection = null;

		return this;
	}

	CDataValidation.prototype.Get_Id = function () {
		return this.Id;
	};
	CDataValidation.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_DataValidation;
	};
	CDataValidation.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.DataValidationInner;
	};
	CDataValidation.prototype._init = function (ws, doNotBuildDependencies) {
		//list convert to formula
		if (this.list) {
			if (!this.formula1 && !this.formula2) {
				this.formula1 = this.list;
			}
			this.list = null;
		}

		if (this.formula1) {
			this.formula1._init(ws, null, doNotBuildDependencies);
		}
		if (this.formula2) {
			this.formula2._init(ws, null, doNotBuildDependencies);
		}
	};
	CDataValidation.prototype._buildDependencies = function (ws, doNotBuildDependencies) {
		if (this.formula1 && this.formula1._formula) {
			this.formula1._formula.buildDependencies();
		}
		if (this.formula2 && this.formula2._formula) {
			this.formula2._formula.buildDependencies();
		}
	};
	CDataValidation.prototype.clone = function (needSaveId) {
		var res = new CDataValidation();
		if (this.ranges) {
			res.ranges = [];
			for (var i = 0; i < this.ranges.length; ++i) {
				res.ranges.push(this.ranges[i].clone());
			}
		}
		res.allowBlank = this.allowBlank;
		res.showDropDown = this.showDropDown;
		res.showErrorMessage = this.showErrorMessage;
		res.showInputMessage = this.showInputMessage;
		res.type = this.type;
		res.errorStyle = this.errorStyle;
		res.imeMode = this.imeMode;
		res.operator = this.operator;
		res.error = this.error;
		res.errorTitle = this.errorTitle;
		res.prompt = this.prompt;
		res.promptTitle = this.promptTitle;
		res.formula1 = this.formula1 ? this.formula1.clone() : null;
		res.formula2 = this.formula2 ? this.formula2.clone() : null;
		if (needSaveId) {
			res.Id = this.Id;
		}
		return res;
	};
	CDataValidation.prototype.set = function (val, ws) {
		this.allowBlank = this.checkProperty(this.allowBlank, val.allowBlank, AscCH.historyitem_DataValidation_AllowBlank, ws);
		this.showDropDown = this.checkProperty(this.showDropDown, val.showDropDown, AscCH.historyitem_DataValidation_ShowDropDown, ws);
		this.showErrorMessage = this.checkProperty(this.showErrorMessage, val.showErrorMessage, AscCH.historyitem_DataValidation_ShowErrorMessage, ws);
		this.showInputMessage = this.checkProperty(this.showInputMessage, val.showInputMessage, AscCH.historyitem_DataValidation_ShowInputMessage, ws);
		this.type = this.checkProperty(this.type, val.type, AscCH.historyitem_DataValidation_Type, ws);
		this.errorStyle = this.checkProperty(this.errorStyle, val.errorStyle, AscCH.historyitem_DataValidation_ErrorStyle, ws);
		this.imeMode = this.checkProperty(this.imeMode, val.imeMode, AscCH.historyitem_DataValidation_ImeMode, ws);
		this.operator = this.checkProperty(this.operator, val.operator, AscCH.historyitem_DataValidation_Operator, ws);
		this.error = this.checkProperty(this.error, val.error, AscCH.historyitem_DataValidation_Error, ws);
		this.errorTitle = this.checkProperty(this.errorTitle, val.errorTitle, AscCH.historyitem_DataValidation_ErrorTitle, ws);
		this.prompt = this.checkProperty(this.prompt, val.prompt, AscCH.historyitem_DataValidation_Prompt, ws);
		this.promptTitle = this.checkProperty(this.promptTitle, val.promptTitle, AscCH.historyitem_DataValidation_PromptTotle, ws);
		this.formula1 = this.checkProperty(this.formula1, val.formula1, AscCH.historyitem_DataValidation_Formula1, ws);
		this.formula2 = this.checkProperty(this.formula2, val.formula2, AscCH.historyitem_DataValidation_Formula2, ws);
	};
	CDataValidation.prototype.checkProperty = function (propOld, propNew, type, ws) {
		var isFormulaType = type === AscCH.historyitem_DataValidation_Formula1 || type === AscCH.historyitem_DataValidation_Formula12;
		var _propOld = isFormulaType ? propOld && propOld.text : propOld;
		var _propNew = isFormulaType ? propNew && propNew.text : propNew;
		if (_propOld !== _propNew && undefined !== _propNew) {
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, type,
				ws.getId(), null, new AscCommonExcel.UndoRedoData_DataValidation(this.Id, _propOld, _propNew));
			return propNew;
		}
		return propOld;
	};
	CDataValidation.prototype.isEqual = function (obj) {
		var errorEqual = obj.error === this.error && this.errorStyle === obj.errorStyle && this.showErrorMessage === obj.showErrorMessage;
		var compareFormulas = function (_f1, _f2) {
			if (_f1 === _f2) {
				return true;
			} else if (_f1 && _f2 && _f1.text === _f2.text) {
				return true;
			}
			return false;
		};

		if (errorEqual) {
			if (obj.allowBlank === this.allowBlank && obj.showDropDown === this.showDropDown && obj.showInputMessage === this.showInputMessage) {
				if (obj.type === this.type && obj.imeMode === this.imeMode && obj.operator === this.operator && obj.prompt === this.prompt) {
					if (obj.promptTitle === this.promptTitle && compareFormulas(obj.formula1, this.formula1) && compareFormulas(obj.formula2, this.formula2)) {
						return true;
					}
				}
			}
		}
		return false;
	};
	CDataValidation.prototype.setOffset = function (offset) {
		if (this.formula1) {
			this.formula1.setOffset(offset);
		}
		if (this.formula2) {
			this.formula2.setOffset(offset);
		}
	};
	CDataValidation.prototype.Write_ToBinary2 = function (writer) {
		//for wrapper
		//writer.WriteLong(this.getObjectType());

		if (null != this.ranges) {
			writer.WriteBool(true);
			writer.WriteLong(this.ranges.length);
			for (var i = 0; i < this.ranges.length; i++) {
				writer.WriteLong(this.ranges[i].r1);
				writer.WriteLong(this.ranges[i].c1);
				writer.WriteLong(this.ranges[i].r2);
				writer.WriteLong(this.ranges[i].c2);
			}
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.allowBlank);
		writer.WriteBool(this.showDropDown);
		writer.WriteBool(this.showErrorMessage);
		writer.WriteBool(this.showInputMessage);
		writer.WriteLong(this.type);
		writer.WriteLong(this.errorStyle);
		writer.WriteLong(this.imeMode);
		writer.WriteLong(this.operator);

		if (null != this.error) {
			writer.WriteBool(true);
			writer.WriteString2(this.error);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.errorTitle) {
			writer.WriteBool(true);
			writer.WriteString2(this.errorTitle);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.prompt) {
			writer.WriteBool(true);
			writer.WriteString2(this.prompt);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.promptTitle) {
			writer.WriteBool(true);
			writer.WriteString2(this.promptTitle);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.formula1) {
			writer.WriteBool(true);
			this.formula1.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.formula2) {
			writer.WriteBool(true);
			this.formula2.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
	};
	CDataValidation.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			var length = reader.GetULong();
			for (var i = 0; i < length; ++i) {
				if (!this.ranges) {
					this.ranges = [];
				}
				var r1 = reader.GetLong();
				var c1 = reader.GetLong();
				var r2 = reader.GetLong();
				var c2 = reader.GetLong();
				this.ranges.push(new Asc.Range(c1, r1, c2, r2));
			}
		}

		this.allowBlank = reader.GetBool();
		this.showDropDown = reader.GetBool();
		this.showErrorMessage = reader.GetBool();
		this.showInputMessage = reader.GetBool();
		this.type = reader.GetLong();
		this.errorStyle = reader.GetLong();
		this.imeMode = reader.GetLong();
		this.operator = reader.GetLong();

		if (reader.GetBool()) {
			this.error = reader.GetString2();
		}
		if (reader.GetBool()) {
			this.errorTitle = reader.GetString2();
		}
		if (reader.GetBool()) {
			this.prompt = reader.GetString2();
		}
		if (reader.GetBool()) {
			this.promptTitle = reader.GetString2();
		}
		var obj;
		if (reader.GetBool()) {
			obj = new CDataFormula();
			obj.Read_FromBinary2(reader);
			this.formula1 = obj;
		}
		if (reader.GetBool()) {
			obj = new CDataFormula();
			obj.Read_FromBinary2(reader);
			this.formula2 = obj;
		}
	};
	CDataValidation.prototype.setSqRef = function (sqRef) {
		this.ranges = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
	};
	CDataValidation.prototype.contains = function (c, r) {
		if (this.ranges) {
			for (var i = 0; i < this.ranges.length; ++i) {
				if (this.ranges[i].contains(c, r)) {
					return true;
				}
			}
		}
		return false;
	};
	CDataValidation.prototype.containsRange = function (range) {
		if (this.ranges) {
			for (var i = 0; i < this.ranges.length; ++i) {
				if (this.ranges[i].containsRange(range)) {
					return true;
				}
			}
		}
		return false;
	};
	CDataValidation.prototype.intersection = function (range) {
		if (this.ranges) {
			for (var i = 0; i < this.ranges.length; ++i) {
				if (this.ranges[i].intersection(range)) {
					return true;
				}
			}
		}
		return false;
	};
	CDataValidation.prototype.getIntersections = function (range, offset) {
		var res = [];
		if (this.ranges) {
			for (var i = 0; i < this.ranges.length; ++i) {
				var intersection = this.ranges[i].intersection(range);
				if (intersection) {
					if (offset) {
						intersection.setOffset(offset);
					}
					res.push(intersection);
				}
			}
		}
		return res.length ? res : null;
	};
	CDataValidation.prototype.checkValue = function (cell, ws) {
		if (!this.showErrorMessage || EDataValidationType.None === this.type) {
			return true;
		}

		let cleanFormulaCaches = function () {
			AscCommonExcel.g_oLOOKUPCache.clean();
			AscCommonExcel.g_oVLOOKUPCache.clean();
			AscCommonExcel.g_oHLOOKUPCache.clean();
			AscCommonExcel.g_oMatchCache.clean();
			AscCommonExcel.g_oSUMIFSCache.clean();
			AscCommonExcel.g_oFormulaRangesCache.clean();
			AscCommonExcel.g_oCountIfCache.clean();
		};

		let cellType = cell.getType();
		let val = cell.getValueWithoutFormat();

		if (EDataValidationType.List === this.type) {
			let list = this._getListValues(ws);
			let aValue = list[0];
			if (!aValue) {
				return false;
			}
			let aData = list[1];
			if (aData) {
				for (let i = 0; i < aData.length; ++i) {
					if (aData[i].isEqualCell(cell)) {
						return true;
					}
				}
			} else {
				return -1 !== aValue.indexOf(val);
			}
		} else if (EDataValidationType.Custom === this.type) {
			cleanFormulaCaches();
			let v = this.formula1 && this.formula1.clone().getValue(ws, true, null, this.calculateOffset(ws));
			v = v && v.tocBool();
			return !!(v && AscCommonExcel.cElementType.bool === v.type && v.toBool());
		} else {
			if (EDataValidationType.TextLength === this.type) {
				val = val.length;
			} else {
				if (AscCommon.CellValueType.Number !== cellType) {
					return false;
				}
				val = Number(val);

				if (isNaN(val) || (EDataValidationType.Whole === this.type && (val >> 0) !== val)) {
					return false;
				}
			}

			cleanFormulaCaches();

			let v1 = this.formula1 && this.formula1.getValue(ws, true);
			let v2 = this.formula2 && this.formula2.getValue(ws, true);


			let res = false;
			if (v1 == null && v2 == null) {
				switch (this.type) {
					case EDataValidationType.None:
					case EDataValidationType.Date:
					case EDataValidationType.Decimal:
					case EDataValidationType.TextLength:
					case EDataValidationType.Time:
					case EDataValidationType.Whole:
						res = true;
						break;
				}
				return res;
			}

			if (!checkIntegerType(v1)) {
				return false;
			}
			v1 = v1.toNumber();

			switch (this.operator) {
				case EDataValidationOperator.Between:
					res = checkIntegerType(v2) && v1 <= val && val <= v2.toNumber();
					break;
				case EDataValidationOperator.NotBetween:
					res = checkIntegerType(v2) && !(v1 <= val && val <= v2.toNumber());
					break;
				case EDataValidationOperator.Equal:
					res = v1 === val;
					break;
				case EDataValidationOperator.NotEqual:
					res = v1 !== val;
					break;
				case EDataValidationOperator.LessThan:
					res = v1 > val;
					break;
				case EDataValidationOperator.LessThanOrEqual:
					res = v1 >= val;
					break;
				case EDataValidationOperator.GreaterThan:
					res = v1 < val;
					break;
				case EDataValidationOperator.GreaterThanOrEqual:
					res = v1 <= val;
					break;
			}
			return res;
		}
		return false;
	};
	CDataValidation.prototype._getListValues = function (ws) {
		var aValue, aData;

		var f = this.formula1;
		var offset;
		if (f && f._formula) {
			//если формула содержит ссылки на диапазоны, то в зависимости от активной области нужно их сдвинуть
			offset = this.calculateOffset(ws);
			if (offset) {
				f = f.clone();
			}
		}

		var list = f && f.getValue(ws, false, null, offset);
		if (list && AscCommonExcel.cElementType.error !== list.type) {
			if (AscCommonExcel.cElementType.string === list.type) {
				aValue = list.getValue().split(AscCommon.FormulaSeparators.functionArgumentSeparatorDef);
				if (aValue && aValue.length) {
					for (var i = 0; i < aValue.length; i++) {
						//обрезаем только вначале строки
						if (aValue[i] && aValue[i].length) {
							var pos = 0;
							while ((pos < aValue[i].length) && (aValue[i][pos] == ' ')) {
								++pos;
							}
							aValue[i] = pos ? aValue[i].substr(pos) : aValue[i];
						}
					}
				}
			} else if (AscCommonExcel.cElementType.array === list.type) {
				let seenValues = {};
				aValue = [];
				list.foreach(function (elem) {
					let _val = elem.getValue && elem.getValue();
					if (_val != null && !seenValues[_val]) {
						aValue.push(_val);
						seenValues[_val] = true;
					}
				});
			} else {
				list = list.getRange && list.getRange();
				if (list) {
					aValue = [];
					aData = [];
					let duplicatedMap = [];
					list._foreachNoEmpty(function (cell) {
						if (!cell.isNullTextString()) {
							let val = cell.getValue();
							if (!duplicatedMap[val]) {
								aValue.push(val);
								aData.push(new AscCommonExcel.CCellValue(cell));
								duplicatedMap[val] = 1;
							}
						}
					});
				}
			}
		}
		return [aValue, aData];
	};
	CDataValidation.prototype.isListValues = function () {
		return (this.type === EDataValidationType.List && !this.showDropDown);
	};
	CDataValidation.prototype.getListValues = function (ws) {
		return this.isListValues() ? this._getListValues(ws) : null;
	};
	CDataValidation.prototype.asc_checkValid = function () {
		var res = Asc.c_oAscError.ID.No;

		var _getNumber = function (_text) {
			var _val = null;
			if (!isNum(_text)) {
				var date = AscCommon.g_oFormatParser.parseDate(_text, AscCommon.g_oDefaultCultureInfo);
				if (date) {
					_val = date.value;
				}
			} else {
				_val = parseFloat(_text);
			}

			return _val;
		};

		if (this.type !== EDataValidationType.Custom && this.type !== EDataValidationType.List) {
			if (this.operator === EDataValidationOperator.Between || this.operator === EDataValidationOperator.NotBetween) {
				if (this.formula1 && this.formula2) {
					var nFormula1 = _getNumber(this.formula1.text);
					var nFormula2 = _getNumber(this.formula2.text);

					if (nFormula1 !== null && nFormula2 !== null && nFormula2 < nFormula1) {
						return Asc.c_oAscError.ID.DataValidateMinGreaterMax;
					}
				}
			}
		}

		return res;
	};
	CDataValidation.prototype.isValidDataRef = function (ws, _val, type) {
		var _checkValidType = function (val) {
			var _res = false;
			if (val.type === cElementType.cell || val.type === cElementType.cell3D) {
				_res = true;
			} else if (type === EDataValidationType.List) {
				if (val.type === cElementType.cellsRange || val.type === cElementType.cellsRange3D) {
					_res = true;
				}
			} else if (val.type === cElementType.number) {
				_res = true;
			}
			return _res;
		};

		var checkDefNames = function (_f) {
			var outStack = _f._formula.outStack;
			if (outStack && outStack.length) {
				for (var i = 0; i < outStack.length; i++) {
					if (outStack[i].type === cElementType.name && outStack[i].Calculate().type === cElementType.error) {
						return false;
					}
				}
			}
			return true;
		};

		var _checkFormulaOnError = function (fValue, _f) {
			//ошибка по именованному диапазону
			if (fValue.type === cElementType.error && fValue.errorType === AscCommonExcel.cErrorType.wrong_name && !checkDefNames(_f)) {
				return asc_error.NamedRangeNotFound;
			}

			//если ссылка на диапазон - в любом случае отдаём ошибку
			if (fValue.type === cElementType.cellsRange || fValue.type === cElementType.cellsRange3D) {
				//в случае списка допустимы строки/столбцы
				if (type === EDataValidationType.List) {
					var _bbox = fValue.getBBox0();
					if (_bbox.c1 !== _bbox.c2 && _bbox.r1 !== _bbox.r2) {
						return asc_error.DataValidateInvalidList;
					}
				} else {
					return asc_error.DataValidateInvalid;
				}
			}

			if (fValue.type === cElementType.array) {
				//в ms другой текст ошибки, мы выдаём общий
				return asc_error.DataValidateInvalid;
			}

			if (type !== EDataValidationType.Custom && !_checkValidType(fValue)) {
				return type === EDataValidationType.List ? asc_error.DataValidateInvalidList : asc_error.DataValidateNotNumeric;
			}

			//если ощибка в подсчете формулы - выдаём предупреждение
			if (fValue.type === cElementType.error) {
				return asc_error.FormulaEvaluateError;
			}

			return null;
		};

		var asc_error = Asc.c_oAscError.ID;
		var formula, fResult, isNumeric, date;
		if (_val[0] === "=") {
			formula = new CDataFormula(_val.slice(1));
			fResult = formula.getValue(ws, null, true);
			var formulaError = _checkFormulaOnError(fResult, formula);
			if (formulaError !== null) {
				return formulaError;
			}
		} else {
			isNumeric = isNum(_val);
			if (!isNumeric) {
				//проверим, может быть это дата или время
				if (type !== EDataValidationType.List) {
					date = AscCommon.g_oFormatParser.parseDate(_val, AscCommon.g_oDefaultCultureInfo);
				}
			}
		}

		var res = asc_error.No;
		switch (type) {
			case EDataValidationType.Date:
				if (fResult) {

				} else {
					if (!isNumeric) {
						if (date) {
							_val = date.value;
						} else {
							return asc_error.DataValidateInvalid;
						}
					}

					//TODO не нашёл константу на максимальную дату
					var maxDate = 2958465;
					if (isNumeric && (_val < 0 || _val > maxDate)) {
						return asc_error.DataValidateInvalid;
					}
				}

				break;
			case EDataValidationType.Decimal:
			case EDataValidationType.Whole:
				if (fResult) {

				} else {
					if (!isNumeric) {
						if (date) {
							_val = date.value;
						} else {
							return asc_error.DataValidateInvalid;
						}
					}
				}

				break;
			case EDataValidationType.List:
				if (fResult) {

				} else {

				}

				break;
			case EDataValidationType.TextLength:
				if (fResult) {

				} else {
					if (!isNumeric) {
						if (date) {
							_val = date.value;
						} else {
							return asc_error.DataValidateNotNumeric;
						}
					}
					if (_val >= 10000000000 || _val < 0) {
						return asc_error.DataValidateNegativeTextLength;
					}
				}

				break;
			case EDataValidationType.Time:
				if (fResult) {

				} else {
					if (!isNumeric) {
						if (date) {
							_val = date.value;
						} else {
							return asc_error.DataValidateInvalid;
						}
					}
					if (_val < 0 || _val >= 1) {
						return asc_error.DataValidateInvalid;
					}
				}

				break;
		}

		return res;
	};

	CDataValidation.prototype.getError = function () {
		return this.error;
	};
	CDataValidation.prototype.getErrorStyle = function () {
		return this.errorStyle;
	};
	CDataValidation.prototype.getErrorTitle = function () {
		return this.errorTitle;
	};
	CDataValidation.prototype.getAllowBlank = function () {
		return this.allowBlank;
	};
	CDataValidation.prototype.getShowDropDown = function () {
		return this.showDropDown;
	};
	CDataValidation.prototype.getShowErrorMessage = function () {
		return this.showErrorMessage;
	};
	CDataValidation.prototype.getShowInputMessage = function () {
		return this.showInputMessage;
	};
	CDataValidation.prototype.asc_getType = function () {
		return this.type;
	};
	CDataValidation.prototype.getImeMode = function () {
		return this.imeMode;
	};
	CDataValidation.prototype.getOperator = function () {
		return this.operator;
	};
	CDataValidation.prototype.getPrompt = function () {
		return this.prompt;
	};
	CDataValidation.prototype.getPromptTitle = function () {
		return this.promptTitle;
	};
	CDataValidation.prototype.getFormula1 = function () {
		return this.formula1;
	};
	CDataValidation.prototype.getFormula2 = function () {
		return this.formula2;
	};
	CDataValidation.prototype.setAllowBlank = function (newVal, addToHistory) {
		this.allowBlank = newVal;
	};
	CDataValidation.prototype.setShowDropDown = function (newVal, addToHistory) {
		this.showDropDown = newVal;
	};
	CDataValidation.prototype.setShowErrorMessage = function (newVal, addToHistory) {
		this.showErrorMessage = newVal;
	};
	CDataValidation.prototype.setShowInputMessage = function (newVal, addToHistory) {
		this.showInputMessage = newVal;
	};
	CDataValidation.prototype.setType = function (newVal, addToHistory) {
		this.type = newVal;
	};
	CDataValidation.prototype.setErrorStyle = function (newVal, addToHistory) {
		this.errorStyle = newVal;
	};
	CDataValidation.prototype.setImeMode = function (newVal, addToHistory) {
		this.imeMode = newVal;
	};
	CDataValidation.prototype.setOperator = function (newVal, addToHistory) {
		this.operator = newVal;
	};
	CDataValidation.prototype.setError = function (newVal, addToHistory) {
		this.error = newVal;
	};
	CDataValidation.prototype.setErrorTitle = function (newVal, addToHistory) {
		this.errorTitle = newVal;
	};
	CDataValidation.prototype.setPrompt = function (newVal, addToHistory) {
		this.prompt = newVal;
	};
	CDataValidation.prototype.setPromptTitle = function (newVal, addToHistory) {
		this.promptTitle = newVal;
	};
	CDataValidation.prototype.setFormula1 = function (newVal, addToHistory) {
		this.formula1 = newVal;
	};
	CDataValidation.prototype.setFormula2 = function (newVal, addToHistory) {
		this.formula2 = newVal;
	};
	CDataValidation.prototype.shift = function (bInsert, type, updateRange) {

		var _setDiff = function (_range) {
			var _newRanges, offset, tempRange, intersection, otherPart, diff;

			switch (type) {
				case c_oAscInsertOptions.InsertCellsAndShiftDown:
					tempRange = new Asc.Range(updateRange.c1, updateRange.r1, updateRange.c2, AscCommon.gc_nMaxRow0);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = updateRange.r2 - updateRange.r1 + 1;

						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						offset = new AscCommon.CellBase(bInsert ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);

					}
					break;

				case c_oAscInsertOptions.InsertCellsAndShiftRight:
					tempRange = new Asc.Range(updateRange.c1, updateRange.r1, AscCommon.gc_nMaxCol0, updateRange.r2);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = updateRange.c2 - updateRange.c1 + 1;
						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						offset = new AscCommon.CellBase(0, bInsert ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);
					}
					break;
			}

			return _newRanges;
		};

		var _offset;
		if (type === c_oAscInsertOptions.InsertCellsAndShiftDown || type === c_oAscInsertOptions.InsertRows) {
			_offset = new AscCommon.CellBase(updateRange.r2 - updateRange.r1 + 1, 0);
			if (!bInsert) {
				_offset.row = -_offset.row;
			}
		} else {
			_offset = new AscCommon.CellBase(0, updateRange.c2 - updateRange.c1 + 1);
			if (!bInsert) {
				_offset.col = -_offset.col;
			}
		}

		var newRanges = [];
		var bDel, isChanged;
		//TODO правлю ошибку. 50521 - попытаться понять, как получился такой файл.
		if (!this.ranges) {
			return -1;
		}
		for (var i = 0; i < this.ranges.length; i++) {
			if (!bInsert && updateRange.containsRange(this.ranges[i])) {
				bDel = true;
			} else {
				if (updateRange.isIntersectForShift(this.ranges[i], _offset)) {
					var cloneRange = this.ranges[i].clone();
					cloneRange.forShift(updateRange, _offset);
					newRanges.push(cloneRange);
					isChanged = true;
				} else {
					var changedRanges = _setDiff(this.ranges[i]);
					if (changedRanges) {
						newRanges = newRanges.concat(changedRanges);
						isChanged = true;
					} else {
						newRanges = newRanges.concat(this.ranges[i].clone());
					}
				}
			}
		}
		if (!newRanges.length && bDel) {
			//удаляем
			return -1;
		} else if (newRanges.length && isChanged) {
			//меняем диапазон
			return newRanges;
		}
	};

	CDataValidation.prototype.clear = function (ranges) {
		if (!this.ranges) {
			return null;
		}

		var newRanges = [];
		var isChanged;
		for (var i = 0; i < this.ranges.length; i++) {
			for (var j = 0; j < ranges.length; j++) {
				var intersection = this.ranges[i].intersection(ranges[j]);
				if (intersection) {
					isChanged = true;
					newRanges = newRanges.concat(intersection.difference(this.ranges[i]));
				} else {
					newRanges.push(this.ranges[i]);
				}
			}
		}

		return isChanged ? newRanges : null;
	};

	CDataValidation.prototype.move = function (oBBoxFrom, copyRange, offset) {
		if (!this.ranges) {
			return null;
		}

		var newRanges = [];
		var isChanged;
		for (var i = 0; i < this.ranges.length; i++) {
			var intersection = this.ranges[i].intersection(oBBoxFrom);
			if (intersection) {
				isChanged = true;
				if (!copyRange) {
					newRanges = newRanges.concat(intersection.difference(this.ranges[i]));
				} else {
					newRanges.push(this.ranges[i]);
				}

				intersection.setOffset(offset);
				newRanges.push(intersection);
			} else {
				newRanges.push(this.ranges[i]);
			}
		}

		return isChanged ? newRanges : null;
	};

	CDataValidation.prototype.prepeareToPaste = function (range, offset) {
		if (!this.ranges) {
			return false;
		}

		var newRanges = [];
		for (var j = 0; j < this.ranges.length; j++) {
			var intersection = range.intersection(this.ranges[j]);
			if (intersection) {
				intersection.setOffset(offset);
				newRanges.push(intersection);
			}
		}
		if (newRanges.length) {
			this.ranges = newRanges;
			return true;
		}
		return false;
	};

	CDataValidation.prototype.applyCollaborative = function (nSheetId, collaborativeEditing) {

	};

	CDataValidation.prototype.correctToInterface = function (ws) {
		var t = this;
		var doCorrect = function (_formula) {
			var _val = _formula.text;

			//если формула
			var _isNum = isNum(_val);
			if (_val[0] === '"' || _isNum) {
				if (!_isNum) {
					_val = _formula.text = _val.slice(1, -1);
					_isNum = isNum(_val);
					if (!_isNum) {
						_val = _formula.text = _val.replace(/\"\"/g, "\"");
					}
				}

				if (_isNum) {
					//переводим в дату
					var _format;
					if (t.type === EDataValidationType.Date) {
						_format = AscCommon.oNumFormatCache.get("m/d/yyyy");
					} else if (t.type === EDataValidationType.Time) {
						_format = AscCommon.oNumFormatCache.get("h:mm:ss AM/PM");
					}
					if (_format) {
						var dateVal = _format.format(_val);
						if (dateVal && dateVal[0] && dateVal[0].text) {
							_formula.text = dateVal[0].text;
						}
					}
				}

			} else {
				if (_formula && _formula._formula) {
					//если формула содержит ссылки на диапазоны, то в зависимости от активной области нужно их сдвинуть
					var offset = t.calculateOffset(ws);
					if (offset) {
						_formula._formula.changeOffset(offset);
					}
					_formula.text = "=" + _formula._formula.assembleLocale(AscCommonExcel.cFormulaFunctionToLocale);
				} else {
					_formula.text = "=" + _val;
				}
			}
		};

		if (this.formula1) {
			doCorrect(this.formula1);
		}
		if (this.formula2) {
			doCorrect(this.formula2);
		}
	};

	CDataValidation.prototype.correctFromInterface = function (ws) {
		var t = this;

		var addQuotes = function (_val) {
			var _res;
			if (_val[0] === '"') {
				_res = _val.replace(/\"/g, "\"\"");
				_res = "\"" + _res + "\"";
			} else {
				_res = "\"" + _val + "\"";
			}
			return _res;
		};

		var doCorrect = function (_formula) {
			var _val = _formula.text;
			var isNumeric = isNum(_val);
			if (isNumeric) {
				if (t.type === EDataValidationType.List) {
					_formula.text = '"' + _formula.text + '"';
				}
			} else {
				var isDate;
				var isFormula;
				if (!isNumeric) {
					if (_val[0] === "=") {
						_val = _val.slice(1);
						_formula.text = _val;

						if (isNum(_val)) {
							if (t.type === EDataValidationType.List) {
								_val = '"' + _val + '"';
							}
							_formula.text = _val;
							return;
						}
						var _tempFormula = new CDataFormula(_val);
						isFormula = _tempFormula.getValue(ws, null, true);
					} else if (t.type !== EDataValidationType.List) {
						isDate = AscCommon.g_oFormatParser.parseDate(_val, AscCommon.g_oDefaultCultureInfo);
					}
				}

				//храним число
				if (isDate) {
					_formula.text = isDate.value;
					return;
				}

				if (!isFormula) {
					_formula.text = addQuotes(_formula.text);
				} else if (_tempFormula && _tempFormula._formula) {
					_formula.text = _tempFormula._formula.assemble();
				}
			}
		};

		if (this.formula1) {
			doCorrect(this.formula1);
		}
		if (this.formula2) {
			doCorrect(this.formula2);
		}
	};

	CDataValidation.prototype.checkFormulaStackOnCell = function (row, col) {
		var stack = this.formula1 && this.formula1._formula && this.formula1._formula.outStack;
		if (stack && stack.length) {
			for (var i = 0; i < stack.length; i++) {
				if (stack[i]) {
					if (stack[i].type === AscCommonExcel.cElementType.cell || stack[i].type === AscCommonExcel.cElementType.cellsRange) {
						if (stack[i].range && stack[i].range.bbox && stack[i].range.bbox.contains(col, row)) {
							return true;
						}
					}
				}
			}
		}
		return false;
	};

	CDataValidation.prototype.calculateOffset = function (ws) {
		if (!this.ranges) {
			return null;
		}

		var res = null;
		//находим левый верхний угол
		var _row = null, _col = null;
		for (var i = 0; i < this.ranges.length; i++) {
			if (_row === null && _col === null) {
				_row = this.ranges[i].r1;
				_col = this.ranges[i].c1;
			} else if (_row > this.ranges[i].r1) {
				_row = this.ranges[i].r1;
			} else if (_col > this.ranges[i].c1) {
				_col = this.ranges[i].c1;
			}
		}
		if (_row !== null && _col !== null) {
			var selectionRange = ws.getSelection();
			var activeCell = selectionRange.activeCell;
			res = new AscCommon.CellBase(activeCell.row - _row, activeCell.col - _col);
		}

		return res;
	};

	function CDataValidations() {
		this.disablePrompts = false;
		this.xWindow = null;
		this.yWindow = null;

		this.elems = [];

		return this;
	}

	CDataValidations.prototype.init = function (ws) {
		for (var i = 0; i < this.elems.length; ++i) {
			this.elems[i]._init(ws);
		}
	};
	CDataValidations.prototype.clone = function () {
		var i, res = new CDataValidations();
		res.disablePrompts = this.disablePrompts;
		res.xWindow = this.xWindow;
		res.yWindow = this.yWindow;
		for (i = 0; i < this.elems.length; ++i) {
			res.elems.push(this.elems[i].clone());
		}
		return res;
	};
	CDataValidations.prototype.shift = function (ws, bInsert, type, updateRange, addToHistory) {
		for (var i = 0; i < this.elems.length; i++) {
			var isUpdate = this.elems[i].shift(bInsert, type, updateRange);
			if (isUpdate === -1) {
				if (this.delete(ws, this.elems[i].Id, addToHistory)) {
					i--;
				}
			} else if (isUpdate) {
				var to = this.elems[i].clone();
				to.ranges = isUpdate;
				this.change(ws, this.elems[i], to, addToHistory);
			}
		}
	};

	CDataValidations.prototype.add = function (ws, val, addToHistory) {
		this.elems.push(val);
		if (addToHistory) {
			History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DataValidationAdd, ws.getId(), null,
				new AscCommonExcel.UndoRedoData_DataValidation(val.Id, null, val));
		}
	};

	CDataValidations.prototype.change = function (ws, from, to, addToHistory) {
		if (!to || !to.ranges || !to.ranges.length) {
			return;
		}
		to.Id = from.Id;
		for (var i = 0; i < this.elems.length; i++) {
			if (this.elems[i].Id === to.Id) {
				this.elems[i] = to;
			}
		}
		if (addToHistory) {
			History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DataValidationChange, ws.getId(), null,
				new AscCommonExcel.UndoRedoData_DataValidation(from.Id, from, to));
		}
	};

	CDataValidations.prototype.delete = function (ws, id, addToHistory) {
		var deleteElem;
		for (var i = 0; i < this.elems.length; i++) {
			if (this.elems[i].Id === id) {
				deleteElem = this.elems[i];
				this.elems.splice(i, 1);
			}
		}

		if (addToHistory && deleteElem) {
			History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DataValidationDelete, ws.getId(), null,
				new AscCommonExcel.UndoRedoData_DataValidation(deleteElem.Id, deleteElem, null));
		}

		return deleteElem;
	};

	CDataValidations.prototype.getById = function (id) {
		for (var i = 0; i < this.elems.length; i++) {
			if (this.elems[i].Id === id) {
				return {data: this.elems[i], index: i};
			}
		}
	};

	CDataValidations.prototype.getIntersections = function (ranges) {
		//выделяем несколько групп
		//первая - если вся активная область находится в пределах одного dataValidation
		//вторая - если пересекаемся с dataValidation

		var checkAdd = function (arr, obj) {
			for (var n = 0; n < arr.length; n++) {
				if (arr[n].isEqual(obj)) {
					return true;
				}
			}
			return false;
		};

		var intersectionArr = [];
		var containArr = [];
		if (this.elems) {
			for (var i = 0; i < this.elems.length; i++) {
				var dataValidation = this.elems[i];

				for (var j = 0; j < ranges.length; j++) {
					if (dataValidation.intersection(ranges[j])) {
						if (dataValidation.containsRange(ranges[j])) {
							if (!checkAdd(dataValidation, containArr)) {
								containArr.push(dataValidation);
							}
						} else {
							if (!checkAdd(dataValidation, intersectionArr)) {
								intersectionArr.push(dataValidation);
							}
						}
					}
				}
			}
		}

		return {intersection: intersectionArr, contain: containArr};
	};

	CDataValidations.prototype.getProps = function (ranges, doExtend, ws) {
		var _obj = this.getIntersections(ranges);
		var dataValidationIntersection = _obj.intersection;
		var dataValidationContain = _obj.contain;
		var needCheck = doExtend === undefined;

		if (needCheck) {
			//если выделено несколько диапазонов с data validation
			if (dataValidationIntersection.length > 1 || dataValidationContain.length > 1) {
				return c_oAscError.ID.MoreOneTypeDataValidate;
			}
			//если в выделение попали диапазоны как с data validation так и без
			if (dataValidationIntersection.length) {
				return c_oAscError.ID.ContainsCellsWithoutDataValidate;
			}
		}

		var getNewObject = function () {
			var _res = new window['AscCommonExcel'].CDataValidation();
			_res.showErrorMessage = true;
			_res.showInputMessage = true;
			_res.allowBlank = true;
			return _res;
		};

		//для передачи в интерфейс использую объект и модели - CDataValidation
		//если doExtend = null -> значит erase === true
		var res;
		if (doExtend === null) {
			res = getNewObject();
		} else if (doExtend !== undefined) {
			res = doExtend ? dataValidationIntersection[0].clone(true) : getNewObject();
		} else if (dataValidationContain.length === 1) {
			res = dataValidationContain[0].clone(true);
		} else {
			//возвращаем новый объект с опциями
			res = getNewObject();
		}

		res._init(ws);
		res.correctToInterface(ws);

		return res;
	};

	CDataValidations.prototype.setProps = function (ws, ranges, props) {
		var _obj = this.getIntersections(ranges);
		var instersection = _obj.intersection;
		var contain = _obj.contain;

		var prepeareAdd = function (_props, modelRanges) {
			var _dataValidation = _props.clone();
			var _ranges = [];
			var needRanges = modelRanges ? modelRanges : ranges;
			for (var i = 0; i < needRanges.length; i++) {
				_ranges.push(needRanges[i].clone());
			}
			_dataValidation.ranges = _ranges;
			_dataValidation._init(ws);
			return _dataValidation;
		};

		props.Id = AscCommon.g_oIdCounter.Get_NewId();
		props.correctFromInterface(ws);

		var equalRangeDataValidation;
		var equalDataValidation;
		var i;
		if (this.elems) {
			for (i = 0; i < this.elems.length; i++) {
				if (this._isPartOfRanges(this.elems[i].ranges, ranges)) {
					if (!equalRangeDataValidation) {
						equalRangeDataValidation = [];
					}
					equalRangeDataValidation.push(this.elems[i]);
				}
				//пока не усложняем логику и не объединяем объекты с одинаковыми настройками
				/*if (props.isEqual(this.dataValidations.elems[i])) {
					equalDataValidation = this.dataValidations.elems[i];
					break;
				}*/
			}
		}

		if (!instersection.length && !contain.length) {
			//самый простой вариант - просто добавляем новый обхект и привязываем его к активной области
			if (equalDataValidation) {
				//в данном случае расширяем диапазон
				//set
			} else {
				this.add(ws, prepeareAdd(props), true);
			}
		} else if (equalRangeDataValidation) {
			for (i = 0; i < equalRangeDataValidation.length; i++) {
				this.change(ws, equalRangeDataValidation[i], prepeareAdd(props, equalRangeDataValidation[i].ranges), true);
			}
		} else {
			var t = this;
			var _split = function (_dataValidation) {
				var _newRanges = [];

				var dataValidationRanges = _dataValidation.ranges;
				for (var i = 0; i < dataValidationRanges.length; i++) {

					var tempRanges = [];
					for (var j = 0; j < ranges.length; j++) {
						if (tempRanges.length) {
							var tempRanges2 = [];
							for (var k = 0; k < tempRanges.length; k++) {
								tempRanges2 = tempRanges2.concat(ranges[j].difference(tempRanges[k]));
							}
							tempRanges = tempRanges2;
						} else {
							tempRanges = ranges[j].difference(dataValidationRanges[i]);
						}
					}
					_newRanges = _newRanges.concat(tempRanges);
				}

				if (!_newRanges.length) {
					t.delete(ws, _dataValidation.Id, true)
				} else {
					var newDataValidation = _dataValidation.clone();
					newDataValidation.ranges = _newRanges;
					t.change(ws, _dataValidation, prepeareAdd(newDataValidation, _newRanges), true);
				}
			};

			var k;
			for (k = 0; k < instersection.length; k++) {
				_split(instersection[k]);
			}
			for (k = 0; k < contain.length; k++) {
				_split(contain[k]);
			}
			//разбиваем диапазон объектов, с которыми пересекаемся + добавляем новый
			this.add(ws, prepeareAdd(props), true);
		}
	};

	CDataValidations.prototype._isPartOfRanges = function (_ranges1, _ranges2) {
		if (_ranges1 && _ranges2 && _ranges1.length <= _ranges2.length) {
			for (var j = 0; j < _ranges1.length; j++) {
				var _equal = false;
				for (var n = 0; n < _ranges2.length; n++) {
					if (_ranges1[j].isEqual(_ranges2[n])) {
						_equal = true;
						break;
					}
				}
				if (!_equal) {
					return false;
				}
			}
		} else {
			return false;
		}

		return true;
	};

	CDataValidations.prototype._containRanges = function (_ranges1, _ranges2) {
		//проверка на то, что диапазон второго range входит в дипапазон первого
		var res = false;
		if (_ranges1 && _ranges2 && _ranges1.length && _ranges2.length) {
			for (var j = 0; j < _ranges1.length; j++) {
				//проверяем, вошёл ли целиком массив диапазонов второго в один из первых
				if (_ranges1[j].containsRanges(_ranges2)) {
					res = true;
					break;
				}
			}
		} else {
			res = false;
		}

		return res;
	};

	CDataValidations.prototype.clear = function (ws, ranges, addToHistory) {
		for (var i = 0; i < this.elems.length; i++) {
			var isEmptyRanges = !this.elems[i].ranges || !this.elems[i].ranges.length;
			if (isEmptyRanges || this._containRanges(ranges, this.elems[i].ranges)) {
				if (this.delete(ws, this.elems[i].Id, addToHistory)) {
					i--;
				}
			} else {
				var changedRanges = this.elems[i].clear(ranges);
				if (changedRanges) {
					if (!changedRanges.length) {
						if (this.delete(ws, this.elems[i].Id, addToHistory)) {
							i--;
						}
					} else {
						var newDataValidation = this.elems[i].clone();
						newDataValidation.ranges = changedRanges;
						this.change(ws, this.elems[i], newDataValidation, addToHistory);
					}
				}
			}
		}
	};

	CDataValidations.prototype.move = function (ws, oBBoxFrom, oBBoxTo, copyRange, offset) {
		for (var i = 0; i < this.elems.length; i++) {
			var changedRanges = this.elems[i].move(oBBoxFrom, copyRange, offset);
			if (changedRanges) {
				var newDataValidation = this.elems[i].clone();
				newDataValidation.ranges = changedRanges;
				this.change(ws, this.elems[i], newDataValidation, true);
			}
		}
	};

	CDataValidations.prototype.getCopyByRange = function (range, offset) {
		var res = [];
		for (var i = 0; i < this.elems.length; i++) {
			var changedRanges = this.elems[i].getIntersections(range, offset);
			if (changedRanges) {
				var newDataValidation = this.elems[i].clone();
				newDataValidation.ranges = changedRanges;
				res.push(newDataValidation);
			}
		}
		return res.length ? res : null;
	};

	CDataValidations.prototype.getIntersectionByRange = function (range) {
		var res = [];
		for (var i = 0; i < this.elems.length; i++) {
			var changedRanges = this.elems[i].getIntersections(range);
			if (changedRanges) {
				res.push({ranges: changedRanges, id: this.elems[i].Id});
			}
		}
		return res.length ? res : null;
	};

	CDataValidations.prototype.expandRanges = function (ranges) {
		var res = [];
		var _notExpandRanges = [];
		for (var k = 0; k < ranges.length; k++) {
			res[k] = ranges[k];
			for (var i = 0; i < this.elems.length; i++) {
				var _expandRange = res[k];
				var isIntersection = false;
				var tempArr = [];
				for (var j = 0; j < this.elems[i].ranges.length; j++) {
					if (this.elems[i].ranges[j].intersection(_expandRange)) {
						isIntersection = true;
						_expandRange = _expandRange.union(this.elems[i].ranges[j]);
					} else {
						tempArr.push(this.elems[i].ranges[j]);
					}
				}
				if (isIntersection) {
					_notExpandRanges = _notExpandRanges.concat(tempArr);
					res[k] = _expandRange;
				}
			}
		}
		return res.concat(_notExpandRanges);
	};

	CDataValidations.prototype.getSameSettingsElems = function (_elem) {
		var res = null;
		if (!_elem) {
			return res;
		}

		for (var i = 0; i < this.elems.length; i++) {
			if (this.elems[i].isEqual(_elem)) {
				if (!res) {
					res = [];
				}
				res.push(this.elems[i]);
			}
		}

		return res;
	};


	/*
	 * Export
	 * -----------------------------------------------------------------------------
	 */
	var prot;
	window['Asc'] = window['Asc'] || {};
	window['Asc']['c_oAscEDataValidationType'] = window['Asc'].EDataValidationType = EDataValidationType;
	prot = EDataValidationType;
	prot['None'] = prot.None;
	prot['Custom'] = prot.Custom;
	prot['Date'] = prot.Date;
	prot['Decimal'] = prot.Decimal;
	prot['List'] = prot.List;
	prot['TextLength'] = prot.TextLength;
	prot['Time'] = prot.Time;
	prot['Whole'] = prot.Whole;

	window['Asc']['c_oAscEDataValidationErrorStyle'] = window['Asc'].EDataValidationErrorStyle = EDataValidationErrorStyle;
	prot = EDataValidationErrorStyle;
	prot['Stop'] = prot.Stop;
	prot['Warning'] = prot.Warning;
	prot['Information'] = prot.Information;

	window['Asc'].EDataValidationImeMode = EDataValidationImeMode;

	window['Asc']['CDataFormula'] = window['Asc'].CDataFormula = CDataFormula;
	prot = CDataFormula.prototype;
	prot['asc_getValue'] = prot.asc_getValue;
	prot['asc_setValue'] = prot.asc_setValue;

	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CDataValidation = CDataValidation;
	prot = CDataValidation.prototype;
	prot['asc_getError'] = prot.getError;
	prot['asc_getErrorStyle'] = prot.getErrorStyle;
	prot['asc_getErrorTitle'] = prot.getErrorTitle;
	prot['asc_getErrorTitle'] = prot.getErrorTitle;
	prot['asc_getAllowBlank'] = prot.getAllowBlank;
	prot['asc_getShowDropDown'] = prot.getShowDropDown;
	prot['asc_getShowErrorMessage'] = prot.getShowErrorMessage;
	prot['asc_getShowInputMessage'] = prot.getShowInputMessage;
	prot['asc_getType'] = prot.asc_getType;
	//prot['asc_getImeMode'] = prot.getImeMode;
	prot['asc_getOperator'] = prot.getOperator;
	prot['asc_getPrompt'] = prot.getPrompt;
	prot['asc_getPromptTitle'] = prot.getPromptTitle;
	prot['asc_getFormula1'] = prot.getFormula1;
	prot['asc_getFormula2'] = prot.getFormula2;

	prot['asc_setError'] = prot.setError;
	prot['asc_setErrorStyle'] = prot.setErrorStyle;
	prot['asc_setErrorTitle'] = prot.setErrorTitle;
	prot['asc_setErrorTitle'] = prot.setErrorTitle;
	prot['asc_setAllowBlank'] = prot.setAllowBlank;
	prot['asc_setShowDropDown'] = prot.setShowDropDown;
	prot['asc_setShowErrorMessage'] = prot.setShowErrorMessage;
	prot['asc_setShowInputMessage'] = prot.setShowInputMessage;
	prot['asc_setType'] = prot.setType;
	//prot['asc_setImeMode'] = prot.setImeMode;
	prot['asc_setOperator'] = prot.setOperator;
	prot['asc_setPrompt'] = prot.setPrompt;
	prot['asc_setPromptTitle'] = prot.setPromptTitle;
	prot['asc_setFormula1'] = prot.setFormula1;
	prot['asc_setFormula2'] = prot.setFormula2;
	prot['asc_checkValid'] = prot.asc_checkValid;

	window['Asc']['EDataValidationOperator'] = window['Asc'].EDataValidationOperator = EDataValidationOperator;
	prot = EDataValidationOperator;
	prot['Between'] = prot.Between;
	prot['NotBetween'] = prot.NotBetween;
	prot['Equal'] = prot.Equal;
	prot['NotEqual'] = prot.NotEqual;
	prot['LessThan'] = prot.LessThan;
	prot['LessThanOrEqual'] = prot.LessThanOrEqual;
	prot['GreaterThan'] = prot.GreaterThan;
	prot['GreaterThanOrEqual'] = prot.GreaterThanOrEqual;

	window['AscCommonExcel'].CDataValidations = CDataValidations;
})(window);
