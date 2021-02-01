/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
	var FT_Common = AscFonts.FT_Common;
	var CellValueType = AscCommon.CellValueType;
	var EIconSetType = Asc.EIconSetType;

	/**
	 * Отвечает за условное форматирование
	 * -----------------------------------------------------------------------------
	 *
	 * @constructor
	 * @memberOf Asc
	 */
	function CConditionalFormatting() {
		this.pivot = false;
		this.ranges = null;
		this.aRules = [];

		return this;
	}

	CConditionalFormatting.prototype.setSqRef = function (sqRef) {
		this.ranges = AscCommonExcel.g_oRangeCache.getRangesFromSqRef(sqRef);
	};
	CConditionalFormatting.prototype.isValid = function () {
		//todo more checks
		return this.ranges && this.ranges.length > 0;
	};
	CConditionalFormatting.prototype.initRules = function () {
		for (var i = 0; i < this.aRules.length; ++i) {
			this.aRules[i].updateConditionalFormatting(this);
		}
	};

	//todo need another approach
	function CConditionalFormattingFormulaParent(ws, rule, isDefName) {
		this.ws = ws;
		this.rule = rule;
		this.isDefName = isDefName;
	}

	CConditionalFormattingFormulaParent.prototype.onFormulaEvent = function (type, eventData) {
		if (AscCommon.c_oNotifyParentType.IsDefName === type && this.isDefName) {
			return {bbox: this.rule.getBBox(), ranges: this.rule.ranges};
		} else if (AscCommon.c_oNotifyParentType.Change === type) {
			this.ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(this.rule.ranges));
		}
	};
	CConditionalFormattingFormulaParent.prototype.clone = function () {
		return new CConditionalFormattingFormulaParent(this.ws, this.rule, this.isDefName);
	};

	function CConditionalFormattingRule() {
		this.aboveAverage = true;
		this.activePresent = false;
		this.bottom = false;
		this.dxf = null;
		this.equalAverage = false;
		this.id = AscCommon.g_oIdCounter.Get_NewId();
		this.operator = null;
		this.percent = false;
		this.priority = null;
		this.rank = null;
		this.stdDev = null;
		this.stopIfTrue = false;
		this.text = null;
		this.timePeriod = null;
		this.type = null;

		this.aRuleElements = [];

		// from CConditionalFormatting
		// Combined all the rules into one array to sort the priorities,
		// so they transferred these properties to the rule
		this.pivot = false;
		this.ranges = null;

		this.isLock = null;

		return this;
	}

	CConditionalFormattingRule.prototype.Get_Id = function () {
		return this.id;
	};

	CConditionalFormattingRule.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.CFDataInner;
	};

	CConditionalFormattingRule.prototype.clone = function () {
		var i, res = new CConditionalFormattingRule();
		res.aboveAverage = this.aboveAverage;
		res.bottom = this.bottom;
		if (this.dxf) {
			res.dxf = this.dxf.clone();
		}
		res.equalAverage = this.equalAverage;
		res.operator = this.operator;
		res.percent = this.percent;
		res.priority = this.priority;
		res.rank = this.rank;
		res.stdDev = this.stdDev;
		res.stopIfTrue = this.stopIfTrue;
		res.text = this.text;
		res.timePeriod = this.timePeriod;
		res.type = this.type;

		res.updateConditionalFormatting(this);

		for (i = 0; i < this.aRuleElements.length; ++i) {
			res.aRuleElements.push(this.aRuleElements[i].clone());
		}
		return res;
	};
	CConditionalFormattingRule.prototype.Write_ToBinary2 = function (writer) {
		//for wrapper
		//writer.WriteLong(this.getObjectType());

		writer.WriteBool(this.aboveAverage);
		writer.WriteBool(this.activePresent);
		writer.WriteBool(this.bottom);

		if (null != this.dxf) {
			var dxf = this.dxf;
			writer.WriteBool(true);
			var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(writer);
			oBinaryStylesTableWriter.bs.WriteItem(0, function () {
				oBinaryStylesTableWriter.WriteDxf(dxf);
			});
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.equalAverage);

		if (null != this.operator) {
			writer.WriteBool(true);
			writer.WriteLong(this.operator);
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.percent);

		if (null != this.priority) {
			writer.WriteBool(true);
			writer.WriteLong(this.priority);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.rank) {
			writer.WriteBool(true);
			writer.WriteLong(this.rank);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.stdDev) {
			writer.WriteBool(true);
			writer.WriteLong(this.stdDev);
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.stopIfTrue);

		if (null != this.text) {
			writer.WriteBool(true);
			writer.WriteString2(this.text);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.timePeriod) {
			writer.WriteBool(true);
			writer.WriteString2(this.timePeriod);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.type) {
			writer.WriteBool(true);
			writer.WriteLong(this.type);
		} else {
			writer.WriteBool(false);
		}

		var i;
		if (null != this.aRuleElements) {
			writer.WriteBool(true);
			writer.WriteLong(this.aRuleElements.length);
			for (i = 0; i < this.aRuleElements.length; i++) {
				writer.WriteLong(this.aRuleElements[i].type);
				this.aRuleElements[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.pivot);

		if (null != this.ranges) {
			writer.WriteBool(true);
			writer.WriteLong(this.ranges.length);
			for (i = 0; i < this.ranges.length; i++) {
				writer.WriteLong(this.ranges[i].r1);
				writer.WriteLong(this.ranges[i].c1);
				writer.WriteLong(this.ranges[i].r2);
				writer.WriteLong(this.ranges[i].c2);
			}
		} else {
			writer.WriteBool(false);
		}
	};

	CConditionalFormattingRule.prototype.Read_FromBinary2 = function (reader) {
		this.aboveAverage = reader.GetBool();
		this.activePresent = reader.GetBool();
		this.bottom = reader.GetBool();

		var length, i;
		if (reader.GetBool()) {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var bsr = new AscCommonExcel.Binary_StylesTableReader(reader, wb);
			var bcr = new AscCommon.Binary_CommonReader(reader);
			var oDxf = new AscCommonExcel.CellXfs();
			reader.GetUChar();
			length = reader.GetULongLE();
			bcr.Read1(length, function (t, l) {
				return bsr.ReadDxf(t, l, oDxf);
			});
			this.dxf = oDxf;
		}

		this.equalAverage = reader.GetBool();

		if (reader.GetBool()) {
			this.operator = reader.GetLong();
		}

		this.percent = reader.GetBool();

		if (reader.GetBool()) {
			this.priority = reader.GetLong();
		}

		if (reader.GetBool()) {
			this.rank = reader.GetLong();
		}

		if (reader.GetBool()) {
			this.stdDev = reader.GetLong();
		}

		this.stopIfTrue = reader.GetBool();

		if (reader.GetBool()) {
			this.text = reader.GetString2();
		}

		if (reader.GetBool()) {
			this.timePeriod = reader.GetString2();
		}

		if (reader.GetBool()) {
			this.type = reader.GetLong();
		}

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aRuleElements) {
					this.aRuleElements = [];
				}
				var type = reader.GetLong();
				var elem;
				switch (type) {
					case Asc.ECfType.colorScale:
						elem = new CColorScale();
						break;
					case Asc.ECfType.dataBar:
						elem = new CDataBar();
						break;
					case Asc.ECfType.iconSet:
						elem = new CIconSet();
						break;
					default:
						elem = new CFormulaCF();
						break;
					//TODO ?CFormulaCF
				}
				this.aRuleElements.push(elem.Read_FromBinary2(reader));
			}
		}

		this.pivot = reader.GetBool();

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
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
	};

	CConditionalFormattingRule.prototype.set = function (val, addToHistory, ws) {

		this.aboveAverage = this.checkProperty(this.aboveAverage, val.aboveAverage, AscCH.historyitem_CFRule_SetAboveAverage, ws, addToHistory);
		this.activePresent = this.checkProperty(this.activePresent, val.activePresent, AscCH.historyitem_CFRule_SetActivePresent, ws, addToHistory);
		this.bottom = this.checkProperty(this.bottom, val.bottom, AscCH.historyitem_CFRule_SetBottom, ws, addToHistory);

		this.equalAverage = this.checkProperty(this.equalAverage, val.equalAverage, AscCH.historyitem_CFRule_SetEqualAverage, ws, addToHistory);

		this.operator = this.checkProperty(this.operator, val.operator, AscCH.historyitem_CFRule_SetOperator, ws, addToHistory);
		this.percent = this.checkProperty(this.percent, val.percent, AscCH.historyitem_CFRule_SetPercent, ws, addToHistory);
		this.priority = this.checkProperty(this.priority, val.priority, AscCH.historyitem_CFRule_SetPriority, ws, addToHistory);
		this.rank = this.checkProperty(this.rank, val.rank, AscCH.historyitem_CFRule_SetRank, ws, addToHistory);
		this.stdDev = this.checkProperty(this.stdDev, val.stdDev, AscCH.historyitem_CFRule_SetStdDev, ws, addToHistory);
		this.stopIfTrue = this.checkProperty(this.stopIfTrue, val.stopIfTrue, AscCH.historyitem_CFRule_SetStopIfTrue, ws, addToHistory);
		this.text = this.checkProperty(this.text, val.text, AscCH.historyitem_CFRule_SetText, ws, addToHistory);
		this.timePeriod = this.checkProperty(this.timePeriod, val.timePeriod, AscCH.historyitem_CFRule_SetTimePeriod, ws, addToHistory);
		this.type = this.checkProperty(this.type, val.type, AscCH.historyitem_CFRule_SetType, ws, addToHistory);
		this.pivot = this.checkProperty(this.pivot, val.pivot, AscCH.historyitem_CFRule_SetPivot, ws, addToHistory);

		var compareElements = function (_elem1, _elem2) {
			if (_elem1.length === _elem2.length) {
				for (var i = 0; i < _elem1.length; i++) {
					if (!_elem1[i].isEqual(_elem2[i])) {
						return false;
					}
				}
				return true;
			}
			return false;
		};


		if (!compareElements(this.aRuleElements, val.aRuleElements)) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoCF, AscCH.historyitem_CFRule_SetRuleElements,
					ws.getId(), null, new AscCommonExcel.UndoRedoData_CF(this.id, this.aRuleElements, val.aRuleElements));
			}

			this.aRuleElements = val.aRuleElements;
		}

		if (!(this.dxf && val.dxf && this.dxf.isEqual(val.dxf))) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoCF, AscCH.historyitem_CFRule_SetDxf,
					ws.getId(), null, new AscCommonExcel.UndoRedoData_CF(this.id, this.dxf, val.dxf));
			}

			this.xfs = val.dxf;
		}

		if (this.ranges && val.ranges && !compareElements(this.ranges, val.ranges)) {
			if (addToHistory) {
				var getUndoRedoRange = function (_ranges) {
					var needRanges = [];
					for (var i = 0; i < _ranges.length; i++) {
						needRanges.push(new AscCommonExcel.UndoRedoData_BBox(_ranges[i]));
					}
					return needRanges;
				};

				History.Add(AscCommonExcel.g_oUndoRedoCF, AscCH.historyitem_CFRule_SetRanges,
					ws.getId(), null, new AscCommonExcel.UndoRedoData_CF(this.id, getUndoRedoRange(this.ranges), getUndoRedoRange(val.ranges)));
			}
			this.ranges = val.ranges;
		}
	};

	CConditionalFormattingRule.prototype.checkProperty = function (propOld, propNew, type, ws, addToHistory) {
		if (propOld !== propNew) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoCF, type,
					ws.getId(), null, new AscCommonExcel.UndoRedoData_CF(this.id, propOld, propNew));
			}
			return propNew;
		}
		return propOld;
	};

	CConditionalFormattingRule.prototype.getTimePeriod = function () {
		var start, end;
		var now = new Asc.cDate();
		now.setUTCHours(0, 0, 0, 0);
		switch (this.timePeriod) {
			case AscCommonExcel.ST_TimePeriod.last7Days:
				now.setUTCDate(now.getUTCDate() + 1);
				end = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() - 7);
				start = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.lastMonth:
				now.setUTCDate(1);
				end = now.getExcelDate();
				now.setUTCMonth(now.getUTCMonth() - 1);
				start = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.thisMonth:
				now.setUTCDate(1);
				start = now.getExcelDate();
				now.setUTCMonth(now.getUTCMonth() + 1);
				end = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.nextMonth:
				now.setUTCDate(1);
				now.setUTCMonth(now.getUTCMonth() + 1);
				start = now.getExcelDate();
				now.setUTCMonth(now.getUTCMonth() + 1);
				end = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.lastWeek:
				now.setUTCDate(now.getUTCDate() - now.getUTCDay());
				end = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() - 7);
				start = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.thisWeek:
				now.setUTCDate(now.getUTCDate() - now.getUTCDay());
				start = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() + 7);
				end = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.nextWeek:
				now.setUTCDate(now.getUTCDate() - now.getUTCDay() + 7);
				start = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() + 7);
				end = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.yesterday:
				end = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() - 1);
				start = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.today:
				start = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() + 1);
				end = now.getExcelDate();
				break;
			case AscCommonExcel.ST_TimePeriod.tomorrow:
				now.setUTCDate(now.getUTCDate() + 1);
				start = now.getExcelDate();
				now.setUTCDate(now.getUTCDate() + 1);
				end = now.getExcelDate();
				break;
		}
		return {start: start, end: end};
	};
	CConditionalFormattingRule.prototype.getValueCellIs = function (ws, opt_parent, opt_bbox, opt_offset, opt_returnRaw) {
		var res;
		if (null !== this.text) {
			res = new AscCommonExcel.cString(this.text);
		} else if (this.aRuleElements[1]) {
			res = this.aRuleElements[1].getValue(ws, opt_parent, opt_bbox, opt_offset, opt_returnRaw);
		}
		return res;
	};
	CConditionalFormattingRule.prototype.getFormulaCellIs = function () {
		return null === this.text && this.aRuleElements[1];
	};
	CConditionalFormattingRule.prototype.cellIs = function (operator, cell, v1, v2) {
		if (operator === AscCommonExcel.ECfOperator.Operator_beginsWith ||
			operator === AscCommonExcel.ECfOperator.Operator_endsWith ||
			operator === AscCommonExcel.ECfOperator.Operator_containsText ||
			operator === AscCommonExcel.ECfOperator.Operator_notContains) {
			return this._cellIsText(operator, cell, v1);
		} else {
			return this._cellIsNumber(operator, cell, v1, v2);
		}
	};
	CConditionalFormattingRule.prototype._cellIsText = function (operator, cell, v1) {
		if (!v1 || AscCommonExcel.cElementType.empty === v1.type) {
			v1 = new AscCommonExcel.cString("");
		}
		if (AscCommonExcel.ECfOperator.Operator_notContains === operator) {
			return !this._cellIsText(AscCommonExcel.ECfOperator.Operator_containsText, cell, v1);
		}
		var cellType = cell ? cell.type : null;
		if (cellType === CellValueType.Error || AscCommonExcel.cElementType.error === v1.type) {
			return false;
		}
		var res = false;
		var cellVal = cell ? cell.getValueWithoutFormat().toLowerCase() : "";
		var v1Val = v1.toLocaleString().toLowerCase();
		switch (operator) {
			case AscCommonExcel.ECfOperator.Operator_beginsWith:
			case AscCommonExcel.ECfOperator.Operator_endsWith:
				if (AscCommonExcel.cElementType.string === v1.type && (cellType === CellValueType.String || "" === v1Val)) {
					if (AscCommonExcel.ECfOperator.Operator_beginsWith === operator) {
						res = cellVal.startsWith(v1Val);
					} else {
						res = cellVal.endsWith(v1Val);
					}
				} else {
					res = false;
				}
				break;
			case AscCommonExcel.ECfOperator.Operator_containsText:
				if ("" === cellVal) {
					res = false;
				} else {
					res = -1 !== cellVal.indexOf(v1Val);
				}
				break;
		}
		return res;
	};
	CConditionalFormattingRule.prototype._cellIsNumber = function (operator, cell, v1, v2) {
		if (!v1 || AscCommonExcel.cElementType.empty === v1.type) {
			v1 = new AscCommonExcel.cNumber(0);
		}
		if ((cell && cell.type === CellValueType.Error) || AscCommonExcel.cElementType.error === v1.type) {
			return false;
		}
		var cellVal;
		var res = false;
		switch (operator) {
			case AscCommonExcel.ECfOperator.Operator_equal:
				if (AscCommonExcel.cElementType.number === v1.type) {
					if (!cell || cell.isNullTextString()) {
						res = 0 === v1.getValue();
					} else if (cell.type === CellValueType.Number) {
						res = cell.getNumberValue() === v1.getValue();
					} else {
						res = false;
					}
				} else if (AscCommonExcel.cElementType.string === v1.type) {
					if (!cell || cell.isNullTextString()) {
						res = "" === v1.getValue().toLowerCase();
					} else if (cell.type === CellValueType.String) {
						cellVal = cell.getValueWithoutFormat().toLowerCase();
						res = cellVal === v1.getValue().toLowerCase();
					} else {
						res = false;
					}
				} else if (AscCommonExcel.cElementType.bool === v1.type) {
					if (cell && cell.type === CellValueType.Bool) {
						res = cell.getBoolValue() === v1.toBool();
					} else {
						res = false;
					}
				}
				break;
			case AscCommonExcel.ECfOperator.Operator_notEqual:
				res = !this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_equal, cell, v1);
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThan:
				if (AscCommonExcel.cElementType.number === v1.type) {
					if (!cell || cell.isNullTextString()) {
						res = 0 > v1.getValue();
					} else if (cell.type === CellValueType.Number) {
						res = cell.getNumberValue() > v1.getValue();
					} else {
						res = true;
					}
				} else if (AscCommonExcel.cElementType.string === v1.type) {
					if (!cell || cell.isNullTextString()) {
						res = "" > v1.getValue().toLowerCase();
					} else if (cell.type === CellValueType.Number) {
						res = false;
					} else if (cell.type === CellValueType.String) {
						cellVal = cell.getValueWithoutFormat().toLowerCase();
						//todo Excel uses different string compare function
						res = cellVal > v1.getValue().toLowerCase();
					} else if (cell.type === CellValueType.Bool) {
						res = true;
					}
				} else if (AscCommonExcel.cElementType.bool === v1.type) {
					if (cell && cell.type === CellValueType.Bool) {
						res = cell.getBoolValue() > v1.toBool();
					} else {
						res = false;
					}
				}
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual:
				res = this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_greaterThan, cell, v1) ||
					this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_equal, cell, v1);
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThan:
				res = !this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual, cell, v1);
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThanOrEqual:
				res = !this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_greaterThan, cell, v1);
				break;
			case AscCommonExcel.ECfOperator.Operator_between:
				res = this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual, cell, v1) &&
					this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_lessThanOrEqual, cell, v2);
				break;
			case AscCommonExcel.ECfOperator.Operator_notBetween:
				res = !this._cellIsNumber(AscCommonExcel.ECfOperator.Operator_between, cell, v1, v2);
				break;
		}
		return res;
	};
	CConditionalFormattingRule.prototype.getAverage = function (val, average, stdDev) {
		var res = false;
		/*if (this.stdDev) {
			average += (this.aboveAverage ? 1 : -1) * this.stdDev + stdDev;
		}*/
		if (this.aboveAverage) {
			res = val > average;
		} else {
			res = val < average;
		}
		res = res || (this.equalAverage && val == average);
		return res;
	};
	CConditionalFormattingRule.prototype.hasStdDev = function () {
		return null !== this.stdDev;
	};
	CConditionalFormattingRule.prototype.updateConditionalFormatting = function (cf) {
		var i;
		this.pivot = cf.pivot;
		if (cf.ranges) {
			this.ranges = [];
			for (i = 0; i < cf.ranges.length; ++i) {
				this.ranges.push(cf.ranges[i].clone());
			}
		}
	};
	CConditionalFormattingRule.prototype.getBBox = function () {
		var bbox = null;
		if (this.ranges && this.ranges.length > 0) {
			bbox = this.ranges[0].clone();
			for (var i = 1; i < this.ranges.length; ++i) {
				bbox.union2(this.ranges[i]);
			}
		}
		return bbox;
	};
	CConditionalFormattingRule.prototype.getIndexRule = function (values, ws, value) {
		var valueCFVO;
		var aCFVOs = this._getCFVOs();
		for (var i = aCFVOs.length - 1; i >= 0; --i) {
			valueCFVO = this._getValue(values, aCFVOs[i], ws);
			if (value > valueCFVO || (aCFVOs[i].Gte && value === valueCFVO)) {
				return i;
			}
		}
		return 0;
	};
	CConditionalFormattingRule.prototype.getMin = function (values, ws) {
		var aCFVOs = this._getCFVOs();
		var oCFVO = (aCFVOs && 0 < aCFVOs.length) ? aCFVOs[0] : null;
		return this._getValue(values, oCFVO, ws);
	};
	CConditionalFormattingRule.prototype.getMid = function (values, ws) {
		var aCFVOs = this._getCFVOs();
		var oCFVO = (aCFVOs && 2 < aCFVOs.length) ? aCFVOs[1] : null;
		return this._getValue(values, oCFVO, ws);
	};
	CConditionalFormattingRule.prototype.getMax = function (values, ws) {
		var aCFVOs = this._getCFVOs();
		var oCFVO = (aCFVOs && 2 === aCFVOs.length) ? aCFVOs[1] : ((aCFVOs && 2 < aCFVOs.length) ? aCFVOs[2] : null);
		return this._getValue(values, oCFVO, ws);
	};
	CConditionalFormattingRule.prototype._getCFVOs = function () {
		var oRuleElement = this.aRuleElements[0];
		return oRuleElement && oRuleElement.aCFVOs;
	};
	CConditionalFormattingRule.prototype._getValue = function (values, oCFVO, ws) {
		var res, min;
		if (oCFVO) {
			if (oCFVO.Val) {
				res = 0;
				if (null === oCFVO.formula) {
					oCFVO.formulaParent = new CConditionalFormattingFormulaParent(ws, this, false);
					oCFVO.formula = new CFormulaCF();
					oCFVO.formula.Text = oCFVO.Val;
				}
				var calcRes = oCFVO.formula.getValue(ws, oCFVO.formulaParent, null, null, true);
				if (calcRes && calcRes.tocNumber) {
					calcRes = calcRes.tocNumber();
					if (calcRes && calcRes.toNumber) {
						res = calcRes.toNumber();
					}
				}
			}
			switch (oCFVO.Type) {
				case AscCommonExcel.ECfvoType.Minimum:
					res = AscCommonExcel.getArrayMin(values);
					break;
				case AscCommonExcel.ECfvoType.Maximum:
					res = AscCommonExcel.getArrayMax(values);
					break;
				case AscCommonExcel.ECfvoType.Number:
					break;
				case AscCommonExcel.ECfvoType.Percent:
					min = AscCommonExcel.getArrayMin(values);
					res = min + (AscCommonExcel.getArrayMax(values) - min) * res / 100;
					break;
				case AscCommonExcel.ECfvoType.Percentile:
					res = AscCommonExcel.getPercentile(values, res / 100.0);
					if (AscCommonExcel.cElementType.number === res.type) {
						res = res.getValue();
					} else {
						res = AscCommonExcel.getArrayMin(values);
					}
					break;
				case AscCommonExcel.ECfvoType.Formula:
					break;
				case AscCommonExcel.ECfvoType.AutoMin:
					res = Math.min(0, AscCommonExcel.getArrayMin(values));
					break;
				case AscCommonExcel.ECfvoType.AutoMax:
					res = Math.max(0, AscCommonExcel.getArrayMax(values));
					break;
				default:
					res = -Number.MAX_VALUE;
					break;
			}
		}
		return res;
	};

	CConditionalFormattingRule.prototype.asc_getType = function () {
		return this.type;
	};
	CConditionalFormattingRule.prototype.asc_getDxf = function () {
		return this.dxf;
	};
	CConditionalFormattingRule.prototype.asc_getLocation = function () {
		var arrResult = [];
		var t = this;

		if (this.ranges) {
			var wb = Asc['editor'].wbModel;
			var isActive = true, sheet;
			for (var i = 0; i < wb.aWorksheets.length; i++) {
				if (i !== wb.nActive) {
					wb.aWorksheets[i].aConditionalFormattingRules.forEach(function (item) {
						if (item.id === t.id) {
							isActive = false;
						}
					});
					if (!isActive) {
						sheet = wb.aWorksheets[i];
						break;
					}
				}
			}

			this.ranges.forEach(function (item) {
				arrResult.push((sheet ? sheet.sName + "!" : "") + item.getAbsName());
			});
		}
		return [isActive, arrResult.join(AscCommon.FormulaSeparators.functionArgumentSeparator)];
	};
	CConditionalFormattingRule.prototype.asc_getContainsText = function () {
		if (null !== this.text) {
			return this.text;
		}
		var ruleElement = this.aRuleElements[1];
		return ruleElement && ruleElement.getFormula ? ruleElement.Text : null;
	};
	CConditionalFormattingRule.prototype.asc_getTimePeriod = function () {
		return this.timePeriod;
	};
	CConditionalFormattingRule.prototype.asc_getOperator = function () {
		return this.operator;
	};
	CConditionalFormattingRule.prototype.asc_getPriority = function () {
		return this.priority;
	};
	CConditionalFormattingRule.prototype.asc_getRank = function () {
		return this.rank;
	};
	CConditionalFormattingRule.prototype.asc_getBottom = function () {
		return this.bottom;
	};
	CConditionalFormattingRule.prototype.asc_getPercent = function () {
		return this.percent;
	};
	CConditionalFormattingRule.prototype.asc_getAboveAverage = function () {
		return this.aboveAverage;
	};
	CConditionalFormattingRule.prototype.asc_getEqualAverage = function () {
		return this.equalAverage;
	};
	CConditionalFormattingRule.prototype.asc_getStdDev = function () {
		return this.stdDev;
	};
	CConditionalFormattingRule.prototype.asc_getStopIfTrue = function () {
		return this.stopIfTrue;
	};
	CConditionalFormattingRule.prototype.asc_getText = function () {
		return this.text;
	};
	CConditionalFormattingRule.prototype.asc_getValue1 = function () {
		var ruleElement = this.aRuleElements[0];
		return ruleElement && ruleElement.getFormula ? ruleElement.Text : null;
	};
	CConditionalFormattingRule.prototype.asc_getValue2 = function () {
		var ruleElement = this.aRuleElements[1];
		return ruleElement && ruleElement.getFormula ? ruleElement.Text : null;
	};
	CConditionalFormattingRule.prototype.asc_getColorScaleOrDataBarOrIconSetRule = function () {
		if ((Asc.ECfType.dataBar === this.type || Asc.ECfType.iconSet === this.type ||
			Asc.ECfType.colorScale === this.type) && 1 === this.aRuleElements.length) {
			var res = this.aRuleElements[0];
			if (res && this.type === res.type) {
				return res;
			}
		}
		return null;
	};
	CConditionalFormattingRule.prototype.asc_getId = function () {
		return this.id;
	};
	CConditionalFormattingRule.prototype.asc_getIsLock = function () {
		return this.isLock;
	};
	CConditionalFormattingRule.prototype.asc_getPreview = function (id, text) {
		var api_sheet = Asc['editor'];
		var res;
		if (this.dxf) {
			res = this.dxf.asc_getPreview2(api_sheet, id, text);
		} else if (Asc.ECfType.colorScale === this.type && 1 === this.aRuleElements.length) {
			res = this.aRuleElements[0].asc_getPreview();
		}
		return res;
	};
	
	CConditionalFormattingRule.prototype.asc_setType = function (val) {
		this.type = val;
	};
	CConditionalFormattingRule.prototype.asc_setDxf = function (val) {
		this.dxf = val;
	};
	CConditionalFormattingRule.prototype.asc_setLocation = function (val) {
		var t = this;
		if (val) {
			val = val.split(",");
			this.ranges = [];
			val.forEach(function (item) {
				t.ranges.push(AscCommonExcel.g_oRangeCache.getAscRange(item));
			});
		}
	};

	//TODO?
	CConditionalFormattingRule.prototype.asc_setContainsText = function (val) {
		this.text = val;
		/*if (null !== this.text) {
			return this.text;
		}
		var ruleElement = this.aRuleElements[1];
		return ruleElement && ruleElement.getFormula ? ruleElement.Text : null;*/
	};

	CConditionalFormattingRule.prototype.asc_setTimePeriod = function (val) {
		this.timePeriod = val;
	};
	CConditionalFormattingRule.prototype.asc_setOperator = function (val) {
		this.operator = val;
	};
	CConditionalFormattingRule.prototype.asc_setPriority = function (val) {
		this.priority = val;
	};
	CConditionalFormattingRule.prototype.asc_setRank = function (val) {
		this.rank = val;
	};
	CConditionalFormattingRule.prototype.asc_setBottom = function (val) {
		this.bottom = val;
	};
	CConditionalFormattingRule.prototype.asc_setPercent = function (val) {
		this.percent = val;
	};
	CConditionalFormattingRule.prototype.asc_setAboveAverage = function (val) {
		this.aboveAverage = val;
	};
	CConditionalFormattingRule.prototype.asc_setEqualAverage = function (val) {
		this.equalAverage = val;
	};
	CConditionalFormattingRule.prototype.asc_setStdDev = function (val) {
		this.stdDev = val;
	};
	CConditionalFormattingRule.prototype.asc_setStopIfTrue = function (val) {
		this.stopIfTrue = val;
	};
	CConditionalFormattingRule.prototype.asc_setText = function (val) {
		this.text = val;
	};
	CConditionalFormattingRule.prototype.asc_setValue1 = function (val) {
		if (!this.aRuleElements) {
			this.aRuleElements = [];
		}
		//TODO
		/*<cfRule type="containsText" dxfId="6" priority="9" operator="containsText" text="14">
			<formula>NOT(ISERROR(SEARCH("14",A1)))</formula>
			</cfRule>
			<cfRule type="timePeriod" dxfId="7" priority="8" timePeriod="yesterday">
		<formula>FLOOR(A1,1)=TODAY()-1</formula>
		</cfRule>*/

		this.aRuleElements[0] = new CFormulaCF();
		this.aRuleElements[0].Text = val;
	};
	CConditionalFormattingRule.prototype.asc_setValue2 = function (val) {
		if (!this.aRuleElements) {
			this.aRuleElements = [];
		}
		this.aRuleElements[1] = new CFormulaCF();
		this.aRuleElements[1].Text = val;
	};

	//TODO?
	CConditionalFormattingRule.prototype.asc_setColorScaleOrDataBarOrIconSetRule = function () {
		/*if ((Asc.ECfType.dataBar === this.type || Asc.ECfType.iconSet === this.type ||
			Asc.ECfType.colorScale === this.type) && 1 === this.aRuleElements.length) {
			var res = this.aRuleElements[0];
			if (res && this.type === res.type) {
				return res;
			}
		}
		return null;*/
	};

	CConditionalFormattingRule.prototype.asc_checkScope = function (type, tableId) {
		var sheet, range;
		var api_sheet = Asc['editor'];
		var wb = api_sheet.wbModel;

		switch (type) {
			case Asc.c_oAscSelectionForCFType.selection:
				sheet = wb.getActiveWs();
				// ToDo multiselect
				range = sheet.selectionRange.getLast();
				break;
			case Asc.c_oAscSelectionForCFType.worksheet:
				break;
			case Asc.c_oAscSelectionForCFType.table:
				var oTable;
				if (tableId) {
					oTable = wb.getTableByName(tableId, true);
					if (oTable) {
						sheet = wb.aWorksheets[oTable.index];
						range = oTable.table.Ref;
					}
				} else {
					//this table
					sheet = wb.getActiveWs();
					var thisTableIndex = sheet.autoFilters.searchRangeInTableParts(sheet.selectionRange.getLast());
					if (thisTableIndex >= 0) {
						range = sheet.TableParts[thisTableIndex].Ref;
					} else {
						sheet = null;
					}
				}
				break;
			case Asc.c_oAscSelectionForCFType.pivot:
				// ToDo
				break;
		}

		if (range) {
			var multiplyRange = new AscCommonExcel.MultiplyRange(this.ranges);
			if (multiplyRange.isIntersect(range)) {
				return true;
			}
		}

		return false;
	};

	function CColorScale() {
		this.aCFVOs = [];
		this.aColors = [];

		return this;
	}

	CColorScale.prototype.type = Asc.ECfType.colorScale;
	CColorScale.prototype.clone = function () {
		var i, res = new CColorScale();
		for (i = 0; i < this.aCFVOs.length; ++i) {
			res.aCFVOs.push(this.aCFVOs[i].clone());
		}
		for (i = 0; i < this.aColors.length; ++i) {
			res.aColors.push(this.aColors[i].clone());
		}
		return res;
	};
	CColorScale.prototype.Write_ToBinary2 = function (writer) {
		//CConditionalFormatValueObject
		var i;
		if (null != this.aCFVOs) {
			writer.WriteBool(true);
			writer.WriteLong(this.aCFVOs.length);
			for (i = 0; i < this.aCFVOs.length; i++) {
				this.aCFVOs[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}
		//rgbcolor,...
		if (null != this.aColors) {
			writer.WriteBool(true);
			writer.WriteLong(this.aColors.length);
			for (i = 0; i < this.aColors.length; i++) {
				writer.WriteLong(this.aColors[i].asc_getType());
				this.aColors[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}
	};
	CColorScale.prototype.Read_FromBinary2 = function (reader) {
		var i, length, elem;
		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aCFVOs) {
					this.aCFVOs = [];
				}
				elem = new CConditionalFormatValueObject();
				this.aCFVOs.push(elem.Read_FromBinary2(reader));
			}
		}

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aColors) {
					this.aColors = [];
				}
				//TODO colors!!!
				var type = reader.GetLong();
				switch (type) {
					case Asc.ECfType.colorScale:
						break;
				}
				this.aColors.push(elem.Read_FromBinary2(reader));
			}
		}
	};
	CColorScale.prototype.asc_getCFVOs = function () {
		return this.aCFVOs;
	};
	CColorScale.prototype.asc_getColors = function () {
		var res = [];
		for (var i = 0; i < this.aColors.length; ++i) {
			res.push(Asc.colorObjToAscColor(this.aColors[i]));
		}
		return res;
	};
	CColorScale.prototype.asc_getPreview = function (api, width, height) {
		return AscCommonExcel.drawColorScalePreview(api.wb, width, height, this.aColors);
	};
	CColorScale.prototype.asc_setCFVOs = function (val) {
		this.aCFVOs = val;
	};
	CColorScale.prototype.asc_setColors = function (val) {
		this.aColors = [];
		for (var i = 0; i < val.length; ++i) {
			this.aColors.push(AscCommonExcel.CorrectAscColor(val[i]));
		}
	};

	function CDataBar() {
		this.MaxLength = 90;
		this.MinLength = 10;
		this.ShowValue = true;
		this.AxisPosition = AscCommonExcel.EDataBarAxisPosition.automatic;
		this.Gradient = true;
		this.Direction = AscCommonExcel.EDataBarDirection.context;
		this.NegativeBarColorSameAsPositive = false;
		this.NegativeBarBorderColorSameAsPositive = true;

		this.aCFVOs = [];
		this.Color = null;
		this.NegativeColor = null;
		this.BorderColor = null;
		this.NegativeBorderColor = null;
		this.AxisColor = null;
		return this;
	}

	CDataBar.prototype.type = Asc.ECfType.dataBar;
	CDataBar.prototype.clone = function () {
		var i, res = new CDataBar();
		res.MaxLength = this.MaxLength;
		res.MinLength = this.MinLength;
		res.ShowValue = this.ShowValue;
		res.AxisPosition = this.AxisPosition;
		res.Gradient = this.Gradient;
		res.Direction = this.Direction;
		res.NegativeBarColorSameAsPositive = this.NegativeBarColorSameAsPositive;
		res.NegativeBarBorderColorSameAsPositive = this.NegativeBarBorderColorSameAsPositive;
		for (i = 0; i < this.aCFVOs.length; ++i) {
			res.aCFVOs.push(this.aCFVOs[i].clone());
		}
		if (this.Color) {
			res.Color = this.Color.clone();
		}
		if (this.NegativeColor) {
			res.NegativeColor = this.NegativeColor.clone();
		}
		if (this.BorderColor) {
			res.BorderColor = this.BorderColor.clone();
		}
		if (this.NegativeBorderColor) {
			res.NegativeBorderColor = this.NegativeBorderColor.clone();
		}
		if (this.AxisColor) {
			res.AxisColor = this.AxisColor.clone();
		}
		return res;
	};
	CDataBar.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.MaxLength) {
			writer.WriteBool(true);
			writer.WriteLong(this.MaxLength);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.MinLength) {
			writer.WriteBool(true);
			writer.WriteLong(this.MinLength);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.ShowValue) {
			writer.WriteBool(true);
			writer.WriteBool(this.ShowValue);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.AxisPosition) {
			writer.WriteBool(true);
			writer.WriteLong(this.AxisPosition);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Gradient) {
			writer.WriteBool(true);
			writer.WriteBool(this.Gradient);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Direction) {
			writer.WriteBool(true);
			writer.WriteLong(this.Direction);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeBarColorSameAsPositive) {
			writer.WriteBool(true);
			writer.WriteBool(this.NegativeBarColorSameAsPositive);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeBarBorderColorSameAsPositive) {
			writer.WriteBool(true);
			writer.WriteBool(this.NegativeBarBorderColorSameAsPositive);
		} else {
			writer.WriteBool(false);
		}

		//CConditionalFormatValueObject
		if (null != this.aCFVOs) {
			writer.WriteBool(true);
			writer.WriteLong(this.aCFVOs.length);
			for (var i = 0; i < this.aCFVOs.length; i++) {
				this.aCFVOs[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}

		if (null != this.Color) {
			writer.WriteBool(true);
			writer.WriteLong(this.Color.asc_getType());
			this.Color.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.NegativeColor.asc_getType());
			this.NegativeColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.BorderColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.Color.asc_getType());
			this.BorderColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeBorderColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.NegativeBorderColor.asc_getType());
			this.NegativeBorderColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.AxisColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.AxisColor.asc_getType());
			this.AxisColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
	};
	CDataBar.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.MaxLength = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.MinLength = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.ShowValue = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.AxisPosition = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Gradient = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.Direction = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.NegativeBarColorSameAsPositive = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.NegativeBarBorderColorSameAsPositive = reader.GetLong();
		}

		var i, length, type, elem;
		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aCFVOs) {
					this.aCFVOs = [];
				}
				elem = new CConditionalFormatValueObject();
				this.aCFVOs.push(elem.Read_FromBinary2(reader));
			}
		}

		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.Color = elem.Read_FromBinary2(reader);
		}
		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.Color = elem.Read_FromBinary2(reader);
		}
		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.NegativeColor = elem.Read_FromBinary2(reader);
		}
		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.BorderColor = elem.Read_FromBinary2(reader);
		}
		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.NegativeBorderColor = elem.Read_FromBinary2(reader);
		}
		if (reader.GetBool()) {
			//TODO colors!!!
			type = reader.GetLong();
			switch (type) {
				case Asc.ECfType.colorScale:
					elem = null;
					break;
			}
			this.AxisColor = elem.Read_FromBinary2(reader);
		}
	};
	CDataBar.prototype.asc_getShowValue = function () {
		return this.ShowValue;
	};
	CDataBar.prototype.asc_getAxisPosition = function () {
		return this.AxisPosition;
	};
	CDataBar.prototype.asc_getGradient = function () {
		return this.Gradient;
	};
	CDataBar.prototype.asc_getDirection = function () {
		return this.Direction;
	};
	CDataBar.prototype.asc_getNegativeBarColorSameAsPositive = function () {
		return this.NegativeBarColorSameAsPositive;
	};
	CDataBar.prototype.asc_getNegativeBarBorderColorSameAsPositive = function () {
		return this.NegativeBarBorderColorSameAsPositive;
	};
	CDataBar.prototype.asc_getCFVOs = function () {
		return this.aCFVOs;
	};
	CDataBar.prototype.asc_getColor = function () {
		return Asc.colorObjToAscColor(this.Color);
	};
	CDataBar.prototype.asc_getNegativeColor = function () {
		return Asc.colorObjToAscColor(this.NegativeColor);
	};
	CDataBar.prototype.asc_getBorderColor = function () {
		return Asc.colorObjToAscColor(this.BorderColor);
	};
	CDataBar.prototype.asc_getNegativeBorderColor = function () {
		return Asc.colorObjToAscColor(this.NegativeBorderColor);
	};
	CDataBar.prototype.asc_getAxisColor = function () {
		return Asc.colorObjToAscColor(this.AxisColor);
	};

	CDataBar.prototype.asc_setShowValue = function (val) {
		this.ShowValue = val;
	};
	CDataBar.prototype.asc_getAxisPosition = function (val) {
		return this.AxisPosition;
	};
	CDataBar.prototype.asc_getGradient = function (val) {
		this.Gradient = val;
	};
	CDataBar.prototype.asc_getDirection = function (val) {
		this.Direction = val;
	};
	CDataBar.prototype.asc_getNegativeBarColorSameAsPositive = function (val) {
		this.NegativeBarColorSameAsPositive = val;
	};
	CDataBar.prototype.asc_getNegativeBarBorderColorSameAsPositive = function (val) {
		this.NegativeBarBorderColorSameAsPositive = val;
	};
	CDataBar.prototype.asc_getCFVOs = function (val) {
		this.aCFVOs = val;
	};
	CDataBar.prototype.asc_getColor = function (val) {
		this.Color = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_getNegativeColor = function (val) {
		this.NegativeColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_getBorderColor = function (val) {
		this.BorderColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_getNegativeBorderColor = function (val) {
		this.NegativeBorderColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_getAxisColor = function (val) {
		this.AxisColor = AscCommonExcel.CorrectAscColor(val);
	};

	function CFormulaCF() {
		this.Text = null;
		this._f = null;

		return this;
	}

	CFormulaCF.prototype.clone = function () {
		var res = new CFormulaCF();
		res.Text = this.Text;
		return res;
	};
	CFormulaCF.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.Text) {
			writer.WriteBool(true);
			writer.WriteString2(this.Text);
		} else {
			writer.WriteBool(false);
		}
	};
	CFormulaCF.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.Text = reader.GetString2();
		}
	};
	CFormulaCF.prototype.isEqual = function (val) {
		return val.Text === this.Text;
	};
	CFormulaCF.prototype.init = function (ws, opt_parent) {
		if (!this._f) {
			this._f = new AscCommonExcel.parserFormula(this.Text, opt_parent, ws);
			this._f.parse();
			if (opt_parent) {
				//todo realize removeDependencies
				this._f.buildDependencies();
			}
		}
	};
	CFormulaCF.prototype.getFormula = function (ws, opt_parent) {
		this.init(ws, opt_parent);
		return this._f;
	};
	CFormulaCF.prototype.getValue = function (ws, opt_parent, opt_bbox, opt_offset, opt_returnRaw) {
		this.init(ws, opt_parent);
		var res = this._f.calculate(null, opt_bbox, opt_offset);
		if (!opt_returnRaw) {
			res = this._f.simplifyRefType(res);
		}
		return res;
	};

	function CIconSet() {
		this.IconSet = EIconSetType.Traffic3Lights1;
		this.Percent = true;
		this.Reverse = false;
		this.ShowValue = true;

		this.aCFVOs = [];
		this.aIconSets = [];

		return this;
	}

	CIconSet.prototype.type = Asc.ECfType.iconSet;
	CIconSet.prototype.clone = function () {
		var i, res = new CIconSet();
		res.IconSet = this.IconSet;
		res.Percent = this.Percent;
		res.Reverse = this.Reverse;
		res.ShowValue = this.ShowValue;
		for (i = 0; i < this.aCFVOs.length; ++i)
			res.aCFVOs.push(this.aCFVOs[i].clone());
		for (i = 0; i < this.aIconSets.length; ++i)
			res.aIconSets.push(this.aIconSets[i].clone());
		return res;
	};
	CIconSet.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.IconSet) {
			writer.WriteBool(true);
			writer.WriteLong(this.IconSet);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Percent) {
			writer.WriteBool(true);
			writer.WriteBool(this.Percent);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Reverse) {
			writer.WriteBool(true);
			writer.WriteBool(this.Reverse);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.ShowValue) {
			writer.WriteBool(true);
			writer.WriteBool(this.ShowValue);
		} else {
			writer.WriteBool(false);
		}

		//CConditionalFormatValueObject
		var i;
		if (null != this.aCFVOs) {
			writer.WriteBool(true);
			writer.WriteLong(this.aCFVOs.length);
			for (i = 0; i < this.aCFVOs.length; i++) {
				this.aCFVOs[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}

		//new AscCommonExcel.CConditionalFormatIconSet()
		if (null != this.aIconSets) {
			writer.WriteBool(true);
			writer.WriteLong(this.aIconSets.length);
			for (i = 0; i < this.aIconSets.length; i++) {
				this.aIconSets[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}
	};
	CIconSet.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.IconSet = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Percent = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.Reverse = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.ShowValue = reader.GetBool();
		}


		var i, length, elem;
		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aCFVOs) {
					this.aCFVOs = [];
				}
				elem = new CConditionalFormatValueObject();
				this.aCFVOs.push(elem.Read_FromBinary2(reader));
			}
		}

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aIconSets) {
					this.aIconSets = [];
				}
				elem = new CConditionalFormatIconSet();
				this.aIconSets.push(elem.Read_FromBinary2(reader));
			}
		}
	};
	CIconSet.prototype.asc_getIconSet = function () {
		return this.IconSet;
	};
	CIconSet.prototype.asc_getReverse = function () {
		return this.Reverse;
	};
	CIconSet.prototype.asc_getShowValue = function () {
		return this.ShowValue;
	};
	CIconSet.prototype.asc_getCFVOs = function () {
		return this.aCFVOs;
	};
	CIconSet.prototype.asc_getIconSets = function () {
		return this.aIconSets;
	};

	CIconSet.prototype.asc_setIconSet = function (val) {
		this.IconSet = val;
	};
	CIconSet.prototype.asc_setReverse = function (val) {
		this.Reverse = val;
	};
	CIconSet.prototype.asc_setShowValue = function (val) {
		this.ShowValue = val;
	};
	CIconSet.prototype.asc_setCFVOs = function (val) {
		this.aCFVOs = val;
	};
	CIconSet.prototype.asc_setIconSets = function (val) {
		this.aIconSets = val;
	};

	function CConditionalFormatValueObject() {
		this.Gte = true;
		this.Type = null;
		this.Val = null;
		this.formulaParent = null;
		this.formula = null;

		return this;
	}

	CConditionalFormatValueObject.prototype.clone = function () {
		var res = new CConditionalFormatValueObject();
		res.Gte = this.Gte;
		res.Type = this.Type;
		res.Val = this.Val;
		res.formulaParent = this.formulaParent ? this.formulaParent.clone() : null;
		res.formula = this.formula ? this.formula.clone() : null;
		return res;
	};
	CConditionalFormatValueObject.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.Gte) {
			writer.WriteBool(true);
			writer.WriteBool(this.IconSet);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Type) {
			writer.WriteBool(true);
			writer.WriteLong(this.Type);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.Val) {
			writer.WriteBool(true);
			writer.WriteLong(this.Val);
		} else {
			writer.WriteBool(false);
		}

		//not format
		/*if (null != this.formulaParent) {
			writer.WriteBool(true);
			//writer.WriteLong(this.Color.asc_getType());
			this.formulaParent.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.formula) {
			writer.WriteBool(true);
			//writer.WriteLong(this.Color.asc_getType());
			this.formula.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}*/
	};
	CConditionalFormatValueObject.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.Gte = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.Type = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Val = reader.GetBool();
		}

		//not format
		/*var i, length, elem;
		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aCFVOs) {
					this.aCFVOs = [];
				}
				elem = new CConditionalFormatValueObject();
				this.aCFVOs.push(elem.Read_FromBinary2(reader));
			}
		}

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aIconSets) {
					this.aIconSets = [];
				}
				elem = new CConditionalFormatIconSet();
				this.aIconSets.push(elem.Read_FromBinary2(reader));
			}
		}*/
	};
	CConditionalFormatValueObject.prototype.asc_getGte = function () {
		return this.Gte;
	};
	CConditionalFormatValueObject.prototype.asc_getType = function () {
		return this.Type;
	};
	CConditionalFormatValueObject.prototype.asc_getVal = function () {
		return this.Val;
	};
	CConditionalFormatValueObject.prototype.asc_setGte = function (val) {
		this.Gte = val;
	};
	CConditionalFormatValueObject.prototype.asc_setType = function (val) {
		this.Type = val;
	};
	CConditionalFormatValueObject.prototype.asc_setVal = function (val) {
		this.Val = val;
	};

	function CConditionalFormatIconSet() {
		this.IconSet = null;
		this.IconId = null;

		return this;
	}

	CConditionalFormatIconSet.prototype.clone = function () {
		var res = new CConditionalFormatIconSet();
		res.IconSet = this.IconSet;
		res.IconId = this.IconId;
		return res;
	};
	CConditionalFormatIconSet.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.IconSet) {
			writer.WriteBool(true);
			writer.WriteBool(this.IconSet);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.IconId) {
			writer.WriteBool(true);
			writer.WriteLong(this.IconId);
		} else {
			writer.WriteBool(false);
		}
	};
	CConditionalFormatIconSet.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.IconSet = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.IconId = reader.GetLong();
		}
	};

	function CGradient(c1, c2) {
		this.MaxColorIndex = 512;
		this.base_shift = 8;

		this.c1 = c1;
		this.c2 = c2;

		this.min = this.max = 0;
		this.koef = null;
		this.r1 = this.r2 = 0;
		this.g1 = this.g2 = 0;
		this.b1 = this.b2 = 0;

		return this;
	}

	CGradient.prototype.init = function (min, max) {
		var distance = max - min;

		this.min = min;
		this.max = max;
		this.koef = distance ? this.MaxColorIndex / (2.0 * distance) : 0;
		this.r1 = this.c1.getR();
		this.g1 = this.c1.getG();
		this.b1 = this.c1.getB();
		this.r2 = this.c2.getR();
		this.g2 = this.c2.getG();
		this.b2 = this.c2.getB();
	};
	CGradient.prototype.calculateColor = function (indexColor) {
		indexColor = ((indexColor - this.min) * this.koef) >> 0;

		var r = (this.r1 + ((FT_Common.IntToUInt(this.r2 - this.r1) * indexColor) >> this.base_shift)) & 0xFF;
		var g = (this.g1 + ((FT_Common.IntToUInt(this.g2 - this.g1) * indexColor) >> this.base_shift)) & 0xFF;
		var b = (this.b1 + ((FT_Common.IntToUInt(this.b2 - this.b1) * indexColor) >> this.base_shift)) & 0xFF;
		//console.log("index=" + indexColor + ": r=" + r + " g=" + g + " b=" + b);
		return new AscCommonExcel.RgbColor((r << 16) + (g << 8) + b);
	};
	CGradient.prototype.getMinColor = function () {
		return new AscCommonExcel.RgbColor((this.r1 << 16) + (this.g1 << 8) + this.b1);
	};
	CGradient.prototype.getMaxColor = function () {
		return new AscCommonExcel.RgbColor((this.r2 << 16) + (this.g2 << 8) + this.b2);
	};

	var cDefIconSize = 16;
	var cDefIconFont = 11;

	var iCheckGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgOEMxNSAxMS44NjYgMTEuODY2IDE1IDggMTVDNC4xMzQwMSAxNSAxIDExLjg2NiAxIDhDMSA0LjEzNDAxIDQuMTM0MDEgMSA4IDFDMTEuODY2IDEgMTUgNC4xMzQwMSAxNSA4WiIgZmlsbD0iIzJFOTk1RiIvPjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNMTIuODA1MSA1LjU5MzJMNy42MzkxOCAxMi42MDQxTDQuMjQxNyA4LjY1MTg5TDUuNzU4MzMgNy4zNDgxMUw3LjUxODc1IDkuMzk1OTRMMTEuMTk1IDQuNDA2OEwxMi44MDUxIDUuNTkzMloiIGZpbGw9IndoaXRlIi8+PC9zdmc+';
	var iCheckSymbolGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTE0IDQuMjk1NzhMNi42ODIwMyAxNEwyIDguNjc4MjFMMy42MzQxNyA3LjI1NjA4TDYuNTUyMTggMTAuNTcyOEwxMi4yNjI4IDNMMTQgNC4yOTU3OFoiIGZpbGw9IiMyRTk5NUYiLz48L3N2Zz4=';
	var iCircleTwoWhiteQuarters = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNi40MDg3IDIgNC44ODI1OCAyLjYzMjE0IDMuNzU3MzYgMy43NTczNkMyLjYzMjE0IDQuODgyNTggMiA2LjQwODcgMiA4QzIgOS41OTEzIDIuNjMyMTQgMTEuMTE3NCAzLjc1NzM2IDEyLjI0MjZDNC44ODI1OCAxMy4zNjc5IDYuNDA4NyAxNCA4IDE0TDggOEw4IDJaIiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==';
	var iCircleBlack = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjwvc3ZnPg==';
	var iCircleGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzlCOUI5QiIvPjwvc3ZnPg==';
	var iCircleGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgOEMxNSAxMS44NjYgMTEuODY2IDE1IDggMTVDNC4xMzQwMSAxNSAxIDExLjg2NiAxIDhDMSA0LjEzNDAxIDQuMTM0MDEgMSA4IDFDMTEuODY2IDEgMTUgNC4xMzQwMSAxNSA4WiIgZmlsbD0iIzJFOTk1RiIvPjwvc3ZnPg==';
	var iCircleOneWhiteQuarter = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNy4yMTIwNyAyIDYuNDMxODUgMi4xNTUxOSA1LjcwMzkgMi40NTY3MkM0Ljk3NTk1IDIuNzU4MjUgNC4zMTQ1MSAzLjIwMDIxIDMuNzU3MzYgMy43NTczNkMzLjIwMDIxIDQuMzE0NTEgMi43NTgyNSA0Ljk3NTk1IDIuNDU2NzIgNS43MDM5QzIuMTU1MTkgNi40MzE4NSAyIDcuMjEyMDcgMiA4TDggOEw4IDJaIiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==';
	var iCircleLightRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48ZWxsaXBzZSBjeD0iOCIgY3k9IjgiIHJ4PSI3IiByeT0iNyIgdHJhbnNmb3JtPSJyb3RhdGUoLTE4MCA4IDgpIiBmaWxsPSIjRkY4MDgwIi8+PC9zdmc+';
	var iCircleRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgOEMxNSAxMS44NjYgMTEuODY2IDE1IDggMTVDNC4xMzQwMSAxNSAxIDExLjg2NiAxIDhDMSA0LjEzNDAxIDQuMTM0MDEgMSA4IDFDMTEuODY2IDEgMTUgNC4xMzQwMSAxNSA4WiIgZmlsbD0iI0ZGMTExMSIvPjwvc3ZnPg==';
	var iCircleThreeWhiteQuarters = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNi44MTMzMSAyIDUuNjUzMjggMi4zNTE4OSA0LjY2NjU4IDMuMDExMThDMy42Nzk4OSAzLjY3MDQ3IDIuOTEwODUgNC42MDc1NCAyLjQ1NjczIDUuNzAzOUMyLjAwMjYgNi44MDAyNSAxLjg4Mzc4IDguMDA2NjUgMi4xMTUyOSA5LjE3MDU0QzIuMzQ2OCAxMC4zMzQ0IDIuOTE4MjUgMTEuNDAzNSAzLjc1NzM2IDEyLjI0MjZDNC41OTY0OCAxMy4wODE4IDUuNjY1NTggMTMuNjUzMiA2LjgyOTQ2IDEzLjg4NDdDNy45OTMzNSAxNC4xMTYyIDkuMTk5NzUgMTMuOTk3NCAxMC4yOTYxIDEzLjU0MzNDMTEuMzkyNSAxMy4wODkxIDEyLjMyOTUgMTIuMzIwMSAxMi45ODg4IDExLjMzMzRDMTMuNjQ4MSAxMC4zNDY3IDE0IDkuMTg2NjggMTQgOEw4IDhMOCAyWiIgZmlsbD0id2hpdGUiLz48L3N2Zz4=';
	var iCircleWhite = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxjaXJjbGUgY3g9IjgiIGN5PSI4IiByPSI2IiB0cmFuc2Zvcm09InJvdGF0ZSgtOTAgOCA4KSIgZmlsbD0id2hpdGUiLz48L3N2Zz4=';
	var iCircleYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgOEMxNSAxMS44NjYgMTEuODY2IDE1IDggMTVDNC4xMzQwMSAxNSAxIDExLjg2NiAxIDhDMSA0LjEzNDAxIDQuMTM0MDEgMSA4IDFDMTEuODY2IDEgMTUgNC4xMzQwMSAxNSA4WiIgZmlsbD0iI0ZGQ0YzMyIvPjwvc3ZnPg==';
	var iCrossRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTQuOTk5OSA3Ljk5OTk2QzE0Ljk5OTkgMTEuODY1OSAxMS44NjU5IDE0Ljk5OTkgNy45OTk5NiAxNC45OTk5QzQuMTMzOTkgMTQuOTk5OSAxIDExLjg2NTkgMSA3Ljk5OTk2QzEgNC4xMzM5OSA0LjEzMzk5IDEgNy45OTk5NiAxQzExLjg2NTkgMSAxNC45OTk5IDQuMTMzOTkgMTQuOTk5OSA3Ljk5OTk2WiIgZmlsbD0iI0ZGMTExMSIvPjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNi41ODU4IDhMNC4yOTI5MSA1LjcwNzExTDUuNzA3MTIgNC4yOTI4OUw4LjAwMDAxIDYuNTg1NzlMMTAuMjkyOSA0LjI5Mjg5TDExLjcwNzEgNS43MDcxMUw5LjQxNDIzIDhMMTEuNzA3MSAxMC4yOTI5TDEwLjI5MjkgMTEuNzA3MUw4LjAwMDAxIDkuNDE0MjFMNS43MDcxMiAxMS43MDcxTDQuMjkyOTEgMTAuMjkyOUw2LjU4NTggOFoiIGZpbGw9IndoaXRlIi8+PC9zdmc+';
	var iCrossSymbolRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTYuMTk3MzUgOEwyIDMuODAyNjVMMy44MDI2NSAyTDggNi4xOTczNUwxMi4xOTczIDJMMTQgMy44MDI2NUw5LjgwMjY1IDhMMTQgMTIuMTk3M0wxMi4xOTczIDE0TDggOS44MDI2NUwzLjgwMjY1IDE0TDIgMTIuMTk3M0w2LjE5NzM1IDhaIiBmaWxsPSIjRkYxMTExIi8+PC9zdmc+';
	var iDashYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSA1SDE1VjExSDFWNVoiIGZpbGw9IiNGRkNGMzMiLz48L3N2Zz4=';
	var iDiamondRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSI4IiB5PSIxIiB3aWR0aD0iOS44OTk1IiBoZWlnaHQ9IjkuODk5NSIgcng9IjIiIHRyYW5zZm9ybT0icm90YXRlKDQ1IDggMSkiIGZpbGw9IiNGRjExMTEiLz48L3N2Zz4=';
	var iDown = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTcuNSAxNUwwLjUgOEg1VjFIMTBWOEgxNC41TDcuNSAxNVoiIGZpbGw9IiNGRjExMTEiLz48L3N2Zz4=';
	var iDownGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTcuNSAxNUwwLjUgOEg1VjFIMTBWOEgxNC41TDcuNSAxNVoiIGZpbGw9IiM1MDUwNTAiLz48L3N2Zz4=';
	var iDownIncline = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTEzIDE1TDMuMTAwNTMgMTVMNi4yODI1MSAxMS44MThMMS4zMzI3NiA2Ljg2ODI3TDQuODY4MyAzLjMzMjczTDkuODE4MDUgOC4yODI0OEwxMyA1LjEwMDVMMTMgMTVaIiBmaWxsPSIjRkZDRjMzIi8+PC9zdmc+';
	var iDownInclineGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48ZyBjbGlwLXBhdGg9InVybCgjY2xpcDApIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTE0IDE0TDQuMTAwNTEgMTRMNy4yODI0OSAxMC44MThMMi4zMzI3NCA1Ljg2ODI3TDUuODY4MjcgMi4zMzI3NEwxMC44MTggNy4yODI0OUwxNCA0LjEwMDUxTDE0IDE0WiIgZmlsbD0iIzUwNTA1MCIvPjwvZz48ZGVmcz48Y2xpcFBhdGggaWQ9ImNsaXAwIj48cmVjdCB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIGZpbGw9IndoaXRlIi8+PC9jbGlwUGF0aD48L2RlZnM+PC9zdmc+';
	var iExclamationSymbolYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNy40MDcyOCA4LjkxNTQ2TDcuMDg5NCA0LjQ4NjEzQzcuMDI5OCAzLjYyMzA3IDcgMy4wMDM1MiA3IDIuNjI3NDhDNyAyLjExNTgxIDcuMTQyMzggMS43MTgxOSA3LjQyNzE1IDEuNDM0NjFDNy43MTg1NCAxLjE0NDg3IDguMDk5MzQgMSA4LjU2OTU0IDFDOS4xMzkwNyAxIDkuNTE5ODcgMS4xODQ5NCA5LjcxMTkyIDEuNTU0ODJDOS45MDM5NyAxLjkxODU0IDEwIDIuNDQ1NjIgMTAgMy4xMzYwNkMxMCAzLjU0MjkzIDkuOTc2ODIgMy45NTU5NyA5LjkzMDQ2IDQuMzc1MTZMOS41MDMzMSA4LjkzMzk1QzkuNDU2OTUgOS40NzY0NCA5LjM1NzYyIDkuODkyNTYgOS4yMDUzIDEwLjE4MjNDOS4wNTI5OCAxMC40NzIgOC44MDEzMiAxMC42MTY5IDguNDUwMzMgMTAuNjE2OUM4LjA5MjcyIDEwLjYxNjkgNy44NDQzNyAxMC40NzgyIDcuNzA1MyAxMC4yMDA4QzcuNTY2MjMgOS45MTcyMiA3LjQ2Njg5IDkuNDg4NzcgNy40MDcyOCA4LjkxNTQ2Wk04LjUwOTkzIDE1QzguMTA1OTYgMTUgNy43NTE2NiAxNC44Nzk4IDcuNDQ3MDIgMTQuNjM5NEM3LjE0OTAxIDE0LjM5MjggNyAxNC4wNTA2IDcgMTMuNjEyOUM3IDEzLjIzMDcgNy4xNDIzOCAxMi45MDcxIDcuNDI3MTUgMTIuNjQyQzcuNzE4NTQgMTIuMzcwOCA4LjA3Mjg1IDEyLjIzNTEgOC40OTAwNyAxMi4yMzUxQzguOTA3MjggMTIuMjM1MSA5LjI2MTU5IDEyLjM3MDggOS41NTI5OCAxMi42NDJDOS44NTA5OSAxMi45MDcxIDEwIDEzLjIzMDcgMTAgMTMuNjEyOUMxMCAxNC4wNDQ1IDkuODUwOTkgMTQuMzgzNSA5LjU1Mjk4IDE0LjYzMDFDOS4yNTQ5NyAxNC44NzY3IDguOTA3MjggMTUgOC41MDk5MyAxNVoiIGZpbGw9IiNGRkNGMzMiLz48L3N2Zz4=';
	var iExclamationYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgOEMxNSAxMS44NjYgMTEuODY2IDE1IDggMTVDNC4xMzQwMSAxNSAxIDExLjg2NiAxIDhDMSA0LjEzNDAxIDQuMTM0MDEgMSA4IDFDMTEuODY2IDEgMTUgNC4xMzQwMSAxNSA4WiIgZmlsbD0iI0ZGQ0YzMyIvPjxwYXRoIGQ9Ik03LjI3MTUyIDguNjUzOUw3LjA1OTYgNS40OTAwOUM3LjAxOTg3IDQuODczNjIgNyA0LjQzMTA5IDcgNC4xNjI0OEM3IDMuNzk3MDEgNy4wOTQ5MiAzLjUxMjk5IDcuMjg0NzcgMy4zMTA0NEM3LjQ3OTAzIDMuMTAzNDggNy43MzI4OSAzIDguMDQ2MzYgM0M4LjQyNjA1IDMgOC42Nzk5MSAzLjEzMjEgOC44MDc5NSAzLjM5NjNDOC45MzU5OCAzLjY1NjEgOSA0LjAzMjU4IDkgNC41MjU3NkM5IDQuODE2MzggOC45ODQ1NSA1LjExMTQgOC45NTM2NCA1LjQxMDgzTDguNjY4ODcgOC42NjcxMUM4LjYzNzk3IDkuMDU0NiA4LjU3MTc0IDkuMzUxODMgOC40NzAyIDkuNTU4NzhDOC4zNjg2NSA5Ljc2NTc0IDguMjAwODggOS44NjkyMiA3Ljk2Njg5IDkuODY5MjJDNy43Mjg0OCA5Ljg2OTIyIDcuNTYyOTEgOS43NzAxNSA3LjQ3MDIgOS41NzE5OUM3LjM3NzQ4IDkuMzY5NDQgNy4zMTEyNiA5LjA2MzQxIDcuMjcxNTIgOC42NTM5Wk04LjAwNjYyIDEzQzcuNzM3MzEgMTMgNy41MDExIDEyLjkxNDEgNy4yOTgwMSAxMi43NDI0QzcuMDk5MzQgMTIuNTY2MyA3IDEyLjMyMTkgNyAxMi4wMDkyQzcgMTEuNzM2MiA3LjA5NDkyIDExLjUwNTEgNy4yODQ3NyAxMS4zMTU3QzcuNDc5MDMgMTEuMTIyIDcuNzE1MjMgMTEuMDI1MSA3Ljk5MzM4IDExLjAyNTFDOC4yNzE1MiAxMS4wMjUxIDguNTA3NzMgMTEuMTIyIDguNzAxOTkgMTEuMzE1N0M4LjkwMDY2IDExLjUwNTEgOSAxMS43MzYyIDkgMTIuMDA5MkM5IDEyLjMxNzUgOC45MDA2NiAxMi41NTk3IDguNzAxOTkgMTIuNzM1OEM4LjUwMzMxIDEyLjkxMTkgOC4yNzE1MiAxMyA4LjAwNjYyIDEzWiIgZmlsbD0id2hpdGUiLz48L3N2Zz4=';
	var iFlagGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNSAxTDUgMTBMMTQgNS41TDUgMVoiIGZpbGw9IiMyRTk5NUYiLz48cmVjdCB4PSIyIiB5PSIwLjk5OTk5NiIgd2lkdGg9IjIiIGhlaWdodD0iMTQiIGZpbGw9IiM3MjcyNzIiLz48L3N2Zz4=';
	var iFlagRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNSAxTDUgMTBMMTQgNS41TDUgMVoiIGZpbGw9IiNGRjExMTEiLz48cmVjdCB4PSIyIiB5PSIwLjk5OTk5NiIgd2lkdGg9IjIiIGhlaWdodD0iMTQiIGZpbGw9IiM3MjcyNzIiLz48L3N2Zz4=';
	var iFlagYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNSAxTDUgMTBMMTQgNS41TDUgMVoiIGZpbGw9IiNGRkNGMzMiLz48cmVjdCB4PSIyIiB5PSIwLjk5OTk5NiIgd2lkdGg9IjIiIGhlaWdodD0iMTQiIGZpbGw9IiM3MjcyNzIiLz48L3N2Zz4=';
	var iFourFilledBars = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIxMCIgd2lkdGg9IjMiIGhlaWdodD0iNSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSI5IiB5PSI0IiB3aWR0aD0iMyIgaGVpZ2h0PSIxMSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iIzIzNjFCRSIvPjwvc3ZnPg==';
	var iFourFilledBoxes = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAySDdWOEgxVjJaIiBmaWxsPSIjMjM2MUJFIi8+PHJlY3QgeD0iOCIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+PC9zdmc+';
	var iOneFilledBars = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIxMCIgd2lkdGg9IjMiIGhlaWdodD0iNSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiNDQ0NDQ0MiLz48cmVjdCB4PSI5IiB5PSI0IiB3aWR0aD0iMyIgaGVpZ2h0PSIxMSIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPjwvc3ZnPg==';
	var iOneFilledBoxes = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+PHJlY3QgeD0iOCIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiNDQ0NDQ0MiLz48cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+PC9zdmc+';
	var iSide = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTE1IDguNUw4IDE1LjVMOCAxMUwxIDExTDEgNkw4IDZMOCAxLjVMMTUgOC41WiIgZmlsbD0iI0ZGQ0YzMyIvPjwvc3ZnPg==';
	var iSideGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTE1IDguNUw4IDE1LjVMOCAxMUwxIDExTDEgNkw4IDZMOCAxLjVMMTUgOC41WiIgZmlsbD0iIzUwNTA1MCIvPjwvc3ZnPg==';
	var iStarGold = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNy41IDIuMTI5NzhMOS4xMzk0OSA1LjQ1MTc1QzkuMjg1MTUgNS43NDY4OSA5LjU2NjcyIDUuOTUxNDYgOS44OTI0MyA1Ljk5ODc5TDEzLjU1ODQgNi41MzE0OUwxMC45MDU3IDkuMTE3MjlDMTAuNjcgOS4zNDcwMiAxMC41NjI1IDkuNjc4MDIgMTAuNjE4MSAxMC4wMDI0TDExLjI0NDMgMTMuNjUzNkw3Ljk2NTM0IDExLjkyOThDNy42NzQwMiAxMS43NzY2IDcuMzI1OTggMTEuNzc2NiA3LjAzNDY2IDExLjkyOThMMy43NTU2OCAxMy42NTM2TDQuMzgxOTEgMTAuMDAyNEM0LjQzNzU0IDkuNjc4MDMgNC4zMyA5LjM0NzAyIDQuMDk0MzEgOS4xMTcyOUwxLjQ0MTU2IDYuNTMxNDlMNS4xMDc1NyA1Ljk5ODc5QzUuNDMzMjggNS45NTE0NiA1LjcxNDg1IDUuNzQ2ODkgNS44NjA1MSA1LjQ1MTc1TDcuNSAyLjEyOTc4WiIgZmlsbD0iI0ZGQ0YzMyIgc3Ryb2tlPSIjRERBMTA5Ii8+PC9zdmc+';
	var iStarHalf = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48bWFzayBpZD0ibWFzazAiIG1hc2stdHlwZT0iYWxwaGEiIG1hc2tVbml0cz0idXNlclNwYWNlT25Vc2UiIHg9IjEiIHk9IjEiIHdpZHRoPSI3IiBoZWlnaHQ9IjE1Ij48cmVjdCB3aWR0aD0iNyIgaGVpZ2h0PSIxNSIgdHJhbnNmb3JtPSJtYXRyaXgoLTEgMCAwIDEgOCAxKSIgZmlsbD0iI0M0QzRDNCIvPjwvbWFzaz48ZyBtYXNrPSJ1cmwoI21hc2swKSI+PHBhdGggZD0iTTggMi4xMjk3OEw2LjM2MDUxIDUuNDUxNzVDNi4yMTQ4NSA1Ljc0Njg5IDUuOTMzMjggNS45NTE0NiA1LjYwNzU3IDUuOTk4NzlMMS45NDE1NiA2LjUzMTQ5TDQuNTk0MzEgOS4xMTcyOUM0LjgzIDkuMzQ3MDIgNC45Mzc1NCA5LjY3ODAyIDQuODgxOTEgMTAuMDAyNEw0LjI1NTY4IDEzLjY1MzZMNy41MzQ2NiAxMS45Mjk4QzcuODI1OTggMTEuNzc2NiA4LjE3NDAyIDExLjc3NjYgOC40NjUzNCAxMS45Mjk4TDExLjc0NDMgMTMuNjUzNkwxMS4xMTgxIDEwLjAwMjRDMTEuMDYyNSA5LjY3ODAzIDExLjE3IDkuMzQ3MDIgMTEuNDA1NyA5LjExNzI5TDE0LjA1ODQgNi41MzE0OUwxMC4zOTI0IDUuOTk4NzlDMTAuMDY2NyA1Ljk1MTQ2IDkuNzg1MTUgNS43NDY4OSA5LjYzOTQ5IDUuNDUxNzVMOCAyLjEyOTc4WiIgZmlsbD0iI0ZGQ0YzMyIgc3Ryb2tlPSIjRERBMTA5Ii8+PC9nPjxtYXNrIGlkPSJtYXNrMSIgbWFzay10eXBlPSJhbHBoYSIgbWFza1VuaXRzPSJ1c2VyU3BhY2VPblVzZSIgeD0iOCIgeT0iMSIgd2lkdGg9IjciIGhlaWdodD0iMTUiPjxyZWN0IHdpZHRoPSI3IiBoZWlnaHQ9IjE1IiB0cmFuc2Zvcm09Im1hdHJpeCgtMSAwIDAgMSAxNSAxKSIgZmlsbD0iI0M0QzRDNCIvPjwvbWFzaz48ZyBtYXNrPSJ1cmwoI21hc2sxKSI+PHBhdGggZD0iTTggMi4xMjk3OEw2LjM2MDUxIDUuNDUxNzVDNi4yMTQ4NSA1Ljc0Njg5IDUuOTMzMjggNS45NTE0NiA1LjYwNzU3IDUuOTk4NzlMMS45NDE1NiA2LjUzMTQ5TDQuNTk0MzEgOS4xMTcyOUM0LjgzIDkuMzQ3MDIgNC45Mzc1NSA5LjY3ODAyIDQuODgxOTEgMTAuMDAyNEw0LjI1NTY4IDEzLjY1MzZMNy41MzQ2NiAxMS45Mjk4QzcuODI1OTggMTEuNzc2NiA4LjE3NDAyIDExLjc3NjYgOC40NjUzNCAxMS45Mjk4TDExLjc0NDMgMTMuNjUzNkwxMS4xMTgxIDEwLjAwMjRDMTEuMDYyNSA5LjY3ODAzIDExLjE3IDkuMzQ3MDIgMTEuNDA1NyA5LjExNzI5TDE0LjA1ODQgNi41MzE0OUwxMC4zOTI0IDUuOTk4NzlDMTAuMDY2NyA1Ljk1MTQ2IDkuNzg1MTUgNS43NDY4OSA5LjYzOTQ5IDUuNDUxNzVMOCAyLjEyOTc4WiIgZmlsbD0id2hpdGUiIHN0cm9rZT0iIzczNzM3MyIvPjwvZz48L3N2Zz4=';
	var iStarSilver = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNy41IDIuMTI5NzhMOS4xMzk0OSA1LjQ1MTc1QzkuMjg1MTUgNS43NDY4OSA5LjU2NjcyIDUuOTUxNDYgOS44OTI0MyA1Ljk5ODc5TDEzLjU1ODQgNi41MzE0OUwxMC45MDU3IDkuMTE3MjlDMTAuNjcgOS4zNDcwMiAxMC41NjI1IDkuNjc4MDIgMTAuNjE4MSAxMC4wMDI0TDExLjI0NDMgMTMuNjUzNkw3Ljk2NTM0IDExLjkyOThDNy42NzQwMiAxMS43NzY2IDcuMzI1OTggMTEuNzc2NiA3LjAzNDY2IDExLjkyOThMMy43NTU2OCAxMy42NTM2TDQuMzgxOTEgMTAuMDAyNEM0LjQzNzU0IDkuNjc4MDMgNC4zMyA5LjM0NzAyIDQuMDk0MzEgOS4xMTcyOUwxLjQ0MTU2IDYuNTMxNDlMNS4xMDc1NyA1Ljk5ODc5QzUuNDMzMjggNS45NTE0NiA1LjcxNDg1IDUuNzQ2ODkgNS44NjA1MSA1LjQ1MTc1TDcuNSAyLjEyOTc4WiIgZmlsbD0id2hpdGUiIHN0cm9rZT0iIzczNzM3MyIvPjwvc3ZnPg==';
	var iThreeFilledBars = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIxMCIgd2lkdGg9IjMiIGhlaWdodD0iNSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSI5IiB5PSI0IiB3aWR0aD0iMyIgaGVpZ2h0PSIxMSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPjwvc3ZnPg==';
	var iThreeFilledBoxes = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+PHJlY3QgeD0iOCIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+PC9zdmc+';
	var iTrafficLightGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAzQzEgMS44OTU0MyAxLjg5NTQzIDEgMyAxSDEzQzE0LjEwNDYgMSAxNSAxLjg5NTQzIDE1IDNWMTNDMTUgMTQuMTA0NiAxNC4xMDQ2IDE1IDEzIDE1SDNDMS44OTU0MyAxNSAxIDE0LjEwNDYgMSAxM1YzWiIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik0xNCA4QzE0IDQuNjg2MjkgMTEuMzEzNyAyIDggMkM0LjY4NjI5IDIgMiA0LjY4NjI5IDIgOEMyIDExLjMxMzcgNC42ODYyOSAxNCA4IDE0QzExLjMxMzcgMTQgMTQgMTEuMzEzNyAxNCA4WiIgZmlsbD0id2hpdGUiLz48cGF0aCBkPSJNMTMgOEMxMyA1LjIzODU4IDEwLjc2MTQgMyA4IDNDNS4yMzg1OCAzIDMgNS4yMzg1OCAzIDhDMyAxMC43NjE0IDUuMjM4NTggMTMgOCAxM0MxMC43NjE0IDEzIDEzIDEwLjc2MTQgMTMgOFoiIGZpbGw9IiMyRTk5NUYiLz48L3N2Zz4=';
	var iTrafficLightRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAzQzEgMS44OTU0MyAxLjg5NTQzIDEgMyAxSDEzQzE0LjEwNDYgMSAxNSAxLjg5NTQzIDE1IDNWMTNDMTUgMTQuMTA0NiAxNC4xMDQ2IDE1IDEzIDE1SDNDMS44OTU0MyAxNSAxIDE0LjEwNDYgMSAxM1YzWiIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik0xNCA4QzE0IDQuNjg2MjkgMTEuMzEzNyAyIDggMkM0LjY4NjI5IDIgMiA0LjY4NjI5IDIgOEMyIDExLjMxMzcgNC42ODYyOSAxNCA4IDE0QzExLjMxMzcgMTQgMTQgMTEuMzEzNyAxNCA4WiIgZmlsbD0id2hpdGUiLz48cGF0aCBkPSJNMTMgOEMxMyA1LjIzODU4IDEwLjc2MTQgMyA4IDNDNS4yMzg1OCAzIDMgNS4yMzg1OCAzIDhDMyAxMC43NjE0IDUuMjM4NTggMTMgOCAxM0MxMC43NjE0IDEzIDEzIDEwLjc2MTQgMTMgOFoiIGZpbGw9IiNGRjExMTEiLz48L3N2Zz4=';
	var iTrafficLightYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAzQzEgMS44OTU0MyAxLjg5NTQzIDEgMyAxSDEzQzE0LjEwNDYgMSAxNSAxLjg5NTQzIDE1IDNWMTNDMTUgMTQuMTA0NiAxNC4xMDQ2IDE1IDEzIDE1SDNDMS44OTU0MyAxNSAxIDE0LjEwNDYgMSAxM1YzWiIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik0xNCA4QzE0IDQuNjg2MjkgMTEuMzEzNyAyIDggMkM0LjY4NjI5IDIgMiA0LjY4NjI5IDIgOEMyIDExLjMxMzcgNC42ODYyOSAxNCA4IDE0QzExLjMxMzcgMTQgMTQgMTEuMzEzNyAxNCA4WiIgZmlsbD0id2hpdGUiLz48cGF0aCBkPSJNMTMgOEMxMyA1LjIzODU4IDEwLjc2MTQgMyA4IDNDNS4yMzg1OCAzIDMgNS4yMzg1OCAzIDhDMyAxMC43NjE0IDUuMjM4NTggMTMgOCAxM0MxMC43NjE0IDEzIDEzIDEwLjc2MTQgMTMgOFoiIGZpbGw9IiNGRkNGMzMiLz48L3N2Zz4=';
	var iTriangleYellow = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAxNUgxNUw4IDFMMSAxNVoiIGZpbGw9IiNGRkNGMzMiLz48L3N2Zz4=';
	var iTriangleGreen = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMSAxMkgxNUw4IDNMMSAxMloiIGZpbGw9IiMyRTk5NUYiLz48L3N2Zz4=';
	var iTriangleRed = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTUgNEwxIDRMOCAxM0wxNSA0WiIgZmlsbD0iI0ZGMTExMSIvPjwvc3ZnPg==';
	var iTwoFilledBars = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIxMCIgd2lkdGg9IjMiIGhlaWdodD0iNSIgZmlsbD0iIzIzNjFCRSIvPjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSI5IiB5PSI0IiB3aWR0aD0iMyIgaGVpZ2h0PSIxMSIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPjwvc3ZnPg==';
	var iTwoFilledBoxes = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+PHJlY3QgeD0iOCIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz48cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+PC9zdmc+';
	var iUp = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTcuNSAxTDE0LjUgOEwxMCA4VjE1TDUgMTVMNSA4TDAuNSA4TDcuNSAxWiIgZmlsbD0iIzJFOTk1RiIvPjwvc3ZnPg==';
	var iUpGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTcuNSAxTDE0LjUgOEwxMCA4VjE1TDUgMTVMNSA4TDAuNSA4TDcuNSAxWiIgZmlsbD0iIzUwNTA1MCIvPjwvc3ZnPg==';
	var iUpIncline = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTEzIDNMMy4xMDA1MyAzTDYuMjgyNTEgNi4xODE5OEwxLjMzMjc2IDExLjEzMTdMNC44NjgzIDE0LjY2NzNMOS44MTgwNCA5LjcxNzUxTDEzIDEyLjg5OTVWM1oiIGZpbGw9IiNGRkNGMzMiLz48L3N2Zz4=';
	var iUpInclineGray = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48ZyBjbGlwLXBhdGg9InVybCgjY2xpcDApIj48cGF0aCBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsaXAtcnVsZT0iZXZlbm9kZCIgZD0iTTE0IDJMNC4xMDA1MSAyTDcuMjgyNDkgNS4xODE5OEwyLjMzMjc0IDEwLjEzMTdMNS44NjgyNyAxMy42NjczTDEwLjgxOCA4LjcxNzUxTDE0IDExLjg5OTVMMTQgMloiIGZpbGw9IiM1MDUwNTAiLz48L2c+PGRlZnM+PGNsaXBQYXRoIGlkPSJjbGlwMCI+PHJlY3Qgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2IiBmaWxsPSJ3aGl0ZSIvPjwvY2xpcFBhdGg+PC9kZWZzPjwvc3ZnPg==';
	var iZeroFilledBars = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIxMCIgd2lkdGg9IjMiIGhlaWdodD0iNSIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiNDQ0NDQ0MiLz48cmVjdCB4PSI5IiB5PSI0IiB3aWR0aD0iMyIgaGVpZ2h0PSIxMSIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPjwvc3ZnPg==';
	var iZeroFilledBoxes = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB4PSIxIiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+PHJlY3QgeD0iOCIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiNDQ0NDQ0MiLz48cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+PC9zdmc+';

	var c_arrIcons = [20];
	c_arrIcons[EIconSetType.Arrows3] = [iDown, iSide, iUp];
	c_arrIcons[EIconSetType.Arrows3Gray] = [iDownGray, iSideGray, iUpGray];
	c_arrIcons[EIconSetType.Flags3] = [iFlagRed, iFlagYellow, iFlagGreen];
	c_arrIcons[EIconSetType.Signs3] = [iDiamondRed, iTriangleYellow, iCircleGreen];
	c_arrIcons[EIconSetType.Symbols3] = [iCrossRed, iExclamationYellow, iCheckGreen];
	c_arrIcons[EIconSetType.Symbols3_2] = [iCrossSymbolRed, iExclamationSymbolYellow, iCheckSymbolGreen];
	c_arrIcons[EIconSetType.Traffic3Lights1] = [iCircleRed, iCircleYellow, iCircleGreen];
	c_arrIcons[EIconSetType.Traffic3Lights2] = [iTrafficLightRed, iTrafficLightYellow, iTrafficLightGreen];
	c_arrIcons[EIconSetType.Arrows4] = [iDown, iDownIncline, iUpIncline, iUp];
	c_arrIcons[EIconSetType.Arrows4Gray] = [iDownGray, iDownInclineGray, iUpInclineGray, iUpGray];
	c_arrIcons[EIconSetType.Rating4] = [iOneFilledBars, iTwoFilledBars, iThreeFilledBars, iFourFilledBars];
	c_arrIcons[EIconSetType.RedToBlack4] = [iCircleBlack, iCircleGray, iCircleLightRed, iCircleRed];
	c_arrIcons[EIconSetType.Traffic4Lights] = [iCircleBlack, iCircleRed, iCircleYellow, iCircleGreen];
	c_arrIcons[EIconSetType.Arrows5] = [iDown, iDownIncline, iSide, iUpIncline, iUp];
	c_arrIcons[EIconSetType.Arrows5Gray] = [iDownGray, iDownInclineGray, iSideGray, iUpInclineGray, iUpGray];
	c_arrIcons[EIconSetType.Quarters5] = [iCircleWhite, iCircleThreeWhiteQuarters, iCircleTwoWhiteQuarters, iCircleOneWhiteQuarter, iCircleBlack];
	c_arrIcons[EIconSetType.Rating5] = [iZeroFilledBars, iOneFilledBars, iTwoFilledBars, iThreeFilledBars, iFourFilledBars];
	c_arrIcons[EIconSetType.Triangles3] = [iTriangleRed, iDashYellow, iTriangleGreen];
	c_arrIcons[EIconSetType.Stars3] = [iStarSilver, iStarHalf, iStarGold];
	c_arrIcons[EIconSetType.Boxes5] = [iZeroFilledBoxes, iOneFilledBoxes, iTwoFilledBoxes, iThreeFilledBoxes, iFourFilledBoxes];

	function getCFIconsForLoad() {
		return [iCheckGreen, iCheckSymbolGreen, iCircleTwoWhiteQuarters, iCircleBlack, iCircleGray, iCircleGreen,
			iCircleLightRed, iCircleOneWhiteQuarter, iCircleRed, iCircleThreeWhiteQuarters, iCircleWhite,
			iCircleYellow, iCrossRed, iCrossSymbolRed, iDashYellow, iDiamondRed, iDown, iDownGray, iDownIncline,
			iDownInclineGray, iExclamationSymbolYellow, iExclamationYellow, iFlagGreen, iFlagRed, iFlagYellow,
			iFourFilledBars, iFourFilledBoxes, iOneFilledBars, iOneFilledBoxes, iSide, iSideGray, iStarGold, iStarHalf,
			iStarSilver, iThreeFilledBars, iThreeFilledBoxes, iTrafficLightGreen, iTrafficLightRed,
			iTrafficLightYellow, iTriangleYellow, iTriangleGreen, iTriangleRed, iTwoFilledBars, iTwoFilledBoxes, iUp,
			iUpGray, iUpIncline, iUpInclineGray, iZeroFilledBars, iZeroFilledBoxes];
	}

	function getCFIcon(oRuleElement, index) {
		var oIconSet = oRuleElement.aIconSets[index];
		var iconSetType = (oIconSet && null !== oIconSet.IconSet) ? oIconSet.IconSet : oRuleElement.IconSet;
		if (EIconSetType.NoIcons === iconSetType) {
			return null;
		}
		var icons = c_arrIcons[iconSetType] || c_arrIcons[EIconSetType.Traffic3Lights1];
		return icons[(oIconSet && null !== oIconSet.IconId) ? oIconSet.IconId : index] || icons[icons.length - 1];
	}

	function getDataBarGradientColor(color) {
		if (!color) {
			return null;
		}
		var RGB = {R: 0xFF, G: 0xFF, B: 0xFF};
		var bCoeff = 0.828;
		RGB.R = Math.min(255, (color.getR() + bCoeff * (0xFF - color.getR()) + 0.5) >> 0);
		RGB.G = Math.min(255, (color.getG() + bCoeff * (0xFF - color.getG()) + 0.5) >> 0);
		RGB.B = Math.min(255, (color.getB() + bCoeff * (0xFF - color.getB()) + 0.5) >> 0);
		return AscCommonExcel.createRgbColor(RGB.R, RGB.G, RGB.B);
	}

	/*
	 * Export
	 * -----------------------------------------------------------------------------
	 */
	var prot;
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CConditionalFormatting = CConditionalFormatting;
	window['AscCommonExcel'].CConditionalFormattingFormulaParent = CConditionalFormattingFormulaParent;
	window['AscCommonExcel'].CConditionalFormattingRule = CConditionalFormattingRule;
	window['AscCommonExcel'].CColorScale = CColorScale;
	window['AscCommonExcel'].CDataBar = CDataBar;
	window['AscCommonExcel'].CFormulaCF = CFormulaCF;
	window['AscCommonExcel'].CIconSet = CIconSet;
	window['AscCommonExcel'].CConditionalFormatValueObject = CConditionalFormatValueObject;
	window['AscCommonExcel'].CConditionalFormatIconSet = CConditionalFormatIconSet;
	window['AscCommonExcel'].CGradient = CGradient;

	window['AscCommonExcel'].cDefIconSize = cDefIconSize;
	window['AscCommonExcel'].cDefIconFont = cDefIconFont;
	window['AscCommonExcel'].getCFIconsForLoad = getCFIconsForLoad;
	window['AscCommonExcel'].getCFIcon = getCFIcon;
	window['AscCommonExcel'].getDataBarGradientColor = getDataBarGradientColor;

	prot = CConditionalFormattingRule;
	prot['asc_getDxf'] = prot.asc_getDxf;
	prot['asc_getType'] = prot.asc_getType;
	prot['asc_getLocation'] = prot.asc_getLocation;
	prot['asc_getContainsText'] = prot.asc_getContainsText;
	prot['asc_getTimePeriod'] = prot.asc_getTimePeriod;
	prot['asc_getOperator'] = prot.asc_getOperator;
	prot['asc_getRank'] = prot.asc_getRank;
	prot['asc_getBottom'] = prot.asc_getBottom;
	prot['asc_getPercent'] = prot.asc_getPercent;
	prot['asc_getAboveAverage'] = prot.asc_getAboveAverage;
	prot['asc_getEqualAverage'] = prot.asc_getEqualAverage;
	prot['asc_getStdDev'] = prot.asc_getStdDev;
	prot['asc_getValue1'] = prot.asc_getValue1;
	prot['asc_getValue2'] = prot.asc_getValue2;
	prot['asc_getColorScaleOrDataBarOrIconSetRule'] = prot.asc_getColorScaleOrDataBarOrIconSetRule;
	prot['asc_getId'] = prot.asc_getId;
	prot['asc_getIsLock'] = prot.asc_getIsLock;
	prot['asc_getPreview'] = prot.asc_getPreview;
	prot['asc_setDxf'] = prot.asc_setDxf;
	prot['asc_setType'] = prot.asc_setType;
	prot['asc_setLocation'] = prot.asc_setLocation;
	prot['asc_setContainsText'] = prot.asc_setContainsText;
	prot['asc_setTimePeriod'] = prot.asc_setTimePeriod;
	prot['asc_setOperator'] = prot.asc_setOperator;
	prot['asc_setRank'] = prot.asc_setRank;
	prot['asc_setBottom'] = prot.asc_setBottom;
	prot['asc_setPercent'] = prot.asc_setPercent;
	prot['asc_setAboveAverage'] = prot.asc_setAboveAverage;
	prot['asc_setEqualAverage'] = prot.asc_setEqualAverage;
	prot['asc_setStdDev'] = prot.asc_setStdDev;
	prot['asc_setValue1'] = prot.asc_setValue1;
	prot['asc_setValue2'] = prot.asc_setValue2;
	prot['asc_checkScope'] = prot.asc_checkScope;
	//prot['asc_setColorScaleOrDataBarOrIconSetRule'] = prot.asc_setColorScaleOrDataBarOrIconSetRule;


	prot = CColorScale;
	prot['asc_getCFVOs'] = prot.asc_getCFVOs;
	prot['asc_getColors'] = prot.asc_getColors;
	prot['asc_getPreview'] = prot.asc_getPreview;
	prot['asc_setCFVOs'] = prot.asc_setCFVOs;
	prot['asc_setColors'] = prot.asc_setColors;

	prot = CDataBar;
	prot['asc_getShowValue'] = prot.asc_getShowValue;
	prot['asc_getAxisPosition'] = prot.asc_getAxisPosition;
	prot['asc_getGradient'] = prot.asc_getGradient;
	prot['asc_getDirection'] = prot.asc_getDirection;
	prot['asc_getNegativeBarColorSameAsPositive'] = prot.asc_getNegativeBarColorSameAsPositive;
	prot['asc_getNegativeBarBorderColorSameAsPositive'] = prot.asc_getNegativeBarBorderColorSameAsPositive;
	prot['asc_getCFVOs'] = prot.asc_getCFVOs;
	prot['asc_getColor'] = prot.asc_getColor;
	prot['asc_getNegativeColor'] = prot.asc_getNegativeColor;
	prot['asc_getBorderColor'] = prot.asc_getBorderColor;
	prot['asc_getNegativeBorderColor'] = prot.asc_getNegativeBorderColor;
	prot['asc_getAxisColor'] = prot.asc_getAxisColor;
	prot['asc_setShowValue'] = prot.asc_setShowValue;
	prot['asc_setAxisPosition'] = prot.asc_setAxisPosition;
	prot['asc_setGradient'] = prot.asc_setGradient;
	prot['asc_setDirection'] = prot.asc_setDirection;
	prot['asc_setNegativeBarColorSameAsPositive'] = prot.asc_setNegativeBarColorSameAsPositive;
	prot['asc_setNegativeBarBorderColorSameAsPositive'] = prot.asc_setNegativeBarBorderColorSameAsPositive;
	prot['asc_setCFVOs'] = prot.asc_setCFVOs;
	prot['asc_setColor'] = prot.asc_setColor;
	prot['asc_setNegativeColor'] = prot.asc_setNegativeColor;
	prot['asc_setBorderColor'] = prot.asc_setBorderColor;
	prot['asc_setNegativeBorderColor'] = prot.asc_setNegativeBorderColor;
	prot['asc_setAxisColor'] = prot.asc_setAxisColor;

	prot = CIconSet;
	prot['asc_getIconSet'] = prot.asc_getIconSet;
	prot['asc_getReverse'] = prot.asc_getReverse;
	prot['asc_getShowValue'] = prot.asc_getShowValue;
	prot['asc_getCFVOs'] = prot.asc_getCFVOs;
	prot['asc_getIconSets'] = prot.asc_getIconSets;
	prot['asc_setIconSet'] = prot.asc_setIconSet;
	prot['asc_setReverse'] = prot.asc_setReverse;
	prot['asc_setShowValue'] = prot.asc_setShowValue;
	prot['asc_setCFVOs'] = prot.asc_setCFVOs;
	prot['asc_setIconSets'] = prot.asc_setIconSets;

	prot = CConditionalFormatValueObject;
	prot['asc_getGte'] = prot.asc_getGte;
	prot['asc_getType'] = prot.asc_getType;
	prot['asc_getVal'] = prot.asc_getVal;
	prot['asc_setGte'] = prot.asc_setGte;
	prot['asc_setType'] = prot.asc_setType;
	prot['asc_setVal'] = prot.asc_setVal;
})(window);
