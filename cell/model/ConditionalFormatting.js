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
	var asc_error = Asc.c_oAscError.ID;

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
				elem.Read_FromBinary2(reader);
				this.aRuleElements.push(elem);
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
					ws.getId(), this.getUnionRange(), new AscCommonExcel.UndoRedoData_CF(this.id, this.aRuleElements, val.aRuleElements));
			}

			this.aRuleElements = val.aRuleElements;
		}

		if ((this.dxf && val.dxf && !this.dxf.isEqual(val.dxf)) || (this.dxf && !val.dxf) || (!this.dxf && val.dxf)) {
			var elem = val.dxf ? val.dxf.clone() : null;
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoCF, AscCH.historyitem_CFRule_SetDxf,
					ws.getId(), this.getUnionRange(), new AscCommonExcel.UndoRedoData_CF(this.id, this.dxf, elem));
			}

			this.dxf = elem;
		}

		if (this.ranges && val.ranges && !compareElements(this.ranges, val.ranges)) {
			this.setLocation(val.ranges, ws, true);
		}
	};

	CConditionalFormattingRule.prototype.setLocation = function (location, ws, addToHistory) {
		if (addToHistory) {
			var getUndoRedoRange = function (_ranges) {
				var needRanges = [];
				for (var i = 0; i < _ranges.length; i++) {
					needRanges.push(new AscCommonExcel.UndoRedoData_BBox(_ranges[i]));
				}
				return needRanges;
			};

			History.Add(AscCommonExcel.g_oUndoRedoCF, AscCH.historyitem_CFRule_SetRanges,
				ws.getId(), this.getUnionRange(location), new AscCommonExcel.UndoRedoData_CF(this.id, getUndoRedoRange(this.ranges), getUndoRedoRange(location)));
		}
		this.ranges = location;
		if (ws) {
			ws.cleanConditionalFormattingRangeIterator();
		}
	};

	CConditionalFormattingRule.prototype.checkProperty = function (propOld, propNew, type, ws, addToHistory) {
		if (propOld !== propNew) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoCF, type, ws.getId(), this.getUnionRange(),
					new AscCommonExcel.UndoRedoData_CF(this.id, propOld, propNew));
			}
			return propNew;
		}
		return propOld;
	};

	CConditionalFormattingRule.prototype.setOffset = function(offset, range, ws, addToHistory) {
		var newRanges = [];
		var isChange = false;

		var _setDiff = function (_range) {
			//TODO объединть в одну функцию с dataValidation(.shift)
			var _newRanges, _offset, tempRange, intersection, otherPart, diff;

			if (range && range.getType() === Asc.c_oAscSelectionType.RangeCells) {
				if (offset.row !== 0) {
					//c_oAscInsertOptions.InsertCellsAndShiftDown
					tempRange = new Asc.Range(range.c1, range.r1, range.c2, AscCommon.gc_nMaxRow0);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = range.r2 - range.r1 + 1;

						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						_offset = new AscCommon.CellBase(offset.row > 0 ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(_offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);

					}
				} else if (offset.col !== 0) {
					//c_oAscInsertOptions.InsertCellsAndShiftRight
					tempRange = new Asc.Range(range.c1, range.r1, AscCommon.gc_nMaxCol0, range.r2);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = range.c2 - range.c1 + 1;
						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						_offset = new AscCommon.CellBase(0, offset.col > 0 ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(_offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);
					}
				}
			}

			return _newRanges;
		};

		for (var i = 0; i < this.ranges.length; i++) {
			var newRange = this.ranges[i].clone();
			if (range.isIntersectForShift(newRange, offset)) {
				if (newRange.forShift(range, offset)) {
					if (ws.autoFilters.isAddTotalRow && newRange.containsRange(this.ranges[i])) {
						newRange = this.ranges[i].clone();
					} else {
						isChange = true;
					}
				}
				newRanges.push(newRange);
			} else {
				if (ws.autoFilters.isAddTotalRow && newRange.containsRange(this.ranges[i])) {
					newRange = this.ranges[i].clone();
				} else {
					var changedRanges = _setDiff(this.ranges[i]);
					if (changedRanges) {
						newRanges = newRanges.concat(changedRanges);
						isChange = true;
					} else {
						newRanges = newRanges.concat(this.ranges[i].clone());
					}
				}
			}
		}
		if (isChange) {
			this.setLocation(newRanges, ws, addToHistory);
			if (ws) {
				ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(newRanges))
			}
		}
	};

	CConditionalFormattingRule.prototype.getUnionRange = function (opt_ranges) {
		var res = null;

		var _getUnionRanges = function (_ranges) {
			if (!_ranges) {
				return null;
			}

			var _res = null;
			for (var i = 0; i < _ranges.length; i++) {
				if (!_res) {
					_res = _ranges[i].clone();
				} else {
					_res.union2(_ranges[i]);
				}
			}
			return _res;
		};

		if (opt_ranges) {
			res = _getUnionRanges(opt_ranges);
		}
		if (this.ranges) {
			var tempRange = _getUnionRanges(this.ranges);
			if (tempRange) {
				if (res) {
					res.union2(tempRange)
				} else {
					res = tempRange;
				}
			}
		}

		return res;
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
		if (this.stdDev && stdDev) {
			average += ((this.aboveAverage ? 1 : -1) * this.stdDev) * stdDev;
		}
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
	CConditionalFormattingRule.prototype.containsIntoRange = function (range) {
		var res = null;
		if (this.ranges && this.ranges.length > 0) {
			res = true;
			for (var i = 0; i < this.ranges.length; ++i) {
				if (!range.containsRange(this.ranges[i])) {
					res = false;
					break;
				}
			}
		}
		return res;
	};
	CConditionalFormattingRule.prototype.containsIntoRange = function (range) {
		var res = null;
		if (this.ranges && this.ranges.length > 0) {
			res = true;
			for (var i = 0; i < this.ranges.length; ++i) {
				if (!range.containsRange(this.ranges[i])) {
					res = false;
					break;
				}
			}
		}
		return res;
	};
	CConditionalFormattingRule.prototype.getIntersections = function (range) {
		var res = [];
		if (this.ranges) {
			for (var i = 0; i < this.ranges.length; ++i) {
				var intersection = this.ranges[i].intersection(range);
				if (intersection) {
					res.push(intersection);
				}
			}
		}
		return res.length ? res : null;
	};
	CConditionalFormattingRule.prototype.getIndexRule = function (values, ws, value) {
		var valueCFVO;
		var aCFVOs = this._getCFVOs();
		var bReverse = this.aRuleElements && this.aRuleElements[0] && this.aRuleElements[0].Reverse;
		for (var i = aCFVOs.length - 1; i >= 0; --i) {
			valueCFVO = this._getValue(values, aCFVOs[i], ws);
			if (value > valueCFVO || (aCFVOs[i].Gte && value === valueCFVO)) {
				return bReverse ? aCFVOs.length - 1 - i : i;
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

	CConditionalFormattingRule.prototype.applyPreset = function (presetId) {
		var presetType = presetId[0];
		var styleIndex = presetId[1];

		var elem;
		switch (presetType) {
			case Asc.c_oAscCFRuleTypeSettings.dataBar:
				elem = new CDataBar();
				elem.applyPreset(styleIndex);
				this.type = Asc.ECfType.dataBar;
				break;
			case Asc.c_oAscCFRuleTypeSettings.colorScale:
				elem = new CColorScale();
				elem.applyPreset(styleIndex);
				this.type = Asc.ECfType.colorScale;
				break;
			case Asc.c_oAscCFRuleTypeSettings.icons:
				elem = new CIconSet();
				elem.applyPreset(styleIndex);
				this.type = Asc.ECfType.iconSet;
				break;
		}
		if (elem) {
			this.aRuleElements.push(elem);
		}
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
		return [isActive, "=" + arrResult.join(AscCommon.FormulaSeparators.functionArgumentSeparator)];
	};
	CConditionalFormattingRule.prototype.asc_getContainsText = function () {
		if (null !== this.text) {
			return this.text;
		}
		var ruleElement = this.aRuleElements[1];
		return ruleElement && ruleElement.getFormula ? "=" + ruleElement.Text : null;
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
		return ruleElement && ruleElement.getFormula ? "=" + ruleElement.Text : null;
	};
	CConditionalFormattingRule.prototype.asc_getValue2 = function () {
		var ruleElement = this.aRuleElements[1];
		return ruleElement && ruleElement.getFormula ? "=" + ruleElement.Text : null;
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
		if (Asc.ECfType.colorScale === this.type && 1 === this.aRuleElements.length) {
			res = this.aRuleElements[0].asc_getPreview(api_sheet, id, text);
		} else if (Asc.ECfType.dataBar === this.type && 1 === this.aRuleElements.length) {
			res = this.aRuleElements[0].asc_getPreview(api_sheet, id);
		} else if (Asc.ECfType.iconSet === this.type && 1 === this.aRuleElements.length) {
			res = this.aRuleElements[0].asc_getPreview(api_sheet, id);
		} else {
			if (this.dxf) {
				res = this.dxf.asc_getPreview2(api_sheet, id, text);
			} else {
				var tempXfs = new AscCommonExcel.CellXfs();
				res = tempXfs.asc_getPreview2(api_sheet, id, text);
			}
		}
		return res;
	};

	CConditionalFormattingRule.prototype.asc_setType = function (val) {
		this.type = val;
		this._cleanAfterChangeType();
		var formula = this.getFormulaByType(this.text);
		if (formula) {
			this.aRuleElements = [];
			this.aRuleElements[0] = new CFormulaCF();
			this.aRuleElements[0].Text = formula;
		}
	};
	CConditionalFormattingRule.prototype._cleanAfterChangeType = function () {
		switch (this.type) {
			case Asc.ECfType.notContainsErrors:
			case Asc.ECfType.containsErrors:
			case Asc.ECfType.notContainsBlanks:
			case Asc.ECfType.containsBlanks:
			case Asc.ECfType.timePeriod:
				this.aRuleElements = [];
				this.percent = null;
				this.text = null;
				this.rank = null;
				break;
			case Asc.ECfType.colorScale:
				this.dxf = null;
				break;
		}
		if (this.type !== Asc.ECfType.top10) {
			this.rank = null;
		}
	};
	CConditionalFormattingRule.prototype.asc_setDxf = function (val) {
		this.dxf = val;
	};
	CConditionalFormattingRule.prototype.asc_setLocation = function (val) {
		var t = this;
		if (val) {
			if (val[0] === "=") {
				val = val.slice(1);
			}
			val = val.split(",");
			this.ranges = [];
			val.forEach(function (item) {
				if (-1 !== item.indexOf("!")) {
					var is3DRef = AscCommon.parserHelp.parse3DRef(item);
					if (is3DRef) {
						item = is3DRef.range;
					}
				}
				t.ranges.push(AscCommonExcel.g_oRangeCache.getAscRange(item));
			});
		}
	};
	
	CConditionalFormattingRule.prototype.asc_setContainsText = function (val) {
		if (val[0] === "=") {
			val = val.slice(1);
			//генерируем массив
			this.aRuleElements = [];
			this.aRuleElements[0] = new CFormulaCF();
			this.aRuleElements[0].Text = this.getFormulaByType(val);
			this.aRuleElements[1] = new CFormulaCF();
			this.aRuleElements[1].Text = val;
			this.text = null;
		} else {
			this.aRuleElements = [];
			this.aRuleElements[0] = new CFormulaCF();
			this.aRuleElements[0].Text = this.getFormulaByType(val);
			this.text = val;
		}
	};

	CConditionalFormattingRule.prototype.getFormulaByType = function (val) {
		var t = this;
		var _generateTimePeriodFunction = function () {
			switch (t.timePeriod) {
				case AscCommonExcel.ST_TimePeriod.yesterday:
					res = "FLOOR(" + firstCellInRange + ",1)" + "=TODAY()-1";
					break;
				case AscCommonExcel.ST_TimePeriod.today:
					res = "FLOOR(" + firstCellInRange + ",1)" + "=TODAY()";
					break;
				case AscCommonExcel.ST_TimePeriod.tomorrow:
					res = "FLOOR(" + firstCellInRange + ",1)" + "=TODAY()+1";
					break;
				case AscCommonExcel.ST_TimePeriod.last7Days:
					res = "AND(TODAY()-FLOOR(" + firstCellInRange + ",1)<=6,FLOOR(" + firstCellInRange + ",1)<=TODAY())";
					break;
				case AscCommonExcel.ST_TimePeriod.lastWeek:
					res = "AND(TODAY()-ROUNDDOWN(" + firstCellInRange + ",0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(" + firstCellInRange + ",0)<(WEEKDAY(TODAY())+7))";
					break;
				case AscCommonExcel.ST_TimePeriod.thisWeek:
					res = "AND(TODAY()-ROUNDDOWN(" + firstCellInRange + ",0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(" + firstCellInRange + ",0)-TODAY()<=7-WEEKDAY(TODAY()))";
					break;
				case AscCommonExcel.ST_TimePeriod.nextWeek:
					res = "AND(ROUNDDOWN(" + firstCellInRange + ",0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(" + firstCellInRange + ",0)-TODAY()<(15-WEEKDAY(TODAY())))";
					break;
				case AscCommonExcel.ST_TimePeriod.lastMonth:
					res = "AND(MONTH(" +firstCellInRange + ")=MONTH(EDATE(TODAY(),0-1)),YEAR(" + firstCellInRange + ")=YEAR(EDATE(TODAY(),0-1)))";
					break;
				case AscCommonExcel.ST_TimePeriod.thisMonth:
					res = "AND(MONTH(" + firstCellInRange + ")=MONTH(TODAY()),YEAR(" + firstCellInRange + ")=YEAR(TODAY()))";
					break;
				case AscCommonExcel.ST_TimePeriod.nextMonth:
					res = "AND(MONTH(" + firstCellInRange + ")=MONTH(EDATE(TODAY(),0+1)),YEAR(" + firstCellInRange + ")=YEAR(EDATE(TODAY(),0+1)))";
					break;
			}
		};

		var res = null;
		var range;
		if (val !== null && val !== undefined) {
			val = addQuotes(val);
		}
		if (this.ranges && this.ranges[0]) {
			range = this.ranges[0];
		} else {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var sheet = wb.getActiveWs();
			range = sheet && sheet.selectionRange && sheet.selectionRange.ranges && sheet.selectionRange.ranges[0];
		}

		if (range) {
			var firstCellInRange = new Asc.Range(range.c1, range.r1, range.c1, range.r1);

			AscCommonExcel.executeInR1C1Mode(false, function () {
				firstCellInRange = firstCellInRange.getName();
			});

			switch (this.type) {
				case Asc.ECfType.notContainsText:
					if (val !== null && val !== undefined) {
						res = "ISERROR(SEARCH(" + val + "," + firstCellInRange + "))";
					}
					break;
				case Asc.ECfType.containsText:
					if (val !== null && val !== undefined) {
						res = "NOT(ISERROR(SEARCH(" + val + "," + firstCellInRange + ")))";
					}
					break;
				case Asc.ECfType.endsWith:
					if (val !== null && val !== undefined) {
						res = "RIGHT(" + firstCellInRange + ",LEN(" + val + "))" + "=" + val;
					}
					break;
				case Asc.ECfType.beginsWith:
					if (val !== null && val !== undefined) {
						res = "LEFT(" + firstCellInRange + ",LEN(" + val + "))" + "=" + val;
					}
					break;
				case Asc.ECfType.notContainsErrors:
					res = "NOT(ISERROR(" + firstCellInRange + "))";
					break;
				case Asc.ECfType.containsErrors:
					res = "ISERROR(" + firstCellInRange + ")";
					break;
				case Asc.ECfType.notContainsBlanks:
					res = "LEN(TRIM(" + firstCellInRange + "))>0";
					break;
				case Asc.ECfType.containsBlanks:
					res = "LEN(TRIM(" + firstCellInRange + "))=0";
					break;
				case Asc.ECfType.timePeriod:
					res = _generateTimePeriodFunction();
					break;
			}
		}
		return res;
	};

	CConditionalFormattingRule.prototype.asc_setTimePeriod = function (val) {
		this.timePeriod = val;
		var formula = this.getFormulaByType();
		if (formula) {
			this.aRuleElements = [];
			this.aRuleElements[0] = new CFormulaCF();
			this.aRuleElements[0].Text = formula;
		}
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
		//чищу всегда, поскольку от интерфейса всегда заново выставляются оба значения
		this.aRuleElements = [];
		val = correctFromInterface(val);

		this.aRuleElements[0] = new CFormulaCF();
		this.aRuleElements[0].Text = val;
	};
	CConditionalFormattingRule.prototype.asc_setValue2 = function (val) {
		if (!this.aRuleElements) {
			this.aRuleElements = [];
		}

		val = correctFromInterface(val);

		this.aRuleElements[1] = new CFormulaCF();
		this.aRuleElements[1].Text = val;
	};

	CConditionalFormattingRule.prototype.asc_setColorScaleOrDataBarOrIconSetRule = function (val) {
		this.aRuleElements = [];
		this.aRuleElements.push(val);
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
				sheet = wb.getActiveWs();
				var _activeCell = sheet.selectionRange.activeCell;
				var _pivot = sheet.getPivotTable(_activeCell.col, _activeCell.row);
				if (_pivot) {
					range = _pivot.location && _pivot.location.ref;
				}

				if (!range) {
					sheet = null;
				}

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
	CConditionalFormattingRule.sStartLockCFId = 'cfrule_';

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
	CColorScale.prototype.applyPreset = function (styleIndex) {
		var presetStyles = conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.colorScale][styleIndex];
		for (var i = 0; i < presetStyles.length; i++) {
			var formatValueObject = new CConditionalFormatValueObject();
			formatValueObject.Type = presetStyles[i][0] ? presetStyles[i][0] : null;
			formatValueObject.Val = presetStyles[i][1] ? presetStyles[i][1] + "" : null;
			var colorObject = new AscCommonExcel.RgbColor(presetStyles[i][2] ? presetStyles[i][2] : 0);
			this.aCFVOs.push(formatValueObject);
			this.aColors.push(colorObject);
		}
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
				writer.WriteLong(this.aColors[i].getType());
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
				elem.Read_FromBinary2(reader);
				this.aCFVOs.push(elem);
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
					case AscCommonExcel.UndoRedoDataTypes.RgbColor:
						elem = new AscCommonExcel.RgbColor();
						break;
					case AscCommonExcel.UndoRedoDataTypes.ThemeColor:
						elem = new AscCommonExcel.ThemeColor();
						break;
				}
				if (null != elem.Read_FromBinary2) {
					elem.Read_FromBinary2(reader);
				} else if (null != elem.Read_FromBinary2AndReplace) {
					elem = elem.Read_FromBinary2AndReplace(reader);
				}
				this.aColors.push(elem);
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
	CColorScale.prototype.asc_getPreview = function (api, id) {
		return AscCommonExcel.drawGradientPreview(id, api.wb, this.aColors);
	};
	CColorScale.prototype.asc_setCFVOs = function (val) {
		this.aCFVOs = val;
	};
	CColorScale.prototype.asc_setColors = function (val) {
		var newArr = [];
		for (var i = 0; i < val.length; ++i) {
			if (this.aColors[i] && this.aColors[i].getR() === val[i].asc_getR() && this.aColors[i].getG() === val[i].asc_getG() && this.aColors[i].getB() === val[i].asc_getB()) {
				newArr.push(this.aColors[i]);
			} else {
				newArr.push(AscCommonExcel.CorrectAscColor(val[i]));
			}
		}
		this.aColors = newArr;
	};
	CColorScale.prototype.isEqual = function (elem) {
		if ((elem.aCFVOs && elem.aCFVOs.length === this.aCFVOs.length) || (!elem.aCFVOs && !this.aCFVOs)) {
			var i;
			if (elem.aCFVOs) {
				for (i = 0; i < elem.aCFVOs.length; i++) {
					if (!elem.aCFVOs[i].isEqual(this.aCFVOs[i])) {
						return false;
					}
				}
			}
			if ((elem.aColors && elem.aColors.length === this.aColors.length) || (!elem.aColors && !this.aColors)) {
				if (elem.aColors) {
					for (i = 0; i < elem.aColors.length; i++) {
						if (!elem.aColors[i].isEqual(this.aColors[i])) {
							return false;
						}
					}
				}
				return true;
			}
		}

		return false;
	};
	CColorScale.prototype.getType = function () {
		return window['AscCommonExcel'].UndoRedoDataTypes.ColorScale;
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
	CDataBar.prototype.isEqual = function (elem) {
		var _compareColors = function (_color1, _color2) {
			if (!_color1 && !_color2) {
				return true;
			}
			if (_color1 && _color2 && _color1.isEqual(_color2)) {
				return true;
			}
			return false;
		};

		if (this.MaxLength === elem.MaxLength && this.MinLength === elem.MinLength &&
			this.ShowValue === elem.ShowValue && this.AxisPosition === elem.AxisPosition &&
			this.Gradient === elem.Gradient && this.Direction === elem.Direction &&
			this.NegativeBarColorSameAsPositive === elem.NegativeBarColorSameAsPositive &&
			this.NegativeBarBorderColorSameAsPositive === elem.NegativeBarBorderColorSameAsPositive) {
			if (elem.aCFVOs && elem.aCFVOs.length === this.aCFVOs.length) {
				var i;
				for (i = 0; i < elem.aCFVOs.length; i++) {
					if (!elem.aCFVOs[i].isEqual(this.aCFVOs[i])) {
						return false;
					}
				}

				if (_compareColors(this.Color, elem.Color) && _compareColors(this.NegativeColor, elem.NegativeColor) &&
					_compareColors(this.BorderColor, elem.BorderColor) &&
					_compareColors(this.NegativeBorderColor, elem.NegativeBorderColor) &&
					_compareColors(this.AxisColor, elem.AxisColor)) {
					return true;
				}
			}
		}

		return false;
	};
	CDataBar.prototype.applyPreset = function (styleIndex) {
		var _generateRgbColor = function (_color) {
			if (_color === undefined || _color === null) {
				return null;
			}

			return new AscCommonExcel.RgbColor(_color);
		};

		var presetStyles = conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.dataBar][styleIndex];

		this.AxisColor = _generateRgbColor(presetStyles[0]);
		this.AxisPosition = 0;
		this.BorderColor = _generateRgbColor(presetStyles[1]);
		this.Color = _generateRgbColor(presetStyles[2]);
		this.Direction = 0;
		this.Gradient = presetStyles[3];
		this.MaxLength = 100;
		this.MinLength = 0;
		this.NegativeBarBorderColorSameAsPositive = false;
		this.NegativeBarColorSameAsPositive = false;
		this.NegativeBorderColor = _generateRgbColor(presetStyles[4]);
		this.NegativeColor = _generateRgbColor(presetStyles[5]);
		this.ShowValue = true;

		var formatValueObject1 = new CConditionalFormatValueObject();
		formatValueObject1.Type = AscCommonExcel.ECfvoType.AutoMin;
		this.aCFVOs.push(formatValueObject1);
		var formatValueObject2 = new CConditionalFormatValueObject();
		formatValueObject2.Type = AscCommonExcel.ECfvoType.AutoMax;
		this.aCFVOs.push(formatValueObject2);
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
			writer.WriteLong(this.Color.getType());
			this.Color.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.NegativeColor.getType());
			this.NegativeColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.BorderColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.BorderColor.getType());
			this.BorderColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.NegativeBorderColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.NegativeBorderColor.getType());
			this.NegativeBorderColor.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		if (null != this.AxisColor) {
			writer.WriteBool(true);
			writer.WriteLong(this.AxisColor.getType());
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
			this.NegativeBarBorderColorSameAsPositive = reader.GetBool();
		}

		var i, length, type, elem;
		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aCFVOs) {
					this.aCFVOs = [];
				}
				elem = new CConditionalFormatValueObject();
				elem.Read_FromBinary2(reader);
				this.aCFVOs.push(elem);
			}
		}

		var readColor = function () {
			var type = reader.GetLong();
			var _color;
			switch (type) {
				case AscCommonExcel.UndoRedoDataTypes.RgbColor:
					_color = new AscCommonExcel.RgbColor();
					break;
				case AscCommonExcel.UndoRedoDataTypes.ThemeColor:
					_color = new AscCommonExcel.ThemeColor();
					break;
			}
			if (null != _color.Read_FromBinary2) {
				_color.Read_FromBinary2(reader);
			} else if (null != _color.Read_FromBinary2AndReplace) {
				_color = _color.Read_FromBinary2AndReplace(reader);
			}
			return _color;
		};

		if (reader.GetBool()) {
			this.Color = readColor();
		}
		if (reader.GetBool()) {
			this.NegativeColor = readColor();
		}
		if (reader.GetBool()) {
			this.BorderColor = readColor();
		}
		if (reader.GetBool()) {
			this.NegativeBorderColor = readColor();
		}
		if (reader.GetBool()) {
			this.AxisColor = readColor();
		}
	};
	CDataBar.prototype.asc_setInterfaceDefault = function () {
		//ms всегда создаёт правило с такими настройками, хотя в документации други дефолтовые значения
		//дёргаем этот метод при создании нового правила из интерфейса
		this.MinLength = 0;
		this.MaxLength = 100;
	};
	CDataBar.prototype.asc_getPreview = function (api, id) {
		var color = this.Color;
		var aColors = [];
		var isReverse = this.Direction === AscCommonExcel.EDataBarDirection.rightToLeft;
		if (color) {
			if (this.Gradient) {
				var endColor = getDataBarGradientColor(color);
				if (isReverse) {
					aColors = [endColor, color];
				} else {
					aColors = [color, endColor];
				}
			} else {
				aColors = [color];
			}
		}

		AscCommonExcel.drawGradientPreview(id, api.wb, aColors, new AscCommon.CColor(202, 202, 202)/*this.settings.cells.defaultState.border*/, this.BorderColor, isReverse ? - 0.75 : 0.75, 2);
	};
	CDataBar.prototype.asc_getShowValue = function () {
		return this.ShowValue;
	};
	CDataBar.prototype.asc_getAxisPosition = function () {
		//TODO после открытия менять значения для условного формтирования без ext
		if (this.AxisPosition === AscCommonExcel.EDataBarAxisPosition.automatic && !this.AxisColor) {
			this.AxisPosition = AscCommonExcel.EDataBarAxisPosition.none;
		}
		return this.AxisPosition;
	};
	CDataBar.prototype.asc_getGradient = function () {
		return this.Gradient;
	};
	CDataBar.prototype.asc_getDirection = function () {
		return this.Direction;
	};
	CDataBar.prototype.asc_getNegativeBarColorSameAsPositive = function () {
		//TODO после открытия менять значения для условного формтирования без ext
		//в старом формате эта опция не используется
		//буду ориентироваться что если не задан NegativeColor, то эта опция выставляется в true
		if (!this.NegativeColor) {
			this.NegativeBarColorSameAsPositive = true;
		}
		return this.NegativeBarColorSameAsPositive;
	};
	CDataBar.prototype.asc_getNegativeBarBorderColorSameAsPositive = function () {
		return this.NegativeBarBorderColorSameAsPositive;
	};
	CDataBar.prototype.asc_getCFVOs = function () {
		return this.aCFVOs;
	};
	CDataBar.prototype.asc_getColor = function () {
		return this.Color ? Asc.colorObjToAscColor(this.Color) : null;
	};
	CDataBar.prototype.asc_getNegativeColor = function () {
		return this.NegativeColor ? Asc.colorObjToAscColor(this.NegativeColor) : null;
	};
	CDataBar.prototype.asc_getBorderColor = function () {
		return this.BorderColor ? Asc.colorObjToAscColor(this.BorderColor) : null;
	};
	CDataBar.prototype.asc_getNegativeBorderColor = function () {
		return this.NegativeBorderColor ? Asc.colorObjToAscColor(this.NegativeBorderColor) : null;
	};
	CDataBar.prototype.asc_getAxisColor = function () {
		return this.AxisColor ? Asc.colorObjToAscColor(this.AxisColor) : null;
	};

	CDataBar.prototype.asc_setShowValue = function (val) {
		this.ShowValue = val;
	};
	CDataBar.prototype.asc_setAxisPosition = function (val) {
		this.AxisPosition = val;
	};
	CDataBar.prototype.asc_setGradient = function (val) {
		this.Gradient = val;
	};
	CDataBar.prototype.asc_setDirection = function (val) {
		this.Direction = val;
	};
	CDataBar.prototype.asc_setNegativeBarColorSameAsPositive = function (val) {
		this.NegativeBarColorSameAsPositive = val;
	};
	CDataBar.prototype.asc_setNegativeBarBorderColorSameAsPositive = function (val) {
		this.NegativeBarBorderColorSameAsPositive = val;
	};
	CDataBar.prototype.asc_setCFVOs = function (val) {
		this.aCFVOs = val;
	};
	CDataBar.prototype.asc_setColor = function (val) {
		this.Color = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_setNegativeColor = function (val) {
		this.NegativeColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_setBorderColor = function (val) {
		this.BorderColor = AscCommonExcel.CorrectAscColor(val);
		if (val === null) {
			this.asc_setNegativeBorderColor(val);
		}
	};
	CDataBar.prototype.asc_setNegativeBorderColor = function (val) {
		this.NegativeBorderColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.asc_setAxisColor = function (val) {
		this.AxisColor = AscCommonExcel.CorrectAscColor(val);
	};
	CDataBar.prototype.getType = function () {
		return window['AscCommonExcel'].UndoRedoDataTypes.DataBar;
	};

	function CFormulaCF() {
		this.Text = null;
		this._f = null;

		return this;
	}

	CFormulaCF.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.CFormulaCF;
	};

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
	CFormulaCF.prototype.asc_getText = function () {
		return this.Text;
	};
	CFormulaCF.prototype.asc_setText = function (val) {
		this.Text = val;
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
		if (this.aCFVOs) {
			for (i = 0; i < this.aCFVOs.length; ++i) {
				res.aCFVOs.push(this.aCFVOs[i].clone());
			}
		}
		if (this.aIconSets) {
			for (i = 0; i < this.aIconSets.length; ++i) {
				res.aIconSets.push(this.aIconSets[i].clone());
			}
		}
		return res;
	};
	CIconSet.prototype.isEqual = function (elem) {
		if (this.IconSet === elem.IconSet && this.Percent === elem.Percent && this.Reverse === elem.Reverse &&
			this.ShowValue === elem.ShowValue) {
			if ((elem.aCFVOs && elem.aCFVOs.length === this.aCFVOs.length) || (!elem.aCFVOs && !this.aCFVOs)) {
				var i;
				if (elem.aCFVOs) {
					for (i = 0; i < elem.aCFVOs.length; i++) {
						if (!elem.aCFVOs[i].isEqual(this.aCFVOs[i])) {
							return false;
						}
					}
				}
				if (elem.aIconSets && elem.aIconSets.length === this.aIconSets.length) {
					for (i = 0; i < elem.aIconSets.length; i++) {
						if (!elem.aIconSets[i].isEqual(this.aIconSets[i])) {
							return false;
						}
					}
					return true;
				} else if (!elem.aIconSets && !this.aIconSets) {
					return true;
				}
			}
		}
		return false;
	};

	CIconSet.prototype.applyPreset = function (styleIndex) {
		var presetStyles = conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.icons][styleIndex];

		this.IconSet = styleIndex;
		for (var i = 0; i < presetStyles.length; i++) {
			var formatValueObject = new CConditionalFormatValueObject();
			formatValueObject.Type = presetStyles[i][0];
			formatValueObject.Val = presetStyles[i][1];
			if (presetStyles[i][2]) {
				formatValueObject.formula = new CFormulaCF();
				formatValueObject.formula.Text = presetStyles[i][2];
			}
			this.aCFVOs.push(formatValueObject);
		}
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
				elem.Read_FromBinary2(reader);
				this.aCFVOs.push(elem);
			}
		}

		if (reader.GetBool()) {
			length = reader.GetULong();
			for (i = 0; i < length; ++i) {
				if (!this.aIconSets) {
					this.aIconSets = [];
				}
				elem = new CConditionalFormatIconSet();
				elem.Read_FromBinary2(reader);
				this.aIconSets.push(elem);
			}
		}
	};
	CIconSet.prototype.asc_getPreview = function (api, id) {
		var i, aIconImgs = [];
		if (!this.Reverse) {
			for (i = this.aCFVOs.length - 1; i >= 0; i--) {
				aIconImgs.push(getCFIcon(this, i));
			}
		} else {
			for (i = 0; i < this.aCFVOs.length; i++) {
				aIconImgs.push(getCFIcon(this, i));
			}
		}
		AscCommonExcel.drawIconSetPreview(id, api.wb, aIconImgs);
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
		this.aIconSets = val == null ? [] : val;
	};
	CIconSet.prototype.getType = function () {
		return window['AscCommonExcel'].UndoRedoDataTypes.IconSet;
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
			writer.WriteBool(this.Gte);
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
			writer.WriteString2(this.Val);
		} else {
			writer.WriteBool(false);
		}
	};
	CConditionalFormatValueObject.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			this.Gte = reader.GetBool();
		}
		if (reader.GetBool()) {
			this.Type = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Val = reader.GetString2();
		}
	};
	CConditionalFormatValueObject.prototype.asc_getGte = function () {
		return this.Gte;
	};
	CConditionalFormatValueObject.prototype.asc_getType = function () {
		return this.Type;
	};
	CConditionalFormatValueObject.prototype.asc_getVal = function () {
		return !isNumeric(this.Val) ? "=" + this.Val : this.Val;
	};
	CConditionalFormatValueObject.prototype.asc_setGte = function (val) {
		this.Gte = val;
	};
	CConditionalFormatValueObject.prototype.asc_setType = function (val) {
		this.Type = val;
	};
	CConditionalFormatValueObject.prototype.asc_setVal = function (val) {
		val = correctFromInterface(val);
		this.Val = (val !== undefined && val !== null) ? val + "" : val;
	};
	CConditionalFormatValueObject.prototype.isEqual = function (elem) {
		if (this.Gte === elem.Gte && this.Type === elem.Type && this.Val === elem.Val && this.Type === elem.Type) {
			return true;
		}
		return false;
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
	CConditionalFormatIconSet.prototype.isEqual = function (val) {
		return this.IconSet === val.IconSet && this.IconId === val.IconId;
	};
	CConditionalFormatIconSet.prototype.Write_ToBinary2 = function (writer) {
		if (null != this.IconSet) {
			writer.WriteBool(true);
			writer.WriteLong(this.IconSet);
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
			this.IconSet = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.IconId = reader.GetLong();
		}
	};
	CConditionalFormatIconSet.prototype.asc_getIconSet = function () {
		return this.IconSet;
	};
	CConditionalFormatIconSet.prototype.asc_getIconId = function () {
		return this.IconId;
	};
	CConditionalFormatIconSet.prototype.asc_getIndex = function () {
		this.IconId = val;
	};
	CConditionalFormatIconSet.prototype.asc_setIconSet = function (val) {
		this.IconSet = val;
	};
	CConditionalFormatIconSet.prototype.asc_setIconId = function (val) {
		this.IconId = val;
	};
	CConditionalFormatIconSet.prototype.asc_setIndex = function (val) {
		this.IconId = val;
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

	function isValidDataRefCf(type, props) {
		var i;
		var ws;

		var checkFormulaStack = function (_f) {
			if (!_f) {
				return null;
			}
			var stack = _f.outStack;
			if (stack && stack.length) {
				//если идут фрифметические операции, использования диапазонов внутри формул - ошибки на это нет
				//поэтому я проверяю на одиночный диапазон
				if (stack.length === 1 && (stack[0].type === AscCommonExcel.cElementType.cellsRange ||
					stack[0].type === AscCommonExcel.cElementType.cellsRange3D)) {
					return asc_error.NotSingleReferenceCannotUsed;
				}

				if (type === Asc.ECfType.colorScale || type === Asc.ECfType.dataBar || type === Asc.ECfType.iconSet) {
					for (var i = 0; i < stack.length; i++) {
						if (stack[i]) {
							//допускаются только абсолютные ссылки
							if (stack[i].type === AscCommonExcel.cElementType.cellsRange ||
								stack[i].type === AscCommonExcel.cElementType.cellsRange3D ||
								stack[i].type === AscCommonExcel.cElementType.cell ||
								stack[i].type === AscCommonExcel.cElementType.cell3D) {
								//ссылки должны быть только абсолютные
								var _range = stack[i].getRange();
								if (_range.bbox) {
									_range = _range.bbox;
								}
								var isAbsRow1 = _range.isAbsRow(_range.refType1);
								var isAbsCol1 = _range.isAbsCol(_range.refType1);
								var isAbsRow2 = _range.isAbsRow(_range.refType2);
								var isAbsCol2 = _range.isAbsCol(_range.refType2);

								if (!isAbsRow1 || !isAbsCol1 || !isAbsRow2 || !isAbsCol2) {
									return asc_error.CannotUseRelativeReference;
								}
							}
						}
					}
				}
			}
			return null;
		};

		var _doParseFormula = function (sFormula) {
			if(!(typeof sFormula === "string" && sFormula.length > 0)) {
				return;
			}
			if (!ws) {
				var oWB = Asc.editor && Asc.editor.wbModel;
				if(!oWB) {
					return;
				}
				ws = oWB.getWorksheet(0);
			}

			if(sFormula.charAt(0) === '=') {
				sFormula = sFormula.slice(1);
			}

			var _formulaParsed = new AscCommonExcel.parserFormula(sFormula, null, ws);
			var _parseResultArg = new AscCommonExcel.ParseResult([], []);
			_formulaParsed.parse(true, true, _parseResultArg, true);

			return _formulaParsed;
		};


		var _checkValue = function(_val, _type, _isNumeric) {
			var fParser, _error;
			switch (_type) {
				case AscCommonExcel.ECfvoType.Formula:
					if (_isNumeric) {

					} else if (_val && _val[0] !== "=") {
						_val = '"' + _val + '"';
					} else {
						fParser = _doParseFormula(_val);
						//если внутри диапазон - проверяем его
						_error = fParser && checkFormulaStack(fParser);
						if (_error !== null) {
							return _error;
						}
					}
					break;
				case AscCommonExcel.ECfvoType.Number:
					if (_isNumeric) {

					} else if (_val && _val[0] !== "=") {
						_val = '"' + _val + '"';
					} else {
						fParser = _doParseFormula(_val);
						//если внутри диапазон - проверяем его
						_error = fParser && checkFormulaStack(fParser);
						if (_error !== null) {
							return _error;
						}
					}
					break;
				case AscCommonExcel.ECfvoType.Percent:
					if (_isNumeric) {
						if (_val < 0 && _val > 100) {
							//is not valid precentile
							return asc_error.NotValidPercentage;
						}
					} else if (_val && _val[0] !== "=") {
						_val = '"' + _val + '"';
					} else {
						fParser = _doParseFormula(_val);
						//если внутри диапазон - проверяем его
						_error = fParser && checkFormulaStack(fParser);
						if (_error !== null) {
							return _error;
						}
					}
					break;
				case AscCommonExcel.ECfvoType.Percentile:
					//в случае с индивидуальное проверкой Percentile - выдаём только 2 ошибки
					if (_isNumeric) {
						if (_val < 0 && _val > 100) {
							//is not valid precentile
							return asc_error.NotValidPercentile;
						}
					} else {
						return asc_error.CannotAddConditionalFormatting;
					}
					break;
			}

			return null;
		};

		var compareRefs = function (_prevVal, _prevType, _prevNum, _val, _type, _isNum) {
			//далее сравниваем ближайшие значения с одним типом, предыдущее должно быть меньше следующего
			//в databar ошибка для подобного сравнения не возникает
			//для iconSet сравниваем числа для типов Number/Percent/Percentile - должны идти по убыванию, сраниваем только соседние

			if (_prevNum && _isNum) {
				if (type === Asc.ECfType.colorScale) {
					if (_prevType === _type && type !== AscCommonExcel.ECfvoType.Formula && _prevVal > _val) {
						return asc_error.ValueMustBeGreaterThen;
					}
				} else if (type === Asc.ECfType.iconSet) {
					if (_prevType !== AscCommonExcel.ECfvoType.Formula && type !== AscCommonExcel.ECfvoType.Formula &&
						_val < _prevVal) {
						return asc_error.IconDataRangesOverlap;
					}
				}
			}

			return null;
		};

		//value, type
		var prevType, prevVal, prevNum;
		var nError;
		var _isNumeric;
		for (i = 0; i < props.length; i++) {
			if (undefined !== props[i][1] && type !== Asc.ECfType.top10) {
				_isNumeric = isNumeric(props[i][0]);
				nError = _checkValue(props[i][0], props[i][1], _isNumeric);
				if (nError !== null) {
					return [nError, i];
				}

				if (prevType === undefined) {
					prevType = props[i][1];
					prevVal = props[i][0];
					prevNum = _isNumeric;
				} else {
					nError = compareRefs(prevVal, prevType, prevNum, props[i][0], props[i][1], _isNumeric);
					if (nError !== null) {
						return [nError, i];
					}

					prevType = props[i][1];
					prevVal = props[i][0];
					prevNum = _isNumeric;
				}
			} else {
				//в этом случае должны быть следующие типы
				if (type === Asc.ECfType.expression) {
					nError = _checkValue(props[i][0], AscCommonExcel.ECfvoType.Formula);
					if (nError !== null) {
						return [nError, i];
					}
				} else if (type === Asc.ECfType.cellIs) {
					nError = _checkValue(props[i][0], AscCommonExcel.ECfvoType.Formula);
					if (nError !== null) {
						return [nError, i];
					}
				} else if (type === Asc.ECfType.containsText) {
					nError = _checkValue(props[i][0], AscCommonExcel.ECfvoType.Formula);
					if (nError !== null) {
						return [nError, i];
					}
				} else if (type === Asc.ECfType.top10) {
					_isNumeric = isNumeric(props[i][0]);
					var isPrecent = props[i][1];
					if (!_isNumeric) {
						return [asc_error.ErrorTop10Between, i];
					} else if (!isPrecent && (props[i][0] < 0 || props[i][0] > 1000)) {
						return [asc_error.ErrorTop10Between, i];
					} else if (isPrecent && (props[i][0] < 0 || props[i][0] > 100)) {
						return [asc_error.ErrorTop10Between, i];
					}
				}

			}
		}
	}

	function correctFromInterface(val) {
		var _isNumeric = isNumeric(val);
		if (!_isNumeric) {
			var isDate;
			var isFormula;

			if (val[0] === "=") {
				val = val.slice(1);
				isFormula = true;
			} else {
				isDate = AscCommon.g_oFormatParser.parseDate(val, AscCommon.g_oDefaultCultureInfo);
			}

			//храним число
			if (isDate) {
				val = isDate.value;
				return val;
			}

			if (!isFormula) {
				val = addQuotes(val);
			}
		}

		return val;
	}

	function addQuotes (val) {
		var _res;
		if (val[0] === '"') {
			_res = val.replace(/\"/g, "\"\"");
			_res = "\"" + _res + "\"";
		} else {
			_res = "\"" + val + "\"";
		}
		return _res;
	}

	var isNumeric = function (_val) {
		return !isNaN(parseFloat(_val)) && isFinite(_val);
	};

	var cDefIconSize = 16;
	var cDefIconFont = 11;

	var fullIconArray = ["data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik03LjUgMTVMMC41IDhINVYxSDEwVjhIMTQuNUw3LjUgMTVaIiBmaWxsPSIjRkYxMTExIi8+CjxwYXRoIGQ9Ik0xMCA4LjVIMTMuMjkyOUw3LjUgMTQuMjkyOUwxLjcwNzExIDguNUg1SDUuNVY4VjEuNUg5LjVWOFY4LjVIMTBaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xNSA4LjVMOCAxNS41TDggMTFMMSAxMUwxIDZMOCA2TDggMS41TDE1IDguNVoiIGZpbGw9IiNGRkNGMzMiLz4KPHBhdGggZD0iTTguNSA2TDguNSAyLjcwNzExTDE0LjI5MjkgOC41TDguNSAxNC4yOTI5TDguNSAxMUw4LjUgMTAuNUw4IDEwLjVMMS41IDEwLjVMMS41IDYuNUw4IDYuNUw4LjUgNi41TDguNSA2WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik03LjUgMUwxNC41IDhMMTAgOFYxNUw1IDE1TDUgOEwwLjUgOEw3LjUgMVoiIGZpbGw9IiMyRTk5NUYiLz4KPHBhdGggZD0iTTUgNy41TDEuNzA3MTEgNy41TDcuNSAxLjcwNzExTDEzLjI5MjkgNy41TDEwIDcuNUg5LjVWOFYxNC41TDUuNSAxNC41TDUuNSA4VjcuNUg1WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGcgY2xpcC1wYXRoPSJ1cmwoI2NsaXAwKSI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNMTQgMTRMNC4xMDA1MSAxNEw3LjI4MjQ5IDEwLjgxOEwyLjMzMjc0IDUuODY4MjdMNS44NjgyNyAyLjMzMjc0TDEwLjgxOCA3LjI4MjQ5TDE0IDQuMTAwNTFMMTQgMTRaIiBmaWxsPSIjNTA1MDUwIi8+CjxwYXRoIGQ9Ik0xMS4xNzE2IDcuNjM2MDRMMTMuNSA1LjMwNzYxTDEzLjUgMTMuNUw1LjMwNzYxIDEzLjVMNy42MzYwNCAxMS4xNzE2TDcuOTg5NTkgMTAuODE4TDcuNjM2MDQgMTAuNDY0NUwzLjAzOTg0IDUuODY4MjdMNS44NjgyNyAzLjAzOTg0TDEwLjQ2NDUgNy42MzYwNEwxMC44MTggNy45ODk1OUwxMS4xNzE2IDcuNjM2MDRaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjwvZz4KPGRlZnM+CjxjbGlwUGF0aCBpZD0iY2xpcDAiPgo8cmVjdCB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIGZpbGw9IndoaXRlIi8+CjwvY2xpcFBhdGg+CjwvZGVmcz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xNSA4LjVMOCAxNS41TDggMTFMMSAxMUwxIDZMOCA2TDggMS41TDE1IDguNVoiIGZpbGw9IiM1MDUwNTAiLz4KPHBhdGggZD0iTTguNSA2TDguNSAyLjcwNzExTDE0LjI5MjkgOC41TDguNSAxNC4yOTI5TDguNSAxMUw4LjUgMTAuNUw4IDEwLjVMMS41IDEwLjVMMS41IDYuNUw4IDYuNUw4LjUgNi41TDguNSA2WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik03LjUgMUwxNC41IDhMMTAgOFYxNUw1IDE1TDUgOEwwLjUgOEw3LjUgMVoiIGZpbGw9IiM1MDUwNTAiLz4KPHBhdGggZD0iTTUgNy41TDEuNzA3MTEgNy41TDcuNSAxLjcwNzExTDEzLjI5MjkgNy41TDEwIDcuNUg5LjVWOFYxNC41TDUuNSAxNC41TDUuNSA4VjcuNUg1WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTUgMUw1IDEwTDE0IDUuNUw1IDFaIiBmaWxsPSIjRkYxMTExIi8+CjxwYXRoIGQ9Ik0xMi44ODIgNS41TDUuNSA5LjE5MDk4TDUuNSAxLjgwOTAyTDEyLjg4MiA1LjVaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjIiIHk9IjAuOTk5OTk2IiB3aWR0aD0iMiIgaGVpZ2h0PSIxNCIgZmlsbD0iIzcyNzI3MiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTUgMUw1IDEwTDE0IDUuNUw1IDFaIiBmaWxsPSIjRkZDRjMzIi8+CjxwYXRoIGQ9Ik0xMi44ODIgNS41TDUuNSA5LjE5MDk4TDUuNSAxLjgwOTAyTDEyLjg4MiA1LjVaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjIiIHk9IjAuOTk5OTk2IiB3aWR0aD0iMiIgaGVpZ2h0PSIxNCIgZmlsbD0iIzcyNzI3MiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTUgMUw1IDEwTDE0IDUuNUw1IDFaIiBmaWxsPSIjMkU5OTVGIi8+CjxwYXRoIGQ9Ik0xMi44ODIgNS41TDUuNSA5LjE5MDk4TDUuNSAxLjgwOTAyTDEyLjg4MiA1LjVaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjIiIHk9IjAuOTk5OTk2IiB3aWR0aD0iMiIgaGVpZ2h0PSIxNCIgZmlsbD0iIzcyNzI3MiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iOCIgeT0iMSIgd2lkdGg9IjkuODk5NSIgaGVpZ2h0PSI5Ljg5OTUiIHJ4PSIyIiB0cmFuc2Zvcm09InJvdGF0ZSg0NSA4IDEpIiBmaWxsPSIjRkYxMTExIi8+CjxyZWN0IHg9IjgiIHk9IjEuNzA3MTEiIHdpZHRoPSI4Ljg5OTUiIGhlaWdodD0iOC44OTk1IiByeD0iMS41IiB0cmFuc2Zvcm09InJvdGF0ZSg0NSA4IDEuNzA3MTEpIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgMTVIMTVMOCAxTDEgMTVaIiBmaWxsPSIjRkZDRjMzIi8+CjxwYXRoIGQ9Ik04IDIuMTE4MDNMMTQuMTkxIDE0LjVIMS44MDkwMkw4IDIuMTE4MDNaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDhDMTUgMTEuODY2IDExLjg2NiAxNSA4IDE1QzQuMTM0MDEgMTUgMSAxMS44NjYgMSA4QzEgNC4xMzQwMSA0LjEzNDAxIDEgOCAxQzExLjg2NiAxIDE1IDQuMTM0MDEgMTUgOFoiIGZpbGw9IiMyRTk5NUYiLz4KPHBhdGggZD0iTTE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTkgMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOFoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE0Ljk5OTkgNy45OTk5NkMxNC45OTk5IDExLjg2NTkgMTEuODY1OSAxNC45OTk5IDcuOTk5OTYgMTQuOTk5OUM0LjEzMzk5IDE0Ljk5OTkgMSAxMS44NjU5IDEgNy45OTk5NkMxIDQuMTMzOTkgNC4xMzM5OSAxIDcuOTk5OTYgMUMxMS44NjU5IDEgMTQuOTk5OSA0LjEzMzk5IDE0Ljk5OTkgNy45OTk5NloiIGZpbGw9IiNGRjExMTEiLz4KPHBhdGggZD0iTTE0LjQ5OTkgNy45OTk5NkMxNC40OTk5IDExLjU4OTggMTEuNTg5OCAxNC40OTk5IDcuOTk5OTYgMTQuNDk5OUM0LjQxMDEzIDE0LjQ5OTkgMS41IDExLjU4OTggMS41IDcuOTk5OTZDMS41IDQuNDEwMTMgNC40MTAxMyAxLjUgNy45OTk5NiAxLjVDMTEuNTg5OCAxLjUgMTQuNDk5OSA0LjQxMDEzIDE0LjQ5OTkgNy45OTk5NloiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik02LjU4NTggOEw0LjI5MjkxIDUuNzA3MTFMNS43MDcxMiA0LjI5Mjg5TDguMDAwMDEgNi41ODU3OUwxMC4yOTI5IDQuMjkyODlMMTEuNzA3MSA1LjcwNzExTDkuNDE0MjMgOEwxMS43MDcxIDEwLjI5MjlMMTAuMjkyOSAxMS43MDcxTDguMDAwMDEgOS40MTQyMUw1LjcwNzEyIDExLjcwNzFMNC4yOTI5MSAxMC4yOTI5TDYuNTg1OCA4WiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDhDMTUgMTEuODY2IDExLjg2NiAxNSA4IDE1QzQuMTM0MDEgMTUgMSAxMS44NjYgMSA4QzEgNC4xMzQwMSA0LjEzNDAxIDEgOCAxQzExLjg2NiAxIDE1IDQuMTM0MDEgMTUgOFoiIGZpbGw9IiNGRkNGMzMiLz4KPHBhdGggZD0iTTE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTkgMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOFoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHBhdGggZD0iTTcuMjcxNTIgOC42NTM5TDcuMDU5NiA1LjQ5MDA5QzcuMDE5ODcgNC44NzM2MiA3IDQuNDMxMDkgNyA0LjE2MjQ4QzcgMy43OTcwMSA3LjA5NDkyIDMuNTEyOTkgNy4yODQ3NyAzLjMxMDQ0QzcuNDc5MDMgMy4xMDM0OCA3LjczMjg5IDMgOC4wNDYzNiAzQzguNDI2MDUgMyA4LjY3OTkxIDMuMTMyMSA4LjgwNzk1IDMuMzk2M0M4LjkzNTk4IDMuNjU2MSA5IDQuMDMyNTggOSA0LjUyNTc2QzkgNC44MTYzOCA4Ljk4NDU1IDUuMTExNCA4Ljk1MzY0IDUuNDEwODNMOC42Njg4NyA4LjY2NzExQzguNjM3OTcgOS4wNTQ2IDguNTcxNzQgOS4zNTE4MyA4LjQ3MDIgOS41NTg3OEM4LjM2ODY1IDkuNzY1NzQgOC4yMDA4OCA5Ljg2OTIyIDcuOTY2ODkgOS44NjkyMkM3LjcyODQ4IDkuODY5MjIgNy41NjI5MSA5Ljc3MDE1IDcuNDcwMiA5LjU3MTk5QzcuMzc3NDggOS4zNjk0NCA3LjMxMTI2IDkuMDYzNDEgNy4yNzE1MiA4LjY1MzlaTTguMDA2NjIgMTNDNy43MzczMSAxMyA3LjUwMTEgMTIuOTE0MSA3LjI5ODAxIDEyLjc0MjRDNy4wOTkzNCAxMi41NjYzIDcgMTIuMzIxOSA3IDEyLjAwOTJDNyAxMS43MzYyIDcuMDk0OTIgMTEuNTA1MSA3LjI4NDc3IDExLjMxNTdDNy40NzkwMyAxMS4xMjIgNy43MTUyMyAxMS4wMjUxIDcuOTkzMzggMTEuMDI1MUM4LjI3MTUyIDExLjAyNTEgOC41MDc3MyAxMS4xMjIgOC43MDE5OSAxMS4zMTU3QzguOTAwNjYgMTEuNTA1MSA5IDExLjczNjIgOSAxMi4wMDkyQzkgMTIuMzE3NSA4LjkwMDY2IDEyLjU1OTcgOC43MDE5OSAxMi43MzU4QzguNTAzMzEgMTIuOTExOSA4LjI3MTUyIDEzIDguMDA2NjIgMTNaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDhDMTUgMTEuODY2IDExLjg2NiAxNSA4IDE1QzQuMTM0MDEgMTUgMSAxMS44NjYgMSA4QzEgNC4xMzQwMSA0LjEzNDAxIDEgOCAxQzExLjg2NiAxIDE1IDQuMTM0MDEgMTUgOFoiIGZpbGw9IiMyRTk5NUYiLz4KPHBhdGggZD0iTTE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTkgMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOFoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMi44MDUgNS41OTMyTDcuNjM5MTUgMTIuNjA0MUw0LjI0MTY3IDguNjUxODlMNS43NTgzIDcuMzQ4MTFMNy41MTg3MiA5LjM5NTk0TDExLjE5NDkgNC40MDY4TDEyLjgwNSA1LjU5MzJaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik02LjE5NzM1IDhMMiAzLjgwMjY1TDMuODAyNjUgMkw4IDYuMTk3MzVMMTIuMTk3MyAyTDE0IDMuODAyNjVMOS44MDI2NSA4TDE0IDEyLjE5NzNMMTIuMTk3MyAxNEw4IDkuODAyNjVMMy44MDI2NSAxNEwyIDEyLjE5NzNMNi4xOTczNSA4WiIgZmlsbD0iI0ZGMTExMSIvPgo8cGF0aCBkPSJNOC4zNTM1NSA2LjU1MDlMMTIuMTk3MyAyLjcwNzExTDEzLjI5MjkgMy44MDI2NUw5LjQ0OTEgNy42NDY0NUw5LjA5NTU1IDhMOS40NDkxIDguMzUzNTVMMTMuMjkyOSAxMi4xOTczTDEyLjE5NzMgMTMuMjkyOUw4LjM1MzU1IDkuNDQ5MUw4IDkuMDk1NTVMNy42NDY0NSA5LjQ0OTFMMy44MDI2NSAxMy4yOTI5TDIuNzA3MTEgMTIuMTk3M0w2LjU1MDkgOC4zNTM1NUw2LjkwNDQ1IDhMNi41NTA5IDcuNjQ2NDVMMi43MDcxMSAzLjgwMjY1TDMuODAyNjUgMi43MDcxMUw3LjY0NjQ1IDYuNTUwOUw4IDYuOTA0NDVMOC4zNTM1NSA2LjU1MDlaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTcuNDA3MjggOC45MTU0Nkw3LjA4OTQgNC40ODYxM0M3LjAyOTggMy42MjMwNyA3IDMuMDAzNTIgNyAyLjYyNzQ4QzcgMi4xMTU4MSA3LjE0MjM4IDEuNzE4MTkgNy40MjcxNSAxLjQzNDYxQzcuNzE4NTQgMS4xNDQ4NyA4LjA5OTM0IDEgOC41Njk1NCAxQzkuMTM5MDcgMSA5LjUxOTg3IDEuMTg0OTQgOS43MTE5MiAxLjU1NDgyQzkuOTAzOTcgMS45MTg1NCAxMCAyLjQ0NTYyIDEwIDMuMTM2MDZDMTAgMy41NDI5MyA5Ljk3NjgyIDMuOTU1OTcgOS45MzA0NiA0LjM3NTE2TDkuNTAzMzEgOC45MzM5NUM5LjQ1Njk1IDkuNDc2NDQgOS4zNTc2MiA5Ljg5MjU2IDkuMjA1MyAxMC4xODIzQzkuMDUyOTggMTAuNDcyIDguODAxMzIgMTAuNjE2OSA4LjQ1MDMzIDEwLjYxNjlDOC4wOTI3MiAxMC42MTY5IDcuODQ0MzcgMTAuNDc4MiA3LjcwNTMgMTAuMjAwOEM3LjU2NjIzIDkuOTE3MjIgNy40NjY4OSA5LjQ4ODc3IDcuNDA3MjggOC45MTU0NlpNOC41MDk5MyAxNUM4LjEwNTk2IDE1IDcuNzUxNjYgMTQuODc5OCA3LjQ0NzAyIDE0LjYzOTRDNy4xNDkwMSAxNC4zOTI4IDcgMTQuMDUwNiA3IDEzLjYxMjlDNyAxMy4yMzA3IDcuMTQyMzggMTIuOTA3MSA3LjQyNzE1IDEyLjY0MkM3LjcxODU0IDEyLjM3MDggOC4wNzI4NSAxMi4yMzUxIDguNDkwMDcgMTIuMjM1MUM4LjkwNzI4IDEyLjIzNTEgOS4yNjE1OSAxMi4zNzA4IDkuNTUyOTggMTIuNjQyQzkuODUwOTkgMTIuOTA3MSAxMCAxMy4yMzA3IDEwIDEzLjYxMjlDMTAgMTQuMDQ0NSA5Ljg1MDk5IDE0LjM4MzUgOS41NTI5OCAxNC42MzAxQzkuMjU0OTcgMTQuODc2NyA4LjkwNzI4IDE1IDguNTA5OTMgMTVaIiBmaWxsPSIjRkZDRjMzIi8+CjxwYXRoIGQ9Ik05LjI2ODE3IDEuNzg1MjNMOS4yNjgxNiAxLjc4NTIzTDkuMjY5NzcgMS43ODgyOUM5LjQwNjI4IDIuMDQ2ODEgOS41IDIuNDc4NTEgOS41IDMuMTM2MDZDOS41IDMuNTIzOTIgOS40Nzc5MSAzLjkxODYgOS40MzM0OSA0LjMyMDIxTDkuNDMzNDIgNC4zMjAyTDkuNDMyNjQgNC4zMjg1Mkw5LjAwNTQ5IDguODg3M0w5LjAwNTQ4IDguODg3M0w5LjAwNTEzIDguODkxMzhDOC45NjExNyA5LjQwNTczIDguODcwMDkgOS43NDU0MSA4Ljc2MjczIDkuOTQ5NjRDOC43MjQ4NyAxMC4wMjE3IDguNjg2NDkgMTAuMDU1NiA4LjY1Mjg3IDEwLjA3NDlDOC42MTczMSAxMC4wOTU0IDguNTU2NTIgMTAuMTE2OSA4LjQ1MDMzIDEwLjExNjlDOC4zMzUyOSAxMC4xMTY5IDguMjcyNTEgMTAuMDk0NyA4LjIzOTY3IDEwLjA3NjRDOC4yMTEyNCAxMC4wNjA1IDguMTgxNzggMTAuMDM0OSA4LjE1MzE2IDkuOTc4NDdDOC4wNTQxMSA5Ljc3NTMxIDcuOTYyOTUgOS40MiA3LjkwNTQyIDguODcxNThMNy41ODgyMiA0LjQ1MTY4QzcuNTg4MiA0LjQ1MTQ1IDcuNTg4MTggNC40NTEyMyA3LjU4ODE3IDQuNDUxQzcuNTI4NjQgMy41ODg5NSA3LjUgMi45ODQ2NiA3LjUgMi42Mjc0OEM3LjUgMi4yMTA3MiA3LjYxMzM0IDEuOTU0OSA3Ljc3OTg0IDEuNzg5MDNDNy45NjM4MiAxLjYwNjE2IDguMjEwODcgMS41IDguNTY5NTQgMS41QzkuMDMxMTYgMS41IDkuMTkzMiAxLjY0MDg0IDkuMjY4MTcgMS43ODUyM1pNOS4yMTIzIDEzLjAwOEw5LjIxMjIyIDEzLjAwODFMOS4yMjA2NyAxMy4wMTU2QzkuNDE3MjkgMTMuMTkwNSA5LjUgMTMuMzgwNCA5LjUgMTMuNjEyOUM5LjUgMTMuOTE0OSA5LjQwMzI1IDE0LjEwNSA5LjIzNDIzIDE0LjI0NDlDOS4wMjg1IDE0LjQxNTEgOC43OTQzNyAxNC41IDguNTA5OTMgMTQuNUM4LjIxNTU5IDE0LjUgNy45NzI5NyAxNC40MTU5IDcuNzYxNDkgMTQuMjUwNkM3LjU5Njg0IDE0LjExMjUgNy41IDEzLjkyMSA3LjUgMTMuNjEyOUM3LjUgMTMuMzcyNCA3LjU4MjYxIDEzLjE4MDQgNy43Njc4MyAxMy4wMDhDNy45NjE5MSAxMi44MjczIDguMTkyNTggMTIuNzM1MSA4LjQ5MDA3IDEyLjczNTFDOC43ODc1NSAxMi43MzUxIDkuMDE4MjMgMTIuODI3MyA5LjIxMjMgMTMuMDA4WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xNCA0LjI5NTc4TDYuNjgyMDMgMTRMMiA4LjY3ODIxTDMuNjM0MTcgNy4yNTYwOEw2LjU1MjE4IDEwLjU3MjhMMTIuMjYyOCAzTDE0IDQuMjk1NzhaIiBmaWxsPSIjMkU5OTVGIi8+CjxwYXRoIGQ9Ik02Ljk1MTM5IDEwLjg3MzlMMTIuMzYyNiAzLjY5ODE4TDEzLjI5ODIgNC4zOTYwNEw2LjY1MjIgMTMuMjA5MUwyLjcwNzQ2IDguNzI1MzdMMy41ODcyNyA3Ljk1OTczTDYuMTc2NzggMTAuOTAzMUw2LjU4MjAyIDExLjM2MzdMNi45NTEzOSAxMC44NzM5WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDhDMTUgMTEuODY2IDExLjg2NiAxNSA4IDE1QzQuMTM0MDEgMTUgMSAxMS44NjYgMSA4QzEgNC4xMzQwMSA0LjEzNDAxIDEgOCAxQzExLjg2NiAxIDE1IDQuMTM0MDEgMTUgOFoiIGZpbGw9IiNGRjExMTEiLz4KPHBhdGggZD0iTTE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTkgMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOFoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDhDMTUgMTEuODY2IDExLjg2NiAxNSA4IDE1QzQuMTM0MDEgMTUgMSAxMS44NjYgMSA4QzEgNC4xMzQwMSA0LjEzNDAxIDEgOCAxQzExLjg2NiAxIDE1IDQuMTM0MDEgMTUgOFoiIGZpbGw9IiNGRkNGMzMiLz4KPHBhdGggZD0iTTE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTkgMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOFoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgM0MxIDEuODk1NDMgMS44OTU0MyAxIDMgMUgxM0MxNC4xMDQ2IDEgMTUgMS44OTU0MyAxNSAzVjEzQzE1IDE0LjEwNDYgMTQuMTA0NiAxNSAxMyAxNUgzQzEuODk1NDMgMTUgMSAxNC4xMDQ2IDEgMTNWM1oiIGZpbGw9IiM1MDUwNTAiLz4KPHBhdGggZD0iTTMgMS41SDEzQzEzLjgyODQgMS41IDE0LjUgMi4xNzE1NyAxNC41IDNWMTNDMTQuNSAxMy44Mjg0IDEzLjgyODQgMTQuNSAxMyAxNC41SDNDMi4xNzE1NyAxNC41IDEuNSAxMy44Mjg0IDEuNSAxM1YzQzEuNSAyLjE3MTU3IDIuMTcxNTcgMS41IDMgMS41WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cGF0aCBkPSJNMTQgOEMxNCA0LjY4NjI5IDExLjMxMzcgMiA4IDJDNC42ODYyOSAyIDIgNC42ODYyOSAyIDhDMiAxMS4zMTM3IDQuNjg2MjkgMTQgOCAxNEMxMS4zMTM3IDE0IDE0IDExLjMxMzcgMTQgOFoiIGZpbGw9IndoaXRlIi8+CjxwYXRoIGQ9Ik0xMyA4QzEzIDUuMjM4NTggMTAuNzYxNCAzIDggM0M1LjIzODU4IDMgMyA1LjIzODU4IDMgOEMzIDEwLjc2MTQgNS4yMzg1OCAxMyA4IDEzQzEwLjc2MTQgMTMgMTMgMTAuNzYxNCAxMyA4WiIgZmlsbD0iI0ZGMTExMSIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgM0MxIDEuODk1NDMgMS44OTU0MyAxIDMgMUgxM0MxNC4xMDQ2IDEgMTUgMS44OTU0MyAxNSAzVjEzQzE1IDE0LjEwNDYgMTQuMTA0NiAxNSAxMyAxNUgzQzEuODk1NDMgMTUgMSAxNC4xMDQ2IDEgMTNWM1oiIGZpbGw9IiM1MDUwNTAiLz4KPHBhdGggZD0iTTMgMS41SDEzQzEzLjgyODQgMS41IDE0LjUgMi4xNzE1NyAxNC41IDNWMTNDMTQuNSAxMy44Mjg0IDEzLjgyODQgMTQuNSAxMyAxNC41SDNDMi4xNzE1NyAxNC41IDEuNSAxMy44Mjg0IDEuNSAxM1YzQzEuNSAyLjE3MTU3IDIuMTcxNTcgMS41IDMgMS41WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cGF0aCBkPSJNMTQgOEMxNCA0LjY4NjI5IDExLjMxMzcgMiA4IDJDNC42ODYyOSAyIDIgNC42ODYyOSAyIDhDMiAxMS4zMTM3IDQuNjg2MjkgMTQgOCAxNEMxMS4zMTM3IDE0IDE0IDExLjMxMzcgMTQgOFoiIGZpbGw9IndoaXRlIi8+CjxwYXRoIGQ9Ik0xMyA4QzEzIDUuMjM4NTggMTAuNzYxNCAzIDggM0M1LjIzODU4IDMgMyA1LjIzODU4IDMgOEMzIDEwLjc2MTQgNS4yMzg1OCAxMyA4IDEzQzEwLjc2MTQgMTMgMTMgMTAuNzYxNCAxMyA4WiIgZmlsbD0iI0ZGQ0YzMyIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgM0MxIDEuODk1NDMgMS44OTU0MyAxIDMgMUgxM0MxNC4xMDQ2IDEgMTUgMS44OTU0MyAxNSAzVjEzQzE1IDE0LjEwNDYgMTQuMTA0NiAxNSAxMyAxNUgzQzEuODk1NDMgMTUgMSAxNC4xMDQ2IDEgMTNWM1oiIGZpbGw9IiM1MDUwNTAiLz4KPHBhdGggZD0iTTMgMS41SDEzQzEzLjgyODQgMS41IDE0LjUgMi4xNzE1NyAxNC41IDNWMTNDMTQuNSAxMy44Mjg0IDEzLjgyODQgMTQuNSAxMyAxNC41SDNDMi4xNzE1NyAxNC41IDEuNSAxMy44Mjg0IDEuNSAxM1YzQzEuNSAyLjE3MTU3IDIuMTcxNTcgMS41IDMgMS41WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cGF0aCBkPSJNMTQgOEMxNCA0LjY4NjI5IDExLjMxMzcgMiA4IDJDNC42ODYyOSAyIDIgNC42ODYyOSAyIDhDMiAxMS4zMTM3IDQuNjg2MjkgMTQgOCAxNEMxMS4zMTM3IDE0IDE0IDExLjMxMzcgMTQgOFoiIGZpbGw9IndoaXRlIi8+CjxwYXRoIGQ9Ik0xMyA4QzEzIDUuMjM4NTggMTAuNzYxNCAzIDggM0M1LjIzODU4IDMgMyA1LjIzODU4IDMgOEMzIDEwLjc2MTQgNS4yMzg1OCAxMyA4IDEzQzEwLjc2MTQgMTMgMTMgMTAuNzYxNCAxMyA4WiIgZmlsbD0iIzJFOTk1RiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMyAxNUwzLjEwMDUzIDE1TDYuMjgyNTEgMTEuODE4TDEuMzMyNzYgNi44NjgyN0w0Ljg2ODMgMy4zMzI3M0w5LjgxODA1IDguMjgyNDhMMTMgNS4xMDA1TDEzIDE1WiIgZmlsbD0iI0ZGQ0YzMyIvPgo8cGF0aCBkPSJNNi42MzYwNyAxMS40NjQ1TDIuMDM5ODcgNi44NjgyN0w0Ljg2ODMgNC4wMzk4NEw5LjQ2NDQ5IDguNjM2MDNMOS44MTgwNSA4Ljk4OTU5TDEwLjE3MTYgOC42MzYwM0wxMi41IDYuMzA3NjFMMTIuNSAxNC41TDQuMzA3NjQgMTQuNUw2LjYzNjA3IDEyLjE3MTZMNi45ODk2MiAxMS44MThMNi42MzYwNyAxMS40NjQ1WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMyAzTDMuMTAwNSAzTDYuMjgyNDggNi4xODE5OEwxLjMzMjczIDExLjEzMTdMNC44NjgyNyAxNC42NjczTDkuODE4MDEgOS43MTc1MUwxMyAxMi44OTk1VjNaIiBmaWxsPSIjRkZDRjMzIi8+CjxwYXRoIGQ9Ik02LjYzNjAzIDUuODI4NDNMNC4zMDc2MSAzLjVMMTIuNSAzLjVWMTEuNjkyNEwxMC4xNzE2IDkuMzYzOTZMOS44MTgwMSA5LjAxMDQxTDkuNDY0NDYgOS4zNjM5Nkw0Ljg2ODI3IDEzLjk2MDJMMi4wMzk4NCAxMS4xMzE3TDYuNjM2MDMgNi41MzU1M0w2Ljk4OTU5IDYuMTgxOThMNi42MzYwMyA1LjgyODQzWiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGcgY2xpcC1wYXRoPSJ1cmwoI2NsaXAwKSI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNMTQgMTRMNC4xMDA1MSAxNEw3LjI4MjQ5IDEwLjgxOEwyLjMzMjc0IDUuODY4MjdMNS44NjgyNyAyLjMzMjc0TDEwLjgxOCA3LjI4MjQ5TDE0IDQuMTAwNTFMMTQgMTRaIiBmaWxsPSIjNTA1MDUwIi8+CjxwYXRoIGQ9Ik0xMS4xNzE2IDcuNjM2MDRMMTMuNSA1LjMwNzYxTDEzLjUgMTMuNUw1LjMwNzYxIDEzLjVMNy42MzYwNCAxMS4xNzE2TDcuOTg5NTkgMTAuODE4TDcuNjM2MDQgMTAuNDY0NUwzLjAzOTg0IDUuODY4MjdMNS44NjgyNyAzLjAzOTg0TDEwLjQ2NDUgNy42MzYwNEwxMC44MTggNy45ODk1OUwxMS4xNzE2IDcuNjM2MDRaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjwvZz4KPGRlZnM+CjxjbGlwUGF0aCBpZD0iY2xpcDAiPgo8cmVjdCB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIGZpbGw9IndoaXRlIi8+CjwvY2xpcFBhdGg+CjwvZGVmcz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGcgY2xpcC1wYXRoPSJ1cmwoI2NsaXAwKSI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNMTQgMkw0LjEwMDUxIDJMNy4yODI0OSA1LjE4MTk4TDIuMzMyNzQgMTAuMTMxN0w1Ljg2ODI3IDEzLjY2NzNMMTAuODE4IDguNzE3NTFMMTQgMTEuODk5NUwxNCAyWiIgZmlsbD0iIzUwNTA1MCIvPgo8cGF0aCBkPSJNMTEuMTcxNiA4LjM2Mzk2TDEzLjUgMTAuNjkyNEwxMy41IDIuNUw1LjMwNzYxIDIuNUw3LjYzNjA0IDQuODI4NDNMNy45ODk1OSA1LjE4MTk4TDcuNjM2MDQgNS41MzU1M0wzLjAzOTg1IDEwLjEzMTdMNS44NjgyNyAxMi45NjAyTDEwLjQ2NDUgOC4zNjM5NkwxMC44MTggOC4wMTA0MUwxMS4xNzE2IDguMzYzOTZaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjwvZz4KPGRlZnM+CjxjbGlwUGF0aCBpZD0iY2xpcDAiPgo8cmVjdCB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIGZpbGw9IndoaXRlIi8+CjwvY2xpcFBhdGg+CjwvZGVmcz4KPC9zdmc+Cg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMTAiIHdpZHRoPSIzIiBoZWlnaHQ9IjUiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iMS41IiB5PSIxMC41IiB3aWR0aD0iMiIgaGVpZ2h0PSI0IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iNS41IiB5PSI3LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjciIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOSIgeT0iNCIgd2lkdGg9IjMiIGhlaWdodD0iMTEiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iOS41IiB5PSI0LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEwIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxMy41IiB5PSIxLjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEzIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMTAiIHdpZHRoPSIzIiBoZWlnaHQ9IjUiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iMS41IiB5PSIxMC41IiB3aWR0aD0iMiIgaGVpZ2h0PSI0IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iNS41IiB5PSI3LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjciIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOSIgeT0iNCIgd2lkdGg9IjMiIGhlaWdodD0iMTEiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iOS41IiB5PSI0LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEwIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxMy41IiB5PSIxLjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEzIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMTAiIHdpZHRoPSIzIiBoZWlnaHQ9IjUiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iMS41IiB5PSIxMC41IiB3aWR0aD0iMiIgaGVpZ2h0PSI0IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iNS41IiB5PSI3LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjciIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOSIgeT0iNCIgd2lkdGg9IjMiIGhlaWdodD0iMTEiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iOS41IiB5PSI0LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEwIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxMy41IiB5PSIxLjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEzIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMTAiIHdpZHRoPSIzIiBoZWlnaHQ9IjUiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iMS41IiB5PSIxMC41IiB3aWR0aD0iMiIgaGVpZ2h0PSI0IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iNS41IiB5PSI3LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjciIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOSIgeT0iNCIgd2lkdGg9IjMiIGhlaWdodD0iMTEiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iOS41IiB5PSI0LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEwIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSIxMy41IiB5PSIxLjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEzIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iOCIgY3k9IjgiIHI9IjciIGZpbGw9IiM1MDUwNTAiLz4KPGNpcmNsZSBjeD0iOCIgY3k9IjgiIHI9IjYuNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGNpcmNsZSBjeD0iOCIgY3k9IjgiIHI9IjciIGZpbGw9IiM5QjlCOUIiLz4KPGNpcmNsZSBjeD0iOCIgY3k9IjgiIHI9IjYuNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGVsbGlwc2UgY3g9IjgiIGN5PSI4IiByeD0iNyIgcnk9IjciIHRyYW5zZm9ybT0icm90YXRlKC0xODAgOCA4KSIgZmlsbD0iI0ZGODA4MCIvPgo8cGF0aCBkPSJNMS41IDhDMS41IDQuNDEwMTUgNC40MTAxNSAxLjUgOCAxLjVDMTEuNTg5OSAxLjUgMTQuNSA0LjQxMDE1IDE0LjUgOEMxNC41IDExLjU4OTkgMTEuNTg5OSAxNC41IDggMTQuNUM0LjQxMDE1IDE0LjUgMS41IDExLjU4OTggMS41IDhaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxjaXJjbGUgY3g9IjgiIGN5PSI4IiByPSI2IiB0cmFuc2Zvcm09InJvdGF0ZSgtOTAgOCA4KSIgZmlsbD0id2hpdGUiLz48L3N2Zz4=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNi44MTMzMSAyIDUuNjUzMjggMi4zNTE4OSA0LjY2NjU4IDMuMDExMThDMy42Nzk4OSAzLjY3MDQ3IDIuOTEwODUgNC42MDc1NCAyLjQ1NjczIDUuNzAzOUMyLjAwMjYgNi44MDAyNSAxLjg4Mzc4IDguMDA2NjUgMi4xMTUyOSA5LjE3MDU0QzIuMzQ2OCAxMC4zMzQ0IDIuOTE4MjUgMTEuNDAzNSAzLjc1NzM2IDEyLjI0MjZDNC41OTY0OCAxMy4wODE4IDUuNjY1NTggMTMuNjUzMiA2LjgyOTQ2IDEzLjg4NDdDNy45OTMzNSAxNC4xMTYyIDkuMTk5NzUgMTMuOTk3NCAxMC4yOTYxIDEzLjU0MzNDMTEuMzkyNSAxMy4wODkxIDEyLjMyOTUgMTIuMzIwMSAxMi45ODg4IDExLjMzMzRDMTMuNjQ4MSAxMC4zNDY3IDE0IDkuMTg2NjggMTQgOEw4IDhMOCAyWiIgZmlsbD0id2hpdGUiLz48L3N2Zz4=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNi40MDg3IDIgNC44ODI1OCAyLjYzMjE0IDMuNzU3MzYgMy43NTczNkMyLjYzMjE0IDQuODgyNTggMiA2LjQwODcgMiA4QzIgOS41OTEzIDIuNjMyMTQgMTEuMTE3NCAzLjc1NzM2IDEyLjI0MjZDNC44ODI1OCAxMy4zNjc5IDYuNDA4NyAxNCA4IDE0TDggOEw4IDJaIiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48Y2lyY2xlIGN4PSI4IiBjeT0iOCIgcj0iNyIgZmlsbD0iIzUwNTA1MCIvPjxwYXRoIGQ9Ik04IDJDNy4yMTIwNyAyIDYuNDMxODUgMi4xNTUxOSA1LjcwMzkgMi40NTY3MkM0Ljk3NTk1IDIuNzU4MjUgNC4zMTQ1MSAzLjIwMDIxIDMuNzU3MzYgMy43NTczNkMzLjIwMDIxIDQuMzE0NTEgMi43NTgyNSA0Ljk3NTk1IDIuNDU2NzIgNS43MDM5QzIuMTU1MTkgNi40MzE4NSAyIDcuMjEyMDcgMiA4TDggOEw4IDJaIiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMTAiIHdpZHRoPSIzIiBoZWlnaHQ9IjUiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iMS41IiB5PSIxMC41IiB3aWR0aD0iMiIgaGVpZ2h0PSI0IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjUiIHk9IjciIHdpZHRoPSIzIiBoZWlnaHQ9IjgiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iNS41IiB5PSI3LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjciIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOSIgeT0iNCIgd2lkdGg9IjMiIGhlaWdodD0iMTEiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iOS41IiB5PSI0LjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEwIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjEzIiB5PSIxIiB3aWR0aD0iMyIgaGVpZ2h0PSIxNCIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxMy41IiB5PSIxLjUiIHdpZHRoPSIyIiBoZWlnaHQ9IjEzIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTE1IDRMMSA0TDggMTNMMTUgNFoiIGZpbGw9IiNGRjExMTEiLz4KPHBhdGggZD0iTTggMTIuMTg1NkwyLjAyMjMyIDQuNUwxMy45Nzc3IDQuNUw4IDEyLjE4NTZaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo=",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgNUgxNVYxMUgxVjVaIiBmaWxsPSIjRkZDRjMzIi8+CjxwYXRoIGQ9Ik0xLjUgNS41SDE0LjVWMTAuNUgxLjVWNS41WiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgMTJIMTVMOCAzTDEgMTJaIiBmaWxsPSIjMkU5OTVGIi8+CjxwYXRoIGQ9Ik04IDMuODE0NDFMMTMuOTc3NyAxMS41SDIuMDIyMzJMOCAzLjgxNDQxWiIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNy41IDIuMTI5NzhMOS4xMzk0OSA1LjQ1MTc1QzkuMjg1MTUgNS43NDY4OSA5LjU2NjcyIDUuOTUxNDYgOS44OTI0MyA1Ljk5ODc5TDEzLjU1ODQgNi41MzE0OUwxMC45MDU3IDkuMTE3MjlDMTAuNjcgOS4zNDcwMiAxMC41NjI1IDkuNjc4MDIgMTAuNjE4MSAxMC4wMDI0TDExLjI0NDMgMTMuNjUzNkw3Ljk2NTM0IDExLjkyOThDNy42NzQwMiAxMS43NzY2IDcuMzI1OTggMTEuNzc2NiA3LjAzNDY2IDExLjkyOThMMy43NTU2OCAxMy42NTM2TDQuMzgxOTEgMTAuMDAyNEM0LjQzNzU0IDkuNjc4MDMgNC4zMyA5LjM0NzAyIDQuMDk0MzEgOS4xMTcyOUwxLjQ0MTU2IDYuNTMxNDlMNS4xMDc1NyA1Ljk5ODc5QzUuNDMzMjggNS45NTE0NiA1LjcxNDg1IDUuNzQ2ODkgNS44NjA1MSA1LjQ1MTc1TDcuNSAyLjEyOTc4WiIgZmlsbD0id2hpdGUiIHN0cm9rZT0iIzczNzM3MyIvPjwvc3ZnPg==",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTggMi4xMjk3OEw2LjM2MDUxIDUuNDUxNzVDNi4yMTQ4NSA1Ljc0Njg5IDUuOTMzMjggNS45NTE0NiA1LjYwNzU3IDUuOTk4NzlMMS45NDE1NiA2LjUzMTQ5TDQuNTk0MzEgOS4xMTcyOUM0LjgzIDkuMzQ3MDIgNC45Mzc1NCA5LjY3ODAyIDQuODgxOTEgMTAuMDAyNEw0LjI1NTY4IDEzLjY1MzZMNy41MzQ2NiAxMS45Mjk4QzcuODI1OTggMTEuNzc2NiA4LjE3NDAyIDExLjc3NjYgOC40NjUzNCAxMS45Mjk4TDExLjc0NDMgMTMuNjUzNkwxMS4xMTgxIDEwLjAwMjRDMTEuMDYyNSA5LjY3ODAzIDExLjE3IDkuMzQ3MDIgMTEuNDA1NyA5LjExNzI5TDE0LjA1ODQgNi41MzE0OUwxMC4zOTI0IDUuOTk4NzlDMTAuMDY2NyA1Ljk1MTQ2IDkuNzg1MTUgNS43NDY4OSA5LjYzOTQ5IDUuNDUxNzVMOCAyLjEyOTc4WiIgZmlsbD0iI0ZGQ0YzMyIgc3Ryb2tlPSIjRERBMTA5Ii8+CjxwYXRoIGQ9Ik04IDEyLjMxNDlDOC4wNzk5MSAxMi4zMTQ5IDguMTU5ODIgMTIuMzM0IDguMjMyNjQgMTIuMzcyM0wxMS41MTE2IDE0LjA5NjJDMTEuODc4NCAxNC4yODkgMTIuMzA3MiAxMy45Nzc2IDEyLjIzNzEgMTMuNTY5MUwxMS42MTA5IDkuOTE3OUMxMS41ODMgOS43NTU3IDExLjYzNjggOS41OTAyIDExLjc1NDcgOS40NzUzM0wxNC40MDc0IDYuODg5NTRDMTQuNzA0MiA2LjYwMDI3IDE0LjU0MDQgNi4wOTYyOCAxNC4xMzAzIDYuMDM2NjlMMTAuNDY0MyA1LjUwMzk5QzEwLjMwMTQgNS40ODAzMiAxMC4xNjA3IDUuMzc4MDQgMTAuMDg3OCA1LjIzMDQ3TDguNDQ4MzQgMS45MDg0OUM4LjM1NjY0IDEuNzIyNjkgOC4xNzgzMiAxLjYyOTc5IDggMS42Mjk3OFYxMi4zMTQ5WiIgZmlsbD0id2hpdGUiLz4KPHBhdGggZD0iTTggMTIuMzE0OUM4LjA3OTkxIDEyLjMxNDkgOC4xNTk4MiAxMi4zMzQgOC4yMzI2NCAxMi4zNzIzTDExLjUxMTYgMTQuMDk2MkMxMS44Nzg0IDE0LjI4OSAxMi4zMDcyIDEzLjk3NzYgMTIuMjM3MSAxMy41NjkxTDExLjYxMDkgOS45MTc5QzExLjU4MyA5Ljc1NTcgMTEuNjM2OCA5LjU5MDIgMTEuNzU0NyA5LjQ3NTMzTDE0LjQwNzQgNi44ODk1NEMxNC43MDQyIDYuNjAwMjcgMTQuNTQwNCA2LjA5NjI4IDE0LjEzMDMgNi4wMzY2OUwxMC40NjQzIDUuNTAzOTlDMTAuMzAxNCA1LjQ4MDMyIDEwLjE2MDcgNS4zNzgwNCAxMC4wODc4IDUuMjMwNDdMOC40NDgzNCAxLjkwODQ5QzguMzU2NjQgMS43MjI2OSA4LjE3ODMyIDEuNjI5NzkgOCAxLjYyOTc4VjMuMjU5NjFMOS4xOTEwOSA1LjY3MzAzQzkuNDA5NTkgNi4xMTU3NSA5LjgzMTk0IDYuNDIyNiAxMC4zMjA1IDYuNDkzNTlMMTIuOTgzOSA2Ljg4MDYxTDExLjA1NjcgOC43NTkyNEMxMC43MDMxIDkuMTAzODUgMTAuNTQxOCA5LjYwMDM1IDEwLjYyNTMgMTAuMDg2OUwxMS4wODAyIDEyLjczOTZMOC42OTc5OCAxMS40ODcyQzguNDc5NSAxMS4zNzIzIDguMjM5NzUgMTEuMzE0OSA4IDExLjMxNDlWMTIuMzE0OVoiIGZpbGw9IiM3MzczNzMiLz4KPHBhdGggZD0iTTggMTIuMzE0OUM4LjA3OTkxIDEyLjMxNDkgOC4xNTk4MiAxMi4zMzQgOC4yMzI2NCAxMi4zNzIzTDExLjUxMTYgMTQuMDk2MkMxMS44Nzg0IDE0LjI4OSAxMi4zMDcyIDEzLjk3NzYgMTIuMjM3MSAxMy41NjkxTDExLjYxMDkgOS45MTc5QzExLjU4MyA5Ljc1NTcgMTEuNjM2OCA5LjU5MDIgMTEuNzU0NyA5LjQ3NTMzTDE0LjQwNzQgNi44ODk1NEMxNC43MDQyIDYuNjAwMjcgMTQuNTQwNCA2LjA5NjI4IDE0LjEzMDMgNi4wMzY2OUwxMC40NjQzIDUuNTAzOTlDMTAuMzAxNCA1LjQ4MDMyIDEwLjE2MDcgNS4zNzgwNCAxMC4wODc4IDUuMjMwNDdMOC40NDgzNCAxLjkwODQ5QzguMzU2NjQgMS43MjI2OSA4LjE3ODMyIDEuNjI5NzkgOCAxLjYyOTc4VjMuMjU5NjFMOS4xOTEwOSA1LjY3MzAzQzkuNDA5NTkgNi4xMTU3NSA5LjgzMTk0IDYuNDIyNiAxMC4zMjA1IDYuNDkzNTlMMTIuOTgzOSA2Ljg4MDYxTDExLjA1NjcgOC43NTkyNEMxMC43MDMxIDkuMTAzODUgMTAuNTQxOCA5LjYwMDM1IDEwLjYyNTMgMTAuMDg2OUwxMS4wODAyIDEyLjczOTZMOC42OTc5OCAxMS40ODcyQzguNDc5NSAxMS4zNzIzIDguMjM5NzUgMTEuMzE0OSA4IDExLjMxNDlWMTIuMzE0OVoiIGZpbGw9IiM3MzczNzMiLz4KPC9zdmc+",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNNy41IDIuMTI5NzhMOS4xMzk0OSA1LjQ1MTc1QzkuMjg1MTUgNS43NDY4OSA5LjU2NjcyIDUuOTUxNDYgOS44OTI0MyA1Ljk5ODc5TDEzLjU1ODQgNi41MzE0OUwxMC45MDU3IDkuMTE3MjlDMTAuNjcgOS4zNDcwMiAxMC41NjI1IDkuNjc4MDIgMTAuNjE4MSAxMC4wMDI0TDExLjI0NDMgMTMuNjUzNkw3Ljk2NTM0IDExLjkyOThDNy42NzQwMiAxMS43NzY2IDcuMzI1OTggMTEuNzc2NiA3LjAzNDY2IDExLjkyOThMMy43NTU2OCAxMy42NTM2TDQuMzgxOTEgMTAuMDAyNEM0LjQzNzU0IDkuNjc4MDMgNC4zMyA5LjM0NzAyIDQuMDk0MzEgOS4xMTcyOUwxLjQ0MTU2IDYuNTMxNDlMNS4xMDc1NyA1Ljk5ODc5QzUuNDMzMjggNS45NTE0NiA1LjcxNDg1IDUuNzQ2ODkgNS44NjA1MSA1LjQ1MTc1TDcuNSAyLjEyOTc4WiIgZmlsbD0iI0ZGQ0YzMyIgc3Ryb2tlPSIjRERBMTA5Ii8+PC9zdmc+",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxLjUiIHk9IjIuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cmVjdCB4PSI4IiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+CjxyZWN0IHg9IjguNSIgeT0iMi41IiB3aWR0aD0iNSIgaGVpZ2h0PSI1IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iOC41IiB5PSI5LjUiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iMSIgeT0iOSIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxLjUiIHk9IjkuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxLjUiIHk9IjIuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cmVjdCB4PSI4IiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+CjxyZWN0IHg9IjguNSIgeT0iMi41IiB3aWR0aD0iNSIgaGVpZ2h0PSI1IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiNDQ0NDQ0MiLz4KPHJlY3QgeD0iOC41IiB5PSI5LjUiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iMSIgeT0iOSIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSIxLjUiIHk9IjkuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iI0NDQ0NDQyIvPgo8cmVjdCB4PSIxLjUiIHk9IjIuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cmVjdCB4PSI4IiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+CjxyZWN0IHg9IjguNSIgeT0iMi41IiB3aWR0aD0iNSIgaGVpZ2h0PSI1IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iOC41IiB5PSI5LjUiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iMSIgeT0iOSIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSIxLjUiIHk9IjkuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3QgeD0iMSIgeT0iMiIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSIxLjUiIHk9IjIuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cmVjdCB4PSI4IiB5PSIyIiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjQ0NDQ0NDIi8+CjxyZWN0IHg9IjguNSIgeT0iMi41IiB3aWR0aD0iNSIgaGVpZ2h0PSI1IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjgiIHk9IjkiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iOC41IiB5PSI5LjUiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iMSIgeT0iOSIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSIxLjUiIHk9IjkuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8L3N2Zz4K",
		"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgMkg3VjhIMVYyWiIgZmlsbD0iIzIzNjFCRSIvPgo8cGF0aCBkPSJNMS41IDIuNUg2LjVWNy41SDEuNVYyLjVaIiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+CjxyZWN0IHg9IjgiIHk9IjIiIHdpZHRoPSI2IiBoZWlnaHQ9IjYiIGZpbGw9IiMyMzYxQkUiLz4KPHJlY3QgeD0iOC41IiB5PSIyLjUiIHdpZHRoPSI1IiBoZWlnaHQ9IjUiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjIiLz4KPHJlY3QgeD0iOCIgeT0iOSIgd2lkdGg9IjYiIGhlaWdodD0iNiIgZmlsbD0iIzIzNjFCRSIvPgo8cmVjdCB4PSI4LjUiIHk9IjkuNSIgd2lkdGg9IjUiIGhlaWdodD0iNSIgc3Ryb2tlPSJibGFjayIgc3Ryb2tlLW9wYWNpdHk9IjAuMiIvPgo8cmVjdCB4PSIxIiB5PSI5IiB3aWR0aD0iNiIgaGVpZ2h0PSI2IiBmaWxsPSIjMjM2MUJFIi8+CjxyZWN0IHg9IjEuNSIgeT0iOS41IiB3aWR0aD0iNSIgaGVpZ2h0PSI1IiBzdHJva2U9ImJsYWNrIiBzdHJva2Utb3BhY2l0eT0iMC4yIi8+Cjwvc3ZnPgo="];


	var iDown = 1, iSide = 2, iUp = 3;
	var iDownGray = 4, iSideGray = 5, iUpGray = 6;
	var iFlagRed = 7, iFlagYellow = 8, iFlagGreen = 9;
	var iDiamondRed = 10, iTriangleYellow = 11, iCircleGreen = 12;
	var iCrossRed = 13, iExclamationYellow = 14, iCheckGreen = 15;
	var iCrossSymbolRed = 16, iExclamationSymbolYellow = 17, iCheckSymbolGreen = 18;
	var iCircleRed = 19, iCircleYellow = 20;
	var iTrafficLightRed = 21, iTrafficLightYellow = 22, iTrafficLightGreen = 23;
	var iDownIncline = 24, iUpIncline = 25;
	var iDownInclineGray = 26, iUpInclineGray = 27;
	var iOneFilledBars = 28, iTwoFilledBars = 29, iThreeFilledBars = 30, iFourFilledBars = 31;
	var iCircleBlack = 32, iCircleGray = 33, iCircleLightRed = 34;
	var iCircleWhite = 35, iCircleThreeWhiteQuarters = 36, iCircleTwoWhiteQuarters = 37, iCircleOneWhiteQuarter = 38;
	var iZeroFilledBars = 39;
	var iTriangleRed = 40, iDashYellow = 41, iTriangleGreen = 42;
	var iStarSilver = 43, iStarHalf = 44, iStarGold = 45;
	var iZeroFilledBoxes = 46, iOneFilledBoxes = 47, iTwoFilledBoxes = 48, iThreeFilledBoxes = 49, iFourFilledBoxes = 50;

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
		return fullIconArray;
	}

	function getCFIcon(oRuleElement, index) {
		var oIconSet = oRuleElement.aIconSets && oRuleElement.aIconSets[index];
		var iconSetType = (oIconSet && null !== oIconSet.IconSet) ? oIconSet.IconSet : oRuleElement.IconSet;
		if (EIconSetType.NoIcons === iconSetType) {
			return null;
		}
		var icons = c_arrIcons[iconSetType] || c_arrIcons[EIconSetType.Traffic3Lights1];
		return fullIconArray[(icons[(oIconSet && null !== oIconSet.IconId) ? oIconSet.IconId : index] || icons[icons.length - 1]) - 1];
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

	function getFullCFIcons() {
		return fullIconArray;
	}
	function getCFIconsByType() {
		return c_arrIcons;
	}
	function getFullCFPresets() {
		return conditionalFormattingPresets;
	}

	//[AxisColor, BorderColor, Color, Gradient, NegativeBorderColor, NegativeColor]
	var aDataBarStyles = [[0, 6524614, 6524614, true, 16711680, 16711680],
		[0, 6538116, 6538116, true, 16711680, 16711680], [0, 16733530, 16733530, true, 16711680, 16711680],
		[0, 16758312, 16758312, true, 16711680, 16711680], [0, 35567, 35567, true, 16711680, 16711680],
		[0, 14024827, 14024827, true, 16711680, 16711680], [0, , 6524614, false, , 16711680],
		[0, , 6538116, false, , 16711680], [0, , 16733530, false, , 16711680], [0, , 16758312, false, , 16711680],
		[0, , 35567, false, , 16711680], [0, , 14024827, false, , 16711680]];
	//[[Type, Val, Rgb], [Type, Val, Rgb], ...]
	var aColorScaleStyles = [[[2, null, 16279915], [5, 50, 16771972], [1, null, 6536827]],
		[[2, null, 6536827], [5, 50, 16771972], [1, null, 16279915]],
		[[2, null, 16279915], [5, 50, 16579839], [1, null, 6536827]],
		[[2, null, 6536827], [5, 50, 16579839], [1, null, 16279915]],
		[[2, null, 16279915], [5, 50, 16579839], [1, null, 5933766]],
		[[2, null, 5933766], [5, 50, 16579839], [1, null, 16279915]], [[2, null, 16279915], [1, null, 16579839]],
		[[2, null, 16579839], [1, null, 16279915]], [[2, null, 16579839], [1, null, 6536827]],
		[[2, null, 6536827], [1, null, 16579839]], [[2, null, 16773020], [1, null, 6536827]],
		[[2, null, 6536827], [1, null, 16773020]]];
	//[[[Type, Val, Formula], [Type, Val, Formula], ...], [[Type, Val, Formula], [Type, Val, Formula], ...]]
	var aIconsStyles = [[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '25', null], [4, '50', null], [4, '75', '75']],
		[[4, '0', null], [4, '25', null], [4, '50', null], [4, '75', '75']],
		[[4, '0', null], [4, '25', null], [4, '50', null], [4, '75', '75']],
		[[4, '0', null], [4, '25', null], [4, '50', null], [4, '75', '75']],
		[[4, '0', null], [4, '25', null], [4, '50', null], [4, '75', '75']],
		[[4, '0', null], [4, '20', null], [4, '40', null], [4, '60', null], [4, '80', '80']],
		[[4, '0', null], [4, '20', null], [4, '40', null], [4, '60', null], [4, '80', '80']],
		[[4, '0', null], [4, '20', null], [4, '40', null], [4, '60', null], [4, '80', '80']],
		[[4, '0', null], [4, '20', null], [4, '40', null], [4, '60', null], [4, '80', '80']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '33', null], [4, '67', '67']],
		[[4, '0', null], [4, '20', null], [4, '40', null], [4, '60', null], [4, '80', '80']]];

	//[[fontColor, fillColor, borderColor]]
	var aFormatStyles = [["9C0006", "FFC7CE"], ["9C5700", "FFEB9C"], ["006100", "C6EFCE"], [, "FFC7CE"], ["9C0006"],
		[, , "9C0006"]];

	var conditionalFormattingPresets = {};
	conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.dataBar] = aDataBarStyles;
	conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.colorScale] = aColorScaleStyles;
	conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.icons] = aIconsStyles;
	conditionalFormattingPresets[Asc.c_oAscCFRuleTypeSettings.format] = aFormatStyles;

	/*
	 * Export
	 * -----------------------------------------------------------------------------
	 */
	var prot;
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CConditionalFormatting = CConditionalFormatting;
	window['AscCommonExcel'].CConditionalFormattingFormulaParent = CConditionalFormattingFormulaParent;
	window['Asc']["asc_CConditionalFormattingRule"] = window['AscCommonExcel'].CConditionalFormattingRule = CConditionalFormattingRule;
	window['Asc']["asc_CColorScale"] = window['AscCommonExcel'].CColorScale = CColorScale;
	window['Asc']["asc_CDataBar"] = window['AscCommonExcel'].CDataBar = CDataBar;
	window['AscCommonExcel'].CFormulaCF = CFormulaCF;
	window['Asc']["asc_CIconSet"] = window['AscCommonExcel'].CIconSet = CIconSet;
	window['Asc']["asc_CConditionalFormatValueObject"] = window['AscCommonExcel'].CConditionalFormatValueObject = CConditionalFormatValueObject;
	window['Asc']["asc_CConditionalFormatIconSet"] = window['AscCommonExcel'].CConditionalFormatIconSet = CConditionalFormatIconSet;
	window['AscCommonExcel'].CGradient = CGradient;

	window['AscCommonExcel'].cDefIconSize = cDefIconSize;
	window['AscCommonExcel'].cDefIconFont = cDefIconFont;
	window['AscCommonExcel'].getCFIconsForLoad = getCFIconsForLoad;
	window['AscCommonExcel'].getCFIcon = getCFIcon;
	window['AscCommonExcel'].getDataBarGradientColor = getDataBarGradientColor;
	window['AscCommonExcel'].getFullCFIcons = getFullCFIcons;
	window['AscCommonExcel'].getFullCFPresets = getFullCFPresets;
	window['AscCommonExcel'].getCFIconsByType = getCFIconsByType;
	window['AscCommonExcel'].isValidDataRefCf = isValidDataRefCf;

	prot = CConditionalFormattingRule.prototype;
	prot['asc_getDxf'] = prot.asc_getDxf;
	prot['asc_getType'] = prot.asc_getType;
	prot['asc_getLocation'] = prot.asc_getLocation;
	prot['asc_getContainsText'] = prot.asc_getContainsText;
	prot['asc_getTimePeriod'] = prot.asc_getTimePeriod;
	prot['asc_getOperator'] = prot.asc_getOperator;
	prot['asc_getPriority'] = prot.asc_getPriority;
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
	prot['asc_setPriority'] = prot.asc_setPriority;
	prot['asc_setRank'] = prot.asc_setRank;
	prot['asc_setBottom'] = prot.asc_setBottom;
	prot['asc_setPercent'] = prot.asc_setPercent;
	prot['asc_setAboveAverage'] = prot.asc_setAboveAverage;
	prot['asc_setEqualAverage'] = prot.asc_setEqualAverage;
	prot['asc_setStdDev'] = prot.asc_setStdDev;
	prot['asc_setValue1'] = prot.asc_setValue1;
	prot['asc_setValue2'] = prot.asc_setValue2;
	prot['asc_checkScope'] = prot.asc_checkScope;
	prot['asc_setColorScaleOrDataBarOrIconSetRule'] = prot.asc_setColorScaleOrDataBarOrIconSetRule;


	prot = CColorScale.prototype;
	prot['asc_getCFVOs'] = prot.asc_getCFVOs;
	prot['asc_getColors'] = prot.asc_getColors;
	prot['asc_getPreview'] = prot.asc_getPreview;
	prot['asc_setCFVOs'] = prot.asc_setCFVOs;
	prot['asc_setColors'] = prot.asc_setColors;

	prot = CDataBar.prototype;
	prot['asc_setInterfaceDefault'] = prot.asc_setInterfaceDefault;
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

	prot = CIconSet.prototype;
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

	prot = CConditionalFormatValueObject.prototype;
	prot['asc_getGte'] = prot.asc_getGte;
	prot['asc_getType'] = prot.asc_getType;
	prot['asc_getVal'] = prot.asc_getVal;
	prot['asc_setGte'] = prot.asc_setGte;
	prot['asc_setType'] = prot.asc_setType;
	prot['asc_setVal'] = prot.asc_setVal;

	prot = CFormulaCF.prototype;
	prot['asc_getText'] = prot.asc_getText;
	prot['asc_setText'] = prot.asc_setText;

	prot = CConditionalFormatIconSet.prototype;
	prot['asc_getIconSet'] = prot.asc_getIconSet;
	prot['asc_getIconId'] = prot.asc_getIconId;
	prot['asc_setIconSet'] = prot.asc_setIconSet;
	prot['asc_setIconId'] = prot.asc_setIconId;

})(window);
