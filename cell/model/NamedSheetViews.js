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

(function (undefined) {

	var prot;
	var asc = window["Asc"];
	var c_oAscError = asc.c_oAscError;
	var History = AscCommon.History;


	// 1. интерфейс
	// asc_addNamedSheetView - создание нового вью и дублирование текущего.
	// лочим созднанный лист, проверяем лок менеджера. при принятии лока другими пользователями - лочится менеджер.
	// для добавления в историю используем historyitem_Worksheet_SheetViewAdd, в историю кладём весь объект.

	// asc_getNamedSheetViews - отдаём массив отображений активного листа
	// asc_getActiveNamedSheetView - отдаём имя активного листа

	// asc_deleteNamedSheetViews - удаление массива отображений. лочим удаляемое отображение.
	// в данном случае(как и в именованных диапазонах) у других пользователей нельзя добавить новое отображение. Удалить другие можно.
	// для истории при удалении использую UndoRedoData_NamedSheetViewRedo, потому что весь объект необходим только при undo, пересылать его не нужно.

	// asc_setActiveNamedSheetView - выставление активного отображения. ничего не лочим. внутри функции описана процедура взаимодействия со скрытыми строками при переходе между вью.

	// 1.1 Ограничения строгого совместного редактирования:
	// 	- Локи работают следующим образом: при переходе между вью локов нет. при примении а/ф в режиме вью ничего не лочится.
	// - при взаимных изменениях с одним а/ф, применяем тот а/ф, который был последним сохраненным.
	// - при скрытии строк в режиме дефолт - лочится лист, но в режиме вью можно использовать а/ф для скрытия строчек, скрывать строки через контекстное меню после скрытия строк в дефолте - нельзя.
	// - при добавлении нового вью - лочим менеджер
	// - при удалении вью - лочим менеджер. но при удалении не проверяем залочен ли менеджер. проверяем только залоченность конкретного вью.
	// - при переименовании - лочим менеджер

	// 2. служебные функции в приватном апи
	// _isLockedNamedSheetView - проверка лока массива отображений
	// _onUpdateNamedSheetViewLock - вызывается из onLocksAcquired, добавляем информацию о локах
	// _onUpdateAllSheetViewLock - снимаем локи со всех листов и отображний. !!! вызывается через "updateAllSheetViewLock"(пересмотреть). возможно, необходимо добавить вызов в onLocksReleased !!!
	// 	isNamedSheetViewManagerLocked - проверка лока листа. !!! храним в прототипе апи - sheetViewManagerLocks. пересмореть!!!
	// 	updateAllFilters


	function CT_NamedSheetViews() {
		this.namedSheetView = [];
		// this.extLst = null;
		return this;
	}

	CT_NamedSheetViews.prototype.toStream = function (s, tableIds) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.StartRecord(0);
		var len = this.namedSheetView.length;
		s.WriteULong(len);
		for (var i = 0; i < len; i++) {
			s.StartRecord(0);
			this.namedSheetView[i].toStream(s, tableIds);
			s.EndRecord();
		}
		s.EndRecord();
	};
	CT_NamedSheetViews.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_NamedSheetView();
						tmp.fromStream(s, wb);
						this.namedSheetView.push(tmp);
					}
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};
	CT_NamedSheetViews.prototype.hasColorFilter = function(){
		return this.namedSheetView.some(function(namedSheetView) {
			return namedSheetView.hasColorFilter();
		});
	}

	function CT_NamedSheetView() {
		this.nsvFilters = [];
		//this.extLst
		this.name = null;
		this.id = AscCommon.CreateGUID();

		this.ws = null;
		this.isLock = null;
		this._isActive = null;

		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		return this;
	}

	CT_NamedSheetView.prototype.Get_Id = function () {
		return this.Id;
	};

	CT_NamedSheetView.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_NamedSheetView;
	};

	CT_NamedSheetView.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.NamedSheetView;
	};

	CT_NamedSheetView.prototype.toStream = function (s, tableIds) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.name);
		s._WriteString2(1, this.id);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.StartRecord(0);
		var len = this.nsvFilters.length;
		s.WriteULong(len);
		for (var i = 0; i < len; i++) {
			s.StartRecord(0);
			this.nsvFilters[i].toStream(s, tableIds);
			s.EndRecord();
		}
		s.EndRecord();
	};
	CT_NamedSheetView.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				case 0: {
					this.name = s.GetString2();
					break;
				}
				case 1: {
					this.id = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_NsvFilter();
						tmp.fromStream(s, wb);
						this.nsvFilters.push(tmp);
					}
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};
	CT_NamedSheetView.prototype.Read_FromBinary2 = function (reader) {
		var length = reader.GetLong();
		for (var i = 0; i < length; ++i) {
			var _filter = new CT_NsvFilter();
			_filter.Read_FromBinary2(reader)
			this.nsvFilters.push(_filter);
		}

		this.name = reader.GetString2();
		this.id = reader.GetString2();
	};
	CT_NamedSheetView.prototype.initPostOpen = function (tableIds, ws) {
		for (var i = 0; i < this.nsvFilters.length; ++i) {
			this.nsvFilters[i].initPostOpen(tableIds);
		}
		if (!this.ws) {
			this.setWS(ws);
		}
	};
	CT_NamedSheetView.prototype.clone = function (tableNameMap) {
		var res = new CT_NamedSheetView(true);

		for (var i = 0; i < this.nsvFilters.length; ++i) {
			res.nsvFilters[i] = this.nsvFilters[i].clone(tableNameMap);
		}

		res.name = this.name;
		res.ws = this.ws;

		return res;
	};

	CT_NamedSheetView.prototype.setWS = function (ws) {
		this.ws = ws;
	};
	CT_NamedSheetView.prototype.hasColorFilter = function(){
		return this.nsvFilters.some(function(nsvFilter) {
			return nsvFilter.hasColorFilter();
		});
	}
	CT_NamedSheetView.prototype.asc_getName = function () {
		return this.name;
	};

	CT_NamedSheetView.prototype.asc_setName = function (val) {
		var t = this;
		var api = window["Asc"]["editor"];
		if (this.name !== val) {
			if (api.isNamedSheetViewManagerLocked(t.ws.getId())) {
				t.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.LockedEditView, c_oAscError.Level.NoCritical);
				return;
			}

			api._isLockedNamedSheetView([t], function(success) {
				if (!success) {
					t.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.LockedEditView, c_oAscError.Level.NoCritical);
					return;
				}

				History.Create_NewPoint();
				History.StartTransaction();

				var oldVal = t.name;
				t.setName(val);

				History.Add(AscCommonExcel.g_oUndoRedoNamedSheetViews, AscCH.historyitem_NamedSheetView_SetName,
					t.ws.getId(), null, new AscCommonExcel.UndoRedoData_NamedSheetView(t.Id, oldVal, val));

				History.EndTransaction();

				api.handlers.trigger("asc_onRefreshNamedSheetViewList", t.ws.index);
			});
		}
	};

	CT_NamedSheetView.prototype.setName = function (val) {
		this.name = val;
	};


	CT_NamedSheetView.prototype.asc_getIsActive = function () {
		return this._isActive;
	};

	CT_NamedSheetView.prototype.generateName = function () {
		var ws = this.ws;
		if (!ws) {
			return;
		}

		var mapNames = [], isContains, name = this.name;
		for (var i = 0; i < ws.aNamedSheetViews.length; i++) {
			if (name && name === ws.aNamedSheetViews[i].name) {
				isContains = true;
			}
			mapNames[ws.aNamedSheetViews[i].name] = 1;
		}

		var baseName, counter;
		if (!name) {
			//TODO перевод
			name = "View";

			baseName = name;
			counter = 1;
			while (mapNames[baseName + counter]) {
				counter++;
			}
			name = baseName + counter;
		} else if (isContains) {
			//так делаяем при создании дубликата
			baseName = name + " ";
			counter = 2;
			while (mapNames[baseName + "(" + counter + ")"]) {
				counter++;
			}
			name = baseName + "(" + counter + ")";
		}

		return name;
	};

	CT_NamedSheetView.prototype.asc_getIsLock = function () {
		return this.isLock;
	};

	CT_NamedSheetView.prototype.addFilter = function (filter) {
		var nsvFilter = new CT_NsvFilter();
		nsvFilter.init(filter);
		this.nsvFilters.push(nsvFilter);
		//TODO history

	};

	CT_NamedSheetView.prototype.deleteFilter = function (filter) {
		if (!this.nsvFilters || !this.nsvFilters.length || !filter) {
			return;
		}

		for (var i = 0; i < this.nsvFilters.length; i++) {
			var isAutoFilter = filter.isAutoFilter();
			var isDelete = false;
			if (isAutoFilter && this.nsvFilters[i].tableId === "0") {
				isDelete = true;
			} else if (!isAutoFilter && this.nsvFilters[i].tableId === filter.DisplayName) {
				isDelete = true;
			}

			if (isDelete) {
				var historyFilter = this.nsvFilters[i].clone();
				this.nsvFilters.splice(i, 1);
				History.Add(AscCommonExcel.g_oUndoRedoNamedSheetViews, AscCH.historyitem_NamedSheetView_DeleteFilter,
					this.ws.getId(), null, new AscCommonExcel.UndoRedoData_NamedSheetViewRedo(this.Id, historyFilter, null));
				break;
			}
		}
	};

	CT_NamedSheetView.prototype.getNsvFiltersByTableId = function (val) {
		if (!this.nsvFilters) {
			return null;
		}
		if (!val) {
			val = "0";
		}
		for (var i = 0; i < this.nsvFilters.length; i++) {
			if (this.nsvFilters[i].tableId === val) {
				return this.nsvFilters[i];
			}
		}
		return null;
	};

	CT_NamedSheetView.prototype.Write_ToBinary2 = function (writer) {
		//for wrapper
		writer.WriteLong(this.getObjectType());

		writer.WriteLong(this.nsvFilters ? this.nsvFilters.length : 0);

		if (this.nsvFilters) {
			for(var i = 0, length = this.nsvFilters.length; i < length; ++i) {
				this.nsvFilters[i].Write_ToBinary2(writer);
			}
		}

		writer.WriteString2(this.name);
		writer.WriteString2(this.id);
	};

	function CT_NsvFilter() {
		this.columnsFilter = [];
		this.sortRules = null;
		//this.extLst
		this.filterId = AscCommon.CreateGUID();
		this.ref = null;
		this.tableId = null;
		this.tableIdOpen = null;

		return this;
	}

	CT_NsvFilter.prototype.toStream = function (s, tableIds) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.filterId);
		if (null !== this.ref) {
			s._WriteString2(1, this.ref.getName(AscCommonExcel.referenceType.R));
		}
		if ("0" === this.tableId) {
			s._WriteUInt2(2, 0);
		} else {
			var elem = tableIds && tableIds[this.tableId];
			if (elem) {
				s._WriteUInt2(2, elem.id);
			}
		}
		s._WriteUInt2(2, this.tableIdOpen);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.columnsFilter);
		if (null !== this.sortRules) {
			var sortRules = new CT_SortRules();
			sortRules.sortRule = this.sortRules;
			s.WriteRecord4(1, sortRules);
		}
	};
	CT_NsvFilter.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				case 0: {
					this.filterId = s.GetString2();
					break;
				}
				case 1: {
					this.ref = AscCommonExcel.g_oRangeCache.getAscRange(s.GetString2());
					break;
				}
				case 2: {
					this.tableIdOpen = s.GetULong();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_ColumnFilter();
						tmp.fromStream(s, wb);
						this.columnsFilter.push(tmp);
					}
					break;
				}
				case 1: {
					var sortRules = new CT_SortRules();
					sortRules.fromStream(s, wb);
					this.sortRules = sortRules.sortRule;
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};
	CT_NsvFilter.prototype.Write_ToBinary2 = function (writer) {
		writer.WriteLong(this.columnsFilter ? this.columnsFilter.length : 0);

		var i, length;
		if (this.columnsFilter) {
			for(i = 0, length = this.columnsFilter.length; i < length; ++i) {
				this.columnsFilter[i].Write_ToBinary2(writer);
			}
		}

		writer.WriteLong(this.sortRules ? this.sortRules.length : 0);

		if (this.sortRules) {
			for(i = 0, length = this.sortRules.length; i < length; ++i) {
				this.sortRules[i].Write_ToBinary2(writer);
			}
		}

		writer.WriteString2(this.filterId);

		if (null != this.Ref) {
			writer.WriteBool(true);
			writer.WriteLong(this.Ref.r1);
			writer.WriteLong(this.Ref.c1);
			writer.WriteLong(this.Ref.r2);
			writer.WriteLong(this.Ref.c2);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.tableId) {
			writer.WriteBool(true);
			writer.WriteString2(this.tableId);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.tableIdOpen) {
			writer.WriteBool(true);
			writer.WriteString2(this.tableIdOpen);
		} else {
			writer.WriteBool(false);
		}
	};
	CT_NsvFilter.prototype.Read_FromBinary2 = function (reader) {
		var i, obj;
		var length = reader.GetLong();
		for (i = 0; i < length; ++i) {
			_obj = new CT_ColumnFilter();
			_obj.Read_FromBinary2(reader)
			this.columnsFilter.push(_obj);
		}

		length = reader.GetLong();
		for (i = 0; i < length; ++i) {
			var _obj = new CT_SortRule();
			if (!this.sortRules) {
				this.sortRules = [];
			}
			_obj.Read_FromBinary2(reader)
			this.sortRules.push(_obj);
		}

		this.filterId = reader.GetString2();

		if (reader.GetBool()) {
			var r1 = reader.GetLong();
			var c1 = reader.GetLong();
			var r2 = reader.GetLong();
			var c2 = reader.GetLong();

			this.ref = new Asc.Range(c1, r1, c2, r2);
		}

		if (reader.GetBool()) {
			this.tableId = reader.GetString2();
		}

		if (reader.GetBool()) {
			this.tableIdOpen = reader.GetString2();
		}
	};
	CT_NsvFilter.prototype.clone = function (tableNameMap) {
		var res = new CT_NsvFilter();
		var i;
		if (this.columnsFilter) {
			for (i = 0; i < this.columnsFilter.length; ++i) {
				res.columnsFilter[i] = this.columnsFilter[i].clone();
			}
		}
		if (this.sortRules) {
			for (i = 0; i < this.sortRules.length; ++i) {
				res.sortRules[i] = this.sortRules[i].clone();
			}
		}

		//res.filterId = this.filterId;
		res.ref = this.ref;
		res.tableId = tableNameMap && tableNameMap[this.tableId] ? tableNameMap[this.tableId] : this.tableId;
		res.tableIdOpen = this.tableIdOpen;

		return res;
	};

	CT_NsvFilter.prototype.initPostOpen = function (tableIds) {
		var table = null;
		if (null != this.tableIdOpen) {
			if (0 !== this.tableIdOpen) {
				table = tableIds[this.tableIdOpen];
				if (table) {
					this.tableId = table.DisplayName;
				}
			} else {
				this.tableId = "0";
			}
			this.tableIdOpen = null;
		}
		return table;
	};

	CT_NsvFilter.prototype.getColumnFilterByColId = function (id, isGetIndex) {
		for (var i = 0; i < this.columnsFilter.length; ++i) {
			if (this.columnsFilter[i].filter && this.columnsFilter[i].filter.ColId === id) {
				return !isGetIndex ? this.columnsFilter[i].filter : {filter: this.columnsFilter[i].filter, index: i};
			}
		}
		return null;
	};

	CT_NsvFilter.prototype.init = function (obj) {
		if (obj) {
			var af;
			if (!obj.isAutoFilter()) {
				this.ref = obj.Ref;
				this.tableId = obj.DisplayName;
				af = obj.AutoFilter;
			} else {
				this.ref = obj.Ref;
				this.tableId = "0";
				af = obj;
			}

			if (af && af.FilterColumns) {
				for (var i = 0; i < af.FilterColumns.length; i++) {
					var newColumnFilter = new CT_ColumnFilter();
					newColumnFilter.colId = af.FilterColumns[i].ColId;
					newColumnFilter.filter = af.FilterColumns[i].clone();
					this.columnsFilter.push(newColumnFilter);
				}
			}
		}
	};

	CT_NsvFilter.prototype.isApplyAutoFilter = function (obj) {
		var res = null
		if (this.columnsFilter && this.columnsFilter.length) {
			for (var i = 0; i < this.columnsFilter.length; i++) {
				var _filterColumn = this.columnsFilter[i] && this.columnsFilter[i].filter;
				if (_filterColumn.isApplyAutoFilter()) {
					res = true;
					break;
				}
			}
		}
		return res;
	};
	CT_NsvFilter.prototype.getFilterColumnByIndex = function (index) {
		return this.columnsFilter && this.columnsFilter[index] && this.columnsFilter[index].filter;
	};
	CT_NsvFilter.prototype.getIndexByColId = function (colId) {
		var res = null;

		if (!this.columnsFilter) {
			return res;
		}

		for (var i = 0; i < this.columnsFilter.length; i++) {
			if (this.columnsFilter[i].filter && this.columnsFilter[i].filter.ColId === colId) {
				res = i;
				break;
			}
		}

		return res;
	};
	CT_NsvFilter.prototype.deleteFilterColumn = function(index) {
		if (this.columnsFilter && this.columnsFilter[index]) {
			this.columnsFilter.splice(index, 1)
		}
	};
	CT_NsvFilter.prototype.hasColorFilter = function(){
		return this.columnsFilter.some(function(columnsFilter) {
			return columnsFilter.hasColorFilter();
		});
	}


	function CT_ColumnFilter() {
		this.filter = null;
		//this.extLst

		//нужно ли?
		this.colId = null;
		this.id = AscCommon.CreateGUID();

		return this;
	}

	CT_ColumnFilter.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.colId);
		s._WriteString2(1, this.id);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		var dxfs = [];
		if (null !== this.filter) {
			s.StartRecord(1);
			s.WriteULong(1);
			s.StartRecord(0);
			var tmp = new AscCommon.CMemory(true);
			s.ExportToMemory(tmp);
			var btw = new AscCommonExcel.BinaryTableWriter(tmp, dxfs, false, {});
			btw.WriteFilterColumn(this.filter);
			s.ImportFromMemory(tmp);
			s.EndRecord();
			s.EndRecord();
		}
		if (dxfs.length > 0) {
			s.StartRecord(0);
			var tmp = new AscCommon.CMemory(true);
			s.ExportToMemory(tmp);
			var bstw = new AscCommonExcel.BinaryStylesTableWriter(tmp, null, null);
			bstw.WriteDxf(dxfs[0]);
			s.ImportFromMemory(tmp);
			s.EndRecord();
		}
	};
	CT_ColumnFilter.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				case 0: {
					this.colId = s.GetULong();
					break;
				}
				case 1: {
					this.id = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var dxfs = [];
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					var _stream = new AscCommon.FT_Stream2();
					_stream.FromFileStream(s);
					var bsr = new AscCommonExcel.Binary_StylesTableReader(_stream, wb);
					dxfs.push(bsr.ReadDxfExternal());
					_stream.ToFileStream2(s);
					break;
				}
				case 1: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var _stream = new AscCommon.FT_Stream2();
						_stream.FromFileStream(s);
						var oReadResult = new AscCommonWord.DocReadResult(null);
						var bwtr = new AscCommonExcel.Binary_TableReader(_stream, oReadResult, null, dxfs);
						this.filter = bwtr.ReadFilterColumnExternal();
						_stream.ToFileStream2(s);
					}
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};
	CT_ColumnFilter.prototype.Write_ToBinary2 = function (writer) {
		if(null != this.dxf) {
			var dxf = this.dxf;
			writer.WriteBool(true);
			var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(writer);
			oBinaryStylesTableWriter.bs.WriteItem(0, function(){oBinaryStylesTableWriter.WriteDxf(dxf);});
		}else {
			writer.WriteBool(false);
		}

		if(null != this.filter) {
			writer.WriteBool(true);
			this.filter.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
		//?
		/*	this.colId = null;
		this.id = null;*/
	};
	CT_ColumnFilter.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var bsr = new AscCommonExcel.Binary_StylesTableReader(reader, wb);
			var bcr = new AscCommon.Binary_CommonReader(reader);
			var oDxf = new AscCommonExcel.CellXfs();
			reader.GetUChar();
			var length = reader.GetULongLE();
			bcr.Read1(length, function (t, l) {
				return bsr.ReadDxf(t, l, oDxf);
			});
			this.dxf = oDxf;
		}
		if (reader.GetBool()) {
			var obj = new window['AscCommonExcel'].FilterColumn();
			obj.Read_FromBinary2(reader);
			this.colId = obj ? obj.ColId : null;
			this.filter = obj;
		}
	};
	CT_ColumnFilter.prototype.clone = function () {
		var res = new CT_ColumnFilter();
		res.filter = this.filter ? this.filter.clone() : null;
		res.colId = this.colId;

		return res;
	};
	CT_ColumnFilter.prototype.hasColorFilter = function(){
		return null !== this.filter && this.filter.isColorFilter();
	};

	function CT_SortRules() {
		this.sortMethod = null;//none
		this.caseSensitive = null;//False
		this.sortRule = [];
		// this.extLst = null;
		return this;
	}

	CT_SortRules.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUChar2(0, this.sortMethod);
		s._WriteBool2(1, this.caseSensitive);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.sortRule);
	};
	CT_SortRules.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				case 0: {
					this.sortMethod = s.GetUChar();
					break;
				}
				case 1: {
					this.caseSensitive = s.GetBool();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_SortRule();
						tmp.fromStream(s, wb);
						this.sortRule.push(tmp);
					}
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};

	function CT_SortRule() {
		this.sortCondition = null;
		this.richSortCondition = null;

		//нужно ли?
		this.colId = null;
		this.id = AscCommon.CreateGUID();

		return this;
	}

	CT_SortRule.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.colId);
		s._WriteString2(1, this.id);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		// s.WriteRecord4(1, this.richSortCondition);
		var dxfs = [];
		if (null !== this.sortCondition) {
			s.StartRecord(2);
			var tmp = new AscCommon.CMemory(true);
			s.ExportToMemory(tmp);
			var btw = new AscCommonExcel.BinaryTableWriter(tmp, dxfs, false, {});
			//dxfId is absent in sortCondition
			if (this.sortCondition.dxf) {
				dxfs.push(this.sortCondition.dxf);
				this.sortCondition.dxf = null;
			}
			btw.WriteSortCondition(this.sortCondition);
			if (dxfs.length > 0) {
				this.sortCondition.dxf = dxfs[0];
			}
			s.ImportFromMemory(tmp);
			s.EndRecord();
		}
		if (dxfs.length > 0) {
			s.StartRecord(0);
			var tmp = new AscCommon.CMemory(true);
			s.ExportToMemory(tmp);
			var bstw = new AscCommonExcel.BinaryStylesTableWriter(tmp, null, null);
			bstw.WriteDxf(dxfs[0]);
			s.ImportFromMemory(tmp);
			s.EndRecord();
		}
	};
	CT_SortRule.prototype.fromStream = function (s, wb) {
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd)
				break;
			switch (_at) {
				case 0: {
					this.colId = s.GetULong();
					break;
				}
				case 1: {
					this.id = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		var dxfs = [];
		var _type;
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					var _stream = new AscCommon.FT_Stream2();
					_stream.FromFileStream(s);
					var bsr = new AscCommonExcel.Binary_StylesTableReader(_stream, wb);
					dxfs.push(bsr.ReadDxfExternal());
					_stream.ToFileStream2(s);
					break;
				}
				// case 1:
				// {
				// 	this.richSortCondition = new CT_RichSortCondition();
				// 	this.richSortCondition.fromStream(s);
				// 	break;
				// }
				case 2: {
					var _stream = new AscCommon.FT_Stream2();
					_stream.FromFileStream(s);
					var oReadResult = new AscCommonWord.DocReadResult(null);
					var bwtr = new AscCommonExcel.Binary_TableReader(_stream, oReadResult, null, dxfs);
					this.sortCondition = bwtr.ReadSortConditionExternal();
					//dxfId is absent in sortCondition
					if ((Asc.ESortBy.sortbyCellColor === this.sortCondition.ConditionSortBy ||
						Asc.ESortBy.sortbyFontColor === this.sortCondition.ConditionSortBy)
						&& null === this.sortCondition.dxf && dxfs.length > 0) {
						this.sortCondition.dxf = dxfs[0];
					}
					_stream.ToFileStream2(s);
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};
	CT_SortRule.prototype.Write_ToBinary2 = function (writer) {
		if(null != this.dxf) {
			var dxf = this.dxf;
			writer.WriteBool(true);
			var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(writer);
			oBinaryStylesTableWriter.bs.WriteItem(0, function(){oBinaryStylesTableWriter.WriteDxf(dxf);});
		}else {
			writer.WriteBool(false);
		}

		if(null != this.sortCondition) {
			writer.WriteBool(true);
			this.sortCondition.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		if(null != this.richSortCondition) {
			writer.WriteBool(true);
			this.richSortCondition.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}
	};
	CT_SortRule.prototype.Read_FromBinary2 = function (reader) {
		if (reader.GetBool()) {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var bsr = new AscCommonExcel.Binary_StylesTableReader(reader, wb);
			var bcr = new AscCommon.Binary_CommonReader(reader);
			var oDxf = new AscCommonExcel.CellXfs();
			reader.GetUChar();
			var length = reader.GetULongLE();
			bcr.Read1(length, function (t, l) {
				return bsr.ReadDxf(t, l, oDxf);
			});
			this.dxf = oDxf;
		}

		var obj;
		if (reader.GetBool()) {
			obj = new AscCommonExcel.FilterColumn();
			obj.Read_FromBinary2(reader);
			this.sortCondition = obj;
		}

		if (reader.GetBool()) {
			obj = new AscCommonExcel.SortCondition();
			obj.Read_FromBinary2(reader);
			//TODO CT_RichSortCondition ?
			this.richSortCondition = obj;
		}
	};
	CT_SortRule.prototype.clone = function () {
		var res = new CT_SortRule();
		res.sortCondition = this.sortCondition ? this.sortCondition.clone() : null;
		res.richSortCondition = this.richSortCondition ? this.richSortCondition.clone() : null;

		return res;
	};


	window["Asc"]["CT_NamedSheetViews"] = window['Asc'].CT_NamedSheetViews = CT_NamedSheetViews;
	prot = CT_NamedSheetView.prototype;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_getIsActive"] = prot.asc_getIsActive;
	prot["asc_setIsActive"] = prot.asc_setIsActive;
	prot["asc_getIsLock"] = prot.asc_getIsLock;

	window["Asc"]["CT_NamedSheetView"] = window['Asc'].CT_NamedSheetView = CT_NamedSheetView;
	window["Asc"]["CT_NsvFilter"] = window['Asc'].CT_NsvFilter = CT_NsvFilter;
	window["Asc"]["CT_ColumnFilter"] = window['Asc'].CT_ColumnFilter = CT_ColumnFilter;
	window["Asc"]["CT_SortRule"] = window['Asc'].CT_SortRule = CT_SortRule;

})(window);
