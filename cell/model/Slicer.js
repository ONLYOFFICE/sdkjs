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

	var insertSlicerType = {
		table: 0,
		pivotTable: 1
	};

	var ST_olapSlicerCacheSortOrder = {
		Natural: 0,
		Ascending: 1,
		Descending: 2
	};
	var ST_tabularSlicerCacheSortOrder = {
		Ascending: 0,
		Descending: 1
	};

	var ST_slicerCacheCrossFilter = {
		None: 0,
		ShowItemsWithDataAtTop: 1,
		ShowItemsWithNoData: 2
	};

	var ST_slicerStyleType = {
		unselectedItemWithData: 0,
		selectedItemWithData: 1,
		unselectedItemWithNoData: 2,
		selectedItemWithNoData: 3,
		hoveredUnselectedItemWithData: 4,
		hoveredSelectedItemWithData: 5,
		hoveredUnselectedItemWithNoData: 6,
		hoveredSelectedItemWithNoData: 7
	};

	function CT_slicerStyleElement() {
		this.type = null;
		this.dxfId = null;
		return this;
	}

	CT_slicerStyleElement.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUChar2(0, this.type);
		s._WriteUInt2(1, this.dxfId);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

	};
	CT_slicerStyleElement.prototype.fromStream = function (s) {
		var _type;
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd) {
				break;
			}
			switch (_at) {
				case 0: {
					this.type = s.GetUChar();
					break;
				}
				case 1: {
					this.dxfId = s.GetULong();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		while (true) {
			if (s.cur >= _end_pos) {
				break;
			}
			_type = s.GetUChar();
			switch (_type) {
				default: {
					s.SkipRecord();
					break;
				}
			}
		}
		s.Seek2(_end_pos);
	};

	function CT_slicerStyle() {
		this.name = null;
		this.slicerStyleElements = [];
		return this;
	}

	CT_slicerStyle.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.name);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.slicerStyleElements);
	};
	CT_slicerStyle.prototype.fromStream = function (s) {
		var _type;
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd) {
				break;
			}
			switch (_at) {
				case 0: {
					this.name = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		while (true) {
			if (s.cur >= _end_pos) {
				break;
			}
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_slicerStyleElement();
						tmp.fromStream(s);
						this.slicerStyleElements.push(tmp);
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

	function CT_slicerStyles() {
		this.defaultSlicerStyle = null;
		this.slicerStyle = [];
		return this;
	}

	CT_slicerStyles.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.defaultSlicerStyle);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.slicerStyle);
	};
	CT_slicerStyles.prototype.fromStream = function (s) {
		var _type;
		var _len = s.GetULong();
		var _start_pos = s.cur;
		var _end_pos = _len + _start_pos;
		var _at;
// attributes
		s.GetUChar();
		while (true) {
			_at = s.GetUChar();
			if (_at === AscCommon.g_nodeAttributeEnd) {
				break;
			}
			switch (_at) {
				case 0: {
					this.defaultSlicerStyle = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		while (true) {
			if (s.cur >= _end_pos) {
				break;
			}
			_type = s.GetUChar();
			switch (_type) {
				case 0: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_slicerStyle();
						tmp.fromStream(s);
						this.slicerStyle.push(tmp);
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

	function CT_slicers(opt_ws) {
		this.slicer = [];

		this._ws = opt_ws;
		return this;
	}

	CT_slicers.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.slicer);
	};
	CT_slicers.prototype.fromStream = function (s, ws, slicerCaches) {
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
						var tmp = new CT_slicer(ws);
						tmp.fromStream(s, slicerCaches);
						this.slicer.push(tmp);
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

	function CT_slicer(ws) {
		//from documentation
		this.name = null;
		this.uid = AscCommon.CreateGUID();
		this.cacheDefinition = null;
		this.caption = null;
		this.startItem = 0;
		this.columnCount = 1;
		this.showCaption = true;
		this.level = 0;
		this.style = null;
		this.lockedPosition = false;
		this.rowHeight = null;

		this.ws = ws;

		this._writeBinaryForHistory = null;

		this._ascSourceName = null;
		this._ascNameInFormulas = null;
		this._ascSortOrder = null;
		this._ascCustomListSort = null;
		this._ascHideItemsWithNoData = null;
		this._ascIndicateItemsWithNoData = null;
		this._ascShowItemsWithNoDataLast = null;
		this._ascButtonWidth = null;

		return this;
	}

	CT_slicer.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.Slicer;
	};
	CT_slicer.prototype.clone = function (ws) {
		//если ws-> undefined, то клонирование для интерфейса, только CT_slicer без внутренних структур
		var res = new CT_slicer(ws);

		res.name = this.name;
		//res.uid = this.uid;
		res.caption = this.caption;
		res.startItem = this.startItem;
		res.columnCount = this.columnCount;
		res.showCaption = this.showCaption;
		res.level = this.level;
		res.style = this.style;
		res.lockedPosition = this.lockedPosition;
		res.rowHeight = this.rowHeight;

		if (ws) {
			res.uid = AscCommon.CreateGUID();
			res.cacheDefinition = this.cacheDefinition.clone(ws.workbook);
			return res;
		}

		this.initInterfaceOptions();

		res._ascSourceName = this._ascSourceName;
		res._ascNameInFormulas = this._ascNameInFormulas;
		res._ascSortOrder = this._ascSortOrder;
		res._ascCustomListSort = this._ascCustomListSort;
		res._ascHideItemsWithNoData = this._ascHideItemsWithNoData;

		res._ascIndicateItemsWithNoData = this._ascIndicateItemsWithNoData;
		res._ascShowItemsWithNoDataLast = this._ascShowItemsWithNoDataLast;
		res._ascButtonWidth = this._ascButtonWidth;

		return res;
	};
	CT_slicer.prototype.merge = function (val) {
		if (val.name !== this.name) {
			this.name = undefined;
		}
		if (val.caption !== this.caption) {
			this.caption = undefined;
		}
		if (val.startItem !== this.startItem) {
			this.startItem = undefined;
		}
		if (val.columnCount !== this.columnCount) {
			this.columnCount = undefined;
		}
		if (val.showCaption !== this.showCaption) {
			this.showCaption = undefined;
		}
		if (val.level !== this.level) {
			this.level = undefined;
		}
		if (val.style !== this.style) {
			this.style = undefined;
		}
		if (val.lockedPosition !== this.lockedPosition) {
			this.lockedPosition = undefined;
		}
		if (val.rowHeight !== this.rowHeight) {
			this.rowHeight = undefined;
		}

		if (val._ascSourceName !== this._ascSourceName) {
			this._ascSourceName = undefined;
		}
		if (val._ascNameInFormulas !== this._ascNameInFormulas) {
			this._ascNameInFormulas = undefined;
		}
		if (val._ascSortOrder !== this._ascSortOrder) {
			this._ascSortOrder = undefined;
		}
		if (val._ascCustomListSort !== this._ascCustomListSort) {
			this._ascCustomListSort = undefined;
		}
		if (val._ascHideItemsWithNoData !== this._ascHideItemsWithNoData) {
			this._ascHideItemsWithNoData = undefined;
		}
		if (val._ascIndicateItemsWithNoData !== this._ascIndicateItemsWithNoData) {
			this._ascIndicateItemsWithNoData = undefined;
		}
		if (val._ascShowItemsWithNoDataLast !== this._ascShowItemsWithNoDataLast) {
			this._ascShowItemsWithNoDataLast = undefined;
		}
		if (val._ascButtonWidth !== this._ascButtonWidth) {
			this._ascButtonWidth = undefined;
		}
	};
	CT_slicer.prototype.init = function (name, obj_name, type, ws, pivotTable) {
		var isNewCache = false;
		if (name) {
			this.name = this.generateName(name);
			this.caption = name;
		}
		if (!this.rowHeight) {
			this.rowHeight = 241300;
		}

		if (!this.cacheDefinition) {
			//необходимо проверить, возможно данный кэш уже существует
			var cache;
			var caches = this.ws.getSlicerCachesBySourceName(name);
			if (caches) {
				for (var i = 0; i < caches.length; i++) {
					if (caches[i].checkObjApply(name, obj_name, type, pivotTable)) {
						cache = caches[i];
						break;
					}
				}
			}
			if (!cache) {
				cache = new CT_slicerCacheDefinition(this.ws.workbook);
				cache.init(name, obj_name, type, pivotTable);
				isNewCache = true;
			}

			this.cacheDefinition = cache;
		}

		if (ws && !this.ws) {
			this.ws = ws;
			if (!this.cacheDefinition.isInit()) {
				this.cacheDefinition.initAfterSerialize(ws.workbook);
			}
		}
		return isNewCache;
	};
	CT_slicer.prototype.initPostOpen = function (tableIds, sheetIds) {
		if (this.cacheDefinition) {
			this.cacheDefinition.initAfterSerialize(this.ws.workbook);
			var tableCache = this.cacheDefinition.tableSlicerCache;
			if (tableCache) {
				var _obj = tableCache.initPostOpen(tableIds, this.cacheDefinition.sourceName);
				if (_obj) {
					this.cacheDefinition._type = insertSlicerType.table;
				}
			}
			var tabular = this.cacheDefinition.getTabular();
			if (tabular) {
				this.cacheDefinition._type = insertSlicerType.pivotTable;
			}
			for (var i = 0; i < this.cacheDefinition.pivotTables.length; ++i) {
				this.cacheDefinition.pivotTables[i].initPostOpen(sheetIds);
			}
		}
	};
	CT_slicer.prototype.getIndexSheetCache = function () {
		if (this.cacheDefinition) {
			return this.cacheDefinition.getIndexSheetCache();
		}
		return null;
	};
	CT_slicer.prototype.initPostOpenZip = function (pivotCaches) {
		if (this.cacheDefinition) {
			this.cacheDefinition.initAfterSerialize(this.ws.workbook, pivotCaches);
		}
	};
	CT_slicer.prototype.initInterfaceOptions = function () {
		this._ascSourceName = this.getSourceName();
		this._ascNameInFormulas = this.getNameInFormulas();
		this._ascSortOrder = this.getSortOrder();
		this._ascCustomListSort = this.getCustomListSort();
		this._ascHideItemsWithNoData = this.getHideItemsWithNoData();
		this._ascIndicateItemsWithNoData = this.getIndicateItemsWithNoData();
		this._ascShowItemsWithNoDataLast = this.getShowItemsWithNoDataLast();
	};

	CT_slicer.prototype.Write_ToBinary2 = function (w) {
		this._writeBinaryForHistory = true;

		var old = new AscCommon.CMemory(true);
		pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
		pptx_content_writer.BinaryFileWriter.ImportFromMemory(w);
		pptx_content_writer.BinaryFileWriter.WriteRecord4(0, this);
		pptx_content_writer.BinaryFileWriter.ExportToMemory(w);
		pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);

		this._writeBinaryForHistory = false;
	};

	CT_slicer.prototype.Read_FromBinary2 = function (r) {
		var fileStream = r.ToFileStream();
		fileStream.GetUChar();
		this.fromStream(fileStream, null, true);
		r.FromFileStream(fileStream);
	}

	CT_slicer.prototype.toStream = function (s) {
		var historySerialize = this._writeBinaryForHistory;
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.name);
		s._WriteString2(1, this.uid);
		if (this.cacheDefinition && !historySerialize) {
			s._WriteString2(2, this.cacheDefinition.name);
		}
		s._WriteString2(3, this.caption);
		s._WriteUInt2(4, this.startItem);
		s._WriteUInt2(5, this.columnCount);
		s._WriteBool2(6, this.showCaption);
		s._WriteUInt2(7, this.level);
		s._WriteString2(8, this.style);
		s._WriteBool2(9, this.lockedPosition);
		s._WriteUInt2(10, this.rowHeight);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (null != this.cacheDefinition && historySerialize) {
			s.StartRecord(0);
			this.cacheDefinition.toStream(s, null, null, true);
			s.EndRecord();
		}
	};
	CT_slicer.prototype.fromStream = function (s, slicerCaches, historySerialize) {
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
					this.uid = s.GetString2();
					break;
				}
				case 2: {
					var cache = s.GetString2();
					this.cacheDefinition = slicerCaches[cache] || null;
					break;
				}
				case 3: {
					this.caption = s.GetString2();
					if (typeof this.caption === "string") {
						if (AscFonts.IsCheckSymbols) {
							AscFonts.FontPickerByCharacter.getFontsByString(this.caption);
						}
					}
					break;
				}
				case 4: {
					this.startItem = s.GetULong();
					break;
				}
				case 5: {
					this.columnCount = s.GetULong();
					break;
				}
				case 6: {
					this.showCaption = s.GetBool();
					break;
				}
				case 7: {
					this.level = s.GetULong();
					break;
				}
				case 8: {
					this.style = s.GetString2();
					break;
				}
				case 9: {
					this.lockedPosition = s.GetBool();
					break;
				}
				case 10: {
					this.rowHeight = s.GetULong();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}

		while (true) {
			if (s.cur >= _end_pos)
				break;
			_at = s.GetUChar();
			switch (_at) {
				case 0: {
					this.cacheDefinition = new CT_slicerCacheDefinition();
					this.cacheDefinition.fromStream(s, historySerialize);
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
	CT_slicer.prototype.isExt = function () {
		return !!this.getTableSlicerCache();
	};
	CT_slicer.prototype.setName = function (val) {
		var oldVal = this.name;
		this.name = val;
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetName,
			this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(null, oldVal, val));
		this.ws.workbook.onSlicerChangeName(oldVal, val);
	};
	CT_slicer.prototype.setCaption = function (val) {
		var oldVal = this.caption;
		this.caption = val;
		if (typeof this.caption === "string") {
			if (AscFonts.IsCheckSymbols) {
				AscFonts.FontPickerByCharacter.getFontsByString(this.caption);
			}
		}
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCaption,
			this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
	};
	CT_slicer.prototype.getSlicerCache = function () {
		return this.cacheDefinition;
	};
	CT_slicer.prototype.getTableSlicerCache = function () {
		return this.cacheDefinition && this.cacheDefinition.getTableSlicerCache();
	};
	CT_slicer.prototype.getTabular = function () {
		return this.cacheDefinition && this.cacheDefinition.getTabular();
	};
	CT_slicer.prototype.getCacheDefinition = function () {
		return this.cacheDefinition;
	};
	CT_slicer.prototype.generateName = function (name) {
		var wb = this.ws.workbook;
		var mapNames = [];
		var isContainName = false;
		for (var i = 0; i < wb.aWorksheets.length; i++) {
			if (!wb.aWorksheets[i].aSlicers) {
				continue;
			}
			for (var j = 0; j < wb.aWorksheets[i].aSlicers.length; j++) {
				if (name === wb.aWorksheets[i].aSlicers[j].name) {
					isContainName = true;
				}
				mapNames[wb.aWorksheets[i].aSlicers[j].name] = 1;
			}
		}
		if (isContainName) {
			var baseName = name + " ";
			var counter = 1;
			while (mapNames[baseName + counter]) {
				counter++;
			}
			name = baseName + counter;
		}

		return name;
	};

	CT_slicer.prototype.setTableColName = function (oldVal, newVal) {
		History.Create_NewPoint();
		History.StartTransaction();

		//TODO передать информацию во view о смене caption
		if (this.caption === oldVal) {
			this.setCaption(newVal);
		}
		this.setSourceName(newVal);
		this.setTableCacheColName(newVal);


		History.EndTransaction();
	};

	CT_slicer.prototype.setSourceName = function (val) {
		if (this.cacheDefinition) {
			var oldVal = this.cacheDefinition.sourceName;
			this.cacheDefinition.setSourceName(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheSourceName,
				this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
		}
	};

	CT_slicer.prototype.setTableCacheColName = function (val) {
		if (!this.cacheDefinition) {
			return;
		}

		var _tableCache = this.cacheDefinition.getTableSlicerCache();
		if (_tableCache) {
			var oldVal = _tableCache.column;
			_tableCache.setColName(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetTableColName,
				this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
		}
	};

	CT_slicer.prototype.setTableName = function (val) {
		if (!this.cacheDefinition) {
			return;
		}

		var _tableCache = this.cacheDefinition.getTableSlicerCache();
		if (_tableCache) {
			var oldVal = _tableCache.tableId;
			_tableCache.setTableName(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetTableName,
				this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
		}
	};

	CT_slicer.prototype.setCacheName = function (val) {
		if (!this.cacheDefinition) {
			return;
		}

		var oldVal = this.cacheDefinition.name;
		this.cacheDefinition.setName(val);
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheName,
			this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));

	};

	CT_slicer.prototype.set = function (val, opt_moveSheet) {
		if (!opt_moveSheet && this.name !== val.name && val.name !== undefined) {
			this.setName(val.name);
		}

		this.caption = this.checkProperty(this.caption, val.caption, AscCH.historyitem_Slicer_SetCaption);
		this.startItem = this.checkProperty(this.startItem, val.startItem, AscCH.historyitem_Slicer_SetStartItem);
		this.columnCount = this.checkProperty(this.columnCount, val.columnCount, AscCH.historyitem_Slicer_SetColumnCount);
		this.showCaption = this.checkProperty(this.showCaption, val.showCaption, AscCH.historyitem_Slicer_SetShowCaption);
		this.level = this.checkProperty(this.level, val.level, AscCH.historyitem_Slicer_SetLevel);
		this.style = this.checkProperty(this.style, val.style, AscCH.historyitem_Slicer_SetStyle);
		this.lockedPosition = this.checkProperty(this.lockedPosition, val.lockedPosition, AscCH.historyitem_Slicer_SetLockedPosition);
		this.rowHeight = this.checkProperty(this.rowHeight, val.rowHeight, AscCH.historyitem_Slicer_SetRowHeight);

		var isChangedSortOrder = false;
		if (val._ascSortOrder !== undefined) {
			isChangedSortOrder = this.setSortOrder(val._ascSortOrder);
		}
		if (val._ascCustomListSort !== undefined) {
			this.setCustomListSort(val._ascCustomListSort);
		}
		var isChangedHideItemsWithNoData = false;
		if (val._ascHideItemsWithNoData !== undefined) {
			isChangedHideItemsWithNoData = this.setHideItemsWithNoData(val._ascHideItemsWithNoData);
		}

		var isChangedCrossFilter = false;
		var crossFilter = this._ascGenerateCrossFilter(val._ascHideItemsWithNoData, val._ascIndicateItemsWithNoData, val._ascShowItemsWithNoDataLast);
		if (crossFilter !== undefined) {
			isChangedCrossFilter = this.setCrossFilter(crossFilter);
		}
		if (isChangedSortOrder || isChangedHideItemsWithNoData || isChangedCrossFilter) {
			this.updateItemsWithNoData();
		}

		//TODO ws?
		if (!opt_moveSheet) {
			var slicers = this.ws.getSlicersByCacheName(this.cacheDefinition.name);
			if (slicers) {
				for (var i = 0; i < slicers.length; i++) {
					this.ws.workbook.onSlicerUpdate(slicers[i].name);
				}
			}
		}
	};

	CT_slicer.prototype.checkProperty = function (propOld, propNew, type) {
		if (propOld !== propNew && undefined !== propNew) {
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, type,
				this.ws.getId(), null, new AscCommonExcel.UndoRedoData_Slicer(this.name, propOld, propNew));
			return propNew;
		}
		return propOld;
	};

	CT_slicer.prototype.getSourceName = function () {
		return this.cacheDefinition.sourceName;
	};

	CT_slicer.prototype.getNameInFormulas = function () {
		return this.cacheDefinition.name;
	};

	CT_slicer.prototype.getSortOrder = function () {
		//TODO может быть не только таблица
		var table = this.cacheDefinition.getTableSlicerCache() || this.cacheDefinition.getTabular();
		if (table) {
			return table.sortOrder;
		}
		return ST_tabularSlicerCacheSortOrder.Ascending;
	};

	CT_slicer.prototype.getCustomListSort = function () {
		var table = this.cacheDefinition.getTableSlicerCache() || this.cacheDefinition.getTabular();
		if (table) {
			return table.customListSort;
		}
		return true;
	};

	CT_slicer.prototype.getHideItemsWithNoData = function () {
		return this.cacheDefinition.getHideItemsWithNoData();
	};

	CT_slicer.prototype.getIndicateItemsWithNoData = function () {
		return this.cacheDefinition.getIndicateItemsWithNoData();
	};

	CT_slicer.prototype.getShowItemsWithNoDataLast = function () {
		return this.cacheDefinition.getShowItemsWithNoDataLast();
	};

	CT_slicer.prototype.setSortOrder = function (val) {
		return this.cacheDefinition.setSortOrder(val);
	};

	CT_slicer.prototype.setCustomListSort = function (val) {
		this.cacheDefinition.setCustomListSort(val);
	};

	CT_slicer.prototype.setHideItemsWithNoData = function (val) {
		return this.cacheDefinition.setHideItemsWithNoData(val);
	};

	CT_slicer.prototype.setCrossFilter = function (val) {
		return this.cacheDefinition.setCrossFilter(val);
	};

	CT_slicer.prototype.updateItemsWithNoData = function () {
		return this.cacheDefinition.updateItemsWithNoData();
	};

	CT_slicer.prototype._ascGenerateCrossFilter = function (_ascHideItemsWithNoData, _ascIndicateItemsWithNoData, _ascShowItemsWithNoDataLast) {
		if (_ascHideItemsWithNoData || (_ascIndicateItemsWithNoData && _ascShowItemsWithNoDataLast)) {
			return ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
		}
		if (!_ascIndicateItemsWithNoData && !_ascShowItemsWithNoDataLast) {
			return ST_slicerCacheCrossFilter.None;
		}
		if (_ascIndicateItemsWithNoData && !_ascShowItemsWithNoDataLast) {
			return ST_slicerCacheCrossFilter.ShowItemsWithNoData;
		}
		return ST_slicerCacheCrossFilter.None;
	};

	CT_slicer.prototype.asc_setName = function (val) {
		this.name = val;
	};

	CT_slicer.prototype.asc_setCaption = function (val) {
		this.caption = val;
	};

	CT_slicer.prototype.asc_setStartItem = function (val) {
		this.startItem = val;
	};

	CT_slicer.prototype.asc_setColumnCount = function (val) {
		this.columnCount = val;
	};

	CT_slicer.prototype.asc_setShowCaption = function (val) {
		this.showCaption = val;
	};

	CT_slicer.prototype.asc_setLevel = function (val) {
		this.level = val;
	};

	CT_slicer.prototype.asc_setStyle = function (val) {
		this.style = val;
	};

	CT_slicer.prototype.asc_setLockedPosition = function (val) {
		this.lockedPosition = val;
	};

	CT_slicer.prototype.asc_setRowHeight = function (val) {
		this.rowHeight = val;
	};

	CT_slicer.prototype.asc_setSortOrder = function (val) {
		this._ascSortOrder = val;
	};

	CT_slicer.prototype.asc_setHideItemsWithNoData = function (val) {
		this._ascHideItemsWithNoData = val;
	};

	CT_slicer.prototype.asc_setCustomListSort = function (val) {
		this._ascCustomListSort = val;
	};

	CT_slicer.prototype.asc_setIndicateItemsWithNoData = function (val) {
		this._ascIndicateItemsWithNoData = val;
	};

	CT_slicer.prototype.asc_setShowItemsWithNoDataLast = function (val) {
		this._ascShowItemsWithNoDataLast = val;
	};

	CT_slicer.prototype.asc_setButtonWidth = function (val) {
		this._ascButtonWidth = val;
	};

	CT_slicer.prototype.asc_getName = function () {
		return this.name;
	};

	CT_slicer.prototype.asc_getCaption = function () {
		return this.caption;
	};

	CT_slicer.prototype.asc_getStartItem = function () {
		return this.startItem;
	};

	CT_slicer.prototype.asc_getColumnCount = function () {
		return this.columnCount;
	};

	CT_slicer.prototype.asc_getShowCaption = function () {
		return this.showCaption;
	};

	CT_slicer.prototype.asc_getLevel = function () {
		return this.level;
	};

	CT_slicer.prototype.asc_getStyle = function () {
		var sStyle = this.style;
		if (sStyle) {
			return sStyle;
		}
		var wb = Asc.editor.wbModel;
		if (wb) {
			return wb.SlicerStyles.DefaultStyle;
		}
		return null;
	};

	CT_slicer.prototype.asc_getLockedPosition = function () {
		return this.lockedPosition;
	};

	CT_slicer.prototype.asc_getRowHeight = function () {
		return this.rowHeight;
	};

	CT_slicer.prototype.asc_getSourceName = function () {
		return this._ascSourceName;
	};

	CT_slicer.prototype.asc_getNameInFormulas = function () {
		return this._ascNameInFormulas;
	};

	CT_slicer.prototype.asc_getSortOrder = function () {
		return this._ascSortOrder;
	};

	CT_slicer.prototype.asc_getCrossFilter = function () {
		var table = this.cacheDefinition.getTableSlicerCache() || this.cacheDefinition.getTabular();
		if (table) {
			return table.crossFilter;
		}
		return ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
	};

	CT_slicer.prototype.asc_getCustomListSort = function () {
		return this._ascCustomListSort;
	};

	CT_slicer.prototype.asc_getHideItemsWithNoData = function () {
		return this._ascHideItemsWithNoData;
	};

	CT_slicer.prototype.asc_getIndicateItemsWithNoData = function () {
		return this._ascIndicateItemsWithNoData;
	};

	CT_slicer.prototype.asc_getShowItemsWithNoDataLast = function () {
		return this._ascShowItemsWithNoDataLast;
	};
	CT_slicer.prototype.asc_getButtonWidth = function () {
		return this._ascButtonWidth;
	};

	CT_slicer.prototype.asc_getStylesPictures = function () {
		return Asc.editor.asc_getSlicerPictures();
	};

	CT_slicer.prototype.checkModelContent = function (model) {
		return this.cacheDefinition.checkModelContent(model);
	};
	CT_slicer.prototype.getPivotCache = function () {
		return this.cacheDefinition.getPivotCache();
	};


	function CT_slicerCacheDefinition(wb) {
		this.pivotTables = [];//SlicerCachePivotTable
		this.data = null;//CSlicerCacheData
		this.name = null;
		this.uid = AscCommon.CreateGUID();
		this.sourceName = null;
		this.tableSlicerCache = null;
		this.slicerCacheHideItemsWithNoData = null;

		this.wb = wb;

		//пока добавил объект для хранения типа, чтобы не проходится по внутреннему дереву
		this._type = null;

		return this;
	}

	CT_slicerCacheDefinition.prototype.isInit = function () {
		return !!this.wb && (insertSlicerType.pivotTable !== this.getType() || !!this.getPivotCache());
	};
	CT_slicerCacheDefinition.prototype.init = function (name, obj_name, type, pivotTable) {
		switch (type) {
			case insertSlicerType.table: {
				this.sourceName = name;
				//TODO для генерации имени нужна отдельная функция
				this.name = this.generateSlicerCacheName(name);
				this.tableSlicerCache = new CT_tableSlicerCache();
				this.tableSlicerCache.tableId = obj_name;
				this.tableSlicerCache.column = name;
				break;
			}
			case insertSlicerType.pivotTable: {
				var cacheDefinition = pivotTable.cacheDefinition;
				var cacheFields = cacheDefinition.getFields();
				var fieldIndex = pivotTable.asc_getFieldIndexByName(name);
				var cacheField = -1 !== fieldIndex && cacheFields[fieldIndex];
				if (cacheField) {
					this.sourceName = cacheField.asc_getName();
					//TODO для генерации имени нужна отдельная функция
					this.name = this.generateSlicerCacheName(name);
					this.data = new CT_slicerCacheData();
					this.data.tabular = new CT_tabularSlicerCache();
					pivotTable.checkPivotCacheId();
					this.setPivotCacheDefinition(cacheDefinition);

					var table = new CT_slicerCachePivotTable();
					table.name = pivotTable.asc_getName();
					table.sheetId = pivotTable.GetWS().getId();
					this.pivotTables.push(table);
				}

				break;
			}
		}
		this._type = type;
	};
	CT_slicerCacheDefinition.prototype.initAfterSerialize = function (wb, pivotCachesOpen) {
		this.wb = wb;
		var tabular = this.getTabular();
		if (tabular) {
			tabular.initAfterSerialize(this.wb, pivotCachesOpen);
		}
		var olap = this.getOlap();
		if (olap) {
			olap.initAfterSerialize(this.wb, pivotCachesOpen);
		}
	};
	CT_slicerCacheDefinition.prototype.setPivotCacheDefinition = function (pivotCacheDefinition) {
		var tabular = this.getTabular();
		if (tabular) {
			tabular.setPivotCacheDefinition(pivotCacheDefinition);
		}
		var olap = this.getOlap();
		if (olap) {
			olap.setPivotCacheDefinition(pivotCacheDefinition);
		}
	};
	CT_slicerCacheDefinition.prototype.moveToWb = function (wb) {
		this.initAfterSerialize(wb);
		var realPivotTables = wb.getPivotTablesByCacheId(this.getPivotCacheId());
		var oldPivotTables = this.pivotTables;
		this.pivotTables = [];
		for (var i = 0; i < oldPivotTables.length; ++i) {
			for (var j = 0; j < realPivotTables.length; ++j) {
				if (oldPivotTables[i].name === realPivotTables[j].name) {
					this.addPivotTable(realPivotTables[j]);
				}
			}
		}
		return !this.getTabular() || -1 !== this.getPivotFieldIndex();
	};
	CT_slicerCacheDefinition.prototype.addPivotTable = function (pivotTable) {
		var table = new CT_slicerCachePivotTable();
		table.name = pivotTable.asc_getName();
		table.sheetId = pivotTable.GetWS().getId();
		this.pivotTables.push(table);
	};
	CT_slicerCacheDefinition.prototype.clone = function (wb) {
		var res = new CT_slicerCacheDefinition(wb);

		var i
		for (i = 0; i < this.pivotTables.length; i++) {
			res.pivotTables[i] = this.pivotTables[i].clone();//SlicerCachePivotTable
		}
		res.data = this.data ? this.data.clone() : null;//CSlicerCacheData

		res.name = this.name;
		res.uid = AscCommon.CreateGUID();
		res.sourceName = this.sourceName;

		res.tableSlicerCache = this.tableSlicerCache ? this.tableSlicerCache.clone() : null;
		if (this.slicerCacheHideItemsWithNoData) {
			if (this.slicerCacheHideItemsWithNoData.length === 0) {
				res.slicerCacheHideItemsWithNoData = this.slicerCacheHideItemsWithNoData;
			} else {
				//TODO проверить структуру
				for (i = 0; i < this.slicerCacheHideItemsWithNoData.length; i++) {
					if (!res.slicerCacheHideItemsWithNoData) {
						res.slicerCacheHideItemsWithNoData = [];
					}
					res.slicerCacheHideItemsWithNoData[i] = this.slicerCacheHideItemsWithNoData[i];
				}
			}
		}

		return res;
	};
	CT_slicerCacheDefinition.prototype.copyFrom = function (slicerCacheDefinition) {
		this.pivotTables = slicerCacheDefinition.pivotTables;
		this.data = slicerCacheDefinition.data;
		this.name = slicerCacheDefinition.name;
		this.uid = slicerCacheDefinition.uid;
		this.sourceName = slicerCacheDefinition.sourceName;
		this.tableSlicerCache = slicerCacheDefinition.tableSlicerCache;
		this.slicerCacheHideItemsWithNoData = slicerCacheDefinition.slicerCacheHideItemsWithNoData;
	};

	CT_slicerCacheDefinition.prototype.generateSlicerCacheName = function (name) {
		var wb = this.wb;
		var checkAlreadyAdd = function (_name) {
			var _res = false;
			if (wb.getSlicerCacheByCacheName(_name)) {
				_res = true;
			} else if (wb.getDefinesNames(_name)) {
				_res = true;
			}

			return _res;
		};

		//replace not valid symbols
		name = name.replace(/[-+*\/^&%<=>: ;//),]/g, "_");
		//TODO дополнительная проверка - пересмотреть
		if (!AscCommon.rx_defName.test(name)) {
			name = name.replace(/[^a-zA-ZА-Яа-яЁё0-9]/gi, "_")
		}

		var index = 1;
		name = AscCommon.translateManager.getValue("Slicer") + "_" + name;
		var newName = name;
		while (checkAlreadyAdd(newName)) {
			newName = name + index;
			index++;
		}
		return newName;
	};

	CT_slicerCacheDefinition.prototype.Write_ToBinary2 = function (w) {
		var old = new AscCommon.CMemory(true);
		pptx_content_writer.BinaryFileWriter.ExportToMemory(old);
		pptx_content_writer.BinaryFileWriter.ImportFromMemory(w);
		pptx_content_writer.BinaryFileWriter.StartRecord(0);
		this.toStream(pptx_content_writer.BinaryFileWriter, null, null, true);
		pptx_content_writer.BinaryFileWriter.EndRecord();
		pptx_content_writer.BinaryFileWriter.ExportToMemory(w);
		pptx_content_writer.BinaryFileWriter.ImportFromMemory(old);
	};

	CT_slicerCacheDefinition.prototype.Read_FromBinary2 = function (r) {
		var fileStream = r.ToFileStream();
		fileStream.GetUChar();
		this.fromStream(fileStream, true, null);
		r.FromFileStream(fileStream);
	};
	CT_slicerCacheDefinition.prototype.toStream = function (s, tableIds, sheetIds, historySerialize) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.name);
		s._WriteString2(1, this.uid);
		s._WriteString2(2, this.sourceName);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (this.pivotTables.length > 0) {
			s.StartRecord(0);
			s.WriteULong(this.pivotTables.length);
			for (var i = 0; i < this.pivotTables.length; i++) {
				s.StartRecord(0);
				this.pivotTables[i].toStream(s, sheetIds, historySerialize);
				s.EndRecord();
			}
			s.EndRecord();
		}
		s.WriteRecord4(1, this.data);
		if (null != this.tableSlicerCache) {
			s.StartRecord(3);
			this.tableSlicerCache.toStream(s, tableIds, historySerialize);
			s.EndRecord();
		}
		if (null != this.slicerCacheHideItemsWithNoData) {
			var hideNoData = new CT_slicerCacheHideNoData();
			hideNoData.slicerCacheOlapLevelName = this.slicerCacheHideItemsWithNoData;
			hideNoData.count = this.slicerCacheHideItemsWithNoData.length;
			s.WriteRecord4(4, hideNoData);
		}
	};
	CT_slicerCacheDefinition.prototype.fromStream = function (s, historySerialize) {
		var _type;
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
					this.uid = s.GetString2();
					break;
				}
				case 2: {
					this.sourceName = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
//members
		while (true) {
			if (s.cur >= _end_pos)
				break;
			_type = s.GetUChar();
			switch (_type) {
				case 0:
				case 2: {
					s.Skip2(4);
					var _c = s.GetULong();
					for (var i = 0; i < _c; ++i) {
						s.Skip2(1); // type
						var tmp = new CT_slicerCachePivotTable();
						tmp.fromStream(s, historySerialize);
						this.pivotTables.push(tmp);
					}
					break;
				}
				case 1: {
					this.data = new CT_slicerCacheData();
					this.data.fromStream(s);
					break;
				}
				case 3: {
					this.tableSlicerCache = new CT_tableSlicerCache();
					this.tableSlicerCache.fromStream(s, historySerialize);
					break;
				}
				case 4: {
					var hideNoData = new CT_slicerCacheHideNoData();
					hideNoData.fromStream(s);
					this.slicerCacheHideItemsWithNoData = hideNoData.slicerCacheOlapLevelName;
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

	CT_slicerCacheDefinition.prototype.getIndexSheetCache = function () {
		//TODO позже можно использовать данную функцию в функции getFilterValues. сейчас часть кода дублируется.
		var res = null;
		var type = this.getType();
		var wb = this.wb;
		switch (type) {
			case insertSlicerType.table: {
				//пока беру первый элемент, поскольку не очень понятно в каких случаях их вообще может быть несколько
				var tableCache = this.tableSlicerCache;
				var tableObj = wb.getTableByName(tableCache.tableId, true);
				res = tableObj ? tableObj.index : null;
				break;
			}
			case insertSlicerType.pivotTable: {
				//slicer can connect to pivots in different worksheet or can have zero connection
				res = 0;
				break;
			}
		}
		return res;
	};

	CT_slicerCacheDefinition.prototype.getFilterValues = function () {
		var res = null;
		var type = this.getType();
		var wb = this.wb;
		switch (type) {
			case insertSlicerType.table: {
				//пока беру первый элемент, поскольку не очень понятно в каких случаях их вообще может быть несколько
				var tableCache = this.tableSlicerCache;
				var tableObj = wb.getTableByName(tableCache.tableId, true);
				var table = tableObj ? tableObj.table : null;
				if (table) {
					var ws = wb.aWorksheets[tableObj.index];
					var colId = table.getColIdByName(tableCache.column);
					if (colId !== null) {
						var sortObj = {};
						sortObj.fullValues = true;
						sortObj.sortOrder = this.tableSlicerCache.sortOrder === ST_tabularSlicerCacheSortOrder.Ascending;
						sortObj.hideItemsWithNoData = this.getHideItemsWithNoData();
						sortObj.indicateItemsWithNoData = this.getIndicateItemsWithNoData();
						sortObj.showItemsWithNoDataLast = this.getShowItemsWithNoDataLast();

						res = ws.autoFilters.getOpenAndClosedValues(table, colId, null, sortObj);

					}
				}
				break;
			}
			case insertSlicerType.pivotTable: {
				var tabular = this.data.tabular;
				var cacheDefinition = tabular.pivotCacheDefinition;
				let pivotTables = this.getPivotTables();
				if (cacheDefinition) {
					var fieldIndex = cacheDefinition.getFieldIndexByName(this.sourceName);
					if (-1 !== fieldIndex) {
						let pivotField;
						if (pivotTables.length > 0) {
							const slicerCachePivotTable = this.pivotTables[0];
							const pivotTable = slicerCachePivotTable.getPivotTable(wb);
							pivotField = pivotTable.asc_getPivotFields()[fieldIndex];
						}
						var cacheField = cacheDefinition.getFields()[fieldIndex];
						res = {
							values: tabular.getFilterObject(cacheField, this.slicerCacheHideItemsWithNoData, pivotField),
							automaticRowCount: null,
							ignoreCustomFilter: null
						};
					}
				}
				break;
			}
		}
		return res;
	};

	CT_slicerCacheDefinition.prototype.getFilterObj = function () {
		var res = null;
		var type = this.getType();
		var wb = this.wb;
		switch (type) {
			case insertSlicerType.table: {
				//пока беру первый элемент, поскольку не очень понятно в каких случаях их вообще может быть несколько
				var tableCache = this.tableSlicerCache;
				var table = wb.getTableByName(tableCache.tableId);
				if (table) {
					var colId = table.getColIdByName(tableCache.column);
					res = {obj: table, colId: colId}
				}
				break;
			}
			case insertSlicerType.pivotTable: {
				res = this;
				break;
			}
		}
		return res;
	};
	CT_slicerCacheDefinition.prototype.getLinkedPivotSlicers = function () {
		var res = {};
		var wb = this.wb;
		var pivotTables = this.getPivotTables();
		pivotTables.forEach(function (pivotTable) {
			var slicerCaches = wb.getSlicerCachesByPivotTable(pivotTable.GetWS().getId(), pivotTable.asc_getName());
			slicerCaches.forEach(function (slicerCache) {
				var slicers = wb.getSlicersByCacheName(slicerCache.name);
				slicers.forEach(function (slicer) {
					res[slicer.name] = slicer;
				});
			});
		});
		return res;
	};
	CT_slicerCacheDefinition.prototype.applyPivotFilterWithLock = function (api, values, slicerName, excludePivot, confirmation) {
		var t = this;
		var fld = this.getPivotFieldIndex();
		var pivotTables = this.getPivotTables();
		if (-1 !== fld && pivotTables.length > 0) {
			api._isLockedPivotAndConnectedBySlicer(pivotTables[0], [fld], function (res) {
				if (!res) {
					api.wbModel.onSlicerUpdate(slicerName);
					return;
				}
				History.Create_NewPoint();
				History.StartTransaction();
				api.wbModel.dependencyFormulas.lockRecal();
				var changeRes = t.applyPivotFilter(api, values, excludePivot, confirmation);
				api.wbModel.dependencyFormulas.unlockRecal();
				History.EndTransaction();

				let success = api._changePivotEndCheckError(changeRes, function () {
					t.applyPivotFilterWithLock(api, values, slicerName, excludePivot, true);
				});
				if (success) {
					//todo pivot
					api.updateWorksheetByPivotTable(pivotTables[0], changeRes, true, false);
				}
			});
		} else {
			api.wbModel.onSlicerUpdate(slicerName);
		}
	};
	CT_slicerCacheDefinition.prototype.applyPivotFilter = function (api, values, excludePivot, confirmation) {
		var changeRes =  new AscCommonExcel.PivotChangeResult();
		var fld = this.getPivotFieldIndex();
		if (-1 === fld) {
			return;
		}
		var pivotTables = this.getPivotTables();
		if (pivotTables.length === 0) {
			return;
		}
		//todo assign values instead of setVisibleFromValues ?
		var visible = {};
		for (var i = 0; i < values.length; ++i) {
			if (values[i].visible) {
				visible[values[i].val] = 1;
			}
		}
		var wb = this.wb;
		for (var i = 0; i < pivotTables.length; ++i) {
			var pivotTable = pivotTables[i];
			if (pivotTable === excludePivot) {
				continue;
			}
			var autoFilterObject = new Asc.AutoFiltersOptions();
			pivotTable.fillAutoFiltersOptions(autoFilterObject, fld);
			autoFilterObject.setVisibleFromValues(visible);
			autoFilterObject.filter.type = Asc.c_oAscAutoFilterTypes.Filters;
			let changeResCur = api._changePivot(pivotTable, confirmation, true, function (ws) {
				pivotTable.filterPivotItems(fld, autoFilterObject);
			});
			changeRes.merge(changeResCur);
			if (c_oAscError.ID.No !== changeRes.error || c_oAscError.ID.No !== changeRes.warning) {
				return changeRes;
			}
			if (changeResCur.updateRes) {
				pivotTable.syncSlicersWithPivot(changeResCur.updateRes.cacheFieldsWithData);
			}
		}
		var oldVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
		this.data.tabular.fromAutoFiltersOptionsElements(visible);
		//add to history for 0 connection case(usually duplicate of syncSlicersWithPivot)
		var newVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheData,
			null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, newVal));
		var slicers = this.getLinkedPivotSlicers();
		for (var name in slicers) {
			if (slicers.hasOwnProperty(name)) {
				wb.onSlicerUpdate(name);
			}
		}
		return changeRes;
	};
	CT_slicerCacheDefinition.prototype.getSourceName = function () {
		return this.sourceName;
	};
	CT_slicerCacheDefinition.prototype.setSourceName = function (val) {
		this.sourceName = val;
	};

	CT_slicerCacheDefinition.prototype.getTableSlicerCache = function () {
		return this.tableSlicerCache;
	};
	CT_slicerCacheDefinition.prototype.getTabular = function () {
		return this.data && this.data.tabular;
	};
	CT_slicerCacheDefinition.prototype.getOlap = function () {
		return this.data && this.data.olap;
	};

	CT_slicerCacheDefinition.prototype.checkObjApply = function (name, obj_name, type, pivotTable) {
		var res = false;
		var _obj;
		switch (type) {
			case window['AscCommonExcel'].insertSlicerType.table : {
				_obj = this.getTableSlicerCache();
				if (_obj && _obj.tableId === obj_name && _obj.column === name) {
					res = true;
				}
				break;
			}
			case window['AscCommonExcel'].insertSlicerType.pivotTable : {
				res = this.indexOfPivotTable(pivotTable.GetWS().getId(), pivotTable.asc_getName()) >= 0;
				break;
			}
		}
		return res;
	};

	CT_slicerCacheDefinition.prototype.setName = function (val) {
		this.name = val;
	};

	CT_slicerCacheDefinition.prototype.getType = function () {
		if (this._type !== null) {
			return this._type;
		}

		var tableCache = this.getTableSlicerCache();
		if (tableCache) {
			this._type = insertSlicerType.table;
		}
		var tabular = this.getTabular();
		if (tabular) {
			this._type = insertSlicerType.pivotTable;
		}

		return this._type;
	};

	CT_slicerCacheDefinition.prototype.getRange = function () {
		var res = null;
		var type = this.getType();
		var wb = this.wb;
		switch (type) {
			case insertSlicerType.table: {
				//пока беру первый элемент, поскольку не очень понятно в каких случаях их вообще может быть несколько
				var tableCache = this.tableSlicerCache;
				var table = wb.getTableByName(tableCache.tableId);
				if (table) {
					var colId = table.getColIdByName(tableCache.column);
					res = new Asc.Range(table.Ref.c1 + colId, table.Ref.r1, table.Ref.c1 + colId, table.Ref.r2);
				}
				break;
			}
			case insertSlicerType.pivotTable: {
				//
				break;
			}
		}
		return res;
	};

	CT_slicerCacheDefinition.prototype.setSortOrder = function (val) {
		var obj = this.getTableSlicerCache() || this.getTabular();
		if (obj && obj.sortOrder !== val) {
			var oldVal = obj.sortOrder;
			obj.setSortOrder(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheSortOrder,
				null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
			return true;
		}
		return false;
	};

	CT_slicerCacheDefinition.prototype.setCustomListSort = function (val) {
		var obj = this.getTableSlicerCache() || this.getTabular();
		if (obj && obj.customListSort !== val) {
			var oldVal = obj.customListSort;
			obj.setCustomListSort(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheCustomListSort,
				null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
		}
	};

	CT_slicerCacheDefinition.prototype.setCrossFilter = function (val) {
		var obj = this.getTableSlicerCache() || this.getTabular();
		if (obj && obj.crossFilter !== val) {
			var oldVal = obj.crossFilter;
			obj.setCrossFilter(val);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheCrossFilter,
				null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
			return true;
		}
		return false;
	};

	CT_slicerCacheDefinition.prototype.setHideItemsWithNoData = function (val) {
		//TODO ?
		var oldVal = this.slicerCacheHideItemsWithNoData;
		var newVal = val ? [] : null;
		if (oldVal !== newVal) {
			this.slicerCacheHideItemsWithNoData = val ? [] : null;
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheHideItemsWithNoData,
				null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, val));
			return true;
		}
		return false;
	};

	CT_slicerCacheDefinition.prototype.updateItemsWithNoData = function () {
		var tabular = this.getTabular();
		var pivotTables = this.getPivotTables();
		if (tabular && pivotTables.length > 0) {
			var pivotTable = pivotTables[0];
			var calculateRes = pivotTable.calculateDataRow();
			this.syncWithPivot(pivotTable, calculateRes.cacheFieldsWithData);
		}
	};

	CT_slicerCacheDefinition.prototype.getHideItemsWithNoData = function () {
		return null !== this.slicerCacheHideItemsWithNoData;
	};

	CT_slicerCacheDefinition.prototype.getIndicateItemsWithNoData = function () {
		var hideItemsWithNoData = this.getHideItemsWithNoData();
		if (hideItemsWithNoData) {
			return true;
		} else {
			var table = this.getTableSlicerCache() || this.getTabular();
			if (table) {
				return table.crossFilter !== ST_slicerCacheCrossFilter.None;
			}
			return true;
		}
	};

	CT_slicerCacheDefinition.prototype.getShowItemsWithNoDataLast = function () {
		var hideItemsWithNoData = this.getHideItemsWithNoData();
		if (hideItemsWithNoData) {
			return true;
		} else {
			var table = this.getTableSlicerCache() || this.getTabular();
			if (table) {
				return table.crossFilter === ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
			}
			return true;
		}
	};

	CT_slicerCacheDefinition.prototype.checkModelContent = function (model) {
		var res = null;
		var type = this.getType();
		var wb = model.workbook;
		switch (type) {
			case insertSlicerType.table: {
				var tableCache = this.tableSlicerCache;
				var table = wb.getTableByName(tableCache.tableId);
				if (table) {
					if (null !== table.getColIdByName(tableCache.column)) {
						res = true;
					}
				}
				break;
			}
			case insertSlicerType.pivotTable: {
				var tabular = this.getTabular();
				var olap = this.getOlap();
				if ((!tabular || tabular.checkModelContent(wb)) && (!olap || olap.checkModelContent(wb))) {
					res = true;
				}
				break;
			}
		}
		return res;
	};
	CT_slicerCacheDefinition.prototype.getPivotCache = function () {
		var tabular = this.getTabular();
		return tabular && tabular.pivotCacheDefinition || null;
	};
	CT_slicerCacheDefinition.prototype.getPivotCacheId = function () {
		//todo olap
		var tabular = this.getTabular();
		return tabular && tabular.pivotCacheDefinition && tabular.pivotCacheDefinition.getPivotCacheId() || null;
	};
	CT_slicerCacheDefinition.prototype.getPivotFieldIndex = function () {
		var tabular = this.getTabular();
		if (tabular && tabular.pivotCacheDefinition) {
			var fieldIndex = tabular.pivotCacheDefinition.getFieldIndexByName(this.sourceName);
			if (-1 !== fieldIndex) {
				return fieldIndex;
			}
		}
		return -1;
	};
	CT_slicerCacheDefinition.prototype.getPivotTables = function () {
		var res = [];
		var wb = this.wb;
		this.pivotTables.forEach(function (table) {
			var pivotTable = table.getPivotTable(wb);
			if (pivotTable) {
				res.push(pivotTable);
			}
		});
		return res;
	};
	CT_slicerCacheDefinition.prototype.indexOfPivotTable = function (sheetId, pivotName) {
		return this.pivotTables.findIndex(function (elem) {
			return elem.sheetId === sheetId && elem.name === pivotName;
		});
	};
	CT_slicerCacheDefinition.prototype.deletePivotTable = function (sheetId, pivotName) {
		var oldVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
		var index = this.indexOfPivotTable(sheetId, pivotName);
		if (index >= 0) {
			this.pivotTables.splice(index, 1);
		}
		var newVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheData,
			null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, newVal));
	};
	CT_slicerCacheDefinition.prototype.getPivotTablesCount = function () {
		return this.pivotTables.length;
	};
	CT_slicerCacheDefinition.prototype.syncWithPivot = function (pivotTable, cacheFieldsWithData) {
		var tabular = this.getTabular();
		if (!tabular) {
			return;
		}
		var cacheFields = pivotTable.cacheDefinition.getFields();
		var fieldIndex = pivotTable.cacheDefinition.getFieldIndexByName(this.sourceName);
		var cacheField = -1 !== fieldIndex && cacheFields[fieldIndex];
		if (cacheField) {
			pivotTable.checkPivotFieldItems(fieldIndex);
			var pivotField = pivotTable.asc_getPivotFields()[fieldIndex];

			var oldVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
			tabular.syncWithCache(cacheField, pivotField, cacheFieldsWithData[fieldIndex]);
			var newVal = new AscCommonExcel.UndoRedoData_BinaryWrapper2(this);
			History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheData,
				null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, oldVal, newVal));
		}
	};
	CT_slicerCacheDefinition.prototype.movePivotTable = function (sheetIdFrom, nameFrom, sheetIdTo, nameTo) {
		this.pivotTables.forEach(function (table) {
			if (table.sheetId === sheetIdFrom && table.name === nameFrom) {
				table.sheetId = sheetIdTo;
				table.name = nameTo;
			}
		});
		var from = new AscCommonExcel.UndoRedoData_CellData(sheetIdFrom, nameFrom);
		var to = new AscCommonExcel.UndoRedoData_CellData(sheetIdTo, nameTo);
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheMovePivot,
			null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, from, to));
	};
	CT_slicerCacheDefinition.prototype.forCopySheet = function (sheetIdFrom, sheetIdTo) {
		this.pivotTables.forEach(function (table) {
			if (table.sheetId === sheetIdFrom) {
				table.sheetId = sheetIdTo;
			}
		});
		History.Add(AscCommonExcel.g_oUndoRedoSlicer, AscCH.historyitem_Slicer_SetCacheCopySheet,
			null, null, new AscCommonExcel.UndoRedoData_Slicer(this.name, sheetIdFrom, sheetIdTo));
	};

	function CT_slicerCacheData() {
		this.olap = null;//OlapSlicerCache
		this.tabular = null;//TabularSlicerCache
	}

	CT_slicerCacheData.prototype.clone = function () {
		var res = new CT_slicerCacheData();

		if (this.olap) {
			res.olap = this.olap.clone();
		}
		if (this.tabular) {
			res.tabular = this.tabular.clone();
		}
		return res;
	};
	CT_slicerCacheData.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecord4(0, this.olap);
		s.WriteRecord4(1, this.tabular);
	};
	CT_slicerCacheData.prototype.fromStream = function (s) {
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
					this.olap = new CT_olapSlicerCache();
					this.olap.fromStream(s);
					break;
				}
				case 1: {
					this.tabular = new CT_tabularSlicerCache();
					this.tabular.fromStream(s);
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

	function CT_slicerCachePivotTable() {
		this.sheetId = null;
		this.tabIdOpen = null;
		this.name = null;
	}

	CT_slicerCachePivotTable.prototype.initPostOpen = function (sheetIds) {
		var ws = null;
		if (null != this.tabIdOpen) {
			ws = sheetIds[this.tabIdOpen];
			if (ws) {
				this.sheetId = ws.getId();
			}
		}
		return ws;
	};
	CT_slicerCachePivotTable.prototype.clone = function () {
		var res = new CT_slicerCachePivotTable();

		res.sheetId = this.sheetId;
		res.tabIdOpen = this.tabIdOpen;
		res.name = this.name;

		return res;
	};
	CT_slicerCachePivotTable.prototype.toStream = function (s, sheetIds, historySerialize) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		if (historySerialize) {
			s._WriteString2(0, this.sheetId);
		} else {
			s._WriteUInt2(0, sheetIds[this.sheetId] || 1);
		}
		s._WriteString2(1, this.name);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

	};
	CT_slicerCachePivotTable.prototype.fromStream = function (s, historySerialize) {
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
					if (historySerialize) {
						this.sheetId = s.GetString2();
					} else {
						this.tabIdOpen = s.GetULong();
					}
					break;
				}
				case 1: {
					this.name = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
		s.Seek2(_end_pos);
	};
	CT_slicerCachePivotTable.prototype.getPivotTable = function (wb) {
		var ws = wb.getWorksheetById(this.sheetId);
		if (ws) {
			return ws.getPivotTableByName(this.name);
		}
		return null;
	};


	function CT_olapSlicerCacheItem() {
		this.p = [];//OlapSlicerCacheItemParent - состоит из одного поля, поэтому данную структуру не добавляю
		this.n = null;
		this.c = null;
		this.nd = false;
	}

	CT_olapSlicerCacheItem.prototype.clone = function () {
		var res = new CT_olapSlicerCacheItem();

		//TODO ?p

		res.n = this.n;
		res.c = this.c;
		res.nd = this.nd;

		return res;
	};
	CT_olapSlicerCacheItem.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.n);
		s._WriteString2(1, this.c);
		s._WriteBool2(2, this.nd);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (this.p.length > 0) {
			var p = this.p.map(function (currentValue) {
				var res = new CT_olapSlicerCacheItemParent();
				res.n = currentValue;
				return res;
			});
			s.WriteRecordArray4(0, 0, p);
		}
	};
	CT_olapSlicerCacheItem.prototype.fromStream = function (s) {
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
					this.n = s.GetString2();
					break;
				}
				case 1: {
					this.c = s.GetString2();
					break;
				}
				case 2: {
					this.nd = s.GetBool();
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
						var tmp = new CT_olapSlicerCacheItemParent();
						tmp.fromStream(s);
						this.p.push(tmp.n);
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

	function CT_olapSlicerCacheRange() {
		this.i = [];//OlapSlicerCacheItem
		this.startItem = null
	}

	CT_olapSlicerCacheRange.prototype.clone = function () {
		var res = new CT_olapSlicerCacheRange();

		for (var i = 0; i < this.i.length; i++) {
			res.i[i] = this.i[i].clone();//OlapSlicerCacheItem
		}

		return res;
	};
	CT_olapSlicerCacheRange.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.startItem);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.i);
	};
	CT_olapSlicerCacheRange.prototype.fromStream = function (s) {
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
					this.startItem = s.GetULong();
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
						var tmp = new CT_olapSlicerCacheItem();
						tmp.fromStream(s);
						this.i.push(tmp);
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

	function CT_olapSlicerCacheLevelData() {
		this.ranges = [];//OlapSlicerCacheRange
		this.uniqueName = null;
		this.sourceCaption = null;
		this.count = null;
		this.sortOrder = ST_olapSlicerCacheSortOrder.Natural;
		this.crossFilter = ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
	}

	CT_olapSlicerCacheLevelData.prototype.clone = function () {
		var res = new CT_olapSlicerCacheLevelData();

		for (var i = 0; i < this.ranges.length; i++) {
			res.ranges[i] = this.ranges[i].clone();//OlapSlicerCacheItem
		}

		res.uniqueName = this.uniqueName;
		res.sourceCaption = this.sourceCaption;
		res.count = this.count;
		res.sortOrder = this.sortOrder;
		res.crossFilter = this.crossFilter;

		return res;
	};
	CT_olapSlicerCacheLevelData.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.uniqueName);
		s._WriteString2(1, this.sourceCaption);
		s._WriteUInt2(2, this.count);
		s._WriteUChar2(3, this.sortOrder);
		s._WriteUChar2(4, this.crossFilter);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.ranges);
	};
	CT_olapSlicerCacheLevelData.prototype.fromStream = function (s) {
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
					this.uniqueName = s.GetString2();
					break;
				}
				case 1: {
					this.sourceCaption = s.GetString2();
					break;
				}
				case 2: {
					this.count = s.GetULong();
					break;
				}
				case 3: {
					this.sortOrder = s.GetUChar();
					break;
				}
				case 4: {
					this.crossFilter = s.GetUChar();
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
						var tmp = new CT_olapSlicerCacheRange();
						tmp.fromStream(s);
						this.ranges.push(tmp);
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

	function CT_olapSlicerCache() {
		this.levels = [];//OlapSlicerCacheLevelData
		this.selections = [];//OlapSlicerCacheSelection
		this.pivotCacheId = null;
		this.pivotCacheDefinition = null;
	}

	CT_olapSlicerCache.prototype.initAfterSerialize = function (wb, pivotCachesOpen) {
		if (null !== this.pivotCacheId) {
			this.pivotCacheDefinition = wb.getPivotCacheById(this.pivotCacheId, pivotCachesOpen);
			this.pivotCacheId = null;
		}
	};
	CT_olapSlicerCache.prototype.setPivotCacheDefinition = function (pivotCacheDefinition) {
		this.pivotCacheDefinition = pivotCacheDefinition;
		this.pivotCacheId = null;
	};
	CT_olapSlicerCache.prototype.checkModelContent = function (wb) {
		return null !== this.pivotCacheId && !!wb.getPivotCacheById(this.pivotCacheId);
	};
	CT_olapSlicerCache.prototype.clone = function () {
		var res = new CT_olapSlicerCache();

		var i;
		for (i = 0; i < this.levels.length; i++) {
			res.levels[i] = this.levels[i].clone();
		}
		for (i = 0; i < this.selections.length; i++) {
			res.selections[i] = this.selections[i].clone();
		}
		res.pivotCacheId = this.pivotCacheId;
		res.pivotCacheDefinition = this.pivotCacheDefinition;

		return res;
	};
	CT_olapSlicerCache.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.getPivotCacheId());
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (this.levels.length > 0) {
			var levelsData = new CT_olapSlicerCacheLevelsData();
			levelsData.level = this.levels;
			levelsData.count = this.levels.length;
			s.WriteRecord4(0, levelsData);
		}
		if (this.selections.length > 0) {
			var selections = new CT_olapSlicerCacheSelections();
			selections.selection = this.selections;
			selections.count = this.selections.length;
			s.WriteRecord4(1, selections);
		}
	};
	CT_olapSlicerCache.prototype.fromStream = function (s) {
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
					this.pivotCacheId = s.GetULong();
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
					var levelsData = new CT_olapSlicerCacheLevelsData();
					levelsData.fromStream(s);
					this.levels = levelsData.level;
					break;
				}
				case 1: {
					var selections = new CT_olapSlicerCacheSelections();
					selections.fromStream(s);
					this.selections = selections.selection;
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
	CT_olapSlicerCache.prototype.getPivotCacheId = function () {
		return this.pivotCacheDefinition && this.pivotCacheDefinition.getPivotCacheId() || 0;
	};

	function CT_olapSlicerCacheLevelsData() {
		this.count = null;
		this.level = [];
		return this;
	}

	CT_olapSlicerCacheLevelsData.prototype.clone = function () {
		var res = new CT_olapSlicerCacheLevelsData();

		res.count = this.count;
		for (var i = 0; i < this.level.length; i++) {
			//TODO need clone?
			res.level[i] = this.level[i].clone();
		}

		return res;
	};
	CT_olapSlicerCacheLevelsData.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.count);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.level);
	};
	CT_olapSlicerCacheLevelsData.prototype.fromStream = function (s) {
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
					this.count = s.GetULong();
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
						var tmp = new CT_olapSlicerCacheLevelData();
						tmp.fromStream(s);
						this.level.push(tmp);
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

	function CT_olapSlicerCacheSelections() {
		this.count = null;
		this.selection = [];
		return this;
	}

	CT_olapSlicerCacheSelections.prototype.clone = function () {
		var res = new CT_olapSlicerCacheSelections();

		res.count = this.count;
		for (var i = 0; i < this.selection.length; i++) {
			//TODO need clone?
			res.selection[i] = this.selection[i].clone();
		}

		return res;
	};
	CT_olapSlicerCacheSelections.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.count);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.selection);
	};
	CT_olapSlicerCacheSelections.prototype.fromStream = function (s) {
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
					this.count = s.GetULong();
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
						var tmp = new CT_olapSlicerCacheSelection();
						tmp.fromStream(s);
						this.selection.push(tmp);
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

	function CT_olapSlicerCacheSelection() {
		this.p = [];//OlapSlicerCacheItemParent - состоит из одного поля, поэтому данную структуру не добавляю
		this.n = null;
	}

	CT_olapSlicerCacheSelections.prototype.clone = function () {
		var res = new CT_olapSlicerCacheSelection();

		for (var i = 0; i < this.p.length; i++) {
			//TODO need clone?
			res.p[i] = this.p[i].clone();
		}
		res.n = this.n;

		return res;
	};
	CT_olapSlicerCacheSelection.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.n);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (this.p.length > 0) {
			var p = this.p.map(function (currentValue) {
				var res = new CT_olapSlicerCacheItemParent();
				res.n = currentValue;
				return res;
			});
			s.WriteRecordArray4(0, 0, p);
		}
	};
	CT_olapSlicerCacheSelection.prototype.fromStream = function (s) {
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
					this.n = s.GetString2();
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
						var tmp = new CT_olapSlicerCacheItemParent();
						tmp.fromStream(s);
						this.p.push(tmp.n);
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

	function CT_olapSlicerCacheItemParent() {
		this.n = null;
		return this;
	}

	CT_olapSlicerCacheItemParent.prototype.clone = function () {
		var res = new CT_olapSlicerCacheItemParent();

		res.n = this.n;

		return res;
	};
	CT_olapSlicerCacheItemParent.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.n);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

	};
	CT_olapSlicerCacheItemParent.prototype.fromStream = function (s) {
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
					this.n = s.GetString2();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
		s.Seek2(_end_pos);
	};

	function CT_tableSlicerCache() {
		this.tableId = null;
		this.tableIdOpen = null;//?
		this.column = null;
		this.columnOpen = null;//?
		this.sortOrder = ST_tabularSlicerCacheSortOrder.Ascending;
		this.customListSort = true;
		this.crossFilter = ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;

		return this;
	}

	CT_tableSlicerCache.prototype.clone = function () {
		var res = new CT_tableSlicerCache();

		res.tableId = this.tableId;
		res.tableIdOpen = this.tableIdOpen;
		res.column = this.column;
		res.columnOpen = this.columnOpen;
		res.sortOrder = this.sortOrder;
		res.customListSort = this.customListSort;
		res.crossFilter = this.crossFilter;

		return res;
	};
	CT_tableSlicerCache.prototype.initPostOpen = function (tableIds, sourceName) {
		var table = null;
		if (null != this.tableIdOpen && null != this.columnOpen) {
			table = tableIds[this.tableIdOpen];
			if (table) {
				this.tableId = table.DisplayName;
				if (sourceName && null !== table.getIndexByColumnName(sourceName)) {
					this.column = sourceName;
				} else {
					this.column = table.getTableNameColumnByIndex(this.columnOpen - 1);
				}
			}
		}
		return table;
	};
	CT_tableSlicerCache.prototype.toStream = function (s, tableIds, historySerialize) {
		var tableIdOpen = null;
		var columnOpen = null;
		var elem = !historySerialize && tableIds && tableIds[this.tableId];
		if (elem) {
			tableIdOpen = elem.id;
			columnOpen = (elem.table.getTableIndexColumnByName(this.column) + 1) || null;
		} else if (historySerialize) {
			tableIdOpen = this.tableId;
			columnOpen = this.column;
		}

		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		if (historySerialize) {
			s._WriteString2(0, tableIdOpen);
			s._WriteString2(1, columnOpen);
		} else {
			s._WriteUInt2(0, tableIdOpen);
			s._WriteUInt2(1, columnOpen);
		}
		s._WriteUChar2(2, this.sortOrder);
		s._WriteBool2(3, this.customListSort);
		s._WriteUChar2(4, this.crossFilter);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);
	};
	CT_tableSlicerCache.prototype.fromStream = function (s, historySerialize) {
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
					if (historySerialize) {
						this.tableId = s.GetString2();
					} else {
						this.tableIdOpen = s.GetULong();
					}
					break;
				}
				case 1: {
					if (historySerialize) {
						this.column = s.GetString2();
					} else {
						this.columnOpen = s.GetULong();
					}
					break;
				}
				case 2: {
					this.sortOrder = s.GetUChar();
					break;
				}
				case 3: {
					this.customListSort = s.GetBool();
					break;
				}
				case 4: {
					this.crossFilter = s.GetUChar();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
		s.Seek2(_end_pos);
	};
	CT_tableSlicerCache.prototype.setTableName = function (val) {
		//TODO history
		this.tableId = val;
	};
	CT_tableSlicerCache.prototype.setColName = function (val) {
		//TODO history
		this.column = val;
	};
	CT_tableSlicerCache.prototype.setSortOrder = function (val) {
		this.sortOrder = val;
	};

	CT_tableSlicerCache.prototype.setCustomListSort = function (val) {
		this.customListSort = val;
	};

	CT_tableSlicerCache.prototype.setCrossFilter = function (val) {
		this.crossFilter = val;
	};

	function CT_tabularSlicerCache() {
		this.items = [];//TabularSlicerCacheItem
		this.pivotCacheId = null;
		this.pivotCacheDefinition = null;
		this.sortOrder = ST_tabularSlicerCacheSortOrder.Ascending;
		this.customListSort = true;
		this.showMissing = true;
		this.crossFilter = ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
	}

	CT_tabularSlicerCache.prototype.initAfterSerialize = function (wb, pivotCachesOpen) {
		if (null !== this.pivotCacheId) {
			this.pivotCacheDefinition = wb.getPivotCacheById(this.pivotCacheId, pivotCachesOpen);
			if (this.pivotCacheDefinition) {
				this.pivotCacheId = null;
			}
		}
	};
	CT_tabularSlicerCache.prototype.setPivotCacheDefinition = function (pivotCacheDefinition) {
		this.pivotCacheDefinition = pivotCacheDefinition;
		this.pivotCacheId = null;
	};
	CT_tabularSlicerCache.prototype.checkModelContent = function (wb) {
		return null !== this.pivotCacheId && !!wb.getPivotCacheById(this.pivotCacheId);
	};
	CT_tabularSlicerCache.prototype.clone = function () {
		var res = new CT_tabularSlicerCache();

		for (var i = 0; i < this.items.length; i++) {
			res.items[i] = this.items[i].clone();
		}
		res.pivotCacheId = this.pivotCacheId;
		res.pivotCacheDefinition = this.pivotCacheDefinition;
		res.sortOrder = this.sortOrder;
		res.customListSort = this.customListSort;
		res.showMissing = this.showMissing;
		res.crossFilter = this.crossFilter;

		return res;
	};
	CT_tabularSlicerCache.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.getPivotCacheId());
		s._WriteUChar2(1, this.sortOrder);
		s._WriteBool2(2, this.customListSort);
		s._WriteBool2(3, this.showMissing);
		s._WriteUChar2(4, this.crossFilter);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		if (this.items.length > 0) {
			var items = new CT_tabularSlicerCacheItems();
			items.i = this.items;
			items.count = this.items.length;
			s.WriteRecord4(0, items);
		}
	};
	CT_tabularSlicerCache.prototype.fromStream = function (s) {
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
					this.pivotCacheId = s.GetULong();
					break;
				}
				case 1: {
					this.sortOrder = s.GetUChar();
					break;
				}
				case 2: {
					this.customListSort = s.GetBool();
					break;
				}
				case 3: {
					this.showMissing = s.GetBool();
					break;
				}
				case 4: {
					this.crossFilter = s.GetUChar();
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
					var items = new CT_tabularSlicerCacheItems();
					items.fromStream(s);
					this.items = items.i;
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
	CT_tabularSlicerCache.prototype.syncWithCache = function (cacheField, pivotField, cacheFieldWithData) {
		var pivotItems = pivotField.getItems();
		var count = Math.min(cacheField.getGroupOrSharedSize(), pivotItems.length);
		if (cacheFieldWithData) {
			count = Math.min(count, cacheFieldWithData.length);
		}
		this.items = [];
		for (var i = 0; i < count; ++i) {
			var pivotItem = pivotItems[i];
			if (Asc.c_oAscItemType.Data === pivotItem.t || Asc.c_oAscItemType.Blank === pivotItem.t) {
				var item = new CT_tabularSlicerCacheItem();
				item.x = pivotItem.x;
				item.s = !pivotItem.h;
				if (this.crossFilter !== ST_slicerCacheCrossFilter.None) {
					item.nd = !(!cacheFieldWithData || cacheFieldWithData[item.x]);
				}
				this.items.push(item);
			}
		}
		this.sortItems(this.sortOrder, cacheField.getGroupOrSharedItems());
	};
	CT_tabularSlicerCache.prototype.sortItems = function (type, sharedItems) {
		var sign = ST_tabularSlicerCacheSortOrder.Ascending === type ? 1 : -1;
		if (this.crossFilter === ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop) {
			this.items.sort(function (a, b) {
				if (a.nd === b.nd) {
					return sign * AscCommonExcel.cmpPivotItems(sharedItems, a, b);
				} else {
					return a.nd - b.nd;
				}
			});
		} else {
			this.items.sort(function (a, b) {
				return sign * AscCommonExcel.cmpPivotItems(sharedItems, a, b);
			});
		}
	};
	CT_tabularSlicerCache.prototype.getFilterObject = function (cacheField, slicerCacheHideItemsWithNoData, opt_pivotField) {
		var values = [];
		let pivotItems;
		if (opt_pivotField) {
			pivotItems = opt_pivotField.getItems();
		}
		for (var i = 0; i < this.items.length; ++i) {
			var item = this.items[i];
			var elem = AscCommonExcel.AutoFiltersOptionsElements();
			var sharedItem = cacheField.getGroupOrSharedItem(item.x);
			var num = sharedItem.isDateOrNum() && cacheField.getNumFormat();
			let name;
			if (pivotItems) {
				const pivotItem = pivotItems[opt_pivotField.getItemIndexByValue(item.x)];
				name = pivotItem.getName(cacheField, num);
			} else {
				name = sharedItem.getCellValue().getTextValue(num);
			}
			elem.val = item.x;
			elem.text = name;
			elem.visible = item.s;
			elem.hiddenByOtherColumns = item.nd || undefined;//todo
			elem.isDateFormat = false;
			elem.repeats = undefined;
			//todo isDateFormat
			if (!slicerCacheHideItemsWithNoData || !elem.hiddenByOtherColumns) {
				values.push(elem);
			}
		}
		return values;
	};
	CT_tabularSlicerCache.prototype.fromAutoFiltersOptionsElements = function (visible) {
		for (var i = 0; i < this.items.length; ++i) {
			this.items[i].s = !!visible[this.items[i].x];
		}
	};
	CT_tabularSlicerCache.prototype.setSortOrder = function (val) {
		this.sortOrder = val;
	};
	CT_tabularSlicerCache.prototype.setCustomListSort = function (val) {
		this.customListSort = val;
	};
	CT_tabularSlicerCache.prototype.setCrossFilter = function (val) {
		this.crossFilter = val;
	};
	CT_tabularSlicerCache.prototype.getPivotCacheId = function () {
		return this.pivotCacheDefinition && this.pivotCacheDefinition.getPivotCacheId() || 0;
	};

	function CT_slicerCacheOlapLevelName() {
		this.uniqueName = null;
		this.count = null;
		return this;
	}

	CT_slicerCacheOlapLevelName.prototype.clone = function () {
		var res = new CT_slicerCacheOlapLevelName();

		res.uniqueName = this.uniqueName;
		res.count = this.count;

		return res;
	};
	CT_slicerCacheOlapLevelName.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteString2(0, this.uniqueName);
		s._WriteUInt2(1, this.count);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

	};
	CT_slicerCacheOlapLevelName.prototype.fromStream = function (s) {
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
					this.uniqueName = s.GetString2();
					break;
				}
				case 1: {
					this.count = s.GetULong();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
		s.Seek2(_end_pos);
	};

	function CT_slicerCacheHideNoData() {
		this.count = 0;
		this.slicerCacheOlapLevelName = [];
		return this;
	}

	CT_slicerCacheHideNoData.prototype.clone = function () {
		var res = new CT_slicerCacheHideNoData();

		res.count = this.count;
		for (var i = 0; i < this.slicerCacheOlapLevelName.length; i++) {
			res.slicerCacheOlapLevelName[i] = this.slicerCacheOlapLevelName[i].clone();
		}

		return res;
	};
	CT_slicerCacheHideNoData.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.count);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.slicerCacheOlapLevelName);
	};
	CT_slicerCacheHideNoData.prototype.fromStream = function (s) {
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
					this.count = s.GetULong();
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
						var tmp = new CT_slicerCacheOlapLevelName();
						tmp.fromStream(s);
						this.slicerCacheOlapLevelName.push(tmp);
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


	function CT_tabularSlicerCacheItems() {
		this.count = null;
		this.i = [];
		return this;
	}

	CT_tabularSlicerCacheItems.prototype.clone = function () {
		var res = new CT_tabularSlicerCacheItems();

		res.count = this.count;
		for (var i = 0; i < this.i.length; i++) {
			res.i[i] = this.i[i].clone();
		}

		return res;
	};
	CT_tabularSlicerCacheItems.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.count);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

		s.WriteRecordArray4(0, 0, this.i);
	};
	CT_tabularSlicerCacheItems.prototype.fromStream = function (s) {
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
					this.count = s.GetULong();
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
						var tmp = new CT_tabularSlicerCacheItem();
						tmp.fromStream(s);
						this.i.push(tmp);
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

	function CT_tabularSlicerCacheItem() {
		this.x = null;
		this.s = false;
		this.nd = false;
	}

	CT_tabularSlicerCacheItem.prototype.clone = function () {
		var res = new CT_tabularSlicerCacheItem();
		res.x = this.x;
		res.s = this.s;
		res.nd = this.nd;
		return res;
	};
	CT_tabularSlicerCacheItem.prototype.toStream = function (s) {
		s.WriteUChar(AscCommon.g_nodeAttributeStart);
		s._WriteUInt2(0, this.x);
		s._WriteBool2(1, this.s);
		s._WriteBool2(2, this.nd);
		s.WriteUChar(AscCommon.g_nodeAttributeEnd);

	};
	CT_tabularSlicerCacheItem.prototype.fromStream = function (s) {
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
					this.x = s.GetULong();
					break;
				}
				case 1: {
					this.s = s.GetBool();
					break;
				}
				case 2: {
					this.nd = s.GetBool();
					break;
				}
				default:
					s.Seek2(_end_pos);
					return;
			}
		}
		s.Seek2(_end_pos);
	};

	var prot;
	window['Asc'] = window['Asc'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window['Asc']['CT_slicer'] = window['Asc'].CT_slicer = CT_slicer;
	prot = CT_slicer.prototype;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_setCaption"] = prot.asc_setCaption;
	prot["asc_setStartItem"] = prot.asc_setStartItem;
	prot["asc_setColumnCount"] = prot.asc_setColumnCount;
	prot["asc_setShowCaption"] = prot.asc_setShowCaption;
	prot["asc_setStyle"] = prot.asc_setStyle;
	prot["asc_setLockedPosition"] = prot.asc_setLockedPosition;
	prot["asc_setRowHeight"] = prot.asc_setRowHeight;
	prot["asc_setLockedPosition"] = prot.asc_setLockedPosition;
	prot["asc_setSortOrder"] = prot.asc_setSortOrder;
	prot["asc_setHideItemsWithNoData"] = prot.asc_setHideItemsWithNoData;
	prot["asc_setCustomListSort"] = prot.asc_setCustomListSort;
	prot["asc_setIndicateItemsWithNoData"] = prot.asc_setIndicateItemsWithNoData;
	prot["asc_setShowItemsWithNoDataLast"] = prot.asc_setShowItemsWithNoDataLast;
	prot["asc_setButtonWidth"] = prot.asc_setButtonWidth;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getCaption"] = prot.asc_getCaption;
	prot["asc_getStartItem"] = prot.asc_getStartItem;
	prot["asc_getColumnCount"] = prot.asc_getColumnCount;
	prot["asc_getShowCaption"] = prot.asc_getShowCaption;
	prot["asc_getStyle"] = prot.asc_getStyle;
	prot["asc_getLockedPosition"] = prot.asc_getLockedPosition;
	prot["asc_getRowHeight"] = prot.asc_getRowHeight;
	prot["asc_getSourceName"] = prot.asc_getSourceName;
	prot["asc_getNameInFormulas"] = prot.asc_getNameInFormulas;
	prot["asc_getSortOrder"] = prot.asc_getSortOrder;
	prot["asc_getCrossFilter"] = prot.asc_getCrossFilter;
	prot["asc_getHideItemsWithNoData"] = prot.asc_getHideItemsWithNoData;
	prot["asc_getCustomListSort"] = prot.asc_getCustomListSort;
	prot["asc_getIndicateItemsWithNoData"] = prot.asc_getIndicateItemsWithNoData;
	prot["asc_getShowItemsWithNoDataLast"] = prot.asc_getShowItemsWithNoDataLast;
	prot["asc_getButtonWidth"] = prot.asc_getButtonWidth;
	prot["asc_getStylesPictures"] = prot.asc_getStylesPictures;

	window['Asc']['CT_slicers'] = window['Asc'].CT_slicers = CT_slicers;
	window['Asc']['CT_slicerCacheDefinition'] = window['Asc'].CT_slicerCacheDefinition = CT_slicerCacheDefinition;
	window['Asc']['CT_slicerCacheData'] = window['Asc'].CT_slicerCacheData = CT_slicerCacheData;
	window['Asc']['CT_slicerCachePivotTable'] = window['Asc'].CT_slicerCachePivotTable = CT_slicerCachePivotTable;
	window['Asc']['CT_olapSlicerCacheItem'] = window['Asc'].CT_olapSlicerCacheItem = CT_olapSlicerCacheItem;
	window['Asc']['CT_olapSlicerCacheRange'] = window['Asc'].CT_olapSlicerCacheRange = CT_olapSlicerCacheRange;
	window['Asc']['CT_olapSlicerCacheLevelData'] = window['Asc'].CT_olapSlicerCacheLevelData = CT_olapSlicerCacheLevelData;
	window['Asc']['CT_olapSlicerCache'] = window['Asc'].CT_olapSlicerCache = CT_olapSlicerCache;
	window['Asc']['CT_olapSlicerCacheSelection'] = window['Asc'].CT_olapSlicerCacheSelection = CT_olapSlicerCacheSelection;
	window['Asc']['CT_tableSlicerCache'] = window['Asc'].CT_tableSlicerCache = CT_tableSlicerCache;
	window['Asc']['CT_tabularSlicerCache'] = window['Asc'].CT_tabularSlicerCache = CT_tabularSlicerCache;
	window['Asc']['CT_tabularSlicerCacheItem'] = window['Asc'].CT_tabularSlicerCacheItem = CT_tabularSlicerCacheItem;
	window['Asc']['CT_slicerStyles'] = window['Asc'].CT_slicerStyles = CT_slicerStyles;
	window['Asc']['CT_slicerStyle'] = window['Asc'].CT_slicerStyle = CT_slicerStyle;
	window['Asc']['CT_slicerStyleElement'] = window['Asc'].CT_slicerStyleElement = CT_slicerStyleElement;
	window['Asc']['CT_slicerCacheHideNoData'] = window['Asc'].CT_slicerCacheHideNoData = CT_slicerCacheHideNoData;
	window['Asc']['CT_slicerCacheOlapLevelName'] = window['Asc'].CT_slicerCacheOlapLevelName = CT_slicerCacheOlapLevelName;
	window['Asc']['CT_olapSlicerCacheItemParent'] = window['Asc'].CT_olapSlicerCacheItemParent = CT_olapSlicerCacheItemParent;

	window['Asc']['ST_olapSlicerCacheSortOrder'] = window['AscCommonExcel'].ST_olapSlicerCacheSortOrder = ST_olapSlicerCacheSortOrder;
	prot = ST_olapSlicerCacheSortOrder;
	prot['Natural'] = prot.Natural;
	prot['Ascending'] = prot.Ascending;
	prot['Descending'] = prot.Descending;

	window['Asc']['ST_tabularSlicerCacheSortOrder'] = window['Asc'].ST_tabularSlicerCacheSortOrder = ST_tabularSlicerCacheSortOrder;
	prot = ST_tabularSlicerCacheSortOrder;
	prot['Ascending'] = prot.Ascending;
	prot['Descending'] = prot.Descending;

	window['Asc']['ST_slicerCacheCrossFilter'] = window['Asc'].ST_slicerCacheCrossFilter = ST_slicerCacheCrossFilter;
	prot = ST_slicerCacheCrossFilter;
	prot['None'] = prot.None;
	prot['ShowItemsWithDataAtTop'] = prot.ShowItemsWithDataAtTop;
	prot['ShowItemsWithNoData'] = prot.ShowItemsWithNoData;

	window['Asc']['ST_slicerStyleType'] = window['Asc'].ST_slicerStyleType = ST_slicerStyleType;
	prot = ST_slicerStyleType;
	prot['unselectedItemWithData'] = prot.unselectedItemWithData;
	prot['selectedItemWithData'] = prot.selectedItemWithData;
	prot['unselectedItemWithNoData'] = prot.unselectedItemWithNoData;
	prot['selectedItemWithNoData'] = prot.selectedItemWithNoData;
	prot['hoveredUnselectedItemWithData'] = prot.hoveredUnselectedItemWithData;
	prot['hoveredSelectedItemWithData'] = prot.hoveredSelectedItemWithData;
	prot['hoveredUnselectedItemWithNoData'] = prot.hoveredUnselectedItemWithNoData;
	prot['hoveredSelectedItemWithNoData'] = prot.hoveredSelectedItemWithNoData;

	window['AscCommonExcel'].insertSlicerType = insertSlicerType;
	prot = insertSlicerType;
	prot['table'] = prot.table;
	prot['pivotTable'] = prot.pivotTable;

})(window);
