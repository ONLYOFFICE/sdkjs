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

// Range / cell iterators that merge SheetMemory data and cellStylesByCol
// direct-style entries. Loaded after Workbook.js so AscCommonExcel.Cell /
// Range / Row / RowIterator exist.
//
//   OccupiedRowIterator                   - data + direct-style-only cells in row order.
//   maxStyleOnlyRow                       - bound for occupied iteration.
//   Range._foreachNoEmpty (occupied)      - row-major occupied iteration.
//   Range._foreachDataOnly                - row-major data-only iteration.
//   Range._foreachNoEmptyByCol (occupied) - column-major occupied iteration.
//   Range._foreachDataOnlyByCol           - column-major data-only iteration.
(function (window, undefined) {
	var ns = window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	// Raw tuple iterator. Emits (row, col, type, value, xfIndex[, fHandle])
	// without allocating a Cell or pushing workbook.loadCells. Visitor
	// returning non-null stops iteration (matches _foreachDataOnly).
	// Split into 5- and 6-arg specializations so the hot call site stays
	// monomorphic — V8 deopts when a single site sees both arities.
	function _fedv5(bbox, ws, excludeHidden, visitor) {
		var cellsByCol = ws.cellsByCol;
		var rowsData = ws.rowsData;
		var rdInt32 = rowsData ? rowsData.dataInt32 : null;
		var rdIdxA = rowsData ? rowsData.indexA : 0;
		var rdStruct = rowsData ? rowsData.structSize : 0;
		var c1 = bbox.c1;
		var c2 = Math.min(bbox.c2, cellsByCol.length - 1);
		var r1 = bbox.r1;
		var r2 = bbox.r2;
		var aCols = [], aDatas = [];
		var effMin = Number.MAX_SAFE_INTEGER, effMax = -1;
		for (var ci = c1; ci <= c2; ci++) {
			var cd = cellsByCol[ci];
			if (cd && r1 <= cd.getMaxIndex() && cd.getMinIndex() <= r2) {
				aCols.push(ci); aDatas.push(cd);
				var mn = cd.getMinIndex();
				var mx = cd.getMaxIndex();
				if (mn < effMin) effMin = mn;
				if (mx > effMax) effMax = mx;
			}
		}
		var aLen = aCols.length;
		if (aLen === 0) return undefined;
		// Clip the outer row loop to the active-column extent.
		var startR = r1 > effMin ? r1 : effMin;
		var stopR = r2 < effMax ? r2 : effMax;
		for (var r = startR; r <= stopR; r++) {
			if (excludeHidden && rdInt32) {
				var rowFlags = rdInt32[((r - rdIdxA) * rdStruct) >> 2] & 0xFF;
				if (rowFlags & 0x02) continue;
			}
			for (var i = 0; i < aLen; i++) {
				var colData = aDatas[i];
				if (!colData.isNonEmpty(r)) continue;
				var col = aCols[i];
				var byteOff = (r - colData.indexA) * colData.structSize;
				var int32Off = byteOff >> 2;
				var mix = colData.dataInt32[int32Off];
				var flagsHi = (mix >>> 24) & 0xFF;
				var type = (flagsHi >>> 1) & 0x3;
				var xfIndex = mix & 0xFFFFFF;
				var value;
				if (flagsHi & 0x08) {
					value = colData.dataFloat[(byteOff + 8) >> 3];
				} else {
					value = colData.dataInt32[(byteOff + 8) >> 2];
				}
				var stop = visitor(r, col, type, value, xfIndex);
				if (stop != null) return stop;
			}
		}
		return undefined;
	}

	function _fedv6(bbox, ws, excludeHidden, visitor) {
		var cellsByCol = ws.cellsByCol;
		var rowsData = ws.rowsData;
		var rdInt32 = rowsData ? rowsData.dataInt32 : null;
		var rdIdxA = rowsData ? rowsData.indexA : 0;
		var rdStruct = rowsData ? rowsData.structSize : 0;
		var workbookFormulas = ws.workbook.workbookFormulas;
		var c1 = bbox.c1;
		var c2 = Math.min(bbox.c2, cellsByCol.length - 1);
		var r1 = bbox.r1;
		var r2 = bbox.r2;
		var aCols = [], aDatas = [];
		var effMin = Number.MAX_SAFE_INTEGER, effMax = -1;
		for (var ci = c1; ci <= c2; ci++) {
			var cd = cellsByCol[ci];
			if (cd && r1 <= cd.getMaxIndex() && cd.getMinIndex() <= r2) {
				aCols.push(ci); aDatas.push(cd);
				var mn = cd.getMinIndex();
				var mx = cd.getMaxIndex();
				if (mn < effMin) effMin = mn;
				if (mx > effMax) effMax = mx;
			}
		}
		var aLen = aCols.length;
		if (aLen === 0) return undefined;
		var startR = r1 > effMin ? r1 : effMin;
		var stopR = r2 < effMax ? r2 : effMax;
		for (var r = startR; r <= stopR; r++) {
			if (excludeHidden && rdInt32) {
				var rowFlags = rdInt32[((r - rdIdxA) * rdStruct) >> 2] & 0xFF;
				if (rowFlags & 0x02) continue;
			}
			for (var i = 0; i < aLen; i++) {
				var colData = aDatas[i];
				if (!colData.isNonEmpty(r)) continue;
				var col = aCols[i];
				var byteOff = (r - colData.indexA) * colData.structSize;
				var int32Off = byteOff >> 2;
				var mix = colData.dataInt32[int32Off];
				var flagsHi = (mix >>> 24) & 0xFF;
				var type = (flagsHi >>> 1) & 0x3;
				var xfIndex = mix & 0xFFFFFF;
				var formulaIndex = colData.dataInt32[int32Off + 1];
				var value;
				if (flagsHi & 0x08) {
					value = colData.dataFloat[(byteOff + 8) >> 3];
				} else {
					value = colData.dataInt32[(byteOff + 8) >> 2];
				}
				var fHandle = (formulaIndex && workbookFormulas) ? workbookFormulas.get(formulaIndex) : null;
				var stop = visitor(r, col, type, value, xfIndex, fHandle);
				if (stop != null) return stop;
			}
		}
		return undefined;
	}

	function forEachDataValue(range, visitor, opts) {
		opts = opts || {};
		var ws = range.worksheet;
		var excludeHidden = !!(ws.bExcludeHiddenRows || opts.excludeHiddenRows);
		return opts.includeFormulaHandle
			? _fedv6(range.bbox, ws, excludeHidden, visitor)
			: _fedv5(range.bbox, ws, excludeHidden, visitor);
	}

	ns.forEachDataValue = forEachDataValue;
	ns._forEachDataValue5 = _fedv5;
	ns._forEachDataValue6 = _fedv6;

	// Read-only row cursor: init snapshots active columns once, setRow swaps
	// the current row, read/readInto/get* are O(1) per (row, col) probe.
	// Does NOT push workbook.loadCells and does NOT allocate a Cell — not
	// usable for mutation, history-aware writes, or formula rebinding.
	function RowCursor() {
		this.cellsByCol = null;
		this.c1 = 0;
		this.c2 = 0;
		this.row = -1;
		this._cdByCol = null;
		this._effMinR = 0;
		this._effMaxR = -1;
	}
	RowCursor.prototype.init = function (cellsByCol, c1, c2) {
		this.cellsByCol = cellsByCol;
		this.c1 = c1;
		var cap = Math.min(c2, cellsByCol.length - 1);
		this.c2 = cap;
		var sparse = new Array(cap - c1 + 1);
		var effMin = Number.MAX_SAFE_INTEGER, effMax = -1;
		for (var c = c1; c <= cap; c++) {
			var cd = cellsByCol[c];
			if (cd) {
				sparse[c - c1] = cd;
				var mn = cd.getMinIndex();
				var mx = cd.getMaxIndex();
				if (mn < effMin) effMin = mn;
				if (mx > effMax) effMax = mx;
			}
		}
		this._cdByCol = sparse;
		this._effMinR = effMin;
		this._effMaxR = effMax;
		return this;
	};
	RowCursor.prototype.setRow = function (r) { this.row = r; };
	RowCursor.prototype.isOccupied = function (c) {
		if (this.row < this._effMinR || this.row > this._effMaxR) return false;
		var cd = this._cdByCol[c - this.c1];
		return cd ? cd.isNonEmpty(this.row) : false;
	};
	RowCursor.prototype.read = function (c) {
		if (this.row < this._effMinR || this.row > this._effMaxR) return null;
		var cd = this._cdByCol[c - this.c1];
		if (!cd || !cd.isNonEmpty(this.row)) return null;
		var byteOff = (this.row - cd.indexA) * cd.structSize;
		var int32Off = byteOff >> 2;
		var mix = cd.dataInt32[int32Off];
		var flagsHi = (mix >>> 24) & 0xFF;
		return {
			type: (flagsHi >>> 1) & 0x3,
			xfIndex: mix & 0xFFFFFF,
			value: (flagsHi & 0x08)
				? cd.dataFloat[(byteOff + 8) >> 3]
				: cd.dataInt32[(byteOff + 8) >> 2]
		};
	};
	// Allocation-free variant. The out parameter shape must stay stable
	// across calls; recommended pre-init: { type:0, xfIndex:0, value:0, isOccupied:false }.
	RowCursor.prototype.readInto = function (c, out) {
		if (this.row < this._effMinR || this.row > this._effMaxR) {
			out.isOccupied = false;
			return false;
		}
		var cd = this._cdByCol[c - this.c1];
		if (!cd || !cd.isNonEmpty(this.row)) {
			out.isOccupied = false;
			return false;
		}
		var byteOff = (this.row - cd.indexA) * cd.structSize;
		var int32Off = byteOff >> 2;
		var mix = cd.dataInt32[int32Off];
		var flagsHi = (mix >>> 24) & 0xFF;
		out.type = (flagsHi >>> 1) & 0x3;
		out.xfIndex = mix & 0xFFFFFF;
		out.value = (flagsHi & 0x08)
			? cd.dataFloat[(byteOff + 8) >> 3]
			: cd.dataInt32[(byteOff + 8) >> 2];
		out.isOccupied = true;
		return true;
	};
	// Single-field accessors. Each returns the field if (row, c) is occupied,
	// else `undefined`. No allocation.
	RowCursor.prototype.getType = function (c) {
		if (this.row < this._effMinR || this.row > this._effMaxR) return undefined;
		var cd = this._cdByCol[c - this.c1];
		if (!cd || !cd.isNonEmpty(this.row)) return undefined;
		var byteOff = (this.row - cd.indexA) * cd.structSize;
		var mix = cd.dataInt32[byteOff >> 2];
		return ((mix >>> 24) >>> 1) & 0x3;
	};
	RowCursor.prototype.getXfIndex = function (c) {
		if (this.row < this._effMinR || this.row > this._effMaxR) return undefined;
		var cd = this._cdByCol[c - this.c1];
		if (!cd || !cd.isNonEmpty(this.row)) return undefined;
		var byteOff = (this.row - cd.indexA) * cd.structSize;
		return cd.dataInt32[byteOff >> 2] & 0xFFFFFF;
	};
	RowCursor.prototype.getValue = function (c) {
		if (this.row < this._effMinR || this.row > this._effMaxR) return undefined;
		var cd = this._cdByCol[c - this.c1];
		if (!cd || !cd.isNonEmpty(this.row)) return undefined;
		var byteOff = (this.row - cd.indexA) * cd.structSize;
		var mix = cd.dataInt32[byteOff >> 2];
		return ((mix >>> 24) & 0x08)
			? cd.dataFloat[(byteOff + 8) >> 3]
			: cd.dataInt32[(byteOff + 8) >> 2];
	};

	ns.RowCursor = RowCursor;

	// Largest row index with any direct style entry in [c1..c2], or -1.
	function maxStyleOnlyRow(ws, c1, c2) {
		var stylesByCol = ws.cellStylesByCol;
		if (!stylesByCol || stylesByCol.length === 0) {
			return -1;
		}
		var maxC = stylesByCol.length - 1;
		if (maxC > c2) {
			maxC = c2;
		}
		var maxR = -1;
		for (var c = c1; c <= maxC; c++) {
			var store = stylesByCol[c];
			if (!store || store.isEmpty()) {
				continue;
			}
			var last = store.lastRow();
			if (last > maxR) {
				maxR = last;
			}
		}
		return maxR;
	}

	// Yields data and style-only cells in column order. Style-only
	// emissions use a shared read-only transient Cell. Data+style at the
	// same coord emits once through the data path (Cell.loadContent
	// resolves the direct xf).
	function OccupiedRowIterator() {
	}
	// r2 is optional; forwarded to the inner RowIterator so the adaptive
	// active-extent clip is correct when adaptive mode is active.
	OccupiedRowIterator.prototype.init = function (ws, r1, c1, c2, r2) {
		this.ws = ws;
		this.c1 = c1;
		var dataMaxC = (ws.cellsByCol && ws.cellsByCol.length > 0) ? ws.cellsByCol.length - 1 : -1;
		var stylesByCol = ws.cellStylesByCol || [];
		var styleMaxC = stylesByCol.length - 1;
		var effectiveC2 = dataMaxC > styleMaxC ? dataMaxC : styleMaxC;
		if (c2 > effectiveC2) {
			c2 = effectiveC2;
		}
		this.c2 = c2;
		this.dataIter = null;
		if (c2 >= c1) {
			this.dataIter = new ns.RowIterator();
			this.dataIter.init(ws, r1, c1, c2, r2);
		}
		this.transientCell = new ns.Cell(ws);
		this.transientCell._isTransient = true;
		this.row = r1 - 1;
		// Snapshot the columns that carry any style entry in [c1..c2] once,
		// so per-row setRow does not re-walk the full cellStylesByCol array.
		this._activeStoreCols = null;
		this._activeStores = null;
		var maxStoreC = stylesByCol.length - 1;
		if (maxStoreC > c2) {
			maxStoreC = c2;
		}
		for (var sc = c1; sc <= maxStoreC; sc++) {
			var s = stylesByCol[sc];
			if (s && !s.isEmpty()) {
				if (this._activeStoreCols === null) {
					this._activeStoreCols = [];
					this._activeStores = [];
				}
				this._activeStoreCols.push(sc);
				this._activeStores.push(s);
			}
		}
		this._styleCols = null;
		this._styleColsIdx = 0;
		this._pendingData = null;
		this._pendingDataCol = -1;
		this._dataExhausted = (this.dataIter === null);
	};
	OccupiedRowIterator.prototype.release = function () {
		if (this.dataIter) {
			this.dataIter.release();
			this.dataIter = null;
		}
	};
	// Combined data + style row extent for the _foreachNoEmpty outer-loop clip.
	// Returns -1 when the clip is not safe: dataIter.getEffMinR() < 0 means
	// baseline mode (SweepLine doesn't publish an extent), so we cannot
	// know which rows have data. The caller's gate `effMin/effMax >= 0`
	// then keeps the legacy r1..minR walk.
	OccupiedRowIterator.prototype.getEffMinR = function () {
		if (!this.dataIter) return -1;
		var dataMin = this.dataIter.getEffMinR();
		if (dataMin < 0) return -1;
		var dataKnown = this.dataIter.getEffMaxR() >= 0;
		var styleMin = -1;
		if (this._activeStores) {
			for (var i = 0; i < this._activeStores.length; i++) {
				var s = this._activeStores[i];
				var f = s.firstRow();
				if (f >= 0 && (styleMin === -1 || f < styleMin)) styleMin = f;
			}
		}
		if (!dataKnown) return styleMin;
		if (styleMin < 0) return dataMin;
		return dataMin < styleMin ? dataMin : styleMin;
	};
	OccupiedRowIterator.prototype.getEffMaxR = function () {
		if (!this.dataIter) return -1;
		var dataMin = this.dataIter.getEffMinR();
		if (dataMin < 0) return -1;
		var dataMax = this.dataIter.getEffMaxR();
		var styleMax = -1;
		if (this._activeStores) {
			for (var i = 0; i < this._activeStores.length; i++) {
				var s = this._activeStores[i];
				var l = s.lastRow();
				if (l > styleMax) styleMax = l;
			}
		}
		if (dataMax < 0) return styleMax;
		if (styleMax < 0) return dataMax;
		return dataMax > styleMax ? dataMax : styleMax;
	};
	OccupiedRowIterator.prototype.setRow = function (row) {
		this.row = row;
		if (this.dataIter) {
			this.dataIter.setRow(row);
			this._dataExhausted = false;
		}
		this._pendingData = null;
		this._pendingDataCol = -1;
		this._styleCols = this._computeStyleOnlyCols(row);
		this._styleColsIdx = 0;
	};
	OccupiedRowIterator.prototype._computeStyleOnlyCols = function (row) {
		var stores = this._activeStores;
		if (!stores) {
			return null;
		}
		var cols = null;
		for (var i = 0; i < stores.length; i++) {
			var xf = stores[i].get(row);
			if (xf == null || xf === 0) {
				continue;
			}
			if (cols === null) {
				cols = [];
			}
			cols.push(this._activeStoreCols[i]);
		}
		return cols;
	};
	OccupiedRowIterator.prototype._peekData = function () {
		if (this._pendingData) {
			return this._pendingData;
		}
		if (this._dataExhausted || !this.dataIter) {
			return null;
		}
		var cell = this.dataIter.next();
		if (!cell) {
			this._dataExhausted = true;
			return null;
		}
		this._pendingData = cell;
		this._pendingDataCol = cell.nCol;
		return cell;
	};
	OccupiedRowIterator.prototype._emitStyleOnly = function (col) {
		var c = this.transientCell;
		c.clear();
		c.nRow = this.row;
		c.nCol = col;
		c._isTransient = true;
		var xfs = this.ws._directOrInheritedXfs(this.row, col);
		if (xfs) {
			c.xfs = xfs;
		}
		return c;
	};
	OccupiedRowIterator.prototype.next = function () {
		while (true) {
			var dataCell = this._peekData();
			var styleColsLen = this._styleCols ? this._styleCols.length : 0;
			var styleCol = (this._styleColsIdx < styleColsLen) ? this._styleCols[this._styleColsIdx] : -1;
			if (!dataCell && styleCol < 0) {
				return null;
			}
			if (!dataCell) {
				this._styleColsIdx++;
				return this._emitStyleOnly(styleCol);
			}
			if (styleCol < 0 || this._pendingDataCol < styleCol) {
				this._pendingData = null;
				return dataCell;
			}
			if (this._pendingDataCol === styleCol) {
				// data + style at same col: emit once via data path.
				this._styleColsIdx++;
				this._pendingData = null;
				return dataCell;
			}
			this._styleColsIdx++;
			return this._emitStyleOnly(styleCol);
		}
	};

	ns.OccupiedRowIterator = OccupiedRowIterator;
	ns.maxStyleOnlyRow = maxStyleOnlyRow;

	var Range = ns.Range;
	var Cell = ns.Cell;
	var Row = ns.Row;
	var RowIterator = ns.RowIterator;

	// Range-level entry points for the raw tuple iterator and row cursor.
	// Both are additive and don't alter _foreach* semantics.
	Range.prototype.forEachDataValue = function(visitor, opts) {
		return ns.forEachDataValue(this, visitor, opts);
	};
	Range.prototype.createRowCursor = function () {
		return new RowCursor().init(this.worksheet.cellsByCol, this.bbox.c1, this.bbox.c2);
	};

	// Occupied iteration: data cells AND direct-style-only cells.
	// Style-only cells arrive as a shared _isTransient Cell with cell.xfs
	// preloaded; callers must treat them as read-only (no saveContent,
	// clearData, setStyle, _removeCell). Use _foreachDataOnly when only
	// SheetMemory init rows should be visited.
	Range.prototype._foreachNoEmpty = function(actionCell, actionRow, excludeHiddenRows) {
		var oRes, i, oBBox = this.bbox;
		var ws = this.worksheet;
		var dataMaxR = Math.max(ws.cellsByColRowsCount - 1, ws.rowsData.getMaxIndex());
		var styleMaxR = maxStyleOnlyRow(ws, oBBox.c1, oBBox.c2);
		var minR = Math.max(dataMaxR, styleMaxR);
		minR = Math.min(minR, oBBox.r2);
		if (actionCell || actionRow) {
			var itRow = null;
			if (actionCell) {
				itRow = new OccupiedRowIterator();
				itRow.init(ws, oBBox.r1, oBBox.c1, oBBox.c2, oBBox.r2);
			}
			// Outer-loop clip. Only safe when actionRow is null: an actionRow
			// callback may read excludedCount and may rely on per-row row
			// metadata over the full bbox, so we keep the legacy walk there.
			var startR = oBBox.r1;
			var stopR = minR;
			if (itRow && !actionRow && itRow.getEffMinR) {
				var effMin = itRow.getEffMinR();
				var effMax = itRow.getEffMaxR();
				if (effMin >= 0 && effMax >= 0) {
					if (effMin > startR) startR = effMin;
					if (effMax < stopR) stopR = effMax;
				}
			}
			var bExcludeHiddenRows = (ws.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell;
			var tempRow = new Row(ws);
			var allRow = ws.getAllRow();
			var allRowHidden = allRow && allRow.getHidden();
			for (i = startR; i <= stopR; i++) {
				if (actionRow) {
					if (tempRow.loadContent(i)) {
						if (bExcludeHiddenRows && tempRow.getHidden()) {
							excludedCount++;
							continue;
						}
						oRes = actionRow(tempRow, excludedCount);
						tempRow.saveContent(true);
						if (null != oRes) {
							if (itRow) {
								itRow.release();
							}
							return oRes;
						}
					} else if (bExcludeHiddenRows && allRowHidden) {
						excludedCount++;
						continue;
					}
				} else if (bExcludeHiddenRows && ws.getRowHidden(i)) {
					excludedCount++;
					continue;
				}
				if (itRow) {
					itRow.setRow(i);
					while (tempCell = itRow.next()) {
						oRes = actionCell(tempCell, i, tempCell.nCol, oBBox.r1, oBBox.c1, excludedCount);
						if (null != oRes) {
							itRow.release();
							return oRes;
						}
					}
				}
			}
			if (itRow) {
				itRow.release();
			}
		}
	};
	// Data-only row-major iteration; no transient style-only emissions.
	Range.prototype._foreachDataOnly = function(actionCell, actionRow, excludeHiddenRows) {
		var oRes, i, oBBox = this.bbox, minR = Math.max(this.worksheet.cellsByColRowsCount - 1, this.worksheet.rowsData.getMaxIndex());
		minR = Math.min(minR, oBBox.r2);
		if (actionCell || actionRow) {
			var itRow = new RowIterator();
			if (actionCell) {
				itRow.init(this.worksheet, this.bbox.r1, this.bbox.c1, this.bbox.c2, this.bbox.r2);
			}
			// Outer-loop clip to the active-column extent. Skipped when
			// actionRow is provided so the row callback still sees every row
			// in [r1..minR] (excludedCount accuracy + per-row side effects).
			var startR = oBBox.r1;
			var stopR = minR;
			if (actionCell && !actionRow && itRow.getEffMinR) {
				var effMin = itRow.getEffMinR();
				var effMax = itRow.getEffMaxR();
				if (effMin >= 0 && effMax >= 0) {
					if (effMin > startR) startR = effMin;
					if (effMax < stopR) stopR = effMax;
				}
			}
			var bExcludeHiddenRows = (this.worksheet.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell;
			var tempRow = new Row(this.worksheet);
			var allRow = this.worksheet.getAllRow();
			var allRowHidden = allRow && allRow.getHidden();
			for (i = startR; i <= stopR; i++) {
				if (actionRow) {
					if (tempRow.loadContent(i)) {
						if (bExcludeHiddenRows && tempRow.getHidden()) {
							excludedCount++;
							continue;
						}
						oRes = actionRow(tempRow, excludedCount);
						tempRow.saveContent(true);
						if (null != oRes) {
							if (actionCell) {
								itRow.release();
							}
							return oRes;
						}
					} else if (bExcludeHiddenRows && allRowHidden) {
						excludedCount++;
						continue;
					}
				} else if (bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
					excludedCount++;
					continue;
				}
				if (actionCell) {
					itRow.setRow(i);
					while (tempCell = itRow.next()) {
						oRes = actionCell(tempCell, i, tempCell.nCol, oBBox.r1, oBBox.c1, excludedCount);
						if (null != oRes) {
							if (actionCell) {
								itRow.release();
							}
							return oRes;
						}
					}
				}
			}
			if (actionCell) {
				itRow.release();
			}
		}
	};
	// Column-major occupied (data + direct-style-only) iteration; transient
	// Cell emissions are read-only.
	Range.prototype._foreachNoEmptyByCol = function(actionCell, excludeHiddenRows) {
		if (!actionCell) {
			return;
		}
		var oRes, i, j, colData;
		var ws = this.worksheet;
		var wb = ws.workbook;
		var oBBox = this.bbox;
		var dataColMax = ws.getColDataLength() - 1;
		var stylesByCol = ws.cellStylesByCol || [];
		var styleColMax = stylesByCol.length - 1;
		var minC = oBBox.c2;
		var bigC = dataColMax > styleColMax ? dataColMax : styleColMax;
		if (minC > bigC) {
			minC = bigC;
		}
		if (oBBox.c1 > minC) {
			return;
		}
		var bExcludeHiddenRows = (ws.bExcludeHiddenRows || excludeHiddenRows);
		var excludedCount = 0;
		var tempCell = new Cell(ws);
		tempCell._isTransient = true;
		var transientStyleCell = new Cell(ws);
		transientStyleCell._isTransient = true;
		wb.loadCells.push(tempCell);
		for (j = oBBox.c1; j <= minC; ++j) {
			colData = ws.getColDataNoEmpty(j);
			var styleStore = stylesByCol[j] || null;
			if (!colData && (!styleStore || styleStore.isEmpty())) {
				continue;
			}
			var dataMaxR = colData ? colData.getMaxIndex() : -1;
			var styleMaxR = (styleStore && !styleStore.isEmpty()) ? styleStore.lastRow() : -1;
			var maxR = dataMaxR > styleMaxR ? dataMaxR : styleMaxR;
			var loopMaxR = Math.min(oBBox.r2, maxR);
			for (i = oBBox.r1; i <= loopMaxR; i++) {
				if (bExcludeHiddenRows && ws.getRowHidden(i)) {
					excludedCount++;
					continue;
				}
				var targetCell = null;
				for (var k = 0; k < wb.loadCells.length - 1; ++k) {
					var elem = wb.loadCells[k];
					if (elem.nRow === i && elem.nCol === j && ws === elem.ws) {
						targetCell = elem;
						break;
					}
				}
				if (null !== targetCell) {
					oRes = actionCell(targetCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
					if (null != oRes) {
						wb.loadCells.pop();
						return oRes;
					}
					continue;
				}
				var loaded = false;
				if (colData) {
					loaded = tempCell.loadContent(i, j, colData);
				}
				if (loaded) {
					oRes = actionCell(tempCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
					tempCell.saveContent(true);
					if (null != oRes) {
						wb.loadCells.pop();
						return oRes;
					}
					continue;
				}
				if (styleStore) {
					var styleXf = styleStore.get(i);
					if (styleXf != null && styleXf !== 0) {
						transientStyleCell.clear();
						transientStyleCell.nRow = i;
						transientStyleCell.nCol = j;
						transientStyleCell._isTransient = true;
						var xfs = ws._directOrInheritedXfs(i, j);
						if (xfs) {
							transientStyleCell.xfs = xfs;
						}
						oRes = actionCell(transientStyleCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
						if (null != oRes) {
							wb.loadCells.pop();
							return oRes;
						}
					}
				}
			}
		}
		wb.loadCells.pop();
	};
	// Column-major data-only iteration; direct-style-only cells are skipped.
	Range.prototype._foreachDataOnlyByCol = function(actionCell, excludeHiddenRows) {
		var oRes, i, j, colData;
		var wb = this.worksheet.workbook;
		var oBBox = this.bbox, minR = Math.min(this.worksheet.cellsByColRowsCount - 1, oBBox.r2);
		var minC = Math.min(this.worksheet.getColDataLength() - 1, oBBox.c2);
		if (actionCell && oBBox.c1 <= minC && oBBox.r1 <= minR) {
			var bExcludeHiddenRows = (this.worksheet.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell = new Cell(this.worksheet);
			tempCell._isTransient = true;
			wb.loadCells.push(tempCell);
			for (j = oBBox.c1; j <= minC; ++j) {
				colData = this.worksheet.getColDataNoEmpty(j);
				if (colData) {
					// Clip the inner loop to the column's [getMinIndex..getMaxIndex].
					// Callers do not rely on excludedCount accumulating over
					// the empty prefix; the (cell, row, col) signature audit
					// confirmed this is safe.
					var startRow = oBBox.r1;
					var colMin = colData.getMinIndex();
					if (colMin > startRow) startRow = colMin;
					for (i = startRow; i <= Math.min(minR, colData.getMaxIndex()); i++) {
						if (bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
							excludedCount++;
							continue;
						}
						var targetCell = null;
						for (var k = 0; k < wb.loadCells.length - 1; ++k) {
							var elem = wb.loadCells[k];
							if (elem.nRow == i && elem.nCol == j && this.worksheet === elem.ws) {
								targetCell = elem;
								break;
							}
						}
						if (null === targetCell) {
							if (tempCell.loadContent(i, j, colData)) {
								oRes = actionCell(tempCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
								tempCell.saveContent(true);
							}
						} else {
							oRes = actionCell(targetCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
						}
						if (null != oRes) {
							wb.loadCells.pop();
							return oRes;
						}
					}
				}
			}
			wb.loadCells.pop();
		}
	};
})(window);
