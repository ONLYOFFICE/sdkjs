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

// Core read/write for cellStylesByCol direct cell styles.
//
// Sibling files in this folder share one runtime namespace
// `window.AscCommonExcel.CellStyleStorage`; this file installs it and
// exposes the private surface via `CSS._internals`. Helpers here and in
// sibling files touch only `ws.cellStylesByCol[col]` (a CRangeAttrArray)
// -- never Cell, SheetMemory, or History -- so the load/save, range ops,
// and serializer paths all share one source of truth without re-entering
// Cell materialization.
(function (window, undefined) {
	var DEFAULT_MAX_ROW = (window['AscCommon'] && window['AscCommon'].gc_nMaxRow0 != null)
		? window['AscCommon'].gc_nMaxRow0
		: 1048575;

	// SheetMemory word 0 packs `(cellFlag << 24) | xfIndex` (see Cell._toFlags).
	// We read only the init flag in bit 0 of the high byte.
	var _CELL_FLAG_INIT_SHIFT = 24;

	function _getCRangeAttrArrayCtor() {
		return window['AscCommonExcel'].CRangeAttrArray;
	}

	function _isDataInitRow(sheetMemory, row) {
		var mix = sheetMemory.getInt32(row, 0);
		return ((mix >>> _CELL_FLAG_INIT_SHIFT) & 1) !== 0;
	}

	// Normalize a number-or-CellXfs into a positive xfIndex, returning 0
	// for "no direct cell style" (matches Cell.loadContent, which only
	// restores style when the saved index is > 0). An unregistered CellXfs
	// is added to g_StyleCache so a fresh builder-side xf is not dropped.
	function _toXfIndex(xfIndexOrXfs) {
		if (xfIndexOrXfs == null) {
			return 0;
		}
		if (typeof xfIndexOrXfs === 'number') {
			return xfIndexOrXfs > 0 ? (xfIndexOrXfs | 0) : 0;
		}
		if (typeof xfIndexOrXfs.getIndexNumber === 'function') {
			var idx = xfIndexOrXfs.getIndexNumber();
			if (typeof idx === 'number' && idx > 0) {
				return idx | 0;
			}
		}
		var cache = window['AscCommonExcel'].g_StyleCache;
		if (cache && typeof cache.addXf === 'function') {
			var cached = cache.addXf(xfIndexOrXfs);
			if (cached && typeof cached.getIndexNumber === 'function') {
				var idx2 = cached.getIndexNumber();
				return (typeof idx2 === 'number' && idx2 > 0) ? (idx2 | 0) : 0;
			}
		}
		return 0;
	}

	// Lazy creation only. Legacy SheetMemory xf bits are migrated by the
	// eager file-open sweep (LegacyMigration); a store materialized here
	// starts empty.
	function getCellStyleStore(ws, col, opt_create) {
		if (col < 0) {
			return null;
		}
		var store = ws.cellStylesByCol[col];
		if (!store && opt_create) {
			var Ctor = _getCRangeAttrArrayCtor();
			store = new Ctor(DEFAULT_MAX_ROW);
			ws.cellStylesByCol[col] = store;
		}
		return store || null;
	}

	// Upper bound (exclusive) for column scans that must reach past
	// `cellStylesByCol.length` into data-only columns. Used by
	// `shiftCellXfs` left/right modes.
	function _maxLegacyColLength(ws) {
		if (!ws) {
			return 0;
		}
		if (typeof ws.getColDataLength === 'function') {
			return ws.getColDataLength();
		}
		if (ws.cellsByCol && typeof ws.cellsByCol.length === 'number') {
			return ws.cellsByCol.length;
		}
		return 0;
	}

	// Returns the existing store for `col` or null. No SheetMemory
	// hydration: LegacyMigration has already run for every column at file
	// open. Callers treat null as a no-op source.
	function _resolveSourceStore(ws, col) {
		if (col < 0) {
			return null;
		}
		return ws.cellStylesByCol[col] || null;
	}

	// Returns 0 for "no direct cell style" so callers can compare with the
	// existing Cell.saveContent / Cell.loadContent semantics directly.
	function getCellXf(ws, row, col) {
		if (col < 0) {
			return 0;
		}
		var store = ws.cellStylesByCol[col];
		if (!store) {
			return 0;
		}
		var idx = store.get(row);
		return (idx == null) ? 0 : idx;
	}

	function setCellXf(ws, row, col, xfIndexOrXfs) {
		var idx = _toXfIndex(xfIndexOrXfs);
		if (idx === 0) {
			var existing = _resolveSourceStore(ws, col);
			if (existing) {
				existing.clearRange(row, row);
			}
			return;
		}
		var store = getCellStyleStore(ws, col, true);
		store.setRange(row, row, idx);
	}

	function setCellXfRange(ws, r1, c1, r2, c2, xfIndexOrXfs) {
		if (c2 < c1 || r2 < r1) {
			return;
		}
		var idx = _toXfIndex(xfIndexOrXfs);
		if (idx === 0) {
			for (var c = c1; c <= c2; c++) {
				var existing = _resolveSourceStore(ws, c);
				if (existing) {
					existing.clearRange(r1, r2);
				}
			}
			return;
		}
		for (var c2col = c1; c2col <= c2; c2col++) {
			var store = getCellStyleStore(ws, c2col, true);
			store.setRange(r1, r2, idx);
		}
	}

	function clearCellXfRange(ws, r1, c1, r2, c2) {
		if (c2 < c1 || r2 < r1) {
			return;
		}
		for (var c = c1; c <= c2; c++) {
			var existing = _resolveSourceStore(ws, c);
			if (existing) {
				existing.clearRange(r1, r2);
			}
		}
	}

	// Copy a 2D range of direct cell styles between worksheets (or within one).
	// fromBBox and toBBox are expected to have matching dimensions; if they
	// differ, the smaller dimension is used so the call is bounded.
	function copyCellXfRange(wsTo, wsFrom, fromBBox, toBBox) {
		var fw = fromBBox.c2 - fromBBox.c1 + 1;
		var fh = fromBBox.r2 - fromBBox.r1 + 1;
		var tw = toBBox.c2 - toBBox.c1 + 1;
		var th = toBBox.r2 - toBBox.r1 + 1;
		var w = fw < tw ? fw : tw;
		var h = fh < th ? fh : th;
		if (w <= 0 || h <= 0) {
			return;
		}
		for (var dc = 0; dc < w; dc++) {
			var srcCol = fromBBox.c1 + dc;
			var dstCol = toBBox.c1 + dc;
			// A null source store means no direct style to copy; clear the
			// destination band so it stays consistent with the source.
			var src = _resolveSourceStore(wsFrom, srcCol);
			var dst = wsTo.cellStylesByCol[dstCol];
			if (!src) {
				if (dst) {
					dst.clearRange(toBBox.r1, toBBox.r1 + h - 1);
				}
				continue;
			}
			if (!dst) {
				dst = getCellStyleStore(wsTo, dstCol, true);
			}
			dst.copyFrom(src, fromBBox.r1, toBBox.r1, h);
		}
	}

	// Mirror `cell.xfs` into `ws.cellStylesByCol` so the store stays the
	// source of truth for direct cell xf. Called from Cell.setStyleInternal
	// and Cell.clearDataKeepXf. Skips transient cells, cells without a real
	// (row, col), and cells without a Worksheet host.
	function mirrorCellStyle(cell) {
		if (!cell || cell._isTransient) {
			return;
		}
		if (cell.nRow < 0 || cell.nCol < 0) {
			return;
		}
		var ws = cell.ws;
		if (!ws || !ws.cellStylesByCol) {
			return;
		}
		setCellXf(ws, cell.nRow, cell.nCol, cell.xfs);
	}

	// Wrappers dispatch through `CSS.<name>` at call time so methods
	// defined in sibling files resolve regardless of file load order.
	function installOnWorksheet(WorksheetCtor) {
		WorksheetCtor.prototype.getCellStyleStore = function (col, opt_create) {
			return CSS.getCellStyleStore(this, col, opt_create);
		};
		WorksheetCtor.prototype.getCellXf = function (row, col) {
			return CSS.getCellXf(this, row, col);
		};
		WorksheetCtor.prototype.setCellXf = function (row, col, xfIndexOrXfs) {
			return CSS.setCellXf(this, row, col, xfIndexOrXfs);
		};
		WorksheetCtor.prototype.deleteCellXfRowsAllCols = function (start, count) {
			return CSS.deleteRowsAllCols(this, start, count);
		};
		WorksheetCtor.prototype.insertCellXfRowsAllCols = function (index, count) {
			return CSS.insertRowsAllCols(this, index, count);
		};
		WorksheetCtor.prototype.copyCellXfRowInAllCols = function (fromRow, toRow, toCount) {
			return CSS.copyRowInAllCols(this, fromRow, toRow, toCount);
		};
		WorksheetCtor.prototype.deleteCellXfCols = function (start, count) {
			return CSS.deleteCols(this, start, count);
		};
		WorksheetCtor.prototype.insertCellXfCols = function (index, count, maxCol) {
			return CSS.insertCols(this, index, count, maxCol);
		};
		WorksheetCtor.prototype.moveCellXfBand = function (fromCol, toCol, r1, count, clearSource) {
			return CSS.moveColRowBand(this, fromCol, toCol, r1, count, clearSource);
		};
		WorksheetCtor.prototype.sortCellXfs = function (oBBox, oSortedIndexes, opt_by_row) {
			return CSS.sortCellXfs(this, oBBox, oSortedIndexes, opt_by_row);
		};
	}

	var ns = window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	var CSS = ns.CellStyleStorage = ns.CellStyleStorage || {};
	CSS.getCellStyleStore = getCellStyleStore;
	CSS.getCellXf = getCellXf;
	CSS.setCellXf = setCellXf;
	CSS.setCellXfRange = setCellXfRange;
	CSS.clearCellXfRange = clearCellXfRange;
	CSS.copyCellXfRange = copyCellXfRange;
	CSS.mirrorCellStyle = mirrorCellStyle;
	CSS.installOnWorksheet = installOnWorksheet;

	// Shared private surface for sibling CellStyleStorage files.
	CSS._internals = CSS._internals || {};
	CSS._internals.DEFAULT_MAX_ROW = DEFAULT_MAX_ROW;
	CSS._internals.getCRangeAttrArrayCtor = _getCRangeAttrArrayCtor;
	CSS._internals.isDataInitRow = _isDataInitRow;
	CSS._internals.resolveSourceStore = _resolveSourceStore;
	CSS._internals.maxLegacyColLength = _maxLegacyColLength;
	CSS._internals.getCellStyleStore = getCellStyleStore;
})(window);
