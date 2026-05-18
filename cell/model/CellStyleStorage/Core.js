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

// Core read/write for cellStylesByCol direct cell styles. Sibling files
// share the AscCommonExcel.CellStyleStorage namespace; private internals
// for them live on CSS._internals. Helpers here never touch Cell,
// SheetMemory, or History.
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

	// Number / CellXfs -> positive xfIndex (0 == no direct style). An
	// unregistered CellXfs is added to g_StyleCache.
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

	// Lazy creation. Stores start empty (LegacyMigration hydrates first).
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

	// Existing store for `col`, or null (no on-demand hydration).
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

	// Mirror cell.xfs into ws.cellStylesByCol. Called from
	// setStyleInternal / clearDataKeepXf. Skips transient cells.
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

	// Wrappers resolve through CSS.<name> at call time so sibling files
	// may install methods in any load order.
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
	CSS.mirrorCellStyle = mirrorCellStyle;
	CSS.installOnWorksheet = installOnWorksheet;

	// Shared private surface for sibling CellStyleStorage files.
	CSS._internals = CSS._internals || {};
	CSS._internals.DEFAULT_MAX_ROW = DEFAULT_MAX_ROW;
	CSS._internals.getCRangeAttrArrayCtor = _getCRangeAttrArrayCtor;
	CSS._internals.isDataInitRow = _isDataInitRow;
	CSS._internals.resolveSourceStore = _resolveSourceStore;
	CSS._internals.getCellStyleStore = getCellStyleStore;
})(window);
