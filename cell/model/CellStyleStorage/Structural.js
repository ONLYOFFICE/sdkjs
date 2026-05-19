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

// Structural cellStylesByCol mutations: row/column shifts and inserts,
// band moves, sort permutations, plus the style-only cleaners that pair
// with Range.cleanFormat/cleanAll and the format-painter/autofill
// promote path. Every helper here pairs with a matching SheetMemory
// mutation in Workbook.js; the two storages move together.
(function (window, undefined) {
	var CSS = window['AscCommonExcel'].CellStyleStorage;
	var _internals = CSS._internals;
	var DEFAULT_MAX_ROW = _internals.DEFAULT_MAX_ROW;
	var _getCRangeAttrArrayCtor = _internals.getCRangeAttrArrayCtor;
	var _resolveSourceStore = _internals.resolveSourceStore;
	var getCellStyleStore = _internals.getCellStyleStore;

	// Drop `count` rows starting at `start` from every column store.
	// Pairs with `_removeRows` / `_shiftCellsUp`.
	function deleteRowsAllCols(ws, start, count) {
		if (count <= 0) {
			return;
		}
		var stores = ws.cellStylesByCol;
		for (var c = 0; c < stores.length; c++) {
			var s = stores[c];
			if (s) {
				s.deleteRows(start, count);
			}
		}
	}

	// Open `count` empty rows at `index` in every column store.
	// Pairs with `_insertRowsBefore` / `_shiftCellsBottom`.
	function insertRowsAllCols(ws, index, count) {
		if (count <= 0) {
			return;
		}
		var stores = ws.cellStylesByCol;
		for (var c = 0; c < stores.length; c++) {
			var s = stores[c];
			if (s) {
				s.insertRows(index, count);
			}
		}
	}

	// Mirror the row-inheritance copy that _insertRowsBefore performs on
	// the data side, so inserted rows pick up the style of the row above.
	function copyRowInAllCols(ws, fromRow, toRow, toCount) {
		if (toCount <= 0 || fromRow < 0) {
			return;
		}
		var stores = ws.cellStylesByCol;
		for (var c = 0; c < stores.length; c++) {
			var s = stores[c];
			if (s) {
				s.setAreaByRow(fromRow, toRow, toCount);
			}
		}
	}

	// Splice out [start..start+count-1] columns to keep cellStylesByCol
	// index-aligned with cellsByCol after `_removeCols`.
	function deleteCols(ws, start, count) {
		if (count <= 0) {
			return;
		}
		ws.cellStylesByCol.splice(start, count);
	}

	// Open `count` empty column slots before `index` to keep
	// cellStylesByCol index-aligned with cellsByCol after
	// `_insertColsBefore`. Trailing entries beyond `maxCol` are dropped
	// to mirror `cellsByCol.splice(maxCol - count + 1, count)`.
	function insertCols(ws, index, count, maxCol) {
		if (count <= 0) {
			return;
		}
		var stores = ws.cellStylesByCol;
		if (maxCol != null) {
			var dropFrom = maxCol - count + 1;
			if (dropFrom < stores.length) {
				stores.splice(dropFrom, stores.length - dropFrom);
			}
		}
		for (var i = stores.length - 1; i >= index; --i) {
			stores[i + count] = stores[i];
			stores[i] = undefined;
		}
	}

	// Move a row band [r1..r1+count-1] between columns of one worksheet,
	// optionally clearing the source band. Destination is materialized
	// lazily; a null source means there is no direct style to move.
	function moveColRowBand(ws, fromCol, toCol, r1, count, clearSource) {
		if (count <= 0) {
			return;
		}
		var src = _resolveSourceStore(ws, fromCol);
		var dst = ws.cellStylesByCol[toCol];
		if (!src) {
			if (dst) {
				dst.clearRange(r1, r1 + count - 1);
			}
			return;
		}
		if (!dst) {
			dst = getCellStyleStore(ws, toCol, true);
		}
		dst.copyFrom(src, r1, r1, count);
		if (clearSource && fromCol !== toCol) {
			src.clearRange(r1, r1 + count - 1);
		}
	}

	// Mirror _moveCells on cellStylesByCol. SheetMemory no longer owns
	// the xf, so this is the only path that preserves locked-only xfs on
	// the source band of a protected-sheet move. `clearStart`/`clearEnd`
	// are inclusive-exclusive; pass equal values to skip the source clear.
	// Lazily materializes a source store from SheetMemory presence so
	// data-only source columns clear properly.
	function moveCellsBetweenWorksheets(wsFrom, wsTo, fromCol, toCol, r1From, r1To,
										count, clearStart, clearEnd, getLockedOnlyXfIndex) {
		if (count <= 0) {
			return;
		}
		var srcStore = wsFrom.cellStylesByCol[fromCol];
		if (!srcStore && typeof wsFrom.getColDataNoEmpty === 'function'
			&& wsFrom.getColDataNoEmpty(fromCol)) {
			srcStore = getCellStyleStore(wsFrom, fromCol, true);
		}
		var dstStore = wsTo.cellStylesByCol[toCol];
		if (srcStore) {
			if (!dstStore) {
				dstStore = getCellStyleStore(wsTo, toCol, true);
			}
			dstStore.copyFrom(srcStore, r1From, r1To, count);
		} else if (dstStore) {
			dstStore.clearRange(r1To, r1To + count - 1);
		}
		if (!srcStore || clearEnd <= clearStart) {
			return;
		}
		if (getLockedOnlyXfIndex) {
			srcStore.mapRange(clearStart, clearEnd - 1, function (v) {
				if (v == null || v === 0) {
					return null;
				}
				var locked = getLockedOnlyXfIndex(v);
				return (locked != null && locked > 0) ? locked : null;
			});
		} else {
			srcStore.clearRange(clearStart, clearEnd - 1);
		}
	}

	// Mirror Range._sortByArray's permutation on cellStylesByCol. Must
	// run BEFORE the SheetMemory mutation. Horizontal sort uses a
	// destination-band snapshot so cycles resolve regardless of order.
	function sortCellXfs(ws, oBBox, oSortedIndexes, opt_by_row) {
		if (!ws || !oBBox || !oSortedIndexes) {
			return;
		}
		var height = oBBox.r2 - oBBox.r1 + 1;
		if (height <= 0) {
			return;
		}
		var Ctor = _getCRangeAttrArrayCtor();

		if (!opt_by_row) {
			for (var c = oBBox.c1; c <= oBBox.c2; c++) {
				var store = _resolveSourceStore(ws, c);
				if (!store) {
					continue;
				}
				var temp = new Ctor(store.maxRow);
				temp.copyFrom(store, oBBox.r1, oBBox.r1, height);
				for (var fk in oSortedIndexes) {
					if (!Object.prototype.hasOwnProperty.call(oSortedIndexes, fk)) {
						continue;
					}
					temp.copyFrom(store, fk | 0, oSortedIndexes[fk] | 0, 1);
				}
				store.copyFrom(temp, oBBox.r1, oBBox.r1, height);
			}
			return;
		}

		var snapshots = {};
		for (var key in oSortedIndexes) {
			if (!Object.prototype.hasOwnProperty.call(oSortedIndexes, key)) {
				continue;
			}
			var from = key | 0;
			var to = oSortedIndexes[key] | 0;

			var storeFrom = _resolveSourceStore(ws, from);
			var storeTo = _resolveSourceStore(ws, to);

			// Snapshot the destination band BEFORE mutation so a later
			// iteration reading `to` as a source sees the pre-shift state.
			var snapTo = new Ctor(storeTo ? storeTo.maxRow : DEFAULT_MAX_ROW);
			if (storeTo) {
				snapTo.copyFrom(storeTo, oBBox.r1, oBBox.r1, height);
			}
			snapshots[to] = snapTo;

			var src;
			if (Object.prototype.hasOwnProperty.call(snapshots, from)) {
				src = snapshots[from];
			} else {
				src = storeFrom;
			}

			if (src) {
				if (!storeTo) {
					storeTo = getCellStyleStore(ws, to, true);
				}
				storeTo.copyFrom(src, oBBox.r1, oBBox.r1, height);
			} else if (storeTo) {
				storeTo.clearRange(oBBox.r1, oBBox.r1 + height - 1);
			}
		}
	}

	// Clear direct styles for cells that exist only in `cellStylesByCol`.
	// Emits `historyitem_Cell_SetStyleOnly` so undo/redo replays through
	// `ws.setCellXf`, avoiding Cell allocation / SheetMemory init.
	function cleanStyleOnlyDirectStyles(ws, bbox, opt_excludeHiddenRows) {
		if (!ws || !ws.cellStylesByCol || !bbox) {
			return;
		}
		if (typeof CSS.forEachStyleOnlyCell !== 'function') {
			return;
		}
		var AscCommon = window['AscCommon'];
		var AscCommonExcel = window['AscCommonExcel'];
		var AscCH = window['AscCH'];
		var entries = null;
		CSS.forEachStyleOnlyCell(ws, bbox, function (row, col, xfIndex) {
			if (!entries) {
				entries = [];
			}
			entries.push(row, col, xfIndex);
		}, {excludeHiddenRows: !!opt_excludeHiddenRows});
		if (!entries) {
			return;
		}
		var historyOn = AscCommon.History.Is_On();
		var sheetId = historyOn ? ws.getId() : null;
		var styleCache = AscCommonExcel.g_StyleCache;
		var UndoRedoData_CellSimpleData = AscCommonExcel.UndoRedoData_CellSimpleData;
		for (var i = 0; i < entries.length; i += 3) {
			var row = entries[i];
			var col = entries[i + 1];
			var xfIndex = entries[i + 2];
			if (historyOn) {
				var oldXfs = styleCache.getXf(xfIndex);
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetStyleOnly,
					sheetId, new Asc.Range(col, row, col, row),
					new UndoRedoData_CellSimpleData(row, col, oldXfs, null));
			}
			ws.setCellXf(row, col, null);
		}
	}

	// Tile style-only sources over destination tiles. Runs after the data
	// pass; forEachStyleOnlyCell skips data init rows so the two passes
	// stay disjoint.
	function promoteStyleOnlyDirectStyles(wsFrom, from, wsTo, to, nDx, nDy) {
		if (!wsFrom || !wsTo || nDx <= 0 || nDy <= 0) {
			return;
		}
		if (typeof CSS.forEachStyleOnlyCell !== 'function') {
			return;
		}
		var AscCommon = window['AscCommon'];
		var AscCommonExcel = window['AscCommonExcel'];
		var AscCH = window['AscCH'];
		// Gather sources first so per-tile cost is O(|sources|) and we can
		// early-return before touching history.
		var sources = null;
		CSS.forEachStyleOnlyCell(wsFrom, from, function (row, col, xfIndex) {
			if (xfIndex > 0) {
				if (!sources) { sources = []; }
				sources.push(row - from.r1, col - from.c1, xfIndex);
			}
		}, {excludeHiddenRows: !!wsFrom.bExcludeHiddenRows});
		if (!sources) {
			return;
		}
		var historyOn = AscCommon.History.Is_On();
		var sheetId = historyOn ? wsTo.getId() : null;
		var styleCache = AscCommonExcel.g_StyleCache;
		var UndoRedoData_CellSimpleData = AscCommonExcel.UndoRedoData_CellSimpleData;
		var checkHiddenDst = !!wsTo.bExcludeHiddenRows;
		for (var ti = to.c1; ti <= to.c2; ti += nDx) {
			for (var tj = to.r1; tj <= to.r2; tj += nDy) {
				for (var k = 0; k < sources.length; k += 3) {
					var dr = sources[k];
					var dc = sources[k + 1];
					var xfIndex = sources[k + 2];
					var destRow = tj + dr;
					var destCol = ti + dc;
					if (destRow < to.r1 || destRow > to.r2 ||
						destCol < to.c1 || destCol > to.c2) {
						continue;
					}
					if (checkHiddenDst && wsTo.getRowHidden(destRow)) {
						continue;
					}
					var oldXf = wsTo.getCellXf(destRow, destCol);
					if (oldXf === xfIndex) {
						continue;
					}
					if (historyOn) {
						var oldXfs = oldXf > 0 ? styleCache.getXf(oldXf) : null;
						var newXfs = styleCache.getXf(xfIndex);
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell,
							AscCH.historyitem_Cell_SetStyleOnly,
							sheetId,
							new Asc.Range(destCol, destRow, destCol, destRow),
							new UndoRedoData_CellSimpleData(destRow, destCol, oldXfs, newXfs));
					}
					wsTo.setCellXf(destRow, destCol, xfIndex);
				}
			}
		}
	}

	// Style-only complement of clearDataKeepXf for post-insert borders.
	// History is suppressed; the enclosing structural item replays on undo.
	function applyInsertedBorderToStyleOnly(ws, bbox, borders, bRow) {
		if (typeof CSS.forEachStyleOnlyCell !== 'function') {
			return;
		}
		var AscCommonExcel = window['AscCommonExcel'];
		var styleCache = AscCommonExcel.g_StyleCache;
		var entries = null;
		CSS.forEachStyleOnlyCell(ws, bbox, function (row, col, xfIndex) {
			if (!entries) { entries = []; }
			entries.push(row, col, xfIndex);
		});
		if (!entries) {
			return;
		}
		for (var i = 0; i < entries.length; i += 3) {
			var row = entries[i];
			var col = entries[i + 1];
			var xfIndex = entries[i + 2];
			var oldXfs = styleCache.getXf(xfIndex);
			if (!oldXfs) {
				continue;
			}
			var key = bRow ? col : row;
			var newBorder = (borders && borders[key]) ? borders[key] : null;
			var newXfs = oldXfs.clone();
			newXfs.border = (newBorder != null) ? styleCache.addBorder(newBorder) : null;
			var registered = styleCache.addXf(newXfs);
			if (registered === oldXfs) {
				continue;
			}
			ws.setCellXf(row, col, registered);
		}
	}

	// Emit SetStyleOnly history for every style-only entry in `bbox`
	// before a structural shift drops it. Data cells in the same bbox
	// are covered by the enclosing structural item's history.
	function recordStyleOnlyClearHistory(ws, bbox, opt_excludeHiddenRows) {
		var AscCommon = window['AscCommon'];
		if (!AscCommon.History.Is_On() || !ws || !ws.cellStylesByCol || !bbox) {
			return;
		}
		if (typeof CSS.forEachStyleOnlyCell !== 'function') {
			return;
		}
		var AscCommonExcel = window['AscCommonExcel'];
		var AscCH = window['AscCH'];
		var styleCache = AscCommonExcel.g_StyleCache;
		var UndoRedoData_CellSimpleData = AscCommonExcel.UndoRedoData_CellSimpleData;
		var sheetId = ws.getId();
		CSS.forEachStyleOnlyCell(ws, bbox, function (row, col, xfIndex) {
			if (xfIndex <= 0) {
				return;
			}
			var oldXfs = styleCache.getXf(xfIndex);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetStyleOnly,
				sheetId, new Asc.Range(col, row, col, row),
				new UndoRedoData_CellSimpleData(row, col, oldXfs, null));
		}, {excludeHiddenRows: !!opt_excludeHiddenRows});
	}

	// Style-only complement of Range._setBorderEdge. Iterates style-only
	// cells in `edgeBbox`, lets the caller-provided Range apply the edge
	// override through its own `_setBorderEdge`, and emits SetStyleOnly
	// history for any cell whose xfs changed. The Range is borrowed only
	// for its `_setBorderEdge` prototype method.
	function applyBorderEdgeToStyleOnly(range, edgeBbox, bbox, oNewBorder) {
		if (typeof CSS.forEachStyleOnlyCell !== 'function') {
			return;
		}
		var ws = range.worksheet;
		var entries = null;
		CSS.forEachStyleOnlyCell(ws, edgeBbox, function (row, col, xfIndex) {
			if (!entries) { entries = []; }
			entries.push(row, col, xfIndex);
		});
		if (!entries) {
			return;
		}
		var AscCommon = window['AscCommon'];
		var AscCommonExcel = window['AscCommonExcel'];
		var AscCH = window['AscCH'];
		var styleCache = AscCommonExcel.g_StyleCache;
		var Cell = AscCommonExcel.Cell;
		var UndoRedoData_CellSimpleData = AscCommonExcel.UndoRedoData_CellSimpleData;
		var historyOn = AscCommon.History.Is_On();
		var sheetId = historyOn ? ws.getId() : null;
		var tempCell = new Cell(ws);
		tempCell._isTransient = true;
		var wb = ws.workbook;
		wb.loadCells.push(tempCell);
		try {
			for (var i = 0; i < entries.length; i += 3) {
				var row = entries[i];
				var col = entries[i + 1];
				var xfIndex = entries[i + 2];
				var oldXfs = styleCache.getXf(xfIndex);
				tempCell.clear();
				tempCell.setRowCol(row, col);
				tempCell.xfs = oldXfs;
				AscCommon.History.TurnOff();
				range._setBorderEdge(bbox, tempCell, row, col, oNewBorder);
				AscCommon.History.TurnOn();
				if (tempCell.xfs !== oldXfs) {
					var newXfs = tempCell.xfs;
					if (historyOn) {
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell,
							AscCH.historyitem_Cell_SetStyleOnly, sheetId,
							new Asc.Range(col, row, col, row),
							new UndoRedoData_CellSimpleData(row, col, oldXfs, newXfs));
					}
					ws.setCellXf(row, col, newXfs);
				}
			}
		} finally {
			wb.loadCells.pop();
		}
	}

	CSS.deleteRowsAllCols = deleteRowsAllCols;
	CSS.insertRowsAllCols = insertRowsAllCols;
	CSS.copyRowInAllCols = copyRowInAllCols;
	CSS.deleteCols = deleteCols;
	CSS.insertCols = insertCols;
	CSS.moveColRowBand = moveColRowBand;
	CSS.moveCellsBetweenWorksheets = moveCellsBetweenWorksheets;
	CSS.sortCellXfs = sortCellXfs;
	CSS.cleanStyleOnlyDirectStyles = cleanStyleOnlyDirectStyles;
	CSS.promoteStyleOnlyDirectStyles = promoteStyleOnlyDirectStyles;
	CSS.applyInsertedBorderToStyleOnly = applyInsertedBorderToStyleOnly;
	CSS.applyBorderEdgeToStyleOnly = applyBorderEdgeToStyleOnly;
	CSS.recordStyleOnlyClearHistory = recordStyleOnlyClearHistory;
})(window);
