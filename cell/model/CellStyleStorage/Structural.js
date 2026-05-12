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
	var _maxLegacyColLength = _internals.maxLegacyColLength;
	var getCellStyleStore = _internals.getCellStyleStore;

	// Shift cell styles inside / through `bbox` in one of four directions.
	// `offset` is currently informational; `mode` drives the behavior.
	// Modes:
	//   'up'    -- delete the bbox rows in cols [c1..c2]; rows below shift up.
	//   'down'  -- insert empty rows of bbox height in cols [c1..c2] at r1.
	//   'left'  -- within rows [r1..r2], pull style from columns to the right of c2 into [c1..]; clear the gap at the far right.
	//   'right' -- within rows [r1..r2], shift style in cols starting at c1 by +width; clear the inserted block at [c1..c1+width-1].
	function shiftCellXfs(ws, bbox, offset, mode) {
		var c1 = bbox.c1;
		var c2 = bbox.c2;
		var r1 = bbox.r1;
		var r2 = bbox.r2;
		if (c2 < c1 || r2 < r1) {
			return;
		}
		var w = c2 - c1 + 1;
		var h = r2 - r1 + 1;

		if (mode === 'up') {
			for (var c = c1; c <= c2; c++) {
				var sUp = ws.cellStylesByCol[c];
				if (sUp) {
					sUp.deleteRows(r1, h);
				}
			}
			return;
		}
		if (mode === 'down') {
			for (var cd = c1; cd <= c2; cd++) {
				var sDown = ws.cellStylesByCol[cd];
				if (sDown) {
					sDown.insertRows(r1, h);
				}
			}
			return;
		}
		if (mode === 'left') {
			// Bound must reach data-only columns to match `_shiftCellsLeft`.
			var L = Math.max(ws.cellStylesByCol.length, _maxLegacyColLength(ws));
			for (var cl = c1; cl < L; cl++) {
				var srcL = _resolveSourceStore(ws, cl + w);
				var dstL = ws.cellStylesByCol[cl];
				if (!srcL) {
					if (dstL) {
						dstL.clearRange(r1, r2);
					}
					continue;
				}
				if (!dstL) {
					dstL = getCellStyleStore(ws, cl, true);
				}
				dstL.copyFrom(srcL, r1, r1, h);
			}
			return;
		}
		if (mode === 'right') {
			// Rightmost source Lr-1 lands at Lr-1+w; destinations past the
			// current array length are materialized lazily. Same data-only
			// widening as the 'left' branch.
			var Lr = Math.max(ws.cellStylesByCol.length, _maxLegacyColLength(ws));
			for (var cr = Lr - 1 + w; cr >= c1 + w; cr--) {
				var srcR = _resolveSourceStore(ws, cr - w);
				var dstR = ws.cellStylesByCol[cr];
				if (!srcR) {
					if (dstR) {
						dstR.clearRange(r1, r2);
					}
					continue;
				}
				if (!dstR) {
					dstR = getCellStyleStore(ws, cr, true);
				}
				dstR.copyFrom(srcR, r1, r1, h);
			}
			for (var cg = c1; cg < c1 + w; cg++) {
				var dstGap = ws.cellStylesByCol[cg];
				if (dstGap) {
					dstGap.clearRange(r1, r2);
				}
			}
			return;
		}
	}

	// Iterate every style run intersecting bbox in column-major order.
	// callback(loRow, hiRow, col, xfIndex). Return false from the callback to stop.
	function iterCellXfs(ws, bbox, callback) {
		var c1 = bbox.c1;
		var c2 = bbox.c2;
		if (c2 < c1) {
			return;
		}
		for (var c = c1; c <= c2; c++) {
			var store = ws.cellStylesByCol[c];
			if (!store) {
				continue;
			}
			var stopped = false;
			var colCaptured = c;
			store.iter(bbox.r1, bbox.r2, function (lo, hi, v) {
				if (callback(lo, hi, colCaptured, v) === false) {
					stopped = true;
					return false;
				}
			});
			if (stopped) {
				return;
			}
		}
	}

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

	// Mirror the permutation that `Range._sortByArray` applies to
	// SheetMemory. Must run BEFORE the SheetMemory mutation so the
	// snapshot logic resolves cycles against pre-sort state.
	//
	// Vertical (`opt_by_row` falsy): per column, reorder rows in
	// [oBBox.r1..oBBox.r2] using `oSortedIndexes` (from-row -> to-row).
	// Reads come from the pre-mutation store, so partial permutations
	// resolve regardless of iteration order.
	//
	// Horizontal (`opt_by_row` truthy): swap the row band between
	// columns named by `oSortedIndexes` (from-col -> to-col). Each
	// destination's pre-mutation band is snapshotted before the
	// destination is overwritten, so later iterations that read the
	// same column as a source see the snapshot, not the mutated state.
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

			// Snapshot the destination band BEFORE any mutation. Even when
			// `to` has no store yet, record an empty snapshot so a later
			// iteration that reads `to` as a source sees the pre-mutation
			// "no styles in band" state rather than the mutated column.
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

	// Tile direct styles from pure style-only source cells over a
	// destination range. Runs AFTER `_promoteFromTo`'s data pass; the
	// two passes target disjoint (row, col) sets because
	// `forEachStyleOnlyCell` skips data init rows.
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
		// Gather source offsets first so we don't re-walk `cellStylesByCol`
		// per tile and so we can early-return without history side effects.
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

	CSS.shiftCellXfs = shiftCellXfs;
	CSS.iterCellXfs = iterCellXfs;
	CSS.deleteRowsAllCols = deleteRowsAllCols;
	CSS.insertRowsAllCols = insertRowsAllCols;
	CSS.copyRowInAllCols = copyRowInAllCols;
	CSS.deleteCols = deleteCols;
	CSS.insertCols = insertCols;
	CSS.moveColRowBand = moveColRowBand;
	CSS.sortCellXfs = sortCellXfs;
	CSS.cleanStyleOnlyDirectStyles = cleanStyleOnlyDirectStyles;
	CSS.promoteStyleOnlyDirectStyles = promoteStyleOnlyDirectStyles;
})(window);
