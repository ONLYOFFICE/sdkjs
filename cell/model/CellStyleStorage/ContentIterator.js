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

// Content-aware iteration: visits the union of SheetMemory data cells
// and cellStylesByCol style runs over `bbox` without materializing
// persistent Cell objects. See docs/cell-style-storage-iterator-design.md
// for the full contract.
(function (window, undefined) {
	var CSS = window['AscCommonExcel'].CellStyleStorage;
	var _internals = CSS._internals;
	var _isDataInitRow = _internals.isDataInitRow;

	// Per-column cursor that merges a SheetMemory data scan and a
	// CRangeAttrArray run cursor lazily: each consume() advances only as
	// far as the next emittable event.
	function _makeContentColumnCursor(ws, col, c1, r1, c2, r2, opts) {
		var sheetMemory = (typeof ws.getColDataNoEmpty === 'function')
			? ws.getColDataNoEmpty(col)
			: (ws.cellsByCol ? ws.cellsByCol[col] : null);
		var styleStore = ws.cellStylesByCol[col] || null;

		var dataRow = Infinity;
		var dataEnd = -1;
		if (sheetMemory && sheetMemory.dataBuffer && sheetMemory.indexA >= 0
			&& sheetMemory.indexB >= sheetMemory.indexA) {
			var dStart = sheetMemory.indexA > r1 ? sheetMemory.indexA : r1;
			var dStop = sheetMemory.indexB < r2 ? sheetMemory.indexB : r2;
			if (dStart <= dStop) {
				dataRow = dStart - 1;
				dataEnd = dStop;
			}
		}

		function _advanceData() {
			if (dataRow === Infinity) {
				return;
			}
			var r = dataRow + 1;
			while (r <= dataEnd) {
				if (_isDataInitRow(sheetMemory, r)) {
					dataRow = r;
					return;
				}
				r++;
			}
			dataRow = Infinity;
		}

		var ranges = styleStore ? styleStore._ranges : null;
		var runIdx = 0;
		var runLo = Infinity;
		var runHi = -1;
		var runVal = 0;

		function _loadRun() {
			while (ranges && runIdx < ranges.length) {
				var rg = ranges[runIdx];
				if (rg.r1 > r2) {
					break;
				}
				if (rg.r2 < r1) {
					runIdx++;
					continue;
				}
				runLo = rg.r1 > r1 ? rg.r1 : r1;
				runHi = rg.r2 < r2 ? rg.r2 : r2;
				runVal = rg.v;
				return;
			}
			runLo = Infinity;
			runHi = -1;
		}

		function _advanceRun() {
			runIdx++;
			_loadRun();
		}

		if (ranges && ranges.length) {
			runIdx = styleStore._lowerBound(r1);
			_loadRun();
		}
		_advanceData();

		var pending = null;
		var dataOnly = !!opts.dataOnly;
		var styleOnly = !!opts.styleOnly;

		function _makeData(row, hasStyle, xfIndex) {
			return {
				type: 'data',
				row: row,
				col: col,
				hasData: true,
				hasStyle: hasStyle,
				xfIndex: xfIndex,
				colData: sheetMemory,
				colStyleStore: styleStore
			};
		}

		function _makeStyleRun(lo, hi, xfIndex) {
			return {
				type: 'styleRun',
				row: lo,
				col: col,
				hi: hi,
				hasData: false,
				hasStyle: true,
				xfIndex: xfIndex,
				colData: sheetMemory,
				colStyleStore: styleStore
			};
		}

		function _pullRaw() {
			while (true) {
				if (dataRow === Infinity && runLo === Infinity) {
					return null;
				}
				if (runLo === Infinity || dataRow < runLo) {
					var dOnly = dataRow;
					_advanceData();
					return _makeData(dOnly, false, 0);
				}
				if (dataRow === Infinity || dataRow > runHi) {
					var sLoFull = runLo;
					var sHiFull = runHi;
					var sValFull = runVal;
					_advanceRun();
					return _makeStyleRun(sLoFull, sHiFull, sValFull);
				}
				// Overlap: runLo <= dataRow <= runHi.
				if (runLo < dataRow) {
					var sLo = runLo;
					var sHi = dataRow - 1;
					var sVal = runVal;
					runLo = dataRow;
					return _makeStyleRun(sLo, sHi, sVal);
				}
				// runLo === dataRow: emit data+style and consume one row of the run.
				var dr = dataRow;
				var xf = runVal;
				_advanceData();
				if (runLo === runHi) {
					_advanceRun();
				} else {
					runLo++;
				}
				return _makeData(dr, true, xf);
			}
		}

		function _refill() {
			while (true) {
				var ev = _pullRaw();
				if (!ev) {
					pending = null;
					return;
				}
				if (dataOnly && ev.type === 'styleRun') {
					continue;
				}
				if (styleOnly && ev.type === 'data') {
					continue;
				}
				pending = ev;
				return;
			}
		}
		_refill();

		return {
			col: col,
			peek: function () {
				return pending;
			},
			consume: function () {
				var e = pending;
				_refill();
				return e;
			}
		};
	}

	function _columnHasAnyContent(ws, col) {
		if (ws.cellStylesByCol[col]) {
			return true;
		}
		if (typeof ws.getColDataNoEmpty === 'function') {
			return !!ws.getColDataNoEmpty(col);
		}
		return ws.cellsByCol ? !!ws.cellsByCol[col] : false;
	}

	function iterContent(ws, bbox, callback, options) {
		if (!ws || !bbox || typeof callback !== 'function') {
			return;
		}
		var opts = options || {};
		var columnMajor = (opts.columnMajor !== false);
		var c1 = bbox.c1 | 0;
		var c2 = bbox.c2 | 0;
		var r1 = bbox.r1 | 0;
		var r2 = bbox.r2 | 0;
		if (c1 < 0) {
			c1 = 0;
		}
		if (r1 < 0) {
			r1 = 0;
		}
		if (c2 < c1 || r2 < r1) {
			return;
		}

		if (columnMajor) {
			for (var c = c1; c <= c2; c++) {
				if (!_columnHasAnyContent(ws, c)) {
					continue;
				}
				var cursor = _makeContentColumnCursor(ws, c, c1, r1, c2, r2, opts);
				while (cursor.peek()) {
					var ev = cursor.consume();
					if (callback(ev) === false) {
						return;
					}
				}
			}
			return;
		}

		// Row-major: streaming k-way merge over active per-column cursors.
		// Sort key is (event.row, col); for styleRun events event.row is the
		// run's current lo (clipped to bbox). See iterator-design section 3
		// for the stability rule.
		var cursors = [];
		for (var rc = c1; rc <= c2; rc++) {
			if (!_columnHasAnyContent(ws, rc)) {
				continue;
			}
			var cur = _makeContentColumnCursor(ws, rc, c1, r1, c2, r2, opts);
			if (cur.peek()) {
				cursors.push(cur);
			}
		}
		while (cursors.length) {
			var bestIdx = -1;
			var bestRow = Infinity;
			for (var i = 0; i < cursors.length; i++) {
				var p = cursors[i].peek();
				if (p && p.row < bestRow) {
					bestRow = p.row;
					bestIdx = i;
				}
			}
			if (bestIdx === -1) {
				return;
			}
			var emitted = cursors[bestIdx].consume();
			if (!cursors[bestIdx].peek()) {
				cursors.splice(bestIdx, 1);
			}
			if (callback(emitted) === false) {
				return;
			}
		}
	}

	// Per-cell wrapper over _makeContentColumnCursor. iterContent emits
	// styleRun events as multi-row spans; XLSB/JSON writers need per-row
	// events because XLSB requires each row to appear in ascending order
	// exactly once. The wrapped styleRun stays pending in the inner
	// cursor until every row in [lo..hi] has been emitted.
	function _makeByCellCursor(ws, col, c1, r1, c2, r2, opts) {
		var inner = _makeContentColumnCursor(ws, col, c1, r1, c2, r2, opts);
		var curRow = Infinity;

		function _sync() {
			var ev = inner.peek();
			curRow = ev ? ev.row : Infinity;
		}
		_sync();

		return {
			col: col,
			peekRow: function () { return curRow; },
			hasMore: function () { return curRow !== Infinity; },
			consume: function () {
				var ev = inner.peek();
				if (!ev) {
					return null;
				}
				var emitted;
				if (ev.type === 'data') {
					emitted = ev;
					inner.consume();
					_sync();
				} else {
					emitted = {
						type: 'styleRun',
						row: curRow,
						col: ev.col,
						hi: curRow,
						hasData: false,
						hasStyle: true,
						xfIndex: ev.xfIndex,
						colData: ev.colData,
						colStyleStore: ev.colStyleStore
					};
					if (curRow < ev.hi) {
						curRow++;
					} else {
						inner.consume();
						_sync();
					}
				}
				return emitted;
			}
		};
	}

	function _normalizeBBox(bbox) {
		var c1 = bbox.c1 | 0;
		var c2 = bbox.c2 | 0;
		var r1 = bbox.r1 | 0;
		var r2 = bbox.r2 | 0;
		if (c1 < 0) { c1 = 0; }
		if (r1 < 0) { r1 = 0; }
		return (c2 < c1 || r2 < r1) ? null : { c1: c1, r1: r1, c2: c2, r2: r2 };
	}

	// Pollable per-cell row-major cursor over (data + style-only) events.
	// Writers use this to interleave style-only emissions with the existing
	// data-cell stream by draining cursor events that come before each data
	// cell or row-metadata event.
	function createContentByCellCursor(ws, bbox, options) {
		var b = _normalizeBBox(bbox);
		if (!ws || !b) {
			return { peek: function () { return null; }, consume: function () { return null; } };
		}
		var opts = options || {};
		var styleOnly = !!opts.styleOnly;
		var cursors = [];
		for (var c = b.c1; c <= b.c2; c++) {
			// styleOnly skips columns with no usable style runs so a
			// full-sheet save doesn't pay for a per-column data scan only
			// to drop every event in the filter.
			if (styleOnly) {
				var styleStoreCol = ws.cellStylesByCol[c];
				if (!styleStoreCol || styleStoreCol.isEmpty()) {
					continue;
				}
				if (styleStoreCol.lastRow() < b.r1 || styleStoreCol.firstRow() > b.r2) {
					continue;
				}
			}
			if (!_columnHasAnyContent(ws, c)) {
				continue;
			}
			var cur = _makeByCellCursor(ws, c, b.c1, b.r1, b.c2, b.r2, opts);
			if (cur.hasMore()) {
				cursors.push(cur);
			}
		}

		function _pickBest() {
			var bestIdx = -1;
			var bestRow = Infinity;
			for (var i = 0; i < cursors.length; i++) {
				var pr = cursors[i].peekRow();
				if (pr < bestRow) {
					bestRow = pr;
					bestIdx = i;
				}
			}
			return bestIdx;
		}

		return {
			peek: function () {
				var idx = _pickBest();
				if (idx === -1) {
					return null;
				}
				return { row: cursors[idx].peekRow(), col: cursors[idx].col };
			},
			consume: function () {
				var idx = _pickBest();
				if (idx === -1) {
					return null;
				}
				var emitted = cursors[idx].consume();
				if (!cursors[idx].hasMore()) {
					cursors.splice(idx, 1);
				}
				return emitted;
			}
		};
	}

	// Push-based row-major per-cell iteration. styleRun events are split
	// per row (lo == hi for every emit). Callback returning false aborts.
	function iterContentByCell(ws, bbox, callback, options) {
		if (!ws || !bbox || typeof callback !== 'function') {
			return;
		}
		var cursor = createContentByCellCursor(ws, bbox, options);
		while (true) {
			var ev = cursor.consume();
			if (!ev) {
				return;
			}
			if (callback(ev) === false) {
				return;
			}
		}
	}

	// Visit every (row, col) inside `bbox` that has a cellStylesByCol entry
	// but no SheetMemory data init flag -- the cells that `_foreachNoEmpty`
	// would miss. Callback receives `(row, col, xfIndex)`; returning false
	// stops iteration. Pure read: callers can clear store entries and emit
	// history afterwards. Skips cells that have both data and direct
	// style; those are covered by the data-only path, and a second pass
	// would emit duplicate history. `opts.excludeHiddenRows` mirrors
	// `_foreachNoEmpty`'s filter-view behavior.
	function forEachStyleOnlyCell(ws, bbox, callback, opts) {
		if (!ws || !bbox || typeof callback !== 'function') {
			return;
		}
		var stores = ws.cellStylesByCol;
		if (!stores) {
			return;
		}
		var c1 = bbox.c1 | 0;
		var c2 = bbox.c2 | 0;
		var r1 = bbox.r1 | 0;
		var r2 = bbox.r2 | 0;
		if (c1 < 0) { c1 = 0; }
		if (r1 < 0) { r1 = 0; }
		if (c2 < c1 || r2 < r1) {
			return;
		}
		var maxCol = stores.length - 1;
		if (c2 < maxCol) {
			maxCol = c2;
		}
		var excludeHiddenRows = !!(opts && opts.excludeHiddenRows);
		var canCheckHidden = excludeHiddenRows && typeof ws.getRowHidden === 'function';

		for (var c = c1; c <= maxCol; c++) {
			var store = stores[c];
			if (!store) {
				continue;
			}
			var sheetMemory = (typeof ws.getColDataNoEmpty === 'function')
				? ws.getColDataNoEmpty(c)
				: null;
			var stopped = false;
			var colCaptured = c;
			store.iter(r1, r2, function (lo, hi, xfIndex) {
				for (var r = lo; r <= hi; r++) {
					if (canCheckHidden && ws.getRowHidden(r)) {
						continue;
					}
					if (sheetMemory && sheetMemory.hasIndex(r) && _isDataInitRow(sheetMemory, r)) {
						continue;
					}
					if (callback(r, colCaptured, xfIndex) === false) {
						stopped = true;
						return false;
					}
				}
			});
			if (stopped) {
				return;
			}
		}
	}

	CSS.iterContent = iterContent;
	CSS.iterContentByCell = iterContentByCell;
	CSS.createContentByCellCursor = createContentByCellCursor;
	CSS.forEachStyleOnlyCell = forEachStyleOnlyCell;
})(window);
