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

/*
 * Adaptive RowIterator mode. Keeps emitted cells ordered by row, then column.
 * Loaded after Workbook.js and before CellIterators.js.
 */
(function (window, undefined) {
	"use strict";

	var ns = window["AscCommonExcel"] = window["AscCommonExcel"] || {};
	var RowIterator = ns.RowIterator;
	if (!RowIterator || !RowIterator.prototype) {
		return;
	}
	if (RowIterator.prototype._adaptivePatched) {
		return;
	}
	RowIterator.prototype._adaptivePatched = true;

	var SweepLineRowIterator = ns.SweepLineRowIterator;
	var Cell = ns.Cell;

	var MODE_BASELINE = "baseline";
	var MODE_ADAPTIVE = "adaptive";

	function normalizeMode(value) {
		if (value === "p1" || value === MODE_ADAPTIVE) return MODE_ADAPTIVE;
		return MODE_BASELINE;
	}

	if (ns.g_rowIteratorMode == null && ns.g_rowIteratorProto != null) {
		ns.g_rowIteratorMode = normalizeMode(ns.g_rowIteratorProto);
	}
	if (ns.g_rowIteratorMode == null) {
		ns.g_rowIteratorMode = MODE_BASELINE;
	}
	ns.g_rowIteratorProto = ns.g_rowIteratorMode;

	ns.setRowIteratorMode = function (mode) {
		var m = normalizeMode(mode);
		ns.g_rowIteratorMode = m;
		ns.g_rowIteratorProto = m;
	};
	ns.getRowIteratorMode = function () {
		return ns.g_rowIteratorMode || MODE_BASELINE;
	};

	// Compatibility alias for older test runners.
	ns.setRowIteratorPrototype = function (name) {
		ns.setRowIteratorMode(name);
	};

	function initAdaptiveColumns(it, ws, r1, c1, c2, r2) {
		var src = ws.cellsByCol;
		var cap = Math.min(c2, src.length - 1);
		var aCols = [], aDatas = [];
		var effMin = Number.MAX_SAFE_INTEGER, effMax = -1;
		var maxOfMax = -1, minOfMax = Number.MAX_SAFE_INTEGER;
		for (var ci = c1; ci <= cap; ci++) {
			var cd = src[ci];
			if (cd && r1 <= cd.getMaxIndex() && cd.getMinIndex() <= r2) {
				aCols.push(ci);
				aDatas.push(cd);
				var mn = cd.getMinIndex();
				var mx = cd.getMaxIndex();
				if (mn < effMin) effMin = mn;
				if (mx > effMax) effMax = mx;
				if (mx > maxOfMax) maxOfMax = mx;
				if (mx < minOfMax) minOfMax = mx;
			}
		}
		var aLen = aCols.length;
		it._activeCols = aCols;
		it._activeDatas = aDatas;
		it._activeLen = aLen;
		it._effMinR = effMax < 0 ? r2 + 1 : (r1 > effMin ? r1 : effMin);
		it._effMaxR = effMax < 0 ? -1 : (r2 < effMax ? r2 : effMax);
		return { aLen: aLen, maxOfMax: maxOfMax, minOfMax: minOfMax };
	}

	// Tapered-shape trigger: build the alive list + maxIndex-sorted prune
	// order only when the per-row prune is worth the extra link traversal.
	function initTaperedPrune(it, stats, bboxSpan) {
		if (stats.aLen < 8 || (stats.maxOfMax - stats.minOfMax) <= (bboxSpan >> 2)) {
			return;
		}
		var aLen = stats.aLen;
		var aDatas = it._activeDatas;
		var aliveNext = new Int32Array(aLen);
		var alivePrev = new Int32Array(aLen);
		for (var k = 0; k < aLen; k++) {
			alivePrev[k] = k - 1;
			aliveNext[k] = (k + 1 < aLen) ? (k + 1) : -1;
		}
		// Sort positions by getMaxIndex ascending via a plain Array (typed
		// arrays don't accept closure-capturing comparators), then copy into
		// Int32Arrays so the hot per-row prune stays monomorphic.
		var permArr = new Array(aLen);
		for (var p = 0; p < aLen; p++) permArr[p] = p;
		permArr.sort(function (a, b) { return aDatas[a].getMaxIndex() - aDatas[b].getMaxIndex(); });
		var pruneOrder = new Int32Array(aLen);
		var pruneMax = new Int32Array(aLen);
		for (var q = 0; q < aLen; q++) {
			pruneOrder[q] = permArr[q];
			pruneMax[q] = aDatas[permArr[q]].getMaxIndex();
		}
		it._tapered = true;
		it._aliveNext = aliveNext;
		it._alivePrev = alivePrev;
		it._aliveFirst = aLen > 0 ? 0 : -1;
		it._pruneOrder = pruneOrder;
		it._pruneMax = pruneMax;
		it._pruneCursor = 0;
	}

	// Advance the alive-list head past every column whose getMaxIndex is
	// now below `row`. Correct only under monotone non-decreasing setRow
	// (which is what Range._foreach{DataOnly,NoEmpty} drives).
	function pruneInactiveColumns(it, row) {
		var pruneMax = it._pruneMax;
		var pruneOrder = it._pruneOrder;
		var aliveNext = it._aliveNext;
		var alivePrev = it._alivePrev;
		var aLen = it._activeLen;
		var cursor = it._pruneCursor;
		var first = it._aliveFirst;
		while (cursor < aLen && pruneMax[cursor] < row) {
			var pos = pruneOrder[cursor];
			var prev = alivePrev[pos];
			var next = aliveNext[pos];
			if (prev >= 0) aliveNext[prev] = next;
			else first = next;
			if (next >= 0) alivePrev[next] = prev;
			cursor++;
		}
		it._pruneCursor = cursor;
		it._aliveFirst = first;
		return first;
	}

	// Save the baseline prototype methods; the patched dispatchers below
	// delegate to them in baseline mode instead of re-implementing the
	// SweepLine wrapper.
	var baselineInit = RowIterator.prototype.init;

	function nullAdaptiveFields(it) {
		it._activeCols = null;
		it._activeDatas = null;
		it._activeLen = 0;
		it._activeIdx = 0;
		it._effMinR = 0;
		it._effMaxR = -1;
		it.curRow = 0;
		it._tapered = false;
		it._aliveNext = null;
		it._alivePrev = null;
		it._aliveFirst = -1;
		it._pruneOrder = null;
		it._pruneMax = null;
		it._pruneCursor = 0;
	}

	RowIterator.prototype.init = function (ws, r1, c1, c2, r2) {
		this._mode = ns.getRowIteratorMode();
		this._proto = this._mode === MODE_ADAPTIVE ? "p1" : "baseline";
		if (this._mode !== MODE_ADAPTIVE) {
			baselineInit.call(this, ws, r1, c1, c2);
			// Initialize the adaptive fields too so every RowIterator instance
			// shares one hidden class regardless of mode.
			nullAdaptiveFields(this);
			return;
		}
		if (r2 == null) r2 = window["AscCommon"].gc_nMaxRow0;
		this.ws = ws;
		this.cell = new Cell(ws);
		this.ws.workbook.loadCells.push(this.cell);
		this.iter = null;
		nullAdaptiveFields(this);
		this.curRow = r1 - 1;
		var stats = initAdaptiveColumns(this, ws, r1, c1, c2, r2);
		initTaperedPrune(this, stats, r2 - r1 + 1);
	};

	// Adaptive mode publishes its active-extent so Range._foreach* can clip
	// the outer row loop. Baseline returns -1/-1: SweepLine has no extent
	// to publish, so the caller keeps the legacy full-bbox walk.
	RowIterator.prototype.getEffMinR = function () {
		return this._mode === MODE_ADAPTIVE ? this._effMinR : -1;
	};
	RowIterator.prototype.getEffMaxR = function () {
		return this._mode === MODE_ADAPTIVE ? this._effMaxR : -1;
	};

	RowIterator.prototype.release = function () {
		if (this._mode === MODE_ADAPTIVE) {
			// Skip the history pass for read-only iteration. Every public
			// mutator sets Cell._hasChanged.
			if (this.cell._hasChanged) {
				this.cell.saveContent();
			}
		} else {
			this.cell.saveContent(true);
		}
		this.ws.workbook.loadCells.pop();
	};

	RowIterator.prototype.setRow = function (index) {
		if (this._mode === MODE_ADAPTIVE) {
			this.curRow = index;
			if (index < this._effMinR || index > this._effMaxR) {
				this._activeIdx = this._tapered ? -1 : this._activeLen;
				return;
			}
			if (this._tapered) {
				this._activeIdx = pruneInactiveColumns(this, index);
			} else {
				this._activeIdx = 0;
			}
		} else {
			this.iter.setRow(index);
		}
	};

	// Walks the alive linked list (tapered) or the plain column-ascending
	// array. Either way emission stays (row, col) monotone non-decreasing.
	RowIterator.prototype.nextAdaptive = function () {
		// Inline the _hasChanged guard: saveContent(true) still pays the
		// call cost when there is nothing to save.
		if (this.cell._hasChanged) {
			this.cell.saveContent();
		}
		var ws = this.ws;
		var wb = ws.workbook;
		var row = this.curRow;
		var aCols = this._activeCols;
		var aDatas = this._activeDatas;
		var loadCells = wb.loadCells;
		// Skip the loadCells identity walk at depth 1 (no outer cell to clash).
		var depthOk = loadCells.length === 1;
		if (this._tapered) {
			var aliveNext = this._aliveNext;
			var pos = this._activeIdx;
			while (pos >= 0) {
				var cdT = aDatas[pos];
				if (cdT.isNonEmpty(row)) {
					var cT = aCols[pos];
					this._activeIdx = aliveNext[pos];
					if (!depthOk) {
						for (var kT = 0; kT < loadCells.length - 1; ++kT) {
							var elemT = loadCells[kT];
							if (elemT.nRow === row && elemT.nCol === cT && ws === elemT.ws) {
								return elemT;
							}
						}
					}
					if (this.cell.loadContent(row, cT, cdT)) {
						return this.cell;
					}
				}
				pos = aliveNext[pos];
			}
			this._activeIdx = -1;
			return undefined;
		}
		var aLen = this._activeLen;
		for (var i = this._activeIdx; i < aLen; i++) {
			var colData = aDatas[i];
			if (!colData.isNonEmpty(row)) continue;
			var c = aCols[i];
			this._activeIdx = i + 1;
			if (!depthOk) {
				for (var k = 0; k < loadCells.length - 1; ++k) {
					var elem = loadCells[k];
					if (elem.nRow === row && elem.nCol === c && ws === elem.ws) {
						return elem;
					}
				}
			}
			if (this.cell.loadContent(row, c, colData)) {
				return this.cell;
			}
		}
		this._activeIdx = aLen;
		return undefined;
	};

	var baselineNext = RowIterator.prototype.next;
	RowIterator.prototype.next = function () {
		if (this._mode === MODE_ADAPTIVE) {
			return this.nextAdaptive();
		}
		return baselineNext.call(this);
	};
})(window);
