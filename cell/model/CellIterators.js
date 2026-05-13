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

// Cell-shaped iterators composing SheetMemory data scans with
// cellStylesByCol cursors. Loaded after Workbook.js to reference the
// exported Cell and RowIterator via AscCommonExcel.
(function (window, undefined) {
	var ns = window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	// Largest row index with any direct cell style entry in [c1..c2],
	// or -1 when none.
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

	// Pull-based row iterator yielding data and direct-style-only cells.
	// Direct-style-only emissions return a shared transient Cell (read-only,
	// no saveContent). Within a row, columns emit in ascending order; a cell
	// with both data and a direct cellStylesByCol entry emits once via the
	// data path because Cell.loadContent resolves the direct xf.
	function OccupiedRowIterator() {
	}
	OccupiedRowIterator.prototype.init = function (ws, r1, c1, c2) {
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
			this.dataIter.init(ws, r1, c1, c2);
		}
		this.transientCell = new ns.Cell(ws);
		this.transientCell._isTransient = true;
		this.row = r1 - 1;
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
		var stylesByCol = this.ws.cellStylesByCol;
		if (!stylesByCol || stylesByCol.length === 0) {
			return null;
		}
		var maxC = stylesByCol.length - 1;
		if (maxC > this.c2) {
			maxC = this.c2;
		}
		var cols = null;
		for (var c = this.c1; c <= maxC; c++) {
			var store = stylesByCol[c];
			if (!store) {
				continue;
			}
			var xf = store.get(row);
			if (xf == null || xf === 0) {
				continue;
			}
			if (cols === null) {
				cols = [];
			}
			cols.push(c);
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
				// data + direct xf at same col: emit once via data path.
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
})(window);
