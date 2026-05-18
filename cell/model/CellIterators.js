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
				itRow.init(ws, oBBox.r1, oBBox.c1, oBBox.c2);
			}
			var bExcludeHiddenRows = (ws.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell;
			var tempRow = new Row(ws);
			var allRow = ws.getAllRow();
			var allRowHidden = allRow && allRow.getHidden();
			for (i = oBBox.r1; i <= minR; i++) {
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
				itRow.init(this.worksheet, this.bbox.r1, this.bbox.c1, this.bbox.c2);
			}
			var bExcludeHiddenRows = (this.worksheet.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell;
			var tempRow = new Row(this.worksheet);
			var allRow = this.worksheet.getAllRow();
			var allRowHidden = allRow && allRow.getHidden();
			for (i = oBBox.r1; i <= minR; i++) {
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
					for (i = oBBox.r1; i <= Math.min(minR, colData.getMaxIndex()); i++) {
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
