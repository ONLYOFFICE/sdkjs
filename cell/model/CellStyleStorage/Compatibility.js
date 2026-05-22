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

// Temporary compatibility flag for direct-cell-style storage:
//   useSeparatedCellStyles = true  -> current behavior; direct xf lives in
//                                     ws.cellStylesByCol, SheetMemory low-24
//                                     bits stay zero.
//   useSeparatedCellStyles = false -> legacy behavior; direct xf lives in
//                                     SheetMemory low-24 bits, cellStylesByCol
//                                     stays empty, occupied iteration
//                                     collapses to data-only.
//
// The whole legacy path is concentrated in this file. To delete it later:
//   1. Drop this file and the configs/cell.json entry.
//   2. Drop the `clearMovedSourceData` call site in Worksheet._moveCells back
//      to plain `fromData.clear(clearStart, clearEnd)`.
//   3. Drop `CSS.clearMovedSourceData` from Core/Compatibility.
(function (window, undefined) {
	var ns = window['AscCommonExcel'];
	var CSS = ns.CellStyleStorage;
	var Cell = ns.Cell;
	var Range = ns.Range;
	var Worksheet = ns.Worksheet;
	var Row = ns.Row;
	var RowIterator = ns.RowIterator;
	var SheetMemory = ns.SheetMemory;

	// Snapshot the separated-mode implementations that the sibling
	// CellStyleStorage files, Workbook.js, and CellIterators.js have already
	// installed by the time this file runs. Keeping the original references
	// here means installSeparatedMode is fully reversible.
	var sep = {
		cellSaveContent: Cell.prototype.saveContent,
		cellLoadContent: Cell.prototype.loadContent,
		rangeForeachNoEmpty: Range.prototype._foreachNoEmpty,
		rangeForeachNoEmptyByCol: Range.prototype._foreachNoEmptyByCol,
		worksheetGetRowIterator: Worksheet.prototype.getRowIterator,
		mirrorCellStyle: CSS.mirrorCellStyle,
		hydrateAllColumnsFromSheetMemory: CSS.hydrateAllColumnsFromSheetMemory,
		forEachStyleOnlyCell: CSS.forEachStyleOnlyCell,
		createStyleOnlyDrain: CSS.createStyleOnlyDrain,
		createContentByCellCursor: CSS.createContentByCellCursor,
		getDirectCellXfs: CSS.getDirectCellXfs,
		resolveLoadXfIndex: CSS.resolveLoadXfIndex,
		isStyleOnlyCell: CSS.isStyleOnlyCell,
		getWriterCellXfIndex: CSS.getWriterCellXfIndex,
		getWriterCellXfs: CSS.getWriterCellXfs,
		moveCellsBetweenWorksheets: CSS.moveCellsBetweenWorksheets,
		saveContentSkipStyleOnly: CSS.saveContentSkipStyleOnly
	};

	// SheetMemory.prototype.clearExceptLocked is unique because separated mode
	// does not define it at all -- legacy mode installs it lazily. Snapshot
	// whether the baseline owned the property and the current implementation
	// (if any) so installSeparatedMode can return the prototype to the exact
	// shape it had before installLegacyMode ran.
	var sepHadClearExceptLocked = Object.prototype.hasOwnProperty.call(
		SheetMemory.prototype, 'clearExceptLocked');
	var sepClearExceptLocked = sepHadClearExceptLocked
		? SheetMemory.prototype.clearExceptLocked
		: null;

	// _moveCells calls this AFTER mirroring the band to cellStylesByCol.
	// Separated mode: xf is already off the SheetMemory shadow, so just zero
	// the data bytes. Legacy mode: xf lives in the shadow, so preserve the
	// locked-only transform when the sheet is protected.
	function sepClearMovedSourceData(fromData, clearStart, clearEnd, getLockedOnlyXfIndex) {
		fromData.clear(clearStart, clearEnd);
	}
	function legacyClearMovedSourceData(fromData, clearStart, clearEnd, getLockedOnlyXfIndex) {
		if (getLockedOnlyXfIndex && typeof fromData.clearExceptLocked === 'function') {
			fromData.clearExceptLocked(clearStart, clearEnd, getLockedOnlyXfIndex);
		} else {
			fromData.clear(clearStart, clearEnd);
		}
	}

	// ======== Legacy Cell save / load ========

	// Mirror of g_nCellFlag_init from Workbook.js. That constant lives in a
	// module-scoped IIFE and is not exported, so the legacy loadContent
	// re-states its value here. Keep in sync with Workbook.js if it ever
	// changes (the SheetMemory flag layout is stable).
	var g_nCellFlag_init_legacy = 1;

	function legacyCellSaveContent(opt_inCaseOfChange) {
		if (this.hasRowCol() && (!opt_inCaseOfChange || this._hasChanged)) {
			this._hasChanged = false;
			var wb = this.ws.workbook;
			var sheetMemory = this.ws.getColData(this.nCol);
			sheetMemory.checkIndex(this.nRow);
			var xfSave = this.xfs ? this.xfs.getIndexNumber() : 0;
			var numberSave = 0;
			var formulaSave = this.formulaParsed
				? wb.workbookFormulas.add(this.formulaParsed).getIndexNumber()
				: 0;
			var flagValue = 0;
			if (null != this.number) {
				flagValue = 1;
				var flagsN = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flagsN << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
				sheetMemory.setFloat64(this.nRow, 8, this.number);
			} else if (null != this.text || null != this.multiText) {
				flagValue = 2;
				var flagsT = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flagsT << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
				numberSave = this.getTextIndex();
				sheetMemory.setInt32(this.nRow, 8, numberSave);
			} else {
				var flagsF = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flagsF << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
			}
		}
	}

	function legacyCellLoadContent(row, col, opt_sheetMemory) {
		var res = false;
		this.clear();
		this.nRow = row;
		this.nCol = col;
		var sheetMemory = opt_sheetMemory;
		if (!sheetMemory) {
			sheetMemory = this.ws.getColDataNoEmpty(this.nCol);
			if (!sheetMemory) {
				return res;
			}
		}
		if (sheetMemory.hasIndex(this.nRow)) {
			var mix = sheetMemory.getInt32(this.nRow, 0);
			var flags = (mix >> 24) & 0xff;
			var xfIndex = mix & 0xffffff;
			if (0 !== (g_nCellFlag_init_legacy & flags)) {
				var wb = this.ws.workbook;
				var flagValue = this._fromFlags(flags);
				if (xfIndex > 0) {
					this.xfs = ns.g_StyleCache.getXf(xfIndex);
				}
				var formulaIndex = sheetMemory.getInt32(this.nRow, 4);
				if (formulaIndex > 0) {
					this.formulaParsed = wb.workbookFormulas.get(formulaIndex);
				}
				if (1 === flagValue) {
					this.number = sheetMemory.getFloat64(this.nRow, 8);
				} else if (2 === flagValue) {
					this.textIndex = sheetMemory.getInt32(this.nRow, 8);
					var text = wb.sharedStrings.get(this.textIndex);
					typeof text === 'string' ? this.text = text : this.multiText = text;
				}
				res = true;
			}
		}
		return res;
	}

	// ======== Legacy Range iterators ========

	// Baseline _foreachNoEmpty: plain RowIterator, no style-only emissions.
	// Identical in behavior to the separated-mode _foreachDataOnly.
	function legacyForeachNoEmpty(actionCell, actionRow, excludeHiddenRows) {
		var oRes, i, oBBox = this.bbox;
		var ws = this.worksheet;
		var minR = Math.max(ws.cellsByColRowsCount - 1, ws.rowsData.getMaxIndex());
		minR = Math.min(minR, oBBox.r2);
		if (actionCell || actionRow) {
			var itRow = new RowIterator();
			if (actionCell) {
				itRow.init(ws, oBBox.r1, oBBox.c1, oBBox.c2, oBBox.r2);
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
							if (actionCell) {
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
				if (actionCell) {
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
			if (actionCell) {
				itRow.release();
			}
		}
	}

	// Baseline _foreachNoEmptyByCol: data-only column-major scan.
	function legacyForeachNoEmptyByCol(actionCell, excludeHiddenRows) {
		var oRes, i, j, colData;
		var ws = this.worksheet;
		var wb = ws.workbook;
		var oBBox = this.bbox;
		var minR = Math.min(ws.cellsByColRowsCount - 1, oBBox.r2);
		var minC = Math.min(ws.getColDataLength() - 1, oBBox.c2);
		if (actionCell && oBBox.c1 <= minC && oBBox.r1 <= minR) {
			var bExcludeHiddenRows = (ws.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell = new Cell(ws);
			tempCell._isTransient = true;
			wb.loadCells.push(tempCell);
			for (j = oBBox.c1; j <= minC; ++j) {
				colData = ws.getColDataNoEmpty(j);
				if (colData) {
					for (i = oBBox.r1; i <= Math.min(minR, colData.getMaxIndex()); i++) {
						if (bExcludeHiddenRows && ws.getRowHidden(i)) {
							excludedCount++;
							continue;
						}
						var targetCell = null;
						for (var k = 0; k < wb.loadCells.length - 1; ++k) {
							var elem = wb.loadCells[k];
							if (elem.nRow == i && elem.nCol == j && ws === elem.ws) {
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
	}

	// ======== Legacy Worksheet iterator entry ========

	// Legacy mode has no style-only cells, so getRowIterator collapses to the
	// data-only RowIterator. Callers see plain Cell instances (no transients).
	function legacyWorksheetGetRowIterator(r1, c1, c2, callback) {
		var it = new RowIterator();
		it.init(this, r1, c1, c2);
		callback(it);
		it.release();
	}

	// ======== Legacy CellStyleStorage helpers ========

	// Read direct xf from SheetMemory low-24 bits without materializing a Cell.
	// Returns 0 when the row has no init flag: legacy Cell.saveContent always
	// stamps the init flag alongside an xf write, so an init-clear row carries
	// no committed direct xf even if stale low-24 bits remain from a deleted
	// cell.
	function legacyReadXfIndexFromSheetMemory(ws, row, col) {
		if (!ws || row < 0 || col < 0) {
			return 0;
		}
		var sm = (typeof ws.getColDataNoEmpty === 'function')
			? ws.getColDataNoEmpty(col)
			: null;
		if (!sm) {
			return 0;
		}
		if (typeof sm.hasIndex === 'function' && !sm.hasIndex(row)) {
			return 0;
		}
		var mix = sm.getInt32(row, 0);
		if (((mix >>> 24) & 1) === 0) {
			return 0;
		}
		return mix & 0xFFFFFF;
	}

	function legacyGetDirectCellXfs(ws, row, col) {
		var idx = legacyReadXfIndexFromSheetMemory(ws, row, col);
		if (idx <= 0) {
			return null;
		}
		var cache = ns.g_StyleCache;
		return cache ? (cache.getXf(idx) || null) : null;
	}

	function legacyResolveLoadXfIndex(ws, row, col) {
		return legacyReadXfIndexFromSheetMemory(ws, row, col);
	}

	function legacyIsStyleOnlyCell(/*ws, row, col*/) {
		// In legacy mode every direct xf rides on a SheetMemory data init row;
		// there is no style-only cell.
		return false;
	}

	function legacyGetWriterCellXfIndex(ws, row, col, cell) {
		if (cell && cell.xfs && typeof cell.xfs.getIndexNumber === 'function') {
			var idx = cell.xfs.getIndexNumber();
			if (idx > 0) {
				return idx;
			}
		}
		return legacyReadXfIndexFromSheetMemory(ws, row, col);
	}

	function legacyGetWriterCellXfs(ws, row, col, cell) {
		if (cell && cell.xfs) {
			return cell.xfs;
		}
		var idx = legacyReadXfIndexFromSheetMemory(ws, row, col);
		if (idx <= 0) {
			return null;
		}
		var cache = ns.g_StyleCache;
		return cache ? (cache.getXf(idx) || null) : null;
	}

	function legacyMirrorCellStyle(/*cell*/) {
		// Legacy mode: xf lives in SheetMemory, persisted by Cell.saveContent.
	}

	function legacyHydrateAllColumnsFromSheetMemory(/*ws*/) {
		// Legacy mode: low-24 xf bits stay in SheetMemory and are read directly
		// by Cell.loadContent. Nothing to migrate.
	}

	function legacyForEachStyleOnlyCell(/*ws, bbox, callback, opts*/) {
		// No style-only cells in legacy mode.
	}

	function legacyCreateStyleOnlyDrain(/*ws, bbox, opts, onStyleOnly*/) {
		return {
			drainBefore: function () {},
			drainTail: function () {}
		};
	}

	function legacyCreateContentByCellCursor(/*ws, bbox, options*/) {
		return {
			peek: function () { return null; },
			consume: function () { return null; }
		};
	}

	function legacyMoveCellsBetweenWorksheets(/*wsFrom, wsTo, fromCol, toCol, r1From, r1To, count, clearStart, clearEnd, getLockedOnlyXfIndex*/) {
		// Legacy mode: xf travels with the SheetMemory copyRange called by
		// Worksheet._moveCells, and the locked-only transform happens through
		// SheetMemory.clearExceptLocked via clearMovedSourceData. No mirroring
		// is needed.
	}

	function legacySaveContentSkipStyleOnly(cell, opt_inCaseOfChange) {
		// Legacy mode has no pure style-only cells; persist normally so the
		// xf bits land in SheetMemory.
		cell.saveContent(opt_inCaseOfChange);
	}

	// ======== Legacy SheetMemory.clearExceptLocked ========

	// Installed only when entering legacy mode. Matches baseline a82cb624
	// behavior: clears the row's bytes but restores `getLockedOnlyXfIndex(xf)`
	// into the low 24 xf bits when the source xf carries any locked attribute.
	function installLegacySheetMemoryClearExceptLocked() {
		if (typeof SheetMemory.prototype.clearExceptLocked === 'function') {
			return;
		}
		SheetMemory.prototype.clearExceptLocked = function (start, end, getLockedOnlyXfIndex) {
			start = Math.max(start, this.indexA);
			end = Math.min(end, this.indexB + 1);
			if (start >= end) {
				return;
			}
			var initBit = 1;
			for (var i = start; i < end; i++) {
				var mix = this.getInt32(i, 0);
				var xfIndex = mix & 0xFFFFFF;
				var startOffset = (i - this.indexA) * this.structSize;
				var endOffset = startOffset + this.structSize;
				this.dataUint8.fill(0, startOffset, endOffset);
				if (xfIndex > 0 && getLockedOnlyXfIndex) {
					var newXf = getLockedOnlyXfIndex(xfIndex);
					if (newXf != null) {
						this.setInt32(i, 0, newXf | (initBit << 24));
					}
				}
			}
		};
	}

	// ======== Installers ========
	//
	// installSeparatedMode / installLegacyMode are PROCESS-GLOBAL mode
	// installers: they swap prototype methods on Cell / Range / Worksheet /
	// SheetMemory and replace function references on CellStyleStorage. They
	// do NOT migrate existing workbook data between the two storage layouts
	// (SheetMemory low-24 xf bits vs. ws.cellStylesByCol). The selected mode
	// must therefore be set BEFORE any workbook is opened or created in this
	// process; toggling mid-session against an already-loaded workbook will
	// leave that workbook half-migrated and is not supported.

	function installSeparatedMode() {
		CSS.useSeparatedCellStyles = true;
		Cell.prototype.saveContent = sep.cellSaveContent;
		Cell.prototype.loadContent = sep.cellLoadContent;
		Range.prototype._foreachNoEmpty = sep.rangeForeachNoEmpty;
		Range.prototype._foreachNoEmptyByCol = sep.rangeForeachNoEmptyByCol;
		Worksheet.prototype.getRowIterator = sep.worksheetGetRowIterator;
		CSS.mirrorCellStyle = sep.mirrorCellStyle;
		CSS.hydrateAllColumnsFromSheetMemory = sep.hydrateAllColumnsFromSheetMemory;
		CSS.forEachStyleOnlyCell = sep.forEachStyleOnlyCell;
		CSS.createStyleOnlyDrain = sep.createStyleOnlyDrain;
		CSS.createContentByCellCursor = sep.createContentByCellCursor;
		CSS.getDirectCellXfs = sep.getDirectCellXfs;
		CSS.resolveLoadXfIndex = sep.resolveLoadXfIndex;
		CSS.isStyleOnlyCell = sep.isStyleOnlyCell;
		CSS.getWriterCellXfIndex = sep.getWriterCellXfIndex;
		CSS.getWriterCellXfs = sep.getWriterCellXfs;
		CSS.moveCellsBetweenWorksheets = sep.moveCellsBetweenWorksheets;
		CSS.saveContentSkipStyleOnly = sep.saveContentSkipStyleOnly;
		CSS.clearMovedSourceData = sepClearMovedSourceData;
		// Return SheetMemory.prototype to the shape it had before any
		// legacy install ran: restore the original implementation when one
		// existed, otherwise drop the property entirely.
		if (sepHadClearExceptLocked) {
			SheetMemory.prototype.clearExceptLocked = sepClearExceptLocked;
		} else {
			delete SheetMemory.prototype.clearExceptLocked;
		}
	}

	function installLegacyMode() {
		CSS.useSeparatedCellStyles = false;
		installLegacySheetMemoryClearExceptLocked();
		Cell.prototype.saveContent = legacyCellSaveContent;
		Cell.prototype.loadContent = legacyCellLoadContent;
		Range.prototype._foreachNoEmpty = legacyForeachNoEmpty;
		Range.prototype._foreachNoEmptyByCol = legacyForeachNoEmptyByCol;
		Worksheet.prototype.getRowIterator = legacyWorksheetGetRowIterator;
		CSS.mirrorCellStyle = legacyMirrorCellStyle;
		CSS.hydrateAllColumnsFromSheetMemory = legacyHydrateAllColumnsFromSheetMemory;
		CSS.forEachStyleOnlyCell = legacyForEachStyleOnlyCell;
		CSS.createStyleOnlyDrain = legacyCreateStyleOnlyDrain;
		CSS.createContentByCellCursor = legacyCreateContentByCellCursor;
		CSS.getDirectCellXfs = legacyGetDirectCellXfs;
		CSS.resolveLoadXfIndex = legacyResolveLoadXfIndex;
		CSS.isStyleOnlyCell = legacyIsStyleOnlyCell;
		CSS.getWriterCellXfIndex = legacyGetWriterCellXfIndex;
		CSS.getWriterCellXfs = legacyGetWriterCellXfs;
		CSS.moveCellsBetweenWorksheets = legacyMoveCellsBetweenWorksheets;
		CSS.saveContentSkipStyleOnly = legacySaveContentSkipStyleOnly;
		CSS.clearMovedSourceData = legacyClearMovedSourceData;
	}

	function installMode(separated) {
		if (separated) {
			installSeparatedMode();
		} else {
			installLegacyMode();
		}
	}

	CSS.installSeparatedMode = installSeparatedMode;
	CSS.installLegacyMode = installLegacyMode;
	CSS.installMode = installMode;

	// Ensure the _moveCells call site always finds a hook, even if no
	// installer has run yet (separated mode is the default).
	if (typeof CSS.clearMovedSourceData !== 'function') {
		CSS.clearMovedSourceData = sepClearMovedSourceData;
	}

	// Default flag value if nothing set it before this file loaded.
	if (CSS.useSeparatedCellStyles == null) {
		CSS.useSeparatedCellStyles = true;
	}
	// Honor a pre-load override (e.g. an integrator sets the flag to false
	// before scripts execute) so the right mode is active for the very first
	// workbook open.
	if (!CSS.useSeparatedCellStyles) {
		installLegacyMode();
	}
})(window);
