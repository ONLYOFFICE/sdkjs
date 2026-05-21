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

// Writer / loader / protection hooks against cellStylesByCol.
// cellStylesByCol is the sole source of truth for direct cell xf at
// write and read time; LegacyMigration runs first at file open.
(function (window, undefined) {
	var CSS = window['AscCommonExcel'].CellStyleStorage;
	var _internals = CSS._internals;
	var _isDataInitRow = _internals.isDataInitRow;

	// `cell` is kept for call-site stability and not consulted.
	function getWriterCellXfIndex(ws, row, col, cell) {
		if (!ws || row < 0 || col < 0) {
			return 0;
		}
		var store = ws.cellStylesByCol[col];
		if (!store) {
			return 0;
		}
		var idx = store.get(row);
		return (idx == null) ? 0 : idx;
	}

	// CellXfs for `stylesForWrite.add(...)`, or null when no direct style.
	function getWriterCellXfs(ws, row, col, cell) {
		var idx = getWriterCellXfIndex(ws, row, col, cell);
		if (idx <= 0) {
			return null;
		}
		var cache = window['AscCommonExcel'].g_StyleCache;
		return cache ? cache.getXf(idx) : null;
	}

	// Read-time xfIndex for Cell.loadContent. 0 when no entry exists.
	function resolveLoadXfIndex(ws, row, col) {
		if (!ws || !ws.cellStylesByCol || row < 0 || col < 0) {
			return 0;
		}
		var store = ws.cellStylesByCol[col];
		if (!store) {
			return 0;
		}
		var idx = store.get(row);
		return (idx == null) ? 0 : idx;
	}

	// Direct CellXfs at (row, col), or null. Never materializes a Cell;
	// data cells reach style through _foreachNoEmpty / cell.xfs.
	function getDirectCellXfs(ws, row, col) {
		if (!ws || !ws.cellStylesByCol || row < 0 || col < 0) {
			return null;
		}
		var store = ws.cellStylesByCol[col];
		if (!store) {
			return null;
		}
		var idx = store.get(row);
		if (idx == null || idx <= 0) {
			return null;
		}
		var cache = window['AscCommonExcel'].g_StyleCache;
		return cache ? (cache.getXf(idx) || null) : null;
	}

	// True when (row, col) has no SheetMemory data init flag.
	function isStyleOnlyCell(ws, row, col) {
		if (!ws || row < 0 || col < 0) {
			return false;
		}
		var sm = (typeof ws.getColDataNoEmpty === 'function')
			? ws.getColDataNoEmpty(col)
			: null;
		if (!sm) {
			return true;
		}
		if (typeof sm.hasIndex === 'function' && !sm.hasIndex(row)) {
			return true;
		}
		return !_isDataInitRow(sm, row);
	}

	// saveContent variant that skips the SheetMemory write for pure
	// style-only cells (no number / text / multiText / formula); their xf
	// lives only in ws.cellStylesByCol. Used by open paths so a
	// roundtripped pure-style cell matches in-session shape (no SheetMemory
	// init row). Free function: called as
	// `AscCommonExcel.CellStyleStorage.saveContentSkipStyleOnly(cell, ...)`
	// from fromToJSON.js and Serialize.js after Workbook.js has defined
	// the Cell prototype.
	function saveContentSkipStyleOnly(cell, opt_inCaseOfChange) {
		if (!cell.hasRowCol()) {
			return;
		}
		if (opt_inCaseOfChange && !cell._hasChanged) {
			return;
		}
		if (null == cell.number && null == cell.text && null == cell.multiText
			&& null == cell.formulaParsed) {
			cell._hasChanged = false;
			return;
		}
		cell._hasChanged = false;
		var wb = cell.ws.workbook;
		var sheetMemory = cell.ws.getColData(cell.nCol);
		sheetMemory.checkIndex(cell.nRow);
		var numberSave = 0;
		var formulaSave = cell.formulaParsed ? wb.workbookFormulas.add(cell.formulaParsed).getIndexNumber() : 0;
		var flagValue = 0;
		if (null != cell.number) {
			flagValue = 1;
			var flagsN = cell._toFlags(flagValue);
			sheetMemory.setInt32(cell.nRow, 0, (flagsN << 24));
			sheetMemory.setInt32(cell.nRow, 4, formulaSave);
			sheetMemory.setFloat64(cell.nRow, 8, cell.number);
		} else if (null != cell.text || null != cell.multiText) {
			flagValue = 2;
			var flagsT = cell._toFlags(flagValue);
			sheetMemory.setInt32(cell.nRow, 0, (flagsT << 24));
			sheetMemory.setInt32(cell.nRow, 4, formulaSave);
			numberSave = cell.getTextIndex();
			sheetMemory.setInt32(cell.nRow, 8, numberSave);
		} else {
			// Formula-only: mirror saveContent's empty-branch write.
			var flagsF = cell._toFlags(flagValue);
			sheetMemory.setInt32(cell.nRow, 0, (flagsF << 24));
			sheetMemory.setInt32(cell.nRow, 4, formulaSave);
		}
	}

	CSS.getWriterCellXfIndex = getWriterCellXfIndex;
	CSS.getWriterCellXfs = getWriterCellXfs;
	CSS.resolveLoadXfIndex = resolveLoadXfIndex;
	CSS.getDirectCellXfs = getDirectCellXfs;
	CSS.isStyleOnlyCell = isStyleOnlyCell;
	CSS.saveContentSkipStyleOnly = saveContentSkipStyleOnly;
})(window);
