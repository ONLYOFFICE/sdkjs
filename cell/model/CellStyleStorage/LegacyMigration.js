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

// Legacy SheetMemory xf-shadow migration.
//
// Pre-split files put direct cell xf into the low 24 bits of
// SheetMemory word 0. This module replays those bits into
// cellStylesByCol exactly once per worksheet, at file open, before any
// consumer (calc, save, range op) reads direct cell styles. After the
// sweep, the legacy shadow is dead code.
(function (window, undefined) {
	var CSS = window['AscCommonExcel'].CellStyleStorage;
	var _internals = CSS._internals;
	var getCellStyleStore = _internals.getCellStyleStore;

	// True if any row in the populated [indexA..indexB] window carries
	// legacy direct xf bits. Lets the sweep skip clean columns without
	// allocating an empty store.
	function _columnHasLegacyXfBits(sheetMemory) {
		if (!sheetMemory || !sheetMemory.dataBuffer || !sheetMemory.dataInt32) {
			return false;
		}
		var minRow = sheetMemory.indexA;
		var maxRow = sheetMemory.indexB;
		if (typeof minRow !== 'number' || typeof maxRow !== 'number'
			|| maxRow < minRow || minRow < 0) {
			return false;
		}
		for (var r = minRow; r <= maxRow; r++) {
			if ((sheetMemory.getInt32(r, 0) & 0xFFFFFF) > 0) {
				return true;
			}
		}
		return false;
	}

	// Merge legacy xf values into `store`, filling ONLY rows that have no
	// store entry yet. Rows already present in the store are untouched.
	//
	// CRangeAttrArray treats `null` identically for "never set" and
	// "explicitly cleared" -- the merge is safe only because this runs
	// before any clear can happen (see migration-only contract on
	// hydrateAllColumnsFromSheetMemory). Same-column mixed legacy case:
	// the per-cell open-time mirror may have migrated some rows already;
	// this fills the rest.
	function _mergeLegacyXfBitsIntoStore(sheetMemory, store) {
		if (!sheetMemory || !sheetMemory.dataBuffer || !sheetMemory.dataInt32) {
			return;
		}
		var minRow = sheetMemory.indexA;
		var maxRow = sheetMemory.indexB;
		if (typeof minRow !== 'number' || typeof maxRow !== 'number'
			|| maxRow < minRow || minRow < 0) {
			return;
		}
		for (var r = minRow; r <= maxRow; r++) {
			var xfIdx = sheetMemory.getInt32(r, 0) & 0xFFFFFF;
			if (xfIdx > 0 && store.get(r) === null) {
				store.setRange(r, r, xfIdx);
			}
		}
	}

	// Eager post-open sweep. Walks every populated column and merges
	// legacy xf bits into the column's store; allocates a store only when
	// there is real work to do. Bounded to `ws.cellsByCol.length`, so a
	// sparse worksheet stays sparse.
	//
	// Migration-only contract: MUST run at file-open completion, BEFORE
	// any user edit or tombstone. After edits begin, `store.get(r) ===
	// null` can mean either "never styled" or "explicitly cleared"; a
	// repeat sweep could resurrect a legacy value and undo a user clear.
	// Idempotent if re-run immediately (every row gets a store entry on
	// the first pass), but do not use as a late-session repair tool.
	//
	// Wired into binary and JSON open paths once per worksheet at the
	// completion of ReadSheetData / SheetDataFromJSON.
	function hydrateAllColumnsFromSheetMemory(ws) {
		if (!ws || !ws.cellsByCol || !ws.cellStylesByCol) {
			return;
		}
		var cells = ws.cellsByCol;
		var len = cells.length;
		for (var c = 0; c < len; c++) {
			var sm = cells[c];
			if (!sm) {
				continue;
			}
			if (!_columnHasLegacyXfBits(sm)) {
				continue;
			}
			var store = getCellStyleStore(ws, c, true);
			_mergeLegacyXfBitsIntoStore(sm, store);
		}
	}

	CSS.hydrateAllColumnsFromSheetMemory = hydrateAllColumnsFromSheetMemory;
})(window);
