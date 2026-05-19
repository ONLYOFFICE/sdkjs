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

// Legacy SheetMemory xf-shadow migration. Old files carried direct xf
// in the low 24 bits of SheetMemory word 0; this sweep moves them into
// cellStylesByCol once per worksheet at open, before any other reader.
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

	// Fill legacy xf values into `store` for rows the store has not
	// already covered. Safe only at file-open because a CRangeAttrArray
	// null entry cannot distinguish "never set" from "user-cleared".
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

	// File-open-only sweep. Must run BEFORE any user edit or tombstone,
	// otherwise it can resurrect a cleared row. Called once per worksheet
	// from ReadSheetData / SheetDataFromJSON.
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
