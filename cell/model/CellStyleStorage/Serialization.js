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
//
// cellStylesByCol is the sole source of truth for direct cell xf at
// write and read time. No `cell.xfs` fallback for writers, no
// SheetMemory low-24-bit fallback for the loader: LegacyMigration
// runs once per worksheet at file open, before any consumer here.
(function (window, undefined) {
	var CSS = window['AscCommonExcel'].CellStyleStorage;
	var _internals = CSS._internals;
	var _isDataInitRow = _internals.isDataInitRow;

	// Writer-side xfIndex resolver. `cell` is kept in the signature for
	// call-site stability; it is not consulted.
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

	// Same as getWriterCellXfIndex but returns the CellXfs object for
	// `stylesForWrite.add(...)`, or null when there is no direct style.
	function getWriterCellXfs(ws, row, col, cell) {
		var idx = getWriterCellXfIndex(ws, row, col, cell);
		if (idx <= 0) {
			return null;
		}
		var cache = window['AscCommonExcel'].g_StyleCache;
		return cache ? cache.getXf(idx) : null;
	}

	// Read-time xfIndex resolver for Cell.loadContent. Returns 0 when no
	// store or entry exists. No SheetMemory fallback (LegacyMigration
	// runs first at file open).
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

	// Read-only direct-xf resolver for protection checks. Returns the
	// CellXfs at (row, col) when a positive direct xfIndex is present,
	// else null. Never materializes a Cell. Data cells should reach
	// their direct style via the existing `_foreachNoEmpty` / `cell.xfs`
	// path (which routes through `resolveLoadXfIndex`).
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

	// True when (row, col) has no SheetMemory data init flag, i.e. is a
	// pure style-only cell. Used by protection checks; data cells are
	// covered by the `_foreachNoEmpty` / `cell.xfs` path.
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

	CSS.getWriterCellXfIndex = getWriterCellXfIndex;
	CSS.getWriterCellXfs = getWriterCellXfs;
	CSS.resolveLoadXfIndex = resolveLoadXfIndex;
	CSS.getDirectCellXfs = getDirectCellXfs;
	CSS.isStyleOnlyCell = isStyleOnlyCell;
})(window);
