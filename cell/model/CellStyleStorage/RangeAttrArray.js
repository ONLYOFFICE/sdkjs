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

(function (window, undefined) {
	var DEFAULT_MAX_ROW = (window['AscCommon'] && window['AscCommon'].gc_nMaxRow0 != null)
		? window['AscCommon'].gc_nMaxRow0
		: 1048575;

	/**
	 * Range-based attribute storage for a single column.
	 *
	 * Stores values (intended for direct cell xfIndex, but the class is type-agnostic)
	 * over inclusive row ranges [r1..r2]. The class is a plain in-memory data
	 * structure. It does not depend on Worksheet, Cell, SheetMemory, or History.
	 *
	 * Invariants:
	 *  - Stored ranges are sorted by r1 and non-overlapping.
	 *  - Adjacent ranges with strictly equal values are merged.
	 *  - null means "no value" and is never stored explicitly (the default tail
	 *    is implicit).
	 *  - All operations are clamped to the closed interval [0..maxRow].
	 *
	 * @constructor
	 * @param {number} [maxRow] Maximum allowed row index (inclusive).
	 */
	function CRangeAttrArray(maxRow) {
		this.maxRow = (maxRow != null) ? maxRow : DEFAULT_MAX_ROW;
		// Array of plain objects { r1, r2, v }. r1, r2 inclusive; v never null.
		this._ranges = [];
	}

	/**
	 * Returns true if no value is stored.
	 * @returns {boolean}
	 */
	CRangeAttrArray.prototype.isEmpty = function () {
		return this._ranges.length === 0;
	};

	/**
	 * First row carrying any value, or -1 if empty.
	 * @returns {number}
	 */
	CRangeAttrArray.prototype.firstRow = function () {
		return this._ranges.length === 0 ? -1 : this._ranges[0].r1;
	};

	/**
	 * Last row carrying any value, or -1 if empty.
	 * @returns {number}
	 */
	CRangeAttrArray.prototype.lastRow = function () {
		return this._ranges.length === 0 ? -1 : this._ranges[this._ranges.length - 1].r2;
	};

	/**
	 * Lookup the value at a given row.
	 * @param {number} row
	 * @returns {*} the stored value, or null if no value is set at row.
	 */
	CRangeAttrArray.prototype.get = function (row) {
		if (row < 0 || row > this.maxRow) {
			return null;
		}
		var i = this._lowerBound(row);
		if (i < this._ranges.length) {
			var r = this._ranges[i];
			if (r.r1 <= row && row <= r.r2) {
				return r.v;
			}
		}
		return null;
	};

	/**
	 * Set value over inclusive [r1..r2]. Passing null clears that range.
	 * @param {number} r1
	 * @param {number} r2
	 * @param {*} value any non-null value, or null to clear.
	 */
	CRangeAttrArray.prototype.setRange = function (r1, r2, value) {
		var clamped = this._clamp(r1, r2);
		if (!clamped) {
			return;
		}
		r1 = clamped[0];
		r2 = clamped[1];
		if (value === undefined) {
			value = null;
		}

		var i = this._lowerBound(r1);
		var removeFrom = i;
		var removeTo = i;
		var leftSplit = null;
		var rightSplit = null;

		while (removeTo < this._ranges.length) {
			var cur = this._ranges[removeTo];
			if (cur.r1 > r2) {
				break;
			}
			if (cur.r1 < r1) {
				leftSplit = { r1: cur.r1, r2: r1 - 1, v: cur.v };
			}
			if (cur.r2 > r2) {
				rightSplit = { r1: r2 + 1, r2: cur.r2, v: cur.v };
			}
			removeTo++;
		}

		var replace = [];
		if (leftSplit) {
			replace.push(leftSplit);
		}
		if (value !== null) {
			replace.push({ r1: r1, r2: r2, v: value });
		}
		if (rightSplit) {
			replace.push(rightSplit);
		}

		var args = [removeFrom, removeTo - removeFrom];
		for (var k = 0; k < replace.length; k++) {
			args.push(replace[k]);
		}
		Array.prototype.splice.apply(this._ranges, args);

		this._mergeNeighbors(removeFrom, removeFrom + replace.length);
	};

	/**
	 * Clear inclusive [r1..r2]. Equivalent to setRange(r1, r2, null).
	 */
	CRangeAttrArray.prototype.clearRange = function (r1, r2) {
		this.setRange(r1, r2, null);
	};

	/**
	 * Insert `count` rows at `start`. Existing ranges at or after `start` shift
	 * by +count; ranges that straddle `start` split (rows below `start` stay,
	 * rows >= `start` shift). The newly inserted rows are left empty.
	 * @param {number} start
	 * @param {number} count
	 */
	CRangeAttrArray.prototype.insertRows = function (start, count) {
		if (count <= 0 || start > this.maxRow) {
			return;
		}
		if (start < 0) {
			start = 0;
		}
		for (var k = 0; k < this._ranges.length; k++) {
			var r = this._ranges[k];
			if (r.r2 < start) {
				continue;
			}
			if (r.r1 >= start) {
				r.r1 += count;
				r.r2 += count;
			} else {
				// Range straddles start: split into [r.r1..start-1] and [start+count..r.r2+count].
				var tail = { r1: start + count, r2: r.r2 + count, v: r.v };
				r.r2 = start - 1;
				this._ranges.splice(k + 1, 0, tail);
				k++;
			}
		}
		this._clampTailToMaxRow();
	};

	/**
	 * Delete `count` rows starting at `start`. Surviving rows above the deletion
	 * shift by -count. Ranges that straddled the deletion may merge across the
	 * gap if their value matches.
	 * @param {number} start
	 * @param {number} count
	 */
	CRangeAttrArray.prototype.deleteRows = function (start, count) {
		if (count <= 0) {
			return;
		}
		if (start < 0) {
			count += start;
			start = 0;
			if (count <= 0) {
				return;
			}
		}
		if (start > this.maxRow) {
			return;
		}
		var end = start + count - 1;

		var newRanges = [];
		for (var k = 0; k < this._ranges.length; k++) {
			var r = this._ranges[k];
			if (r.r2 < start) {
				newRanges.push(r);
			} else if (r.r1 > end) {
				newRanges.push({ r1: r.r1 - count, r2: r.r2 - count, v: r.v });
			} else {
				if (r.r1 < start) {
					newRanges.push({ r1: r.r1, r2: start - 1, v: r.v });
				}
				if (r.r2 > end) {
					newRanges.push({ r1: start, r2: r.r2 - count, v: r.v });
				}
			}
		}
		this._ranges = newRanges;
		this._mergeAll();
	};

	/**
	 * Copy values for `count` consecutive rows from `other` (starting at fromR1)
	 * into this storage starting at toR1. Destination range is cleared first.
	 * Safe when this === other.
	 * @param {CRangeAttrArray} other
	 * @param {number} fromR1
	 * @param {number} toR1
	 * @param {number} count
	 */
	CRangeAttrArray.prototype.copyFrom = function (other, fromR1, toR1, count) {
		if (count <= 0) {
			return;
		}
		var fromR2 = fromR1 + count - 1;
		var toR2 = toR1 + count - 1;
		var offset = toR1 - fromR1;

		var snapshot = [];
		other.iter(fromR1, fromR2, function (lo, hi, v) {
			snapshot.push({ r1: lo, r2: hi, v: v });
		});

		this.clearRange(toR1, toR2);
		for (var k = 0; k < snapshot.length; k++) {
			var s = snapshot[k];
			this.setRange(s.r1 + offset, s.r2 + offset, s.v);
		}
	};

	/**
	 * Replicate the value at `fromRow` across rows [toRow..toRow+toCount-1].
	 * Equivalent to SheetMemory.setAreaByRow for a single source row.
	 * @param {number} fromRow
	 * @param {number} toRow
	 * @param {number} toCount
	 */
	CRangeAttrArray.prototype.setAreaByRow = function (fromRow, toRow, toCount) {
		if (toCount <= 0) {
			return;
		}
		var v = this.get(fromRow);
		if (v === null) {
			this.clearRange(toRow, toRow + toCount - 1);
		} else {
			this.setRange(toRow, toRow + toCount - 1, v);
		}
	};

	/**
	 * Iterate forward over stored runs intersecting [r1..r2].
	 * Callback receives (lo, hi, value) clamped to [r1..r2]. Return false to stop.
	 * @param {number} r1
	 * @param {number} r2
	 * @param {function(number, number, *):(boolean|undefined)} callback
	 */
	CRangeAttrArray.prototype.iter = function (r1, r2, callback) {
		var clamped = this._clamp(r1, r2);
		if (!clamped) {
			return;
		}
		r1 = clamped[0];
		r2 = clamped[1];
		var i = this._lowerBound(r1);
		while (i < this._ranges.length) {
			var r = this._ranges[i];
			if (r.r1 > r2) {
				break;
			}
			var lo = r.r1 < r1 ? r1 : r.r1;
			var hi = r.r2 > r2 ? r2 : r.r2;
			if (callback(lo, hi, r.v) === false) {
				return;
			}
			i++;
		}
	};

	/**
	 * Apply `mapFn(value, lo, hi)` to every stored run intersecting [r1..r2].
	 * If the result differs from the original value, that segment is rewritten
	 * to the new value (or cleared, if the function returns null). Existing runs
	 * outside [r1..r2] are not touched.
	 *
	 * Used for protected-sheet "clear except locked" semantics.
	 *
	 * @param {number} r1
	 * @param {number} r2
	 * @param {function(*, number, number):*} mapFn
	 */
	CRangeAttrArray.prototype.mapRange = function (r1, r2, mapFn) {
		var clamped = this._clamp(r1, r2);
		if (!clamped) {
			return;
		}
		r1 = clamped[0];
		r2 = clamped[1];

		var segs = [];
		this.iter(r1, r2, function (lo, hi, v) {
			segs.push({ lo: lo, hi: hi, v: v });
		});
		for (var k = 0; k < segs.length; k++) {
			var s = segs[k];
			var nv = mapFn(s.v, s.lo, s.hi);
			if (nv !== s.v) {
				this.setRange(s.lo, s.hi, nv);
			}
		}
	};

	/**
	 * Deep clone this storage.
	 * @returns {CRangeAttrArray}
	 */
	CRangeAttrArray.prototype.clone = function () {
		var c = new CRangeAttrArray(this.maxRow);
		for (var i = 0; i < this._ranges.length; i++) {
			var r = this._ranges[i];
			c._ranges.push({ r1: r.r1, r2: r.r2, v: r.v });
		}
		return c;
	};

	/**
	 * @returns {number} number of stored runs (mainly for tests).
	 */
	CRangeAttrArray.prototype.getRangeCount = function () {
		return this._ranges.length;
	};

	// --- internal helpers ---

	CRangeAttrArray.prototype._clamp = function (r1, r2) {
		if (r1 > r2) {
			return null;
		}
		if (r1 < 0) {
			r1 = 0;
		}
		if (r2 > this.maxRow) {
			r2 = this.maxRow;
		}
		if (r1 > r2) {
			return null;
		}
		return [r1, r2];
	};

	// First index i with this._ranges[i].r2 >= row (or length if none).
	CRangeAttrArray.prototype._lowerBound = function (row) {
		var lo = 0;
		var hi = this._ranges.length;
		while (lo < hi) {
			var mid = (lo + hi) >> 1;
			if (this._ranges[mid].r2 < row) {
				lo = mid + 1;
			} else {
				hi = mid;
			}
		}
		return lo;
	};

	// First index i with this._ranges[i].r1 > row (or length if none).
	CRangeAttrArray.prototype._upperBound = function (row) {
		var lo = 0;
		var hi = this._ranges.length;
		while (lo < hi) {
			var mid = (lo + hi) >> 1;
			if (this._ranges[mid].r1 <= row) {
				lo = mid + 1;
			} else {
				hi = mid;
			}
		}
		return lo;
	};

	// Merge any same-value adjacent runs whose join sits in [from-1..to].
	CRangeAttrArray.prototype._mergeNeighbors = function (from, to) {
		var start = from > 0 ? from - 1 : 0;
		var i = start;
		while (i + 1 < this._ranges.length && i <= to) {
			var a = this._ranges[i];
			var b = this._ranges[i + 1];
			if (a.r2 + 1 === b.r1 && a.v === b.v) {
				a.r2 = b.r2;
				this._ranges.splice(i + 1, 1);
				to--;
			} else {
				i++;
			}
		}
	};

	CRangeAttrArray.prototype._mergeAll = function () {
		var i = 0;
		while (i + 1 < this._ranges.length) {
			var a = this._ranges[i];
			var b = this._ranges[i + 1];
			if (a.r2 + 1 === b.r1 && a.v === b.v) {
				a.r2 = b.r2;
				this._ranges.splice(i + 1, 1);
			} else {
				i++;
			}
		}
	};

	CRangeAttrArray.prototype._clampTailToMaxRow = function () {
		while (this._ranges.length) {
			var last = this._ranges[this._ranges.length - 1];
			if (last.r1 > this.maxRow) {
				this._ranges.pop();
				continue;
			}
			if (last.r2 > this.maxRow) {
				last.r2 = this.maxRow;
			}
			break;
		}
	};

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CRangeAttrArray = CRangeAttrArray;
})(window);
