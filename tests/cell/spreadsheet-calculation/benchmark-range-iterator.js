/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

$( function () {

	function RangeTopBottomIterator(){
		this.size = 0;
		this.rangesTop = null;
		this.indexTop = 0;
		this.rangesBottom = null;
		this.indexBottom = 0;
		this.lastRow = -1;
		this.tree = null;
		this.treeWithId = null;
		this.mmap = null;
		this.mmapCache = null;
		this.mmapCacheLeftLast = Number.MAX_VALUE;
		this.mmapCacheRightLast = -1;
		this.mmapCacheLeft = null;
		this.mmapCacheRight = null;
	}
	RangeTopBottomIterator.prototype.init = function (arr, fGetRanges) {
		window.rr = [];
		var rangesTop = this.rangesTop = [];
		var rangesBottom = this.rangesBottom = [];
		var nextId = 0;
		this.size = arr.length;
		arr.forEach(function(elem) {
			var ranges = fGetRanges(elem);
			for (var i = 0; i < ranges.length; i++) {
				window.rr.push([ranges[i].r1, ranges[i].c1, ranges[i].r2, ranges[i].c2]);
				var rangeElem = {id: nextId++, bbox: ranges[i], data: elem, isInsert: false};
				rangesTop.push(rangeElem);
				rangesBottom.push(rangeElem);
			}
		});
		//Array.sort is stable in all browsers
		//https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort#browser_compatibility
		this.rangesTop.sort(RangeTopBottomIterator.prototype.compareByLeftTop);
		this.rangesBottom.sort(RangeTopBottomIterator.prototype.compareByRightBottom);
		this.reset();
	};
	RangeTopBottomIterator.prototype.compareByLeftTop = function (a, b) {
		return Asc.Range.prototype.compareByLeftTop(a.bbox, b.bbox);
	};
	RangeTopBottomIterator.prototype.compareByRightBottom = function (a, b) {
		return Asc.Range.prototype.compareByRightBottom(a.bbox, b.bbox);
	};
	RangeTopBottomIterator.prototype.getSize = function () {
		return this.size;
	};
	RangeTopBottomIterator.prototype.reset = function() {
		this.indexTop = 0;
		this.indexBottom = 0;
		this.lastRow = -1;
		if(this.tree) {
			var intervals = this.tree.searchNodes(0, AscCommon.gc_nMaxCol0);
			for(var i = 0; i < intervals.length; ++i) {
				intervals[i].isInsert = false;
			}
		}
		this.tree = new AscCommon.DataIntervalTree();
		if(this.treeWithId) {
			var intervals = this.treeWithId.searchNodes(0, AscCommon.gc_nMaxCol0);
			for (var i = 0; i < intervals.length; ++i) {
				intervals[i].isInsert = false;
			}
		}
		this.treeWithId = new AscCommon.DataIntervalTreeWithId();
		if (this.mmap) {
			this.mmap.forEach(function(rangeElem) {
				rangeElem.isInsert = false;
			});
		}
		this.mmap = new Map();
		this.mmapCache = null;
		this.mmapCacheLeft = null;
		this.mmapCacheRight = null;
	};
	RangeTopBottomIterator.prototype.getWithForEach = function(row, col) {
		var res = [];
		this.rangesTop.forEach(function(elem) {
			if (elem.bbox.r1 <= row && row <= elem.bbox.r2 && elem.bbox.c1 <= col && col <= elem.bbox.c2) {
				res.push(elem.data);
			}
		});
		return res;
	};
	RangeTopBottomIterator.prototype.getWithTree = function(row, col) {
		var res = [];
		if (this.lastRow > row) {
			this.reset();
		}
		var rangeElem;
		while (this.indexTop < this.rangesTop.length && row >= this.rangesTop[this.indexTop].bbox.r1) {
			rangeElem = this.rangesTop[this.indexTop++];
			if (row <= rangeElem.bbox.r2) {
				rangeElem.isInsert = true;
				this.tree.insert(rangeElem.bbox.c1, rangeElem.bbox.c2, rangeElem.data);
			}
		}
		while (this.indexBottom < this.rangesBottom.length && row > this.rangesBottom[this.indexBottom].bbox.r2) {
			rangeElem = this.rangesBottom[this.indexBottom++];
			if (rangeElem.isInsert) {
				rangeElem.isInsert = false;
				this.tree.remove(rangeElem.bbox.c1, rangeElem.bbox.c2, rangeElem.data);
			}
		}
		var intervals = this.tree.searchNodes(col, col);
		for(var i = 0; i < intervals.length; ++i) {
			res.push(intervals[i].data);
		}
		this.lastRow = row;
		return res;
	};
	RangeTopBottomIterator.prototype.getWithTreeWithId = function(row, col) {
		var res = [];
		if (this.lastRow > row) {
			this.reset();
		}
		var rangeElem;
		while (this.indexTop < this.rangesTop.length && row >= this.rangesTop[this.indexTop].bbox.r1) {
			rangeElem = this.rangesTop[this.indexTop++];
			if (row <= rangeElem.bbox.r2) {
				rangeElem.isInsert = true;
				this.treeWithId.insert(rangeElem.bbox.c1, rangeElem.bbox.c2, rangeElem.data, rangeElem.id);
			}
		}
		while (this.indexBottom < this.rangesBottom.length && row > this.rangesBottom[this.indexBottom].bbox.r2) {
			rangeElem = this.rangesBottom[this.indexBottom++];
			if (rangeElem.isInsert) {
				rangeElem.isInsert = false;
				this.treeWithId.remove(rangeElem.bbox.c1, rangeElem.bbox.c2, rangeElem.data, rangeElem.id);
			}
		}
		var intervals = this.treeWithId.searchNodes(col, col);
		for(var i = 0; i < intervals.length; ++i) {
			res.push(intervals[i].data);
		}
		this.lastRow = row;
		return res;
	};
	RangeTopBottomIterator.prototype.getWithMap = function(row, col) {
		var res = [];
		if (this.lastRow > row) {
			this.reset();
		}
		var rangeElem;
		while (this.indexTop < this.rangesTop.length && row >= this.rangesTop[this.indexTop].bbox.r1) {
			rangeElem = this.rangesTop[this.indexTop++];
			if (row <= rangeElem.bbox.r2) {
				rangeElem.isInsert = true;
				this.mmap.set(rangeElem.id, rangeElem);
			}
		}
		while (this.indexBottom < this.rangesBottom.length && row > this.rangesBottom[this.indexBottom].bbox.r2) {
			rangeElem = this.rangesBottom[this.indexBottom++];
			if (rangeElem.isInsert) {
				rangeElem.isInsert = false;
				this.mmap.delete(rangeElem.id);
			}
		}
		this.mmap.forEach(function(rangeElem) {
			if(rangeElem.bbox.c1 <= col && col <= rangeElem.bbox.c2) {
				res.push(rangeElem.data);
			}
		});
		this.lastRow = row;
		return res;
	};
	RangeTopBottomIterator.prototype.getWithMapWithCache = function(row, col) {
		var res = [];
		if (this.lastRow > row) {
			this.reset();
		}
		var rangeElem;
		while (this.indexTop < this.rangesTop.length && row >= this.rangesTop[this.indexTop].bbox.r1) {
			rangeElem = this.rangesTop[this.indexTop++];
			if (row <= rangeElem.bbox.r2) {
				rangeElem.isInsert = true;
				this.mmap.set(rangeElem.id, rangeElem);
				this.mmapCache = null;
			}
		}
		while (this.indexBottom < this.rangesBottom.length && row > this.rangesBottom[this.indexBottom].bbox.r2) {
			rangeElem = this.rangesBottom[this.indexBottom++];
			if (rangeElem.isInsert) {
				rangeElem.isInsert = false;
				this.mmap.delete(rangeElem.id);
				this.mmapCache = null;
			}
		}
		var t = this;
		if(!this.mmapCache) {
			this.mmapCache = [];
			this.mmap.forEach(function(rangeElem) {
				for(var i = rangeElem.bbox.c1; i <= rangeElem.bbox.c2; ++i ) {
					if(!t.mmapCache[i]) {
						t.mmapCache[i] = [];
					}
					t.mmapCache[i].push(rangeElem.data);
				}
			});
		}
		this.lastRow = row;
		return t.mmapCache[col] || [];
	};
	RangeTopBottomIterator.prototype.getWithMapWithCacheWithRange = function(row, col) {
		var res = [];
		if (this.lastRow > row) {
			this.reset();
		}
		var rangeElem;
		while (this.indexTop < this.rangesTop.length && row >= this.rangesTop[this.indexTop].bbox.r1) {
			rangeElem = this.rangesTop[this.indexTop++];
			if (row <= rangeElem.bbox.r2) {
				rangeElem.isInsert = true;
				this.mmap.set(rangeElem.id, rangeElem);
				this.mmapCache = null;
			}
		}
		while (this.indexBottom < this.rangesBottom.length && row > this.rangesBottom[this.indexBottom].bbox.r2) {
			rangeElem = this.rangesBottom[this.indexBottom++];
			if (rangeElem.isInsert) {
				rangeElem.isInsert = false;
				this.mmap.delete(rangeElem.id);
				this.mmapCache = null;
			}
		}
		var t = this;
		if(!this.mmapCache) {
			this.mmapCache = [];
			this.mmapCacheLeft = this.mmapCacheLeftLast;
			this.mmapCacheRight = this.mmapCacheRightLast;
			if (this.mmapCacheLeftLast <= this.mmapCacheRightLast) {
				this.mmap.forEach(function (rangeElem) {
					for (var i = Math.max(t.mmapCacheLeftLast, rangeElem.bbox.c1); i <= Math.min(t.mmapCacheRightLast, rangeElem.bbox.c2); ++i) {
						if (!t.mmapCache[i]) {
							t.mmapCache[i] = [];
						}
						t.mmapCache[i].push(rangeElem.data);
					}
				});
			}
		}
		this.lastRow = row;
		this.mmapCacheRightLast = Math.max(col, this.mmapCacheRightLast);
		this.mmapCacheLeftLast = Math.min(col, this.mmapCacheLeftLast);

		if (this.mmapCacheLeft <= col && col <= this.mmapCacheRight) {
			return t.mmapCache[col] || [];
		} else {
			this.mmap.forEach(function(rangeElem) {
				if(rangeElem.bbox.c1 <= col && col <= rangeElem.bbox.c2) {
					res.push(rangeElem.data);
				}
			});
			return res;
		}
	};
	RangeTopBottomIterator.prototype.get = RangeTopBottomIterator.prototype.getWithMapWithCache;

	function bench(rows, cols, ranges, name){
		test( "Bench:"+name, function () {

			var iterator = new RangeTopBottomIterator();
			var t0, t1, i, j, res;
//---------------------------------------------------
			t0 = performance.now();
			iterator.init(ranges, function(elem){
				return [elem];
			});
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'init');
			//---------------------------------------------------
			t0 = performance.now();
			for(i = 0; i < rows; ++i) {
				for(j = 0; j < cols; ++j) {
					res = iterator.getWithTree(i, j);
				}
			}
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'getWithTree');
			//---------------------------------------------------
			t0 = performance.now();
			for(i = 0; i < rows; ++i) {
				for(j = 0; j < cols; ++j) {
					res = iterator.getWithTreeWithId(i, j);
				}
			}
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'getWithTreeWithId');
			//---------------------------------------------------
			t0 = performance.now();
			for(i = 0; i < rows; ++i) {
				for(j = 0; j < cols; ++j) {
					res = iterator.getWithMap(i, j);
				}
			}
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'getWithMap');
			//---------------------------------------------------
			t0 = performance.now();
			for(i = 0; i < rows; ++i) {
				for(j = 0; j < cols; ++j) {
					res = iterator.getWithMapWithCache(i, j);
				}
			}
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'getWithMapWithCache');
			//---------------------------------------------------
			t0 = performance.now();
			for(i = 0; i < rows; ++i) {
				for(j = 0; j < cols; ++j) {
					res = iterator.getWithMapWithCacheWithRange(i, j);
				}
			}
			t1 = performance.now();
			strictEqual( t1 - t0, 0 , 'getWithMapWithCacheWithRange');
		} );
	}
	module( "Range iterator" );

	test( "Proof: 1384*15", function () {

		var rows = 100;
		var cols = 30;
		var ranges = window.ranges_1384_15.map(function(elem){return new Asc.Range(elem[1], elem[0], elem[3], elem[2]);});
		var iterator = new RangeTopBottomIterator();
		var i, j, res;

		iterator.init(ranges, function(elem){
			return [elem];
		});

		var getForEachRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
				getForEachRes.push(iterator.getWithForEach(i, j));
			}
		}
		var getWithTreeRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
				getWithTreeRes.push(iterator.getWithTree(i, j));
			}
		}
		var getWithTreeWithIdRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
			getWithTreeWithIdRes.push(iterator.getWithTreeWithId(i, j));
			}
		}
		var getWithMapRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
				getWithMapRes.push(iterator.getWithMap(i, j));
			}
		}
		var getWithMapWithCacheRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
				getWithMapWithCacheRes.push(iterator.getWithMapWithCache(i, j));
			}
		}
		var getWithMapWithCacheWithRangeRes = [];
		for(i = 0; i < rows; ++i) {
			for(j = 0; j < cols; ++j) {
				getWithMapWithCacheWithRangeRes.push(iterator.getWithMapWithCacheWithRange(i, j));
			}
		}
		ok( getForEachRes.length > 0,  'getForEachRes-len');
		ok( getForEachRes.length === getWithTreeRes.length,  'getWithTreeRes-len');
		ok( getForEachRes.length === getWithTreeWithIdRes.length,  'getWithTreeWithIdRes-len');
		ok( getForEachRes.length === getWithMapRes.length,  'getWithMapRes-len');
		ok( getForEachRes.length === getWithMapWithCacheRes.length,  'getWithMapWithCacheRes-len');
		ok( getForEachRes.length === getWithMapWithCacheWithRangeRes.length,  'getWithMapWithCacheWithRangeRes-len');

		for (i = 0; i < getForEachRes.length; ++i) {
			if(getForEachRes[i].length !== getWithTreeRes[i].length) {
				ok( false,  'getWithTreeRes-len-'+i);
				break;
			}
			for (j = 0; j < getForEachRes[i].length; ++j) {
				if (getForEachRes[i][j].getName() !== getWithTreeRes[i][j].getName()) {
					ok(false, 'getWithTreeRes');
					break;
				}
			}
		}
		for (i = 0; i < getForEachRes.length; ++i) {
			if(getForEachRes[i].length !== getWithTreeWithIdRes[i].length) {
				ok( false,  'getWithTreeWithIdRes-len-'+i);
				break;
			}
			for (j = 0; j < getForEachRes[i].length; ++j) {
				if (getForEachRes[i][j].getName() !== getWithTreeWithIdRes[i][j].getName()) {
					ok(false, 'getWithTreeWithIdRes');
					break;
				}
			}
		}
		for (i = 0; i < getForEachRes.length; ++i) {
			if(getForEachRes[i].length !== getWithMapRes[i].length) {
				ok( false,  'getWithMapRes-len-'+i);
				break;
			}
			for (j = 0; j < getForEachRes[i].length; ++j) {
				if (getForEachRes[i][j].getName() !== getWithMapRes[i][j].getName()) {
					ok(false, 'getWithMapRes');
					break;
				}
			}
		}
		for (i = 0; i < getForEachRes.length; ++i) {
			if(getForEachRes[i].length !== getWithMapWithCacheRes[i].length) {
				ok( false,  'getWithMapWithCacheRes-len-'+i);
				break;
			}
			for (j = 0; j < getForEachRes[i].length; ++j) {
				if (getForEachRes[i][j].getName() !== getWithMapWithCacheRes[i][j].getName()) {
					ok(false, 'getWithMapWithCacheRes');
					break;
				}
			}
		}
		for (i = 0; i < getForEachRes.length; ++i) {
			if(getForEachRes[i].length !== getWithMapWithCacheWithRangeRes[i].length) {
				ok( false,  'getWithMapWithCacheWithRangeRes-len-'+i);
				break;
			}
			for (j = 0; j < getForEachRes[i].length; ++j) {
				if (getForEachRes[i][j].getName() !== getWithMapWithCacheWithRangeRes[i][j].getName()) {
					ok(false, 'getWithMapWithCacheWithRangeRes');
					break;
				}
			}
		}
	} );

	(function () {
		var rows = -1;
		var cols = -1;
		var ranges = [];
		window.ranges_1384_15.forEach(function(elem){
			var r1 = elem[0];
			var c1 = elem[1];
			var r2 = elem[2];
			var c2 = elem[3];
			rows = Math.max(rows, r2);
			cols = Math.max(cols, c2);
			ranges.push(new Asc.Range(c1, r1, c2, r2));
		});
		bench(rows + 1, cols + 1, ranges, "ranges_1384_15");
	}());

	(function () {
		var rows = -1;
		var cols = -1;
		var ranges = [];
		for(var i = 0; i < 60; ++i) {
			window.ranges_1384_15.forEach(function(elem){
				var r1 = elem[0];
				var c1 = elem[1];
				var r2 = elem[2];
				var c2 = elem[3];
				rows = Math.max(rows, r2);
				cols = Math.max(cols, c2);
				ranges.push(new Asc.Range(c1+ i * 15, r1, c2+ i * 15, r2));
			});
		}
		bench(rows + 1, cols + 1, ranges, "ranges_1384_900");
	}());
} );
