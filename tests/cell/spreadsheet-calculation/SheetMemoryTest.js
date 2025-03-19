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

$(function () {
	let SheetMemory = AscCommonExcel.SheetMemory;

	QUnit.test("Test: \"checkIndex\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		assert.strictEqual(sheetMemory.hasIndex(10), false);
		assert.strictEqual(sheetMemory.getMaxIndex(), -1);
		assert.strictEqual(sheetMemory.hasIndex(200), false);

		sheetMemory.checkIndex(10);
		assert.strictEqual(sheetMemory.hasIndex(9), false);
		assert.strictEqual(sheetMemory.hasIndex(10), true);
		assert.strictEqual(sheetMemory.hasIndex(11), false);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 10);

		sheetMemory.checkIndex(200);
		assert.strictEqual(sheetMemory.hasIndex(200), false);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 100);

		sheetMemory.checkIndex(15);
		assert.strictEqual(sheetMemory.hasIndex(15), true);
		assert.strictEqual(sheetMemory.hasIndex(16), true);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 100);

		sheetMemory.checkIndex(5);
		assert.strictEqual(sheetMemory.hasIndex(4), false);
		assert.strictEqual(sheetMemory.hasIndex(5), true);
		assert.strictEqual(sheetMemory.getMinIndex(), 5);
		assert.strictEqual(sheetMemory.getMaxIndex(), 100);

		let allocThreshOld = sheetMemory.getMinIndex() + sheetMemory.getAllocatedCount() - 1;
		sheetMemory.checkIndex(allocThreshOld);
		sheetMemory.setUint8(allocThreshOld, 0, 1);
		assert.strictEqual(sheetMemory.getMinIndex(), 5);
		assert.strictEqual(sheetMemory.getMaxIndex(), allocThreshOld);
		assert.strictEqual(sheetMemory.getUint8(allocThreshOld, 0), 1);
	});

	QUnit.test("Test: \"clone\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(5);
		sheetMemory.checkIndex(15);
		sheetMemory.setUint8(5, 0, 1);
		sheetMemory.setUint8(10, 0, 2);
		sheetMemory.setUint8(15, 0, 3);
		let sheetMemory2 = sheetMemory.clone();
		assert.strictEqual(sheetMemory2.getMinIndex(), 5);
		assert.strictEqual(sheetMemory2.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory2.getUint8(5, 0), 1);
		assert.strictEqual(sheetMemory2.getUint8(10, 0), 2);
		assert.strictEqual(sheetMemory2.getUint8(15, 0), 3);
	});

	QUnit.test("Test: \"deleteRange\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(12);
		sheetMemory.checkIndex(20);
		sheetMemory.setUint8(12, 0, 1);
		sheetMemory.setUint8(13, 0, 2);
		sheetMemory.setUint8(14, 0, 3);
		sheetMemory.setUint8(15, 0, 4);
		sheetMemory.setUint8(16, 0, 5);
		sheetMemory.setUint8(17, 0, 6);
		sheetMemory.setUint8(18, 0, 7);
		sheetMemory.setUint8(19, 0, 8);
		sheetMemory.setUint8(20, 0, 9);

		sheetMemory.deleteRange(30, 100);
		assert.strictEqual(sheetMemory.getMinIndex(), 12);
		assert.strictEqual(sheetMemory.getMaxIndex(), 20);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(16, 0), 5);
		assert.strictEqual(sheetMemory.getUint8(17, 0), 6);
		assert.strictEqual(sheetMemory.getUint8(18, 0), 7);
		assert.strictEqual(sheetMemory.getUint8(19, 0), 8);
		assert.strictEqual(sheetMemory.getUint8(20, 0), 9);

		sheetMemory.deleteRange(2, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 18);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 5);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 6);
		assert.strictEqual(sheetMemory.getUint8(16, 0), 7);
		assert.strictEqual(sheetMemory.getUint8(17, 0), 8);
		assert.strictEqual(sheetMemory.getUint8(18, 0), 9);

		sheetMemory.deleteRange(8, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 5);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 6);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 7);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 8);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 9);

		sheetMemory.deleteRange(14, 4);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 13);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 5);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 6);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 7);

		sheetMemory.deleteRange(10, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 11);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 6);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 7);

		sheetMemory.deleteRange(8, 20);
		assert.strictEqual(sheetMemory.getMinIndex(), -1);
		assert.strictEqual(sheetMemory.getMaxIndex(), -1);
	});

	QUnit.test("Test: \"insertRange\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(8);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(8, 0, 1);
		sheetMemory.setUint8(9, 0, 2);
		sheetMemory.setUint8(10, 0, 3);
		sheetMemory.setUint8(11, 0, 4);
		sheetMemory.setUint8(12, 0, 5);

		sheetMemory.insertRange(30, 100);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 5);

		sheetMemory.insertRange(8, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 14);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 5);

		sheetMemory.insertRange(11, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 16);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(16, 0), 5);

		sheetMemory.insertRange(15, 4);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 20);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(16, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(17, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(18, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(19, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(20, 0), 5);
	});

	QUnit.test("Test: \"copyRange1\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(10);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(10, 0, 1);
		sheetMemory.setUint8(11, 0, 2);
		sheetMemory.setUint8(12, 0, 3);

		let sheetMemory2 = new SheetMemory(2, 100);
		sheetMemory2.checkIndex(8);
		sheetMemory2.checkIndex(12);
		sheetMemory2.setUint8(8, 0, 1);
		sheetMemory2.setUint8(9, 0, 2);
		sheetMemory2.setUint8(10, 0, 3);
		sheetMemory2.setUint8(11, 0, 4);
		sheetMemory2.setUint8(12, 0, 5);

		sheetMemory.copyRange(sheetMemory2, 7, 8, 7);
		assert.strictEqual(sheetMemory.getMinIndex(), 9);
		assert.strictEqual(sheetMemory.getMaxIndex(), 13);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 5);

		sheetMemory.copyRange(sheetMemory2, 10, 9, 4);
		assert.strictEqual(sheetMemory.getMinIndex(), 9);
		assert.strictEqual(sheetMemory.getMaxIndex(), 13);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 5);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 5);
	});

	QUnit.test("Test: \"copyRange2\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(10);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(10, 0, 1);
		sheetMemory.setUint8(11, 0, 2);
		sheetMemory.setUint8(12, 0, 3);

		sheetMemory.copyRange(sheetMemory, 10, 13, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 3);

		sheetMemory.copyRange(sheetMemory, 10, 12, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 3);

		sheetMemory.copyRange(sheetMemory, 10, 11, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 3);

		sheetMemory.copyRange(sheetMemory, 10, 10, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 3);


		sheetMemory.copyRange(sheetMemory, 10, 9, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 9);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 3);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 3);
	});

	QUnit.test("Test: \"copyRange3\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(10);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(10, 0, 1);
		sheetMemory.setUint8(11, 0, 2);
		sheetMemory.setUint8(12, 0, 3);

		sheetMemory.copyRange(sheetMemory, 9, 11, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 1);

		sheetMemory.copyRange(sheetMemory, 12, 11, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
	});

	QUnit.test("Test: \"copyRangeByChunk\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(8);
		sheetMemory.checkIndex(9);
		sheetMemory.setUint8(8, 0, 1);
		sheetMemory.setUint8(9, 0, 2);

		sheetMemory.copyRangeByChunk(8, 1, 10, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 11);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);

		sheetMemory.copyRangeByChunk(8, 5, 13, 7);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 19);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(16, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(17, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(18, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(19, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(20, 0), 1);
	});

	QUnit.test("Test: \"clear1\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(8);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(8, 0, 1);
		sheetMemory.setUint8(9, 0, 2);
		sheetMemory.setUint8(10, 0, 3);
		sheetMemory.setUint8(11, 0, 4);
		sheetMemory.setUint8(12, 0, 5);

		sheetMemory.clear(5, 15);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
	});

	QUnit.test("Test: \"clear2\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(8);
		sheetMemory.checkIndex(12);
		sheetMemory.setUint8(8, 0, 1);
		sheetMemory.setUint8(9, 0, 2);
		sheetMemory.setUint8(10, 0, 3);
		sheetMemory.setUint8(11, 0, 4);
		sheetMemory.setUint8(12, 0, 5);

		sheetMemory.clear(8, 11);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 4);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 5);

		sheetMemory.clear(11, 13);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 12);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
	});

	testForeachNoEmpty("Test: SweepLineRowIterator step" + 1, 1, 1);
	testForeachNoEmpty("Test: SweepLineRowIterator step" + 2, 1, 1);

	QUnit.module("SheetMemory");
	function testForeachNoEmpty(name, offset, stepRow) {
		QUnit.test(name, function (assert) {
			// console.profile('testForeachNoEmpty');

			// let rowsTest = offset + 2 + stepRow;
			// let colsTest = 2;
			// let dataTest = [0,2,3,4,0,0,0,0];
			// testCellsByCol(dataTest, rowsTest + 1, colsTest, offset, stepRow, assert);

			let rows = offset + 3 + stepRow;//stepRow rows are needed to properly clean
			let cols = 3;
			let baseLen = (rows - stepRow) * cols;//last row for correct cleaning
			let base = [...Array(baseLen + 1).keys()].slice(1);
			let data = new Array(rows * cols);
			data.fill(0);
			let iterations = Math.pow(2, baseLen);
			for (let i = 0; i < iterations; ++i) {
				for (let j = 0; j < baseLen; ++j) {
					let bit = ((i >> j) % 2 !== 0);
					data[j] = bit ? base[j] : 0;
				}
				testCellsByCol(data, rows, cols, offset, stepRow, assert);
			}

			// console.profileEnd('testForeachNoEmpty');

			assert.ok(true);
		});
	}

	function testCellsByCol(data, rows, cols, offset, stepRow, assert) {
		let res = '';
		let testData = getTestDataFromArray(data, rows, cols, offset, stepRow);
		let r1 = offset;

		let sweepLine = new AscCommonExcel.SweepLineRowIterator();
		sweepLine.init(testData.cellsByCol, r1, 0, cols);
		for (let i = r1; i < rows; i += stepRow) {
			sweepLine.setRow(i);
			while (sweepLine.nextCol()) {
				res += `${i}-${sweepLine.col}-${sweepLine.colData.getUint8(i, 0)};`;
			}
		}
		//many asserts processes very slow
		if (res !== testData.expected) {
			assert.strictEqual(res, testData.expected, JSON.stringify(data));
		}
		if (sweepLine.colDatas.length !== sweepLine.colDatasIndex) {
			assert.strictEqual(sweepLine.colDatas.length, sweepLine.colDatasIndex, "colDatas");
		}
		if (sweepLine.toInsert.length !== sweepLine.toInsertIndex) {
			assert.strictEqual(sweepLine.toInsert.length, sweepLine.toInsertIndex, "toInsert");
		}
		if (sweepLine.toDelete.length !== sweepLine.toDeleteIndex) {
			assert.strictEqual(sweepLine.toDelete.length, sweepLine.toDeleteIndex, "toDelete");
		}
		if (sweepLine.r1.length !== sweepLine.r1Index) {
			assert.strictEqual(sweepLine.r1.length, sweepLine.r1Index, "r1");
		}
		if (sweepLine.r2.length !== sweepLine.r2Index) {
			assert.strictEqual(sweepLine.r2.length, sweepLine.r2Index, "r2");
		}
	}

	function getTestDataFromArray(data, rows, cols, offset, stepRow) {
		let expected = "";
		let cellsByCol = Array.from(Array(cols), () => {
			return new SheetMemory(2, rows);
		});
		let dataLen = data.length;
		for (let i = 0; i < dataLen; ++i) {
			if (data[i] > 0) {
				let row = Math.trunc(i / cols);
				let col = i % cols;
				let sheetMemory = cellsByCol[col];
				sheetMemory.checkIndex(row);
				sheetMemory.setUint8(row, 0, data[i]);
			}
		}
		for (let i = 0; i < cellsByCol.length; ++i) {
			if (cellsByCol[i].getAllocatedCount() === 0) {
				delete cellsByCol[i];
			}
		}
		for (let i = offset * cols; i < dataLen; ++i) {
			let row = Math.trunc(i / cols);
			let col = i % cols;
			if(row % stepRow === 0) {
				let sheetMemory = cellsByCol[col];
				if (sheetMemory && sheetMemory.hasIndex(row)) {
					expected += `${row}-${col}-${data[i]};`;
				}
			}
		}
		return {cellsByCol, expected}
	}
});
