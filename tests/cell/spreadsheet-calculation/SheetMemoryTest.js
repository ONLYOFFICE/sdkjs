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

	QUnit.module("SheetMemory");
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
		assert.strictEqual(sheetMemory.getMaxIndex(), 10);

		sheetMemory.checkIndex(15);
		assert.strictEqual(sheetMemory.hasIndex(15), true);
		assert.strictEqual(sheetMemory.hasIndex(16), false);
		assert.strictEqual(sheetMemory.getMinIndex(), 10);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);

		sheetMemory.checkIndex(5);
		assert.strictEqual(sheetMemory.hasIndex(4), false);
		assert.strictEqual(sheetMemory.hasIndex(5), true);
		assert.strictEqual(sheetMemory.getMinIndex(), 5);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);

		let allocThreshOld = sheetMemory.getMinIndex() + sheetMemory.getAllocatedCount();
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

	QUnit.test("Test: \"setAreaByRow\"", function (assert) {

		let sheetMemory = new SheetMemory(2, 100);
		sheetMemory.checkIndex(8);
		sheetMemory.checkIndex(9);
		sheetMemory.setUint8(8, 0, 1);
		sheetMemory.setUint8(9, 0, 2);

		sheetMemory.setAreaByRow(8, 10, 2);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 11);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);

		sheetMemory.setAreaByRow(8, 13, 3);
		assert.strictEqual(sheetMemory.getMinIndex(), 8);
		assert.strictEqual(sheetMemory.getMaxIndex(), 15);
		assert.strictEqual(sheetMemory.getUint8(8, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(9, 0), 2);
		assert.strictEqual(sheetMemory.getUint8(10, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(11, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(12, 0), 0);
		assert.strictEqual(sheetMemory.getUint8(13, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(14, 0), 1);
		assert.strictEqual(sheetMemory.getUint8(15, 0), 1);
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

	testForeachNoEmpty("Test: SweepLineRowIterator data only step" + 1, 1, 1);
	testForeachNoEmpty("Test: SweepLineRowIterator data only step" + 2, 1, 1);

	testForeachNoEmptyWithAttr("Test: SweepLineRowIterator data and styles step" + 1, 1, 1);
	testForeachNoEmptyWithAttr("Test: SweepLineRowIterator data and styles step" + 2, 1, 1);

	function testForeachNoEmpty(name, offset, stepRow) {
		QUnit.test(name, function (assert) {
			// console.profile('testForeachNoEmpty');

			// let rowsTest = offset + 2 + stepRow;
			// let colsTest = 2;
			// let dataTest = [0,2,3,4,0,0,0,0];
			// testCellsByCol(dataTest, [], rowsTest, colsTest, offset, stepRow, assert);

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
				testCellsByCol(data, [], rows, cols, offset, stepRow, assert);
			}

			// console.profileEnd('testForeachNoEmpty');

			assert.ok(true);
		});
	}
	function testForeachNoEmptyWithAttr(name, offset, stepRow) {
		QUnit.test(name, function (assert) {
			AscCommonExcel.g_StyleCache.xfs.list[0] = 1;

			// console.profile('testForeachNoEmpty');

			// let rowsTest = offset + 2 + stepRow;
			// let colsTest = 2;
			// let dataTest = [0,0,3,0,0,0,0,0];
			// let dataAttr = [null,null,null,null,1,null,null,null];
			// testCellsByCol(dataTest, dataAttr, rowsTest, colsTest, offset, stepRow, assert);

			let rows = offset + 2 + stepRow;//stepRow rows are needed to properly clean
			let cols = 2;
			let baseLen = (rows - stepRow) * cols;//last row for correct cleaning
			let base = [...Array(baseLen + 1).keys()].slice(1);
			let data = new Array(rows * cols);
			data.fill(0);
			let dataAttr = new Array(rows * cols);
			dataAttr.fill(null);
			let iterations = Math.pow(2, baseLen);
			for (let i = 0; i < iterations; ++i) {
				for (let j = 0; j < baseLen; ++j) {
					let bit = ((i >> j) % 2 !== 0);
					data[j] = bit ? base[j] : 0;
				}
				for (let k = 0; k < iterations; ++k) {
					for (let l = 0; l < baseLen; ++l) {
						let bit = ((k >> l) % 2 !== 0);
						dataAttr[l] = bit ? 1 : null;
					}
					testCellsByCol(data, dataAttr, rows, cols, offset, stepRow, assert);
				}
			}
			// console.profileEnd('testForeachNoEmpty');

			assert.ok(true);
		});
	}

	function testCellsByCol(data, dataAttr, rows, cols, offset, stepRow, assert) {
		let res = '';
		let testData = getTestDataFromArray(data, dataAttr, rows, cols, offset, stepRow);
		let r1 = offset;

		let sweepLine = new AscCommonExcel.SweepLineRowIterator(testData.cellsByCol, testData.xfsByCol, r1, 0, rows, cols);
		for (let i = r1; i < rows; i += stepRow) {
			sweepLine.setRow(i);
			while (sweepLine.nextCol()) {
				let dataStr = sweepLine.sheetMemory ? sweepLine.sheetMemory.getUint8(i, 0) : null;
				let xfStr = sweepLine.xf ? sweepLine.xf : null;
				if (dataStr || xfStr) {
					res += `${i}-${sweepLine.col}-${dataStr}-${xfStr};`;
				}
			}
		}
		//many asserts processes very slow
		if (res !== testData.expected) {
			assert.strictEqual(res, testData.expected, JSON.stringify(data) + "-" + JSON.stringify(dataAttr));
		}
		if (sweepLine.rowDataLen !== sweepLine.rowDataIndex) {
			assert.strictEqual(sweepLine.rowDataLen, sweepLine.rowDataIndex, "rowData");
		}
		if (sweepLine.toInsert.length !== sweepLine.toInsertIndex) {
			assert.strictEqual(sweepLine.toInsert.length, sweepLine.toInsertIndex, "toInsert");
		}
		if (sweepLine.toDelete.length !== sweepLine.toDeleteIndex) {
			assert.strictEqual(sweepLine.toDelete.length, sweepLine.toDeleteIndex, "toDelete");
		}
		if (sweepLine.events.length !== sweepLine.eventsIndex) {
			assert.strictEqual(sweepLine.events.length, sweepLine.eventsIndex, "events");
		}
	}

	function getTestDataFromArray(data, dataAttr, rows, cols, offset, stepRow) {
		//SheetMemory
		let cellsByCol = Array.from(Array(cols), () => {
			return new SheetMemory(2, rows);
		});
		for (let i = 0; i < data.length; ++i) {
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
		//xf
		let xfsByCol = Array.from(Array(cols), () => {
			return new AscCommonExcel.CAttrArray(null);
		});
		for (let i = 0; i < dataAttr.length; ++i) {
			if (null !== dataAttr[i]) {
				let row = Math.trunc(i / cols);
				let col = i % cols;
				let attrArray = xfsByCol[col];
				attrArray.set(row, dataAttr[i]);
			}
		}
		for (let i = 0; i < xfsByCol.length; ++i) {
			if (xfsByCol[i].size() === 0) {
				delete xfsByCol[i];
			}
		}
		//expected
		let expected = "";
		for (let i = offset * cols; i < data.length; ++i) {
			let row = Math.trunc(i / cols);
			let col = i % cols;
			if(row % stepRow === 0) {
				let sheetMemory = cellsByCol[col];
				let dataStr = null;
				if (sheetMemory && sheetMemory.hasIndex(row)) {
					dataStr = data[i];
				}
				let attrArray = xfsByCol[col];
				let xfStr = null;
				if (attrArray) {
					xfStr = attrArray.search(row);
				}
				if (dataStr || xfStr) {
					expected += `${row}-${col}-${dataStr}-${xfStr};`;
				}
			}
		}
		return {cellsByCol, xfsByCol, expected}
	}

	QUnit.module("CAttrArray");
	QUnit.test("Test: \"set\"", function (assert) {
		let attrArray = new AscCommonExcel.CAttrArray(null);
		attrArray.set(0, null, true);
		attrArray.set(1, 1, true);
		attrArray.set(2, null, true);
		attrArray.set(3, 1, true);
		attrArray.set(4, 1, true);
		attrArray.set(5, null, true);
		attrArray.set(6, 1, true);
		attrArray.set(7, 1, true);
		attrArray.set(8, 2, true);
		attrArray.set(9, 2, true);
		let expected = [null, 1, null, 1, 1, null, 1, 1, 2, 2];
		checkAttrArrayByArray(attrArray, expected, 0, assert, "");
		assert.ok(true);
	});
	QUnit.test("Test: \"setArea-cell\"", function (assert) {
		// let data = [null,1,1,1,null];
		// let attrArray = testAttrArraySetAreaCell(data);
		// checkAttrArrayByArray(attrArray, data, 0, assert, "");

		let baseLen = 5;
		let data = new Array(baseLen);
		data.fill(null);
		let iterations = Math.pow(2, baseLen);
		for (let i = 0; i < iterations; ++i) {
			for (let j = 0; j < baseLen; ++j) {
				let bit = ((i >> j) % 2 !== 0);
				data[j] = bit ? 1 : null;
			}
			let attrArray = testAttrArraySetAreaCell(data);
			checkAttrArrayByArray(attrArray, data, 0, assert, "");
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"setArea-range\"", function (assert) {
		// let data = [null, null, null, null, 1];
		// let attrArray = testAttrArraySetAreaCell(data);
		// checkAttrArrayByArray(attrArray, data, 0, assert, "");

		let baseLen = 6;
		let tail = 3;
		let data = new Array(baseLen + tail);
		for(let i = 0; i < baseLen; i++) {
			for(let j = 0; j < baseLen; j++) {
				for(let k = 0; k < baseLen; k++) {
					let attrArray = new AscCommonExcel.CAttrArray(null);
					data.fill(null);
					attrArray.setArea(j, j + 1, 1);
					data.fill(1, j, j + 2);
					attrArray.setArea(k, k + 2, 2);
					data.fill(2, k, k + 3);
					attrArray.setArea(i, i, 1);
					data.fill(1, i, i + 1);
					checkAttrArrayByArray(attrArray, data, tail, assert, i+"-"+j+"-"+k);
				}
			}
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"insertRange\"", function (assert) {
		let baseLen = 5;
		let iterations = Math.pow(2, baseLen);
		for (let k = 1; k < baseLen; ++k) {
			for (let i = 1; i < iterations; ++i) {
				let data = new Array(baseLen);
				data.fill(null);
				for (let j = 0; j < baseLen; ++j) {
					let bit = ((i >> j) % 2 !== 0);
					data[j] = bit ? 1 : null;
				}
				let attrArray = testAttrArraySetAreaCell(data);
				attrArray.insertRange(k, 1);
				data.splice(k, 0, null);
				checkAttrArrayByArray(attrArray, data, 0, assert, "");
			}
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"deleteRange\"", function (assert) {
		let baseLen = 5;
		let iterations = Math.pow(2, baseLen);
		for (let k = 0; k < baseLen; ++k) {
			for (let l = 0; l < baseLen - k - 1; ++l) {
				let data = new Array(baseLen);
				data.fill(null);
				for (let i = 0; i < iterations; ++i) {
					for (let j = 0; j < baseLen; ++j) {
						let bit = ((i >> j) % 2 !== 0);
						data[j] = bit ? 1 : null;
					}
					let attrArray = testAttrArraySetAreaCell(data);
					attrArray.deleteRange(k, l);
					data.splice(k, l);
					checkAttrArrayByArray(attrArray, data, 0, assert, "");
				}
			}
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"copyRange\"", function (assert) {
		let data = [null, 1, 2, 2, 3, 3, 3, null, 4, 4, 4, 4];
		let attrArray = testAttrArraySetAreaCell(data);
		for(let i = 0; i < data.length; i++) {
			for(let j = i; j < data.length; j++) {
				let count = j - i + 1;
				for(let k = 0; k < data.length - count; k++) {
					let attrArrayTemp = new AscCommonExcel.CAttrArray(null);
					attrArrayTemp.copyRange(attrArray, i, k, count);
					let dataExpected = new Array(data.length);
					dataExpected.fill(null);
					for (let l = i; l <= j; l++) {
						dataExpected[k + l - i] = data[l];
					}
					checkAttrArrayByArray(attrArrayTemp, dataExpected, 0, assert, i + "-" + j);
				}
			}
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"setAreaByRow\"", function (assert) {
		let data = [null, 1, 2, 2, 3, 3, 3, null, 4, 4, 4, 4];
		let dataExpected = data.concat();
		let attrArray = testAttrArraySetAreaCell(data);
		for (let i = 0; i < data.length; i += 3) {
			attrArray.setAreaByRow(i, i + 1, 2);
			dataExpected[i + 1] = dataExpected[i];
			dataExpected[i + 2] = dataExpected[i];
			checkAttrArrayByArray(attrArray, dataExpected, 0, assert, i);
		}
		assert.ok(true);
	});
	QUnit.test("Test: \"clear\"", function (assert) {
		let data = [null, 1, 2, 2, 3, 3, 3, null, 4, 4, 4, 4];
		let attrArray = testAttrArraySetAreaCell(data);
		attrArray.clear(0, 3);
		data.fill(null, 0, 3);
		checkAttrArrayByArray(attrArray, data, 0, assert, "clear1");
		attrArray.clear(6, 9);
		data.fill(null, 6, 9);
		checkAttrArrayByArray(attrArray, data, 0, assert, "clear2");
		assert.ok(true);
	});

	QUnit.module("CAttrArrayIterator");
	QUnit.test("Test: \"CAttrArrayIterator\"", function (assert) {
		let data = [null, 1, 2, 2, 3, 3, 3, null, 4, 4, 4, 4];
		let attrArray = testAttrArraySetAreaCell(data);

		for (let i = 0; i < data.length; i++) {
			for (let j = i; j < data.length; j++) {
				let colXfIterForward = new AscCommonExcel.CAttrArrayIterator(attrArray, i, j, false, false);
				let colXfIterForwardNoEmpty = new AscCommonExcel.CAttrArrayIterator(attrArray, i, j, true, false);
				let colXfIterBackward = new AscCommonExcel.CAttrArrayIterator(attrArray, i, j, false, true);
				let colXfIterBackwardNoEmpty = new AscCommonExcel.CAttrArrayIterator(attrArray, i, j, true, true);

				let minIndex = -1;
				for (let k = i; k >= 0 && data[k] === data[i]; k--) {
					minIndex = k;
				}
				let maxIndex = -1;
				for (let k = j; k < data.length && data[k] === data[j]; k++) {
					maxIndex = k;
				}
				let minIndexNoEmpty = -1;
				if (null !== data[i]) {
					minIndexNoEmpty = minIndex;
				} else {
					for (let k = i; k <= j; k++) {
						if (null !== data[k]) {
							minIndexNoEmpty = k;
							break;
						}
					}
				}
				let maxIndexNoEmpty = -1;
				if (null !== data[j]) {
					maxIndexNoEmpty = maxIndex;
				} else {
					for (let k = j; k >= i; k--) {
						if (null !== data[k]) {
							maxIndexNoEmpty = k;
							break;
						}
					}
				}

				if (colXfIterForward.getMinIndex() !== minIndex) {
					assert.deepEqual(colXfIterForward.getMinIndex(), minIndex, "colXfIterForward.getMinIndex " + i + "-" + j);
				}
				if (colXfIterForward.getMaxIndex() !== maxIndex) {
					assert.deepEqual(colXfIterForward.getMaxIndex(), maxIndex, "colXfIterForward.getMaxIndex " + i + "-" + j);
				}
				if (colXfIterForwardNoEmpty.getMinIndex() !== minIndexNoEmpty) {
					assert.deepEqual(colXfIterForwardNoEmpty.getMinIndex(), minIndexNoEmpty, "colXfIterForwardNoEmpty.getMinIndex " + i + "-" + j);
				}
				if (colXfIterForwardNoEmpty.getMaxIndex() !== maxIndexNoEmpty) {
					assert.deepEqual(colXfIterForwardNoEmpty.getMaxIndex(), maxIndexNoEmpty, "colXfIterForwardNoEmpty.getMaxIndex " + i + "-" + j);
				}
				if (colXfIterBackward.getMinIndex() !== minIndex) {
					assert.deepEqual(colXfIterBackward.getMinIndex(), minIndex, "colXfIterBackward.getMinIndex " + i + "-" + j);
				}
				if (colXfIterBackward.getMaxIndex() !== maxIndex) {
					assert.deepEqual(colXfIterBackward.getMaxIndex(), maxIndex, "colXfIterBackward.getMaxIndex " + i + "-" + j);
				}
				if (colXfIterBackwardNoEmpty.getMinIndex() !== minIndexNoEmpty) {
					assert.deepEqual(colXfIterBackwardNoEmpty.getMinIndex(), minIndexNoEmpty, "colXfIterBackwardNoEmpty.getMinIndex " + i + "-" + j);
				}
				if (colXfIterBackwardNoEmpty.getMaxIndex() !== maxIndexNoEmpty) {
					assert.deepEqual(colXfIterBackwardNoEmpty.getMaxIndex(), maxIndexNoEmpty, "colXfIterBackwardNoEmpty.getMaxIndex " + i + "-" + j);
				}

				let expected = data.slice(i, j - i + 1);
				let dataIterations = [];
				while (colXfIterForward.next()) {
					dataIterations.push({val: colXfIterForward.getCurVal(), from: colXfIterForward.getCurFrom(), to: colXfIterForward.getCurTo()});
				}
				let dataForward = new Array(expected.length);
				dataIterations.forEach(function(elem) {
					dataForward.fill(elem.val, elem.from - i, elem.to - i + 1);
				})
				if (!dataForward.every((val, idx) => val === expected[idx])) {
					assert.deepEqual(dataForward, expected, "colXfIterForward " + i + "-" + j);
				}

				let expectedIterationsFilter = dataIterations.filter(function(elem) {
					return null !== elem.val;
				});
				let dataIterationsNoEmpty = [];
				while (colXfIterForwardNoEmpty.nextNoEmpty()) {
					dataIterationsNoEmpty.push({val: colXfIterForwardNoEmpty.getCurVal(), from: colXfIterForwardNoEmpty.getCurFrom(), to: colXfIterForwardNoEmpty.getCurTo()});
				}
				if (!dataIterationsNoEmpty.every((val, idx) => val.val === expectedIterationsFilter[idx].val && val.from === expectedIterationsFilter[idx].from && val.to === expectedIterationsFilter[idx].to)) {
					assert.deepEqual(dataIterationsNoEmpty, expectedIterationsFilter, "colXfIterForwardNoEmpty " + i + "-" + j);
				}

				let expectedIterationsBackward = dataIterations.concat().reverse();
				let dataIterationsBackward = [];
				while (colXfIterBackward.prev()) {
					dataIterationsBackward.push({val: colXfIterBackward.getCurVal(), from: colXfIterBackward.getCurFrom(), to: colXfIterBackward.getCurTo()});
				}
				if (!dataIterationsBackward.every((val, idx) => val.val === expectedIterationsBackward[idx].val && val.from === expectedIterationsBackward[idx].from && val.to === expectedIterationsBackward[idx].to)) {
					assert.deepEqual(dataIterationsBackward, expectedIterationsBackward, "colXfIterBackward " + i + "-" + j);
				}

				let expectedIterationsBackwardFilter = expectedIterationsBackward.filter(function(elem) {
					return null !== elem.val;
				});
				let dataIterationsBackwardFilter = [];
				while (colXfIterBackwardNoEmpty.prevNoEmpty()) {
					dataIterationsBackwardFilter.push({val: colXfIterBackwardNoEmpty.getCurVal(), from: colXfIterBackwardNoEmpty.getCurFrom(), to: colXfIterBackwardNoEmpty.getCurTo()});
				}
				if (!dataIterationsBackwardFilter.every((val, idx) => val.val === expectedIterationsBackwardFilter[idx].val && val.from === expectedIterationsBackwardFilter[idx].from && val.to === expectedIterationsBackwardFilter[idx].to)) {
					assert.deepEqual(dataIterationsBackwardFilter, expectedIterationsBackwardFilter, "colXfIterBackwardNoEmpty " + i + "-" + j);
				}
			}
		}
		assert.ok(true);
	});

	function testAttrArraySetAreaCell(data) {
		let attrArray = new AscCommonExcel.CAttrArray(null);
		for (let i = 0, j = data.length - 1; i <= j; ++i, --j) {
			if (data[i]) {
				attrArray.setArea(i, i, data[i]);
			}
			if (data[j]) {
				attrArray.setArea(j, j, data[j]);
			}
		}
		return attrArray;
	}

	function checkAttrArrayByArray(attrArray, expected, expectedTail, assert, message) {
		let res = new Array(expected.length - expectedTail);
		let chunks = 0;
		let colXfIter = new AscCommonExcel.CAttrArrayIterator(attrArray, 0, expected.length - expectedTail - 1);
		while (colXfIter.next()) {
			chunks++;
			res.fill(colXfIter.getCurVal(), colXfIter.getCurFrom(), colXfIter.getCurTo() + 1);
		}
		let chunksExpected = 1;
		let prevElem = expected[0];
		for (let i = 1; i < expected.length - expectedTail; ++i) {
			if (expected[i] !== prevElem) {
				prevElem = expected[i];
				chunksExpected++;
			}
		}
		//many asserts processes very slow
		if (!res.every((val, idx) => val === expected[idx])) {
			assert.deepEqual(res, expected, "checkAttrArrayByArray data" + message);
		}
		if (chunks !== chunksExpected) {
			assert.equal(chunks, chunksExpected, "checkAttrArrayByArray chunks" + message + JSON.stringify(expected));
		}
		if (attrArray.data[attrArray.data.length - 1].endRow !== AscCommon.gc_nMaxRow0) {
			assert.equal(attrArray.data[attrArray.data.length - 1].endRow, AscCommon.gc_nMaxRow0, "checkAttrArrayByArray endRow" + message + JSON.stringify(expected));
		}
		return res;
	}
});
