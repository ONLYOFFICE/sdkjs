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

	function cmpArrays(array1, array2) {
		return array1.length === array2.length && array1.every((value, index) => value === array2[index]);
	}

	QUnit.module('BroadcastHelper');
	let sortsInsert = [
		[AscCommonExcel.BroadcastHelper.prototype.cmpInsertFrom,
			AscCommonExcel.BroadcastHelper.prototype.addInsertFrom],
		[AscCommonExcel.BroadcastHelper.prototype.cmpInsertTo, AscCommonExcel.BroadcastHelper.prototype.addInsertTo]
	];
	let sortsDelete = [
		[AscCommonExcel.BroadcastHelper.prototype.cmpDeleteFrom,
			AscCommonExcel.BroadcastHelper.prototype.addDeleteFrom],
		[AscCommonExcel.BroadcastHelper.prototype.cmpDeleteTo, AscCommonExcel.BroadcastHelper.prototype.addDeleteTo]
	];
	QUnit.test('BroadcastHelper simple', function (assert) {
		let helper = new AscCommonExcel.BroadcastHelper();
		helper.nextRow();
		assert.strictEqual(helper.from, -1, "from");
		assert.strictEqual(helper.to, -1, "to");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");

		let elems = [
			{index: 0, isActive: true, bbox: new Asc.Range(0, 0, 6, 0)},
			{index: 1, isActive: true, bbox: new Asc.Range(2, 0, 4, 0)}
		];
		insertElems(helper, elems);
		assert.strictEqual(helper.from, 0, "from");
		assert.strictEqual(helper.to, 6, "to");
		assert.strictEqual(helper.next(), true, "next0");
		assert.strictEqual(helper.curFrom, 0, "curFrom");
		assert.strictEqual(helper.curTo, 1, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		for (let i = 0; i < helper.curElemsLen; ++i) {
			let elem = helper.curElems[i];
			if (elem) {
				assert.strictEqual(elem, elems[i]);
			}
		}
		assert.strictEqual(helper.next(), true, "next1");
		assert.strictEqual(helper.curFrom, 2, "curFrom");
		assert.strictEqual(helper.curTo, 4, "curTo");
		assert.strictEqual(helper.curElemsCount, 2, "curElemsCount");
		for (let i = 0; i < helper.curElemsLen; ++i) {
			let elem = helper.curElems[i];
			if (elem) {
				assert.strictEqual(elem, elems[i]);
			}
		}
		assert.strictEqual(helper.next(), true, "next2");
		assert.strictEqual(helper.curFrom, 5, "curFrom");
		assert.strictEqual(helper.curTo, 6, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		for (let i = 0; i < helper.curElemsLen; ++i) {
			let elem = helper.curElems[i];
			if (elem) {
				assert.strictEqual(elem, elems[i]);
			}
		}
		assert.strictEqual(helper.next(), false, "next3");

		deleteElems(helper, elems);
		assert.strictEqual(helper.from, -1, "from");
		assert.strictEqual(helper.to, -1, "to");
		assert.strictEqual(helper.next(), false, "next4");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");
	});
	QUnit.test('BroadcastHelper sparse', function (assert) {
		let helper = new AscCommonExcel.BroadcastHelper();
		let elems = [
			{index: 0, isActive: true, bbox: new Asc.Range(0, 0, 2, 1)},
			{index: 1, isActive: true, bbox: new Asc.Range(4, 0, 6, 0)},
			{index: 2, isActive: true, bbox: new Asc.Range(9, 0, 10, 1)}
		];
		insertElems(helper, elems);
		assert.strictEqual(helper.from, 0, "from");
		assert.strictEqual(helper.to, 10, "to");
		assert.strictEqual(helper.next(), true, "next0");
		assert.strictEqual(helper.curFrom, 0, "curFrom");
		assert.strictEqual(helper.curTo, 2, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.curElems[0], elems[0]);
		assert.strictEqual(helper.next(), true, "next1");
		assert.strictEqual(helper.curFrom, 4, "curFrom");
		assert.strictEqual(helper.curTo, 6, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.curElems[1], elems[1]);
		assert.strictEqual(helper.next(), true, "next2");
		assert.strictEqual(helper.curFrom, 9, "curFrom");
		assert.strictEqual(helper.curTo, 10, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.curElems[2], elems[2]);
		assert.strictEqual(helper.next(), false, "next3");
	});
	QUnit.test('BroadcastHelper Insert Delete', function (assert) {
		let helper = new AscCommonExcel.BroadcastHelper();

		let elems1 = [
			{index: 0, isActive: true, bbox: new Asc.Range(1, 0, 5, 2)},
			{index: 1, isActive: true, bbox: new Asc.Range(1, 0, 4, 2)}
		];
		insertElems(helper, elems1);
		assert.strictEqual(helper.from, 1, "from");
		assert.strictEqual(helper.to, 5, "to");
		assert.strictEqual(helper.next(), true, "next1");
		assert.strictEqual(helper.curFrom, 1, "curFrom");
		assert.strictEqual(helper.curTo, 4, "curTo");
		assert.strictEqual(helper.curElemsCount, 2, "curElemsCount");
		assert.strictEqual(helper.next(), true, "next2");
		assert.strictEqual(helper.curFrom, 5, "curFrom");
		assert.strictEqual(helper.curTo, 5, "curTo");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.next(), false, "next3");

		let elems2 = [
			{index: 2, isActive: true, bbox: new Asc.Range(1, 1, 3, 2)},
			{index: 3, isActive: true, bbox: new Asc.Range(1, 1, 2, 2)}
		];
		insertElems(helper, elems2);
		assert.strictEqual(helper.from, 1, "from");
		assert.strictEqual(helper.to, 5, "to");
		assert.strictEqual(helper.next(), true, "next4");
		assert.strictEqual(helper.curFrom, 1, "from");
		assert.strictEqual(helper.curTo, 2, "to");
		assert.strictEqual(helper.curElemsCount, 4, "curElemsCount");
		assert.strictEqual(helper.next(), true, "next5");
		assert.strictEqual(helper.curFrom, 3, "from");
		assert.strictEqual(helper.curTo, 3, "to");
		assert.strictEqual(helper.curElemsCount, 3, "curElemsCount");
		assert.strictEqual(helper.next(), true, "next6");
		assert.strictEqual(helper.curFrom, 4, "from");
		assert.strictEqual(helper.curTo, 4, "to");
		assert.strictEqual(helper.curElemsCount, 2, "curElemsCount");
		assert.strictEqual(helper.next(), true, "next7");
		assert.strictEqual(helper.curFrom, 5, "from");
		assert.strictEqual(helper.curTo, 5, "to");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.next(), false, "next8");

		let elems = elems1.concat(elems2);
		deleteElems(helper, elems);
		assert.strictEqual(helper.from, -1, "from");
		assert.strictEqual(helper.to, -1, "to");
		assert.strictEqual(helper.next(), false, "next9");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");
	});
	QUnit.test('BroadcastHelper curElems', function (assert) {
		let helper = new AscCommonExcel.BroadcastHelper();

		let elems = [
			{index: 0, isActive: true, bbox: new Asc.Range(1, 0, 5, 2)},
			{index: 1, isActive: true, bbox: new Asc.Range(1, 0, 4, 2)},
			{index: 3, isActive: true, bbox: new Asc.Range(1, 0, 3, 2)},
			{index: 4, isActive: true, bbox: new Asc.Range(1, 0, 2, 2)}
		];
		insertElems(helper, elems);
		assert.strictEqual(helper.next(), true, "next1");
		assert.strictEqual(helper.curElemsCount, 4, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		for (let i = 1; i < 3; ++i) {
			let elem = helper.curElems[i];
			elem.isActive = false;
			helper.addDeleteFrom(elem);
			helper.addDeleteToUnsorted(elem);
			// helper.addDeleteTo(elem);
		}
		helper.finishDelete();
		assert.strictEqual(helper.from, 1, "from");
		assert.strictEqual(helper.to, 5, "to");
		assert.strictEqual(helper.next(), true, "next2");
		assert.strictEqual(helper.curElemsCount, 2, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 2, "curElemsCount");
		assert.strictEqual(helper.curElems[0], elems[0], "curElems[0]");
		assert.strictEqual(helper.curElems[1], elems[3], "curElems[1]");

		deleteElems(helper, elems);
		assert.strictEqual(helper.from, -1, "from");
		assert.strictEqual(helper.to, -1, "to");
		assert.strictEqual(helper.next(), false, "next3");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 0, "curElemsLen");
	});
	QUnit.test('BroadcastHelper curElems next', function (assert) {
		let helper = new AscCommonExcel.BroadcastHelper();

		let elems = [
			{index: 0, isActive: true, bbox: new Asc.Range(1, 0, 2, 1)},
			{index: 1, isActive: true, bbox: new Asc.Range(1, 0, 3, 1)},
			{index: 3, isActive: true, bbox: new Asc.Range(1, 0, 4, 1)},
			{index: 4, isActive: true, bbox: new Asc.Range(1, 0, 5, 1)}
		];
		insertElems(helper, elems);
		assert.strictEqual(helper.next(), true, "next1");
		assert.strictEqual(helper.curElemsCount, 4, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		assert.strictEqual(helper.curElems[0], elems[0], "curElems[0]");
		assert.strictEqual(helper.curElems[1], elems[1], "curElems[1]");
		assert.strictEqual(helper.curElems[2], elems[2], "curElems[2]");
		assert.strictEqual(helper.curElems[3], elems[3], "curElems[3]");

		assert.strictEqual(helper.next(), true, "next2");
		assert.strictEqual(helper.curElemsCount, 3, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		assert.strictEqual(helper.curElems[0], undefined, "curElems[0]");
		assert.strictEqual(helper.curElems[1], elems[1], "curElems[1]");
		assert.strictEqual(helper.curElems[2], elems[2], "curElems[2]");
		assert.strictEqual(helper.curElems[3], elems[3], "curElems[3]");

		assert.strictEqual(helper.next(), true, "next3");
		assert.strictEqual(helper.curElemsCount, 2, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		assert.strictEqual(helper.curElems[0], undefined, "curElems[0]");
		assert.strictEqual(helper.curElems[1], undefined, "curElems[1]");
		assert.strictEqual(helper.curElems[2], elems[2], "curElems[2]");
		assert.strictEqual(helper.curElems[3], elems[3], "curElems[3]");

		assert.strictEqual(helper.next(), true, "next4");
		assert.strictEqual(helper.curElemsCount, 1, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		assert.strictEqual(helper.curElems[0], undefined, "curElems[0]");
		assert.strictEqual(helper.curElems[1], undefined, "curElems[1]");
		assert.strictEqual(helper.curElems[2], undefined, "curElems[2]");
		assert.strictEqual(helper.curElems[3], elems[3], "curElems[3]");

		assert.strictEqual(helper.next(), false, "next5");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 4, "curElemsLen");
		assert.strictEqual(helper.curElems[0], undefined, "curElems[0]");
		assert.strictEqual(helper.curElems[1], undefined, "curElems[1]");
		assert.strictEqual(helper.curElems[2], undefined, "curElems[2]");
		assert.strictEqual(helper.curElems[3], undefined, "curElems[3]");

		deleteElems(helper, elems);
		assert.strictEqual(helper.from, -1, "from");
		assert.strictEqual(helper.to, -1, "to");
		assert.strictEqual(helper.next(), false, "next6");
		assert.strictEqual(helper.curElemsCount, 0, "curElemsCount");
		assert.strictEqual(helper.curElemsLen, 0, "curElemsLen");
	});
	QUnit.test('BroadcastHelper generated', function (assert) {
		assert.ok(true);
		let helper = new AscCommonExcel.BroadcastHelper();
		let rows = 3;
		let cols = 3;

		// let elems = [{"index":0,"id":1,"isActive":true,"bbox":{"c1":0,"r1":0,"c2":0,"r2":0,"refType1":3,"refType2":3}},{"index":1,"id":2,"isActive":true,"bbox":{"c1":0,"r1":0,"c2":0,"r2":3,"refType1":3,"refType2":3}}];
		// testElems(assert, elems, rows, cols);
		// return;

		for (let a_r1 = 0; a_r1 < rows; ++a_r1) {
			for (let a_r2 = a_r1; a_r2 < rows; ++a_r2) {
				for (let a_c1 = 0; a_c1 < cols; ++a_c1) {
					for (let a_c2 = a_c1; a_c2 < cols; ++a_c2) {
						for (let b_r1 = 0; b_r1 < rows; ++b_r1) {
							for (let b_r2 = b_r1; b_r2 < rows; ++b_r2) {
								for (let b_c1 = 0; b_c1 < cols; ++b_c1) {
									for (let b_c2 = b_c1; b_c2 < cols; ++b_c2) {
										let elems = [
											{
												index: 0, id: 1, isActive: true,
												bbox: new Asc.Range(a_c1, a_r1, a_c2, a_r2)
											},
											{
												index: 1, id: 2, isActive: true,
												bbox: new Asc.Range(b_c1, b_r1, b_c2, b_r2)
											}
										];
										testElems(assert, elems, rows, cols);
									}
								}
							}
						}
					}
				}
			}
		}
		// for (let a_r1 = 0; a_r1 < rows; ++a_r1) {
		//     for (let a_r2 = a_r1; a_r2 < rows; ++a_r2) {
		//         for (let a_c1 = 0; a_c1 < cols; ++a_c1) {
		//             for (let a_c2 = a_c1; a_c2 < cols; ++a_c2) {
		//                 for (let b_r1 = 0; b_r1 < rows; ++b_r1) {
		//                     for (let b_r2 = b_r1; b_r2 < rows; ++b_r2) {
		//                         for (let b_c1 = 0; b_c1 < cols; ++b_c1) {
		//                             for (let b_c2 = b_c1; b_c2 < cols; ++b_c2) {
		//                                 for (let c_r1 = 0; c_r1 < rows; ++c_r1) {
		//                                     for (let c_r2 = c_r1; c_r2 < rows; ++c_r2) {
		//                                         for (let c_c1 = 0; c_c1 < cols; ++c_c1) {
		//                                             for (let c_c2 = c_c1; c_c2 < cols; ++c_c2) {
		//                                                 let elems = [
		//                                                     {index: 0, id: 1, isActive: true, bbox: new Asc.Range(a_c1, a_r1, a_c2, a_r2)},
		//                                                     {index: 1, id: 2, isActive: true, bbox: new Asc.Range(b_c1, b_r1, b_c2, b_r2)},
		//                                                     {index: 2, id: 4, isActive: true, bbox: new Asc.Range(c_c1, c_r1, c_c2, c_r2)}
		//                                                 ];
		//                                                 testElems(assert, elems, rows, cols);
		//                                             }
		//                                         }
		//                                     }
		//                                 }
		//                             }
		//                         }
		//                     }
		//                 }
		//             }
		//         }
		//     }
		// }
	});

	function insertElems(helper, elems) {
		let elemsTemp = elems.concat();
		for (let i = 0; i < sortsInsert.length; ++i) {
			elemsTemp.sort(sortsInsert[i][0]);
			for (let j = 0; j < elemsTemp.length; ++j) {
				sortsInsert[i][1].call(helper, elemsTemp[j]);
			}
		}
		helper.finishInsert();
	}

	function deleteElems(helper, elems) {
		let elemsTemp = elems.concat();
		for (let i = 0; i < sortsDelete.length; ++i) {
			elemsTemp.sort(sortsDelete[i][0]);
			for (let j = 0; j < elemsTemp.length; ++j) {
				if (elemsTemp[j].isActive) {
					sortsDelete[i][1].call(helper, elemsTemp[j]);
				}
			}
		}
		helper.finishDelete();
	}

	function testElems(assert, elems, rows, cols) {
		let namePrefix = JSON.stringify(elems);
		let helper = new AscCommonExcel.BroadcastHelper();
		let expected = getTestDataFromArray(elems, rows, cols);
		let insertFromSorted = elems.concat();
		insertFromSorted.sort(helper.cmpInsertFrom);
		let insertFromSortedIndex = 0;
		let insertToSorted = elems.concat();
		insertToSorted.sort(helper.cmpInsertTo);
		let insertToSortedIndex = 0;
		let deleteFromSorted = elems.concat();
		deleteFromSorted.sort(helper.cmpDeleteFrom);
		let deleteFromSortedIndex = 0;
		let deleteToSorted = elems.concat();
		deleteToSorted.sort(helper.cmpDeleteTo);
		let deleteToSortedIndex = 0;
		let res = new Array(rows * cols);
		res.fill(0);

		let curYOld = 0;
		var fillRes = function (row, helper) {
			helper.nextRow();
			while (helper.next()) {
				let sum = 0;
				if (0 === helper.curElemsLen) {
					assert.notStrictEqual(helper.curElemsLen, 0, namePrefix + "_curElemsLen");
				}
				for (let i = 0; i < helper.curElemsLen; ++i) {
					if (helper.curElems[i]) {
						sum += helper.curElems[i].id;
					}
				}
				for (let i = helper.curFrom; i <= helper.curTo; ++i) {
					res[row * cols + i] = sum;
				}
			}
		}
		while (deleteFromSortedIndex < deleteFromSorted.length) {
			let curY = 0;
			if (insertFromSortedIndex < insertFromSorted.length) {
				curY = Math.min(insertFromSorted[insertFromSortedIndex].bbox.r1,
					deleteFromSorted[deleteFromSortedIndex].bbox.r2);
			} else {
				curY = deleteFromSorted[deleteFromSortedIndex].bbox.r2;
			}
			for (let i = curYOld + 1; i < curY; ++i) {
				fillRes(i, helper);
			}

			insertFromSortedIndex = helper.addInsertByRow(curY, curY, insertFromSorted, insertFromSortedIndex, helper.addInsertFrom);
			insertToSortedIndex = helper.addInsertByRow(curY, curY, insertToSorted, insertToSortedIndex, helper.addInsertTo);
			helper.finishInsert();

			fillRes(curY, helper);

			deleteFromSortedIndex =
				helper.addDeleteByRow(curY, deleteFromSorted, deleteFromSortedIndex, helper.addDeleteFrom);
			deleteToSortedIndex = helper.addDeleteByRow(curY, deleteToSorted, deleteToSortedIndex, helper.addDeleteTo);
			helper.finishDelete();
			curYOld = curY;
		}
		if (!cmpArrays(res, expected)) {
			assert.deepEqual(res, expected, namePrefix + "_expected");
			//throw 1;
		}
		//many asserts processes very slow
		if (helper.from !== -1) {
			assert.strictEqual(helper.from, -1, namePrefix + "_from");
		}
		if (helper.to !== -1) {
			assert.strictEqual(helper.to, -1, namePrefix + "_to");
		}
		if (helper.next() !== false) {
			assert.strictEqual(true, false, namePrefix + "_next");
		}
		if (helper.curElemsCount !== 0) {
			assert.strictEqual(helper.curElemsCount, 0, namePrefix + "_curElemsCount");
		}
	}

	function getTestDataFromArray(elems, rows, cols) {
		let expected = new Array(rows * cols);
		expected.fill(0);
		for (let i = 0; i < elems.length; ++i) {
			let elem = elems[i];
			for (let r = elem.bbox.r1; r <= elem.bbox.r2; ++r) {
				for (let c = elem.bbox.c1; c <= elem.bbox.c2; ++c) {
					expected[r * cols + c] += elem.id;
				}
			}
		}
		return expected;
	}

	QUnit.module('DependencyGraph');
	let api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.initCollaborativeEditing({});
	let wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);
	let ws = new AscCommonExcel.Worksheet(wb, 0);
	let sheetId = ws.getId();
	QUnit.test('DependencyGraph _broadcastCellsByCells', function (assert) {
		assert.ok(true);
		let listeningBbox = ["B1", "C1", "A2", "D2", "E2"];
		let changedBbox = ["A1", "B1", "C1", "D1", "D2", "A3"];
		//let expected = ["B1", "C1", "D2"];

		testListeningNames(assert, listeningBbox, changedBbox);
	});
	QUnit.test('DependencyGraph _broadcastRangesByCells', function (assert) {
		assert.ok(true);
		let listeningBbox = ["B2:C2", "B2:D2", "C2:D2", "E2:E3", "G2:G3"];
		let changedBbox = ["A1", "B1", "A2", "C2", "H2", "E3", "F3"];
		//let expected = ["B2:C2", "B2:D2", "C2:D2", "E2:E3"];

		testListeningNames(assert, listeningBbox, changedBbox);
	});
	QUnit.test('DependencyGraph _broadcastCellsByRanges', function (assert) {
		assert.ok(true);
		let listeningBbox = ["B1", "C1", "A2", "D2", "E2"];
		let changedBbox = ["C1:D2", "A2:B3", "E2:F3"];
		//let expected = ["C1", "A2", "D2", "E2"];

		testListeningNames(assert, listeningBbox, changedBbox);
	});
	QUnit.test('DependencyGraph _broadcastRangesByRanges', function (assert) {
		assert.ok(true);
		let listeningBbox = ["B2:C2", "B2:D2", "C2:D2", "E2:E3", "G2:G3"];
		let changedBbox = ["A1:B3", "D3:H4"];
		//let expected = ["B2:C2", "B2:D2", "E2:E3", "G2:G3"];

		testListeningNames(assert, listeningBbox, changedBbox);
	});
	QUnit.test('DependencyGraph generated', function (assert) {
		assert.ok(true);
		let rows = 3;
		let cols = 3;

		// let listeningBbox = ["A1:A2","A3:B3"]
		// let changedBbox = ["A2:B2","A2:C2"]
		// testListeningNames(assert, listeningBbox, changedBbox);
		// return;

		for (let a_r1 = 0; a_r1 < rows; ++a_r1) {
			for (let a_r2 = a_r1; a_r2 < rows; ++a_r2) {
				for (let a_c1 = 0; a_c1 < cols; ++a_c1) {
					for (let a_c2 = a_c1; a_c2 < cols; ++a_c2) {
						for (let b_r1 = a_r1; b_r1 < rows; ++b_r1) {
							for (let b_r2 = b_r1; b_r2 < rows; ++b_r2) {
								for (let b_c1 = a_c1; b_c1 < cols; ++b_c1) {
									for (let b_c2 = b_c1; b_c2 < cols; ++b_c2) {
										for (let c_r1 = 0; c_r1 < rows; ++c_r1) {
											for (let c_r2 = c_r1; c_r2 < rows; ++c_r2) {
												for (let c_c1 = 0; c_c1 < cols; ++c_c1) {
													for (let c_c2 = c_c1; c_c2 < cols; ++c_c2) {
														for (let d_r1 = c_r1; d_r1 < rows; ++d_r1) {
															for (let d_r2 = d_r1; d_r2 < rows; ++d_r2) {
																for (let d_c1 = c_c1; d_c1 < cols; ++d_c1) {
																	for (let d_c2 = d_c1; d_c2 < cols; ++d_c2) {
																		if(!(a_r1 === b_r1 && a_c1 === b_c1 && a_r2 === b_r2 && a_c2 === b_c2) &&
																			!(c_r1 === d_r1 && c_c1 === d_c1 && c_r2 === d_r2 && c_c2 === d_c2))
																		{
																			let listeningBbox = [
																				new Asc.Range(a_c1, a_r1, a_c2, a_r2),
																				new Asc.Range(b_c1, b_r1, b_c2, b_r2)
																			];
																			let changedBbox = [
																				new Asc.Range(c_c1, c_r1, c_c2, c_r2),
																				new Asc.Range(d_c1, d_r1, d_c2, d_r2)
																			];
																			testListening(assert, listeningBbox,changedBbox);
																		}
																	}
																}
															}
														}
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
		// for (let a_r1 = 0; a_r1 < rows; ++a_r1) {
		// 	for (let a_r2 = a_r1; a_r2 < rows; ++a_r2) {
		// 		for (let a_c1 = 0; a_c1 < cols; ++a_c1) {
		// 			for (let a_c2 = a_c1; a_c2 < cols; ++a_c2) {
		// 				for (let b_r1 = a_r1; b_r1 < rows; ++b_r1) {
		// 					for (let b_r2 = b_r1; b_r2 < rows; ++b_r2) {
		// 						for (let b_c1 = a_c1; b_c1 < cols; ++b_c1) {
		// 							for (let b_c2 = b_c1; b_c2 < cols; ++b_c2) {
		// 								for (let c_r1 = 0; c_r1 < rows; ++c_r1) {
		// 									for (let c_r2 = c_r1; c_r2 < rows; ++c_r2) {
		// 										for (let c_c1 = 0; c_c1 < cols; ++c_c1) {
		// 											for (let c_c2 = c_c1; c_c2 < cols; ++c_c2) {
		// 												for (let d_r1 = c_r1; d_r1 < rows; ++d_r1) {
		// 													for (let d_r2 = d_r1; d_r2 < rows; ++d_r2) {
		// 														for (let d_c1 = c_c1; d_c1 < cols; ++d_c1) {
		// 															for (let d_c2 = d_c1; d_c2 < cols; ++d_c2) {
		// 																for (let e_r1 = c_r1; e_r1 < rows; ++e_r1) {
		// 																	for (let e_r2 = e_r1; e_r2 < rows; ++e_r2) {
		// 																		for (let e_c1 = c_c1; e_c1 < cols; ++e_c1) {
		// 																			for (let e_c2 = e_c1; e_c2 < cols; ++e_c2) {
		// 																				if(!(a_r1 === b_r1 && a_c1 === b_c1 && a_r2 === b_r2 && a_c2 === b_c2) &&
		// 																					!(a_r1 === c_r1 && a_c1 === c_c1 && a_r2 === c_r2 && a_c2 === c_c2) &&
		// 																					!(c_r1 === b_r1 && c_c1 === b_c1 && c_r2 === b_r2 && c_c2 === b_c2) &&
		// 																					!(e_r1 === d_r1 && e_c1 === d_c1 && e_r2 === d_r2 && e_c2 === d_c2))
		// 																				{
		// 																					let listeningBbox = [
		// 																						new Asc.Range(a_c1, a_r1, a_c2, a_r2),
		// 																						new Asc.Range(b_c1, b_r1, b_c2, b_r2),
		// 																						new Asc.Range(c_c1, c_r1, c_c2, c_r2),
		// 																					];
		// 																					let changedBbox = [
		// 																						new Asc.Range(d_c1, d_r1, d_c2, d_r2),
		// 																						new Asc.Range(e_c1, e_r1, e_c2, e_r2)
		// 																					];
		// 																					testListening(assert, listeningBbox,changedBbox);
		// 																				}
		// 																			}
		// 																		}
		// 																	}
		// 																}
		// 															}
		// 														}
		// 													}
		// 												}
		// 											}
		// 										}
		// 									}
		// 								}
		// 							}
		// 						}
		// 					}
		// 				}
		// 			}
		// 		}
		// 	}
		// }
	});

	function getExpected(listeningBbox, changedBbox) {
		let indexes = new Array(listeningBbox.length);
		for (let i = 0; i < listeningBbox.length; ++i) {
			for (let j = 0; j < changedBbox.length; ++j) {
				if (listeningBbox[i].isIntersect(changedBbox[j])) {
					indexes[i] = 1;
				}
			}
		}
		return listeningBbox.filter(function (elem, index) {
			return 1 === indexes[index];
		}).map(function(elem){
			return elem.getName();
		});
	}
	function getNames(arr) {
		return arr.map(function(elem){
			return elem.getName();
		});
	}

	function testListeningNames(assert, listeningBbox, changedBbox) {
		listeningBbox = listeningBbox.map(function (elem) {
			return AscCommonExcel.g_oRangeCache.getAscRange(elem);
		});
		changedBbox = changedBbox.map(function (elem) {
			return AscCommonExcel.g_oRangeCache.getAscRange(elem);
		});
		testListening(assert, listeningBbox, changedBbox);
	}

	function testListening(assert, listeningBbox, changedBbox) {
		let res = new Set();
		let graph = new AscCommonExcel.DependencyGraph(wb);
		let listeners = [];
		for (let i = 0; i < listeningBbox.length; ++i) {
			let bbox = listeningBbox[i];
			let listener = {
				bbox: bbox,
				getListenerId: function () {
					return i;
				},
				notify: function (data) {
					res.add(bbox.getName());
				}
			};
			listeners.push(listener);
			graph.startListeningRange(sheetId, bbox, listener);
		}
		for (let i = 0; i < changedBbox.length; ++i) {
			let bbox = changedBbox[i];
			if (bbox.isOneCell()) {
				graph.addToChangedCell2(sheetId, bbox.r1, bbox.c1);
			} else {
				graph.addToChangedRange2(sheetId, bbox);
			}
		}
		graph.calcTree();
		let expected = getExpected(listeningBbox, changedBbox);
		let expectedSorted = expected.sort();
		let resSorted = Array.from(res.keys()).sort();
		if (!cmpArrays(resSorted, expected)) {
			assert.deepEqual(resSorted, expected, "testListening");
			var listeningBboxNames = JSON.stringify(getNames(listeningBbox));
			var changedBboxNames = JSON.stringify(getNames(changedBbox));
			//throw 1;
		}

		for (let i = 0; i < listeningBbox.length; ++i) {
			let bbox = listeningBbox[i];
			graph.endListeningRange(sheetId, bbox, listeners[i]);
		}
		res.clear();
		graph.calcTree();
		resSorted = Array.from(res.keys()).sort();
		if (!cmpArrays(resSorted, [])) {
			assert.deepEqual(resSorted, [], "testListening endListeningRange");
		}
	}

});
