/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

"use strict";

(function (window)
{
	const editor = Asc.editor;
	const {
		wbView,
		executeTestWithCatchEvent,
		getFragments,
		getSelectionCellEditor,
		moveToStartCellEditor,
		moveToEndCellEditor,
		moveToCell,
		selectToCell,
		checkOpenCellEditor,
		onKeyDown,
		closeCellEditor,
		enterTextWithoutClose,
		setCheckOpenCellEditor,
		enterText,
		cellEditor,
		getCellText,
		wsView,
		ws,
		moveAndEnterText,
		createTest,
		moveAndGetCellText,
		goToSheet,
		createWorksheet,
		removeCurrentWorksheet,
		cleanSelection,
		cleanActiveCell,
		checkRange,
		openCellEditor,
		checkActiveCell,
		selectAll,
		cleanAll,
		setCellFormat,
		xfs,
		undo,
		createEvent,
		selectAllCell,
		cellPosition,
		getCellEditMode,
		testPreventDefaultAndStopPropagation,
		addHyperlink,
		createMathInShape,
		moveCursorRight,
		moveCursorLeft,
		moveCursorUp,
		moveCursorDown,
		checkMoveContentShapeHelper,
		contentPosition,
		selectedContent,
		checkSelectContentShapeHelper,
		round,
		moveShapeHelper,
		createTable,
		createSlicer,
		moveToShapeParagraph,
		addProperty,
		textPr,
		moveToParagraph,
		selectedObjects,
		graphicController,
		createShape,
		selectOnlyObjects,
		createShapeWithContent,
		checkTextAfterKeyDownHelperHelloWorld,
		getParagraphText,
		checkDirectTextPrAfterKeyDown,
		checkDirectParaPrAfterKeyDown,
		startCellEditorTest,
		startTableTest,
		startGraphicTest,
		oTableTypes,
		oCellEditorTypes,
		oGraphicTypes,
		checkRemoveObject,
		checkTextAfterKeyDownHelperEmpty,
		createGroup,
		createChart
	} = window.AscTestShortcut;
	const pxToMm = AscCommon.g_dKoef_pix_to_mm;

	$(
		function ()
		{
			QUnit.module('test worksheet shortcuts');
			function createTwoPivotTables()
			{
				cleanAll();
				moveAndEnterText('ad', 0, 0);
				moveAndEnterText('1', 1, 0);
				moveAndEnterText('2', 2, 0);
				editor.asc_insertPivotExistingWorksheet("Sheet1!$A$1:$A$3", "Sheet1!$C$1");
				const oPivotTable1 = ws().getPivotTable(2, 1);
				oPivotTable1.asc_addField(editor, 0);

				moveAndEnterText('ap', 0, 1);
				moveAndEnterText('2', 1, 1);
				moveAndEnterText('3', 2, 1);
				editor.asc_insertPivotExistingWorksheet("Sheet1!$B$1:$B$3", "Sheet1!$D$1");
				const oPivotTable2 = ws().getPivotTable(3, 1);
				oPivotTable2.asc_addField(editor, 0);
			}
			QUnit.test('Test refresh all connections', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					createTwoPivotTables();

					moveAndEnterText('4', 2, 0);
					moveAndEnterText('4', 2, 1);

					moveToCell(1, 2);
					onKeyDown(oEvent);
					equal(getCellText(), '5');

					moveToCell(1, 3);
					equal(getCellText(), '6');
				}, oTableTypes.refreshAllConnections);

			});

			QUnit.test('Test refresh selected connections', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					createTwoPivotTables();

					moveAndEnterText('4', 2, 0);
					moveAndEnterText('4', 2, 1);

					moveToCell(1, 2);
					onKeyDown(oEvent);
					equal(getCellText(), '5');

					moveToCell(1, 3);
					equal(getCellText(), '5');
				}, oTableTypes.refreshSelectedConnections);
			});


			QUnit.test('Test change format table info', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					createTable();
					moveToCell(4, 0);
					onKeyDown(oEvent);
					moveToCell(5, 0);
					equal(getCellText(), '10');
				}, oTableTypes.changeFormatTableInfo);
			});
			// todo
			// QUnit.test.todo('Test calculate all', (oAssert) =>
			// {
			// 	const {deep, equal} = createTest(oAssert);
			// 	startTableTest((oEvent) =>
			// 	{
			//
			// 	}, oTableTypes.calculateAll);
			//
			// });
			//
			//
			// QUnit.test.todo('Test calculate active sheet', (oAssert) =>
			// {
			// 	const {deep, equal} = createTest(oAssert);
			// 	startTableTest((oEvent) =>
			// 	{
			//
			// 	}, oTableTypes.calculateActiveSheet);
			//
			// });
			//
			// QUnit.test.todo('Test calculate only changed', (oAssert) =>
			// {
			// 	const {deep, equal} = createTest(oAssert);
			// 	startTableTest((oEvent) =>
			// 	{
			//
			// 	}, oTableTypes.calculateOnlyChanged);
			//
			// });

			QUnit.test('Test focus on cell editor', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					equal(getCellEditMode(), true);
				}, oTableTypes.focusOnCellEditor);
			});

			QUnit.test('Test add date', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					const oDate = new Asc.cDate();
					equal(getCellText(), oDate.getDateString(editor), 'Check insert current date');
				}, oTableTypes.addDate);


			});
			QUnit.test('Test add time', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					const oDate = new Asc.cDate();
					onKeyDown(oEvent);
					equal(getCellText(), oDate.getTimeString(editor).split(' ').join(':00 '), 'Check insert current time');
				}, oTableTypes.addTime);


			});


			QUnit.test('Test remove active cell text', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(0, 0);
					enterText('hello World');
					onKeyDown(oEvent);
					equal(getCellText(), '', 'Check remove active cell');
				}, oTableTypes.removeActiveCell);
			});


			QUnit.test('Test empty range', (oAssert) =>
			{

				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveAndEnterText('Hello World', 0, 0);
					moveAndEnterText('Hello World', 1, 1);
					moveAndEnterText('Hello World', 2, 2);
					moveToCell(0, 0);
					selectToCell(5, 5);
					onKeyDown(oEvent);
					const arrSteps = [];
					closeCellEditor();
					arrSteps.push(moveAndGetCellText(0, 0));
					arrSteps.push(moveAndGetCellText(1, 1));
					arrSteps.push(moveAndGetCellText(2, 2));
					deep(arrSteps, ['', '', ''], 'Check empty shortcut');
				}, oTableTypes.emptyRange);

			});


			QUnit.test('Test move active cell to left', (oAssert) =>
			{

				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 1);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0), 'Check move left active cell');
				}, oTableTypes.moveActiveCellToLeft);

			});

			QUnit.test('Test move active cell to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 1), 'Check move right active cell');
				}, oTableTypes.moveActiveCellToRight);

			});

			QUnit.test('Test move active cell to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(1, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(2, 0), 'Check move down active cell');
				}, oTableTypes.moveActiveCellToDown);
			});

			QUnit.test('Test move active cell to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(1, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0), 'Check move up active cell');
				}, oTableTypes.moveActiveCellToUp);

			});

			QUnit.test('Test reset', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					editor.asc_startAddShape('rect');
					onKeyDown(oEvent);
					equal(graphicController().checkEndAddShape(), false);
					moveToCell(0, 0);
					editor.asc_SelectionCut();
					onKeyDown(oEvent);
					equal(wsView().copyCutRange, null);

					moveToCell(0, 0);
					editor.asc_Copy();
					onKeyDown(oEvent);
					equal(wsView().copyCutRange, null);
				}, oTableTypes.reset);

			});

			QUnit.test('Test disable num lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					// run test for macOS and Windows
					moveToCell(0, 0);
					AscCommon.AscBrowser.isOpera = true;
					testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert);

					AscCommon.AscBrowser.isOpera = false;
					testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert, true);
				}, oTableTypes.disableNumLock);


			});

			QUnit.test('Test disable scroll lock', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					// run test for macOS and Windows
					moveToCell(0, 0);
					AscCommon.AscBrowser.isOpera = true;
					testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert);

					AscCommon.AscBrowser.isOpera = false;
					testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert, true);
				}, oTableTypes.disableScrollLock);

			});

			QUnit.test('Test select column', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 1048575, 0, 0), 'Check move up');
				}, oTableTypes.selectColumn);

			});

			QUnit.test('Test select row', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 0, 0, 16383), 'Check move up');
				}, oTableTypes.selectRow);
			});

			QUnit.test('Test select sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 1048575, 0, 16383), 'Check move up');
				}, oTableTypes.selectSheet);
			});

			QUnit.test('Test add separator', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					equal(getCellText(), '.');
				}, oTableTypes.addSeparator);
			});


			QUnit.test('Test go to previous sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					createWorksheet();
					onKeyDown(oEvent);
					equal(wbView().wsActive, 0, 'Check got to previous worksheet');
					goToSheet(1);
					removeCurrentWorksheet();
				}, oTableTypes.goToPreviousSheet);
			});

			QUnit.test('Test move to next sheet', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					createWorksheet();
					goToSheet(0);
					onKeyDown(oEvent);
					equal(wbView().wsActive, 1, 'Check got to next worksheet');
					removeCurrentWorksheet();
				}, oTableTypes.moveToNextSheet);

			});

			QUnit.test('Test move to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 39);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0), 'check move left');
				}, oTableTypes.moveToLeftEdgeCell);
			});

			QUnit.test('Test select to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 39);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 0, 0, 0), 'check move left');
				}, oTableTypes.selectToLeftEdgeCell);
			});

			QUnit.test('Test move to left cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 1);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0), 'check move left');
				}, oTableTypes.moveToLeftCell);
			});

			QUnit.test('Test select to left cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 1);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 0, 0, 1), 'check move left');
				}, oTableTypes.selectToLeftCell);
			});

			QUnit.test('Test move to right cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 1), 'check move left');
				}, oTableTypes.moveToRightCell);
			});

			QUnit.test('Test select to right cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 0, 0, 1), 'check move left');
				}, oTableTypes.selectToRightCell);
			});

			QUnit.test('Test move to top cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{

					moveToCell(25, 0);
					enterText('Hello');
					moveToCell(27, 0);
					enterText('Hello');
					moveToCell(35, 0);

					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(27, 0));

					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(25, 0));
				}, oTableTypes.moveToTopCell);
			});

			QUnit.test('Test move to upper cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{

					moveAndEnterText('Hello', 22, 0);
					moveToCell(33, 0);
					wsView().visibleRange = new Asc.Range(0, 0, 23, 35);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0));

				}, oTableTypes.moveToUpperCell);
			});

			QUnit.test('Test select to top cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{

					moveToCell(25, 0);
					enterText('Hello');
					moveToCell(27, 0);
					enterText('Hello');
					moveToCell(35, 0);

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(27, 35, 0, 0));

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(25, 35, 0, 0));
				}, oTableTypes.selectToTopCell);

			});

			QUnit.test('Test move to up cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(5, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(4, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(3, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(2, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(1, 0));
				}, oTableTypes.moveToUpCell);

			});

			QUnit.test('Test select to up cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(5, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 4, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 3, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 2, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 1, 0, 0));
				}, oTableTypes.selectToUpCell);

			});

			QUnit.test('Test move to bottom cell', (oAssert) =>
			{

				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(45, 0);
					enterText('Hello');
					moveToCell(47, 0);
					enterText('Hello');
					moveToCell(42, 0);

					wsView().visibleRange = new Asc.Range(0, 0, 49, 49);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(45, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(47, 0));
				}, oTableTypes.moveToBottomCell);

				startTableTest((oEvent) =>
				{
					moveToCell(45, 0);
					enterText('Hello');
					moveToCell(47, 0);
					enterText('Hello');
					moveToCell(42, 0);

					wsView().visibleRange = new Asc.Range(0, 0, 49, 49);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(92, 0));
				}, oTableTypes.moveToLowerCell);
			});

			QUnit.test('Test select to bottom cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(45, 0);
					enterText('Hello');
					moveToCell(47, 0);
					enterText('Hello');
					moveToCell(42, 0);
					wsView().visibleRange = new Asc.Range(0, 0, 100, 100);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(42, 45, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(42, 47, 0, 0));
				}, oTableTypes.selectToBottomCell);
			});

			QUnit.test('Test select to lower cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(45, 0);
					enterText('Hello');
					moveToCell(47, 0);
					enterText('Hello');
					moveToCell(42, 0);

					wsView().visibleRange = new Asc.Range(0, 0, 100, 100);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(42, 143, 0, 0));
				}, oTableTypes.selectToLowerCell);
			});

			QUnit.test('Test move to down cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(1, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(2, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(3, 0));
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(4, 0));
				}, oTableTypes.moveToDownCell);
			});

			QUnit.test('Test select to down cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					moveToCell(0, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 1, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 2, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 3, 0, 0));
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 4, 0, 0));
				}, oTableTypes.selectToDownCell);
			});


			QUnit.test('Test move to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(5, 25);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(5, 0));
				}, oTableTypes.moveToFirstColumn);
			});

			QUnit.test('Test select to left edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(5, 25);

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 5, 0, 25));
				}, oTableTypes.selectToFirstColumn);

			});

			QUnit.test('Test move to left edge top', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(5, 25);

					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0));
				}, oTableTypes.moveToLeftEdgeTop);
			});

			QUnit.test('Test select to left edge top', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(5, 25);

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 5, 0, 25));
				}, oTableTypes.selectToLeftEdgeTop);
			});


			QUnit.test('Test select to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll()
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(0, 0);

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 0, 0, 23));

					moveToCell(4, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(4, 4, 0, 8));

					moveToCell(5, 0);
					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(5, 5, 0, 5));
				}, oTableTypes.selectToRightEdgeCell);
			});

			QUnit.test('Test move to right edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(0, 0);

					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 23));

					moveToCell(4, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(4, 8));

					moveToCell(5, 0);
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(5, 5));
				}, oTableTypes.moveToRightEdgeCell);
			});

			QUnit.test('Test move to right bottom edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll()
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(0, 0);

					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(5, 8));
				}, oTableTypes.moveToRightBottomEdge);
			});

			QUnit.test('Test select to right bottom edge cell', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('Hello', 5, 5);
					moveAndEnterText('Hello', 4, 8);
					moveToCell(0, 0);

					onKeyDown(oEvent);
					deep(cleanSelection(), checkRange(0, 5, 0, 8));
				}, oTableTypes.selectToRightBottomEdge);
			});

			QUnit.test('Test set number format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveAndEnterText('49990', 5, 5);
					onKeyDown(oEvent);
					equal(getCellText(), '49990.00', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setNumberFormat);
			});


			QUnit.test('Test set time format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('49990');
					onKeyDown(oEvent);
					equal(getCellText(), '12:00:00 AM', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setTimeFormat);
			});

			QUnit.test('Test set date format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('49990');
					onKeyDown(oEvent);
					equal(getCellText(), '11/11/2036', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setDateFormat);
			});

			QUnit.test('Test set currency format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('49990');
					onKeyDown(oEvent);
					equal(getCellText(), '$49,990.00', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setCurrencyFormat);
			});

			QUnit.test('Test set percent format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(getCellText(), '10.00%', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setPercentFormat);
			});


			QUnit.test('Test strikethrough', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(xfs().asc_getFontStrikeout(), true);
					onKeyDown(oEvent);
					equal(xfs().asc_getFontStrikeout(), false);
				}, oTableTypes.setStrikethrough);
			});

			QUnit.test('Test set exponential format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(getCellText(), '1.00E-01', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setExponentialFormat);
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(xfs().asc_getFontBold(), true);
					onKeyDown(oEvent);
					equal(xfs().asc_getFontBold(), false);
				}, oTableTypes.setBold);
			});

			QUnit.test('Test italic', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(xfs().asc_getFontItalic(), true);
					onKeyDown(oEvent);
					equal(xfs().asc_getFontItalic(), false);
				}, oTableTypes.setItalic);
			});

			QUnit.test('Test underline', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(xfs().asc_getFontUnderline(), true);
					onKeyDown(oEvent);
					equal(xfs().asc_getFontUnderline(), false);
				}, oTableTypes.setUnderline);
			});


			QUnit.test('Test set general format', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(5, 5);
					enterText('0.1');
					setCellFormat(Asc.c_oAscNumFormatType.Time);

					onKeyDown(oEvent);
					equal(getCellText(), '0.1', 'set number format');
					setCellFormat(Asc.c_oAscNumFormatType.General);
				}, oTableTypes.setGeneralFormat);
			});


			QUnit.test('Test redo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					undo();
					onKeyDown(oEvent);
					equal(getCellText(), '0.1');
				}, oTableTypes.redo);
			});

			QUnit.test('Test undo', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					moveToCell(6, 6);
					enterText('0.1');
					onKeyDown(oEvent);
					equal(getCellText(), '');
				}, oTableTypes.undo);
			});

			QUnit.test('Test print', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					executeTestWithCatchEvent('asc_onPrint', () => true, true, oEvent, oAssert);
				}, oTableTypes.print);
			});

			QUnit.test('Test add sum function', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					cleanAll();
					goToSheet(0);
					moveToCell(0, 0);
					enterText('1');
					moveToCell(0, 1);
					enterText('2');
					moveToCell(7, 7);
					setCellFormat(Asc.c_oAscNumFormatType.General);
					onKeyDown(oEvent);
					enterTextWithoutClose('A1,B1');
					equal(getCellText(), "3");
				}, oTableTypes.addSum);
			});

			QUnit.test('Test context menu', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startTableTest((oEvent) =>
				{
					executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
				}, oTableTypes.contextMenu);
			});

			QUnit.module("test cell editor shortcuts", {
				before    : function ()
				{
					goToSheet(0);
				},
				beforeEach: function ()
				{
					setCheckOpenCellEditor(false);
					goToSheet(0);
					cleanAll();
				},
				afterEach : function ()
				{
					if (!checkOpenCellEditor())
					{
						throw new Error('cell editor must be opened in cell editor module');
					}
				}
			});
			QUnit.test('Test close cell editor', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					equal(getCellText(), '');
				}, oCellEditorTypes.closeWithoutSave);

			});
			QUnit.test('Test add new line', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					equal(cellEditor().textRender.getLinesCount(), 2);

					onKeyDown(oEvent);
					equal(cellEditor().textRender.getLinesCount(), 3);

					onKeyDown(oEvent);
					equal(cellEditor().textRender.getLinesCount(), 4);

					onKeyDown(oEvent);
					equal(cellEditor().textRender.getLinesCount(), 5);
				}, oCellEditorTypes.addNewLine);
			});

			QUnit.test('Try close editor', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(1, 0));

					moveToCell(0, 0);
					equal(getCellText(), 'Hello');
				}, oCellEditorTypes.saveAndMoveDown);
				cleanAll();
				startCellEditorTest((oEvent) =>
				{
					moveToCell(1, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0));

					moveToCell(1, 0);
					equal(getCellText(), 'Hello');
				}, oCellEditorTypes.saveAndMoveUp);
			});

			QUnit.test('Test sync and close editor', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 1));
					moveToCell(0, 0);
					equal(getCellText(), 'Hello');
				}, oCellEditorTypes.saveAndMoveRight);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 1);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					deep(cleanActiveCell(), checkActiveCell(0, 0));
					moveToCell(0, 1);
					equal(getCellText(), 'Hello');
				}, oCellEditorTypes.saveAndMoveLeft);
			});
			QUnit.test('Test remove char back', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello');
					onKeyDown(oEvent);
					equal(getCellText(), 'Hell');
				}, oCellEditorTypes.removeCharBack);

			});
			QUnit.test('Test remove word back', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World');
					onKeyDown(oEvent);
					equal(getCellText(), 'Hello ');
				}, oCellEditorTypes.removeWordBack);

			});

			QUnit.test('Test add space in cell editor', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					onKeyDown(oEvent);
					equal(getCellText(), ' ');
				}, oCellEditorTypes.addSpace);

			});

			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 18);
				}, oCellEditorTypes.moveToEndLine);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 47);
				}, oCellEditorTypes.moveToEndDocument);
			});

			QUnit.test('Test select to end', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello World Hello ');
				}, oCellEditorTypes.selectToEndLine);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
				}, oCellEditorTypes.selectToEndDocument);
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 36);
				}, oCellEditorTypes.moveToStartLine);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 0);
				}, oCellEditorTypes.moveToStartDocument);
			});
			QUnit.test('Test select to start', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello World');
				}, oCellEditorTypes.selectToStartLine);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
				}, oCellEditorTypes.selectToStartDocument);
			});
			QUnit.test('Test move cursor to left', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 46);
				}, oCellEditorTypes.moveCursorLeftChar);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 42);
				}, oCellEditorTypes.moveCursorLeftWord);
			});
			QUnit.test('Test select to left', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'd');
				}, oCellEditorTypes.selectLeftChar);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToEndCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'World');
				}, oCellEditorTypes.selectLeftWord);
			});
			QUnit.test('Test move cursor to up', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					onKeyDown(oEvent);
					equal(cellPosition(), 29);
				}, oCellEditorTypes.moveToUpLine);

			});
			QUnit.test('Test select to up', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), ' World Hello World');
				}, oCellEditorTypes.selectToUpLine);
			});
			QUnit.test('Test move cursor to right', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 1);
				}, oCellEditorTypes.moveToRightChar);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 6);
				}, oCellEditorTypes.moveToRightWord);
			});
			QUnit.test('Test select to right', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'H');
				}, oCellEditorTypes.selectRightChar);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello ');
				}, oCellEditorTypes.selectRightWord);
			});

			QUnit.test('Test move cursor to down', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(cellPosition(), 18);
				}, oCellEditorTypes.moveToDownLine);
			});
			QUnit.test('Test select to down', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getSelectionCellEditor(), 'Hello World Hello ');
				}, oCellEditorTypes.selectToDownLine);
			});

			QUnit.test('Test delete front', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getCellText(), 'ello World Hello World Hello World Hello World');
				}, oCellEditorTypes.deleteFrontChar);

				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					equal(getCellText(), 'World Hello World Hello World Hello World');
				}, oCellEditorTypes.deleteFrontWord);
			});

			QUnit.test('Test strikethrough', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					enterTextWithoutClose('hihih');
					const arrFragments = getFragments(0, 5);
					equal(arrFragments.length, 1);
					equal(arrFragments[0].format.getStrikeout(), true);
				}, oCellEditorTypes.setStrikethrough);
			});

			QUnit.test('Test select all text', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);

					equal(getSelectionCellEditor(), 'Hello World Hello World Hello World Hello World');
				}, oCellEditorTypes.selectAll);
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					enterTextWithoutClose('hihih');
					const arrFragments = getFragments(0, 5);
					equal(arrFragments.length, 1);
					equal(arrFragments[0].format.getBold(), true);
				}, oCellEditorTypes.setBold);
			});

			QUnit.test('Test italic', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					enterTextWithoutClose('hihih');
					const arrFragments = getFragments(0, 5);
					equal(arrFragments.length, 1);
					equal(arrFragments[0].format.getItalic(), true);
				}, oCellEditorTypes.setItalic);
			});

			QUnit.test('Test underline', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('Hello World Hello World Hello World Hello World');
					moveToStartCellEditor();
					onKeyDown(oEvent);
					enterTextWithoutClose('hihih');
					const arrFragments = getFragments(0, 5);
					equal(arrFragments.length, 1);
					equal(arrFragments[0].format.getUnderline(), Asc.EUnderline.underlineSingle);
				}, oCellEditorTypes.setUnderline);
			});

			QUnit.test('Test disable scroll lock', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					AscCommon.AscBrowser.isOpera = true;
					testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert);

					AscCommon.AscBrowser.isOpera = false;
					testPreventDefaultAndStopPropagation(createEvent(145, false, false, false, false, false), oAssert, true);
				}, oCellEditorTypes.disableScrollLock);
			});
			QUnit.test('Test disable num lock', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					AscCommon.AscBrowser.isOpera = true;
					testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert);

					AscCommon.AscBrowser.isOpera = false;
					testPreventDefaultAndStopPropagation(createEvent(144, false, false, false, false, false), oAssert, true);
				}, oCellEditorTypes.disableNumLock);
			});

			QUnit.test('Test print', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					testPreventDefaultAndStopPropagation(createEvent(80, true, false, false, false, false), oAssert);
				}, oCellEditorTypes.disablePrint);
			});

			QUnit.test('Test undo', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('H');
					enterTextWithoutClose('e');
					enterTextWithoutClose('l');
					enterTextWithoutClose('l');
					enterTextWithoutClose('o');
					onKeyDown(oEvent);
					equal(getCellText(), 'Hell');
				}, oCellEditorTypes.undo);
			});

			QUnit.test('Test redo', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					enterTextWithoutClose('H');
					enterTextWithoutClose('e');
					enterTextWithoutClose('l');
					enterTextWithoutClose('l');
					enterTextWithoutClose('o');
					cellEditor().undo();
					onKeyDown(oEvent);
					equal(getCellText(), 'Hello');
				}, oCellEditorTypes.redo);
			});

			QUnit.test('Test add separator', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					onKeyDown(oEvent);
					equal(getCellText(), '.');
				}, oCellEditorTypes.addSeparator);
			});
			QUnit.test('Test disable F2', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();

					AscCommon.AscBrowser.isOpera = true;
					testPreventDefaultAndStopPropagation(createEvent(113, false, false, false, false, false), oAssert);

					AscCommon.AscBrowser.isOpera = false;
					testPreventDefaultAndStopPropagation(createEvent(113, false, false, false, false, false), oAssert, true);
				}, oCellEditorTypes.disableF2);
			});

			QUnit.test('Test switch reference', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					enterTextWithoutClose('=F4');

					onKeyDown(oEvent);
					selectAllCell();
					equal(getSelectionCellEditor(), '=$F$4');
					closeCellEditor();
				}, oCellEditorTypes.switchReference);

			});

			QUnit.test('Test add time', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					onKeyDown(oEvent);
					const oDate = new Asc.cDate();
					equal(getCellText(), oDate.getTimeString(editor).split(' ').join(':00 '), 'Check insert current time');
				}, oCellEditorTypes.addTime);
			});

			QUnit.test('Test add date', (oAssert) =>
			{
				const {equal, deep} = createTest(oAssert);
				startCellEditorTest((oEvent) =>
				{
					moveToCell(0, 0);
					openCellEditor();
					onKeyDown(oEvent);
					const oDate = new Asc.cDate();
					equal(getCellText(), oDate.getDateString(editor), 'Check insert current date');
				}, oCellEditorTypes.addDate);
			});

			QUnit.module('Test graphic objects shortcuts');
			QUnit.test('Test remove back text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkTextAfterKeyDownHelperHelloWorld('Hello Worl', oEvent, oAssert, 'Check delete with backspace');
				}, oGraphicTypes.removeBackChar);
			});
			QUnit.test('Test remove back text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkTextAfterKeyDownHelperHelloWorld('Hello ', oEvent, oAssert, 'Check delete word with backspace')
				}, oGraphicTypes.removeBackWord);
			});

			QUnit.test('Test remove chart', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oChart = createChart();
					selectOnlyObjects([oChart]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oChart, ws().Drawings.map((el) => el.graphicObject)), "Check remove group");
				}, oGraphicTypes.removeChart);
			});

			QUnit.test('Test remove shape', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oShape, ws().Drawings.map((el) => el.graphicObject)), 'Check remove shape');
				}, oGraphicTypes.removeShape);
			});

			QUnit.test('Test remove group', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oGroup = createGroup([createShape(), createShape()]);
					selectOnlyObjects([oGroup]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oGroup, ws().Drawings.map((el) => el.graphicObject)), 'Check remove group');
				}, oGraphicTypes.removeGroup);
			});
			QUnit.test('Test remove shape in group', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oGroupedGroup = createGroup([createShape(), createShape()]);
					const oRemovedShape = createShape();
					const oGroup = createGroup([oGroupedGroup, oRemovedShape]);
					selectOnlyObjects([oRemovedShape]);
					onKeyDown(oEvent);
					oAssert.true(checkRemoveObject(oRemovedShape, oGroup.spTree), 'Check remove shape in group');
				}, oGraphicTypes.removeShapeInGroup);
			});

			QUnit.test('Test add tab', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);
					onKeyDown(oEvent);
					selectAll();
					equal(selectedContent(), '\t');
				}, oGraphicTypes.addTab);
			});
			QUnit.test('Test select next object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape2, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape3, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape1, true);
				}, oGraphicTypes.selectNextObject);
			});

			QUnit.test('Test select previous object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape3, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape2, true);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 1 && selectedObjects()[0] === oShape1, true);
				}, oGraphicTypes.selectPreviousObject);
			});

			QUnit.test('Test visit hyperlink', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('', 0, false);
					moveToParagraph(oParagraph);
					addHyperlink("https://www.onlyoffice.com/", "Hello");
					moveCursorLeft();
					moveCursorLeft();
					executeTestWithCatchEvent('asc_onHyperlinkClick', (sValue) => sValue, 'https://www.onlyoffice.com/', oEvent, oAssert);
					moveToParagraph(oParagraph);
					moveCursorLeft();
					const oSelectedInfo = new CSelectedElementsInfo();
					graphicController().getTargetDocContent().GetSelectedElementsInfo(oSelectedInfo);
					equal(oSelectedInfo.m_oHyperlink.Visited, true);
				}, oGraphicTypes.visitHyperink);
			});

			QUnit.test('Test add line in math', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oParagraph} = createMathInShape();
					moveToParagraph(oParagraph, true);
					moveCursorRight();
					moveCursorRight();
					enterText('fffffff');
					moveCursorLeft();
					moveCursorLeft();
					onKeyDown(oEvent);
					const oParaMath = oParagraph.GetAllParaMaths()[0];
					const oFraction = oParaMath.Root.GetFirstElement();
					const oNumerator = oFraction.getNumerator();
					const oEqArray = oNumerator.GetFirstElement();
					oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');
				}, oGraphicTypes.addLineInMath);
			});
			QUnit.test('Test add break line', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);

					onKeyDown(oEvent);
					oAssert.true(oShape.getDocContent().Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line');
				}, oGraphicTypes.addBreakLine);
			});

			QUnit.test('Test add paragraph', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('');
					moveToParagraph(oParagraph);

					onKeyDown(oEvent);
					oAssert.true(oShape.getDocContent().Content.length === 2, 'Check add new paragraph');
				}, oGraphicTypes.addParagraph);
			});

			QUnit.test('Test create text body', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oShape = createShape();
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(!!oShape.txBody, 'Check creating txBody');
				}, oGraphicTypes.createTxBody);
			});

			QUnit.test('Test move cursor to start position in empty content', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape, oParagraph} = createShapeWithContent('', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.true(oParagraph.IsCursorAtBegin(), 'Check move cursor to start position in shape');
				}, oGraphicTypes.moveToStartInEmptyContent);
			});
			QUnit.test('Test select all after enter', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello Word', true);
					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					oAssert.strictEqual(selectedContent(), 'Hello Word', 'Check select all content in shape');
				}, oGraphicTypes.selectAllAfterEnter);
			});
			QUnit.test('Test select all after enter in title', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oChart = createChart();
					selectOnlyObjects([oChart]);
					const oTitles = oChart.getAllTitles();
					const oController = graphicController();
					oController.selection.chartSelection = oChart;
					oChart.selectTitle(oTitles[0], 0);

					onKeyDown(oEvent);
					oAssert.strictEqual(selectedContent(), 'Diagram Title', 'Check select all title');
				}, oGraphicTypes.selectAllTitleAfterEnter);
			});
			QUnit.test('Test move cursor to start position in empty title', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oChart = createChart();
					const oTitles = oChart.getAllTitles();
					const oContent = AscFormat.CreateDocContentFromString('', graphicController().drawingObjects.getDrawingDocument(), oTitles[0].txBody);
					oTitles[0].txBody.content = oContent;
					selectOnlyObjects([oChart]);

					const oController = graphicController();
					oController.selection.chartSelection = oChart;
					oChart.selectTitle(oTitles[0], 0);

					onKeyDown(oEvent);
					oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');
				}, oGraphicTypes.moveCursorToStartPositionInTitle);
			});
			QUnit.test('Test reset text selection', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent("Hello world");
					const oParagraph = moveToShapeParagraph(oShape, 0);
					onKeyDown(oEvent);
					equal(!graphicController().selection.textSelection, true);
				}, oGraphicTypes.resetTextSelection);
			});

			QUnit.test('Test reset step selection', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					let oGroupedShape1 = createShape();
					let oGroupedShape2 = createShape();
					createGroup([oGroupedShape1, oGroupedShape2]);
					let oTestGroup = oGroupedShape1.group;
					const oController = graphicController();
					selectOnlyObjects([oTestGroup, oGroupedShape1]);
					onKeyDown(oEvent);
					oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oTestGroup && oTestGroup.selectedObjects.length === 0, 'Check reset step selection');

				}, oGraphicTypes.resetStepSelection);
			});
			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 18], oEvent, oAssert);
				}, oGraphicTypes.moveCursorToEndLine);
			});

			QUnit.test('Test move cursor to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 59], oEvent, oAssert);
				}, oGraphicTypes.moveCursorToEndDocument);
			});


			QUnit.test('Test select to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello ', ''], oEvent, oAssert);
				}, oGraphicTypes.selectToEndLine);
			});

			QUnit.test('Test select to end', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello World Hello World Hello World Hello World', ''], oEvent, oAssert);
				}, oGraphicTypes.selectToEndDocument);
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([54, 0], oEvent, oAssert);
				}, oGraphicTypes.moveCursorToStartLine);
			});

			QUnit.test('Test move cursor to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([0, 0], oEvent, oAssert);
				}, oGraphicTypes.moveCursorToStartDocument);
			});

			QUnit.test('Test select to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'World'], oEvent, oAssert);
				}, oGraphicTypes.selectToStartLine);
			});

			QUnit.test('Test select to start', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'Hello World Hello World Hello World Hello World Hello World'], oEvent, oAssert);
				}, oGraphicTypes.selectToStartDocument);
			});

			QUnit.test('Test move cursor to left char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([58, 0], oEvent, oAssert);

				}, oGraphicTypes.moveCursorLeftChar);
			});

			QUnit.test('Test select to left char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'd'], oEvent, oAssert);
				}, oGraphicTypes.selectCursorLeftChar);
			});

			QUnit.test('Test move cursor to left word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([54, 0], oEvent, oAssert);
				}, oGraphicTypes.moveCursorLeftWord);
			});

			QUnit.test('Test select to left word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', 'World'], oEvent, oAssert);
				}, oGraphicTypes.selectCursorLeftWord);
			});

			QUnit.test('Test move object to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100 - pxToMm, y: 100}, oEvent, oAssert);
				}, oGraphicTypes.littleMoveGraphicObjectLeft);
			});

			QUnit.test('Test move object to left', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100 - 5 * pxToMm, y: 100}, oEvent, oAssert);

				}, oGraphicTypes.bigMoveGraphicObjectLeft);
			});

			QUnit.test('Test move cursor to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([41, 0], oEvent, oAssert);

				}, oGraphicTypes.moveCursorUp);
			});

			QUnit.test('Test select to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['', ' World Hello World'], oEvent, oAssert);
				}, oGraphicTypes.selectCursorUp);
			});

			QUnit.test('Test move object to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 - pxToMm}, oEvent, oAssert);
				}, oGraphicTypes.littleMoveGraphicObjectUp);
			});

			QUnit.test('Test move object to up', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 - 5 * pxToMm}, oEvent, oAssert);
				}, oGraphicTypes.bigMoveGraphicObjectUp);
			});

			QUnit.test('Test move cursor to right char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 1], oEvent, oAssert);

				}, oGraphicTypes.moveCursorRightChar);
			});

			QUnit.test('Test select to right char', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['H', ''], oEvent, oAssert);
				}, oGraphicTypes.selectCursorRightChar);
			});

			QUnit.test('Test move cursor to right word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 6], oEvent, oAssert);
				}, oGraphicTypes.moveCursorRightWord);
			});

			QUnit.test('Test select to right word', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello ', ''], oEvent, oAssert);
				}, oGraphicTypes.selectCursorRightWord);
			});

			QUnit.test('Test move object to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100 + pxToMm, y: 100}, oEvent, oAssert);

				}, oGraphicTypes.littleMoveGraphicObjectRight);
			});

			QUnit.test('Test move object to right', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100 + 5 * pxToMm, y: 100}, oEvent, oAssert);

				}, oGraphicTypes.bigMoveGraphicObjectRight);
			});

			QUnit.test('Test move cursor to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkMoveContentShapeHelper([59, 18], oEvent, oAssert);
				}, oGraphicTypes.moveCursorDown);
			});

			QUnit.test('Test select to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkSelectContentShapeHelper(['Hello World Hello ', ''], oEvent, oAssert);
				}, oGraphicTypes.selectCursorDown);
			});

			QUnit.test('Test move object to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 + pxToMm}, oEvent, oAssert);

				}, oGraphicTypes.littleMoveGraphicObjectDown);
			});

			QUnit.test('Test move object to down', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					moveShapeHelper({x: 100, y: 100 + 5 * pxToMm}, oEvent, oAssert);
				}, oGraphicTypes.bigMoveGraphicObjectDown);
			});

			QUnit.test('Test remove front text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('Hello World');
					moveToParagraph(oParagraph, true);
					onKeyDown(oEvent);
					equal(getParagraphText(oParagraph), 'ello World');
				}, oGraphicTypes.removeFrontChar);
			});

			QUnit.test('Test remove front text graphic object', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oParagraph} = createShapeWithContent('Hello World');
					moveToParagraph(oParagraph, true);
					onKeyDown(oEvent);
					equal(getParagraphText(oParagraph), 'World');
				}, oGraphicTypes.removeFrontWord);
			});
			QUnit.test('Test select all content in shape', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello World');
					const oParagraph = moveToShapeParagraph(oShape, 0);
					onKeyDown(oEvent);
					equal(selectedContent(), 'Hello World');
				}, oGraphicTypes.selectAllContent);
			});
			QUnit.test('Test select all graphic objects', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oShape1 = createShape();
					const oShape2 = createShape();
					const oShape3 = createShape();
					selectOnlyObjects([oShape1]);
					onKeyDown(oEvent);
					equal(selectedObjects().length === 3 && (oShape1.selected && selectedObjects().indexOf(oShape1) !== -1) && (oShape2.selected && selectedObjects().indexOf(oShape2) !== -1) && (oShape3.selected && selectedObjects().indexOf(oShape3) !== -1), true);
				}, oGraphicTypes.selectAllDrawings);
			});

			QUnit.test('Test bold', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetBold(), oEvent, oAssert, true);

				}, oGraphicTypes.bold);
			});

			QUnit.test('Test clear slicer', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oSlicer = createSlicer();
					oSlicer.buttonsContainer.buttons[0].setInvertSelectTmpState();
					oSlicer.onViewUpdate();
					onKeyDown(oEvent);
					equal(oSlicer.buttonsContainer.buttons[0].isSelected(), true);
				}, oGraphicTypes.cleanSlicer);
			});

			QUnit.test('Test center align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Center);

				}, oGraphicTypes.centerAlign);
			});
			QUnit.test('Test italic', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetItalic(), oEvent, oAssert, true);

				}, oGraphicTypes.italic);
			});
			QUnit.test('Test justify align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Justify);

				}, oGraphicTypes.justifyAlign);
			});
			QUnit.test('Test left align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const continueCheck = checkDirectParaPrAfterKeyDown((oPara) => oPara.GetJc(), oEvent, oAssert, align_Justify);
					continueCheck((oPara) => oPara.GetJc(), oEvent, oAssert, align_Left);

				}, oGraphicTypes.leftAlign);
			});
			QUnit.test('Test right align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oEvent, oAssert, align_Right);
				}, oGraphicTypes.rightAlign);
			});

			QUnit.test('Test invert multiselect slicer', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const oSlicer = createSlicer();
					onKeyDown(oEvent);
					equal(oSlicer.isMultiSelect(), true);
				}, oGraphicTypes.invertMultiselectSlicer);
			});
			QUnit.test('Test underline', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetUnderline(), oEvent, oAssert, true);
				}, oGraphicTypes.underline);
			});

			QUnit.test('Test superscript vertical align', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oEvent, oAssert, AscCommon.vertalign_SuperScript);

				}, oGraphicTypes.superscript);
			});

			QUnit.test('Test add en dash', (oAssert) =>
			{
				startGraphicTest((oEvent) =>
				{
					checkTextAfterKeyDownHelperEmpty('–', oEvent, oAssert);
				}, oGraphicTypes.enDash);
			});

			QUnit.test('Test add subscript', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oEvent, oAssert, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
				}, oGraphicTypes.subscript);
			});

			QUnit.test('Test decrease font size', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello');
					selectAll();
					addProperty({FontSize: 16});
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 14);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 12);

					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 11);

					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 10);
				}, oGraphicTypes.decreaseFontSize);
			});

			QUnit.test('Test increase font size', (oAssert) =>
			{
				const {deep, equal} = createTest(oAssert);
				startGraphicTest((oEvent) =>
				{
					const {oShape} = createShapeWithContent('Hello');
					moveToShapeParagraph(oShape, 0);
					selectAll();
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 11);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 12);

					selectOnlyObjects([oShape]);
					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 14);

					onKeyDown(oEvent);
					equal(textPr().GetFontSize(), 16);
				}, oGraphicTypes.increaseFontSize);
			});
		}
	)
})(window);
