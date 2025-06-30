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

QUnit.config.autostart = false;
(function (window)
{
	const {
		tableHotkeyTypes,
		cellEditorHotkeyTypes,
		graphicHotkeyTypes,
		tableEvents,
		graphicEvents,
		ExecuteCellEditorHotkey,
		ExecuteTableHotkey,
		ExecuteGraphicHotkey,
		ExecuteShortcut,
		InitEditor
	} = AscTestShortcut;

	let editor, wb, wbView, ws, wsView, cellEditor;
	InitEditor(function ()
	{
		editor = window["Asc"]["editor"];
		wb = editor.wbModel;
		wbView = editor.wb;
		ws = wb.aWorksheets[0];
		wsView = wbView.getWorksheet();
		cellEditor = wbView.cellEditor;
		QUnit.start();
	});

	function UpdateView()
	{
		wsView._cleanCache(new Asc.Range(0, 0, wsView.cols.length - 1, wsView.rows.length - 1));
		wsView.changeWorksheet("update", {reinitRanges: true});
	}

	function GetCellCacheText(column, row)
	{
		return wsView._getCellTextCache(column, row).state.chars.map((e) => String.fromCharCode(e)).join('');
	}

	function GetParagraphText(paragraph)
	{
		return paragraph.GetText({ParaSeparator : ""});
	}

	function ClearShapeAndAddParagraph(sText)
	{
		const textShape = AddShape(0, 0, 100, 100);
		const txBody = AscFormat.CreateTextBodyFromString(sText, wb.DrawingDocument, textShape);
		textShape.setTxBody(txBody);
		textShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
		const paragraph = txBody.content.Content[0];
		paragraph.SetThisElementCurrent();
		paragraph.MoveCursorToStartPos();
		textShape.recalculate();
		return {shape: textShape, paragraph: paragraph};
	}

	function AddShape(x, y, height, width)
	{
		const shapeTrack = new AscFormat.NewShapeTrack('rect', x, y, wb.theme, null, null, null, 0);
		shapeTrack.track({}, x + width, y + height);
		const shape = shapeTrack.getShape(false, wb.DrawingDocument, null);
		shape.setBDeleted(false);
		shape.setWorksheet(ws);
		shape.setParent(GetDrawingObjects().drawingObjects);
		shape.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
		shape.checkDrawingBaseCoords();

		shape.select(GetDrawingObjects(), 0);
		return shape;
	}

	function EnterText(text)
	{
		const codePoints = text.split('').map((e) => e.charCodeAt());
		wbView.EnterText(codePoints);
	}

	function AddChart()
	{
		ws.getRange4(0, 0).fillData([['1']]);
		const props = editor.asc_getChartSettings(Asc.c_oAscChartTypeSettings.lineNormal);
		props.left = 0;
		props.top = 0;
		props.width = 100;
		props.height = 100;
		return GetDrawingObjects().addChartDrawingObject(props);
	}

	function CreateGroup(arrObjects)
	{
		const controller = GetDrawingObjects();
		controller.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			arrObjects[i].select(controller, 0);
		}

		return controller.createGroup();
	}

	function GetDirectGraphicTextPr()
	{
		return GetDrawingObjects().getParagraphTextPr();
	}

	function GetDirectGraphicParaPr()
	{
		return GetDrawingObjects().getParagraphParaPr();
	}

	function Round(number, amount)
	{
		const power = Math.pow(10, amount);
		return Math.round(number * power) / power;
	}

	function SelectDrawings(arrDrawings)
	{
		const drawingController = GetDrawingObjects();
		drawingController.resetSelection()
		for (let i = 0; i < arrDrawings.length; i += 1)
		{
			arrDrawings[i].select(drawingController, 0);
		}
	}

	function GetDrawingObjects()
	{
		return editor.getGraphicController();
	}

	function FillActiveCell(text)
	{
		const activeCell = ws.selectionRange.activeCell;
		ws.getRange4(activeCell.row, activeCell.col).fillData([[text]]);

	}

	function SetVisibleRange(r1, c1, r2, c2)
	{
		wsView.visibleRange = new Asc.Range(r1, c1, r2, c2);
		wsView.nRowsCount = 49;
		wsView.nColsCount = 49;
	}

	function FillExampleData()
	{
		ws.getRange3(0, 0, 2, 2).fillData([
			['test11', 'test12', 'test13'],
			['1', '2', '3'],
			['1', '2', '3']
		]);

		Select(0, 0, 0, 0, 2, 2);
	}

	function GetRange(c1, r1, c2, r2)
	{
		return new Asc.Range(c1, r1, c2, r2);
	}

	function Select(activeR, activeC, r1, c1, r2, c2)
	{
		CloseCellEditor(true);
		ws.selectionRange.ranges = [GetRange(c1, r1, c2, r2)];
		ws.selectionRange.activeCell = new AscCommon.CellBase(activeR, activeC);
		wbView._updateSelectionInfo();
	}

	function ResetData(r1, c1, r2, c2)
	{
		ws.autoFilters.deleteAutoFilter(GetRange(0, 0, 0, 0));
		while (ws.TableParts.length)
		{
			ws.deleteTablePart(0);
		}
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	}

	function GetCellText(row, column)
	{
		return ws.getRange4(row, column).getValueWithFormat();
	}

	function CloseCellEditor(save)
	{
		wbView.closeCellEditor(!save);
	}

	function CreateTable()
	{
		FillExampleData();
		ws.autoFilters.addAutoFilter('TableStyleMedium2', GetRange(0, 0, 2, 2));
	}

	function CreateSlicer()
	{
		CreateTable();
		editor.asc_insertSlicer(['test11']);
		return GetDrawingObjects().selectedObjects[0];
	}

	function RemoveGraphicObjects()
	{
		GetDrawingObjects().resetSelection();
		GetDrawingObjects().selectAll();
		GetDrawingObjects().remove();
	}

	function GetCellEditorText()
	{
		return AscCommonExcel.getFragmentsText(cellEditor.options.fragments);
	}

	function OpenCellEditor()
	{
		const enterOptions = new AscCommonExcel.CEditorEnterOptions();
		enterOptions.newText = '';
		enterOptions.quickInput = true;
		enterOptions.focus = true;
		wbView._onEditCell(enterOptions);
	}

	function GetSelectedCellEditorText()
	{
		const fragments = cellEditor.copySelection() || [];
		return AscCommonExcel.getFragmentsText(fragments)
	}

	function CreateWorksheets(names)
	{
		editor.asc_insertWorksheet(names);
	}

	function RemoveWorksheets(indexes)
	{
		editor.asc_deleteWorksheet(indexes);
	}

	QUnit.module('Test cell editor shortcut', {
		afterEach: function ()
		{
			CloseCellEditor();
		}
	});
	QUnit.test('Check actions with text movements', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'World Hello World');

		cellEditor._moveCursor(-2);

		function CheckCursorPosition(expected, description)
		{
			assert.strictEqual(cellEditor.cursorPos, expected, description);
		}

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine);
		CheckCursorPosition(18, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftChar);
		CheckCursorPosition(18, 'Check move to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftWord);
		CheckCursorPosition(12, 'Check move to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveCursorLeftWord, 1);
		CheckCursorPosition(6, 'Check move to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord, 1);
		CheckCursorPosition(12, 'Check move to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord);
		CheckCursorPosition(18, 'Check move to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToRightWord);
		CheckCursorPosition(24, 'Check move to right word');


		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToDownLine);
		CheckCursorPosition(36, 'Check move down');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToUpLine);
		CheckCursorPosition(18, 'Check move up');*/

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToEndDocument);
		CheckCursorPosition(161, 'Check move to end document');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartDocument);
		CheckCursorPosition(0, 'Check move to start document');

		function CheckSelectedText(expectedText, description)
		{
			assert.strictEqual(GetSelectedCellEditorText(), expectedText, description);
		}

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine);
		CheckSelectedText('Hello World Hello ', 'Select to end line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftChar);
		CheckSelectedText('Hello World Hello ', 'Select to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord, 1);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord, 1);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello World ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectRightWord);
		CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartLine);
		CheckSelectedText('Hello World Hello ', 'Select to start line');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToDownLine);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToUpLine);
		CheckSelectedText('Hello World Hello ', 'Select up');*/

		cellEditor._moveCursor(-4);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftChar);
		CheckSelectedText('d', 'Select to left char');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectLeftWord);
		CheckSelectedText('World', 'Select to left word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToStartDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to start content');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndDocument);
		CheckSelectedText('', 'Select to end content');

		cellEditor._moveCursor(-2);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.selectToEndDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to end content');
		cellEditor._moveCursor(-4);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(144, 'Check move to start line');

		//TODO ****temporarily commenting on tests. this is due to the lack of logical separation of the end of the line and the beginning of the next line in the cell editor****
		/*ExecuteCellEditorHotkey(cellEditorHotkeyTypes.moveToStartLine);
		CheckCursorPosition(144, 'Check move to start line');*/
	});

	QUnit.test('Check remove parts of text', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();
		EnterText('Hello Hello Hello Hello Hello Hello Hello');
		cellEditor._moveCursor(-4);

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeCharBack);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeWordBack);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.removeWordBack, 1);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

		cellEditor._moveCursor(-2);
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontChar);
		assert.strictEqual(GetCellEditorText(), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontWord);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello Hello ', 'Check removing front word');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.deleteFrontWord, 1);
		assert.strictEqual(GetCellEditorText(), 'Hello Hello Hello ', 'Check removing front word');
	});

	QUnit.test('Check save and moving from cell', (assert) =>
	{

		function CheckSaveAndMoving(row, column, text)
		{
			assert.false(wbView.getCellEditMode(), 'Check closing cell editor');
			assert.deepEqual([ws.selectionRange.activeCell.row, ws.selectionRange.activeCell.col], [row, column], 'Check active cell');
			assert.strictEqual(GetCellText(5, 5), text, 'Check initial text');
		}

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello1');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveDown);
		CheckSaveAndMoving(6, 5, 'Hello1');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello2');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveUp);
		CheckSaveAndMoving(4, 5, 'Hello2');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello3');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveRight);
		CheckSaveAndMoving(5, 6, 'Hello3');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello4');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.saveAndMoveLeft);
		CheckSaveAndMoving(5, 4, 'Hello4');

		Select(5, 5, 5, 5, 5, 5);
		EnterText('Hello5');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.closeWithoutSave);
		CheckSaveAndMoving(5, 5, 'Hello4');
	});

	QUnit.test('Check change text formatting', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello Hello');
		cellEditor.selectAll();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getStrikeout()), 'Check cell editor strikeout format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getStrikeout()), 'Check cell editor strikeout format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getBold()), 'Check cell editor bold format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getBold()), 'Check cell editor bold format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getItalic()), 'Check cell editor italic format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.false(cellEditor._getFragments(0, 11).every((e) => e.format.getItalic()), 'Check cell editor italic format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getUnderline() === Asc.EUnderline.underlineSingle), 'Check cell editor underline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getUnderline() === Asc.EUnderline.underlineNone), 'Check cell editor underline format');

		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check init font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 12), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 14), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 16), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 18), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 20), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 22), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 24), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 26), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 28), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 36), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 48), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 72), "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 72), "Check increase font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 48), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 36), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 28), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 26), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 24), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 22), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 20), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 18), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 16), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 14), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 12), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 10), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 9), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 8), "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 8), "Check decrease font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.true(cellEditor._getFragments(0, 11).every((e) => e.format.getSize() === 11), "Check initial value");
	});

	QUnit.test('Check enter text in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		OpenCellEditor();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');
		EnterText('');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), '.', 'Check add separator');
		EnterText('');

		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 2, 'Check add new line');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 3, 'Check add new line');
		ExecuteCellEditorHotkey(cellEditorHotkeyTypes.addNewLine);
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 4, 'Check add new line');
	});

	QUnit.test('Check disabling shortcuts', (assert) =>
	{
		OpenCellEditor();
		assert.strictEqual(ExecuteCellEditorHotkey(cellEditorHotkeyTypes.disableScrollLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		assert.strictEqual(ExecuteCellEditorHotkey(cellEditorHotkeyTypes.disableNumLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		AscCommon.AscBrowser.isOpera = true;
		assert.strictEqual(ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		AscCommon.AscBrowser.isOpera = false;
		assert.strictEqual(ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Print) & keydownresult_PreventDefault, keydownresult_PreventDefault);
	});

	QUnit.test('Check select all', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('Hello');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		assert.strictEqual(GetSelectedCellEditorText(), 'Hello', 'Check select all');
	});

	QUnit.test('Check undo/redo in cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('H');
		EnterText('e');
		EnterText('l');
		EnterText('l');
		EnterText('o');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo);
		assert.strictEqual(GetCellEditorText(), 'Hell', 'Check undo');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo);
		assert.strictEqual(GetCellEditorText(), 'Hel', 'Check undo');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo);
		assert.strictEqual(GetCellEditorText(), 'Hell', 'Check redo');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo);
		assert.strictEqual(GetCellEditorText(), 'Hello', 'Check redo');
	});
	QUnit.test('Check switch reference of formula', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('=F4');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=$F$4', 'Check switch to absolute reference');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=F$4', 'Check switch to absolute row');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=$F4', 'Check switch to absolute column');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference);
		assert.strictEqual(GetCellEditorText(), '=F4', 'Check switch to relative reference');
	});

	QUnit.test('Test catch events', (assert) =>
	{
		function ExecuteTestWithCatchEvent(sSendEvent, fCustomCheck, customExpectedValue, oEvent)
		{
			let check = false;
			const Check = function (...args)
			{
				if (fCustomCheck)
				{
					check = fCustomCheck(...args);
				}
				else
				{
					check = true;
				}
			}
			editor.asc_registerCallback(sSendEvent, Check);
			ExecuteShortcut(oEvent);
			assert.strictEqual(check, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
			editor.asc_unregisterCallback(sSendEvent, Check);
		}

		ExecuteTestWithCatchEvent('asc_onPrint', () => true, true, Asc.c_oAscSpreadsheetShortcutType.Print);
		ExecuteTestWithCatchEvent('asc_onContextMenu', () => true, true, tableEvents[tableHotkeyTypes.contextMenu][0]);


		ClearShapeAndAddParagraph('');
		const hyperlink = new AscCommonExcel.Hyperlink();
		hyperlink.Hyperlink = "https://www.onlyoffice.com/";
		const ascHyperlink = new Asc.asc_CHyperlink(hyperlink);
		ascHyperlink.text = "Hello";
		editor.asc_insertHyperlink(ascHyperlink);
		GetDrawingObjects().cursorMoveLeft();
		GetDrawingObjects().cursorMoveLeft();

		ExecuteTestWithCatchEvent('asc_onHyperlinkClick', (sValue) => sValue, 'https://www.onlyoffice.com/', graphicEvents[graphicHotkeyTypes.visitHyperink][0]);
		const selectedInfo = new CSelectedElementsInfo();
		GetDrawingObjects().getTargetDocContent().GetSelectedElementsInfo(selectedInfo);
		assert.true(selectedInfo.m_oHyperlink.Visited, 'Check hyperlink visited');
		RemoveGraphicObjects();

	});
	QUnit.module('Test graphic shortcuts', {
		afterEach: function ()
		{
			RemoveGraphicObjects();
		}
	});
	QUnit.test('Check remove parts of text', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('Hello Hello Hello Hello Hello Hello Hello');
		GetDrawingObjects().cursorMoveToEndPos();

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackChar);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackWord);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeBackWord, 1);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

		GetDrawingObjects().cursorMoveToStartPos();
		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontChar);
		assert.strictEqual(GetParagraphText(paragraph), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontWord);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello Hello ', 'Check removing front word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.removeFrontWord, 1);
		assert.strictEqual(GetParagraphText(paragraph), 'Hello Hello Hello ', 'Check removing front word');
	});

	QUnit.test('Check remove graphic objects', (assert) =>
	{

		function CheckRemoveObject(spTree, shape)
		{
			assert.true(spTree.indexOf(shape) === -1, 'Check remove shape from spTree');
		}

		function CheckRemoveDrawingObject(shape)
		{
			assert.false(ws.Drawings.some((e) => e.graphicObject === shape), 'Check remove shape from spTree');
		}

		const CheckActionsWithEvent = (eventIndex) =>
		{
			const shape = AddShape(0, 0, 100, 100);
			SelectDrawings([shape]);
			assert.true(ws.Drawings.some((e) => e.graphicObject === shape), 'Check available shape in spTree');
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeShape, eventIndex);
			CheckRemoveDrawingObject(shape);

			const chart = AddChart();
			SelectDrawings([chart]);
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeChart, eventIndex);
			CheckRemoveDrawingObject(chart);


			const shape1 = AddShape(0, 0, 100, 100);
			const shape2 = AddShape(0, 0, 100, 100);
			const shape3 = AddShape(0, 0, 100, 100);

			const group = CreateGroup([shape1, shape2, shape3]);
			group.selectObject(shape1, 0);
			GetDrawingObjects().selection.groupSelection = group;

			ExecuteGraphicHotkey(graphicHotkeyTypes.removeShapeInGroup, eventIndex);
			CheckRemoveObject(group.spTree, shape1);

			SelectDrawings([group]);
			ExecuteGraphicHotkey(graphicHotkeyTypes.removeGroup, eventIndex);
			CheckRemoveDrawingObject(group);
		};

		CheckActionsWithEvent(0);
		CheckActionsWithEvent(1);
	});

	QUnit.test('Check main actions with shapes', (assert) =>
	{
		const drawing1 = AddShape(100, 100, 100, 200);

		function CheckShapePosition(X, Y)
		{
			assert.deepEqual([Round(drawing1.x, 12), Round(drawing1.y, 12), Round(drawing1.extX, 12), Round(drawing1.extY, 12)], [Round(X, 12), Round(Y, 12), 200, 100], 'Check shape position after movement ' + X + ' ' + Y);
		}

		SelectDrawings([drawing1]);


		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectLeft);
		CheckShapePosition(100 - 5 * AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectLeft);
		CheckShapePosition(100 - 6 * AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectRight);
		CheckShapePosition(100 - AscCommon.g_dKoef_pix_to_mm, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectRight);
		CheckShapePosition(100, 100);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectDown);
		CheckShapePosition(100, 100 + 5 * AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectDown);
		CheckShapePosition(100, 100 + 6 * AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.bigMoveGraphicObjectUp);
		CheckShapePosition(100, 100 + AscCommon.g_dKoef_pix_to_mm);

		ExecuteGraphicHotkey(graphicHotkeyTypes.littleMoveGraphicObjectUp);
		CheckShapePosition(100, 100);

		function CheckSelectedObjects(arrOfDrawings)
		{
			const length = Math.max(arrOfDrawings.length, GetDrawingObjects().selectedObjects.length);
			for (let i = 0; i < length; i++)
			{
				assert.true(GetDrawingObjects().selectedObjects[i] === arrOfDrawings[i], 'Check selection movement between objects');
			}
		}

		GetDrawingObjects().resetSelection();
		const drawing2 = AddShape(0, 0, 10, 10);
		const drawing3 = AddShape(0, 0, 10, 10);
		SelectDrawings([drawing3]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing1]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectNextObject);
		CheckSelectedObjects([drawing3]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing1]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectPreviousObject);
		CheckSelectedObjects([drawing3]);

		drawing2.setWordShape(false);
		SelectDrawings([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.createTxBody);
		assert.true(!!drawing2.txBody, 'Check create textBody after enter');

		SelectDrawings([drawing2]);
		ExecuteGraphicHotkey(graphicHotkeyTypes.moveToStartInEmptyContent);
		const paragraph = drawing2.getDocContent().Content[0];
		assert.true(drawing2.parent.controller.selection.textSelection === drawing2, 'Check text selection');
		assert.true(paragraph.IsCursorAtBegin(), 'Check movement to start position in empty content');

		EnterText('Hello');
		SelectDrawings([drawing2]);

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectAllAfterEnter);
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Hello\r\n', 'Check select non empty content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetTextSelection);
		assert.true(GetDrawingObjects().selectedObjects[0] === drawing2, 'Check reset text selection');
		assert.true(!GetDrawingObjects().selection.textSelection, 'Check reset text selection');


		let group = CreateGroup([drawing1, drawing2]);
		group.selectObject(drawing1, 0);
		GetDrawingObjects().selection.groupSelection = group;

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetStepSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects[0].getObjectType(), AscDFH.historyitem_type_GroupShape, 'Check reset step drawing selection');
		assert.strictEqual(group.selectedObjects.length, 0, 'Check reset step selection');

		ExecuteGraphicHotkey(graphicHotkeyTypes.resetStepSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects.length, 0, 'Check reset all drawing selection');
	});

	QUnit.test('Check actions with text movements', (assert) =>
	{
		const {shape} = ClearShapeAndAddParagraph(
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'World Hello World ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'Hello World Hello ' +
			'World Hello World');

		function CheckCursorPosition(expected, description)
		{
			const position = shape.txBody.content.GetContentPosition();
			assert.strictEqual(position[position.length - 1].Position, expected, description);
		}

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine);
		CheckCursorPosition(18, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndLine, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightChar);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftChar);
		CheckCursorPosition(18, 'Check move to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftWord);
		CheckCursorPosition(12, 'Check move to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorLeftWord, 1);
		CheckCursorPosition(6, 'Check move to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord, 1);
		CheckCursorPosition(12, 'Check move to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord);
		CheckCursorPosition(18, 'Check move to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorRightWord);
		CheckCursorPosition(24, 'Check move to right word');


		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartLine);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorDown);
		CheckCursorPosition(36, 'Check move down');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorUp);
		CheckCursorPosition(18, 'Check move up');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToEndDocument);
		CheckCursorPosition(161, 'Check move to end document');

		ExecuteGraphicHotkey(graphicHotkeyTypes.moveCursorToStartDocument);
		CheckCursorPosition(0, 'Check move to start document');

		function CheckSelectedText(expectedText, description)
		{
			const selectedText = GetDrawingObjects().GetSelectedText();
			assert.strictEqual(selectedText, expectedText, description);
		}

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine);
		CheckSelectedText('Hello World Hello ', 'Select to end line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndLine, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightChar);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftChar);
		CheckSelectedText('Hello World Hello ', 'Select to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord, 1);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord, 1);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello World ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorRightWord);
		CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartLine);
		CheckSelectedText('Hello World Hello ', 'Select to start line');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorDown);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorUp);
		CheckSelectedText('Hello World Hello ', 'Select up');

		GetDrawingObjects().cursorMoveToEndPos();
		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftChar);
		CheckSelectedText('d', 'Select to left char');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectCursorLeftWord);
		CheckSelectedText('World', 'Select to left word');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToStartDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to start content');

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndDocument);
		CheckSelectedText('\r\n', 'Select to end content');

		GetDrawingObjects().cursorMoveToStartPos();

		ExecuteGraphicHotkey(graphicHotkeyTypes.selectToEndDocument);
		CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World\r\n', 'Select to end content');
	});

	QUnit.test('Check paragraph property change', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello world');

		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check align left");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingCenterPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Center, "Check turn on center para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingCenterPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on center para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingJustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingJustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on justify para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingJustifyPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingLeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingLeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Justify, "Check turn on left para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingLeftPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingRightPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Right, "Check turn on right para");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingRightPara);
		assert.strictEqual(GetDirectGraphicParaPr().GetJc(), AscCommon.align_Left, "Check turn on right para");
	});

	QUnit.test('Check text property change', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello world');
		GetDrawingObjects().selectAll();

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.strictEqual(GetDirectGraphicTextPr().GetBold(), true, 'Check turn on bold');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.strictEqual(GetDirectGraphicTextPr().GetBold(), false, 'Check turn off bold');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.strictEqual(GetDirectGraphicTextPr().GetItalic(), true, 'Check turn on italic');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.strictEqual(GetDirectGraphicTextPr().GetItalic(), false, 'Check turn off italic');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.strictEqual(GetDirectGraphicTextPr().GetUnderline(), true, 'Check turn on underline');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.strictEqual(GetDirectGraphicTextPr().GetUnderline(), false, 'Check turn off underline');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingSuperscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_SuperScript, 'Check turn on superscript');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingSuperscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off superscript');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingSubscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_SubScript, 'Check turn on subscript');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingSubscript);
		assert.strictEqual(GetDirectGraphicTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off subscript');

		// defaultSize = 10
		// 10 -> 11 -> 12 -> 14 -> 16 -> 14 -> 12 -> 11 -> 10
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 11, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 12, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 14, 'Check increase font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 16, 'Check increase font size');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 14, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 12, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 11, 'Check decrease font size');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetDirectGraphicTextPr().GetFontSize(), 10, 'Check decrease font size');
	});

	QUnit.test('Check add various characters', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingEnDash);
		assert.strictEqual(GetParagraphText(paragraph), String.fromCharCode(0x2013), 'Check add en dash');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingEnDash);
		assert.strictEqual(GetParagraphText(paragraph), String.fromCharCode(0x2013, 0x2013), 'Check add en dash');
	});

	QUnit.test('Check select all', (assert) =>
	{
		const {paragraph, shape} = ClearShapeAndAddParagraph('Hello');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);

		assert.true(shape.txBody.content.IsSelectionUse(), 'Check content selection');
		assert.true(paragraph.IsSelectedAll(), 'Check paragraph selection');
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Hello\r\n', 'Check selected text');

		const drawing2 = AddShape();

		SelectDrawings([shape]);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		assert.strictEqual(GetDrawingObjects().selectedObjects.length, 2, 'Check amount of selected shapes');
		assert.true(GetDrawingObjects().selectedObjects[0] === drawing2, 'Check selected shape');
		assert.true(GetDrawingObjects().selectedObjects[1] === shape, 'Check selected shape');
	});

	QUnit.test('Check select all in chart title', (assert) =>
	{
		const chart = AddChart();
		SelectDrawings([chart]);
		const titles = chart.getAllTitles();
		const controller = GetDrawingObjects();
		controller.selection.chartSelection = chart;
		chart.selectTitle(titles[0], 0);
		ExecuteGraphicHotkey(graphicHotkeyTypes.selectAllTitleAfterEnter);
		assert.strictEqual(GetDrawingObjects().GetSelectedText(), 'Diagram Title\r\n', 'Check select all content in chart title');
	});

	QUnit.test('Check add tab symbol', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteGraphicHotkey(graphicHotkeyTypes.addTab);
		assert.true(paragraph.Content[0].Content[0].IsTab(), 'Check add tab symbol');
	});

	QUnit.test('Check add new paragraph', (assert) =>
	{
		const {shape} = ClearShapeAndAddParagraph('');
		ExecuteGraphicHotkey(graphicHotkeyTypes.addParagraph);
		assert.strictEqual(shape.txBody.content.Content.length, 2, 'Check add new paragraph');
	});

	QUnit.test('Check add break line', (assert) =>
	{
		function CheckAddBreakLine(shape)
		{
			assert.strictEqual(shape.txBody.content.Content.length, 1, 'Check amount paragraphs');
			assert.strictEqual(shape.txBody.content.Content[0].GetLinesCount(), 2, 'Check amount lines');
		}

		{

			const {shape} = ClearShapeAndAddParagraph('Hello Hello');
			GetDrawingObjects().cursorMoveRight(false, true);
			ExecuteGraphicHotkey(graphicHotkeyTypes.addBreakLine);
			CheckAddBreakLine(shape);

		}

		{
			const {shape} = ClearShapeAndAddParagraph('');
			editor.asc_AddMath2(c_oAscMathType.FractionVertical);
			GetDrawingObjects().recalculate();
			const paragraph = shape.txBody.content.Content[0];
			GetDrawingObjects().cursorMoveLeft();
			GetDrawingObjects().cursorMoveLeft();
			EnterText('8888888');
			GetDrawingObjects().cursorMoveLeft();
			GetDrawingObjects().cursorMoveLeft();
			ExecuteGraphicHotkey(graphicHotkeyTypes.addLineInMath);
			const paraMath = paragraph.GetAllParaMaths()[0];
			const fraction = paraMath.Root.GetFirstElement();
			const numerator = fraction.getNumerator();
			const eqArray = numerator.GetFirstElement();
			assert.strictEqual(eqArray.getRowsCount(), 2, 'Check add new line in math');
		}
	});

	QUnit.test('Check slicer shortcuts', (assert) =>
	{
		const ExecuteHotkeyWithEvent = (index) =>
		{
			const slicer = CreateSlicer();
			ExecuteGraphicHotkey(graphicHotkeyTypes.invertMultiselectSlicer, index);
			assert.true(slicer.isMultiSelect(), 'Check multiselect slicer');

			slicer.buttonsContainer.buttons[0].setUnselectTmpState();
			ExecuteGraphicHotkey(graphicHotkeyTypes.cleanSlicer, index);
			assert.true(slicer.buttonsContainer.buttons[0].isSelected(), 'Check clean slicer');
			RemoveGraphicObjects();
			ResetData(0, 0, 49, 49);
		};
		ExecuteHotkeyWithEvent(0);
		ExecuteHotkeyWithEvent(1);
	});


	QUnit.module('Test table hotkeys', {
		before   : function f()
		{
			SetVisibleRange(0, 0, 49, 49);
		},
		afterEach: function ()
		{
			CloseCellEditor(true);
			ResetData(0, 0, 49, 49);
		}
	});
	QUnit.test('Check drop down menu', (assert) =>
	{
		let check = false;
		const ShowAutoComplete = function () {check = true;};
		wbView.handlers.add("asc_onEntriesListMenu", ShowAutoComplete);
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showAutoComplete);
		assert.true(check);
		wbView.handlers.remove("asc_onEntriesListMenu", ShowAutoComplete);
		check = false;

		const props = ws.getDataValidationProps();
		props.setType(Asc.EDataValidationType.List);
		props.setFormula1(new Asc.CDataFormula('"1"'));
		ws.setDataValidationProps(props);
		const ShowDataValidations = function () {check = true;};
		wbView.handlers.add("asc_onValidationListMenu", ShowDataValidations);
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showDataValidation);
		assert.true(check);
		check = false;
		ExecuteTableHotkey(tableHotkeyTypes.showDataValidation, 1);
		assert.true(check);
		ws.deleteDataValidationById(props.Get_Id());
		wbView.handlers.remove("asc_onValidationListMenu", ShowDataValidations);
		check = false;

		const SetAutoFiltersDialog = function () {check = true;};
		wbView.handlers.add("setAutoFiltersDialog", SetAutoFiltersDialog);
		CreateTable();
		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.showFilterOptions);
		assert.true(check);
		wbView.handlers.remove("setAutoFiltersDialog", SetAutoFiltersDialog);
	});
	QUnit.test('Check cell movement', (assert) =>
	{
		function CheckActiveCell(expectedRow, expectedColumn, description)
		{
			assert.deepEqual([ws.selectionRange.activeCell.row, ws.selectionRange.activeCell.col], [expectedRow, expectedColumn], description);
		}

		function CheckActiveRange(r1, c1, r2, c2, description)
		{
			const lastRange = ws.selectionRange.getLast();
			assert.deepEqual([lastRange.r1, lastRange.c1, lastRange.r2, lastRange.c2], [r1, c1, r2, c2], description);
		}


		Select(0, 15, 0, 15, 0, 15);
		FillActiveCell('Hello');
		Select(25, 0, 25, 0, 25, 0);
		FillActiveCell('Hello');
		Select(27, 0, 27, 0, 27, 0);
		FillActiveCell('Hello');

		Select(5, 5, 5, 5, 5, 5);
		FillActiveCell('Hello');
		Select(4, 8, 4, 8, 4, 8);
		FillActiveCell('Hello');

		Select(35, 0, 35, 0, 35, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(27, 0, 'Check move to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(25, 0, 'Check move to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToTopCell);
		CheckActiveCell(0, 0, 'Check move to top cell');


		Select(0, 39, 0, 39, 0, 39);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeCell);
		CheckActiveCell(0, 15, 'Check move to left edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeCell);
		CheckActiveCell(0, 0, 'Check move to left edge cell');

		Select(0, 2, 0, 2, 0, 2);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftCell);
		CheckActiveCell(0, 1, 'Check move to left cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell);
		CheckActiveCell(0, 15, 'Check move to right edge cell');

		Select(4, 0, 4, 0, 4, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell, 1);
		CheckActiveCell(4, 8, 'Check move to right edge cell');

		Select(5, 0, 5, 0, 5, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightEdgeCell);
		CheckActiveCell(5, 5, 'Check move to right edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToRightCell);
		CheckActiveCell(5, 6, 'Check move to right cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToRightCell);
		CheckActiveCell(5, 7, 'Check move to right cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToDownCell);
		CheckActiveCell(6, 7, 'Check move to down cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToDownCell);
		CheckActiveCell(7, 7, 'Check move to down cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToUpCell);
		CheckActiveCell(6, 7, 'Check move to up cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToUpCell);
		CheckActiveCell(5, 7, 'Check move to up cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.moveToBottomCell);
		CheckActiveCell(25, 0, 'Check move to bottom cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToBottomCell);
		CheckActiveCell(27, 0, 'Check move to bottom cell');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.moveToFirstColumn);
		CheckActiveCell(5, 0, 'Check move to first column');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.moveToLeftEdgeTop);
		CheckActiveCell(0, 0, 'Check move to left-top edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToRightBottomEdge);
		CheckActiveCell(27, 15, 'Check move to right-bottom edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.moveToLowerCell);
		CheckActiveCell(77, 15, 'Check move to lower cell');

		ExecuteTableHotkey(tableHotkeyTypes.moveToUpperCell);
		CheckActiveCell(27, 15, 'Check move to upper cell');

		Select(35, 0, 35, 0, 35, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(27, 0, 35, 0, 'Check select to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(25, 0, 35, 0, 'Check select to top cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToTopCell);
		CheckActiveRange(0, 0, 35, 0, 'Check select to top cell');

		Select(0, 39, 0, 39, 0, 39);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeCell);
		CheckActiveRange(0, 15, 0, 39, 'Check select to left edge cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeCell);
		CheckActiveRange(0, 0, 0, 39, 'Check select to left edge cell');

		Select(0, 2, 0, 2, 0, 2);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftCell);
		CheckActiveRange(0, 1, 0, 2, 'Check select to left cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell);
		CheckActiveRange(0, 0, 0, 15, 'Check select to right edge cell');

		Select(4, 0, 4, 0, 4, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell, 1);
		CheckActiveRange(4, 0, 4, 8, 'Check select to right edge cell');

		Select(5, 0, 5, 0, 5, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightEdgeCell);
		CheckActiveRange(5, 0, 5, 5, 'Check select to right edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToRightCell);
		CheckActiveRange(5, 0, 5, 6, 'Check select to right cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToRightCell);
		CheckActiveRange(5, 0, 5, 7, 'Check select to right cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToDownCell);
		CheckActiveRange(5, 0, 6, 7, 'Check select to down cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToDownCell);
		CheckActiveRange(5, 0, 7, 7, 'Check select to down cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToUpCell);
		CheckActiveRange(5, 0, 6, 7, 'Check select to up cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToUpCell);
		CheckActiveRange(5, 0, 5, 7, 'Check select to up cell');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectToBottomCell);
		CheckActiveRange(0, 0, 25, 0, 'Check select to bottom cell');
		ExecuteTableHotkey(tableHotkeyTypes.selectToBottomCell);
		CheckActiveRange(0, 0, 27, 0, 'Check select to bottom cell');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.selectToFirstColumn);
		CheckActiveRange(5, 0, 5, 25, 'Check select to first column');

		Select(5, 25, 5, 25, 5, 25);
		ExecuteTableHotkey(tableHotkeyTypes.selectToLeftEdgeTop);
		CheckActiveRange(0, 0, 5, 25, 'Check select to left-top edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToRightBottomEdge);
		CheckActiveRange(5, 15, 27, 25, 'Check select to right-bottom edge cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToLowerCell);
		CheckActiveRange(5, 15, 77, 25, 'Check select to lower cell');

		ExecuteTableHotkey(tableHotkeyTypes.selectToUpperCell);
		CheckActiveRange(5, 15, 27, 25, 'Check select to upper cell');

		Select(0, 0, 0, 0, 3, 3);

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 1, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 2, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(0, 3, 'Check move active cell to right');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToRight);
		CheckActiveCell(1, 0, 'Check move active cell to right');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 3, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 2, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 1, 'Check move active cell to left');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToLeft);
		CheckActiveCell(0, 0, 'Check move active cell to left');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(1, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(2, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(3, 0, 'Check move active cell to down');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToDown);
		CheckActiveCell(0, 1, 'Check move active cell to down');

		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(3, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(2, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(1, 0, 'Check move active cell to up');
		ExecuteTableHotkey(tableHotkeyTypes.moveActiveCellToUp);
		CheckActiveCell(0, 0, 'Check move active cell to up');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectRow);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, 0, 'Check select row');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectColumn);
		CheckActiveRange(0, 0, 0, AscCommon.gc_nMaxCol0, 'Check select column');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteTableHotkey(tableHotkeyTypes.selectSheet);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0, 'Check select sheet');

		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll);
		CheckActiveRange(0, 0, AscCommon.gc_nMaxRow0, AscCommon.gc_nMaxCol0, 'Check select sheet');
	});

	QUnit.test('Check reset actions', (assert) =>
	{
		editor.asc_startAddShape('rect');
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(GetDrawingObjects().checkEndAddShape(), false, 'Check resetting add shape');
		Select(0, 0, 0, 0, 0, 0);
		editor.asc_SelectionCut();
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(wsView.getCutRange(), null, 'Check resetting cut range');

		Select(0, 0, 0, 0, 0, 0);
		editor.asc_Copy();
		ExecuteTableHotkey(tableHotkeyTypes.reset);
		assert.strictEqual(wsView.getCutRange(), null, 'Check resetting copy range');
	});


	QUnit.test('Check actions with removing', (assert) =>
	{
		Select(0, 15, 0, 15, 0, 15);
		FillActiveCell('Hello');
		Select(25, 0, 25, 0, 25, 0);
		FillActiveCell('Hello');
		Select(27, 0, 27, 0, 27, 0);
		FillActiveCell('Hello');

		Select(0, 15, 0, 0, 30, 30);
		ExecuteTableHotkey(tableHotkeyTypes.removeActiveCell);
		CloseCellEditor(true);

		assert.strictEqual(GetCellText(0, 15), '', 'Check remove active cell');
		assert.strictEqual(GetCellText(25, 0), 'Hello', 'Check remove active cell');
		assert.strictEqual(GetCellText(27, 0), 'Hello', 'Check remove active cell');

		ExecuteTableHotkey(tableHotkeyTypes.emptyRange);
		assert.strictEqual(GetCellText(0, 15), '', 'Check remove range');
		assert.strictEqual(GetCellText(25, 0), '', 'Check remove range');
		assert.strictEqual(GetCellText(27, 0), '', 'Check remove range');
	});

	QUnit.test('Check actions with filling cell', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getDateString(editor), 'Check add date');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), (new Asc.cDate()).getTimeString(editor), 'Check add time');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator);
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(0, 0), '.', 'Check add separator');
	});

	QUnit.test('Check actions with formatting cell', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell('0.1');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat);
		assert.strictEqual(GetCellText(0, 0), '1.00E-01', 'set exponential format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat);
		assert.strictEqual(GetCellText(0, 0), '10.00%', 'set percent format');

		FillActiveCell('49990');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat);
		assert.strictEqual(GetCellText(0, 0), '$49,990.00', 'set currency format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat);
		assert.strictEqual(GetCellText(0, 0), '12:00:00 AM', 'set time format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellDateFormat);
		assert.strictEqual(GetCellText(0, 0), '11/11/2036', 'set date format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat);
		assert.strictEqual(GetCellText(0, 0), '49990.00', 'set number format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat);
		assert.strictEqual(GetCellText(0, 0), '49990', 'set general format');

		function GetCellFormatting(row, column)
		{
			return ws.getRange4(row, column).getXfs();
		}

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.true(GetCellFormatting(0, 0).asc_getFontStrikeout(), 'Check turn on strikeout format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout);
		assert.false(GetCellFormatting(0, 0).asc_getFontStrikeout(), 'Check turn of strikeout format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.true(GetCellFormatting(0, 0).asc_getFontBold(), 'Check turn on bold format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold);
		assert.false(GetCellFormatting(0, 0).asc_getFontBold(), 'Check turn off bold format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.true(GetCellFormatting(0, 0).asc_getFontItalic(), 'Check turn on italic format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic);
		assert.false(GetCellFormatting(0, 0).asc_getFontItalic(), 'Check turn off italic format');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.true(GetCellFormatting(0, 0).asc_getFontUnderline(), 'Check turn on underline format');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline);
		assert.false(GetCellFormatting(0, 0).asc_getFontUnderline(), 'Check turn off underline format');

		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check init font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 12, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 14, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 16, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 18, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 20, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 22, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 24, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 26, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 28, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 36, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 48, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 72, "Check increase font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 72, "Check increase font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 48, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 36, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 28, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 26, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 24, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 22, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 20, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 18, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 16, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 14, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 12, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 10, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 9, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 8, "Check decrease font size");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 8, "Check decrease font size");

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
		assert.strictEqual(GetCellFormatting(0, 0).asc_getFontSize(), 11, "Check initial value");
	});

	QUnit.test('Check undo/redo', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('0.1');
		CloseCellEditor(true);

		assert.strictEqual(GetCellText(0, 0), '0.1', 'Check cell value');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo);
		assert.strictEqual(GetCellText(0, 0), '', 'Check undo');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo);
		assert.strictEqual(GetCellText(0, 0), '0.1', 'Check redo');
	});

	QUnit.test('Check focus on cell editor', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();

		ExecuteTableHotkey(tableHotkeyTypes.openCellEditor);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();

		ExecuteTableHotkey(tableHotkeyTypes.openCellEditor, 1);
		assert.true(wsView.getCellEditMode(), 'Check open cell editor');
		assert.strictEqual(cellEditor.textRender.getLinesCount(), 1, "Check lines count");
		CloseCellEditor();
	});
	QUnit.test('Check show formulas shortcut', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		FillActiveCell("=sum(1+2)");
		UpdateView();
		assert.strictEqual(GetCellCacheText(0, 0), "3", "Check show value");
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.ShowFormulas);
		assert.strictEqual(GetCellCacheText(0, 0), "=SUM(1+2)", "Check show formula");
	});

	QUnit.test('Check prevent default in Opera', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		assert.strictEqual(ExecuteTableHotkey(tableHotkeyTypes.disableNumLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
		assert.strictEqual(ExecuteTableHotkey(tableHotkeyTypes.disableScrollLock) & keydownresult_PreventDefault, keydownresult_PreventDefault);
	});

	QUnit.test('Check add sum formula', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('1');
		Select(0, 1, 0, 1, 0, 1);
		EnterText('2');

		Select(7, 7, 7, 7, 7, 7);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction);
		EnterText('A1,B1');
		CloseCellEditor(true);
		assert.strictEqual(GetCellText(7, 7), "3", 'Check sum result');
	});

	QUnit.test('Check move on worksheets', (assert) =>
	{
		CreateWorksheets(['name1', 'name2']);
		const ws1 = wb.getWorksheetByName(['name1'], 0);
		const ws2 = wb.getWorksheetByName(['name2'], 0);
		assert.strictEqual(wb.getActive(), 0, 'Check select first sheet');
		assert.true(wb.getWorksheet(0) === ws1, 'Check select first sheet');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet);
		assert.strictEqual(wb.getActive(), 1, 'Check select second sheet');
		assert.true(wb.getWorksheet(1) === ws2, 'Check select second sheet');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet);
		assert.strictEqual(wb.getActive(), 2, 'Check select third sheet');
		assert.true(wb.getWorksheet(2) === ws, 'Check select third sheet');

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet);
		assert.strictEqual(wb.getActive(), 1, 'Check select second sheet');
		assert.true(wb.getWorksheet(1) === ws2, 'Check select second sheet');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet);
		assert.strictEqual(wb.getActive(), 0, 'Check select first sheet');
		assert.true(wb.getWorksheet(0) === ws1, 'Check select first sheet');

		RemoveWorksheets([0, 1]);
	});

	QUnit.test('Check refresh connections', (assert) =>
	{
		Select(0, 0, 0, 0, 0, 0);
		EnterText('ad');
		Select(1, 0, 1, 0, 1, 0);
		EnterText('1');
		Select(2, 0, 2, 0, 2, 0);
		EnterText('2');
		CloseCellEditor(true);
		editor.asc_insertPivotExistingWorksheet("Sheet1!$A$1:$A$3", "Sheet1!$C$1");
		const pivotTable1 = ws.getPivotTable(2, 1);
		pivotTable1.asc_addField(editor, 0);

		Select(0, 1, 0, 1, 0, 1);
		EnterText('ap');
		Select(1, 1, 1, 1, 1, 1);
		EnterText('2');
		Select(2, 1, 2, 1, 2, 1);
		EnterText('3');
		CloseCellEditor(true);
		editor.asc_insertPivotExistingWorksheet("Sheet1!$B$1:$B$3", "Sheet1!$D$1");
		const pivotTable2 = ws.getPivotTable(3, 1);
		pivotTable2.asc_addField(editor, 0);

		Select(2, 0, 2, 0, 2, 0);
		EnterText('4');
		Select(2, 1, 2, 1, 2, 1);
		EnterText('4');

		Select(1, 2, 1, 2, 1, 2);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots);
		assert.strictEqual(GetCellText(1, 2), '5', 'Check refresh selected connections');

		Select(1, 3, 1, 3, 1, 3);
		assert.strictEqual(GetCellText(1, 3), '5', 'Check refresh selected connections');

		Select(2, 0, 2, 0, 2, 0);
		EnterText('5');

		Select(1, 2, 1, 2, 1, 2);
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots);
		assert.strictEqual(GetCellText(1, 2), '6', 'Check refresh all connections');

		Select(1, 3, 1, 3, 1, 3);
		assert.strictEqual(GetCellText(1, 3), '6', 'Check refresh all connections');

	});
	QUnit.test('Change format table info', (assert) =>
	{
		CreateTable();
		Select(1, 0, 1, 0, 1, 0);
		ExecuteTableHotkey(tableHotkeyTypes.changeFormatTableInfo);
		assert.strictEqual(GetCellText(3, 2), '6', 'Check change format table info');
		ExecuteTableHotkey(tableHotkeyTypes.changeFormatTableInfo);
	});

	QUnit.test('Check recalculating data', (assert) =>
	{
		CreateWorksheets(['name1']);
		const wsName = wb.getWorksheetByName(['name1'], 0);

		ws.getRange4(0, 2).fillData([['=sum(a1+b1)']]);
		ws.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(1);
		});
		ws.getRange4(0, 1)._foreach(function (cell)
		{
			cell.setValueNumberInternal(2);
		});

		wsName.getRange4(0, 2).fillData([['=sum(a1+b1)']]);
		wsName.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(1);
		});
		wsName.getRange4(0, 1)._foreach(function (cell)
		{
			cell.setValueNumberInternal(2);
		});

		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating values');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating values');
		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet);
		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '0', 'Check non calculating value');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '3', 'Check calculating value');

		wsName.getRange4(0, 0)._foreach(function (cell)
		{
			cell.setValueNumberInternal(0);
		});

		ExecuteShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateAll);
		assert.strictEqual(ws.getRange4(0, 2).getValueWithFormat(), '3', 'Check calculating value');
		assert.strictEqual(wsName.getRange4(0, 2).getValueWithFormat(), '2', 'Check calculating value');
		RemoveWorksheets([0]);
	});
})(window);
