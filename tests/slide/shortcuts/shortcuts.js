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

'use strict';

(function (window)
{
	const {
		mainShortcutTypes,
		thumbnailsTypes,
		demonstrationTypes,
		demonstrationEvents,
		ExecuteDemonstrationShortcut,
		ExecuteThumbnailShortcut,
		ExecuteMainShortcut,
		ExecuteThumbnailHotkey,
		ExecuteMainHotkey,
		CNativeEvent
	} = AscTest;

	const logicDocument = AscTest.CreateLogicDocument();
	AscFormat.CHART_STYLE_MANAGER.init();

	editor.sendEvent = Asc.asc_docs_api.prototype.sendEvent.bind(editor);
	editor.asc_registerCallback = Asc.asc_docs_api.prototype.asc_registerCallback.bind(editor);
	editor.asc_unregisterCallback = Asc.asc_docs_api.prototype.asc_unregisterCallback.bind(editor);
	editor.sync_ContextMenuCallback = Asc.asc_docs_api.prototype.sync_ContextMenuCallback.bind(editor);
	editor.sync_DemonstrationSlideChanged = Asc.asc_docs_api.prototype.sync_DemonstrationSlideChanged.bind(editor);
	editor.sync_endDemonstration = Asc.asc_docs_api.prototype.sync_endDemonstration.bind(editor);
	editor.sync_DialogAddHyperlink = Asc.asc_docs_api.prototype.sync_DialogAddHyperlink.bind(editor);
	editor.getGraphicController = Asc.asc_docs_api.prototype.getGraphicController.bind(editor);
	editor.FontSizeIn = Asc.asc_docs_api.prototype.FontSizeIn.bind(editor);
	editor.FontSizeOut = Asc.asc_docs_api.prototype.FontSizeOut.bind(editor);
	editor.asc_getChartObject = Asc.asc_docs_api.prototype.asc_getChartObject.bind(editor);
	editor.asc_AddMath2 = Asc.asc_docs_api.prototype.asc_AddMath2.bind(editor);
	editor.IncreaseIndent = Asc.asc_docs_api.prototype.IncreaseIndent.bind(editor);
	editor.DecreaseIndent = Asc.asc_docs_api.prototype.DecreaseIndent.bind(editor);
	editor.get_ShowParaMarks = Asc.asc_docs_api.prototype.get_ShowParaMarks.bind(editor);
	editor.retrieveFormatPainterData = Asc.asc_docs_api.prototype.retrieveFormatPainterData.bind(editor);
	editor.sync_EndAddShape = Asc.asc_docs_api.prototype.sync_EndAddShape.bind(editor);
	editor.asc_getShowGuides = Asc.asc_docs_api.prototype.asc_getShowGuides.bind(editor);
	editor.sync_HyperlinkClickCallback = Asc.asc_docs_api.prototype.sync_HyperlinkClickCallback.bind(editor);
	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		window.setTimeout(function ()
		{
			callback({"saveLock": false});
		}, 0);
	};
	function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	function GoToSlide(slide)
	{
		editor.WordControl.GoToPage(slide);
	}

	function CleanPresentation()
	{
		editor.WordControl.Thumbnails.SelectAll();
		logicDocument.deleteSlides(logicDocument.GetSelectedSlides());
	}

	function AddShape(x, y, height, width)
	{
		AscCommon.History.Create_NewPoint();
		const shapeTrack = new AscFormat.NewShapeTrack('rect', x, y, AscFormat.GetDefaultTheme(), null, null, null, 0);
		shapeTrack.track({}, x+ width, y + height);
		const shape = shapeTrack.getShape(false, AscTest.DrawingDocument, null);
		shape.setBDeleted(false);
		shape.setParent(logicDocument.Slides[0]);
		shape.addToDrawingObjects();
		shape.select(GetDrawingObjects(), 0);
		return shape;
	}

	function AddChart()
	{
		const chart = editor.asc_getChartObject(Asc.c_oAscChartTypeSettings.lineNormal);
		chart.setParent(logicDocument.Slides[0]);

		chart.addToDrawingObjects();
		chart.spPr.setXfrm(new AscFormat.CXfrm());
		chart.spPr.xfrm.setOffX(0);
		chart.spPr.xfrm.setOffY(0);
		chart.spPr.xfrm.setExtX(100);
		chart.spPr.xfrm.setExtY(100);
		logicDocument.Recalculate();
		return chart;
	}

	function AddPlaceholderShape()
	{
		const shape = AddShape(0, 0, 100, 100);
		shape.setNvSpPr(new AscFormat.UniNvPr());
		let placeholder = new AscFormat.Ph();
		placeholder.setType(AscFormat.phType_title);
		shape.nvSpPr.nvPr.setPh(placeholder);
		shape.txBody = AscFormat.CreateTextBodyFromString('', shape.getDrawingDocument(), shape);

		shape.recalculateContentWitCompiledPr();
		return shape;
	}

	function AddTable(columns, rows)
	{
		return logicDocument.Add_FlowTable(columns, rows);
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

	function ClearShapeAndAddParagraph(sText)
	{
		const textShape = AddShape(0, 0, 100, 100);
		const txBody = AscFormat.CreateTextBodyFromString(sText, editor.WordControl.m_oDrawingDocument, textShape)
		textShape.setTxBody(txBody);
		textShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
		const content = txBody.content;
		content.SetThisElementCurrent();
		content.MoveCursorToStartPos();
		textShape.recalculate();
		return {shape: textShape, paragraph: content.Content[0]};
	}

	function GetDirectTextPr()
	{
		return logicDocument.GetDirectTextPr();
	}

	function GetDirectParaPr()
	{
		return logicDocument.GetDirectParaPr();
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

	QUnit.module("Test shortcuts",
		{
			afterEach : function ()
			{
				CleanPresentation();
			},
			beforeEach: function ()
			{
				CreateSlide();
			},
			before: function ()
			{
				AscTest.TurnOffRecalculate();
			},
			after: function ()
			{
				AscTest.TurnOnRecalculate();
			}
		});
	QUnit.test('Check actions with slides', (assert) =>
	{
		AscTest.TurnOnRecalculate();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();

		function CheckSelectedSlides(arrSelectedSlides)
		{
			const presentationSelectedSlides = logicDocument.GetSelectedSlides();
			assert.deepEqual(presentationSelectedSlides, arrSelectedSlides, 'Check selected slides ' + arrSelectedSlides.join(', '));
		}

		GoToSlide(0);

		let oldSlides = logicDocument.Slides.slice();
		ExecuteThumbnailHotkey(thumbnailsTypes.addNextSlide, 0);
		CheckSelectedSlides([1]);
		assert.strictEqual(oldSlides.indexOf(logicDocument.Slides[1]), -1, 'Check add next slide');

		oldSlides = logicDocument.Slides.slice();
		ExecuteThumbnailHotkey(thumbnailsTypes.addNextSlide, 1);
		CheckSelectedSlides([2]);
		assert.strictEqual(oldSlides.indexOf(logicDocument.Slides[2]), -1, 'Check add next slide');

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToNextSlide, 0);
		CheckSelectedSlides([3]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToNextSlide, 1);
		CheckSelectedSlides([4]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToPreviousSlide, 0);
		CheckSelectedSlides([3]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToPreviousSlide, 1);
		CheckSelectedSlides([2]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToLastSlide);
		CheckSelectedSlides([6]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToFirstSlide);
		CheckSelectedSlides([0]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToLastSlide);
		CheckSelectedSlides([6]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveToFirstSlide);
		CheckSelectedSlides([0]);


		ExecuteThumbnailHotkey(thumbnailsTypes.selectNextSlide, 0);
		CheckSelectedSlides([0, 1]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectNextSlide, 1);
		CheckSelectedSlides([0, 1, 2]);

		GoToSlide(2);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectPreviousSlide, 0);
		CheckSelectedSlides([1, 2]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectPreviousSlide, 0);
		CheckSelectedSlides([0, 1, 2]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectToLastSlide);
		CheckSelectedSlides([2, 3, 4, 5, 6]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectPreviousSlide, 1);
		CheckSelectedSlides([2, 3, 4, 5]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectPreviousSlide, 1);
		CheckSelectedSlides([2, 3, 4]);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectToLastSlide);
		CheckSelectedSlides([2, 3, 4, 5, 6]);

		GoToSlide(3);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectToFirstSlide);
		CheckSelectedSlides([0, 1, 2, 3]);

		GoToSlide(3);

		ExecuteThumbnailHotkey(thumbnailsTypes.selectToFirstSlide);
		CheckSelectedSlides([0, 1, 2, 3]);

		let startPositionSlides = logicDocument.Slides.slice();

		function CheckMoveSlides(oldPositions, newPositions)
		{
			for (let i = 0; i < oldPositions.length; i += 1)
			{
				assert.true(startPositionSlides[oldPositions[i]] === logicDocument.Slides[newPositions[i]], 'Check position ' + oldPositions[i] + '->' + newPositions[i]);
			}
			startPositionSlides = logicDocument.Slides.slice();
		}

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToNextPosition, 0);
		CheckSelectedSlides([1, 2, 3, 4]);
		CheckMoveSlides([0, 1, 2, 3], [1, 2, 3, 4]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToNextPosition, 1);
		CheckSelectedSlides([2, 3, 4, 5]);
		CheckMoveSlides([1, 2, 3, 4], [2, 3, 4, 5]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToStart, 0);
		CheckSelectedSlides([0, 1, 2, 3]);
		CheckMoveSlides([2, 3, 4, 5], [0, 1, 2, 3]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToEnd, 0);
		CheckSelectedSlides([3, 4, 5, 6]);
		CheckMoveSlides([0, 1, 2, 3], [3, 4, 5, 6]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToStart, 1);
		CheckSelectedSlides([0, 1, 2, 3]);
		CheckMoveSlides([3, 4, 5, 6], [0, 1, 2, 3]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToEnd, 1);
		CheckSelectedSlides([3, 4, 5, 6]);
		CheckMoveSlides([0, 1, 2, 3], [3, 4, 5, 6]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToPreviousPosition, 0);
		CheckSelectedSlides([2, 3, 4, 5]);
		CheckMoveSlides([3, 4, 5, 6], [2, 3, 4, 5]);

		ExecuteThumbnailHotkey(thumbnailsTypes.moveSelectedSlidesToPreviousPosition, 1);
		CheckSelectedSlides([1, 2, 3, 4]);
		CheckMoveSlides([2, 3, 4, 5], [1, 2, 3, 4]);

		ExecuteThumbnailHotkey(thumbnailsTypes.removeSelectedSlides, 0);
		assert.strictEqual(logicDocument.Slides.length, 3, 'Check slide amount');
		assert.true(logicDocument.Slides[0] === startPositionSlides[0], 'Check saved slides');
		assert.true(logicDocument.Slides[1] === startPositionSlides[5], 'Check saved slides');
		assert.true(logicDocument.Slides[2] === startPositionSlides[6], 'Check saved slides');

		ExecuteThumbnailHotkey(thumbnailsTypes.removeSelectedSlides, 1);
		assert.strictEqual(logicDocument.Slides.length, 2, 'Check slide amount');
		assert.true(logicDocument.Slides[0] === startPositionSlides[0], 'Check saved slides');
		assert.true(logicDocument.Slides[1] === startPositionSlides[6], 'Check saved slides');

		ExecuteThumbnailHotkey(thumbnailsTypes.removeSelectedSlides, 1);
		assert.strictEqual(logicDocument.Slides.length, 1, 'Check slide amount');
		assert.true(logicDocument.Slides[0] === startPositionSlides[0], 'Check saved slides');

		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();

		GoToSlide(0);

		oldSlides = logicDocument.Slides.slice();
		ExecuteMainHotkey(mainShortcutTypes.addNextSlide);
		CheckSelectedSlides([1]);
		assert.strictEqual(oldSlides.indexOf(logicDocument.Slides[1]), -1, 'Check add next slide');

		ExecuteThumbnailShortcut(Asc.c_oAscPresentationShortcutType.EditSelectAll);
		CheckSelectedSlides([0, 1, 2, 3, 4, 5, 6]);

		GoToSlide(1);

		ExecuteMainHotkey(mainShortcutTypes.moveToNextSlide, 0);
		CheckSelectedSlides([2]);

		ExecuteMainHotkey(mainShortcutTypes.moveToPreviousSlide, 0);
		CheckSelectedSlides([1]);

		ExecuteMainHotkey(mainShortcutTypes.moveToNextSlide, 1);
		CheckSelectedSlides([2]);

		ExecuteMainHotkey(mainShortcutTypes.moveToPreviousSlide, 1);
		CheckSelectedSlides([1]);

		ExecuteMainHotkey(mainShortcutTypes.moveToNextSlide, 2);
		CheckSelectedSlides([2]);

		ExecuteMainHotkey(mainShortcutTypes.moveToPreviousSlide, 2);
		CheckSelectedSlides([1]);

		ExecuteMainHotkey(mainShortcutTypes.moveToLastSlide);
		CheckSelectedSlides([6]);

		ExecuteMainHotkey(mainShortcutTypes.moveToFirstSlide);
		CheckSelectedSlides([0]);

		ExecuteMainHotkey(mainShortcutTypes.selectToLastSlide);
		CheckSelectedSlides([0, 1, 2, 3, 4, 5, 6]);

		GoToSlide(3);

		ExecuteMainHotkey(mainShortcutTypes.selectToFirstSlide);
		CheckSelectedSlides([0, 1, 2, 3]);
		AscTest.TurnOffRecalculate();
	});

	QUnit.test('Check actions with catch events', (oAssert) =>
	{
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();
		CreateSlide();

		let Execute;

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

			Execute(oEvent);
			oAssert.strictEqual(check, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
			editor.asc_unregisterCallback(sSendEvent, Check);
		}

		editor.WordControl.DemonstrationManager.Start(null, 0);
		Execute = ExecuteDemonstrationShortcut;
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 1, demonstrationEvents[demonstrationTypes.moveToNextSlide][0]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, demonstrationEvents[demonstrationTypes.moveToNextSlide][1]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, demonstrationEvents[demonstrationTypes.moveToNextSlide][2]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, demonstrationEvents[demonstrationTypes.moveToNextSlide][3]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 5, demonstrationEvents[demonstrationTypes.moveToNextSlide][4]);

		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, demonstrationEvents[demonstrationTypes.moveToPreviousSlide][0]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, demonstrationEvents[demonstrationTypes.moveToPreviousSlide][1]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, demonstrationEvents[demonstrationTypes.moveToPreviousSlide][2]);

		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 0, demonstrationEvents[demonstrationTypes.moveToFirstSlide][0]);

		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 11, demonstrationEvents[demonstrationTypes.moveToLastSlide][0]);

		ExecuteDemonstrationShortcut(new CNativeEvent(53));
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, demonstrationEvents[demonstrationTypes.moveToNumberSlide][0]);
		ExecuteDemonstrationShortcut(new CNativeEvent(56));
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 7, demonstrationEvents[demonstrationTypes.moveToNumberSlide][0]);
		ExecuteDemonstrationShortcut(new CNativeEvent(49));
		ExecuteDemonstrationShortcut(new CNativeEvent(48));
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 9, demonstrationEvents[demonstrationTypes.moveToNumberSlide][0]);
		ExecuteTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 10, demonstrationEvents[demonstrationTypes.moveToNumberSlide][0]);

		
		ExecuteTestWithCatchEvent('asc_onEndDemonstration', () => true, true, demonstrationEvents[demonstrationTypes.exitFromDemonstrationMode][0]);

		Execute = ExecuteMainShortcut;
		ExecuteTestWithCatchEvent('asc_onPrint', () => true, true, Asc.c_oAscPresentationShortcutType.Print);
		ExecuteTestWithCatchEvent('asc_onContextMenu', () => true, true, Asc.c_oAscPresentationShortcutType.ShowContextMenu);

		ClearShapeAndAddParagraph('')
		ExecuteTestWithCatchEvent('asc_onDialogAddHyperlink', () => true, true, Asc.c_oAscPresentationShortcutType.AddHyperlink);

		Execute = ExecuteThumbnailShortcut;
		GoToSlide(0, FOCUS_OBJECT_THUMBNAILS);
		ExecuteTestWithCatchEvent('asc_onPrint', () => true, true, Asc.c_oAscPresentationShortcutType.Print);
		ExecuteTestWithCatchEvent('asc_onContextMenu', () => true, true, Asc.c_oAscPresentationShortcutType.ShowContextMenu);

		GoToSlide(0, FOCUS_OBJECT_MAIN);
	});

	QUnit.test('Check add various characters', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteMainHotkey(mainShortcutTypes.checkNonBreakingSpace);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0), 'Check add non breaking space');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.EuroSign);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x20AC), 'Check add euro sign');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.EnDash);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x20AC, 0x2013), 'Check add en dash');

		ExecuteMainHotkey(mainShortcutTypes.checkAddSpace);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x20AC, 0x2013, 0x0020), 'Check add space');
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

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToEndLineContent);
		CheckCursorPosition(18, 'Check move to end line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorRight);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToEndLineContent, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToEndLineContent, 1);
		CheckCursorPosition(36, 'Check move to end line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToStartLineContent, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToStartLineContent, 1);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorRight);
		CheckCursorPosition(19, 'Check move to right char');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorLeft);
		CheckCursorPosition(18, 'Check move to left char');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorWordLeft);
		CheckCursorPosition(12, 'Check move to left word');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorWordLeft, 1);
		CheckCursorPosition(6, 'Check move to left word');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorWordRight, 1);
		CheckCursorPosition(12, 'Check move to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorWordRight);
		CheckCursorPosition(18, 'Check move to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorWordRight);
		CheckCursorPosition(24, 'Check move to right word');


		ExecuteMainHotkey(mainShortcutTypes.checkMoveToStartLineContent);
		CheckCursorPosition(18, 'Check move to start line');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorBottom);
		CheckCursorPosition(36, 'Check move down');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorTop);
		CheckCursorPosition(18, 'Check move up');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToEndPosContent);
		CheckCursorPosition(161, 'Check move to end document');

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToStartPosContent);
		CheckCursorPosition(0, 'Check move to start document');

		function CheckSelectedText(expectedText, description)
		{
			const selectedText = logicDocument.GetSelectedText();
			assert.strictEqual(selectedText, expectedText, description);
		}

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToEndLineContent);
		CheckSelectedText('Hello World Hello ', 'Select to end line');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorRight);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToEndLineContent, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToEndLineContent, 1);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToStartLineContent, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToStartLineContent, 1);
		CheckSelectedText('Hello World Hello ', 'Select to start line content');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorRight);
		CheckSelectedText('Hello World Hello W', 'Select to right char');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorLeft);
		CheckSelectedText('Hello World Hello ', 'Select to left char');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorLeft);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorRight);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorLeft, 1);
		CheckSelectedText('Hello World ', 'Select to left word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorRight, 1);
		CheckSelectedText('Hello World Hello ', 'Select to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorRight);
		CheckSelectedText('Hello World Hello World ', 'Select to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorRight);
		CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectToStartLineContent);
		CheckSelectedText('Hello World Hello ', 'Select to start line');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorBottom);
		CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorTop);
		CheckSelectedText('Hello World Hello ', 'Select up');


		logicDocument.MoveCursorToEndPos();
		ExecuteMainHotkey(mainShortcutTypes.checkSelectCursorLeft);
		CheckSelectedText('d', 'Select to left char');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectWordCursorLeft);
		CheckSelectedText('World', 'Select to left word');
	});

	QUnit.test('Check remove parts of text', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('Hello Hello Hello Hello Hello Hello Hello');
		logicDocument.MoveCursorToEndPos();

		ExecuteMainHotkey(mainShortcutTypes.checkDeleteBack);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

		ExecuteMainHotkey(mainShortcutTypes.checkDeleteWordBack);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

		ExecuteMainHotkey(mainShortcutTypes.checkDeleteWordBack, 1);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

		logicDocument.MoveCursorToStartPos();
		ExecuteMainHotkey(mainShortcutTypes.checkDeleteFront);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
		ExecuteMainHotkey(mainShortcutTypes.checkDeleteWordFront);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello ', 'Check removing front word');

		ExecuteMainHotkey(mainShortcutTypes.checkDeleteWordFront, 1);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello ', 'Check removing front word');
	});

	QUnit.test('Check text property change', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello world');
		logicDocument.SelectAll();

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Bold);
		assert.strictEqual(GetDirectTextPr().GetBold(), true, 'Check turn on bold');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Bold);
		assert.strictEqual(GetDirectTextPr().GetBold(), false, 'Check turn off bold');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Italic);
		assert.strictEqual(GetDirectTextPr().GetItalic(), true, 'Check turn on italic');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Italic);
		assert.strictEqual(GetDirectTextPr().GetItalic(), false, 'Check turn off italic');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Strikethrough);
		assert.strictEqual(GetDirectTextPr().GetStrikeout(), true, 'Check turn on strikeout');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Strikethrough);
		assert.strictEqual(GetDirectTextPr().GetStrikeout(), false, 'Check turn off strikeout');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Underline);
		assert.strictEqual(GetDirectTextPr().GetUnderline(), true, 'Check turn on underline');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Underline);
		assert.strictEqual(GetDirectTextPr().GetUnderline(), false, 'Check turn off underline');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Superscript);
		assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_SuperScript, 'Check turn on superscript');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Superscript);
		assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off superscript');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Subscript);
		assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_SubScript, 'Check turn on subscript');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Subscript);
		assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off subscript');

		// defaultSize = 10
		// 10 -> 11 -> 12 -> 14 -> 16 -> 14 -> 12 -> 11 -> 10
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 11, 'Check increase font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 12, 'Check increase font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 14, 'Check increase font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 16, 'Check increase font size');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 14, 'Check decrease font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 12, 'Check decrease font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 11, 'Check decrease font size');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFont);
		assert.strictEqual(GetDirectTextPr().GetFontSize(), 10, 'Check decrease font size');
	});

	QUnit.test('Check paragraph property change', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('Hello world');

		assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check align left");
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.CenterAlign);
		assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Center, "Check turn on center para");

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.JustifyAlign);
		assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.LeftAlign);
		assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.RightAlign);
		assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Right, "Check turn on right para");


		paragraph.Set_PresentationLevel(0);
		paragraph.Pr.SetInd(0, 0, 0);
		ExecuteMainHotkey(mainShortcutTypes.checkIncreaseIndent);
		assert.strictEqual(GetDirectParaPr().GetIndLeft(), 11.1125, "Check indent");

		ExecuteMainHotkey(mainShortcutTypes.checkDecreaseIndent);
		assert.strictEqual(GetDirectParaPr().GetIndLeft(), 0, "Check unindent");

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.BulletList);
		const bullet = paragraph.Get_PresentationNumbering();
		assert.strictEqual(bullet.m_nType, AscFormat.numbering_presentationnumfrmt_Char, 'Check bullet list shortcut');

		paragraph.Pr.SetInd(0, 0, 0);
		ExecuteMainHotkey(mainShortcutTypes.checkIncreaseBulletIndent);
		assert.strictEqual(paragraph.Pr.Get_IndLeft(), 11.1125, 'Check bullet indent shortcut');

		ExecuteMainHotkey(mainShortcutTypes.checkDecreaseBulletIndent);
		assert.strictEqual(paragraph.Pr.Get_IndLeft(), 0, 'Check bullet unindent shortcut');
	});

	QUnit.test('Check main actions with shapes', (assert) =>
	{
		const drawing1 = AddShape(0, 0, 100, 200);

		const dotsPerMM = logicDocument.DrawingDocument.GetDotsPerMM();

		function CheckShapePosition(X, Y)
		{
			assert.deepEqual([Round(drawing1.x * dotsPerMM, 10), Round(drawing1.y * dotsPerMM, 10), drawing1.extX, drawing1.extY], [X, Y, 200, 100], 'Check shape position after movement ' + X + ' ' + Y);
		}

		SelectDrawings([drawing1]);

		AscTest.TurnOnRecalculate();
		ExecuteMainHotkey(mainShortcutTypes.checkMoveShapeLeft);
		CheckShapePosition(-5, 0);

		ExecuteMainHotkey(mainShortcutTypes.checkLittleMoveShapeLeft);
		CheckShapePosition(-6, 0);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveShapeRight);
		CheckShapePosition(-1, 0);

		ExecuteMainHotkey(mainShortcutTypes.checkLittleMoveShapeRight);
		CheckShapePosition(0, 0);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveShapeBottom);
		CheckShapePosition(0, 5);

		ExecuteMainHotkey(mainShortcutTypes.checkLittleMoveShapeBottom);
		CheckShapePosition(0, 6);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveShapeTop);
		CheckShapePosition(0, 1);

		ExecuteMainHotkey(mainShortcutTypes.checkLittleMoveShapeTop);
		CheckShapePosition(0, 0);
		AscTest.TurnOffRecalculate();


		function CheckSelectedObjects(arrOfDrawings)
		{
			const length = Math.max(arrOfDrawings.length, GetDrawingObjects().selectedObjects.length);
			for (let i = 0; i < length; i++)
			{
				assert.true(GetDrawingObjects().selectedObjects[i] === arrOfDrawings[i], 'Check selection movement between objects');
			}
		}

		logicDocument.RemoveSelection();
		const drawing2 = AddShape(0, 0, 10, 10);
		const drawing3 = AddShape(0, 0, 10, 10);
		SelectDrawings([drawing3]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectNextObject);
		CheckSelectedObjects([drawing1]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectNextObject);
		CheckSelectedObjects([drawing2]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectNextObject);
		CheckSelectedObjects([drawing3]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectPreviousObject);
		CheckSelectedObjects([drawing2]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectPreviousObject);
		CheckSelectedObjects([drawing1]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectPreviousObject);
		CheckSelectedObjects([drawing3]);

		drawing2.setWordShape(false);
		SelectDrawings([drawing2]);

		ExecuteMainHotkey(mainShortcutTypes.checkAddTxBodyShape);
		assert.true(!!drawing2.txBody, 'Check create textBody after enter');

		SelectDrawings([drawing2]);
		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorToStartPosShape);
		const paragraph = drawing2.getDocContent().Content[0];
		assert.true(paragraph.IsThisElementCurrent() && paragraph.IsCursorAtBegin(), 'Check movement to start position in empty content');

		AscTest.EnterText([72, 101, 108, 108, 111]);
		SelectDrawings([drawing2]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectAllContentShape);
		assert.strictEqual(logicDocument.GetSelectedText(), 'Hello', 'Check select non empty content');

		SelectDrawings([drawing1, drawing2]);

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Group);
		assert.strictEqual(GetDrawingObjects().selectedObjects[0].getObjectType(), AscDFH.historyitem_type_GroupShape, 'Check group selected objects');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.UnGroup);
		assert.strictEqual(GetDrawingObjects().selectedObjects[0].getObjectType(), AscDFH.historyitem_type_Shape, 'Check ungroup selected group');
		assert.strictEqual(GetDrawingObjects().selectedObjects[1].getObjectType(), AscDFH.historyitem_type_Shape, 'Check ungroup selected group');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Group);
		let group = GetDrawingObjects().selectedObjects[0];
		group.selectObject(drawing1, 0);
		GetDrawingObjects().selection.groupSelection = group;

		ExecuteMainHotkey(mainShortcutTypes.checkResetStepDrawingSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects[0].getObjectType(), AscDFH.historyitem_type_GroupShape, 'Check reset step drawing selection');
		assert.strictEqual(group.selectedObjects.length, 0);

		group.selectObject(drawing1, 0);
		GetDrawingObjects().selection.groupSelection = group;

		ExecuteMainHotkey(mainShortcutTypes.checkResetAllDrawingSelection);
		assert.strictEqual(GetDrawingObjects().selectedObjects.length, 0, 'Check reset all drawing selection');

		const graphicFrame = AddTable(4, 4);

		function CheckTablePosition(row, column)
		{
			assert.strictEqual(graphicFrame.graphicObject.CurCell.Index, column, 'Check column active cell position');
			assert.strictEqual(graphicFrame.graphicObject.CurCell.Row.Index, row, 'Check row active cell position');
		}

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorRightTable);
		CheckTablePosition(0, 1);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorRightTable);
		CheckTablePosition(0, 2);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorRightTable);
		CheckTablePosition(0, 3);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorLeftTable);
		CheckTablePosition(0, 2);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorLeftTable);
		CheckTablePosition(0, 1);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorBottomTable);
		CheckTablePosition(1, 1);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveCursorTopTable);
		CheckTablePosition(0, 1);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToNextCell);
		CheckTablePosition(0, 2);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToNextCell);
		CheckTablePosition(0, 3);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToPreviousCell);
		CheckTablePosition(0, 2);

		ExecuteMainHotkey(mainShortcutTypes.checkMoveToPreviousCell);
		CheckTablePosition(0, 1);

		const cell = graphicFrame.graphicObject.Content[0].Get_Cell(0);
		const content = cell.GetContent();
		AscFormat.AddToContentFromString(content, 'Hello Hello');
		SelectDrawings([graphicFrame]);
		ExecuteMainHotkey(mainShortcutTypes.checkSelectFirstCellContent);
		CheckTablePosition(0, 0);
		assert.strictEqual(logicDocument.GetSelectedText(), 'Hello Hello', 'Check select first cell content');

		graphicFrame.MoveCursorToStartPos();
		logicDocument.MoveCursorRight(true, true);
		logicDocument.MoveCursorRight(true, true);

		logicDocument.MoveCursorRight(true, true);

		logicDocument.MoveCursorRight(true, true);
		ExecuteMainHotkey(mainShortcutTypes.checkRemoveAndMoveToStartPosTable);
		assert.true(graphicFrame.graphicObject.IsCursorAtBegin(), 'Check removing in table');
		assert.strictEqual(logicDocument.GetSelectedText(), '', 'Check removing in table');
	});

	QUnit.test('Check prevent default', (assert) =>
	{
		let res = ExecuteMainHotkey(mainShortcutTypes.checkNumLock);
		assert.strictEqual(res & keydownresult_PreventDefault, keydownresult_PreventDefault, 'Check prevent default');
		res = ExecuteMainHotkey(mainShortcutTypes.checkScrollLock);
		assert.strictEqual(res & keydownresult_PreventDefault, keydownresult_PreventDefault, 'Check prevent default');
		res = ExecuteMainHotkey(mainShortcutTypes.disableBrowserZoomIn);
		assert.strictEqual(res & keydownresult_PreventDefault, keydownresult_PreventDefault, 'Check prevent default');
	});

	QUnit.test('Check remove graphic objects', (assert) =>
	{

		function CheckRemoveObject(spTree, shape)
		{
			assert.true(spTree.indexOf(shape) === -1, 'Check remove shape from spTree');
		}

		const checkActionsWithEvent = (eventIndex) =>
		{
			const shape = AddShape(0, 0, 100, 100);
			SelectDrawings([shape]);
			logicDocument.AddAnimation(1, 1, 0, null, false, false);
			GetDrawingObjects().resetSelection();
			const timing = logicDocument.GetCurTiming();
			let effects = timing.getObjectEffects(shape.GetId());
			effects[0].select();
			ExecuteMainHotkey(mainShortcutTypes.checkRemoveAnimation, eventIndex)
			effects = timing.getObjectEffects(shape.GetId());
			assert.strictEqual(effects.length, 0, 'Check remove animation');

			SelectDrawings([shape]);
			assert.true(logicDocument.GetCurrentSlide().cSld.spTree.indexOf(shape) !== -1, 'Check available shape in spTree');
			ExecuteMainHotkey(mainShortcutTypes.checkRemoveShape, eventIndex);
			CheckRemoveObject(logicDocument.GetCurrentSlide().cSld.spTree, shape);

			const chart = AddChart();
			SelectDrawings([chart]);
			ExecuteMainHotkey(mainShortcutTypes.checkRemoveChart, eventIndex);
			CheckRemoveObject(logicDocument.GetCurrentSlide().cSld.spTree, chart);


			const shape1 = AddShape(0, 0, 100, 100);
			const shape2 = AddShape(0, 0, 100, 100);
			const shape3 = AddShape(0, 0, 100, 100);

			const group = CreateGroup([shape1, shape2, shape3]);
			group.selectObject(shape1, 0);
			GetDrawingObjects().selection.groupSelection = group;

			ExecuteMainHotkey(mainShortcutTypes.checkRemoveShapeInGroup, eventIndex);
			CheckRemoveObject(group.spTree, shape1);

			SelectDrawings([group]);
			ExecuteMainHotkey(mainShortcutTypes.checkRemoveGroup, eventIndex);
			CheckRemoveObject(logicDocument.GetCurrentSlide().cSld.spTree, group);
		};

		checkActionsWithEvent(0);
		checkActionsWithEvent(1);
	});

	QUnit.test('Check select all in chart title', (assert) =>
	{
		AscTest.TurnOnRecalculate();
		const chart = AddChart();
		SelectDrawings([chart]);
		const titles = chart.getAllTitles();
		const controller = GetDrawingObjects();
		controller.selection.chartSelection = chart;
		chart.selectTitle(titles[0], 0);
		ExecuteMainHotkey(mainShortcutTypes.checkSelectAllContentChartTitle);
		assert.strictEqual(logicDocument.GetSelectedText(), 'Diagram Title', 'Check select all content in chart title');
		logicDocument.Remove();
		AscTest.TurnOffRecalculate();
	});

	QUnit.test('Check actions for objects with placeholder', (assert) =>
	{
		const shape1 = AddPlaceholderShape();
		const shape2 = AddPlaceholderShape();
		const shape3 = AddPlaceholderShape();

		SelectDrawings([shape1]);

		ExecuteMainHotkey(mainShortcutTypes.checkSelectNextObjectWithPlaceholder);
		assert.true(GetDrawingObjects().selectedObjects[0] === shape2, 'Check select next object with placeholder');

		ExecuteMainHotkey(mainShortcutTypes.checkSelectNextObjectWithPlaceholder);
		assert.true(GetDrawingObjects().selectedObjects[0] === shape3, 'Check select next object with placeholder');

		ExecuteMainHotkey(mainShortcutTypes.checkAddNextSlideAfterSelectLastPlaceholderObject);

		assert.strictEqual(logicDocument.Slides.length, 2, 'Check creating new slide after selecting last shape with placeholder');
		assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check creating new slide after selecting last shape with placeholder');
	});

	QUnit.test('Check add break line', (assert) =>
	{
		function CheckAddBreakLine(shape)
		{
			assert.strictEqual(shape.txBody.content.Content.length, 1, 'Check amount paragraphs');
			assert.strictEqual(shape.txBody.content.Content[0].GetLinesCount(), 2, 'Check amount lines');
		}

		{
			AscTest.TurnOnRecalculate();
			const {shape} = ClearShapeAndAddParagraph('Hello Hello');
			logicDocument.MoveCursorRight(false, true);

			ExecuteMainHotkey(mainShortcutTypes.checkAddBreakLine);
			CheckAddBreakLine(shape);

			const placeholderShape = AddPlaceholderShape();
			placeholderShape.setTxBody(AscFormat.CreateTextBodyFromString('Hello hello', AscTest.DrawingDocument, placeholderShape));
			placeholderShape.txBody.content.SetThisElementCurrent();
			placeholderShape.txBody.content.MoveCursorToStartPos();
			logicDocument.MoveCursorRight(false, true);

			ExecuteMainHotkey(mainShortcutTypes.checkAddTitleBreakLine);
			CheckAddBreakLine(placeholderShape);
			AscTest.TurnOffRecalculate();
		}

		const checkAddMathBreakLine = (eventIndex) =>
		{
			const {shape} = ClearShapeAndAddParagraph('');
			editor.asc_AddMath2(c_oAscMathType.FractionVertical);
			logicDocument.Recalculate();
			const paragraph = shape.txBody.content.Content[0];
			logicDocument.MoveCursorLeft();
			logicDocument.MoveCursorLeft();
			AscTest.EnterText([56, 56, 56, 56, 56, 56, 56]);
			logicDocument.MoveCursorLeft();
			logicDocument.MoveCursorLeft();
			ExecuteMainHotkey(mainShortcutTypes.checkAddMathBreakLine, eventIndex);
			const paraMath = paragraph.GetAllParaMaths()[0];
			const fraction = paraMath.Root.GetFirstElement();
			const numerator = fraction.getNumerator();
			const eqArray = numerator.GetFirstElement();
			assert.strictEqual(eqArray.getRowsCount(), 2, 'Check add new line in math');
		}

		checkAddMathBreakLine(0);
		checkAddMathBreakLine(1);
	});

	QUnit.test('Check add new paragraph', (assert) =>
	{
		const {shape} = ClearShapeAndAddParagraph('');
		ExecuteMainHotkey(mainShortcutTypes.checkAddParagraph);
		assert.strictEqual(shape.txBody.content.Content.length, 2, 'Check add new paragraph');
	});

	QUnit.test('Check add tab symbol', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		ExecuteMainHotkey(mainShortcutTypes.checkAddTab);
		assert.true(paragraph.Content[0].Content[0].IsTab(), 'Check add tab symbol');
	});

	QUnit.test('Check show paragraph marks', (assert) =>
	{
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.ShowParaMarks);
		assert.true(editor.get_ShowParaMarks(), 'Check show para marks');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.ShowParaMarks);
		assert.false(editor.get_ShowParaMarks(), 'Check hide para marks');
	});

	QUnit.test('Check copy/paste format and clear formatting actions', (assert) =>
	{
		ClearShapeAndAddParagraph('Hello');
		logicDocument.SelectAll();

		logicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Bold: true, Italic: true, Strikeout: true}), true);
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.CopyFormat);

		let textPr = editor.getFormatPainterData().TextPr;
		assert.true(textPr.GetBold(), 'Check copy paragraph property');
		assert.true(textPr.GetItalic(), 'Check copy paragraph property');
		assert.true(textPr.GetStrikeout(), 'Check copy paragraph property');


		ClearShapeAndAddParagraph('Hello');
		logicDocument.SelectAll();
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.PasteFormat);

		textPr = GetDirectTextPr();
		assert.true(textPr.GetBold(), 'Check paste paragraph property');
		assert.true(textPr.GetItalic(), 'Check paste paragraph property');
		assert.true(textPr.GetStrikeout(), 'Check paste paragraph property');

		ExecuteMainHotkey(mainShortcutTypes.checkClearParagraphFormatting);
		textPr = GetDirectTextPr();
		assert.false(textPr.GetBold(), 'Check clear paragraph property');
		assert.false(textPr.GetItalic(), 'Check clear paragraph property');
		assert.false(textPr.GetStrikeout(), 'Check clear paragraph property');
	});

	QUnit.test('Check undo/redo', (assert) =>
	{
		const {paragraph} = ClearShapeAndAddParagraph('');
		AscTest.EnterText([56, 56, 56, 56]);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), '8888', 'Check enter text');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.EditUndo);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), '', 'Check undo text');

		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.EditRedo);
		assert.strictEqual(AscTest.GetParagraphText(paragraph), '8888', 'Check redo text');
	});

	QUnit.test('Check select all', (assert) =>
	{
		const {paragraph, shape} = ClearShapeAndAddParagraph('Hello');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.EditSelectAll);

		assert.true(shape.txBody.content.IsSelectionUse(), 'Check content selection');
		assert.true(paragraph.IsSelectedAll(), 'Check paragraph selection');
		assert.strictEqual(logicDocument.GetSelectedText(), 'Hello', 'Check selected text');
	});

	QUnit.test('Check reset action with adding new shape', (assert) =>
	{
		logicDocument.StartAddShape('rect', true);
		ExecuteMainHotkey(mainShortcutTypes.checkResetAddShape);
		assert.true(!GetDrawingObjects().checkTrackDrawings(), 'Check reset add new shape');
	});

	QUnit.test('Check visit hyperlink', (assert) =>
	{
		CreateSlide();
		ClearShapeAndAddParagraph('Hello');
		logicDocument.AddHyperlink({
			Text   : 'abcd',
			ToolTip: 'abcd',
			Value  : "ppaction://hlinkshowjump?jump=lastslide"
		});
		logicDocument.MoveCursorLeft();
		logicDocument.MoveCursorLeft();
		const selectedInfo = logicDocument.IsCursorInHyperlink();
		ExecuteMainHotkey(mainShortcutTypes.checkVisitHyperlink);
		assert.true(selectedInfo.Visited && logicDocument.GetSelectedSlides()[0] === 1, 'Check visit hyperlink');
	});

	QUnit.test('Check duplicate presentation objects', (assert) =>
	{
		const shape = AddShape(0, 0, 100, 100);
		SelectDrawings([shape]);
		assert.strictEqual(logicDocument.Slides[0].cSld.spTree.length, 1, 'Check duplicate graphic object');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Duplicate);
		assert.strictEqual(logicDocument.Slides[0].cSld.spTree.length, 2, 'Check duplicate graphic object');

		GetDrawingObjects().resetSelection();
		assert.strictEqual(logicDocument.Slides.length, 1, 'Check duplicate slide');
		ExecuteMainShortcut(Asc.c_oAscPresentationShortcutType.Duplicate);
		assert.strictEqual(logicDocument.Slides.length, 2, 'Check duplicate slide');
	});

	QUnit.test('Check save action', (assert) =>
	{
		assert.timeout(1000);
		const done1 = assert.async();
		const done2 = assert.async();

		function CheckSave(Execute, resolve)
		{
			const OldSave = editor._onSaveCallbackInner;
			editor._onSaveCallbackInner = function ()
			{
				resolve();
				editor._onSaveCallbackInner = OldSave;
			};
			editor._saveCheck = () => true;
			editor.asc_isDocumentCanSave = () => true;
			editor.canSave = true;
			Execute(Asc.c_oAscPresentationShortcutType.Save);
		}

		new Promise((resolve) => CheckSave(ExecuteThumbnailShortcut, resolve)).then(function ()
		{
			assert.true(true, 'Check save shortcut');
			done1();
			return new Promise((resolve) => CheckSave(ExecuteMainShortcut, resolve))
		}).then(function ()
		{
			assert.true(true, 'Check save shortcut');
			done2();
		});
	});
})(window);
