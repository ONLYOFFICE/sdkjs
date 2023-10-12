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

'use strict';
(function (window)
{
	const {
		createMathInShape,
		moveCursorDown,
		moveCursorLeft,
		moveCursorRight,
		onInput,
		goToPageWithFocus,
		goToPage,
		checkDirectTextPrAfterKeyDown,
		checkDirectParaPrAfterKeyDown,
		createShape,
		addSlide,
		onKeyDown,
		executeTestWithCatchEvent,
		createNativeEvent,
		checkSelectedSlides,
		cleanPresentation,
		executeCheckMoveShape,
		addPropertyToDocument,
		testMoveHelper,
		createShapeWithTitlePlaceholder,
		createChart,
		createTable,
		checkRemoveObject,
		checkTextAfterKeyDownHelperHelloWorld,
		checkTextAfterKeyDownHelperEmpty,
		createEvent,
		getShapeWithParagraphHelper,
		moveToParagraph,
		getFirstSlide,
		selectOnlyObjects,
		addToSelection,
		createGroup,
		getController,
		oGlobalLogicDocument,
		oMainShortcutTypes,
		oDemonstrationEvents,
		oDemonstrationTypes,
		oThumbnailsTypes,
		oThumbnailsMainFocusTypes,
		startThumbnailsFocusTest,
		startThumbnailsMainFocusTest,
		startMainTest
	} = window.AscTestShortcut;
	let oMockEvent = createNativeEvent();

	function checkSave(oAssert)
	{
		const fOldSave = editor._onSaveCallbackInner;
		let bCheck = false;
		editor._onSaveCallbackInner = function ()
		{
			bCheck = true;
			editor.canSave = true;
		};
		onKeyDown(oMockEvent);
		oAssert.strictEqual(bCheck, true, 'Check save shortcut');
		editor._onSaveCallbackInner = fOldSave;
	}

	$(function ()
	{
		QUnit.module('Check shortcut focus', {
			before: function ()
			{
				addSlide();
			},
			after : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('check shortcut focus', (oAssert) =>
		{
			editor.StartDemonstration("presentation-preview", 0);
			let bCheck = false;
			let fOldKeyDown;
			fOldKeyDown = editor.WordControl.DemonstrationManager.onKeyDown;
			editor.WordControl.DemonstrationManager.onKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check demonstration onKeyDown');
			editor.WordControl.DemonstrationManager.onKeyDown = fOldKeyDown;
			editor.EndDemonstration();

			bCheck = false;
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			fOldKeyDown = editor.WordControl.Thumbnails.onKeyDown;
			editor.WordControl.Thumbnails.onKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check thumbnails onKeyDown');
			editor.WordControl.Thumbnails.onKeyDown = fOldKeyDown;

			bCheck = false;
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			fOldKeyDown = editor.WordControl.m_oLogicDocument.OnKeyDown;
			editor.WordControl.m_oLogicDocument.OnKeyDown = function ()
			{
				bCheck = true;
			}
			editor.WordControl.onKeyDown(createNativeEvent());
			oAssert.true(bCheck, 'Check logic document onKeyDown');
			editor.WordControl.m_oLogicDocument.OnKeyDown = fOldKeyDown;
		});

		QUnit.module("Test thumbnails main focus shortcuts", {
			beforeEach: function ()
			{
				cleanPresentation();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('Test add next slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				const arrOldSlides = oGlobalLogicDocument.Slides.slice();
				onKeyDown(oEvent);
				const arrSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
				oAssert.true(checkSelectedSlides([1]) && (arrOldSlides.indexOf(oGlobalLogicDocument.Slides[arrSelectedSlides[0]]) === -1));

			}, oThumbnailsMainFocusTypes.addNextSlide);
		});

		QUnit.test('Test move to previous slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([3]));
			}, oThumbnailsMainFocusTypes.moveToPreviousSlide);
		});

		QUnit.test('Test move to next slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]));
			}, oThumbnailsMainFocusTypes.moveToNextSlide);
		});

		QUnit.test('Test move to first slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0]));
			}, oThumbnailsMainFocusTypes.moveToFirstSlide);
		});

		QUnit.test('Test select to first slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(4, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1, 2, 3, 4]));
			}, oThumbnailsMainFocusTypes.selectToFirstSlide);
		});

		QUnit.test('Test move to last slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([4]));
			}, oThumbnailsMainFocusTypes.moveToLastSlide);
		});

		QUnit.test('Test select to last slide', (oAssert) =>
		{
			startThumbnailsMainFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1, 2, 3, 4]));
			}, oThumbnailsMainFocusTypes.selectToLastSlide);
		});

		QUnit.test('Test thumbnails shortcut actions', (oAssert) =>
		{
			cleanPresentation();
			addSlide();
			addSlide();
			addSlide();
			addSlide();
			const fOldShortcut = editor.getShortcut;
			goToPage(0);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditSelectAll;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			onKeyDown(createNativeEvent());
			oAssert.true(checkSelectedSlides([0, 1, 2, 3]), 'Check select all slides');

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Duplicate;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			const arrOldSlides = oGlobalLogicDocument.Slides.slice();
			onKeyDown(createNativeEvent());
			oAssert.true(checkSelectedSlides([1]) && oGlobalLogicDocument.Slides.length === 5 && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check duplicate slides');

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Print;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			executeTestWithCatchEvent('asc_onPrint', () => true, true, createNativeEvent(), oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Save;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			checkSave(oAssert);

			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowContextMenu;};
			goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oMockEvent, oAssert);

			cleanPresentation();
			editor.getShortcut = fOldShortcut;
		});

		QUnit.module('Test thumbnails hotkeys', {
			before: function ()
			{
				cleanPresentation();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
			},
			after : function ()
			{
				cleanPresentation();
			}
		});
		let arrOldSlides;
		let oOldSlide;
		QUnit.test('Test thumbnails focus event', (oAssert) =>
		{
			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				arrOldSlides = oGlobalLogicDocument.Slides.slice();
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]) && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check add next slide');
			}, oThumbnailsTypes.addNextSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				oOldSlide = getFirstSlide();
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0]) && oGlobalLogicDocument.Slides.indexOf(oOldSlide) === -1, 'Check remove selected slides');
			}, oThumbnailsTypes.removeSelectedSlides);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				const oFirstSlide = getFirstSlide();
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[3] === oFirstSlide, 'Check move selected slides to end');
			}, oThumbnailsTypes.moveSelectedSlidesToEnd);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				const oFirstSlide = getFirstSlide();
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[1] === oFirstSlide, 'Check move selected slides to next pos');
			}, oThumbnailsTypes.moveSelectedSlidesToNextPosition);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1]), 'Check select next slide');
			}, oThumbnailsTypes.selectNextSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]), 'Check move to next slide');
			}, oThumbnailsTypes.moveToNextSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0]), 'Check move to first slide');
			}, oThumbnailsTypes.moveToFirstSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1, 2]), 'Check select from current position to first slide');
			}, oThumbnailsTypes.selectToFirstSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([3]), 'Check move to last slide');
			}, oThumbnailsTypes.moveToLastSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([0, 1, 2, 3]), 'Check select from current position to last slide');
			}, oThumbnailsTypes.selectToLastSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				const oLastSlide = oGlobalLogicDocument.Slides[2];
				onKeyDown(oEvent);
				oAssert.true(getFirstSlide() === oLastSlide, 'Check move selected slides to start');
			}, oThumbnailsTypes.moveSelectedSlidesToStart);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				const oLastSlide = oGlobalLogicDocument.Slides[2];
				onKeyDown(oEvent);
				oAssert.true(oGlobalLogicDocument.Slides[1] === oLastSlide, 'Check move selected slides to previous pos');
			}, oThumbnailsTypes.moveSelectedSlidesToPreviousPosition);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1, 2]), 'Check select previous slide');
			}, oThumbnailsTypes.selectPreviousSlide);

			startThumbnailsFocusTest((oEvent) =>
			{
				goToPageWithFocus(2, FOCUS_OBJECT_THUMBNAILS);
				onKeyDown(oEvent);
				oAssert.true(checkSelectedSlides([1]), 'Check move to previous slide');
			}, oThumbnailsTypes.moveToPreviousSlide);
		});

		QUnit.module('Test demonstration mode shortcuts', {
			beforeEach: function ()
			{
				cleanPresentation();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
				addSlide();
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});

		QUnit.test('Test demonstration mode shortcuts', (oAssert) =>
		{
			editor.StartDemonstration("presentation-preview", 0);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 1, oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide][0].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide][1].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide][2].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide][3].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 5, oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide][4].event, oAssert);

			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 4, oDemonstrationEvents[oDemonstrationTypes.moveToPreviousSlide][0].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 3, oDemonstrationEvents[oDemonstrationTypes.moveToPreviousSlide][1].event, oAssert);
			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 2, oDemonstrationEvents[oDemonstrationTypes.moveToPreviousSlide][2].event, oAssert);

			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 0, oDemonstrationEvents[oDemonstrationTypes.moveToFirstSlide][0].event, oAssert);

			executeTestWithCatchEvent('asc_onDemonstrationSlideChanged', (nSlideNum) => nSlideNum, 5, oDemonstrationEvents[oDemonstrationTypes.moveToLastSlide][0].event, oAssert);

			executeTestWithCatchEvent('asc_onEndDemonstration', () => true, true, oDemonstrationEvents[oDemonstrationTypes.exitFromDemonstrationMode][0].event, oAssert);

			editor.EndDemonstration();
		});

		QUnit.module("Test main focus shortcuts", {
			beforeEach: function ()
			{
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			},
			afterEach : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('Test if the desired action is received by the keyboard shortcut.', (oAssert) =>
		{
			let oEvent;
			oEvent = createEvent(65, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditSelectAll, 'Check getting select all shortcut action');

			oEvent = createEvent(90, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditUndo, 'Check getting undo shortcut action');

			oEvent = createEvent(89, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EditRedo, 'Check getting redo shortcut action');

			oEvent = createEvent(88, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Cut, 'Check getting cut shortcut action');

			oEvent = createEvent(67, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Copy, 'Check getting copy shortcut action');

			oEvent = createEvent(86, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Paste, 'Check getting paste shortcut action');

			oEvent = createEvent(68, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Duplicate, 'Check getting duplicate shortcut action');

			oEvent = createEvent(80, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Print, 'Check getting print shortcut action');

			oEvent = createEvent(83, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Save, 'Check getting save shortcut action');

			oEvent = createEvent(93, false, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(121, false, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(57351, false, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowContextMenu, 'Check getting show context menu shortcut action');

			oEvent = createEvent(56, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.ShowParaMarks, 'Check getting show paragraph marks shortcut action');

			oEvent = createEvent(66, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Bold, 'Check getting bold shortcut action');

			oEvent = createEvent(67, true, false, true, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.CopyFormat, 'Check getting copy format shortcut action');

			oEvent = createEvent(69, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.CenterAlign, 'Check getting center align shortcut action');

			oEvent = createEvent(69, true, false, true, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EuroSign, 'Check getting euro sign shortcut action');

			oEvent = createEvent(71, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Group, 'Check getting group shortcut action');

			oEvent = createEvent(71, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.UnGroup, 'Check getting ungroup shortcut action');

			oEvent = createEvent(73, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Italic, 'Check getting italic shortcut action');

			oEvent = createEvent(74, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.JustifyAlign, 'Check getting justify align shortcut action');

			oEvent = createEvent(75, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.AddHyperlink, 'Check getting add hyperlink shortcut action');

			oEvent = createEvent(76, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.BulletList, 'Check getting bullet list shortcut action');

			oEvent = createEvent(76, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.LeftAlign, 'Check getting left align shortcut action');

			oEvent = createEvent(82, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.RightAlign, 'Check getting right align shortcut action');

			oEvent = createEvent(85, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Underline, 'Check getting underline shortcut action');

			oEvent = createEvent(53, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Strikethrough, 'Check getting strikethrough shortcut action');

			oEvent = createEvent(86, true, false, true, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.PasteFormat, 'Check getting paste format shortcut action');

			oEvent = createEvent(188, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Superscript, 'Check getting superscript shortcut action');

			oEvent = createEvent(190, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.Subscript, 'Check getting subscript shortcut action');

			oEvent = createEvent(189, true, true, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.EnDash, 'Check getting en dash shortcut action');

			oEvent = createEvent(219, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.DecreaseFont, 'Check getting decrease font size shortcut action');

			oEvent = createEvent(221, true, false, false, false, false);
			oAssert.strictEqual(editor.getShortcut(oEvent), Asc.c_oAscPresentationShortcutType.IncreaseFont, 'Check getting increase font size shortcut action');
		});

		let fOldGetShortcut;
		QUnit.module('Test main shortcut actions', {
			before: function ()
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				fOldGetShortcut = editor.getShortcut;
			},
			after : function ()
			{
				cleanPresentation();
				editor.getShortcut = fOldGetShortcut;
			}
		});

		QUnit.test('Test undo shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditUndo;};
			createShape();
			onKeyDown(oMockEvent);
			oAssert.strictEqual(getFirstSlide().cSld.spTree.length, 0);
		});

		QUnit.test('Test redo shortcut', (oAssert) =>
		{
			const oUndoShape = createShape();
			editor.Undo();
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditRedo;};
			onKeyDown(oMockEvent);
			oAssert.strictEqual(getFirstSlide().cSld.spTree.length === 1 && getFirstSlide().cSld.spTree[0] === oUndoShape, true);
		});

		QUnit.test('Test select all shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EditSelectAll;};
			cleanPresentation();
			addSlide();
			goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			const oFirstShape = createShape();
			const oSecondShape = createShape();
			onKeyDown(oMockEvent);
			const oController = getController();
			oAssert.strictEqual(oController.selectedObjects.length === 2 && oController.selectedObjects.indexOf(oFirstShape) !== -1 && oController.selectedObjects.indexOf(oSecondShape) !== -1, true);
		});

		QUnit.test('Test duplicate shape', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Duplicate;};
			const {oShape} = getShapeWithParagraphHelper('', true);
			const arrOldSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree.slice();
			selectOnlyObjects([oShape]);
			onKeyDown(oMockEvent);
			const arrUpdatedSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
			const oNewShape = arrUpdatedSpTree[arrUpdatedSpTree.length - 1];
			oAssert.true(arrOldSpTree.indexOf(oNewShape) === -1);
		});

		QUnit.test('Test print', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Print;};
			executeTestWithCatchEvent('asc_onPrint', () => true, true, oMockEvent, oAssert);
		});

		QUnit.test('Test save', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Save;};
			checkSave(oAssert);
		});

		QUnit.test('Test context menu', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowContextMenu;};
			executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oMockEvent, oAssert, () =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('');
				oParagraph.SetThisElementCurrent();
			});
		});

		QUnit.test('Test show para marks', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.ShowParaMarks;};
			editor.put_ShowParaMarks(false);
			onKeyDown(oMockEvent);
			oAssert.true(!!editor.get_ShowParaMarks(), 'Check show para marks shortcut');
		});

		QUnit.test('Test bold shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Bold;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Bold(), oMockEvent, oAssert, true, 'Check bold shortcut');
		});

		function getCopyParagraphPrTest()
		{
			const oCopyParagraphTextPr = new AscCommonWord.CTextPr();
			oCopyParagraphTextPr.SetUnderline(true);
			oCopyParagraphTextPr.SetBold(true);
			oCopyParagraphTextPr.BoldCS = true;
			oCopyParagraphTextPr.SetItalic(true);
			oCopyParagraphTextPr.ItalicCS = true;
			return oCopyParagraphTextPr;
		}

		QUnit.test('Test copy format', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.CopyFormat;};
			const {oParagraph} = getShapeWithParagraphHelper('Hello World');
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});

			onKeyDown(oMockEvent);
			const oTextPr = editor.getFormatPainterData().TextPr;
			oAssert.strictEqual(oTextPr.Get_Bold(), true, 'Check copy format shortcut');
			oAssert.strictEqual(oTextPr.Get_Italic(), true, 'Check copy format shortcut');
			oAssert.strictEqual(oTextPr.Get_Underline(), true, 'Check copy format shortcut');
		});

		QUnit.test('Test paste format', (oAssert) =>
		{
			let oParagraph;
			oParagraph = getShapeWithParagraphHelper('Hello World').oParagraph;
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});
			oGlobalLogicDocument.Document_Format_Copy();
			
			
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.PasteFormat;};
			oParagraph = getShapeWithParagraphHelper('Hello World').oParagraph;
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();

			onKeyDown(oMockEvent);
			const oDirectTextPr = oParagraph.GetDirectTextPr();
			oAssert.strictEqual(oDirectTextPr.Get_Bold(), true, 'Check copy format shortcut');
			oAssert.strictEqual(oDirectTextPr.Get_Italic(), true, 'Check copy format shortcut');
			oAssert.strictEqual(oDirectTextPr.Get_Underline(), true, 'Check copy format shortcut');
		});

		QUnit.test('Test center align', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.CenterAlign;};
			checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Center, 'Check center align shortcut');
		});

		QUnit.test('Test insert euro sign', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EuroSign;};
			checkTextAfterKeyDownHelperEmpty('€', oMockEvent, oAssert, 'Check euro sign shortcut');
		});

		QUnit.test('Test group', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Group;};
			const oFirstShape = createShape();
			const oSecondShape = createShape();
			selectOnlyObjects([oFirstShape, oSecondShape]);

			onKeyDown(oMockEvent);
			const oGroup = oFirstShape.group;
			oAssert.true(oFirstShape.group && (oFirstShape.group === oSecondShape.group), 'Check group shortcut');
		});

		QUnit.test('Test ungroup', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.UnGroup;};
			const oFirstShape = createShape();
			const oSecondShape = createShape();
			const oGroup = createGroup([oFirstShape, oSecondShape]);
			selectOnlyObjects([oGroup]);

			onKeyDown(oMockEvent);
			oAssert.true(!oFirstShape.group && !oSecondShape.group && oGroup.bDeleted, 'Check ungroup shortcut');
		});

		QUnit.test('Test italic', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Italic;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Italic(), oMockEvent, oAssert, true, 'Check italic shortcut');
		});

		QUnit.test('Test justify align', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.JustifyAlign;};
			checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Justify, 'check justify align shortcut');
		});

		QUnit.test('Test open hyperlink dialog', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.AddHyperlink;};
			executeTestWithCatchEvent('asc_onDialogAddHyperlink', () => true, true, oMockEvent, oAssert, () =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello World');
				moveToParagraph(oParagraph);
				oGlobalLogicDocument.SelectAll();
			});
		});

		QUnit.test('Test add bullet list to paragraphs', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.BulletList;};
			const {oParagraph} = getShapeWithParagraphHelper('Hello World');
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();

			onKeyDown(oMockEvent);
			const oBullet = oParagraph.Get_PresentationNumbering();
			oAssert.true(oBullet.m_nType === AscFormat.numbering_presentationnumfrmt_Char, 'Check bullet list shortcut');
		});

		QUnit.test('Test left align', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.LeftAlign;};
			checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Left, 'check right align shortcut');
		});

		QUnit.test('Test right align', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.RightAlign;};
			checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.GetJc(), oMockEvent, oAssert, AscCommon.align_Right, 'check right align shortcut');
		});

		QUnit.test('Test underline shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Underline;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Underline(), oMockEvent, oAssert, true, 'Check underline shortcut');
		});

		QUnit.test('Test strikeout shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Strikethrough;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Strikeout(), oMockEvent, oAssert, true, 'Check strikeout shortcut');
		});

		QUnit.test('Test superscript vertalign', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Superscript;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oMockEvent, oAssert, AscCommon.vertalign_SuperScript, 'Check superscript shortcut');
		});

		QUnit.test('Test subscript vertalign', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.Subscript;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.GetVertAlign(), oMockEvent, oAssert, AscCommon.vertalign_SubScript, 'Check subscript shortcut');
		});

		QUnit.test('Test en dash shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.EnDash;};
			checkTextAfterKeyDownHelperEmpty('–', oMockEvent, oAssert, 'Check en dash shortcut');
		});

		QUnit.test('Test decrease font size', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.DecreaseFont;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), oMockEvent, oAssert, 9, 'Check decrease font size shortcut');
		});

		QUnit.test('Test increase font size', (oAssert) =>
		{
			editor.getShortcut = function () {return Asc.c_oAscPresentationShortcutType.IncreaseFont;};
			checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), oMockEvent, oAssert, 11, 'Check increase font size shortcut');
		});

		QUnit.module('Test hotkeys', {
			before: function ()
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
			},
			after : function ()
			{
				cleanPresentation();
			}
		});
		QUnit.test('Test delete back char', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				checkTextAfterKeyDownHelperHelloWorld('Hello Worl', oEvent, oAssert, 'Check delete with backspace')
			}, oMainShortcutTypes.checkDeleteBack);
		});

		QUnit.test('Test delete back word', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				checkTextAfterKeyDownHelperHelloWorld('Hello ', oEvent, oAssert, 'Check delete word with backspace')
			}, oMainShortcutTypes.checkDeleteWordBack);
		});

		QUnit.test('Test remove animation', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape} = getShapeWithParagraphHelper('', true);
				selectOnlyObjects([oShape]);
				oGlobalLogicDocument.AddAnimation(1, 1, 0, null, false, false);

				onKeyDown(oEvent);
				const oTiming = oGlobalLogicDocument.GetCurTiming();
				const arrEffects = oTiming.getObjectEffects(oShape.GetId());
				oAssert.true(arrEffects.length === 0, 'Check remove animation');
			}, oMainShortcutTypes.checkRemoveAnimation);
		});

		QUnit.test('Test remove chart', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oChart1 = createChart(getFirstSlide());
				selectOnlyObjects([oChart1]);
				onKeyDown(oEvent);
				oAssert.true(checkRemoveObject(oChart1, getFirstSlide().cSld.spTree), "Check remove chart");
			}, oMainShortcutTypes.checkRemoveChart);
		});

		QUnit.test('Test remove shape', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape} = getShapeWithParagraphHelper('', true);
				selectOnlyObjects([oShape]);

				onKeyDown(oEvent);
				const arrSpTree = oGlobalLogicDocument.Slides[0].cSld.spTree;
				oAssert.true(checkRemoveObject(oShape, arrSpTree), 'Check remove shape');
			}, oMainShortcutTypes.checkRemoveShape);
		});

		QUnit.test('Test remove table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oGraphicFrame = createTable(3, 3);
				selectOnlyObjects([oGraphicFrame]);
				onKeyDown(oEvent);
				const arrSpTree = getFirstSlide().cSld.spTree;
				oAssert.true(checkRemoveObject(oGraphicFrame, arrSpTree), "Check remove table");
			}, oMainShortcutTypes.checkRemoveTable);
		});

		QUnit.test('Test remove group', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oGroup1 = createGroup([createShape(), createShape()]);
				selectOnlyObjects([oGroup1]);
				onKeyDown(oEvent);
				const arrSpTree = getFirstSlide().cSld.spTree;
				oAssert.true(checkRemoveObject(oGroup1, arrSpTree), 'Check remove group');
			}, oMainShortcutTypes.checkRemoveGroup);
		});

		QUnit.test('Test remove shape in group', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oGroupedGroup = createGroup([createShape(), createShape()]);
				const oRemovedShape = createShape();
				const oGroup1 = createGroup([oGroupedGroup, oRemovedShape]);
				selectOnlyObjects([oRemovedShape]);
				onKeyDown(oEvent);
				oAssert.true(checkRemoveObject(oRemovedShape, oGroup1.spTree), 'Check remove shape in group');
			}, oMainShortcutTypes.checkRemoveShapeInGroup);
		});

		//Tab
		QUnit.test('Test go to next cell', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oGraphicFrame = createTable(3, 3);
				const oTable = oGraphicFrame.graphicObject;
				onKeyDown(oEvent);
				oAssert.strictEqual(oTable.CurCell.Index, 1, 'check go to next cell shortcut');
			}, oMainShortcutTypes.checkMoveToNextCell);
		});

		QUnit.test('Test go to previous cell', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oGraphicFrame = createTable(3, 3);
				const oTable = oGraphicFrame.graphicObject;
				moveCursorRight();
				onKeyDown(oEvent);
				oAssert.strictEqual(oTable.CurCell.Index, 0, 'check go to previous cell shortcut');
			}, oMainShortcutTypes.checkMoveToPreviousCell);
		});

		QUnit.test('Test bullet indent', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello');
				const oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
				oParagraph.Add_PresentationNumbering(oBullet);
				moveToParagraph(oParagraph, true);
				oParagraph.Set_Ind({Left: 0});
				onKeyDown(oEvent);
				oAssert.strictEqual(oParagraph.Pr.Get_IndLeft(), 11.1125, 'Check bullet indent shortcut');
			}, oMainShortcutTypes.checkIncreaseBulletIndent);
		});

		QUnit.test('Test bullet unindent', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello');
				const oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
				oParagraph.Add_PresentationNumbering(oBullet);
				oParagraph.Set_PresentationLevel(1);
				moveToParagraph(oParagraph, true);
				oParagraph.Set_Ind({Left: 11.1125});
				onKeyDown(oEvent);
				oAssert.strictEqual(oParagraph.Pr.Get_IndLeft(), 0, 'Check bullet indent shortcut');
			}, oMainShortcutTypes.checkDecreaseBulletIndent);
		});

		QUnit.test('Test add tab', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('');
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oGlobalLogicDocument.SelectAll();
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(false, {TabSymbol: '\t'}), '\t', 'Check add tab');
			}, oMainShortcutTypes.checkAddTab);
		});

		QUnit.test('Test select next object', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				const oFirstShape = createShape();
				const oSecondShape = createShape();
				const oThirdShape = createShape();
				selectOnlyObjects([oFirstShape]);
				onKeyDown(oEvent);
				oAssert.true(getController().getSelectedArray().length === 1 && getController().getSelectedArray()[0] === oSecondShape, 'Check select previous object');
			}, oMainShortcutTypes.checkSelectNextObject);
		});

		QUnit.test('Test select previous object', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				const oFirstShape = createShape();
				const oSecondShape = createShape();
				const oThirdShape = createShape();
				selectOnlyObjects([oFirstShape]);
				onKeyDown(oEvent);
				oAssert.true(getController().getSelectedArray().length === 1 && getController().getSelectedArray()[0] === oThirdShape, 'Check select previous object');
			}, oMainShortcutTypes.checkSelectPreviousObject);
		});
		// Enter
		QUnit.test('Test check visit hyperlink', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				goToPage(1);
				const {oParagraph} = getShapeWithParagraphHelper('Hello');
				moveToParagraph(oParagraph);
				oGlobalLogicDocument.AddHyperlink({
					Text   : 'abcd',
					ToolTip: 'abcd',
					Value  : 'ppaction://hlinkshowjump?jump=firstslide'
				});
				moveCursorLeft();
				moveCursorLeft();
				onKeyDown(oEvent);
				const oSelectedInfo = oGlobalLogicDocument.IsCursorInHyperlink();
				oAssert.true(oSelectedInfo.Visited && oGlobalLogicDocument.GetSelectedSlides()[0] === 0, 'Check visit hyperlink');
				goToPage(0);
			}, oMainShortcutTypes.checkVisitHyperlink);
		});

		QUnit.test('Test select shapes with placeholder', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);
				const oFirstShapeWithPlaceholder = createShapeWithTitlePlaceholder();
				const oSecondShapeWithPlaceholder = createShapeWithTitlePlaceholder();

				const oController = getController();
				oController.resetSelection();
				onKeyDown(oEvent);
				oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oFirstShapeWithPlaceholder && oFirstShapeWithPlaceholder.selected, 'Check select first shape with placeholder');

				onKeyDown(oEvent);
				oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oSecondShapeWithPlaceholder && oSecondShapeWithPlaceholder.selected, 'Check select second shape with placeholder');
			}, oMainShortcutTypes.checkSelectNextObjectWithPlaceholder);


		});

		QUnit.test('Test add next slide after placeholder shape', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				cleanPresentation();
				addSlide();
				goToPageWithFocus(0, FOCUS_OBJECT_MAIN);

				const oFirstShapeWithPlaceholder = createShapeWithTitlePlaceholder();
				selectOnlyObjects([oFirstShapeWithPlaceholder]);
				const arrOldSlides = oGlobalLogicDocument.Slides.slice();
				onKeyDown(oEvent);
				const arrSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
				oAssert.true(arrSelectedSlides.length === 1 && arrSelectedSlides[0] === 1 && arrOldSlides.indexOf(oGlobalLogicDocument.Slides[1]) === -1, 'Check add next slide after selecting last placeholder on current slide');
				goToPage(0);
			}, oMainShortcutTypes.checkAddNextSlideAfterSelectLastPlaceholderObject);
		});

		QUnit.test('Test add break line', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape, oParagraph} = getShapeWithParagraphHelper('');
				moveToParagraph(oParagraph);

				onKeyDown(oEvent);
				oAssert.true(oShape.getDocContent().Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line');
			}, oMainShortcutTypes.checkAddBreakLine);
		});

		QUnit.test('Test add break line in title', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShapeWithPlaceholder = createShapeWithTitlePlaceholder();
				const oContent = oShapeWithPlaceholder.getDocContent();
				const oParagraph = oContent.GetAllParagraphs()[0];
				oParagraph.SetThisElementCurrent();
				onKeyDown(oEvent);
				oAssert.true(oContent.Content.length === 1 && oParagraph.GetLinesCount() === 2, 'Check add break line in title');
			}, oMainShortcutTypes.checkAddTitleBreakLine);
		});

		QUnit.test('Test add new line in math equation', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = createMathInShape();
				oGlobalLogicDocument.MoveCursorToStartPos();
				moveCursorRight();
				moveCursorRight();
				onInput([56, 56, 56, 56, 56, 56, 56]);
				moveCursorLeft();
				moveCursorLeft();
				onKeyDown(oEvent);
				const oParaMath = oParagraph.GetAllParaMaths()[0];
				const oFraction = oParaMath.Root.GetFirstElement();
				const oNumerator = oFraction.getNumerator();
				const oEqArray = oNumerator.GetFirstElement();
				oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');
			}, oMainShortcutTypes.checkAddMathBreakLine);
		});

		QUnit.test('Test add new paragraph', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape, oParagraph} = getShapeWithParagraphHelper('');
				moveToParagraph(oParagraph);

				onKeyDown(oEvent);
				oAssert.true(oShape.getDocContent().Content.length === 2, 'Check add new paragraph');
			}, oMainShortcutTypes.checkAddParagraph);
		});

		QUnit.test('Test creating txBody', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = createShape();
				selectOnlyObjects([oShape]);
				onKeyDown(oEvent);
				oAssert.true(!!oShape.txBody, 'Check creating txBody');
			}, oMainShortcutTypes.checkAddTxBodyShape);
		});
		QUnit.test('Test move cursor to start position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape, oParagraph} = getShapeWithParagraphHelper('', true);
				selectOnlyObjects([oShape]);
				onKeyDown(oEvent);
				oAssert.true(oParagraph.IsCursorAtBegin(), 'Check move cursor to start position in shape');
			}, oMainShortcutTypes.checkMoveCursorToStartPosShape);
		});
		QUnit.test('Test select all content in shape', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oShape} = getShapeWithParagraphHelper('Hello Word', true);
				selectOnlyObjects([oShape]);
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Hello Word', 'Check select all content in shape');
			}, oMainShortcutTypes.checkSelectAllContentShape);
		});
		QUnit.test('Test select all title', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oChart = createChart();
				selectOnlyObjects([oChart]);
				const oTitles = oChart.getAllTitles();
				const oController = getController();
				oController.selection.chartSelection = oChart;
				oChart.selectTitle(oTitles[0], 0);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Diagram Title', 'Check select all title');
			}, oMainShortcutTypes.checkSelectAllContentChartTitle);
		});
		QUnit.test('Test move cursor to begin pos in title', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oChart = createChart();
				const oTitles = oChart.getAllTitles();
				const oContent = AscFormat.CreateDocContentFromString('', editor.WordControl.m_oDrawingDocument, oTitles[0].txBody);
				oTitles[0].txBody.content = oContent;
				selectOnlyObjects([oChart]);

				const oController = getController();
				oController.selection.chartSelection = oChart;
				oChart.selectTitle(oTitles[0], 0);

				onKeyDown(oEvent);
				oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');
			}, oMainShortcutTypes.checkMoveCursorToStartPosChartTitle);
		});
		QUnit.test('Test remove and move to start position in table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const arrSteps = [];
				const oFrame = createTable(3, 3);
				oFrame.Set_CurrentElement();
				const oTable1 = oFrame.graphicObject;
				oTable1.MoveCursorToStartPos();
				// First cell
				moveCursorRight(true, true);
				moveCursorRight(true, true);
				// Second cell
				moveCursorRight(true, true);
				// Third cell
				moveCursorRight(true, true);

				onKeyDown(oEvent);
				arrSteps.push(oTable1.IsCursorAtBegin());
				moveCursorRight(true, true);
				moveCursorRight(true, true);
				moveCursorRight(true, true);
				arrSteps.push(oGlobalLogicDocument.GetSelectedText());
				oAssert.deepEqual(arrSteps, [true, ''], 'Check remove and move to start position in table');
			}, oMainShortcutTypes.checkRemoveAndMoveToStartPosTable);
		});
		QUnit.test('Test select first cell content', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oFrame = createTable(3, 3);
				selectOnlyObjects([oFrame]);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedText(), 'Cell0x0', 'Check select first cell content');
			}, oMainShortcutTypes.checkSelectFirstCellContent);
		});
		// Esc
		QUnit.test('Test reset add new shape', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oController = getController();
				oGlobalLogicDocument.StartAddShape('rect', true);
				onKeyDown(oEvent);
				oAssert.true(!oController.checkTrackDrawings(), 'Check reset add new shape');
			}, oMainShortcutTypes.checkResetAddShape);
		});

		QUnit.test('Test reset all selection', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oController = getController();
				oController.resetSelection();
				const oGroupedShape1 = createShape();
				const oGroupedShape2 = createShape();
				createGroup([oGroupedShape1, oGroupedShape2]);
				addToSelection(oGroupedShape1);
				const oTestGroup = oGroupedShape1.group;
				onKeyDown(oEvent);
				oAssert.true(oController.selectedObjects.length === 0, 'Check reset all selection');
			}, oMainShortcutTypes.checkResetAllDrawingSelection);
		});

		QUnit.test('Test reset step selection', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oController = getController();
				const oGroupedShape1 = createShape();
				const oGroupedShape2 = createShape();
				createGroup([oGroupedShape1, oGroupedShape2]);
				addToSelection(oGroupedShape1);
				const oTestGroup = oGroupedShape1.group;

				selectOnlyObjects([oTestGroup, oGroupedShape1]);
				onKeyDown(oEvent);
				oAssert.true(oController.selectedObjects.length === 1 && oController.selectedObjects[0] === oTestGroup && oTestGroup.selectedObjects.length === 0, 'Check reset step selection');
			}, oMainShortcutTypes.checkResetStepDrawingSelection);
		});

		// Space
		QUnit.test('Test add non breaking space', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A0), oEvent, oAssert, 'Check add non breaking space');
			}, oMainShortcutTypes.checkNonBreakingSpace);
		});

		QUnit.test('Test clear paragraph formatting', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello World');
				oParagraph.SetThisElementCurrent();
				oGlobalLogicDocument.SelectAll();
				addPropertyToDocument({Bold: true, Italic: true, Underline: true});

				onKeyDown(oEvent);
				const oTextPr = oGlobalLogicDocument.GetDirectTextPr();
				oAssert.true(!(oTextPr.GetBold() || oTextPr.GetItalic() || oTextPr.GetUnderline()), 'Check clear paragraph formatting');
			}, oMainShortcutTypes.checkClearParagraphFormatting);
		});

		QUnit.test('Test add space', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				checkTextAfterKeyDownHelperEmpty(' ', oEvent, oAssert, 'Check add space')
			}, oMainShortcutTypes.checkAddSpace);
		});
		//pgUp

		//End
		QUnit.test('Test move cursor to end position shortcut', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, false, true, false);
				oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to end position shortcut');
			}, oMainShortcutTypes.checkMoveToEndPosContent);
		});

		QUnit.test('Test move cursor to end line shortcut', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, false, true, false);
				oAssert.true(oPos.X === 100 && oPos.Y === 15, 'Check move cursor to end line shortcut');
			}, oMainShortcutTypes.checkMoveToEndLineContent);
		});

		QUnit.test('Test select text to end line', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
				oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to end line shortcut');
			}, oMainShortcutTypes.checkSelectToEndLineContent);
		});

		// Home
		QUnit.test('Test move to start position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move to start position shortcut');
			}, oMainShortcutTypes.checkMoveToStartPosContent);
		});

		QUnit.test('Test move to start line', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 0 && oPos.Y === 75, 'Check move to start line shortcut');
			}, oMainShortcutTypes.checkMoveToStartLineContent);
		});

		QUnit.test('Test select to start line', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
				oAssert.strictEqual(sSelectedText, 'Hello', 'Check select to start line shortcut');
			}, oMainShortcutTypes.checkSelectToStartLineContent);
		});

		//Left arrow
		QUnit.test('Test move cursor to end position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 20 && oPos.Y === 75, 'Check move cursor to end position shortcut');
			}, oMainShortcutTypes.checkMoveCursorLeft);
		});

		QUnit.test('Test select text to left position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
				oAssert.strictEqual(sSelectedText, 'o', 'Check select text to left position shortcut');
			}, oMainShortcutTypes.checkSelectCursorLeft);
		});

		QUnit.test('Test select word text to left position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
				oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to left position shortcut');
			}, oMainShortcutTypes.checkSelectWordCursorLeft);
		});

		QUnit.test('Test move cursor to left word position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 0 && oPos.Y === 15, 'Check move cursor to left word position shortcut');
			}, oMainShortcutTypes.checkMoveCursorWordLeft);
		});
		QUnit.test('Test move left in table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oFrame = createTable(3, 3);
				oFrame.Set_CurrentElement();
				const oTable1 = oFrame.graphicObject;
				oTable1.MoveCursorToStartPos();
				moveCursorRight(true);
				moveCursorRight(true);
				onKeyDown(oEvent);
				oAssert.deepEqual([oTable1.CurCell.Row.Index, oTable1.CurCell.Index], [0, 0], 'Check move left in table');
			}, oMainShortcutTypes.checkMoveCursorLeftTable);
		});
		//Right arrow
		QUnit.test('Test move cursor to right position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, false, true, false);
				oAssert.true(oPos.X === 5 && oPos.Y === 15, 'Check move cursor to right position shortcut');
			}, oMainShortcutTypes.checkMoveCursorRight);
		});
		QUnit.test('Test move right in table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oFrame = createTable(3, 3);
				oFrame.Set_CurrentElement();
				const oTable1 = oFrame.graphicObject;
				oTable1.MoveCursorToStartPos();
				moveCursorRight(true);
				onKeyDown(oEvent);
				oAssert.deepEqual([oTable1.CurCell.Row.Index, oTable1.CurCell.Index], [0, 1], 'Check move right in table');
			}, oMainShortcutTypes.checkMoveCursorRightTable);
		});

		QUnit.test('Test select text to right position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
				oAssert.strictEqual(sSelectedText, 'H', 'Check select text to right position shortcut');
			}, oMainShortcutTypes.checkSelectCursorRight);
		});

		QUnit.test('Test select word text to right position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
				oAssert.strictEqual(sSelectedText, 'HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', 'Check select word text to right position shortcut');
			}, oMainShortcutTypes.checkSelectWordCursorRight);
		});

		QUnit.test('Test move cursor to right word position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 25 && oPos.Y === 75, 'Check move cursor to right word position shortcut');
			}, oMainShortcutTypes.checkMoveCursorWordRight);
		});
		//Top arrow
		QUnit.test('Test move cursor to top position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, true, true, false);
				oAssert.true(oPos.X === 25 && oPos.Y === 55, 'Check move cursor to top position shortcut');
			}, oMainShortcutTypes.checkMoveCursorTop);
		});
		QUnit.test('Test move top in table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oFrame = createTable(3, 3);
				oFrame.Set_CurrentElement();
				const oTable1 = oFrame.graphicObject;
				oTable1.MoveCursorToStartPos();
				moveCursorDown();
				onKeyDown(oEvent);
				oAssert.deepEqual([oTable1.CurCell.Row.Index, oTable1.CurCell.Index], [0, 0], 'Check move top in table');
			}, oMainShortcutTypes.checkMoveCursorTopTable);
		});

		QUnit.test('Test select text to top position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, true, false, true);
				oAssert.strictEqual(sSelectedText, 'worldHelloworldHello', 'Check select text to top position shortcut');
			}, oMainShortcutTypes.checkSelectCursorTop);
		});
		// Bottom arrow
		QUnit.test('Test move cursor to bottom position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oPos} = testMoveHelper(oEvent, false, true, false);
				oAssert.true(oPos.X === 0 && oPos.Y === 35, 'Check move cursor to bottom position shortcut');
			}, oMainShortcutTypes.checkMoveCursorBottom);
		});
		QUnit.test('Test move bottom in table', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oFrame = createTable(3, 3);
				oFrame.Set_CurrentElement();
				const oTable1 = oFrame.graphicObject;
				oTable1.MoveCursorToStartPos();
				onKeyDown(oEvent);
				oAssert.deepEqual([oTable1.CurCell.Row.Index, oTable1.CurCell.Index], [1, 0], 'Check move bottom in table');
			}, oMainShortcutTypes.checkMoveCursorBottomTable);
		});

		QUnit.test('Test select text to bottom position', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {sSelectedText} = testMoveHelper(oEvent, false, false, true);
				oAssert.strictEqual(sSelectedText, 'HelloworldHelloworld', 'Check select text to bottom position shortcut');
			}, oMainShortcutTypes.checkSelectCursorBottom);
		});

		// Check move shape
		QUnit.test('Test big move shape bottom', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.y, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape bottom');
			}, oMainShortcutTypes.checkMoveShapeBottom);
		});

		QUnit.test('Test little move shape bottom', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.y, 1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape bottom');
			}, oMainShortcutTypes.checkLittleMoveShapeBottom);
		});

		QUnit.test('Test big move shape top', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.y, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape top');
			}, oMainShortcutTypes.checkMoveShapeTop);
		});

		QUnit.test('Test little move shape top', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.y, -1 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape top');
			}, oMainShortcutTypes.checkLittleMoveShapeTop);
		});

		QUnit.test('Test big move shape right', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.x, 5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape right');
			}, oMainShortcutTypes.checkMoveShapeRight);
		});

		QUnit.test('Test little move shape right', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.x, 1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape right');
			}, oMainShortcutTypes.checkLittleMoveShapeRight);
		});

		QUnit.test('Test big move shape left', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.x, -5 * AscCommon.g_dKoef_pix_to_mm, 'Check move shape left');
			}, oMainShortcutTypes.checkMoveShapeLeft);
		});

		QUnit.test('Test little move shape left', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const oShape = executeCheckMoveShape(oEvent);
				oAssert.strictEqual(oShape.x, -1 * AscCommon.g_dKoef_pix_to_mm, 'Check little move shape left');
			}, oMainShortcutTypes.checkLittleMoveShapeLeft);
		});

		//Delete
		QUnit.test('Test delete front symbol', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello world');
				moveToParagraph(oParagraph, true);

				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'ello world', 'Check delete front shortcut');
			}, oMainShortcutTypes.checkDeleteFront);
		});

		QUnit.test('Test delete front word', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello world');
				moveToParagraph(oParagraph, true);

				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(oParagraph), 'world', 'Check delete front word shortcut');
			}, oMainShortcutTypes.checkDeleteWordFront);
		});

		QUnit.test('Test increase indent', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello');
				oParagraph.Pr.SetInd(0, 0, 0);
				oParagraph.Set_PresentationLevel(0);
				moveToParagraph(oParagraph, true);

				onKeyDown(oEvent);
				const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
				oAssert.strictEqual(oParaPr.GetIndLeft(), 11.1125, 'Check increase indent');
			}, oMainShortcutTypes.checkIncreaseIndent);
		});

		QUnit.test('Test decrease indent', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				const {oParagraph} = getShapeWithParagraphHelper('Hello');
				oParagraph.Pr.SetInd(0, 12, 0);
				oParagraph.Set_PresentationLevel(1);
				moveToParagraph(oParagraph, true);

				onKeyDown(oEvent);
				const oParaPr = oGlobalLogicDocument.GetDirectParaPr();
				oAssert.true(AscFormat.fApproxEqual(oParaPr.GetIndLeft(), 0.8875), 'Check decrease indent');
			}, oMainShortcutTypes.checkDecreaseIndent);
		});

		QUnit.test('Test prevent default on num lock', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				onKeyDown(oEvent);
				oAssert.true(oEvent.isDefaultPrevented, 'Check prevent default on num lock');
			}, oMainShortcutTypes.checkNumLock);
		});

		QUnit.test('Test prevent default on scroll lock', (oAssert) =>
		{
			startMainTest((oEvent) =>
			{
				onKeyDown(oEvent);
				oAssert.true(oEvent.isDefaultPrevented, 'Check prevent default on scroll lock');
			}, oMainShortcutTypes.checkScrollLock);
		});
	});
})(window);
