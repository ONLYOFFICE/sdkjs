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


(function (window)
{
	const testFlags = {
		nothing: 0x0000,
		macOs: 0x0001
	};

	const mainShortcutTypes = {
		checkDeleteBack                                  : 0,
		checkDeleteWordBack                              : 1,
		checkRemoveAnimation                             : 2,
		checkRemoveChart                                 : 3,
		checkRemoveShape                                 : 4,
		checkRemoveTable                                 : 5,
		checkRemoveGroup                                 : 6,
		checkRemoveShapeInGroup                          : 7,
		checkMoveToNextCell                              : 8,
		checkMoveToPreviousCell                          : 9,
		checkIncreaseBulletIndent                        : 10,
		checkDecreaseBulletIndent                        : 11,
		checkAddTab                                      : 12,
		checkSelectNextObject                            : 13,
		checkSelectPreviousObject                        : 14,
		checkVisitHyperlink                              : 15,
		checkSelectNextObjectWithPlaceholder             : 16,
		checkAddNextSlideAfterSelectLastPlaceholderObject: 17,
		checkAddBreakLine                                : 18,
		checkAddTitleBreakLine                           : 19,
		checkAddMathBreakLine                            : 20,
		checkAddParagraph                                : 21,
		checkAddTxBodyShape                              : 22,
		checkMoveCursorToStartPosShape                   : 23,
		checkSelectAllContentShape                       : 24,
		checkSelectAllContentChartTitle                  : 25,
		checkRemoveAndMoveToStartPosTable                : 27,
		checkSelectFirstCellContent                      : 28,
		checkResetAddShape                               : 29,
		checkResetAllDrawingSelection                    : 30,
		checkResetStepDrawingSelection                   : 31,
		checkNonBreakingSpace                            : 32,
		checkClearParagraphFormatting                    : 33,
		checkAddSpace                                    : 34,
		checkMoveToEndPosContent                         : 35,
		checkMoveToEndLineContent                        : 36,
		checkSelectToEndLineContent                      : 37,
		checkMoveToStartPosContent                       : 38,
		checkMoveToStartLineContent                      : 39,
		checkSelectToStartLineContent                    : 40,
		checkMoveCursorLeft                              : 41,
		checkSelectCursorLeft                            : 42,
		checkSelectWordCursorLeft                        : 43,
		checkMoveCursorWordLeft                          : 44,
		checkMoveCursorLeftTable                         : 45,
		checkMoveCursorRight                             : 46,
		checkMoveCursorRightTable                        : 47,
		checkSelectCursorRight                           : 48,
		checkSelectWordCursorRight                       : 49,
		checkMoveCursorWordRight                         : 50,
		checkMoveCursorTop                               : 51,
		checkMoveCursorTopTable                          : 52,
		checkSelectCursorTop                             : 53,
		checkMoveCursorBottom                            : 54,
		checkMoveCursorBottomTable                       : 55,
		checkSelectCursorBottom                          : 56,
		checkMoveShapeBottom                             : 57,
		checkLittleMoveShapeBottom                       : 58,
		checkMoveShapeTop                                : 59,
		checkLittleMoveShapeTop                          : 60,
		checkMoveShapeRight                              : 61,
		checkLittleMoveShapeRight                        : 62,
		checkMoveShapeLeft                               : 63,
		checkLittleMoveShapeLeft                         : 64,
		checkDeleteFront                                 : 65,
		checkDeleteWordFront                             : 66,
		checkIncreaseIndent                              : 67,
		checkDecreaseIndent                              : 68,
		checkNumLock                                     : 69,
		checkScrollLock                                  : 70,
		addNextSlide                                     : 71,
		moveToPreviousSlide                              : 72,
		moveToNextSlide                                  : 73,
		moveToFirstSlide                                 : 74,
		selectToFirstSlide                               : 75,
		moveToLastSlide                                  : 76,
		selectToLastSlide                                : 77,
		disableBrowserZoomIn                             : 78,
	};
	const mainEvents = {};
	mainEvents[mainShortcutTypes.checkDeleteBack] = [CreateKeyboardEvent(8, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkDeleteWordBack] = [
		CreateKeyboardEvent(8, true, false, false, false),
		CreateKeyboardEvent(8, false, false, true, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkRemoveAnimation] = [
		CreateKeyboardEvent(8, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveChart] = [
		CreateKeyboardEvent(8, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveShape] = [
		CreateKeyboardEvent(8, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveTable] = [
		CreateKeyboardEvent(8, false, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveGroup] = [
		CreateKeyboardEvent(8, false, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveShapeInGroup] = [
		CreateKeyboardEvent(8, false, false, false, false),
		CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveToNextCell] = [CreateKeyboardEvent(9, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveToPreviousCell] = [CreateKeyboardEvent(9, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkIncreaseBulletIndent] = [CreateKeyboardEvent(9, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkDecreaseBulletIndent] = [CreateKeyboardEvent(9, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkAddTab] = [CreateKeyboardEvent(9, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectNextObject] = [CreateKeyboardEvent(9, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectPreviousObject] = [CreateKeyboardEvent(9, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkVisitHyperlink] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectNextObjectWithPlaceholder] = [CreateKeyboardEvent(13, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddNextSlideAfterSelectLastPlaceholderObject] = [CreateKeyboardEvent(13, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddBreakLine] = [CreateKeyboardEvent(13, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkAddMathBreakLine] = [
		CreateKeyboardEvent(13, false, true, false, false),
		CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddTitleBreakLine] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddParagraph] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddTxBodyShape] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorToStartPosShape] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectAllContentShape] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectAllContentChartTitle] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkRemoveAndMoveToStartPosTable] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectFirstCellContent] = [CreateKeyboardEvent(13, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkResetAddShape] = [CreateKeyboardEvent(27, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkResetAllDrawingSelection] = [CreateKeyboardEvent(27, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkResetStepDrawingSelection] = [CreateKeyboardEvent(27, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkNonBreakingSpace] = [CreateKeyboardEvent(32, true, true, false, false)];
	mainEvents[mainShortcutTypes.checkClearParagraphFormatting] = [CreateKeyboardEvent(32, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkAddSpace] = [CreateKeyboardEvent(32, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveToEndPosContent] = [CreateKeyboardEvent(35, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveToEndLineContent] = [
		CreateKeyboardEvent(35, false, false, false, false),
		CreateKeyboardEvent(39, true, false, false, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkSelectToEndLineContent] = [
		CreateKeyboardEvent(35, false, true, false, false),
		CreateKeyboardEvent(39, true, true, false, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkMoveToStartPosContent] = [CreateKeyboardEvent(36, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveToStartLineContent] = [
		CreateKeyboardEvent(36, false, false, false, false),
		CreateKeyboardEvent(37, true, false, false, false, testFlags.macOs)
	];
	mainEvents[mainShortcutTypes.checkSelectToStartLineContent] = [
		CreateKeyboardEvent(36, false, true, false, false),
		CreateKeyboardEvent(37, true, true, false, false, testFlags.macOs)
	];
	mainEvents[mainShortcutTypes.checkMoveCursorLeft] = [CreateKeyboardEvent(37, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectCursorLeft] = [CreateKeyboardEvent(37, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkSelectWordCursorLeft] = [
		CreateKeyboardEvent(37, true, true, false, false),
		CreateKeyboardEvent(37, false, true, true, false, testFlags.macOs)
	];
	mainEvents[mainShortcutTypes.checkMoveCursorWordLeft] = [
		CreateKeyboardEvent(37, true, false, false, false),
		CreateKeyboardEvent(37, false, false, true, false, testFlags.macOs)
	];
	mainEvents[mainShortcutTypes.checkMoveCursorLeftTable] = [CreateKeyboardEvent(37, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorRight] = [CreateKeyboardEvent(39, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorRightTable] = [CreateKeyboardEvent(39, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectCursorRight] = [CreateKeyboardEvent(39, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkSelectWordCursorRight] = [
		CreateKeyboardEvent(39, true, true, false, false),
		CreateKeyboardEvent(39, false, true, true, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkMoveCursorWordRight] = [
		CreateKeyboardEvent(39, true, false, false, false),
		CreateKeyboardEvent(39, false, false, true, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkMoveCursorTop] = [CreateKeyboardEvent(38, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorTopTable] = [CreateKeyboardEvent(38, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectCursorTop] = [CreateKeyboardEvent(38, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorBottom] = [CreateKeyboardEvent(40, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveCursorBottomTable] = [CreateKeyboardEvent(40, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkSelectCursorBottom] = [CreateKeyboardEvent(40, false, true, false, false)];
	mainEvents[mainShortcutTypes.checkMoveShapeBottom] = [CreateKeyboardEvent(40, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkLittleMoveShapeBottom] = [CreateKeyboardEvent(40, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveShapeTop] = [CreateKeyboardEvent(38, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkLittleMoveShapeTop] = [CreateKeyboardEvent(38, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveShapeRight] = [CreateKeyboardEvent(39, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkLittleMoveShapeRight] = [CreateKeyboardEvent(39, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkMoveShapeLeft] = [CreateKeyboardEvent(37, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkLittleMoveShapeLeft] = [CreateKeyboardEvent(37, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkDeleteFront] = [CreateKeyboardEvent(46, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkDeleteWordFront] = [
		CreateKeyboardEvent(46, true, false, false, false),
		CreateKeyboardEvent(46, false, false, true, false, testFlags.macOs),
	];
	mainEvents[mainShortcutTypes.checkIncreaseIndent] = [CreateKeyboardEvent(77, true, false, false, false)];
	mainEvents[mainShortcutTypes.checkDecreaseIndent] = [CreateKeyboardEvent(77, true, true, false, false)];
	mainEvents[mainShortcutTypes.checkNumLock] = [CreateKeyboardEvent(144, false, false, false, false)];
	mainEvents[mainShortcutTypes.checkScrollLock] = [CreateKeyboardEvent(145, false, false, false, false)];
	mainEvents[mainShortcutTypes.addNextSlide] = [
		CreateKeyboardEvent(77, true, false, false, false)
	];
	mainEvents[mainShortcutTypes.moveToPreviousSlide] = [
		CreateKeyboardEvent(38, false, false, false, false),
		CreateKeyboardEvent(37, false, false, false, false),
		CreateKeyboardEvent(33, false, false, false, false)
	];
	mainEvents[mainShortcutTypes.moveToNextSlide] = [
		CreateKeyboardEvent(39, false, false, false, false),
		CreateKeyboardEvent(40, false, false, false, false),
		CreateKeyboardEvent(34, false, false, false, false)
	];
	mainEvents[mainShortcutTypes.moveToFirstSlide] = [
		CreateKeyboardEvent(36, false, false, false, false)
	];
	mainEvents[mainShortcutTypes.selectToFirstSlide] = [
		CreateKeyboardEvent(36, false, true, false, false)
	];
	mainEvents[mainShortcutTypes.moveToLastSlide] = [
		CreateKeyboardEvent(35, false, false, false, false)
	];
	mainEvents[mainShortcutTypes.selectToLastSlide] = [
		CreateKeyboardEvent(35, false, true, false, false)
	];
	mainEvents[mainShortcutTypes.disableBrowserZoomIn] = [
		CreateKeyboardEvent(187, true, true, false, false)
	];

	const demonstrationTypes = {
		moveToNextSlide          : 0,
		moveToPreviousSlide      : 1,
		moveToFirstSlide         : 2,
		moveToLastSlide          : 3,
		exitFromDemonstrationMode: 4,
		moveToNumberSlide        : 5
	};
	const demonstrationEvents = {};
	demonstrationEvents[demonstrationTypes.moveToNextSlide] = [
		new CNativeEvent(13, false, false, false, false),
		new CNativeEvent(32, false, false, false, false),
		new CNativeEvent(34, false, false, false, false),
		new CNativeEvent(39, false, false, false, false),
		new CNativeEvent(40, false, false, false, false)
	];
	demonstrationEvents[demonstrationTypes.moveToNumberSlide] = [
		new CNativeEvent(13, false, false, false, false),
	];
	demonstrationEvents[demonstrationTypes.moveToPreviousSlide] = [
		new CNativeEvent(33, false, false, false, false),
		new CNativeEvent(37, false, false, false, false),
		new CNativeEvent(38, false, false, false, false)
	];
	demonstrationEvents[demonstrationTypes.moveToFirstSlide] = [
		new CNativeEvent(36, false, false, false, false)
	];
	demonstrationEvents[demonstrationTypes.moveToLastSlide] = [
		new CNativeEvent(35, false, false, false, false)
	];
	demonstrationEvents[demonstrationTypes.exitFromDemonstrationMode] = [
		new CNativeEvent(27, false, false, false, false)
	];

	const thumbnailsTypes = {
		addNextSlide                        : 0,
		removeSelectedSlides                : 1,
		moveSelectedSlidesToEnd             : 2,
		moveSelectedSlidesToNextPosition    : 3,
		selectNextSlide                     : 4,
		moveToNextSlide                     : 5,
		moveToFirstSlide                    : 6,
		selectToFirstSlide                  : 7,
		moveToLastSlide                     : 8,
		selectToLastSlide                   : 9,
		moveSelectedSlidesToStart           : 10,
		moveSelectedSlidesToPreviousPosition: 11,
		selectPreviousSlide                 : 12,
		moveToPreviousSlide                 : 13
	};
	const thumbnailsEvents = {};
	thumbnailsEvents[thumbnailsTypes.addNextSlide] = [
		new CNativeEvent(13, false, false, false, false),
		new CNativeEvent(77, true, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.removeSelectedSlides] = [
		new CNativeEvent(8, false, false, false, false),
		new CNativeEvent(46, false, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveSelectedSlidesToEnd] = [
		new CNativeEvent(40, true, true, false, false),
		new CNativeEvent(34, true, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveSelectedSlidesToNextPosition] = [
		new CNativeEvent(40, true, false, false, false),
		new CNativeEvent(34, true, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.selectNextSlide] = [
		new CNativeEvent(40, false, true, false, false),
		new CNativeEvent(34, false, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveToNextSlide] = [
		new CNativeEvent(40, true, false, false, false),
		new CNativeEvent(34, true, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveToFirstSlide] = [
		new CNativeEvent(36, false, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.selectToFirstSlide] = [
		new CNativeEvent(36, false, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveToLastSlide] = [
		new CNativeEvent(35, false, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.selectToLastSlide] = [
		new CNativeEvent(35, false, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveSelectedSlidesToStart] = [
		new CNativeEvent(33, true, true, false, false),
		new CNativeEvent(38, true, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveSelectedSlidesToPreviousPosition] = [
		new CNativeEvent(33, true, false, false, false),
		new CNativeEvent(38, true, false, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.selectPreviousSlide] = [
		new CNativeEvent(38, false, true, false, false),
		new CNativeEvent(33, false, true, false, false)
	];
	thumbnailsEvents[thumbnailsTypes.moveToPreviousSlide] = [
		new CNativeEvent(33, true, false, false, false),
		new CNativeEvent(38, true, false, false, false)
	];


	function CreateKeyboardEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey, flags)
	{
		flags = flags || testFlags.nothing;
		const event = new AscCommon.CKeyboardEvent();
		event.KeyCode = nKeyCode;
		event.CtrlKey = !!bIsCtrl;
		event.ShiftKey = !!bIsShift;
		event.AltKey = !!bIsAlt;
		event.MacCmdKey = !!bIsMetaKey;
		event.isDefaultPrevented = false;
		event.isPropagationStopped = false;
		event.preventDefault = function ()
		{
			event.isDefaultPrevented = true;
		}
		event.stopPropagation = function ()
		{
			event.isPropagationStopped = true;
		}
		event.enableFlags = function ()
		{
			if (flags & testFlags.macOs)
			{
				AscCommon.AscBrowser.isMacOs = true;
			}
		}
		event.disableFlags = function ()
		{
			if (flags & testFlags.macOs)
			{
				AscCommon.AscBrowser.isMacOs = false;
			}
		}
		return event;
	}

	function CNativeEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey)
	{
		this.keyCode = nKeyCode;
		this.ctrlKey = !!bIsCtrl;
		this.shiftKey = !!bIsShift;
		this.altKey = !!bIsAlt;
		this.metaKey = !!bIsMetaKey;
		this.isDefaultPrevented = false;
		this.isPropagationStopped = false;
	}

	CNativeEvent.prototype.preventDefault = function ()
	{
		this.isDefaultPrevented = true;
	};
	CNativeEvent.prototype.stopPropagation = function ()
	{
		this.isPropagationStopped = true;
	};

	editor.getShortcut = function (e)
	{
		if (typeof e === 'number')
		{
			return e;
		}
	};

	function ExecuteMainHotkey(type, index)
	{
		const event = mainEvents[type][index || 0];
		return ExecuteMainShortcut(event);
	}

	function ExecuteThumbnailHotkey(e, index)
	{
		const event = thumbnailsEvents[e][index || 0];
		return ExecuteThumbnailShortcut(event);
	}

	function ExecuteMainShortcut(e)
	{
		e.enableFlags && e.enableFlags();
		const res = editor.WordControl.m_oLogicDocument.OnKeyDown(e);
		e.disableFlags && e.disableFlags();
		return res;
	}

	function ExecuteThumbnailShortcut(e)
	{
		const OldGetShortcut = editor.getShortcut;
		if (typeof e === 'number')
		{
			const shortcutType = e;
			editor.getShortcut = function ()
			{
				return shortcutType;
			};
			e = new CNativeEvent();
		}
		const res = editor.WordControl.Thumbnails.onKeyDown(e);
		editor.getShortcut = OldGetShortcut;
		return res;
	}

	function ExecuteDemonstrationShortcut(e)
	{
		return editor.WordControl.DemonstrationManager.onKeyDown(e);
	}

	AscTest.mainShortcutTypes = mainShortcutTypes;
	AscTest.thumbnailsTypes = thumbnailsTypes;
	AscTest.demonstrationTypes = demonstrationTypes;
	AscTest.demonstrationEvents = demonstrationEvents;
	AscTest.ExecuteDemonstrationShortcut = ExecuteDemonstrationShortcut;
	AscTest.ExecuteThumbnailShortcut = ExecuteThumbnailShortcut;
	AscTest.ExecuteMainShortcut = ExecuteMainShortcut;
	AscTest.ExecuteThumbnailHotkey = ExecuteThumbnailHotkey;
	AscTest.ExecuteMainHotkey = ExecuteMainHotkey;
	AscTest.CNativeEvent = CNativeEvent;
})(window);
