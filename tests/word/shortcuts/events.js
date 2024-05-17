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
	window.AscTestShortcut = {};
	const testFlags = {
		nothing: 0x0000,
		macOs: 0x0001
	};
	const testHotkeyActions = {
		removeBackSymbol                   : 0,
		removeBackWord                     : 1,
		removeShape                        : 2,
		removeForm                         : 3,
		moveToNextForm                     : 4,
		moveToPreviousForm                 : 5,
		handleTab                          : 6,
		moveToNextCell                     : 7,
		moveToPreviousCell                 : 8,
		selectNextObject                   : 9,
		selectPreviousObject               : 10,
		testIndent                         : 11,
		testUnIndent                       : 12,
		addTabToParagraph                  : 13,
		visitHyperlink                     : 14,
		addBreakLineInlineLvlSdt           : 15,
		createTextBoxContent               : 16,
		createTextBody                     : 17,
		addNewLineToMath                   : 18,
		moveCursorToStartPositionShapeEnter: 19,
		selectAllShapeEnter                : 20,
		selectAllInChartTitle              : 22,
		addNewParagraphContent             : 23,
		addNewParagraphMath                : 24,
		closeAllWindowsPopups              : 25,
		resetShapeSelection                : 26,
		resetStartAddShape                 : 27,
		resetFormattingByExample           : 28,
		resetMarkerFormat                  : 29,
		resetDragNDrop                     : 30,
		endEditing                         : 31,
		toggleCheckBox                     : 32,
		pageUp                             : 33,
		pageDown                           : 34,
		moveToEndDocument                  : 35,
		moveToEndLine                      : 36,
		selectToEndDocument                : 37,
		selectToEndLine                    : 38,
		selectToStartLine                  : 39,
		selectToStartDocument              : 40,
		moveToStartLine                    : 41,
		moveToStartDocument                : 42,
		selectLeftWord                     : 43,
		moveToLeftWord                     : 44,
		selectLeftChar                     : 45,
		moveToLeftChar                     : 46,
		moveToRightChar                    : 47,
		selectRightChar                    : 48,
		moveToRightWord                    : 49,
		selectRightWord                    : 50,
		moveUp                             : 51,
		selectUp                           : 52,
		previousOptionComboBox             : 53,
		moveDown                           : 54,
		selectDown                         : 55,
		nextOptionComboBox                 : 56,
		removeFrontSymbol                  : 57,
		removeFrontWord                    : 58,
		unicodeToChar                      : 59,
		showContextMenu                    : 60,
		disableNumLock                     : 61,
		disableScrollLock                  : 62,
		addSJKSpace                        : 63,
		bigMoveGraphicObjectLeft           : 64,
		littleMoveGraphicObjectLeft        : 65,
		bigMoveGraphicObjectRight          : 66,
		littleMoveGraphicObjectRight       : 67,
		bigMoveGraphicObjectDown           : 68,
		littleMoveGraphicObjectDown        : 69,
		bigMoveGraphicObjectUp             : 70,
		littleMoveGraphicObjectUp          : 71,
		moveToPreviousPage                 : 72,
		selectToPreviousPage               : 73,
		moveToStartPreviousPage            : 74,
		selectToStartPreviousPage          : 75,
		moveToPreviousHeaderFooter         : 76,
		moveToPreviousHeader               : 77,
		moveToNextPage                     : 78,
		selectToNextPage                   : 79,
		moveToStartNextPage                : 80,
		selectToStartNextPage              : 81,
		moveToNextHeaderFooter             : 82,
		moveToNextHeader                   : 83,
		disableBrowserZoomIn               : 84,
	};

	const testHotkeyEvents = {};
	testHotkeyEvents[testHotkeyActions.bigMoveGraphicObjectLeft] = [CreateTestEvent(37, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.littleMoveGraphicObjectLeft] = [CreateTestEvent(37, true, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.bigMoveGraphicObjectRight] = [CreateTestEvent(39, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.littleMoveGraphicObjectRight] = [CreateTestEvent(39, true, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.bigMoveGraphicObjectDown] = [CreateTestEvent(40, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.littleMoveGraphicObjectDown] = [CreateTestEvent(40, true, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.bigMoveGraphicObjectUp] = [CreateTestEvent(38, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.littleMoveGraphicObjectUp] = [CreateTestEvent(38, true, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.removeBackSymbol] = [CreateTestEvent(8, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.removeBackWord] = [
		CreateTestEvent(8, true, false, false, false),
		CreateTestEvent(8, false, false, true, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.removeShape] = [
		CreateTestEvent(8, false, false, false, false, false),
		CreateTestEvent(46, false, false, false, false, false)
	];
	testHotkeyEvents[testHotkeyActions.removeForm] = [
		CreateTestEvent(8, false, false, false, false, false),
		CreateTestEvent(46, false, false, false, false, false)
	];
	testHotkeyEvents[testHotkeyActions.moveToNextForm] = [CreateTestEvent(9, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToPreviousForm] = [CreateTestEvent(9, false, true, false, false, false)];
	testHotkeyEvents[testHotkeyActions.handleTab] = [CreateTestEvent(9, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToNextCell] = [CreateTestEvent(9, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToPreviousCell] = [CreateTestEvent(9, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.selectNextObject] = [CreateTestEvent(9, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectPreviousObject] = [CreateTestEvent(9, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.testIndent] = [CreateTestEvent(9, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.testUnIndent] = [CreateTestEvent(9, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.addTabToParagraph] = [CreateTestEvent(9, false, false, false)];
	testHotkeyEvents[testHotkeyActions.visitHyperlink] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.addBreakLineInlineLvlSdt] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.createTextBoxContent] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.createTextBody] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.addNewLineToMath] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveCursorToStartPositionShapeEnter] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectAllShapeEnter] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectAllInChartTitle] = [CreateTestEvent(13, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.addNewParagraphContent] = [CreateTestEvent(13, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.addNewParagraphMath] = [CreateTestEvent(13, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.closeAllWindowsPopups] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.resetShapeSelection] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.resetStartAddShape] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.resetFormattingByExample] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.resetMarkerFormat] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.resetDragNDrop] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.endEditing] = [CreateTestEvent(27, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.toggleCheckBox] = [CreateTestEvent(32, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToEndDocument] = [CreateTestEvent(35, true, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToEndLine] = [
		CreateTestEvent(35, false, false, false, false),
		CreateTestEvent(39, true, false, false, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.selectToEndDocument] = [CreateTestEvent(35, true, true, false, false)];
	testHotkeyEvents[testHotkeyActions.selectToEndLine] = [
		CreateTestEvent(35, false, true, false, false),
		CreateTestEvent(39, true, true, false, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.selectToStartLine] = [
		CreateTestEvent(36, false, true, false, false),
		CreateTestEvent(37, true, true, false, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.selectToStartDocument] = [CreateTestEvent(36, true, true, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToStartLine] = [
		CreateTestEvent(36, false, false, false, false),
		CreateTestEvent(37, true, false, false, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.moveToStartDocument] = [CreateTestEvent(36, true, false, false)];
	testHotkeyEvents[testHotkeyActions.selectLeftWord] = [
		CreateTestEvent(37, true, true, false, false),
		CreateTestEvent(37, false, true, true, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.moveToLeftWord] = [
		CreateTestEvent(37, true, false, false, false),
		CreateTestEvent(37, false, false, true, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.selectLeftChar] = [CreateTestEvent(37, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToLeftChar] = [CreateTestEvent(37, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToRightChar] = [CreateTestEvent(39, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectRightChar] = [CreateTestEvent(39, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToRightWord] = [
		CreateTestEvent(39, true, false, false, false),
		CreateTestEvent(39, false, false, true, false, testFlags.macOs)
	];
	testHotkeyEvents[testHotkeyActions.selectRightWord] = [
		CreateTestEvent(39, true, true, false, false),
		CreateTestEvent(39, false, true, true, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.moveUp] = [CreateTestEvent(38, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectUp] = [CreateTestEvent(38, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.previousOptionComboBox] = [CreateTestEvent(38, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveDown] = [CreateTestEvent(40, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.selectDown] = [CreateTestEvent(40, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.nextOptionComboBox] = [CreateTestEvent(40, false, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.removeFrontSymbol] = [CreateTestEvent(46, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.removeFrontWord] = [
		CreateTestEvent(46, true, false, false, false),
		CreateTestEvent(46, false, false, true, false, testFlags.macOs),
	];
	testHotkeyEvents[testHotkeyActions.unicodeToChar] = [
		CreateTestEvent(88, false, false, true, false),
		CreateTestEvent(88, true, false, true, false, testFlags.macOs)
	];
	testHotkeyEvents[testHotkeyActions.showContextMenu] = [
		CreateTestEvent(93, false, false, false, false),
		CreateTestEvent(57351, false, false, false, false),
		CreateTestEvent(121, false, true, false, false)
	];
	testHotkeyEvents[testHotkeyActions.disableNumLock] = [CreateTestEvent(144, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.disableScrollLock] = [CreateTestEvent(145, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.addSJKSpace] = [CreateTestEvent(12288, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToStartPreviousPage] = [CreateTestEvent(33, true, false, true, false)];
	testHotkeyEvents[testHotkeyActions.moveToPreviousPage] = [CreateTestEvent(33, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToPreviousHeaderFooter] = [CreateTestEvent(33, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToPreviousHeader] = [
		CreateTestEvent(33, true, false, true, false),
		CreateTestEvent(33, false, false, true, false)
	];
	testHotkeyEvents[testHotkeyActions.selectToStartPreviousPage] = [CreateTestEvent(33, true, true, false, false)];
	testHotkeyEvents[testHotkeyActions.selectToPreviousPage] = [CreateTestEvent(33, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToStartNextPage] = [CreateTestEvent(34, true, false, true, false)];
	testHotkeyEvents[testHotkeyActions.moveToNextPage] = [CreateTestEvent(34, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToNextHeaderFooter] = [CreateTestEvent(34, false, false, false, false)];
	testHotkeyEvents[testHotkeyActions.moveToNextHeader] = [
		CreateTestEvent(34, true, false, true, false),
		CreateTestEvent(34, false, false, true, false)
	];
	testHotkeyEvents[testHotkeyActions.selectToStartNextPage] = [CreateTestEvent(34, true, true, false, false)];
	testHotkeyEvents[testHotkeyActions.selectToNextPage] = [CreateTestEvent(34, false, true, false, false)];
	testHotkeyEvents[testHotkeyActions.disableBrowserZoomIn] = [CreateTestEvent(187, true, true, false, false)];

function CreateTestEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey, flags)
{
	flags = flags || testFlags.nothing;
	const event = new AscCommon.CKeyboardEvent();
	event.KeyCode = nKeyCode;
	event.CtrlKey = !!bIsCtrl;
	event.ShiftKey = !!bIsShift;
	event.AltKey = !!bIsAlt;
	event.MacCmdKey = !!bIsMetaKey;
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

	function ExecuteShortcut(e)
	{
		e.enableFlags && e.enableFlags();
		const res = editor.WordControl.m_oLogicDocument.OnKeyDown(e);
		e.disableFlags && e.disableFlags();
		return res;
	}

	function ExecuteHotkey(type, eventIndex)
	{
		const event = testHotkeyEvents[type][eventIndex || 0];
		return ExecuteShortcut(event);
	}

	AscTestShortcut.testHotkeyActions = testHotkeyActions;
	AscTestShortcut.testHotkeyEvents = testHotkeyEvents;
	AscTestShortcut.ExecuteShortcut = ExecuteShortcut;
	AscTestShortcut.ExecuteHotkey = ExecuteHotkey;
})(window);
