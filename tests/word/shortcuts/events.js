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

(function (window)
{
	const {createNativeEvent} = AscTestShortcut;
	const testAll = 0;
	const testMacOs = 1;
	const testWindows = 2;
	function CTestEvent(oEvent, nType)
	{
		this.type = nType || testAll;
		this.event = oEvent;
	}
	const oTestTypes = {
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
		moveCursorToStartPositionTitleEnter: 21,
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
		selectLeftSymbol                   : 45,
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
		moveToNextHeader                   : 83
	};

	const oTestEvents = {};
	oTestEvents[oTestTypes.bigMoveGraphicObjectLeft] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false))];
	oTestEvents[oTestTypes.littleMoveGraphicObjectLeft] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false))];
	oTestEvents[oTestTypes.bigMoveGraphicObjectRight] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false))];
	oTestEvents[oTestTypes.littleMoveGraphicObjectRight] = [new CTestEvent(createNativeEvent(39, true, false, false, false, false))];
	oTestEvents[oTestTypes.bigMoveGraphicObjectDown] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false))];
	oTestEvents[oTestTypes.littleMoveGraphicObjectDown] = [new CTestEvent(createNativeEvent(40, true, false, false, false, false))];
	oTestEvents[oTestTypes.bigMoveGraphicObjectUp] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false))];
	oTestEvents[oTestTypes.littleMoveGraphicObjectUp] = [new CTestEvent(createNativeEvent(38, true, false, false, false, false))];
	oTestEvents[oTestTypes.removeBackSymbol] = [new CTestEvent(createNativeEvent(8, false, false, false, false))];
	oTestEvents[oTestTypes.removeBackWord] = [new CTestEvent(createNativeEvent(8, true, false, false, false))];
	oTestEvents[oTestTypes.removeShape] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false))
	];
	oTestEvents[oTestTypes.removeForm] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false))
	];
	oTestEvents[oTestTypes.moveToNextForm] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false))];
	oTestEvents[oTestTypes.moveToPreviousForm] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false))];
	oTestEvents[oTestTypes.handleTab] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false))];
	oTestEvents[oTestTypes.moveToNextCell] = [new CTestEvent(createNativeEvent(9, false, false, false, false))];
	oTestEvents[oTestTypes.moveToPreviousCell] = [new CTestEvent(createNativeEvent(9, false, true, false, false))];
	oTestEvents[oTestTypes.selectNextObject] = [new CTestEvent(createNativeEvent(9, false, false, false, false))];
	oTestEvents[oTestTypes.selectPreviousObject] = [new CTestEvent(createNativeEvent(9, false, true, false, false))];
	oTestEvents[oTestTypes.testIndent] = [new CTestEvent(createNativeEvent(9, false, false, false, false))];
	oTestEvents[oTestTypes.testUnIndent] = [new CTestEvent(createNativeEvent(9, false, true, false, false))];
	oTestEvents[oTestTypes.addTabToParagraph] = [new CTestEvent(createNativeEvent(9, false, false, false))];
	oTestEvents[oTestTypes.visitHyperlink] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.addBreakLineInlineLvlSdt] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.createTextBoxContent] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.createTextBody] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.addNewLineToMath] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.moveCursorToStartPositionShapeEnter] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.selectAllShapeEnter] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.moveCursorToStartPositionTitleEnter] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.selectAllInChartTitle] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false))];
	oTestEvents[oTestTypes.addNewParagraphContent] = [new CTestEvent(createNativeEvent(13, false, false, false, false))];
	oTestEvents[oTestTypes.addNewParagraphMath] = [new CTestEvent(createNativeEvent(13, false, false, false, false))];
	oTestEvents[oTestTypes.closeAllWindowsPopups] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.resetShapeSelection] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.resetStartAddShape] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.resetFormattingByExample] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.resetMarkerFormat] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.resetDragNDrop] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.endEditing] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false))];
	oTestEvents[oTestTypes.toggleCheckBox] = [new CTestEvent(createNativeEvent(32, false, false, false, false, false))];
	oTestEvents[oTestTypes.pageUp] = [new CTestEvent()];
	oTestEvents[oTestTypes.pageDown] = [new CTestEvent()];
	oTestEvents[oTestTypes.moveToEndDocument] = [new CTestEvent(createNativeEvent(35, true, false, false))];
	oTestEvents[oTestTypes.moveToEndLine] = [new CTestEvent(createNativeEvent(35, false, false, false, false))];
	oTestEvents[oTestTypes.selectToEndDocument] = [new CTestEvent(createNativeEvent(35, true, true, false, false))];
	oTestEvents[oTestTypes.selectToEndLine] = [new CTestEvent(createNativeEvent(35, false, true, false, false))];
	oTestEvents[oTestTypes.selectToStartLine] = [new CTestEvent(createNativeEvent(36, false, true, false, false))];
	oTestEvents[oTestTypes.selectToStartDocument] = [new CTestEvent(createNativeEvent(36, true, true, false, false))];
	oTestEvents[oTestTypes.moveToStartLine] = [new CTestEvent(createNativeEvent(36, false, false, false, false))];
	oTestEvents[oTestTypes.moveToStartDocument] = [new CTestEvent(createNativeEvent(36, true, false, false))];
	oTestEvents[oTestTypes.selectLeftWord] = [new CTestEvent(createNativeEvent(37, true, true, false, false))];
	oTestEvents[oTestTypes.moveToLeftWord] = [new CTestEvent(createNativeEvent(37, true, false, false, false))];
	oTestEvents[oTestTypes.selectLeftSymbol] = [new CTestEvent(createNativeEvent(37, false, true, false, false))];
	oTestEvents[oTestTypes.moveToLeftChar] = [new CTestEvent(createNativeEvent(37, false, false, false, false))];
	oTestEvents[oTestTypes.moveToRightChar] = [new CTestEvent(createNativeEvent(39, false, false, false, false))];
	oTestEvents[oTestTypes.selectRightChar] = [new CTestEvent(createNativeEvent(39, false, true, false, false))];
	oTestEvents[oTestTypes.moveToRightWord] = [new CTestEvent(createNativeEvent(39, true, false, false, false))];
	oTestEvents[oTestTypes.selectRightWord] = [new CTestEvent(createNativeEvent(39, true, true, false, false))];
	oTestEvents[oTestTypes.moveUp] = [new CTestEvent(createNativeEvent(38, false, false, false, false))];
	oTestEvents[oTestTypes.selectUp] = [new CTestEvent(createNativeEvent(38, false, true, false, false))];
	oTestEvents[oTestTypes.previousOptionComboBox] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false))];
	oTestEvents[oTestTypes.moveDown] = [new CTestEvent(createNativeEvent(40, false, false, false, false))];
	oTestEvents[oTestTypes.selectDown] = [new CTestEvent(createNativeEvent(40, false, true, false, false))];
	oTestEvents[oTestTypes.nextOptionComboBox] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false))];
	oTestEvents[oTestTypes.removeFrontSymbol] = [new CTestEvent(createNativeEvent(46, false, false, false, false))];
	oTestEvents[oTestTypes.removeFrontWord] = [new CTestEvent(createNativeEvent(46, true, false, false, false))];
	oTestEvents[oTestTypes.unicodeToChar] = [
		new CTestEvent(createNativeEvent(88, false, false, true, false), testWindows),
		new CTestEvent(createNativeEvent(88, true, false, true, false), testMacOs)
	];
	oTestEvents[oTestTypes.showContextMenu] = [
		new CTestEvent(createNativeEvent(93, false, false, false, false)),
		new CTestEvent(createNativeEvent(57351, false, false, false, false)),
		new CTestEvent(createNativeEvent(121, false, true, false, false))
	];
	oTestEvents[oTestTypes.disableNumLock] = [new CTestEvent(createNativeEvent(144, false, false, false, false))];
	oTestEvents[oTestTypes.disableScrollLock] = [new CTestEvent(createNativeEvent(145, false, false, false, false))];
	oTestEvents[oTestTypes.addSJKSpace] = [new CTestEvent(createNativeEvent(12288, false, false, false, false))];
	oTestEvents[oTestTypes.moveToStartPreviousPage] = [new CTestEvent(createNativeEvent(33, true, false, true, false))];
	oTestEvents[oTestTypes.moveToPreviousPage] = [new CTestEvent(createNativeEvent(33, false, false, false, false))];
	oTestEvents[oTestTypes.moveToPreviousHeaderFooter] = [new CTestEvent(createNativeEvent(33, false, false, false, false))];
	oTestEvents[oTestTypes.moveToPreviousHeader] = [
		new CTestEvent(createNativeEvent(33, true, false, true, false)),
		new CTestEvent(createNativeEvent(33, false, false, true, false))
	];
	oTestEvents[oTestTypes.selectToStartPreviousPage] = [new CTestEvent(createNativeEvent(33, true, true, false, false))];
	oTestEvents[oTestTypes.selectToPreviousPage] = [new CTestEvent(createNativeEvent(33, false, true, false, false))];
	oTestEvents[oTestTypes.moveToStartNextPage] = [new CTestEvent(createNativeEvent(34, true, false, true, false))];
	oTestEvents[oTestTypes.moveToNextPage] = [new CTestEvent(createNativeEvent(34, false, false, false, false))];
	oTestEvents[oTestTypes.moveToNextHeaderFooter] = [new CTestEvent(createNativeEvent(34, false, false, false, false))];
	oTestEvents[oTestTypes.moveToNextHeader] = [
		new CTestEvent(createNativeEvent(34, true, false, true, false)),
		new CTestEvent(createNativeEvent(34, false, false, true, false))
	];
	oTestEvents[oTestTypes.selectToStartNextPage] = [new CTestEvent(createNativeEvent(34, true, true, false, false))];
	oTestEvents[oTestTypes.selectToNextPage] = [new CTestEvent(createNativeEvent(34, false, true, false, false))];

	function startTest(fCallback, nShortcutType)
	{
		const arrTestEvents = oTestEvents[nShortcutType];

		for (let i = 0; i < arrTestEvents.length; i += 1)
		{
			const nTestType = arrTestEvents[i].type;
			if (nTestType === testAll)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);

				AscCommon.AscBrowser.isMacOs = false;
				fCallback(arrTestEvents[i].event);
			} else if (nTestType === testMacOs)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);
				AscCommon.AscBrowser.isMacOs = false;
			} else if (nTestType === testWindows)
			{
				fCallback(arrTestEvents[i].event);
			}
		}
	}
	AscTestShortcut.oTestTypes = oTestTypes;
	AscTestShortcut.startTest = startTest;
})(window);
