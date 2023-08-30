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

	const oMainShortcutTypes = {
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
		checkMoveCursorToStartPosChartTitle              : 26,
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
		checkScrollLock                                  : 70
	};
	const oMainEvents = {};
	oMainEvents[oMainShortcutTypes.checkDeleteBack] = [new CTestEvent(createNativeEvent(8, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteWordBack] = [new CTestEvent(createNativeEvent(8, true, false, false, false, false, false), testWindows),
		new CTestEvent(createNativeEvent(8, false, false, true, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkRemoveAnimation] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveChart] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveShape] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveTable] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveGroup] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveShapeInGroup] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToNextCell] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToPreviousCell] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkIncreaseBulletIndent] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDecreaseBulletIndent] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTab] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectNextObject] = [new CTestEvent(createNativeEvent(9, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectPreviousObject] = [new CTestEvent(createNativeEvent(9, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkVisitHyperlink] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectNextObjectWithPlaceholder] = [new CTestEvent(createNativeEvent(13, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddNextSlideAfterSelectLastPlaceholderObject] = [new CTestEvent(createNativeEvent(13, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddBreakLine] = [new CTestEvent(createNativeEvent(13, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddMathBreakLine] = [
		new CTestEvent(createNativeEvent(13, false, true, false, false, false, false)),
		new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTitleBreakLine] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddParagraph] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddTxBodyShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorToStartPosShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectAllContentShape] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectAllContentChartTitle] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorToStartPosChartTitle] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkRemoveAndMoveToStartPosTable] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectFirstCellContent] = [new CTestEvent(createNativeEvent(13, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetAddShape] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetAllDrawingSelection] = [new CTestEvent(createNativeEvent(27, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkResetStepDrawingSelection] = [new CTestEvent(createNativeEvent(27, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkNonBreakingSpace] = [new CTestEvent(createNativeEvent(32, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkClearParagraphFormatting] = [new CTestEvent(createNativeEvent(32, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkAddSpace] = [new CTestEvent(createNativeEvent(32, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToEndPosContent] = [new CTestEvent(createNativeEvent(35, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToEndLineContent] = [new CTestEvent(createNativeEvent(35, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(39, true, false, false, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkSelectToEndLineContent] = [new CTestEvent(createNativeEvent(35, false, true, false, false, false, false)),
		new CTestEvent(createNativeEvent(39, true, true, false, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveToStartPosContent] = [new CTestEvent(createNativeEvent(36, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveToStartLineContent] = [new CTestEvent(createNativeEvent(36, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, true, false, false, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkSelectToStartLineContent] = [new CTestEvent(createNativeEvent(36, false, true, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, true, true, false, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveCursorLeft] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorLeft] = [new CTestEvent(createNativeEvent(37, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectWordCursorLeft] = [new CTestEvent(createNativeEvent(37, true, true, false, false, false, false), testWindows),
		new CTestEvent(createNativeEvent(37, false, true, true, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveCursorWordLeft] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false, false), testWindows),
		new CTestEvent(createNativeEvent(37, false, false, true, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveCursorLeftTable] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorRight] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorRightTable] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorRight] = [new CTestEvent(createNativeEvent(39, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectWordCursorRight] = [new CTestEvent(createNativeEvent(39, true, true, false, false, false, false), testWindows),
		new CTestEvent(createNativeEvent(39, false, true, true, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveCursorWordRight] = [new CTestEvent(createNativeEvent(39, true, false, false, false, false, false), testWindows),
		new CTestEvent(createNativeEvent(39, false, false, true, false, false, false), testMacOs)];
	oMainEvents[oMainShortcutTypes.checkMoveCursorTop] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorTopTable] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorTop] = [new CTestEvent(createNativeEvent(38, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorBottom] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveCursorBottomTable] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkSelectCursorBottom] = [new CTestEvent(createNativeEvent(40, false, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeBottom] = [new CTestEvent(createNativeEvent(40, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeBottom] = [new CTestEvent(createNativeEvent(40, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeTop] = [new CTestEvent(createNativeEvent(38, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeTop] = [new CTestEvent(createNativeEvent(38, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeRight] = [new CTestEvent(createNativeEvent(39, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeRight] = [new CTestEvent(createNativeEvent(39, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkMoveShapeLeft] = [new CTestEvent(createNativeEvent(37, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkLittleMoveShapeLeft] = [new CTestEvent(createNativeEvent(37, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteFront] = [new CTestEvent(createNativeEvent(46, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDeleteWordFront] = [new CTestEvent(createNativeEvent(46, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkIncreaseIndent] = [new CTestEvent(createNativeEvent(77, true, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkDecreaseIndent] = [new CTestEvent(createNativeEvent(77, true, true, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkNumLock] = [new CTestEvent(createNativeEvent(144, false, false, false, false, false, false))];
	oMainEvents[oMainShortcutTypes.checkScrollLock] = [new CTestEvent(createNativeEvent(145, false, false, false, false, false, false))];

	const oDemonstrationTypes = {
		moveToNextSlide          : 0,
		moveToPreviousSlide      : 1,
		moveToFirstSlide         : 2,
		moveToLastSlide          : 3,
		exitFromDemonstrationMode: 4
	};
	const oDemonstrationEvents = {};
	oDemonstrationEvents[oDemonstrationTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(13, false, false, false, false)),
		new CTestEvent(createNativeEvent(32, false, false, false, false)),
		new CTestEvent(createNativeEvent(34, false, false, false, false)),
		new CTestEvent(createNativeEvent(39, false, false, false, false)),
		new CTestEvent(createNativeEvent(40, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(33, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, false, false, false, false)),
		new CTestEvent(createNativeEvent(38, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false))
	];
	oDemonstrationEvents[oDemonstrationTypes.exitFromDemonstrationMode] = [
		new CTestEvent(createNativeEvent(27, false, false, false, false))
	];

	const oThumbnailsTypes = {
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
	const oThumbnailsEvents = {};
	oThumbnailsEvents[oThumbnailsTypes.addNextSlide] = [
		new CTestEvent(createNativeEvent(13, false, false, false, false)),
		new CTestEvent(createNativeEvent(77, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.removeSelectedSlides] = [
		new CTestEvent(createNativeEvent(8, false, false, false, false)),
		new CTestEvent(createNativeEvent(46, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToEnd] = [
		new CTestEvent(createNativeEvent(40, true, true, false, false)),
		new CTestEvent(createNativeEvent(34, true, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToNextPosition] = [
		new CTestEvent(createNativeEvent(40, true, false, false, false)),
		new CTestEvent(createNativeEvent(34, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectNextSlide] = [
		new CTestEvent(createNativeEvent(40, false, true, false, false)),
		new CTestEvent(createNativeEvent(34, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(40, true, false, false, false)),
		new CTestEvent(createNativeEvent(34, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToStart] = [
		new CTestEvent(createNativeEvent(33, true, true, false, false)),
		new CTestEvent(createNativeEvent(38, true, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveSelectedSlidesToPreviousPosition] = [
		new CTestEvent(createNativeEvent(33, true, false, false, false)),
		new CTestEvent(createNativeEvent(38, true, false, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.selectPreviousSlide] = [
		new CTestEvent(createNativeEvent(38, false, true, false, false)),
		new CTestEvent(createNativeEvent(33, false, true, false, false))
	];
	oThumbnailsEvents[oThumbnailsTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(33, true, false, false, false)),
		new CTestEvent(createNativeEvent(38, true, false, false, false))
	];

	const oThumbnailsMainFocusTypes = {
		addNextSlide                    : 0,
		moveToPreviousSlide             : 1,
		moveToNextSlide                 : 2,
		moveToFirstSlide                : 3,
		selectToFirstSlide              : 4,
		moveSelectedSlidesToEnd         : 5,
		moveSelectedSlidesToNextPosition: 6,
		moveToLastSlide                 : 7,
		selectToLastSlide               : 8
	};
	const oThumbnailsMainFocusEvents = {};
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.addNextSlide] = [
		new CTestEvent(createNativeEvent(77, true, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToPreviousSlide] = [
		new CTestEvent(createNativeEvent(38, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(37, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(33, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToNextSlide] = [
		new CTestEvent(createNativeEvent(39, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(40, false, false, false, false, false, false)),
		new CTestEvent(createNativeEvent(34, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.selectToFirstSlide] = [
		new CTestEvent(createNativeEvent(36, false, true, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.moveToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, false, false, false, false, false))
	];
	oThumbnailsMainFocusEvents[oThumbnailsMainFocusTypes.selectToLastSlide] = [
		new CTestEvent(createNativeEvent(35, false, true, false, false, false, false))
	];

	function privateStartTest(fCallback, nShortcutType, oTestEvents)
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

	function startThumbnailsMainFocusTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oThumbnailsMainFocusEvents);
	}

	function startMainTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oMainEvents);
	}

	function startThumbnailsFocusTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oThumbnailsEvents);
	}

	AscTestShortcut.startThumbnailsMainFocusTest = startThumbnailsMainFocusTest;
	AscTestShortcut.startMainTest = startMainTest;
	AscTestShortcut.startThumbnailsFocusTest = startThumbnailsFocusTest;
	AscTestShortcut.oThumbnailsMainFocusTypes = oThumbnailsMainFocusTypes;
	AscTestShortcut.oMainShortcutTypes = oMainShortcutTypes;
	AscTestShortcut.oThumbnailsTypes = oThumbnailsTypes;
	AscTestShortcut.oDemonstrationTypes = oDemonstrationTypes;
	AscTestShortcut.oDemonstrationEvents = oDemonstrationEvents;
})(window);
