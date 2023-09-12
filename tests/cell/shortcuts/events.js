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
	const createEvent = AscTestShortcut.createEvent;
	const cleanGraphic = AscTestShortcut.cleanGraphic;
	const testAll = 0;
	const testWindows = 1;
	const testMacOs = 2;

	function CTestEvent(oEvent, nType)
	{
		this.type = nType || testAll;
		this.event = oEvent;
	}

	const oKeyCode =
		{
			BackSpace       : 8,
			Tab             : 9,
			Enter           : 13,
			Esc             : 27,
			End             : 35,
			Home            : 36,
			ArrowLeft       : 37,
			ArrowTop        : 38,
			ArrowRight      : 39,
			ArrowBottom     : 40,
			Delete          : 46,
			A               : 65,
			B               : 66,
			C               : 67,
			E               : 69,
			I               : 73,
			J               : 74,
			K               : 75,
			L               : 76,
			M               : 77,
			P               : 80,
			R               : 82,
			S               : 83,
			U               : 85,
			V               : 86,
			X               : 88,
			Y               : 89,
			Z               : 90,
			OperaContextMenu: 57351,
			F10             : 121,
			NumLock         : 144,
			ScrollLock      : 145,
			Equal           : 187,
			EqualFirefox    : 61,
			Comma           : 188,
			Minus           : 189,
			Period          : 190,
			BracketLeft     : 219,
			BracketRight    : 221,
			F2              : 113
		}


	function privateStartTest(fCallback, nShortcutType, oEvents)
	{
		cleanGraphic();
		const arrTestEvents = oEvents[nShortcutType];

		for (let i = 0; i < arrTestEvents.length; i += 1)
		{
			cleanGraphic();
			const nTestType = arrTestEvents[i].type;
			if (nTestType === testAll)
			{
				AscCommon.AscBrowser.isMacOs = true;
				fCallback(arrTestEvents[i].event);

				cleanGraphic();
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

	const oGraphicTypes = {
		removeBackChar                  : 1,
		removeBackWord                  : 1.5,
		removeChart                     : 2,
		removeShape                     : 3,
		removeGroup                     : 4,
		removeShapeInGroup              : 5,
		addTab                          : 6,
		selectNextObject                : 7,
		selectPreviousObject            : 8,
		visitHyperink                   : 9,
		addLineInMath                   : 10,
		addBreakLine                    : 11,
		addParagraph                    : 12,
		createTxBody                    : 13,
		moveToStartInEmptyContent       : 14,
		selectAllAfterEnter             : 15,
		moveCursorToStartPositionInTitle: 16,
		selectAllTitleAfterEnter        : 17,
		resetTextSelection              : 18,
		resetStepSelection              : 18.5,
		moveCursorToEndDocument         : 19,
		moveCursorToEndLine             : 19.5,
		selectToEndDocument             : 20,
		selectToEndLine                 : 20.5,
		moveCursorToStartDocument       : 21,
		moveCursorToStartLine           : 21.5,
		selectToStartDocument           : 22,
		selectToStartLine               : 22.5,
		moveCursorLeftChar              : 23,
		selectCursorLeftChar            : 24,
		moveCursorLeftWord              : 25,
		selectCursorLeftWord            : 26,
		bigMoveGraphicObjectLeft        : 27,
		littleMoveGraphicObjectLeft     : 28,
		moveCursorRightChar             : 29,
		selectCursorRightChar           : 30,
		moveCursorRightWord             : 31,
		selectCursorRightWord           : 32,
		bigMoveGraphicObjectRight       : 33,
		littleMoveGraphicObjectRight    : 34,
		moveCursorUp                    : 35,
		selectCursorUp                  : 36,
		bigMoveGraphicObjectUp          : 37,
		littleMoveGraphicObjectUp       : 38,
		moveCursorDown                  : 39,
		selectCursorDown                : 40,
		bigMoveGraphicObjectDown        : 41,
		littleMoveGraphicObjectDown     : 42,
		removeFrontWord                 : 43,
		removeFrontChar                 : 44,
		selectAllContent                : 45,
		selectAllDrawings               : 46,
		bold                            : 47,
		cleanSlicer                     : 48,
		centerAlign                     : 49,
		italic                          : 50,
		justifyAlign                    : 51,
		leftAlign                       : 52,
		rightAlign                      : 53,
		invertMultiselectSlicer         : 54,
		underline                       : 55,
		superscriptAndSubscript         : 56,
		superscript                     : 57,
		enDash                          : 58,
		subscript                       : 61,
		increaseFontSize                : 62,
		decreaseFontSize                : 63
	};

	const oGraphicTestEvents = {};
	oGraphicTestEvents[oGraphicTypes.removeBackChar] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.removeBackWord] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, true, false, false), testMacOs)
	];
	oGraphicTestEvents[oGraphicTypes.removeChart] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.removeShape] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.removeGroup] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.removeShapeInGroup] = [
		new CTestEvent(createEvent(oKeyCode.BackSpace, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.addTab] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.selectNextObject] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.selectPreviousObject] = [
		new CTestEvent(createEvent(oKeyCode.Tab, false, true, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.visitHyperink] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.addLineInMath] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.addBreakLine] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, true, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.addParagraph] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.createTxBody] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.moveToStartInEmptyContent] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.selectAllAfterEnter] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.moveCursorToStartPositionInTitle] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectAllTitleAfterEnter] = [
		new CTestEvent(createEvent(oKeyCode.Enter, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.resetTextSelection] = [
		new CTestEvent(createEvent(oKeyCode.Esc, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.resetStepSelection] = [
		new CTestEvent(createEvent(oKeyCode.Esc, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.moveCursorToEndDocument] = [
		new CTestEvent(createEvent(oKeyCode.End, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorToEndLine] = [
		new CTestEvent(createEvent(oKeyCode.End, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, false, false, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.selectToEndDocument] = [
		new CTestEvent(createEvent(oKeyCode.End, true, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectToEndLine] = [
		new CTestEvent(createEvent(oKeyCode.End, false, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, true, false, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorToStartDocument] = [
		new CTestEvent(createEvent(oKeyCode.Home, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorToStartLine] = [
		new CTestEvent(createEvent(oKeyCode.Home, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, false, false, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.selectToStartDocument] = [
		new CTestEvent(createEvent(oKeyCode.Home, true, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectToStartLine] = [
		new CTestEvent(createEvent(oKeyCode.Home, false, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, true, false, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorLeftChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.selectCursorLeftChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorLeftWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, true, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.selectCursorLeftWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, true, false, false, false), testWindows),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, true, true, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.bigMoveGraphicObjectLeft] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.littleMoveGraphicObjectLeft] = [
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorRightChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectCursorRightChar] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorRightWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, true, false, false), testMacOs)

	];
	oGraphicTestEvents[oGraphicTypes.selectCursorRightWord] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, true, false, false, false), testWindows),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, true, true, false, false), testMacOs),

	];
	oGraphicTestEvents[oGraphicTypes.bigMoveGraphicObjectRight] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.littleMoveGraphicObjectRight] = [
		new CTestEvent(createEvent(oKeyCode.ArrowRight, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectCursorUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.bigMoveGraphicObjectUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, false, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.littleMoveGraphicObjectUp] = [
		new CTestEvent(createEvent(oKeyCode.ArrowTop, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.moveCursorDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectCursorDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.bigMoveGraphicObjectDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.littleMoveGraphicObjectDown] = [
		new CTestEvent(createEvent(oKeyCode.ArrowBottom, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.removeFrontWord] = [
		new CTestEvent(createEvent(oKeyCode.Delete, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.removeFrontChar] = [
		new CTestEvent(createEvent(oKeyCode.Delete, false, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectAllContent] = [
		new CTestEvent(createEvent(oKeyCode.A, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.selectAllDrawings] = [
		new CTestEvent(createEvent(oKeyCode.A, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.bold] = [
		new CTestEvent(createEvent(oKeyCode.B, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.cleanSlicer] = [
		new CTestEvent(createEvent(oKeyCode.C, true, false, true, false, false), testMacOs),
		new CTestEvent(createEvent(oKeyCode.C, false, false, true, false, false), testWindows)
	];
	oGraphicTestEvents[oGraphicTypes.centerAlign] = [
		new CTestEvent(createEvent(oKeyCode.E, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.italic] = [
		new CTestEvent(createEvent(oKeyCode.I, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.justifyAlign] = [
		new CTestEvent(createEvent(oKeyCode.J, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.leftAlign] = [
		new CTestEvent(createEvent(oKeyCode.L, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.rightAlign] = [
		new CTestEvent(createEvent(oKeyCode.R, true, false, false, false, false))
	];
	oGraphicTestEvents[oGraphicTypes.invertMultiselectSlicer] = [
		new CTestEvent(createEvent(oKeyCode.S, true, false, true, false, false), testMacOs),
		new CTestEvent(createEvent(oKeyCode.S, false, false, true, false, false), testWindows)
	];
	oGraphicTestEvents[oGraphicTypes.underline] = [
		new CTestEvent(createEvent(oKeyCode.U, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.superscript] = [
		new CTestEvent(createEvent(oKeyCode.Comma, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.enDash] = [
		new CTestEvent(createEvent(oKeyCode.Minus, true, true, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.subscript] = [
		new CTestEvent(createEvent(oKeyCode.Period, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.increaseFontSize] = [
		new CTestEvent(createEvent(oKeyCode.BracketRight, true, false, false, false, false))

	];
	oGraphicTestEvents[oGraphicTypes.decreaseFontSize] = [
		new CTestEvent(createEvent(oKeyCode.BracketLeft, true, false, false, false, false))

	];

	function startGraphicTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oGraphicTestEvents);
	}

	const oTableTypes = {
		refreshAllConnections     : 0,
		refreshSelectedConnections: 1,
		changeFormatTableInfo     : 2,
		calculateAll              : 3,
		calculateWorkbook         : 4,
		calculateActiveSheet      : 5,
		calculateOnlyChanged      : 6,
		focusOnCellEditor         : 7,
		addDate                   : 8,
		addTime                   : 9,
		removeActiveCell          : 10,
		emptyRange                : 11,
		moveActiveCellToLeft      : 12,
		moveActiveCellToRight     : 13,
		moveActiveCellToDown      : 14,
		moveActiveCellToUp        : 15,
		reset                     : 16,
		disableNumLock            : 17,
		disableScrollLock         : 18,
		selectColumn              : 19,
		selectRow                 : 20,
		selectSheet               : 21,
		addSeparator              : 22,
		goToPreviousSheet         : 23,
		moveToTopCell             : 24,
		moveToNextSheet           : 25,
		moveToLeftEdgeCell        : 26,
		selectToLeftEdgeCell      : 27,
		moveToLeftCell            : 28,
		selectToLeftCell          : 29,
		moveToRightEdgeCell       : 30,
		selectToRightEdgeCell     : 31,
		moveToRightCell           : 32,
		selectToRightCell         : 33,
		selectToTopCell           : 34,
		moveToUpCell              : 35,
		selectToUpCell            : 36,
		moveToBottomCell          : 37,
		selectToBottomCell        : 38,
		moveToDownCell            : 39,
		selectToDownCell          : 40,
		moveToFirstColumn         : 41,
		selectToFirstColumn       : 42,
		moveToLeftEdgeTop         : 43,
		selectToLeftEdgeTop       : 44,
		moveToRightBottomEdge     : 45,
		selectToRightBottomEdge   : 46,
		setNumberFormat           : 47,
		setTimeFormat             : 48,
		setDateFormat             : 49,
		setCurrencyFormat         : 50,
		setPercentFormat          : 51,
		setStrikethrough          : 52,
		setExponentialFormat      : 53,
		setBold                   : 54,
		setItalic                 : 55,
		setUnderline              : 56,
		setGeneralFormat          : 57,
		redo                      : 58,
		undo                      : 59,
		print                     : 60,
		addSum                    : 61,
		moveToUpperCell           : 62,
		contextMenu               : 63,
		moveToLowerCell           : 64,
		selectToLowerCell         : 65
	};

	const oTableEvents = {};
	oTableEvents[oTableTypes.refreshAllConnections] = [
		new CTestEvent(createEvent(116, true, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.refreshSelectedConnections] = [
		new CTestEvent(createEvent(116, false, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.changeFormatTableInfo] = [
		new CTestEvent(createEvent(82, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.calculateAll] = [
		new CTestEvent(createEvent(120, true, true, true, false, false))
	];

	oTableEvents[oTableTypes.calculateActiveSheet] = [
		new CTestEvent(createEvent(120, false, true, false, false, false))

	];
	oTableEvents[oTableTypes.calculateOnlyChanged] = [
		new CTestEvent(createEvent(120, false, false, false, false, false))

	];
	oTableEvents[oTableTypes.focusOnCellEditor] = [
		new CTestEvent(createEvent(113, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.addDate] = [
		new CTestEvent(createEvent(186, true, false, false, false, false)),
		new CTestEvent(createEvent(59, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.addTime] = [
		new CTestEvent(createEvent(186, true, true, false, false, false)),
		new CTestEvent(createEvent(59, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.removeActiveCell] = [
		new CTestEvent(createEvent(8, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.emptyRange] = [
		new CTestEvent(createEvent(46, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveActiveCellToLeft] = [
		new CTestEvent(createEvent(9, false, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveActiveCellToRight] = [
		new CTestEvent(createEvent(9, false, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveActiveCellToDown] = [
		new CTestEvent(createEvent(13, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveActiveCellToUp] = [
		new CTestEvent(createEvent(13, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.reset] = [
		new CTestEvent(createEvent(27, false, false, false, false))
	];
	oTableEvents[oTableTypes.disableNumLock] = [
		new CTestEvent(createEvent(144, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.disableScrollLock] = [
		new CTestEvent(createEvent(145, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectColumn] = [
		new CTestEvent(createEvent(32, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectRow] = [
		new CTestEvent(createEvent(32, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.selectSheet] = [
		new CTestEvent(createEvent(32, true, true, false, false, false), testWindows),
		new CTestEvent(createEvent(65, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.addSeparator] = [
		new CTestEvent(createEvent(110, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.goToPreviousSheet] = [
		new CTestEvent(createEvent(33, false, false, true, false, false))
	];
	oTableEvents[oTableTypes.moveToUpperCell] = [
		new CTestEvent(createEvent(33, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveToTopCell] = [
		new CTestEvent(createEvent(38, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveToNextSheet] = [
		new CTestEvent(createEvent(33, false, false, true, false, false))
	];
	oTableEvents[oTableTypes.moveToBottomCell] = [
		new CTestEvent(createEvent(40, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveToLowerCell] = [
		new CTestEvent(createEvent(34, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToLowerCell] = [
		new CTestEvent(createEvent(34, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToLeftEdgeCell] = [
		new CTestEvent(createEvent(37, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToLeftEdgeCell] = [
		new CTestEvent(createEvent(37, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.moveToLeftCell] = [
		new CTestEvent(createEvent(37, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToLeftCell] = [
		new CTestEvent(createEvent(37, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToRightEdgeCell] = [
		new CTestEvent(createEvent(39, true, false, false, false, false)),
		new CTestEvent(createEvent(35, false, false, false, false, false))

	];
	oTableEvents[oTableTypes.selectToRightEdgeCell] = [
		new CTestEvent(createEvent(39, true, true, false, false, false)),
		new CTestEvent(createEvent(35, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToRightCell] = [
		new CTestEvent(createEvent(39, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToRightCell] = [
		new CTestEvent(createEvent(39, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.selectToTopCell] = [
		new CTestEvent(createEvent(38, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToUpCell] = [
		new CTestEvent(createEvent(38, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToUpCell] = [
		new CTestEvent(createEvent(38, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.selectToBottomCell] = [
		new CTestEvent(createEvent(40, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToDownCell] = [
		new CTestEvent(createEvent(40, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToDownCell] = [
		new CTestEvent(createEvent(40, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToFirstColumn] = [
		new CTestEvent(createEvent(36, false, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToFirstColumn] = [
		new CTestEvent(createEvent(36, false, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToLeftEdgeTop] = [
		new CTestEvent(createEvent(36, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToLeftEdgeTop] = [
		new CTestEvent(createEvent(36, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.moveToRightBottomEdge] = [
		new CTestEvent(createEvent(35, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.selectToRightBottomEdge] = [
		new CTestEvent(createEvent(35, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setNumberFormat] = [
		new CTestEvent(createEvent(49, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setTimeFormat] = [
		new CTestEvent(createEvent(50, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setDateFormat] = [
		new CTestEvent(createEvent(51, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setCurrencyFormat] = [
		new CTestEvent(createEvent(52, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setPercentFormat] = [
		new CTestEvent(createEvent(53, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setStrikethrough] = [
		new CTestEvent(createEvent(53, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.setExponentialFormat] = [
		new CTestEvent(createEvent(54, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.setBold] = [
		new CTestEvent(createEvent(66, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.setItalic] = [
		new CTestEvent(createEvent(73, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.setUnderline] = [
		new CTestEvent(createEvent(85, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.setGeneralFormat] = [
		new CTestEvent(createEvent(192, true, true, false, false, false))
	];
	oTableEvents[oTableTypes.redo] = [
		new CTestEvent(createEvent(89, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.undo] = [
		new CTestEvent(createEvent(90, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.print] = [
		new CTestEvent(createEvent(80, true, false, false, false, false))
	];
	oTableEvents[oTableTypes.addSum] = [
		new CTestEvent(createEvent(61, false, false, true, false, false), testWindows),
		new CTestEvent(createEvent(61, true, false, true, false, false), testMacOs)
	];
	oTableEvents[oTableTypes.contextMenu] = [
		new CTestEvent(createEvent(93, false, false, false, false, false))
	];

	function startTableTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oTableEvents);
	}

	const oCellEditorTypes = {
		closeWithoutSave     : 0,
		addNewLine           : 1,
		saveAndMoveDown      : 2,
		saveAndMoveUp        : 3,
		saveAndMoveRight     : 4,
		saveAndMoveLeft      : 5,
		removeCharBack       : 6,
		removeWordBack       : 7,
		addSpace             : 8,
		moveToEndLine        : 9,
		moveToEndDocument    : 10,
		selectToEndLine      : 11,
		selectToEndDocument  : 12,
		moveToStartLine      : 13,
		moveToStartDocument  : 14,
		selectToStartLine    : 15,
		selectToStartDocument: 16,
		moveCursorLeftChar   : 17,
		moveCursorLeftWord   : 18,
		selectLeftChar       : 19,
		selectLeftWord       : 20,
		moveToUpLine         : 21,
		selectToUpLine       : 22,
		moveToRightChar      : 23,
		moveToRightWord      : 24,
		selectRightChar      : 25,
		selectRightWord      : 26,
		moveToDownLine       : 27,
		selectToDownLine     : 28,
		deleteFrontChar      : 29,
		deleteFrontWord      : 30,
		setStrikethrough     : 31,
		selectAll            : 32,
		setBold              : 33,
		setItalic            : 34,
		setUnderline         : 35,
		disableScrollLock    : 36,
		disableNumLock       : 37,
		disablePrint         : 38,
		undo                 : 39,
		redo                 : 40,
		addSeparator         : 41,
		disableF2            : 42,
		switchReference      : 43,
		addTime              : 44,
		addDate              : 45
	};

	const oCellEditorEvents = {};
	oCellEditorEvents[oCellEditorTypes.closeWithoutSave] = [
		new CTestEvent(createEvent(27, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.addNewLine] = [
		new CTestEvent(createEvent(13, false, false, true, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.saveAndMoveDown] = [
		new CTestEvent(createEvent(13, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.saveAndMoveUp] = [
		new CTestEvent(createEvent(13, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.saveAndMoveRight] = [
		new CTestEvent(createEvent(9, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.saveAndMoveLeft] = [
		new CTestEvent(createEvent(9, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.removeCharBack] = [
		new CTestEvent(createEvent(8, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.removeWordBack] = [
		new CTestEvent(createEvent(8, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(8, false, false, true, false, false), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.addSpace] = [
		new CTestEvent(createEvent(32, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveToEndLine] = [
		new CTestEvent(createEvent(35, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, false, false, false, true), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.moveToEndDocument] = [
		new CTestEvent(createEvent(35, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectToEndLine] = [
		new CTestEvent(createEvent(35, false, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowRight, false, true, false, false, true), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.selectToEndDocument] = [
		new CTestEvent(createEvent(35, true, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveToStartLine] = [
		new CTestEvent(createEvent(36, false, false, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, false, false, false, true), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.moveToStartDocument] = [
		new CTestEvent(createEvent(36, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectToStartLine] = [
		new CTestEvent(createEvent(36, false, true, false, false, false)),
		new CTestEvent(createEvent(oKeyCode.ArrowLeft, false, true, false, false, true), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.selectToStartDocument] = [
		new CTestEvent(createEvent(36, true, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveCursorLeftChar] = [
		new CTestEvent(createEvent(37, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveCursorLeftWord] = [
		new CTestEvent(createEvent(37, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(37, false, false, true, false, false), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.selectLeftChar] = [
		new CTestEvent(createEvent(37, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectLeftWord] = [
		new CTestEvent(createEvent(37, true, true, false, false, false), testWindows),
		new CTestEvent(createEvent(37, false, true, true, false, false), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.moveToUpLine] = [
		new CTestEvent(createEvent(38, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectToUpLine] = [
		new CTestEvent(createEvent(38, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveToRightChar] = [
		new CTestEvent(createEvent(39, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.moveToRightWord] = [
		new CTestEvent(createEvent(39, true, false, false, false, false), testWindows),
		new CTestEvent(createEvent(39, false, false, true, false, false), testMacOs),
	];
	oCellEditorEvents[oCellEditorTypes.selectRightChar] = [
		new CTestEvent(createEvent(39, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectRightWord] = [
		new CTestEvent(createEvent(39, true, true, false, false, false), testWindows),
		new CTestEvent(createEvent(39, false, true, true, false, false), testMacOs)
	];
	oCellEditorEvents[oCellEditorTypes.moveToDownLine] = [
		new CTestEvent(createEvent(40, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectToDownLine] = [
		new CTestEvent(createEvent(40, false, true, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.deleteFrontChar] = [
		new CTestEvent(createEvent(46, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.deleteFrontWord] = [
		new CTestEvent(createEvent(46, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.setStrikethrough] = [
		new CTestEvent(createEvent(53, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.selectAll] = [
		new CTestEvent(createEvent(65, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.setBold] = [
		new CTestEvent(createEvent(66, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.setItalic] = [
		new CTestEvent(createEvent(73, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.setUnderline] = [
		new CTestEvent(createEvent(85, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.disableScrollLock] = [
		new CTestEvent(createEvent(145, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.disableNumLock] = [
		new CTestEvent(createEvent(144, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.disablePrint] = [
		new CTestEvent(createEvent(80, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.undo] = [
		new CTestEvent(createEvent(90, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.redo] = [
		new CTestEvent(createEvent(89, true, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.addSeparator] = [
		new CTestEvent(createEvent(110, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.disableF2] = [
		new CTestEvent(createEvent(113, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.switchReference] = [
		new CTestEvent(createEvent(115, false, false, false, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.addTime] = [
		new CTestEvent(createEvent(186, true, true, false, false))
	];
	oCellEditorEvents[oCellEditorTypes.addDate] = [
		new CTestEvent(createEvent(186, true, false, false, false))
	];

	function startCellEditorTest(fCallback, nShortcutType)
	{
		privateStartTest(fCallback, nShortcutType, oCellEditorEvents);
	}

	AscTestShortcut.startCellEditorTest = startCellEditorTest;
	AscTestShortcut.startTableTest = startTableTest;
	AscTestShortcut.startGraphicTest = startGraphicTest;
	AscTestShortcut.oTableTypes = oTableTypes;
	AscTestShortcut.oCellEditorTypes = oCellEditorTypes;
	AscTestShortcut.oGraphicTypes = oGraphicTypes;
})(window);
