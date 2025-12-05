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

(function (window)
{
	const testFlags = {
		nothing: 0x0000,
		macOs  : 0x1000,
		opera  : 0x0110,
	};

	function CNativeEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey, flags)
	{
		this.keyCode = nKeyCode;
		this.which = nKeyCode;
		this.ctrlKey = !!bIsCtrl;
		this.shiftKey = !!bIsShift;
		this.altKey = !!bIsAlt;
		this.metaKey = !!bIsMetaKey;
		this.isDefaultPrevented = false;
		this.isPropagationStopped = false;

		this.flags = flags || testFlags.nothing;
	}

	CNativeEvent.prototype.preventDefault = function ()
	{
		this.isDefaultPrevented = true;
	};
	CNativeEvent.prototype.stopPropagation = function ()
	{
		this.isPropagationStopped = true;
	};

	CNativeEvent.prototype.enableFlags = function ()
	{
		if (this.flags & testFlags.macOs)
		{
			AscCommon.AscBrowser.isMacOs = true;
		}

		if (this.flags & testFlags.opera)
		{
			AscCommon.AscBrowser.isOpera = true;
		}
	}
	CNativeEvent.prototype.disableFlags = function ()
	{
		if (this.flags & testFlags.macOs)
		{
			AscCommon.AscBrowser.isMacOs = false;
		}

		if (this.flags & testFlags.opera)
		{
			AscCommon.AscBrowser.isOpera = false;
		}
	}
	const keyCodes =
		{
			BackSpace       : 8,
			Tab             : 9,
			Enter           : 13,
			Esc             : 27,
			Space           : 32,
			PageUp          : 33,
			PageDown        : 34,
			End             : 35,
			Home            : 36,
			ArrowLeft       : 37,
			ArrowTop        : 38,
			ArrowRight      : 39,
			ArrowBottom     : 40,
			Digit1          : 49,
			Digit2          : 50,
			Digit3          : 51,
			Digit4          : 52,
			Digit5          : 53,
			Digit6          : 54,
			SemicolonFirefox: 59,
			Delete          : 46,
			EqualFirefox    : 61,
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
			ContextMenu     : 93,
			NumpadDecimal   : 110,
			F2              : 113,
			F4              : 115,
			F5              : 116,
			F9              : 120,
			F10             : 121,
			NumLock         : 144,
			ScrollLock      : 145,
			MinusFirefox    : 173,
			Semicolon       : 186,
			Equal           : 187,
			Comma           : 188,
			Minus           : 189,
			Period          : 190,
			Backquote       : 192,
			BracketLeft     : 219,
			BracketRight    : 221,
			OperaContextMenu: 57351,
		}

	const graphicHotkeyTypes = {
		removeBackChar              : 1,
		removeChart                 : 2,
		removeShape                 : 3,
		removeGroup                 : 4,
		removeShapeInGroup          : 5,
		addTab                      : 6,
		selectNextObject            : 7,
		selectPreviousObject        : 8,
		visitHyperink               : 9,
		addLineInMath               : 10,
		addBreakLine                : 11,
		addParagraph                : 12,
		createTxBody                : 13,
		moveToStartInEmptyContent   : 14,
		selectAllAfterEnter         : 15,
		selectAllTitleAfterEnter    : 17,
		resetTextSelection          : 18,
		moveCursorToEndDocument     : 19,
		selectToEndDocument         : 20,
		moveCursorToStartDocument   : 21,
		selectToStartDocument       : 22,
		moveCursorLeftChar          : 23,
		selectCursorLeftChar        : 24,
		moveCursorLeftWord          : 25,
		selectCursorLeftWord        : 26,
		bigMoveGraphicObjectLeft    : 27,
		littleMoveGraphicObjectLeft : 28,
		moveCursorRightChar         : 29,
		selectCursorRightChar       : 30,
		moveCursorRightWord         : 31,
		selectCursorRightWord       : 32,
		bigMoveGraphicObjectRight   : 33,
		littleMoveGraphicObjectRight: 34,
		moveCursorUp                : 35,
		selectCursorUp              : 36,
		bigMoveGraphicObjectUp      : 37,
		littleMoveGraphicObjectUp   : 38,
		moveCursorDown              : 39,
		selectCursorDown            : 40,
		bigMoveGraphicObjectDown    : 41,
		littleMoveGraphicObjectDown : 42,
		removeFrontWord             : 43,
		removeFrontChar             : 44,
		cleanSlicer                 : 48,
		invertMultiselectSlicer     : 54,
		removeBackWord              : 64,
		selectToStartLine           : 65,
		moveCursorToStartLine       : 66,
		selectToEndLine             : 67,
		moveCursorToEndLine         : 68,
		resetStepSelection          : 69
	};

	const graphicEvents = {};
	graphicEvents[graphicHotkeyTypes.removeBackChar] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.removeBackWord] = [
		new CNativeEvent(keyCodes.BackSpace, true, false, false, false),
		new CNativeEvent(keyCodes.BackSpace, false, false, true, false, testFlags.macOs)
	];
	graphicEvents[graphicHotkeyTypes.removeChart] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.removeShape] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.removeGroup] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.removeShapeInGroup] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.addTab] = [
		new CNativeEvent(keyCodes.Tab, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.selectNextObject] = [
		new CNativeEvent(keyCodes.Tab, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.selectPreviousObject] = [
		new CNativeEvent(keyCodes.Tab, false, true, false, false)
	];
	graphicEvents[graphicHotkeyTypes.visitHyperink] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.addLineInMath] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.addBreakLine] = [
		new CNativeEvent(keyCodes.Enter, false, true, false, false)
	];
	graphicEvents[graphicHotkeyTypes.addParagraph] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.createTxBody] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.moveToStartInEmptyContent] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.selectAllAfterEnter] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.selectAllTitleAfterEnter] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.resetTextSelection] = [
		new CNativeEvent(keyCodes.Esc, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.resetStepSelection] = [
		new CNativeEvent(keyCodes.Esc, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.moveCursorToEndDocument] = [
		new CNativeEvent(keyCodes.End, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorToEndLine] = [
		new CNativeEvent(keyCodes.End, false, false, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, true, testFlags.macOs)

	];
	graphicEvents[graphicHotkeyTypes.selectToEndDocument] = [
		new CNativeEvent(keyCodes.End, true, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.selectToEndLine] = [
		new CNativeEvent(keyCodes.End, false, true, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, true, false, true, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.moveCursorToStartDocument] = [
		new CNativeEvent(keyCodes.Home, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorToStartLine] = [
		new CNativeEvent(keyCodes.Home, false, false, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, true, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.selectToStartDocument] = [
		new CNativeEvent(keyCodes.Home, true, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.selectToStartLine] = [
		new CNativeEvent(keyCodes.Home, false, true, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, true, false, true, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.moveCursorLeftChar] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.selectCursorLeftChar] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorLeftWord] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, false, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, false, true, false, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.selectCursorLeftWord] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, true, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, true, true, false, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.bigMoveGraphicObjectLeft] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.littleMoveGraphicObjectLeft] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorRightChar] = [
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.selectCursorRightChar] = [
		new CNativeEvent(keyCodes.ArrowRight, false, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorRightWord] = [
		new CNativeEvent(keyCodes.ArrowRight, true, false, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, false, true, false, testFlags.macOs)
	];
	graphicEvents[graphicHotkeyTypes.selectCursorRightWord] = [
		new CNativeEvent(keyCodes.ArrowRight, true, true, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, true, true, false, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.bigMoveGraphicObjectRight] = [
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.littleMoveGraphicObjectRight] = [
		new CNativeEvent(keyCodes.ArrowRight, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorUp] = [
		new CNativeEvent(keyCodes.ArrowTop, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.selectCursorUp] = [
		new CNativeEvent(keyCodes.ArrowTop, false, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.bigMoveGraphicObjectUp] = [
		new CNativeEvent(keyCodes.ArrowTop, false, false, false, false)
	];
	graphicEvents[graphicHotkeyTypes.littleMoveGraphicObjectUp] = [
		new CNativeEvent(keyCodes.ArrowTop, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.moveCursorDown] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.selectCursorDown] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, true, false, false)

	];
	graphicEvents[graphicHotkeyTypes.bigMoveGraphicObjectDown] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.littleMoveGraphicObjectDown] = [
		new CNativeEvent(keyCodes.ArrowBottom, true, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.removeFrontWord] = [
		new CNativeEvent(keyCodes.Delete, true, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, true, false, testFlags.macOs),

	];
	graphicEvents[graphicHotkeyTypes.removeFrontChar] = [
		new CNativeEvent(keyCodes.Delete, false, false, false, false)

	];
	graphicEvents[graphicHotkeyTypes.cleanSlicer] = [
		new CNativeEvent(keyCodes.C, true, false, true, false, testFlags.macOs),
		new CNativeEvent(keyCodes.C, false, false, true, false)
	];
	graphicEvents[graphicHotkeyTypes.invertMultiselectSlicer] = [
		new CNativeEvent(keyCodes.S, true, false, true, false, testFlags.macOs),
		new CNativeEvent(keyCodes.S, false, false, true, false)
	];

	const tableHotkeyTypes = {
		changeFormatTableInfo     : 2,
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
		moveToTopCell             : 24,
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
		moveToUpperCell           : 62,
		contextMenu               : 63,
		moveToLowerCell           : 64,
		selectToLowerCell         : 65,
		selectToUpperCell         : 66,
		showFilterOptions         : 67,
		showAutoComplete          : 68,
		showDataValidation        : 69,
		openCellEditor            : 70,
		increaseFontSize          : 71,
		decreaseFontSize          : 72,
		};

	const tableEvents = {};
	tableEvents[tableHotkeyTypes.openCellEditor] = [
		new CNativeEvent(keyCodes.Enter, true, false, false, false),
		new CNativeEvent(keyCodes.Enter, false, false, true, false)
	];
	tableEvents[tableHotkeyTypes.changeFormatTableInfo] = [
		new CNativeEvent(keyCodes.R, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.removeActiveCell] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.emptyRange] = [
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveActiveCellToLeft] = [
		new CNativeEvent(keyCodes.Tab, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveActiveCellToRight] = [
		new CNativeEvent(keyCodes.Tab, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveActiveCellToDown] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveActiveCellToUp] = [
		new CNativeEvent(keyCodes.Enter, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.reset] = [
		new CNativeEvent(keyCodes.Esc, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.disableNumLock] = [
		new CNativeEvent(keyCodes.NumLock, false, false, false, false, testFlags.opera)
	];
	tableEvents[tableHotkeyTypes.disableScrollLock] = [
		new CNativeEvent(keyCodes.ScrollLock, false, false, false, false, testFlags.opera)
	];
	tableEvents[tableHotkeyTypes.selectColumn] = [
		new CNativeEvent(keyCodes.Space, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.selectRow] = [
		new CNativeEvent(keyCodes.Space, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectSheet] = [
		new CNativeEvent(keyCodes.Space, true, true, false, false),
		new CNativeEvent(keyCodes.A, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToUpperCell] = [
		new CNativeEvent(keyCodes.PageUp, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToUpperCell] = [
		new CNativeEvent(keyCodes.PageUp, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToTopCell] = [
		new CNativeEvent(keyCodes.ArrowTop, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToBottomCell] = [
		new CNativeEvent(keyCodes.ArrowBottom, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToLowerCell] = [
		new CNativeEvent(keyCodes.PageDown, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToLowerCell] = [
		new CNativeEvent(keyCodes.PageDown, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToLeftEdgeCell] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToLeftEdgeCell] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToLeftCell] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToLeftCell] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToRightEdgeCell] = [
		new CNativeEvent(keyCodes.ArrowRight, true, false, false, false),
		new CNativeEvent(keyCodes.End, false, false, false, false)

	];
	tableEvents[tableHotkeyTypes.selectToRightEdgeCell] = [
		new CNativeEvent(keyCodes.ArrowRight, true, true, false, false),
		new CNativeEvent(keyCodes.End, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToRightCell] = [
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToRightCell] = [
		new CNativeEvent(keyCodes.ArrowRight, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToTopCell] = [
		new CNativeEvent(keyCodes.ArrowTop, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToUpCell] = [
		new CNativeEvent(keyCodes.ArrowTop, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToUpCell] = [
		new CNativeEvent(keyCodes.ArrowTop, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToBottomCell] = [
		new CNativeEvent(keyCodes.ArrowBottom, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToDownCell] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToDownCell] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToFirstColumn] = [
		new CNativeEvent(keyCodes.Home, false, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToFirstColumn] = [
		new CNativeEvent(keyCodes.Home, false, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToLeftEdgeTop] = [
		new CNativeEvent(keyCodes.Home, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToLeftEdgeTop] = [
		new CNativeEvent(keyCodes.Home, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.moveToRightBottomEdge] = [
		new CNativeEvent(keyCodes.End, true, false, false, false)
	];
	tableEvents[tableHotkeyTypes.selectToRightBottomEdge] = [
		new CNativeEvent(keyCodes.End, true, true, false, false)
	];
	tableEvents[tableHotkeyTypes.contextMenu] = [new CNativeEvent(keyCodes.ContextMenu, false, false, false, false)];
	tableEvents[tableHotkeyTypes.showFilterOptions] = [new CNativeEvent(keyCodes.ArrowBottom, false, false, true, false)];
	tableEvents[tableHotkeyTypes.showAutoComplete] = [new CNativeEvent(keyCodes.ArrowBottom, false, false, true, false)];
	tableEvents[tableHotkeyTypes.showDataValidation] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, false, true, false),
		new CNativeEvent(keyCodes.ArrowTop, false, false, true, false)
	];
	tableEvents[tableHotkeyTypes.increaseFontSize] = [new CNativeEvent(keyCodes.BracketRight, true, false, false, false)];
	tableEvents[tableHotkeyTypes.decreaseFontSize] = [new CNativeEvent(keyCodes.BracketLeft, true, false, false, false)];

	const cellEditorHotkeyTypes = {
		closeWithoutSave     : 0,
		addNewLine           : 1,
		saveAndMoveDown      : 2,
		saveAndMoveUp        : 3,
		saveAndMoveRight     : 4,
		saveAndMoveLeft      : 5,
		removeCharBack       : 6,
		removeWordBack       : 7,
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
		disableScrollLock    : 36,
		disableNumLock       : 37,
		disableF2            : 42,
		increaseFontSize     : 46,
		decreaseFontSize     : 47
	};

	const oCellEditorEvents = {};
	oCellEditorEvents[cellEditorHotkeyTypes.closeWithoutSave] = [
		new CNativeEvent(keyCodes.Esc, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.addNewLine] = [
		new CNativeEvent(keyCodes.Enter, false, false, true, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.saveAndMoveDown] = [
		new CNativeEvent(keyCodes.Enter, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.saveAndMoveUp] = [
		new CNativeEvent(keyCodes.Enter, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.saveAndMoveRight] = [
		new CNativeEvent(keyCodes.Tab, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.saveAndMoveLeft] = [
		new CNativeEvent(keyCodes.Tab, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.removeCharBack] = [
		new CNativeEvent(keyCodes.BackSpace, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.removeWordBack] = [
		new CNativeEvent(keyCodes.BackSpace, true, false, false, false),
		new CNativeEvent(keyCodes.BackSpace, false, false, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToEndLine] = [
		new CNativeEvent(keyCodes.End, false, false, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, true, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToEndDocument] = [
		new CNativeEvent(keyCodes.End, true, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToEndLine] = [
		new CNativeEvent(keyCodes.End, false, true, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, true, false, true, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToEndDocument] = [
		new CNativeEvent(keyCodes.End, true, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToStartLine] = [
		new CNativeEvent(keyCodes.Home, false, false, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, true, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToStartDocument] = [
		new CNativeEvent(keyCodes.Home, true, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToStartLine] = [
		new CNativeEvent(keyCodes.Home, false, true, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, true, false, true, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToStartDocument] = [
		new CNativeEvent(keyCodes.Home, true, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveCursorLeftChar] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveCursorLeftWord] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, false, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, false, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectLeftChar] = [
		new CNativeEvent(keyCodes.ArrowLeft, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectLeftWord] = [
		new CNativeEvent(keyCodes.ArrowLeft, true, true, false, false),
		new CNativeEvent(keyCodes.ArrowLeft, false, true, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToUpLine] = [
		new CNativeEvent(keyCodes.ArrowTop, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToUpLine] = [
		new CNativeEvent(keyCodes.ArrowTop, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToRightChar] = [
		new CNativeEvent(keyCodes.ArrowRight, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToRightWord] = [
		new CNativeEvent(keyCodes.ArrowRight, true, false, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, false, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectRightChar] = [
		new CNativeEvent(keyCodes.ArrowRight, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectRightWord] = [
		new CNativeEvent(keyCodes.ArrowRight, true, true, false, false),
		new CNativeEvent(keyCodes.ArrowRight, false, true, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.moveToDownLine] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.selectToDownLine] = [
		new CNativeEvent(keyCodes.ArrowBottom, false, true, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.deleteFrontChar] = [
		new CNativeEvent(keyCodes.Delete, false, false, false, false)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.deleteFrontWord] = [
		new CNativeEvent(keyCodes.Delete, true, false, false, false),
		new CNativeEvent(keyCodes.Delete, false, false, true, false, testFlags.macOs),
	];
	oCellEditorEvents[cellEditorHotkeyTypes.disableScrollLock] = [
		new CNativeEvent(keyCodes.ScrollLock, false, false, false, false, testFlags.opera)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.disableNumLock] = [
		new CNativeEvent(keyCodes.NumLock, false, false, false, false, testFlags.opera)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.disableF2] = [
		new CNativeEvent(keyCodes.F2, false, false, false, false, testFlags.opera)
	];
	oCellEditorEvents[cellEditorHotkeyTypes.increaseFontSize] = [new CNativeEvent(keyCodes.BracketRight, true, false, false, false)];
	oCellEditorEvents[cellEditorHotkeyTypes.decreaseFontSize] = [new CNativeEvent(keyCodes.BracketLeft, true, false, false, false)];
	function ExecuteGraphicHotkey(type, eventIndex)
	{
		const event = graphicEvents[type][eventIndex || 0];
		return ExecuteShortcut(event);
	}

	function ExecuteTableHotkey(type, eventIndex)
	{
		const event = tableEvents[type][eventIndex || 0];
		return ExecuteShortcut(event);
	}

	function ExecuteCellEditorHotkey(type, eventIndex)
	{
		const event = oCellEditorEvents[type][eventIndex || 0];
		return ExecuteShortcut(event);
	}

	const MockShortcutEvent = {
		nShortcutType: 0,
		IsShift: function () {},
		IsCtrl: function () {},
		GetKeyCode: function () {}
	};
	function ExecuteShortcut(e)
	{
		if (AscFormat.isRealNumber(e))
		{
			MockShortcutEvent.nShortcutType = e;
			return Asc.editor.wb.onKeyDown(MockShortcutEvent);
		}
		e.enableFlags();
		AscCommon.check_KeyboardEvent(e);
		const res = Asc.editor.wb.onKeyDown(AscCommon.global_keyboardEvent);
		e.disableFlags();
		return res;
	}

	window.AscTestShortcut = {};
	AscTestShortcut.tableHotkeyTypes = tableHotkeyTypes;
	AscTestShortcut.cellEditorHotkeyTypes = cellEditorHotkeyTypes;
	AscTestShortcut.graphicHotkeyTypes = graphicHotkeyTypes;
	AscTestShortcut.graphicEvents = graphicEvents;
	AscTestShortcut.tableEvents = tableEvents;
	AscTestShortcut.ExecuteCellEditorHotkey = ExecuteCellEditorHotkey;
	AscTestShortcut.ExecuteTableHotkey = ExecuteTableHotkey;
	AscTestShortcut.ExecuteGraphicHotkey = ExecuteGraphicHotkey;
	AscTestShortcut.ExecuteShortcut = ExecuteShortcut;
})(window);
