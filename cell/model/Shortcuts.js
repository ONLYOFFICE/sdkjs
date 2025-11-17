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

(function(undefined) {
	const AscShortcut = Asc.CAscShortcut;
	const keyCodes = Asc.c_oAscKeyCodes;

	const c_oAscUnlockedShortcutActionTypes = {};
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.EditSelectAll] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.RecalculateAll] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellInsertDate] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellInsertTime] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.NextWorksheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Strikeout] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Italic] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Bold] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Underline] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.EditUndo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.EditRedo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.SpeechWorker] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellDateFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.ShowFormulas] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Subscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.Superscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CenterPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.JustifyPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.LeftPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.RightPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscSpreadsheetShortcutType.EnDash] = true;

	const c_oAscDefaultShortcuts = {};

	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog, keyCodes.KeyF, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFindAndReplaceMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFindAndReplaceMenu, keyCodes.KeyH, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel, keyCodes.KeyH, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Save] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Save, keyCodes.KeyS, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint, keyCodes.KeyP, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DownloadAs] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DownloadAs, keyCodes.KeyS, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenHelpMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenHelpMenu, keyCodes.F1, false, false, false, false)];

	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenContextMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenContextMenu, keyCodes.F10, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CloseMenu] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CloseMenu, keyCodes.Escape, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Zoom100] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Zoom100, keyCodes.Digit0, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveDown, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellDown, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellUp, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellRight, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellLeft, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstColumn, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomEdge] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomEdge, keyCodes.PageDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveTopEdge] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveTopEdge, keyCodes.PageUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell, keyCodes.Home, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet, keyCodes.End, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet, keyCodes.PageUp, false, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet, keyCodes.PageUp, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.NextWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet, keyCodes.PageDown, false, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet, keyCodes.PageDown, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ZoomIn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomIn, keyCodes.KeyEqual, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomIn, keyCodes.NumpadPlus, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ZoomOut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomOut, keyCodes.KeyMinus, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomOut, keyCodes.NumpadMinus, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.NavigatePreviousControl] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NavigatePreviousControl, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.NavigateNextControl] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NavigateNextControl, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectColumn, keyCodes.Space, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectRow, keyCodes.Space, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectOneCellRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectOneCellRight, keyCodes.ArrowRight, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectOneCellLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectOneCellLeft, keyCodes.ArrowLeft, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectOneCellUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectOneCellUp, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectOneCellDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectOneCellDown, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectCursorBeginningRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectCursorBeginningRow, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectCursorEndRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectCursorEndRow, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight, keyCodes.ArrowRight, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft, keyCodes.ArrowLeft, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp, keyCodes.ArrowUp, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown, keyCodes.ArrowDown, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet, keyCodes.Home, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell, keyCodes.End, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectFirstColumn] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectFirstColumn, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectDownOneScreen] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectDownOneScreen, keyCodes.PageDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectUpOneScreen] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectUpOneScreen, keyCodes.PageUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditUndo] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo, keyCodes.KeyZ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditRedo] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo, keyCodes.KeyY, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormula] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormula, keyCodes.KeyF, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNumberFormat, keyCodes.KeyO, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteFormulaAllFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteFormulaAllFormatting, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNoBorders] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNoBorders, keyCodes.KeyB, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteFormulaColumnWidth] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteFormulaColumnWidth, keyCodes.KeyW, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Transpose] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Transpose, keyCodes.KeyT, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteOnlyValue] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteOnlyValue, keyCodes.KeyV, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteValueNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteValueNumberFormat, keyCodes.KeyA, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteValueAllFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteValueAllFormatting, keyCodes.KeyE, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormatting] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormatting, keyCodes.KeyR, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PasteLink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PasteLink, keyCodes.KeyN, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.VisitHyperlink] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.VisitHyperlink, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Bold] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold, keyCodes.KeyB, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Underline] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline, keyCodes.KeyU, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Strikeout] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout, keyCodes.Digit5, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor, keyCodes.F2, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter, keyCodes.KeyL, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFilterWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFilterWindow, keyCodes.ArrowDown, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate, keyCodes.KeyL, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveDown, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveUp, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveRight, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveLeft, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryStay] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryStay, keyCodes.Enter, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryStay, keyCodes.Enter, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.FillSelectedCellRange] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.FillSelectedCellRange, keyCodes.Enter, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellStartNewLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellStartNewLine, keyCodes.Enter, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EquationAddPlaceholder] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EquationAddPlaceholder, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellEntryCancel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEntryCancel, keyCodes.Escape, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteLeftChar] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteLeftChar, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteRightChar] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteRightChar, keyCodes.Delete, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ClearActiveCellContent] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ClearActiveCellContent, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ClearSelectedCellsContent] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ClearSelectedCellsContent, keyCodes.Delete, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow, keyCodes.KeyEqual, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow, keyCodes.NumpadPlus, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow, keyCodes.KeyMinus, true, true, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow, keyCodes.NumpadMinus, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertDate] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate, keyCodes.KeySemicolon, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertTime] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime, keyCodes.KeySemicolon, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.AutoFill] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.AutoFill, keyCodes.ArrowDown, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditSelectAll] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll, keyCodes.KeyA, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveCharacterLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveCharacterLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveCharacterRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveCharacterRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineDown, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectCharacterRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectCharacterRight, keyCodes.ArrowRight, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectCharacterLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectCharacterLeft, keyCodes.ArrowLeft, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText, keyCodes.Home, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveEndText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveEndText, keyCodes.End, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText, keyCodes.Home, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectEndText] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectEndText, keyCodes.End, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveEndLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveEndLine, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectBeginningLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectBeginningLine, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectEndLine] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectEndLine, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectLineUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectLineUp, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectLineDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectLineDown, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots, keyCodes.F5, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots, keyCodes.F5, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow, keyCodes.KeyR, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenInsertFunctionDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenInsertFunctionDialog, keyCodes.F3, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RecalculateAll] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateAll, keyCodes.F9, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet, keyCodes.F9, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ShowFormulas] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ShowFormulas, keyCodes.KeyBackquote, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference, keyCodes.F4, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog, keyCodes.Digit1, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat, keyCodes.KeyBackquote, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat, keyCodes.Digit4, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat, keyCodes.Digit5, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat, keyCodes.Digit6, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellDateFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellDateFormat, keyCodes.Digit3, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat, keyCodes.Digit2, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat, keyCodes.Digit1, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditShape] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditShape, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditChart] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditChart, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepBottom, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveFocusNextObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveFocusNextObject, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveFocusPreviousObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveFocusPreviousObject, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DrawingAddTab] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DrawingAddTab, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Subscript] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript, keyCodes.Period, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Superscript] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript, keyCodes.Comma, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize, keyCodes.BracketRight, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize, keyCodes.BracketLeft, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CenterPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CenterPara, keyCodes.KeyE, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.JustifyPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.JustifyPara, keyCodes.KeyJ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RightPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RightPara, keyCodes.KeyR, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.LeftPara] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.LeftPara, keyCodes.KeyL, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EndParagraph] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EndParagraph, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.AddLineBreak] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.AddLineBreak, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RemoveGraphicalObject] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RemoveGraphicalObject, keyCodes.Delete, false, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RemoveGraphicalObject, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ExitAddingShapesMode] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ExitAddingShapesMode, keyCodes.Escape, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SpeechWorker] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SpeechWorker, keyCodes.KeyZ, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EnDash] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EnDash, keyCodes.KeyMinus, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft, keyCodes.ArrowLeft, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown, keyCodes.ArrowDown, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp, keyCodes.ArrowUp, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Italic] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic, keyCodes.KeyI, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator, keyCodes.NumpadDecimal, false, false, false, false)];

	if (AscCommon.AscBrowser.isMacOs) {
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel, keyCodes.KeyF, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenCommentField] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenCommentField, keyCodes.KeyA, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel, keyCodes.KeyQ, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank, keyCodes.ArrowLeft, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Cut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Cut, keyCodes.KeyX, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Copy] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Copy, keyCodes.KeyC, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Paste] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Paste, keyCodes.KeyV, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord, keyCodes.Delete, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord, keyCodes.Delete, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft, keyCodes.ArrowLeft, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveWordRight, keyCodes.ArrowRight, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft, keyCodes.ArrowLeft, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectWordRight, keyCodes.ArrowRight, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues, keyCodes.KeyC, true, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues, keyCodes.KeyC, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect, keyCodes.KeyS, true, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect, keyCodes.KeyS, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction, keyCodes.KeyEqual, true, false, true, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction, keyCodes.KeyEqual, false, false, true, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight, keyCodes.ArrowRight, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft, keyCodes.ArrowLeft, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog, keyCodes.KeyF, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel, keyCodes.KeyH, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Save].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Save, keyCodes.KeyS, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint, keyCodes.KeyP, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DownloadAs].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DownloadAs, keyCodes.KeyS, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Zoom100].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Zoom100, keyCodes.Digit0, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank, keyCodes.ArrowRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell, keyCodes.Home, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet, keyCodes.End, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ZoomIn].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomIn, keyCodes.KeyEqual, false, false, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomIn, keyCodes.NumpadPlus, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ZoomOut].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomOut, keyCodes.KeyMinus, false, false, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ZoomOut, keyCodes.NumpadMinus, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectColumn].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectColumn, keyCodes.Space, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight, keyCodes.ArrowRight, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft, keyCodes.ArrowLeft, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp, keyCodes.ArrowUp, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown, keyCodes.ArrowDown, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet, keyCodes.Home, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell, keyCodes.End, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditUndo].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditUndo, keyCodes.KeyZ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditRedo].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditRedo, keyCodes.KeyY, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink, keyCodes.KeyK, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Bold].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Bold, keyCodes.KeyB, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Underline].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Underline, keyCodes.KeyU, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Strikeout].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Strikeout, keyCodes.Digit5, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter, keyCodes.KeyL, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate, keyCodes.KeyL, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow, keyCodes.KeyEqual, false, true, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow, keyCodes.NumpadPlus, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow, keyCodes.KeyMinus, false, true, false, true), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow, keyCodes.NumpadMinus, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertDate].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertDate, keyCodes.KeySemicolon, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertTime].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertTime, keyCodes.KeySemicolon, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EditSelectAll].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EditSelectAll, keyCodes.KeyA, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText, keyCodes.Home, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveEndText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveEndText, keyCodes.End, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText, keyCodes.Home, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectEndText].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectEndText, keyCodes.End, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine, keyCodes.ArrowLeft, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveEndLine].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveEndLine, keyCodes.ArrowRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow, keyCodes.KeyR, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog, keyCodes.Digit1, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Subscript].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Subscript, keyCodes.Period, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Superscript].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Superscript, keyCodes.Comma, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize, keyCodes.BracketRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize, keyCodes.BracketLeft, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CenterPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CenterPara, keyCodes.KeyE, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.JustifyPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.JustifyPara, keyCodes.KeyJ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RightPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RightPara, keyCodes.KeyR, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.LeftPara].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.LeftPara, keyCodes.KeyL, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SpeechWorker].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SpeechWorker, keyCodes.KeyZ, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.EnDash].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.EnDash, keyCodes.KeyMinus, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft, keyCodes.ArrowLeft, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown, keyCodes.ArrowDown, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp, keyCodes.ArrowUp, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet, keyCodes.PageUp, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.NextWorksheet].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NextWorksheet, keyCodes.PageDown, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat, keyCodes.KeyBackquote, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat, keyCodes.Digit4, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat, keyCodes.Digit5, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat, keyCodes.Digit6, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellDateFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellDateFormat, keyCodes.Digit3, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat, keyCodes.Digit2, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat, keyCodes.Digit1, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Italic].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Italic, keyCodes.KeyI, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots, keyCodes.F5, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.ShowFormulas].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.ShowFormulas, keyCodes.KeyBackquote, false, false, false, true));
	} else {
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel, keyCodes.KeyF, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenCommentField] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenCommentField, keyCodes.KeyH, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel, keyCodes.KeyQ, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Cut] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Cut, keyCodes.KeyX, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Cut, keyCodes.Delete, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Copy] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Copy, keyCodes.KeyC, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Copy, keyCodes.Insert, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.Paste] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Paste, keyCodes.KeyV, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.Paste, keyCodes.Insert, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord, keyCodes.Backspace, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord, keyCodes.Delete, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveWordRight, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft, keyCodes.ArrowLeft, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectWordRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectWordRight, keyCodes.ArrowRight, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues, keyCodes.KeyC, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect, keyCodes.KeyS, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction, keyCodes.KeyEqual, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank, keyCodes.ArrowRight, true, false, false, false));
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight].push(new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight, keyCodes.ArrowRight, true, true, false, false));
	}

	if (window["AscDesktopEditor"]) {
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.NextFileTab] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.NextFileTab, keyCodes.Tab, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.PreviousFileTab] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.PreviousFileTab, keyCodes.Tab, true, true, false, false)];
		if (AscCommon.AscBrowser.isMacOs) {
			c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile, keyCodes.KeyO, false, false, false, true)];
			c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CloseFile, keyCodes.KeyW, false, false, false, true)];
		} else {
			c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile, keyCodes.KeyO, true, false, false, false)];
			c_oAscDefaultShortcuts[Asc.c_oAscSpreadsheetShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CloseFile, keyCodes.KeyW, true, false, false, false), new AscShortcut(Asc.c_oAscSpreadsheetShortcutType.CloseFile, keyCodes.F4, true, false, false, false)];
		}
	}

	function getStringFromShortcutType(type) {
		switch (type) {
			case Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel:
				return "OpenFilePanel";
			case Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog:
				return "OpenFindDialog";
			case Asc.c_oAscSpreadsheetShortcutType.OpenFindAndReplaceMenu:
				return "OpenFindAndReplaceMenu";
			case Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel:
				return "OpenCommentsPanel";
			case Asc.c_oAscSpreadsheetShortcutType.OpenCommentField:
				return "OpenCommentField";
			case Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel:
				return "OpenChatPanel";
			case Asc.c_oAscSpreadsheetShortcutType.Save:
				return "Save";
			case Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint:
				return "PrintPreviewAndPrint";
			case Asc.c_oAscSpreadsheetShortcutType.DownloadAs:
				return "DownloadAs";
			case Asc.c_oAscSpreadsheetShortcutType.OpenHelpMenu:
				return "OpenHelpMenu";
			case Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile:
				return "OpenExistingFile";
			case Asc.c_oAscSpreadsheetShortcutType.NextFileTab:
				return "NextFileTab";
			case Asc.c_oAscSpreadsheetShortcutType.PreviousFileTab:
				return "PreviousFileTab";
			case Asc.c_oAscSpreadsheetShortcutType.CloseFile:
				return "CloseFile";
			case Asc.c_oAscSpreadsheetShortcutType.OpenContextMenu:
				return "OpenContextMenu";
			case Asc.c_oAscSpreadsheetShortcutType.CloseMenu:
				return "CloseMenu";
			case Asc.c_oAscSpreadsheetShortcutType.Zoom100:
				return "Zoom100";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveUp:
				return "CellMoveUp";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveDown:
				return "CellMoveDown";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveLeft:
				return "CellMoveLeft";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveRight:
				return "CellMoveRight";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellDown:
				return "CellMoveActiveCellDown";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellUp:
				return "CellMoveActiveCellUp";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellRight:
				return "CellMoveActiveCellRight";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellLeft:
				return "CellMoveActiveCellLeft";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank:
				return "CellMoveLeftNonBlank";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstColumn:
				return "CellMoveFirstColumn";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank:
				return "CellMoveRightNonBlank";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank:
				return "CellMoveBottomNonBlank";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomEdge:
				return "CellMoveBottomEdge";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank:
				return "CellMoveTopNonBlank";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveTopEdge:
				return "CellMoveTopEdge";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell:
				return "CellMoveFirstCell";
			case Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet:
				return "CellMoveEndSpreadsheet";
			case Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet:
				return "PreviousWorksheet";
			case Asc.c_oAscSpreadsheetShortcutType.NextWorksheet:
				return "NextWorksheet";
			case Asc.c_oAscSpreadsheetShortcutType.ZoomIn:
				return "ZoomIn";
			case Asc.c_oAscSpreadsheetShortcutType.ZoomOut:
				return "ZoomOut";
			case Asc.c_oAscSpreadsheetShortcutType.NavigatePreviousControl:
				return "NavigatePreviousControl";
			case Asc.c_oAscSpreadsheetShortcutType.NavigateNextControl:
				return "NavigateNextControl";
			case Asc.c_oAscSpreadsheetShortcutType.SelectColumn:
				return "SelectColumn";
			case Asc.c_oAscSpreadsheetShortcutType.SelectRow:
				return "SelectRow";
			case Asc.c_oAscSpreadsheetShortcutType.SelectOneCellRight:
				return "SelectOneCellRight";
			case Asc.c_oAscSpreadsheetShortcutType.SelectOneCellLeft:
				return "SelectOneCellLeft";
			case Asc.c_oAscSpreadsheetShortcutType.SelectOneCellUp:
				return "SelectOneCellUp";
			case Asc.c_oAscSpreadsheetShortcutType.SelectOneCellDown:
				return "SelectOneCellDown";
			case Asc.c_oAscSpreadsheetShortcutType.SelectCursorBeginningRow:
				return "SelectCursorBeginningRow";
			case Asc.c_oAscSpreadsheetShortcutType.SelectCursorEndRow:
				return "SelectCursorEndRow";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight:
				return "SelectNextNonblankRight";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft:
				return "SelectNextNonblankLeft";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp:
				return "SelectNextNonblankUp";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown:
				return "SelectNextNonblankDown";
			case Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet:
				return "SelectBeginningWorksheet";
			case Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell:
				return "SelectLastUsedCell";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight:
				return "SelectNearestNonblankRight";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft:
				return "SelectNonblankLeft";
			case Asc.c_oAscSpreadsheetShortcutType.SelectFirstColumn:
				return "SelectFirstColumn";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown:
				return "SelectNearestNonblankDown";
			case Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp:
				return "SelectNearestNonblankUp";
			case Asc.c_oAscSpreadsheetShortcutType.SelectDownOneScreen:
				return "SelectDownOneScreen";
			case Asc.c_oAscSpreadsheetShortcutType.SelectUpOneScreen:
				return "SelectUpOneScreen";
			case Asc.c_oAscSpreadsheetShortcutType.EditUndo:
				return "EditUndo";
			case Asc.c_oAscSpreadsheetShortcutType.EditRedo:
				return "EditRedo";
			case Asc.c_oAscSpreadsheetShortcutType.Cut:
				return "Cut";
			case Asc.c_oAscSpreadsheetShortcutType.Copy:
				return "Copy";
			case Asc.c_oAscSpreadsheetShortcutType.Paste:
				return "Paste";
			case Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormula:
				return "PasteOnlyFormula";
			case Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNumberFormat:
				return "PasteFormulaNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutType.PasteFormulaAllFormatting:
				return "PasteFormulaAllFormatting";
			case Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNoBorders:
				return "PasteFormulaNoBorders";
			case Asc.c_oAscSpreadsheetShortcutType.PasteFormulaColumnWidth:
				return "PasteFormulaColumnWidth";
			case Asc.c_oAscSpreadsheetShortcutType.Transpose:
				return "Transpose";
			case Asc.c_oAscSpreadsheetShortcutType.PasteOnlyValue:
				return "PasteOnlyValue";
			case Asc.c_oAscSpreadsheetShortcutType.PasteValueNumberFormat:
				return "PasteValueNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutType.PasteValueAllFormatting:
				return "PasteValueAllFormatting";
			case Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormatting:
				return "PasteOnlyFormatting";
			case Asc.c_oAscSpreadsheetShortcutType.PasteLink:
				return "PasteLink";
			case Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink:
				return "InsertHyperlink";
			case Asc.c_oAscSpreadsheetShortcutType.VisitHyperlink:
				return "VisitHyperlink";
			case Asc.c_oAscSpreadsheetShortcutType.Bold:
				return "Bold";
			case Asc.c_oAscSpreadsheetShortcutType.Italic:
				return "Italic";
			case Asc.c_oAscSpreadsheetShortcutType.Underline:
				return "Underline";
			case Asc.c_oAscSpreadsheetShortcutType.Strikeout:
				return "Strikeout";
			case Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor:
				return "EditOpenCellEditor";
			case Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter:
				return "ToggleAutoFilter";
			case Asc.c_oAscSpreadsheetShortcutType.OpenFilterWindow:
				return "OpenFilterWindow";
			case Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate:
				return "FormatAsTableTemplate";
			case Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveDown:
				return "CompleteCellEntryMoveDown";
			case Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveUp:
				return "CompleteCellEntryMoveUp";
			case Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveRight:
				return "CompleteCellEntryMoveRight";
			case Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveLeft:
				return "CompleteCellEntryMoveLeft";
			case Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryStay:
				return "CompleteCellEntryStay";
			case Asc.c_oAscSpreadsheetShortcutType.FillSelectedCellRange:
				return "FillSelectedCellRange";
			case Asc.c_oAscSpreadsheetShortcutType.CellStartNewLine:
				return "CellStartNewLine";
			case Asc.c_oAscSpreadsheetShortcutType.EquationAddPlaceholder:
				return "EquationAddPlaceholder";
			case Asc.c_oAscSpreadsheetShortcutType.CellEntryCancel:
				return "CellEntryCancel";
			case Asc.c_oAscSpreadsheetShortcutType.DeleteLeftChar:
				return "DeleteLeftChar";
			case Asc.c_oAscSpreadsheetShortcutType.DeleteRightChar:
				return "DeleteRightChar";
			case Asc.c_oAscSpreadsheetShortcutType.ClearActiveCellContent:
				return "ClearActiveCellContent";
			case Asc.c_oAscSpreadsheetShortcutType.ClearSelectedCellsContent:
				return "ClearSelectedCellsContent";
			case Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow:
				return "OpenInsertCellsWindow";
			case Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow:
				return "OpenDeleteCellsWindow";
			case Asc.c_oAscSpreadsheetShortcutType.CellInsertDate:
				return "CellInsertDate";
			case Asc.c_oAscSpreadsheetShortcutType.CellInsertTime:
				return "CellInsertTime";
			case Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator:
				return "CellAddSeparator";
			case Asc.c_oAscSpreadsheetShortcutType.AutoFill:
				return "AutoFill";
			case Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord:
				return "DeleteLeftWord";
			case Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord:
				return "DeleteRightWord";
			case Asc.c_oAscSpreadsheetShortcutType.EditSelectAll:
				return "EditSelectAll";
			case Asc.c_oAscSpreadsheetShortcutType.MoveCharacterLeft:
				return "MoveCharacterLeft";
			case Asc.c_oAscSpreadsheetShortcutType.MoveCharacterRight:
				return "MoveCharacterRight";
			case Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineUp:
				return "MoveCursorLineUp";
			case Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineDown:
				return "MoveCursorLineDown";
			case Asc.c_oAscSpreadsheetShortcutType.SelectCharacterRight:
				return "SelectCharacterRight";
			case Asc.c_oAscSpreadsheetShortcutType.SelectCharacterLeft:
				return "SelectCharacterLeft";
			case Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft:
				return "MoveWordLeft";
			case Asc.c_oAscSpreadsheetShortcutType.MoveWordRight:
				return "MoveWordRight";
			case Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft:
				return "SelectWordLeft";
			case Asc.c_oAscSpreadsheetShortcutType.SelectWordRight:
				return "SelectWordRight";
			case Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText:
				return "MoveBeginningText";
			case Asc.c_oAscSpreadsheetShortcutType.MoveEndText:
				return "MoveEndText";
			case Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText:
				return "SelectBeginningText";
			case Asc.c_oAscSpreadsheetShortcutType.SelectEndText:
				return "SelectEndText";
			case Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine:
				return "MoveBeginningLine";
			case Asc.c_oAscSpreadsheetShortcutType.MoveEndLine:
				return "MoveEndLine";
			case Asc.c_oAscSpreadsheetShortcutType.SelectBeginningLine:
				return "SelectBeginningLine";
			case Asc.c_oAscSpreadsheetShortcutType.SelectEndLine:
				return "SelectEndLine";
			case Asc.c_oAscSpreadsheetShortcutType.SelectLineUp:
				return "SelectLineUp";
			case Asc.c_oAscSpreadsheetShortcutType.SelectLineDown:
				return "SelectLineDown";
			case Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots:
				return "RefreshSelectedPivots";
			case Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots:
				return "RefreshAllPivots";
			case Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues:
				return "SlicerClearSelectedValues";
			case Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect:
				return "SlicerSwitchMultiSelect";
			case Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow:
				return "FormatTableAddSummaryRow";
			case Asc.c_oAscSpreadsheetShortcutType.OpenInsertFunctionDialog:
				return "OpenInsertFunctionDialog";
			case Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction:
				return "CellInsertSumFunction";
			case Asc.c_oAscSpreadsheetShortcutType.RecalculateAll:
				return "RecalculateAll";
			case Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet:
				return "RecalculateActiveSheet";
			case Asc.c_oAscSpreadsheetShortcutType.ShowFormulas:
				return "ShowFormulas";
			case Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference:
				return "CellEditorSwitchReference";
			case Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog:
				return "OpenNumberFormatDialog";
			case Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat:
				return "CellGeneralFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat:
				return "CellCurrencyFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat:
				return "CellPercentFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat:
				return "CellExponentialFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellDateFormat:
				return "CellDateFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat:
				return "CellTimeFormat";
			case Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat:
				return "CellNumberFormat";
			case Asc.c_oAscSpreadsheetShortcutType.EditShape:
				return "EditShape";
			case Asc.c_oAscSpreadsheetShortcutType.EditChart:
				return "EditChart";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight:
				return "MoveShapeLittleStepRight";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft:
				return "MoveShapeLittleStepLeft";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp:
				return "MoveShapeLittleStepUp";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom:
				return "MoveShapeLittleStepBottom";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepLeft:
				return "MoveShapeBigStepLeft";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepRight:
				return "MoveShapeBigStepRight";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepUp:
				return "MoveShapeBigStepUp";
			case Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepBottom:
				return "MoveShapeBigStepBottom";
			case Asc.c_oAscSpreadsheetShortcutType.MoveFocusNextObject:
				return "MoveFocusNextObject";
			case Asc.c_oAscSpreadsheetShortcutType.MoveFocusPreviousObject:
				return "MoveFocusPreviousObject";
			case Asc.c_oAscSpreadsheetShortcutType.DrawingAddTab:
				return "DrawingAddTab";
			case Asc.c_oAscSpreadsheetShortcutType.Subscript:
				return "Subscript";
			case Asc.c_oAscSpreadsheetShortcutType.Superscript:
				return "Superscript";
			case Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize:
				return "IncreaseFontSize";
			case Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize:
				return "DecreaseFontSize";
			case Asc.c_oAscSpreadsheetShortcutType.CenterPara:
				return "CenterPara";
			case Asc.c_oAscSpreadsheetShortcutType.JustifyPara:
				return "JustifyPara";
			case Asc.c_oAscSpreadsheetShortcutType.RightPara:
				return "RightPara";
			case Asc.c_oAscSpreadsheetShortcutType.LeftPara:
				return "LeftPara";
			case Asc.c_oAscSpreadsheetShortcutType.EndParagraph:
				return "EndParagraph";
			case Asc.c_oAscSpreadsheetShortcutType.AddLineBreak:
				return "AddLineBreak";
			case Asc.c_oAscSpreadsheetShortcutType.RemoveGraphicalObject:
				return "RemoveGraphicalObject";
			case Asc.c_oAscSpreadsheetShortcutType.ExitAddingShapesMode:
				return "ExitAddingShapesMode";
			case Asc.c_oAscSpreadsheetShortcutType.SpeechWorker:
				return "SpeechWorker";
			case Asc.c_oAscSpreadsheetShortcutType.EnDash:
				return "EnDash";

			default:
				return null;
		}
	}

	function getShortcutTypeFromString(str) {
		switch (str) {
			case "OpenFilePanel":
				return Asc.c_oAscSpreadsheetShortcutType.OpenFilePanel;
			case "OpenFindDialog":
				return Asc.c_oAscSpreadsheetShortcutType.OpenFindDialog;
			case "OpenFindAndReplaceMenu":
				return Asc.c_oAscSpreadsheetShortcutType.OpenFindAndReplaceMenu;
			case "OpenCommentsPanel":
				return Asc.c_oAscSpreadsheetShortcutType.OpenCommentsPanel;
			case "OpenCommentField":
				return Asc.c_oAscSpreadsheetShortcutType.OpenCommentField;
			case "OpenChatPanel":
				return Asc.c_oAscSpreadsheetShortcutType.OpenChatPanel;
			case "Save":
				return Asc.c_oAscSpreadsheetShortcutType.Save;
			case "PrintPreviewAndPrint":
				return Asc.c_oAscSpreadsheetShortcutType.PrintPreviewAndPrint;
			case "DownloadAs":
				return Asc.c_oAscSpreadsheetShortcutType.DownloadAs;
			case "OpenHelpMenu":
				return Asc.c_oAscSpreadsheetShortcutType.OpenHelpMenu;
			case "OpenExistingFile":
				return Asc.c_oAscSpreadsheetShortcutType.OpenExistingFile;
			case "NextFileTab":
				return Asc.c_oAscSpreadsheetShortcutType.NextFileTab;
			case "PreviousFileTab":
				return Asc.c_oAscSpreadsheetShortcutType.PreviousFileTab;
			case "CloseFile":
				return Asc.c_oAscSpreadsheetShortcutType.CloseFile;
			case "OpenContextMenu":
				return Asc.c_oAscSpreadsheetShortcutType.OpenContextMenu;
			case "CloseMenu":
				return Asc.c_oAscSpreadsheetShortcutType.CloseMenu;
			case "Zoom100":
				return Asc.c_oAscSpreadsheetShortcutType.Zoom100;
			case "CellMoveUp":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveUp;
			case "CellMoveDown":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveDown;
			case "CellMoveLeft":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveLeft;
			case "CellMoveRight":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveRight;
			case "CellMoveActiveCellDown":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellDown;
			case "CellMoveActiveCellUp":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellUp;
			case "CellMoveActiveCellRight":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellRight;
			case "CellMoveActiveCellLeft":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveActiveCellLeft;
			case "CellMoveLeftNonBlank":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveLeftNonBlank;
			case "CellMoveFirstColumn":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstColumn;
			case "CellMoveRightNonBlank":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveRightNonBlank;
			case "CellMoveBottomNonBlank":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomNonBlank;
			case "CellMoveBottomEdge":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveBottomEdge;
			case "CellMoveTopNonBlank":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveTopNonBlank;
			case "CellMoveTopEdge":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveTopEdge;
			case "CellMoveFirstCell":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveFirstCell;
			case "CellMoveEndSpreadsheet":
				return Asc.c_oAscSpreadsheetShortcutType.CellMoveEndSpreadsheet;
			case "PreviousWorksheet":
				return Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet;
			case "NextWorksheet":
				return Asc.c_oAscSpreadsheetShortcutType.NextWorksheet;
			case "ZoomIn":
				return Asc.c_oAscSpreadsheetShortcutType.ZoomIn;
			case "ZoomOut":
				return Asc.c_oAscSpreadsheetShortcutType.ZoomOut;
			case "NavigatePreviousControl":
				return Asc.c_oAscSpreadsheetShortcutType.NavigatePreviousControl;
			case "NavigateNextControl":
				return Asc.c_oAscSpreadsheetShortcutType.NavigateNextControl;
			case "SelectColumn":
				return Asc.c_oAscSpreadsheetShortcutType.SelectColumn;
			case "SelectRow":
				return Asc.c_oAscSpreadsheetShortcutType.SelectRow;
			case "SelectOneCellRight":
				return Asc.c_oAscSpreadsheetShortcutType.SelectOneCellRight;
			case "SelectOneCellLeft":
				return Asc.c_oAscSpreadsheetShortcutType.SelectOneCellLeft;
			case "SelectOneCellUp":
				return Asc.c_oAscSpreadsheetShortcutType.SelectOneCellUp;
			case "SelectOneCellDown":
				return Asc.c_oAscSpreadsheetShortcutType.SelectOneCellDown;
			case "SelectCursorBeginningRow":
				return Asc.c_oAscSpreadsheetShortcutType.SelectCursorBeginningRow;
			case "SelectCursorEndRow":
				return Asc.c_oAscSpreadsheetShortcutType.SelectCursorEndRow;
			case "SelectNextNonblankRight":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankRight;
			case "SelectNextNonblankLeft":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankLeft;
			case "SelectNextNonblankUp":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankUp;
			case "SelectNextNonblankDown":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNextNonblankDown;
			case "SelectBeginningWorksheet":
				return Asc.c_oAscSpreadsheetShortcutType.SelectBeginningWorksheet;
			case "SelectLastUsedCell":
				return Asc.c_oAscSpreadsheetShortcutType.SelectLastUsedCell;
			case "SelectNearestNonblankRight":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankRight;
			case "SelectNonblankLeft":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNonblankLeft;
			case "SelectFirstColumn":
				return Asc.c_oAscSpreadsheetShortcutType.SelectFirstColumn;
			case "SelectNearestNonblankDown":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankDown;
			case "SelectNearestNonblankUp":
				return Asc.c_oAscSpreadsheetShortcutType.SelectNearestNonblankUp;
			case "SelectDownOneScreen":
				return Asc.c_oAscSpreadsheetShortcutType.SelectDownOneScreen;
			case "SelectUpOneScreen":
				return Asc.c_oAscSpreadsheetShortcutType.SelectUpOneScreen;
			case "EditUndo":
				return Asc.c_oAscSpreadsheetShortcutType.EditUndo;
			case "EditRedo":
				return Asc.c_oAscSpreadsheetShortcutType.EditRedo;
			case "Cut":
				return Asc.c_oAscSpreadsheetShortcutType.Cut;
			case "Copy":
				return Asc.c_oAscSpreadsheetShortcutType.Copy;
			case "Paste":
				return Asc.c_oAscSpreadsheetShortcutType.Paste;
			case "PasteOnlyFormula":
				return Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormula;
			case "PasteFormulaNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNumberFormat;
			case "PasteFormulaAllFormatting":
				return Asc.c_oAscSpreadsheetShortcutType.PasteFormulaAllFormatting;
			case "PasteFormulaNoBorders":
				return Asc.c_oAscSpreadsheetShortcutType.PasteFormulaNoBorders;
			case "PasteFormulaColumnWidth":
				return Asc.c_oAscSpreadsheetShortcutType.PasteFormulaColumnWidth;
			case "Transpose":
				return Asc.c_oAscSpreadsheetShortcutType.Transpose;
			case "PasteOnlyValue":
				return Asc.c_oAscSpreadsheetShortcutType.PasteOnlyValue;
			case "PasteValueNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutType.PasteValueNumberFormat;
			case "PasteValueAllFormatting":
				return Asc.c_oAscSpreadsheetShortcutType.PasteValueAllFormatting;
			case "PasteOnlyFormatting":
				return Asc.c_oAscSpreadsheetShortcutType.PasteOnlyFormatting;
			case "PasteLink":
				return Asc.c_oAscSpreadsheetShortcutType.PasteLink;
			case "InsertHyperlink":
				return Asc.c_oAscSpreadsheetShortcutType.InsertHyperlink;
			case "VisitHyperlink":
				return Asc.c_oAscSpreadsheetShortcutType.VisitHyperlink;
			case "Bold":
				return Asc.c_oAscSpreadsheetShortcutType.Bold;
			case "Italic":
				return Asc.c_oAscSpreadsheetShortcutType.Italic;
			case "Underline":
				return Asc.c_oAscSpreadsheetShortcutType.Underline;
			case "Strikeout":
				return Asc.c_oAscSpreadsheetShortcutType.Strikeout;
			case "EditOpenCellEditor":
				return Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor;
			case "ToggleAutoFilter":
				return Asc.c_oAscSpreadsheetShortcutType.ToggleAutoFilter;
			case "OpenFilterWindow":
				return Asc.c_oAscSpreadsheetShortcutType.OpenFilterWindow;
			case "FormatAsTableTemplate":
				return Asc.c_oAscSpreadsheetShortcutType.FormatAsTableTemplate;
			case "CompleteCellEntryMoveDown":
				return Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveDown;
			case "CompleteCellEntryMoveUp":
				return Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveUp;
			case "CompleteCellEntryMoveRight":
				return Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveRight;
			case "CompleteCellEntryMoveLeft":
				return Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryMoveLeft;
			case "CompleteCellEntryStay":
				return Asc.c_oAscSpreadsheetShortcutType.CompleteCellEntryStay;
			case "FillSelectedCellRange":
				return Asc.c_oAscSpreadsheetShortcutType.FillSelectedCellRange;
			case "CellStartNewLine":
				return Asc.c_oAscSpreadsheetShortcutType.CellStartNewLine;
			case "EquationAddPlaceholder":
				return Asc.c_oAscSpreadsheetShortcutType.EquationAddPlaceholder;
			case "CellEntryCancel":
				return Asc.c_oAscSpreadsheetShortcutType.CellEntryCancel;
			case "DeleteLeftChar":
				return Asc.c_oAscSpreadsheetShortcutType.DeleteLeftChar;
			case "DeleteRightChar":
				return Asc.c_oAscSpreadsheetShortcutType.DeleteRightChar;
			case "ClearActiveCellContent":
				return Asc.c_oAscSpreadsheetShortcutType.ClearActiveCellContent;
			case "ClearSelectedCellsContent":
				return Asc.c_oAscSpreadsheetShortcutType.ClearSelectedCellsContent;
			case "OpenInsertCellsWindow":
				return Asc.c_oAscSpreadsheetShortcutType.OpenInsertCellsWindow;
			case "OpenDeleteCellsWindow":
				return Asc.c_oAscSpreadsheetShortcutType.OpenDeleteCellsWindow;
			case "CellInsertDate":
				return Asc.c_oAscSpreadsheetShortcutType.CellInsertDate;
			case "CellInsertTime":
				return Asc.c_oAscSpreadsheetShortcutType.CellInsertTime;
			case "CellAddSeparator":
				return Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator;
			case "AutoFill":
				return Asc.c_oAscSpreadsheetShortcutType.AutoFill;
			case "DeleteLeftWord":
				return Asc.c_oAscSpreadsheetShortcutType.DeleteLeftWord;
			case "DeleteRightWord":
				return Asc.c_oAscSpreadsheetShortcutType.DeleteRightWord;
			case "EditSelectAll":
				return Asc.c_oAscSpreadsheetShortcutType.EditSelectAll;
			case "MoveCharacterLeft":
				return Asc.c_oAscSpreadsheetShortcutType.MoveCharacterLeft;
			case "MoveCharacterRight":
				return Asc.c_oAscSpreadsheetShortcutType.MoveCharacterRight;
			case "MoveCursorLineUp":
				return Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineUp;
			case "MoveCursorLineDown":
				return Asc.c_oAscSpreadsheetShortcutType.MoveCursorLineDown;
			case "SelectCharacterRight":
				return Asc.c_oAscSpreadsheetShortcutType.SelectCharacterRight;
			case "SelectCharacterLeft":
				return Asc.c_oAscSpreadsheetShortcutType.SelectCharacterLeft;
			case "MoveWordLeft":
				return Asc.c_oAscSpreadsheetShortcutType.MoveWordLeft;
			case "MoveWordRight":
				return Asc.c_oAscSpreadsheetShortcutType.MoveWordRight;
			case "SelectWordLeft":
				return Asc.c_oAscSpreadsheetShortcutType.SelectWordLeft;
			case "SelectWordRight":
				return Asc.c_oAscSpreadsheetShortcutType.SelectWordRight;
			case "MoveBeginningText":
				return Asc.c_oAscSpreadsheetShortcutType.MoveBeginningText;
			case "MoveEndText":
				return Asc.c_oAscSpreadsheetShortcutType.MoveEndText;
			case "SelectBeginningText":
				return Asc.c_oAscSpreadsheetShortcutType.SelectBeginningText;
			case "SelectEndText":
				return Asc.c_oAscSpreadsheetShortcutType.SelectEndText;
			case "MoveBeginningLine":
				return Asc.c_oAscSpreadsheetShortcutType.MoveBeginningLine;
			case "MoveEndLine":
				return Asc.c_oAscSpreadsheetShortcutType.MoveEndLine;
			case "SelectBeginningLine":
				return Asc.c_oAscSpreadsheetShortcutType.SelectBeginningLine;
			case "SelectEndLine":
				return Asc.c_oAscSpreadsheetShortcutType.SelectEndLine;
			case "SelectLineUp":
				return Asc.c_oAscSpreadsheetShortcutType.SelectLineUp;
			case "SelectLineDown":
				return Asc.c_oAscSpreadsheetShortcutType.SelectLineDown;
			case "RefreshSelectedPivots":
				return Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots;
			case "RefreshAllPivots":
				return Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots;
			case "SlicerClearSelectedValues":
				return Asc.c_oAscSpreadsheetShortcutType.SlicerClearSelectedValues;
			case "SlicerSwitchMultiSelect":
				return Asc.c_oAscSpreadsheetShortcutType.SlicerSwitchMultiSelect;
			case "FormatTableAddSummaryRow":
				return Asc.c_oAscSpreadsheetShortcutType.FormatTableAddSummaryRow;
			case "OpenInsertFunctionDialog":
				return Asc.c_oAscSpreadsheetShortcutType.OpenInsertFunctionDialog;
			case "CellInsertSumFunction":
				return Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction;
			case "RecalculateAll":
				return Asc.c_oAscSpreadsheetShortcutType.RecalculateAll;
			case "RecalculateActiveSheet":
				return Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet;
			case "ShowFormulas":
				return Asc.c_oAscSpreadsheetShortcutType.ShowFormulas;
			case "CellEditorSwitchReference":
				return Asc.c_oAscSpreadsheetShortcutType.CellEditorSwitchReference;
			case "OpenNumberFormatDialog":
				return Asc.c_oAscSpreadsheetShortcutType.OpenNumberFormatDialog;
			case "CellGeneralFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat;
			case "CellCurrencyFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat;
			case "CellPercentFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat;
			case "CellExponentialFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat;
			case "CellDateFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellDateFormat;
			case "CellTimeFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat;
			case "CellNumberFormat":
				return Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat;
			case "EditShape":
				return Asc.c_oAscSpreadsheetShortcutType.EditShape;
			case "EditChart":
				return Asc.c_oAscSpreadsheetShortcutType.EditChart;
			case "MoveShapeLittleStepRight":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepRight;
			case "MoveShapeLittleStepLeft":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepLeft;
			case "MoveShapeLittleStepUp":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepUp;
			case "MoveShapeLittleStepBottom":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeLittleStepBottom;
			case "MoveShapeBigStepLeft":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepLeft;
			case "MoveShapeBigStepRight":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepRight;
			case "MoveShapeBigStepUp":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepUp;
			case "MoveShapeBigStepBottom":
				return Asc.c_oAscSpreadsheetShortcutType.MoveShapeBigStepBottom;
			case "MoveFocusNextObject":
				return Asc.c_oAscSpreadsheetShortcutType.MoveFocusNextObject;
			case "MoveFocusPreviousObject":
				return Asc.c_oAscSpreadsheetShortcutType.MoveFocusPreviousObject;
			case "DrawingAddTab":
				return Asc.c_oAscSpreadsheetShortcutType.DrawingAddTab;
			case "Subscript":
				return Asc.c_oAscSpreadsheetShortcutType.Subscript;
			case "Superscript":
				return Asc.c_oAscSpreadsheetShortcutType.Superscript;
			case "IncreaseFontSize":
				return Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize;
			case "DecreaseFontSize":
				return Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize;
			case "CenterPara":
				return Asc.c_oAscSpreadsheetShortcutType.CenterPara;
			case "JustifyPara":
				return Asc.c_oAscSpreadsheetShortcutType.JustifyPara;
			case "RightPara":
				return Asc.c_oAscSpreadsheetShortcutType.RightPara;
			case "LeftPara":
				return Asc.c_oAscSpreadsheetShortcutType.LeftPara;
			case "EndParagraph":
				return Asc.c_oAscSpreadsheetShortcutType.EndParagraph;
			case "AddLineBreak":
				return Asc.c_oAscSpreadsheetShortcutType.AddLineBreak;
			case "RemoveGraphicalObject":
				return Asc.c_oAscSpreadsheetShortcutType.RemoveGraphicalObject;
			case "ExitAddingShapesMode":
				return Asc.c_oAscSpreadsheetShortcutType.ExitAddingShapesMode;
			case "SpeechWorker":
				return Asc.c_oAscSpreadsheetShortcutType.SpeechWorker;
			case "EnDash":
				return Asc.c_oAscSpreadsheetShortcutType.EnDash;
			default:
				return null;
		}
	}


	window["Asc"]["c_oAscDefaultShortcuts"] = window["Asc"].c_oAscDefaultShortcuts = c_oAscDefaultShortcuts;
	window["Asc"]["c_oAscUnlockedShortcutActionTypes"] = window["Asc"].c_oAscUnlockedShortcutActionTypes = c_oAscUnlockedShortcutActionTypes;
	window["AscCommon"].getStringFromShortcutType = getStringFromShortcutType;
	window["AscCommon"].getShortcutTypeFromString = getShortcutTypeFromString;
})();
