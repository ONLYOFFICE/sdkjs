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
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.EditSelectAll] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.EditUndo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.EditRedo] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Duplicate] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Save] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.OpenContextMenu] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.ShowParaMarks] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Bold] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.CopyFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.CenterPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.EuroSign] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Group] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.UnGroup] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Italic] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.JustifyPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.InsertHyperlink] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.BulletList] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.LeftPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.RightPara] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Underline] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Strikeout] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.PasteFormat] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Superscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.Subscript] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.EnDash] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.DecreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.IncreaseFontSize] = true;
	c_oAscUnlockedShortcutActionTypes[Asc.c_oAscPresentationShortcutType.SpeechWorker] = true;

	const c_oAscDefaultShortcuts = {};
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindDialog] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenFindDialog, keyCodes.KeyF, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu, keyCodes.KeyH, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentsPanel] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenCommentsPanel, keyCodes.KeyH, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentField] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenCommentField, keyCodes.KeyH, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenChatPanel] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenChatPanel, keyCodes.KeyQ, false, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Save] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Save, keyCodes.KeyS, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint, keyCodes.KeyP, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SaveAs] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SaveAs, keyCodes.KeyS, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenHelpMenu] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenHelpMenu, keyCodes.F1, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenContextMenu] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenContextMenu, keyCodes.F10, false, true, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenContextMenu, keyCodes.ContextMenu, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseMenu] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseMenu, keyCodes.Escape, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Zoom100] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Zoom100, keyCodes.Digit0, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToFirstSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToFirstSlide, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToLastSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToLastSlide, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToNextSlide, keyCodes.PageDown, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToNextSlide, keyCodes.ArrowDown, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToNextSlide, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToPreviousSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToPreviousSlide, keyCodes.PageUp, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToPreviousSlide, keyCodes.ArrowUp, false, false, false, false),new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToPreviousSlide, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomIn] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.ZoomIn, keyCodes.KeyEqual, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomOut] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.ZoomOut, keyCodes.KeyMinus, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NextModalControl] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.NextModalControl, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PreviousModalControl] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PreviousModalControl, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NewSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.NewSlide, keyCodes.KeyM, true, false, false, false),new AscShortcut(Asc.c_oAscPresentationShortcutType.NewSlide, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RemoveSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.RemoveSlide, keyCodes.Delete, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.RemoveSlide, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Duplicate] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Duplicate, keyCodes.KeyD, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToBegin] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToBegin, keyCodes.ArrowUp, true, true, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToBegin, keyCodes.PageUp, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToEnd] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToEnd, keyCodes.ArrowDown, true, true, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToEnd, keyCodes.PageDown, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditShape] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditShape, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditChart] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditChart, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Group] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Group, keyCodes.KeyG, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnGroup] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.UnGroup, keyCodes.KeyG, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectRight] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.BigMoveObjectRight, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.BigMoveObjectUp, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BigMoveObjectDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.BigMoveObjectDown, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToNextCell] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToNextCell, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToPreviousCell] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToPreviousCell, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToNextRow] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToNextRow, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToPreviousRow] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToPreviousRow, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EndParagraphCell] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EndParagraphCell, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.AddNewRow] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.AddNewRow, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, keyCodes.Enter, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, keyCodes.PageDown, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, keyCodes.ArrowRight, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, keyCodes.ArrowDown, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide, keyCodes.Space, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide, keyCodes.PageUp, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide, keyCodes.ArrowLeft, false, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview, keyCodes.Escape, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditUndo] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditUndo, keyCodes.KeyZ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditRedo] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditRedo, keyCodes.KeyY, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CopyFormat] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CopyFormat, keyCodes.KeyC, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteFormat] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PasteFormat, keyCodes.KeyV, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UseDestinationTheme] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.UseDestinationTheme, keyCodes.KeyH, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.KeepSourceFormat] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.KeepSourceFormat, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteAsPicture] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PasteAsPicture, keyCodes.KeyU, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.KeepTextOnly] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.KeepTextOnly, keyCodes.KeyT, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertHyperlink] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.InsertHyperlink, keyCodes.KeyK, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.VisitHyperlink] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.VisitHyperlink, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditSelectAll] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditSelectAll, keyCodes.KeyA, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectNextSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectNextSlide, keyCodes.PageDown, false, true, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectNextSlide, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectPreviousSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectPreviousSlide, keyCodes.PageUp, false, true, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectPreviousSlide, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToFirstSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectToFirstSlide, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToLastSlide] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectToLastSlide, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToStartLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectToStartLine, keyCodes.Home, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectToEndLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectToEndLine, keyCodes.End, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectRightChar, keyCodes.ArrowRight, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectLeftChar, keyCodes.ArrowLeft, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLineUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectLineUp, keyCodes.ArrowUp, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLineDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectLineDown, keyCodes.ArrowDown, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditDeselectAll] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EditDeselectAll, keyCodes.Escape, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ShowParaMarks] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.ShowParaMarks, keyCodes.Digit8, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Bold] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Bold, keyCodes.KeyB, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Italic] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Italic, keyCodes.KeyI, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Underline] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Underline, keyCodes.KeyU, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Strikeout] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Strikeout, keyCodes.Digit5, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Subscript] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Subscript, keyCodes.Period, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Superscript] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Superscript, keyCodes.Comma, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BulletList] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.BulletList, keyCodes.KeyL, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ResetChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.ResetChar, keyCodes.Space, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.IncreaseFontSize] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFontSize, keyCodes.BracketRight, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DecreaseFontSize] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFontSize, keyCodes.BracketLeft, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CenterPara] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CenterPara, keyCodes.KeyE, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.JustifyPara] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.JustifyPara, keyCodes.KeyJ, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RightPara] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.RightPara, keyCodes.KeyR, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LeftPara] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LeftPara, keyCodes.KeyL, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Indent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Indent, keyCodes.KeyM, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnIndent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.UnIndent, keyCodes.KeyM, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteLeftChar, keyCodes.Backspace, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteRightChar, keyCodes.Delete, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.StartIndent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.StartIndent, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.StartUnIndent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.StartUnIndent, keyCodes.Tab, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertTab] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.InsertTab, keyCodes.Tab, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertLineBreak] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.InsertLineBreak, keyCodes.Enter, false, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EndParagraph] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EndParagraph, keyCodes.Enter, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EuroSign] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EuroSign, keyCodes.KeyE, true, false, true, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NonBreakingSpace] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.NonBreakingSpace, keyCodes.Space, true, true, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToRightChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToRightChar, keyCodes.ArrowRight, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToLeftChar] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToLeftChar, keyCodes.ArrowLeft, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToUpLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToUpLine, keyCodes.ArrowUp, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToDownLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToDownLine, keyCodes.ArrowDown, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder, keyCodes.Enter, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToStartLine, keyCodes.Home, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndLine] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToEndLine, keyCodes.End, false, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartContent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToStartContent, keyCodes.Home, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndContent] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToEndContent, keyCodes.End, true, false, false, false)];
	c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SpeechWorker] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SpeechWorker, keyCodes.KeyZ, true, false, true, false)];

	if (AscCommon.AscBrowser.isMacOs) {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFilePanel] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenFilePanel, keyCodes.KeyF, true, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideUp, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideDown, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft, keyCodes.ArrowLeft, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight, keyCodes.ArrowRight, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp, keyCodes.ArrowUp, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown, keyCodes.ArrowDown, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation, keyCodes.Enter, false, true, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Cut] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Cut, keyCodes.KeyX, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Copy] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Copy, keyCodes.KeyC, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Paste] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Paste, keyCodes.KeyV, false, false, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat, keyCodes.KeyV, false, true, false, true)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectRightWord, keyCodes.ArrowRight, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectLeftWord, keyCodes.ArrowLeft, false, true, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteLeftWord, keyCodes.Backspace, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteRightWord, keyCodes.Delete, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToStartWord, keyCodes.ArrowLeft, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToEndWord, keyCodes.ArrowRight, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EnDash] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EnDash, keyCodes.KeyMinus, false, false, true, false, true)];

		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFindDialog].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenFindDialog, keyCodes.KeyF, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentsPanel].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenCommentsPanel, keyCodes.KeyH, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenCommentField].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenCommentField, keyCodes.KeyA, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenChatPanel].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenChatPanel, keyCodes.KeyQ, true, false, true, false));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Save].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Save, keyCodes.KeyS, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint, keyCodes.KeyP, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SaveAs].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.SaveAs, keyCodes.KeyS, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Zoom100].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Zoom100, keyCodes.Digit0, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomIn].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.ZoomIn, keyCodes.KeyEqual, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ZoomOut].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.ZoomOut, keyCodes.KeyMinus, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NewSlide].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.NewSlide, keyCodes.KeyM, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Duplicate].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Duplicate, keyCodes.KeyD, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToBegin].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToBegin, keyCodes.ArrowUp, false, true, false, true), new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToBegin, keyCodes.PageUp, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideToEnd].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToEnd, keyCodes.ArrowDown, false, true, false, true), new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideToEnd, keyCodes.PageDown, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Group].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Group, keyCodes.KeyG, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnGroup].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.UnGroup, keyCodes.KeyG, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditUndo].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.EditUndo, keyCodes.KeyZ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditRedo].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.EditRedo, keyCodes.KeyY, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CopyFormat].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.CopyFormat, keyCodes.KeyC, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteFormat].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.PasteFormat, keyCodes.KeyV, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.InsertHyperlink].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.InsertHyperlink, keyCodes.KeyK, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EditSelectAll].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.EditSelectAll, keyCodes.KeyA, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ShowParaMarks].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.ShowParaMarks, keyCodes.Digit8, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Bold].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Bold, keyCodes.KeyB, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Italic].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Italic, keyCodes.KeyI, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Underline].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Underline, keyCodes.KeyU, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Strikeout].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Strikeout, keyCodes.Digit5, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Subscript].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Subscript, keyCodes.Period, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Superscript].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Superscript, keyCodes.Comma, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.BulletList].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.BulletList, keyCodes.KeyL, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.ResetChar].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.ResetChar, keyCodes.Space, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.IncreaseFontSize].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.IncreaseFontSize, keyCodes.BracketRight, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DecreaseFontSize].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.DecreaseFontSize, keyCodes.BracketLeft, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CenterPara].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.CenterPara, keyCodes.KeyE, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.JustifyPara].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.JustifyPara, keyCodes.KeyJ, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.RightPara].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.RightPara, keyCodes.KeyR, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LeftPara].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.LeftPara, keyCodes.KeyL, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Indent].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.Indent, keyCodes.KeyM, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.UnIndent].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.UnIndent, keyCodes.KeyM, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EuroSign].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.EuroSign, keyCodes.KeyE, false, false, true, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NonBreakingSpace].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.NonBreakingSpace, keyCodes.Space, false, true, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder, keyCodes.Enter, false, false, false, true));
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SpeechWorker].push(new AscShortcut(Asc.c_oAscPresentationShortcutType.SpeechWorker, keyCodes.KeyZ, false, false, true, true));
	} else {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenFilePanel] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenFilePanel, keyCodes.KeyF, false, false, true, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideUp, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveSlideDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveSlideDown, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp, keyCodes.ArrowUp, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown, keyCodes.ArrowDown, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation, keyCodes.F5, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Cut] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Cut, keyCodes.KeyX, true, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.Cut, keyCodes.Delete, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Copy] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Copy, keyCodes.KeyC, true, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.Copy, keyCodes.Insert, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.Paste] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.Paste, keyCodes.KeyV, true, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.Paste, keyCodes.Insert, false, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat, keyCodes.KeyV, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectRightWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectRightWord, keyCodes.ArrowRight, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.SelectLeftWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.SelectLeftWord, keyCodes.ArrowLeft, true, true, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteLeftWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteLeftWord, keyCodes.Backspace, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.DeleteRightWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.DeleteRightWord, keyCodes.Delete, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToStartWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToStartWord, keyCodes.ArrowLeft, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.MoveToEndWord] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.MoveToEndWord, keyCodes.ArrowRight, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.EnDash] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.EnDash, keyCodes.KeyMinus, true, true, false, false)];
	}

	if (window["AscDesktopEditor"]) {
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.NextFileTab] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.NextFileTab, keyCodes.Tab, true, false, false, false)];
		c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.PreviousFileTab] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.PreviousFileTab, keyCodes.Tab, true, true, false, false)];
		if (AscCommon.AscBrowser.isMacOs) {
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenExistingFile, keyCodes.KeyO, false, false, false, true)];
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.KeyW, false, false, false, true)];
		} else {
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.OpenExistingFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.OpenExistingFile, keyCodes.KeyO, true, false, false, false)];
			c_oAscDefaultShortcuts[Asc.c_oAscPresentationShortcutType.CloseFile] = [new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.KeyW, true, false, false, false), new AscShortcut(Asc.c_oAscPresentationShortcutType.CloseFile, keyCodes.F4, true, false, false, false)];
		}
	}
	function getStringFromShortcutType(type) {
		switch (type) {
			case Asc.c_oAscPresentationShortcutType.OpenFilePanel:
				return "OpenFilePanel";
			case Asc.c_oAscPresentationShortcutType.OpenFindDialog:
				return "OpenFindDialog";
			case Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu:
				return "OpenFindAndReplaceMenu";
			case Asc.c_oAscPresentationShortcutType.OpenCommentsPanel:
				return "OpenCommentsPanel";
			case Asc.c_oAscPresentationShortcutType.OpenCommentField:
				return "OpenCommentField";
			case Asc.c_oAscPresentationShortcutType.OpenChatPanel:
				return "OpenChatPanel";
			case Asc.c_oAscPresentationShortcutType.Save:
				return "Save";
			case Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint:
				return "PrintPreviewAndPrint";
			case Asc.c_oAscPresentationShortcutType.SaveAs:
				return "SaveAs";
			case Asc.c_oAscPresentationShortcutType.OpenHelpMenu:
				return "OpenHelpMenu";
			case Asc.c_oAscPresentationShortcutType.OpenExistingFile:
				return "OpenExistingFile";
			case Asc.c_oAscPresentationShortcutType.NextFileTab:
				return "NextFileTab";
			case Asc.c_oAscPresentationShortcutType.PreviousFileTab:
				return "PreviousFileTab";
			case Asc.c_oAscPresentationShortcutType.CloseFile:
				return "CloseFile";
			case Asc.c_oAscPresentationShortcutType.OpenContextMenu:
				return "OpenContextMenu";
			case Asc.c_oAscPresentationShortcutType.CloseMenu:
				return "CloseMenu";
			case Asc.c_oAscPresentationShortcutType.Zoom100:
				return "Zoom100";
			case Asc.c_oAscPresentationShortcutType.GoToFirstSlide:
				return "GoToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.GoToLastSlide:
				return "GoToLastSlide";
			case Asc.c_oAscPresentationShortcutType.GoToNextSlide:
				return "GoToNextSlide";
			case Asc.c_oAscPresentationShortcutType.GoToPreviousSlide:
				return "GoToPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.ZoomIn:
				return "ZoomIn";
			case Asc.c_oAscPresentationShortcutType.ZoomOut:
				return "ZoomOut";
			case Asc.c_oAscPresentationShortcutType.NextModalControl:
				return "NextModalControl";
			case Asc.c_oAscPresentationShortcutType.PreviousModalControl:
				return "PreviousModalControl";
			case Asc.c_oAscPresentationShortcutType.NewSlide:
				return "NewSlide";
			case Asc.c_oAscPresentationShortcutType.RemoveSlide:
				return "RemoveSlide";
			case Asc.c_oAscPresentationShortcutType.Duplicate:
				return "Duplicate";
			case Asc.c_oAscPresentationShortcutType.MoveSlideUp:
				return "MoveSlideUp";
			case Asc.c_oAscPresentationShortcutType.MoveSlideDown:
				return "MoveSlideDown";
			case Asc.c_oAscPresentationShortcutType.MoveSlideToBegin:
				return "MoveSlideToBegin";
			case Asc.c_oAscPresentationShortcutType.MoveSlideToEnd:
				return "MoveSlideToEnd";
			case Asc.c_oAscPresentationShortcutType.EditShape:
				return "EditShape";
			case Asc.c_oAscPresentationShortcutType.EditChart:
				return "EditChart";
			case Asc.c_oAscPresentationShortcutType.Group:
				return "Group";
			case Asc.c_oAscPresentationShortcutType.UnGroup:
				return "UnGroup";
			case Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject:
				return "MoveFocusToNextObject";
			case Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject:
				return "MoveFocusToPreviousObject";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft:
				return "LittleMoveObjectLeft";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight:
				return "LittleMoveObjectRight";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp:
				return "LittleMoveObjectUp";
			case Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown:
				return "LittleMoveObjectDown";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft:
				return "BigMoveObjectLeft";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectRight:
				return "BigMoveObjectRight";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectUp:
				return "BigMoveObjectUp";
			case Asc.c_oAscPresentationShortcutType.BigMoveObjectDown:
				return "BigMoveObjectDown";
			case Asc.c_oAscPresentationShortcutType.MoveToNextCell:
				return "MoveToNextCell";
			case Asc.c_oAscPresentationShortcutType.MoveToPreviousCell:
				return "MoveToPreviousCell";
			case Asc.c_oAscPresentationShortcutType.MoveToNextRow:
				return "MoveToNextRow";
			case Asc.c_oAscPresentationShortcutType.MoveToPreviousRow:
				return "MoveToPreviousRow";
			case Asc.c_oAscPresentationShortcutType.EndParagraphCell:
				return "EndParagraphCell";
			case Asc.c_oAscPresentationShortcutType.AddNewRow:
				return "AddNewRow";
			case Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation:
				return "DemonstrationStartPresentation";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide:
				return "DemonstrationGoToNextSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide:
				return "DemonstrationGoToPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide:
				return "DemonstrationGoToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide:
				return "DemonstrationGoToLastSlide";
			case Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview:
				return "DemonstrationClosePreview";
			case Asc.c_oAscPresentationShortcutType.EditUndo:
				return "EditUndo";
			case Asc.c_oAscPresentationShortcutType.EditRedo:
				return "EditRedo";
			case Asc.c_oAscPresentationShortcutType.Cut:
				return "Cut";
			case Asc.c_oAscPresentationShortcutType.Copy:
				return "Copy";
			case Asc.c_oAscPresentationShortcutType.Paste:
				return "Paste";
			case Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat:
				return "PasteTextWithoutFormat";
			case Asc.c_oAscPresentationShortcutType.CopyFormat:
				return "CopyFormat";
			case Asc.c_oAscPresentationShortcutType.PasteFormat:
				return "PasteFormat";
			case Asc.c_oAscPresentationShortcutType.UseDestinationTheme:
				return "UseDestinationTheme";
			case Asc.c_oAscPresentationShortcutType.KeepSourceFormat:
				return "KeepSourceFormat";
			case Asc.c_oAscPresentationShortcutType.PasteAsPicture:
				return "PasteAsPicture";
			case Asc.c_oAscPresentationShortcutType.KeepTextOnly:
				return "KeepTextOnly";
			case Asc.c_oAscPresentationShortcutType.InsertHyperlink:
				return "InsertHyperlink";
			case Asc.c_oAscPresentationShortcutType.VisitHyperlink:
				return "VisitHyperlink";
			case Asc.c_oAscPresentationShortcutType.EditSelectAll:
				return "EditSelectAll";
			case Asc.c_oAscPresentationShortcutType.SelectNextSlide:
				return "SelectNextSlide";
			case Asc.c_oAscPresentationShortcutType.SelectPreviousSlide:
				return "SelectPreviousSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToFirstSlide:
				return "SelectToFirstSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToLastSlide:
				return "SelectToLastSlide";
			case Asc.c_oAscPresentationShortcutType.SelectToStartLine:
				return "SelectToStartLine";
			case Asc.c_oAscPresentationShortcutType.SelectToEndLine:
				return "SelectToEndLine";
			case Asc.c_oAscPresentationShortcutType.SelectRightChar:
				return "SelectRightChar";
			case Asc.c_oAscPresentationShortcutType.SelectLeftChar:
				return "SelectLeftChar";
			case Asc.c_oAscPresentationShortcutType.SelectRightWord:
				return "SelectRightWord";
			case Asc.c_oAscPresentationShortcutType.SelectLeftWord:
				return "SelectLeftWord";
			case Asc.c_oAscPresentationShortcutType.SelectLineUp:
				return "SelectLineUp";
			case Asc.c_oAscPresentationShortcutType.SelectLineDown:
				return "SelectLineDown";
			case Asc.c_oAscPresentationShortcutType.EditDeselectAll:
				return "EditDeselectAll";
			case Asc.c_oAscPresentationShortcutType.ShowParaMarks:
				return "ShowParaMarks";
			case Asc.c_oAscPresentationShortcutType.Bold:
				return "Bold";
			case Asc.c_oAscPresentationShortcutType.Italic:
				return "Italic";
			case Asc.c_oAscPresentationShortcutType.Underline:
				return "Underline";
			case Asc.c_oAscPresentationShortcutType.Strikeout:
				return "Strikeout";
			case Asc.c_oAscPresentationShortcutType.Subscript:
				return "Subscript";
			case Asc.c_oAscPresentationShortcutType.Superscript:
				return "Superscript";
			case Asc.c_oAscPresentationShortcutType.BulletList:
				return "BulletList";
			case Asc.c_oAscPresentationShortcutType.ResetChar:
				return "ResetChar";
			case Asc.c_oAscPresentationShortcutType.IncreaseFontSize:
				return "IncreaseFontSize";
			case Asc.c_oAscPresentationShortcutType.DecreaseFontSize:
				return "DecreaseFontSize";
			case Asc.c_oAscPresentationShortcutType.CenterPara:
				return "CenterPara";
			case Asc.c_oAscPresentationShortcutType.JustifyPara:
				return "JustifyPara";
			case Asc.c_oAscPresentationShortcutType.RightPara:
				return "RightPara";
			case Asc.c_oAscPresentationShortcutType.LeftPara:
				return "LeftPara";
			case Asc.c_oAscPresentationShortcutType.Indent:
				return "Indent";
			case Asc.c_oAscPresentationShortcutType.UnIndent:
				return "UnIndent";
			case Asc.c_oAscPresentationShortcutType.DeleteLeftChar:
				return "DeleteLeftChar";
			case Asc.c_oAscPresentationShortcutType.DeleteRightChar:
				return "DeleteRightChar";
			case Asc.c_oAscPresentationShortcutType.DeleteLeftWord:
				return "DeleteLeftWord";
			case Asc.c_oAscPresentationShortcutType.DeleteRightWord:
				return "DeleteRightWord";
			case Asc.c_oAscPresentationShortcutType.StartIndent:
				return "StartIndent";
			case Asc.c_oAscPresentationShortcutType.StartUnIndent:
				return "StartUnIndent";
			case Asc.c_oAscPresentationShortcutType.InsertTab:
				return "InsertTab";
			case Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder:
				return "EquationAddPlaceholder";
			case Asc.c_oAscPresentationShortcutType.InsertLineBreak:
				return "InsertLineBreak";
			case Asc.c_oAscPresentationShortcutType.EndParagraph:
				return "EndParagraph";
			case Asc.c_oAscPresentationShortcutType.EuroSign:
				return "EuroSign";
			case Asc.c_oAscPresentationShortcutType.EnDash:
				return "EnDash";
			case Asc.c_oAscPresentationShortcutType.NonBreakingSpace:
				return "NonBreakingSpace";
			case Asc.c_oAscPresentationShortcutType.MoveToRightChar:
				return "MoveToRightChar";
			case Asc.c_oAscPresentationShortcutType.MoveToLeftChar:
				return "MoveToLeftChar";
			case Asc.c_oAscPresentationShortcutType.MoveToUpLine:
				return "MoveToUpLine";
			case Asc.c_oAscPresentationShortcutType.MoveToDownLine:
				return "MoveToDownLine";
			case Asc.c_oAscPresentationShortcutType.MoveToStartWord:
				return "MoveToStartWord";
			case Asc.c_oAscPresentationShortcutType.MoveToEndWord:
				return "MoveToEndWord";
			case Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder:
				return "GoToNextPlaceholder";
			case Asc.c_oAscPresentationShortcutType.MoveToStartLine:
				return "MoveToStartLine";
			case Asc.c_oAscPresentationShortcutType.MoveToEndLine:
				return "MoveToEndLine";
			case Asc.c_oAscPresentationShortcutType.MoveToStartContent:
				return "MoveToStartContent";
			case Asc.c_oAscPresentationShortcutType.MoveToEndContent:
				return "MoveToEndContent";
			case Asc.c_oAscPresentationShortcutType.SpeechWorker:
				return "SpeechWorker";
			default:
				return null;
		}
	}
	function getShortcutTypeFromString(str) {
		switch (str) {
			case "OpenFilePanel":
				return Asc.c_oAscPresentationShortcutType.OpenFilePanel;
			case "OpenFindDialog":
				return Asc.c_oAscPresentationShortcutType.OpenFindDialog;
			case "OpenFindAndReplaceMenu":
				return Asc.c_oAscPresentationShortcutType.OpenFindAndReplaceMenu;
			case "OpenCommentsPanel":
				return Asc.c_oAscPresentationShortcutType.OpenCommentsPanel;
			case "OpenCommentField":
				return Asc.c_oAscPresentationShortcutType.OpenCommentField;
			case "OpenChatPanel":
				return Asc.c_oAscPresentationShortcutType.OpenChatPanel;
			case "Save":
				return Asc.c_oAscPresentationShortcutType.Save;
			case "PrintPreviewAndPrint":
				return Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint;
			case "SaveAs":
				return Asc.c_oAscPresentationShortcutType.SaveAs;
			case "OpenHelpMenu":
				return Asc.c_oAscPresentationShortcutType.OpenHelpMenu;
			case "OpenExistingFile":
				return Asc.c_oAscPresentationShortcutType.OpenExistingFile;
			case "NextFileTab":
				return Asc.c_oAscPresentationShortcutType.NextFileTab;
			case "PreviousFileTab":
				return Asc.c_oAscPresentationShortcutType.PreviousFileTab;
			case "CloseFile":
				return Asc.c_oAscPresentationShortcutType.CloseFile;
			case "OpenContextMenu":
				return Asc.c_oAscPresentationShortcutType.OpenContextMenu;
			case "CloseMenu":
				return Asc.c_oAscPresentationShortcutType.CloseMenu;
			case "Zoom100":
				return Asc.c_oAscPresentationShortcutType.Zoom100;
			case "GoToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.GoToFirstSlide;
			case "GoToLastSlide":
				return Asc.c_oAscPresentationShortcutType.GoToLastSlide;
			case "GoToNextSlide":
				return Asc.c_oAscPresentationShortcutType.GoToNextSlide;
			case "GoToPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.GoToPreviousSlide;
			case "ZoomIn":
				return Asc.c_oAscPresentationShortcutType.ZoomIn;
			case "ZoomOut":
				return Asc.c_oAscPresentationShortcutType.ZoomOut;
			case "NextModalControl":
				return Asc.c_oAscPresentationShortcutType.NextModalControl;
			case "PreviousModalControl":
				return Asc.c_oAscPresentationShortcutType.PreviousModalControl;
			case "NewSlide":
				return Asc.c_oAscPresentationShortcutType.NewSlide;
			case "RemoveSlide":
				return Asc.c_oAscPresentationShortcutType.RemoveSlide;
			case "Duplicate":
				return Asc.c_oAscPresentationShortcutType.Duplicate;
			case "MoveSlideUp":
				return Asc.c_oAscPresentationShortcutType.MoveSlideUp;
			case "MoveSlideDown":
				return Asc.c_oAscPresentationShortcutType.MoveSlideDown;
			case "MoveSlideToBegin":
				return Asc.c_oAscPresentationShortcutType.MoveSlideToBegin;
			case "MoveSlideToEnd":
				return Asc.c_oAscPresentationShortcutType.MoveSlideToEnd;
			case "EditShape":
				return Asc.c_oAscPresentationShortcutType.EditShape;
			case "EditChart":
				return Asc.c_oAscPresentationShortcutType.EditChart;
			case "Group":
				return Asc.c_oAscPresentationShortcutType.Group;
			case "UnGroup":
				return Asc.c_oAscPresentationShortcutType.UnGroup;
			case "MoveFocusToNextObject":
				return Asc.c_oAscPresentationShortcutType.MoveFocusToNextObject;
			case "MoveFocusToPreviousObject":
				return Asc.c_oAscPresentationShortcutType.MoveFocusToPreviousObject;
			case "LittleMoveObjectLeft":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectLeft;
			case "LittleMoveObjectRight":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectRight;
			case "LittleMoveObjectUp":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectUp;
			case "LittleMoveObjectDown":
				return Asc.c_oAscPresentationShortcutType.LittleMoveObjectDown;
			case "BigMoveObjectLeft":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectLeft;
			case "BigMoveObjectRight":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectRight;
			case "BigMoveObjectUp":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectUp;
			case "BigMoveObjectDown":
				return Asc.c_oAscPresentationShortcutType.BigMoveObjectDown;
			case "MoveToNextCell":
				return Asc.c_oAscPresentationShortcutType.MoveToNextCell;
			case "MoveToPreviousCell":
				return Asc.c_oAscPresentationShortcutType.MoveToPreviousCell;
			case "MoveToNextRow":
				return Asc.c_oAscPresentationShortcutType.MoveToNextRow;
			case "MoveToPreviousRow":
				return Asc.c_oAscPresentationShortcutType.MoveToPreviousRow;
			case "EndParagraphCell":
				return Asc.c_oAscPresentationShortcutType.EndParagraphCell;
			case "AddNewRow":
				return Asc.c_oAscPresentationShortcutType.AddNewRow;
			case "DemonstrationStartPresentation":
				return Asc.c_oAscPresentationShortcutType.DemonstrationStartPresentation;
			case "DemonstrationGoToNextSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToNextSlide;
			case "DemonstrationGoToPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToPreviousSlide;
			case "DemonstrationGoToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToFirstSlide;
			case "DemonstrationGoToLastSlide":
				return Asc.c_oAscPresentationShortcutType.DemonstrationGoToLastSlide;
			case "DemonstrationClosePreview":
				return Asc.c_oAscPresentationShortcutType.DemonstrationClosePreview;
			case "EditUndo":
				return Asc.c_oAscPresentationShortcutType.EditUndo;
			case "EditRedo":
				return Asc.c_oAscPresentationShortcutType.EditRedo;
			case "Cut":
				return Asc.c_oAscPresentationShortcutType.Cut;
			case "Copy":
				return Asc.c_oAscPresentationShortcutType.Copy;
			case "Paste":
				return Asc.c_oAscPresentationShortcutType.Paste;
			case "PasteTextWithoutFormat":
				return Asc.c_oAscPresentationShortcutType.PasteTextWithoutFormat;
			case "CopyFormat":
				return Asc.c_oAscPresentationShortcutType.CopyFormat;
			case "PasteFormat":
				return Asc.c_oAscPresentationShortcutType.PasteFormat;
			case "UseDestinationTheme":
				return Asc.c_oAscPresentationShortcutType.UseDestinationTheme;
			case "KeepSourceFormat":
				return Asc.c_oAscPresentationShortcutType.KeepSourceFormat;
			case "PasteAsPicture":
				return Asc.c_oAscPresentationShortcutType.PasteAsPicture;
			case "KeepTextOnly":
				return Asc.c_oAscPresentationShortcutType.KeepTextOnly;
			case "InsertHyperlink":
				return Asc.c_oAscPresentationShortcutType.InsertHyperlink;
			case "VisitHyperlink":
				return Asc.c_oAscPresentationShortcutType.VisitHyperlink;
			case "EditSelectAll":
				return Asc.c_oAscPresentationShortcutType.EditSelectAll;
			case "SelectNextSlide":
				return Asc.c_oAscPresentationShortcutType.SelectNextSlide;
			case "SelectPreviousSlide":
				return Asc.c_oAscPresentationShortcutType.SelectPreviousSlide;
			case "SelectToFirstSlide":
				return Asc.c_oAscPresentationShortcutType.SelectToFirstSlide;
			case "SelectToLastSlide":
				return Asc.c_oAscPresentationShortcutType.SelectToLastSlide;
			case "SelectToStartLine":
				return Asc.c_oAscPresentationShortcutType.SelectToStartLine;
			case "SelectToEndLine":
				return Asc.c_oAscPresentationShortcutType.SelectToEndLine;
			case "SelectRightChar":
				return Asc.c_oAscPresentationShortcutType.SelectRightChar;
			case "SelectLeftChar":
				return Asc.c_oAscPresentationShortcutType.SelectLeftChar;
			case "SelectRightWord":
				return Asc.c_oAscPresentationShortcutType.SelectRightWord;
			case "SelectLeftWord":
				return Asc.c_oAscPresentationShortcutType.SelectLeftWord;
			case "SelectLineUp":
				return Asc.c_oAscPresentationShortcutType.SelectLineUp;
			case "SelectLineDown":
				return Asc.c_oAscPresentationShortcutType.SelectLineDown;
			case "EditDeselectAll":
				return Asc.c_oAscPresentationShortcutType.EditDeselectAll;
			case "ShowParaMarks":
				return Asc.c_oAscPresentationShortcutType.ShowParaMarks;
			case "Bold":
				return Asc.c_oAscPresentationShortcutType.Bold;
			case "Italic":
				return Asc.c_oAscPresentationShortcutType.Italic;
			case "Underline":
				return Asc.c_oAscPresentationShortcutType.Underline;
			case "Strikeout":
				return Asc.c_oAscPresentationShortcutType.Strikeout;
			case "Subscript":
				return Asc.c_oAscPresentationShortcutType.Subscript;
			case "Superscript":
				return Asc.c_oAscPresentationShortcutType.Superscript;
			case "BulletList":
				return Asc.c_oAscPresentationShortcutType.BulletList;
			case "ResetChar":
				return Asc.c_oAscPresentationShortcutType.ResetChar;
			case "IncreaseFontSize":
				return Asc.c_oAscPresentationShortcutType.IncreaseFontSize;
			case "DecreaseFontSize":
				return Asc.c_oAscPresentationShortcutType.DecreaseFontSize;
			case "CenterPara":
				return Asc.c_oAscPresentationShortcutType.CenterPara;
			case "JustifyPara":
				return Asc.c_oAscPresentationShortcutType.JustifyPara;
			case "RightPara":
				return Asc.c_oAscPresentationShortcutType.RightPara;
			case "LeftPara":
				return Asc.c_oAscPresentationShortcutType.LeftPara;
			case "Indent":
				return Asc.c_oAscPresentationShortcutType.Indent;
			case "UnIndent":
				return Asc.c_oAscPresentationShortcutType.UnIndent;
			case "DeleteLeftChar":
				return Asc.c_oAscPresentationShortcutType.DeleteLeftChar;
			case "DeleteRightChar":
				return Asc.c_oAscPresentationShortcutType.DeleteRightChar;
			case "DeleteLeftWord":
				return Asc.c_oAscPresentationShortcutType.DeleteLeftWord;
			case "DeleteRightWord":
				return Asc.c_oAscPresentationShortcutType.DeleteRightWord;
			case "StartIndent":
				return Asc.c_oAscPresentationShortcutType.StartIndent;
			case "StartUnIndent":
				return Asc.c_oAscPresentationShortcutType.StartUnIndent;
			case "InsertTab":
				return Asc.c_oAscPresentationShortcutType.InsertTab;
			case "EquationAddPlaceholder":
				return Asc.c_oAscPresentationShortcutType.EquationAddPlaceholder;
			case "InsertLineBreak":
				return Asc.c_oAscPresentationShortcutType.InsertLineBreak;
			case "EndParagraph":
				return Asc.c_oAscPresentationShortcutType.EndParagraph;
			case "EuroSign":
				return Asc.c_oAscPresentationShortcutType.EuroSign;
			case "EnDash":
				return Asc.c_oAscPresentationShortcutType.EnDash;
			case "NonBreakingSpace":
				return Asc.c_oAscPresentationShortcutType.NonBreakingSpace;
			case "MoveToRightChar":
				return Asc.c_oAscPresentationShortcutType.MoveToRightChar;
			case "MoveToLeftChar":
				return Asc.c_oAscPresentationShortcutType.MoveToLeftChar;
			case "MoveToUpLine":
				return Asc.c_oAscPresentationShortcutType.MoveToUpLine;
			case "MoveToDownLine":
				return Asc.c_oAscPresentationShortcutType.MoveToDownLine;
			case "MoveToStartWord":
				return Asc.c_oAscPresentationShortcutType.MoveToStartWord;
			case "MoveToEndWord":
				return Asc.c_oAscPresentationShortcutType.MoveToEndWord;
			case "GoToNextPlaceholder":
				return Asc.c_oAscPresentationShortcutType.GoToNextPlaceholder;
			case "MoveToStartLine":
				return Asc.c_oAscPresentationShortcutType.MoveToStartLine;
			case "MoveToEndLine":
				return Asc.c_oAscPresentationShortcutType.MoveToEndLine;
			case "MoveToStartContent":
				return Asc.c_oAscPresentationShortcutType.MoveToStartContent;
			case "MoveToEndContent":
				return Asc.c_oAscPresentationShortcutType.MoveToEndContent;
			case "SpeechWorker":
				return Asc.c_oAscPresentationShortcutType.SpeechWorker;
			default:
				return null;
		}
	}

	function isTypeEvent(nType, oKeyboardEvent) {
		const arrShortcuts = c_oAscDefaultShortcuts[nType];
		if (arrShortcuts) {
			const nKeyboardIndex = AscCommon.CShortcuts.GetShortcutIndex(oKeyboardEvent.GetKeyCode(), oKeyboardEvent.IsShortcutCtrl(), oKeyboardEvent.IsShift(), oKeyboardEvent.IsAlt(), oKeyboardEvent.IsCmd());
			for (let i = 0; i < arrShortcuts.length; i += 1) {
				const nShortcutIndex = arrShortcuts[i].asc_GetShortcutIndex();
				if (nShortcutIndex === nKeyboardIndex) {
					return true;
				}
			}
		}
		return false;
	}
	function isCopyPasteEvent(oKeyboardEvent) {
		return isTypeEvent(Asc.c_oAscPresentationShortcutType.Cut, oKeyboardEvent) ||
			isTypeEvent(Asc.c_oAscPresentationShortcutType.Copy, oKeyboardEvent) ||
			isTypeEvent(Asc.c_oAscPresentationShortcutType.Paste, oKeyboardEvent);
	}
	window["Asc"]["c_oAscDefaultShortcuts"] = window["Asc"].c_oAscDefaultShortcuts = c_oAscDefaultShortcuts;
	window["Asc"]["c_oAscUnlockedShortcutActionTypes"] = window["Asc"].c_oAscUnlockedShortcutActionTypes = c_oAscUnlockedShortcutActionTypes;
	window["AscCommon"].getStringFromShortcutType = getStringFromShortcutType;
	window["AscCommon"].getShortcutTypeFromString = getShortcutTypeFromString;
	window["AscCommon"].isCopyPasteEvent = isCopyPasteEvent;
})();
