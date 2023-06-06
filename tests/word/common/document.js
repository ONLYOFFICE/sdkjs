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

(function(window)
{
	let logicDocument = null;

	const Key = {
		_0 : 48,
		_1 : 49,
		_2 : 50,
		_3 : 51,
		_4 : 52,
		_5 : 53,
		_6 : 54,
		_7 : 55,
		_8 : 56,
		_9 : 57,

		A : 65,
		B : 66,
		C : 67,
		D : 68,
		E : 69,
		F : 70,

		a : 97,
		b : 98,
		c : 99,
		d : 100,
		e : 101,
		f : 102,

		space : 32,
		backspace : 8,
		minus : 45,
		delete : 46,
		enter : 13
	};

	function IsKeyDown(key)
	{
		return (Key.space === key
			|| Key.backspace === key
			|| Key.enter === key
			|| Key.delete === key);
	}

	function CreateLogicDocument()
	{
		if (logicDocument)
			return logicDocument;

		logicDocument = new AscWord.CDocument(AscTest.DrawingDocument, true);
		logicDocument.Api = AscTest.Editor;
		logicDocument.On_EndLoad();

		AscTest.DrawingDocument.m_oLogicDocument = logicDocument;

		return logicDocument;
	}
	function CreateParagraph()
	{
		return new AscWord.CParagraph(AscTest.DrawingDocument);
	}
	function CreateTable(rows, cols)
	{
		return new AscWord.CTable(AscTest.DrawingDocument, null, true, rows, cols);
	}
	function GetParagraphText(paragraph)
	{
		return paragraph.GetText({ParaEndToSpace : false});
	}
	function RemoveTableBorders(table)
	{
		function CreateNoneBorder()
		{
			let border = new AscWord.CBorder();
			border.SetNone();
			return border;
		}

		table.Set_TableBorder_Left(CreateNoneBorder());
		table.Set_TableBorder_Top(CreateNoneBorder());
		table.Set_TableBorder_Right(CreateNoneBorder());
		table.Set_TableBorder_Bottom(CreateNoneBorder());
		table.Set_TableBorder_InsideH(CreateNoneBorder());
		table.Set_TableBorder_InsideV(CreateNoneBorder());
	}
	function SetFillingFormMode(isOForm)
	{
		editor.restrictions = Asc.c_oAscRestrictionType.OnlyForms;

		if (isOForm)
			editor.DocInfo = {Format : "oform"};
		else
			editor.DocInfo = {Format : "docx"};
	}
	function SetEditingMode()
	{
		editor.restrictions = Asc.c_oAscRestrictionType.None;

		editor.DocInfo = {Format : "docx"};
	}
	function SetTrackRevisions(turnOn)
	{
		if (!logicDocument)
			return;
		
		logicDocument.SetLocalTrackRevisions(turnOn);
	}
	function PressKey(keyCode, isCtrl, isShift, isAlt)
	{
		if (!logicDocument)
			return;

		global_mouseEvent.CtrlKey  = !!isCtrl;
		global_mouseEvent.ShiftKey = !!isShift;
		global_mouseEvent.AltKey   = !!isAlt;
		global_mouseEvent.KeyCode  = keyCode;

		if (IsKeyDown(keyCode))
			logicDocument.OnKeyDown(global_mouseEvent);
		else
			logicDocument.OnKeyPress(global_mouseEvent);
	}
	function MoveCursorLeft(isShift, isCtrl, count)
	{
		if (!logicDocument)
			return;
		
		if (!count)
			count = 1;
		
		while (count)
		{
			logicDocument.MoveCursorLeft(!!isShift, !!isCtrl);
			--count;
		}
	}
	function MoveCursorRight(isShift, isCtrl, count)
	{
		if (!logicDocument)
			return;

		if (!count)
			count = 1;

		while (count)
		{
			logicDocument.MoveCursorRight(!!isShift, !!isCtrl, false);
			--count;
		}
	}
	function ClickMouseButton(x, y, page, isRight, count)
	{
		if (!logicDocument)
			return;

		let e = new AscCommon.CMouseEventHandler();

		e.Button     = isRight ? AscCommon.g_mouse_button_right : AscCommon.g_mouse_button_left;
		e.ClickCount = count ? count : 1;

		e.Type = AscCommon.g_mouse_event_type_down;
		logicDocument.OnMouseDown(e, x, y, page);

		e.Type = AscCommon.g_mouse_event_type_up;
		logicDocument.OnMouseUp(e, x, y, page);
	}
	function Recalculate()
	{
		if (!logicDocument)
			return;

		logicDocument.RecalculateFromStart(false);
	}
	function ClearDocument()
	{
		if (!logicDocument)
			return;

		logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	}
	function ClearParagraph(p)
	{
		if (!p)
			return;
		
		p.RemoveFromContent(0, p.GetElementsCount());
	}
	function EnterText(text)
	{
		if (!logicDocument)
			return;

		logicDocument.EnterText(text);
	}
	function CorrectEnterText(oldText, newText)
	{
		if (!logicDocument)
			return;

		logicDocument.CorrectEnterText(oldText, newText);
	}
	function BeginCompositeInput()
	{
		if (!logicDocument)
			return;

		logicDocument.Begin_CompositeInput();
	}
	function ReplaceCompositeInput(text)
	{
		if (!logicDocument)
			return;

		logicDocument.Replace_CompositeText(text);
	}
	function EndCompositeInput()
	{
		if (!logicDocument)
			return;

		logicDocument.End_CompositeInput();
	}
	function EnterTextCompositeInput(text)
	{
		BeginCompositeInput();
		ReplaceCompositeInput(text);
		EndCompositeInput();
	}
	function MoveCursorToParagraph(paragraph, isToStart)
	{
		if (!paragraph || !(paragraph instanceof AscWord.CParagraph))
			return;
		
		paragraph.SetThisElementCurrent();
		
		if (false === isToStart)
			paragraph.MoveCursorToEndPos();
		else
			paragraph.MoveCursorToStartPos();
	}
	function AddNumbering(type, subtype)
	{
		if (!logicDocument)
			return;
		
		let numObject = AscWord.GetNumberingObjectByDeprecatedTypes(type, subtype);
		if (!numObject)
			return;
		
		let _numInfo = numObject;
		if (typeof _numInfo === "string" || _numInfo instanceof String)
		{
			try
			{
				_numInfo = JSON.parse(numInfo);
			}
			catch (e)
			{
				return;
			}
		}
		
		logicDocument.SetParagraphNumbering(_numInfo);
	}
	function SetParagraphNumberingLvl(paragraph, iLvl)
	{
		let numPr = paragraph.GetNumPr();
		if (!numPr)
			return;
		
		paragraph.SetNumPr(numPr.NumId, iLvl)
	}
	function SelectDocumentRange(startPos, endPos)
	{
		if (!logicDocument)
			return;
		
		logicDocument.RemoveSelection();
		logicDocument.Selection.Use = true;
		logicDocument.Selection.StartPos = startPos;
		logicDocument.Selection.EndPos   = endPos;

		let direction = startPos < endPos ? 1 : -1;
		for (let pos = startPos; pos <= endPos; ++pos)
		{
			logicDocument.Content[pos].SelectAll(direction);
		}
	}
	function GetFinalSection()
	{
		if (!logicDocument)
			return null;
		
		return logicDocument.SectPr;
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.CreateLogicDocument      = CreateLogicDocument;
	AscTest.CreateParagraph          = CreateParagraph;
	AscTest.CreateTable              = CreateTable;
	AscTest.GetParagraphText         = GetParagraphText;
	AscTest.RemoveTableBorders       = RemoveTableBorders;
	AscTest.SetFillingFormMode       = SetFillingFormMode;
	AscTest.SetEditingMode           = SetEditingMode;
	AscTest.SetTrackRevisions        = SetTrackRevisions;
	AscTest.PressKey                 = PressKey;
	AscTest.MoveCursorLeft           = MoveCursorLeft;
	AscTest.MoveCursorRight          = MoveCursorRight;
	AscTest.Recalculate              = Recalculate;
	AscTest.ClickMouseButton         = ClickMouseButton;
	AscTest.ClearDocument            = ClearDocument;
	AscTest.ClearParagraph           = ClearParagraph;
	AscTest.EnterText                = EnterText;
	AscTest.CorrectEnterText         = CorrectEnterText;
	AscTest.BeginCompositeInput      = BeginCompositeInput;
	AscTest.ReplaceCompositeInput    = ReplaceCompositeInput;
	AscTest.EndCompositeInput        = EndCompositeInput;
	AscTest.EnterTextCompositeInput  = EnterTextCompositeInput;
	AscTest.Key                      = Key;
	AscTest.MoveCursorToParagraph    = MoveCursorToParagraph;
	AscTest.AddNumbering             = AddNumbering;
	AscTest.SetParagraphNumberingLvl = SetParagraphNumberingLvl;
	AscTest.SelectDocumentRange      = SelectDocumentRange;
	AscTest.GetFinalSection          = GetFinalSection;

})(window);

