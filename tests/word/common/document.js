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

(function(window)
{
	let logicDocument = null;
	let styleManager  = null;
	let styleCounter  = 0;

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
		
		styleManager = logicDocument.GetStyleManager();

		return logicDocument;
	}
	function CreateParagraph()
	{
		return new AscWord.Paragraph();
	}
	function CreateRun()
	{
		return new AscWord.CRun();
	}
	function CreateTable(rows, cols, tableGrid)
	{
		let t = new AscWord.CTable(AscTest.DrawingDocument, null, true, rows, cols, tableGrid);
		t.CorrectBadTable();
		return t;
	}
	function CreateImage(w, h)
	{
		if (!logicDocument)
			return null;
		
		let drawingObjects = logicDocument.GetDrawingObjects();
		let drawing = new ParaDrawing(w, h, null, logicDocument.DrawingDocument, logicDocument, null);
		let image   = drawingObjects.createImage(AscCommon.g_sWordPlaceholderImage, 0, 0, w, h);
		image.setParent(drawing);
		drawing.Set_GraphicObject(image);
		return drawing;
	}
	function CreateStyle(styleType, name)
	{
		if (!styleManager)
			return null;
		
		if (!name)
			name = "style" + (++styleCounter);
		
		let style = new AscWord.CStyle(name, null, null, styleType);
		styleManager.Add(style);
		return style;
	}
	function CreateParagraphStyle(name)
	{
		return CreateStyle(styletype_Paragraph, name);
	}
	function CreateRunStyle(name)
	{
		return CreateStyle(styletype_Character, name);
	}
	function CreateMath()
	{
		let math = new AscWord.ParaMath();
		let run = new AscWord.CRun(null, true);
		math.Root.AddToContent(0, run);
		return math;
	}
	function CreateBlockLevelSdt()
	{
		return new AscWord.CBlockLevelSdt()
	}
	function CreateDefaultHeader(sectPr)
	{
		if (!sectPr || !logicDocument)
			return null;
		
		let header = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Header);
		sectPr.Set_Header_Default(header);
		return header.GetContent();
	}
	function GetParagraphText(paragraph)
	{
		return paragraph.GetText({
			ParaSeparator : "",
			TableCellSeparator : "",
			TableRowSeparator : "",
		});
	}
	function GetParagraphReviewText(paragraph)
	{
		let result = [];
		paragraph.CheckRunContent(function(run)
		{
			let text = run.GetText({
				ParaSeparator : "",
				Text : ""
			});
			if (!text || !text.length)
				return;
			
			let reviewType = run.GetReviewType();
			if (result.length && reviewType === result[result.length - 1][0])
				result[result.length - 1][1] += text;
			else
				result.push([reviewType, text]);
		});
		return result;
	}
	function GetBlockLevelSdtText(cc)
	{
		let result = "";
		cc.CheckRunContent(function(run)
		{
			result += run.GetText();
		});
		return result;
	}
	function GetBlockLevelSdtReviewText(cc)
	{
		let result = [];
		cc.CheckRunContent(function(run)
		{
			let text = run.GetText({
				ParaSeparator : "",
				Text : ""
			});
			if (!text || !text.length)
				return;
			
			let reviewType = run.GetReviewType();
			if (result.length && reviewType === result[result.length - 1][0])
				result[result.length - 1][1] += text;
			else
				result.push([reviewType, text]);
		});
		return result;
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
	function RemoveTableMargins(table)
	{
		table.Set_TableCellMar(0, 0, 0, 0);
		
		for (let rowIndex = 0, rowCount = table.GetRowsCount(); rowIndex < rowIndex; ++rowIndex)
		{
			let row = table.GetRow(rowIndex);
			for (let cellIndex = 0, cellCount = table.GetCells; cellIndex < cellCount; ++cellIndex)
			{
				row.GetCell(cellIndex).Set_Margins(null);
			}
		}
	}
	function SetFillingFormMode(isOForm)
	{
		editor.restrictions = Asc.c_oAscRestrictionType.OnlyForms;

		if (isOForm)
			editor.DocInfo = {Format : "oform", isFormatWithForms : function() {return true;}};
		else
			editor.DocInfo = {Format : "docx", isFormatWithForms : function() {return false;}};
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
	function AcceptAllRevisionChanges()
	{
		if (!logicDocument)
			return;
		
		logicDocument.AcceptAllRevisionChanges();
	}
	function RejectAllRevisionChanges()
	{
		if (!logicDocument)
			return;
		
		logicDocument.RejectAllRevisionChanges();
	}
	function AcceptRevisionChangesBySelection()
	{
		if (!logicDocument)
			return;
		
		logicDocument.AcceptRevisionChangesBySelection();
	}
	function RejectRevisionChangesBySelection()
	{
		if (!logicDocument)
			return;
		
		logicDocument.RejectRevisionChangesBySelection();
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
	function AddTextToParagraph(p, text)
	{
		let run = CreateRun();
		p.AddToContentToEnd(run);
		run.AddText(text);
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

		logicDocument.getCompositeInput().begin();
	}
	function ReplaceCompositeInput(text)
	{
		if (!logicDocument)
			return;

		logicDocument.getCompositeInput().replace(text);
	}
	function EndCompositeInput()
	{
		if (!logicDocument)
			return;

		logicDocument.getCompositeInput().end();
	}
	function EnterTextCompositeInput(text)
	{
		BeginCompositeInput();
		ReplaceCompositeInput(text);
		EndCompositeInput();
	}
	function MoveCursorToParagraph(paragraph, isToStart)
	{
		if (!paragraph || !(paragraph instanceof AscWord.Paragraph))
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
	function SelectParagraphRange(paragraph, start, end)
	{
		if (!paragraph || start >= end)
			return;
		
		if (logicDocument)
			logicDocument.RemoveSelection();
		
		paragraph.RemoveSelection();
		paragraph.MoveCursorToStartPos();
		for (let i = 0; i < start; ++i)
			paragraph.MoveCursorRight(false, false);
		
		let startPos = paragraph.getCurrentPos();
		for (let i = start; i < end; ++i)
			paragraph.MoveCursorRight(false, false);
		
		let endPos = paragraph.getCurrentPos();
		paragraph.StartSelectionFromCurPos();
		paragraph.SetSelectionContentPos(startPos, endPos, false);
		paragraph.Document_SetThisElementCurrent();
	}
	function SelectParagraph(paragraph)
	{
		if (logicDocument)
			logicDocument.RemoveSelection();
		
		paragraph.SelectAll();
		paragraph.Document_SetThisElementCurrent();
	}
	function GetFinalSection()
	{
		if (!logicDocument)
			return null;
		
		return logicDocument.SectPr;
	}
	function SetCompatibilityMode(mode)
	{
		if (!logicDocument)
			return;
		
		logicDocument.Settings.CompatibilityMode = mode;
	}
	function StartCollaboration()
	{
		AscCommon.CollaborativeEditing.Start_CollaborationEditing();
		if (logicDocument)
			logicDocument.StartCollaborationEditing();
		
		SyncCollaboration();
	}
	function SyncCollaboration()
	{
		AscCommon.CollaborativeEditing.Send_Changes();
		
	}
	function EndCollaboration()
	{
		AscCommon.CollaborativeEditing.End_CollaborationEditing();
	}
	function StartTextSpeaker()
	{
		AscCommon.EditorActionSpeaker.run();
	}
	function StopTextSpeaker()
	{
		AscCommon.EditorActionSpeaker.stop();
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.CreateLogicDocument              = CreateLogicDocument;
	AscTest.GetLogicDocument                 = CreateLogicDocument;
	AscTest.CreateParagraph                  = CreateParagraph;
	AscTest.CreateRun                        = CreateRun;
	AscTest.CreateTable                      = CreateTable;
	AscTest.CreateImage                      = CreateImage;
	AscTest.CreateStyle                      = CreateStyle;
	AscTest.CreateParagraphStyle             = CreateParagraphStyle;
	AscTest.CreateRunStyle                   = CreateRunStyle;
	AscTest.CreateMath                       = CreateMath;
	AscTest.CreateBlockLvlSdt                = CreateBlockLevelSdt;
	AscTest.CreateDefaultHeader              = CreateDefaultHeader;
	AscTest.GetParagraphText                 = GetParagraphText;
	AscTest.GetParagraphReviewText           = GetParagraphReviewText;
	AscTest.GetBlockLevelSdtText             = GetBlockLevelSdtText;
	AscTest.GetBlockLevelSdtReviewText       = GetBlockLevelSdtReviewText;
	AscTest.RemoveTableBorders               = RemoveTableBorders;
	AscTest.RemoveTableMargins               = RemoveTableMargins;
	AscTest.SetFillingFormMode               = SetFillingFormMode;
	AscTest.SetEditingMode                   = SetEditingMode;
	AscTest.SetTrackRevisions                = SetTrackRevisions;
	AscTest.AcceptAllRevisionChanges         = AcceptAllRevisionChanges;
	AscTest.RejectAllRevisionChanges         = RejectAllRevisionChanges;
	AscTest.AcceptRevisionChangesBySelection = AcceptRevisionChangesBySelection;
	AscTest.RejectRevisionChangesBySelection = RejectRevisionChangesBySelection;
	AscTest.PressKey                         = PressKey;
	AscTest.MoveCursorLeft                   = MoveCursorLeft;
	AscTest.MoveCursorRight                  = MoveCursorRight;
	AscTest.Recalculate                      = Recalculate;
	AscTest.ClickMouseButton                 = ClickMouseButton;
	AscTest.ClearDocument                    = ClearDocument;
	AscTest.ClearParagraph                   = ClearParagraph;
	AscTest.AddTextToParagraph               = AddTextToParagraph;
	AscTest.EnterText                        = EnterText;
	AscTest.CorrectEnterText                 = CorrectEnterText;
	AscTest.BeginCompositeInput              = BeginCompositeInput;
	AscTest.ReplaceCompositeInput            = ReplaceCompositeInput;
	AscTest.EndCompositeInput                = EndCompositeInput;
	AscTest.EnterTextCompositeInput          = EnterTextCompositeInput;
	AscTest.Key                              = Key;
	AscTest.MoveCursorToParagraph            = MoveCursorToParagraph;
	AscTest.AddNumbering                     = AddNumbering;
	AscTest.SetParagraphNumberingLvl         = SetParagraphNumberingLvl;
	AscTest.SelectDocumentRange              = SelectDocumentRange;
	AscTest.GetFinalSection                  = GetFinalSection;
	AscTest.SetCompatibilityMode             = SetCompatibilityMode;
	AscTest.StartCollaboration               = StartCollaboration;
	AscTest.SyncCollaboration                = SyncCollaboration;
	AscTest.EndCollaboration                 = EndCollaboration;
	AscTest.SelectParagraphRange             = SelectParagraphRange;
	AscTest.SelectParagraph                  = SelectParagraph;
	AscTest.StartTextSpeaker                 = StartTextSpeaker;
	AscTest.StopTextSpeaker                  = StopTextSpeaker;

})(window);

