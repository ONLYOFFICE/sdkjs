/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window)
{
	let logicDocument = null;

	const KeyCode = {
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
		f : 102
	};

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
	function PressKey(keyCode, isCtrl, isShift, isAlt)
	{
		if (!logicDocument)
			return;

		global_mouseEvent.CtrlKey  = !!isCtrl;
		global_mouseEvent.ShiftKey = !!isShift;
		global_mouseEvent.AltKey   = !!isAlt;
		global_mouseEvent.KeyCode  = keyCode;

		if (0x20 !== keyCode)
			logicDocument.OnKeyPress(global_mouseEvent);
		else
			logicDocument.OnKeyDown(global_mouseEvent);
	}
	function MoveCursorLeft(isShift, isCtrl)
	{
		if (!logicDocument)
			return;

		logicDocument.MoveCursorLeft(!!isShift, !!isCtrl);
	}
	function MoveCursorRight(isShift, isCtrl)
	{
		if (!logicDocument)
			return;

		logicDocument.MoveCursorRight(!!isShift, !!isCtrl, false);
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
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.CreateLogicDocument = CreateLogicDocument;
	AscTest.SetFillingFormMode  = SetFillingFormMode;
	AscTest.SetEditingMode      = SetEditingMode;
	AscTest.PressKey            = PressKey;
	AscTest.MoveCursorLeft      = MoveCursorLeft;
	AscTest.MoveCursorRight     = MoveCursorRight;
	AscTest.Recalculate         = Recalculate;
	AscTest.ClickMouseButton    = ClickMouseButton;
	AscTest.KeyCode             = KeyCode;

})(window);


