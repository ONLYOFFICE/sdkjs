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

const test_CharWidth   = 0.5;
const test_FontSize    = 10;
const test_FontHeight  = 20;
const test_FontAscent  = 15;
const test_FontDescent = 5;

// TODO: Сейчас fontSize=10, потому что по умолчанию такой в редакторе, нужно, чтобы здесь он выставлялся

window.AscNotLoadAllScript = true;

const test_GraphemeNormal       = 20;
const test_GraphemeLigature_ffi = 25;
const test_GraphemeLigature_ff  = 30;
const test_GraphemeLigature_fi  = 35;

const test_GraphemeCombining_xyz = 40;
const test_GraphemeCombining_xy  = 45;

const test_LETTER = {
	f : 102,
	i : 105,

	x : 120,
	y : 121,
	z : 122
};

var drawingDocument = {
	OnStartRecalculate : function(){},
	OnRecalculatePage : function(){},
	OnEndRecalculate : function(){},
	UpdateTargetTransform : function(){},
	SelectEnabled : function(){},
	SelectShow : function(){},
	TargetStart : function(){},
	TargetShow : function(){},
	TargetEnd : function(){},
	Set_RulerState_Start : function(){},
	Set_RulerState_Paragraph : function(){},
	Set_RulerState_End : function(){},
	Update_MathTrack : function(){},
	OnDrawContentControl : function(){},
	Update_FieldTrack : function(){},
	SetTargetColor : function(){},
	SetTargetSize : function(){},
	UpdateTarget : function(){}
};

var editor = {
	WordControl : drawingDocument
};

var XRegExp = XRegExp | new function(){};

if (!AscCommon.g_oIdCounter)
{
	AscCommon.g_oIdCounter = {
		Counter : 0,
		Get_NewId : function()
		{
			return ++this.Counter;
		}
	};
}


AscCommon.g_oTableId = {
	Add : function(c, id) {},

	TurnOff : function(){},
	TurnOn : function(){}
};
AscCommon.g_oIdCounter.m_bLoad = false;
AscCommon.g_oIdCounter.m_bRead = false;

let HB_String = [];
AscFonts.HB_StartString = function()
{
	HB_String.length = 0;
};
AscFonts.HB_EndString = function()
{

};
AscFonts.HB_AppendToString = function(u)
{
	HB_String.push(u);
};
AscFonts.CTextShaper.prototype.FlushWord = function()
{
	AscFonts.HB_EndString();

	for (let nIndex = 0, nCount = HB_String.length; nIndex < nCount; ++nIndex)
	{
		if (nCount - nIndex >= 3
			&& test_LETTER.f === HB_String[nIndex]
			&& test_LETTER.f === HB_String[nIndex + 1]
			&& test_LETTER.i === HB_String[nIndex + 2])
		{
			this.FlushGrapheme(test_GraphemeLigature_ffi, test_CharWidth  * 3, 3, true);
			nIndex += 2;
		}
		else if (nCount - nIndex >= 2
			&& test_LETTER.f === HB_String[nIndex]
			&& test_LETTER.f === HB_String[nIndex + 1])
		{
			this.FlushGrapheme(test_GraphemeLigature_ff, test_CharWidth  * 2, 2, true);
			nIndex += 1;
		}
		else if (nCount - nIndex >= 2
			&& test_LETTER.f === HB_String[nIndex]
			&& test_LETTER.i === HB_String[nIndex + 1])
		{
			this.FlushGrapheme(test_GraphemeLigature_fi, test_CharWidth  * 2, 2, true);
			nIndex += 1;
		}
		else if (nCount - nIndex >= 3
			&& test_LETTER.x === HB_String[nIndex]
			&& test_LETTER.y === HB_String[nIndex + 1]
			&& test_LETTER.z === HB_String[nIndex + 2])
		{
			this.FlushGrapheme(test_GraphemeCombining_xyz, test_CharWidth  * 3, 3, false);
			nIndex += 2;
		}
		else if (nCount - nIndex >= 2
			&& test_LETTER.x === HB_String[nIndex]
			&& test_LETTER.y === HB_String[nIndex + 1])
		{
			this.FlushGrapheme(test_GraphemeCombining_xy, test_CharWidth  * 2, 2, false);
			nIndex += 1;
		}
		else
		{
			this.FlushGrapheme(test_GraphemeNormal, test_CharWidth, 1, false);
		}
	}

	AscFonts.HB_StartString();
};

g_oTextMeasurer.SetFontInternal = function() {};
g_oTextMeasurer.SetTextPr = function() {};
g_oTextMeasurer.SetFontSlot = function() {};
g_oTextMeasurer.GetHeight = function() {return test_FontHeight;};
g_oTextMeasurer.GetAscender = function() {return test_FontAscent;};
g_oTextMeasurer.GetDescender = function() {return test_FontDescent;};
g_oTextMeasurer.MeasureCode = function() {return {fAdvanceX : test_CharWidth};};

var AscTest = AscTest || {};

(function(window)
{
	function CreateLogicDocument()
	{
		let logicDocument = new AscWord.CDocument(drawingDocument, true);
		logicDocument.Api = null;
		logicDocument.On_EndLoad();

		drawingDocument.m_oLogicDocument = logicDocument;

		return logicDocument;
	}
	function SetFillingFormMode(oLogicDocument)
	{
		oLogicDocument.IsFillingOFormMode = function()
		{
			return true;
		}
		oLogicDocument.IsFillingFormMode = function()
		{
			return true;
		}
	}
	function SetEditingMode(oLogicDocument)
	{
		oLogicDocument.IsFillingOFormMode = function()
		{
			return false;
		}
		oLogicDocument.IsFillingFormMode = function()
		{
			return false;
		}
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.CreateLogicDocument = CreateLogicDocument;
	AscTest.SetFillingFormMode  = SetFillingFormMode;
	AscTest.SetEditingMode      = SetEditingMode;

})(window);


