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
	/**
	 * @constructor
	 */
	function CTextFontInfo(sName, nStyle, nSize)
	{
		this.Name  = undefined !== sName ? sName : "Arial";
		this.Style = undefined !== nStyle ? nStyle : 0;
		this.Size  = undefined !== nSize ? nSize : 10;
	}

	const DEFAULT_TEXTFONTINFO = new CTextFontInfo();

	// Функции для возможной перегрузки
	// 1. FlushGrapheme - основная функция, которую нужно ОБЯЗАТЕЛЬНО реализовывать в дочернем классе
	// 2. Shape - простая функция для шейпинга текста
	// 3. GetFontInfo - получаем информацию о шрифте на момент составления графем
	// 4. GetCodePoint - данная функция нужна, чтобы в буфере можно было хранить не сам юникод, а его содержащий класс
	// 5. GetLigaturesType - какие лигатуры поддерживать

	/**
	 * @constructor
	 */
	function CTextShaper()
	{
		this.Buffer         = [];
		this.BufferIndex    = 0;
		this.Script         = -1;
		this.FontId         = -1;
		this.FontSubst      = false;
		this.FontSlot       = AscWord.fontslot_None;
		this.FontSize       = 10;
		this.ForceCheckFont = false;
		this.Direction      = AscFonts.HB_DIRECTION.HB_DIRECTION_LTR;
	}
	CTextShaper.prototype.ClearBuffer = function()
	{
		this.Buffer.length = 0;
		this.BufferIndex   = 0;

		this.Script    = -1;
		this.FontId    = -1;
		this.FontSlot  = AscWord.fontslot_None;
		this.Direction = AscFonts.HB_DIRECTION.HB_DIRECTION_LTR;

		this.StartString();
	};
	CTextShaper.prototype.ResetBuffer = function()
	{
		if (this.IsRtlDirection())
			this.BufferIndex = this.Buffer.length;
	};
	CTextShaper.prototype.StartString = function()
	{
		AscFonts.HB_StartString();
	};
	CTextShaper.prototype.AppendToString = function(oItem)
	{
		let nCodePoint = this.GetCodePoint(oItem);
		nCodePoint = this.private_CheckNewSegment(nCodePoint);
		this.Buffer.push(oItem);
		AscFonts.HB_AppendToString(nCodePoint);
	};
	CTextShaper.prototype.EndString = function()
	{
		AscFonts.HB_EndString();
	};
	CTextShaper.prototype.FlushWord = function()
	{
		if (this.Buffer.length <= 0)
			return this.ClearBuffer();

		this.EndString();

		let nScript = AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED === this.Script ? AscFonts.HB_SCRIPT.HB_SCRIPT_COMMON : this.Script;
		
		this.Direction = this.GetDirection(nScript);
		this.ResetBuffer();
		
		let oFontInfo = this.GetFontInfo(this.FontSlot);
		let nFontId   = AscCommon.FontNameMap.GetId(this.FontId.m_pFaceInfo.family_name);
		AscCommon.g_oTextMeasurer.SetFontInternal(this.FontId.m_pFaceInfo.family_name, AscFonts.MEASURE_FONTSIZE, oFontInfo.Style);

		this.FontSize = oFontInfo.Size;

		AscFonts.HB_ShapeString(this, nFontId, oFontInfo.Style, this.FontId, this.GetLigaturesType(), nScript, this.Direction, "en");

		// Значит шрифт был подобран, возвращаем назад состояние отрисовщика
		if (this.FontId.m_pFaceInfo.family_name !== oFontInfo.Name)
		{
			AscCommon.g_oTextMeasurer.SetFontInternal(oFontInfo.Name, AscFonts.MEASURE_FONTSIZE, oFontInfo.Style);
			this.ForceCheckFont = true;
		}

		this.ClearBuffer();
	};
	CTextShaper.prototype.GetLigaturesType = function()
	{
		return Asc.LigaturesType.None;
	};
	CTextShaper.prototype.GetTextScript = function(nUnicode)
	{
		return AscFonts.hb_get_script_by_unicode(nUnicode);
	};
	CTextShaper.prototype.GetFontSlot = function(nUnicode)
	{
		return AscWord.GetFontSlot(nUnicode, AscWord.fonthint_Default, lcid_unknown, false, false);
	};
	CTextShaper.prototype.GetDirection = function(nScript)
	{
		return AscFonts.hb_get_script_horizontal_direction(nScript);
	};
	CTextShaper.prototype.private_CheckNewSegment = function(nUnicode)
	{
		if (this.Buffer.length >= AscFonts.HB_STRING_MAX_LEN)
			this.FlushWord();
		
		// TODO: Check bugs 66317 66435
		let nScript = this.GetTextScript(nUnicode);
		if (nScript !== this.Script
			&& -1 !== this.Script
			&& AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED !== nScript
			&& AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED !== this.Script)
			this.FlushWord();
		
		if (this.GetDirection(this.Script) !== this.GetDirection(nScript))
			this.FlushWord();

		let nFontSlot = this.GetFontSlot(nUnicode);
		this.private_CheckFont(nFontSlot);

		let nFontId = this.FontId;
		let isSubst;
		if (AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED === nScript && this.FontSubst)
		{
			let oInfo = AscCommon.g_oTextMeasurer.GetFontBySymbol(nUnicode, -1 !== nFontId ? nFontId : null, true);
			nFontId   = oInfo.Font;
			nUnicode  = oInfo.CodePoint;
			isSubst   = true;
		}
		else
		{
			isSubst = !AscCommon.g_oTextMeasurer.CheckUnicodeInCurrentFont(nUnicode);
			if (-1 === nFontId || isSubst)
			{
				let oInfo = AscCommon.g_oTextMeasurer.GetFontBySymbol(nUnicode, -1 !== nFontId ? nFontId : null);
				nFontId   = oInfo.Font;
				nUnicode  = oInfo.CodePoint;
			}
			else
			{
				let nCurFontId = AscCommon.g_oTextMeasurer.GetCurrentFont();
				if (nCurFontId)
					nFontId = nCurFontId;
			}
		}

		if (this.FontId !== nFontId
			&& -1 !== this.FontId
			&& this.FontSubst
			&& isSubst
			&& this.private_CheckBufferInFont(nFontId))
			this.FontId = nFontId;

		if (this.FontId !== nFontId && -1 !== this.FontId)
			this.FlushWord();

		this.Script    = AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED !== nScript || -1 === this.Script ? nScript : this.Script;
		this.FontSlot  = AscWord.fontslot_None !== nFontSlot ? nFontSlot : this.FontSlot;
		this.FontId    = nFontId;
		this.FontSubst = isSubst;

		return nUnicode;
	};
	CTextShaper.prototype.private_CheckFont = function(nFontSlot)
	{
		if (this.FontSlot === nFontSlot && !this.ForceCheckFont)
			return;
		
		let oFontInfo = this.GetFontInfo(nFontSlot);
		AscCommon.g_oTextMeasurer.SetFontInternal(oFontInfo.Name, AscFonts.MEASURE_FONTSIZE, oFontInfo.Style);
		this.ForceCheckFont = false;
	};
	CTextShaper.prototype.private_CheckBufferInFont = function(nFontId)
	{
		for (let nPos = 0, nCount = this.Buffer.length; nPos < nCount; ++nPos)
		{
			if (!nFontId.GetGIDByUnicode(this.GetCodePoint(this.Buffer[nPos])))
				return false;
		}

		return true;
	};
	CTextShaper.prototype.GetFontInfo = function(nFontSlot)
	{
		return DEFAULT_TEXTFONTINFO;
	};
	CTextShaper.prototype.Shape = function(sString)
	{
		for (var oIterator = sString.getUnicodeIterator(); oIterator.check(); oIterator.next())
		{
			let nCodePoint = oIterator.value();
			this.AppendToString(nCodePoint);
		}
		this.FlushWord();
	};
	CTextShaper.prototype.GetCodePoint = function(oItem)
	{
		return oItem;
	};
	CTextShaper.prototype.FlushGrapheme = function(nGrapheme, nWidth, nCodePointsCount, isLigature)
	{
		if (this.IsRtlDirection())
			this.BufferIndex -= nCodePointsCount;
		else
			this.BufferIndex += nCodePointsCount;
	};
	CTextShaper.prototype.IsRtlDirection = function()
	{
		return this.Direction === AscFonts.HB_DIRECTION.HB_DIRECTION_RTL;
	};
	CTextShaper.prototype.PrintAllUnicodesByScript = function(nScript)
	{
		let log = "";

		function Flush(isForce)
		{
			if (log.length > 100 || isForce)
			{
				console.log(log);
				log = "";
			}
		}

		for (let u = 0; u < 0x10FFFF; ++u)
		{
			if (nScript === this.GetTextScript(u))
			{
				log += AscCommon.IntToHex(u) + " " + String.fromCodePoint(u) + " ";
				Flush(false);
			}
		}

		Flush(true);
	};
	
	function isRtlScript(unicode)
	{
		return AscFonts.hb_get_script_horizontal_direction(AscFonts.hb_get_script_by_unicode(unicode)) === AscFonts.HB_DIRECTION.HB_DIRECTION_RTL;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscFonts'] = window['AscFonts'] || {};
	window['AscFonts'].CTextFontInfo        = CTextFontInfo;
	window['AscFonts'].CTextShaper          = CTextShaper;
	window['AscFonts'].DEFAULT_TEXTFONTINFO = DEFAULT_TEXTFONTINFO;
	window['AscFonts'].isRtlScript          = isRtlScript;

})(window);
