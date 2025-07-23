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
	const FLAGS_MASK               = 0xFF;
	const FLAGS_FONTKOEF_SCRIPT    = 0x01;
	const FLAGS_FONTKOEF_SMALLCAPS = 0x02;
	const FLAGS_GAPS               = 0x04;
	
	// 16-31 биты зарезервированы под FontSize

	const FLAGS_NON_FONTKOEF_SCRIPT    = FLAGS_MASK ^ FLAGS_FONTKOEF_SCRIPT;
	const FLAGS_NON_FONTKOEF_SMALLCAPS = FLAGS_MASK ^ FLAGS_FONTKOEF_SMALLCAPS;
	const FLAGS_NON_GAPS               = FLAGS_MASK ^ FLAGS_GAPS;

	/**
	 * Класс представляющий пробелбный символ
	 * @param {number} [nCharCode=0x20] - Юникодное значение символа
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunSpace(nCharCode)
	{
		AscWord.CRunElementBase.call(this);

		this.Value        = undefined !== nCharCode ? nCharCode : 0x20;
		this.Flags        = 0x00000000 | 0;
		this.Width        = 0x00000000 | 0;
		this.WidthVisible = 0x00000000 | 0;
		this.WidthOrigin  = 0x00000000 | 0;
		this.Grapheme     = AscFonts.NO_GRAPHEME;

		if (AscFonts.IsCheckSymbols)
			AscFonts.FontPickerByCharacter.getFontBySymbol(this.Value);
	}
	CRunSpace.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunSpace.prototype.constructor = CRunSpace;

	CRunSpace.prototype.Type = para_Space;
	CRunSpace.prototype.IsSpace = function()
	{
		return true;
	};
	CRunSpace.prototype.SetWidth = function(width)
	{
		this.Width = ((width * (((this.Flags >> 16) & 0xFFFF) / 64)) * AscWord.TEXTWIDTH_DIVIDER) | 0;
	};
	CRunSpace.prototype.GetWidth = function()
	{
		let nWidth = (this.Width / AscWord.TEXTWIDTH_DIVIDER);

		if (this.Flags & FLAGS_GAPS)
			nWidth += this.LGap + this.RGap;

		return (nWidth > 0 ? nWidth : 0);
	};
	CRunSpace.prototype.GetCodePoint = function()
	{
		return this.Value;
	};
	CRunSpace.prototype.getBidiType = function()
	{
		return AscBidi.TYPE.WS;
	};
	CRunSpace.prototype.SetGrapheme = function(grapheme)
	{
		this.Grapheme = grapheme;
	};
	CRunSpace.prototype.SetMetrics = function(fontSize, fontSlot, textPr)
	{
		this.SetScriptFlag(textPr.VertAlign !== AscCommon.vertalign_Baseline);
		this.SetSmallCapsFlag(!textPr.Caps && textPr.SmallCaps);
		let fontCoef = this.GetFontCoef();
		// Разрешенные размеры шрифта только либо целое, либо целое/2. Даже после применения fontCoef, поэтому
		// мы должны подкрутить коэффициент так, чтобы после домножения на него, у на получался разрешенный размер.
		fontSize = (((fontSize * fontCoef * 2 + 0.5) | 0) / 2);
		this.Flags = (this.Flags & 0xFFFF) | (((fontSize * 64) & 0xFFFF) << 16);
	};
	CRunSpace.prototype.Draw = function(X, Y, Context, PDSE, oTextPr)
	{
		if (this.Flags & FLAGS_GAPS)
		{
			this.DrawGapsBackground(X, Y, Context, PDSE, oTextPr);
			X += this.LGap;
		}

		if(Context.m_bIsTextDrawer)
		{
			Context.CheckSpaceDraw(this);
		}
		
		if (AscFonts.NO_GRAPHEME !== this.Grapheme)
		{
			let fontSize = (((this.Flags >> 16) & 0xFFFF) / 64);
			AscFonts.DrawGrapheme(this.Grapheme, Context, X, Y, fontSize);
		}
		else if (undefined !== editor && editor.ShowParaMarks)
		{
			Context.SetFontSlot(AscWord.fontslot_ASCII, this.GetFontCoef());

			if (this.SpaceGap)
				X += this.SpaceGap;

			if (0x2003 === this.Value || 0x2002 === this.Value)
				Context.FillText(X, Y, String.fromCharCode(0x00B0));
			else if (0x2005 === this.Value)
				Context.FillText(X, Y, String.fromCharCode(0x007C));
			else
				Context.FillText(X, Y, String.fromCharCode(0x00B7));
		}
	};
	CRunSpace.prototype.SetWidth = function(width, textPr, enWidth)
	{
		let fontSize = (((this.Flags >> 16) & 0xFFFF) / 64);
		let Temp = width * fontSize;
		
		var ResultWidth  = (Math.max((Temp + textPr.Spacing), 0) * AscWord.TEXTWIDTH_DIVIDER) | 0;
		this.Width       = ResultWidth;
		this.WidthOrigin = ResultWidth;
		
		if (Math.abs(Temp) > 0.001)
			this.WidthEn = (ResultWidth * enWidth / width) | 0;
		else
			this.WidthEn = ResultWidth;
		
		if (0x2003 === this.Value || 0x2002 === this.Value || 0x2005 === this.Value)
		{
			g_oTextMeasurer.SetTextPr(textPr);
			g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, this.GetFontCoef());
			let code = 0x2005 === this.Value ? 0x007C : 0x00B0;
			this.SpaceGap = Math.max((Temp - g_oTextMeasurer.MeasureCode(code).Width) / 2, 0);
		}
		else if (undefined !== this.SpaceGap)
		{
			this.SpaceGap = 0;
		}
		
		// Не меняем здесь WidthVisible, это значение для пробела высчитывается отдельно, и не должно меняться при пересчете
	};
	CRunSpace.prototype.Measure = function(Context, TextPr)
	{
		if (this.Flags & FLAGS_GAPS)
		{
			this.Flags &= FLAGS_NON_GAPS;
			this.LGap = 0;
			this.RGap = 0;
		}
	};
	CRunSpace.prototype.GetFontCoef = function()
	{
		if (this.Flags & FLAGS_FONTKOEF_SCRIPT && this.Flags & FLAGS_FONTKOEF_SMALLCAPS)
			return smallcaps_and_script_koef;
		else if (this.Flags & FLAGS_FONTKOEF_SCRIPT)
			return AscCommon.vaKSize;
		else if (this.Flags & FLAGS_FONTKOEF_SMALLCAPS)
			return smallcaps_Koef;
		else
			return 1;
	};
	CRunSpace.prototype.SetScriptFlag = function(isScript)
	{
		if (isScript)
			this.Flags |= FLAGS_FONTKOEF_SCRIPT;
		else
			this.Flags &= FLAGS_NON_FONTKOEF_SCRIPT;
	};
	CRunSpace.prototype.SetSmallCapsFlag = function(isSmallCaps)
	{
		if (isSmallCaps)
			this.Flags |= FLAGS_FONTKOEF_SMALLCAPS;
		else
			this.Flags &= FLAGS_NON_FONTKOEF_SMALLCAPS;
	};
	CRunSpace.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunSpace.prototype.Copy = function()
	{
		return new CRunSpace(this.Value);
	};
	CRunSpace.prototype.Write_ToBinary = function(Writer)
	{
		// Long : Type
		// Long : Value

		Writer.WriteLong(para_Space);
		Writer.WriteLong(this.Value);
	};
	CRunSpace.prototype.Read_FromBinary = function(Reader)
	{
		this.Value = Reader.GetLong();
	};
	CRunSpace.prototype.GetAutoCorrectFlags = function()
	{
		return AscWord.AUTOCORRECT_FLAGS_ALL;
	};
	CRunSpace.prototype.SetCondensedWidth = function(nKoef)
	{
		this.Width = this.WidthOrigin * nKoef;
	};
	CRunSpace.prototype.ResetCondensedWidth = function()
	{
		this.Width = this.WidthOrigin;
	};
	CRunSpace.prototype.BalanceSingleByteDoubleByteWidth = function()
	{
		// ea-space doesn't need to be balanced (bug 58483)
		if (this.Value === 0x3000)
			return;
		
		this.Width = this.WidthEn;
	};
	CRunSpace.prototype.SetGaps = function(nLeftGap, nRightGap, nCellWidth)
	{
		this.Flags |= FLAGS_GAPS;

		this.LGap = nLeftGap;
		this.RGap = nRightGap;
	};
	CRunSpace.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextItemChar(0x20);
	};
	CRunSpace.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	CRunSpace.prototype.ToMathElement = function()
	{
		let oSpace = new CMathText();
		oSpace.add(0x0020);
		return oSpace;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunSpace = CRunSpace;

})(window);
