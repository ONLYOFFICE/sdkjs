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
	 * Данный класс рассчитывает текстовые настройки шрифта для выделенного текста
	 * @constructor
	 */
	function CFontCalculator()
	{
		this.Bold     = null;
		this.Italic   = null;
		this.FontSize = null;
		this.FontName = null;
		this.Lang     = null;

		// Для случая, когда в выделение попали только элементы со слотом AscWord.fontslot_Unknown
		// (например, только знак конца параграфа)
		this.BoldUnknown     = null;
		this.ItalicUnkown    = null;
		this.FontSizeUnknown = null;
		this.FontNameUnknown = null;
		this.LangUnknown     = null;
	}
	CFontCalculator.prototype.Calculate = function(oDocContent, oTextPr)
	{
		this.Reset();

		if (oDocContent.IsNumberingSelection())
			this.HandleNumberingSelection(oDocContent);
		else if (oDocContent.IsTextSelectionUse() && !oDocContent.IsSelectionEmpty())
			this.HandleRegularSelection(oDocContent);
		else
			this.HandleNoSelection(oDocContent);

		this.CheckResult();

		oTextPr.Bold       = this.Bold;
		oTextPr.Italic     = this.Italic;
		oTextPr.FontSize   = this.FontSize;
		oTextPr.FontFamily = {Name : this.FontName, Index : -1};
		oTextPr.Lang.Val   = this.Lang;
	};
	//------------------------------------------------------------------------------------------------------------------
	CFontCalculator.prototype.Reset = function()
	{
		this.Bold     = null;
		this.Italic   = null;
		this.FontSize = null;
		this.FontName = null;
		this.Lang     = null;

		this.BoldUnknown     = null;
		this.ItalicUnkown    = null;
		this.FontSizeUnknown = null;
		this.FontNameUnknown = null;
		this.LangUnknown     = null;
	};
	CFontCalculator.prototype.HandleNumberingSelection = function(docContent)
	{
		let paragraph = docContent.GetCurrentParagraph();
		if (!paragraph)
			return;

		let textPr = paragraph.GetNumberingTextPr();

		this.Bold     = textPr.Bold;
		this.Italic   = textPr.Italic;
		this.FontSize = textPr.FontSize;
		this.FontName = textPr.RFonts.Ascii ? textPr.RFonts.Ascii.Name : null;
		this.Lang     = textPr.Lang.Val;
	};
	CFontCalculator.prototype.HandleRegularSelection = function(oDocContent)
	{
		let oThis = this;
		oDocContent.CheckSelectedRunContent(function(oRun, nStartPos, nEndPos)
		{
			if (oThis.IsStop())
				return true;

			if (oRun.IsEmpty())
				return false;
			
			let oTextPr;
			if (oRun.IsParaEndRun() && oRun.GetParagraph())
				oTextPr = oRun.GetParagraph().GetParaEndCompiledPr();
			else
				oTextPr = oRun.Get_CompiledPr(false);
			
			let oParagraph = oRun.GetParagraph();
			if (oParagraph)
				oTextPr.ReplaceThemeFonts(oParagraph.GetTheme().themeElements.fontScheme);
			
			let nFontSlot = oRun.GetFontSlotInRange(nStartPos, nEndPos);
			if (AscWord.fontslot_None !== nFontSlot)
			{
				// TODO: Пока оставим эту проверку здесь. Проблема в том, что настройки в последнем ране не совпадают
				//       с настройками конца параграфа (если это исправится, тогда здесь можно просто оставить
				//       простое получение компилированных настроек рана)
				nFontSlot = oThis.CheckUnknownFontSlot(nFontSlot, oTextPr);
				
				if (nFontSlot & AscWord.fontslot_CS)
				{
					oThis.CheckBold(oTextPr.BoldCS);
					oThis.CheckItalic(oTextPr.ItalicCS);
					oThis.CheckFontSize(oTextPr.FontSizeCS);
					oThis.CheckFontName(oTextPr.RFonts.CS.Name);
				}
				
				if (nFontSlot !== AscWord.fontslot_CS && nFontSlot !== AscWord.fontslot_None)
				{
					oThis.CheckBold(oTextPr.Bold);
					oThis.CheckItalic(oTextPr.Italic);
					oThis.CheckFontSize(oTextPr.FontSize);
				}
				
				if (nFontSlot & AscWord.fontslot_ASCII)
					oThis.CheckFontName(oTextPr.RFonts.Ascii.Name);
				
				if (nFontSlot & AscWord.fontslot_HAnsi)
					oThis.CheckFontName(oTextPr.RFonts.HAnsi.Name);
				
				if (nFontSlot & AscWord.fontslot_EastAsia)
					oThis.CheckFontName(oTextPr.RFonts.EastAsia.Name);
			}
			
			let direction = oRun.GetDirectionFlagInRange(nStartPos, nEndPos);
			if (AscBidi.DIRECTION_FLAG.None !== direction)
			{
				if (direction & AscBidi.DIRECTION_FLAG.Other)
					oThis.CheckLangUnknown(oParagraph && oParagraph.isRtlDirection() ? oTextPr.Lang.Bidi : oTextPr.Lang.Val);
				
				if (direction & AscBidi.DIRECTION_FLAG.LTR)
					oThis.CheckLang(oTextPr.Lang.Val);
				
				if (direction & AscBidi.DIRECTION_FLAG.RTL)
					oThis.CheckLang(oTextPr.Lang.Bidi);
			}

			return false;
		});
	};
	CFontCalculator.prototype.HandleNoSelection = function(oDocContent)
	{
		let oParagraph = oDocContent.GetCurrentParagraph();
		if (!oParagraph)
			return;

		let oParaContentPos = oParagraph.Get_ParaContentPos(false, false);
		if (!oParaContentPos)
			return;

		let oRun = oParagraph.GetClassByPos(oParaContentPos);
		if (!oRun || !(oRun instanceof AscWord.ParaRun))
			return;

		let oTextPr   = oRun.Get_CompiledPr(false);
		let nInRunPos = oParaContentPos.Get(oParaContentPos.GetDepth());
		let nFontSlot = this.CheckUnknownFontSlot(oRun.GetFontSlotByPosition(nInRunPos), oTextPr);

		oTextPr.ReplaceThemeFonts(oParagraph.GetTheme().themeElements.fontScheme);

		if (AscWord.fontslot_None === nFontSlot)
			nFontSlot = AscWord.fontslot_ASCII;

		if (nFontSlot & AscWord.fontslot_ASCII)
			this.FontName = oTextPr.RFonts.Ascii.Name;
		else if (nFontSlot & AscWord.fontslot_HAnsi)
			this.FontName = oTextPr.RFonts.HAnsi.Name;
		else if (nFontSlot & AscWord.fontslot_EastAsia)
			this.FontName = oTextPr.RFonts.EastAsia.Name;

		if (nFontSlot & AscWord.fontslot_CS)
		{
			this.Bold     = oTextPr.BoldCS;
			this.Italic   = oTextPr.ItalicCS;
			this.FontSize = oTextPr.FontSizeCS;
			this.FontName = oTextPr.RFonts.CS.Name;
		}
		else if (nFontSlot & AscWord.fontslot_ASCII
			|| nFontSlot & AscWord.fontslot_HAnsi
			|| nFontSlot & AscWord.fontslot_EastAsia)
		{
			this.Bold     = oTextPr.Bold;
			this.Italic   = oTextPr.Italic;
			this.FontSize = oTextPr.FontSize;
		}
		
		let nextEl = oParagraph.GetNextRunElement();
		let prevEl = oParagraph.GetPrevRunElement();
		
		let dir = AscBidi.DIRECTION_FLAG.None;
		if (!nextEl && !prevEl)
		{
			// Use paragraph flag
		}
		else if (!nextEl || !prevEl)
		{
			let el = nextEl ? nextEl : prevEl;
			dir = el.GetDirectionFlag();
		}
		else
		{
			let prevDir = prevEl.GetDirectionFlag();
			let nextDir = nextEl.GetDirectionFlag();
			if (prevDir === nextDir || prevDir === AscBidi.DIRECTION_FLAG.LTR || prevDir === AscBidi.DIRECTION_FLAG.RTL)
				dir = prevDir;
			else
				dir = nextDir;
		}
		
		if (AscBidi.DIRECTION_FLAG.LTR === dir)
			this.Lang = oTextPr.Lang.Val;
		else if (AscBidi.DIRECTION_FLAG.RTL === dir)
			this.Lang = oTextPr.Lang.Bidi;
		else
			this.Lang = oParagraph.isRtlDirection() ? oTextPr.Lang.Bidi : oTextPr.Lang.Val;
	};
	CFontCalculator.prototype.IsStop = function()
	{
		return (undefined === this.Bold
			&& undefined === this.Italic
			&& undefined === this.FontSize
			&& undefined === this.FontName
			&& undefined === this.Lang);
	};
	CFontCalculator.prototype.CheckBold = function(isBold)
	{
		if (undefined === this.Bold)
			return;

		if (null === this.Bold)
			this.Bold = isBold;
		else if (this.Bold !== isBold)
			this.Bold = undefined;
	};
	CFontCalculator.prototype.CheckItalic = function(isItalic)
	{
		if (undefined === this.Italic)
			return;

		if (null === this.Italic)
			this.Italic = isItalic;
		else if (this.Italic !== isItalic)
			this.Italic = undefined;
	};
	CFontCalculator.prototype.CheckFontSize = function(nFontSize)
	{
		if (undefined === this.FontSize)
			return;

		if (null === this.FontSize)
			this.FontSize = nFontSize;
		else if (this.FontSize !== nFontSize)
			this.FontSize = undefined;
	};
	CFontCalculator.prototype.CheckFontName = function(sFontName)
	{
		if (undefined === this.FontName)
			return;

		if (null === this.FontName)
			this.FontName = sFontName;
		else if (this.FontName !== sFontName)
			this.FontName = undefined;
	};
	CFontCalculator.prototype.CheckLang = function(lang)
	{
		if (undefined === this.Lang)
			return;
		
		if (null === this.Lang)
			this.Lang = lang;
		else if (this.Lang !== lang)
			this.Lang = undefined;
	};
	CFontCalculator.prototype.CheckBoldUnknown = function(isBold)
	{
		if (null === this.BoldUnknown)
			this.BoldUnknown = isBold;
	};
	CFontCalculator.prototype.CheckItalicUnknown = function(isItalic)
	{
		if (null === this.ItalicUnkown)
			this.ItalicUnkown = isItalic;
	};
	CFontCalculator.prototype.CheckFontSizeUnknown = function(nFontSize)
	{
		if (null === this.FontSizeUnknown)
			this.FontSizeUnknown = nFontSize;
	};
	CFontCalculator.prototype.CheckFontNameUnknown = function(sFontName)
	{
		if (null === this.FontNameUnknown)
			this.FontNameUnknown = sFontName;
	};
	CFontCalculator.prototype.CheckLangUnknown = function(lang)
	{
		if (undefined === this.LangUnknown)
			return;
		
		if (null === this.LangUnknown)
			this.LangUnknown = lang;
		else if (lang !== this.LangUnknown)
			this.LangUnknown = undefined;
	};
	CFontCalculator.prototype.CheckUnknownFontSlot = function(nFontSlot, oTextPr)
	{
		let _nFontSlot = nFontSlot;
		if (AscWord.fontslot_Unknown === nFontSlot)
		{
			if (oTextPr.CS || oTextPr.RTL)
			{
				this.CheckBoldUnknown(oTextPr.BoldCS);
				this.CheckItalicUnknown(oTextPr.ItalicCS);
				this.CheckFontSizeUnknown(oTextPr.FontSizeCS);
				this.CheckFontNameUnknown(oTextPr.RFonts.CS.Name);
			}
			else
			{
				this.CheckBoldUnknown(oTextPr.Bold);
				this.CheckItalicUnknown(oTextPr.Italic);
				this.CheckFontSizeUnknown(oTextPr.FontSize);
				this.CheckFontNameUnknown(oTextPr.RFonts.Ascii.Name);
			}

			_nFontSlot = AscWord.fontslot_None;
		}
		else if ((AscWord.fontslot_Unknown | AscWord.fontslot_CS) === nFontSlot)
		{
			_nFontSlot = AscWord.fontslot_CS;
		}

		return _nFontSlot;
	};
	CFontCalculator.prototype.CheckResult = function()
	{
		if (null === this.Bold)
		{
			if (null !== this.BoldUnknown)
			{
				this.Bold     = this.BoldUnknown;
				this.Italic   = this.ItalicUnkown;
				this.FontSize = this.FontSizeUnknown;
				this.FontName = this.FontNameUnknown;
			}
			else
			{
				this.Bold     = undefined;
				this.Italic   = undefined;
				this.FontSize = undefined;
				this.FontName = undefined;
			}
		}
		
		if (null === this.Lang)
			this.Lang = this.LangUnknown ? this.LangUnknown : undefined;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].FontCalculator = new CFontCalculator();

})(window);
