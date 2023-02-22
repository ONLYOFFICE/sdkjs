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
	const DEFAULT_LANG = lcid_enUS;
	
	/**
	 * Класс для автоматической расстановки переносов в тексте
	 * @constructor
	 */
	function CTextHyphenator()
	{
		this.Word     = false;
		this.FontSlot = fontslot_Unknown;
		this.Lang     = lcid_enUS;
	}
	CTextHyphenator.prototype.Hyphenate = function(paragraph)
	{
		let self = this;
		paragraph.CheckRunContent(function(run, startPos, endPos)
		{
			self.HyphenateRun(run, startPos, endPos);
		});
		this.FlushWord();
	};
	CTextHyphenator.prototype.HyphenateRun = function(run, startPos, endPos)
	{
		for (let pos = startPos; nPos < endPos; ++pos)
		{
			let item = run.GetElement(pos);
			if (!item.IsText())
			{
				this.FlushWord();
			}
			else if (item.IsNBSP() || item.IsPunctuation())
			{
				this.FlushWord();
			}
			else
			{
				if (!this.Word)
					this.GetLang(run, item.GetFontSlot());
				
				if (this.Word)
					this.AppendToWord(item);
				
				if (oItem.IsSpaceAfter())
					this.FlushWord();
			}
		}
	};
	CTextHyphenator.prototype.GetLang = function(run, fontSlot)
	{
		let textPr = run.Get_CompiledPr(false);
		let lang;
		switch (fontSlot)
		{
			case fontslot_EastAsia:
				lang = textPr.Lang.EastAsia;
				break;
			case fontslot_CS:
				lang = textPr.Lang.Bidi;
				break;
			case fontslot_HAnsi:
			case fontslot_ASCII:
			default:
				lang = textPr.Lang.Val;
				break;
		}
		
		if (textPr.CS)
		{
			lang     = textPr.Lang.Bidi;
			fontSlot = fontslot_CS;
		}
		
		this.Lang     = lang;
		this.FontSlot = fontSlot;
	};
	CTextHyphenator.prototype.AppendToWord = function(textItem)
	{
	
	};
	CTextHyphenator.prototype.FlushWord = function()
	{
		if (!this.Word)
			return;
		
		
		this.Word = false;
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTextHyphenator = CTextHyphenator;
	window['AscWord'].TextHyphenator  = new  CTextHyphenator();
	
})(window);

