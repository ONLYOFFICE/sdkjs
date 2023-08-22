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

var AscHyphenation = AscHyphenation || {};
(function()
{
	// Это тестовые функции, которые должны быть заменены на нормальные
	let BUFFER_STRING = "";
	
	AscHyphenation.addCodePoint = function(codePoint)
	{
		BUFFER_STRING += String.fromCodePoint(codePoint);
	};
	AscHyphenation.hyphenate = function()
	{
		let checkString = BUFFER_STRING.toLowerCase();
		
		if ("reenter" === checkString)
			return [1];
		else if ("coauthor" === checkString)
			return [1];
		else if ("привет" === checkString)
			return [2];
		else if ("участвовавшими" === checkString)
			return [3, 6, 9, 11];
		else if ("припри" === checkString)
			return [2];
		else if ("приприпри" === checkString)
			return [2, 5];
		else if ("приприприпри" === checkString)
			return [2, 5, 8];
		else if ("приприприприпри" === checkString)
			return [2, 5, 8, 11];
		else if ("приприприприприпри" === checkString)
			return [2, 5, 8, 11, 14];
		else if ("приприприприприприпри" === checkString)
			return [2, 5, 8, 11, 14, 17];
		else if ("приприприприприприприпри" === checkString)
			return [2, 5, 8, 11, 14, 17, 20];
		
		return [];
	};
	AscHyphenation.clear = function()
	{
		BUFFER_STRING = "";
	};
})();

(function(window)
{
	let hyphenator = null;
	function getInstance()
	{
		if (!hyphenator)
			hyphenator = new TextHyphenator();
		
		return hyphenator;
	}
	
	const DEFAULT_LANG = lcid_enUS;
	
	/**
	 * Класс для автоматической расстановки переносов в тексте
	 * @constructor
	 */
	function TextHyphenator()
	{
		this.word     = false;
		this.fontSlot = fontslot_Unknown;
		this.lang     = DEFAULT_LANG;
		this.buffer   = [];
		
		this.hyphenateCaps = true;
	}
	TextHyphenator.hyphenate = function(paragraph)
	{
		let hyphenator = getInstance();
		hyphenator.hyphenateParagraph(paragraph);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	TextHyphenator.prototype.hyphenateParagraph = function(paragraph)
	{
		this.checkHyphenateCaps(paragraph);
		
		let self = this;
		paragraph.CheckRunContent(function(run, startPos, endPos)
		{
			self.hyphenateRun(run, startPos, endPos);
		});
		this.flushWord();
	};
	TextHyphenator.prototype.hyphenateRun = function(run, startPos, endPos)
	{
		for (let pos = startPos; pos < endPos; ++pos)
		{
			let item = run.GetElement(pos);
			if (!item.IsText())
			{
				this.flushWord();
			}
			else if (item.IsNBSP() || item.IsPunctuation())
			{
				this.flushWord();
			}
			else
			{
				if (!this.word)
					this.getLang(run, item.GetFontSlot(run.Get_CompiledPr(false)));
				
				this.appendToWord(item);
				
				if (item.IsSpaceAfter())
					this.flushWord();
			}
		}
	};
	TextHyphenator.prototype.resetBuffer = function()
	{
		this.buffer.length = 0;
		AscHyphenation.clear();
	};
	TextHyphenator.prototype.getLang = function(run, fontSlot)
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
		
		this.lang     = lang;
		this.fontSlot = fontSlot;
	};
	TextHyphenator.prototype.appendToWord = function(textItem)
	{
		this.word = true;
		this.buffer.push(textItem);
		AscHyphenation.addCodePoint(textItem.GetCodePoint());
		textItem.SetHyphenAfter(false);
	};
	TextHyphenator.prototype.flushWord = function()
	{
		if (!this.word)
			return;
		
		this.word = false;
		
		if (!this.isHyphenateCaps() && this.isAllCaps())
			return this.resetBuffer();
		
		let result = AscHyphenation.hyphenate();
		for (let i = 0, len = result.length; i < len; ++i)
		{
			this.buffer[result[i]].SetHyphenAfter(true);
		}
		
		this.resetBuffer();
	};
	TextHyphenator.prototype.isAllCaps = function()
	{
		for (let i = 0, len = this.buffer.length; i < len; ++i)
		{
			let char = String.fromCodePoint(this.buffer[i].GetCodePoint());
			if (char.toUpperCase() !== char)
				return false;
		}
		
		return true;
	};
	TextHyphenator.prototype.checkHyphenateCaps = function(paragraph)
	{
		this.hyphenateCaps = true;
		
		let logicDocument = paragraph.GetLogicDocument();
		if (logicDocument && logicDocument.IsDocumentEditor())
			this.hyphenateCaps = logicDocument.GetDocumentSettings().isHyphenateCaps();
	};
	TextHyphenator.prototype.isHyphenateCaps = function()
	{
		return this.hyphenateCaps;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].TextHyphenator = TextHyphenator;
	
})(window);

