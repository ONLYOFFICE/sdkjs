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
	/**
	 * Класс для шейпинга простой текстовой строки, идущей с одном шрифте
	 * @constructor
	 */
	function StringShaper()
	{
		AscFonts.CTextShaper.call(this);
		
		this.textPr    = null;
		this.graphemes = [];
		this.widths    = [];
	}
	StringShaper.prototype = Object.create(AscFonts.CTextShaper.prototype);
	StringShaper.prototype.constructor = StringShaper;
	StringShaper.prototype.Shape = function(codePoints, textPr)
	{
		this.textPr     = textPr;
		this.codePoints = codePoints.slice();
		this.graphemes  = [];
		this.widths     = [];
		
		for (let i = 0; i < codePoints.length; ++i)
		{
			this.AppendToString(codePoints[i]);
		}
		this.FlushWord();
	};
	StringShaper.prototype.FlushGrapheme = function(grapheme, width, codePointCount, isLigature)
	{
		if (codePointCount <= 0)
			return
		
		this.graphemes.push(grapheme);
		this.widths.push(width);
		
		if (this.IsRtlDirection())
			this.BufferIndex -= codePointCount;
		else
			this.BufferIndex += codePointCount;
	};
	StringShaper.prototype.GetFontSlot = function(codePoint)
	{
		if (!this.textPr)
			return AscWord.fontslot_ASCII;
		
		return AscWord.GetFontSlotByTextPr(codePoint, this.textPr);
	};
	StringShaper.prototype.GetDirection = function(script)
	{
		return AscFonts.HB_DIRECTION.HB_DIRECTION_LTR;
	};
	StringShaper.prototype.GetFontInfo = function(fontSlot)
	{
		if (!this.textPr)
			return AscFonts.DEFAULT_TEXTFONTINFO;
		
		return this.textPr.GetFontInfo(fontSlot);
	};
	StringShaper.prototype.GetGraphemes = function()
	{
		return this.graphemes;
	};
	StringShaper.prototype.GetWidths = function()
	{
		return this.widths;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscFonts'] = window['AscFonts'] || {};
	window['AscFonts'].StringShaper = StringShaper;
})(window);
