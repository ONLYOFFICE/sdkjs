/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window)
{
	/**
	 * Class for shaping simple text string using single font
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
