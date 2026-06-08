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
	 * Данный класс считает количество графем у заданной строки в заданном шрифте
	 * @constructor
	 */
	function CGraphemesCounter()
	{
		AscFonts.CTextShaper.call(this);

		this.TextPr     = null;
		this.CharsCount = 0;
		this.TrimResult = null;
		this.TrimLength = 0;
	}
	CGraphemesCounter.prototype = Object.create(AscFonts.CTextShaper.prototype);
	CGraphemesCounter.prototype.constructor = CGraphemesCounter;
	CGraphemesCounter.prototype.GetCount = function(sString, oTextPr)
	{
		this.Init();
		
		if (oTextPr)
			this.TextPr = oTextPr;
		else
			this.TextPr = null;

		this.CharsCount = 0;
		this.TrimResult = null;

		this.Shape(sString);
		return this.CharsCount;
	};
	CGraphemesCounter.prototype.Trim = function(sString, nLen, oTextPr)
	{
		this.Init();
		
		if (oTextPr)
			this.TextPr = oTextPr;
		else
			this.TextPr = null;

		this.CharsCount = 0;
		this.TrimResult = [];
		this.TrimLength = nLen;

		this.Shape(sString);
		return (typeof(sString) === "string" ? String.fromCodePoint.apply(String, this.TrimResult) : this.TrimResult);
	};
	CGraphemesCounter.prototype.Init = function()
	{
		this.ClearBuffer();
	};
	CGraphemesCounter.prototype.Shape = function(sString)
	{
		if (typeof(sString) === "string")
		{
			for (let oIter = sString.getUnicodeIterator(); oIter.check(); oIter.next())
			{
				this.HandleCodePoint(oIter.value());
			}
		}
		else if (Array.isArray(sString))
		{
			for (let nPos = 0, nCount = sString.length; nPos < nCount; ++nPos)
			{
				this.HandleCodePoint(sString[nPos]);
			}
		}

		this.FlushWord();
	};
	CGraphemesCounter.prototype.HandleCodePoint = function(nCodePoint)
	{
		if (AscCommon.IsSpace(nCodePoint))
		{
			this.FlushWord();

			if (this.TrimResult && this.CharsCount < this.TrimLength)
				this.TrimResult.push(nCodePoint);

			this.CharsCount++;
		}
		else
		{
			this.AppendToString(nCodePoint);
		}
	};
	CGraphemesCounter.prototype.GetFontInfo = function(nFontSlot)
	{
		if (this.TextPr)
			return this.TextPr.GetFontInfo(nFontSlot);

		return AscFonts.DEFAULT_TEXTFONTINFO;
	};
	CGraphemesCounter.prototype.FlushGrapheme = function(nGrapheme, nWidth, nCodePointsCount, isLigature)
	{
		if (this.TrimResult && this.CharsCount < this.TrimLength)
		{
			let nCount = 0;
			if (isLigature)
				nCount = Math.min(this.TrimLength - this.CharsCount, nCodePointsCount);
			else
				nCount = nCodePointsCount;

			for (let nPos = 0; nPos < nCount; ++nPos)
			{
				this.TrimResult.push(this.GetCodePoint(this.Buffer[this.BufferIndex + nPos]));
			}
		}

		this.CharsCount += isLigature ? nCodePointsCount : 1;
		this.BufferIndex += nCodePointsCount;
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].GraphemesCounter = new CGraphemesCounter();

})(window);
