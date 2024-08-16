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
