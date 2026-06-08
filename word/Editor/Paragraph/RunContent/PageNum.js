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
	 * Класс представляющий элемент номер страницы
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunPageNum()
	{
		AscWord.CRunElementBase.call(this);
		
		this.Width        = 0;
		this.WidthVisible = 0;
		
		this.parent    = null;
		this.numFormat = -1;
		this.numText   = "1";
		this.pageNum   = 1;
		this.textPr    = null;
		this.graphemes = [];
		this.widths    = [];
	}
	CRunPageNum.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunPageNum.prototype.constructor = CRunPageNum;

	CRunPageNum.prototype.Type = para_PageNum;
	CRunPageNum.prototype.Draw = function(x, y, context)
	{
		let fontSize = this.textPr.FontSize * this.textPr.getFontCoef();
		for (let index = 0; index < this.graphemes.length; ++index)
		{
			AscFonts.DrawGrapheme(this.graphemes[index], context, x, y, fontSize);
			x += this.widths[index] * fontSize;
		}
	};
	CRunPageNum.prototype.Measure = function (Context, TextPr)
	{
		this.textPr = TextPr;
		this._measure();
	};
	CRunPageNum.prototype.SetNumFormat = function(format)
	{
		this.numFormat = format;
	};
	CRunPageNum.prototype.SetValue = function(pageNum, numFormat)
	{
		if (-1 !== this.numFormat)
			numFormat = this.numFormat;
		
		this.pageNum = pageNum;
		this.numText = AscCommon.IntToNumberFormat(pageNum, numFormat);
		
		this._measure();
	};
	CRunPageNum.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunPageNum.prototype.SaveRecalculateObject = function(isCopy)
	{
		return new AscWord.PageNumRecalculateObject(this.Type, this.graphemes, this.widths, this.Width, isCopy);
	};
	CRunPageNum.prototype.LoadRecalculateObject = function(recalcObj)
	{
		this.graphemes    = recalcObj.graphemes;
		this.widths       = recalcObj.widths;
		this.Width        = recalcObj.width;
		this.WidthVisible = this.Width;
	};
	CRunPageNum.prototype.PrepareRecalculateObject = function()
	{
		this.graphemes = [];
		this.widths    = [];
	};
	CRunPageNum.prototype.Document_CreateFontCharMap = function(charMap)
	{
		let numFormat = (-1 !== this.numFormat ? this.numFormat : Asc.c_oAscNumberingFormat.Decimal);
		let symbols = AscWord.GetNumberingSymbolsByFormat(numFormat);
		for (let iter = symbols.getUnicodeIterator(); iter.check(); iter.next())
		{
			charMap.AddChar(String.fromCodePoint(iter.value()));
		}
	};
	CRunPageNum.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunPageNum.prototype.GetValue = function()
	{
		return this.pageNum;
	};
	/**
	 * Выставляем родительский класс
	 * @param {ParaRun} parent
	 */
	CRunPageNum.prototype.SetParent = function(parent)
	{
		this.parent = parent;
	};
	/**
	 * Получаем родительский класс
	 * @returns {?ParaRun}
	 */
	CRunPageNum.prototype.GetParent = function()
	{
		return this.parent;
	};
	CRunPageNum.prototype.GetFontSlot = function(textPr)
	{
		return AscWord.fontslot_Unknown;
	};
	CRunPageNum.prototype.ToString = function()
	{
		return this.numText;
	};
	CRunPageNum.prototype._measure = function()
	{
		if (!this.textPr)
			return;
		
		AscWord.stringShaper.Shape(this.numText.codePointsArray(), this.textPr);
		
		this.graphemes = AscWord.stringShaper.GetGraphemes();
		this.widths    = AscWord.stringShaper.GetWidths();
		
		let totalWidth = 0;
		for (let index = 0; index < this.widths.length; ++index)
		{
			totalWidth += this.widths[index];
		}
		let fontSize = this.textPr.FontSize * this.textPr.getFontCoef();
		totalWidth = (totalWidth * fontSize * AscWord.TEXTWIDTH_DIVIDER) | 0;
		
		this.Width        = totalWidth;
		this.WidthVisible = totalWidth;
	};

	/**
	 * @constructor
	 */
	function PageNumRecalculateObject(type, graphemes, widths, totalWidth, isCopy)
	{
		this.type      = type;
		this.graphemes = graphemes && isCopy ? graphemes.slice() : graphemes;
		this.widths    = widths && isCopy ? widths.slice() : widths;
		this.width     = totalWidth;
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunPageNum               = CRunPageNum;
	window['AscWord'].PageNumRecalculateObject  = PageNumRecalculateObject;

})(window);
