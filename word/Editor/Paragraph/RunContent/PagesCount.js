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
	 * Класс представляющий элемент "количество страниц"
	 * @param PageCount
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunPagesCount(PageCount)
	{
		AscWord.CRunElementBase.call(this);

		this.FontKoef  = 1;
		this.NumWidths = [];
		this.Widths    = [];
		this.String    = "";
		this.PageCount = undefined !== PageCount ? PageCount : 1;
		this.Parent    = null;
	}
	CRunPagesCount.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunPagesCount.prototype.constructor = CRunPagesCount;

	CRunPagesCount.prototype.Type = para_PageCount;
	CRunPagesCount.prototype.Copy = function()
	{
		return new CRunPagesCount();
	};
	CRunPagesCount.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunPagesCount.prototype.Measure = function(Context, TextPr)
	{
		this.FontKoef = TextPr.getFontCoef();
		Context.SetFontSlot(AscWord.fontslot_ASCII, this.FontKoef);

		for (var Index = 0; Index < 10; Index++)
		{
			this.NumWidths[Index] = Context.Measure("" + Index).Width;
		}

		this.private_UpdateWidth();
	};
	CRunPagesCount.prototype.Draw = function(X, Y, Context)
	{
		var Len = this.String.length;

		var _X = X;
		var _Y = Y;

		Context.SetFontSlot(AscWord.fontslot_ASCII, this.FontKoef);
		for (var Index = 0; Index < Len; Index++)
		{
			var Char = this.String.charAt(Index);
			Context.FillText(_X, _Y, Char);
			_X += this.Widths[Index];
		}
	};
	CRunPagesCount.prototype.Document_CreateFontCharMap = function(FontCharMap)
	{
		var sValue = "1234567890";
		for (var Index = 0; Index < sValue.length; Index++)
		{
			var Char = sValue.charAt(Index);
			FontCharMap.AddChar(Char);
		}
	};
	CRunPagesCount.prototype.Update_PageCount = function(nPageCount)
	{
		this.PageCount = nPageCount;
		this.private_UpdateWidth();
	};
	CRunPagesCount.prototype.SetNumValue = function(nValue)
	{
		this.Update_PageCount(nValue);
	};
	CRunPagesCount.prototype.private_UpdateWidth = function()
	{
		this.String = "" + this.PageCount;

		var RealWidth = 0;
		for (var Index = 0, Len = this.String.length; Index < Len; Index++)
		{
			var Char = parseInt(this.String.charAt(Index));

			this.Widths[Index] = this.NumWidths[Char];
			RealWidth += this.NumWidths[Char];
		}

		RealWidth = (RealWidth * AscWord.TEXTWIDTH_DIVIDER) | 0;

		this.Width        = RealWidth;
		this.WidthVisible = RealWidth;
	};
	CRunPagesCount.prototype.Write_ToBinary = function(Writer)
	{
		// Long : Type
		// Long : PageCount
		Writer.WriteLong(this.Type);
		Writer.WriteLong(this.PageCount);
	};
	CRunPagesCount.prototype.Read_FromBinary = function(Reader)
	{
		this.PageCount = Reader.GetLong();
	};
	CRunPagesCount.prototype.GetPageCountValue = function()
	{
		return this.PageCount;
	};
	/**
	 * Выставляем родительский класс
	 * @param {ParaRun} oParent
	 */
	CRunPagesCount.prototype.SetParent = function(oParent)
	{
		this.Parent = oParent;
	};
	/**
	 * Получаем родительский класс
	 * @returns {?ParaRun}
	 */
	CRunPagesCount.prototype.GetParent = function()
	{
		return this.Parent;
	};
	CRunPagesCount.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunPagesCount = CRunPagesCount;

})(window);
