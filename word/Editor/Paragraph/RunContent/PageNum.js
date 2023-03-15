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
	 * Класс представляющий элемент номер страницы
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunPageNum()
	{
		AscWord.CRunElementBase.call(this);

		this.FontKoef = 1;

		this.NumWidths = [];

		this.Widths = [];
		this.String = [];

		this.Width        = 0;
		this.WidthVisible = 0;

		this.Parent = null;
	}
	CRunPageNum.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunPageNum.prototype.constructor = CRunPageNum;

	CRunPageNum.prototype.Type = para_PageNum;
	CRunPageNum.prototype.Draw = function(X, Y, Context)
	{
		// Value - реальное значение, которое должно быть отрисовано.
		// Align - прилегание параграфа, в котором лежит данный номер страницы.

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
	CRunPageNum.prototype.Measure = function (Context, TextPr)
	{
		this.FontKoef = TextPr.Get_FontKoef();
		Context.SetFontSlot(AscWord.fontslot_ASCII, this.FontKoef);

		for (var Index = 0; Index < 10; Index++)
		{
			this.NumWidths[Index] = Context.Measure("" + Index).Width;
		}

		this.Width        = 0;
		this.Height       = 0;
		this.WidthVisible = 0;
	};
	CRunPageNum.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CRunPageNum.prototype.GetWidthVisible = function()
	{
		return this.WidthVisible;
	};
	CRunPageNum.prototype.SetWidthVisible = function(WidthVisible)
	{
		this.WidthVisible = WidthVisible;
	};
	CRunPageNum.prototype.Set_Page = function(PageNum)
	{
		this.String = "" + PageNum;
		var Len     = this.String.length;

		var RealWidth = 0;
		for (var Index = 0; Index < Len; Index++)
		{
			var Char = parseInt(this.String.charAt(Index));

			this.Widths[Index] = this.NumWidths[Char];
			RealWidth += this.NumWidths[Char];
		}

		this.Width        = RealWidth;
		this.WidthVisible = RealWidth;
	};
	CRunPageNum.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunPageNum.prototype.SaveRecalculateObject = function(Copy)
	{
		return new CPageNumRecalculateObject(this.Type, this.Widths, this.String, this.Width, Copy);
	};
	CRunPageNum.prototype.LoadRecalculateObject = function(RecalcObj)
	{
		this.Widths = RecalcObj.Widths;
		this.String = RecalcObj.String;

		this.Width        = RecalcObj.Width;
		this.WidthVisible = this.Width;
	};
	CRunPageNum.prototype.PrepareRecalculateObject = function()
	{
		this.Widths = [];
		this.String = "";
	};
	CRunPageNum.prototype.Document_CreateFontCharMap = function(FontCharMap)
	{
		var sValue = "1234567890";
		for (var Index = 0; Index < sValue.length; Index++)
		{
			var Char = sValue.charAt(Index);
			FontCharMap.AddChar(Char);
		}
	};
	CRunPageNum.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunPageNum.prototype.Copy = function()
	{
		return new CRunPageNum();
	};
	CRunPageNum.prototype.Write_ToBinary = function(Writer)
	{
		// Long   : Type
		Writer.WriteLong(para_PageNum);
	}
	CRunPageNum.prototype.Read_FromBinary = function(Reader)
	{
	};
	CRunPageNum.prototype.GetPageNumValue = function()
	{
		var nPageNum = parseInt(this.String);
		if (isNaN(nPageNum))
			return 1;

		return nPageNum;
	};
	CRunPageNum.prototype.GetType = function()
	{
		return this.Type;
	};
	/**
	 * Выставляем родительский класс
	 * @param {ParaRun} oParent
	 */
	CRunPageNum.prototype.SetParent = function(oParent)
	{
		this.Parent = oParent;
	};
	/**
	 * Получаем родительский класс
	 * @returns {?ParaRun}
	 */
	CRunPageNum.prototype.GetParent = function()
	{
		return this.Parent;
	};
	CRunPageNum.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};

	/**
	 * @constructor
	 */
	function CPageNumRecalculateObject(Type, Widths, String, Width, Copy)
	{
		this.Type   = Type;
		this.Widths = Widths;
		this.String = String;
		this.Width  = Width;

		if ( true === Copy )
		{
			this.Widths = [];
			var Len = Widths.length;
			for ( var Index = 0; Index < Len; Index++ )
				this.Widths[Index] = Widths[Index];
		}
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunPageNum               = CRunPageNum;
	window['AscWord'].CPageNumRecalculateObject = CPageNumRecalculateObject;

})(window);
