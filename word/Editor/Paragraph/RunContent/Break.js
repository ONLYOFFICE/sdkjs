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
	const break_Line   = 0x01;
	const break_Page   = 0x02;
	const break_Column = 0x03;

	const break_Clear_None  = 0x00;
	const break_Clear_All   = 0x01;
	const break_Clear_Left  = 0x02;
	const break_Clear_Right = 0x03;

	/**
	 * Класс представляющий разрыв строки/колонки/страницы
	 * @param nBreakType
	 * @param nClear {break_Clear_None | break_Clear_All | break_Clear_Left | break_Clear_Right}
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunBreak(nBreakType, nClear)
	{
		AscWord.CRunElementBase.call(this);

		this.BreakType = nBreakType;
		this.Clear     = nClear ? nClear : break_Clear_None;

		this.Flags = {}; // специальные флаги для разных break
		this.Flags.Use = true;

		if (break_Page === this.BreakType || break_Column === this.BreakType)
			this.Flags.NewLine = true;

		this.Height       = 0;
		this.Width        = 0;
		this.WidthVisible = 0;
	}
	CRunBreak.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunBreak.prototype.constructor = CRunBreak;

	CRunBreak.prototype.Type = para_NewLine;
	CRunBreak.prototype.IsBreak = function()
	{
		return true;
	};
	CRunBreak.prototype.Draw = function(X, Y, Context)
	{
		if (false === this.Flags.Use)
			return;

		if(Context.m_bIsTextDrawer)
		{
			Context.CheckSpaceDraw();
		}
		if (undefined !== editor && editor.ShowParaMarks)
		{
			// Цвет и шрифт можно не запоминать и не выставлять старый, т.к. на данном элемента всегда заканчивается
			// отрезок обтекания или целая строка.

			switch (this.BreakType)
			{
				case break_Line:
				{
					Context.b_color1(0, 0, 0, 255);
					Context.SetFont({
						FontFamily : {Name : "ASCW3", Index : -1},
						FontSize   : 10,
						Italic     : false,
						Bold       : false
					});
					Context.FillText(X, Y, String.fromCharCode(0x0038/*0x21B5*/));
					break;
				}
				case break_Page:
				case break_Column:
				{
					var strPageBreak = this.Flags.BreakPageInfo.Str;
					var Widths       = this.Flags.BreakPageInfo.Widths;

					Context.b_color1(0, 0, 0, 255);
					Context.SetFont({
						FontFamily : {Name : "Courier New", Index : -1},
						FontSize   : 8,
						Italic     : false,
						Bold       : false
					});

					var Len = strPageBreak.length;
					for (var Index = 0; Index < Len; Index++)
					{
						Context.FillText(X, Y, strPageBreak[Index]);
						X += Widths[Index];
					}

					break;
				}
			}

		}
	};
	CRunBreak.prototype.Measure = function(Context)
	{
		if (false === this.Flags.Use)
		{
			this.Width        = 0;
			this.WidthVisible = 0;
			this.Height       = 0;
			return;
		}

		switch (this.BreakType)
		{
			case break_Line:
			{
				this.Width  = 0;
				this.Height = 0;

				Context.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});
				var Temp = Context.Measure(String.fromCharCode(0x0038));

				// Почему-то в шрифте Wingding 3 символ 0x0038 имеет неправильную ширину
				this.WidthVisible = Temp.Width * 1.7;

				break;
			}
			case break_Page:
			case break_Column:
			{
				this.Width  = 0;
				this.Height = 0;

				break;
			}
		}
	};
	CRunBreak.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CRunBreak.prototype.GetWidthVisible = function()
	{
		return this.WidthVisible;
	};
	CRunBreak.prototype.SetWidthVisible = function(WidthVisible)
	{
		this.WidthVisible = WidthVisible;
	};
	CRunBreak.prototype.Update_String = function(_W)
	{
		if (false === this.Flags.Use)
		{
			this.Width        = 0;
			this.WidthVisible = 0;
			this.Height       = 0;
			return;
		}

		if (break_Page === this.BreakType || break_Column === this.BreakType)
		{
			var W = false === this.Flags.NewLine ? 50 : Math.max(_W, 50);

			g_oTextMeasurer.SetFont({
				FontFamily : {Name : "Courier New", Index : -1},
				FontSize   : 8,
				Italic     : false,
				Bold       : false
			});

			var Widths = [];

			var nStrWidth    = 0;
			var strBreakPage = break_Page === this.BreakType ? " Page Break " : " Column Break ";
			var Len          = strBreakPage.length;
			for (var Index = 0; Index < Len; Index++)
			{
				var Val       = g_oTextMeasurer.Measure(strBreakPage[Index]).Width;
				nStrWidth += Val;
				Widths[Index] = Val;
			}

			var strSymbol = String.fromCharCode("0x00B7");
			var nSymWidth = g_oTextMeasurer.Measure(strSymbol).Width * 2 / 3;

			var strResult = "";
			if (W - 6 * nSymWidth >= nStrWidth)
			{
				var Count     = parseInt((W - nStrWidth) / ( 2 * nSymWidth ));
				var strResult = strBreakPage;
				for (var Index = 0; Index < Count; Index++)
				{
					strResult = strSymbol + strResult + strSymbol;
					Widths.splice(0, 0, nSymWidth);
					Widths.splice(Widths.length, 0, nSymWidth);
				}
			}
			else
			{
				var Count = parseInt(W / nSymWidth);
				for (var Index = 0; Index < Count; Index++)
				{
					strResult += strSymbol;
					Widths[Index] = nSymWidth;
				}
			}

			var ResultW = 0;
			var Count   = Widths.length;
			for (var Index = 0; Index < Count; Index++)
			{
				ResultW += Widths[Index];
			}

			var AddW = 0;
			if (ResultW < W && Count > 1)
			{
				AddW = (W - ResultW) / (Count - 1);
			}

			for (var Index = 0; Index < Count - 1; Index++)
			{
				Widths[Index] += AddW;
			}

			this.Flags.BreakPageInfo        = {};
			this.Flags.BreakPageInfo.Str    = strResult;
			this.Flags.BreakPageInfo.Widths = Widths;

			this.Width        = W;
			this.WidthVisible = W;
		}
	};
	CRunBreak.prototype.CanAddNumbering = function()
	{
		return (break_Line === this.BreakType);
	};
	CRunBreak.prototype.Copy = function()
	{
		return new CRunBreak(this.BreakType);
	};
	CRunBreak.prototype.IsEqual = function(oElement)
	{
		return (oElement.Type === this.Type && this.BreakType === oElement.BreakType);
	};
	/**
	 * Функция проверяет особый случай, когда у нас PageBreak, после которого в параграфе ничего не идет
	 * @returns {boolean}
	 */
	CRunBreak.prototype.Is_NewLine = function()
	{
		if (break_Line === this.BreakType || ((break_Page === this.BreakType || break_Column === this.BreakType) && true === this.Flags.NewLine))
			return true;

		return false;
	};
	CRunBreak.prototype.Write_ToBinary = function(Writer)
	{
		// Long   : Type
		// Long   : BreakType
		// Optional :
		// Long   : Flags (breakPage)
		Writer.WriteLong(para_NewLine);
		Writer.WriteLong(this.BreakType);

		if (break_Page === this.BreakType || break_Column === this.BreakType)
			Writer.WriteBool(this.Flags.NewLine);
		else
			Writer.WriteLong(this.Clear);
	};
	CRunBreak.prototype.Read_FromBinary = function(Reader)
	{
		this.BreakType = Reader.GetLong();

		if (break_Page === this.BreakType || break_Column === this.BreakType)
			this.Flags = {NewLine : Reader.GetBool()};
		else
			this.Clear = Reader.GetLong();
	};
	/**
	 * Разрыв страницы или колонки?
	 * @returns {boolean}
	 */
	CRunBreak.prototype.IsPageOrColumnBreak = function()
	{
		return (break_Page === this.BreakType || break_Column === this.BreakType);
	};
	/**
	 * Разрыв страницы?
	 * @returns {boolean}
	 */
	CRunBreak.prototype.IsPageBreak = function()
	{
		return (break_Page === this.BreakType);
	};
	/**
	 * Разрыв колонки?
	 * @returns {boolean}
	 */
	CRunBreak.prototype.IsColumnBreak = function()
	{
		return (break_Column === this.BreakType);
	};
	/**
	 * Перенос строки?
	 * @returns {boolean}
	 */
	CRunBreak.prototype.IsLineBreak = function()
	{
		return (break_Line === this.BreakType);
	};
	CRunBreak.prototype.GetAutoCorrectFlags = function()
	{
		return (AscWord.AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE
			| AscWord.AUTOCORRECT_FLAGS_HYPERLINK);
	};
	CRunBreak.prototype.ToSearchElement = function(oProps)
	{
		if (break_Page === this.BreakType)
			return new AscCommonWord.CSearchTextSpecialPageBreak();
		else if (break_Column === this.BreakType)
			return new AscCommonWord.CSearchTextSpecialColumnBreak();
		else
			return new AscCommonWord.CSearchTextSpecialLineBreak();
	};
	CRunBreak.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunBreak         = CRunBreak;
	window['AscWord'].break_Line        = break_Line;
	window['AscWord'].break_Page        = break_Page;
	window['AscWord'].break_Column      = break_Column;
	window['AscWord'].break_Clear_None  = break_Clear_None;
	window['AscWord'].break_Clear_All   = break_Clear_All;
	window['AscWord'].break_Clear_Left  = break_Clear_Left;
	window['AscWord'].break_Clear_Right = break_Clear_Right;
})(window);
