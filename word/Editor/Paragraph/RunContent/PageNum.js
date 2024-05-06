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

		this.numFormat = -1;
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
		this.Set_Page(1, Asc.c_oAscNumberingFormat.Decimal);
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
	CRunPageNum.prototype.SetNumFormat = function(format)
	{
		this.numFormat = format;
	};
	CRunPageNum.prototype.Set_Page = function(pageNum, numFormat)
	{
		if (-1 !== this.numFormat)
			numFormat = this.numFormat;
		
		let numText = AscCommon.IntToNumberFormat(pageNum, numFormat);
		AscWord.stringShaper.Shape(numText.codePointsArray(), this.textPr);
		
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
	
	/**
	 * @constructor
	 */
	function PageNumRecalculateObject(type, graphemes, widths, totalWidth, isCopy)
	{
		this.type      = type;
		this.graphemes = graphemes;
		this.widths    = widths;
		this.width     = totalWidth;
		
		if (isCopy)
		{
			this.graphemes = graphemes.slice();
			this.widths    = widths.slice();
		}
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunPageNum               = CRunPageNum;
	window['AscWord'].CPageNumRecalculateObject = CPageNumRecalculateObject;
	window['AscWord'].PageNumRecalculateObject  = PageNumRecalculateObject;

})(window);
