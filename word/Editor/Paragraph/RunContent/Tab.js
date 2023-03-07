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

var tab_Bar     = Asc.c_oAscTabType.Bar;
var tab_Center  = Asc.c_oAscTabType.Center;
var tab_Clear   = Asc.c_oAscTabType.Clear;
var tab_Decimal = Asc.c_oAscTabType.Decimal;
var tab_Num     = Asc.c_oAscTabType.Num;
var tab_Right   = Asc.c_oAscTabType.Right;
var tab_Left    = Asc.c_oAscTabType.Left;

var tab_Symbol = 0x0022;//0x2192;

(function(window)
{

	// TODO: Реализовать табы по точке и с чертой (tab_Bar tab_Decimal)


	/**
	 * Класс представляющий элемент табуляции.
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunTab()
	{
		AscWord.CRunElementBase.call(this);

		this.Width        = 0;
		this.WidthVisible = 0;
		this.RealWidth    = 0;

		this.DotWidth        = 0;
		this.UnderscoreWidth = 0;
		this.HyphenWidth     = 0;
		this.Leader          = Asc.c_oAscTabLeader.None;
	}
	CRunTab.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunTab.prototype.constructor = CRunTab;

	CRunTab.prototype.Type = para_Tab;
	CRunTab.prototype.IsTab = function()
	{
		return true;
	};
	CRunTab.prototype.Draw = function(X, Y, Context)
	{
		if (this.WidthVisible > 0.01)
		{
			var sChar = null, nCharWidth = 0;
			switch (this.Leader)
			{
				case Asc.c_oAscTabLeader.Dot:
					sChar      = '.';
					nCharWidth = this.DotWidth;
					break;
				case Asc.c_oAscTabLeader.Heavy:
				case Asc.c_oAscTabLeader.Underscore:
					sChar      = '_';
					nCharWidth = this.UnderscoreWidth;
					break;
				case Asc.c_oAscTabLeader.Hyphen:
					sChar      = '-';
					nCharWidth = this.HyphenWidth;
					break;
				case Asc.c_oAscTabLeader.MiddleDot:
					sChar      = '·';
					nCharWidth = this.MiddleDotWidth;
					break;
			}

			if (null !== sChar && nCharWidth > 0.001)
			{
				Context.SetFontSlot(AscWord.fontslot_ASCII, 1);
				var nCharsCount = Math.floor(this.WidthVisible / nCharWidth);

				var _X = X + (this.WidthVisible - nCharsCount * nCharWidth) / 2;
				for (var nIndex = 0; nIndex < nCharsCount; ++nIndex, _X += nCharWidth)
					Context.FillText(_X, Y, sChar);
			}
		}

		if (editor && editor.ShowParaMarks)
		{
			Context.p_color(0, 0, 0, 255);
			Context.b_color1(0, 0, 0, 255);

			var X0 = this.Width / 2 - this.RealWidth / 2;

			Context.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});

			if (X0 > 0)
				Context.FillText2(X + X0, Y, String.fromCharCode(tab_Symbol), 0, this.Width);
			else
				Context.FillText2(X, Y, String.fromCharCode(tab_Symbol), this.RealWidth - this.Width, this.Width);
		}
	};
	CRunTab.prototype.Measure = function(Context)
	{
		Context.SetFontSlot(AscWord.fontslot_ASCII, 1);

		this.DotWidth        = Context.Measure(".").Width;
		this.UnderscoreWidth = Context.Measure("_").Width;
		this.HyphenWidth     = Context.Measure("-").Width * 1.5;
		this.MiddleDotWidth  = Context.Measure("·").Width;

		Context.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});
		this.RealWidth = Context.Measure(String.fromCharCode(tab_Symbol)).Width;
	};
	CRunTab.prototype.SetLeader = function(nLeaderType)
	{
		this.Leader = nLeaderType;
	};
	CRunTab.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CRunTab.prototype.GetWidthVisible = function()
	{
		return this.WidthVisible;
	};
	CRunTab.prototype.SetWidthVisible = function(WidthVisible)
	{
		this.WidthVisible = WidthVisible;
	};
	CRunTab.prototype.Copy = function()
	{
		return new CRunTab();
	};
	CRunTab.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunTab.prototype.SaveRecalculateObject = function()
	{
		return {
			Width        : this.Width,
			WidthVisible : this.WidthVisible
		};
	};
	CRunTab.prototype.LoadRecalculateObject = function(RecalcObj)
	{
		this.Width        = RecalcObj.Width;
		this.WidthVisible = RecalcObj.WidthVisible;
	};
	CRunTab.prototype.PrepareRecalculateObject = function()
	{
	};
	CRunTab.prototype.GetAutoCorrectFlags = function()
	{
		return AscWord.AUTOCORRECT_FLAGS_ALL;
	};
	CRunTab.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextSpecialTab();
	};
	CRunTab.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunTab = CRunTab;

})(window);
