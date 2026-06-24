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

var tab_Bar     = Asc.c_oAscTabType.Bar;
var tab_Center  = Asc.c_oAscTabType.Center;
var tab_Clear   = Asc.c_oAscTabType.Clear;
var tab_Decimal = Asc.c_oAscTabType.Decimal;
var tab_Num     = Asc.c_oAscTabType.Num;
var tab_Right   = Asc.c_oAscTabType.Right;
var tab_Left    = Asc.c_oAscTabType.Left;

let tab_Symbol     = 0x0022;//0x2192;
let tab_Symbol_Rtl = 0x0021;

(function(window)
{
	let tabWidth = null;
	
	function getTabActualWidth(tabCode)
	{
		if (null !== tabWidth)
			return tabWidth;
		
		g_oTextMeasurer.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});
		tabWidth = g_oTextMeasurer.Measure(String.fromCharCode(tabCode)).Width;
		return tabWidth;
	}

	// TODO: Реализовать табы по точке и с чертой (tab_Bar tab_Decimal)


	/**
	 * Класс представляющий элемент табуляции.
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunTab()
	{
		AscWord.CRunElementBase.call(this);
		
		this.Width       = 0;
		this.LeaderCode  = 0x00;
		this.LeaderWidth = 0;
	}
	CRunTab.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunTab.prototype.constructor = CRunTab;

	CRunTab.prototype.Type = para_Tab;
	CRunTab.prototype.IsTab = function()
	{
		return true;
	};
	CRunTab.prototype.Draw = function(X, Y, Context, drawState)
	{
		if (this.Width > 0.01 && 0 !== this.LeaderCode)
		{
			Context.SetFontSlot(AscWord.fontslot_ASCII, 1);
			var nCharsCount = Math.floor(this.Width / this.LeaderWidth);
			
			var _X = X + (this.Width - nCharsCount * this.LeaderWidth) / 2;
			for (var nIndex = 0; nIndex < nCharsCount; ++nIndex, _X += this.LeaderWidth)
				Context.FillTextCode(_X, Y, this.LeaderCode);
		}

		if(Context.m_bIsTextDrawer)
		{
			Context.CheckSpaceDraw(this);
		}
		
		if (editor && editor.ShowParaMarks)
		{
			let isRtl = drawState.isRtlMainDirection();
			
			Context.p_color(0, 0, 0, 255);
			Context.b_color1(0, 0, 0, 255);
			
			let tabCode = drawState.isRtlMainDirection() ? tab_Symbol_Rtl : tab_Symbol;
			let tabActualWidth = getTabActualWidth(tabCode);
			var X0 = this.Width / 2 - tabActualWidth / 2;

			Context.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});

			if (X0 > 0)
				Context.FillText2(X + X0, Y, String.fromCharCode(tabCode), 0, this.Width);
			else
				Context.FillText2(X, Y, String.fromCharCode(tabCode), isRtl ? 0 : tabActualWidth - this.Width, this.Width);
		}
	};
	CRunTab.prototype.Measure = function(Context)
	{
	};
	CRunTab.prototype.SetLeader = function(leaderType, textPr)
	{
		g_oTextMeasurer.SetTextPr(textPr);
		g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, 1);
		
		switch (leaderType)
		{
			case Asc.c_oAscTabLeader.Dot: // "."
				this.LeaderCode  = 0x002E;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x002E).Width;
				break;
			case Asc.c_oAscTabLeader.Heavy: // "_"
			case Asc.c_oAscTabLeader.Underscore:
				this.LeaderCode  = 0x005F;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x005F).Width;
				break;
			case Asc.c_oAscTabLeader.Hyphen: // "-"
				this.LeaderCode  = 0x002D;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x002D).Width * 1.5;
				break;
			case Asc.c_oAscTabLeader.MiddleDot: // "·"
				this.LeaderCode  = 0x00B7;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x00B7).Width;
				break;
			default:
				this.LeaderCode  = 0x00;
				this.LeaderWidth = 0;
				break;
		}
	};
	CRunTab.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CRunTab.prototype.GetWidthVisible = function()
	{
		return this.Width;
	};
	CRunTab.prototype.SetWidthVisible = function(WidthVisible)
	{
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
			LeaderCode   : this.LeaderCode,
			LeaderWidth  : this.LeaderWidth
		};
	};
	CRunTab.prototype.LoadRecalculateObject = function(RecalcObj)
	{
		this.Width        = RecalcObj.Width;
		this.LeaderCode   = RecalcObj.LeaderCode;
		this.LeaderWidth  = RecalcObj.LeaderWidth;
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
	CRunTab.prototype.getBidiType = function()
	{
		return AscBidi.TYPE.PM;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunTab = CRunTab;

})(window);
