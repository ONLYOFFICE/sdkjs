/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
/**
 * User: Ilja.Kirillov
 * Date: 08.05.2018
 * Time: 15:52
 */

// Import
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var History         = AscCommon.History;
var IntToNumberFormat = window["AscCommon"].IntToNumberFormat;


var numbering_presentationnumfrmt_AlphaLcParenBoth = 0;
var numbering_presentationnumfrmt_AlphaLcParenR = 1;
var numbering_presentationnumfrmt_AlphaLcPeriod = 2;
var numbering_presentationnumfrmt_AlphaUcParenBoth = 3;
var numbering_presentationnumfrmt_AlphaUcParenR = 4;
var numbering_presentationnumfrmt_AlphaUcPeriod = 5;
var numbering_presentationnumfrmt_Arabic1Minus = 6;
var numbering_presentationnumfrmt_Arabic2Minus = 7;
var numbering_presentationnumfrmt_ArabicDbPeriod = 8;
var numbering_presentationnumfrmt_ArabicDbPlain = 9;
var numbering_presentationnumfrmt_ArabicParenBoth = 10;
var numbering_presentationnumfrmt_ArabicParenR = 11;
var numbering_presentationnumfrmt_ArabicPeriod = 12;
var numbering_presentationnumfrmt_ArabicPlain = 13;
var numbering_presentationnumfrmt_CircleNumDbPlain = 14;
var numbering_presentationnumfrmt_CircleNumWdBlackPlain = 15;
var numbering_presentationnumfrmt_CircleNumWdWhitePlain = 16;
var numbering_presentationnumfrmt_Ea1ChsPeriod = 17;
var numbering_presentationnumfrmt_Ea1ChsPlain = 18;
var numbering_presentationnumfrmt_Ea1ChtPeriod = 19;
var numbering_presentationnumfrmt_Ea1ChtPlain = 20;
var numbering_presentationnumfrmt_Ea1JpnChsDbPeriod = 21;
var numbering_presentationnumfrmt_Ea1JpnKorPeriod = 22;
var numbering_presentationnumfrmt_Ea1JpnKorPlain = 23;
var numbering_presentationnumfrmt_Hebrew2Minus = 24;
var numbering_presentationnumfrmt_HindiAlpha1Period = 25;
var numbering_presentationnumfrmt_HindiAlphaPeriod = 26;
var numbering_presentationnumfrmt_HindiNumParenR = 27;
var numbering_presentationnumfrmt_HindiNumPeriod = 28;
var numbering_presentationnumfrmt_RomanLcParenBoth = 29;
var numbering_presentationnumfrmt_RomanLcParenR = 30;
var numbering_presentationnumfrmt_RomanLcPeriod = 31;
var numbering_presentationnumfrmt_RomanUcParenBoth = 32;
var numbering_presentationnumfrmt_RomanUcParenR = 33;
var numbering_presentationnumfrmt_RomanUcPeriod = 34;
var numbering_presentationnumfrmt_ThaiAlphaParenBoth = 35;
var numbering_presentationnumfrmt_ThaiAlphaParenR = 36;
var numbering_presentationnumfrmt_ThaiAlphaPeriod = 37;
var numbering_presentationnumfrmt_ThaiNumParenBoth = 38;
var numbering_presentationnumfrmt_ThaiNumParenR = 39;
var numbering_presentationnumfrmt_ThaiNumPeriod = 40;

var numbering_presentationnumfrmt_None          =     100;
var numbering_presentationnumfrmt_Char          =     101;

function IsAlphaPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_AlphaLcParenBoth &&
				nType <= numbering_presentationnumfrmt_AlphaUcPeriod;
	}
	return false;
}
function IsArabicPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_Arabic1Minus &&
			nType <= numbering_presentationnumfrmt_ArabicPlain;
	}
	return false;
}
function IsCirclePrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_CircleNumDbPlain &&
			nType <= numbering_presentationnumfrmt_CircleNumWdWhitePlain;
	}
	return false;
}
function IsEaPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_Ea1ChsPeriod &&
			nType <= numbering_presentationnumfrmt_Ea1JpnKorPlain;
	}
	return false;
}

function IsHebrewPrNumbering(nType)
{
	return nType === numbering_presentationnumfrmt_Hebrew2Minus;
}

function IsHindiPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_HindiAlpha1Period &&
			nType <= numbering_presentationnumfrmt_HindiNumPeriod;
	}
	return false;
}
function IsRomanPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_RomanLcParenBoth &&
			nType <= numbering_presentationnumfrmt_RomanUcPeriod;
	}
	return false;
}
function IsThaiPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= numbering_presentationnumfrmt_ThaiAlphaParenBoth &&
			nType <= numbering_presentationnumfrmt_ThaiNumPeriod;
	}
	return false;
}

function IsPrNumberingSameType(nType1, nType2)
{
	return IsAlphaPrNumbering(nType1) && IsAlphaPrNumbering(nType2) ||
		IsArabicPrNumbering(nType1) && IsArabicPrNumbering(nType2) ||
		IsCirclePrNumbering(nType1) && IsCirclePrNumbering(nType2) ||
		IsEaPrNumbering(nType1) && IsEaPrNumbering(nType2) ||
		IsHebrewPrNumbering(nType1) && IsHebrewPrNumbering(nType2) ||
		IsHindiPrNumbering(nType1) && IsHindiPrNumbering(nType2) ||
		IsRomanPrNumbering(nType1) && IsRomanPrNumbering(nType2) ||
		IsThaiPrNumbering(nType1) && IsThaiPrNumbering(nType2);
}

// Класс для работы с нумерацией в презентациях
function CPresentationBullet()
{
	this.m_nType    = numbering_presentationnumfrmt_None;  // Тип
	this.m_nStartAt = null;                                // Стартовое значение для нумерованных списков
	this.m_sChar    = null;                                // Значение для символьных списков

	this.m_oColor   = { r : 0, g : 0, b : 0, a: 255 };     // Цвет
	this.m_bColorTx = true;                                // Использовать ли цвет первого рана в параграфе
	this.Unifill    = null;

	this.m_sFont    = "Arial";                             // Шрифт
	this.m_bFontTx  = true;                                // Использовать ли шрифт первого рана в параграфе

	this.m_dSize    = 1;                                   // Размер шрифта, в пунктах или в процентах (зависит от флага m_bSizePct)
	this.m_bSizeTx  = false;                               // Использовать ли размер шрифта первого рана в параграфе
	this.m_bSizePct = true;                                // Задан ли размер шрифта в процентах

	this.m_oTextPr = null;
	this.m_nNum    = null;
	this.m_sString = null;
}

CPresentationBullet.prototype.getHighlightForNumbering = function(intFormat) {
	switch (this.m_nType) {
		case numbering_presentationnumfrmt_AlphaLcParenBoth:
		case numbering_presentationnumfrmt_AlphaUcParenBoth:
		case numbering_presentationnumfrmt_ArabicParenBoth:
		case numbering_presentationnumfrmt_RomanLcParenBoth:
		case numbering_presentationnumfrmt_RomanUcParenBoth:
		case numbering_presentationnumfrmt_ThaiAlphaParenBoth:
		case numbering_presentationnumfrmt_ThaiNumParenBoth:
			return '(' + intFormat + ')';
		case numbering_presentationnumfrmt_AlphaLcParenR:
		case numbering_presentationnumfrmt_AlphaUcParenR:
		case numbering_presentationnumfrmt_ArabicParenR:
		case numbering_presentationnumfrmt_HindiNumParenR:
		case numbering_presentationnumfrmt_RomanLcParenR:
		case numbering_presentationnumfrmt_RomanUcParenR:
		case numbering_presentationnumfrmt_ThaiAlphaParenR:
		case numbering_presentationnumfrmt_ThaiNumParenR:
			return '' + intFormat + ')';
		case numbering_presentationnumfrmt_AlphaLcPeriod:
		case numbering_presentationnumfrmt_AlphaUcPeriod:
		case numbering_presentationnumfrmt_ArabicDbPeriod:
		case numbering_presentationnumfrmt_ArabicPeriod:
		case numbering_presentationnumfrmt_Ea1ChsPeriod:
		case numbering_presentationnumfrmt_Ea1ChtPeriod:
		case numbering_presentationnumfrmt_Ea1JpnChsDbPeriod:
		case numbering_presentationnumfrmt_Ea1JpnKorPeriod:
		case numbering_presentationnumfrmt_HindiAlpha1Period:
		case numbering_presentationnumfrmt_HindiAlphaPeriod:
		case numbering_presentationnumfrmt_HindiNumPeriod:
		case numbering_presentationnumfrmt_RomanLcPeriod:
		case numbering_presentationnumfrmt_RomanUcPeriod:
		case numbering_presentationnumfrmt_ThaiAlphaPeriod:
		case numbering_presentationnumfrmt_ThaiNumPeriod:
			return '' + intFormat + '.';
		case numbering_presentationnumfrmt_Arabic1Minus:
		case numbering_presentationnumfrmt_Arabic2Minus:
		case numbering_presentationnumfrmt_Hebrew2Minus:
			return '' + intFormat + '-';
		case numbering_presentationnumfrmt_ArabicDbPlain:
		case numbering_presentationnumfrmt_ArabicPlain:
		case numbering_presentationnumfrmt_CircleNumDbPlain:
		case numbering_presentationnumfrmt_CircleNumWdBlackPlain:
		case numbering_presentationnumfrmt_CircleNumWdWhitePlain:
		case numbering_presentationnumfrmt_Ea1ChsPlain:
		case numbering_presentationnumfrmt_Ea1ChtPlain:
		case numbering_presentationnumfrmt_Ea1JpnKorPlain:
		case numbering_presentationnumfrmt_Char:
		case numbering_presentationnumfrmt_None:
		default:
			return '' + intFormat;
	}
}

function getAdaptedNumberingFormat(nType) {
	switch (nType) {
		case numbering_presentationnumfrmt_AlphaLcParenBoth:
		case numbering_presentationnumfrmt_AlphaLcParenR:
		case numbering_presentationnumfrmt_AlphaLcPeriod:
			return Asc.c_oAscNumberingFormat.LowerLetter;

		case numbering_presentationnumfrmt_AlphaUcParenBoth:
		case numbering_presentationnumfrmt_AlphaUcParenR:
		case numbering_presentationnumfrmt_AlphaUcPeriod:
			return Asc.c_oAscNumberingFormat.UpperLetter;

		case numbering_presentationnumfrmt_Arabic1Minus:
			return Asc.c_oAscNumberingFormat.ArabicAlpha;

		case numbering_presentationnumfrmt_Arabic2Minus:
			return Asc.c_oAscNumberingFormat.ArabicAbjad;

		case numbering_presentationnumfrmt_ArabicDbPeriod:
		case numbering_presentationnumfrmt_ArabicDbPlain:
			return Asc.c_oAscNumberingFormat.DecimalFullWidth;

		case numbering_presentationnumfrmt_ArabicParenBoth:
		case numbering_presentationnumfrmt_ArabicParenR:
		case numbering_presentationnumfrmt_ArabicPeriod:
		case numbering_presentationnumfrmt_ArabicPlain:
			return Asc.c_oAscNumberingFormat.Decimal;

		case numbering_presentationnumfrmt_CircleNumDbPlain:
			return Asc.c_oAscNumberingFormat.DecimalEnclosedCircle;

		case numbering_presentationnumfrmt_Ea1ChsPeriod:
		case numbering_presentationnumfrmt_Ea1ChsPlain:
			return Asc.c_oAscNumberingFormat.ChineseCounting;

		case numbering_presentationnumfrmt_Hebrew2Minus:
			return Asc.c_oAscNumberingFormat.Hebrew2;

		case numbering_presentationnumfrmt_HindiAlpha1Period:
			return Asc.c_oAscNumberingFormat.HindiConsonants;

		case numbering_presentationnumfrmt_HindiAlphaPeriod:
			return Asc.c_oAscNumberingFormat.HindiVowels;

		case numbering_presentationnumfrmt_HindiNumParenR:
		case numbering_presentationnumfrmt_HindiNumPeriod:
			return Asc.c_oAscNumberingFormat.HindiNumbers;

		case numbering_presentationnumfrmt_RomanLcParenBoth:
		case numbering_presentationnumfrmt_RomanLcParenR:
		case numbering_presentationnumfrmt_RomanLcPeriod:
			return Asc.c_oAscNumberingFormat.LowerRoman;

		case numbering_presentationnumfrmt_RomanUcParenBoth:
		case numbering_presentationnumfrmt_RomanUcParenR:
		case numbering_presentationnumfrmt_RomanUcPeriod:
			return Asc.c_oAscNumberingFormat.UpperRoman;

		case numbering_presentationnumfrmt_ThaiAlphaParenBoth:
		case numbering_presentationnumfrmt_ThaiAlphaParenR:
		case numbering_presentationnumfrmt_ThaiAlphaPeriod:
			return Asc.c_oAscNumberingFormat.ThaiLetters;

		case numbering_presentationnumfrmt_ThaiNumParenBoth:
		case numbering_presentationnumfrmt_ThaiNumParenR:
		case numbering_presentationnumfrmt_ThaiNumPeriod:
			return Asc.c_oAscNumberingFormat.ThaiNumbers;

		case numbering_presentationnumfrmt_Ea1ChtPeriod:
		case numbering_presentationnumfrmt_Ea1ChtPlain:
			break;

		case numbering_presentationnumfrmt_Ea1JpnChsDbPeriod:
			break;

		case numbering_presentationnumfrmt_CircleNumWdBlackPlain:
			break;

		case numbering_presentationnumfrmt_CircleNumWdWhitePlain: // TODO: new break type
			break;

		case numbering_presentationnumfrmt_Ea1JpnKorPeriod:
		case numbering_presentationnumfrmt_Ea1JpnKorPlain:
			break;

		case numbering_presentationnumfrmt_None:
		default:
			return Asc.c_oAscNumberingFormat.None;
	}
}

CPresentationBullet.prototype.Get_Type = function()
{
	return this.m_nType;
};
CPresentationBullet.prototype.Get_StartAt = function()
{
	return this.m_nStartAt;
};
CPresentationBullet.prototype.Measure = function(Context, FirstTextPr, Num, Theme, ColorMap)
{
	var sT = "";
	if (this.m_nType === numbering_presentationnumfrmt_Char) {
		if ( null != this.m_sChar ) {
			sT = this.m_sChar;
		}
	} else {
		var typeOfNum = getAdaptedNumberingFormat(this.m_nType);
		var formatNum = IntToNumberFormat(Num, typeOfNum);
		sT = this.getHighlightForNumbering(formatNum);
	}

	this.m_sString = sT;
	this.m_nNum = Num;
	if(sT.length === 0)
	{
		return { Width : 0 };
	}
	var dFontSize = FirstTextPr.FontSize;
	if ( false === this.m_bSizeTx )
	{
		if ( true === this.m_bSizePct )
			dFontSize *= this.m_dSize;
		else
			dFontSize = this.m_dSize;
	}

	var RFonts;
	if(!this.m_bFontTx)
	{
		RFonts = {
			Ascii: {
				Name: this.m_sFont,
				Index: -1
			},
			EastAsia: {
				Name: this.m_sFont,
				Index: -1
			},
			CS: {
				Name: this.m_sFont,
				Index: -1
			},
			HAnsi: {
				Name: this.m_sFont,
				Index: -1
			}
		};
	}
	else
	{
		RFonts = FirstTextPr.RFonts;
	}

	var FirstTextPr_ = FirstTextPr.Copy();
	if(FirstTextPr_.Underline)
	{
		FirstTextPr_.Underline = false;
	}

	if ( true === this.m_bColorTx || !this.Unifill)
	{
		if(FirstTextPr.Unifill)
		{
			this.Unifill = FirstTextPr_.Unifill;
		}
		else
		{
			this.Unifill = AscFormat.CreteSolidFillRGB(FirstTextPr.Color.r, FirstTextPr.Color.g, FirstTextPr.Color.b);
		}
	}

	var TextPr_ = new CTextPr();
	var bIsNumbered = this.IsNumbered();
	TextPr_.Set_FromObject({
		RFonts: RFonts,
		Unifill: this.Unifill,
		FontSize : dFontSize,
		Bold     : ( bIsNumbered ? FirstTextPr.Bold   : false ),
		Italic   : ( bIsNumbered ? FirstTextPr.Italic : false )
	});
	FirstTextPr_.Merge(TextPr_);
	this.m_oTextPr = FirstTextPr_;


	var X = 0;
	var OldTextPr = Context.GetTextPr();
	var Hint =  this.m_oTextPr.RFonts.Hint;
	var bCS  =  this.m_oTextPr.CS;
	var bRTL =  this.m_oTextPr.RTL;
	var lcid =  this.m_oTextPr.Lang.EastAsia;

	var FontSlot = g_font_detector.Get_FontClass( sT.getUnicodeIterator().value(), Hint, lcid, bCS, bRTL );
	Context.SetTextPr( this.m_oTextPr, Theme );
	Context.SetFontSlot( FontSlot );
	for (var iter = sT.getUnicodeIterator(); iter.check(); iter.next()) {
		var charCode = iter.value();
		X += Context.MeasureCode(charCode).Width;
	}

	if(OldTextPr)
	{
		Context.SetTextPr( OldTextPr, Theme );
	}
	return { Width : X };
};
CPresentationBullet.prototype.Copy = function()
{
	var Bullet = new CPresentationBullet();

	Bullet.m_nType    = this.m_nType;
	Bullet.m_nStartAt = this.m_nStartAt;
	Bullet.m_sChar    = this.m_sChar;

	Bullet.m_oColor.r = this.m_oColor.r;
	Bullet.m_oColor.g = this.m_oColor.g;
	Bullet.m_oColor.b = this.m_oColor.b;
	Bullet.m_bColorTx = this.m_bColorTx;

	Bullet.m_sFont    = this.m_sFont;
	Bullet.m_bFontTx  = this.m_bFontTx;

	Bullet.m_dSize    = this.m_dSize;
	Bullet.m_bSizeTx  = this.m_bSizeTx;
	Bullet.m_bSizePct = this.m_bSizePct;

	return Bullet;
};
CPresentationBullet.prototype.Draw = function(X, Y, Context, PDSE)
{
	if ( null === this.m_oTextPr || null === this.m_nNum || null == this.m_sString || this.m_sString.length == 0)
		return;



	var OldTextPr  = Context.GetTextPr();
	var OldTextPr2 = g_oTextMeasurer.GetTextPr();

	var Hint =  this.m_oTextPr.RFonts.Hint;
	var bCS  =  this.m_oTextPr.CS;
	var bRTL =  this.m_oTextPr.RTL;
	var lcid =  this.m_oTextPr.Lang.EastAsia;

	var sT = this.m_sString;
	var FontSlot = g_font_detector.Get_FontClass( sT.getUnicodeIterator().value(), Hint, lcid, bCS, bRTL );

	if(this.m_oTextPr.Unifill){
		this.m_oTextPr.Unifill.check(PDSE.Theme, PDSE.ColorMap);
	}
	Context.SetTextPr( this.m_oTextPr, PDSE.Theme );
	Context.SetFontSlot( FontSlot );
	if(!Context.Start_Command){
		if(this.m_oTextPr.Unifill){
			var RGBA = this.m_oTextPr.Unifill.getRGBAColor();
			this.m_oColor.r = RGBA.R;
			this.m_oColor.g = RGBA.G;
			this.m_oColor.b = RGBA.B;
		}
		var r = this.m_oColor.r;
		var g = this.m_oColor.g;
		var b = this.m_oColor.b;
		if(PDSE.Paragraph && PDSE.Paragraph.IsEmpty())
		{
			var dAlpha = 0.4;
			var rB, gB, bB;
			if(PDSE.BgColor)
			{
				rB = PDSE.BgColor.r;
				gB = PDSE.BgColor.g;
				bB = PDSE.BgColor.b;
			}
			else
			{
				rB = 255;
				gB = 255;
				bB = 255;
			}
			r = Math.min(255, (r  * dAlpha + rB * (1 - dAlpha) + 0.5) >> 0);
			g = Math.min(255, (g  * dAlpha + gB * (1 - dAlpha) + 0.5) >> 0);
			b = Math.min(255, (b  * dAlpha + bB * (1 - dAlpha) + 0.5) >> 0);
		}
		Context.p_color(r, g, b, 255 );
		Context.b_color1(r, g, b, 255 );
	}
	g_oTextMeasurer.SetTextPr( this.m_oTextPr, PDSE.Theme  );
	g_oTextMeasurer.SetFontSlot( FontSlot );

	for (var iter = sT.getUnicodeIterator(); iter.check(); iter.next()) {
		var charCode = iter.value();
		Context.FillTextCode( X, Y, charCode );
		X += g_oTextMeasurer.MeasureCode(charCode).Width;
	}

	if(OldTextPr)
	{
		Context.SetTextPr( OldTextPr, PDSE.Theme );
	}
	if(OldTextPr2)
	{
		g_oTextMeasurer.SetTextPr( OldTextPr2, PDSE.Theme  );
	}
};
CPresentationBullet.prototype.IsNumbered = function()
{
	return this.m_nType >= numbering_presentationnumfrmt_AlphaLcParenBoth
		&& this.m_nType <= numbering_presentationnumfrmt_ThaiNumPeriod;
};
CPresentationBullet.prototype.IsNone = function()
{
	return this.m_nType === numbering_presentationnumfrmt_None;
};
CPresentationBullet.prototype.IsAlpha = function()
{
	return IsAlphaPrNumbering(this.m_nType);
};


//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].getAdaptedNumberingFormat = getAdaptedNumberingFormat;
