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

// Import
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var History         = AscCommon.History;
var IntToNumberFormat = window["AscCommon"].IntToNumberFormat;

var kHeightImageBullet = 1 / 4.043478260869565;



function IsAlphaPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth &&
				nType <= AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod;
	}
	return false;
}
function IsArabicPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_Arabic1Minus &&
			nType <= AscFormat.numbering_presentationnumfrmt_ArabicPlain;
	}
	return false;
}
function IsCirclePrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain &&
			nType <= AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain;
	}
	return false;
}
function IsEaPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod &&
			nType <= AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain;
	}
	return false;
}

function IsHebrewPrNumbering(nType)
{
	return nType === AscFormat.numbering_presentationnumfrmt_Hebrew2Minus;
}

function IsHindiPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period &&
			nType <= AscFormat.numbering_presentationnumfrmt_HindiNumPeriod;
	}
	return false;
}
function IsRomanPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth &&
			nType <= AscFormat.numbering_presentationnumfrmt_RomanUcPeriod;
	}
	return false;
}
function IsThaiPrNumbering(nType)
{
	if(AscFormat.isRealNumber(nType))
	{
		return nType >= AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth &&
			nType <= AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod;
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
	this.m_nType    = AscFormat.numbering_presentationnumfrmt_None;  // Тип
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

	this.m_sSrc = null;
}

CPresentationBullet.prototype.convertFromAscTypeToPresentation = function (nType) {
	switch (nType) {
			case c_oAscNumberingLevel.DecimalBracket_Right    :
			case c_oAscNumberingLevel.DecimalBracket_Left     :
				return AscFormat.numbering_presentationnumfrmt_ArabicParenR;
			case c_oAscNumberingLevel.DecimalDot_Right        :
			case c_oAscNumberingLevel.DecimalDot_Left         :
				return AscFormat.numbering_presentationnumfrmt_ArabicPeriod;
			case c_oAscNumberingLevel.UpperRomanDot_Right     :
				return AscFormat.numbering_presentationnumfrmt_RomanUcPeriod;
			case c_oAscNumberingLevel.UpperLetterDot_Left     :
				return AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod;
			case c_oAscNumberingLevel.LowerLetterBracket_Left :
				return AscFormat.numbering_presentationnumfrmt_AlphaLcParenR;
			case c_oAscNumberingLevel.LowerLetterDot_Left     :
				return AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod;
			case c_oAscNumberingLevel.LowerRomanDot_Right     :
				return AscFormat.numbering_presentationnumfrmt_RomanLcPeriod;
			case c_oAscNumberingLevel.UpperRomanBracket_Left  :
				return AscFormat.numbering_presentationnumfrmt_RomanUcParenR;
			case c_oAscNumberingLevel.LowerRomanBracket_Left  :
				return AscFormat.numbering_presentationnumfrmt_RomanLcParenR;
			case c_oAscNumberingLevel.UpperLetterBracket_Left :
				return AscFormat.numbering_presentationnumfrmt_AlphaUcParenR;
			default:
				break;
	}
};

CPresentationBullet.prototype.getHighlightForNumbering = function(intFormat) {
	switch (this.m_nType) {
		case AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ArabicParenBoth:
		case AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ThaiNumParenBoth:
			return '(' + intFormat + ')';
		case AscFormat.numbering_presentationnumfrmt_AlphaLcParenR:
		case AscFormat.numbering_presentationnumfrmt_AlphaUcParenR:
		case AscFormat.numbering_presentationnumfrmt_ArabicParenR:
		case AscFormat.numbering_presentationnumfrmt_HindiNumParenR:
		case AscFormat.numbering_presentationnumfrmt_RomanLcParenR:
		case AscFormat.numbering_presentationnumfrmt_RomanUcParenR:
		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenR:
		case AscFormat.numbering_presentationnumfrmt_ThaiNumParenR:
			return '' + intFormat + ')';
		case AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod:
		case AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod:
		case AscFormat.numbering_presentationnumfrmt_ArabicDbPeriod:
		case AscFormat.numbering_presentationnumfrmt_ArabicPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChtPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1JpnChsDbPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPeriod:
		case AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period:
		case AscFormat.numbering_presentationnumfrmt_HindiAlphaPeriod:
		case AscFormat.numbering_presentationnumfrmt_HindiNumPeriod:
		case AscFormat.numbering_presentationnumfrmt_RomanLcPeriod:
		case AscFormat.numbering_presentationnumfrmt_RomanUcPeriod:
		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaPeriod:
		case AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod:
			return '' + intFormat + '.';
		case AscFormat.numbering_presentationnumfrmt_Arabic1Minus:
		case AscFormat.numbering_presentationnumfrmt_Arabic2Minus:
		case AscFormat.numbering_presentationnumfrmt_Hebrew2Minus:
			return '' + intFormat + '-';
		case AscFormat.numbering_presentationnumfrmt_ArabicDbPlain:
		case AscFormat.numbering_presentationnumfrmt_ArabicPlain:
		case AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain:
		case AscFormat.numbering_presentationnumfrmt_CircleNumWdBlackPlain:
		case AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChsPlain:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChtPlain:
		case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain:
		case AscFormat.numbering_presentationnumfrmt_Char:
		case AscFormat.numbering_presentationnumfrmt_None:
		default:
			return '' + intFormat;
	}
}

function getAdaptedNumberingFormat(nType) {
	switch (nType) {
		case AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_AlphaLcParenR:
		case AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod:
			return Asc.c_oAscNumberingFormat.LowerLetter;

		case AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_AlphaUcParenR:
		case AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod:
			return Asc.c_oAscNumberingFormat.UpperLetter;

		case AscFormat.numbering_presentationnumfrmt_Arabic1Minus:
			return Asc.c_oAscNumberingFormat.ArabicAlpha;

		case AscFormat.numbering_presentationnumfrmt_Arabic2Minus:
			return Asc.c_oAscNumberingFormat.ArabicAbjad;

		case AscFormat.numbering_presentationnumfrmt_ArabicDbPeriod:
		case AscFormat.numbering_presentationnumfrmt_ArabicDbPlain:
			return Asc.c_oAscNumberingFormat.DecimalFullWidth;

		case AscFormat.numbering_presentationnumfrmt_ArabicParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ArabicParenR:
		case AscFormat.numbering_presentationnumfrmt_ArabicPeriod:
		case AscFormat.numbering_presentationnumfrmt_ArabicPlain:
			return Asc.c_oAscNumberingFormat.Decimal;

		case AscFormat.numbering_presentationnumfrmt_CircleNumDbPlain:
			return Asc.c_oAscNumberingFormat.DecimalEnclosedCircle;

		case AscFormat.numbering_presentationnumfrmt_Ea1ChsPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChsPlain:
			return Asc.c_oAscNumberingFormat.ChineseCounting;

		case AscFormat.numbering_presentationnumfrmt_Hebrew2Minus:
			return Asc.c_oAscNumberingFormat.Hebrew2;

		case AscFormat.numbering_presentationnumfrmt_HindiAlpha1Period:
			return Asc.c_oAscNumberingFormat.HindiConsonants;

		case AscFormat.numbering_presentationnumfrmt_HindiAlphaPeriod:
			return Asc.c_oAscNumberingFormat.HindiVowels;

		case AscFormat.numbering_presentationnumfrmt_HindiNumParenR:
		case AscFormat.numbering_presentationnumfrmt_HindiNumPeriod:
			return Asc.c_oAscNumberingFormat.HindiNumbers;

		case AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_RomanLcParenR:
		case AscFormat.numbering_presentationnumfrmt_RomanLcPeriod:
			return Asc.c_oAscNumberingFormat.LowerRoman;

		case AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth:
		case AscFormat.numbering_presentationnumfrmt_RomanUcParenR:
		case AscFormat.numbering_presentationnumfrmt_RomanUcPeriod:
			return Asc.c_oAscNumberingFormat.UpperRoman;

		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaParenR:
		case AscFormat.numbering_presentationnumfrmt_ThaiAlphaPeriod:
			return Asc.c_oAscNumberingFormat.ThaiLetters;

		case AscFormat.numbering_presentationnumfrmt_ThaiNumParenBoth:
		case AscFormat.numbering_presentationnumfrmt_ThaiNumParenR:
		case AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod:
			return Asc.c_oAscNumberingFormat.ThaiNumbers;

		case AscFormat.numbering_presentationnumfrmt_Ea1ChtPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1ChtPlain:
			break;

		case AscFormat.numbering_presentationnumfrmt_Ea1JpnChsDbPeriod:
			break;

		case AscFormat.numbering_presentationnumfrmt_CircleNumWdBlackPlain:
			break;

		case AscFormat.numbering_presentationnumfrmt_CircleNumWdWhitePlain: // TODO: new break type
			break;

		case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPeriod:
		case AscFormat.numbering_presentationnumfrmt_Ea1JpnKorPlain:
			break;

		case AscFormat.numbering_presentationnumfrmt_None:
		default:
			return Asc.c_oAscNumberingFormat.None;
	}
}

CPresentationBullet.prototype.GetDrawingContent = function (arrLvls, nLvl, nNum)
{
	const oApi = Asc.editor || editor;
	if (this.m_sSrc)
	{
		const oImage = oApi.ImageLoader.map_image_index[this.m_sSrc];
		return oImage ? {image: oImage, amount: 1} : {amount: 0};
	}
	else
	{
		return this.GetDrawingText(nNum);
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
CPresentationBullet.prototype.GetIndentSize = function ()
{
	return 0;
};
CPresentationBullet.prototype.GetNumberPosition = function ()
{
	return 0;
};

CPresentationBullet.prototype.GetDrawingText = function () {
	let nNum;
	if (arguments.length === 1) {
		nNum = arguments[0];
	} else if (arguments.length === 3) {
		nNum = arguments[2];
	} else {
		nNum = 1;
	}
	var sT = "";
	if (this.m_nType === AscFormat.numbering_presentationnumfrmt_Char)
	{
		if ( null != this.m_sChar )
		{
			sT = this.m_sChar;
		}
	} else if (this.m_nType !== AscFormat.numbering_presentationnumfrmt_Blip)
	{
		var nTypeOfNum = getAdaptedNumberingFormat(this.m_nType);
		var nFormatNum = IntToNumberFormat(nNum, nTypeOfNum);
		sT = this.getHighlightForNumbering(nFormatNum);
	}
	return sT;
}

CPresentationBullet.prototype.MergeTextPr = function (FirstTextPr)
{
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
			this.Unifill = AscFormat.CreateSolidFillRGB(FirstTextPr.Color.r, FirstTextPr.Color.g, FirstTextPr.Color.b);
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
}

CPresentationBullet.prototype.Measure = function(Context, FirstTextPr, Num, Theme)
{
	this.m_nNum = Num;
	this.m_sString = this.GetDrawingText(Num);
	this.MergeTextPr(FirstTextPr);
	if (this.m_nType === AscFormat.numbering_presentationnumfrmt_Blip)
	{
		var sizes = AscCommon.getSourceImageSize(this.m_sSrc);
		var x_height = this.m_oTextPr.FontSize * kHeightImageBullet;
		var adaptImageHeight = x_height;
		var adaptImageWidth = sizes.width * adaptImageHeight / (sizes.height ? sizes.height : 1);
		return { Width: adaptImageWidth };
	}

	var X = 0;
	var OldTextPr = Context.GetTextPr();
	var FontSlot;
	var sT = this.m_sString;
	if (sT)
	{
		FontSlot = AscWord.GetFontSlotByTextPr( sT.getUnicodeIterator().value(), this.m_oTextPr);
	}
	Context.SetTextPr( this.m_oTextPr, Theme );
	Context.SetFontSlot( FontSlot );

	if(sT.length === 0)
	{
		return { Width : 0 };
	}
	for (var iter = sT.getUnicodeIterator(); iter.check(); iter.next())
	{
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
	Bullet.m_sSrc     = this.m_sSrc;

	return Bullet;
};

CPresentationBullet.prototype.IsErrorInNumeration = function()
{
	return ((null === this.m_oTextPr
		|| null === this.m_nNum
		|| null == this.m_sString
		|| this.m_sString.length === 0)
		&& this.m_sSrc === null);
}

CPresentationBullet.prototype.Draw = function(X, Y, Context, PDSE)
{
	if (this.IsErrorInNumeration())
		return;

	if (this.m_nType === AscFormat.numbering_presentationnumfrmt_Blip)
	{
		var sizes = AscCommon.getSourceImageSize(this.m_sSrc);
		var x_height = this.m_oTextPr.FontSize * kHeightImageBullet;
		var adaptImageHeight = x_height;
		var adaptImageWidth = sizes.width * adaptImageHeight / (sizes.height ? sizes.height : 1);
		Context.drawImage(this.m_sSrc, X, Y - adaptImageHeight, adaptImageWidth, adaptImageHeight);
		return;
	}

	var OldTextPr  = Context.GetTextPr();
	var OldTextPr2 = g_oTextMeasurer.GetTextPr();

	var sT = this.m_sString;
	var FontSlot;
	if (sT)
	{
		FontSlot = AscWord.GetFontSlotByTextPr( sT.getUnicodeIterator().value(), this.m_oTextPr);
	}

	if(this.m_oTextPr.Unifill){
		this.m_oTextPr.Unifill.check(PDSE.Theme, PDSE.ColorMap);
	}
	Context.SetTextPr( this.m_oTextPr, PDSE.Theme );
	Context.SetFontSlot( FontSlot );
	if(!Context.isTextDrawer()){
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

	for (var iter = sT.getUnicodeIterator(); iter.check(); iter.next())
	{
		var charCode = iter.value();
		if (Context.m_bIsTextDrawer === true)
		{
			Context.CheckAddNewPath(X, Y, charCode);
		}
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
	return this.m_nType >= AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth
		&& this.m_nType <= AscFormat.numbering_presentationnumfrmt_ThaiNumPeriod;
};
CPresentationBullet.prototype.GetStringByLvlText = CPresentationBullet.prototype.GetDrawingText;
CPresentationBullet.prototype.GetTextPr = function ()
{
	if (!this.m_oTextPr)
		this.m_oTextPr = new AscCommonWord.CTextPr();
	return this.m_oTextPr;
};
CPresentationBullet.prototype.SetTextPr = function (oTextPr)
{
	this.m_oTextPr = oTextPr;
};
CPresentationBullet.prototype.SetJc = function (nJc) {};
CPresentationBullet.prototype.GetJc = function () {return AscCommon.align_Left};
CPresentationBullet.prototype.GetSymbols = function () {};
CPresentationBullet.prototype.GetSuff = function () {};

CPresentationBullet.prototype.IsNone = function()
{
	return this.m_nType === AscFormat.numbering_presentationnumfrmt_None;
};
CPresentationBullet.prototype.IsAlpha = function()
{
	return IsAlphaPrNumbering(this.m_nType);
};


//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].getAdaptedNumberingFormat = getAdaptedNumberingFormat;
window['AscCommonWord'].CPresentationBullet = CPresentationBullet;
