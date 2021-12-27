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
 * Time: 17:06
 */

function CNumberingLvl()
{
	this.Jc      = AscCommon.align_Left;
	this.Format  = Asc.c_oAscNumberingFormat.Bullet;
	this.PStyle  = undefined;
	this.Start   = 1;
	this.Restart = -1; // -1 - делаем нумерацию сначала всегда, 0 - никогда не начинаем нумерацию заново
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.TextPr  = new CTextPr();
	this.ParaPr  = new CParaPr();
	this.LvlText = [];
	this.Legacy  = undefined;
	this.IsLgl   = false;

	this.private_CheckSymbols();
}
/**
 * Доступ к типу прилегания данного уровня
 * @returns {AscCommon.align_Left | AscCommon.align_Right | AscCommon.align_Center}
 */
CNumberingLvl.prototype.GetJc = function()
{
	return this.Jc;
};
/**
 * Доступ к типу данного уровня
 * @returns {Asc.c_oAscNumberingFormat}
 */
CNumberingLvl.prototype.GetFormat = function()
{
	return this.Format;
};
/**
 * Доступ к связанному стилю
 * @returns {?string}
 */
CNumberingLvl.prototype.GetPStyle = function()
{
	return this.PStyle;
};
/**
 * Устанавливаем связанный стиль
 * @param {string} sStyleId
 */
CNumberingLvl.prototype.SetPStyle = function(sStyleId)
{
	this.PStyle = sStyleId;
};
/**
 * Доступ к начальному значению для данного уровня
 * @returns {number}
 */
CNumberingLvl.prototype.GetStart = function()
{
	return this.Start;
};
/**
 * Доступ к параметру, означающему нужно ли перестартовывать нумерации при смене уровня или оставлять её сквозной
 * @returns {number}
 */
CNumberingLvl.prototype.GetRestart = function()
{
	return this.Restart;
};
/**
 * Доступ к типу разделителя между нумерацией и содержимым параграфа
 * @returns {Asc.c_oAscNumberingSuff}
 */
CNumberingLvl.prototype.GetSuff = function()
{
	return this.Suff;
};
/**
 * Доуступ к текстовым настройкам уровня
 * @returns {CTextPr}
 */
CNumberingLvl.prototype.GetTextPr = function()
{
	return this.TextPr;
};
/**
 * Доступ к настройкам параграфа данного уровня
 * @returns {CParaPr}
 */
CNumberingLvl.prototype.GetParaPr = function()
{
	return this.ParaPr;
};
/**
 * Доступ к содержимому нумерации
 * @returns {[CNumberingLvlTextString | CNumberingLvlTextNum]}
 */
CNumberingLvl.prototype.GetLvlText = function()
{
	return this.LvlText;
};
/**
 * Получение языка нумерации
 * @returns {CLang}
 */
CNumberingLvl.prototype.GetOLang = function()
{
	return this.TextPr && this.TextPr.Lang;
};
/**
 * Выставляем содержимое нумерации
 * @constructor
 */
CNumberingLvl.prototype.SetLvlText = function(arrLvlText)
{
	this.LvlText = arrLvlText;
};
/**
 * Проверяем совместимость с устаревшей нумерацией
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsLegacy = function()
{
	return !!(this.Legacy instanceof CNumberingLvlLegacy && this.Legacy.Legacy);
};
/**
 * Получаем расстояние между символом нумерации и текстом
 * @return {twips}
 */
CNumberingLvl.prototype.GetLegacySpace = function()
{
	if (this.Legacy)
		return this.Legacy.Space;

	return 0;
};
/**
 * Получаем расстояние выделенное под знак нумерации (вместе с расстоянием до текста)
 * @return {twips}
 */
CNumberingLvl.prototype.GetLegacyIndent = function()
{
	if (this.Legacy)
		return this.Legacy.Indent;

	return 0;
};
/**
 * Использовать ли только арабскую нумерацию для предыдущих уровней, используемых на данном уровне
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsLegalStyle = function()
{
	return this.IsLgl;
};
/**
 * Выставляем значения по умолчанию для заданного уровня
 * @param nLvl {number} 0..8
 * @param nType {c_oAscMultiLevelNumbering}
 */
CNumberingLvl.prototype.InitDefault = function(nLvl, nType)
{
	switch (nType)
	{
		case c_oAscMultiLevelNumbering.Numbered:
			this.private_InitDefaultNumbered(nLvl);
			break;
		case c_oAscMultiLevelNumbering.Bullet:
			this.private_InitDefaultBullet(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel1:
			this.private_InitDefaultMultilevel1(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel2:
			this.private_InitDefaultMultilevel2(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel3:
			this.private_InitDefaultMultiLevel3(nLvl);
			break;
		default:
			this.private_InitDefault(nLvl);
	}
};
CNumberingLvl.prototype.private_InitDefault = function(nLvl)
{
	this.Jc      = AscCommon.align_Left;
	this.SetFormat(Asc.c_oAscNumberingFormat.Bullet);
	this.PStyle  = undefined;
	this.Start   = 1;
	this.Restart = -1; // -1 - делаем нумерацию сначала всегда, 0 - никогда не начинаем нумерацию заново
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = 36 * (nLvl + 1) * g_dKoef_pt_to_mm;
	this.ParaPr.Ind.FirstLine = -18 * g_dKoef_pt_to_mm;

	this.TextPr = new CTextPr();
	this.LvlText = [];

	if (0 == nLvl % 3)
	{
		this.TextPr.RFonts.SetAll("Symbol", -1);
		this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00B7)));
	}
	else if (1 == nLvl % 3)
	{
		this.TextPr.RFonts.SetAll("Courier New", -1);
		this.LvlText.push(new CNumberingLvlTextString("o"));
	}
	else
	{
		this.TextPr.RFonts.SetAll("Wingdings", -1);
		this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A7)));
	}
};
/**
 * Выставляем значения по умолчанию для заданного уровня для нумерованного списка
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultNumbered = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;

	var nLeft      = 36 * (nLvl + 1) * g_dKoef_pt_to_mm;
	var nFirstLine = -18 * g_dKoef_pt_to_mm;

	if (0 === nLvl % 3)
	{
		this.Jc = AscCommon.align_Left;
		this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
	}
	else if (1 === nLvl % 3)
	{
		this.Jc = AscCommon.align_Left;
		this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
	}
	else
	{
		this.Jc = AscCommon.align_Right;
		this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
		nFirstLine = -9 * g_dKoef_pt_to_mm;
	}

	this.LvlText = [];
	this.LvlText.push(new CNumberingLvlTextNum(nLvl));
	this.LvlText.push(new CNumberingLvlTextString("."));

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = nLeft;
	this.ParaPr.Ind.FirstLine = nFirstLine;

	this.TextPr = new CTextPr();
};
/**
 * Многоуровневый символьный список
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultBullet = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.Jc      = AscCommon.align_Left;
	this.SetFormat(Asc.c_oAscNumberingFormat.Bullet);

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = 36 * (nLvl + 1) * g_dKoef_pt_to_mm;
	this.ParaPr.Ind.FirstLine = -18 * g_dKoef_pt_to_mm;

	this.TextPr  = new CTextPr();
	this.LvlText = [];
	if (0 === nLvl % 3)
	{
		this.TextPr.RFonts.SetAll("Symbol", -1);
		this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00B7)));
	}
	else if (1 === nLvl % 3)
	{
		this.TextPr.RFonts.SetAll("Courier New", -1);
		this.LvlText.push(new CNumberingLvlTextString("o"));
	}
	else
	{
		this.TextPr.RFonts.SetAll("Wingdings", -1);
		this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A7)));
	}
};
/**
 * Многоуровневый список 1) a) i) 1) a) i) 1) a) i)
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultilevel1 = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.Jc      = AscCommon.align_Left;

	if (0 === nLvl % 3)
	{
		this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
	}
	else if (1 === nLvl % 3)
	{
		this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
	}
	else
	{
		this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
	}

	this.LvlText = [];
	this.LvlText.push(new CNumberingLvlTextNum(nLvl));
	this.LvlText.push(new CNumberingLvlTextString(")"));

	var nLeft      = 18 * (nLvl + 1) * g_dKoef_pt_to_mm;
	var nFirstLine = -18 * g_dKoef_pt_to_mm;

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = nLeft;
	this.ParaPr.Ind.FirstLine = nFirstLine;

	this.TextPr = new CTextPr();
};
/**
 * Многоуровневый список 1. 1.1. 1.1.1. и т.д.
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultilevel2 = function(nLvl)
{
	this.Jc     = AscCommon.align_Left;
	this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;

	var nLeft      = 0;
	var nFirstLine = 0;

	switch (nLvl)
	{
		case 0 :
			nLeft      = 18 * g_dKoef_pt_to_mm;
			nFirstLine = -18 * g_dKoef_pt_to_mm;
			break;
		case 1 :
			nLeft      = 39.6 * g_dKoef_pt_to_mm;
			nFirstLine = -21.6 * g_dKoef_pt_to_mm;
			break;
		case 2 :
			nLeft      = 61.2 * g_dKoef_pt_to_mm;
			nFirstLine = -25.2 * g_dKoef_pt_to_mm;
			break;
		case 3 :
			nLeft      = 86.4 * g_dKoef_pt_to_mm;
			nFirstLine = -32.4 * g_dKoef_pt_to_mm;
			break;
		case 4 :
			nLeft      = 111.6 * g_dKoef_pt_to_mm;
			nFirstLine = -39.6 * g_dKoef_pt_to_mm;
			break;
		case 5 :
			nLeft      = 136.8 * g_dKoef_pt_to_mm;
			nFirstLine = -46.8 * g_dKoef_pt_to_mm;
			break;
		case 6 :
			nLeft      = 162 * g_dKoef_pt_to_mm;
			nFirstLine = -54 * g_dKoef_pt_to_mm;
			break;
		case 7 :
			nLeft      = 187.2 * g_dKoef_pt_to_mm;
			nFirstLine = -61.2 * g_dKoef_pt_to_mm;
			break;
		case 8 :
			nLeft      = 216 * g_dKoef_pt_to_mm;
			nFirstLine = -72 * g_dKoef_pt_to_mm;
			break;
	}

	this.LvlText = [];
	for (var nIndex = 0; nIndex <= nLvl; ++nIndex)
	{
		this.LvlText.push(new CNumberingLvlTextNum(nIndex));
		this.LvlText.push(new CNumberingLvlTextString("."));
	}

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = nLeft;
	this.ParaPr.Ind.FirstLine = nFirstLine;

	this.TextPr = new CTextPr();
};
/**
 * Многоуровневый символьный список
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultiLevel3 = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.SetFormat(Asc.c_oAscNumberingFormat.Bullet);
	this.Jc      = AscCommon.align_Left;
	this.LvlText = [];

	switch (nLvl)
	{
		case 0:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x0076)));
			break;
		case 1:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00D8)));
			break;
		case 2:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A7)));
			break;
		case 3:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00B7)));
			break;
		case 4:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A8)));
			break;
		case 5:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00D8)));
			break;
		case 6:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A7)));
			break;
		case 7:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00B7)));
			break;
		case 8:
			this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(0x00A8)));
			break;
	}

	var nLeft      = 18 * (nLvl + 1) * g_dKoef_pt_to_mm;
	var nFirstLine = -18 * g_dKoef_pt_to_mm;

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = nLeft;
	this.ParaPr.Ind.FirstLine = nFirstLine;

	this.TextPr = new CTextPr();
	if (3 === nLvl || 4 === nLvl || 7 === nLvl || 8 === nLvl)
		this.TextPr.RFonts.SetAll("Symbol", -1);
	else
		this.TextPr.RFonts.SetAll("Wingdings", -1);
};
/**
 * Создаем копию
 * @returns {CNumberingLvl}
 */
CNumberingLvl.prototype.Copy = function()
{
	var oLvl = new CNumberingLvl();

	oLvl.Jc      = this.Jc;
	oLvl.Format  = this.Format;
	oLvl.PStyle  = this.PStyle;
	oLvl.Start   = this.Start;
	oLvl.Restart = this.Restart;
	oLvl.Suff    = this.Suff;
	oLvl.LvlText = [];

	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		oLvl.LvlText.push(this.LvlText[nIndex].Copy());
	}
	oLvl.TextPr = this.TextPr.Copy();
	oLvl.ParaPr = this.ParaPr.Copy();

	if (this.Legacy)
		oLvl.Legacy = this.Legacy.Copy();

	oLvl.IsLgl = this.IsLgl;

	return oLvl;
};
/**
 * Выставляем значения по заданному пресету
 * @param nType {c_oAscNumberingLevel}
 * @param nLvl {number} 0..8
 * @param [sText=undefined] Используется для типа c_oAscNumberingLevel.Bullet
 * @param [oTextPr=undefined] {CTextPr} Используется для типа c_oAscNumberingLevel.Bullet
 */
CNumberingLvl.prototype.SetByType = function(nType, nLvl, sText, oTextPr)
{
	switch (nType)
	{
		case c_oAscNumberingLevel.None:
			this.SetFormat(Asc.c_oAscNumberingFormat.None);
			this.LvlText = [];
			this.TextPr  = new CTextPr();
			break;
		case c_oAscNumberingLevel.Bullet:
			this.SetFormat(Asc.c_oAscNumberingFormat.Bullet);
			this.TextPr  = oTextPr;
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextString(sText));
			break;
		case c_oAscNumberingLevel.DecimalBracket_Right:
			this.Jc      = AscCommon.align_Right;
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.DecimalBracket_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.DecimalDot_Right:
			this.Jc      = AscCommon.align_Right;
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.DecimalDot_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperRomanDot_Right:
			this.Jc      = AscCommon.align_Right;
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperRoman);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperLetterDot_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperLetter);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.LowerLetterBracket_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.LowerLetterDot_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.LowerRomanDot_Right:
			this.Jc      = AscCommon.align_Right;
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperRomanBracket_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperRoman);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.LowerRomanBracket_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperLetterBracket_Left:
			this.Jc      = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperLetter);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
	}
};
/**
 * Получаем тип пресета (если это возможно)
 * @returns {{Type: number, SubType: number}}
 */
CNumberingLvl.prototype.GetPresetType = function()
{
	var nType    = -1;
	var nSubType = -1;

	if (Asc.c_oAscNumberingFormat.Bullet === this.Format)
	{
		nType    = 0;
		nSubType = 0;

		if (1 === this.LvlText.length && numbering_lvltext_Text === this.LvlText[0].Type)
		{
			var nNumVal = this.LvlText[0].Value.charCodeAt(0);

			if (0x00B7 === nNumVal)
				nSubType = 1;
			else if (0x006F === nNumVal)
				nSubType = 2;
			else if (0x00A7 === nNumVal)
				nSubType = 3;
			else if (0x0076 === nNumVal)
				nSubType = 4;
			else if (0x00D8 === nNumVal)
				nSubType = 5;
			else if (0x00FC === nNumVal)
				nSubType = 6;
			else if (0x00A8 === nNumVal)
				nSubType = 7;
			else if (0x2013 === nNumVal)
				nSubType = 8;
		}
	}
	else
	{
		nType    = 1;
		nSubType = 0;

		if (2 === this.LvlText.length && numbering_lvltext_Num === this.LvlText[0].Type && numbering_lvltext_Text === this.LvlText[1].Type)
		{
			var nNumVal2 = this.LvlText[1].Value;

			if (Asc.c_oAscNumberingFormat.Decimal === this.Format)
			{
				if ("." === nNumVal2)
					nSubType = 1;
				else if (")" === nNumVal2)
					nSubType = 2;
			}
			else if (Asc.c_oAscNumberingFormat.UpperRoman === this.Format)
			{
				if ("." === nNumVal2)
					nSubType = 3;
			}
			else if (Asc.c_oAscNumberingFormat.UpperLetter === this.Format)
			{
				if ("." === nNumVal2)
					nSubType = 4;
			}
			else if (Asc.c_oAscNumberingFormat.LowerLetter === this.Format)
			{
				if (")" === nNumVal2)
					nSubType = 5;
				else if ("." === nNumVal2)
					nSubType = 6;
			}
			else if (Asc.c_oAscNumberingFormat.LowerRoman === this.Format)
			{
				if ("." === nNumVal2)
					nSubType = 7;
			}
		}
	}

	return {Type : nType, SubType : nSubType};
};
/**
 * Выставляем значения по заданному формату
 * @param nLvl {number} 0..8
 * @param nType
 * @param sFormatText
 * @param nAlign
 */
CNumberingLvl.prototype.SetByFormat = function(nLvl, nType, sFormatText, nAlign)
{
	this.Jc      = nAlign;
	this.SetFormat(nType);
	this.LvlText = [];

	var nLastPos = 0;
	var nPos     = 0;
	while (-1 !== (nPos = sFormatText.indexOf("%", nPos)) && nPos < sFormatText.length)
	{
		if (nPos < sFormatText.length - 1 && sFormatText.charCodeAt(nPos + 1) >= 49 && sFormatText.charCodeAt(nPos + 1) <= 49 + nLvl)
		{
			if (nPos > nLastPos)
			{
				var sSubString = sFormatText.substring(nLastPos, nPos);
				for (var nSubIndex = 0, nSubLen = sSubString.length; nSubIndex < nSubLen; ++nSubIndex)
					this.LvlText.push(new CNumberingLvlTextString(sSubString.charAt(nSubIndex)));
			}

			this.LvlText.push(new CNumberingLvlTextNum(sFormatText.charCodeAt(nPos + 1) - 49));
			nPos += 2;
			nLastPos = nPos;
		}
		else
		{
			nPos++;
		}
	}
	nPos = sFormatText.length;
	if (nPos > nLastPos)
	{
		var sSubString = sFormatText.substring(nLastPos, nPos);
		for (var nSubIndex = 0, nSubLen = sSubString.length; nSubIndex < nSubLen; ++nSubIndex)
			this.LvlText.push(new CNumberingLvlTextString(sSubString.charAt(nSubIndex)));
	}

	this.TextPr = new CTextPr();
};
/**
 * Собираем статистику документа о количестве слов, букв и т.д.
 * @param oStats объект статистики
 */
CNumberingLvl.prototype.CollectDocumentStatistics = function(oStats)
{
	var bWord = false;
	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		var bSymbol  = false;
		var bSpace   = false;
		var bNewWord = false;

		if (numbering_lvltext_Text === this.LvlText[nIndex].Type && (sp_string === this.LvlText[nIndex].Value || nbsp_string === this.LvlText[nIndex].Value))
		{
			bWord   = false;
			bSymbol = true;
			bSpace  = true;
		}
		else
		{
			if (false === bWord)
				bNewWord = true;

			bWord   = true;
			bSymbol = true;
			bSpace  = false;
		}

		if (true === bSymbol)
			oStats.Add_Symbol(bSpace);

		if (true === bNewWord)
			oStats.Add_Word();
	}

	if (Asc.c_oAscNumberingSuff.Tab === this.Suff || Asc.c_oAscNumberingSuff.Space === this.Suff)
		oStats.Add_Symbol(true);
};
/**
 * Все нумерованные значение переделываем на заданный уровень
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.ResetNumberedText = function(nLvl)
{
	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		if (numbering_lvltext_Num === this.LvlText[nIndex].Type)
			this.LvlText[nIndex].Value = nLvl;
	}
};
/**
 * Проверяем, совпадает ли тип и текст в заданных нумерациях
 * @param oLvl {CNumberingLvl}
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsSimilar = function(oLvl)
{
	if (!oLvl || this.Format !== oLvl.Format || this.LvlText.length !== oLvl.LvlText.length)
		return false;

	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		if (this.LvlText[nIndex].Type !== oLvl.LvlText[nIndex].Type
			|| this.LvlText[nIndex].Value !== oLvl.LvlText[nIndex].Value)
			return false;
	}

	return true;
};
/**
 * Заполняем специальный класс для работы с интерфейсом
 * @param oAscLvl {CAscNumberingLvl}
 */
CNumberingLvl.prototype.FillToAscNumberingLvl = function(oAscLvl)
{
	oAscLvl.put_Format(this.GetFormat());
	oAscLvl.put_Align(this.GetJc());
	oAscLvl.put_Restart(this.GetRestart());
	oAscLvl.put_Start(this.GetStart());
	oAscLvl.put_Suff(this.GetSuff());
	oAscLvl.put_PStyle(this.GetPStyle());

	var arrText = [];
	for (var nPos = 0, nCount = this.LvlText.length; nPos < nCount; ++nPos)
	{
		var oTextElement = this.LvlText[nPos];
		var oAscElement  = new Asc.CAscNumberingLvlText();

		if (numbering_lvltext_Text === oTextElement.Type)
			oAscElement.put_Type(Asc.c_oAscNumberingLvlTextType.Text);
		else
			oAscElement.put_Type(Asc.c_oAscNumberingLvlTextType.Num);

		oAscElement.put_Value(oTextElement.Value);
		arrText.push(oAscElement);
	}
	oAscLvl.put_Text(arrText);

	oAscLvl.TextPr = this.TextPr.Copy();
	oAscLvl.ParaPr = this.ParaPr.Copy();
};
/**
 * Заполняем настройки уровня из интерфейсного класса
 * @param oAscLvl {CAscNumberingLvl}
 */
CNumberingLvl.prototype.FillFromAscNumberingLvl = function(oAscLvl)
{
	if (undefined !== oAscLvl.get_Format())
		this.SetFormat(oAscLvl.get_Format());

	if (undefined !== oAscLvl.get_Align())
		this.Jc = oAscLvl.get_Align();

	if (undefined !== oAscLvl.get_Restart())
		this.Restart = oAscLvl.get_Restart();

	if (undefined !== oAscLvl.get_Start())
		this.Start = oAscLvl.get_Start();

	if (undefined !== oAscLvl.get_Suff())
		this.Suff = oAscLvl.get_Suff();

	if (undefined !== oAscLvl.get_Text())
	{
		var arrAscText = oAscLvl.get_Text();
		for (var nPos = 0, nCount = arrAscText.length; nPos < nCount; ++nPos)
		{
			var oTextElement = arrAscText[nPos];
			var oElement;
			if (Asc.c_oAscNumberingLvlTextType.Text === oTextElement.get_Type())
			{
				oElement = new CNumberingLvlTextString(oTextElement.get_Value());
			}
			else if (Asc.c_oAscNumberingLvlTextType.Num === oTextElement.get_Type())
			{
				oElement = new CNumberingLvlTextNum(oTextElement.get_Value());
			}

			if (oElement)
				this.LvlText.push(oElement);
		}
	}

	if (undefined !== oAscLvl.get_TextPr())
		this.TextPr = oAscLvl.get_TextPr().Copy();

	if (undefined !== oAscLvl.get_ParaPr())
		this.ParaPr = oAscLvl.get_ParaPr().Copy();

	if (undefined !== oAscLvl.get_PStyle())
		this.PStyle = oAscLvl.get_PStyle();
};
CNumberingLvl.prototype.WriteToBinary = function(oWriter)
{
	// Long               : Jc
	// Long               : Format
	// String             : PStyle
	// Long               : Start
	// Long               : Restart
	// Long               : Suff
	// Variable           : TextPr
	// Variable           : ParaPr
	// Long               : количество элементов в LvlText
	// Array of variables : массив LvlText
	// Bool               : true -> CNumberingLegacy
	//                    : false -> Legacy = undefined
	// Bool               : IsLgl

	oWriter.WriteLong(this.Jc);
	oWriter.WriteLong(this.Format);

	oWriter.WriteString2(this.PStyle ? this.PStyle : "");

	oWriter.WriteLong(this.Start);
	oWriter.WriteLong(this.Restart);
	oWriter.WriteLong(this.Suff);

	this.TextPr.WriteToBinary(oWriter);
	this.ParaPr.WriteToBinary(oWriter);

	var nCount = this.LvlText.length;
	oWriter.WriteLong(nCount);

	for (var nIndex = 0; nIndex < nCount; ++nIndex)
		this.LvlText[nIndex].WriteToBinary(oWriter);

	if (this.Legacy instanceof CNumberingLvlLegacy)
	{
		oWriter.WriteBool(true);
		this.Legacy.WriteToBinary(oWriter);
	}
	else
	{
		oWriter.WriteBool(false);
	}

	oWriter.WriteBool(this.IsLgl);
};
CNumberingLvl.prototype.ReadFromBinary = function(oReader)
{
	// Long               : Jc
	// Long               : Format
	// String             : PStyle
	// Long               : Start
	// Long               : Restart
	// Long               : Suff
	// Variable           : TextPr
	// Variable           : ParaPr
	// Long               : количество элементов в LvlText
	// Array of variables : массив LvlText
	// Bool               : true -> CNumberingLegacy
	//                    : false -> Legacy = undefined
	// Bool               : IsLgl


	this.Jc     = oReader.GetLong();
	this.SetFormat(oReader.GetLong());

	this.PStyle = oReader.GetString2();
	if ("" === this.PStyle)
		this.PStyle = undefined;

	this.Start   = oReader.GetLong();
	this.Restart = oReader.GetLong();
	this.Suff    = oReader.GetLong();

	this.TextPr = new CTextPr();
	this.ParaPr = new CParaPr();
	this.TextPr.ReadFromBinary(oReader);
	this.ParaPr.ReadFromBinary(oReader);

	var nCount = oReader.GetLong();
	this.LvlText = [];
	for (var nIndex = 0; nIndex < nCount; ++nIndex)
	{
		var oElement = this.private_ReadLvlTextFromBinary(oReader);
		if (oElement)
			this.LvlText.push(oElement);
	}

	if (oReader.GetBool())
	{
		this.Legacy = new CNumberingLvlLegacy();
		this.Legacy.ReadFromBinary(oReader);
	}

	this.IsLgl = oReader.GetBool();
};
CNumberingLvl.prototype.private_ReadLvlTextFromBinary = function(oReader)
{
	var nType = oReader.GetLong();

	var oElement = null;
	if (numbering_lvltext_Num === nType)
		oElement = new CNumberingLvlTextNum();
	else if (numbering_lvltext_Text === nType)
		oElement = new CNumberingLvlTextString();

	oElement.ReadFromBinary(oReader);
	return oElement;
};
/**
 * Проверяем является ли данный уровень маркированным
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsBulleted = function()
{
	return this.GetFormat() === Asc.c_oAscNumberingFormat.Bullet;
};
/**
 * Проверяем является ли данный уровень нумерованным
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsNumbered = function()
{
	var nFormat = this.GetFormat();
	return (nFormat !== Asc.c_oAscNumberingFormat.Bullet && nFormat !== Asc.c_oAscNumberingFormat.None);
};
/**
 * Получаем список связанных уровней с данным
 * @returns {number[]}
 */
CNumberingLvl.prototype.GetRelatedLvlList = function()
{
	var arrLvls = [];
	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		if (numbering_lvltext_Num === this.LvlText[nIndex].Type)
		{
			var nLvl  = this.LvlText[nIndex].Value;

			if (arrLvls.length <= 0)
				arrLvls.push(nLvl);

			for (var nLvlIndex = 0, nLvlsCount = arrLvls.length; nLvlIndex < nLvlsCount; ++nLvlIndex)
			{
				if (arrLvls[nLvlIndex] === nLvl)
					break;
				else if (arrLvls[nLvlIndex] > nLvl)
					arrLvls.splice(nLvlIndex, 0, nLvl);
			}
		}
	}

	return arrLvls;
};
CNumberingLvl.prototype.SetFormat = function(nFormat)
{
	this.Format = nFormat;
	this.private_CheckSymbols();
};
CNumberingLvl.prototype.private_CheckSymbols = function()
{

	function pickDecimal() {
		for (var i = 0; i < 10; i += 1) {
			AscFonts.FontPickerByCharacter.getFontBySymbol(0x30 + i);
		}
	}

	if (AscFonts.IsCheckSymbols)
	{
		for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
		{
			var oItem = this.LvlText[nIndex];
			if (oItem.IsText())
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(oItem.GetValue().charCodeAt(0));
			}
		}

		switch (this.Format)
		{
			case Asc.c_oAscNumberingFormat.Bullet:
			{
				break;
			}
			case Asc.c_oAscNumberingFormat.Custom:
			case Asc.c_oAscNumberingFormat.BahtText:
			case Asc.c_oAscNumberingFormat.Decimal:
			case Asc.c_oAscNumberingFormat.DecimalZero:
			case Asc.c_oAscNumberingFormat.DollarText:
			case Asc.c_oAscNumberingFormat.DecimalHalfWidth:
			{
				pickDecimal();
				break;
			}
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese: {
				pickDecimal();
				for (var nValue = 0; nValue < 10; ++nValue)
				{
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x2460 + nValue);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle:
			{
				pickDecimal();
				for (var nValue = 0; nValue < 20; ++nValue)
				{
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x2460 + nValue);
				}
				break;
			}

			case Asc.c_oAscNumberingFormat.LowerLetter:
			case Asc.c_oAscNumberingFormat.UpperLetter:
			case Asc.c_oAscNumberingFormat.LowerRoman:
			case Asc.c_oAscNumberingFormat.UpperRoman:
			{
				for (var nValue = 0; nValue < 26; ++nValue)
				{
					AscFonts.FontPickerByCharacter.getFontBySymbol(97 + nValue);
					AscFonts.FontPickerByCharacter.getFontBySymbol(65 + nValue);
				}
				break;
			}

			case Asc.c_oAscNumberingFormat.RussianLower:
			case Asc.c_oAscNumberingFormat.RussianUpper:
			{
				for (var nValue = 0; nValue < 32; ++nValue)
				{
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x0430 + nValue);
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x0410 + nValue);
				}
				break;
			}


			case Asc.c_oAscNumberingFormat.ChineseCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x25CB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5341);
				break;
			}
			case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x25CB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5341);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x767E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5343);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E07);
				break;
			}
			case Asc.c_oAscNumberingFormat.ChineseLegalSimplified:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x96F6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x58F9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8D30);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x53C1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8086);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4F0D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x9646);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x67D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x634C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7396);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x62FE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4F70);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4EDF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x842C);
				break;
			}
			case Asc.c_oAscNumberingFormat.ThaiNumbers:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E50);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E51);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E52);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E53);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E54);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E55);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E56);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E57);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E58);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E59);
				break;
			}
			case Asc.c_oAscNumberingFormat.Aiueo:
			{
				var iter = 0xFF71;
				while (iter !== 0xFF9C) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(iter);
					iter += 1;
				}
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF66);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF9D);
				break;
			}
			case Asc.c_oAscNumberingFormat.AiueoFullWidth:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30ED);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F3);
				break;
			}
			case Asc.c_oAscNumberingFormat.ArabicAbjad:
			case Asc.c_oAscNumberingFormat.ArabicAlpha:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0623);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0628);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x062F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0630);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0631);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0632);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0633);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0634);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0635);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0636);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0637);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0638);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0639);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x063A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0641);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0642);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0643);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0644);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0645);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0646);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0647);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0648);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x064A);
				break;
			}
			case Asc.c_oAscNumberingFormat.CardinalText:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(32);
				AscFonts.FontPickerByCharacter.getFontBySymbol(45);
				AscFonts.FontPickerByCharacter.getFontBySymbol(97);
				AscFonts.FontPickerByCharacter.getFontBySymbol(98);
				AscFonts.FontPickerByCharacter.getFontBySymbol(99);
				AscFonts.FontPickerByCharacter.getFontBySymbol(100);
				AscFonts.FontPickerByCharacter.getFontBySymbol(101);
				AscFonts.FontPickerByCharacter.getFontBySymbol(102);
				AscFonts.FontPickerByCharacter.getFontBySymbol(103);
				AscFonts.FontPickerByCharacter.getFontBySymbol(104);
				AscFonts.FontPickerByCharacter.getFontBySymbol(105);
				AscFonts.FontPickerByCharacter.getFontBySymbol(106);
				AscFonts.FontPickerByCharacter.getFontBySymbol(107);
				AscFonts.FontPickerByCharacter.getFontBySymbol(108);
				AscFonts.FontPickerByCharacter.getFontBySymbol(109);
				AscFonts.FontPickerByCharacter.getFontBySymbol(110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(111);
				AscFonts.FontPickerByCharacter.getFontBySymbol(112);
				AscFonts.FontPickerByCharacter.getFontBySymbol(113);
				AscFonts.FontPickerByCharacter.getFontBySymbol(114);
				AscFonts.FontPickerByCharacter.getFontBySymbol(115);
				AscFonts.FontPickerByCharacter.getFontBySymbol(116);
				AscFonts.FontPickerByCharacter.getFontBySymbol(117);
				AscFonts.FontPickerByCharacter.getFontBySymbol(118);
				AscFonts.FontPickerByCharacter.getFontBySymbol(119);
				AscFonts.FontPickerByCharacter.getFontBySymbol(120);
				AscFonts.FontPickerByCharacter.getFontBySymbol(121);
				AscFonts.FontPickerByCharacter.getFontBySymbol(122);
				AscFonts.FontPickerByCharacter.getFontBySymbol(223);
				AscFonts.FontPickerByCharacter.getFontBySymbol(225);
				AscFonts.FontPickerByCharacter.getFontBySymbol(228);
				AscFonts.FontPickerByCharacter.getFontBySymbol(229);
				AscFonts.FontPickerByCharacter.getFontBySymbol(233);
				AscFonts.FontPickerByCharacter.getFontBySymbol(234);
				AscFonts.FontPickerByCharacter.getFontBySymbol(235);
				AscFonts.FontPickerByCharacter.getFontBySymbol(237);
				AscFonts.FontPickerByCharacter.getFontBySymbol(243);
				AscFonts.FontPickerByCharacter.getFontBySymbol(246);
				AscFonts.FontPickerByCharacter.getFontBySymbol(252);
				AscFonts.FontPickerByCharacter.getFontBySymbol(261);
				AscFonts.FontPickerByCharacter.getFontBySymbol(263);
				AscFonts.FontPickerByCharacter.getFontBySymbol(269);
				AscFonts.FontPickerByCharacter.getFontBySymbol(281);
				AscFonts.FontPickerByCharacter.getFontBySymbol(283);
				AscFonts.FontPickerByCharacter.getFontBySymbol(299);
				AscFonts.FontPickerByCharacter.getFontBySymbol(326);
				AscFonts.FontPickerByCharacter.getFontBySymbol(345);
				AscFonts.FontPickerByCharacter.getFontBySymbol(347);
				AscFonts.FontPickerByCharacter.getFontBySymbol(353);
				AscFonts.FontPickerByCharacter.getFontBySymbol(357);
				AscFonts.FontPickerByCharacter.getFontBySymbol(363);
				AscFonts.FontPickerByCharacter.getFontBySymbol(940);
				AscFonts.FontPickerByCharacter.getFontBySymbol(941);
				AscFonts.FontPickerByCharacter.getFontBySymbol(942);
				AscFonts.FontPickerByCharacter.getFontBySymbol(943);
				AscFonts.FontPickerByCharacter.getFontBySymbol(945);
				AscFonts.FontPickerByCharacter.getFontBySymbol(946);
				AscFonts.FontPickerByCharacter.getFontBySymbol(947);
				AscFonts.FontPickerByCharacter.getFontBySymbol(948);
				AscFonts.FontPickerByCharacter.getFontBySymbol(949);
				AscFonts.FontPickerByCharacter.getFontBySymbol(953);
				AscFonts.FontPickerByCharacter.getFontBySymbol(954);
				AscFonts.FontPickerByCharacter.getFontBySymbol(955);
				AscFonts.FontPickerByCharacter.getFontBySymbol(956);
				AscFonts.FontPickerByCharacter.getFontBySymbol(957);
				AscFonts.FontPickerByCharacter.getFontBySymbol(958);
				AscFonts.FontPickerByCharacter.getFontBySymbol(959);
				AscFonts.FontPickerByCharacter.getFontBySymbol(960);
				AscFonts.FontPickerByCharacter.getFontBySymbol(961);
				AscFonts.FontPickerByCharacter.getFontBySymbol(962);
				AscFonts.FontPickerByCharacter.getFontBySymbol(963);
				AscFonts.FontPickerByCharacter.getFontBySymbol(964);
				AscFonts.FontPickerByCharacter.getFontBySymbol(966);
				AscFonts.FontPickerByCharacter.getFontBySymbol(967);
				AscFonts.FontPickerByCharacter.getFontBySymbol(972);
				AscFonts.FontPickerByCharacter.getFontBySymbol(973);
				AscFonts.FontPickerByCharacter.getFontBySymbol(974);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1072);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1074);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1076);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1077);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1080);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1082);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1083);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1084);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1085);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1086);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1087);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1088);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1089);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1090);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1093);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1094);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1095);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1096);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1099);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1100);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1103);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(45);
				AscFonts.FontPickerByCharacter.getFontBySymbol(65);
				AscFonts.FontPickerByCharacter.getFontBySymbol(66);
				AscFonts.FontPickerByCharacter.getFontBySymbol(67);
				AscFonts.FontPickerByCharacter.getFontBySymbol(68);
				AscFonts.FontPickerByCharacter.getFontBySymbol(69);
				AscFonts.FontPickerByCharacter.getFontBySymbol(70);
				AscFonts.FontPickerByCharacter.getFontBySymbol(71);
				AscFonts.FontPickerByCharacter.getFontBySymbol(72);
				AscFonts.FontPickerByCharacter.getFontBySymbol(73);
				AscFonts.FontPickerByCharacter.getFontBySymbol(74);
				AscFonts.FontPickerByCharacter.getFontBySymbol(75);
				AscFonts.FontPickerByCharacter.getFontBySymbol(76);
				AscFonts.FontPickerByCharacter.getFontBySymbol(77);
				AscFonts.FontPickerByCharacter.getFontBySymbol(78);
				AscFonts.FontPickerByCharacter.getFontBySymbol(79);
				AscFonts.FontPickerByCharacter.getFontBySymbol(80);
				AscFonts.FontPickerByCharacter.getFontBySymbol(81);
				AscFonts.FontPickerByCharacter.getFontBySymbol(82);
				AscFonts.FontPickerByCharacter.getFontBySymbol(83);
				AscFonts.FontPickerByCharacter.getFontBySymbol(84);
				AscFonts.FontPickerByCharacter.getFontBySymbol(85);
				AscFonts.FontPickerByCharacter.getFontBySymbol(86);
				AscFonts.FontPickerByCharacter.getFontBySymbol(87);
				AscFonts.FontPickerByCharacter.getFontBySymbol(88);
				AscFonts.FontPickerByCharacter.getFontBySymbol(89);
				AscFonts.FontPickerByCharacter.getFontBySymbol(90);
				AscFonts.FontPickerByCharacter.getFontBySymbol(193);
				AscFonts.FontPickerByCharacter.getFontBySymbol(196);
				AscFonts.FontPickerByCharacter.getFontBySymbol(197);
				AscFonts.FontPickerByCharacter.getFontBySymbol(201);
				AscFonts.FontPickerByCharacter.getFontBySymbol(202);
				AscFonts.FontPickerByCharacter.getFontBySymbol(203);
				AscFonts.FontPickerByCharacter.getFontBySymbol(205);
				AscFonts.FontPickerByCharacter.getFontBySymbol(211);
				AscFonts.FontPickerByCharacter.getFontBySymbol(214);
				AscFonts.FontPickerByCharacter.getFontBySymbol(220);
				AscFonts.FontPickerByCharacter.getFontBySymbol(260);
				AscFonts.FontPickerByCharacter.getFontBySymbol(262);
				AscFonts.FontPickerByCharacter.getFontBySymbol(268);
				AscFonts.FontPickerByCharacter.getFontBySymbol(280);
				AscFonts.FontPickerByCharacter.getFontBySymbol(282);
				AscFonts.FontPickerByCharacter.getFontBySymbol(298);
				AscFonts.FontPickerByCharacter.getFontBySymbol(325);
				AscFonts.FontPickerByCharacter.getFontBySymbol(344);
				AscFonts.FontPickerByCharacter.getFontBySymbol(346);
				AscFonts.FontPickerByCharacter.getFontBySymbol(352);
				AscFonts.FontPickerByCharacter.getFontBySymbol(356);
				AscFonts.FontPickerByCharacter.getFontBySymbol(362);
				AscFonts.FontPickerByCharacter.getFontBySymbol(902);
				AscFonts.FontPickerByCharacter.getFontBySymbol(904);
				AscFonts.FontPickerByCharacter.getFontBySymbol(905);
				AscFonts.FontPickerByCharacter.getFontBySymbol(906);
				AscFonts.FontPickerByCharacter.getFontBySymbol(908);
				AscFonts.FontPickerByCharacter.getFontBySymbol(910);
				AscFonts.FontPickerByCharacter.getFontBySymbol(911);
				AscFonts.FontPickerByCharacter.getFontBySymbol(913);
				AscFonts.FontPickerByCharacter.getFontBySymbol(914);
				AscFonts.FontPickerByCharacter.getFontBySymbol(915);
				AscFonts.FontPickerByCharacter.getFontBySymbol(916);
				AscFonts.FontPickerByCharacter.getFontBySymbol(917);
				AscFonts.FontPickerByCharacter.getFontBySymbol(921);
				AscFonts.FontPickerByCharacter.getFontBySymbol(922);
				AscFonts.FontPickerByCharacter.getFontBySymbol(923);
				AscFonts.FontPickerByCharacter.getFontBySymbol(924);
				AscFonts.FontPickerByCharacter.getFontBySymbol(925);
				AscFonts.FontPickerByCharacter.getFontBySymbol(926);
				AscFonts.FontPickerByCharacter.getFontBySymbol(927);
				AscFonts.FontPickerByCharacter.getFontBySymbol(928);
				AscFonts.FontPickerByCharacter.getFontBySymbol(929);
				AscFonts.FontPickerByCharacter.getFontBySymbol(931);
				AscFonts.FontPickerByCharacter.getFontBySymbol(932);
				AscFonts.FontPickerByCharacter.getFontBySymbol(934);
				AscFonts.FontPickerByCharacter.getFontBySymbol(935);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1030);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1040);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1042);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1044);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1045);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1048);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1050);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1051);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1052);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1053);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1054);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1055);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1056);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1057);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1058);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1061);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1062);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1063);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1064);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1067);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1068);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1071);
				break;
			}
			case Asc.c_oAscNumberingFormat.Chicago:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x002A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x00A7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x2020);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x2021);
				break;
			}
			case Asc.c_oAscNumberingFormat.Chosung:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3131);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3134);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3137);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3139);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3141);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3142);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3145);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3147);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3148);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x314A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x314B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x314C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x314D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x314E);
				break;
			}
			case Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop:
			{
				for (var i  = 0;i < 20; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x2488 + i);
				}
				pickDecimal();

				break;
			}
			case Asc.c_oAscNumberingFormat.DecimalEnclosedParen:
			{
				for (var i = 0; i < 20; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x2474 + i);
				}
				pickDecimal();
				break;
			}
			case Asc.c_oAscNumberingFormat.DecimalFullWidth:
			{
				for (var i = 0; i < 10; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF10 + i);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.Ganada:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xAC00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB098);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB2E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB77C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB9C8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xBC14);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC0AC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC544);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC790);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xCC28);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xCE74);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD0C0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD30C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD558);
				break;
			}
			case Asc.c_oAscNumberingFormat.Hebrew1:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05EA);
				break;
			}
			case Asc.c_oAscNumberingFormat.Hebrew2:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x05EA);
				break;
			}
			case Asc.c_oAscNumberingFormat.Hex:
			{
				for (var i = 0; i < 6; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(0x0041 + i);
				}
				pickDecimal();
				break;
			}
			case Asc.c_oAscNumberingFormat.HindiConsonants:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(2306);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2307);
				for (var i = 0x0905; i <= 0x0914; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(i);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.HindiCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(2306);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2309);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2310);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2311);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2312);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2313);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2319);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2325);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2327);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2330);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2331);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2332);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2335);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2336);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2337);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2340);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2342);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2344);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2346);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2348);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2351);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2352);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2354);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2357);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2360);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2361);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2364);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2366);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2367);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2368);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2375);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2376);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2379);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2380);
				AscFonts.FontPickerByCharacter.getFontBySymbol(2381);
				break;
			}
			case Asc.c_oAscNumberingFormat.HindiNumbers:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0967);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0968);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0969);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x096F);
				break;
			}
			case Asc.c_oAscNumberingFormat.HindiVowels:
			{
				for (var i = 0x0915; i <= 0x0939; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(i);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographDigital:
			case Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3007);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographEnclosedCircle:
			{
				for (var i = 0x3220; i <= 0x3229; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(i);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographLegalTraditional:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x58F9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8CB3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x53C3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8086);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4F0D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x9678);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x67D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x634C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7396);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x62FE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4F70);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4EDF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x842C);
				AscFonts.FontPickerByCharacter.getFontBySymbol('萬'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('仟'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('佰'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('拾'.charCodeAt(0));
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographTraditional:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E01);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E19);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E59);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x58EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5DF1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5E9A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x620A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7532);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7678);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8F9B);
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographZodiac:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E11);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4EA5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5348);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x536F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5B50);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5BC5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5DF3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x620C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x672A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7533);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8FB0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x9149);
				pickDecimal();
				break;
			}
			case Asc.c_oAscNumberingFormat.IdeographZodiacTraditional:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E01);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E11);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E19);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E59);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4EA5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5348);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x536F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x58EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5B50);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5BC5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5DF1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5DF3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5E9A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x620A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x620D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x672A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7532);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7533);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x7678);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8F9B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x8FB0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x9149);
				break;
			}
			case Asc.c_oAscNumberingFormat.Iroha:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF66);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF71);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF72);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF73);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF74);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF75);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF76);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF77);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF78);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF79);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF7F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF80);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF81);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF82);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF83);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF84);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF85);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF86);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF87);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF88);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF89);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF8F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF90);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF91);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF92);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF93);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF95);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF96);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF97);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF98);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF99);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF9A);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF9B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF9C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF9D);
				break;
			}
			case Asc.c_oAscNumberingFormat.IrohaFullWidth:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30A8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30AF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B3);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B7);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30B9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30BF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30C8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CD);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30CF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D5);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30D8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30DF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E6);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30E9);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EA);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30ED);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30EF);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x30F3);
				break;
			}
			case Asc.c_oAscNumberingFormat.JapaneseCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x3007);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5341);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5343);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x767E);
				AscFonts.FontPickerByCharacter.getFontBySymbol('千'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('百'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('十'.charCodeAt(0));
				break;
			}
			case Asc.c_oAscNumberingFormat.JapaneseLegal:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x58F1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5F10);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x53C2);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4F0D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x62FE);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x767E);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x842C);
				AscFonts.FontPickerByCharacter.getFontBySymbol('萬'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('阡'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('百'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('拾'.charCodeAt(0));
				break;
			}
			case Asc.c_oAscNumberingFormat.KoreanCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC77C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC774);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC0BC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC0AC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC624);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC721);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xCE60);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD314);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xAD6C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC2ED);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB9CC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xCC9C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xBC31);
				AscFonts.FontPickerByCharacter.getFontBySymbol('만'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('천'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('백'.charCodeAt(0));
				AscFonts.FontPickerByCharacter.getFontBySymbol('십'.charCodeAt(0));
				break;
			}
			case Asc.c_oAscNumberingFormat.KoreanDigital:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xAD6C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC0AC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC0BC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC601);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC624);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC721);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC774);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC77C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xCE60);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD314);
				break;
			}
			case Asc.c_oAscNumberingFormat.KoreanDigital2:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x96F6);
				break;
			}
			case Asc.c_oAscNumberingFormat.KoreanLegal:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD558);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB098);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB458);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC14B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB137);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB2E4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC12F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC5EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC12F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC77C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xACF1);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC5EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB35F);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC544);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD649);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC5F4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC2A4);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xBB3C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC11C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB978);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB9C8);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD754);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC270);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC608);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC21C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC77C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD754);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC5EC);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xB4E0);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xC544);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0xD754);
				break;
			}
			case Asc.c_oAscNumberingFormat.None:
			{
				break;
			}
			case Asc.c_oAscNumberingFormat.NumberInDash:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x002D);
				pickDecimal();
				break;
			}
			case Asc.c_oAscNumberingFormat.Ordinal:
			{
				pickDecimal();
				AscFonts.FontPickerByCharacter.getFontBySymbol(45);
				AscFonts.FontPickerByCharacter.getFontBySymbol(46);
				AscFonts.FontPickerByCharacter.getFontBySymbol(58);
				AscFonts.FontPickerByCharacter.getFontBySymbol(97);
				AscFonts.FontPickerByCharacter.getFontBySymbol(100);
				AscFonts.FontPickerByCharacter.getFontBySymbol(101);
				AscFonts.FontPickerByCharacter.getFontBySymbol(104);
				AscFonts.FontPickerByCharacter.getFontBySymbol(110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(114);
				AscFonts.FontPickerByCharacter.getFontBySymbol(115);
				AscFonts.FontPickerByCharacter.getFontBySymbol(116);
				AscFonts.FontPickerByCharacter.getFontBySymbol(176);
				AscFonts.FontPickerByCharacter.getFontBySymbol(186);
				AscFonts.FontPickerByCharacter.getFontBySymbol(959);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1081);
				break;
			}
			case Asc.c_oAscNumberingFormat.OrdinalText:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(32);
				AscFonts.FontPickerByCharacter.getFontBySymbol(45);
				AscFonts.FontPickerByCharacter.getFontBySymbol(97);
				AscFonts.FontPickerByCharacter.getFontBySymbol(98);
				AscFonts.FontPickerByCharacter.getFontBySymbol(99);
				AscFonts.FontPickerByCharacter.getFontBySymbol(100);
				AscFonts.FontPickerByCharacter.getFontBySymbol(101);
				AscFonts.FontPickerByCharacter.getFontBySymbol(102);
				AscFonts.FontPickerByCharacter.getFontBySymbol(103);
				AscFonts.FontPickerByCharacter.getFontBySymbol(104);
				AscFonts.FontPickerByCharacter.getFontBySymbol(105);
				AscFonts.FontPickerByCharacter.getFontBySymbol(106);
				AscFonts.FontPickerByCharacter.getFontBySymbol(107);
				AscFonts.FontPickerByCharacter.getFontBySymbol(108);
				AscFonts.FontPickerByCharacter.getFontBySymbol(109);
				AscFonts.FontPickerByCharacter.getFontBySymbol(110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(111);
				AscFonts.FontPickerByCharacter.getFontBySymbol(112);
				AscFonts.FontPickerByCharacter.getFontBySymbol(113);
				AscFonts.FontPickerByCharacter.getFontBySymbol(114);
				AscFonts.FontPickerByCharacter.getFontBySymbol(115);
				AscFonts.FontPickerByCharacter.getFontBySymbol(116);
				AscFonts.FontPickerByCharacter.getFontBySymbol(117);
				AscFonts.FontPickerByCharacter.getFontBySymbol(118);
				AscFonts.FontPickerByCharacter.getFontBySymbol(119);
				AscFonts.FontPickerByCharacter.getFontBySymbol(120);
				AscFonts.FontPickerByCharacter.getFontBySymbol(121);
				AscFonts.FontPickerByCharacter.getFontBySymbol(122);
				AscFonts.FontPickerByCharacter.getFontBySymbol(223);
				AscFonts.FontPickerByCharacter.getFontBySymbol(225);
				AscFonts.FontPickerByCharacter.getFontBySymbol(228);
				AscFonts.FontPickerByCharacter.getFontBySymbol(229);
				AscFonts.FontPickerByCharacter.getFontBySymbol(233);
				AscFonts.FontPickerByCharacter.getFontBySymbol(234);
				AscFonts.FontPickerByCharacter.getFontBySymbol(235);
				AscFonts.FontPickerByCharacter.getFontBySymbol(237);
				AscFonts.FontPickerByCharacter.getFontBySymbol(243);
				AscFonts.FontPickerByCharacter.getFontBySymbol(246);
				AscFonts.FontPickerByCharacter.getFontBySymbol(252);
				AscFonts.FontPickerByCharacter.getFontBySymbol(261);
				AscFonts.FontPickerByCharacter.getFontBySymbol(263);
				AscFonts.FontPickerByCharacter.getFontBySymbol(269);
				AscFonts.FontPickerByCharacter.getFontBySymbol(281);
				AscFonts.FontPickerByCharacter.getFontBySymbol(283);
				AscFonts.FontPickerByCharacter.getFontBySymbol(299);
				AscFonts.FontPickerByCharacter.getFontBySymbol(326);
				AscFonts.FontPickerByCharacter.getFontBySymbol(345);
				AscFonts.FontPickerByCharacter.getFontBySymbol(347);
				AscFonts.FontPickerByCharacter.getFontBySymbol(353);
				AscFonts.FontPickerByCharacter.getFontBySymbol(357);
				AscFonts.FontPickerByCharacter.getFontBySymbol(363);
				AscFonts.FontPickerByCharacter.getFontBySymbol(940);
				AscFonts.FontPickerByCharacter.getFontBySymbol(941);
				AscFonts.FontPickerByCharacter.getFontBySymbol(942);
				AscFonts.FontPickerByCharacter.getFontBySymbol(943);
				AscFonts.FontPickerByCharacter.getFontBySymbol(945);
				AscFonts.FontPickerByCharacter.getFontBySymbol(946);
				AscFonts.FontPickerByCharacter.getFontBySymbol(947);
				AscFonts.FontPickerByCharacter.getFontBySymbol(948);
				AscFonts.FontPickerByCharacter.getFontBySymbol(949);
				AscFonts.FontPickerByCharacter.getFontBySymbol(953);
				AscFonts.FontPickerByCharacter.getFontBySymbol(954);
				AscFonts.FontPickerByCharacter.getFontBySymbol(955);
				AscFonts.FontPickerByCharacter.getFontBySymbol(956);
				AscFonts.FontPickerByCharacter.getFontBySymbol(957);
				AscFonts.FontPickerByCharacter.getFontBySymbol(958);
				AscFonts.FontPickerByCharacter.getFontBySymbol(959);
				AscFonts.FontPickerByCharacter.getFontBySymbol(960);
				AscFonts.FontPickerByCharacter.getFontBySymbol(961);
				AscFonts.FontPickerByCharacter.getFontBySymbol(962);
				AscFonts.FontPickerByCharacter.getFontBySymbol(963);
				AscFonts.FontPickerByCharacter.getFontBySymbol(964);
				AscFonts.FontPickerByCharacter.getFontBySymbol(966);
				AscFonts.FontPickerByCharacter.getFontBySymbol(967);
				AscFonts.FontPickerByCharacter.getFontBySymbol(972);
				AscFonts.FontPickerByCharacter.getFontBySymbol(973);
				AscFonts.FontPickerByCharacter.getFontBySymbol(974);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1072);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1074);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1076);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1077);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1080);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1082);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1083);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1084);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1085);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1086);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1087);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1088);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1089);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1090);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1093);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1094);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1095);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1096);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1099);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1100);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1103);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(65);
				AscFonts.FontPickerByCharacter.getFontBySymbol(66);
				AscFonts.FontPickerByCharacter.getFontBySymbol(67);
				AscFonts.FontPickerByCharacter.getFontBySymbol(68);
				AscFonts.FontPickerByCharacter.getFontBySymbol(69);
				AscFonts.FontPickerByCharacter.getFontBySymbol(70);
				AscFonts.FontPickerByCharacter.getFontBySymbol(71);
				AscFonts.FontPickerByCharacter.getFontBySymbol(72);
				AscFonts.FontPickerByCharacter.getFontBySymbol(73);
				AscFonts.FontPickerByCharacter.getFontBySymbol(74);
				AscFonts.FontPickerByCharacter.getFontBySymbol(75);
				AscFonts.FontPickerByCharacter.getFontBySymbol(76);
				AscFonts.FontPickerByCharacter.getFontBySymbol(77);
				AscFonts.FontPickerByCharacter.getFontBySymbol(78);
				AscFonts.FontPickerByCharacter.getFontBySymbol(79);
				AscFonts.FontPickerByCharacter.getFontBySymbol(80);
				AscFonts.FontPickerByCharacter.getFontBySymbol(81);
				AscFonts.FontPickerByCharacter.getFontBySymbol(82);
				AscFonts.FontPickerByCharacter.getFontBySymbol(83);
				AscFonts.FontPickerByCharacter.getFontBySymbol(84);
				AscFonts.FontPickerByCharacter.getFontBySymbol(85);
				AscFonts.FontPickerByCharacter.getFontBySymbol(86);
				AscFonts.FontPickerByCharacter.getFontBySymbol(87);
				AscFonts.FontPickerByCharacter.getFontBySymbol(88);
				AscFonts.FontPickerByCharacter.getFontBySymbol(89);
				AscFonts.FontPickerByCharacter.getFontBySymbol(90);
				AscFonts.FontPickerByCharacter.getFontBySymbol(193);
				AscFonts.FontPickerByCharacter.getFontBySymbol(196);
				AscFonts.FontPickerByCharacter.getFontBySymbol(197);
				AscFonts.FontPickerByCharacter.getFontBySymbol(201);
				AscFonts.FontPickerByCharacter.getFontBySymbol(202);
				AscFonts.FontPickerByCharacter.getFontBySymbol(203);
				AscFonts.FontPickerByCharacter.getFontBySymbol(205);
				AscFonts.FontPickerByCharacter.getFontBySymbol(211);
				AscFonts.FontPickerByCharacter.getFontBySymbol(214);
				AscFonts.FontPickerByCharacter.getFontBySymbol(220);
				AscFonts.FontPickerByCharacter.getFontBySymbol(260);
				AscFonts.FontPickerByCharacter.getFontBySymbol(262);
				AscFonts.FontPickerByCharacter.getFontBySymbol(268);
				AscFonts.FontPickerByCharacter.getFontBySymbol(280);
				AscFonts.FontPickerByCharacter.getFontBySymbol(282);
				AscFonts.FontPickerByCharacter.getFontBySymbol(298);
				AscFonts.FontPickerByCharacter.getFontBySymbol(325);
				AscFonts.FontPickerByCharacter.getFontBySymbol(344);
				AscFonts.FontPickerByCharacter.getFontBySymbol(346);
				AscFonts.FontPickerByCharacter.getFontBySymbol(352);
				AscFonts.FontPickerByCharacter.getFontBySymbol(356);
				AscFonts.FontPickerByCharacter.getFontBySymbol(362);
				AscFonts.FontPickerByCharacter.getFontBySymbol(902);
				AscFonts.FontPickerByCharacter.getFontBySymbol(904);
				AscFonts.FontPickerByCharacter.getFontBySymbol(905);
				AscFonts.FontPickerByCharacter.getFontBySymbol(906);
				AscFonts.FontPickerByCharacter.getFontBySymbol(908);
				AscFonts.FontPickerByCharacter.getFontBySymbol(910);
				AscFonts.FontPickerByCharacter.getFontBySymbol(911);
				AscFonts.FontPickerByCharacter.getFontBySymbol(913);
				AscFonts.FontPickerByCharacter.getFontBySymbol(914);
				AscFonts.FontPickerByCharacter.getFontBySymbol(915);
				AscFonts.FontPickerByCharacter.getFontBySymbol(916);
				AscFonts.FontPickerByCharacter.getFontBySymbol(917);
				AscFonts.FontPickerByCharacter.getFontBySymbol(921);
				AscFonts.FontPickerByCharacter.getFontBySymbol(922);
				AscFonts.FontPickerByCharacter.getFontBySymbol(923);
				AscFonts.FontPickerByCharacter.getFontBySymbol(924);
				AscFonts.FontPickerByCharacter.getFontBySymbol(925);
				AscFonts.FontPickerByCharacter.getFontBySymbol(926);
				AscFonts.FontPickerByCharacter.getFontBySymbol(927);
				AscFonts.FontPickerByCharacter.getFontBySymbol(928);
				AscFonts.FontPickerByCharacter.getFontBySymbol(929);
				AscFonts.FontPickerByCharacter.getFontBySymbol(931);
				AscFonts.FontPickerByCharacter.getFontBySymbol(932);
				AscFonts.FontPickerByCharacter.getFontBySymbol(934);
				AscFonts.FontPickerByCharacter.getFontBySymbol(935);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1030);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1040);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1042);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1044);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1045);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1048);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1050);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1051);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1052);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1053);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1054);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1055);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1056);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1057);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1058);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1061);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1062);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1063);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1064);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1067);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1068);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1071);
				AscFonts.FontPickerByCharacter.getFontBySymbol(200);
				AscFonts.FontPickerByCharacter.getFontBySymbol(212);
				AscFonts.FontPickerByCharacter.getFontBySymbol(221);
				AscFonts.FontPickerByCharacter.getFontBySymbol(919);
				AscFonts.FontPickerByCharacter.getFontBySymbol(937);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1043);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1049);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1059);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1066);
				AscFonts.FontPickerByCharacter.getFontBySymbol(232);
				AscFonts.FontPickerByCharacter.getFontBySymbol(244);
				AscFonts.FontPickerByCharacter.getFontBySymbol(253);
				AscFonts.FontPickerByCharacter.getFontBySymbol(951);
				AscFonts.FontPickerByCharacter.getFontBySymbol(969);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1075);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1081);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1091);
				AscFonts.FontPickerByCharacter.getFontBySymbol(1098);
				break;
			}
			case Asc.c_oAscNumberingFormat.TaiwaneseCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x5341);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				break;
			}
			case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);

				break;
			}
			case Asc.c_oAscNumberingFormat.TaiwaneseDigital:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x25CB);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
				break;
			}
			case Asc.c_oAscNumberingFormat.ThaiCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(3627);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3609);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3638);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3656);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3591);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3626);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3629);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3634);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3617);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3637);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3657);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3585);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3648);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3592);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3655);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3604);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3649);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3611);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3636);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3610);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3618);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3621);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3639);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3619);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3614);
				AscFonts.FontPickerByCharacter.getFontBySymbol(3633);

				break;
			}
			case Asc.c_oAscNumberingFormat.ThaiLetters:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E01);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E02);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E04);
				AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E25);
				for (var i = 0x0E07; i <= 0x0E23; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(i);
				}
				for (var i = 0x0E27; i <= 0x0E2E; i += 1) {
					AscFonts.FontPickerByCharacter.getFontBySymbol(i);
				}
				break;
			}
			case Asc.c_oAscNumberingFormat.VietnameseCounting:
			{
				AscFonts.FontPickerByCharacter.getFontBySymbol(97);
				AscFonts.FontPickerByCharacter.getFontBySymbol(98);
				AscFonts.FontPickerByCharacter.getFontBySymbol(99);
				AscFonts.FontPickerByCharacter.getFontBySymbol(104);
				AscFonts.FontPickerByCharacter.getFontBySymbol(105);
				AscFonts.FontPickerByCharacter.getFontBySymbol(109);
				AscFonts.FontPickerByCharacter.getFontBySymbol(110);
				AscFonts.FontPickerByCharacter.getFontBySymbol(115);
				AscFonts.FontPickerByCharacter.getFontBySymbol(116);
				AscFonts.FontPickerByCharacter.getFontBySymbol(117);
				AscFonts.FontPickerByCharacter.getFontBySymbol(121);
				AscFonts.FontPickerByCharacter.getFontBySymbol(225);
				AscFonts.FontPickerByCharacter.getFontBySymbol(237);
				AscFonts.FontPickerByCharacter.getFontBySymbol(259);
				AscFonts.FontPickerByCharacter.getFontBySymbol(432);
				AscFonts.FontPickerByCharacter.getFontBySymbol(7843);
				AscFonts.FontPickerByCharacter.getFontBySymbol(7889);
				AscFonts.FontPickerByCharacter.getFontBySymbol(7897);
				AscFonts.FontPickerByCharacter.getFontBySymbol(7901);
				break;
			}
		}
	}
}



function CNumberingLvlTextString(Val)
{
	if ("string" == typeof(Val))
		this.Value = Val;
	else
		this.Value = "";

	if (AscFonts.IsCheckSymbols)
		AscFonts.FontPickerByCharacter.getFontsByString(this.Value);

	this.Type = numbering_lvltext_Text;
}
CNumberingLvlTextString.prototype.IsLvl = function()
{
	return false;
};
CNumberingLvlTextString.prototype.IsText = function()
{
	return true;
};
CNumberingLvlTextString.prototype.GetValue = function()
{
	return this.Value;
};
CNumberingLvlTextString.prototype.Copy = function()
{
	return new CNumberingLvlTextString(this.Value);
};
CNumberingLvlTextString.prototype.WriteToBinary = function(Writer)
{
	// Long   : numbering_lvltext_Text
	// String : Value

	Writer.WriteLong(numbering_lvltext_Text);
	Writer.WriteString2(this.Value);
};
CNumberingLvlTextString.prototype.ReadFromBinary = function(Reader)
{
	this.Value = Reader.GetString2();

	if (AscFonts.IsCheckSymbols)
		AscFonts.FontPickerByCharacter.getFontsByString(this.Value);
};

function CNumberingLvlTextNum(Lvl)
{
	if ("number" == typeof(Lvl))
		this.Value = Lvl;
	else
		this.Value = 0;

	this.Type = numbering_lvltext_Num;
}
CNumberingLvlTextNum.prototype.IsLvl = function()
{
	return true;
};
CNumberingLvlTextNum.prototype.IsText = function()
{
	return false;
};
CNumberingLvlTextNum.prototype.GetValue = function()
{
	return this.Value;
};
CNumberingLvlTextNum.prototype.Copy = function()
{
	return new CNumberingLvlTextNum(this.Value);
};
CNumberingLvlTextNum.prototype.WriteToBinary = function(Writer)
{
	// Long : numbering_lvltext_Text
	// Long : Value

	Writer.WriteLong(numbering_lvltext_Num);
	Writer.WriteLong(this.Value);
};
CNumberingLvlTextNum.prototype.ReadFromBinary = function(Reader)
{
	this.Value = Reader.GetLong();
};

function CNumberingLvlLegacy(isUse, twIndent, twSpace)
{
	this.Legacy = !!isUse;
	this.Indent = twIndent ? twIndent : 0; // Значение в твипсах
	this.Space  = twSpace ? twSpace : 0;   // Значение в твипсах
}
CNumberingLvlLegacy.prototype.Copy = function()
{
	return new CNumberingLvlLegacy(this.Legacy, this.Indent, this.Space);
};
CNumberingLvlLegacy.prototype.WriteToBinary = function(oWriter)
{
	// Bool : Legacy
	// Long : Indent
	// Long : Space
	oWriter.WriteBool(this.Legacy);
	oWriter.WriteLong(this.Indent);
	oWriter.WriteLong(this.Space);
};
CNumberingLvlLegacy.prototype.ReadFromBinary = function(oReader)
{
	// Bool : Legacy
	// Long : Indent
	// Long : Space
	this.Legacy = oReader.GetBool();
	this.Indent = oReader.GetLong();
	this.Space  = oReader.GetLong();
};
