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
 * Устанавливаем тип прилегания
 * @param nJc {AscCommon.align_Left | AscCommon.align_Right | AscCommon.align_Center}
 */
CNumberingLvl.prototype.SetJc = function(nJc)
{
	this.Jc = nJc;
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
 * @param paraPr {AscWord.CTextPr}
 */
CNumberingLvl.prototype.SetTextPr = function(oTextPr)
{
	this.TextPr = oTextPr;
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
 * @param paraPr {AscWord.CParaPr}
 */
CNumberingLvl.prototype.SetParaPr = function(paraPr)
{
	this.ParaPr = paraPr;
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
 */
CNumberingLvl.prototype.SetLvlText = function(arrLvlText)
{
	this.LvlText = arrLvlText;
};
CNumberingLvl.prototype.AddStringToLvlText = function(text)
{
	for (let iterator = text.getUnicodeIterator(); iterator.check(); iterator.next())
	{
		this.LvlText.push(new CNumberingLvlTextString(String.fromCharCode(iterator.value())));
	}
};
CNumberingLvl.prototype.AddLvlToLvlText = function(iLvl)
{
	this.LvlText.push(new CNumberingLvlTextNum(iLvl));
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
 * @param iLvl {number} 0..8
 * @param type {c_oAscMultiLevelNumbering}
 */
CNumberingLvl.CreateDefault = function(iLvl, type)
{
	let numLvl = new CNumberingLvl();
	numLvl.InitDefault(iLvl, type);
	return numLvl;
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
		case c_oAscMultiLevelNumbering.MultiLevel_1_a_i:
			this.private_InitDefaultMultiLevel_1_a_i(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_1_11_111:
			this.private_InitDefaultMultiLevel_1_11_111(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_Bullet:
			this.private_InitDefaultMultiLevel_Bullet(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_Article_Section:
			this.private_InitDefaultMultiLevel_Article_Section(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_Chapter:
			this.private_InitDefaultMultiLevel_Chapter(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_I_A_1:
			this.private_InitDefaultMultiLevel_I_A_1(nLvl);
			break;
		case c_oAscMultiLevelNumbering.MultiLevel_1_11_111_NoInd:
			this.private_InitDefaultMultiLevel_1_11_111_NoInd(nLvl);
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
CNumberingLvl.prototype.private_InitDefaultMultiLevel_1_a_i = function(nLvl)
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
CNumberingLvl.prototype.private_InitDefaultMultiLevel_1_11_111 = function(nLvl)
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
CNumberingLvl.prototype.private_InitDefaultMultiLevel_Bullet = function(nLvl)
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
 * Многоуровневый список
 * Article I
 *   Section 1.01
 *     (a)
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultiLevel_Article_Section = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.LvlText = [];
	this.ParaPr  = new CParaPr();
	this.TextPr  = new CTextPr();
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;

	switch (nLvl)
	{
		case 0:

			this.SetFormat(Asc.c_oAscNumberingFormat.UpperRoman);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 0;
			this.ParaPr.Ind.FirstLine = 0;

			this.AddStringToLvlText("Article ");
			this.AddLvlToLvlText(0);
			this.AddStringToLvlText(".");

			break;
		case 1:

			this.SetFormat(Asc.c_oAscNumberingFormat.DecimalZero);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 0;
			this.ParaPr.Ind.FirstLine = 0;

			this.AddStringToLvlText("Section ");
			this.AddLvlToLvlText(0);
			this.AddStringToLvlText(".");
			this.AddLvlToLvlText(1);

			break;

		case 2:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 720 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -432 * g_dKoef_twips_to_mm;

			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(2);
			this.AddStringToLvlText(")");

			break;
		case 3:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.Jc      = AscCommon.align_Right;

			this.ParaPr.Ind.Left      = 864 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -144 * g_dKoef_twips_to_mm;

			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(3);
			this.AddStringToLvlText(")");
			break;
		case 4:
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 1008 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -432 * g_dKoef_twips_to_mm;

			this.AddLvlToLvlText(4);
			this.AddStringToLvlText(")");
			break;
		case 5:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 1152 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -432 * g_dKoef_twips_to_mm;

			this.AddLvlToLvlText(5);
			this.AddStringToLvlText(")");
			break;
		case 6:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.Jc      = AscCommon.align_Right;

			this.ParaPr.Ind.Left      = 1296 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -288 * g_dKoef_twips_to_mm;

			this.AddLvlToLvlText(6);
			this.AddStringToLvlText(")");
			break;
		case 7:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.Jc      = AscCommon.align_Left;

			this.ParaPr.Ind.Left      = 1440 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -432 * g_dKoef_twips_to_mm;

			this.AddLvlToLvlText(7);
			this.AddStringToLvlText(".");
			break;
		case 8:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.Jc      = AscCommon.align_Right;

			this.ParaPr.Ind.Left      = 1584 * g_dKoef_twips_to_mm;
			this.ParaPr.Ind.FirstLine = -144 * g_dKoef_twips_to_mm;

			this.AddLvlToLvlText(8);
			this.AddStringToLvlText(".");
			break;
	}
};
/**
 * Многоуровневый список
 * Chapter 1
 * (none)
 * (none)
 * ...
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultiLevel_Chapter = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.LvlText = [];
	this.ParaPr  = new CParaPr();
	this.TextPr  = new CTextPr();
	this.Suff    = Asc.c_oAscNumberingSuff.None;
	this.Jc      = AscCommon.align_Left;

	this.ParaPr.Ind.Left      = 0;
	this.ParaPr.Ind.FirstLine = 0;

	if (0 === nLvl)
	{
		this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
		this.Suff = Asc.c_oAscNumberingSuff.Space;
		this.AddStringToLvlText("Chapter ");
		this.AddLvlToLvlText(0);
	}
	else
	{
		this.SetFormat(Asc.c_oAscNumberingFormat.None);
	}


};
/**
 * Многоуровневый список
 * I. A. 1. a) (1) (a) (i) (a) (i)
 * @param nLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultiLevel_I_A_1 = function(nLvl)
{
	this.Start   = 1;
	this.Restart = -1;
	this.LvlText = [];
	this.ParaPr  = new CParaPr();
	this.TextPr  = new CTextPr();
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.Jc      = AscCommon.align_Left;

	this.ParaPr.Ind.Left      = nLvl * 720 * g_dKoef_twips_to_mm;
	this.ParaPr.Ind.FirstLine = 0;

	switch (nLvl)
	{
		case 0:
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperRoman);
			this.AddLvlToLvlText(0);
			this.AddStringToLvlText(".");
			break;
		case 1:
			this.SetFormat(Asc.c_oAscNumberingFormat.UpperLetter);
			this.AddLvlToLvlText(1);
			this.AddStringToLvlText(".");
			break;
		case 2:
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.AddLvlToLvlText(2);
			this.AddStringToLvlText(".");
			break;
		case 3:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.AddLvlToLvlText(3);
			this.AddStringToLvlText(")");
			break;
		case 4:
			this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(4);
			this.AddStringToLvlText(")");
			break;
		case 5:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(5);
			this.AddStringToLvlText(")");
			break;
		case 6:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(6);
			this.AddStringToLvlText(")");
			break;
		case 7:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerLetter);
			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(7);
			this.AddStringToLvlText(")");
			break;
		case 8:
			this.SetFormat(Asc.c_oAscNumberingFormat.LowerRoman);
			this.AddStringToLvlText("(");
			this.AddLvlToLvlText(8);
			this.AddStringToLvlText(")");
			break;
	}
};
/**
 * Многоуровневый список 1. 1.1. 1.1.1. и т.д.
 * @param iLvl {number} 0..8
 */
CNumberingLvl.prototype.private_InitDefaultMultiLevel_1_11_111_NoInd = function(iLvl)
{
	this.SetFormat(Asc.c_oAscNumberingFormat.Decimal);
	this.Jc      = AscCommon.align_Left;
	this.Start   = 1;
	this.Restart = -1;
	this.Suff    = Asc.c_oAscNumberingSuff.Tab;
	this.TextPr  = new CTextPr();

	this.ParaPr               = new CParaPr();
	this.ParaPr.Ind.Left      = (432 + 144 * iLvl) * g_dKoef_twips_to_mm;
	this.ParaPr.Ind.FirstLine = -this.ParaPr.Ind.Left;

	this.LvlText = [];
	for (let index = 0; index <= iLvl; ++index)
	{
		this.LvlText.push(new CNumberingLvlTextNum(index));
		this.LvlText.push(new CNumberingLvlTextString("."));
	}
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
		case c_oAscNumberingLevel.LowerRussian_Dot_Left:
			this.Jc = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.RussianLower);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.LowerRussian_Bracket_Left:
			this.Jc = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.RussianLower);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString(")"));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperRussian_Dot_Left:
			this.Jc = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.RussianUpper);
			this.LvlText = [];
			this.LvlText.push(new CNumberingLvlTextNum(nLvl));
			this.LvlText.push(new CNumberingLvlTextString("."));
			this.TextPr = new CTextPr();
			break;
		case c_oAscNumberingLevel.UpperRussian_Bracket_Left:
			this.Jc = AscCommon.align_Left;
			this.SetFormat(Asc.c_oAscNumberingFormat.RussianUpper);
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

			switch (this.Format)
			{
				case Asc.c_oAscNumberingFormat.Decimal:
					if ("." === nNumVal2)
						nSubType = 1;
					else if (")" === nNumVal2)
						nSubType = 2;
					break;
				case Asc.c_oAscNumberingFormat.UpperRoman:
					if ("." === nNumVal2)
						nSubType = 3;
					break;
				case Asc.c_oAscNumberingFormat.UpperLetter:
					if ("." === nNumVal2)
						nSubType = 4;
					break;
				case Asc.c_oAscNumberingFormat.LowerLetter:
					if (")" === nNumVal2)
						nSubType = 5;
					else if ("." === nNumVal2)
						nSubType = 6;
					break;
				case Asc.c_oAscNumberingFormat.LowerRoman:
					if ("." === nNumVal2)
						nSubType = 7;
					break;
				case Asc.c_oAscNumberingFormat.RussianUpper:
					if ("." === nNumVal2)
						nSubType = 8;
					else if (")" === nNumVal2)
						nSubType = 9;
					break;
				case Asc.c_oAscNumberingFormat.RussianLower:
					if ("." === nNumVal2)
						nSubType = 10;
					else if (")" === nNumVal2)
						nSubType = 11;
					break;

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
	this.SetLvlTextFormat(nLvl, sFormatText);
	this.TextPr = new CTextPr();
};
/**
 * Выставляем LvlText по заданному формату
 * @param nLvl {number} 0..8
 * @param sFormatText
 */
CNumberingLvl.prototype.SetLvlTextFormat = function(nLvl, sFormatText)
{
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
};
/**
 * Получаем LvlText в виде строки для записи в xml
 * @returns {string}
 */
CNumberingLvl.prototype.GetLvlTextFormat = function() {
	var res = "";
	for (var i = 0; i < this.LvlText.length; ++i) {
		if (this.LvlText[i].IsLvl()) {
			res += "%" + (this.LvlText[i].Value + 1)
		} else {
			res += this.LvlText[i].Value
		}
	}
	return res;
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
CNumberingLvl.prototype.IsEqual = function(numLvl)
{
	// Формат и текст проверяются в IsSimilar
	if (!this.IsSimilar(numLvl))
		return false;
	
	return (this.Jc === numLvl.Jc
		&& this.PStyle === numLvl.PStyle
		&& this.Start === numLvl.Start
		&& this.Restart === numLvl.Restart
		&& this.Suff === numLvl.Suff
		&& this.TextPr.IsEqual(numLvl.TextPr)
		&& this.ParaPr.IsEqual(numLvl.ParaPr)
		&& this.Legacy === numLvl.Legacy
		&& this.IsLgl === numLvl.IsLgl);
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
CNumberingLvl.prototype.FillLvlText = function(arrOfInfo)
{
	for (let i = 0; i < arrOfInfo.length; i += 1)
	{
		if (AscFormat.isRealNumber(arrOfInfo[i]))
		{
			this.LvlText.push(new CNumberingLvlTextNum(arrOfInfo[i]));
		}
		else if (typeof arrOfInfo[i] === "string")
		{
			for (const oUnicodeIterator = arrOfInfo[i].getUnicodeIterator(); oUnicodeIterator.check(); oUnicodeIterator.next())
			{
				this.LvlText.push(new CNumberingLvlTextString(AscCommon.encodeSurrogateChar(oUnicodeIterator.value())));
			}
		}
	}
};
// TODO: исправить при добавлении картиночных буллетов
CNumberingLvl.prototype.IsImageBullet = function ()
{
	return false;
};
/**
 *
 * @returns {AscFonts.CImage}
 */
CNumberingLvl.prototype.GetImage = function ()
{

};
/**
 *
 * @returns {String | Object}
 */
CNumberingLvl.prototype.GetDrawingContent = function (arrLvls, nLvl, nNum, oLang)
{
	if (this.IsImageBullet())
	{
		const oImage = this.GetImage();
		const oResult = {image: oImage, amount: 0};
		if (oImage)
		{
			for (let i = 0; i < this.LvlText.length; i += 1)
			{
				const oNumberingLvlText = this.LvlText[i];
				if (oNumberingLvlText.IsText())
				{
					oResult.amount += 1;
				}
			}
		}
		return oResult;
	}
	else
	{
		return this.GetStringByLvlText(arrLvls, nLvl, nNum, oLang);
	}
}

CNumberingLvl.prototype.GetStringByLvlText = function (arrLvls, nLvl, nNum, oLang)
{
	const arrResult = [];
	for (let i = 0; i < this.LvlText.length; i += 1)
	{
		const oNumberingLvlText = this.LvlText[i];

		if (oNumberingLvlText.IsText())
		{
			arrResult.push(oNumberingLvlText.GetValue());
		}
		else
		{
			if (AscFormat.isRealNumber(nNum))
			{
				const nNumberingLvl = oNumberingLvlText.GetValue();
				let nFormat = this.GetFormat();
				if (nNumberingLvl === nLvl)
				{
					nNum = (this.GetStart() - 1) + nNum;
				}
				else if (arrLvls[nNumberingLvl] && nLvl > nNumberingLvl)
				{
					nFormat = arrLvls[nNumberingLvl].GetFormat();
					nNum = arrLvls[nNumberingLvl].GetStart();
				}
				arrResult.push(AscCommon.IntToNumberFormat(nNum, nFormat, oLang));
			}
		}
	}

	return arrResult.join('');
};
CNumberingLvl.prototype.GetNumberPosition = function ()
{
	const nLeft = this.GetIndentSize() || 0;
	if (AscFormat.isRealNumber(this.ParaPr.Ind.FirstLine))
	{
		return nLeft + this.ParaPr.Ind.FirstLine;
	}
	return nLeft;
};
CNumberingLvl.prototype.GetIndentSize = function ()
{
	return this.ParaPr && this.ParaPr.Ind ? this.ParaPr.Ind.Left : 0;
};
CNumberingLvl.prototype.GetStopTab = function ()
{
	const oParaPr = this.GetParaPr();
	if (oParaPr)
	{
		const oTabs = oParaPr.GetTabs();
		if (oTabs)
		{
			if (oTabs && oTabs.GetCount() === 1)
			{
				return oTabs.Get(0).Pos;
			}
		}
	}
	return null;
};

CNumberingLvl.prototype.SetStopTab = function (nValue)
{
	var oParaPr = this.ParaPr;
	if (!oParaPr)
	{
		oParaPr = new AscCommonWord.CParaPr;
		this.ParaPr = oParaPr;
	}
	if (AscFormat.isRealNumber(nValue))
	{
		var oTabs = new AscCommonWord.CParaTabs;
		oTabs.Add(new AscCommonWord.CParaTab(Asc.c_oAscTabType.Num, nValue));
		oParaPr.Tabs = oTabs;
	}
	else
	{
		delete oParaPr.Tabs;
	}
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
	return AscWord.IsBulletedNumbering(this.GetFormat());
};
/**
 * Проверяем является ли данный уровень нумерованным
 * @returns {boolean}
 */
CNumberingLvl.prototype.IsNumbered = function()
{
	return AscWord.IsNumberedNumbering(this.GetFormat());
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

CNumberingLvl.prototype.private_CheckSymbols = function ()
{
	if (AscFonts.IsCheckSymbols)
	{
		const sSymbols = this.GetSymbols();
		AscFonts.FontPickerByCharacter.checkTextLight(sSymbols);
	}
}

CNumberingLvl.prototype.GetSymbols = function()
{
	let arrSymbols = [];
	for (let index = 0, count = this.LvlText.length; index < count; ++index)
	{
		let textItem = this.LvlText[index];
		if (textItem.IsText())
			arrSymbols.push(textItem.GetValue());
	}

	function appendDecimal() {
		arrSymbols.push("0123456789");
	}

	if (this.IsLgl)
		appendDecimal();

	let arrCodesOfSymbols = [];
	switch (this.Format)
	{
		case Asc.c_oAscNumberingFormat.Aiueo:
		{
			arrCodesOfSymbols = [0xFF71, 0xFF72, 0xFF73, 0xFF74, 0xFF75, 0xFF76, 0xFF77, 0xFF78, 0xFF79, 0xFF7A, 0xFF7B,
				0xFF7C, 0xFF7D, 0xFF7E, 0xFF7F, 0xFF80, 0xFF81, 0xFF82, 0xFF83, 0xFF84, 0xFF85, 0xFF86, 0xFF87, 0xFF88,
				0xFF89, 0xFF8A, 0xFF8B, 0xFF8C, 0xFF8D, 0xFF8E, 0xFF8F, 0xFF90, 0xFF91, 0xFF92, 0xFF93, 0xFF94, 0xFF95,
				0xFF96, 0xFF97, 0xFF98, 0xFF99, 0xFF9A, 0xFF9B, 0xFF66, 0xFF9D];
			break;
		}
		case Asc.c_oAscNumberingFormat.ArabicAbjad:
		case Asc.c_oAscNumberingFormat.ArabicAlpha:
		{
			arrCodesOfSymbols = [0x0623, 0x0628, 0x062A, 0x062B, 0x062C, 0x062D, 0x062E, 0x062F, 0x0630, 0x0631, 0x0632,
				0x0633, 0x0634, 0x0635, 0x0636, 0x0637, 0x0638, 0x0639, 0x063A, 0x0641, 0x0642, 0x0643, 0x0644, 0x0645,
				0x0646, 0x0647, 0x0648, 0x064A];
			break;
		}
		case Asc.c_oAscNumberingFormat.CardinalText:
		{
			arrCodesOfSymbols = [0x0020, 0x002D, 0x0061, 0x0062, 0x0063, 0x0064, 0x0065, 0x0066, 0x0067, 0x0068, 0x0069,
				0x006A, 0x006B, 0x006C, 0x006D, 0x006E, 0x006F, 0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075, 0x0076,
				0x0077, 0x0078, 0x0079, 0x007A, 0x00DF, 0x00E1, 0x00E4, 0x00E5, 0x00E9, 0x00EA, 0x00EB, 0x00ED, 0x00F3,
				0x00F6, 0x00FC, 0x0105, 0x0107, 0x010D, 0x0119, 0x011B, 0x012B, 0x0146, 0x0159, 0x015B, 0x0161, 0x0165,
				0x016B, 0x03AC, 0x03AD, 0x03AE, 0x03AF, 0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B9, 0x03BA, 0x03BB,
				0x03BC, 0x03BD, 0x03BE, 0x03BF, 0x03C0, 0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C6, 0x03C7, 0x03CC, 0x03CD,
				0x03CE, 0x0430, 0x0432, 0x0434, 0x0435, 0x0438, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F, 0x0440,
				0x0441, 0x0442, 0x0445, 0x0446, 0x0447, 0x0448, 0x044B, 0x044C, 0x044F, 0x0456, 0x002D, 0x0041, 0x0042,
				0x0043, 0x0044, 0x0045, 0x0046, 0x0047, 0x0048, 0x0049, 0x004A, 0x004B, 0x004C, 0x004D, 0x004E, 0x004F,
				0x0050, 0x0051, 0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057, 0x0058, 0x0059, 0x005A, 0x00C1, 0x00C4,
				0x00C5, 0x00C9, 0x00CA, 0x00CB, 0x00CD, 0x00D3, 0x00D6, 0x00DC, 0x0104, 0x0106, 0x010C, 0x0118, 0x011A,
				0x012A, 0x0145, 0x0158, 0x015A, 0x0160, 0x0164, 0x016A, 0x0386, 0x0388, 0x0389, 0x038A, 0x038C, 0x038E,
				0x038F, 0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F,
				0x03A0, 0x03A1, 0x03A3, 0x03A4, 0x03A6, 0x03A7, 0x0406, 0x0410, 0x0412, 0x0414, 0x0415, 0x0418, 0x041A,
				0x041B, 0x041C, 0x041D, 0x041E, 0x041F, 0x0420, 0x0421, 0x0422, 0x0425, 0x0426, 0x0427, 0x0428, 0x042B,
				0x042C, 0x042F];
			break;
		}
		case Asc.c_oAscNumberingFormat.Chicago:
		{
			arrCodesOfSymbols = [0x002A, 0x00A7, 0x2020, 0x2021];
			break;
		}
		case Asc.c_oAscNumberingFormat.ChineseCounting:
		{
			arrCodesOfSymbols = [0x25CB, 0x4E00, 0x4E8C, 0x4E09, 0x56DB, 0x4E94, 0x516D, 0x4E03, 0x516B, 0x4E5D, 0x5341];
			break;
		}
		case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
		{
			arrCodesOfSymbols = [0x25CB, 0x4E00, 0x4E8C, 0x4E09, 0x56DB, 0x4E94, 0x516D, 0x4E03, 0x516B, 0x4E5D, 0x5341,
				0x767E, 0x5343, 0x4E07];
			break;
		}
		case Asc.c_oAscNumberingFormat.ChineseLegalSimplified:
		{
			arrCodesOfSymbols = [0x96F6, 0x58F9, 0x8D30, 0x53C1, 0x8086, 0x4F0D, 0x9646, 0x67D2, 0x634C, 0x7396, 0x62FE,
				0x4F70, 0x4EDF, 0x842C];
			break;
		}
		case Asc.c_oAscNumberingFormat.Chosung:
		{
			arrCodesOfSymbols = [0x3131, 0x3134, 0x3137, 0x3139, 0x3141, 0x3142, 0x3145, 0x3147, 0x3148, 0x314A, 0x314B,
				0x314C, 0x314D, 0x314E];
			break;
		}
		case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle:
		{
			arrCodesOfSymbols = [0x2460, 0x2461, 0x2462, 0x2463, 0x2464, 0x2465, 0x2466, 0x2467, 0x2468, 0x2469, 0x246A,
				0x246B, 0x246C, 0x246D, 0x246E, 0x246F, 0x2470, 0x2471, 0x2472, 0x2473];
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese:
		{
			arrCodesOfSymbols = [0x2460, 0x2461, 0x2462, 0x2463, 0x2464, 0x2465, 0x2466, 0x2467, 0x2468, 0x2469];
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop:
		{
			arrCodesOfSymbols = [0x2488, 0x2489, 0x248A, 0x248B, 0x248C, 0x248D, 0x248E, 0x248F, 0x2490, 0x2491, 0x2492,
				0x2493, 0x2494, 0x2495, 0x2496, 0x2497, 0x2498, 0x2499, 0x249A, 0x249B];
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.DecimalEnclosedParen:
		{
			arrCodesOfSymbols = [0x2474, 0x2475, 0x2476, 0x2477, 0x2478, 0x2479, 0x247A, 0x247B, 0x247C, 0x247D, 0x247E,
				0x247F, 0x2480, 0x2481, 0x2482, 0x2483, 0x2484, 0x2485, 0x2486, 0x2487];
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.DecimalFullWidth2:
		case Asc.c_oAscNumberingFormat.DecimalFullWidth:
		{
			arrCodesOfSymbols = [0xFF10, 0xFF11, 0xFF12, 0xFF13, 0xFF14, 0xFF15, 0xFF16, 0xFF17, 0xFF18, 0xFF19];
			break;
		}
		case Asc.c_oAscNumberingFormat.Ganada:
		{
			arrCodesOfSymbols = [0xAC00, 0xB098, 0xB2E4, 0xB77C, 0xB9C8, 0xBC14, 0xC0AC, 0xC544, 0xC790, 0xCC28, 0xCE74,
				0xD0C0, 0xD30C, 0xD558];
			break;
		}
		case Asc.c_oAscNumberingFormat.Hebrew1:
		{
			arrCodesOfSymbols = [0x05D0, 0x05D1, 0x05D2, 0x05D3, 0x05D4, 0x05D5, 0x05D6, 0x05D7, 0x05D8, 0x05D9, 0x05DA,
				0x05DB, 0x05DC, 0x05DD, 0x05DE, 0x05DF, 0x05E0, 0x05E1, 0x05E2, 0x05E3, 0x05E4, 0x05E5, 0x05E6, 0x05E7,
				0x05E8, 0x05E9, 0x05EA];
			break;
		}
		case Asc.c_oAscNumberingFormat.Hebrew2:
		{
			arrCodesOfSymbols = [0x05D0, 0x05D1, 0x05D2, 0x05D3, 0x05D4, 0x05D5, 0x05D6, 0x05D7, 0x05D8, 0x05D9, 0x05DB,
				0x05DC, 0x05DE, 0x05E0, 0x05E1, 0x05E2, 0x05E4, 0x05E6, 0x05E7, 0x05E8, 0x05E9, 0x05EA];
			break;
		}
		case Asc.c_oAscNumberingFormat.Hex:
		{
			arrCodesOfSymbols = [0x0041, 0x0042, 0x0043, 0x0044, 0x0045, 0x0046];
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.HindiConsonants:
		{
			arrCodesOfSymbols = [0x0902, 0x0903, 0x0905, 0x0906, 0x0907, 0x0908, 0x0909, 0x090A, 0x090B, 0x090C, 0x090D,
				0x090E, 0x090F, 0x0910, 0x0911, 0x0912, 0x0913, 0x0914];
			break;
		}
		case Asc.c_oAscNumberingFormat.HindiCounting:
		{
			arrCodesOfSymbols = [0x0902, 0x0905, 0x0906, 0x0907, 0x0908, 0x0909, 0x090F, 0x0915, 0x0917, 0x091A, 0x091B,
				0x091C, 0x091F, 0x0920, 0x0921, 0x0924, 0x0926, 0x0928, 0x092A, 0x092C, 0x092F, 0x0930, 0x0932, 0x0935,
				0x0938, 0x0939, 0x093C, 0x093E, 0x093F, 0x0940, 0x0947, 0x0948, 0x094B, 0x094C, 0x094D];
			break;
		}
		case Asc.c_oAscNumberingFormat.HindiNumbers:
		{
			arrCodesOfSymbols = [0x0967, 0x0968, 0x0969, 0x096A, 0x096B, 0x096C, 0x096D, 0x096E, 0x096F];
			break;
		}
		case Asc.c_oAscNumberingFormat.HindiVowels:
		{
			arrCodesOfSymbols = [0x0915, 0x0916, 0x0917, 0x0918, 0x0919, 0x091A, 0x091B, 0x091C, 0x091D, 0x091E, 0x091F,
				0x0920, 0x0921, 0x0922, 0x0923, 0x0924, 0x0925, 0x0926, 0x0927, 0x0928, 0x0929, 0x092A, 0x092B, 0x092C,
				0x092D, 0x092E, 0x092F, 0x0930, 0x0931, 0x0932, 0x0933, 0x0934, 0x0935, 0x0936, 0x0937, 0x0938, 0x0939];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographEnclosedCircle:
		{
			arrCodesOfSymbols = [0x3220, 0x3221, 0x3222, 0x3223, 0x3224, 0x3225, 0x3226, 0x3227, 0x3228, 0x3229];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographLegalTraditional:
		{
			arrCodesOfSymbols = [0x58F9, 0x8CB3, 0x53C3, 0x8086, 0x4F0D, 0x9678, 0x67D2, 0x634C, 0x7396, 0x62FE, 0x4F70,
				0x4EDF, 0x842C];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographTraditional:
		{
			arrCodesOfSymbols = [0x4E01, 0x4E19, 0x4E59, 0x58EC, 0x5DF1, 0x5E9A, 0x620A, 0x7532, 0x7678, 0x8F9B];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographZodiac:
		{
			appendDecimal();
			arrCodesOfSymbols = [0x4E11, 0x4EA5, 0x5348, 0x536F, 0x5B50, 0x5BC5, 0x5DF3, 0x620C, 0x672A, 0x7533, 0x8FB0,
				0x9149];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographZodiacTraditional:
		{
			arrCodesOfSymbols = [0x4E01, 0x4E11, 0x4E19, 0x4E59, 0x4EA5, 0x5348, 0x536F, 0x58EC, 0x5B50, 0x5BC5, 0x5DF1,
				0x5DF3, 0x5E9A, 0x620A, 0x620D, 0x672A, 0x7532, 0x7533, 0x7678, 0x8F9B, 0x8FB0, 0x9149];
			break;
		}
		case Asc.c_oAscNumberingFormat.Iroha:
		{
			arrCodesOfSymbols = [0x30F0, 0x30F1, 0xFF66, 0xFF71, 0xFF72, 0xFF73, 0xFF74, 0xFF75, 0xFF76, 0xFF77, 0xFF78,
				0xFF79, 0xFF7A, 0xFF7B, 0xFF7C, 0xFF7D, 0xFF7E, 0xFF7F, 0xFF80, 0xFF81, 0xFF82, 0xFF83, 0xFF84, 0xFF85,
				0xFF86, 0xFF87, 0xFF88, 0xFF89, 0xFF8A, 0xFF8B, 0xFF8C, 0xFF8D, 0xFF8E, 0xFF8F, 0xFF90, 0xFF91, 0xFF92,
				0xFF93, 0xFF94, 0xFF95, 0xFF96, 0xFF97, 0xFF98, 0xFF99, 0xFF9A, 0xFF9B, 0xFF9C, 0xFF9D];
			break;
		}
		case Asc.c_oAscNumberingFormat.AiueoFullWidth:
		case Asc.c_oAscNumberingFormat.IrohaFullWidth:
		{
			arrCodesOfSymbols = [0x30A2, 0x30A4, 0x30A6, 0x30A8, 0x30AA, 0x30AB, 0x30AD, 0x30AF, 0x30B1, 0x30B3, 0x30B5,
				0x30B7, 0x30B9, 0x30BB, 0x30BD, 0x30BF, 0x30C1, 0x30C4, 0x30C6, 0x30C8, 0x30CA, 0x30CB, 0x30CC, 0x30CD,
				0x30CE, 0x30CF, 0x30D2, 0x30D5, 0x30D8, 0x30DB, 0x30DE, 0x30DF, 0x30E0, 0x30E1, 0x30E2, 0x30E4, 0x30E6,
				0x30E8, 0x30E9, 0x30EA, 0x30EB, 0x30EC, 0x30ED, 0x30EF, 0x30F0, 0x30F1, 0x30F2, 0x30F3];
			break;
		}
		case Asc.c_oAscNumberingFormat.JapaneseCounting:
		{
			arrCodesOfSymbols = [0x3007, 0x4E00, 0x4E8C, 0x4E09, 0x56DB, 0x4E94, 0x516D, 0x4E03, 0x516B, 0x4E5D, 0x5341,
				0x5343, 0x767E];
			break;
		}
		case Asc.c_oAscNumberingFormat.IdeographDigital:
		case Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand:
		{
			arrCodesOfSymbols = [0x3007, 0x4E00, 0x4E03, 0x4E09, 0x4E5D, 0x4E8C, 0x4E94, 0x516B, 0x516D, 0x56DB];
			break;
		}
		case Asc.c_oAscNumberingFormat.JapaneseLegal:
		{
			arrCodesOfSymbols = [0x58F1, 0x5F10, 0x53C2, 0x56DB, 0x4F0D, 0x516D, 0x4E03, 0x516B, 0x4E5D, 0x62FE, 0x767E,
				0x842C, 0x9621];
			break;
		}
		case Asc.c_oAscNumberingFormat.KoreanCounting:
		{
			arrCodesOfSymbols = [0xC77C, 0xC774, 0xC0BC, 0xC0AC, 0xC624, 0xC721, 0xCE60, 0xD314, 0xAD6C, 0xC2ED, 0xB9CC,
				0xCC9C, 0xBC31];
			break;
		}
		case Asc.c_oAscNumberingFormat.KoreanDigital:
		{
			arrCodesOfSymbols = [0xAD6C, 0xC0AC, 0xC0BC, 0xC601, 0xC624, 0xC721, 0xC774, 0xC77C, 0xCE60, 0xD314];
			break;
		}
		case Asc.c_oAscNumberingFormat.KoreanDigital2:
		{
			arrCodesOfSymbols = [0x4E00, 0x4E03, 0x4E09, 0x4E5D, 0x4E8C, 0x4E94, 0x516B, 0x516D, 0x56DB, 0x96F6];
			break;
		}
		case Asc.c_oAscNumberingFormat.KoreanLegal:
		{
			arrCodesOfSymbols = [0xD558, 0xB098, 0xB458, 0xC14B, 0xB137, 0xB2E4, 0xC12F, 0xC5EC, 0xC12F, 0xC77C, 0xACF1,
				0xC5EC, 0xB35F, 0xC544, 0xD649, 0xC5F4, 0xC2A4, 0xBB3C, 0xC11C, 0xB978, 0xB9C8, 0xD754, 0xC270, 0xC608,
				0xC21C, 0xC77C, 0xD754, 0xC5EC, 0xB4E0, 0xC544, 0xD754];
			break;
		}
		case Asc.c_oAscNumberingFormat.UpperRoman:
		case Asc.c_oAscNumberingFormat.UpperLetter:
		{
			arrCodesOfSymbols = [0x0041, 0x0042, 0x0043, 0x0044, 0x0045, 0x0046, 0x0047, 0x0048, 0x0049, 0x004A, 0x004B,
				0x004C, 0x004D, 0x004E, 0x004F, 0x0050, 0x0051, 0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057, 0x0058,
				0x0059, 0x005A];
			break;
		}
		case Asc.c_oAscNumberingFormat.LowerLetter:
		case Asc.c_oAscNumberingFormat.LowerRoman:

		{
			arrCodesOfSymbols = [0x0061, 0x0062, 0x0063, 0x0064, 0x0065, 0x0066, 0x0067, 0x0068, 0x0069, 0x006A, 0x006B,
				0x006C, 0x006D, 0x006E, 0x006F, 0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075, 0x0076, 0x0077, 0x0078,
				0x0079, 0x007A];
			break;
		}
		case Asc.c_oAscNumberingFormat.NumberInDash:
		{
			appendDecimal();
			arrCodesOfSymbols = [0x002D];
			break;
		}
		case Asc.c_oAscNumberingFormat.Ordinal:
		{
			appendDecimal();
			arrCodesOfSymbols = [0x002D, 0x002E, 0x003A, 0x0061, 0x0064, 0x0065, 0x0068, 0x006E, 0x0072, 0x0073, 0x0074,
				0x00B0, 0x00BA, 0x03BF, 0x0439];
			break;
		}
		case Asc.c_oAscNumberingFormat.OrdinalText:
		{
			arrCodesOfSymbols = [0x0020, 0x002D, 0x0061, 0x0062, 0x0063, 0x0064, 0x0065, 0x0066, 0x0067, 0x0068, 0x0069,
				0x006A, 0x006B, 0x006C, 0x006D, 0x006E, 0x006F, 0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075, 0x0076,
				0x0077, 0x0078, 0x0079, 0x007A, 0x00DF, 0x00E1, 0x00E4, 0x00E5, 0x00E9, 0x00EA, 0x00EB, 0x00ED, 0x00F3,
				0x00F6, 0x00FC, 0x0105, 0x0107, 0x010D, 0x0119, 0x011B, 0x012B, 0x0146, 0x0159, 0x015B, 0x0161, 0x0165,
				0x016B, 0x03AC, 0x03AD, 0x03AE, 0x03AF, 0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B9, 0x03BA, 0x03BB,
				0x03BC, 0x03BD, 0x03BE, 0x03BF, 0x03C0, 0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C6, 0x03C7, 0x03CC, 0x03CD,
				0x03CE, 0x0430, 0x0432, 0x0434, 0x0435, 0x0438, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F, 0x0440,
				0x0441, 0x0442, 0x0445, 0x0446, 0x0447, 0x0448, 0x044B, 0x044C, 0x044F, 0x0456, 0x0041, 0x0042, 0x0043,
				0x0044, 0x0045, 0x0046, 0x0047, 0x0048, 0x0049, 0x004A, 0x004B, 0x004C, 0x004D, 0x004E, 0x004F, 0x0050,
				0x0051, 0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057, 0x0058, 0x0059, 0x005A, 0x00C1, 0x00C4, 0x00C5,
				0x00C9, 0x00CA, 0x00CB, 0x00CD, 0x00D3, 0x00D6, 0x00DC, 0x0104, 0x0106, 0x010C, 0x0118, 0x011A, 0x012A,
				0x0145, 0x0158, 0x015A, 0x0160, 0x0164, 0x016A, 0x0386, 0x0388, 0x0389, 0x038A, 0x038C, 0x038E, 0x038F,
				0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F, 0x03A0,
				0x03A1, 0x03A3, 0x03A4, 0x03A6, 0x03A7, 0x0406, 0x0410, 0x0412, 0x0414, 0x0415, 0x0418, 0x041A, 0x041B,
				0x041C, 0x041D, 0x041E, 0x041F, 0x0420, 0x0421, 0x0422, 0x0425, 0x0426, 0x0427, 0x0428, 0x042B, 0x042C,
				0x042F, 0x00C8, 0x00D4, 0x00DD, 0x0397, 0x03A9, 0x0413, 0x0419, 0x0423, 0x042A, 0x00E8, 0x00F4, 0x00FD,
				0x03B7, 0x03C9, 0x0433, 0x0439, 0x0443, 0x044A];
			break;
		}
		case Asc.c_oAscNumberingFormat.RussianLower:
		{
			arrCodesOfSymbols = [0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437, 0x0438, 0x0439, 0x043A,
				0x043B, 0x043C, 0x043D, 0x043E, 0x043F, 0x0440, 0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447,
				0x0448, 0x0449, 0x044A, 0x044B, 0x044C, 0x044D, 0x044E, 0x044F];
			break;
		}
		case Asc.c_oAscNumberingFormat.RussianUpper:
		{
			arrCodesOfSymbols = [0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417, 0x0418, 0x0419, 0x041A,
				0x041B, 0x041C, 0x041D, 0x041E, 0x041F, 0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
				0x0428, 0x0429, 0x042A, 0x042B, 0x042C, 0x042D, 0x042E, 0x042F];
			break;
		}
		case Asc.c_oAscNumberingFormat.TaiwaneseCounting:
		{
			arrCodesOfSymbols = [0x4E00, 0x4E03, 0x4E09, 0x4E5D, 0x4E8C, 0x4E94, 0x516B, 0x516D, 0x5341, 0x56DB];
			break;
		}
		case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand:
		{
			arrCodesOfSymbols = [0x4E00, 0x4E8C, 0x4E09, 0x56DB, 0x4E94, 0x516D, 0x4E03, 0x516B, 0x4E5D];
			break;
		}
		case Asc.c_oAscNumberingFormat.TaiwaneseDigital:
		{
			arrCodesOfSymbols = [0x25CB, 0x4E00, 0x4E03, 0x4E09, 0x4E5D, 0x4E8C, 0x4E94, 0x516B, 0x516D, 0x56DB];
			break;
		}
		case Asc.c_oAscNumberingFormat.ThaiCounting:
		{
			arrCodesOfSymbols = [0x0E2B, 0x0E19, 0x0E36, 0x0E48, 0x0E07, 0x0E2A, 0x0E2D, 0x0E32, 0x0E21, 0x0E35, 0x0E49,
				0x0E01, 0x0E40, 0x0E08, 0x0E47, 0x0E14, 0x0E41, 0x0E1B, 0x0E34, 0x0E1A, 0x0E22, 0x0E25, 0x0E37, 0x0E23,
				0x0E1E, 0x0E31];
			break;
		}
		case Asc.c_oAscNumberingFormat.ThaiLetters:
		{
			arrCodesOfSymbols = [0x0E01, 0x0E02, 0x0E04, 0x0E25, 0x0E07, 0x0E08, 0x0E09, 0x0E0A, 0x0E0B, 0x0E0C, 0x0E0D,
				0x0E0E, 0x0E0F, 0x0E10, 0x0E11, 0x0E12, 0x0E13, 0x0E14, 0x0E15, 0x0E16, 0x0E17, 0x0E18, 0x0E19, 0x0E1A,
				0x0E1B, 0x0E1C, 0x0E1D, 0x0E1E, 0x0E1F, 0x0E20, 0x0E21, 0x0E22, 0x0E23, 0x0E27, 0x0E28, 0x0E29, 0x0E2A,
				0x0E2B, 0x0E2C, 0x0E2D, 0x0E2E];
			break;
		}
		case Asc.c_oAscNumberingFormat.ThaiNumbers:
		{
			arrCodesOfSymbols = [0x0E50, 0x0E51, 0x0E52, 0x0E53, 0x0E54, 0x0E55, 0x0E56, 0x0E57, 0x0E58, 0x0E59];
			break;
		}
		case Asc.c_oAscNumberingFormat.VietnameseCounting:
		{
			arrCodesOfSymbols = [0x0061, 0x0062, 0x0063, 0x0068, 0x0069, 0x006D, 0x006E, 0x0073, 0x0074, 0x0075, 0x0079,
				0x00E1, 0x00ED, 0x0103, 0x01B0, 0x1EA3, 0x1ED1, 0x1ED9, 0x1EDD];
			break;
		}
		case Asc.c_oAscNumberingFormat.CustomGreece:
		{
			arrCodesOfSymbols = [0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03C3, 0x03C4, 0x03B6, 0x03B7, 0x03B8, 0x03B9,
				0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF, 0x03C0, 0x03DF, 0x03C1, 0x03C3, 0x03C4, 0x03C5, 0x03C6,
				0x03C7, 0x03C8, 0x03C9, 0x03E1, 0x002C];
			break;
		}
		case Asc.c_oAscNumberingFormat.CustomDecimalTwoZero:
		case Asc.c_oAscNumberingFormat.CustomDecimalThreeZero:
		case Asc.c_oAscNumberingFormat.CustomDecimalFourZero:
		case Asc.c_oAscNumberingFormat.Custom:
		case Asc.c_oAscNumberingFormat.BahtText:
		case Asc.c_oAscNumberingFormat.Decimal:
		case Asc.c_oAscNumberingFormat.DecimalZero:
		case Asc.c_oAscNumberingFormat.DollarText:
		case Asc.c_oAscNumberingFormat.DecimalHalfWidth:
		{
			appendDecimal();
			break;
		}
		case Asc.c_oAscNumberingFormat.Bullet:
		case Asc.c_oAscNumberingFormat.None:
		default:
			break;
	}

	return arrSymbols.concat(arrCodesOfSymbols.map(AscCommon.encodeSurrogateChar)).join('');
};

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
CNumberingLvlTextString.prototype.IsEqual = function (oAnotherElement)
{
	if (this.Type !== oAnotherElement.Type)
		return false;

	if (this.Value !==  oAnotherElement.Value)
		return false;

	return true;
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
CNumberingLvlTextNum.prototype.IsEqual = function (oAnotherElement, oPr)
{
	const bIsSingleLvlPreviewPresetEquals = oPr && oPr.isSingleLvlPreviewPreset;
	if (this.Type !== oAnotherElement.Type)
		return false;

	if (!bIsSingleLvlPreviewPresetEquals && this.Value !==  oAnotherElement.Value)
		return false;

	return true;
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
//---------------------------------------------------------export---------------------------------------------------
window['AscWord'] = window['AscWord'] || {};
window["AscWord"].CNumberingLvl = CNumberingLvl;

window["AscCommonWord"] = window.AscCommonWord = window["AscCommonWord"] || {};
window["AscCommonWord"].CNumberingLvl = CNumberingLvl;
