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

function CNumberingLvl()
{
	this.Jc      = AscCommon.align_Left;
	this.Format  = Asc.c_oAscNumberingFormat.Bullet;
	this.PStyle  = undefined;
	this.Start   = 0;
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
				arrResult.push(AscCommon.IntToNumberFormat(nNum, nFormat, {lang: oLang}));
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
	let relatedLvl = [];
	for (var nIndex = 0, nCount = this.LvlText.length; nIndex < nCount; ++nIndex)
	{
		if (numbering_lvltext_Num !== this.LvlText[nIndex].Type)
			continue;
		
		var nLvl  = this.LvlText[nIndex].Value;
		let insertIndex = 0;
		for (let lvlCount = relatedLvl.length; insertIndex < lvlCount; ++insertIndex)
		{
			if (relatedLvl[insertIndex] < nLvl)
				continue;
			
			if (relatedLvl[insertIndex] === nLvl)
				insertIndex = -1;
			
			break;
		}
		
		if (insertIndex === relatedLvl.length)
			relatedLvl.push(nLvl);
		else if (-1 !== insertIndex)
			relatedLvl.splice(insertIndex, 0, nLvl);
	}
	
	return relatedLvl;
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
	
	return arrSymbols.join('') + AscWord.GetNumberingSymbolsByFormat(this.Format);
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
