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
/**
 * User: Ilja.Kirillov
 * Date: 08.05.2018
 * Time: 15:41
 */

var numbering_lvltext_Text = 1;
var numbering_lvltext_Num  = 2;

/** enum {number} */
var c_oAscMultiLevelNumbering = {
	Numbered    : 0,
	Bullet      : 1,

	MultiLevel_1_a_i           : 101,
	MultiLevel_1_11_111        : 102,
	MultiLevel_Bullet          : 103,
	MultiLevel_Article_Section : 104,
	MultiLevel_Chapter         : 105,
	MultiLevel_I_A_1           : 106,
	MultiLevel_1_11_111_NoInd  : 107,
};
window["Asc"]["c_oAscMultiLevelNumbering"] = window["Asc"].c_oAscMultiLevelNumbering = c_oAscMultiLevelNumbering;
c_oAscMultiLevelNumbering["Bullet"]      = c_oAscMultiLevelNumbering.Bullet;
c_oAscMultiLevelNumbering["Numbered"]    = c_oAscMultiLevelNumbering.Numbered;
c_oAscMultiLevelNumbering["MultiLevel1"] = c_oAscMultiLevelNumbering.MultiLevel1;
c_oAscMultiLevelNumbering["MultiLevel2"] = c_oAscMultiLevelNumbering.MultiLevel2;
c_oAscMultiLevelNumbering["MultiLevel3"] = c_oAscMultiLevelNumbering.MultiLevel3;

/** enum {number} */
var c_oAscNumberingLevel = {
	None       : 0,
	Bullet     : 0x1000,
	Numbered   : 0x2000,

	DecimalBracket_Right      : 0x2001,
	DecimalBracket_Left       : 0x2002,
	DecimalDot_Right          : 0x2003,
	DecimalDot_Left           : 0x2004,
	UpperRomanDot_Right       : 0x2005,
	UpperLetterDot_Left       : 0x2006,
	LowerLetterBracket_Left   : 0x2007,
	LowerLetterDot_Left       : 0x2008,
	LowerRomanDot_Right       : 0x2009,
	UpperRomanBracket_Left    : 0x200A,
	LowerRomanBracket_Left    : 0x200B,
	UpperLetterBracket_Left   : 0x200C,
	LowerRussian_Dot_Left     : 0x3001,
	LowerRussian_Bracket_Left : 0x3002,
	UpperRussian_Dot_Left     : 0x3003,
	UpperRussian_Bracket_Left : 0x3004
};

window["Asc"]["asc_oAscNumberingLevel"]           = window["Asc"].c_oAscNumberingLevel = c_oAscNumberingLevel;
c_oAscNumberingLevel["None"]                      = c_oAscNumberingLevel.None;
c_oAscNumberingLevel["Bullet"]                    = c_oAscNumberingLevel.Bullet;
c_oAscNumberingLevel["Numbered"]                  = c_oAscNumberingLevel.Numbered;
c_oAscNumberingLevel["DecimalBracket_Right"]      = c_oAscNumberingLevel.DecimalBracket_Right;
c_oAscNumberingLevel["DecimalBracket_Left"]       = c_oAscNumberingLevel.DecimalBracket_Left;
c_oAscNumberingLevel["DecimalDot_Right"]          = c_oAscNumberingLevel.DecimalDot_Right;
c_oAscNumberingLevel["DecimalDot_Left"]           = c_oAscNumberingLevel.DecimalDot_Left;
c_oAscNumberingLevel["UpperRomanDot_Right"]       = c_oAscNumberingLevel.UpperRomanDot_Right;
c_oAscNumberingLevel["UpperLetterDot_Left"]       = c_oAscNumberingLevel.UpperLetterDot_Left;
c_oAscNumberingLevel["LowerLetterBracket_Left"]   = c_oAscNumberingLevel.LowerLetterBracket_Left;
c_oAscNumberingLevel["LowerLetterDot_Left"]       = c_oAscNumberingLevel.LowerLetterDot_Left;
c_oAscNumberingLevel["LowerRomanDot_Right"]       = c_oAscNumberingLevel.LowerRomanDot_Right;
c_oAscNumberingLevel["UpperRomanBracket_Left"]    = c_oAscNumberingLevel.UpperRomanBracket_Left;
c_oAscNumberingLevel["LowerRomanBracket_Left"]    = c_oAscNumberingLevel.LowerRomanBracket_Left;
c_oAscNumberingLevel["UpperLetterBracket_Left"]   = c_oAscNumberingLevel.UpperLetterBracket_Left;
c_oAscNumberingLevel["LowerRussian_Dot_Left"]     = c_oAscNumberingLevel.LowerRussian_Dot_Left;
c_oAscNumberingLevel["LowerRussian_Bracket_Left"] = c_oAscNumberingLevel.LowerRussian_Bracket_Left;
c_oAscNumberingLevel["UpperRussian_Dot_Left"]     = c_oAscNumberingLevel.UpperRussian_Dot_Left;
c_oAscNumberingLevel["UpperRussian_Bracket_Left"] = c_oAscNumberingLevel.UpperRussian_Bracket_Left;

// Преобразовываем число в буквенную строку :
//  1 -> a
//  2 -> b
//   ...
// 26 -> z
// 27 -> aa
//   ...
// 52 -> zz
// 53 -> aaa
//   ...
function Numbering_Number_To_Alpha(Num, bLowerCase)
{
	var _Num = Num - 1;
	var Count = (_Num - _Num % 26) / 26;
	var Ost   = _Num % 26;

	var T = "";

	var Letter;
	if ( true === bLowerCase )
		Letter = String.fromCharCode( Ost + 97 );
	else
		Letter = String.fromCharCode( Ost + 65 );

	for ( var Index2 = 0; Index2 < Count + 1; Index2++ )
		T += Letter;

	return T;
}

// Преобразовываем число в обычную строку :
function Numbering_Number_To_String(Num)
{
	return "" + Num;
}

// Преобразовываем число в римскую систему исчисления :
//    1 -> i
//    4 -> iv
//    5 -> v
//    9 -> ix
//   10 -> x
//   40 -> xl
//   50 -> l
//   90 -> xc
//  100 -> c
//  400 -> cd
//  500 -> d
//  900 -> cm
// 1000 -> m
function Numbering_Number_To_Roman(Num, bLowerCase)
{
	// Переводим число Num в римскую систему исчисления
	var Rims;

	if ( true === bLowerCase )
		Rims = [  'm', 'cm', 'd', 'cd', 'c', 'xc', 'l', 'xl', 'x', 'ix', 'v', 'iv', 'i', ' '];
	else
		Rims = [  'M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I', ' '];

	var Vals = [ 1000,  900, 500,  400, 100,   90,  50,   40,  10,    9,   5,    4,   1,   0];

	var T = "";
	var Index2 = 0;
	while ( Num > 0 )
	{
		while ( Vals[Index2] <= Num )
		{
			T   += Rims[Index2];
			Num -= Vals[Index2];
		}

		Index2++;

		if ( Index2 >= Rims.length )
			break;
	}

	return T;
}

(function(window)
{
	/**
	 * Получаем набор символов используемых в заданной нумерации
	 * @param numInfo {object} нумерация заданная в виде объекта
	 * @returns {string}
	 */
	function GetNumberingSymbols(numInfo)
	{
		if (!numInfo || !numInfo.Lvl || !Array.isArray(numInfo.Lvl))
			return "";

		let symbols = "";
		for (let ilvl = 0, count = numInfo.Lvl.length; ilvl < count; ++ilvl)
		{
			let numLvl = AscWord.CNumberingLvl.FromJson(numInfo.Lvl[ilvl]);
			if (numLvl)
				symbols += numLvl.GetSymbols();
		}

		return symbols;
	}
	/**
	 * Получаем объект нумерации по устаревшим типам
	 * @param type {number}
	 * @param subtype {number}
	 * @returns {object}
	 */
	function GetNumberingObjectByDeprecatedTypes(type, subtype)
	{
		// Во всех случаях SubType = 0 означает, что нажали просто на кнопку
		// c выбором типа списка, без выбора подтипа.
		//
		// Маркированный список Type = 0
		// нет          - SubType = -1
		// черная точка - SubType = 1
		// круг         - SubType = 2
		// квадрат      - SubType = 3
		// картинка     - SubType = -1
		// 4 ромба      - SubType = 4
		// ч/б стрелка  - SubType = 5
		// галка        - SubType = 6
		// ромб         - SubType = 7
		// минус        - SubType = 8
		//
		// Нумерованный список Type = 1
		// нет - SubType = -1
		// 1.  - SubType = 1
		// 1)  - SubType = 2
		// I.  - SubType = 3
		// A.  - SubType = 4
		// a)  - SubType = 5
		// a.  - SubType = 6
		// i.  - SubType = 7
		// Ж.  - SubType = 8
		// Ж)  - SubType = 9
		// ж.  - SubType = 10
		// ж)  - SubType = 11
		//
		//
		// Многоуровневый список Type = 2
		// нет           - SubType = -1
		// 1)a)i)        - SubType = 1
		// 1.1.1         - SubType = 2
		// маркированный - SubType = 3
		//
		// SubType = 4: (Headings)
		//  Article I
		// 	 Section 1.01
		// 		 (a)
		//
		// SubType = 5 : (Headings)
		// 1.1.1  (аналогичен SubType = 2)
		//
		// SubType = 6: (Headings)
		// I. A. 3. a) (1)
		//
		// SubType = 7: (Headings)
		//  Chapter 1
		//  (none)
		//  (none)



		let value = "";

		if (-1 === subtype)
		{
			value = '{"Type":"remove"}';
		}
		else if (0 === type)
		{
			switch (subtype)
			{
				case 0: value = '{"Type":"bullet"}'; break;
				case 1: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"·","rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}}]}'; break;
				case 2: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"o","rPr":{"rFonts":{"ascii":"Courier New","cs":"Courier New","eastAsia":"Courier New","hAnsi":"Courier New"}}}]}'; break;
				case 3: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"§","rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}}]}'; break;
				case 4: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"v","rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}}]}'; break;
				case 5: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"Ø","rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}}]}'; break;
				case 6: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"ü","rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}}]}'; break;
				case 7: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"¨","rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}}]}'; break;
				case 8: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"–","rPr":{"rFonts":{"ascii":"Arial","cs":"Arial","eastAsia":"Arial","hAnsi":"Arial"}}}]}'; break;
			}
		}
		else if (1 === type)
		{
			switch (subtype)
			{
				case 0: value = '{"Type":"number"}'; break;
				case 1: value = '{"Type":"number","Lvl":[{"lvlJc":"right","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1."}]}'; break;
				case 2: value = '{"Type":"number","Lvl":[{"lvlJc":"right","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1)"}]}'; break;
				case 3: value = '{"Type":"number","Lvl":[{"lvlJc":"right","suff":"tab","numFmt":{"val":"upperRoman"},"lvlText":"%1."}]}'; break;
				case 4: value = '{"Type":"number","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"upperLetter"},"lvlText":"%1."}]}'; break;
				case 5: value = '{"Type":"number","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%1)"}]}'; break;
				case 6: value = '{"Type":"number","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%1."}]}'; break;
				case 7: value = '{"Type":"number","Lvl":[{"lvlJc":"right","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%1."}]}'; break;
				case 8: value = '{"Type":"number","Lvl":[{"lvlJc":"left","numFmt":{"val":"russianUpper"},"lvlText":"%1."}]}'; break;
				case 9: value = '{"Type":"number","Lvl":[{"lvlJc":"left","numFmt":{"val":"russianUpper"},"lvlText":"%1)"}]}'; break;
				case 10: value = '{"Type":"number","Lvl":[{"lvlJc":"left","numFmt":{"val":"russianLower"},"lvlText":"%1."}]}'; break;
				case 11: value = '{"Type":"number","Lvl":[{"lvlJc":"left","numFmt":{"val":"russianLower"},"lvlText":"%1)"}]}'; break;
			}
		}
		else if (2 === type)
		{
			switch (subtype)
			{
				case 1: value = '{"Type":"number","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1)","pPr":{"ind":{"left":360,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%2)","pPr":{"ind":{"left":720,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%3)","pPr":{"ind":{"left":1080,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%4)","pPr":{"ind":{"left":1440,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%5)","pPr":{"ind":{"left":1800,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%6)","pPr":{"ind":{"left":2160,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%7)","pPr":{"ind":{"left":2520,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%8)","pPr":{"ind":{"left":2880,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%9)","pPr":{"ind":{"left":3240,"firstLine":-360}}}]}'; break;
				case 2: value = '{"Type":"number","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.","pPr":{"ind":{"left":360,"firstLine":-360}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.","pPr":{"ind":{"left":792,"firstLine":-432}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.","pPr":{"ind":{"left":1224,"firstLine":-504}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.","pPr":{"ind":{"left":1728,"firstLine":-648}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.","pPr":{"ind":{"left":2232,"firstLine":-792}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.","pPr":{"ind":{"left":2736,"firstLine":-936}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.","pPr":{"ind":{"left":3240,"firstLine":-1080}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.%8.","pPr":{"ind":{"left":3744,"firstLine":-1224}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.%8.%9.","pPr":{"ind":{"left":4320,"firstLine":-1440}}}]}'; break;
				case 3: value = '{"Type":"bullet","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"v","pPr":{"ind":{"left":360,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"Ø","pPr":{"ind":{"left":720,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"§","pPr":{"ind":{"left":1080,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"·","pPr":{"ind":{"left":1440,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"¨","pPr":{"ind":{"left":1800,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"Ø","pPr":{"ind":{"left":2160,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"§","pPr":{"ind":{"left":2520,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Wingdings","cs":"Wingdings","eastAsia":"Wingdings","hAnsi":"Wingdings"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"·","pPr":{"ind":{"left":2880,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"bullet"},"lvlText":"¨","pPr":{"ind":{"left":3240,"firstLine":-360}},"rPr":{"rFonts":{"ascii":"Symbol","cs":"Symbol","eastAsia":"Symbol","hAnsi":"Symbol"}}}]}'; break;
				case 4: value = '{"Type":"number","Headings":"true","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"upperRoman"},"lvlText":"Article %1.","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimalZero"},"lvlText":"Section %1.%2","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"(%3)","pPr":{"ind":{"left":720,"firstLine":-432}}},{"lvlJc":"right","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"(%4)","pPr":{"ind":{"left":864,"firstLine":-144}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%5)","pPr":{"ind":{"left":1008,"firstLine":-432}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%6)","pPr":{"ind":{"left":1152,"firstLine":-432}}},{"lvlJc":"right","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%7)","pPr":{"ind":{"left":1296,"firstLine":-288}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%8.","pPr":{"ind":{"left":1440,"firstLine":-432}}},{"lvlJc":"right","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"%9.","pPr":{"ind":{"left":1584,"firstLine":-144}}}]}'; break;
				case 5: value = '{"Type":"number","Headings":"true","Lvl":[{"lvlJc":"left","suff":"space","numFmt":{"val":"decimal"},"lvlText":"Chapter %1","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"nothing","numFmt":{"val":"none"},"lvlText":"","pPr":{"ind":{"left":0,"firstLine":0}}}]}'; break;
				case 6: value = '{"Type":"number","Headings":"true","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"upperRoman"},"lvlText":"%1.","pPr":{"ind":{"left":0,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"upperLetter"},"lvlText":"%2.","pPr":{"ind":{"left":720,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%3.","pPr":{"ind":{"left":1440,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"%4)","pPr":{"ind":{"left":2160,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"(%5)","pPr":{"ind":{"left":2880,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"(%6)","pPr":{"ind":{"left":3600,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"(%7)","pPr":{"ind":{"left":4320,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerLetter"},"lvlText":"(%8)","pPr":{"ind":{"left":5040,"firstLine":0}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"lowerRoman"},"lvlText":"(%9)","pPr":{"ind":{"left":5760,"firstLine":0}}}]}'; break;
				case 7: value = '{"Type":"number","Headings":"true","Lvl":[{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.","pPr":{"ind":{"left":432,"firstLine":-432}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.","pPr":{"ind":{"left":576,"firstLine":-576}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.","pPr":{"ind":{"left":720,"firstLine":-720}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.","pPr":{"ind":{"left":864,"firstLine":-864}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.","pPr":{"ind":{"left":1008,"firstLine":-1008}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.","pPr":{"ind":{"left":1152,"firstLine":-1152}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.","pPr":{"ind":{"left":1296,"firstLine":-1296}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.%8.","pPr":{"ind":{"left":1440,"firstLine":-1440}}},{"lvlJc":"left","suff":"tab","numFmt":{"val":"decimal"},"lvlText":"%1.%2.%3.%4.%5.%6.%7.%8.%9.","pPr":{"ind":{"left":1584,"firstLine":-1584}}}]}'; break;
			}
		}

		let numObject = {};

		try
		{
			numObject = JSON.parse(value);
		}
		catch (e)
		{
		}

		return numObject;
	}
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].GetNumberingSymbols                 = GetNumberingSymbols;
	window["AscWord"].GetNumberingObjectByDeprecatedTypes = GetNumberingObjectByDeprecatedTypes;

})(window);
