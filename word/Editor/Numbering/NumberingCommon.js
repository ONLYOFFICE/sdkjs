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

		return AscWord.CNumInfo.Parse(value);
	}
	/**
	 * Проверяем является ли данный тип списка маркированным
	 * @param {Asc.c_oAscNumberingFormat} type
	 * @returns {boolean}
	 */
	function IsBulletedNumbering(type)
	{
		return (Asc.c_oAscNumberingFormat.Bullet === type || Asc.c_oAscNumberingFormat.None === type);
	}
	/**
	 * Проверяем является ли данный тип списка нумерованным
	 * @param {Asc.c_oAscNumberingFormat} type
	 * @returns {boolean}
	 */
	function IsNumberedNumbering(type)
	{
		return !IsBulletedNumbering(type);
	}
	
	/**
	 * Получаем набор символов используемых в заданном формате
	 * @param format {Asc.c_oAscNumberingFormat}
	 * @returns {string}
	 */
	function GetNumberingSymbolsByFormat(format)
	{
		let arrSymbols = [];
		function appendDecimal() {
			arrSymbols.push("0123456789");
		}
		
		let arrCodesOfSymbols = [];
		switch (format)
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
			case Asc.c_oAscNumberingFormat.DollarText:
			case Asc.c_oAscNumberingFormat.OrdinalText:
			case Asc.c_oAscNumberingFormat.CardinalText:
			{
				appendDecimal();
				arrCodesOfSymbols = [0x0020, 0x0027, 0x002D, 0x0041, 0x0042, 0x0043, 0x0044, 0x0045, 0x0046, 0x0047, 0x0048,
					0x0049, 0x004A, 0x004B, 0x004C, 0x004D, 0x004E, 0x004F, 0x0050, 0x0051, 0x0052, 0x0053, 0x0054, 0x0055,
					0x0056, 0x0057, 0x0058, 0x0059, 0x005A, 0x0061, 0x0062, 0x0063, 0x0064, 0x0065, 0x0066, 0x0067, 0x0068,
					0x0069, 0x006A, 0x006B, 0x006C, 0x006D, 0x006E, 0x006F, 0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075,
					0x0076, 0x0077, 0x0078, 0x0079, 0x007A, 0x00C1, 0x00C4, 0x00C5, 0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CD,
					0x00D3, 0x00D4, 0x00D6, 0x00DC, 0x00DD, 0x00DF, 0x00E1, 0x00E4, 0x00E5, 0x00E7, 0x00E8, 0x00E9, 0x00EA,
					0x00EB, 0x00ED, 0x00F3, 0x00F4, 0x00F6, 0x00FC, 0x00FD, 0x0104, 0x0105, 0x0106, 0x0107, 0x010C, 0x010D,
					0x0118, 0x0119, 0x011A, 0x011B, 0x012A, 0x012B, 0x0130, 0x0131, 0x0145, 0x0146, 0x0158, 0x0159, 0x015A,
					0x015B, 0x015F, 0x0160, 0x0161, 0x0164, 0x0165, 0x016A, 0x016B, 0x0386, 0x0388, 0x0389, 0x038A, 0x038C,
					0x038E, 0x038F, 0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0397, 0x0399, 0x039A, 0x039B, 0x039C, 0x039D,
					0x039E, 0x039F, 0x03A0, 0x03A1, 0x03A3, 0x03A4, 0x03A6, 0x03A7, 0x03A9, 0x03AC, 0x03AD, 0x03AE, 0x03AF,
					0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B7, 0x03B9, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF,
					0x03C0, 0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C6, 0x03C7, 0x03C9, 0x03CC, 0x03CD, 0x03CE, 0x0406, 0x0408,
					0x0410, 0x0412, 0x0413, 0x0414, 0x0415, 0x0418, 0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F,
					0x0420, 0x0421, 0x0422, 0x0423, 0x0425, 0x0426, 0x0427, 0x0428, 0x042A, 0x042B, 0x042C, 0x042F, 0x0430,
					0x0432, 0x0433, 0x0434, 0x0435, 0x0437, 0x0438, 0x0439, 0x043A, 0x043B, 0x043C, 0x043D, 0x043E, 0x043F,
					0x0440, 0x0441, 0x0442, 0x0443, 0x0445, 0x0446, 0x0447, 0x0448, 0x044A, 0x044B, 0x044C, 0x044F, 0x0456,
					0x0458, 0x045B];
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
			case Asc.c_oAscNumberingFormat.BahtText:
			case Asc.c_oAscNumberingFormat.ThaiCounting:
			{
				arrCodesOfSymbols = [0x0E01, 0x0E04, 0x0E07, 0x0E08, 0x0E14, 0x0E15, 0x0E16, 0x0E17, 0x0E19, 0x0E1A, 0x0E1B,
					0x0E1E, 0x0E21, 0x0E22, 0x0E23, 0x0E25, 0x0E27, 0x0E2A, 0x0E2B, 0x0E2D, 0x0E31, 0x0E32, 0x0E34, 0x0E35,
					0x0E36, 0x0E37, 0x0E40, 0x0E41, 0x0E47, 0x0E48, 0x0E49, 0x0E4C];
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
			case Asc.c_oAscNumberingFormat.CustomUpperTurkish:
			{
				arrCodesOfSymbols = [0x0041, 0x0042, 0x0043, 0x00c7, 0x0044, 0x0045, 0x0046, 0x0047, 0x011e, 0x0048, 0x0049, 0x0130, 0x004a,
					0x004b, 0x004c, 0x004d, 0x004e, 0x004f, 0x00d6, 0x0050, 0x0052, 0x0053, 0x015e, 0x0054, 0x0055, 0x00dc, 0x0056,
					0x0059, 0x005a];
				break;
			}
			case Asc.c_oAscNumberingFormat.CustomLowerTurkish:
			{
				arrCodesOfSymbols = [0x0061, 0x0062, 0x0063, 0x00e7, 0x0064, 0x0065, 0x0066, 0x0067, 0x011f, 0x0068, 0x0131, 0x0069, 0x006a,
					0x006b, 0x006c, 0x006d, 0x006e, 0x006f, 0x00f6, 0x0070, 0x0072, 0x0073, 0x015f, 0x0074, 0x0075, 0x00fc, 0x0076,
					0x0079, 0x007a];
				break;
			}
			case Asc.c_oAscNumberingFormat.CustomDecimalTwoZero:
			case Asc.c_oAscNumberingFormat.CustomDecimalThreeZero:
			case Asc.c_oAscNumberingFormat.CustomDecimalFourZero:
			case Asc.c_oAscNumberingFormat.Custom:
			case Asc.c_oAscNumberingFormat.Decimal:
			case Asc.c_oAscNumberingFormat.DecimalZero:
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
	}
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].GetNumberingSymbols                 = GetNumberingSymbols;
	window["AscWord"].GetNumberingSymbolsByFormat         = GetNumberingSymbolsByFormat;
	window["AscWord"].GetNumberingObjectByDeprecatedTypes = GetNumberingObjectByDeprecatedTypes;
	window["AscWord"].IsBulletedNumbering                 = IsBulletedNumbering;
	window["AscWord"].IsNumberedNumbering                 = IsNumberedNumbering;

})(window);
