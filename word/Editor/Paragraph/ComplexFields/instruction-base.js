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

(function(window)
{
	const FLAGS_MERGEFORMAT = 0x00000001;
	const NO_NUMERIC        = -1;
	
	/**
	 * Base class for complex field instruction
	 * @constructor
	 */
	function FieldInstructionBase()
	{
		this.ComplexField    = null;
		this.InstructionLine = "";
		this.generalSwitches = 0x0000;
		this.generalNumeric  = NO_NUMERIC;
	}
	FieldInstructionBase.prototype.Type = AscWord.fieldtype_UNKNOWN;
	FieldInstructionBase.prototype.GetType = function()
	{
		return this.Type;
	};
	FieldInstructionBase.prototype.SetComplexField = function(oComplexField)
	{
		this.ComplexField = oComplexField;
	};
	FieldInstructionBase.prototype.GetComplexField = function()
	{
		return this.ComplexField;
	};
	FieldInstructionBase.prototype.ToString = function()
	{
		let result = this.writeField();
		result += this.writeGeneralSwitches();
		return result;
	};
	FieldInstructionBase.prototype.SetPr = function()
	{
	};
	FieldInstructionBase.prototype.SetInstructionLine = function(sLine)
	{
		this.InstructionLine = sLine;
	};
	FieldInstructionBase.prototype.CheckInstructionLine = function(sLine)
	{
		return (this.InstructionLine === sLine);
	};
	FieldInstructionBase.prototype.addGeneralSwitches = function(switches)
	{
		if (!switches || !switches.length)
			return;
		
		for (let i = 0; i < switches.length; ++i)
		{
			let curSwitch = switches[i].toUpperCase();
			let numeric = NO_NUMERIC;
			if ("MERGEFORMAT" === curSwitch)
				this.generalSwitches |= FLAGS_MERGEFORMAT;
			else if (NO_NUMERIC !== (numeric = generalNumericToNumFormat(switches[i])))
				this.generalNumeric = numeric;
		}
	};
	FieldInstructionBase.prototype.isMergeFormat = function()
	{
		return !!(this.generalSwitches & FLAGS_MERGEFORMAT);
	};
	FieldInstructionBase.prototype.getNumericFormat = function()
	{
		return this.generalNumeric;
	};
	FieldInstructionBase.prototype.haveNumericFormat = function()
	{
		return this.generalNumeric !== NO_NUMERIC;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	FieldInstructionBase.prototype.writeField = function()
	{
		// This method must be overridden
	};
	FieldInstructionBase.prototype.writeGeneralSwitches = function()
	{
		let result = "";
		
		if (this.GeneralSwitches & FLAGS_MERGEFORMAT)
			result += " \\* MERGEFORMAT";
		
		if (NO_NUMERIC !== this.generalNumeric)
			result += " \\* " + numFormatToGeneralNumeric(this.generalNumeric);
		
		return result;
	};
	
	
	function generalNumericToNumFormat(value, lang)
	{
		switch (value)
		{
			case "AIUEO": return Asc.c_oAscNumberingFormat.AiueoFullWidth;
			case "ALPHABETIC": return Asc.c_oAscNumberingFormat.UpperLetter;
			case "alphabetic": return Asc.c_oAscNumberingFormat.LowerLetter;
			case "Arabic": return Asc.c_oAscNumberingFormat.Decimal;
			case "ARABICABJAD": return Asc.c_oAscNumberingFormat.ArabicAbjad;
			case "ARABICALPHA": return Asc.c_oAscNumberingFormat.ArabicAlpha;
			case "ArabicDash": return Asc.c_oAscNumberingFormat.NumberInDash;
			case "BAHTTEXT": return Asc.c_oAscNumberingFormat.BahtText;
			case "CardText": return Asc.c_oAscNumberingFormat.CardinalText;
			case "CHINESENUM1": return lang === lcid_zhTW ? Asc.c_oAscNumberingFormat.TaiwaneseCounting : Asc.c_oAscNumberingFormat.ChineseCounting;
			case "CHINESENUM2": return lang === lcid_zhTW ? Asc.c_oAscNumberingFormat.IdeographLegalTraditional : Asc.c_oAscNumberingFormat.ChineseLegalSimplified;
			case "CHINESENUM3": return lang === lcid_zhTW ? Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand : Asc.c_oAscNumberingFormat.ChineseCountingThousand;
			case "CHOSUNG": return Asc.c_oAscNumberingFormat.Chosung;
			case "CIRCLENUM": return Asc.c_oAscNumberingFormat.DecimalEnclosedCircle;
			case "DBCHAR": return Asc.c_oAscNumberingFormat.DecimalFullWidth;
			case "DBNUM1": return lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.IdeographDigital : Asc.c_oAscNumberingFormat.KoreanDigital;
			case "DBNUM2": return lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.JapaneseCounting : Asc.c_oAscNumberingFormat.KoreanCounting;
			case "DBNUM3": return lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.JapaneseLegal : Asc.c_oAscNumberingFormat.KoreanLegal;
			case "DBNUM4": return lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand : lang === lcid_koKR ? Asc.c_oAscNumberingFormat.KoreanDigital2 : Asc.c_oAscNumberingFormat.TaiwaneseDigital;
			case "DollarText": return Asc.c_oAscNumberingFormat.DollarText;
			case "GANADA": return Asc.c_oAscNumberingFormat.Ganada;
			case "GB1": return Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop;
			case "GB2": return Asc.c_oAscNumberingFormat.DecimalEnclosedParen;
			case "GB3": return Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese;
			case "GB4": return Asc.c_oAscNumberingFormat.IdeographEnclosedCircle;
			case "HEBREW1": return Asc.c_oAscNumberingFormat.Hebrew1;
			case "HEBREW2": return Asc.c_oAscNumberingFormat.Hebrew2;
			case "Hex": return Asc.c_oAscNumberingFormat.Hex;
			case "HINDIARABIC": return Asc.c_oAscNumberingFormat.HindiNumbers;
			case "HINDICARDTEXT": return Asc.c_oAscNumberingFormat.HindiCounting;
			case "HINDILETTER1": return Asc.c_oAscNumberingFormat.HindiVowels;
			case "HINDILETTER2": return Asc.c_oAscNumberingFormat.HindiConsonants;
			case "IROHA": return Asc.c_oAscNumberingFormat.IrohaFullWidth;
			case "KANJINUM1": return lang === lcid_koKR ? Asc.c_oAscNumberingFormat.KoreanDigital : lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.IdeographDigital : lang === lcid_zhCN ? Asc.c_oAscNumberingFormat.ChineseCounting :Asc.c_oAscNumberingFormat.TaiwaneseCounting;
			case "KANJINUM2": return lang === lcid_koKR ? Asc.c_oAscNumberingFormat.KoreanCounting : lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.ChineseCountingThousand : lang === lcid_zhCN ? Asc.c_oAscNumberingFormat.ChineseLegalSimplified :Asc.c_oAscNumberingFormat.IdeographLegalTraditional;
			case "KANJINUM3": return lang === lcid_koKR ? Asc.c_oAscNumberingFormat.KoreanLegal : lang === lcid_jaJP ? Asc.c_oAscNumberingFormat.JapaneseLegal : lang === lcid_zhCN ? Asc.c_oAscNumberingFormat.ChineseCountingThousand :Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand;
			case "Ordinal": return Asc.c_oAscNumberingFormat.Ordinal;
			case "OrdText": return Asc.c_oAscNumberingFormat.OrdinalText;
			case "ROMAN":
			case "Roman": return Asc.c_oAscNumberingFormat.UpperRoman;
			case "roman": return Asc.c_oAscNumberingFormat.LowerRoman;
			case "SBCHAR": return Asc.c_oAscNumberingFormat.DecimalHalfWidth;
			case "THAIARABIC": return Asc.c_oAscNumberingFormat.ThaiNumbers;
			case "THAICARDTEXT": return Asc.c_oAscNumberingFormat.ThaiCounting;
			case "THAILETTER": return Asc.c_oAscNumberingFormat.ThaiLetters;
			case "VIETCARDTEXT": return Asc.c_oAscNumberingFormat.VietnameseCounting;
			case "ZODIAC1": return Asc.c_oAscNumberingFormat.IdeographTraditional;
			case "ZODIAC2": return Asc.c_oAscNumberingFormat.IdeographZodiac;
			case "ZODIAC3": return Asc.c_oAscNumberingFormat.IdeographZodiacTraditional;
		}
		return NO_NUMERIC;
	}
	function numFormatToGeneralNumeric(format)
	{
		switch (format)
		{
			case Asc.c_oAscNumberingFormat.AiueoFullWidth: return "AIUEO";
			case Asc.c_oAscNumberingFormat.UpperLetter: return "ALPHABETIC";
			case Asc.c_oAscNumberingFormat.LowerLetter: return "alphabetic";
			case Asc.c_oAscNumberingFormat.Decimal: return "Arabic";
			case Asc.c_oAscNumberingFormat.ArabicAbjad: return "ARABICABJAD";
			case Asc.c_oAscNumberingFormat.ArabicAlpha: return "ARABICALPHA";
			case Asc.c_oAscNumberingFormat.NumberInDash: return "ArabicDash";
			case Asc.c_oAscNumberingFormat.BahtText: return "BAHTTEXT";
			case Asc.c_oAscNumberingFormat.CardinalText: return "CardText";
			
			case Asc.c_oAscNumberingFormat.TaiwaneseCounting:
			case Asc.c_oAscNumberingFormat.ChineseCounting: return "CHINESENUM1";
			
			case Asc.c_oAscNumberingFormat.IdeographLegalTraditional:
			case Asc.c_oAscNumberingFormat.ChineseLegalSimplified: return "CHINESENUM2";
			
			case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand:
			case Asc.c_oAscNumberingFormat.ChineseCountingThousand: return "CHINESENUM3";
			
			case Asc.c_oAscNumberingFormat.Chosung: return "CHOSUNG";
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle: return "CIRCLENUM";
			case Asc.c_oAscNumberingFormat.DecimalFullWidth: return "DBCHAR";
			
			case Asc.c_oAscNumberingFormat.IdeographDigital:
			case Asc.c_oAscNumberingFormat.KoreanDigital: return "DBNUM1";
			
			case Asc.c_oAscNumberingFormat.JapaneseCounting:
			case Asc.c_oAscNumberingFormat.KoreanCounting: return "DBNUM2";
			
			case Asc.c_oAscNumberingFormat.JapaneseLegal:
			case Asc.c_oAscNumberingFormat.KoreanLegal: return "DBNUM3";
			
			case Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand:
			case Asc.c_oAscNumberingFormat.KoreanDigital2:
			case Asc.c_oAscNumberingFormat.TaiwaneseDigital: return "DBNUM4";
			
			case Asc.c_oAscNumberingFormat.DollarText: return "DollarText";
			case Asc.c_oAscNumberingFormat.Ganada: return "GANADA";
			case Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop: return "GB1";
			case Asc.c_oAscNumberingFormat.DecimalEnclosedParen: return "GB2";
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese: return "GB3";
			case Asc.c_oAscNumberingFormat.IdeographEnclosedCircle: return "GB4";
			case Asc.c_oAscNumberingFormat.Hebrew1: return "HEBREW1";
			case Asc.c_oAscNumberingFormat.Hebrew2: return "HEBREW2";
			case Asc.c_oAscNumberingFormat.Hex: return "Hex";
			case Asc.c_oAscNumberingFormat.HindiNumbers: return "HINDIARABIC";
			case Asc.c_oAscNumberingFormat.HindiCounting: return "HINDICARDTEXT";
			case Asc.c_oAscNumberingFormat.HindiVowels: return "HINDILETTER1";
			case Asc.c_oAscNumberingFormat.HindiConsonants: return "HINDILETTER2";
			case Asc.c_oAscNumberingFormat.IrohaFullWidth: return "IROHA";
			
			// case Asc.c_oAscNumberingFormat.KoreanDigital:
			// case Asc.c_oAscNumberingFormat.IdeographDigital:
			// case Asc.c_oAscNumberingFormat.ChineseCounting:
			// case Asc.c_oAscNumberingFormat.TaiwaneseCounting: return "KANJINUM1";
			// case Asc.c_oAscNumberingFormat.KoreanCounting:
			// case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
			// case Asc.c_oAscNumberingFormat.ChineseLegalSimplified:
			// case Asc.c_oAscNumberingFormat.IdeographLegalTraditional: return "KANJINUM2";
			// case Asc.c_oAscNumberingFormat.KoreanLegal:
			// case Asc.c_oAscNumberingFormat.JapaneseLegal:
			// case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
			// case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand: return "KANJINUM3";
			
			case Asc.c_oAscNumberingFormat.Ordinal: return "Ordinal";
			case Asc.c_oAscNumberingFormat.OrdinalText: return "OrdText";
			case Asc.c_oAscNumberingFormat.UpperRoman: return "ROMAN";
			case Asc.c_oAscNumberingFormat.LowerRoman: return "roman";
			case Asc.c_oAscNumberingFormat.DecimalHalfWidth: return "SBCHAR";
			case Asc.c_oAscNumberingFormat.ThaiNumbers: return "THAIARABIC";
			case Asc.c_oAscNumberingFormat.ThaiCounting: return "THAICARDTEXT";
			case Asc.c_oAscNumberingFormat.ThaiLetters: return "THAILETTER";
			case Asc.c_oAscNumberingFormat.VietnameseCounting: return "VIETCARDTEXT";
			case Asc.c_oAscNumberingFormat.IdeographTraditional: return "ZODIAC1";
			case Asc.c_oAscNumberingFormat.IdeographZodiac: return "ZODIAC2";
			case Asc.c_oAscNumberingFormat.IdeographZodiacTraditional: return "ZODIAC3";
		}
		return "Arabic";
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].FieldInstructionBase = FieldInstructionBase;
	
})(window);
