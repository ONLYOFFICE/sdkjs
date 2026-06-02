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

(/**
* @param {Window} window
* @param {undefined} undefined
*/
function(window, undefined) {
// Import
var CellValueType = AscCommon.CellValueType;
var c_oAscNumFormatType = Asc.c_oAscNumFormatType;
var g_aCultureInfos = AscCommon.g_aCultureInfos;
var g_aAdditionalCurrencySymbols = AscCommon.g_aAdditionalCurrencySymbols;
var c_oAscDateFormatExcel = AscCommon.c_oAscDateFormatExcel;
var c_oAscTimeFormatExcel = AscCommon.c_oAscTimeFormatExcel;

var gc_sFormatDecimalPoint = ".";
var gc_sFormatThousandSeparator = ",";
var LocaleFormatSymbol ={};
var numFormat_Text = 0;
var numFormat_TextPlaceholder = 1;
var numFormat_Bracket = 2;
var numFormat_Digit = 3;
var numFormat_DigitNoDisp = 4;
var numFormat_DigitSpace = 5;
var numFormat_DecimalPoint = 6;
var numFormat_DecimalFrac = 7;
var numFormat_Thousand = 8;
var numFormat_Scientific = 9;
var numFormat_Repeat = 10;
var numFormat_Skip = 11;
var numFormat_Year = 12;
var numFormat_Month = 13;
var numFormat_Minute = 14;
var numFormat_Hour = 15;
var numFormat_Day = 16;
var numFormat_Second = 17;
var numFormat_Milliseconds = 18;
var numFormat_AmPm = 19;
var numFormat_DateSeparator = 20;
var numFormat_TimeSeparator = 21;
var numFormat_DecimalPointText = 22;
//Helper types that will be replaced in _prepareFormat
var numFormat_MonthMinute = 101;
var numFormat_Percent = 102;
var numFormat_General = 103;
var numFormat_DigitDrop = 104;
var numFormat_Plus = 105;
var numFormat_Minus = 106;
var numFormat_ThousandText = 107;
var numFormat_JapanEra = 108;
var numFormat_JapanEraYear = 109;
var numFormat_DayOfWeek = 110;

var FormatStates = {Decimal: 1, Frac: 2, Scientific: 3, Slash: 4, SlashFrac: 5};
var SignType = {Negative: 1, Null:2, Positive: 3};

var gc_nMaxDigCount = 15;//Maximum number of precision digits
var gc_nMaxDigCountView = 11;//Maximum number of digits in a cell
var gc_nMaxMantissa = Math.pow(10, gc_nMaxDigCount);
var gc_log10of2 = Math.log10(2); // log10(2): converts binary exponent to decimal exponent

// Pre-allocated buffers for IEEE 754 bit-level exponent extraction in getNumberParts.
// Avoids Math.log() floating-point precision issues and is slightly faster.
const _g_numPartsBuf = new ArrayBuffer(8);
const _g_numPartsF64 = new Float64Array(_g_numPartsBuf);
const _g_numPartsI32 = new Int32Array(_g_numPartsBuf);

// Set of Arabic locale LCIDs used for the RLM (right-to-left mark) check in format().
// Defined at module level so it is created once, not on every format() call.
const _g_arabicLCIDs = new Set([
	lcid_ar, lcid_arSY, lcid_arSA, lcid_arAE,
	lcid_arBH, lcid_arDZ, lcid_arEG, lcid_arIQ,
	lcid_arJO, lcid_arKW, lcid_arQA
]);

// Module-level helper used by NumFormat.prototype.toString().
// Converts an array of digit-type format items back to a format string fragment.
// Extracted from the inner closure to avoid re-creating it on every toString() call.
function _formatDigitArrayToString(aFormat) {
	var res = "";
	for (var i = 0, length = aFormat.length; i < length; ++i) {
		var item = aFormat[i];
		if (numFormat_Digit === item.type) {
			res += (null != item.val) ? item.val : "0";
		} else if (numFormat_DigitNoDisp === item.type) {
			res += "#";
		} else if (numFormat_DigitSpace === item.type) {
			res += "?";
		} else if (numFormat_DigitDrop === item.type) {
			res += "x";
		}
	}
	return res;
}
var gc_aTimeFormats = ['[$-F400]h:mm:ss AM/PM', 'h:mm;@', 'h:mm AM/PM;@', 'h:mm:ss;@', 'h:mm:ss AM/PM;@', 'mm:ss.0;@',
	'[h]:mm:ss;@'];
var gc_aFractionFormats = ['#\\ ?/?', '#\\ ??/??', '#\\ ???/???', '#\\ ?/2', '#\\ ?/4', '#\\ ?/8', '#\\ ??/16', '#\\ ?/10', '#\\ ??/100'];
//important for shortcuts
var gc_oParseDateOverrideFormats = {
    "d-mmm": 1,
    "d-mmm-yy": 1,
    "mmm-yy": 1,
    "h:mm": 1,
    "h:mm AM/PM": 1,
    "h:mm:ss": 1,
    "h:mm:ss AM/PM": 1,
    "mm:ss.0": 1,
    "[h]:mm:ss": 1
};
const dBNum1Numbers = ['\u3007','\u4E00','\u4E8C','\u4E09','\u56DB','\u4E94','\u516D','\u4E03','\u516B','\u4E5D'];
const interfaceFormatScientific = '0.00E+00';
const interfaceFormatPercent = '0.00%';
let interfaceShortDateFormat = 'm/d/yyyy';

var NumComporationOperators =
{
	equal: 1,
	greater: 2,
	less: 3,
	greaterorequal: 4,
	lessorequal: 5,
	notequal: 6
};
var NumFormatType =
{
	Excel: 1,
	WordFieldDate: 2,
	WordFieldNumeric: 3,
	PDFFormDate: 4
};

function getNumberParts(x)
{
    var sig = SignType.Null;
    if (!isFinite(x))
        x = 0;
	if(x > 0)
		sig = SignType.Positive;
	else if(x < 0)
	{
		sig = SignType.Negative;
		x = -x;
	}
    var exp = -gc_nMaxDigCount;
	var man = 0;
	if(SignType.Null != sig)
	{
		// Extract the base-2 exponent directly from the IEEE 754 representation.
		// Bits [62:52] of a float64 are the biased exponent; subtract the bias 1023
		// to get floor(log2(x)).  Multiplying by log10(2) ≈ 0.30103 gives an
		// approximation of floor(log10(x)) that is always <= the true value.
		// When it is one too small, man ends up >= gc_nMaxMantissa and the
		// existing correction below fixes it.
		_g_numPartsF64[0] = x;
		const binaryExp = ((_g_numPartsI32[1] >>> 20) & 0x7FF) - 1023;
		exp = Math.floor(binaryExp * gc_log10of2) - gc_nMaxDigCount + 1;
		man = Math.round(x / Math.pow(10, exp));
		if(man >= gc_nMaxMantissa)
		{
			exp++;
			man = Math.floor(man / 10);
		}
	}
    return {mantissa: man, exponent: exp, sign: sig};//for 0.123 exponent == - gc_nMaxDigCount
}

	function compareNumbers(val1, val2) {
		var res = 0;
		var parts1 = getNumberParts(val1);
		var parts2 = getNumberParts(val2);
		if (parts1.sign === parts2.sign) {
			if (parts1.exponent === parts2.exponent) {
				res = parts1.mantissa - parts2.mantissa;
			} else {
				res = parts1.exponent - parts2.exponent;
			}
			if (SignType.Negative === parts1.sign) {
				res = -res;
			}
		} else {
			res = parts1.sign - parts2.sign;
		}
		return res;
	}

    function isNumber(n) {
        return !isNaN(parseFloat(n)) && isFinite(n);
    }
	function round10(value, exp1, exp2) {
		//todo use Math.round10
		// Shift
		value = value.toString().split('e');
		value = Math.round(+(value[0] + 'e' + (value[1] ? (+value[1] + exp1) : exp1)));
		// Shift back
		value = value.toString().split('e');
		return +(value[0] + 'e' + (value[1] ? (+value[1] - exp2) : -exp2));
	}

function FormatObj(type, val)
{
    this.type = type;
    this.val = val;//what is stored here is determined by the type
}
function FormatObjScientific(val, format, sign)
{
    this.type = numFormat_Scientific;
    this.val = val;//E or e
    this.format = format;//format array
    this.sign = sign;
}
function FormatObjDecimalFrac(aLeft, aRight)
{
    this.type = numFormat_DecimalFrac;
    this.aLeft = aLeft;//left part format array
    this.aRight = aRight;//right part format array
    this.bNumRight = false;
	this.numerator = 0;
	this.denominator = 0;
}
function FormatObjDateVal(type, nCount, bElapsed)
{
    this.type = type;
    this.val = nCount;//Number of consecutive characters
    this.bElapsed = bElapsed;//true == [hhh]; in square brackets
}
//Parses the body of a [$...] locale modifier into target fields.
//Grammar: <currency>?(-<lid|bcp47>(,<calId>)?(-x-<priv>)?)?
//Fills: CurrencyString, Lid, LidName, CalendarId, bGannen.
function parseBracketLocaleModifier(data, target)
{
	var rest = data.substring(1);
	var dashIdx = rest.indexOf('-');
	var localeStr = null;
	if (dashIdx === -1) {
		if (rest.length > 0) {
			target.CurrencyString = rest;
		}
		return;
	}
	if (dashIdx > 0) {
		target.CurrencyString = rest.substring(0, dashIdx);
	}
	localeStr = rest.substring(dashIdx + 1);
	if (!localeStr || localeStr.length === 0) {
		return;
	}
	var tokens = localeStr.split(/[-,]/);
	var ti = 0;
	if (ti < tokens.length) {
		var first0 = tokens[ti];
		if (/^[0-9a-fA-F]{1,8}$/.test(first0)) {
			if (first0.toLowerCase() === "87f70000") {
				target.Lid = "411";
				target.LidName = "ja-jp";
				target.bGannen = true;
			} else {
				target.Lid = first0;
			}
			ti++;
		} else if (/^[a-zA-Z]{2,3}$/.test(first0)) {
			var lang = first0;
			ti++;
			//Region: 2-4 letters or 3 digits, but not 'x' (private-use marker)
			if (ti < tokens.length
				&& /^([a-zA-Z]{2,4}|\d{3})$/.test(tokens[ti])
				&& tokens[ti].toLowerCase() !== "x") {
				lang += "-" + tokens[ti];
				ti++;
			}
			var langLower = lang.toLowerCase();
			target.LidName = langLower;
			var lcid = null;
			if (typeof Asc !== "undefined" && Asc.g_oLcidNameToIdMap) {
				lcid = Asc.g_oLcidNameToIdMap[lang] || Asc.g_oLcidNameToIdMap[langLower];
			}
			//Hardcoded ja-JP fallback: parser can run in workers/native
			//before Asc.g_oLcidNameToIdMap is initialised.
			if (!lcid && langLower === "ja-jp") {
				lcid = 0x0411;
			}
			if (lcid) {
				target.Lid = lcid.toString(16);
			}
		}
	}
	while (ti < tokens.length) {
		var tok = tokens[ti];
		if (tok === "x" || tok === "X") {
			ti++;
			if (ti < tokens.length) {
				var ext = tokens[ti].toLowerCase();
				if (ext === "gannen" && target.LidName === "ja-jp") {
					target.bGannen = true;
				}
				ti++;
			}
		} else if (/^\d+$/.test(tok)) {
			target.CalendarId = parseInt(tok, 10);
			ti++;
		} else {
			ti++;
		}
	}
}
function FormatObjBracket(sData)
{
    this.type = numFormat_Bracket;
    this.val = sData;
    this.parse = function(data)
    {
        var length = data.length;
        if(length > 0)
        {
            var first = data[0];
            if("$" == first)
            {
                //Examples include currency, LCID, BCP-47, private-use, and calendar modifiers.
                parseBracketLocaleModifier(data, this);
            }
			else if("=" == first || ">" == first || "<" == first)
			{
				var nIndex = 1;
				var sOperator = first;
				if(length > 1 && (">" == first || "<" == first))
				{
					var second = data[1];
					if("=" == second || (">" == second && "<" == first))
					{
						sOperator += second;
						nIndex = 2;
					}
				}
				switch(sOperator)
				{
					case "=": this.operator = NumComporationOperators.equal;break;
					case ">": this.operator = NumComporationOperators.greater;break;
					case "<": this.operator = NumComporationOperators.less;break;
					case ">=": this.operator = NumComporationOperators.greaterorequal;break;
					case "<=": this.operator = NumComporationOperators.lessorequal;break;
					case "<>": this.operator = NumComporationOperators.notequal;break;
				}
				this.operatorValue = parseFloat(data.substring(nIndex));
			}
            else
            {
				var sLowerColor = data.toLowerCase();
                //todo Color1-56
                if("black" == sLowerColor)
                    this.color = 0x000000;
                else if("blue" == sLowerColor)
                    this.color = 0x0000ff;
                else if("cyan" == sLowerColor)
                    this.color = 0x00ffff;
                else if("green" == sLowerColor)
                    this.color = 0x00ff00;
                else if("magenta" == sLowerColor)
                    this.color = 0xff00ff;
                else if("red" == sLowerColor)
                    this.color = 0xff0000;
                else if("white" == sLowerColor)
                    this.color = 0xffffff;
                else if("yellow" == sLowerColor)
                    this.color = 0xffff00;
                else if("y" == first || "m" == first || "d" == first || "h" == first || "s" == first ||
                    "Y" == first || "M" == first || "D" == first || "H" == first || "S" == first ||
					"a" == first)
                {
                    var bSame = true;
                    var nCount = 1;
                    for(var i = 1; i < length; ++i)
                    {
                        if(first != data[i])
                        {
                            bSame = false;
                            break;
                        }
                        nCount++;
                    }
                    if(true == bSame)
                    {
                        switch(first)
                        {
                            case "Y":
                            case "y": this.dataObj = new FormatObjDateVal(numFormat_Year, nCount, true);break;
                            case "M":
                            case "m": this.dataObj = new FormatObjDateVal(numFormat_MonthMinute, nCount, true);break;
                            case "D":
                            case "d": this.dataObj = new FormatObjDateVal(numFormat_Day, nCount, true);break;
                            case "H":
                            case "h": this.dataObj = new FormatObjDateVal(numFormat_Hour, nCount, true);break;
                            case "S":
                            case "s": this.dataObj = new FormatObjDateVal(numFormat_Second, nCount, true);break;
                            case "a": this.dataObj = new FormatObjDateVal(numFormat_DayOfWeek, nCount, true);break;
                        }
                    }
                }
            }
        }
    };
    this.parse(sData);
}
function ParseLocalFormatSymbol(Name)
{
	LocaleFormatSymbol['Y'] = 'Y';
	LocaleFormatSymbol['y'] = 'y';
	LocaleFormatSymbol['M'] = 'M';
	LocaleFormatSymbol['m'] = 'm';
	LocaleFormatSymbol['D'] = 'D';
	LocaleFormatSymbol['d'] = 'd';
	LocaleFormatSymbol['H'] = 'H';
	LocaleFormatSymbol['h'] = 'h';
	LocaleFormatSymbol['Minute'] = 'M';
	LocaleFormatSymbol['minute'] = 'm';
	LocaleFormatSymbol['S'] = 'S';
	LocaleFormatSymbol['s'] = 's';
	LocaleFormatSymbol['a'] = 'a';
	LocaleFormatSymbol['G'] = 'G';
	LocaleFormatSymbol['g'] = 'g';
	LocaleFormatSymbol['general'] = 'General';
	var overrides = AscCommon.g_oLocaleFormatSymbols && AscCommon.g_oLocaleFormatSymbols[Name];
	if (typeof overrides === 'string') {
		overrides = AscCommon.g_oLocaleFormatSymbols[overrides];
	}
	if (overrides) {
		for (var k in overrides) {
			if (overrides.hasOwnProperty(k)) {
				LocaleFormatSymbol[k] = overrides[k];
			}
		}
	}
	return true;
}
function NumFormat(bAddMinusIfNes)
{
    //Format reading stream
    this.formatString = "";
    this.length = this.formatString.length;
    this.index = 0;
    this.EOF = -1;
    
    //Format
    this.aRawFormat = [];
    this.aDecFormat = [];
    this.aFracFormat = [];
    this.bDateTime = false;
	this.bDate = false;
	this.bTime = false;//flag to distinguish date format with time from simple date
	this.bDay = false;//to distinguish when to use MonthGenitiveNames
    this.nPercent = 0;
    this.bScientific = false;
    this.bThousandSep = false;
    this.nThousandScale = 0;
    this.bTextFormat = false;
    this.bTimePeriod = false;
    this.bMillisec = false;
    this.bSlash = false;
    this.bWhole = false;
	this.bCurrency = false;
	this.bRepeat = false;
    this.Color = -1;
	this.ComporationOperator = null;
	this.LCID = null;
	this.CurrencyString = null;
	this.DBNum = 0;
	//Set from canonical [$-ja-JP-x-gannen]. Enables Gannen year-one substitution.
	this.bGannen = false;
	//Format contains gg or ggg. Excel keeps Latin-width era names numeric.
	this.bHasKanjiEra = false;
	//Numeric [$-411,80] is a non-era calendar override in Excel; the BCP-47
	//[$-ja-JP,80] keeps era rendering. Falls back to Gregorian when set.
	this.bJapanEraCalendarOverride = false;
	//Set by a parsed Japanese LCID bracket; applies to later localized era-name
	//tokens in this sub-format (g/G normally, x/X when local g/G is day/hour).
	this.bJapanEraTokenContext = false;

	this.bGeneralChart = false;//if the format contains only one text (e.g. "General" in chart)
    this.bAddMinusIfNes = bAddMinusIfNes;//when formatting for negative numbers is not specified, sometimes a minus sign needs to be inserted
}
NumFormat.prototype =
{
    _getChar : function()
    {
        if(this.index < this.length)
        {
            return this.formatString[this.index];
        }
        return this.EOF;
    },
    _readChar : function()
    {
        var curChar = this._getChar();
        if(this.index < this.length)
            this.index++;
        return curChar;
    },
    _skip : function(val)
    {
        var nNewIndex = this.index + val;
        if(nNewIndex >= 0)
            this.index = nNewIndex;
    },
    _addToFormat : function(type, val)
    {
        var oFormatObj = new FormatObj(type, val);
        this.aRawFormat.push(oFormatObj);
    },
    _addToFormat2 : function(oFormatObj)
    {
        this.aRawFormat.push(oFormatObj);
    },
    _ReadText : function(endChar)
    {
        var sText = "";
        while(true)
        {
            var next = this._readChar();
            if(this.EOF == next || endChar == next)
                break;
            else
            {
                sText += next;
            }
        }
        this._addToFormat(numFormat_Text, sText);
    },
    _GetText : function(len)
    {
        return this.formatString.substr(this.index, len);
    },
    _ReadChar : function()
    {
        var next = this._readChar();
        if(this.EOF != next)
            this._addToFormat(numFormat_Text, next);
    },
    _ReadBracket : function()
    {
        var sBracket = "";
        while(true)
        {
            var next = this._readChar();
            if(this.EOF == next || "]" == next)
                break;
            else
            {
                sBracket += next;
            }
        }
		var oFormatObjBracket = new FormatObjBracket(sBracket);
		if(null != oFormatObjBracket.operator)
			this.ComporationOperator = oFormatObjBracket;
		if (AscCommon.isJapanEraLid && AscCommon.isJapanEraLid(oFormatObjBracket.Lid))
			this.bJapanEraTokenContext = true;
        this._addToFormat2(oFormatObjBracket);
    },
    _ReadAmPm : function(next)
    {
		if ("A" === next || "a" === next) {
			let ampm = "AM/PM";
			if (ampm.substring(1) === this._GetText(ampm.length - 1).toUpperCase()) {
				this._addToFormat2(new FormatObj(numFormat_AmPm));
				this.bTimePeriod = true;
				this.bDateTime = true;
				this._skip(ampm.length - 1);
				return true;
			}
		}
		if ("上" === next) {
			let ampm = "上午/下午";
			if (ampm.substring(1) === this._GetText(ampm.length - 1).toUpperCase()) {
				this._addToFormat2(new FormatObj(numFormat_AmPm));
				this.bTimePeriod = true;
				this.bDateTime = true;
				this._skip(ampm.length - 1);
				return true;
			}
		}
		return false;
    },
	_ReadAmPmPDF : function(next)
    {
		let bAmPm = true;
		let nttCount = 1;
        while(true)
        {
            next = this._readChar();
            if(this.EOF == next)
                break;
            else if ("t" == next)
            {
				nttCount++;
            }
            else
            {
				// if more than two tt, don't add am/pm
				if (nttCount > 2) {
					bAmPm = false;
				}

				this._skip(-1);
				break;
            }
        }
        if(bAmPm == true)
        {
            this._addToFormat2(new FormatObj(numFormat_AmPm));
            this.bTimePeriod = true;
            this.bDateTime = true;
        }
    },
    _parseFormat : function(digitSpaceSymbol, useLocaleFormat)
    {
        var sGeneral;
        var DecimalSeparator;
        var GroupSeparator;
        var TimeSeparator;
        var Year;
        var Month;
        var Day;
        var Hour;
        var year;
        var month;
        var day;
        var hour;
        var Minute;
        var minute;
        var Second;
        var second;
		var dayOfWeek;
		var JapanEra;
		var japanEra;
		if (useLocaleFormat) {
			sGeneral = LocaleFormatSymbol['general'].toLowerCase();
			DecimalSeparator = g_oDefaultCultureInfo.NumberDecimalSeparator;
			TimeSeparator = g_oDefaultCultureInfo.TimeSeparator;
			GroupSeparator = g_oDefaultCultureInfo.NumberGroupSeparator;
			Year = LocaleFormatSymbol['Y'];
			year = LocaleFormatSymbol['y'];
			Month = LocaleFormatSymbol['M'];
			month = LocaleFormatSymbol['m'];
			Day = LocaleFormatSymbol['D'];
			day = LocaleFormatSymbol['d'];
			Hour = LocaleFormatSymbol['H'];
			hour = LocaleFormatSymbol['h'];
			Minute = LocaleFormatSymbol['Minute'];
			minute = LocaleFormatSymbol['minute'];
			Second = LocaleFormatSymbol['S'];
			second = LocaleFormatSymbol['s'];
			dayOfWeek = LocaleFormatSymbol['a'];
			JapanEra = LocaleFormatSymbol['G'];
			japanEra = LocaleFormatSymbol['g'];
		} else {
			sGeneral = AscCommon.g_cGeneralFormat.toLowerCase();
			DecimalSeparator = gc_sFormatDecimalPoint;
			TimeSeparator = ':';
			GroupSeparator = gc_sFormatThousandSeparator;
			Year = 'Y';
			year = 'y';
			Month = 'M';
			month = 'm';
			Day = 'D';
			day = 'd';
			Hour = 'H';
			hour = 'h';
			Minute = 'M';
			minute = 'm';
			Second = 'S';
			second = 's';
			dayOfWeek = 'a';
			JapanEra = 'G';
			japanEra = 'g';
		}
        var sGeneralFirst = sGeneral[0];
        this.bGeneralChart = true;
        while(true)
        {
            var next = this._readChar();
            var bNoFormat = false;
            if(this.EOF == next)
                break;
            else if("[" == next)
                this._ReadBracket();
            else if("\"" == next)
                this._ReadText("\"");
            else if("\\" == next)
                this._ReadChar();
            else if("%" == next)
            {
                this._addToFormat(numFormat_Percent);
            }
            else if(TimeSeparator == next)
            {
                this._addToFormat(numFormat_TimeSeparator);
            }
            else if('0' === next)
            {
                this._addToFormat(numFormat_Digit, 0);
            }
            else if("#" == next)
            {
                this._addToFormat(numFormat_DigitNoDisp);
            }
            else if(digitSpaceSymbol == next)
            {
                this._addToFormat(numFormat_DigitSpace);
            }
            else if(DecimalSeparator == next)
            {
                this._addToFormat(numFormat_DecimalPoint);
            }
            else if("/" == next)
            {
                this._addToFormat2(new FormatObjDecimalFrac([], []));
            }
            else if(GroupSeparator == next)
            {
                this._addToFormat(numFormat_Thousand, 1);
            }
            else if("$" == next || "+" == next || "-" == next || "(" == next || ")" == next || " " == next)
            {
                this._addToFormat(numFormat_Text, next);
            }
            else if (sGeneralFirst === next.toLowerCase() &&
                sGeneral === (next + this._GetText(sGeneral.length - 1)).toLowerCase()) {
                this._addToFormat(numFormat_General);
                this._skip(sGeneral.length - 1);
            }
			else if (this._ReadAmPm(next))
			{

			}
            else if("E" == next || "e" == next)
            {
                //Precedence: E/e+sign -> scientific; lowercase e -> era-year; uppercase E -> literal.
                //Peek instead of consume so the non-sign char (e.g. '.' in "ge.m.d") survives.
                var peekSign = this._GetText(1);
                if (peekSign === "+" || peekSign === "-") {
                    this._readChar();
                    var sign = ("+" === peekSign) ? SignType.Positive : SignType.Negative;
                    this._addToFormat2(new FormatObjScientific(next, "", sign));
                } else if ("e" === next) {
                    this._addToFormat2(new FormatObjDateVal(numFormat_JapanEraYear, 1, false));
                } else {
                    bNoFormat = true;
                    this._addToFormat(numFormat_Text, next);
                }
            }
            else if("*" == next)
            {
                var nextnext = this._readChar();
                if(this.EOF != nextnext)
                    this._addToFormat(numFormat_Repeat, nextnext);
            }
            else if("_" == next)
            {
                var nextnext = this._readChar();
                if(this.EOF != nextnext)
                    this._addToFormat(numFormat_Skip, nextnext);
            }
            else if("@" == next)
            {
                this._addToFormat(numFormat_TextPlaceholder);
            }
            else if((JapanEra == next || japanEra == next) && this.bJapanEraTokenContext)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_JapanEra, 1, false));
            }
            else if(Year == next || year == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
            }
            else if(Month == next || month == next)
            {
                if (Month === Minute) {
                    this._addToFormat2(new FormatObjDateVal(numFormat_MonthMinute, 1, false));
                } else {
                    this._addToFormat2(new FormatObjDateVal(numFormat_Month, 1, false));
                }
            }
            else if(Day == next || day == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
            }
            else if(Hour == next || hour == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
            }
            else if(Minute == next || minute == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Minute, 1, false));
            }
            else if(Second == next || second == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
            }
			else if (dayOfWeek == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_DayOfWeek, 1, false));
			}
            else if ("g" == next || "G" == next)
            {
                this._addToFormat2(new FormatObjDateVal(numFormat_JapanEra, 1, false));
            }
            else {
                bNoFormat = true;
                this._addToFormat(numFormat_Text, next);
            }
            if (!bNoFormat)
                this.bGeneralChart = false;
        }
        return true;
    },
    _parseFormatWordDateTime : function()
    {
        while(true)
        {
            var next = this._readChar();
			if(this.EOF == next)
				break;
			else if("\'" == next)
				this._ReadText("\'");
			else if (this._ReadAmPm(next))
			{

			}
			else if("Y" == next || "y" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
			}
			else if("M" == next || "m" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_MonthMinute, 1, false));
			}
			else if("D" == next || "d" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
			}
			else if("H" == next || "h" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
			}
			else if("S" == next || "s" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
			}
			else if ("a" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_DayOfWeek, 1, false));
			}
			else {
					this._addToFormat(numFormat_Text, next);
			}
        }
        return true;
    },
	_parseFormatPDFDateTime : function()
    {
        while(true)
        {
            var next = this._readChar();
			if(this.EOF == next)
				break;
			else if("\'" == next)
				this._ReadText("\'");
			else if ("y" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Year, 1, false));
			}
			else if ("m" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Month, 1, false));
			}
			else if ("M" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Minute, 1, false));
			}
			else if ("d" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Day, 1, false));
			}
			else if ("h" == next || "H" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Hour, 1, false));
			}
			else if ("s" == next)
			{
				this._addToFormat2(new FormatObjDateVal(numFormat_Second, 1, false));
			}
			else if ("t" == next) {
				this._ReadAmPmPDF(next);
			}
			else {
				this._addToFormat(numFormat_Text, next);
			}
        }
        return true;
    },
	_parseFormatWordNumeric : function(digitSpaceSymbol)
	{
		while(true)
		{
			var next = this._readChar();
			if (this.EOF == next) {
				break;
			} else if ("\'" === next) {
				this._ReadText("\'");
			} else if ('0' === next) {
				this._addToFormat(numFormat_Digit, 0);
			} else if (digitSpaceSymbol === next) {
				this._addToFormat(numFormat_DigitSpace);
			} else if ('x' === next || 'X' === next) {
				this._addToFormat(numFormat_DigitDrop);
			} else if (gc_sFormatDecimalPoint === next) {
				this._addToFormat(numFormat_DecimalPoint);
			} else if (gc_sFormatThousandSeparator === next) {
				this._addToFormat(numFormat_Thousand, 1);
			} else if ('+' === next) {
				this._addToFormat(numFormat_Plus);
			} else if ('-' === next) {
				this._addToFormat(numFormat_Minus);
			} else {
				this._addToFormat(numFormat_Text, next);
			}
		}
		return true;
	},
	_isDigitType: function(type) {
		return numFormat_Digit === type || numFormat_DigitNoDisp === type || numFormat_DigitSpace === type ||
			numFormat_DigitDrop === type;
	},
    _prepareFormat : function()
    {
        //Color
		for(var i = 0, length = this.aRawFormat.length; i < length; ++i)
        {
            var oCurItem = this.aRawFormat[i];
            if(numFormat_Bracket == oCurItem.type && null != oCurItem.color)
                this.Color = oCurItem.color;
        }
        this.bRepeat = false;
        var nFormatLength = this.aRawFormat.length;

        //Group several consecutive elements into one special symbol
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Repeat == item.type)
            {
                //Keep only the last numFormat_Repeat
                if(false == this.bRepeat)
                    this.bRepeat = true;
                else
                {
                    this.aRawFormat.splice(i, 1);
                    nFormatLength--;
                }
            }
            else if(numFormat_Bracket == item.type)
            {
                //Handle [hhh]
                var oNewObj = item.dataObj;
                if(null != oNewObj)
                {
                    this.aRawFormat.splice(i, 1, oNewObj);
                    this.bDateTime = true;
                    if(numFormat_Hour == oNewObj.type || numFormat_Minute == oNewObj.type || numFormat_Second == oNewObj.type || numFormat_Milliseconds == oNewObj.type)
                        this.bTime = true;
                    else if (numFormat_Year == oNewObj.type || numFormat_Month == oNewObj.type || numFormat_Day == oNewObj.type) {
                        this.bDate = true;
                        if (numFormat_Day == oNewObj.type)
                            this.bDay = true;
                    }
                }
            }
            else if(numFormat_Year == item.type || numFormat_MonthMinute == item.type || numFormat_Month == item.type || numFormat_Day == item.type || numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type || numFormat_Thousand == item.type ||
				numFormat_DayOfWeek == item.type || numFormat_JapanEra == item.type || numFormat_JapanEraYear == item.type)
            {
                //Combine hhh sequences into one
                var nStartType = item.type;
                var nEndIndex = i;
                for(var j = i + 1; j < nFormatLength; ++j)
                {
                    if(nStartType == this.aRawFormat[j].type)
                        nEndIndex = j;
                    else
                        break;
                }
                if(i != nEndIndex)
                {
                    item.val = nEndIndex - i + 1;
                    var nDelCount = item.val - 1;
                    this.aRawFormat.splice(i + 1, nDelCount);
                    nFormatLength -= nDelCount;
                }
                if(numFormat_Thousand != item.type)
                {
                    this.bDateTime = true;
                    if(numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type || numFormat_Milliseconds == item.type)
                        this.bTime = true;
                    else if (numFormat_Year == item.type || numFormat_Month == item.type || numFormat_Day == item.type) {
                        this.bDate = true;
                        if (numFormat_Day == item.type)
                            this.bDay = true;
                    }
                    //Only kanji-width era names enable Gannen substitution.
                    else if (numFormat_JapanEra == item.type && item.val >= 2) {
                        this.bHasKanjiEra = true;
                    }
                }
            }
            else if(numFormat_Scientific == item.type)
            {
                var bAsText = false;
                if(true == this.bScientific)
                {
                    bAsText = true;
                }
                else
                {
                    var aDigitArray = [];
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var nextItem = this.aRawFormat[j];
                        if(this._isDigitType(nextItem.type))
                            aDigitArray.push(nextItem);
                    }
                    if(aDigitArray.length > 0)
                    {
                        item.format = aDigitArray;
                        this.bScientific = true;
                    }
                    else
                        bAsText = true;
                }
                if(false != bAsText)
                {
                    //replace with text
                    item.type = numFormat_Text;
                    item.val = item.val + "+";
                }
            }
            else if(numFormat_DecimalFrac == item.type)
            {
                var bValid = false;
                //collect left and right parts of the fraction
                var nLeft = i;
                for(var j = i - 1; j >= 0; --j)
                {
                    var subitem = this.aRawFormat[j];
                    if(this._isDigitType(subitem.type))
                        nLeft = j;
                    else
                        break;
                }
                var nRight = i;
                if(nLeft < i)
                {
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var subitem = this.aRawFormat[j];
                        if(this._isDigitType(subitem.type) || (numFormat_Text === subitem.type && '0' <= subitem.val && subitem.val <= '9'))
                            nRight = j;
                        else
                            break;
                    }
                    if(nRight > i)
                    {
                        bValid = true;
                        item.aRight = this.aRawFormat.splice(i + 1, nRight - i);
                        item.aLeft = this.aRawFormat.splice(nLeft, i - nLeft);
                        nFormatLength -= nRight - nLeft;
                        i -= i - nLeft;
                        this.bSlash = true;

                        var flag = (item.aRight.length > 0) && (item.aRight[0].type == numFormat_Digit || item.aRight[0].type == numFormat_Text) && (parseInt(item.aRight[0].val) > 0);
                        if(flag)
                        {
                            var rPart = 0;
                            for(var j = 0; j< item.aRight.length; j++)
                            {
                                if(item.aRight[j].type == numFormat_Digit || item.aRight[j].type == numFormat_Text)
                                    rPart = rPart*10 + parseInt(item.aRight[j].val);
                                else
                                {
                                    bValid = false;
                                    this.bSlash = false;
                                    break;
                                }
                            }
                            if(bValid == true)
                            {
                                item.aRight = [];
                                item.aRight.push(new FormatObj(numFormat_Digit, rPart));
                                item.bNumRight = true;
                            }
                        }
                    }

                }

                if(false == bValid)
                {
                    item.type = numFormat_DateSeparator;
                }
            }
        }
        
        var nReadState = FormatStates.Decimal;
        var bDecimal = true;
        nFormatLength = this.aRawFormat.length;
        //Resolve conflicts, set property values
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_DecimalPoint == item.type)
            {
                //milliseconds
                //If numFormat_Digit follows DecimalPoint, and there is a date/time format, then these are milliseconds
                if(this.bDateTime)
                {
                    var nStartIndex = i;
                    var nEndIndex = nStartIndex;
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var subItem = this.aRawFormat[j];
                        if(numFormat_Digit == subItem.type)
                            nEndIndex = j;
                        else
                            break;
                    }
                    if(nStartIndex < nEndIndex)
                    {
                        var nDigCount = nEndIndex - nStartIndex;
                        var oNewItem = new FormatObjDateVal(numFormat_Milliseconds, nDigCount, false);
                        var nDelCount = nDigCount;
                        oNewItem.format = this.aRawFormat.splice(i + 1, nDelCount, oNewItem);
                        nFormatLength -= (nDigCount - 1);
                        i++;
                        this.bMillisec = true;

                    }
                    //convert all subsequent to text
                    item.type = numFormat_DecimalPointText;
                    item.val = null;
                }
                else if(FormatStates.Decimal == nReadState)
                    nReadState = FormatStates.Frac;
            }
            else if(numFormat_MonthMinute == item.type)
            {
                //Resolve numFormat_MonthMinute conflicts
                var bRightCond = false;
                if (item.bElapsed)
                {
                    bRightCond = true;
                }
                else
                {
                    //search forward for the first element with datetime type
                    for(var j = i + 1; j < nFormatLength; ++j)
                    {
                        var subItem = this.aRawFormat[j];
                        if(numFormat_Year == subItem.type || numFormat_Month == subItem.type || numFormat_Day == subItem.type || numFormat_MonthMinute == subItem.type ||
                        numFormat_Hour == subItem.type || numFormat_Minute == subItem.type || numFormat_Second == subItem.type || numFormat_Milliseconds == subItem.type)
                        {
                            if(numFormat_Second == subItem.type)
                                bRightCond = true;
                            break;
                        }
                    }
                }
                var bLeftCond = false;
                if(false == bRightCond)
                {
                    //search backward for the first element with type hh or ss
                    var bFindSec = false;//to resolve the case mm:ss:mm it should be Minutes:Seconds:Months
                    for(var j = i - 1; j >= 0; --j)
                    {
                        var subItem = this.aRawFormat[j];
                        
                        if(numFormat_Hour == subItem.type)
                        {
                            bLeftCond = true;
                            break;
                        }
                        else if(numFormat_Second == subItem.type)
                        {
                            //continue looking further until the next date time object is found
                            bFindSec = true;
                        }
                        else if(numFormat_Minute == subItem.type || numFormat_Month == subItem.type || numFormat_MonthMinute == subItem.type)
                        {
                            if(true == bFindSec && numFormat_Minute == subItem.type)
                                bFindSec = false;
                            break;
                        }
                        else if(numFormat_Year == subItem.type || numFormat_Day == subItem.type || numFormat_Hour == subItem.type || numFormat_Second == subItem.type || numFormat_Milliseconds == subItem.type)
                        {
                            if(true == bFindSec)
                                break;
                        }
                    }
                    if(true == bFindSec)
                        bLeftCond = true;
                }
                
                if((true == bLeftCond || true == bRightCond) && item.val <= 2)
				{
                    item.type = numFormat_Minute;
					this.bTime = true;
				}
                else
				{
                    item.type = numFormat_Month;
					this.bDate = true;
				}
            }
            else if(numFormat_Percent == item.type)
            {
                this.nPercent++;
                //replace with text
                item.type = numFormat_Text;
                item.val = "%";
            }
            else if(numFormat_Thousand == item.type)
            {
                var isPrevDigit = i > 0 && this._isDigitType(this.aRawFormat[i - 1].type);
                var isPrevDecimalPoint = i > 0 && numFormat_DecimalPoint === this.aRawFormat[i - 1].type;
                var isNextDigit = i + 1 < nFormatLength && this._isDigitType(this.aRawFormat[i + 1].type);
                if (isPrevDigit && isNextDigit) {
                    if(FormatStates.Decimal == nReadState) {
                        this.bThousandSep = true;
                    }
                } else if (isPrevDigit || isPrevDecimalPoint) {
                    this.nThousandScale = item.val;
                } else {
                    item.type = numFormat_ThousandText;
                }
            }
            else if(this._isDigitType(item.type))
            {
                this.nThousandScale = 0;
                if(FormatStates.Decimal == nReadState)
                {
                    this.aDecFormat.push(item);

                    if(this.bSlash === true)
                        this.bWhole = true;
                }
                else if(FormatStates.Frac == nReadState)
                    this.aFracFormat.push(item);

            }
            else if(numFormat_Scientific == item.type)
                nReadState = FormatStates.Scientific;
            else if(numFormat_TextPlaceholder == item.type)
            {
                this.bTextFormat = true;
            }
        }
        return true;
    },
	_prepareFormatDatePDF : function()
    {
		var nFormatLength = this.aRawFormat.length;
        //Group several consecutive elements into one special symbol
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Year == item.type || numFormat_Month == item.type || numFormat_Day == item.type)
            {
                // Remove "yyy" (3 y's) for year type, or any date item with val > 4.
                // Use else-if because both conditions are mutually exclusive (3 is not > 4),
                // and after splice the index must be decremented so the next element is not skipped.
				if(item.val === 3 && numFormat_Year == item.type)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength--;
					i--;
                }
                else if(item.val > 4)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength--;
					i--;
                }
            }
			else if(numFormat_Hour == item.type || numFormat_Minute == item.type || numFormat_Second == item.type)
            {
				// Remove time items with val > 2 (only 1 or 2 repetitions are valid).
                if(item.val > 2)
                {
                    this.aRawFormat.splice(i, 1);
					nFormatLength--;
					i--;
                }
            }
        }
    },
	_calsScientific : function(nDecLen, nRealExp)
	{
		var nKoef = 0;
		if(true == this.bThousandSep)
			nKoef = 4;
		if(nDecLen > nKoef)
			nKoef = nDecLen;
		if(nRealExp > 0 && nKoef > 0)
		{
			var nTemp = nRealExp % nKoef;
			if(0 == nTemp)
				nTemp = nKoef;
			nKoef = nTemp;
		}
		return nKoef;
	},
	_parseNumber : function(number, aDecFormat, nFracLen, nValType)
    {
        var res = {bDigit: false, dec: 0, frac: 0, fraction: 0, exponent: 0, exponentFrac: 0, scientific: 0, sign: SignType.Positive, date: {}};
        if(CellValueType.String != nValType)
            res.bDigit = (number == number - 0);
        if(res.bDigit)
        {
			var numberAbs = Math.abs(number);
			res.fraction = numberAbs - Math.floor(numberAbs);
			//Round
			var parts = getNumberParts(number);
			res.sign = parts.sign;
			var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0 for 0.123
			if(SignType.Null != parts.sign)
			{
				if(true == this.bScientific)
				{
					var nKoef = this._calsScientific(aDecFormat.length, nRealExp);
					res.scientific = nRealExp - nKoef;
					nRealExp = nKoef;
				}
				else
				{
					//Percent
					for(var i = 0; i < this.nPercent; ++i)
						nRealExp += 2;
					//Thousands separator
					for(var i = 0; i < this.nThousandScale; ++i)
						nRealExp -= 3;		
				}
				//round after operations that may change nRealExp
				if(false == this.bSlash)
				{
					var nOldRealExp = nRealExp;
					parts = getNumberParts(round10(parts.mantissa, nFracLen + nRealExp - gc_nMaxDigCount, nFracLen));
					if(SignType.Null != parts.sign)
					{
						nRealExp = gc_nMaxDigCount + parts.exponent;
						if(nOldRealExp != nRealExp && true == this.bScientific)
						{
							var nKoef = this._calsScientific(aDecFormat.length, nRealExp);
							res.scientific += nRealExp - nOldRealExp;
							nRealExp = nKoef;
						}
					}
				}
				res.exponent = nRealExp;
				res.exponentFrac = nRealExp;
				if(nRealExp > 0 && nRealExp < gc_nMaxDigCount)
				{
					var sNumber = parts.mantissa.toString();
					var nExponentFrac = 0;
					for(var i = nRealExp, length = sNumber.length; i < length; ++i)
					{
						if("0" == sNumber[i])
							nExponentFrac++;
						else
							break;
					}
					if(nRealExp + nExponentFrac < sNumber.length)
						res.exponentFrac = - nExponentFrac;
				}
				if(SignType.Null != parts.sign)
				{
					if(nRealExp <= 0)
					{
						if(this.bSlash == true)
						{
							res.dec = 0;
							res.frac = parts.mantissa;
						}
						else
						{
							if(nFracLen > 0)
							{
								res.dec = 0;
								res.frac = 0;
								if(nFracLen + nRealExp > 0)
								{
									var sTemp = parts.mantissa.toString();
									res.frac = sTemp.substring(0, nFracLen + nRealExp) - 0;
								}
							}
							else
							{
								res.dec = 0;
								res.frac = 0;
							}
						}
					}
					else if(nRealExp >= gc_nMaxDigCount)
					{
						res.dec = parts.mantissa;
						res.frac = 0;
					}
					else
					{
						var sTemp = parts.mantissa.toString();
						if(this.bSlash == true)
						{
							res.dec = sTemp.substring(0, nRealExp) - 0;
							if(nRealExp < sTemp.length)
								res.frac = sTemp.substring(nRealExp) - 0;
							else
								res.frac = 0;
						}
						else
						{
							if(nFracLen > 0 )
							{
								res.dec = sTemp.substring(0, nRealExp) - 0;
								res.frac = 0;
								var nStart = nRealExp;
								var nEnd = nRealExp + nFracLen;
								if(nStart < sTemp.length)
									res.frac = sTemp.substring(nStart, nEnd) - 0;
							}
							else
							{
								res.dec = sTemp.substring(0, nRealExp) - 0;
								res.frac = 0;
							}
						}
					}
				}
				if(0 == res.frac && 0 == res.dec && false === this.bDateTime)
					res.sign = SignType.Null;
			}
			//After rounding the result may be zero,
			//but didn't move the sign check here because rounding requires a non-negative number

            if(this.bDateTime === true)
				res.date = this.parseDate(number);
        }
        return res;
    },
	_parseNumberForPDFDate : function(number) {
		let oDateTmp = new Date();
		oDateTmp.setTime(number * (86400 * 1000));
	 
		return {
			date: {
				d:			oDateTmp.getDate(),
				dayWeek:	oDateTmp.getDay(),
				hour:		oDateTmp.getHours(),
				min:		oDateTmp.getMinutes(),
				month:		oDateTmp.getMonth(),
				ms:			0,
				//ms:			oDateTmp.getMilliseconds(),
				sec:		oDateTmp.getSeconds(),
				year:		oDateTmp.getFullYear()
			}
		}
	},
	parseDate : function(number)
	{
        var d = {val: 0, coeff: 1}, h = {val: 0, coeff: 24},
            min = {val: 0, coeff: 60}, s = {val: 0, coeff: 60}, ms = {val: 0, coeff: 1000};
        //number is negative in case of bDate1904
        var numberAbs = this.formatType == AscCommon.NumFormatType.PDFFormDate ? number : Math.abs(number);
        var tmp = numberAbs;
        var ttimes = [d, h, min, s, ms];
        for(var i = 0; i < 4; i++)
        {
            var v = tmp*ttimes[i].coeff;
            ttimes[i].val = Math.floor(v);
            tmp = v - ttimes[i].val;
        }
        ms.val = Math.round(tmp*1000);
        for(i = 4; i > 0 && (ttimes[i].val === ttimes[i].coeff); i--)
        {
            ttimes[i].val = 0;
            ttimes[i-1].val++;
        }
        var stDate, day, month, year, dayWeek;
		if(AscCommon.bDate1904)
		{
			stDate = new Date(Date.UTC(1904,0,1,0,0,0));
			if(d.val)
				stDate.setUTCDate( stDate.getUTCDate() + d.val );
			day = stDate.getUTCDate();
			dayWeek = stDate.getUTCDay();
			month = stDate.getUTCMonth();
			year = stDate.getUTCFullYear();
		}
		else
		{
			if (60 <= numberAbs && numberAbs < 61)
			{
				day = 29;
				month = 1;
				year = 1900;
				dayWeek = 3;
			}
			else if (0 <= numberAbs && numberAbs < 1)
			{
				//TODO need to use cDate everywhere
				stDate = new Asc.cDate(Date.UTC(1899,11,31,0,0,0));
				day = stDate.getUTCDate();
				dayWeek = ( stDate.getUTCDay() > 0) ? stDate.getUTCDay() - 1 : 6;
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
			else if(numberAbs < 60 && number > 0)
			{
				stDate = new Date(Date.UTC(1899,11,31,0,0,0));
				if(d.val)
				// setUTCDate doesn't consider the transition from 1899 to 1900 when adding d.val
					stDate.setUTCDate( stDate.getUTCDate() + d.val );
				day = stDate.getUTCDate();
				dayWeek = ( stDate.getUTCDay() > 0) ? stDate.getUTCDay() - 1 : 6;
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
			else
			{
				stDate = new Date(Date.UTC(1899,11,30,0,0,0));
				if(d.val)
					stDate.setUTCDate( stDate.getUTCDate() + d.val );
				day = stDate.getUTCDate();
				dayWeek = stDate.getUTCDay();
				month = stDate.getUTCMonth();
				year = stDate.getUTCFullYear();
			}
		}
        return {d: day, month: month, year: year, dayWeek: dayWeek, hour: h.val, min: min.val, sec: s.val, ms: ms.val, countDay: d.val };
	},
	_FormatNumber: function (number, exponent, format, nReadState, cultureInfo, opt_forceNull)
	{
        var aRes = [];
        var nFormatLen = format.length;
        if(nFormatLen > 0)
        {
            if(FormatStates.Frac != nReadState && FormatStates.SlashFrac != nReadState)
            {
				var sNumber = number + "";
				var nNumberLen = sNumber.length;
				// Bug 14325: number like "1.23456789123456e+23" with format "0.000…" (30 zeros)
				if(exponent > nNumberLen)
				{
					for(var i = 0; i < exponent - nNumberLen; ++i)
						sNumber += "0";
					nNumberLen = sNumber.length;
				}
                var bIsNUll = false;
                if("0" == sNumber && !opt_forceNull)
                    bIsNUll = true;
                // Use an index instead of Array.shift() to avoid O(n) per call.
                // format is always a fresh .concat() copy, so mutation is safe, but
                // shifting re-indexes the whole array each time (O(n²) overall).
                var fmtIdx = 0;
                //align length
                if(nNumberLen > nFormatLen)
                {
                    if(false === bIsNUll)
                    {
						var item = format[fmtIdx++];
						if (numFormat_DigitDrop !== item.type) {
							var nSplitIndex = nNumberLen - nFormatLen + 1;
							aRes.push(new FormatObj(numFormat_Text, sNumber.slice(0, nSplitIndex)));
							sNumber = sNumber.substring(nSplitIndex);
						} else {
							sNumber = sNumber.substring(nNumberLen - nFormatLen);
						}
                    }
                }
                else if(nNumberLen < nFormatLen)
                {
                    // leading padding — only zeros and spaces here
                    for(var i = 0, length = nFormatLen - nNumberLen; i < length; ++i)
                    {
                        var item = format[fmtIdx++];
                        aRes.push(new FormatObj(item.type));
                    }
                }
                // fill digit-by-digit
                for(var i = 0, length = sNumber.length; i < length; ++i)
                {
                    var sCurNumber = sNumber[i];
					var numFormat = numFormat_Text;
                    var item = format[fmtIdx++];
                    if(true == bIsNUll && null != item && FormatStates.Scientific != nReadState)
					{
						if(numFormat_DigitNoDisp == item.type)
							sCurNumber = "";
						else if(numFormat_DigitSpace == item.type)
						{
							numFormat = numFormat_DigitSpace;
							sCurNumber = null;
						}
					}
                    aRes.push(new FormatObj(numFormat, sCurNumber));
                }
                
                //Insert separators
                if(true == this.bThousandSep && FormatStates.Slash != nReadState)
                {
					var sThousandSep = cultureInfo.NumberGroupSeparator;
					var aGroupSize = cultureInfo.NumberGroupSizes;
					var nCurGroupIndex = 0;
					var nCurGroupSize = 0;
					if (nCurGroupIndex < aGroupSize.length)
					    nCurGroupSize = aGroupSize[nCurGroupIndex++];
					else
					    nCurGroupSize = 0;
                    var nIndex = 0;
                    for(var i = aRes.length - 1; i >= 0; --i)
                    {
                        var item = aRes[i];
                        if(numFormat_Text == item.type)
                        {
                            var aNewText = [];
                            var nTextLength = item.val.length;
                            for(var j = nTextLength - 1; j >= 0; --j)
                            {
                                if (nCurGroupSize == nIndex)
                                {
                                    aNewText.push(sThousandSep);
                                    nTextLength++;
                                }
                                aNewText.push(item.val[j]);
                                if(0 != j)
                                {
                                    nIndex++;
                                    if (nCurGroupSize + 1 == nIndex) {
                                        nIndex = 1;
                                        if (nCurGroupIndex < aGroupSize.length)
                                            nCurGroupSize = aGroupSize[nCurGroupIndex++];
                                    }
                                }
                            }
                            if(nTextLength > 1)
                                aNewText.reverse();
                            item.val = aNewText.join("");
                        }
                        else if(numFormat_DigitNoDisp != item.type)
                        {
                            //don't add space only before numFormat_DigitNoDisp
                            if (nCurGroupSize == nIndex)
                            {
                                item.val = sThousandSep;
                                aRes[i] = item;
                            }
                        }
                        nIndex++;
                        if (nCurGroupSize + 1 == nIndex) {
                            nIndex = 1;
                            if (nCurGroupIndex < aGroupSize.length)
                                nCurGroupSize = aGroupSize[nCurGroupIndex++];
                        }
                    }
                }
            }
            else
            {
				var val = number;
				var exp = exponent;
                //Count the number of leading zeros
                var nStartNulls = 0;
				if(exp < 0)
					nStartNulls = Math.abs(exp);
                var sNumber = val.toString();
                var nNumberLen = sNumber.length;
				//remove trailing zeros
				var nLastNoNull = nNumberLen;
                for(var i = nNumberLen - 1; i >= 0; --i)
                {
					if ("0" != sNumber[i])
						break;
					nLastNoNull = i;
				}
				if (nLastNoNull < nNumberLen && (FormatStates.SlashFrac != nReadState || 0 == nLastNoNull)) {
					sNumber = sNumber.substring(0, nLastNoNull);
					nNumberLen = sNumber.length;
				}
                //fill leading zeros
                for(var i = 0; i < nStartNulls; ++i)
                    aRes.push(new FormatObj(numFormat_Text, "0"));
                //simply fill with text
                for(var i = 0, length = nNumberLen; i < length; ++i)
                    aRes.push(new FormatObj(numFormat_Text, sNumber[i]));
                //simply copy, here will be only zeros and spaces
                for(var i = nNumberLen + nStartNulls; i < nFormatLen; ++i)
                {
                    var item = format[i];
                    aRes.push(new FormatObj(item.type));
                }
            }
        }
        return aRes;
    },
	_replaceDBNumDigit: function (val) {
		//todo DBNum 1-4
		if (1 !== this.DBNum) {
			return val;
		}
		let locale = Asc.g_oLcidIdToNameMap[this.LCID];
		if (!locale) {
			return val;
		}
		locale = locale.substring(0, 2);
		if ('zh' === locale || 'ja' === locale || 'ko' === locale) {
			let dBNumVal = '';
			for (let j = 0; j < val.length; ++j) {
				if ('0' <= val[j] && val[j] <= '9') {
					dBNumVal += dBNum1Numbers[val[j] - '0'];
				} else {
					dBNumVal += val[j];
				}
			}
			val = dBNumVal;
		}
		return val;
	},
    _AddDigItem : function(res, oCurText, item)
    {
        if(numFormat_Text == item.type)
            oCurText.text += item.val;
        else if(numFormat_Digit == item.type)
        {
            //text.val may be filled in Thousand
            oCurText.text += "0";
            if(null != item.val)
                oCurText.text += item.val;
        }
        else if(numFormat_DigitNoDisp == item.type)
        {
            // No visible placeholder character — only append the value if present.
            if(null != item.val)
                oCurText.text += item.val;
        }
        else if(numFormat_DigitSpace == item.type || numFormat_DigitDrop == item.type)
        {
            var oNewFont = new AscCommonExcel.Font();
			oNewFont.skip = true;
            this._CommitText(res, oCurText, "0", oNewFont);
            if(null != item.val)
                oCurText.text += item.val;
        }
    },
    _ZeroPad: function(n)
    {
        // Always return a string so callers get consistent string concatenation.
        return n < 10 ? "0" + n : String(n);
    },
    //era is null when era rendering is disabled (non-Japanese LCID, calendar
    //override) or when the date predates Meiji.
    _resolveJapanEra: function (oParsedNumber)
    {
        var bEraLcid = !this.bJapanEraCalendarOverride
            && AscCommon.isJapanEraLcid && AscCommon.isJapanEraLcid(this.LCID);
        var era = bEraLcid ? AscCommon.getJapanEraByDate(
            oParsedNumber.date.year,
            oParsedNumber.date.month + 1,
            oParsedNumber.date.d) : null;
        return {bEraLcid: bEraLcid, era: era};
    },
    _CommitText: function(res, oCurText, textVal, format)
    {
        // Flush accumulated text first (always with format=null).
        // Previously this called _CommitText recursively with (res, null, text, null).
        // Inlined here to avoid the function call overhead on every format operation.
        if(null != oCurText && oCurText.text.length > 0)
        {
            var sAccum = oCurText.text;
            oCurText.text = "";
            var accumFmt = null;
            if(-1 != this.Color)
            {
                accumFmt = new AscCommonExcel.Font();
                accumFmt.c = new AscCommonExcel.RgbColor(this.Color);
            }
            var prevLen = res.length;
            var prevItem = prevLen > 0 ? res[prevLen - 1] : null;
            if(null != prevItem && ((null == prevItem.format && null == accumFmt) ||
                (null != prevItem.format && null != accumFmt && accumFmt.isEqual(prevItem.format))))
            {
                prevItem.text += sAccum;
            }
            else
            {
                res.push(null == accumFmt ? {text: sAccum} : {text: sAccum, format: accumFmt});
            }
        }
        if(null != textVal && textVal.length > 0)
        {
            var length = res.length;
            var prev = length > 0 ? res[length - 1] : null;
            if(-1 != this.Color)
            {
                if(null == format)
                    format = new AscCommonExcel.Font();
                format.c = new AscCommonExcel.RgbColor(this.Color);
            }
            if(null != prev && ((null == prev.format && null == format) || (null != prev.format && null != format && format.isEqual(prev.format))))
            {
                prev.text += textVal;
            }
            else
            {
                res.push(null == format ? {text: textVal} : {text: textVal, format: format});
            }
        }
    },
    setFormat: function(format, cultureInfo, formatType, useLocaleFormat) {
        if (null == cultureInfo) {
            cultureInfo = g_oDefaultCultureInfo;
        }
        this.formatString = format;
        this.length = this.formatString.length;
        //string -> tokens
		if (NumFormatType.WordFieldDate === formatType) {
			this.valid = this._parseFormatWordDateTime();
		} else if (NumFormatType.PDFFormDate === formatType) {
			this.valid = this._parseFormatPDFDateTime();
		} else if (NumFormatType.WordFieldNumeric === formatType) {
			this.valid = this._parseFormatWordNumeric("#");
		} else {
			this.valid = this._parseFormat("?", useLocaleFormat);
		}
        if (true == this.valid) {
            //prepare tokens
            // this.valid = formatType != NumFormatType.PDFFormDate ? this._prepareFormat() : this._prepareFormatPDF();
            this.valid = this._prepareFormat();
            if (this.valid) {
                //additional prepare
                var aCurrencySymbols = ["$", "€", "£", "¥", "р.", cultureInfo.CurrencySymbol];
                var sText = "";
                for (var i = 0, length = this.aRawFormat.length; i < length; ++i) {
                    var item = this.aRawFormat[i];
                    if (numFormat_Text == item.type) {
                        sText += item.val;
                    } else if (numFormat_Bracket == item.type) {
						let dbnum = item.val.match(/DBNum(\d)/);
						if (dbnum) {
							this.DBNum = parseInt(dbnum[1]);
						} else {
							if (null != item.CurrencyString) {
								this.bCurrency = true;
								this.CurrencyString = item.CurrencyString;
								sText += item.CurrencyString;
							}
							if (null != item.Lid) {
								//Excel sometimes add 0x10000(0x442 and 0x10442)
								this.LCID = parseInt(item.Lid, 16) & 0xFFFF;
							}
							if (item.bGannen) {
								this.bGannen = true;
							}
							if (AscCommon.isJapanEraLid && AscCommon.isJapanEraLid(item.Lid)
								&& null != item.CalendarId && !item.LidName) {
								this.bJapanEraCalendarOverride = true;
							}
						}
                    }
                    else if (numFormat_DecimalPoint == item.type) {
                        sText += gc_sFormatDecimalPoint;
                    } else if (numFormat_DecimalPointText == item.type) {
                        sText += gc_sFormatDecimalPoint;
                    }
                }
                if ("" != sText) {
                    for (var i = 0, length = aCurrencySymbols.length; i < length; ++i) {
                        if (-1 != sText.indexOf(aCurrencySymbols[i])) {
                            this.bCurrency = true;
                            break;
                        }
                    }
                }
                    }
                }
        return this.valid;
    },
    isInvalidDateValue : function(number)
    {
        return (number == number - 0) && ((number < 0 && !AscCommon.bDate1904) || number > 2958465.9999884);
    },
    _applyGeneralFormat: function(number, nValType, dDigitsCount, bChart, cultureInfo){
        var res = null;
        //todo fIsFitMeasurer and decrease dDigitsCount by other format tokens
        var sGeneral = DecodeGeneralFormat(number, nValType, dDigitsCount);
        if (null != sGeneral) {
            var numFormat = oNumFormatCache.get(sGeneral);
            if (null != numFormat) {
                res = numFormat.format(number, nValType, dDigitsCount, bChart, cultureInfo, true);
            }
        }
        if(!res){
            res = [{text: number.toString()}];
        }
        if (-1 != this.Color) {
            for (var i = 0; i < res.length; ++i) {
                var elem = res[i];
                if (null == elem.format) {
                    elem.format = new AscCommonExcel.Font();
                }
                elem.format.c = new AscCommonExcel.RgbColor(this.Color);
            }
        }
        return res;
    },
	_formatDecimalFrac: function(oParsedNumber) {
		var forceNull = false;
		for (var i = 0; i < this.aRawFormat.length; ++i) {
			var item = this.aRawFormat[i];
			if (numFormat_DecimalFrac == item.type) {
				var frac = oParsedNumber.fraction;
				var numerator = 0;
				var denominator = 0;
				if (item.bNumRight === true) {
					//todo max denominator - 99999
					denominator = item.aRight[0].val;
					numerator = Math.round(denominator * frac);
				} else if (frac > 0) {
					//Continued fraction
					//7 - excel max denominator length
					var denominatorLen = Math.min(7, item.aRight.length);
					var denominatorBound = Math.pow(10, denominatorLen);
					var an = Math.floor(frac);
					var xn1 = frac - an;
					var pn1 = an;
					var qn1 = 1;
					var pn2 = 1;
					var qn2 = 0;
					do {
						an = Math.floor(1 / xn1);
						xn1 = 1 / xn1 - an;
						var pn = an * pn1 + pn2;
						var qn = an * qn1 + qn2;
						pn2 = pn1;
						pn1 = pn;
						qn2 = qn1;
						qn1 = qn;
					} while (qn < denominatorBound);
					numerator = pn2;
					denominator = qn2;
				}
				if (numerator <= 0) {
					numerator = 0;
					if (this.bWhole === false) {
						if (denominator <= 0) {
							denominator = 1;
						}
					} else {
						denominator = 0;
					}
				}
				if (this.bWhole === false) {
					numerator += denominator * oParsedNumber.dec;
				} else if (numerator === denominator && 0 !== numerator) {
					oParsedNumber.dec++;
					numerator = 0;
					denominator = 0;
				}
				if (0 === numerator && 0 === denominator) {
					forceNull = true;
				}
				item.numerator = numerator;
				item.denominator = denominator;
			}
		}
		return forceNull;
	},
    format: function (number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull)
    {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        var cultureInfoLCID = cultureInfo;
        if (null != this.LCID) {
            cultureInfoLCID = g_aCultureInfos[this.LCID] || cultureInfo;
        }
        if(null == nValType)
            nValType = CellValueType.Number;
        var res = [];
        var oCurText = {text: ""};
        if(true == this.valid)
        {
            if(true === this.bDateTime)
            {
                if(this.isInvalidDateValue(number) && this.formatType != AscCommon.NumFormatType.PDFFormDate)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.repeat = true;
                    this._CommitText(res, null, "#", oNewFont);
                    return res;
                }
            }
            var oParsedNumber;
			if (this.formatType == AscCommon.NumFormatType.PDFFormDate)
				oParsedNumber = this._parseNumberForPDFDate(number);
			else
				oParsedNumber = this._parseNumber(number, this.aDecFormat, this.aFracFormat.length, nValType);

            if (true == this.isGeneral() || (true == oParsedNumber.bDigit && true == this.bTextFormat) || (false == oParsedNumber.bDigit && false == this.bTextFormat) || (bChart && this.bGeneralChart))
            {
                return this._applyGeneralFormat(number, nValType, dDigitsCount, bChart, cultureInfo);
            }
			var forceNull = !!opt_forceNull;
			if (this.bSlash) {
				forceNull = this._formatDecimalFrac(oParsedNumber);
			}
            var aDec = [];
            var aFrac = [];
            var aScientific = [];
            if(true == oParsedNumber.bDigit)
            {
                aDec = this._FormatNumber(oParsedNumber.dec, oParsedNumber.exponent, this.aDecFormat.concat(), FormatStates.Decimal, cultureInfo, forceNull);
                aFrac = this._FormatNumber(oParsedNumber.frac, oParsedNumber.exponentFrac, this.aFracFormat.concat(), FormatStates.Frac, cultureInfo);
            }

            var bNoDecFormat = false;
            if((null == aDec || 0 == aDec.length) && 0 != oParsedNumber.dec)
            {
                //case ".00"
                bNoDecFormat = true;
            }
            var hasSign = false;
            var nReadState = FormatStates.Decimal;
            var nFormatLength = this.aRawFormat.length;
			// _g_arabicLCIDs is a module-level Set — O(1) lookup vs O(n) OR-chain.
			const isArabic = _g_arabicLCIDs.has(cultureInfoLCID.LCID);
			
			let _t = this;
			function checkRLM(prev)
			{
				if (!isArabic)
					return;
				
				if (undefined === prev
					|| prev < 0
					|| (numFormat_TimeSeparator !== _t.aRawFormat[prev].type
						&& (numFormat_Text !== _t.aRawFormat[prev].type || ":" !== _t.aRawFormat[prev].val)))
					oCurText.text += "‏";
			}
			
            for(var i = 0; i < nFormatLength; ++i)
            {
                var item = this.aRawFormat[i];
                if(numFormat_Bracket == item.type)
                {
                    if(null != item.CurrencyString)
                        oCurText.text += item.CurrencyString;
                }
                else if(numFormat_DecimalPoint == item.type)
                {
                    if(bNoDecFormat && null != oParsedNumber.dec && FormatStates.Decimal == nReadState)
                    {
                        oCurText.text += oParsedNumber.dec;
                    }
					oCurText.text += cultureInfo.NumberDecimalSeparator;
                    nReadState = FormatStates.Frac;
                }
                else if (numFormat_DecimalPointText == item.type) {
                    oCurText.text += cultureInfo.NumberDecimalSeparator;
                }
                else if (numFormat_ThousandText == item.type) {
                    oCurText.text += cultureInfo.NumberGroupSeparator;
                }
                else if(this._isDigitType(item.type))
                {
                    var text = null;
                    if(nReadState == FormatStates.Decimal)
                        text = aDec.shift();
                    else if(nReadState == FormatStates.Frac)
                        text = aFrac.shift();
                    else if(nReadState == FormatStates.Scientific)
                        text = aScientific.shift();
                    if(null != text)
                    {
                        this._AddDigItem(res, oCurText, text);
                    }
                }
                else if(numFormat_Text == item.type)
                {
					if(',' === item.val && isArabic) {
						oCurText.text += "،";
					} else {
						oCurText.text += item.val;
					}
                }
                else if(numFormat_TextPlaceholder == item.type)
                {
                    oCurText.text += number;
                }
                else if(numFormat_Scientific == item.type)
                {
                    if(null != item.format)
                    {
                        oCurText.text += item.val;

                        if(oParsedNumber.scientific < 0)
                            oCurText.text += "-";
                        else if(item.sign == SignType.Positive)
                            oCurText.text += "+";

                        
                        aScientific = this._FormatNumber(Math.abs(oParsedNumber.scientific), 0, item.format.concat(), FormatStates.Scientific, cultureInfo);
                        nReadState = FormatStates.Scientific;
                    }
                }
                else if(numFormat_DecimalFrac == item.type)
                {
                    var curForceNull = this.bWhole === false;
					var aLeft = this._FormatNumber(item.numerator, 0, item.aLeft.concat(), FormatStates.Slash, cultureInfo, curForceNull);
					for (var j = 0, length = aLeft.length; j < length; ++j) {
						var subitem = aLeft[j];
						if (subitem) {
							this._AddDigItem(res, oCurText, subitem);
						}
					}
					if ((item.numerator > 0 && item.denominator > 0) || curForceNull) {
						oCurText.text += "/";
					} else {
						var oNewFont = new AscCommonExcel.Font();
						oNewFont.skip = true;
						this._CommitText(res, oCurText, "/", oNewFont);
					}
					if (item.bNumRight === true) {
						var rightVal = item.aRight[0].val;
						if (rightVal) {
							if (item.denominator > 0) {
								oCurText.text += rightVal;
							} else {
								for (var rightIdx = 0; rightIdx < rightVal.toString().length; ++rightIdx) {
									var oNewFont = new AscCommonExcel.Font();
									oNewFont.skip = true;
									this._CommitText(res, oCurText, "0", oNewFont);
								}
							}
						}
					} else {
						var aRight = this._FormatNumber(item.denominator, 0, item.aRight.concat(), FormatStates.SlashFrac, cultureInfo);
						for (var j = 0, length = aRight.length; j < length; ++j) {
							var subitem = aRight[j];
							if (subitem) {
								this._AddDigItem(res, oCurText, subitem);
							}
						}
					}
                }
                else if(numFormat_Repeat == item.type)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.repeat = true;
                    this._CommitText(res, oCurText, item.val, oNewFont);
                }
                else if(numFormat_Skip == item.type)
                {
                    var oNewFont = new AscCommonExcel.Font();
					oNewFont.skip = true;
                    this._CommitText(res, oCurText, item.val, oNewFont);
                }
				else if(numFormat_DateSeparator == item.type)
                {
                    oCurText.text += cultureInfo.DateSeparator;
				}
				else if(numFormat_TimeSeparator == item.type)
                {
                    oCurText.text += cultureInfo.TimeSeparator;
				}
				else if(numFormat_DayOfWeek == item.type)
				{
					if (item.val === 3)
					{
						oCurText.text += cultureInfoLCID.AbbreviatedDayNames[oParsedNumber.date.dayWeek];
					}
					else if (item.val > 3)
					{
						oCurText.text += cultureInfoLCID.DayNames[oParsedNumber.date.dayWeek];
					}
					else
					{
						checkRLM();
						oCurText.text += 'a'.repeat(item.val);
					}
				}
                else if(numFormat_JapanEra == item.type)
                {
                    var nVal = Math.min(item.val, 3);
                    var resolved = this._resolveJapanEra(oParsedNumber);
                    var era = resolved.era;
                    if (era) {
                        if (nVal === 1) {
                            oCurText.text += era.latinShort;
                        } else if (nVal === 2) {
                            oCurText.text += era.kanjiShort;
                        } else {
                            oCurText.text += era.kanjiFull;
                        }
                    } else if (resolved.bEraLcid) {
                        //Pre-Meiji date under Japanese LCID: keep literal g.
                        //Non-era LCID: token disappears (Excel oracle).
                        oCurText.text += 'g'.repeat(nVal);
                    }
                }
                else if(numFormat_JapanEraYear == item.type)
                {
                    var nVal = Math.min(item.val, 2);
                    var resolved = this._resolveJapanEra(oParsedNumber);
                    var era = resolved.era;
                    if (era) {
                        var eraYear = oParsedNumber.date.year - era.startYear + 1;
                        if (this.bGannen && this.bHasKanjiEra && eraYear === 1) {
                            oCurText.text += '\u5143';
                        } else if (nVal === 1) {
                            oCurText.text += eraYear;
                        } else {
                            oCurText.text += this._ZeroPad(eraYear);
                        }
                    } else if (resolved.bEraLcid) {
                        oCurText.text += 'e'.repeat(nVal);
                    } else {
                        //Non-era LCID: e/ee renders Gregorian year (Excel oracle).
                        oCurText.text += oParsedNumber.date.year;
                    }
                }
                else if(numFormat_Year == item.type)
                {
                  if (item.val > 0) {
					  checkRLM();
                    if (item.val <= 2) {
						oCurText.text += (oParsedNumber.date.year.toString().slice(-2));
                    } else {
						if (oParsedNumber.date.year.toString().length < 4)
                    		oCurText.text += '0' + oParsedNumber.date.year;
						else
							oCurText.text += oParsedNumber.date.year;
                    }
                  }
                }
                else if(numFormat_Month == item.type)
                {
                    var m = oParsedNumber.date.month;
					if (item.val === 1) {
						checkRLM();
						oCurText.text += m + 1;
					} else if (item.val === 2) {
						checkRLM();
						oCurText.text += this._ZeroPad(m + 1);
					}
                    else if (item.val == 3) {
                        if (this.bDay && cultureInfoLCID.AbbreviatedMonthGenitiveNames.length > 0)
                            oCurText.text += cultureInfoLCID.AbbreviatedMonthGenitiveNames[m];
                        else
                            oCurText.text += cultureInfoLCID.AbbreviatedMonthNames[m];
                    }
                    else if (item.val == 5) {
                        var sMonthName = cultureInfoLCID.MonthNames[m];
                        if (sMonthName.length > 0)
                            oCurText.text += sMonthName[0];
                    }
                    else if (item.val > 0){
                        if (this.bDay && cultureInfoLCID.MonthGenitiveNames.length > 0)
                            oCurText.text += cultureInfoLCID.MonthGenitiveNames[m];
                        else
                            oCurText.text += cultureInfoLCID.MonthNames[m];
                    }
                }
                else if(numFormat_Day == item.type)
                {
                    if(item.val == 1) {
						checkRLM();
						oCurText.text += oParsedNumber.date.d;
					} else if(item.val === 2) {
						checkRLM();
						oCurText.text += this._ZeroPad(oParsedNumber.date.d);
					}
                    else if(item.val == 3)
                        oCurText.text += cultureInfoLCID.AbbreviatedDayNames[oParsedNumber.date.dayWeek];
                    else if(item.val > 0)
                        oCurText.text += cultureInfoLCID.DayNames[oParsedNumber.date.dayWeek];
                    
                }
                else if(numFormat_Hour == item.type)
                {
                    var h = oParsedNumber.date.hour;
                    if(item.bElapsed === true)
                        h = oParsedNumber.date.countDay*24 + oParsedNumber.date.hour;
                    if(this.bTimePeriod === true)
                        h = h%12||12;
					
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += h;
						else
							oCurText.text += this._ZeroPad(h);
					}
                }
                else if(numFormat_Minute == item.type)
                {
                    var min = oParsedNumber.date.min;
                    if(item.bElapsed === true)
                        min = oParsedNumber.date.countDay*24*60 + oParsedNumber.date.hour*60 + oParsedNumber.date.min;
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += min;
						else
							oCurText.text += this._ZeroPad(min);
					}
                }
                else if(numFormat_Second == item.type)
                {
                    var s = oParsedNumber.date.sec;
                    if(this.bMillisec === false)
                        s = oParsedNumber.date.sec + Math.round(oParsedNumber.date.ms/1000);
                    if(item.bElapsed === true)
                        s = oParsedNumber.date.countDay*24*60*60 + oParsedNumber.date.hour*60*60 + oParsedNumber.date.min*60 + s;
	
					if (item.val > 0) {
						checkRLM(i - 1);
						if (item.val === 1)
							oCurText.text += s;
						else
							oCurText.text += this._ZeroPad(s);
					}
                }
                else if (numFormat_AmPm == item.type) {
                    if (cultureInfoLCID.AMDesignator.length > 0 && cultureInfoLCID.PMDesignator.length > 0)
                        oCurText.text += (oParsedNumber.date.hour < 12) ? cultureInfoLCID.AMDesignator : cultureInfoLCID.PMDesignator;
                    else
                        oCurText.text += (oParsedNumber.date.hour < 12) ? "AM" : "PM";
                }
                else if (numFormat_Milliseconds == item.type) {
                    var nMsFormatLength = item.format.length;
                    var dMs = oParsedNumber.date.ms;
                    //Round
                    if (nMsFormatLength < 3) {
                        var dTemp = dMs / Math.pow(10, 3 - nMsFormatLength);
                        dTemp = Math.round(dTemp);
                        dMs = dTemp * Math.pow(10, 3 - nMsFormatLength);
                    }
                    var nExponent = 0;
                    if(0 == dMs)
                        nExponent = -1;
                    else if (dMs < 10)
                        nExponent = -2;
                    else if (dMs < 100)
                        nExponent = -1;
                    var aMilSec = this._FormatNumber(dMs, nExponent, item.format.concat(), FormatStates.Frac, cultureInfo);
					checkRLM(i - 1);
                    for (var k = 0; k < aMilSec.length; k++)
                        this._AddDigItem(res, oCurText, aMilSec[k]);
                }
                else if (numFormat_General == item.type) {
                    this._CommitText(res, oCurText, null, null);
                    //todo minus sign
                    res = res.concat(this._applyGeneralFormat(Math.abs(number), nValType, dDigitsCount, bChart, cultureInfo));
                } else if (numFormat_Plus == item.type) {
					hasSign = true;
					if (number > 0) {
						oCurText.text += '+';
					} else if (number < 0) {
						oCurText.text += '-';
					} else {
						oCurText.text += ' ';
					}
				} else if (numFormat_Minus == item.type) {
					hasSign = true;
					if (number < 0) {
						oCurText.text += '-';
					} else {
						oCurText.text += ' ';
					}
				}
            }

			if (true == this.bAddMinusIfNes && SignType.Negative == oParsedNumber.sign && !hasSign) {
				// No explicit sign placeholder consumed the minus — prepend it now.
				res.unshift({text: "-"});
			}
            this._CommitText(res, oCurText, null, null);
			if(0 == res.length)
                res = [{text: ""}];
        }
        else
        {
            if(0 == res.length)
                res = [{text: number.toString()}];
        }
		//the length of the resulting string should not exceed c_oAscMaxColumnWidth
		var nLen = 0;
		for(var i = 0; i < res.length; ++i){
			var elem = res[i];
			if (elem.text) {
				elem.text = this._replaceDBNumDigit(elem.text);
				nLen += elem.text.length;
			}
		}
		if(nLen > Asc.c_oAscMaxColumnWidth){
			var oNewFont = new AscCommonExcel.Font();
			oNewFont.repeat = true;
			res = [{text: "#", format: oNewFont}];
		}
        return res;
    },
	shiftFormat : function(output, nShift, useLocaleFormat) {
		if (this.bDateTime || this.bSlash || this.bTextFormat)
			return false;
		output.format = this.toString(nShift, useLocaleFormat);
		return true;
	},
    toString : function(nShift, useLocaleFormat, options)
    {
		var sGeneral;
		var DecimalSeparator;
		var GroupSeparator;
		var TimeSeparator;
		var year;
		var month;
		var day;
		var hour;
		var minute;
		var second;
		var dayOfWeek;
		var era;
		if (useLocaleFormat) {
			sGeneral = LocaleFormatSymbol['general'];
			DecimalSeparator = g_oDefaultCultureInfo.NumberDecimalSeparator;
			TimeSeparator = g_oDefaultCultureInfo.TimeSeparator;
			GroupSeparator = g_oDefaultCultureInfo.NumberGroupSeparator;
			if (LocaleFormatSymbol['M'] === LocaleFormatSymbol['m']) {
				year = LocaleFormatSymbol['Y'];
				month = LocaleFormatSymbol['M'];
				day = LocaleFormatSymbol['D'];
			} else {
				year = LocaleFormatSymbol['y'];
				month = LocaleFormatSymbol['m'];
				day = LocaleFormatSymbol['d'];
			}
			hour = LocaleFormatSymbol['h'];
			minute = LocaleFormatSymbol['minute'];
			second = LocaleFormatSymbol['s'];
			dayOfWeek = LocaleFormatSymbol['a'];
			era = LocaleFormatSymbol['g'];
		} else {
			sGeneral = AscCommon.g_cGeneralFormat;
			DecimalSeparator = gc_sFormatDecimalPoint;
			TimeSeparator = ':';
			GroupSeparator = gc_sFormatThousandSeparator;
			year = 'y';
			month = 'm';
			day = 'd';
			hour = 'h';
			minute = 'm';
			second = 's';
			dayOfWeek = 'a';
			era = 'g';
		}
        var nDecLength = this.aDecFormat.length;
        var nDecIndex = 0;
        var nFracLength = this.aFracFormat.length;
        var nFracIndex = 0;
        var nNewFracLength = nFracLength + nShift;
        if(nNewFracLength < 0)
            nNewFracLength = 0;
        var nReadState = FormatStates.Decimal;
        var res = "";
        // Use the module-level helper to avoid re-creating the function on every toString() call.
        const fFormatToString = _formatDigitArrayToString;
        //Color
        if(null != this.Color)
        {
            switch(this.Color)
            {
            case 0x000000: res += "[Black]";break;
            case 0x0000ff: res += "[Blue]";break;
            case 0x00ffff: res += "[Cyan]";break;
            case 0x00ff00: res += "[Green]";break;
            case 0xff00ff: res += "[Magenta]";break;
            case 0xff0000: res += "[Red]";break;
            case 0xffffff: res += "[White]";break;
            case 0xffff00: res += "[Yellow]";break;
            }
        }
		//Comporation operator
        if(null != this.ComporationOperator)
        {
			switch(this.ComporationOperator.operator)
			{
				case NumComporationOperators.equal: res += "[=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.greater: res += "[>" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.less: res += "[<" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.greaterorequal: res += "[>=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.lessorequal: res += "[<=" + this.ComporationOperator.operatorValue +"]";break;
				case NumComporationOperators.notequal: res += "[<>" + this.ComporationOperator.operatorValue +"]";break;
			}
		}
		if (this.DBNum > 0)
		{
			res += '[DBNum' + this.DBNum + ']';
		}

		var bGannenFallback = options && options.gannenFallback;
		var bStripGannenFallback = false;
		if (bGannenFallback && this.bHasKanjiEra) {
			for (var nGannenItem = 0; nGannenItem < this.aRawFormat.length; ++nGannenItem) {
				if (this.aRawFormat[nGannenItem] && this.aRawFormat[nGannenItem].type === numFormat_JapanEraYear) {
					bStripGannenFallback = true;
					break;
				}
			}
		}

        var nFormatLength = this.aRawFormat.length;
        for(var i = 0; i < nFormatLength; ++i)
        {
            var item = this.aRawFormat[i];
            if(numFormat_Bracket == item.type)
            {
                if (bGannenFallback && item.bGannen && bStripGannenFallback)
                {
                    res += "[$]";
                }
                else if(null != item.CurrencyString || null != item.Lid)
                {
                    res += "[$";
                    if(null != item.CurrencyString)
                        res += item.CurrencyString;
                    if (null != item.Lid) {
                        res += "-";
                        //Canonical Gannen form is BCP-47; numeric [$-411,x-gannen] is not
                        //a valid Excel Gannen format and is dropped on round-trip.
                        if (!bGannenFallback && item.bGannen && AscCommon.isJapanEraLid && AscCommon.isJapanEraLid(item.Lid)) {
                            res += "ja-JP-x-gannen";
                        } else if (item.LidName === "ja-jp" && null != item.CalendarId) {
                            //[$-ja-JP,80] keeps era rendering; [$-411,80] does not.
                            res += "ja-JP";
                        } else {
                            res += item.Lid;
                        }
                        if (null != item.CalendarId) {
                            res += ",";
                            res += item.CalendarId;
                        }
                    }
                    res += "]";
                }
            }
            else if(numFormat_DecimalPoint == item.type)
            {
                nReadState = FormatStates.Frac;
                if(0 != nNewFracLength)
                    res += DecimalSeparator;
            }
            else if (numFormat_DecimalPointText == item.type) {
                res += DecimalSeparator;
            }
            else if(numFormat_Thousand == item.type || numFormat_ThousandText == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += GroupSeparator;
            }
            else if(this._isDigitType(item.type))
            {
                if(FormatStates.Decimal == nReadState)
                    nDecIndex++;
                else
                    nFracIndex++;
                if(nReadState == FormatStates.Frac && nFracIndex > nNewFracLength)
                    ;
                else
                {
                    var sCurSimbol;
                    if(numFormat_Digit == item.type)
                        sCurSimbol = "0";
                    else if(numFormat_DigitNoDisp == item.type)
                        sCurSimbol = "#";
                    else if(numFormat_DigitSpace == item.type)
                        sCurSimbol = "?";
					else if(numFormat_DigitDrop == item.type)
						sCurSimbol = "x";
                    res += sCurSimbol;
                    if(nReadState == FormatStates.Frac && nFracIndex == nFracLength)
                    {
                        for(var j = 0; j < nShift; ++j)
                            res += sCurSimbol;
                    }
                }
                if(0 == nFracLength && nShift > 0 && FormatStates.Decimal == nReadState && nDecIndex == nDecLength)
                {
                    res += gc_sFormatDecimalPoint;
                    for(var j = 0; j < nShift; ++j)
                        res += "0";
                }
            }
            else if(numFormat_Text == item.type)
            {
                if("%" == item.val)
                    res += item.val;
                else
                    res += "\"" + item.val + "\"";
            }
            else if(numFormat_TextPlaceholder == item.type)
                res += "@";
            else if(numFormat_Scientific == item.type)
            {
                nReadState = FormatStates.Scientific;
                res += item.val;
                if(item.sign == SignType.Positive)
                    res += "+";
                else
                    res += "-";
            }
            else if(numFormat_DecimalFrac == item.type)
            {
                res += fFormatToString(item.aLeft);
                res += "/";
                res += fFormatToString(item.aRight);
            }
            else if(numFormat_Repeat == item.type)
                res += "*" + item.val;
            else if(numFormat_Skip == item.type)
                res += "_" + item.val;
			else if(numFormat_DateSeparator == item.type)
                res += "/";
			else if(numFormat_TimeSeparator == item.type)
                res += TimeSeparator;
            else if(numFormat_Year == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += year;
            }
            else if(numFormat_Month == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += month;
            }
            else if(numFormat_Day == item.type)
            {
                for(var j = 0; j < item.val; ++j)
                    res += day;
            }
            else if(numFormat_Hour == item.type)
            {
				if (item.bElapsed) {
					res += "[";
				}
				for(var j = 0; j < item.val; ++j)
					res += hour;
				if (item.bElapsed) {
					res += "]";
				}
            }
            else if(numFormat_Minute == item.type)
            {
                if (item.bElapsed) {
                    res += "[";
                }
                for(var j = 0; j < item.val; ++j)
                    res += minute;
                if (item.bElapsed) {
                    res += "]";
                }
            }
            else if(numFormat_Second == item.type)
            {
                if (item.bElapsed) {
                    res += "[";
                }
                for(var j = 0; j < item.val; ++j)
                    res += second;
                if (item.bElapsed) {
                    res += "]";
                }
            }
			else if(numFormat_DayOfWeek == item.type)
			{
				var nIndex = (item.val > 3) ? 3 : item.val;
				for(var j = 0; j < nIndex; ++j)
					res += dayOfWeek;
			}
            else if(numFormat_JapanEra == item.type)
            {
                var nIndex = Math.min(item.val, 3);
                for(var j = 0; j < nIndex; ++j)
                    res += era;
            }
            else if(numFormat_JapanEraYear == item.type)
            {
                //Lowercase: parser precedence is e -> era-year, E -> scientific/literal.
                var nIndex = Math.min(item.val, 2);
                for(var j = 0; j < nIndex; ++j)
                    res += "e";
            }
            else if(numFormat_AmPm == item.type)
                res += "AM/PM";
            else if(numFormat_Milliseconds == item.type)
                res += fFormatToString(item.format);
			else if(numFormat_Plus == item.type)
				res += "+";
			else if(numFormat_Minus == item.type)
				res += "-";
			else if(numFormat_General == item.type)
				res += sGeneral;
        }
        return res;
    },
	getFormatCellsInfo: function() {
		var info = new Asc.asc_CFormatCellsInfo();
		info.asc_setDecimalPlaces(this.aFracFormat.length);
		info.asc_setSeparator(this.bThousandSep);
		info.asc_setSymbol(this.LCID);
		info.asc_setCurrencySymbol(this.CurrencyString);
		return info;
	},
	isGeneral: function() {
		return 1 == this.aRawFormat.length && numFormat_General == this.aRawFormat[0].type;
	}
};
function NumFormatCache()
{
    this.oNumFormats = {};
}
NumFormatCache.prototype =
{
	cleanCache : function(){
		this.oNumFormats = {};
	},
    get : function(format, formatType)
    {
		var key = format + String.fromCharCode(5) + formatType;
        var res = this.oNumFormats[key];
        if(null == res)
        {
            res = new CellFormat(format, formatType, false);
            this.oNumFormats[key] = res;
        }
        return res;
    }
};
//cache of structures by format string
var oNumFormatCache = new NumFormatCache();

// Strip \x escaping and "text" quoting from a format string.
// '#\ ?/?' and '#" "?/?' and '# ?/?' all normalize to '# ?/?'.
function stripFormatEscaping(s) {
	var r = '';
	for (var i = 0; i < s.length; i++) {
		var c = s[i];
		if (c === '\\' && i + 1 < s.length) { r += s[++i]; }
		else if (c === '"') { for (++i; i < s.length && s[i] !== '"'; i++) { r += s[i]; } }
		else { r += c; }
	}
	return r;
}

function CellFormat(format, formatType, useLocaleFormat)
{
    this.sFormat = format;
    this.oPositiveFormat = null;
    this.oNegativeFormat = null;
    this.oNullFormat = null;
    this.oTextFormat = null;
	this.aComporationFormats = null;
    var aFormats = format.split(";");
	var aParsedFormats = [];
	for(var i = 0; i < aFormats.length; ++i)
	{
    var sNewFormat = aFormats[i];
    //if sNewFormat ends with an odd number of '\', it means ';' was escaped and this is text
    while(true){
      var formatTail = sNewFormat.match(/\\+$/g);
      if (formatTail && formatTail.length > 0 && 1 === formatTail[0].length % 2 && i + 1 < aFormats.length) {
        sNewFormat += ';';
        sNewFormat += aFormats[++i];
      } else {
        break;
      }
    }
		var oNewFormat = new NumFormat(false);
		oNewFormat.setFormat(sNewFormat, undefined, formatType, useLocaleFormat);
		if (oNewFormat.LCID === 0xF800) {
			sNewFormat = '[$-F800]' + g_oDefaultCultureInfo.LongDatePattern;
			oNewFormat = new NumFormat(false);
			oNewFormat.setFormat(sNewFormat, undefined, formatType, useLocaleFormat);
		}
		aParsedFormats.push(oNewFormat);
	}
  var nFormatsLength = aParsedFormats.length;
	var noComparisonn = aParsedFormats.every(function(format) {return !format.ComporationOperator});
	if(noComparisonn)
	{
		if(4 <= nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = aParsedFormats[2];
			this.oTextFormat = aParsedFormats[3];
			//for ';;;' format, if 4 formats exist fourth always used for text
			this.oTextFormat.bTextFormat = true;
		}
		else if(3 == nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = aParsedFormats[2];
			this.oTextFormat = this.oPositiveFormat;
			if (this.oNullFormat.bTextFormat) {
				this.oTextFormat = this.oNullFormat;
				this.oNullFormat = this.oPositiveFormat;
			}
		}
		else if(2 == nFormatsLength)
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oNegativeFormat = aParsedFormats[1];
			this.oNullFormat = this.oPositiveFormat;
			this.oTextFormat = this.oPositiveFormat;
			if (this.oNegativeFormat.bTextFormat) {
				this.oTextFormat = this.oNegativeFormat;
				this.oNegativeFormat = this.oPositiveFormat;
				this.oPositiveFormat.bAddMinusIfNes = true;
			}
		}
		else
		{
			this.oPositiveFormat = aParsedFormats[0];
			this.oPositiveFormat.bAddMinusIfNes = true;
			this.oNegativeFormat = this.oPositiveFormat;
			this.oNullFormat = this.oPositiveFormat;
			this.oTextFormat = this.oPositiveFormat;
		}
	}
	else
	{
		this.oTextFormat = new NumFormat(false);
		this.oTextFormat.setFormat("@", undefined, undefined, useLocaleFormat);
		//based on experiments, if the comparison operator crosses 0, then the minus sign needs to be added depending on the value
		//example [<100] needs to add sign, [<-100] no need to add sign
		for (let i = 0; i < aParsedFormats.length && i < 2; ++i) {
			let oCurFormat = aParsedFormats[i];
			if (oCurFormat.ComporationOperator) {
				let operator = oCurFormat.ComporationOperator.operator;
				let operatorValue = oCurFormat.ComporationOperator.operatorValue;
				if (0 < operatorValue && (operator === NumComporationOperators.less || operator === NumComporationOperators.lessorequal))
					oCurFormat.bAddMinusIfNes = true;
				else if (0 > operatorValue && (operator === NumComporationOperators.greater || operator === NumComporationOperators.greaterorequal))
					oCurFormat.bAddMinusIfNes = true;
			}
		}
		if (aParsedFormats.length > 2) {
			aParsedFormats[2].bAddMinusIfNes = true;
		}
		this.aComporationFormats = aParsedFormats.slice(0, 3);
	}
    this.formatCache = {};
}
CellFormat.prototype =
{
	isTextFormat : function()
	{
		if (this.oPositiveFormat  != null) {
			return this.oPositiveFormat.bTextFormat;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bTextFormat;
		}
		return false;
	},
	isGeneralFormat : function()
	{
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.isGeneral();
		} else if (this.aComporationFormats != null  && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].isGeneral();
		}
		return false;
	},
	isDateTimeFormat : function()
	{
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.bDateTime;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bDateTime;
		}
		return false;
	},
	isTimeFormat : function() {
		if (this.oPositiveFormat != null) {
			return this.oPositiveFormat.bTime;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bTime;
		}
		return false;
	},
	isDateFormat : function() {
		if ( this.oPositiveFormat != null) {
			return this.oPositiveFormat.bDate;
		} else if (this.aComporationFormats != null && this.aComporationFormats.length > 0) {
			return this.aComporationFormats[0].bDate;
		}
		return false;
	},
	getTextFormat: function () {
	    var oRes = null;
	    if (null == this.aComporationFormats) {
	        if (null != this.oTextFormat && this.oTextFormat.bTextFormat)
	            oRes = this.oTextFormat;
	    } else {
	        for (var i = 0, length = this.aComporationFormats.length; i < length ; ++i) {
	            var oCurFormat = this.aComporationFormats[i];
	            if (null == oCurFormat.ComporationOperator && oCurFormat.bTextFormat) {
	                oRes = oCurFormat;
	                break;
	            }
	        }
	    }
	    return oRes;
	},
	getFormatByValue : function(dNumber)
	{
		var oRes = null;
		if(null == this.aComporationFormats)
		{
			if(dNumber > 0 && null != this.oPositiveFormat)
				oRes = this.oPositiveFormat;
			else if(dNumber < 0 && null != this.oNegativeFormat)
				oRes = this.oNegativeFormat;
			else if(null != this.oNullFormat)
				oRes = this.oNullFormat;
		}
		else
		{
			//todo only 4 formats allowed in aComporationFormats
			// Handle text values - use text format if available
			if(typeof dNumber === 'string')
			{
				// Look for text format (usually at index 3)
				for (let i = 0; i < this.aComporationFormats.length; ++i) {
					let oCurFormat = this.aComporationFormats[i];
					if (null == oCurFormat.ComporationOperator && oCurFormat.bTextFormat) {
						oRes = oCurFormat;
						break;
					}
				}
				if (null == oRes && this.aComporationFormats.length > 3 && this.aComporationFormats[3]) {
					oRes = this.aComporationFormats[3];
				}
			}
			else
			{
				// Process all conditional formats in order
				for (let i = 0; i < this.aComporationFormats.length; ++i)
				{
					let oCurFormat = this.aComporationFormats[i];
					let oOperationValue, operator;
					
					// Skip text format
					if (null == oCurFormat.ComporationOperator && oCurFormat.bTextFormat) {
						continue;
					}
					
					if (null != oCurFormat.ComporationOperator) {
						operator = oCurFormat.ComporationOperator.operator;
						oOperationValue = oCurFormat.ComporationOperator.operatorValue;
					} else if(0 === i) {
						oOperationValue = 0;
						operator = NumComporationOperators.greater;
					} else if(1 === i && this.aComporationFormats.length > 2 && !this.aComporationFormats[2].bTextFormat) {
						oOperationValue = 0;
						operator = NumComporationOperators.less;
					} else if(!oCurFormat.bTextFormat) {
						//fallback
						oRes = oCurFormat;
					} else {
						break;
					}
					
					let isMatch = (operator === NumComporationOperators.equal && dNumber === oOperationValue) ||
						(operator === NumComporationOperators.greater && dNumber > oOperationValue) ||
						(operator === NumComporationOperators.less && dNumber < oOperationValue) ||
						(operator === NumComporationOperators.greaterorequal && dNumber >= oOperationValue) ||
						(operator === NumComporationOperators.lessorequal && dNumber <= oOperationValue) ||
						(operator === NumComporationOperators.notequal && dNumber !== oOperationValue);
					if (isMatch) {
						oRes = oCurFormat;
						break;
					}
				}
			}
		}
		return oRes;
	},
    format : function(number, nValType, dDigitsCount, bChart, cultureInfo, opt_withoutCache, opt_forceNull)
    {
        var res = null;
        if (null == bChart)
            bChart = false;
        var lcid = cultureInfo ? cultureInfo.LCID : 0;
        var cacheKey, cacheVal;
        if (!opt_withoutCache) {
            cacheKey = number + '-' + nValType + '-' + dDigitsCount + '-' + lcid;
            cacheVal = this.formatCache[cacheKey];
            if(null != cacheVal)
            {
                if (bChart)
                    res = cacheVal.chart;
                else
                    res = cacheVal.nochart;
                if (null != res)
                    return res;
            }
        }
        res = [{text: number.toString()}];
        var dNumber = number - 0;
        var oFormat = null;
		if(CellValueType.String != nValType && number == dNumber)
		{
			oFormat = this.getFormatByValue(dNumber);
			if(null != oFormat)
			    res = oFormat.format(number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull);
			else if(null != this.aComporationFormats)
			{
			    var oNewFont = new AscCommonExcel.Font();
				oNewFont.repeat = true;
				res = [{text: "#", format: oNewFont}];
			}
		}
		else
		{
			//text
		    if (null != this.oTextFormat) {
		        oFormat = this.oTextFormat;
		        res = oFormat.format(number, nValType, dDigitsCount, cultureInfo, bChart, opt_forceNull);
		    }
		}
        if (!opt_withoutCache) {
            if (null == cacheVal) {
                cacheVal = {chart: null, nochart: null};
                this.formatCache[cacheKey] = cacheVal;
            }
            if (null != oFormat && oFormat.bGeneralChart) {
                if (bChart)
                    cacheVal.chart = res;
                else
                    cacheVal.nochart = res;
            }
            else {
                cacheVal.chart = res;
                cacheVal.nochart = res;
            }
        }
        return res;
    },
    shiftFormat : function(output, nShift, useLocaleFormat)
    {
        var bRes = false;
        var bCurRes = true;
		if(null == this.aComporationFormats)
		{
			bCurRes = this.oPositiveFormat.shiftFormat(output, nShift, useLocaleFormat);
			if(false == bCurRes)
				output.format = this.oPositiveFormat.formatString;
			bRes |= bCurRes;
			if(null != this.oNegativeFormat && this.oPositiveFormat != this.oNegativeFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oNegativeFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oNegativeFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
			if(null != this.oNullFormat && this.oPositiveFormat != this.oNullFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oNullFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oNullFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
			if(null != this.oTextFormat && this.oPositiveFormat != this.oTextFormat)
			{
				var oTempOutput = {};
				bCurRes = this.oTextFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(false == bCurRes)
					output.format += ";" + this.oTextFormat.formatString;
				else
					output.format += ";" + oTempOutput.format;
				bRes |= bCurRes;
			}
		}
		else
		{
			var length = this.aComporationFormats.length;
			output.format = "";
			for(var i = 0; i < length; ++i)
			{
				var oTempOutput = {};
				var oCurFormat = this.aComporationFormats[i];
				var bCurRes = oCurFormat.shiftFormat(oTempOutput, nShift, useLocaleFormat);
				if(0 != i)
					output.format += ";";
				if(false == bCurRes)
					output.format += oCurFormat.formatString;
				else
					output.format += oTempOutput.format;
				bRes |= bCurRes;
			}
		}
        return bRes;
    },
	toString: function(nShift, useLocaleFormat, options) {
		var res = '';
		if (null == this.aComporationFormats) {
			res += this.oPositiveFormat.toString(nShift, useLocaleFormat, options);
			if (null != this.oNegativeFormat && this.oPositiveFormat != this.oNegativeFormat) {
				res += ";" + this.oNegativeFormat.toString(nShift, useLocaleFormat, options);
			}
			if (null != this.oNullFormat && this.oPositiveFormat != this.oNullFormat) {
				res += ";" + this.oNullFormat.toString(nShift, useLocaleFormat, options);
			}
			if (null != this.oTextFormat && this.oPositiveFormat != this.oTextFormat) {
				res += ";" + this.oTextFormat.toString(nShift, useLocaleFormat, options);
			}
		}
		else {
			var length = this.aComporationFormats.length;
			for (var i = 0; i < length; ++i) {
				var oCurFormat = this.aComporationFormats[i];
				if (0 != i) {
					res += ";";
				}
				res += oCurFormat.toString(nShift, useLocaleFormat, options);
			}
		}
		return res;
	},
	formatToMathInfo : function(number, nValType, dDigitsCount)
	{
		return this._formatToText(number, nValType, dDigitsCount, false);
	},
	formatToChart : function(number, dDigitsCount, cultureInfo)
	{
		return this._formatToText(number, CellValueType.Number, dDigitsCount || gc_nMaxDigCount, true, cultureInfo);
	},
	formatToWord : function(number, dDigitsCount, cultureInfo)
	{
		return this._formatToText(number, CellValueType.Number, dDigitsCount || gc_nMaxDigCount, false, cultureInfo, true);
	},
	_formatToText : function(number, nValType, dDigitsCount, bChart, cultureInfo, opt_forceNull)
	{
		var result = "";
		var arrFormat = this.format(number, nValType, dDigitsCount, bChart, cultureInfo, undefined, opt_forceNull);
		for (var i = 0, item; i < arrFormat.length; ++i) {
			item = arrFormat[i];
			if (item.format) {
				if (item.format.repeat)
					continue;
				if (item.format.skip) {
					result += " ";
					continue;
				}
			}
			if (item.text)
				result += item.text;
		}
		return result;
	},
	getType: function() {
		return this.getTypeInfo().type;
	},
	getTypeInfo: function() {
		var info;
		if (null != this.oPositiveFormat) {
			info = this.oPositiveFormat.getFormatCellsInfo();
			info.asc_setType(this._getType(this.oPositiveFormat));
		} else if (null != this.aComporationFormats && this.aComporationFormats.length > 0) {
			info = this.aComporationFormats[0].getFormatCellsInfo();
			info.asc_setType(this._getType(this.aComporationFormats[0]));
		} else {
			info = new Asc.asc_CFormatCellsInfo();
			info.asc_setType(c_oAscNumFormatType.General);
			info.asc_setDecimalPlaces(0);
			info.asc_setSeparator(false);
			info.asc_setSymbol(null);
		}
		return info;
	},
	_getType: function(format) {
		var nType = c_oAscNumFormatType.Custom;
		if (format.isGeneral()) {
			nType = c_oAscNumFormatType.General;
		}
		else if (format.bDateTime) {
			if (format.bDate) {
				nType = c_oAscNumFormatType.Date;
			} else {
				nType = c_oAscNumFormatType.Time;
			}
		}
		else if (format.bCurrency) {
			if (format.bRepeat) {
				nType = c_oAscNumFormatType.Accounting;
			} else {
				nType = c_oAscNumFormatType.Currency;
			}
		} else {
			var info = format.getFormatCellsInfo();
			var normalized = stripFormatEscaping(this.sFormat);
			if (format.bScientific && /^0\.0*,*E\+00$/.test(normalized)) {
				nType = c_oAscNumFormatType.Scientific;
			} else {
				var types = [c_oAscNumFormatType.Text, c_oAscNumFormatType.Percent,
				c_oAscNumFormatType.Number, c_oAscNumFormatType.Fraction, c_oAscNumFormatType.Currency,
				c_oAscNumFormatType.Accounting
				];
				for (var i = 0; i < types.length; ++i) {
					var type = types[i];
					info.asc_setType(type);
					var formats = getFormatCells(info);
					for (var j = 0; j < formats.length; j++) {
						if (stripFormatEscaping(formats[j]) === normalized) {
							nType = type;
							break;
						}
					}
					if (nType !== c_oAscNumFormatType.Custom) {
						break;
					}
				}
			}
		}
		return nType;
	},
	checkCultureInfoFontPicker: function() {
		if (null !== this.sFormat) {
			AscFonts.FontPickerByCharacter.getFontsByString(this.sFormat);
		}
		if (null !== this.oPositiveFormat && null !== this.oPositiveFormat.LCID) {
			checkCultureInfoFontPicker(this.oPositiveFormat.LCID);
		}
		if (null !== this.oNegativeFormat && null !== this.oNegativeFormat.LCID) {
			checkCultureInfoFontPicker(this.oNegativeFormat.LCID);
		}
		if (null !== this.oNullFormat && null !== this.oNullFormat.LCID) {
			checkCultureInfoFontPicker(this.oNullFormat.LCID);
		}
		if (null !== this.oTextFormat && null !== this.oTextFormat.LCID) {
			checkCultureInfoFontPicker(this.oTextFormat.LCID);
		}
		if (this.aComporationFormats) {
			for (var i = 0, length = this.aComporationFormats.length; i < length; ++i) {
				var oCurFormat = this.aComporationFormats[i];
				if (null !== oCurFormat.LCID) {
					checkCultureInfoFontPicker(oCurFormat.LCID);
				}

			}
		}
	}
};
var oDecodeGeneralFormatCache = {};
function DecodeGeneralFormat(val, nValType, dDigitsCount)
{
    var cacheVal = oDecodeGeneralFormatCache[val];
    if(null != cacheVal)
    {
        cacheVal = cacheVal[nValType];
        if(null != cacheVal)
        {
            cacheVal = cacheVal[dDigitsCount];
            if(null != cacheVal)
                return cacheVal;
        }
    }
    var res = DecodeGeneralFormat_Raw(val, nValType, dDigitsCount);
    var cacheVal = oDecodeGeneralFormatCache[val];
    if(null == cacheVal)
    {
        cacheVal = {};
        oDecodeGeneralFormatCache[val] = cacheVal;
    }
    var cacheType = cacheVal[nValType];
    if(null == cacheType)
    {
        cacheType = {};
        cacheVal[nValType] = cacheType;
    }
    cacheType[dDigitsCount] = res;
    return res;
}
function DecodeGeneralFormat_Raw(val, nValType, dDigitsCount)
{
    if(CellValueType.String == nValType)
        return "@";
    var number = val - 0;
    if(number != val)
        return "@";
    if(0 == number)
        return "0";
    var nDigitsCount;
    if(null == dDigitsCount || dDigitsCount > gc_nMaxDigCountView)
        nDigitsCount = gc_nMaxDigCountView;
    else
        nDigitsCount = parseInt(dDigitsCount);//while measurer is not connected, we don't use non-integer metrics
    if(number < 0)
    {
        //todo maybe need nDigitsCount--
        //nDigitsCount--;//for '-' sign
        number = -number;
    }
    if(nDigitsCount < 1)
        return "0";//can return any numeric format, it won't be considered anyway when nDigitsCount < 1
	var bContinue = true;
	var parts = getNumberParts(number);
	while(bContinue)
	{
		bContinue = false;
		var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0 for 0.123
		var nRealExpAbs = Math.abs(nRealExp);
		var nExpMinDigitsCount;//number of digits in 'E+00' format
		if(nRealExpAbs < 100)
			nExpMinDigitsCount = 4;
		else
			nExpMinDigitsCount = 2 + nRealExpAbs.toString().length;
		
		var suffix = "";
		if (nRealExp > 0)
		{
			if(nRealExp > nDigitsCount)
			{
				if(nDigitsCount >= nExpMinDigitsCount + 1)//1 for one more character before E (*E+00)
				{
					suffix = "E+";
					for(var i = 2; i < nExpMinDigitsCount; ++i)
						suffix += "0";
					nDigitsCount -= nExpMinDigitsCount;
				}
				else
					return "0";//can return any numeric format, there will be hashes anyway
			}
		}
		else
		{
			var nVarian1 = nDigitsCount - 2 + nRealExp;//without E+00, 2 for "0." characters
			var nVarian2 = nDigitsCount - nExpMinDigitsCount;// with E+00
			if(nVarian2 > 2)
				nVarian2--;//for '.' character
			else if(nVarian2 > 0)
				nVarian2 = 1;
			if(nVarian1 <= 0 && nVarian2 <= 0)
				return "0";
			if(nVarian1 < nVarian2)
			{
				//if the number fits completely in nVarian1, then use nVarian1
				var bUseVarian1 = false;
				if(nVarian1 > 0 && 0 == (parts.mantissa % Math.pow(10, gc_nMaxDigCount - nVarian1)))
					bUseVarian1 = true;
				if(false == bUseVarian1)
				{
					if(nDigitsCount >= nExpMinDigitsCount + 1)
					{
						suffix = "E+";
						for(var i = 2; i < nExpMinDigitsCount; ++i)
							suffix += "0";
						nDigitsCount -= nExpMinDigitsCount;
					}
					else
						return "0";//can return any numeric format, there will be hashes anyway
				}
			}
		}
		var dec_num_digits = nRealExp;
		if(suffix)
			dec_num_digits = 1;
		//round the mantissa to correctly handle the situation 0.999 when nDigitsCount = 4
		var nRoundDigCount = 0;
		if(dec_num_digits <= 0)
		{
			//2 for '0.' characters
			var nTemp = nDigitsCount + dec_num_digits - 2;
			if(nTemp > 0)
				nRoundDigCount = nTemp;
		}
		else if(dec_num_digits < gc_nMaxDigCount)
		{
			if(dec_num_digits <= nDigitsCount)
			{
				//1 for '.' character
				if(dec_num_digits + 1 < nDigitsCount)
					nRoundDigCount = nDigitsCount - 1;
				else
					nRoundDigCount = dec_num_digits;
			}
		}
		if(nRoundDigCount > 0)
		{
			var nTemp = Math.pow(10, gc_nMaxDigCount - nRoundDigCount);
			number = Math.round(parts.mantissa / nTemp) * nTemp * Math.pow(10, parts.exponent);
			
			var oNewParts = getNumberParts(number);
			//if the number of digits changed as a result of rounding, need to start over
			if(oNewParts.exponent != parts.exponent)
				bContinue = true;
			else
				bContinue = false;
			parts = oNewParts;
		}
	}
	
    var frac_num_digits;
    if(dec_num_digits > 0)
        frac_num_digits = nDigitsCount - 1 - dec_num_digits;//1 for '.' character
    else
        frac_num_digits = nDigitsCount - 2 + dec_num_digits;//2 for '0.' characters

    //calculate the required number of digits after decimal point
    if(frac_num_digits > 0)
    {
		var sTempNumber = parts.mantissa.toString();
		if(dec_num_digits > 0)
			sTempNumber = sTempNumber.substring(dec_num_digits, dec_num_digits + frac_num_digits);
		else
			sTempNumber = sTempNumber.substring(0, frac_num_digits);
        var nTempNumberLength = sTempNumber.length;
        var nreal_frac_num_digits = frac_num_digits;
        for(var i = frac_num_digits - 1; i >= 0; --i)
        {
            if("0" == sTempNumber[i])
                nreal_frac_num_digits--;
            else
                break;
        }
        frac_num_digits = nreal_frac_num_digits;
		if(dec_num_digits < 0)
			frac_num_digits += (-dec_num_digits);
    }
    if(frac_num_digits <= 0)
        return "0" + suffix;

    //build the format
    var number_format_string = "0" + gc_sFormatDecimalPoint;
    for(var i = 0; i < frac_num_digits; ++i)
        number_format_string += "0";
    number_format_string += suffix;
    return number_format_string;
}
function GeneralEditFormatCache()
{
    this.oCache = {};
}
GeneralEditFormatCache.prototype =
{
	cleanCache : function(){
		this.oCache = {};
	},
    format: function (number, cultureInfo)
    {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        //convert the number so that the string contains only 15 significant digits.
        var value = this.oCache[number];
        if(null == value)
        {
			if(0 == number)
				value = "0";
			else
			{
				var sRes = "";
				var parts = getNumberParts(number);
				var nRealExp = gc_nMaxDigCount + parts.exponent;//nRealExp == 0 for 0.123
				if(parts.exponent >= 0)//nRealExp >= -gc_nMaxDigCount
				{
					if(nRealExp <= 21)
					{
						sRes = parts.mantissa.toString();
						for(var i = 0; i < parts.exponent; ++i)
							sRes += "0";
					}
					else
					{
					    sRes = this._removeTileZeros(parts.mantissa.toString(), cultureInfo);
						if(sRes.length > 1)
						{
							var temp = sRes.substring(0, 1);
							temp += cultureInfo.NumberDecimalSeparator;
							temp += sRes.substring(1);
							sRes = temp;
						}
						sRes += "E+" + (nRealExp - 1);
					}
				}
				else
				{
					if(nRealExp > 0)
					{
						sRes = parts.mantissa.toString();
						if(sRes.length > nRealExp)
						{
							var temp = sRes.substring(0, nRealExp);
							temp += cultureInfo.NumberDecimalSeparator;
							temp += sRes.substring(nRealExp);
							sRes = temp;
						}
						sRes = this._removeTileZeros(sRes, cultureInfo);
					}
					else
					{
						if(nRealExp >= -18)
						{
							sRes = "0";
							sRes += cultureInfo.NumberDecimalSeparator;
							for(var i = 0; i < -nRealExp; ++i)
								sRes += "0";
							var sTemp = parts.mantissa.toString();
							sTemp = sTemp.substring(0, 19 + nRealExp);
							sRes += this._removeTileZeros(sTemp, cultureInfo);
						}
						else
						{
							sRes = parts.mantissa.toString();
							if(sRes.length > 1)
							{
								var temp = sRes.substring(0, 1);
								temp += cultureInfo.NumberDecimalSeparator;
								temp += sRes.substring(1);
								temp = this._removeTileZeros(temp, cultureInfo);
								sRes = temp;
							}
							sRes += "E-" + (1 - nRealExp);
						}
					}
				}
				if( SignType.Negative == parts.sign)
					value = "-" + sRes;
				else
					value = sRes;
			}
            this.oCache[number] = value;
        }
        return value;
    },
    _removeTileZeros: function (val, cultureInfo)
    {
		var res = val;
		var nLength = val.length;
		var nLastNoZero = nLength - 1;
		for(var i = val.length - 1; i >= 0; --i)
		{
			nLastNoZero = i;
			if("0" != val[i])
				break;
		}
		if(nLastNoZero != nLength - 1)
		{
		    if (cultureInfo.NumberDecimalSeparator == res[nLastNoZero])
				res = res.substring(0, nLastNoZero);
			else
				res = res.substring(0, nLastNoZero + 1);
		}
		return res;
	}
};
var oGeneralEditFormatCache = new GeneralEditFormatCache();

function FormatParser()
{
	this.days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
	this.daysLeap = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
}
FormatParser.prototype =
{
    // Replaces the locale decimal separator with '.' (JS format).
    // '.' is first replaced with 'q' so it is not recognized as a decimal separator, matching Excel behavior.
    _normalizeDecimalSep: function (val, cultureInfo) {
        if (typeof val !== "string")
            val = String(val);
        if ("." != cultureInfo.NumberDecimalSeparator) {
            val = val.replace(".", "q");
            val = val.replace(cultureInfo.NumberDecimalSeparator, ".");
        }
        return val;
    },
    isLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        val = this._normalizeDecimalSep(val, cultureInfo);
        //parseNum excludes hex number notation.
        return AscCommonExcel.parseNum(val) && Asc.isNumberInfinity(val);
    },
    parseLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        return this._normalizeDecimalSep(val, cultureInfo) - 0;
    },
    // Reverse of parseLocaleNumber: converts a JS-format number string ("0.2") to locale display string ("0,2").
    toLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        if (typeof val !== "string")
            val = String(val);
        var sep = cultureInfo.NumberDecimalSeparator;
        return sep !== "." ? val.replace(".", sep) : val;
    },
    // Combines isLocaleNumber + parseLocaleNumber in a single pass.
    // Returns the parsed number, or null if val is not a valid locale number.
    tryParseLocaleNumber: function (val, cultureInfo) {
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        val = this._normalizeDecimalSep(val, cultureInfo);
        if (!AscCommonExcel.parseNum(val) || !Asc.isNumberInfinity(val))
            return null;
        return val - 0;
    },
    /**
     * Get cached regex for number parsing
     * @private
     */
    _getNumberRegex: function (cultureInfo) {
        // Cache key based on culture-specific separators
        var cacheKey = cultureInfo.NumberGroupSeparator + "|" + 
                       cultureInfo.NumberDecimalSeparator + "|" + 
                       cultureInfo.CurrencySymbol;
        
        if (!this._numberRegexCache) {
            this._numberRegexCache = {};
        }
        let regex = this._numberRegexCache[cacheKey];
        if (!regex) {
            regex = new RegExp(
                "^(([ \\+\\-%\\$€£¥\\(]|" + escapeRegExp(cultureInfo.CurrencySymbol) + 
                ")*)((?:\\d+(?:" + escapeRegExp(cultureInfo.NumberGroupSeparator) + 
                "\\d+)*(?:" + escapeRegExp(cultureInfo.NumberDecimalSeparator) + 
                "\\d*)?)?(?:\\s*\\d+/\\d+)?)(([ %\\)]|р.|" + 
                escapeRegExp(cultureInfo.CurrencySymbol) + ")*)$"
            );
            this._numberRegexCache[cacheKey] = regex;
        }
        return regex;
    },
    /**
     * Check if format should be preserved (not auto-detected from input)
     * @private
     */
    _shouldPreserveFormat: function (currentFormat, stringFormat) {
        switch (currentFormat) {
            case Asc.c_oAscNumFormatType.Number:
            case Asc.c_oAscNumFormatType.Currency:
            case Asc.c_oAscNumFormatType.Accounting:
            case Asc.c_oAscNumFormatType.Date:
            case Asc.c_oAscNumFormatType.Time:
            case Asc.c_oAscNumFormatType.LongDate:
                return true;
            case Asc.c_oAscNumFormatType.Fraction:
                // Preserve unless it's a "detect" format like "# ?/?"
                return -1 === stringFormat.indexOf('/?');
            case Asc.c_oAscNumFormatType.Percent:
                return stringFormat !== interfaceFormatPercent;
            case Asc.c_oAscNumFormatType.Scientific:
                return stringFormat !== interfaceFormatScientific;
            default:
                return false;
        }
    },
    /**
     * Check if format is numeric (for simple fraction "1/2" disambiguation with dates)
     * @private
     */
    _isNumericFormat: function (currentFormat) {
        switch (currentFormat) {
			case Asc.c_oAscNumFormatType.Number:
			case Asc.c_oAscNumFormatType.Currency:
			case Asc.c_oAscNumFormatType.Accounting:
			case Asc.c_oAscNumFormatType.Fraction:
			case Asc.c_oAscNumFormatType.Percent:
			case Asc.c_oAscNumFormatType.Scientific:
                return true;
            default:
                return false;
        }
    },
    /**
     * Build currency format string based on pattern
     * @private
     */
    _buildCurrencyFormat: function (sCurrency, sFracFormat, nPattern) {
        var sNumberFormat = "#" + gc_sFormatThousandSeparator + "##0" + sFracFormat;
        var sCurrencyFormat = (sCurrency.length > 1) ? "\"" + sCurrency + "\"" : "\\" + sCurrency;
        var sPositivePattern, sNegativePattern;

        switch (nPattern) {
            case 0:
                sPositivePattern = sCurrencyFormat + sNumberFormat + "_)";
                sNegativePattern = "[Red](" + sCurrencyFormat + sNumberFormat + ")";
                break;
            case 1:
                sPositivePattern = sCurrencyFormat + sNumberFormat;
                sNegativePattern = "[Red]-" + sCurrencyFormat + sNumberFormat;
                break;
            case 2:
                sPositivePattern = sCurrencyFormat + sNumberFormat;
                sNegativePattern = "[Red]" + sCurrencyFormat + "-" + sNumberFormat;
                break;
            case 3:
                sPositivePattern = sCurrencyFormat + sNumberFormat + "_-";
                sNegativePattern = "[Red]" + sCurrencyFormat + sNumberFormat + "-";
                break;
            case 4:
                sPositivePattern = sNumberFormat + sCurrencyFormat + "_)";
                sNegativePattern = "[Red](" + sNumberFormat + sCurrencyFormat + ")";
                break;
            case 5:
                sPositivePattern = sNumberFormat + sCurrencyFormat;
                sNegativePattern = "[Red]-" + sNumberFormat + sCurrencyFormat;
                break;
            case 6:
                sPositivePattern = sNumberFormat + "-" + sCurrencyFormat;
                sNegativePattern = "[Red]" + sNumberFormat + "-" + sCurrencyFormat;
                break;
            case 7:
                sPositivePattern = sNumberFormat + sCurrencyFormat + "_-";
                sNegativePattern = "[Red]" + sNumberFormat + sCurrencyFormat + "-";
                break;
            case 8:
                sPositivePattern = sNumberFormat + " " + sCurrencyFormat;
                sNegativePattern = "[Red]-" + sNumberFormat + " " + sCurrencyFormat;
                break;
            case 9:
                sPositivePattern = sCurrencyFormat + " " + sNumberFormat;
                sNegativePattern = "[Red]-" + sCurrencyFormat + " " + sNumberFormat;
                break;
            case 10:
                sPositivePattern = sNumberFormat + " " + sCurrencyFormat + "_-";
                sNegativePattern = "[Red]" + sNumberFormat + " " + sCurrencyFormat + "-";
                break;
            case 11:
                sPositivePattern = sCurrencyFormat + " " + sNumberFormat + "_-";
                sNegativePattern = "[Red]" + sCurrencyFormat + " " + sNumberFormat + "-";
                break;
            case 12:
                sPositivePattern = sCurrencyFormat + " " + sNumberFormat;
                sNegativePattern = "[Red]" + sCurrencyFormat + " -" + sNumberFormat;
                break;
            case 13:
                sPositivePattern = sNumberFormat + " " + sCurrencyFormat;
                sNegativePattern = "[Red]" + sNumberFormat + "- " + sCurrencyFormat;
                break;
            case 14:
                sPositivePattern = sCurrencyFormat + " " + sNumberFormat + "_)";
                sNegativePattern = "[Red](" + sCurrencyFormat + " " + sNumberFormat + ")";
                break;
            case 15:
                sPositivePattern = sNumberFormat + " " + sCurrencyFormat + "_)";
                sNegativePattern = "[Red](" + sNumberFormat + " " + sCurrencyFormat + ")";
                break;
            default:
                sPositivePattern = sCurrencyFormat + sNumberFormat;
                sNegativePattern = "[Red]-" + sCurrencyFormat + sNumberFormat;
        }
        return sPositivePattern + ";" + sNegativePattern;
    },
    parse: function (value, cultureInfo, currentFormat, stringFormat)
    {
        if (currentFormat === Asc.c_oAscNumFormatType.Text)
            return null;
        if (null == cultureInfo)
            cultureInfo = g_oDefaultCultureInfo;
        if (!stringFormat)
            stringFormat = AscCommon.g_cGeneralFormat;

        // Replace Non-breaking space (0xA0) with White-space (0x20)
        if (" " == cultureInfo.NumberGroupSeparator)
            value = value.replace(new RegExp(String.fromCharCode(0xA0), "g"), "");

        var shouldPreserveFormat = this._shouldPreserveFormat(currentFormat, stringFormat);
		const isDateOverrideFormat = stringFormat in gc_oParseDateOverrideFormats || stringFormat === interfaceShortDateFormat;

        // Regex that matches numbers, mixed fractions ("1 1/2"), and simple fractions ("1/2")
        // Cached per cultureInfo to avoid recompilation on every call
        var rx_thouthand = this._getNumberRegex(cultureInfo);
        var match = value.match(rx_thouthand);

        // Variables for fraction parsing
        var sVal = null;
        var sNumerator = null;
        var sDenominator = null;
        var sBefore = null;
        var sAfter = null;
        var res = null;
        var bError = false;

        if (null != match) {
            // If the third group has "/" symbol parse it like a fraction
            if (match[3] && match[3].indexOf('/') !== -1) {
                // Find fraction at end: must be at start OR after space
                // This rejects "0.5/2" (no space before 5) but accepts "1 1/2", "1,234 1/2"
                var fractionMatch = match[3].match(/(?:^|\s)(\d+)\/(\d+)$/);
                if (!fractionMatch) {
                    match = null;
                } else {
                    sNumerator = fractionMatch[1];
                    sDenominator = fractionMatch[2];
                    
                    // Get part before fraction - let existing locale-aware parsing handle it
                    var beforeFraction = match[3].substring(0, fractionMatch.index).trim();
                    
                    if (beforeFraction === '') {
                        // Simple fraction like "1/2"
                        if (!this._isNumericFormat(currentFormat)) {
                            // Let parseDate handle it (could be date like "1/2" = Jan 2)
                            match = null;
                        } else {
                            sVal = '0';
                        }
                    } else {
                        // Mixed fraction - whole part validated by _parseThouthand (locale-aware)
                        sVal = beforeFraction;
                    }
                }
            } else {
                sVal = match[3];
            }
        }

        if (null != match) {
            sBefore = match[1];
            sAfter = match[4];

            var oChartCount = {};
            if (null != sBefore)
                this._parseStringLetters(sBefore, cultureInfo.CurrencySymbol, true, oChartCount);
            if (null != sAfter)
                this._parseStringLetters(sAfter, cultureInfo.CurrencySymbol, false, oChartCount);
            if (sNumerator && sDenominator)
                this._parseStringLetters('/', cultureInfo.CurrencySymbol, false, oChartCount);
            var bMinus = false;
            var bPercent = false;
            var bFraction = false;
            var sCurrency = null;
			var oCurrencyElem = null;
			var nBracket = 0;
			for(var sChar in oChartCount){
				var elem = oChartCount[sChar];
				if(" " == sChar)
					continue;
				else if("+" == sChar){
					if(elem.all > 1)
						bError = true;
				}
				else if("-" == sChar){
					if(elem.all > 1)
						bError = true;
					else
						bMinus = true;
				}
				else if("(" == sChar){
					if(1 == elem.all && 1 == elem.before)
						nBracket++;
					else
						bError = true;
				}
				else if(")" == sChar){
					if(1 == elem.all && 1 == elem.after)
						nBracket++;
					else
						bError = true;
				}
				else if("%" == sChar){
					if(1 == elem.all)
						bPercent = true;
					else
						bError = true;
				}
                else if ('/' == sChar) {
                    if (sVal) {
                        if (1 == elem.all)
                            bFraction = true;
                        else
                            bError = true;
                    } else {
                        // "/" without value - treat as error
                        bError = true;
                    }
                }
				else{
					if(null == sCurrency && 1 == elem.all){
						sCurrency = sChar;
						oCurrencyElem = elem;
					}
					else
						bError = true;
				}
			}
			if (nBracket > 0) {
			    if (2 == nBracket)
			        bMinus = true;
			    else
			        bError = true;
			}
			var CurrencyNegativePattern = cultureInfo.CurrencyNegativePattern;
			if(null != sCurrency){
			    if (sCurrency == cultureInfo.CurrencySymbol) {
			        var nPattern = cultureInfo.CurrencyNegativePattern;
			        if (0 == nPattern || 1 == nPattern || 2 == nPattern || 3 == nPattern || 9 == nPattern || 11 == nPattern || 12 == nPattern || 14 == nPattern) {
			            if (1 != oCurrencyElem.before)
			                bError = true;
			        }
			        else if (1 != oCurrencyElem.after)
			            bError = true;
			    }
			    else if(-1 != "$€£¥".indexOf(sCurrency)){
			        if (1 == oCurrencyElem.before) {
			            CurrencyNegativePattern = 0;
			        }
                    else
						bError = true;
				}
				else if(-1 != "р.".indexOf(sCurrency)){
				    if (1 == oCurrencyElem.after) {
				        CurrencyNegativePattern = 5;
				    }
                    else
						bError = true;
				}
				else
				    bError = true;
			}
			if(!bError){
				var oVal = this._parseThouthand(sVal, sNumerator, sDenominator, cultureInfo);

                if (oVal) {
					res = {format: null, value: null, bDateTime: false, bDate: false, bTime: false, bPercent: false, bCurrency: false};
					var dVal = oVal.number;
					if (bMinus)
						dVal = -dVal;
					var sFracFormat = "";
					if (parseInt(dVal) != dVal)
						sFracFormat = gc_sFormatDecimalPoint + "00";
					var sFormat = null;
					
					// Percent: always divide by 100, but only change format for non-preserving types
					if (bPercent) {
						res.bPercent = true;
						dVal /= 100;
						if (shouldPreserveFormat && !isDateOverrideFormat) {
							sFormat = stringFormat;
						} else {
							sFormat = "0" + sFracFormat + "%";
						}
					}
                    else if (bFraction) 
                    {
                        res.bFraction = true;
                        sFormat = this._selectFractionFormat(sNumerator, sDenominator, shouldPreserveFormat, isDateOverrideFormat, currentFormat, stringFormat);
                    }
					else if (sCurrency && !(shouldPreserveFormat && !isDateOverrideFormat)) {
						res.bCurrency = true;
						sFormat = this._buildCurrencyFormat(sCurrency, sFracFormat, CurrencyNegativePattern);
					}
					else if (oVal.thouthand && (currentFormat === undefined || currentFormat == Asc.c_oAscNumFormatType.General)) {
						// Only apply thousand-separator format for General format type (or when no format specified)
						sFormat = "#" + gc_sFormatThousandSeparator + "##0" + sFracFormat;
					}
					else
						sFormat = stringFormat;
					res.format = sFormat;
					res.value = dVal;
                    if (!sFormat) 
                        res = null;
				}
			}
        }
        // Handle special cases after main parsing
        if (res == null && value[0] == ' ') {
            return null;
        }
        if (res == null && !bError) {
            res = this.parseDate(value, cultureInfo, shouldPreserveFormat, isDateOverrideFormat, currentFormat, stringFormat);
        }

        return res;
    },
    _parseStringLetters: function (sVal, currencySymbol, bBefore, oRes) {
        //separately handle 'р.' and currencySymbol because they may not be single characters
        var aTemp = ["р.", currencySymbol];
        for (var i = 0, length = aTemp.length; i < length; i++){
            var sChar = aTemp[i];
            var nIndex = -1;
            var nCount = 0;
            while(-1 != (nIndex = sVal.indexOf(sChar, nIndex + 1)))
                nCount++;
            if(nCount > 0)
            {
                sVal = sVal.replace(new RegExp(escapeRegExp(sChar), "g"), "");
                var elem = oRes[sChar];
                if(!elem){
                    elem = {before: 0, after: 0, all: 0};
                    oRes[sChar] = elem;
                }
                if(bBefore)
                    elem.before += nCount;
                else
                    elem.after += nCount;
                elem.all += nCount;
            }
        }
        for (var i = 0, length = sVal.length; i < length; i++) {
            var sChar = sVal[i];
            var elem = oRes[sChar];
            if (!elem) {
                elem = {before: 0, after: 0, all: 0};
                oRes[sChar] = elem;
            }
            if (bBefore)
                elem.before++;
            else
                elem.after++;
            elem.all++;
        }
    },
    /**
     * Select appropriate fraction format based on numerator/denominator lengths and current format
     * @param {string} sNumerator - numerator string
     * @param {string} sDenominator - denominator string
	 * @param {boolean} shouldPreserveFormat - true if format should be preserved
	 * @param {boolean} isDateOverrideFormat - true if format is a date override format
     * @param {number} currentFormat - current cell format type
	 * @param {string} stringFormat - current format string
     * @returns {string|null} - fraction format string or null
     */
    _selectFractionFormat: function (sNumerator, sDenominator, shouldPreserveFormat, isDateOverrideFormat, currentFormat, stringFormat) {
        var numLength = sNumerator.length;
        var denomLength = sDenominator.length;
        var maxLength = Math.max(numLength, denomLength);
        if ((shouldPreserveFormat && !isDateOverrideFormat) || currentFormat == Asc.c_oAscNumFormatType.Fraction) {
            return stringFormat;
        }
        
        // Limit to 3 digits max - larger fractions return null (will be treated as text)
        if (numLength > 3 || denomLength > 3) {
            return null;
        }
        
        // Simple rule: single digit denominators use "# ?/?", otherwise "# ??/??"
        // Exception: when current format is already fraction-like, prefer simpler format
        if (maxLength <= 1) {
            return "# ?/?";
        } else if (denomLength <= 1) {
            // Denominator is single digit - use simpler format
            return "# ?/?";
        } else {
            // Multi-digit denominator - use double-digit format
            return "# ??/??";
        }
    },
    _parseThouthand: function (val, sNumerator, sDenominator, cultureInfo)
    {
        var oRes = null;
        var bThouthand = false;
        // Reverse the string to scan group separators from the least-significant end.
        const sReverseVal = val.split("").reverse().join("");
        var nGroupSizeIndex = 0;
        var nGroupSize = cultureInfo.NumberGroupSizes[nGroupSizeIndex];
        var nPrevIndex = 0;
        var nIndex = -1;
        var bError = false;
        while (-1 != (nIndex = sReverseVal.indexOf(cultureInfo.NumberGroupSeparator, nIndex + 1))) {
            var nCurLength = nIndex - nPrevIndex;
            if (nCurLength < nGroupSize) {
                bError = true;
                break;
            }
            if (nGroupSizeIndex < cultureInfo.NumberGroupSizes.length - 1) {
                nGroupSizeIndex++;
                nGroupSize = cultureInfo.NumberGroupSizes[nGroupSizeIndex];
            }
            nPrevIndex = nIndex + 1;
        }
        if (!bError) {
            if (0 != nPrevIndex) {
                //so that 0,001 is not recognized
                if (nPrevIndex < val.length && parseInt(val.substr(0, val.length - nPrevIndex)) > 0) {
                    val = val.replace(new RegExp(escapeRegExp(cultureInfo.NumberGroupSeparator), "g"), '');
                    bThouthand = true;
                }
            }
			if (g_oFormatParser.isLocaleNumber(val, cultureInfo)) {
				var dNumber = g_oFormatParser.parseLocaleNumber(val, cultureInfo);
                if(sNumerator && sDenominator) {
                    // Mixed fractions must have integer whole part (e.g., "0.5 1/2" is invalid)
                    if (Number.isInteger(dNumber)) {
                        oRes = { number: dNumber + (sNumerator / sDenominator), thouthand: bThouthand };
                    }
                    // else oRes stays null - invalid mixed fraction
                }
				else {
                    oRes = { number: dNumber, thouthand: bThouthand };
                }
			}
        }
		return oRes;
	},
    _parseDateFromArray: function (match, oDataTypes, cultureInfo)
	{
        var res = null;
        var bError = false;
        //in the first pass, separate date and time using delimiter
        for (var i = 0, length = match.length; i < length; i++) {
            var elem = match[i];
            if (elem.type == oDataTypes.delimiter) {
                bError = true;
                if(i - 1 >= 0 && i + 1 < length){
                    var prev = match[i - 1];
                    var next = match[i + 1];
                    if(prev.type != oDataTypes.delimiter && next.type != oDataTypes.delimite){
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if(false == prev.date && false == next.date){
                                bError = false;
                                prev.time = true;
                                next.time = true;
                            }
                        }
                        else{
                            if(false == prev.time && false == next.time){
                                bError = false;
                                prev.date = true;
                                next.date = true;
                            }
                        }
                    }
                }
                else if (i - 1 >= 0 && i + 1 == length) {
                    //case "10:"
                    var prev = match[i - 1];
                    if (prev.type != oDataTypes.delimiter) {
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if (false == prev.date) {
                                bError = false;
                                prev.time = true;
                            }
                        }
                    }
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            //separate date and time using Am/Pm and month names
            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.type == oDataTypes.letter){
                    var valLower = elem.val.toLowerCase();
                    if (elem.am || elem.pm) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.date) {
                                prev.time = true;
                            }
                        }
                        //AmPm should be the last entry
                        if (i + 1 != length) {
                            bError = true;
                        }
                    }
                    else if (null != elem.month) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.time)
                                prev.date = true;
                        }
                        if (i + 1 < length) {
                            let next = match[i + 1];
                            // processing the option when the date is given as the format "October 11, 2008"
                            if (i === 0 && i + 2 < length) {
                                let afterNext = match[i + 2];
                                if (oDataTypes.digit == afterNext.type && false == afterNext.time) {
                                    afterNext.date = true;
                                }
                            }
                            if (oDataTypes.digit == next.type && false == next.time)
                                next.date = true;
                        }
                    }
                    else
                        bError = true;
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            var aDate = [];
            var nMonthIndex = null;
			var sMonthFormat = null;
            var aTime = [];
            var am = false;
            var pm = false;

            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.date) {
                    if (elem.type == oDataTypes.digit)
                        aDate.push(elem.val);
                    else if (elem.type == oDataTypes.letter && null != elem.month) {
                        nMonthIndex = aDate.length;
                        sMonthFormat = elem.month.format;
                        aDate.push(elem.month.val);
                    }
                    else
                        bError = true;
                }
                else if (elem.time) {
                    if (elem.type == oDataTypes.digit)
                        aTime.push(elem.val);
                    else if (elem.type == oDataTypes.letter && (elem.am || elem.pm)) {
                        am = elem.am;
                        pm = elem.pm;
                    }
                    else
                        bError = true;
                }
                else if (oDataTypes.digit == elem.type)
                    bError = true;//case "1-2-3 10"
            }
            var nDateLength = aDate.length;
            if (nDateLength > 0 && !(2 <= nDateLength && nDateLength <= 3 && (null == nMonthIndex || (3 == nDateLength && 1 == nMonthIndex) || 2 == nDateLength || (3 == nDateLength && 0 == nMonthIndex))))
                bError = true;
            var nTimeLength = aTime.length;
            if (nTimeLength > 3)
                bError = true;
            if(!bError){
                res = { d: null, m: null, y: null, h: null, min: null, s: null, am: am, pm: pm, sDateFormat: null };
                if (nDateLength > 0) {
                    var nIndexD = Math.max(cultureInfo.ShortDatePattern.indexOf("0"), cultureInfo.ShortDatePattern.indexOf("1"));
                    var nIndexM = Math.max(cultureInfo.ShortDatePattern.indexOf("2"), cultureInfo.ShortDatePattern.indexOf("3"));
                    var nIndexY = Math.max(cultureInfo.ShortDatePattern.indexOf("4"), cultureInfo.ShortDatePattern.indexOf("5"));
                    if (null != nMonthIndex) {
                        if (2 == nDateLength) {
                            res.d = aDate[nDateLength - 1 - nMonthIndex];
                            res.m = aDate[nMonthIndex];
                            //priority goes to d-mmm format, but if it doesn't fit we try mmm-yy
                            if (this.isValidDate((new Date()).getFullYear(), res.m - 1, res.d))
                                res.sDateFormat = "d-mmm";
                            else {
                                //in non-classic case (!= dd/mm/yyyy) swap d and m before trying y
                                if (!isDMY(cultureInfo) && this.isValidDate((new Date()).getFullYear(), res.d - 1, res.m)) {
                                    res.sDateFormat = "d-mmm";
                                    var temp = res.d;
                                    res.d = res.m;
                                    res.m = temp;
                                }
                                else {
                                    //if text month is second, then the first parameter can only be a day
                                    if (0 == nMonthIndex) {
                                        res.sDateFormat = "mmm-yy";
                                        res.d = null;
                                        res.m = aDate[0];
                                        res.y = aDate[1];
                                    }
                                    else
                                        bError = true;
                                }
                            }
                        } else {
                            if (nMonthIndex == 0) {
                                res.sDateFormat = "dd-mmm-yy";
                                res.m = aDate[0];
                                res.d = aDate[1];
                                res.y = aDate[2];
                            } else {
                                res.sDateFormat = "d-mmm-yy";
                                res.d = aDate[0];
                                res.m = aDate[1];
                                res.y = aDate[2];
                            }
                        }
                    }
                    else {
                        //check the order in default format
                        if (2 == nDateLength) {
                            //d and m have priority
                            if (nIndexD < nIndexM) {
                                res.d = aDate[0];
                                res.m = aDate[1];
                            }
                            else {
                                res.m = aDate[0];
                                res.d = aDate[1];
                            }
                            if (this.isValidDate((new Date()).getFullYear(), res.m - 1, res.d))
                                res.sDateFormat = "d-mmm";
                            else{
                                //in reverse notation (== yyyy/mm/dd) swap d and m before trying y
                                if (isYMD(cultureInfo) && this.isValidDate((new Date()).getFullYear(), res.d - 1, res.m)) {
                                    res.sDateFormat = "d-mmm";
                                    var temp = res.d;
                                    res.d = res.m;
                                    res.m = temp;
                                }
                                else{
                                    res.sDateFormat = "mmm-yy";
                                    res.d = null;
                                    if (nIndexM < nIndexY) {
                                        res.m = aDate[0];
                                        res.y = aDate[1];
                                    }
                                    else {
                                        res.y = aDate[0];
                                        res.m = aDate[1];
                                    }
                                }
                            }
                        } else if(3 == nDateLength && aDate[0] > 1000) {
                            res.y = aDate[0];
                            res.m = aDate[1];
                            res.d = aDate[2];
                            res.sDateFormat = getShortDateFormat(cultureInfo);
                        } else {
                            for (var i = 0, length = cultureInfo.ShortDatePattern.length; i < length; i++)
                            {
                                var nIndex = cultureInfo.ShortDatePattern[i] - 0;
                                var val = aDate[i];
                                if (0 == nIndex || 1 == nIndex) {
                                    res.d = val;
                                } else if (2 == nIndex || 3 == nIndex) {
                                    res.m = val;
                                } else if (4 == nIndex || 5 == nIndex) {
                                    res.y = val;
                                }
                            }
                            res.sDateFormat = getShortDateFormat(cultureInfo);
                        }
                    }
                    if(null != res.y)
                    {
                        if(res.y < 30)
                            res.y = 2000 + res.y;
                        else if(res.y < 100)
                            res.y = 1900 + res.y;
                    }
                }
                if(nTimeLength > 0){
                    res.h = aTime[0];
                    if(nTimeLength > 1)
                        res.min = aTime[1];
                    if(nTimeLength > 2)
                        res.s = aTime[2];
                }
                if(bError)
                    res = null;
            }
        }
		return res;
    },
	_parseDateFromArrayPDF: function (match, oDataTypes, cultureInfo, oFormat)
	{
        var res = null;
        var bError = false;
        //in the first pass, separate date and time using delimiter
        for (var i = 0, length = match.length; i < length; i++) {
            var elem = match[i];
            if (elem.type == oDataTypes.delimiter) {
                bError = true;
                if(i - 1 >= 0 && i + 1 < length){
                    var prev = match[i - 1];
                    var next = match[i + 1];
                    if(prev.type != oDataTypes.delimiter && next.type != oDataTypes.delimite){
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if(false == prev.date && false == next.date){
                                bError = false;
                                prev.time = true;
                                next.time = true;
                            }
                        }
                        else{
                            if(false == prev.time && false == next.time){
                                bError = false;
                                prev.date = true;
                                next.date = true;
                            }
                        }
                    }
                }
                else if (i - 1 >= 0 && i + 1 == length) {
                    //case "10:"
                    var prev = match[i - 1];
                    if (prev.type != oDataTypes.delimiter) {
                        if (cultureInfo.TimeSeparator == elem.val || (":" == elem.val && cultureInfo.DateSeparator != elem.val)) {
                            if (false == prev.date) {
                                bError = false;
                                prev.time = true;
                            }
                        }
                    }
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            //separate date and time using Am/Pm and month names
            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.type == oDataTypes.letter){
                    var valLower = elem.val.toLowerCase();
                    if (elem.am || elem.pm) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.date) {
                                prev.time = true;
                            }
                        }
                        //AmPm should be the last entry
                        if (i + 1 != length) {
                            bError = true;
                        }
                    }
                    else if (null != elem.month) {
                        if (i - 1 >= 0) {
                            var prev = match[i - 1];
                            if (oDataTypes.digit == prev.type && false == prev.time)
                                prev.date = true;
                        }
                        if (i + 1 < length) {
                            var next = match[i + 1];
                            if (oDataTypes.digit == next.type && false == next.time)
                                next.date = true;
                        }
                    }
                    else
                        bError = true;
                }
                if(bError)
                    break;
            }
        }
        if(!bError){
            var aDate = [];
            var nMonthIndex = null;
			var sMonthFormat = null;
			var monthDone = false;
            var aTime = [];
            var am = false;
            var pm = false;

			var nIndexD = Math.max(cultureInfo.ShortDatePattern.indexOf("0"), cultureInfo.ShortDatePattern.indexOf("1"));
			var nIndexM = Math.max(cultureInfo.ShortDatePattern.indexOf("2"), cultureInfo.ShortDatePattern.indexOf("3"));
            var nIndexY = Math.max(cultureInfo.ShortDatePattern.indexOf("4"), cultureInfo.ShortDatePattern.indexOf("5"));

            for (var i = 0, length = match.length; i < length; i++) {
                var elem = match[i];
                if (elem.date || (elem.time == false && elem.type == oDataTypes.digit)) {
                    if (elem.type == oDataTypes.digit)
                        aDate.push(elem.val);
                    else if (elem.type == oDataTypes.letter && null != elem.month) {
                        if (aDate.length >= 3)
							continue;
							
						nMonthIndex = aDate.length;
                        sMonthFormat = elem.month.format;
                        aDate.push(elem.month.val);
						monthDone = true;
                    }
                    else
                        bError = true;
                }
                else if (elem.time) {
                    if (elem.type == oDataTypes.digit)
                        aTime.push(elem.val);
                    else if (elem.type == oDataTypes.letter && (elem.am || elem.pm)) {
                        am = elem.am;
                        pm = elem.pm;
                    }
                    else
                        bError = true;
                }
            }
			if (aDate.length > 3)
				aDate.length = 3;

            var nDateLength = aDate.length;
            var nTimeLength = aTime.length;
            if (nTimeLength > 3)
                aTime.length = 3;
            if(!bError){
                res = { d: null, m: null, y: null, h: null, min: null, s: null, am: am, pm: pm, sDateFormat: null };
                if (nDateLength > 0) {
                    if (null != nMonthIndex) {
                        res.m = aDate[nMonthIndex];

						if (nIndexD != -1) {
							if (nIndexD != nMonthIndex) {
								res.d = aDate[nIndexD];
							}
							else {
								if (aDate[0] <= 31) {
									res.d = aDate[0];
									res.y = aDate[2];
								}
								else {
									res.d = aDate[2];
									res.y = aDate[0];
								}
							}
						}
						
						if (nIndexY != -1 && res.y == null) {
							if (nIndexY != nMonthIndex) {
								res.y = aDate[nIndexY];
							}
							else {
								res.d = aDate[0];
								res.y = aDate[2];
							}
						}
                    }
                    else {
                        res.m = aDate[nIndexM];
						res.d = aDate[nIndexD];
						res.y = aDate[nIndexY];
                    }
                    if(null != res.y)
                    {
                        if(res.y < 30)
                            res.y = 2000 + res.y;
                        else if(res.y < 100)
                            res.y = 1900 + res.y;
                    }
                }
                if(nTimeLength > 0){
                    res.h = aTime[0];
                    if(nTimeLength > 1)
                        res.min = aTime[1];
                    if(nTimeLength > 2)
                        res.s = aTime[2];
                }
                if(bError)
                    res = null;
            }
        }
		return res;
    },
    strcmp: function (s1, s2, index1, length, index2) {
        if (null == index2)
            index2 = 0;
        var bRes = true;
        for (var i = 0; i < length; ++i) {
            if (s1[index1 + i] != s2[index2 + i]) {
                bRes = false;
                break;
            }
        }
        return length === 0 ? false: bRes;
    },
	parseDate: function (value, cultureInfo, shouldPreserveFormat, isDateOverrideFormat, currentFormat, stringFormat)
	{
		//todo "11: AM" should fail
		var res = null;
		var match = [];
		var sCurValue = null;
		var oCurDataType = null;
		var oPrevType = null;
		var bAmPm = false;
		var bMonth = false;
		var bError = false;
		var oDataTypes = {letter: {id: 0, min: 2, max: 9}, digit: {id: 1, min: 1, max: 4}, delimiter: {id: 2, min: 1, max: 1}, space: {id: 3, min: null, max: null}};
		var valueLower = value.toLowerCase();
		for(var i = 0, length = value.length; i < length; i++)
		{
		    var sChar = value[i];
		    var oDataType = null;
		    if("0" <= sChar && sChar <= "9")
		        oDataType = oDataTypes.digit;
		    else if(" " == sChar || "," == sChar)
		        oDataType = oDataTypes.space;
		    else if ("/" == sChar || "-" == sChar || ":" == sChar || cultureInfo.DateSeparator == sChar || cultureInfo.TimeSeparator == sChar)
		        oDataType = oDataTypes.delimiter;
		    else
		        oDataType = oDataTypes.letter;
			    
		    if(null != oDataType)
		    {
		        if(null == oCurDataType)
		            sCurValue = sChar;
		        else
		        {
		            if(oCurDataType == oDataType)
		            {
		                if(null == oCurDataType.max || sCurValue.length < oCurDataType.max)
		                    sCurValue += sChar;
		                else
		                    bError = true;
		            }
		            else
		            {
		                if (null == oCurDataType.min || sCurValue.length >= oCurDataType.min) {
		                    if (oDataTypes.space != oCurDataType) {
		                        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		                        if (oDataTypes.digit == oCurDataType)
		                            oNewElem.val = oNewElem.val - 0;
		                        match.push(oNewElem);
		                    }
		                    sCurValue = sChar;
		                    oPrevType = oCurDataType;
		                }
		                else
		                    bError = true;
		            }
		        }
		        oCurDataType = oDataType;
		    }
		    else
		        bError = true;
		    if(oDataTypes.letter == oDataType){
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        var bAm = false;
		        var bPm = false;
		        if (!bAmPm && ((bAm = this.strcmp(valueLower, "am", i, 2)) || (bPm = this.strcmp(valueLower, "pm", i, 2)))) {
		            bAmPm = true;
		            oNewElem.am = bAm;
		            oNewElem.pm = bPm;
		            oNewElem.time = true;
		            match.push(oNewElem);
		            i += 2 - 1;
		            if (oPrevType != oDataTypes.space)
		                bError = true;
		        }
		        else if (!bMonth) {
		            bMonth = true;
					let aArraysToCheck = [{ arr: cultureInfo.MonthNames, format: "mmmm" }, { arr: cultureInfo.AbbreviatedMonthNames, format: "mmm" }];
		            var bFound = false;
		            for (var index in aArraysToCheck) {
		                var aArrayTemp = aArraysToCheck[index];
		                for (var j = 0, length2 = aArrayTemp.arr.length; j < length2; j++) {
		                    var sCmpVal = aArrayTemp.arr[j].toLowerCase();
		                    var sCmpValCrop = sCmpVal.replace(/\./g, "");
		                    var bCrop = false;
		                    if (this.strcmp(valueLower, sCmpVal, i, sCmpVal.length) || (bCrop = (sCmpVal != sCmpValCrop && this.strcmp(valueLower, sCmpValCrop, i, sCmpValCrop.length)))) {
		                        bFound = true;
		                        oNewElem.month = { val: j + 1, format: aArrayTemp.format };
		                        oNewElem.date = true;
		                        if (bCrop)
		                            i += sCmpValCrop.length - 1;
		                        else
		                            i += sCmpVal.length - 1;
		                        break;
		                    }
		                }
		                if (bFound)
		                    break;
		            }
		            //nothing other than month name can be present
		            if (bFound)
		                match.push(oNewElem);
		            else
		                bError = true;
		        }
		        else
		            bError = true;
		        oCurDataType = null;
		        sCurValue = null;
		    }
			if (bError)
			{
				match = null;
				break;
			}
		}
		if (null != match && null != sCurValue) {
		    if (oDataTypes.space != oCurDataType) {
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        if (oDataTypes.digit == oCurDataType)
		            oNewElem.val = oNewElem.val - 0;
		        match.push(oNewElem);
		    }
		}
		if(null != match && match.length > 0)
		{
		    var oParsedDate = this._parseDateFromArray(match, oDataTypes, cultureInfo);
			if(null != oParsedDate)
			{
				var d = oParsedDate.d;
				var m = oParsedDate.m;
				var y = oParsedDate.y;
				var h = oParsedDate.h;
				var min = oParsedDate.min;
				var s = oParsedDate.s;
				var am = oParsedDate.am;
				var pm = oParsedDate.pm;
				var sDateFormat = oParsedDate.sDateFormat;
				
				var bDate = false;
				var bTime = false;
				var bSeconds = false;
				var nDay;
				var nMounth;
				var nYear;
				if(AscCommon.bDate1904)
				{
					nDay = 1;
					nMounth = 0;
					nYear = 1904;
				}
				else
				{
					nDay = 31;
					nMounth = 11;
					nYear = 1899;
				}
				var nHour = 0;
				var nMinute = 0;
				var nSecond = 0;
				var dValue = 0;
				var bValidDate = true;
				if(null != m && (null != d || null != y))
				{
					bDate = true;
					var oNowDate;
					if(null != d)
						nDay = d - 0;
					else
						nDay = 1;
					nMounth = m - 1;
					if(null != y)
						nYear = y - 0;
					else
                    {
                        oNowDate = new Date();
						nYear = oNowDate.getFullYear();
                    }
					
					//check date validity
					bValidDate = this.isValidDate(nYear, nMounth, nDay);
				}
				if(null != h)
				{
					bTime = true;
					nHour = h - 0;
					if (am || pm)
					{
						if(nHour <= 23)
						{
							//convert 24
							nHour = nHour % 12;
							if(pm)
								nHour += 12;
						}
						else
							bValidDate = false;
					}
					if(null != min)
					{
						nMinute = min - 0;
						if(nMinute > 59)
							bValidDate = false;
					}
					if(null != s)
					{
						nSecond = s - 0;
						if (0 <= nSecond && nSecond < 60) {
							bSeconds = true;
						} else {
							bValidDate = false;
						}
					}
				}
				if(true == bValidDate && (true == bDate || true == bTime))
				{
					if(AscCommon.bDate1904)
						dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1904,0,1,0,0,0)) / (86400 * 1000);
					else
					{
						if(1900 < nYear || (1900 == nYear && 1 < nMounth ))
							dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1899,11,30,0,0,0)) / (86400 * 1000);
						else if(1900 == nYear && 1 == nMounth && 29 == nDay)
							dValue = 60;
						else
							dValue = (Date.UTC(nYear,nMounth,nDay,nHour,nMinute,nSecond) - Date.UTC(1899,11,31,0,0,0)) / (86400 * 1000);
					}
					if(dValue >= 0)
					{
						var sFormat = "";

						const needOverrideFormat = isDateOverrideFormat &&
							((currentFormat === Asc.c_oAscNumFormatType.Date && !bDate) ||
							 (currentFormat === Asc.c_oAscNumFormatType.LongDate && bTime) ||
							 (currentFormat === Asc.c_oAscNumFormatType.Time && bDate));
                        // Check if current format should be preserved (not converted to date)
                        if (shouldPreserveFormat && !needOverrideFormat) {
                            sFormat = stringFormat;
                        } else if (bDate) {
							if (bTime && nHour > 23) {
								sFormat = AscCommon.g_cGeneralFormat;
							} else {
								sFormat += sDateFormat;
								if (bTime) {
									sFormat += " h:mm";
								}
							}
						} else {
							if (dValue > 1) {
								sFormat += "[h]:mm";
							} else {
								sFormat += "h:mm";
							}
							if (bSeconds || dValue > 1) {
								sFormat += ":ss";
							}
							if (am || pm)
								sFormat += " AM/PM";
						}
						res = {format: sFormat, value: dValue, bDateTime: true, bDate: bDate, bTime: bTime, bPercent: false, bCurrency: false};
					}
				}
            }
        }
		return res;
	},
	parseDatePDF: function (value, cultureInfo, oFormat)
	{
		if (null == cultureInfo)
			cultureInfo = g_oDefaultCultureInfo;
		let res = null;
		let match = [];
		let sCurValue = null;
		let oCurDataType = null;
		let oPrevType = null;
		let bAmPm = false;
		let bMonth = false;
		let bError = false;
		let oDataTypes = {letter: {id: 0, min: 2, max: 9}, digit: {id: 1, min: 1, max: 4}, delimiter: {id: 2, min: 1, max: 1}, space: {id: 3, min: null, max: null}};
		let valueLower = value.toLowerCase();
		for(var i = 0, length = value.length; i < length; i++)
		{
		    var sChar = value[i];
		    var oDataType = null;
		    if("0" <= sChar && sChar <= "9")
		        oDataType = oDataTypes.digit;
		    else if(" " == sChar)
		        oDataType = oDataTypes.space;
		    else if ("." == sChar || "/" == sChar || "-" == sChar || ":" == sChar || "," == sChar || cultureInfo.DateSeparator == sChar || cultureInfo.TimeSeparator == sChar)
		        oDataType = oDataTypes.delimiter;
		    else
		        oDataType = oDataTypes.letter;
			    
			// after delimiter there can be month again
			if (oDataType == oDataTypes.delimiter)
				bMonth = false;

		    if(null != oDataType)
		    {
		        if(null == oCurDataType)
		            sCurValue = sChar;
		        else
		        {
		            if(oCurDataType == oDataType)
		            {
		                if(null == oCurDataType.max || sCurValue.length < oCurDataType.max)
		                    sCurValue += sChar;
		                else
		                    bError = true;
		            }
		            else
		            {
		                if (null == oCurDataType.min || sCurValue.length >= oCurDataType.min) {
		                    if (oDataTypes.space != oCurDataType) {
		                        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		                        if (oDataTypes.digit == oCurDataType)
		                            oNewElem.val = oNewElem.val - 0;
								if (oNewElem.val < 100 && sCurValue.length == 4)
									bError = true; // year less than hundred, example: year 0001
		                        
								match.push(oNewElem);
		                    }
		                    sCurValue = sChar;
		                    oPrevType = oCurDataType;
		                }
		                else
		                    bError = true;
		            }
		        }
		        oCurDataType = oDataType;
		    }
		    else
		        bError = true;
		    if(oDataTypes.letter == oDataType){
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        var bAm = false;
		        var bPm = false;
		        if (!bAmPm && ((bAm = this.strcmp(valueLower, "am", i, 2)) || (bPm = this.strcmp(valueLower, "pm", i, 2)))) {
		            bAmPm = true;
		            oNewElem.am = bAm;
		            oNewElem.pm = bPm;
		            oNewElem.time = true;
		            match.push(oNewElem);
		            i += 2 - 1;
		            if (oPrevType != oDataTypes.space)
		                bError = true;
		        }
		        else if (!bMonth) {
		            bMonth = true;
		            var aArraysToCheck = [{ arr: cultureInfo.MonthNames, format: "mmmm" }, { arr: cultureInfo.AbbreviatedMonthNames, format: "mmm" }];
		            var bFound = false;
		            for (var index in aArraysToCheck) {
		                var aArrayTemp = aArraysToCheck[index];
		                for (var j = 0, length2 = aArrayTemp.arr.length; j < length2; j++) {
		                    var sCmpVal = aArrayTemp.arr[j].toLowerCase();
		                    var sCmpValCrop = sCmpVal.replace(/\./g, "");
		                    var bCrop = false;
		                    if (this.strcmp(valueLower, sCmpVal, i, sCmpVal.length) || (bCrop = (sCmpVal != sCmpValCrop && this.strcmp(valueLower, sCmpValCrop, i, sCmpValCrop.length)))) {
		                        bFound = true;
		                        oNewElem.month = { val: j + 1, format: aArrayTemp.format };
		                        oNewElem.date = true;
		                        if (bCrop)
		                            i += sCmpValCrop.length - 1;
		                        else
		                            i += sCmpVal.length - 1;
		                        break;
		                    }
		                }
		                if (bFound)
		                    break;
		            }
		            //nothing other than month name can be present
		            if (bFound)
		                match.push(oNewElem);
		            else
		                bError = true;
		        }
		        else
		            bError = true;
		        oCurDataType = null;
		        sCurValue = null;
		    }
			if (bError)
			{
				match = null;
				break;
			}
		}
		if (null != match && null != sCurValue) {
		    if (oDataTypes.space != oCurDataType) {
		        var oNewElem = { val: sCurValue, type: oCurDataType, month: null, am: false, pm: false, date: false, time: false };
		        if (oDataTypes.digit == oCurDataType)
		            oNewElem.val = oNewElem.val - 0;

		        match.push(oNewElem);
		    }
		}
		if(null != match && match.length > 0)
		{
		    var oParsedDate = this._parseDateFromArrayPDF(match, oDataTypes, cultureInfo, oFormat);
			if(null != oParsedDate)
			{
				var d = oParsedDate.d;
				var m = oParsedDate.m;
				var y = oParsedDate.y;
				var h = oParsedDate.h;
				var min = oParsedDate.min;
				var s = oParsedDate.s;
				var am = oParsedDate.am;
				var pm = oParsedDate.pm;
				var sDateFormat = oParsedDate.sDateFormat;
				
				var bDate = false;
				var bTime = false;
				var nDay;
				var nMounth;
				var nYear;
				if(AscCommon.bDate1904)
				{
					nDay = 1;
					nMounth = 0;
					nYear = 1904;
				}
				else
				{
					nDay = 31;
					nMounth = 11;
					nYear = 1899;
				}
				var nHour = 0;
				var nMinute = 0;
				var nSecond = 0;
				var dValue = 0;
				var bValidDate = true;
				if(null != m)
				{
					bDate = true;
					var oNowDate;
					if(null != d)
						nDay = d - 0;
					else
						nDay = 1;
					nMounth = m - 1;
					if(null != y)
						nYear = y - 0;
					else
                    {
                        oNowDate = new Date();
						nYear = oNowDate.getFullYear();
                    }
					
					//check date validity
					bValidDate = this.isValidDatePDF(nYear, nMounth, nDay);
				}
				if(null != h)
				{
					bTime = true;
					nHour = h - 0;
					if (am || pm)
					{
						if(nHour <= 23)
						{
							//convert 24
							nHour = nHour % 12;
							if(pm)
								nHour += 12;
						}
						else
							bValidDate = false;
					}
					if(null != min)
					{
						nMinute = min - 0;
						if(nMinute > 59)
							bValidDate = false;
					}
					if(null != s)
					{
						nSecond = s - 0;
						if(nSecond > 59)
							bValidDate = false;
					}
				}
				if(true == bValidDate && (true == bDate || true == bTime))
				{
					var oDateTmp = new Date();
					oDateTmp.setFullYear(nYear, nMounth, nDay);
					oDateTmp.setHours(nHour, nMinute, nSecond);
					dValue = oDateTmp.getTime() / (86400 * 1000);

					var sFormat;
					if(true == bDate && true == bTime)
					{
						sFormat = sDateFormat + " h:mm:ss";
						if (am || pm)
							sFormat += " AM/PM";
					}
					else if(true == bDate)
						sFormat = sDateFormat;
					else
					{
						if(dValue > 1)
							sFormat = "[h]:mm:ss";
						else if (am || pm)
							sFormat = "h:mm:ss AM/PM";
						else
							sFormat = "h:mm:ss";
					}
					res = {format: sFormat, value: dValue, bDateTime: true, bDate: bDate, bTime: bTime, bPercent: false, bCurrency: false};
				}
            }
        }
		return res;
	},
	isValidDate : function(nYear, nMounth, nDay)
	{
		if(nYear < 1900 && !(1899 === nYear && 11 == nMounth && 31 == nDay))
			return false;
		else
		{
			if(nMounth < 0 || nMounth > 11)
				return false;
			else if(this.isValidDay(nYear, nMounth, nDay))
				return true;
			else if(1900 == nYear && 1 == nMounth && 29 == nDay)
				return true;
		}
		return false;
	},
	isValidDatePDF : function(nYear, nMounth, nDay)
	{
		if(nMounth < 0 || nMounth > 11)
			return false;
		else if(this.isValidDay(nYear, nMounth, nDay))
			return true;
		else if(1900 == nYear && 1 == nMounth && 29 == nDay)
			return true;
		return false;
	},
	isValidDay : function(nYear, nMounth, nDay){
		if(this.isLeapYear(nYear))
		{
			if(nDay <= 0 || nDay > this.daysLeap[nMounth])
				return false;
		}
		else
		{
			if(nDay <= 0 || nDay > this.days[nMounth])
				return false;
		}
		return true;
	},
	isLeapYear : function(year)
	{
		return (0 == (year % 4)) && (0 != (year % 100) || 0 == (year % 400))
	}
};
var g_oFormatParser = new FormatParser();
function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}
function makeStringCompare(lang) {
	const locale = lang || "en";
	return function(s1, s2) {
		if (s1 === s2) {
			return 0;
		}
		return s1.localeCompare(s2, locale);
	};
}
AscCommon.stringCompare = makeStringCompare(AscCommon.g_oDefaultCultureInfo ? AscCommon.g_oDefaultCultureInfo.Name : "en");
function setCurrentCultureInfo (LCID, decimalSeparator, groupSeparator) {
	var res = false;
	var cultureInfoNew = g_aCultureInfos[LCID];
	if (cultureInfoNew) {
		if (LCID !== g_oLCID) {
			g_oLCID = LCID;
			AscCommon.g_oDefaultCultureInfo = g_oDefaultCultureInfo = JSON.parse(JSON.stringify(cultureInfoNew)); // ToDo clone
			AscCommon.stringCompare = makeStringCompare(g_oDefaultCultureInfo.Name);
			interfaceShortDateFormat = getShortDateFormat(g_oDefaultCultureInfo);
			res = true;
		}
		ParseLocalFormatSymbol(g_oDefaultCultureInfo.Name);
		decimalSeparator = (null != decimalSeparator) ? decimalSeparator : cultureInfoNew.NumberDecimalSeparator;
		if (decimalSeparator !== g_oDefaultCultureInfo.NumberDecimalSeparator) {
			g_oDefaultCultureInfo.NumberDecimalSeparator = decimalSeparator;
			res = true;
		}
		groupSeparator = (null != groupSeparator) ? groupSeparator : cultureInfoNew.NumberGroupSeparator;
		if (groupSeparator !== g_oDefaultCultureInfo.NumberGroupSeparator) {
			g_oDefaultCultureInfo.NumberGroupSeparator = groupSeparator;
			res = true;
		}
	}
	return res;
}
	function checkCultureInfoFontPicker(LCID) {
		var ci = g_aCultureInfos[LCID] || g_oDefaultCultureInfo;
		AscFonts.FontPickerByCharacter.getFontsByString(ci.CurrencySymbol);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.NumberDecimalSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.NumberGroupSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.AMDesignator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.PMDesignator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.DateSeparator);
		AscFonts.FontPickerByCharacter.getFontsByString(ci.TimeSeparator);
		var arrays = [ci.DayNames, ci.AbbreviatedDayNames, ci.MonthNames, ci.AbbreviatedMonthNames,
			ci.MonthGenitiveNames, ci.AbbreviatedMonthGenitiveNames
		];
		arrays.forEach(function(arr){
			arr.forEach(function(text) {
				AscFonts.FontPickerByCharacter.getFontsByString(text);
			});
		});
	}

	function isDMY(cultureInfo) {
		//day month year
		var res = true;
		for (var i = 0; i < cultureInfo.ShortDatePattern.length - 1; ++i) {
			if (cultureInfo.ShortDatePattern.charCodeAt(i) > cultureInfo.ShortDatePattern.charCodeAt(i + 1)) {
				return false;
			}
		}
		return true;
	}
	function isYMD(cultureInfo) {
		//year month day
		var res = true;
		for (var i = 0; i < cultureInfo.ShortDatePattern.length - 1; ++i) {
			if (cultureInfo.ShortDatePattern.charCodeAt(i) < cultureInfo.ShortDatePattern.charCodeAt(i + 1)) {
				return false;
			}
		}
		return true;
	}
	function getShortDateMonthFormat(bDate, bYear, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var separator;
		if ('/' == g_oDefaultCultureInfo.DateSeparator) {
			separator = '-';
		} else {
			separator = '/';
		}
		var sRes = '';
		if (bDate) {
			if (-1 != cultureInfo.ShortDatePattern.indexOf('1')) {
				sRes += 'dd';
			} else {
				sRes += 'd';
			}
			sRes += separator;
		}
		sRes += 'mmm';
		if (bYear) {
			sRes += separator;
			sRes += 'yy';
		}
		return sRes;
	}
	function getShortDateFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var dateElems = [];
		for (var i = 0; i < cultureInfo.ShortDatePattern.length; ++i) {
			switch (cultureInfo.ShortDatePattern[i]) {
				case '0':
					dateElems.push('d');
					break;
				case '1':
					dateElems.push('dd');
					break;
				case '2':
					dateElems.push('m');
					break;
				case '3':
					dateElems.push('mm');
					break;
				case '4':
					dateElems.push('yy');
					break;
				case '5':
					dateElems.push('yyyy');
					break;
			}
		}
		return dateElems.join('/');
	}

	function getShortDateFormat2(day, month, year, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var dateElems = [];
		for (var i = 0; i < cultureInfo.ShortDatePattern.length; ++i) {
			switch (cultureInfo.ShortDatePattern[i]) {
				case '0':
				case '1':
					if (day > 0) {
						dateElems.push('d'.repeat(day));
					}
					break;
				case '2':
				case '3':
					if (month > 0) {
						dateElems.push('m'.repeat(month));
					}
					break;
				case '4':
				case '5':
					if (year > 0) {
						dateElems.push('y'.repeat(year));
					}
					break;
			}
		}
		return dateElems.join('/');
	}

	function getShortTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		if (AscCommon.is12HourTimeFormat(cultureInfo)) {
			return 'h:mm AM/PM;@';
		} else {
			return 'h:mm;@'
		}
	}
	function getLongTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		if (AscCommon.is12HourTimeFormat(cultureInfo)) {
			return 'h:mm:ss AM/PM;@';
		} else {
			return 'h:mm:ss;@'
		}
	}

	function getNumberFormatSimple(opt_separate, opt_fraction) {
		var numberFormat = opt_separate ? '#,##0' : '0';
		if (opt_fraction > 0) {
			numberFormat += '.' + '0'.repeat(opt_fraction);
		}
		return numberFormat;
	}

	function getNumberFormat(opt_cultureInfo, opt_separate, opt_fraction, opt_red) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(opt_separate, opt_fraction);
		var red = opt_red ? '[Red]' : '';

		var positiveFormat;
		var negativeFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
			case 4:
			case 14:
			case 15:
				positiveFormat = numberFormat + '_)';
				negativeFormat = '\\(' + numberFormat + '\\)';
				break;
			default:
				positiveFormat = numberFormat + '_ ';
				negativeFormat = '\\-' + numberFormat + '\\ ';
				break;
		}
		return positiveFormat + ';' + red + negativeFormat;
	}

	function getLocaleFormat(opt_cultureInfo, opt_currency) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var symbol = opt_currency ? cultureInfo.CurrencySymbol : '';
		return '[$' + symbol + '-' + cultureInfo.LCID.toString(16).toUpperCase() + ']';
	}
	function getCurrencyCustomFormat(symbol) {
		return '[$' + symbol + ']';
	}

	function getCurrencyFormatSimple(opt_cultureInfo, opt_fraction, opt_currency, opt_currencyLocale, opt_currencySymbol, opt_red) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				if (opt_currencyLocale) {
					signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				} else {
					signCurrencyFormat = '"' + cultureInfo.CurrencySymbol + '"';
				}
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}
		var red = opt_red ? '[Red]' : '';

		var prefixs = ['_ ', '_-', '_(', '_)'];
		var postfix = '';
		var positiveFormat;
		var negativeFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
				postfix = prefixs[3];
				negativeFormat = '\\(' + signCurrencyFormat + numberFormat + '\\)';
				break;
			case 1:
				negativeFormat = '\\-' + signCurrencyFormat + numberFormat;
				break;
			case 2:
				negativeFormat = signCurrencyFormatSpace + '\\-' + numberFormat;
				break;
			case 3:
				postfix = prefixs[1];
				negativeFormat = signCurrencyFormatSpace + numberFormat + '\\-';
				break;
			case 4:
				postfix = prefixs[3];
				negativeFormat = '\\(' + numberFormat + signCurrencyFormatEnd + '\\)';
				break;
			case 5:
				negativeFormat = '\\-' + numberFormat + signCurrencyFormatEnd;
				break;
			case 6:
				negativeFormat = numberFormat + '\\-' + signCurrencyFormatEnd;
				break;
			case 7:
				postfix = prefixs[1];
				negativeFormat = numberFormat + signCurrencyFormatEnd + '\\-';
				break;
			case 8:
				negativeFormat = '\\-' + numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
			case 9:
				negativeFormat = '\\-' + signCurrencyFormatSpace + numberFormat;
				break;
			case 10:
				postfix = prefixs[1];
				negativeFormat = numberFormat + '\\ ' + signCurrencyFormatEnd + '\\-';
				break;
			case 11:
				postfix = prefixs[1];
				negativeFormat = signCurrencyFormatSpace + numberFormat + '\\-';
				break;
			case 12:
				negativeFormat = signCurrencyFormatSpace + '\\-' + numberFormat;
				break;
			case 13:
				negativeFormat = numberFormat + '\\-\\ ' + signCurrencyFormatEnd;
				break;
			case 14:
				postfix = prefixs[3];
				negativeFormat = '(' + signCurrencyFormat + numberFormat + '\\)';
				break;
			case 15:
				postfix = prefixs[3];
				negativeFormat = '\\(' + numberFormat + signCurrencyFormatEnd + '\\)';
				break;
		}
		switch (cultureInfo.CurrencyPositivePattern) {
			case 0:
				positiveFormat = signCurrencyFormat + numberFormat;
				break;
			case 1:
				positiveFormat = numberFormat + signCurrencyFormatEnd;
				break;
			case 2:
				positiveFormat = signCurrencyFormatSpace + numberFormat;
				break;
			case 3:
				positiveFormat = numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		positiveFormat = positiveFormat + postfix;
		return positiveFormat + ';' + red + negativeFormat;
	}

	function getCurrencyFormatSimple2(opt_cultureInfo, opt_fraction, opt_currency, opt_currencySymbol, opt_negative) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}
		var positiveFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
			case 1:
			case 14:
				positiveFormat = signCurrencyFormat + numberFormat;
				break;
			case 2:
			case 3:
			case 9:
			case 10:
			case 11:
			case 12:
				positiveFormat = signCurrencyFormatSpace + numberFormat;
				break;
			case 4:
			case 5:
			case 6:
			case 7:
			case 15:
				positiveFormat = numberFormat + signCurrencyFormatEnd;
				break;
			case 8:
			case 13:
				positiveFormat = numberFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		return opt_negative ? positiveFormat + ';[Red]' + positiveFormat : positiveFormat;
	}

	function getCurrencyFormat(opt_cultureInfo, opt_fraction, opt_currency, opt_currencyLocale, opt_currencySymbol) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		var numberFormat = getNumberFormatSimple(true, opt_fraction);
		var nullSignFormat = '* "-"';
		if (opt_fraction) {
			nullSignFormat += '?'.repeat(opt_fraction);
		}
		var signCurrencyFormat;
		var signCurrencyFormatEnd;
		var signCurrencyFormatSpace;
		if (opt_currency) {
			if (opt_currencySymbol) {
				signCurrencyFormat = getCurrencyCustomFormat(opt_currencySymbol);
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormat = signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			} else {
				if (opt_currencyLocale) {
					signCurrencyFormat = getLocaleFormat(cultureInfo, true);
				} else {
					signCurrencyFormat = '"' + cultureInfo.CurrencySymbol + '"';
				}
				signCurrencyFormatEnd = signCurrencyFormat;
				signCurrencyFormatSpace = signCurrencyFormat + '\\ ';
			}
		} else {
			signCurrencyFormatEnd = signCurrencyFormat = signCurrencyFormatSpace = '';
			for (var i = 0; i < cultureInfo.CurrencySymbol.length; ++i) {
				signCurrencyFormatEnd += '_' + cultureInfo.CurrencySymbol[i];
			}
		}

		var prefixs = ['_ ', '_-', '_(', '_)'];
		var prefix = prefixs[0];
		var postfix = prefixs[0];
		var positiveNumberFormat = '* ' + numberFormat;
		var positiveFormat;
		var negativeFormat;
		var nullFormat;
		switch (cultureInfo.CurrencyNegativePattern) {
			case 0:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + signCurrencyFormat + '* \\(' + numberFormat + '\\)';
				break;
			case 1:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-' + signCurrencyFormat + '* ' + numberFormat + postfix;
				break;
			case 2:
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\-' + numberFormat + postfix;
				break;
			case 3:
				prefix = postfix = prefixs[1];
				negativeFormat = prefix + signCurrencyFormatSpace + '* ' + numberFormat + '\\-';
				break;
			case 4:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + '* \\(' + numberFormat + '\\)' + signCurrencyFormatEnd + postfix;
				break;
			case 5:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-* ' + numberFormat + signCurrencyFormatEnd + postfix;
				break;
			case 6:
				negativeFormat = prefix + '* ' + numberFormat + '\\-' + signCurrencyFormatEnd + postfix;
				break;
			case 7:
				negativeFormat = prefix + '* ' + numberFormat + signCurrencyFormatEnd + '\\-';
				break;
			case 8:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-* ' + numberFormat + '\\ ' + signCurrencyFormatEnd + postfix;
				break;
			case 9:
				prefix = postfix = prefixs[1];
				negativeFormat = '\\-' + signCurrencyFormatSpace + '* ' + numberFormat + postfix;
				break;
			case 10:
				negativeFormat = prefix + '* ' + numberFormat + '\\ ' + signCurrencyFormatEnd + '\\-';
				break;
			case 11:
				negativeFormat = prefix + signCurrencyFormatSpace + '* ' + numberFormat + '\\-';
				break;
			case 12:
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\-' + numberFormat + postfix;
				break;
			case 13:
				negativeFormat = prefix + '* ' + numberFormat + '\\-\\ ' + signCurrencyFormatEnd + postfix;
				break;
			case 14:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + signCurrencyFormatSpace + '* \\(' + numberFormat + '\\)';
				break;
			case 15:
				prefix = prefixs[2];
				postfix = prefixs[3];
				negativeFormat = prefix + '* \\(' + numberFormat + '\\)\\ ' + signCurrencyFormatEnd + postfix;
				break;
		}
		switch (cultureInfo.CurrencyPositivePattern) {
			case 0:
				positiveFormat = signCurrencyFormat + positiveNumberFormat;
				nullFormat = signCurrencyFormat + nullSignFormat;
				break;
			case 1:
				positiveFormat = positiveNumberFormat + signCurrencyFormatEnd;
				nullFormat = nullSignFormat + signCurrencyFormatEnd;
				break;
			case 2:
				positiveFormat = signCurrencyFormatSpace + positiveNumberFormat;
				nullFormat = signCurrencyFormatSpace + nullSignFormat;
				break;
			case 3:
				positiveFormat = positiveNumberFormat + '\\ ' + signCurrencyFormatEnd;
				nullFormat = nullSignFormat + '\\ ' + signCurrencyFormatEnd;
				break;
		}
		positiveFormat = prefix + positiveFormat + postfix;
		nullFormat = prefix + nullFormat + postfix;
		var textFormat = prefix + '@' + postfix;
		return positiveFormat + ';' + negativeFormat + ';' + nullFormat + ';' + textFormat;
	}

	function getFormatCells(info) {
		var res = [];
		if (info) {
			var format;
			var i;
			var currencySymbol = info.currency;
			var cultureInfo = g_aCultureInfos[info.symbol];
			var hasCurrency = !!cultureInfo || !!currencySymbol;
			if (Asc.c_oAscNumFormatType.General === info.type) {
				res.push(AscCommon.g_cGeneralFormat);
			} else if (Asc.c_oAscNumFormatType.Number === info.type) {
				var numberFormat = getNumberFormatSimple(info.separator, info.decimalPlaces);
				res.push(numberFormat);
				res.push(numberFormat + ';[Red]' + numberFormat);
				res.push(getNumberFormat(cultureInfo, info.separator, info.decimalPlaces, false));
				res.push(getNumberFormat(cultureInfo, info.separator, info.decimalPlaces, true));
			} else if (Asc.c_oAscNumFormatType.Currency === info.type) {
				res.push(getCurrencyFormatSimple2(cultureInfo, info.decimalPlaces, hasCurrency, currencySymbol, false));
				res.push(getCurrencyFormatSimple2(cultureInfo, info.decimalPlaces, hasCurrency, currencySymbol, true));
				res.push(getCurrencyFormatSimple(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol, false));
				res.push(getCurrencyFormatSimple(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol, true));
			} else if (Asc.c_oAscNumFormatType.Accounting === info.type) {
				res.push(getCurrencyFormat(cultureInfo, info.decimalPlaces, hasCurrency, true, currencySymbol));
			} else if (Asc.c_oAscNumFormatType.Date === info.type) {
				//todo locale dependence
				if (info.symbol === g_oDefaultCultureInfo.LCID) {
					res.push(getShortDateFormat(cultureInfo));
					res.push('[$-F800]' + cultureInfo.LongDatePattern);
				}
				if (c_oAscDateFormatExcel[info.symbol]) {
					res = res.concat(c_oAscDateFormatExcel[info.symbol]);
				}
				//todo remove (backward compat)
				res.push(getShortDateFormat2(1, 1, 0, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 0, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 2, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 2, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 4, cultureInfo) + ';@');
				res.push(getShortDateFormat2(2, 2, 4, cultureInfo) + ';@');
				res.push(getShortDateFormat2(1, 1, 2, cultureInfo) + ' h:mm;@');
				res.push(getShortDateFormat2(2, 2, 2, cultureInfo) + ' h:mm;@');
				res.push('[$-409]' + getShortDateFormat2(1, 1, 2, cultureInfo) + ' h:mm AM/PM;@');
				var locale = getLocaleFormat(cultureInfo, false);
				res.push(locale + 'mmmmm;@');
				res.push(locale + 'mmmm d, yyyy;@');
				var separators = ['-', '/', ' '];
				for (i = 0; i < separators.length; ++i) {
					var separator = separators[i];
					res.push(locale + 'd' + separator + 'mmm;@');
					res.push(locale + 'd' + separator + 'mmm' + separator + 'yy;@');
					res.push(locale + 'dd' + separator + 'mmm' + separator + 'yy;@');
					res.push(locale + 'mmm' + separator + 'yy;@');
					res.push(locale + 'mmmm' + separator + 'yy;@');
					res.push(locale + 'mmmmm' + separator + 'yy;@');
					res.push(locale + 'yy' + separator + 'mmm;@');
					res.push(locale + 'd' + separator + 'mmm' + separator + 'yyyy;@');
					res.push(locale + 'yyyy' + separator + 'mmm' + separator + 'd;@');
					res.push(locale + 'yy' + separator + 'mmm' + separator + 'd;@');
					res.push('yy' + separator + 'm' + separator + 'd;@');
					res.push('yy' + separator + 'mm' + separator + 'dd;@');
					res.push('yyyy' + separator + 'm' + separator + 'd;@');
					res.push('yyyy' + separator + 'mm' + separator + 'dd;@');
				}
			} else if (Asc.c_oAscNumFormatType.Time === info.type) {
				if (info.symbol === g_oDefaultCultureInfo.LCID) {
					if (AscCommon.is12HourTimeFormat(cultureInfo)) {
						res.push('[$-F400]h:mm:ss AM/PM');
					} else {
						res.push('[$-F400]h:mm:ss');
					}
				}
				if (c_oAscTimeFormatExcel[info.symbol]) {
					res = res.concat(c_oAscTimeFormatExcel[info.symbol]);
				}
				res = res.concat(['h:mm;@', 'h:mm AM/PM;@', 'h:mm:ss;@', 'h:mm:ss AM/PM;@', 'mm:ss.0;@', '[h]:mm:ss;@']);
			} else if (Asc.c_oAscNumFormatType.Percent === info.type) {
				format = '0';
				if (info.decimalPlaces > 0) {
					format += '.' + '0'.repeat(info.decimalPlaces);
				}
				format += '%';
				res.push(format);
			} else if (Asc.c_oAscNumFormatType.Fraction === info.type) {
				res = gc_aFractionFormats;
			} else if (Asc.c_oAscNumFormatType.Scientific === info.type) {
				format = '0.' + '0'.repeat(info.decimalPlaces) + 'E+00';
				res.push(format);
			} else if (Asc.c_oAscNumFormatType.Text === info.type) {
				res.push('@');
			} else if (Asc.c_oAscNumFormatType.Custom === info.type) {
				for (i = 0; i <= 4; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getCurrencyFormatSimple(null, 0, false, false, null, false));
				res.push(getCurrencyFormatSimple(null, 0, false, false, null, true));
				res.push(getCurrencyFormatSimple(null, 2, false, false, null, false));
				res.push(getCurrencyFormatSimple(null, 2, false, false, null, true));
				res.push(getCurrencyFormatSimple(null, 0, true, false, null, false));
				res.push(getCurrencyFormatSimple(null, 0, true, false, null, true));
				res.push(getCurrencyFormatSimple(null, 2, true, false, null, false));
				res.push(getCurrencyFormatSimple(null, 2, true, false, null, true));
				for (i = 9; i <= 13; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getShortDateFormat(null));
				res.push(getShortDateMonthFormat(true, true, null));
				res.push(getShortDateMonthFormat(true, false, null));
				res.push(getShortDateMonthFormat(false, true, null));
				for (i = 18; i <= 21; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(getShortDateFormat(null) + " h:mm");
				for (i = 45; i <= 49; ++i) {
					res.push(AscCommonExcel.aStandartNumFormats[i]);
				}
				res.push(AscCommon.getCurrencyFormat(null, 0, true, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 0, false, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 2, true, false, null));
				res.push(AscCommon.getCurrencyFormat(null, 2, false, false, null));
			} else {
				res.push(AscCommon.g_cGeneralFormat);
				res.push('0.00');
				res.push(interfaceFormatScientific);
				res.push(getCurrencyFormat(cultureInfo, 2, hasCurrency, true, currencySymbol));
				res.push(getCurrencyFormatSimple2(cultureInfo, 2, hasCurrency, currencySymbol, false));
				res.push(interfaceShortDateFormat);
				res.push('[$-F800]' + cultureInfo.LongDatePattern);
				//todo F400
				if (AscCommon.is12HourTimeFormat(cultureInfo)) {
					res.push('[$-F400]h:mm:ss AM/PM');
				} else {
					res.push('[$-F400]h:mm:ss');
				}
				res.push(interfaceFormatPercent);
				res.push('# ?/?');
				res.push('@');
			}
		}
		return res;
	}
	function getFormatByCulturalStandardId(id, opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		let formats;
		let localeStart = cultureInfo.Name.substring(0, 2);
		let LCID = cultureInfo.LCID;
		if ('zh' === localeStart) {
			if (4 === LCID || 2052 === LCID || 4100 === LCID || 30724 === LCID) {
				// zh
				// zh-Hans
				// zh-CN
				// zh-SG
				formats = {
					27: 'yyyy"年"m"月"',
					28: 'm"月"d"日"',
					29: 'm"月"d"日"',
					30: 'm-d-yy',
					31: 'yyyy"年"m"月"d"日"',
					32: 'h"时"mm"分"',
					33: 'h"时"mm"分"ss"秒"',
					34: '上午/下午h"时"mm"分"',
					35: '上午/下午h"时"mm"分"ss"秒"',
					36: 'yyyy"年"m"月"',
					50: 'yyyy"年"m"月"',
					51: 'm"月"d"日"',
					52: 'yyyy"年"m"月"',
					53: 'm"月"d"日"',
					54: 'm"月"d"日"',
					55: '上午/下午h"时"mm"分"',
					56: '上午/下午h"时"mm"分"ss"秒"',
					57: 'yyyy"年"m"月"',
					58: 'm"月"d"日"'
				}
			} else {
				// zh-Hant
				// zh-TW
				// zh-HK
				// zh-MO
				formats = {
					27: '[$-404]e/m/d',
					28: '[$-404]e"年"m"月"d"日"',
					29: '[$-404]e"年"m"月"d"日"',
					30: 'm/d/yy',
					31: 'yyyy"年"m"月"d"日"',
					32: 'hh"時"mm"分"',
					33: 'hh"時"mm"分"ss"秒"',
					34: '上午/下午hh"時"mm"分"',
					35: '上午/下午hh"時"mm"分"ss"秒"',
					36: '[$-404]e/m/d',
					50: '[$-404]e/m/d',
					51: '[$-404]e"年"m"月"d"日"',
					52: '上午/下午hh"時"mm"分"',
					53: '上午/下午hh"時"mm"分"ss"秒"',
					54: '上午/下午hh"時"mm"分"',
					55: '上午/下午hh"時"mm"分"ss"秒"',
					56: '[$-404]e/m/d',
					57: '[$-404]e"年"m"月"d"日"',
					58: '[$-404]e"年"m"月"d"日"'
				}
			}
		} else if ('ja' === localeStart) {
			//"ja-jp"
			formats = {
				27: '[$-411]ge.m.d',
				28: '[$-411]ggge"年"m"月"d"日"',
				29: '[$-411]ggge"年"m"月"d"日"',
				30: 'm/d/yy',
				31: 'yyyy"年"m"月"d"日"',
				32: 'h"時"mm"分"',
				33: 'h"時"mm"分"ss"秒"',
				34: 'yyyy"年"m"月"',
				35: 'm"月"d"日"',
				36: '[$-411]ge.m.d',
				50: '[$-411]ge.m.d',
				51: '[$-411]ggge"年"m"月"d"日"',
				52: 'yyyy"年"m"月"',
				53: 'm"月"d"日"',
				54: '[$-411]ggge"年"m"月"d"日"',
				55: 'yyyy"年"m"月"',
				56: 'm"月"d"日"',
				57: '[$-411]ge.m.d',
				58: '[$-411]ggge"年"m"月"d"日"'
			}
		} else if ('ko' === localeStart) {
			//"ko-kr"
			formats = {
				27: 'yyyy"年" mm"月" dd"日"',
				28: 'mm-dd',
				29: 'mm-dd',
				30: 'mm-dd-yy',
				31: 'yyyy"년" mm"월" dd"일"',
				32: 'h"시" mm"분"',
				33: 'h"시" mm"분" ss"초"',
				34: 'yyyy-mm-dd',
				35: 'yyyy-mm-dd',
				36: 'yyyy"年" mm"月" dd"日"',
				50: 'yyyy"年" mm"月" dd"日"',
				51: 'mm-dd',
				52: 'yyyy-mm-dd',
				53: 'yyyy-mm-dd',
				54: 'mm-dd',
				55: 'yyyy-mm-dd',
				56: 'yyyy-mm-dd',
				57: 'yyyy"年" mm"月" dd"日"',
				58: 'mm-dd'
			}
		} else if ('th' === localeStart) {
			//"th-th"
			formats = {
				59: 't0',
					60: 't0.00',
					61: 't#,##0',
					62: 't#,##0.00',
					67: 't0%',
					68: 't0.00%',
					69: 't# ?/?',
					70: 't# ??/??',
					71: 'ว/ด/ปปปป',
					72: 'ว-ดดด-ปป',
					73: 'ว-ดดด',
					74: 'ดดด-ปป',
					75: 'ช:นน',
					76: 'ช:นน:ทท',
					77: 'ว/ด/ปปปป ช:นน',
					78: 'นน:ทท',
					79: '[ช]:นน:ทท',
					80: '80 นน:ทท.0',
					81: 'd/m/bb'
			}
		}
		return formats && formats[id] || null;
	}
	function getFormatByStandardId(id, opt_cultureInfo) {
		var res = getFormatByCulturalStandardId(id, opt_cultureInfo);
		if (res) {
			return res;
		}
		if (59 <= id && id <= 78) {
			if (69 <= id && id <= 71) {
				id += 1;
			}
			id -= 58;
		} else if (79 <= id && id <= 81) {
			id -= 34;
		}
			//todo currencyLocale true/false?
			var currencyLocale = true;
			switch (id) {
				case 5:
					res = AscCommon.getCurrencyFormatSimple(null, 0, true, currencyLocale, null, false);
					break;
				case 6:
					res = AscCommon.getCurrencyFormatSimple(null, 0, true, currencyLocale, null, true);
					break;
				case 7:
					res = AscCommon.getCurrencyFormatSimple(null, 2, true, currencyLocale, null, false);
					break;
				case 8:
					res = AscCommon.getCurrencyFormatSimple(null, 2, true, currencyLocale, null, true);
					break;
				case 14:
					res = AscCommon.getShortDateFormat(null);
					break;
			case 15:
				res = AscCommon.getShortDateMonthFormat(true, true, null);
				break;
			case 16:
				res = AscCommon.getShortDateMonthFormat(true, false, null);
				break;
			case 17:
				res = AscCommon.getShortDateMonthFormat(false, true, null);
				break;
				case 22:
					res = AscCommon.getShortDateFormat(null) + " h:mm";
					break;
			case 23:
			case 24:
			case 25:
			case 26:
				//like 0
				res = "General";
				break;
				case 27:
				case 28:
				case 29:
				case 30:
				case 31:
				//like 14
				res = AscCommon.getShortDateFormat(null);
				break;
			case 32:
			case 33:
			case 34:
			case 35:
				//like 21
				res = AscCommonExcel.aStandartNumFormats[21];
				break;
				case 36:
				//like 14
					res = AscCommon.getShortDateFormat(null);
					break;
				case 37:
					res = AscCommon.getCurrencyFormatSimple(null, 0, false, currencyLocale, null, false);
					break;
				case 38:
					res = AscCommon.getCurrencyFormatSimple(null, 0, false, currencyLocale, null, true);
					break;
				case 39:
					res = AscCommon.getCurrencyFormatSimple(null, 2, false, currencyLocale, null, false);
					break;
				case 40:
					res = AscCommon.getCurrencyFormatSimple(null, 2, false, currencyLocale, null, true);
					break;
				case 41:
					res = AscCommon.getCurrencyFormat(null, 0, false, currencyLocale, null);
					break;
				case 42:
					res = AscCommon.getCurrencyFormat(null, 0, true, currencyLocale, null);
					break;
				case 43:
					res = AscCommon.getCurrencyFormat(null, 2, false, currencyLocale, null);
					break;
				case 44:
					res = AscCommon.getCurrencyFormat(null, 2, true, currencyLocale, null);
					break;
			case 50:
			case 51:
			case 52:
			case 53:
			case 54:
			case 55:
			case 56:
			case 57:
			case 58:
				//like 14
				res = AscCommon.getShortDateFormat(null);
				break;
                default:
                    res = AscCommonExcel.aStandartNumFormats[id];
                    break;
			}
		return res;
	}
	function canGetFormatByStandardId(id) {
		return (5 <= id && id <= 8) || (14 <= id && id <= 17) || 22 == id || (27 <= id && id <= 81);
	}
	function is12HourTimeFormat(opt_cultureInfo) {
		var cultureInfo = opt_cultureInfo ? opt_cultureInfo : g_oDefaultCultureInfo;
		return cultureInfo.UseAMPM > 0;
	}

// Import locale data from NumFormatData.js
var g_aCultureInfos = AscCommon.g_aCultureInfos;
var g_aAdditionalCurrencySymbols = AscCommon.g_aAdditionalCurrencySymbols;
var c_oAscDateFormatExcel = AscCommon.c_oAscDateFormatExcel;

var g_oDefaultCultureInfo, g_oLCID;
setCurrentCultureInfo(1033);//en-US//1033//fr-FR//1036//basq//1069//ru-Ru//1049//hindi//1081

function getGannenFormatCodes(format) {
	var sLowerFormat = format ? format.toLowerCase() : "";
	if (!sLowerFormat
		|| (sLowerFormat.indexOf("x-gannen") === -1 && sLowerFormat.indexOf("87f70000") === -1)) {
		return null;
	}

	var cf = new CellFormat(format);
	var canonical = cf.toString();
	if (canonical.toLowerCase().indexOf("x-gannen") === -1) {
		return null;
	}

	var fallback = cf.toString(undefined, undefined, {gannenFallback: true});
	return fallback !== canonical ? {fallback: fallback, formatCode16: canonical} : null;
}

	//---------------------------------------------------------export---------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].isNumber = isNumber;
    window["AscCommon"].NumFormat = NumFormat;
    window["AscCommon"].CellFormat = CellFormat;
    window["AscCommon"].getGannenFormatCodes = getGannenFormatCodes;
    window["AscCommon"].DecodeGeneralFormat = DecodeGeneralFormat;
    window["AscCommon"].setCurrentCultureInfo = setCurrentCultureInfo;
	window["AscCommon"].checkCultureInfoFontPicker = checkCultureInfoFontPicker;
	window['AscCommon'].getShortDateFormat = getShortDateFormat;
	window['AscCommon'].getShortDateFormat2 = getShortDateFormat2;
	window['AscCommon'].getShortTimeFormat = getShortTimeFormat;
	window['AscCommon'].getLongTimeFormat = getLongTimeFormat;
	window['AscCommon'].getShortDateMonthFormat = getShortDateMonthFormat;
	window['AscCommon'].getNumberFormatSimple = getNumberFormatSimple;
	window['AscCommon'].getNumberFormat = getNumberFormat;
	window['AscCommon'].getLocaleFormat = getLocaleFormat;
	window['AscCommon'].getCurrencyFormatSimple = getCurrencyFormatSimple;
	window['AscCommon'].getCurrencyFormatSimple2 = getCurrencyFormatSimple2;
	window['AscCommon'].getCurrencyFormat = getCurrencyFormat;
	window['AscCommon'].getFormatCells = getFormatCells;
	window['AscCommon'].canGetFormatByStandardId = canGetFormatByStandardId;
	window['AscCommon'].getFormatByStandardId = getFormatByStandardId;
	window['AscCommon'].is12HourTimeFormat = is12HourTimeFormat;
	window['AscCommon'].compareNumbers = compareNumbers;

    window["AscCommon"].gc_nMaxDigCount = gc_nMaxDigCount;
    window["AscCommon"].gc_nMaxDigCountView = gc_nMaxDigCountView;
    window["AscCommon"].oNumFormatCache = oNumFormatCache;
    window["AscCommon"].oGeneralEditFormatCache = oGeneralEditFormatCache;
    window["AscCommon"].g_oFormatParser = g_oFormatParser;
    window["AscCommon"].g_aCultureInfos = g_aCultureInfos;
    window["AscCommon"].g_oDefaultCultureInfo = g_oDefaultCultureInfo;
	window["AscCommon"].g_aAdditionalCurrencySymbols = g_aAdditionalCurrencySymbols;
	window["AscCommon"].NumFormatType = NumFormatType;

	window["AscCommon"].escapeRegExp = escapeRegExp;


})(window);
