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

(function (window)
{
	/**
	 * @constructor
	 */
	function CAutoCorrectSettings()
	{
		this.SmartQuotes            = true;
		this.HyphensWithDash        = true;
		this.AutomaticBulletedLists = true;
		this.AutomaticNumberedLists = true;
		this.FrenchPunctuation      = true;
		this.DoubleSpaceWithPeriod  = false;
		this.FirstLetterOfSentences = true;
		this.FirstLetterOfCells     = true;
		this.Hyperlinks             = true;
		this.FirstLetterExceptions  = {};
		this.FirstLetterExcMaxLen   = 0;
		this.InitDefaultFirstLetterAutoCorrectExceptions();
	}
	//getters
	CAutoCorrectSettings.prototype.IsSmartQuotes  = function()
	{
		return this.SmartQuotes;
	};
	CAutoCorrectSettings.prototype.IsHyphensWithDash = function()
	{
		return this.HyphensWithDash;
	};
	CAutoCorrectSettings.prototype.IsAutomaticBulletedLists = function()
	{
		return this.AutomaticBulletedLists;
	};
	CAutoCorrectSettings.prototype.IsAutomaticNumberedLists = function()
	{
		return this.AutomaticNumberedLists;
	};
	CAutoCorrectSettings.prototype.IsFrenchPunctuation = function()
	{
		return this.FrenchPunctuation;
	};
	CAutoCorrectSettings.prototype.IsDoubleSpaceWithPeriod = function()
	{
		return this.DoubleSpaceWithPeriod;
	};
	CAutoCorrectSettings.prototype.IsFirstLetterOfSentences = function()
	{
		return this.FirstLetterOfSentences;
	};
	CAutoCorrectSettings.prototype.IsFirstLetterOfCells = function()
	{
		return this.FirstLetterOfCells;
	};
	CAutoCorrectSettings.prototype.IsHyperlinks = function()
	{
		return this.Hyperlinks;
	};
	//setters
	CAutoCorrectSettings.prototype.SetSmartQuotes  = function(bVal)
	{
		this.SmartQuotes = bVal;
	};
	CAutoCorrectSettings.prototype.SetHyphensWithDash = function(bVal)
	{
		this.HyphensWithDash = bVal;
	};
	CAutoCorrectSettings.prototype.SetAutomaticBulletedLists = function(bVal)
	{
		this.AutomaticBulletedLists = bVal;
	};
	CAutoCorrectSettings.prototype.SetAutomaticNumberedLists = function(bVal)
	{
		this.AutomaticNumberedLists = bVal;
	};
	CAutoCorrectSettings.prototype.SetFrenchPunctuation = function(bVal)
	{
		this.FrenchPunctuation = bVal;
	};
	CAutoCorrectSettings.prototype.SetDoubleSpaceWithPeriod = function(bVal)
	{
		this.DoubleSpaceWithPeriod = bVal;
	};
	CAutoCorrectSettings.prototype.SetFirstLetterOfSentences = function(bVal)
	{
		this.FirstLetterOfSentences = bVal;
	};
	CAutoCorrectSettings.prototype.SetFirstLetterOfCells = function(bVal)
	{
		this.FirstLetterOfCells = bVal;
	};
	CAutoCorrectSettings.prototype.SetHyperlinks = function(bVal)
	{
		this.Hyperlinks = bVal;
	};
	CAutoCorrectSettings.prototype.SetFirstLetterAutoCorrectExceptions = function(arrExceptions)
	{
		this.FirstLetterExceptions = {};
		var nMaxLen = 0;
		for (var nIndex = 0, nCount = arrExceptions.length; nIndex < nCount; ++nIndex)
		{
			if (!arrExceptions[nIndex].length)
				continue;

			if (arrExceptions[nIndex].length > nMaxLen)
				nMaxLen = arrExceptions[nIndex].length;

			var nChar = arrExceptions[nIndex].charAt(0);

			if (!this.FirstLetterExceptions[nChar])
				this.FirstLetterExceptions[nChar] = [];

			this.FirstLetterExceptions[nChar].push(arrExceptions[nIndex]);
		}

		this.FirstLetterExcMaxLen = nMaxLen;
	};
	CAutoCorrectSettings.prototype.InitDefaultFirstLetterAutoCorrectExceptions = function()
	{
		// Init default for Latin and Cyrillic
		this.SetFirstLetterAutoCorrectExceptions([
			"a", "abbr", "abs", "acct", "addn", "adj", "advt", "al", "alt", "amt", "anon", "approx", "appt", "apr", "apt", "assn", "assoc", "asst", "attn", "attrib", "aug", "aux", "ave", "avg",
			"b", "bal", "bldg", "blvd", "bot", "bro", "bros",
			"c", "ca", "calc", "cc", "cert", "certif", "cf", "cit", "cm", "co", "comp", "conf", "confed", "const", "cont", "contrib", "coop", "corp", "ct",
			"d", "dbl", "dec", "decl", "def", "defn", "dept", "deriv", "diag", "diff", "div", "dm", "dr", "dup", "dupl",
			"e", "encl", "eq", "eqn", "equip", "equiv", "esp", "esq", "est", "etc", "excl", "ext",
			"f", "feb", "ff", "fig", "freq", "fri", "ft", "fwd",
			"g", "gal", "gen", "gov", "govt",
			"h", "hdqrs", "hgt", "hist", "hosp", "hq", "hr", "hrs", "ht", "hwy",
			"i", "ib", "ibid", "illus", "in", "inc", "incl", "incr", "int", "intl", "irreg", "ital",
			"j", "jan", "jct", "jr", "jul", "jun",
			"k", "kg", "km", "kmh",
			"l", "lang", "lb", "lbs", "lg", "lit", "ln", "lt",
			"m", "mar", "masc", "max", "mfg", "mg", "mgmt", "mgr", "mgt", "mhz", "mi", "min", "misc", "mkt", "mktg", "ml", "mm", "mngr", "mon", "mph", "mr", "mrs", "msec", "msg", "mt", "mtg", "mtn", "mun",
			"n", "na", "name", "nat", "natl", "ne", "neg", "ng", "no", "norm", "nos", "nov", "num", "nw",
			"o", "obj", "occas", "oct", "op", "opt", "ord", "org", "orig", "oz",
			"p", "pa", "pg", "pkg", "pl", "pls", "pos", "pp", "ppt", "pred", "pref", "prepd", "prev", "priv", "prof", "proj", "pseud", "psi", "pt", "publ",
			"q", "qlty", "qt", "qty",
			"r", "rd", "re", "rec", "ref", "reg", "rel", "rep", "req", "reqd", "resp", "rev",
			"s", "sat", "sci", "se", "sec", "sect", "sep", "sept", "seq", "sig", "soln", "soph", "spec", "specif", "sq", "sr", "st", "sta", "stat", "std", "subj", "subst", "sun", "supvr", "sw",
			"t", "tbs", "tbsp", "tech", "tel", "temp", "thur", "thurs", "tkt", "tot", "transf", "transl", "tsp", "tues",
			"u", "univ", "util",
			"v", "var", "veg", "vert", "viz", "vol", "vs",
			"w", "wed", "wk", "wkly", "wt",
			"x",
			"y", "yd", "yr",
			"z",

			"а",
			"б",
			"вв",
			"гг", "гл",
			"д", "др",
			"е", "ед",
			"ё",
			"ж",
			"з",
			"и",
			"й",
			"к", "кв", "кл", "коп", "куб",
			"лл",
			"м", "мл", "млн", "млрд",
			"н", "наб", "нач",
			"о", "обл", "обр", "ок",
			"п", "пер", "пл", "пос", "пр",
			"руб",
			"сб", "св", "см", "соч", "ср", "ст", "стр",
			"тт", "тыс",
			"у",
			"ф",
			"х",
			"ц",
			"ш", "шт",
			"щ",
			"ъ",
			"ы",
			"ь",
			"э", "экз",
			"ю"]
		);
	};
	CAutoCorrectSettings.prototype.GetFirstLetterAutoCorrectExceptions = function()
	{
		var arrResult = [];
		for (var nChar in this.FirstLetterExceptions)
		{
			arrResult = arrResult.concat(this.FirstLetterExceptions[nChar]);
		}
		return arrResult;
	};
	CAutoCorrectSettings.prototype.CheckFirstLetterAutoCorrectException = function(sWord)
	{
		var _sWord = sWord.toLowerCase();

		var nChar = _sWord.charAt(0);
		if (!this.FirstLetterExceptions[nChar])
			return false;

		var arrExceptions = this.FirstLetterExceptions[nChar];
		for (var nIndex = 0, nCount = arrExceptions.length; nIndex < nCount; ++nIndex)
		{
			if (_sWord === arrExceptions[nIndex])
				return true;
		}

		return false;
	};
	CAutoCorrectSettings.prototype.GetFirstLetterAutoCorrectExceptionsMaxLen = function()
	{
		return this.FirstLetterExcMaxLen;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CAutoCorrectSettings    = CAutoCorrectSettings;

})(window);
