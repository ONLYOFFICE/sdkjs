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
	/**
	 * Паттерн поиска
	 * @constructor
	 */
	function CSearchPatternEngine()
	{
		this.Elements = [];
	}
	CSearchPatternEngine.prototype.Set = function(sString)
	{
		this.Elements = [];
		for (var oIterator = sString.getUnicodeIterator(); oIterator.check(); oIterator.next())
		{
			var nCharCode = oIterator.value();

			if (0x005E === nCharCode)
			{
				oIterator.next();
				if (!oIterator.check())
				{
					this.Elements.push(new AscCommonWord.CSearchTextItemChar(nCharCode));
					break;
				}

				var nNextCharCode = oIterator.value();
				var oSpecialElement = this.private_GetSpecialElement(nNextCharCode);
				if (oSpecialElement)
				{
					this.Elements.push(oSpecialElement);
				}
				else
				{
					this.Elements.push(new AscCommonWord.CSearchTextItemChar(nCharCode));
					this.Elements.push(new AscCommonWord.CSearchTextItemChar(nNextCharCode));
				}
			}
			else
			{
				this.Elements.push(new AscCommonWord.CSearchTextItemChar(nCharCode));
			}
		}
	};
	CSearchPatternEngine.prototype.private_GetSpecialElement = function(nCharCode)
	{
		switch(nCharCode)
		{
			case 0x006C: return new AscCommonWord.CSearchTextSpecialLineBreak();	      // ^l - new line
			case 0x0074: return new AscCommonWord.CSearchTextSpecialTab();                // ^t - tab
			case 0x0070: return new AscCommonWord.CSearchTextSpecialParaEnd();            // ^p - paraEnd
			case 0x003F: return new AscCommonWord.CSearchTextSpecialAnySymbol();          // ^? - any symbol
			case 0x0023: return new AscCommonWord.CSearchTextSpecialAnyDigit();           // ^# - any digit
			case 0x0024: return new AscCommonWord.CSearchTextSpecialAnyLetter();          // ^$ - any letter
			case 0x006E: return new AscCommonWord.CSearchTextSpecialColumnBreak();        // ^n - column Break
			case 0x0065: return new AscCommonWord.CSearchTextSpecialEndnoteMark();        // ^e - endnote mark
			case 0x0064: return new AscCommonWord.CSearchTextSpecialField();              // ^d - field
			case 0x0066: return new AscCommonWord.CSearchTextSpecialFootnoteMark();       // ^f - footnote mark
			case 0x0067: return new AscCommonWord.CSearchTextSpecialGraphicObject();      // ^g - graphic object
			case 0x006D: return new AscCommonWord.CSearchTextSpecialPageBreak();          // ^m - break page
			case 0x007E: return new AscCommonWord.CSearchTextSpecialNonBreakingHyphen();  // ^~ - nonbreaking hyphen
			case 0x0073: return new AscCommonWord.CSearchTextSpecialNonBreakingSpace();   // ^s - nonbreaking space
			case 0x005E: return new AscCommonWord.CSearchTextItemChar(0x5E);              // ^^ - caret character
			case 0x0077: return new AscCommonWord.CSearchTextSpecialAnySpace();           // ^w - any space
			case 0x002B: return new AscCommonWord.CSearchTextSpecialEmDash();             // ^+ - em dash
			case 0x003D: return new AscCommonWord.CSearchTextSpecialEnDash();             // ^= - en dash
			case 0x0025: return new AscCommonWord.CSearchTextSpecialSectionCharacter();   // ^% - § Section Character
			case 0x0076: return new AscCommonWord.CSearchTextSpecialParagraphCharacter(); // ^v - ¶ Paragraph Character
			case 0x0079: return new AscCommonWord.CSearchTextSpecialAnyDash();            // ^y - any dash
		}

		return null;
	};
	CSearchPatternEngine.prototype.Get = function(nIndex)
	{
		return this.Elements[nIndex];
	};
	CSearchPatternEngine.prototype.GetLength = function()
	{
		return this.Elements.length;
	};
	CSearchPatternEngine.prototype.Check = function(nPos, oRunItem, oEngine)
	{
		var oSearchElement = oRunItem.ToSearchElement(oEngine);

		if (!oSearchElement)
			return false;

		return this.Elements[nPos].IsMatch(oSearchElement);
	};
	CSearchPatternEngine.prototype.GetErrorForReplaceString = function(sString)
	{
		for (var oIterator = sString.getUnicodeIterator(); oIterator.check(); oIterator.next())
		{
			var nCharCode = oIterator.value();

			if (0x005E === nCharCode)
			{
				oIterator.next();
				if (!oIterator.check())
					break;

				var nNextCharCode   = oIterator.value();
				var oSpecialElement = this.private_GetSpecialElement(nNextCharCode);
				if (oSpecialElement && !oSpecialElement.ToRunElement(false))
					return String.fromCodePoint(0x005E, nNextCharCode);
			}
		}

		return null;
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CSearchPatternEngine = CSearchPatternEngine;

})(window);
