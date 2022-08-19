/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window)
{
	function IsEscapingChar(unicode)
	{
		return (94 === unicode); // ^
	}
	function IsDigit(unicode)
	{
		return (57 === unicode); // 9
	}
	function IsLetter(unicode)
	{
		return (97 === unicode); // a
	}
	function IsDigitOrLetter(unicode)
	{
		return (42 === unicode); // *
	}

	function CTextItem(unicode)
	{
		this.Value = unicode;
	}
	CTextItem.prototype.Check = function(unicode)
	{
		return (this.Value === unicode);
	}

	function CDigitItem()
	{
	}
	CDigitItem.prototype.Check = function(unicode)
	{
		return AscCommon.IsDigit(unicode);
	}

	function CLetterItem()
	{
	}
	CLetterItem.prototype.Check = function(unicode)
	{
		return AscCommon.IsLetter(unicode);
	}

	function CDigitOrLetterItem()
	{
	}
	CDigitOrLetterItem.prototype.Check = function(unicode)
	{
		return (AscCommon.IsDigit(unicode) || AscCommon.IsLetter(unicode));
	}

	/**
	 * Класс представляющий маску для текстовой формы
	 * @constructor
	 */
	function CTextFormMask()
	{
		this.Mask    = "";
		this.Pattern = [];

		this.Parse();
	}
	CTextFormMask.prototype.Set = function(sMask)
	{
		this.Mask = sMask;
		this.Parse();
	};
	CTextFormMask.prototype.Get = function()
	{
		return this.Mask;
	};
	CTextFormMask.prototype.Check = function(arrBuffer)
	{
		for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
		{
			if (nIndex >= this.Pattern.length || !this.Pattern[nIndex].Check(arrBuffer[nIndex]))
				return false;
		}

		return true;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CTextFormMask.prototype.Parse = function()
	{
		this.Pattern = [];
		for (let iterator = this.Mask.getUnicodeIterator(); iterator.check(); iterator.next())
		{
			let unicode = iterator.value();
			if (IsEscapingChar(unicode))
			{
				iterator.next();
				if (!iterator.check())
					break;

				this.Pattern.push(new CTextItem(iterator.value()));
			}
			else if (IsLetter(unicode))
			{
				this.Pattern.push(new CLetterItem());
			}
			else if (IsDigit(unicode))
			{
				this.Pattern.push(new CDigitItem());
			}
			else if (IsDigitOrLetter(unicode))
			{
				this.Pattern.push(new CDigitOrLetterItem());
			}
			else
			{
				this.Pattern.push(new CTextItem(unicode));
			}
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTextFormMask = CTextFormMask;

})(window);
