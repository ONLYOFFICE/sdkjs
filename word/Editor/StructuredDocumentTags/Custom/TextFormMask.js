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

(function(window)
{
	function IsEscapingChar(unicode)
	{
		return (92 === unicode); // \
	}
	function IsDigit(unicode)
	{
		return (57 === unicode); // 9
	}
	function IsLetter(unicode)
	{
		return (97 === unicode || 65 === unicode); // a A
	}
	function IsDigitOrLetter(unicode)
	{
		return (79 === unicode); // O
	}
	function IsAnyCharacter(unicode)
	{
		return (88 === unicode); // X
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

	function CAnyCharacter()
	{
	}
	CAnyCharacter.prototype.Check = function(unicode)
	{
		return true;
	}

	function BufferIterator (buffer)
	{
		this.buffer = buffer;
		this.intCursor = 0;
		this.intContent = this.GetNext();
		this.arrOutputContent = [];
	}
	BufferIterator.prototype.GetNext = function ()
	{
		if (this.intCursor > this.buffer.length)
			return false;

		let oContent = this.buffer[this.intCursor];
		this.intCursor++;

		if (undefined !== oContent) {
			return oContent;
		}

		return false;
	};
	BufferIterator.prototype.CheckRule = function (oRule)
	{
		if (undefined === oRule)
			return false;

		if (oRule.Check(this.intContent))
		{
			this.arrOutputContent.push(this.intContent);
			this.intContent = this.GetNext();
		}
		else
		{
			if (oRule.Value)
				this.arrOutputContent.push(oRule.Value);
			else
				return false;
		}

		return true;
	};
	BufferIterator.prototype.IsReturnContent = function (intLengthPatternArr)
	{
		return this.intCursor - 1 === this.buffer.length ||
			this.arrOutputContent.length === intLengthPatternArr
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
	CTextFormMask.prototype.GetLength = function()
	{
		return this.Pattern.length;
	};
	CTextFormMask.prototype.Check = function(arrBuffer, isFullCheck)
	{
		if (isFullCheck && arrBuffer.length !== this.Pattern.length)
			return false;

		for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
		{
			if (nIndex >= this.Pattern.length || !this.Pattern[nIndex].Check(arrBuffer[nIndex]))
				return false;
		}

		return true;
	};
	CTextFormMask.prototype.Correct = function(text)
	{
		let buffer;
		if (typeof(text) === "string")
			buffer = text.codePointsArray();
		else if (Array.isArray(text))
			buffer = Array.from(text);

		buffer = this.CorrectBuffer(buffer);

		if (!buffer)
			return text;

		if (typeof(text) === "string")
		{
			let sResult = "";
			for (let index = 0, count = buffer.length; index < count; ++index)
			{
				sResult += String.fromCodePoint(buffer[index]);
			}
			return sResult;
		}

		return buffer;
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
			else if (IsAnyCharacter(unicode))
			{
				this.Pattern.push(new CAnyCharacter());
			}
			else
			{
				this.Pattern.push(new CTextItem(unicode));
			}
		}
	};
	CTextFormMask.prototype.CorrectBuffer = function(buffer)
	{
		if (!this.Pattern.length  || !buffer || !buffer.length)
			return buffer;

		let oBufferIterator = new BufferIterator(buffer, this.Pattern.length);

		for (let i = 0, isContinue = true; i < this.Pattern.length && isContinue; i++)
		{
			isContinue = oBufferIterator.CheckRule(this.Pattern[i]);

			if (!isContinue)
				break;
		}

		if (oBufferIterator.IsReturnContent(this.Pattern.length))
			return Array.from(oBufferIterator.arrOutputContent);

		return false
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTextFormMask = CTextFormMask;

})(window);
