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
	const FormatType = {
		None   : 0,
		Digit  : 1,
		Letter : 2,
		Mask   : 3,
		RegExp : 4
	};

	/**
	 * Базовый класс для всех форматов
	 * @constructor
	 */
	function CTextFormFormat()
	{
		this.BaseFormat = FormatType.None;
		this.Symbols    = []; // Специальный параметр для возможного ограничения на ввод символов

		this.Mask   = new AscWord.CTextFormMask();
		this.RegExp = "";
	}
	CTextFormFormat.prototype.SetSymbols = function(value)
	{
		this.Symbols = [];

		if (Array.isArray(value))
		{
			this.Symbols = Array.from(value);
		}
		else if (typeof(value) === "string")
		{
			for (let oIter = value.getUnicodeIterator(); oIter.check(); oIter.next())
			{
				this.Symbols.push(oIter.value());
			}
		}
	};
	CTextFormFormat.prototype.GetSymbols = function(isToString)
	{
		if (isToString)
		{
			// wait for support by IE
			//return String.fromCodePoint(...this.Symbols);
			let sResult = "";
			for (let nIndex = 0, nCount = this.Symbols.length; nIndex < nCount; ++nIndex)
			{
				sResult += String.fromCodePoint(this.Symbols[nIndex]);
			}
			return sResult;
		}

		return this.Symbols;
	};
	CTextFormFormat.prototype.SetNone = function()
	{
		this.BaseFormat = FormatType.None;
	};
	CTextFormFormat.prototype.GetType = function()
	{
		return this.BaseFormat;
	};
	CTextFormFormat.prototype.SetDigit = function()
	{
		this.BaseFormat = FormatType.Digit;
	};
	CTextFormFormat.prototype.SetLetter = function()
	{
		this.BaseFormat = FormatType.Letter;
	};
	/**
	 * Выствляем маску. Маску используем как в Adobe:
	 * 9 - число, a - текстовый символ, * - либо число, либо текстовый символ
	 * @param sMask
	 */
	CTextFormFormat.prototype.SetMask = function(sMask)
	{
		this.BaseFormat = FormatType.Mask;
		this.Mask.Set(sMask);
	};
	CTextFormFormat.prototype.SetRegExp = function(sRegExp)
	{
		this.BaseFormat = FormatType.RegExp;
		this.RegExp     = sRegExp;
	};
	CTextFormFormat.prototype.CheckFormat = function(sText)
	{
		switch (this.BaseFormat)
		{
			case FormatType.Digit:
				return this.CheckDigit(sText);
			case FormatType.Letter:
				return this.CheckLetter(sText);
			case FormatType.Mask:
				return this.CheckMask(sText);
			case FormatType.RegExp:
				return this.CheckRegExp(sText);
		}

		return true;
	};
	CTextFormFormat.prototype.CheckSymbols = function(sText)
	{
		if (sText && this.Symbols.length)
		{
			for (let oIter = sText.getUnicodeIterator(); oIter.check(); oIter.next())
			{
				if (-1 === this.Symbols.indexOf(oIter.value()))
					return false;
			}
		}

		return true;
	};
	CTextFormFormat.prototype.Check = function(sText)
	{
		return (this.CheckFormat(sText) && this.CheckSymbols(sText));
	};
	CTextFormFormat.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteLong(this.BaseFormat);
		oWriter.WriteString2(this.GetSymbols(true));
		oWriter.WriteString2(this.Mask.Get());
		oWriter.WriteString2(this.RegExp);
	};
	CTextFormFormat.prototype.ReadFromBinary = function(oReader)
	{
		this.BaseFormat = oReader.GetLong();
		this.SetSymbols(oReader.GetString2());
		this.Mask.Set(oReader.GetString2());
		this.RegExp = oReader.GetString2();
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CTextFormFormat.prototype.CheckDigit = function(sText)
	{
		for (let oIter = sText.getUnicodeIterator(); oIter.check(); oIter.next())
		{
			if (!AscCommon.IsDigit(oIter.value()))
				return false;
		}
		return true;
	};
	CTextFormFormat.prototype.CheckLetter = function(sText)
	{
		for (let oIter = sText.getUnicodeIterator(); oIter.check(); oIter.next())
		{
			if (!AscCommon.IsLetter(oIter.value()))
				return false;
		}
		return true;
	};
	CTextFormFormat.prototype.CheckMask = function(sText)
	{
		return this.Mask.Check(sText);
	};
	CTextFormFormat.prototype.CheckRegExp = function(sText)
	{
		return (!!sText.match(this.RegExp));
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTextFormFormat = CTextFormFormat;

	let exportPrototype       = window['Asc']['TextFormFormatType'] = window['Asc'].TextFormFormatType = FormatType;
	exportPrototype['None']   = exportPrototype.None;
	exportPrototype['Digit']  = exportPrototype.Digit;
	exportPrototype['Letter'] = exportPrototype.Letter;
	exportPrototype['Mask']   = exportPrototype.Mask;
	exportPrototype['RegExp'] = exportPrototype.RegExp;

})(window);
