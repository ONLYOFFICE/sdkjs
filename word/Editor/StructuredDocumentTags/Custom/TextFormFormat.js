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

		this.FullCheck = false;
	}
	CTextFormFormat.prototype.Copy = function()
	{
		let format = new CTextFormFormat();
		format.BaseFormat = this.BaseFormat;
		format.Symbols    = this.Symbols.slice();
		format.RegExp     = this.RegExp;
		format.Mask.Set(this.Mask.Get());
		return format;
	};
	CTextFormFormat.prototype.IsEqual = function(format)
	{
		if (!format)
			return false;
		
		return (this.BaseFormat === format.BaseFormat
			&& (FormatType.RegExp !== this.BaseFormat || format.RegExp === this.RegExp)
			&& (FormatType.Mask !== this.BaseFormat || format.Mask.Get() === this.Mask.Get()));
	};
	CTextFormFormat.prototype.IsEmpty = function()
	{
		return (FormatType.None === this.BaseFormat && !this.Symbols.length);
	};
	CTextFormFormat.prototype.SetSymbols = function(value)
	{
		this.Symbols = [];

		if (Array.isArray(value))
		{
			this.Symbols = value.slice();
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
	CTextFormFormat.prototype.SetType = function(type, opt_val)
	{
		switch (type)
		{
			case Asc.TextFormFormatType.Digit: this.SetDigit(); break;
			case Asc.TextFormFormatType.Letter: this.SetLetter(); break;
			case Asc.TextFormFormatType.Mask: this.SetMask(opt_val); break;
			case Asc.TextFormFormatType.RegExp: this.SetRegExp(opt_val); break;
			case Asc.TextFormFormatType.None:
			default: this.SetNone(); break;
		}
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
	CTextFormFormat.prototype.GetMask = function()
	{
		return this.Mask.Get();
	};
	CTextFormFormat.prototype.IsMask = function()
	{
		return (this.BaseFormat === FormatType.Mask);
	};
	CTextFormFormat.prototype.GetMaskLength = function()
	{
		return this.Mask.GetLength();
	};
	CTextFormFormat.prototype.SetRegExp = function(sRegExp)
	{
		this.BaseFormat = FormatType.RegExp;
		this.RegExp     = sRegExp;
	};
	CTextFormFormat.prototype.GetRegExp = function()
	{
		return this.RegExp;
	};
	CTextFormFormat.prototype.IsRegExp = function()
	{
		return (this.BaseFormat === FormatType.RegExp);
	};
	CTextFormFormat.prototype.CheckFormat = function(arrBuffer)
	{
		switch (this.BaseFormat)
		{
			case FormatType.Digit:
				return this.CheckDigit(arrBuffer);
			case FormatType.Letter:
				return this.CheckLetter(arrBuffer);
			case FormatType.Mask:
				return this.CheckMask(arrBuffer);
			case FormatType.RegExp:
				return this.CheckRegExp(arrBuffer);
		}

		return true;
	};
	CTextFormFormat.prototype.CheckSymbols = function(arrBuffer)
	{
		if (arrBuffer && this.Symbols.length)
		{
			for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
			{
				if (-1 === this.Symbols.indexOf(arrBuffer[nIndex]))
					return false;
			}
		}

		return true;
	};
	CTextFormFormat.prototype.Correct = function(sText)
	{
		if (this.BaseFormat === FormatType.Mask)
			sText = this.Mask.Correct(sText);

		return sText;
	};
	CTextFormFormat.prototype.Check = function(sText, isFullCheck)
	{
		this.FullCheck = !!isFullCheck;

		let arrBuffer = this.GetBuffer(sText);
		return (this.CheckFormat(arrBuffer) && this.CheckSymbols(arrBuffer));
	};
	CTextFormFormat.prototype.CheckOnFly = function(sText)
	{
		let arrBuffer = this.GetBuffer(sText);

		if (!this.CheckSymbols(arrBuffer))
			return false;

		if (FormatType.Digit === this.BaseFormat || FormatType.Letter === this.BaseFormat)
			return this.CheckFormat(arrBuffer);

		return true;
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
	CTextFormFormat.prototype.ToJson = function()
	{
		switch (this.BaseFormat)
		{
			case FormatType.Digit:
				return {
					"type" : "digit"
				};
			case FormatType.Letter:
				return {
					"type" : "letter"
				};
			case FormatType.Mask:
				return {
					"type"  : "mask",
					"value" : this.Mask.Get()
				};
			case FormatType.RegExp:
				return {
					"type"  : "regExp",
					"value" : this.RegExp
				};
		}

		return {
			"type" : "none"
		};
	};
	CTextFormFormat.prototype.FromJson = function(json)
	{
		this.SetNone();

		if (!json || !json["type"])
			return;

		let sType = json["type"];
		switch (sType)
		{
			case "digit":
				return this.SetDigit();
			case "letter":
				return this.SetLetter();
			case "mask":
				return this.SetMask(json["value"]);
			case "regExp":
				return this.SetRegExp(json["value"]);
		}
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CTextFormFormat.prototype.CheckDigit = function(arrBuffer)
	{
		for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
		{
			if (!AscCommon.IsDigit(arrBuffer[nIndex]))
				return false;
		}
		return true;
	};
	CTextFormFormat.prototype.CheckLetter = function(arrBuffer)
	{
		for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
		{
			if (!AscCommon.IsLetter(arrBuffer[nIndex]))
				return false;
		}
		return true;
	};
	CTextFormFormat.prototype.CheckMask = function(arrBuffer)
	{
		return this.Mask.Check(arrBuffer, this.FullCheck);
	};
	CTextFormFormat.prototype.CheckRegExp = function(arrBuffer)
	{
		let sText = "";
		for (let nIndex = 0, nCount = arrBuffer.length; nIndex < nCount; ++nIndex)
		{
			sText += String.fromCodePoint(arrBuffer[nIndex]);
		}

		return (!!sText.match(this.RegExp));
	};
	CTextFormFormat.prototype.GetBuffer = function(sText)
	{
		let arrBuffer = [];
		if (Array.isArray(sText))
		{
			arrBuffer = sText.slice();
		}
		else if (typeof(sText) === "string")
		{
			for (let oIter = sText.getUnicodeIterator(); oIter.check(); oIter.next())
			{
				arrBuffer.push(oIter.value());
			}
		}

		return arrBuffer;
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
