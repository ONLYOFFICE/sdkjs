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
	/**
	 * Класс с настройками для даты
	 * @constructor
	 */
	function CSdtDatePickerPr()
	{
		this.FullDate   = (new Date()).toISOString().slice(0, 19) + 'Z';
		this.LangId     = 1033;
		this.DateFormat = "mm/dd/yyyy";
		this.Calendar   = Asc.c_oAscCalendarType.Gregorian;
		
		// Специальный параметр, чтобы при выставлении значения из интерфейса, но не из календаря мы помечали, что реально дата не проставлена
		this.NullFullDate = false;
	}
	CSdtDatePickerPr.prototype.Copy = function()
	{
		var oDate = new CSdtDatePickerPr();

		oDate.FullDate   = this.FullDate;
		oDate.LangId     = this.LangId;
		oDate.DateFormat = this.DateFormat;
		oDate.Calendar   = this.Calendar;
		
		oDate.NullFullDate = this.NullFullDate;

		return oDate;
	};
	CSdtDatePickerPr.prototype.IsEqual = function(oDate)
	{
		return (oDate && this.FullDate === oDate.FullDate && this.LangId === oDate.LangId && this.DateFormat === oDate.DateFormat && this.Calendar === oDate.Calendar && this.NullFullDate === oDate.NullFullDate);
	};
	CSdtDatePickerPr.prototype.ToString = function(sFormat, sFullDate, nLangId)
	{
		if (undefined === sFormat)
			sFormat = this.DateFormat;

		if (undefined === sFullDate)
			sFullDate = this.FullDate;

		if (undefined === nLangId)
			nLangId = this.LangId;

		var oFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.WordFieldDate);
		if (oFormat)
		{
			var oCultureInfo = AscCommon.g_aCultureInfos[nLangId];
			if (!oCultureInfo)
				oCultureInfo = AscCommon.g_aCultureInfos[1033];

			var oDateTime = new Asc.cDate(sFullDate);
			return oFormat.formatToChart(oDateTime.getExcelDate(true) + (oDateTime.getHours() * 60 * 60 + oDateTime.getMinutes() * 60 + oDateTime.getSeconds()) / AscCommonExcel.c_sPerDay, 15, oCultureInfo);
		}

		return sFullDate;
	};
	CSdtDatePickerPr.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteString2(this.FullDate);
		oWriter.WriteLong(this.LangId);
		oWriter.WriteString2(this.DateFormat);
		oWriter.WriteLong(this.Calendar);
	};
	CSdtDatePickerPr.prototype.ReadFromBinary = function(oReader)
	{
		this.FullDate   = oReader.GetString2();
		this.LangId     = oReader.GetLong();
		this.DateFormat = oReader.GetString2();
		this.Calendar   = oReader.GetLong();
	};
	CSdtDatePickerPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtDatePickerPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtDatePickerPr.prototype.GetFullDate = function()
	{
		if (this.IsNullFullDate())
			return null;
		
		return this.FullDate;
	};
	CSdtDatePickerPr.prototype.SetFullDate = function(fullDate)
	{
		let date;
		if (fullDate instanceof Date)
			date = fullDate;
		else if (!fullDate)
			date = new Date();
		else
			date = new Date(fullDate);
		
		if (isNaN(date))
			date = new Date();
		
		this.FullDate = date.toISOString().slice(0, 19) + 'Z';
		this.NullFullDate = false;
	};
	CSdtDatePickerPr.prototype.SetNullFullDate = function(isNull)
	{
		this.NullFullDate = isNull;
	};
	CSdtDatePickerPr.prototype.IsNullFullDate = function()
	{
		return this.NullFullDate;
	};
	CSdtDatePickerPr.prototype.GetLangId = function()
	{
		return this.LangId;
	};
	CSdtDatePickerPr.prototype.SetLangId = function(nLangId)
	{
		this.LangId = nLangId;
	};
	CSdtDatePickerPr.prototype.GetDateFormat = function()
	{
		return this.DateFormat;
	};
	CSdtDatePickerPr.prototype.SetDateFormat = function(sDateFormat)
	{
		this.DateFormat = sDateFormat;
	};
	CSdtDatePickerPr.prototype.GetCalendar = function()
	{
		return this.Calendar;
	};
	CSdtDatePickerPr.prototype.SetCalendar = function(nCalendar)
	{
		this.Calendar = nCalendar;
	};
	CSdtDatePickerPr.prototype.GetFormatsExamples = function()
	{
		if (Asc.c_oAscDateTimeFormat[this.LangId])
			return Asc.c_oAscDateTimeFormat[this.LangId];
		
		return [
			"MM/DD/YYYY",
			"dddd\,\ mmmm\ dd\,\ yyyy",
			"DD\ MMMM\ YYYY",
			"MMMM\ DD\,\ YYYY",
			"DD-MMM-YY",
			"MMMM\ YY",
			"MMM-YY",
			"MM/DD/YYYY\ hh:mm\ AM/PM",
			"MM/DD/YYYY\ hh:mm:ss\ AM/PM",
			"hh:mm",
			"hh:mm:ss",
			"hh:mm\ AM/PM",
			"hh:mm:ss:\ AM/PM"
		];
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtDatePickerPr = CSdtDatePickerPr;

	window['AscCommon'].CSdtDatePickerPr    = CSdtDatePickerPr;
	window['AscCommon']['CSdtDatePickerPr'] = CSdtDatePickerPr;

	CSdtDatePickerPr.prototype['get_FullDate']        = CSdtDatePickerPr.prototype.GetFullDate;
	CSdtDatePickerPr.prototype['put_FullDate']        = CSdtDatePickerPr.prototype.SetFullDate;
	CSdtDatePickerPr.prototype['get_LangId']          = CSdtDatePickerPr.prototype.GetLangId;
	CSdtDatePickerPr.prototype['put_LangId']          = CSdtDatePickerPr.prototype.SetLangId;
	CSdtDatePickerPr.prototype['get_DateFormat']      = CSdtDatePickerPr.prototype.GetDateFormat;
	CSdtDatePickerPr.prototype['put_DateFormat']      = CSdtDatePickerPr.prototype.SetDateFormat;
	CSdtDatePickerPr.prototype['get_Calendar']        = CSdtDatePickerPr.prototype.GetCalendar;
	CSdtDatePickerPr.prototype['put_Calendar']        = CSdtDatePickerPr.prototype.SetCalendar;
	CSdtDatePickerPr.prototype['get_FormatsExamples'] = CSdtDatePickerPr.prototype.GetFormatsExamples;
	CSdtDatePickerPr.prototype['get_String']          = CSdtDatePickerPr.prototype.ToString;

})(window);
