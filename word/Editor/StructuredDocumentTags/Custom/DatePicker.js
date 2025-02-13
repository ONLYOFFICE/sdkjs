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
