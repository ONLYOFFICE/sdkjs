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
	const CASE_NONE     = 0;
	const CASE_UPPER    = 1;
	const CASE_LOWER    = 2;
	const CASE_FIRSTCAP = 3;
	const CASE_CAPS     = 4;

	// Built-in property names accepted by DOCPROPERTY (case-insensitive) and their source in the model
	const BUILTIN_PROPERTIES = {
		"title"                : {part : "core", key : "title"},
		"subject"              : {part : "core", key : "subject"},
		"author"               : {part : "core", key : "creator"},
		"keywords"             : {part : "core", key : "keywords"},
		"comments"             : {part : "core", key : "description"},
		"category"             : {part : "core", key : "category"},
		"lastsavedby"          : {part : "core", key : "lastModifiedBy"},
		"revisionnumber"       : {part : "core", key : "revision"},
		"createtime"           : {part : "core", key : "created", isDate : true},
		"lastsavedtime"        : {part : "core", key : "modified", isDate : true},
		"lastprinted"          : {part : "core", key : "lastPrinted", isDate : true},
		"company"              : {part : "app", key : "Company"},
		"manager"              : {part : "app", key : "Manager"},
		"template"             : {part : "app", key : "Template"},
		"hyperlinkbase"        : {part : "app", key : "HyperlinkBase"},
		"nameofapplication"    : {part : "app", key : "Application"},
		"totaleditingtime"     : {part : "app", key : "TotalTime"},
		"pages"                : {part : "app", key : "Pages"},
		"words"                : {part : "app", key : "Words"},
		"characters"           : {part : "app", key : "Characters"},
		"characterswithspaces" : {part : "app", key : "CharactersWithSpaces"},
		"lines"                : {part : "app", key : "Lines"},
		"paragraphs"           : {part : "app", key : "Paragraphs"},
		"security"             : {part : "app", key : "DocSecurity"}
	};

	// Document information fields which are shortcuts for a built-in property
	const INFO_FIELDS = {
		"TITLE"       : "Title",
		"SUBJECT"     : "Subject",
		"AUTHOR"      : "Author",
		"KEYWORDS"    : "Keywords",
		"COMMENTS"    : "Comments",
		"LASTSAVEDBY" : "LastSavedBy",
		"REVNUM"      : "RevisionNumber",
		"CREATEDATE"  : "CreateTime",
		"SAVEDATE"    : "LastSavedTime",
		"PRINTDATE"   : "LastPrinted",
		"TEMPLATE"    : "Template",
		"EDITTIME"    : "TotalEditingTime"
	};

	/**
	 * DOCPROPERTY field and document information fields (TITLE, AUTHOR, SUBJECT, ...)
	 * @param {string} [fieldName="DOCPROPERTY"]
	 * @constructor
	 */
	function FieldInstructionDOCPROPERTY(fieldName)
	{
		AscWord.FieldInstructionBase.call(this);

		this.FieldName     = fieldName ? fieldName.toUpperCase() : "DOCPROPERTY";
		this.PropertyName  = INFO_FIELDS[this.FieldName] ? INFO_FIELDS[this.FieldName] : "";
		this.Argument      = null; // Optional new value of an info field (TITLE "New title"), it is kept but not applied
		this.DateFormat    = null;
		this.NumericFormat = null;
		this.TextCase      = CASE_NONE;
	}
	FieldInstructionDOCPROPERTY.prototype = Object.create(AscWord.FieldInstructionBase.prototype);
	FieldInstructionDOCPROPERTY.prototype.constructor = FieldInstructionDOCPROPERTY;
	FieldInstructionDOCPROPERTY.prototype.Type = AscWord.fieldtype_DOCPROPERTY;

	FieldInstructionDOCPROPERTY.IsInfoField = function(fieldName)
	{
		return !!INFO_FIELDS[fieldName.toUpperCase()];
	};
	FieldInstructionDOCPROPERTY.prototype.IsInfoField = function()
	{
		return "DOCPROPERTY" !== this.FieldName;
	};
	FieldInstructionDOCPROPERTY.prototype.SetPropertyName = function(name)
	{
		this.PropertyName = name;
	};
	FieldInstructionDOCPROPERTY.prototype.GetPropertyName = function()
	{
		return this.PropertyName;
	};
	FieldInstructionDOCPROPERTY.prototype.SetArgument = function(value)
	{
		this.Argument = value;
	};
	FieldInstructionDOCPROPERTY.prototype.SetDateFormat = function(format)
	{
		this.DateFormat = format;
	};
	FieldInstructionDOCPROPERTY.prototype.SetNumericFormat = function(format)
	{
		this.NumericFormat = format;
	};
	FieldInstructionDOCPROPERTY.prototype.addGeneralSwitches = function(switches)
	{
		AscWord.FieldInstructionBase.prototype.addGeneralSwitches.call(this, switches);

		if (!switches)
			return;

		for (let i = 0; i < switches.length; ++i)
		{
			switch (switches[i].toUpperCase())
			{
				case "UPPER": this.TextCase = CASE_UPPER; break;
				case "LOWER": this.TextCase = CASE_LOWER; break;
				case "FIRSTCAP": this.TextCase = CASE_FIRSTCAP; break;
				case "CAPS": this.TextCase = CASE_CAPS; break;
			}
		}
	};
	/**
	 * Calculates the text of the field
	 * @returns {?string} null if the property doesn't exist
	 */
	FieldInstructionDOCPROPERTY.prototype.GetText = function()
	{
		let logicDocument = this.ComplexField ? this.ComplexField.LogicDocument : null;
		if (!logicDocument || !this.PropertyName)
			return null;

		let property = this.private_FindBuiltInProperty(logicDocument);
		if (!property)
			property = this.private_FindCustomProperty(logicDocument);

		if (!property)
			return null;

		return this.private_ApplyTextCase(this.private_FormatValue(property.value, property.isDate));
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	FieldInstructionDOCPROPERTY.prototype.writeField = function()
	{
		let result = " " + this.FieldName + " ";
		if (!this.IsInfoField())
			result += "\"" + this.PropertyName + "\" ";
		else if (null !== this.Argument)
			result += "\"" + this.Argument + "\" ";

		if (null !== this.DateFormat)
			result += "\\@ \"" + this.DateFormat + "\" ";

		if (null !== this.NumericFormat)
			result += "\\# \"" + this.NumericFormat + "\" ";

		return result;
	};
	FieldInstructionDOCPROPERTY.prototype.private_FindBuiltInProperty = function(logicDocument)
	{
		let builtIn = BUILTIN_PROPERTIES[this.PropertyName.toLowerCase()];
		if (!builtIn)
			return null;

		let part = "core" === builtIn.part ? logicDocument.Core : logicDocument.App;
		let value = part ? part[builtIn.key] : null;
		return {value : (undefined === value || null === value) ? "" : value, isDate : !!builtIn.isDate};
	};
	FieldInstructionDOCPROPERTY.prototype.private_FindCustomProperty = function(logicDocument)
	{
		let customProperties = logicDocument.CustomProperties;
		if (!customProperties)
			return null;

		let name       = this.PropertyName.toLowerCase();
		let properties = customProperties.getAllProperties();
		for (let i = 0, count = properties.length; i < count; ++i)
		{
			let property = properties[i];
			if (!property || !property.asc_getName() || property.asc_getName().toLowerCase() !== name)
				continue;

			let type  = property.asc_getType();
			let value = property.asc_getValue();
			return {
				value  : (undefined === value || null === value) ? "" : value,
				isDate : (AscCommon.c_oVariantTypes.vtDate === type || AscCommon.c_oVariantTypes.vtFiletime === type)
			};
		}
		return null;
	};
	FieldInstructionDOCPROPERTY.prototype.private_FormatValue = function(value, isDate)
	{
		if (isDate)
			return this.private_FormatDate(value);

		if ("boolean" === typeof(value))
			return value ? "Y" : "N";

		if ("number" === typeof(value))
			return this.private_FormatNumber(value);

		return "" + value;
	};
	FieldInstructionDOCPROPERTY.prototype.private_FormatDate = function(value)
	{
		let date = value instanceof Date ? value : (value ? new Date(value) : null);
		if (!date || isNaN(date.getTime()))
			return "";

		let langId = this.private_GetLang();
		let format = this.DateFormat;
		if (!format)
			format = 1033 === langId ? "M/d/yyyy h:mm:ss am/pm" : "dd/MM/yyyy HH:mm:ss";

		let numFormat = AscCommon.oNumFormatCache.get(format, AscCommon.NumFormatType.WordFieldDate);
		if (!numFormat)
			return "";

		let dateTime = new Asc.cDate(date.getTime());
		let excelDate = dateTime.getExcelDate(true) + (dateTime.getHours() * 60 * 60 + dateTime.getMinutes() * 60 + dateTime.getSeconds()) / AscCommonExcel.c_sPerDay;
		return numFormat.formatToWord(excelDate, 15, AscCommon.g_aCultureInfos[langId]);
	};
	FieldInstructionDOCPROPERTY.prototype.private_FormatNumber = function(value)
	{
		if (this.NumericFormat)
		{
			let numFormat = AscCommon.oNumFormatCache.get(this.NumericFormat, AscCommon.NumFormatType.WordFieldNumeric);
			if (numFormat)
				return numFormat.formatToWord(value, 15, AscCommon.g_aCultureInfos[this.private_GetLang()]);
		}

		if (this.haveNumericFormat() && Number.isInteger(value))
		{
			let textPr = this.private_GetTextPr();
			return "" + AscCommon.IntToNumberFormat(value, this.getNumericFormat(), {lang : textPr && textPr.Lang, isFromField : true, isSkipFractPart : true});
		}

		return "" + value;
	};
	FieldInstructionDOCPROPERTY.prototype.private_ApplyTextCase = function(text)
	{
		switch (this.TextCase)
		{
			case CASE_UPPER:
				return text.toUpperCase();
			case CASE_LOWER:
				return text.toLowerCase();
			case CASE_FIRSTCAP:
				return text.charAt(0).toUpperCase() + text.slice(1);
			case CASE_CAPS:
				return text.replace(/(^|\s)(\S)/g, function(match, space, letter) {return space + letter.toUpperCase();});
		}
		return text;
	};
	FieldInstructionDOCPROPERTY.prototype.private_GetTextPr = function()
	{
		let sepChar = this.ComplexField ? this.ComplexField.GetSeparateChar() : null;
		if (sepChar && sepChar.GetRun())
			return sepChar.GetRun().Get_CompiledPr(false);

		return null;
	};
	FieldInstructionDOCPROPERTY.prototype.private_GetLang = function()
	{
		let textPr = this.private_GetTextPr();
		return textPr && textPr.Lang && textPr.Lang.Val ? textPr.Lang.Val : 1033;
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].FieldInstructionDOCPROPERTY = FieldInstructionDOCPROPERTY;

})(window);
