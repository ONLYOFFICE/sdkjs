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
	 * Class for working with adding field data
	 * @constructor
	 */
	function CAddinFieldData()
	{
		this.FieldId = undefined;
		this.Value   = undefined;
		this.Content = undefined;
	}
	CAddinFieldData.prototype.SetFieldId = function(fieldId)
	{
		this.FieldId = "" + fieldId;
	};
	CAddinFieldData.prototype.GetFieldId = function()
	{
		return this.FieldId;
	};
	CAddinFieldData.prototype.SetValue = function(value)
	{
		this.Value = "" + value;
	};
	CAddinFieldData.prototype.GetValue = function()
	{
		return this.Value;
	};
	CAddinFieldData.prototype.SetContent = function(content)
	{
		this.Content = "" + content;
	};
	CAddinFieldData.prototype.GetContent = function()
	{
		return this.Content;
	};
	CAddinFieldData.prototype.ToJson = function()
	{
		return {
			"FieldId" : this.FieldId,
			"Value"   : this.Value,
			"Content" : this.Content
		};
	};
	CAddinFieldData.FromJson = function(obj)
	{
		let newData = new CAddinFieldData();
		if (!obj)
			return newData;
		
		if (undefined !== obj.FieldId)
			newData.SetFieldId(obj.FieldId);
		else if (undefined !== obj["FieldId"])
			newData.SetFieldId(obj["FieldId"]);
		
		if (undefined !== obj.Value)
			newData.SetValue(obj.Value);
		else if (undefined !== obj["Value"])
			newData.SetValue(obj["Value"]);
		
		if (undefined !== obj.Content)
			newData.SetContent(obj.Content);
		else if (undefined !== obj["Content"])
			newData.SetContent(obj["Content"]);
		
		return newData;
	};
	CAddinFieldData.FromField = function(field)
	{
		if (!field
			|| !(field instanceof AscWord.CComplexField)
			|| !field.IsAddin()
			|| null === field.GetFieldId())
			return null;
		
		let data = new CAddinFieldData();
		data.SetFieldId(field.GetFieldId());
		data.SetValue(field.GetInstruction().GetValue());
		data.SetContent(field.GetFieldValueText());
		return data;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CAddinFieldData = CAddinFieldData;
	
})(window);

