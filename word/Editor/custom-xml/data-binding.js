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

(function()
{
	/**
	 * @constructor
	 */
	function DataBinding(prefix, itemID, xpath, checkSum)
	{
		this.prefixMappings		= prefix	?	prefix		: undefined;
		this.storeItemID 		= itemID	?	itemID		: undefined;
		this.xpath				= xpath		?	xpath		: undefined;
		this.storeItemCheckSum	= checkSum	?	checkSum	: undefined;
	}
	DataBinding.prototype.copy = function()
	{
		return new DataBinding(this.prefixMappings, this.storeItemID, this.xpath, this.storeItemCheckSum);
	};
	DataBinding.prototype.fillFromObject = function (obj)
	{
		this.prefixMappings		= obj.prefixMappings;
		this.storeItemID 		= obj.storeItemID;
		this.xpath				= obj.xpath;
		this.storeItemCheckSum	= obj.storeItemCheckSum;
	};
	DataBinding.prototype.getPrefix = function()
	{
		return this.prefixMappings;
	};
	DataBinding.prototype.getItemId = function()
	{
		return this.storeItemID;
	};
	DataBinding.prototype.getXPath = function()
	{
		return this.xpath;
	};
	DataBinding.prototype.getCheckSum = function()
	{
		return this.storeItemCheckSum;
	};
	DataBinding.prototype.recalculateCheckSum = function (stringOfCustomXMlContent)
	{
	};
	DataBinding.prototype.Refresh_RecalcData = function(Data)
	{
		// Ничего не делаем
	};
	DataBinding.prototype.toBinary = function(writer)
	{
		return this.Write_ToBinary(writer);
	};
	DataBinding.fromBinary = function(reader)
	{
		let data = new DataBinding();
		data.Read_FromBinary(reader);
		return data;
	};
	DataBinding.fromObject = function(obj)
	{
		let data = new DataBinding();
		data.fillFromObject(obj);
		return data;
	};
	DataBinding.prototype.Write_ToBinary = function(writer)
	{
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		let flags = 0;
		
		if (undefined !== this.prefixMappings)
		{
			writer.WriteString2(this.prefixMappings);
			flags |= 1;
		}
		
		if (undefined !== this.storeItemID)
		{
			writer.WriteString2(this.storeItemID);
			flags |= 2;
		}
		
		if (undefined !== this.xpath)
		{
			writer.WriteString2(this.xpath);
			flags |= 4;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	DataBinding.prototype.Read_FromBinary = function(reader)
	{
		let flags = reader.GetLong();
		if (flags & 1)
			this.prefixMappings = reader.GetString2();
		if (flags & 2)
			this.storeItemID = reader.GetString2();
		if (flags & 4)
			this.xpath = reader.GetString2();
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DataBinding = DataBinding;
	
})();
