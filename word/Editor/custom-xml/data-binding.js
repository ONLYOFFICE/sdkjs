/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	DataBinding.prototype.recalculateCheckSum = function (stringOfCustomXMlContent)
	{
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
			writer.WriteLong(this.storeItemID);
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
