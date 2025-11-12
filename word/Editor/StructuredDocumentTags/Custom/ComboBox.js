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
	 * Класс, представляющий элемент списка Combobox или DropDownList
	 * @constructor
	 */
	function CSdtListItem()
	{
		this.DisplayText = "";
		this.Value       = "";
	}
	CSdtListItem.prototype.Copy = function()
	{
		var oItem = new CSdtListItem();
		oItem.DisplayText = this.DisplayText;
		oItem.Value       = this.Value;
		return oItem;
	};
	CSdtListItem.prototype.IsEqual = function(oItem)
	{
		return (this.DisplayText === oItem.DisplayText && this.Value === oItem.Value);
	};
	CSdtListItem.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteString2(this.DisplayText);
		oWriter.WriteString2(this.Value);
	};
	CSdtListItem.prototype.ReadFromBinary = function(oReader)
	{
		this.DisplayText = oReader.GetString2();
		this.Value       = oReader.GetString2();
	};

	/**
	 * Класс с настройками для выпадающего списка
	 * @constructor
	 */
	function CSdtComboBoxPr()
	{
		this.ListItems = [];
		this.LastValue = -1;
		this.AutoFit   = false;
		this.Format    = new AscWord.CTextFormFormat();
	}
	CSdtComboBoxPr.prototype.Copy = function()
	{
		var oList = new CSdtComboBoxPr();

		oList.LastValue = this.LastValue;
		oList.ListItems = [];

		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			oList.ListItems.push(this.ListItems[nIndex].Copy());
		}

		oList.AutoFit = this.AutoFit;
		oList.Format  = this.Format.Copy();

		return oList;
	};
	CSdtComboBoxPr.prototype.IsEqual = function(oOther)
	{
		if (!oOther
			|| this.LastValue !== oOther.LastValue
			|| this.ListItems.length !== oOther.ListItems.length
			|| !this.Format.IsEqual(oOther.Format))
			return false;

		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			if (!this.ListItems[nIndex].IsEqual(oOther.ListItems[nIndex]))
				return false;
		}

		return true;
	};
	CSdtComboBoxPr.prototype.AddItem = function(sDisplay, sValue, nPosition)
	{
		if (null !== this.GetTextByValue(sValue))
			return false;
		
		var oItem = new CSdtListItem();
		oItem.DisplayText = sDisplay;
		oItem.Value       = sValue;
		
		if (undefined === nPosition || nPosition < 0 || nPosition >= this.ListItems.length)
			this.ListItems.push(oItem);
		else
			this.ListItems.splice(nPosition, 0, oItem);
		
		return true;
	};
	CSdtComboBoxPr.prototype.RemoveItem = function(nIndex)
	{
		if (nIndex < 0 || nIndex >= this.ListItems.length)
			return;
		
		this.ListItems.splice(nIndex, 1);
	};
	CSdtComboBoxPr.prototype.Clear = function()
	{
		this.ListItems = [];
		this.LastValue = -1;
	};
	CSdtComboBoxPr.prototype.GetTextByValue = function(sValue)
	{
		if (!sValue || "" === sValue)
			return null;

		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			if (this.ListItems[nIndex].Value === sValue)
				return this.ListItems[nIndex].DisplayText;
		}

		return null;
	};
	CSdtComboBoxPr.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteLong(this.LastValue);
		oWriter.WriteLong(this.ListItems.length);
		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			this.ListItems[nIndex].WriteToBinary(oWriter);
		}
		oWriter.WriteBool(this.AutoFit);
		this.Format.WriteToBinary(oWriter);
	};
	CSdtComboBoxPr.prototype.ReadFromBinary = function(oReader)
	{
		this.LastValue = oReader.GetLong();

		var nCount = oReader.GetLong();
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			var oItem = new CSdtListItem();
			oItem.ReadFromBinary(oReader);
			this.ListItems.push(oItem);
		}
		this.AutoFit = oReader.GetBool();
		this.Format.ReadFromBinary(oReader);
	};
	CSdtComboBoxPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtComboBoxPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtComboBoxPr.prototype.Get = function(nIndex)
	{
		return this.ListItems[nIndex];
	};
	CSdtComboBoxPr.prototype.GetIndex = function(sValue)
	{
		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			if (this.ListItems[nIndex].Value === sValue)
				return nIndex;
		}
		
		return -1;
	};
	CSdtComboBoxPr.prototype.GetItemsCount = function()
	{
		return this.ListItems.length;
	};
	CSdtComboBoxPr.prototype.GetItemDisplayText = function(nIndex)
	{
		if (!this.ListItems[nIndex])
			return "";

		return this.ListItems[nIndex].DisplayText;
	};
	CSdtComboBoxPr.prototype.GetItemValue = function(nIndex)
	{
		if (!this.ListItems[nIndex])
			return "";

		return this.ListItems[nIndex].Value;
	};
	CSdtComboBoxPr.prototype.FindByText = function(sValue)
	{
		for (var nIndex = 0, nCount = this.ListItems.length; nIndex < nCount; ++nIndex)
		{
			if (this.ListItems[nIndex].DisplayText === sValue)
				return nIndex;
		}

		return -1;
	};
	CSdtComboBoxPr.prototype.GetAutoFit = function()
	{
		return this.AutoFit;
	};
	CSdtComboBoxPr.prototype.SetAutoFit = function(isAutoFit)
	{
		this.AutoFit = isAutoFit;
	};
	CSdtComboBoxPr.prototype.GetFormat = function()
	{
		return this.Format;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtComboBoxPr = CSdtComboBoxPr;
	window['AscWord'].CSdtListItem   = CSdtListItem;

	window['AscCommon'].CSdtComboBoxPr    = CSdtComboBoxPr;
	window['AscCommon']['CSdtComboBoxPr'] = CSdtComboBoxPr;

	CSdtComboBoxPr.prototype['add_Item']            = CSdtComboBoxPr.prototype.AddItem;
	CSdtComboBoxPr.prototype['clear']               = CSdtComboBoxPr.prototype.Clear;
	CSdtComboBoxPr.prototype['get_TextByValue']     = CSdtComboBoxPr.prototype.GetTextByValue;
	CSdtComboBoxPr.prototype['get_ItemsCount']      = CSdtComboBoxPr.prototype.GetItemsCount;
	CSdtComboBoxPr.prototype['get_ItemDisplayText'] = CSdtComboBoxPr.prototype.GetItemDisplayText;
	CSdtComboBoxPr.prototype['get_ItemValue']       = CSdtComboBoxPr.prototype.GetItemValue;
	CSdtComboBoxPr.prototype['get_AutoFit']         = CSdtComboBoxPr.prototype.GetAutoFit;
	CSdtComboBoxPr.prototype['put_AutoFit']         = CSdtComboBoxPr.prototype.SetAutoFit;


})(window);
