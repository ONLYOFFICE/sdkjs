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
	 * Класс с настройками чекбокса
	 * @constructor
	 */
	function CSdtCheckBoxPr()
	{
		this.Checked         = false;
		this.CheckedSymbol   = Asc.c_oAscSdtCheckBoxDefaults.CheckedSymbol;
		this.UncheckedSymbol = Asc.c_oAscSdtCheckBoxDefaults.UncheckedSymbol;
		this.CheckedFont     = Asc.c_oAscSdtCheckBoxDefaults.CheckedFont;
		this.UncheckedFont   = Asc.c_oAscSdtCheckBoxDefaults.UncheckedFont;
		this.GroupKey        = undefined;
		this.ChoiceName      = undefined;
		this.Label           = undefined;
	}
	CSdtCheckBoxPr.prototype.Copy = function()
	{
		var oCopy = new CSdtCheckBoxPr();

		oCopy.Checked         = this.Checked;
		oCopy.CheckedSymbol   = this.CheckedSymbol;
		oCopy.CheckedFont     = this.CheckedFont;
		oCopy.UncheckedSymbol = this.UncheckedSymbol;
		oCopy.UncheckedFont   = this.UncheckedFont;
		oCopy.GroupKey        = this.GroupKey;

		return oCopy;
	};
	CSdtCheckBoxPr.prototype.IsEqual = function(oOther)
	{
		if (!oOther
			|| oOther.Checked !== this.Checked
			|| oOther.CheckedSymbol !== this.CheckedSymbol
			|| oOther.CheckedFont !== this.CheckedFont
			|| oOther.UncheckedSymbol !== this.UncheckedSymbol
			|| oOther.UncheckedFont !== this.UncheckedFont
			|| oOther.GroupKey !== this.GroupKey)
			return false;

		return true;
	};
	CSdtCheckBoxPr.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteBool(this.Checked);
		oWriter.WriteString2(this.CheckedFont);
		oWriter.WriteLong(this.CheckedSymbol);
		oWriter.WriteString2(this.UncheckedFont);
		oWriter.WriteLong(this.UncheckedSymbol);

		if (undefined !== this.GroupKey)
		{
			oWriter.WriteBool(true);
			oWriter.WriteString2(this.GroupKey);
		}
		else
		{
			oWriter.WriteBool(false);
		}
	};
	CSdtCheckBoxPr.prototype.ReadFromBinary = function(oReader)
	{
		this.Checked         = oReader.GetBool();
		this.CheckedFont     = oReader.GetString2();
		this.CheckedSymbol   = oReader.GetLong();
		this.UncheckedFont   = oReader.GetString2();
		this.UncheckedSymbol = oReader.GetLong();

		if (oReader.GetBool())
			this.GroupKey = oReader.GetString2();
	};
	CSdtCheckBoxPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtCheckBoxPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtCheckBoxPr.prototype.SetChecked = function(isChecked)
	{
		this.Checked = isChecked;
	};
	CSdtCheckBoxPr.prototype.GetChecked = function()
	{
		return this.Checked;
	};
	CSdtCheckBoxPr.prototype.GetCheckedSymbol = function()
	{
		return this.CheckedSymbol;
	};
	CSdtCheckBoxPr.prototype.SetCheckedSymbol = function(nSymbol)
	{
		this.CheckedSymbol = nSymbol;
	};
	CSdtCheckBoxPr.prototype.GetCheckedFont = function()
	{
		return this.CheckedFont;
	};
	CSdtCheckBoxPr.prototype.SetCheckedFont = function(sFont)
	{
		this.CheckedFont = sFont;
	};
	CSdtCheckBoxPr.prototype.GetUncheckedSymbol = function()
	{
		return this.UncheckedSymbol;
	};
	CSdtCheckBoxPr.prototype.SetUncheckedSymbol = function(nSymbol)
	{
		this.UncheckedSymbol = nSymbol;
	};
	CSdtCheckBoxPr.prototype.GetUncheckedFont = function()
	{
		return this.UncheckedFont;
	};
	CSdtCheckBoxPr.prototype.SetUncheckedFont = function(sFont)
	{
		this.UncheckedFont = sFont;
	};
	CSdtCheckBoxPr.prototype.GetGroupKey = function()
	{
		return this.GroupKey;
	};
	CSdtCheckBoxPr.prototype.SetGroupKey = function(sKey)
	{
		this.GroupKey = sKey;
	};
	CSdtCheckBoxPr.prototype.SetChoiceName = function(name)
	{
		this.ChoiceName = name;
	};
	CSdtCheckBoxPr.prototype.GetChoiceName = function(returnRaw)
	{
		if (returnRaw)
			return this.ChoiceName;
		
		return this.ChoiceName ? this.ChoiceName : "";
	};
	CSdtCheckBoxPr.prototype.SetLabel = function(text)
	{
		this.Label = text;
	};
	CSdtCheckBoxPr.prototype.GetLabel = function()
	{
		return this.Label ? this.Label : "";
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'].CSdtCheckBoxPr    = CSdtCheckBoxPr;
	window['AscCommon']['CSdtCheckBoxPr'] = CSdtCheckBoxPr;

	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtCheckBoxPr = CSdtCheckBoxPr;

	CSdtCheckBoxPr.prototype['get_Checked']         = CSdtCheckBoxPr.prototype.GetChecked;
	CSdtCheckBoxPr.prototype['put_Checked']         = CSdtCheckBoxPr.prototype.SetChecked;
	CSdtCheckBoxPr.prototype['get_CheckedSymbol']   = CSdtCheckBoxPr.prototype.GetCheckedSymbol;
	CSdtCheckBoxPr.prototype['put_CheckedSymbol']   = CSdtCheckBoxPr.prototype.SetCheckedSymbol;
	CSdtCheckBoxPr.prototype['get_CheckedFont']     = CSdtCheckBoxPr.prototype.GetCheckedFont;
	CSdtCheckBoxPr.prototype['put_CheckedFont']     = CSdtCheckBoxPr.prototype.SetCheckedFont;
	CSdtCheckBoxPr.prototype['get_UncheckedSymbol'] = CSdtCheckBoxPr.prototype.GetUncheckedSymbol;
	CSdtCheckBoxPr.prototype['put_UncheckedSymbol'] = CSdtCheckBoxPr.prototype.SetUncheckedSymbol;
	CSdtCheckBoxPr.prototype['get_UncheckedFont']   = CSdtCheckBoxPr.prototype.GetUncheckedFont;
	CSdtCheckBoxPr.prototype['put_UncheckedFont']   = CSdtCheckBoxPr.prototype.SetUncheckedFont;
	CSdtCheckBoxPr.prototype['get_GroupKey']        = CSdtCheckBoxPr.prototype.GetGroupKey;
	CSdtCheckBoxPr.prototype['put_GroupKey']        = CSdtCheckBoxPr.prototype.SetGroupKey;
	CSdtCheckBoxPr.prototype['get_ChoiceName']      = CSdtCheckBoxPr.prototype.GetChoiceName;
	CSdtCheckBoxPr.prototype['put_ChoiceName']      = CSdtCheckBoxPr.prototype.SetChoiceName;
	CSdtCheckBoxPr.prototype['get_Label']           = CSdtCheckBoxPr.prototype.GetLabel;
	CSdtCheckBoxPr.prototype['put_Label']           = CSdtCheckBoxPr.prototype.SetLabel;
	
})(window);
