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
	const ComplexFormType = {
		Custom          : 0,
		Telephone       : 1,
		Email           : 2,
		LabeledCheckBox : 3
	};

	/**
	 *
	 * @constructor
	 */
	function CSdtComplexFormPr(type)
	{
		this.Type = undefined !== type ? type : ComplexFormType.Custom;
	}
	CSdtComplexFormPr.prototype.Copy = function()
	{
		let formPr = new CSdtComplexFormPr();
		formPr.Type = this.Type;
		return formPr;
	};
	CSdtComplexFormPr.prototype.IsEqual = function(formPr)
	{
		return (formPr instanceof CSdtComplexFormPr
			&& formPr.Type === this.Type);
	};
	CSdtComplexFormPr.prototype.GetType = function()
	{
		return this.Type;
	};
	CSdtComplexFormPr.prototype.Write_ToBinary = function(writer)
	{
		this.WriteToBinary(writer);
	};
	CSdtComplexFormPr.prototype.Read_FromBinary = function(reader)
	{
		this.ReadFromBinary(reader);
	};
	CSdtComplexFormPr.prototype.WriteToBinary = function(writer)
	{
		writer.WriteLong(this.Type);
	};
	CSdtComplexFormPr.prototype.ReadFromBinary = function(reader)
	{
		this.Type = reader.GetLong();
	};
	CSdtComplexFormPr.prototype.IsLabeledCheckBox = function()
	{
		return this.Type === ComplexFormType.LabeledCheckBox;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtComplexFormPr = CSdtComplexFormPr

	let exportPrototype          = window['Asc']['ComplexFormType'] = window['Asc'].ComplexFormType = ComplexFormType;
	exportPrototype['Custom']          = exportPrototype.Custom;
	exportPrototype['Telephone']       = exportPrototype.Telephone;
	exportPrototype['Email']           = exportPrototype.Email;
	exportPrototype['LabeledCheckBox'] = exportPrototype.LabeledCheckBox;

})(window);
