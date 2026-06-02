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
	 * Генератор ключей специальных форм
	 * @param {AscWord.CFormsManager} formManager
	 * @constructor
	 */
	function CFormKeyGenerator(formManager)
	{
		this.GlobalCounter = 1;
		this.FormManager   = formManager
	}

	CFormKeyGenerator.prototype.GetNewKey = function(form)
	{
		let key = this.GenerateKey(form);
		if (form && form.IsRadioButton())
		{
			while (!this.CheckRadioButtonGroup(key))
			{
				key = this.GenerateKey(form);
			}
		}
		else
		{
			while (!this.CheckKey(key))
			{
				key = this.GenerateKey(form);
			}
		}
		return key;
	};
	CFormKeyGenerator.prototype.GetNewChoice = function(form)
	{
		if (!form || !form.IsRadioButton())
			return this.GetNewKey(form);
		
		let checkBoxPr = form.GetCheckBoxPr();
		return this.GetNewChoiceByGroupKey(checkBoxPr.GetGroupKey());
	};
	CFormKeyGenerator.prototype.GetNewChoiceByGroupKey = function(groupKey)
	{
		let buttons  = this.FormManager.GetRadioButtons(groupKey);
		let choiceKeys = {};
		for (let index = 0, count = buttons.length; index < count; ++index)
		{
			choiceKeys[buttons[index].GetFormKey()] = buttons;
		}
		
		let choiceNum = buttons.length;
		let newKey    = "Choice" + choiceNum;
		while (choiceKeys[newKey])
		{
			choiceNum++;
			newKey = "Choice" + choiceNum;
		}
		
		return newKey;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CFormKeyGenerator.prototype.CheckKey = function(key)
	{
		if (!key || "" === key)
			return true;

		let forms = this.FormManager.GetAllFormsByKey(key);
		return (!forms.length);
	};
	CFormKeyGenerator.prototype.CheckRadioButtonGroup = function(groupKey)
	{
		if (!groupKey || "" === groupKey)
			return true;
		
		let forms = this.FormManager.GetRadioButtons(groupKey);
		return (!forms.length);
	};
	CFormKeyGenerator.prototype.GenerateKey = function(form)
	{
		let counter = this.GlobalCounter++;

		if (!form)
			return "Form" + counter;
		else if (form.IsSignatureForm())
			return "Signature" + counter;
		else if (form.IsRadioButton())
			return "Group " + counter; // Добавил пробел, потому что в UI с пробелом генерится, чтобы не было разницы
		else if (form.IsComplexForm())
			return "Complex" + counter;
		else if (form.IsTextForm())
			return "Text" + counter;
		else if (form.IsCheckBox())
			return "CheckBox" + counter;
		else if (form.IsDropDownList())
			return "DropDown" + counter;
		else if (form.IsComboBox())
			return "ComboBox" + counter;
		else if (form.IsPicture())
			return "Image" + counter;

		return "Form" + counter;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CFormKeyGenerator = CFormKeyGenerator;

})(window);
