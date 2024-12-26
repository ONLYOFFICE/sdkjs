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
		else if (form.IsDropDownList() || form.IsComboBox())
			return "DropDown" + counter;
		else if (form.IsPicture())
			return "Image" + counter;

		return "Form" + counter;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CFormKeyGenerator = CFormKeyGenerator;

})(window);
