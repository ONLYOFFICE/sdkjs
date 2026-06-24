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
	 * Функция конвертации формы в Json представление
	 * @param oForm
	 * @returns {object}
	 */
	function FormToJson(oForm)
	{
		if (!oForm || !oForm.IsForm())
			return {};

		Buffer = [];
		FormToBuffer(oForm);

		return {
			"preview" : "",
			"type"    : "custom",
			"format"  : Buffer
		};
	}

	/**
	 * Создание формы по Json объекту
	 * @param json
	 * @param {?AscWord.CInlineLevelSdt} form
	 * @returns {?CInlineLevelSdt}
	 */
	function JsonToForm(json, form)
	{
		let arrFormat = json["format"];

		if (!arrFormat
			|| !Array.isArray(arrFormat)
			|| !arrFormat.length)
			return null;

		let oComplexForm;
		if (form)
		{
			oComplexForm = form;
		}
		else
		{
			oComplexForm = new AscWord.CInlineLevelSdt();
			oComplexForm.SetComplexFormPr(new AscWord.CSdtComplexFormPr());
		}

		oComplexForm.RemoveFromContent(0, oComplexForm.GetElementsCount());

		for (let nIndex = 0, nCount = arrFormat.length; nIndex < nCount; ++nIndex)
		{
			let oElement = arrFormat[nIndex];
			if (typeof(oElement) === "string")
			{
				let oRun = new AscWord.CRun();
				oRun.AddText(oElement);
				oComplexForm.AddToContentToEnd(oRun);
			}
			else
			{
				let oForm = JsonToSimpleForm(oElement);
				if (oForm)
					oComplexForm.AddToContentToEnd(oForm);
			}
		}

		if (oComplexForm.GetElementsCount())
			oComplexForm.SetShowingPlcHdr(false);
		else
			oComplexForm.ReplaceContentWithPlaceHolder(false);

		return oComplexForm;
	}

	/**
	 * Получем массив всех юниковод, используемых в заданной форме
	 * @param json
	 * @returns {Array.number}
	 */
	function GetUnicodesFromJsonToForm(json)
	{
		if (!json)
			return [];

		let format = json["format"];

		if (!format
			|| !Array.isArray(format)
			|| !format.length)
			return [];

		let codePoints = [];
		for (let index = 0, count = format.length; index < count; ++index)
		{
			let element = format[index];
			if (typeof(element) === "string")
			{
				AppendString(codePoints, element);
			}
			else
			{
				let placeholder = element["placeholder"];
				if (placeholder)
					AppendString(codePoints, placeholder);

				let type = element["type"];
				if ("checkBox" === type)
				{
					if (element["checkedSymbol"])
						AppendCodePoint(codePoints, codePoints.push(element["checkedSymbol"]));

					if (element["uncheckedSymbol"])
						AppendCodePoint(codePoints, codePoints.push(element["uncheckedSymbol"]));
				}
				else if ("comboBox" === type)
				{
					let choice = element["choice"];
					if (Array.isArray(choice))
					{
						for (let choiceIndex = 0, choiceLen = choice.length; choiceIndex < choiceLen; ++choiceIndex)
						{
							let value = choice[choiceIndex];
							if (typeof(value) === "string" && value)
								AppendString(codePoints, value);
						}
					}
				}
			}
		}

		return codePoints;
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	let Buffer = [];

	function FormToBuffer(oForm)
	{
		if (oForm.IsComplexForm())
			ComplexFormToBuffer(oForm);
		else if (oForm.IsTextForm())
			TextFormToBuffer(oForm);
		else if (oForm.IsCheckBox())
			CheckBoxToBuffer(oForm);
		else if (oForm.IsComboBox() || oForm.IsDropDownList())
			ComboBoxToBuffer(oForm);
		else if (oForm.IsPictureForm())
			PictureToBuffer(oForm);
	}
	function ComplexFormToBuffer(oForm)
	{
		for (let nPos = 0, nCount = oForm.GetElementsCount(); nPos < nCount; ++nPos)
		{
			let oElement = oForm.GetElement(nPos);
			if (oElement instanceof AscWord.CInlineLevelSdt)
				FormToBuffer(oElement);
			else if (oElement instanceof AscWord.CRun && !oElement.IsEmpty())
				Buffer.push(oElement.GetText());
		}
	}
	function TextFormToBuffer(oForm)
	{
		let oPr = oForm.GetTextFormPr();
		Buffer.push({
			"type"          : "text",
			"placeholder"   : oForm.GetPlaceholderText(),
			"maxCharacters" : oPr.GetMaxCharacters(),
			"comb"          : oPr.IsComb(),
			"format"        : oPr.GetFormat().ToJson()
		});
	}
	function CheckBoxToBuffer(oForm)
	{
		let pr = oForm.GetCheckBoxPr();
		Buffer.push({
			"type"            : "checkBox",
			"checked"         : pr.GetChecked(),
			"checkedSymbol"   : pr.GetCheckedSymbol(),
			"uncheckedSymbol" : pr.GetUncheckedSymbol()
		});
	}
	function ComboBoxToBuffer(oForm)
	{
		let oPr = oForm.IsComboBox() ? oForm.GetComboBoxPr() : oForm.GetDropDownListPr();
		let object = {
			"type"          : "comboBox",
			"edit"          : oForm.IsComboBox(),
			"placeholder"   : oForm.GetPlaceholderText(),
			"format"        : oPr.GetFormat().ToJson(),
			"choice"        : []
		};

		for (let nIndex = 0, nCount = oPr.GetItemsCount(); nIndex < nCount; ++nIndex)
		{
			object.choice.push(oPr.GetItemDisplayText(nIndex));
		}

		Buffer.push(object);
	}
	function PictureToBuffer(oForm)
	{
		Buffer.push({
			"type" : "picture"
		});
	}
	function JsonToSimpleForm(json)
	{
		let form = null;

		let nType = json["type"];
		if ("text" === nType)
			form = JsonToTextForm(json);
		else if ("checkBox" === nType)
			form = JsonToCheckBox(json);
		else if ("comboBox" === nType)
			form = JsonToComboBox(json);
		else if ("picture" === nType)
			form = JsonToPicture(json);

		if (form)
			form.SetFormPr(new AscWord.CSdtFormPr());

		return form;
	}
	function JsonToTextForm(json)
	{
		let textForm   = new AscWord.CInlineLevelSdt();
		let textFormPr = new AscWord.CSdtTextFormPr();

		if (json["maxCharacters"] && json["maxCharacters"] > 0)
		{
			textFormPr.SetMaxCharacters(json["maxCharacters"]);
			textFormPr.SetComb(!!json["comb"]);
		}

		textFormPr.GetFormat().FromJson(json["format"]);

		textForm.ApplyTextFormPr(textFormPr);

		let placeholder = json["placeholder"];
		if (placeholder && placeholder !== textForm.GetPlaceholderText())
			textForm.SetPlaceholderText(placeholder);

		textForm.ReplaceContentWithPlaceHolder();

		return textForm;
	}
	function JsonToCheckBox(json)
	{
		let checkBox   = new AscWord.CInlineLevelSdt();
		let checkBoxPr = new AscWord.CSdtCheckBoxPr();
		checkBoxPr.SetChecked(!!json["checked"]);

		if (json["checkedSymbol"])
			checkBoxPr.SetCheckedSymbol(json["checkedSymbol"]);

		if (json["uncheckedSymbol"])
			checkBoxPr.SetUncheckedSymbol(json["uncheckedSymbol"]);

		checkBox.ApplyCheckBoxPr(checkBoxPr);
		return checkBox;
	}
	function JsonToComboBox(json)
	{
		let comboBox   = new AscWord.CInlineLevelSdt();
		let comboBoxPr = new AscWord.CSdtComboBoxPr();

		let items = json["choice"];
		if (Array.isArray(items))
		{
			for (let index = 0, count = items.length; index < count; ++index)
			{
				let value = items[index];
				if (typeof(value) === "string" && value)
					comboBoxPr.AddItem(value, value);
			}
		}
		comboBoxPr.GetFormat().ToJson(json["format"]);

		if (json["edit"])
			comboBox.ApplyComboBoxPr(comboBoxPr);
		else
			comboBox.ApplyDropDownListPr(comboBoxPr);

		let placeholder = json["placeholder"];
		if (placeholder && placeholder !== comboBox.GetPlaceholderText())
			comboBox.SetPlaceholderText(placeholder);

		comboBox.ReplaceContentWithPlaceHolder();

		return comboBox;
	}
	function JsonToPicture(json)
	{
		// TODO: Реализовать
		return null;
	}
	function AppendString(buffer, string)
	{
		for (let iterator = string.getUnicodeIterator(); iterator.check(); iterator.next())
		{
			AppendCodePoint(buffer, iterator.value());
		}
	}
	function AppendCodePoint(buffer, codePoint)
	{
		if (-1 === buffer.indexOf(codePoint))
			buffer.push(codePoint);
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].FormToJson                = FormToJson;
	window['AscWord'].JsonToForm                = JsonToForm;
	window['AscWord'].GetUnicodesFromJsonToForm = GetUnicodesFromJsonToForm;

})(window);
