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
	// TODO: Сейчас массив Forms содержит все формы, те что используются и те, что нет, из-за того что
	//       нет нормальной событийной системы, о том, что содержимое документа изменилось. Как только добавим такую
	//       систему, то данный класс на неё подпишется и будет обновлять массив Forms с проверкой IsUseInDocument, и
	//       на функции GetForms мы будем отдавать его сразу, а не формировать каждый раз новый массив с проверками
	//       IsUseInDocument. Это ускорит ситуации, где многократно вызываеется GetForms без изменения струкутуры
	//       документа.



	/**
	 * Класс для работы со всеми специальными формами внунтри документа
	 * @param {AscWord.CDocument} oLogicDocument
	 * @constructor
	 */
	function CFormsManager(oLogicDocument)
	{
		this.LogicDocument = oLogicDocument;

		// В мапе форм находятся вообще все формы. В списке находятся только самостоятельные формы, которые
		// не являются частью другой формы
		this.FormsMap     = {};
		this.Forms        = [];
		this.UpdateList   = false;
		this.KeyGenerator = new AscWord.CFormKeyGenerator(this);
	}
	CFormsManager.prototype.Register = function(oForm)
	{
		if (!oForm)
			return;

		let sId = oForm.GetId();

		if (oForm.IsForm())
			this.FormsMap[sId] = oForm;
		else
			delete this.FormsMap[sId];

		this.UpdateList = true;
	};
	CFormsManager.prototype.Unregister = function(oForm)
	{
		if (!oForm)
			return;

		delete this.FormsMap[oForm.GetId()];

		this.UpdateList = true;
	};
	/**
	 * @returns {[]}
	 */
	CFormsManager.prototype.GetAllForms = function()
	{
		this.CheckFormsList();

		let arrResult = [];
		for (let nIndex = 0, nCount = this.Forms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = this.Forms[nIndex];
			if (oForm.IsUseInDocument())
				arrResult.push(oForm);
		}

		return arrResult;
	};
	/**
	 * Получаем ключи форм по заданным параметрам
	 * @param oPr
	 * @returns {Array.string}
	 */
	CFormsManager.prototype.GetAllKeys = function(oPr)
	{
		let isText       = oPr && oPr.Text;
		let isComboBox   = oPr && oPr.ComboBox;
		let isDropDown   = oPr && oPr.DropDownList;
		let isCheckBox   = oPr && oPr.CheckBox;
		let isPicture    = oPr && oPr.Picture;
		let isRadioGroup = oPr && oPr.RadioGroup;
		let isComplex    = oPr && oPr.Complex;

		let arrKeys  = [];
		let arrForms = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = arrForms[nIndex];

			let sKey = null;

			let isComplexForm = oForm.IsComplexForm();
			if (isComplexForm && !isComplex)
				continue;

			if ((isComplex && isComplexForm)
				|| (isText && oForm.IsTextForm())
				|| (isComboBox && oForm.IsComboBox())
				|| (isDropDown && oForm.IsDropDownList())
				|| (isCheckBox && oForm.IsCheckBox() && !oForm.IsRadioButton())
				|| (isPicture && oForm.IsPicture()))
			{
				sKey = oForm.GetFormKey();
			}
			else if (isRadioGroup && oForm.IsRadioButton())
			{
				sKey = oForm.GetRadioButtonGroupKey();
			}

			if (sKey && -1 === arrKeys.indexOf(sKey))
				arrKeys.push(sKey);
		}

		return arrKeys;
	};
	/**
	 * Получаем массив всех специальных форм с заданным ключом
	 * @param sKey
	 * @param [formType=undefined] {Asc.c_oAscContentControlSpecificType}
	 * @returns {[]}
	 */
	CFormsManager.prototype.GetAllFormsByKey = function(sKey, formType)
	{
		let arrForms  = this.GetAllForms();
		let arrResult = [];
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = arrForms[nIndex];
			if (sKey === oForm.GetFormKey() && (undefined === formType || formType === oForm.GetSpecificType()))
				arrResult.push(oForm);
		}

		return arrResult;
	};
	/**
	 * Получаем массив всех специальных радио кнопок
	 * @param sGroupKey
	 * @returns {[]}
	 */
	CFormsManager.prototype.GetRadioButtons = function(sGroupKey)
	{
		let arrForms  = this.GetAllForms();
		let arrResult = [];
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = arrForms[nIndex];
			if (oForm.IsRadioButton() && sGroupKey === oForm.GetCheckBoxPr().GetGroupKey())
				arrResult.push(oForm);
		}

		return arrResult;
	};
	/**
	 * Все ли обязательные поля заполнены
	 * @param {string} [roleName]
	 * @returns {boolean}
	 */
	CFormsManager.prototype.IsAllRequiredFormsFilled = function(roleName)
	{
		// TODO: Сейчас у нас здесь идет проверка и на правильность заполнения форм с форматом
		// Возможно стоит разделить на 2 разные проверки и добавить одну общую проверку на правильность
		// заполненности формы, куда будут входить обе предыдущие проверки

		let arrForms = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = arrForms[nIndex];
			
			if (roleName && roleName !== oForm.GetFormRole())
				continue;
			
			if (oForm.IsFormRequired() && !oForm.IsFormFilled())
				return false;

			if (oForm.IsTextForm()
				&& !oForm.IsPlaceHolder()
				&& !oForm.GetTextFormPr().CheckFormat(oForm.GetInnerText(), true))
				return false;
		}
		return true;
	};
	/**
	 * Проверяем залоченность форм по заданному ключу
	 * @param nCheckType
	 * @param sKey
	 * @param oSkipParagraph
	 */
	CFormsManager.prototype.CheckLockByKey = function(nCheckType, sKey, oSkipParagraph)
	{
		let arrForms = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm      = arrForms[nIndex];
			let oParagraph = oForm.GetParagraph ? oForm.GetParagraph() : null;
			if (oParagraph
				&& oParagraph !== oSkipParagraph
				&& oForm.GetFormKey() === sKey)
			{
				if (oForm.IsPicture())
				{
					let arrDrawings = oForm.GetAllDrawingObjects();
					for (var nDrawingIndex = 0, nDrawingsCount = arrDrawings.length; nDrawingIndex < nDrawingsCount; ++nDrawingIndex)
					{
						let oDrawing = arrDrawings[nDrawingIndex];
						oDrawing.Lock.Check(oDrawing.GetId());
					}
				}
				else
				{
					oParagraph.Lock.Check(oParagraph.GetId());
				}
			}
		}
	};
	/**
	 * Изменяем другие формы, из-за изменения заданной формы
	 * @param oForm
	 */
	CFormsManager.prototype.OnChange = function(oForm)
	{
		if (!this.IsValidForm(oForm))
			return;
		
		if (oForm.IsComplexForm())
			this.OnChangeComplexForm(oForm);
		else if (oForm.IsCheckBox())
			this.OnChangeCheckBox(oForm);
		else if (oForm.IsPicture())
			this.OnChangePictureForm(oForm);
		else
			this.OnChangeTextForm(oForm);
	};
	/**
	 * Проверяем корректность изменения формы
	 * @param oForm
	 */
	CFormsManager.prototype.ValidateChangeOnFly = function(oForm)
	{
		if (!oForm.IsPlaceHolder() && !oForm.IsComplexForm() && oForm.IsTextForm())
		{
			let oTextFormPr = oForm.GetTextFormPr();
			if (!oTextFormPr.CheckFormatOnFly(oForm.GetInnerText()))
				return false;
		}

		return true;
	};
	/**
	 * @returns {AscWord.CFormKeyGenerator}
	 */
	CFormsManager.prototype.GetKeyGenerator = function()
	{
		return this.KeyGenerator;
	};
	/**
	 * Получаем данные всех форм
	 * @returns {array}
	 */
	CFormsManager.prototype.GetAllFormsData = function()
	{
		let data = [];

		let allForms = this.GetAllForms();
		let passedKeys = {};
		for (let index = 0, count = allForms.length; index < count; ++index)
		{
			let form = allForms[index];
			let key  = form.GetFormKey();
			let type = form.GetSpecificType();
			
			if (form.IsRadioButton())
				key = form.GetCheckBoxPr().GetGroupKey();
			
			if (!key || (passedKeys[key] && passedKeys[key][type]))
				continue;
			
			if (!passedKeys[key])
				passedKeys[key] = {};
			
			passedKeys[key][type] = form;
			
			let stringType = Asc.c_oAscContentControlSpecificType.toString(type);
			if (form.IsRadioButton())
				stringType = "radio";
			
			let roleColor = form.GetRoleColor();
			
			data.push({
				"key"       : key,
				"tag"       : form.GetTag(),
				"value"     : this.GetFormValue(form),
				"type"      : stringType,
				"role"      : form.GetFormRole(),
				"roleColor" : roleColor ? roleColor.ToHexColor() : undefined
			});
		}
		
		return data;
	};
	CFormsManager.prototype.SetAllFormsData = function(data)
	{
		if (!data || !Array.isArray(data))
			return;
		
		for (let index = 0, count = data.length; index < count; ++index)
		{
			let key   = data[index]["key"];
			let value = data[index]["value"];
			let type  = data[index]["type"];
			
			if (undefined !== type && null !== type)
				type = Asc.c_oAscContentControlSpecificType.fromString(type);
			
			let forms = this.GetAllFormsByKey(key, type);
			let form  = forms[0];
			if (!form)
			{
				let radioGroup = this.GetRadioButtons(key);
				if (!radioGroup.length)
					continue;
				
				this.SetRadioGroupValue(key, value);
			}
			else
			{
				form.SetFormValue(value);
			}
			
			this.OnChange(form);
		}
	};
	CFormsManager.prototype.GetFormValue = function(form)
	{
		if (!form)
			return null;

		if (form.IsRadioButton())
			return this.GetRadioGroupValue(form.GetRadioButtonGroupKey());

		return form.GetFormValue();
	};
	/**
	 * Получем роль по текущему ключу и возможно заданному типу формы
	 * @param key
	 * @param [formType=undefined] {Asc.c_oAscContentControlSpecificType}
	 * @returns {string}
	 */
	CFormsManager.prototype.GetRoleByKey = function(key, formType)
	{
		let allForms = this.GetAllFormsByKey(key);
		for (let index = 0, count = allForms.length; index < count; ++index)
		{
			let form = allForms[index];
			if (undefined !== formType && formType !== form.GetSpecificType())
				continue;
			
			let role = form.GetFormRole();
			if (role)
				return role;
		}
		
		return "";
	};
	CFormsManager.prototype.OnEndLoad = function()
	{
		// Проверим, что у всех форм есть ключи
		let keyGenerator = this.GetKeyGenerator();
		let allForms     = this.GetAllForms();
		for (let index = 0, count = allForms.length; index < count; ++index)
		{
			let form = allForms[index];
			
			if (form.IsRadioButton())
			{
				let groupKey = form.GetRadioButtonGroupKey();
				if (!groupKey || "" === groupKey)
				{
					groupKey = keyGenerator.GetNewKey(form);
					form.SetRadioButtonGroupKey(groupKey);
				}
				
				let choice = form.GetFormKey();
				if (!choice || "" === choice)
				{
					choice = keyGenerator.GetNewChoice(form);
					form.SetFormKey(choice);
				}
			}
			else
			{
				let key = form.GetFormKey();
				if (key && "" !== key)
					continue;
				
				key = keyGenerator.GetNewKey(form);
				form.SetFormKey(key);
			}
		}
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CFormsManager.prototype.CheckFormsList = function()
	{
		_flushFormToCheck();
		if (!this.UpdateList)
			return;

		this.Forms.length = 0;
		for (let sId in this.FormsMap)
		{
			let oForm = this.FormsMap[sId];
			if (oForm.IsForm() && oForm === oForm.GetMainForm())
				this.Forms.push(oForm);
		}

		this.UpdateList = false;
	};
	CFormsManager.prototype.OnChangeCheckBox = function(oForm)
	{
		let isChecked  = oForm.GetCheckBoxPr().Checked;
		let userMaster = this.GetUserMasterByForm(oForm);
		let arrForms   = this.GetAllForms();

		if (oForm.IsRadioButton())
		{
			if (!oForm.GetCheckBoxPr().GetChecked())
				return;
			
			let sKey = oForm.GetCheckBoxPr().GetGroupKey();
			for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
			{
				let oTempForm = arrForms[nIndex];
				if (oTempForm.IsComplexForm()
					|| oTempForm === oForm
					|| !oTempForm.IsRadioButton()
					|| sKey !== oTempForm.GetCheckBoxPr().GetGroupKey()
					|| userMaster !== this.GetUserMasterByForm(oTempForm))
					continue;

				if (oTempForm.GetCheckBoxPr().GetChecked())
					oTempForm.SetCheckBoxChecked(false);
			}
		}
		else
		{
			let sKey = oForm.GetFormKey();
			for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
			{
				let oTempForm = arrForms[nIndex];
				if (oTempForm.IsComplexForm()
					|| oTempForm === oForm
					|| !oTempForm.IsCheckBox()
					|| oTempForm.IsRadioButton()
					|| sKey !== oTempForm.GetFormKey()
					|| isChecked === oTempForm.GetCheckBoxPr().GetChecked()
					|| userMaster !== this.GetUserMasterByForm(oTempForm))
					continue;

				oTempForm.ToggleCheckBox();
			}
		}
	};
	CFormsManager.prototype.OnChangePictureForm = function(oForm)
	{
		let sImageUrl;
		let arrDrawings = oForm.GetAllDrawingObjects();
		if (arrDrawings.length)
			sImageUrl = arrDrawings[0].GetPictureUrl();

		if (!sImageUrl)
			return;

		let sKey          = oForm.GetFormKey();
		let isPlaceHolder = oForm.IsPlaceHolder();
		let userMaster    = this.GetUserMasterByForm(oForm);
		let arrForms      = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oTempForm = arrForms[nIndex];
			if (oTempForm.IsComplexForm()
				|| oTempForm === oForm
				|| sKey !== oTempForm.GetFormKey()
				|| !oTempForm.IsPicture()
				|| userMaster !== this.GetUserMasterByForm(oTempForm))
				continue;

			let arrDrawings = oTempForm.GetAllDrawingObjects();
			if (arrDrawings.length > 0)
			{
				let oPicture = arrDrawings[0].GetPicture();
				if (oPicture)
					oPicture.setBlipFill(AscFormat.CreateBlipFillRasterImageId(sImageUrl));
			}
			oTempForm.SetShowingPlcHdr(isPlaceHolder);
			oTempForm.UpdatePictureFormLayout();
		}
	};
	CFormsManager.prototype.OnChangeTextForm = function(oForm)
	{
		let sKey          = oForm.GetFormKey();
		let isPlaceHolder = oForm.IsPlaceHolder();
		let oSrcRun       = !isPlaceHolder ? oForm.MakeSingleRunElement(false) : null;
		let userMaster    = this.GetUserMasterByForm(oForm);
		let arrForms      = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oTempForm = arrForms[nIndex];

			if (oTempForm.IsComplexForm()
				|| oTempForm.IsPicture()
				|| oTempForm.IsCheckBox()
				|| oTempForm === oForm
				|| sKey !== oTempForm.GetFormKey()
				|| userMaster !== this.GetUserMasterByForm(oTempForm))
				continue;

			if (isPlaceHolder)
			{
				if (!oTempForm.IsPlaceHolder())
					oTempForm.ReplaceContentWithPlaceHolder(false);
			}
			else
			{
				if (oTempForm.IsPlaceHolder())
					oTempForm.ReplacePlaceHolderWithContent();

				let oDstRun = oTempForm.MakeSingleRunElement(false);
				oDstRun.CopyTextFormContent(oSrcRun);
			}
		}
	};
	CFormsManager.prototype.OnChangeComplexForm = function(oForm)
	{
		let sKey          = oForm.GetFormKey();
		let isPlaceholder = oForm.IsPlaceHolder();
		let userMaster    = this.GetUserMasterByForm(oForm);
		let arrForms      = this.GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oTempForm = arrForms[nIndex];
			if (!oTempForm.IsComplexForm()
				|| oTempForm === oForm
				|| sKey !== oTempForm.GetFormKey()
				|| userMaster !== this.GetUserMasterByForm(oTempForm))
				continue;

			// TODO: Сейчас мы полностью перезаписываем содержимое поля. Можно проверить, что поле состоит из таких
			//       же базовых подклассов и попробовать обновить их каждый по отдельности, что бы было меньше изменений

			oTempForm.SetShowingPlcHdr(isPlaceholder);
			oTempForm.RemoveAll();
			for (let nPos = 0, nItemsCount = oForm.GetElementsCount(); nPos < nItemsCount; ++nPos)
			{
				oTempForm.AddToContent(nPos, oForm.GetElement(nPos).Copy());
			}
		}
	};
	CFormsManager.prototype.GetRadioGroupValue = function(groupKey)
	{
		let group = this.GetRadioButtons(groupKey);
		for (let index = 0, count = group.length; index < count; ++index)
		{
			let radioButton = group[index];
			if (radioButton.IsCheckBoxChecked() && radioButton.GetFormKey())
				return radioButton.GetFormKey();
		}

		return "";
	};
	CFormsManager.prototype.SetRadioGroupValue = function(groupKey, value)
	{
		let group = this.GetRadioButtons(groupKey);
		for (let index = 0, count = group.length; index < count; ++index)
		{
			let radioButton = group[index];
			if (radioButton.GetFormKey() !== value)
				radioButton.SetCheckBoxChecked(false);
			else
				radioButton.SetCheckBoxChecked(true);
		}
		
		return "";
	};
	CFormsManager.prototype.GetUserMasterByForm = function(form)
	{
		if (!form)
			return null;
		
		let fieldMaster = form.GetFieldMaster();
		if (!fieldMaster)
			return null;
		
		return fieldMaster.getFirstUser();
	};
	CFormsManager.prototype.IsValidForm = function(form)
	{
		return (form && form.IsUseInDocument());
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Здесь мы копим список форм, которые еще никуда не добавлены и надо будет их обработать при первом обращении
	let formToCheck = [];
	function _registerForm(form)
	{
		let formId = form.GetId();
		if (formToCheck[formId])
			delete formToCheck[formId];
		
		let para = form.GetParagraph();
		if (!para)
			return;
		
		let logicDocument = para.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		logicDocument.GetFormsManager().Register(form);
	}
	function _flushFormToCheck()
	{
		formToCheck.forEach(_registerForm);
		formToCheck.length = 0;
	}
	function registerForm(form)
	{
		let para = form.GetParagraph();
		if (!para || !para.GetLogicDocument())
		{
			formToCheck.push(form);
			return;
		}
		
		_registerForm(form);
	}
	function unregisterForm(form)
	{
		let para = form.GetParagraph();
		let logicDocument = para.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		logicDocument.GetFormsManager().Unregister(form);
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CFormsManager  = CFormsManager;
	window['AscWord'].registerForm   = registerForm;
	window['AscWord'].unregisterForm = unregisterForm;
	
	

})(window);
