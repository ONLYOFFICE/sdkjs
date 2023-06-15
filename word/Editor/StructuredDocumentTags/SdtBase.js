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
/**
 * User: Ilja.Kirillov
 * Date: 02.04.2020
 * Time: 15:50
 */

/**
 * Базовый класс для контент контролов
 * @constructor
 */
function CSdtBase()
{}

/**
 * Получаем текст плейсхолдера
 * @returns {string}
 */
CSdtBase.prototype.GetPlaceholderText = function()
{
	var oDocPart = this.GetLogicDocument().GetGlossaryDocument().GetDocPartByName(this.GetPlaceholder());
	if (oDocPart)
	{
		var oFirstParagraph = oDocPart.GetFirstParagraph();
		return oFirstParagraph.GetText({ParaEndToSpace : false});
	}

	return String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode);
};
/**
 * Выставляем простой плейсхолдера в виде текста
 * @param {string} sText
 */
CSdtBase.prototype.SetPlaceholderText = function(sText)
{
	if (!sText)
		return this.SetPlaceholder(undefined);

	if (sText === this.GetPlaceholderText())
		return;

	var oLogicDocument = this.GetLogicDocument();
	var oGlossary      = oLogicDocument.GetGlossaryDocument();

	var oDocPart = oGlossary.GetDocPartByName(this.GetPlaceholder());
	if (!oDocPart || oGlossary.IsDefaultDocPart(oDocPart))
	{
		var oNewDocPart;
		if (!oDocPart)
			oNewDocPart = oGlossary.CreateDocPart(oGlossary.GetNewName());
		else
			oNewDocPart = oDocPart.Copy(oGlossary.GetNewName());

		this.SetPlaceholder(oNewDocPart.GetDocPartName());
		oDocPart = oNewDocPart;
	}

	oDocPart.RemoveSelection();
	oDocPart.MoveCursorToStartPos();

	var oTextPr = oDocPart.GetDirectTextPr();
	var oParaPr = oDocPart.GetDirectParaPr();

	oDocPart.ClearContent(true);
	if (this.IsForm())
		oDocPart.MakeSingleParagraphContent();

	oDocPart.SelectAll();

	var oParagraph = oDocPart.GetFirstParagraph();
	oParagraph.CorrectContent();

	var oRun = null;
	if (this.IsForm())
		oRun = oParagraph.MakeSingleRunParagraph();

	oParagraph.SetDirectParaPr(oParaPr);
	oParagraph.SetDirectTextPr(oTextPr);

	if (oRun)
		oRun.AddText(sText);
	else
		oDocPart.AddText(sText);

	oDocPart.RemoveSelection();

	var isPlaceHolder = this.IsPlaceHolder();
	if (isPlaceHolder && this.IsPicture())
		this.private_UpdatePictureContent();
	else if (isPlaceHolder)
		this.private_FillPlaceholderContent();

	return oDocPart;
};
/**
 * Выставляем параметр, что данный контрол должен быть простым текстовым
 * @param {boolean} isText
 */
CSdtBase.prototype.SetContentControlText = function(isText)
{
	if (this.Pr.Text !== isText)
	{
		History.Add(new CChangesSdtPrText(this, this.Pr.Text, isText));
		this.Pr.Text = isText;
	}
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsContentControlText = function()
{
	// Временно отключаем запись этого флага, т.к. мы пока его не обрабатываем в редакторе вообще
	// и можем создать некорректный файл для других редакторов. См. баг #51589
	return false;
	return this.Pr.Text;
};
/**
 * Выставляем параметр, что данный контрол специальный для вставки формул
 * @param {boolean} isEquation
 */
CSdtBase.prototype.SetContentControlEquation = function(isEquation)
{
	if (this.Pr.Equation !== isEquation)
	{
		History.Add(new CChangesSdtPrEquation(this, this.Pr.Equation, isEquation));
		this.Pr.Equation = isEquation;
	}
};
/**
 * Применяем настройку, что данный контрол будет специальным для формул
 * @constructor
 */
CSdtBase.prototype.ApplyContentControlEquationPr = function()
{
	let textPr = this.GetDefaultTextPr().Copy();
	textPr.SetItalic(true);
	textPr.SetFontFamily("Cambria Math");

	this.SetDefaultTextPr(textPr);
	this.SetContentControlEquation(true);
	this.SetContentControlTemporary(true);

	if (!this.IsContentControlEquation())
		return;

	this.SetPlaceholder(c_oAscDefaultPlaceholderName.Equation);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsContentControlEquation = function()
{
	return this.Pr.Equation;
};
/**
 * Выставляем параметр, что данный контрол при редактировании должен быть удален
 * @param {boolean} isTemporary
 */
CSdtBase.prototype.SetContentControlTemporary = function(isTemporary)
{
	if (this.Pr.Temporary !== isTemporary)
	{
		History.Add(new CChangesSdtPrTemporary(this, this.Pr.Temporary, isTemporary));
		this.Pr.Temporary = isTemporary;
	}
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsContentControlTemporary = function()
{
	return this.Pr.Temporary;
};
/**
 * @param {AscWord.CSdtFormPr} oFormPr
 */
CSdtBase.prototype.SetFormPr = function(oFormPr)
{
	this.private_CheckFieldMasterBeforeSet(oFormPr);
	this.private_CheckKeyValueBeforeSet(oFormPr);

	if ((!this.Pr.FormPr && oFormPr) || !this.Pr.FormPr.IsEqual(oFormPr))
	{
		let change = new CChangesSdtPrFormPr(this, this.Pr.FormPr, oFormPr);
		AscCommon.History.Add(change);
		change.Redo();

		this.private_OnAddFormPr();
	}
};
CSdtBase.prototype.private_CheckKeyValueBeforeSet = function(formPr)
{
	if (!this.Pr.FormPr || !formPr)
		return;

	let newKey = formPr.GetKey();
	if (!newKey)
		newKey = "";

	newKey = newKey.trim();

	if ("" === newKey)
		formPr.SetKey(this.Pr.FormPr.GetKey());
	else
		formPr.SetKey(newKey);
};
CSdtBase.prototype.private_CheckFieldMasterBeforeSet = function(formPr)
{
	if (!formPr || !formPr.GetRole())
		return;
	
	let roleName = formPr.GetRole();
	if (!roleName)
		return;
	
	// Настройки formPr могут прийти в интерфейс с заполненным fieldMaster, если в интерфейсе меняется роль, значит
	// она будет здесь выставлена и имеет больший приоритет, чем выставленный fieldMaster
	
	formPr.SetFieldMaster(null);
	formPr.SetRole(null);
	
	let logicDocument = this.GetLogicDocument();
	let oform;
	
	if (!logicDocument
		|| !AscCommon.IsSupportOFormFeature()
		|| !(oform = logicDocument.GetOFormDocument())
		|| !logicDocument.IsActionStarted())
		return;
	
	let role = oform.getRole(roleName);
	let userMaster;
	if (!role || !(userMaster = role.getUserMaster()))
		return;
	
	let fieldMaster;
	if (this.Pr.FormPr && this.Pr.FormPr.GetFieldMaster())
		fieldMaster = this.Pr.FormPr.GetFieldMaster();
	
	if (!fieldMaster)
		fieldMaster = oform.getFormat().createFieldMaster();
	else
		oform.getFormat().addFieldMaster(fieldMaster);
	
	if (1 !== fieldMaster.getUserCount() || userMaster !== fieldMaster.getUser(0))
	{
		fieldMaster.clearUsers();
		fieldMaster.addUser(userMaster);
	}
	
	fieldMaster.setLogicField(this);
	formPr.SetFieldMaster(fieldMaster);
};
/**
 * Удаляем настройки специальных форм
 */
CSdtBase.prototype.RemoveFormPr = function()
{
	if (this.Pr.FormPr)
	{
		History.Add(new CChangesSdtPrFormPr(this, this.Pr.FormPr, undefined));
		this.Pr.FormPr = undefined;

		var oLogicDocument = this.GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.GetFormsManager().Unregister(this);

		this.private_OnAddFormPr();
	}
};
/**
 * @returns {?AscWord.CSdtFormPr}
 */
CSdtBase.prototype.GetFormPr = function()
{
	return this.Pr.FormPr;
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsForm = function()
{
	return (undefined !== this.Pr.FormPr);
};
CSdtBase.prototype.IsComplexForm = function()
{
	return (undefined !== this.Pr.ComplexFormPr);
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsFixedForm = function()
{
	return false;
};
/**
 * returns {boolean}
 */
CSdtBase.prototype.IsFormRequired = function()
{
	return (this.Pr.FormPr ? this.Pr.FormPr.GetRequired() : false);
};
/**
 * Устанавливаем флаг Required
 * @param {boolean} isRequired
 */
CSdtBase.prototype.SetFormRequired = function(isRequired)
{
	var oFormPr = this.GetFormPr();
	if (oFormPr && isRequired !== oFormPr.GetRequired())
	{
		var oNewPr = oFormPr.Copy();
		oNewPr.SetRequired(isRequired);
		this.SetFormPr(oNewPr);
	}
};
/**
 * Получаем ключ для специальной формы, если он задан
 * @returns {?string}
 */
CSdtBase.prototype.GetFormKey = function()
{
	if (!this.IsForm())
		return undefined;

	return (this.Pr.FormPr.Key);
};
/**
 * Задаем новый ключ для специальной формы
 * @param key {string}
 */
CSdtBase.prototype.SetFormKey = function(key)
{
	if (this.GetFormKey() === key)
		return;
	
	let formPr = this.GetFormPr();
	if (!formPr)
		return;
	
	formPr = formPr.Copy();
	formPr.Key = key;
	this.SetFormPr(formPr);
};
/**
 * Проверяем, является ли данный контейнер чекбоксом
 * @returns {boolean}
 */
CSdtBase.prototype.IsCheckBox = function()
{
	return false;
};
/**
 * Проверяем, является ли заданный контрол радио-кнопкой
 * @returns {boolean}
 */
CSdtBase.prototype.IsRadioButton = function()
{
	return !!(this.IsCheckBox() && this.Pr.CheckBox && this.Pr.CheckBox.GroupKey);
};
/**
 * Является ли данный контейнер специальной текстовой формой
 * @returns {boolean}
 */
CSdtBase.prototype.IsTextForm = function()
{
	return false;
};
/**
 * Получаем ключ для группы радио-кнопок
 * @returns {?string}
 */
CSdtBase.prototype.GetRadioButtonGroupKey = function()
{
	if (!this.IsRadioButton())
		return undefined;

	return (this.Pr.CheckBox.GroupKey);
};
/**
 * Задаем групповой ключ для радио кнопки
 * @param {string }groupKey
 */
CSdtBase.prototype.SetRadioButtonGroupKey = function(groupKey)
{
	let checkBoxPr = this.Pr.CheckBox;
	if (!this.IsRadioButton() || !checkBoxPr)
		return;
	
	checkBoxPr = checkBoxPr.Copy();
	checkBoxPr.SetGroupKey(groupKey);
	this.SetCheckBoxPr(checkBoxPr);
};
/**
 * Для чекбоксов и радио-кнопок получаем состояние
 * @returns {bool}
 */
CSdtBase.prototype.IsCheckBoxChecked = function()
{
	if (this.IsCheckBox())
		return this.Pr.CheckBox.Checked;

	return false;
};
/**
 * Копируем placeholder
 * @return {string}
 */
CSdtBase.prototype.private_CopyPlaceholder = function(oPr)
{
	var oLogicDocument = this.GetLogicDocument();
	if (!oLogicDocument || !this.Pr.Placeholder)
		return;

	var oGlossary = oLogicDocument.GetGlossaryDocument();
	var oDocPart  = oGlossary.GetDocPartByName(this.Pr.Placeholder);
	if (!oDocPart)
		return;

	if (oGlossary.IsDefaultDocPart(oDocPart))
	{
		return this.Pr.Placeholder;
	}
	else
	{
		var sCopyName;
		if(oPr && oPr.Comparison && oPr.Comparison.originalDocument) 
		{
			var oPrGlossary = oPr.Comparison.originalDocument.GetGlossaryDocument();
			sCopyName = oPrGlossary.GetNewName();
			oDocPart.Glossary = oPrGlossary;
			oGlossary.AddDocPart(oDocPart.Copy(sCopyName));
			oDocPart.Glossary = oGlossary;
			return sCopyName;
		}
		else 
		{
			sCopyName = oGlossary.GetNewName();
			oGlossary.AddDocPart(oDocPart.Copy(sCopyName));
			return sCopyName;
		}
	}
};
/**
 * Проверяем является ли данный контрол текущим
 * @return {boolean}
 */
CSdtBase.prototype.IsCurrent = function()
{
	return this.Current;
};
/**
 * Выставляем, является ли данный контрол текущим
 * @param {boolean} isCurrent
 */
CSdtBase.prototype.SetCurrent = function(isCurrent)
{
	this.Current = isCurrent;
	
	if (this.IsForm() && this.IsFixedForm())
	{
		let logicDocument   = this.GetLogicDocument();
		let drawingDocument = logicDocument ? logicDocument.GetDrawingDocument() : null;
		if (drawingDocument && !logicDocument.IsFillingOFormMode())
		{
			drawingDocument.OnDrawContentControl(null, AscCommon.ContentControlTrack.In);
			drawingDocument.OnDrawContentControl(null, AscCommon.ContentControlTrack.Hover);
		}
	}
};
/**
 * Специальная функция, которая обновляет текстовые настройки у плейсхолдера для форм
 */
CSdtBase.prototype.UpdatePlaceHolderTextPrForForm = function()
{
};
/**
 * Проверяем попадание в контейнер
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {boolean}
 */
CSdtBase.prototype.CheckHitInContentControlByXY = function(X, Y, nPageAbs)
{
	return false;
};
/**
 * Ищем ближаюшую позицию, которая попадала бы в контейнер
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {?{X:number,Y:number}}
 */
CSdtBase.prototype.CorrectXYToHitIn = function(X, Y, nPageAbs)
{
	return null;
};
/**
 * Расширенное очищение контрола, с учетом типа контрола
 */
CSdtBase.prototype.ClearContentControlExt = function()
{
	if (this.IsComplexForm())
	{
		let arrSubForms = this.GetAllChildForms();
		for (let nIndex = 0, nCount = arrSubForms.length; nIndex < nCount; ++nIndex)
		{
			arrSubForms[nIndex].ClearContentControlExt();
		}
	}
	else if (this.IsCheckBox())
	{
		this.SetCheckBoxChecked(false);
	}
	else if (this.IsPicture())
	{
		if (this.IsPlaceHolder())
			return;

		var arrDrawings = this.GetAllDrawingObjects();
		if (arrDrawings.length > 0 && arrDrawings[0].IsPicture())
		{
			var nW = arrDrawings[0].Extent.W;
			var nH = arrDrawings[0].Extent.H;

			this.ReplaceContentWithPlaceHolder();
			this.ApplyPicturePr(true, nW, nH);
		}
		else
		{
			this.ReplaceContentWithPlaceHolder();
			this.ApplyPicturePr(true);
		}
	}
	else
	{
		this.ReplaceContentWithPlaceHolder();
	}
};
/**
 * Проверяем правильно ли заполнена форма
 * @returns {boolean}
 */
CSdtBase.prototype.IsFormFilled = function()
{
	return true;
}
/**
 * Проверка заполненности формы для составных форм
 * @returns {boolean}
 */
CSdtBase.prototype.IsComplexFormFilled = function()
{
	let oMainForm = this.GetMainForm();
	if (!oMainForm)
		return false;

	let arrForms = oMainForm.GetAllSubForms();
	if (!arrForms.length)
		return true;

	for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
	{
		if (!arrForms[nIndex].IsFormFilled())
			return false;
	}

	return true;
};
/**
 * Оборачиваем форму в графический контейнер
 * @returns {?ParaDrawing}
 */
CSdtBase.prototype.ConvertFormToFixed = function()
{
	return null;
};
/**
 * Уладаляем графичейский контейнер у формы
 * @returns {?CSdtBase}
 */
CSdtBase.prototype.ConvertFormToInline = function()
{
	return this;
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsMultiLineForm = function()
{
	return true;
};
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsPictureForm = function()
{
	return false;
};
/**
 * Функция обновления картиночной формы
 */
CSdtBase.prototype.UpdatePictureFormLayout = function()
{
};
/**
 * Выставляем настройку, что заданный контрол является составным полем
 * @param oPr {AscWord.CSdtComplexFormPr}
 */
CSdtBase.prototype.SetComplexFormPr = function(oPr)
{
	if (!this.Pr.ComplexFormPr || !this.Pr.ComplexFormPr.IsEqual(oPr))
	{
		let _oPr    = oPr ? oPr.Copy() : undefined;
		let oChange = new CChangesSdtPrComplexFormPr(this, this.Pr.ComplexFormPr, _oPr);
		AscCommon.History.Add(oChange);
		oChange.Redo();
	}
};
/**
 * @returns {?AscWord.CSdtComplexFormPr}
 */
CSdtBase.prototype.GetComplexFormPr = function()
{
	return this.Pr.ComplexFormPr;
}
/**
 * Получаем главную родительскую сложную форму
 * @returns {?AscWord.CInlineLevelSdt}
 */
CSdtBase.prototype.GetMainComplexForm = function()
{
	let oMain = null;
	if (this instanceof AscWord.CInlineLevelSdt || this.IsComplexForm())
		oMain = this;

	let oCur  = this;
	while (true)
	{
		oCur = oCur.GetParent();
		if (!oCur || !(oCur instanceof AscWord.CInlineLevelSdt) || !oCur.IsComplexForm())
			break;

		oMain = oCur;
	}

	return oMain;
};
CSdtBase.prototype.GetMainForm = function()
{
	if (!this.IsForm())
		return null;

	let oMain = this.GetMainComplexForm();
	return oMain ? oMain : this;
};
/**
 * Данная функция возвращает все дочерние формы по отношению к данной форме
 * @returns {CSdtBase[]}
 */
CSdtBase.prototype.GetAllChildForms = function()
{
	let arrForms    = [];
	let arrControls = this.GetAllContentControls();
	for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
	{
		let oControl = arrControls[nIndex];
		if (oControl.IsForm())
			arrForms.push(oControl);
	}

	return arrForms;
};
/**
 * Получаем все простые подформы (т.е. если подформы - составное поле, то мы пробегаемся по её простым подформам)
 * @param arrForms
 * @returns {CSdtBase[]}
 */
CSdtBase.prototype.GetAllSubForms = function(arrForms)
{
	if (!arrForms)
		arrForms = [];

	let arrControls = this.GetAllContentControls();
	for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
	{
		let oControl = arrControls[nIndex];
		if (!oControl.IsForm())
			continue;

		if (!oControl.IsComplexForm())
			arrForms.push(oControl);
	}

	return arrForms;
};
/**
 * Получаем порядковый номер данного подполя в родительском сложном поле
 * Если данный объект не является полем или подполем в сложном поле, то вернется -1
 * @returns {number}
 */
CSdtBase.prototype.GetSubFormIndex = function()
{
	if (!this.IsForm())
		return -1;
	
	let mainForm = this.GetMainForm();
	if (this === mainForm)
		return -1;
	
	let subForms = mainForm.GetAllSubForms();
	for (let index = 0, count = subForms.length; index < count; ++index)
	{
		if (subForms[index] === this)
			return index;
	}
	
	return -1;
};
/**
 * Провяеряем является ли данная форма текущей, с учетом того, что она либо сама является составной формой, либо
 * лежит в составной
 * @returns {boolean}
 */
CSdtBase.prototype.IsCurrentComplexForm = function()
{
	// Текущая форма есть только в режиме заполнения. В режиме редактирования не даем заполнять форму
	let logicDocument = this.GetLogicDocument();
	if (logicDocument && logicDocument.IsDocumentEditor() && !logicDocument.IsFillingFormMode())
		return false;
	
	if (this.IsCurrent())
		return true;

	let oMainComplexForm = this.GetMainComplexForm();
	if (!oMainComplexForm)
		return false;

	if (oMainComplexForm.IsCurrent())
		return true;

	let arrForms = oMainComplexForm.GetAllChildForms();
	for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
	{
		if (arrForms[nIndex].IsCurrent())
			return true;
	}

	return false;
};
/**
 * Является ли данная форма основной (а не подформой внутри другой формы)
 * @returns {boolean}
 */
CSdtBase.prototype.IsMainForm = function()
{
	return (this === this.GetMainForm());
};
/**
 * Возвращаем следующую простую подформу в составе сложной формы
 * @returns {?CSdtBase}
 */
CSdtBase.prototype.GetNextSubForm = function()
{
	let oMainForm;
	if (!this.IsForm()
		|| this.IsComplexForm()
		|| !(oMainForm = this.GetMainComplexForm())
		|| oMainForm === this)
		return null;

	let arrForms = oMainForm.GetAllSubForms();
	if (!arrForms.length)
		return null;

	let nIndex = arrForms.indexOf(this);
	if (-1 === nIndex)
		return arrForms[0];

	return (nIndex < arrForms.length - 1 ? arrForms[nIndex + 1] : this);
};
/**
 * Возвращаем предыдущую простую подформу в составе сложной формы
 * @returns {?CSdtBase}
 */
CSdtBase.prototype.GetPrevSubForm = function()
{
	let oMainForm;
	if (!this.IsForm()
		|| this.IsComplexForm()
		|| !(oMainForm = this.GetMainComplexForm())
		|| oMainForm === this)
		return null;

	let arrForms = oMainForm.GetAllSubForms();
	if (!arrForms.length)
		return null;

	let nIndex = arrForms.indexOf(this);
	if (-1 === nIndex)
		return arrForms[arrForms.length - 1];

	return (nIndex > 0 ? arrForms[nIndex - 1] : this);
};
CSdtBase.prototype.GetSubFormFromCurrentPosition = function(isForward)
{
	let oMainForm;
	if (!this.IsForm() || !(oMainForm = this.GetMainComplexForm()))
		return null;

	let arrForms = oMainForm.GetAllSubForms();
	if (!arrForms.length)
		return null;

	if (!this.IsComplexForm())
		return this;

	let nCurPos = this.State.ContentPos;
	if (isForward)
	{
		for (let nPos = nCurPos + 1, nCount = this.Content.length; nPos < nCount; ++nPos)
		{
			let oElement = this.GetElement(nPos);
			if (oElement instanceof AscWord.CInlineLevelSdt && oElement.IsForm())
			{
				if (!oElement.IsComplexForm())
					return oElement;

				let arrSubForms = oElement.GetAllSubForms();
				if (arrSubForms.length)
					return arrSubForms[0];
			}
		}
	}
	else
	{
		for (let nPos = nCurPos - 1; nPos >= 0; --nPos)
		{
			let oElement = this.GetElement(nPos);
			if (oElement instanceof AscWord.CInlineLevelSdt && oElement.IsForm())
			{
				if (!oElement.IsComplexForm())
					return oElement;

				let arrSubForms = oElement.GetAllSubForms();
				if (arrSubForms.length)
					return arrSubForms[arrSubForms.length - 1];
			}
		}
	}

	let oParent = this.GetParent();
	if (this === oMainForm
		|| !oParent
		|| !(oParent instanceof AscWord.CInlineLevelSdt)
		|| !oParent.IsForm())
		return null;

	return oParent.GetSubFormFromCurrentPosition(isForward);
};
CSdtBase.prototype.IsBuiltInTableOfContents = function()
{
	return (this.Pr && this.Pr.DocPartObj && this.Pr.DocPartObj.Gallery === "Table of Contents");
};
CSdtBase.prototype.IsBuiltInWatermark = function()
{
	return (this.Pr && this.Pr.DocPartObj && (this.Pr.DocPartObj.Gallery === "Watermarks" || this.Pr.DocPartObj.Gallery === "Watermark"));
};
CSdtBase.prototype.IsBuiltInUnique = function()
{
	return (this.Pr && this.Pr.DocPartObj && true === this.Pr.DocPartObj.Unique);
};
CSdtBase.prototype.GetInnerText = function()
{
	return "";
};
CSdtBase.prototype.GetFormValue = function()
{
	if (!this.IsForm())
		return null;

	if (this.IsPlaceHolder())
		return this.IsCheckBox() ? false : "";

	if (this.IsComplexForm())
	{
		return this.GetInnerText();
	}
	else if (this.IsCheckBox())
	{
		return this.IsCheckBoxChecked();
	}
	else if (this.IsPictureForm())
	{
		let oImg;
		let allDrawings = this.GetAllDrawingObjects();
		for (let nDrawing = 0; nDrawing < allDrawings.length; ++nDrawing)
		{
			if (allDrawings[nDrawing].IsPicture())
			{
				oImg = allDrawings[nDrawing].GraphicObj;
				break;
			}
		}

		return oImg ? oImg.getBase64Img() : "";
	}

	return this.GetInnerText();
};
CSdtBase.prototype.MoveCursorOutsideForm = function(isBefore)
{
};
CSdtBase.prototype.GetFieldMaster = function()
{
	let formPr = this.GetFormPr();
	if (!formPr)
		return null;
	
	return formPr.GetFieldMaster();
};
/**
 * Получаем название роли данного поля
 * @returns {string}
 */
CSdtBase.prototype.GetFormRole = function()
{
	let fieldMaster = this.GetFieldMaster();
	let userMaster  = fieldMaster ? fieldMaster.getFirstUser() : null;
	return userMaster ? userMaster.getRole() : "";
};
CSdtBase.prototype.SetFormRole = function(roleName)
{
	if (!this.IsForm() || roleName === this.GetFormRole())
		return;
	
	let formPr = this.GetFormPr().Copy();
	formPr.SetRole(roleName);
	this.SetFormPr(formPr);
};
CSdtBase.prototype.SetFieldMaster = function(fieldMaster)
{
	if (!fieldMaster)
		return;
	
	let formPr = this.GetFormPr();
	if (!formPr)
		return;
	
	let newFormPr = formPr.Copy();
	newFormPr.Field = fieldMaster;
	this.SetFormPr(newFormPr);
};
CSdtBase.prototype.GetFormShd = function()
{
	let formPr = this.GetFormPr();
	if (!formPr)
		return null;
	
	return formPr.GetShd();
};
CSdtBase.prototype.GetFormHighlightColor = function(defaultColor)
{
	if (undefined === defaultColor)
	{
		let logicDocument = this.GetLogicDocument();
		defaultColor = logicDocument && logicDocument.GetSpecialFormsHighlight ? logicDocument.GetSpecialFormsHighlight() : null;
	}
	
	let logicDocument = this.GetLogicDocument();
	
	if (!logicDocument || !this.IsForm())
		return defaultColor;
	
	let formPr = this.GetFormPr();
	if (!this.IsMainForm())
	{
		let mainForm = this.GetMainForm();
		if (!mainForm)
			return defaultColor;
		
		formPr = mainForm.GetFormPr();
	}
	
	if (!formPr)
		return defaultColor;
	
	let fieldMaster = formPr.GetFieldMaster();
	let userMaster  = fieldMaster ? fieldMaster.getFirstUser() : null;
	let userColor   = userMaster ? userMaster.getColor() : null;
	
	let oform         = logicDocument ? logicDocument.GetOFormDocument() : null;
	let currentUser   = oform ? oform.getCurrentUserMaster() : null;
	
	if (!currentUser || currentUser === userMaster)
		return userColor ? userColor : defaultColor;
	
	return new AscWord.CDocumentColor(0xF2, 0xF2, 0xF2);
};
CSdtBase.prototype.CheckOFormUserMaster = function()
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument)
		return true;
	
	return logicDocument.CheckOFormUserMaster(this);
};
/**
 * Проверяем, можно ли ставить курсор внутрь
 */
CSdtBase.prototype.CanPlaceCursorInside = function()
{
	let logicDocument = this.GetLogicDocument();
	return (!this.IsPicture() && (!this.IsForm() || this.IsComplexForm() || !logicDocument || !logicDocument.IsDocumentEditor() || logicDocument.IsFillingFormMode()))
};
CSdtBase.prototype.SkipFillingFormModeCheck = function(isSkip)
{
};
