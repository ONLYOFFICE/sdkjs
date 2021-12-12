/*
 * (c) Copyright Ascensio System SIA 2010-2020
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
	var oTextPr = new CTextPr();
	oTextPr.SetItalic(true);
	oTextPr.SetFontFamily("Cambria Math");

	this.SetDefaultTextPr(oTextPr);
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
 * @param {CSdtFormPr} oFormPr
 */
CSdtBase.prototype.SetFormPr = function(oFormPr)
{
	if ((!this.Pr.FormPr && oFormPr) || !this.Pr.FormPr.IsEqual(oFormPr))
	{
		History.Add(new CChangesSdtPrFormPr(this, this.Pr.FormPr, oFormPr));
		this.Pr.FormPr = oFormPr;

		var oLogicDocument = this.GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.RegisterForm(this);

		this.private_OnAddFormPr();
	}
}
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
			oLogicDocument.UnregisterForm(this);

		this.private_OnAddFormPr();
	}
};
/**
 * @returns {?CSdtFormPr}
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
	if (this.IsCheckBox())
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
