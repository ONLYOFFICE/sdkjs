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
	oDocPart.SelectAll();

	var oParagraph = oDocPart.GetFirstParagraph();
	oParagraph.CorrectContent();
	oParagraph.SetDirectParaPr(oParaPr);
	oParagraph.SetDirectTextPr(oTextPr);

	oDocPart.AddText(sText);
	oDocPart.RemoveSelection();

	if (this.IsPlaceHolder())
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
 *
 * @param {CSdtFormPr} oFormPr
 */
CSdtBase.prototype.SetFormPr = function(oFormPr)
{
	if ((!this.Pr.FormPr && oFormPr) || this.Pr.FormPr.Is)
	{
		History.Add(new CChangesSdtPrFormPr(this, this.Pr.FormPr, oFormPr));
		this.Pr.FormPr = oFormPr;

		var oLogicDocument = this.GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.RegisterForm(this);
	}
}
/**
 * @returns {boolean}
 */
CSdtBase.prototype.IsForm = function()
{
	return (undefined !== this.Pr.FormPr);
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
