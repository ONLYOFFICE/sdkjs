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

function CSdtPr()
{
	this.Alias = undefined;
	this.Id    = undefined;
	this.Tag   = undefined;
	this.Label = undefined;
	this.Lock  = undefined;
	
	this.DataBinding = undefined;

	this.DocPartObj = {
		Gallery  : undefined,
		Category : undefined,
		Unique   : undefined
	};

	this.Appearance = Asc.c_oAscSdtAppearance.Frame;
	this.Color      = undefined;

	this.CheckBox = undefined;
	this.Picture  = false;
	this.ComboBox = undefined;
	this.DropDown = undefined;
	this.Date     = undefined;
	this.Equation = false;
	this.TextForm = undefined;

	this.TextPr = new CTextPr();

	this.Placeholder   = undefined;
	this.ShowingPlcHdr = false;

	this.Text      = false;
	this.Temporary = false;

	this.FormPr        = undefined;
	this.PictureFormPr = undefined;
	this.ComplexFormPr = undefined;

	this.OForm         = undefined;
	
	this.BorderColor = undefined;
	this.ShdColor    = undefined;
}

CSdtPr.prototype.Copy = function()
{
	var oPr = new CSdtPr();

	oPr.Alias      = this.Alias;
	oPr.Id         = this.Id;
	oPr.Tag        = this.Tag;
	oPr.Label      = this.Label;
	oPr.Lock       = this.Lock;
	oPr.Appearance = this.Appearance;
	oPr.Color      = (this.Color ? this.Color.Copy() : undefined);
	
	oPr.BorderColor = this.BorderColor ? this.BorderColor.Copy() : undefined;
	oPr.ShdColor    = this.ShdColor ? this.ShdColor.Copy() : undefined;

	if (this.DataBinding)
		oPr.DataBinding = this.DataBinding.copy();

	if (this.CheckBox)
		oPr.CheckBox = this.CheckBox.Copy();

	oPr.Picture = this.Picture;

	if (this.ComboBox)
		oPr.ComboBox = this.ComboBox.Copy();

	if (this.DropDown)
		oPr.DropDown = this.DropDown.Copy();

	if (this.Date)
		oPr.Date = this.Date.Copy();

	if (this.TextForm)
		oPr.TextForm = this.TextForm.Copy();

	if (this.PictureFormPr)
		oPr.PictureFormPr = this.PictureFormPr.Copy();

	if (this.ComplexFormPr)
		oPr.ComplexFormPr = this.ComplexFormPr.Copy();

	oPr.TextPr = this.TextPr.Copy();

	oPr.Placeholder   = this.Placeholder;
	oPr.ShowingPlcHdr = this.ShowingPlcHdr;

	oPr.Equation  = this.Equation;
	oPr.Text      = this.Text;
	oPr.Temporary = this.Temporary;

	return oPr;
};
CSdtPr.prototype.Write_ToBinary = function(Writer)
{
	this.TextPr.WriteToBinary(Writer);

	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);
	var Flags = 0;

	if (undefined !== this.Alias)
	{
		Writer.WriteString2(this.Alias);
		Flags |= 1;
	}

	if (undefined !== this.Id)
	{
		Writer.WriteLong(this.Id);
		Flags |= 2;
	}

	if (undefined !== this.Tag)
	{
		Writer.WriteString2(this.Tag);
		Flags |= 4;
	}

	if (undefined !== this.Label)
	{
		Writer.WriteLong(this.Tag);
		Flags |= 8;
	}

	if (undefined !== this.Lock)
	{
		Writer.WriteLong(this.Lock);
		Flags |= 16;
	}

	if (undefined !== this.DocPartObj.Unique)
	{
		Writer.WriteBool(this.DocPartObj.Unique);
		Flags |= 32;
	}

	if (undefined !== this.DocPartObj.Gallery)
	{
		Writer.WriteString2(this.DocPartObj.Gallery);
		Flags |= 64;
	}

	if (undefined !== this.DocPartObj.Category)
	{
		Writer.WriteString2(this.DocPartObj.Category);
		Flags |= 128;
	}

	if (undefined !== this.Appearance)
	{
		Writer.WriteLong(this.Appearance);
		Flags |= 256;
	}

	if (undefined !== this.Color)
	{
		this.Color.WriteToBinary(Writer);
		Flags |= 512;
	}

	if (undefined !== this.CheckBox)
	{
		this.CheckBox.WriteToBinary(Writer);
		Flags |= 1024;
	}

	if (undefined !== this.Picture)
	{
		Writer.WriteBool(this.Picture);
		Flags |= 2048;
	}

	if (undefined !== this.ComboBox)
	{
		this.ComboBox.WriteToBinary(Writer);
		Flags |= 4096;
	}

	if (undefined !== this.DropDown)
	{
		this.DropDown.WriteToBinary(Writer);
		Flags |= 8192;
	}

	if (undefined !== this.Date)
	{
		this.Date.WriteToBinary(Writer);
		Flags |= 16384;
	}

	if (undefined !== this.Placeholder)
	{
		Writer.WriteString2(this.Placeholder);
		Flags |= 32768;
	}

	if (undefined !== this.ShowingPlcHdr)
	{
		Writer.WriteBool(this.ShowingPlcHdr);
		Flags |= 65536;
	}

	if (undefined !== this.Equation)
	{
		Writer.WriteBool(this.Equation);
		Flags |= 131072;
	}

	if (undefined !== this.Text)
	{
		Writer.WriteBool(this.Text);
		Flags |= 262144;
	}

	if (undefined !== this.Temporary)
	{
		Writer.WriteBool(this.Temporary);
		Flags |= 524288;
	}

	if (undefined !== this.TextForm)
	{
		this.TextForm.WriteToBinary(Writer);
		Flags |= 1048576;
	}

	if (undefined !== this.PictureFormPr)
	{
		this.PictureFormPr.WriteToBinary(Writer);
		Flags |= 2097152;
	}

	if (this.ComplexFormPr)
	{
		this.ComplexFormPr.WriteToBinary(Writer);
		Flags |= (1 << 22);
	}
	
	if (this.DataBinding)
	{
		this.DataBinding.toBinary(Writer);
		Flags |= (1 << 23);
	}
	
	if (this.BorderColor)
	{
		this.BorderColor.toBinary(Writer);
		Flags |= (1 << 24);
	}
	
	if (this.ShdColor)
	{
		this.ShdColor.toBinary(Writer);
		Flags |= (1 << 25);
	}
	
	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CSdtPr.prototype.Read_FromBinary = function(Reader)
{
	this.TextPr = new CTextPr();
	this.TextPr.ReadFromBinary(Reader);

	var Flags = Reader.GetLong();

	if (Flags & 1)
		this.Alias = Reader.GetString2();

	if (Flags & 2)
		this.Id = Reader.GetLong();

	if (Flags & 4)
		this.Tag = Reader.GetString2();

	if (Flags & 8)
		this.Tag = Reader.GetLong();

	if (Flags & 16)
		this.Lock = Reader.GetLong();

	if (Flags & 32)
		this.DocPartObj.Unique = Reader.GetBool();

	if (Flags & 64)
		this.DocPartObj.Gallery = Reader.GetString2();

	if (Flags & 128)
		this.DocPartObj.Category = Reader.GetString2();

	if (Flags & 256)
		this.Appearance = Reader.GetLong();

	if (Flags & 512)
	{
		this.Color = new CDocumentColor();
		this.Color.ReadFromBinary(Reader);
	}

	if (Flags & 1024)
	{
		this.CheckBox = new AscWord.CSdtCheckBoxPr();
		this.CheckBox.ReadFromBinary(Reader);
	}

	if (Flags & 2048)
		this.Picture = Reader.GetBool();

	if (Flags & 4096)
	{
		this.ComboBox = new AscWord.CSdtComboBoxPr();
		this.ComboBox.ReadFromBinary(Reader);
	}

	if (Flags & 8192)
	{
		this.DropDown = new AscWord.CSdtComboBoxPr();
		this.DropDown.ReadFromBinary(Reader);
	}

	if (Flags & 16384)
	{
		this.Date = new AscWord.CSdtDatePickerPr();
		this.Date.ReadToBinary(Reader);
	}

	if (Flags & 32768)
		this.Placeholder = Reader.GetString2();

	if (Flags & 65536)
		this.ShowingPlcHdr = Reader.GetBool();

	if (Flags & 131072)
		this.Equation = Reader.GetBool();

	if (Flags & 262144)
		this.Text = Reader.GetBool();

	if (Flags & 524288)
		this.Temporary = Reader.GetBool();

	if (Flags & 1048576)
	{
		this.TextForm = new AscWord.CSdtTextFormPr();
		this.TextForm.ReadFromBinary(Reader);
	}

	if (Flags & 2097152)
	{
		this.PictureFormPr = new AscWord.CSdtPictureFormPr();
		this.PictureFormPr.ReadFromBinary(Reader);
	}

	if (Flags & (1 << 22))
	{
		this.ComplexFormPr = new AscWord.CSdtComplexFormPr();
		this.ComplexFormPr.ReadFromBinary(Reader);
	}
	
	if (Flags & (1 << 23))
		this.DataBinding = AscWord.DataBinding.fromBinary(Reader);
	
	if (Flags & (1 << 24))
		this.BorderColor = AscWord.CDocumentColorA.fromBinary(Reader);
	
	if (Flags & (1 << 25))
		this.ShdColor = AscWord.CDocumentColorA.fromBinary(Reader);
};
CSdtPr.prototype.IsBuiltInDocPart = function()
{
	if (this.DocPartObj && (this.DocPartObj.Category || this.DocPartObj.Gallery))
		return true;

	return false;
};
CSdtPr.prototype.GetDocPartGallery = function()
{
	return this.DocPartObj ? this.DocPartObj.Gallery : undefined;
};

function CContentControlPr(nType)
{
	this.CC         = null;
	this.Id         = undefined;
	this.Tag        = undefined;
	this.Alias      = undefined;
	this.Lock       = undefined;
	this.InternalId = undefined;
	this.CCType     = undefined !== nType ? nType : c_oAscSdtLevelType.Inline;
	
	this.Temporary  = undefined;
	this.DataBinding = undefined;
	
	// section property
	this.SectionBreak = undefined;
	this.PageSizeW    = undefined;
	this.PageSizeH    = undefined;
	this.Orient       = undefined;
	
	// Margins 
	this.MarginT = undefined;
	this.MarginL = undefined;
	this.MarginR = undefined;
	this.MarginB = undefined;
	
	
	this.Appearance = Asc.c_oAscSdtAppearance.Frame;
	this.Color      = undefined;
	
	this.CheckBoxPr    = undefined;
	this.ComboBoxPr    = undefined;
	this.DropDownPr    = undefined;
	this.DateTimePr    = undefined;
	this.TextFormPr    = undefined;
	this.PictureFormPr = undefined;
	this.ComplexFormPr = undefined;
	
	this.PlaceholderText = undefined;
	
	this.FormPr = undefined;
	
	this.BorderColor = undefined;
	this.ShdColor    = undefined;
}
CContentControlPr.prototype.GetEventObject = function()
{
	return {
		"Tag"        : this.Tag,
		"Id"         : this.Id,
		"Lock"       : this.Lock,
		"InternalId" : this.InternalId
	};
};
CContentControlPr.prototype.FillFromObject = function(oPr)
{
	if (undefined !== oPr.Id)
		this.Id = oPr.Id;

	if (undefined !== oPr.Tag)
		this.Tag = oPr.Tag;

	if (undefined !== oPr.Alias)
		this.Alias = oPr.Alias;

	if (undefined !== oPr.Lock)
		this.Lock  = oPr.Lock;

	if (undefined !== oPr.InternalId)
		this.InternalId = oPr.InternalId;

	if (undefined !== oPr.Appearance)
		this.Appearance = oPr.Appearance;

	if (undefined !== oPr.Color)
		this.Color = oPr.Color;

	if (undefined !== oPr.PlaceholderText)
		this.PlaceholderText = oPr.PlaceholderText;
	
	if (undefined !== oPr.ShdColor)
		this.ShdColor = AscWord.CDocumentColorA.fromObjectRgba(oPr.ShdColor);
	
	if (undefined !== oPr.BorderColor)
		this.BorderColor = AscWord.CDocumentColorA.fromObjectRgba(oPr.BorderColor);

	if (undefined !== oPr.DataBinding)
		this.DataBinding = AscWord.DataBinding.fromObject(oPr.DataBinding);
};
CContentControlPr.prototype.FillFromContentControl = function(oContentControl)
{
	if (!oContentControl)
		return;

	this.CC         = oContentControl;
	this.CCType     = oContentControl.IsBlockLevel() ? c_oAscSdtLevelType.Block : c_oAscSdtLevelType.Inline;
	this.Id         = oContentControl.Pr.Id;
	this.Lock       = oContentControl.Pr.Lock;
	this.InternalId = oContentControl.GetId();
	this.Tag        = oContentControl.GetTag();
	this.Alias      = oContentControl.GetAlias();
	this.Appearance = oContentControl.GetAppearance();
	this.Color      = oContentControl.GetColor();
	this.Temporary  = oContentControl.IsContentControlTemporary();

	if (oContentControl.Pr.DataBinding)
		this.DataBinding = oContentControl.GetDataBinding();

	if (oContentControl.IsCheckBox())
	{
		this.CheckBoxPr = oContentControl.GetCheckBoxPr().Copy();
		if (oContentControl.IsRadioButton())
			this.CheckBoxPr.SetChoiceName(oContentControl.GetFormKey());
	}
	else if (oContentControl.IsComboBox())
		this.ComboBoxPr = oContentControl.GetComboBoxPr().Copy();
	else if (oContentControl.IsDropDownList())
		this.DropDownPr = oContentControl.GetDropDownListPr().Copy();
	else if (oContentControl.IsDatePicker())
	{
		this.DateTimePr = oContentControl.GetDatePickerPr().Copy();
		if (oContentControl.GetInnerText() !== this.DateTimePr.ToString())
			this.DateTimePr.SetNullFullDate(true);
	}
	else if (oContentControl.IsTextForm())
		this.TextFormPr = oContentControl.GetTextFormPr().Copy();
	else if (oContentControl.IsPictureForm())
		this.PictureFormPr = oContentControl.GetPictureFormPr().Copy();
	else if (oContentControl.IsComplexForm())
		this.ComplexFormPr = oContentControl.GetComplexFormPr().Copy();

	this.PlaceholderText = oContentControl.GetPlaceholderText();

	if (oContentControl.IsForm())
	{
		let mainForm = oContentControl.IsMainForm() ? oContentControl : oContentControl.GetMainForm();
		
		this.FormPr = mainForm.GetFormPr().Copy();
		this.FormPr.SetFixed(mainForm.IsFixedForm());
		
		if (mainForm !== oContentControl)
		{
			let subFormPr = oContentControl.GetFormPr();
			this.FormPr.SetAscBorder(subFormPr.GetAscBorder());
			this.FormPr.SetShd(subFormPr.GetShd());
		}
		
		if (oContentControl.IsSignatureForm())
			this.FormPr.SetRequired(true);
	}
	
	if (oContentControl.getShdColor())
		this.ShdColor = oContentControl.getShdColor().getAscColor();
	
	if (oContentControl.getBorderColor())
		this.BorderColor = oContentControl.getBorderColor().getAscColor();
};
CContentControlPr.prototype.SetToContentControl = function(oContentControl)
{
	if (!oContentControl)
		return;

	if (undefined !== this.FormPr
		&& oContentControl.IsRadioButton()
		&& oContentControl.IsFormRequired() !== this.FormPr.GetRequired()
		&& oContentControl.GetLogicDocument())
	{
		oContentControl.GetLogicDocument().OnChangeRadioRequired(oContentControl.GetRadioButtonGroupKey(), this.FormPr.GetRequired());
	}
	
	if (undefined !== this.Tag)
		oContentControl.SetTag(this.Tag);
	
	if (undefined !== this.Id)
		oContentControl.SetContentControlId(this.Id);

	if (undefined !== this.Lock)
		oContentControl.SetContentControlLock(this.Lock);

	if (undefined !== this.Alias)
		oContentControl.SetAlias(this.Alias);

	if (undefined !== this.Appearance)
		oContentControl.SetAppearance(this.Appearance);

	// Тут может быть как CDocumentColor так и Asc.asc_CColor
	if (undefined !== this.Color)
	{
		if (!this.Color)
			oContentControl.SetColor(undefined);
		else
			oContentControl.SetColor(new CDocumentColor(this.Color.r, this.Color.g, this.Color.b));
	}
	
	if (undefined !== this.Temporary)
		oContentControl.SetContentControlTemporary(this.Temporary);

	if (undefined !== this.CheckBoxPr)
	{
		let prevGroupKey = oContentControl.GetCheckBoxPr() ? oContentControl.GetCheckBoxPr().GetGroupKey() : undefined;
		if (prevGroupKey !== this.CheckBoxPr.GroupKey && undefined !== this.CheckBoxPr.Checked)
			this.CheckBoxPr.Checked = false;
		
		let choiceName = this.CheckBoxPr.GetChoiceName(true);
		if (undefined !== choiceName)
		{
			oContentControl.SetFormKey(choiceName);
			if (this.FormPr)
				this.FormPr.SetKey(choiceName);
		}
		
		oContentControl.SetCheckBoxPr(this.CheckBoxPr);
		oContentControl.private_UpdateCheckBoxContent();
	}

	if (undefined !== this.ComboBoxPr)
		oContentControl.SetComboBoxPr(this.ComboBoxPr);

	if (undefined !== this.DropDownPr)
		oContentControl.SetDropDownListPr(this.DropDownPr);

	if (undefined !== this.DateTimePr)
	{
		let dateTimePr = this.DateTimePr;
		if (dateTimePr.IsNullFullDate())
		{
			dateTimePr = dateTimePr.Copy();
			dateTimePr.SetNullFullDate(false);
			dateTimePr.SetFullDate(oContentControl.GetDatePickerPr().GetFullDate());
		}
		
		oContentControl.ApplyDatePickerPr(dateTimePr);
	}

	if (undefined !== this.TextFormPr && oContentControl.IsInlineLevel())
	{
		let isCombChanged   = (!oContentControl.Pr.TextForm || this.TextFormPr.Comb !== oContentControl.Pr.TextForm.Comb);
		let isMaxChanged    = (!oContentControl.Pr.TextForm || this.TextFormPr.MaxCharacters !== oContentControl.Pr.TextForm.MaxCharacters);
		let isFormatChanged = (!oContentControl.Pr.TextForm || !this.TextFormPr.Format.IsEqual(oContentControl.Pr.TextForm.Format));

		if (oContentControl.IsFixedForm() && isCombChanged)
			oContentControl.UpdateFixedFormCombWidthByFormSize(this.TextFormPr);

		oContentControl.SetTextFormPr(this.TextFormPr);

		if (this.TextFormPr.Comb && (isCombChanged || isMaxChanged))
			oContentControl.TrimTextForm();
		else if (!this.TextFormPr.Comb && isCombChanged && oContentControl.IsPlaceHolder())
			oContentControl.ReplaceContentWithPlaceHolder(false, true);

		if (oContentControl.IsFixedForm() && oContentControl.IsMainForm() && !isCombChanged)
			oContentControl.UpdateFixedFormSizeByCombWidth();

		if (!this.TextFormPr.MultiLine)
			oContentControl.CorrectSingleLineFormContent();
		
		if (isFormatChanged
			&& !oContentControl.IsPlaceHolder()
			&& !oContentControl.Pr.TextForm.CheckFormat(oContentControl.GetInnerText(), true))
		{
			oContentControl.ClearContentControlExt();
		}
	}

	if (undefined !== this.PlaceholderText)
		oContentControl.SetPlaceholderText(this.PlaceholderText);

	this.SetFormPrToContentControl(oContentControl);

	if (undefined !== this.PictureFormPr && oContentControl.IsInlineLevel())
	{
		oContentControl.SetPictureFormPr(this.PictureFormPr);
		oContentControl.UpdatePictureFormLayout();
	}

	if (undefined !== this.ComplexFormPr)
		oContentControl.SetComplexFormPr(this.ComplexFormPr);
	
	if (this.ShdColor)
		oContentControl.setShdColor(AscWord.CDocumentColorA.fromObjectRgba(this.ShdColor));
	
	if (this.BorderColor)
		oContentControl.setBorderColor(AscWord.CDocumentColorA.fromObjectRgba(this.BorderColor));
};
CContentControlPr.prototype.SetFormPrToContentControl = function(contentControl)
{
	if (!this.FormPr)
		return;
	
	let formPr = this.FormPr;
	
	let newRole = formPr.GetRole();
	if (contentControl.IsForm() && !contentControl.IsMainForm())
	{
		// Ключ у подформы должен сохраняться
		let oldKey    = contentControl.GetFormKey();
		let newFormPr = formPr.Copy();
		newFormPr.SetKey(oldKey);
		contentControl.SetFormPr(newFormPr);

		contentControl = contentControl.GetMainForm();
		if (!contentControl)
			return;

		formPr = contentControl.GetFormPr().Copy();
		formPr.SetRole(newRole);
	}

	if (contentControl.IsComplexForm())
	{
		let subForms = contentControl.GetAllSubForms();
		for (let index = 0, count = subForms.length; index < count; ++index)
		{
			subForms[index].SetFormRole(newRole);
		}
	}
	
	let newKey = formPr.GetKey();
	let oldKey = contentControl.GetFormKey();
	
	let fieldMaster = contentControl.GetFieldMaster();
	let userMaster  = fieldMaster ? fieldMaster.getFirstUser() : null;
	let oldRole     = userMaster ? userMaster.getRole() : null;
	
	let isKeyChanged = newKey && newKey !== oldKey;
	let isRoleChanged = newRole && newRole !== oldRole;
	
	if (isKeyChanged && isRoleChanged)
	{
		// Такого не должно быть, ключ и роль не должны меняться одновременно, но если все же произошло,
		// то выставляем все как задано в настройках, не делая ничего дополнительного
	}
	else if (isKeyChanged)
	{
		this.OnSetKeyToForm(newKey, contentControl);
	}
	else if (isRoleChanged)
	{
		this.OnSetRoleToForm(newRole, contentControl);
	}
	
	contentControl.SetFormPr(formPr);
};
CContentControlPr.prototype.OnSetKeyToForm = function(newKey, form)
{
	let logicDocument = form.GetLogicDocument();
	if (!logicDocument)
		return;
	
	let formManager = logicDocument.GetFormsManager();
	let allForms = formManager.GetAllFormsByKey(newKey, form.GetSpecificType());
	let firstForm = null;
	for (let iForm = 0, nForms = allForms.length; iForm < nForms; ++iForm)
	{
		if (allForms[iForm] === form)
			continue;
		
		firstForm = allForms[iForm];
		
		// Напрямую у formManager не вызываем, т.к. еще может быть не выставлен ключ у текущей формы
		logicDocument.OnChangeForm(allForms[iForm]);
		break;
	}
	
	if (firstForm)
		form.SyncFormPrWithSameKey(firstForm);
	
	let role = formManager.GetRoleByKey(newKey, form.GetSpecificType());
	if (!role)
		return;
	
	this.FormPr.SetRole(role);
};
CContentControlPr.prototype.OnSetRoleToForm = function(newRole, form)
{
	let logicDocument = form.GetLogicDocument();
	if (!logicDocument)
		return;
	
	let formManager = logicDocument.GetFormsManager();
	
	let formKey = form.GetFormKey();
	if (!formKey || "" === formKey)
		return;
	
	let allForms = formManager.GetAllFormsByKey(formKey, form.GetSpecificType());
	for (let index = 0, formCount = allForms.length; index < formCount; ++index)
	{
		let curForm = allForms[index];
		if (curForm !== form)
		{
			let formPr = curForm.GetFormPr();
			if (!formPr)
				continue;
			
			let newFormPr = formPr.Copy();
			newFormPr.SetRole(newRole);
			newFormPr.SetFieldMaster(null);
			curForm.SetFormPr(newFormPr);
		}
	}
};
CContentControlPr.prototype.GetId = function()
{
	return this.Id;
};
CContentControlPr.prototype.SetId = function(Id)
{
	this.Id = Id;
};
CContentControlPr.prototype.GetTag = function()
{
	return this.Tag;
};
CContentControlPr.prototype.SetTag = function(sTag)
{
	this.Tag = sTag;
};
CContentControlPr.prototype.GetLock = function()
{
	return this.Lock;
};
CContentControlPr.prototype.SetLock = function(nLock)
{
	this.Lock = nLock;
};
CContentControlPr.prototype.GetInternalId = function()
{
	return this.InternalId;
};
CContentControlPr.prototype.GetContentControlType = function()
{
	return this.CCType;
};
CContentControlPr.prototype.GetAlias = function()
{
	return this.Alias;
};
CContentControlPr.prototype.SetAlias = function(sAlias)
{
	this.Alias = sAlias;
};
CContentControlPr.prototype.GetAppearance = function()
{
	return this.Appearance;
};
CContentControlPr.prototype.SetAppearance = function(nAppearance)
{
	this.Appearance = nAppearance;
};
CContentControlPr.prototype.GetColor = function()
{
	if (!this.Color)
		return null;

	return new Asc.asc_CColor(this.Color.r, this.Color.g, this.Color.b);
};
CContentControlPr.prototype.SetColor = function(r, g, b)
{
	if (undefined === r)
		this.Color = undefined;
	else if (null === r)
		this.Color = null;
	else
		this.Color = new CDocumentColor(r, g, b);
};
CContentControlPr.prototype.GetTemporary = function()
{
	return this.Temporary;
};
CContentControlPr.prototype.SetTemporary = function(isTemporary)
{
	this.Temporary = isTemporary;
};
CContentControlPr.prototype.GetSpecificType = function()
{
	if (this.CC)
		return this.CC.GetSpecificType();

	return Asc.c_oAscContentControlSpecificType.None;
};
CContentControlPr.prototype.GetCheckBoxPr = function()
{
	if (this.CC && this.CC.IsCheckBox())
		return this.CheckBoxPr;

	return null;
};
CContentControlPr.prototype.SetCheckBoxPr = function(oPr)
{
	this.CheckBoxPr = oPr;
};
CContentControlPr.prototype.GetComboBoxPr = function()
{
	if (this.CC && this.CC.IsComboBox())
		return this.ComboBoxPr;

	return null;
};
CContentControlPr.prototype.SetComboBoxPr = function(oPr)
{
	this.ComboBoxPr = oPr;
};
CContentControlPr.prototype.GetDropDownListPr = function()
{
	if (this.CC && this.CC.IsDropDownList())
		return this.DropDownPr;

	return null;
};
CContentControlPr.prototype.SetDropDownListPr = function(oPr)
{
	this.DropDownPr = oPr;
};
CContentControlPr.prototype.GetDateTimePr = function()
{
	if (this.CC && this.CC.IsDatePicker())
		return this.DateTimePr;

	return null;
};
CContentControlPr.prototype.SetDateTimePr = function(oPr)
{
	this.DateTimePr = oPr;
};
CContentControlPr.prototype.GetTextFormPr = function()
{
	if (this.CC && this.CC.IsTextForm())
		return this.TextFormPr;

	return null;
};
CContentControlPr.prototype.SetTextFormPr = function(oPr)
{
	this.TextFormPr = oPr;
};
CContentControlPr.prototype.GetPlaceholderText = function()
{
	return this.PlaceholderText;
};
CContentControlPr.prototype.SetPlaceholderText = function(sText)
{
	this.PlaceholderText = sText;
};
CContentControlPr.prototype.GetFormPr = function()
{
	if (this.CC && this.CC.IsForm())
		return this.FormPr;

	return null;
};
CContentControlPr.prototype.SetFormPr = function(oPr)
{
	this.FormPr = oPr;
};
CContentControlPr.prototype.SetPictureFormPr = function(oPr)
{
	this.PictureFormPr = oPr;
};
CContentControlPr.prototype.GetPictureFormPr = function()
{
	return this.PictureFormPr;
};
CContentControlPr.prototype.SetComplexFormPr = function(oPr)
{
	this.ComplexFormPr = oPr;
};
CContentControlPr.prototype.GetComplexFormPr = function()
{
	return this.ComplexFormPr;
};
CContentControlPr.prototype.GetNewKey = function()
{
	let logicDocument = editor && editor.WordControl && editor.WordControl.m_oLogicDocument;
	if (!logicDocument || !logicDocument.IsDocumentEditor())
		return "";

	let formManager  = logicDocument.GetFormsManager();
	let keyGenerator = formManager.GetKeyGenerator();

	return keyGenerator.GetNewKey(this.CC);
};
CContentControlPr.prototype.IsSignature = function()
{
	let pictPr = this.GetPictureFormPr();
	return (pictPr && pictPr.IsSignature());
};

//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord']        = window['AscCommonWord'] || {};
window['AscCommonWord'].CSdtPr = CSdtPr;

window['AscCommon'] = window['AscCommon'] || {};

window['AscCommon'].CContentControlPr    = CContentControlPr;
window['AscCommon']['CContentControlPr'] = CContentControlPr;

CContentControlPr.prototype['get_Id']                 = CContentControlPr.prototype.GetId;
CContentControlPr.prototype['put_Id']                 = CContentControlPr.prototype.SetId;
CContentControlPr.prototype['get_Tag']                = CContentControlPr.prototype.GetTag;
CContentControlPr.prototype['put_Tag']                = CContentControlPr.prototype.SetTag;
CContentControlPr.prototype['get_Lock']               = CContentControlPr.prototype.GetLock;
CContentControlPr.prototype['put_Lock']               = CContentControlPr.prototype.SetLock;
CContentControlPr.prototype['get_InternalId']         = CContentControlPr.prototype.GetInternalId;
CContentControlPr.prototype['get_ContentControlType'] = CContentControlPr.prototype.GetContentControlType;
CContentControlPr.prototype['get_Alias']              = CContentControlPr.prototype.GetAlias;
CContentControlPr.prototype['put_Alias']              = CContentControlPr.prototype.SetAlias;
CContentControlPr.prototype['get_Appearance']         = CContentControlPr.prototype.GetAppearance;
CContentControlPr.prototype['put_Appearance']         = CContentControlPr.prototype.SetAppearance;
CContentControlPr.prototype['get_Color']              = CContentControlPr.prototype.GetColor;
CContentControlPr.prototype['put_Color']              = CContentControlPr.prototype.SetColor;
CContentControlPr.prototype['get_Temporary']          = CContentControlPr.prototype.GetTemporary;
CContentControlPr.prototype['put_Temporary']          = CContentControlPr.prototype.SetTemporary;
CContentControlPr.prototype['get_SpecificType']       = CContentControlPr.prototype.GetSpecificType;
CContentControlPr.prototype['get_CheckBoxPr']         = CContentControlPr.prototype.GetCheckBoxPr;
CContentControlPr.prototype['put_CheckBoxPr']         = CContentControlPr.prototype.SetCheckBoxPr;
CContentControlPr.prototype['get_ComboBoxPr']         = CContentControlPr.prototype.GetComboBoxPr;
CContentControlPr.prototype['put_ComboBoxPr']         = CContentControlPr.prototype.SetComboBoxPr;
CContentControlPr.prototype['get_DropDownListPr']     = CContentControlPr.prototype.GetDropDownListPr;
CContentControlPr.prototype['put_DropDownListPr']     = CContentControlPr.prototype.SetDropDownListPr;
CContentControlPr.prototype['get_DateTimePr']         = CContentControlPr.prototype.GetDateTimePr;
CContentControlPr.prototype['put_DateTimePr']         = CContentControlPr.prototype.SetDateTimePr;
CContentControlPr.prototype['get_TextFormPr']         = CContentControlPr.prototype.GetTextFormPr;
CContentControlPr.prototype['put_TextFormPr']         = CContentControlPr.prototype.SetTextFormPr;
CContentControlPr.prototype['get_PlaceholderText']    = CContentControlPr.prototype.GetPlaceholderText;
CContentControlPr.prototype['put_PlaceholderText']    = CContentControlPr.prototype.SetPlaceholderText;
CContentControlPr.prototype['get_FormPr']             = CContentControlPr.prototype.GetFormPr;
CContentControlPr.prototype['put_FormPr']             = CContentControlPr.prototype.SetFormPr;
CContentControlPr.prototype['get_PictureFormPr']      = CContentControlPr.prototype.GetPictureFormPr;
CContentControlPr.prototype['put_PictureFormPr']      = CContentControlPr.prototype.SetPictureFormPr;
CContentControlPr.prototype['get_ComplexFormPr']      = CContentControlPr.prototype.GetComplexFormPr;
CContentControlPr.prototype['put_ComplexFormPr']      = CContentControlPr.prototype.SetComplexFormPr;
CContentControlPr.prototype['get_NewKey']             = CContentControlPr.prototype.GetNewKey;
CContentControlPr.prototype['is_Signature']           = CContentControlPr.prototype.IsSignature;
