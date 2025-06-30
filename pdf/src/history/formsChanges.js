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

// common
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Value]			= CChangesPDFFormValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Parent_Value]		= CChangesPDFFormParentValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Format_Value]		= CChangesPDFFormFormatValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Changed]			= CChangesPDFFormChanged;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Add_Kid]			= CChangesPDFFormKidsContent;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Change_Display]	= CChangesPDFFormDisplay;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Border_Color]		= CChangesPDFFormBorderColor;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_BG_Color]			= CChangesPDFFormBGrColor;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Border_Style]		= CChangesPDFFormBorderStyle;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Required]			= CChangesPDFFormRequired;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Text_Color]		= CChangesPDFFormTextColor;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Text_Font]		= CChangesPDFFormTextFont;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Text_Size]		= CChangesPDFFormTextSize;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Default_Value]	= CChangesPDFFormDefaultValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Rect]				= CChangesPDFFormRect;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Actions]			= CChangesPDFFormActions;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Partial_Name]		= CChangesPDFFormPartialName;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Meta]				= CChangesPDFFormMeta;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Read_Only]		= CChangesPDFFormReadOnly;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_No_Export]		= CChangesPDFFormNoExport;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Border_Width]		= CChangesPDFFormBorderWidth;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Locked]			= CChangesPDFFormLocked;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Rotate]			= CChangesPDFFormRotate;

// text
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_Multiline]			= CChangesPDFTextFormMultiline;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_Align]				= CChangesPDFTextFormAlign;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_Char_Limit]			= CChangesPDFTextCharLimit;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_Comb]				= CChangesPDFTextComb;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_DoNot_Scroll]		= CChangesPDFTextFormDoNotScroll;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_Password]			= CChangesPDFTextFormPassword;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_File_Select]			= CChangesPDFTextFormFileSelect;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Text_Form_DoNot_Spell_Check]	= CChangesPDFTextFormDoNotSpellCheck;

// combobox
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Combobox_Form_Editable]	= CChangesPDFComboboxFieldEditable;

// listbox
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Cur_Idxs]			= CChangesPDFListFormCurIdxs;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Parent_Cur_Idxs]		= CChangesPDFListFormParentCurIdxs;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Top_Idx]				= CChangesPDFListTopIndex;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Option]				= CChangesPDFListOption;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Content_Option]		= CChangesPDFListContentOption;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Commit_On_Sel_Change]= CChangesPDFListCommitOnSelChange;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Multiple_Selection]	= CChangesPDFListMultipleSelection;

// button
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Image]			= CChangesPDFPushbuttonImage;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Layout]			= CChangesPDFPushbuttonLayout;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Icon_Pos]		= CChangesPDFPushbuttonIconPos;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Highlight_Type]	= CChangesPDFPushbuttonHighlightType;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Scale_When_Type]= CChangesPDFPushbuttonScaleWhenType;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Scale_How_Type]	= CChangesPDFPushbuttonScaleHowType;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Fit_Bounds]		= CChangesPDFPushbuttonFitBounds;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Caption]		= CChangesPDFPushbuttonCaption;


// checbox/radio
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Checkbox_No_Toggle_To_Off]	= CChangesPDFCheckboxNoToggleToOff;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Checkbox_Style]			= CChangesPDFCheckboxStyle;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Checkbox_Export_Value]		= CChangesPDFCheckboxExpValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Checkbox_Options]			= CChangesPDFCheckOptions;

// radio
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Radiobutton_Is_Unison]	= CChangesPDFRadiobuttonIsUnison;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormValue.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormValue.prototype.constructor = CChangesPDFFormValue;
CChangesPDFFormValue.prototype.Type = AscDFH.historyitem_Pdf_Form_Value;
CChangesPDFFormValue.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.private_SetValue(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormFormatValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormFormatValue.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormFormatValue.prototype.constructor = CChangesPDFFormFormatValue;
CChangesPDFFormFormatValue.prototype.Type = AscDFH.historyitem_Pdf_Form_Format_Value;
CChangesPDFFormFormatValue.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetFormatValue(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormParentValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormParentValue.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormParentValue.prototype.constructor = CChangesPDFFormParentValue;
CChangesPDFFormParentValue.prototype.Type = AscDFH.historyitem_Pdf_Form_Parent_Value;
CChangesPDFFormParentValue.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetParentValue(Value);
};

CChangesPDFFormValue.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	var nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	if (Array.isArray(this.New) || Array.isArray(this.Old))
		nFlags |= 8;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		Writer.WriteString2(Array.isArray(this.New) ? JSON.stringify(this.New) : this.New);
	}
		
	if (undefined !== this.Old) {
		Writer.WriteString2(Array.isArray(this.Old) ? JSON.stringify(this.Old) : this.Old);
	}
};
CChangesPDFFormValue.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old


	var nFlags = Reader.GetLong();

	let isArrayValue = false;
	if (nFlags & 8)
		isArrayValue = true;

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.New = undefined;
	else
		this.New = isArrayValue ? JSON.parse(Reader.GetString2()) : Reader.GetString2();

	if (nFlags & 4)
		this.Old = undefined;
	else
		this.Old = isArrayValue ? JSON.parse(Reader.GetString2()) : Reader.GetString2();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFFormChanged(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormChanged.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFFormChanged.prototype.constructor = CChangesPDFFormChanged;
CChangesPDFFormChanged.prototype.Type = AscDFH.historyitem_Pdf_Form_Changed;
CChangesPDFFormChanged.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._wasChanged = Value;
	oForm.IsWidget() && oForm.SetDrawFromStream(!Value);
};


/**
 * @constructor
 * @extends {AscDFH.CChangesDrawingsContent}
 */
function CChangesPDFFormKidsContent(Class, Pos, Items, bAdd)
{
	AscDFH.CChangesDrawingsContent.call(this, Class, this.Type, Pos, Items, bAdd);
}
CChangesPDFFormKidsContent.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesPDFFormKidsContent.prototype.constructor = CChangesPDFFormKidsContent;
CChangesPDFFormKidsContent.prototype.Type = AscDFH.historyitem_Pdf_Form_Add_Kid;

CChangesPDFFormKidsContent.prototype.Undo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	if (this.IsAdd()) {
		for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
			let oKid = this.Items[nIndex];
	
			oForm.RemoveKid(oKid);
			oKid.AddToRedraw();
		}
	}
	else {
		for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
			let oKid = this.Items[nIndex];
	
			oForm.AddKid(oKid);
			oKid.AddToRedraw();
		}
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFFormKidsContent.prototype.Redo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	if (this.IsAdd()) {
		for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
			let oKid = this.Items[nIndex];
	
			oForm.AddKid(oKid);
			oKid.AddToRedraw();
		}
	}
	else {
		for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
			let oKid = this.Items[nIndex];
	
			oForm.RemoveKid(oKid);
			oKid.AddToRedraw();
		}
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFFormKidsContent.prototype.private_InsertInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

    let oParentField = this.Class;
	let aKids = oParentField.GetKids();

    let oDocument = this.Class.GetDocument();
    let oContentChanges = this.private_GetContentChanges();

    for (let i = 0; i < this.Items.length; ++i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

		aKids.splice(nPos, 0, oItem);
		oItem._parent = oParentField;

        oItem.AddToRedraw();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFFormKidsContent.prototype.private_RemoveInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

	let oParentField = this.Class;
	let aKids = oParentField.GetKids();

    let oDocument = this.Class.GetDocument();
    let oContentChanges = this.private_GetContentChanges();

    // Remove items in reverse order to maintain indices
    for (let i = this.Items.length - 1; i >= 0; --i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove,  true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

		aKids.splice(nPos, 1);
		oItem._parent = null;

		oItem.AddToRedraw();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFFormKidsContent.prototype.private_GetContentChanges = function() {
    return this.Class.kidsContentChanges;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFFormDisplay(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormDisplay.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFFormDisplay.prototype.constructor = CChangesPDFFormDisplay;
CChangesPDFFormDisplay.prototype.Type = AscDFH.historyitem_Pdf_Form_Change_Display;
CChangesPDFFormDisplay.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetDisplay(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFFormBorderColor(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormBorderColor.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFFormBorderColor.prototype.constructor = CChangesPDFFormBorderColor;
CChangesPDFFormBorderColor.prototype.Type = AscDFH.historyitem_Pdf_Form_Border_Color;
CChangesPDFFormBorderColor.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetBorderColor(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFFormBGrColor(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormBGrColor.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFFormBGrColor.prototype.constructor = CChangesPDFFormBGrColor;
CChangesPDFFormBGrColor.prototype.Type = AscDFH.historyitem_Pdf_Form_BG_Color;
CChangesPDFFormBGrColor.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetBackgroundColor(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFFormBorderStyle(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormBorderStyle.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFFormBorderStyle.prototype.constructor = CChangesPDFFormBorderStyle;
CChangesPDFFormBorderStyle.prototype.Type = AscDFH.historyitem_Pdf_Form_Border_Style;
CChangesPDFFormBorderStyle.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetBorderStyle(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFFormRequired(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormRequired.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFFormRequired.prototype.constructor = CChangesPDFFormRequired;
CChangesPDFFormRequired.prototype.Type = AscDFH.historyitem_Pdf_Form_Required;
CChangesPDFFormRequired.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._required = Value;
	oForm.AddToRedraw();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFFormTextColor(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormTextColor.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFFormTextColor.prototype.constructor = CChangesPDFFormTextColor;
CChangesPDFFormTextColor.prototype.Type = AscDFH.historyitem_Pdf_Form_Text_Color;
CChangesPDFFormTextColor.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetTextColor(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormTextFont(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormTextFont.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormTextFont.prototype.constructor = CChangesPDFFormTextFont;
CChangesPDFFormTextFont.prototype.Type = AscDFH.historyitem_Pdf_Form_Text_Font;
CChangesPDFFormTextFont.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetTextFontActual(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFFormTextSize(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormTextSize.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFFormTextSize.prototype.constructor = CChangesPDFFormTextSize;
CChangesPDFFormTextSize.prototype.Type = AscDFH.historyitem_Pdf_Form_Text_Size;
CChangesPDFFormTextSize.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetTextSize(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormDefaultValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormDefaultValue.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormDefaultValue.prototype.constructor = CChangesPDFFormDefaultValue;
CChangesPDFFormDefaultValue.prototype.Type = AscDFH.historyitem_Pdf_Form_Default_Value;
CChangesPDFFormDefaultValue.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetDefaultValue(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFFormRect(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormRect.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFFormRect.prototype.constructor = CChangesPDFFormRect;
CChangesPDFFormRect.prototype.Type = AscDFH.historyitem_Pdf_Form_Rect;
CChangesPDFFormRect.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetRect(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormActions(Class, oOldActionsInfo, oNewActionsInfo, nTriggerType, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, JSON.stringify(oOldActionsInfo), JSON.stringify(oNewActionsInfo), Color);
	this.TriggerType = nTriggerType;
}
CChangesPDFFormActions.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormActions.prototype.constructor = CChangesPDFFormActions;
CChangesPDFFormActions.prototype.Type = AscDFH.historyitem_Pdf_Form_Actions;
CChangesPDFFormActions.prototype.CreateReverseChange = function() {
	return new this.constructor(this.Class, this.New, this.Old, this.TriggerType, this.Color);
};
CChangesPDFFormActions.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetActions(this.TriggerType, JSON.parse(Value));
};

CChangesPDFFormActions.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.TriggerType)
		nFlags |= 2;

	if (undefined === this.New)
		nFlags |= 4;

	if (undefined === this.Old)
		nFlags |= 8;
	

	Writer.WriteLong(nFlags);

	if (undefined !== this.TriggerType)
		Writer.WriteLong(this.TriggerType);

	if (undefined !== this.New)
		Writer.WriteString2(this.New);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old);
};
CChangesPDFFormActions.prototype.ReadFromBinary = function(Reader)
{
	this.FromLoad = true;

	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// long : New
	// long : Old


	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.TriggerType = undefined;
	else
		this.TriggerType = Reader.GetLong();

	if (nFlags & 4)
		this.New = undefined;
	else
		this.New = Reader.GetString2();

	if (nFlags & 8)
		this.Old = undefined;
	else
		this.Old = Reader.GetString2();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFFormPartialName(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormPartialName.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFFormPartialName.prototype.constructor = CChangesPDFFormPartialName;
CChangesPDFFormPartialName.prototype.Type = AscDFH.historyitem_Pdf_Form_Partial_Name;
CChangesPDFFormPartialName.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField._partialName = Value;
	oField.SetNeedUpdateEditShape(true);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFFormMeta(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormMeta.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFFormMeta.prototype.constructor = CChangesPDFFormMeta;
CChangesPDFFormMeta.prototype.Type = AscDFH.historyitem_Pdf_Form_Meta;
CChangesPDFFormMeta.prototype.private_SetValue = function(Value)
{
	var oField = this.Class;
	oField.SetMeta(Value);
};

CChangesPDFFormMeta.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		Writer.WriteString2(JSON.stringify(this.New));
	}
		
	if (undefined !== this.Old) {
		Writer.WriteString2(JSON.stringify(this.Old));
	}
};
CChangesPDFFormMeta.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.New = undefined;
	else
		this.New = JSON.parse(Reader.GetString2());

	if (nFlags & 4)
		this.Old = undefined;
	else
		this.Old = JSON.parse(Reader.GetString2());
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFFormReadOnly(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormReadOnly.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFFormReadOnly.prototype.constructor = CChangesPDFFormReadOnly;
CChangesPDFFormReadOnly.prototype.Type = AscDFH.historyitem_Pdf_Form_Read_Only;
CChangesPDFFormReadOnly.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._readOnly = Value;
	oForm.AddToRedraw();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFFormNoExport(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormNoExport.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFFormNoExport.prototype.constructor = CChangesPDFFormNoExport;
CChangesPDFFormNoExport.prototype.Type = AscDFH.historyitem_Pdf_Form_No_Export;
CChangesPDFFormNoExport.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._noExport = Value;
	oForm.AddToRedraw();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFFormBorderWidth(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormBorderWidth.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFFormBorderWidth.prototype.constructor = CChangesPDFFormBorderWidth;
CChangesPDFFormBorderWidth.prototype.Type = AscDFH.historyitem_Pdf_Form_Border_Width;
CChangesPDFFormBorderWidth.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetBorderWidth(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFFormLocked(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormLocked.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFFormLocked.prototype.constructor = CChangesPDFFormLocked;
CChangesPDFFormLocked.prototype.Type = AscDFH.historyitem_Pdf_Form_Locked;
CChangesPDFFormLocked.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._locked = Value;
	oForm.AddToRedraw();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFFormRotate(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFFormRotate.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFFormRotate.prototype.constructor = CChangesPDFFormRotate;
CChangesPDFFormRotate.prototype.Type = AscDFH.historyitem_Pdf_Form_Rotate;
CChangesPDFFormRotate.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetRotate(Value);
};


//------------------------------------------------------------------------------------------------------------------
//
// Text Form
//
//------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextFormMultiline(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormMultiline.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextFormMultiline.prototype.constructor = CChangesPDFTextFormMultiline;
CChangesPDFTextFormMultiline.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_Multiline;
CChangesPDFTextFormMultiline.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetMultiline(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFTextFormAlign(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormAlign.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFTextFormAlign.prototype.constructor = CChangesPDFTextFormAlign;
CChangesPDFTextFormAlign.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_Align;
CChangesPDFTextFormAlign.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetAlign(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFTextCharLimit(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextCharLimit.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFTextCharLimit.prototype.constructor = CChangesPDFTextCharLimit;
CChangesPDFTextCharLimit.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_Char_Limit;
CChangesPDFTextCharLimit.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetCharLimit(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextComb(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextComb.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextComb.prototype.constructor = CChangesPDFTextComb;
CChangesPDFTextComb.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_Comb;
CChangesPDFTextComb.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetComb(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextFormDoNotScroll(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormDoNotScroll.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextFormDoNotScroll.prototype.constructor = CChangesPDFTextFormDoNotScroll;
CChangesPDFTextFormDoNotScroll.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_DoNot_Scroll;
CChangesPDFTextFormDoNotScroll.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetDoNotScroll(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextFormPassword(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormPassword.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextFormPassword.prototype.constructor = CChangesPDFTextFormPassword;
CChangesPDFTextFormPassword.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_Password;
CChangesPDFTextFormPassword.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetPassword(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextFormFileSelect(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormFileSelect.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextFormFileSelect.prototype.constructor = CChangesPDFTextFormFileSelect;
CChangesPDFTextFormFileSelect.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_File_Select;
CChangesPDFTextFormFileSelect.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetFileSelect(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFTextFormDoNotSpellCheck(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFTextFormDoNotSpellCheck.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFTextFormDoNotSpellCheck.prototype.constructor = CChangesPDFTextFormDoNotSpellCheck;
CChangesPDFTextFormDoNotSpellCheck.prototype.Type = AscDFH.historyitem_Pdf_Text_Form_DoNot_Spell_Check;
CChangesPDFTextFormDoNotSpellCheck.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetDoNotSpellCheck(Value);
};

//------------------------------------------------------------------------------------------------------------------
//
// List Form
//
//------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFComboboxFieldEditable(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFComboboxFieldEditable.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFComboboxFieldEditable.prototype.constructor = CChangesPDFComboboxFieldEditable;
CChangesPDFComboboxFieldEditable.prototype.Type = AscDFH.historyitem_Pdf_Combobox_Form_Editable;
CChangesPDFComboboxFieldEditable.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._editable = Value;
};

//------------------------------------------------------------------------------------------------------------------
//
// List Form
//
//------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFListFormCurIdxs(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListFormCurIdxs.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFListFormCurIdxs.prototype.constructor = CChangesPDFListFormCurIdxs;
CChangesPDFListFormCurIdxs.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Cur_Idxs;
CChangesPDFListFormCurIdxs.prototype.private_SetValue = function(Value)
{
	var oField = this.Class;
	oField.SetCurIdxs(Value);
};

CChangesPDFListFormCurIdxs.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		Writer.WriteString2(JSON.stringify(this.New));
	}
		
	if (undefined !== this.Old) {
		Writer.WriteString2(JSON.stringify(this.Old));
	}
};
CChangesPDFListFormCurIdxs.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.New = undefined;
	else
		this.New = JSON.parse(Reader.GetString2());

	if (nFlags & 4)
		this.Old = undefined;
	else
		this.Old = JSON.parse(Reader.GetString2());
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFListFormParentCurIdxs(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListFormParentCurIdxs.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFListFormParentCurIdxs.prototype.constructor = CChangesPDFListFormParentCurIdxs;
CChangesPDFListFormParentCurIdxs.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Parent_Cur_Idxs;
CChangesPDFListFormParentCurIdxs.prototype.private_SetValue = function(Value)
{
	var oField = this.Class;
	oField.SetParentCurIdxs(Value);
};

CChangesPDFListFormParentCurIdxs.prototype.WriteToBinary = function(Writer)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		Writer.WriteString2(JSON.stringify(this.New));
	}
		
	if (undefined !== this.Old) {
		Writer.WriteString2(JSON.stringify(this.Old));
	}
};
CChangesPDFListFormParentCurIdxs.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// string : New
	// string : Old

	let nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.New = undefined;
	else
		this.New = JSON.parse(Reader.GetString2());

	if (nFlags & 4)
		this.Old = undefined;
	else
		this.Old = JSON.parse(Reader.GetString2());
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFListTopIndex(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListTopIndex.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFListTopIndex.prototype.constructor = CChangesPDFListTopIndex;
CChangesPDFListTopIndex.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Top_Idx;
CChangesPDFListTopIndex.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField._topIdx = Value;
	if (Value != undefined) {
		oField._bAutoShiftContentView = false;
		oField._bShiftByTopIndex = true;
	}
	else {
		oField._bAutoShiftContentView = true;
		oField._bShiftByTopIndex = false;
	}

	oField.AddToRedraw();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFListOption(Class, Pos, Items, isAdd) {
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, isAdd);
}

CChangesPDFListOption.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFListOption.prototype.constructor = CChangesPDFListOption;
CChangesPDFListOption.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Option;

CChangesPDFListOption.prototype.WriteToBinary = function (writer) {
	writer.WriteBool(this.IsAdd());
	writer.WriteLong(this.Pos);
	writer.WriteString2(JSON.stringify(this.Items));
};
CChangesPDFListOption.prototype.ReadFromBinary = function (reader) {
	reader.Seek2(reader.GetCurPos() - 4);
	this.Type = reader.GetLong();
	this.Add = reader.GetBool();
	this.Pos = reader.GetLong();
	
	this.Items = JSON.parse(reader.GetString2());
};
CChangesPDFListOption.prototype.private_GetChangedArray = function () {
	return this.Class._options;
};
CChangesPDFListOption.prototype.private_InsertInArrayLoad = function () {
	if (this.Items.length <= 0)
		return;

	let aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		for (let i = this.Items.length - 1; i >= 0; i--) {
			this.Class.AddOption(this.Items[i], this.Pos);
		}
	}
};
CChangesPDFListOption.prototype.private_RemoveInArrayLoad = function () {

	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		this.Class.RemoveOption(this.Pos);
	}
};
CChangesPDFListOption.prototype.private_InsertInArrayUndoRedo = function () {
	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		for (let i = this.Items.length - 1; i >= 0; i--) {
			this.Class.AddOption(this.Items[i], this.Pos);
		}
	}
};
CChangesPDFListOption.prototype.private_RemoveInArrayUndoRedo = function () {

	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		this.Class.RemoveOption(this.Pos);
	}
};
CChangesPDFListOption.prototype.Load = function () {
	if (this.IsAdd()) {
		this.private_InsertInArrayLoad();
	}
	else {
		this.private_RemoveInArrayLoad();
	}
	this.RefreshRecalcData();
};
CChangesPDFListOption.prototype.Undo = function () {
	if (this.IsAdd()) {
		this.private_RemoveInArrayUndoRedo();
	}
	else {
		this.private_InsertInArrayUndoRedo();
	}
};
CChangesPDFListOption.prototype.Redo = function () {
	if (this.IsAdd()) {
		this.private_InsertInArrayUndoRedo();
	}
	else {
		this.private_RemoveInArrayUndoRedo();
	}
};
CChangesPDFListOption.prototype.IsContentChange = function () {
	return false;
};
CChangesPDFListOption.prototype.Copy = function()
{
	var oChanges = new this.constructor(this.Class, this.Type, this.Pos, this.Items, this.Add);

	oChanges.UseArray = this.UseArray;

	for (var nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
		oChanges.PosArray[nIndex] = this.PosArray[nIndex];

	return oChanges;
};
CChangesPDFListOption.prototype.ConvertToSimpleChanges = function()
{
	let arrSimpleActions = this.ConvertToSimpleActions();
	let arrChanges       = [];
	for (let nIndex = 0, nCount = arrSimpleActions.length; nIndex < nCount; ++nIndex)
	{
		let oAction = arrSimpleActions[nIndex];
		let oChange = new this.constructor(this.Class, this.Type, oAction.Pos, [oAction.Item], oAction.Add);
		arrChanges.push(oChange);
	}
	return arrChanges;
};
CChangesPDFListOption.prototype.CreateReverseChange = function(){
	var oRet = this.private_CreateReverseChange(this.constructor);
	oRet.Type = this.Type;
	oRet.Pos = this.Pos;
	return oRet;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFListContentOption(Class, Pos, Items, isAdd) {
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, isAdd);
}

CChangesPDFListContentOption.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFListContentOption.prototype.constructor = CChangesPDFListContentOption;
CChangesPDFListContentOption.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Content_Option;

CChangesPDFListContentOption.prototype.WriteToBinary = function (writer) {
	writer.WriteBool(this.IsAdd());
	writer.WriteLong(this.Pos);
	writer.WriteString2(JSON.stringify(this.Items));
};
CChangesPDFListContentOption.prototype.ReadFromBinary = function (reader) {
	reader.Seek2(reader.GetCurPos() - 4);
	this.Type = reader.GetLong();
	this.Add = reader.GetBool();
	this.Pos = reader.GetLong();
	
	this.Items = JSON.parse(reader.GetString2());
};
CChangesPDFListContentOption.prototype.private_GetChangedArray = function () {
	return this.Class._options;
};
CChangesPDFListContentOption.prototype.private_InsertInArrayLoad = function () {
	if (this.Items.length <= 0)
		return;

	let aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		for (let i = this.Items.length - 1; i >= 0; i--) {
			this.Class.private_AddOptionToContent(this.Items[i], this.Pos);
		}
	}
};
CChangesPDFListContentOption.prototype.private_RemoveInArrayLoad = function () {

	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		this.Class.private_RemoveOptionFromContent(this.Pos);
	}
};
CChangesPDFListContentOption.prototype.private_InsertInArrayUndoRedo = function () {
	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		for (let i = this.Items.length - 1; i >= 0; i--) {
			this.Class.private_AddOptionToContent(this.Items[i], this.Pos);
		}
	}
};
CChangesPDFListContentOption.prototype.private_RemoveInArrayUndoRedo = function () {

	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		this.Class.private_RemoveOptionFromContent(this.Pos);
	}
};
CChangesPDFListContentOption.prototype.Load = function () {
	if (this.IsAdd()) {
		this.private_InsertInArrayLoad();
	}
	else {
		this.private_RemoveInArrayLoad();
	}
	this.RefreshRecalcData();
};
CChangesPDFListContentOption.prototype.Undo = function () {
	if (this.IsAdd()) {
		this.private_RemoveInArrayUndoRedo();
	}
	else {
		this.private_InsertInArrayUndoRedo();
	}
};
CChangesPDFListContentOption.prototype.Redo = function () {
	if (this.IsAdd()) {
		this.private_InsertInArrayUndoRedo();
	}
	else {
		this.private_RemoveInArrayUndoRedo();
	}
};
CChangesPDFListContentOption.prototype.IsContentChange = function () {
	return false;
};
CChangesPDFListContentOption.prototype.Copy = function()
{
	var oChanges = new this.constructor(this.Class, this.Type, this.Pos, this.Items, this.Add);

	oChanges.UseArray = this.UseArray;

	for (var nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
		oChanges.PosArray[nIndex] = this.PosArray[nIndex];

	return oChanges;
};
CChangesPDFListContentOption.prototype.ConvertToSimpleChanges = function()
{
	let arrSimpleActions = this.ConvertToSimpleActions();
	let arrChanges       = [];
	for (let nIndex = 0, nCount = arrSimpleActions.length; nIndex < nCount; ++nIndex)
	{
		let oAction = arrSimpleActions[nIndex];
		let oChange = new this.constructor(this.Class, this.Type, oAction.Pos, [oAction.Item], oAction.Add);
		arrChanges.push(oChange);
	}
	return arrChanges;
};
CChangesPDFListContentOption.prototype.CreateReverseChange = function(){
	var oRet = this.private_CreateReverseChange(this.constructor);
	oRet.Type = this.Type;
	oRet.Pos = this.Pos;
	return oRet;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFListCommitOnSelChange(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListCommitOnSelChange.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFListCommitOnSelChange.prototype.constructor = CChangesPDFListCommitOnSelChange;
CChangesPDFListCommitOnSelChange.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Commit_On_Sel_Change;
CChangesPDFListCommitOnSelChange.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._commitOnSelChange = Value;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFListMultipleSelection(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListMultipleSelection.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFListMultipleSelection.prototype.constructor = CChangesPDFListMultipleSelection;
CChangesPDFListMultipleSelection.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Multiple_Selection;
CChangesPDFListMultipleSelection.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._multipleSelection = Value;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFPushbuttonImage(Class, sOldRasterId, sNewRasterId, nAPType, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, sOldRasterId, sNewRasterId, Color);
	this.APType = nAPType;
}
CChangesPDFPushbuttonImage.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFPushbuttonImage.prototype.constructor = CChangesPDFPushbuttonImage;
CChangesPDFPushbuttonImage.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Image;
CChangesPDFPushbuttonImage.prototype.CreateReverseChange = function() {
	return new this.constructor(this.Class, this.New, this.Old, this.APType, this.Color);
};
CChangesPDFPushbuttonImage.prototype.private_SetValue = function(Value)
{
	let oButtonField = this.Class;
	oButtonField.lastValue = Value;

	oButtonField.SetImageRasterId(Value, this.APType);

	if (this.FromLoad && typeof Value === "string" && Value.length > 0) {
		let sImageId = AscCommon.getFullImageSrc2(Value);
		let _img = Asc.editor.ImageLoader.map_image_index[sImageId];
		if (_img && _img.Status === AscFonts.ImageLoadStatus.Complete) {
			oButtonField.SetImage(Value);
			return;
		}

		AscCommon.CollaborativeEditing.Add_NewImage(Value);
		AscCommon.CollaborativeEditing.m_aEndLoadCallbacks.push(function() {
			if (oButtonField.lastValue === Value) {
				oButtonField.SetImage(Value);
			}
		}.bind(oButtonField));
	}
	else {
		oButtonField.SetImage(Value);
	}
};

CChangesPDFPushbuttonImage.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.APType)
		nFlags |= 2;

	if (undefined === this.New)
		nFlags |= 4;

	if (undefined === this.Old)
		nFlags |= 8;
	

	Writer.WriteLong(nFlags);

	if (undefined !== this.APType)
		Writer.WriteLong(this.APType);

	if (undefined !== this.New)
		Writer.WriteString2(this.New);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old);
};
CChangesPDFPushbuttonImage.prototype.ReadFromBinary = function(Reader)
{
	this.FromLoad = true;

	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// long : New
	// long : Old


	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.APType = undefined;
	else
		this.APType = Reader.GetLong();

	if (nFlags & 4)
		this.New = undefined;
	else
		this.New = Reader.GetString2();

	if (nFlags & 8)
		this.Old = undefined;
	else
		this.Old = Reader.GetString2();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFPushbuttonLayout(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonLayout.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFPushbuttonLayout.prototype.constructor = CChangesPDFPushbuttonLayout;
CChangesPDFPushbuttonLayout.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Layout;
CChangesPDFPushbuttonLayout.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetLayout(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFPushbuttonIconPos(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonIconPos.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFPushbuttonIconPos.prototype.constructor = CChangesPDFPushbuttonIconPos;
CChangesPDFPushbuttonIconPos.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Icon_Pos;
CChangesPDFPushbuttonIconPos.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetIconPosition(Value[0], Value[1]);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFPushbuttonHighlightType(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonHighlightType.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFPushbuttonHighlightType.prototype.constructor = CChangesPDFPushbuttonHighlightType;
CChangesPDFPushbuttonHighlightType.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Highlight_Type;
CChangesPDFPushbuttonHighlightType.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetHighlight(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFPushbuttonScaleWhenType(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonScaleWhenType.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFPushbuttonScaleWhenType.prototype.constructor = CChangesPDFPushbuttonScaleWhenType;
CChangesPDFPushbuttonScaleWhenType.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Scale_When_Type;
CChangesPDFPushbuttonScaleWhenType.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetScaleWhen(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFPushbuttonScaleHowType(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonScaleHowType.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFPushbuttonScaleHowType.prototype.constructor = CChangesPDFPushbuttonScaleHowType;
CChangesPDFPushbuttonScaleHowType.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Scale_How_Type;
CChangesPDFPushbuttonScaleHowType.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetScaleHow(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFPushbuttonFitBounds(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonFitBounds.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFPushbuttonFitBounds.prototype.constructor = CChangesPDFPushbuttonFitBounds;
CChangesPDFPushbuttonFitBounds.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Fit_Bounds;
CChangesPDFPushbuttonFitBounds.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetButtonFitBounds(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFPushbuttonCaption(Class, sOldRasterId, sNewRasterId, nAPType, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, sOldRasterId, sNewRasterId, Color);
	this.APType = nAPType;
}
CChangesPDFPushbuttonCaption.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFPushbuttonCaption.prototype.constructor = CChangesPDFPushbuttonCaption;
CChangesPDFPushbuttonCaption.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Caption;
CChangesPDFPushbuttonCaption.prototype.CreateReverseChange = function() {
	return new this.constructor(this.Class, this.New, this.Old, this.APType, this.Color);
};
CChangesPDFPushbuttonCaption.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetCaption(Value, this.APType);
};

CChangesPDFPushbuttonCaption.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.APType)
		nFlags |= 2;

	if (undefined === this.New)
		nFlags |= 4;

	if (undefined === this.Old)
		nFlags |= 8;
	

	Writer.WriteLong(nFlags);

	if (undefined !== this.APType)
		Writer.WriteLong(this.APType);

	if (undefined !== this.New)
		Writer.WriteString2(this.New);

	if (undefined !== this.Old)
		Writer.WriteString2(this.Old);
};
CChangesPDFPushbuttonCaption.prototype.ReadFromBinary = function(Reader)
{
	this.FromLoad = true;

	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// long : New
	// long : Old


	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.APType = undefined;
	else
		this.APType = Reader.GetLong();

	if (nFlags & 4)
		this.New = undefined;
	else
		this.New = Reader.GetString2();

	if (nFlags & 8)
		this.Old = undefined;
	else
		this.Old = Reader.GetString2();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFCheckboxNoToggleToOff(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCheckboxNoToggleToOff.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFCheckboxNoToggleToOff.prototype.constructor = CChangesPDFCheckboxNoToggleToOff;
CChangesPDFCheckboxNoToggleToOff.prototype.Type = AscDFH.historyitem_Pdf_Checkbox_No_Toggle_To_Off;
CChangesPDFCheckboxNoToggleToOff.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._noToggleToOff = Value;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFCheckboxStyle(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCheckboxStyle.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFCheckboxStyle.prototype.constructor = CChangesPDFCheckboxStyle;
CChangesPDFCheckboxStyle.prototype.Type = AscDFH.historyitem_Pdf_Checkbox_Style;
CChangesPDFCheckboxStyle.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField.SetStyle(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFCheckboxExpValue(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCheckboxExpValue.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFCheckboxExpValue.prototype.constructor = CChangesPDFCheckboxExpValue;
CChangesPDFCheckboxExpValue.prototype.Type = AscDFH.historyitem_Pdf_Checkbox_Export_Value;
CChangesPDFCheckboxExpValue.prototype.private_SetValue = function(Value)
{
	let oField = this.Class;
	oField._exportValue = Value;
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseBoolProperty}
 */
function CChangesPDFRadiobuttonIsUnison(Class, Old, New, Color)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New, Color);
}
CChangesPDFRadiobuttonIsUnison.prototype = Object.create(AscDFH.CChangesBaseBoolProperty.prototype);
CChangesPDFRadiobuttonIsUnison.prototype.constructor = CChangesPDFRadiobuttonIsUnison;
CChangesPDFRadiobuttonIsUnison.prototype.Type = AscDFH.historyitem_Pdf_Radiobutton_Is_Unison;
CChangesPDFRadiobuttonIsUnison.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm._radiosInUnison = Value;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfStringProperty}
 */
function CChangesPDFCheckOptions(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCheckOptions.prototype = Object.create(AscDFH.CChangesPDFArrayOfStringProperty.prototype);
CChangesPDFCheckOptions.prototype.constructor = CChangesPDFCheckOptions;
CChangesPDFCheckOptions.prototype.Type = AscDFH.historyitem_Pdf_Checkbox_Options;
CChangesPDFCheckOptions.prototype.private_SetValue = function(Value)
{
	let oForm = this.Class;
	oForm.SetOptions(Value);
};
