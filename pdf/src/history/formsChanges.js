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

AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Value]			= CChangesPDFFormValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Parent_Value]		= CChangesPDFFormParentValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Format_Value]		= CChangesPDFFormFormatValue;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Changed]			= CChangesPDFFormChanged;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Add_Kid]			= CChangesPDFFormAddKid;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Remove_Kid]		= CChangesPDFFormRemoveKid;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Change_Display]	= CChangesPDFFormDisplay;

AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Cur_Idxs]		= CChangesPDFListFormCurIdxs;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Parent_Cur_Idxs]	= CChangesPDFListFormParentCurIdxs;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Top_Idx]			= CChangesPDFListTopIndex;

AscDFH.changesFactory[AscDFH.historyitem_Pdf_Pushbutton_Image]		= CChangesPDFPushbuttonImage;

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
function CChangesPDFFormAddKid(Class, Pos, Items)
{
	AscDFH.CChangesDrawingsContent.call(this, Class, this.Type, Pos, Items, true);
}
CChangesPDFFormAddKid.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesPDFFormAddKid.prototype.constructor = CChangesPDFFormAddKid;
CChangesPDFFormAddKid.prototype.Type = AscDFH.historyitem_Pdf_Form_Add_Kid;

CChangesPDFFormAddKid.prototype.Undo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oKid = this.Items[nIndex];

		oForm.RemoveKid(oKid);
		oKid.AddToRedraw();
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFFormAddKid.prototype.Redo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oKid = this.Items[nIndex];

		oForm.AddKid(oKid);
		oKid.AddToRedraw();
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesDrawingsContent}
 */
function CChangesPDFFormRemoveKid(Class, Pos, Items)
{
	AscDFH.CChangesDrawingsContent.call(this, Class, this.Type, Pos, Items, false);
}
CChangesPDFFormRemoveKid.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesPDFFormRemoveKid.prototype.constructor = CChangesPDFFormRemoveKid;
CChangesPDFFormRemoveKid.prototype.Type = AscDFH.historyitem_Pdf_Form_Remove_Kid;

CChangesPDFFormRemoveKid.prototype.Undo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oKid = this.Items[nIndex];

		oForm.AddKid(oKid);
		oKid.AddToRedraw();
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFFormRemoveKid.prototype.Redo = function()
{
	let oForm	= this.Class;
	let oDocument = Asc.editor.getPDFDoc();
	let oDrDoc	= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oKid = this.Items[nIndex];

		oForm.RemoveKid(oKid);
		oKid.AddToRedraw();
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
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
	oField.SetApiCurIdxs(Value);
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
	if (this.FromLoad && typeof Value === "string" && Value.length > 0) {
		let sImageId = AscCommon.getFullImageSrc2(Value);
		let _img = Asc.editor.ImageLoader.map_image_index[sImageId];
		if (_img && _img.Status === AscFonts.ImageLoadStatus.Complete) {
			oButtonField.AddImage2(Value, this.APType);
			return;
		}

		AscCommon.CollaborativeEditing.Add_NewImage(Value);
		AscCommon.CollaborativeEditing.m_aEndLoadCallbacks.push(oButtonField.AddImage2.bind(oButtonField, Value, this.APType));
	}
	else {
		oButtonField.AddImage2(Value, this.APType);
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
