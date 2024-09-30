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
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Add_Kid]			= CChangesPDFFormAddKid;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Remove_Kid]		= CChangesPDFFormRemoveKid;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Form_Change_Display]	= CChangesPDFFormDisplay;

AscDFH.changesFactory[AscDFH.historyitem_Pdf_List_Form_Cur_Idxs]	= CChangesPDFListFormCurIdxs;

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
	oField.SetValue(Value);
	oField.Commit();
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
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFListFormCurIdxs(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFListFormCurIdxs.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFListFormCurIdxs.prototype.constructor = CChangesPDFListFormCurIdxs;
CChangesPDFListFormCurIdxs.prototype.Type = AscDFH.historyitem_Pdf_List_Form_Cur_Idxs;
CChangesPDFListFormCurIdxs.prototype.private_SetValue = function(Value)
{
	var oField = this.Class;
	oField.SetApiCurIdxs(Value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFPushbuttonImage(Class, Old, New, Color)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New, Color);
}
CChangesPDFPushbuttonImage.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFPushbuttonImage.prototype.constructor = CChangesPDFPushbuttonImage;
CChangesPDFPushbuttonImage.prototype.Type = AscDFH.historyitem_Pdf_Pushbutton_Image;
CChangesPDFPushbuttonImage.prototype.private_SetValue = function(Value)
{
	var oField = this.Class;
	oField.AddImage2(Value[0], Value[1]);
};
