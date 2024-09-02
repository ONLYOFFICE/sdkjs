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

AscDFH.changesFactory[AscDFH.historyitem_type_Pdf_Drawing_Page]	= CChangesPDFDrawingPage;
AscDFH.changesFactory[AscDFH.historyitem_type_Pdf_Drawing_Document]	= CChangesPdfDrawingObjectProperty;

let pdfDrawingsChangesMap = {};
window['AscDFH'].pdfDrawingsChangesMap = pdfDrawingsChangesMap;

function CChangesPdfDrawingObjectProperty(Class, Type, OldPr, NewPr) {
	this.Type = Type;
	var _OldPr = OldPr && OldPr.Get_Id ? OldPr.Get_Id() : undefined;
	var _NewPr = NewPr && NewPr.Get_Id ? NewPr.Get_Id() : undefined;
	AscDFH.CChangesBaseStringProperty.call(this, Class, _OldPr, _NewPr);
}
CChangesPdfDrawingObjectProperty.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPdfDrawingObjectProperty.prototype.constructor = CChangesPdfDrawingObjectProperty;

CChangesPdfDrawingObjectProperty.prototype.ReadFromBinary = function (reader) {
	reader.Seek2(reader.GetCurPos() - 4);
	this.Type = reader.GetLong();
	AscDFH.CChangesBaseStringProperty.prototype.ReadFromBinary.call(this, reader);
};
CChangesPdfDrawingObjectProperty.prototype.private_SetValue = function (Value) {
	var oObject = null;
	if (typeof Value === "string") {
		oObject = AscCommon.g_oTableId.Get_ById(Value);
		if (!oObject) {
			oObject = null;
		}
	}
	if (AscDFH.pdfDrawingsChangesMap[this.Type]) {
		AscDFH.pdfDrawingsChangesMap[this.Type](this.Class, oObject);
	}
};
CChangesPdfDrawingObjectProperty.prototype.Load = function(){
	this.Redo();
	this.RefreshRecalcData();
};

pdfDrawingsChangesMap[AscDFH.historyitem_type_Pdf_Drawing_Document] = function (oDrawing, value) {
	oDrawing.SetDocument(value);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFDrawingPage(Class, Old, New, Color)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New, Color);
}
CChangesPDFDrawingPage.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFDrawingPage.prototype.constructor = CChangesPDFDrawingPage;
CChangesPDFDrawingPage.prototype.Type = AscDFH.historyitem_type_Pdf_Drawing_Page;
CChangesPDFDrawingPage.prototype.private_SetValue = function(Value)
{
	let oDrawing = this.Class;
	oDrawing.SetPage(Value, true);
};
