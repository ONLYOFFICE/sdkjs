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


AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_AddItem]			= CChangesPDFDocumentAddItem;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RemoveItem]		= CChangesPDFDocumentRemoveItem;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_AddPage]			= CChangesPDFDocumentAddPage;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RemovePage]		= CChangesPDFDocumentRemovePage;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RotatePage]		= CChangesPDFDocumentRotatePage;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RecognizePage]	= CChangesPDFDocumentRecognizePage;

/**
 * @constructor
 * @extends {AscDFH.CChangesDrawingsContent}
 */
function CChangesPDFDocumentAddItem(Class, Pos, Items)
{
	AscDFH.CChangesDrawingsContent.call(this, Class, this.Type, Pos, Items, true);
}
CChangesPDFDocumentAddItem.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesPDFDocumentAddItem.prototype.constructor = CChangesPDFDocumentAddItem;
CChangesPDFDocumentAddItem.prototype.Type = AscDFH.historyitem_PDF_Document_AddItem;

CChangesPDFDocumentAddItem.prototype.Undo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	let oViewer		= Asc.editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oItem = this.Items[nIndex];

		if (oItem.IsAnnot()) {
			let nPage = oItem.GetPage();
			
			oItem.AddToRedraw();
			oDocument.annots.splice(oDocument.annots.indexOf(oItem), 1);

			oViewer.pagesInfo.pages[nPage].annots.splice(this.Pos, 1);
			if (oItem.IsComment())
				editor.sync_RemoveComment(oItem.GetId());
			
			oViewer.DrawingObjects.resetSelection();
			oItem.AddToRedraw();
		}
		else if (oItem.IsDrawing()) {
			let nPage = oItem.GetPage();
			
			oItem.AddToRedraw();
			oDocument.drawings.splice(oDocument.annots.indexOf(oItem), 1);
			oViewer.pagesInfo.pages[nPage].drawings.splice(this.Pos, 1);
			oViewer.DrawingObjects.resetSelection();
			oItem.AddToRedraw();
		}
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentAddItem.prototype.Redo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	let oViewer		= Asc.editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oItem = this.Items[nIndex];

		if (oItem.IsAnnot()) {
			let nPage = oItem.GetPage();
			
			oItem.AddToRedraw();
			oDocument.annots.push(oItem);
			oViewer.pagesInfo.pages[nPage].annots.splice(this.Pos, 0, oItem);
			if (oItem.IsComment())
				editor.sendEvent("asc_onAddComment", oItem.GetId(), oItem.GetAscCommentData());
			oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
			oViewer.DrawingObjects.resetSelection();
			oItem.AddToRedraw();
		}
		else if (oItem.IsDrawing()) {
			let nPage = oItem.GetPage();
			
			oItem.AddToRedraw();
			oDocument.drawings.push(oItem);
			oViewer.pagesInfo.pages[nPage].drawings.splice(this.Pos, 0, oItem);
			oViewer.DrawingObjects.resetSelection();
			oItem.AddToRedraw();
		}
	}

	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentAddItem.prototype.private_InsertInArrayLoad = function () {
	if (this.Items.length <= 0)
		return;

	this.Redo();
	return;
	let aChangedArray = this.private_GetChangedArray() || [];

	if (null !== aChangedArray) {
		let oContentChanges = this.private_GetContentChanges(), nPos;
		for (let i = 0; i < this.Items.length; ++i) {
			if (oContentChanges) {
				nPos = oContentChanges.Check(AscCommon.contentchanges_Add, this.Pos + i);
			}
			else {
				nPos = this.Pos + i;
			}

			let oElement = this.Items[i];

			nPos = Math.min(nPos, aChangedArray.length);
			aChangedArray.splice(nPos, 0, oElement);
		}
	}
};

/**
 * @constructor
 * @extends {AscDFH.CChangesDrawingsContent}
 */
function CChangesPDFDocumentRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesDrawingsContent.call(this, Class, this.Type, Pos, Items, false);
}
CChangesPDFDocumentRemoveItem.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesPDFDocumentRemoveItem.prototype.constructor = CChangesPDFDocumentRemoveItem;
CChangesPDFDocumentRemoveItem.prototype.Type = AscDFH.historyitem_PDF_Document_RemoveItem;

CChangesPDFDocumentRemoveItem.prototype.Undo = function()
{
	let oDocument = this.Class;
	let oViewer = editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oItem = this.Items[nIndex];

		if (oItem.IsForm()) {
			if (oItem.IsWidget()) {
				let nPage = oItem.GetPage();
				oItem.AddToRedraw();

				oDocument.widgets.push(oItem);
				oViewer.pagesInfo.pages[nPage].fields.splice(this.Pos, 0, oItem);
			}
			else {
				oDocument.widgetsParents.push(oItem);
			}
		}
		else if (oItem.IsAnnot()) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.push(oItem);
			oViewer.pagesInfo.pages[nPage].annots.splice(this.Pos, 0, oItem);
			if (oItem.GetReply(0) != null || oItem.GetType() != AscPDF.ANNOTATIONS_TYPES.FreeText && oItem.GetContents())
				editor.sendEvent("asc_onAddComment", oItem.GetId(), oItem.GetAscCommentData());

			oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
		}
		else if (oItem.IsDrawing()) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.drawings.push(oItem);
			oViewer.pagesInfo.pages[nPage].drawings.splice(this.Pos, 0, oItem);
		}
	}

	oDocument.mouseDownAnnot = null;
};
CChangesPDFDocumentRemoveItem.prototype.Redo = function()
{
	var oDocument = this.Class;
	let oViewer = editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let oItem = this.Items[nIndex];

		if (oItem.IsForm()) {
			if (oItem.IsWidget()) {
				oDocument.RemoveForm(oItem);
			}
			else {
				let nIdx = oDocument.widgetsParents.indexOf(oItem);
				if (nIdx != -1) {
					oDocument.widgetsParents.splice(nIdx, oItem);
				}
			}
		}
		else if (oItem.IsAnnot()) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.splice(oDocument.annots.indexOf(oItem), 1);
			oViewer.pagesInfo.pages[nPage].annots.splice(this.Pos, 1);
			if (oItem.GetReply(0) != null || oItem.GetType() != AscPDF.ANNOTATIONS_TYPES.FreeText && oItem.GetContents())
				editor.sync_RemoveComment(oItem.GetId());
		}
		else if (oItem.IsDrawing()) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.drawings.splice(oDocument.annots.indexOf(oItem), 1);
			oViewer.pagesInfo.pages[nPage].drawings.splice(this.Pos, 1);
		}
	}
	
	oDocument.mouseDownAnnot = null;
};
CChangesPDFDocumentRemoveItem.prototype.private_RemoveInArrayLoad = function()
{
	this.Redo();
	
	var aChangedArray = this.private_GetChangedArray();
	if (null !== aChangedArray) {
		var oContentChanges = this.private_GetContentChanges(), nPos;
		for (var i = 0; i < this.Items.length; ++i) {
			if (oContentChanges) {
				nPos = oContentChanges.Check(AscCommon.contentchanges_Remove, this.Pos + i);
			}
			else {
				nPos = this.Pos + i;
			}
			if (false === nPos) {
				continue;
			}
			aChangedArray.splice(nPos, 1);
		}
	}
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentAddPage(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesPDFDocumentAddPage.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentAddPage.prototype.constructor = CChangesPDFDocumentAddPage;
CChangesPDFDocumentAddPage.prototype.Type = AscDFH.historyitem_PDF_Document_AddPage;

CChangesPDFDocumentAddPage.prototype.Undo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		oDocument.RemovePage(nPos);
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentAddPage.prototype.Redo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		let oItem = this.Items[nIndex];
		oDocument.AddPage(nPos, oItem)
	}

	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentAddPage.prototype.private_WriteItem = function(Writer, oPage)
{
	Writer.WriteLong(oPage.Rotate);
	Writer.WriteLong(oPage.Dpi);
	Writer.WriteLong(oPage.W);
	Writer.WriteLong(oPage.H);
};
CChangesPDFDocumentAddPage.prototype.private_ReadItem = function(Reader)
{
	return {
		Rotate: Reader.GetLong(),
		Dpi: Reader.GetLong(),
		W: Reader.GetLong(),
		H: Reader.GetLong()
	};
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentRemovePage(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesPDFDocumentRemovePage.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentRemovePage.prototype.constructor = CChangesPDFDocumentRemovePage;
CChangesPDFDocumentRemovePage.prototype.Type = AscDFH.historyitem_PDF_Document_RemovePage;

CChangesPDFDocumentRemovePage.prototype.Undo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		let oItem = this.Items[nIndex];
		oDocument.AddPage(nPos, oItem);
	}
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentRemovePage.prototype.Redo = function()
{
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		oDocument.RemovePage(nPos)
	}

	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentRemovePage.prototype.private_WriteItem = function(Writer, oPage)
{
	Writer.WriteLong(oPage.Rotate);
	Writer.WriteLong(oPage.Dpi);
	Writer.WriteLong(oPage.W);
	Writer.WriteLong(oPage.H);
};
CChangesPDFDocumentRemovePage.prototype.private_ReadItem = function(Reader)
{
	return {
		Rotate: Reader.GetLong(),
		Dpi: Reader.GetLong(),
		W: Reader.GetLong(),
		H: Reader.GetLong()
	};
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFDocumentRotatePage(Class, nPage, Old, New)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	this.Page = nPage;
}
CChangesPDFDocumentRotatePage.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFDocumentRotatePage.prototype.constructor = CChangesPDFDocumentRotatePage;
CChangesPDFDocumentRotatePage.prototype.Type = AscDFH.historyitem_PDF_Document_RotatePage;
CChangesPDFDocumentRotatePage.prototype.private_SetValue = function(Value)
{
	let oDoc = this.Class;
	oDoc.SetPageRotate(this.Page, Value);
};
CChangesPDFDocumentRotatePage.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (undefined === this.Page)
		nFlags |= 1;

	if (undefined === this.New)
		nFlags |= 2;

	if (undefined === this.Old)
		nFlags |= 3;

	Writer.WriteLong(nFlags);

	if (undefined !== this.Page)
		Writer.WriteLong(this.Page);

	if (undefined !== this.New)
		Writer.WriteLong(this.New);

	if (undefined !== this.Old)
		Writer.WriteLong(this.Old);
};
CChangesPDFDocumentRotatePage.prototype.ReadFromBinary = function(Reader)
{
	// Long  : Flag
	// 1-bit : Подсвечивать ли данные изменения
	// 2-bit : IsUndefined New
	// 3-bit : IsUndefined Old
	// long : New
	// long : Old


	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Page = undefined;
	else
		this.Page = Reader.GetLong();

	if (nFlags & 2)
		this.New = undefined;
	else
		this.New = Reader.GetLong();

	if (nFlags & 3)
		this.Old = undefined;
	else
		this.Old = Reader.GetLong();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFDocumentRecognizePage(Class, nPage, Old, New)
{
	AscDFH.CChangesBaseBoolProperty.call(this, Class, Old, New);
	this.Page = nPage;
}
CChangesPDFDocumentRecognizePage.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentRecognizePage.prototype.constructor = CChangesPDFDocumentRecognizePage;
CChangesPDFDocumentRecognizePage.prototype.Type = AscDFH.historyitem_PDF_Document_RecognizePage;
CChangesPDFDocumentRecognizePage.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (undefined === this.Page)
		nFlags |= 1;
	if (undefined === this.New)
		nFlags |= 2;
	if (undefined === this.Old)
		nFlags |= 3;

	
	Writer.WriteLong(nFlags);
	if (undefined !== this.Page)
		Writer.WriteLong(this.Page);
};
CChangesPDFDocumentRecognizePage.prototype.ReadFromBinary = function(Reader)
{
	let nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Page = undefined;
	else
		this.Page = Reader.GetLong();

	if (nFlags & 2)
		this.New = false;
	else
		this.New = true;

	if (nFlags & 3)
		this.Old = false;
	else
		this.Old = true;
};
CChangesPDFDocumentRecognizePage.prototype.private_SetValue = function(bRecognize)
{
	let oDoc = this.Class;
	let oFile = oDoc.Viewer.file;
	let nPage = this.Page;

	oFile.pages[nPage].isConvertedToShapes = bRecognize;
	delete oDoc.Viewer.drawingPages[nPage].Image;

	oDoc.Viewer.paint(function() {
		oDoc.Viewer.thumbnails._repaintPage(nPage);
	});
};
