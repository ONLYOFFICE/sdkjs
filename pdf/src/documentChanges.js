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


AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_AddItem]		= CChangesPDFDocumentAddItem;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RemoveItem]	= CChangesPDFDocumentRemoveItem;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentAddItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesPDFDocumentAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentAddItem.prototype.constructor = CChangesPDFDocumentAddItem;
CChangesPDFDocumentAddItem.prototype.Type = AscDFH.historyitem_PDF_Document_AddItem;

CChangesPDFDocumentAddItem.prototype.Undo = function()
{
	var oDocument = this.Class;
	let oViewer = editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		let oItem = this.Items[nIndex];

		if (oItem instanceof AscPDF.CAnnotationBase) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.splice(nPos, 1);
			this.PosInPage = oViewer.pagesInfo.pages[nPage].annots.indexOf(oItem);
			oViewer.pagesInfo.pages[nPage].annots.splice(this.PosInPage, 1);
			if (oItem.IsComment())
				editor.sync_RemoveComment(oItem.GetId());
			
		}
	}

	oDocument.mouseDownAnnot = null;
};
CChangesPDFDocumentAddItem.prototype.Redo = function()
{
	var oDocument = this.Class;
	let oViewer = editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		let oItem = this.Items[nIndex];

		if (oItem instanceof AscPDF.CAnnotationBase) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.splice(nPos, 0, oItem);
			oViewer.pagesInfo.pages[nPage].annots.splice(this.PosInPage, 0, oItem);
			if (oItem.IsComment())
				editor.sendEvent("asc_onAddComment", oItem.GetId(), oItem.GetAscCommentData());

			oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
		}
	}
	oDocument.mouseDownAnnot = null;
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesPDFDocumentRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentRemoveItem.prototype.constructor = CChangesPDFDocumentRemoveItem;
CChangesPDFDocumentRemoveItem.prototype.Type = AscDFH.historyitem_PDF_Document_RemoveItem;

CChangesPDFDocumentRemoveItem.prototype.Undo = function()
{
	let oDocument = this.Class;
	let oViewer = editor.getDocumentRenderer();
	
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let nPos = this.Pos[0];
		let nPosInPage = this.Pos[1];

		let oItem = this.Items[nIndex];

		if (oItem instanceof AscPDF.CAnnotationBase) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.splice(nPos, 0, oItem);
			oViewer.pagesInfo.pages[nPage].annots.splice(nPosInPage, 0, oItem);
			if (oItem.GetReply(0) != null || oItem.GetType() != AscPDF.ANNOTATIONS_TYPES.FreeText && oItem.GetContents())
				editor.sendEvent("asc_onAddComment", oItem.GetId(), oItem.GetAscCommentData());

			oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
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
		let nPos		= this.Pos[0];
		let nPosInPage	= this.Pos[1];

		let oItem = this.Items[nIndex];

		if (oItem instanceof AscPDF.CAnnotationBase) {
			let nPage = oItem.GetPage();
			oItem.AddToRedraw();

			oDocument.annots.splice(nPos, 1);
			oViewer.pagesInfo.pages[nPage].annots.splice(nPosInPage, 1);
			if (oItem.GetReply(0) != null || oItem.GetType() != AscPDF.ANNOTATIONS_TYPES.FreeText && oItem.GetContents())
				editor.sync_RemoveComment(oItem.GetId());
		}
	}
	
	oDocument.mouseDownAnnot = null;
};
