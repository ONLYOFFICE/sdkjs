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


// Update the changes factory mappings
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_AnnotsContent]	= CChangesPDFDocumentAnnotsContent;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_FieldsContent]	= CChangesPDFDocumentFieldsContent;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_DrawingsContent]	= CChangesPDFDocumentDrawingsContent;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_PagesContent]		= CChangesPDFDocumentPagesContent;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RotatePage]		= CChangesPDFDocumentRotatePage;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_RecognizePage]	= CChangesPDFDocumentRecognizePage;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_SetDocument]      = CChangesPDFObjectSetDocument;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_PageLocks]        = CChangesPDFDocumentPageLocks;
AscDFH.changesFactory[AscDFH.historyitem_PDF_PropLocker_ObjectId]	    = CChangesPDFPropLockerObjectId;
AscDFH.changesFactory[AscDFH.historyitem_PDF_Document_MovePage]         = CChangesPDFDocumentMovePage;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Document_Start_Merge_Pages]= CChangesPDFDocumentStartMergePages;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Document_Part_Merge_Pages] = CChangesPDFDocumentPartMergePages;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Document_End_Merge_Pages]  = CChangesPDFDocumentEndMergePages;
AscDFH.changesFactory[AscDFH.historyitem_Pdf_Document_Calc_Order]       = CChangesPDFCalcOrder;

function CChangesPDFArrayOfDoubleProperty(Class, Old, New) {
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesPDFArrayOfDoubleProperty.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFArrayOfDoubleProperty.prototype.constructor = CChangesPDFArrayOfDoubleProperty;

CChangesPDFArrayOfDoubleProperty.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		var nNewCount = this.New.length;
		Writer.WriteLong(nNewCount);
		for (var nIndex = 0; nIndex < nNewCount; ++nIndex)
			Writer.WriteDouble(this.New[nIndex]);
	}
	
	if (undefined !== this.Old) {
		var nOldCount = this.Old.length;
		Writer.WriteLong(nOldCount);
		for (var nIndex = 0; nIndex < nOldCount; ++nIndex)
			Writer.WriteDouble(this.Old[nIndex]);
	}
};
CChangesPDFArrayOfDoubleProperty.prototype.ReadFromBinary = function(Reader)
{
	// Long : Count of the columns in the new grid
	// Array of double : widths of columns in the new grid
	// Long : Count of the columns in the old grid
	// Array of double : widths of columns in the old grid

	let nFlags = Reader.GetLong();
	
	if (!(nFlags & 1)) {
		let nCount = Reader.GetLong();
		this.New = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
			this.New[nIndex] = Reader.GetDouble();
	}

	if (!(nFlags & 2)) {
		let nCount = Reader.GetLong();
		this.Old = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
			this.Old[nIndex] = Reader.GetDouble();
	} 
};

CChangesPDFArrayOfDoubleProperty.prototype.Load = function(){
	this.Redo();
	this.RefreshRecalcData();
};

window['AscDFH'].CChangesPDFArrayOfDoubleProperty = CChangesPDFArrayOfDoubleProperty;

function CChangesPDFArrayOfStringProperty(Class, Old, New) {
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesPDFArrayOfStringProperty.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFArrayOfStringProperty.prototype.constructor = CChangesPDFArrayOfStringProperty;

CChangesPDFArrayOfStringProperty.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New) {
		var nNewCount = this.New.length;
		Writer.WriteLong(nNewCount);
		for (var nIndex = 0; nIndex < nNewCount; ++nIndex)
			Writer.WriteString2(this.New[nIndex]);
	}
	
	if (undefined !== this.Old) {
		var nOldCount = this.Old.length;
		Writer.WriteLong(nOldCount);
		for (var nIndex = 0; nIndex < nOldCount; ++nIndex)
			Writer.WriteString2(this.Old[nIndex]);
	}
};
CChangesPDFArrayOfStringProperty.prototype.ReadFromBinary = function(Reader)
{
	// Long : Count of the columns in the new grid
	// Array of double : widths of columns in the new grid
	// Long : Count of the columns in the old grid
	// Array of double : widths of columns in the old grid

	let nFlags = Reader.GetLong();
	
	if (!(nFlags & 1)) {
		let nCount = Reader.GetLong();
		this.New = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
			this.New[nIndex] = Reader.GetString2();
	}

	if (!(nFlags & 2)) {
		let nCount = Reader.GetLong();
		this.Old = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
			this.Old[nIndex] = Reader.GetString2();
	} 
};

CChangesPDFArrayOfStringProperty.prototype.Load = function(){
	this.Redo();
	this.RefreshRecalcData();
};

window['AscDFH'].CChangesPDFArrayOfStringProperty = CChangesPDFArrayOfStringProperty;

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentAnnotsContent(Class, Pos, Items, bAdd)
{
    AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, bAdd);
}
CChangesPDFDocumentAnnotsContent.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentAnnotsContent.prototype.constructor = CChangesPDFDocumentAnnotsContent;
CChangesPDFDocumentAnnotsContent.prototype.Type = AscDFH.historyitem_PDF_Document_AnnotsContent;
CChangesPDFDocumentAnnotsContent.prototype.IsContentChange = function(){
    return true;
};

CChangesPDFDocumentAnnotsContent.prototype.ReadFromBinary = function (reader) {
    this.Type = reader.GetLong();
    this.Add = reader.GetBool();
    this.Pos = reader.GetLong();
    AscDFH.CChangesBaseContentChange.prototype.ReadFromBinary.call(this, reader);
};
CChangesPDFDocumentAnnotsContent.prototype.WriteToBinary = function (writer) {
    writer.WriteLong(this.Type);
    writer.WriteBool(this.IsAdd());
    writer.WriteLong(this.Pos);
    AscDFH.CChangesBaseContentChange.prototype.WriteToBinary.call(this, writer);
};

CChangesPDFDocumentAnnotsContent.prototype.Undo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();

    if (this.IsAdd()) {
        // Undo addition by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.annots.splice(oDocument.annots.indexOf(oItem), 1);
            oPage.annots.splice(this.Pos, 1);

            oItem.parentPage = null;
            oItem._page = -1;
            oItem.selectStartPage = -1;

            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
            Asc.editor.sync_RemoveComment(oItem.GetId());
        }
    } else {
        // Undo removal by adding items back
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.annots.push(oItem);
            oPage.annots.splice(this.Pos, 0, oItem);

            oItem.parentPage = oPage;
            oItem._page = oPage.GetIndex();
            oItem.selectStartPage = oItem._page;

            oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Types.display["hidden"] : window["AscPDF"].Api.Types.display["visible"]);
            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
            oDocument.CheckComment(oItem);
        }
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentAnnotsContent.prototype.Redo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();

    if (this.IsAdd()) {
        // Redo addition by adding items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.annots.push(oItem);
            oPage.annots.splice(this.Pos, 0, oItem);

            oItem.parentPage = oPage;
            oItem._page = oPage.GetIndex();
            oItem.selectStartPage = oItem._page;

            oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Types.display["hidden"] : window["AscPDF"].Api.Types.display["visible"]);
            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
            oDocument.CheckComment(oItem);
        }
    } else {
        // Redo removal by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.annots.splice(oDocument.annots.indexOf(oItem), 1);
            oPage.annots.splice(this.Pos, 1);

            oItem.parentPage = null;
            oItem._page = -1;
            oItem.selectStartPage = -1;

            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
            Asc.editor.sync_RemoveComment(oItem.GetId());
        }
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentAnnotsContent.prototype.Load = function() {
    if (this.IsAdd()) {
        this.private_InsertInArrayLoad();
    }
    else {
        this.private_RemoveInArrayLoad();
    }
};
CChangesPDFDocumentAnnotsContent.prototype.private_InsertInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();
    let oContentChanges = this.private_GetContentChanges();

    for (let i = 0; i < this.Items.length; ++i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        // Ensure position is within bounds
        nPos = Math.min(nPos, oDocument.annots.length);

        // Insert into document annots array
        oDocument.annots.splice(nPos, 0, oItem);

        // Insert into viewer annots array
        let annotsArray = oPage.annots;
        nPos = Math.min(nPos, annotsArray.length);
        annotsArray.splice(nPos, 0, oItem);

        oItem.parentPage = oPage;
        oItem._page = oPage.GetIndex();
        oItem.selectStartPage = oItem._page;

        oDocument.CheckComment(oItem);

        oItem.SetDisplay(oDocument.IsAnnotsHidden() ? window["AscPDF"].Api.Types.display["hidden"] : window["AscPDF"].Api.Types.display["visible"]);
        oViewer.DrawingObjects.resetSelection();
        oItem.AddToRedraw();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentAnnotsContent.prototype.private_RemoveInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();
    let oContentChanges = this.private_GetContentChanges();

    // Remove items in reverse order to maintain indices
    for (let i = this.Items.length - 1; i >= 0; --i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove,  true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        oItem.AddToRedraw();
        // Remove from document annots array
        let indexInAnnots = oDocument.annots.indexOf(oItem);
        if (indexInAnnots !== -1)
            oDocument.annots.splice(indexInAnnots, 1);

        // Remove from viewer annots array
        let annotsArray = oPage.annots;
        let indexInPageAnnots = annotsArray.indexOf(oItem);
        if (indexInPageAnnots !== -1)
            annotsArray.splice(indexInPageAnnots, 1);

        oItem.parentPage = null;
        oItem._page = -1;
        oItem.selectStartPage = -1;

        Asc.editor.sync_RemoveComment(oItem.GetId());

        oViewer.DrawingObjects.resetSelection();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentAnnotsContent.prototype.private_GetContentChanges = function() {
    return this.Class.annotsContentChanges;
};
CChangesPDFDocumentAnnotsContent.prototype.GetContentChangesClass = function() {
	return this.private_GetContentChanges();
};
CChangesPDFDocumentAnnotsContent.prototype.private_WriteItem = function (Writer, Item) {
    Writer.WriteString2(Item.Get_Id());
};
CChangesPDFDocumentAnnotsContent.prototype.private_ReadItem = function (Reader) {
    var Id = Reader.GetString2();
    return AscCommon.g_oTableId.Get_ById(Id);
};
CChangesPDFDocumentAnnotsContent.prototype.Copy = function() {
    let oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

    oChanges.UseArray = this.UseArray;

    for (let nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
        oChanges.PosArray[nIndex] = this.PosArray[nIndex];

    return oChanges;
};
CChangesPDFDocumentAnnotsContent.prototype.CreateReverseChange = function(){
    return this.private_CreateReverseChange(this.constructor);
};

// Similarly, implement classes for Fields and Drawings

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentFieldsContent(Class, Pos, Items, bAdd)
{
    AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, bAdd);
}
CChangesPDFDocumentFieldsContent.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentFieldsContent.prototype.constructor = CChangesPDFDocumentFieldsContent;
CChangesPDFDocumentFieldsContent.prototype.Type = AscDFH.historyitem_PDF_Document_FieldsContent;
CChangesPDFDocumentFieldsContent.prototype.IsContentChange = function(){
    return true;
};

CChangesPDFDocumentFieldsContent.prototype.ReadFromBinary = function (reader) {
    this.Type = reader.GetLong();
    this.Add = reader.GetBool();
    this.Pos = reader.GetLong();
    AscDFH.CChangesBaseContentChange.prototype.ReadFromBinary.call(this, reader);
};
CChangesPDFDocumentFieldsContent.prototype.WriteToBinary = function (writer) {
    writer.WriteLong(this.Type);
    writer.WriteBool(this.IsAdd());
    writer.WriteLong(this.Pos);
    AscDFH.CChangesBaseContentChange.prototype.WriteToBinary.call(this, writer);
};

CChangesPDFDocumentFieldsContent.prototype.Undo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();

    if (this.IsAdd()) {
        // Undo addition by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.widgets.splice(oDocument.widgets.indexOf(oItem), 1);
            oPage.fields.splice(this.Pos, 1);

            oItem.parentPage = null;
            oItem._page = -1;
            oItem.selectStartPage = -1;

            oItem.AddToRedraw();
        }
    }
    else {
        // Undo removal by adding items back
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oPage.fields.splice(this.Pos, 0, oItem);
            oDocument.widgets.push(oItem);

            oItem.parentPage = oPage;
            oItem._page = oPage.GetIndex();
            oItem.selectStartPage = oItem._page;

            oItem.AddToRedraw();
        }
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentFieldsContent.prototype.Redo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();

    if (this.IsAdd()) {
        // Redo addition by adding items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.widgets.push(oItem);
            oPage.fields.splice(this.Pos, 0, oItem);

            oItem.parentPage = oPage;
            oItem._page = oPage.GetIndex();
            oItem.selectStartPage = oItem._page;
            
            oItem.AddToRedraw();
        }
    }
    else {
        // Redo removal by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.widgets.splice(oDocument.widgets.indexOf(oItem), 1);
            oPage.fields.splice(this.Pos, 1);

            oItem.parentPage = null;
            oItem._page = -1;
            oItem.selectStartPage = -1;

            oItem.AddToRedraw();
        }
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentFieldsContent.prototype.Load = function() {
    if (this.IsAdd()) {
        this.private_InsertInArrayLoad();
    }
    else {
        this.private_RemoveInArrayLoad();
    }
};
CChangesPDFDocumentFieldsContent.prototype.private_InsertInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = oPage.GetDocument();
    let oContentChanges = this.private_GetContentChanges();

    for (let i = 0; i < this.Items.length; ++i) {
        let oItem = this.Items[i];
        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        // Insert into document widgets array
        nPos = Math.min(nPos, oDocument.widgets.length);
        oDocument.widgets.splice(nPos, 0, oItem);

        // Insert into viewer fields array
        let fieldsArray = oPage.fields;
        nPos = Math.min(nPos, fieldsArray.length);
        fieldsArray.splice(nPos, 0, oItem);

        oItem.parentPage = oPage;
        oItem._page = oPage.GetIndex();
        oItem.selectStartPage = oItem._page;

        oItem.AddToRedraw();
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentFieldsContent.prototype.private_RemoveInArrayLoad = function()
{
    if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = oPage.GetDocument();
    let oContentChanges = this.private_GetContentChanges();

    // Remove items in reverse order to maintain indices
    for (let i = this.Items.length - 1; i >= 0; --i) {
        let oItem = this.Items[i];
        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove,  true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        oItem.AddToRedraw();

        // Remove from document widgets array
        let indexInWidgets = oDocument.widgets.indexOf(oItem);
        if (indexInWidgets !== -1)
            oDocument.widgets.splice(indexInWidgets, 1);

        // Remove from viewer fields array
        let fieldsArray = oPage.fields;
        let indexInFields = fieldsArray.indexOf(oItem);
        if (indexInFields !== -1)
            fieldsArray.splice(indexInFields, 1);

        oItem.parentPage = null;
        oItem._page = -1;
        oItem.selectStartPage = -1;
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentFieldsContent.prototype.private_GetContentChanges = function() {
    return this.Class.fieldsContentChanges;
};
CChangesPDFDocumentFieldsContent.prototype.GetContentChangesClass = function() {
	return this.private_GetContentChanges();
};
CChangesPDFDocumentFieldsContent.prototype.private_WriteItem = function (Writer, Item) {
    Writer.WriteString2(Item.Get_Id());
};
CChangesPDFDocumentFieldsContent.prototype.private_ReadItem = function (Reader) {
    var Id = Reader.GetString2();
    return AscCommon.g_oTableId.Get_ById(Id);
};
CChangesPDFDocumentFieldsContent.prototype.Copy = function() {
    let oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

    oChanges.UseArray = this.UseArray;

    for (let nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
        oChanges.PosArray[nIndex] = this.PosArray[nIndex];

    return oChanges;
};
CChangesPDFDocumentFieldsContent.prototype.CreateReverseChange = function(){
    return this.private_CreateReverseChange(this.constructor);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentDrawingsContent(Class, Pos, Items, bAdd)
{
    AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, bAdd);
}
CChangesPDFDocumentDrawingsContent.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentDrawingsContent.prototype.constructor = CChangesPDFDocumentDrawingsContent;
CChangesPDFDocumentDrawingsContent.prototype.Type = AscDFH.historyitem_PDF_Document_DrawingsContent;
CChangesPDFDocumentDrawingsContent.prototype.IsContentChange = function(){
    return true;
};

CChangesPDFDocumentDrawingsContent.prototype.ReadFromBinary = function (reader) {
    this.Type = reader.GetLong();
    this.Add = reader.GetBool();
    this.Pos = reader.GetLong();
    AscDFH.CChangesBaseContentChange.prototype.ReadFromBinary.call(this, reader);
};
CChangesPDFDocumentDrawingsContent.prototype.WriteToBinary = function (writer) {
    writer.WriteLong(this.Type);
    writer.WriteBool(this.IsAdd());
    writer.WriteLong(this.Pos);
    AscDFH.CChangesBaseContentChange.prototype.WriteToBinary.call(this, writer);
};
CChangesPDFDocumentDrawingsContent.prototype.Undo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();

    if (this.IsAdd()) {
        // Undo addition by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.drawings.splice(oDocument.drawings.indexOf(oItem), 1);
            oPage.drawings.splice(this.Pos, 1);

            oItem.parent = oPage;
            oItem.selectStartPage = -1;

            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
        }
    } else {
        // Undo removal by adding items back
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            oItem.AddToRedraw();
            oDocument.drawings.push(oItem);
            oPage.drawings.splice(this.Pos, 0, oItem);

            oItem.parent = oPage;
            oItem.selectStartPage = oItem._page;

            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
        }
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentDrawingsContent.prototype.Redo = function()
{
    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();

    if (this.IsAdd()) {
        // Redo addition by adding items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];

            oItem.AddToRedraw();
            oDocument.drawings.push(oItem);
            oPage.drawings.splice(this.Pos, 0, oItem);

            oItem.parent = oPage;
            oItem.selectStartPage = oItem._page;

            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
        }
    } else {
        // Redo removal by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
        {
            let oItem = this.Items[nIndex];
            
            oItem.AddToRedraw();
            oDocument.drawings.splice(oDocument.drawings.indexOf(oItem), 1);
            oPage.drawings.splice(this.Pos, 1);

            oItem.parent = null;
            oItem.selectStartPage = -1;
            
            oViewer.DrawingObjects.resetSelection();
            oItem.AddToRedraw();
        }
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentDrawingsContent.prototype.Load = function() {
    if (this.IsAdd()) {
        this.private_InsertInArrayLoad();
    }
    else {
        this.private_RemoveInArrayLoad();
    }
};
CChangesPDFDocumentDrawingsContent.prototype.private_InsertInArrayLoad = function()
{
	if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();
    let oContentChanges = this.private_GetContentChanges();

    for (let i = 0; i < this.Items.length; ++i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        // Ensure position is within bounds
        nPos = Math.min(nPos, oDocument.drawings.length);

        // Insert into document drawings array
        oDocument.drawings.splice(nPos, 0, oItem);

        // Insert into viewer drawings array
        let drawingsArray = oPage.drawings;
        nPos = Math.min(nPos, drawingsArray.length);
        drawingsArray.splice(nPos, 0, oItem);
        oItem.parent = oPage;
        oItem.selectStartPage = oItem._page;

        oViewer.DrawingObjects.resetSelection();
        oItem.AddToRedraw();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentDrawingsContent.prototype.private_RemoveInArrayLoad = function()
{
	if (this.Items.length <= 0)
        return;

    let oPage = this.Class;
    let oDocument = this.Class.GetDocument();
    let oViewer = Asc.editor.getDocumentRenderer();
    let oContentChanges = this.private_GetContentChanges();

    // Remove items in reverse order to maintain indices
    for (let i = this.Items.length - 1; i >= 0; --i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove,  true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        oItem.AddToRedraw();
        // Remove from document drawings array
        let indexInDrawings = oDocument.drawings.indexOf(oItem);
        if (indexInDrawings !== -1)
            oDocument.drawings.splice(indexInDrawings, 1);

        // Remove from viewer drawings array
        let drawingsArray = oPage.drawings;
        let indexInDrawingsArray = drawingsArray.indexOf(oItem);
        if (indexInDrawingsArray !== -1)
            drawingsArray.splice(indexInDrawingsArray, 1);

        oItem.parent = null;
        oItem.selectStartPage = -1;

        oViewer.DrawingObjects.resetSelection();
    }

    oDocument.SetMouseDownObject(null);
    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDocumentDrawingsContent.prototype.private_GetContentChanges = function() {
    return this.Class.drawingsContentChanges;
};
CChangesPDFDocumentDrawingsContent.prototype.GetContentChangesClass = function() {
	return this.private_GetContentChanges();
};
CChangesPDFDocumentDrawingsContent.prototype.private_WriteItem = function (Writer, Item) {
    Writer.WriteString2(Item.Get_Id());
};
CChangesPDFDocumentDrawingsContent.prototype.private_ReadItem = function (Reader) {
    var Id = Reader.GetString2();
    return AscCommon.g_oTableId.Get_ById(Id);
};
CChangesPDFDocumentDrawingsContent.prototype.Copy = function() {
    let oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

    oChanges.UseArray = this.UseArray;

    for (let nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
        oChanges.PosArray[nIndex] = this.PosArray[nIndex];

    return oChanges;
};
CChangesPDFDocumentDrawingsContent.prototype.CreateReverseChange = function(){
    return this.private_CreateReverseChange(this.constructor);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDocumentPagesContent(Class, Pos, Items, bAdd)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, bAdd);
}
CChangesPDFDocumentPagesContent.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDocumentPagesContent.prototype.constructor = CChangesPDFDocumentPagesContent;
CChangesPDFDocumentPagesContent.prototype.Type = AscDFH.historyitem_PDF_Document_PagesContent;

CChangesPDFDocumentPagesContent.prototype.Undo = function() {
	let oDocument   = this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	
    if (this.IsAdd()) {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
            oDocument.RemovePage(nPos);
        }
    }
    else {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
            let oItem = this.Items[nIndex];
            oDocument.AddPage(nPos, oItem);
        }
    }
	
	
	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentPagesContent.prototype.Redo = function() {
	let oDocument	= this.Class;
	let oDrDoc		= oDocument.GetDrawingDocument();
	let oContentChanges = this.private_GetContentChanges();

    if (this.IsAdd()) {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex]);
            if (nPos === false) continue;

            let oItem = this.Items[nIndex];
            oDocument.AddPage(nPos, oItem)
        }
    }
    else {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove, true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex]);
            if (nPos === false) continue;
            oDocument.RemovePage(nPos)
        }
    }
	

	oDocument.SetMouseDownObject(null);
	oDrDoc.TargetEnd();
};
CChangesPDFDocumentPagesContent.prototype.Load = function()
{
	let pdfDocument = this.Class;
	let pageChanges = pdfDocument.pagesContentChanges;

    if (this.IsAdd()) {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
            nPos = pageChanges.Check(AscCommon.contentchanges_Add, nPos);
            pdfDocument.AddPage(nPos, this.Items[nIndex]);
        }
    }
    else {
        for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let nPos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
            nPos = pageChanges.Check(AscCommon.contentchanges_Remove, nPos);
            if (false === nPos)
                continue;
            
            pdfDocument.RemovePage(nPos);
        }
    }
	
};
CChangesPDFDocumentPagesContent.prototype.private_WriteItem = function(Writer, oPage)
{
    let nFlags = 0;
    if (undefined == oPage.originIndex) {
        nFlags |= 1;
    }

    Writer.WriteLong(nFlags);

    Writer.WriteString2(oPage.Id);
	Writer.WriteLong(oPage.Rotate);
    if (undefined != oPage.originIndex) {
        Writer.WriteLong(oPage.originIndex);
        Writer.WriteLong(oPage.originRotate);
    }
	Writer.WriteBool(!!oPage.isRecognized);
	Writer.WriteLong(oPage.Dpi);
	Writer.WriteLong(oPage.W);
	Writer.WriteLong(oPage.H);
};
CChangesPDFDocumentPagesContent.prototype.private_ReadItem = function(Reader)
{
    let nFlags = Reader.GetLong();
    let hasOriginIndex = !(nFlags & 1);

    let oPage = {
        Id: Reader.GetString2(),
		Rotate: Reader.GetLong(),
        originIndex: hasOriginIndex ? Reader.GetLong() : undefined,
        originRotate: hasOriginIndex ? Reader.GetLong() : undefined,
        isRecognized: Reader.GetBool(),
		Dpi: Reader.GetLong(),
		W: Reader.GetLong(),
		H: Reader.GetLong(),
        fonts: []
    }
	
    oPage["originIndex"] = oPage.originIndex;

    return oPage;
};
CChangesPDFDocumentPagesContent.prototype.ReadFromBinary = function (reader) {
    this.Add = reader.GetBool();
    this.Pos = reader.GetLong();
    AscDFH.CChangesBaseContentChange.prototype.ReadFromBinary.call(this, reader);
};
CChangesPDFDocumentPagesContent.prototype.WriteToBinary = function (writer) {
    writer.WriteBool(this.IsAdd());
    writer.WriteLong(this.Pos);
    AscDFH.CChangesBaseContentChange.prototype.WriteToBinary.call(this, writer);
};
CChangesPDFDocumentPagesContent.prototype.Copy = function() {
    let oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

    oChanges.UseArray = this.UseArray;

    for (let nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
        oChanges.PosArray[nIndex] = this.PosArray[nIndex];

    return oChanges;
};
CChangesPDFDocumentPagesContent.prototype.CreateReverseChange = function(){
    return this.private_CreateReverseChange(this.constructor);
};
CChangesPDFDocumentPagesContent.prototype.private_GetContentChanges = function() {
    return this.Class.pagesContentChanges;
};
CChangesPDFDocumentPagesContent.prototype.GetContentChangesClass = function() {
	return this.private_GetContentChanges();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFDocumentRotatePage(Class, Old, New)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesPDFDocumentRotatePage.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentRotatePage.prototype.constructor = CChangesPDFDocumentRotatePage;
CChangesPDFDocumentRotatePage.prototype.Type = AscDFH.historyitem_PDF_Document_RotatePage;
CChangesPDFDocumentRotatePage.prototype.private_SetValue = function(Value)
{
	let oPage       = this.Class;
    let oDoc        = oPage.GetDocument();
    let oFile       = oDoc.Viewer.file;
    let nPageIdx    = oPage.GetIndex();

    if (-1 !== nPageIdx) {
        oFile.pages[nPageIdx].Rotate = Value;

        // sticky note всегда неповернуты
        oDoc.Viewer.pagesInfo.pages[nPageIdx].annots.forEach(function(annot) {
            if (annot.IsComment()) {
                annot.AddToRedraw();
            }
        });
        
		oDoc.Viewer.resize(true);
        oDoc.Viewer.paint();
    }
};
CChangesPDFDocumentRotatePage.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (undefined === this.New)
		nFlags |= 1;

	if (undefined === this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (undefined !== this.New)
		Writer.WriteLong(this.New);

	if (undefined !== this.Old)
		Writer.WriteLong(this.Old);
};
CChangesPDFDocumentRotatePage.prototype.ReadFromBinary = function(Reader)
{
	let nFlags = Reader.GetLong();
	
	if (nFlags & 1)
		this.New = undefined;
	else
		this.New = Reader.GetLong();
	
	if (nFlags & 2)
		this.Old = undefined;
	else
		this.Old = Reader.GetLong();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFDocumentRecognizePage(Class, Old, New)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesPDFDocumentRecognizePage.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentRecognizePage.prototype.constructor = CChangesPDFDocumentRecognizePage;
CChangesPDFDocumentRecognizePage.prototype.Type = AscDFH.historyitem_PDF_Document_RecognizePage;
CChangesPDFDocumentRecognizePage.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (true === this.New)
		nFlags |= 1;
	
	if (true === this.Old)
		nFlags |= 2;
	
	Writer.WriteLong(nFlags);
};
CChangesPDFDocumentRecognizePage.prototype.ReadFromBinary = function(Reader)
{
	let nFlags = Reader.GetLong();
	
	this.New = !!(nFlags & 1);
	this.Old = !!(nFlags & 2);
};
CChangesPDFDocumentRecognizePage.prototype.private_SetValue = function(bRecognize)
{
	let oPage       = this.Class;
    let oDoc        = oPage.GetDocument();
    let oFile       = oDoc.Viewer.file;
    let nPageIdx    = oPage.GetIndex();
	
    if (-1 !== nPageIdx) {
        oFile.pages[nPageIdx].isRecognized = bRecognize;
        if (oDoc.Viewer.drawingPages[nPageIdx]) {
            delete oDoc.Viewer.drawingPages[nPageIdx].Image;
        }

        oDoc.Viewer.paint(function() {
            let oThumbnails = oDoc.Viewer.thumbnails;
            oThumbnails && oThumbnails._repaintPage(nPageIdx);
        });
    }
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFObjectSetDocument(Class, Old, New)
{
	let oldId = Old && Old.Get_Id ? Old.Get_Id() : undefined;
	let newId = New && New.Get_Id ? New.Get_Id() : undefined;
	AscDFH.CChangesBaseStringProperty.call(this, Class, oldId, newId);
}
CChangesPDFObjectSetDocument.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFObjectSetDocument.prototype.constructor = CChangesPDFObjectSetDocument;
CChangesPDFObjectSetDocument.prototype.Type = AscDFH.historyitem_PDF_Document_SetDocument;
CChangesPDFObjectSetDocument.prototype.private_SetValue = function(value)
{
	let doc = AscCommon.g_oTableId.Get_ById(value);
	this.Class.SetDocument(doc);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFDocumentPageLocks(Class, deleteLock, rotateLock, editPageLock) {
    this.deleteLock = deleteLock;
    this.rotateLock = rotateLock;
    this.editPageLock = editPageLock;
    AscDFH.CChangesBaseProperty.call(this, Class);
};

CChangesPDFDocumentPageLocks.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentPageLocks.prototype.constructor = CChangesPDFDocumentPageLocks;
CChangesPDFDocumentPageLocks.prototype.Type = AscDFH.historyitem_PDF_Document_PageLocks;

CChangesPDFDocumentPageLocks.prototype.WriteToBinary = function(Writer){
    AscFormat.writeObject(Writer, this.deleteLock);
    AscFormat.writeObject(Writer, this.rotateLock);
    AscFormat.writeObject(Writer, this.editPageLock);
};

CChangesPDFDocumentPageLocks.prototype.ReadFromBinary = function(Reader){
    this.deleteLock = AscFormat.readObject(Reader);
    this.rotateLock = AscFormat.readObject(Reader);
    this.editPageLock = AscFormat.readObject(Reader);
};

CChangesPDFDocumentPageLocks.prototype.Undo = function(){
    let oPage = this.Class;
    oPage.deleteLock = null;
    oPage.rotateLock = null;
    oPage.editPageLock = null;
};

CChangesPDFDocumentPageLocks.prototype.Redo = function(){
    let oPage = this.Class;
    oPage.deleteLock = this.deleteLock;
    oPage.rotateLock = this.rotateLock;
    oPage.editPageLock = this.editPageLock;
};
CChangesPDFDocumentPageLocks.prototype.Load = function(){
    this.Redo();
    this.RefreshRecalcData();
};
CChangesPDFDocumentPageLocks.prototype.CreateReverseChange = function()
{
    return new this.constructor(this.Class, null, null, null, null, null, null);
};

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseStringProperty}
 */
function CChangesPDFPropLockerObjectId(Class, Old, New)
{
	AscDFH.CChangesBaseStringProperty.call(this, Class, Old, New);
}
CChangesPDFPropLockerObjectId.prototype = Object.create(AscDFH.CChangesBaseStringProperty.prototype);
CChangesPDFPropLockerObjectId.prototype.constructor = CChangesPDFPropLockerObjectId;
CChangesPDFPropLockerObjectId.prototype.Type = AscDFH.historyitem_PDF_PropLocker_ObjectId;
CChangesPDFPropLockerObjectId.prototype.private_SetValue = function(value)
{
	this.Class.objectId = value;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseLongProperty}
 */
function CChangesPDFDocumentMovePage(Class, Old, New)
{
	AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
}
CChangesPDFDocumentMovePage.prototype = Object.create(AscDFH.CChangesBaseLongProperty.prototype);
CChangesPDFDocumentMovePage.prototype.constructor = CChangesPDFDocumentMovePage;
CChangesPDFDocumentMovePage.prototype.Type = AscDFH.historyitem_PDF_Document_MovePage;
CChangesPDFDocumentMovePage.prototype.private_SetValue = function(nNewPos)
{
    let oDoc        = this.Class.GetDocument();
    let nCurPos     = this.Class.GetIndex();
    let aFilePages  = oDoc.Viewer.file.pages;
    let aPagesInfo  = oDoc.Viewer.pagesInfo.pages;

    let oMovedFilePage = aFilePages.splice(nCurPos, 1)[0];
    let oMovedPageInfo = aPagesInfo.splice(nCurPos, 1)[0];

    aFilePages.splice(nNewPos, 0, oMovedFilePage);
    aPagesInfo.splice(nNewPos, 0, oMovedPageInfo);

    let aPagesRange = [];
    let nStart = Math.min(nCurPos, nNewPos);
    let nEnd = Math.max(nCurPos, nNewPos);
    for (let i = nStart; i <= nEnd; i++) {
        aPagesRange.push(i);
    }

    oDoc.Viewer.resize(true);
    oDoc.Viewer.onUpdatePages(aPagesRange);
    oDoc.Viewer.onRepaintForms(aPagesRange);
    oDoc.Viewer.onRepaintAnnots(aPagesRange);
};

function CChangesPDFDocumentStartMergePages(Class, Old, New, Color) {
    AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFDocumentStartMergePages.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentStartMergePages.prototype.constructor = CChangesPDFDocumentStartMergePages;
CChangesPDFDocumentStartMergePages.prototype.Type = AscDFH.historyitem_Pdf_Document_Start_Merge_Pages;

CChangesPDFDocumentStartMergePages.prototype.Undo = function () {
    delete this.Class.partsOfBinaryData;
};
CChangesPDFDocumentStartMergePages.prototype.Redo = function () {
    this.Class.partsOfBinaryData = [];
};

function CChangesPDFDocumentPartMergePages(Class, Old, New, Color) {
    AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesPDFDocumentPartMergePages.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentPartMergePages.prototype.constructor = CChangesPDFDocumentPartMergePages;
CChangesPDFDocumentPartMergePages.prototype.Type = AscDFH.historyitem_ImageShapeSetPartBinaryData;

CChangesPDFDocumentPartMergePages.prototype.private_SetValue = function (aUint8Array) {
    if (aUint8Array.length) {
        this.Class.partsOfBinaryData.push(aUint8Array);
    }
};
CChangesPDFDocumentPartMergePages.prototype.WriteToBinary = function(Writer) {
    Writer.WriteLong(this.Old.length);
    Writer.WriteBuffer(this.Old, 0, this.Old.length);

    Writer.WriteLong(this.New.length);
    Writer.WriteBuffer(this.New, 0, this.New.length);
};
CChangesPDFDocumentPartMergePages.prototype.ReadFromBinary = function(Reader) {
    let length = Reader.GetLong();
    this.Old = new Uint8Array(Reader.GetBuffer(length));

    length = Reader.GetLong();
    this.New = new Uint8Array(Reader.GetBuffer(length));
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesPDFDocumentEndMergePages(Class, nMaxIdx, sMergeName, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, undefined, undefined, Color);
    this.MaxIdx = nMaxIdx;
    this.MergeName = sMergeName;
}
CChangesPDFDocumentEndMergePages.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesPDFDocumentEndMergePages.prototype.constructor = CChangesPDFDocumentEndMergePages;
CChangesPDFDocumentEndMergePages.prototype.Type = AscDFH.historyitem_Pdf_Document_End_Merge_Pages;
CChangesPDFDocumentEndMergePages.prototype.Undo = function() {
    let oDoc = this.Class;
	let oFile = oDoc.Viewer.file;

    // clear united binary
    delete oDoc.unitedBinary;
    
    oDoc.mergedPagesData.pop();
    oFile.nativeFile["UndoMergePages"]();

    let aPages = oFile.nativeFile["getPagesInfo"]();
    oFile.originalPagesCount = aPages.length;
};
CChangesPDFDocumentEndMergePages.prototype.Redo = function()
{
    let oDoc = this.Class;
	let oFile = oDoc.Viewer.file;

    // union binary
    if (!oDoc.partsOfBinaryData) {
        return this.Undo();
    }
    let lenOfAllBinaryData = 0;
    for (let i = 0; i < oDoc.partsOfBinaryData.length; i += 1) {
        lenOfAllBinaryData += oDoc.partsOfBinaryData[i].length;
    }
    oDoc.unitedBinary = new Uint8Array(lenOfAllBinaryData);

    let indexOfInsert = 0;
    for (let i = 0; i < oDoc.partsOfBinaryData.length; i += 1) {
        const partOfBinaryData = oDoc.partsOfBinaryData[i];
        for (let j = 0; j < partOfBinaryData.length; j += 1) {
            oDoc.unitedBinary[indexOfInsert] = partOfBinaryData[j];
            indexOfInsert += 1;
        }
    }

    delete oDoc.partsOfBinaryData;

    oFile.nativeFile["MergePages"](oDoc.unitedBinary, this.MaxIdx, this.MergeName);
    let aPages = oFile.nativeFile["getPagesInfo"]();
    oFile.originalPagesCount = aPages.length;

    oDoc.mergedPagesData.push({
        mergeName: this.MergeName,
        maxId: this.MaxIdx,
        binary: oDoc.unitedBinary
    });

    oDoc.Viewer.checkLoadCMap();
};
CChangesPDFDocumentEndMergePages.prototype.WriteToBinary = function(Writer)
{
	let nFlags = 0;

	if (false !== this.Color)
		nFlags |= 1;

	if (undefined === this.MaxIdx)
		nFlags |= 2;

	if (undefined === this.MergeName)
		nFlags |= 4;

	Writer.WriteLong(nFlags);

	if (undefined !== this.MaxIdx)
		Writer.WriteLong(this.MaxIdx);

	if (undefined !== this.MergeName)
		Writer.WriteString2(this.MergeName);
};
CChangesPDFDocumentEndMergePages.prototype.ReadFromBinary = function(Reader)
{
	this.FromLoad = true;

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.Color = true;
	else
		this.Color = false;

	if (nFlags & 2)
		this.MaxIdx = undefined;
	else
		this.MaxIdx = Reader.GetLong();

	if (nFlags & 4)
		this.MergeName = undefined;
	else
		this.MergeName = Reader.GetString2();
};

/**
 * @constructor
 * @extends {AscDFH.CChangesPDFArrayOfDoubleProperty}
 */
function CChangesPDFCalcOrder(Class, Old, New, Color)
{
	AscDFH.CChangesPDFArrayOfDoubleProperty.call(this, Class, Old, New, Color);
}
CChangesPDFCalcOrder.prototype = Object.create(AscDFH.CChangesPDFArrayOfDoubleProperty.prototype);
CChangesPDFCalcOrder.prototype.constructor = CChangesPDFCalcOrder;
CChangesPDFCalcOrder.prototype.Type = AscDFH.historyitem_Pdf_Document_Calc_Order;
CChangesPDFCalcOrder.prototype.private_SetValue = function(Value)
{
	let oCalcInfo = this.Class;
	oCalcInfo.ids = Value;
};
