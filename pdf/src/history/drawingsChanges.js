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

AscDFH.changesFactory[AscDFH.historyitem_type_Pdf_Drawing_Page]		= CChangesPDFDrawingPage;
AscDFH.changesFactory[AscDFH.historyitem_type_Pdf_Drawing_Redacts]	= CChangesPDFDrawingRedacts;

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

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesPDFDrawingRedacts(Class, Pos, Items, bAdd)
{
    AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, bAdd);
}
CChangesPDFDrawingRedacts.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesPDFDrawingRedacts.prototype.constructor = CChangesPDFDrawingRedacts;
CChangesPDFDrawingRedacts.prototype.Type = AscDFH.historyitem_type_Pdf_Drawing_Redacts;
CChangesPDFDrawingRedacts.prototype.IsContentChange = function(){
    return true;
};

CChangesPDFDrawingRedacts.prototype.ReadFromBinary = function (reader) {
    this.Type = reader.GetLong();
    this.Add = reader.GetBool();
    this.Pos = reader.GetLong();
    AscDFH.CChangesBaseContentChange.prototype.ReadFromBinary.call(this, reader);
};
CChangesPDFDrawingRedacts.prototype.WriteToBinary = function (writer) {
    writer.WriteLong(this.Type);
    writer.WriteBool(this.IsAdd());
    writer.WriteLong(this.Pos);
    AscDFH.CChangesBaseContentChange.prototype.WriteToBinary.call(this, writer);
};
CChangesPDFDrawingRedacts.prototype.Undo = function()
{
    let oDrawing = this.Class;
    let oDocument = Asc.editor.getPDFDoc();

    if (this.IsAdd()) {
        // Undo addition by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            oDrawing._redactIds.splice(this.Pos, 1);
        }
    } else {
        // Undo removal by adding items back
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let sId = this.Items[nIndex];
            oDrawing._redactIds.splice(this.Pos, 0, sId);
        }
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDrawingRedacts.prototype.Redo = function()
{
    let oDrawing = this.Class;
    let oDocument = Asc.editor.getPDFDoc();

    if (this.IsAdd()) {
        // Redo addition by adding items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            let sId = this.Items[nIndex];
            oDrawing._redactIds.splice(this.Pos, 0, sId);
        }
    } else {
        // Redo removal by removing items
        for (let nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex) {
            oDrawing._redactIds.splice(this.Pos, 1);
        }
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDrawingRedacts.prototype.Load = function() {
    if (this.IsAdd()) {
        this.private_InsertInArrayLoad();
    }
    else {
        this.private_RemoveInArrayLoad();
    }
};
CChangesPDFDrawingRedacts.prototype.private_InsertInArrayLoad = function()
{
	if (this.Items.length <= 0)
        return;

    let oDrawing = this.Class;
    let oDocument = Asc.editor.getPDFDoc();
    let oContentChanges = this.private_GetContentChanges();

    for (let i = 0; i < this.Items.length; ++i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        nPos = Math.min(nPos, oDrawing._redactIds.length);
        oDrawing._redactIds.splice(nPos, 0, oItem);
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDrawingRedacts.prototype.private_RemoveInArrayLoad = function()
{
	if (this.Items.length <= 0)
        return;

    let oDrawing = this.Class;
    let oDocument = Asc.editor.getPDFDoc();
    let oContentChanges = this.private_GetContentChanges();

    // Remove items in reverse order to maintain indices
    for (let i = this.Items.length - 1; i >= 0; --i) {
        let oItem = this.Items[i];

        // Adjust position based on content changes
        let nPos = oContentChanges.Check(AscCommon.contentchanges_Remove,  true !== this.UseArray ? this.Pos + i : this.PosArray[i]);
        if (nPos === false) continue;

        let indexInArray = oDrawing._redactIds.indexOf(oItem);
        if (indexInArray !== -1)
            oDrawing._redactIds.splice(indexInArray, 1);
    }

    oDocument.private_UpdateTargetForCollaboration(true);
};
CChangesPDFDrawingRedacts.prototype.private_GetContentChanges = function() {
    return this.Class.redactsArrayChanges;
};
CChangesPDFDrawingRedacts.prototype.GetContentChangesClass = function() {
	return this.private_GetContentChanges();
};
CChangesPDFDrawingRedacts.prototype.private_WriteItem = function (Writer, sId) {
    Writer.WriteString2(sId);
};
CChangesPDFDrawingRedacts.prototype.private_ReadItem = function (Reader) {
    return Reader.GetString2();
};
CChangesPDFDrawingRedacts.prototype.Copy = function() {
    let oChanges = new this.constructor(this.Class, this.Pos, this.Items, this.Add);

    oChanges.UseArray = this.UseArray;

    for (let nIndex = 0, nCount = this.PosArray.length; nIndex < nCount; ++nIndex)
        oChanges.PosArray[nIndex] = this.PosArray[nIndex];

    return oChanges;
};
CChangesPDFDrawingRedacts.prototype.CreateReverseChange = function(){
    return this.private_CreateReverseChange(this.constructor);
};
