/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

/**
 *
 * @constructor
 * @extends {AscCommon.CCollaborativeEditingBase}
 */
function CPDFCollaborativeEditing(oDoc) {
	AscCommon.CWordCollaborativeEditing.call(this);
    this.m_aSkipContentControlsOnCheckEditingLock = {};
    this.m_oLogicDocument = oDoc;
    this.m_oSelectedObjects = {};
    this.m_aEndLoadCallbacks = [];
    this.m_aFieldNameConflicts = [];
}

CPDFCollaborativeEditing.prototype = Object.create(AscCommon.CWordCollaborativeEditing.prototype);
CPDFCollaborativeEditing.prototype.constructor = CPDFCollaborativeEditing;
CPDFCollaborativeEditing.prototype.Clear = function() {
    AscCommon.CWordCollaborativeEditing.prototype.Clear.apply(this, arguments);
    this.m_aFieldNameConflicts.length = 0;
};

CPDFCollaborativeEditing.prototype.CheckWaitingImages = function (aImages) {
    if (aImages.length !== 0) {
        this.waitingImagesForLoad = true;
    }
};
CPDFCollaborativeEditing.prototype.SendImagesCallback = function (aImages) {
    this.waitingImagesForLoad = false;
    let oApi = Asc.editor || Asc['editor'];
    oApi.pre_Save(aImages);
};
CPDFCollaborativeEditing.prototype.Add_ForeignSelectedObject = function(UserId, oObject, UserShortId) {
    if (!this.m_oSelectedObjects[UserId]) {
        this.m_oSelectedObjects[UserId] = [];
    }

    this.m_aForeignCursorsId[UserId] = UserShortId;
    this.m_oSelectedObjects[UserId].push(oObject);
};
CPDFCollaborativeEditing.prototype.Remove_FiregnSelectedObject = function(UserId, oObject) {
    if (!oObject) {
        this.m_oSelectedObjects[UserId] = [];
    }
    else {
        this.m_oSelectedObjects[UserId].splice(this.m_oSelectedObjects[UserId].indexOf(oObject), 1);
    }
};
CPDFCollaborativeEditing.prototype.Check_ForeignSelectedObjectsLabels = function(X, Y, PageIndex) {
    let oDoc = this.GetLogicDocument();

    for (let UserId in this.m_oSelectedObjects) {
        let aObjects = this.m_oSelectedObjects[UserId];
        for (let i = 0; i < aObjects.length; i++) {
            if (aObjects[i].GetPage() !== PageIndex) {
                continue;
            }

            let aOrigRect = aObjects[i].GetRect();
            if (X >= aOrigRect[0] && X < aOrigRect[2] && Y > aOrigRect[1] && Y < aOrigRect[3]) {
                let color = AscCommon.getUserColorById(this.m_aForeignCursorsId[UserId], null, true);
                oDoc.Show_ForeignSelectedObjectLabel(UserId, aObjects[i], color);
            }
        }
    }
};
CPDFCollaborativeEditing.prototype.Update_ForeignSelectedObjectsLabelsPositions = function(PageIndex) {
    let oDoc = this.GetLogicDocument();

    for (let UserId in this.m_oSelectedObjects) {
        let aObjects = this.m_oSelectedObjects[UserId];
        for (let i = 0; i < aObjects.length; i++) {
            if (aObjects[i].GetPage() !== PageIndex || aObjects[i].LabelTimer == null) {
                continue;
            }

            let color = AscCommon.getUserColorById(this.m_aForeignCursorsId[UserId], null, true);
            oDoc.Show_ForeignSelectedObjectLabel(UserId, aObjects[i], color);
        }
    }
};
CPDFCollaborativeEditing.prototype.Update_ForeignCursorPosition = function(UserId, Run, InRunPos, isRemoveLabel) {
    let DrawingDocument = this.m_oLogicDocument.DrawingDocument;

    if (!(Run instanceof AscCommonWord.ParaRun))
        return;

    let Paragraph = Run.GetParagraph();

    if (!Paragraph) {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }

    let ParaContentPos = Paragraph.Get_PosByElement(Run);
    if (!ParaContentPos) {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }
    ParaContentPos.Update(InRunPos, ParaContentPos.GetDepth() + 1);

    let XY = Paragraph.Get_XYByContentPos(ParaContentPos);
    if (XY && XY.Height > 0.001 && XY.PageNum >= 0) {
        let ShortId = this.m_aForeignCursorsId[UserId] ? this.m_aForeignCursorsId[UserId] : UserId;
        DrawingDocument.Collaborative_UpdateTarget(UserId, ShortId, XY.X, XY.Y, XY.Height, XY.PageNum, Paragraph.Get_ParentTextTransform());
        this.Add_ForeignCursorXY(UserId, XY.X, XY.Y, XY.PageNum, XY.Height, Paragraph, isRemoveLabel);

        if (true === this.m_aForeignCursorsToShow[UserId]) {
            this.Show_ForeignCursorLabel(UserId);
            this.Remove_ForeignCursorToShow(UserId);
        }
    }
    else {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        this.Remove_ForeignCursorXY(UserId);
        this.Remove_ForeignCursorToShow(UserId);
    }
};
CPDFCollaborativeEditing.prototype.private_LockByMe = function() {
	for (let nIndex = 0, nCount = this.m_aCheckLocks.length; nIndex < nCount; ++nIndex) {
		let oItem = this.m_aCheckLocks[nIndex];
        
		if (true !== oItem && false !== oItem) {
			let oClass = AscCommon.g_oTableId.Get_ById(oItem["guid"]);
			if (oClass) {
				oClass.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeMine);
				this.Add_Unlock2(oClass);
			}
		}
	}
};
CPDFCollaborativeEditing.prototype.GetDocumentPositionBinary = function(oWriter, PosInfo) {
    if (!PosInfo)
        return '';
    var BinaryPos = oWriter.GetCurPosition();
    oWriter.WriteString2(PosInfo.Class.Get_Id());
    oWriter.WriteLong(PosInfo.Position);
    var BinaryLen = oWriter.GetCurPosition() - BinaryPos;
    return  (BinaryLen + ";" + oWriter.GetBase64Memory2(BinaryPos, BinaryLen));
};
CPDFCollaborativeEditing.prototype.GetDocument = function() {
    return this.m_oLogicDocument;
};
CPDFCollaborativeEditing.prototype.private_StartCheckFieldNames = function() {
	let oHistory = this.GetDocument().History;
    let nStartPoint = null === oHistory.SavedIndex ? 0 : oHistory.SavedIndex + 1;
    let nLastPoint = this.m_nUseType <= 0 ? oHistory.Points.length - 1 : oHistory.Index;
    let oFields = {};
    let oParents = {};

    function addField(oField) {
        if (!oField || !oField.IsForm || !oField.IsForm())
            return;

        let aWidgets = oField.GetAllWidgets();
        for (let nWidget = 0; nWidget < aWidgets.length; ++nWidget)
            oFields[aWidgets[nWidget].GetId()] = aWidgets[nWidget];
    }

    function addFields(aFields) {
        if (!aFields)
            return;

        for (let nField = 0; nField < aFields.length; ++nField)
            addField(aFields[nField]);
    }

    for (let nPoint = nStartPoint; nPoint <= nLastPoint; ++nPoint) {
        let oPoint = oHistory.Points[nPoint];
        if (!oPoint)
            continue;

        for (let nItem = 0; nItem < oPoint.Items.length; ++nItem) {
            let oChange = oPoint.Items[nItem].Data;
            if (!oChange)
                continue;

            if (oChange.Type === AscDFH.historyitem_TableId_Add) {
                addField(oChange.NewClass);
                if (oChange.NewClass && oChange.NewClass.IsForm && oChange.NewClass.IsForm()
                    && !oChange.NewClass.IsWidget()) {
                    oParents[oChange.NewClass.GetId()] = oChange.NewClass;
                }
            }
            else if (oChange.Type === AscDFH.historyitem_PDF_Document_FieldsContent && oChange.IsAdd()) {
                addFields(oChange.Items);
            }
            else if (oChange.Type === AscDFH.historyitem_Pdf_Form_Kids_Content && oChange.IsAdd()) {
                addFields(oChange.Items);
            }
            else if (oChange.Type === AscDFH.historyitem_Pdf_Form_Partial_Name) {
                addField(oChange.Class);
            }
        }
    }

    return {
        OwnFields:        oFields,
        OwnParents:       oParents,
        FieldNodesBefore: private_CollectFieldNodes(this.GetDocument().widgets)
    };
};
CPDFCollaborativeEditing.prototype.private_EndCheckFieldNames = function(oCheckInfo) {
    let oDoc = this.GetDocument();
    
	let oOwnFields			= oCheckInfo.OwnFields;
    let oOwnParents			= oCheckInfo.OwnParents;
    let oFieldNodesBefore	= oCheckInfo.FieldNodesBefore;

    let aFieldsAfter = oDoc.widgets;
    let aForeignFields = [];
    let oFieldNodesAfter = private_CollectFieldNodes(aFieldsAfter);
    let oForeignParentsByName = {};

    for (let nField = 0; nField < aFieldsAfter.length; ++nField) {
        let oField = aFieldsAfter[nField];
        let sId = oField.GetId();
        let oFieldBefore = oFieldNodesBefore[sId];

        if (!oOwnFields[sId]
            && (!oFieldBefore || oFieldBefore.Name !== oField.GetFullName())) {
            aForeignFields.push(oField);
        }
    }

    for (let sNodeId in oFieldNodesAfter) {
        let oNodeInfo = oFieldNodesAfter[sNodeId];
        let oField = oNodeInfo.Field;
        let oNodeBefore = oFieldNodesBefore[sNodeId];

        if (!oField.IsWidget() && !oOwnParents[sNodeId]
            && (!oNodeBefore || oNodeBefore.Name !== oNodeInfo.Name)) {
            if (!oForeignParentsByName[oNodeInfo.Name])
                oForeignParentsByName[oNodeInfo.Name] = [];
            oForeignParentsByName[oNodeInfo.Name].push(oField);
        }
    }

    let oParentConflicts = {};
    let oFieldsWithParentConflict = {};
    for (let sOwnId in oOwnFields) {
        let oOwnField = oOwnFields[sOwnId];
        let aOwnParents = [];
        let oParent = oOwnField.GetParent();

        while (oParent) {
            if (oOwnParents[oParent.GetId()] === oParent)
                aOwnParents.push(oParent);
            oParent = oParent.GetParent();
        }

        aOwnParents.reverse();
        for (let nParent = 0; nParent < aOwnParents.length; ++nParent) {
            let oOwnParent = aOwnParents[nParent];
            let sParentName = oOwnParent.GetFullName();
            let aForeignParents = oForeignParentsByName[sParentName];
            if (!aForeignParents)
                continue;

            oFieldsWithParentConflict[sOwnId] = true;
            if (!oParentConflicts[oOwnParent.GetId()]) {
                oParentConflicts[oOwnParent.GetId()] = true;
                this.m_aFieldNameConflicts.push({
                    Foreign:    aForeignParents[0],
                    Parent:     oOwnParent,
                    ParentName: sParentName
                });
            }
            break;
        }
    }

    for (let sOwnId in oOwnFields) {
        if (oFieldsWithParentConflict[sOwnId])
            continue;

        let oOwnField = oOwnFields[sOwnId];
        if (-1 === aFieldsAfter.indexOf(oOwnField))
            continue;

        let sOwnName = oOwnField.GetFullName();
        for (let nForeign = 0; nForeign < aForeignFields.length; ++nForeign) {
            let oForeignField = aForeignFields[nForeign];
            if (sOwnName === oForeignField.GetFullName()) {
                this.m_aFieldNameConflicts.push({
                    Own:     oOwnField,
                    Foreign: oForeignField,
                    Name:    sOwnName
                });
            }
        }
    }
};
CPDFCollaborativeEditing.prototype.Apply_OtherChanges = function() {
    let oCheckInfo = this.private_StartCheckFieldNames();

    AscCommon.CWordCollaborativeEditing.prototype.Apply_OtherChanges.apply(this, arguments);

    this.private_EndCheckFieldNames(oCheckInfo);
};
CPDFCollaborativeEditing.prototype.private_ResolveFieldNameConflicts = function() {
    if (0 === this.m_aFieldNameConflicts.length)
        return;

    let oDoc = this.GetDocument();
    let aFields = oDoc.widgets;
    let oFieldNodes = {};
    let oConflictingFields = {};
    let aParentConflicts = [];
    let bParentRenamed = false;

    function createNameWithPostfix(sName, isNameUsed) {
        let oPostfixMatch = /^(.*)_(\d+)$/.exec(sName);
        let sBaseName = oPostfixMatch ? oPostfixMatch[1] : sName;
        let nPostfix = oPostfixMatch ? Number(oPostfixMatch[2]) + 1 : 1;
        let sNewName = sBaseName + "_" + nPostfix;
        
        while (isNameUsed(sNewName)) {
            sNewName = sBaseName + "_" + ++nPostfix;
        }

        return sNewName;
    }

    for (let nField = 0; nField < aFields.length; ++nField) {
        let oField = aFields[nField];
        while (oField) {
            oFieldNodes[oField.GetId()] = oField;
            oField = oField.GetParent();
        }
    }

    for (let nConflict = 0; nConflict < this.m_aFieldNameConflicts.length; ++nConflict) {
        if (this.m_aFieldNameConflicts[nConflict].Parent)
            aParentConflicts.push(this.m_aFieldNameConflicts[nConflict]);
    }

    aParentConflicts.sort(function(oFirst, oSecond) {
        return oFirst.ParentName.split(".").length - oSecond.ParentName.split(".").length;
    });

    for (let nConflict = 0; nConflict < aParentConflicts.length; ++nConflict) {
        let oConflict = aParentConflicts[nConflict];
        let oOwnParent = oConflict.Parent;
        let oForeignParent = oConflict.Foreign;

        if (oFieldNodes[oOwnParent.GetId()] !== oOwnParent
            || oFieldNodes[oForeignParent.GetId()] !== oForeignParent
            || oConflict.ParentName !== oOwnParent.GetFullName()
            || oConflict.ParentName !== oForeignParent.GetFullName()) {
            continue;
        }

        let oOwnGrandParent = oOwnParent.GetParent();
        let sParentPrefix = oOwnGrandParent ? oOwnGrandParent.GetFullName() + "." : "";
        let sNewPartialName = createNameWithPostfix(oOwnParent.GetPartialName(), function(sName) {
            return !!oDoc.GetField(sParentPrefix + sName);
        });

        oOwnParent.SetPartialName(sNewPartialName);
        bParentRenamed = true;
    }

    for (let nConflict = 0; nConflict < this.m_aFieldNameConflicts.length; ++nConflict) {
        let oConflict = this.m_aFieldNameConflicts[nConflict];
        if (oConflict.Parent)
            continue;

        let oOwnField = oConflict.Own;
        let oForeignField = oConflict.Foreign;

        if (-1 === aFields.indexOf(oOwnField)
            || -1 === aFields.indexOf(oForeignField)
            || oConflict.Name !== oOwnField.GetFullName()
            || oConflict.Name !== oForeignField.GetFullName()) {
            continue;
        }

        let sName = oConflict.Name;
        let nFieldType = oOwnField.GetType();
        if (!oConflictingFields[sName])
            oConflictingFields[sName] = {};
        if (!oConflictingFields[sName][nFieldType])
            oConflictingFields[sName][nFieldType] = [];
        if (-1 === oConflictingFields[sName][nFieldType].indexOf(oOwnField))
            oConflictingFields[sName][nFieldType].push(oOwnField);
    }

    for (let sName in oConflictingFields) {
        for (let nFieldType in oConflictingFields[sName]) {
            let aOwnFields = oConflictingFields[sName][nFieldType];
            let sNewName = createNameWithPostfix(sName, function(sCandidate) {
                return !!oDoc.GetField(sCandidate);
            });

            for (let nField = 0; nField < aOwnFields.length; ++nField) {
                if (aOwnFields[nField].GetFullName() === sName)
                    aOwnFields[nField].SetName(sNewName);
            }
        }
    }

    if (bParentRenamed && oDoc.private_UpdateTargetForCollaboration)
        oDoc.private_UpdateTargetForCollaboration(true);

    this.m_aFieldNameConflicts.length = 0;
};
CPDFCollaborativeEditing.prototype.Send_Changes = function(IsUserSave, AdditionalInfo, IsUpdateInterface, isAfterAskSave) {
	if (!this.canSendChanges())
		return;
	
	let oDoc        = this.GetDocument();
	let oHistory    = oDoc.History;
	
	let localHistory = AscCommon.History;
	AscCommon.History = oHistory;

    this.private_ResolveFieldNameConflicts();
	
	// Recalculate positions
	this.Refresh_DCChanges();

	
	AscCommon.DocumentEditorApi.prototype.asc_Save.apply(this, arguments);
	
	
    // Generate own changes
    let StartPoint = ( null === oHistory.SavedIndex ? 0 : oHistory.SavedIndex + 1 );
    let LastPoint = -1;

    if (this.m_nUseType <= 0) {
        // (remove unnecessary points in advance)
        oHistory.Clear_Redo();
        LastPoint = oHistory.Points.length - 1;
    }
    else {
        LastPoint = oHistory.Index;
    }

    // Calculate how many changes do not need to be sent to the server
    let SumIndex = 0;
    let StartPoint2 = Math.min(StartPoint, LastPoint + 1);

    for (let PointIndex = 0; PointIndex < StartPoint2; PointIndex++) {
        let Point = oHistory.Points[PointIndex];
        SumIndex += Point.Items.length;
    }
    let deleteIndex = ( null === oHistory.SavedIndex ? null : SumIndex );

    let aChanges = [], aChanges2 = [];
    for (let PointIndex = StartPoint; PointIndex <= LastPoint; PointIndex++) {
        let Point = oHistory.Points[PointIndex];
        oHistory.Update_PointInfoItem(PointIndex, StartPoint, LastPoint, SumIndex, deleteIndex);

        for (let Index = 0; Index < Point.Items.length; Index++)
        {
            let Item = Point.Items[Index];
            let oChanges = new AscCommon.CCollaborativeChanges();
            oChanges.Set_FromUndoRedo(Item.Class, Item.Data, Item.Binary);

            aChanges2.push(Item.Data);

            aChanges.push(oChanges.m_pData);
        }
    }

    let UnlockCount = 0;

    // While the user is alone, we don't clear their locks until another user joins
    let bCollaborative = this.getCollaborativeEditing();
    if (bCollaborative)
	{
		UnlockCount = this.m_aNeedUnlock.length;
		this.Release_Locks();

		let UnlockCount2 = this.m_aNeedUnlock2.length;
		for (let Index = 0; Index < UnlockCount2; Index++)
		{
			let Class = this.m_aNeedUnlock2[Index];
			Class.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
			editor.CoAuthoringApi.releaseLocks(Class.Get_Id());
		}

		this.m_aNeedUnlock.length  = 0;
		this.m_aNeedUnlock2.length = 0;
	}

	deleteIndex = ( null === oHistory.SavedIndex ? null : SumIndex );
	if (0 < aChanges.length || null !== deleteIndex) {
		this.CoHistory.AddOwnChanges(aChanges2, deleteIndex);
		editor.CoAuthoringApi.saveChanges(aChanges, deleteIndex, AdditionalInfo, editor.canUnlockDocument2, bCollaborative);
		oHistory.CanNotAddChanges = true;
	}
	else {
		editor.CoAuthoringApi.unLockDocument(!!isAfterAskSave, editor.canUnlockDocument2, null, bCollaborative);
	}

	editor.canUnlockDocument2 = false;

    if (-1 === this.m_nUseType) {
        // Clear Undo/Redo only during collaborative editing
        oHistory.Clear();
        oHistory.SavedIndex = null;
    }
    else if (0 === this.m_nUseType) {
        // Clear Undo/Redo only during collaborative editing
        oHistory.Clear();
        oHistory.SavedIndex = null;

        this.m_nUseType = 1;
    }
    else {
        // Update the last save point in history
        oHistory.Reset_SavedIndex(IsUserSave);
    }

    if (false !== IsUpdateInterface)
        editor.WordControl.m_oLogicDocument.UpdateInterface(undefined, true);

    // TODO: While history is cleared on save, we need to update Undo/Redo
    editor.WordControl.m_oLogicDocument.Document_UpdateUndoRedoState();

    // We don't check own locks. When all users leave, redraw happens and own locks are no longer drawn.
    if (0 !== UnlockCount || 1 !== this.m_nUseType) {
        // Redraw the document (to update locks)
        editor.getDocumentRenderer().paint();
    }

    editor.WordControl.m_oLogicDocument.getCompositeInput().checkState();
	AscCommon.History = localHistory;
};
CPDFCollaborativeEditing.prototype.OnEnd_Load_Objects = function()
{
    // This function is called when external objects (images and fonts) have loaded

    // Remove the lock
    AscCommon.CollaborativeEditing.Set_GlobalLock(false);
    AscCommon.CollaborativeEditing.Set_GlobalLockSelection(false);

	if (this.m_fEndLoadCallBack)
	{
		this.m_fEndLoadCallBack();
		this.m_fEndLoadCallBack = null;
	}

	this.m_oLogicDocument.ResumeRecalculate();
	this.m_oLogicDocument.RecalculateByChanges(this.CoHistory.GetAllChanges(), this.m_nRecalcIndexStart, this.m_nRecalcIndexEnd, false, undefined);
	this.m_oLogicDocument.UpdateTracks();
	
    this.m_aEndLoadCallbacks.forEach(function(callback) {
        callback();
    });
    this.m_aEndLoadCallbacks.length = 0;

    editor.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ApplyChanges);
};
CPDFCollaborativeEditing.prototype.canSendChanges = function(){
    let oApi = this.GetEditorApi();
    let oDoc = oApi.getPDFDoc();
    let oActionQueue = oDoc.GetActionsQueue();

    return oApi && oApi.canSendChanges() && !oActionQueue.IsInProgress();
};
CPDFCollaborativeEditing.prototype.Apply_Changes = function(fEndCallBack) {
	if (!this.m_aChanges.length)
		return fEndCallBack ? fEndCallBack() : null;

	let docHistory = this.GetDocument().History;
	docHistory.StartNoHistoryMode();
	
	this.GetDocument().currInkInDrawingProcess = null; // stop ink drawing
	AscCommon.CCollaborativeEditingBase.prototype.Apply_Changes.call(this, fEndCallBack);
	
	docHistory.EndNoHistoryMode();
};
CPDFCollaborativeEditing.prototype.OnEnd_ReadForeignChanges = function() {
	AscCommon.CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges.apply(this, arguments);
};
CPDFCollaborativeEditing.prototype.Check_MergeData = function() {};
CPDFCollaborativeEditing.prototype.Release_Locks = function() {
    let UnlockCount = this.m_aNeedUnlock.length;
    for (let Index = 0; Index < UnlockCount; Index++) {
        let Class = this.m_aNeedUnlock[Index];
        let CurLockType = Class.Lock.Get_Type();
        
        if (AscCommon.c_oAscLockTypes.kLockTypeOther3 != CurLockType && AscCommon.c_oAscLockTypes.kLockTypeOther != CurLockType) {
            Class.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
            Class.AddToRedraw && Class.AddToRedraw();

            if (Class.IsAnnot && Class.IsAnnot()) {
                Asc.editor.sync_UnLockComment(Class.Get_Id());
            }
        }
        else if (AscCommon.c_oAscLockTypes.kLockTypeOther3 === CurLockType)
        {
            Class.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
            Class.AddToRedraw && Class.AddToRedraw();
        }
    }
};
CPDFCollaborativeEditing.prototype._PreUndo = function() {
    return this.private_SaveDocumentState()
};
CPDFCollaborativeEditing.prototype._PostUndo = function(state, changes) {
    let logicDocument = this.m_oLogicDocument;
    this.private_RestoreDocumentState(state);
    logicDocument.History.Get_RecalcData(null, changes)
};

function private_CollectFieldNodes(aWidgets) {
    let oNodes = {};

    for (let nWidget = 0; nWidget < aWidgets.length; ++nWidget) {
        let oField = aWidgets[nWidget];
        while (oField) {
            let sId = oField.GetId();
            if (!oNodes[sId]) {
                oNodes[sId] = {
                    Field: oField,
                    Name:  oField.GetFullName()
                };
            }

            oField = oField.GetParent();
        }
    }

    return oNodes;
}

//--------------------------------------------------------export----------------------------------------------------
window['AscPDF'] = window['AscPDF'] || {};
window['AscPDF'].CPDFCollaborativeEditing = CPDFCollaborativeEditing;
