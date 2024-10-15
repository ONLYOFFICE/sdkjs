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
}

CPDFCollaborativeEditing.prototype = Object.create(AscCommon.CWordCollaborativeEditing.prototype);
CPDFCollaborativeEditing.prototype.constructor = CPDFCollaborativeEditing;

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

            let aOrigRect = aObjects[i].GetOrigRect();
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
CPDFCollaborativeEditing.prototype.Send_Changes = function(IsUserSave, AdditionalInfo, IsUpdateInterface, isAfterAskSave) {
	if (!this.canSendChanges())
		return;
	
	let oDoc        = this.GetDocument();
	let oHistory    = oDoc.History;
	
	let localHistory = AscCommon.History;
	AscCommon.History = oHistory;
	
	// Пересчитываем позиции
	this.Refresh_DCChanges();

	
	AscCommon.DocumentEditorApi.prototype.asc_Save.apply(this, arguments);
	
	
    // Генерируем свои изменения
    let StartPoint = ( null === oHistory.SavedIndex ? 0 : oHistory.SavedIndex + 1 );
    let LastPoint = -1;

    if (this.m_nUseType <= 0) {
        // (ненужные точки предварительно удаляем)
        oHistory.Clear_Redo();
        LastPoint = oHistory.Points.length - 1;
    }
    else {
        LastPoint = oHistory.Index;
    }

    // Просчитаем сколько изменений на сервер пересылать не надо
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

    // Пока пользователь сидит один, мы не чистим его локи до тех пор пока не зайдет второй
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
        // Чистим Undo/Redo только во время совместного редактирования
        oHistory.Clear();
        oHistory.SavedIndex = null;
    }
    else if (0 === this.m_nUseType) {
        // Чистим Undo/Redo только во время совместного редактирования
        oHistory.Clear();
        oHistory.SavedIndex = null;

        this.m_nUseType = 1;
    }
    else {
        // Обновляем точку последнего сохранения в истории
        oHistory.Reset_SavedIndex(IsUserSave);
    }

    if (false !== IsUpdateInterface)
        editor.WordControl.m_oLogicDocument.UpdateInterface(undefined, true);

    // TODO: Пока у нас обнуляется история на сохранении нужно обновлять Undo/Redo
    editor.WordControl.m_oLogicDocument.Document_UpdateUndoRedoState();

    // Свои локи не проверяем. Когда все пользователи выходят, происходит перерисовка и свои локи уже не рисуются.
    if (0 !== UnlockCount || 1 !== this.m_nUseType) {
        // Перерисовываем документ (для обновления локов)
        editor.WordControl.m_oLogicDocument.DrawingDocument.ClearCachePages();
        editor.WordControl.m_oLogicDocument.DrawingDocument.FirePaint();
    }

    editor.WordControl.m_oLogicDocument.getCompositeInput().checkState();
	AscCommon.History = localHistory;
};
CPDFCollaborativeEditing.prototype.OnEnd_Load_Objects = function()
{
    // Данная функция вызывается, когда загрузились внешние объекты (картинки и шрифты)

    // Снимаем лок
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
	docHistory.TurnOff();
	
	this.GetDocument().currInkInDrawingProcess = null; // останавливаем ink рисование
	AscCommon.CCollaborativeEditingBase.prototype.Apply_Changes.call(this, fEndCallBack);
	
	docHistory.TurnOn();
};
CPDFCollaborativeEditing.prototype.OnEnd_ReadForeignChanges = function() {
	AscCommon.CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges.apply(this, arguments);
};
CPDFCollaborativeEditing.prototype.Check_MergeData = function() {};

//--------------------------------------------------------export----------------------------------------------------
window['AscPDF'] = window['AscPDF'] || {};
window['AscPDF'].CPDFCollaborativeEditing = CPDFCollaborativeEditing;
