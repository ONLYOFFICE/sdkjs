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
function CPDFCollaborativeEditing() {
	AscCommon.CWordCollaborativeEditing.call(this);
    this.m_aSkipContentControlsOnCheckEditingLock = {};
}

CPDFCollaborativeEditing.prototype = Object.create(AscCommon.CWordCollaborativeEditing.prototype);
CPDFCollaborativeEditing.prototype.constructor = CPDFCollaborativeEditing;

CPDFCollaborativeEditing.prototype.GetDocument = function() {
    return this.m_oLogicDocument;
};
CPDFCollaborativeEditing.prototype.Send_Changes = function(IsUserSave, AdditionalInfo, IsUpdateInterface, isAfterAskSave) {
	if (!this.canSendChanges())
		return;
	
    // Пересчитываем позиции
    this.Refresh_DCChanges();

    let oDoc        = this.GetDocument();
    let oHistory    = oDoc.History;

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
	
	let oform = this.m_oLogicDocument.GetOFormDocument();
	if (oform)
		oform.onEndLoadChanges();

    editor.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ApplyChanges);
};
CPDFCollaborativeEditing.prototype.canSendChanges = function(){
    let oApi = this.GetEditorApi();
    let oDoc = oApi.getPDFDoc();
    let oActionQueue = oDoc.GetActionsQueue();

    return oApi && oApi.canSendChanges() && !oActionQueue.IsInProgress();
};
CPDFCollaborativeEditing.prototype.OnEnd_ReadForeignChanges = function() {
	AscCommon.CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges.apply(this, arguments);
};
CPDFCollaborativeEditing.prototype.Check_MergeData = function() {};

//--------------------------------------------------------export----------------------------------------------------
window['AscPDF'] = window['AscPDF'] || {};
window['AscPDF'].CPDFCollaborativeEditing = CPDFCollaborativeEditing;
