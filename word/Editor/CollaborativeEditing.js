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
function CWordCollaborativeEditing()
{
	AscCommon.CCollaborativeEditingBase.call(this);
    this.m_aSkipContentControlsOnCheckEditingLock = {};
}

CWordCollaborativeEditing.prototype = Object.create(AscCommon.CCollaborativeEditingBase.prototype);
CWordCollaborativeEditing.prototype.constructor = CWordCollaborativeEditing;
CWordCollaborativeEditing.prototype.GetEditorApi = function()
{
    if(this.m_oLogicDocument)
    {
        return this.m_oLogicDocument.GetApi();
    }
    return null;
};
CWordCollaborativeEditing.prototype.GetDrawingDocument = function()
{
    if(this.m_oLogicDocument)
    {
        return this.m_oLogicDocument.GetDrawingDocument();
    }
    return null;
};
CWordCollaborativeEditing.prototype.Clear = function()
{
	AscCommon.CCollaborativeEditingBase.prototype.Clear.apply(this, arguments);
	this.Remove_AllForeignCursors();
};
CWordCollaborativeEditing.prototype.Send_Changes = function(IsUserSave, AdditionalInfo, IsUpdateInterface, isAfterAskSave)
{
	if (!this.canSendChanges())
		return;
	
    // Пересчитываем позиции
    this.Refresh_DCChanges();

    // Генерируем свои изменения
    var StartPoint = ( null === AscCommon.History.SavedIndex ? 0 : AscCommon.History.SavedIndex + 1 );
    var LastPoint = -1;

    if (this.m_nUseType <= 0)
    {
        // (ненужные точки предварительно удаляем)
        AscCommon.History.Clear_Redo();
        LastPoint = AscCommon.History.Points.length - 1;
    }
    else
    {
        LastPoint = AscCommon.History.Index;
    }

    // Просчитаем сколько изменений на сервер пересылать не надо
    var SumIndex = 0;
    var StartPoint2 = Math.min(StartPoint, LastPoint + 1);
    for (var PointIndex = 0; PointIndex < StartPoint2; PointIndex++)
    {
        var Point = AscCommon.History.Points[PointIndex];
        SumIndex += Point.Items.length;
    }
    var deleteIndex = ( null === AscCommon.History.SavedIndex ? null : SumIndex );

    var aChanges = [], aChanges2 = [];
    for (var PointIndex = StartPoint; PointIndex <= LastPoint; PointIndex++)
    {
        var Point = AscCommon.History.Points[PointIndex];
        AscCommon.History.Update_PointInfoItem(PointIndex, StartPoint, LastPoint, SumIndex, deleteIndex);

        for (var Index = 0; Index < Point.Items.length; Index++)
        {
            var Item = Point.Items[Index];
            var oChanges = new AscCommon.CCollaborativeChanges();
            oChanges.Set_FromUndoRedo(Item.Class, Item.Data, Item.Binary);

            aChanges2.push(Item.Data);

            aChanges.push(oChanges.m_pData);
        }
    }

    var UnlockCount = 0;

    // Пока пользователь сидит один, мы не чистим его локи до тех пор пока не зайдет второй
    var bCollaborative = this.getCollaborativeEditing();
    if (bCollaborative)
	{
		UnlockCount = this.m_aNeedUnlock.length;
		this.Release_Locks();

		var UnlockCount2 = this.m_aNeedUnlock2.length;
		for (var Index = 0; Index < UnlockCount2; Index++)
		{
			var Class = this.m_aNeedUnlock2[Index];
			Class.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
			editor.CoAuthoringApi.releaseLocks(Class.Get_Id());
		}

		this.m_aNeedUnlock.length  = 0;
		this.m_aNeedUnlock2.length = 0;
	}

	var deleteIndex = ( null === AscCommon.History.SavedIndex ? null : SumIndex );
	if (0 < aChanges.length || null !== deleteIndex)
	{
		this.CoHistory.AddOwnChanges(aChanges2, deleteIndex);
		editor.CoAuthoringApi.saveChanges(aChanges, deleteIndex, AdditionalInfo, editor.canUnlockDocument2, bCollaborative);
		AscCommon.History.CanNotAddChanges = true;
	}
	else
	{
		editor.CoAuthoringApi.unLockDocument(!!isAfterAskSave, editor.canUnlockDocument2, null, bCollaborative);
	}
	editor.canUnlockDocument2 = false;

    if (-1 === this.m_nUseType)
    {
        // Чистим Undo/Redo только во время совместного редактирования
        AscCommon.History.Clear();
        AscCommon.History.SavedIndex = null;
    }
    else if (0 === this.m_nUseType)
    {
        // Чистим Undo/Redo только во время совместного редактирования
        AscCommon.History.Clear();
        AscCommon.History.SavedIndex = null;

        this.m_nUseType = 1;
    }
    else
    {
        // Обновляем точку последнего сохранения в истории
        AscCommon.History.Reset_SavedIndex(IsUserSave);
    }

    if (false !== IsUpdateInterface)
        editor.WordControl.m_oLogicDocument.UpdateInterface(undefined, true);

    // TODO: Пока у нас обнуляется история на сохранении нужно обновлять Undo/Redo
    editor.WordControl.m_oLogicDocument.Document_UpdateUndoRedoState();

    // Свои локи не проверяем. Когда все пользователи выходят, происходит перерисовка и свои локи уже не рисуются.
    if (0 !== UnlockCount || 1 !== this.m_nUseType)
    {
        // Перерисовываем документ (для обновления локов)
        editor.WordControl.m_oLogicDocument.DrawingDocument.ClearCachePages();
        editor.WordControl.m_oLogicDocument.DrawingDocument.FirePaint();
    }

    editor.WordControl.m_oLogicDocument.getCompositeInput().checkState();
};
CWordCollaborativeEditing.prototype.Release_Locks = function()
{
    var UnlockCount = this.m_aNeedUnlock.length;
    for (var Index = 0; Index < UnlockCount; Index++)
    {
        var CurLockType = this.m_aNeedUnlock[Index].Lock.Get_Type();
        if (AscCommon.c_oAscLockTypes.kLockTypeOther3 != CurLockType && AscCommon.c_oAscLockTypes.kLockTypeOther != CurLockType)
        {
            this.m_aNeedUnlock[Index].Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);

            if (this.m_aNeedUnlock[Index] instanceof AscCommonWord.CHeaderFooterController)
                editor.sync_UnLockHeaderFooters();
            else if (this.m_aNeedUnlock[Index] instanceof AscCommonWord.CDocument)
                editor.sync_UnLockDocumentProps();
            else if (this.m_aNeedUnlock[Index] instanceof AscCommon.CComment)
                editor.sync_UnLockComment(this.m_aNeedUnlock[Index].Get_Id());
            else if (this.m_aNeedUnlock[Index] instanceof AscCommonWord.CGraphicObjects)
                editor.sync_UnLockDocumentSchema();
            else if (this.m_aNeedUnlock[Index] instanceof AscCommon.CCore)
                editor.sendEvent("asc_onLockCore", false);
            else if (this.m_aNeedUnlock[Index] instanceof AscCommonWord.CDocProtect)
                editor.sendEvent("asc_onLockDocumentProtection", false);
        }
        else if (AscCommon.c_oAscLockTypes.kLockTypeOther3 === CurLockType)
        {
            this.m_aNeedUnlock[Index].Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
        }
    }
};
CWordCollaborativeEditing.prototype.OnEnd_Load_Objects = function()
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

	var nPageIndex  = undefined;
    if (this.Is_Fast())
	{
		var oParagraph = this.m_oLogicDocument.GetCurrentParagraph();
		nPageIndex     = oParagraph ? Math.max(oParagraph.GetCurrentPageAbsolute(), editor.GetCurrentVisiblePage()) : undefined;
	}

	this.m_oLogicDocument.ResumeRecalculate();
	this.m_oLogicDocument.RecalculateByChanges(this.CoHistory.GetAllChanges(), this.m_nRecalcIndexStart, this.m_nRecalcIndexEnd, false, nPageIndex);
	this.m_oLogicDocument.UpdateTracks();
	
	let oform = this.m_oLogicDocument.GetOFormDocument();
	if (oform)
		oform.onEndLoadChanges();

    editor.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ApplyChanges);
};
CWordCollaborativeEditing.prototype.Check_MergeData = function()
{
    var LogicDocument = editor.WordControl.m_oLogicDocument;
    LogicDocument.Comments.Check_MergeData();
};
CWordCollaborativeEditing.prototype.OnEnd_CheckLock = function(isDontLockInFastMode, fCallback)
{
	// Если задан fCallback, тогда действие нужно выполнять именно на нём, поэотому возвращаемое значение true

	var aIds = [];
	for (var nIndex = 0, nCount = this.m_aCheckLocks.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.m_aCheckLocks[nIndex];

		if (true === oItem) // сравниваем по значению и типу обязательно
		{
			if (fCallback)
				fCallback(true);

			return true;
		}
		else if (false !== oItem)
		{
			aIds.push(oItem);
		}
	}

	if ((true === isDontLockInFastMode && true === this.Is_Fast()) || !this.canSendChanges())
	{
		if (fCallback)
		{
			fCallback(false);
			return true;
		}
		else
		{
			return false;
		}
	}

	if (fCallback)
	{
		if (aIds.length > 0)
		{
			var oThis = this;
			editor.CoAuthoringApi.askLock(aIds, function(result)
			{
				oThis.Set_GlobalLock(false);

				if (result["lock"])
				{
					oThis.private_LockByMe();
					fCallback(false);
				}
				else if (result["error"])
				{
					fCallback(true);
				}
			});

			this.Set_GlobalLock(true);
		}
		else
		{
			fCallback(false);
		}

		return true;
	}
	else
	{
		if (aIds.length > 0)
		{
			// Отправляем запрос на сервер со списком Id
			editor.CoAuthoringApi.askLock(aIds, this.OnCallback_AskLock);

			// Ставим глобальный лок, только во время совместного редактирования
			if (-1 === this.m_nUseType)
			{
				this.Set_GlobalLock(true);
			}
			else
			{
				this.private_LockByMe();
				this.m_aCheckLocks.length = 0;
			}
		}

		return false;
	}
};
CWordCollaborativeEditing.prototype.private_LockByMe = function()
{
	for (var nIndex = 0, nCount = this.m_aCheckLocks.length; nIndex < nCount; ++nIndex)
	{
		var oItem = this.m_aCheckLocks[nIndex];
		if (true !== oItem && false !== oItem) // сравниваем по значению и типу обязательно
		{
			var oClass = AscCommon.g_oTableId.Get_ById(oItem);
			if (oClass)
			{
				oClass.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeMine);
				this.Add_Unlock2(oClass);
			}
		}
	}
};
CWordCollaborativeEditing.prototype.OnCallback_AskLock = function(result)
{
    var oThis   = AscCommon.CollaborativeEditing;
    var oEditor = editor;

    if (true === oThis.Get_GlobalLock())
    {
        // Здесь проверяем есть ли длинная операция, если она есть, то до ее окончания нельзя делать
        // Undo, иначе точка истории уберется, а изменения допишутся в предыдущую.
        if (false == oEditor.checkLongActionCallback(oThis.OnCallback_AskLock, result))
            return;

        // Снимаем глобальный лок
        oThis.Set_GlobalLock(false);

        if (result["lock"])
        {
        	oThis.private_LockByMe();
		}
        else if (result["error"])
        {
            // Если у нас началось редактирование диаграммы, а вернулось, что ее редактировать нельзя,
            // посылаем сообщение о закрытии редактора диаграмм.
            if (oEditor.isOpenedChartFrame)
                oEditor.sync_closeChartEditor();

          if (oEditor.isOleEditor)
            oEditor.sync_closeOleEditor();

            // Делаем откат на 1 шаг назад и удаляем из Undo/Redo эту последнюю точку
            oEditor.WordControl.m_oLogicDocument.Document_Undo();
            AscCommon.History.Clear_Redo();
        }

        oEditor.isChartEditor = false;
        oEditor.isOleEditor = false;
    }
};
CWordCollaborativeEditing.prototype.AddContentControlForSkippingOnCheckEditingLock = function(oContentControl)
{
	this.m_aSkipContentControlsOnCheckEditingLock[oContentControl.GetId()] = oContentControl;
};
CWordCollaborativeEditing.prototype.RemoveContentControlForSkippingOnCheckEditingLock = function(oContentControl)
{
	delete this.m_aSkipContentControlsOnCheckEditingLock[oContentControl.GetId()];
};
CWordCollaborativeEditing.prototype.IsNeedToSkipContentControlOnCheckEditingLock = function(oContentControl)
{
	if (this.m_aSkipContentControlsOnCheckEditingLock[oContentControl.GetId()] === oContentControl)
		return true;

	return false;
};
CWordCollaborativeEditing.prototype.Start_CollaborationEditing = function()
{
	this.m_nUseType = -1;
};
CWordCollaborativeEditing.prototype.End_CollaborationEditing = function()
{
	if (this.m_nUseType <= 0)
	{
		if (this.m_oLogicDocument && !this.m_oLogicDocument.GetHistory().Have_Changes() && !this.Have_OtherChanges())
			this.m_nUseType = 1;
		else
			this.m_nUseType = 0;
	}
};
CWordCollaborativeEditing.prototype._PreUndo = function()
{
	let logicDocument = this.m_oLogicDocument;
	
	logicDocument.DrawingDocument.EndTrackTable(null, true);
	logicDocument.TurnOffCheckChartSelection();
	
	return this.private_SaveDocumentState()
};
CWordCollaborativeEditing.prototype._PostUndo = function(state, changes)
{
	this.private_RestoreDocumentState(state);
	this.private_RecalculateDocument(changes);
	
	let logicDocument = this.m_oLogicDocument;
	logicDocument.TurnOnCheckChartSelection();
	logicDocument.UpdateSelection();
	logicDocument.UpdateInterface();
	logicDocument.UpdateRulers();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с сохраненными позициями документа.
//----------------------------------------------------------------------------------------------------------------------
CWordCollaborativeEditing.prototype.Remove_ForeignCursor = function(UserId)
{
    AscCommon.CCollaborativeEditingBase.prototype.Remove_ForeignCursor.call(this, UserId);
    var DrawingDocument = this.m_oLogicDocument.DrawingDocument;
    DrawingDocument.Collaborative_RemoveTarget(UserId);
    this.Remove_ForeignCursorXY(UserId);
    this.Remove_ForeignCursorToShow(UserId);
};
CWordCollaborativeEditing.prototype.Update_ForeignCursorsPositions = function()
{
    for (var UserId in this.m_aForeignCursors)
    {
        var DocPos = this.m_aForeignCursors[UserId];
        if (!DocPos || DocPos.length <= 0)
            continue;

        this.m_aForeignCursorsPos.Update_DocumentPosition(DocPos);

        var Run      = DocPos[DocPos.length - 1].Class;
        var InRunPos = DocPos[DocPos.length - 1].Position;

        this.Update_ForeignCursorPosition(UserId, Run, InRunPos, false);
    }
};
CWordCollaborativeEditing.prototype.Update_ForeignCursorPosition = function(UserId, Run, InRunPos, isRemoveLabel)
{
    var DrawingDocument = this.m_oLogicDocument.DrawingDocument;

    if (!(Run instanceof AscCommonWord.ParaRun))
        return;

    var Paragraph = Run.GetParagraph();

    if (!Paragraph)
    {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }

    var ParaContentPos = Paragraph.Get_PosByElement(Run);
    if (!ParaContentPos)
    {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }
    ParaContentPos.Update(InRunPos, ParaContentPos.GetDepth() + 1);

    var XY = Paragraph.Get_XYByContentPos(ParaContentPos);
    if (XY && XY.Height > 0.001)
    {
        var ShortId = this.m_aForeignCursorsId[UserId] ? this.m_aForeignCursorsId[UserId] : UserId;
        DrawingDocument.Collaborative_UpdateTarget(UserId, ShortId, XY.X, XY.Y, XY.Height, XY.PageNum, Paragraph.Get_ParentTextTransform());
        this.Add_ForeignCursorXY(UserId, XY.X, XY.Y, XY.PageNum, XY.Height, Paragraph, isRemoveLabel);

        if (true === this.m_aForeignCursorsToShow[UserId])
        {
            this.Show_ForeignCursorLabel(UserId);
            this.Remove_ForeignCursorToShow(UserId);
        }
    }
    else
    {
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        this.Remove_ForeignCursorXY(UserId);
        this.Remove_ForeignCursorToShow(UserId);
    }
};
CWordCollaborativeEditing.prototype.OnEnd_ReadForeignChanges = function()
{
	AscCommon.CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges.apply(this, arguments);

	if (this.m_oLogicDocument && this.m_oLogicDocument.GetApi())
	{
		var oApi = this.m_oLogicDocument.GetApi();
		if (this.m_oLogicDocument.GetBookmarksManager().IsNeedUpdate())
		{
			oApi.asc_OnBookmarksUpdate();
		}

		if (this.m_oLogicDocument && this.m_oLogicDocument.ClearListsCache)
			this.m_oLogicDocument.ClearListsCache();

		if (this.m_oLogicDocument && this.m_oLogicDocument.Settings && this.m_oLogicDocument.Settings.DocumentProtection)
		{
			let updateFromUser = this.m_oLogicDocument.Settings.DocumentProtection.GetNeedUpdate();
			if (updateFromUser) {
				oApi.asc_OnProtectionUpdate(updateFromUser);
				this.m_oLogicDocument.Settings.DocumentProtection.SetNeedUpdate(null);
			}
		}
	}
};
CWordCollaborativeEditing.prototype.private_RecalculateDocument = function(arrChanges)
{
	this.m_oLogicDocument.RecalculateByChanges(arrChanges);
};
CWordCollaborativeEditing.prototype.private_UpdateForeignCursor = function(CursorInfo, UserId, Show, UserShortId)
{
    this.m_oLogicDocument.Update_ForeignCursor(CursorInfo, UserId, Show, UserShortId);
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].CWordCollaborativeEditing = CWordCollaborativeEditing;
